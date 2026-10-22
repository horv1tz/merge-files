//! Конвертация документов: LibreOffice (headless), а в Windows при его отсутствии —
//! установленный Microsoft Office через PowerShell/COM.
//!
//! Все внешние процессы запускаются с ограничением по времени и отменой пользователем.

use anyhow::{Context, Result, anyhow, bail};
use std::io::Read;
use std::path::{Path, PathBuf};
use std::process::{Command, Stdio};
use std::sync::Mutex;
use std::sync::atomic::{AtomicBool, Ordering};
use std::time::{Duration, Instant};

use crate::shell::hide_window;

/// Сколько ждём одну конвертацию, прежде чем считать, что конвертер завис.
const TIMEOUT: Duration = Duration::from_secs(10 * 60);

/// Результат работы внешней команды.
struct Output {
    ok: bool,
    stdout: String,
    stderr: String,
}

/// Запускает команду, опрашивая флаг отмены и таймаут; при отмене убивает всё дерево процессов.
fn run_cmd(cmd: &mut Command, cancel: &AtomicBool, timeout: Duration) -> Result<Output> {
    hide_window(cmd);
    #[cfg(unix)]
    {
        // Отдельная группа процессов: при отмене убиваем и дочерний soffice.bin.
        use std::os::unix::process::CommandExt;
        cmd.process_group(0);
    }
    let mut child = cmd
        .stdin(Stdio::null())
        .stdout(Stdio::piped())
        .stderr(Stdio::piped())
        .spawn()
        .context("не удалось запустить внешнюю программу")?;
    let drain = |pipe: Option<Box<dyn Read + Send>>| {
        std::thread::spawn(move || {
            let mut buf = Vec::new();
            if let Some(mut p) = pipe {
                let _ = p.read_to_end(&mut buf);
            }
            String::from_utf8_lossy(&buf).into_owned()
        })
    };
    let out_t = drain(child.stdout.take().map(|p| Box::new(p) as Box<dyn Read + Send>));
    let err_t = drain(child.stderr.take().map(|p| Box::new(p) as Box<dyn Read + Send>));
    let started = Instant::now();
    let status = loop {
        if let Some(st) = child.try_wait()? {
            break Some(st);
        }
        if cancel.load(Ordering::Relaxed) || started.elapsed() > timeout {
            kill_tree(&mut child);
            break None;
        }
        std::thread::sleep(Duration::from_millis(80));
    };
    let Some(st) = status else {
        // Потоки чтения не ждём: их трубы может держать оставшийся дочерний процесс.
        return if cancel.load(Ordering::Relaxed) {
            Err(anyhow!(CANCELLED))
        } else {
            Err(anyhow!("Конвертер не ответил за {} мин. и был остановлен", timeout.as_secs() / 60))
        };
    };
    let stdout = out_t.join().unwrap_or_default();
    let stderr = err_t.join().unwrap_or_default();
    Ok(Output { ok: st.success(), stdout, stderr })
}

/// Текст ошибки отмены (по нему интерфейс отличает отмену от сбоя).
pub const CANCELLED: &str = "Операция отменена";

fn kill_tree(child: &mut std::process::Child) {
    if cfg!(unix) {
        let _ = Command::new("kill").args(["-KILL", &format!("-{}", child.id())]).output();
    }
    if cfg!(windows) {
        // soffice.exe запускает дочерний soffice.bin — убиваем всё дерево.
        let mut k = Command::new("taskkill");
        hide_window(&mut k);
        let _ = k.args(["/PID", &child.id().to_string(), "/T", "/F"]).output();
    }
    let _ = child.kill();
    let _ = child.wait();
}

// ───────────────────────── поиск LibreOffice ─────────────────────────

/// Найденный LibreOffice кэшируется; «не найден» помним лишь несколько секунд — пользователь мог его установить.
static SOFFICE: Mutex<(Option<PathBuf>, Option<Instant>)> = Mutex::new((None, None));
const NEGATIVE_TTL: Duration = Duration::from_secs(5);

fn find_soffice() -> Option<PathBuf> {
    let mut g = SOFFICE.lock().unwrap_or_else(|e| e.into_inner());
    if g.0.is_some() {
        return g.0.clone();
    }
    if g.1.is_some_and(|t| t.elapsed() < NEGATIVE_TTL) {
        return None;
    }
    g.0 = discover_soffice();
    g.1 = Some(Instant::now());
    g.0.clone()
}

fn discover_soffice() -> Option<PathBuf> {
    let mut candidates: Vec<PathBuf> = vec![];
    if cfg!(windows) {
        for var in ["ProgramFiles", "ProgramW6432", "ProgramFiles(x86)"] {
            if let Some(base) = std::env::var_os(var) {
                push_libreoffice_dirs(&mut candidates, Path::new(&base));
            }
        }
        if let Some(local) = std::env::var_os("LOCALAPPDATA") {
            push_libreoffice_dirs(&mut candidates, &Path::new(&local).join("Programs"));
        }
        for root in [
            r"HKLM\SOFTWARE\LibreOffice\UNO\InstallPath",
            r"HKLM\SOFTWARE\WOW6432Node\LibreOffice\UNO\InstallPath",
            r"HKCU\SOFTWARE\LibreOffice\UNO\InstallPath",
        ] {
            if let Some(dir) = reg_default_value(root) {
                candidates.push(PathBuf::from(dir).join("soffice.exe"));
            }
        }
    } else {
        candidates.push("/Applications/LibreOffice.app/Contents/MacOS/soffice".into());
        candidates.push("/opt/libreoffice/program/soffice".into());
        candidates.push("/usr/bin/soffice".into());
        candidates.push("/usr/bin/libreoffice".into());
        candidates.push("/usr/local/bin/soffice".into());
        candidates.push("/snap/bin/libreoffice".into());
    }
    if let Some(found) = candidates.into_iter().find(|p| p.is_file()) {
        return Some(found);
    }
    // Поиск в PATH.
    let names: &[&str] = if cfg!(windows) { &["soffice.exe", "libreoffice.exe"] } else { &["soffice", "libreoffice"] };
    let path = std::env::var_os("PATH")?;
    std::env::split_paths(&path).flat_map(|d| names.iter().map(move |n| d.join(n))).find(|p| p.is_file())
}

fn push_libreoffice_dirs(out: &mut Vec<PathBuf>, base: &Path) {
    let Ok(rd) = std::fs::read_dir(base) else { return };
    for e in rd.flatten() {
        if e.file_name().to_string_lossy().to_lowercase().starts_with("libreoffice") {
            out.push(e.path().join("program").join("soffice.exe"));
        }
    }
}

/// Значение по умолчанию ключа реестра (через reg.exe).
fn reg_default_value(key: &str) -> Option<String> {
    let mut cmd = Command::new("reg");
    hide_window(&mut cmd);
    let out = cmd.args(["query", key, "/ve"]).output().ok()?;
    if !out.status.success() {
        return None;
    }
    let text = crate::shell::decode_console(&out.stdout);
    let line = text.lines().find(|l| l.contains("REG_SZ"))?;
    Some(line.split("REG_SZ").nth(1)?.trim().to_string())
}

fn office_installed() -> bool {
    cfg!(windows) && ["Word.Application", "Excel.Application"].iter().any(|p| reg_default_value(&format!(r"HKCR\{p}")).is_some())
}

/// Есть ли чем конвертировать (LibreOffice или, в Windows, Microsoft Office).
pub fn available() -> bool {
    static OFFICE: Mutex<Option<(bool, Instant)>> = Mutex::new(None);
    if find_soffice().is_some() {
        return true;
    }
    let mut g = OFFICE.lock().unwrap_or_else(|e| e.into_inner());
    match *g {
        Some((v, t)) if v || t.elapsed() < NEGATIVE_TTL => v,
        _ => {
            let v = office_installed();
            *g = Some((v, Instant::now()));
            v
        }
    }
}

pub const MISSING_HINT: &str =
    "Для этой операции нужен LibreOffice (в Windows подойдёт и Microsoft Office). Установите LibreOffice: https://www.libreoffice.org";

// ───────────────────────── конвертер ─────────────────────────

/// Конвертер с общим профилем LibreOffice (холодный старт происходит один раз).
pub struct Converter<'a> {
    profile: PathBuf,
    cancel: &'a AtomicBool,
}

impl<'a> Converter<'a> {
    pub fn new(work_dir: &Path, cancel: &'a AtomicBool) -> Self {
        Converter { profile: work_dir.join("lo-profile"), cancel }
    }

    /// Конвертирует `input` в формат `fmt` ("pdf", "docx", "xlsx"), результат кладёт в `out_dir`.
    pub fn convert(&self, input: &Path, fmt: &str, out_dir: &Path) -> Result<PathBuf> {
        self.convert_inner(input, fmt, out_dir, false)
    }

    /// Последний шаг: объединённый файл → PDF (ошибка не должна обвинять исходные файлы).
    pub fn convert_result(&self, merged: &Path, fmt: &str, out_dir: &Path) -> Result<PathBuf> {
        self.convert_inner(merged, fmt, out_dir, true)
    }

    fn convert_inner(&self, input: &Path, fmt: &str, out_dir: &Path, is_result: bool) -> Result<PathBuf> {
        std::fs::create_dir_all(out_dir)?;
        let stem = input.file_stem().map(|s| s.to_string_lossy().into_owned()).unwrap_or_default();
        let expected = out_dir.join(format!("{stem}.{fmt}"));
        let name = input.file_name().map(|s| s.to_string_lossy().into_owned()).unwrap_or_default();
        let unreadable = || {
            if is_result {
                anyhow!("Не удалось создать {} из объединённого файла: конвертер завершился с ошибкой", fmt.to_uppercase())
            } else {
                anyhow!("Не удалось преобразовать «{name}»: файл повреждён, защищён паролем или имеет другой формат, чем указано в расширении")
            }
        };

        let Some(soffice) = find_soffice() else {
            return self
                .convert_with_office(input, fmt, &expected)
                .map_err(|e| anyhow!("{MISSING_HINT} (Microsoft Office тоже не сработал: {e})"));
        };
        let url = file_url(&self.profile);
        let mut cmd = Command::new(soffice);
        cmd.arg("--headless")
            .arg("--norestore")
            .arg("--nolockcheck")
            .arg(format!("-env:UserInstallation={url}"))
            .args(["--convert-to", fmt, "--outdir"])
            .arg(out_dir)
            .arg(input);
        let out = run_cmd(&mut cmd, self.cancel, TIMEOUT)?;
        if !expected.exists() {
            let detail = format!("{} {}", out.stdout.trim(), out.stderr.trim());
            let detail = detail.trim();
            if out.ok || detail.is_empty() || detail.chars().count() > 200 {
                return Err(unreadable());
            }
            return Err(anyhow!("{} ({detail})", unreadable()));
        }
        Ok(expected)
    }

    /// Windows: конвертация через Microsoft Word/Excel (COM). На других системах недоступна.
    fn convert_with_office(&self, input: &Path, fmt: &str, out: &Path) -> Result<PathBuf> {
        if !cfg!(windows) {
            bail!("только Windows");
        }
        let ext = input.extension().and_then(|e| e.to_str()).unwrap_or("").to_lowercase();
        let script = match (ext.as_str(), fmt) {
            ("doc" | "docx", "pdf" | "docx") => {
                let code = if fmt == "pdf" { 17 } else { 16 };
                format!(
                    "$ErrorActionPreference='Stop'; $a=New-Object -ComObject Word.Application; $a.Visible=$false; $a.DisplayAlerts=0; \
                     $a.AutomationSecurity=3; \
                     try {{ $d=$a.Documents.Open($env:MF_IN,$false,$true); $d.SaveAs([ref]$env:MF_OUT,[ref]{code}); $d.Close($false) }} finally {{ $a.Quit() }}"
                )
            }
            ("xls" | "xlsx" | "xlsm" | "ods", "pdf" | "xlsx") => {
                let body = if fmt == "pdf" { "$w.ExportAsFixedFormat(0,$env:MF_OUT)" } else { "$w.SaveAs($env:MF_OUT,51)" };
                format!(
                    "$ErrorActionPreference='Stop'; $a=New-Object -ComObject Excel.Application; $a.Visible=$false; $a.DisplayAlerts=$false; \
                     $a.AutomationSecurity=3; \
                     try {{ $w=$a.Workbooks.Open($env:MF_IN,0,$true); {body}; $w.Close($false) }} finally {{ $a.Quit() }}"
                )
            }
            _ => bail!("неподдерживаемое преобразование"),
        };
        let mut cmd = Command::new("powershell");
        cmd.args(["-NoProfile", "-NonInteractive", "-Command", &script]).env("MF_IN", input).env("MF_OUT", out);
        let res = run_cmd(&mut cmd, self.cancel, TIMEOUT)?;
        // Office иногда сообщает об ошибке уже после успешного сохранения — важен сам файл.
        if !out.exists() {
            bail!("{}", res.stderr.lines().next().unwrap_or("файл не создан").trim());
        }
        Ok(out.to_path_buf())
    }
}

/// `file:///C:/dir/with%20space` — путь в виде URL (LibreOffice требует percent-кодирование).
fn file_url(path: &Path) -> String {
    let p = path.to_string_lossy().replace('\\', "/");
    let mut s = String::from("file://");
    if !p.starts_with('/') {
        s.push('/');
    }
    for b in p.bytes() {
        if b.is_ascii_alphanumeric() || b"-._~/:".contains(&b) {
            s.push(b as char);
        } else {
            s.push_str(&format!("%{b:02X}"));
        }
    }
    s
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn url_is_percent_encoded() {
        assert_eq!(file_url(Path::new(r"C:\Users\Иван Петров\lo")), "file:///C:/Users/%D0%98%D0%B2%D0%B0%D0%BD%20%D0%9F%D0%B5%D1%82%D1%80%D0%BE%D0%B2/lo");
        assert_eq!(file_url(Path::new("/tmp/a b#c")), "file:///tmp/a%20b%23c");
    }

    #[test]
    fn dotted_names_convert() {
        if find_soffice().is_none() {
            return;
        }
        let dir = tempfile::tempdir().unwrap();
        let src = dir.path().join("Отчёт 01.03.2024.txt");
        std::fs::write(&src, "hello").unwrap();
        let cancel = AtomicBool::new(false);
        let conv = Converter::new(dir.path(), &cancel);
        let out = conv.convert(&src, "docx", &dir.path().join("o")).unwrap();
        assert!(out.ends_with("Отчёт 01.03.2024.docx"));
    }

    #[test]
    fn cancel_kills_process() {
        let cancel = AtomicBool::new(true);
        let mut cmd = Command::new(if cfg!(windows) { "ping" } else { "sleep" });
        if cfg!(windows) {
            cmd.args(["-n", "30", "127.0.0.1"]);
        } else {
            cmd.arg("30");
        }
        let t = Instant::now();
        assert!(run_cmd(&mut cmd, &cancel, Duration::from_secs(60)).is_err());
        assert!(t.elapsed() < Duration::from_secs(5));
    }
}
