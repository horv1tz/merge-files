//! Интеграция с проводником: единственный экземпляр приложения и пункт контекстного меню.
//!
//! Проводник запускает команду меню отдельно для каждого выделенного файла, поэтому
//! первый запущенный экземпляр слушает локальный порт, а остальные передают ему свои
//! пути и завершаются — все выбранные файлы оказываются в одном окне.
//! Протокол: клиент шлёт строку-заголовок и пути (по одному в строке), закрывает запись
//! и ждёт ответ `OK`; без ответа файлы считаются не доставленными.

use anyhow::{Result, bail};
use std::io::{Read, Write};
use std::net::{Shutdown, SocketAddr, TcpListener, TcpStream};
use std::path::{Path, PathBuf};
use std::process::Command;
use std::time::Duration;

const ADDR: &str = "127.0.0.1:47831";
const MAGIC: &str = "MERGE-FILES-1";
const MENU_KEY: &str = "MergeFiles";
pub const MENU_LABEL: &str = "Объединить файлы…";
pub const EXTS: [&str; 6] = ["xlsx", "xlsm", "xls", "ods", "docx", "doc"];

pub enum Role {
    /// Это первый экземпляр: принимает файлы от остальных.
    Primary(TcpListener),
    /// Пути переданы уже работающему экземпляру, можно выходить.
    Forwarded,
    /// Порт недоступен — работаем сами по себе.
    Standalone,
}

/// Определяет роль экземпляра; при запуске с файлами ждёт первого экземпляра до ~1.5 с.
pub fn claim(paths: &[PathBuf]) -> Role {
    claim_at(ADDR.parse().unwrap(), paths)
}

fn claim_at(addr: SocketAddr, paths: &[PathBuf]) -> Role {
    for _ in 0..30 {
        if let Ok(l) = TcpListener::bind(addr) {
            return Role::Primary(l);
        }
        if paths.is_empty() {
            return Role::Standalone;
        }
        match forward(&addr, paths) {
            Forward::Delivered => return Role::Forwarded,
            // Порт занят чужой программой (или зависшим экземпляром) — не ждём впустую.
            Forward::Foreign => return Role::Standalone,
            Forward::Refused => {}
        }
        std::thread::sleep(Duration::from_millis(50));
    }
    Role::Standalone
}

enum Forward {
    Delivered,
    /// Подключиться не удалось (экземпляр ещё не слушает).
    Refused,
    /// Подключились, но ответа протокола нет.
    Foreign,
}

fn forward(addr: &SocketAddr, paths: &[PathBuf]) -> Forward {
    let Ok(mut s) = TcpStream::connect_timeout(addr, Duration::from_millis(300)) else { return Forward::Refused };
    // Первый экземпляр может быть занят запуском окна — ждём ответ подольше.
    let _ = s.set_read_timeout(Some(Duration::from_secs(10)));
    let _ = s.set_write_timeout(Some(Duration::from_secs(5)));
    let mut text = format!("{MAGIC}\n");
    for p in paths {
        text.push_str(&p.to_string_lossy());
        text.push('\n');
    }
    if s.write_all(text.as_bytes()).is_err() {
        return Forward::Foreign;
    }
    let _ = s.shutdown(Shutdown::Write);
    let mut reply = String::new();
    if s.read_to_string(&mut reply).is_ok() && reply.trim() == "OK" { Forward::Delivered } else { Forward::Foreign }
}

/// Принимает пути от других экземпляров и передаёт пачками в `on_batch`.
pub fn serve(listener: TcpListener, on_batch: impl Fn(Vec<PathBuf>) + Send + 'static) {
    std::thread::spawn(move || {
        for mut stream in listener.incoming().flatten() {
            let _ = stream.set_read_timeout(Some(Duration::from_secs(2)));
            let mut text = String::new();
            let _ = stream.read_to_string(&mut text);
            let mut lines = text.lines();
            if lines.next() != Some(MAGIC) {
                continue;
            }
            let batch: Vec<PathBuf> = lines.filter(|l| !l.is_empty()).map(PathBuf::from).collect();
            if !batch.is_empty() {
                on_batch(batch);
            }
            let _ = stream.write_all(b"OK");
        }
    });
}

/// Скрывает консольное окно дочернего процесса в Windows.
pub fn hide_window(cmd: &mut Command) -> &mut Command {
    #[cfg(windows)]
    {
        use std::os::windows::process::CommandExt;
        cmd.creation_flags(0x0800_0000); // CREATE_NO_WINDOW
    }
    cmd
}

// ───────────────────────── контекстное меню ─────────────────────────

#[derive(Clone, Copy, PartialEq, Eq, Debug)]
pub enum MenuState {
    NotInstalled,
    /// Пункт есть и указывает на эту копию программы.
    Installed,
    /// Пункт есть, но указывает на другой (перенесённый/удалённый) exe.
    Stale,
}

fn key_for(ext: &str) -> String {
    format!(r"HKCU\Software\Classes\SystemFileAssociations\.{ext}\shell\{MENU_KEY}")
}

fn reg(args: &[&str]) -> Result<std::process::Output> {
    let mut cmd = Command::new("reg");
    hide_window(&mut cmd);
    Ok(cmd.args(args).output()?)
}

pub fn menu_supported() -> bool {
    cfg!(windows)
}

fn command_line(exe: &Path) -> String {
    format!("\"{}\" \"%1\"", exe.display())
}

/// Короткий ASCII-идентификатор exe (FNV-1a): путь с кириллицей нельзя надёжно прочитать из вывода reg.exe.
fn exe_id(exe: &Path) -> String {
    let mut h: u64 = 0xcbf2_9ce4_8422_2325;
    for b in exe.to_string_lossy().to_lowercase().bytes() {
        h ^= u64::from(b);
        h = h.wrapping_mul(0x0100_0000_01b3);
    }
    format!("{h:016x}")
}

/// Декодирует вывод консольной программы (в Windows — в кодировке OEM, например cp866).
pub fn decode_console(bytes: &[u8]) -> String {
    #[cfg(windows)]
    {
        unsafe extern "system" {
            fn MultiByteToWideChar(cp: u32, flags: u32, src: *const u8, n: i32, dst: *mut u16, cap: i32) -> i32;
        }
        if bytes.is_empty() || bytes.len() > i32::MAX as usize {
            return String::new();
        }
        let n = bytes.len() as i32;
        // SAFETY: буфер достаточного размера (число UTF-16 единиц не больше числа байт); CP_OEMCP = 1.
        unsafe {
            let mut buf = vec![0u16; bytes.len()];
            let len = MultiByteToWideChar(1, 0, bytes.as_ptr(), n, buf.as_mut_ptr(), n);
            if len > 0 {
                buf.truncate(len as usize);
                return String::from_utf16_lossy(&buf);
            }
        }
    }
    String::from_utf8_lossy(bytes).into_owned()
}

pub fn menu_state() -> MenuState {
    if !cfg!(windows) {
        return MenuState::NotInstalled;
    }
    let Ok(exe) = std::env::current_exe() else { return MenuState::NotInstalled };
    let want = exe_id(&exe);
    let (mut ok, mut any) = (0, false);
    for ext in EXTS {
        let key = key_for(ext);
        let Ok(out) = reg(&["query", &key, "/v", "ExeId"]) else { continue };
        if out.status.success() {
            any = true;
            // Значение — ASCII, кодировка вывода не важна.
            if String::from_utf8_lossy(&out.stdout).split_whitespace().any(|w| w == want) {
                ok += 1;
            }
        } else if reg(&["query", &key]).is_ok_and(|o| o.status.success()) {
            any = true;
        }
    }
    if ok == EXTS.len() {
        MenuState::Installed
    } else if any {
        MenuState::Stale
    } else {
        MenuState::NotInstalled
    }
}

/// Добавляет пункт «Объединить файлы…» в меню файлов Excel/Word (для текущего пользователя).
pub fn install_menu() -> Result<()> {
    if !cfg!(windows) {
        bail!("контекстное меню поддерживается только в Windows");
    }
    let exe = std::env::current_exe()?;
    let command = command_line(&exe);
    let icon = format!("\"{}\",0", exe.display());
    for ext in EXTS {
        let key = key_for(ext);
        let cmd_key = format!(r"{key}\command");
        let ok = reg(&["add", &key, "/ve", "/d", MENU_LABEL, "/f"])?.status.success()
            && reg(&["add", &key, "/v", "Icon", "/d", &icon, "/f"])?.status.success()
            && reg(&["add", &key, "/v", "ExeId", "/d", &exe_id(&exe), "/f"])?.status.success()
            && reg(&["add", &cmd_key, "/ve", "/d", &command, "/f"])?.status.success();
        if !ok {
            bail!("не удалось записать пункт меню в реестр (.{ext})");
        }
    }
    if menu_state() != MenuState::Installed {
        bail!("пункт меню записан, но проверка не прошла");
    }
    Ok(())
}

pub fn uninstall_menu() -> Result<()> {
    if !cfg!(windows) {
        bail!("контекстное меню поддерживается только в Windows");
    }
    for ext in EXTS {
        let key = key_for(ext);
        // Не удалось удалить, но ключ ещё есть — ошибка; ключа уже нет — всё в порядке.
        if !reg(&["delete", &key, "/f"])?.status.success() && reg(&["query", &key])?.status.success() {
            bail!("не удалось удалить пункт меню (.{ext})");
        }
    }
    Ok(())
}

/// Ключ естественной сортировки: числа в именах сравниваются как числа.
pub fn natural_key(name: &str) -> String {
    let mut out = String::new();
    let mut digits = String::new();
    for c in name.to_lowercase().chars() {
        if c.is_ascii_digit() {
            digits.push(c);
        } else {
            if !digits.is_empty() {
                out.push_str(&format!("{digits:0>20}"));
                digits.clear();
            }
            out.push(c);
        }
    }
    if !digits.is_empty() {
        out.push_str(&format!("{digits:0>20}"));
    }
    out
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn second_instance_forwards_paths_with_ack() {
        // Свободный порт, чтобы тест не зависел от запущенной программы.
        let l = TcpListener::bind("127.0.0.1:0").unwrap();
        let addr = l.local_addr().unwrap();
        let (tx, rx) = std::sync::mpsc::channel();
        serve(l, move |b| tx.send(b).unwrap());
        let paths = vec![PathBuf::from("/tmp/a b.docx"), PathBuf::from("/tmp/c.docx")];
        assert!(matches!(claim_at(addr, &paths), Role::Forwarded));
        assert_eq!(rx.recv_timeout(Duration::from_secs(3)).unwrap(), paths);
    }

    #[test]
    fn natural_order() {
        let mut v = vec!["file10.docx", "File2.docx", "file1.docx"];
        v.sort_by_key(|s| natural_key(s));
        assert_eq!(v, ["file1.docx", "File2.docx", "file10.docx"]);
        assert!(natural_key("a9999999999999.x") < natural_key("a10000000000000.x"));
    }

    #[test]
    fn exe_id_is_stable_ascii() {
        let id = exe_id(Path::new(r"C:\Users\Иван\Merge Files.exe"));
        assert_eq!(id.len(), 16);
        assert!(id.bytes().all(|b| b.is_ascii_hexdigit()));
        assert_eq!(id, exe_id(Path::new(r"c:\users\иван\merge files.exe")));
    }
}
