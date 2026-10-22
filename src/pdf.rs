//! Конвертация документов: LibreOffice (headless), а в Windows при его отсутствии —
//! установленный Microsoft Office через PowerShell/COM.

use anyhow::{Context, Result, anyhow, bail};
use std::path::{Path, PathBuf};
use std::process::Command;

fn find_soffice() -> Option<PathBuf> {
    let candidates = [
        "soffice",
        "libreoffice",
        r"C:\Program Files\LibreOffice\program\soffice.exe",
        r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
        "/Applications/LibreOffice.app/Contents/MacOS/soffice",
    ];
    for c in candidates {
        let p = PathBuf::from(c);
        if p.is_absolute() {
            if p.exists() {
                return Some(p);
            }
        } else if Command::new(&p).arg("--version").output().is_ok() {
            return Some(p);
        }
    }
    None
}

/// Конвертирует `input` в формат `fmt` ("pdf", "docx", …), результат кладёт в `out_dir`.
pub fn convert(input: &Path, fmt: &str, out_dir: &Path) -> Result<PathBuf> {
    let Some(soffice) = find_soffice() else {
        return convert_with_office(input, fmt, out_dir).map_err(|e| {
            anyhow!("Не найден LibreOffice (и Microsoft Office не сработал: {e}). Установите LibreOffice: https://www.libreoffice.org")
        });
    };
    // Отдельный профиль, чтобы не конфликтовать с уже запущенным LibreOffice.
    let profile = out_dir.join("lo-profile");
    let url = profile.to_string_lossy().replace('\\', "/");
    let url = format!("file://{}{}", if url.starts_with('/') { "" } else { "/" }, url);
    let status = Command::new(soffice)
        .arg("--headless")
        .arg(format!("-env:UserInstallation={url}"))
        .args(["--convert-to", fmt, "--outdir"])
        .arg(out_dir)
        .arg(input)
        .output()
        .context("не удалось запустить LibreOffice")?;
    if !status.status.success() {
        bail!("LibreOffice завершился с ошибкой: {}", String::from_utf8_lossy(&status.stderr));
    }
    let result = out_dir.join(input.file_stem().unwrap_or_default()).with_extension(fmt);
    if !result.exists() {
        bail!("LibreOffice не создал {}", result.display());
    }
    Ok(result)
}

/// Windows: конвертация через Microsoft Word/Excel (COM). На других системах недоступна.
fn convert_with_office(input: &Path, fmt: &str, out_dir: &Path) -> Result<PathBuf> {
    if !cfg!(windows) {
        bail!("только Windows");
    }
    let ext = input.extension().and_then(|e| e.to_str()).unwrap_or("").to_lowercase();
    let out = out_dir.join(input.file_stem().unwrap_or_default()).with_extension(fmt);
    let script = match (ext.as_str(), fmt) {
        ("doc" | "docx", "pdf" | "docx") => {
            let code = if fmt == "pdf" { 17 } else { 16 };
            format!(
                "$a=New-Object -ComObject Word.Application; $a.Visible=$false; $a.DisplayAlerts=0; \
                 try {{ $d=$a.Documents.Open($env:MF_IN,$false,$true); $d.SaveAs([ref]$env:MF_OUT,[ref]{code}); $d.Close($false) }} finally {{ $a.Quit() }}"
            )
        }
        ("xls" | "xlsx" | "xlsm" | "ods", "pdf" | "xlsx") => {
            let body = if fmt == "pdf" {
                "$w.ExportAsFixedFormat(0,$env:MF_OUT)"
            } else {
                "$w.SaveAs($env:MF_OUT,51)"
            };
            format!(
                "$a=New-Object -ComObject Excel.Application; $a.Visible=$false; $a.DisplayAlerts=$false; \
                 try {{ $w=$a.Workbooks.Open($env:MF_IN); {body}; $w.Close($false) }} finally {{ $a.Quit() }}"
            )
        }
        _ => bail!("неподдерживаемое преобразование"),
    };
    let r = Command::new("powershell")
        .args(["-NoProfile", "-NonInteractive", "-Command", &script])
        .env("MF_IN", input)
        .env("MF_OUT", &out)
        .output()
        .context("не удалось запустить PowerShell")?;
    if !r.status.success() || !out.exists() {
        bail!("{}", String::from_utf8_lossy(&r.stderr).trim());
    }
    Ok(out)
}
