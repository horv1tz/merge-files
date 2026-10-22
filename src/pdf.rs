//! Конвертация через LibreOffice (headless).

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
    let soffice = find_soffice()
        .ok_or_else(|| anyhow!("LibreOffice не найден: он нужен для вывода в PDF и чтения .doc"))?;
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
