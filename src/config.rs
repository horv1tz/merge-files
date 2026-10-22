//! Мелкие настройки пользователя (формат `ключ=значение`, по одному в строке).

use std::path::PathBuf;

#[derive(Default, Clone)]
pub struct Config {
    /// Результат: "pdf" или пусто (тот же формат).
    pub target_pdf: bool,
    pub last_dir: Option<PathBuf>,
}

fn path() -> Option<PathBuf> {
    let base = if cfg!(windows) {
        std::env::var_os("APPDATA").map(PathBuf::from)?
    } else {
        std::env::var_os("XDG_CONFIG_HOME")
            .map(PathBuf::from)
            .or_else(|| std::env::var_os("HOME").map(|h| PathBuf::from(h).join(".config")))?
    };
    Some(base.join("MergeFiles").join("config.txt"))
}

impl Config {
    pub fn load() -> Config {
        let mut c = Config::default();
        let Some(text) = path().and_then(|p| std::fs::read_to_string(p).ok()) else { return c };
        for line in text.lines() {
            match line.split_once('=') {
                Some(("target", v)) => c.target_pdf = v.trim() == "pdf",
                Some(("last_dir", v)) if !v.trim().is_empty() => c.last_dir = Some(PathBuf::from(v.trim())),
                _ => {}
            }
        }
        c
    }

    pub fn save(&self) {
        let Some(p) = path() else { return };
        if let Some(dir) = p.parent() {
            let _ = std::fs::create_dir_all(dir);
        }
        let text = format!(
            "target={}\nlast_dir={}\n",
            if self.target_pdf { "pdf" } else { "same" },
            self.last_dir.as_ref().map(|d| d.to_string_lossy().into_owned()).unwrap_or_default()
        );
        let _ = std::fs::write(p, text);
    }
}

/// Папка для журнала ошибок.
pub fn log_path() -> PathBuf {
    let base = if cfg!(windows) {
        std::env::var_os("LOCALAPPDATA").map(PathBuf::from)
    } else {
        std::env::var_os("XDG_STATE_HOME").map(PathBuf::from)
    };
    base.unwrap_or_else(std::env::temp_dir).join("MergeFiles").join("error.log")
}

/// Дописывает строку в журнал ошибок (молча игнорирует сбои записи).
pub fn log(msg: &str) {
    use std::io::Write;
    let p = log_path();
    if let Some(dir) = p.parent() {
        let _ = std::fs::create_dir_all(dir);
    }
    if let Ok(mut f) = std::fs::OpenOptions::new().create(true).append(true).open(p) {
        let _ = writeln!(f, "{msg}");
    }
}
