//! Интеграция с проводником: единственный экземпляр приложения и пункт контекстного меню.
//!
//! Проводник запускает команду меню отдельно для каждого выделенного файла, поэтому
//! первый запущенный экземпляр слушает локальный порт, а остальные передают ему свои
//! пути и завершаются — все выбранные файлы оказываются в одном окне.

use anyhow::{Result, bail};
use std::io::{Read, Write};
use std::net::{SocketAddr, TcpListener, TcpStream};
use std::path::PathBuf;
use std::process::Command;
use std::time::Duration;

const ADDR: &str = "127.0.0.1:47831";
const MENU_KEY: &str = "MergeFiles";
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
    let addr: SocketAddr = ADDR.parse().unwrap();
    for _ in 0..30 {
        if let Ok(l) = TcpListener::bind(addr) {
            return Role::Primary(l);
        }
        if paths.is_empty() {
            return Role::Standalone;
        }
        if let Ok(mut s) = TcpStream::connect_timeout(&addr, Duration::from_millis(200)) {
            let text: String = paths.iter().map(|p| format!("{}\n", p.display())).collect();
            if s.write_all(text.as_bytes()).is_ok() {
                return Role::Forwarded;
            }
        }
        std::thread::sleep(Duration::from_millis(50));
    }
    Role::Standalone
}

/// Принимает пути от других экземпляров и передаёт пачками в `on_batch`.
pub fn serve(listener: TcpListener, on_batch: impl Fn(Vec<PathBuf>) + Send + 'static) {
    std::thread::spawn(move || {
        for stream in listener.incoming().flatten() {
            let mut stream = stream;
            let _ = stream.set_read_timeout(Some(Duration::from_secs(2)));
            let mut text = String::new();
            if stream.read_to_string(&mut text).is_ok() || !text.is_empty() {
                let batch: Vec<PathBuf> = text.lines().filter(|l| !l.is_empty()).map(PathBuf::from).collect();
                if !batch.is_empty() {
                    on_batch(batch);
                }
            }
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

fn key_for(ext: &str) -> String {
    format!(r"HKCU\Software\Classes\SystemFileAssociations\.{ext}\shell\{MENU_KEY}")
}

fn reg(args: &[&str]) -> Result<bool> {
    let mut cmd = Command::new("reg");
    hide_window(&mut cmd);
    Ok(cmd.args(args).output()?.status.success())
}

pub fn menu_supported() -> bool {
    cfg!(windows)
}

pub fn menu_installed() -> bool {
    cfg!(windows) && reg(&["query", &key_for("docx")]).unwrap_or(false)
}

/// Добавляет пункт «Объединить файлы…» в меню файлов Excel/Word (для текущего пользователя).
pub fn install_menu() -> Result<()> {
    if !cfg!(windows) {
        bail!("контекстное меню поддерживается только в Windows");
    }
    let exe = std::env::current_exe()?;
    let exe = exe.to_string_lossy();
    let command = format!("\"{exe}\" \"%1\"");
    let icon = format!("\"{exe}\",0");
    for ext in EXTS {
        let key = key_for(ext);
        let cmd_key = format!(r"{key}\command");
        let ok = reg(&["add", &key, "/ve", "/d", "Объединить файлы…", "/f"])?
            && reg(&["add", &key, "/v", "Icon", "/d", &icon, "/f"])?
            && reg(&["add", &cmd_key, "/ve", "/d", &command, "/f"])?;
        if !ok {
            bail!("не удалось записать в реестр (.{ext})");
        }
    }
    Ok(())
}

pub fn uninstall_menu() -> Result<()> {
    if !cfg!(windows) {
        bail!("контекстное меню поддерживается только в Windows");
    }
    for ext in EXTS {
        // Код возврата игнорируем: ключа может уже не быть.
        let _ = reg(&["delete", &key_for(ext), "/f"]);
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
                out.push_str(&format!("{digits:0>12}"));
                digits.clear();
            }
            out.push(c);
        }
    }
    if !digits.is_empty() {
        out.push_str(&format!("{digits:0>12}"));
    }
    out
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn second_instance_forwards_paths() {
        let Role::Primary(l) = claim(&[]) else { panic!("порт должен быть свободен") };
        let (tx, rx) = std::sync::mpsc::channel();
        serve(l, move |b| tx.send(b).unwrap());
        let paths = vec![PathBuf::from("/tmp/a b.docx"), PathBuf::from("/tmp/c.docx")];
        assert!(matches!(claim(&paths), Role::Forwarded));
        assert_eq!(rx.recv_timeout(Duration::from_secs(3)).unwrap(), paths);
    }

    #[test]
    fn natural_order() {
        let mut v = vec!["file10.docx", "File2.docx", "file1.docx"];
        v.sort_by_key(|s| natural_key(s));
        assert_eq!(v, ["file1.docx", "File2.docx", "file10.docx"]);
    }
}
