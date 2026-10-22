#![cfg_attr(all(not(debug_assertions), windows), windows_subsystem = "windows")]

mod app;
mod config;
mod excel;
mod job;
mod pdf;
mod shell;
mod word;

use eframe::egui;
use job::{Job, JobEntry, Kind, Target};
use std::path::PathBuf;
use std::sync::atomic::AtomicBool;
use std::sync::mpsc::channel;

const USAGE: &str = "Использование:
  merge-files                         — открыть окно
  merge-files файл1 файл2 …           — открыть окно с этими файлами
  merge-files -o РЕЗУЛЬТАТ файл1 …    — объединить без окна (тип по расширению: .xlsx, .docx или .pdf)
  merge-files --install-menu          — добавить пункт в контекстное меню проводника (Windows)
  merge-files --uninstall-menu        — убрать пункт из контекстного меню";

/// Вывод сообщений: в Windows release-сборке у процесса нет консоли — подключаемся к родительской.
fn say(msg: &str) {
    #[cfg(windows)]
    {
        use std::io::Write;
        unsafe extern "system" {
            fn AttachConsole(pid: u32) -> i32;
        }
        // SAFETY: простой вызов Win32 без указателей; (u32::MAX) = ATTACH_PARENT_PROCESS.
        if unsafe { AttachConsole(u32::MAX) } != 0 {
            if let Ok(mut f) = std::fs::OpenOptions::new().write(true).open("CONOUT$") {
                let _ = writeln!(f, "{msg}");
                return;
            }
        }
        // Консоли нет (запуск из проводника) — показываем окно сообщения.
        rfd::MessageDialog::new().set_title("Merge Files").set_description(msg).show();
    }
    #[cfg(not(windows))]
    eprintln!("{msg}");
}

fn absolute(p: PathBuf) -> PathBuf {
    std::path::absolute(&p).unwrap_or(p)
}

/// Консольный режим: `merge-files -o OUT in1 in2 ...`.
fn cli(args: Vec<String>) -> anyhow::Result<()> {
    let mut out = None;
    let mut inputs = vec![];
    let mut it = args.into_iter();
    while let Some(a) = it.next() {
        if a == "-o" {
            out = it.next().map(|p| absolute(PathBuf::from(p)));
        } else {
            inputs.push(absolute(PathBuf::from(a)));
        }
    }
    let out = out.ok_or_else(|| anyhow::anyhow!("не указан файл результата (-o)\n\n{USAGE}"))?;
    let kind = inputs
        .first()
        .and_then(|p| Kind::of(p))
        .ok_or_else(|| anyhow::anyhow!("нужны файлы .xlsx/.xls/.ods или .docx/.doc"))?;
    if let Some(bad) = inputs.iter().find(|p| Kind::of(p) != Some(kind)) {
        anyhow::bail!("«{}» — другой тип файла, чем первый", bad.display());
    }
    let target = if out.extension().is_some_and(|e| e.eq_ignore_ascii_case("pdf")) { Target::Pdf } else { Target::Same };
    let entries = inputs
        .into_iter()
        .map(|p| {
            let all_sheets: Vec<String> =
                if kind == Kind::Excel { excel::list_sheets(&p)?.into_iter().map(|s| s.name).collect() } else { vec![] };
            let selected = (0..all_sheets.len()).collect();
            Ok(JobEntry { path: p, all_sheets, selected })
        })
        .collect::<anyhow::Result<Vec<_>>>()?;
    let job = Job { kind, target, entries, out };
    let cancel = AtomicBool::new(false);
    let outcome = job::run(&job, &cancel, &|m| say(m))?;
    for w in outcome.warnings {
        say(&format!("Предупреждение: {w}"));
    }
    say(&format!("Готово: {}", job.out.display()));
    Ok(())
}

fn finish(r: anyhow::Result<()>) -> eframe::Result {
    if let Err(e) = r {
        say(&format!("Ошибка: {e:#}"));
        std::process::exit(1);
    }
    Ok(())
}

fn main() -> eframe::Result {
    std::panic::set_hook(Box::new(|info| {
        let thread = std::thread::current();
        config::log(&format!("panic в потоке {:?}: {info}", thread.name().unwrap_or("?")));
    }));

    let args: Vec<String> = std::env::args().skip(1).collect();
    match args.first().map(String::as_str) {
        Some("--help" | "-h" | "/?") => {
            say(USAGE);
            return Ok(());
        }
        Some("--install-menu") => return finish(shell::install_menu().map(|_| say("Пункт меню добавлен"))),
        Some("--uninstall-menu") => return finish(shell::uninstall_menu().map(|_| say("Пункт меню удалён"))),
        _ if args.iter().any(|a| a == "-o") => return finish(cli(args)),
        _ => {}
    }

    // Остальные аргументы — файлы (запуск из контекстного меню).
    let initial: Vec<PathBuf> = args.into_iter().map(|a| absolute(PathBuf::from(a))).collect();
    let listener = match shell::claim(&initial) {
        shell::Role::Forwarded => return Ok(()),
        shell::Role::Primary(l) => Some(l),
        shell::Role::Standalone => None,
    };
    let mut viewport = egui::ViewportBuilder::default()
        .with_title("Merge Files")
        .with_inner_size([580.0, 600.0])
        .with_min_inner_size([460.0, 420.0])
        .with_drag_and_drop(true);
    if let Ok(icon) = eframe::icon_data::from_png_bytes(include_bytes!("../assets/icon_window.png")) {
        viewport = viewport.with_icon(icon);
    }
    let options = eframe::NativeOptions { viewport, ..Default::default() };
    let res = eframe::run_native(
        "Merge Files",
        options,
        Box::new(move |cc| {
            let incoming = listener.map(|l| {
                let (tx, rx) = channel();
                let ctx = cc.egui_ctx.clone();
                shell::serve(l, move |batch| {
                    let _ = tx.send(batch);
                    ctx.request_repaint();
                });
                rx
            });
            Ok(Box::new(app::App::new(cc, incoming, initial)))
        }),
    );
    if let Err(e) = &res {
        let msg = format!("Не удалось запустить окно программы: {e}");
        config::log(&msg);
        say(&msg);
    }
    res
}
