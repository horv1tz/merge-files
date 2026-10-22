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
use std::ffi::OsString;
use std::path::PathBuf;
use std::sync::atomic::AtomicBool;
use std::sync::{Arc, OnceLock};
use std::sync::mpsc::channel;

const USAGE: &str = "Использование:
  merge-files                         — открыть окно
  merge-files файл1 файл2 …           — открыть окно с этими файлами
  merge-files -o РЕЗУЛЬТАТ файл1 …    — объединить без окна (тип по расширению: .xlsx, .docx или .pdf)
  merge-files --install-menu          — добавить пункт в контекстное меню проводника (Windows)
  merge-files --uninstall-menu        — убрать пункт из контекстного меню
  merge-files --version

В Windows exe — оконное приложение: в скриптах запускайте его через «start /wait», чтобы дождаться результата.";

/// Вывод сообщений. В Windows release-сборке у процесса нет консоли — один раз подключаемся к родительской.
/// Если консоли нет (запуск из проводника или планировщика), в окне показываем только `important` (ошибки, справку).
fn say(msg: &str, important: bool) {
    #[cfg(windows)]
    {
        use std::io::Write;
        use std::sync::{Mutex, OnceLock};
        unsafe extern "system" {
            fn AttachConsole(pid: u32) -> i32;
        }
        static CONSOLE: OnceLock<Option<Mutex<std::fs::File>>> = OnceLock::new();
        let console = CONSOLE.get_or_init(|| {
            // SAFETY: простой вызов Win32 без указателей; u32::MAX = ATTACH_PARENT_PROCESS.
            // Если консоль уже есть (отладочная сборка), вызов вернёт ошибку, а CONOUT$ всё равно откроется.
            unsafe { AttachConsole(u32::MAX) };
            std::fs::OpenOptions::new().write(true).open("CONOUT$").ok().map(Mutex::new)
        });
        match console {
            Some(f) => {
                if let Ok(mut f) = f.lock() {
                    let _ = writeln!(f, "{msg}");
                }
            }
            None if important => {
                rfd::MessageDialog::new().set_title("Merge Files").set_description(msg).show();
            }
            None => {}
        }
    }
    #[cfg(not(windows))]
    {
        let _ = important;
        eprintln!("{msg}");
    }
}

fn absolute(p: impl Into<PathBuf>) -> PathBuf {
    let p = p.into();
    std::path::absolute(&p).unwrap_or(p)
}

/// Консольный режим: `merge-files -o OUT in1 in2 ...`.
fn cli(args: Vec<OsString>) -> anyhow::Result<()> {
    let mut out = None;
    let mut inputs = vec![];
    let mut it = args.into_iter();
    while let Some(a) = it.next() {
        if a == "-o" {
            out = it.next().map(absolute);
        } else {
            inputs.push(absolute(a));
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
            let sheets = if kind == Kind::Excel {
                match excel::list_sheets(&p) {
                    Ok(s) => s,
                    // HTML/XML-выгрузка под видом .xls: листы определятся после конвертации.
                    Err(_) if job::sniff(&p).is_ok_and(|f| f == job::Sniffed::Text) => {
                        vec![excel::SheetInfo { name: "Данные".into(), hidden: false }]
                    }
                    Err(e) => return Err(e),
                }
            } else {
                vec![]
            };
            let all_sheets: Vec<String> = sheets.iter().map(|s| s.name.clone()).collect();
            // Как и в окне: скрытые листы по умолчанию не берём (если скрыты все — берём все).
            let mut selected: Vec<usize> = (0..sheets.len()).filter(|&i| !sheets[i].hidden).collect();
            if selected.is_empty() {
                selected = (0..sheets.len()).collect();
            }
            Ok(JobEntry { path: p, all_sheets, selected })
        })
        .collect::<anyhow::Result<Vec<_>>>()?;
    let job = Job { kind, target, entries, out, keep_source_styles: true };
    let cancel = AtomicBool::new(false);
    let outcome = job::run(&job, &cancel, &|m| say(m, false))?;
    for w in outcome.warnings {
        say(&format!("Предупреждение: {w}"), false);
    }
    say(&format!("Готово: {}", job.out.display()), false);
    Ok(())
}

fn finish(r: anyhow::Result<()>) -> eframe::Result {
    if let Err(e) = r {
        say(&format!("Ошибка: {e:#}"), true);
        std::process::exit(1);
    }
    Ok(())
}

fn main() -> eframe::Result {
    std::panic::set_hook(Box::new(|info| {
        let thread = std::thread::current();
        let line = format!("panic в потоке {:?}: {info}", thread.name().unwrap_or("?"));
        config::log(&line);
        if thread.name() == Some("main") {
            say(&format!("Внутренняя ошибка программы. Подробности записаны в журнал: {}\n{line}", config::log_path().display()), true);
        }
    }));

    let args: Vec<OsString> = std::env::args_os().skip(1).collect();
    match args.first().and_then(|a| a.to_str()) {
        Some("--help" | "-h" | "/?") => {
            say(USAGE, true);
            return Ok(());
        }
        Some("--version") => {
            say(&format!("Merge Files {}", env!("CARGO_PKG_VERSION")), true);
            return Ok(());
        }
        Some("--install-menu") => return finish(shell::install_menu().map(|_| say("Пункт меню добавлен", true))),
        Some("--uninstall-menu") => return finish(shell::uninstall_menu().map(|_| say("Пункт меню удалён", true))),
        Some(flag) if flag.starts_with("--") => {
            say(&format!("Неизвестный параметр «{flag}»\n\n{USAGE}"), true);
            std::process::exit(2);
        }
        _ if args.iter().any(|a| a == "-o") => return finish(cli(args)),
        _ => {}
    }

    // Остальные аргументы — файлы (запуск из контекстного меню).
    let initial: Vec<PathBuf> = args.into_iter().map(absolute).collect();
    // Принимать файлы от других экземпляров начинаем сразу, не дожидаясь окна: ответ нужен им в течение секунд.
    let ctx_slot: Arc<OnceLock<egui::Context>> = Arc::new(OnceLock::new());
    let incoming = match shell::claim(&initial) {
        shell::Role::Forwarded => return Ok(()),
        shell::Role::Primary(l) => {
            let (tx, rx) = channel();
            let slot = ctx_slot.clone();
            shell::serve(l, move |batch| {
                let _ = tx.send(batch);
                if let Some(ctx) = slot.get() {
                    ctx.request_repaint();
                }
            });
            Some(rx)
        }
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
            let _ = ctx_slot.set(cc.egui_ctx.clone());
            Ok(Box::new(app::App::new(cc, incoming, initial)))
        }),
    );
    if let Err(e) = &res {
        let msg = format!("Не удалось запустить окно программы: {e}");
        config::log(&msg);
        say(&msg, true);
    }
    res
}
