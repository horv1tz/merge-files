//! Окно приложения (egui).

use crate::config::{self, Config};
use crate::job::{self, Job, JobEntry, Kind, Outcome, Target};
use crate::shell::{self, MenuState};
use crate::{excel, pdf};
use eframe::egui::{self, Align, Color32, Layout, RichText};
use std::path::{Path, PathBuf};
use std::sync::Arc;
use std::sync::atomic::{AtomicBool, Ordering};
use std::sync::mpsc::{Receiver, TryRecvError, channel};
use std::thread::JoinHandle;
use std::time::{Duration, Instant};

struct SheetItem {
    name: String,
    hidden: bool,
    selected: bool,
}

struct Entry {
    path: PathBuf,
    sheets: Vec<SheetItem>,
}

enum Status {
    Idle,
    Busy(String),
    Done { path: PathBuf },
    Error(String),
}

enum Msg {
    Progress(String),
    Done(Result<Outcome, String>),
}

struct Worker {
    rx: Receiver<Msg>,
    cancel: Arc<AtomicBool>,
    handle: Option<JoinHandle<()>>,
}

pub struct App {
    entries: Vec<Entry>,
    kind: Option<Kind>,
    target: Target,
    status: Status,
    warnings: Vec<String>,
    worker: Option<Worker>,
    incoming: Option<Receiver<Vec<PathBuf>>>,
    /// Файлы, пришедшие во время работы (проводник, перетаскивание): `bool` — отсортировать пачку по имени.
    pending: Vec<(Vec<PathBuf>, bool)>,
    menu: Option<MenuState>,
    menu_rx: Option<Receiver<Result<MenuState, String>>>,
    show_settings: bool,
    config: Config,
    last_output: Option<PathBuf>,
}

impl App {
    pub fn new(cc: &eframe::CreationContext<'_>, incoming: Option<Receiver<Vec<PathBuf>>>, initial: Vec<PathBuf>) -> Self {
        setup_fonts(&cc.egui_ctx);
        let config = Config::load();
        let mut app = App {
            entries: vec![],
            kind: None,
            target: if config.target_pdf { Target::Pdf } else { Target::Same },
            status: Status::Idle,
            warnings: vec![],
            worker: None,
            incoming,
            pending: vec![],
            menu: None,
            menu_rx: None,
            show_settings: false,
            config,
            last_output: None,
        };
        if !initial.is_empty() {
            app.pending.push((initial, true));
        }
        app
    }

    fn busy(&self) -> bool {
        self.worker.is_some()
    }

    fn error(&mut self, msg: impl Into<String>) {
        self.status = Status::Error(msg.into());
    }

    // ───────────── список файлов ─────────────

    /// Добавляет файлы; возвращает сообщения об отклонённых файлах.
    fn add(&mut self, paths: Vec<PathBuf>, sort_batch: bool) {
        let first_new = self.entries.len();
        let mut problems: Vec<String> = vec![];
        for path in paths {
            let name = path.file_name().unwrap_or_default().to_string_lossy().into_owned();
            let Some(k) = Kind::of(&path) else {
                problems.push(format!("«{name}»: неподдерживаемый тип файла"));
                continue;
            };
            if self.kind.is_some_and(|cur| cur != k) {
                problems.push(format!("«{name}»: нельзя смешивать Excel и Word в одном объединении"));
                continue;
            }
            if self.entries.iter().any(|e| job::same_file(&e.path, &path)) {
                problems.push(format!("«{name}» уже в списке"));
                continue;
            }
            let sheets = if k == Kind::Excel {
                match excel::list_sheets(&path) {
                    Ok(s) if s.is_empty() => {
                        problems.push(format!("«{name}»: в книге нет листов"));
                        continue;
                    }
                    Ok(s) => s.into_iter().map(|i| SheetItem { selected: !i.hidden, name: i.name, hidden: i.hidden }).collect(),
                    Err(e) => {
                        problems.push(format!("«{name}»: {e:#}"));
                        continue;
                    }
                }
            } else {
                vec![]
            };
            self.kind = Some(k);
            self.entries.push(Entry { path, sheets });
        }
        if sort_batch {
            self.entries[first_new..]
                .sort_by_key(|e| shell::natural_key(&e.path.file_name().unwrap_or_default().to_string_lossy()));
        }
        if self.entries.is_empty() {
            self.kind = None;
        }
        if !problems.is_empty() {
            self.error(problems.join("\n"));
        }
    }

    fn ready(&self) -> bool {
        match self.kind {
            Some(Kind::Excel) => self.entries.iter().any(|e| e.sheets.iter().any(|s| s.selected)),
            Some(Kind::Word) => !self.entries.is_empty(),
            None => false,
        }
    }

    // ───────────── фоновые события ─────────────

    fn poll(&mut self, ctx: &egui::Context) {
        // Файлы из проводника.
        let mut got = false;
        if let Some(rx) = &self.incoming {
            while let Ok(b) = rx.try_recv() {
                self.pending.push((b, true));
                got = true;
            }
        }
        if got {
            ctx.send_viewport_cmd(egui::ViewportCommand::Minimized(false));
            ctx.send_viewport_cmd(egui::ViewportCommand::Focus);
        }
        // Перетаскивание в окно.
        let dropped: Vec<PathBuf> = ctx.input(|i| i.raw.dropped_files.iter().map(|f| f.path().to_path_buf()).collect());
        if !dropped.is_empty() {
            self.pending.push((dropped, false));
        }
        if !self.busy() && !self.pending.is_empty() {
            for (batch, sort) in std::mem::take(&mut self.pending) {
                self.add(batch, sort);
            }
        }

        // Ход работы.
        let mut finished: Option<Result<Outcome, String>> = None;
        if let Some(w) = &self.worker {
            loop {
                match w.rx.try_recv() {
                    Ok(Msg::Progress(m)) => self.status = Status::Busy(m),
                    Ok(Msg::Done(r)) => {
                        finished = Some(r);
                        break;
                    }
                    Err(TryRecvError::Empty) => break,
                    Err(TryRecvError::Disconnected) => {
                        finished = Some(Err(format!("Работа неожиданно прервалась. Журнал: {}", config::log_path().display())));
                        break;
                    }
                }
            }
        }
        if let Some(r) = finished {
            if let Some(mut w) = self.worker.take() {
                if let Some(h) = w.handle.take() {
                    let _ = h.join();
                }
            }
            match r {
                Ok(o) => {
                    self.warnings = o.warnings;
                    let out = self.last_output.take();
                    self.status = match out {
                        Some(path) => Status::Done { path },
                        None => Status::Idle,
                    };
                }
                Err(e) => {
                    self.warnings.clear();
                    self.error(e);
                }
            }
        }

        // Пункт меню.
        if let Some(rx) = &self.menu_rx {
            match rx.try_recv() {
                Ok(Ok(state)) => {
                    self.menu = Some(state);
                    self.menu_rx = None;
                }
                Ok(Err(e)) => {
                    self.menu = Some(shell::menu_state());
                    self.menu_rx = None;
                    self.error(e);
                }
                Err(TryRecvError::Empty) => {}
                Err(TryRecvError::Disconnected) => self.menu_rx = None,
            }
        }
    }

    // ───────────── запуск объединения ─────────────

    fn start_merge(&mut self, ctx: &egui::Context) {
        let Some(kind) = self.kind else { return };
        let paths: Vec<PathBuf> = self.entries.iter().map(|e| e.path.clone()).collect();
        if job::needs_converter(self.target, &paths) && !pdf::available() {
            self.error(pdf::MISSING_HINT);
            return;
        }
        let ext = self.target.ext(kind);
        let dir = self
            .config
            .last_dir
            .clone()
            .filter(|d| d.is_dir())
            .or_else(|| paths.first().and_then(|p| p.parent().map(Path::to_path_buf)));
        let mut dlg = rfd::FileDialog::new().set_title("Сохранить результат").set_file_name(format!("merged.{ext}")).add_filter(ext, &[ext]);
        if let Some(d) = dir {
            dlg = dlg.set_directory(d);
        }
        let Some(mut out) = dlg.save_file() else { return };
        if !out.extension().and_then(|e| e.to_str()).is_some_and(|e| e.eq_ignore_ascii_case(ext)) {
            out = PathBuf::from(format!("{}.{ext}", out.display()));
        }
        self.config.last_dir = out.parent().map(Path::to_path_buf);
        self.config.target_pdf = self.target == Target::Pdf;
        self.config.save();

        let entries: Vec<JobEntry> = self
            .entries
            .iter()
            .map(|e| JobEntry {
                path: e.path.clone(),
                all_sheets: e.sheets.iter().map(|s| s.name.clone()).collect(),
                selected: e.sheets.iter().enumerate().filter(|(_, s)| s.selected).map(|(i, _)| i).collect(),
            })
            .collect();
        let job = Job { kind, target: self.target, entries, out: out.clone() };
        if let Err(e) = job::validate(&job) {
            self.error(format!("{e:#}"));
            return;
        }

        let (tx, rx) = channel();
        let cancel = Arc::new(AtomicBool::new(false));
        let cancel2 = cancel.clone();
        let ctx = ctx.clone();
        let handle = std::thread::spawn(move || {
            let progress_tx = tx.clone();
            let progress_ctx = ctx.clone();
            let progress = move |m: &str| {
                let _ = progress_tx.send(Msg::Progress(m.to_string()));
                progress_ctx.request_repaint();
            };
            let res = std::panic::catch_unwind(std::panic::AssertUnwindSafe(|| job::run(&job, &cancel2, &progress)));
            let msg = match res {
                Ok(Ok(o)) => Ok(o),
                Ok(Err(e)) => Err(format!("{e:#}")),
                Err(_) => Err(format!("Внутренняя ошибка программы. Подробности в журнале: {}", config::log_path().display())),
            };
            let _ = tx.send(Msg::Done(msg));
            ctx.request_repaint();
        });
        self.last_output = Some(out);
        self.warnings.clear();
        self.status = Status::Busy("Подготовка…".into());
        self.worker = Some(Worker { rx, cancel, handle: Some(handle) });
    }

    fn toggle_menu(&mut self, install: bool) {
        let (tx, rx) = channel();
        self.menu_rx = Some(rx);
        std::thread::spawn(move || {
            let r = if install { shell::install_menu() } else { shell::uninstall_menu() };
            let _ = tx.send(r.map(|_| shell::menu_state()).map_err(|e| format!("{e:#}")));
        });
    }

    fn refresh_menu(&mut self) {
        let (tx, rx) = channel();
        self.menu_rx = Some(rx);
        std::thread::spawn(move || {
            let _ = tx.send(Ok(shell::menu_state()));
        });
    }
}

// ───────────── отрисовка ─────────────

impl eframe::App for App {
    fn ui(&mut self, ui: &mut egui::Ui, _frame: &mut eframe::Frame) {
        let ctx = ui.ctx().clone();
        self.poll(&ctx);
        let busy = self.busy();
        let dark = ui.visuals().dark_mode;
        let ok_color = if dark { Color32::LIGHT_GREEN } else { Color32::from_rgb(0, 110, 40) };
        let err_color = ui.visuals().error_fg_color;

        // Нижние панели должны добавляться ДО центральной — иначе для них не останется места.
        egui::Panel::bottom("controls").show(ui, |ui| {
            ui.add_space(6.0);
            let paths: Vec<PathBuf> = self.entries.iter().map(|e| e.path.clone()).collect();
            let missing_converter = self.kind.is_some() && job::needs_converter(self.target, &paths) && !pdf::available();
            ui.horizontal(|ui| {
                ui.label("Результат:");
                let same = match self.kind {
                    Some(Kind::Excel) => "Excel (.xlsx)",
                    Some(Kind::Word) => "Word (.docx)",
                    None => "Тот же формат",
                };
                ui.add_enabled_ui(!busy, |ui| {
                    ui.radio_value(&mut self.target, Target::Same, same);
                    ui.radio_value(&mut self.target, Target::Pdf, "PDF");
                });
            });
            ui.add_space(4.0);
            ui.horizontal(|ui| {
                let can = self.ready() && !busy;
                if ui.add_enabled(can, egui::Button::new("Объединить…").min_size([160.0, 32.0].into())).clicked() {
                    self.start_merge(&ctx);
                }
                if let Some(w) = &self.worker {
                    ui.spinner();
                    if ui.button("Отмена").clicked() {
                        w.cancel.store(true, Ordering::Relaxed);
                    }
                }
            });
            if missing_converter {
                ui.add_space(4.0);
                ui.add(egui::Label::new(RichText::new(pdf::MISSING_HINT).color(err_color)).wrap());
            }
            ui.add_space(4.0);
            egui::ScrollArea::vertical().max_height(130.0).auto_shrink([false, true]).show(ui, |ui| {
                match &self.status {
                    Status::Idle => {}
                    Status::Busy(m) => {
                        ui.add(egui::Label::new(m.as_str()).wrap());
                    }
                    Status::Error(m) => {
                        ui.add(egui::Label::new(RichText::new(m.as_str()).color(err_color)).wrap());
                    }
                    Status::Done { path } => {
                        ui.add(egui::Label::new(RichText::new(format!("Готово: {}", path.display())).color(ok_color)).wrap());
                        ui.horizontal(|ui| {
                            if ui.button("Открыть файл").clicked() {
                                open_path(path, false);
                            }
                            if ui.button("Показать в папке").clicked() {
                                open_path(path, true);
                            }
                        });
                    }
                }
                if !self.warnings.is_empty() {
                    ui.add_space(4.0);
                    egui::CollapsingHeader::new(RichText::new(format!("⚠ Предупреждения ({})", self.warnings.len())).strong())
                        .default_open(true)
                        .show(ui, |ui| {
                            for w in &self.warnings {
                                ui.add(egui::Label::new(format!("• {w}")).wrap());
                            }
                        });
                }
            });
            ui.add_space(6.0);
        });

        egui::CentralPanel::default().show(ui, |ui| {
            ui.heading("Объединение файлов");
            ui.label("Листы Excel или документы Word — в один файл того же типа или в PDF.");
            ui.add_space(8.0);

            ui.horizontal(|ui| {
                if ui.add_enabled(!busy, egui::Button::new("➕ Добавить файлы…")).clicked() {
                    let mut dlg = rfd::FileDialog::new().add_filter("Excel / Word", &["xlsx", "xlsm", "xls", "ods", "docx", "doc"]);
                    if let Some(d) = self.config.last_dir.clone().filter(|d| d.is_dir()) {
                        dlg = dlg.set_directory(d);
                    }
                    if let Some(p) = dlg.pick_files() {
                        self.pending.push((p, false));
                    }
                }
                if ui.add_enabled(!busy && !self.entries.is_empty(), egui::Button::new("Очистить")).clicked() {
                    self.entries.clear();
                    self.kind = None;
                    self.status = Status::Idle;
                    self.warnings.clear();
                }
                ui.with_layout(Layout::right_to_left(Align::Center), |ui| {
                    if ui.button("⚙ Настройки").clicked() {
                        self.show_settings = true;
                        if self.menu_rx.is_none() {
                            self.refresh_menu();
                        }
                    }
                });
            });
            if self.entries.is_empty() {
                ui.add_space(4.0);
                ui.weak("Перетащите файлы в окно или нажмите «Добавить файлы…»");
            }
            ui.separator();

            let mut action: Option<(usize, i8)> = None; // -1 вверх, 1 вниз, 0 удалить
            egui::ScrollArea::vertical().auto_shrink([false, false]).show(ui, |ui| {
                ui.add_enabled_ui(!busy, |ui| {
                    let n = self.entries.len();
                    for (i, e) in self.entries.iter_mut().enumerate() {
                        ui.horizontal(|ui| {
                            let name = e.path.file_name().unwrap_or_default().to_string_lossy().into_owned();
                            // Имя занимает всё свободное место слева, кнопки остаются справа и не уезжают за край.
                            let w = (ui.available_width() - 130.0).max(60.0);
                            ui.scope(|ui| {
                                ui.set_width(w);
                                ui.add(egui::Label::new(RichText::new(format!("{}. {name}", i + 1)).strong()).truncate())
                                    .on_hover_text(e.path.display().to_string());
                            });
                            if ui.add_enabled(i > 0, egui::Button::new("⬆")).on_hover_text("Выше").clicked() {
                                action = Some((i, -1));
                            }
                            if ui.add_enabled(i + 1 < n, egui::Button::new("⬇")).on_hover_text("Ниже").clicked() {
                                action = Some((i, 1));
                            }
                            if ui.button("✖").on_hover_text("Убрать из списка").clicked() {
                                action = Some((i, 0));
                            }
                        });
                        if !e.sheets.is_empty() {
                            ui.indent(("sheets", i), |ui| {
                                for s in &mut e.sheets {
                                    let label = if s.hidden { format!("{} (скрытый)", s.name) } else { s.name.clone() };
                                    ui.checkbox(&mut s.selected, label);
                                }
                            });
                        }
                    }
                });
            });
            if let Some((i, a)) = action {
                match a {
                    -1 if i > 0 => self.entries.swap(i, i - 1),
                    1 if i + 1 < self.entries.len() => self.entries.swap(i, i + 1),
                    0 => {
                        self.entries.remove(i);
                        if self.entries.is_empty() {
                            self.kind = None;
                        }
                    }
                    _ => {}
                }
            }
        });

        self.settings_window(&ctx);
    }

    fn on_exit(&mut self) {
        // Закрыли окно во время работы: отменяем и даём конвертеру до 3 с, чтобы убрать за собой.
        if let Some(w) = &mut self.worker {
            w.cancel.store(true, Ordering::Relaxed);
            let t = Instant::now();
            if let Some(h) = w.handle.take() {
                while !h.is_finished() && t.elapsed() < Duration::from_secs(3) {
                    std::thread::sleep(Duration::from_millis(50));
                }
            }
        }
    }
}

impl App {
    fn settings_window(&mut self, ctx: &egui::Context) {
        let mut open = self.show_settings;
        egui::Window::new("Настройки").open(&mut open).collapsible(false).resizable(false).show(ctx, |ui| {
            ui.strong("Интеграция с проводником");
            let mut on = self.menu == Some(MenuState::Installed);
            let enabled = shell::menu_supported() && self.menu_rx.is_none() && self.menu.is_some();
            let resp = ui.add_enabled(enabled, egui::Checkbox::new(&mut on, format!("Пункт «{}» в контекстном меню", shell::MENU_LABEL)));
            if resp.changed() {
                self.toggle_menu(on);
            }
            if self.menu_rx.is_some() {
                ui.horizontal(|ui| {
                    ui.spinner();
                    ui.weak("Применяю…");
                });
            }
            if self.menu == Some(MenuState::Stale) {
                ui.colored_label(ui.visuals().warn_fg_color, "Пункт указывает на другую копию программы — включите галочку заново.");
            }
            if shell::menu_supported() {
                ui.weak("Файлы Excel и Word → ПКМ → «Объединить файлы…». Права администратора не нужны.");
                ui.weak("Windows 11: пункт в «Показать дополнительные параметры». Для 15+ выделенных файлов Windows пункт скрывает.");
                ui.weak("Если переместите exe — включите пункт заново.");
            } else {
                ui.weak("Доступно только в Windows.");
            }
            ui.separator();
            ui.strong("О программе");
            ui.label(format!("Merge Files {}", env!("CARGO_PKG_VERSION")));
            ui.weak(format!("Журнал ошибок: {}", config::log_path().display()));
        });
        self.show_settings = open;
    }
}

fn open_path(path: &Path, select: bool) {
    use std::process::Command;
    #[cfg(windows)]
    {
        use std::os::windows::process::CommandExt;
        let mut cmd = Command::new("explorer");
        if select {
            cmd.raw_arg(format!("/select,\"{}\"", path.display()));
        } else {
            cmd.arg(path);
        }
        let _ = cmd.spawn();
    }
    #[cfg(not(windows))]
    {
        let target = if select { path.parent().unwrap_or(path) } else { path };
        let opener = if cfg!(target_os = "macos") { "open" } else { "xdg-open" };
        let _ = Command::new(opener).arg(target).spawn();
    }
}

/// В Windows подключает системные шрифты как запасные: кириллица и значки (⬆ ⬇ ✖ ⚙) отображаются везде одинаково.
fn setup_fonts(ctx: &egui::Context) {
    let Some(windir) = std::env::var_os("WINDIR") else { return };
    let mut fonts = egui::FontDefinitions::default();
    for (name, file) in [("segoeui", "segoeui.ttf"), ("seguisym", "seguisym.ttf")] {
        if let Ok(data) = std::fs::read(Path::new(&windir).join("Fonts").join(file)) {
            fonts.font_data.insert(name.into(), Arc::new(egui::FontData::from_owned(data)));
            for family in [egui::FontFamily::Proportional, egui::FontFamily::Monospace] {
                if let Some(list) = fonts.families.get_mut(&family) {
                    list.push(name.into());
                }
            }
        }
    }
    ctx.set_fonts(fonts);
}
