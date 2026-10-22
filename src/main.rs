#![cfg_attr(all(not(debug_assertions), windows), windows_subsystem = "windows")]

mod excel;
mod pdf;
mod word;

use eframe::egui;
use std::path::{Path, PathBuf};
use std::sync::mpsc::{Receiver, channel};

#[derive(Clone, Copy, PartialEq)]
enum Kind {
    Excel,
    Word,
}

impl Kind {
    fn of(path: &Path) -> Option<Kind> {
        match path.extension()?.to_str()?.to_lowercase().as_str() {
            "xlsx" | "xlsm" | "xls" | "ods" => Some(Kind::Excel),
            "docx" | "doc" => Some(Kind::Word),
            _ => None,
        }
    }
    fn ext(self) -> &'static str {
        match self {
            Kind::Excel => "xlsx",
            Kind::Word => "docx",
        }
    }
}

#[derive(Clone, Copy, PartialEq)]
enum Target {
    Same,
    Pdf,
}

struct Entry {
    path: PathBuf,
    sheets: Vec<(String, bool)>,
}

struct App {
    entries: Vec<Entry>,
    kind: Option<Kind>,
    target: Target,
    status: String,
    is_error: bool,
    rx: Option<Receiver<Result<PathBuf, String>>>,
}

impl App {
    fn new() -> Self {
        App { entries: vec![], kind: None, target: Target::Same, status: String::new(), is_error: false, rx: None }
    }

    fn error(&mut self, msg: String) {
        self.status = msg;
        self.is_error = true;
    }

    fn add(&mut self, paths: Vec<PathBuf>) {
        for path in paths {
            let Some(k) = Kind::of(&path) else {
                self.error(format!("Неподдерживаемый файл: {}", path.display()));
                continue;
            };
            if self.kind.is_some_and(|cur| cur != k) {
                self.error("Нельзя смешивать Excel и Word в одном объединении".into());
                continue;
            }
            let sheets = if k == Kind::Excel {
                match excel::list_sheets(&path) {
                    Ok(s) => s.into_iter().map(|n| (n, true)).collect(),
                    Err(e) => {
                        self.error(format!("{e:#}"));
                        continue;
                    }
                }
            } else {
                vec![]
            };
            self.kind = Some(k);
            self.entries.push(Entry { path, sheets });
        }
        if self.entries.is_empty() {
            self.kind = None;
        }
    }

    fn merge(&mut self, ctx: &egui::Context) {
        let Some(kind) = self.kind else { return };
        let (ext, filter) = match self.target {
            Target::Same => (kind.ext(), kind.ext()),
            Target::Pdf => ("pdf", "pdf"),
        };
        let Some(out) = rfd::FileDialog::new()
            .set_file_name(format!("merged.{ext}"))
            .add_filter(filter, &[ext])
            .save_file()
        else {
            return;
        };
        let target = self.target;
        let entries: Vec<(PathBuf, Vec<String>)> = self
            .entries
            .iter()
            .map(|e| (e.path.clone(), e.sheets.iter().filter(|s| s.1).map(|s| s.0.clone()).collect()))
            .collect();
        let (tx, rx) = channel();
        self.rx = Some(rx);
        self.status = "Объединение…".into();
        self.is_error = false;
        let ctx = ctx.clone();
        std::thread::spawn(move || {
            let r = run(kind, target, entries, &out).map(|_| out).map_err(|e| format!("{e:#}"));
            let _ = tx.send(r);
            ctx.request_repaint();
        });
    }
}

fn run(kind: Kind, target: Target, entries: Vec<(PathBuf, Vec<String>)>, out: &Path) -> anyhow::Result<()> {
    let tmp = tempfile::tempdir()?;
    // Промежуточный файл того же типа; для PDF он конвертируется через LibreOffice.
    let merged = match target {
        Target::Same => out.to_path_buf(),
        Target::Pdf => tmp.path().join(format!("merged.{}", kind.ext())),
    };
    match kind {
        Kind::Excel => {
            let mut sources = vec![];
            for (path, sheets) in entries.into_iter().filter(|(_, s)| !s.is_empty()) {
                // Форматы кроме xlsx/xlsm (xls, ods) сначала конвертируются в xlsx.
                let native = path
                    .extension()
                    .is_some_and(|e| e.eq_ignore_ascii_case("xlsx") || e.eq_ignore_ascii_case("xlsm"));
                let path = if native {
                    path
                } else {
                    let dir = tmp.path().join(format!("in{}", sources.len()));
                    std::fs::create_dir_all(&dir)?;
                    pdf::convert(&path, "xlsx", &dir)?
                };
                sources.push(excel::Source { path, sheets });
            }
            excel::merge(&sources, &merged)?;
        }
        Kind::Word => {
            let mut docs = vec![];
            for (i, (path, _)) in entries.into_iter().enumerate() {
                if path.extension().is_some_and(|e| e.eq_ignore_ascii_case("doc")) {
                    let dir = tmp.path().join(format!("in{i}"));
                    std::fs::create_dir_all(&dir)?;
                    docs.push(pdf::convert(&path, "docx", &dir)?);
                } else {
                    docs.push(path);
                }
            }
            word::merge(&docs, &merged)?;
        }
    }
    if target == Target::Pdf {
        let dir = tmp.path().join("pdf");
        std::fs::create_dir_all(&dir)?;
        let result = pdf::convert(&merged, "pdf", &dir)?;
        std::fs::copy(result, out)?;
    }
    Ok(())
}

impl eframe::App for App {
    fn ui(&mut self, ui: &mut egui::Ui, _frame: &mut eframe::Frame) {
        let ctx = ui.ctx().clone();

        let dropped: Vec<PathBuf> = ctx.input(|i| i.raw.dropped_files.iter().filter_map(|f| Some(f.path().to_path_buf())).collect());
        if !dropped.is_empty() {
            self.add(dropped);
        }
        if let Some(rx) = &self.rx {
            if let Ok(r) = rx.try_recv() {
                match r {
                    Ok(p) => {
                        self.status = format!("Готово: {}", p.display());
                        self.is_error = false;
                    }
                    Err(e) => self.error(e),
                }
                self.rx = None;
            }
        }
        let busy = self.rx.is_some();

        egui::Panel::bottom("bottom").show(ui, |ui| {
            ui.add_space(6.0);
            let color = if self.is_error { egui::Color32::LIGHT_RED } else { egui::Color32::LIGHT_GREEN };
            ui.horizontal(|ui| {
                if busy {
                    ui.spinner();
                }
                ui.colored_label(color, &self.status);
            });
            ui.add_space(6.0);
        });

        egui::CentralPanel::default().show(ui, |ui| {
            ui.heading("Объединение файлов");
            ui.label("Листы Excel или документы Word — в один файл того же типа или в PDF.");
            ui.add_space(8.0);

            ui.horizontal(|ui| {
                if ui.add_enabled(!busy, egui::Button::new("➕ Добавить файлы…")).clicked() {
                    if let Some(p) = rfd::FileDialog::new()
                        .add_filter("Excel / Word", &["xlsx", "xlsm", "xls", "ods", "docx", "doc"])
                        .pick_files()
                    {
                        self.add(p);
                    }
                }
                if ui.add_enabled(!busy && !self.entries.is_empty(), egui::Button::new("Очистить")).clicked() {
                    self.entries.clear();
                    self.kind = None;
                }
                ui.label("или перетащите файлы в окно");
            });
            ui.separator();

            let mut action: Option<(usize, i8)> = None; // -1 вверх, 1 вниз, 0 удалить
            egui::ScrollArea::vertical().auto_shrink([false, false]).show(ui, |ui| {
                if self.entries.is_empty() {
                    ui.weak("Файлы не выбраны");
                }
                let n = self.entries.len();
                for (i, e) in self.entries.iter_mut().enumerate() {
                    ui.horizontal(|ui| {
                        ui.label(format!("{}.", i + 1));
                        let name = e.path.file_name().unwrap_or_default().to_string_lossy();
                        ui.strong(name.as_ref()).on_hover_text(e.path.display().to_string());
                        if ui.add_enabled(i > 0, egui::Button::new("⬆")).clicked() {
                            action = Some((i, -1));
                        }
                        if ui.add_enabled(i + 1 < n, egui::Button::new("⬇")).clicked() {
                            action = Some((i, 1));
                        }
                        if ui.button("✖").clicked() {
                            action = Some((i, 0));
                        }
                    });
                    if !e.sheets.is_empty() {
                        ui.indent(("sheets", i), |ui| {
                            for (name, on) in &mut e.sheets {
                                ui.checkbox(on, name.as_str());
                            }
                        });
                    }
                }
            });
            if !busy {
                if let Some((i, a)) = action {
                    match a {
                        -1 => self.entries.swap(i, i - 1),
                        1 => self.entries.swap(i, i + 1),
                        _ => {
                            self.entries.remove(i);
                            if self.entries.is_empty() {
                                self.kind = None;
                            }
                        }
                    }
                }
            }
        });

        egui::Panel::bottom("controls").show(ui, |ui| {
            ui.add_space(6.0);
            ui.horizontal(|ui| {
                ui.label("Результат:");
                let same = self.kind.map(|k| match k {
                    Kind::Excel => "Excel (.xlsx)",
                    Kind::Word => "Word (.docx)",
                });
                ui.radio_value(&mut self.target, Target::Same, same.unwrap_or("Тот же формат"));
                ui.radio_value(&mut self.target, Target::Pdf, "PDF");
            });
            ui.add_space(4.0);
            let ready = self.kind.is_some_and(|k| match k {
                Kind::Excel => self.entries.iter().any(|e| e.sheets.iter().any(|s| s.1)),
                Kind::Word => !self.entries.is_empty(),
            });
            if ui.add_enabled(ready && !busy, egui::Button::new("Объединить…").min_size([160.0, 32.0].into())).clicked() {
                self.merge(&ctx);
            }
            if self.target == Target::Pdf || self.entries.iter().any(|e| e.path.extension().is_some_and(|x| x == "doc")) {
                ui.weak("PDF, .doc, .xls и .ods требуют LibreOffice (в Windows подойдёт и Microsoft Office).");
            }
        });
    }
}

/// Консольный режим: `merge-files -o OUT in1 in2 ...` (тип результата по расширению OUT).
fn cli(args: Vec<String>) -> anyhow::Result<()> {
    let mut out = None;
    let mut inputs = vec![];
    let mut it = args.into_iter();
    while let Some(a) = it.next() {
        if a == "-o" {
            out = it.next().map(PathBuf::from);
        } else {
            inputs.push(PathBuf::from(a));
        }
    }
    let out = out.ok_or_else(|| anyhow::anyhow!("использование: merge-files -o OUT файл1 файл2 …"))?;
    let kind = inputs.first().and_then(|p| Kind::of(p)).ok_or_else(|| anyhow::anyhow!("нужны файлы xlsx/docx"))?;
    let target = if out.extension().is_some_and(|e| e.eq_ignore_ascii_case("pdf")) { Target::Pdf } else { Target::Same };
    let entries = inputs
        .into_iter()
        .map(|p| {
            let sheets = if kind == Kind::Excel { excel::list_sheets(&p)? } else { vec![] };
            Ok((p, sheets))
        })
        .collect::<anyhow::Result<Vec<_>>>()?;
    run(kind, target, entries, &out)
}

fn main() -> eframe::Result {
    let args: Vec<String> = std::env::args().skip(1).collect();
    if !args.is_empty() {
        if let Err(e) = cli(args) {
            eprintln!("Ошибка: {e:#}");
            std::process::exit(1);
        }
        return Ok(());
    }
    let options = eframe::NativeOptions {
        viewport: egui::ViewportBuilder::default().with_inner_size([560.0, 520.0]).with_drag_and_drop(true),
        ..Default::default()
    };
    eframe::run_native("Merge Files", options, Box::new(|_cc| Ok(Box::new(App::new()))))
}
