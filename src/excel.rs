//! Объединение листов Excel в одну книгу.
//!
//! Читаются значения ячеек (формулы заменяются вычисленными значениями,
//! стили и объединённые ячейки не переносятся), ширина столбцов подбирается автоматически.

use anyhow::{Context, Result, anyhow};
use calamine::{Data, Reader, open_workbook_auto};
use rust_xlsxwriter::{Format, Workbook};
use std::collections::HashSet;
use std::path::{Path, PathBuf};

/// Список листов книги.
pub fn list_sheets(path: &Path) -> Result<Vec<String>> {
    let wb = open_workbook_auto(path).with_context(|| format!("не удалось открыть {}", path.display()))?;
    Ok(wb.sheet_names())
}

/// Файл и имена листов, которые нужно взять из него.
pub struct Source {
    pub path: PathBuf,
    pub sheets: Vec<String>,
}

pub fn merge(sources: &[Source], out: &Path) -> Result<()> {
    let mut book = Workbook::new();
    let mut used: HashSet<String> = HashSet::new();
    let date = Format::new().set_num_format("yyyy-mm-dd");
    let datetime = Format::new().set_num_format("yyyy-mm-dd hh:mm:ss");
    let mut count = 0;

    for src in sources {
        let mut wb = open_workbook_auto(&src.path)
            .with_context(|| format!("не удалось открыть {}", src.path.display()))?;
        let stem = src.path.file_stem().map(|s| s.to_string_lossy().into_owned()).unwrap_or_default();
        for sheet in &src.sheets {
            let range = wb
                .worksheet_range(sheet)
                .with_context(|| format!("лист «{sheet}» в {}", src.path.display()))?;
            let name = unique_name(&mut used, sheet, &stem);
            let ws = book.add_worksheet();
            ws.set_name(&name)?;
            let (r0, c0) = range.start().unwrap_or((0, 0));
            for (r, c, cell) in range.cells() {
                let (row, col) = (r as u32 + r0, c as u16 + c0 as u16);
                match cell {
                    Data::Empty => {}
                    Data::String(s) | Data::DateTimeIso(s) | Data::DurationIso(s) => {
                        ws.write_string(row, col, s)?;
                    }
                    Data::Int(i) => {
                        ws.write_number(row, col, *i as f64)?;
                    }
                    Data::Float(f) => {
                        ws.write_number(row, col, *f)?;
                    }
                    Data::Bool(b) => {
                        ws.write_boolean(row, col, *b)?;
                    }
                    Data::DateTime(d) => {
                        let v = d.as_f64();
                        let fmt = if v.fract() == 0.0 { &date } else { &datetime };
                        ws.write_number_with_format(row, col, v, fmt)?;
                    }
                    Data::Error(e) => {
                        ws.write_string(row, col, e.to_string())?;
                    }
                }
            }
            ws.autofit();
            count += 1;
        }
    }
    if count == 0 {
        return Err(anyhow!("не выбрано ни одного листа"));
    }
    book.save(out).with_context(|| format!("не удалось сохранить {}", out.display()))?;
    Ok(())
}

/// Имя листа Excel: до 31 символа, без `[]:*?/\`, уникальное без учёта регистра.
fn unique_name(used: &mut HashSet<String>, sheet: &str, file_stem: &str) -> String {
    let clean = |s: &str| -> String {
        s.chars().filter(|c| !"[]:*?/\\".contains(*c)).take(31).collect()
    };
    let try_insert = |used: &mut HashSet<String>, n: String| -> Option<String> {
        if !n.is_empty() && used.insert(n.to_lowercase()) { Some(n) } else { None }
    };
    if let Some(n) = try_insert(used, clean(sheet)) {
        return n;
    }
    if let Some(n) = try_insert(used, clean(&format!("{file_stem}_{sheet}"))) {
        return n;
    }
    for i in 2.. {
        let suffix = format!("_{i}");
        let base: String = clean(sheet).chars().take(31 - suffix.chars().count()).collect();
        if let Some(n) = try_insert(used, format!("{base}{suffix}")) {
            return n;
        }
    }
    unreachable!()
}
