//! Объединение листов Excel в одну книгу.
//!
//! Листы копируются вместе с форматированием (стили, объединённые ячейки,
//! ширина столбцов, формулы). Нужен .xlsx; прочие форматы конвертируются заранее.

use anyhow::{Context, Result, anyhow};
use calamine::{Reader, open_workbook_auto};
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
    let mut base: Option<umya_spreadsheet::Workbook> = None;
    let mut used: HashSet<String> = HashSet::new();

    for src in sources {
        let mut wb = umya_spreadsheet::reader::xlsx::read(&src.path)
            .map_err(|e| anyhow!("не удалось прочитать {}: {e:?}", src.path.display()))?;
        let stem = src.path.file_stem().map(|s| s.to_string_lossy().into_owned()).unwrap_or_default();
        let all: Vec<String> = wb.sheet_collection().iter().map(|s| s.name().to_string()).collect();
        for name in &src.sheets {
            if !all.contains(name) {
                return Err(anyhow!("в {} нет листа «{name}»", src.path.display()));
            }
        }
        match base.as_mut() {
            None => {
                // Первая книга — основа: сохраняются её тема и стили.
                for name in all.iter().filter(|n| !src.sheets.contains(n)) {
                    wb.remove_sheet_by_name(name).map_err(|e| anyhow!("{e:?}"))?;
                }
                for name in &src.sheets {
                    let new = unique_name(&mut used, name, &stem);
                    let ws = wb.sheet_by_name_mut(name).map_err(|e| anyhow!("{e:?}"))?;
                    ws.set_name(new);
                }
                base = Some(wb);
            }
            Some(book) => {
                for name in &src.sheets {
                    let mut ws = wb.sheet_by_name(name).map_err(|e| anyhow!("{e:?}"))?.clone();
                    ws.set_name(unique_name(&mut used, name, &stem));
                    book.add_sheet(ws).map_err(|e| anyhow!("{e:?}"))?;
                }
            }
        }
    }
    let mut book = base.ok_or_else(|| anyhow!("не выбрано ни одного листа"))?;
    if book.sheet_collection().is_empty() {
        return Err(anyhow!("не выбрано ни одного листа"));
    }
    // Выделенными должны быть не все листы сразу (иначе Excel откроет групповой режим).
    for (i, ws) in book.sheet_collection_mut().iter_mut().enumerate() {
        for view in ws.sheet_views_mut().sheet_view_list_mut() {
            view.set_tab_selected(i == 0);
        }
    }
    book.set_active_sheet(0);
    umya_spreadsheet::writer::xlsx::write(&book, out)
        .map_err(|e| anyhow!("не удалось сохранить {}: {e:?}", out.display()))?;
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

#[cfg(test)]
mod tests {
    use super::*;

    fn make(path: &Path, cell: &str) {
        let mut b = umya_spreadsheet::new_file();
        {
            let ws = b.sheet_by_name_mut("Sheet1").unwrap();
            ws.cell_mut("A1").set_value(cell);
            ws.cell_mut("A1").style_mut().font_mut().set_bold(true);
        }
        umya_spreadsheet::writer::xlsx::write(&b, path).unwrap();
    }

    #[test]
    fn merges_sheets_with_unique_names_and_styles() {
        let dir = tempfile::tempdir().unwrap();
        let (a, b, out) = (dir.path().join("a.xlsx"), dir.path().join("b.xlsx"), dir.path().join("o.xlsx"));
        make(&a, "first");
        make(&b, "second");
        let src = |p: &PathBuf| Source { path: p.clone(), sheets: vec!["Sheet1".into()] };
        merge(&[src(&a), src(&b)], &out).unwrap();
        let book = umya_spreadsheet::reader::xlsx::read(&out).unwrap();
        let names: Vec<_> = book.sheet_collection().iter().map(|s| s.name().to_string()).collect();
        assert_eq!(names, ["Sheet1", "b_Sheet1"]);
        let ws = book.sheet_by_name("b_Sheet1").unwrap();
        assert_eq!(ws.cell("A1").unwrap().value(), "second");
        assert_eq!(ws.cell("A1").unwrap().style().font().unwrap().bold(), true);
    }

    #[test]
    fn sheet_names_are_valid() {
        let mut used = HashSet::new();
        assert_eq!(unique_name(&mut used, "a/b", "f"), "ab");
        assert_eq!(unique_name(&mut used, "AB", "f"), "f_AB");
        assert_eq!(unique_name(&mut used, "AB", "f"), "AB_2");
    }
}
