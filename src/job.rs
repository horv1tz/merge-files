//! Конвейер объединения: проверки, конвертация исходников, слияние, вывод в PDF.

use crate::{excel, pdf, word};
use anyhow::{Context, Result, anyhow, bail};
use std::io::Read;
use std::path::{Path, PathBuf};
use std::sync::atomic::{AtomicBool, Ordering};

#[derive(Clone, Copy, PartialEq, Debug)]
pub enum Kind {
    Excel,
    Word,
}

impl Kind {
    pub fn of(path: &Path) -> Option<Kind> {
        match path.extension()?.to_str()?.to_lowercase().as_str() {
            "xlsx" | "xlsm" | "xls" | "ods" => Some(Kind::Excel),
            "docx" | "doc" => Some(Kind::Word),
            _ => None,
        }
    }

    pub fn ext(self) -> &'static str {
        match self {
            Kind::Excel => "xlsx",
            Kind::Word => "docx",
        }
    }
}

#[derive(Clone, Copy, PartialEq, Debug)]
pub enum Target {
    Same,
    Pdf,
}

impl Target {
    pub fn ext(self, kind: Kind) -> &'static str {
        match self {
            Target::Same => kind.ext(),
            Target::Pdf => "pdf",
        }
    }
}

/// Один входной файл задания.
#[derive(Clone)]
pub struct JobEntry {
    pub path: PathBuf,
    /// Все листы исходного файла (в порядке файла) — для Excel; пусто для Word.
    pub all_sheets: Vec<String>,
    /// Индексы выбранных листов в `all_sheets`.
    pub selected: Vec<usize>,
}

#[derive(Clone)]
pub struct Job {
    pub kind: Kind,
    pub target: Target,
    pub entries: Vec<JobEntry>,
    pub out: PathBuf,
}

pub struct Outcome {
    pub warnings: Vec<String>,
}

/// Реальный формат файла по содержимому (расширению верить нельзя).
#[derive(PartialEq, Debug)]
pub enum Sniffed {
    Xlsx,
    Docx,
    /// Старый бинарный формат OLE (.xls/.doc) — в том числе зашифрованные xlsx/docx.
    Ole,
    /// Другой zip-контейнер (например .ods).
    OtherZip,
    Unknown,
}

pub fn sniff(path: &Path) -> Result<Sniffed> {
    let mut f = std::fs::File::open(path).with_context(|| format!("не удалось открыть «{}»", path.display()))?;
    let mut head = [0u8; 8];
    let n = f.read(&mut head)?;
    if n >= 4 && head[..4] == *b"PK\x03\x04" {
        let Ok(zip) = zip::ZipArchive::new(std::fs::File::open(path)?) else { return Ok(Sniffed::Unknown) };
        let names: Vec<String> = zip.file_names().filter_map(Result::ok).map(|c| c.into_owned()).collect();
        let has = |n: &str| names.iter().any(|x| x == n);
        return Ok(if has("xl/workbook.xml") {
            Sniffed::Xlsx
        } else if has("word/document.xml") {
            Sniffed::Docx
        } else {
            Sniffed::OtherZip
        });
    }
    if n >= 8 && head == [0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1] {
        return Ok(Sniffed::Ole);
    }
    Ok(Sniffed::Unknown)
}

/// Нужен ли LibreOffice/Office для этого задания (до выбора файла результата).
pub fn needs_converter(job_target: Target, paths: &[PathBuf]) -> bool {
    job_target == Target::Pdf
        || paths.iter().any(|p| {
            !p.extension()
                .and_then(|e| e.to_str())
                .is_some_and(|e| matches!(e.to_lowercase().as_str(), "xlsx" | "xlsm" | "docx"))
        })
}

pub fn same_file(a: &Path, b: &Path) -> bool {
    match (std::fs::canonicalize(a), std::fs::canonicalize(b)) {
        (Ok(x), Ok(y)) => x == y,
        _ => a == b,
    }
}

/// Проверки, которые можно сделать до начала работы.
pub fn validate(job: &Job) -> Result<()> {
    if job.entries.is_empty() {
        bail!("Не выбрано ни одного файла");
    }
    for e in &job.entries {
        if same_file(&e.path, &job.out) {
            bail!("Файл результата совпадает с исходным файлом «{}». Выберите другое имя.", e.path.display());
        }
    }
    let want = job.target.ext(job.kind);
    let ok = job.out.extension().and_then(|e| e.to_str()).is_some_and(|e| e.eq_ignore_ascii_case(want));
    if !ok {
        bail!("Расширение файла результата должно быть .{want}");
    }
    if let Some(dir) = job.out.parent().filter(|d| !d.as_os_str().is_empty()) {
        if !dir.is_dir() {
            bail!("Папка «{}» не существует", dir.display());
        }
    }
    Ok(())
}

fn check_cancel(cancel: &AtomicBool) -> Result<()> {
    if cancel.load(Ordering::Relaxed) { Err(anyhow!("Операция отменена")) } else { Ok(()) }
}

pub fn run(job: &Job, cancel: &AtomicBool, progress: &dyn Fn(&str)) -> Result<Outcome> {
    validate(job)?;
    let tmp = tempfile::Builder::new().prefix("merge-files-").tempdir().context("не удалось создать временную папку")?;
    let conv = pdf::Converter::new(tmp.path(), cancel);
    let merged = tmp.path().join(format!("merged.{}", job.kind.ext()));
    let mut warnings = vec![];

    match job.kind {
        Kind::Excel => {
            let mut sources = vec![];
            for (i, e) in job.entries.iter().enumerate() {
                check_cancel(cancel)?;
                let name = e.path.file_name().unwrap_or_default().to_string_lossy().into_owned();
                if e.selected.is_empty() {
                    continue;
                }
                let (path, names) = match sniff(&e.path)? {
                    Sniffed::Xlsx => (e.path.clone(), e.selected.iter().map(|&k| e.all_sheets[k].clone()).collect()),
                    Sniffed::Docx | Sniffed::Unknown => bail!("«{name}» не является книгой Excel"),
                    Sniffed::Ole | Sniffed::OtherZip => {
                        progress(&format!("Конвертация «{name}»…"));
                        let converted = conv.convert(&e.path, "xlsx", &tmp.path().join(format!("in{i}")))?;
                        // LibreOffice может изменить имена листов (обрезает до 31 символа), поэтому выбираем по номеру.
                        let after = excel::list_sheets(&converted)?;
                        let names = e
                            .selected
                            .iter()
                            .map(|&k| after.get(k).map(|s| s.name.clone()).ok_or_else(|| anyhow!("в «{name}» не найден лист №{}", k + 1)))
                            .collect::<Result<Vec<_>>>()?;
                        (converted, names)
                    }
                };
                sources.push(excel::Source { path, sheets: names });
            }
            check_cancel(cancel)?;
            progress("Объединение листов…");
            let opts = excel::Options { fit_to_width: job.target == Target::Pdf };
            warnings.extend(excel::merge(&sources, &merged, &opts)?);
        }
        Kind::Word => {
            let mut docs = vec![];
            for (i, e) in job.entries.iter().enumerate() {
                check_cancel(cancel)?;
                let name = e.path.file_name().unwrap_or_default().to_string_lossy().into_owned();
                match sniff(&e.path)? {
                    Sniffed::Docx => docs.push(e.path.clone()),
                    Sniffed::Xlsx | Sniffed::Unknown => bail!("«{name}» не является документом Word"),
                    Sniffed::Ole | Sniffed::OtherZip => {
                        progress(&format!("Конвертация «{name}»…"));
                        docs.push(conv.convert(&e.path, "docx", &tmp.path().join(format!("in{i}")))?);
                    }
                }
            }
            check_cancel(cancel)?;
            progress("Объединение документов…");
            warnings.extend(word::merge(&docs, &merged)?);
        }
    }

    let result = match job.target {
        Target::Same => merged,
        Target::Pdf => {
            check_cancel(cancel)?;
            progress("Создание PDF…");
            conv.convert(&merged, "pdf", &tmp.path().join("pdf"))?
        }
    };
    check_cancel(cancel)?;
    // Результат кладётся в итоговый файл только целиком — исходный файл не портится при сбое.
    std::fs::copy(&result, &job.out).map_err(|e| {
        anyhow!("Не удалось сохранить «{}»: {e}. Возможно, файл открыт в другой программе или нет прав на запись.", job.out.display())
    })?;
    Ok(Outcome { warnings })
}

#[cfg(test)]
mod tests {
    use super::*;

    fn entry(p: &Path) -> JobEntry {
        JobEntry { path: p.to_path_buf(), all_sheets: vec![], selected: vec![] }
    }

    #[test]
    fn output_must_not_be_input_and_must_match_type() {
        let dir = tempfile::tempdir().unwrap();
        let a = dir.path().join("a.docx");
        std::fs::write(&a, b"x").unwrap();
        let mut job = Job { kind: Kind::Word, target: Target::Same, entries: vec![entry(&a)], out: a.clone() };
        assert!(validate(&job).unwrap_err().to_string().contains("совпадает"));
        job.out = dir.path().join("o.pdf");
        assert!(validate(&job).unwrap_err().to_string().contains(".docx"));
        job.target = Target::Pdf;
        assert!(validate(&job).is_ok());
        job.out = dir.path().join("нет").join("o.pdf");
        assert!(validate(&job).unwrap_err().to_string().contains("не существует"));
    }

    #[test]
    fn sniffs_real_format() {
        let dir = tempfile::tempdir().unwrap();
        let p = dir.path().join("fake.xlsx");
        std::fs::write(&p, b"just text").unwrap();
        assert_eq!(sniff(&p).unwrap(), Sniffed::Unknown);
        let ole = dir.path().join("a.xls");
        std::fs::write(&ole, [0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1, 0]).unwrap();
        assert_eq!(sniff(&ole).unwrap(), Sniffed::Ole);
    }

    #[test]
    fn converter_needed_only_for_legacy_or_pdf() {
        let p = |s: &str| PathBuf::from(s);
        assert!(!needs_converter(Target::Same, &[p("a.XLSX"), p("b.docx")]));
        assert!(needs_converter(Target::Same, &[p("a.xls")]));
        assert!(needs_converter(Target::Pdf, &[p("a.docx")]));
    }
}
