//! Объединение документов .docx: содержимое следующих документов дописывается
//! в конец первого (с разрывом страницы). Картинки и гиперссылки переносятся;
//! стили, нумерация, колонтитулы берутся из первого документа.

use anyhow::{Context, Result, anyhow};
use std::io::{Read, Write};
use std::path::{Path, PathBuf};

const DOC: &str = "word/document.xml";
const RELS: &str = "word/_rels/document.xml.rels";
const TYPES: &str = "[Content_Types].xml";
const PAGE_BREAK: &str = r#"<w:p><w:r><w:br w:type="page"/></w:r></w:p>"#;

struct Docx {
    entries: Vec<(String, Vec<u8>)>,
}

impl Docx {
    fn read(path: &Path) -> Result<Docx> {
        let f = std::fs::File::open(path).with_context(|| format!("не удалось открыть {}", path.display()))?;
        let mut zip = zip::ZipArchive::new(f).with_context(|| format!("{} — не docx", path.display()))?;
        let mut entries = Vec::new();
        for i in 0..zip.len() {
            let mut e = zip.by_index(i)?;
            if e.is_dir() {
                continue;
            }
            let mut buf = Vec::new();
            e.read_to_end(&mut buf)?;
            entries.push((e.name().map(|n| n.to_string()).unwrap_or_default(), buf));
        }
        Ok(Docx { entries })
    }

    fn get(&self, name: &str) -> Option<&[u8]> {
        self.entries.iter().find(|(n, _)| n == name).map(|(_, d)| d.as_slice())
    }

    fn text(&self, name: &str) -> Result<String> {
        let data = self.get(name).ok_or_else(|| anyhow!("в docx нет {name}"))?;
        Ok(String::from_utf8_lossy(data).into_owned())
    }

    fn set(&mut self, name: &str, data: Vec<u8>) {
        match self.entries.iter_mut().find(|(n, _)| n == name) {
            Some(e) => e.1 = data,
            None => self.entries.push((name.to_string(), data)),
        }
    }

    fn write(&self, path: &Path) -> Result<()> {
        let f = std::fs::File::create(path).with_context(|| format!("не удалось создать {}", path.display()))?;
        let mut zip = zip::ZipWriter::new(f);
        let opts = zip::write::SimpleFileOptions::default().compression_method(zip::CompressionMethod::Deflated);
        for (name, data) in &self.entries {
            zip.start_file(name, opts)?;
            zip.write_all(data)?;
        }
        zip.finish()?;
        Ok(())
    }
}

pub fn merge(inputs: &[PathBuf], out: &Path) -> Result<()> {
    let (first, rest) = inputs.split_first().ok_or_else(|| anyhow!("нет файлов"))?;
    let mut base = Docx::read(first)?;
    for (i, p) in rest.iter().enumerate() {
        let other = Docx::read(p)?;
        append(&mut base, &other, i + 1).with_context(|| format!("при добавлении {}", p.display()))?;
    }
    base.write(out)
}

/// Позиции содержимого `<w:body>`: (начало, конец до финального sectPr).
fn body_bounds(doc: &str) -> Result<(usize, usize)> {
    let start = doc.find("<w:body>").ok_or_else(|| anyhow!("нет <w:body>"))? + "<w:body>".len();
    let end = doc.rfind("</w:body>").ok_or_else(|| anyhow!("нет </w:body>"))?;
    let end = match doc[start..end].rfind("<w:sectPr") {
        // sectPr уровня body идёт последним и после него нет абзацев
        Some(i) if !doc[start + i..end].contains("</w:p>") => start + i,
        _ => end,
    };
    Ok((start, end))
}

fn attr(tag: &str, name: &str) -> Option<String> {
    let key = format!(" {name}=\"");
    let s = tag.find(&key)? + key.len();
    let e = tag[s..].find('"')?;
    Some(tag[s..s + e].to_string())
}

/// Удаляет все самозакрывающиеся теги `<name .../>`.
fn strip_tags(s: &mut String, name: &str) {
    let open = format!("<{name} ");
    while let Some(i) = s.find(&open) {
        match s[i..].find("/>") {
            Some(j) => s.replace_range(i..i + j + 2, ""),
            None => break,
        }
    }
}

fn append(base: &mut Docx, other: &Docx, n: usize) -> Result<()> {
    let mut base_doc = base.text(DOC)?;
    let other_doc = other.text(DOC)?;
    let mut base_rels = base.text(RELS)?;
    let other_rels = other.text(RELS).unwrap_or_default();
    let mut base_types = base.text(TYPES)?;
    let other_types = other.text(TYPES)?;

    let (s, e) = body_bounds(&other_doc)?;
    let mut content = other_doc[s..e].to_string();
    // Ссылки на колонтитулы другого документа в базе не существуют.
    strip_tags(&mut content, "w:headerReference");
    strip_tags(&mut content, "w:footerReference");

    // Картинки и гиперссылки: переносим связи с новыми идентификаторами.
    let mut new_rels = String::new();
    for (k, tag) in other_rels.split("<Relationship ").skip(1).enumerate() {
        let tag = tag.split("/>").next().unwrap_or("");
        let (Some(id), Some(ty), Some(target)) = (attr(tag, "Id"), attr(tag, "Type"), attr(tag, "Target")) else {
            continue;
        };
        let new_id = format!("rIdM{n}x{k}");
        let is_image = ty.ends_with("/image");
        let is_link = ty.ends_with("/hyperlink");
        if !is_image && !is_link {
            continue;
        }
        let mut target_out = target.clone();
        let mut mode = String::new();
        if let Some(m) = attr(tag, "TargetMode") {
            mode = format!(r#" TargetMode="{m}""#);
        } else if is_image {
            let src = format!("word/{}", target.trim_start_matches('/').trim_start_matches("word/"));
            let Some(data) = other.get(&src) else { continue };
            let base_name = target.rsplit('/').next().unwrap_or(&target);
            let new_name = format!("media/m{n}_{base_name}");
            base.set(&format!("word/{new_name}"), data.to_vec());
            target_out = new_name;
            if let Some(ext) = base_name.rsplit_once('.').map(|(_, e)| e.to_lowercase()) {
                ensure_default_type(&mut base_types, &other_types, &ext);
            }
        }
        for a in ["r:embed", "r:id", "r:link"] {
            content = content.replace(&format!("{a}=\"{id}\""), &format!("{a}=\"{new_id}\""));
        }
        new_rels.push_str(&format!(
            r#"<Relationship Id="{new_id}" Type="{ty}" Target="{target_out}"{mode}/>"#
        ));
    }
    let pos = base_rels.rfind("</Relationships>").ok_or_else(|| anyhow!("повреждён {RELS}"))?;
    base_rels.insert_str(pos, &new_rels);

    // Пространства имён корневого элемента другого документа, которых нет в базе.
    let other_root = root_tag(&other_doc)?;
    let mut decls = String::new();
    for part in other_root.split_whitespace() {
        if let Some(rest) = part.strip_prefix("xmlns:") {
            let prefix = rest.split('=').next().unwrap_or("");
            if !root_tag(&base_doc)?.contains(&format!("xmlns:{prefix}=")) {
                decls.push(' ');
                decls.push_str(part.trim_end_matches('>'));
            }
        }
    }
    if !decls.is_empty() {
        let r = root_tag(&base_doc)?.len();
        let at = base_doc.find("<w:document").unwrap() + r - 1;
        base_doc.insert_str(at, &decls);
    }

    let (_, be) = body_bounds(&base_doc)?;
    base_doc.insert_str(be, &format!("{PAGE_BREAK}{content}"));

    base.set(DOC, base_doc.into_bytes());
    base.set(RELS, base_rels.into_bytes());
    base.set(TYPES, base_types.into_bytes());
    Ok(())
}

/// Открывающий тег `<w:document ...>` целиком.
fn root_tag(doc: &str) -> Result<&str> {
    let s = doc.find("<w:document").ok_or_else(|| anyhow!("нет <w:document>"))?;
    let e = doc[s..].find('>').ok_or_else(|| anyhow!("повреждён document.xml"))?;
    Ok(&doc[s..=s + e])
}

fn ensure_default_type(base_types: &mut String, other_types: &str, ext: &str) {
    let key = format!("Extension=\"{ext}\"");
    if base_types.to_lowercase().contains(&key.to_lowercase()) {
        return;
    }
    let found = other_types
        .split("<Default ")
        .skip(1)
        .map(|t| t.split("/>").next().unwrap_or(""))
        .find(|t| attr(&format!(" {t}"), "Extension").is_some_and(|e| e.eq_ignore_ascii_case(ext)));
    if let (Some(t), Some(pos)) = (found, base_types.rfind("</Types>")) {
        base_types.insert_str(pos, &format!("<Default {t}/>"));
    }
}
