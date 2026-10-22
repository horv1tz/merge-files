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
        let tag = format!(" {}", tag.split("/>").next().unwrap_or(""));
        let tag = tag.as_str();
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

    merge_namespaces(&mut base_doc, &other_doc, "<w:document")?;

    // Стили и нумерация другого документа, которых нет в базе.
    let shift = n as u64 * 1000;
    merge_styles(base, other, shift)?;
    merge_numbering(base, other, shift, &mut base_rels, &mut base_types)?;
    content = shift_ids(&content, shift, false);

    let (_, be) = body_bounds(&base_doc)?;
    base_doc.insert_str(be, &format!("{PAGE_BREAK}{content}"));

    base.set(DOC, base_doc.into_bytes());
    base.set(RELS, base_rels.into_bytes());
    base.set(TYPES, base_types.into_bytes());
    Ok(())
}

/// Добавляет в корневой элемент `base` объявления xmlns, которые есть в `other`, но нет в `base`.
fn merge_namespaces(base: &mut String, other: &str, root: &str) -> Result<()> {
    let other_root = root_tag(other, root)?;
    let mut decls = String::new();
    for part in other_root.split_whitespace() {
        if let Some(rest) = part.strip_prefix("xmlns:") {
            let prefix = rest.split('=').next().unwrap_or("");
            if !root_tag(base, root)?.contains(&format!("xmlns:{prefix}=")) {
                decls.push(' ');
                decls.push_str(part.trim_end_matches('>').trim_end_matches('/'));
            }
        }
    }
    if !decls.is_empty() {
        let at = base.find(root).unwrap() + root_tag(base, root)?.len() - 1;
        base.insert_str(at, &decls);
    }
    Ok(())
}

/// Открывающий тег корневого элемента (например `<w:document ...>`) целиком.
fn root_tag<'a>(doc: &'a str, root: &str) -> Result<&'a str> {
    let s = doc.find(root).ok_or_else(|| anyhow!("нет {root}"))?;
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

/// Увеличивает числовые идентификаторы нумерации (`numId`, при `abstract_too` и `abstractNumId`).
/// Нулевой `numId` («без списка») не меняется.
fn shift_ids(s: &str, offset: u64, abstract_too: bool) -> String {
    let mut pats = vec!["w:numId=\"", "<w:numId w:val=\""];
    if abstract_too {
        pats.extend(["w:abstractNumId=\"", "<w:abstractNumId w:val=\""]);
    }
    let mut out = s.to_string();
    for pat in pats {
        let mut res = String::with_capacity(out.len());
        let mut rest = out.as_str();
        while let Some(i) = rest.find(pat) {
            let after = i + pat.len();
            res.push_str(&rest[..after]);
            rest = &rest[after..];
            let digits = rest.bytes().take_while(u8::is_ascii_digit).count();
            match rest[..digits].parse::<u64>() {
                Ok(v) if v != 0 || pat.contains("abstract") => res.push_str(&(v + offset).to_string()),
                _ => res.push_str(&rest[..digits]),
            }
            rest = &rest[digits..];
        }
        res.push_str(rest);
        out = res;
    }
    out
}

/// Вырезает блоки `<open ...>...</close>` (для абзацев нумерации и стилей).
fn blocks<'a>(xml: &'a str, open: &str, close: &str) -> Vec<&'a str> {
    let mut v = vec![];
    let mut rest = xml;
    while let Some(i) = rest.find(open) {
        let Some(j) = rest[i..].find(close) else { break };
        let end = i + j + close.len();
        v.push(&rest[i..end]);
        rest = &rest[end..];
    }
    v
}

fn merge_styles(base: &mut Docx, other: &Docx, shift: u64) -> Result<()> {
    let (Some(_), Some(_)) = (base.get("word/styles.xml"), other.get("word/styles.xml")) else {
        return Ok(());
    };
    let mut b = base.text("word/styles.xml")?;
    let o = other.text("word/styles.xml")?;
    let mut add = String::new();
    for st in blocks(&o, "<w:style ", "</w:style>") {
        let Some(id) = attr(st, "w:styleId") else { continue };
        if !b.contains(&format!("w:styleId=\"{id}\"")) {
            add.push_str(&shift_ids(st, shift, false));
        }
    }
    if !add.is_empty() {
        merge_namespaces(&mut b, &o, "<w:styles")?;
        let pos = b.rfind("</w:styles>").ok_or_else(|| anyhow!("повреждён styles.xml"))?;
        b.insert_str(pos, &add);
        base.set("word/styles.xml", b.into_bytes());
    }
    Ok(())
}

fn merge_numbering(base: &mut Docx, other: &Docx, shift: u64, rels: &mut String, types: &mut String) -> Result<()> {
    let Some(o) = other.get("word/numbering.xml") else { return Ok(()) };
    let o = String::from_utf8_lossy(o).into_owned();
    let abstracts: String = blocks(&o, "<w:abstractNum ", "</w:abstractNum>").concat();
    let nums: String = blocks(&o, "<w:num ", "</w:num>").concat();
    if abstracts.is_empty() && nums.is_empty() {
        return Ok(());
    }
    let mut b = match base.get("word/numbering.xml") {
        Some(d) => String::from_utf8_lossy(d).into_owned(),
        None => {
            // В базе нет нумерации: создаём часть и связи с ней.
            let at = rels.rfind("</Relationships>").ok_or_else(|| anyhow!("повреждён {RELS}"))?;
            rels.insert_str(
                at,
                r#"<Relationship Id="rIdMnum" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/>"#,
            );
            let at = types.rfind("</Types>").ok_or_else(|| anyhow!("повреждён {TYPES}"))?;
            types.insert_str(
                at,
                r#"<Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>"#,
            );
            let root = root_tag(&o, "<w:numbering")?;
            format!("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n{}</w:numbering>", root.trim_end_matches("/>"))
        }
    };
    merge_namespaces(&mut b, &o, "<w:numbering")?;
    let abstracts = shift_ids(&abstracts, shift, true);
    let nums = shift_ids(&nums, shift, true);
    // abstractNum должны идти перед num.
    let pos = b.find("<w:num ").unwrap_or_else(|| b.rfind("</w:numbering>").unwrap_or(b.len()));
    b.insert_str(pos, &abstracts);
    let pos = b.rfind("</w:numbering>").ok_or_else(|| anyhow!("повреждён numbering.xml"))?;
    b.insert_str(pos, &nums);
    base.set("word/numbering.xml", b.into_bytes());
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;

    fn make_docx(path: &Path, text: &str, with_image: bool) {
        let f = std::fs::File::create(path).unwrap();
        let mut z = zip::ZipWriter::new(f);
        let o = zip::write::SimpleFileOptions::default();
        let ns = r#"xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships""#;
        let files = [
            ("[Content_Types].xml", r#"<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/></Types>"#.to_string()),
            ("word/_rels/document.xml.rels", if with_image {
                r#"<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/i.png"/></Relationships>"#.to_string()
            } else {
                r#"<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>"#.to_string()
            }),
            ("word/document.xml", format!(
                r#"<?xml version="1.0"?><w:document {ns}><w:body><w:p><w:r><w:t>{text}</w:t></w:r></w:p>{}<w:sectPr><w:pgSz w:w="1"/></w:sectPr></w:body></w:document>"#,
                if with_image { r#"<w:p><w:r><w:drawing><a:blip r:embed="rId1"/></w:drawing></w:r></w:p>"# } else { "" }
            )),
        ];
        for (n, c) in files {
            z.start_file(n, o).unwrap();
            z.write_all(c.as_bytes()).unwrap();
        }
        if with_image {
            z.start_file("word/media/i.png", o).unwrap();
            z.write_all(b"PNG").unwrap();
        }
        z.finish().unwrap();
    }

    #[test]
    fn merges_text_images_and_page_breaks() {
        let dir = tempfile::tempdir().unwrap();
        let (a, b, out) = (dir.path().join("a.docx"), dir.path().join("b.docx"), dir.path().join("o.docx"));
        make_docx(&a, "ONE", false);
        make_docx(&b, "TWO", true);
        merge(&[a, b], &out).unwrap();
        let d = Docx::read(&out).unwrap();
        let doc = d.text(DOC).unwrap();
        assert!(doc.find("ONE").unwrap() < doc.find("TWO").unwrap());
        assert!(doc.contains(r#"w:type="page""#));
        assert_eq!(doc.matches("<w:sectPr>").count(), 1);
        assert!(doc.contains(r#"r:embed="rIdM1x0""#));
        assert!(d.get("word/media/m1_i.png").is_some());
        assert!(d.text(RELS).unwrap().contains(r#"Id="rIdM1x0""#));
    }

    #[test]
    fn shifts_num_ids_but_not_zero() {
        let s = r#"<w:numId w:val="0"/><w:numId w:val="3"/>"#;
        assert_eq!(shift_ids(s, 1000, false), r#"<w:numId w:val="0"/><w:numId w:val="1003"/>"#);
    }
}
