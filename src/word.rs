//! Объединение документов .docx: документы идут друг за другом, каждый в своём разделе
//! (свой размер и ориентация страницы, поля, колонки, колонтитулы).
//!
//! Пакеты разбираются структурно (собственный небольшой XML-парсер без потерь), а не заменой
//! подстрок. Переносятся: картинки, диаграммы, SmartArt, OLE-объекты и прочие связанные части
//! (рекурсивно, вместе со своими связями и типами содержимого), гиперссылки, колонтитулы,
//! сноски, концевые сноски, примечания (с цепочками ответов), стили и нумерация списков.
//! Все идентификаторы (связи, имена частей, `docPr`, закладки, правки, `sdt`, номера списков,
//! стилей и абзацев) выбираются так, чтобы не пересекаться с уже существующими, поэтому результат
//! слияния можно снова использовать как входной файл.
//!
//! Что не переносится без потерь, не пропадает молча: `merge` возвращает предупреждения для
//! пользователя. Это: стили с тем же именем, но другим оформлением (используется стиль первого
//! документа), «шрифт по умолчанию», тема, настройки всего документа (зеркальные поля, разные
//! колонтитулы чётных/нечётных страниц), встроенные шрифты, привязка элементов управления к
//! данным, «битые» ссылки исходных файлов. Формат Strict Open XML не поддерживается (ошибка).

use anyhow::{Context, Result, anyhow, bail};
use std::collections::{HashMap, HashSet};
use std::io::{Read, Write};
use std::path::{Path, PathBuf};
use xml::{Elem, Node};

// ───────────────────────────── XML без потерь ─────────────────────────────

mod xml {
    use anyhow::{Result, anyhow, bail};

    #[derive(Clone, Debug)]
    pub enum Node {
        Elem(Elem),
        /// Текст как в файле (с экранированием).
        Text(String),
        /// Комментарий, CDATA или инструкция обработки — как есть.
        Raw(String),
    }

    #[derive(Clone, Debug, Default)]
    pub struct Elem {
        pub name: String,
        /// Значения атрибутов хранятся без экранирования.
        pub attrs: Vec<(String, String)>,
        pub children: Vec<Node>,
    }

    #[derive(Clone, Debug)]
    pub struct Doc {
        pub prolog: Vec<String>,
        pub root: Elem,
    }

    impl Elem {
        pub fn new(name: &str) -> Elem {
            Elem { name: name.to_string(), ..Default::default() }
        }
        pub fn with(mut self, k: &str, v: &str) -> Elem {
            self.set_attr(k, v);
            self
        }
        pub fn with_child(mut self, c: Elem) -> Elem {
            self.push(c);
            self
        }
        pub fn attr(&self, k: &str) -> Option<&str> {
            self.attrs.iter().find(|(n, _)| n == k).map(|(_, v)| v.as_str())
        }
        pub fn set_attr(&mut self, k: &str, v: impl Into<String>) {
            let v = v.into();
            match self.attrs.iter_mut().find(|(n, _)| n == k) {
                Some(a) => a.1 = v,
                None => self.attrs.push((k.to_string(), v)),
            }
        }
        pub fn remove_attr(&mut self, k: &str) -> Option<String> {
            let i = self.attrs.iter().position(|(n, _)| n == k)?;
            Some(self.attrs.remove(i).1)
        }
        pub fn local(&self) -> &str {
            self.name.rsplit(':').next().unwrap_or("")
        }
        pub fn prefix(&self) -> &str {
            self.name.split_once(':').map(|x| x.0).unwrap_or("")
        }
        pub fn elems(&self) -> impl Iterator<Item = &Elem> + '_ {
            self.children.iter().filter_map(|n| if let Node::Elem(e) = n { Some(e) } else { None })
        }
        pub fn child(&self, name: &str) -> Option<&Elem> {
            self.elems().find(|e| e.name == name)
        }
        pub fn child_mut(&mut self, name: &str) -> Option<&mut Elem> {
            self.children.iter_mut().find_map(|n| match n {
                Node::Elem(e) if e.name == name => Some(e),
                _ => None,
            })
        }
        pub fn push(&mut self, e: Elem) {
            self.children.push(Node::Elem(e));
        }
        pub fn val(&self) -> Option<&str> {
            self.attr("w:val")
        }
        /// Текст внутри элемента (без экранирования).
        pub fn text(&self) -> String {
            let mut s = String::new();
            for n in &self.children {
                if let Node::Text(t) = n {
                    s.push_str(&unescape(t));
                }
            }
            s
        }
        pub fn set_text(&mut self, s: &str) {
            self.children = vec![Node::Text(escape_text(s))];
        }
    }

    pub fn unescape(s: &str) -> String {
        if !s.contains('&') {
            return s.to_string();
        }
        let mut out = String::with_capacity(s.len());
        let mut rest = s;
        while let Some(i) = rest.find('&') {
            out.push_str(&rest[..i]);
            rest = &rest[i..];
            let Some(j) = rest.find(';') else { break };
            let ent = &rest[1..j];
            let rep = match ent {
                "lt" => Some('<'),
                "gt" => Some('>'),
                "amp" => Some('&'),
                "quot" => Some('"'),
                "apos" => Some('\''),
                _ => ent
                    .strip_prefix("#x")
                    .or_else(|| ent.strip_prefix("#X"))
                    .and_then(|h| u32::from_str_radix(h, 16).ok())
                    .or_else(|| ent.strip_prefix('#').and_then(|d| d.parse::<u32>().ok()))
                    .and_then(char::from_u32),
            };
            match rep {
                Some(c) => {
                    out.push(c);
                    rest = &rest[j + 1..];
                }
                None => {
                    out.push('&');
                    rest = &rest[1..];
                }
            }
        }
        out.push_str(rest);
        out
    }

    pub fn escape_text(s: &str) -> String {
        let mut o = String::with_capacity(s.len());
        for c in s.chars() {
            match c {
                '&' => o.push_str("&amp;"),
                '<' => o.push_str("&lt;"),
                '>' => o.push_str("&gt;"),
                c => o.push(c),
            }
        }
        o
    }

    fn escape_attr(s: &str, o: &mut String) {
        for c in s.chars() {
            match c {
                '&' => o.push_str("&amp;"),
                '<' => o.push_str("&lt;"),
                '>' => o.push_str("&gt;"),
                '"' => o.push_str("&quot;"),
                '\t' => o.push_str("&#9;"),
                '\n' => o.push_str("&#10;"),
                '\r' => o.push_str("&#13;"),
                c => o.push(c),
            }
        }
    }

    pub fn parse(data: &[u8]) -> Result<Doc> {
        let data = data.strip_prefix(&[0xEF, 0xBB, 0xBF][..]).unwrap_or(data);
        if data.starts_with(&[0xFF, 0xFE]) || data.starts_with(&[0xFE, 0xFF]) {
            bail!("XML в кодировке UTF-16 не поддерживается");
        }
        let s = std::str::from_utf8(data).map_err(|_| anyhow!("XML не в кодировке UTF-8"))?;
        let b = s.as_bytes();
        let mut i = 0;
        let mut stack: Vec<Elem> = vec![];
        let mut root: Option<Elem> = None;
        let mut prolog = vec![];
        let is_ws = |c: u8| c == b' ' || c == b'\t' || c == b'\n' || c == b'\r';
        while i < b.len() {
            if b[i] != b'<' {
                let j = s[i..].find('<').map(|x| x + i).unwrap_or(b.len());
                let t = &s[i..j];
                match stack.last_mut() {
                    Some(p) => p.children.push(Node::Text(t.to_string())),
                    None if !t.trim().is_empty() => bail!("XML повреждён: текст вне корневого элемента"),
                    None => {}
                }
                i = j;
                continue;
            }
            let rest = &s[i..];
            if rest.starts_with("<!--") {
                let e = rest.find("-->").ok_or_else(|| anyhow!("XML повреждён: не закрыт комментарий"))? + 3;
                if let Some(p) = stack.last_mut() {
                    p.children.push(Node::Raw(rest[..e].to_string()));
                } else if root.is_none() {
                    prolog.push(rest[..e].to_string());
                }
                i += e;
            } else if rest.starts_with("<![CDATA[") {
                let e = rest.find("]]>").ok_or_else(|| anyhow!("XML повреждён: не закрыт CDATA"))? + 3;
                match stack.last_mut() {
                    Some(p) => p.children.push(Node::Raw(rest[..e].to_string())),
                    None => bail!("XML повреждён: CDATA вне корневого элемента"),
                }
                i += e;
            } else if rest.starts_with("<?") {
                let e = rest.find("?>").ok_or_else(|| anyhow!("XML повреждён: не закрыта инструкция"))? + 2;
                let raw = &rest[..e];
                if let Some(p) = stack.last_mut() {
                    p.children.push(Node::Raw(raw.to_string()));
                } else if root.is_none() && !raw.starts_with("<?xml ") && !raw.starts_with("<?xml?") {
                    prolog.push(raw.to_string());
                }
                i += e;
            } else if rest.starts_with("<!") {
                bail!("XML с DOCTYPE не поддерживается");
            } else if rest.starts_with("</") {
                let e = rest.find('>').ok_or_else(|| anyhow!("XML повреждён: не закрыт тег"))?;
                let name = rest[2..e].trim();
                let el = stack.pop().ok_or_else(|| anyhow!("XML повреждён: лишний закрывающий тег </{name}>"))?;
                if el.name != name {
                    bail!("XML повреждён: <{}> закрыт как </{name}>", el.name);
                }
                match stack.last_mut() {
                    Some(p) => p.children.push(Node::Elem(el)),
                    None => {
                        if root.is_some() {
                            bail!("XML повреждён: несколько корневых элементов");
                        }
                        root = Some(el);
                    }
                }
                i += e + 1;
            } else {
                // открывающий тег
                let mut j = 1;
                let tb = rest.as_bytes();
                while j < tb.len() && !is_ws(tb[j]) && tb[j] != b'>' && tb[j] != b'/' {
                    j += 1;
                }
                let mut el = Elem::new(&rest[1..j]);
                if el.name.is_empty() {
                    bail!("XML повреждён: пустое имя тега");
                }
                let self_closing;
                loop {
                    while j < tb.len() && is_ws(tb[j]) {
                        j += 1;
                    }
                    if j >= tb.len() {
                        bail!("XML повреждён: не закрыт тег <{}", el.name);
                    }
                    if tb[j] == b'>' {
                        self_closing = false;
                        j += 1;
                        break;
                    }
                    if tb[j] == b'/' {
                        if tb.get(j + 1) != Some(&b'>') {
                            bail!("XML повреждён: ошибка в теге <{}>", el.name);
                        }
                        self_closing = true;
                        j += 2;
                        break;
                    }
                    let ns = j;
                    while j < tb.len() && tb[j] != b'=' && !is_ws(tb[j]) && tb[j] != b'>' {
                        j += 1;
                    }
                    let an = rest[ns..j].to_string();
                    while j < tb.len() && is_ws(tb[j]) {
                        j += 1;
                    }
                    if tb.get(j) != Some(&b'=') {
                        bail!("XML повреждён: атрибут {an} без значения в <{}>", el.name);
                    }
                    j += 1;
                    while j < tb.len() && is_ws(tb[j]) {
                        j += 1;
                    }
                    let q = *tb.get(j).ok_or_else(|| anyhow!("XML повреждён: обрыв в теге"))?;
                    if q != b'"' && q != b'\'' {
                        bail!("XML повреждён: значение атрибута {an} без кавычек");
                    }
                    j += 1;
                    let vs = j;
                    while j < tb.len() && tb[j] != q {
                        j += 1;
                    }
                    if j >= tb.len() {
                        bail!("XML повреждён: не закрыто значение атрибута {an}");
                    }
                    el.attrs.push((an, unescape(&rest[vs..j])));
                    j += 1;
                }
                i += j;
                if self_closing {
                    match stack.last_mut() {
                        Some(p) => p.children.push(Node::Elem(el)),
                        None => {
                            if root.is_some() {
                                bail!("XML повреждён: несколько корневых элементов");
                            }
                            root = Some(el);
                        }
                    }
                } else {
                    if stack.len() >= 400 {
                        bail!("XML слишком глубоко вложен");
                    }
                    stack.push(el);
                }
            }
        }
        if let Some(e) = stack.last() {
            bail!("XML повреждён: не закрыт <{}>", e.name);
        }
        Ok(Doc { prolog, root: root.ok_or_else(|| anyhow!("XML пуст"))? })
    }

    fn write_elem(e: &Elem, o: &mut String) {
        o.push('<');
        o.push_str(&e.name);
        for (k, v) in &e.attrs {
            o.push(' ');
            o.push_str(k);
            o.push_str("=\"");
            escape_attr(v, o);
            o.push('"');
        }
        if e.children.is_empty() {
            o.push_str("/>");
            return;
        }
        o.push('>');
        for n in &e.children {
            match n {
                Node::Elem(c) => write_elem(c, o),
                Node::Text(t) | Node::Raw(t) => o.push_str(t),
            }
        }
        o.push_str("</");
        o.push_str(&e.name);
        o.push('>');
    }

    pub fn to_bytes(d: &Doc) -> Vec<u8> {
        let mut o = String::from("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n");
        for p in &d.prolog {
            o.push_str(p);
        }
        write_elem(&d.root, &mut o);
        o.into_bytes()
    }

    /// Обход потомков (сам `el` не посещается); `f` возвращает false, чтобы удалить элемент.
    pub fn walk(el: &mut Elem, f: &mut dyn FnMut(&mut Elem) -> bool) {
        el.children.retain_mut(|n| match n {
            Node::Elem(c) => {
                if !f(c) {
                    return false;
                }
                walk(c, f);
                true
            }
            _ => true,
        });
    }

    /// Обход элемента и всех его потомков.
    pub fn walk_ref(el: &Elem, f: &mut dyn FnMut(&Elem)) {
        f(el);
        for c in el.elems() {
            walk_ref(c, f);
        }
    }
}

// ───────────────────────────── Пакет OPC ─────────────────────────────

struct Pkg {
    entries: Vec<(String, Vec<u8>)>,
    idx: HashMap<String, usize>,
}

impl Pkg {
    fn from_entries(entries: Vec<(String, Vec<u8>)>) -> Pkg {
        let idx = entries.iter().enumerate().map(|(i, (n, _))| (n.to_lowercase(), i)).collect();
        Pkg { entries, idx }
    }

    fn read(path: &Path) -> Result<Pkg> {
        let f = std::fs::File::open(path).with_context(|| format!("не удалось открыть {}", path.display()))?;
        let mut zip = zip::ZipArchive::new(f)
            .with_context(|| format!("{} — не docx (файл повреждён или защищён паролем)", path.display()))?;
        let mut entries: Vec<(String, Vec<u8>)> = Vec::new();
        let mut seen = HashSet::new();
        for i in 0..zip.len() {
            let mut e = zip.by_index(i).with_context(|| format!("{} — не удалось прочитать архив", path.display()))?;
            if e.is_dir() {
                continue;
            }
            let name = e.name().map(|n| n.to_string()).unwrap_or_default().replace('\\', "/");
            let name = name.trim_start_matches('/').to_string();
            if name.is_empty() || !seen.insert(name.to_lowercase()) {
                continue;
            }
            let mut buf = Vec::new();
            e.read_to_end(&mut buf)?;
            entries.push((name, buf));
        }
        Ok(Pkg::from_entries(entries))
    }

    fn get(&self, name: &str) -> Option<&[u8]> {
        self.idx.get(&name.to_lowercase()).map(|&i| self.entries[i].1.as_slice())
    }

    /// Имя записи в архиве (с настоящим регистром), если она есть.
    fn name_of(&self, name: &str) -> Option<String> {
        self.idx.get(&name.to_lowercase()).map(|&i| self.entries[i].0.clone())
    }

    fn set(&mut self, name: &str, data: Vec<u8>) {
        match self.idx.get(&name.to_lowercase()) {
            Some(&i) => self.entries[i].1 = data,
            None => {
                self.idx.insert(name.to_lowercase(), self.entries.len());
                self.entries.push((name.to_string(), data));
            }
        }
    }

    fn write(&self, path: &Path) -> Result<()> {
        let f = std::fs::File::create(path).with_context(|| format!("не удалось создать {}", path.display()))?;
        let mut zip = zip::ZipWriter::new(std::io::BufWriter::new(f));
        let opts = zip::write::SimpleFileOptions::default().compression_method(zip::CompressionMethod::Deflated);
        // [Content_Types].xml и корневые связи — первыми, остальное в исходном порядке.
        let first = ["[Content_Types].xml", "_rels/.rels"];
        let mut order: Vec<&(String, Vec<u8>)> = vec![];
        for f in first {
            if let Some(&i) = self.idx.get(&f.to_lowercase()) {
                order.push(&self.entries[i]);
            }
        }
        for e in &self.entries {
            if !first.iter().any(|f| f.eq_ignore_ascii_case(&e.0)) {
                order.push(e);
            }
        }
        for (name, data) in order {
            zip.start_file(name.as_str(), opts.large_file(data.len() as u64 > 3_900_000_000))?;
            zip.write_all(data)?;
        }
        zip.finish()?.flush()?;
        Ok(())
    }
}

fn dir_of(p: &str) -> &str {
    p.rfind('/').map(|i| &p[..i]).unwrap_or("")
}

fn file_of(p: &str) -> &str {
    p.rsplit('/').next().unwrap_or(p)
}

fn join(dir: &str, f: &str) -> String {
    if dir.is_empty() { f.to_string() } else { format!("{dir}/{f}") }
}

fn normalize(p: &str) -> String {
    let mut out: Vec<&str> = vec![];
    for seg in p.split('/') {
        match seg {
            "" | "." => {}
            ".." => {
                out.pop();
            }
            s => out.push(s),
        }
    }
    out.join("/")
}

/// Относительный путь `to` из каталога `from_dir` (оба без ведущего «/»).
fn relative(from_dir: &str, to: &str) -> String {
    let f: Vec<&str> = from_dir.split('/').filter(|s| !s.is_empty()).collect();
    let t: Vec<&str> = to.split('/').collect();
    let mut i = 0;
    while i < f.len() && i + 1 < t.len() && f[i] == t[i] {
        i += 1;
    }
    let mut parts: Vec<&str> = vec![".."; f.len() - i];
    parts.extend(&t[i..]);
    parts.join("/")
}

fn rels_path(part: &str) -> String {
    join(&join(dir_of(part), "_rels"), &format!("{}.rels", file_of(part)))
}

fn ext_of(p: &str) -> &str {
    file_of(p).rsplit_once('.').map(|x| x.1).unwrap_or("")
}

fn pct_decode(s: &str) -> String {
    let b = s.as_bytes();
    let mut out = Vec::with_capacity(b.len());
    let mut i = 0;
    while i < b.len() {
        if b[i] == b'%' && i + 3 <= b.len() {
            if let Some(v) = std::str::from_utf8(&b[i + 1..i + 3]).ok().and_then(|h| u8::from_str_radix(h, 16).ok()) {
                out.push(v);
                i += 3;
                continue;
            }
        }
        out.push(b[i]);
        i += 1;
    }
    String::from_utf8_lossy(&out).into_owned()
}

fn pct_encode_path(s: &str) -> String {
    let mut o = String::new();
    for &b in s.as_bytes() {
        if b.is_ascii_alphanumeric() || b"-._~/!$&'()*+,;=:@".contains(&b) {
            o.push(b as char);
        } else {
            o.push_str(&format!("%{b:02X}"));
        }
    }
    o
}

/// Имя существующей части пакета, на которую указывает связь `target` из части `owner`.
fn resolve_part(pkg: &Pkg, owner: &str, target: &str) -> Option<String> {
    let target = target.split('#').next().unwrap_or(target);
    let raw = match target.strip_prefix('/') {
        Some(t) => normalize(t),
        None => normalize(&join(dir_of(owner), target)),
    };
    pkg.name_of(&pct_decode(&raw)).or_else(|| pkg.name_of(&raw))
}

fn rel_suffix(ty: &str) -> &str {
    ty.rsplit('/').next().unwrap_or(ty)
}

#[derive(Clone, Debug)]
struct Rel {
    id: String,
    ty: String,
    target: String,
    external: bool,
}

#[derive(Clone, Debug, Default)]
struct Rels {
    items: Vec<Rel>,
    ids: HashSet<String>,
    next: u64,
    dirty: bool,
}

impl Rels {
    fn parse(bytes: &[u8]) -> Result<Rels> {
        let doc = xml::parse(bytes)?;
        let mut r = Rels::default();
        for e in doc.root.elems().filter(|e| e.local() == "Relationship") {
            let (Some(id), Some(ty), Some(target)) = (e.attr("Id"), e.attr("Type"), e.attr("Target")) else {
                continue;
            };
            r.push(Rel {
                id: id.to_string(),
                ty: ty.to_string(),
                target: target.to_string(),
                external: e.attr("TargetMode").is_some_and(|m| m.eq_ignore_ascii_case("External")),
            });
        }
        r.dirty = false;
        Ok(r)
    }

    fn push(&mut self, rel: Rel) {
        if let Some(n) = rel.id.strip_prefix("rId").and_then(|n| n.parse::<u64>().ok()) {
            self.next = self.next.max(n + 1);
        }
        self.ids.insert(rel.id.clone());
        self.items.push(rel);
        self.dirty = true;
    }

    fn get(&self, id: &str) -> Option<&Rel> {
        self.items.iter().find(|r| r.id == id)
    }

    /// Добавляет связь с новым свободным идентификатором.
    fn add(&mut self, ty: &str, target: &str, external: bool) -> String {
        if self.next == 0 {
            self.next = 1;
        }
        loop {
            let id = format!("rId{}", self.next);
            self.next += 1;
            if !self.ids.contains(&id) {
                self.push(Rel { id: id.clone(), ty: ty.to_string(), target: target.to_string(), external });
                return id;
            }
        }
    }

    fn to_bytes(&self) -> Vec<u8> {
        let mut root = Elem::new("Relationships").with("xmlns", "http://schemas.openxmlformats.org/package/2006/relationships");
        for r in &self.items {
            let mut e = Elem::new("Relationship").with("Id", &r.id).with("Type", &r.ty).with("Target", &r.target);
            if r.external {
                e.set_attr("TargetMode", "External");
            }
            root.push(e);
        }
        xml::to_bytes(&xml::Doc { prolog: vec![], root })
    }
}

#[derive(Clone, Debug, Default)]
struct Types {
    defaults: Vec<(String, String)>,
    overrides: Vec<(String, String)>,
}

impl Types {
    fn parse(bytes: &[u8]) -> Result<Types> {
        let doc = xml::parse(bytes)?;
        let mut t = Types::default();
        for e in doc.root.elems() {
            match (e.local(), e.attr("Extension"), e.attr("PartName"), e.attr("ContentType")) {
                ("Default", Some(x), _, Some(c)) => t.defaults.push((x.to_lowercase(), c.to_string())),
                ("Override", _, Some(p), Some(c)) => t.overrides.push((p.to_string(), c.to_string())),
                _ => {}
            }
        }
        Ok(t)
    }

    fn ct_override(&self, part: &str) -> Option<&str> {
        let p = format!("/{part}");
        self.overrides.iter().find(|(n, _)| n.eq_ignore_ascii_case(&p)).map(|(_, c)| c.as_str())
    }

    fn ct_default(&self, ext: &str) -> Option<&str> {
        self.defaults.iter().find(|(e, _)| e.eq_ignore_ascii_case(ext)).map(|(_, c)| c.as_str())
    }

    /// Гарантирует тип содержимого части: через `Default` по расширению (если разрешено и
    /// расширение свободно) или через `Override`.
    fn ensure(&mut self, part: &str, ct: &str, as_default: bool) {
        let ext = ext_of(part).to_lowercase();
        let cur = self.ct_override(part).or_else(|| self.ct_default(&ext));
        if cur.is_some_and(|c| c.eq_ignore_ascii_case(ct)) {
            return;
        }
        if as_default && !ext.is_empty() && self.ct_default(&ext).is_none() && self.ct_override(part).is_none() {
            self.defaults.push((ext, ct.to_string()));
        } else {
            self.set_override(part, ct);
        }
    }

    fn set_override(&mut self, part: &str, ct: &str) {
        let p = format!("/{part}");
        match self.overrides.iter_mut().find(|(n, _)| n.eq_ignore_ascii_case(&p)) {
            Some(o) => o.1 = ct.to_string(),
            None => self.overrides.push((p, ct.to_string())),
        }
    }

    fn to_bytes(&self) -> Vec<u8> {
        let mut root = Elem::new("Types").with("xmlns", "http://schemas.openxmlformats.org/package/2006/content-types");
        for (e, c) in &self.defaults {
            root.push(Elem::new("Default").with("Extension", e).with("ContentType", c));
        }
        for (p, c) in &self.overrides {
            root.push(Elem::new("Override").with("PartName", p).with("ContentType", c));
        }
        xml::to_bytes(&xml::Doc { prolog: vec![], root })
    }
}

fn builtin_ct(ext: &str) -> Option<&'static str> {
    Some(match ext.to_lowercase().as_str() {
        "png" => "image/png",
        "jpg" | "jpeg" | "jpe" => "image/jpeg",
        "gif" => "image/gif",
        "bmp" => "image/bmp",
        "tif" | "tiff" => "image/tiff",
        "emf" => "image/x-emf",
        "wmf" => "image/x-wmf",
        "svg" => "image/svg+xml",
        "webp" => "image/webp",
        "xml" => "application/xml",
        "bin" => "application/vnd.openxmlformats-officedocument.oleObject",
        "xlsx" => "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        "docx" => "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        "pptx" => "application/vnd.openxmlformats-officedocument.presentationml.presentation",
        "htm" | "html" => "text/html",
        "rtf" => "application/rtf",
        "txt" => "text/plain",
        _ => return None,
    })
}

// ───────────────────────────── Пространства имён ─────────────────────────────

const CANON: &[(&str, &str)] = &[
    ("http://schemas.openxmlformats.org/wordprocessingml/2006/main", "w"),
    ("http://schemas.openxmlformats.org/officeDocument/2006/relationships", "r"),
    ("http://schemas.openxmlformats.org/drawingml/2006/main", "a"),
    ("http://schemas.openxmlformats.org/drawingml/2006/picture", "pic"),
    ("http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing", "wp"),
    ("http://schemas.openxmlformats.org/drawingml/2006/chart", "c"),
    ("http://schemas.openxmlformats.org/drawingml/2006/diagram", "dgm"),
    ("http://schemas.openxmlformats.org/markup-compatibility/2006", "mc"),
    ("http://schemas.openxmlformats.org/officeDocument/2006/math", "m"),
    ("urn:schemas-microsoft-com:vml", "v"),
    ("urn:schemas-microsoft-com:office:office", "o"),
    ("urn:schemas-microsoft-com:office:word", "w10"),
    ("http://schemas.microsoft.com/office/word/2010/wordml", "w14"),
    ("http://schemas.microsoft.com/office/word/2012/wordml", "w15"),
    ("http://schemas.microsoft.com/office/word/2006/wordml", "wne"),
    ("http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing", "wp14"),
    ("http://schemas.microsoft.com/office/word/2010/wordprocessingShape", "wps"),
    ("http://schemas.microsoft.com/office/word/2010/wordprocessingGroup", "wpg"),
    ("http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas", "wpc"),
    ("http://schemas.microsoft.com/office/word/2010/wordprocessingInk", "wpi"),
    ("http://schemas.microsoft.com/office/word/2018/wordml", "w16"),
    ("http://schemas.microsoft.com/office/word/2016/wordml/cid", "w16cid"),
    ("http://schemas.microsoft.com/office/word/2018/wordml/cex", "w16cex"),
    ("http://schemas.microsoft.com/office/word/2015/wordml/symex", "w16se"),
    ("http://schemas.microsoft.com/office/word/2020/wordml/sdtdatahash", "w16sdtdh"),
    ("http://schemas.microsoft.com/office/word/2023/wordml/word16du", "w16du"),
    ("http://schemas.microsoft.com/office/drawing/2008/diagram", "dsp"),
    ("http://schemas.microsoft.com/office/drawing/2010/main", "a14"),
    ("http://schemas.microsoft.com/office/drawing/2014/chartex", "cx"),
    ("http://schemas.openxmlformats.org/drawingml/2006/chartDrawing", "cdr"),
    ("http://schemas.openxmlformats.org/drawingml/2006/lockedCanvas", "lc"),
];

fn canon_prefix(uri: &str) -> Option<&'static str> {
    CANON.iter().find(|(u, _)| *u == uri).map(|(_, p)| *p)
}

fn is_canon_prefix(p: &str) -> bool {
    CANON.iter().any(|(_, c)| *c == p)
}

fn rename_tokens(v: &str, map: &HashMap<String, String>) -> String {
    v.split_whitespace().map(|t| map.get(t).map(|s| s.as_str()).unwrap_or(t)).collect::<Vec<_>>().join(" ")
}

fn rename_prefixes(e: &mut Elem, map: &HashMap<String, String>) {
    if let Some((p, l)) = e.name.split_once(':') {
        if let Some(t) = map.get(p) {
            e.name = format!("{t}:{l}");
        }
    }
    for (k, v) in e.attrs.iter_mut() {
        if let Some(p) = k.strip_prefix("xmlns:") {
            if let Some(t) = map.get(p) {
                *k = format!("xmlns:{t}");
            }
        } else if let Some((p, l)) = k.split_once(':') {
            if l == "Ignorable" || l == "Requires" {
                *v = rename_tokens(v, map);
            }
            if let Some(t) = map.get(p) {
                *k = format!("{t}:{l}");
            }
        }
    }
    for n in e.children.iter_mut() {
        if let Node::Elem(c) = n {
            rename_prefixes(c, map);
        }
    }
}

/// Приводит префиксы пространств имён корня к общепринятым (`w`, `r`, `wp`, ...), чтобы остальной
/// код мог опираться на имена вида `w:p`. Strict OOXML и пространство имён по умолчанию не поддерживаются.
fn harmonize(doc: &mut xml::Doc) -> Result<()> {
    let root = &mut doc.root;
    if !root.name.contains(':') {
        bail!("неподдерживаемая структура XML (пространство имён по умолчанию)");
    }
    let decls: Vec<(String, String)> = root
        .attrs
        .iter()
        .filter_map(|(k, v)| k.strip_prefix("xmlns:").map(|p| (p.to_string(), v.clone())))
        .collect();
    if decls.iter().any(|(_, u)| u.starts_with("http://purl.oclc.org/ooxml/")) {
        bail!("формат Strict Open XML не поддерживается (сохраните документ как обычный .docx)");
    }
    let mut rename: HashMap<String, String> = HashMap::new();
    let mut taken: HashMap<String, String> = HashMap::new();
    for (p, u) in &decls {
        if let Some(c) = canon_prefix(u) {
            taken.insert(c.to_string(), u.clone());
            if p != c {
                rename.insert(p.clone(), c.to_string());
            }
        }
    }
    let mut n = 0;
    for (p, u) in &decls {
        if canon_prefix(u).is_some() {
            continue;
        }
        let mut t = p.clone();
        while is_canon_prefix(&t) || taken.get(&t).is_some_and(|x| x != u) {
            n += 1;
            t = format!("x{n}");
        }
        taken.insert(t.clone(), u.clone());
        if &t != p {
            rename.insert(p.clone(), t);
        }
    }
    // Объявления внутри документа: допустимы, только если они не конфликтуют с общепринятыми префиксами.
    let mut nested = false;
    let mut bad: Option<String> = None;
    let root_ptr = root as *const Elem;
    xml::walk_ref(root, &mut |e| {
        if std::ptr::eq(e, root_ptr) {
            return;
        }
        for (k, v) in &e.attrs {
            if let Some(p) = k.strip_prefix("xmlns:") {
                nested = true;
                if is_canon_prefix(p) && canon_prefix(v) != Some(p) {
                    bad = Some(p.to_string());
                }
            }
        }
    });
    if let Some(p) = bad {
        bail!("неподдерживаемая структура XML (префикс {p} переопределён внутри документа)");
    }
    if !rename.is_empty() {
        if nested {
            bail!("неподдерживаемая структура XML (нестандартные префиксы пространств имён)");
        }
        rename_prefixes(root, &rename);
        let mut seen = HashSet::new();
        root.attrs.retain(|(k, _)| !k.starts_with("xmlns:") || seen.insert(k.clone()));
    }
    Ok(())
}

/// Добавляет в `dst` объявления xmlns и `mc:Ignorable` из `src`.
fn merge_root_decls(dst: &mut Elem, src: &Elem) -> Result<()> {
    for (k, v) in &src.attrs {
        if k.starts_with("xmlns:") {
            match dst.attr(k) {
                None => dst.attrs.push((k.clone(), v.clone())),
                Some(x) if x == v => {}
                Some(_) => bail!("конфликт пространств имён ({k})"),
            }
        }
    }
    if let Some(ign) = src.attr("mc:Ignorable") {
        let mut toks: Vec<String> = dst.attr("mc:Ignorable").unwrap_or("").split_whitespace().map(String::from).collect();
        for t in ign.split_whitespace() {
            if !toks.iter().any(|x| x == t) {
                toks.push(t.to_string());
            }
        }
        dst.set_attr("mc:Ignorable", toks.join(" "));
        if dst.attr("xmlns:mc").is_none() {
            dst.attrs.insert(0, ("xmlns:mc".into(), "http://schemas.openxmlformats.org/markup-compatibility/2006".into()));
        }
    }
    Ok(())
}

// ───────────────────────────── Идентификаторы ─────────────────────────────

#[derive(Default)]
struct Pool {
    used: HashSet<i64>,
    next: i64,
}

impl Pool {
    fn add(&mut self, v: i64) {
        self.used.insert(v);
        if v >= self.next {
            self.next = v.saturating_add(1);
        }
    }
    fn fresh(&mut self) -> i64 {
        if self.next < 1 {
            self.next = 1;
        }
        while self.used.contains(&self.next) && self.next < i64::MAX {
            self.next += 1;
        }
        let v = self.next;
        self.add(v);
        v
    }
    fn keep_or_fresh(&mut self, v: i64) -> i64 {
        if self.used.contains(&v) {
            self.fresh()
        } else {
            self.add(v);
            v
        }
    }
}

#[derive(Default)]
struct HexPool {
    used: HashSet<u32>,
    next: u32,
}

impl HexPool {
    fn add_str(&mut self, s: &str) {
        if let Ok(v) = u32::from_str_radix(s.trim(), 16) {
            self.used.insert(v);
        }
    }
    fn fresh(&mut self) -> u32 {
        if self.next < 0x1000_0000 {
            self.next = 0x1000_0000;
        }
        while self.used.contains(&self.next) {
            self.next += 1;
            if self.next >= 0x7FFF_FFFF {
                self.next = 0x1000_0000;
            }
        }
        self.used.insert(self.next);
        self.next
    }
    /// Значение для `old`: прежнее, если оно свободно, иначе новое.
    fn map(&mut self, old: &str) -> String {
        match u32::from_str_radix(old.trim(), 16) {
            Ok(v) if old.len() == 8 && v < 0x8000_0000 && !self.used.contains(&v) => {
                self.used.insert(v);
                old.to_string()
            }
            _ => format!("{:08X}", self.fresh()),
        }
    }
}

/// Элементы-аннотации с уникальным `w:id` (отслеживание правок).
const ANN_ELEMS: &[&str] = &[
    "w:ins", "w:del", "w:moveFrom", "w:moveTo", "w:rPrChange", "w:pPrChange", "w:sectPrChange", "w:tblPrChange",
    "w:tblPrExChange", "w:trPrChange", "w:tcPrChange", "w:tblGridChange", "w:numberingChange", "w:cellIns",
    "w:cellDel", "w:cellMerge",
];

/// Парные элементы диапазонов: (семейство, начало, конец).
const RANGE_ELEMS: &[(&str, &str, &str)] = &[
    ("bm", "w:bookmarkStart", "w:bookmarkEnd"),
    ("mf", "w:moveFromRangeStart", "w:moveFromRangeEnd"),
    ("mt", "w:moveToRangeStart", "w:moveToRangeEnd"),
    ("pm", "w:permStart", "w:permEnd"),
];

#[derive(Default)]
struct Ids {
    ann: Pool,
    docpr: Pool,
    sdt: Pool,
    para: HexPool,
    durable: HexPool,
    nsid: HexPool,
    bm_names: HashSet<String>,
    vml: HashSet<String>,
}

impl Ids {
    fn scan(&mut self, root: &Elem) {
        xml::walk_ref(root, &mut |e| {
            let n = e.name.as_str();
            let id = |k: &str| e.attr(k).and_then(|v| v.trim().parse::<i64>().ok());
            if n == "wp:docPr" {
                if let Some(v) = id("id") {
                    self.docpr.add(v);
                }
            } else if n == "w:id" {
                if let Some(v) = e.val().and_then(|v| v.trim().parse::<i64>().ok()) {
                    self.sdt.add(v);
                }
            } else if n == "w:nsid" {
                if let Some(v) = e.val() {
                    self.nsid.add_str(v);
                }
            } else if ANN_ELEMS.contains(&n) || RANGE_ELEMS.iter().any(|r| r.1 == n || r.2 == n) {
                if let Some(v) = id("w:id") {
                    self.ann.add(v);
                }
                if n == "w:bookmarkStart" {
                    if let Some(name) = e.attr("w:name") {
                        self.bm_names.insert(name.to_lowercase());
                    }
                }
            }
            for (k, v) in &e.attrs {
                match k.as_str() {
                    "w14:paraId" | "w15:paraId" | "w16cid:paraId" => self.para.add_str(v),
                    "w16cid:durableId" | "w16cex:durableId" => self.durable.add_str(v),
                    _ => {}
                }
            }
            if e.prefix() == "v" {
                if let Some(i) = e.attr("id") {
                    self.vml.insert(i.to_string());
                }
            }
        });
    }
}

// ───────────────────────────── Исходный документ ─────────────────────────────

const MISSING_RID: &str = "rIdMissing";
const CT_DOC: &str = "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml";

struct KindInfo {
    key: &'static str,
    ty: &'static str,
    ct: &'static str,
    file: &'static str,
}

const KINDS: &[KindInfo] = &[
    KindInfo {
        key: "styles",
        ty: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles",
        ct: "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml",
        file: "styles.xml",
    },
    KindInfo {
        key: "numbering",
        ty: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering",
        ct: "application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml",
        file: "numbering.xml",
    },
    KindInfo {
        key: "settings",
        ty: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings",
        ct: "application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml",
        file: "settings.xml",
    },
    KindInfo {
        key: "footnotes",
        ty: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes",
        ct: "application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml",
        file: "footnotes.xml",
    },
    KindInfo {
        key: "endnotes",
        ty: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes",
        ct: "application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml",
        file: "endnotes.xml",
    },
    KindInfo {
        key: "comments",
        ty: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments",
        ct: "application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml",
        file: "comments.xml",
    },
    KindInfo {
        key: "commentsExtended",
        ty: "http://schemas.microsoft.com/office/2011/relationships/commentsExtended",
        ct: "application/vnd.openxmlformats-officedocument.wordprocessingml.commentsExtended+xml",
        file: "commentsExtended.xml",
    },
    KindInfo {
        key: "commentsIds",
        ty: "http://schemas.microsoft.com/office/2016/09/relationships/commentsIds",
        ct: "application/vnd.openxmlformats-officedocument.wordprocessingml.commentsIds+xml",
        file: "commentsIds.xml",
    },
    KindInfo {
        key: "commentsExtensible",
        ty: "http://schemas.microsoft.com/office/2018/08/relationships/commentsExtensible",
        ct: "application/vnd.openxmlformats-officedocument.wordprocessingml.commentsExtensible+xml",
        file: "commentsExtensible.xml",
    },
    KindInfo {
        key: "people",
        ty: "http://schemas.microsoft.com/office/2011/relationships/people",
        ct: "application/vnd.openxmlformats-officedocument.wordprocessingml.people+xml",
        file: "people.xml",
    },
];

fn kind(key: &str) -> &'static KindInfo {
    KINDS.iter().find(|k| k.key == key).expect("известный вид части")
}

struct Src {
    pkg: Pkg,
    /// Имя файла для сообщений.
    name: String,
    main: String,
    types: Types,
    rels: Rels,
    doc: xml::Doc,
}

impl Src {
    fn load(path: &Path) -> Result<Src> {
        let pkg = Pkg::read(path)?;
        let name = path.file_name().map(|n| n.to_string_lossy().into_owned()).unwrap_or_default();
        let types = Types::parse(pkg.get("[Content_Types].xml").ok_or_else(|| anyhow!("в docx нет [Content_Types].xml"))?)?;
        let root_rels = match pkg.get("_rels/.rels") {
            Some(b) => Rels::parse(b)?,
            None => Rels::default(),
        };
        let main = root_rels
            .items
            .iter()
            .filter(|r| !r.external && rel_suffix(&r.ty) == "officeDocument")
            .find_map(|r| resolve_part(&pkg, "", &r.target))
            .or_else(|| pkg.name_of("word/document.xml"))
            .ok_or_else(|| anyhow!("в docx нет основной части документа (word/document.xml)"))?;
        let rels = match pkg.get(&rels_path(&main)) {
            Some(b) => Rels::parse(b)?,
            None => Rels::default(),
        };
        let mut doc = xml::parse(pkg.get(&main).unwrap_or_default()).with_context(|| format!("{main} — {name}"))?;
        harmonize(&mut doc).with_context(|| format!("{main} — {name}"))?;
        if doc.root.name != "w:document" {
            bail!("{name} — не документ Word");
        }
        if doc.root.child("w:body").is_none() {
            bail!("в документе {name} нет <w:body>");
        }
        Ok(Src { pkg, name, main, types, rels, doc })
    }

    /// Путь части документа заданного вида (по связям основной части).
    fn part_path(&self, key: &str) -> Option<String> {
        let ty = kind(key).ty;
        let suffix = rel_suffix(ty);
        self.rels
            .items
            .iter()
            .filter(|r| !r.external && rel_suffix(&r.ty) == suffix)
            .find_map(|r| resolve_part(&self.pkg, &self.main, &r.target))
    }

    fn load_xml(&self, path: &str) -> Result<xml::Doc> {
        let mut d = xml::parse(self.pkg.get(path).unwrap_or_default()).with_context(|| format!("{path} — {}", self.name))?;
        harmonize(&mut d).with_context(|| format!("{path} — {}", self.name))?;
        Ok(d)
    }
}

// ───────────────────────────── Слияние ─────────────────────────────

struct Part {
    path: String,
    doc: xml::Doc,
    dirty: bool,
}

/// Состояние переноса одного исходного документа.
#[derive(Default)]
struct Ctx {
    file: String,
    style_map: HashMap<String, String>,
    inject_pstyle: Option<String>,
    num_off: u64,
    abs_off: u64,
    pic_off: u64,
    /// "footnotes" / "endnotes" / "comments" -> старый id -> новый id
    note_maps: HashMap<&'static str, HashMap<String, String>>,
    para_map: HashMap<String, String>,
    durable_map: HashMap<String, String>,
    bm_rename: HashMap<String, String>,
    ranges: HashMap<(&'static str, String), String>,
    vml_map: HashMap<String, String>,
    rel_memo: HashMap<(String, String), String>,
    part_memo: HashMap<String, String>,
    srels: HashMap<String, Rels>,
    empty_story: [Option<String>; 2],
    dropped_bindings: usize,
    dropped_refs: usize,
    k: usize,
}

struct Mode {
    /// (часть-источник, часть-приёмник): ссылки r:* перепривязываются к связям приёмника.
    rels: Option<(String, String)>,
}

struct Merger {
    pkg: Pkg,
    types: Types,
    main: String,
    doc: xml::Doc,
    rels: HashMap<String, Rels>,
    parts: HashMap<&'static str, Part>,
    used: HashSet<String>,
    hint: HashMap<String, usize>,
    ids: Ids,
    body: Vec<Node>,
    final_sect: Elem,
    /// Определённые на данный момент ссылки на колонтитулы: (это нижний?, тип).
    hf_state: HashSet<(bool, String)>,
    warnings: Vec<String>,
    touched: bool,
}

const HF_TYPES: [&str; 3] = ["default", "first", "even"];

fn default_sect() -> Elem {
    Elem::new("w:sectPr")
        .with_child(Elem::new("w:pgSz").with("w:w", "11906").with("w:h", "16838"))
        .with_child(
            Elem::new("w:pgMar")
                .with("w:top", "1134")
                .with("w:right", "850")
                .with("w:bottom", "1134")
                .with("w:left", "1701")
                .with("w:header", "708")
                .with("w:footer", "708")
                .with("w:gutter", "0"),
        )
}

fn is_ws_text(n: &Node) -> bool {
    matches!(n, Node::Text(t) if t.trim().is_empty())
}

fn para_sect_mut(n: &mut Node) -> Option<&mut Elem> {
    match n {
        Node::Elem(p) if p.name == "w:p" => p.child_mut("w:pPr")?.child_mut("w:sectPr"),
        _ => None,
    }
}

fn para_sect(n: &Node) -> Option<&Elem> {
    match n {
        Node::Elem(p) if p.name == "w:p" => p.child("w:pPr")?.child("w:sectPr"),
        _ => None,
    }
}

fn sect_refs(s: &Elem) -> Vec<(bool, String)> {
    s.elems()
        .filter(|e| e.name == "w:headerReference" || e.name == "w:footerReference")
        .map(|e| (e.name == "w:footerReference", e.attr("w:type").unwrap_or("default").to_string()))
        .collect()
}

/// LibreOffice иногда пишет `sectPr` в начало `pPr`; по схеме он идёт после `rPr`, перед `pPrChange`.
fn fix_ppr_order(ppr: &mut Elem) {
    let Some(i) = ppr.children.iter().position(|n| matches!(n, Node::Elem(e) if e.name == "w:sectPr")) else { return };
    let after_ok = ppr.children[i + 1..].iter().all(|n| match n {
        Node::Elem(e) => e.name == "w:pPrChange",
        _ => true,
    });
    if after_ok {
        return;
    }
    let sect = ppr.children.remove(i);
    let pos = ppr
        .children
        .iter()
        .position(|n| matches!(n, Node::Elem(e) if e.name == "w:pPrChange"))
        .unwrap_or(ppr.children.len());
    ppr.children.insert(pos, sect);
}

/// Ставит `sectPr` в конец абзаца (`pPr`), соблюдая порядок элементов схемы.
fn put_sect_in_para(p: &mut Elem, sect: Elem) {
    if p.child("w:pPr").is_none() {
        p.children.insert(0, Node::Elem(Elem::new("w:pPr")));
    }
    let ppr = p.child_mut("w:pPr").expect("pPr создан выше");
    let pos = ppr
        .children
        .iter()
        .position(|n| matches!(n, Node::Elem(e) if e.name == "w:pPrChange"))
        .unwrap_or(ppr.children.len());
    ppr.children.insert(pos, Node::Elem(sect));
}

/// Закрывает раздел: sectPr уходит в последний абзац содержимого (или в новый пустой абзац).
fn close_section(body: &mut Vec<Node>, sect: Elem) {
    const MARKS: &[&str] = &["w:bookmarkStart", "w:bookmarkEnd", "w:proofErr", "w:permStart", "w:permEnd"];
    for n in body.iter_mut().rev() {
        match n {
            Node::Elem(e) if MARKS.contains(&e.name.as_str()) => continue,
            Node::Elem(e) if e.name == "w:p" && e.child("w:pPr").and_then(|p| p.child("w:sectPr")).is_none() => {
                put_sect_in_para(e, sect);
                return;
            }
            Node::Text(_) | Node::Raw(_) => continue,
            _ => break,
        }
    }
    let mut p = Elem::new("w:p");
    put_sect_in_para(&mut p, sect);
    body.push(Node::Elem(p));
}

fn instr_rename(s: &str, map: &HashMap<String, String>) -> String {
    fn flush(tok: &mut String, out: &mut String, map: &HashMap<String, String>) {
        if tok.is_empty() {
            return;
        }
        let core = tok.trim_matches('"');
        let pre = &tok[..tok.len() - tok.trim_start_matches('"').len()];
        let post = &tok[tok.trim_end_matches('"').len()..];
        match map.get(&core.to_lowercase()) {
            Some(n) if !core.is_empty() => {
                out.push_str(pre);
                out.push_str(n);
                out.push_str(post);
            }
            _ => out.push_str(tok),
        }
        tok.clear();
    }
    let mut out = String::new();
    let mut tok = String::new();
    for ch in s.chars() {
        if ch.is_whitespace() {
            flush(&mut tok, &mut out, map);
            out.push(ch);
        } else {
            tok.push(ch);
        }
    }
    flush(&mut tok, &mut out, map);
    out
}

/// Каноническая запись элемента для сравнения определений (без rsid и служебных атрибутов).
fn canon_xml(e: &Elem, skip: &[&str], out: &mut String) {
    out.push('<');
    out.push_str(&e.name);
    let mut attrs: Vec<&(String, String)> =
        e.attrs
            .iter()
            .filter(|(k, _)| !k.starts_with("w:rsid") && !k.starts_with("w14:") && !k.starts_with("xmlns") && k != "w:styleId" && k != "w:default")
            .collect();
    attrs.sort();
    for (k, v) in attrs {
        out.push_str(&format!(" {k}={v:?}"));
    }
    out.push('>');
    for c in e.elems() {
        if skip.contains(&c.name.as_str()) || c.name == "w:rsid" {
            continue;
        }
        canon_xml(c, skip, out);
    }
    for n in &e.children {
        if let Node::Text(t) = n {
            if !t.trim().is_empty() {
                out.push_str(t.trim());
            }
        }
    }
    out.push_str("</>");
}

const STYLE_SKIP: &[&str] = &[
    "w:name", "w:aliases", "w:uiPriority", "w:qFormat", "w:semiHidden", "w:unhideWhenUsed", "w:locked", "w:hidden",
    "w:personal", "w:personalCompose", "w:personalReply", "w:autoRedefine", "w:link", "w:next", "w:basedOn",
];

fn style_sig(st: &Elem, names: &HashMap<String, String>) -> String {
    let mut s = String::new();
    if let Some(b) = st.child("w:basedOn").and_then(|b| b.val()) {
        s.push_str(&format!("basedOn={};", names.get(b).cloned().unwrap_or_else(|| b.to_lowercase())));
    }
    canon_xml(st, STYLE_SKIP, &mut s);
    s
}

fn truncate_chars(s: &str, n: usize) -> &str {
    match s.char_indices().nth(n) {
        Some((i, _)) => &s[..i],
        None => s,
    }
}

impl Merger {
    fn new(base: Src) -> Result<Merger> {
        let Src { pkg, types, main, rels, doc, .. } = base;
        let used = pkg.entries.iter().map(|(n, _)| n.to_lowercase()).collect();
        let mut m = Merger {
            pkg,
            types,
            main: main.clone(),
            doc,
            rels: HashMap::new(),
            parts: HashMap::new(),
            used,
            hint: HashMap::new(),
            ids: Ids::default(),
            body: vec![],
            final_sect: default_sect(),
            hf_state: HashSet::new(),
            warnings: vec![],
            touched: false,
        };
        m.rels.insert(main, rels);
        // Идентификаторы, уже занятые в основном документе, колонтитулах и примечаниях.
        m.ids.scan(&m.doc.root);
        for key in ["numbering", "footnotes", "endnotes", "comments", "commentsExtended", "commentsIds", "commentsExtensible"] {
            if m.load_part(key)? {
                m.ids.scan(&m.parts[key].doc.root);
            }
        }
        let hf: Vec<String> = m.rel_targets(&m.main.clone(), &["header", "footer"]);
        for p in hf {
            if let Some(b) = m.pkg.get(&p) {
                if let Ok(d) = xml::parse(b) {
                    m.ids.scan(&d.root);
                }
            }
        }
        let body = m.doc.root.child_mut("w:body").ok_or_else(|| anyhow!("нет <w:body>"))?;
        let mut kids = std::mem::take(&mut body.children);
        kids.retain(|n| !is_ws_text(n));
        let fin = kids.iter().rposition(|n| matches!(n, Node::Elem(e) if e.name == "w:sectPr"));
        let fin = fin.map(|i| kids.remove(i)).and_then(|n| if let Node::Elem(e) = n { Some(e) } else { None });
        m.final_sect = fin.unwrap_or_else(default_sect);
        m.body = kids;
        for n in &m.body {
            if let Some(s) = para_sect(n) {
                for r in sect_refs(s) {
                    m.hf_state.insert(r);
                }
            }
        }
        for r in sect_refs(&m.final_sect) {
            m.hf_state.insert(r);
        }
        Ok(m)
    }

    fn warn(&mut self, msg: String) {
        if !self.warnings.contains(&msg) {
            self.warnings.push(msg);
        }
    }

    fn rels_mut(&mut self, owner: &str) -> &mut Rels {
        if !self.rels.contains_key(owner) {
            let r = self.pkg.get(&rels_path(owner)).and_then(|b| Rels::parse(b).ok()).unwrap_or_default();
            self.rels.insert(owner.to_string(), r);
        }
        self.rels.get_mut(owner).expect("вставлено выше")
    }

    /// Пути частей, на которые ссылается `owner` связями заданных видов.
    fn rel_targets(&mut self, owner: &str, suffixes: &[&str]) -> Vec<String> {
        let items = self.rels_mut(owner).items.clone();
        items
            .iter()
            .filter(|r| !r.external && suffixes.contains(&rel_suffix(&r.ty)))
            .filter_map(|r| resolve_part(&self.pkg, owner, &r.target))
            .collect()
    }

    fn mdir(&self) -> String {
        dir_of(&self.main).to_string()
    }

    /// Загружает часть базового документа (если она есть).
    fn load_part(&mut self, key: &'static str) -> Result<bool> {
        if self.parts.contains_key(key) {
            return Ok(true);
        }
        let suffix = rel_suffix(kind(key).ty);
        let main = self.main.clone();
        let Some(path) = self.rel_targets(&main, &[suffix]).into_iter().next() else { return Ok(false) };
        let mut doc = xml::parse(self.pkg.get(&path).unwrap_or_default()).with_context(|| format!("{path} (первый документ)"))?;
        harmonize(&mut doc).with_context(|| format!("{path} (первый документ)"))?;
        self.parts.insert(key, Part { path, doc, dirty: false });
        Ok(true)
    }

    /// Создаёт часть в базовом документе (с типом содержимого и связью).
    fn create_part(&mut self, key: &'static str, root: Elem) -> String {
        let k = kind(key);
        let path = self.unique_name(&join(&self.mdir(), k.file));
        self.types.ensure(&path, k.ct, false);
        let target = pct_encode_path(&relative(&self.mdir(), &path));
        let main = self.main.clone();
        self.rels_mut(&main).add(k.ty, &target, false);
        self.parts.insert(key, Part { path: path.clone(), doc: xml::Doc { prolog: vec![], root }, dirty: true });
        path
    }

    fn unique_name(&mut self, want: &str) -> String {
        let want = want.trim_start_matches('/').to_string();
        if self.used.insert(want.to_lowercase()) {
            return want;
        }
        let dir = dir_of(&want).to_string();
        let file = file_of(&want);
        let (stem, ext) = match file.rsplit_once('.') {
            Some((s, e)) => (s.to_string(), format!(".{e}")),
            None => (file.to_string(), String::new()),
        };
        let prefix = stem.trim_end_matches(|c: char| c.is_ascii_digit()).to_string();
        let key = format!("{dir}/{prefix}{ext}").to_lowercase();
        let mut k = *self.hint.get(&key).unwrap_or(&1);
        loop {
            let cand = join(&dir, &format!("{prefix}{k}{ext}"));
            k += 1;
            if self.used.insert(cand.to_lowercase()) {
                self.hint.insert(key, k);
                return cand;
            }
        }
    }

    // ---------- перенос связанных частей ----------

    fn copy_type(&mut self, src: &Src, src_path: &str, new_path: &str) {
        let ext = ext_of(src_path).to_lowercase();
        let (ct, as_default) = if let Some(c) = src.types.ct_override(src_path) {
            // Обычные типы картинок и т.п. лучше объявлять через Default по расширению.
            (c.to_string(), builtin_ct(&ext).is_some_and(|b| b.eq_ignore_ascii_case(c)))
        } else if let Some(c) = src.types.ct_default(&ext) {
            (c.to_string(), true)
        } else if let Some(c) = builtin_ct(&ext) {
            (c.to_string(), true)
        } else {
            ("application/octet-stream".to_string(), false)
        };
        self.types.ensure(new_path, &ct, as_default);
    }

    /// Копирует часть `src_path` со всеми её связями; возвращает новый путь.
    fn import_part(&mut self, src: &Src, cx: &mut Ctx, src_path: &str, rel_ty: &str) -> Result<Option<String>> {
        if let Some(p) = cx.part_memo.get(&src_path.to_lowercase()) {
            return Ok(Some(p.clone()));
        }
        if src_path.eq_ignore_ascii_case(&src.main) {
            return Ok(None);
        }
        let Some(data) = src.pkg.get(src_path) else { return Ok(None) };
        let new_path = self.unique_name(src_path);
        cx.part_memo.insert(src_path.to_lowercase(), new_path.clone());
        self.copy_type(src, src_path, &new_path);
        let suffix = rel_suffix(rel_ty);
        let mut bytes = data.to_vec();
        if suffix == "header" || suffix == "footer" {
            let mut d = src.load_xml(src_path)?;
            self.plan_bookmarks(cx, &[&d.root]);
            let saved = std::mem::take(&mut cx.ranges);
            let mode = Mode { rels: None };
            xml::walk(&mut d.root, &mut |e| self.xform_elem(src, cx, &mode, e));
            cx.ranges = saved;
            bytes = xml::to_bytes(&d);
        } else if suffix == "diagramData" {
            bytes = self.patch_diagram_data(src, cx, &bytes);
        }
        if let Some(rb) = src.pkg.get(&rels_path(src_path)) {
            let srels = Rels::parse(rb).with_context(|| format!("{} — {}", rels_path(src_path), src.name))?;
            let mut out = Rels::default();
            for r in &srels.items {
                if r.external {
                    out.push(r.clone());
                    continue;
                }
                let child = resolve_part(&src.pkg, src_path, &r.target);
                let new_child = match child {
                    Some(c) => self.import_part(src, cx, &c, &r.ty)?,
                    None => None,
                };
                match new_child {
                    Some(nc) => out.push(Rel {
                        id: r.id.clone(),
                        ty: r.ty.clone(),
                        target: pct_encode_path(&relative(dir_of(&new_path), &nc)),
                        external: false,
                    }),
                    None => {
                        let f = src.name.clone();
                        self.warn(format!(
                            "«{f}»: часть «{}» ссылается на отсутствующий объект «{}» — связь пропущена.",
                            file_of(src_path),
                            r.target
                        ));
                    }
                }
            }
            if !out.items.is_empty() {
                self.pkg.set(&rels_path(&new_path), out.to_bytes());
                self.used.insert(rels_path(&new_path).to_lowercase());
            }
        }
        self.pkg.set(&new_path, bytes);
        Ok(Some(new_path))
    }

    /// В данных SmartArt `dsp:dataModelExt relId` ссылается на связь основного документа.
    fn patch_diagram_data(&mut self, src: &Src, cx: &mut Ctx, bytes: &[u8]) -> Vec<u8> {
        let Ok(s) = std::str::from_utf8(bytes) else { return bytes.to_vec() };
        let key = "relId=\"";
        let mut out = String::with_capacity(s.len());
        let mut rest = s;
        while let Some(i) = rest.find(key) {
            let a = i + key.len();
            let Some(l) = rest[a..].find('"') else { break };
            out.push_str(&rest[..a]);
            let old = rest[a..a + l].to_string();
            let main = src.main.clone();
            let dest = self.main.clone();
            out.push_str(&self.import_rel(src, cx, &main, &dest, &old));
            rest = &rest[a + l..];
        }
        out.push_str(rest);
        out.into_bytes()
    }

    /// Переносит связь `id` части `src_owner` в часть `dest_owner` базы; возвращает новый идентификатор.
    fn import_rel(&mut self, src: &Src, cx: &mut Ctx, src_owner: &str, dest_owner: &str, id: &str) -> String {
        let key = (src_owner.to_string(), id.to_string());
        if let Some(n) = cx.rel_memo.get(&key) {
            return n.clone();
        }
        if !cx.srels.contains_key(src_owner) {
            let r = if src_owner.eq_ignore_ascii_case(&src.main) {
                src.rels.clone()
            } else {
                src.pkg.get(&rels_path(src_owner)).and_then(|b| Rels::parse(b).ok()).unwrap_or_default()
            };
            cx.srels.insert(src_owner.to_string(), r);
        }
        let rel = cx.srels[src_owner].get(id).cloned();
        let file = src.name.clone();
        let new_id = match rel {
            None => {
                self.warn(format!("«{file}»: в документе есть ссылка на несуществующий объект ({id}) — ссылка удалена."));
                MISSING_RID.to_string()
            }
            Some(r) if r.external => self.rels_mut(dest_owner).add(&r.ty, &r.target, true),
            Some(r) => {
                let part = resolve_part(&src.pkg, src_owner, &r.target);
                let imported = match part {
                    Some(p) => self.import_part(src, cx, &p, &r.ty),
                    None => Ok(None),
                };
                match imported {
                    Ok(Some(np)) => {
                        let target = pct_encode_path(&relative(dir_of(dest_owner), &np));
                        self.rels_mut(dest_owner).add(&r.ty, &target, false)
                    }
                    Ok(None) => {
                        self.warn(format!(
                            "«{file}»: объект «{}» ({}) отсутствует в файле и не перенесён.",
                            r.target,
                            rel_suffix(&r.ty)
                        ));
                        MISSING_RID.to_string()
                    }
                    Err(e) => {
                        self.warn(format!("«{file}»: не удалось перенести объект «{}»: {e}", r.target));
                        MISSING_RID.to_string()
                    }
                }
            }
        };
        cx.rel_memo.insert(key, new_id.clone());
        new_id
    }

    // ---------- преобразование содержимого ----------

    fn plan_bookmarks(&mut self, cx: &mut Ctx, roots: &[&Elem]) {
        let mut names = vec![];
        for r in roots {
            xml::walk_ref(r, &mut |e| {
                if e.name == "w:bookmarkStart" {
                    if let Some(n) = e.attr("w:name") {
                        names.push(n.to_string());
                    }
                }
            });
        }
        for name in names {
            let low = name.to_lowercase();
            if cx.bm_rename.contains_key(&low) {
                continue;
            }
            if self.ids.bm_names.contains(&low) {
                let new = loop {
                    cx.k += 1;
                    let suf = format!("_m{}", cx.k);
                    let cand = format!("{}{suf}", truncate_chars(&name, 40usize.saturating_sub(suf.chars().count())));
                    if !self.ids.bm_names.contains(&cand.to_lowercase()) {
                        break cand;
                    }
                };
                self.ids.bm_names.insert(new.to_lowercase());
                cx.bm_rename.insert(low, new);
            } else {
                self.ids.bm_names.insert(low);
            }
        }
    }

    fn range_id(&mut self, cx: &mut Ctx, fam: &'static str, start: bool, old: &str) -> Option<String> {
        if start {
            let new = match old.trim().parse::<i64>() {
                Ok(v) => self.ids.ann.keep_or_fresh(v).to_string(),
                Err(_) => old.to_string(),
            };
            cx.ranges.insert((fam, old.to_string()), new.clone());
            Some(new)
        } else {
            cx.ranges.get(&(fam, old.to_string())).cloned()
        }
    }

    fn xform_elem(&mut self, src: &Src, cx: &mut Ctx, mode: &Mode, e: &mut Elem) -> bool {
        if let Some((so, dop)) = &mode.rels {
            let mut missing = false;
            let mut i = 0;
            while i < e.attrs.len() {
                if e.attrs[i].0.starts_with("r:") || e.attrs[i].0 == "o:relid" {
                    let old = e.attrs[i].1.clone();
                    let new = self.import_rel(src, cx, so, dop, &old);
                    if new == MISSING_RID {
                        // Ссылка в исходном файле «битая»: оставлять её нельзя (она могла бы
                        // указать на чужую связь), поэтому атрибут убирается.
                        e.attrs.remove(i);
                        missing = true;
                        continue;
                    }
                    e.attrs[i].1 = new;
                }
                i += 1;
            }
            if missing
                && matches!(e.name.as_str(), "w:headerReference" | "w:footerReference" | "c:chart" | "w:altChunk" | "dgm:relIds" | "o:OLEObject")
            {
                return false;
            }
        }
        let name = e.name.clone();
        match name.as_str() {
            "w:pStyle" | "w:rStyle" | "w:tblStyle" | "w:numStyleLink" | "w:styleLink" | "w:basedOn" | "w:next" | "w:link" => {
                if let Some(n) = e.val().and_then(|v| cx.style_map.get(v)) {
                    let n = n.clone();
                    e.set_attr("w:val", n);
                }
            }
            "w:numId" => {
                if let Some(v) = e.val().and_then(|v| v.trim().parse::<u64>().ok()) {
                    if v != 0 {
                        e.set_attr("w:val", v.saturating_add(cx.num_off).to_string());
                    }
                }
            }
            "w:footnoteReference" | "w:endnoteReference" | "w:commentRangeStart" | "w:commentRangeEnd" | "w:commentReference" => {
                let key = match name.as_str() {
                    "w:footnoteReference" => "footnotes",
                    "w:endnoteReference" => "endnotes",
                    _ => "comments",
                };
                let old = e.attr("w:id").unwrap_or("").to_string();
                match cx.note_maps.get(key).and_then(|m| m.get(&old)) {
                    Some(n) => e.set_attr("w:id", n.clone()),
                    None => {
                        cx.dropped_refs += 1;
                        return false;
                    }
                }
            }
            "wp:docPr" => {
                if let Some(v) = e.attr("id").and_then(|v| v.trim().parse::<i64>().ok()) {
                    let n = self.ids.docpr.keep_or_fresh(v);
                    e.set_attr("id", n.to_string());
                }
            }
            "w:id" => {
                if let Some(v) = e.val().and_then(|v| v.trim().parse::<i64>().ok()) {
                    let n = self.ids.sdt.keep_or_fresh(v);
                    e.set_attr("w:val", n.to_string());
                }
            }
            "w:dataBinding" => {
                cx.dropped_bindings += 1;
                return false;
            }
            "w:hyperlink" => {
                if let Some(a) = e.attr("w:anchor") {
                    if let Some(n) = cx.bm_rename.get(&a.to_lowercase()) {
                        let n = n.clone();
                        e.set_attr("w:anchor", n);
                    }
                }
            }
            "w:fldSimple" => {
                if let Some(a) = e.attr("w:instr") {
                    let n = instr_rename(a, &cx.bm_rename);
                    if n != a {
                        e.set_attr("w:instr", n);
                    }
                }
            }
            "w:instrText" => {
                if !cx.bm_rename.is_empty() {
                    let t = e.text();
                    let n = instr_rename(&t, &cx.bm_rename);
                    if n != t {
                        e.set_text(&n);
                    }
                }
            }
            "w:pPr" => fix_ppr_order(e),
            "w:p" => {
                if let Some(id) = &cx.inject_pstyle {
                    let has = e.child("w:pPr").is_some_and(|p| p.child("w:pStyle").is_some());
                    if !has {
                        if e.child("w:pPr").is_none() {
                            e.children.insert(0, Node::Elem(Elem::new("w:pPr")));
                        }
                        if let Some(ppr) = e.child_mut("w:pPr") {
                            ppr.children.insert(0, Node::Elem(Elem::new("w:pStyle").with("w:val", id)));
                        }
                    }
                }
            }
            "o:OLEObject" => {
                if let Some(n) = e.attr("ShapeID").and_then(|s| cx.vml_map.get(s)) {
                    let n = n.clone();
                    e.set_attr("ShapeID", n);
                }
            }
            _ => {}
        }
        if ANN_ELEMS.contains(&name.as_str()) {
            if let Some(v) = e.attr("w:id").and_then(|v| v.trim().parse::<i64>().ok()) {
                let n = self.ids.ann.keep_or_fresh(v);
                e.set_attr("w:id", n.to_string());
            }
        }
        for &(fam, s, en) in RANGE_ELEMS {
            if name == s || name == en {
                if let Some(old) = e.attr("w:id").map(String::from) {
                    if let Some(n) = self.range_id(cx, fam, name == s, &old) {
                        e.set_attr("w:id", n);
                    }
                }
                if fam == "bm" && name == s {
                    if let Some(nm) = e.attr("w:name").and_then(|n| cx.bm_rename.get(&n.to_lowercase())) {
                        let nm = nm.clone();
                        e.set_attr("w:name", nm);
                    }
                }
            }
        }
        for a in e.attrs.iter_mut() {
            if a.0 == "w14:paraId" {
                if !cx.para_map.contains_key(&a.1) {
                    let n = self.ids.para.map(&a.1);
                    cx.para_map.insert(a.1.clone(), n);
                }
                a.1 = cx.para_map[&a.1].clone();
            }
        }
        if e.prefix() == "v" && e.local() != "shapetype" {
            if let Some(id) = e.attr("id").map(String::from) {
                if self.ids.vml.contains(&id) {
                    let new = loop {
                        cx.k += 1;
                        let c = format!("_x0000_s{}", 90000 + cx.k);
                        if !self.ids.vml.contains(&c) {
                            break c;
                        }
                    };
                    self.ids.vml.insert(new.clone());
                    cx.vml_map.insert(id.clone(), new.clone());
                    e.set_attr("id", new.clone());
                    if e.attr("o:spid").is_some() {
                        e.set_attr("o:spid", new);
                    }
                } else {
                    self.ids.vml.insert(id);
                }
            }
        }
        true
    }

    // ---------- стили ----------

    fn merge_styles(&mut self, src: &Src, cx: &mut Ctx) -> Result<()> {
        let Some(spath) = src.part_path("styles") else { return Ok(()) };
        let sdoc = src.load_xml(&spath)?;
        if !self.load_part("styles")? {
            // В базе нет стилей: берём настройки по умолчанию, стили импортируются ниже.
            let mut root = sdoc.root.clone();
            root.children.retain(|n| matches!(n, Node::Elem(e) if e.name == "w:docDefaults" || e.name == "w:latentStyles"));
            self.create_part("styles", root);
        }
        let part = self.parts.get_mut("styles").expect("часть стилей загружена");
        let broot = &mut part.doc.root;
        let sroot = &sdoc.root;

        fn id_names(root: &Elem) -> HashMap<String, String> {
            root.elems()
                .filter(|e| e.name == "w:style")
                .filter_map(|e| {
                    let id = e.attr("w:styleId")?;
                    let n = e.child("w:name").and_then(|n| n.val()).unwrap_or(id);
                    Some((id.to_string(), n.to_lowercase()))
                })
                .collect()
        }
        let bnames = id_names(broot);
        let snames = id_names(sroot);
        let key_of = |e: &Elem, names: &HashMap<String, String>| {
            let id = e.attr("w:styleId").unwrap_or("");
            (e.attr("w:type").unwrap_or("paragraph").to_string(), names.get(id).cloned().unwrap_or_else(|| id.to_lowercase()))
        };
        let mut by_key: HashMap<(String, String), (String, String)> = HashMap::new();
        let mut base_ids: HashSet<String> = HashSet::new();
        let mut default_types: HashSet<String> = HashSet::new();
        // Стили «по умолчанию» базы: тип -> (id, подпись)
        let mut base_defaults: HashMap<String, (String, String)> = HashMap::new();
        for e in broot.elems().filter(|e| e.name == "w:style") {
            let id = e.attr("w:styleId").unwrap_or("").to_string();
            let entry = (id.clone(), style_sig(e, &bnames));
            by_key.insert(key_of(e, &bnames), entry.clone());
            base_ids.insert(id.to_lowercase());
            if matches!(e.attr("w:default"), Some("1" | "true" | "on")) {
                let ty = e.attr("w:type").unwrap_or("paragraph").to_string();
                default_types.insert(ty.clone());
                base_defaults.insert(ty, entry);
            }
        }
        let mut conflicts: Vec<String> = vec![];
        let mut to_import: Vec<Elem> = vec![];
        let mut src_default_para: Option<String> = None;
        for st in sroot.elems().filter(|e| e.name == "w:style") {
            let Some(id) = st.attr("w:styleId") else { continue };
            if st.attr("w:type").unwrap_or("paragraph") == "paragraph" && matches!(st.attr("w:default"), Some("1" | "true" | "on")) {
                src_default_para = Some(id.to_string());
            }
            let key = key_of(st, &snames);
            let is_default = matches!(st.attr("w:default"), Some("1" | "true" | "on"));
            let base_default = if is_default { base_defaults.get(&key.0) } else { None };
            if let Some((bid, bsig)) = by_key.get(&key).or(base_default) {
                cx.style_map.insert(id.to_string(), bid.clone());
                if *bsig != style_sig(st, &snames) {
                    conflicts.push(st.child("w:name").and_then(|n| n.val()).unwrap_or(id).to_string());
                }
            } else {
                let mut new_id = id.to_string();
                if base_ids.contains(&new_id.to_lowercase()) {
                    let mut k = 1;
                    loop {
                        new_id = format!("{id}_m{k}");
                        if !base_ids.contains(&new_id.to_lowercase()) {
                            break;
                        }
                        k += 1;
                    }
                }
                base_ids.insert(new_id.to_lowercase());
                cx.style_map.insert(id.to_string(), new_id.clone());
                let mut c = st.clone();
                c.set_attr("w:styleId", new_id);
                by_key.insert(key, (c.attr("w:styleId").unwrap_or("").to_string(), String::new()));
                to_import.push(c);
            }
        }
        let tmode = Mode { rels: None };
        let default_inject_saved = cx.inject_pstyle.take();
        let mut imported = vec![];
        for mut c in to_import {
            let ty = c.attr("w:type").unwrap_or("paragraph").to_string();
            if default_types.contains(&ty) {
                c.remove_attr("w:default");
            } else if matches!(c.attr("w:default"), Some("1" | "true" | "on")) {
                default_types.insert(ty);
            }
            imported.push(c);
        }
        // Преобразуем ссылки внутри импортируемых стилей (их тела не зависят от структуры базы).
        for c in imported.iter_mut() {
            // basedOn/next/link, ссылки на нумерацию и стили внутри определения
            xml::walk(c, &mut |e| self.xform_elem(src, cx, &tmode, e));
        }
        cx.inject_pstyle = default_inject_saved;
        let part = self.parts.get_mut("styles").expect("часть стилей загружена");
        if !imported.is_empty() {
            for c in imported {
                part.doc.root.push(c);
            }
            part.dirty = true;
        }
        let sdr = sroot.child("w:docDefaults").map(|d| {
            let mut s = String::new();
            canon_xml(d, &[], &mut s);
            s
        });
        let bdr = part.doc.root.child("w:docDefaults").map(|d| {
            let mut s = String::new();
            canon_xml(d, &[], &mut s);
            s
        });
        merge_root_decls(&mut part.doc.root, sroot)?;
        part.dirty = true;
        // Стиль абзаца «по умолчанию» исходного документа.
        let base_default = part
            .doc
            .root
            .elems()
            .find(|e| {
                e.name == "w:style" && e.attr("w:type").unwrap_or("paragraph") == "paragraph" && matches!(e.attr("w:default"), Some("1" | "true" | "on"))
            })
            .and_then(|e| e.attr("w:styleId"))
            .map(String::from);
        if let (Some(sd), Some(bd)) = (&src_default_para, &base_default) {
            if let Some(mapped) = cx.style_map.get(sd) {
                if mapped != bd {
                    cx.inject_pstyle = Some(mapped.clone());
                }
            }
        }
        let f = cx.file.clone();
        if !conflicts.is_empty() {
            conflicts.sort();
            conflicts.dedup();
            let shown: Vec<&str> = conflicts.iter().take(6).map(|s| s.as_str()).collect();
            let more = if conflicts.len() > 6 { format!(" и ещё {}", conflicts.len() - 6) } else { String::new() };
            self.warn(format!(
                "«{f}»: стили с теми же именами, но другим оформлением ({}{more}) заменены стилями первого документа — внешний вид текста может отличаться.",
                shown.join(", ")
            ));
        }
        if let (Some(a), Some(b)) = (sdr, bdr) {
            if a != b {
                self.warn(format!(
                    "«{f}»: шрифт, размер и интервалы «по умолчанию» отличаются от первого документа — текст без явного оформления может выглядеть иначе."
                ));
            }
        }
        Ok(())
    }

    // ---------- нумерация ----------

    fn numbering_offsets(&mut self, cx: &mut Ctx) -> Result<()> {
        let (mut abs, mut num, mut pic) = (-1i64, 0i64, -1i64);
        if self.load_part("numbering")? {
            for e in self.parts["numbering"].doc.root.elems() {
                let get = |k: &str| e.attr(k).and_then(|v| v.trim().parse::<i64>().ok());
                match e.name.as_str() {
                    "w:abstractNum" => abs = abs.max(get("w:abstractNumId").unwrap_or(-1)),
                    "w:num" => num = num.max(get("w:numId").unwrap_or(0)),
                    "w:numPicBullet" => pic = pic.max(get("w:numPicBulletId").unwrap_or(-1)),
                    _ => {}
                }
            }
        }
        cx.abs_off = abs.saturating_add(1) as u64;
        cx.num_off = num as u64;
        cx.pic_off = pic.saturating_add(1) as u64;
        Ok(())
    }

    fn merge_numbering(&mut self, src: &Src, cx: &mut Ctx) -> Result<()> {
        let Some(npath) = src.part_path("numbering") else { return Ok(()) };
        let ndoc = src.load_xml(&npath)?;
        if !self.load_part("numbering")? {
            let mut root = ndoc.root.clone();
            root.children.clear();
            self.create_part("numbering", root);
        }
        let dest = self.parts["numbering"].path.clone();
        let mode = Mode { rels: Some((npath, dest)) };
        let (mut pics, mut abss, mut nums) = (vec![], vec![], vec![]);
        for e in ndoc.root.elems() {
            let mut c = e.clone();
            let shift = |c: &mut Elem, k: &str, off: u64| {
                if let Some(v) = c.attr(k).and_then(|v| v.trim().parse::<u64>().ok()) {
                    c.set_attr(k, v.saturating_add(off).to_string());
                }
            };
            match e.name.as_str() {
                "w:numPicBullet" | "w:abstractNum" | "w:num" => {}
                _ => continue,
            }
            cx.ranges.clear();
            xml::walk(&mut c, &mut |x| self.xform_elem(src, cx, &mode, x));
            xml::walk(&mut c, &mut |x| {
                if x.name == "w:lvlPicBulletId" {
                    if let Some(v) = x.val().and_then(|v| v.trim().parse::<u64>().ok()) {
                        x.set_attr("w:val", v.saturating_add(cx.pic_off).to_string());
                    }
                }
                true
            });
            match e.name.as_str() {
                "w:numPicBullet" => {
                    shift(&mut c, "w:numPicBulletId", cx.pic_off);
                    pics.push(c);
                }
                "w:abstractNum" => {
                    shift(&mut c, "w:abstractNumId", cx.abs_off);
                    if let Some(ns) = c.child_mut("w:nsid") {
                        if let Some(v) = ns.val().map(String::from) {
                            let n = self.ids.nsid.map(&v);
                            ns.set_attr("w:val", n);
                        }
                    }
                    abss.push(c);
                }
                _ => {
                    shift(&mut c, "w:numId", cx.num_off);
                    if let Some(a) = c.child_mut("w:abstractNumId") {
                        if let Some(v) = a.val().and_then(|v| v.trim().parse::<u64>().ok()) {
                            a.set_attr("w:val", v.saturating_add(cx.abs_off).to_string());
                        }
                    }
                    nums.push(c);
                }
            }
        }
        let part = self.parts.get_mut("numbering").expect("нумерация загружена");
        merge_root_decls(&mut part.doc.root, &ndoc.root)?;
        let old = std::mem::take(&mut part.doc.root.children);
        let (mut bp, mut ba, mut bn, mut rest) = (vec![], vec![], vec![], vec![]);
        for n in old {
            if let Node::Elem(e) = n {
                match e.name.as_str() {
                    "w:numPicBullet" => bp.push(e),
                    "w:abstractNum" => ba.push(e),
                    "w:num" => bn.push(e),
                    _ => rest.push(e),
                }
            }
        }
        for e in bp.into_iter().chain(pics).chain(ba).chain(abss).chain(bn).chain(nums).chain(rest) {
            part.doc.root.push(e);
        }
        part.dirty = true;
        Ok(())
    }

    // ---------- сноски, примечания ----------

    fn plan_notes(&mut self, src: &Src, cx: &mut Ctx) -> Result<Vec<(&'static str, String, xml::Doc)>> {
        let mut v = vec![];
        for (key, item) in [("footnotes", "w:footnote"), ("endnotes", "w:endnote"), ("comments", "w:comment")] {
            let Some(path) = src.part_path(key) else { continue };
            let doc = src.load_xml(&path)?;
            let mut max: i64 = if key == "comments" { -1 } else { 0 };
            if self.load_part(key)? {
                for e in self.parts[key].doc.root.elems().filter(|e| e.name == item) {
                    if let Some(i) = e.attr("w:id").and_then(|v| v.trim().parse::<i64>().ok()) {
                        max = max.max(i);
                    }
                }
            }
            let mut map = HashMap::new();
            for e in doc.root.elems().filter(|e| e.name == item) {
                let normal = key == "comments" || matches!(e.attr("w:type"), None | Some("normal"));
                if let (true, Some(old)) = (normal, e.attr("w:id")) {
                    max = max.saturating_add(1);
                    map.insert(old.to_string(), max.to_string());
                }
            }
            cx.note_maps.insert(key, map);
            v.push((key, path, doc));
        }
        Ok(v)
    }

    fn ensure_note_settings(&mut self, key: &'static str) -> Result<()> {
        if !self.load_part("settings")? {
            let root = Elem::new("w:settings").with("xmlns:w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
            self.create_part("settings", root);
        }
        let (tag, item) = if key == "footnotes" { ("w:footnotePr", "w:footnote") } else { ("w:endnotePr", "w:endnote") };
        let part = self.parts.get_mut("settings").expect("настройки загружены");
        let root = &mut part.doc.root;
        if root.child(tag).is_some() {
            return Ok(());
        }
        let pr = Elem::new(tag)
            .with_child(Elem::new(item).with("w:id", "-1"))
            .with_child(Elem::new(item).with("w:id", "0"));
        const AFTER: &[&str] = &[
            "w:endnotePr", "w:compat", "w:docVars", "w:rsids", "m:mathPr", "w:attachedSchema", "w:themeFontLang",
            "w:clrSchemeMapping", "w:doNotIncludeSubdocsInStats", "w:doNotAutoCompressPictures", "w:forceUpgrade",
            "w:captions", "w:readModeInkLockDown", "w:smartTagType", "w:schemaLibrary", "w:shapeDefaults",
            "w:doNotEmbedSmartTags", "w:decimalSymbol", "w:listSeparator",
        ];
        let pos = root
            .children
            .iter()
            .position(|n| matches!(n, Node::Elem(e) if AFTER.contains(&e.name.as_str())))
            .unwrap_or(root.children.len());
        root.children.insert(pos, Node::Elem(pr));
        part.dirty = true;
        Ok(())
    }

    fn import_notes(&mut self, src: &Src, cx: &mut Ctx, notes: Vec<(&'static str, String, xml::Doc)>) -> Result<()> {
        for (key, spath, sdoc) in notes {
            let item = match key {
                "footnotes" => "w:footnote",
                "endnotes" => "w:endnote",
                _ => "w:comment",
            };
            if !self.load_part(key)? {
                let mut root = sdoc.root.clone();
                root.children.clear();
                if key != "comments" {
                    let sep = |ty: &str, id: &str, inner: &str| {
                        Elem::new(item).with("w:type", ty).with("w:id", id).with_child(
                            Elem::new("w:p")
                                .with_child(Elem::new("w:pPr").with_child(
                                    Elem::new("w:spacing").with("w:after", "0").with("w:line", "240").with("w:lineRule", "auto"),
                                ))
                                .with_child(Elem::new("w:r").with_child(Elem::new(inner))),
                        )
                    };
                    root.push(sep("separator", "-1", "w:separator"));
                    root.push(sep("continuationSeparator", "0", "w:continuationSeparator"));
                }
                self.create_part(key, root);
                if key != "comments" {
                    self.ensure_note_settings(key)?;
                }
            }
            let dest = self.parts[key].path.clone();
            let mode = Mode { rels: Some((spath, dest)) };
            let mut items = vec![];
            for e in sdoc.root.elems().filter(|e| e.name == item) {
                let Some(new) = e.attr("w:id").and_then(|o| cx.note_maps.get(key).and_then(|m| m.get(o))).cloned() else { continue };
                let mut c = e.clone();
                c.set_attr("w:id", new);
                cx.ranges.clear();
                xml::walk(&mut c, &mut |x| self.xform_elem(src, cx, &mode, x));
                items.push(c);
            }
            let part = self.parts.get_mut(key).expect("часть загружена");
            merge_root_decls(&mut part.doc.root, &sdoc.root)?;
            for c in items {
                part.doc.root.push(c);
            }
            part.dirty = true;
        }
        Ok(())
    }

    /// commentsExtended / commentsIds / commentsExtensible / people.
    fn import_comment_parts(&mut self, src: &Src, cx: &mut Ctx) -> Result<()> {
        let has_comments = cx.note_maps.get("comments").is_some_and(|m| !m.is_empty());
        for (key, item) in [
            ("commentsIds", "w16cid:commentId"),
            ("commentsExtended", "w15:commentEx"),
            ("commentsExtensible", "w16cex:commentExtensible"),
            ("people", "w15:person"),
        ] {
            if !has_comments && key != "people" {
                continue;
            }
            let Some(path) = src.part_path(key) else { continue };
            let doc = src.load_xml(&path)?;
            if !self.load_part(key)? {
                let mut root = doc.root.clone();
                root.children.clear();
                self.create_part(key, root);
            }
            let existing_people: HashSet<String> = if key == "people" {
                self.parts[key].doc.root.elems().filter_map(|p| p.attr("w15:author").map(String::from)).collect()
            } else {
                HashSet::new()
            };
            let mut items = vec![];
            for e in doc.root.elems().filter(|e| e.name == item) {
                let mut c = e.clone();
                match key {
                    "people" => {
                        if c.attr("w15:author").is_some_and(|a| existing_people.contains(a)) {
                            continue;
                        }
                    }
                    "commentsIds" => {
                        let Some(p) = c.attr("w16cid:paraId").and_then(|p| cx.para_map.get(p)).cloned() else { continue };
                        c.set_attr("w16cid:paraId", p);
                        if let Some(d) = c.attr("w16cid:durableId").map(String::from) {
                            let n = self.ids.durable.map(&d);
                            cx.durable_map.insert(d, n.clone());
                            c.set_attr("w16cid:durableId", n);
                        }
                    }
                    "commentsExtended" => {
                        let Some(p) = c.attr("w15:paraId").and_then(|p| cx.para_map.get(p)).cloned() else { continue };
                        c.set_attr("w15:paraId", p);
                        if let Some(pp) = c.attr("w15:paraIdParent").map(String::from) {
                            match cx.para_map.get(&pp) {
                                Some(n) => c.set_attr("w15:paraIdParent", n.clone()),
                                None => {
                                    c.remove_attr("w15:paraIdParent");
                                }
                            }
                        }
                    }
                    _ => {
                        let Some(d) = c.attr("w16cex:durableId").map(String::from) else { continue };
                        let n = match cx.durable_map.get(&d) {
                            Some(n) => n.clone(),
                            None => self.ids.durable.map(&d),
                        };
                        c.set_attr("w16cex:durableId", n);
                    }
                }
                items.push(c);
            }
            let part = self.parts.get_mut(key).expect("часть загружена");
            merge_root_decls(&mut part.doc.root, &doc.root)?;
            for c in items {
                part.doc.root.push(c);
            }
            part.dirty = true;
        }
        Ok(())
    }

    // ---------- разделы и колонтитулы ----------

    fn empty_story(&mut self, cx: &mut Ctx, footer: bool) -> String {
        let i = footer as usize;
        if let Some(r) = &cx.empty_story[i] {
            return r.clone();
        }
        let (root_name, file, ty, ct) = if footer {
            ("w:ftr", "footer.xml", "footer", "application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml")
        } else {
            ("w:hdr", "header.xml", "header", "application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml")
        };
        let root = Elem::new(root_name)
            .with("xmlns:w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main")
            .with("xmlns:r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships")
            .with_child(
                Elem::new("w:p").with_child(
                    Elem::new("w:pPr")
                        .with_child(Elem::new("w:spacing").with("w:before", "0").with("w:after", "0").with("w:line", "20").with("w:lineRule", "exact"))
                        .with_child(Elem::new("w:rPr").with_child(Elem::new("w:sz").with("w:val", "2"))),
                ),
            );
        let path = self.unique_name(&join(&self.mdir(), file));
        self.types.ensure(&path, ct, false);
        self.pkg.set(&path, xml::to_bytes(&xml::Doc { prolog: vec![], root }));
        let target = pct_encode_path(&relative(&self.mdir(), &path));
        let main = self.main.clone();
        let rid = self.rels_mut(&main).add(&format!("http://schemas.openxmlformats.org/officeDocument/2006/relationships/{ty}"), &target, false);
        cx.empty_story[i] = Some(rid.clone());
        rid
    }

    /// Первый раздел следующего документа не должен унаследовать колонтитулы предыдущего:
    /// недостающие ссылки заменяются пустыми колонтитулами. Затем учитываются определённые ссылки.
    fn section_step(&mut self, cx: &mut Ctx, s: &mut Elem, first: &mut bool, adjust: bool) {
        if *first && adjust {
            let keep = matches!(s.child("w:type").and_then(|t| t.val()), Some("oddPage" | "evenPage"));
            if !keep {
                s.children.retain(|n| !matches!(n, Node::Elem(e) if e.name == "w:type"));
            }
            let defined: HashSet<(bool, String)> = sect_refs(s).into_iter().collect();
            let mut add = vec![];
            for footer in [false, true] {
                for t in HF_TYPES {
                    let key = (footer, t.to_string());
                    if self.hf_state.contains(&key) && !defined.contains(&key) {
                        let rid = self.empty_story(cx, footer);
                        add.push(
                            Elem::new(if footer { "w:footerReference" } else { "w:headerReference" })
                                .with("w:type", t)
                                .with("r:id", &rid),
                        );
                    }
                }
            }
            for (i, a) in add.into_iter().enumerate() {
                s.children.insert(i, Node::Elem(a));
            }
        }
        *first = false;
        for r in sect_refs(s) {
            self.hf_state.insert(r);
        }
    }

    // ---------- проверки совместимости ----------

    fn compat_checks(&mut self, src: &Src, cx: &Ctx) -> Result<()> {
        let f = cx.file.clone();
        // Настройки документа, которые действуют на весь файл.
        let flags = [
            ("w:evenAndOddHeaders", "разные колонтитулы для чётных и нечётных страниц"),
            ("w:mirrorMargins", "зеркальные поля"),
            ("w:gutterAtTop", "поле переплёта сверху"),
        ];
        let base_has = |this: &Self, tag: &str| this.parts.get("settings").is_some_and(|p| p.doc.root.child(tag).is_some());
        self.load_part("settings")?;
        if let Some(sp) = src.part_path("settings") {
            let sdoc = src.load_xml(&sp)?;
            for (tag, what) in flags {
                if sdoc.root.child(tag).is_some() != base_has(self, tag) {
                    self.warn(format!(
                        "«{f}»: параметр «{what}» в этом документе отличается от первого документа; используется настройка первого документа."
                    ));
                }
            }
        } else {
            for (tag, what) in flags {
                if base_has(self, tag) {
                    self.warn(format!(
                        "«{f}»: параметр «{what}» включён в первом документе и будет применён и к этому документу."
                    ));
                }
            }
        }
        // Встроенные (внедрённые) шрифты не переносятся.
        let embedded = src
            .rels
            .items
            .iter()
            .filter(|r| !r.external && rel_suffix(&r.ty) == "fontTable")
            .filter_map(|r| resolve_part(&src.pkg, &src.main, &r.target))
            .any(|p| src.pkg.get(&rels_path(&p)).is_some_and(|b| String::from_utf8_lossy(b).contains("/font\"")));
        if embedded {
            self.warn(format!(
                "«{f}»: встроенные в документ шрифты не переносятся; если этих шрифтов нет на компьютере, они будут заменены."
            ));
        }
        // Тема оформления (цвета и шрифты темы).
        let uses_theme = |bytes: Option<&[u8]>| {
            bytes.is_some_and(|b| {
                let s = String::from_utf8_lossy(b);
                ["Theme=\"", "themeColor=\"", "themeFill=\"", "themeShade=\"", "themeTint=\""].iter().any(|k| s.contains(k))
            })
        };
        let styles_bytes = src.part_path("styles").and_then(|p| src.pkg.get(&p));
        if uses_theme(src.pkg.get(&src.main)) || uses_theme(styles_bytes) {
            let theme_sig = |pkg: &Pkg, owner_rels: &Rels, owner: &str| -> Option<String> {
                let p = owner_rels
                    .items
                    .iter()
                    .filter(|r| !r.external && rel_suffix(&r.ty) == "theme")
                    .find_map(|r| resolve_part(pkg, owner, &r.target))?;
                let d = xml::parse(pkg.get(&p)?).ok()?;
                let te = d.root.child("a:themeElements")?;
                let mut s = String::new();
                for c in te.elems().filter(|c| c.name == "a:clrScheme" || c.name == "a:fontScheme") {
                    canon_xml(c, &[], &mut s);
                }
                Some(s)
            };
            let main = self.main.clone();
            let brels = self.rels_mut(&main).clone();
            let a = theme_sig(&src.pkg, &src.rels, &src.main);
            let b = theme_sig(&self.pkg, &brels, &main);
            if let (Some(a), Some(b)) = (a, b) {
                if a != b {
                    self.warn(format!(
                        "«{f}»: тема оформления (цвета и шрифты темы) отличается от первого документа; текст, оформленный цветами и шрифтами темы, может выглядеть иначе."
                    ));
                }
            }
        }
        Ok(())
    }

    // ---------- добавление документа ----------

    fn append(&mut self, mut src: Src) -> Result<()> {
        if !self.touched {
            // Первый документ при слиянии трогаем минимально: только порядок sectPr в pPr.
            let mut tmp = Elem::new("w:body");
            tmp.children = std::mem::take(&mut self.body);
            tmp.push(std::mem::replace(&mut self.final_sect, Elem::new("w:sectPr")));
            xml::walk(&mut tmp, &mut |e| {
                if e.name == "w:pPr" {
                    fix_ppr_order(e);
                }
                true
            });
            if let Some(Node::Elem(f)) = tmp.children.pop() {
                self.final_sect = f;
            }
            self.body = tmp.children;
        }
        self.touched = true;
        let mut cx = Ctx { file: src.name.clone(), ..Default::default() };
        merge_root_decls(&mut self.doc.root, &src.doc.root)?;
        self.compat_checks(&src, &cx)?;
        self.numbering_offsets(&mut cx)?;
        self.merge_styles(&src, &mut cx)?;
        self.merge_numbering(&src, &mut cx)?;
        let notes = self.plan_notes(&src, &mut cx)?;

        // Тело документа.
        let body = src.doc.root.child_mut("w:body").ok_or_else(|| anyhow!("нет <w:body>"))?;
        let mut kids = std::mem::take(&mut body.children);
        kids.retain(|n| !is_ws_text(n));
        let fin = kids.iter().rposition(|n| matches!(n, Node::Elem(e) if e.name == "w:sectPr"));
        let fin = fin.map(|i| kids.remove(i)).and_then(|n| if let Node::Elem(e) = n { Some(e) } else { None });
        let mut fin = fin.unwrap_or_else(default_sect);
        let mut tmp = Elem::new("w:body");
        tmp.children = kids;
        tmp.push(fin);

        let mut roots: Vec<&Elem> = vec![&tmp];
        for (_, _, d) in &notes {
            roots.push(&d.root);
        }
        self.plan_bookmarks(&mut cx, &roots);

        // Сноски и примечания переносятся до тела: они задают карту идентификаторов.
        self.import_notes(&src, &mut cx, notes)?;
        let mode = Mode { rels: Some((src.main.clone(), self.main.clone())) };
        xml::walk(&mut tmp, &mut |e| self.xform_elem(&src, &mut cx, &mode, e));
        self.import_comment_parts(&src, &mut cx)?;

        let Some(Node::Elem(f)) = tmp.children.pop() else { bail!("внутренняя ошибка: потерян раздел документа") };
        fin = f;
        let mut nodes = tmp.children;
        let mut first = true;
        for i in 0..nodes.len() {
            if let Some(s) = para_sect_mut(&mut nodes[i]) {
                self.section_step(&mut cx, s, &mut first, true);
            }
        }
        self.section_step(&mut cx, &mut fin, &mut first, true);

        // Предыдущий документ закрывается своим разделом, новый становится последним.
        let prev = std::mem::replace(&mut self.final_sect, fin);
        close_section(&mut self.body, prev);
        self.body.extend(nodes);

        let f = cx.file.clone();
        if cx.dropped_bindings > 0 {
            self.warn(format!(
                "«{f}»: привязка элементов управления к данным (XML-хранилище) не переносится — содержимое сохранено как обычный текст ({} шт.).",
                cx.dropped_bindings
            ));
        }
        if cx.dropped_refs > 0 {
            self.warn(format!(
                "«{f}»: найдены ссылки на сноски или примечания, которых нет в файле; они удалены ({} шт.).",
                cx.dropped_refs
            ));
        }
        Ok(())
    }

    fn finish(mut self, out: &Path) -> Result<Vec<String>> {
        if self.touched {
            let mut kids = std::mem::take(&mut self.body);
            kids.push(Node::Elem(std::mem::replace(&mut self.final_sect, Elem::new("w:sectPr"))));
            self.doc.root.child_mut("w:body").ok_or_else(|| anyhow!("нет <w:body>"))?.children = kids;
            let main = self.main.clone();
            self.pkg.set(&main, xml::to_bytes(&self.doc));
            self.types.set_override(&main, CT_DOC);
            for part in self.parts.values().filter(|p| p.dirty) {
                self.pkg.set(&part.path, xml::to_bytes(&part.doc));
            }
            for (owner, rels) in &self.rels {
                if rels.dirty {
                    self.pkg.set(&rels_path(owner), rels.to_bytes());
                }
            }
            self.pkg.set("[Content_Types].xml", self.types.to_bytes());
        }
        self.pkg.write(out)?;
        Ok(self.warnings)
    }
}

/// Объединяет документы `.docx` по порядку; возвращает предупреждения для пользователя.
/// Параметры объединения документов.
#[derive(Clone, Debug)]
pub struct Options {
    /// Сохранять оформление каждого документа (стили с одним именем, но разным видом не подменяются стилями первого).
    pub keep_source_styles: bool,
}

impl Default for Options {
    fn default() -> Self {
        Options { keep_source_styles: true }
    }
}

pub fn merge(inputs: &[PathBuf], out: &Path, _opts: &Options) -> Result<Vec<String>> {
    let (first, rest) = inputs.split_first().ok_or_else(|| anyhow!("нет файлов"))?;
    let base = Src::load(first)?;
    let mut m = Merger::new(base).with_context(|| format!("при разборе {}", first.display()))?;
    for p in rest {
        let src = Src::load(p)?;
        m.append(src).with_context(|| format!("при добавлении {}", p.display()))?;
    }
    m.finish(out)
}

#[cfg(test)]
mod tests {
    use super::*;

    const NS: &str = r#"xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture" xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" mc:Ignorable="w14""#;
    const RT: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/";
    const CT_NOTES: &str = "application/vnd.openxmlformats-officedocument.wordprocessingml.";

    /// Конструктор минимального docx из сырого XML.
    struct Fx {
        body: String,
        ns: String,
        root: String,
        parts: Vec<(String, Vec<u8>)>,
        rels: Vec<(String, String, String, bool)>,
        types: Vec<(String, String)>,
        styles: String,
        numbering: Option<String>,
    }

    fn styles_xml(extra: &str) -> String {
        format!(
            r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:docDefaults><w:rPrDefault><w:rPr><w:sz w:val="22"/></w:rPr></w:rPrDefault></w:docDefaults><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style><w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:basedOn w:val="Normal"/><w:rPr><w:b/></w:rPr></w:style>{extra}</w:styles>"#
        )
    }

    fn sect(extra: &str) -> String {
        format!(r#"<w:sectPr>{extra}<w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1000" w:right="1000" w:bottom="1000" w:left="1000" w:header="500" w:footer="500" w:gutter="0"/></w:sectPr>"#)
    }

    fn para(t: &str) -> String {
        format!(r#"<w:p><w:r><w:t>{t}</w:t></w:r></w:p>"#)
    }

    fn fx(body: &str) -> Fx {
        Fx {
            body: body.to_string(),
            ns: NS.to_string(),
            root: "w:document".into(),
            parts: vec![],
            rels: vec![],
            types: vec![],
            styles: styles_xml(""),
            numbering: None,
        }
    }

    impl Fx {
        fn part(mut self, name: &str, data: impl AsRef<[u8]>) -> Fx {
            self.parts.push((name.to_string(), data.as_ref().to_vec()));
            self
        }
        fn rel(mut self, id: &str, ty: &str, target: &str) -> Fx {
            self.rels.push((id.into(), format!("{RT}{ty}"), target.into(), false));
            self
        }
        fn ext(mut self, id: &str, ty: &str, target: &str) -> Fx {
            self.rels.push((id.into(), format!("{RT}{ty}"), target.into(), true));
            self
        }
        fn ct(mut self, part: &str, ct: &str) -> Fx {
            self.types.push((part.into(), ct.into()));
            self
        }
        fn write(&self, path: &Path) {
            let mut e: Vec<(String, Vec<u8>)> = vec![];
            let mut ct = String::from(
                r#"<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/>"#,
            );
            ct.push_str(&format!(r#"<Override PartName="/word/document.xml" ContentType="{CT_NOTES}document.main+xml"/><Override PartName="/word/styles.xml" ContentType="{CT_NOTES}styles+xml"/>"#));
            for (p, c) in &self.types {
                ct.push_str(&format!(r#"<Override PartName="{p}" ContentType="{c}"/>"#));
            }
            if self.numbering.is_some() {
                ct.push_str(&format!(r#"<Override PartName="/word/numbering.xml" ContentType="{CT_NOTES}numbering+xml"/>"#));
            }
            ct.push_str("</Types>");
            e.push(("[Content_Types].xml".into(), ct.into_bytes()));
            e.push((
                "_rels/.rels".into(),
                format!(r#"<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="{RT}officeDocument" Target="word/document.xml"/></Relationships>"#).into_bytes(),
            ));
            let mut rels = format!(r#"<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdSty" Type="{RT}styles" Target="styles.xml"/>"#);
            if self.numbering.is_some() {
                rels.push_str(&format!(r#"<Relationship Id="rIdNum" Type="{RT}numbering" Target="numbering.xml"/>"#));
            }
            for (id, ty, t, ext) in &self.rels {
                rels.push_str(&format!(r#"<Relationship Id="{id}" Type="{ty}" Target="{t}"{}/>"#, if *ext { r#" TargetMode="External""# } else { "" }));
            }
            rels.push_str("</Relationships>");
            e.push(("word/_rels/document.xml.rels".into(), rels.into_bytes()));
            e.push((
                "word/document.xml".into(),
                format!(r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><{r} {ns}><w:body>{b}</w:body></{r}>"#, r = self.root, ns = self.ns, b = self.body).into_bytes(),
            ));
            e.push(("word/styles.xml".into(), self.styles.clone().into_bytes()));
            if let Some(n) = &self.numbering {
                e.push(("word/numbering.xml".into(), n.clone().into_bytes()));
            }
            e.extend(self.parts.iter().cloned());
            Pkg::from_entries(e).write(path).unwrap();
        }
    }

    fn png(c: u8) -> Vec<u8> {
        vec![0x89, b'P', b'N', b'G', c]
    }

    fn img(rid: &str, id: u32) -> String {
        format!(
            r#"<w:r><w:drawing><wp:inline><wp:extent cx="100" cy="100"/><wp:docPr id="{id}" name="p"/><a:graphic><a:graphicData uri="x"><pic:pic><pic:blipFill><a:blip r:embed="{rid}"/></pic:blipFill></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>"#
        )
    }

    fn hdr(root: &str, t: &str) -> String {
        format!(r#"<?xml version="1.0"?><{root} {NS}><w:p><w:r><w:t>{t}</w:t></w:r></w:p></{root}>"#)
    }

    fn run(dir: &Path, name: &str, docs: &[&Fx]) -> (PathBuf, Vec<String>) {
        let mut paths = vec![];
        for (i, d) in docs.iter().enumerate() {
            let p = dir.join(format!("{name}_in{i}.docx"));
            d.write(&p);
            paths.push(p);
        }
        let out = dir.join(format!("{name}_out.docx"));
        let w = merge(&paths, &out, &Options::default()).unwrap();
        let errs = check(&out);
        assert!(errs.is_empty(), "{name}: пакет некорректен: {errs:#?}");
        (out, w)
    }

    fn doc_of(p: &Path) -> (Pkg, xml::Doc) {
        let pkg = Pkg::read(p).unwrap();
        let d = xml::parse(pkg.get("word/document.xml").unwrap()).unwrap();
        (pkg, d)
    }

    fn all(e: &Elem, name: &str) -> Vec<Elem> {
        let mut v = vec![];
        xml::walk_ref(e, &mut |x| {
            if x.name == name {
                v.push(x.clone());
            }
        });
        v
    }

    fn texts(e: &Elem) -> Vec<String> {
        all(e, "w:t").iter().map(|t| t.text()).collect()
    }

    /// Проверка целостности пакета (аналог проверок Word): связи, типы, уникальность id, порядок элементов.
    fn check(path: &Path) -> Vec<String> {
        let mut errs = vec![];
        let pkg = Pkg::read(path).unwrap();
        if pkg.entries[0].0 != "[Content_Types].xml" {
            errs.push("[Content_Types].xml не первый".into());
        }
        let types = Types::parse(pkg.get("[Content_Types].xml").unwrap()).unwrap();
        let mut docs: HashMap<String, xml::Doc> = HashMap::new();
        for (n, d) in &pkg.entries {
            if n.ends_with(".xml") || n.ends_with(".rels") {
                match xml::parse(d) {
                    Ok(x) => {
                        docs.insert(n.clone(), x);
                    }
                    Err(e) => errs.push(format!("{n}: {e}")),
                }
            }
            if n != "[Content_Types].xml" && types.ct_override(n).is_none() && types.ct_default(ext_of(n)).is_none() {
                errs.push(format!("нет типа содержимого: {n}"));
            }
        }
        for (_, ov) in types.overrides.iter().map(|(p, c)| (p, c)) {
            let _ = ov;
        }
        for (p, _) in &types.overrides {
            if pkg.get(p.trim_start_matches('/')).is_none() {
                errs.push(format!("Override на отсутствующую часть {p}"));
            }
        }
        let mut relmap: HashMap<String, Rels> = HashMap::new();
        for (n, _) in &pkg.entries {
            if let Some(owner) = n.strip_suffix(".rels") {
                let owner = owner.replace("_rels/", "");
                let owner = if owner == "." || owner.is_empty() { String::new() } else { owner };
                let r = Rels::parse(pkg.get(n).unwrap()).unwrap();
                let mut seen = HashSet::new();
                for x in &r.items {
                    if !seen.insert(x.id.clone()) {
                        errs.push(format!("{n}: повторный Id {}", x.id));
                    }
                    if !x.external && resolve_part(&pkg, &owner, &x.target).is_none() {
                        errs.push(format!("{n}: цель {} отсутствует", x.target));
                    }
                }
                relmap.insert(owner, r);
            }
        }
        for (n, d) in &docs {
            if n.ends_with(".rels") || n == "[Content_Types].xml" {
                continue;
            }
            xml::walk_ref(&d.root, &mut |e| {
                for (k, v) in &e.attrs {
                    if (k.starts_with("r:") || k == "o:relid") && relmap.get(n.as_str()).and_then(|r| r.get(v)).is_none() {
                        errs.push(format!("{n}: {k}={v} без связи"));
                    }
                }
            });
        }
        let main = docs.get("word/document.xml").unwrap();
        fn dups(what: &str, ids: Vec<String>) -> Vec<String> {
            let mut s = HashSet::new();
            ids.into_iter().filter(|i| !s.insert(i.clone())).map(|i| format!("{what}: повтор {i}")).collect()
        }
        let story_names: Vec<String> = pkg
            .entries
            .iter()
            .map(|e| e.0.clone())
            .filter(|n| n.starts_with("word/") && (n.contains("header") || n.contains("footer") || n.contains("footnotes") || n.contains("endnotes") || n.contains("comments")) && n.ends_with(".xml") && !n.contains("_rels"))
            .collect();
        let mut stories: Vec<&Elem> = vec![&main.root];
        for n in &story_names {
            if let Some(d) = docs.get(n) {
                if !n.contains("commentsExt") && !n.contains("commentsIds") {
                    stories.push(&d.root);
                }
            }
        }
        errs.extend(dups("docPr", stories.iter().flat_map(|r| all(r, "wp:docPr")).filter_map(|e| e.attr("id").map(String::from)).collect()));
        errs.extend(dups("bookmark id", stories.iter().flat_map(|r| all(r, "w:bookmarkStart")).filter_map(|e| e.attr("w:id").map(String::from)).collect()));
        errs.extend(dups("bookmark name", stories.iter().flat_map(|r| all(r, "w:bookmarkStart")).filter_map(|e| e.attr("w:name").map(String::from)).collect()));
        for (what, key, item, refname) in [("сноска", "footnotes", "w:footnote", "w:footnoteReference"), ("концевая", "endnotes", "w:endnote", "w:endnoteReference"), ("примечание", "comments", "w:comment", "w:commentReference")] {
            let part = docs.get(&format!("word/{key}.xml"));
            let ids: Vec<String> = part.map(|d| all(&d.root, item).iter().filter_map(|e| e.attr("w:id").map(String::from)).collect()).unwrap_or_default();
            errs.extend(dups(what, ids.clone()));
            for r in all(&main.root, refname) {
                let id = r.attr("w:id").unwrap_or("").to_string();
                if !ids.contains(&id) {
                    errs.push(format!("{refname} {id} без {what}"));
                }
            }
        }
        if let Some(n) = docs.get("word/numbering.xml") {
            let order: Vec<u8> = n
                .root
                .elems()
                .map(|e| match e.name.as_str() {
                    "w:numPicBullet" => 0,
                    "w:abstractNum" => 1,
                    "w:num" => 2,
                    _ => 3,
                })
                .collect();
            if order.windows(2).any(|w| w[0] > w[1]) {
                errs.push("numbering: нарушен порядок элементов".into());
            }
            errs.extend(dups("numId", all(&n.root, "w:num").iter().filter_map(|e| e.attr("w:numId").map(String::from)).collect()));
            errs.extend(dups("abstractNumId", all(&n.root, "w:abstractNum").iter().filter_map(|e| e.attr("w:abstractNumId").map(String::from)).collect()));
        }
        let body = main.root.child("w:body").unwrap();
        if body.elems().filter(|e| e.name == "w:sectPr").count() != 1 || body.elems().last().map(|e| e.name.as_str()) != Some("w:sectPr") {
            errs.push("в body должен быть ровно один sectPr, и он последний".into());
        }
        errs
    }

    #[test]
    fn xml_roundtrip_keeps_content() {
        let src = "<?xml version=\"1.0\"?><a:r xmlns:a=\"u\" x='1\"2' y=\"a&amp;b&lt;\"><!-- c --><a:t xml:space=\"preserve\"> a &amp; b </a:t><![CDATA[<raw>]]><a:e/></a:r>";
        let d = xml::parse(src.as_bytes()).unwrap();
        assert_eq!(d.root.attr("x"), Some("1\"2"));
        assert_eq!(d.root.attr("y"), Some("a&b<"));
        assert_eq!(d.root.child("a:t").unwrap().text(), " a & b ");
        let out = String::from_utf8(xml::to_bytes(&d)).unwrap();
        assert!(out.contains(r#"x="1&quot;2""#) && out.contains("<![CDATA[<raw>]]>") && out.contains("<!-- c -->"));
        assert_eq!(xml::parse(out.as_bytes()).unwrap().root.children.len(), d.root.children.len());
        assert!(xml::parse(b"<a><b></a>").is_err());
        assert!(xml::parse(b"<a>").is_err());
        assert!(xml::parse(b"<!DOCTYPE a><a/>").is_err());
    }

    #[test]
    fn each_document_keeps_its_own_section_and_headers() {
        let dir = tempfile::tempdir().unwrap();
        let a = fx(&format!(
            "{}{}",
            para("AAA"),
            sect(r#"<w:headerReference w:type="default" r:id="rIdH"/>"#)
        ))
        .part("word/header1.xml", hdr("w:hdr", "HEADER-A"))
        .rel("rIdH", "header", "header1.xml")
        .ct("/word/header1.xml", &format!("{CT_NOTES}header+xml"));
        let b = fx(&format!(
            "{}{}",
            para("BBB"),
            r#"<w:sectPr><w:pgSz w:w="16838" w:h="11906" w:orient="landscape"/></w:sectPr>"#
        ));
        let (out, w) = run(dir.path(), "sec", &[&a, &b, &a]);
        assert!(w.is_empty(), "{w:?}");
        let (pkg, d) = doc_of(&out);
        let body = d.root.child("w:body").unwrap();
        let sects = all(body, "w:sectPr");
        assert_eq!(sects.len(), 3);
        // нет искусственных абзацев с разрывом страницы
        assert!(!String::from_utf8_lossy(pkg.get("word/document.xml").unwrap()).contains(r#"w:type="page""#));
        // первый раздел: A4 портрет и колонтитул A
        assert_eq!(sects[0].child("w:pgSz").unwrap().attr("w:w"), Some("11906"));
        let h1 = sects[0].child("w:headerReference").unwrap().attr("r:id").unwrap().to_string();
        let rels = Rels::parse(pkg.get("word/_rels/document.xml.rels").unwrap()).unwrap();
        let t1 = resolve_part(&pkg, "word/document.xml", &rels.get(&h1).unwrap().target).unwrap();
        assert!(String::from_utf8_lossy(pkg.get(&t1).unwrap()).contains("HEADER-A"));
        // второй документ: альбомный, и он НЕ наследует колонтитул A (ссылка на пустой колонтитул)
        assert_eq!(sects[1].child("w:pgSz").unwrap().attr("w:orient"), Some("landscape"));
        let h2 = sects[1].child("w:headerReference").expect("пустой колонтитул").attr("r:id").unwrap().to_string();
        assert_ne!(h1, h2);
        let t2 = resolve_part(&pkg, "word/document.xml", &rels.get(&h2).unwrap().target).unwrap();
        assert!(!String::from_utf8_lossy(pkg.get(&t2).unwrap()).contains("HEADER-A"));
        // третий (копия первого) снова со своим колонтитулом, скопированным отдельной частью
        let h3 = sects[2].child("w:headerReference").unwrap().attr("r:id").unwrap().to_string();
        let t3 = resolve_part(&pkg, "word/document.xml", &rels.get(&h3).unwrap().target).unwrap();
        assert_ne!(t1, t3);
        assert!(String::from_utf8_lossy(pkg.get(&t3).unwrap()).contains("HEADER-A"));
    }

    #[test]
    fn section_goes_into_last_paragraph_without_extra_page_break() {
        let dir = tempfile::tempdir().unwrap();
        let a = fx(&format!("{}{}", para("ONE"), sect("")));
        let b = fx(&format!("{}{}", para("TWO"), sect("")));
        let (out, _) = run(dir.path(), "para", &[&a, &b]);
        let (_, d) = doc_of(&out);
        let body = d.root.child("w:body").unwrap();
        let ps: Vec<&Elem> = body.elems().filter(|e| e.name == "w:p").collect();
        assert_eq!(ps.len(), 2, "лишних абзацев быть не должно");
        assert!(ps[0].child("w:pPr").unwrap().child("w:sectPr").is_some());
        assert_eq!(texts(body), ["ONE", "TWO"]);
    }

    #[test]
    fn first_section_of_later_document_always_starts_on_new_page() {
        let dir = tempfile::tempdir().unwrap();
        let a = fx(&format!("{}{}", para("ONE"), sect("")));
        let b = fx(&format!("{}{}", para("TWO"), r#"<w:sectPr><w:type w:val="continuous"/><w:pgSz w:w="1" w:h="1"/></w:sectPr>"#));
        let (out, _) = run(dir.path(), "cont", &[&a, &b]);
        let (_, d) = doc_of(&out);
        let last = d.root.child("w:body").unwrap().child("w:sectPr").unwrap().clone();
        assert!(last.child("w:type").is_none());
        assert_eq!(last.child("w:pgSz").unwrap().attr("w:w"), Some("1"));
    }

    #[test]
    fn images_hyperlinks_and_content_types_are_copied() {
        let dir = tempfile::tempdir().unwrap();
        let a = fx(&format!("{}{}", para("A"), sect("")));
        let b = fx(&format!(
            r#"<w:p>{}<w:hyperlink r:id="rIdL"><w:r><w:t>l</w:t></w:r></w:hyperlink></w:p>{}"#,
            img("rIdI", 1),
            sect("")
        ))
        .rel("rIdI", "image", "media/i.png")
        .ext("rIdL", "hyperlink", "https://example.com")
        .part("word/media/i.png", png(1));
        let (out, w) = run(dir.path(), "img", &[&a, &b]);
        assert!(w.is_empty());
        let (pkg, d) = doc_of(&out);
        let blip = all(&d.root, "a:blip").pop().unwrap();
        let rels = Rels::parse(pkg.get("word/_rels/document.xml.rels").unwrap()).unwrap();
        let r = rels.get(blip.attr("r:embed").unwrap()).unwrap();
        assert_eq!(pkg.get(&format!("word/{}", r.target)).unwrap(), png(1).as_slice());
        let link = all(&d.root, "w:hyperlink").pop().unwrap();
        let lr = rels.get(link.attr("r:id").unwrap()).unwrap();
        assert!(lr.external && lr.target == "https://example.com");
    }

    #[test]
    fn remerging_an_output_never_overwrites_parts_or_ids() {
        let dir = tempfile::tempdir().unwrap();
        let mk = |c: u8| {
            let mut f = fx(&format!("<w:p>{}</w:p>{}", img("rIdI", 1), sect(""))).rel("rIdI", "image", "media/image1.png").part("word/media/image1.png", png(c));
            f.numbering = None;
            f
        };
        let (a, b, c) = (mk(1), mk(2), mk(3));
        let (ab, _) = run(dir.path(), "ab", &[&a, &b]);
        let c_path = dir.path().join("c.docx");
        c.write(&c_path);
        let out = dir.path().join("abc.docx");
        merge(&[ab.clone(), c_path.clone()], &out, &Options::default()).unwrap();
        assert!(check(&out).is_empty(), "{:?}", check(&out));
        let out2 = dir.path().join("cab.docx");
        merge(&[c_path, ab], &out2, &Options::default()).unwrap();
        assert!(check(&out2).is_empty());
        for (o, order) in [(&out, [1u8, 2, 3]), (&out2, [3, 1, 2])] {
            let (pkg, d) = doc_of(o);
            let rels = Rels::parse(pkg.get("word/_rels/document.xml.rels").unwrap()).unwrap();
            let seen: Vec<u8> = all(&d.root, "a:blip")
                .iter()
                .map(|b| {
                    let t = &rels.get(b.attr("r:embed").unwrap()).unwrap().target;
                    pkg.get(&format!("word/{t}")).unwrap()[4]
                })
                .collect();
            assert_eq!(seen, order);
        }
    }

    #[test]
    fn notes_and_comments_are_renumbered_and_created() {
        let dir = tempfile::tempdir().unwrap();
        let notes = |t: &str| {
            format!(
                r#"<?xml version="1.0"?><w:footnotes {NS}><w:footnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:footnote><w:footnote w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote><w:footnote w:id="1"><w:p><w:r><w:footnoteRef/></w:r><w:r><w:t>FN-{t}</w:t></w:r></w:p></w:footnote></w:footnotes>"#
            )
        };
        let end = |t: &str| notes(t).replace("footnote", "endnote").replace("footnotes", "endnotes");
        let cm = |t: &str| {
            format!(r#"<?xml version="1.0"?><w:comments {NS}><w:comment w:id="0" w:author="X" w:date="2024-01-01T00:00:00Z"><w:p><w:r><w:t>CM-{t}</w:t></w:r></w:p></w:comment></w:comments>"#)
        };
        let with_notes = |t: &str| {
            fx(&format!(
                r#"<w:p><w:r><w:t>{t}</w:t></w:r><w:r><w:footnoteReference w:id="1"/></w:r><w:r><w:endnoteReference w:id="1"/></w:r><w:commentRangeStart w:id="0"/><w:r><w:t>x</w:t></w:r><w:commentRangeEnd w:id="0"/><w:r><w:commentReference w:id="0"/></w:r></w:p>{}"#,
                sect("")
            ))
            .part("word/footnotes.xml", notes(t))
            .part("word/endnotes.xml", end(t))
            .part("word/comments.xml", cm(t))
            .rel("rF", "footnotes", "footnotes.xml")
            .rel("rE", "endnotes", "endnotes.xml")
            .rel("rC", "comments", "comments.xml")
            .ct("/word/footnotes.xml", &format!("{CT_NOTES}footnotes+xml"))
            .ct("/word/endnotes.xml", &format!("{CT_NOTES}endnotes+xml"))
            .ct("/word/comments.xml", &format!("{CT_NOTES}comments+xml"))
        };
        let plain = fx(&format!("{}{}", para("PLAIN"), sect("")));
        let (a, b) = (with_notes("ONE"), with_notes("TWO"));
        // база без сносок и примечаний — части создаются
        let (out, w) = run(dir.path(), "notes_new", &[&plain, &b]);
        assert!(w.is_empty(), "{w:?}");
        let (pkg, _) = doc_of(&out);
        assert!(pkg.get("word/footnotes.xml").is_some() && pkg.get("word/endnotes.xml").is_some() && pkg.get("word/comments.xml").is_some());
        let fnx = xml::parse(pkg.get("word/footnotes.xml").unwrap()).unwrap();
        assert!(all(&fnx.root, "w:footnote").iter().any(|f| f.attr("w:type") == Some("separator")));
        // обе стороны со сносками — id не пересекаются, текст у своего документа
        let (out, _) = run(dir.path(), "notes_both", &[&a, &b]);
        let (pkg, d) = doc_of(&out);
        let fnx = xml::parse(pkg.get("word/footnotes.xml").unwrap()).unwrap();
        let find = |id: &str, root: &Elem, item: &str| texts(all(root, item).iter().find(|f| f.attr("w:id") == Some(id)).unwrap()).concat();
        let refs: Vec<String> = all(&d.root, "w:footnoteReference").iter().map(|r| r.attr("w:id").unwrap().to_string()).collect();
        assert_eq!(refs.len(), 2);
        assert_eq!(find(&refs[0], &fnx.root, "w:footnote"), "FN-ONE");
        assert_eq!(find(&refs[1], &fnx.root, "w:footnote"), "FN-TWO");
        let cx = xml::parse(pkg.get("word/comments.xml").unwrap()).unwrap();
        let crefs: Vec<String> = all(&d.root, "w:commentReference").iter().map(|r| r.attr("w:id").unwrap().to_string()).collect();
        assert_eq!(find(&crefs[0], &cx.root, "w:comment"), "CM-ONE");
        assert_eq!(find(&crefs[1], &cx.root, "w:comment"), "CM-TWO");
        let starts: Vec<String> = all(&d.root, "w:commentRangeStart").iter().map(|r| r.attr("w:id").unwrap().to_string()).collect();
        assert_eq!(starts, crefs);
        // сноски, у которых нет части в файле: ссылки удаляются с предупреждением, файл не портится
        let broken = fx(&format!(r#"<w:p><w:r><w:footnoteReference w:id="7"/></w:r></w:p>{}"#, sect("")));
        let (out, w) = run(dir.path(), "notes_broken", &[&plain, &broken]);
        assert!(!w.is_empty());
        assert!(all(&doc_of(&out).1.root, "w:footnoteReference").is_empty());
    }

    #[test]
    fn charts_ole_and_other_parts_are_copied_recursively() {
        let dir = tempfile::tempdir().unwrap();
        let a = fx(&format!("{}{}", para("A"), sect("")));
        let b = fx(&format!(
            r#"<w:p><w:r><w:drawing><wp:inline><wp:docPr id="3" name="c"/><a:graphic><a:graphicData uri="c"><c:chart r:id="rId50"/></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>{}"#,
            sect("")
        ))
        .rel("rId50", "chart", "charts/chart1.xml")
        .part("word/charts/chart1.xml", format!(r#"<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><c:externalData r:id="rId1"/></c:chartSpace>"#))
        .part("word/charts/_rels/chart1.xml.rels", format!(r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="{RT}package" Target="../embeddings/Sheet1.xlsx"/></Relationships>"#))
        .part("word/embeddings/Sheet1.xlsx", b"XLSX")
        .ct("/word/charts/chart1.xml", "application/vnd.openxmlformats-officedocument.drawingml.chart+xml")
        .ct("/word/embeddings/Sheet1.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
        // два раза: имена частей не должны пересекаться
        let (out, w) = run(dir.path(), "chart", &[&a, &b, &b]);
        assert!(w.is_empty(), "{w:?}");
        let (pkg, _) = doc_of(&out);
        assert!(pkg.get("word/charts/chart1.xml").is_some() && pkg.get("word/charts/chart2.xml").is_some());
        assert!(pkg.get("word/embeddings/Sheet1.xlsx").is_some() && pkg.get("word/embeddings/Sheet2.xlsx").is_some());
        let r2 = String::from_utf8_lossy(pkg.get("word/charts/_rels/chart2.xml.rels").unwrap()).into_owned();
        assert!(r2.contains("Sheet2.xlsx"));
    }

    #[test]
    fn ole_vml_and_smartart_references_are_remapped() {
        let dir = tempfile::tempdir().unwrap();
        let ole = |n: u8| {
            fx(&format!(
                r#"<w:p><w:r><w:object><v:shape id="_x0000_i1025"><v:imagedata r:id="rIdI" o:title=""/></v:shape><o:OLEObject Type="Embed" ShapeID="_x0000_i1025" r:id="rIdO"/></w:object></w:r></w:p>{}"#,
                sect("")
            ))
            .rel("rIdI", "image", "media/o.png")
            .rel("rIdO", "oleObject", "embeddings/o.bin")
            .part("word/media/o.png", png(n))
            .part("word/embeddings/o.bin", vec![n])
            .ct("/word/embeddings/o.bin", "application/vnd.openxmlformats-officedocument.oleObject")
        };
        let dgm = fx(&format!(
            r#"<w:p><w:r><w:drawing><wp:inline><wp:docPr id="5" name="d"/><a:graphic><a:graphicData uri="d"><dgm:relIds xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram" r:dm="rD"/></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>{}"#,
            sect("")
        ))
        .rel("rD", "diagramData", "diagrams/data1.xml")
        .rel("rDr", "http://schemas.microsoft.com/office/2007/relationships/diagramDrawing".trim_start_matches(RT), "diagrams/drawing1.xml")
        .part("word/diagrams/data1.xml", r#"<dgm:dataModel xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram"><x relId="rDr"/></dgm:dataModel>"#)
        .part("word/diagrams/drawing1.xml", "<d/>")
        .ct("/word/diagrams/data1.xml", "application/vnd.openxmlformats-officedocument.drawingml.diagramData+xml")
        .ct("/word/diagrams/drawing1.xml", "application/vnd.ms-office.drawingml.diagramDrawing+xml");
        let (o1, o2) = (ole(1), ole(2));
        let (out, w) = run(dir.path(), "ole", &[&o1, &o2, &dgm, &dgm]);
        assert!(w.is_empty(), "{w:?}");
        let (pkg, d) = doc_of(&out);
        // идентификаторы VML-фигур различны, ShapeID каждого OLE-объекта указывает на свою фигуру
        let shapes: Vec<String> = all(&d.root, "v:shape").iter().map(|s| s.attr("id").unwrap().to_string()).collect();
        let ole_ids: Vec<String> = all(&d.root, "o:OLEObject").iter().map(|s| s.attr("ShapeID").unwrap().to_string()).collect();
        assert_eq!(shapes, ole_ids);
        assert_eq!(shapes.iter().collect::<HashSet<_>>().len(), 2);
        // relId внутри данных SmartArt указывает на новую связь с drawing
        let rels = Rels::parse(pkg.get("word/_rels/document.xml.rels").unwrap()).unwrap();
        let data2 = String::from_utf8_lossy(pkg.get("word/diagrams/data2.xml").unwrap()).into_owned();
        let rid = data2.split("relId=\"").nth(1).unwrap().split('"').next().unwrap();
        let t = &rels.get(rid).unwrap().target;
        assert!(t.ends_with("drawing2.xml"), "{t}");
    }

    #[test]
    fn sectprchange_is_not_taken_for_body_sectpr() {
        let dir = tempfile::tempdir().unwrap();
        let tracked = r#"<w:sectPr><w:pgSz w:w="1" w:h="2"/><w:sectPrChange w:id="9" w:author="a" w:date="2024-01-01T00:00:00Z"><w:sectPr><w:pgSz w:w="3" w:h="4"/></w:sectPr></w:sectPrChange></w:sectPr>"#;
        let a = fx(&format!("{}{tracked}", para("ONE")));
        let b = fx(&format!("{}{tracked}", para("TWO")));
        let plain = fx(&format!("{}{}", para("PLAIN"), sect("")));
        for (n, docs) in [("spc_base", [&a, &plain]), ("spc_other", [&plain, &b]), ("spc_both", [&a, &b])] {
            let (out, _) = run(dir.path(), n, &docs);
            let (_, d) = doc_of(&out);
            let body = d.root.child("w:body").unwrap();
            let last = body.elems().last().unwrap();
            assert_eq!(last.name, "w:sectPr");
            // вложенный sectPr остался внутри sectPrChange
            for s in all(body, "w:sectPrChange") {
                assert_eq!(s.elems().count(), 1);
            }
            assert_eq!(texts(body).concat().matches("PLAIN").count() + texts(body).concat().matches("ONE").count() + texts(body).concat().matches("TWO").count(), 2);
        }
    }

    #[test]
    fn drawing_bookmark_tracked_and_sdt_ids_are_unique() {
        let dir = tempfile::tempdir().unwrap();
        let mk = |t: &str| {
            fx(&format!(
                r#"<w:p><w:bookmarkStart w:id="0" w:name="_Toc1"/><w:r><w:t>{t}</w:t></w:r><w:bookmarkEnd w:id="0"/></w:p><w:p><w:hyperlink w:anchor="_Toc1"><w:r><w:t>go</w:t></w:r></w:hyperlink><w:r><w:instrText> PAGEREF _Toc1 \h </w:instrText></w:r><w:fldSimple w:instr=" REF _Toc1 "><w:r><w:t>r</w:t></w:r></w:fldSimple></w:p><w:p><w:ins w:id="5" w:author="x" w:date="2024-01-01T00:00:00Z"><w:r><w:t>i</w:t></w:r></w:ins></w:p><w:p>{}</w:p><w:sdt><w:sdtPr><w:id w:val="77"/><w:dataBinding w:xpath="/a"/></w:sdtPr><w:sdtContent><w:p><w:r><w:t>s</w:t></w:r></w:p></w:sdtContent></w:sdt>{}"#,
                img("rI", 1),
                sect("")
            ))
            .rel("rI", "image", "media/i.png")
            .part("word/media/i.png", png(1))
        };
        let (a, b) = (mk("A"), mk("B"));
        let (out, w) = run(dir.path(), "ids", &[&a, &b, &a]);
        assert_eq!(w.len(), 2, "{w:?}");
        assert!(w.iter().all(|m| m.contains("привязка")));
        let (_, d) = doc_of(&out);
        let names: Vec<String> = all(&d.root, "w:bookmarkStart").iter().map(|e| e.attr("w:name").unwrap().to_string()).collect();
        assert_eq!(names, ["_Toc1", "_Toc1_m1", "_Toc1_m2"]);
        let anchors: Vec<String> = all(&d.root, "w:hyperlink").iter().map(|e| e.attr("w:anchor").unwrap().to_string()).collect();
        assert_eq!(anchors, names);
        let instr: Vec<String> = all(&d.root, "w:instrText").iter().map(|e| e.text()).collect();
        assert!(instr[1].contains("_Toc1_m1") && instr[2].contains("_Toc1_m2"));
        let simple: Vec<String> = all(&d.root, "w:fldSimple").iter().map(|e| e.attr("w:instr").unwrap().to_string()).collect();
        assert!(simple[1].contains("_Toc1_m1"));
        assert!(all(&d.root, "w:dataBinding").len() == 1, "привязка у первого документа остаётся");
        let uniq = |v: Vec<String>| v.iter().collect::<HashSet<_>>().len() == v.len();
        assert!(uniq(all(&d.root, "w:ins").iter().map(|e| e.attr("w:id").unwrap().to_string()).collect()));
        assert!(uniq(all(&d.root, "w:id").iter().map(|e| e.val().unwrap().to_string()).collect()) || all(&d.root, "w:id").len() == 1);
        assert!(uniq(all(&d.root, "wp:docPr").iter().map(|e| e.attr("id").unwrap().to_string()).collect()));
    }

    #[test]
    fn namespaces_ignorable_and_numbering_order_are_merged() {
        let dir = tempfile::tempdir().unwrap();
        let num = |extra: &str| {
            format!(
                r#"<?xml version="1.0"?><w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:abstractNum w:abstractNumId="0"><w:nsid w:val="AAAAAAAA"/><w:lvl w:ilvl="0"><w:numFmt w:val="decimal"/></w:lvl></w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>{extra}</w:numbering>"#
            )
        };
        let li = r#"<w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr><w:r><w:t>i</w:t></w:r></w:p>"#;
        let mut a = fx(&format!("{li}{}", sect("")));
        a.numbering = Some(num(r#"<w:numIdMacAtCleanup w:val="1"/>"#));
        let mut b = fx(&format!(r#"<w:p zz:tag="x"><w:r><w:t>i</w:t></w:r></w:p>{li}{}"#, sect("")));
        b.ns = b.ns.replace(r#"mc:Ignorable="w14""#, r#"xmlns:zz="urn:vendor:private" mc:Ignorable="w14 zz""#);
        b.numbering = Some(num(""));
        let (out, _) = run(dir.path(), "ign", &[&a, &b, &b]);
        let (pkg, d) = doc_of(&out);
        let ig = d.root.attr("mc:Ignorable").unwrap();
        assert!(ig.split(' ').any(|t| t == "zz") && d.root.attr("xmlns:zz").is_some());
        // номера списков: у каждого документа свой, numIdMacAtCleanup остаётся последним
        let ids: Vec<String> = all(&d.root, "w:numId").iter().map(|e| e.val().unwrap().to_string()).collect();
        assert_eq!(ids.len(), 3);
        assert_eq!(ids.iter().collect::<HashSet<_>>().len(), 3);
        let n = xml::parse(pkg.get("word/numbering.xml").unwrap()).unwrap();
        assert_eq!(n.root.elems().last().unwrap().name, "w:numIdMacAtCleanup");
        let nsids: Vec<String> = all(&n.root, "w:nsid").iter().map(|e| e.val().unwrap().to_string()).collect();
        assert_eq!(nsids.iter().collect::<HashSet<_>>().len(), 3);
    }

    #[test]
    fn content_type_for_media_declared_only_by_override() {
        let dir = tempfile::tempdir().unwrap();
        let mk = |c: u8| {
            let mut f = fx(&format!("<w:p>{}</w:p>{}", img("rI", c as u32), sect("")))
                .rel("rI", "image", "media/image1.gif")
                .part("word/media/image1.gif", vec![c])
                .ct("/word/media/image1.gif", "image/gif");
            f.parts.retain(|_| true);
            f
        };
        let (a, b) = (mk(1), mk(2));
        let (out, _) = run(dir.path(), "ov", &[&a, &b]);
        let pkg = Pkg::read(&out).unwrap();
        let t = Types::parse(pkg.get("[Content_Types].xml").unwrap()).unwrap();
        assert_eq!(t.ct_default("gif"), Some("image/gif"), "типичные картинки объявляются через Default");
    }

    #[test]
    fn style_identity_by_name_and_conflict_warning() {
        let dir = tempfile::tempdir().unwrap();
        let a = fx(&format!(r#"<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>A</w:t></w:r></w:p>{}"#, sect("")));
        // у второго документа «Normal» имеет id "a" (локализованный Word), Heading1 — id "1" и другое оформление
        let mut b = fx(&format!(r#"<w:p><w:pPr><w:pStyle w:val="1"/></w:pPr><w:r><w:t>B</w:t></w:r></w:p><w:p><w:r><w:t>plain</w:t></w:r></w:p>{}"#, sect("")));
        b.styles = styles_xml("").replace(r#"w:styleId="Normal""#, r#"w:styleId="a""#).replace(r#"w:val="Normal""#, r#"w:val="a""#).replace(r#"w:styleId="Heading1""#, r#"w:styleId="1""#).replace("<w:b/>", "<w:i/>")
            .replace("</w:styles>", r#"<w:style w:type="paragraph" w:customStyle="1" w:styleId="Mine"><w:name w:val="Mine"/><w:basedOn w:val="a"/></w:style></w:styles>"#);
        b.body = b.body.replace("plain", "plain</w:t></w:r></w:p><w:p><w:pPr><w:pStyle w:val=\"Mine\"/></w:pPr><w:r><w:t>mine");
        let (out, w) = run(dir.path(), "styles", &[&a, &b]);
        let (pkg, d) = doc_of(&out);
        let st = xml::parse(pkg.get("word/styles.xml").unwrap()).unwrap();
        let styles: Vec<&Elem> = st.root.elems().filter(|e| e.name == "w:style").collect();
        // дубликатов «Normal» и второго стиля по умолчанию нет
        assert_eq!(styles.iter().filter(|s| s.attr("w:default").is_some()).count(), 1);
        assert_eq!(styles.iter().filter(|s| s.child("w:name").unwrap().val() == Some("heading 1")).count(), 1);
        // ссылки на стили второго документа перенаправлены на стиль первого / импортированный
        let ps: Vec<String> = all(&d.root, "w:pStyle").iter().map(|e| e.val().unwrap().to_string()).collect();
        assert_eq!(ps, ["Heading1", "Heading1", "Mine"]);
        let mine = styles.iter().find(|s| s.attr("w:styleId") == Some("Mine")).unwrap();
        assert_eq!(mine.child("w:basedOn").unwrap().val(), Some("Normal"));
        assert!(w.iter().any(|m| m.contains("heading 1")), "{w:?}");
    }

    #[test]
    fn default_formatting_difference_is_reported() {
        // docDefaults отличаются -> предупреждение
        let dir = tempfile::tempdir().unwrap();
        let a = fx(&format!("{}{}", para("A"), sect("")));
        let mut b = fx(&format!("{}{}", para("B"), sect("")));
        b.styles = styles_xml("").replace(r#"w:val="22""#, r#"w:val="30""#);
        let (_, w) = run(dir.path(), "dd", &[&a, &b]);
        assert!(w.iter().any(|m| m.contains("по умолчанию")), "{w:?}");
    }

    #[test]
    fn nonstandard_prefixes_and_strict_format() {
        let dir = tempfile::tempdir().unwrap();
        let a = fx(&format!("{}{}", para("A"), sect("")));
        let mut b = fx("");
        b.ns = r#"xmlns:x="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:rel="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:w="urn:other""#.into();
        b.root = "x:document".into();
        b.body = String::new();
        let p = dir.path().join("b.docx");
        b.write(&p);
        // подменяем тело на x:-префиксы
        let mut pkg = Pkg::read(&p).unwrap();
        pkg.set(
            "word/document.xml",
            format!(r#"<x:document {}><x:body><x:p><x:hyperlink rel:id="rL"><x:r><x:t>PFX</x:t></x:r></x:hyperlink></x:p><w:odd/><x:sectPr/></x:body></x:document>"#, b.ns).into_bytes(),
        );
        let mut rels = String::from_utf8(pkg.get("word/_rels/document.xml.rels").unwrap().to_vec()).unwrap();
        rels = rels.replace("</Relationships>", &format!(r#"<Relationship Id="rL" Type="{RT}hyperlink" Target="https://x.y" TargetMode="External"/></Relationships>"#));
        pkg.set("word/_rels/document.xml.rels", rels.into_bytes());
        pkg.write(&p).unwrap();
        let pa = dir.path().join("a.docx");
        a.write(&pa);
        let out = dir.path().join("o.docx");
        merge(&[pa.clone(), p], &out, &Options::default()).unwrap();
        assert!(check(&out).is_empty(), "{:?}", check(&out));
        let (_, d) = doc_of(&out);
        assert_eq!(texts(d.root.child("w:body").unwrap()), ["A", "PFX"]);
        assert!(all(&d.root, "x1:odd").len() == 1, "чужое пространство имён переименовано, а не смешано с w:");
        // Strict
        let mut s = fx(&para("S"));
        s.ns = NS.replace("http://schemas.openxmlformats.org/wordprocessingml/2006/main", "http://purl.oclc.org/ooxml/wordprocessingml/main");
        let ps = dir.path().join("s.docx");
        s.write(&ps);
        let err = merge(&[pa, ps], &dir.path().join("e.docx"), &Options::default()).unwrap_err();
        assert!(format!("{err:#}").contains("Strict"));
    }

    #[test]
    fn empty_missing_and_dangling_parts_do_not_corrupt_output() {
        let dir = tempfile::tempdir().unwrap();
        let a = fx(&format!("{}{}", para("A"), sect("")));
        let empty = fx("");
        let nosect = fx(&para("NOSECT"));
        let bad = fx(&format!(
            r#"<w:p><w:hyperlink r:id="rNone"><w:r><w:t>h</w:t></w:r></w:hyperlink>{}</w:p>{}"#,
            img("rGone", 4),
            sect("")
        ))
        .rel("rGone", "image", "media/gone.png");
        let (out, w) = run(dir.path(), "odd", &[&a, &empty, &nosect, &bad]);
        assert_eq!(w.len(), 2, "{w:?}");
        let (_, d) = doc_of(&out);
        assert_eq!(all(&d.root, "w:sectPr").len(), 4);
        assert!(all(&d.root, "w:hyperlink").iter().all(|h| h.attr("r:id").is_none()));
    }

    #[test]
    fn unicode_part_names_and_single_input() {
        let dir = tempfile::tempdir().unwrap();
        let mk = |c: u8| {
            fx(&format!("<w:p>{}</w:p>{}", img("rU", c as u32), sect("")))
                .rel("rU", "image", "media/%D0%BA%D0%B0%D1%80%D1%82%20%D1%8B.png")
                .part("word/media/карт ы.png", png(c))
        };
        let (a, b) = (mk(1), mk(2));
        let (out, _) = run(dir.path(), "uni", &[&a, &b]);
        let pkg = Pkg::read(&out).unwrap();
        assert!(pkg.get("word/media/карт ы.png").is_some() && pkg.get("word/media/карт ы1.png").is_some());
        // один файл: результат — тот же пакет
        let src = dir.path().join("single.docx");
        a.write(&src);
        let o = dir.path().join("single_out.docx");
        assert!(merge(&[src.clone()], &o, &Options::default()).unwrap().is_empty());
        let (p1, p2) = (Pkg::read(&src).unwrap(), Pkg::read(&o).unwrap());
        for (n, d) in &p1.entries {
            assert_eq!(p2.get(n).unwrap(), d.as_slice(), "{n}");
        }
        assert!(merge(&[], &o, &Options::default()).is_err());
    }

    #[test]
    fn main_part_content_type_is_normalized_and_empty_header_is_valid() {
        let dir = tempfile::tempdir().unwrap();
        let a = fx(&format!("{}{}", para("A"), sect(r#"<w:headerReference w:type="first" r:id="rH"/><w:footerReference w:type="even" r:id="rF"/>"#)))
            .part("word/h.xml", hdr("w:hdr", "H"))
            .part("word/f.xml", hdr("w:ftr", "F"))
            .rel("rH", "header", "h.xml")
            .rel("rF", "footer", "f.xml")
            .ct("/word/h.xml", &format!("{CT_NOTES}header+xml"))
            .ct("/word/f.xml", &format!("{CT_NOTES}footer+xml"));
        let b = fx(&format!("{}{}", para("B"), sect("")));
        let (out, _) = run(dir.path(), "hf", &[&a, &b]);
        let (_, d) = doc_of(&out);
        let last = d.root.child("w:body").unwrap().child("w:sectPr").unwrap().clone();
        let refs = sect_refs(&last);
        assert!(refs.contains(&(false, "first".into())) && refs.contains(&(true, "even".into())) && refs.len() == 2, "{refs:?}");
    }
    #[test]
    fn settings_theme_and_embedded_font_differences_are_reported() {
        let dir = tempfile::tempdir().unwrap();
        let theme = |c: &str| format!(r#"<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:themeElements><a:clrScheme name="x"><a:accent1><a:srgbClr val="{c}"/></a:accent1></a:clrScheme></a:themeElements></a:theme>"#);
        let settings = |extra: &str| format!(r#"<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">{extra}</w:settings>"#);
        let base = |extra: &str, color: &str| {
            fx(&format!(r#"<w:p><w:r><w:rPr><w:color w:themeColor="accent1" w:val="000000"/></w:rPr><w:t>T</w:t></w:r></w:p>{}"#, sect("")))
                .part("word/settings.xml", settings(extra))
                .rel("rS", "settings", "settings.xml")
                .ct("/word/settings.xml", &format!("{CT_NOTES}settings+xml"))
                .part("word/theme/theme1.xml", theme(color))
                .rel("rT", "theme", "theme/theme1.xml")
                .ct("/word/theme/theme1.xml", "application/vnd.openxmlformats-officedocument.theme+xml")
        };
        let a = base("", "FF0000");
        let b = base("<w:evenAndOddHeaders/>", "00FF00")
            .part("word/fontTable.xml", r#"<w:fonts xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>"#)
            .part("word/_rels/fontTable.xml.rels", format!(r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="{RT}font" Target="fonts/font1.odttf"/></Relationships>"#))
            .part("word/fonts/font1.odttf", b"F")
            .rel("rFt", "fontTable", "fontTable.xml")
            .ct("/word/fontTable.xml", &format!("{CT_NOTES}fontTable+xml"))
            .ct("/word/fonts/font1.odttf", "application/vnd.openxmlformats-officedocument.obfuscatedFont");
        let (_, w) = run(dir.path(), "compat", &[&a, &b]);
        assert!(w.iter().any(|m| m.contains("чётных")), "{w:?}");
        assert!(w.iter().any(|m| m.contains("встроенные")), "{w:?}");
        assert!(w.iter().any(|m| m.contains("тема оформления")), "{w:?}");
        // одинаковые файлы — без предупреждений
        let (_, w) = run(dir.path(), "same", &[&a, &a]);
        assert!(w.is_empty(), "{w:?}");
    }
    #[test]
    fn sectpr_written_first_in_ppr_is_moved_to_schema_position() {
        let dir = tempfile::tempdir().unwrap();
        let odd = |t: &str| {
            fx(&format!(
                r#"<w:p><w:pPr><w:sectPr><w:pgSz w:w="1" w:h="1"/></w:sectPr><w:pStyle w:val="Normal"/><w:rPr/></w:pPr><w:r><w:t>{t}</w:t></w:r></w:p>{}"#,
                sect("")
            ))
        };
        let (a, b) = (odd("A"), odd("B"));
        let (out, _) = run(dir.path(), "ppr", &[&a, &b]);
        let (_, d) = doc_of(&out);
        for ppr in all(&d.root, "w:pPr") {
            let names: Vec<&str> = ppr.elems().map(|e| e.name.as_str()).collect();
            assert_eq!(names.last(), Some(&"w:sectPr"), "{names:?}");
        }
    }
}
