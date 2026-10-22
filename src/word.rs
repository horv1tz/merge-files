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
//! Оформление: по умолчанию (`Options::keep_source_styles`) каждый документ выглядит так же, как и
//! отдельно. Стиль с тем же именем, но другим определением добавляется под новым именем
//! («Заголовок 1 (2)»), ссылки на него переписываются, нумерация, заданная стилем, получает свой
//! список (и начинается заново в каждом документе), отличия docDefaults/«обычного» стиля/темы
//! переносятся в добавленные стили. Одинаковые стили по-прежнему сливаются в один. Режим «как в Word»
//! (`keep_source_styles: false`): побеждают стили первого документа, о замене предупреждается.
//!
//! Что не переносится без потерь, не пропадает молча: `merge` возвращает предупреждения для
//! пользователя. Это: настройки, общие для всего файла (зеркальные поля, фон страницы, режим
//! совместимости), встроенные шрифты, привязка элементов управления к данным, «битые» ссылки
//! исходных файлов, обновление полей (оглавление). Формат Strict Open XML не поддерживается (ошибка).
//! Повреждённые файлы (обрезанный архив, неправильный XML) отклоняются с понятным сообщением.

use anyhow::{Result, anyhow, bail};
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

    /// Допустимый в XML 1.0 символ.
    fn is_xml_char(c: char) -> bool {
        matches!(c, '\t' | '\n' | '\r' | '\u{20}'..='\u{D7FF}' | '\u{E000}'..='\u{FFFD}' | '\u{10000}'..='\u{10FFFF}')
    }

    /// Байты части XML -> текст: UTF-8 (с BOM или без) и UTF-16 (с BOM или без; OPC допускает обе кодировки).
    pub fn decode(data: &[u8]) -> Result<String> {
        let utf16 = if data.starts_with(&[0xFF, 0xFE]) {
            Some((false, &data[2..]))
        } else if data.starts_with(&[0xFE, 0xFF]) {
            Some((true, &data[2..]))
        } else if data.len() >= 2 && data[0] == b'<' && data[1] == 0 {
            Some((false, data))
        } else if data.len() >= 2 && data[0] == 0 && data[1] == b'<' {
            Some((true, data))
        } else {
            None
        };
        if let Some((be, body)) = utf16 {
            if body.len() % 2 != 0 {
                bail!("XML обрезан: файл оборван на середине символа (UTF-16)");
            }
            let units = body.chunks_exact(2).map(|c| if be { u16::from_be_bytes([c[0], c[1]]) } else { u16::from_le_bytes([c[0], c[1]]) });
            return char::decode_utf16(units)
                .collect::<std::result::Result<String, _>>()
                .map_err(|_| anyhow!("XML (UTF-16) содержит недопустимые символы — файл повреждён"));
        }
        let data = data.strip_prefix(&[0xEF, 0xBB, 0xBF][..]).unwrap_or(data);
        match std::str::from_utf8(data) {
            Ok(s) => Ok(s.to_string()),
            // Оборванный на середине символа конец — признак обрезанного файла, а не другой кодировки.
            Err(e) if e.error_len().is_none() => bail!("XML обрезан: файл оборван на середине символа"),
            Err(_) => bail!("XML содержит байты, недопустимые в UTF-8 (файл повреждён или сохранён в другой кодировке)"),
        }
    }

    /// Проверяет текст или значение атрибута: ссылки на сущности и допустимость символов.
    fn check_text(s: &str) -> Result<()> {
        let b = s.as_bytes();
        let mut i = 0;
        while i < b.len() {
            match b[i] {
                b'&' => {
                    let end = b[i + 1..].iter().take(12).position(|&x| x == b';');
                    let ent = end.map(|j| &s[i + 1..i + 1 + j]);
                    let ok = match ent {
                        Some("lt" | "gt" | "amp" | "quot" | "apos") => true,
                        Some(e) => e
                            .strip_prefix("#x")
                            .and_then(|h| u32::from_str_radix(h, 16).ok().filter(|_| h.bytes().all(|c| c.is_ascii_hexdigit())))
                            .or_else(|| e.strip_prefix('#').and_then(|d| d.parse::<u32>().ok().filter(|_| d.bytes().all(|c| c.is_ascii_digit()))))
                            .and_then(char::from_u32)
                            .is_some_and(is_xml_char),
                        None => false,
                    };
                    if !ok {
                        let shown: String = s[i..].chars().take_while(|c| !c.is_whitespace() && *c != '<').take(10).collect();
                        bail!("XML повреждён: недопустимая ссылка «{shown}» (символ «&» нужно записывать как «&amp;»)");
                    }
                    i += end.map_or(1, |j| j + 2);
                }
                0..=8 | 11 | 12 | 14..=31 => bail!("XML повреждён: недопустимый управляющий символ (код {})", b[i]),
                _ => i += 1,
            }
        }
        if !s.is_ascii() && !s.chars().all(is_xml_char) {
            bail!("XML повреждён: недопустимый символ");
        }
        Ok(())
    }

    pub fn parse(data: &[u8]) -> Result<Doc> {
        let text = decode(data)?;
        let s = text.as_str();
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
                    Some(p) => {
                        check_text(t)?;
                        p.children.push(Node::Text(t.to_string()))
                    }
                    None if !t.trim().is_empty() => bail!("XML повреждён: текст вне корневого элемента"),
                    None => {}
                }
                i = j;
                continue;
            }
            let rest = &s[i..];
            if rest.starts_with("<!--") {
                let e = rest.find("-->").ok_or_else(|| anyhow!("XML обрезан: не закрыт комментарий"))? + 3;
                if let Some(p) = stack.last_mut() {
                    p.children.push(Node::Raw(rest[..e].to_string()));
                } else if root.is_none() {
                    prolog.push(rest[..e].to_string());
                }
                i += e;
            } else if rest.starts_with("<![CDATA[") {
                let e = rest.find("]]>").ok_or_else(|| anyhow!("XML обрезан: не закрыт блок CDATA"))? + 3;
                match stack.last_mut() {
                    Some(p) => p.children.push(Node::Raw(rest[..e].to_string())),
                    None => bail!("XML повреждён: CDATA вне корневого элемента"),
                }
                i += e;
            } else if rest.starts_with("<?") {
                let e = rest.find("?>").ok_or_else(|| anyhow!("XML обрезан: не закрыта инструкция обработки"))? + 2;
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
                let e = rest.find('>').ok_or_else(|| anyhow!("XML обрезан: не закрыт тег"))?;
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
                        bail!("XML обрезан: не закрыт тег <{}", el.name);
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
                    let q = *tb.get(j).ok_or_else(|| anyhow!("XML обрезан: оборван тег <{}", el.name))?;
                    if q != b'"' && q != b'\'' {
                        bail!("XML повреждён: значение атрибута {an} без кавычек");
                    }
                    j += 1;
                    let vs = j;
                    while j < tb.len() && tb[j] != q {
                        j += 1;
                    }
                    if j >= tb.len() {
                        bail!("XML обрезан: не закрыто значение атрибута {an}");
                    }
                    check_text(&rest[vs..j])?;
                    if rest[vs..j].contains('<') {
                        bail!("XML повреждён: символ «<» внутри значения атрибута {an}");
                    }
                    if el.attrs.iter().any(|(n, _)| *n == an) {
                        bail!("XML повреждён: атрибут {an} повторяется в <{}>", el.name);
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
            bail!("XML обрезан: не закрыт элемент <{}>", e.name);
        }
        Ok(Doc { prolog, root: root.ok_or_else(|| anyhow!("XML пуст или обрезан"))? })
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

/// Ошибка ввода-вывода без «сырого» системного текста.
fn io_msg(e: &std::io::Error) -> String {
    use std::io::ErrorKind as K;
    match e.kind() {
        K::NotFound => "файл не найден".into(),
        K::PermissionDenied => "нет доступа к файлу (возможно, он открыт в другой программе)".into(),
        K::UnexpectedEof => "файл обрезан (неожиданный конец данных)".into(),
        K::InvalidData => "данные повреждены (ошибка контрольной суммы или сжатия)".into(),
        _ => match e.raw_os_error() {
            Some(c) => format!("ошибка ввода-вывода (код {c})"),
            None => "ошибка ввода-вывода".into(),
        },
    }
}

fn zip_msg(e: &zip::result::ZipError) -> String {
    use zip::result::ZipError as Z;
    match e {
        Z::Io(e) => io_msg(e),
        Z::InvalidPassword => "документ защищён паролем".into(),
        Z::UnsupportedArchive(m) if m.to_lowercase().contains("password") => "документ защищён паролем".into(),
        Z::UnsupportedArchive(_) | Z::CompressionMethodNotSupported(_) => "неподдерживаемый вариант архива ZIP".into(),
        _ => "файл повреждён (обрезан или не является архивом ZIP)".into(),
    }
}

/// Ошибка разбора части пакета для пользователя.
fn part_err(part: &str, e: impl std::fmt::Display) -> anyhow::Error {
    anyhow!("часть «{part}» повреждена: {e}")
}

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
        let f = std::fs::File::open(path).map_err(|e| anyhow!("не удалось открыть файл: {}", io_msg(&e)))?;
        let mut zip = zip::ZipArchive::new(f).map_err(|e| anyhow!("{}", zip_msg(&e)))?;
        let mut entries: Vec<(String, Vec<u8>)> = Vec::new();
        let mut seen = HashSet::new();
        for i in 0..zip.len() {
            let mut e = zip.by_index(i).map_err(|er| anyhow!("не удалось прочитать запись №{} архива: {}", i + 1, zip_msg(&er)))?;
            if e.is_dir() {
                continue;
            }
            let name = e.name().map(|n| n.to_string()).unwrap_or_default().replace('\\', "/");
            let name = name.trim_start_matches('/').to_string();
            if name.is_empty() || !seen.insert(name.to_lowercase()) {
                continue;
            }
            let mut buf = Vec::new();
            e.read_to_end(&mut buf).map_err(|er| anyhow!("файл повреждён: не удалось прочитать «{name}» — {}", io_msg(&er)))?;
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
        let f = std::fs::File::create(path).map_err(|e| anyhow!("не удалось создать файл «{}»: {}", path.display(), io_msg(&e)))?;
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
            zip.start_file(name.as_str(), opts.large_file(data.len() as u64 > 3_900_000_000))
                .map_err(|e| anyhow!("не удалось записать результат: {}", zip_msg(&e)))?;
            zip.write_all(data).map_err(|e| anyhow!("не удалось записать результат: {}", io_msg(&e)))?;
        }
        zip.finish().map_err(|e| anyhow!("не удалось записать результат: {}", zip_msg(&e)))?.flush().map_err(|e| anyhow!("не удалось записать результат: {}", io_msg(&e)))?;
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

/// Токены-префиксы в значении `mc:Ignorable` / `Requires` -> общепринятые префиксы.
fn map_prefix_tokens(v: &str, f: &dyn Fn(&str) -> Option<String>) -> String {
    v.split_whitespace().map(|t| f(t).unwrap_or_else(|| t.to_string())).collect::<Vec<_>>().join(" ")
}

/// Приводит все префиксы пространств имён документа к общепринятым (`w`, `r`, `wp`, ...), чтобы остальной
/// код мог опираться на имена вида `w:p`: учитываются объявления на любой глубине и пространство имён по
/// умолчанию (`xmlns="…"`, элементы без префикса), все объявления собираются в корне. Атрибуты без
/// префикса не принадлежат никакому пространству имён и не меняются.
fn normalize_ns(root: &mut Elem) -> Result<()> {
    let root_ptr = root as *const Elem;
    // Быстрый путь: всё уже в общепринятом виде.
    let mut simple = true;
    let mut decls: Vec<(String, String)> = vec![];
    xml::walk_ref(root, &mut |e| {
        let is_root = std::ptr::eq(e, root_ptr);
        for (k, v) in &e.attrs {
            if k == "xmlns" {
                simple = false;
            } else if let Some(p) = k.strip_prefix("xmlns:") {
                decls.push((p.to_string(), v.clone()));
                let canon = canon_prefix(v);
                if !is_root || canon.is_some_and(|c| c != p) || (is_canon_prefix(p) && canon != Some(p)) {
                    simple = false;
                }
            }
        }
    });
    if simple {
        return check_prefixes(root, &decls);
    }
    // Итоговый префикс для каждого адреса.
    let mut chosen: HashMap<String, String> = HashMap::new();
    let mut taken: HashSet<String> = HashSet::new();
    for (_, u) in &decls {
        if let Some(c) = canon_prefix(u) {
            chosen.insert(u.clone(), c.to_string());
            taken.insert(c.to_string());
        }
    }
    let mut n = 0;
    let mut want: Vec<(String, String)> = decls.clone();
    // Адреса из xmlns="…" тоже нужны в таблице.
    xml::walk_ref(root, &mut |e| {
        if let Some(u) = e.attr("xmlns") {
            if !u.is_empty() {
                want.push((String::new(), u.to_string()));
            }
        }
    });
    for (p, u) in &want {
        if chosen.contains_key(u) {
            continue;
        }
        let mut t = p.clone();
        while t.is_empty() || is_canon_prefix(&t) || taken.contains(&t) {
            n += 1;
            t = format!("x{n}");
        }
        taken.insert(t.clone());
        chosen.insert(u.clone(), t);
    }

    fn lookup<'a>(scope: &'a [(String, String)], p: &str) -> Option<&'a str> {
        scope.iter().rev().find(|(k, _)| k == p).map(|(_, v)| v.as_str())
    }
    fn rw(e: &mut Elem, scope: &mut Vec<(String, String)>, chosen: &HashMap<String, String>) -> Result<()> {
        let mark = scope.len();
        for (k, v) in &e.attrs {
            if k == "xmlns" {
                scope.push((String::new(), v.clone()));
            } else if let Some(p) = k.strip_prefix("xmlns:") {
                scope.push((p.to_string(), v.clone()));
            }
        }
        let map_name = |name: &str, scope: &[(String, String)], is_attr: bool| -> Result<Option<String>> {
            match name.split_once(':') {
                Some(("xml", _)) | Some(("xmlns", _)) => Ok(None),
                Some((p, l)) => {
                    let u = lookup(scope, p).ok_or_else(|| anyhow!("XML повреждён: префикс «{p}:» не объявлен"))?;
                    Ok(Some(format!("{}:{l}", chosen.get(u).map(String::as_str).unwrap_or(p))))
                }
                None if is_attr => Ok(None),
                None => Ok(match lookup(scope, "") {
                    Some(u) if !u.is_empty() => chosen.get(u).map(|c| format!("{c}:{name}")),
                    _ => None,
                }),
            }
        };
        if let Some(n) = map_name(&e.name, scope, false)? {
            e.name = n;
        }
        let tokens = |v: &str, scope: &[(String, String)]| {
            map_prefix_tokens(v, &|t| lookup(scope, t).and_then(|u| chosen.get(u)).cloned())
        };
        let is_choice = e.name == "mc:Choice";
        let old = std::mem::take(&mut e.attrs);
        for (k, v) in old {
            if k == "xmlns" || k.starts_with("xmlns:") {
                continue;
            }
            let nk = map_name(&k, scope, true)?.unwrap_or_else(|| k.clone());
            let nv = if nk == "mc:Ignorable" || nk == "mc:ProcessContent" || nk == "mc:PreserveElements" || (is_choice && nk == "Requires") {
                tokens(&v, scope)
            } else {
                v
            };
            e.attrs.push((nk, nv));
        }
        for c in e.children.iter_mut() {
            if let Node::Elem(c) = c {
                rw(c, scope, chosen)?;
            }
        }
        scope.truncate(mark);
        Ok(())
    }
    rw(root, &mut vec![], &chosen)?;
    // Объявления - в корне: сначала в порядке первого появления, затем остальные атрибуты.
    let mut seen = HashSet::new();
    let mut head: Vec<(String, String)> = vec![];
    for (_, u) in decls.iter().chain(want.iter()) {
        if let Some(p) = chosen.get(u) {
            if seen.insert(p.clone()) {
                head.push((format!("xmlns:{p}"), u.clone()));
            }
        }
    }
    head.extend(std::mem::take(&mut root.attrs));
    root.attrs = head;
    Ok(())
}

/// Все ли префиксы элементов и атрибутов объявлены (иначе XML не является корректным).
fn check_prefixes(root: &Elem, decls: &[(String, String)]) -> Result<()> {
    let declared: HashSet<&str> = decls.iter().map(|(p, _)| p.as_str()).collect();
    let mut bad: Option<String> = None;
    xml::walk_ref(root, &mut |e| {
        let names = std::iter::once(e.name.as_str()).chain(e.attrs.iter().map(|(k, _)| k.as_str()));
        for n in names {
            if let Some((p, _)) = n.split_once(':') {
                if p != "xml" && p != "xmlns" && !declared.contains(p) && bad.is_none() {
                    bad = Some(p.to_string());
                }
            }
        }
    });
    match bad {
        Some(p) => bail!("XML повреждён: префикс «{p}:» не объявлен"),
        None => Ok(()),
    }
}

/// Приводит документ к виду, на который опирается остальной код (см. `normalize_ns`).
/// Strict Open XML не поддерживается.
fn harmonize(doc: &mut xml::Doc) -> Result<()> {
    let mut strict = false;
    xml::walk_ref(&doc.root, &mut |e| {
        strict |= e.attrs.iter().any(|(k, v)| k.starts_with("xmlns") && v.starts_with("http://purl.oclc.org/ooxml/"));
    });
    if strict {
        bail!("формат Strict Open XML не поддерживается (сохраните документ как обычный .docx)");
    }
    normalize_ns(&mut doc.root)
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
        let types = Types::parse(
            pkg.get("[Content_Types].xml").ok_or_else(|| anyhow!("в файле нет [Content_Types].xml — это не документ Word или файл повреждён"))?,
        )
        .map_err(|e| part_err("[Content_Types].xml", e))?;
        let root_rels = match pkg.get("_rels/.rels") {
            Some(b) => Rels::parse(b).map_err(|e| part_err("_rels/.rels", e))?,
            None => Rels::default(),
        };
        let main = root_rels
            .items
            .iter()
            .filter(|r| !r.external && rel_suffix(&r.ty) == "officeDocument")
            .find_map(|r| resolve_part(&pkg, "", &r.target))
            .or_else(|| pkg.name_of("word/document.xml"))
            .ok_or_else(|| anyhow!("в файле нет основной части документа (word/document.xml) — это не документ Word или файл повреждён"))?;
        let rels = match pkg.get(&rels_path(&main)) {
            Some(b) => Rels::parse(b).map_err(|e| part_err(&rels_path(&main), e))?,
            None => Rels::default(),
        };
        let mut doc = xml::parse(pkg.get(&main).unwrap_or_default()).map_err(|e| part_err(&main, e))?;
        harmonize(&mut doc).map_err(|e| part_err(&main, e))?;
        if doc.root.name != "w:document" {
            bail!("это не документ Word (основная часть «{main}» не содержит w:document)");
        }
        if doc.root.child("w:body").is_none() {
            bail!("в документе нет <w:body>");
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
        let mut d = xml::parse(self.pkg.get(path).unwrap_or_default()).map_err(|e| part_err(path, e))?;
        harmonize(&mut d).map_err(|e| part_err(path, e))?;
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
    /// Переименованные стили: прежнее имя (строчными) -> новое имя (для полей STYLEREF).
    style_names: HashMap<String, String>,
    /// Порядковый номер документа в результате (2 — второй): для имён вида «Заголовок 1 (2)».
    ordinal: usize,
    /// Стиль абзаца, который получают абзацы без `pStyle` (в стилях результата).
    default_target: Option<String>,
    /// Стиль таблицы по умолчанию, который нужно указать таблицам без `tblStyle`.
    inject_tblstyle: Option<String>,
    /// Цепочки стилей результата (для проверки нумерации, заданной стилем).
    style_info: HashMap<String, StyleLite>,
    /// Номер списка результата -> новый номер списка этого документа (с началом с 1); `None` — не нужен.
    restart: HashMap<u64, Option<u64>>,
    next_num: u64,
    /// У документа включены разные колонтитулы чётных и нечётных страниц.
    src_even: bool,
    /// Копия «обычного» стиля для абзацев внутри таблиц (см. `merge_styles`) и идентификатор исходного.
    cell_target: Option<String>,
    src_default_id: Option<String>,
    /// Тема документа отличается от темы результата, и документ ею пользуется.
    theme_diff: bool,
    /// Тема документа, если её цвета и шрифты нужно заменить явными значениями.
    theme: Option<Theme>,
    /// Сколько полей оглавления/указателей найдено в документе.
    toc: usize,
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

/// Что стиль говорит о нумерации и отступах абзаца.
#[derive(Clone, Debug, Default)]
struct StyleLite {
    based: Option<String>,
    num: Option<u64>,
    ilvl: Option<String>,
    ind: Option<Vec<(String, String)>>,
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
    /// Сохранять оформление каждого документа (см. `Options`).
    keep: bool,
    /// Сколько документов уже добавлено.
    appended: usize,
    /// Хотя бы у одного из документов включены разные колонтитулы чётных и нечётных страниц:
    /// настройка действует на весь файл, поэтому остальным документам колонтитулы копируются на чётные страницы.
    any_even: bool,
    /// Копии стилей под новыми именами: (тип, прежнее имя строчными) -> id копий.
    variants: HashMap<(String, String), Vec<String>>,
    /// docDefaults, с которыми стиль был добавлен в результат (для стилей, добавленных из других документов).
    style_dsig: HashMap<String, String>,
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

/// Для колонтитулов «по умолчанию» раздела, у которых нет колонтитула чётных страниц, добавляет ссылку
/// на тот же колонтитул. Возвращает true, если что-то добавлено.
fn add_even_refs(s: &mut Elem) -> bool {
    let mut add = vec![];
    for tag in ["w:headerReference", "w:footerReference"] {
        let has = |ty: &str| s.elems().find(|e| e.name == tag && e.attr("w:type").unwrap_or("default") == ty);
        if let (Some(d), None) = (has("default"), has("even")) {
            if let Some(rid) = d.attr("r:id") {
                add.push(Elem::new(tag).with("w:type", "even").with("r:id", rid));
            }
        }
    }
    if add.is_empty() {
        return false;
    }
    let at = s
        .children
        .iter()
        .rposition(|n| matches!(n, Node::Elem(e) if e.name == "w:headerReference" || e.name == "w:footerReference"))
        .map_or(0, |i| i + 1);
    for (k, a) in add.into_iter().enumerate() {
        s.children.insert(at + k, Node::Elem(a));
    }
    true
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

/// `STYLEREF "Имя стиля"`: имя стиля заменяется новым, если стиль был переименован.
fn styleref_rename(s: &str, names: &HashMap<String, String>) -> String {
    let t = s.trim_start();
    let lead = &s[..s.len() - t.len()];
    let Some(head) = t.get(..8).filter(|h| h.eq_ignore_ascii_case("STYLEREF")) else { return s.to_string() };
    let rest = &t[8..];
    let body = rest.trim_start();
    let gap = &rest[..rest.len() - body.len()];
    let (name, tail) = match body.strip_prefix('"') {
        Some(b) => match b.find('"') {
            Some(i) => (&b[..i], &b[i + 1..]),
            None => return s.to_string(),
        },
        None => {
            let i = body.find(char::is_whitespace).unwrap_or(body.len());
            (&body[..i], &body[i..])
        }
    };
    match names.get(&name.to_lowercase()) {
        Some(n) => format!("{lead}{head}{gap}\"{n}\"{tail}"),
        None => s.to_string(),
    }
}

/// Инструкция поля целиком: закладки и стили, на которые она ссылается, получают новые имена.
fn rename_instr(s: &str, bm: &HashMap<String, String>, styles: &HashMap<String, String>) -> String {
    let mut r = if bm.is_empty() { s.to_string() } else { instr_rename(s, bm) };
    if !styles.is_empty() {
        r = styleref_rename(&r, styles);
    }
    r
}

/// Поле, которое нужно обновить после слияния (содержимое зависит от нумерации страниц).
fn is_toc_instr(s: &str) -> bool {
    let w = s.trim_start().split(|c: char| c.is_whitespace() || c == '\\').next().unwrap_or("");
    ["TOC", "TOA", "INDEX"].iter().any(|k| w.eq_ignore_ascii_case(k))
}

/// Переименовывает закладки и стили в инструкциях полей. Word разбивает инструкцию на несколько
/// `w:instrText` (на границах правок), поэтому для каждой сложной области (`fldChar begin … separate`)
/// фрагменты склеиваются, результат записывается в первый из них, остальные удаляются.
/// Возвращает число найденных полей оглавления/указателей.
fn rename_field_instrs(root: &mut Elem, bm: &HashMap<String, String>, styles: &HashMap<String, String>) -> usize {
    struct Field {
        idx: Vec<usize>,
        text: String,
        in_instr: bool,
    }
    let mut count = 0usize;
    let mut texts: Vec<String> = vec![];
    let mut groups: Vec<Vec<usize>> = vec![];
    let mut stack: Vec<Field> = vec![];
    let mut toc = 0usize;
    let finish = |f: Field, groups: &mut Vec<Vec<usize>>, toc: &mut usize| {
        if is_toc_instr(&f.text) {
            *toc += 1;
        }
        groups.push(f.idx);
    };
    xml::walk_ref(root, &mut |e| match e.name.as_str() {
        "w:fldChar" => match e.attr("w:fldCharType") {
            Some("begin") => stack.push(Field { idx: vec![], text: String::new(), in_instr: true }),
            Some("separate") => {
                if let Some(f) = stack.last_mut() {
                    f.in_instr = false;
                }
            }
            Some("end") => {
                if let Some(f) = stack.pop() {
                    finish(f, &mut groups, &mut toc);
                }
            }
            _ => {}
        },
        "w:instrText" => {
            let t = e.text();
            texts.push(t.clone());
            if let Some(f) = stack.last_mut().filter(|f| f.in_instr) {
                f.idx.push(count);
                f.text.push_str(&t);
            } else {
                groups.push(vec![count]);
            }
            count += 1;
        }
        "w:fldSimple" if e.attr("w:instr").is_some_and(is_toc_instr) => toc += 1,
        _ => {}
    });
    // Незакрытые поля тоже обрабатываются.
    while let Some(f) = stack.pop() {
        finish(f, &mut groups, &mut toc);
    }
    // План: индекс фрагмента -> Some(новый текст) | None (удалить).
    let mut plan: HashMap<usize, Option<String>> = HashMap::new();
    for g in &groups {
        if g.is_empty() {
            continue;
        }
        let joined: String = g.iter().map(|&i| texts[i].as_str()).collect();
        let renamed = rename_instr(&joined, bm, styles);
        if renamed != joined {
            plan.insert(g[0], Some(renamed));
            for &i in &g[1..] {
                plan.insert(i, None);
            }
        }
    }
    if !plan.is_empty() {
        let mut i = 0usize;
        xml::walk(root, &mut |e| {
            if e.name == "w:instrText" {
                let k = i;
                i += 1;
                match plan.get(&k) {
                    Some(Some(t)) => {
                        e.set_text(t);
                        if t.starts_with(char::is_whitespace) || t.ends_with(char::is_whitespace) {
                            e.set_attr("xml:space", "preserve");
                        }
                    }
                    Some(None) => return false,
                    None => {}
                }
            }
            true
        });
    }
    toc
}

/// Каноническая запись элемента для сравнения определений (без rsid и служебных атрибутов).
fn canon_xml(e: &Elem, skip: &[&str], out: &mut String) {
    out.push('<');
    out.push_str(&e.name);
    let mut attrs: Vec<&(String, String)> =
        e.attrs
            .iter()
            .filter(|(k, _)| !k.starts_with("w:rsid") && !k.starts_with("w14:") && !k.starts_with("xmlns") && k != "w:styleId" && k != "w:default" && k != "w:customStyle" && k != "w:abstractNumId")
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

fn truncate_chars(s: &str, n: usize) -> &str {
    match s.char_indices().nth(n) {
        Some((i, _)) => &s[..i],
        None => s,
    }
}

// ───────────────────────────── Стили ─────────────────────────────

const PPR_ORDER: &[&str] = &[
    "w:pStyle", "w:keepNext", "w:keepLines", "w:pageBreakBefore", "w:framePr", "w:widowControl", "w:numPr",
    "w:suppressLineNumbers", "w:pBdr", "w:shd", "w:tabs", "w:suppressAutoHyphens", "w:kinsoku", "w:wordWrap",
    "w:overflowPunct", "w:topLinePunct", "w:autoSpaceDE", "w:autoSpaceDN", "w:bidi", "w:adjustRightInd",
    "w:snapToGrid", "w:spacing", "w:ind", "w:contextualSpacing", "w:mirrorIndents", "w:suppressOverlap", "w:jc",
    "w:textDirection", "w:textAlignment", "w:textboxTightWrap", "w:outlineLvl", "w:divId", "w:cnfStyle", "w:rPr",
    "w:sectPr", "w:pPrChange",
];

const RPR_ORDER: &[&str] = &[
    "w:ins", "w:del", "w:moveFrom", "w:moveTo", "w:rStyle", "w:rFonts", "w:b", "w:bCs", "w:i", "w:iCs", "w:caps",
    "w:smallCaps", "w:strike", "w:dstrike", "w:outline", "w:shadow", "w:emboss", "w:imprint", "w:noProof",
    "w:snapToGrid", "w:vanish", "w:webHidden", "w:color", "w:spacing", "w:w", "w:kern", "w:position", "w:sz",
    "w:szCs", "w:highlight", "w:u", "w:effect", "w:bdr", "w:shd", "w:fitText", "w:vertAlign", "w:rtl", "w:cs",
    "w:em", "w:lang", "w:eastAsianLayout", "w:specVanish", "w:oMath", "w:rPrChange",
];

const STYLE_ORDER: &[&str] = &[
    "w:name", "w:aliases", "w:basedOn", "w:next", "w:link", "w:autoRedefine", "w:hidden", "w:uiPriority",
    "w:semiHidden", "w:unhideWhenUsed", "w:qFormat", "w:locked", "w:personal", "w:personalCompose",
    "w:personalReply", "w:rsid", "w:pPr", "w:rPr", "w:tblPr", "w:trPr", "w:tcPr", "w:tblStylePr",
];

/// Вставляет дочерний элемент на место, предусмотренное схемой (`order` — порядок имён).
fn insert_ordered(parent: &mut Elem, child: Elem, order: &[&str]) {
    let rank = |n: &str| order.iter().position(|x| *x == n);
    let Some(r) = rank(&child.name) else {
        parent.push(child);
        return;
    };
    let pos = parent
        .children
        .iter()
        .position(|n| matches!(n, Node::Elem(e) if rank(&e.name).is_some_and(|x| x > r)))
        .unwrap_or(parent.children.len());
    parent.children.insert(pos, Node::Elem(child));
}

fn flag_on(e: &Elem, name: &str) -> bool {
    matches!(e.attr(name), Some("1" | "true" | "on"))
}

/// Подпись нумераций документа: номер списка -> описание уровней без идентификаторов.
fn num_sigs(root: &Elem) -> HashMap<String, String> {
    const SKIP: &[&str] = &["w:nsid", "w:tmpl", "w:name", "w:pStyle", "w:styleLink", "w:numStyleLink"];
    let mut abs: HashMap<String, String> = HashMap::new();
    for e in root.elems().filter(|e| e.name == "w:abstractNum") {
        if let Some(id) = e.attr("w:abstractNumId") {
            let mut s = String::new();
            canon_xml(e, SKIP, &mut s);
            abs.insert(id.to_string(), s);
        }
    }
    let mut out = HashMap::new();
    for e in root.elems().filter(|e| e.name == "w:num") {
        let (Some(id), Some(a)) = (e.attr("w:numId"), e.child("w:abstractNumId").and_then(|a| a.val())) else { continue };
        let mut s = abs.get(a).cloned().unwrap_or_default();
        for o in e.elems().filter(|o| o.name == "w:lvlOverride") {
            s.push('|');
            canon_xml(o, SKIP, &mut s);
        }
        out.insert(id.trim().to_string(), s);
    }
    out
}

/// Собственное определение стиля (без родителя), номера списков заменены описанием списка.
fn style_own_sig(st: &Elem, nums: &HashMap<String, String>) -> String {
    let mut c = st.clone();
    xml::walk(&mut c, &mut |e| {
        if e.name == "w:numId" {
            if let Some(v) = e.val().map(|v| v.trim().to_string()) {
                let n = if v == "0" { v } else { format!("num:{}", nums.get(&v).map(String::as_str).unwrap_or("")) };
                e.set_attr("w:val", n);
            }
        }
        true
    });
    let mut s = String::new();
    canon_xml(&c, STYLE_SKIP, &mut s);
    s
}

// Нейтральные значения свойств «по умолчанию» (когда документ их не задаёт).
fn neutral_attr(elem: &str, attr: &str) -> Option<&'static str> {
    match (elem, attr) {
        ("w:spacing", "w:before" | "w:after" | "w:beforeLines" | "w:afterLines" | "w:val") => Some("0"),
        ("w:spacing", "w:line") => Some("240"),
        ("w:spacing", "w:lineRule") => Some("auto"),
        ("w:spacing", "w:beforeAutospacing" | "w:afterAutospacing") => Some("0"),
        ("w:ind", "w:left" | "w:right" | "w:start" | "w:end" | "w:firstLine" | "w:hanging") => Some("0"),
        ("w:ind", "w:leftChars" | "w:rightChars" | "w:firstLineChars" | "w:hangingChars") => Some("0"),
        _ => None,
    }
}

/// Элемент «выключено / по умолчанию» для свойства, которое задано в docDefaults первого документа.
fn neutral_elem(b: &Elem, is_rpr: bool) -> Option<Elem> {
    let n = b.name.as_str();
    let toggle_off = |e: &Elem| flag_on(e, "w:val") || e.attr("w:val").is_none();
    let off = || Elem::new(n).with("w:val", "0");
    if is_rpr {
        match n {
            "w:b" | "w:bCs" | "w:i" | "w:iCs" | "w:caps" | "w:smallCaps" | "w:strike" | "w:dstrike" | "w:vanish"
            | "w:shadow" | "w:emboss" | "w:imprint" | "w:outline"
                if toggle_off(b) =>
            {
                Some(off())
            }
            "w:sz" | "w:szCs" => Some(Elem::new(n).with("w:val", "20")),
            "w:color" => Some(Elem::new(n).with("w:val", "auto")),
            "w:kern" | "w:position" | "w:spacing" => Some(Elem::new(n).with("w:val", "0")),
            "w:w" => Some(Elem::new(n).with("w:val", "100")),
            _ => None,
        }
    } else {
        match n {
            "w:spacing" | "w:ind" => {
                let mut e = Elem::new(n);
                for (k, _) in &b.attrs {
                    if let Some(v) = neutral_attr(n, k) {
                        e.set_attr(k, v);
                    }
                }
                (!e.attrs.is_empty()).then_some(e)
            }
            "w:jc" => Some(Elem::new(n).with("w:val", "left")),
            "w:keepNext" | "w:keepLines" | "w:pageBreakBefore" | "w:suppressAutoHyphens" | "w:contextualSpacing" | "w:bidi"
            | "w:mirrorIndents"
                if toggle_off(b) =>
            {
                Some(off())
            }
            _ => None,
        }
    }
}

/// Свойство `c` из docDefaults документа-источника переносится в определение стиля `t`.
fn absorb_default(t: &mut Elem, c: &Elem, order: &[&str]) {
    let Some(i) = t.children.iter().position(|n| matches!(n, Node::Elem(e) if e.name == c.name)) else {
        insert_ordered(t, c.clone(), order);
        return;
    };
    let Node::Elem(te) = &mut t.children[i] else { return };
    if c.name == "w:rFonts" {
        // Каждая «группа» (обычный шрифт / шрифт темы) наследуется целиком.
        for group in [["w:ascii", "w:asciiTheme"], ["w:hAnsi", "w:hAnsiTheme"], ["w:eastAsia", "w:eastAsiaTheme"], ["w:cs", "w:cstheme"]] {
            if group.iter().all(|g| te.attr(g).is_none()) {
                for g in group {
                    if let Some(v) = c.attr(g) {
                        te.set_attr(g, v);
                    }
                }
            }
        }
    } else if matches!(c.name.as_str(), "w:spacing" | "w:ind" | "w:lang") {
        for (k, v) in &c.attrs {
            if te.attr(k).is_none() {
                te.set_attr(k, v.clone());
            }
        }
    }
}

/// Свойство `b` задано в docDefaults первого документа, а у источника его нет: чтобы оно не «просочилось»,
/// в определении стиля задаётся нейтральное значение.
fn neutralize_default(t: &mut Elem, b: &Elem, order: &[&str], is_rpr: bool) {
    match t.children.iter().position(|n| matches!(n, Node::Elem(e) if e.name == b.name)) {
        None => {
            if let Some(n) = neutral_elem(b, is_rpr) {
                insert_ordered(t, n, order);
            }
        }
        Some(i) => {
            let Node::Elem(te) = &mut t.children[i] else { return };
            if matches!(b.name.as_str(), "w:spacing" | "w:ind") {
                for (k, _) in &b.attrs {
                    if te.attr(k).is_none() {
                        if let Some(v) = neutral_attr(&b.name, k) {
                            te.set_attr(k, v);
                        }
                    }
                }
            } else if b.name == "w:rFonts" {
                for (g, th) in [("w:ascii", "w:asciiTheme"), ("w:hAnsi", "w:hAnsiTheme")] {
                    if (b.attr(g).is_some() || b.attr(th).is_some()) && te.attr(g).is_none() && te.attr(th).is_none() {
                        te.set_attr(g, "Times New Roman");
                    }
                }
            }
        }
    }
}

/// Переносит в корневой стиль абзаца (без `basedOn`) свойства docDefaults документа-источника: после
/// слияния действуют docDefaults первого документа, и текст без явного оформления иначе изменил бы вид.
fn push_defaults(st: &mut Elem, ds: Option<&Elem>, db: Option<&Elem>, skip: Option<&PprSkip>) {
    for (cont, prop, order, is_rpr) in [("w:pPrDefault", "w:pPr", PPR_ORDER, false), ("w:rPrDefault", "w:rPr", RPR_ORDER, true)] {
        let skip = skip.filter(|_| !is_rpr);
        let sp = ds.and_then(|d| d.child(cont)).and_then(|c| c.child(prop));
        let bp = db.and_then(|d| d.child(cont)).and_then(|c| c.child(prop));
        if sp.is_none() && bp.is_none() {
            continue;
        }
        if st.child(prop).is_none() {
            insert_ordered(st, Elem::new(prop), STYLE_ORDER);
        }
        let Some(t) = st.child_mut(prop) else { continue };
        if let Some(sp) = sp {
            for c in sp.elems() {
                if let Some(c) = without_skipped(c, skip) {
                    absorb_default(t, &c, order);
                }
            }
        }
        if let Some(bp) = bp {
            for c in bp.elems() {
                if let Some(c) = without_skipped(c, skip) {
                    neutralize_default(t, &c, order, is_rpr);
                }
            }
        }
    }
}

/// Свойства абзаца (имя элемента -> атрибуты), которые задают стили таблиц: у абзацев в ячейках таблицы
/// они сильнее docDefaults, но слабее явных свойств стиля абзаца.
type PprSkip = HashMap<String, HashSet<String>>;

fn without_skipped(c: &Elem, skip: Option<&PprSkip>) -> Option<Elem> {
    let Some(set) = skip.and_then(|s| s.get(&c.name)) else { return Some(c.clone()) };
    if !matches!(c.name.as_str(), "w:spacing" | "w:ind") {
        return None;
    }
    let mut c = c.clone();
    c.attrs.retain(|(k, _)| !set.contains(k));
    (!c.attrs.is_empty()).then_some(c)
}

/// Что задают в свойствах абзаца стили таблиц `tables` (включая условное оформление строк и столбцов).
fn table_ppr_props(tables: &[&Elem]) -> PprSkip {
    let mut out: PprSkip = HashMap::new();
    for st in tables {
        let pprs = st.child("w:pPr").into_iter().chain(st.elems().filter(|e| e.name == "w:tblStylePr").filter_map(|e| e.child("w:pPr")));
        for ppr in pprs {
            for c in ppr.elems() {
                let e = out.entry(c.name.clone()).or_default();
                if matches!(c.name.as_str(), "w:spacing" | "w:ind") {
                    e.extend(c.attrs.iter().map(|(k, _)| k.clone()));
                } else {
                    e.insert("*".into());
                }
            }
        }
    }
    out
}

/// Помечает абзацы внутри ячеек таблиц временным атрибутом `mf:cell` (его снимает `xform_elem`).
fn mark_cell_paragraphs(e: &mut Elem, in_cell: bool) {
    for n in e.children.iter_mut() {
        if let Node::Elem(c) = n {
            let inside = in_cell || c.name == "w:tc";
            if inside && c.name == "w:p" {
                c.set_attr("mf:cell", "1");
            }
            mark_cell_paragraphs(c, inside);
        }
    }
}

/// Какие стили используются в содержимом документа (основная часть, колонтитулы, сноски, примечания,
/// нумерация) и есть ли в нём таблицы.
fn used_styles(src: &Src) -> (HashSet<String>, bool) {
    let mut used = HashSet::new();
    let mut tables = false;
    let mut scan = |root: &Elem| {
        xml::walk_ref(root, &mut |e| match e.name.as_str() {
            "w:pStyle" | "w:rStyle" | "w:tblStyle" | "w:numStyleLink" | "w:styleLink" => {
                if let Some(v) = e.val() {
                    used.insert(v.to_string());
                }
            }
            "w:tbl" => tables = true,
            _ => {}
        });
    };
    scan(&src.doc.root);
    let mut parts = vec![];
    for r in src.rels.items.iter().filter(|r| !r.external) {
        if matches!(rel_suffix(&r.ty), "header" | "footer" | "footnotes" | "endnotes" | "comments" | "numbering") {
            if let Some(p) = resolve_part(&src.pkg, &src.main, &r.target) {
                parts.push(p);
            }
        }
    }
    parts.sort();
    parts.dedup();
    for p in parts {
        if let Some(d) = src.pkg.get(&p).and_then(|b| xml::parse(b).ok()) {
            scan(&d.root);
        }
    }
    (used, tables)
}

struct SRec<'a> {
    el: &'a Elem,
    id: String,
    ty: String,
    name: String,
    name_l: String,
    default: bool,
    based: Option<usize>,
    link: Option<usize>,
}

struct BRec {
    id: String,
    ty: String,
    default: bool,
    based: Option<String>,
    sig: String,
    named_sig: String,
    dsig: String,
}

#[derive(Clone, Debug)]
enum Dec {
    /// Использовать стиль базы с таким id.
    Merge(String),
    /// Добавить стиль в базу: новый id и (если стиль переименован) новое имя.
    Import { id: String, name: Option<String> },
    /// Стиль нигде не используется — не переносится.
    Skip,
}

impl Dec {
    fn target(&self) -> Option<&str> {
        match self {
            Dec::Merge(id) | Dec::Import { id, .. } => Some(id),
            Dec::Skip => None,
        }
    }
}

/// Решает судьбу каждого стиля документа-источника.
struct Planner<'a> {
    keep: bool,
    def_para: Option<usize>,
    src: Vec<SRec<'a>>,
    base: Vec<BRec>,
    by_key: HashMap<(String, String), Vec<usize>>,
    base_default: HashMap<String, usize>,
    needed: HashSet<usize>,
    forced: HashSet<usize>,
    dsig: String,
    sigs: Vec<String>,
    named_sigs: Vec<String>,
    decs: Vec<Option<Dec>>,
    busy: Vec<bool>,
    base_ids: HashSet<String>,
    base_names: HashSet<String>,
    src_ids: HashSet<String>,
    src_names: HashSet<String>,
    taken_ids: HashSet<String>,
    taken_names: HashSet<String>,
    ordinal: usize,
    conflicts: Vec<String>,
}

impl<'a> Planner<'a> {
    fn equiv(&self, i: usize, c: usize, parent: &Option<String>) -> bool {
        let b = &self.base[c];
        b.ty == self.src[i].ty && b.dsig == self.dsig && b.sig == self.sigs[i] && b.based == *parent && !self.forced.contains(&i)
    }

    /// Свободный идентификатор стиля. Новый идентификатор не должен совпадать с идентификатором другого стиля
    /// источника: ссылки источника переписываются по карте «старый -> новый» ровно один раз, и новое значение
    /// не должно быть принято за старое (у источника, который сам был результатом слияния, бывают id вида `_m1`).
    fn unique_id(&mut self, id: &str) -> String {
        let mut new = id.to_string();
        let mut k = 1;
        loop {
            let l = new.to_lowercase();
            let other_src = self.src_ids.contains(&l) && new != id;
            if !self.taken_ids.contains(&l) && !other_src {
                break;
            }
            new = format!("{id}_m{k}");
            k += 1;
        }
        self.taken_ids.insert(new.to_lowercase());
        new
    }

    /// То же для стиля, которого нет в источнике (созданного для результата).
    fn unique_new_id(&mut self, id: &str) -> String {
        let mut new = id.to_string();
        let mut k = 1;
        while self.taken_ids.contains(&new.to_lowercase()) || self.src_ids.contains(&new.to_lowercase()) {
            new = format!("{id}_m{k}");
            k += 1;
        }
        self.taken_ids.insert(new.to_lowercase());
        new
    }

    fn fresh_name(&mut self, name: &str) -> String {
        let base = truncate_chars(name, 200).to_string();
        let mut k = self.ordinal.max(2);
        loop {
            let cand = format!("{base} ({k})");
            k += 1;
            // Имя не должно совпасть и с именем другого стиля источника (он переносится под своим именем).
            if !self.src_names.contains(&cand.to_lowercase()) && self.taken_names.insert(cand.to_lowercase()) {
                return cand;
            }
        }
    }

    fn decide(&mut self, i: usize) {
        if self.decs[i].is_some() || self.busy[i] {
            return;
        }
        self.busy[i] = true;
        let parent = self.src[i].based.and_then(|p| {
            self.decide(p);
            self.decs[p].as_ref().and_then(|d| d.target().map(String::from))
        });
        let key = (self.src[i].ty.clone(), self.src[i].name_l.clone());
        let mut cands: Vec<usize> = self.by_key.get(&key).cloned().unwrap_or_default();
        let same_name = !cands.is_empty();
        if !same_name && self.src[i].default {
            cands.extend(self.base_default.get(&key.0).copied());
        }
        let dec = if self.keep {
            if let Some(&c) = cands.iter().find(|&&c| self.equiv(i, c, &parent)) {
                Dec::Merge(self.base[c].id.clone())
            } else if !same_name && !self.taken_names.contains(&self.src[i].name_l) {
                let id = self.unique_id(&self.src[i].id.clone());
                self.taken_names.insert(self.src[i].name_l.clone());
                Dec::Import { id, name: None }
            } else if self.needed.contains(&i) || self.forced.contains(&i) || !same_name {
                // Имя занято (стилем другого типа или копией) — стиль переносится под новым именем.
                let id = self.unique_id(&self.src[i].id.clone());
                let name = self.fresh_name(&self.src[i].name.clone());
                Dec::Import { id, name: Some(name) }
            } else {
                Dec::Skip
            }
        } else if let Some(&c) = cands.first() {
            // Побеждает стиль первого документа.
            if self.base[c].named_sig != self.named_sigs[i] {
                self.conflicts.push(self.src[i].name.clone());
            }
            Dec::Merge(self.base[c].id.clone())
        } else {
            let id = self.unique_id(&self.src[i].id.clone());
            Dec::Import { id, name: None }
        };
        self.decs[i] = Some(dec);
        self.busy[i] = false;
    }

    fn run(&mut self) {
        loop {
            for i in 0..self.src.len() {
                self.decide(i);
            }
            if !self.keep {
                return;
            }
            // Связанные стили (абзаца и знака) переносятся парой: если один добавлен под новым именем,
            // второй тоже, иначе ссылка `link` указывала бы на стиль с другим оформлением.
            let mut grew = false;
            for i in 0..self.src.len() {
                let Some(j) = self.src[i].link else { continue };
                let renamed = |d: &Option<Dec>| matches!(d, Some(Dec::Import { name: Some(_), .. }));
                if renamed(&self.decs[j]) && matches!(self.decs[i], Some(Dec::Merge(_))) && self.forced.insert(i) {
                    grew = true;
                }
            }
            if !grew {
                return;
            }
            // Повторяем с нуля: решения о родителях и детях могли измениться.
            for d in self.decs.iter_mut() {
                *d = None;
            }
            self.taken_ids = self.base_ids.clone();
            self.taken_names = self.base_names.clone();
        }
    }
}

/// Шрифты и цвета темы оформления документа.
#[derive(Clone, Debug, Default)]
struct Theme {
    /// Описание темы для сравнения (без названий).
    sig: String,
    /// "majorHAnsi", "minorEastAsia", ... -> название шрифта
    fonts: HashMap<String, String>,
    /// "dk1", "lt1", "accent1", ... -> цвет RRGGBB
    colors: HashMap<String, String>,
}

fn load_theme(pkg: &Pkg, rels: &Rels, owner: &str) -> Option<Theme> {
    let p = rels.items.iter().filter(|r| !r.external && rel_suffix(&r.ty) == "theme").find_map(|r| resolve_part(pkg, owner, &r.target))?;
    let d = xml::parse(pkg.get(&p)?).ok()?;
    let te = d.root.elems().find(|e| e.local() == "themeElements")?;
    let mut th = Theme::default();
    if let Some(fs) = te.elems().find(|e| e.local() == "fontScheme") {
        for (group, pre) in [("majorFont", "major"), ("minorFont", "minor")] {
            let Some(g) = fs.elems().find(|e| e.local() == group) else { continue };
            for (tag, suffix) in [("latin", "HAnsi"), ("ea", "EastAsia"), ("cs", "Bidi")] {
                if let Some(f) = g.elems().find(|e| e.local() == tag).and_then(|e| e.attr("typeface")) {
                    th.fonts.insert(format!("{pre}{suffix}"), f.to_string());
                }
            }
        }
    }
    if let Some(cs) = te.elems().find(|e| e.local() == "clrScheme") {
        for c in cs.elems() {
            let hex = c.elems().next().and_then(|v| v.attr("val").filter(|_| v.local() == "srgbClr").or_else(|| v.attr("lastClr")));
            if let Some(h) = hex {
                th.colors.insert(c.local().to_string(), h.to_uppercase());
            }
        }
    }
    let mut parts: Vec<String> = th.fonts.iter().map(|(k, v)| format!("f:{k}={v}")).chain(th.colors.iter().map(|(k, v)| format!("c:{k}={v}"))).collect();
    parts.sort();
    th.sig = parts.join(";");
    Some(th)
}

/// Имя цвета темы (Word или DrawingML) -> ключ схемы цветов.
fn theme_color_key(n: &str) -> &str {
    match n {
        "dark1" | "text1" | "tx1" => "dk1",
        "light1" | "background1" | "bg1" => "lt1",
        "dark2" | "text2" | "tx2" => "dk2",
        "light2" | "background2" | "bg2" => "lt2",
        "hyperlink" => "hlink",
        "followedHyperlink" => "folHlink",
        o => o,
    }
}

/// Цвета и шрифты темы заменяются явными значениями (когда тема документа отличается от темы результата).
fn flatten_theme(e: &mut Elem, th: &Theme) {
    match e.name.as_str() {
        "w:rFonts" => {
            for (ta, explicit) in [("w:asciiTheme", "w:ascii"), ("w:hAnsiTheme", "w:hAnsi"), ("w:eastAsiaTheme", "w:eastAsia"), ("w:cstheme", "w:cs")] {
                if let Some(v) = e.remove_attr(ta) {
                    if let Some(f) = th.fonts.get(&v).filter(|f| !f.is_empty()) {
                        e.set_attr(explicit, f.clone());
                    }
                }
            }
        }
        "a:schemeClr" => {
            if let Some(h) = e.attr("val").and_then(|v| th.colors.get(theme_color_key(v))).cloned() {
                e.name = "a:srgbClr".into();
                e.set_attr("val", h);
            }
        }
        _ => {
            let color_target = if e.name == "w:color" { "w:val" } else { "w:color" };
            for (tc, shade, tint, target) in
                [("w:themeColor", "w:themeShade", "w:themeTint", color_target), ("w:themeFill", "w:themeFillShade", "w:themeFillTint", "w:fill")]
            {
                if let Some(n) = e.remove_attr(tc) {
                    e.remove_attr(shade);
                    e.remove_attr(tint);
                    // Word пишет и вычисленный цвет, и ссылку на тему; тема нужна, только если цвета нет.
                    if matches!(e.attr(target), None | Some("auto")) {
                        if let Some(h) = th.colors.get(theme_color_key(&n)) {
                            e.set_attr(target, h.clone());
                        }
                    }
                }
            }
        }
    }
}

/// Порядок дочерних элементов `w:settings` по схеме.
const SETTINGS_ORDER: &[&str] = &[
    "w:writeProtection", "w:view", "w:zoom", "w:removePersonalInformation", "w:removeDateAndTime",
    "w:doNotDisplayPageBoundaries", "w:displayBackgroundShape", "w:printPostScriptOverText",
    "w:printFractionalCharacterWidth", "w:printFormsData", "w:embedTrueTypeFonts", "w:embedSystemFonts",
    "w:saveSubsetFonts", "w:saveFormsData", "w:mirrorMargins", "w:alignBordersAndEdges",
    "w:bordersDoNotSurroundHeader", "w:bordersDoNotSurroundFooter", "w:gutterAtTop", "w:hideSpellingErrors",
    "w:hideGrammaticalErrors", "w:activeWritingStyle", "w:proofState", "w:formsDesign", "w:attachedTemplate",
    "w:linkStyles", "w:stylePaneFormatFilter", "w:stylePaneSortMethod", "w:documentType", "w:mailMerge",
    "w:revisionView", "w:trackRevisions", "w:doNotTrackMoves", "w:doNotTrackFormatting", "w:documentProtection",
    "w:autoFormatOverride", "w:styleLockTheme", "w:styleLockQFSet", "w:defaultTabStop", "w:autoHyphenation",
    "w:consecutiveHyphenLimit", "w:hyphenationZone", "w:doNotHyphenateCaps", "w:showEnvelope", "w:summaryLength",
    "w:clickAndTypeStyle", "w:defaultTableStyle", "w:evenAndOddHeaders", "w:bookFoldRevPrinting",
    "w:bookFoldPrinting", "w:bookFoldPrintingSheets", "w:drawingGridHorizontalSpacing",
    "w:drawingGridVerticalSpacing", "w:displayHorizontalDrawingGridEvery", "w:displayVerticalDrawingGridEvery",
    "w:doNotUseMarginsForDrawingGridOrigin", "w:drawingGridHorizontalOrigin", "w:drawingGridVerticalOrigin",
    "w:doNotShadeFormData", "w:noPunctuationKerning", "w:characterSpacingControl", "w:printTwoOnOne",
    "w:strictFirstAndLastChars", "w:noLineBreaksAfter", "w:noLineBreaksBefore", "w:savePreviewPicture",
    "w:doNotValidateAgainstSchema", "w:saveInvalidXml", "w:ignoreMixedContent", "w:alwaysShowPlaceholderText",
    "w:doNotDemarcateInvalidXml", "w:saveXmlDataOnly", "w:useXSLTWhenSaving", "w:saveThroughXslt",
    "w:showXMLTags", "w:alwaysMergeEmptyNamespace", "w:updateFields", "w:hdrShapeDefaults", "w:footnotePr",
    "w:endnotePr", "w:compat", "w:docVars", "w:rsids", "m:mathPr", "w:attachedSchema", "w:themeFontLang",
    "w:clrSchemeMapping", "w:doNotIncludeSubdocsInStats", "w:doNotAutoCompressPictures", "w:forceUpgrade",
    "w:captions", "w:readModeInkLockDown", "w:smartTagType", "w:schemaLibrary", "w:shapeDefaults",
    "w:doNotEmbedSmartTags", "w:decimalSymbol", "w:listSeparator",
];

/// Включена ли настройка `tag` в корне `settings` (элемент есть и не выключен через `w:val`).
fn setting_on(root: &Elem, tag: &str) -> bool {
    root.child(tag).is_some_and(|e| !matches!(e.val(), Some("0" | "false" | "off")))
}

/// Включены ли у документа разные колонтитулы чётных и нечётных страниц — без разбора всего пакета.
fn peek_even_odd(path: &Path) -> bool {
    let Ok(f) = std::fs::File::open(path) else { return false };
    let Ok(mut zip) = zip::ZipArchive::new(f) else { return false };
    let names: Vec<String> = zip.file_names().filter_map(|n| n.ok().map(|c| c.into_owned())).collect();
    let mut read = |want: &str| -> Option<Vec<u8>> {
        let n = names.iter().find(|n| n.trim_start_matches('/').replace('\\', "/").eq_ignore_ascii_case(want))?;
        let mut buf = vec![];
        zip.by_name(n).ok()?.read_to_end(&mut buf).ok()?;
        Some(buf)
    };
    let main = read("_rels/.rels")
        .and_then(|b| Rels::parse(&b).ok())
        .and_then(|r| {
            r.items
                .iter()
                .find(|x| !x.external && rel_suffix(&x.ty) == "officeDocument")
                .map(|x| normalize(x.target.trim_start_matches('/')))
        })
        .unwrap_or_else(|| "word/document.xml".to_string());
    let Some(rels) = read(&rels_path(&main)).and_then(|b| Rels::parse(&b).ok()) else { return false };
    let Some(target) = rels.items.iter().find(|x| !x.external && rel_suffix(&x.ty) == "settings").map(|x| x.target.clone()) else {
        return false;
    };
    let path = match target.strip_prefix('/') {
        Some(t) => normalize(t),
        None => normalize(&join(dir_of(&main), &target)),
    };
    read(&path).and_then(|b| xml::parse(&b).ok()).is_some_and(|d| setting_on(&d.root, "w:evenAndOddHeaders"))
}

/// Нумерация, которую абзац получает от своего стиля (с учётом родительских стилей).
struct ChainNum {
    num: u64,
    ilvl: Option<String>,
    ind: Option<Vec<(String, String)>>,
    chain: Vec<String>,
}

fn chain_numbering(info: &HashMap<String, StyleLite>, id: &str) -> Option<ChainNum> {
    let mut cur = Some(id.to_string());
    let mut chain: Vec<String> = vec![];
    let (mut num, mut ilvl): (Option<u64>, Option<String>) = (None, None);
    let mut ind: Vec<(String, String)> = vec![];
    while let Some(c) = cur.take() {
        if chain.len() > 32 || chain.contains(&c) {
            break;
        }
        let Some(s) = info.get(&c) else { break };
        chain.push(c);
        if num.is_none() {
            num = s.num;
        }
        if ilvl.is_none() {
            ilvl = s.ilvl.clone();
        }
        for (k, v) in s.ind.iter().flatten() {
            if !ind.iter().any(|(kk, _)| kk == k) {
                ind.push((k.clone(), v.clone()));
            }
        }
        cur = s.based.clone();
    }
    let n = num.filter(|n| *n != 0)?;
    Some(ChainNum { num: n, ilvl, ind: (!ind.is_empty()).then_some(ind), chain })
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
            keep: true,
            appended: 0,
            any_even: false,
            variants: HashMap::new(),
            style_dsig: HashMap::new(),
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
        let mut doc = xml::parse(self.pkg.get(&path).unwrap_or_default()).map_err(|e| part_err(&path, e))?;
        harmonize(&mut doc).map_err(|e| part_err(&path, e))?;
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
            if cx.cell_target.is_some() {
                mark_cell_paragraphs(&mut d.root, false);
            }
            xml::walk(&mut d.root, &mut |e| self.xform_elem(src, cx, &mode, e));
            cx.toc += rename_field_instrs(&mut d.root, &cx.bm_rename, &cx.style_names);
            cx.ranges = saved;
            bytes = xml::to_bytes(&d);
        } else if suffix == "diagramData" {
            bytes = self.patch_diagram_data(src, cx, &bytes);
        }
        if let Some(rb) = src.pkg.get(&rels_path(src_path)) {
            let srels = Rels::parse(rb).map_err(|e| part_err(&rels_path(src_path), e))?;
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
        let Ok(decoded) = xml::decode(bytes) else { return bytes.to_vec() };
        let s = decoded.as_str();
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
        let was_utf8 = std::str::from_utf8(bytes.strip_prefix(&[0xEF, 0xBB, 0xBF][..]).unwrap_or(bytes)).is_ok();
        if !was_utf8 {
            // Часть была в UTF-16: теперь это UTF-8, объявление кодировки должно совпадать.
            let head_end = out.char_indices().nth(200).map_or(out.len(), |(i, _)| i);
            let head = out[..head_end].replace("UTF-16", "UTF-8").replace("utf-16", "UTF-8");
            out = format!("{head}{}", &out[head_end..]);
        }
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
        if let Some(th) = &cx.theme {
            flatten_theme(e, th);
        }
        let name = e.name.clone();
        match name.as_str() {
            "w:pStyle" | "w:rStyle" | "w:tblStyle" | "w:numStyleLink" | "w:styleLink" => {
                if let Some(n) = e.val().and_then(|v| cx.style_map.get(v)) {
                    let n = n.clone();
                    e.set_attr("w:val", n);
                }
            }
            // Родитель, следующий и связанный стили: если стиль не перенесён, ссылка убирается
            // (иначе она указала бы на другой стиль базы с тем же идентификатором).
            "w:basedOn" | "w:next" | "w:link" => match e.val().and_then(|v| cx.style_map.get(v)) {
                Some(n) => {
                    let n = n.clone();
                    e.set_attr("w:val", n);
                }
                None => return false,
            },
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
                    let n = rename_instr(a, &cx.bm_rename, &cx.style_names);
                    if n != a {
                        e.set_attr("w:instr", n);
                    }
                }
            }
            "w:pPr" => fix_ppr_order(e),
            "w:tblPr" => {
                if let Some(id) = &cx.inject_tblstyle {
                    if e.child("w:tblStyle").is_none() {
                        e.children.insert(0, Node::Elem(Elem::new("w:tblStyle").with("w:val", id)));
                    }
                }
            }
            "w:p" => {
                self.restart_style_numbering(cx, e);
                let in_cell = e.remove_attr("mf:cell").is_some();
                let mut done = false;
                if let (true, Some(cell)) = (in_cell, cx.cell_target.clone()) {
                    // В ячейке таблицы «обычный» стиль (явный или неявный) — особая копия.
                    let cur = e.child("w:pPr").and_then(|p| p.child("w:pStyle")).and_then(|s| s.val()).map(String::from);
                    if cur.is_none() || cur == cx.src_default_id {
                        if e.child("w:pPr").is_none() {
                            e.children.insert(0, Node::Elem(Elem::new("w:pPr")));
                        }
                        if let Some(ppr) = e.child_mut("w:pPr") {
                            ppr.children.retain(|n| !matches!(n, Node::Elem(x) if x.name == "w:pStyle"));
                            ppr.children.insert(0, Node::Elem(Elem::new("w:pStyle").with("w:val", &cell)));
                        }
                        done = true;
                    }
                }
                if let (false, Some(id)) = (done, &cx.inject_pstyle) {
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
        let sroot = &sdoc.root;
        let snum = match src.part_path("numbering") {
            Some(p) => num_sigs(&src.load_xml(&p)?.root),
            None => HashMap::new(),
        };
        let bnum = self.parts.get("numbering").map(|p| num_sigs(&p.doc.root)).unwrap_or_default();
        let keep = self.keep;
        let canon_dd = |d: Option<&Elem>| {
            let mut s = String::new();
            if let Some(d) = d {
                canon_xml(d, &[], &mut s);
            }
            s
        };
        let part = self.parts.get("styles").expect("часть стилей загружена");
        let broot = &part.doc.root;
        let (mut sdd, bdd) = (canon_dd(sroot.child("w:docDefaults")), canon_dd(broot.child("w:docDefaults")));
        if let (true, Some(th)) = (cx.theme_diff, &cx.theme) {
            // Тема другая: даже одинаковые стили, пользующиеся её шрифтами и цветами, выглядят иначе.
            sdd.push_str(&format!("|theme:{}", th.sig));
        }
        let d_diff = sdd != bdd;

        // Стили базы.
        let bstyles: Vec<&Elem> = broot.elems().filter(|e| e.name == "w:style" && e.attr("w:styleId").is_some()).collect();
        let bnames: HashMap<String, String> = bstyles
            .iter()
            .map(|e| {
                let id = e.attr("w:styleId").unwrap_or("");
                (id.to_string(), e.child("w:name").and_then(|n| n.val()).unwrap_or(id).to_lowercase())
            })
            .collect();
        let named = |e: &Elem, names: &HashMap<String, String>, sig: &str| {
            let b = e.child("w:basedOn").and_then(|b| b.val()).map(|b| names.get(b).cloned().unwrap_or_else(|| b.to_lowercase()));
            format!("basedOn={};{sig}", b.unwrap_or_default())
        };
        let mut base: Vec<BRec> = vec![];
        let mut by_key: HashMap<(String, String), Vec<usize>> = HashMap::new();
        let mut base_default: HashMap<String, usize> = HashMap::new();
        for (k, e) in bstyles.iter().enumerate() {
            let id = e.attr("w:styleId").unwrap_or("").to_string();
            let ty = e.attr("w:type").unwrap_or("paragraph").to_string();
            let sig = style_own_sig(e, &bnum);
            let dsig = self.style_dsig.get(&id).cloned().unwrap_or_else(|| bdd.clone());
            let name_l = bnames.get(&id).cloned().unwrap_or_else(|| id.to_lowercase());
            by_key.entry((ty.clone(), name_l.clone())).or_default().push(k);
            if flag_on(e, "w:default") {
                base_default.entry(ty.clone()).or_insert(k);
            }
            base.push(BRec {
                named_sig: named(e, &bnames, &sig),
                sig,
                dsig,
                based: e.child("w:basedOn").and_then(|b| b.val()).map(String::from),
                default: flag_on(e, "w:default"),
                ty,
                id,
            });
        }
        // Копии стилей, добавленные из предыдущих документов под другими именами.
        for (key, ids) in &self.variants {
            for id in ids {
                if let Some(k) = base.iter().position(|b| b.id == *id) {
                    by_key.entry(key.clone()).or_default().push(k);
                }
            }
        }

        // Стили источника.
        let sstyles: Vec<&Elem> = sroot.elems().filter(|e| e.name == "w:style" && e.attr("w:styleId").is_some()).collect();
        let snames: HashMap<String, String> = sstyles
            .iter()
            .map(|e| {
                let id = e.attr("w:styleId").unwrap_or("");
                (id.to_string(), e.child("w:name").and_then(|n| n.val()).unwrap_or(id).to_string())
            })
            .collect();
        let snames_l: HashMap<String, String> = snames.iter().map(|(k, v)| (k.clone(), v.to_lowercase())).collect();
        let sidx: HashMap<String, usize> = sstyles.iter().enumerate().map(|(i, e)| (e.attr("w:styleId").unwrap_or("").to_string(), i)).collect();
        let mut srec: Vec<SRec> = vec![];
        for e in &sstyles {
            let id = e.attr("w:styleId").unwrap_or("").to_string();
            let name = snames.get(&id).cloned().unwrap_or_else(|| id.clone());
            let target = |tag: &str| e.child(tag).and_then(|b| b.val()).and_then(|v| sidx.get(v)).copied();
            srec.push(SRec {
                el: e,
                ty: e.attr("w:type").unwrap_or("paragraph").to_string(),
                name_l: name.to_lowercase(),
                default: flag_on(e, "w:default"),
                based: target("w:basedOn"),
                link: target("w:link"),
                name,
                id,
            });
        }
        // Корневые стили абзацев получают docDefaults источника, если они отличаются от базовых.
        let push = keep && d_diff;
        let mut pushed: HashMap<usize, Elem> = HashMap::new();
        let mut sigs = vec![];
        let mut named_sigs = vec![];
        for (i, r) in srec.iter().enumerate() {
            let own = style_own_sig(r.el, &snum);
            named_sigs.push(named(r.el, &snames_l, &own));
            if push && r.ty == "paragraph" && r.based.is_none() {
                let mut c = r.el.clone();
                push_defaults(&mut c, sroot.child("w:docDefaults"), broot.child("w:docDefaults"), None);
                sigs.push(style_own_sig(&c, &snum));
                pushed.insert(i, c);
            } else {
                sigs.push(own);
            }
        }
        // Стиль абзаца «по умолчанию» источника. LibreOffice не помечает «Обычный» атрибутом `w:default`,
        // но применяет его к абзацам без стиля, поэтому без пометки это стиль с именем «Normal».
        let def_para = srec
            .iter()
            .position(|r| r.ty == "paragraph" && r.default)
            .or_else(|| srec.iter().position(|r| r.ty == "paragraph" && r.name_l == "normal"));
        // Какие стили нужны: используемые в содержимом, их родители и связанные стили.
        let (used, has_tables) = used_styles(src);
        let mut needed: HashSet<usize> = HashSet::new();
        let mut stack: Vec<usize> = vec![];
        for (i, r) in srec.iter().enumerate() {
            if used.contains(&r.id) || Some(i) == def_para || (r.default && r.ty == "table" && has_tables) {
                stack.push(i);
            }
        }
        while let Some(i) = stack.pop() {
            if needed.insert(i) {
                stack.extend(srec[i].based);
                stack.extend(srec[i].link);
                stack.extend(srec.iter().enumerate().filter(|(_, r)| r.link == Some(i)).map(|(j, _)| j));
            }
        }
        let base_ids: HashSet<String> = base.iter().map(|b| b.id.to_lowercase()).collect();
        let base_names: HashSet<String> = bnames.values().cloned().collect();
        let n = srec.len();
        let srec_ids: HashSet<String> = srec.iter().map(|r| r.id.to_lowercase()).collect();
        let srec_names: HashSet<String> = srec.iter().map(|r| r.name_l.clone()).collect();
        let mut pl = Planner {
            keep,
            src: srec,
            def_para,
            base,
            by_key,
            base_default,
            needed,
            forced: HashSet::new(),
            dsig: sdd.clone(),
            sigs,
            named_sigs,
            decs: vec![None; n],
            busy: vec![false; n],
            taken_ids: base_ids.clone(),
            taken_names: base_names.clone(),
            src_ids: srec_ids,
            src_names: srec_names,
            base_ids,
            base_names,
            ordinal: cx.ordinal,
            conflicts: vec![],
        };
        pl.run();

        // Что делать с каждым стилем: карта идентификаторов, список добавляемых стилей.
        let mut default_types: HashSet<String> = pl.base.iter().filter(|b| b.default).map(|b| b.ty.clone()).collect();
        let mut imported: Vec<Elem> = vec![];
        let mut registered: Vec<(Option<(String, String)>, String)> = vec![];
        for (i, r) in pl.src.iter().enumerate() {
            let Some(dec) = &pl.decs[i] else { continue };
            if let Some(t) = dec.target() {
                cx.style_map.insert(r.id.clone(), t.to_string());
            }
            let Dec::Import { id, name } = dec else { continue };
            let mut c = pushed.remove(&i).unwrap_or_else(|| r.el.clone());
            c.set_attr("w:styleId", id.clone());
            if let Some(nn) = name {
                match c.child_mut("w:name") {
                    Some(e) => e.set_attr("w:val", nn.clone()),
                    None => insert_ordered(&mut c, Elem::new("w:name").with("w:val", nn), STYLE_ORDER),
                }
                c.set_attr("w:customStyle", "1");
                c.remove_attr("w:default");
                c.children.retain(|n| !matches!(n, Node::Elem(e) if e.name == "w:aliases" || e.name == "w:qFormat"));
                cx.style_names.insert(r.name_l.clone(), nn.clone());
                registered.push((Some((r.ty.clone(), r.name_l.clone())), id.clone()));
            } else {
                if default_types.contains(&r.ty) {
                    c.remove_attr("w:default");
                } else if flag_on(&c, "w:default") {
                    default_types.insert(r.ty.clone());
                }
                registered.push((None, id.clone()));
            }
            imported.push(c);
        }
        // Нужен ли отдельный «обычный» стиль, если у источника его нет, а docDefaults отличаются.
        let src_default_para = pl.def_para;
        let src_default_tbl = pl.src.iter().position(|r| r.ty == "table" && r.default);
        // У источника нет стиля абзаца «по умолчанию» (абзацы без стиля получают только docDefaults), а у результата
        // он есть или docDefaults другие: создаётся пустой стиль с настройками docDefaults источника.
        if keep && src_default_para.is_none() && (push || pl.base_default.contains_key("paragraph")) {
            let id = pl.unique_new_id("Normal");
            let name = pl.fresh_name("Normal");
            let mut c = Elem::new("w:style")
                .with("w:type", "paragraph")
                .with("w:customStyle", "1")
                .with("w:styleId", &id)
                .with_child(Elem::new("w:name").with("w:val", &name));
            if push {
                push_defaults(&mut c, sroot.child("w:docDefaults"), broot.child("w:docDefaults"), None);
            }
            registered.push((None, id.clone()));
            imported.push(c);
            cx.default_target = Some(id);
        }
        // У абзацев в таблицах свойства стиля таблицы сильнее docDefaults, но слабее явных свойств стиля абзаца.
        // Перенос docDefaults в «обычный» стиль перекрыл бы стиль таблицы, поэтому для ячеек делается копия
        // «обычного» стиля без свойств, которые задают стили таблиц.
        if push && has_tables {
            let tables: Vec<&Elem> = pl.src.iter().filter(|r| r.ty == "table" && (used.contains(&r.id) || r.default)).map(|r| r.el).collect();
            let skip = table_ppr_props(&tables);
            let root_default = match src_default_para {
                Some(i) => pl.src[i].based.is_none().then(|| (Some(pl.src[i].el.clone()), pl.src[i].id.clone(), pl.src[i].name.clone())),
                None => Some((None, "Normal".to_string(), "Normal".to_string())),
            };
            if let (false, Some((el, id0, name0))) = (skip.is_empty(), root_default) {
                let id = pl.unique_new_id(&format!("{id0}Cell"));
                let name = pl.fresh_name(&format!("{name0} в таблице"));
                let mut c = el.unwrap_or_else(|| Elem::new("w:style").with("w:type", "paragraph").with_child(Elem::new("w:name").with("w:val", &name0)));
                c.set_attr("w:styleId", id.clone());
                c.set_attr("w:customStyle", "1");
                c.remove_attr("w:default");
                match c.child_mut("w:name") {
                    Some(e) => e.set_attr("w:val", name),
                    None => insert_ordered(&mut c, Elem::new("w:name").with("w:val", &name), STYLE_ORDER),
                }
                c.children.retain(|n| !matches!(n, Node::Elem(e) if e.name == "w:aliases" || e.name == "w:qFormat"));
                push_defaults(&mut c, sroot.child("w:docDefaults"), broot.child("w:docDefaults"), Some(&skip));
                registered.push((None, id.clone()));
                imported.push(c);
                cx.cell_target = Some(id);
                cx.src_default_id = src_default_para.map(|i| pl.src[i].id.clone());
            }
        }
        let conflicts = std::mem::take(&mut pl.conflicts);
        let tbl_target = src_default_tbl.and_then(|i| pl.decs[i].as_ref().and_then(|d| d.target().map(String::from)));
        if let Some(i) = src_default_para {
            cx.default_target = pl.decs[i].as_ref().and_then(|d| d.target().map(String::from));
        }
        drop(pl);
        for (key, id) in registered {
            self.style_dsig.insert(id.clone(), sdd.clone());
            if let Some(k) = key {
                self.variants.entry(k).or_default().push(id);
            }
        }

        // Ссылки внутри добавляемых определений (basedOn/next/link, нумерация) переписываются как в содержимом.
        let tmode = Mode { rels: None };
        let inject_saved = cx.inject_pstyle.take();
        for c in imported.iter_mut() {
            xml::walk(c, &mut |e| self.xform_elem(src, cx, &tmode, e));
        }
        cx.inject_pstyle = inject_saved;
        let part = self.parts.get_mut("styles").expect("часть стилей загружена");
        for c in imported {
            part.doc.root.push(c);
        }
        merge_root_decls(&mut part.doc.root, sroot)?;
        part.dirty = true;

        // Стиль абзаца «по умолчанию» исходного документа: абзацы без pStyle получают его явно.
        let find_default = |ty: &str| {
            part.doc
                .root
                .elems()
                .find(|e| e.name == "w:style" && e.attr("w:type").unwrap_or("paragraph") == ty && flag_on(e, "w:default"))
                .and_then(|e| e.attr("w:styleId"))
                .map(String::from)
        };
        let (bd, bt) = (find_default("paragraph"), find_default("table"));
        if let Some(t) = &cx.default_target {
            if Some(t) != bd.as_ref() {
                cx.inject_pstyle = Some(t.clone());
            }
        }
        if has_tables {
            if let Some(t) = tbl_target.filter(|t| Some(t) != bt.as_ref()) {
                cx.inject_tblstyle = Some(t);
            }
        }
        // Описание цепочек стилей для проверки нумерации абзацев.
        cx.style_info = part
            .doc
            .root
            .elems()
            .filter(|e| e.name == "w:style")
            .filter_map(|e| {
                let id = e.attr("w:styleId")?.to_string();
                let ppr = e.child("w:pPr");
                let numpr = ppr.and_then(|p| p.child("w:numPr"));
                let ind = ppr.and_then(|p| p.child("w:ind")).map(|i| i.attrs.clone());
                Some((
                    id,
                    StyleLite {
                        based: e.child("w:basedOn").and_then(|b| b.val()).map(String::from),
                        num: numpr.and_then(|n| n.child("w:numId")).and_then(|n| n.val()).and_then(|v| v.trim().parse().ok()),
                        ilvl: numpr.and_then(|n| n.child("w:ilvl")).and_then(|n| n.val()).map(String::from),
                        ind,
                    },
                ))
            })
            .collect();

        let f = cx.file.clone();
        if !keep {
            let mut conflicts = conflicts;
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
            if d_diff {
                self.warn(format!(
                    "«{f}»: шрифт, размер и интервалы «по умолчанию» отличаются от первого документа — текст без явного оформления может выглядеть иначе."
                ));
            }
        }
        Ok(())
    }

    // ---------- нумерация, заданная стилем ----------

    /// Абзацы, нумерация которых задана стилем, общим с другим документом, получают собственный список
    /// (копия списка с «начать с»), чтобы нумерация не продолжалась из предыдущих документов.
    fn restart_style_numbering(&mut self, cx: &mut Ctx, e: &mut Elem) {
        if cx.style_info.is_empty() || e.child("w:pPr").is_some_and(|p| p.child("w:numPr").is_some()) {
            return;
        }
        let target = match e.child("w:pPr").and_then(|p| p.child("w:pStyle")).and_then(|s| s.val()) {
            Some(v) => cx.style_map.get(v).cloned(),
            None => cx.default_target.clone(),
        };
        let Some(cn) = target.and_then(|t| chain_numbering(&cx.style_info, &t)) else { return };
        // Списки самого документа (со сдвигом номеров) уже независимы.
        if cn.num > cx.num_off {
            return;
        }
        let Some((new_num, ilvl)) = self.restart_num(cx, cn.num, &cn.chain, cn.ilvl.clone()) else { return };
        let mut numpr = Elem::new("w:numPr");
        if let Some(l) = ilvl {
            numpr.push(Elem::new("w:ilvl").with("w:val", &l));
        }
        // Номера списков при переносе сдвигаются на `num_off` (см. xform_elem), поэтому пишем со смещением.
        numpr.push(Elem::new("w:numId").with("w:val", &new_num.saturating_sub(cx.num_off).to_string()));
        if e.child("w:pPr").is_none() {
            e.children.insert(0, Node::Elem(Elem::new("w:pPr")));
        }
        let Some(ppr) = e.child_mut("w:pPr") else { return };
        insert_ordered(ppr, numpr, PPR_ORDER);
        // Отступы из стиля сохраняются явно: у абзаца с собственной нумерацией отступ уровня списка сильнее стиля.
        if let (Some(ind), None) = (cn.ind, ppr.child("w:ind")) {
            let mut i = Elem::new("w:ind");
            i.attrs = ind;
            insert_ordered(ppr, i, PPR_ORDER);
        }
    }

    /// Новый список на основе списка `n` результата: те же уровни, но нумерация начинается заново.
    fn restart_num(&mut self, cx: &mut Ctx, n: u64, chain: &[String], ilvl: Option<String>) -> Option<(u64, Option<String>)> {
        let part = self.parts.get("numbering")?;
        let root = &part.doc.root;
        let parse = |v: Option<&str>| v.and_then(|v| v.trim().parse::<u64>().ok());
        let num_el = root.elems().find(|e| e.name == "w:num" && parse(e.attr("w:numId")) == Some(n))?;
        let aid = num_el.child("w:abstractNumId")?.val()?.to_string();
        let abs = root.elems().find(|e| e.name == "w:abstractNum" && e.attr("w:abstractNumId") == Some(aid.as_str()))?;
        // Уровень, привязанный к стилю абзаца, если стиль его не указывает.
        let ilvl = ilvl.or_else(|| {
            abs.elems()
                .filter(|l| l.name == "w:lvl")
                .find(|l| l.child("w:pStyle").and_then(|p| p.val()).is_some_and(|p| chain.iter().any(|c| c == p)))
                .and_then(|l| l.attr("w:ilvl").map(String::from))
        });
        if let Some(r) = cx.restart.get(&n) {
            return r.map(|id| (id, ilvl));
        }
        let lvls: Vec<&Elem> = abs.elems().filter(|l| l.name == "w:lvl").collect();
        // Маркированный список не нумерует — перезапуск не нужен.
        let counts = lvls.iter().any(|l| !matches!(l.child("w:numFmt").and_then(|f| f.val()), Some("bullet" | "none")));
        if !counts {
            cx.restart.insert(n, None);
            return None;
        }
        if cx.next_num == 0 {
            let max = root.elems().filter(|e| e.name == "w:num").filter_map(|e| parse(e.attr("w:numId"))).max().unwrap_or(0);
            cx.next_num = max + 1;
        }
        let id = cx.next_num;
        cx.next_num += 1;
        let mut nn = Elem::new("w:num").with("w:numId", &id.to_string()).with_child(Elem::new("w:abstractNumId").with("w:val", &aid));
        let own: Vec<&Elem> = num_el.elems().filter(|o| o.name == "w:lvlOverride").collect();
        let mut done: HashSet<String> = HashSet::new();
        for l in &lvls {
            let Some(k) = l.attr("w:ilvl") else { continue };
            done.insert(k.to_string());
            let start = |lv: &Elem| lv.child("w:start").and_then(|s| s.val()).unwrap_or("1").to_string();
            match own.iter().find(|o| o.attr("w:ilvl") == Some(k)) {
                Some(o) => {
                    let mut o = (*o).clone();
                    if o.child("w:startOverride").is_none() {
                        let st = o.child("w:lvl").map(start).unwrap_or_else(|| start(l));
                        o.children.insert(0, Node::Elem(Elem::new("w:startOverride").with("w:val", &st)));
                    }
                    nn.push(o);
                }
                None => nn.push(
                    Elem::new("w:lvlOverride").with("w:ilvl", k).with_child(Elem::new("w:startOverride").with("w:val", &start(l))),
                ),
            }
        }
        for o in own {
            if o.attr("w:ilvl").is_some_and(|k| !done.contains(k)) {
                nn.push(o.clone());
            }
        }
        let part = self.parts.get_mut("numbering")?;
        let at = part
            .doc
            .root
            .children
            .iter()
            .rposition(|c| matches!(c, Node::Elem(e) if e.name == "w:num"))
            .or_else(|| part.doc.root.children.iter().rposition(|c| matches!(c, Node::Elem(e) if e.name == "w:abstractNum")))
            .map_or(part.doc.root.children.len(), |i| i + 1);
        part.doc.root.children.insert(at, Node::Elem(nn));
        part.dirty = true;
        cx.restart.insert(n, Some(id));
        Some((id, ilvl))
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
                cx.toc += rename_field_instrs(&mut c, &cx.bm_rename, &cx.style_names);
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
            // Страница с особым колонтитулом первой страницы без него остаётся пустой (а не берёт чужой).
            let title_pg = s.child("w:titlePg").is_some_and(|t| !matches!(t.val(), Some("0" | "false" | "off")));
            let mut add = vec![];
            for footer in [false, true] {
                for t in HF_TYPES {
                    let key = (footer, t.to_string());
                    if (self.hf_state.contains(&key) || (title_pg && t == "first")) && !defined.contains(&key) {
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

    // ---------- колонтитулы чётных и нечётных страниц ----------

    /// Настройка «разные колонтитулы для чётных и нечётных страниц» действует на весь файл. Если она нужна
    /// хотя бы одному документу, она включается, а документам без неё на чётные страницы копируется
    /// колонтитул «по умолчанию» (тот же объект): их страницы выглядят как раньше.
    fn enable_even_odd(&mut self) -> Result<()> {
        if !self.any_even {
            return Ok(());
        }
        if !self.load_part("settings")? {
            let root = Elem::new("w:settings").with("xmlns:w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
            self.create_part("settings", root);
        }
        let part = self.parts.get_mut("settings").expect("настройки загружены");
        if setting_on(&part.doc.root, "w:evenAndOddHeaders") {
            return Ok(());
        }
        part.doc.root.children.retain(|n| !matches!(n, Node::Elem(e) if e.name == "w:evenAndOddHeaders"));
        insert_ordered(&mut part.doc.root, Elem::new("w:evenAndOddHeaders"), SETTINGS_ORDER);
        part.dirty = true;
        // Разделы первого документа.
        let mut added = add_even_refs(&mut self.final_sect);
        for n in self.body.iter_mut() {
            if let Some(s) = para_sect_mut(n) {
                added |= add_even_refs(s);
            }
        }
        if added {
            self.hf_state.insert((false, "even".into()));
            self.hf_state.insert((true, "even".into()));
        }
        Ok(())
    }

    // ---------- проверки совместимости ----------

    fn compat_checks(&mut self, src: &Src, cx: &mut Ctx) -> Result<()> {
        let f = cx.file.clone();
        // Настройки документа, которые действуют на весь файл.
        let flags = [("w:mirrorMargins", "зеркальные поля"), ("w:gutterAtTop", "поле переплёта сверху")];
        let base_has = |this: &Self, tag: &str| this.parts.get("settings").is_some_and(|p| p.doc.root.child(tag).is_some());
        self.load_part("settings")?;
        let mut src_bg_shape = false;
        let mut src_settings: Option<Elem> = None;
        if let Some(sp) = src.part_path("settings") {
            let sdoc = src.load_xml(&sp)?;
            cx.src_even = setting_on(&sdoc.root, "w:evenAndOddHeaders");
            src_bg_shape = setting_on(&sdoc.root, "w:displayBackgroundShape");
            src_settings = Some(sdoc.root.clone());
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
        // Режим совместимости (параметры вёрстки) общий для всего файла: слегка влияет на интервалы и разбивку.
        let compat_of = |root: Option<&Elem>| -> (String, Vec<String>) {
            let Some(c) = root.and_then(|r| r.child("w:compat")) else { return ("12".into(), vec![]) };
            let mode = c
                .elems()
                .find(|e| e.name == "w:compatSetting" && e.attr("w:name") == Some("compatibilityMode"))
                .and_then(|e| e.attr("w:val"))
                .unwrap_or("12")
                .to_string();
            let mut flags: Vec<String> = c.elems().filter(|e| e.name != "w:compatSetting" && !matches!(e.val(), Some("0" | "false" | "off"))).map(|e| e.name.clone()).collect();
            flags.sort();
            (mode, flags)
        };
        let base_compat = compat_of(self.parts.get("settings").map(|p| &p.doc.root));
        if compat_of(src_settings.as_ref()) != base_compat {
            self.warn(format!(
                "«{f}»: режим совместимости документа (параметры вёрстки) отличается от первого документа; используется режим первого документа — интервалы и разбивка на страницы могут немного измениться."
            ));
        }
        // Фон страницы общий для всего файла: используется фон первого документа.
        let bg_sig = |root: &Elem, shape: bool| {
            let mut s = String::new();
            if let Some(b) = root.child("w:background") {
                canon_xml(b, &[], &mut s);
            }
            (s, shape)
        };
        let base_shape = self.parts.get("settings").is_some_and(|p| setting_on(&p.doc.root, "w:displayBackgroundShape"));
        let (sb, bb) = (bg_sig(&src.doc.root, src_bg_shape), bg_sig(&self.doc.root, base_shape));
        if sb.0 != bb.0 {
            self.warn(if sb.0.is_empty() {
                format!("«{f}»: у первого документа задан фон страницы — он будет применён и к страницам этого документа (в Word фон общий для всего файла).")
            } else if bb.0.is_empty() {
                format!("«{f}»: фон страницы этого документа не перенесён (в Word фон общий для всего файла, у первого документа фона нет).")
            } else {
                format!("«{f}»: фон страницы этого документа отличается от первого документа; используется фон первого документа (в Word фон общий для всего файла).")
            });
        }
        // Встроенные (внедрённые) шрифты не переносятся.
        let embedded = src
            .rels
            .items
            .iter()
            .filter(|r| !r.external && rel_suffix(&r.ty) == "fontTable")
            .filter_map(|r| resolve_part(&src.pkg, &src.main, &r.target))
            .any(|p| src.pkg.get(&rels_path(&p)).and_then(|b| Rels::parse(b).ok()).is_some_and(|r| r.items.iter().any(|x| rel_suffix(&x.ty) == "font")));
        if embedded {
            self.warn(format!(
                "«{f}»: встроенные в документ шрифты не переносятся; если этих шрифтов нет на компьютере, они будут заменены."
            ));
        }
        // Тема оформления (цвета и шрифты темы).
        let uses_theme = |pkg: &Pkg, path: Option<String>| {
            path.and_then(|p| pkg.get(&p)).and_then(|b| xml::decode(b).ok()).is_some_and(|s| {
                ["Theme=\"", "themeColor=\"", "themeFill=\"", "themeShade=\"", "themeTint=\"", "schemeClr"].iter().any(|k| s.contains(k))
            })
        };
        let mut parts = vec![Some(src.main.clone()), src.part_path("styles"), src.part_path("numbering")];
        for r in src.rels.items.iter().filter(|r| !r.external && matches!(rel_suffix(&r.ty), "header" | "footer" | "footnotes" | "endnotes")) {
            parts.push(resolve_part(&src.pkg, &src.main, &r.target));
        }
        if parts.into_iter().any(|p| uses_theme(&src.pkg, p)) {
            let main = self.main.clone();
            let brels = self.rels_mut(&main).clone();
            let a = load_theme(&src.pkg, &src.rels, &src.main);
            let b = load_theme(&self.pkg, &brels, &main);
            if let Some(a) = a {
                if b.as_ref().is_none_or(|b| b.sig != a.sig) {
                    cx.theme_diff = true;
                    if self.keep {
                        // Цвета и шрифты темы заменяются явными значениями — вид текста сохраняется.
                        cx.theme = Some(a);
                        let objects = src.rels.items.iter().any(|r| matches!(rel_suffix(&r.ty), "chart" | "diagramData" | "diagramColors"));
                        if objects {
                            self.warn(format!(
                                "«{f}»: тема оформления отличается от первого документа; цвета и шрифты темы заменены явными значениями, кроме диаграмм и SmartArt — их цвета могут отличаться."
                            ));
                        }
                    } else {
                        self.warn(format!(
                            "«{f}»: тема оформления (цвета и шрифты темы) отличается от первого документа; текст, оформленный цветами и шрифтами темы, может выглядеть иначе."
                        ));
                    }
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
        self.appended += 1;
        let mut cx = Ctx { file: src.name.clone(), ordinal: self.appended + 1, ..Default::default() };
        merge_root_decls(&mut self.doc.root, &src.doc.root)?;
        self.compat_checks(&src, &mut cx)?;
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
        if cx.cell_target.is_some() {
            mark_cell_paragraphs(&mut tmp, false);
        }
        xml::walk(&mut tmp, &mut |e| self.xform_elem(&src, &mut cx, &mode, e));
        cx.toc += rename_field_instrs(&mut tmp, &cx.bm_rename, &cx.style_names);
        self.import_comment_parts(&src, &mut cx)?;

        let Some(Node::Elem(f)) = tmp.children.pop() else { bail!("внутренняя ошибка: потерян раздел документа") };
        fin = f;
        let mut nodes = tmp.children;
        // Настройка чётных/нечётных колонтитулов общая: документ без неё получает колонтитулы и на чётных страницах.
        if self.any_even && !cx.src_even {
            for n in nodes.iter_mut() {
                if let Some(s) = para_sect_mut(n) {
                    add_even_refs(s);
                }
            }
            add_even_refs(&mut fin);
        }
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
        if cx.toc > 0 {
            self.warn(format!(
                "«{f}»: в документе есть оглавление или указатель — номера страниц в них остались прежними; после объединения обновите поля в Word (Ctrl+A, затем F9)."
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

/// Параметры объединения документов.
#[derive(Clone, Debug)]
pub struct Options {
    /// Сохранять оформление каждого документа: стиль следующего документа с тем же именем, но другим
    /// определением, добавляется под новым именем («Заголовок 1 (2)»), а не заменяется стилем первого.
    /// Выключено — как при вставке файла в Word: побеждают стили первого документа.
    pub keep_source_styles: bool,
}

impl Default for Options {
    fn default() -> Self {
        Options { keep_source_styles: true }
    }
}

/// Объединяет документы `.docx` по порядку; возвращает предупреждения для пользователя.
pub fn merge(inputs: &[PathBuf], out: &Path, opts: &Options) -> Result<Vec<String>> {
    let (first, rest) = inputs.split_first().ok_or_else(|| anyhow!("нет файлов"))?;
    // Ко всем ошибкам добавляется имя файла, к которому они относятся.
    let of = |p: &Path, e: anyhow::Error| {
        anyhow!("«{}»: {e:#}", p.file_name().map(|n| n.to_string_lossy().into_owned()).unwrap_or_default())
    };
    let base = Src::load(first).map_err(|e| of(first, e))?;
    let mut m = Merger::new(base).map_err(|e| of(first, e))?;
    m.keep = opts.keep_source_styles;
    m.any_even = inputs.iter().any(|p| peek_even_odd(p));
    m.enable_even_odd()?;
    for p in rest {
        let src = Src::load(p).map_err(|e| of(p, e))?;
        m.append(src).map_err(|e| of(p, e))?;
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
        run_opts(dir, name, docs, &Options::default())
    }

    /// Режим «как в Word»: побеждают стили первого документа.
    fn run_base_wins(dir: &Path, name: &str, docs: &[&Fx]) -> (PathBuf, Vec<String>) {
        run_opts(dir, name, docs, &Options { keep_source_styles: false })
    }

    fn run_opts(dir: &Path, name: &str, docs: &[&Fx], opts: &Options) -> (PathBuf, Vec<String>) {
        let mut paths = vec![];
        for (i, d) in docs.iter().enumerate() {
            let p = dir.join(format!("{name}_in{i}.docx"));
            d.write(&p);
            paths.push(p);
        }
        let out = dir.join(format!("{name}_out.docx"));
        let w = merge(&paths, &out, opts).unwrap();
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
        // стили: уникальные id и имена, все ссылки указывают на существующие стили
        if let Some(st) = docs.get("word/styles.xml") {
            let styles: Vec<&Elem> = st.root.elems().filter(|e| e.name == "w:style").collect();
            let ids: Vec<String> = styles.iter().filter_map(|e| e.attr("w:styleId").map(String::from)).collect();
            errs.extend(dups("styleId", ids.iter().map(|i| i.to_lowercase()).collect()));
            errs.extend(dups("имя стиля", styles.iter().filter_map(|e| e.child("w:name").and_then(|n| n.val()).map(|n| n.to_lowercase())).collect()));
            for e in &styles {
                for tag in ["w:basedOn", "w:next", "w:link"] {
                    if let Some(v) = e.child(tag).and_then(|c| c.val()) {
                        if !ids.iter().any(|i| i == v) {
                            errs.push(format!("стиль {:?}: {tag} -> {v} не существует", e.attr("w:styleId")));
                        }
                    }
                }
            }
            for r in &stories {
                for tag in ["w:pStyle", "w:rStyle", "w:tblStyle"] {
                    for e in all(r, tag) {
                        let v = e.val().unwrap_or("");
                        if !ids.iter().any(|i| i == v) {
                            errs.push(format!("{tag} {v} без стиля"));
                        }
                    }
                }
            }
        }
        // списки: используемые номера существуют
        let num_ids: Vec<String> = docs
            .get("word/numbering.xml")
            .map(|n| all(&n.root, "w:num").iter().filter_map(|e| e.attr("w:numId").map(String::from)).collect())
            .unwrap_or_default();
        let mut num_roots: Vec<&Elem> = stories.clone();
        if let Some(st) = docs.get("word/styles.xml") {
            num_roots.push(&st.root);
        }
        for r in num_roots {
            for e in all(r, "w:numId") {
                let v = e.val().unwrap_or("");
                if v != "0" && !num_ids.iter().any(|i| i == v) {
                    errs.push(format!("numId {v} не определён"));
                }
            }
        }
        // колонтитулы одного раздела: типы не повторяются
        for sp in all(&main.root, "w:sectPr") {
            let mut seen = HashSet::new();
            for r in sp.elems().filter(|e| e.name == "w:headerReference" || e.name == "w:footerReference") {
                if !seen.insert((r.name.clone(), r.attr("w:type").unwrap_or("default").to_string())) {
                    errs.push("повторная ссылка на колонтитул одного типа".into());
                }
            }
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
        let (out, w) = run_base_wins(dir.path(), "styles", &[&a, &b]);
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
        let (_, w) = run_base_wins(dir.path(), "dd", &[&a, &b]);
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
        let b = base("<w:mirrorMargins/><w:evenAndOddHeaders/>", "00FF00")
            .part("word/fontTable.xml", r#"<w:fonts xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>"#)
            .part("word/_rels/fontTable.xml.rels", format!(r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="{RT}font" Target="fonts/font1.odttf"/></Relationships>"#))
            .part("word/fonts/font1.odttf", b"F")
            .rel("rFt", "fontTable", "fontTable.xml")
            .ct("/word/fontTable.xml", &format!("{CT_NOTES}fontTable+xml"))
            .ct("/word/fonts/font1.odttf", "application/vnd.openxmlformats-officedocument.obfuscatedFont");
        let (out, w) = run(dir.path(), "compat", &[&a, &b]);
        assert!(w.iter().any(|m| m.contains("зеркальные поля")), "{w:?}");
        // чётные/нечётные колонтитулы — не потеря: настройка включается для всего файла, предупреждения нет
        assert!(!w.iter().any(|m| m.contains("чётных")), "{w:?}");
        let pkg = Pkg::read(&out).unwrap();
        let settings = xml::parse(pkg.get("word/settings.xml").unwrap()).unwrap();
        assert!(setting_on(&settings.root, "w:evenAndOddHeaders"));
        assert!(w.iter().any(|m| m.contains("встроенные")), "{w:?}");
        // тема отличается: цвета и шрифты темы второго документа заменены явными значениями, предупреждения нет
        assert!(!w.iter().any(|m| m.contains("тема оформления")), "{w:?}");
        let d = doc_of(&out).1;
        let colors = all(&d.root, "w:color");
        assert_eq!(colors.len(), 2);
        assert!(colors[0].attr("w:themeColor").is_some(), "первый документ не меняется");
        assert!(colors[1].attr("w:themeColor").is_none() && colors[1].val() == Some("000000"));
        // в режиме «как в Word» тема первого документа — с предупреждением
        let (_, w) = run_base_wins(dir.path(), "compat_wins", &[&a, &b]);
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

    // ───────── оформление каждого документа, нумерация стилей, поля, колонтитулы, ошибки ─────────

    const WNS: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    fn numbering_xml(fmt: &str) -> String {
        format!(
            r#"<?xml version="1.0"?><w:numbering xmlns:w="{WNS}"><w:abstractNum w:abstractNumId="0"><w:nsid w:val="AAAAAAAA"/><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="{fmt}"/><w:lvlText w:val="%1."/></w:lvl></w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num></w:numbering>"#
        )
    }

    fn style_list(pkg: &Pkg) -> Vec<Elem> {
        let st = xml::parse(pkg.get("word/styles.xml").unwrap()).unwrap();
        st.root.elems().filter(|e| e.name == "w:style").cloned().collect()
    }

    fn style_name(s: &Elem) -> String {
        s.child("w:name").and_then(|n| n.val()).unwrap_or("").to_string()
    }

    fn pstyles(d: &xml::Doc) -> Vec<String> {
        all(&d.root, "w:pStyle").iter().map(|e| e.val().unwrap().to_string()).collect()
    }

    /// Документ со стилем «List Number» (нумерация в стиле) и тремя пунктами.
    fn numbered_doc(fmt: &str, items: &str) -> Fx {
        let mut f = fx(&format!(
            r#"{}{}"#,
            ["1", "2"].iter().map(|_| format!(r#"<w:p><w:pPr><w:pStyle w:val="ListNumber"/></w:pPr><w:r><w:t>{items}</w:t></w:r></w:p>"#)).collect::<String>(),
            sect("")
        ));
        f.styles = styles_xml(
            r#"<w:style w:type="paragraph" w:styleId="ListNumber"><w:name w:val="List Number"/><w:basedOn w:val="Normal"/><w:pPr><w:numPr><w:numId w:val="1"/></w:numPr><w:ind w:left="720"/></w:pPr></w:style>"#,
        );
        f.numbering = Some(numbering_xml(fmt));
        f
    }

    #[test]
    fn conflicting_styles_are_kept_under_new_names() {
        let dir = tempfile::tempdir().unwrap();
        let a = fx(&format!(r#"<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>A</w:t></w:r></w:p>{}"#, sect("")));
        // у второго документа другой «Normal» и другой «heading 1»; ещё один свой стиль опирается на «heading 1»
        let mut b = fx(&format!(
            r#"<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>B</w:t></w:r></w:p><w:p><w:r><w:t>plain</w:t></w:r></w:p><w:p><w:pPr><w:pStyle w:val="Mine"/></w:pPr><w:r><w:t>mine</w:t></w:r></w:p>{}"#,
            sect("")
        ));
        b.styles = styles_xml(r#"<w:style w:type="paragraph" w:customStyle="1" w:styleId="Mine"><w:name w:val="Mine"/><w:basedOn w:val="Heading1"/><w:next w:val="Normal"/></w:style>"#)
            .replace("<w:b/>", "<w:i/>")
            .replace(r#"<w:name w:val="Normal"/></w:style>"#, r#"<w:name w:val="Normal"/><w:rPr><w:sz w:val="40"/></w:rPr></w:style>"#);
        let (out, w) = run(dir.path(), "keep", &[&a, &b, &b]);
        assert!(w.is_empty(), "оформление сохранено — предупреждений нет: {w:?}");
        let (pkg, d) = doc_of(&out);
        let styles = style_list(&pkg);
        let by_name = |n: &str| styles.iter().find(|s| style_name(s) == n).unwrap_or_else(|| panic!("нет стиля {n}"));
        // стили первого документа не тронуты, у копий новые id и имена, повторов нет
        assert_eq!(by_name("heading 1").child("w:rPr").unwrap().elems().next().unwrap().name, "w:b");
        let h1 = by_name("heading 1 (2)");
        assert_ne!(h1.attr("w:styleId"), Some("Heading1"));
        assert_eq!(h1.attr("w:customStyle"), Some("1"));
        assert!(h1.attr("w:default").is_none());
        assert_eq!(h1.child("w:rPr").unwrap().elems().next().unwrap().name, "w:i");
        assert_eq!(styles.len(), 2 + 3, "Normal, heading 1 и копии Normal, heading 1, а также Mine: {:?}", styles.iter().map(style_name).collect::<Vec<_>>());
        // «Mine» есть только у второго документа (имя не занято) — добавляется как есть и опирается на копию «heading 1»
        let mine = by_name("Mine");
        assert_eq!(mine.child("w:basedOn").unwrap().val(), h1.attr("w:styleId"));
        let normal2 = by_name("Normal (2)");
        assert_eq!(mine.child("w:next").unwrap().val(), normal2.attr("w:styleId"));
        // ссылки второго и третьего документов идут на копии; абзац без стиля получает копию «Normal»
        let ps = pstyles(&d);
        let n2 = normal2.attr("w:styleId").unwrap();
        assert_eq!(ps, ["Heading1", h1.attr("w:styleId").unwrap(), n2, mine.attr("w:styleId").unwrap(), h1.attr("w:styleId").unwrap(), n2, mine.attr("w:styleId").unwrap()]);
        // третий документ совпадает со вторым — своих копий («(3)») он не заводит
        assert!(!styles.iter().any(|s| style_name(s).ends_with("(3)")));
        // режим «как в Word»: побеждают стили первого документа, предупреждение остаётся
        let (out, w) = run_base_wins(dir.path(), "wins", &[&a, &b]);
        assert!(w.iter().any(|m| m.contains("heading 1")), "{w:?}");
        let (pkg, d) = doc_of(&out);
        assert_eq!(style_list(&pkg).len(), 3);
        assert_eq!(pstyles(&d), ["Heading1", "Heading1", "Mine"]);
    }

    #[test]
    fn identical_styles_are_not_duplicated() {
        let dir = tempfile::tempdir().unwrap();
        let a = fx(&format!(r#"<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>A</w:t></w:r></w:p>{}"#, sect("")));
        let (out, w) = run(dir.path(), "same_styles", &[&a, &a, &a]);
        assert!(w.is_empty(), "{w:?}");
        let (pkg, d) = doc_of(&out);
        assert_eq!(style_list(&pkg).len(), 2);
        assert_eq!(pstyles(&d), ["Heading1"; 3]);
        // абзацы без стиля остаются без стиля
        let b = fx(&format!("{}{}", para("P"), sect("")));
        let (out, _) = run(dir.path(), "same_plain", &[&b, &b]);
        assert!(pstyles(&doc_of(&out).1).is_empty());
    }

    #[test]
    fn default_formatting_is_pushed_into_normal_copy() {
        let dir = tempfile::tempdir().unwrap();
        // у первого документа есть «лишний» интервал и размер 11, у второго — размер 15 и без интервала
        let mut a = fx(&format!("{}{}", para("A"), sect("")));
        a.styles = styles_xml("").replace("</w:rPrDefault>", r#"</w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:after="200" w:line="276" w:lineRule="auto"/></w:pPr></w:pPrDefault>"#);
        let mut b = fx(&format!("{}<w:tbl><w:tblPr/><w:tr><w:tc>{}</w:tc></w:tr></w:tbl>{}", para("B"), para("cell"), sect("")));
        b.styles = styles_xml("").replace(r#"w:val="22""#, r#"w:val="30""#);
        let (out, w) = run(dir.path(), "dd_keep", &[&a, &b]);
        assert!(w.is_empty(), "{w:?}");
        let (pkg, d) = doc_of(&out);
        let styles = style_list(&pkg);
        let n2 = styles.iter().find(|s| style_name(s) == "Normal (2)").expect("копия Normal");
        let rpr = n2.child("w:rPr").unwrap();
        assert_eq!(rpr.child("w:sz").unwrap().val(), Some("30"), "размер по умолчанию второго документа перенесён в стиль");
        let sp = n2.child("w:pPr").unwrap().child("w:spacing").unwrap();
        assert_eq!((sp.attr("w:after"), sp.attr("w:line")), (Some("0"), Some("240")), "интервалы первого документа нейтрализованы");
        // все абзацы второго документа (в том числе в таблице) получают копию, абзац первого — нет
        let ps = pstyles(&d);
        assert_eq!(ps, [n2.attr("w:styleId").unwrap(); 2]);
        assert_eq!(texts(d.root.child("w:body").unwrap()), ["A", "B", "cell"]);
        // в режиме «как в Word» остаётся предупреждение
        let (_, w) = run_base_wins(dir.path(), "dd_wins", &[&a, &b]);
        assert!(w.iter().any(|m| m.contains("по умолчанию")), "{w:?}");
        // одинаковые docDefaults — без копий
        let (out, w) = run(dir.path(), "dd_same", &[&a, &a]);
        assert!(w.is_empty());
        assert_eq!(style_list(&doc_of(&out).0).len(), 2);
    }

    #[test]
    fn style_numbering_restarts_in_every_document() {
        let dir = tempfile::tempdir().unwrap();
        let (a, b) = (numbered_doc("decimal", "x"), numbered_doc("decimal", "y"));
        for (name, keep) in [("num_keep", true), ("num_wins", false)] {
            let (out, w) = run_opts(dir.path(), name, &[&a, &b, &b], &Options { keep_source_styles: keep });
            assert!(w.is_empty(), "{w:?}");
            let (pkg, d) = doc_of(&out);
            // стили не дублируются: список общий, но у абзацев последующих документов — собственные списки
            assert_eq!(style_list(&pkg).len(), 3, "{name}");
            let nums: Vec<Option<String>> = all(&d.root, "w:p").iter().map(|p| p.child("w:pPr").and_then(|x| x.child("w:numPr")).and_then(|n| n.child("w:numId")).and_then(|n| n.val()).map(String::from)).collect();
            assert_eq!(nums.len(), 6);
            assert_eq!(nums[0], None, "первый документ не меняется");
            assert_eq!(nums[1], None);
            assert!(nums[2].is_some() && nums[2] == nums[3], "{nums:?}");
            assert!(nums[4].is_some() && nums[4] == nums[5] && nums[4] != nums[2], "{nums:?}");
            // новые списки: тот же абстрактный список и «начать с 1» на каждом уровне
            let n = xml::parse(pkg.get("word/numbering.xml").unwrap()).unwrap();
            let id = nums[2].clone().unwrap();
            let num = all(&n.root, "w:num").into_iter().find(|e| e.attr("w:numId") == Some(id.as_str())).unwrap();
            assert_eq!(num.child("w:lvlOverride").unwrap().child("w:startOverride").unwrap().val(), Some("1"));
            // отступ стиля сохранён явно (у абзаца с собственным списком отступ уровня списка сильнее стиля)
            let ind = all(&d.root, "w:p")[2].child("w:pPr").unwrap().child("w:ind").unwrap().clone();
            assert_eq!(ind.attr("w:left"), Some("720"));
        }
        // маркеры не нумеруются — перезапуск не нужен
        let (a, b) = (numbered_doc("bullet", "x"), numbered_doc("bullet", "y"));
        let (out, _) = run(dir.path(), "num_bullet", &[&a, &b]);
        assert!(all(&doc_of(&out).1.root, "w:numPr").is_empty());
    }

    #[test]
    fn numbered_style_conflict_brings_its_own_list() {
        let dir = tempfile::tempdir().unwrap();
        // у первого документа «heading 1» без нумерации, у второго — с нумерацией «Статья %1.»
        let h = |t: &str| fx(&format!(r#"<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>{t}</w:t></w:r></w:p>{}"#, sect("")));
        let plain = h("P");
        let mut numbered = h("N");
        numbered.styles = styles_xml("").replace(r#"<w:rPr><w:b/></w:rPr>"#, r#"<w:pPr><w:numPr><w:numId w:val="1"/></w:numPr></w:pPr><w:rPr><w:b/></w:rPr>"#);
        numbered.numbering = Some(numbering_xml("decimal").replace("%1.", "Статья %1."));
        for (first, second) in [(&plain, &numbered), (&numbered, &plain), (&numbered, &numbered)] {
            let (out, w) = run(dir.path(), "num_conflict", &[first, second, first]);
            assert!(w.is_empty(), "{w:?}");
            let (pkg, d) = doc_of(&out);
            let styles = style_list(&pkg);
            let numbered_style = |s: &Elem| s.child("w:pPr").and_then(|p| p.child("w:numPr")).is_some();
            // каждый абзац ссылается на стиль, нумерация которого — как в его исходном документе
            let ps = pstyles(&d);
            let want = [first, second, first].map(|f| std::ptr::eq(f, &numbered));
            for (id, want_num) in ps.iter().zip(want) {
                let st = styles.iter().find(|s| s.attr("w:styleId") == Some(id.as_str())).unwrap();
                let has_direct = false;
                assert_eq!(numbered_style(st) || has_direct, want_num, "{id}: {:?}", styles.iter().map(style_name).collect::<Vec<_>>());
            }
        }
    }

    #[test]
    fn linked_character_style_is_renamed_together_with_its_paragraph_style() {
        let dir = tempfile::tempdir().unwrap();
        let link = |extra: &str| {
            let mut f = fx(&format!(r#"<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:rPr><w:rStyle w:val="H1Char"/></w:rPr><w:t>x</w:t></w:r></w:p>{}"#, sect("")));
            f.styles = styles_xml(&format!(r#"<w:style w:type="character" w:customStyle="1" w:styleId="H1Char"><w:name w:val="Heading 1 Char"/><w:link w:val="Heading1"/>{extra}</w:style>"#))
                .replace(r#"<w:basedOn w:val="Normal"/>"#, r#"<w:basedOn w:val="Normal"/><w:link w:val="H1Char"/>"#);
            f
        };
        let a = link("<w:rPr><w:b/></w:rPr>");
        // «Heading 1 Char» одинаков, но связанный «heading 1» отличается
        let mut b = link("<w:rPr><w:b/></w:rPr>");
        b.styles = b.styles.replace("<w:rPr><w:b/></w:rPr></w:style><w:style w:type=\"character\"", "<w:rPr><w:i/></w:rPr></w:style><w:style w:type=\"character\"");
        let (out, w) = run(dir.path(), "linked", &[&a, &b]);
        assert!(w.is_empty(), "{w:?}");
        let (pkg, d) = doc_of(&out);
        let styles = style_list(&pkg);
        let h = styles.iter().find(|s| style_name(s) == "heading 1 (2)").unwrap();
        let c = styles.iter().find(|s| style_name(s) == "Heading 1 Char (2)").expect("связанный знаковый стиль перенесён парой");
        assert_eq!(h.child("w:link").unwrap().val(), c.attr("w:styleId"));
        assert_eq!(c.child("w:link").unwrap().val(), h.attr("w:styleId"));
        let rs = all(&d.root, "w:rStyle");
        assert_eq!(rs[1].val(), c.attr("w:styleId"));
    }

    #[test]
    fn styleref_follows_renamed_style_and_split_instructions_are_renamed() {
        let dir = tempfile::tempdir().unwrap();
        // закладка и поле, разбитые на несколько w:instrText, + STYLEREF на переименованный стиль
        let mut b = fx(&format!(
            r#"<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:bookmarkStart w:id="0" w:name="_Toc123456789"/><w:r><w:t>T</w:t></w:r><w:bookmarkEnd w:id="0"/></w:p><w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> PAGEREF _Toc</w:instrText></w:r><w:r><w:instrText>123456789 \h </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>1</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> STYLEREF "heading 1" </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>T</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>{}"#,
            sect("")
        ));
        b.styles = styles_xml("").replace("<w:b/>", "<w:i/>");
        // у первого документа та же закладка
        let a = fx(&format!(r#"<w:p><w:bookmarkStart w:id="0" w:name="_Toc123456789"/><w:r><w:t>A</w:t></w:r><w:bookmarkEnd w:id="0"/></w:p>{}"#, sect("")));
        let (out, _) = run(dir.path(), "fields", &[&a, &b]);
        let (_, d) = doc_of(&out);
        let instr: Vec<String> = all(&d.root, "w:instrText").iter().map(|e| e.text()).collect();
        assert_eq!(instr.iter().filter(|t| t.trim().is_empty()).count(), 0, "лишние пустые фрагменты удалены: {instr:?}");
        assert!(instr[0].contains("PAGEREF _Toc123456789_m1 \\h"), "{instr:?}");
        assert!(instr[1].contains("STYLEREF \"heading 1 (2)\""), "{instr:?}");
        let names: Vec<String> = all(&d.root, "w:bookmarkStart").iter().map(|e| e.attr("w:name").unwrap().to_string()).collect();
        assert_eq!(names, ["_Toc123456789", "_Toc123456789_m1"]);
    }

    #[test]
    fn toc_fields_in_appended_documents_are_reported() {
        let dir = tempfile::tempdir().unwrap();
        let a = fx(&format!("{}{}", para("A"), sect("")));
        let toc = |instr: &str| {
            fx(&format!(
                r#"<w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> TOC \o "1-3" </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>1</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p><w:fldSimple w:instr="{instr}"><w:r><w:t>x</w:t></w:r></w:fldSimple>{}"#,
                sect("")
            ))
        };
        let (_, w) = run(dir.path(), "toc", &[&a, &toc(" PAGE ")]);
        assert_eq!(w.len(), 1, "{w:?}");
        assert!(w[0].contains("F9") && w[0].contains("оглавлен"));
        // у первого документа оглавление остаётся верным — предупреждения нет
        let (_, w) = run(dir.path(), "toc_first", &[&toc(" PAGE "), &a]);
        assert!(w.is_empty(), "{w:?}");
    }

    fn settings_part(f: Fx, extra: &str) -> Fx {
        f.part("word/settings.xml", format!(r#"<w:settings xmlns:w="{WNS}">{extra}</w:settings>"#))
            .rel("rS", "settings", "settings.xml")
            .ct("/word/settings.xml", &format!("{CT_NOTES}settings+xml"))
    }

    fn with_header(f: Fx, text: &str) -> Fx {
        f.part("word/hdr.xml", hdr("w:hdr", text))
            .rel("rH", "header", "hdr.xml")
            .ct("/word/hdr.xml", &format!("{CT_NOTES}header+xml"))
            .part("word/ftr.xml", hdr("w:ftr", &format!("{text}-F")))
            .rel("rF", "footer", "ftr.xml")
            .ct("/word/ftr.xml", &format!("{CT_NOTES}footer+xml"))
    }

    fn sect_hf() -> String {
        sect(r#"<w:headerReference w:type="default" r:id="rH"/><w:footerReference w:type="default" r:id="rF"/>"#)
    }

    #[test]
    fn even_odd_headers_are_shared_by_all_documents() {
        let dir = tempfile::tempdir().unwrap();
        let plain = with_header(fx(&format!("{}{}", para("P"), sect_hf())), "HDR-P");
        let even = settings_part(with_header(fx(&format!("{}{}", para("E"), sect_hf())), "HDR-E"), "<w:evenAndOddHeaders/>");
        for (name, docs) in [("eo_later", [&even, &plain]), ("eo_first", [&plain, &even])] {
            let (out, w) = run(dir.path(), name, &docs);
            assert!(!w.iter().any(|m| m.contains("чётных")), "{w:?}");
            let (pkg, d) = doc_of(&out);
            let settings = xml::parse(pkg.get("word/settings.xml").unwrap()).unwrap();
            assert!(setting_on(&settings.root, "w:evenAndOddHeaders"), "{name}");
            // у документа без настройки чётная страница показывает его обычный колонтитул и нижний колонтитул
            let sects = all(&d.root, "w:sectPr");
            let plain_idx = if name == "eo_later" { 1 } else { 0 };
            let sp = &sects[plain_idx];
            for tag in ["w:headerReference", "w:footerReference"] {
                let get = |ty: &str| sp.elems().find(|e| e.name == tag && e.attr("w:type").unwrap_or("default") == ty).and_then(|e| e.attr("r:id")).map(String::from);
                assert!(get("even").is_some() && get("even") == get("default"), "{name}: {tag}");
            }
            // у документа с настройкой чётный колонтитул не берётся у соседнего документа
            let other = &sects[1 - plain_idx];
            let ev = |s: &Elem| s.elems().find(|e| e.name == "w:headerReference" && e.attr("w:type") == Some("even")).and_then(|e| e.attr("r:id")).map(String::from);
            assert!(ev(other).is_none() || ev(other) != ev(sp), "{name}");
        }
        // настройка есть у всех — ничего не добавляется
        let (out, _) = run(dir.path(), "eo_all", &[&even, &even]);
        let (_, d) = doc_of(&out);
        assert!(all(&d.root, "w:headerReference").iter().all(|e| e.attr("w:type") == Some("default")));
    }

    #[test]
    fn page_background_difference_is_reported() {
        let dir = tempfile::tempdir().unwrap();
        // w:background — дочерний элемент w:document перед w:body
        let write_bg = |dir: &Path, name: &str, color: Option<&str>| {
            let f = fx(&format!("{}{}", para("A"), sect("")));
            let p = dir.join(name);
            f.write(&p);
            let mut pkg = Pkg::read(&p).unwrap();
            let doc = String::from_utf8(pkg.get("word/document.xml").unwrap().to_vec()).unwrap();
            let doc = match color {
                Some(c) => doc.replace("<w:body>", &format!(r#"<w:background w:color="{c}"/><w:body>"#)),
                None => doc,
            };
            pkg.set("word/document.xml", doc.into_bytes());
            pkg.write(&p).unwrap();
            p
        };
        let (plain, yellow, blue) = (write_bg(dir.path(), "p.docx", None), write_bg(dir.path(), "y.docx", Some("FFF2CC")), write_bg(dir.path(), "b.docx", Some("0000FF")));
        let out = dir.path().join("o.docx");
        let o = Options::default();
        let w = merge(&[plain.clone(), yellow.clone()], &out, &o).unwrap();
        assert!(w.iter().any(|m| m.contains("фон страницы") && m.contains("не перенесён")), "{w:?}");
        let w = merge(&[yellow.clone(), plain.clone()], &out, &o).unwrap();
        assert!(w.iter().any(|m| m.contains("фон страницы") && m.contains("применён")), "{w:?}");
        let w = merge(&[yellow.clone(), blue], &out, &o).unwrap();
        assert!(w.iter().any(|m| m.contains("фон страницы") && m.contains("отличается")), "{w:?}");
        assert!(String::from_utf8_lossy(Pkg::read(&out).unwrap().get("word/document.xml").unwrap()).contains("FFF2CC"));
        assert!(merge(&[yellow.clone(), yellow], &out, &o).unwrap().is_empty());
    }

    #[test]
    fn default_namespace_and_utf16_parts_are_accepted() {
        let dir = tempfile::tempdir().unwrap();
        let a = fx(&format!("{}{}", para("A"), sect("")));
        // пространство имён по умолчанию + атрибуты с вложенным объявлением (как делает ElementTree)
        let pd = dir.path().join("d.docx");
        fx("").write(&pd);
        let mut pkg = Pkg::read(&pd).unwrap();
        pkg.set(
            "word/document.xml",
            format!(r#"<?xml version="1.0"?><document xmlns="{WNS}" xmlns:rel="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><body xmlns:ns0="{WNS}"><p><pPr><pStyle ns0:val="Heading1"/></pPr><r><t>DEF</t></r></p><sectPr><pgSz ns0:w="1" ns0:h="2"/></sectPr></body></document>"#).into_bytes(),
        );
        pkg.write(&pd).unwrap();
        // UTF-16 с BOM (LE) и без BOM (BE)
        let utf16 = |s: &str, be: bool, bom: bool| {
            let mut v = vec![];
            if bom {
                v.extend(if be { [0xFE, 0xFF] } else { [0xFF, 0xFE] });
            }
            for u in s.encode_utf16() {
                v.extend(if be { u.to_be_bytes() } else { u.to_le_bytes() });
            }
            v
        };
        let mk16 = |name: &str, be: bool, bom: bool| {
            let f = fx(&format!("{}{}", para("Юникод"), sect("")));
            let p = dir.path().join(name);
            f.write(&p);
            let mut pkg = Pkg::read(&p).unwrap();
            for part in ["word/document.xml", "word/styles.xml", "word/_rels/document.xml.rels", "[Content_Types].xml"] {
                let text = String::from_utf8(pkg.get(part).unwrap().to_vec()).unwrap();
                pkg.set(part, utf16(&text, be, bom));
            }
            pkg.write(&p).unwrap();
            p
        };
        let out = dir.path().join("o.docx");
        for (n, p) in [("def", pd), ("u16le", mk16("le.docx", false, true)), ("u16be", mk16("be.docx", true, false))] {
            let pa = dir.path().join("a.docx");
            a.write(&pa);
            merge(&[pa, p], &out, &Options::default()).unwrap_or_else(|e| panic!("{n}: {e:#}"));
            let errs = check(&out);
            assert!(errs.is_empty(), "{n}: {errs:#?}");
            let (_, d) = doc_of(&out);
            let t = texts(d.root.child("w:body").unwrap());
            assert_eq!(t.len(), 2, "{n}");
            assert!(t[1] == "DEF" || t[1] == "Юникод", "{t:?}");
        }
        // результат всегда в UTF-8
        assert!(Pkg::read(&out).unwrap().get("word/document.xml").unwrap().starts_with(b"<?xml"));
    }

    #[test]
    fn malformed_xml_is_rejected_with_a_clear_message() {
        let dir = tempfile::tempdir().unwrap();
        let a = fx(&format!("{}{}", para("A"), sect("")));
        let pa = dir.path().join("a.docx");
        a.write(&pa);
        let out = dir.path().join("o.docx");
        let bad = |name: &str, part: &str, f: &dyn Fn(String) -> Vec<u8>| {
            let p = dir.path().join(name);
            fx(&format!("{}{}", para("B &amp; C"), sect(""))).write(&p);
            let mut pkg = Pkg::read(&p).unwrap();
            let text = String::from_utf8(pkg.get(part).unwrap().to_vec()).unwrap();
            pkg.set(part, f(text));
            pkg.write(&p).unwrap();
            p
        };
        let cases: Vec<(&str, PathBuf, &str)> = vec![
            ("амперсанд", bad("amp.docx", "word/document.xml", &|t| t.replace("B &amp; C", "B & C").into_bytes()), "&amp;"),
            ("ссылка на символ", bad("ref.docx", "word/styles.xml", &|t| t.replace("<w:b/>", "<w:b/>&#x0;").into_bytes()), "недопустим"),
            ("обрыв", bad("cut.docx", "word/document.xml", &|t| t[..t.len() / 2].as_bytes().to_vec()), "обрезан"),
            (
                "обрыв на символе",
                bad("cut2.docx", "word/document.xml", &|t| {
                    let mut v = t[..t.find("B &amp; C").unwrap()].as_bytes().to_vec();
                    v.push(0xD0);
                    v
                }),
                "обрезан",
            ),
            (
                "не UTF-8",
                bad("enc.docx", "word/document.xml", &|t| {
                    let mut v = t.into_bytes();
                    let i = v.windows(9).position(|w| w == b"B &amp; C").unwrap();
                    v[i] = 0xE9;
                    v
                }),
                "UTF-8",
            ),
            ("префикс", bad("pfx.docx", "word/document.xml", &|t| t.replace("<w:t>", "<q:t>").replace("</w:t>", "</q:t>").into_bytes()), "не объявлен"),
            ("повтор атрибута", bad("dup.docx", "word/document.xml", &|t| t.replace("<w:t>", r#"<w:t a="1" a="2">"#).into_bytes()), "повторяется"),
        ];
        for (what, p, key) in cases {
            let e = merge(&[pa.clone(), p.clone()], &out, &Options::default()).unwrap_err();
            let msg = format!("{e:#}");
            assert!(msg.contains(&p.file_name().unwrap().to_string_lossy().to_string()), "{what}: нет имени файла: {msg}");
            assert!(msg.contains(key), "{what}: {msg}");
            assert!(!msg.contains("UTF-8") || what == "не UTF-8", "{what}: обрыв нельзя выдавать за кодировку: {msg}");
            assert!(!msg.chars().any(|c| c.is_ascii_alphabetic()) || msg.contains("«") , "{msg}");
        }
    }

    #[test]
    fn damaged_archives_are_described_in_russian() {
        let dir = tempfile::tempdir().unwrap();
        let a = fx(&format!("{}{}", para("A"), sect("")));
        let pa = dir.path().join("a.docx");
        a.write(&pa);
        let out = dir.path().join("o.docx");
        let full = std::fs::read(&pa).unwrap();
        let cases: Vec<(&str, Vec<u8>, &str)> = vec![
            ("обрезан", full[..full.len() / 2].to_vec(), "обрезан"),
            ("не zip", b"just some text, not an archive".to_vec(), "повреждён"),
            ("пустой", vec![], "повреждён"),
        ];
        for (what, bytes, key) in cases {
            let p = dir.path().join(format!("{what}.docx"));
            std::fs::write(&p, bytes).unwrap();
            let msg = format!("{:#}", merge(&[pa.clone(), p], &out, &Options::default()).unwrap_err());
            assert!(msg.contains(key) && msg.contains(".docx"), "{what}: {msg}");
            for raw in ["invalid Zip", "Could not find", "os error", "Os {", "Error {"] {
                assert!(!msg.contains(raw), "{what}: сырой текст библиотеки: {msg}");
            }
        }
        // отсутствующий файл
        let msg = format!("{:#}", merge(&[pa.clone(), dir.path().join("нет.docx")], &out, &Options::default()).unwrap_err());
        assert!(msg.contains("нет.docx") && msg.contains("не найден") && !msg.contains("os error"), "{msg}");
        // повреждённая контрольная сумма записи
        let mut broken = full.clone();
        let pos = broken.windows(4).position(|w| w == b"PK\x03\x04").unwrap();
        let data_start = pos + 30 + u16::from_le_bytes([broken[pos + 26], broken[pos + 27]]) as usize + u16::from_le_bytes([broken[pos + 28], broken[pos + 29]]) as usize;
        for b in broken[data_start..data_start + 8].iter_mut() {
            *b ^= 0xFF;
        }
        let p = dir.path().join("crc.docx");
        std::fs::write(&p, broken).unwrap();
        let msg = format!("{:#}", merge(&[pa, p], &out, &Options::default()).unwrap_err());
        assert!(msg.contains("crc.docx") && msg.contains("повреждён"), "{msg}");
        assert!(!msg.contains("checksum") && !msg.contains("deflate"), "{msg}");
    }

    #[test]
    fn xml_decode_handles_bom_and_utf16() {
        assert_eq!(xml::decode(b"\xEF\xBB\xBF<a/>").unwrap(), "<a/>");
        assert_eq!(xml::decode(&[0xFF, 0xFE, b'<', 0, b'a', 0, b'/', 0, b'>', 0]).unwrap(), "<a/>");
        assert_eq!(xml::decode(&[0, b'<', 0, b'a', 0, b'/', 0, b'>']).unwrap(), "<a/>");
        assert!(xml::decode(&[0xFF, 0xFE, b'<']).is_err());
        assert!(xml::parse("<a b=\"1\" c='&#1;'/>".as_bytes()).is_err());
        assert!(xml::parse("<a>&#x41;&lt;&#65;</a>".as_bytes()).is_ok());
        assert!(xml::parse("<a>a &b; c</a>".as_bytes()).is_err());
        assert!(xml::parse("<a b=\"x<y\"/>".as_bytes()).is_err());
    }

    fn theme_xml(accent: &str, minor: &str) -> String {
        format!(
            r#"<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:themeElements><a:clrScheme name="x"><a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1><a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1><a:accent1><a:srgbClr val="{accent}"/></a:accent1></a:clrScheme><a:fontScheme name="y"><a:majorFont><a:latin typeface="Major{minor}"/><a:ea typeface=""/><a:cs typeface=""/></a:majorFont><a:minorFont><a:latin typeface="{minor}"/><a:ea typeface=""/><a:cs typeface=""/></a:minorFont></a:fontScheme></a:themeElements></a:theme>"#
        )
    }

    fn with_theme(f: Fx, accent: &str, minor: &str) -> Fx {
        f.part("word/theme/theme1.xml", theme_xml(accent, minor))
            .rel("rT", "theme", "theme/theme1.xml")
            .ct("/word/theme/theme1.xml", "application/vnd.openxmlformats-officedocument.theme+xml")
    }

    #[test]
    fn theme_fonts_and_colors_are_made_explicit_when_themes_differ() {
        let dir = tempfile::tempdir().unwrap();
        let body = |t: &str| {
            format!(
                r#"<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:rPr><w:color w:val="auto" w:themeColor="accent1"/><w:rFonts w:asciiTheme="minorHAnsi" w:ascii="Arial"/></w:rPr><w:t>{t}</w:t></w:r></w:p><w:p><w:r><w:drawing><wp:inline><wp:docPr id="1" name="s"/><a:graphic><a:graphicData uri="u"><a:solidFill><a:schemeClr val="accent1"><a:lumMod val="50000"/></a:schemeClr></a:solidFill></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>{}"#,
                sect("")
            )
        };
        let styles = |extra: &str| styles_xml("").replace("<w:sz w:val=\"22\"/>", &format!("<w:rFonts w:asciiTheme=\"minorHAnsi\"/><w:sz w:val=\"22\"/>{extra}"));
        let mut a = with_theme(fx(&body("A")), "FF0000", "Calibri");
        a.styles = styles("");
        let mut b = with_theme(fx(&body("B")), "00FF00", "Cambria");
        b.styles = styles("");
        let (out, w) = run(dir.path(), "theme_keep", &[&a, &b]);
        assert!(w.is_empty(), "{w:?}");
        let (pkg, d) = doc_of(&out);
        // первый документ не тронут, у второго цвет темы стал явным, шрифт темы — названием шрифта
        let runs: Vec<Elem> = all(&d.root, "w:rPr").into_iter().filter(|r| r.child("w:color").is_some()).collect();
        assert_eq!(runs[0].child("w:color").unwrap().attr("w:themeColor"), Some("accent1"));
        let c = runs[1].child("w:color").unwrap();
        assert_eq!((c.val(), c.attr("w:themeColor")), (Some("00FF00"), None));
        let f = runs[1].child("w:rFonts").unwrap();
        assert_eq!((f.attr("w:ascii"), f.attr("w:asciiTheme")), (Some("Cambria"), None));
        // DrawingML: цвет схемы заменён на RGB с теми же преобразованиями
        let clr = all(&d.root, "a:srgbClr");
        assert_eq!(clr.len(), 1);
        assert_eq!(clr[0].val(), None);
        assert_eq!(clr[0].attr("val"), Some("00FF00"));
        assert!(clr[0].child("a:lumMod").is_some());
        assert_eq!(all(&d.root, "a:schemeClr").len(), 1, "у первого документа цвет схемы остаётся");
        // стиль по умолчанию второго документа: шрифт темы стал явным
        let styles = style_list(&pkg);
        let n2 = styles.iter().find(|s| style_name(s) == "Normal (2)").expect("копия Normal");
        let f = n2.child("w:rPr").unwrap().child("w:rFonts").unwrap();
        assert_eq!((f.attr("w:ascii"), f.attr("w:asciiTheme")), (Some("Cambria"), None));
        // одинаковая тема — ничего не заменяется и стили не дублируются
        let b2 = with_theme(fx(&body("B")), "FF0000", "Calibri");
        let mut b2 = b2;
        b2.styles = styles_xml("").replace("<w:sz w:val=\"22\"/>", "<w:rFonts w:asciiTheme=\"minorHAnsi\"/><w:sz w:val=\"22\"/>");
        let (out, w) = run(dir.path(), "theme_same", &[&a, &b2]);
        assert!(w.is_empty());
        assert_eq!(all(&doc_of(&out).1.root, "a:schemeClr").len(), 2);
        assert_eq!(style_list(&doc_of(&out).0).len(), 2);
    }

    #[test]
    fn table_cell_paragraphs_keep_table_style_over_pushed_defaults() {
        let dir = tempfile::tempdir().unwrap();
        let table_style = r#"<w:style w:type="table" w:styleId="TableGrid"><w:name w:val="Table Grid"/><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr></w:style>"#;
        let a = fx(&format!("{}{}", para("A"), sect("")));
        let mut b = fx(&format!(
            r#"{}<w:tbl><w:tblPr><w:tblStyle w:val="TableGrid"/></w:tblPr><w:tr><w:tc><w:p><w:r><w:t>cell</w:t></w:r></w:p><w:p><w:pPr><w:pStyle w:val="Normal"/></w:pPr><w:r><w:t>n</w:t></w:r></w:p><w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>h</w:t></w:r></w:p></w:tc></w:tr></w:tbl>{}"#,
            para("B"),
            sect("")
        ));
        b.styles = styles_xml(table_style).replace("</w:rPrDefault>", r#"</w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:after="100" w:line="300" w:lineRule="auto"/><w:jc w:val="both"/></w:pPr></w:pPrDefault>"#);
        let (out, w) = run(dir.path(), "cells", &[&a, &b]);
        assert!(w.is_empty(), "{w:?}");
        let (pkg, d) = doc_of(&out);
        let styles = style_list(&pkg);
        let n2 = styles.iter().find(|s| style_name(s) == "Normal (2)").unwrap();
        let cell = styles.iter().find(|s| style_name(s) == "Normal в таблице (2)").expect("копия для ячеек");
        let sp = |s: &Elem| s.child("w:pPr").and_then(|p| p.child("w:spacing")).map(|e| (e.attr("w:after").map(String::from), e.attr("w:line").map(String::from)));
        // у обычного абзаца — интервалы docDefaults, у абзаца в ячейке — только то, чего стиль таблицы не задаёт
        assert_eq!(sp(n2), Some((Some("100".into()), Some("300".into()))));
        assert_eq!(sp(cell), None, "и after, и line, и lineRule заданы стилем таблицы: {:?}", cell.child("w:pPr"));
        assert!(cell.child("w:pPr").unwrap().child("w:jc").is_some(), "выравнивание стиль таблицы не задаёт — переносится");
        let ps = pstyles(&d);
        let (n2id, cid) = (n2.attr("w:styleId").unwrap(), cell.attr("w:styleId").unwrap());
        // абзац без стиля и с явным «Normal» в ячейке — копия для ячеек; заголовок — своя копия; вне таблицы — копия Normal
        assert_eq!(ps[0], n2id);
        assert_eq!(&ps[1..3], [cid, cid]);
        assert_ne!(ps[3], cid);
        assert!(all(&d.root, "w:p").iter().all(|p| p.attr("mf:cell").is_none()));
    }

    #[test]
    fn style_with_missing_parent_or_default_is_handled() {
        let dir = tempfile::tempdir().unwrap();
        let a = fx(&format!("{}{}", para("A"), sect("")));
        // у второго документа нет стиля «по умолчанию», а docDefaults отличаются; стиль ссылается на несуществующий
        let mut b = fx(&format!(r#"{}<w:p><w:pPr><w:pStyle w:val="X"/></w:pPr><w:r><w:t>x</w:t></w:r></w:p>{}"#, para("B"), sect("")));
        b.styles = format!(
            r#"<?xml version="1.0"?><w:styles xmlns:w="{WNS}"><w:docDefaults><w:rPrDefault><w:rPr><w:sz w:val="40"/></w:rPr></w:rPrDefault></w:docDefaults><w:style w:type="paragraph" w:styleId="X"><w:name w:val="X"/><w:basedOn w:val="Gone"/><w:next w:val="Gone"/><w:link w:val="Gone"/></w:style></w:styles>"#
        );
        let (out, w) = run(dir.path(), "nodef", &[&a, &b]);
        assert!(w.is_empty(), "{w:?}");
        let (pkg, d) = doc_of(&out);
        let styles = style_list(&pkg);
        let x = styles.iter().find(|s| style_name(s) == "X").unwrap();
        assert!(x.child("w:basedOn").is_none() && x.child("w:next").is_none() && x.child("w:link").is_none());
        // абзацы без стиля получили созданный «обычный» стиль с размером шрифта второго документа
        let n2 = styles.iter().find(|s| style_name(s) == "Normal (2)").expect("созданный стиль");
        assert_eq!(n2.child("w:rPr").unwrap().child("w:sz").unwrap().val(), Some("40"));
        let ps = pstyles(&d);
        assert_eq!(ps, [n2.attr("w:styleId").unwrap(), "X"]);
    }

    #[test]
    fn title_page_without_own_first_footer_stays_blank() {
        let dir = tempfile::tempdir().unwrap();
        // у первого документа есть колонтитул «по умолчанию», у второго — особая первая страница только с верхним
        let a = with_header(fx(&format!("{}{}", para("A"), sect_hf())), "HDR-A");
        let b = with_header(fx(&format!("{}{}", para("B"), sect(r#"<w:headerReference w:type="first" r:id="rH"/><w:titlePg/>"#))), "HDR-B");
        let (out, _) = run(dir.path(), "titlepg", &[&a, &b]);
        let (pkg, d) = doc_of(&out);
        let last = d.root.child("w:body").unwrap().child("w:sectPr").unwrap().clone();
        let r = |tag: &str, ty: &str| last.elems().find(|e| e.name == tag && e.attr("w:type") == Some(ty)).and_then(|e| e.attr("r:id")).map(String::from);
        let rels = Rels::parse(pkg.get("word/_rels/document.xml.rels").unwrap()).unwrap();
        // нижний колонтитул первой страницы — пустой, а не колонтитул предыдущего документа
        let f = r("w:footerReference", "first").expect("ссылка на пустой нижний колонтитул первой страницы");
        let target = resolve_part(&pkg, "word/document.xml", &rels.get(&f).unwrap().target).unwrap();
        assert!(!String::from_utf8_lossy(pkg.get(&target).unwrap()).contains("HDR-A"));
        // обычные колонтитулы предыдущего документа тоже не наследуются
        assert!(r("w:headerReference", "default").is_some() && r("w:footerReference", "default").is_some());
    }

    #[test]
    fn compatibility_mode_difference_is_reported() {
        let dir = tempfile::tempdir().unwrap();
        let compat = |mode: &str, extra: &str| {
            settings_part(
                fx(&format!("{}{}", para("P"), sect(""))),
                &format!(r#"<w:compat>{extra}<w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="{mode}"/></w:compat>"#),
            )
        };
        let (_, w) = run(dir.path(), "compat_same", &[&compat("15", ""), &compat("15", "")]);
        assert!(w.is_empty(), "{w:?}");
        let (_, w) = run(dir.path(), "compat_mode", &[&compat("15", ""), &compat("14", "")]);
        assert!(w.iter().any(|m| m.contains("режим совместимости")), "{w:?}");
        let (_, w) = run(dir.path(), "compat_flag", &[&compat("15", ""), &compat("15", "<w:usePrinterMetrics/>")]);
        assert!(w.iter().any(|m| m.contains("режим совместимости")), "{w:?}");
    }

    #[test]
    fn remerging_an_output_with_renamed_styles_keeps_every_look() {
        let dir = tempfile::tempdir().unwrap();
        // три документа с разным «Normal»; результат слияния двух снова входит в слияние (id вида Normal_m1 уже заняты)
        let mk = |t: &str, prop: &str| {
            let mut f = fx(&format!("{}{}", para(t), sect("")));
            f.styles = styles_xml("").replace(r#"<w:name w:val="Normal"/></w:style>"#, &format!(r#"<w:name w:val="Normal"/><w:rPr><{prop}/></w:rPr></w:style>"#));
            f
        };
        let (x, y, z) = (mk("X", "w:b"), mk("Y", "w:i"), mk("Z", "w:u"));
        let (xy, _) = run(dir.path(), "re_xy", &[&x, &y]);
        let zp = dir.path().join("z.docx");
        z.write(&zp);
        for order in [[zp.clone(), xy.clone()], [xy.clone(), zp.clone()]] {
            let out = dir.path().join("re_out.docx");
            merge(&order, &out, &Options::default()).unwrap();
            let errs = check(&out);
            assert!(errs.is_empty(), "{errs:#?}");
            let (pkg, d) = doc_of(&out);
            let styles = style_list(&pkg);
            let default_id = styles.iter().find(|s| s.attr("w:default") == Some("1")).and_then(|s| s.attr("w:styleId")).unwrap().to_string();
            let look: HashMap<String, String> = all(d.root.child("w:body").unwrap(), "w:p")
                .iter()
                .map(|p| {
                    let sid = p.child("w:pPr").and_then(|x| x.child("w:pStyle")).and_then(|s| s.val()).map(String::from).unwrap_or(default_id.clone());
                    let st = styles.iter().find(|s| s.attr("w:styleId") == Some(sid.as_str())).unwrap();
                    (texts(p).concat(), st.child("w:rPr").unwrap().elems().next().unwrap().name.clone())
                })
                .collect();
            assert_eq!(look["X"], "w:b");
            assert_eq!(look["Y"], "w:i");
            assert_eq!(look["Z"], "w:u");
        }
    }

    #[test]
    fn normal_without_default_flag_is_the_default_paragraph_style() {
        let dir = tempfile::tempdir().unwrap();
        let a = fx(&format!("{}{}", para("A"), sect("")));
        // как у LibreOffice: «Normal» без w:default, абзац без стиля получает именно его
        let mut b = fx(&format!("{}{}", para("B"), sect("")));
        b.styles = styles_xml("")
            .replace(r#" w:default="1" w:styleId="Normal""#, r#" w:styleId="Normal""#)
            .replace(r#"<w:name w:val="Normal"/></w:style>"#, r#"<w:name w:val="Normal"/><w:rPr><w:sz w:val="40"/></w:rPr></w:style>"#);
        let (out, w) = run(dir.path(), "lo_normal", &[&a, &b]);
        assert!(w.is_empty(), "{w:?}");
        let (pkg, d) = doc_of(&out);
        let styles = style_list(&pkg);
        let n2 = styles.iter().find(|s| style_name(s) == "Normal (2)").expect("копия Normal");
        assert_eq!(n2.child("w:rPr").unwrap().child("w:sz").unwrap().val(), Some("40"));
        assert_eq!(pstyles(&d), [n2.attr("w:styleId").unwrap()]);
        // стиль по умолчанию у результата остаётся один
        assert_eq!(styles.iter().filter(|s| s.attr("w:default") == Some("1")).count(), 1);
    }

    #[test]
    fn document_without_any_default_style_does_not_inherit_the_base_normal() {
        let dir = tempfile::tempdir().unwrap();
        // у первого документа «Normal» с размером 30; у второго стилей абзаца «по умолчанию» нет вообще
        let mut a = fx(&format!("{}{}", para("A"), sect("")));
        a.styles = styles_xml("").replace(r#"<w:name w:val="Normal"/></w:style>"#, r#"<w:name w:val="Normal"/><w:rPr><w:sz w:val="30"/></w:rPr></w:style>"#);
        let mut b = fx(&format!("{}{}", para("B"), sect("")));
        b.styles = styles_xml("").replace(r#"<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>"#, "");
        let (out, w) = run(dir.path(), "nodefault", &[&a, &b]);
        assert!(w.is_empty(), "{w:?}");
        let (pkg, d) = doc_of(&out);
        let styles = style_list(&pkg);
        let ps = pstyles(&d);
        assert_eq!(ps.len(), 1, "абзац второго документа получил пустой стиль, а не «Normal» первого");
        let own = styles.iter().find(|s| s.attr("w:styleId") == Some(ps[0].as_str())).unwrap();
        assert!(own.child("w:rPr").is_none_or(|r| r.child("w:sz").is_none()), "{own:?}");
    }
}
