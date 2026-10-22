//! Объединение листов Excel в одну книгу.
//!
//! Работает на уровне пакета Office Open XML: части выбранных листов (ячейки, стили, рисунки,
//! диаграммы, таблицы, сводные таблицы, примечания…) копируются как есть, а общие ресурсы
//! (стили, общие строки, имена листов, именованные диапазоны, номера таблиц, медиафайлы)
//! объединяются с перенумерацией. Так ничего не теряется на промежуточной модели,
//! а то, что перенести нельзя, возвращается пользователю в виде предупреждений.

use anyhow::{Result, anyhow, bail};
use std::collections::{BTreeMap, BTreeSet, HashMap, HashSet};
use std::io::Write;
use std::path::{Path, PathBuf};

/// Минимальный XML-слой: разбор с пространствами имён и запись.
/// Нужен, чтобы копировать части пакета Office Open XML без потерь
/// (в том числе при нестандартных префиксах: `x:sheet`, `<xdr:wsDr>` и т.п.).
mod xml {
    use anyhow::{Result, anyhow, bail};
    use std::io::Write;

    pub const NS_MAIN: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    pub const NS_R: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    pub const NS_XML: &str = "http://www.w3.org/XML/1998/namespace";
    pub const NS_CT: &str = "http://schemas.openxmlformats.org/package/2006/content-types";
    pub const NS_PKG_REL: &str = "http://schemas.openxmlformats.org/package/2006/relationships";

    #[derive(Clone, Debug)]
    pub struct Attr {
        /// Имя как в файле (`r:id`, `s`, `xmlns:r`).
        pub q: String,
        /// URI пространства имён префикса (для атрибутов без префикса и `xmlns*` — пусто).
        pub ns: String,
        pub local: String,
        pub v: String,
    }

    #[derive(Clone, Debug, Default)]
    pub struct Elem {
        pub ns: String,
        pub prefix: String,
        pub local: String,
        pub attrs: Vec<Attr>,
        pub kids: Vec<Node>,
    }

    #[derive(Clone, Debug)]
    pub enum Node {
        E(Elem),
        T(String),
    }

    fn is_decl(a: &Attr) -> bool {
        a.q == "xmlns" || a.q.starts_with("xmlns:")
    }

    impl Elem {
        pub fn new(ns: &str, local: &str) -> Elem {
            Elem { ns: ns.into(), local: local.into(), ..Default::default() }
        }
        pub fn main(local: &str) -> Elem {
            Elem::new(NS_MAIN, local)
        }
        /// Атрибут без префикса.
        pub fn attr(&self, name: &str) -> Option<&str> {
            self.attrs.iter().find(|a| a.q == name).map(|a| a.v.as_str())
        }
        /// Атрибут с пространством имён (например `r:id`).
        pub fn attr_ns(&self, ns: &str, local: &str) -> Option<&str> {
            self.attrs.iter().find(|a| a.ns == ns && a.local == local && !is_decl(a)).map(|a| a.v.as_str())
        }
        pub fn set(&mut self, name: &str, v: impl Into<String>) {
            let v = v.into();
            match self.attrs.iter_mut().find(|a| a.q == name) {
                Some(a) => a.v = v,
                None => self.attrs.push(Attr { q: name.into(), ns: String::new(), local: name.into(), v }),
            }
        }
        pub fn set_ns(&mut self, ns: &str, prefix: &str, local: &str, v: impl Into<String>) {
            let v = v.into();
            match self.attrs.iter_mut().find(|a| a.ns == ns && a.local == local && !is_decl(a)) {
                Some(a) => a.v = v,
                None => self.attrs.push(Attr { q: format!("{prefix}:{local}"), ns: ns.into(), local: local.into(), v }),
            }
        }
        pub fn remove(&mut self, name: &str) -> Option<String> {
            let i = self.attrs.iter().position(|a| a.q == name)?;
            Some(self.attrs.remove(i).v)
        }
        pub fn remove_ns(&mut self, ns: &str, local: &str) -> Option<String> {
            let i = self.attrs.iter().position(|a| a.ns == ns && a.local == local && !is_decl(a))?;
            Some(self.attrs.remove(i).v)
        }
        pub fn child(&self, local: &str) -> Option<&Elem> {
            self.children().find(|e| e.local == local)
        }
        pub fn child_mut(&mut self, local: &str) -> Option<&mut Elem> {
            self.kids.iter_mut().find_map(|n| match n {
                Node::E(e) if e.local == local => Some(e),
                _ => None,
            })
        }
        pub fn children(&self) -> impl Iterator<Item = &Elem> {
            self.kids.iter().filter_map(|n| if let Node::E(e) = n { Some(e) } else { None })
        }
        pub fn children_mut(&mut self) -> impl Iterator<Item = &mut Elem> {
            self.kids.iter_mut().filter_map(|n| if let Node::E(e) = n { Some(e) } else { None })
        }
        pub fn text(&self) -> String {
            let mut s = String::new();
            for n in &self.kids {
                if let Node::T(t) = n {
                    s.push_str(t);
                }
            }
            s
        }
        pub fn set_text(&mut self, t: impl Into<String>) {
            self.kids.retain(|n| matches!(n, Node::E(_)));
            self.kids.push(Node::T(t.into()));
        }
        pub fn push(&mut self, e: Elem) {
            self.kids.push(Node::E(e));
        }
        /// Обход всех потомков (включая сам элемент) с изменением.
        pub fn walk_mut(&mut self, f: &mut dyn FnMut(&mut Elem)) {
            f(self);
            for k in &mut self.kids {
                if let Node::E(e) = k {
                    e.walk_mut(f);
                }
            }
        }
        pub fn walk(&self, f: &mut dyn FnMut(&Elem)) {
            f(self);
            for e in self.children() {
                e.walk(f);
            }
        }
    }

    // ------------------------------------------------------------------ разбор

    pub enum Ev {
        Start(Elem),
        Empty(Elem),
        End,
        Text(String),
        Eof,
    }

    pub struct Parser<'a> {
        s: &'a str,
        pos: usize,
        scopes: Vec<Vec<(String, String)>>,
        names: Vec<&'a str>,
    }

    const CP1252_HI: [u16; 32] = [
        0x20AC, 0x0081, 0x201A, 0x0192, 0x201E, 0x2026, 0x2020, 0x2021, 0x02C6, 0x2030, 0x0160, 0x2039, 0x0152, 0x008D, 0x017D, 0x008F, 0x0090, 0x2018,
        0x2019, 0x201C, 0x201D, 0x2022, 0x2013, 0x2014, 0x02DC, 0x2122, 0x0161, 0x203A, 0x0153, 0x009D, 0x017E, 0x0178,
    ];
    const CP1251_HI: [u16; 64] = [
        0x0402, 0x0403, 0x201A, 0x0453, 0x201E, 0x2026, 0x2020, 0x2021, 0x20AC, 0x2030, 0x0409, 0x2039, 0x040A, 0x040C, 0x040B, 0x040F, 0x0452, 0x2018,
        0x2019, 0x201C, 0x201D, 0x2022, 0x2013, 0x2014, 0x0098, 0x2122, 0x0459, 0x203A, 0x045A, 0x045C, 0x045B, 0x045F, 0x00A0, 0x040E, 0x045E, 0x0408,
        0x00A4, 0x0490, 0x00A6, 0x00A7, 0x0401, 0x00A9, 0x0404, 0x00AB, 0x00AC, 0x00AD, 0x00AE, 0x0407, 0x00B0, 0x00B1, 0x0406, 0x0456, 0x0491, 0x00B5,
        0x00B6, 0x00B7, 0x0451, 0x2116, 0x0454, 0x00BB, 0x0458, 0x0405, 0x0455, 0x0457,
    ];

    /// Объявленная в `<?xml … encoding="…"?>` кодировка (нижний регистр).
    fn declared_encoding(b: &[u8]) -> Option<String> {
        let head = &b[..b.len().min(200)];
        let h = String::from_utf8_lossy(head).to_lowercase();
        let i = h.find("encoding")?;
        let rest = &h[i + 8..];
        let q = rest.find(['"', '\''])?;
        let quote = rest.as_bytes()[q] as char;
        let rest2 = &rest[q + 1..];
        Some(rest2[..rest2.find(quote)?].to_string())
    }

    /// Декодирует байты части (UTF-8 / UTF-16 с BOM / однобайтовые кодировки из объявления) в строку.
    pub fn decode(bytes: &[u8]) -> Result<std::borrow::Cow<'_, str>> {
        use std::borrow::Cow;
        if bytes.starts_with(&[0xFF, 0xFE]) || bytes.starts_with(&[0xFE, 0xFF]) {
            let le = bytes[0] == 0xFF;
            let units: Vec<u16> = bytes[2..]
                .chunks_exact(2)
                .map(|c| if le { u16::from_le_bytes([c[0], c[1]]) } else { u16::from_be_bytes([c[0], c[1]]) })
                .collect();
            return Ok(Cow::Owned(String::from_utf16(&units).map_err(|_| anyhow!("некорректная кодировка XML"))?));
        }
        let b = bytes.strip_prefix(&[0xEF, 0xBB, 0xBF]).unwrap_or(bytes);
        if let Ok(s) = std::str::from_utf8(b) {
            return Ok(Cow::Borrowed(s));
        }
        let enc = declared_encoding(b).unwrap_or_default();
        let table: Option<(&[u16], u8)> = match enc.as_str() {
            "iso-8859-1" | "latin1" | "latin-1" => Some((&[], 0)),
            "windows-1252" | "cp1252" | "iso-8859-15" => Some((&CP1252_HI, 0x80)),
            "windows-1251" | "cp1251" => Some((&CP1251_HI, 0x80)),
            _ => None,
        };
        let Some((tbl, start)) = table else {
            if enc.is_empty() {
                bail!("некорректная кодировка XML (ожидается UTF-8)");
            }
            bail!("неподдерживаемая кодировка XML «{enc}»")
        };
        let mut out = String::with_capacity(b.len() + b.len() / 8);
        for &x in b {
            let ch = if x < 0x80 {
                x as char
            } else if enc.contains("1251") {
                if x >= 0xC0 { char::from_u32(0x0410 + (x - 0xC0) as u32).unwrap_or('?') } else { char::from_u32(tbl[(x - start) as usize] as u32).unwrap_or('?') }
            } else if !tbl.is_empty() && x < 0xA0 {
                char::from_u32(tbl[(x - start) as usize] as u32).unwrap_or('?')
            } else {
                x as char
            };
            out.push(ch);
        }
        Ok(Cow::Owned(out))
    }

    pub fn unescape(s: &str) -> Result<String> {
        if !s.contains('&') {
            return Ok(s.to_string());
        }
        let mut out = String::with_capacity(s.len());
        let mut rest = s;
        while let Some(i) = rest.find('&') {
            out.push_str(&rest[..i]);
            rest = &rest[i + 1..];
            let j = rest.find(';').ok_or_else(|| anyhow!("некорректная ссылка & в XML"))?;
            let ent = &rest[..j];
            match ent {
                "amp" => out.push('&'),
                "lt" => out.push('<'),
                "gt" => out.push('>'),
                "quot" => out.push('"'),
                "apos" => out.push('\''),
                _ if ent.starts_with("#x") || ent.starts_with("#X") => {
                    let c = u32::from_str_radix(&ent[2..], 16).ok().and_then(char::from_u32);
                    out.push(c.ok_or_else(|| anyhow!("некорректный символ &{ent};"))?);
                }
                _ if ent.starts_with('#') => {
                    let c = ent[1..].parse::<u32>().ok().and_then(char::from_u32);
                    out.push(c.ok_or_else(|| anyhow!("некорректный символ &{ent};"))?);
                }
                _ => bail!("неизвестная сущность &{ent}; в XML"),
            }
            rest = &rest[j + 1..];
        }
        out.push_str(rest);
        Ok(out)
    }

    fn split_q(q: &str) -> (&str, &str) {
        match q.split_once(':') {
            Some((p, l)) => (p, l),
            None => ("", q),
        }
    }

    impl<'a> Parser<'a> {
        pub fn new(s: &'a str) -> Self {
            Parser { s, pos: 0, scopes: vec![], names: vec![] }
        }

        fn lookup(&self, prefix: &str) -> &str {
            for sc in self.scopes.iter().rev() {
                for (p, u) in sc.iter().rev() {
                    if p == prefix {
                        return u;
                    }
                }
            }
            if prefix == "xml" { NS_XML } else { "" }
        }

        pub fn next(&mut self) -> Result<Ev> {
            loop {
                let b = self.s.as_bytes();
                if self.pos >= b.len() {
                    if !self.scopes.is_empty() {
                        bail!("неожиданный конец XML");
                    }
                    return Ok(Ev::Eof);
                }
                let rest = &self.s[self.pos..];
                if !rest.starts_with('<') {
                    let end = rest.find('<').map_or(self.s.len(), |i| self.pos + i);
                    let raw = &self.s[self.pos..end];
                    self.pos = end;
                    if self.scopes.is_empty() {
                        if raw.trim().is_empty() {
                            continue;
                        }
                        bail!("текст вне корневого элемента XML");
                    }
                    return Ok(Ev::Text(unescape(raw)?));
                }
                if rest.starts_with("<?") {
                    let i = rest.find("?>").ok_or_else(|| anyhow!("незакрытая инструкция XML"))?;
                    self.pos += i + 2;
                    continue;
                }
                if rest.starts_with("<!--") {
                    let i = rest.find("-->").ok_or_else(|| anyhow!("незакрытый комментарий XML"))?;
                    self.pos += i + 3;
                    continue;
                }
                if rest.starts_with("<![CDATA[") {
                    let i = rest.find("]]>").ok_or_else(|| anyhow!("незакрытый CDATA"))?;
                    let t = rest[9..i].to_string();
                    self.pos += i + 3;
                    return Ok(Ev::Text(t));
                }
                if rest.starts_with("<!") {
                    // DOCTYPE: пропускаем (в том числе внутренний набор в [...]).
                    let mut depth = 0i32;
                    let mut end = None;
                    for (i, c) in rest.char_indices() {
                        match c {
                            '[' => depth += 1,
                            ']' => depth -= 1,
                            '>' if depth <= 0 => {
                                end = Some(i);
                                break;
                            }
                            _ => {}
                        }
                    }
                    self.pos += end.ok_or_else(|| anyhow!("незакрытый DOCTYPE"))? + 1;
                    continue;
                }
                if rest.starts_with("</") {
                    let i = rest.find('>').ok_or_else(|| anyhow!("незакрытый тег XML"))?;
                    let name = rest[2..i].trim_end();
                    self.pos += i + 1;
                    if self.scopes.pop().is_none() {
                        bail!("лишний закрывающий тег </{name}>");
                    }
                    match self.names.pop() {
                        Some(open) if open == name => {}
                        Some(open) => bail!("тег <{open}> закрыт как </{name}>"),
                        None => bail!("лишний закрывающий тег </{name}>"),
                    }
                    return Ok(Ev::End);
                }
                return self.start_tag();
            }
        }

        fn start_tag(&mut self) -> Result<Ev> {
            let s = self.s;
            let mut i = self.pos + 1;
            let bytes = s.as_bytes();
            let is_ws = |c: u8| c == b' ' || c == b'\t' || c == b'\n' || c == b'\r';
            let name_start = i;
            while i < bytes.len() && !is_ws(bytes[i]) && bytes[i] != b'/' && bytes[i] != b'>' {
                i += 1;
            }
            let qname = &s[name_start..i];
            if qname.is_empty() {
                bail!("пустое имя тега XML");
            }
            let mut raw: Vec<(String, String)> = vec![];
            let empty;
            loop {
                while i < bytes.len() && is_ws(bytes[i]) {
                    i += 1;
                }
                if i >= bytes.len() {
                    bail!("незакрытый тег <{qname}>");
                }
                if bytes[i] == b'>' {
                    empty = false;
                    i += 1;
                    break;
                }
                if bytes[i] == b'/' {
                    if bytes.get(i + 1) == Some(&b'>') {
                        empty = true;
                        i += 2;
                        break;
                    }
                    bail!("некорректный тег <{qname}>");
                }
                let an = i;
                while i < bytes.len() && !is_ws(bytes[i]) && bytes[i] != b'=' && bytes[i] != b'>' && bytes[i] != b'/' {
                    i += 1;
                }
                let aname = &s[an..i];
                while i < bytes.len() && is_ws(bytes[i]) {
                    i += 1;
                }
                if bytes.get(i) != Some(&b'=') {
                    bail!("атрибут {aname} без значения в <{qname}>");
                }
                i += 1;
                while i < bytes.len() && is_ws(bytes[i]) {
                    i += 1;
                }
                let q = *bytes.get(i).ok_or_else(|| anyhow!("обрыв тега"))?;
                if q != b'"' && q != b'\'' {
                    bail!("значение атрибута {aname} без кавычек");
                }
                let vs = i + 1;
                let ve = s[vs..].find(q as char).ok_or_else(|| anyhow!("незакрытое значение атрибута"))? + vs;
                raw.push((aname.to_string(), unescape(&s[vs..ve])?));
                i = ve + 1;
            }
            self.pos = i;
            let decls: Vec<(String, String)> = raw
                .iter()
                .filter_map(|(n, v)| {
                    if n == "xmlns" {
                        Some((String::new(), v.clone()))
                    } else {
                        n.strip_prefix("xmlns:").map(|p| (p.to_string(), v.clone()))
                    }
                })
                .collect();
            self.scopes.push(decls);
            if !empty {
                self.names.push(qname);
            }
            let (prefix, local) = split_q(qname);
            let mut e = Elem {
                ns: self.lookup(prefix).to_string(),
                prefix: prefix.to_string(),
                local: local.to_string(),
                attrs: Vec::with_capacity(raw.len()),
                kids: vec![],
            };
            for (n, v) in raw {
                if n == "xmlns" || n.starts_with("xmlns:") {
                    e.attrs.push(Attr { q: n, ns: String::new(), local: String::new(), v });
                } else {
                    let (p, l) = split_q(&n);
                    let ns = if p.is_empty() { String::new() } else { self.lookup(p).to_string() };
                    e.attrs.push(Attr { local: l.to_string(), ns, q: n, v });
                }
            }
            if empty {
                self.scopes.pop();
                Ok(Ev::Empty(e))
            } else {
                Ok(Ev::Start(e))
            }
        }

        /// Читает содержимое элемента, начало которого уже прочитано.
        pub fn tree(&mut self, start: Elem) -> Result<Elem> {
            let mut stack = vec![start];
            loop {
                match self.next()? {
                    Ev::Start(e) => {
                        if stack.len() > 400 {
                            bail!("слишком глубокая вложенность XML");
                        }
                        stack.push(e)
                    }
                    Ev::Empty(e) => stack.last_mut().unwrap().kids.push(Node::E(e)),
                    Ev::Text(t) => {
                        let top = stack.last_mut().unwrap();
                        match top.kids.last_mut() {
                            Some(Node::T(prev)) => prev.push_str(&t),
                            _ => top.kids.push(Node::T(t)),
                        }
                    }
                    Ev::End => {
                        let mut done = stack.pop().unwrap();
                        if done.kids.iter().any(|k| matches!(k, Node::E(_))) {
                            done.kids.retain(|k| !matches!(k, Node::T(t) if t.trim().is_empty()));
                        }
                        match stack.last_mut() {
                            Some(p) => p.kids.push(Node::E(done)),
                            None => return Ok(done),
                        }
                    }
                    Ev::Eof => bail!("неожиданный конец XML"),
                }
            }
        }
    }

    /// Разбирает весь документ и возвращает корневой элемент.
    pub fn parse(bytes: &[u8]) -> Result<Elem> {
        let s = decode(bytes)?;
        let mut p = Parser::new(&s);
        loop {
            match p.next()? {
                Ev::Start(e) => return p.tree(e),
                Ev::Empty(e) => return Ok(e),
                Ev::Eof => bail!("пустой XML-документ"),
                _ => {}
            }
        }
    }

    // ------------------------------------------------------------------ запись

    pub fn esc_text(s: &str, out: &mut String) {
        for c in s.chars() {
            match c {
                '&' => out.push_str("&amp;"),
                '<' => out.push_str("&lt;"),
                '>' => out.push_str("&gt;"),
                '\r' => out.push_str("&#13;"),
                c => out.push(c),
            }
        }
    }

    pub fn esc_attr(s: &str, out: &mut String) {
        for c in s.chars() {
            match c {
                '&' => out.push_str("&amp;"),
                '<' => out.push_str("&lt;"),
                '>' => out.push_str("&gt;"),
                '"' => out.push_str("&quot;"),
                '\n' => out.push_str("&#10;"),
                '\r' => out.push_str("&#13;"),
                '\t' => out.push_str("&#9;"),
                c => out.push(c),
            }
        }
    }

    fn lookup_scope<'a>(scope: &'a [(String, String)], prefix: &str) -> Option<&'a str> {
        scope.iter().rev().find(|(p, _)| p == prefix).map(|(_, u)| u.as_str())
    }

    /// Записывает начальный тег. `main` — пространство имён, элементы которого пишутся без префикса
    /// (пусто — не менять префиксы). Недостающие объявления префиксов (`xmlns:p`) добавляются сами,
    /// поэтому поддерево можно переносить из любого места документа.
    /// Возвращает (имя тега, действующее пространство по умолчанию).
    fn start_tag(e: &Elem, out: &mut String, parent_default: &str, main: &str, close: bool, scope: &mut Vec<(String, String)>) -> (String, String) {
        let is_main = !main.is_empty() && e.ns == main;
        let qname = if is_main || e.prefix.is_empty() { e.local.clone() } else { format!("{}:{}", e.prefix, e.local) };
        let required: &str = if is_main {
            main
        } else if e.prefix.is_empty() {
            &e.ns
        } else {
            parent_default
        };
        // Объявления на самом элементе.
        for a in &e.attrs {
            if let Some(p) = a.q.strip_prefix("xmlns:") {
                scope.push((p.to_string(), a.v.clone()));
            }
        }
        let mut extra: Vec<(String, String)> = vec![];
        let mut need = |prefix: &str, uri: &str, scope: &mut Vec<(String, String)>| {
            if prefix.is_empty() || prefix == "xml" || uri.is_empty() {
                return;
            }
            if lookup_scope(scope, prefix) != Some(uri) {
                scope.push((prefix.to_string(), uri.to_string()));
                extra.push((prefix.to_string(), uri.to_string()));
            }
        };
        if !is_main {
            need(&e.prefix, &e.ns, scope);
        }
        for a in &e.attrs {
            if a.q.starts_with("xmlns") {
                continue;
            }
            if let Some((p, _)) = a.q.split_once(':') {
                need(p, &a.ns, scope);
            }
        }
        out.push('<');
        out.push_str(&qname);
        if required != parent_default {
            out.push_str(" xmlns=\"");
            esc_attr(required, out);
            out.push('"');
        }
        for (p, u) in &extra {
            out.push_str(" xmlns:");
            out.push_str(p);
            out.push_str("=\"");
            esc_attr(u, out);
            out.push('"');
        }
        for a in &e.attrs {
            if a.q == "xmlns" {
                continue;
            }
            out.push(' ');
            out.push_str(&a.q);
            out.push_str("=\"");
            esc_attr(&a.v, out);
            out.push('"');
        }
        out.push_str(if close { "/>" } else { ">" });
        (qname, required.to_string())
    }

    fn write_rec(e: &Elem, out: &mut String, parent_default: &str, main: &str, scope: &mut Vec<(String, String)>) {
        let mark = scope.len();
        if e.kids.is_empty() {
            start_tag(e, out, parent_default, main, true, scope);
            scope.truncate(mark);
            return;
        }
        let (qname, def) = start_tag(e, out, parent_default, main, false, scope);
        for k in &e.kids {
            match k {
                Node::E(c) => write_rec(c, out, &def, main, scope),
                Node::T(t) => esc_text(t, out),
            }
        }
        out.push_str("</");
        out.push_str(&qname);
        out.push('>');
        scope.truncate(mark);
    }

    pub fn write_elem(e: &Elem, out: &mut String, parent_default: &str, main: &str) {
        write_rec(e, out, parent_default, main, &mut vec![]);
    }

    const DECL: &str = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n";

    /// Документ целиком (с XML-декларацией).
    pub fn doc(e: &Elem, main: &str) -> Vec<u8> {
        let mut s = String::from(DECL);
        write_elem(e, &mut s, "", main);
        s.into_bytes()
    }

    /// Элемент в виде строки (без декларации).
    pub fn to_string(e: &Elem, main: &str) -> String {
        let mut s = String::new();
        write_elem(e, &mut s, main, main);
        s
    }

    /// Потоковая запись: начало и конец тегов вручную, готовые поддеревья целиком.
    pub struct Sw<'w> {
        w: &'w mut dyn Write,
        buf: String,
        stack: Vec<(String, String, usize)>,
        scope: Vec<(String, String)>,
        main: &'static str,
    }

    impl<'w> Sw<'w> {
        pub fn new(w: &'w mut dyn Write, main: &'static str) -> Self {
            Sw { w, buf: String::from(DECL), stack: vec![], scope: vec![], main }
        }
        fn default(&self) -> String {
            self.stack.last().map(|x| x.1.clone()).unwrap_or_default()
        }
        pub fn start(&mut self, e: &Elem) {
            let d = self.default();
            let mark = self.scope.len();
            let r = start_tag(e, &mut self.buf, &d, self.main, false, &mut self.scope);
            self.stack.push((r.0, r.1, mark));
        }
        pub fn end(&mut self) -> std::io::Result<()> {
            if let Some((q, _, mark)) = self.stack.pop() {
                self.buf.push_str("</");
                self.buf.push_str(&q);
                self.buf.push('>');
                self.scope.truncate(mark);
            }
            self.flush_if_big()
        }
        pub fn elem(&mut self, e: &Elem) -> std::io::Result<()> {
            let d = self.default();
            write_rec(e, &mut self.buf, &d, self.main, &mut self.scope);
            self.flush_if_big()
        }
        fn flush_if_big(&mut self) -> std::io::Result<()> {
            if self.buf.len() > 1 << 16 {
                self.w.write_all(self.buf.as_bytes())?;
                self.buf.clear();
            }
            Ok(())
        }
        pub fn finish(mut self) -> std::io::Result<()> {
            while !self.stack.is_empty() {
                self.end()?;
            }
            self.w.write_all(self.buf.as_bytes())?;
            Ok(())
        }
    }
}

// ---------------------------------------------------------------------------
// Пакет OOXML: части, связи, типы содержимого
// ---------------------------------------------------------------------------

const CT_WORKSHEET: &str = "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml";
const CT_WORKBOOK: &str = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml";
const CT_STYLES: &str = "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml";
const CT_SST: &str = "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml";
const CT_RELS: &str = "application/vnd.openxmlformats-package.relationships+xml";
const CT_THEME: &str = "application/vnd.openxmlformats-officedocument.theme+xml";
const REL_OFFICE_DOC: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument";
const REL_BASE: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/";
const NS_MAIN_STRICT: &str = "http://purl.oclc.org/ooxml/spreadsheetml/main";

#[derive(Clone, Debug)]
struct Rel {
    id: String,
    typ: String,
    target: String,
    external: bool,
}

impl Rel {
    /// Последний сегмент типа связи: `worksheet`, `drawing`, `hyperlink`…
    fn kind(&self) -> &str {
        self.typ.rsplit('/').next().unwrap_or("")
    }
}

fn percent_decode(s: &str) -> String {
    if !s.contains('%') {
        return s.to_string();
    }
    let b = s.as_bytes();
    let hex = |c: u8| (c as char).to_digit(16);
    let mut out = Vec::with_capacity(b.len());
    let mut i = 0;
    while i < b.len() {
        if b[i] == b'%' && i + 2 < b.len() {
            if let (Some(h), Some(l)) = (hex(b[i + 1]), hex(b[i + 2])) {
                out.push((h * 16 + l) as u8);
                i += 3;
                continue;
            }
        }
        out.push(b[i]);
        i += 1;
    }
    String::from_utf8_lossy(&out).into_owned()
}

/// Нормализует путь части: без ведущего `/`, без `.`/`..`, с `/`.
fn norm_path(p: &str) -> String {
    let mut parts: Vec<&str> = vec![];
    for seg in p.split(['/', '\\']) {
        match seg {
            "" | "." => {}
            ".." => {
                parts.pop();
            }
            s => parts.push(s),
        }
    }
    parts.join("/")
}

/// Часть, на которую указывает связь `target` из части `base`.
fn resolve_target(base: &str, target: &str) -> String {
    let t = percent_decode(target);
    if t.starts_with('/') {
        return norm_path(&t);
    }
    let dir = base.rsplit_once('/').map_or("", |x| x.0);
    norm_path(&format!("{dir}/{t}"))
}

fn rels_path(part: &str) -> String {
    match part.rsplit_once('/') {
        Some((d, n)) => format!("{d}/_rels/{n}.rels"),
        None => format!("_rels/{part}.rels"),
    }
}

/// Относительный путь от части `from` к части `to` (для Target связи).
fn relative_target(from: &str, to: &str) -> String {
    let fd: Vec<&str> = from.split('/').collect();
    let fd = &fd[..fd.len() - 1];
    let td: Vec<&str> = to.split('/').collect();
    let mut i = 0;
    while i < fd.len() && i + 1 < td.len() && fd[i] == td[i] {
        i += 1;
    }
    let mut r = String::new();
    for _ in i..fd.len() {
        r.push_str("../");
    }
    r.push_str(&td[i..].join("/"));
    r
}

fn parse_rels(bytes: &[u8]) -> Result<Vec<Rel>> {
    let root = xml::parse(bytes)?;
    Ok(root
        .children()
        .filter(|e| e.local == "Relationship")
        .filter_map(|e| {
            Some(Rel {
                id: e.attr("Id")?.to_string(),
                typ: e.attr("Type")?.to_string(),
                target: e.attr("Target")?.to_string(),
                external: e.attr("TargetMode").is_some_and(|m| m.eq_ignore_ascii_case("External")),
            })
        })
        .collect())
}

fn rels_xml(rels: &[Rel]) -> Vec<u8> {
    let mut root = xml::Elem::new(xml::NS_PKG_REL, "Relationships");
    for r in rels {
        let mut e = xml::Elem::new(xml::NS_PKG_REL, "Relationship");
        e.set("Id", r.id.as_str());
        e.set("Type", r.typ.as_str());
        e.set("Target", r.target.as_str());
        if r.external {
            e.set("TargetMode", "External");
        }
        root.push(e);
    }
    xml::doc(&root, xml::NS_PKG_REL)
}

/// Читаемый пакет (zip) с таблицей типов содержимого.
struct Pkg {
    zip: zip::ZipArchive<std::fs::File>,
    names: HashMap<String, usize>,
    ct_default: HashMap<String, String>,
    ct_override: HashMap<String, String>,
}

/// Что лежит в файле по первым байтам.
enum Container {
    Zip,
    /// Составной документ OLE: старый .xls или зашифрованный пакет.
    Ole { encrypted: bool },
    Other,
}

fn file_label(path: &Path) -> String {
    path.file_name().map_or_else(|| path.display().to_string(), |n| n.to_string_lossy().into_owned())
}

fn sniff(path: &Path) -> Result<Container> {
    use std::io::Read;
    let label = file_label(path);
    let mut f = std::fs::File::open(path).map_err(|e| anyhow!("«{label}»: не удалось открыть файл ({})", io_reason(&e)))?;
    let mut head = [0u8; 8];
    let n = f.read(&mut head).map_err(|e| anyhow!("«{label}»: не удалось прочитать файл ({})", io_reason(&e)))?;
    if n == 0 {
        bail!("«{label}»: файл пустой");
    }
    if head.starts_with(b"PK") {
        return Ok(Container::Zip);
    }
    if head == [0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1] {
        // Зашифрованный пакет OOXML хранится в OLE-контейнере с потоком EncryptedPackage (UTF-16LE).
        let needle: Vec<u8> = "EncryptedPackage".encode_utf16().flat_map(|u| u.to_le_bytes()).collect();
        let mut encrypted = false;
        let mut buf = vec![0u8; 1 << 20];
        let mut tail: Vec<u8> = vec![];
        let _ = std::io::Seek::rewind(&mut f);
        loop {
            let k = f.read(&mut buf).unwrap_or(0);
            if k == 0 {
                break;
            }
            tail.extend_from_slice(&buf[..k]);
            if tail.windows(needle.len()).any(|w| w == needle.as_slice()) {
                encrypted = true;
                break;
            }
            let keep = needle.len() - 1;
            let cut = tail.len().saturating_sub(keep);
            tail.drain(..cut);
        }
        return Ok(Container::Ole { encrypted });
    }
    Ok(Container::Other)
}

fn io_reason(e: &std::io::Error) -> String {
    match e.kind() {
        std::io::ErrorKind::NotFound => "файл не найден".into(),
        std::io::ErrorKind::PermissionDenied => "нет доступа к файлу или он занят другой программой".into(),
        _ => e.to_string(),
    }
}

fn not_xlsx_error(path: &Path, c: &Container) -> anyhow::Error {
    let label = file_label(path);
    match c {
        Container::Ole { encrypted: true } => {
            anyhow!("«{label}»: файл защищён паролем. Снимите защиту в Excel и добавьте файл снова")
        }
        Container::Ole { encrypted: false } => {
            anyhow!("«{label}»: это не книга .xlsx (похоже на старый бинарный формат .xls); откройте файл в Excel и сохраните как .xlsx")
        }
        Container::Zip => anyhow!("«{label}»: это архив, но не книга Excel (.xlsx): нет xl/workbook.xml"),
        Container::Other => anyhow!("«{label}»: это не книга Excel (.xlsx): файл повреждён или имеет другой формат"),
    }
}

impl Pkg {
    fn open(path: &Path) -> Result<Pkg> {
        let c = sniff(path)?;
        if !matches!(c, Container::Zip) {
            return Err(not_xlsx_error(path, &c));
        }
        let label = file_label(path);
        let f = std::fs::File::open(path).map_err(|e| anyhow!("«{label}»: не удалось открыть файл ({})", io_reason(&e)))?;
        let mut zip = zip::ZipArchive::new(f)
            .map_err(|_| anyhow!("«{label}»: файл повреждён (не удаётся прочитать архив .xlsx)"))?;
        let mut names = HashMap::new();
        for i in 0..zip.len() {
            if let Ok(e) = zip.by_index_raw(i) {
                if e.is_dir() {
                    continue;
                }
                if let Ok(n) = e.name() {
                    names.insert(norm_path(&n).to_lowercase(), i);
                }
            }
        }
        let mut pkg = Pkg { zip, names, ct_default: HashMap::new(), ct_override: HashMap::new() };
        if pkg.has("[Content_Types].xml") {
            let b = pkg.read("[Content_Types].xml").map_err(|e| anyhow!("«{label}»: {e}"))?;
            let root = xml::parse(&b).map_err(|e| anyhow!("«{label}»: повреждён [Content_Types].xml ({e})"))?;
            for e in root.children() {
                match (e.local.as_str(), e.attr("ContentType")) {
                    ("Default", Some(ct)) => {
                        if let Some(x) = e.attr("Extension") {
                            pkg.ct_default.insert(x.to_lowercase(), ct.to_string());
                        }
                    }
                    ("Override", Some(ct)) => {
                        if let Some(p) = e.attr("PartName") {
                            pkg.ct_override.insert(norm_path(p).to_lowercase(), ct.to_string());
                        }
                    }
                    _ => {}
                }
            }
        }
        Ok(pkg)
    }

    fn has(&self, part: &str) -> bool {
        self.names.contains_key(&part.to_lowercase())
    }

    fn read(&mut self, part: &str) -> Result<Vec<u8>> {
        use std::io::Read;
        let i = *self.names.get(&part.to_lowercase()).ok_or_else(|| anyhow!("в файле нет части {part}"))?;
        let mut e = self.zip.by_index(i).map_err(|_| anyhow!("часть {part} повреждена"))?;
        let mut v = Vec::with_capacity(e.size().min(1 << 28) as usize);
        e.read_to_end(&mut v).map_err(|_| anyhow!("часть {part} повреждена (ошибка распаковки или контрольной суммы)"))?;
        Ok(v)
    }

    fn rels(&mut self, part: &str) -> Result<Vec<Rel>> {
        let p = rels_path(part);
        if !self.has(&p) {
            return Ok(vec![]);
        }
        let b = self.read(&p)?;
        parse_rels(&b).map_err(|e| anyhow!("повреждены связи {p} ({e})"))
    }

    fn content_type(&self, part: &str) -> Option<String> {
        let lp = part.to_lowercase();
        if let Some(c) = self.ct_override.get(&lp) {
            return Some(c.clone());
        }
        let ext = lp.rsplit_once('.')?.1;
        self.ct_default.get(ext).cloned()
    }
}

/// Лист исходной книги (запись `<sheet>` из workbook.xml).
#[derive(Clone, Debug)]
struct SheetEntry {
    name: String,
    state: String,
    kind: String,
    part: String,
}

impl SheetEntry {
    fn is_worksheet(&self) -> bool {
        self.kind == "worksheet"
    }
    fn hidden(&self) -> bool {
        self.state == "hidden" || self.state == "veryHidden"
    }
}

/// Исходная книга: пакет и разобранный workbook.xml.
struct Book {
    pkg: Pkg,
    label: String,
    wb_part: String,
    wb_rels: Vec<Rel>,
    wb: xml::Elem,
    sheets: Vec<SheetEntry>,
}

impl Book {
    fn open(path: &Path) -> Result<Book> {
        let label = file_label(path);
        let mut pkg = Pkg::open(path)?;
        let mut wb_part = String::from("xl/workbook.xml");
        if pkg.has("_rels/.rels") {
            if let Ok(b) = pkg.read("_rels/.rels") {
                if let Ok(rels) = parse_rels(&b) {
                    if let Some(r) = rels.iter().find(|r| r.typ == REL_OFFICE_DOC && !r.external) {
                        wb_part = norm_path(&percent_decode(&r.target));
                    }
                }
            }
        }
        if !pkg.has(&wb_part) {
            return Err(not_xlsx_error(path, &Container::Zip));
        }
        let b = pkg.read(&wb_part).map_err(|e| anyhow!("«{label}»: {e}"))?;
        let wb = xml::parse(&b).map_err(|e| anyhow!("«{label}»: повреждён workbook.xml ({e})"))?;
        if wb.ns == NS_MAIN_STRICT {
            bail!("«{label}»: формат Strict Open XML не поддерживается; сохраните книгу в Excel как обычную «Книга Excel (*.xlsx)»");
        }
        if wb.local != "workbook" || wb.ns != xml::NS_MAIN {
            bail!("«{label}»: это не книга Excel (.xlsx): неожиданное содержимое workbook.xml");
        }
        let wb_rels = pkg.rels(&wb_part).map_err(|e| anyhow!("«{label}»: {e}"))?;
        let mut sheets = vec![];
        if let Some(ss) = wb.child("sheets") {
            for s in ss.children().filter(|e| e.local == "sheet") {
                let (Some(name), Some(rid)) = (s.attr("name"), s.attr_ns(xml::NS_R, "id")) else { continue };
                let (kind, part) = match wb_rels.iter().find(|r| r.id == rid && !r.external) {
                    Some(rel) => (rel.kind().to_string(), resolve_target(&wb_part, &rel.target)),
                    None => ("missing".to_string(), String::new()),
                };
                sheets.push(SheetEntry { name: name.to_string(), state: s.attr("state").unwrap_or("visible").to_string(), kind, part });
            }
        }
        Ok(Book { pkg, label, wb_part, wb_rels, wb, sheets })
    }

    fn date1904(&self) -> bool {
        self.wb
            .child("workbookPr")
            .and_then(|p| p.attr("date1904"))
            .is_some_and(|v| v == "1" || v.eq_ignore_ascii_case("true"))
    }

    fn rel_of_kind(&self, kind: &str) -> Option<&Rel> {
        self.wb_rels.iter().find(|r| r.kind() == kind && !r.external)
    }
}

// ---------------------------------------------------------------------------
// Переписывание ссылок в формулах
// ---------------------------------------------------------------------------

/// Как поступать со ссылками на листы, которых нет в результате.
#[derive(Clone, Copy, PartialEq)]
enum Missing {
    /// Оставить текст как есть (вызывающий решает, что делать).
    Keep,
    /// Заменить ссылку на `#REF!` (так делает сам Excel при удалении листа).
    Ref,
}

/// Правила переписывания формул одной исходной книги.
#[derive(Default)]
struct FormulaCtx {
    /// Имя листа источника (в нижнем регистре) → (новое имя, позиция в результате).
    sheets: HashMap<String, (String, usize)>,
    /// Все листы источника в порядке книги (имена в нижнем регистре) — для ссылок 3D.
    order: Vec<String>,
    /// Смещение номеров внешних книг `[n]` и их число в исходной книге.
    ext_offset: usize,
    ext_count: usize,
    /// Переименованные имена уровня книги (нижний регистр → новое имя).
    names: HashMap<String, String>,
    /// Таблицы источника: Some(новое имя) или None, если таблица не попадёт в результат.
    tables: HashMap<String, Option<String>>,
}

struct Rewritten {
    text: String,
    /// Листы/таблицы, на которые есть ссылки, но которых нет в результате.
    missing: Vec<String>,
}

fn is_ident_char(c: char) -> bool {
    c.is_alphanumeric() || c == '_' || c == '.'
}

fn is_ident_start(c: char) -> bool {
    c.is_alphabetic() || c == '_' || c == '\\'
}

fn quote_sheet(n: &str) -> String {
    format!("'{}'", n.replace('\'', "''"))
}

/// Индекс символа после `]`, парной к `[` в позиции `i` (в ссылках на таблицы `'` экранирует символ).
fn bracket_end(ch: &[char], i: usize) -> usize {
    let mut depth = 0i32;
    let mut j = i;
    while j < ch.len() {
        match ch[j] {
            '\'' => j += 1,
            '[' => depth += 1,
            ']' => {
                depth -= 1;
                if depth == 0 {
                    return j + 1;
                }
            }
            _ => {}
        }
        j += 1;
    }
    ch.len()
}

/// Пропускает ссылку на ячейки/диапазон после префикса листа (`$A$1:$B$2`).
fn skip_ref(ch: &[char], mut j: usize) -> usize {
    let run = |mut k: usize| {
        while k < ch.len() && (ch[k].is_alphanumeric() || ch[k] == '$' || ch[k] == '_' || ch[k] == '.') {
            k += 1;
        }
        k
    };
    j = run(j);
    if j < ch.len() && ch[j] == ':' {
        let k = run(j + 1);
        if k > j + 1 {
            j = k;
        }
    }
    j
}

impl FormulaCtx {
    /// Префикс листа (`name` без `!`); None — ссылку нужно заменить на `#REF!`.
    fn prefix(&self, name: &str, quoted: bool, orig: &str, mode: Missing, missing: &mut Vec<String>) -> Option<String> {
        // Внешняя книга: [n]Лист
        if let Some(rest) = name.strip_prefix('[') {
            if let Some((num, sheet)) = rest.split_once(']') {
                if let Ok(k) = num.parse::<usize>() {
                    if k == 0 || k > self.ext_count {
                        missing.push(format!("[{k}]{sheet}"));
                        return if mode == Missing::Keep { Some(orig.to_string()) } else { None };
                    }
                    let nn = format!("[{}]{}", k + self.ext_offset, sheet);
                    return Some(if quoted { format!("{}!", quote_sheet(&nn)) } else { format!("{nn}!") });
                }
            }
        }
        if let Some((a, b)) = name.split_once(':') {
            let (la, lb) = (a.to_lowercase(), b.to_lowercase());
            let pa = self.order.iter().position(|x| *x == la);
            let pb = self.order.iter().position(|x| *x == lb);
            if let (Some(pa), Some(pb), Some((na, ia)), Some((nb, _))) = (pa, pb, self.sheets.get(&la), self.sheets.get(&lb)) {
                let (lo, hi) = (pa.min(pb), pa.max(pb));
                let base = self.sheets.get(&self.order[lo]).map(|x| x.1);
                let contiguous = base.is_some_and(|base| {
                    (lo..=hi).all(|k| self.sheets.get(&self.order[k]).is_some_and(|x| x.1 == base + (k - lo)))
                });
                let _ = ia;
                if contiguous {
                    if na == a && nb == b {
                        return Some(orig.to_string());
                    }
                    return Some(format!("{}!", quote_sheet(&format!("{na}:{nb}"))));
                }
            }
            missing.push(name.to_string());
            return if mode == Missing::Keep { Some(orig.to_string()) } else { None };
        }
        match self.sheets.get(&name.to_lowercase()) {
            Some((new, _)) => {
                if new == name {
                    Some(orig.to_string())
                } else {
                    Some(format!("{}!", quote_sheet(new)))
                }
            }
            None => {
                missing.push(name.to_string());
                if mode == Missing::Keep { Some(orig.to_string()) } else { None }
            }
        }
    }

    fn rewrite(&self, f: &str, mode: Missing) -> Rewritten {
        let ch: Vec<char> = f.chars().collect();
        let n = ch.len();
        let mut out = String::with_capacity(f.len() + 8);
        let mut missing: Vec<String> = vec![];
        let mut i = 0;
        let mut skip_ident = false;
        let push = |out: &mut String, a: usize, b: usize| out.extend(&ch[a..b.min(n)]);
        while i < n {
            let c = ch[i];
            // Строковый литерал.
            if c == '"' {
                let mut j = i + 1;
                while j < n {
                    if ch[j] == '"' {
                        if j + 1 < n && ch[j + 1] == '"' {
                            j += 2;
                            continue;
                        }
                        break;
                    }
                    j += 1;
                }
                push(&mut out, i, j + 1);
                i = (j + 1).min(n);
                skip_ident = false;
                continue;
            }
            // Ошибки: #REF!, #N/A …
            if c == '#' {
                let mut j = i + 1;
                while j < n && (ch[j].is_ascii_alphanumeric() || ch[j] == '/') {
                    j += 1;
                }
                if j < n && (ch[j] == '!' || ch[j] == '?') {
                    j += 1;
                }
                push(&mut out, i, j);
                i = j;
                continue;
            }
            // Имя листа в апострофах.
            if c == '\'' {
                let mut j = i + 1;
                let mut name = String::new();
                while j < n {
                    if ch[j] == '\'' {
                        if j + 1 < n && ch[j + 1] == '\'' {
                            name.push('\'');
                            j += 2;
                            continue;
                        }
                        break;
                    }
                    name.push(ch[j]);
                    j += 1;
                }
                if j + 1 < n && ch[j + 1] == '!' {
                    let end = j + 2;
                    let orig: String = ch[i..end].iter().collect();
                    match self.prefix(&name, true, &orig, mode, &mut missing) {
                        Some(t) => {
                            out.push_str(&t);
                            i = end;
                            skip_ident = true;
                        }
                        None => {
                            out.push_str("#REF!");
                            i = skip_ref(&ch, end);
                        }
                    }
                } else {
                    push(&mut out, i, j + 1);
                    i = (j + 1).min(n);
                }
                continue;
            }
            // Внешняя книга [n]Лист! или ссылка на таблицу без имени.
            if c == '[' {
                let end = bracket_end(&ch, i);
                let inner: String = ch[i + 1..end.saturating_sub(1).max(i + 1)].iter().collect();
                if !inner.is_empty() && inner.chars().all(|d| d.is_ascii_digit()) {
                    let mut k = end;
                    while k < n && is_ident_char(ch[k]) {
                        k += 1;
                    }
                    if k < n && ch[k] == '!' {
                        let name: String = ch[i..k].iter().collect();
                        let orig: String = ch[i..=k].iter().collect();
                        if let Some(t) = self.prefix(&name, false, &orig, mode, &mut missing) {
                            out.push_str(&t);
                        }
                        i = k + 1;
                        skip_ident = true;
                        continue;
                    }
                }
                push(&mut out, i, end);
                i = end;
                continue;
            }
            // Числа (чтобы `1E5` не принять за имя).
            if c.is_ascii_digit() || (c == '.' && i + 1 < n && ch[i + 1].is_ascii_digit()) {
                let mut j = i;
                while j < n && (ch[j].is_ascii_digit() || ch[j] == '.') {
                    j += 1;
                }
                if j < n && (ch[j] == 'e' || ch[j] == 'E') {
                    let mut k = j + 1;
                    if k < n && (ch[k] == '+' || ch[k] == '-') {
                        k += 1;
                    }
                    if k < n && ch[k].is_ascii_digit() {
                        while k < n && ch[k].is_ascii_digit() {
                            k += 1;
                        }
                        j = k;
                    }
                }
                push(&mut out, i, j);
                i = j;
                skip_ident = false;
                continue;
            }
            if is_ident_start(c) {
                let mut j = i;
                while j < n && is_ident_char(ch[j]) {
                    j += 1;
                }
                // Имя листа без апострофов: Лист1!A1
                if j < n && ch[j] == '!' {
                    let name: String = ch[i..j].iter().collect();
                    let orig: String = ch[i..=j].iter().collect();
                    match self.prefix(&name, false, &orig, mode, &mut missing) {
                        Some(t) => {
                            out.push_str(&t);
                            i = j + 1;
                            skip_ident = true;
                        }
                        None => {
                            out.push_str("#REF!");
                            i = skip_ref(&ch, j + 1);
                        }
                    }
                    continue;
                }
                // Ссылка 3D без апострофов: Лист1:Лист3!A1
                if j < n && ch[j] == ':' {
                    let mut k = j + 1;
                    while k < n && is_ident_char(ch[k]) {
                        k += 1;
                    }
                    if k > j + 1 && k < n && ch[k] == '!' {
                        let name: String = ch[i..k].iter().collect();
                        let orig: String = ch[i..=k].iter().collect();
                        match self.prefix(&name, false, &orig, mode, &mut missing) {
                            Some(t) => {
                                out.push_str(&t);
                                i = k + 1;
                                skip_ident = true;
                            }
                            None => {
                                out.push_str("#REF!");
                                i = skip_ref(&ch, k + 1);
                            }
                        }
                        continue;
                    }
                }
                let run: String = ch[i..j].iter().collect();
                if skip_ident {
                    skip_ident = false;
                    out.push_str(&run);
                    i = j;
                    continue;
                }
                if j < n && ch[j] == '(' {
                    out.push_str(&run);
                    i = j;
                    continue;
                }
                if j < n && ch[j] == '[' {
                    // Таблица: Таблица1[Столбец]
                    let end = bracket_end(&ch, j);
                    match self.tables.get(&run.to_lowercase()) {
                        Some(Some(new)) => {
                            out.push_str(new);
                            push(&mut out, j, end);
                        }
                        Some(None) => {
                            missing.push(run.clone());
                            if mode == Missing::Ref {
                                out.push_str("#REF!");
                            } else {
                                out.push_str(&run);
                                push(&mut out, j, end);
                            }
                        }
                        None => {
                            out.push_str(&run);
                            push(&mut out, j, end);
                        }
                    }
                    i = end;
                    continue;
                }
                match self.names.get(&run.to_lowercase()) {
                    Some(new) => out.push_str(new),
                    None => out.push_str(&run),
                }
                i = j;
                continue;
            }
            if c != '$' {
                skip_ident = false;
            }
            out.push(c);
            i += 1;
        }
        missing.sort();
        missing.dedup();
        Rewritten { text: out, missing }
    }
}

// ---------------------------------------------------------------------------
// Тема и стили
// ---------------------------------------------------------------------------

#[derive(Clone, PartialEq, Debug)]
struct Theme {
    /// В порядке индексов `theme=` в SpreadsheetML: lt1, dk1, lt2, dk2, accent1..6, hlink, folHlink.
    colors: Vec<[u8; 3]>,
    major: String,
    minor: String,
}

fn parse_theme(bytes: &[u8]) -> Option<Theme> {
    let root = xml::parse(bytes).ok()?;
    let mut colors = vec![];
    let mut major = String::new();
    let mut minor = String::new();
    root.walk(&mut |e| {
        if e.local == "clrScheme" {
            for name in ["lt1", "dk1", "lt2", "dk2", "accent1", "accent2", "accent3", "accent4", "accent5", "accent6", "hlink", "folHlink"] {
                let v = e
                    .child(name)
                    .and_then(|c| c.children().next())
                    .and_then(|c| c.attr("val").filter(|v| v.len() == 6 && c.local == "srgbClr").or_else(|| c.attr("lastClr")))
                    .and_then(|h| u32::from_str_radix(h, 16).ok())
                    .map_or([0, 0, 0], |h| [(h >> 16) as u8, (h >> 8) as u8, h as u8]);
                colors.push(v);
            }
        }
        if e.local == "majorFont" {
            major = e.child("latin").and_then(|l| l.attr("typeface")).unwrap_or("").to_string();
        }
        if e.local == "minorFont" {
            minor = e.child("latin").and_then(|l| l.attr("typeface")).unwrap_or("").to_string();
        }
    });
    if colors.len() == 12 { Some(Theme { colors, major, minor }) } else { None }
}

fn rgb_to_hls([r, g, b]: [u8; 3]) -> (f64, f64, f64) {
    let (r, g, b) = (r as f64 / 255.0, g as f64 / 255.0, b as f64 / 255.0);
    let max = r.max(g).max(b);
    let min = r.min(g).min(b);
    let l = (max + min) / 2.0;
    if (max - min).abs() < 1e-12 {
        return (0.0, l, 0.0);
    }
    let d = max - min;
    let s = if l > 0.5 { d / (2.0 - max - min) } else { d / (max + min) };
    let h = if (max - r).abs() < 1e-12 {
        ((g - b) / d + if g < b { 6.0 } else { 0.0 }) / 6.0
    } else if (max - g).abs() < 1e-12 {
        ((b - r) / d + 2.0) / 6.0
    } else {
        ((r - g) / d + 4.0) / 6.0
    };
    (h, l, s)
}

fn hls_to_rgb(h: f64, l: f64, s: f64) -> [u8; 3] {
    if s.abs() < 1e-12 {
        let v = (l * 255.0).round() as u8;
        return [v, v, v];
    }
    let q = if l < 0.5 { l * (1.0 + s) } else { l + s - l * s };
    let p = 2.0 * l - q;
    let f = |mut t: f64| {
        if t < 0.0 {
            t += 1.0;
        }
        if t > 1.0 {
            t -= 1.0;
        }
        let v = if t < 1.0 / 6.0 {
            p + (q - p) * 6.0 * t
        } else if t < 0.5 {
            q
        } else if t < 2.0 / 3.0 {
            p + (q - p) * (2.0 / 3.0 - t) * 6.0
        } else {
            p
        };
        (v * 255.0).round().clamp(0.0, 255.0) as u8
    };
    [f(h + 1.0 / 3.0), f(h), f(h - 1.0 / 3.0)]
}

fn apply_tint(rgb: [u8; 3], tint: f64) -> [u8; 3] {
    if tint.abs() < 1e-9 {
        return rgb;
    }
    let (h, l, s) = rgb_to_hls(rgb);
    let l = if tint < 0.0 { l * (1.0 + tint) } else { l * (1.0 - tint) + tint };
    hls_to_rgb(h, l.clamp(0.0, 1.0), s)
}

/// Что нужно «запечь» при переносе стилей из книги с другой темой.
struct ThemeFix {
    src: Theme,
    colors: bool,
    fonts: bool,
}

impl ThemeFix {
    /// Заменяет цвета темы на RGB и (при другой теме шрифтов) шрифты темы на явные.
    fn apply(&self, e: &mut xml::Elem) {
        let colors = self.colors;
        let src = &self.src;
        e.walk_mut(&mut |x| {
            if !colors || x.ns != xml::NS_MAIN {
                return;
            }
            if let Some(t) = x.attr("theme").and_then(|t| t.parse::<usize>().ok()) {
                if let Some(base) = src.colors.get(t) {
                    let tint = x.attr("tint").and_then(|v| v.parse::<f64>().ok()).unwrap_or(0.0);
                    let c = apply_tint(*base, tint);
                    x.remove("theme");
                    x.remove("tint");
                    x.set("rgb", format!("FF{:02X}{:02X}{:02X}", c[0], c[1], c[2]));
                }
            }
        });
        if self.fonts && e.local == "font" {
            let scheme = e.child("scheme").and_then(|s| s.attr("val")).map(str::to_string);
            if let Some(sc) = scheme {
                let face = if sc == "major" { &self.src.major } else { &self.src.minor };
                if !face.is_empty() {
                    if let Some(nm) = e.child_mut("name") {
                        nm.set("val", face.as_str());
                    }
                    e.kids.retain(|k| !matches!(k, xml::Node::E(c) if c.local == "scheme"));
                }
            }
        }
    }
}

fn is_date_format(id: u32, code: Option<&str>) -> bool {
    match id {
        14..=17 | 22 | 27..=36 | 50..=58 => return true,
        0..=13 | 18..=21 | 37..=49 => return false,
        _ => {}
    }
    let Some(code) = code else { return false };
    // Убираем литералы в кавычках, экранированные символы и секции в [] (цвета, условия, локаль).
    let mut s = String::new();
    let mut in_q = false;
    let mut in_b = false;
    let mut esc = false;
    for c in code.chars() {
        if esc {
            esc = false;
            continue;
        }
        match c {
            '\\' | '_' | '*' if !in_q && !in_b => esc = true,
            '"' => in_q = !in_q,
            '[' if !in_q => in_b = true,
            ']' if !in_q => in_b = false,
            c if !in_q && !in_b => s.push(c.to_ascii_lowercase()),
            _ => {}
        }
    }
    s.contains('y') || s.contains('d') || (s.contains('m') && !s.contains('h') && !s.contains('s'))
}

#[derive(Default)]
struct Interner {
    items: Vec<String>,
    idx: HashMap<String, usize>,
}

impl Interner {
    fn add(&mut self, s: String) -> usize {
        if let Some(&i) = self.idx.get(&s) {
            return i;
        }
        self.items.push(s.clone());
        self.idx.insert(s, self.items.len() - 1);
        self.items.len() - 1
    }
}

/// Соответствие индексов стилей исходной книги индексам результата.
#[derive(Default)]
struct StyleMap {
    xf: Vec<u32>,
    dxf: Vec<u32>,
    numfmt: HashMap<u32, u32>,
    /// xf исходной книги с форматом даты.
    date_xf: Vec<bool>,
}

impl StyleMap {
    fn xf_for(&self, s: Option<&str>) -> u32 {
        let i = s.and_then(|v| v.parse::<usize>().ok()).unwrap_or(0);
        self.xf.get(i).or_else(|| self.xf.first()).copied().unwrap_or(0)
    }
    fn is_date(&self, s: Option<&str>) -> bool {
        let i = s.and_then(|v| v.parse::<usize>().ok()).unwrap_or(0);
        self.date_xf.get(i).copied().unwrap_or(false)
    }
    /// Новый номер dxf; None, если в исходной книге такого формата нет (повреждённая ссылка).
    fn dxf_for(&self, id: u32) -> Option<u32> {
        self.dxf.get(id as usize).copied()
    }
    fn numfmt_for(&self, id: u32) -> u32 {
        self.numfmt.get(&id).copied().unwrap_or(id)
    }
}

/// Накопитель общих стилей результата.
#[derive(Default)]
struct StyleOut {
    numfmt_by_code: HashMap<String, u32>,
    numfmt_ids: HashMap<u32, String>,
    numfmts: Vec<(u32, String)>,
    fonts: Interner,
    fills: Interner,
    borders: Interner,
    csxfs: Interner,
    xfs: Interner,
    dxfs: Interner,
    /// ключ стиля → (xfId в результате)
    style_keys: HashMap<String, u32>,
    cell_styles: Vec<String>,
    table_styles: Vec<(String, String)>,
    default_table_style: Option<(String, String)>,
    indexed_colors: Option<String>,
    started: bool,
}

fn ser(e: &xml::Elem) -> String {
    xml::to_string(e, xml::NS_MAIN)
}

fn default_styles_root() -> xml::Elem {
    let parse = |s: &str| xml::parse(s.as_bytes()).unwrap();
    parse(&format!(
        "<styleSheet xmlns=\"{}\"><fonts count=\"1\"><font><sz val=\"11\"/><name val=\"Calibri\"/><family val=\"2\"/></font></fonts>\
         <fills count=\"2\"><fill><patternFill patternType=\"none\"/></fill><fill><patternFill patternType=\"gray125\"/></fill></fills>\
         <borders count=\"1\"><border><left/><right/><top/><bottom/><diagonal/></border></borders>\
         <cellStyleXfs count=\"1\"><xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\"/></cellStyleXfs>\
         <cellXfs count=\"1\"><xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\"/></cellXfs>\
         <cellStyles count=\"1\"><cellStyle name=\"Normal\" xfId=\"0\" builtinId=\"0\"/></cellStyles></styleSheet>",
        xml::NS_MAIN
    ))
}

impl StyleOut {
    fn map_numfmt(&mut self, id: u32, code: &str) -> u32 {
        if let Some(&o) = self.numfmt_by_code.get(code) {
            return o;
        }
        let new = if id < 164 && !self.numfmt_ids.contains_key(&id) {
            id
        } else {
            let mut n = 164;
            while self.numfmt_ids.contains_key(&n) {
                n += 1;
            }
            n
        };
        self.numfmt_by_code.insert(code.to_string(), new);
        self.numfmt_ids.insert(new, code.to_string());
        self.numfmts.push((new, code.to_string()));
        new
    }

    /// Добавляет стили источника; возвращает таблицы соответствия.
    fn add_source(&mut self, root: &xml::Elem, fix: Option<&ThemeFix>, warns: &mut Warnings, label: &str) -> Result<StyleMap> {
        let mut map = StyleMap::default();
        let mut codes: HashMap<u32, String> = HashMap::new();
        if !self.started {
            self.fills.add("<fill><patternFill patternType=\"none\"/></fill>".into());
            self.fills.add("<fill><patternFill patternType=\"gray125\"/></fill>".into());
        }
        if let Some(nf) = root.child("numFmts") {
            for e in nf.children().filter(|e| e.local == "numFmt") {
                if let (Some(id), Some(code)) = (e.attr("numFmtId").and_then(|v| v.parse::<u32>().ok()), e.attr("formatCode")) {
                    codes.insert(id, code.to_string());
                    let o = self.map_numfmt(id, code);
                    map.numfmt.insert(id, o);
                }
            }
        }
        let prep = |e: &xml::Elem| -> xml::Elem {
            let mut c = e.clone();
            c.kids.retain(|k| !matches!(k, xml::Node::E(x) if x.local == "extLst"));
            if let Some(f) = fix {
                f.apply(&mut c);
            }
            c
        };
        let list = |name: &str| -> Vec<xml::Elem> { root.child(name).map(|c| c.children().cloned().collect()).unwrap_or_default() };
        let fonts: Vec<usize> = list("fonts").iter().map(|e| self.fonts.add(ser(&prep(e)))).collect();
        let fills: Vec<usize> = list("fills").iter().map(|e| self.fills.add(ser(&prep(e)))).collect();
        let borders: Vec<usize> = list("borders").iter().map(|e| self.borders.add(ser(&prep(e)))).collect();

        let remap_xf = |e: &mut xml::Elem, map: &StyleMap| {
            if let Some(v) = e.attr("numFmtId").and_then(|v| v.parse::<u32>().ok()) {
                e.set("numFmtId", map.numfmt_for(v).to_string());
            }
            for (attr, tbl) in [("fontId", &fonts), ("fillId", &fills), ("borderId", &borders)] {
                if let Some(v) = e.attr(attr).and_then(|v| v.parse::<usize>().ok()) {
                    e.set(attr, tbl.get(v).copied().unwrap_or(0).to_string());
                }
            }
        };

        // Именованные стили и их базовые xf.
        let csx_src = list("cellStyleXfs");
        let mut csx_map: Vec<Option<u32>> = vec![None; csx_src.len()];
        if let Some(cs) = root.child("cellStyles") {
            for st in cs.children().filter(|e| e.local == "cellStyle") {
                let Some(xid) = st.attr("xfId").and_then(|v| v.parse::<usize>().ok()) else { continue };
                if xid >= csx_src.len() {
                    continue;
                }
                let name_key = format!("n:{}", st.attr("name").unwrap_or("").to_lowercase());
                let builtin_key = match st.attr("builtinId") {
                    Some(b) if st.attr("customBuiltin") != Some("1") => Some(format!("b{b}:{}", st.attr("iLevel").unwrap_or(""))),
                    _ => None,
                };
                let existing = builtin_key.as_ref().and_then(|k| self.style_keys.get(k)).or_else(|| self.style_keys.get(&name_key)).copied();
                if let Some(existing) = existing {
                    csx_map[xid].get_or_insert(existing);
                    if let Some(k) = builtin_key {
                        self.style_keys.entry(k).or_insert(existing);
                    }
                    self.style_keys.entry(name_key).or_insert(existing);
                    continue;
                }
                if csx_map[xid].is_some() {
                    continue;
                }
                let mut xf = prep(&csx_src[xid]);
                remap_xf(&mut xf, &map);
                let idx = self.push_csxf(ser(&xf));
                if let Some(k) = builtin_key {
                    self.style_keys.insert(k, idx);
                }
                self.style_keys.insert(name_key, idx);
                csx_map[xid] = Some(idx);
                let mut st2 = st.clone();
                st2.kids.clear();
                st2.set("xfId", idx.to_string());
                self.cell_styles.push(ser(&st2));
            }
        }
        for (i, xf) in csx_src.iter().enumerate() {
            if csx_map[i].is_none() {
                let mut x = prep(xf);
                remap_xf(&mut x, &map);
                let idx = self.push_csxf_shared(ser(&x));
                csx_map[i] = Some(idx);
            }
        }

        for xf in list("cellXfs") {
            let mut x = prep(&xf);
            remap_xf(&mut x, &map);
            let old_id = xf.attr("numFmtId").and_then(|v| v.parse::<u32>().ok()).unwrap_or(0);
            map.date_xf.push(is_date_format(old_id, codes.get(&old_id).map(String::as_str)));
            let xid = x.attr("xfId").and_then(|v| v.parse::<usize>().ok()).unwrap_or(0);
            let new_xid = csx_map.get(xid).copied().flatten().or_else(|| csx_map.first().copied().flatten()).unwrap_or(0);
            x.set("xfId", new_xid.to_string());
            map.xf.push(self.xfs.add(ser(&x)) as u32);
        }
        if map.xf.is_empty() {
            // В книге нет формата ячеек по умолчанию — берём «пустой» (первый шрифт, без заливки и границ).
            let font0 = fonts.first().copied().unwrap_or(0);
            let borders0 = borders.first().copied().unwrap_or(0);
            let xid = csx_map.first().copied().flatten().unwrap_or(0);
            map.date_xf.push(false);
            map.xf.push(self.xfs.add(format!("<xf numFmtId=\"0\" fontId=\"{font0}\" fillId=\"0\" borderId=\"{borders0}\" xfId=\"{xid}\"/>")) as u32);
        }
        for e in list("dxfs") {
            let mut d = prep(&e);
            d.walk_mut(&mut |x| {
                if x.local == "numFmt" && x.ns == xml::NS_MAIN {
                    if let (Some(id), Some(code)) = (x.attr("numFmtId").and_then(|v| v.parse::<u32>().ok()), x.attr("formatCode").map(str::to_string)) {
                        let o = self.map_numfmt(id, &code);
                        x.set("numFmtId", o.to_string());
                    }
                }
            });
            map.dxf.push(self.dxfs.add(ser(&d)) as u32);
        }
        if let Some(ts) = root.child("tableStyles") {
            if !self.started {
                self.default_table_style =
                    Some((ts.attr("defaultTableStyle").unwrap_or("TableStyleMedium2").to_string(), ts.attr("defaultPivotStyle").unwrap_or("PivotStyleLight16").to_string()));
            }
            for t in ts.children().filter(|e| e.local == "tableStyle") {
                let name = t.attr("name").unwrap_or("").to_string();
                if self.table_styles.iter().any(|(n, _)| *n == name) {
                    continue;
                }
                let mut t2 = prep(t);
                for el in t2.children_mut() {
                    if let Some(d) = el.attr("dxfId").and_then(|v| v.parse::<u32>().ok()) {
                        if let Some(n) = map.dxf_for(d) {
                            el.set("dxfId", n.to_string());
                        } else {
                            el.remove("dxfId");
                        }
                    }
                }
                self.table_styles.push((name, ser(&t2)));
            }
        }
        if let Some(ic) = root.child("colors").and_then(|c| c.child("indexedColors")) {
            let s = ser(ic);
            match &self.indexed_colors {
                None if !self.started => self.indexed_colors = Some(s),
                Some(cur) if *cur == s => {}
                _ => warns.add(format!("«{label}»: в книге своя палитра индексированных цветов; цвета ячеек могут отличаться от оригинала")),
            }
        } else if self.indexed_colors.is_some() {
            warns.add(format!("«{label}»: у первой книги своя палитра индексированных цветов; цвета ячеек из «{label}» могут отличаться от оригинала"));
        }
        self.started = true;
        Ok(map)
    }

    fn push_csxf(&mut self, s: String) -> u32 {
        // Именованный стиль всегда получает собственный xf (без объединения одинаковых).
        self.csxfs.items.push(s);
        (self.csxfs.items.len() - 1) as u32
    }

    fn push_csxf_shared(&mut self, s: String) -> u32 {
        if let Some(i) = self.csxfs.items.iter().position(|x| *x == s) {
            return i as u32;
        }
        self.push_csxf(s)
    }

    fn finish(&self) -> Result<Vec<u8>> {
        if self.xfs.items.len() > 64000 {
            bail!("слишком много различных форматов ячеек ({}): Excel допускает не более 64000", self.xfs.items.len());
        }
        let mut s = String::from("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n");
        s.push_str(&format!("<styleSheet xmlns=\"{}\">", xml::NS_MAIN));
        if !self.numfmts.is_empty() {
            s.push_str(&format!("<numFmts count=\"{}\">", self.numfmts.len()));
            for (id, code) in &self.numfmts {
                s.push_str(&format!("<numFmt numFmtId=\"{id}\" formatCode=\""));
                xml::esc_attr(code, &mut s);
                s.push_str("\"/>");
            }
            s.push_str("</numFmts>");
        }
        let mut fonts = self.fonts.items.clone();
        if fonts.is_empty() {
            fonts.push("<font><sz val=\"11\"/><name val=\"Calibri\"/><family val=\"2\"/></font>".into());
        }
        let mut borders = self.borders.items.clone();
        if borders.is_empty() {
            borders.push("<border><left/><right/><top/><bottom/><diagonal/></border>".into());
        }
        let mut csxfs = self.csxfs.items.clone();
        let mut cell_styles = self.cell_styles.clone();
        if csxfs.is_empty() {
            csxfs.push("<xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\"/>".into());
        }
        if cell_styles.is_empty() {
            cell_styles.push("<cellStyle name=\"Normal\" xfId=\"0\" builtinId=\"0\"/>".into());
        }
        let section = |s: &mut String, tag: &str, items: &[String]| {
            s.push_str(&format!("<{tag} count=\"{}\">", items.len()));
            for i in items {
                s.push_str(i);
            }
            s.push_str(&format!("</{tag}>"));
        };
        section(&mut s, "fonts", &fonts);
        section(&mut s, "fills", &self.fills.items);
        section(&mut s, "borders", &borders);
        section(&mut s, "cellStyleXfs", &csxfs);
        section(&mut s, "cellXfs", &self.xfs.items);
        section(&mut s, "cellStyles", &cell_styles);
        section(&mut s, "dxfs", &self.dxfs.items);
        let (dts, dps) = self.default_table_style.clone().unwrap_or(("TableStyleMedium2".into(), "PivotStyleLight16".into()));
        s.push_str(&format!("<tableStyles count=\"{}\" defaultTableStyle=\"{dts}\" defaultPivotStyle=\"{dps}\">", self.table_styles.len()));
        for (_, t) in &self.table_styles {
            s.push_str(t);
        }
        s.push_str("</tableStyles>");
        if let Some(ic) = &self.indexed_colors {
            s.push_str(&format!("<colors>{ic}</colors>"));
        }
        s.push_str("</styleSheet>");
        Ok(s.into_bytes())
    }
}

/// Ширина цифры (в пикселях при 96 dpi) шрифта «Обычный» — единица измерения ширины столбцов Excel.
fn max_digit_width(name: &str, size_pt: f64) -> u32 {
    let ratio = match name.to_lowercase().as_str() {
        "calibri" | "calibri light" => 0.5068,
        "cambria" => 0.5547,
        "arial" | "liberation sans" | "microsoft sans serif" | "helvetica" => 0.556,
        "arial narrow" => 0.456,
        "times new roman" | "liberation serif" => 0.5,
        "verdana" => 0.636,
        "tahoma" => 0.546,
        "segoe ui" => 0.55,
        "courier new" | "lucida console" | "liberation mono" => 0.6,
        "consolas" => 0.55,
        "trebuchet ms" => 0.53,
        "aptos narrow" => 0.47,
        "aptos" => 0.53,
        _ => 0.55,
    };
    ((size_pt * 96.0 / 72.0 * ratio).round() as u32).max(1)
}

/// Шрифт стиля «Обычный» (имя, размер) — по нему Excel переводит ширину столбцов в пиксели.
fn normal_font(root: &xml::Elem) -> (String, f64) {
    let fonts: Vec<&xml::Elem> = root.child("fonts").map(|f| f.children().collect()).unwrap_or_default();
    let xfs: Vec<&xml::Elem> = root.child("cellStyleXfs").map(|f| f.children().collect()).unwrap_or_default();
    let normal_xf = root
        .child("cellStyles")
        .and_then(|cs| cs.children().find(|c| c.attr("builtinId") == Some("0")))
        .and_then(|c| c.attr("xfId").and_then(|v| v.parse::<usize>().ok()))
        .unwrap_or(0);
    let font_id = xfs.get(normal_xf).and_then(|x| x.attr("fontId")).and_then(|v| v.parse::<usize>().ok()).unwrap_or(0);
    match fonts.get(font_id).or(fonts.first()) {
        Some(f) => (
            f.child("name").and_then(|n| n.attr("val")).unwrap_or("Calibri").to_string(),
            f.child("sz").and_then(|n| n.attr("val")).and_then(|v| v.parse().ok()).unwrap_or(11.0),
        ),
        None => ("Calibri".into(), 11.0),
    }
}

// ---------------------------------------------------------------------------
// Состояние объединения
// ---------------------------------------------------------------------------

type Zw = zip::ZipWriter<tempfile::NamedTempFile>;

#[derive(Default)]
struct Warnings {
    list: Vec<String>,
    seen: HashSet<String>,
}

impl Warnings {
    fn add(&mut self, s: impl Into<String>) {
        let s = s.into();
        if self.seen.insert(s.clone()) {
            self.list.push(s);
        }
    }
}

struct OutSheet {
    name: String,
    state: &'static str,
    part: String,
}

struct DefName {
    name: String,
    local: Option<usize>,
    attrs: Vec<(String, String)>,
    text: String,
}

/// Всё, что накапливается в результате (кроме самого zip-файла).
#[derive(Default)]
struct St {
    pending: Vec<(String, Vec<u8>)>,
    used: HashSet<String>,
    counters: HashMap<String, usize>,
    overrides: Vec<(String, String)>,
    defaults: HashMap<String, String>,
    styles: StyleOut,
    sst: Interner,
    sst_refs: usize,
    sheets: Vec<OutSheet>,
    defined: Vec<DefName>,
    ext_links: Vec<String>,
    pivot_caches: Vec<(u32, String)>,
    theme: Option<(Theme, Vec<u8>)>,
    names_used: HashSet<String>,
    global_names: HashMap<String, String>,
    table_id: u32,
    date1904: bool,
    warns: Warnings,
    fit_to_width: bool,
    wb_pr: Option<xml::Elem>,
    wb_view: Option<xml::Elem>,
    calc_pr: Option<xml::Elem>,
    active: usize,
    /// Application / AppVersion первой книги (LibreOffice по ним выбирает способ расчёта высоты строк).
    app: Option<(String, Option<String>)>,
    /// Ширина цифры шрифта «Обычный» первой книги.
    base_mdw: u32,
    /// Нужен ли полный пересчёт при открытии (значения могли измениться).
    recalc: bool,
}

/// Исходная книга со всем контекстом переноса.
struct Src {
    book: Book,
    label: String,
    stem: String,
    selected: Vec<SheetEntry>,
    out_names: Vec<String>,
    first_out_idx: usize,
    styles: StyleMap,
    sst: Option<Vec<String>>,
    sst_map: Vec<u32>,
    fcx: FormulaCtx,
    fix: Option<ThemeFix>,
    date_shift: f64,
    /// Ширина цифры «Обычного» шрифта источника и коэффициент пересчёта ширины столбцов.
    mdw: u32,
    width_scale: f64,
    memo: HashMap<String, String>,
    /// лист результата → (формул заменено значениями, нет листов)
    converted: BTreeMap<String, (usize, BTreeSet<String>)>,
    vm_cells: usize,
    cm_cells: usize,
    neg_dates: usize,
    dropped_ext: BTreeSet<String>,
    missing_refs: BTreeMap<String, BTreeSet<String>>,
    has_drawings: bool,
    theme_differs: bool,
    /// Части таблиц источника (нижний регистр) → новые (displayName, name).
    table_names: HashMap<String, (String, String)>,
}

fn guess_content_type(ext: &str) -> &'static str {
    match ext {
        "png" => "image/png",
        "jpg" | "jpeg" | "jpe" => "image/jpeg",
        "gif" => "image/gif",
        "bmp" => "image/bmp",
        "tif" | "tiff" => "image/tiff",
        "emf" => "image/x-emf",
        "wmf" => "image/x-wmf",
        "svg" => "image/svg+xml",
        "webp" => "image/webp",
        "vml" => "application/vnd.openxmlformats-officedocument.vmlDrawing",
        "bin" => "application/vnd.openxmlformats-officedocument.oleObject",
        "xlsx" => "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        "docx" => "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        "pptx" => "application/vnd.openxmlformats-officedocument.presentationml.presentation",
        "xml" => "application/xml",
        _ => "application/octet-stream",
    }
}

/// Расширения, для которых тип содержимого пишется как Default (общий для всех частей).
const DEFAULT_EXTS: &[&str] = &["png", "jpg", "jpeg", "jpe", "gif", "bmp", "tif", "tiff", "emf", "wmf", "svg", "webp", "vml", "bin"];

impl St {
    fn reserve(&mut self, part: &str) {
        self.used.insert(part.to_lowercase());
    }

    /// Свободное имя части по образцу исходной (`xl/charts/chart3.xml` → `xl/charts/chart1.xml`).
    fn alloc_part_name(&mut self, src_part: &str) -> String {
        let (dir, file) = src_part.rsplit_once('/').unwrap_or(("", src_part));
        let (stem, ext) = match file.rsplit_once('.') {
            Some((s, e)) => (s, format!(".{}", e.to_lowercase())),
            None => (file, String::new()),
        };
        let base = stem.trim_end_matches(|c: char| c.is_ascii_digit());
        let key = format!("{}/{}{}", dir.to_lowercase(), base.to_lowercase(), ext);
        loop {
            let n = self.counters.entry(key.clone()).or_insert(0);
            *n += 1;
            let name = if dir.is_empty() { format!("{base}{n}{ext}") } else { format!("{dir}/{base}{n}{ext}") };
            if self.used.insert(name.to_lowercase()) {
                return name;
            }
        }
    }

    fn register_ct(&mut self, part: &str, ct: &str) {
        let ext = part.rsplit_once('.').map(|x| x.1.to_lowercase()).unwrap_or_default();
        if DEFAULT_EXTS.contains(&ext.as_str()) {
            match self.defaults.get(&ext) {
                None => {
                    self.defaults.insert(ext, ct.to_string());
                    return;
                }
                Some(c) if c == ct => return,
                _ => {}
            }
        }
        self.overrides.push((part.to_string(), ct.to_string()));
    }

    fn put(&mut self, part: &str, ct: Option<&str>, bytes: Vec<u8>) {
        if let Some(ct) = ct {
            self.register_ct(part, ct);
        }
        self.pending.push((part.to_string(), bytes));
    }

    /// Следующий номер для таблиц, уникальный в книге.
    fn next_table_id(&mut self) -> u32 {
        self.table_id += 1;
        self.table_id
    }
}

fn zip_opts() -> zip::write::SimpleFileOptions {
    zip::write::SimpleFileOptions::default().compression_method(zip::CompressionMethod::Deflated).large_file(false)
}

fn flush(zip: &mut Zw, st: &mut St) -> Result<()> {
    for (name, bytes) in st.pending.drain(..) {
        zip.start_file(name.as_str(), zip_opts()).map_err(|e| anyhow!("не удалось записать результат: {e}"))?;
        zip.write_all(&bytes).map_err(|e| anyhow!("не удалось записать результат: {}", io_reason_str(&e)))?;
    }
    Ok(())
}

fn io_reason_str(e: &std::io::Error) -> String {
    match e.kind() {
        std::io::ErrorKind::StorageFull => "на диске не хватает места".into(),
        _ => io_reason(e),
    }
}

fn col_number(letters: &str) -> usize {
    letters.bytes().filter(u8::is_ascii_alphabetic).fold(0usize, |a, b| a * 26 + (b.to_ascii_uppercase() - b'A' + 1) as usize)
}

/// Связи, которые не переносятся (с пояснениями в предупреждениях делает вызывающий).
fn skipped_rel(kind: &str) -> bool {
    matches!(
        kind,
        "printerSettings"
            | "threadedComment"
            | "person"
            | "queryTable"
            | "vbaProject"
            | "vbaProjectSignature"
            | "slicer"
            | "slicerCache"
            | "timeline"
            | "timelineCache"
            | "customXml"
            | "customProperty"
            | "calcChain"
            | "connections"
    )
}

/// Копирует связи части: целевые части копируются в результат под новыми именами,
/// идентификаторы связей остаются прежними (ссылки `r:id` в XML не меняются).
fn copy_rels(
    zip_st: &mut St,
    s: &mut Src,
    src_part: &str,
    new_part: &str,
) -> Result<(Vec<Rel>, Vec<(String, String, String)>)> {
    let rels = s.book.pkg.rels(src_part).map_err(|e| anyhow!("«{}»: {e}", s.label))?;
    let mut out = vec![];
    let mut mapped = vec![];
    for r in rels {
        if r.external {
            out.push(r);
            continue;
        }
        let kind = r.kind().to_string();
        if skipped_rel(&kind) {
            match kind.as_str() {
                "threadedComment" => {
                    s.dropped_ext.insert("обсуждения (threaded comments) — остались обычные примечания".into());
                }
                "slicer" | "timeline" => {
                    s.dropped_ext.insert("срезы и временные шкалы".into());
                }
                _ => {}
            }
            continue;
        }
        if kind == "hyperlink" && r.target.starts_with('#') {
            // Внутренняя ссылка на место в книге: Target="#Лист!A1".
            let loc = r.target[1..].to_string();
            let rw = s.fcx.rewrite(&loc, Missing::Keep);
            let mut r2 = r.clone();
            r2.target = format!("#{}", rw.text);
            out.push(r2);
            continue;
        }
        let target = resolve_target(src_part, &r.target);
        match copy_part(zip_st, s, &target, &kind)? {
            Some(np) => {
                mapped.push((r.id.clone(), kind, np.clone()));
                out.push(Rel { id: r.id, typ: r.typ, target: relative_target(new_part, &np), external: false });
            }
            None => {} // целевой части нет в файле — связь повреждена, пропускаем
        }
    }
    Ok((out, mapped))
}

// ---------------------------------------------------------------------------
// Копирование частей пакета (рисунки, диаграммы, таблицы, примечания, сводные…)
// ---------------------------------------------------------------------------

fn copy_part(st: &mut St, s: &mut Src, part: &str, kind: &str) -> Result<Option<String>> {
    let key = part.to_lowercase();
    if let Some(n) = s.memo.get(&key) {
        return Ok(Some(n.clone()));
    }
    if !s.book.pkg.has(part) {
        return Ok(None);
    }
    if kind == "pivotTable" {
        let has_cache = s.book.pkg.rels(part).unwrap_or_default().iter().any(|r| {
            r.kind() == "pivotCacheDefinition" && !r.external && s.book.pkg.has(&resolve_target(part, &r.target))
        });
        if !has_cache {
            s.dropped_ext.insert("сводные таблицы без кэша данных (в файле повреждены)".into());
            return Ok(None);
        }
    }
    let new_part = st.alloc_part_name(part);
    s.memo.insert(key, new_part.clone());
    let (rels, mapped) = copy_rels(st, s, part, &new_part)?;
    let bytes = s.book.pkg.read(part).map_err(|e| anyhow!("«{}»: {e}", s.label))?;
    let transformed = match kind {
        "drawing" => transform_drawing(s, &bytes),
        "chart" => transform_chart(s, &bytes),
        "table" => transform_table(st, s, part, &bytes),
        "pivotTable" => transform_pivot_table(st, s, &bytes, &mapped),
        "pivotCacheDefinition" => transform_pivot_cache(s, &bytes),
        _ => Ok(bytes.clone()),
    };
    // Если часть не удалось разобрать, копируем её как есть, чтобы ничего не потерять.
    let bytes = transformed.unwrap_or(bytes);
    let ext = part.rsplit_once('.').map(|x| x.1.to_lowercase()).unwrap_or_default();
    let ct = s.book.pkg.content_type(part).unwrap_or_else(|| guess_content_type(&ext).to_string());
    st.put(&new_part, Some(&ct), bytes);
    if !rels.is_empty() {
        st.put(&rels_path(&new_part), None, rels_xml(&rels));
    }
    Ok(Some(new_part))
}

fn note_missing(s: &mut Src, place: &str, missing: Vec<String>) {
    if !missing.is_empty() {
        s.missing_refs.entry(place.to_string()).or_default().extend(missing);
    }
}

fn transform_drawing(s: &mut Src, bytes: &[u8]) -> Result<Vec<u8>> {
    let mut root = xml::parse(bytes)?;
    s.has_drawings = true;
    let before = root.kids.len();
    root.kids.retain(|k| match k {
        xml::Node::E(a) => {
            let mut bad = false;
            a.walk(&mut |x| {
                if x.ns.contains("slicer") {
                    bad = true;
                }
            });
            !bad
        }
        _ => true,
    });
    if root.kids.len() != before {
        s.dropped_ext.insert("срезы и временные шкалы".into());
    }
    let mut edits: Vec<(String, String)> = vec![];
    root.walk(&mut |x| {
        for a in &x.attrs {
            if a.q == "textlink" && !a.v.is_empty() {
                edits.push((a.v.clone(), String::new()));
            }
        }
    });
    if !edits.is_empty() {
        let mut missing = vec![];
        let mut rewritten: HashMap<String, String> = HashMap::new();
        for (orig, _) in &edits {
            let r = s.fcx.rewrite(orig, Missing::Keep);
            missing.extend(r.missing);
            rewritten.insert(orig.clone(), r.text);
        }
        root.walk_mut(&mut |x| {
            for a in &mut x.attrs {
                if a.q == "textlink" {
                    if let Some(n) = rewritten.get(&a.v) {
                        a.v = n.clone();
                    }
                }
            }
        });
        note_missing(s, "рисунки", missing);
    }
    Ok(xml::doc(&root, ""))
}

fn transform_chart(s: &mut Src, bytes: &[u8]) -> Result<Vec<u8>> {
    let mut root = xml::parse(bytes)?;
    s.has_drawings = true;
    let mut missing = vec![];
    let fcx = &s.fcx;
    root.walk_mut(&mut |x| {
        if x.local == "f" && !x.ns.is_empty() {
            let t = x.text();
            if !t.trim().is_empty() {
                let r = fcx.rewrite(&t, Missing::Keep);
                missing.extend(r.missing);
                if r.text != t {
                    x.set_text(r.text);
                }
            }
        }
    });
    note_missing(s, "диаграммы", missing);
    Ok(xml::doc(&root, ""))
}

/// Допустимое имя таблицы/имени Excel из произвольной строки.
fn sanitize_ident(s: &str) -> String {
    let mut r: String = s.chars().map(|c| if c.is_alphanumeric() || c == '_' || c == '.' { c } else { '_' }).collect();
    if r.is_empty() || r.chars().next().is_some_and(|c| c.is_ascii_digit() || c == '.') {
        r.insert(0, '_');
    }
    r
}

/// Свободное имя вида `имя_суффикс`, `имя_2`… среди занятых (в нижнем регистре).
fn unique_ident(used: &HashSet<String>, base: &str, stem: &str) -> String {
    let first = format!("{base}_{}", sanitize_ident(stem));
    if !used.contains(&first.to_lowercase()) {
        return first;
    }
    for i in 2.. {
        let c = format!("{base}_{i}");
        if !used.contains(&c.to_lowercase()) {
            return c;
        }
    }
    unreachable!()
}

fn transform_table(st: &mut St, s: &mut Src, part: &str, bytes: &[u8]) -> Result<Vec<u8>> {
    let mut root = xml::parse(bytes)?;
    if root.local != "table" {
        bail!("не таблица");
    }
    root.set("id", st.next_table_id().to_string());
    if let Some((display, name)) = s.table_names.get(&part.to_lowercase()) {
        root.set("displayName", display.as_str());
        root.set("name", name.as_str());
    }
    let is_query = root.attr("tableType") == Some("queryTable");
    if is_query {
        root.remove("tableType");
        s.dropped_ext.insert("запросы Power Query/внешние данные таблиц (таблицы превращены в обычные)".into());
    }
    let styles = &s.styles;
    let fcx = &s.fcx;
    let mut missing = vec![];
    root.walk_mut(&mut |x| {
        for a in &mut x.attrs {
            if a.ns.is_empty() && a.q.ends_with("DxfId") {
                if let Ok(v) = a.v.parse::<u32>() {
                    if let Some(n) = styles.dxf_for(v) {
                        a.v = n.to_string();
                    }
                }
            }
        }
        if is_query {
            x.remove("queryTableFieldId");
        }
        if x.local == "calculatedColumnFormula" || x.local == "totalsRowFormula" {
            let t = x.text();
            if !t.trim().is_empty() {
                let r = fcx.rewrite(&t, Missing::Keep);
                missing.extend(r.missing);
                if r.text != t {
                    x.set_text(r.text);
                }
            }
        }
    });
    note_missing(s, "таблицы", missing);
    Ok(xml::doc(&root, xml::NS_MAIN))
}

fn transform_pivot_table(st: &mut St, s: &mut Src, bytes: &[u8], mapped: &[(String, String, String)]) -> Result<Vec<u8>> {
    let mut root = xml::parse(bytes)?;
    let cache_part = mapped
        .iter()
        .find(|m| m.1 == "pivotCacheDefinition")
        .map(|m| m.2.clone())
        .ok_or_else(|| anyhow!("у сводной таблицы нет кэша"))?;
    let id = match st.pivot_caches.iter().find(|(_, p)| *p == cache_part) {
        Some((id, _)) => *id,
        None => {
            let id = st.pivot_caches.len() as u32 + 1;
            st.pivot_caches.push((id, cache_part));
            id
        }
    };
    root.set("cacheId", id.to_string());
    let styles = &s.styles;
    root.walk_mut(&mut |x| {
        if let Some(v) = x.attr("dxfId").and_then(|v| v.parse::<u32>().ok()) {
            if let Some(n) = styles.dxf_for(v) {
                x.set("dxfId", n.to_string());
            }
        }
        if let Some(v) = x.attr("numFmtId").and_then(|v| v.parse::<u32>().ok()) {
            x.set("numFmtId", styles.numfmt_for(v).to_string());
        }
    });
    Ok(xml::doc(&root, ""))
}

fn transform_pivot_cache(s: &mut Src, bytes: &[u8]) -> Result<Vec<u8>> {
    let mut root = xml::parse(bytes)?;
    let styles = &s.styles;
    let fcx = &s.fcx;
    let mut missing = vec![];
    root.walk_mut(&mut |x| {
        if x.local == "worksheetSource" {
            if let Some(sh) = x.attr("sheet").map(str::to_string) {
                match fcx.sheets.get(&sh.to_lowercase()) {
                    Some((new, _)) => x.set("sheet", new.as_str()),
                    None => missing.push(sh),
                }
            }
            if let Some(n) = x.attr("name").map(str::to_lowercase) {
                if let Some(Some(new)) = fcx.tables.get(&n) {
                    x.set("name", new.as_str());
                } else if let Some(new) = fcx.names.get(&n) {
                    x.set("name", new.as_str());
                }
            }
        }
        if x.local == "cacheField" {
            if let Some(v) = x.attr("numFmtId").and_then(|v| v.parse::<u32>().ok()) {
                x.set("numFmtId", styles.numfmt_for(v).to_string());
            }
        }
    });
    note_missing(s, "сводные таблицы (источник данных)", missing);
    Ok(xml::doc(&root, ""))
}

// ---------------------------------------------------------------------------
// Перенос листа
// ---------------------------------------------------------------------------

/// Порядок дочерних элементов `worksheet` по схеме.
fn top_rank(e: &xml::Elem) -> u32 {
    if e.ns != xml::NS_MAIN {
        return match e.local.as_str() {
            "AlternateContent" => 35,
            _ => 99,
        };
    }
    match e.local.as_str() {
        "sheetPr" => 0,
        "dimension" => 1,
        "sheetViews" => 2,
        "sheetFormatPr" => 3,
        "cols" => 4,
        "sheetData" => 5,
        "sheetCalcPr" => 6,
        "sheetProtection" => 7,
        "protectedRanges" => 8,
        "scenarios" => 9,
        "autoFilter" => 10,
        "sortState" => 11,
        "dataConsolidate" => 12,
        "customSheetViews" => 13,
        "mergeCells" => 14,
        "phoneticPr" => 15,
        "conditionalFormatting" => 16,
        "dataValidations" => 17,
        "hyperlinks" => 18,
        "printOptions" => 19,
        "pageMargins" => 20,
        "pageSetup" => 21,
        "headerFooter" => 22,
        "rowBreaks" => 23,
        "colBreaks" => 24,
        "customProperties" => 25,
        "cellWatches" => 26,
        "ignoredErrors" => 27,
        "smartTags" => 28,
        "drawing" => 29,
        "legacyDrawing" => 30,
        "legacyDrawingHF" => 31,
        "drawingHF" => 32,
        "picture" => 33,
        "oleObjects" => 34,
        "controls" => 35,
        "webPublishItems" => 36,
        "tableParts" => 37,
        "extLst" => 38,
        _ => 99,
    }
}

#[derive(PartialEq, Clone, Copy)]
enum Pend {
    SheetPr,
    SheetViews,
    SheetFormat,
    PageSetup,
}

impl Pend {
    fn rank(self) -> u32 {
        match self {
            Pend::SheetPr => 0,
            Pend::SheetViews => 2,
            Pend::SheetFormat => 3,
            Pend::PageSetup => 21,
        }
    }
    fn local(self) -> &'static str {
        match self {
            Pend::SheetPr => "sheetPr",
            Pend::SheetViews => "sheetViews",
            Pend::SheetFormat => "sheetFormatPr",
            Pend::PageSetup => "pageSetup",
        }
    }
}

#[derive(Default)]
struct SheetState {
    name: String,
    dead_si: HashSet<String>,
    max_col: usize,
    cols: Vec<(usize, usize, f64, bool)>,
    default_width: Option<f64>,
    /// Ширина по умолчанию для вставляемого sheetFormatPr.
    insert_width: f64,
    converted: usize,
    no_value: usize,
    missing: BTreeSet<String>,
}

const NS_XM: &str = "http://schemas.microsoft.com/office/excel/2006/main";
const URI_SPARKLINES: &str = "{05C60535-1F16-4fd2-B633-F4F36F0B64E0}";
const URI_X14_CF: &str = "{78C0D931-6437-4328-A5C5-2E7C1D0A9C6D}";
const URI_X14_DV: &str = "{CCE6A557-97BC-4b89-ADB6-D9C93CAAB3DF}";

fn inline_str(text: &str) -> xml::Elem {
    let mut t = xml::Elem::main("t");
    t.set_ns(xml::NS_XML, "xml", "space", "preserve");
    t.set_text(text);
    let mut is = xml::Elem::main("is");
    is.push(t);
    is
}

/// Превращает формулу в значение (по кэшированному результату). Возвращает false, если кэша нет.
fn formula_to_value(c: &mut xml::Elem) -> bool {
    c.kids.retain(|k| !matches!(k, xml::Node::E(x) if x.local == "f"));
    let t = c.attr("t").map(str::to_string);
    let v = c.child("v").map(|v| v.text());
    match (t.as_deref(), v) {
        (Some("str"), Some(v)) => {
            c.kids.retain(|k| !matches!(k, xml::Node::E(x) if x.local == "v"));
            c.set("t", "inlineStr");
            c.push(inline_str(&v));
            true
        }
        (_, Some(_)) => true,
        (_, None) => {
            c.remove("t");
            false
        }
    }
}

fn load_sst(s: &mut Src) -> Result<()> {
    if s.sst.is_some() {
        return Ok(());
    }
    let mut items = vec![];
    if let Some(r) = s.book.rel_of_kind("sharedStrings") {
        let part = resolve_target(&s.book.wb_part, &r.target);
        if s.book.pkg.has(&part) {
            let bytes = s.book.pkg.read(&part).map_err(|e| anyhow!("«{}»: {e}", s.label))?;
            let text = xml::decode(&bytes).map_err(|e| anyhow!("«{}»: sharedStrings: {e}", s.label))?;
            let mut p = xml::Parser::new(&text);
            let mut depth = 0usize;
            loop {
                match p.next().map_err(|e| anyhow!("«{}»: повреждён sharedStrings.xml ({e})", s.label))? {
                    xml::Ev::Start(e) => {
                        depth += 1;
                        if depth == 2 && e.local == "si" {
                            let mut t = p.tree(e).map_err(|e| anyhow!("«{}»: повреждён sharedStrings.xml ({e})", s.label))?;
                            if let Some(f) = &s.fix {
                                f.apply(&mut t);
                            }
                            items.push(ser(&t));
                            depth -= 1;
                        }
                    }
                    xml::Ev::Empty(e) => {
                        if depth == 1 && e.local == "si" {
                            items.push(ser(&e));
                        }
                    }
                    xml::Ev::End => depth = depth.saturating_sub(1),
                    xml::Ev::Eof => break,
                    xml::Ev::Text(_) => {}
                }
            }
        }
    }
    s.sst_map = vec![u32::MAX; items.len()];
    s.sst = Some(items);
    Ok(())
}

fn map_sst(st: &mut St, s: &mut Src, idx: usize) -> Result<u32> {
    load_sst(s)?;
    if s.sst_map[..].get(idx).is_none() {
        bail!("«{}»: в листе есть ссылка на несуществующую общую строку №{idx}; файл повреждён", s.label);
    }
    if s.sst_map[idx] == u32::MAX {
        let raw = s.sst.as_ref().unwrap()[idx].clone();
        s.sst_map[idx] = st.sst.add(raw) as u32;
    }
    st.sst_refs += 1;
    Ok(s.sst_map[idx])
}

fn fmt_num(v: f64) -> String {
    if v.fract() == 0.0 && v.abs() < 1e15 { format!("{}", v as i64) } else { format!("{v}") }
}

fn process_cell(st: &mut St, s: &mut Src, sh: &mut SheetState, c: &mut xml::Elem) -> Result<()> {
    let sattr = c.attr("s").map(str::to_string);
    let newxf = s.styles.xf_for(sattr.as_deref());
    if newxf != 0 || sattr.is_some() {
        c.set("s", newxf.to_string());
    }
    if c.remove("cm").is_some() {
        s.cm_cells += 1;
    }
    if c.remove("vm").is_some() {
        s.vm_cells += 1;
    }
    if let Some(r) = c.attr("r") {
        let letters: String = r.chars().take_while(|c| c.is_ascii_alphabetic()).collect();
        sh.max_col = sh.max_col.max(col_number(&letters));
    }
    // Формула.
    let mut convert = false;
    if let Some(f) = c.child_mut("f") {
        let ftype = f.attr("t").unwrap_or("").to_string();
        let si = f.attr("si").map(str::to_string);
        let text = f.text();
        if ftype == "shared" && text.is_empty() {
            if si.as_ref().is_some_and(|x| sh.dead_si.contains(x)) {
                convert = true;
            }
        } else if !text.trim().is_empty() {
            let r = s.fcx.rewrite(&text, Missing::Keep);
            if !r.missing.is_empty() {
                convert = true;
                sh.missing.extend(r.missing);
                if ftype == "shared" {
                    if let Some(x) = si {
                        sh.dead_si.insert(x);
                    }
                }
            } else if r.text != text {
                f.set_text(r.text);
            }
        }
    }
    if convert {
        sh.converted += 1;
        if !formula_to_value(c) {
            sh.no_value += 1;
        }
    }
    // Значение.
    let t = c.attr("t").map(str::to_string);
    match t.as_deref() {
        Some("s") => {
            if let Some(v) = c.child_mut("v") {
                let idx: usize = v.text().trim().parse().map_err(|_| anyhow!("«{}»: некорректный индекс строки в ячейке", s.label))?;
                let n = map_sst(st, s, idx)?;
                v.set_text(n.to_string());
            }
        }
        None | Some("n") => {
            if s.date_shift != 0.0 && s.styles.is_date(sattr.as_deref()) {
                if let Some(v) = c.child_mut("v") {
                    if let Ok(x) = v.text().trim().parse::<f64>() {
                        let y = x + s.date_shift;
                        if y < 0.0 {
                            s.neg_dates += 1;
                        } else {
                            v.set_text(fmt_num(y));
                        }
                    }
                }
            }
        }
        _ => {}
    }
    Ok(())
}

fn process_row(st: &mut St, s: &mut Src, sh: &mut SheetState, row: &mut xml::Elem) -> Result<()> {
    if let Some(v) = row.attr("s").map(str::to_string) {
        row.set("s", s.styles.xf_for(Some(&v)).to_string());
    }
    for c in row.children_mut() {
        if c.local == "c" {
            process_cell(st, s, sh, c)?;
        }
    }
    Ok(())
}

fn prune_dangling(e: &mut xml::Elem, valid: &HashSet<String>) -> bool {
    let own = e.attrs.iter().any(|a| a.ns == xml::NS_R && !valid.contains(&a.v));
    e.kids.retain_mut(|k| match k {
        xml::Node::E(c) => !prune_dangling(c, valid),
        _ => true,
    });
    own
}

/// Обработка элемента верхнего уровня листа. None — элемент не переносится.
fn top_level(s: &mut Src, sh: &mut SheetState, mut e: xml::Elem, valid: &HashSet<String>, active: bool, fit: bool) -> Option<xml::Elem> {
    let place = sh.name.clone();
    if e.ns == xml::NS_MAIN {
        match e.local.as_str() {
            "sheetPr" => {
                e.remove("codeName");
                if let Some(f) = &s.fix {
                    f.apply(&mut e);
                }
                if fit {
                    match e.child_mut("pageSetUpPr") {
                        Some(p) => p.set("fitToPage", "1"),
                        None => {
                            let mut p = xml::Elem::main("pageSetUpPr");
                            p.set("fitToPage", "1");
                            e.push(p);
                        }
                    }
                }
            }
            "sheetViews" => {
                for (i, v) in e.children_mut().enumerate() {
                    v.remove("tabSelected");
                    if v.attr("workbookViewId").is_some() {
                        v.set("workbookViewId", "0");
                    }
                    if active && i == 0 {
                        v.set("tabSelected", "1");
                    }
                }
            }
            "sheetFormatPr" => {
                sh.default_width = e.attr("defaultColWidth").and_then(|v| v.parse().ok());
                if s.width_scale != 1.0 {
                    // Без явной ширины по умолчанию Excel берёт baseColWidth цифр + 5 пикселей.
                    let w = sh.default_width.unwrap_or_else(|| {
                        let b: f64 = e.attr("baseColWidth").and_then(|v| v.parse().ok()).unwrap_or(8.0);
                        let mdw = s.mdw as f64;
                        ((b * mdw + 5.0) / mdw * 256.0).trunc() / 256.0
                    });
                    let scaled = w * s.width_scale;
                    sh.default_width = Some(scaled);
                    e.set("defaultColWidth", format!("{scaled:.6}"));
                }
            }
            "cols" => {
                let xf0 = s.styles.xf_for(None);
                for c in e.children_mut() {
                    let st_attr = c.attr("style").map(str::to_string);
                    let n = s.styles.xf_for(st_attr.as_deref());
                    if st_attr.is_some() || xf0 != 0 {
                        c.set("style", n.to_string());
                    }
                    let min = c.attr("min").and_then(|v| v.parse().ok()).unwrap_or(1);
                    let max = c.attr("max").and_then(|v| v.parse().ok()).unwrap_or(min);
                    let mut w: f64 = c.attr("width").and_then(|v| v.parse().ok()).unwrap_or(0.0);
                    if s.width_scale != 1.0 && c.attr("width").is_some() {
                        w *= s.width_scale;
                        c.set("width", format!("{w:.6}"));
                    }
                    let hidden = matches!(c.attr("hidden"), Some("1" | "true"));
                    sh.cols.push((min, max, w, hidden));
                }
            }
            "conditionalFormatting" => {
                if let Some(f) = &s.fix {
                    f.apply(&mut e);
                }
                let mut missing = vec![];
                let styles = &s.styles;
                let fcx = &s.fcx;
                e.walk_mut(&mut |x| {
                    if x.local == "cfRule" {
                        if let Some(v) = x.attr("dxfId").and_then(|v| v.parse::<u32>().ok()) {
                            match styles.dxf_for(v) {
                                Some(n) => x.set("dxfId", n.to_string()),
                                None => {
                                    x.remove("dxfId");
                                }
                            }
                        }
                    }
                    let is_formula = x.local == "formula" || (x.local == "cfvo" && x.attr("type") == Some("formula"));
                    if is_formula {
                        let (text, in_attr) = if x.local == "cfvo" { (x.attr("val").unwrap_or("").to_string(), true) } else { (x.text(), false) };
                        if !text.trim().is_empty() {
                            let r = fcx.rewrite(&text, Missing::Ref);
                            missing.extend(r.missing);
                            if r.text != text {
                                if in_attr {
                                    x.set("val", r.text);
                                } else {
                                    x.set_text(r.text);
                                }
                            }
                        }
                    }
                });
                note_missing(s, &format!("лист «{place}» (условное форматирование)"), missing);
            }
            "dataValidations" => {
                let mut missing = vec![];
                let fcx = &s.fcx;
                e.walk_mut(&mut |x| {
                    if x.local == "formula1" || x.local == "formula2" {
                        let t = x.text();
                        if !t.trim().is_empty() {
                            let r = fcx.rewrite(&t, Missing::Ref);
                            missing.extend(r.missing);
                            if r.text != t {
                                x.set_text(r.text);
                            }
                        }
                    }
                });
                note_missing(s, &format!("лист «{place}» (проверка данных)"), missing);
            }
            "hyperlinks" => {
                let mut missing = vec![];
                for h in e.children_mut() {
                    if let Some(loc) = h.attr("location").map(str::to_string) {
                        let r = s.fcx.rewrite(&loc, Missing::Keep);
                        missing.extend(r.missing);
                        if r.text != loc {
                            h.set("location", r.text);
                        }
                    }
                }
                note_missing(s, &format!("лист «{place}» (гиперссылки)"), missing);
            }
            "pageSetup" => {
                e.remove_ns(xml::NS_R, "id");
            }
            "customSheetViews" => return None,
            _ if top_rank(&e) == 99 => return None, // неизвестный элемент в основном пространстве имён: файл не открылся бы
            "colBreaks" if fit => return None,
            "extLst" => {
                e.kids.retain_mut(|k| {
                    let xml::Node::E(ext) = k else { return false };
                    let uri = ext.attr("uri").unwrap_or("").to_string();
                    let known = uri.eq_ignore_ascii_case(URI_SPARKLINES) || uri.eq_ignore_ascii_case(URI_X14_CF) || uri.eq_ignore_ascii_case(URI_X14_DV);
                    if !known {
                        s.dropped_ext.insert("расширения листа (срезы, временные шкалы и пр.)".into());
                        return false;
                    }
                    true
                });
                let mut missing = vec![];
                let fcx = &s.fcx;
                let fix = &s.fix;
                let styles = &s.styles;
                e.walk_mut(&mut |x| {
                    if x.ns == NS_XM && x.local == "f" {
                        let t = x.text();
                        if !t.trim().is_empty() {
                            let r = fcx.rewrite(&t, Missing::Keep);
                            missing.extend(r.missing);
                            if r.text != t {
                                x.set_text(r.text);
                            }
                        }
                    }
                    if x.local == "dxf" {
                        if let Some(f) = fix {
                            f.apply(x);
                        }
                        let _ = styles;
                    }
                });
                note_missing(s, &format!("лист «{place}» (спарклайны и расширенные правила)"), missing);
                if e.children().next().is_none() {
                    return None;
                }
            }
            _ => {}
        }
    }
    if e.ns == xml::NS_MAIN && e.local == "tableParts" {
        // Ссылки на таблицы, которые не удалось перенести, убираются ниже общей чисткой.
    }
    if prune_dangling(&mut e, valid) {
        return None;
    }
    if e.ns == xml::NS_MAIN {
        if matches!(e.local.as_str(), "controls" | "oleObjects" | "hyperlinks" | "tableParts" | "customProperties") {
            if e.children().next().is_none() {
                return None;
            }
            if e.local == "tableParts" {
                let n = e.children().count();
                e.set("count", n.to_string());
            }
        }
    }
    Some(e)
}

fn make_page_setup(sh: &SheetState) -> xml::Elem {
    let mut ps = xml::Elem::main("pageSetup");
    apply_fit(&mut ps, sh);
    ps
}

/// Ширина (в символах) используемой области листа.
fn used_width(sh: &SheetState) -> f64 {
    let def = sh.default_width.unwrap_or(9.14);
    let mut total = 0.0;
    for col in 1..=sh.max_col {
        let mut w = def;
        for &(min, max, cw, hidden) in &sh.cols {
            if col >= min && col <= max {
                w = if hidden { 0.0 } else { cw };
            }
        }
        total += w;
    }
    total
}

fn apply_fit(ps: &mut xml::Elem, sh: &SheetState) {
    ps.remove("scale");
    ps.set("fitToWidth", "1");
    ps.set("fitToHeight", "0");
    if used_width(sh) > 100.0 {
        ps.set("orientation", "landscape");
    }
}

/// Переносит один лист исходной книги в результат.
#[allow(clippy::too_many_arguments)]
fn transform_sheet(
    zip: &mut Zw,
    st: &mut St,
    s: &mut Src,
    entry: &SheetEntry,
    out_idx: usize,
    new_name: &str,
    state: &'static str,
    active: bool,
) -> Result<()> {
    let out_part = format!("xl/worksheets/sheet{}.xml", out_idx + 1);
    st.reserve(&out_part);
    let (rels, _mapped) = copy_rels(st, s, &entry.part, &out_part)?;
    let valid: HashSet<String> = rels.iter().map(|r| r.id.clone()).collect();
    flush(zip, st)?;

    let bytes = s.book.pkg.read(&entry.part).map_err(|e| anyhow!("«{}», лист «{}»: {e}", s.label, entry.name))?;
    let text = xml::decode(&bytes).map_err(|e| anyhow!("«{}», лист «{}»: {e}", s.label, entry.name))?;
    let (lbl, sname) = (s.label.clone(), entry.name.clone());
    let bad = |e: anyhow::Error| anyhow!("«{lbl}», лист «{sname}»: повреждён XML листа ({e})");
    let mut p = xml::Parser::new(&text);
    let root = loop {
        match p.next().map_err(bad)? {
            xml::Ev::Start(e) => break e,
            xml::Ev::Empty(e) => break e,
            xml::Ev::Eof => bail!("«{}», лист «{}»: пустая часть листа", s.label, entry.name),
            _ => {}
        }
    };
    if root.local != "worksheet" {
        bail!("«{}», лист «{}»: это не рабочий лист", s.label, entry.name);
    }
    let fit = st.fit_to_width;
    let mut sh = SheetState { name: new_name.to_string(), ..Default::default() };
    let mut pending: Vec<Pend> = vec![];
    if fit {
        pending.push(Pend::SheetPr);
        pending.push(Pend::PageSetup);
    }
    if active {
        pending.push(Pend::SheetViews);
    }
    if s.width_scale != 1.0 {
        let mdw = s.mdw as f64;
        sh.insert_width = ((8.0 * mdw + 5.0) / mdw * 256.0).trunc() / 256.0 * s.width_scale;
        pending.push(Pend::SheetFormat);
    }

    zip.start_file(out_part.as_str(), zip_opts()).map_err(|e| anyhow!("не удалось записать результат: {e}"))?;
    let werr = |e: std::io::Error| anyhow!("не удалось записать результат: {}", io_reason_str(&e));
    let mut sw = xml::Sw::new(zip, xml::NS_MAIN);
    let mut root_out = root.clone();
    root_out.kids.clear();
    sw.start(&root_out);

    loop {
        let ev = p.next().map_err(bad)?;
        match ev {
            xml::Ev::Start(e) if e.local == "sheetData" && e.ns == xml::NS_MAIN => {
                flush_pending(&mut sw, &sh, &mut pending, 5, &werr)?;
                sw.start(&e);
                loop {
                    match p.next().map_err(bad)? {
                        xml::Ev::Start(r) => {
                            let mut row = p.tree(r).map_err(bad)?;
                            process_row(st, s, &mut sh, &mut row)?;
                            sw.elem(&row).map_err(werr)?;
                        }
                        xml::Ev::Empty(mut row) => {
                            process_row(st, s, &mut sh, &mut row)?;
                            sw.elem(&row).map_err(werr)?;
                        }
                        xml::Ev::End => break,
                        xml::Ev::Text(_) => {}
                        xml::Ev::Eof => bail!("«{}», лист «{}»: обрыв XML", s.label, entry.name),
                    }
                }
                sw.end().map_err(werr)?;
            }
            xml::Ev::Start(e) => {
                let t = p.tree(e).map_err(bad)?;
                handle_top(&mut sw, s, &mut sh, &mut pending, t, &valid, active, fit, &werr)?;
            }
            xml::Ev::Empty(e) => {
                if e.local == "sheetData" && e.ns == xml::NS_MAIN {
                    flush_pending(&mut sw, &sh, &mut pending, 5, &werr)?;
                }
                handle_top(&mut sw, s, &mut sh, &mut pending, e, &valid, active, fit, &werr)?;
            }
            xml::Ev::End | xml::Ev::Eof => break,
            xml::Ev::Text(_) => {}
        }
    }
    flush_pending(&mut sw, &sh, &mut pending, u32::MAX, &werr)?;
    sw.finish().map_err(werr)?;

    st.register_ct(&out_part, CT_WORKSHEET);
    if !rels.is_empty() {
        st.put(&rels_path(&out_part), None, rels_xml(&rels));
        flush(zip, st)?;
    }
    st.sheets.push(OutSheet { name: new_name.to_string(), state, part: out_part });
    if sh.converted > 0 {
        let e = s.converted.entry(new_name.to_string()).or_default();
        e.0 += sh.converted;
        e.1.extend(sh.missing.iter().cloned());
        if sh.no_value > 0 {
            s.missing_refs.entry(format!("лист «{new_name}» (формулы без сохранённого значения)")).or_default().insert(sh.no_value.to_string());
        }
    }
    Ok(())
}

fn make_pending(kind: Pend, sh: &SheetState) -> xml::Elem {
    match kind {
        Pend::SheetPr => {
            let mut x = xml::Elem::main("sheetPr");
            let mut p = xml::Elem::main("pageSetUpPr");
            p.set("fitToPage", "1");
            x.push(p);
            x
        }
        Pend::SheetViews => {
            let mut v = xml::Elem::main("sheetView");
            v.set("workbookViewId", "0");
            v.set("tabSelected", "1");
            let mut x = xml::Elem::main("sheetViews");
            x.push(v);
            x
        }
        Pend::SheetFormat => {
            let mut x = xml::Elem::main("sheetFormatPr");
            x.set("defaultColWidth", format!("{:.6}", sh.insert_width));
            x.set("defaultRowHeight", "15");
            x
        }
        Pend::PageSetup => make_page_setup(sh),
    }
}

/// Записывает вставляемые элементы, которые по схеме идут раньше элемента с рангом `rank`.
fn flush_pending(
    sw: &mut xml::Sw,
    sh: &SheetState,
    pending: &mut Vec<Pend>,
    rank: u32,
    werr: &dyn Fn(std::io::Error) -> anyhow::Error,
) -> Result<()> {
    let mut i = 0;
    while i < pending.len() {
        if pending[i].rank() < rank {
            let k = pending.remove(i);
            sw.elem(&make_pending(k, sh)).map_err(werr)?;
        } else {
            i += 1;
        }
    }
    Ok(())
}

#[allow(clippy::too_many_arguments)]
fn handle_top(
    sw: &mut xml::Sw,
    s: &mut Src,
    sh: &mut SheetState,
    pending: &mut Vec<Pend>,
    e: xml::Elem,
    valid: &HashSet<String>,
    active: bool,
    fit: bool,
    werr: &dyn Fn(std::io::Error) -> anyhow::Error,
) -> Result<()> {
    let rank = top_rank(&e);
    // Элемент уже есть в исходнике — вставка вместо него не нужна.
    if e.ns == xml::NS_MAIN {
        pending.retain(|k| k.local() != e.local);
    }
    flush_pending(sw, sh, pending, rank, werr)?;
    let is_page_setup = e.ns == xml::NS_MAIN && e.local == "pageSetup";
    if let Some(mut out) = top_level(s, sh, e, valid, active, fit) {
        if is_page_setup && fit {
            apply_fit(&mut out, sh);
        }
        sw.elem(&out).map_err(werr)?;
    }
    Ok(())
}

// ---------------------------------------------------------------------------
// Публичный интерфейс
// ---------------------------------------------------------------------------

/// Файл и имена листов, которые нужно взять из него (в порядке списка).
#[derive(Clone, Debug)]
pub struct Source {
    pub path: PathBuf,
    pub sheets: Vec<String>,
}

/// Параметры объединения.
#[derive(Default, Clone, Debug)]
pub struct Options {
    /// Подготовка к печати в PDF: все листы вписываются по ширине страницы (при широкой таблице —
    /// альбомная ориентация), а скрытые листы, которые выбрал пользователь, делаются видимыми.
    pub fit_to_width: bool,
}

/// Лист рабочей книги.
#[derive(Clone, Debug, PartialEq)]
pub struct SheetInfo {
    pub name: String,
    pub hidden: bool,
}

fn calamine_sheets(path: &Path) -> Result<Vec<SheetInfo>> {
    use calamine::{Reader, SheetType, SheetVisible, open_workbook_auto};
    let label = file_label(path);
    let wb = open_workbook_auto(path).map_err(|e| anyhow!("«{label}»: не удалось прочитать книгу ({e})"))?;
    Ok(wb
        .sheets_metadata()
        .iter()
        .filter(|s| s.typ == SheetType::WorkSheet)
        .map(|s| SheetInfo { name: s.name.clone(), hidden: s.visible != SheetVisible::Visible })
        .collect())
}

/// Листы книги: только рабочие листы (листы диаграмм не входят), с признаком скрытого.
/// Для .xlsx/.xlsm имена те же, что увидит `merge`; .xls/.ods читаются через calamine.
pub fn list_sheets(path: &Path) -> Result<Vec<SheetInfo>> {
    match sniff(path)? {
        Container::Zip => match Book::open(path) {
            Ok(book) => Ok(book
                .sheets
                .iter()
                .filter(|s| s.is_worksheet())
                .map(|s| SheetInfo { name: s.name.clone(), hidden: s.hidden() })
                .collect()),
            Err(e) => {
                let other = zip::ZipArchive::new(std::fs::File::open(path)?).ok().is_some_and(|z| {
                    z.file_names().filter_map(Result::ok).any(|n| n == "content.xml" || n == "xl/workbook.bin")
                });
                if other { calamine_sheets(path) } else { Err(e) }
            }
        },
        Container::Ole { encrypted: true } => Err(not_xlsx_error(path, &Container::Ole { encrypted: true })),
        Container::Ole { encrypted: false } => calamine_sheets(path),
        Container::Other => {
            // Текстовые форматы (SpreadsheetML 2003, CSV) calamine не читает; остальное пробуем.
            calamine_sheets(path).map_err(|_| not_xlsx_error(path, &Container::Other))
        }
    }
}

/// Имя листа Excel: до 31 символа, без `[]:*?/\`, уникальное без учёта регистра.
fn unique_name(used: &mut HashSet<String>, sheet: &str, file_stem: &str) -> String {
    let clean = |s: &str| -> String {
        let c: String = s.chars().filter(|c| !"[]:*?/\\".contains(*c) && !c.is_control()).take(31).collect();
        c.trim_matches('\'').to_string()
    };
    let try_insert = |used: &mut HashSet<String>, n: String| -> Option<String> {
        if !n.is_empty() && !n.eq_ignore_ascii_case("history") && used.insert(n.to_lowercase()) { Some(n) } else { None }
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
        let base = if base.is_empty() { "Лист".to_string() } else { base };
        if let Some(n) = try_insert(used, format!("{base}{suffix}")) {
            return n;
        }
    }
    unreachable!()
}

fn sanitize_for_message(s: &str) -> String {
    s.chars().filter(|c| !c.is_control()).collect()
}

/// Объединяет выбранные листы в одну книгу `.xlsx`. Возвращает предупреждения для пользователя
/// о том, что не удалось сохранить без изменений.
pub fn merge(sources: &[Source], out: &Path, opts: &Options) -> Result<Vec<String>> {
    let active: Vec<&Source> = sources.iter().filter(|s| !s.sheets.is_empty()).collect();
    if active.is_empty() {
        bail!("не выбрано ни одного листа");
    }
    let mut st = St { fit_to_width: opts.fit_to_width, ..Default::default() };
    for p in ["[Content_Types].xml", "_rels/.rels", "xl/workbook.xml", "xl/_rels/workbook.xml.rels", "xl/styles.xml", "xl/sharedStrings.xml", "xl/theme/theme1.xml"] {
        st.reserve(p);
    }

    // Проход 0: открываем книги, находим листы, назначаем имена результата.
    let mut srcs: Vec<Src> = vec![];
    let mut used_sheet_names: HashSet<String> = HashSet::new();
    let mut total = 0usize;
    for source in &active {
        let book = Book::open(&source.path)?;
        let label = book.label.clone();
        let stem = source.path.file_stem().map(|s| s.to_string_lossy().into_owned()).unwrap_or_default();
        let mut selected: Vec<SheetEntry> = vec![];
        for want in &source.sheets {
            let found = book
                .sheets
                .iter()
                .find(|e| e.name == *want)
                .or_else(|| book.sheets.iter().find(|e| e.name.to_lowercase() == want.to_lowercase()));
            match found {
                Some(e) if e.is_worksheet() => {
                    if !selected.iter().any(|x| x.part == e.part) {
                        selected.push(e.clone());
                    }
                }
                Some(e) => bail!(
                    "в файле «{label}» лист «{}» — {}, такие листы не переносятся",
                    sanitize_for_message(&e.name),
                    match e.kind.as_str() {
                        "chartsheet" => "лист диаграммы",
                        "dialogsheet" => "лист диалога",
                        "macrosheet" => "лист макросов",
                        _ => "служебный лист",
                    }
                ),
                None => bail!("в файле «{label}» нет листа «{}»", sanitize_for_message(want)),
            }
        }
        let mut out_names = vec![];
        for e in &selected {
            let n = unique_name(&mut used_sheet_names, &e.name, &stem);
            out_names.push(n);
        }
        let first = total;
        total += selected.len();
        let mut fcx = FormulaCtx::default();
        fcx.order = book.sheets.iter().map(|e| e.name.to_lowercase()).collect();
        for (i, e) in selected.iter().enumerate() {
            fcx.sheets.insert(e.name.to_lowercase(), (out_names[i].clone(), first + i));
        }
        srcs.push(Src {
            book,
            label,
            stem,
            selected,
            out_names,
            first_out_idx: first,
            styles: StyleMap::default(),
            sst: None,
            sst_map: vec![],
            fcx,
            fix: None,
            date_shift: 0.0,
            mdw: 7,
            width_scale: 1.0,
            memo: HashMap::new(),
            converted: BTreeMap::new(),
            vm_cells: 0,
            cm_cells: 0,
            neg_dates: 0,
            dropped_ext: BTreeSet::new(),
            missing_refs: BTreeMap::new(),
            has_drawings: false,
            theme_differs: false,
            table_names: HashMap::new(),
        });
    }

    // Состояние листов и активный лист: первый видимый.
    let mut states: Vec<&'static str> = vec![];
    for s in &srcs {
        for e in &s.selected {
            states.push(if opts.fit_to_width {
                "visible"
            } else {
                match e.state.as_str() {
                    "hidden" => "hidden",
                    "veryHidden" => "veryHidden",
                    _ => "visible",
                }
            });
        }
    }
    if !states.contains(&"visible") {
        states[0] = "visible";
        st.warns.add("Все выбранные листы были скрыты — первый лист сделан видимым, иначе книга не откроется");
    }
    st.active = states.iter().position(|x| *x == "visible").unwrap_or(0);

    // Проход 1: запись.
    let dir = match out.parent() {
        Some(p) if !p.as_os_str().is_empty() => p.to_path_buf(),
        _ => PathBuf::from("."),
    };
    let tmp = tempfile::Builder::new()
        .prefix(".merge-")
        .suffix(".tmp")
        .tempfile_in(&dir)
        .map_err(|e| anyhow!("не удалось создать файл результата в папке «{}» ({})", dir.display(), io_reason(&e)))?;
    let mut zip = zip::ZipWriter::new(tmp);
    for (i, s) in srcs.iter_mut().enumerate() {
        add_source(&mut zip, &mut st, s, &states, i == 0)?;
    }
    finish_package(&mut zip, &mut st)?;
    let tmp = zip.finish().map_err(|e| anyhow!("не удалось записать результат: {e}"))?;
    verify_package(tmp.path(), st.sheets.len())?;
    #[cfg(unix)]
    {
        use std::os::unix::fs::PermissionsExt;
        let _ = std::fs::set_permissions(tmp.path(), std::fs::Permissions::from_mode(0o644));
    }
    tmp.persist(out).map_err(|e| {
        anyhow!("не удалось сохранить «{}»: {} (возможно, файл открыт в другой программе)", out.display(), io_reason(&e.error))
    })?;
    Ok(st.warns.list)
}

const NS_EXT_PROPS: &str = "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties";

fn read_app_props(pkg: &mut Pkg) -> Option<(String, Option<String>)> {
    let rels = pkg.rels("").ok()?;
    let r = rels.iter().find(|r| r.kind() == "extended-properties" && !r.external)?;
    let part = resolve_target("", &r.target);
    let root = xml::parse(&pkg.read(&part).ok()?).ok()?;
    let app = root.child("Application")?.text();
    if app.trim().is_empty() {
        return None;
    }
    Some((app, root.child("AppVersion").map(|v| v.text())))
}

fn read_part_by_kind(s: &mut Src, kind: &str) -> Result<Option<(String, Vec<u8>)>> {
    let Some(r) = s.book.rel_of_kind(kind) else { return Ok(None) };
    let part = resolve_target(&s.book.wb_part, &r.target);
    if !s.book.pkg.has(&part) {
        return Ok(None);
    }
    let b = s.book.pkg.read(&part).map_err(|e| anyhow!("«{}»: {e}", s.label))?;
    Ok(Some((part, b)))
}

fn add_source(zip: &mut Zw, st: &mut St, s: &mut Src, states: &[&'static str], first_source: bool) -> Result<()> {
    let label = s.label.clone();
    if first_source {
        st.wb_pr = s.book.wb.child("workbookPr").cloned();
        st.wb_view = s.book.wb.child("bookViews").and_then(|b| b.child("workbookView")).cloned();
        st.calc_pr = s.book.wb.child("calcPr").cloned().map(|mut c| {
            c.remove("fullCalcOnLoad");
            c
        });
        st.app = read_app_props(&mut s.book.pkg);
    }

    // Что не переносится вообще.
    if s.book.pkg.has("xl/vbaProject.bin") || s.book.pkg.content_type(&s.book.wb_part).is_some_and(|c| c.contains("macroEnabled")) {
        st.warns.add(format!("«{label}»: макросы (VBA) не сохраняются — результат записывается как обычная книга .xlsx без макросов"));
    }
    for e in s.book.sheets.iter().filter(|e| !e.is_worksheet()) {
        let what = match e.kind.as_str() {
            "chartsheet" => "лист диаграммы",
            "dialogsheet" => "лист диалога",
            "macrosheet" => "лист макросов",
            _ => "служебный лист",
        };
        st.warns.add(format!("«{label}»: {what} «{}» не перенесён — переносятся только обычные листы", sanitize_for_message(&e.name)));
    }
    if s.book.pkg.has("xl/connections.xml") || s.book.pkg.names.keys().any(|n| n.starts_with("xl/model/")) {
        st.warns.add(format!("«{label}»: подключения к внешним данным и модель данных (Power Pivot/Power Query) не переносятся"));
    }
    if s.book.wb.child("workbookProtection").is_some_and(|p| {
        p.attrs.iter().any(|a| matches!(a.q.as_str(), "lockStructure" | "lockWindows" | "workbookPassword" | "workbookHashValue") && a.v != "0" && a.v != "false")
    }) {
        st.warns.add(format!("«{label}»: защита структуры книги не переносится"));
    }
    if s.book.wb.child("extLst").is_some_and(|x| {
        let mut slicer = false;
        x.walk(&mut |e| {
            if e.local.to_lowercase().contains("slicer") || e.local.to_lowercase().contains("timeline") {
                slicer = true;
            }
        });
        slicer
    }) {
        s.dropped_ext.insert("срезы и временные шкалы".into());
    }

    // Система дат.
    let src_1904 = s.book.date1904();
    if first_source {
        st.date1904 = src_1904;
    } else if src_1904 != st.date1904 {
        s.date_shift = if src_1904 { 1462.0 } else { -1462.0 };
        st.recalc = true;
        st.warns.add(format!(
            "«{label}»: книга использует систему дат {} — даты пересчитаны под систему первой книги ({}); числа-константы в формулах и условиях не пересчитываются",
            if src_1904 { "1904" } else { "1900" },
            if st.date1904 { "1904" } else { "1900" }
        ));
    }

    // Тема.
    let theme = read_part_by_kind(s, "theme")?;
    if let Some((_, bytes)) = &theme {
        if let Some(t) = parse_theme(bytes) {
            match &st.theme {
                None => st.theme = Some((t, bytes.clone())),
                Some((ot, _)) if *ot != t => {
                    s.fix = Some(ThemeFix { colors: t.colors != ot.colors, fonts: t.major != ot.major || t.minor != ot.minor, src: t });
                    s.theme_differs = true;
                }
                _ => {}
            }
        }
    }

    // Стили.
    let styles_root = match read_part_by_kind(s, "styles")? {
        Some((_, b)) => xml::parse(&b).map_err(|e| anyhow!("«{label}»: повреждён styles.xml ({e})"))?,
        None => default_styles_root(),
    };
    let (font_name, font_size) = normal_font(&styles_root);
    s.mdw = max_digit_width(&font_name, font_size);
    if first_source {
        st.base_mdw = s.mdw;
    } else if s.mdw != st.base_mdw {
        // Ширина столбцов измеряется в цифрах шрифта «Обычный»; у книги он другой — пересчитываем.
        s.width_scale = s.mdw as f64 / st.base_mdw as f64;
    }
    s.styles = st.styles.add_source(&styles_root, s.fix.as_ref(), &mut st.warns, &label)?;

    // Внешние книги.
    s.fcx.ext_offset = st.ext_links.len();
    let ext_ids: Vec<String> = s
        .book
        .wb
        .child("externalReferences")
        .map(|x| x.children().filter_map(|e| e.attr_ns(xml::NS_R, "id").map(str::to_string)).collect())
        .unwrap_or_default();
    s.fcx.ext_count = ext_ids.len();
    for rid in ext_ids {
        let Some(rel) = s.book.wb_rels.iter().find(|r| r.id == rid && !r.external).cloned() else { continue };
        let part = resolve_target(&s.book.wb_part, &rel.target);
        if let Some(np) = copy_part(st, s, &part, "externalLink")? {
            st.ext_links.push(np);
        }
    }
    if s.fcx.ext_offset != st.ext_links.len() {
        st.warns.add(format!("«{label}»: в книге есть ссылки на другие книги — они сохранены как есть и могут не работать из нового расположения файла"));
    }

    prescan_tables(st, s);
    process_names(st, s);

    // Листы.
    for i in 0..s.selected.len() {
        let entry = s.selected[i].clone();
        let new_name = s.out_names[i].clone();
        let out_idx = s.first_out_idx + i;
        if sanitize_for_message(&new_name) != entry.name && !new_name.contains('_') {
            // имя изменено не префиксом, а очисткой
        }
        let clean_changed = entry.name.chars().count() > 31 || entry.name.chars().any(|c| "[]:*?/\\".contains(c));
        if clean_changed {
            st.warns.add(format!("«{label}»: имя листа «{}» недопустимо для Excel — в результате лист называется «{new_name}»", sanitize_for_message(&entry.name)));
        }
        transform_sheet(zip, st, s, &entry, out_idx, &new_name, states[out_idx], out_idx == st.active)?;
    }
    flush(zip, st)?;
    emit_source_warnings(st, s);
    Ok(())
}

fn list_short(items: &BTreeSet<String>) -> String {
    let v: Vec<String> = items.iter().take(5).map(|x| format!("«{x}»")).collect();
    let mut r = v.join(", ");
    if items.len() > 5 {
        r.push_str(&format!(" и ещё {}", items.len() - 5));
    }
    r
}

fn emit_source_warnings(st: &mut St, s: &mut Src) {
    let label = s.label.clone();
    for (sheet, (n, missing)) in &s.converted {
        st.warns.add(format!(
            "«{label}», лист «{sheet}»: {n} формул ссылаются на листы или таблицы, которые не вошли в результат ({}) — они заменены вычисленными значениями",
            list_short(missing)
        ));
    }
    for (place, items) in &s.missing_refs {
        if place.contains("формулы без сохранённого значения") {
            st.warns.add(format!("«{label}», {place}: {} ячеек остались пустыми, так как в файле не было вычисленных значений", list_short(items).replace(['«', '»'], "")));
            continue;
        }
        st.warns.add(format!("«{label}», {place}: есть ссылки на листы или таблицы, которых нет в результате ({}) — такие ссылки работать не будут", list_short(items)));
    }
    if s.cm_cells > 0 {
        st.warns.add(format!("«{label}»: динамические формулы ({} ячеек) сохранены как обычные формулы массива", s.cm_cells));
    }
    if s.vm_cells > 0 {
        st.warns.add(format!("«{label}»: картинки в ячейках и расширенные типы данных ({} ячеек) не сохраняются", s.vm_cells));
    }
    if s.neg_dates > 0 {
        st.warns.add(format!("«{label}»: {} дат раньше 1904 года нельзя пересчитать в систему дат первой книги — значения оставлены как есть", s.neg_dates));
    }
    if !s.dropped_ext.is_empty() {
        let list: Vec<&str> = s.dropped_ext.iter().map(String::as_str).collect();
        st.warns.add(format!("«{label}»: не перенесено: {}", list.join("; ")));
    }
    if s.theme_differs && s.has_drawings {
        st.warns.add(format!("«{label}»: тема оформления отличается от темы первой книги — цвета ячеек сохранены, но цвета диаграмм и фигур могут измениться"));
    }
}

/// Проверяет имена и ссылки таблиц источника и назначает таблицам уникальные имена.
fn prescan_tables(st: &mut St, s: &mut Src) {
    let entries: Vec<SheetEntry> = s.book.sheets.iter().filter(|e| e.is_worksheet()).cloned().collect();
    for e in entries {
        let selected = s.fcx.sheets.contains_key(&e.name.to_lowercase());
        let Ok(rels) = s.book.pkg.rels(&e.part) else { continue };
        for r in rels.iter().filter(|r| r.kind() == "table" && !r.external) {
            let part = resolve_target(&e.part, &r.target);
            let Ok(bytes) = s.book.pkg.read(&part) else { continue };
            let Ok(root) = xml::parse(&bytes) else { continue };
            let name = root.attr("name").unwrap_or("").to_string();
            let display = root.attr("displayName").unwrap_or(&name).to_string();
            if display.is_empty() {
                continue;
            }
            let lname = display.to_lowercase();
            if !selected {
                s.fcx.tables.insert(lname, None);
                if !name.is_empty() {
                    s.fcx.tables.insert(name.to_lowercase(), None);
                }
                continue;
            }
            let new = if st.names_used.contains(&lname) {
                let n = unique_ident(&st.names_used, &display, &s.stem);
                st.warns.add(format!("«{}»: таблица «{display}» переименована в «{n}», так как такое имя уже занято", s.label));
                n
            } else {
                display.clone()
            };
            st.names_used.insert(new.to_lowercase());
            s.fcx.tables.insert(lname, Some(new.clone()));
            let new_name = if name.is_empty() || name.eq_ignore_ascii_case(&display) { new.clone() } else { name.clone() };
            if !name.is_empty() && !name.eq_ignore_ascii_case(&display) {
                s.fcx.tables.insert(name.to_lowercase(), Some(new.clone()));
            }
            s.table_names.insert(part.to_lowercase(), (new, new_name));
        }
    }
}

fn is_macro_name(e: &xml::Elem) -> bool {
    matches!(e.attr("vbProcedure"), Some("1" | "true")) || matches!(e.attr("xlm"), Some("1" | "true")) || matches!(e.attr("function"), Some("1" | "true"))
}

/// Именованные диапазоны: привязка к листам результата, уникальность, переписывание ссылок.
fn process_names(st: &mut St, s: &mut Src) {
    let Some(dn) = s.book.wb.child("definedNames") else { return };
    let items: Vec<xml::Elem> = dn.children().filter(|e| e.local == "definedName").cloned().collect();
    let own: HashSet<String> = items
        .iter()
        .filter(|e| e.attr("localSheetId").is_none())
        .filter_map(|e| e.attr("name").map(str::to_lowercase))
        .collect();
    let mut dup: HashSet<String> = HashSet::new();
    // Проход 1: решаем, какие глобальные имена переименовать.
    for e in items.iter().filter(|e| e.attr("localSheetId").is_none()) {
        let Some(name) = e.attr("name") else { continue };
        if name.starts_with("_xlnm.") || is_macro_name(e) {
            continue;
        }
        let lname = name.to_lowercase();
        if st.names_used.contains(&lname) {
            let text = s.fcx.rewrite(&e.text(), Missing::Ref).text;
            if st.global_names.get(&lname) == Some(&text) {
                dup.insert(lname);
                continue;
            }
            let mut used = st.names_used.clone();
            used.extend(own.iter().cloned());
            used.extend(s.fcx.names.values().map(|v| v.to_lowercase()));
            let new = unique_ident(&used, name, &s.stem);
            st.warns.add(format!("«{}»: имя «{name}» уже есть в результате — переименовано в «{new}»", s.label));
            s.fcx.names.insert(lname, new);
        }
    }
    // Проход 2: перенос.
    let mut missing_all: BTreeSet<String> = BTreeSet::new();
    let mut emitted: HashSet<(String, Option<usize>)> = HashSet::new();
    for e in &items {
        let Some(name) = e.attr("name") else { continue };
        if is_macro_name(e) {
            st.warns.add(format!("«{}»: имя «{name}» связано с макросом и не перенесено", s.label));
            continue;
        }
        let local = e.attr("localSheetId").and_then(|v| v.parse::<usize>().ok());
        let (final_name, local_out) = match local {
            Some(l) => {
                let Some(entry) = s.book.sheets.get(l) else { continue };
                let Some((_, out_idx)) = s.fcx.sheets.get(&entry.name.to_lowercase()) else { continue };
                (name.to_string(), Some(*out_idx))
            }
            None => {
                if name.starts_with("_xlnm.") {
                    continue;
                }
                let l = name.to_lowercase();
                if dup.contains(&l) {
                    continue;
                }
                (s.fcx.names.get(&l).cloned().unwrap_or_else(|| name.to_string()), None)
            }
        };
        if !emitted.insert((final_name.to_lowercase(), local_out)) {
            continue; // повторное определение того же имени в исходном файле
        }
        let r = s.fcx.rewrite(&e.text(), Missing::Ref);
        missing_all.extend(r.missing);
        let attrs: Vec<(String, String)> = e
            .attrs
            .iter()
            .filter(|a| a.ns.is_empty() && a.q != "name" && a.q != "localSheetId" && !a.q.starts_with("xmlns"))
            .map(|a| (a.q.clone(), a.v.clone()))
            .collect();
        if local_out.is_none() {
            st.names_used.insert(final_name.to_lowercase());
            st.global_names.insert(final_name.to_lowercase(), r.text.clone());
        }
        st.defined.push(DefName { name: final_name, local: local_out, attrs, text: r.text });
    }
    if !missing_all.is_empty() {
        st.recalc = true;
        st.warns.add(format!(
            "«{}»: именованные диапазоны ссылаются на листы, которых нет в результате ({}) — ссылки заменены на #REF!",
            s.label,
            list_short(&missing_all)
        ));
    }
}

// ---------------------------------------------------------------------------
// Завершение пакета и проверка
// ---------------------------------------------------------------------------

fn finish_package(zip: &mut Zw, st: &mut St) -> Result<()> {
    // Стили, строки, тема.
    let styles = st.styles.finish()?;
    st.put("xl/styles.xml", Some(CT_STYLES), styles);
    let mut sst = format!(
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n<sst xmlns=\"{}\" count=\"{}\" uniqueCount=\"{}\">",
        xml::NS_MAIN,
        st.sst_refs,
        st.sst.items.len()
    );
    for i in &st.sst.items {
        sst.push_str(i);
    }
    sst.push_str("</sst>");
    st.put("xl/sharedStrings.xml", Some(CT_SST), sst.into_bytes());
    let theme_present = if let Some((_, bytes)) = st.theme.take() {
        st.put("xl/theme/theme1.xml", Some(CT_THEME), bytes);
        true
    } else {
        false
    };

    // Связи книги.
    let mut rels: Vec<Rel> = vec![];
    let mut next = 1;
    let rid = |next: &mut usize| {
        let r = format!("rId{next}");
        *next += 1;
        r
    };
    let mut sheet_rids = vec![];
    for sh in &st.sheets {
        let id = rid(&mut next);
        rels.push(Rel { id: id.clone(), typ: format!("{REL_BASE}worksheet"), target: relative_target("xl/workbook.xml", &sh.part), external: false });
        sheet_rids.push(id);
    }
    let add_rel = |rels: &mut Vec<Rel>, next: &mut usize, kind: &str, target: &str| -> String {
        let id = rid(next);
        rels.push(Rel { id: id.clone(), typ: format!("{REL_BASE}{kind}"), target: relative_target("xl/workbook.xml", target), external: false });
        id
    };
    add_rel(&mut rels, &mut next, "styles", "xl/styles.xml");
    add_rel(&mut rels, &mut next, "sharedStrings", "xl/sharedStrings.xml");
    if theme_present {
        add_rel(&mut rels, &mut next, "theme", "xl/theme/theme1.xml");
    }
    let ext_ids: Vec<String> = st.ext_links.clone().iter().map(|p| add_rel(&mut rels, &mut next, "externalLink", p)).collect();
    let cache_ids: Vec<(u32, String)> = st.pivot_caches.clone().iter().map(|(id, p)| (*id, add_rel(&mut rels, &mut next, "pivotCacheDefinition", p))).collect();

    // workbook.xml
    let mut wb = xml::Elem::main("workbook");
    wb.attrs.push(xml::Attr { q: "xmlns:r".into(), ns: String::new(), local: String::new(), v: xml::NS_R.into() });
    let mut pr = st.wb_pr.take().unwrap_or_else(|| xml::Elem::main("workbookPr"));
    pr.remove("codeName");
    pr.remove("date1904");
    if st.date1904 {
        pr.set("date1904", "1");
    }
    wb.push(pr);
    let mut view = st.wb_view.take().unwrap_or_else(|| xml::Elem::main("workbookView"));
    view.set("activeTab", st.active.to_string());
    view.remove("firstSheet");
    wb.kids.push(xml::Node::E({
        let mut v = xml::Elem::main("bookViews");
        v.push(view);
        v
    }));
    let mut sheets = xml::Elem::main("sheets");
    for (i, sh) in st.sheets.iter().enumerate() {
        let mut e = xml::Elem::main("sheet");
        e.set("name", sh.name.as_str());
        e.set("sheetId", (i + 1).to_string());
        if sh.state != "visible" {
            e.set("state", sh.state);
        }
        e.set_ns(xml::NS_R, "r", "id", sheet_rids[i].as_str());
        sheets.push(e);
    }
    wb.push(sheets);
    if !ext_ids.is_empty() {
        let mut x = xml::Elem::main("externalReferences");
        for id in &ext_ids {
            let mut e = xml::Elem::main("externalReference");
            e.set_ns(xml::NS_R, "r", "id", id.as_str());
            x.push(e);
        }
        wb.push(x);
    }
    if !st.defined.is_empty() {
        let mut dn = xml::Elem::main("definedNames");
        for d in &st.defined {
            let mut e = xml::Elem::main("definedName");
            e.set("name", d.name.as_str());
            if let Some(l) = d.local {
                e.set("localSheetId", l.to_string());
            }
            for (k, v) in &d.attrs {
                e.set(k, v.as_str());
            }
            e.set_text(d.text.as_str());
            dn.push(e);
        }
        wb.push(dn);
    }
    let mut calc = st.calc_pr.take().unwrap_or_else(|| xml::Elem::main("calcPr"));
    if st.recalc {
        calc.set("fullCalcOnLoad", "1");
    }
    wb.push(calc);
    if !cache_ids.is_empty() {
        let mut x = xml::Elem::main("pivotCaches");
        for (id, rid) in &cache_ids {
            let mut e = xml::Elem::main("pivotCache");
            e.set("cacheId", id.to_string());
            e.set_ns(xml::NS_R, "r", "id", rid.as_str());
            x.push(e);
        }
        wb.push(x);
    }
    st.put("xl/workbook.xml", Some(CT_WORKBOOK), xml::doc(&wb, xml::NS_MAIN));
    st.put("xl/_rels/workbook.xml.rels", None, rels_xml(&rels));
    let mut root_rels = vec![Rel { id: "rId1".into(), typ: REL_OFFICE_DOC.into(), target: "xl/workbook.xml".into(), external: false }];
    // Application/AppVersion первой книги: LibreOffice по ним решает, доверять ли высоте строк.
    if let Some((app, ver)) = st.app.take() {
        let mut p = xml::Elem::new(NS_EXT_PROPS, "Properties");
        let mut a = xml::Elem::new(NS_EXT_PROPS, "Application");
        a.set_text(app);
        p.push(a);
        if let Some(v) = ver {
            let mut a = xml::Elem::new(NS_EXT_PROPS, "AppVersion");
            a.set_text(v);
            p.push(a);
        }
        st.put("docProps/app.xml", Some("application/vnd.openxmlformats-officedocument.extended-properties+xml"), xml::doc(&p, NS_EXT_PROPS));
        root_rels.push(Rel {
            id: "rId2".into(),
            typ: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties".into(),
            target: "docProps/app.xml".into(),
            external: false,
        });
    }
    st.put("_rels/.rels", None, rels_xml(&root_rels));

    // [Content_Types].xml
    let mut ct = xml::Elem::new(xml::NS_CT, "Types");
    let def = |ext: &str, t: &str| {
        let mut e = xml::Elem::new(xml::NS_CT, "Default");
        e.set("Extension", ext);
        e.set("ContentType", t);
        e
    };
    ct.push(def("rels", CT_RELS));
    ct.push(def("xml", "application/xml"));
    let mut defaults: Vec<(&String, &String)> = st.defaults.iter().collect();
    defaults.sort();
    for (ext, t) in defaults {
        ct.push(def(ext, t));
    }
    let mut seen = HashSet::new();
    for (part, t) in &st.overrides {
        if !seen.insert(part.to_lowercase()) {
            continue;
        }
        let mut e = xml::Elem::new(xml::NS_CT, "Override");
        e.set("PartName", format!("/{part}"));
        e.set("ContentType", t.as_str());
        ct.push(e);
    }
    st.put("[Content_Types].xml", None, xml::doc(&ct, xml::NS_CT));
    flush(zip, st)
}

/// Проверка готового пакета: целостность архива, XML, связи, типы содержимого.
fn verify_package(path: &Path, expected_sheets: usize) -> Result<()> {
    let mut pkg = Pkg::open(path).map_err(|e| anyhow!("внутренняя ошибка: результат не прошёл проверку ({e})"))?;
    let fail = |m: String| anyhow!("внутренняя ошибка: результат не прошёл проверку ({m}); исходные файлы не изменены");
    let names: Vec<String> = pkg.names.keys().cloned().collect();
    for lname in &names {
        let real = lname.as_str();
        if lname.ends_with(".xml") || lname.ends_with(".rels") {
            let bytes = pkg.read(real).map_err(|e| fail(e.to_string()))?;
            let text = xml::decode(&bytes).map_err(|e| fail(format!("{lname}: {e}")))?;
            let mut p = xml::Parser::new(&text);
            loop {
                match p.next() {
                    Ok(xml::Ev::Eof) => break,
                    Ok(_) => {}
                    Err(e) => return Err(fail(format!("{lname}: {e}"))),
                }
            }
        }
        if lname != "[content_types].xml" && !lname.ends_with(".rels") && pkg.content_type(lname).is_none() {
            return Err(fail(format!("для части {lname} не задан тип содержимого")));
        }
        if lname.ends_with(".rels") {
            let owner = {
                let n = lname.replace("/_rels/", "/");
                let n = n.strip_prefix("_rels/").unwrap_or(&n).to_string();
                n.trim_end_matches(".rels").to_string()
            };
            let rels = pkg.rels(&owner).map_err(|e| fail(e.to_string()))?;
            for r in rels.iter().filter(|r| !r.external && !r.target.starts_with('#')) {
                let t = resolve_target(&owner, &r.target);
                if !pkg.has(&t) {
                    return Err(fail(format!("связь {} из {owner} указывает на отсутствующую часть {t}", r.id)));
                }
            }
        }
    }
    let book = Book::open(path).map_err(|e| fail(e.to_string()))?;
    if book.sheets.len() != expected_sheets {
        return Err(fail(format!("листов {} вместо {expected_sheets}", book.sheets.len())));
    }
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;

    const WB_NS: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    const R_NS: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    const STYLES: &str = "<styleSheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">\
        <fonts count=\"2\"><font><sz val=\"11\"/><name val=\"Calibri\"/></font><font><b/><sz val=\"11\"/><color theme=\"4\"/><name val=\"Calibri\"/></font></fonts>\
        <fills count=\"2\"><fill><patternFill patternType=\"none\"/></fill><fill><patternFill patternType=\"gray125\"/></fill></fills>\
        <borders count=\"1\"><border><left/><right/><top/><bottom/><diagonal/></border></borders>\
        <cellStyleXfs count=\"1\"><xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\"/></cellStyleXfs>\
        <cellXfs count=\"3\"><xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\"/>\
        <xf numFmtId=\"14\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\" applyNumberFormat=\"1\"/>\
        <xf numFmtId=\"0\" fontId=\"1\" fillId=\"0\" borderId=\"0\" xfId=\"0\" applyFont=\"1\"/></cellXfs>\
        <cellStyles count=\"1\"><cellStyle name=\"Normal\" xfId=\"0\" builtinId=\"0\"/></cellStyles>\
        <dxfs count=\"1\"><dxf><font><b/></font></dxf></dxfs></styleSheet>";

    fn theme(accent1: &str) -> String {
        format!(
            "<a:theme xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" name=\"T\"><a:themeElements><a:clrScheme name=\"c\">\
             <a:dk1><a:sysClr val=\"windowText\" lastClr=\"000000\"/></a:dk1><a:lt1><a:sysClr val=\"window\" lastClr=\"FFFFFF\"/></a:lt1>\
             <a:dk2><a:srgbClr val=\"1F497D\"/></a:dk2><a:lt2><a:srgbClr val=\"EEECE1\"/></a:lt2>\
             <a:accent1><a:srgbClr val=\"{accent1}\"/></a:accent1><a:accent2><a:srgbClr val=\"C0504D\"/></a:accent2>\
             <a:accent3><a:srgbClr val=\"9BBB59\"/></a:accent3><a:accent4><a:srgbClr val=\"8064A2\"/></a:accent4>\
             <a:accent5><a:srgbClr val=\"4BACC6\"/></a:accent5><a:accent6><a:srgbClr val=\"F79646\"/></a:accent6>\
             <a:hlink><a:srgbClr val=\"0000FF\"/></a:hlink><a:folHlink><a:srgbClr val=\"800080\"/></a:folHlink></a:clrScheme>\
             <a:fontScheme name=\"f\"><a:majorFont><a:latin typeface=\"Cambria\"/></a:majorFont><a:minorFont><a:latin typeface=\"Calibri\"/></a:minorFont></a:fontScheme>\
             </a:themeElements></a:theme>"
        )
    }

    /// Лист для тестов: (имя, состояние, xml листа, связи листа).
    struct Sh {
        name: String,
        state: &'static str,
        xml: String,
        rels: Vec<(String, String, String)>,
    }

    fn sheet_xml(rows: &str, tail: &str) -> String {
        format!("<worksheet xmlns=\"{WB_NS}\" xmlns:r=\"{R_NS}\"><sheetData>{rows}</sheetData>{tail}</worksheet>")
    }

    fn sh(name: &str, rows: &str, tail: &str) -> Sh {
        Sh { name: name.into(), state: "visible", xml: sheet_xml(rows, tail), rels: vec![] }
    }

    fn cs(r: &str, text: &str) -> String {
        format!("<c r=\"{r}\" t=\"inlineStr\"><is><t>{text}</t></is></c>")
    }
    fn cn(r: &str, v: f64) -> String {
        format!("<c r=\"{r}\"><v>{v}</v></c>")
    }
    fn cf(r: &str, f: &str, v: f64) -> String {
        format!("<c r=\"{r}\"><f>{f}</f><v>{v}</v></c>")
    }
    fn row(n: u32, cells: &[String]) -> String {
        format!("<row r=\"{n}\">{}</row>", cells.concat())
    }

    #[derive(Default)]
    struct Fx {
        sheets: Vec<Sh>,
        names: String,
        wb_pr: String,
        extra: Vec<(String, Vec<u8>, Option<String>)>,
        wb_rels: Vec<(String, String, String)>,
        theme: Option<String>,
        wb_tail: String,
        styles: Option<String>,
    }

    fn write_zip(path: &Path, parts: &[(String, Vec<u8>)]) {
        let f = std::fs::File::create(path).unwrap();
        let mut z = zip::ZipWriter::new(f);
        for (n, b) in parts {
            z.start_file(n.as_str(), zip_opts()).unwrap();
            z.write_all(b).unwrap();
        }
        z.finish().unwrap();
    }

    impl Fx {
        fn build(&self, path: &Path) {
            let mut parts: Vec<(String, Vec<u8>)> = vec![];
            let mut ct = String::from(
                "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\"><Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>\
                 <Default Extension=\"xml\" ContentType=\"application/xml\"/><Default Extension=\"png\" ContentType=\"image/png\"/>\
                 <Override PartName=\"/xl/workbook.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml\"/>\
                 <Override PartName=\"/xl/styles.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml\"/>",
            );
            let mut wb = format!("<workbook xmlns=\"{WB_NS}\" xmlns:r=\"{R_NS}\"><workbookPr {}/><sheets>", self.wb_pr);
            let mut rels = String::from("<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">");
            for (i, s) in self.sheets.iter().enumerate() {
                let st = if s.state == "visible" { String::new() } else { format!(" state=\"{}\"", s.state) };
                wb.push_str(&format!("<sheet name=\"{}\" sheetId=\"{}\"{st} r:id=\"rId{}\"/>", s.name, i + 1, i + 1));
                if s.xml.starts_with("<chartsheet") {
                    rels.push_str(&format!("<Relationship Id=\"rId{}\" Type=\"{REL_BASE}chartsheet\" Target=\"chartsheets/sheet{}.xml\"/>", i + 1, i + 1));
                    parts.push((format!("xl/chartsheets/sheet{}.xml", i + 1), s.xml.clone().into_bytes()));
                    ct.push_str(&format!("<Override PartName=\"/xl/chartsheets/sheet{}.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.chartsheet+xml\"/>", i + 1));
                    continue;
                }
                rels.push_str(&format!("<Relationship Id=\"rId{}\" Type=\"{REL_BASE}worksheet\" Target=\"worksheets/sheet{}.xml\"/>", i + 1, i + 1));
                parts.push((format!("xl/worksheets/sheet{}.xml", i + 1), s.xml.clone().into_bytes()));
                ct.push_str(&format!("<Override PartName=\"/xl/worksheets/sheet{}.xml\" ContentType=\"{CT_WORKSHEET}\"/>", i + 1));
                if !s.rels.is_empty() {
                    let mut r = String::from("<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">");
                    for (id, t, tg) in &s.rels {
                        r.push_str(&format!("<Relationship Id=\"{id}\" Type=\"{t}\" Target=\"{tg}\"/>"));
                    }
                    r.push_str("</Relationships>");
                    parts.push((format!("xl/worksheets/_rels/sheet{}.xml.rels", i + 1), r.into_bytes()));
                }
            }
            wb.push_str("</sheets>");
            wb.push_str(&self.wb_tail);
            if !self.names.is_empty() {
                wb.push_str(&format!("<definedNames>{}</definedNames>", self.names));
            }
            wb.push_str("</workbook>");
            let n = self.sheets.len();
            rels.push_str(&format!("<Relationship Id=\"rId{}\" Type=\"{REL_BASE}styles\" Target=\"styles.xml\"/>", n + 1));
            if self.theme.is_some() {
                rels.push_str(&format!("<Relationship Id=\"rId{}\" Type=\"{REL_BASE}theme\" Target=\"theme/theme1.xml\"/>", n + 2));
                ct.push_str(&format!("<Override PartName=\"/xl/theme/theme1.xml\" ContentType=\"{CT_THEME}\"/>"));
            }
            for (id, t, tg) in &self.wb_rels {
                rels.push_str(&format!("<Relationship Id=\"{id}\" Type=\"{t}\" Target=\"{tg}\"/>"));
            }
            rels.push_str("</Relationships>");
            for (name, bytes, ctype) in &self.extra {
                if let Some(c) = ctype {
                    ct.push_str(&format!("<Override PartName=\"/{name}\" ContentType=\"{c}\"/>"));
                }
                parts.push((name.clone(), bytes.clone()));
            }
            ct.push_str("</Types>");
            parts.push(("[Content_Types].xml".into(), ct.into_bytes()));
            parts.push((
                "_rels/.rels".into(),
                format!("<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"rId1\" Type=\"{REL_OFFICE_DOC}\" Target=\"xl/workbook.xml\"/></Relationships>").into_bytes(),
            ));
            parts.push(("xl/workbook.xml".into(), wb.into_bytes()));
            parts.push(("xl/_rels/workbook.xml.rels".into(), rels.into_bytes()));
            parts.push(("xl/styles.xml".into(), self.styles.clone().unwrap_or_else(|| STYLES.into()).into_bytes()));
            if let Some(t) = &self.theme {
                parts.push(("xl/theme/theme1.xml".into(), t.clone().into_bytes()));
            }
            write_zip(path, &parts);
        }
    }

    fn part(path: &Path, name: &str) -> String {
        let mut z = zip::ZipArchive::new(std::fs::File::open(path).unwrap()).unwrap();
        let mut s = String::new();
        std::io::Read::read_to_string(&mut z.by_name(name).unwrap_or_else(|_| panic!("нет части {name}")), &mut s).unwrap();
        s
    }

    fn part_bytes(path: &Path, name: &str) -> Vec<u8> {
        let mut z = zip::ZipArchive::new(std::fs::File::open(path).unwrap()).unwrap();
        let mut v = vec![];
        std::io::Read::read_to_end(&mut z.by_name(name).unwrap(), &mut v).unwrap();
        v
    }

    fn has_part(path: &Path, name: &str) -> bool {
        let z = zip::ZipArchive::new(std::fs::File::open(path).unwrap()).unwrap();
        z.file_names().filter_map(Result::ok).any(|n| n == name)
    }

    fn src(p: &Path, sheets: &[&str]) -> Source {
        Source { path: p.to_path_buf(), sheets: sheets.iter().map(|s| s.to_string()).collect() }
    }

    fn ctx(map: &[(&str, &str)]) -> FormulaCtx {
        let mut c = FormulaCtx::default();
        for (i, (old, new)) in map.iter().enumerate() {
            c.sheets.insert(old.to_lowercase(), (new.to_string(), i));
            c.order.push(old.to_lowercase());
        }
        c
    }

    fn rw(c: &FormulaCtx, f: &str) -> String {
        c.rewrite(f, Missing::Keep).text
    }

    // ---------------------------------------------------------------- имена листов

    #[test]
    fn sheet_names_are_valid() {
        let mut used = HashSet::new();
        assert_eq!(unique_name(&mut used, "a/b", "f"), "ab");
        assert_eq!(unique_name(&mut used, "AB", "f"), "f_AB");
        assert_eq!(unique_name(&mut used, "AB", "f"), "AB_2");
        assert_eq!(unique_name(&mut used, &"x".repeat(40), "f").chars().count(), 31);
        assert_eq!(unique_name(&mut used, "History", "f"), "f_History");
    }

    // ---------------------------------------------------------------- формулы

    #[test]
    fn rewrites_sheet_references() {
        let c = ctx(&[("Data", "b_Data"), ("It's", "b_It's"), ("Other", "Other")]);
        assert_eq!(rw(&c, "Data!A1+1"), "'b_Data'!A1+1");
        assert_eq!(rw(&c, "SUM('Data'!A1:B2)"), "SUM('b_Data'!A1:B2)");
        assert_eq!(rw(&c, "'It''s'!A1"), "'b_It''s'!A1");
        assert_eq!(rw(&c, "Other!A1+Data!$B$2"), "Other!A1+'b_Data'!$B$2");
        assert_eq!(rw(&c, "\"Data!A1\"&Data!A1"), "\"Data!A1\"&'b_Data'!A1");
        assert_eq!(rw(&c, "Data!$A$1:Data!$B$2"), "'b_Data'!$A$1:'b_Data'!$B$2");
        assert_eq!(rw(&c, "DATA!A1"), "'b_Data'!A1", "регистр не важен");
        assert_eq!(rw(&c, "#REF!+Data!A1"), "#REF!+'b_Data'!A1");
        assert_eq!(rw(&c, "1E5+Data!A1"), "1E5+'b_Data'!A1");
        assert_eq!(rw(&c, "Datax!A1"), "Datax!A1", "чужой лист остаётся");
    }

    #[test]
    fn missing_sheets_are_reported() {
        let c = ctx(&[("Data", "Data")]);
        let r = c.rewrite("Data!A1+Gone!B2", Missing::Keep);
        assert_eq!(r.text, "Data!A1+Gone!B2");
        assert_eq!(r.missing, ["Gone"]);
        let r = c.rewrite("Data!A1+'Gone x'!$B$2:$C$3*2", Missing::Ref);
        assert_eq!(r.text, "Data!A1+#REF!*2");
        assert_eq!(r.missing, ["Gone x"]);
    }

    #[test]
    fn rewrites_3d_names_tables_and_external() {
        let mut c = ctx(&[("S1", "N1"), ("S2", "N2"), ("S3", "N3")]);
        assert_eq!(rw(&c, "SUM(S1:S3!A1)"), "SUM('N1:N3'!A1)");
        c.sheets.get_mut("s2").unwrap().1 = 7; // порядок нарушен — ссылка 3D уже не диапазон
        assert_eq!(c.rewrite("SUM(S1:S3!A1)", Missing::Keep).missing, ["S1:S3"]);
        let mut c = ctx(&[("Data", "Data")]);
        c.names.insert("total".into(), "Total_b".into());
        c.tables.insert("tbl".into(), Some("Tbl_b".into()));
        c.tables.insert("gone".into(), None);
        c.ext_offset = 2;
        c.ext_count = 1;
        assert_eq!(rw(&c, "Total*2+SUM(Tbl[Qty])+Tbl[[#This Row],[a]]"), "Total_b*2+SUM(Tbl_b[Qty])+Tbl_b[[#This Row],[a]]");
        assert_eq!(rw(&c, "Total(1)+Data!Total"), "Total(1)+Data!Total");
        assert_eq!(rw(&c, "[1]Sheet1!A1+'[1]Sheet 2'!B2+[1]!Nm"), "[3]Sheet1!A1+'[3]Sheet 2'!B2+[3]!Nm");
        assert_eq!(c.rewrite("[2]Sheet1!A1", Missing::Keep).missing, ["[2]Sheet1"], "нет такой внешней книги");
        assert_eq!(c.rewrite("SUM(Gone[x])", Missing::Keep).missing, ["Gone"]);
        assert_eq!(rw(&c, "[@Col]*2"), "[@Col]*2");
    }

    // ---------------------------------------------------------------- слияние

    fn two_files(dir: &Path) -> (PathBuf, PathBuf) {
        let mk = |name: &str, k: f64, tag: &str| {
            let fx = Fx {
                sheets: vec![
                    sh(
                        "Data",
                        &(row(1, &[cs("A1", tag), cn("B1", k)]) + &row(2, &[cf("A2", "Summary!A1+Data!B1", k), cf("B2", "'Data'!B1*2", k * 2.0)])),
                        "<hyperlinks><hyperlink ref=\"A1\" location=\"Summary!A1\"/></hyperlinks>",
                    ),
                    sh(
                        "Summary",
                        &(row(1, &[cf("A1", "SUM(Data!B1:B2)", k), cs("B1", "Data!A1 text")]) + &row(2, &[cf("A2", "Total*2", k)])),
                        "<conditionalFormatting sqref=\"A1\"><cfRule type=\"expression\" dxfId=\"0\" priority=\"1\"><formula>Data!$B$1&gt;1</formula></cfRule></conditionalFormatting>\
                         <dataValidations count=\"1\"><dataValidation type=\"list\" sqref=\"C1\"><formula1>Data!$A$1:$A$2</formula1></dataValidation></dataValidations>",
                    ),
                ],
                names: "<definedName name=\"Total\">Data!$B$1</definedName>\
                        <definedName name=\"_xlnm.Print_Area\" localSheetId=\"0\">Data!$A$1:$B$2</definedName>\
                        <definedName name=\"_xlnm.Print_Titles\" localSheetId=\"0\">Data!$1:$1</definedName>\
                        <definedName name=\"_xlnm.Print_Area\" localSheetId=\"1\">Summary!$A$1:$B$2</definedName>\
                        <definedName name=\"Loc\" localSheetId=\"1\">Summary!$A$1</definedName>"
                    .into(),
                ..Default::default()
            };
            let p = dir.join(format!("{name}.xlsx"));
            fx.build(&p);
            p
        };
        (mk("a", 1.0, "first"), mk("b", 10.0, "second"))
    }

    #[test]
    fn merge_rewrites_references_names_and_scopes() {
        let dir = tempfile::tempdir().unwrap();
        let (a, b) = two_files(dir.path());
        let out = dir.path().join("o.xlsx");
        let w = merge(&[src(&a, &["Data", "Summary"]), src(&b, &["Data", "Summary"])], &out, &Options::default()).unwrap();
        let wb = part(&out, "xl/workbook.xml");
        assert!(wb.contains("name=\"Data\"") && wb.contains("name=\"b_Data\"") && wb.contains("name=\"b_Summary\""), "{wb}");
        // Листы второго файла ссылаются на свои копии.
        let s3 = part(&out, "xl/worksheets/sheet3.xml");
        assert!(s3.contains("<f>'b_Summary'!A1+'b_Data'!B1</f>"), "{s3}");
        assert!(s3.contains("<f>'b_Data'!B1*2</f>"));
        assert!(s3.contains("location=\"'b_Summary'!A1\""));
        let s4 = part(&out, "xl/worksheets/sheet4.xml");
        assert!(s4.contains("SUM('b_Data'!B1:B2)") && s4.contains("Data!A1 text"), "{s4}");
        assert!(s4.contains("<formula>'b_Data'!$B$1&gt;1</formula>"));
        assert!(s4.contains("<formula1>'b_Data'!$A$1:$A$2</formula1>"));
        assert!(s4.contains("<f>Total_b*2</f>"), "{s4}");
        // Первый файл не тронут.
        let s1 = part(&out, "xl/worksheets/sheet1.xml");
        assert!(s1.contains("<f>Summary!A1+Data!B1</f>"), "{s1}");
        // Имена: привязка к листам результата, глобальное имя второго файла переименовано.
        assert!(wb.contains("<definedName name=\"Total\">Data!$B$1</definedName>"));
        assert!(wb.contains("<definedName name=\"Total_b\">'b_Data'!$B$1</definedName>"), "{wb}");
        assert!(wb.contains("name=\"_xlnm.Print_Area\" localSheetId=\"2\">'b_Data'!$A$1:$B$2<"), "{wb}");
        assert!(wb.contains("name=\"_xlnm.Print_Titles\" localSheetId=\"2\">'b_Data'!$1:$1<"));
        assert!(wb.contains("name=\"_xlnm.Print_Area\" localSheetId=\"3\">'b_Summary'!$A$1:$B$2<"));
        assert!(wb.contains("name=\"Loc\" localSheetId=\"3\">'b_Summary'!$A$1<"));
        assert!(wb.contains("name=\"_xlnm.Print_Area\" localSheetId=\"0\">'Data'!$A$1:$B$2<") || wb.contains("name=\"_xlnm.Print_Area\" localSheetId=\"0\">Data!$A$1:$B$2<"));
        assert!(w.iter().any(|x| x.contains("Total_b")), "{w:?}");
        // Нет повторяющихся пар (имя, область).
        let n = wb.matches("<definedName ").count();
        assert_eq!(n, 10);
    }

    #[test]
    fn removed_sheets_shift_local_scopes() {
        let dir = tempfile::tempdir().unwrap();
        let (a, _) = two_files(dir.path());
        let out = dir.path().join("o.xlsx");
        // Берём только второй лист — его область печати должна указывать на индекс 0.
        merge(&[src(&a, &["Summary"])], &out, &Options::default()).unwrap();
        let wb = part(&out, "xl/workbook.xml");
        assert!(wb.contains("localSheetId=\"0\">Summary!$A$1:$B$2<"), "{wb}");
        assert!(!wb.contains("Data!$A$1:$B$2"), "имя удалённого листа не должно остаться: {wb}");
        assert!(wb.contains("name=\"Loc\" localSheetId=\"0\""));
    }

    #[test]
    fn formulas_to_removed_sheets_become_values() {
        let dir = tempfile::tempdir().unwrap();
        let fx = Fx {
            sheets: vec![
                sh(
                    "A",
                    &(row(1, &[cf("A1", "B!A1+1", 5.0), "<c r=\"B1\" t=\"str\"><f>B!B1&amp;\"x\"</f><v>hello</v></c>".into(), cf("C1", "SUM(A1:A1)", 5.0)])
                        + "<row r=\"2\"><c r=\"A2\"><f t=\"shared\" ref=\"A2:A3\" si=\"0\">B!A1*2</f><v>1</v></c></row><row r=\"3\"><c r=\"A3\"><f t=\"shared\" si=\"0\"/><v>2</v></c></row>"),
                    "",
                ),
                sh("B", &row(1, &[cn("A1", 4.0)]), ""),
            ],
            ..Default::default()
        };
        let p = dir.path().join("a.xlsx");
        fx.build(&p);
        let out = dir.path().join("o.xlsx");
        let w = merge(&[src(&p, &["A"])], &out, &Options::default()).unwrap();
        let s = part(&out, "xl/worksheets/sheet1.xml");
        assert!(!s.contains("B!A1") && !s.contains("<f t=\"shared\""), "{s}");
        assert!(s.contains("<c r=\"A1\"><v>5</v></c>"), "{s}");
        assert!(s.contains("t=\"inlineStr\"") && s.contains("hello"), "{s}");
        assert!(s.contains("<f>SUM(A1:A1)</f>"));
        assert!(w.iter().any(|x| x.contains("заменены вычисленными значениями")), "{w:?}");
    }

    #[test]
    fn pictures_from_different_files_stay_distinct() {
        let dir = tempfile::tempdir().unwrap();
        let drawing = format!(
            "<xdr:wsDr xmlns:xdr=\"http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing\" xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" xmlns:r=\"{R_NS}\">\
             <xdr:oneCellAnchor><xdr:from><xdr:col>1</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>1</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:ext cx=\"100\" cy=\"100\"/>\
             <xdr:pic><xdr:nvPicPr><xdr:cNvPr id=\"2\" name=\"P\"/><xdr:cNvPicPr/></xdr:nvPicPr><xdr:blipFill><a:blip r:embed=\"rId1\"/><a:stretch><a:fillRect/></a:stretch></xdr:blipFill>\
             <xdr:spPr><a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom></xdr:spPr></xdr:pic><xdr:clientData/></xdr:oneCellAnchor></xdr:wsDr>"
        );
        let mk = |name: &str, img: &[u8]| {
            let mut s = sh("S", &row(1, &[cs("A1", name)]), "<drawing r:id=\"rId1\"/>");
            s.rels = vec![("rId1".into(), format!("{REL_BASE}drawing"), "../drawings/drawing1.xml".into())];
            let fx = Fx {
                sheets: vec![s],
                extra: vec![
                    ("xl/drawings/drawing1.xml".into(), drawing.clone().into_bytes(), Some("application/vnd.openxmlformats-officedocument.drawing+xml".into())),
                    (
                        "xl/drawings/_rels/drawing1.xml.rels".into(),
                        format!("<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"rId1\" Type=\"{REL_BASE}image\" Target=\"../media/image1.png\"/></Relationships>").into_bytes(),
                        None,
                    ),
                    ("xl/media/image1.png".into(), img.to_vec(), None),
                ],
                ..Default::default()
            };
            let p = dir.path().join(format!("{name}.xlsx"));
            fx.build(&p);
            p
        };
        let (a, b) = (mk("a", b"PNG-AAAA"), mk("b", b"PNG-BBBB"));
        let out = dir.path().join("o.xlsx");
        merge(&[src(&a, &["S"]), src(&b, &["S"])], &out, &Options::default()).unwrap();
        assert_eq!(part_bytes(&out, "xl/media/image1.png"), b"PNG-AAAA");
        assert_eq!(part_bytes(&out, "xl/media/image2.png"), b"PNG-BBBB");
        assert!(part(&out, "xl/drawings/_rels/drawing2.xml.rels").contains("../media/image2.png"));
        assert!(part(&out, "[Content_Types].xml").contains("Extension=\"png\""));
    }

    fn table_xml(id: u32, name: &str) -> String {
        format!(
            "<table xmlns=\"{WB_NS}\" id=\"{id}\" name=\"{name}\" displayName=\"{name}\" ref=\"A1:B2\" totalsRowShown=\"0\"><autoFilter ref=\"A1:B2\"/>\
             <tableColumns count=\"2\"><tableColumn id=\"1\" name=\"x\"/><tableColumn id=\"2\" name=\"y\"/></tableColumns>\
             <tableStyleInfo name=\"TableStyleMedium2\" showFirstColumn=\"1\" showLastColumn=\"1\" showRowStripes=\"1\" showColumnStripes=\"1\"/></table>"
        )
    }

    #[test]
    fn tables_get_unique_names_ids_and_keep_flags() {
        let dir = tempfile::tempdir().unwrap();
        let mk = |name: &str| {
            let mut s = sh(
                "T",
                &(row(1, &[cs("A1", "x"), cs("B1", "y")]) + &row(2, &[cn("A2", 1.0), cf("B2", "SUM(Table1[x])", 1.0)])),
                "<tableParts count=\"1\"><tablePart r:id=\"rId1\"/></tableParts>",
            );
            s.rels = vec![("rId1".into(), format!("{REL_BASE}table"), "../tables/table1.xml".into())];
            let fx = Fx {
                sheets: vec![s],
                extra: vec![("xl/tables/table1.xml".into(), table_xml(1, "Table1").into_bytes(), Some("application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml".into()))],
                ..Default::default()
            };
            let p = dir.path().join(format!("{name}.xlsx"));
            fx.build(&p);
            p
        };
        let (a, b) = (mk("a"), mk("b"));
        let out = dir.path().join("o.xlsx");
        let w = merge(&[src(&a, &["T"]), src(&b, &["T"])], &out, &Options::default()).unwrap();
        let t1 = part(&out, "xl/tables/table1.xml");
        let t2 = part(&out, "xl/tables/table2.xml");
        assert!(t1.contains("id=\"1\"") && t1.contains("displayName=\"Table1\""), "{t1}");
        assert!(t2.contains("id=\"2\"") && t2.contains("displayName=\"Table1_b\"") && t2.contains("name=\"Table1_b\""), "{t2}");
        for t in [&t1, &t2] {
            assert!(t.contains("showFirstColumn=\"1\"") && t.contains("showLastColumn=\"1\"") && t.contains("showRowStripes=\"1\"") && t.contains("showColumnStripes=\"1\""));
        }
        assert!(part(&out, "xl/worksheets/sheet2.xml").contains("SUM(Table1_b[x])"));
        assert!(part(&out, "xl/worksheets/sheet1.xml").contains("SUM(Table1[x])"));
        assert!(w.iter().any(|x| x.contains("Table1_b")), "{w:?}");
    }

    #[test]
    fn xlsm_macros_are_dropped_with_warning() {
        let dir = tempfile::tempdir().unwrap();
        let mut fx = Fx { sheets: vec![sh("S", &row(1, &[cs("A1", "m")]), "")], ..Default::default() };
        fx.extra.push(("xl/vbaProject.bin".into(), vec![1, 2, 3], Some("application/vnd.ms-office.vbaProject".into())));
        fx.wb_rels.push(("rId9".into(), format!("{REL_BASE}vbaProject"), "vbaProject.bin".into()));
        let p = dir.path().join("m.xlsm");
        fx.build(&p);
        // Содержимое книги объявлено как макро-книга.
        let mut parts = vec![];
        {
            let mut z = zip::ZipArchive::new(std::fs::File::open(&p).unwrap()).unwrap();
            for i in 0..z.len() {
                let mut e = z.by_index(i).unwrap();
                let name = e.name().unwrap().to_string();
                let mut v = vec![];
                std::io::Read::read_to_end(&mut e, &mut v).unwrap();
                if name == "[Content_Types].xml" {
                    v = String::from_utf8(v).unwrap().replace("spreadsheetml.sheet.main+xml", "ms-excel.sheet.macroEnabled.main+xml").into_bytes();
                }
                parts.push((name, v));
            }
        }
        write_zip(&p, &parts);
        let out = dir.path().join("o.xlsx");
        let w = merge(&[src(&p, &["S"])], &out, &Options::default()).unwrap();
        assert!(!has_part(&out, "xl/vbaProject.bin"));
        let ct = part(&out, "[Content_Types].xml");
        assert!(ct.contains("spreadsheetml.sheet.main+xml") && !ct.contains("macroEnabled") && !ct.contains("vbaProject"), "{ct}");
        assert!(w.iter().any(|x| x.contains("макросы")), "{w:?}");
    }

    #[test]
    fn array_formulas_iconsets_outline_and_sparklines_survive() {
        let dir = tempfile::tempdir().unwrap();
        let tail = "<conditionalFormatting sqref=\"A1:A3\"><cfRule type=\"iconSet\" priority=\"1\"><iconSet iconSet=\"3Arrows\"><cfvo type=\"percent\" val=\"0\"/><cfvo type=\"percent\" val=\"33\"/><cfvo type=\"percent\" val=\"67\"/></iconSet></cfRule></conditionalFormatting>\
            <extLst><ext uri=\"{05C60535-1F16-4fd2-B633-F4F36F0B64E0}\" xmlns:x14=\"http://schemas.microsoft.com/office/spreadsheetml/2009/9/main\"><x14:sparklineGroups xmlns:xm=\"http://schemas.microsoft.com/office/excel/2006/main\"><x14:sparklineGroup><x14:sparklines><x14:sparkline><xm:f>Data!A1:A3</xm:f><xm:sqref>D1</xm:sqref></x14:sparkline></x14:sparklines></x14:sparklineGroup></x14:sparklineGroups></ext>\
            <ext uri=\"{A8765BA9-456A-4dab-B4F3-ACF838C121DE}\" xmlns:x14=\"http://schemas.microsoft.com/office/spreadsheetml/2009/9/main\"><x14:slicerList/></ext></extLst>";
        let rows = "<row r=\"1\"><c r=\"A1\"><f t=\"array\" ref=\"A1:A3\">ROW(1:3)*2</f><v>2</v></c></row><row r=\"2\" outlineLevel=\"1\" hidden=\"1\"><c r=\"A2\"><v>4</v></c></row><row r=\"3\"><c r=\"A3\"><v>6</v></c></row>";
        let mk = |name: &str, sheet: &str| {
            let mut s = sh(sheet, rows, tail);
            s.xml = s.xml.replace("<sheetData>", "<sheetFormatPr defaultRowHeight=\"15\" outlineLevelRow=\"1\"/><cols><col min=\"1\" max=\"1\" width=\"12\" outlineLevel=\"1\"/></cols><sheetData>");
            let p = dir.path().join(format!("{name}.xlsx"));
            Fx { sheets: vec![s], ..Default::default() }.build(&p);
            p
        };
        let (a, b) = (mk("a", "Data"), mk("b", "Data"));
        let out = dir.path().join("o.xlsx");
        let w = merge(&[src(&a, &["Data"]), src(&b, &["Data"])], &out, &Options::default()).unwrap();
        for n in [1, 2] {
            let s = part(&out, &format!("xl/worksheets/sheet{n}.xml"));
            assert!(s.contains("<f t=\"array\" ref=\"A1:A3\">ROW(1:3)*2</f>"), "{s}");
            assert!(s.contains("<iconSet iconSet=\"3Arrows\">"), "{s}");
            assert!(s.contains("outlineLevel=\"1\" hidden=\"1\"") && s.contains("outlineLevelRow=\"1\"") && s.contains("<col min=\"1\" max=\"1\" width=\"12\" outlineLevel=\"1\"/>"), "{s}");
            assert!(s.contains("sparklineGroups") && !s.contains("slicerList"), "{s}");
        }
        assert!(part(&out, "xl/worksheets/sheet2.xml").contains("<xm:f>'b_Data'!A1:A3</xm:f>"));
        assert!(part(&out, "xl/worksheets/sheet1.xml").contains("<xm:f>Data!A1:A3</xm:f>"));
        assert!(w.iter().any(|x| x.contains("не перенесено")), "{w:?}");
    }

    #[test]
    fn date_system_1904_is_converted() {
        let dir = tempfile::tempdir().unwrap();
        let mk = |name: &str, d1904: bool| {
            let fx = Fx {
                sheets: vec![sh(name, "<row r=\"1\"><c r=\"A1\" s=\"1\"><v>100</v></c><c r=\"B1\"><v>100</v></c></row>", "")],
                wb_pr: if d1904 { "date1904=\"1\"".into() } else { String::new() },
                ..Default::default()
            };
            let p = dir.path().join(format!("{name}.xlsx"));
            fx.build(&p);
            p
        };
        let (a, b) = (mk("a", false), mk("b", true));
        let out = dir.path().join("o.xlsx");
        let w = merge(&[src(&a, &["a"]), src(&b, &["b"])], &out, &Options::default()).unwrap();
        assert!(!part(&out, "xl/workbook.xml").contains("date1904"));
        let s2 = part(&out, "xl/worksheets/sheet2.xml");
        assert!(s2.contains("<v>1562</v>"), "дата +1462: {s2}");
        assert!(s2.contains("<c r=\"B1\"><v>100</v></c>"), "обычные числа не трогаем: {s2}");
        assert!(part(&out, "xl/worksheets/sheet1.xml").contains("<v>100</v>"));
        assert!(w.iter().any(|x| x.contains("1904")), "{w:?}");
        // Обратный порядок: основа в системе 1904.
        let out2 = dir.path().join("o2.xlsx");
        merge(&[src(&b, &["b"]), src(&a, &["a"])], &out2, &Options::default()).unwrap();
        assert!(part(&out2, "xl/workbook.xml").contains("date1904=\"1\""));
        assert!(part(&out2, "xl/worksheets/sheet2.xml").contains("<v>-1362</v>") || part(&out2, "xl/worksheets/sheet2.xml").contains("<v>100</v>"));
    }

    #[test]
    fn hidden_sheets_and_active_tab() {
        let dir = tempfile::tempdir().unwrap();
        let mut h = sh("H", &row(1, &[cs("A1", "h")]), "");
        h.state = "hidden";
        let mut vh = sh("VH", &row(1, &[cs("A1", "vh")]), "");
        vh.state = "veryHidden";
        let vis = sh("V", &row(1, &[cs("A1", "v")]), "<pageMargins left=\"1\" right=\"1\" top=\"1\" bottom=\"1\" header=\"0\" footer=\"0\"/>");
        let p = dir.path().join("a.xlsx");
        Fx { sheets: vec![h, vh, vis], ..Default::default() }.build(&p);
        let info = list_sheets(&p).unwrap();
        assert_eq!(info.iter().map(|s| (s.name.as_str(), s.hidden)).collect::<Vec<_>>(), [("H", true), ("VH", true), ("V", false)]);
        let out = dir.path().join("o.xlsx");
        merge(&[src(&p, &["H", "VH", "V"])], &out, &Options::default()).unwrap();
        let wb = part(&out, "xl/workbook.xml");
        assert!(wb.contains("activeTab=\"2\"") && wb.contains("state=\"hidden\"") && wb.contains("state=\"veryHidden\""), "{wb}");
        assert!(part(&out, "xl/worksheets/sheet3.xml").contains("tabSelected=\"1\""));
        assert!(!part(&out, "xl/worksheets/sheet1.xml").contains("tabSelected"));
        // Только скрытые листы: первый становится видимым.
        let out2 = dir.path().join("o2.xlsx");
        let w = merge(&[src(&p, &["H", "VH"])], &out2, &Options::default()).unwrap();
        let wb2 = part(&out2, "xl/workbook.xml");
        assert!(wb2.contains("activeTab=\"0\"") && !wb2.contains("name=\"H\" sheetId=\"1\" state"), "{wb2}");
        assert!(w.iter().any(|x| x.contains("скрыты")), "{w:?}");
        // Для PDF скрытые листы показываются.
        let out3 = dir.path().join("o3.xlsx");
        merge(&[src(&p, &["H", "VH", "V"])], &out3, &Options { fit_to_width: true }).unwrap();
        assert!(!part(&out3, "xl/workbook.xml").contains("state="));
    }

    #[test]
    fn fit_to_width_sets_page_setup() {
        let dir = tempfile::tempdir().unwrap();
        let wide: String = (1..=30u8).map(|i| cn(&format!("{}1", (b'A' + (i - 1) % 26) as char), i as f64)).collect::<Vec<_>>().join("");
        let narrow = sh("N", &row(1, &[cn("A1", 1.0)]), "<pageMargins left=\"1\" right=\"1\" top=\"1\" bottom=\"1\" header=\"0\" footer=\"0\"/><pageSetup scale=\"50\" orientation=\"portrait\" r:id=\"rId1\"/><colBreaks count=\"1\" manualBreakCount=\"1\"><brk id=\"3\" max=\"1048575\" man=\"1\"/></colBreaks>");
        let mut wide_sh = sh("W", &format!("<row r=\"1\">{}</row>", wide.replace("AA1", "A1")), "");
        wide_sh.xml = wide_sh.xml.replace("<sheetData>", "<cols><col min=\"1\" max=\"26\" width=\"14\" customWidth=\"1\"/></cols><sheetData>");
        let p = dir.path().join("a.xlsx");
        Fx { sheets: vec![narrow, wide_sh], ..Default::default() }.build(&p);
        let out = dir.path().join("o.xlsx");
        merge(&[src(&p, &["N", "W"])], &out, &Options { fit_to_width: true }).unwrap();
        let n = part(&out, "xl/worksheets/sheet1.xml");
        assert!(n.contains("<sheetPr><pageSetUpPr fitToPage=\"1\"/></sheetPr>"), "{n}");
        assert!(n.contains("fitToWidth=\"1\"") && n.contains("fitToHeight=\"0\"") && !n.contains("scale="), "{n}");
        assert!(!n.contains("orientation=\"landscape\"") && !n.contains("colBreaks"), "{n}");
        let w = part(&out, "xl/worksheets/sheet2.xml");
        assert!(w.contains("<pageSetup orientation=\"landscape\"") || w.contains("orientation=\"landscape\""), "{w}");
        assert!(w.find("</sheetData>").unwrap() < w.find("<pageSetup").unwrap());
        // Без параметра страница не меняется.
        let out2 = dir.path().join("o2.xlsx");
        merge(&[src(&p, &["N"])], &out2, &Options::default()).unwrap();
        assert!(part(&out2, "xl/worksheets/sheet1.xml").contains("scale=\"50\""));
    }

    #[test]
    fn theme_colors_are_resolved_when_themes_differ() {
        let dir = tempfile::tempdir().unwrap();
        let mk = |name: &str, accent: &str| {
            let fx = Fx {
                sheets: vec![sh(name, "<row r=\"1\"><c r=\"A1\" s=\"2\" t=\"inlineStr\"><is><t>c</t></is></c></row>", "")],
                theme: Some(theme(accent)),
                ..Default::default()
            };
            let p = dir.path().join(format!("{name}.xlsx"));
            fx.build(&p);
            p
        };
        let (a, b) = (mk("a", "4F81BD"), mk("b", "00FF00"));
        let out = dir.path().join("o.xlsx");
        merge(&[src(&a, &["a"]), src(&b, &["b"])], &out, &Options::default()).unwrap();
        let st = part(&out, "xl/styles.xml");
        assert!(st.contains("<color theme=\"4\"/>"), "первая книга сохраняет тему: {st}");
        assert!(st.contains("<color rgb=\"FF00FF00\"/>"), "цвет второй книги должен стать RGB: {st}");
        assert!(part(&out, "xl/theme/theme1.xml").contains("4F81BD"));
    }

    #[test]
    fn prefixed_namespaces_are_accepted() {
        let dir = tempfile::tempdir().unwrap();
        let p = dir.path().join("x.xlsx");
        let wb = format!(
            "<x:workbook xmlns:x=\"{WB_NS}\" xmlns:r=\"{R_NS}\"><x:sheets><x:sheet name=\"S1\" sheetId=\"1\" r:id=\"rId1\"/></x:sheets></x:workbook>"
        );
        let sheet = format!("<x:worksheet xmlns:x=\"{WB_NS}\"><x:sheetData><x:row r=\"1\"><x:c r=\"A1\" t=\"inlineStr\"><x:is><x:t>hi</x:t></x:is></x:c></x:row></x:sheetData></x:worksheet>");
        let parts: Vec<(String, Vec<u8>)> = vec![
            ("[Content_Types].xml".into(), format!("<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\"><Default Extension=\"rels\" ContentType=\"{CT_RELS}\"/><Default Extension=\"xml\" ContentType=\"application/xml\"/><Override PartName=\"/xl/workbook.xml\" ContentType=\"{CT_WORKBOOK}\"/><Override PartName=\"/xl/worksheets/sheet1.xml\" ContentType=\"{CT_WORKSHEET}\"/></Types>").into_bytes()),
            ("_rels/.rels".into(), format!("<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"rId1\" Type=\"{REL_OFFICE_DOC}\" Target=\"xl/workbook.xml\"/></Relationships>").into_bytes()),
            ("xl/_rels/workbook.xml.rels".into(), format!("<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"rId1\" Type=\"{REL_BASE}worksheet\" Target=\"worksheets/sheet1.xml\"/></Relationships>").into_bytes()),
            ("xl/workbook.xml".into(), wb.into_bytes()),
            ("xl/worksheets/sheet1.xml".into(), sheet.into_bytes()),
        ];
        write_zip(&p, &parts);
        assert_eq!(list_sheets(&p).unwrap()[0].name, "S1");
        let out = dir.path().join("o.xlsx");
        merge(&[src(&p, &["S1"])], &out, &Options::default()).unwrap();
        let s = part(&out, "xl/worksheets/sheet1.xml");
        assert!(s.contains("<c r=\"A1\" t=\"inlineStr\"><is><t>hi</t></is></c>") && !s.contains("x:"), "{s}");
    }

    #[test]
    fn chart_sheets_are_excluded_and_reported() {
        let dir = tempfile::tempdir().unwrap();
        let cs_xml = format!("<chartsheet xmlns=\"{WB_NS}\"><sheetViews><sheetView workbookViewId=\"0\"/></sheetViews></chartsheet>");
        let p = dir.path().join("c.xlsx");
        Fx { sheets: vec![sh("Data", &row(1, &[cs("A1", "d")]), ""), Sh { name: "Chart1".into(), state: "visible", xml: cs_xml, rels: vec![] }], ..Default::default() }.build(&p);
        assert_eq!(list_sheets(&p).unwrap().iter().map(|s| s.name.as_str()).collect::<Vec<_>>(), ["Data"]);
        let out = dir.path().join("o.xlsx");
        let w = merge(&[src(&p, &["Data"])], &out, &Options::default()).unwrap();
        assert!(w.iter().any(|x| x.contains("лист диаграммы «Chart1»")), "{w:?}");
        let e = merge(&[src(&p, &["Chart1"])], &out, &Options::default()).unwrap_err();
        assert!(e.to_string().contains("лист диаграммы"), "{e}");
    }

    #[test]
    fn shared_strings_and_styles_are_remapped() {
        let dir = tempfile::tempdir().unwrap();
        let mk = |name: &str, text: &str| {
            let mut fx = Fx { sheets: vec![sh(name, "<row r=\"1\"><c r=\"A1\" t=\"s\" s=\"2\"><v>1</v></c><c r=\"B1\" t=\"s\"><v>0</v></c></row>", "")], ..Default::default() };
            fx.extra.push((
                "xl/sharedStrings.xml".into(),
                format!("<sst xmlns=\"{WB_NS}\" count=\"2\" uniqueCount=\"2\"><si><t>zero</t></si><si><t>{text}</t></si></sst>").into_bytes(),
                Some(CT_SST.into()),
            ));
            fx.wb_rels.push(("rId9".into(), format!("{REL_BASE}sharedStrings"), "sharedStrings.xml".into()));
            let p = dir.path().join(format!("{name}.xlsx"));
            fx.build(&p);
            p
        };
        let (a, b) = (mk("a", "one"), mk("b", "uno"));
        let out = dir.path().join("o.xlsx");
        merge(&[src(&a, &["a"]), src(&b, &["b"])], &out, &Options::default()).unwrap();
        let sst = part(&out, "xl/sharedStrings.xml");
        assert!(sst.contains("uniqueCount=\"3\"") && sst.contains("one") && sst.contains("uno"), "{sst}");
        let s2 = part(&out, "xl/worksheets/sheet2.xml");
        assert!(s2.contains("<c r=\"A1\" t=\"s\" s=\"2\"><v>2</v></c>"), "{s2}");
        assert!(s2.contains("<c r=\"B1\" t=\"s\"><v>1</v></c>"), "{s2}");
    }

    // ---------------------------------------------------------------- ошибки и список листов

    #[test]
    fn friendly_errors() {
        let dir = tempfile::tempdir().unwrap();
        let out = dir.path().join("o.xlsx");
        let check = |p: &Path, needle: &str| {
            let e = list_sheets(p).unwrap_err().to_string();
            assert!(e.contains(needle) && e.contains(&file_label(p)), "{e}");
            assert!(!e.contains("Error") && !e.contains("{"), "{e}");
            let e = merge(&[src(p, &["S"])], &out, &Options::default()).unwrap_err().to_string();
            assert!(e.contains(needle) && e.contains(&file_label(p)), "{e}");
        };
        let enc = dir.path().join("enc.xlsx");
        let mut v = vec![0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1];
        v.extend(vec![0u8; 600]);
        v.extend("EncryptedPackage".encode_utf16().flat_map(|u| u.to_le_bytes()));
        std::fs::write(&enc, v).unwrap();
        check(&enc, "паролем");
        let txt = dir.path().join("t.xlsx");
        std::fs::write(&txt, "a,b,c\n1,2,3\n").unwrap();
        check(&txt, "не книга Excel");
        let bad = dir.path().join("bad.xlsx");
        std::fs::write(&bad, b"PK\x03\x04 not really a zip").unwrap();
        check(&bad, "повреждён");
        let empty = dir.path().join("empty.xlsx");
        std::fs::write(&empty, b"").unwrap();
        check(&empty, "пустой");
        let missing = dir.path().join("none.xlsx");
        check(&missing, "не найден");
        let docx = dir.path().join("d.xlsx");
        write_zip(&docx, &[("[Content_Types].xml".into(), b"<Types/>".to_vec()), ("word/document.xml".into(), b"<w/>".to_vec())]);
        check(&docx, "не книга Excel");
        assert!(!out.exists(), "при ошибке результат не создаётся");
    }

    #[test]
    fn errors_name_missing_sheet_and_leave_no_output() {
        let dir = tempfile::tempdir().unwrap();
        let p = dir.path().join("a.xlsx");
        Fx { sheets: vec![sh("S", &row(1, &[cs("A1", "x")]), "")], ..Default::default() }.build(&p);
        let out = dir.path().join("o.xlsx");
        let e = merge(&[src(&p, &["Nope"])], &out, &Options::default()).unwrap_err().to_string();
        assert!(e.contains("Nope") && e.contains("a.xlsx"), "{e}");
        assert!(merge(&[], &out, &Options::default()).is_err());
        assert!(merge(&[src(&p, &[])], &out, &Options::default()).is_err());
        // Повреждённый лист: ошибка, а не пустой лист, и старый результат не затирается.
        std::fs::write(&out, b"old").unwrap();
        let broken = dir.path().join("b.xlsx");
        let mut fx = Fx { sheets: vec![sh("S", &row(1, &[cs("A1", "x")]), "")], ..Default::default() };
        fx.sheets[0].xml = fx.sheets[0].xml.replace("</row>", "</rowx>");
        fx.build(&broken);
        let e = merge(&[src(&broken, &["S"])], &out, &Options::default()).unwrap_err().to_string();
        assert!(e.contains("b.xlsx") && e.contains("повреждён"), "{e}");
        assert_eq!(std::fs::read(&out).unwrap(), b"old");
        let leftovers: Vec<_> = std::fs::read_dir(dir.path()).unwrap().filter_map(|e| e.ok()).filter(|e| e.file_name().to_string_lossy().starts_with(".merge-")).collect();
        assert!(leftovers.is_empty(), "временные файлы не должны оставаться");
    }

    #[test]
    fn list_sheets_agrees_with_merge() {
        let dir = tempfile::tempdir().unwrap();
        let (a, _) = two_files(dir.path());
        let names: Vec<String> = list_sheets(&a).unwrap().into_iter().map(|s| s.name).collect();
        assert_eq!(names, ["Data", "Summary"]);
        let out = dir.path().join("o.xlsx");
        merge(&[Source { path: a, sheets: names }], &out, &Options::default()).unwrap();
        let wb = part(&out, "xl/workbook.xml");
        assert!(wb.contains("name=\"Data\"") && wb.contains("name=\"Summary\""));
    }

    #[test]
    fn long_and_invalid_sheet_names_are_fixed_with_warning() {
        let dir = tempfile::tempdir().unwrap();
        let long = "Очень длинное название листа из тридцати пяти символов";
        let p = dir.path().join("a.xlsx");
        Fx { sheets: vec![sh(long, &row(1, &[cf("A1", &format!("'{long}'!B1"), 1.0)]), "")], ..Default::default() }.build(&p);
        let out = dir.path().join("o.xlsx");
        let w = merge(&[src(&p, &[long])], &out, &Options::default()).unwrap();
        let wb = part(&out, "xl/workbook.xml");
        let short: String = long.chars().take(31).collect();
        assert!(wb.contains(&format!("name=\"{short}\"")), "{wb}");
        assert!(part(&out, "xl/worksheets/sheet1.xml").contains(&format!("<f>'{short}'!B1</f>")));
        assert!(w.iter().any(|x| x.contains("недопустимо")), "{w:?}");
    }

    #[test]
    fn single_file_roundtrip_is_stable() {
        let dir = tempfile::tempdir().unwrap();
        let (a, _) = two_files(dir.path());
        let out = dir.path().join("o.xlsx");
        merge(&[src(&a, &["Data", "Summary"])], &out, &Options::default()).unwrap();
        let out2 = dir.path().join("o2.xlsx");
        merge(&[src(&out, &["Data", "Summary"])], &out2, &Options::default()).unwrap();
        assert_eq!(part(&out, "xl/worksheets/sheet1.xml"), part(&out2, "xl/worksheets/sheet1.xml"));
        assert_eq!(part(&out, "xl/styles.xml"), part(&out2, "xl/styles.xml"));
    }

    #[test]
    fn xml_layer_roundtrip_and_namespaces() {
        let src = "<?xml version=\"1.0\" encoding=\"UTF-8\"?><r:a xmlns:r=\"urn:r\" xmlns=\"urn:d\" k=\"v&amp;&quot;\"><b x=\"1\"/><r:c>t &lt; &#65;<![CDATA[<x>]]></r:c><!-- c --></r:a>";
        let e = xml::parse(src.as_bytes()).unwrap();
        assert_eq!((e.ns.as_str(), e.local.as_str()), ("urn:r", "a"));
        let b = e.child("b").unwrap();
        assert_eq!(b.ns, "urn:d");
        assert_eq!(e.child("c").unwrap().text(), "t < A<x>");
        let out = String::from_utf8(xml::doc(&e, "")).unwrap();
        assert!(out.contains("<r:a xmlns=\"urn:d\" xmlns:r=\"urn:r\" k=\"v&amp;&quot;\">") || out.contains("<r:a"), "{out}");
        let e2 = xml::parse(out.as_bytes()).unwrap();
        assert_eq!(e2.child("c").unwrap().text(), "t < A<x>");
        assert_eq!(e2.child("b").unwrap().ns, "urn:d");
        // Поддерево переносится с недостающими объявлениями префиксов.
        let sub = e.child("c").unwrap().clone();
        let s = xml::to_string(&sub, "");
        assert!(s.contains("xmlns:r=\"urn:r\"") || s.starts_with("<r:c"), "{s}");
        // Кодировки.
        assert_eq!(xml::decode(b"<?xml version=\"1.0\" encoding=\"windows-1251\"?><a>\xcf\xf0\xe8\xe2\xe5\xf2</a>").unwrap().to_string(), "<?xml version=\"1.0\" encoding=\"windows-1251\"?><a>Привет</a>");
        assert!(xml::parse(b"<a><b></a>").is_err());
    }

    #[test]
    fn external_links_are_renumbered() {
        let dir = tempfile::tempdir().unwrap();
        let mk = |name: &str, target: &str| {
            let mut fx = Fx { sheets: vec![sh(name, &row(1, &[cf("A1", "[1]Ext!A1+1", 3.0)]), "")], ..Default::default() };
            fx.wb_tail = "<externalReferences><externalReference r:id=\"rId9\"/></externalReferences>".into();
            fx.wb_rels.push(("rId9".into(), format!("{REL_BASE}externalLink"), "externalLinks/externalLink1.xml".into()));
            fx.extra.push((
                "xl/externalLinks/externalLink1.xml".into(),
                format!("<externalLink xmlns=\"{WB_NS}\"><externalBook xmlns:r=\"{R_NS}\" r:id=\"rId1\"><sheetNames><sheetName val=\"Ext\"/></sheetNames></externalBook></externalLink>").into_bytes(),
                Some("application/vnd.openxmlformats-officedocument.spreadsheetml.externalLink+xml".into()),
            ));
            fx.extra.push((
                "xl/externalLinks/_rels/externalLink1.xml.rels".into(),
                format!("<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"rId1\" Type=\"{REL_BASE}externalLinkPath\" Target=\"{target}\" TargetMode=\"External\"/></Relationships>").into_bytes(),
                None,
            ));
            let p = dir.path().join(format!("{name}.xlsx"));
            fx.build(&p);
            p
        };
        let (a, b) = (mk("a", "one.xlsx"), mk("b", "two.xlsx"));
        let out = dir.path().join("o.xlsx");
        let w = merge(&[src(&a, &["a"]), src(&b, &["b"])], &out, &Options::default()).unwrap();
        assert!(part(&out, "xl/worksheets/sheet1.xml").contains("<f>[1]Ext!A1+1</f>"));
        assert!(part(&out, "xl/worksheets/sheet2.xml").contains("<f>[2]Ext!A1+1</f>"));
        assert_eq!(part(&out, "xl/workbook.xml").matches("<externalReference ").count(), 2);
        assert!(part(&out, "xl/externalLinks/_rels/externalLink2.xml.rels").contains("two.xlsx"));
        assert!(w.iter().any(|x| x.contains("ссылки на другие книги")), "{w:?}");
    }

    #[test]
    fn named_styles_are_not_duplicated() {
        let dir = tempfile::tempdir().unwrap();
        let st = |name: &str, builtin: &str| {
            format!(
                "<styleSheet xmlns=\"{WB_NS}\"><fonts count=\"1\"><font><sz val=\"11\"/><name val=\"Calibri\"/></font></fonts><fills count=\"2\"><fill><patternFill patternType=\"none\"/></fill><fill><patternFill patternType=\"gray125\"/></fill></fills><borders count=\"1\"><border><left/><right/><top/><bottom/><diagonal/></border></borders>\
                 <cellStyleXfs count=\"2\"><xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\"/><xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\"/></cellStyleXfs>\
                 <cellXfs count=\"1\"><xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\"/></cellXfs>\
                 <cellStyles count=\"2\"><cellStyle name=\"{name}\" xfId=\"0\" builtinId=\"0\"/><cellStyle name=\"Good\" xfId=\"1\" {builtin}/></cellStyles></styleSheet>"
            )
        };
        let mk = |name: &str, normal: &str, builtin: &str| {
            let p = dir.path().join(format!("{name}.xlsx"));
            Fx { sheets: vec![sh(name, &row(1, &[cn("A1", 1.0)]), "")], styles: Some(st(normal, builtin)), ..Default::default() }.build(&p);
            p
        };
        let (a, b) = (mk("a", "Normal", ""), mk("b", "Обычный", "builtinId=\"26\""));
        let out = dir.path().join("o.xlsx");
        merge(&[src(&a, &["a"]), src(&b, &["b"])], &out, &Options::default()).unwrap();
        let styles = part(&out, "xl/styles.xml");
        assert_eq!(styles.matches("<cellStyle ").count(), 2, "{styles}");
        assert!(!styles.contains("Обычный"), "{styles}");
    }
}
