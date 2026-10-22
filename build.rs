//! Встраивает в Windows-exe значок и сведения о версии.

fn main() {
    println!("cargo:rerun-if-changed=assets/icon.ico");
    if std::env::var("CARGO_CFG_TARGET_OS").as_deref() == Ok("windows") {
        let mut res = winresource::WindowsResource::new();
        res.set_icon("assets/icon.ico");
        res.set("ProductName", "Merge Files");
        res.set("FileDescription", "Объединение файлов Excel и Word");
        res.set("LegalCopyright", "MIT License");
        if let Err(e) = res.compile() {
            // Без ресурсов программа работает, только без значка — не ломаем сборку.
            println!("cargo:warning=не удалось встроить ресурсы Windows: {e}");
        }
    }
}
