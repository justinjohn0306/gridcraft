//! Windows only: embed the app icon and version info (VERSIONINFO) into `sheetcraft.exe`, so it
//! shows in Explorer, the taskbar, the Start menu and Alt-Tab.
//!
//! On every other target this does nothing. A missing resource compiler is a warning, so a
//! cross-compile from macOS or Linux still links, unless `SHEETCRAFT_REQUIRE_WINRES=1` turns it
//! into an error (for release builds).

fn main() {
    println!("cargo:rerun-if-changed=build.rs");
    println!("cargo:rerun-if-changed=../../assets/app-icon/sheetcraft.ico");
    println!("cargo:rerun-if-env-changed=SHEETCRAFT_REQUIRE_WINRES");
    if std::env::var("CARGO_CFG_TARGET_OS").as_deref() != Ok("windows") {
        return;
    }
    let mut res = winresource::WindowsResource::new();
    res.set_icon("../../assets/app-icon/sheetcraft.ico")
        .set("ProductName", "SheetCraft")
        .set("FileDescription", "SheetCraft spreadsheet")
        .set("CompanyName", "Learning Machines LLC")
        .set("LegalCopyright", "Copyright (c) the SheetCraft authors. MIT OR Apache-2.0.")
        .set("OriginalFilename", "sheetcraft.exe")
        .set("InternalName", "sheetcraft");
    if let Err(e) = res.compile() {
        if std::env::var_os("SHEETCRAFT_REQUIRE_WINRES").is_some() {
            eprintln!("embedding Windows resources failed: {e}");
            std::process::exit(1);
        }
        println!("cargo:warning=sheetcraft.exe built without icon/version resources: {e}");
    }
}
