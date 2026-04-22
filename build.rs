fn main() {
    // Load .env from project root for local dev. In CI the vars are injected
    // as repository secrets and don't need a file. Either way they end up
    // baked into the binary via env!() — users never need to supply them.
    if let Ok(contents) = std::fs::read_to_string(".env") {
        for line in contents.lines() {
            let line = line.trim();
            if line.is_empty() || line.starts_with('#') { continue; }
            if let Some((key, val)) = line.split_once('=') {
                println!("cargo:rustc-env={}={}", key.trim(), val.trim());
            }
        }
    }
    println!("cargo:rerun-if-changed=.env");

    slint_build::compile("ui/app.slint").expect("Slint build failed");

    #[cfg(target_os = "windows")]
    embed_exe_icon();
}

/// Generate a multi-resolution `.ico` from `assets/icon_120.png` and bake it
/// into the exe as the primary icon resource. Without this, Windows falls
/// back to the generic exe glyph in the Start menu, taskbar, and file
/// shortcuts — the runtime tray icon doesn't help the shell.
#[cfg(target_os = "windows")]
fn embed_exe_icon() {
    use image::imageops::FilterType;
    use ico::{IconDir, IconDirEntry, IconImage, ResourceType};

    const SOURCE: &str = "assets/icon_120.png";
    println!("cargo:rerun-if-changed={SOURCE}");

    let out_dir  = std::env::var("OUT_DIR").expect("OUT_DIR not set");
    let ico_path = std::path::Path::new(&out_dir).join("ruscal.ico");

    let src = image::open(SOURCE).expect("failed to read icon PNG source");

    // Shell-relevant sizes: 16 list-view, 32 small tile, 48 large list, 256
    // Explorer jumbo. Intermediate sizes keep the rescale crisp at HiDPI.
    let mut dir = IconDir::new(ResourceType::Icon);
    for size in [16u32, 24, 32, 48, 64, 128, 256] {
        let resized = src.resize_exact(size, size, FilterType::Lanczos3).to_rgba8();
        let img     = IconImage::from_rgba_data(size, size, resized.into_raw());
        dir.add_entry(IconDirEntry::encode(&img).expect("encode .ico entry"));
    }
    let file = std::fs::File::create(&ico_path).expect("create .ico");
    dir.write(file).expect("write .ico");

    let mut res = winresource::WindowsResource::new();
    res.set_icon(ico_path.to_str().expect("ico path non-UTF8"));

    // winresource's Windows SDK auto-detect only checks the standard
    // `Program Files` locations. For portable MSVC/SDK installs (e.g.
    // `%USERPROFILE%\Apps\msvc\...`) that lookup fails. If `rc.exe` is on
    // PATH we already know where the toolkit lives — hand the parent
    // directory to winresource so it builds the `.res` with an absolute
    // invocation.
    if let Some(rc_dir) = find_rc_dir() {
        res.set_toolkit_path(&rc_dir);
    }

    res.compile().expect("winresource compile failed");
}

#[cfg(target_os = "windows")]
fn find_rc_dir() -> Option<String> {
    let output = std::process::Command::new("where").arg("rc.exe").output().ok()?;
    if !output.status.success() { return None; }
    let stdout = String::from_utf8(output.stdout).ok()?;
    let first  = stdout.lines().next()?.trim();
    if first.is_empty() { return None; }
    std::path::Path::new(first)
        .parent()
        .and_then(|p| p.to_str())
        .map(str::to_owned)
}
