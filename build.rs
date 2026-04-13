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
}
