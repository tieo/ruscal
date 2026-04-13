/// OAuth 2.0 PKCE flow for installed desktop applications.
use std::io::{BufRead, BufReader, Write as _};
use std::net::{TcpListener, TcpStream};
use std::path::PathBuf;
use std::time::{Duration, SystemTime, UNIX_EPOCH};

use rand::Rng;
use sha2::{Digest, Sha256};

use super::GoogleError;

// ── Public types ──────────────────────────────────────────────────────────────

pub struct GoogleCreds {
    pub client_id:     String,
    pub client_secret: String,
}

#[derive(serde::Serialize, serde::Deserialize, Clone)]
pub struct Tokens {
    pub access_token:  String,
    pub refresh_token: String,
    pub expires_at:    i64,
    #[serde(default)]
    pub email:         String,
}

impl Tokens {
    pub fn is_expired(&self) -> bool {
        let now = SystemTime::now()
            .duration_since(UNIX_EPOCH)
            .unwrap_or_default()
            .as_secs() as i64;
        now >= self.expires_at - 60
    }
}

// ── Token persistence (per-account) ──────────────────────────────────────────

fn tokens_dir() -> Option<PathBuf> {
    dirs::data_local_dir().map(|d| d.join("ruscal").join("tokens"))
}

fn token_path_for(email: &str) -> Option<PathBuf> {
    let safe = email.replace('@', "_at_").replace('.', "_dot_");
    tokens_dir().map(|d| d.join(format!("{safe}.json")))
}

/// Legacy single-account path — kept for migration only.
pub fn legacy_token_path() -> Option<PathBuf> {
    dirs::data_local_dir().map(|d| d.join("ruscal").join("google_token.json"))
}

pub fn load_tokens_for(email: &str) -> Option<Tokens> {
    let data = std::fs::read_to_string(token_path_for(email)?).ok()?;
    serde_json::from_str(&data).ok()
}

pub fn save_tokens_for(email: &str, tokens: &Tokens) -> Result<(), GoogleError> {
    let path = token_path_for(email)
        .ok_or_else(|| GoogleError::Config("cannot locate config directory".into()))?;
    if let Some(parent) = path.parent() {
        std::fs::create_dir_all(parent)?;
    }
    // Embed the email in the token file so we can recover it without filename parsing.
    let mut t = tokens.clone();
    t.email = email.to_owned();
    std::fs::write(path, serde_json::to_string_pretty(&t)?)?;
    Ok(())
}

pub fn revoke_tokens_for(email: &str) {
    if let Some(path) = token_path_for(email) {
        let _ = std::fs::remove_file(path);
    }
}

pub fn load_legacy_tokens() -> Option<Tokens> {
    let data = std::fs::read_to_string(legacy_token_path()?).ok()?;
    serde_json::from_str(&data).ok()
}

pub fn delete_legacy_tokens() {
    if let Some(path) = legacy_token_path() {
        let _ = std::fs::remove_file(path);
    }
}

/// Returns all emails that have stored token files, sorted alphabetically.
pub fn list_stored_accounts() -> Vec<String> {
    let Some(dir) = tokens_dir() else { return vec![] };
    let Ok(entries) = std::fs::read_dir(&dir) else { return vec![] };

    let mut emails = Vec::new();
    for entry in entries.flatten() {
        let path = entry.path();
        if path.extension().and_then(|e| e.to_str()) != Some("json") {
            continue;
        }
        if let Ok(data) = std::fs::read_to_string(&path) {
            if let Ok(tokens) = serde_json::from_str::<Tokens>(&data) {
                let email = if !tokens.email.is_empty() {
                    tokens.email
                } else {
                    // Fallback: decode email from filename for pre-existing tokens.
                    path.file_stem()
                        .and_then(|s| s.to_str())
                        .map(|s| s.replace("_at_", "@").replace("_dot_", "."))
                        .unwrap_or_default()
                };
                if !email.is_empty() {
                    emails.push(email);
                }
            }
        }
    }
    emails.sort();
    emails
}

// ── Token refresh ─────────────────────────────────────────────────────────────

pub fn refresh(creds: &GoogleCreds, tokens: &Tokens) -> Result<Tokens, GoogleError> {
    let resp: serde_json::Value = reqwest::blocking::Client::new()
        .post("https://oauth2.googleapis.com/token")
        .form(&[
            ("client_id",     creds.client_id.as_str()),
            ("client_secret", creds.client_secret.as_str()),
            ("refresh_token", tokens.refresh_token.as_str()),
            ("grant_type",    "refresh_token"),
        ])
        .send()?
        .json()?;

    let access_token = resp["access_token"]
        .as_str()
        .ok_or_else(|| GoogleError::Auth(format!("refresh failed: {resp}")))?
        .to_owned();

    let expires_in = resp["expires_in"].as_i64().unwrap_or(3600);
    Ok(Tokens { access_token, refresh_token: tokens.refresh_token.clone(), expires_at: now_secs() + expires_in, email: tokens.email.clone() })
}

// ── Full OAuth flow ───────────────────────────────────────────────────────────

/// Run the browser-based PKCE flow. Returns `(tokens, email)`.
pub fn authorize(creds: &GoogleCreds, browser_path: Option<&str>) -> Result<(Tokens, String), GoogleError> {
    let listener = TcpListener::bind("127.0.0.1:0")?;
    let port     = listener.local_addr()?.port();

    let (verifier, challenge) = pkce_pair();
    let state        = random_hex(16);
    let redirect_uri = format!("http://localhost:{port}/callback");

    let auth_url = format!(
        "https://accounts.google.com/o/oauth2/auth\
         ?client_id={cid}\
         &redirect_uri={redir}\
         &response_type=code\
         &scope=https%3A%2F%2Fwww.googleapis.com%2Fauth%2Fcalendar%20email\
         &code_challenge={ch}\
         &code_challenge_method=S256\
         &state={state}\
         &access_type=offline\
         &prompt=consent",
        cid   = creds.client_id,
        redir = percent_encode(&redirect_uri),
        ch    = challenge,
        state = state,
    );

    if let Some(browser) = browser_path {
        std::process::Command::new(browser)
            .arg(&auth_url)
            .spawn()
            .map_err(|e| GoogleError::Auth(format!("cannot open browser '{browser}': {e}")))?;
    } else {
        open::that(&auth_url).map_err(|e| GoogleError::Auth(format!("cannot open browser: {e}")))?;
    }

    let code   = wait_for_code(listener, &state)?;
    let tokens = exchange_code(creds, &code, &verifier, &redirect_uri)?;
    let email  = get_user_email(&tokens.access_token)?;
    Ok((tokens, email))
}

// ── User info ─────────────────────────────────────────────────────────────────

pub fn get_user_email(access_token: &str) -> Result<String, super::GoogleError> {
    let response = reqwest::blocking::Client::new()
        .get("https://www.googleapis.com/oauth2/v1/userinfo")
        .bearer_auth(access_token)
        .send()?;

    let status = response.status();
    let body   = response.text()?;

    let resp: serde_json::Value = serde_json::from_str(&body)
        .map_err(|e| super::GoogleError::Auth(format!("userinfo parse error: {e} — body: {body}")))?;

    resp["email"]
        .as_str()
        .map(str::to_owned)
        .ok_or_else(|| super::GoogleError::Auth(format!("no email in userinfo (status={status}): {body}")))
}

// ── Helpers ───────────────────────────────────────────────────────────────────

fn pkce_pair() -> (String, String) {
    let bytes: Vec<u8> = (0..32).map(|_| rand::thread_rng().r#gen()).collect();
    let verifier  = base64url(&bytes);
    let challenge = base64url(&Sha256::digest(verifier.as_bytes()));
    (verifier, challenge)
}

fn base64url(input: &[u8]) -> String {
    const C: &[u8] = b"ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789-_";
    let mut out = String::with_capacity((input.len() * 4 + 2) / 3);
    for chunk in input.chunks(3) {
        let b0 = chunk[0] as u32;
        let b1 = chunk.get(1).copied().unwrap_or(0) as u32;
        let b2 = chunk.get(2).copied().unwrap_or(0) as u32;
        let n  = (b0 << 16) | (b1 << 8) | b2;
        out.push(C[((n >> 18) & 63) as usize] as char);
        out.push(C[((n >> 12) & 63) as usize] as char);
        if chunk.len() > 1 { out.push(C[((n >> 6) & 63) as usize] as char); }
        if chunk.len() > 2 { out.push(C[(n & 63) as usize] as char); }
    }
    out
}

fn percent_encode(s: &str) -> String {
    let mut out = String::new();
    for b in s.bytes() {
        if b.is_ascii_alphanumeric() || matches!(b, b'-' | b'_' | b'.' | b'~') {
            out.push(b as char);
        } else {
            out.push_str(&format!("%{b:02X}"));
        }
    }
    out
}

fn random_hex(bytes: usize) -> String {
    let v: Vec<u8> = (0..bytes).map(|_| rand::thread_rng().r#gen()).collect();
    hex::encode(v)
}

fn now_secs() -> i64 {
    SystemTime::now()
        .duration_since(UNIX_EPOCH)
        .unwrap_or_default()
        .as_secs() as i64
}

fn wait_for_code(listener: TcpListener, expected_state: &str) -> Result<String, GoogleError> {
    let expected = expected_state.to_owned();
    let (tx, rx) = std::sync::mpsc::channel::<Result<String, GoogleError>>();
    std::thread::spawn(move || {
        let result = listener
            .accept()
            .map_err(GoogleError::from)
            .and_then(|(stream, _)| handle_callback(stream, &expected));
        let _ = tx.send(result);
    });
    rx.recv_timeout(Duration::from_secs(300))
        .map_err(|_| GoogleError::Auth("OAuth flow timed out (5 min)".into()))?
}

fn handle_callback(mut stream: TcpStream, expected_state: &str) -> Result<String, GoogleError> {
    let mut request_line = String::new();
    {
        let mut reader = BufReader::new(&stream);
        reader.read_line(&mut request_line)?;
        for line in reader.lines() {
            if line?.is_empty() { break; }
        }
    }
    let body = "<!doctype html><html><body style='font-family:sans-serif;padding:40px'>\
                <h2>Authorized!</h2><p>You can close this tab and return to ruscal.</p>\
                </body></html>";
    write!(
        stream,
        "HTTP/1.1 200 OK\r\nContent-Length: {}\r\nContent-Type: text/html\r\n\r\n{}",
        body.len(), body
    )?;
    parse_callback_line(&request_line, expected_state)
}

fn parse_callback_line(line: &str, expected_state: &str) -> Result<String, GoogleError> {
    let path = line.split_whitespace().nth(1)
        .ok_or_else(|| GoogleError::Auth("malformed HTTP request line".into()))?;
    let query = path.splitn(2, '?').nth(1).unwrap_or("");
    let mut code  = None;
    let mut state = None;
    for param in query.split('&') {
        if let Some((k, v)) = param.split_once('=') {
            match k {
                "code"  => code  = Some(v.to_owned()),
                "state" => state = Some(v.to_owned()),
                _       => {}
            }
        }
    }
    if state.as_deref() != Some(expected_state) {
        return Err(GoogleError::Auth("CSRF state mismatch".into()));
    }
    code.ok_or_else(|| GoogleError::Auth("no 'code' parameter in callback URL".into()))
}

fn exchange_code(
    creds:        &GoogleCreds,
    code:         &str,
    verifier:     &str,
    redirect_uri: &str,
) -> Result<Tokens, GoogleError> {
    let resp: serde_json::Value = reqwest::blocking::Client::new()
        .post("https://oauth2.googleapis.com/token")
        .form(&[
            ("client_id",     creds.client_id.as_str()),
            ("client_secret", creds.client_secret.as_str()),
            ("code",          code),
            ("code_verifier", verifier),
            ("redirect_uri",  redirect_uri),
            ("grant_type",    "authorization_code"),
        ])
        .send()?
        .json()?;

    let access_token = resp["access_token"].as_str()
        .ok_or_else(|| GoogleError::Auth(format!("token exchange failed: {resp}")))?
        .to_owned();
    let refresh_token = resp["refresh_token"].as_str()
        .ok_or_else(|| GoogleError::Auth("no refresh_token in response".into()))?
        .to_owned();
    let expires_in = resp["expires_in"].as_i64().unwrap_or(3600);
    Ok(Tokens { access_token, refresh_token, expires_at: now_secs() + expires_in, email: String::new() })
}
