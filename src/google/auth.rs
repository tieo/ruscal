/// OAuth 2.0 PKCE flow for installed desktop applications.
///
/// # Flow
/// 1. Bind a random-port localhost TCP listener.
/// 2. Build the Google authorization URL with a PKCE challenge.
/// 3. Open the URL in the system browser via [`open`].
/// 4. Wait (up to 5 min) for Google to redirect to `localhost:<port>/callback?code=…`.
/// 5. Exchange the code + PKCE verifier for access + refresh tokens.
/// 6. Persist tokens to disk; return access token.
use std::io::{BufRead, BufReader, Write as _};
use std::net::{TcpListener, TcpStream};
use std::time::{Duration, SystemTime, UNIX_EPOCH};

use rand::Rng;
use sha2::{Digest, Sha256};

use super::GoogleError;

// ── Public types ──────────────────────────────────────────────────────────────

/// Client credentials from Google Cloud Console.
pub struct GoogleCreds {
    pub client_id:     String,
    pub client_secret: String,
}

/// Access + refresh token pair with expiry.
#[derive(serde::Serialize, serde::Deserialize, Clone)]
pub struct Tokens {
    pub access_token:  String,
    pub refresh_token: String,
    /// Unix timestamp (seconds) when `access_token` expires.
    pub expires_at:    i64,
}

impl Tokens {
    /// True when the access token is within 60 s of expiry.
    pub fn is_expired(&self) -> bool {
        let now = SystemTime::now()
            .duration_since(UNIX_EPOCH)
            .unwrap_or_default()
            .as_secs() as i64;
        now >= self.expires_at - 60
    }
}

// ── Token persistence ─────────────────────────────────────────────────────────

fn token_path() -> Option<std::path::PathBuf> {
    dirs::config_dir().map(|d| d.join("ruscal").join("google_token.json"))
}

/// Load tokens from disk. Returns `None` if not found or unreadable.
pub fn load_tokens() -> Option<Tokens> {
    let data = std::fs::read_to_string(token_path()?).ok()?;
    serde_json::from_str(&data).ok()
}

/// Persist tokens to `~/.config/ruscal/google_token.json`.
pub fn save_tokens(tokens: &Tokens) -> Result<(), GoogleError> {
    let path = token_path()
        .ok_or_else(|| GoogleError::Config("cannot locate config directory".into()))?;
    if let Some(parent) = path.parent() {
        std::fs::create_dir_all(parent)?;
    }
    std::fs::write(path, serde_json::to_string_pretty(tokens)?)?;
    Ok(())
}

// ── Token refresh ─────────────────────────────────────────────────────────────

/// Use the stored refresh token to get a new access token.
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
    let expires_at = now_secs() + expires_in;

    Ok(Tokens { access_token, refresh_token: tokens.refresh_token.clone(), expires_at })
}

// ── Full OAuth flow ───────────────────────────────────────────────────────────

/// Run the browser-based OAuth 2.0 PKCE authorization flow.
///
/// Opens the system browser for the user to grant access, then captures the
/// redirect on a local listener and exchanges the code for tokens.
pub fn authorize(creds: &GoogleCreds) -> Result<Tokens, GoogleError> {
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

    open::that(&auth_url).map_err(|e| GoogleError::Auth(format!("cannot open browser: {e}")))?;

    let code = wait_for_code(listener, &state)?;
    exchange_code(creds, &code, &verifier, &redirect_uri)
}

// ── Helpers ───────────────────────────────────────────────────────────────────

/// Generate a (code_verifier, code_challenge) PKCE pair.
fn pkce_pair() -> (String, String) {
    let bytes: Vec<u8> = (0..32).map(|_| rand::thread_rng().r#gen()).collect();
    let verifier  = base64url(&bytes);
    let challenge = base64url(&Sha256::digest(verifier.as_bytes()));
    (verifier, challenge)
}

/// URL-safe base64 encoding without padding (RFC 4648 §5).
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

/// Percent-encode all bytes that are not RFC 3986 unreserved characters.
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

// ── Local callback listener ───────────────────────────────────────────────────

/// Wait (up to 5 minutes) for Google's redirect to our local callback URL.
///
/// Spawns an inner thread so we can apply a timeout without blocking forever
/// if the user cancels in the browser.
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

/// Read the HTTP request from the browser, send a success page, extract the code.
fn handle_callback(mut stream: TcpStream, expected_state: &str) -> Result<String, GoogleError> {
    // Read the first line of the HTTP request (GET /callback?… HTTP/1.1).
    let mut request_line = String::new();
    {
        let mut reader = BufReader::new(&stream);
        reader.read_line(&mut request_line)?;
        // Drain remaining headers so the browser doesn't show "connection reset".
        for line in reader.lines() {
            if line?.is_empty() { break; }
        }
    }

    // Always send a response so the browser shows something meaningful.
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

/// Extract `code` from a request line like `GET /callback?code=X&state=Y HTTP/1.1`.
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

// ── User info ─────────────────────────────────────────────────────────────────

/// Fetch the authenticated user's email address.
///
/// Used to build the CalDAV home URL: `https://apidata.googleusercontent.com/caldav/v1/{email}/`
#[allow(dead_code)] // will be used for displaying the signed-in account
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

// ── Code exchange ─────────────────────────────────────────────────────────────

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

    Ok(Tokens { access_token, refresh_token, expires_at: now_secs() + expires_in })
}
