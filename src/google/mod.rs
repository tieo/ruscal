/// Google Calendar integration — OAuth 2.0 PKCE + CalDAV calendar list.
pub mod auth;
pub mod calendar;

/// A Google Calendar entry returned from CalDAV discovery.
pub struct GoogleCalendar {
    /// CalDAV collection URL — used for event sync.
    pub id:      String,
    /// Human-readable display name shown in the picker.
    pub summary: String,
}

use auth::GoogleCreds;

// ── Error type ────────────────────────────────────────────────────────────────

#[derive(Debug)]
pub enum GoogleError {
    /// Problem reading the config file or finding credentials.
    Config(String),
    /// OAuth flow failure (browser, CSRF, token exchange, …).
    Auth(String),
    /// Google Calendar API returned an unexpected response.
    Api(String),
    /// I/O error (network, file system).
    Io(std::io::Error),
    /// HTTP-layer error from reqwest.
    Http(reqwest::Error),
    /// JSON parse error.
    Json(serde_json::Error),
}

impl std::fmt::Display for GoogleError {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::Config(s) => write!(f, "config: {s}"),
            Self::Auth(s)   => write!(f, "auth: {s}"),
            Self::Api(s)    => write!(f, "API: {s}"),
            Self::Io(e)     => write!(f, "I/O: {e}"),
            Self::Http(e)   => write!(f, "HTTP: {e}"),
            Self::Json(e)   => write!(f, "JSON: {e}"),
        }
    }
}

impl From<std::io::Error>    for GoogleError { fn from(e: std::io::Error)    -> Self { Self::Io(e)   } }
impl From<reqwest::Error>    for GoogleError { fn from(e: reqwest::Error)    -> Self { Self::Http(e) } }
impl From<serde_json::Error> for GoogleError { fn from(e: serde_json::Error) -> Self { Self::Json(e) } }

// ── Credential loading ────────────────────────────────────────────────────────

/// Load OAuth client credentials from `~/.config/ruscal/.env`.
///
/// Expected format:
/// ```env
/// GOOGLE_CLIENT_ID=…
/// GOOGLE_CLIENT_SECRET=…
/// ```
fn load_creds() -> Result<GoogleCreds, GoogleError> {
    let path = dirs::config_dir()
        .ok_or_else(|| GoogleError::Config("cannot locate config directory".into()))?
        .join("ruscal")
        .join(".env");

    dotenvy::from_path(&path).map_err(|_| {
        GoogleError::Config(format!(
            "credentials file not found at {}.\n\
             Create it with:\n\
             GOOGLE_CLIENT_ID=<your-client-id>\n\
             GOOGLE_CLIENT_SECRET=<your-client-secret>",
            path.display()
        ))
    })?;

    let client_id = std::env::var("GOOGLE_CLIENT_ID")
        .map_err(|_| GoogleError::Config("GOOGLE_CLIENT_ID not set in .env".into()))?;

    let client_secret = std::env::var("GOOGLE_CLIENT_SECRET")
        .map_err(|_| GoogleError::Config("GOOGLE_CLIENT_SECRET not set in .env".into()))?;

    Ok(GoogleCreds { client_id, client_secret })
}

// ── Public API ────────────────────────────────────────────────────────────────

/// Get a valid access token.
///
/// Loads stored tokens and refreshes if expired. If no tokens are stored,
/// runs the full OAuth 2.0 PKCE flow (opens a browser window).
pub fn get_access_token() -> Result<String, GoogleError> {
    let creds = load_creds()?;

    if let Some(mut tokens) = auth::load_tokens() {
        if tokens.is_expired() {
            tokens = auth::refresh(&creds, &tokens)?;
            auth::save_tokens(&tokens)?;
        }
        return Ok(tokens.access_token);
    }

    // No stored tokens — do the full browser-based OAuth flow.
    let tokens = auth::authorize(&creds)?;
    auth::save_tokens(&tokens)?;
    Ok(tokens.access_token)
}

/// List all Google Calendars for the authenticated user.
pub fn list_google_calendars() -> Result<Vec<GoogleCalendar>, GoogleError> {
    let token = get_access_token()?;
    calendar::list_calendars(&token)
}
