/// Google Calendar integration — OAuth 2.0 PKCE + CalDAV calendar list.
pub mod auth;
pub mod calendar;

/// A Google Calendar entry returned from CalDAV discovery.
pub struct GoogleCalendar {
    pub id:      String,
    pub summary: String,
}

use auth::GoogleCreds;

// ── Error type ────────────────────────────────────────────────────────────────

#[derive(Debug)]
pub enum GoogleError {
    Config(String),
    Auth(String),
    /// Google rejected the stored refresh token (`invalid_grant`). Happens on
    /// password change, explicit revoke, 6-month inactivity, or 7-day expiry
    /// for OAuth clients still in "Testing" verification mode. The caller
    /// must delete the stored token and run the full OAuth flow again.
    AuthRevoked(String),
    Api(String),
    Io(std::io::Error),
    Http(reqwest::Error),
    Json(serde_json::Error),
}

impl std::fmt::Display for GoogleError {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::Config(s)      => write!(f, "config: {s}"),
            Self::Auth(s)        => write!(f, "auth: {s}"),
            Self::AuthRevoked(s) => write!(f, "auth revoked: {s}"),
            Self::Api(s)         => write!(f, "API: {s}"),
            Self::Io(e)          => write!(f, "I/O: {e}"),
            Self::Http(e)        => write!(f, "HTTP: {e}"),
            Self::Json(e)        => write!(f, "JSON: {e}"),
        }
    }
}

impl From<std::io::Error>    for GoogleError { fn from(e: std::io::Error)    -> Self { Self::Io(e)   } }
impl From<reqwest::Error>    for GoogleError { fn from(e: reqwest::Error)    -> Self { Self::Http(e) } }
impl From<serde_json::Error> for GoogleError { fn from(e: serde_json::Error) -> Self { Self::Json(e) } }

// ── Credential loading ────────────────────────────────────────────────────────

fn load_creds() -> Result<GoogleCreds, GoogleError> {
    let client_id = option_env!("GOOGLE_CLIENT_ID")
        .ok_or_else(|| GoogleError::Config(
            "GOOGLE_CLIENT_ID not set at build time".into()
        ))?;
    let client_secret = option_env!("GOOGLE_CLIENT_SECRET")
        .ok_or_else(|| GoogleError::Config(
            "GOOGLE_CLIENT_SECRET not set at build time".into()
        ))?;
    Ok(GoogleCreds { client_id: client_id.to_owned(), client_secret: client_secret.to_owned() })
}

// ── Public API ────────────────────────────────────────────────────────────────

/// Get a valid access token for a known account.
///
/// Loads the stored token for `email`. If expired, refreshes it.
/// Falls back to the legacy single-account token file for migration.
/// Returns an error if no token exists — the caller must call
/// `authorize_new_account()` to do the OAuth flow.
pub fn get_access_token_for(email: &str) -> Result<String, GoogleError> {
    let creds = load_creds()?;

    // Try per-account token first.
    if let Some(mut tokens) = auth::load_tokens_for(email) {
        if tokens.is_expired() {
            match auth::refresh(&creds, &tokens) {
                Ok(new) => {
                    tokens = new;
                    auth::save_tokens_for(email, &tokens)?;
                }
                Err(e @ GoogleError::AuthRevoked(_)) => {
                    // Google has permanently rejected this refresh token.
                    // Delete it so the next attempt starts clean and the
                    // user isn't stuck in a retry loop against dead creds.
                    auth::revoke_tokens_for(email);
                    return Err(e);
                }
                Err(e) => return Err(e),
            }
        }
        return Ok(tokens.access_token);
    }

    // Migration: try the old single-account file and adopt it if it matches.
    if let Some(mut tokens) = auth::load_legacy_tokens() {
        if tokens.is_expired() {
            match auth::refresh(&creds, &tokens) {
                Ok(new) => tokens = new,
                Err(e @ GoogleError::AuthRevoked(_)) => {
                    auth::delete_legacy_tokens();
                    return Err(e);
                }
                Err(e) => return Err(e),
            }
        }
        let legacy_email = auth::get_user_email(&tokens.access_token).unwrap_or_default();
        if legacy_email == email {
            auth::save_tokens_for(email, &tokens)?;
            auth::delete_legacy_tokens();
            return Ok(tokens.access_token);
        }
    }

    Err(GoogleError::Auth(format!(
        "No stored token for {email} — please re-authenticate"
    )))
}

/// Run a full browser OAuth flow for a new (or switched) account.
///
/// Returns `(access_token, email)`. Tokens are saved to disk keyed by email.
pub fn authorize_new_account(browser_path: Option<&str>) -> Result<(String, String), GoogleError> {
    let creds = load_creds()?;
    let (tokens, email) = auth::authorize(&creds, browser_path)?;
    auth::save_tokens_for(&email, &tokens)?;
    Ok((tokens.access_token, email))
}

/// List calendars, optionally reusing an existing account's stored tokens.
///
/// - `email_hint = Some(email)` → tries stored tokens for that account first;
///   if not found runs OAuth (in case the user wants to reauth).
/// - `email_hint = None` → runs OAuth immediately (new account setup).
///
/// Returns `(calendars, email)` so the caller can record which account was used.
pub fn list_google_calendars(
    email_hint:   Option<&str>,
    browser_path: Option<&str>,
) -> Result<(Vec<GoogleCalendar>, String), GoogleError> {
    let (token, email) = match email_hint {
        Some(email) => {
            match get_access_token_for(email) {
                Ok(token) => (token, email.to_string()),
                // Token missing/stale — fall through to OAuth.
                Err(_) => authorize_new_account(browser_path)?,
            }
        }
        None => {
            // Check legacy token for seamless migration.
            let creds = load_creds()?;
            if let Some(mut tokens) = auth::load_legacy_tokens() {
                if tokens.is_expired() {
                    tokens = auth::refresh(&creds, &tokens)?;
                }
                let email = auth::get_user_email(&tokens.access_token)?;
                auth::save_tokens_for(&email, &tokens)?;
                auth::delete_legacy_tokens();
                (tokens.access_token, email)
            } else {
                authorize_new_account(browser_path)?
            }
        }
    };

    let calendars = calendar::list_calendars(&token)?;
    Ok((calendars, email))
}

/// List all emails that have stored token files.
pub fn list_stored_accounts() -> Vec<String> {
    auth::list_stored_accounts()
}

/// Remove the stored token for one account, forcing re-authentication.
pub fn sign_out_account(email: &str) {
    auth::revoke_tokens_for(email);
}
