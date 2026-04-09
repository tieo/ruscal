/// Generic CalDAV client — calendar discovery via PROPFIND.
///
/// Works with any CalDAV server: Google Calendar, Apple iCloud, Nextcloud, etc.
/// The caller supplies the calendar-home URL and a bearer token (or other auth).
use reqwest::blocking::Client;
use reqwest::header::{AUTHORIZATION, CONTENT_TYPE};

// ── Public types ──────────────────────────────────────────────────────────────

#[derive(Debug)]
pub struct CalDavCalendar {
    /// Absolute URL of the calendar collection (used for event sync later).
    pub href: String,
    /// Human-readable display name shown in the picker.
    pub display_name: String,
}

#[derive(Debug)]
pub enum CalDavError {
    Http(reqwest::Error),
    Xml(roxmltree::Error),
    Protocol(String),
}

impl std::fmt::Display for CalDavError {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::Http(e)      => write!(f, "HTTP: {e}"),
            Self::Xml(e)       => write!(f, "XML: {e}"),
            Self::Protocol(s)  => write!(f, "CalDAV: {s}"),
        }
    }
}

impl From<reqwest::Error>   for CalDavError { fn from(e: reqwest::Error)   -> Self { Self::Http(e) } }
impl From<roxmltree::Error> for CalDavError { fn from(e: roxmltree::Error) -> Self { Self::Xml(e)  } }

// ── Public API ────────────────────────────────────────────────────────────────

/// Discover the calendar-home-set URL given a known principal URL.
///
/// Sends one PROPFIND (Depth: 0) asking for `calendar-home-set` and returns
/// the absolute URL of the calendar home collection, which can be passed
/// directly to [`list_calendars`].
pub fn home_url_from_principal(
    principal_url: &str,
    access_token:  &str,
) -> Result<String, CalDavError> {
    let origin = url_origin(principal_url);

    let body = r#"<?xml version="1.0" encoding="UTF-8"?>
<D:propfind xmlns:D="DAV:" xmlns:C="urn:ietf:params:xml:ns:caldav">
  <D:prop><C:calendar-home-set/></D:prop>
</D:propfind>"#;

    let text = propfind(principal_url, access_token, "0", body)?;
    let doc  = roxmltree::Document::parse(&text)?;

    let home_href = doc
        .descendants()
        .find(|n| n.tag_name().name() == "calendar-home-set"
               && n.tag_name().namespace() == Some(NS_CALDAV))
        .and_then(|n| n.descendants().find(|c| c.tag_name().name() == "href"
                                            && c.tag_name().namespace() == Some(NS_DAV)))
        .and_then(|n| n.text())
        .ok_or_else(|| CalDavError::Protocol(
            format!("no calendar-home-set in response: {}", &text[..text.len().min(400)])
        ))?
        .trim()
        .to_owned();

    Ok(abs_url(&home_href, &origin))
}

/// Discover all calendar collections under `home_url` (PROPFIND Depth: 1).
///
/// `access_token` is sent as a Bearer token. Filters the multi-status response
/// to entries whose `resourcetype` contains a CalDAV `<calendar/>` element.
pub fn list_calendars(
    home_url:     &str,
    access_token: &str,
) -> Result<Vec<CalDavCalendar>, CalDavError> {
    let body = r#"<?xml version="1.0" encoding="UTF-8"?>
<D:propfind xmlns:D="DAV:" xmlns:C="urn:ietf:params:xml:ns:caldav">
  <D:prop>
    <D:displayname/>
    <D:resourcetype/>
  </D:prop>
</D:propfind>"#;

    let text = propfind(home_url, access_token, "1", body)?;
    parse_multistatus(&text, home_url)
}

// ── Shared HTTP helper ────────────────────────────────────────────────────────

fn propfind(url: &str, access_token: &str, depth: &str, body: &str) -> Result<String, CalDavError> {
    let method = reqwest::Method::from_bytes(b"PROPFIND").expect("valid method");
    let resp   = Client::new()
        .request(method, url)
        .header(AUTHORIZATION, format!("Bearer {access_token}"))
        .header(CONTENT_TYPE, "application/xml; charset=utf-8")
        .header("Depth", depth)
        .body(body.to_owned())
        .send()?;

    let status = resp.status();
    let text   = resp.text()?;

    if !status.is_success() && status.as_u16() != 207 {
        return Err(CalDavError::Protocol(format!(
            "server returned {status}: {}",
            text.chars().take(400).collect::<String>()
        )));
    }
    Ok(text)
}

fn abs_url(href: &str, origin: &str) -> String {
    if href.starts_with("http") { href.to_owned() } else { format!("{origin}{href}") }
}

// ── XML parsing ───────────────────────────────────────────────────────────────

const NS_DAV:    &str = "DAV:";
const NS_CALDAV: &str = "urn:ietf:params:xml:ns:caldav";

fn parse_multistatus(xml: &str, home_url: &str) -> Result<Vec<CalDavCalendar>, CalDavError> {
    let doc  = roxmltree::Document::parse(xml)?;
    let root = doc.root_element();

    // Derive the server origin so we can resolve relative hrefs.
    let origin = url_origin(home_url);

    let mut calendars = Vec::new();

    for response in root.descendants().filter(|n| {
        n.tag_name().name() == "response" && n.tag_name().namespace() == Some(NS_DAV)
    }) {
        // Skip entries whose resourcetype doesn't include <C:calendar/>.
        if !has_calendar_type(&response) {
            continue;
        }

        let href = response
            .descendants()
            .find(|n| n.tag_name().name() == "href" && n.tag_name().namespace() == Some(NS_DAV))
            .and_then(|n| n.text())
            .unwrap_or("")
            .trim()
            .to_owned();

        // Skip the home collection itself (its href == the request path).
        if href.is_empty() || is_home_url(&href, home_url) {
            continue;
        }

        let display_name = response
            .descendants()
            .find(|n| {
                n.tag_name().name() == "displayname"
                    && n.tag_name().namespace() == Some(NS_DAV)
            })
            .and_then(|n| n.text())
            .map(str::trim)
            .filter(|s| !s.is_empty())
            .unwrap_or(&href)
            .to_owned();

        // Make the href absolute.
        let abs_href = if href.starts_with("http") {
            href
        } else {
            format!("{origin}{href}")
        };

        calendars.push(CalDavCalendar { href: abs_href, display_name });
    }

    Ok(calendars)
}

/// True if a `<response>` element's `resourcetype` contains `<C:calendar/>`.
fn has_calendar_type(response: &roxmltree::Node) -> bool {
    response.descendants().any(|n| {
        n.tag_name().name() == "calendar"
            && n.tag_name().namespace() == Some(NS_CALDAV)
    })
}

/// Extract `scheme://host:port` from a URL.
fn url_origin(url: &str) -> String {
    // Find the third slash: "https://host/path" → "https://host"
    let after_scheme = url.find("://").map(|i| i + 3).unwrap_or(0);
    let path_start   = url[after_scheme..].find('/').map(|i| i + after_scheme)
        .unwrap_or(url.len());
    url[..path_start].to_owned()
}

/// True if `href` refers to the same resource as `home_url` (ignoring trailing slash).
fn is_home_url(href: &str, home_url: &str) -> bool {
    fn norm(s: &str) -> &str { s.trim_end_matches('/') }
    // Compare path portions or full URLs.
    let home_path = home_url.find("://")
        .and_then(|i| home_url[i + 3..].find('/').map(|j| &home_url[i + 3 + j..]))
        .unwrap_or(home_url);
    norm(href) == norm(home_url) || norm(href) == norm(home_path)
}
