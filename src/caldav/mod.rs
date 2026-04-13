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

impl CalDavError {
    /// Extract the HTTP status code from a Protocol error, if any.
    ///
    /// Protocol errors are formatted as "PUT {uid} returned {status}: {body}".
    /// This parses out the numeric status so callers can branch on it without
    /// string-matching.
    pub fn http_status(&self) -> Option<u16> {
        if let Self::Protocol(s) = self {
            // Find "returned NNN" in the message — our own format from put_event / delete_event.
            s.split_whitespace()
                .skip_while(|w| *w != "returned")
                .nth(1)
                .and_then(|w| w.trim_end_matches(':').parse::<u16>().ok())
        } else {
            None
        }
    }
}

impl From<reqwest::Error>   for CalDavError { fn from(e: reqwest::Error)   -> Self { Self::Http(e) } }
impl From<roxmltree::Error> for CalDavError { fn from(e: roxmltree::Error) -> Self { Self::Xml(e)  } }

// ── Shared HTTP client ────────────────────────────────────────────────────────

/// Build a reqwest blocking client with a 30-second timeout.
///
/// All CalDAV operations use this so the sync cannot hang indefinitely if
/// Google Calendar is unresponsive.
fn client() -> Client {
    Client::builder()
        .timeout(std::time::Duration::from_secs(30))
        .build()
        .unwrap_or_else(|_| Client::new())
}

// ── Public API ────────────────────────────────────────────────────────────────

/// Discover the calendar-home-set URL given a known principal URL.
///
/// Sends one PROPFIND (Depth: 0) asking for `calendar-home-set` and returns
/// the absolute URL of the calendar home collection, which can be passed
/// directly to [`list_calendars`].
#[allow(dead_code)]
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

/// Create or update a single iCalendar event in a CalDAV collection.
///
/// Uses HTTP PUT with the event UID as the filename. CalDAV servers treat
/// PUT as an upsert — existing events with the same UID filename are
/// overwritten, so repeated syncs are idempotent.
pub fn put_event(
    calendar_url: &str,
    uid:          &str,
    ical:         &str,
    access_token: &str,
) -> Result<(), CalDavError> {
    put_event_inner(calendar_url, uid, ical, access_token, None)
}

fn put_event_inner(
    calendar_url: &str,
    uid:          &str,
    ical:         &str,
    access_token: &str,
    if_match:     Option<&str>,
) -> Result<(), CalDavError> {
    let url  = format!("{}/{uid}.ics", calendar_url.trim_end_matches('/'));
    let mut req = client()
        .put(&url)
        .header(AUTHORIZATION, format!("Bearer {access_token}"))
        .header(CONTENT_TYPE, "text/calendar; charset=utf-8");
    if let Some(etag) = if_match {
        req = req.header("If-Match", etag);
    }
    let resp = req.body(ical.to_owned()).send()?;

    let status = resp.status();
    if !status.is_success() {
        let body = resp.text()?;
        return Err(CalDavError::Protocol(format!(
            "PUT {uid} returned {status}: {}",
            body.chars().take(400).collect::<String>()
        )));
    }
    Ok(())
}

/// List the UIDs of all ruscal-managed events in a CalDAV collection.
///
/// Uses PROPFIND Depth:1 to enumerate .ics resources. Ruscal events are
/// identified by their UID filename suffix `@ruscal.ics`.
pub fn list_ruscal_event_uids(
    calendar_url: &str,
    access_token: &str,
) -> Result<Vec<String>, CalDavError> {
    let body = r#"<?xml version="1.0" encoding="UTF-8"?>
<D:propfind xmlns:D="DAV:">
  <D:prop><D:getetag/></D:prop>
</D:propfind>"#;

    let text   = propfind(calendar_url, access_token, "1", body)?;
    let doc    = roxmltree::Document::parse(&text)?;

    let mut uids = Vec::new();
    for node in doc.descendants().filter(|n| {
        n.tag_name().name() == "href" && n.tag_name().namespace() == Some(NS_DAV)
    }) {
        let href = node.text().unwrap_or("").trim();
        // Ruscal events are stored as "{uid}.ics"; uid ends in "@ruscal".
        // Google URL-encodes @ as %40 in the href.
        if href.ends_with("%40ruscal.ics") || href.ends_with("@ruscal.ics") {
            // Strip ".ics" and URL-decode "%40" back to "@"
            let uid = href
                .trim_end_matches(".ics")
                .rsplit('/')
                .next()
                .unwrap_or("")
                .replace("%40", "@");
            if !uid.is_empty() {
                uids.push(uid);
            }
        }
    }

    Ok(uids)
}

/// Find all hrefs for events with a specific UID via CalDAV calendar-query REPORT.
///
/// Google may store an event under a URL we did not choose (e.g., when ORGANIZER
/// is an external address). This function finds the actual server-side href(s)
/// so we can DELETE them before re-PUTting.
pub fn find_event_hrefs_by_uid(
    calendar_url: &str,
    uid:          &str,
    access_token: &str,
) -> Result<Vec<String>, CalDavError> {
    let body = format!(
        r#"<?xml version="1.0" encoding="UTF-8"?>
<C:calendar-query xmlns:D="DAV:" xmlns:C="urn:ietf:params:xml:ns:caldav">
  <D:prop><D:getetag/></D:prop>
  <C:filter>
    <C:comp-filter name="VCALENDAR">
      <C:comp-filter name="VEVENT">
        <C:prop-filter name="UID">
          <C:text-match>{uid}</C:text-match>
        </C:prop-filter>
      </C:comp-filter>
    </C:comp-filter>
  </C:filter>
</C:calendar-query>"#
    );

    let method = reqwest::Method::from_bytes(b"REPORT").expect("valid method");
    let resp   = client()
        .request(method, calendar_url)
        .header(AUTHORIZATION, format!("Bearer {access_token}"))
        .header(CONTENT_TYPE, "application/xml; charset=utf-8")
        .header("Depth", "1")
        .body(body)
        .send()?;

    let status = resp.status();
    let text   = resp.text()?;

    // 207 Multi-Status = results; 404 / other = no results or not supported.
    if status.as_u16() != 207 {
        return Ok(Vec::new()); // graceful: REPORT may not be supported
    }

    let doc = match roxmltree::Document::parse(&text) {
        Ok(d)  => d,
        Err(_) => return Ok(Vec::new()),
    };

    let origin = url_origin(calendar_url);
    let mut hrefs = Vec::new();

    for node in doc.descendants().filter(|n| {
        n.tag_name().name() == "href" && n.tag_name().namespace() == Some(NS_DAV)
    }) {
        let href = node.text().unwrap_or("").trim();
        if href.ends_with(".ics") {
            hrefs.push(abs_url(href, &origin));
        }
    }

    Ok(hrefs)
}

/// Delete a CalDAV resource at an arbitrary absolute URL.
///
/// Used to remove events that were stored by the server at a server-generated
/// URL rather than the UID-based filename we would normally use.
pub fn delete_event_at_url(
    url:          &str,
    access_token: &str,
) -> Result<(), CalDavError> {
    let resp = client()
        .delete(url)
        .header(AUTHORIZATION, format!("Bearer {access_token}"))
        .send()?;

    let status = resp.status();
    if !status.is_success() && status.as_u16() != 404 {
        let body = resp.text()?;
        return Err(CalDavError::Protocol(format!(
            "DELETE {url} returned {status}: {}",
            body.chars().take(400).collect::<String>()
        )));
    }
    Ok(())
}

/// Delete a single event from a CalDAV collection.
///
/// Returns `true` if the server confirmed deletion (2xx), `false` if the
/// resource was already absent (404). Errors on any other status.
pub fn delete_event(
    calendar_url: &str,
    uid:          &str,
    access_token: &str,
) -> Result<bool, CalDavError> {
    let url  = format!("{}/{uid}.ics", calendar_url.trim_end_matches('/'));
    let resp = client()
        .delete(&url)
        .header(AUTHORIZATION, format!("Bearer {access_token}"))
        .send()?;

    let status = resp.status();
    if status.as_u16() == 404 {
        return Ok(false); // already gone
    }
    if !status.is_success() {
        let body = resp.text()?;
        return Err(CalDavError::Protocol(format!(
            "DELETE {uid} returned {status}: {}",
            body.chars().take(400).collect::<String>()
        )));
    }
    Ok(true) // actually deleted
}

#[allow(dead_code)]
/// Fetch a single event by UID from a CalDAV collection.
///
/// Returns the raw iCalendar text as stored on the server. Useful for verifying
/// that the server accepted and stored what we PUT.
pub fn get_event(
    calendar_url: &str,
    uid:          &str,
    access_token: &str,
) -> Result<String, CalDavError> {
    let url = format!("{}/{uid}.ics", calendar_url.trim_end_matches('/'));
    let resp = client()
        .get(&url)
        .header(AUTHORIZATION, format!("Bearer {access_token}"))
        .send()?;

    let status = resp.status();
    let body   = resp.text()?;

    if !status.is_success() {
        return Err(CalDavError::Protocol(format!(
            "GET {uid} returned {status}: {}",
            body.chars().take(400).collect::<String>()
        )));
    }
    Ok(body)
}

// ── Shared HTTP helper ────────────────────────────────────────────────────────

fn propfind(url: &str, access_token: &str, depth: &str, body: &str) -> Result<String, CalDavError> {
    let method = reqwest::Method::from_bytes(b"PROPFIND").expect("valid method");
    let resp   = client()
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

#[allow(dead_code)]
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
