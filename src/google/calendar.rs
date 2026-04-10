/// Google Calendar via CalDAV.
///
/// Google's CalDAV home URL has the form:
/// `https://apidata.googleusercontent.com/caldav/v2/{email%40domain}/`
///
/// We fetch the user's email from the userinfo endpoint, build the home URL
/// directly, then list calendar collections with a PROPFIND Depth:1.
use super::{auth, GoogleCalendar, GoogleError};
use crate::caldav;

const CALDAV_ORIGIN: &str = "https://apidata.googleusercontent.com";

pub fn list_calendars(access_token: &str) -> Result<Vec<GoogleCalendar>, GoogleError> {
    let email = auth::get_user_email(access_token)?;
    log::debug!("got email: {email}");

    // Google's CalDAV home URL — @ must be percent-encoded as %40.
    let home_url = format!(
        "{}/caldav/v2/{}/",
        CALDAV_ORIGIN,
        email.replace('@', "%40"),
    );
    log::debug!("home_url: {home_url}");

    let cals = caldav::list_calendars(&home_url, access_token)
        .map_err(|e| GoogleError::Api(e.to_string()))?;

    Ok(cals
        .into_iter()
        .map(|c| GoogleCalendar { id: c.href, summary: c.display_name })
        .collect())
}
