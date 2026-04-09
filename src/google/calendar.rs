/// Google Calendar via CalDAV.
///
/// Google's CalDAV principal URL requires the user's email address in the path:
/// `https://apidata.googleusercontent.com/caldav/v1/{email}/user/`
///
/// We get the email from the userinfo endpoint, then PROPFIND the principal for
/// `calendar-home-set`, then list calendar collections from the home URL.
use super::{auth, GoogleCalendar, GoogleError};
use crate::caldav;

const CALDAV_ORIGIN: &str = "https://apidata.googleusercontent.com";

pub fn list_calendars(access_token: &str) -> Result<Vec<GoogleCalendar>, GoogleError> {
    let email = auth::get_user_email(access_token)?;
    log::debug!("got email: {email}");

    // Google's CalDAV principal path requires the email with @ encoded as %40.
    let principal_url = format!(
        "{}/caldav/v1/{}/user/",
        CALDAV_ORIGIN,
        email.replace('@', "%40"),
    );
    log::debug!("principal_url: {principal_url}");

    let home_url = caldav::home_url_from_principal(&principal_url, access_token)
        .map_err(|e| GoogleError::Api(e.to_string()))?;

    let cals = caldav::list_calendars(&home_url, access_token)
        .map_err(|e| GoogleError::Api(e.to_string()))?;

    Ok(cals
        .into_iter()
        .map(|c| GoogleCalendar { id: c.href, summary: c.display_name })
        .collect())
}
