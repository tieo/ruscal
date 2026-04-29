# Privacy Policy

`ruscal` is a Free and Open Source desktop tool that synchronises calendar events between supported calendar systems on your own machine. The application is provided at no cost and is intended for use as is.

Calendar systems currently supported include Microsoft Outlook (via the locally-installed Outlook MAPI interface), Google Calendar (via the Google Calendar API), and generic CalDAV servers configured by the user. Support for additional calendar systems may be added in future versions.

This page describes how `ruscal` handles personal information.

For calendar systems that authenticate via OAuth (such as Google), `ruscal` requires you to authorize it to access your calendar. The resulting OAuth tokens are stored on your computer and are not shared with anyone else. `ruscal` uses these tokens to read and write calendar events on your behalf. For calendar systems that authenticate locally (such as Outlook MAPI on Windows), no separate credential is held by `ruscal`; it uses the existing local session.

**None of your data is collected, stored, processed or shared with the developer or any third parties.** You are welcome to inspect the source code or open an issue if you have any questions.

The app does rely on third-party services that you choose to connect (e.g. the [Google Calendar API](https://developers.google.com/calendar/), Microsoft Outlook, or a CalDAV server you configure), each of which may collect information used to identify you. We recommend you read the privacy policies of any service you connect, for example the [Google Privacy Policy](https://policies.google.com/privacy).

## Links to Other Services

`ruscal` only communicates with calendar systems that you yourself connect (Outlook on the local machine, Google Calendar after explicit OAuth authorization, or a CalDAV server URL you provide). We have no control over and assume no responsibility for the content, privacy policies, or practices of those services.

## How `ruscal` Works

When you connect a Google account, `ruscal` requests the following Google OAuth scopes:

- `https://www.googleapis.com/auth/calendar`: read and write your Google calendars and events, so they can be kept in sync with the other calendar(s) you've configured.
- `email`: read your Google account email address, so the app can label which Google account a stored token belongs to.

Authorized OAuth tokens are stored on your local machine in `%LOCALAPPDATA%\ruscal\tokens` (Windows). You can revoke a token at any time by deleting the relevant file in that folder, by using the in-app sign-out, or (for Google specifically) by removing the application from your [Google Account permissions page](https://myaccount.google.com/permissions). Other than the token, no account data is persisted on disk by `ruscal`.

`ruscal` reads events from each connected calendar, compares them, and writes the differences to keep them in sync. **No third-party services are used in the process. None of your calendar data is shared with the developer or any third parties.** Synchronisation runs entirely on your own machine.

## Changes to This Privacy Policy

This Privacy Policy may be updated from time to time. You are advised to review this page periodically for any changes. Changes are effective immediately after they are posted on this page.

## Contact

If you have any questions or suggestions about this Privacy Policy, please open an issue at [github.com/tieo/ruscal](https://github.com/tieo/ruscal/issues).

---

*This policy is adapted from [`slgobinath/gcalendar`'s privacy policy](https://github.com/slgobinath/gcalendar/blob/master/privacy_policy.md), which describes a comparable local-only desktop calendar tool.*
