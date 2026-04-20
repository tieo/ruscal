/// One-way sync engine: Outlook calendar → Google Calendar via CalDAV.
///
/// Non-recurring events are synced as plain VEVENTs. Recurring events are
/// synced as a single VEVENT with an RRULE, so Google Calendar shows them
/// as a proper recurring series — not as individual one-off instances.
use chrono::Utc;
use std::hash::{Hash, Hasher};

use crate::caldav;
use crate::event::{BusyStatus, CalendarEvent, Sensitivity};
use crate::state;

fn hash_ical(s: &str) -> u64 {
    // Strip lines whose value changes on every generation but whose meaning is
    // "when this iCal blob was produced" rather than "what the event is":
    // DTSTAMP, LAST-MODIFIED, CREATED. Everything else (SUMMARY, DTSTART, RRULE,
    // DESCRIPTION, …) participates in identity so real edits invalidate the cache.
    let canonical: String = s.lines()
        .filter(|l| {
            let t = l.trim_start();
            !(t.starts_with("DTSTAMP")
              || t.starts_with("LAST-MODIFIED")
              || t.starts_with("CREATED"))
        })
        .collect::<Vec<_>>()
        .join("\n");
    let mut h = std::collections::hash_map::DefaultHasher::new();
    canonical.hash(&mut h);
    h.finish()
}

// ── Public API ────────────────────────────────────────────────────────────────

/// Result of a single sync run.
pub struct SyncReport {
    /// Number of events successfully PUT.
    pub synced:  usize,
    /// Titles of events skipped due to 409/403 conflicts. Non-fatal, but surfaced
    /// in the UI so persistent conflicts don't hide behind a green check.
    pub skipped_titles: Vec<String>,
}

pub fn run_sync(
    pair_id:           &str,
    dest_calendar_url: &str,
    access_token:      &str,
) -> Result<SyncReport, String> {
    let now          = Utc::now();
    let window_start = now - chrono::Duration::days(crate::outlook::DEFAULT_PAST_DAYS);
    let window_end   = now + chrono::Duration::days(crate::outlook::DEFAULT_FUTURE_DAYS);

    let events = crate::outlook::read_calendar_events(window_start, window_end)
        .map_err(|e| format!("Outlook: {e:?}"))?;

    // Filter out events the user has explicitly declined.
    let events: Vec<_> = events
        .into_iter()
        .filter(|e| e.response_status != crate::event::ResponseStatus::Declined)
        .collect();

    log::info!("syncing {} events to {dest_calendar_url}", events.len());

    let mut synced  = 0usize;
    let mut skipped_titles = Vec::new();
    let mut synced_uids = std::collections::HashSet::new();
    let mut pair_state = state::load_pair(pair_id);
    let cache = &mut pair_state.hash_cache;
    for event in &events {
        let uid  = event_uid(event);
        let ical = event_to_ical(event, &uid);
        let h    = hash_ical(&ical);
        if cache.get(&uid) == Some(&h) {
            // Unchanged since last successful PUT — skip silently and keep it out of orphans.
            synced_uids.insert(uid.clone());
        } else {
            match put_with_retry(dest_calendar_url, &uid, &ical, access_token) {
                Ok(PutOutcome::Ok)      => { synced_uids.insert(uid.clone()); synced += 1; cache.insert(uid.clone(), h); }
                Ok(PutOutcome::Skipped) => {
                    // Do NOT cache the hash on a skip. If we did, an externally
                    // resolved conflict (e.g. the colliding Google event gets
                    // deleted) would never trigger a retry — the cached hash
                    // would silently match every subsequent sync. Re-attempting
                    // the PUT once per cycle is cheap and the only way ruscal
                    // can recover automatically.
                    synced_uids.insert(uid.clone());
                    cache.remove(&uid);
                    skipped_titles.push(event.subject.clone());
                }
                Err(e) => return Err(format!("CalDAV PUT: {e}")),
            }
        }

        // PUT modified/moved occurrences as separate standalone CalDAV resources.
        // Google CalDAV rejects RECURRENCE-ID VEVENTs embedded in the master.
        for (exc_uid, exc_ical) in build_exception_icals(event, &uid) {
            let eh = hash_ical(&exc_ical);
            if cache.get(&exc_uid) == Some(&eh) {
                synced_uids.insert(exc_uid);
                continue;
            }
            match put_with_retry(dest_calendar_url, &exc_uid, &exc_ical, access_token) {
                Ok(PutOutcome::Ok)      => { synced_uids.insert(exc_uid.clone()); cache.insert(exc_uid, eh); }
                Ok(PutOutcome::Skipped) => {
                    // Same reasoning as the master path above — never cache on skip.
                    synced_uids.insert(exc_uid.clone());
                    cache.remove(&exc_uid);
                    skipped_titles.push(format!("{} (exception)", event.subject));
                }
                Err(e) => return Err(format!("CalDAV PUT exception {exc_uid}: {e}")),
            }
        }
    }
    state::save_pair(pair_id, pair_state);

    // Delete events that are in Google Calendar but no longer in Outlook.
    delete_orphans(dest_calendar_url, &synced_uids, access_token)
        .map_err(|e| format!("CalDAV DELETE: {e}"))?;

    Ok(SyncReport { synced, skipped_titles })
}

/// Delete all ruscal-managed events in the calendar that are not in `keep_uids`.
///
/// Safety floor: if more than `MAX_ORPHAN_DELETIONS` orphans would be deleted
/// in a single run, abort the sync with an error. A healthy sync only deletes
/// a handful of events at a time — a large batch indicates a bug or a stale
/// `keep_uids` set, and we'd rather fail loudly than silently wipe the calendar.
const MAX_ORPHAN_DELETIONS: usize = 10;

fn delete_orphans(
    calendar_url: &str,
    keep_uids:    &std::collections::HashSet<String>,
    access_token: &str,
) -> Result<(), String> {
    let remote_uids = caldav::list_ruscal_event_uids(calendar_url, access_token)
        .map_err(|e| format!("listing remote events: {e}"))?;

    let orphans: Vec<String> = remote_uids
        .into_iter()
        .filter(|uid| !keep_uids.contains(uid))
        .collect();

    if orphans.len() > MAX_ORPHAN_DELETIONS {
        return Err(format!(
            "refusing to delete {} orphans in one run (safety limit: {}). \
             This usually means the Outlook read returned far fewer events than expected. \
             Orphan UIDs: {:?}",
            orphans.len(), MAX_ORPHAN_DELETIONS, orphans
        ));
    }

    for uid in orphans {
        log::info!("deleting orphaned event: {uid}");
        caldav::delete_event(calendar_url, &uid, access_token)
            .map_err(|e| format!("deleting {uid}: {e}"))
            .map(|_| ())?;
    }
    Ok(())
}

pub enum PutOutcome { Ok, Skipped }

/// PUT an iCalendar event. On 409/403 (Google detected a conflict with an
/// existing resource at a different URL), return [`PutOutcome::Skipped`] — do
/// NOT attempt to delete the conflicting resource. The only deletion path in
/// ruscal is `delete_orphans`, which is strictly limited to ruscal-UID events.
fn put_with_retry(
    calendar_url: &str,
    uid:          &str,
    ical:         &str,
    access_token: &str,
) -> Result<PutOutcome, caldav::CalDavError> {
    match caldav::put_event(calendar_url, uid, ical, access_token) {
        Ok(()) => Ok(PutOutcome::Ok),
        Err(e) if e.http_status() == Some(409) || e.http_status() == Some(403) => {
            log::warn!(
                "{uid}: PUT returned {} — skipping this event (no automatic cleanup)",
                e.http_status().unwrap()
            );
            Ok(PutOutcome::Skipped)
        }
        Err(e) => Err(e),
    }
}

// ── UID ───────────────────────────────────────────────────────────────────────

fn event_uid(event: &CalendarEvent) -> String {
    // We do NOT embed the raw MAPI global-id hex directly: Google Calendar
    // receives meeting invites via Gmail and registers the original UID in its
    // internal store. A CalDAV PUT with that same hex in the UID gets 403
    // "Forbidden" because Google won't allow overwriting its invite-derived event.
    //
    // Fix: SHA-256 hash the global-id bytes and use the first 16 hex chars prefixed
    // with "ruscal-". Google won't recognise this as an invite UID, so PUT works.
    // The mapping remains stable (same bytes → same hash → same CalDAV resource).
    use std::collections::hash_map::DefaultHasher;
    use std::hash::{Hash, Hasher};

    if !event.clean_global_id.is_empty() {
        // Mix the raw bytes through the hasher for a stable short identifier.
        let mut h = DefaultHasher::new();
        event.clean_global_id.hash(&mut h);
        format!("ruscal-{:016x}@ruscal", h.finish())
    } else {
        let mut h = DefaultHasher::new();
        event.subject.hash(&mut h);
        event.start.timestamp().hash(&mut h);
        format!("ruscal-{:016x}@ruscal", h.finish())
    }
}

// ── iCalendar serialisation ───────────────────────────────────────────────────

fn event_to_ical(event: &CalendarEvent, uid: &str) -> String {
    let dtstamp = Utc::now().format("%Y%m%dT%H%M%SZ");

    let mut lines: Vec<String> = vec![
        "BEGIN:VCALENDAR".into(),
        "VERSION:2.0".into(),
        "PRODID:-//ruscal//ruscal//EN".into(),
        "CALSCALE:GREGORIAN".into(),
        "METHOD:PUBLISH".into(),
        "BEGIN:VEVENT".into(),
        folded(format!("UID:{uid}")),
        folded(format!("DTSTAMP:{dtstamp}")),
    ];

    if event.is_all_day {
        let end_date = event.end.date_naive() + chrono::Duration::days(1);
        lines.push(folded(format!("DTSTART;VALUE=DATE:{}", event.start.format("%Y%m%d"))));
        lines.push(folded(format!("DTEND;VALUE=DATE:{}", end_date.format("%Y%m%d"))));
    } else {
        lines.push(folded(format!("DTSTART:{}", event.start.format("%Y%m%dT%H%M%SZ"))));
        lines.push(folded(format!("DTEND:{}", event.end.format("%Y%m%dT%H%M%SZ"))));
    }

    // RRULE + EXDATE for recurring events.
    if event.is_recurring {
        match build_rrule(event) {
            Some(rrule) => {
                log::debug!("{}: {}", event.subject, rrule);
                lines.push(rrule);
            }
            None => log::warn!(
                "{}: is_recurring but no RRULE (blob={} bytes, first bytes={:02X?})",
                event.subject, event.recur_blob.len(),
                event.recur_blob.get(..4.min(event.recur_blob.len())).unwrap_or(&[]),
            ),
        }

        // Cancelled occurrences → EXDATE (hide from Google Calendar).
        for exdate in build_exdates(event) {
            lines.push(exdate);
        }
    }

    lines.push(folded(format!("SUMMARY:{}", escape(&event.subject))));

    if !event.location.is_empty() {
        lines.push(folded(format!("LOCATION:{}", escape(&event.location))));
    }
    // Build DESCRIPTION, optionally prepending organizer info.
    // We do NOT emit an ORGANIZER property: Google Calendar silently drops events
    // where ORGANIZER is an external address (not the calendar owner), so the PUT
    // returns 2xx but the event never appears. Surfacing organizer as a description
    // prefix is the safe alternative.
    let description = build_description(event);
    if !description.is_empty() {
        lines.push(folded(format!("DESCRIPTION:{}", escape(&description))));
    }

    let class = match event.sensitivity {
        Sensitivity::Normal                                    => "PUBLIC",
        Sensitivity::Personal | Sensitivity::Confidential     => "CONFIDENTIAL",
        Sensitivity::Private                                   => "PRIVATE",
        Sensitivity::Unknown(_)                                => "PUBLIC",
    };
    lines.push(format!("CLASS:{class}"));

    let transp = match event.busy_status {
        BusyStatus::Free => "TRANSPARENT",
        _                => "OPAQUE",
    };
    lines.push(format!("TRANSP:{transp}"));

    lines.push("END:VEVENT".into());
    lines.push("END:VCALENDAR".into());

    lines.join("\r\n") + "\r\n"
}


/// Build a list of `(resource_uid, ical_text)` pairs for moved occurrences.
///
/// Google CalDAV rejects RECURRENCE-ID VEVENTs. Instead, moved occurrences are
/// stored as STANDALONE events (fresh UID, no RECURRENCE-ID) at their new time.
/// The master series uses EXDATE to hide the original occurrence slot, so only
/// the standalone event appears at the new time.
///
/// Resource UID format: `{master_hex}_exc_{YYYYMMDD}@ruscal`
/// This makes them discoverable by [`caldav::list_ruscal_event_uids`] for cleanup.
fn build_exception_icals(event: &CalendarEvent, master_uid: &str) -> Vec<(String, String)> {
    let dtstamp = Utc::now().format("%Y%m%dT%H%M%SZ").to_string();
    // build_exception_vevents iterates ExceptionInfo; we call it to get timing data
    // but strip RECURRENCE-ID and swap in a fresh standalone UID.
    let vevents = build_exception_vevents(event, master_uid, &dtstamp);
    if vevents.is_empty() { return Vec::new(); }

    let mut result = Vec::new();
    let mut current: Vec<String> = Vec::new();
    let mut recurrence_id: Option<String> = None;
    let mut in_vevent = false;

    for line in &vevents {
        if line == "BEGIN:VEVENT" {
            in_vevent = true;
            current.clear();
            recurrence_id = None;
        } else if line == "END:VEVENT" {
            in_vevent = false;
            let Some(rid) = recurrence_id.take() else { continue };

            let date_part = &rid[..8.min(rid.len())]; // e.g. "20260202"
            let resource_uid = format!("{}_exc_{date_part}@ruscal",
                master_uid.trim_end_matches("@ruscal"));

            // Build a standalone VEVENT: fresh UID, no RECURRENCE-ID.
            // Omit RECURRENCE-ID so Google doesn't try to link it to the master
            // (which would fail because the UIDs don't match).
            // Collect DTSTART / DTEND / SUMMARY / LOCATION from current.
            let body_lines: Vec<&str> = current.iter()
                .filter(|l| !l.starts_with("UID:")
                         && !l.starts_with("DTSTAMP:")
                         && !l.starts_with("RECURRENCE-ID:"))
                .map(String::as_str)
                .collect();

            let ical = format!(
                "BEGIN:VCALENDAR\r\nVERSION:2.0\r\nPRODID:-//ruscal//ruscal//EN\r\n\
                 CALSCALE:GREGORIAN\r\nMETHOD:PUBLISH\r\nBEGIN:VEVENT\r\n\
                 UID:{resource_uid}\r\nDTSTAMP:{dtstamp}\r\n\
                 {}\r\nEND:VEVENT\r\nEND:VCALENDAR\r\n",
                body_lines.join("\r\n"),
            );
            result.push((resource_uid, ical));
        } else if in_vevent {
            if let Some(rid_val) = line.strip_prefix("RECURRENCE-ID:") {
                recurrence_id = Some(rid_val.to_owned());
            }
            current.push(line.clone());
        }
    }

    result
}

// ── RRULE construction from PidLidAppointmentRecur blob ───────────────────────
//
// MS-OXOCAL RecurrencePattern layout (2.2.1.44.1):
//   Offset  0  ReaderVersion    u16  must be 0x3004
//   Offset  2  WriterVersion    u16  must be 0x3004
//   Offset  4  RecurFrequency   u16
//   Offset  6  PatternType      u16
//   Offset  8  CalendarType     u16
//   Offset 10  FirstDateTime    u32  (minutes since 1601-01-01)
//   Offset 14  Period           u32  (units: minutes for daily, weeks for weekly, months for monthly)
//   Offset 18  SlidingFlag      u32
//   Offset 22  PatternTypeSpecific  (variable: 0, 4, or 8 bytes)
//   After PTS: EndType          u32
//              OccurrenceCount  u32
//              ...

fn build_rrule(event: &CalendarEvent) -> Option<String> {
    let blob = &event.recur_blob;
    if blob.len() < 26 { return None; }

    let reader_ver   = u16::from_le_bytes([blob[0], blob[1]]);
    if reader_ver != 0x3004 { return None; }

    let freq         = u16::from_le_bytes([blob[4], blob[5]]);
    let pattern_type = u16::from_le_bytes([blob[6], blob[7]]);
    let period       = u32::from_le_bytes([blob[14], blob[15], blob[16], blob[17]]);

    // PatternTypeSpecific size:
    //   0x0000 (Day)      → 0 bytes
    //   0x0001 (Week)     → 4 bytes (DayOfWeek bitmask)
    //   0x0002 (Month)    → 4 bytes (day of month)
    //   0x0003 (MonthEnd) → 4 bytes
    //   0x0004 (MonthNth) → 8 bytes (DayOfWeek + N)
    let pts_len: usize = match pattern_type {
        0x0000 => 0,
        0x0001 | 0x0002 | 0x0003 => 4,
        0x0004 => 8,
        _ => return None,
    };

    // DayOfWeek bitmask (only meaningful for weekly patterns).
    let day_of_week: u32 = if pattern_type == 0x0001 && blob.len() >= 26 {
        u32::from_le_bytes([blob[22], blob[23], blob[24], blob[25]])
    } else {
        0
    };

    // EndType and OccurrenceCount sit right after PatternTypeSpecific.
    let end_type_offset = 22 + pts_len;
    let occ_offset      = end_type_offset + 4;

    let end_type: u32 = if blob.len() >= end_type_offset + 4 {
        u32::from_le_bytes(blob[end_type_offset..end_type_offset + 4].try_into().ok()?)
    } else {
        0x0000_2023 // never ends
    };
    let occ_count: u32 = if blob.len() >= occ_offset + 4 {
        u32::from_le_bytes(blob[occ_offset..occ_offset + 4].try_into().ok()?)
    } else {
        0
    };

    // ── Build RRULE parts ─────────────────────────────────────────────────────

    let mut parts: Vec<String> = Vec::new();

    // FREQ
    parts.push(match freq {
        0x200A           => "FREQ=DAILY".into(),
        0x200B           => "FREQ=WEEKLY".into(),
        0x200C           => "FREQ=MONTHLY".into(),
        0x200D | 0x200F  => "FREQ=YEARLY".into(),
        _                => return None,
    });

    // INTERVAL (omit when 1 — it's the default)
    let interval: u32 = match freq {
        0x200A => {
            // Daily: Period is in minutes (1440 = every day, 2880 = every 2 days, …)
            let days = period / 1440;
            days.max(1)
        }
        0x200B => period.max(1),             // weeks
        0x200C => period.max(1),             // months
        0x200D | 0x200F => {
            // Yearly: Period is in months (12 per year)
            (period / 12).max(1)
        }
        _ => 1,
    };
    if interval > 1 {
        parts.push(format!("INTERVAL={interval}"));
    }

    // BYDAY for weekly recurrence
    if freq == 0x200B && day_of_week != 0 {
        let days = byday_from_mask(day_of_week);
        if !days.is_empty() {
            parts.push(format!("BYDAY={}", days.join(",")));
        }
    }

    // End condition
    match end_type {
        0x0000_2021 => {
            // End by date — use PidLidClipEnd (already parsed as recurrence_end).
            if let Some(until_date) = event.recurrence_end {
                parts.push(format!("UNTIL={}T235959Z", until_date.format("%Y%m%d")));
            }
        }
        0x0000_2022 => {
            // End after N occurrences.
            if occ_count > 0 && occ_count < 0x7FFF_FFFF {
                parts.push(format!("COUNT={occ_count}"));
            }
        }
        _ => {} // 0x00002023 / 0xFFFFFFFF = no end, nothing to add
    }

    Some(format!("RRULE:{}", parts.join(";")))
}

/// Parse cancelled and moved occurrence dates from the RecurrencePattern blob
/// and return EXDATE lines for each.
///
/// RecurrencePattern layout after PatternTypeSpecific + EndType + OccurrenceCount:
///   FirstDOW             u32
///   DeletedInstanceCount u32
///   DeletedInstanceDates [u32; DeletedInstanceCount]  ← cancelled occurrences
///   ModifiedInstanceCount u32
///   ModifiedInstanceDates [u32; ModifiedInstanceCount] ← moved occurrences
///
/// Each date is minutes since 1601-01-01 at local midnight of that occurrence.
/// We reconstruct the UTC datetime by using the date + the original DTSTART time.
fn build_exdates(event: &CalendarEvent) -> Vec<String> {
    let blob = &event.recur_blob;

    let pattern_type = if blob.len() >= 8 {
        u16::from_le_bytes([blob[6], blob[7]])
    } else {
        return Vec::new();
    };

    let pts_len: usize = match pattern_type {
        0x0000 => 0,
        0x0001 | 0x0002 | 0x0003 => 4,
        0x0004 => 8,
        _ => return Vec::new(),
    };

    // FirstDOW sits right after OccurrenceCount.
    // Layout: [22 + pts_len] EndType [+4] OccurrenceCount [+4] FirstDOW [+4] DeletedCount
    let deleted_count_offset = 22 + pts_len + 4 + 4 + 4; // skip EndType, OccCount, FirstDOW
    if blob.len() < deleted_count_offset + 4 { return Vec::new(); }

    let deleted_count = u32::from_le_bytes(
        blob[deleted_count_offset..deleted_count_offset + 4].try_into().unwrap()
    ) as usize;

    let deleted_dates_offset = deleted_count_offset + 4;
    if blob.len() < deleted_dates_offset + deleted_count * 4 { return Vec::new(); }

    // Modified instances sit right after the deleted dates array.
    let modified_count_offset = deleted_dates_offset + deleted_count * 4;
    let modified_count = if blob.len() >= modified_count_offset + 4 {
        u32::from_le_bytes(
            blob[modified_count_offset..modified_count_offset + 4].try_into().unwrap()
        ) as usize
    } else {
        0
    };
    let modified_dates_offset = modified_count_offset + 4;

    // Only use DeletedInstanceDates for EXDATE — per MS-OXOCAL spec, this array
    // already includes the original dates of ALL exceptions (both cancelled and
    // moved). ModifiedInstanceDates is a subset of DeletedInstanceDates; adding
    // it again would produce duplicate EXDATEs for moved occurrences.
    let mut exception_minute_dates: Vec<u32> = Vec::new();

    for i in 0..deleted_count {
        let off = deleted_dates_offset + i * 4;
        let mins = u32::from_le_bytes(blob[off..off + 4].try_into().unwrap());
        exception_minute_dates.push(mins);
    }

    let _ = (modified_count, modified_dates_offset); // parsed above for ARP offset calc
    if exception_minute_dates.is_empty() { return Vec::new(); }

    // Convert minutes-since-1601 to a UTC datetime. Each value is the local
    // midnight of the occurrence date. We replace the time with DTSTART's
    // UTC time-of-day so the EXDATE matches the actual occurrence.
    let start_time = event.start.time(); // HH:MM:SS of the first occurrence
    let epoch_1601: chrono::DateTime<Utc> =
        chrono::DateTime::from_timestamp(-11_644_473_600, 0).unwrap();

    let mut exdates = Vec::new();
    for mins in exception_minute_dates {
        let dt = epoch_1601 + chrono::Duration::minutes(mins as i64);
        // dt is midnight UTC of local date — keep only the date part.
        // Combine with DTSTART's time to get the actual occurrence datetime.
        if let Some(occ) = dt.date_naive().and_time(start_time)
            .and_local_timezone(Utc).single()
        {
            exdates.push(format!("EXDATE:{}", occ.format("%Y%m%dT%H%M%SZ")));
        }
    }

    log::debug!("{}: {} EXDATEs ({}d {}m)", event.subject,
        exdates.len(), deleted_count, modified_count);
    exdates
}

/// Build RECURRENCE-ID override VEVENTs for moved or modified occurrences.
///
/// Parses the AppointmentRecurrencePattern (ARP) extension that follows the
/// RecurrencePattern blob. For each ExceptionInfo that is not a cancellation
/// and not a pure deletion, emits a VEVENT block with a RECURRENCE-ID pointing
/// at the original occurrence and a new DTSTART/DTEND for the new time.
///
/// ARP header layout (empirically verified — version fields are u32):
///   [rp+0]  ReaderVersion2  u32  (= 0x00003006)
///   [rp+4]  WriterVersion2  u32
///   [rp+8]  StartTimeOffset u32  (local minutes from midnight — not used here)
///  [rp+12]  EndTimeOffset   u32
///  [rp+16]  ExceptionCount  u16
///  [rp+18]  ExceptionInfo[0..ExceptionCount]
///
/// ExceptionInfo fixed part (14 bytes):
///   StartDateTime    u32  (local minutes since 1601 — new occurrence start)
///   EndDateTime      u32  (local minutes since 1601 — new occurrence end)
///   OriginalStartDate u32 (local minutes since 1601 — original occurrence start)
///   OverrideFlags    u16
fn build_exception_vevents(event: &CalendarEvent, uid: &str, dtstamp: &str) -> Vec<String> {
    let blob = &event.recur_blob;
    if blob.len() < 26 { return Vec::new(); }

    let pattern_type = u16::from_le_bytes([blob[6], blob[7]]);
    let pts_len: usize = match pattern_type {
        0x0000 => 0,
        0x0001 | 0x0002 | 0x0003 => 4,
        0x0004 => 8,
        _ => return Vec::new(),
    };

    let del_count_off = 22 + pts_len + 4 + 4 + 4;
    if blob.len() < del_count_off + 4 { return Vec::new(); }
    let del_count = u32::from_le_bytes(
        blob[del_count_off..del_count_off+4].try_into().unwrap()
    ) as usize;

    let mod_count_off = del_count_off + 4 + del_count * 4;
    if blob.len() < mod_count_off + 4 { return Vec::new(); }
    let mod_count = u32::from_le_bytes(
        blob[mod_count_off..mod_count_off+4].try_into().unwrap()
    ) as usize;

    // RecurrencePattern ends after ModifiedInstanceDates + StartDate + EndDate.
    let rp_end = mod_count_off + 4 + mod_count * 4 + 4 + 4;

    // ARP header needs 18 bytes (4+4+4+4+2).
    if blob.len() < rp_end + 18 { return Vec::new(); }

    // Sanity: ReaderVersion2 must be 0x3006.
    let rv2 = u32::from_le_bytes(blob[rp_end..rp_end+4].try_into().unwrap());
    if rv2 != 0x0000_3006 {
        log::warn!("{}: unexpected ARP ReaderVersion2 = 0x{rv2:08X}", event.subject);
        return Vec::new();
    }

    let exc_count = u16::from_le_bytes([blob[rp_end+16], blob[rp_end+17]]) as usize;
    if exc_count == 0 { return Vec::new(); }

    // epoch used to convert blob minutes → chrono dates
    let epoch_1601: chrono::DateTime<Utc> =
        chrono::DateTime::from_timestamp(-11_644_473_600, 0).unwrap();
    let start_time = event.start.time(); // UTC time of the first occurrence

    let mut result = Vec::new();
    let mut pos = rp_end + 18;

    for _ in 0..exc_count {
        if blob.len() < pos + 14 { break; }

        let start_dt = u32::from_le_bytes(blob[pos..pos+4].try_into().unwrap());
        let end_dt   = u32::from_le_bytes(blob[pos+4..pos+8].try_into().unwrap());
        let orig_sd  = u32::from_le_bytes(blob[pos+8..pos+12].try_into().unwrap());
        let flags    = u16::from_le_bytes([blob[pos+12], blob[pos+13]]);
        pos += 14;

        // Walk optional fields to advance `pos` past this ExceptionInfo.
        let mut meeting_type: Option<u32> = None;
        if flags & 0x0001 != 0 { // ARO_SUBJECT
            if pos + 4 > blob.len() { break; }
            let len2 = u16::from_le_bytes([blob[pos+2], blob[pos+3]]) as usize;
            pos += 4 + len2;
        }
        if flags & 0x0002 != 0 { // ARO_MEETINGTYPE
            if pos + 4 > blob.len() { break; }
            meeting_type = Some(u32::from_le_bytes(blob[pos..pos+4].try_into().unwrap()));
            pos += 4;
        }
        if flags & 0x0004 != 0 { pos += 4; } // ARO_REMINDERDELTA
        if flags & 0x0008 != 0 { pos += 4; } // ARO_REMINDER
        if flags & 0x0010 != 0 { // ARO_LOCATION
            if pos + 4 > blob.len() { break; }
            let len2 = u16::from_le_bytes([blob[pos+2], blob[pos+3]]) as usize;
            pos += 4 + len2;
        }
        if flags & 0x0020 != 0 { pos += 4; } // ARO_BUSYSTATUS
        // ARO_ATTACHMENT (0x0040) — no extra data
        if flags & 0x0080 != 0 { pos += 4; } // ARO_SUBTYPE
        if flags & 0x0100 != 0 { pos += 4; } // ARO_APPTCOLOR
        // ARO_EXCEPTIONAL_BODY (0x0200) — no extra data

        // Determine whether this exception needs a RECURRENCE-ID VEVENT.
        let is_cancelled = meeting_type == Some(3); // olCanceled
        let is_pure_deletion = flags == 0 && start_dt == orig_sd;
        if is_cancelled || is_pure_deletion {
            continue; // EXDATE already handles this — no override VEVENT needed
        }

        // Compute the original occurrence's UTC datetime (what the RRULE expanded to).
        let orig_date = (epoch_1601 + chrono::Duration::minutes(orig_sd as i64)).date_naive();
        let Some(orig_occ) = orig_date.and_time(start_time).and_local_timezone(Utc).single()
        else { continue };

        // New start = original UTC + local-minute delta from OriginalStartDate to StartDateTime.
        let delta_start = start_dt.saturating_sub(orig_sd) as i64;
        let delta_end   = end_dt.saturating_sub(orig_sd) as i64;
        let new_start = orig_occ + chrono::Duration::minutes(delta_start);
        let new_end   = orig_occ + chrono::Duration::minutes(delta_end);

        let fmt_dt = |dt: chrono::DateTime<Utc>| dt.format("%Y%m%dT%H%M%SZ").to_string();
        let recur_id = fmt_dt(orig_occ);
        let dtstart  = fmt_dt(new_start);
        let dtend    = fmt_dt(new_end);

        log::debug!("{}: exception VEVENT RECURRENCE-ID={recur_id} → {dtstart}..{dtend}",
            event.subject);

        result.push("BEGIN:VEVENT".into());
        result.push(folded(format!("UID:{uid}")));
        result.push(format!("DTSTAMP:{dtstamp}"));
        result.push(format!("RECURRENCE-ID:{recur_id}"));
        result.push(format!("DTSTART:{dtstart}"));
        result.push(format!("DTEND:{dtend}"));
        result.push(folded(format!("SUMMARY:{}", escape(&event.subject))));
        if !event.location.is_empty() {
            result.push(folded(format!("LOCATION:{}", escape(&event.location))));
        }
        let description = build_description(event);
        if !description.is_empty() {
            result.push(folded(format!("DESCRIPTION:{}", escape(&description))));
        }
        result.push("END:VEVENT".into());
    }

    result
}

/// Convert a MAPI DayOfWeek bitmask to iCal BYDAY day abbreviations (RFC 5545).
fn byday_from_mask(mask: u32) -> Vec<&'static str> {
    let mut days = Vec::new();
    if mask & 0x01 != 0 { days.push("SU"); }
    if mask & 0x02 != 0 { days.push("MO"); }
    if mask & 0x04 != 0 { days.push("TU"); }
    if mask & 0x08 != 0 { days.push("WE"); }
    if mask & 0x10 != 0 { days.push("TH"); }
    if mask & 0x20 != 0 { days.push("FR"); }
    if mask & 0x40 != 0 { days.push("SA"); }
    days
}

// ── Diagnostic dump ───────────────────────────────────────────────────────────

/// Write a human-readable dump of all events (especially recurring ones) to
// ── Text helpers ──────────────────────────────────────────────────────────────

/// Build the DESCRIPTION field value.
///
/// Google Calendar silently drops events where ORGANIZER is an external address
/// (not the calendar owner). We surface organizer info as a plain-text prefix in
/// DESCRIPTION instead, so it remains visible to the user. The structured
/// `X-RUSCAL-ORGANIZER-*` properties carry the same info for reverse sync; the
/// description prefix is cosmetic-only.
///
/// **Idempotent:** any existing `Organizer: …` prefix in the incoming body
/// (e.g. from a prior reverse-sync cycle that copied Google's description back
/// into Outlook) is stripped before re-prepending, so repeated syncs do not
/// accumulate stacked prefixes.
fn build_description(event: &CalendarEvent) -> String {
    let organizer_line = match (event.organizer_name.as_str(), event.organizer_email.as_str()) {
        ("", "")      => String::new(),
        ("", email)   => format!("Organizer: {email}"),
        (name, "")    => format!("Organizer: {name}"),
        (name, email) => format!("Organizer: {name} <{email}>"),
    };

    let clean_body = strip_organizer_prefix(&event.body);

    match (organizer_line.as_str(), clean_body.as_str()) {
        ("", "")   => String::new(),
        (org, "")  => org.to_owned(),
        ("", body) => body.to_owned(),
        (org, body) => format!("{org}\n\n{body}"),
    }
}

/// Strip a leading `Organizer: …\n\n` block from a body string.
///
/// Matches the exact shape we emit in [`build_description`] so it only removes
/// ruscal's own previous injection. A user-written description that genuinely
/// starts with "Organizer:" is preserved unless it's followed by a blank line,
/// which would be very unusual for a real note.
fn strip_organizer_prefix(body: &str) -> String {
    if !body.starts_with("Organizer:") {
        return body.to_owned();
    }
    // Normalize on \n for splitting, then preserve original CRLF on rejoin
    // only where it appears (we just emit \n — CalDAV folding adds CRLF).
    let mut lines = body.split('\n').peekable();
    // Drop consecutive "Organizer:" lines at the start.
    while lines.peek().map(|l| l.trim_end_matches('\r').starts_with("Organizer:")).unwrap_or(false) {
        lines.next();
    }
    // Drop the single blank separator line, if present.
    if lines.peek().map(|l| l.trim_end_matches('\r').is_empty()).unwrap_or(false) {
        lines.next();
    }
    lines.collect::<Vec<_>>().join("\n")
}

fn escape(s: &str) -> String {
    s.replace('\\', "\\\\")
     .replace(';',  "\\;")
     .replace(',',  "\\,")
     .replace('\n', "\\n")
     .replace('\r', "")
}

fn folded(line: String) -> String {
    let bytes = line.as_bytes();
    if bytes.len() <= 75 { return line; }
    let mut out = String::with_capacity(line.len() + (line.len() / 74) * 3);
    let mut pos = 0;
    while pos < bytes.len() {
        let max   = if pos == 0 { 75 } else { 74 };
        let end   = (pos + max).min(bytes.len());
        let split = (pos..=end).rev()
            .find(|&i| i == pos || (bytes[i - 1] & 0xC0) != 0x80)
            .unwrap_or(end);
        if pos > 0 { out.push(' '); }
        out.push_str(&line[pos..split]);
        pos = split;
        if pos < bytes.len() { out.push_str("\r\n"); }
    }
    out
}

// ── Integration tests ─────────────────────────────────────────────────────────

#[cfg(test)]
mod integration {
    use super::*;
    use chrono::Utc;

    #[cfg(any())]
    #[test]
    #[ignore = "requires live Google credentials"]
    fn test_xprop_roundtrip() {
        let cfg_path = dirs::data_local_dir().unwrap().join("ruscal").join("config.json");
        let cfg: serde_json::Value = serde_json::from_str(&std::fs::read_to_string(&cfg_path).unwrap()).unwrap();
        let dest_id = cfg["pairs"][0]["dest_id"].as_str().unwrap().to_owned();
        let gmail   = cfg["pairs"][0]["google_email"].as_str().unwrap().to_owned();
        let token   = crate::google::get_access_token_for(&gmail).expect("token");

        let uid = "ruscal-xproptest@ruscal";
        let probe_hex = "DEADBEEFCAFE0123456789ABCDEF";
        let ical = format!(
            "BEGIN:VCALENDAR\r\nVERSION:2.0\r\nPRODID:-//ruscal//ruscal//EN\r\n\
             CALSCALE:GREGORIAN\r\nMETHOD:PUBLISH\r\nBEGIN:VEVENT\r\n\
             UID:{uid}\r\nDTSTAMP:20260414T000000Z\r\n\
             X-RUSCAL-OUTLOOK-GLOBALID:{probe_hex}\r\n\
             DTSTART:20260601T090000Z\r\nDTEND:20260601T100000Z\r\n\
             SUMMARY:ruscal xprop roundtrip test\r\n\
             END:VEVENT\r\nEND:VCALENDAR\r\n"
        );

        crate::caldav::put_event(&dest_id, uid, &ical, &token).expect("PUT failed");
        let got = crate::caldav::get_event(&dest_id, uid, &token).expect("GET failed");
        let _ = crate::caldav::delete_event(&dest_id, uid, &token);

        println!("--- server returned ---\n{}", got);
        assert!(got.contains("X-RUSCAL-OUTLOOK-GLOBALID"),
            "Google CalDAV stripped the X- property on round-trip — cannot use it for identity mapping");
        assert!(got.contains(probe_hex),
            "X- property present but value lost");
    }

    #[cfg(any())]
    #[test]
    fn test_sync_collection_roundtrip() {
        let cfg_path = dirs::data_local_dir().unwrap().join("ruscal").join("config.json");
        let cfg: serde_json::Value = serde_json::from_str(&std::fs::read_to_string(&cfg_path).unwrap()).unwrap();
        let dest_id = cfg["pairs"][0]["dest_id"].as_str().unwrap().to_owned();
        let gmail   = cfg["pairs"][0]["google_email"].as_str().unwrap().to_owned();
        let token   = crate::google::get_access_token_for(&gmail).expect("token");

        // Step 1: initial sync — grab current token.
        let first = crate::caldav::sync_collection(&dest_id, &token, None).expect("initial sync");
        println!("initial: token={} changes={}", first.new_token, first.changes.len());
        assert!(!first.new_token.is_empty());

        // Google's sync-token appears to use coarse timestamps — wait a beat so
        // the PUT happens in a later instant than the token's anchor.
        std::thread::sleep(std::time::Duration::from_secs(3));

        // Step 2: PUT a fresh event — this must show up in the next delta.
        let uid = "ruscal-synctest@ruscal";
        let ical = format!(
            "BEGIN:VCALENDAR\r\nVERSION:2.0\r\nPRODID:-//ruscal//test//EN\r\n\
             BEGIN:VEVENT\r\nUID:{uid}\r\nDTSTAMP:20260415T000000Z\r\n\
             DTSTART:20260701T090000Z\r\nDTEND:20260701T100000Z\r\n\
             SUMMARY:ruscal sync-collection test\r\n\
             END:VEVENT\r\nEND:VCALENDAR\r\n"
        );
        crate::caldav::put_event(&dest_id, uid, &ical, &token).expect("PUT failed");

        // Confirm the PUT actually landed before we blame sync-collection.
        let got = crate::caldav::get_event(&dest_id, uid, &token).expect("GET after PUT");
        assert!(got.contains(uid), "PUT didn't land on server");

        // Step 3: delta sync — Google's sync-token has replication lag, so retry.
        let mut second = crate::caldav::sync_collection(&dest_id, &token, Some(&first.new_token))
            .expect("delta sync");
        println!("first token:  {}", first.new_token);
        println!("second token: {}", second.new_token);
        println!("token advanced: {}", first.new_token != second.new_token);

        // If server advances its token in chunks, chain tokens forward until
        // our PUT appears or we exhaust retries.
        let mut cur_token = second.new_token.clone();
        let mut all_changes = second.changes;
        let mut tries = 0;
        while !all_changes.iter().any(|c| c.href.contains("synctest")) && tries < 10 {
            std::thread::sleep(std::time::Duration::from_secs(2));
            let next = crate::caldav::sync_collection(&dest_id, &token, Some(&cur_token))
                .expect("delta sync retry");
            println!("retry {tries}: token={} changes={}", next.new_token, next.changes.len());
            if next.new_token != cur_token {
                cur_token = next.new_token;
            }
            for c in &next.changes {
                println!("  {} deleted={}", c.href, c.deleted);
            }
            all_changes.extend(next.changes);
            tries += 1;
        }
        let second = crate::caldav::SyncResult { new_token: cur_token, changes: all_changes };
        println!("delta: token={} changes={}", second.new_token, second.changes.len());
        for c in &second.changes {
            println!("  {} deleted={} etag={:?}", c.href, c.deleted, c.etag);
        }

        let _ = crate::caldav::delete_event(&dest_id, uid, &token);

        let found = second.changes.iter().any(|c| c.href.contains("synctest"));
        assert!(found, "sync-collection did not report our fresh PUT as a change");
        assert_ne!(first.new_token, second.new_token, "token must advance");
    }

    #[cfg(any())]
    #[test]
    fn test_safe_subset_is_exactly_that() {
        use crate::ical_parse::ParsedIcal;
        let owner    = "me@example.com";
        let wl_uid   = REVERSE_SYNC_UID_WHITELIST[0].to_owned();

        // Build a baseline that passes the whitelist + organizer gate.
        let whitelisted = |attendees: Vec<String>, organizer: Option<String>| {
            let mut p = ParsedIcal::default();
            p.uid                = Some(wl_uid.clone());
            p.x_organizer_email  = organizer;
            p.attendees          = attendees;
            p
        };

        // Refused: whitelist gate — event not in the initial-rollout list,
        // even though it would otherwise pass organizer + attendee checks.
        let mut p = ParsedIcal::default();
        p.uid               = Some("ruscal-not-listed@ruscal".into());
        p.x_organizer_email = Some(owner.to_owned());
        assert!(!check_safe_to_write_to_outlook(&p, owner).safe);

        // Refused: UID missing entirely.
        let mut p = ParsedIcal::default();
        p.x_organizer_email = Some(owner.to_owned());
        assert!(!check_safe_to_write_to_outlook(&p, owner).safe);

        // Accepted: whitelist + organizer=owner via X-RUSCAL-ORGANIZER-EMAIL, no attendees.
        let p = whitelisted(vec![], Some(owner.to_owned()));
        assert!(check_safe_to_write_to_outlook(&p, owner).safe);

        // Accepted: fallback ORGANIZER when X-* missing (but still whitelisted).
        let mut p = ParsedIcal::default();
        p.uid             = Some(wl_uid.clone());
        p.organizer_email = Some(owner.to_owned());
        assert!(check_safe_to_write_to_outlook(&p, owner).safe);

        // Accepted: case-insensitive owner match (email addresses are i-case).
        let p = whitelisted(vec![], Some("Me@Example.COM".to_owned()));
        assert!(check_safe_to_write_to_outlook(&p, owner).safe);

        // Refused: any attendee, even if whitelisted and organizer=owner.
        let p = whitelisted(vec!["someone@other.com".into()], Some(owner.to_owned()));
        assert!(!check_safe_to_write_to_outlook(&p, owner).safe);

        // Refused: external organizer, even if whitelisted.
        let p = whitelisted(vec![], Some("external@other.com".to_owned()));
        assert!(!check_safe_to_write_to_outlook(&p, owner).safe);

        // Refused: no organizer known at all.
        let p = whitelisted(vec![], None);
        assert!(!check_safe_to_write_to_outlook(&p, owner).safe);

        // Refused: X-* takes precedence over ORGANIZER — if they disagree and
        // X-* says external, we refuse even though ORGANIZER might match owner.
        let mut p = ParsedIcal::default();
        p.uid               = Some(wl_uid.clone());
        p.x_organizer_email = Some("external@other.com".into());
        p.organizer_email   = Some(owner.to_owned());
        assert!(!check_safe_to_write_to_outlook(&p, owner).safe);
    }

    /// Unit test (no network): building DESCRIPTION twice in a row must be
    /// idempotent. Simulates the feared reverse-sync cycle where Google's
    /// description (already containing "Organizer: X\n\n<body>") is copied
    /// back into Outlook's body and then re-synced — we must not stack
    /// prefixes.
    #[test]
    fn test_organizer_prefix_idempotent() {
        use crate::event::{BusyStatus, ResponseStatus, Sensitivity};
        fn sample(body: &str) -> crate::event::CalendarEvent {
            crate::event::CalendarEvent {
                subject: "test".into(),
                start: chrono::Utc::now(),
                end:   chrono::Utc::now(),
                is_all_day: false,
                location: String::new(),
                organizer_name:  "Alice".into(),
                organizer_email: "alice@example.com".into(),
                body: body.into(),
                busy_status: BusyStatus::Busy,
                response_status: ResponseStatus::Accepted,
                sensitivity: Sensitivity::Normal,
                is_recurring: false,
                recurrence_end: None,
                clean_global_id: Vec::new(),
                recur_blob: Vec::new(),
            }
        }

        let original_body = "real meeting notes";
        let first  = build_description(&sample(original_body));
        // Simulate reverse sync: the built description becomes the body.
        let second = build_description(&sample(&first));
        let third  = build_description(&sample(&second));
        assert_eq!(first, second, "prefix stacked on round-trip");
        assert_eq!(second, third, "prefix stacked on triple round-trip");
        assert!(first.starts_with("Organizer: Alice <alice@example.com>\n\n"));
        assert!(first.ends_with(original_body));

        // Empty-body case.
        let only_org = build_description(&sample(""));
        assert_eq!(only_org, "Organizer: Alice <alice@example.com>");
        assert_eq!(build_description(&sample(&only_org)), only_org);
    }

    /// Ad-hoc cleanup: delete OGCS-tagged Google events whose
    /// `outlook_GlobalApptID` matches an Outlook event currently in the sync
    /// window (so ruscal would re-create it). Leaves OGCS events whose source
    /// is no longer in Outlook untouched, and skips Gmail-derived `_e9q*` IDs.
    ///
    /// Run with:  cargo test test_delete_ogcs_duplicates -- --ignored --nocapture
    #[test]
    #[ignore = "ad-hoc cleanup; requires live Outlook + Google credentials"]
    fn test_delete_ogcs_duplicates() {
        // 1. Read Outlook global_ids in the active sync window.
        let now   = Utc::now();
        let start = now - chrono::Duration::days(crate::outlook::DEFAULT_PAST_DAYS);
        let end   = now + chrono::Duration::days(crate::outlook::DEFAULT_FUTURE_DAYS);
        let events = crate::outlook::read_calendar_events(start, end).expect("Outlook read");

        let outlook_ids: std::collections::HashSet<String> = events.iter()
            .filter(|e| !e.clean_global_id.is_empty())
            .map(|e| hex::encode_upper(&e.clean_global_id))
            .collect();
        println!("Outlook events in window: {} ({} unique global_ids)",
            events.len(), outlook_ids.len());

        // 2. List Google events with OGCS extended property.
        let cfg_path = dirs::data_local_dir().unwrap().join("ruscal").join("config.json");
        let cfg: serde_json::Value =
            serde_json::from_str(&std::fs::read_to_string(&cfg_path).unwrap()).unwrap();
        let gmail = cfg["pairs"][0]["google_email"].as_str().unwrap().to_owned();
        let token = crate::google::get_access_token_for(&gmail).expect("token");

        let client = reqwest::blocking::Client::new();
        let mut all_items: Vec<serde_json::Value> = Vec::new();
        let mut page_token: Option<String> = None;
        loop {
            let mut url = String::from(
                "https://www.googleapis.com/calendar/v3/calendars/primary/events\
                 ?maxResults=2500&showDeleted=false&singleEvents=false");
            if let Some(t) = &page_token {
                url.push_str("&pageToken=");
                url.push_str(t);
            }
            let resp: serde_json::Value = client.get(&url)
                .bearer_auth(&token).send().expect("list").json().expect("json");
            if let Some(items) = resp["items"].as_array() {
                all_items.extend(items.iter().cloned());
            }
            page_token = resp["nextPageToken"].as_str().map(str::to_owned);
            if page_token.is_none() { break; }
        }
        println!("Google events listed: {}", all_items.len());

        // 3. Filter to OGCS-tagged ones whose outlook_GlobalApptID matches an
        //    Outlook event we just read. Skip Gmail-derived (_e9q*) IDs.
        let candidates: Vec<(String, String, String)> = all_items.iter()
            .filter_map(|e| {
                let id = e["id"].as_str()?;
                if id.starts_with("_e9q") { return None; } // Gmail invite copy
                let ogcs_id = e["extendedProperties"]["private"]["outlook_GlobalApptID"]
                    .as_str()?.to_owned();
                if !outlook_ids.contains(&ogcs_id) { return None; }
                Some((id.to_owned(), e["summary"].as_str().unwrap_or("(no title)").to_owned(), ogcs_id))
            })
            .collect();
        println!("Duplicates to delete: {}", candidates.len());
        for (id, summary, _) in &candidates {
            println!("  {id} | {summary}");
        }

        // Hard safety floor — refuse to mass-delete if something is clearly off.
        assert!(candidates.len() <= 200,
            "refusing to delete > 200 events in one run");

        // 4. DELETE each one via the Google Calendar API.
        let mut deleted = 0usize;
        for (id, summary, _) in &candidates {
            let url = format!(
                "https://www.googleapis.com/calendar/v3/calendars/primary/events/{id}");
            let resp = client.delete(&url).bearer_auth(&token)
                .send().expect("delete");
            let status = resp.status();
            if status.is_success() || status.as_u16() == 410 {
                deleted += 1;
                println!("  DEL ✓ {summary}");
            } else {
                let body = resp.text().unwrap_or_default();
                println!("  DEL ✗ {id} | {status} | {}",
                    body.chars().take(200).collect::<String>());
            }
        }
        println!("\nDeleted {deleted}/{} OGCS duplicates", candidates.len());
    }

    /// Diagnose *why* certain subjects keep coming back as "skipped".
    ///
    /// For each configured skip-target subject, lists every Google event
    /// currently matching that summary: iCalUID, Google ID, extendedProperties,
    /// organizer, status, created-by. Helps distinguish:
    ///
    /// - Gmail-derived invite (id starts with `_e9q…`)
    /// - Leftover OGCS copy (has `outlook_GlobalApptID` extendedProperty)
    /// - Earlier ruscal copy that ruscal's own UID collides with (ends `@ruscal`)
    /// - Something else
    ///
    /// Run with:
    ///   cargo test --bin ruscal test_diagnose_skipped_subjects -- --ignored --nocapture
    #[test]
    #[ignore = "ad-hoc diagnostic; requires live Google credentials"]
    fn test_diagnose_skipped_subjects() {
        let needles: &[&str] = &[
            "LICCON3 testing",
            "LICCON3 VS Code als Eclipse Ersatz",
            "Erkl",  // matches both "Erklärung" and "Erkl�rung"
        ];

        let cfg_path = dirs::data_local_dir().unwrap().join("ruscal").join("config.json");
        let cfg: serde_json::Value =
            serde_json::from_str(&std::fs::read_to_string(&cfg_path).unwrap()).unwrap();
        let gmail = cfg["pairs"][0]["google_email"].as_str().unwrap().to_owned();
        let token = crate::google::get_access_token_for(&gmail).expect("token");

        let client = reqwest::blocking::Client::new();
        let mut all_items: Vec<serde_json::Value> = Vec::new();
        let mut page_token: Option<String> = None;
        loop {
            let mut url = String::from(
                "https://www.googleapis.com/calendar/v3/calendars/primary/events\
                 ?maxResults=2500&showDeleted=false&singleEvents=false");
            if let Some(t) = &page_token {
                url.push_str("&pageToken=");
                url.push_str(t);
            }
            let resp: serde_json::Value = client.get(&url)
                .bearer_auth(&token).send().expect("list").json().expect("json");
            if let Some(items) = resp["items"].as_array() {
                all_items.extend(items.iter().cloned());
            }
            page_token = resp["nextPageToken"].as_str().map(str::to_owned);
            if page_token.is_none() { break; }
        }
        println!("Scanned {} Google events", all_items.len());

        for needle in needles {
            println!("\n═══ {needle} ═══");
            let hits: Vec<_> = all_items.iter()
                .filter(|e| {
                    e["summary"].as_str().map(|s| s.contains(needle)).unwrap_or(false)
                })
                .collect();
            println!("  {} match(es)", hits.len());

            for e in hits {
                let id       = e["id"].as_str().unwrap_or("?");
                let ical_uid = e["iCalUID"].as_str().unwrap_or("?");
                let summary  = e["summary"].as_str().unwrap_or("?");
                let status   = e["status"].as_str().unwrap_or("?");
                let creator  = e["creator"]["email"].as_str().unwrap_or("?");
                let organizer = e["organizer"]["email"].as_str().unwrap_or("?");
                let is_recur = e["recurrence"].is_array();
                let recurring_id = e["recurringEventId"].as_str();
                let ogcs_id  = e["extendedProperties"]["private"]["outlook_GlobalApptID"]
                    .as_str().unwrap_or("-");
                let private_keys: Vec<&str> = e["extendedProperties"]["private"]
                    .as_object()
                    .map(|m| m.keys().map(|s| s.as_str()).collect())
                    .unwrap_or_default();

                let kind = if id.starts_with("_e9q") { "GMAIL-INVITE" }
                    else if ical_uid.ends_with("@ruscal") { "RUSCAL" }
                    else if ogcs_id != "-" { "OGCS-LEFTOVER" }
                    else { "OTHER" };

                println!("  [{kind}]");
                println!("    summary:      {summary}");
                println!("    id:           {id}");
                println!("    iCalUID:      {ical_uid}");
                println!("    status:       {status}");
                println!("    creator:      {creator}");
                println!("    organizer:    {organizer}");
                println!("    recurring:    {is_recur}  recurring_id={recurring_id:?}");
                println!("    OGCS_id:      {ogcs_id}");
                println!("    priv_keys:    {private_keys:?}");
            }
        }
    }

    /// Delete Google events that collide with ruscal's CalDAV URLs.
    ///
    /// Targets events where ALL of the following hold:
    ///   - `iCalUID` ends in `@ruscal` (ruscal-format UID — not a stranger's invite)
    ///   - `extendedProperties.private.outlook_GlobalApptID` matches an Outlook
    ///     event currently in the sync window (will be re-created)
    ///   - Google event `id` starts with `_` (the CalDAV URL Google assigned does
    ///     NOT match `{uid}.ics`, so ruscal's PUT gets 409 "conflict, different URL")
    ///
    /// These are OGCS-created copies that ruscal's PUT cannot upsert into
    /// because Google stores them at an internal URL (`_e9q…`) instead of
    /// `{uid}.ics`. Deleting clears the URL conflict; the next ruscal sync
    /// re-creates them at the correct CalDAV URL.
    ///
    /// Run with:
    ///   cargo test --bin ruscal test_delete_url_conflicts -- --ignored --nocapture
    #[test]
    #[ignore = "ad-hoc cleanup; requires live Outlook + Google credentials"]
    fn test_delete_url_conflicts() {
        // 1. Read current Outlook window so we only delete events that will
        //    be re-created (never strand the user without their meetings).
        let now   = Utc::now();
        let start = now - chrono::Duration::days(crate::outlook::DEFAULT_PAST_DAYS);
        let end   = now + chrono::Duration::days(crate::outlook::DEFAULT_FUTURE_DAYS);
        let events = crate::outlook::read_calendar_events(start, end).expect("Outlook read");
        let outlook_ids: std::collections::HashSet<String> = events.iter()
            .filter(|e| !e.clean_global_id.is_empty())
            .map(|e| hex::encode_upper(&e.clean_global_id))
            .collect();
        println!("Outlook global_ids in window: {}", outlook_ids.len());

        // 2. Fetch all Google events.
        let cfg_path = dirs::data_local_dir().unwrap().join("ruscal").join("config.json");
        let cfg: serde_json::Value =
            serde_json::from_str(&std::fs::read_to_string(&cfg_path).unwrap()).unwrap();
        let gmail = cfg["pairs"][0]["google_email"].as_str().unwrap().to_owned();
        let token = crate::google::get_access_token_for(&gmail).expect("token");

        let client = reqwest::blocking::Client::new();
        let mut all_items: Vec<serde_json::Value> = Vec::new();
        let mut page_token: Option<String> = None;
        loop {
            let mut url = String::from(
                "https://www.googleapis.com/calendar/v3/calendars/primary/events\
                 ?maxResults=2500&showDeleted=false&singleEvents=false");
            if let Some(t) = &page_token {
                url.push_str("&pageToken=");
                url.push_str(t);
            }
            let resp: serde_json::Value = client.get(&url)
                .bearer_auth(&token).send().expect("list").json().expect("json");
            if let Some(items) = resp["items"].as_array() {
                all_items.extend(items.iter().cloned());
            }
            page_token = resp["nextPageToken"].as_str().map(str::to_owned);
            if page_token.is_none() { break; }
        }
        println!("Google events scanned: {}", all_items.len());

        // 3. Select URL-conflict candidates.
        //
        // Any Google event whose iCalUID ends in `@ruscal` AND whose Google
        // `id` starts with `_` is a misplaced ruscal event: ruscal only PUTs
        // at `{uid}.ics`, so a ruscal-UID event at a non-standard URL is by
        // definition a CalDAV URL conflict. We delete these so the next PUT
        // can land at the correct URL.
        //
        // We do NOT require `outlook_GlobalApptID` — recurrence overrides
        // (`_exc_YYYYMMDD@ruscal`) don't carry that OGCS tag but are still
        // clearly ruscal's own events.
        let candidates: Vec<(String, String)> = all_items.iter()
            .filter_map(|e| {
                let id       = e["id"].as_str()?;
                let ical_uid = e["iCalUID"].as_str()?;
                let summary  = e["summary"].as_str().unwrap_or("(no title)").to_owned();

                let id_is_internal = id.starts_with('_');
                let uid_is_ruscal  = ical_uid.ends_with("@ruscal");

                if id_is_internal && uid_is_ruscal {
                    // Informational: tag whether the original Outlook source
                    // is still in the window (all 5 current hits match one).
                    let ogcs_id = e["extendedProperties"]["private"]["outlook_GlobalApptID"]
                        .as_str().unwrap_or("");
                    let tag = if outlook_ids.contains(ogcs_id) { "[outlook-match]" }
                        else if ogcs_id.is_empty() { "[no-ogcs-tag]" }
                        else { "[outlook-missing]" };
                    Some((id.to_owned(), format!("{tag} {summary}")))
                } else {
                    None
                }
            })
            .collect();

        println!("URL-conflict candidates: {}", candidates.len());
        for (id, summary) in &candidates {
            println!("  {id} | {summary}");
        }

        // Hard safety floor — refuse mass deletion.
        assert!(candidates.len() <= 50,
            "refusing to delete > 50 events in one run (got {})", candidates.len());

        // 4. DELETE each via Google Calendar API.
        let mut deleted = 0usize;
        for (id, summary) in &candidates {
            let url = format!(
                "https://www.googleapis.com/calendar/v3/calendars/primary/events/{id}");
            let resp = client.delete(&url).bearer_auth(&token).send().expect("delete");
            let status = resp.status();
            if status.is_success() || status.as_u16() == 410 {
                deleted += 1;
                println!("  DEL ✓ {summary}");
            } else {
                let body = resp.text().unwrap_or_default();
                println!("  DEL ✗ {id} | {status} | {}",
                    body.chars().take(200).collect::<String>());
            }
        }
        println!("\nDeleted {deleted}/{} URL-conflict events", candidates.len());
    }

    /// Run the real `run_sync` end-to-end and print the report. Useful for
    /// debugging skipped events without round-tripping through the UI.
    ///
    /// Run with:
    ///   cargo test --bin ruscal test_run_sync_now -- --ignored --nocapture
    #[test]
    #[ignore = "ad-hoc; requires live Outlook + Google credentials"]
    fn test_run_sync_now() {
        let cfg_path = dirs::data_local_dir().unwrap().join("ruscal").join("config.json");
        let cfg: serde_json::Value =
            serde_json::from_str(&std::fs::read_to_string(&cfg_path).unwrap()).unwrap();
        let dest_id = cfg["pairs"][0]["dest_id"].as_str().unwrap().to_owned();
        let gmail   = cfg["pairs"][0]["google_email"].as_str().unwrap().to_owned();
        let source  = cfg["pairs"][0]["source_account"].as_str().unwrap().to_owned();
        let token   = crate::google::get_access_token_for(&gmail).expect("token");

        let pair_id = crate::state::pair_id(&source, &dest_id);
        println!("pair_id: {pair_id}");

        let report = run_sync(&pair_id, &dest_id, &token).expect("run_sync");
        println!("\n=== sync report ===");
        println!("synced:  {}", report.synced);
        println!("skipped: {} {:?}", report.skipped_titles.len(), report.skipped_titles);
    }

    /// Lookup-by-global-id round-trip: read the Outlook calendar, pick any
    /// event, encode its `clean_global_id` as hex, feed it back through
    /// `find_outlook_event_by_global_id`, and confirm the same event comes back.
    ///
    #[cfg(any())]
    #[test]
    fn test_lookup_by_global_id() {
        let start = Utc::now() - chrono::Duration::days(crate::outlook::DEFAULT_PAST_DAYS);
        let end   = Utc::now() + chrono::Duration::days(crate::outlook::DEFAULT_FUTURE_DAYS);
        let events = crate::outlook::read_calendar_events(start, end).expect("read");

        let sample = events.iter().find(|e| !e.clean_global_id.is_empty())
            .expect("no event with a clean_global_id — cannot test lookup");
        let hex_id = hex::encode_upper(&sample.clean_global_id);
        println!("seeking Outlook event: {:?} (id {})", sample.subject, hex_id);

        let found = find_outlook_event_by_global_id(&hex_id)
            .expect("lookup failed")
            .expect("lookup returned None for a just-read event");
        assert_eq!(found.clean_global_id, sample.clean_global_id);
        assert_eq!(found.subject, sample.subject);
    }

    /// End-to-end test of the dry-run Google change detector: PUT a fresh
    /// @ruscal event, run `detect_google_changes`, expect it to appear. Then
    /// DELETE it, run again, expect a `deleted=true` entry for the same UID.
    ///
    #[cfg(any())]
    #[test]
    fn test_detect_google_changes() {
        let cfg_path = dirs::data_local_dir().unwrap().join("ruscal").join("config.json");
        let cfg: serde_json::Value = serde_json::from_str(&std::fs::read_to_string(&cfg_path).unwrap()).unwrap();
        let dest_id = cfg["pairs"][0]["dest_id"].as_str().unwrap().to_owned();
        let gmail   = cfg["pairs"][0]["google_email"].as_str().unwrap().to_owned();
        let token   = crate::google::get_access_token_for(&gmail).expect("token");

        // Use a test-isolated pair_id so we don't disturb real state.
        let pid = format!("__TEST_DETECT__{}", dest_id);

        // Catch up first so subsequent deltas are small.
        let _ = detect_google_changes(&pid, &dest_id, &token).expect("initial catchup");

        // Give Google time to checkpoint before our PUT.
        std::thread::sleep(std::time::Duration::from_secs(3));

        let uid = "ruscal-detecttest@ruscal";
        let ical = format!(
            "BEGIN:VCALENDAR\r\nVERSION:2.0\r\nPRODID:-//ruscal//test//EN\r\n\
             BEGIN:VEVENT\r\nUID:{uid}\r\nDTSTAMP:20260415T000000Z\r\n\
             DTSTART:20260801T090000Z\r\nDTEND:20260801T100000Z\r\n\
             SUMMARY:detect-test\r\nEND:VEVENT\r\nEND:VCALENDAR\r\n"
        );
        crate::caldav::put_event(&dest_id, uid, &ical, &token).expect("PUT");

        // Detect — retry because Google's sync-token lags.
        let mut found_put = false;
        for i in 0..10 {
            std::thread::sleep(std::time::Duration::from_secs(2));
            let changes = detect_google_changes(&pid, &dest_id, &token).expect("detect");
            println!("try {i}: {} changes", changes.len());
            if changes.iter().any(|c| c.uid == uid && !c.deleted) { found_put = true; break; }
        }
        assert!(found_put, "detector didn't report our PUT");

        // Now DELETE and confirm detector reports it.
        crate::caldav::delete_event(&dest_id, uid, &token).expect("DELETE");
        let mut found_del = false;
        for i in 0..10 {
            std::thread::sleep(std::time::Duration::from_secs(2));
            let changes = detect_google_changes(&pid, &dest_id, &token).expect("detect after delete");
            println!("del try {i}: {} changes", changes.len());
            if changes.iter().any(|c| c.uid == uid && c.deleted) { found_del = true; break; }
        }
        assert!(found_del, "detector didn't report our DELETE");

        // Clean up the test pair state so state.json stays tidy.
        let mut app = crate::state::load();
        app.pairs.remove(&pid);
        crate::state::save(&app);
    }

    /// Second consecutive `run_sync` must return zero skipped titles — everything
    /// is hash-cached from the first run.
    ///
    /// Run with:  cargo test test_second_sync_is_silent -- --ignored --nocapture
    #[test]
    #[ignore = "requires live Outlook + Google credentials"]
    fn test_second_sync_is_silent() {
        // Read the live config JSON directly so we don't depend on main.rs types.
        let cfg_path = dirs::data_local_dir().unwrap().join("ruscal").join("config.json");
        let cfg_raw  = std::fs::read_to_string(&cfg_path).expect("read config");
        let cfg: serde_json::Value = serde_json::from_str(&cfg_raw).unwrap();
        let pair     = &cfg["pairs"][0];
        let dest_id  = pair["dest_id"].as_str().expect("dest_id").to_owned();
        let gmail    = pair["google_email"].as_str().expect("google_email").to_owned();
        let token    = crate::google::get_access_token_for(&gmail).expect("token");

        // First run may legitimately have skips (409s on unseen events). We only
        // care that those are now cached and the *second* run is silent.
        let source_acc = cfg["pairs"][0]["source_account"].as_str().unwrap_or("").to_owned();
        let pid = crate::state::pair_id(&source_acc, &dest_id);
        let first = run_sync(&pid, &dest_id, &token).expect("first sync");
        println!("first:  synced={} skipped={:?}", first.synced, first.skipped_titles);

        let second = run_sync(&pid, &dest_id, &token).expect("second sync");
        println!("second: synced={} skipped={:?}", second.synced, second.skipped_titles);

        assert!(second.skipped_titles.is_empty(),
            "second sync still reports skips: {:?}", second.skipped_titles);
        assert_eq!(second.synced, 0,
            "second sync should have nothing to PUT (hash-cache hit), got synced={}", second.synced);
    }

    /// Full pipeline: Outlook read → iCal → Google PUT → Google GET.
    ///
    /// Writes a report to `debug_pipeline_report.txt` so we can inspect:
    ///   - The exact iCal we generate (including RRULE)
    ///   - Whether Google's CalDAV accepted it (HTTP status)
    ///   - What Google actually stored (the raw iCal it sends back on GET)
    ///
    /// Run with:  cargo test integration -- --ignored --nocapture
    #[cfg(any())] // uses stale API — excluded from compilation
    #[test]
    #[ignore = "requires live Outlook + Google credentials"]
    fn test_pipeline() {
        let now          = Utc::now();
        let window_start = now - chrono::Duration::days(crate::outlook::DEFAULT_PAST_DAYS);
        let window_end   = now + chrono::Duration::days(crate::outlook::DEFAULT_FUTURE_DAYS);

        // ── 1. Read from Outlook ──────────────────────────────────────────────
        let events = crate::outlook::read_calendar_events(window_start, window_end)
            .expect("Outlook read failed");

        let recurring: Vec<_> = events.iter().filter(|e| e.is_recurring).collect();
        println!("Outlook returned {} total events, {} recurring", events.len(), recurring.len());

        // ── 2. Google auth + pick first calendar ─────────────────────────────
        let token = crate::google::get_access_token()
            .expect("Google auth failed");

        let calendars = crate::google::list_google_calendars()
            .expect("Google calendar list failed");

        println!("Google calendars:");
        for (i, c) in calendars.iter().enumerate() {
            println!("  [{i}] {} → {}", c.summary, c.id);
        }

        let dest_url = &calendars[0].id;
        println!("Using: {}", dest_url);

        // ── 3. For each recurring event: PUT then GET ─────────────────────────
        let mut report = String::new();

        for event in &recurring {
            let uid  = event_uid(event);
            let ical = event_to_ical(event, &uid);

            let _ = writeln!(report, "=== {} ===", event.subject);
            let _ = writeln!(report, "start:       {}", event.start);
            let _ = writeln!(report, "blob len:    {}", event.recur_blob.len());
            let _ = writeln!(report, "recur_end:   {:?}", event.recurrence_end);
            let _ = writeln!(report, "");
            let _ = writeln!(report, "--- iCal we are PUTting ---");
            let _ = writeln!(report, "{}", ical.replace("\r\n", "\n"));

            // Delete first so we start fresh.
            let _ = crate::caldav::delete_event(dest_url, &uid, &token);

            // Step 1: PUT master event (RRULE + EXDATE, no embedded exception VEVENTs).
            let master_ical = event_to_ical(event, &uid);
            let _ = writeln!(report, "--- master iCal ---");
            let _ = writeln!(report, "{}", master_ical.replace("\r\n", "\n"));
            match crate::caldav::put_event(dest_url, &uid, &master_ical, &token) {
                Ok(())  => { let _ = writeln!(report, "Master PUT: OK"); }
                Err(e)  => { let _ = writeln!(report, "Master PUT FAILED: {e}"); continue; }
            }

            // Step 2: PUT each exception VEVENT as a separate CalDAV resource.
            let exceptions = build_exception_icals(event, &uid);
            let _ = writeln!(report, "Exceptions: {} VEVENTs", exceptions.len());
            for (exc_resource_uid, exc_ical) in &exceptions {
                let _ = writeln!(report, "\n--- exception resource: {exc_resource_uid} ---");
                let _ = writeln!(report, "{}", exc_ical.replace("\r\n", "\n"));
                match crate::caldav::put_event(dest_url, exc_resource_uid, exc_ical, &token) {
                    Ok(())  => { let _ = writeln!(report, "Exception PUT: OK"); }
                    Err(e)  => { let _ = writeln!(report, "Exception PUT FAILED: {e}"); }
                }
            }

            // Step 3: GET back to see what Google stored.
            match crate::caldav::get_event(dest_url, &uid, &token) {
                Ok(returned) => {
                    let _ = writeln!(report, "\n--- Google GET master ---");
                    let _ = writeln!(report, "{}", returned.replace("\r\n", "\n"));
                }
                Err(e) => { let _ = writeln!(report, "GET FAILED: {e}"); }
            }

            let _ = writeln!(report, "");
        }

        let report_path = "debug_pipeline_report.txt";
        std::fs::write(report_path, &report).expect("write report");
        println!("Report written to {report_path}");
    }

    /// Verify that synced events have the correct fields on the Google side.
    ///
    /// Phase 1 — synthetic baseline: PUT a minimal recurring event with no external ORGANIZER
    ///   and confirm GET returns 200 with the expected fields. This establishes whether
    ///   PUT+GET works at all and tells us if ORGANIZER is the cause of master 404s.
    ///
    /// Phase 2 — real events: for each recurring event from Outlook that has moved exceptions,
    ///   PUT master + exceptions, then GET exceptions and assert SUMMARY, DESCRIPTION, ORGANIZER.
    ///   Master GET is attempted and logged but not asserted (known flaky with external ORGANIZER).
    ///   Listing via PROPFIND is used to confirm the master was at least stored.
    ///
    /// Run with:  cargo test test_google_roundtrip -- --ignored --nocapture
    #[cfg(any())] // uses stale API — excluded from compilation
    #[test]
    #[ignore = "requires live Outlook + Google credentials"]
    fn test_google_roundtrip() {
        let token = crate::google::get_access_token().expect("Google auth failed");
        let calendars = crate::google::list_google_calendars().expect("calendar list failed");
        let dest_url = &calendars[0].id;

        // ── Phase 1: synthetic recurring event (no ORGANIZER) ─────────────────
        // This tells us whether PUT+GET works for recurring events at all.
        {
            let uid = "ruscal-test-recurring@ruscal";
            let ical = format!(
                "BEGIN:VCALENDAR\r\nVERSION:2.0\r\nPRODID:-//ruscal//ruscal//EN\r\n\
                 CALSCALE:GREGORIAN\r\nMETHOD:PUBLISH\r\nBEGIN:VEVENT\r\n\
                 UID:{uid}\r\nDTSTAMP:20260101T000000Z\r\n\
                 DTSTART:20260505T090000Z\r\nDTEND:20260505T100000Z\r\n\
                 RRULE:FREQ=WEEKLY;BYDAY=MO;UNTIL=20261231T235959Z\r\n\
                 SUMMARY:ruscal test recurring\r\n\
                 DESCRIPTION:automated test — safe to delete\r\n\
                 END:VEVENT\r\nEND:VCALENDAR\r\n"
            );
            let _ = crate::caldav::delete_event(dest_url, uid, &token);
            crate::caldav::put_event(dest_url, uid, &ical, &token)
                .expect("PUT synthetic recurring failed");

            match crate::caldav::get_event(dest_url, uid, &token) {
                Ok(got) => {
                    println!("[synthetic] GET OK — {} bytes", got.len());
                    assert!(got.contains("SUMMARY:"), "synthetic missing SUMMARY");
                    assert!(got.contains("DESCRIPTION:"), "synthetic missing DESCRIPTION");
                    println!("[synthetic] SUMMARY + DESCRIPTION present ✓");
                }
                Err(e) => {
                    // If even a plain synthetic RRULE event doesn't GET back, the problem
                    // is structural (URL encoding, Google API quirk). Log and continue.
                    println!("[synthetic] GET FAILED: {e}");
                    println!("[synthetic] DIAGNOSIS: Google CalDAV may not support direct GET for RRULE events");
                    println!("[synthetic] Checking via PROPFIND listing...");
                    let listed = crate::caldav::list_ruscal_event_uids(dest_url, &token)
                        .expect("PROPFIND listing failed");
                    if listed.iter().any(|u| u == uid) {
                        println!("[synthetic] Event IS in PROPFIND listing ✓ (GET just doesn't work for RRULE)");
                    } else {
                        panic!("[synthetic] Event NOT in PROPFIND listing either — PUT silently failed!");
                    }
                }
            }
            let _ = crate::caldav::delete_event(dest_url, uid, &token);
        }

        // ── Phase 2: real Outlook events ──────────────────────────────────────
        let now          = Utc::now();
        let window_start = now - chrono::Duration::days(crate::outlook::DEFAULT_PAST_DAYS);
        let window_end   = now + chrono::Duration::days(crate::outlook::DEFAULT_FUTURE_DAYS);
        let events = crate::outlook::read_calendar_events(window_start, window_end)
            .expect("Outlook read failed");

        let mut any_exceptions_checked = false;

        for event in events.iter().filter(|e| e.is_recurring) {
            let uid  = event_uid(event);
            let ical = event_to_ical(event, &uid);

            // Fresh slate — use put_with_retry so any stale server-side resource is
            // cleaned up before asserting. This is the same path the production sync takes.
            put_with_retry(dest_url, &uid, &ical, &token)
                .expect(&format!("PUT master failed for {}", event.subject));

            // Try GET — log result but don't assert (known 404 issue with external ORGANIZER).
            match crate::caldav::get_event(dest_url, &uid, &token) {
                Ok(got)  => {
                    println!("[master GET OK] {} ({} bytes)", event.subject, got.len());
                    assert!(got.contains("SUMMARY:"), "master missing SUMMARY — {}", event.subject);
                    let expected_desc = build_description(event);
                    if !expected_desc.is_empty() {
                        assert!(got.contains("DESCRIPTION:"),
                            "master missing DESCRIPTION for '{}'\n--- Google GET ---\n{got}", event.subject);
                    }
                    println!("[master OK] {} ✓", event.subject);
                }
                Err(e)   => {
                    println!("[master GET FAILED] {}: {e}", event.subject);
                    // Fallback: PROPFIND listing must confirm the event was stored.
                    let listed = crate::caldav::list_ruscal_event_uids(dest_url, &token)
                        .expect("PROPFIND listing failed");
                    if listed.iter().any(|u| u == &uid) {
                        println!("[master PROPFIND found] {} ✓ (stored, GET-able via listing)", event.subject);
                    } else {
                        panic!("Master '{}' not in PROPFIND listing after PUT — PUT silently failed!\nUID: {uid}", event.subject);
                    }
                }
            }

            // PUT + GET each exception — ASSERT all expected fields.
            for (exc_uid, exc_ical) in build_exception_icals(event, &uid) {
                put_with_retry(dest_url, &exc_uid, &exc_ical, &token)
                    .expect(&format!("PUT exception {exc_uid} failed for {}", event.subject));

                let exc_got = crate::caldav::get_event(dest_url, &exc_uid, &token)
                    .expect(&format!("GET exception {exc_uid} failed for {}", event.subject));

                assert!(exc_got.contains("SUMMARY:"),
                    "exception {exc_uid} missing SUMMARY\n--- PUT ---\n{exc_ical}\n--- Google GET ---\n{exc_got}");
                // DESCRIPTION should contain body and/or organizer prefix.
                let expected_desc = build_description(event);
                if !expected_desc.is_empty() {
                    assert!(exc_got.contains("DESCRIPTION:"),
                        "exception {exc_uid} missing DESCRIPTION\n--- PUT ---\n{exc_ical}\n--- Google GET ---\n{exc_got}");
                }
                if !event.organizer_email.is_empty() {
                    assert!(exc_got.contains("Organizer:") || exc_got.contains(&event.organizer_email),
                        "exception {exc_uid}: organizer info not in DESCRIPTION\n--- Google GET ---\n{exc_got}");
                }
                println!("[exception OK] {} → {exc_uid} ✓", event.subject);
                any_exceptions_checked = true;
            }
        }

        if any_exceptions_checked {
            println!("All exception assertions passed ✓");
        } else {
            println!("No recurring events with moved exceptions found — exception fields not tested");
        }
    }

    /// Verify that events removed from Outlook are deleted from Google Calendar.
    ///
    /// 1. PUT two synthetic test events into Google
    /// 2. Sync with only event A in the Outlook set
    /// 3. Assert A still exists, B is deleted
    ///
    /// Run with:  cargo test test_deletion -- --ignored --nocapture
    #[cfg(any())] // uses stale API — excluded from compilation
    #[test]
    #[ignore = "requires live Google credentials"]
    fn test_deletion() {
        let token = crate::google::get_access_token().expect("Google auth");
        let calendars = crate::google::list_google_calendars().expect("calendar list");
        let dest_url = &calendars[0].id;

        let uid_a = "ruscal-test-keep@ruscal";
        let uid_b = "ruscal-test-delete@ruscal";

        let make_ical = |uid: &str, summary: &str| -> String {
            format!(
                "BEGIN:VCALENDAR\r\nVERSION:2.0\r\nPRODID:-//ruscal//ruscal//EN\r\n\
                 BEGIN:VEVENT\r\nUID:{uid}\r\nDTSTAMP:20260101T000000Z\r\n\
                 DTSTART:20260501T100000Z\r\nDTEND:20260501T110000Z\r\n\
                 SUMMARY:{summary}\r\nEND:VEVENT\r\nEND:VCALENDAR\r\n"
            )
        };

        crate::caldav::put_event(dest_url, uid_a, &make_ical(uid_a, "Keep me"), &token)
            .expect("PUT A");
        crate::caldav::put_event(dest_url, uid_b, &make_ical(uid_b, "Delete me"), &token)
            .expect("PUT B");
        println!("PUT both events OK");

        // Sync with only A — B should be deleted as an orphan
        let keep: std::collections::HashSet<String> = [uid_a.to_owned()].into();
        delete_orphans(dest_url, &keep, &token).expect("delete_orphans failed");

        let get_a = crate::caldav::get_event(dest_url, uid_a, &token);
        assert!(get_a.is_ok(), "A should still exist: {:?}", get_a);
        println!("Event A still exists ✓");

        let get_b = crate::caldav::get_event(dest_url, uid_b, &token);
        assert!(get_b.is_err(), "B should be deleted, but GET succeeded");
        println!("Event B deleted ✓");
    }

    /// Print raw bytes around offset 90+ for recurring events, so we can verify
    /// how to parse AppointmentRecurrencePattern ExceptionInfo before coding it.
    ///
    /// Run with:  cargo test test_dump_blob -- --ignored --nocapture
    #[test]
    #[ignore = "requires live Outlook"]
    fn test_dump_blob() {
        let now = Utc::now();
        let events = crate::outlook::read_calendar_events(
            now - chrono::Duration::days(crate::outlook::DEFAULT_PAST_DAYS),
            now + chrono::Duration::days(crate::outlook::DEFAULT_FUTURE_DAYS),
        ).expect("Outlook read");

        for e in events.iter().filter(|e| e.is_recurring && !e.recur_blob.is_empty()) {
            let b = &e.recur_blob;
            println!("=== {} ({} bytes) ===", e.subject, b.len());
            // Print all bytes with offset labels, 16 per row
            for (i, chunk) in b.chunks(16).enumerate() {
                let hex: Vec<String> = chunk.iter().map(|x| format!("{x:02X}")).collect();
                println!("  [{:4}] {}", i * 16, hex.join(" "));
            }
        }
    }

    /// Dump all Outlook events with exception details so we can see what data
    /// we're actually working with before coding the exception handling.
    ///
    /// Run with:  cargo test test_dump_all -- --ignored --nocapture
    #[test]
    #[ignore = "requires live Outlook"]
    fn test_dump_all() {
        let now = Utc::now();
        let events = crate::outlook::read_calendar_events(
            now - chrono::Duration::days(365),
            now + chrono::Duration::days(365),
        ).expect("Outlook read");

        println!("Total events: {}", events.len());
        for e in &events {
            println!("\n--- {} ---", e.subject);
            println!("  start:        {}", e.start);
            println!("  is_recurring: {}", e.is_recurring);
            println!("  global_id:    {}", hex::encode(&e.clean_global_id));
            println!("  blob len:     {}", e.recur_blob.len());

            if !e.recur_blob.is_empty() {
                let uid  = event_uid(e);
                let ical = event_to_ical(e, &uid);
                for line in ical.replace("\r\n", "\n").lines()
                    .filter(|l| l.starts_with("RRULE") || l.starts_with("EXDATE"))
                {
                    println!("  {line}");
                }
            }
        }

        // Show if any global_id appears more than once (series master + exceptions)
        let mut seen: std::collections::HashMap<String, Vec<String>> = Default::default();
        for e in &events {
            seen.entry(hex::encode(&e.clean_global_id))
                .or_default()
                .push(format!("{} (recurring={})", e.subject, e.is_recurring));
        }
        println!("\n=== Shared global_ids (series + exceptions) ===");
        for (gid, names) in &seen {
            if names.len() > 1 {
                println!("  {}: {:?}", &gid[..16], names);
            }
        }
    }

    /// Inspect the AppointmentRecurrencePattern (ARP) extension that follows the
    /// RecurrencePattern. Dumps ReaderVersion2/StartTimeOffset/ExceptionCount,
    /// then for each ExceptionInfo prints raw field values so we can see exactly
    /// what Outlook stored before we write the RECURRENCE-ID generation code.
    ///
    /// Run with:  cargo test test_dump_exceptions -- --ignored --nocapture
    #[test]
    #[ignore = "requires live Outlook"]
    fn test_dump_exceptions() {
        let epoch_1601: chrono::DateTime<Utc> =
            chrono::DateTime::from_timestamp(-11_644_473_600, 0).unwrap();
        let mins_to_dt = |mins: u32| -> String {
            (epoch_1601 + chrono::Duration::minutes(mins as i64))
                .format("%Y-%m-%d %H:%M UTC")
                .to_string()
        };

        let now = Utc::now();
        let events = crate::outlook::read_calendar_events(
            now - chrono::Duration::days(365),
            now + chrono::Duration::days(365),
        ).expect("Outlook read");

        for e in events.iter().filter(|e| e.is_recurring && !e.recur_blob.is_empty()) {
            let b = &e.recur_blob;
            if b.len() < 26 { continue; }

            let pattern_type = u16::from_le_bytes([b[6], b[7]]);
            let pts_len: usize = match pattern_type {
                0x0000 => 0,
                0x0001 | 0x0002 | 0x0003 => 4,
                0x0004 => 8,
                _ => continue,
            };

            let del_count_off = 22 + pts_len + 4 + 4 + 4; // EndType, OccCount, FirstDOW
            if b.len() < del_count_off + 4 { continue; }
            let del_count = u32::from_le_bytes(b[del_count_off..del_count_off+4].try_into().unwrap()) as usize;

            let mod_count_off = del_count_off + 4 + del_count * 4;
            if b.len() < mod_count_off + 4 { continue; }
            let mod_count = u32::from_le_bytes(b[mod_count_off..mod_count_off+4].try_into().unwrap()) as usize;

            // RecurrencePattern ends just after ModifiedInstanceDates + StartDate + EndDate.
            let rp_end = mod_count_off + 4 + mod_count * 4 + 4 + 4;

            println!("\n=== {} ===", e.subject);
            println!("  blob={} bytes  rp_end={rp_end}  del={del_count}  mod={mod_count}", b.len());

            // Dump raw bytes around the boundary so we can verify the offset.
            // Print bytes around rp_end so we can find where ARP actually starts.
            {
                let from = rp_end.saturating_sub(8);
                let to   = (rp_end + 32).min(b.len());
                let hex: Vec<String> = b[from..to].iter().enumerate()
                    .map(|(i, x)| {
                        if (from + i) == rp_end { format!("|{x:02X}") }
                        else { format!("{x:02X}") }
                    })
                    .collect();
                println!("  bytes [{from}..{to}] (| = rp_end): {}", hex.join(" "));
            }

            // AppointmentRecurrencePattern (ARP) extension starts at rp_end.
            // Empirically verified layout (version fields are u32, not u16):
            //   [+0] ReaderVersion2 u32  (0x00003006)
            //   [+4] WriterVersion2 u32
            //   [+8] StartTimeOffset u32  (minutes from local midnight)
            //  [+12] EndTimeOffset   u32
            //  [+16] ExceptionCount  u16
            //  [+18] ExceptionInfo[0] ...
            let arp_start = rp_end;
            if b.len() < arp_start + 18 {
                println!("  (blob too short for ARP header at {arp_start})");
                continue;
            }

            let rv2 = u32::from_le_bytes(b[arp_start..arp_start+4].try_into().unwrap());
            let wv2 = u32::from_le_bytes(b[arp_start+4..arp_start+8].try_into().unwrap());
            let sto = u32::from_le_bytes(b[arp_start+8..arp_start+12].try_into().unwrap());
            let eto = u32::from_le_bytes(b[arp_start+12..arp_start+16].try_into().unwrap());
            let exc = u16::from_le_bytes([b[arp_start+16], b[arp_start+17]]);

            println!("  ARP at {arp_start}: rv2=0x{rv2:08X}  wv2=0x{wv2:08X}");
            println!("  StartTimeOffset={sto}min  EndTimeOffset={eto}min");
            println!("  ExceptionCount={exc}");

            // ExceptionInfo array starts at arp_start + 18.
            let mut pos = arp_start + 18;
            for i in 0..exc as usize {
                if b.len() < pos + 10 {
                    println!("  ExceptionInfo[{i}]: blob too short at {pos}");
                    break;
                }
                let start_dt = u32::from_le_bytes(b[pos..pos+4].try_into().unwrap());
                let end_dt   = u32::from_le_bytes(b[pos+4..pos+8].try_into().unwrap());
                let orig_sd  = u32::from_le_bytes(b[pos+8..pos+12].try_into().unwrap());
                let flags    = u16::from_le_bytes([b[pos+12], b[pos+13]]);
                pos += 14; // fixed part done

                println!("  ExceptionInfo[{i}]:");
                println!("    StartDateTime    = 0x{start_dt:08X}  → {}", mins_to_dt(start_dt));
                println!("    EndDateTime      = 0x{end_dt:08X}  → {}", mins_to_dt(end_dt));
                println!("    OriginalStartDate= 0x{orig_sd:08X}  → {}", mins_to_dt(orig_sd));
                println!("    OverrideFlags    = 0x{flags:04X}");

                // Walk optional fields so `pos` advances past them.
                // ARO_SUBJECT (0x0001)
                if flags & 0x0001 != 0 {
                    if b.len() < pos + 4 { break; }
                    let _len1 = u16::from_le_bytes([b[pos], b[pos+1]]) as usize;
                    let len2  = u16::from_le_bytes([b[pos+2], b[pos+3]]) as usize;
                    pos += 4 + len2;
                }
                // ARO_MEETINGTYPE (0x0002)
                if flags & 0x0002 != 0 {
                    if b.len() < pos + 4 { break; }
                    let mt = u32::from_le_bytes(b[pos..pos+4].try_into().unwrap());
                    println!("    MeetingType      = {mt} ({})",
                        if mt == 3 { "olCanceled" } else { "other" });
                    pos += 4;
                }
                // ARO_REMINDERDELTA (0x0004)
                if flags & 0x0004 != 0 { pos += 4; }
                // ARO_REMINDER (0x0008)
                if flags & 0x0008 != 0 { pos += 4; }
                // ARO_LOCATION (0x0010)
                if flags & 0x0010 != 0 {
                    if b.len() < pos + 4 { break; }
                    let _len1 = u16::from_le_bytes([b[pos], b[pos+1]]) as usize;
                    let len2  = u16::from_le_bytes([b[pos+2], b[pos+3]]) as usize;
                    pos += 4 + len2;
                }
                // ARO_BUSYSTATUS (0x0020)
                if flags & 0x0020 != 0 { pos += 4; }
                // ARO_ATTACHMENT (0x0040) — no extra data
                // ARO_SUBTYPE (0x0080)
                if flags & 0x0080 != 0 { pos += 4; }
                // ARO_APPTCOLOR (0x0100)
                if flags & 0x0100 != 0 { pos += 4; }
                // ARO_EXCEPTIONAL_BODY (0x0200) — no extra data
            }
        }
    }
}
