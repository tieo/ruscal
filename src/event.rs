use std::fmt;

/// A single calendar event read from Outlook, ready to be synced.
#[derive(Debug, Clone)]
pub struct CalendarEvent {
    /// Event title.
    pub subject: String,
    /// UTC start time.
    pub start: chrono::DateTime<chrono::Utc>,
    /// UTC end time.
    pub end: chrono::DateTime<chrono::Utc>,
    /// Whether this is an all-day event (start/end are midnight UTC).
    pub is_all_day: bool,
    /// Free-text location or meeting URL.
    pub location: String,
    /// Organizer display name.
    pub organizer_name: String,
    /// Organizer SMTP email address.
    pub organizer_email: String,
    /// Plain-text body / notes.
    pub body: String,
    /// Busy status on the organizer's calendar.
    pub busy_status: BusyStatus,
    /// This attendee's response to the meeting request.
    pub response_status: ResponseStatus,
    /// Visibility / privacy level.
    pub sensitivity: Sensitivity,
    /// Whether this is the master of a recurring series.
    pub is_recurring: bool,
    /// Local calendar date of the last occurrence for a recurring series.
    ///
    /// `PidLidClipEnd` stores local midnight converted to UTC, so we convert
    /// back to the machine's local timezone to recover the correct calendar date.
    /// `None` for non-recurring events or open-ended series.
    pub recurrence_end: Option<chrono::NaiveDate>,
    /// Stable cross-system identifier (maps to `iCalUID` in Google Calendar).
    ///
    /// This is `PidLidCleanGlobalObjectId` from the MAPI named property set.
    /// It is identical across organizer + attendee copies and across all
    /// occurrences of a recurring series, making it the right key for sync.
    pub clean_global_id: Vec<u8>,
    /// Raw `PidLidAppointmentRecur` blob (MS-OXOCAL RecurrencePattern).
    ///
    /// Only present on recurring series masters. The sync engine reads the
    /// RecurFrequency and Period fields from this blob to expand the series
    /// into individual occurrences within the sync window.
    pub recur_blob: Vec<u8>,
}

/// The organizer's busy/free indication on their calendar.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum BusyStatus {
    Free,
    Tentative,
    Busy,
    OutOfOffice,
    WorkingElsewhere,
    Unknown(u32),
}

impl From<u32> for BusyStatus {
    fn from(v: u32) -> Self {
        match v {
            0 => Self::Free,
            1 => Self::Tentative,
            2 => Self::Busy,
            3 => Self::OutOfOffice,
            4 => Self::WorkingElsewhere,
            n => Self::Unknown(n),
        }
    }
}

impl fmt::Display for BusyStatus {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::Free              => f.write_str("Free"),
            Self::Tentative         => f.write_str("Tentative"),
            Self::Busy              => f.write_str("Busy"),
            Self::OutOfOffice       => f.write_str("OOF"),
            Self::WorkingElsewhere  => f.write_str("Working Elsewhere"),
            Self::Unknown(n)        => write!(f, "Unknown({n})"),
        }
    }
}

/// The local user's response to a meeting request.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ResponseStatus {
    None,
    Organized,
    Tentative,
    Accepted,
    Declined,
    NotResponded,
    Unknown(u32),
}

impl From<u32> for ResponseStatus {
    fn from(v: u32) -> Self {
        match v {
            0 => Self::None,
            1 => Self::Organized,
            2 => Self::Tentative,
            3 => Self::Accepted,
            4 => Self::Declined,
            5 => Self::NotResponded,
            n => Self::Unknown(n),
        }
    }
}

impl fmt::Display for ResponseStatus {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::None          => f.write_str("None"),
            Self::Organized     => f.write_str("Organizer"),
            Self::Tentative     => f.write_str("Tentative"),
            Self::Accepted      => f.write_str("Accepted"),
            Self::Declined      => f.write_str("Declined"),
            Self::NotResponded  => f.write_str("Not responded"),
            Self::Unknown(n)    => write!(f, "Unknown({n})"),
        }
    }
}

/// Visibility / privacy level of an event.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Sensitivity {
    Normal,
    Personal,
    Private,
    Confidential,
    Unknown(u32),
}

impl From<u32> for Sensitivity {
    fn from(v: u32) -> Self {
        match v {
            0 => Self::Normal,
            1 => Self::Personal,
            2 => Self::Private,
            3 => Self::Confidential,
            n => Self::Unknown(n),
        }
    }
}

impl fmt::Display for Sensitivity {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::Normal        => f.write_str("Normal"),
            Self::Personal      => f.write_str("Personal"),
            Self::Private       => f.write_str("Private"),
            Self::Confidential  => f.write_str("Confidential"),
            Self::Unknown(n)    => write!(f, "Unknown({n})"),
        }
    }
}

impl fmt::Display for CalendarEvent {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        let start_local = self.start.with_timezone(&chrono::Local);
        let end_local   = self.end.with_timezone(&chrono::Local);
        write!(
            f,
            "{start}–{end}{all_day}  {subject}",
            start   = start_local.format("%Y-%m-%d %H:%M"),
            end     = end_local.format("%H:%M"),
            all_day = if self.is_all_day { " (all day)" } else { "" },
            subject = self.subject,
        )?;
        if !self.location.is_empty() {
            write!(f, "\n  Location:  {}", self.location)?;
        }
        if !self.organizer_name.is_empty() {
            write!(f, "\n  Organizer: {} <{}>", self.organizer_name, self.organizer_email)?;
        }
        write!(f, "\n  Status:    {}  Busy: {}  Sensitivity: {}",
            self.response_status, self.busy_status, self.sensitivity)?;
        if self.is_recurring {
            match self.recurrence_end {
                Some(end) => write!(f, "\n  Recurring until {end}")?,
                None      => write!(f, "\n  Recurring (no end date)")?,
            }
        }
        if !self.clean_global_id.is_empty() {
            write!(f, "\n  GlobalID:  {}", hex::encode(&self.clean_global_id))?;
        }
        if !self.body.is_empty() {
            if let Some(first_line) = self.body.lines().next() {
                if !first_line.trim().is_empty() {
                    write!(f, "\n  Body:      {first_line}")?;
                }
            }
        }
        Ok(())
    }
}
