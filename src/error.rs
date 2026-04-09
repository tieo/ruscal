use std::fmt;

/// An error returned by the MAPI subsystem, wrapping an HRESULT code.
///
/// HRESULT is Windows' standard 32-bit error code format. Negative values
/// indicate failure; `S_OK` (0) is success. MAPI defines its own codes in
/// the `0x8004xxxx` range on top of the standard Win32 ones.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct MapiError(pub u32);

impl MapiError {
    /// The HRESULT code.
    #[allow(dead_code)]
    pub fn code(self) -> u32 {
        self.0
    }
}

impl fmt::Display for MapiError {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        // Map the most common MAPI HRESULTs to readable names.
        // Full list: https://learn.microsoft.com/en-us/office/client-developer/outlook/mapi/mapi-error-codes
        let name = match self.0 {
            0x80040000 => "MAPI_E_NO_SUPPORT",
            0x80040102 => "MAPI_E_BAD_CHARWIDTH",
            0x80040105 => "MAPI_E_STRING_TOO_LONG",
            0x80040106 => "MAPI_E_CALL_FAILED",
            0x80040107 => "MAPI_E_NOT_ENOUGH_RESOURCES",
            0x8004010F => "MAPI_E_NOT_FOUND",
            0x80040110 => "MAPI_E_VERSION",
            0x80040111 => "MAPI_E_LOGON_FAILED",
            0x80040112 => "MAPI_E_SESSION_LIMIT",
            0x80040113 => "MAPI_E_USER_CANCEL",
            0x80040114 => "MAPI_E_UNABLE_TO_ABORT",
            0x80040115 => "MAPI_E_NETWORK_ERROR",
            0x80040116 => "MAPI_E_DISK_ERROR",
            0x80040117 => "MAPI_E_TOO_COMPLEX",
            0x80040119 => "MAPI_E_OBJECT_CHANGED",
            0x8004011A => "MAPI_E_OBJECT_DELETED",
            0x8004011B => "MAPI_E_BUSY",
            0x8004011D => "MAPI_E_NOT_ENOUGH_DISK",
            0x8004011E => "MAPI_E_NOT_ENOUGH_MEMORY",
            0x8004011F => "MAPI_E_NOT_INITIALIZED",
            0x80040120 => "MAPI_E_NO_ACCESS",
            0x80040121 => "MAPI_E_NOT_ENOUGH_QUOTA",
            0x80040122 => "MAPI_E_INTERFACE_NOT_SUPPORTED",
            0x80040123 => "MAPI_E_TIMEOUT",
            0x80040124 => "MAPI_E_TABLE_EMPTY",
            0x80040125 => "MAPI_E_TABLE_TOO_BIG",
            0x80040129 => "MAPI_E_INVALID_BOOKMARK",
            0x80040200 => "MAPI_E_COLLISION",
            0x80040201 => "MAPI_E_NOT_ME",
            0x80040202 => "MAPI_E_CORRUPT_STORE",
            0x80040203 => "MAPI_E_NOT_IN_QUEUE",
            0x80040204 => "MAPI_E_NO_SUPPRESS",
            0x80040206 => "MAPI_E_COLLISION_HANDLING",
            0x80040207 => "MAPI_E_DECLINE_COPY",
            0x80040208 => "MAPI_E_UNEXPECTED_ID",
            0x80040400 => "MAPI_E_CORRUPT_DATA",
            0x80040401 => "MAPI_E_INVALID_PARAMETER",
            0x80040402 => "MAPI_E_INVALID_ENTRYID",
            0x80040403 => "MAPI_E_INVALID_OBJECT",
            0x80040404 => "MAPI_E_OBJECT_REMOVED",
            0x80040405 => "MAPI_E_INVALID_PROPS",
            0x80040406 => "MAPI_E_REQUIRED_PROPERTY_MISSING",
            0x80040407 => "MAPI_E_INCOMPLETE_RECIPIENTS",
            0x80040408 => "MAPI_E_NO_INTERFACE",
            0x80040600 => "MAPI_E_AMBIGUOUS_RECIP",
            _          => "MAPI_E_UNKNOWN",
        };
        write!(f, "{name} (0x{:08X})", self.0)
    }
}

impl std::error::Error for MapiError {}

/// Turns an i32 HRESULT into `Ok(())` or `Err(MapiError)`.
pub fn check_hr(hr: i32) -> Result<(), MapiError> {
    if hr >= 0 {
        Ok(())
    } else {
        Err(MapiError(hr as u32))
    }
}
