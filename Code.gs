/***************************
 * ICS CALENDAR SYNC FOR GOOGLE CALENDAR
 *
 * This script syncs multiple external ICS calendar feeds into your
 * Google Calendar. It handles:
 * - Multiple calendar sources (Outlook, Google Calendar, etc.)
 * - Proper timezone conversion (Windows to IANA timezone mapping)
 * - Event creation with color coding and prefixes
 * - Automatic removal of events deleted from source calendars
 *
 * SETUP:
 * 1. Configure your ICS feeds in the ICS_FEEDS array below
 * 2. Run syncAllIcsFeeds() to sync all calendars
 * 3. Set up a time-based trigger to run syncAllIcsFeeds() periodically
 *
 * CLEANUP:
 * Run deleteSyncedEvents() to remove all previously synced events
 ***************************/


/***************************
 * CONFIGURATION - ADD YOUR ICS FEEDS HERE
 *
 * Each feed object requires:
 *   - name: Unique identifier for the feed (used internally for tracking)
 *   - url: The ICS feed URL from your calendar provider
 *   - color: Event color in Google Calendar (see available colors below)
 *   - titlePrefix: Text prepended to event titles for easy identification
 *
 * Available CalendarApp.EventColor options:
 *   PALE_BLUE, PALE_GREEN, MAUVE, PALE_RED, YELLOW,
 *   ORANGE, CYAN, GRAY, BLUE, GREEN, RED
 ***************************/

const ICS_FEEDS = [
  {
    name: 'WORK_OUTLOOK',                              // Internal identifier (must be unique)
    url: 'https://outlook.office365.com/owa/calendar/YOUR_CALENDAR_ID@company.com/YOUR_SECRET_TOKEN/calendar.ics',
    color: CalendarApp.EventColor.ORANGE,              // Orange for work events
    titlePrefix: 'Work '                               // Prefix shown before event title
  },
  {
    name: 'TEAM_CALENDAR',
    url: 'https://outlook.office365.com/owa/calendar/TEAM_CALENDAR_ID@company.com/TEAM_SECRET_TOKEN/calendar.ics',
    color: CalendarApp.EventColor.BLUE,                // Blue for team events
    titlePrefix: 'Team '
  },
  {
    name: 'PERSONAL_GOOGLE',
    url: 'https://calendar.google.com/calendar/ical/your.email%40gmail.com/public/basic.ics',
    color: CalendarApp.EventColor.RED,                 // Red for personal events
    titlePrefix: 'Personal '
  },
  {
    name: 'PRIVATE_CALENDAR',
    url: 'https://calendar.google.com/calendar/ical/another.email%40gmail.com/private-YOUR_PRIVATE_TOKEN/basic.ics',
    color: CalendarApp.EventColor.GREEN,               // Green for private events
    titlePrefix: 'Private '
  }
  // Add more feeds by copying the object structure above
];

// Storage key prefix for tracking synced events per feed
const SINGLE_KEYS_PREFIX = 'ICS_SINGLE_EVENT_KEYS_';


/***************************
 * WINDOWS TO IANA TIMEZONE MAPPING
 *
 * Outlook and other Microsoft products use Windows timezone names,
 * while Google Calendar and most systems use IANA timezone names.
 * This map converts Windows timezone IDs to their IANA equivalents.
 *
 * Example: "Pacific Standard Time" -> "America/Los_Angeles"
 ***************************/
const WINDOWS_TO_IANA = {
  // ========== NORTH AMERICA ==========
  'Pacific Standard Time': 'America/Los_Angeles',          // US West Coast
  'Pacific Daylight Time': 'America/Los_Angeles',
  'Mountain Standard Time': 'America/Denver',              // US Mountain (observes DST)
  'Mountain Daylight Time': 'America/Denver',
  'US Mountain Standard Time': 'America/Phoenix',          // Arizona (no DST)
  'Central Standard Time': 'America/Chicago',              // US Central
  'Central Daylight Time': 'America/Chicago',
  'Eastern Standard Time': 'America/New_York',             // US East Coast
  'Eastern Daylight Time': 'America/New_York',
  'US Eastern Standard Time': 'America/Indiana/Indianapolis', // Indiana
  'Atlantic Standard Time': 'America/Halifax',             // Eastern Canada
  'Newfoundland Standard Time': 'America/St_Johns',        // Newfoundland (UTC-3:30)
  'Alaskan Standard Time': 'America/Anchorage',            // Alaska
  'Hawaiian Standard Time': 'Pacific/Honolulu',            // Hawaii (no DST)
  'Canada Central Standard Time': 'America/Regina',        // Saskatchewan (no DST)

  // ========== MEXICO ==========
  'Central Standard Time (Mexico)': 'America/Mexico_City',
  'Mountain Standard Time (Mexico)': 'America/Chihuahua',
  'Pacific Standard Time (Mexico)': 'America/Tijuana',

  // ========== SOUTH AMERICA ==========
  'SA Pacific Standard Time': 'America/Bogota',            // Colombia, Peru
  'SA Eastern Standard Time': 'America/Cayenne',           // French Guiana
  'SA Western Standard Time': 'America/La_Paz',            // Bolivia
  'E. South America Standard Time': 'America/Sao_Paulo',   // Brazil East
  'Argentina Standard Time': 'America/Buenos_Aires',
  'Venezuela Standard Time': 'America/Caracas',            // Venezuela (UTC-4)
  'Central Brazilian Standard Time': 'America/Cuiaba',     // Brazil Central

  // ========== EUROPE ==========
  'GMT Standard Time': 'Europe/London',                    // UK (observes DST)
  'Greenwich Standard Time': 'Atlantic/Reykjavik',         // Iceland (no DST)
  'W. Europe Standard Time': 'Europe/Berlin',              // Germany, Austria
  'Central Europe Standard Time': 'Europe/Budapest',       // Hungary
  'Central European Standard Time': 'Europe/Warsaw',       // Poland
  'Romance Standard Time': 'Europe/Paris',                 // France, Spain
  'E. Europe Standard Time': 'Europe/Chisinau',            // Moldova
  'FLE Standard Time': 'Europe/Kiev',                      // Ukraine, Finland
  'GTB Standard Time': 'Europe/Bucharest',                 // Romania, Greece
  'Russian Standard Time': 'Europe/Moscow',                // Russia (UTC+3)
  'Turkey Standard Time': 'Europe/Istanbul',               // Turkey
  'Belarus Standard Time': 'Europe/Minsk',                 // Belarus
  'Kaliningrad Standard Time': 'Europe/Kaliningrad',       // Russia exclave

  // ========== AFRICA ==========
  'Morocco Standard Time': 'Africa/Casablanca',
  'W. Central Africa Standard Time': 'Africa/Lagos',       // Nigeria, Cameroon
  'South Africa Standard Time': 'Africa/Johannesburg',     // South Africa
  'E. Africa Standard Time': 'Africa/Nairobi',             // Kenya, Ethiopia
  'Egypt Standard Time': 'Africa/Cairo',
  'Libya Standard Time': 'Africa/Tripoli',
  'Namibia Standard Time': 'Africa/Windhoek',

  // ========== MIDDLE EAST ==========
  'Israel Standard Time': 'Asia/Jerusalem',
  'Jordan Standard Time': 'Asia/Amman',
  'Arabic Standard Time': 'Asia/Baghdad',                  // Iraq
  'Arab Standard Time': 'Asia/Riyadh',                     // Saudi Arabia
  'Iran Standard Time': 'Asia/Tehran',                     // Iran (UTC+3:30)
  'Arabian Standard Time': 'Asia/Dubai',                   // UAE, Oman
  'Azerbaijan Standard Time': 'Asia/Baku',
  'Georgian Standard Time': 'Asia/Tbilisi',                // Georgia
  'Caucasus Standard Time': 'Asia/Yerevan',                // Armenia

  // ========== ASIA ==========
  'India Standard Time': 'Asia/Kolkata',                   // India (UTC+5:30)
  'Sri Lanka Standard Time': 'Asia/Colombo',
  'Nepal Standard Time': 'Asia/Kathmandu',                 // Nepal (UTC+5:45)
  'Bangladesh Standard Time': 'Asia/Dhaka',
  'Central Asia Standard Time': 'Asia/Almaty',             // Kazakhstan
  'Ekaterinburg Standard Time': 'Asia/Yekaterinburg',      // Russia (UTC+5)
  'Pakistan Standard Time': 'Asia/Karachi',
  'West Asia Standard Time': 'Asia/Tashkent',              // Uzbekistan
  'N. Central Asia Standard Time': 'Asia/Novosibirsk',     // Russia (UTC+7)
  'Myanmar Standard Time': 'Asia/Yangon',                  // Myanmar (UTC+6:30)
  'SE Asia Standard Time': 'Asia/Bangkok',                 // Thailand, Vietnam
  'North Asia Standard Time': 'Asia/Krasnoyarsk',          // Russia (UTC+7)
  'China Standard Time': 'Asia/Shanghai',                  // China
  'Singapore Standard Time': 'Asia/Singapore',             // Singapore, Malaysia
  'W. Australia Standard Time': 'Australia/Perth',         // Western Australia
  'Taipei Standard Time': 'Asia/Taipei',                   // Taiwan
  'Ulaanbaatar Standard Time': 'Asia/Ulaanbaatar',         // Mongolia
  'North Asia East Standard Time': 'Asia/Irkutsk',         // Russia (UTC+8)
  'Tokyo Standard Time': 'Asia/Tokyo',                     // Japan
  'Korea Standard Time': 'Asia/Seoul',                     // South Korea
  'Yakutsk Standard Time': 'Asia/Yakutsk',                 // Russia (UTC+9)
  'Vladivostok Standard Time': 'Asia/Vladivostok',         // Russia (UTC+10)
  'Magadan Standard Time': 'Asia/Magadan',                 // Russia (UTC+11)

  // ========== AUSTRALIA & PACIFIC ==========
  'Cen. Australia Standard Time': 'Australia/Adelaide',    // South Australia (UTC+9:30)
  'AUS Central Standard Time': 'Australia/Darwin',         // Northern Territory
  'E. Australia Standard Time': 'Australia/Brisbane',      // Queensland (no DST)
  'AUS Eastern Standard Time': 'Australia/Sydney',         // NSW, Victoria (DST)
  'Tasmania Standard Time': 'Australia/Hobart',            // Tasmania
  'West Pacific Standard Time': 'Pacific/Port_Moresby',    // Papua New Guinea
  'Central Pacific Standard Time': 'Pacific/Guadalcanal',  // Solomon Islands
  'Fiji Standard Time': 'Pacific/Fiji',
  'New Zealand Standard Time': 'Pacific/Auckland',         // New Zealand
  'Tonga Standard Time': 'Pacific/Tongatapu',
  'Samoa Standard Time': 'Pacific/Apia',
  'Line Islands Standard Time': 'Pacific/Kiritimati',      // UTC+14 (earliest timezone)

  // ========== UTC ==========
  'UTC': 'Etc/UTC',
  'Coordinated Universal Time': 'Etc/UTC'
};


/***************************
 * DELETE ALL SYNCED EVENTS
 *
 * Removes ALL events that were previously synced by this script.
 * Useful for:
 * - Clean slate before re-syncing
 * - Removing events after disconnecting a calendar
 * - Troubleshooting sync issues
 *
 * Detection methods (any match triggers deletion):
 * 1. Description contains "Synced from ICS feed:"
 * 2. Description has UID, DTSTART, and DTEND fields
 * 3. Title starts with any configured feed prefix
 ***************************/
function deleteSyncedEvents() {
  // Get the user's default (primary) calendar
  const cal = CalendarApp.getDefaultCalendar();
  if (!cal) throw new Error('Default calendar not found');

  // Search a wide date range to catch all synced events
  const from = new Date(2000, 0, 1);   // January 1, 2000
  const to   = new Date(2100, 11, 31); // December 31, 2100

  const events = cal.getEvents(from, to);

  let scanned = 0;
  let deleted = 0;

  // Check each event against multiple detection criteria
  events.forEach(ev => {
    scanned++;

    const desc  = ev.getDescription() || '';
    const title = ev.getTitle() || '';

    let shouldDelete = false;

    // Method 1: Check for sync marker in description
    if (desc.indexOf('Synced from ICS feed:') !== -1) {
      shouldDelete = true;
    }

    // Method 2: Check for ICS metadata fields (legacy detection)
    if (!shouldDelete) {
      const hasUid     = /UID:\s*\S+/.test(desc);
      const hasDtStart = /DTSTART:\s*\S+/.test(desc);
      const hasDtEnd   = /DTEND:\s*\S+/.test(desc);
      if (hasUid && hasDtStart && hasDtEnd) {
        shouldDelete = true;
      }
    }

    // Method 3: Check if title starts with any configured prefix
    if (!shouldDelete) {
      for (let i = 0; i < ICS_FEEDS.length; i++) {
        const feed = ICS_FEEDS[i];
        if (feed.titlePrefix && title.startsWith(feed.titlePrefix)) {
          shouldDelete = true;
          break;
        }
      }
    }

    if (shouldDelete) {
      ev.deleteEvent();
      deleted++;
    }
  });

  // Clear stored event keys for all feeds
  const props = PropertiesService.getScriptProperties();
  ICS_FEEDS.forEach(feed => props.deleteProperty(SINGLE_KEYS_PREFIX + feed.name));

  Logger.log('deleteSyncedEvents: scanned=' + scanned + ', deleted=' + deleted);
}


/***************************
 * SYNC ALL FEEDS
 *
 * Main entry point - iterates through all configured ICS feeds
 * and syncs each one to the Google Calendar.
 *
 * Set up a time-based trigger to run this function:
 * - Every 15 minutes for near real-time sync
 * - Every hour for less frequent updates
 * - Daily for infrequent calendars
 ***************************/
function syncAllIcsFeeds() {
  Logger.log('=== syncAllIcsFeeds START ===');

  // Process each configured feed
  ICS_FEEDS.forEach(feed => {
    if (!feed.url) return;  // Skip feeds without URLs
    syncSingleIcsFeed(feed);
  });

  Logger.log('=== syncAllIcsFeeds END ===');
}


/***************************
 * SYNC ONE FEED (SINGLE EVENTS ONLY)
 *
 * Syncs a single ICS feed to Google Calendar.
 *
 * Features:
 * - Only syncs single (non-recurring) events
 * - Only syncs events from today onwards (future events)
 * - Tracks synced events to avoid duplicates
 * - Removes events that no longer exist in the source
 * - Properly converts timezones from source calendar
 *
 * @param {Object} feed - Feed configuration object from ICS_FEEDS
 ***************************/
function syncSingleIcsFeed(feed) {
  // Get the user's default calendar
  const cal = CalendarApp.getDefaultCalendar();
  if (!cal) throw new Error('Default calendar not found');

  // Calculate the sync window (today and future)
  const now = new Date();
  const windowStart = new Date(now.getFullYear(), now.getMonth(), now.getDate());
  Logger.log('--- Syncing FEED: ' + feed.name + ' ---');

  // Get script properties for persistent storage of synced event keys
  const props = PropertiesService.getScriptProperties();
  const feedKey = SINGLE_KEYS_PREFIX + feed.name;

  // Load previously synced event keys for this feed
  const processedJson = props.getProperty(feedKey) || '[]';
  const oldKeys = new Set(JSON.parse(processedJson) || []);

  // Fetch the ICS file from the remote URL
  const resp = UrlFetchApp.fetch(feed.url);
  if (resp.getResponseCode() !== 200) {
    Logger.log('ICS fetch FAILED for feed ' + feed.name);
    return;
  }

  // Parse the ICS content - split into individual VEVENT blocks
  const ics = resp.getContentText();
  const vevents = ics.split('BEGIN:VEVENT').slice(1);  // Remove header before first event
  Logger.log(feed.name + ': VEVENT blocks: ' + vevents.length);

  let created = 0;
  const newKeys = new Set();  // Track all current event keys

  // Process each VEVENT block
  vevents.forEach(block => {
    // Extract ICS fields from the event block
    const uid      = getIcsField(block, 'UID');              // Unique event identifier
    const dtStartF = getIcsDateFieldRaw(block, 'DTSTART');   // Start date/time with timezone
    const dtEndF   = getIcsDateFieldRaw(block, 'DTEND');     // End date/time with timezone
    const recId    = getIcsField(block, 'RECURRENCE-ID');    // Exception to recurring event
    const rrule    = getIcsField(block, 'RRULE');            // Recurrence rule
    const summary  = decodeIcsText(getIcsField(block, 'SUMMARY') || '');     // Event title
    const location = decodeIcsText(getIcsField(block, 'LOCATION') || '');    // Event location
    const description = decodeIcsText(getIcsField(block, 'DESCRIPTION') || ''); // Event description

    // Skip events missing required fields
    if (!uid || !dtStartF || !dtEndF) return;

    // Parse dates with proper timezone conversion
    const start = parseIcsDateTime(dtStartF);
    const end   = parseIcsDateTime(dtEndF);

    if (!start || !end) {
      Logger.log(feed.name + ' PARSE FAILED: uid=' + uid);
      return;
    }

    // Determine event type - we only sync single (non-recurring) events
    const isRecurring = !!rrule;       // Has recurrence rule
    const isException = !!recId;        // Is an exception to a recurring event
    const isSingle    = !isRecurring && !isException;

    // Skip recurring events and exceptions (only sync single events)
    if (!isSingle) return;

    // Skip past events (only sync future events)
    if (start < windowStart) return;

    // Create unique key for this event instance
    // Format: UID|SINGLE|DTSTART_VALUE
    const key = uid + '|SINGLE|' + dtStartF.value;

    // Track this key as existing in current ICS
    newKeys.add(key);

    // Skip if already synced (avoid duplicates)
    if (oldKeys.has(key)) {
      return;
    }

    // Build event title with feed prefix
    const cleanSummary = summary || '(No title)';
    const title = feed.titlePrefix + cleanSummary;

    Logger.log(feed.name + ' CREATE: ' + title);
    Logger.log('  RAW: ' + dtStartF.value + ' TZID: ' + (dtStartF.tzid || 'none'));
    Logger.log('  CONVERTED: ' + start.toString());

    // Build detailed description with sync metadata and timezone info
    // This helps with debugging and allows the delete function to identify synced events
    const diagDesc =
      'Synced from ICS feed: ' + feed.name + ' (single)\n' +
      'KEY: ' + key + '\n' +
      'UID: ' + uid + '\n' +
      '─────────────────────────────────\n' +
      'DTSTART RAW: ' + dtStartF.rawLine + '\n' +
      'DTEND RAW: ' + dtEndF.rawLine + '\n' +
      '─────────────────────────────────\n' +
      'SOURCE TZID: ' + (dtStartF.tzid || 'NONE') + '\n' +
      'IANA TZID: ' + (dtStartF.ianaTzid || 'NONE') + '\n' +
      'CONVERTED START: ' + start.toString() + '\n' +
      'CONVERTED END: ' + end.toString() + '\n' +
      '─────────────────────────────────\n' +
      (location ? 'LOCATION: ' + location + '\n' : '') +
      (description ? '\nORIGINAL DESC:\n' + description : '');

    // Create the event in Google Calendar
    const ev = cal.createEvent(title, start, end, {
      description: diagDesc
    });

    // Apply feed-specific color
    if (feed.color) ev.setColor(feed.color);

    created++;
  });

  Logger.log(feed.name + ': Created: ' + created);

  /******************************************
   * DELETE EVENTS NO LONGER IN ICS FEED
   *
   * Compare old keys with new keys to find events
   * that were deleted from the source calendar
   ******************************************/
  const removedKeys = [...oldKeys].filter(k => !newKeys.has(k));

  if (removedKeys.length > 0) {
    const removedSet = new Set(removedKeys);

    // Search calendar for events to delete
    const from = new Date(now.getFullYear() - 5, 0, 1);
    const to   = new Date(now.getFullYear() + 5, 11, 31);
    const events = cal.getEvents(from, to);

    events.forEach(ev => {
      const desc = ev.getDescription() || '';

      // Only check events from this specific feed
      if (!desc.includes('Synced from ICS feed: ' + feed.name)) return;

      // Extract the KEY from the event description
      const keyLine = extractLineValue(desc, 'KEY:');
      if (!keyLine) return;

      // Delete if the key is no longer in the source ICS
      if (removedSet.has(keyLine.trim())) {
        ev.deleteEvent();
      }
    });
  }

  // Save the new set of keys for next sync
  props.setProperty(feedKey, JSON.stringify([...newKeys]));
  Logger.log('--- Finished FEED: ' + feed.name + ' ---');
}


/***************************
 * HELPER FUNCTIONS
 ***************************/


/**
 * Extract a simple ICS field value (no parameters needed)
 *
 * Handles ICS line folding (continuation lines starting with space/tab)
 * Works for: UID, SUMMARY, DESCRIPTION, LOCATION, RRULE, etc.
 *
 * @param {string} block - The VEVENT block text
 * @param {string} fieldName - Field name to extract (e.g., 'UID', 'SUMMARY')
 * @returns {string|null} - Field value or null if not found
 */
function getIcsField(block, fieldName) {
  // First unfold the block (join continuation lines)
  // ICS spec: lines starting with space or tab are continuations
  const unfoldedBlock = block.replace(/\r?\n[ \t]/g, '');

  // Match field with optional parameters: FIELDNAME;params:value
  const re = new RegExp(fieldName + '(?:;[^:]*)?:([^\\r\\n]+)', 'i');
  const m = unfoldedBlock.match(re);
  if (!m) return null;

  return m[1].trim();
}


/**
 * Extract a date/time field with full parameter information
 *
 * ICS date fields can have various formats:
 * - DTSTART:20251201T090000Z (UTC time)
 * - DTSTART;TZID=America/Chicago:20251201T090000 (with timezone)
 * - DTSTART;TZID="Eastern Standard Time":20251201T090000 (Windows TZ, quoted)
 * - DTSTART;VALUE=DATE:20251201 (all-day event)
 *
 * @param {string} block - The VEVENT block text
 * @param {string} fieldName - Field name ('DTSTART' or 'DTEND')
 * @returns {Object|null} - Object with value, tzid, ianaTzid, isUtc, rawLine
 */
function getIcsDateFieldRaw(block, fieldName) {
  // Unfold continuation lines first
  const unfoldedBlock = block.replace(/\r?\n[ \t]/g, '');

  // Match: FIELDNAME;params:value
  const re = new RegExp('(' + fieldName + '([^:]*):([^\\r\\n]+))', 'i');
  const m = unfoldedBlock.match(re);
  if (!m) return null;

  const rawLine = m[1].trim();   // Full line for debugging
  const params = m[2] || '';      // Parameters (;TZID=..., etc.)
  const value = m[3].trim();      // The actual date/time value

  // Extract TZID parameter - handles both quoted and unquoted values
  // Examples: TZID=America/Chicago or TZID="Eastern Standard Time"
  let tzid = null;
  const tzidMatch = params.match(/TZID="?([^";]+)"?/i);
  if (tzidMatch) {
    tzid = tzidMatch[1].trim();
  }

  // Map Windows timezone to IANA timezone if needed
  // This is crucial for Outlook calendars which use Windows TZ names
  let ianaTzid = null;
  if (tzid) {
    ianaTzid = WINDOWS_TO_IANA[tzid] || tzid; // Use mapping or assume already IANA
  }

  // Check if time is in UTC (ends with 'Z')
  const isUtc = value.endsWith('Z');

  return {
    rawLine: rawLine,       // Original line for debugging
    params: params,         // Raw parameters string
    value: value,           // Date/time value (may include 'Z')
    tzid: tzid,             // Original timezone ID (may be Windows format)
    ianaTzid: ianaTzid,     // Converted IANA timezone ID
    isUtc: isUtc            // True if UTC time (ends with Z)
  };
}


/**
 * Parse ICS date/time with proper timezone conversion
 *
 * Handles three types of dates:
 * 1. All-day events: YYYYMMDD (8 digits)
 * 2. UTC times: YYYYMMDDTHHMMSSZ (ends with Z)
 * 3. Timezone-specific: YYYYMMDDTHHMMSS with TZID parameter
 * 4. Local time: YYYYMMDDTHHMMSS (no timezone = script timezone)
 *
 * @param {Object} dateField - Object from getIcsDateFieldRaw()
 * @returns {Date|null} - JavaScript Date object in local timezone
 */
function parseIcsDateTime(dateField) {
  if (!dateField) return null;

  let value = dateField.value;
  const tzid = dateField.tzid;
  const ianaTzid = dateField.ianaTzid;
  const isUtc = dateField.isUtc;

  // All-day event: YYYYMMDD (8 digits, no time component)
  if (/^\d{8}$/.test(value)) {
    const y = +value.slice(0, 4);
    const m = +value.slice(4, 6) - 1;  // JS months are 0-indexed
    const d = +value.slice(6, 8);
    return new Date(y, m, d);
  }

  // Strip 'Z' suffix if present (we handle UTC separately)
  if (value.endsWith('Z')) {
    value = value.slice(0, -1);
  }

  // Validate datetime format: YYYYMMDDTHHMMSS
  if (!/^\d{8}T\d{6}$/.test(value)) {
    return null;
  }

  // Parse the datetime components
  const y  = +value.slice(0, 4);        // Year
  const m  = +value.slice(4, 6) - 1;    // Month (0-indexed)
  const d  = +value.slice(6, 8);        // Day
  const hh = +value.slice(9, 11);       // Hour
  const mm = +value.slice(11, 13);      // Minute
  const ss = +value.slice(13, 15);      // Second

  // UTC time: Create date directly in UTC
  if (isUtc) {
    return new Date(Date.UTC(y, m, d, hh, mm, ss));
  }

  // Timezone-specific time: Convert from source timezone to script timezone
  if (ianaTzid) {
    // Format as ISO-like string for parsing
    const isoStr = Utilities.formatString(
      '%04d-%02d-%02d %02d:%02d:%02d',
      y, m + 1, d, hh, mm, ss
    );

    try {
      // Utilities.parseDate interprets the time in the given timezone
      // and converts it to the script's timezone automatically
      const parsed = Utilities.parseDate(isoStr, ianaTzid, 'yyyy-MM-dd HH:mm:ss');
      Logger.log('Timezone conversion: ' + value + ' in ' + ianaTzid + ' -> ' + parsed.toString());
      return parsed;
    } catch (e) {
      Logger.log('Failed to parse with timezone ' + ianaTzid + ': ' + e.message);
      // Fall through to local time interpretation
    }
  }

  // No timezone info: Treat as local time (script timezone)
  return new Date(y, m, d, hh, mm, ss);
}


/**
 * Decode ICS-escaped text
 *
 * ICS format requires escaping of certain characters:
 * - \, -> , (comma)
 * - \n -> newline
 * - \\ -> \ (backslash)
 *
 * @param {string} t - ICS-encoded text
 * @returns {string} - Decoded plain text
 */
function decodeIcsText(t) {
  return (t || '')
    .replace(/\\,/g, ',')      // Unescape commas
    .replace(/\\n/g, '\n')     // Convert \n to actual newline
    .replace(/\\\\/g, '\\');   // Unescape backslashes
}


/**
 * Extract a labeled value from event description
 *
 * Searches for lines like "KEY: some-value" and returns "some-value"
 *
 * @param {string} desc - Event description text
 * @param {string} label - Label to find (e.g., 'KEY:', 'UID:')
 * @returns {string|null} - Value after the label, or null if not found
 */
function extractLineValue(desc, label) {
  const lines = desc.split(/\r?\n/);
  for (let i = 0; i < lines.length; i++) {
    const line = lines[i].trim();
    if (line.toUpperCase().startsWith(label.toUpperCase())) {
      return line.substring(label.length).trim();
    }
  }
  return null;
}
