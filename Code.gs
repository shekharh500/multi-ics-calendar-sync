/***************************
 * ICS CALENDAR SYNC FOR GOOGLE CALENDAR
 * Syncs multiple ICS feeds to your Google Calendar
 * Removes events when they're deleted from source
 ***************************/

/***************************
 * CONFIG - ADD YOUR ICS FEEDS HERE
 ***************************/

const ICS_FEEDS = [
  {
    name: 'WORK_CALENDAR',
    url: 'https://outlook.office365.com/owa/calendar/YOUR_CALENDAR_ID/calendar.ics',
    color: CalendarApp.EventColor.ORANGE,
    titlePrefix: 'Work '
  },
  {
    name: 'PERSONAL_CALENDAR',
    url: 'https://calendar.google.com/calendar/ical/YOUR_EMAIL/public/basic.ics',
    color: CalendarApp.EventColor.BLUE,
    titlePrefix: 'Personal '
  }
  // Add more feeds as needed
  // Available colors: PALE_BLUE, PALE_GREEN, MAUVE, PALE_RED, YELLOW, ORANGE, CYAN, GRAY, BLUE, GREEN, RED
];

const SINGLE_KEYS_PREFIX = 'ICS_SINGLE_EVENT_KEYS_';


/***************************
 * DELETE ALL SYNCED EVENTS
 * Use this to clean up before re-syncing
 ***************************/
function deleteSyncedEvents() {
  const cal = CalendarApp.getDefaultCalendar();
  if (!cal) throw new Error('Default calendar not found');

  const now = new Date();
  const from = new Date(now.getFullYear() - 2, 0, 1);
  const to   = new Date(now.getFullYear() + 3, 11, 31);
  const events = cal.getEvents(from, to);

  let deleted = 0;

  events.forEach(ev => {
    const desc = ev.getDescription() || '';
    if (desc.includes('Synced from ICS feed:')) {
      ev.deleteEvent();
      deleted++;
    }
  });

  const props = PropertiesService.getScriptProperties();
  ICS_FEEDS.forEach(feed => props.deleteProperty(SINGLE_KEYS_PREFIX + feed.name));

  Logger.log('Deleted events: ' + deleted);
}


/***************************
 * SYNC ALL FEEDS
 * Run this function to sync all configured ICS feeds
 ***************************/
function syncAllIcsFeeds() {
  Logger.log('=== syncAllIcsFeeds START ===');
  ICS_FEEDS.forEach(feed => {
    if (!feed.url) return;
    syncSingleIcsFeed(feed);
  });
  Logger.log('=== syncAllIcsFeeds END ===');
}


/***************************
 * SYNC ONE FEED (SINGLE EVENTS ONLY)
 ***************************/
function syncSingleIcsFeed(feed) {
  const cal = CalendarApp.getDefaultCalendar();
  if (!cal) throw new Error('Default calendar not found');

  const now = new Date();
  const windowStart = new Date(now.getFullYear(), now.getMonth(), now.getDate());
  Logger.log('--- Syncing FEED: ' + feed.name + ' ---');
  Logger.log('Future window start (>=): ' + windowStart);

  const props = PropertiesService.getScriptProperties();
  const feedKey = SINGLE_KEYS_PREFIX + feed.name;

  // Previously synced single-event keys for this feed
  const processedJson = props.getProperty(feedKey) || '[]';
  const oldKeys = new Set(JSON.parse(processedJson) || []);
  Logger.log(feed.name + ': Previously stored keys: ' + oldKeys.size);

  // Fetch ICS
  const resp = UrlFetchApp.fetch(feed.url);
  if (resp.getResponseCode() !== 200) {
    Logger.log('ICS fetch FAILED for feed ' + feed.name + ': ' + resp.getResponseCode());
    return;
  }

  const ics = resp.getContentText();
  const vevents = ics.split('BEGIN:VEVENT').slice(1);
  Logger.log(feed.name + ': VEVENT blocks: ' + vevents.length);

  let created = 0;
  const newKeys = new Set(); // keys that still exist in current ICS

  vevents.forEach(block => {
    const uid      = getIcsField(block, 'UID');
    const dtStartF = getIcsDateField(block, 'DTSTART');
    const dtEndF   = getIcsDateField(block, 'DTEND');
    const recId    = getIcsField(block, 'RECURRENCE-ID');
    const rrule    = getIcsField(block, 'RRULE');
    const summary  = decodeIcsText(getIcsField(block, 'SUMMARY') || '');

    if (!uid || !dtStartF || !dtEndF) return;

    const start = parseIcsDate(dtStartF);
    const end   = parseIcsDate(dtEndF);
    if (!start || !end) return;

    const isRecurring = !!rrule;
    const isException = !!recId;
    const isSingle    = !isRecurring && !isException;

    if (!isSingle) return;
    if (start < windowStart) return;

    // Use raw ICS value for the key so it matches exactly
    const dtStartRaw = dtStartF.value;
    const dtEndRaw   = dtEndF.value;
    const key = uid + '|SINGLE|' + dtStartRaw;

    newKeys.add(key); // this key still exists in ICS

    if (oldKeys.has(key)) {
      // already synced, keep it
      return;
    }

    // Title: prefix + SUMMARY
    const cleanSummary = summary || '(No title)';
    const title = feed.titlePrefix + cleanSummary;

    Logger.log(feed.name + ' CREATE SINGLE: ' + key + '  ' + title + '  ' + start + ' -> ' + end);

    const ev = cal.createEvent(title, start, end, {
      description:
        'Synced from ICS feed: ' + feed.name + ' (single)\n' +
        'KEY: ' + key + '\n' +
        'UID: ' + uid + '\n' +
        'DTSTART: ' + dtStartRaw + '\n' +
        'DTEND: ' + dtEndRaw
    });

    if (feed.color) ev.setColor(feed.color);

    created++;
  });

  Logger.log(feed.name + ': New single events created: ' + created);

  /******************************************
   * DELETE EVENTS NO LONGER IN ICS FEED
   ******************************************/
  const removedKeys = [...oldKeys].filter(k => !newKeys.has(k));
  Logger.log(feed.name + ': Keys to remove (no longer in ICS): ' + removedKeys.length);

  if (removedKeys.length > 0) {
    const removedSet = new Set(removedKeys);

    const from = new Date(now.getFullYear() - 2, 0, 1);
    const to   = new Date(now.getFullYear() + 3, 11, 31);
    const events = cal.getEvents(from, to);

    let removedCount = 0;

    events.forEach(ev => {
      const desc = ev.getDescription() || '';
      if (!desc.includes('Synced from ICS feed: ' + feed.name)) return;

      const keyLine = extractLineValue(desc, 'KEY:');
      if (!keyLine) return;

      const eventKey = keyLine.trim();

      if (removedSet.has(eventKey)) {
        Logger.log(feed.name + ' DELETE missing event key=' + eventKey + ' title=' + ev.getTitle());
        ev.deleteEvent();
        removedCount++;
      }
    });

    Logger.log(feed.name + ': Removed events from calendar: ' + removedCount);
  }

  // Persist only the keys that currently exist
  props.setProperty(feedKey, JSON.stringify([...newKeys]));
  Logger.log('--- Finished FEED: ' + feed.name + ' ---');
}


/***************************
 * HELPER FUNCTIONS
 ***************************/

/**
 * Generic ICS field reader (no params needed), e.g. UID, SUMMARY, RRULE.
 */
function getIcsField(block, fieldName) {
  const re = new RegExp(fieldName + '(?:;[^:]*)?:([^\\r\\n]+)', 'i');
  const m = block.match(re);
  if (!m) return null;

  let value = m[1].trim();

  // Handle folded lines (ICS line wrapping)
  const lines = block.split(/\r?\n/);
  const idx = lines.findIndex(l => l.toUpperCase().startsWith(fieldName.toUpperCase()));
  if (idx >= 0) {
    for (let i = idx + 1; i < lines.length; i++) {
      if (/^[ \t]/.test(lines[i])) {
        value += lines[i].trim();
      } else {
        break;
      }
    }
  }

  return value;
}

/**
 * Read a date field with parameters, e.g.
 *   DTSTART;TZID=America/Chicago:20251124T204500
 *   DTSTART:20251124T204500Z
 */
function getIcsDateField(block, fieldName) {
  const re = new RegExp(fieldName + '(?:;([^:]*))?:([^\\r\\n]+)', 'i');
  const m = block.match(re);
  if (!m) return null;

  const params = m[1] || '';
  const value  = m[2].trim();

  let tzid = null;
  let isUtc = false;

  if (params) {
    const parts = params.split(';');
    parts.forEach(p => {
      const [k, v] = p.split('=');
      if (!k || !v) return;
      if (k.toUpperCase() === 'TZID') tzid = v;
    });
  }

  if (value.endsWith('Z')) isUtc = true;

  return { value, tzid, isUtc };
}

/**
 * Decode ICS-escaped text.
 */
function decodeIcsText(t) {
  return (t || '')
    .replace(/\\,/g, ',')
    .replace(/\\n/g, '\n')
    .replace(/\\\\/g, '\\');
}

/**
 * Parse ICS date (DATE or DATE-TIME) into a JS Date,
 * honouring Z and TZID as much as Apps Script allows.
 */
function parseIcsDate(dateField) {
  if (!dateField) return null;

  let { value, tzid, isUtc } = dateField;

  // DATE ONLY (YYYYMMDD)
  if (/^\d{8}$/.test(value)) {
    return new Date(
      +value.slice(0, 4),
      +value.slice(4, 6) - 1,
      +value.slice(6, 8)
    );
  }

  if (value.endsWith('Z')) {
    isUtc = true;
    value = value.slice(0, -1);
  }

  if (!/^\d{8}T\d{6}$/.test(value)) return null;

  const y  = +value.slice(0, 4);
  const m  = +value.slice(4, 6) - 1;
  const d  = +value.slice(6, 8);
  const hh = +value.slice(9, 11);
  const mm = +value.slice(11, 13);
  const ss = +value.slice(13, 15);

  if (isUtc) {
    // Treat value as UTC instant
    return new Date(Date.UTC(y, m, d, hh, mm, ss));
  }

  if (tzid) {
    // Interpret the value in the ICS timezone, convert to script timezone
    const iso = Utilities.formatString(
      '%04d-%02d-%02dT%02d:%02d:%02d',
      y, m + 1, d, hh, mm, ss
    );
    // First, treat as UTC
    const ms = Date.parse(iso + 'Z');
    const date = new Date(ms);

    const scriptTz = Session.getScriptTimeZone();
    const offsetTarget = getOffsetMinutes(scriptTz, date);
    const offsetSource = getOffsetMinutes(tzid, date);
    const diffMs = (offsetTarget - offsetSource) * 60 * 1000;

    return new Date(ms + diffMs);
  }

  // No TZ info: treat as script/calendar local time
  return new Date(y, m, d, hh, mm, ss);
}

/**
 * Get timezone offset in minutes for a given TZ at a given date.
 */
function getOffsetMinutes(tz, date) {
  const s = Utilities.formatDate(date, tz, 'Z'); // e.g. +0530
  const sign = s.startsWith('-') ? -1 : 1;
  const hh = +s.slice(1, 3);
  const mm = +s.slice(3, 5);
  return sign * (hh * 60 + mm);
}

/**
 * Extract value from description by label prefix, e.g. "UID:" or "KEY:"
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
