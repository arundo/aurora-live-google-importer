/**
 * Aurora Live → Google Calendar invite importer
 *
 * Aurora Live sends its 1-to-1 meeting invites from Outlook with a hidden
 * attendee list, so the iCalendar part contains no ATTENDEE line and Gmail
 * does not treat the message as an invitation. This script finds those
 * messages, extracts the text/calendar part and imports the event into your
 * calendar, keeping the original UID so that later updates and cancellations
 * from Aurora Live are applied to the same event.
 *
 * Setup: see the instructions that came with this file. In short:
 *   1. Paste into a new Apps Script project.
 *   2. Services (+) → add "Google Calendar API" (identifier: Calendar).
 *   3. Run importAuroraInvites once with DRY_RUN = true and check the log.
 *   4. Set DRY_RUN = false, run once more, then run installTrigger.
 */

const CONFIG = {
  // Which emails to scan. Narrow on purpose: only Aurora Live's Outlook calendar sender.
  SEARCH_QUERY: 'from:calendar@auroralive.com newer_than:45d',
  CALENDAR_ID: 'primary',
  // Gmail label added to processed threads, for visibility. Set to '' to skip.
  LABEL_NAME: 'Aurora Live/Imported',
  // true = only log what would happen; nothing is written to Calendar or Gmail.
  DRY_RUN: true,
  MAX_THREADS: 50,
  // How often the trigger runs (minutes: 1, 5, 10, 15 or 30).
  TRIGGER_MINUTES: 10,
};

// Outlook uses Windows time zone names; Google Calendar needs IANA names.
const WINDOWS_TZ = {
  'W. Europe Standard Time': 'Europe/Berlin',
  'Romance Standard Time': 'Europe/Paris',
  'Central Europe Standard Time': 'Europe/Budapest',
  'Central European Standard Time': 'Europe/Warsaw',
  'FLE Standard Time': 'Europe/Helsinki',
  'GTB Standard Time': 'Europe/Bucharest',
  'E. Europe Standard Time': 'Europe/Chisinau',
  'GMT Standard Time': 'Europe/London',
  'Greenwich Standard Time': 'Atlantic/Reykjavik',
  'UTC': 'UTC',
  'Eastern Standard Time': 'America/New_York',
  'Central Standard Time': 'America/Chicago',
  'Pacific Standard Time': 'America/Los_Angeles',
  'India Standard Time': 'Asia/Kolkata',
  'Singapore Standard Time': 'Asia/Singapore',
};

/** Main entry point – run manually or from the time-driven trigger. */
function importAuroraInvites() {
  const props = PropertiesService.getUserProperties();
  const processed = JSON.parse(props.getProperty('processedIds') || '[]');
  const seen = new Set(processed);
  const label = (!CONFIG.DRY_RUN && CONFIG.LABEL_NAME)
    ? (GmailApp.getUserLabelByName(CONFIG.LABEL_NAME) || GmailApp.createLabel(CONFIG.LABEL_NAME))
    : null;

  // Collect unprocessed messages, oldest first, so later updates win.
  const messages = [];
  GmailApp.search(CONFIG.SEARCH_QUERY, 0, CONFIG.MAX_THREADS).forEach(thread =>
    thread.getMessages().forEach(m => { if (!seen.has(m.getId())) messages.push(m); }));
  messages.sort((a, b) => a.getDate() - b.getDate());
  console.log(`${messages.length} new message(s) to check.`);

  messages.forEach(msg => {
    try {
      const calendars = extractCalendarParts_(msg.getRawContent());
      if (!calendars.length) console.log(`No calendar data in "${msg.getSubject()}" – skipped.`);
      calendars.forEach(ics => parseIcs_(ics).forEach(applyEvent_));
      if (!CONFIG.DRY_RUN) {
        processed.push(msg.getId());
        if (label) msg.getThread().addLabel(label);
      }
    } catch (e) {
      // Not marked as processed, so it is retried on the next run.
      console.error(`Failed on "${msg.getSubject()}" (${msg.getId()}): ${e.stack || e}`);
    }
  });

  // Keep the list small enough for one user property (9 KB limit).
  if (!CONFIG.DRY_RUN) props.setProperty('processedIds', JSON.stringify(processed.slice(-300)));
}

/** Creates (or replaces) the time-driven trigger. Run once after setup. */
function installTrigger() {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'importAuroraInvites')
    .forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('importAuroraInvites').timeBased().everyMinutes(CONFIG.TRIGGER_MINUTES).create();
  console.log(`Trigger installed: every ${CONFIG.TRIGGER_MINUTES} minutes.`);
}

/** Removes the trigger (stops the automation). */
function uninstallTrigger() {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'importAuroraInvites')
    .forEach(t => ScriptApp.deleteTrigger(t));
  console.log('Trigger removed.');
}

/** Forgets which messages were processed, so the next run re-checks everything in the search window. */
function resetProcessed() {
  PropertiesService.getUserProperties().deleteProperty('processedIds');
  console.log('Processed-message list cleared.');
}

// ---------------------------------------------------------------- Calendar

function applyEvent_(ev) {
  const cal = CONFIG.CALENDAR_ID;
  const found = Calendar.Events.list(cal, { iCalUID: ev.uid }).items || [];
  const existing = found.find(e => e.status !== 'cancelled');
  const when = ev.start.dateTime || ev.start.date;

  if (ev.cancelled) {
    if (!existing) return console.log(`Cancel for "${ev.summary}" – not in calendar, nothing to do.`);
    console.log(`${CONFIG.DRY_RUN ? '[dry run] would remove' : 'Removing'} "${ev.summary}" (${when}).`);
    if (!CONFIG.DRY_RUN) Calendar.Events.remove(cal, existing.id);
    return;
  }

  // Skip if the calendar already holds this or a newer version from Aurora Live.
  const priv = (existing && existing.extendedProperties && existing.extendedProperties.private) || {};
  if (existing && priv.auroraDtstamp && priv.auroraDtstamp >= ev.dtstamp) {
    return console.log(`"${ev.summary}" is up to date – skipped.`);
  }

  const body = {
    summary: ev.summary,
    description: ev.description,
    location: ev.joinUrl || ev.location,
    start: ev.start,
    end: ev.end,
    status: 'confirmed',
    transparency: 'opaque',
    extendedProperties: { private: {
      auroraSource: 'gmail-ics-import',
      auroraDtstamp: ev.dtstamp,
      auroraSequence: String(ev.sequence),
    } },
  };

  if (existing) {
    console.log(`${CONFIG.DRY_RUN ? '[dry run] would update' : 'Updating'} "${ev.summary}" (${when}).`);
    if (!CONFIG.DRY_RUN) Calendar.Events.patch(body, cal, existing.id);
  } else {
    console.log(`${CONFIG.DRY_RUN ? '[dry run] would add' : 'Adding'} "${ev.summary}" (${when}).`);
    if (!CONFIG.DRY_RUN) {
      Calendar.Events.import(Object.assign({
        iCalUID: ev.uid,
        organizer: ev.organizer,
        reminders: { useDefault: true },
      }, body), cal);
    }
  }
}

// ---------------------------------------------------------------- iCalendar

function parseIcs_(text) {
  const lines = text.replace(/\r\n/g, '\n').replace(/\n[ \t]/g, '').split('\n');
  const stack = [];
  const events = [];
  let method = '';
  let cur = null;

  lines.forEach(line => {
    if (!line.trim()) return;
    const p = parseLine_(line);
    if (p.name === 'BEGIN') {
      stack.push(p.value.toUpperCase());
      if (p.value.toUpperCase() === 'VEVENT') cur = {};
      return;
    }
    if (p.name === 'END') {
      const ended = stack.pop();
      if (ended === 'VEVENT' && cur) { events.push(cur); cur = null; }
      return;
    }
    const where = stack[stack.length - 1];
    if (where === 'VCALENDAR' && p.name === 'METHOD') method = p.value.trim().toUpperCase();
    if (where === 'VEVENT' && cur && !(p.name in cur)) cur[p.name] = p;
  });

  return events.map(e => toEvent_(e, method)).filter(Boolean);
}

function parseLine_(line) {
  let inQuotes = false;
  let idx = -1;
  for (let i = 0; i < line.length; i++) {
    if (line[i] === '"') inQuotes = !inQuotes;
    else if (line[i] === ':' && !inQuotes) { idx = i; break; }
  }
  const head = idx < 0 ? line : line.slice(0, idx);
  const value = idx < 0 ? '' : line.slice(idx + 1);
  const parts = head.split(';');
  const params = {};
  parts.slice(1).forEach(s => {
    const eq = s.indexOf('=');
    if (eq > 0) params[s.slice(0, eq).toUpperCase()] = s.slice(eq + 1).replace(/^"|"$/g, '');
  });
  return { name: parts[0].toUpperCase(), params, value };
}

function toEvent_(e, method) {
  if (!e.UID || !e.DTSTART) return null;
  // Ignore replies/counter-proposals; only act on invites, publishes and cancellations.
  if (method && ['REQUEST', 'PUBLISH', 'CANCEL'].indexOf(method) < 0) return null;

  const text = k => (e[k] ? unescapeText_(e[k].value) : '');
  const description = text('DESCRIPTION');
  const start = toGoogleTime_(e.DTSTART);
  const end = e.DTEND ? toGoogleTime_(e.DTEND) : start;
  const orgEmail = e.ORGANIZER ? e.ORGANIZER.value.replace(/^mailto:/i, '').trim() : '';
  const join = description.match(/Join:\s*(https:\/\/teams\.microsoft\.com\/\S+)/i);
  const teamsProp = e['X-MICROSOFT-SKYPETEAMSMEETINGURL'];

  return {
    uid: e.UID.value.trim(),
    summary: text('SUMMARY'),
    description,
    location: text('LOCATION'),
    start,
    end,
    dtstamp: e.DTSTAMP ? e.DTSTAMP.value.trim() : '',
    sequence: e.SEQUENCE ? parseInt(e.SEQUENCE.value, 10) || 0 : 0,
    cancelled: method === 'CANCEL' || !!(e.STATUS && e.STATUS.value.trim().toUpperCase() === 'CANCELLED'),
    organizer: orgEmail ? { email: orgEmail, displayName: e.ORGANIZER.params.CN } : undefined,
    joinUrl: join ? join[1] : (teamsProp ? teamsProp.value.trim() : ''),
  };
}

function toGoogleTime_(prop) {
  const m = prop.value.trim().match(/^(\d{4})(\d{2})(\d{2})(?:T(\d{2})(\d{2})(\d{2})(Z)?)?$/);
  if (!m) throw new Error('Cannot parse date/time: ' + prop.value);
  if (!m[4]) return { date: `${m[1]}-${m[2]}-${m[3]}` };
  const local = `${m[1]}-${m[2]}-${m[3]}T${m[4]}:${m[5]}:${m[6]}`;
  if (m[7]) return { dateTime: local + 'Z', timeZone: 'UTC' };

  const tzid = prop.params.TZID;
  let tz = tzid && (WINDOWS_TZ[tzid] || (tzid.indexOf('/') > 0 ? tzid : null));
  if (!tz) {
    tz = Session.getScriptTimeZone();
    if (tzid) console.warn(`Unknown time zone "${tzid}" – using ${tz}. Add it to WINDOWS_TZ.`);
  }
  return { dateTime: local, timeZone: tz };
}

function unescapeText_(v) {
  return v.replace(/\\n/gi, '\n').replace(/\\([,;\\])/g, '$1');
}

// ---------------------------------------------------------------- MIME

function extractCalendarParts_(raw) {
  const out = [];
  walkMime_(raw, out, 0);
  return out;
}

function walkMime_(entity, out, depth) {
  if (depth > 10) return;
  const sep = entity.search(/\r?\n\r?\n/);
  if (sep < 0) return;
  const headerText = entity.slice(0, sep).replace(/\r?\n[ \t]+/g, ' ');
  const body = entity.slice(sep).replace(/^\r?\n\r?\n/, '');

  const headers = {};
  headerText.split(/\r?\n/).forEach(l => {
    const i = l.indexOf(':');
    if (i > 0) {
      const k = l.slice(0, i).trim().toLowerCase();
      if (!(k in headers)) headers[k] = l.slice(i + 1).trim();
    }
  });

  const ctypeRaw = headers['content-type'] || 'text/plain';
  const ctype = ctypeRaw.toLowerCase();

  if (ctype.indexOf('multipart/') === 0) {
    const b = ctypeRaw.match(/boundary="?([^";\r\n]+)"?/i);
    if (!b) return;
    body.split('--' + b[1]).slice(1).forEach(part => {
      if (part.indexOf('--') === 0) return; // closing delimiter
      walkMime_(part.replace(/^[ \t]*\r?\n/, ''), out, depth + 1);
    });
  } else if (ctype.indexOf('text/calendar') === 0 || ctype.indexOf('application/ics') === 0) {
    const cs = ctypeRaw.match(/charset="?([^";\s]+)"?/i);
    out.push(decodeBody_(body, headers['content-transfer-encoding'], cs ? cs[1] : 'UTF-8'));
  }
}

function decodeBody_(body, cte, charset) {
  const enc = (cte || '').trim().toLowerCase();
  if (enc === 'base64') {
    const bytes = Utilities.base64Decode(body.replace(/[^A-Za-z0-9+/=]/g, ''));
    return Utilities.newBlob(bytes).getDataAsString(charset);
  }
  if (enc === 'quoted-printable') {
    const s = body.replace(/=\r?\n/g, '');
    const bytes = [];
    for (let i = 0; i < s.length; i++) {
      if (s[i] === '=' && /^[0-9A-Fa-f]{2}$/.test(s.substr(i + 1, 2))) {
        const b = parseInt(s.substr(i + 1, 2), 16);
        bytes.push(b > 127 ? b - 256 : b);
        i += 2;
      } else {
        const c = s.charCodeAt(i);
        bytes.push(c > 127 ? c - 256 : c);
      }
    }
    return Utilities.newBlob(bytes).getDataAsString(charset);
  }
  return body;
}
