/**
 * Aligner Tracker — backend (Google Sheets + Apps Script)
 *
 * Nothing personal lives in this file. Your plan lives in the Sheet (Settings,
 * Trays) and is entered in the app's first-run setup screen; the optional
 * Telegram bot token lives in Project Settings → Script properties.
 *
 * First run: select `setup` in the toolbar and click Run. It creates the tabs,
 * asks for permissions, and (if Telegram is configured) sends a test message.
 */

const SS = SpreadsheetApp.getActive();
const DAY_MS = 86400000; // a nominal day; day boundaries use real local midnights, so DST is handled

/* ------------------------------------------------------------------ */
/* Web app entry                                                       */
/* ------------------------------------------------------------------ */
function doGet() {
  ensureTabs_();
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Aligner Tracker')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, viewport-fit=cover');
}

/* ------------------------------------------------------------------ */
/* Helpers                                                             */
/* ------------------------------------------------------------------ */
function sh_(name) {
  const s = SS.getSheetByName(name);
  if (!s) throw new Error('Missing tab: ' + name);
  return s;
}
function tz_() { return SS.getSpreadsheetTimeZone(); }

function settings_() {
  const s = {};
  sh_('Settings').getDataRange().getValues().slice(1).forEach(r => {
    if (r[0] !== '' && r[0] != null) s[String(r[0]).trim()] = r[1];
  });
  return {
    target: Number(s.target_hours) || 22,
    amber: Number(s.amber_hours) || 20,
    nudge: Number(s.nudge_minutes) || 45,
    checkinHour: Number(s.checkin_hour) || 22,
    daysPerTray: Number(s.days_per_tray) || 14,
    totalTrays: Number(s.total_trays) || 15,
    mascot: String(s.mascot_name || 'Pearl').trim() || 'Pearl'
  };
}

/** 'yyyy-MM-dd' for a date, in the spreadsheet's time zone */
function key_(d) { return Utilities.formatDate(new Date(d), tz_(), 'yyyy-MM-dd'); }

/** Midnight (local) of a 'yyyy-MM-dd' key, as a Date */
function dayStart_(key) {
  const off = d => { const z = Utilities.formatDate(d, tz_(), 'Z'); return z.slice(0, 3) + ':' + z.slice(3); }; // +05:30
  const guess = new Date(key + 'T00:00:00' + off(new Date(key + 'T12:00:00Z')));
  return new Date(key + 'T00:00:00' + off(guess));   // second pass corrects daylight-saving switch days
}
function addDays_(key, n) { return key_(dayStart_(key).getTime() + n * DAY_MS + 12 * 3600000); }
function noon_(key) { return new Date(dayStart_(key).getTime() + 12 * 3600000); }

function trays_() {
  return sh_('Trays').getDataRange().getValues().slice(1)
    .map((r, i) => ({
      row: i + 2,
      no: Number(r[0]),
      start: r[1] ? key_(r[1]) : null,
      planned: Number(r[2]) || 14,
      due: r[3] ? key_(r[3]) : null,
      end: r[4] ? key_(r[4]) : null,
      eventId: r[6] || ''
    }))
    .filter(t => t.no && t.start);
}
function currentTray_(trays) {
  trays = trays || trays_();
  if (!trays.length) throw new Error('Add your current tray to the Trays tab first.');
  const open = trays.filter(t => !t.end);
  const cur = open.length ? open[open.length - 1] : trays[trays.length - 1];
  if (!cur.due) cur.due = addDays_(cur.start, cur.planned);
  return cur;
}

function sessions_() {
  const sh = sh_('Sessions'), n = sh.getLastRow();
  if (n < 2) return [];
  return sh.getRange(2, 1, n - 1, 7).getValues()
    .map((r, i) => ({
      row: i + 2,
      id: String(r[0]),
      out: r[1] ? new Date(r[1]).getTime() : null,
      inn: r[2] ? new Date(r[2]).getTime() : null
    }))
    .filter(s => s.out);
}

/* ------------------------------------------------------------------ */
/* Daily summary (wear = 24h − time out; breaks split at midnight)     */
/* ------------------------------------------------------------------ */
function summary_(sess, trays, cfg, nowMs) {
  if (!sess.length) return [];
  const todayKey = key_(nowMs);
  const firstKey = key_(Math.min.apply(null, sess.map(s => s.out)));
  const days = {};
  for (let k = firstKey; k <= todayKey; k = addDays_(k, 1)) days[k] = { out: 0, breaks: 0 };

  sess.forEach(s => {
    const end = s.inn || nowMs;
    let k = key_(s.out);
    if (days[k]) days[k].breaks++;
    while (k <= todayKey) {
      const a = dayStart_(k).getTime(), b = dayStart_(addDays_(k, 1)).getTime();
      const ov = Math.min(end, b) - Math.max(s.out, a);
      if (ov > 0 && days[k]) days[k].out += ov / 60000;
      if (end <= b) break;
      k = addDays_(k, 1);
    }
  });

  return Object.keys(days).sort().map(k => {
    const d = days[k], isToday = k === todayKey;
    const wear = isToday
      ? ((nowMs - dayStart_(k).getTime()) / 60000 - d.out) / 60
      : 24 - d.out / 60;
    const tray = trays.find(t => t.start <= k && (!t.end || k < t.end));
    const status = isToday ? 'today'
      : wear >= cfg.target ? 'green'
      : wear >= cfg.amber ? 'amber' : 'red';
    return { key: k, tray: tray ? tray.no : '', out: Math.round(d.out),
             wear: Math.round(wear * 100) / 100, breaks: d.breaks, status: status };
  });
}

function writeSummary_(rows) {
  const sh = sh_('DailySummary');
  const last = sh.getLastRow();
  if (last > 1) sh.getRange(2, 1, last - 1, 6).clearContent();
  if (rows.length) {
    sh.getRange(2, 1, rows.length, 6).setValues(
      rows.map(r => [dayStart_(r.key), r.tray, r.out, r.wear, r.breaks, r.status]));
  }
}

/* ------------------------------------------------------------------ */
/* API called by the app                                               */
/* ------------------------------------------------------------------ */
/** skipWrite = true when called from reminders (no need to rewrite DailySummary every 15 min) */
function getState(skipWrite) {
  ensureTabs_();
  const cfg = settings_(), now = Date.now();
  const sess = sessions_(), trays = trays_();
  if (!trays.length) return { needsSetup: true, now: now, settings: settingsOut_(cfg) };
  const cur = currentTray_(trays);
  const sum = summary_(sess, trays, cfg, now);
  if (skipWrite !== true) writeSummary_(sum);

  const todayKey = key_(now);
  const t0 = dayStart_(todayKey).getTime(), t1 = dayStart_(addDays_(todayKey, 1)).getTime();
  const open = sess.filter(s => !s.inn).pop();

  let closedOut = 0, breaks = 0;
  sess.forEach(s => {
    if (s.out >= t0) breaks++;
    if (!s.inn) return;
    const ov = Math.min(s.inn, t1) - Math.max(s.out, t0);
    if (ov > 0) closedOut += ov / 60000;
  });

  // Streak (with shields), grades, bests, heatmap — app adds today once the ring closes
  const hist = history_(sum, trays, cfg, now);
  const streak = hist.streak;

  const day = Math.round((t0 - dayStart_(cur.start).getTime()) / DAY_MS) + 1;
  const redDays = sum.filter(d => d.tray === cur.no && d.status === 'red').length;
  const finish = addDays_(cur.due, Math.max(0, cfg.totalTrays - cur.no) * cfg.daysPerTray);

  return {
    now: now, todayStart: t0, todayEnd: t1,
    closedOutMin: closedOut, openOutMs: open ? open.out : null, breaks: breaks,
    target: cfg.target, amber: cfg.amber, nudge: cfg.nudge, streak: streak, shields: hist.shields,
    tray: { no: cur.no, total: cfg.totalTrays, start: cur.start, due: cur.due,
            planned: cur.planned, day: day, redDays: redDays },
    finish: finish,
    history: hist,
    breakList: sess.filter(s => !s.inn || s.out >= now - 35 * DAY_MS)
      .map(s => ({ id: s.id, out: s.out, inn: s.inn })),
    settings: settingsOut_(cfg),
    hasDemo: sess.some(x => x.id.indexOf('demo-') === 0),
    sheetUrl: SS.getUrl()
  };
}

/** action: 'out' | 'in'. clientMs = tap time (server clock), so taps stay accurate. */
function toggle(action, clientMs) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const now = Date.now();
    const okClient = clientMs && Math.abs(now - clientMs) < 6 * 3600000;
    const ts = okClient ? Math.min(clientMs, now) : now;
    const sh = sh_('Sessions');
    const open = sessions_().filter(s => !s.inn).pop();

    if (action === 'out' && !open) {
      sh.appendRow([Utilities.getUuid().slice(0, 8), new Date(ts), '', '', currentTray_().no, '', '']);
    } else if (action === 'in' && open) {
      const t = Math.max(ts, open.out);
      sh.getRange(open.row, 3, 1, 2).setValues([[new Date(t), Math.round((t - open.out) / 60000)]]);
    } // a double tap is ignored: no duplicate rows
    SpreadsheetApp.flush();
  } finally {
    lock.releaseLock();
  }
  return getState();
}

function startNextTray() {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  let closed;
  try {
    const cfg = settings_(), trays = trays_(), cur = currentTray_(trays);
    if (cur.no >= cfg.totalTrays) {
      throw new Error('This is the last tray of the series. Raise total_trays in Settings if you get refinements.');
    }
    const todayKey = key_(Date.now());
    const days = summary_(sessions_(), trays, cfg, Date.now())
      .filter(d => d.tray === cur.no && d.status !== 'today');
    const avg = days.length
      ? Math.round(days.reduce((a, d) => a + d.wear, 0) / days.length * 10) / 10 : '';

    const sh = sh_('Trays');
    sh.getRange(cur.row, 5, 1, 2).setValues([[dayStart_(todayKey), avg]]);

    const no = cur.no + 1, planned = cfg.daysPerTray, due = addDays_(todayKey, planned);
    const eventId = trayEvent_(no, due, cfg.totalTrays);
    sh.appendRow([no, dayStart_(todayKey), planned, dayStart_(due), '', '', eventId, '']);
    SpreadsheetApp.flush();
    closed = { no: cur.no, avg: avg, grade: avg === '' ? '' : grade_(avg, cfg.target) };
  } finally {
    lock.releaseLock();
  }
  return { state: getState(), closed: closed };
}

/* ------------------------------------------------------------------ */
/* Calendar + push                                                     */
/* ------------------------------------------------------------------ */
/** All-day event on the due date, popup at 8 PM the evening before. */
function trayEvent_(no, dueKey, total) {
  try {
    const title = no < total ? 'Aligners: switch to tray ' + (no + 1) : 'Aligners: last tray complete';
    const ev = CalendarApp.getDefaultCalendar().createAllDayEvent(title, noon_(dueKey));
    ev.removeAllReminders();
    ev.addPopupReminder(240);
    return ev.getId();
  } catch (e) {
    console.warn('Calendar event not created: ' + e);
    return '';
  }
}

/**
 * Push via a Telegram bot (free; ntfy.sh blocks Apps Script's servers).
 * opts: { tags: 'alarm_clock' | 'tada' | 'crescent_moon' | 'tooth', priority: 'high' | 'default' }
 * Each message carries an "Open tracker" button that opens the app.
 */
const EMOJI_ = { alarm_clock: '⏰', tada: '🎉', crescent_moon: '🌙', tooth: '🦷' };

function notify_(title, message, opts) {
  opts = opts || {};
  const props = PropertiesService.getScriptProperties();
  const token = props.getProperty('TELEGRAM_TOKEN'), chat = props.getProperty('TELEGRAM_CHAT_ID');
  if (!token || !chat) {            // Telegram is optional: without it, reminders are skipped quietly
    console.log('Telegram not set up; skipped: ' + title);
    return false;
  }
  const esc = s => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
  const payload = {
    chat_id: chat,
    parse_mode: 'HTML',
    text: (EMOJI_[opts.tags] || EMOJI_.tooth) + ' <b>' + esc(title) + '</b>\n' + esc(message)
  };
  let url = '';
  try { url = ScriptApp.getService().getUrl() || ''; } catch (e) {}
  if (url) payload.reply_markup = JSON.stringify({ inline_keyboard: [[{ text: 'Open tracker', url: url }]] });
  const res = UrlFetchApp.fetch('https://api.telegram.org/bot' + token + '/sendMessage',
    { method: 'post', payload: payload, muteHttpExceptions: true });
  if (res.getResponseCode() !== 200) throw new Error('Telegram: ' + res.getContentText());
  return true;
}

/** One-time: after you've sent your bot any message, run this to link your chat. */
function connectTelegram() {
  const props = PropertiesService.getScriptProperties();
  const token = props.getProperty('TELEGRAM_TOKEN');
  if (!token) throw new Error('Add TELEGRAM_TOKEN in Project Settings → Script properties first.');
  const res = JSON.parse(UrlFetchApp.fetch('https://api.telegram.org/bot' + token + '/getUpdates',
    { muteHttpExceptions: true }).getContentText());
  if (!res.ok) throw new Error('Telegram rejected the token: ' + res.description);
  const msg = (res.result || []).map(u => u.message).filter(Boolean).pop();
  if (!msg) throw new Error('Open your bot in Telegram, send it "hi", then run connectTelegram again.');
  props.setProperty('TELEGRAM_CHAT_ID', String(msg.chat.id));
  notify_('Connected', 'Your aligner reminders will arrive here.', { tags: 'tooth' });
  console.log('Telegram linked to chat ' + msg.chat.id);
}

/* ------------------------------------------------------------------ */
/* Phase 2 — reminders (one trigger every 15 min runs all three)       */
/* ------------------------------------------------------------------ */
const NUDGE_REPEAT_MIN = 30, NUDGE_MAX = 3, TRAY_PUSH_HOUR = 9;

/** Run once from the editor to switch reminders on (safe to re-run). */
function installReminders() {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'remindersTick')
    .forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('remindersTick').timeBased().everyMinutes(15).create();
  console.log('Reminders on: checks every 15 min.');
}

/** Run once to switch reminders off. */
function removeReminders() {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'remindersTick')
    .forEach(t => ScriptApp.deleteTrigger(t));
  console.log('Reminders off.');
}

function remindersTick() {
  const cfg = settings_(), s = getState(true), now = Date.now();
  if (s.needsSetup) return;
  const todayKey = key_(now);
  const hour = Number(Utilities.formatDate(new Date(now), tz_(), 'H'));
  const log = sentLog_(todayKey);

  // 1. Out too long — at the threshold, then every 30 min, max 3 per break
  if (s.openOutMs) {
    const mins = (now - s.openOutMs) / 60000;
    const n = log.nudgeFor === s.openOutMs ? log.nudges : 0;
    if (n < NUDGE_MAX && mins >= cfg.nudge + n * NUDGE_REPEAT_MIN) {
      nudgeMsg_(s, mins);
      log.nudgeFor = s.openOutMs; log.nudges = n + 1;
    }
  }

  // 2. Tray change due — first run after 9 AM on or after the due date
  if (!log.tray && hour >= TRAY_PUSH_HOUR && todayKey >= s.tray.due) {
    trayMsg_(s, todayKey);
    log.tray = true;
  }

  // 3. Nightly check-in — first run after the check-in hour
  if (!log.checkin && hour >= cfg.checkinHour) {
    checkinMsg_(s, now);
    log.checkin = true;
  }

  saveLog_(log);
}

/* ---- messages ---- */
function nudgeMsg_(s, mins) {
  const budget = budgetLeft_(s, Date.now());
  notify_('Aligners out for ' + fmtMin_(mins),
    budget >= 0 ? 'Pop them back in. Break budget left today: ' + fmtMin_(budget) + '.'
                : 'Pop them back in. You are ' + fmtMin_(-budget) + ' over today\'s break budget.',
    { tags: 'alarm_clock', priority: 'high' });
}

function trayMsg_(s, todayKey) {
  const t = s.tray;
  const late = Math.round((dayStart_(todayKey) - dayStart_(t.due)) / DAY_MS);
  if (t.no >= t.total) {
    notify_('Last tray done', 'Tray ' + t.no + ' of ' + t.total + ' is complete. Check in with your orthodontist on next steps.',
      { tags: 'tada' });
  } else {
    notify_(late > 0 ? 'Tray ' + (t.no + 1) + ' is ' + late + ' day' + (late > 1 ? 's' : '') + ' overdue'
                     : 'Tray change day',
      'Switch to tray ' + (t.no + 1) + ' today, then tap Start tray ' + (t.no + 1) + ' in the app.',
      { tags: 'tada', priority: 'high' });
  }
}

function checkinMsg_(s, now) {
  const worn = wornMin_(s, now), target = s.target * 60;
  const leftToday = (s.todayEnd - now) / 60000;
  const best = worn + (s.openOutMs ? 0 : leftToday); // most you can still reach
  let title, body;
  if (worn >= target) {
    title = 'Ring closed'; body = fmtMin_(worn) + ' worn today. Your teeth say thanks.';
  } else if (s.openOutMs) {
    title = 'They\'re out right now'; body = fmtMin_(worn) + ' worn so far. Pop them in to close today\'s ring.';
  } else if (best >= target) {
    title = 'On track'; body = fmtMin_(worn) + ' worn. Keep them in and the ring closes tonight.';
  } else {
    title = fmtMin_(worn) + ' worn today';
    body = 'The ring won\'t close today (' + fmtMin_(target - best) + ' short). Keep them in overnight; tomorrow is a fresh start.';
  }
  notify_(title, body, { tags: 'crescent_moon' });
}

/* ---- maths shared with the app ---- */
function outTodayMin_(s, now) {
  return s.closedOutMin + (s.openOutMs ? (now - Math.max(s.openOutMs, s.todayStart)) / 60000 : 0);
}
function wornMin_(s, now) { return Math.max(0, (now - s.todayStart) / 60000 - outTodayMin_(s, now)); }
function budgetLeft_(s, now) { return (24 - s.target) * 60 - outTodayMin_(s, now); }
function fmtMin_(m) {
  m = Math.max(0, Math.round(m));
  const h = Math.floor(m / 60), r = m % 60;
  return h ? h + 'h ' + r + 'm' : r + ' min';
}

/* ---- once-a-day bookkeeping, so each reminder fires once ---- */
function sentLog_(todayKey) {
  let log = {};
  try { log = JSON.parse(PropertiesService.getScriptProperties().getProperty('sentLog') || '{}'); } catch (e) {}
  if (log.day !== todayKey) log = { day: todayKey, tray: false, checkin: false, nudgeFor: log.nudgeFor || null, nudges: log.nudges || 0 };
  return log;
}
function saveLog_(log) { PropertiesService.getScriptProperties().setProperty('sentLog', JSON.stringify(log)); }

/** Gate check: sends one sample of each reminder right now, using your live data. */
function testReminders() {
  if (!PropertiesService.getScriptProperties().getProperty('TELEGRAM_CHAT_ID'))
    throw new Error('Set up Telegram first: add TELEGRAM_TOKEN in Script properties, message your bot, run connectTelegram.');
  const s = getState(true), now = Date.now();
  if (s.needsSetup) throw new Error('Open the app and finish the setup screen first.');
  nudgeMsg_(s, settings_().nudge);
  trayMsg_(s, s.tray.due);
  checkinMsg_(s, now);
  console.log('Sent 3 test pushes. Check your phone.');
}

/* ------------------------------------------------------------------ */
/* One-time setup — run from the editor                                */
/* ------------------------------------------------------------------ */
function setup() {
  ensureTabs_();
  CalendarApp.getDefaultCalendar();          // touching each service makes Google ask for access now
  ScriptApp.getProjectTriggers();
  const trays = trays_();
  if (trays.length) {                         // already set up: make sure the current tray has its event
    const cfg = settings_(), cur = currentTray_(trays);
    if (!cur.eventId) {
      const id = trayEvent_(cur.no, cur.due, cfg.totalTrays);
      sh_('Trays').getRange(cur.row, 4, 1, 4).setValues([[dayStart_(cur.due), '', '', id]]);
    }
  }
  const sent = notify_('Aligner Tracker connected', 'Test message from your tracker.');
  console.log('Setup done. Tabs ready; time zone %s. Telegram: %s. Next: deploy as a web app and open it.',
              tz_(), sent ? 'test message sent' : 'not set up (optional)');
}

/* ------------------------------------------------------------------ */
/* First-run: tabs + the app's setup screen                            */
/* ------------------------------------------------------------------ */
const TABS_ = {
  Settings: ['key', 'value', 'what it means'],
  Trays: ['tray_no', 'start_date', 'planned_days', 'due_date', 'actual_end', 'avg_wear_h', 'calendar_event_id', 'notes'],
  Sessions: ['session_id', 'out_ts', 'in_ts', 'minutes_out', 'tray_no', 'tag', 'edited'],
  DailySummary: ['date', 'tray_no', 'out_min', 'wear_h', 'breaks', 'status'],
  Appointments: ['date', 'type', 'notes', 'trays_added']
};
const DEFAULTS_ = [
  ['target_hours', 22, 'Daily wear target (hours)'],
  ['amber_hours', 20, 'At least this but under target = amber; below = red'],
  ['nudge_minutes', 45, 'Break length that triggers the out-too-long reminder'],
  ['checkin_hour', 22, 'Hour of the nightly check-in (24h clock)'],
  ['days_per_tray', 14, 'Default days per tray'],
  ['total_trays', 15, 'Trays in your current series'],
  ['mascot_name', 'Pearl', 'Your tooth buddy\'s name']
];

/** Creates any missing tab with its headers. Safe to call any time. */
function ensureTabs_() {
  let created = false;
  Object.keys(TABS_).forEach(name => {
    if (SS.getSheetByName(name)) return;
    const sh = SS.insertSheet(name), head = TABS_[name];
    sh.getRange(1, 1, 1, head.length).setValues([head]).setFontWeight('bold');
    sh.setFrozenRows(1);
    if (name === 'Settings') sh.getRange(2, 1, DEFAULTS_.length, 3).setValues(DEFAULTS_);
    if (name === 'Trays') sh.getRange('B:E').setNumberFormat('dd-mmm-yyyy');
    if (name === 'Sessions') sh.getRange('B:C').setNumberFormat('dd-mmm-yyyy hh:mm');
    if (name === 'DailySummary') sh.getRange('A:A').setNumberFormat('dd-mmm-yyyy');
    created = true;
  });
  if (created) {
    const blank = SS.getSheets().find(x => !TABS_[x.getName()] && x.getLastRow() === 0);
    if (blank) SS.deleteSheet(blank);         // tidy away the empty default sheet
    SpreadsheetApp.flush();
  }
}

function settingsOut_(cfg) {
  return { target_hours: cfg.target, amber_hours: cfg.amber, nudge_minutes: cfg.nudge,
           checkin_hour: cfg.checkinHour, days_per_tray: cfg.daysPerTray, total_trays: cfg.totalTrays,
           mascot_name: cfg.mascot };
}

/** Called by the app's first-run screen. */
function completeSetup(obj) {
  const total = Number(obj.total_trays), no = Number(obj.current_tray);
  const days = Number(obj.days_per_tray), target = Number(obj.target_hours);
  const start = String(obj.start_date || '');
  if (!(total >= 1 && total <= 99)) throw new Error('Trays in the series: a number from 1 to 99.');
  if (!(no >= 1 && no <= total)) throw new Error('Current tray must be between 1 and ' + total + '.');
  if (!(days >= 5 && days <= 30)) throw new Error('Days per tray: a number from 5 to 30.');
  if (!(target >= 12 && target <= 24)) throw new Error('Daily target: between 12 and 24 hours.');
  if (!/^\d{4}-\d{2}-\d{2}$/.test(start)) throw new Error('Pick the date you started this tray.');
  if (start > key_(Date.now())) throw new Error('The start date can\'t be in the future.');

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    ensureTabs_();
    if (trays_().length) throw new Error('Setup is already done. Change things in Settings or the Trays tab.');
    const due = addDays_(start, days);
    const eventId = trayEvent_(no, due, total);
    sh_('Trays').appendRow([no, dayStart_(start), days, dayStart_(due), '', '', eventId, '']);
    SpreadsheetApp.flush();
  } finally {
    lock.releaseLock();
  }
  return saveSettings({
    target_hours: target, amber_hours: Math.min(settings_().amber, target - 1),
    days_per_tray: days, total_trays: total, mascot_name: obj.mascot_name || 'Pearl'
  });
}

/* ------------------------------------------------------------------ */
/* Phase 3 — history, grades, shields, bests                           */
/* ------------------------------------------------------------------ */
function grade_(avg, target) { return avg >= target ? 'A' : avg >= target - 1 ? 'B' : 'C'; }

function history_(sum, trays, cfg, nowMs) {
  const done = sum.filter(d => d.status !== 'today');

  // Streak with shields: 1 shield per 7 green days (max 2); a shield absorbs one amber day
  let streak = 0, longest = 0, shields = 0, greens = 0, earnedTotal = 0, lastSaved = null;
  const saved = {};
  done.forEach(d => {
    if (d.status === 'green') {
      streak++; greens++;
      if (greens % 7 === 0 && shields < 2) { shields++; earnedTotal++; }
    } else if (d.status === 'amber' && shields > 0) {
      shields--; saved[d.key] = true; lastSaved = { key: d.key, streak: streak };
    } else {
      streak = 0; greens = 0;
    }
    if (streak > longest) longest = streak;
  });

  // Heatmap: the last 35 days ending today
  const todayKey = key_(nowMs), byKey = {};
  sum.forEach(d => { byKey[d.key] = d; });
  const days = [];
  for (let i = 34; i >= 0; i--) {
    const k = addDays_(todayKey, -i), d = byKey[k];
    days.push(d ? { key: k, wear: d.wear, out: d.out, breaks: d.breaks, status: d.status, saved: !!saved[k] }
                : { key: k, status: 'none' });
  }

  // Per-tray compliance (completed days only)
  const trayStats = trays.map(t => {
    const ds = done.filter(d => d.tray === t.no);
    if (!ds.length) return null;
    const avg = Math.round(ds.reduce((a, d) => a + d.wear, 0) / ds.length * 10) / 10;
    return { no: t.no, avg: avg, days: ds.length, grade: grade_(avg, cfg.target), current: !t.end,
             start: t.start, end: t.end };
  }).filter(Boolean);

  // Last 30 completed days
  const last = done.slice(-30);
  const outSum = last.reduce((a, d) => a + d.out, 0), brSum = last.reduce((a, d) => a + d.breaks, 0);
  const stats = last.length ? {
    days: last.length,
    avgWear: Math.round(last.reduce((a, d) => a + d.wear, 0) / last.length * 10) / 10,
    greenPct: Math.round(last.filter(d => d.status === 'green').length / last.length * 100),
    avgBreak: brSum ? Math.round(outSum / brSum) : 0
  } : null;

  const bestDay = done.reduce((b, d) => (!b || d.wear > b.wear) ? d : b, null);
  const bestTray = trayStats.filter(t => t.days >= 3).reduce((b, t) => (!b || t.avg > b.avg) ? t : b, null);

  return {
    streak: streak, shields: shields, days: days, trays: trayStats, stats: stats,
    brace: {                                   // Brace, the shield buddy
      shields: shields, max: 2, every: 7,
      progress: shields >= 2 ? 7 : greens % 7, // on-target days toward the next shield
      earnedTotal: earnedTotal,                // app compares with what it last saw → "shield earned"
      lastSaved: lastSaved                     // most recent amber day a shield saved → "Blocked!" modal
    },
    bests: {
      longest: longest,
      bestDay: bestDay ? { key: bestDay.key, wear: bestDay.wear } : null,
      bestTray: bestTray ? { no: bestTray.no, avg: bestTray.avg, grade: bestTray.grade } : null
    }
  };
}

function trayFor_(key, trays) {
  trays = trays || trays_();
  const t = trays.find(x => x.start <= key && (!x.end || key < x.end));
  return t ? t.no : currentTray_(trays).no;
}

/* ------------------------------------------------------------------ */
/* Phase 3 — edit / add / delete breaks from the app                   */
/* ------------------------------------------------------------------ */
/** b = { id?: string, out: ms, inn: ms | null } — inn null only keeps a break open */
function saveBreak(b) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const now = Date.now();
    const out = Number(b.out), inn = b.inn ? Number(b.inn) : null;
    if (!out || out > now + 60000) throw new Error('Out time can\'t be in the future.');
    if (inn !== null && inn <= out) throw new Error('In time must be after the out time.');
    if (inn !== null && inn > now + 60000) throw new Error('In time can\'t be in the future.');
    if ((inn || now) - out > DAY_MS) throw new Error('A break can\'t be longer than 24 hours.');

    const sess = sessions_();
    const self = b.id ? sess.find(s => s.id === String(b.id)) : null;
    if (b.id && !self) throw new Error('That break no longer exists. Pull to refresh and try again.');
    if (inn === null && sess.some(s => !s.inn && s !== self)) throw new Error('Another break is already open.');
    const clash = sess.find(s => s !== self && out < (s.inn || now) && (inn || now) > s.out);
    if (clash) throw new Error('Overlaps the break that started ' +
      Utilities.formatDate(new Date(clash.out), tz_(), 'd MMM, HH:mm') + '.');

    const sh = sh_('Sessions');
    const mins = inn ? Math.round((inn - out) / 60000) : '';
    const tray = trayFor_(key_(out));
    const row = self ? self.row : sh.getLastRow() + 1;
    if (!self) sh.getRange(row, 1).setValue('m-' + Utilities.getUuid().slice(0, 6));
    sh.getRange(row, 2, 1, 4).setValues([[new Date(out), inn ? new Date(inn) : '', mins, tray]]);
    sh.getRange(row, 7).setValue('yes');
    SpreadsheetApp.flush();
  } finally {
    lock.releaseLock();
  }
  return getState();
}

function deleteBreak(id) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const s = sessions_().find(x => x.id === String(id));
    if (!s) throw new Error('That break no longer exists.');
    sh_('Sessions').deleteRow(s.row);
    SpreadsheetApp.flush();
  } finally {
    lock.releaseLock();
  }
  return getState();
}

/* ------------------------------------------------------------------ */
/* Phase 3 — settings from the app                                     */
/* ------------------------------------------------------------------ */
const LIMITS_ = {
  target_hours: [12, 24], amber_hours: [10, 24], nudge_minutes: [5, 240],
  checkin_hour: [0, 23], days_per_tray: [5, 30], total_trays: [1, 99]
};

function saveSettings(obj) {
  const clean = {};
  Object.keys(LIMITS_).forEach(k => {
    if (obj[k] === undefined || obj[k] === '' || obj[k] === null) return;
    const v = Number(obj[k]), lo = LIMITS_[k][0], hi = LIMITS_[k][1];
    if (!(v >= lo && v <= hi)) throw new Error(k.replace(/_/g, ' ') + ' must be between ' + lo + ' and ' + hi + '.');
    clean[k] = v;
  });
  const cur = settings_();
  const target = clean.target_hours !== undefined ? clean.target_hours : cur.target;
  const amber = clean.amber_hours !== undefined ? clean.amber_hours : cur.amber;
  if (amber >= target) throw new Error('Amber hours must be below the target.');
  if (clean.total_trays !== undefined && clean.total_trays < currentTray_().no)
    throw new Error('Total trays can\'t be below your current tray.');

  if (obj.mascot_name !== undefined && obj.mascot_name !== null) {
    const name = String(obj.mascot_name).trim();
    if (!/^[\p{L}\p{N} '.-]{1,16}$/u.test(name)) throw new Error('Mascot name: 1–16 letters or numbers.');
    clean.mascot_name = name;
  }

  const sh = sh_('Settings'), vals = sh.getDataRange().getValues();
  Object.keys(clean).forEach(k => {
    const i = vals.findIndex(r => String(r[0]).trim() === k);
    if (i > 0) sh.getRange(i + 1, 2).setValue(clean[k]); else sh.appendRow([k, clean[k]]);
  });
  SpreadsheetApp.flush();
  return getState();
}

/* ------------------------------------------------------------------ */
/* Demo data — run seedDemoData from the editor; clear from the app's  */
/* Settings screen or by running clearDemoData                         */
/* ------------------------------------------------------------------ */
function seedDemoData() {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  let count = 0;
  try {
    const sess = sessions_();
    if (sess.some(s => s.id.indexOf('demo-') === 0)) throw new Error('Demo data is already in. Run clearDemoData first.');
    const trays = trays_(), first = trays[0];
    if (!first) throw new Error('Open the app and finish the setup screen first; demo data builds on your current tray.');

    // Two finished demo trays before your first real tray
    let startKey = first.start;
    if (first.no > 2) {
      const t2 = addDays_(first.start, -first.planned), t1 = addDays_(t2, -first.planned);
      const tsh = sh_('Trays');
      tsh.insertRowsBefore(2, 2);
      tsh.getRange(2, 1, 2, 8).setValues([
        [first.no - 2, dayStart_(t1), first.planned, dayStart_(t2), dayStart_(t2), '', '', 'demo'],
        [first.no - 1, dayStart_(t2), first.planned, dayStart_(first.start), dayStart_(first.start), '', '', 'demo']
      ]);
      startKey = t1;
    }

    // Meals every day (plus some coffees) up to yesterday; sloppier early on, better over time
    const allTrays = trays_();
    const firstReal = sess.length ? Math.min.apply(null, sess.map(s => s.out)) : Infinity;
    const endKey = addDays_(key_(Date.now()), -1);
    let seed = 7;
    const rnd = () => (seed = (seed * 16807) % 2147483647) / 2147483647;
    const rows = [];
    let i = 0;
    for (let k = startKey; k <= endKey; k = addDays_(k, 1), i++) {
      const base = dayStart_(k).getTime(), sloppy = Math.max(0.12, 0.6 - i * 0.016);
      const plan = [[8.5, 15, 12], [13.5, 25, 15], [20.5, 25, 20]];   // [hour, min length, spread]
      if (rnd() < 0.5) plan.push([16.75, 8, 8]);                        // coffee
      plan.forEach(p => {
        let len = p[1] + rnd() * p[2];
        if (rnd() < sloppy * 0.35) len += 40 + rnd() * 90;            // the occasional long dinner
        const out = base + (p[0] + (rnd() - 0.5) * 0.8) * 3600000, inn = out + len * 60000;
        if (inn >= firstReal) return;
        rows.push(['demo-' + Utilities.getUuid().slice(0, 6), new Date(out), new Date(inn),
                   Math.round(len), trayFor_(k, allTrays), '', '']);
      });
    }
    const sh = sh_('Sessions');
    if (rows.length) sh.getRange(sh.getLastRow() + 1, 1, rows.length, 7).setValues(rows);
    count = rows.length;
    SpreadsheetApp.flush();
  } finally {
    lock.releaseLock();
  }
  getState();
  console.log('Added %s demo breaks. Remove them later with clearDemoData (or the button in Settings).', count);
}

function clearDemoData() {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  let s = 0, t = 0;
  try {
    const strip = (name, keep) => {
      const sh = sh_(name), n = sh.getLastRow();
      if (n < 2) return 0;
      const w = sh.getLastColumn(), vals = sh.getRange(2, 1, n - 1, w).getValues(), kept = vals.filter(keep);
      sh.getRange(2, 1, n - 1, w).clearContent();
      if (kept.length) sh.getRange(2, 1, kept.length, w).setValues(kept);
      return vals.length - kept.length;
    };
    s = strip('Sessions', r => String(r[0]).indexOf('demo-') !== 0);
    t = strip('Trays', r => String(r[7]) !== 'demo');
    SpreadsheetApp.flush();
  } finally {
    lock.releaseLock();
  }
  console.log('Removed %s demo breaks and %s demo trays.', s, t);
  return getState();
}
