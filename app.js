/* =====================================================================
   ប្រព័ន្ធគ្រប់គ្រងវត្តមាន · app.js
   ផ្នែក៖ 1 ជំនួយ · 2 ទិន្នន័យ · 3 មើលឃើញ (views) · 4 សកម្មភាព · 5 Excel · 6 ចាប់ផ្តើម
   ===================================================================== */
(function () {
'use strict';

const CFG = window.APP_CONFIG || {};
const $ = (s, r = document) => r.querySelector(s);
const $$ = (s, r = document) => Array.from(r.querySelectorAll(s));
const app = document.getElementById('app');

/* ---------- 1. ជំនួយ ---------- */
const KHD = '០១២៣៤៥៦៧៨៩';
const kh = n => String(n).replace(/\d/g, d => KHD[d]);
const unkh = s => String(s == null ? '' : s).replace(/[០-៩]/g, d => KHD.indexOf(d)).trim();
const DAYS = ['ចន្ទ', 'អង្គារ', 'ពុធ', 'ព្រហស្បតិ៍', 'សុក្រ'];
const MONTHS = ['មករា', 'កុម្ភៈ', 'មីនា', 'មេសា', 'ឧសភា', 'មិថុនា', 'កក្កដា', 'សីហា', 'កញ្ញា', 'តុលា', 'វិច្ឆិកា', 'ធ្នូ'];
const SESS = { am: 'ពេលព្រឹក', pm: 'ពេលរសៀល' };
const SESS_S = { am: 'ព្រឹក', pm: 'រសៀល' };
const esc = s => String(s == null ? '' : s).replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
const pad2 = n => String(n).padStart(2, '0');
const isoOf = d => `${d.getUTCFullYear()}-${pad2(d.getUTCMonth() + 1)}-${pad2(d.getUTCDate())}`;
const dateOf = iso => { const [y, m, d] = iso.split('-').map(Number); return new Date(Date.UTC(y, m - 1, d)); };
const addDays = (iso, n) => { const d = dateOf(iso); d.setUTCDate(d.getUTCDate() + n); return isoOf(d); };
const mondayOf = iso => addDays(iso, -((dateOf(iso).getUTCDay() + 6) % 7));
const todayISO = () => { const n = new Date(); return `${n.getFullYear()}-${pad2(n.getMonth() + 1)}-${pad2(n.getDate())}`; };
const fmtLong = iso => { const d = dateOf(iso); return `${kh(d.getUTCDate())} ${MONTHS[d.getUTCMonth()]} ${kh(d.getUTCFullYear())}`; };
const dnum = iso => dateOf(iso).getUTCDate();
const weekLabel = w => `${kh(dnum(w.start_date))}–${fmtLong(addDays(w.start_date, 4))}`;
const key = (sid, d, s, i) => `${sid}|${d}|${s}|${i}`;
const spKey = (sid, d, type) => `${sid}|${d}|${type}`;
const clone = o => JSON.parse(JSON.stringify(o));

/* ---------- 2. ទិន្នន័យ ---------- */
const DEFAULT_TEMPLATE = {"0": {"am": ["ការដំឡើងនិងថែទុកុំព្យូទ័រ", "ការដំឡើងនិងថែទុកុំព្យូទ័រ", "កុំព្យូទ័ររដ្ឋបាល", "កុំព្យូទ័ររដ្ឋបាល"], "pm": ["មូលដ្ឋានគ្រិះកុំព្យូទ័រ", "មូលដ្ឋានគ្រិះកុំព្យូទ័រ", "អប់រំចិត្ត"]}, "1": {"am": ["មូលដ្ឋានគ្រិះកុំព្យូទ័រ", "មូលដ្ឋានគ្រិះកុំព្យូទ័រ", "ការរចនាគេហទំព័រ១", "ការរចនាគេហទំព័រ១"], "pm": ["ភាសាខ្មែរ", "ភាសាខ្មែរ", "សីលធម៌"]}, "2": {"am": ["គណិតវិទ្យា", "គណិតវិទ្យា", "បណ្ណាល័យ", "បណ្ណាល័យ"], "pm": ["ប្រវត្តិវិទ្យា", "រូបវិទ្យា", "រូបវិទ្យា"]}, "3": {"am": ["បណ្ណាល័យ", "បណ្ណាល័យ", "ការរចនាគេហទំព័រ១", "ការរចនាគេហទំព័រ១"], "pm": ["ភាសាខ្មែរ", "គណិតវិទ្យា", "បណ្ណាល័យ"]}, "4": {"am": ["អេឡិចត្រូនិចសម្រាប់កុំព្យូទ័រ", "អេឡិចត្រូនិចសម្រាប់កុំព្យូទ័រ", "កុំព្យូទ័ររដ្ឋបាល", "កុំព្យូទ័ររដ្ឋបាល"], "pm": ["បណ្ណាល័យ", "ភាសាអង់គ្លេស", "ភាសាអង់គ្លេស"]}};
const DEFAULT_SETTINGS = {
  short_title: 'ថ្នាក់វិទ្យាសាស្ត្រកុំព្យូទ័រ',
  header_lines: [
    'ព្រះរាជាណាចក្រកម្ពុជា',
    'ជាតិ សាសនា ព្រះមហាក្សត្រ',
    'ក្រសួងអប់រំ យុវជន និងកីឡា',
    'សកលវិទ្យាល័យជាតិជាស៊ីមកំចាយមារ',
    'វិទ្យាស្ថានអភិវឌ្ឍន៍សហគមន៍'
  ],
  class_title: 'បញ្ជីវត្តមានសិស្សថ្នាក់អប់រំបច្ចេកទេស ជំនាញ វិទ្យាសាស្ដ្រកុំព្យូទ័រ ជំនាន់ទី៣ កម្រិត១',
  month_title: 'បញ្ជីសរុបវត្តមានប្រចាំខែសិស្សានុសិស្ស ជំនាន់៣ កម្រិត១ ឆ្នាំសិក្សា ២០២៥-២០២៦',
  major: 'ជំនាញ៖ វិទ្យាសាស្ត្រកុំព្យូទ័រ',
  grid_title: 'Attendance 10%',
  place: 'កំចាយមារ',
  taker: '',
  morning: 4,
  afternoon: 3,
  template: DEFAULT_TEMPLATE,
  // ម៉ោងពិសេស (ក្រៅកាលវិភាគ) · weekday: 0=ច័ន្ទ … 4=សុក្រ (ប្រចាំសប្ដាហ៍) · null = មិនទៀងទាត់ (បន្ថែមដោយដៃ)
  specialTypes: [
    { key: 'flag',    label: 'ម៉ោងគោរពទង់ជាតិ',      weekday: 0 },
    { key: 'skill',   label: 'ម៉ោងបំណិន',             weekday: null },
    { key: 'cleanup', label: 'ម៉ោងសម្អាតបរិស្ថាន',     weekday: null },
    { key: 'library', label: 'ម៉ោងបណ្ណាល័យ',          weekday: null }
  ]
};

const S = {
  sb: null, user: null, role: 'owner', settings: clone(DEFAULT_SETTINGS),
  students: [], weeks: [], week: null, abs: {}, specials: {}, sheets: [], gs: null,
  sc: { kind: 'month', month: '', sheetId: null }, rs: { kind: 'all', month: '' },
  tab: 'record', day: 0, per: { sess: 'am', i: 0 }, spType: null,
  rep: 'day', month: '', ttDay: 0, ttDraft: null, tok: 0, pending: 0
};
const isViewer = () => S.role === 'viewer';
const activeStudents = () => S.students.filter(s => s.active);

async function q(p) { const { data, error } = await p; if (error) throw error; return data; }

const ttGet = (tt, d, s, i) => (tt && tt[d] && tt[d][s] && tt[d][s][i]) || '';
function slotsOfDay(tt, d) {
  const out = [];
  for (const sess of ['am', 'pm']) {
    const n = sess === 'am' ? S.settings.morning : S.settings.afternoon;
    for (let i = 0; i < n; i++) out.push({ sess, i, subject: ttGet(tt, d, sess, i) });
  }
  return out;
}

async function loadProfile() {
  // កំណត់តួនាទី (owner/viewer)។ បើគ្មានជួរ profiles ទេ (schema.sql ចាស់, ឬ owner ដែលមិនទាន់ backfill)
  // ចាត់ទុកជា owner ដូចដើម ដើម្បីកុំឲ្យខូចគណនីចាស់។
  try {
    const { data, error } = await S.sb.from('profiles').select('role,owner_id').eq('user_id', S.user.id).maybeSingle();
    if (error) throw error;
    S.role = (data && data.role === 'viewer') ? 'viewer' : 'owner';
  } catch (e) { S.role = 'owner'; }
  if (isViewer()) S.tab = 'reports';
}
async function loadCore() {
  const [st, students, weeks] = await Promise.all([
    q(S.sb.from('settings').select('*')),
    q(S.sb.from('students').select('*').order('seq')),
    q(S.sb.from('weeks').select('*').order('start_date'))
  ]);
  if (st.length) { S.settings = Object.assign(clone(DEFAULT_SETTINGS), st[0].data || {}); S.settingsOwner = st[0].user_id; }
  else if (!isViewer()) { const r = await q(S.sb.from('settings').insert({ data: S.settings }).select()); S.settingsOwner = r[0].user_id; }
  S.students = students;
  S.weeks = weeks;
  try { S.sheets = await q(S.sb.from('grade_sheets').select('*').order('created_at')); } catch (e) { S.sheets = []; }   // មិនទាន់ Run schema_scores.sql → មិនឲ្យខូចទាំងមូល
  S.gs = null;
}
async function saveSettings() {
  await q(S.sb.from('settings').update({ data: S.settings, updated_at: new Date().toISOString() }).eq('user_id', S.settingsOwner || S.user.id));
}
async function fetchAbs(weekId) {
  if (S.abs[weekId]) return S.abs[weekId];
  const m = new Map();
  for (let from = 0; ; from += 1000) {
    const rows = await q(S.sb.from('absences').select('student_id,day,session,period,status')
      .eq('week_id', weekId).order('student_id').order('day').order('session').order('period').range(from, from + 999));
    rows.forEach(r => m.set(key(r.student_id, r.day, r.session, r.period), r.status));
    if (rows.length < 1000) break;
  }
  S.abs[weekId] = m;
  return m;
}
const ensureAbs = ws => Promise.all(ws.map(w => fetchAbs(w.id)));

async function fetchSpecials(weekId) {
  if (S.specials[weekId]) return S.specials[weekId];
  const m = new Map();
  for (let from = 0; ; from += 1000) {
    const rows = await q(S.sb.from('special_absences').select('student_id,day,type,status')
      .eq('week_id', weekId).range(from, from + 999));
    rows.forEach(r => m.set(spKey(r.student_id, r.day, r.type), r.status));
    if (rows.length < 1000) break;
  }
  S.specials[weekId] = m;
  return m;
}
async function applySpChanges(changes) {
  const w = S.week, m = await fetchSpecials(w.id);
  const old = changes.map(c => m.get(spKey(c.sid, c.day, c.type)));
  changes.forEach(c => { const k = spKey(c.sid, c.day, c.type); rtMark(k); c.status == null ? m.delete(k) : m.set(k, c.status); });
  renderView();
  busy(1);
  try {
    const ups = changes.filter(c => c.status != null).map(c => ({ week_id: w.id, student_id: c.sid, day: c.day, type: c.type, status: c.status }));
    if (ups.length) await q(S.sb.from('special_absences').upsert(ups, { onConflict: 'week_id,student_id,day,type' }));
    for (const c of changes.filter(c => c.status == null))
      await q(S.sb.from('special_absences').delete().eq('week_id', w.id).eq('student_id', c.sid).eq('day', c.day).eq('type', c.type));
  } catch (e) {
    changes.forEach((c, n) => { const k = spKey(c.sid, c.day, c.type); old[n] === undefined ? m.delete(k) : m.set(k, old[n]); });
    renderView(); toast('រក្សាទុកមិនបាន៖ ' + errText(e), 'err');
  } finally { busy(-1); }
}

function scheduledSlots(week) {
  const out = [];
  for (let d = 0; d < 5; d++) slotsOfDay(week.timetable, d).forEach(s => s.subject && out.push(Object.assign({ d }, s)));
  return out;
}
function weekTotals(week) {
  const m = S.abs[week.id] || new Map(), slots = scheduledSlots(week), res = new Map();
  for (const st of activeStudents()) {
    let p = 0, u = 0;
    for (const s of slots) { const v = m.get(key(st.id, s.d, s.sess, s.i)); if (v === 1) p++; else if (v === 0) u++; }
    res.set(st.id, { p, u });
  }
  return res;
}

function busy(n) {
  S.pending += n;
  const el = $('#saving'); if (el) { el.classList.toggle('on', S.pending > 0); el.textContent = 'កំពុងរក្សាទុក…'; }
}
let toastTimer;
function toast(msg, kind) {
  const el = $('#toast'); if (!el) return;
  el.textContent = msg; el.className = 'on ' + (kind || '');
  clearTimeout(toastTimer); toastTimer = setTimeout(() => { el.className = ''; }, kind === 'err' ? 4500 : 2200);
}
const errText = e => (e && (e.message || e.error_description)) || String(e);

/* ---------- ការកត់ត្រាអវត្តមាន ---------- */
async function applyChanges(changes) {
  const w = S.week, m = await fetchAbs(w.id);
  const old = changes.map(c => m.get(key(c.sid, c.day, c.sess, c.i)));
  changes.forEach(c => { const k = key(c.sid, c.day, c.sess, c.i); rtMark(k); c.status == null ? m.delete(k) : m.set(k, c.status); });
  renderView();
  busy(1);
  try {
    const ups = changes.filter(c => c.status != null).map(c => ({ week_id: w.id, student_id: c.sid, day: c.day, session: c.sess, period: c.i, status: c.status }));
    if (ups.length) await q(S.sb.from('absences').upsert(ups, { onConflict: 'week_id,student_id,day,session,period' }));
    const groups = new Map();
    changes.filter(c => c.status == null).forEach(c => {
      const g = `${c.sid}|${c.day}|${c.sess}`;
      if (!groups.has(g)) groups.set(g, { sid: c.sid, day: c.day, sess: c.sess, is: [] });
      groups.get(g).is.push(c.i);
    });
    for (const g of groups.values())
      await q(S.sb.from('absences').delete().eq('week_id', w.id).eq('student_id', g.sid).eq('day', g.day).eq('session', g.sess).in('period', g.is));
  } catch (e) {
    changes.forEach((c, n) => { const k = key(c.sid, c.day, c.sess, c.i); old[n] === undefined ? m.delete(k) : m.set(k, old[n]); });
    renderView(); toast('រក្សាទុកមិនបាន៖ ' + errText(e), 'err');
  } finally { busy(-1); }
}

/* ---------- 3. មើលឃើញ ---------- */
const ICONS = {
  record: '<path d="M9 11l3 3L22 4"/><path d="M21 12v7a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h11"/>',
  reports: '<path d="M4 20V10M10 20V4M16 20v-8M22 20H2"/>',
  students: '<circle cx="9" cy="8" r="3.5"/><path d="M2.5 20a6.5 6.5 0 0 1 13 0"/><path d="M16 4.5a3.5 3.5 0 0 1 0 7M18 14a6.5 6.5 0 0 1 3.5 6"/>',
  timetable: '<rect x="3" y="5" width="18" height="16" rx="2"/><path d="M8 3v4M16 3v4M3 10h18"/>',
  settings: '<circle cx="12" cy="12" r="3"/><path d="M19.4 15a1.7 1.7 0 0 0 .3 1.8l.1.1a2 2 0 1 1-2.8 2.8l-.1-.1a1.7 1.7 0 0 0-1.8-.3 1.7 1.7 0 0 0-1 1.5V21a2 2 0 1 1-4 0v-.1a1.7 1.7 0 0 0-1.1-1.5 1.7 1.7 0 0 0-1.8.3l-.1.1a2 2 0 1 1-2.8-2.8l.1-.1a1.7 1.7 0 0 0 .3-1.8 1.7 1.7 0 0 0-1.5-1H3a2 2 0 1 1 0-4h.1a1.7 1.7 0 0 0 1.5-1.1 1.7 1.7 0 0 0-.3-1.8l-.1-.1a2 2 0 1 1 2.8-2.8l.1.1a1.7 1.7 0 0 0 1.8.3H9a1.7 1.7 0 0 0 1-1.5V3a2 2 0 1 1 4 0v.1a1.7 1.7 0 0 0 1 1.5 1.7 1.7 0 0 0 1.8-.3l.1-.1a2 2 0 1 1 2.8 2.8l-.1.1a1.7 1.7 0 0 0-.3 1.8V9a1.7 1.7 0 0 0 1.5 1H21a2 2 0 1 1 0 4h-.1a1.7 1.7 0 0 0-1.5 1z"/>',
  scores: '<path d="M9 3h6l1 2h3v16H5V5h3z"/><path d="M9 12h6M9 16h4"/>',
  specials: '<path d="M12 2.5l2.9 5.88 6.5.94-4.7 4.58 1.1 6.47L12 17.3l-5.8 3.07 1.1-6.47-4.7-4.58 6.5-.94z"/>'
};
const TABS_OWNER = [['record', 'កត់វត្តមាន'], ['specials', 'ម៉ោងពិសេស'], ['scores', 'ពិន្ទុ'], ['reports', 'របាយការណ៍'], ['students', 'សិស្ស'], ['timetable', 'កាលវិភាគ'], ['settings', 'ការកំណត់']];
const TABS_VIEWER = [['reports', 'របាយការណ៍']];
const TABS = () => isViewer() ? TABS_VIEWER : TABS_OWNER;

function showLogin(msg) {
  const missing = !CFG.SUPABASE_URL || !CFG.SUPABASE_ANON_KEY || /YOUR-/.test(CFG.SUPABASE_URL + CFG.SUPABASE_ANON_KEY);
  app.innerHTML = `<main class="login" style="padding:20px"><form id="loginForm" autocomplete="on">
    <h1>ប្រព័ន្ធគ្រប់គ្រងវត្តមាន</h1><p class="muted" style="margin:0 0 8px">ចូលដោយគណនីរបស់អ្នក</p>
    ${missing ? '<p class="err-text">មិនទាន់បានកំណត់ config.js (SUPABASE_URL និង SUPABASE_ANON_KEY)</p>' : ''}
    <label class="f" for="em">អ៊ីមែល</label><input id="em" type="email" autocomplete="username" required>
    <label class="f" for="pw">លេខសម្ងាត់</label><input id="pw" type="password" autocomplete="current-password" required>
    <p class="err-text" id="loginErr">${esc(msg || '')}</p>
    <button class="btn primary" style="width:100%;margin-top:6px" ${missing ? 'disabled' : ''}>ចូល</button></form></main><div id="toast"></div>`;
}

function renderShell() {
  if (isViewer()) S.tab = 'reports';
  const weeks = S.weeks.slice().sort((a, b) => b.start_date.localeCompare(a.start_date));
  const tabs = TABS();
  app.innerHTML = `
  <header class="top">
    <div class="brand"><b>វត្តមាន</b><small>${esc(S.settings.short_title || '')}${isViewer() ? ' · អ្នកមើល' : ''}</small></div>
    <div class="weekpick">
      <select id="weekSel" aria-label="ជ្រើសរើសសប្ដាហ៍">${weeks.map(w => `<option value="${w.id}" ${S.week && w.id === S.week.id ? 'selected' : ''}>${weekLabel(w)}</option>`).join('') || '<option>—</option>'}</select>
      ${isViewer() ? '<button class="icon" data-act="logout" aria-label="ចាកចេញ" title="ចាកចេញ">⎋</button>' : '<button class="icon" data-act="newWeek" aria-label="បន្ថែមសប្ដាហ៍ថ្មី">+</button>'}
    </div>
  </header>
  <div id="saving"></div>
  <main id="view"></main>
  <nav class="tabs" style="grid-template-columns:repeat(${tabs.length},minmax(66px,1fr))" aria-label="ម៉ឺនុយ">${tabs.map(([k, l]) => `<button data-act="tab" data-v="${k}" class="${S.tab === k ? 'on' : ''}" ${S.tab === k ? 'aria-current="page"' : ''}><svg viewBox="0 0 24 24" aria-hidden="true">${ICONS[k]}</svg>${l}</button>`).join('')}</nav>
  <div id="toast" role="status" aria-live="polite"></div>`;
  renderView();
}

async function renderView() {
  if (isViewer() && S.tab !== 'reports') S.tab = 'reports';
  const v = $('#view'); if (!v) return;
  $$('.tabs button').forEach(b => { const on = b.dataset.v === S.tab; b.classList.toggle('on', on); on ? b.setAttribute('aria-current', 'page') : b.removeAttribute('aria-current'); });
  const tok = ++S.tok;
  if (!S.week && (['record', 'specials', 'timetable'].includes(S.tab) || (S.tab === 'reports' && S.rep !== 'score'))) { v.innerHTML = viewNoWeek(); return; }
  switch (S.tab) {
    case 'record':
      if (!S.abs[S.week.id]) { v.innerHTML = '<p class="pad muted">កំពុងផ្ទុក…</p>'; await fetchAbs(S.week.id).catch(e => toast(errText(e), 'err')); if (tok !== S.tok) return; }
      v.innerHTML = viewRecord(); break;
    case 'specials':
      if (!S.specials[S.week.id]) { v.innerHTML = '<p class="pad muted">កំពុងផ្ទុក…</p>'; await fetchSpecials(S.week.id).catch(e => toast(errText(e), 'err')); if (tok !== S.tok) return; }
      v.innerHTML = viewSpecials(); break;
    case 'scores':
      if (!S.gs) { v.innerHTML = '<p class="pad muted">កំពុងផ្ទុក…</p>'; await fetchGS().catch(e => toast(errText(e), 'err')); if (tok !== S.tok) return; }
      v.innerHTML = S.gs ? viewScores() : '<div class="card"><p class="muted">ផ្ទុកពិន្ទុមិនបាន — ពិនិត្យថាបាន Run schema_scores.sql រួចហើយ។</p></div>'; break;
    case 'reports': {
      if (S.rep === 'score') {
        if (!S.gs) { v.innerHTML = viewReports('<p class="pad muted">កំពុងផ្ទុក…</p>'); await fetchGS().catch(e => toast(errText(e), 'err')); if (tok !== S.tok) return; }
        v.innerHTML = viewReports(S.gs ? viewScoreReport() : '<div class="card"><p class="muted">ផ្ទុកពិន្ទុមិនបាន — ពិនិត្យថាបាន Run schema_scores.sql រួចហើយ។</p></div>'); break;
      }
      if (S.rep === 'month') {
        const ws = monthWeeks();
        if (ws.some(w => !S.abs[w.id])) { v.innerHTML = viewReports('<p class="pad muted">កំពុងផ្ទុក…</p>'); await ensureAbs(ws).catch(e => toast(errText(e), 'err')); if (tok !== S.tok) return; }
        v.innerHTML = viewReports(monthTable(ws));
      } else {
        if (!S.abs[S.week.id]) { await fetchAbs(S.week.id).catch(e => toast(errText(e), 'err')); if (tok !== S.tok) return; }
        v.innerHTML = viewReports(S.rep === 'day' ? dayList() : weekTable());
      }
      break;
    }
    case 'students': v.innerHTML = viewStudents(); break;
    case 'timetable': v.innerHTML = viewTimetable(); break;
    case 'settings': v.innerHTML = viewSettings(); break;
  }
}

const viewNoWeek = () => `<div class="card"><h2>មិនទាន់មានសប្ដាហ៍</h2><p class="muted">បង្កើតសប្ដាហ៍ដំបូង ដើម្បីចាប់ផ្តើមកត់វត្តមាន។</p><button class="btn primary" data-act="newWeek">+ បង្កើតសប្ដាហ៍</button></div>`;

function viewRecord() {
  const w = S.week, m = S.abs[w.id] || new Map(), stu = activeStudents(), slots = slotsOfDay(w.timetable, S.day);
  const cur = slots.find(s => s.sess === S.per.sess && s.i === S.per.i) || slots[0];
  const cnt = s => stu.reduce((a, st) => a + (m.has(key(st.id, S.day, s.sess, s.i)) ? 1 : 0), 0);
  const days = DAYS.map((n, d) => `<button class="day ${d === S.day ? 'on' : ''}" data-act="day" data-v="${d}"><span>${n}</span><small>${kh(dnum(addDays(w.start_date, d)))}</small></button>`).join('');
  const chips = slots.map(s => { const c = cnt(s); return `<button class="chip ${s === cur ? 'on' : ''} ${s.subject ? '' : 'empty'}" data-act="per" data-sess="${s.sess}" data-i="${s.i}"><b>${SESS_S[s.sess]} ${kh(s.i + 1)}</b><span>${esc(s.subject || '—')}</span>${c ? `<i class="bdg">${kh(c)}</i>` : ''}</button>`; }).join('');
  let body;
  if (!stu.length) body = '<div class="card"><p class="muted">មិនទាន់មានសិស្ស។ សូមបន្ថែមក្នុងផ្ទាំង "សិស្ស"។</p></div>';
  else if (!cur || !cur.subject) body = '<div class="card"><p class="muted">ម៉ោងនេះមិនទាន់មានមុខវិជ្ជា។ សូមកំណត់ក្នុងផ្ទាំង "កាលវិភាគ"។</p></div>';
  else {
    let p = 0, u = 0;
    const rows = stu.map((st, n) => {
      const v = m.get(key(st.id, S.day, cur.sess, cur.i)); if (v === 1) p++; else if (v === 0) u++;
      const cls = v === 1 ? 'perm' : v === 0 ? 'unperm' : 'ok', txt = v === 1 ? 'មានច្បាប់' : v === 0 ? 'អត់ច្បាប់' : '✓ មក';
      return `<li class="stu"><span class="no">${kh(n + 1)}</span><span class="nm">${esc(st.name)}${st.sex ? `<small>${esc(st.sex)}</small>` : ''}</span><button class="st ${cls}" data-act="cycle" data-sid="${st.id}" aria-label="${esc(st.name)}៖ ${txt}">${txt}</button><button class="more" data-act="more" data-sid="${st.id}" aria-label="ជម្រើសបន្ថែមសម្រាប់ ${esc(st.name)}">⋯</button></li>`;
    }).join('');
    body = `<div class="pcard"><h2>${esc(cur.subject)}</h2>
      <p>${DAYS[S.day]} · ${SESS[cur.sess]} ម៉ោងទី ${kh(cur.i + 1)} — មក ${kh(stu.length - p - u)} · មានច្បាប់ ${kh(p)} · អត់ច្បាប់ ${kh(u)}</p>
      <div class="row"><button class="btn small" data-act="allPresent">មកទាំងអស់</button>${cur.i > 0 ? '<button class="btn small" data-act="copyPrev">ចម្លងពីម៉ោងមុន</button>' : ''}</div></div>
      <ul class="stulist">${rows}</ul>`;
  }
  return `<div class="daybar">${days}</div><div class="chips">${chips}</div>${body}`;
}

function viewSpecials() {
  const w = S.week, m = S.specials[w.id] || new Map(), stu = activeStudents(), types = S.settings.specialTypes || [];
  if (!S.spType || !types.some(t => t.key === S.spType)) S.spType = types[0] ? types[0].key : null;
  const cnt = type => stu.reduce((a, st) => a + (m.has(spKey(st.id, S.day, type)) ? 1 : 0), 0);
  const days = DAYS.map((n, d) => `<button class="day ${d === S.day ? 'on' : ''}" data-act="day" data-v="${d}"><span>${n}</span><small>${kh(dnum(addDays(w.start_date, d)))}</small></button>`).join('');
  const chips = types.map(t => { const c = cnt(t.key); return `<button class="chip ${t.key === S.spType ? 'on' : ''}" data-act="spType" data-k="${t.key}"><b>${esc(t.label)}</b><span>${t.weekday === S.day ? 'ប្រចាំថ្ងៃនេះ' : '—'}</span>${c ? `<i class="bdg">${kh(c)}</i>` : ''}</button>`; }).join('');
  const curT = types.find(x => x.key === S.spType);
  let body;
  if (!types.length) body = '<div class="card"><p class="muted">មិនទាន់កំណត់ប្រភេទម៉ោងពិសេសទេ — កំណត់ក្នុងផ្ទាំង "ការកំណត់"</p></div>';
  else if (!stu.length) body = '<div class="card"><p class="muted">មិនទាន់មានសិស្ស។ សូមបន្ថែមក្នុងផ្ទាំង "សិស្ស"។</p></div>';
  else {
    let p = 0, u = 0;
    const rows = stu.map((st, n) => {
      const v = m.get(spKey(st.id, S.day, S.spType)); if (v === 1) p++; else if (v === 0) u++;
      const cls = v === 1 ? 'perm' : v === 0 ? 'unperm' : 'ok', txt = v === 1 ? 'មានច្បាប់' : v === 0 ? 'អត់ច្បាប់' : '✓ មក';
      return `<li class="stu"><span class="no">${kh(n + 1)}</span><span class="nm">${esc(st.name)}${st.sex ? `<small>${esc(st.sex)}</small>` : ''}</span><button class="st ${cls}" data-act="spCycle" data-sid="${st.id}" aria-label="${esc(st.name)}៖ ${txt}">${txt}</button></li>`;
    }).join('');
    body = `<div class="pcard"><h2>${esc(curT.label)}</h2>
      <p>${DAYS[S.day]} ${fmtLong(addDays(w.start_date, S.day))} — មក ${kh(stu.length - p - u)} · មានច្បាប់ ${kh(p)} · អត់ច្បាប់ ${kh(u)}</p>
      <div class="row"><button class="btn small" data-act="spAllPresent">មកទាំងអស់</button></div></div>
      <ul class="stulist">${rows}</ul>`;
  }
  return `<div class="daybar">${days}</div><div class="chips">${chips}</div>${body}`;
}

/* ---------- ពិន្ទុតាមមុខវិជ្ជា ----------
   មុខវិជ្ជាមួយ = តារាងមួយ (grade_sheets) ដូច Excel៖ សប្តាហ៍ទី១-៤ · វត្តមាន · កិច្ចការ · ប្រឡង → សរុប */
const KIND_L = { month: 'ប្រចាំខែ', term: 'ត្រីមាស/ឆមាស' };
const COMPS = [['w1', 'សប្ដាហ៍ទី១'], ['w2', 'សប្ដាហ៍ទី២'], ['w3', 'សប្ដាហ៍ទី៣'], ['w4', 'សប្ដាហ៍ទី៤'], ['att', 'ពិន្ទុវត្តមាន'], ['hw', 'កិច្ចការ'], ['exam', 'ពិន្ទុប្រឡង']];
const DEF_MAX = { w1: 5, w2: 5, w3: 5, w4: 5, att: 10, hw: 20, exam: 50 };   // ផលបូក = ១០០ (កែបានក្នុងមុខវិជ្ជានីមួយៗ)
const r1 = n => Math.round(n * 10) / 10;
const sn = n => (n == null ? '—' : kh(String(r1(n))));
const maxOf = sh => Object.assign({}, DEF_MAX, sh.maxes || {});
const maxSum = sh => COMPS.reduce((a, [c]) => a + Number(maxOf(sh)[c] || 0), 0);
async function fetchGS() {
  if (S.gs) return S.gs;
  const m = new Map();
  for (let from = 0; ; from += 1000) {
    const rows = await q(S.sb.from('grade_scores').select('sheet_id,student_id,w1,w2,w3,w4,att,hw,exam').order('sheet_id').order('student_id').range(from, from + 999));
    rows.forEach(r => m.set(r.sheet_id + '|' + r.student_id, normScore(r)));
    if (rows.length < 1000) break;
  }
  S.gs = m;
  return m;
}
function normScore(r) { const o = {}; COMPS.forEach(([c]) => { o[c] = r[c] == null ? null : Number(r[c]); }); return o; }
const sheetsOf = (kind, month) => S.sheets.filter(x => (kind === 'all' || x.kind === kind) && (!month || x.month === month))
  .sort((a, b) => String(a.created_at).localeCompare(String(b.created_at)));
const curSheets = () => sheetsOf(S.sc.kind, S.sc.month);
function curSheet() { const l = curSheets(); return l.find(x => x.id === S.sc.sheetId) || l[0] || null; }
function rankRows(rows) {
  const ranked = rows.filter(r => r.pct !== null).sort((a, b) => Math.round(b.pct * 100) - Math.round(a.pct * 100));
  ranked.forEach((r, i) => { r.rank = i > 0 && Math.round(r.pct * 100) === Math.round(ranked[i - 1].pct * 100) ? ranked[i - 1].rank : i + 1; });
}
function sheetRows(sh) {
  const mx = maxSum(sh);
  const rows = activeStudents().map(st => {
    const v = S.gs.get(sh.id + '|' + st.id) || {}, got = COMPS.filter(([c]) => v[c] != null);
    const total = got.length ? got.reduce((a, [c]) => a + v[c], 0) : null;
    return { st, v, total, pct: total === null || !mx ? null : total / mx * 100, rank: 0 };
  });
  rankRows(rows);
  return rows;
}
function viewScores() {
  if (!S.sc.month) S.sc.month = todayISO().slice(0, 7);
  const b = (k, l) => `<button class="${S.sc.kind === k ? 'on' : ''}" data-act="scKind" data-v="${k}">${l}</button>`;
  const list = curSheets(), sh = curSheet(); S.sc.sheetId = sh ? sh.id : null;
  const head = `<div class="tabs2">${b('month', 'ប្រចាំខែ')}${b('term', 'ត្រីមាស/ឆមាស')}</div>
  <div class="card"><div class="row"><div class="grow"><label class="f" for="scMonth" style="margin-top:0">ខែ</label><input id="scMonth" type="month" value="${S.sc.month}"></div><button class="btn primary" data-act="newSheet" style="align-self:flex-end">+ មុខវិជ្ជា</button></div></div>
  <div class="chips">${list.map(x => `<button class="chip ${sh && x.id === sh.id ? 'on' : ''}" data-act="scSheet" data-id="${x.id}"><b>${esc(x.subject)}</b><span>/${kh(r1(maxSum(x)))}</span></button>`).join('')}</div>`;
  if (!sh) return head + '<div class="card"><p class="muted">មិនទាន់មានមុខវិជ្ជាសម្រាប់ខែនេះទេ — ចុច "+ មុខវិជ្ជា" ដើម្បីចាប់ផ្តើមបញ្ចូលពិន្ទុ។</p></div>';
  const mx = maxOf(sh), rows = sheetRows(sh), tmax = maxSum(sh);
  const th = COMPS.map(([c, l]) => `<th>${l}<br><small class="muted">/${kh(Number(mx[c]))}</small></th>`).join('');
  const body = rows.map((t, n) => `<tr><td>${kh(n + 1)}</td><td class="l">${esc(t.st.name)}</td>${COMPS.map(([c, l]) => `<td><input class="scin" type="number" inputmode="decimal" step="any" min="0" max="${Number(mx[c])}" data-sid="${t.st.id}" data-c="${c}" value="${t.v[c] == null ? '' : t.v[c]}" ${Number(mx[c]) > 0 ? '' : 'disabled'} style="width:58px;padding:6px;text-align:center" aria-label="${esc(l)} ${esc(t.st.name)}"></td>`).join('')}<td><b id="t-${t.st.id}">${sn(t.total)}</b></td><td id="p-${t.st.id}">${sn(t.pct)}</td><td id="r-${t.st.id}">${t.rank ? kh(t.rank) : '—'}</td></tr>`).join('');
  return `${head}<div class="row" style="margin-bottom:8px"><b style="font-family:var(--font-kh)">${esc(sh.subject)}</b><span class="grow"></span><button class="btn small" data-act="editSheet" data-id="${sh.id}">កែមុខវិជ្ជា / ពិន្ទុពេញ</button></div>
  <div class="tw"><table class="t"><thead><tr><th>ល.រ</th><th class="l">ឈ្មោះ</th>${th}<th>ពិន្ទុសរុប<br><small class="muted">/${kh(r1(tmax))}</small></th><th>%</th><th>ចំណាត់ថ្នាក់</th></tr></thead><tbody>${body || '<tr><td colspan="11" class="muted">មិនទាន់មានសិស្ស</td></tr>'}</tbody></table></div>
  <p class="muted" style="margin:8px 0 0;font-size:.85rem">ពិន្ទុសរុប = សប្ដាហ៍១-៤ + វត្តមាន + កិច្ចការ + ប្រឡង · % = សរុប ÷ ពិន្ទុពេញសរុប × ១០០ · ចំណាត់ថ្នាក់តាម % · Enter = ចុះទៅសិស្សបន្ទាប់ · ${kh(activeStudents().length)} នាក់</p>`;
}
function sheetDialog(e) {
  const mx = e ? maxOf(e) : DEF_MAX;
  const d = dialog(`<h2>${e ? 'កែមុខវិជ្ជា' : 'បន្ថែមមុខវិជ្ជា'}</h2>
    <label class="f" for="x-sub">មុខវិជ្ជា (ជ្រើសពីកាលវិភាគ ឬវាយខ្លួនឯង)</label><input id="x-sub" type="text" list="x-subs" value="${esc(e ? e.subject : '')}" autocomplete="off"><datalist id="x-subs">${subjectList().map(x => `<option value="${esc(x)}">`).join('')}</datalist>
    <p class="muted" style="margin:10px 0 0;font-size:.85rem">ពិន្ទុពេញនៃសមាសធាតុនីមួយៗ (ដាក់ ០ បើមិនប្រើ)</p>
    <div class="row">${COMPS.map(([c, l]) => `<div style="flex:1 1 120px"><label class="f" for="x-m-${c}">${l}</label><input id="x-m-${c}" type="number" inputmode="decimal" step="any" min="0" value="${Number(mx[c])}"></div>`).join('')}</div>
    <div class="foot">${e ? `<button class="btn danger" data-act="delSheet" data-id="${e.id}">លុប</button>` : ''}<span class="grow"></span>${closeBtn}<button class="btn primary" data-act="saveSheet" data-id="${e ? e.id : ''}">រក្សាទុក</button></div>`);
  setTimeout(() => $('#x-sub', d).focus(), 50);
}
function subjectList() {
  const set = new Set();
  const add = tt => { for (const d in (tt || {})) for (const sess of ['am', 'pm']) ((tt[d] && tt[d][sess]) || []).forEach(x => x && set.add(x)); };
  add(S.settings.template); S.weeks.forEach(w => add(w.timetable)); S.sheets.forEach(x => set.add(x.subject));
  return [...set];
}
async function saveScore(inp) {
  const sh = curSheet(); if (!sh || isViewer()) return;
  const sid = inp.dataset.sid, c = inp.dataset.c, k = sh.id + '|' + sid, raw = inp.value.trim();
  const cur = S.gs.get(k) || {}, old = cur[c] == null ? null : cur[c], mx = Number(maxOf(sh)[c]);
  let v = null;
  if (raw !== '') {
    v = parseFloat(raw);
    if (!(v >= 0) || v > mx) { toast(`ពិន្ទុត្រូវនៅចន្លោះ ០ ដល់ ${kh(mx)}`, 'err'); inp.value = old == null ? '' : old; return; }
  }
  if (v === old) return;
  const had = S.gs.has(k), entry = Object.assign(cur, { [c]: v });
  COMPS.forEach(([x]) => { if (entry[x] === undefined) entry[x] = null; });
  S.gs.set(k, entry);
  refreshGridTotals();
  rtMark('gs|' + k); busy(1);
  try { await q(S.sb.from('grade_scores').upsert({ sheet_id: sh.id, student_id: sid, [c]: v }, { onConflict: 'sheet_id,student_id' })); }
  catch (err) {
    entry[c] = old; if (!had) S.gs.delete(k);
    inp.value = old == null ? '' : old; refreshGridTotals();
    toast('រក្សាទុកមិនបាន៖ ' + errText(err), 'err');
  } finally { busy(-1); }
}
function refreshGridTotals() {
  const sh = curSheet(); if (!sh) return;
  sheetRows(sh).forEach(t => {
    const a = $('#t-' + t.st.id), p = $('#p-' + t.st.id), r = $('#r-' + t.st.id);
    if (a) a.textContent = sn(t.total); if (p) p.textContent = sn(t.pct); if (r) r.textContent = t.rank ? kh(t.rank) : '—';
  });
}
/* របាយការណ៍ពិន្ទុ៖ មុខវិជ្ជានីមួយៗជា % (សរុប ÷ ពិន្ទុពេញ) · រួម សរុប/មធ្យមភាគ/ចំណាត់ថ្នាក់ */
function scoreSummary(kind, month) {
  const sheets = sheetsOf(kind, month), subjects = [...new Set(sheets.map(x => x.subject))];
  const bySheet = new Map(sheets.map(sh => [sh.id, sheetRows(sh)]));
  const rows = activeStudents().map((st, n) => {
    const per = {};
    subjects.forEach(sub => {
      const v = sheets.filter(x => x.subject === sub).map(sh => bySheet.get(sh.id)[n].pct).filter(x => x !== null);
      per[sub] = v.length ? v.reduce((a, b) => a + b, 0) / v.length : null;
    });
    const got = subjects.filter(sub => per[sub] !== null), total = got.reduce((a, sub) => a + per[sub], 0);
    return { st, per, total: got.length ? total : null, pct: got.length ? total / got.length : null, rank: 0 };
  });
  rankRows(rows);
  rows.forEach(r => { r.avg = r.pct; });
  return { sheets, subjects, rows, bySheet };
}
function viewScoreReport() {
  const sum = scoreSummary(S.rs.kind, S.rs.month);
  const kinds = [['all', 'ទាំងអស់'], ...Object.entries(KIND_L)].map(([k, l]) => `<option value="${k}" ${S.rs.kind === k ? 'selected' : ''}>${l}</option>`).join('');
  const ctl = `<div class="card"><div class="row"><div class="grow"><label class="f" for="rsKind" style="margin-top:0">ប្រភេទ</label><select id="rsKind">${kinds}</select></div><div class="grow"><label class="f" for="rsMonth" style="margin-top:0">ខែ</label><input id="rsMonth" type="month" value="${S.rs.month}"></div></div><button class="btn small" style="margin-top:8px" data-act="rsMonthAll">បង្ហាញគ្រប់ខែ</button></div>`;
  if (!sum.sheets.length) return ctl + '<div class="card"><p class="muted">មិនទាន់មានពិន្ទុសម្រាប់ជម្រើសនេះទេ។</p></div>';
  const th = sum.subjects.map(x => `<th>${esc(x)}</th>`).join('');
  const rows = sum.rows.map((r, n) => `<tr><td>${kh(n + 1)}</td><td class="l">${esc(r.st.name)}</td>${sum.subjects.map(x => `<td>${sn(r.per[x])}</td>`).join('')}<td><b>${sn(r.total)}</b></td><td><b>${sn(r.avg)}</b></td><td>${r.rank ? kh(r.rank) : '—'}</td></tr>`).join('');
  return `${ctl}<p class="muted" style="margin:0 0 8px;font-size:.85rem">មុខវិជ្ជានីមួយៗជាភាគរយ (ពិន្ទុសរុបមុខវិជ្ជា ÷ ពិន្ទុពេញ × ១០០) · ការនាំចេញ Excel មានតារាងលម្អិតរបស់មុខវិជ្ជានីមួយៗ (សប្តាហ៍១-៤ វត្តមាន កិច្ចការ ប្រឡង)</p>
  <div class="tw"><table class="t"><thead><tr><th>ល.រ</th><th class="l">ឈ្មោះ</th>${th}<th>សរុប</th><th>មធ្យមភាគ</th><th>ចំណាត់ថ្នាក់</th></tr></thead><tbody>${rows}</tbody></table></div>
  <div class="row" style="margin-top:12px"><button class="btn primary grow" data-act="exportScores">នាំចេញ Excel</button></div>`;
}
function buildScoreWorkbook(sum) {
  const s = S.settings, H = { font: { bold: true }, alignment: { horizontal: 'center', vertical: 'center', wrapText: true }, fill: { fgColor: { rgb: 'EEF3FB' } } };
  const c = (v, st) => ({ v, t: typeof v === 'number' ? 'n' : 's', s: st });
  const monthTxt = S.rs.month ? ' · ខែ' + MONTHS[+S.rs.month.slice(5, 7) - 1] + ' ' + S.rs.month.slice(0, 4) : '';
  const top = (aoa, title) => {
    (s.header_lines || []).forEach(l => aoa.push([c(l, { font: { bold: true } })]));
    aoa.push([c(title, { font: { bold: true, sz: 14 } })]);
    if (s.class_title) aoa.push([c(s.class_title)]);
    aoa.push([]);
  };
  const wb = XLSX.utils.book_new(), used = new Set();
  // សន្លឹកសរុប
  const aoa = [];
  top(aoa, 'តារាងពិន្ទុ' + (S.rs.kind === 'all' ? '' : ' · ' + KIND_L[S.rs.kind]) + monthTxt);
  aoa.push(['ល.រ', 'ឈ្មោះ', 'ភេទ', ...sum.subjects, 'ពិន្ទុសរុប', 'មធ្យមភាគ', 'ចំណាត់ថ្នាក់'].map(h => c(h, H)));
  sum.rows.forEach((r, i) => aoa.push([c(i + 1), c(r.st.name), c(r.st.sex || ''), ...sum.subjects.map(x => c(r.per[x] == null ? '' : r1(r.per[x]))), c(r.total == null ? '' : r1(r.total)), c(r.avg == null ? '' : r1(r.avg)), c(r.rank || '')]));
  aoa.push([]); aoa.push([c('មុខវិជ្ជានីមួយៗជាភាគរយ (ពិន្ទុសរុប ÷ ពិន្ទុពេញ × ១០០)', { font: { italic: true } })]);
  const ws = XLSX.utils.aoa_to_sheet(aoa);
  ws['!cols'] = [{ wch: 6 }, { wch: 28 }, { wch: 8 }, ...sum.subjects.map(() => ({ wch: 16 })), { wch: 12 }, { wch: 12 }, { wch: 12 }];
  XLSX.utils.book_append_sheet(wb, ws, 'សរុប'); used.add('សរុប');
  // មួយសន្លឹកក្នុងមួយមុខវិជ្ជា (ដូច Excel របស់អ្នក)
  sum.sheets.forEach(sh => {
    const mx = maxOf(sh), tmax = maxSum(sh), a = [];
    top(a, `បញ្ជីសម្រង់ពិន្ទុមុខវិជ្ជា ${sh.subject} សម្រាប់ខែ ${MONTHS[+sh.month.slice(5, 7) - 1]} ${sh.month.slice(0, 4)}`);
    a.push([...['ល.រ', 'ឈ្មោះ', 'ភេទ', ...COMPS.map(([, l]) => l)].map(h => c(h, H)), c('ពិន្ទុសរុប', H), c('%', H), c('ចំណាត់ថ្នាក់', H)]);
    a.push([c(''), c(''), c('ពិន្ទុពេញ', { font: { italic: true } }), ...COMPS.map(([k]) => c(Number(mx[k]), { font: { italic: true } })), c(r1(tmax), { font: { italic: true } })]);
    sum.bySheet.get(sh.id).forEach((r, i) => a.push([c(i + 1), c(r.st.name), c(r.st.sex || ''), ...COMPS.map(([k]) => c(r.v[k] == null ? '' : r.v[k])), c(r.total == null ? '' : r1(r.total)), c(r.pct == null ? '' : r1(r.pct)), c(r.rank || '')]));
    const w2 = XLSX.utils.aoa_to_sheet(a);
    w2['!cols'] = [{ wch: 6 }, { wch: 28 }, { wch: 10 }, ...COMPS.map(() => ({ wch: 12 })), { wch: 12 }, { wch: 8 }, { wch: 12 }];
    let nm = (sh.subject + (S.rs.kind === 'all' ? ' ' + (sh.kind === 'term' ? 'ត្រីមាស' : '') + sh.month.slice(5) : '')).replace(/[\\\/?*\[\]:]/g, '').trim().slice(0, 28) || 'សន្លឹក';
    let u = nm, n = 2; while (used.has(u)) u = nm.slice(0, 26) + n++;
    used.add(u); XLSX.utils.book_append_sheet(wb, w2, u);
  });
  return wb;
}
function viewReports(inner) {
  const b = (k, l) => `<button class="${S.rep === k ? 'on' : ''}" data-act="rep" data-v="${k}">${l}</button>`;
  return `<div class="tabs2">${b('day', 'អ្នកឈប់ថ្ងៃ')}${b('week', 'សរុបសប្ដាហ៍')}${b('month', 'សរុបខែ')}${b('score', 'ពិន្ទុ')}</div>${inner}`;
}
function dayList() {
  const w = S.week, m = S.abs[w.id] || new Map(), slots = slotsOfDay(w.timetable, S.day).filter(s => s.subject);
  const days = DAYS.map((n, d) => `<button class="day ${d === S.day ? 'on' : ''}" data-act="day" data-v="${d}"><span>${n}</span><small>${kh(dnum(addDays(w.start_date, d)))}</small></button>`).join('');
  const items = activeStudents().map(st => ({ st, it: slots.map(s => ({ s, v: m.get(key(st.id, S.day, s.sess, s.i)) })).filter(x => x.v !== undefined) })).filter(x => x.it.length);
  const html = items.map(({ st, it }) => {
    const p = it.filter(x => x.v === 1).length, u = it.length - p;
    return `<div class="absitem"><b>${esc(st.name)}</b> <span class="muted">— ${kh(it.length)} ម៉ោង${p ? ` · មានច្បាប់ ${kh(p)}` : ''}${u ? ` · អត់ច្បាប់ ${kh(u)}` : ''}</span><div>${it.map(x => `<span class="pill ${x.v === 1 ? 'perm' : 'unperm'}">${SESS_S[x.s.sess]} ${kh(x.s.i + 1)} · ${esc(x.s.subject)}</span>`).join('')}</div></div>`;
  }).join('');
  return `<div class="daybar">${days}</div><p class="muted" style="margin:0 0 8px">${DAYS[S.day]} ${fmtLong(addDays(w.start_date, S.day))} — អវត្តមាន ${kh(items.length)} នាក់</p>${html || '<div class="card"><p class="muted">គ្មានអ្នកអវត្តមានទេ 🎉</p></div>'}`;
}
function weekTable() {
  const w = S.week, t = weekTotals(w); let P = 0, U = 0;
  const rows = activeStudents().map((st, n) => { const x = t.get(st.id); P += x.p; U += x.u; return `<tr><td>${kh(n + 1)}</td><td class="l">${esc(st.name)}</td><td>${kh(x.p)}</td><td>${kh(x.u)}</td><td><b>${kh(x.p + x.u)}</b></td></tr>`; }).join('');
  return `<div class="tw"><table class="t"><thead><tr><th>ល.រ</th><th class="l">ឈ្មោះ</th><th>មានច្បាប់</th><th>អត់ច្បាប់</th><th>សរុប</th></tr></thead><tbody>${rows}</tbody><tfoot><tr><td colspan="2" class="l">សរុប</td><td>${kh(P)}</td><td>${kh(U)}</td><td>${kh(P + U)}</td></tr></tfoot></table></div>
  <div class="row" style="margin-top:12px"><button class="btn primary grow" data-act="exportWeek">នាំចេញ Excel (សប្ដាហ៍នេះ)</button></div>`;
}
function monthWeeks() {
  if (!S.month) S.month = S.week ? S.week.start_date.slice(0, 7) : todayISO().slice(0, 7);
  return S.weeks.filter(w => w.start_date.slice(0, 7) === S.month).sort((a, b) => a.start_date.localeCompare(b.start_date));
}
function monthTable(ws) {
  const head = `<div class="card"><label class="f" for="monthSel" style="margin-top:0">ជ្រើសរើសខែ</label><input id="monthSel" type="month" value="${S.month}"></div>`;
  if (!ws.length) return head + '<div class="card"><p class="muted">មិនមានសប្ដាហ៍ក្នុងខែនេះទេ។</p></div>';
  const tots = ws.map(weekTotals);
  const th1 = ws.map((w, k) => `<th colspan="2">សប្ដាហ៍ ${kh(k + 1)}</th>`).join('');
  const th2 = ws.map(() => '<th>ច្បាប់</th><th>អត់</th>').join('');
  let P = 0, U = 0; const colT = ws.map(() => [0, 0]);
  const rows = activeStudents().map((st, n) => {
    let p = 0, u = 0;
    const cells = tots.map((t, k) => { const x = t.get(st.id); p += x.p; u += x.u; colT[k][0] += x.p; colT[k][1] += x.u; return `<td>${kh(x.p)}</td><td>${kh(x.u)}</td>`; }).join('');
    P += p; U += u;
    return `<tr><td>${kh(n + 1)}</td><td class="l">${esc(st.name)}</td>${cells}<td>${kh(p)}</td><td>${kh(u)}</td><td><b>${kh(p + u)}</b></td></tr>`;
  }).join('');
  const foot = colT.map(c => `<td>${kh(c[0])}</td><td>${kh(c[1])}</td>`).join('');
  return `${head}<div class="tw"><table class="t"><thead><tr><th rowspan="2">ល.រ</th><th rowspan="2" class="l">ឈ្មោះ</th>${th1}<th colspan="3">សរុប</th></tr><tr>${th2}<th>ច្បាប់</th><th>អត់</th><th>សរុប</th></tr></thead><tbody>${rows}</tbody><tfoot><tr><td colspan="2" class="l">សរុប</td>${foot}<td>${kh(P)}</td><td>${kh(U)}</td><td>${kh(P + U)}</td></tr></tfoot></table></div>
  <div class="row" style="margin-top:12px"><button class="btn primary grow" data-act="exportMonth">នាំចេញ Excel (ខែនេះ)</button></div>`;
}

function viewStudents() {
  const list = S.students.map(s => `<li class="stu"><span class="no">${kh(s.seq)}</span><span class="nm" style="${s.active ? '' : 'opacity:.5'}">${esc(s.name)}<small>${esc([s.sex, s.dob, s.active ? '' : 'លាក់'].filter(Boolean).join(' · '))}</small></span><button class="btn small" data-act="editStu" data-sid="${s.id}">កែ</button></li>`).join('');
  return `<div class="card"><h2>សិស្ស ${kh(S.students.length)} នាក់</h2><div class="row"><button class="btn primary" data-act="addStu">+ បន្ថែម</button><button class="btn" data-act="importStu">នាំចូលពី Excel</button><button class="btn" data-act="renumber">រៀបលេខឡើងវិញ</button></div><input id="fileIn" type="file" accept=".xlsx,.xls,.csv" hidden></div>
  <ul class="stulist">${list || '<li class="muted pad">មិនទាន់មានសិស្ស</li>'}</ul>`;
}

function viewTimetable() {
  const w = S.week;
  if (!S.ttDraft || S.ttDraft.weekId !== w.id) S.ttDraft = { weekId: w.id, tt: clone(w.timetable || {}) };
  const tt = S.ttDraft.tt, d = S.ttDay;
  const subs = new Set(); [S.settings.template, ...S.weeks.map(x => x.timetable), tt].forEach(t => Object.values(t || {}).forEach(day => ['am', 'pm'].forEach(s => (day[s] || []).forEach(x => x && subs.add(x)))));
  const days = DAYS.map((n, i) => `<button class="day ${i === d ? 'on' : ''}" data-act="ttDay" data-v="${i}"><span>${n}</span><small>${kh(dnum(addDays(w.start_date, i)))}</small></button>`).join('');
  const rows = slotsOfDay(tt, d).map(s => `<div class="ttrow"><label for="tt-${s.sess}${s.i}">${SESS_S[s.sess]} ${kh(s.i + 1)}</label><input id="tt-${s.sess}${s.i}" type="text" list="subs" data-tt="1" data-sess="${s.sess}" data-i="${s.i}" value="${esc(s.subject)}" placeholder="មុខវិជ្ជា (ទុកទទេ = គ្មាន)">${s.i > 0 ? `<button data-act="ttUp" data-sess="${s.sess}" data-i="${s.i}" aria-label="ចម្លងពីម៉ោងលើ" title="ចម្លងពីម៉ោងលើ">↓</button>` : '<span style="width:38px;flex:none"></span>'}</div>`).join('');
  return `<div class="card"><h2>កាលវិភាគ · សប្ដាហ៍ ${weekLabel(w)}</h2><p class="muted" style="margin:0 0 8px;font-size:.85rem">ប៊ូតុង ↓ ចម្លងមុខវិជ្ជាពីម៉ោងខាងលើ (សម្រាប់មុខវិជ្ជា ២ ម៉ោង)។ ការកែមិនប៉ះពាល់វត្តមានដែលបានកត់រួច ប៉ុន្តែម៉ោងដែលគ្មានមុខវិជ្ជានឹងមិនរាប់ក្នុងរបាយការណ៍។</p>
  <div class="daybar">${days}</div>${rows}<datalist id="subs">${[...subs].map(x => `<option value="${esc(x)}">`).join('')}</datalist>
  <div class="row"><button class="btn primary grow" data-act="ttSave">រក្សាទុកកាលវិភាគសប្ដាហ៍នេះ</button></div>
  <div class="row" style="margin-top:8px"><button class="btn small" data-act="ttToTemplate">ប្រើជាគំរូសម្រាប់សប្ដាហ៍ថ្មី</button><button class="btn small" data-act="ttFromTemplate">ដាក់គំរូមកសប្ដាហ៍នេះ</button></div></div>`;
}

function viewSettings() {
  const s = S.settings;
  return `<div class="card"><h2>ក្បាលឯកសារ និងចំណងជើង</h2>
    <label class="f" for="s-short">ឈ្មោះខ្លីក្នុងកម្មវិធី</label><input id="s-short" type="text" value="${esc(s.short_title)}">
    <label class="f" for="s-head">ក្បាលអត្ថបទ (មួយបន្ទាត់ក្នុងមួយជួរ)</label><textarea id="s-head">${esc((s.header_lines || []).join('\n'))}</textarea>
    <label class="f" for="s-class">ចំណងជើងរបាយការណ៍សប្ដាហ៍ (ខែ និងថ្ងៃ បន្ថែមដោយស្វ័យប្រវត្តិ)</label><textarea id="s-class">${esc(s.class_title)}</textarea>
    <label class="f" for="s-month">ចំណងជើងរបាយការណ៍ខែ</label><textarea id="s-month">${esc(s.month_title)}</textarea>
    <label class="f" for="s-major">ជំនាញ (បន្ទាត់ក្រោមចំណងជើងខែ)</label><input id="s-major" type="text" value="${esc(s.major)}">
    <label class="f" for="s-grid">ចំណងជើងតារាងវត្តមាន</label><input id="s-grid" type="text" value="${esc(s.grid_title)}">
    <label class="f" for="s-place">ធ្វើនៅ (ទីកន្លែង)</label><input id="s-place" type="text" value="${esc(s.place)}">
    <label class="f" for="s-taker">ឈ្មោះអ្នកស្រង់</label><input id="s-taker" type="text" value="${esc(s.taker)}"></div>
  <div class="card"><h2>ចំនួនម៉ោងក្នុងមួយថ្ងៃ</h2><div class="row"><div class="grow"><label class="f" for="s-am" style="margin-top:0">ពេលព្រឹក</label><input id="s-am" type="number" min="1" max="8" value="${s.morning}"></div><div class="grow"><label class="f" for="s-pm" style="margin-top:0">ពេលរសៀល</label><input id="s-pm" type="number" min="0" max="8" value="${s.afternoon}"></div></div></div>
  <div class="card"><h2>ម៉ោងពិសេស (ក្រៅកាលវិភាគ)</h2><p class="muted" style="margin:0 0 8px;font-size:.85rem">កត់ត្រាម្នាក់ៗ ✓/មានច្បាប់/អត់ច្បាប់ ដូចផ្ទាំង "កត់វត្តមាន" (ក្នុងផ្ទាំង "ម៉ោងពិសេស")។ ជ្រើស "ថ្ងៃប្រចាំ" បើកើតឡើងដដែលៗរៀងរាល់សប្ដាហ៍ ទុក "គ្មាន" បើមិនទៀងទាត់ — ទោះជ្រើសយ៉ាងណា នៅតែអាចកត់ថ្ងៃផ្សេងបានគ្រប់ពេល។</p>
    ${(s.specialTypes || []).map((t, i) => `<div class="row" style="margin-bottom:8px">
      <input class="grow" type="text" data-sp-label="${i}" value="${esc(t.label)}">
      <select data-sp-day="${i}" style="width:132px"><option value="" ${t.weekday == null ? 'selected' : ''}>គ្មាន (មិនទៀងទាត់)</option>${DAYS.map((n, d) => `<option value="${d}" ${t.weekday === d ? 'selected' : ''}>${n}</option>`).join('')}</select>
    </div>`).join('')}</div>
  <button class="btn primary" style="width:100%;margin-bottom:12px" data-act="saveSettings">រក្សាទុកការកំណត់</button>
  <div class="card"><h2>គណនី</h2><p class="muted" style="margin:0 0 8px">${esc(S.user && S.user.email || '')}</p><div class="row"><button class="btn" data-act="logout">ចាកចេញ</button><button class="btn danger" data-act="delWeek">លុបសប្ដាហ៍នេះ</button></div></div>`;
}

/* ---------- ប្រអប់ (dialog) ---------- */
function dialog(html) {
  const d = document.createElement('dialog'); d.className = 'sheet'; d.innerHTML = html;
  document.body.appendChild(d); d.addEventListener('close', () => d.remove()); d.showModal(); return d;
}
const closeBtn = '<button class="btn" data-act="closeDlg">បោះបង់</button>';

/* ---------- 4. សកម្មភាព ---------- */
function curSlot() {
  const slots = slotsOfDay(S.week.timetable, S.day);
  return slots.find(s => s.sess === S.per.sess && s.i === S.per.i) || slots[0];
}
function firstPeriod() {
  const s = slotsOfDay(S.week.timetable, S.day).find(x => x.subject) || slotsOfDay(S.week.timetable, S.day)[0];
  S.per = { sess: s.sess, i: s.i };
}
function pickWeek(w) {
  S.week = w; S.ttDraft = null;
  const t = todayISO(), off = Math.round((dateOf(t) - dateOf(w.start_date)) / 864e5);
  S.day = off >= 0 && off <= 4 ? off : 0; firstPeriod();
  S.month = w.start_date.slice(0, 7);
}

const ACTS = {
  tab: el => { S.tab = el.dataset.v; if (rtNeedShell) { rtNeedShell = false; renderShell(); } else renderView(); window.scrollTo(0, 0); },
  day: el => { S.day = +el.dataset.v; firstPeriod(); renderView(); },
  per: el => { S.per = { sess: el.dataset.sess, i: +el.dataset.i }; renderView(); },
  rep: el => { S.rep = el.dataset.v; renderView(); },
  ttDay: el => { S.ttDay = +el.dataset.v; renderView(); },
  closeDlg: el => el.closest('dialog').close(),

  cycle: el => {
    const s = curSlot(); if (!s || !s.subject) return;
    const cur = (S.abs[S.week.id] || new Map()).get(key(el.dataset.sid, S.day, s.sess, s.i));
    const next = cur === undefined ? 0 : cur === 0 ? 1 : null;   // មក → អត់ច្បាប់ → មានច្បាប់ → មក
    applyChanges([{ sid: el.dataset.sid, day: S.day, sess: s.sess, i: s.i, status: next }]);
  },
  allPresent: async () => {
    const s = curSlot(), m = S.abs[S.week.id], w = S.week;
    const has = activeStudents().some(st => m.has(key(st.id, S.day, s.sess, s.i)));
    if (!has) return toast('ម៉ោងនេះមកទាំងអស់ហើយ');
    if (!confirm('លុបអវត្តមានទាំងអស់ក្នុងម៉ោងនេះ?')) return;
    const ids = activeStudents().map(st => st.id), old = [];
    ids.forEach(id => { const k = key(id, S.day, s.sess, s.i); if (m.has(k)) old.push([k, m.get(k)]); m.delete(k); });
    renderView(); busy(1);
    try { await q(S.sb.from('absences').delete().eq('week_id', w.id).eq('day', S.day).eq('session', s.sess).eq('period', s.i).in('student_id', ids)); }
    catch (e) { old.forEach(([k, v]) => m.set(k, v)); renderView(); toast('រក្សាទុកមិនបាន៖ ' + errText(e), 'err'); }
    finally { busy(-1); }
  },
  copyPrev: () => {
    const s = curSlot(), m = S.abs[S.week.id]; if (!s || s.i < 1) return;
    const ch = [];
    activeStudents().forEach(st => {
      const src = m.get(key(st.id, S.day, s.sess, s.i - 1)), dst = m.get(key(st.id, S.day, s.sess, s.i));
      if (src !== dst) ch.push({ sid: st.id, day: S.day, sess: s.sess, i: s.i, status: src === undefined ? null : src });
    });
    if (!ch.length) return toast('ដូចម៉ោងមុនរួចហើយ');
    applyChanges(ch);
  },
  more: el => {
    const st = S.students.find(x => x.id === el.dataset.sid);
    const d = dialog(`<h2>${esc(st.name)}</h2><p class="muted" style="margin:0">${DAYS[S.day]} · អនុវត្តលើ៖</p>
      <div class="seg" style="margin:8px 0"><label><input type="radio" name="scope" value="one" checked>ម៉ោងនេះ</label><label><input type="radio" name="scope" value="am">ព្រឹកទាំងមូល</label><label><input type="radio" name="scope" value="pm">រសៀលទាំងមូល</label><label><input type="radio" name="scope" value="day">ពេញមួយថ្ងៃ</label></div>
      <p class="muted" style="margin:8px 0 0">ស្ថានភាព៖</p>
      <div class="seg" style="margin:8px 0"><label><input type="radio" name="stat" value="ok">✓ មក</label><label><input type="radio" name="stat" value="1">មានច្បាប់</label><label><input type="radio" name="stat" value="0" checked>អត់ច្បាប់</label></div>
      <div class="foot">${closeBtn}<button class="btn primary" data-act="applyMore" data-sid="${st.id}">អនុវត្ត</button></div>`);
    d.dataset.kind = 'more';
  },
  applyMore: el => {
    const d = el.closest('dialog'), scope = $('input[name=scope]:checked', d).value, stat = $('input[name=stat]:checked', d).value;
    const status = stat === 'ok' ? null : +stat, cur = curSlot();
    const slots = slotsOfDay(S.week.timetable, S.day).filter(s => s.subject && (scope === 'day' || (scope === 'one' ? s.sess === cur.sess && s.i === cur.i : s.sess === scope)));
    const m = S.abs[S.week.id];
    const ch = slots.map(s => ({ sid: el.dataset.sid, day: S.day, sess: s.sess, i: s.i, status })).filter(c => m.get(key(c.sid, c.day, c.sess, c.i)) !== (c.status == null ? undefined : c.status));
    d.close(); if (ch.length) applyChanges(ch); else toast('គ្មានអ្វីត្រូវប្តូរ');
  },

  /* ពិន្ទុ */
  scKind: el => { S.sc.kind = el.dataset.v; S.sc.sheetId = null; renderView(); },
  scSheet: el => { S.sc.sheetId = el.dataset.id; renderView(); },
  rsMonthAll: () => { S.rs.month = ''; renderView(); },
  newSheet: () => sheetDialog(null),
  editSheet: el => sheetDialog(S.sheets.find(x => x.id === el.dataset.id)),
  saveSheet: async el => {
    const d = el.closest('dialog'), id = el.dataset.id;
    const subject = $('#x-sub', d).value.trim(), maxes = {};
    if (!subject) return toast('សូមបញ្ចូលមុខវិជ្ជា', 'err');
    for (const [c, l] of COMPS) {
      const v = parseFloat($('#x-m-' + c, d).value);
      if (!(v >= 0)) return toast(`ពិន្ទុពេញ "${l}" មិនត្រឹមត្រូវ`, 'err');
      maxes[c] = v;
    }
    if (!COMPS.some(([c]) => maxes[c] > 0)) return toast('ត្រូវមានពិន្ទុពេញយ៉ាងតិចមួយសមាសធាតុ', 'err');
    if (curSheets().some(x => x.subject === subject && x.id !== id)) return toast('មុខវិជ្ជានេះមានរួចហើយក្នុងខែនេះ', 'err');
    try {
      if (id) {
        const bad = [...S.gs].find(([k, v]) => k.startsWith(id + '|') && COMPS.some(([c]) => v[c] != null && v[c] > maxes[c]));
        if (bad) return toast('មានពិន្ទុខ្លះធំជាងពិន្ទុពេញថ្មី — សូមកែពិន្ទុជាមុន', 'err');
        await q(S.sb.from('grade_sheets').update({ subject, maxes }).eq('id', id)); Object.assign(S.sheets.find(x => x.id === id), { subject, maxes });
      } else {
        const row = await q(S.sb.from('grade_sheets').insert({ kind: S.sc.kind, month: S.sc.month, subject, maxes }).select().single());
        S.sheets.push(row); S.sc.sheetId = row.id;
      }
      d.close(); renderView(); toast('បានរក្សាទុក');
    } catch (e) { toast('រក្សាទុកមិនបាន៖ ' + errText(e), 'err'); }
  },
  delSheet: async el => {
    const e = S.sheets.find(x => x.id === el.dataset.id);
    if (!confirm(`លុបមុខវិជ្ជា "${e.subject}" រួមទាំងពិន្ទុទាំងអស់របស់មុខវិជ្ជានេះ?`)) return;
    try {
      await q(S.sb.from('grade_sheets').delete().eq('id', e.id));
      S.sheets = S.sheets.filter(x => x.id !== e.id); [...S.gs.keys()].forEach(k => { if (k.startsWith(e.id + '|')) S.gs.delete(k); });
      S.sc.sheetId = null; el.closest('dialog').close(); renderView(); toast('បានលុប');
    } catch (err) { toast('លុបមិនបាន៖ ' + errText(err), 'err'); }
  },
  exportScores: () => {
    try {
      const sum = scoreSummary(S.rs.kind, S.rs.month);
      if (!sum.sheets.length) return toast('មិនមានពិន្ទុសម្រាប់នាំចេញទេ', 'err');
      download(buildScoreWorkbook(sum), `scores_${S.rs.kind}_${S.rs.month || 'all'}.xlsx`);
    } catch (e) { toast('នាំចេញមិនបាន៖ ' + errText(e), 'err'); }
  },

  /* ម៉ោងពិសេស (កត់ម្នាក់ៗ ដូចកត់វត្តមានធម្មតា) */
  spType: el => { S.spType = el.dataset.k; renderView(); },
  spCycle: el => {
    if (!S.spType) return;
    const type = S.spType, sid = el.dataset.sid;
    const cur = (S.specials[S.week.id] || new Map()).get(spKey(sid, S.day, type));
    const next = cur === undefined ? 0 : cur === 0 ? 1 : null;   // មក → អត់ច្បាប់ → មានច្បាប់ → មក
    applySpChanges([{ sid, day: S.day, type, status: next }]);
  },
  spAllPresent: async () => {
    const type = S.spType; if (!type) return;
    const m = S.specials[S.week.id], w = S.week;
    const has = activeStudents().some(st => m.has(spKey(st.id, S.day, type)));
    if (!has) return toast('ម៉ោងនេះមកទាំងអស់ហើយ');
    if (!confirm('លុបអវត្តមានទាំងអស់សម្រាប់ម៉ោងនេះ?')) return;
    const ids = activeStudents().map(st => st.id), old = [];
    ids.forEach(id => { const k = spKey(id, S.day, type); if (m.has(k)) old.push([k, m.get(k)]); m.delete(k); });
    renderView(); busy(1);
    try { await q(S.sb.from('special_absences').delete().eq('week_id', w.id).eq('day', S.day).eq('type', type).in('student_id', ids)); }
    catch (e) { old.forEach(([k, v]) => m.set(k, v)); renderView(); toast('រក្សាទុកមិនបាន៖ ' + errText(e), 'err'); }
    finally { busy(-1); }
  },

  /* សប្ដាហ៍ */
  newWeek: () => {
    const last = S.weeks.slice().sort((a, b) => a.start_date.localeCompare(b.start_date)).pop();
    const def = last ? addDays(last.start_date, 7) : mondayOf(todayISO());
    dialog(`<h2>បន្ថែមសប្ដាហ៍ថ្មី</h2><label class="f" for="nwDate">ជ្រើសថ្ងៃណាមួយក្នុងសប្ដាហ៍ (ប្រព័ន្ធនឹងយកថ្ងៃច័ន្ទ)</label><input id="nwDate" type="date" value="${def}">
      <p class="muted" style="font-size:.85rem">កាលវិភាគនឹងចម្លងពី${last ? 'សប្ដាហ៍ចុងក្រោយ' : 'គំរូ'} ហើយអាចកែបានក្នុងផ្ទាំង "កាលវិភាគ"។</p>
      <div class="foot">${closeBtn}<button class="btn primary" data-act="createWeek">បង្កើត</button></div>`);
  },
  createWeek: async el => {
    const d = el.closest('dialog'), v = $('#nwDate', d).value; if (!v) return;
    const start = mondayOf(v);
    if (S.weeks.some(w => w.start_date === start)) { d.close(); const w = S.weeks.find(x => x.start_date === start); pickWeek(w); renderShell(); return toast('សប្ដាហ៍នេះមានរួចហើយ'); }
    const last = S.weeks.slice().sort((a, b) => a.start_date.localeCompare(b.start_date)).pop();
    const tt = clone(last ? last.timetable : S.settings.template);
    try {
      const row = await q(S.sb.from('weeks').insert({ start_date: start, timetable: tt }).select().single());
      S.weeks.push(row); S.abs[row.id] = new Map(); pickWeek(row); d.close(); renderShell(); toast('បានបង្កើតសប្ដាហ៍');
    } catch (e) { toast('បង្កើតមិនបាន៖ ' + errText(e), 'err'); }
  },
  delWeek: async () => {
    if (!S.week || !confirm(`លុបសប្ដាហ៍ ${weekLabel(S.week)} រួមទាំងវត្តមានទាំងអស់ក្នុងសប្ដាហ៍នេះ? មិនអាចត្រឡប់វិញទេ។`)) return;
    try {
      await q(S.sb.from('weeks').delete().eq('id', S.week.id));
      S.weeks = S.weeks.filter(w => w.id !== S.week.id); delete S.abs[S.week.id]; delete S.specials[S.week.id];
      const last = S.weeks.slice().sort((a, b) => a.start_date.localeCompare(b.start_date)).pop();
      S.week = null; if (last) pickWeek(last); S.tab = 'record'; renderShell(); toast('បានលុប');
    } catch (e) { toast('លុបមិនបាន៖ ' + errText(e), 'err'); }
  },

  /* សិស្ស */
  addStu: () => stuDialog(null),
  editStu: el => stuDialog(S.students.find(s => s.id === el.dataset.sid)),
  saveStu: async el => {
    const d = el.closest('dialog'), id = el.dataset.sid;
    const rec = { name: $('#f-name', d).value.trim(), sex: $('#f-sex', d).value, dob: $('#f-dob', d).value.trim(), seq: parseInt($('#f-seq', d).value, 10) || 0, active: $('#f-act', d).checked };
    if (!rec.name) return toast('សូមបញ្ចូលឈ្មោះ', 'err');
    try {
      if (id) { await q(S.sb.from('students').update(rec).eq('id', id)); Object.assign(S.students.find(s => s.id === id), rec); }
      else { if (!rec.seq) rec.seq = Math.max(0, ...S.students.map(s => s.seq)) + 1; S.students.push(await q(S.sb.from('students').insert(rec).select().single())); }
      S.students.sort((a, b) => a.seq - b.seq); d.close(); renderView(); toast('បានរក្សាទុក');
    } catch (e) { toast('រក្សាទុកមិនបាន៖ ' + errText(e), 'err'); }
  },
  delStu: async el => {
    const st = S.students.find(s => s.id === el.dataset.sid);
    if (!confirm(`លុប "${st.name}" រួមទាំងវត្តមានទាំងអស់របស់សិស្សនេះ? (បើគ្រាន់តែឈប់រៀន សូមប្រើ "លាក់" ជំនួស)`)) return;
    try {
      await q(S.sb.from('students').delete().eq('id', st.id));
      S.students = S.students.filter(s => s.id !== st.id); S.abs = {}; S.specials = {}; S.gs = null; el.closest('dialog').close(); renderView(); toast('បានលុប');
    } catch (e) { toast('លុបមិនបាន៖ ' + errText(e), 'err'); }
  },
  renumber: async () => {
    if (!confirm('រៀបលេខ ល.រ ឡើងវិញ ១, ២, ៣ … តាមលំដាប់បច្ចុប្បន្ន?')) return;
    try {
      const list = S.students.slice().sort((a, b) => a.seq - b.seq);
      for (let i = 0; i < list.length; i++) if (list[i].seq !== i + 1) { await q(S.sb.from('students').update({ seq: i + 1 }).eq('id', list[i].id)); list[i].seq = i + 1; }
      renderView(); toast('រៀបលេខរួច');
    } catch (e) { toast(errText(e), 'err'); }
  },
  importStu: () => $('#fileIn').click(),
  doImport: async el => {
    const d = el.closest('dialog'), rows = d._rows; let seq = Math.max(0, ...S.students.map(s => s.seq));
    try {
      const data = await q(S.sb.from('students').insert(rows.map(r => ({ seq: ++seq, name: r.name, sex: r.sex, dob: r.dob }))).select());
      S.students = S.students.concat(data).sort((a, b) => a.seq - b.seq); d.close(); renderView(); toast(`បាននាំចូល ${kh(data.length)} នាក់`);
    } catch (e) { toast('នាំចូលមិនបាន៖ ' + errText(e), 'err'); }
  },

  /* កាលវិភាគ */
  ttUp: el => { readTT(); const tt = S.ttDraft.tt, d = S.ttDay, a = el.dataset.sess, i = +el.dataset.i; tt[d] = tt[d] || { am: [], pm: [] }; tt[d][a] = tt[d][a] || []; tt[d][a][i] = tt[d][a][i - 1] || ''; renderView(); },
  ttSave: async () => {
    readTT(); const tt = S.ttDraft.tt;
    try { await q(S.sb.from('weeks').update({ timetable: tt }).eq('id', S.week.id)); S.week.timetable = clone(tt); toast('បានរក្សាទុកកាលវិភាគ'); }
    catch (e) { toast('រក្សាទុកមិនបាន៖ ' + errText(e), 'err'); }
  },
  ttToTemplate: async () => {
    readTT(); S.settings.template = clone(S.ttDraft.tt);
    try { await saveSettings(); toast('បានរក្សាទុកជាគំរូ'); } catch (e) { toast(errText(e), 'err'); }
  },
  ttFromTemplate: () => { if (confirm('ជំនួសកាលវិភាគសប្ដាហ៍នេះដោយគំរូ? (ត្រូវរក្សាទុកទើបមានប្រសិទ្ធភាព)')) { S.ttDraft.tt = clone(S.settings.template); renderView(); } },

  /* ការកំណត់ */
  saveSettings: async () => {
    const s = S.settings, v = id => $('#' + id).value;
    const am = Math.min(8, Math.max(1, parseInt(v('s-am'), 10) || 4)), pm = Math.min(8, Math.max(0, parseInt(v('s-pm'), 10) || 0));
    const specialTypes = (s.specialTypes || []).map((t, i) => {
      const lab = $('[data-sp-label="' + i + '"]'), day = $('[data-sp-day="' + i + '"]');
      return { key: t.key, label: (lab ? lab.value.trim() : '') || t.label, weekday: day && day.value !== '' ? +day.value : null };
    });
    Object.assign(s, { short_title: v('s-short').trim(), header_lines: v('s-head').split('\n').map(x => x.trim()).filter(Boolean), class_title: v('s-class').trim(), month_title: v('s-month').trim(), major: v('s-major').trim(), grid_title: v('s-grid').trim(), place: v('s-place').trim(), taker: v('s-taker').trim(), morning: am, afternoon: pm, specialTypes });
    try { await saveSettings(); S.ttDraft = null; renderShell(); toast('បានរក្សាទុកការកំណត់'); } catch (e) { toast('រក្សាទុកមិនបាន៖ ' + errText(e), 'err'); }
  },
  logout: async () => { await S.sb.auth.signOut(); },
  exportWeek: () => { try { download(buildWeekWorkbook(S.week), `attendance_${S.week.start_date}.xlsx`); } catch (e) { toast('នាំចេញមិនបាន៖ ' + errText(e), 'err'); } },
  exportMonth: async () => {
    const ws = monthWeeks(); if (!ws.length) return toast('មិនមានសប្ដាហ៍ក្នុងខែនេះ', 'err');
    try { await ensureAbs(ws); download(buildMonthWorkbook(ws, S.month), `attendance_${S.month}.xlsx`); } catch (e) { toast('នាំចេញមិនបាន៖ ' + errText(e), 'err'); }
  }
};

function stuDialog(st) {
  dialog(`<h2>${st ? 'កែព័ត៌មានសិស្ស' : 'បន្ថែមសិស្ស'}</h2>
    <label class="f" for="f-name">ឈ្មោះ</label><input id="f-name" type="text" value="${esc(st ? st.name : '')}">
    <label class="f" for="f-sex">ភេទ</label><select id="f-sex"><option value=""></option><option ${st && st.sex === 'ប្រុស' ? 'selected' : ''}>ប្រុស</option><option ${st && st.sex === 'ស្រី' ? 'selected' : ''}>ស្រី</option></select>
    <label class="f" for="f-dob">ថ្ងៃខែឆ្នាំកំណើត</label><input id="f-dob" type="text" value="${esc(st ? st.dob : '')}">
    <label class="f" for="f-seq">ល.រ</label><input id="f-seq" type="number" min="1" value="${st ? st.seq : ''}" placeholder="ទុកទទេ = បន្ថែមចុងបញ្ជី">
    <label class="f"><input id="f-act" type="checkbox" ${!st || st.active ? 'checked' : ''}> នៅរៀន (ដកធីក = លាក់ពីរបាយការណ៍)</label>
    <div class="foot">${st ? `<button class="btn danger" data-act="delStu" data-sid="${st.id}">លុប</button>` : ''}<span class="grow"></span>${closeBtn}<button class="btn primary" data-act="saveStu" data-sid="${st ? st.id : ''}">រក្សាទុក</button></div>`);
}
function readTT() {
  if (!S.ttDraft) return;
  const tt = S.ttDraft.tt, d = S.ttDay; tt[d] = tt[d] || { am: [], pm: [] };
  $$('input[data-tt]').forEach(inp => { const a = inp.dataset.sess; tt[d][a] = tt[d][a] || []; tt[d][a][+inp.dataset.i] = inp.value.trim(); });
}

/* នាំចូលសិស្សពី Excel: រកជួរ "ល.រ" រួចអានលេខ/ឈ្មោះ/ភេទ/ថ្ងៃកំណើត */
function parseStudents(wb) {
  for (const name of wb.SheetNames) {
    const aoa = XLSX.utils.sheet_to_json(wb.Sheets[name], { header: 1, defval: '' });
    let h = aoa.findIndex(r => String(r[0]).trim() === 'ល.រ' || /ឈ្មោះ|គោត្តនាម|name/i.test(String(r[1])));
    if (h < 0) continue;
    while (h + 1 < aoa.length && !/^\d+$/.test(unkh(aoa[h + 1][0]))) { if (h > 30) break; h++; }
    const out = [];
    for (let r = h + 1; r < aoa.length && /^\d+$/.test(unkh(aoa[r][0])); r++) {
      const nm = String(aoa[r][1] || '').replace(/\s+/g, ' ').trim(); if (!nm) continue;
      out.push({ name: nm, sex: String(aoa[r][2] || '').trim(), dob: String(aoa[r][3] || '').replace(/[\u200b]/g, '').replace(/\s+/g, ' ').trim() });
    }
    if (out.length) return out;
  }
  return [];
}
function onFile(file) {
  const rd = new FileReader();
  rd.onload = () => {
    try {
      const rows = parseStudents(XLSX.read(rd.result, { type: 'array' }));
      if (!rows.length) return toast('រកមិនឃើញបញ្ជីសិស្ស (ត្រូវមានជួរ "ល.រ" និង "ឈ្មោះ")', 'err');
      const f = rows.filter(r => r.sex === 'ស្រី').length;
      const d = dialog(`<h2>នាំចូលសិស្ស ${kh(rows.length)} នាក់</h2><p class="muted" style="margin:0">ស្រី ${kh(f)} នាក់ · ${S.students.length ? `⚠️ មានសិស្ស ${kh(S.students.length)} នាក់រួចហើយ (នឹងបន្ថែមបន្ទាប់ ប្រយ័ត្នស្ទួន)` : 'បញ្ជីទទេ'}</p>
        <div class="prev"><table class="t"><tbody>${rows.map((r, n) => `<tr><td>${kh(n + 1)}</td><td class="l">${esc(r.name)}</td><td>${esc(r.sex)}</td></tr>`).join('')}</tbody></table></div>
        <div class="foot">${closeBtn}<button class="btn primary" data-act="doImport">នាំចូល</button></div>`);
      d._rows = rows;
    } catch (e) { toast('អានឯកសារមិនបាន៖ ' + errText(e), 'err'); }
  };
  rd.readAsArrayBuffer(file);
}

/* ---------- 5. Excel (xlsx-js-style) ---------- */
const FONT = 'Battambang';
const thin = { style: 'thin', color: { rgb: '000000' } };
const BD = { top: thin, bottom: thin, left: thin, right: thin };
const mk = (o = {}) => ({ font: { name: FONT, sz: o.sz || 11, bold: !!o.bold }, alignment: { horizontal: o.h || 'center', vertical: 'center', wrapText: true, textRotation: o.rot || 0 }, border: o.nb ? undefined : BD, fill: o.fill ? { patternType: 'solid', fgColor: { rgb: o.fill } } : undefined });
const ST = {
  title: mk({ bold: true, nb: true, sz: 12 }), left: mk({ h: 'left', nb: true }), leftB: mk({ h: 'left', nb: true, bold: true }),
  head: mk({ bold: true, fill: 'DDEBF7' }), rot: mk({ bold: true, fill: 'DDEBF7', rot: 90 }), name: mk({ h: 'left' }), cell: mk(),
  p2: mk({ fill: 'E2F0D9' }), p1: mk({ fill: 'FFF2CC' }), p0: mk({ fill: 'F8CBAD' }), foot: mk({ bold: true, fill: 'F2F2F2' }), footL: mk({ bold: true, h: 'left', fill: 'F2F2F2' })
};
class Sheet {
  constructor() { this.c = {}; this.merges = []; this.cols = []; this.rows = []; this.R = 0; this.C = 0; }
  set(r, c, v, s, f) {
    const cell = typeof v === 'number' ? { t: 'n', v } : { t: 's', v: v == null ? '' : String(v) };
    if (f) cell.f = f; if (s) cell.s = s;
    this.c[XLSX.utils.encode_cell({ r, c })] = cell; this.R = Math.max(this.R, r); this.C = Math.max(this.C, c);
  }
  merge(r1, c1, r2, c2, s) {
    this.merges.push({ s: { r: r1, c: c1 }, e: { r: r2, c: c2 } });
    for (let r = r1; r <= r2; r++) for (let c = c1; c <= c2; c++) if ((r !== r1 || c !== c1) && !this.c[XLSX.utils.encode_cell({ r, c })]) this.set(r, c, '', s);
  }
  ws() { const w = Object.assign({}, this.c); w['!ref'] = XLSX.utils.encode_range({ s: { r: 0, c: 0 }, e: { r: this.R, c: this.C } }); w['!merges'] = this.merges; w['!cols'] = this.cols; w['!rows'] = this.rows; return w; }
}
const col = c => XLSX.utils.encode_col(c);

function weekTitle(w) {
  const a = dateOf(w.start_date), b = dateOf(addDays(w.start_date, 4));
  const same = a.getUTCMonth() === b.getUTCMonth();
  const tail = same ? `ប្រចាំខែ ${MONTHS[a.getUTCMonth()]} ថ្ងៃទី ${kh(pad2(a.getUTCDate()))} ដល់ថ្ងៃទី ${kh(pad2(b.getUTCDate()))}`
    : `ចាប់ពីថ្ងៃទី ${kh(pad2(a.getUTCDate()))} ខែ${MONTHS[a.getUTCMonth()]} ដល់ថ្ងៃទី ${kh(pad2(b.getUTCDate()))} ខែ${MONTHS[b.getUTCMonth()]}`;
  return `${S.settings.class_title} ${tail}`;
}
function headerLines(sh, last) {
  (S.settings.header_lines || []).slice(0, 5).forEach((t, r) => { sh.set(r, 0, t, r < 2 ? ST.title : ST.leftB); if (r < 2) sh.merge(r, 0, r, last, ST.title); else sh.merge(r, 0, r, 3, ST.leftB); });
}

function buildWeekSheet(w, students) {
  const m = S.abs[w.id] || new Map(), A = S.settings.morning, B = S.settings.afternoon, per = A + B, G = 5 * per;
  const c0 = 4, cP = c0 + G, cU = cP + 1, sh = new Sheet(), N = students.length, r0 = 14, rEnd = r0 + N - 1;
  headerLines(sh, cU);
  sh.set(5, 0, weekTitle(w), ST.title); sh.merge(5, 0, 5, cU, ST.title);
  ['ល.រ', 'គោត្តនាមនិងនាម', 'ភេទ', 'ថ្ងៃខែឆ្នាំកំណើត'].forEach((t, c) => { sh.set(6, c, t, ST.head); sh.merge(6, c, 13, c, ST.head); });
  sh.set(6, c0, S.settings.grid_title, ST.head); sh.merge(6, c0, 6, cP - 1, ST.head);
  sh.set(6, cP, 'សរុប', ST.head); sh.merge(6, cP, 7, cU, ST.head);
  sh.set(8, cP, 'មានច្បាប់', ST.rot); sh.merge(8, cP, 13, cP, ST.rot);
  sh.set(8, cU, 'គ្មានច្បាប់', ST.rot); sh.merge(8, cU, 13, cU, ST.rot);
  const colInfo = [];
  for (let d = 0; d < 5; d++) {
    const cs = c0 + d * per;
    sh.set(7, cs, `${DAYS[d]} ${kh(pad2(dnum(addDays(w.start_date, d))))}`, ST.head); sh.merge(7, cs, 7, cs + per - 1, ST.head);
    sh.set(8, cs, SESS.am, ST.head); sh.merge(8, cs, 8, cs + A - 1, ST.head);
    if (B) { sh.set(8, cs + A, SESS.pm, ST.head); sh.merge(8, cs + A, 8, cs + per - 1, ST.head); }
    for (const sess of ['am', 'pm']) {
      const n = sess === 'am' ? A : B, base = cs + (sess === 'am' ? 0 : A);
      for (let i = 0; i < n;) {
        const subj = ttGet(w.timetable, d, sess, i); let j = i;
        while (subj && j + 1 < n && ttGet(w.timetable, d, sess, j + 1) === subj) j++;
        sh.set(9, base + i, subj, ST.rot); sh.merge(9, base + i, 13, base + j, ST.rot); i = j + 1;
      }
      for (let i = 0; i < n; i++) colInfo.push({ c: base + i, d, sess, i, subj: ttGet(w.timetable, d, sess, i) });
    }
  }
  students.forEach((st, n) => {
    const r = r0 + n; let p = 0, u = 0;
    sh.set(r, 0, n + 1, ST.cell); sh.set(r, 1, st.name, ST.name); sh.set(r, 2, st.sex || '', ST.cell); sh.set(r, 3, st.dob || '', ST.cell);
    colInfo.forEach(ci => {
      if (!ci.subj) return sh.set(r, ci.c, '', ST.cell);
      const v = m.get(key(st.id, ci.d, ci.sess, ci.i)), val = v === undefined ? 2 : v;
      if (val === 1) p++; else if (val === 0) u++;
      sh.set(r, ci.c, val, val === 2 ? ST.p2 : val === 1 ? ST.p1 : ST.p0);
    });
    const rng = `${col(c0)}${r + 1}:${col(cP - 1)}${r + 1}`;
    sh.set(r, cP, p, ST.cell, `COUNTIF(${rng},1)`); sh.set(r, cU, u, ST.cell, `COUNTIF(${rng},0)`);
  });
  const rf = r0 + N, fem = students.filter(s => s.sex === 'ស្រី').length;
  sh.set(rf, 0, `សិស្សសរុប ${kh(N)}នាក់ ស្រី${kh(fem)}នាក់`, ST.footL); sh.merge(rf, 0, rf, 3, ST.footL);
  sh.set(rf, c0, 'អវត្តមានតាមម៉ោង', ST.foot); sh.merge(rf, c0, rf, cP - 1, ST.foot);
  [['អវត្តមានមានច្បាប់', 1], ['អវត្តមានអត់ច្បាប់', 0]].forEach(([lab, code], k) => {
    const r = rf + 1 + k; sh.set(r, 0, lab, ST.footL); sh.merge(r, 0, r, 3, ST.footL);
    for (let c = c0; c < cP; c++) {
      const ci = colInfo.find(x => x.c === c); let cnt = 0;
      if (ci && ci.subj) students.forEach(st => { if (m.get(key(st.id, ci.d, ci.sess, ci.i)) === code) cnt++; });
      sh.set(r, c, cnt, ST.foot, N ? `COUNTIF(${col(c)}${r0 + 1}:${col(c)}${rEnd + 1},${code})` : undefined);
    }
  });
  const rt = rf + 3; sh.set(rt, 0, 'អវត្តមានសរុបតាមម៉ោងនីមួយៗ', ST.footL); sh.merge(rt, 0, rt, 3, ST.footL);
  for (let c = c0; c < cP; c++) sh.set(rt, c, (sh.c[XLSX.utils.encode_cell({ r: rt - 2, c })].v || 0) + (sh.c[XLSX.utils.encode_cell({ r: rt - 1, c })].v || 0), ST.foot, `SUM(${col(c)}${rt - 1}:${col(c)}${rt})`);
  const fri = addDays(w.start_date, 4), fd = dateOf(fri), sc = Math.max(c0, cP - 12);
  [[`ធ្វើនៅ${S.settings.place}, ថ្ងៃសុក្រ`, ST.title], [`ត្រូវនឹងថ្ងៃទី${kh(fd.getUTCDate())} ខែ${MONTHS[fd.getUTCMonth()]} ឆ្នាំ${kh(fd.getUTCFullYear())}`, ST.title], ['អ្នកស្រង់', ST.title]].forEach(([t, s], k) => { const r = rt + 2 + k; sh.set(r, sc, t, s); sh.merge(r, sc, r, cP - 1, s); });
  sh.set(rt + 7, sc, S.settings.taker || '', ST.title); sh.merge(rt + 7, sc, rt + 7, cP - 1, ST.title);
  sh.cols = [{ wch: 5 }, { wch: 26 }, { wch: 7 }, { wch: 18 }].concat(Array(G).fill({ wch: 5.5 }), [{ wch: 8 }, { wch: 8 }]);
  sh.rows = []; for (let r = 9; r <= 13; r++) sh.rows[r] = { hpt: 26 };
  return sh;
}
function buildMonthSheet(ws, month, students) {
  const W = ws.length, last = 4 + 2 * W + 2, sh = new Sheet(), N = students.length, r0 = 9;
  headerLines(sh, last);
  const mi = +month.slice(5, 7) - 1;
  sh.set(5, 0, S.settings.month_title, ST.title); sh.merge(5, 0, 5, last, ST.title);
  sh.set(6, 0, S.settings.major, ST.title); sh.merge(6, 0, 6, last, ST.title);
  sh.set(7, 0, `អវត្តមានប្រចាំសប្ដាហ៍នីមួយៗ ក្នុងខែ${MONTHS[mi]}`, ST.title); sh.merge(7, 0, 7, last, ST.title);
  ['ល.រ', 'ឈ្មោះ', 'ភេទ', 'ថ្ងៃខែឆ្នាំកំណើត'].forEach((t, c) => { sh.set(8, c, t, ST.head); sh.merge(8, c, 9, c, ST.head); });
  const cT = 4 + 2 * W;
  ws.forEach((_, k) => { sh.set(8, 4 + 2 * k, `សប្ដាហ៍ទី${kh(k + 1)}`, ST.head); sh.merge(8, 4 + 2 * k, 8, 5 + 2 * k, ST.head); sh.set(9, 4 + 2 * k, 'មានច្បាប់', ST.head); sh.set(9, 5 + 2 * k, 'អត់ច្បាប់', ST.head); });
  sh.set(8, cT, 'សរុប', ST.head); sh.merge(8, cT, 8, cT + 2, ST.head); sh.set(9, cT, 'ច្បាប់', ST.head); sh.set(9, cT + 1, 'អត់ច្បាប់', ST.head); sh.set(9, cT + 2, 'សរុប', ST.head);
  const tots = ws.map(weekTotals);
  const pcol = 4 + 5 * (S.settings.morning + S.settings.afternoon);   // ជួរ "មានច្បាប់" ក្នុង Week sheet
  students.forEach((st, n) => {
    const r = r0 + 1 + n; let p = 0, u = 0;
    sh.set(r, 0, n + 1, ST.cell); sh.set(r, 1, st.name, ST.name); sh.set(r, 2, st.sex || '', ST.cell); sh.set(r, 3, st.dob || '', ST.cell);
    const pc = [], uc = [];
    ws.forEach((_, k) => {
      const x = tots[k].get(st.id), wr = 15 + n; p += x.p; u += x.u;
      sh.set(r, 4 + 2 * k, x.p, ST.cell, `Week${k + 1}!${col(pcol)}${wr}`); sh.set(r, 5 + 2 * k, x.u, ST.cell, `Week${k + 1}!${col(pcol + 1)}${wr}`);
      pc.push(col(4 + 2 * k) + (r + 1)); uc.push(col(5 + 2 * k) + (r + 1));
    });
    sh.set(r, cT, p, ST.cell, `SUM(${pc.join(',')})`); sh.set(r, cT + 1, u, ST.cell, `SUM(${uc.join(',')})`);
    sh.set(r, cT + 2, p + u, ST.cell, `SUM(${col(cT)}${r + 1}:${col(cT + 1)}${r + 1})`);
  });
  const rf = r0 + 1 + N, fem = students.filter(s => s.sex === 'ស្រី').length;
  sh.set(rf, 0, `សិស្សសរុប ${kh(N)} នាក់ ស្រី ${kh(fem)} នាក់`, ST.footL); sh.merge(rf, 0, rf, 3, ST.footL);
  sh.cols = [{ wch: 5 }, { wch: 26 }, { wch: 7 }, { wch: 18 }].concat(Array(2 * W + 3).fill({ wch: 10 }));
  return sh;
}
function buildWeekWorkbook(w) {
  const wb = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb, buildWeekSheet(w, activeStudents()).ws(), 'Week1'); return wb;
}
function buildMonthWorkbook(ws, month) {
  const wb = XLSX.utils.book_new(), stu = activeStudents();
  ws.forEach((w, k) => XLSX.utils.book_append_sheet(wb, buildWeekSheet(w, stu).ws(), `Week${k + 1}`));
  XLSX.utils.book_append_sheet(wb, buildMonthSheet(ws, month, stu).ws(), `សរុបប្រចាំខែ${MONTHS[+month.slice(5, 7) - 1]}`);
  return wb;
}
function download(wb, name) { XLSX.writeFile(wb, name); }

/* ---------- 6. ចាប់ផ្តើម ---------- */
// អ្នកមើល (viewer) អាចប្រើបានតែសកម្មភាពអានទាំងនេះ · សកម្មភាពផ្សេងទៀត (កែ/បន្ថែម/លុប) ត្រូវបានទប់ស្កាត់
// នៅត្រង់នេះជាចំណុចតែមួយ បន្ថែមលើ RLS ខាង Supabase ដែលជាស្រទាប់ការពារពិតប្រាកដ។
const VIEWER_SAFE_ACTS = new Set(['tab', 'day', 'rep', 'closeDlg', 'exportWeek', 'exportMonth', 'exportScores', 'rsMonthAll', 'logout']);
document.addEventListener('click', e => {
  const el = e.target.closest('[data-act]'); if (!el) return;
  const act = el.dataset.act;
  if (isViewer() && !VIEWER_SAFE_ACTS.has(act)) { e.preventDefault(); toast('គណនីអ្នកមើល អាចមើលរបាយការណ៍បានតែប៉ុណ្ណោះ', 'err'); return; }
  const f = ACTS[act]; if (f) { e.preventDefault(); Promise.resolve(f(el, e)).catch(err => toast(errText(err), 'err')); }
});
document.addEventListener('change', e => {
  const t = e.target;
  if (t.id === 'weekSel') { const w = S.weeks.find(x => x.id === t.value); if (w) { pickWeek(w); renderView(); } }
  else if (t.id === 'monthSel') { S.month = t.value; renderView(); }
  else if (t.id === 'rsKind') { S.rs.kind = t.value; renderView(); }
  else if (t.id === 'rsMonth') { S.rs.month = t.value; renderView(); }
  else if (t.id === 'scMonth' && !isViewer()) { S.sc.month = t.value; S.sc.sheetId = null; renderView(); }
  else if (t.classList && t.classList.contains('scin') && !isViewer()) saveScore(t);
  else if (t.id === 'fileIn' && t.files[0]) { if (isViewer()) { t.value = ''; return; } onFile(t.files[0]); t.value = ''; }
});
document.addEventListener('keydown', e => {
  if (e.key !== 'Enter' || !e.target.classList || !e.target.classList.contains('scin')) return;
  e.preventDefault();
  const colIn = $$('.scin:not([disabled])').filter(x => x.dataset.c === e.target.dataset.c), i = colIn.indexOf(e.target);
  if (colIn[i + 1]) colIn[i + 1].focus(); else e.target.blur();
});
document.addEventListener('submit', async e => {
  if (e.target.id !== 'loginForm') return; e.preventDefault();
  const btn = $('button', e.target); btn.disabled = true;
  const { data, error } = await S.sb.auth.signInWithPassword({ email: $('#em').value.trim(), password: $('#pw').value });
  if (error) {
    const bad = /invalid login credentials/i.test(error.message);
    const cfgHint = error.status === 404 || error.status === 401 || /path|api key|fetch/i.test(error.message);
    showLogin(bad ? 'អ៊ីមែល ឬលេខសម្ងាត់មិនត្រឹមត្រូវ' : `${error.message}${cfgHint ? ' — ពិនិត្យ SUPABASE_URL និង SUPABASE_ANON_KEY ក្នុង config.js' : ''}`);
    return;
  }
  S.user = data.user; await boot();
});

/* ---------- Realtime: ធ្វើបច្ចុប្បន្នភាពភ្លាមៗពេលមានទិន្នន័យថ្មីពីអ្នកដទៃ ---------- */
const RT_TABLES = ['settings', 'students', 'weeks', 'absences', 'special_absences', 'grade_sheets', 'grade_scores'];
let rtChannel = null, rtTimer = null, rtSeen = false, rtNeedShell = false;
const rtTouch = new Map();
const rtMark = k => rtTouch.set(k, Date.now());               // ការកែរបស់ខ្លួនឯង → មិនត្រូវ echo ត្រឡប់មកជាន់វិញ
const rtOwn = k => Date.now() - (rtTouch.get(k) || 0) < 2500;

function rtSchedule(shell) {
  rtNeedShell = rtNeedShell || !!shell;
  clearTimeout(rtTimer);
  rtTimer = setTimeout(rtFlush, 250);
}
function rtFlush() {
  if (!S.user || !$('#view')) return;
  if (document.querySelector('dialog[open]')) { rtTimer = setTimeout(rtFlush, 1500); return; }   // កុំរំខានពេលបើកផ្ទាំងកែ
  if (document.activeElement && document.activeElement.classList.contains('scin')) { rtTimer = setTimeout(rtFlush, 1500); return; }   // កំពុងវាយពិន្ទុ
  if (S.tab === 'settings' || S.tab === 'timetable') return;                                      // កុំជាន់អ្វីដែលកំពុងវាយ
  if (rtNeedShell) { rtNeedShell = false; renderShell(); } else renderView();
}
function rtHandle(table, p) {
  const ev = p.eventType, n = p.new || {}, o = p.old || {}, r = ev === 'DELETE' ? o : n;
  if (table === 'absences') {
    const m = S.abs[r.week_id]; if (!m) return;
    const k = key(r.student_id, r.day, r.session, r.period); if (rtOwn(k)) return;
    ev === 'DELETE' ? m.delete(k) : m.set(k, n.status);
    if (S.tab === 'record' || S.tab === 'reports') rtSchedule(false);
  } else if (table === 'special_absences') {
    const m = S.specials[r.week_id]; if (!m) return;
    const k = spKey(r.student_id, r.day, r.type); if (rtOwn(k)) return;
    ev === 'DELETE' ? m.delete(k) : m.set(k, n.status);
    if (S.tab === 'specials') rtSchedule(false);
  } else if (table === 'grade_sheets') {
    if (ev === 'DELETE') { S.sheets = S.sheets.filter(x => x.id !== o.id); if (S.gs) [...S.gs.keys()].forEach(k => { if (k.startsWith(o.id + '|')) S.gs.delete(k); }); }
    else { const i = S.sheets.findIndex(x => x.id === n.id); if (i < 0) S.sheets.push(n); else Object.assign(S.sheets[i], n); }
    if (S.tab === 'scores' || S.tab === 'reports') rtSchedule(false);
  } else if (table === 'grade_scores') {
    if (!S.gs) return;
    const k = r.sheet_id + '|' + r.student_id; if (rtOwn('gs|' + k)) return;
    if (ev === 'DELETE') S.gs.delete(k); else S.gs.set(k, normScore(n));
    if (S.tab === 'scores' || S.tab === 'reports') rtSchedule(false);
  } else if (table === 'students') {
    if (ev === 'DELETE') S.students = S.students.filter(x => x.id !== o.id);
    else { const i = S.students.findIndex(x => x.id === n.id); if (i < 0) S.students.push(n); else Object.assign(S.students[i], n); S.students.sort((a, b) => a.seq - b.seq); }
    rtSchedule(true);
  } else if (table === 'weeks') {
    if (ev === 'DELETE') {
      S.weeks = S.weeks.filter(x => x.id !== o.id); delete S.abs[o.id]; delete S.specials[o.id];
      if (S.week && S.week.id === o.id) {
        const last = S.weeks.slice().sort((a, b) => a.start_date.localeCompare(b.start_date)).pop();
        S.week = null; if (last) pickWeek(last);
      }
    } else {
      const i = S.weeks.findIndex(x => x.id === n.id);
      if (i < 0) S.weeks.push(n); else Object.assign(S.weeks[i], n);
    }
    rtSchedule(true);
  } else if (table === 'settings') {
    if (ev !== 'DELETE' && n.data) { S.settings = Object.assign(clone(DEFAULT_SETTINGS), n.data); rtSchedule(true); }
  }
}
async function rtResync() {
  if (!S.user || S.pending > 0) return;
  try {
    const id = S.week && S.week.id;
    await loadCore(); S.abs = {}; S.specials = {}; S.gs = null;
    const w = S.weeks.find(x => x.id === id);
    if (w) S.week = w;
    else { S.week = null; const last = S.weeks.slice().sort((a, b) => a.start_date.localeCompare(b.start_date)).pop(); if (last) pickWeek(last); }
    rtSchedule(true);
  } catch (e) { /* ព្យាយាមម្តងទៀតពេលក្រោយ */ }
}
function startRealtime() {
  stopRealtime();
  let ch = S.sb.channel('att-live');
  RT_TABLES.forEach(t => { ch = ch.on('postgres_changes', { event: '*', schema: 'public', table: t }, p => rtHandle(t, p)); });
  ch.subscribe(status => { if (status === 'SUBSCRIBED') { if (rtSeen) rtResync(); rtSeen = true; } });
  rtChannel = ch;
}
function stopRealtime() { if (rtChannel && S.sb) S.sb.removeChannel(rtChannel); rtChannel = null; rtSeen = false; }
document.addEventListener('visibilitychange', () => { if (!document.hidden) rtResync(); });   // ត្រឡប់មកកម្មវិធីវិញ → ទាញទិន្នន័យថ្មី

async function boot() {
  app.innerHTML = '<p class="boot">កំពុងផ្ទុក…</p>';
  try {
    await loadProfile();
    await loadCore();
    const sorted = S.weeks.slice().sort((a, b) => a.start_date.localeCompare(b.start_date));
    const cur = mondayOf(todayISO());
    const w = sorted.find(x => x.start_date === cur) || sorted[sorted.length - 1] || null;
    S.week = null; if (w) pickWeek(w);
    renderShell();
    startRealtime();
  } catch (e) { app.innerHTML = `<p class="boot" style="color:#b42318">ផ្ទុកទិន្នន័យមិនបាន៖ ${esc(errText(e))}<br><small>ពិនិត្យថាបាន Run schema.sql រួចហើយ</small></p>`; }
}

async function init() {
  if (!window.supabase || !CFG.SUPABASE_URL) { showLogin(); return; }
  let url = String(CFG.SUPABASE_URL).trim();
  if (!/^https?:\/\//i.test(url)) url = 'https://' + url;
  try { url = new URL(url).origin; } catch (e) { /* ទុកដដែល */ }
  S.sb = window.supabase.createClient(url, String(CFG.SUPABASE_ANON_KEY).trim());
  S.sb.auth.onAuthStateChange((ev) => { if (ev === 'SIGNED_OUT') { stopRealtime(); S.user = null; S.abs = {}; S.specials = {}; S.gs = null; S.sheets = []; showLogin(); } });
  const { data } = await S.sb.auth.getSession();
  if (data && data.session) { S.user = data.session.user; await boot(); } else showLogin();
}
window.__att = { S, buildWeekWorkbook, buildMonthWorkbook, monthWeeks, ensureAbs, parseStudents, weekTotals };
init();
})();