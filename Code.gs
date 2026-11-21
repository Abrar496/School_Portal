/**
 * Attendance - school server (free). It runs inside one Google Sheet that belongs to the school.
 *  - Class teachers' apps send one daily summary; the academic coordinator reads them.
 *  - Teachers and students log in (web page or app): notice board, homework, exam marks.
 * See SETUP.md. Student phone numbers are never stored: only a secret-keyed fingerprint of them.
 */
const TEACHERS = 'Teachers';   // A code | B section | C teacher | D login id | E mobile | F password fingerprint | G first password (you type it)
const DAILY = 'Daily';
const SETTINGS = 'Settings';   // B1: the coordinator's code
const TODAY = 'Today';
const STUDENTS = 'Students';   // section | roll | name | phone fingerprint | second fingerprint | captain
const NOTICES = 'Notices';
const CONFIG = 'ExamConfig';   // section | settings (JSON)
const MARKS = 'Marks';         // section | exam | subject | roll | parts (JSON) | total | updated
const ATTENDANCE = 'Attendance'; // section | roll | month (yyyy-MM) | marks (one letter per day: P present, A absent, L leave)
const CLASSINFO = 'ClassInfo';
const PROFILES = 'Student_Profiles';   // section | JSON {fine, weekend, holidays[]}: what a student's page needs to work out % and fine

function json_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
function str_(v) { return v instanceof Date ? Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd') : String(v == null ? '' : v); }
function num_(v) { return Math.max(0, Math.min(100000, Number(v) || 0)); }
function newCode_() { return Utilities.getUuid().replace(/-/g, '').slice(0, 8).toUpperCase(); }
function sheet_(n) { return SpreadsheetApp.getActive().getSheetByName(n); }

/** Run this once (Run > setup). It makes every tab and the secret key. */
function setup() {
  const ss = SpreadsheetApp.getActive();
  [TEACHERS, DAILY, SETTINGS, TODAY, STUDENTS, NOTICES, CONFIG, MARKS, ATTENDANCE, CLASSINFO, PROFILES].forEach(function (n) { if (!ss.getSheetByName(n)) ss.insertSheet(n); });
  const props = PropertiesService.getScriptProperties();
  if (!props.getProperty('SECRET')) props.setProperty('SECRET', Utilities.getUuid() + Utilities.getUuid());

  const t = sheet_(TEACHERS);
  if (t.getLastRow() === 0) {
    t.appendRow(['Code', 'Section', 'Teacher', 'Login id', 'Mobile', 'Password fingerprint (leave)', 'First password']);
    t.appendRow([newCode_(), 'XI-A', 'Teacher name', '01700000001', '01700000001', '', 'change-me-1']);   // example rows: edit them
    t.appendRow([newCode_(), 'XI-B', 'Teacher name', '01700000002', '01700000002', '', 'change-me-2']);
    t.setFrozenRows(1);
  }
  const d = sheet_(DAILY);
  if (d.getLastRow() === 0) { d.appendRow(['Date', 'Section', 'Teacher', 'Present', 'Absent', 'Leave', 'Total', 'Unmarked', 'Absent students', 'Updated', 'Absent data']); d.setFrozenRows(1); }
  d.getRange('A:A').setNumberFormat('@');
  const s = sheet_(SETTINGS);
  if (s.getLastRow() === 0) s.appendRow(['Coordinator code', newCode_()]);
  const st = sheet_(STUDENTS);
  if (st.getLastRow() === 0) { st.appendRow(['Section', 'Roll', 'Name', 'Phone fingerprint', 'Second fingerprint', 'Captain']); st.setFrozenRows(1); }
  const nt = sheet_(NOTICES);
  if (nt.getLastRow() === 0) { nt.appendRow(['Id', 'Created', 'Section', 'Role', 'Author', 'Author roll', 'Kind', 'Subject', 'Title', 'Details', 'Due', 'To rolls', 'Deleted']); nt.setFrozenRows(1); }
  nt.getRange('B:B').setNumberFormat('@'); nt.getRange('K:L').setNumberFormat('@');
  const cf = sheet_(CONFIG);
  if (cf.getLastRow() === 0) { cf.appendRow(['Section', 'Settings (JSON)']); cf.setFrozenRows(1); }
  const mk = sheet_(MARKS);
  if (mk.getLastRow() === 0) { mk.appendRow(['Section', 'Exam', 'Subject', 'Roll', 'Parts', 'Total', 'Updated']); mk.setFrozenRows(1); }

  const at = sheet_(ATTENDANCE);
  if (at.getLastRow() === 0) { at.appendRow(['Section', 'Roll', 'Month', 'Marks']); at.setFrozenRows(1); }
  at.getRange('C:D').setNumberFormat('@');
  const ci = sheet_(CLASSINFO);
  if (ci.getLastRow() === 0) { ci.appendRow(['Section', 'Info (JSON)']); ci.setFrozenRows(1); }

  const td = sheet_(TODAY);
  td.getRange('A1:H1').setValues([['Section', 'Teacher', 'Present', 'Absent', 'Leave', 'Total', 'Unmarked', 'Absent students']]);
  td.getRange('A2').setFormula('=IFERROR(QUERY(Daily!A2:J,"select B,C,D,E,F,G,H,I where A = \'"&TEXT(TODAY(),"yyyy-MM-dd")&"\' order by B",0),"No reports yet today")');
}

// ------------------------------------------------------------------ keys, fingerprints, tokens ----
function hmac_(text) {
  const secret = PropertiesService.getScriptProperties().getProperty('SECRET') || '';
  const sig = Utilities.computeHmacSha256Signature(text, secret);
  return sig.map(function (b) { return ('0' + (b & 0xff).toString(16)).slice(-2); }).join('');
}
function normPhone_(p) { const d = String(p == null ? '' : p).replace(/\D/g, ''); return d.length > 10 ? d.slice(-10) : d; }
function phoneFp_(section, roll, phone) { const n = normPhone_(phone); return n.length < 7 ? '' : hmac_('P|' + section + '|' + roll + '|' + n).slice(0, 24); }
function pwFp_(pw) { return hmac_('W|' + pw).slice(0, 32); }

/** Slows down guessing: at most 10 wrong tries for the same name or roll in 15 minutes. */
function tooMany_(key) { return Number(CacheService.getScriptCache().get('f_' + key) || 0) >= 10; }
function failed_(key) { const c = CacheService.getScriptCache(); c.put('f_' + key, String(Number(c.get('f_' + key) || 0) + 1), 900); }

// ------------------------------------------------------------------ teachers ----
function teacherRows_() { return sheet_(TEACHERS).getDataRange().getValues().map(function (r, i) { return { i: i + 1, r: r }; }).slice(1); }
function loginIdOf_(r) { return (str_(r[3]).trim() || normPhone_(r[4]) || str_(r[0]).trim()).toLowerCase(); }
function teacherInfo_(x) { return { section: str_(x.r[1]).trim(), teacher: str_(x.r[2]).trim(), loginId: loginIdOf_(x.r), mobile: str_(x.r[4]).trim(), row: x.i }; }
function teacherToken_(x) { return hmac_('T|' + loginIdOf_(x.r) + '|' + str_(x.r[5])).slice(0, 32); }

function teacherByCode_(code) {             // the old code, still used for the coordinator's daily report
  if (!code) return null;
  const rows = teacherRows_();
  for (let k = 0; k < rows.length; k++) if (str_(rows[k].r[0]).trim() === code) return teacherInfo_(rows[k]);
  return null;
}
function teacherLogin_(u, pw) {
  const id = String(u || '').trim().toLowerCase(), idPhone = normPhone_(u);
  if (!id || !pw) return { error: 'Enter your login id and password' };
  if (tooMany_('t' + id)) return { error: 'Too many wrong tries. Wait 15 minutes.' };
  const rows = teacherRows_();
  for (let k = 0; k < rows.length; k++) {
    const x = rows[k];
    if (loginIdOf_(x.r) !== id && !(idPhone.length >= 7 && normPhone_(x.r[4]) === idPhone)) continue;
    const fp = str_(x.r[5]), first = str_(x.r[6]);
    if (fp ? fp === pwFp_(pw) : (first && first === pw)) {
      if (!fp) { sheet_(TEACHERS).getRange(x.i, 6).setValue(pwFp_(pw)); sheet_(TEACHERS).getRange(x.i, 7).setValue(''); x.r[5] = pwFp_(pw); }
      return { teacher: teacherInfo_(x), token: teacherToken_(x) };
    }
    break;
  }
  failed_('t' + id);
  return { error: 'Wrong login id or password' };
}
function teacherByToken_(u, tk) {
  const id = String(u || '').trim().toLowerCase();
  if (!id || !tk) return null;
  const rows = teacherRows_();
  for (let k = 0; k < rows.length; k++) if (loginIdOf_(rows[k].r) === id && str_(rows[k].r[5]) && teacherToken_(rows[k]) === tk) return teacherInfo_(rows[k]);
  return null;
}

// ------------------------------------------------------------------ students ----
function studentRows_(section) {
  return sheet_(STUDENTS).getDataRange().getValues().map(function (r, i) { return { i: i + 1, r: r }; }).slice(1)
    .filter(function (x) { return str_(x.r[0]).trim() === section; });
}
function studentToken_(section, x) { return hmac_('S|' + section + '|' + x.r[1] + '|' + str_(x.r[3]) + '|' + str_(x.r[4])).slice(0, 32); }
function studentInfo_(section, x) { return { section: section, roll: Number(x.r[1]) || 0, name: str_(x.r[2]), captain: str_(x.r[5]).toUpperCase() === 'YES' }; }
function studentByToken_(section, roll, tk) {
  if (!section || !roll || !tk) return null;
  const rows = studentRows_(section);
  for (let k = 0; k < rows.length; k++) if (Number(rows[k].r[1]) === Number(roll) && studentToken_(section, rows[k]) === tk) return studentInfo_(section, rows[k]);
  return null;
}
function studentLogin_(section, roll, phone) {
  const key = 's' + section + '_' + roll;
  if (!section || !(Number(roll) > 0)) return { error: 'Choose your class and write your roll number' };
  if (tooMany_(key)) return { error: 'Too many wrong tries. Wait 15 minutes.' };
  const fp = phoneFp_(section, Number(roll), phone);
  const rows = studentRows_(section);
  for (let k = 0; k < rows.length; k++) {
    const x = rows[k];
    if (Number(x.r[1]) !== Number(roll)) continue;
    if (fp && (fp === str_(x.r[3]) || fp === str_(x.r[4]))) return { student: studentInfo_(section, x), token: studentToken_(section, x) };
    break;
  }
  failed_(key);
  return { error: 'Roll number or guardian mobile number is not correct' };
}

/** Who is asking? A teacher (u + tk) or a student (s + r + tk). Returns null when the login is not valid. */
function who_(b) {
  if (b.u) { const t = teacherByToken_(b.u, b.tk); return t ? { role: 'T', section: t.section, name: t.teacher, t: t } : null; }
  if (b.s) { const s = studentByToken_(String(b.s).trim(), b.r, b.tk); return s ? { role: 'S', section: s.section, roll: s.roll, name: s.name, captain: s.captain } : null; }
  return null;
}

// ------------------------------------------------------------------ notice board ----
function cleanRolls_(list) { const out = []; (list || []).slice(0, 200).forEach(function (x) { const n = Number(x); if (n > 0 && out.indexOf(n) < 0) out.push(n); }); return out; }

function noticeRows_(section) {
  const rows = sheet_(NOTICES).getDataRange().getValues().slice(1);
  const cutoff = Utilities.formatDate(new Date(Date.now() - 45 * 86400000), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  return rows.filter(function (r) { return str_(r[2]).trim() === section && String(r[12]).toUpperCase() !== 'YES' && str_(r[1]) >= cutoff; })
    .map(function (r) {
      return { id: str_(r[0]), created: str_(r[1]), role: str_(r[3]), author: str_(r[4]), authorRoll: Number(r[5]) || 0, kind: str_(r[6]),
        subject: str_(r[7]), title: str_(r[8]), body: str_(r[9]), due: str_(r[10]),
        to: str_(r[11]).split(',').map(function (x) { return Number(x); }).filter(function (x) { return x > 0; }) };
    }).sort(function (a, b) { return a.created < b.created ? 1 : -1; });
}
function addNotice_(w, b) {
  const kind = ['HOMEWORK', 'NOTICE', 'FOLLOWUP', 'MESSAGE'].indexOf(String(b.kind)) >= 0 ? String(b.kind) : 'NOTICE';
  const title = String(b.title || '').trim().slice(0, 140);
  if (!title) throw new Error('Write a title');
  const due = /^\d{4}-\d{2}-\d{2}$/.test(String(b.due || '')) ? String(b.due) : '';
  const role = w.role === 'T' ? 'T' : 'C';
  const to = role === 'T' ? cleanRolls_(b.to) : [];
  const id = Utilities.getUuid().slice(0, 8);
  const stamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd'T'HH:mm:ss");
  sheet_(NOTICES).appendRow([id, stamp, w.section, role, w.name, w.roll || 0, role === 'C' && kind === 'MESSAGE' ? 'NOTICE' : kind,
    String(b.subject || '').trim().slice(0, 40), title, String(b.body || '').slice(0, 2000), due, to.join(','), '']);
  return id;
}
function deleteNotice_(w, id) {
  const sh = sheet_(NOTICES), rows = sh.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (str_(rows[i][0]) !== id || str_(rows[i][2]).trim() !== w.section) continue;
    if (w.role === 'S' && Number(rows[i][5]) !== w.roll) throw new Error('You can delete only your own posts');
    sh.getRange(i + 1, 13).setValue('YES');
    return true;
  }
  return false;
}

// ------------------------------------------------------------------ exams and marks ----
const DEFAULT_SUBJECTS = ['Bangla', 'English', 'Physics', 'Chemistry', 'Biology', 'Higher Mathematics'];
const DEFAULT_EXAMS = [
  { id: 'W1', name: 'Weekly Test 1', kind: 'WEEKLY' }, { id: 'W2', name: 'Weekly Test 2', kind: 'WEEKLY' },
  { id: 'HY', name: 'Half Yearly', kind: 'TERM' }, { id: 'AN', name: 'Annual', kind: 'TERM' },
];
/** What a teacher has changed is kept; everything else gets these starting values: weekly 20 (pass 8), term CQ 70 + MCQ 30 (pass 33). */
function defaultSpec_(kind) {
  return kind === 'WEEKLY' ? { parts: [{ n: 'Marks', max: 20, pass: 0 }], pass: 8 } : { parts: [{ n: 'CQ', max: 70, pass: 0 }, { n: 'MCQ', max: 30, pass: 0 }], pass: 33 };
}
function expandConfig_(raw) {
  const subjects = (raw && raw.subjects && raw.subjects.length ? raw.subjects : DEFAULT_SUBJECTS).slice(0, 20).map(function (s) { return String(s).slice(0, 40); });
  const exams = (raw && raw.exams && raw.exams.length ? raw.exams : DEFAULT_EXAMS).slice(0, 12);
  const spec = (raw && raw.spec) || {};
  return {
    subjects: subjects,
    exams: exams.map(function (e) {
      const bySubject = {};
      subjects.forEach(function (s) {
        const o = spec[e.id] && spec[e.id][s] ? spec[e.id][s] : defaultSpec_(e.kind);
        const parts = (o.parts || []).slice(0, 4).map(function (p) { return { n: String(p.n || 'Marks').slice(0, 20), max: num_(p.max), pass: num_(p.pass) }; });
        bySubject[s] = { parts: parts, pass: num_(o.pass), total: parts.reduce(function (a, p) { return a + p.max; }, 0) };
      });
      return { id: String(e.id).slice(0, 12), name: String(e.name).slice(0, 40), kind: e.kind === 'WEEKLY' ? 'WEEKLY' : 'TERM', subjects: bySubject };
    }),
  };
}
function loadConfig_(section) {
  const rows = sheet_(CONFIG).getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) if (str_(rows[i][0]).trim() === section) { try { return expandConfig_(JSON.parse(str_(rows[i][1]))); } catch (e) { break; } }
  return expandConfig_(null);
}
function saveConfig_(section, cfg) {
  const slim = { subjects: cfg.subjects || [], exams: (cfg.exams || []).map(function (e) { return { id: e.id, name: e.name, kind: e.kind }; }), spec: {} };
  (cfg.exams || []).forEach(function (e) { slim.spec[e.id] = {}; Object.keys(e.subjects || {}).forEach(function (s) { slim.spec[e.id][s] = e.subjects[s]; }); });
  const expanded = expandConfig_(slim);                         // validates and limits what is stored
  const sh = sheet_(CONFIG), rows = sh.getDataRange().getValues();
  const row = [section, JSON.stringify(slim)];
  for (let i = 1; i < rows.length; i++) if (str_(rows[i][0]).trim() === section) { sh.getRange(i + 1, 1, 1, 2).setValues([row]); return expanded; }
  sh.appendRow(row);
  return expanded;
}
function loadInfo_(section) {
  const rows = sheet_(CLASSINFO).getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) if (str_(rows[i][0]).trim() === section) { try { return JSON.parse(str_(rows[i][1])); } catch (e) { break; } }
  return { fine: 0, weekend: 'FRI_SAT', holidays: [] };
}
function marksOf_(section, filter) {
  return sheet_(MARKS).getDataRange().getValues().slice(1)
    .filter(function (r) { return str_(r[0]).trim() === section && filter(r); })
    .map(function (r) { let parts = {}; try { parts = JSON.parse(str_(r[4]) || '{}'); } catch (e) {} return { exam: str_(r[1]), subject: str_(r[2]), roll: Number(r[3]) || 0, parts: parts, total: Number(r[5]) || 0 }; });
}

// ------------------------------------------------------------------ POST ----
function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
    const b = JSON.parse(e.postData.contents);
    switch (b.action) {
      case 'login': { const r = teacherLogin_(b.u, String(b.pw || '')); return r.error ? json_({ ok: false, error: r.error }) : json_({ ok: true, tk: r.token, teacher: r.teacher }); }
      case 'slogin': { const r = studentLogin_(String(b.s || '').trim(), b.r, b.phone); return r.error ? json_({ ok: false, error: r.error }) : json_({ ok: true, tk: r.token, student: r.student }); }
      case 'report': return report_(b);
      default: return secured_(b);
    }
  } catch (err) {
    return json_({ ok: false, error: String(err && err.message ? err.message : err) });
  } finally { lock.releaseLock(); }
}

function report_(b) {
  const t = teacherByCode_(String(b.code || '').trim());
  if (!t) return json_({ ok: false, error: 'Unknown code' });
  const date = String(b.date || '').slice(0, 10);
  if (!/^\d{4}-\d{2}-\d{2}$/.test(date)) return json_({ ok: false, error: 'Bad date' });
  const absent = (b.absentees || []).slice(0, 300).map(function (x) { return { r: Number(x.r) || 0, n: String(x.n || '').slice(0, 80) }; });
  const row = [date, t.section, String(b.teacher || t.teacher).slice(0, 60), num_(b.present), num_(b.absent), num_(b.leave), num_(b.total), num_(b.unmarked),
    absent.map(function (x) { return x.r + ' ' + x.n; }).join(', '), new Date().toISOString(), JSON.stringify(absent)];
  const sh = sheet_(DAILY), data = sh.getDataRange().getValues();
  let target = -1;
  for (let i = 1; i < data.length; i++) if (str_(data[i][0]) === date && str_(data[i][1]) === t.section) { target = i + 1; break; }
  if (target < 0) sh.appendRow(row); else sh.getRange(target, 1, 1, row.length).setValues([row]);
  return json_({ ok: true });
}

/** Every action that needs a valid login. */
function secured_(b) {
  const w = who_(b);
  if (!w) return json_({ ok: false, error: 'Please log in again' });
  const teacherOnly = function () { if (w.role !== 'T') throw new Error('Only a teacher can do this'); };

  switch (b.action) {
    case 'notice': {
      if (w.role === 'S' && !w.captain) throw new Error('Only class captains can post');
      return json_({ ok: true, id: addNotice_(w, b) });
    }
    case 'delete': { const ok = deleteNotice_(w, String(b.id || '')); return json_({ ok: ok, error: ok ? '' : 'Not found' }); }

    case 'profile': {
      teacherOnly();
      const sh = sheet_(TEACHERS), row = w.t.row, cur = sh.getRange(row, 1, 1, 7).getValues()[0];
      if (b.newLoginId != null && String(b.newLoginId).trim()) {
        const nid = String(b.newLoginId).trim().toLowerCase();
        if (teacherRows_().some(function (x) { return x.i !== row && loginIdOf_(x.r) === nid; })) throw new Error('That login id is already used');
        sh.getRange(row, 4).setValue(String(b.newLoginId).trim());
      }
      if (b.mobile != null) sh.getRange(row, 5).setValue(String(b.mobile).trim().slice(0, 20));
      if (b.newPassword) {
        if (String(b.newPassword).length < 6) throw new Error('The password needs at least 6 characters');
        if (str_(cur[5]) !== pwFp_(String(b.oldPassword || ''))) throw new Error('The old password is not correct');
        sh.getRange(row, 6).setValue(pwFp_(String(b.newPassword)));
      }
      const x = teacherRows_().filter(function (y) { return y.i === row; })[0];
      return json_({ ok: true, tk: teacherToken_(x), teacher: teacherInfo_(x) });
    }

    case 'roster': {            // the class teacher's student list; phone numbers become fingerprints, they are not stored
      teacherOnly();
      const list = (b.students || []).slice(0, 200), sh = sheet_(STUDENTS), rows = sh.getDataRange().getValues();
      const captain = {};
      for (let i = rows.length - 1; i >= 1; i--) if (str_(rows[i][0]).trim() === w.section) { captain[Number(rows[i][1])] = str_(rows[i][5]); sh.deleteRow(i + 1); }
      let n = 0;
      list.forEach(function (s) {
        const roll = Number(s.roll) || 0; if (roll <= 0) return;
        const cap = s.captain != null ? (s.captain ? 'YES' : '') : (captain[roll] || '');
        sh.appendRow([w.section, roll, String(s.name || '').slice(0, 80), phoneFp_(w.section, roll, s.phone), phoneFp_(w.section, roll, s.phone2), cap]); n++;
      });
      return json_({ ok: true, count: n });
    }
    case 'captains': {          // 4-6 class captains who may post homework and notices
      teacherOnly();
      const rolls = cleanRolls_(b.rolls).slice(0, 8), sh = sheet_(STUDENTS), rows = sh.getDataRange().getValues();
      for (let i = 1; i < rows.length; i++) if (str_(rows[i][0]).trim() === w.section) sh.getRange(i + 1, 6).setValue(rolls.indexOf(Number(rows[i][1])) >= 0 ? 'YES' : '');
      return json_({ ok: true });
    }

    case 'attendance': {        // the class teacher publishes attendance so that students can see their own record and fine
      teacherOnly();
      const info = b.info || {};
      const clean = {
        fine: num_(info.fine), weekend: ['FRI_SAT', 'SAT_SUN', 'SUN', 'FRI'].indexOf(String(info.weekend)) >= 0 ? String(info.weekend) : 'FRI_SAT',
        holidays: (info.holidays || []).slice(0, 400).filter(function (d) { return /^\d{4}-\d{2}-\d{2}$/.test(String(d)); }).map(String),
      };
      if (b.info) {
        const ci = sheet_(CLASSINFO), cr = ci.getDataRange().getValues(); let found = false;
        for (let i = 1; i < cr.length; i++) if (str_(cr[i][0]).trim() === w.section) { ci.getRange(i + 1, 1, 1, 2).setValues([[w.section, JSON.stringify(clean)]]); found = true; break; }
        if (!found) ci.appendRow([w.section, JSON.stringify(clean)]);
      }
      const sh = sheet_(ATTENDANCE), data = sh.getDataRange().getValues(), index = {};
      sh.getRange('C:D').setNumberFormat('@');
      for (let i = 1; i < data.length; i++) if (str_(data[i][0]).trim() === w.section) index[Number(data[i][1]) + '|' + str_(data[i][2]).slice(0, 7)] = i + 1;
      let saved = 0;
      (b.rows || []).slice(0, 1500).forEach(function (r) {
        const roll = Number(r.roll) || 0, month = String(r.month || '').slice(0, 7);
        if (roll <= 0 || !/^\d{4}-\d{2}$/.test(month)) return;
        const marks = String(r.marks || '').toUpperCase().replace(/[^PAL ]/g, ' ').slice(0, 31);
        const key = roll + '|' + month;
        if (index[key]) sh.getRange(index[key], 1, 1, 4).setValues([[w.section, roll, month, marks]]); else sh.appendRow([w.section, roll, month, marks]);
        saved++;
      });
      return json_({ ok: true, saved: saved });
    }
    case 'attendanceMine': {
      if (w.role !== 'S') throw new Error('Open this as a student');
      return json_({ ok: true, info: loadInfo_(w.section), months: sheet_(ATTENDANCE).getDataRange().getValues().slice(1)
        .filter(function (r) { return str_(r[0]).trim() === w.section && Number(r[1]) === w.roll; })
        .map(function (r) { return { month: str_(r[2]).slice(0, 7), marks: str_(r[3]) }; })
        .sort(function (a, c) { return a.month < c.month ? -1 : 1; }) });
    }
    case 'saveConfig': { teacherOnly(); return json_({ ok: true, config: saveConfig_(w.section, b.config || {}) }); }
    case 'marksSave': {
      teacherOnly();
      const cfg = loadConfig_(w.section), exam = cfg.exams.filter(function (e) { return e.id === String(b.exam); })[0];
      const spec = exam && exam.subjects[String(b.subject)];
      if (!spec) throw new Error('Unknown exam or subject');
      const sh = sheet_(MARKS), data = sh.getDataRange().getValues(), index = {};
      for (let i = 1; i < data.length; i++) if (str_(data[i][0]).trim() === w.section && str_(data[i][1]) === exam.id && str_(data[i][2]) === String(b.subject)) index[Number(data[i][3])] = i + 1;
      const stamp = new Date().toISOString(); let saved = 0;
      (b.rows || []).slice(0, 200).forEach(function (r) {
        const roll = Number(r.roll) || 0; if (roll <= 0) return;
        const parts = {}; let total = 0, any = false;
        spec.parts.forEach(function (p) {
          const v = r.parts ? r.parts[p.n] : '';
          if (v === 'AB') { parts[p.n] = 'AB'; any = true; }
          else if (v !== '' && v != null && !isNaN(Number(v))) {
            const x = Number(v); if (x < 0 || x > p.max) throw new Error('Roll ' + roll + ': ' + p.n + ' must be between 0 and ' + p.max);
            parts[p.n] = x; total += x; any = true;
          }
        });
        const row = [w.section, exam.id, String(b.subject), roll, JSON.stringify(parts), total, stamp];
        if (index[roll]) { if (any) sh.getRange(index[roll], 1, 1, 7).setValues([row]); else sh.getRange(index[roll], 5, 1, 3).setValues([['{}', 0, stamp]]); }
        else if (any) sh.appendRow(row);
        saved++;
      });
      return json_({ ok: true, saved: saved });
    }
    case 'marksClass': {
      teacherOnly();
      const roster = studentRows_(w.section).map(function (x) { return { roll: Number(x.r[1]), name: str_(x.r[2]), captain: str_(x.r[5]).toUpperCase() === 'YES' }; })
        .sort(function (a, c) { return a.roll - c.roll; });
      const marks = marksOf_(w.section, function (r) { return str_(r[1]) === String(b.exam) && str_(r[2]) === String(b.subject); });
      return json_({ ok: true, roster: roster, marks: marks });
    }
    case 'marksBulkPublish': {
      teacherOnly();
      const sh = sheet_(MARKS);
      const data = sh.getDataRange().getValues();
      const index = {};
      for (let i = 1; i < data.length; i++) {
        if (str_(data[i][0]).trim() === w.section) {
          index[str_(data[i][1]) + '|' + str_(data[i][2]) + '|' + Number(data[i][3])] = i + 1;
        }
      }
      const stamp = new Date().toISOString();
      let saved = 0;
      (b.sheets || []).slice(0, 60).forEach(function (sheetObj) {
        const examName = String(sheetObj.exam || 'Exam').trim().slice(0, 60);
        const subjName = String(sheetObj.subject || 'General').trim().slice(0, 60);
        const fullMarks = Number(sheetObj.full) || 0;
        const passMarks = Number(sheetObj.pass) || 0;
        (sheetObj.rows || []).slice(0, 250).forEach(function (r) {
          const roll = Number(r.roll) || 0;
          if (roll <= 0) return;
          const rawM = String(r.marks == null ? '' : r.marks).trim();
          if (!rawM) return;
          const isAb = /^a(b|bsent)?$/i.test(rawM);
          const numVal = isAb ? 'AB' : (isNaN(Number(rawM)) ? rawM : Number(rawM));
          const totalNum = typeof numVal === 'number' ? numVal : 0;
          const parts = { Marks: numVal };
          if (fullMarks > 0) parts.Full = fullMarks;
          if (passMarks > 0) parts.Pass = passMarks;
          const row = [w.section, examName, subjName, roll, JSON.stringify(parts), totalNum, stamp];
          const k = examName + '|' + subjName + '|' + roll;
          if (index[k]) {
            sh.getRange(index[k], 1, 1, 7).setValues([row]);
          } else {
            sh.appendRow(row);
            index[k] = sh.getLastRow();
          }
          saved++;
        });
      });
      return json_({ ok: true, saved: saved });
    }
    case 'profilesSave': {
      teacherOnly();
      let sh = sheet_(PROFILES);
      if (!sh) {
        sh = SpreadsheetApp.getActive().insertSheet(PROFILES);
        sh.appendRow(['Section', 'Roll', 'ProfileJSON', 'Updated']);
      }
      const data = sh.getDataRange().getValues();
      const index = {};
      for (let i = 1; i < data.length; i++) {
        if (str_(data[i][0]).trim() === w.section) index[Number(data[i][1])] = i + 1;
      }
      const stamp = new Date().toISOString();
      let saved = 0;
      (b.profiles || []).slice(0, 250).forEach(function (prof) {
        const roll = Number(prof.roll) || 0;
        if (roll <= 0) return;
        const row = [w.section, roll, JSON.stringify(prof), stamp];
        if (index[roll]) {
          sh.getRange(index[roll], 1, 1, 4).setValues([row]);
        } else {
          sh.appendRow(row);
          index[roll] = sh.getLastRow();
        }
        saved++;
      });
      return json_({ ok: true, saved: saved });
    }
    case 'profileMine': {
      if (w.role !== 'S') throw new Error('Open this as a student');
      const sh = sheet_(PROFILES);
      let prof = null;
      if (sh) {
        const data = sh.getDataRange().getValues();
        for (let i = 1; i < data.length; i++) {
          if (str_(data[i][0]).trim() === w.section && Number(data[i][1]) === w.roll) {
            try { prof = JSON.parse(str_(data[i][2]) || 'null'); } catch (e) {}
            break;
          }
        }
      }
      return json_({ ok: true, profile: prof });
    }
    case 'marksMine': {
      if (w.role !== 'S') throw new Error('Open this as a student');
      return json_({ ok: true, config: loadConfig_(w.section), marks: marksOf_(w.section, function (r) { return Number(r[3]) === w.roll; }) });
    }
  }
  throw new Error('Unknown action');
}

// ------------------------------------------------------------------ GET ----
function doGet(e) {
  const p = (e && e.parameter) || {};
  try {
    if (p.action === 'ping') {
      const t = teacherByCode_(String(p.code || '').trim());
      return t ? json_({ ok: true, section: t.section, teacher: t.teacher }) : json_({ ok: false, error: 'Unknown code' });
    }
    if (p.action === 'sections') {                       // the class names, for the student login list
      const seen = {}, out = [];
      teacherRows_().forEach(function (x) { const s = str_(x.r[1]).trim(); if (s && !seen[s]) { seen[s] = 1; out.push(s); } });
      return json_({ ok: true, sections: out });
    }
    if (p.action === 'range') {
      if (!coordinatorOk_(String(p.key || '').trim())) return json_({ ok: false, error: 'Wrong coordinator code' });
      const from = String(p.from || ''), to = String(p.to || '');
      const expected = teacherRows_().filter(function (x) { return str_(x.r[1]).trim(); }).map(function (x) { return { section: str_(x.r[1]).trim(), teacher: str_(x.r[2]).trim() }; });
      const rows = sheet_(DAILY).getDataRange().getValues().slice(1)
        .filter(function (r) { const d = str_(r[0]); return d >= from && d <= to; })
        .map(function (r) {
          let absentees = []; try { absentees = JSON.parse(str_(r[10]) || '[]'); } catch (x) { absentees = []; }
          return { date: str_(r[0]), section: str_(r[1]), teacher: str_(r[2]), present: num_(r[3]), absent: num_(r[4]), leave: num_(r[5]),
            total: num_(r[6]), unmarked: num_(r[7]), updated: str_(r[9]), absentees: absentees };
        });
      return json_({ ok: true, expected: expected, rows: rows });
    }
    const w = who_(p);                                    // the rest needs a login (teacher: u + tk, student: s + r + tk)
    if (!w) return json_({ ok: false, error: 'Please log in again' });
    if (p.action === 'feed') {
      const items = noticeRows_(w.section).filter(function (n) { return w.role === 'T' || n.to.length === 0 || n.to.indexOf(w.roll) >= 0; });
      return json_({ ok: true, section: w.section, mode: w.role === 'T' ? 'T' : (w.captain ? 'C' : 'S'), roll: w.roll || 0, name: w.name, items: items });
    }
    if (p.action === 'config') return json_({ ok: true, config: loadConfig_(w.section) });
    if (p.action === 'marksMine' || p.action === 'attendanceMine' || p.action === 'profileMine') return secured_({ action: p.action, u: p.u, s: p.s, r: p.r, tk: p.tk });
    return json_({ ok: false, error: 'Unknown action' });
  } catch (err) {
    return json_({ ok: false, error: String(err && err.message ? err.message : err) });
  }
}

function coordinatorOk_(key) {
  if (!key) return false;
  return str_(sheet_(SETTINGS).getRange('B1').getValue()).trim() === key;
}
