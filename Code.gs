/**
 * خادم اختبار نظم المعلومات الجغرافية (Google Apps Script)
 * يستقبل نتائج الطلاب ويحفظها في ورقة "Results" ويعطيها للوحة المعلم.
 *
 * 1) غيّر TEACHER_KEY إلى كلمة سر من اختيارك (المعلم فقط يعرفها).
 * 2) إذا غيّرت الأسئلة، غيّر ANSWERS بنفس ترتيب ملف questions.js
 *    (0 = أ، 1 = ب، 2 = ج، 3 = د).
 */
const TEACHER_KEY = 'غير-كلمة-السر-هذه';
const ANSWERS = [1, 0, 2, 0, 3, 1, 2, 0, 0, 1, 2, 3, 0, 1];
const PER_Q = 15;
const SHEET = 'Results';
const HEAD = ['id', 'الاسم', 'اسم الأب', 'القبيلة', 'الشعبة', 'الدرجة', 'من', 'الوقت (ثانية)', 'التاريخ', 'يدوي', 'answers', 'times'];

function sheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET);
  if (!sh) {
    sh = ss.insertSheet(SHEET);
    sh.appendRow(HEAD);
    sh.setFrozenRows(1);
    sh.setRightToLeft(true);
  }
  return sh;
}

function out_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function clean_(s, max) {
  s = String(s == null ? '' : s).replace(/\s+/g, ' ').trim();
  return s.slice(0, max || 30);
}

function rows_() {
  const sh = sheet_();
  const n = sh.getLastRow();
  if (n < 2) return [];
  return sh.getRange(2, 1, n - 1, HEAD.length).getValues().map(function (r) {
    return {
      id: String(r[0]), first: r[1], father: r[2], tribe: r[3], grade: r[4],
      score: Number(r[5]), total: Number(r[6]), timeSec: Number(r[7]),
      at: r[8] instanceof Date ? r[8].toISOString() : String(r[8]),
      manual: r[9] === 'نعم',
      answers: r[10] ? JSON.parse(r[10]) : null
    };
  });
}

function ranked_() {
  return rows_().sort(function (a, b) {
    if (a.score !== b.score) return b.score - a.score;
    if (a.timeSec !== b.timeSec) return a.timeSec - b.timeSec;
    return new Date(a.at) - new Date(b.at);
  });
}

function doGet(e) {
  const p = e.parameter || {};
  if (p.action === 'results') {
    if (p.key !== TEACHER_KEY) return out_({ ok: false, error: 'كلمة سر المعلم غير صحيحة' });
    return out_({ ok: true, total: ANSWERS.length, answers: ANSWERS, results: ranked_() });
  }
  return out_({ ok: true, total: ANSWERS.length });
}

function doPost(e) {
  let b;
  try { b = JSON.parse(e.postData.contents); } catch (err) { return out_({ ok: false, error: 'بيانات غير صالحة' }); }
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const action = b.action || 'submit';
    if (action === 'submit') return out_(submit_(b));
    if (b.key !== TEACHER_KEY) return out_({ ok: false, error: 'كلمة سر المعلم غير صحيحة' });
    if (action === 'add') return out_(add_(b));
    if (action === 'delete') return out_(del_(b.id));
    if (action === 'clear') { const sh = sheet_(); if (sh.getLastRow() > 1) sh.deleteRows(2, sh.getLastRow() - 1); return out_({ ok: true }); }
    return out_({ ok: false, error: 'طلب غير معروف' });
  } finally {
    lock.releaseLock();
  }
}

function submit_(b) {
  const first = clean_(b.first), father = clean_(b.father), tribe = clean_(b.tribe), grade = clean_(b.grade);
  if (!first || !father || !tribe) return { ok: false, error: 'الاسم واسم الأب والقبيلة مطلوبة' };
  const ans = b.answers, tms = b.times;
  if (!Array.isArray(ans) || ans.length !== ANSWERS.length) return { ok: false, error: 'عدد الإجابات غير صحيح' };
  const id = clean_(b.id, 60) || Utilities.getUuid();
  if (rows_().some(function (r) { return r.id === id; })) return { ok: true, duplicate: true };
  let score = 0, time = 0;
  const times = [];
  for (let i = 0; i < ANSWERS.length; i++) {
    const a = Number(ans[i]);
    if (a === ANSWERS[i]) score++;
    let t = Array.isArray(tms) ? Number(tms[i]) : PER_Q;
    if (!(t >= 0)) t = 0;
    if (t > PER_Q || a < 0) t = PER_Q;
    t = Math.floor(t * 10) / 10;
    times.push(t);
    time += t;
  }
  time = Math.round(time * 10) / 10;
  sheet_().appendRow([id, first, father, tribe, grade, score, ANSWERS.length, time, new Date(), '',
    JSON.stringify(ans.map(Number)), JSON.stringify(times)]);
  return { ok: true, score: score, total: ANSWERS.length, timeSec: time };
}

function add_(b) {
  const first = clean_(b.first), father = clean_(b.father), tribe = clean_(b.tribe), grade = clean_(b.grade);
  if (!first || !father || !tribe) return { ok: false, error: 'الاسم واسم الأب والقبيلة مطلوبة' };
  const score = Number(b.score), time = Number(b.timeSec || 0);
  if (!(score >= 0 && score <= ANSWERS.length && Math.floor(score) === score)) return { ok: false, error: 'الدرجة يجب أن تكون بين 0 و ' + ANSWERS.length };
  if (!(time >= 0 && time <= PER_Q * ANSWERS.length)) return { ok: false, error: 'الوقت يجب أن يكون بين 0 و ' + (PER_Q * ANSWERS.length) + ' ثانية' };
  sheet_().appendRow([Utilities.getUuid(), first, father, tribe, grade, score, ANSWERS.length, time, new Date(), 'نعم', '', '']);
  return { ok: true };
}

function del_(id) {
  const sh = sheet_();
  const n = sh.getLastRow();
  if (n < 2) return { ok: true };
  const ids = sh.getRange(2, 1, n - 1, 1).getValues();
  for (let i = 0; i < ids.length; i++) {
    if (String(ids[i][0]) === String(id)) { sh.deleteRow(i + 2); break; }
  }
  return { ok: true };
}
