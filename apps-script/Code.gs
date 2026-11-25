/**
 * Diamond Checklist — หลังบ้าน (Google Apps Script)
 * วางไฟล์นี้ในไฟล์ Google Sheet "แผน 1 ปี 1 ล้าน (Sep 2026-Aug 2027)" → ส่วนขยาย → Apps Script
 * แล้วรันฟังก์ชัน setup() หนึ่งครั้ง (ดูคู่มือใน README.md)
 */

// ไอดีไฟล์ Google Sheets ทั้ง 3 ไฟล์ในโฟลเดอร์ Diamond Amway Project
const IDS = {
  plan: '1aXuHG6nWTuwF3cXc899rLNiBf28u36juraZndcHM9Zo',      // แผน 1 ปี 1 ล้าน
  prospects: '11TcE5BkMsXeG7YiwcnLc3LtqHS4RtdRehr-6Wxfzkmk', // แบบลิสรายชื่อผู้มุ่งหวัง
  calendar: '1bjq0d_evcGaKGLORCcOzsbS6pjNwNOXejXqop7gSCTc'   // ตารางงานเกรียง2026
};

const SHEET_CHECKLIST = 'Checklist รายวัน';
const SHEET_LOG = 'บันทึกเช็กลิสต์';
const SHEET_PROSPECTS = 'แบบฟอร์มใบรายชื่อผู้มุ่งหวัง1';

const HABITS = [
  ['link', 'Link a day'],
  ['body', 'ออกกำลังกาย/อ่าน'],
  ['coffee', 'แก้วกาแฟ'],
  ['meditate', 'สมาธิ'],
  ['clinic', 'คลิปคลินิก'],
  ['plan', 'ฝึกเล่าแผน'],
  ['cyber', 'งานไซเบอร์'],
  ['follow', 'ติดตามลูกค้า'],
  ['online', 'งานช่วงค่ำ/ออนไลน์'],
  ['review', 'สรุปวัน']
];
const NUMS = [
  ['abo', 'ABO Active'],
  ['newpp', 'New PP'],
  ['service', 'Service M'],
  ['selfdev', 'พัฒนาตัวเอง (นาที)'],
  ['cases', 'จำนวนเคส'],
  ['bv', 'BV']
];

const MONTH_FULL = ['มกราคม', 'กุมภาพันธ์', 'มีนาคม', 'เมษายน', 'พฤษภาคม', 'มิถุนายน',
  'กรกฎาคม', 'สิงหาคม', 'กันยายน', 'ตุลาคม', 'พฤศจิกายน', 'ธันวาคม'];
const MONTH_SHORT = ['ม.ค', 'ก.พ', 'มี.ค', 'เม.ย', 'พ.ค', 'มิ.ย', 'ก.ค', 'ส.ค', 'ก.ย', 'ต.ค', 'พ.ย', 'ธ.ค'];
const DOW_SHORT = ['อา', 'จ', 'อ', 'พ', 'พฤ', 'ศ', 'ส'];
const CAL_FIRST_ROW = 2;   // แถว 8:00 น
const CAL_LAST_ROW = 26;   // แถว 20:00 น

// ---------- ติดตั้งครั้งแรก ----------
function setup() {
  const props = PropertiesService.getScriptProperties();
  let key = props.getProperty('APP_KEY');
  if (!key) {
    key = Utilities.getUuid().replace(/-/g, '') + Utilities.getUuid().replace(/-/g, '').slice(0, 8);
    props.setProperty('APP_KEY', key);
  }
  logSheet_();
  // แตะทั้ง 3 ไฟล์เพื่อให้ Google ขอสิทธิ์ครั้งเดียว
  SpreadsheetApp.openById(IDS.prospects).getName();
  SpreadsheetApp.openById(IDS.calendar).getName();
  Logger.log('รหัสลับของแอป (คัดลอกไปใส่ในหน้าแอป): ' + key);
}

/** ใช้เมื่ออยากเปลี่ยนรหัสลับ เช่น มือถือหาย */
function resetKey() {
  PropertiesService.getScriptProperties().deleteProperty('APP_KEY');
  setup();
}

// ---------- จุดรับคำสั่งจากหน้าแอป ----------
function doGet(e) {
  if (e && e.parameter && e.parameter.payload) {
    return json_(handle_(JSON.parse(e.parameter.payload)));
  }
  return ContentService.createTextOutput('Diamond Checklist API ทำงานอยู่');
}

function doPost(e) {
  let req;
  try {
    req = JSON.parse(e.postData.contents);
  } catch (err) {
    return json_({ ok: false, error: 'ข้อมูลที่ส่งมาไม่ถูกต้อง' });
  }
  return json_(handle_(req));
}

function handle_(req) {
  const key = PropertiesService.getScriptProperties().getProperty('APP_KEY');
  if (!key || !req || req.key !== key) return { ok: false, error: 'รหัสลับไม่ถูกต้อง' };
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    switch (req.action) {
      case 'load': return Object.assign({ ok: true }, load_(req.date));
      case 'saveDay': saveDay_(req); return { ok: true };
      case 'addProspect': return Object.assign({ ok: true }, addProspect_(req.p || {}));
      case 'addAppt': addAppt_(req.date, req.time, req.text); return { ok: true };
      default: return { ok: false, error: 'ไม่รู้จักคำสั่ง ' + req.action };
    }
  } catch (err) {
    return { ok: false, error: String(err && err.message || err) };
  } finally {
    lock.releaseLock();
  }
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

// ---------- วันที่ ----------
function parseDate_(s) {
  const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(String(s || ''));
  if (!m) throw new Error('รูปแบบวันที่ต้องเป็น YYYY-MM-DD');
  return { y: +m[1], m: +m[2], d: +m[3], s: m[0] };
}
function addDays_(p, n) {
  const dt = new Date(Date.UTC(p.y, p.m - 1, p.d + n));
  return { y: dt.getUTCFullYear(), m: dt.getUTCMonth() + 1, d: dt.getUTCDate(),
    s: dt.toISOString().slice(0, 10), dow: dt.getUTCDay() };
}
function dow_(p) { return new Date(Date.UTC(p.y, p.m - 1, p.d)).getUTCDay(); }
function daysInMonth_(y, m) { return new Date(Date.UTC(y, m, 0)).getUTCDate(); }
function monthLabel_(p) { return MONTH_FULL[p.m - 1] + ' ' + (p.y + 543); }
function num_(v) { return (v === '' || v === null || v === undefined || isNaN(Number(v))) ? '' : Number(v); }

// ---------- ชีตบันทึกเช็กลิสต์ ----------
function logSheet_() {
  const ss = SpreadsheetApp.openById(IDS.plan);
  let sh = ss.getSheetByName(SHEET_LOG);
  if (!sh) {
    sh = ss.insertSheet(SHEET_LOG);
    const head = ['วันที่'].concat(HABITS.map(h => h[1]), NUMS.map(n => n[1]), ['หมายเหตุ', 'อัปเดตล่าสุด']);
    sh.getRange(1, 1, 1, head.length).setValues([head]).setFontWeight('bold').setBackground('#e8eef7');
    sh.setFrozenRows(1);
    sh.getRange('A:A').setNumberFormat('@');
  }
  return sh;
}

function findLogRow_(sh, dateStr) {
  const last = sh.getLastRow();
  if (last < 2) return -1;
  const vals = sh.getRange(2, 1, last - 1, 1).getDisplayValues();
  for (let i = 0; i < vals.length; i++) if (vals[i][0] === dateStr) return i + 2;
  return -1;
}

// ---------- โหลดข้อมูลวันนี้ ----------
function load_(dateStr) {
  const p = parseDate_(dateStr);
  const sh = logSheet_();
  const row = findLogRow_(sh, p.s);
  const day = { habits: {}, nums: {}, note: '' };
  if (row > 0) {
    const width = 1 + HABITS.length + NUMS.length + 1;
    const v = sh.getRange(row, 1, 1, width).getValues()[0];
    HABITS.forEach((h, i) => { day.habits[h[0]] = v[1 + i] === true; });
    NUMS.forEach((n, i) => { day.nums[n[0]] = v[1 + HABITS.length + i]; });
    day.note = v[width - 1] || '';
  }
  return { day: day, appts: upcomingAppts_(p, 7), prospects: topProspects_(20), streak: streak_(sh, p) };
}

/** จำนวนวันติดกันที่ติ๊กได้ตั้งแต่ 80% ขึ้นไป นับย้อนจากเมื่อวาน */
function streak_(sh, p) {
  const last = sh.getLastRow();
  if (last < 2) return 0;
  const vals = sh.getRange(2, 1, last - 1, 1 + HABITS.length).getValues();
  const map = {};
  vals.forEach(r => {
    const done = r.slice(1).filter(x => x === true).length;
    map[String(r[0])] = done / HABITS.length >= 0.8;
  });
  let n = 0;
  for (let i = 1; i < 400; i++) {
    const q = addDays_(p, -i);
    if (map[q.s]) n++; else break;
  }
  return n;
}

// ---------- บันทึกวัน ----------
function saveDay_(req) {
  const p = parseDate_(req.date);
  const habits = req.habits || {};
  const nums = req.nums || {};
  const note = String(req.note || '').slice(0, 500);

  const sh = logSheet_();
  let row = findLogRow_(sh, p.s);
  if (row < 0) row = sh.getLastRow() + 1;
  const rowVals = [p.s].concat(
    HABITS.map(h => habits[h[0]] === true),
    NUMS.map(n => num_(nums[n[0]])),
    [note, new Date()]
  );
  sh.getRange(row, 1).setNumberFormat('@');
  sh.getRange(row, 1, 1, rowVals.length).setValues([rowVals]);
  sh.getRange(row, 2, 1, HABITS.length).insertCheckboxes();

  writeMonthlyChecklist_(p, nums, note);
}

/** เขียนตัวเลขลงชีต "Checklist รายวัน" (แถว = วันที่ของเดือน) โดยไม่แตะสูตร */
function writeMonthlyChecklist_(p, nums, note) {
  const ss = SpreadsheetApp.openById(IDS.plan);
  const sh = ss.getSheetByName(SHEET_CHECKLIST);
  if (!sh) throw new Error('ไม่พบชีต ' + SHEET_CHECKLIST);
  const label = monthLabel_(p);
  const cur = String(sh.getRange('D3').getDisplayValue()).trim();
  if (cur !== label) {
    const data = sh.getRange('B8:G38').getValues();
    const hasData = data.some(r => r.some(c => c !== '' && c !== null));
    // เก็บเดือนเก่าไว้เป็นชีตสำเนา ก่อนล้างเทมเพลตสำหรับเดือนใหม่
    if (hasData && cur && cur.indexOf('เช่น') !== 0) {
      let name = 'Checklist ' + cur;
      let k = 2;
      while (ss.getSheetByName(name)) name = 'Checklist ' + cur + ' (' + (k++) + ')';
      sh.copyTo(ss).setName(name);
    }
    sh.getRange('B8:G38').clearContent();
    sh.getRange('J8:J38').clearContent();
    sh.getRange('D3').setValue(label);
  }
  const r = 7 + p.d;
  sh.getRange(r, 2, 1, 6).setValues([NUMS.map(n => num_(nums[n[0]]))]);
  sh.getRange(r, 10).setValue(note);
}

// ---------- รายชื่อผู้มุ่งหวัง ----------
function prospectSheet_() {
  const sh = SpreadsheetApp.openById(IDS.prospects).getSheetByName(SHEET_PROSPECTS);
  if (!sh) throw new Error('ไม่พบชีต ' + SHEET_PROSPECTS);
  return sh;
}

function clamp_(v, max) {
  const n = num_(v);
  if (n === '') return '';
  return Math.max(0, Math.min(max, n));
}

function addProspect_(p) {
  const name = String(p.name || '').trim();
  if (!name) throw new Error('กรุณาใส่ชื่อ');
  const sh = prospectSheet_();
  const last = Math.max(sh.getLastRow(), 5);
  const names = sh.getRange(5, 2, last - 4, 1).getValues();
  let row = -1;
  for (let i = 0; i < names.length; i++) {
    if (String(names[i][0]).trim() === '') { row = 5 + i; break; }
  }
  if (row < 0) {
    row = last + 1;
    sh.getRange(row, 1).setValue(row - 4);
    sh.getRange(row, 8).setFormula('=IF(COUNTA(D' + row + ':G' + row + ')>0, SUM(D' + row + ':G' + row + '), "")');
  }
  sh.getRange(row, 2, 1, 6).setValues([[name, String(p.topic || ''),
    clamp_(p.s1, 10), clamp_(p.s2, 5), clamp_(p.s3, 5), clamp_(p.s4, 5)]]);
  sh.getRange(row, 9, 1, 4).setValues([[String(p.job || ''), num_(p.age), String(p.district || ''), String(p.province || '')]]);
  SpreadsheetApp.flush();
  return { prospects: topProspects_(20) };
}

function topProspects_(n) {
  const sh = prospectSheet_();
  const last = sh.getLastRow();
  if (last < 5) return [];
  const v = sh.getRange(5, 2, last - 4, 7).getValues(); // B..H
  return v.filter(r => String(r[0]).trim() !== '')
    .map(r => ({ name: String(r[0]), topic: String(r[1]), score: Number(r[6]) || 0 }))
    .sort((a, b) => b.score - a.score)
    .slice(0, n);
}

// ---------- ตารางงาน ----------
function calSheet_(y, m, create) {
  const ss = SpreadsheetApp.openById(IDS.calendar);
  const be = y + 543;
  const name = be === 2569 ? MONTH_SHORT[m - 1] : MONTH_SHORT[m - 1] + ' ' + String(be).slice(2);
  let sh = ss.getSheetByName(name);
  if (!sh && create) {
    // ปีใหม่: คัดลอกโครงจากเดือน ม.ค (31 วัน) แล้วเขียนหัววันใหม่
    sh = ss.getSheetByName(MONTH_SHORT[0]).copyTo(ss).setName(name);
    sh.getRange(CAL_FIRST_ROW, 2, CAL_LAST_ROW - CAL_FIRST_ROW + 1, 31).clearContent();
    const days = daysInMonth_(y, m);
    const head = [];
    for (let d = 1; d <= 31; d++) {
      head.push(d <= days
        ? DOW_SHORT[dow_({ y: y, m: m, d: d })] + ' ที่ ' + d + ' ' + MONTH_SHORT[m - 1] + ' ' + String(be).slice(2)
        : '');
    }
    sh.getRange(1, 2, 1, 31).setValues([head]);
  }
  return sh;
}

function addAppt_(dateStr, time, text) {
  const p = parseDate_(dateStr);
  const t = /^(\d{1,2}):(\d{2})$/.exec(String(time || ''));
  if (!t) throw new Error('เวลาต้องเป็นรูปแบบ HH:MM');
  const msg = String(text || '').trim();
  if (!msg) throw new Error('กรุณาใส่รายละเอียดนัดหมาย');
  const h = +t[1], mi = +t[2];
  let row = CAL_FIRST_ROW + (h - 8) * 2 + (mi >= 30 ? 1 : 0);
  row = Math.max(CAL_FIRST_ROW, Math.min(CAL_LAST_ROW, row));
  const sh = calSheet_(p.y, p.m, true);
  const cell = sh.getRange(row, 1 + p.d);
  const line = ('0' + h).slice(-2) + ':' + ('0' + mi).slice(-2) + ' ' + msg;
  const old = String(cell.getValue() || '').trim();
  cell.setValue(old ? old + '\n' + line : line);
}

function upcomingAppts_(p, days) {
  const out = [];
  const cache = {};
  for (let i = 0; i < days; i++) {
    const q = addDays_(p, i);
    const k = q.y + '-' + q.m;
    if (!(k in cache)) {
      const sh = calSheet_(q.y, q.m, false);
      cache[k] = sh ? sh.getRange(CAL_FIRST_ROW, 1, CAL_LAST_ROW - CAL_FIRST_ROW + 1, 32).getDisplayValues() : null;
    }
    const grid = cache[k];
    if (!grid) continue;
    grid.forEach(r => {
      const v = String(r[q.d] || '').trim();
      if (v) out.push({ date: q.s, slot: String(r[0]).replace(' น', ''), text: v });
    });
  }
  return out;
}
