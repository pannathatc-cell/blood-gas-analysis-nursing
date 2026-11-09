/**
 * ABG NSSKRU — บันทึกผลการเล่นและการเข้าใช้งานลง Google Sheet
 * คณะพยาบาลศาสตร์ มหาวิทยาลัยราชภัฏศรีสะเกษ
 *
 * วิธีใช้ (ทำครั้งเดียว)
 * 1) เปิด Google Sheet → ส่วนขยาย → Apps Script → วางโค้ดทั้งไฟล์นี้แทนของเดิม → บันทึก
 * 2) เลือกฟังก์ชัน setup แล้วกด "เรียกใช้" → อนุญาตสิทธิ์ (สร้างแท็บ Results และ Visits ให้)
 * 3) ทำให้ใช้งานได้ → การทำให้ใช้งานได้รายการใหม่ → ประเภท: เว็บแอป
 *      เรียกใช้ในฐานะ: ฉัน   |   ผู้ที่มีสิทธิ์เข้าถึง: ทุกคน
 * 4) คัดลอก URL ที่ลงท้ายด้วย /exec ไปวางใน index.html (ค่า SHEET_API)
 */

const SHEET_ID = '1bV3ToecEFKpfnIgWJASQaVAMfgItPQ0hB_uPvSCUJnQ';
const ADMIN_KEY = 'nsskru12345';   // ต้องตรงกับรหัสแอดมินในแอป

const RESULT_HEADERS = ['id', 'เวลาเล่น', 'ชื่อเล่น', 'โหมด', 'คะแนน', 'ถูก', 'ผิด', 'ร้อยละถูก',
  'คอมโบสูงสุด', 'ชีวิตเหลือ', 'หัวข้อที่ตอบ (JSON)', 'บันทึกเมื่อ'];
const VISIT_HEADERS = ['id', 'เวลาเข้าใช้', 'ชื่อเล่น', 'บันทึกเมื่อ', 'ใช้งานล่าสุด', 'เวลาใช้งาน (นาที)'];
const MODE_TH = { balloon: 'แทงลูกโป่ง', clinical: 'สถานการณ์คลินิก', shooter: 'ยิงไข่' };

function setup() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  formatSheets_(ss, sheet_(ss, 'Results', RESULT_HEADERS), sheet_(ss, 'Visits', VISIT_HEADERS));
}

/* เวลาไทย และแสดงวันที่ + เวลาเข้าใช้แบบ 09/11/2026  09.45 น. (ทำครั้งเดียว) */
const TIME_FMT = 'dd/mm/yyyy  HH.mm "น."';
function formatSheets_(ss, rs, vs) {
  ss.setSpreadsheetTimeZone('Asia/Bangkok');
  rs.getRange('B:B').setNumberFormat(TIME_FMT);
  rs.getRange('L:L').setNumberFormat(TIME_FMT);
  vs.getRange('B:B').setNumberFormat(TIME_FMT);
  vs.getRange('D:E').setNumberFormat(TIME_FMT);
  vs.getRange('F:F').setNumberFormat('0.0');
  rs.getRange('1:1').setNumberFormat('@'); vs.getRange('1:1').setNumberFormat('@');
  PropertiesService.getScriptProperties().setProperty('fmt', 'v3');
}

function sheet_(ss, name, headers) {
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  if (sh.getLastRow() === 0) {
    sh.appendRow(headers);
    sh.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#B4241F').setFontColor('#ffffff');
    sh.setFrozenRows(1);
  } else if (sh.getLastColumn() < headers.length) {
    // เพิ่มหัวคอลัมน์ใหม่ให้แท็บเดิม (เช่น เวลาใช้งาน)
    const from = sh.getLastColumn() + 1;
    sh.getRange(1, from, 1, headers.length - from + 1).setValues([headers.slice(from - 1)])
      .setFontWeight('bold').setBackground('#B4241F').setFontColor('#ffffff');
  }
  return sh;
}

/* ---------- write (from students' devices) ---------- */
function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const body = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    const items = body.action === 'batch' ? (body.items || []).slice(0, 200) : [body];
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const rs = sheet_(ss, 'Results', RESULT_HEADERS);
    const vs = sheet_(ss, 'Visits', VISIT_HEADERS);
    if (PropertiesService.getScriptProperties().getProperty('fmt') !== 'v3') formatSheets_(ss, rs, vs);
    const seenR = idSet_(rs), seenV = idSet_(vs);
    const now = new Date();
    const rRows = [], vRows = [], sessions = [];
    items.forEach(it => {
      if (!it) return;
      if (it.action === 'saveResult' && it.record) {
        const r = it.record, id = clean_(r.id, 40);
        if (!id || seenR.has(id)) return;
        seenR.add(id);
        const its = Array.isArray(r.items) ? r.items.slice(0, 300).map(x => ({ t: clean_(x.t, 80), o: x.o ? 1 : 0 })) : [];
        rRows.push([id, toDate_(r.at), clean_(r.name, 24), MODE_TH[r.mode] || clean_(r.mode, 20),
          num_(r.score), num_(r.right), num_(r.wrong), num_(r.acc), num_(r.combo), num_(r.lives),
          JSON.stringify(its), now]);
      } else if (it.action === 'saveVisit' && it.visit) {
        const v = it.visit, id = clean_(v.id, 40);
        if (!id || seenV.has(id)) return;
        seenV.add(id);
        vRows.push([id, toDate_(v.at), clean_(v.name, 24), now, '', '']);
      } else if (it.action === 'saveSession' && it.visit) {
        sessions.push(it);
      }
    });
    if (rRows.length) rs.getRange(rs.getLastRow() + 1, 1, rRows.length, RESULT_HEADERS.length).setValues(rRows);
    if (vRows.length) vs.getRange(vs.getLastRow() + 1, 1, vRows.length, VISIT_HEADERS.length).setValues(vRows);
    sessions.forEach(it => saveSession_(vs, it, seenV, now));
    return json_({ ok: true, saved: rRows.length + vRows.length + sessions.length });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  } finally {
    try { lock.releaseLock(); } catch (e) { }
  }
}

/* เวลาใช้งาน: อัปเดตแถวของการเข้าใช้ครั้งนั้น (นับเฉพาะช่วงที่เปิดหน้าแอปและมีการใช้งาน) */
function saveSession_(vs, it, seenV, now) {
  const v = it.visit, id = clean_(v.id, 40);
  if (!id) return;
  const sec = Math.max(0, Math.min(24 * 3600, Number(it.sec) || 0));
  const minutes = Math.round(sec / 6) / 10;
  const last = toDate_(it.lastAt);
  const found = vs.getRange(2, 1, Math.max(1, vs.getLastRow() - 1), 1)
    .createTextFinder(id).matchEntireCell(true).findNext();
  if (found) {
    const row = found.getRow();
    const prev = Number(vs.getRange(row, 6).getValue()) || 0;
    if (minutes >= prev) vs.getRange(row, 5, 1, 2).setValues([[last, minutes]]);
  } else if (!seenV.has(id)) {
    seenV.add(id);
    vs.appendRow([id, toDate_(v.at), clean_(v.name, 24), now, last, minutes]);
  }
}

/* ---------- read (admin dashboard only) ---------- */
function doGet(e) {
  const p = (e && e.parameter) || {};
  if (p.action === 'ping') return json_({ ok: true, app: 'ABG NSSKRU' });
  if (p.action !== 'list') return json_({ ok: true, app: 'ABG NSSKRU', hint: 'POST to save, GET ?action=list&key=… to read' });
  if (p.key !== ADMIN_KEY) return json_({ ok: false, error: 'รหัสแอดมินไม่ถูกต้อง' });
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const modeKey = {}; Object.keys(MODE_TH).forEach(k => modeKey[MODE_TH[k]] = k);
  const results = rows_(sheet_(ss, 'Results', RESULT_HEADERS), 3000).map(r => ({
    id: String(r[0]), at: iso_(r[1]), name: String(r[2]), mode: modeKey[r[3]] || String(r[3]),
    score: +r[4] || 0, right: +r[5] || 0, wrong: +r[6] || 0, acc: +r[7] || 0, combo: +r[8] || 0, lives: +r[9] || 0,
    items: parse_(r[10]), sid: '', group: '', gender: '', role: ''
  }));
  const visits = rows_(sheet_(ss, 'Visits', VISIT_HEADERS), 6000).map(r => ({
    id: String(r[0]), at: iso_(r[1]), name: String(r[2]), kind: 'visit',
    lastAt: r[4] ? iso_(r[4]) : '', sec: Math.round((Number(r[5]) || 0) * 60),
    sid: '', group: '', gender: '', role: ''
  }));
  return json_({ ok: true, results, visits });
}

/* ---------- helpers ---------- */
function rows_(sh, max) {
  const last = sh.getLastRow();
  if (last < 2) return [];
  const start = Math.max(2, last - max + 1);
  return sh.getRange(start, 1, last - start + 1, sh.getLastColumn()).getValues().filter(r => r[0]);
}
function idSet_(sh) {
  const last = sh.getLastRow();
  if (last < 2) return new Set();
  const start = Math.max(2, last - 5000);
  return new Set(sh.getRange(start, 1, last - start + 1, 1).getValues().map(r => String(r[0])));
}
function clean_(v, n) {
  let s = String(v == null ? '' : v).replace(/[\u0000-\u001f]/g, ' ').trim().slice(0, n);
  if (/^[=+\-@]/.test(s)) s = "'" + s;   // กันสูตรใน Sheet
  return s;
}
function num_(v) { const n = Number(v); return isFinite(n) ? Math.round(n * 100) / 100 : 0; }
function toDate_(v) { const d = new Date(v); return isNaN(d) ? new Date() : d; }
function iso_(v) { const d = v instanceof Date ? v : new Date(v); return isNaN(d) ? '' : d.toISOString(); }
function parse_(v) { try { return JSON.parse(v) || []; } catch (e) { return []; } }
function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
