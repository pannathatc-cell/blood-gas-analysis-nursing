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
const VISIT_HEADERS = ['id', 'เวลาเข้าใช้', 'ชื่อเล่น', 'บันทึกเมื่อ'];
const MODE_TH = { balloon: 'แทงลูกโป่ง', clinical: 'สถานการณ์คลินิก', shooter: 'ยิงไข่' };

function setup() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  sheet_(ss, 'Results', RESULT_HEADERS);
  sheet_(ss, 'Visits', VISIT_HEADERS);
}

function sheet_(ss, name, headers) {
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  if (sh.getLastRow() === 0) {
    sh.appendRow(headers);
    sh.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#B4241F').setFontColor('#ffffff');
    sh.setFrozenRows(1);
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
    const seenR = idSet_(rs), seenV = idSet_(vs);
    const now = new Date();
    const rRows = [], vRows = [];
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
        vRows.push([id, toDate_(v.at), clean_(v.name, 24), now]);
      }
    });
    if (rRows.length) rs.getRange(rs.getLastRow() + 1, 1, rRows.length, RESULT_HEADERS.length).setValues(rRows);
    if (vRows.length) vs.getRange(vs.getLastRow() + 1, 1, vRows.length, VISIT_HEADERS.length).setValues(vRows);
    return json_({ ok: true, saved: rRows.length + vRows.length });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  } finally {
    try { lock.releaseLock(); } catch (e) { }
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
    id: String(r[0]), at: iso_(r[1]), name: String(r[2]), kind: 'visit', sid: '', group: '', gender: '', role: ''
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
