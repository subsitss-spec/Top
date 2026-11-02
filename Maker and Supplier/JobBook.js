/**
 * JobBook — เชื่อมกับชีต "2026 Received and Return Machines Job of NDT-B" (ฐานข้อมูลงานรับ–คืนเครื่อง)
 * NDT-B Service Portal V1.03
 *
 * - แต่ละปีเป็นแท็บแยก (2015-2022, 2023, 2024, 2025, 2026 …) หัวตาราง 2 แถว ข้อมูลเริ่มแถว 3
 * - แท็บ 2023 ขึ้นไปมีคอลัมน์ A = No. ส่วนแท็บ 2015-2022 ไม่มี (คอลัมน์เลื่อนซ้าย 1) → หาตำแหน่งจากหัว "Status"
 * - บันทึกใบ Request → เพิ่ม/เติมแถวงานในแท็บปีนั้นอัตโนมัติ (เติมเฉพาะช่องที่ยังว่าง ไม่ทับข้อมูลที่คนกรอกเอง)
 */
var JOB_BOOK_ID = '1e9Wo5nmeqh2yxKG0gD9bsxIeK2lpktATMGiHdewE2OM';
var JOB_FIRST_ROW = 3;
// ลำดับคอลัมน์ของแท็บแบบใหม่ (มีคอลัมน์ No.) — index เริ่ม 0
var JOB_KEYS = ['no', 'jobNo', 'received', 'quoteDate', 'quoteNo', 'poDate', 'poNo', 'sentOut', 'backFromSupplier',
  'status', 'price', 'sales', 'engineer', 'customer', 'user', 'desc', 'location', 'qty', 'vendor', 'returned', 'picReceive', 'picReturn'];
var JOB_DATE_KEYS = { received: 1, quoteDate: 1, poDate: 1, sentOut: 1, backFromSupplier: 1, returned: 1 };

function jobBook_() { return SpreadsheetApp.openById(JOB_BOOK_ID); }

/** แท็บที่ชื่อเป็นปี หรือช่วงปี เรียงจากใหม่ไปเก่า */
function jobTabs_(book) {
  return book.getSheets()
    .map(function (sh) { var m = /(\d{4})\s*$/.exec(sh.getName()); return m ? { sh: sh, name: sh.getName(), year: +m[1] } : null; })
    .filter(Boolean)
    .sort(function (a, b) { return b.year - a.year; });
}

/** หา offset คอลัมน์ของแท็บ: แท็บใหม่ Status อยู่คอลัมน์ J (index 9) */
function jobLayout_(sh) {
  var head = sh.getRange(1, 1, 1, Math.min(sh.getMaxColumns(), 26)).getDisplayValues()[0];
  var st = -1;
  head.forEach(function (h, i) { if (st < 0 && /^status$/i.test(String(h).trim())) st = i; });
  var offset = st < 0 ? 0 : st - 9;           // 0 = มีคอลัมน์ No., -1 = ไม่มี
  return { offset: offset, width: JOB_KEYS.length + offset };
}

function jobCell_(v, tz) {
  if (v instanceof Date) return Utilities.formatDate(v, tz, 'yyyy-MM-dd');
  return v == null ? '' : String(v).trim();
}

function readJobTab_(t, tz) {
  var sh = t.sh, lay = jobLayout_(sh);
  var last = sh.getLastRow();
  if (last < JOB_FIRST_ROW) return [];
  var vals = sh.getRange(JOB_FIRST_ROW, 1, last - JOB_FIRST_ROW + 1, lay.width).getValues();
  var out = [];
  vals.forEach(function (r, i) {
    var o = { tab: t.name, gid: sh.getSheetId(), row: JOB_FIRST_ROW + i };
    JOB_KEYS.forEach(function (k, j) {
      var c = j + lay.offset;
      o[k] = c < 0 ? '' : jobCell_(r[c], tz);
    });
    if (!/^\d+$/.test(String(o.jobNo).replace(/,/g, ''))) return;
    o.jobNo = String(o.jobNo).replace(/,/g, '');
    out.push(o);
  });
  return out;
}

/** สำหรับหน้า Job Tracker — years: จำนวนแท็บล่าสุดที่จะโหลด (ค่าเริ่มต้น 2), 'all' = ทุกแท็บ */
function getJobs(years) {
  var book = jobBook_(), tz = Session.getScriptTimeZone();
  var tabs = jobTabs_(book);
  var pick = years === 'all' ? tabs : tabs.slice(0, Math.max(1, +years || 2));
  var jobs = [];
  pick.forEach(function (t) { jobs = jobs.concat(readJobTab_(t, tz)); });
  return {
    jobs: jobs,
    tabs: tabs.map(function (t) { return t.name; }),
    loaded: pick.map(function (t) { return t.name; }),
    bookUrl: book.getUrl(),
    today: Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd')
  };
}

/** เลข Job สูงสุดในแท็บล่าสุด 2 แท็บ */
function jobBookMaxNo_() {
  try {
    var book = jobBook_(), tz = Session.getScriptTimeZone(), max = 0;
    jobTabs_(book).slice(0, 2).forEach(function (t) {
      readJobTab_(t, tz).forEach(function (j) { var n = +j.jobNo; if (n > max) max = n; });
    });
    return max;
  } catch (e) { return 0; }
}

/** หางานตามเลข Job ในแท็บล่าสุด 3 แท็บ */
function findJob_(jobNo) {
  var book = jobBook_(), tz = Session.getScriptTimeZone(), hit = null;
  jobTabs_(book).slice(0, 3).some(function (t) {
    readJobTab_(t, tz).some(function (j) { if (j.jobNo === String(jobNo)) { hit = j; return true; } return false; });
    return !!hit;
  });
  return hit;
}

/** แปลงแถวงานในชีตเป็นข้อมูลตั้งต้นของใบ Request (ใช้ตอนเปิดใบของงานที่ยังไม่เคยทำใบในเว็บ) */
function jobToRequest_(j) {
  var eng = j.engineer || '', sales = j.sales || '';
  var serial = '', product = j.desc || '';
  var m = /\s*(?:S\/N|Serial)\s*[:：]?\s*([^\s,()]+)\s*\)?\s*$/i.exec(product);
  if (m) { serial = m[1]; product = product.slice(0, m.index).replace(/[,(]\s*$/, '').trim(); }
  return {
    jobNo: j.jobNo, date: j.received || '', receive: true, ret: !!j.returned,
    staffType: eng ? 'Engineer' : (sales ? 'Sales' : 'Engineer'), staffName: eng || sales,
    customer: j.customer, userName: j.user, product: product, serial: serial,
    sig: { recvStaff: '', recvDate: j.received || '', custSent: '', custSentDate: j.received || '', retStaff: '', retDate: j.returned || '', custRet: '', custRetDate: j.returned || '' },
    _fromJobBook: true
  };
}

function isoToDate_(iso) {
  var m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(String(iso || ''));
  return m ? new Date(+m[1], +m[2] - 1, +m[3]) : null;
}

/** แท็บของปีนั้น ถ้ายังไม่มี (เช่นขึ้นปีใหม่) สร้างจากหัวตารางของแท็บล่าสุด */
function jobTabForYear_(book, year) {
  var name = String(year);
  var sh = book.getSheetByName(name);
  if (sh) return sh;
  var latest = jobTabs_(book)[0];
  if (!latest) throw new Error('ไม่พบแท็บปีในชีตงาน');
  sh = latest.sh.copyTo(book).setName(name);
  var last = sh.getLastRow();
  if (last >= JOB_FIRST_ROW) sh.getRange(JOB_FIRST_ROW, 1, last - JOB_FIRST_ROW + 1, sh.getMaxColumns()).clearContent();
  book.setActiveSheet(sh); book.moveActiveSheet(book.getNumSheets());
  return sh;
}

/**
 * เพิ่ม/เติมแถวงานจากใบ Request
 * - งานใหม่: เพิ่มแถวต่อท้าย (No. ต่อจากเดิม, Status = Receiving)
 * - งานเดิม: เติมเฉพาะช่องที่ยังว่าง ไม่ทับข้อมูลที่กรอกไว้แล้ว
 */
function syncJobBook_(d) {
  var book = jobBook_(), tz = Session.getScriptTimeZone();
  var jobNo = String(d.jobNo || '').replace(/,/g, '').trim();
  if (!/^\d+$/.test(jobNo)) return { ok: false, reason: 'Job No. ไม่ใช่ตัวเลข จึงไม่ได้ลงชีตงาน' };
  var s = d.sig || {};
  var recv = d.receive ? (s.recvDate || d.date) : '';
  var ret = d.ret ? (s.retDate || d.date) : '';
  var desc = [d.maker, d.product, d.model].filter(function (x) { return String(x || '').trim(); }).join(' ') + (d.serial ? ' S/N : ' + d.serial : '');
  var staff = String(d.staffName || '').trim();
  var want = {
    received: recv ? isoToDate_(recv) : null,
    returned: ret ? isoToDate_(ret) : null,
    engineer: d.staffType === 'Engineer' ? staff : '',
    sales: d.staffType === 'Sales' ? staff : '',
    customer: d.customer || '', user: d.userName || '', desc: desc.trim(),
    picReceive: d.receive ? staff : '', picReturn: d.ret ? staff : ''
  };

  var hit = findJob_(jobNo);
  var sh, lay, row, current;
  if (hit) {
    sh = book.getSheetByName(hit.tab); lay = jobLayout_(sh); row = hit.row; current = hit;
  } else {
    var dt = isoToDate_(recv || d.date) || new Date();
    sh = jobTabForYear_(book, dt.getFullYear()); lay = jobLayout_(sh);
    var jobs = readJobTab_({ sh: sh, name: sh.getName() }, tz);
    var lastRow = jobs.length ? Math.max.apply(null, jobs.map(function (j) { return j.row; })) : JOB_FIRST_ROW - 1;
    var maxNo = jobs.reduce(function (m, j) { var n = parseInt(j.no, 10); return n > m ? n : m; }, 0);
    row = lastRow + 1; current = {};
    want.no = maxNo + 1; want.jobNo = +jobNo; want.status = 'Receiving'; want.qty = 1;
  }

  var range = sh.getRange(row, 1, 1, lay.width);
  var vals = range.getValues()[0];
  var filled = [];
  JOB_KEYS.forEach(function (k, j) {
    var c = j + lay.offset;
    if (c < 0 || want[k] == null || want[k] === '') return;
    if (String(vals[c] == null ? '' : vals[c]).trim() !== '') return;     // ไม่ทับของเดิม
    vals[c] = (want[k] instanceof Date || typeof want[k] === 'number') ? want[k] : safeCell_(want[k]);
    filled.push(k);
  });
  if (filled.length) {
    range.setValues([vals]);
    filled.forEach(function (k) {
      if (JOB_DATE_KEYS[k]) sh.getRange(row, JOB_KEYS.indexOf(k) + lay.offset + 1).setNumberFormat('d-mmm-yyyy');
    });
  }
  return { ok: true, created: !hit, tab: sh.getName(), row: row, filled: filled, url: book.getUrl() + '#gid=' + sh.getSheetId() + '&range=A' + row };
}
