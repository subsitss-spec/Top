/**
 * NDT-B Service Portal  V1.03  (2 พ.ย. 2026) — เชื่อมชีตงาน + Job Tracker (JobBook.gs, Jobs.html)
 * หน้าหลัก → (1) Maker Directory  (2) Request Check Sheet (Receive / Return) พิมพ์เป็น PDF
 *
 * ผูกสคริปต์นี้กับสเปรดชีต "Contact Maker and Supplier" (Extensions → Apps Script)
 *  - แท็บแรก (ชีต1)          = รายชื่อ Maker / Supplier  (คอลัมน์ A–K)
 *  - แท็บ RequestCheckSheet  = ใบ Request Check Sheet ที่บันทึกไว้ (สร้างให้อัตโนมัติ)
 *  - แท็บ Staff              = รายชื่อพนักงาน NDT (ชื่อ / มือถือ / อีเมล / ตำแหน่ง) แก้ได้เองในชีต
 */

// ======================= ค่าคงที่ =======================
var APP_VERSION = 'V1.03';
var MAKER_SHEET_NAME = 'ชีต1';           // ถ้าไม่เจอชื่อนี้ จะใช้แท็บแรกของไฟล์
var REQ_SHEET_NAME = 'RequestCheckSheet';
var STAFF_SHEET_NAME = 'Staff';
var MAX_LOGO_BYTES = 4 * 1024 * 1024;

// คอลัมน์ชีต Maker: A Company | B Country | C Website | D Telephone | E Service PIC | F E-mail | G Sales PIC | H E-mail | I Technical website | J New Price machine | K Logo
var MAKER_FIELDS = ['company', 'country', 'website', 'tel', 'servicePic', 'serviceEmail', 'salesPic', 'salesEmail', 'techWeb', 'priceLink', 'logo'];
var MAKER_HEADERS = ['Company', 'Country', 'Website', 'Telephone', 'Service PIC', 'E-mail', 'Sales PIC', 'E-mail', 'Technical website', 'New Price machine', 'Logo'];

// รายชื่อเริ่มต้น (ตามแบบฟอร์มเดิม) — เขียนลงแท็บ Staff ครั้งแรกเท่านั้น หลังจากนั้นแก้ในชีตได้เลย
var DEFAULT_STAFF = [
  ['Mr. Hirofumi Yachi', '', 'yachi@n-denkei.co.th', 'Engineer'],
  ['Mr. Wisut Saphoemsap', '', 'wisut@n-denkei.co.th', 'Engineer'],
  ['Mr. Poontana Munkong', '', 'poontana@n-denkei.co.th', 'Engineer'],
  ['Mr. Suphawat Prakaew', '', 'suphawat@n-denkei.co.th', 'Engineer'],
  ['Mr. Teerapol Rattanachart', '', 'teerapol@n-denkei.co.th', 'Engineer'],
  ['Mr. Thanawat Amornsriserm', '', 'thanawat@n-denkei.co.th', 'Engineer'],
  ['Mr. Witchanon Anonsathian', '', 'witchanon@n-denkei.co.th', 'Engineer'],
  ['Mr. Sabsit Sungsri', '', 'sabsit@n-denkei.co.th', 'Engineer']
];
var STAFF_HEADERS = ['Name', 'Mobile', 'E-mail', 'Role (Engineer/Sales)', 'Show in reply list (Y/N)', 'LINE userId'];

var REQ_HEADERS = ['Job No.', 'Date', 'Receive', 'Return', 'Staff type', 'NDT Staff', 'Warranty', 'Repair', 'Calibrate',
  'Sold by', 'Customer Name', 'User Name', 'Department', 'E-mail', 'Telephone', 'Maker', 'Product Name', 'Model', 'Serial',
  'Input voltage', 'Phase', 'Power cable', 'Symptom', 'Reply to', 'NDT staff (Receive)', 'Receive date',
  'Customer staff (Sent out)', 'Sent out date', 'NDT staff (Return)', 'Return date', 'Customer staff (Return)',
  'Customer return date', 'Updated', 'Updated by', 'Data (JSON)', 'PDF'];
var JSON_COL = 35;   // คอลัมน์ AI = Data (JSON)
var PDF_COL = 36;    // คอลัมน์ AJ = ลิงก์ PDF ล่าสุด
var PDF_FOLDER_PROP = 'PDF_FOLDER_ID';

// ======================= Web app =======================
function doGet(e) {
  var page = (e && e.parameter && e.parameter.page) || 'home';
  var file = { home: 'Home', makers: 'Makers', request: 'Request', jobs: 'Jobs' }[page] || 'Home';
  var title = { Home: 'NDT-B Service Portal', Makers: 'Maker Directory', Request: 'Request Check Sheet', Jobs: 'Job Tracker' }[file];
  var t = HtmlService.createTemplateFromFile(file);
  t.appUrl = ScriptApp.getService().getUrl();
  return t.evaluate()
    .setTitle(title)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(name) {
  return HtmlService.createHtmlOutputFromFile(name).getContent();
}

function onOpen() {
  try {
    SpreadsheetApp.getUi().createMenu('NDT Service')
      .addItem('ตั้งค่าชีต (ครั้งแรก)', 'setupAll')
      .addToUi();
  } catch (e) { /* รันจากตัวแก้ไขสคริปต์ — ไม่มี UI */ }
}

/** รันครั้งแรกครั้งเดียว: เติมหัวคอลัมน์ Logo + สร้างแท็บ Staff และ RequestCheckSheet */
function setupAll() {
  setupMakerSheet();
  setupRequestSheets();
  var msg = 'ตั้งค่าเรียบร้อย: ชีต Maker, Staff, ' + REQ_SHEET_NAME;
  try { SpreadsheetApp.getUi().alert(msg); } catch (e) { Logger.log(msg); }
}

// ======================= ส่วนกลาง =======================
function ss_() { return SpreadsheetApp.getActiveSpreadsheet(); }

function withLock_(fn) {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(20000)) throw new Error('มีคนกำลังบันทึกอยู่ ลองใหม่อีกครั้งในไม่กี่วินาที');
  try { return fn(); } finally { lock.releaseLock(); }
}

/** กันค่าที่ขึ้นต้นด้วย = + - @ ไม่ให้ชีตตีความเป็นสูตร (เช่นเบอร์ +81…) */
function safeCell_(v) {
  var s = v == null ? '' : String(v);
  return /^[=+\-@]/.test(s) ? "'" + s : s;
}

function str_(v) { return v == null ? '' : String(v).trim(); }

// ======================= Maker Directory =======================
function getMakerSheet_() {
  var ss = ss_();
  return ss.getSheetByName(MAKER_SHEET_NAME) || ss.getSheets()[0];
}

function setupMakerSheet() {
  var sh = getMakerSheet_();
  if (sh.getLastRow() === 0) sh.getRange(1, 1, 1, MAKER_HEADERS.length).setValues([MAKER_HEADERS]).setFontWeight('bold');
  else if (!str_(sh.getRange(1, 11).getValue())) sh.getRange(1, 11).setValue('Logo').setFontWeight('bold');
}

/** ตัดวงเล็บท้ายชื่อ + เครื่องหมาย/ช่องว่าง + ตัวพิมพ์เล็ก — ต้องตรงกับฝั่ง Makers.html */
function groupKey_(name) {
  return String(name || '').replace(/\([^)]*\)\s*$/, ' ').toLowerCase().replace(/[^a-z0-9฀-๿]+/g, '');
}

function readMakerRows_() {
  var sh = getMakerSheet_();
  var last = sh.getLastRow();
  if (last < 2) return [];
  var vals = sh.getRange(2, 1, last - 1, 11).getDisplayValues();
  var rows = [];
  vals.forEach(function (v, i) {
    if (!v.some(function (c) { return String(c).trim(); })) return;
    var r = { row: i + 2 };
    MAKER_FIELDS.forEach(function (f, j) { r[f] = v[j] == null ? '' : String(v[j]); });
    rows.push(r);
  });
  return rows;
}

function getData() { return { rows: readMakerRows_() }; }

function saveContact(d) {
  d = d || {};
  var company = str_(d.company);
  if (!company) throw new Error('กรุณากรอกชื่อบริษัท (Company)');
  return withLock_(function () {
    var sh = getMakerSheet_();
    var values = MAKER_FIELDS.slice(0, 10).map(function (f) { return safeCell_(str_(d[f])); });
    var row;
    if (d.row) {
      row = Number(d.row);
      if (!(row >= 2 && row <= sh.getLastRow())) throw new Error('ไม่พบแถวที่ต้องการแก้ไข (ข้อมูลอาจถูกลบไปแล้ว) กรุณารีเฟรชหน้า');
      sh.getRange(row, 1, 1, 10).setValues([values]);
    } else {
      // แผนก/ผู้ติดต่อใหม่ของบริษัทเดิม ใช้โลโก้เดิมของกลุ่ม
      var key = groupKey_(company);
      var existingLogo = '';
      readMakerRows_().some(function (r) { if (groupKey_(r.company) === key && r.logo) { existingLogo = r.logo; return true; } return false; });
      sh.appendRow(values.concat([existingLogo]));
      row = sh.getLastRow();
    }
    if (d.logo && d.logo.dataUrl) setGroupLogo_(company, uploadLogo_(d.logo.dataUrl, company));
    else if (d.logoUrl) setGroupLogo_(company, str_(d.logoUrl));
    return getData();
  });
}

function setLogo(company, logo) {
  return withLock_(function () {
    var url = '';
    if (logo && logo.dataUrl) url = uploadLogo_(logo.dataUrl, company);
    else if (logo && logo.url) {
      url = str_(logo.url);
      if (!/^https?:\/\//i.test(url)) throw new Error('ลิงก์รูปต้องขึ้นต้นด้วย https://');
    }
    setGroupLogo_(company, url);
    return getData();
  });
}

function setGroupLogo_(company, url) {
  var sh = getMakerSheet_();
  var key = groupKey_(company);
  readMakerRows_().forEach(function (r) {
    if (groupKey_(r.company) === key) sh.getRange(r.row, 11).setValue(url);
  });
}

function getLogoFolder_() {
  var props = PropertiesService.getScriptProperties();
  var id = props.getProperty('LOGO_FOLDER_ID');
  if (id) { try { return DriveApp.getFolderById(id); } catch (e) { /* ถูกลบไป สร้างใหม่ */ } }
  var folder = DriveApp.createFolder('Maker Logos');
  props.setProperty('LOGO_FOLDER_ID', folder.getId());
  return folder;
}

function uploadLogo_(dataUrl, company) {
  var m = /^data:(image\/[\w.+-]+);base64,(.+)$/.exec(String(dataUrl || ''));
  if (!m) throw new Error('ไฟล์รูปไม่ถูกต้อง');
  var bytes = Utilities.base64Decode(m[2]);
  if (bytes.length > MAX_LOGO_BYTES) throw new Error('รูปใหญ่เกิน 4 MB');
  var ext = m[1] === 'image/png' ? 'png' : (m[1] === 'image/jpeg' ? 'jpg' : 'img');
  var name = String(company || 'logo').replace(/[\\\/:*?"<>|]+/g, ' ').trim().slice(0, 80) + '.' + ext;
  var file = getLogoFolder_().createFile(Utilities.newBlob(bytes, m[1], name));
  try { file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); }
  catch (e) { try { file.setSharing(DriveApp.Access.DOMAIN_WITH_LINK, DriveApp.Permission.VIEW); } catch (e2) {} }
  return 'https://drive.google.com/thumbnail?id=' + file.getId() + '&sz=w600';
}

/** ใช้ตอนกด "ปรับขนาด / ครอบรูป" โลโก้ที่อยู่ใน Drive แล้ว */
function getLogoData(url) {
  var m = /[?&]id=([\w-]{20,})/.exec(url) || /\/d\/([\w-]{20,})/.exec(url);
  var blob;
  if (m) blob = DriveApp.getFileById(m[1]).getBlob();
  else blob = UrlFetchApp.fetch(url, { muteHttpExceptions: false }).getBlob();
  return 'data:' + blob.getContentType() + ';base64,' + Utilities.base64Encode(blob.getBytes());
}

function deleteContact(row, company) {
  return withLock_(function () {
    var sh = getMakerSheet_();
    row = Number(row);
    if (!(row >= 2 && row <= sh.getLastRow())) throw new Error('ไม่พบแถวนี้ กรุณารีเฟรชหน้า');
    var actual = str_(sh.getRange(row, 1).getDisplayValue());
    if (actual !== str_(company)) throw new Error('ข้อมูลในชีตเปลี่ยนไปแล้ว (แถว ' + row + ' ไม่ใช่ "' + str_(company) + '") กรุณารีเฟรชหน้าแล้วลองใหม่');
    sh.deleteRow(row);
    return getData();
  });
}

// ======================= Request Check Sheet =======================
function setupRequestSheets() {
  var ss = ss_();
  var st = ss.getSheetByName(STAFF_SHEET_NAME);
  if (!st) {
    st = ss.insertSheet(STAFF_SHEET_NAME);
    st.getRange(1, 1, 1, STAFF_HEADERS.length).setValues([STAFF_HEADERS]).setFontWeight('bold');
    st.getRange(2, 1, DEFAULT_STAFF.length, 5).setValues(DEFAULT_STAFF.map(function (r) { return r.concat(['Y']); }));
    st.getRange(2, 2, DEFAULT_STAFF.length, 1).setNumberFormat('@');
    st.setFrozenRows(1);
    st.autoResizeColumns(1, 5);
  }
  if (!str_(st.getRange(1, 6).getValue())) st.getRange(1, 6).setValue('LINE userId').setFontWeight('bold');
  var rq = ss.getSheetByName(REQ_SHEET_NAME);
  if (!rq) {
    rq = ss.insertSheet(REQ_SHEET_NAME);
    rq.getRange(1, 1, 1, REQ_HEADERS.length).setValues([REQ_HEADERS]).setFontWeight('bold');
    rq.setFrozenRows(1);
    rq.getRange('A:A').setNumberFormat('@');
  }
  if (!str_(rq.getRange(1, PDF_COL).getValue())) rq.getRange(1, PDF_COL).setValue('PDF').setFontWeight('bold');
}

function getReqSheet_() {
  var sh = ss_().getSheetByName(REQ_SHEET_NAME);
  if (!sh) { setupRequestSheets(); sh = ss_().getSheetByName(REQ_SHEET_NAME); }
  return sh;
}

function readStaff_() {
  var sh = ss_().getSheetByName(STAFF_SHEET_NAME);
  if (!sh) { setupRequestSheets(); sh = ss_().getSheetByName(STAFF_SHEET_NAME); }
  var last = sh.getLastRow();
  if (last < 2) return [];
  return sh.getRange(2, 1, last - 1, 6).getDisplayValues()
    .filter(function (r) { return str_(r[0]); })
    .map(function (r) {
      return { name: str_(r[0]), mobile: str_(r[1]), email: str_(r[2]), role: str_(r[3]) || 'Engineer', reply: !/^n/i.test(str_(r[4])), lineLinked: /^U[0-9a-f]{32}$/.test(str_(r[5])) };
    });
}

/** ข้อมูลเริ่มต้นของหน้าใบ Request: Maker, พนักงาน, เลข Job ถัดไป, ลูกค้าเดิม, รายการล่าสุด */
function getRequestInit() {
  var list = readRequests_();
  var maxNo = 0;
  list.forEach(function (r) { var n = parseFloat(String(r.jobNo).replace(/,/g, '')); if (!isNaN(n) && n > maxNo) maxNo = n; });
  var customers = {};
  list.forEach(function (r) {
    var d = r.data; if (!d || !d.customer) return;
    var k = (d.customer + '|' + (d.userName || '')).toLowerCase();
    customers[k] = { customer: d.customer, userName: d.userName || '', department: d.department || '', email: d.email || '', tel: d.tel || '' };
  });
  return {
    makers: readMakerRows_(),
    staff: readStaff_(),
    nextJobNo: (function () { var m = Math.max(Math.floor(maxNo) || 0, jobBookMaxNo_()); return m ? String(m + 1) : ''; })(),
    customers: Object.keys(customers).map(function (k) { return customers[k]; }),
    recent: list.slice(-30).reverse().map(summary_),
    lineEnabled: lineReady_()
  };
}

function summary_(r) {
  var d = r.data || {};
  return { jobNo: r.jobNo, date: d.date || '', customer: d.customer || '', product: d.product || '', maker: d.maker || '', model: d.model || '', receive: !!d.receive, ret: !!d.ret };
}

function readRequests_() {
  var sh = getReqSheet_();
  var last = sh.getLastRow();
  if (last < 2) return [];
  var jsonCol = JSON_COL;
  var jobs = sh.getRange(2, 1, last - 1, 1).getDisplayValues();
  var data = sh.getRange(2, jsonCol, last - 1, 1).getValues();
  var out = [];
  jobs.forEach(function (j, i) {
    var no = str_(j[0]); if (!no) return;
    var d = null; try { d = JSON.parse(data[i][0]); } catch (e) {}
    out.push({ row: i + 2, jobNo: no, data: d });
  });
  return out;
}

function getRequest(jobNo) {
  jobNo = str_(jobNo);
  var hit = null;
  readRequests_().forEach(function (r) { if (r.jobNo === jobNo) hit = r; });
  if (!hit) {
    var job = null; try { job = findJob_(jobNo); } catch (e) {}
    if (job) return jobToRequest_(job);
    throw new Error('ไม่พบ Job No. ' + jobNo + ' ทั้งในใบที่บันทึกและในชีตงาน');
  }
  if (!hit.data) throw new Error('ใบ Job No. ' + jobNo + ' ไม่มีข้อมูลสำหรับเปิดซ้ำ (คอลัมน์ Data ว่าง)');
  return hit.data;
}

/** บันทึกใบ (ถ้ามี Job No. เดิมจะเขียนทับแถวเดิม) */
function saveRequest(d) {
  d = d || {};
  var jobNo = str_(d.jobNo);
  if (!jobNo) throw new Error('กรุณากรอก Job No.');
  if (!str_(d.customer)) throw new Error('กรุณากรอกชื่อลูกค้า (Customer Name)');
  return withLock_(function () {
    var sh = getReqSheet_();
    var existing = null;
    readRequests_().forEach(function (r) { if (r.jobNo === jobNo) existing = r; });
    var s = d.sig || {};
    var row = [
      jobNo, d.date || '', d.receive ? 'Y' : '', d.ret ? 'Y' : '', d.staffType || '', d.staffName || '',
      d.warranty === 'in' ? 'In warranty' : (d.warranty === 'out' ? 'Out of warranty' : ''),
      d.repair ? 'Y' : '', d.calibrate ? 'Y' : '', d.soldBy || '', d.customer || '', d.userName || '',
      d.department || '', d.email || '', d.tel || '', d.maker || '', d.product || '', d.model || '', d.serial || '',
      d.voltage || '', d.phase === 'Other' ? 'Other: ' + (d.phaseOther || '') : (d.phase || ''), d.powerCable || '',
      d.symptom || '', (d.replyTo || []).join(', '), s.recvStaff || '', s.recvDate || '', s.custSent || '',
      s.custSentDate || '', s.retStaff || '', s.retDate || '', s.custRet || '', s.custRetDate || '',
      new Date(), Session.getActiveUser().getEmail() || '', JSON.stringify(d)
    ].map(function (v) { return v instanceof Date ? v : safeCell_(v); });
    if (existing) sh.getRange(existing.row, 1, 1, row.length).setValues([row]);
    else sh.appendRow(row);
    var jb; try { jb = syncJobBook_(d); } catch (e) { jb = { ok: false, reason: String(e.message || e) }; }
    return { jobNo: jobNo, updated: !!existing, jobBook: jb, init: getRequestInit() };
  });
}

// ======================= LINE แจ้งเตือนกลุ่ม (Messaging API) =======================
// ตั้งค่าใน Project Settings → Script Properties:
//   LINE_CHANNEL_TOKEN = Channel access token (long-lived) ของ LINE Official Account / Messaging API channel
//   LINE_GROUP_ID      = ไม่ต้องใส่เอง — เชิญบอทเข้ากลุ่มแล้วพิมพ์  /setgroup  ในกลุ่ม
// ผูก Engineer กับ LINE: ให้แต่ละคนพิมพ์ในกลุ่ม  /register ชื่อ  เช่น  /register Sabsit  (บันทึก userId ลงแท็บ Staff คอลัมน์ F)
// หมายเหตุ: บอท LINE ส่งไฟล์ PDF เป็นไฟล์แนบไม่ได้ จึงเก็บ PDF ไว้ใน Google Drive แล้วส่งลิงก์เข้ากลุ่ม

function lineProp_(k) { return PropertiesService.getScriptProperties().getProperty(k) || ''; }
function lineReady_() { return !!(lineProp_('LINE_CHANNEL_TOKEN') && lineProp_('LINE_GROUP_ID')); }

function lineApi_(path, payload, method) {
  var token = lineProp_('LINE_CHANNEL_TOKEN');
  if (!token) throw new Error('ยังไม่ได้ตั้ง LINE_CHANNEL_TOKEN ใน Script Properties');
  var opt = { method: method || 'post', muteHttpExceptions: true, headers: { Authorization: 'Bearer ' + token } };
  if (payload) { opt.contentType = 'application/json'; opt.payload = JSON.stringify(payload); }
  var res = UrlFetchApp.fetch('https://api.line.me/v2/bot/' + path, opt);
  var code = res.getResponseCode(), body = res.getContentText();
  if (code >= 300) {
    var msg = body; try { msg = JSON.parse(body).message || body; } catch (e) {}
    throw new Error('LINE ตอบกลับ ' + code + ': ' + msg);
  }
  return body ? JSON.parse(body) : {};
}

/** Webhook ของบอท — ใช้ URL เว็บแอป (/exec) เป็น Webhook URL ใน LINE Developers Console */
function doPost(e) {
  var events = [];
  try { events = JSON.parse(e.postData.contents).events || []; } catch (err) {}
  events.forEach(function (ev) {
    try { handleLineEvent_(ev); } catch (err) { Logger.log(err); }
  });
  return ContentService.createTextOutput(JSON.stringify({ ok: true })).setMimeType(ContentService.MimeType.JSON);
}

function handleLineEvent_(ev) {
  if (ev.type !== 'message' || !ev.message || ev.message.type !== 'text') return;
  var src = ev.source || {};
  var text = str_(ev.message.text);
  var reply = function (t) { lineApi_('message/reply', { replyToken: ev.replyToken, messages: [{ type: 'text', text: t }] }); };
  if (/^\/setgroup$/i.test(text)) {
    if (src.type !== 'group') return reply('ใช้คำสั่งนี้ในกลุ่ม LINE ที่ต้องการให้แจ้งเตือนเท่านั้น');
    var cur = lineProp_('LINE_GROUP_ID');
    if (cur && cur !== src.groupId) return reply('ตั้งกลุ่มแจ้งเตือนไว้กับกลุ่มอื่นแล้ว ถ้าต้องการย้าย ให้ลบ LINE_GROUP_ID ใน Script Properties ก่อน');
    PropertiesService.getScriptProperties().setProperty('LINE_GROUP_ID', src.groupId);
    return reply('ตั้งกลุ่มนี้เป็นกลุ่มแจ้งเตือน Request Check Sheet แล้ว ✅\nให้ Engineer แต่ละคนพิมพ์  /register ชื่อ  เช่น /register Sabsit');
  }
  var m = /^\/register\s+(.+)$/i.exec(text);
  if (m) {
    if (!src.userId) return reply('อ่าน userId ไม่ได้ (ต้องเพิ่มบอทเป็นเพื่อนก่อน แล้วลองใหม่)');
    var key = m[1].trim().toLowerCase();
    var sh = ss_().getSheetByName(STAFF_SHEET_NAME);
    if (!sh) { setupRequestSheets(); sh = ss_().getSheetByName(STAFF_SHEET_NAME); }
    var names = sh.getRange(2, 1, Math.max(1, sh.getLastRow() - 1), 1).getDisplayValues();
    var hits = [];
    names.forEach(function (r, i) { if (str_(r[0]) && r[0].toLowerCase().indexOf(key) >= 0) hits.push(i + 2); });
    if (hits.length !== 1) return reply(hits.length ? 'พบชื่อซ้ำหลายคน พิมพ์ชื่อให้ละเอียดขึ้น เช่น /register Sabsit Sungsri' : 'ไม่พบชื่อ "' + m[1].trim() + '" ในแท็บ Staff');
    sh.getRange(hits[0], 6).setValue(src.userId);
    return reply('ผูกบัญชี LINE กับ ' + sh.getRange(hits[0], 1).getDisplayValue() + ' แล้ว ✅ จะถูกแท็กเมื่อมีใบงานของคุณ');
  }
  if (/^\/id$/i.test(text)) return reply('userId: ' + (src.userId || '-') + (src.groupId ? '\ngroupId: ' + src.groupId : ''));
}

function getPdfFolder_() {
  var props = PropertiesService.getScriptProperties();
  var id = props.getProperty(PDF_FOLDER_PROP);
  if (id) { try { return DriveApp.getFolderById(id); } catch (e) {} }
  var folder = DriveApp.createFolder('Request Check Sheet PDF');
  props.setProperty(PDF_FOLDER_PROP, folder.getId());
  return folder;
}

/**
 * เก็บ PDF (สร้างจากหน้าเว็บ) ลง Drive แล้วแจ้งเตือนเข้ากลุ่ม LINE พร้อมแท็ก Engineer ของใบนั้น
 * @param {string} jobNo
 * @param {string} pdfDataUrl  data:application/pdf;base64,...
 */
function sendRequestToLine(jobNo, pdfDataUrl) {
  jobNo = str_(jobNo);
  var d = getRequest(jobNo);
  var m = /^data:application\/pdf[^,]*;base64,(.+)$/.exec(String(pdfDataUrl || ''));
  if (!m) throw new Error('ไฟล์ PDF ไม่ถูกต้อง');
  var blob = Utilities.newBlob(Utilities.base64Decode(m[1]), 'application/pdf', 'Request check sheet receive return ' + jobNo + '.pdf');
  var folder = getPdfFolder_();
  var old = folder.getFilesByName(blob.getName());
  while (old.hasNext()) old.next().setTrashed(true);   // เก็บไว้ไฟล์เดียวต่อ Job (ฉบับล่าสุด)
  var file = folder.createFile(blob);
  try { file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); }
  catch (e) { try { file.setSharing(DriveApp.Access.DOMAIN_WITH_LINK, DriveApp.Permission.VIEW); } catch (e2) {} }
  var url = 'https://drive.google.com/file/d/' + file.getId() + '/view';

  var sh = getReqSheet_();
  readRequests_().forEach(function (r) { if (r.jobNo === jobNo) sh.getRange(r.row, PDF_COL).setValue(url); });

  if (!lineReady_()) return { pdfUrl: url, sent: false, reason: 'ยังไม่ได้ตั้งค่าบอท LINE (LINE_CHANNEL_TOKEN / LINE_GROUP_ID)' };
  var staff = null;
  var sSh = ss_().getSheetByName(STAFF_SHEET_NAME);
  if (sSh && sSh.getLastRow() > 1) sSh.getRange(2, 1, sSh.getLastRow() - 1, 6).getDisplayValues().forEach(function (r) { if (str_(r[0]) === str_(d.staffName)) staff = r; });
  var lineId = staff && /^U[0-9a-f]{32}$/.test(str_(staff[5])) ? str_(staff[5]) : '';
  var msg = buildLineMessage_(d, url, lineId);
  lineApi_('message/push', { to: lineProp_('LINE_GROUP_ID'), messages: [msg] });
  return { pdfUrl: url, sent: true, tagged: !!lineId, staff: d.staffName || '' };
}

function buildLineMessage_(d, url, lineId) {
  var nb = function (v) { return String(v == null ? '' : v).replace(/[{}]/g, ''); };   // textV2 ใช้ {} เป็นตัวแทน — ตัดออกจากข้อมูลที่กรอก
  var kind = d.ret && !d.receive ? '📤 คืนเครื่อง' : (d.ret ? '📥📤 รับ/คืนเครื่อง' : '📥 รับเครื่องเข้าซ่อม');
  var jobTxt = /^\d{4,}$/.test(d.jobNo) ? Number(d.jobNo).toLocaleString('en-US') : nb(d.jobNo);
  var status = [d.warranty === 'in' ? 'In warranty' : (d.warranty === 'out' ? 'Out of warranty' : ''), d.repair ? 'Repair' : '', d.calibrate ? 'Calibrate' : ''].filter(String).join(' · ');
  var machine = [d.product, d.maker, d.model ? 'Model ' + d.model : '', d.serial ? 'S/N ' + d.serial : ''].filter(String).map(nb).join(' / ');
  var lines = [
    kind + '  Job No. ' + jobTxt,
    'ลูกค้า: ' + nb(d.customer || '-') + (d.userName ? ' (' + nb(d.userName) + ')' : ''),
    'เครื่อง: ' + (machine || '-'),
    d.symptom ? 'อาการ: ' + nb(d.symptom).replace(/\s+/g, ' ') : '',
    status ? 'สถานะ: ' + status : '',
    'PDF: ' + url
  ].filter(String).join('\n');
  if (!lineId) return { type: 'text', text: (d.staffName ? '@' + d.staffName + ' ' : '') + lines };
  return { type: 'textV2', text: '{eng} ' + lines, substitution: { eng: { type: 'mention', mentionee: { type: 'user', userId: lineId } } } };
}
