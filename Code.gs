// ============================================================
//  海德堡讀書會 Check-in System — Google Apps Script
//  Deploy as Web App: Execute as Me, Anyone access
//  QR Code 內容 = 電話號碼（唯一識別）
// ============================================================
var SHEET_ID     = '1YoOizf3BVTyQXUim-sdJPxp7awzsPWDZrAaXJMmyq54';
var ADMIN_PASS   = 'heidelberg2026';
var MEMBER_SHEET = '會員名冊';
var CONFIG_SHEET = '系統設定';

// 會員名冊欄位: A姓名 B電話 C出生年月日 D身份 E加入日期 F備註
// 活動分頁欄位: A姓名 B電話 C出席狀況 D用餐 E加購便當 F備註 G報到時間

function doGet(e) {
  var p = e.parameter || {};
  var action = p.action || '';
  var result;
  try {
    switch (action) {
      case 'info':               result = getInfo(); break;
      case 'member_lookup':      result = memberLookup(p.phone); break;
      case 'member_qr':          result = memberQR(p.phone); break;
      case 'event_register':     result = eventRegister(p); break;
      case 'event_checkin':      result = eventCheckin(p.phone); break;
      case 'event_search':       result = eventSearch(p.q); break;
      case 'event_stats':        result = eventStats(); break;
      case 'admin_login':        result = adminLogin(p.pass); break;
      case 'admin_event_guests': result = adminEventGuests(p.pass); break;
      case 'admin_members':      result = adminMembers(p.pass); break;
      case 'admin_create_event': result = adminCreateEvent(p); break;
      case 'admin_migrate':      result = adminMigrate(p.pass, p.sheetName); break;
      default: result = {ok: false, msg: 'Unknown action: ' + action};
    }
  } catch (err) {
    result = {ok: false, msg: err.message};
  }
  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

// ── Helpers ───────────────────────────────────────────────
function getSpreadsheet() { return SpreadsheetApp.openById(SHEET_ID); }

function getConfig() {
  var sh = getSpreadsheet().getSheetByName(CONFIG_SHEET);
  if (!sh) return {};
  var cfg = {};
  sh.getDataRange().getValues().forEach(function(r) { if (r[0]) cfg[r[0]] = r[1]; });
  return cfg;
}

function getEventSheet() {
  var name = getConfig()['currentEvent'] || '';
  return name ? getSpreadsheet().getSheetByName(name) : null;
}

function getMemberSheet() { return getSpreadsheet().getSheetByName(MEMBER_SHEET); }

function normalizePhone(p) { return String(p || '').replace(/[-\s]/g, ''); }

function findMemberByPhone(phone) {
  var sh = getMemberSheet();
  if (!sh || sh.getLastRow() < 2) return null;
  var norm = normalizePhone(phone);
  var data = sh.getRange(2, 1, sh.getLastRow() - 1, 6).getValues();
  for (var i = 0; i < data.length; i++) {
    if (normalizePhone(data[i][1]) === norm) {
      return {row: i+2, name: data[i][0], phone: data[i][1], birthday: data[i][2], memberType: data[i][3], joinDate: data[i][4], note: data[i][5]};
    }
  }
  return null;
}

// Detect if event sheet is old Google Form format (col A = timestamp)
function isOldFormat(sh) {
  if (sh.getLastRow() < 2) return false;
  var a1 = String(sh.getRange(2, 1).getValue());
  return a1.match(/^\d{4}-\d{2}-\d{2}T/) !== null;
}

// Convert old format sheet to new format in place
function convertOldFormatSheet(sh) {
  if (!isOldFormat(sh)) return;
  var data = sh.getRange(2, 1, sh.getLastRow() - 1, 8).getValues();
  sh.clear();
  sh.appendRow(['姓名','電話','出席狀況','用餐','加購便當','備註','報到時間']);
  data.forEach(function(r) {
    // old: A=timestamp B=姓名 C=電話 D=生日 E=? F=身份 G=出席狀況
    sh.appendRow([r[1], r[2], r[6]||'出席', '', '0', r[5]||'', '']);
  });
  sh.setFrozenRows(1);
}

function findInEventSheet(sh, phone) {
  if (!sh || sh.getLastRow() < 2) return null;
  var norm = normalizePhone(phone);
  var cols = isOldFormat(sh) ? 8 : 7;
  var phoneCol = isOldFormat(sh) ? 2 : 1; // old=C(idx2), new=B(idx1)
  var data = sh.getRange(2, 1, sh.getLastRow() - 1, cols).getValues();
  for (var i = 0; i < data.length; i++) {
    if (normalizePhone(data[i][phoneCol]) === norm) return {row: i+2, data: data[i], old: isOldFormat(sh)};
  }
  return null;
}

function checkAuth(pass) { return String(pass) === String(ADMIN_PASS); }

// ── Public endpoints ──────────────────────────────────────
function getInfo() { return {ok: true, config: getConfig()}; }

function memberLookup(phone) {
  if (!phone) return {ok: false, msg: '請提供電話'};
  var member = findMemberByPhone(phone);
  if (!member) return {ok: true, found: false};
  var evSh = getEventSheet();
  var alreadyRegistered = evSh ? !!findInEventSheet(evSh, phone) : false;
  return {ok: true, found: true, member: member, alreadyRegistered: alreadyRegistered};
}

function memberQR(phone) {
  if (!phone) return {ok: false, msg: '請提供電話'};
  var member = findMemberByPhone(phone);
  if (!member) return {ok: false, msg: '查無此電話，請先完成活動報名以建立您的會員記錄'};
  return {ok: true, name: member.name, phone: member.phone};
}

function eventRegister(p) {
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var evSh = getEventSheet();
    if (!evSh) return {ok: false, msg: '目前沒有進行中的活動，請聯繫管理員'};
    if (!p.phone) return {ok: false, msg: '請提供電話'};

    var member = findMemberByPhone(p.phone);
    if (!member) {
      if (!p.name) return {ok: false, msg: '請填寫姓名'};
      var mSh = getMemberSheet();
      if (!mSh) {
        mSh = getSpreadsheet().insertSheet(MEMBER_SHEET);
        mSh.appendRow(['姓名','電話','出生年月日','身份','加入日期','備註']);
        mSh.setFrozenRows(1);
      }
      var today = Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyy/MM/dd');
      mSh.appendRow([p.name, p.phone, p.birthday||'', p.memberType||'來賓', today, '']);
      member = {name: p.name, phone: p.phone};
    }

    var attendance = p.attendance || '出席';
    var meal       = attendance === '出席' ? (p.meal || '') : '';
    var lunchbox   = attendance === '出席' ? (p.lunchbox || '0') : '0';
    var note       = p.note || '';

    var existing = findInEventSheet(evSh, p.phone);
    if (existing) {
      evSh.getRange(existing.row, 3, 1, 4).setValues([[attendance, meal, lunchbox, note]]);
    } else {
      evSh.appendRow([member.name, member.phone, attendance, meal, lunchbox, note, '']);
    }

    return {ok: true, name: member.name, phone: member.phone, msg: '報名成功'};
  } finally { lock.releaseLock(); }
}

function eventCheckin(phone) {
  if (!phone) return {ok: false, msg: '缺少電話'};
  var lock = LockService.getScriptLock();
  lock.waitLock(8000);
  try {
    var evSh = getEventSheet();
    if (!evSh) return {ok: false, msg: '目前沒有進行中的活動'};
    var old = isOldFormat(evSh);
    var now = Utilities.formatDate(new Date(), 'Asia/Taipei', 'HH:mm:ss');

    var row = findInEventSheet(evSh, phone);
    if (!row) {
      var member = findMemberByPhone(phone);
      if (!member) return {ok: false, msg: '查無此 QR Code，請改用姓名搜尋'};
      evSh.appendRow([member.name, member.phone, '出席', '', '0', '現場報到', now]);
      return {ok: true, name: member.name, meal: '', lunchbox: '0'};
    }

    var d = row.data;
    if (old) {
      if (d[7]) return {ok: false, msg: d[1] + ' 已於 ' + d[7] + ' 報到', dup: true};
      evSh.getRange(row.row, 8).setValue(now);
      return {ok: true, name: d[1], meal: '', lunchbox: '0'};
    } else {
      if (d[6]) return {ok: false, msg: d[0] + ' 已於 ' + d[6] + ' 報到', dup: true};
      evSh.getRange(row.row, 7).setValue(now);
      return {ok: true, name: d[0], meal: d[3]||'', lunchbox: String(d[4]||'0')};
    }
  } finally { lock.releaseLock(); }
}

function eventSearch(q) {
  if (!q) return {ok: true, results: []};
  var evSh = getEventSheet();
  if (!evSh || evSh.getLastRow() < 2) return {ok: true, results: []};
  var old = isOldFormat(evSh);
  var norm = normalizePhone(q);
  var cols = old ? 8 : 7;
  var data = evSh.getRange(2, 1, evSh.getLastRow() - 1, cols).getValues();
  var results = data
    .filter(function(r) {
      return old ? (String(r[1]).indexOf(q)>=0 || normalizePhone(r[2]).indexOf(norm)>=0)
                 : (String(r[0]).indexOf(q)>=0 || normalizePhone(r[1]).indexOf(norm)>=0);
    })
    .map(function(r) {
      return old ? {name:r[1], phone:r[2], meal:'', lunchbox:'0', checkedIn:!!r[7]}
                 : {name:r[0], phone:r[1], meal:r[3], lunchbox:String(r[4]||'0'), checkedIn:!!r[6]};
    });
  return {ok: true, results: results};
}

function eventStats() {
  var evSh = getEventSheet();
  if (!evSh || evSh.getLastRow() < 2) return {ok:true, total:0, checked:0, unchecked:0};
  var old = isOldFormat(evSh);
  var cols = old ? 8 : 7;
  var data = evSh.getRange(2, 1, evSh.getLastRow()-1, cols).getValues();
  var checkedCol = old ? 7 : 6; // old=col H(idx7), new=col G(idx6)
  var checked = data.filter(function(r){return !!r[checkedCol];}).length;
  return {ok:true, total:data.length, checked:checked, unchecked:data.length-checked};
}

// ── Admin endpoints ───────────────────────────────────────
function adminLogin(pass) {
  return checkAuth(pass) ? {ok:true} : {ok:false, msg:'密碼錯誤'};
}

function adminEventGuests(pass) {
  if (!checkAuth(pass)) return {ok:false, msg:'密碼錯誤'};
  var evSh = getEventSheet();
  if (!evSh || evSh.getLastRow() < 2) return {ok:true, guests:[]};
  var old = isOldFormat(evSh);
  var cols = old ? 8 : 7;
  var data = evSh.getRange(2, 1, evSh.getLastRow()-1, cols).getValues();
  var guests = data.map(function(r) {
    if (old) {
      // A=timestamp B=姓名 C=電話 D=生日 E=? F=身份 G=出席狀況 H=報到時間(新增)
      return {name:r[1], phone:r[2], attendance:r[6]||'出席', meal:'', lunchbox:'0', note:r[5]||'', checkinTime:r[7]?String(r[7]):'', checkedIn:!!r[7]};
    }
    return {name:r[0], phone:r[1], attendance:r[2], meal:r[3], lunchbox:String(r[4]||'0'), note:r[5], checkinTime:r[6]?String(r[6]):'', checkedIn:!!r[6]};
  });
  return {ok:true, guests:guests};
}

function adminMembers(pass) {
  if (!checkAuth(pass)) return {ok:false, msg:'密碼錯誤'};
  var sh = getMemberSheet();
  if (!sh || sh.getLastRow() < 2) return {ok:true, members:[]};
  var data = sh.getRange(2, 1, sh.getLastRow()-1, 6).getValues();
  var members = data.map(function(r) {
    return {name:r[0], phone:r[1], birthday:r[2], memberType:r[3], joinDate:r[4], note:r[5]};
  });
  return {ok:true, members:members};
}

function adminCreateEvent(p) {
  if (!checkAuth(p.pass)) return {ok:false, msg:'密碼錯誤'};
  if (!p.eventName) return {ok:false, msg:'請提供活動名稱'};
  var ss = getSpreadsheet();
  var cfgSh = ss.getSheetByName(CONFIG_SHEET) || ss.insertSheet(CONFIG_SHEET);
  var keys = ['currentEvent','eventName','eventDate','eventTime','location','topic','speaker','desc','guestFee'];
  var vals = [p.eventName,p.eventName,p.eventDate||'',p.eventTime||'',p.location||'',p.topic||'',p.speaker||'',p.desc||'',p.guestFee||''];
  cfgSh.clearContents();
  keys.forEach(function(k,i){ cfgSh.getRange(i+1,1,1,2).setValues([[k,vals[i]]]); });
  var evSh = ss.getSheetByName(p.eventName);
  if (!evSh) {
    evSh = ss.insertSheet(p.eventName);
    evSh.appendRow(['姓名','電話','出席狀況','用餐','加購便當','備註','報到時間']);
    evSh.setFrozenRows(1);
  }
  return {ok:true, msg:'活動「'+p.eventName+'」已建立/更新'};
}

function adminMigrate(pass, sheetName) {
  if (!checkAuth(pass)) return {ok:false, msg:'密碼錯誤'};
  if (!sheetName) return {ok:false, msg:'請提供活動分頁名稱'};
  var ss = getSpreadsheet();
  var srcSh = ss.getSheetByName(sheetName);
  if (!srcSh) return {ok:false, msg:'找不到分頁：'+sheetName};
  if (srcSh.getLastRow() < 2) return {ok:false, msg:'該分頁沒有資料'};

  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var mSh = getMemberSheet();
    if (!mSh) {
      mSh = ss.insertSheet(MEMBER_SHEET);
      mSh.appendRow(['姓名','電話','出生年月日','身份','加入日期','備註']);
      mSh.setFrozenRows(1);
    }
    // Detect column layout: new format A=姓名 B=電話, old format A=會員編號 B=姓名 C=電話
    var header = srcSh.getRange(1,1,1,3).getValues()[0];
    var nameCol=0, phoneCol=1;
    if(String(header[0]).indexOf('編號')>=0||String(header[0]).match(/^M\d/)){nameCol=1;phoneCol=2;}

    var srcData = srcSh.getRange(2, 1, srcSh.getLastRow()-1, 3).getValues();
    var today = Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyy/MM/dd');
    var added=0, skipped=0;

    srcData.forEach(function(row) {
      var name  = String(row[nameCol]||'').trim();
      var phone = String(row[phoneCol]||'').trim();
      if (!name || !phone) return;
      if (findMemberByPhone(phone)) { skipped++; return; }
      mSh.appendRow([name, phone, '', '會員', today, '已從'+sheetName+'匯入']);
      added++;
    });
    return {ok:true, added:added, skipped:skipped};
  } finally { lock.releaseLock(); }
}
