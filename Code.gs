// ============================================================
//  海德堡讀書會 Check-in System — Google Apps Script
//  Deploy as Web App: Execute as Me, Anyone access
// ============================================================
var SHEET_ID    = '1YoOizf3BVTyQXUim-sdJPxp7awzsPWDZrAaXJMmyq54';
var ADMIN_PASS  = 'heidelberg2026';
var MEMBER_SHEET = '會員名冊';
var CONFIG_SHEET = '系統設定';

// ── Column indices (1-based) ──────────────────────────────
// 會員名冊: A會員編號 B姓名 C電話 D出生年月日 E身份 F加入日期 G備註
// Event sheet: A會員編號 B姓名 C電話 D出席狀況 E用餐 F加購便當 G備註 H報到時間 I備註(old) J (unused)

function doGet(e) {
  var p = e.parameter || {};
  var action = p.action || '';
  var result;
  try {
    switch (action) {
      case 'info':            result = getInfo(); break;
      case 'member_lookup':   result = memberLookup(p.phone); break;
      case 'member_qr':       result = memberQR(p.phone); break;
      case 'event_register':  result = eventRegister(p); break;
      case 'event_checkin':   result = eventCheckin(p.id); break;
      case 'event_search':    result = eventSearch(p.q); break;
      case 'event_stats':     result = eventStats(); break;
      case 'admin_login':     result = adminLogin(p.pass); break;
      case 'admin_event_guests': result = adminEventGuests(p.pass); break;
      case 'admin_members':   result = adminMembers(p.pass); break;
      case 'admin_create_event': result = adminCreateEvent(p); break;
      case 'admin_migrate':   result = adminMigrate(p.pass, p.sheetName); break;
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
function getSpreadsheet() {
  return SpreadsheetApp.openById(SHEET_ID);
}

function getConfig() {
  var ss = getSpreadsheet();
  var sh = ss.getSheetByName(CONFIG_SHEET);
  if (!sh) return {};
  var data = sh.getDataRange().getValues();
  var cfg = {};
  data.forEach(function(row) { if (row[0]) cfg[row[0]] = row[1]; });
  return cfg;
}

function getEventSheet() {
  var cfg = getConfig();
  var name = cfg['currentEvent'] || '';
  if (!name) return null;
  return getSpreadsheet().getSheetByName(name);
}

function getMemberSheet() {
  return getSpreadsheet().getSheetByName(MEMBER_SHEET);
}

function genMemberId() {
  var sh = getMemberSheet();
  if (!sh || sh.getLastRow() < 2) return 'M001';
  var ids = sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues()
    .map(function(r) { return String(r[0]); })
    .filter(function(id) { return /^M\d+$/.test(id); })
    .map(function(id) { return parseInt(id.slice(1), 10); });
  var max = ids.length ? Math.max.apply(null, ids) : 0;
  return 'M' + String(max + 1).padStart(3, '0');
}

function normalizePhone(p) {
  return String(p || '').replace(/[-\s]/g, '');
}

function findMemberByPhone(phone) {
  var sh = getMemberSheet();
  if (!sh || sh.getLastRow() < 2) return null;
  var norm = normalizePhone(phone);
  var data = sh.getRange(2, 1, sh.getLastRow() - 1, 7).getValues();
  for (var i = 0; i < data.length; i++) {
    if (normalizePhone(data[i][2]) === norm) {
      return {
        row: i + 2,
        memberId: data[i][0],
        name: data[i][1],
        phone: data[i][2],
        birthday: data[i][3],
        memberType: data[i][4],
        joinDate: data[i][5],
        note: data[i][6]
      };
    }
  }
  return null;
}

function findMemberById(memberId) {
  var sh = getMemberSheet();
  if (!sh || sh.getLastRow() < 2) return null;
  var data = sh.getRange(2, 1, sh.getLastRow() - 1, 7).getValues();
  for (var i = 0; i < data.length; i++) {
    if (String(data[i][0]) === String(memberId)) {
      return {
        row: i + 2,
        memberId: data[i][0],
        name: data[i][1],
        phone: data[i][2],
        birthday: data[i][3],
        memberType: data[i][4],
        joinDate: data[i][5],
        note: data[i][6]
      };
    }
  }
  return null;
}

function findInEventSheet(sh, memberId) {
  if (!sh || sh.getLastRow() < 2) return null;
  var data = sh.getRange(2, 1, sh.getLastRow() - 1, 8).getValues();
  for (var i = 0; i < data.length; i++) {
    if (String(data[i][0]) === String(memberId)) {
      return {row: i + 2, data: data[i]};
    }
  }
  return null;
}

function checkAuth(pass) {
  return String(pass) === String(ADMIN_PASS);
}

// ── Public endpoints ──────────────────────────────────────
function getInfo() {
  var cfg = getConfig();
  return {ok: true, config: cfg};
}

function memberLookup(phone) {
  if (!phone) return {ok: false, msg: '請提供電話'};
  var member = findMemberByPhone(phone);
  if (!member) return {ok: true, found: false};
  var evSh = getEventSheet();
  var alreadyRegistered = false;
  if (evSh) {
    var inEvent = findInEventSheet(evSh, member.memberId);
    alreadyRegistered = !!inEvent;
  }
  return {ok: true, found: true, member: member, alreadyRegistered: alreadyRegistered};
}

function memberQR(phone) {
  if (!phone) return {ok: false, msg: '請提供電話'};
  var member = findMemberByPhone(phone);
  if (!member) return {ok: false, msg: '查無此電話，請先完成活動報名以建立您的會員記錄'};
  return {ok: true, memberId: member.memberId, name: member.name};
}

function eventRegister(p) {
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var evSh = getEventSheet();
    if (!evSh) return {ok: false, msg: '目前沒有進行中的活動，請聯繫管理員'};

    var member;
    if (p.memberId) {
      member = findMemberById(p.memberId);
      if (!member) return {ok: false, msg: '找不到會員 ' + p.memberId};
    } else {
      if (!p.name || !p.phone) return {ok: false, msg: '請填寫姓名和電話'};
      member = findMemberByPhone(p.phone);
      if (!member) {
        var mSh = getMemberSheet();
        var newId = genMemberId();
        var today = Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyy/MM/dd');
        mSh.appendRow([newId, p.name, p.phone, p.birthday || '', p.memberType || '來賓', today, '']);
        member = {memberId: newId, name: p.name, phone: p.phone};
      }
    }

    var attendance = p.attendance || '出席';
    var meal = attendance === '出席' ? (p.meal || '') : '';
    var lunchbox = attendance === '出席' ? (p.lunchbox || '0') : '0';
    var note = p.note || '';

    var existing = findInEventSheet(evSh, member.memberId);
    if (existing) {
      evSh.getRange(existing.row, 4, 1, 4).setValues([[attendance, meal, lunchbox, note]]);
    } else {
      evSh.appendRow([member.memberId, member.name, member.phone, attendance, meal, lunchbox, note, '']);
    }

    return {ok: true, memberId: member.memberId, name: member.name, msg: '報名成功'};
  } finally {
    lock.releaseLock();
  }
}

function eventCheckin(id) {
  if (!id) return {ok: false, msg: '缺少 ID'};
  var lock = LockService.getScriptLock();
  lock.waitLock(8000);
  try {
    var evSh = getEventSheet();
    if (!evSh) return {ok: false, msg: '目前沒有進行中的活動'};

    var row = findInEventSheet(evSh, id);
    if (!row) {
      var member = findMemberById(id);
      if (!member) return {ok: false, msg: '查無此 QR Code，請改用姓名搜尋'};
      var now = Utilities.formatDate(new Date(), 'Asia/Taipei', 'HH:mm:ss');
      evSh.appendRow([member.memberId, member.name, member.phone, '出席', '', '0', '現場報到', now]);
      return {ok: true, name: member.name, meal: '', lunchbox: '0'};
    }

    var d = row.data;
    if (d[7]) return {ok: false, msg: d[1] + ' 已於 ' + d[7] + ' 報到', dup: true};

    var now = Utilities.formatDate(new Date(), 'Asia/Taipei', 'HH:mm:ss');
    evSh.getRange(row.row, 8).setValue(now);
    return {ok: true, name: d[1], meal: d[4] || '', lunchbox: String(d[5] || '0')};
  } finally {
    lock.releaseLock();
  }
}

function eventSearch(q) {
  if (!q) return {ok: true, results: []};
  var evSh = getEventSheet();
  if (!evSh || evSh.getLastRow() < 2) return {ok: true, results: []};
  var norm = normalizePhone(q);
  var data = evSh.getRange(2, 1, evSh.getLastRow() - 1, 8).getValues();
  var results = data.filter(function(r) {
    return String(r[1]).indexOf(q) >= 0 || normalizePhone(r[2]).indexOf(norm) >= 0;
  }).map(function(r) {
    return {
      memberId: r[0], name: r[1], phone: r[2],
      meal: r[4], lunchbox: String(r[5] || '0'),
      checkedIn: !!r[7]
    };
  });
  return {ok: true, results: results};
}

function eventStats() {
  var evSh = getEventSheet();
  if (!evSh || evSh.getLastRow() < 2) return {ok: true, total: 0, checked: 0, unchecked: 0};
  var data = evSh.getRange(2, 1, evSh.getLastRow() - 1, 8).getValues();
  var checked = data.filter(function(r) { return !!r[7]; }).length;
  return {ok: true, total: data.length, checked: checked, unchecked: data.length - checked};
}

// ── Admin endpoints ───────────────────────────────────────
function adminLogin(pass) {
  if (!checkAuth(pass)) return {ok: false, msg: '密碼錯誤'};
  return {ok: true};
}

function adminEventGuests(pass) {
  if (!checkAuth(pass)) return {ok: false, msg: '密碼錯誤'};
  var evSh = getEventSheet();
  if (!evSh) return {ok: true, guests: []};
  if (evSh.getLastRow() < 2) return {ok: true, guests: []};
  var data = evSh.getRange(2, 1, evSh.getLastRow() - 1, 8).getValues();
  var guests = data.map(function(r) {
    return {
      memberId: r[0], name: r[1], phone: r[2],
      attendance: r[3], meal: r[4], lunchbox: String(r[5] || '0'),
      note: r[6], checkinTime: r[7] ? String(r[7]) : '',
      checkedIn: !!r[7]
    };
  });
  return {ok: true, guests: guests};
}

function adminMembers(pass) {
  if (!checkAuth(pass)) return {ok: false, msg: '密碼錯誤'};
  var sh = getMemberSheet();
  if (!sh || sh.getLastRow() < 2) return {ok: true, members: []};
  var data = sh.getRange(2, 1, sh.getLastRow() - 1, 7).getValues();
  var members = data.map(function(r) {
    return {
      memberId: r[0], name: r[1], phone: r[2],
      birthday: r[3], memberType: r[4], joinDate: r[5], note: r[6]
    };
  });
  return {ok: true, members: members};
}

function adminCreateEvent(p) {
  if (!checkAuth(p.pass)) return {ok: false, msg: '密碼錯誤'};
  if (!p.eventName) return {ok: false, msg: '請提供活動名稱'};

  var ss = getSpreadsheet();
  var cfgSh = ss.getSheetByName(CONFIG_SHEET);
  if (!cfgSh) {
    cfgSh = ss.insertSheet(CONFIG_SHEET);
  }

  var keys = ['currentEvent','eventName','eventDate','eventTime','location','topic','speaker','desc','guestFee'];
  var vals = [p.eventName, p.eventName, p.eventDate||'', p.eventTime||'', p.location||'', p.topic||'', p.speaker||'', p.desc||'', p.guestFee||''];
  cfgSh.clearContents();
  keys.forEach(function(k, i) { cfgSh.getRange(i+1, 1, 1, 2).setValues([[k, vals[i]]]); });

  var evSh = ss.getSheetByName(p.eventName);
  if (!evSh) {
    evSh = ss.insertSheet(p.eventName);
    evSh.appendRow(['會員編號','姓名','電話','出席狀況','用餐','加購便當','備註','報到時間']);
    evSh.setFrozenRows(1);
  }

  return {ok: true, msg: '活動「' + p.eventName + '」已建立/更新'};
}

function adminMigrate(pass, sheetName) {
  if (!checkAuth(pass)) return {ok: false, msg: '密碼錯誤'};
  if (!sheetName) return {ok: false, msg: '請提供活動分頁名稱'};

  var ss = getSpreadsheet();
  var srcSh = ss.getSheetByName(sheetName);
  if (!srcSh) return {ok: false, msg: '找不到分頁：' + sheetName};
  if (srcSh.getLastRow() < 2) return {ok: false, msg: '該分頁沒有資料'};

  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var mSh = getMemberSheet();
    if (!mSh) {
      mSh = ss.insertSheet(MEMBER_SHEET);
      mSh.appendRow(['會員編號','姓名','電話','出生年月日','身份','加入日期','備註']);
      mSh.setFrozenRows(1);
    }

    var srcData = srcSh.getRange(2, 1, srcSh.getLastRow() - 1, 8).getValues();
    var today = Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyy/MM/dd');
    var added = 0, skipped = 0;

    srcData.forEach(function(row) {
      var srcId = String(row[0] || '').trim();
      var name  = String(row[1] || '').trim();
      var phone = String(row[2] || '').trim();
      if (!name && !phone) return;

      // If row already has a member ID, check if it exists in 會員名冊
      if (srcId && /^M\d+$/.test(srcId)) {
        var exists = findMemberById(srcId);
        if (exists) { skipped++; return; }
        // ID doesn't exist yet — add it
        mSh.appendRow([srcId, name, phone, '', '會員', today, '已從'+sheetName+'匯入']);
        added++;
        return;
      }

      // No member ID — search by phone
      var existing = findMemberByPhone(phone);
      if (existing) {
        // Already in 會員名冊; update the source sheet row with the member ID
        var colA = srcSh.getRange(srcData.indexOf(row) + 2, 1);
        colA.setValue(existing.memberId);
        skipped++;
        return;
      }

      // New member — create
      var newId = genMemberId();
      mSh.appendRow([newId, name, phone, '', '會員', today, '已從'+sheetName+'匯入']);
      // Back-fill the member ID into the source event sheet
      srcSh.getRange(srcData.indexOf(row) + 2, 1).setValue(newId);
      added++;
    });

    return {ok: true, added: added, skipped: skipped};
  } finally {
    lock.releaseLock();
  }
}
