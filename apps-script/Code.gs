/**
 * Bé Xem Hay – nhận đánh giá, video bị chặn, góp ý từ app và ghi vào Sheet này.
 * Gắn vào Google Sheet: Tiện ích mở rộng > Apps Script > dán mã > Triển khai > Ứng dụng web (Bất kỳ ai).
 * PHIEN_BAN để biết bản nào đang chạy.
 */
var PHIEN_BAN = 'v1';

var TABS = {
  session: ['Thời gian', 'Mã máy', 'Tuổi bé (tháng)', 'Số phút đặt', 'Phút đã xem', 'Lúc tắt', 'Con nhắc lại?', 'Tình huống', 'Số video', 'Mã video'],
  block:   ['Thời gian', 'Mã máy', 'Tuổi bé (tháng)', 'Mã video', 'Tên video'],
  feedback:['Thời gian', 'Mã máy', 'Tuổi bé (tháng)', 'Góp ý', 'Liên hệ']
};
var TAB_NAME = { session: 'Buổi xem', block: 'Video bị chặn', feedback: 'Góp ý' };

function sheet_(type) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(TAB_NAME[type]);
  if (!sh) {
    sh = ss.insertSheet(TAB_NAME[type]);
    sh.appendRow(TABS[type]);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, TABS[type].length).setFontWeight('bold');
  }
  return sh;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function clip_(s, n) { return String(s == null ? '' : s).slice(0, n); }

var MOOD = { vui: '😊 Vui vẻ', tiec: '😐 Hơi tiếc', anvai: '😫 Ăn vạ' };

function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var d = JSON.parse(e.postData.contents);
    var now = new Date();
    var dev = clip_(d.device, 20), m = Number(d.months) || '';
    if (d.type === 'session') {
      var vids = (d.videos || []).slice(0, 20).map(function (x) { return clip_(x, 15); });
      sheet_('session').appendRow([now, dev, m, Number(d.minutes) || '', Math.round((Number(d.watched) || 0) / 6) / 10,
        MOOD[d.mood] || '', d.imitate === 'yes' ? 'Có' : d.imitate === 'no' ? 'Chưa' : '', clip_(d.situation, 20), vids.length, vids.join(',')]);
    } else if (d.type === 'block') {
      sheet_('block').appendRow([now, dev, m, clip_(d.id, 15), clip_(d.title, 120)]);
    } else if (d.type === 'feedback') {
      sheet_('feedback').appendRow([now, dev, m, clip_(d.text, 2000), clip_(d.contact, 100)]);
    } else {
      return json_({ ok: false, loi: 'type lạ' });
    }
    return json_({ ok: true, phienBan: PHIEN_BAN });
  } catch (err) {
    return json_({ ok: false, loi: String(err) });
  } finally {
    lock.releaseLock();
  }
}

/** GET ?action=scores → điểm gom từ mọi bố mẹ, để app xếp hạng video. */
function doGet(e) {
  var out = { ok: true, phienBan: PHIEN_BAN, videos: {}, blocked: {}, sessions: 0 };
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var s = ss.getSheetByName(TAB_NAME.session);
  if (s && s.getLastRow() > 1) {
    var rows = s.getRange(2, 1, s.getLastRow() - 1, 10).getValues();
    out.sessions = rows.length;
    rows.forEach(function (r) {
      var mood = String(r[5]), imit = String(r[6]);
      String(r[9]).split(',').forEach(function (id) {
        if (!id) return;
        var o = out.videos[id] || (out.videos[id] = { n: 0, good: 0, bad: 0, imit: 0 });
        if (mood) { o.n++; if (mood.indexOf('Vui') >= 0) o.good++; if (mood.indexOf('Ăn vạ') >= 0) o.bad++; }
        if (imit === 'Có') o.imit++;
      });
    });
  }
  var b = ss.getSheetByName(TAB_NAME.block);
  if (b && b.getLastRow() > 1) {
    b.getRange(2, 4, b.getLastRow() - 1, 1).getValues().forEach(function (r) {
      var id = String(r[0]); if (id) out.blocked[id] = (out.blocked[id] || 0) + 1;
    });
  }
  return json_(out);
}

/** Chạy tay 1 lần trong trình soạn thảo để tạo sẵn 3 tab (không bắt buộc). */
function taoTab() { sheet_('session'); sheet_('block'); sheet_('feedback'); }
