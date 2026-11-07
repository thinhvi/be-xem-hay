/**
 * Bé Xem Hay – nhận đánh giá, video bị chặn, góp ý từ app và ghi vào Sheet này.
 * v2: mỗi tuần tự quét video mới của các kênh đã duyệt, đưa vào tab "Chờ duyệt".
 *     Bố mẹ/người duyệt đổi cột Trạng thái sang "Duyệt" là app tự lấy về.
 * Gắn vào Google Sheet: Tiện ích mở rộng > Apps Script > dán mã > Triển khai > Ứng dụng web (Bất kỳ ai).
 */
var PHIEN_BAN = 'v2';

var TABS = {
  session: ['Thời gian', 'Mã máy', 'Tuổi bé (tháng)', 'Số phút đặt', 'Phút đã xem', 'Lúc tắt', 'Con nhắc lại?', 'Tình huống', 'Số video', 'Mã video'],
  block:   ['Thời gian', 'Mã máy', 'Tuổi bé (tháng)', 'Mã video', 'Tên video'],
  feedback:['Thời gian', 'Mã máy', 'Tuổi bé (tháng)', 'Góp ý', 'Liên hệ'],
  pending: ['Ngày thấy', 'Trạng thái', 'Mã video', 'Link xem', 'Tên gốc', 'Kênh', 'Phút', 'Tên tiếng Việt',
            'Mục (mã)', 'Tuổi từ', 'Tuổi đến', 'Nhịp', 'Hỏi – chờ con', 'Chủ đề tiếng Anh (mã)', 'Chủ đề lớp (mã)', 'Câu hỏi cho bố mẹ']
};
var TAB_NAME = { session: 'Buổi xem', block: 'Video bị chặn', feedback: 'Góp ý', pending: 'Chờ duyệt' };
var TRANG_THAI = ['Chờ duyệt', 'Duyệt', 'Bỏ'];

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

/** GET ?action=scores → điểm gom từ mọi bố mẹ + các video đã duyệt thêm. */
function doGet(e) {
  var out = { ok: true, phienBan: PHIEN_BAN, videos: {}, blocked: {}, sessions: 0, approved: [] };
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
  out.approved = approved_();
  return json_(out);
}

var codes_ = function (s) { return String(s || '').split(/[,\s]+/).filter(String); };

/** Các dòng "Duyệt" trong tab Chờ duyệt → định dạng video của app. */
function approved_() {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(TAB_NAME.pending);
  if (!sh || sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, TABS.pending.length).getValues()
    .filter(function (r) { return r[1] === 'Duyệt' && r[2]; })
    .map(function (r) {
      return {
        id: String(r[2]), t: String(r[7] || r[4]), ch: String(r[5]), cat: codes_(r[8]),
        age: [Number(r[9]) || 2, Number(r[10]) || 6], pace: r[11] === 'vua' ? 'vua' : 'cham',
        inter: String(r[12]) === 'Có', sit: [], ent: codes_(r[13]), lop: codes_(r[14]),
        ask: String(r[15] || ''), dur: Math.round((Number(r[6]) || 3) * 60)
      };
    });
}

/* ================== Quét video mới hằng tuần ================== */

// Kênh đã duyệt: mục mặc định, tuổi, nhịp, có hỏi – chờ con trả lời không
var KENH = [
  ['UCLsooMJoIpl_7ux2jvdPB-Q', 'Super Simple Songs', 'en', 2, 5, 'cham', true],
  ['UCxG6Tbopv4XcHfem65bgFeg', 'Super Simple Play with Caitie', 'en', 2, 5, 'cham', true],
  ['UCNR06PEAvBj_siYZCBtk1hw', 'Super Simple Storytime', 'story,en', 2, 5, 'cham', true],
  ['UCdHK6g8ddMEcMu9B3LgeKYg', 'Maple Leaf Learning', 'en', 2, 5, 'cham', true],
  ['UC6LKuH7RPkvRmzS9-8URtqA', 'Dream English Kids', 'en', 2, 6, 'cham', true],
  ['UCERmUv0brIpW1JEBbEtNGdw', "The Kids' Picture Show", 'en,explore', 3, 6, 'cham', false],
  ['UCXRPWm72nrs5sK7fNTSXZrw', 'Little Fox', 'en,story', 3, 6, 'cham', false],
  ['UCGfA6UdZEmpIwBZ4btAKkgA', 'Vooks', 'story', 3, 6, 'cham', false],
  ['UCDqgSnRMGVx3dP4sn3ATZMA', 'Daniel Tiger', 'feel', 2, 5, 'cham', true],
  ['UCoookXUzPciGrEZEXmh4Jjg', 'Sesame Street', 'feel', 2, 5, 'cham', true],
  ['UCj_mFUb-47d9QNiJ5556LjQ', 'Hey Duggee', 'story', 2, 6, 'vua', false],
  ['UCVzLLZkDuFGAE2BGdBuBNBg', 'Bluey', 'story,feel', 3, 7, 'vua', false],
  ['UCnBdzaRy-Ky9Vh54XJlFz1Q', 'Storyline Online', 'story', 4, 7, 'cham', false],
  ['UCPlwvN0w4qFSP1FllALB92w', 'Numberblocks', 'math', 3, 6, 'vua', false],
  ['UC_qs3c0ehDvZkbiEbOj6Drg', 'Alphablocks', 'en', 4, 6, 'vua', false],
  ['UC5uIZ2KOZZeQDQo_Gsi_qbQ', 'Cosmic Kids Yoga', 'move', 3, 6, 'cham', true],
  ['UCRFIPG2u1DxKLNuE3y2SjHA', 'SciShow Kids', 'explore', 3, 6, 'cham', false],
  ['UC5XMF3Inoi8R9nSI8ChOsdQ', 'Art for Kids Hub', 'story', 3, 6, 'cham', false]
];

// Bỏ bản tổng hợp dài, trailer, livestream…
var LOAI = /compilation|hour|\bmix\b|non.?stop|collection|best of|top \d+|\+ ?more|marathon|live|trailer|teaser|behind the scenes|unboxing|full movie|#shorts|\bshorts\b|episodes? \d+ ?- ?\d+/i;

// Đoán chủ đề tiếng Anh từ tên video
var TU_KHOA = [
  ['mau', /colou?r|rainbow/i], ['convat', /animal|farm|zoo|bug|insect|fish|dog|cat\b|bird|duck|frog|octopus|shark|bear/i],
  ['doan', /fruit|vegetable|food|eat|hungry|apple|banana|snack|lunch|breakfast|dinner|cook/i],
  ['quanao', /cloth|wear|shoes|dress|hat|socks|boots|coat/i],
  ['thoitiet', /weather|rain|sunny|snow|wind|season|spring|summer|autumn|fall\b|winter|days of the week|months/i],
  ['xe', /car\b|bus\b|train|truck|vehicle|boat|plane|bike|tractor|digger|fire engine/i],
  ['giadinh', /family|mommy|mummy|daddy|grandma|grandpa|baby brother|sister|brother/i],
  ['camxuc', /feel|happy|sad|angry|emotion|scared|calm/i],
  ['hangngay', /wash|brush|bath|morning|bedtime|sleep|clean up|tidy|potty|get dressed|routine/i],
  ['chao', /hello|goodbye|please|thank|sorry|birthday|manners/i],
  ['so', /count|number|\b1 to 10\b|how many/i], ['hinh', /shape|circle|square|triangle/i],
  ['chu', /abc|alphabet|letter|phonics/i], ['cothe', /body|head|face|nose|toes|hands|eyes/i],
  ['nha', /house|home|room|toy/i], ['dongtu', /big|small|open|shut|in on under|walk|jump|run|stop|go\b|fast|slow/i],
  ['noichon', /job|doctor|teacher|firefighter|police|playground|park|community|school|shop|store/i]
];
var LOP_THEO_CHU_DE = { convat: 'dongvat', doan: 'thucvat', xe: 'giaothong', giadinh: 'giadinh', noichon: 'nghe',
  thoitiet: 'tunhien', cothe: 'banthan', camxuc: 'banthan', hangngay: 'banthan', chu: 'truong', so: 'truong', chao: 'truong' };
var MUC_THEO_CHU_DE = { camxuc: 'feel', hangngay: 'habit', so: 'math', hinh: 'math', dongtu: 'move' };
var CAU_HOI = {
  mau: 'What color is it? Chỉ đồ trong nhà và hỏi con màu gì.', convat: 'Con bắt chước tiếng con vật trong video nhé.',
  doan: 'Bữa ăn tới, gọi tên món ăn bằng tiếng Anh cùng con.', quanao: 'Lúc thay đồ, gọi tên từng món bằng tiếng Anh.',
  thoitiet: "How's the weather? Nhìn ra cửa sổ cùng con.", xe: 'Ra đường đếm xe và gọi tên cùng con.',
  giadinh: 'Chỉ ảnh gia đình và gọi tên từng người.', camxuc: 'Hôm nay con thấy thế nào? Happy hay sad?',
  hangngay: 'Hát lại bài này lúc làm việc đó ở nhà.', chao: 'Tập nói câu chào/cảm ơn này với con hôm nay.',
  so: 'Đếm đồ chơi bằng tiếng Anh cùng con.', hinh: 'Tìm đồ vật hình tròn, vuông quanh nhà.', chu: 'Chữ đầu tên con là chữ gì?',
  cothe: 'Where is your nose? Chơi chỉ bộ phận cơ thể.', nha: 'Đi quanh nhà gọi tên đồ vật bằng tiếng Anh.',
  dongtu: 'Chơi làm theo: jump, run, stop!', noichon: 'Lớn lên con muốn làm gì?'
};

function giaiMa_(s) { try { return JSON.parse('"' + s + '"'); } catch (e) { return s; } }

/** PT1H2M3S → giây */
function isoGiay_(s) { var m = /PT(?:(\d+)H)?(?:(\d+)M)?(?:(\d+)S)?/.exec(s || '') || []; return (+m[1] || 0) * 3600 + (+m[2] || 0) * 60 + (+m[3] || 0); }

/** 30 video mới nhất của kênh qua YouTube Data API (dịch vụ nâng cao "YouTube" trong dự án).
 *  Chỉ giữ video nhúng được, không phải livestream, và kênh đã gắn nhãn "dành cho trẻ em". */
function videoCuaKenh_(channelId) {
  var list = YouTube.PlaylistItems.list('contentDetails', { playlistId: 'UU' + channelId.slice(2), maxResults: 30 });
  var ids = (list.items || []).map(function (i) { return i.contentDetails.videoId; });
  if (!ids.length) return [];
  var vs = YouTube.Videos.list('snippet,contentDetails,status', { id: ids.join(',') });
  return (vs.items || []).filter(function (x) {
    return x.status && x.status.embeddable !== false && x.status.madeForKids !== false &&
      x.snippet.liveBroadcastContent === 'none';
  }).map(function (x) { return { id: x.id, title: x.snippet.title, dur: isoGiay_(x.contentDetails.duration) }; });
}

/** Mã video đã có trong app hoặc đã nằm trong Chờ duyệt. */
function daBiet_() {
  var known = {};
  try {
    var js = UrlFetchApp.fetch('https://thinhvi.github.io/be-xem-hay/data.js', { muteHttpExceptions: true }).getContentText();
    (js.match(/"id": ?"[^"]+"/g) || []).forEach(function (m) { known[m.replace(/"id": ?"|"/g, '')] = 1; });
  } catch (e) {}
  var sh = sheet_('pending');
  if (sh.getLastRow() > 1) sh.getRange(2, 3, sh.getLastRow() - 1, 1).getValues().forEach(function (r) { known[String(r[0])] = 1; });
  return known;
}

/** Chạy mỗi tuần (hoặc bấm Chạy tay): thêm video mới hợp chuẩn vào tab Chờ duyệt. */
function quetVideoMoi() {
  var known = daBiet_(), sh = sheet_('pending'), rows = [], now = new Date();
  KENH.forEach(function (k) {
    var dem = 0;
    try {
      videoCuaKenh_(k[0]).forEach(function (v) {
        if (dem >= 6 || known[v.id] || LOAI.test(v.title) || v.dur < 60 || v.dur > 900) return;
        known[v.id] = 1; dem++;
        var cat = codes_(k[2]), ent = [], lop = [];
        TU_KHOA.forEach(function (tk) { if (tk[1].test(v.title)) ent.push(tk[0]); });
        ent = ent.slice(0, 2);
        ent.forEach(function (e) {
          if (LOP_THEO_CHU_DE[e] && lop.indexOf(LOP_THEO_CHU_DE[e]) < 0) lop.push(LOP_THEO_CHU_DE[e]);
          if (MUC_THEO_CHU_DE[e] && cat.indexOf(MUC_THEO_CHU_DE[e]) < 0) cat.push(MUC_THEO_CHU_DE[e]);
        });
        var vi = v.title;
        try { vi = LanguageApp.translate(v.title.replace(/\|.*$/, '').trim(), 'en', 'vi'); } catch (e) {}
        rows.push([now, 'Chờ duyệt', v.id, 'https://youtu.be/' + v.id, v.title, k[1], Math.round(v.dur / 6) / 10, vi,
          cat.join(','), k[3], k[4], k[5], k[6] ? 'Có' : 'Không', cat.indexOf('en') >= 0 ? ent.join(',') : '', lop.join(','),
          ent.length ? CAU_HOI[ent[0]] : '']);
      });
    } catch (e) { Logger.log(k[1] + ': ' + e); }
  });
  if (rows.length) {
    var start = sh.getLastRow() + 1;
    sh.getRange(start, 1, rows.length, rows[0].length).setValues(rows);
    sh.getRange(start, 2, rows.length, 1).setDataValidation(
      SpreadsheetApp.newDataValidation().requireValueInList(TRANG_THAI, true).build());
  }
  Logger.log('Thêm ' + rows.length + ' video vào Chờ duyệt');
  return rows.length;
}

/** Bấm Chạy 1 lần: đặt lịch quét sáng Chủ nhật hằng tuần + tạo tab hướng dẫn. */
function batLichQuet() {
  ScriptApp.getProjectTriggers().forEach(function (t) { if (t.getHandlerFunction() === 'quetVideoMoi') ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('quetVideoMoi').timeBased().onWeekDay(ScriptApp.WeekDay.SUNDAY).atHour(7).create();
  huongDan_();
  return quetVideoMoi();
}

function huongDan_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet(), sh = ss.getSheetByName('Hướng dẫn duyệt') || ss.insertSheet('Hướng dẫn duyệt');
  sh.clear();
  var r = [
    ['Cách duyệt video mới', ''],
    ['1', 'Mỗi sáng Chủ nhật máy tự thêm video mới của các kênh đã duyệt vào tab "Chờ duyệt".'],
    ['2', 'Bấm Link xem, xem lướt: nhịp có chậm không, có la hét/doạ/quảng cáo không, hợp tuổi không.'],
    ['3', 'Sửa Tên tiếng Việt, tuổi, câu hỏi nếu cần. Rồi đổi Trạng thái sang "Duyệt" (hoặc "Bỏ").'],
    ['4', 'App tự lấy video đã Duyệt khi bố mẹ mở app, không cần sửa gì thêm.'],
    ['', ''],
    ['Mã Mục', 'en = Tiếng Anh, math = Toán, explore = Khám phá, feel = Cảm xúc, habit = Thói quen, move = Vận động, story = Truyện'],
    ['Mã Nhịp', 'cham = nhịp chậm, vua = nhịp vừa (video nhịp nhanh thì Bỏ)'],
    ['Mã Chủ đề tiếng Anh', 'chao, hangngay, giadinh, cothe, camxuc, mau, so, convat, doan, quanao, thoitiet, xe, nha, dongtu, noichon, hinh, chu'],
    ['Mã Chủ đề lớp', 'truong, banthan, giadinh, nghe, dongvat, thucvat, giaothong, tunhien']
  ];
  sh.getRange(1, 1, r.length, 2).setValues(r);
  sh.getRange(1, 1).setFontWeight('bold').setFontSize(14);
  sh.setColumnWidth(1, 160); sh.setColumnWidth(2, 720);
}

/** Chạy tay 1 lần trong trình soạn thảo để tạo sẵn các tab (không bắt buộc). */
function taoTab() { sheet_('session'); sheet_('block'); sheet_('feedback'); sheet_('pending'); }
