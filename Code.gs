var SHEET_NAME = 'RSVP';
var COLUMNS = [
  'เวลาบันทึก',
  'ชื่อ-นามสกุลของท่าน',
  'ชื่อเล่น',
  'เป็นแขกฝั่งเจ้าบ่าวหรือเจ้าสาว',
  'ท่านจะมาร่วมงานหรือไม่',
  'มีผู้ติดตามมาด้วยกันกี่ท่าน',
  'ชื่อผู้ติดตามคนที่ 1',
  'ชื่อผู้ติดตามคนที่ 2',
  'ชื่อผู้ติดตามคนที่ 3',
  'ชื่อผู้ติดตามคนที่ 4',
  'ชื่อผู้ติดตามคนที่ 5',
  'อาหารที่ทานไม่ได้ / แพ้อาหาร',
  'อยากบอกอะไรบ่าวสาวไหม'
];

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Ka-nan & Poom RSVP')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover');
}

function getSheet_() {
  var id = PropertiesService.getScriptProperties().getProperty('SHEET_ID');
  if (!id) {
    throw new Error('ยังไม่ได้ตั้งค่า SHEET_ID ใน Script Properties');
  }
  var ss = SpreadsheetApp.openById(id);
  var sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
  }
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(COLUMNS);
  }
  return sheet;
}

function submitRsvp(form) {
  form = form || {};

  // honeypot: บอทมักกรอกฟิลด์ที่ซ่อนไว้ด้วย CSS แต่คนจริงมองไม่เห็นจึงไม่กรอก
  if (form.hp) {
    return { ok: true };
  }

  var name = String(form.name || '').trim().slice(0, 100);
  var nickname = String(form.nickname || '').trim().slice(0, 50);
  var side = String(form.side || '').trim().slice(0, 20);
  var attend = String(form.attend || '').trim().slice(0, 20);

  if (name.length < 2) {
    throw new Error('กรุณากรอกชื่อ-นามสกุลของท่าน');
  }
  if (['ฝั่งเจ้าบ่าว', 'ฝั่งเจ้าสาว'].indexOf(side) === -1) {
    throw new Error('กรุณาเลือกฝั่งแขก');
  }
  if (['สะดวกร่วมงาน', 'ไม่สะดวกร่วมงาน'].indexOf(attend) === -1) {
    throw new Error('กรุณาเลือกสถานะการเข้าร่วม');
  }

  var count = '';
  var companionCols = ['', '', '', '', ''];
  var allergy = '';

  if (attend === 'สะดวกร่วมงาน') {
    count = Math.max(1, Math.min(10, parseInt(form.count, 10) || 1));

    var validRelations = ['คู่สมรส/แฟน', 'เพื่อน', 'ญาติ'];
    var companionsArr = Array.isArray(form.companions) ? form.companions : [];
    companionsArr = companionsArr
      .map(function (c) {
        var n = String((c && c.name) || '').trim().slice(0, 100);
        var r = String((c && c.relation) || '').trim();
        if (validRelations.indexOf(r) === -1) r = '';
        return n ? (r ? n + ' (' + r + ')' : n) : '';
      })
      .filter(function (n) { return n; });

    // เผื่อไว้ 5 คอลัมน์ในชีต คนที่ 5 เป็นต้นไปยัดรวมกันในคอลัมน์สุดท้าย คั่นด้วย ", "
    for (var i = 0; i < 4; i++) {
      companionCols[i] = companionsArr[i] || '';
    }
    companionCols[4] = companionsArr.slice(4).join(', ').slice(0, 500);

    allergy = String(form.allergy || '').trim().slice(0, 500);
  }

  var wish = String(form.wish || '').trim().slice(0, 500);

  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var sheet = getSheet_();
    sheet.appendRow([new Date(), name, nickname, side, attend, count]
      .concat(companionCols)
      .concat([allergy, wish]));
  } finally {
    lock.releaseLock();
  }

  return { ok: true };
}
