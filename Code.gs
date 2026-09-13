var SHEET_NAME = 'RSVP';
var COLUMNS = [
  'เวลาบันทึก',
  'ชื่อของท่าน',
  'เป็นแขกฝั่งเจ้าบ่าวหรือเจ้าสาว',
  'ท่านจะมาร่วมงานหรือไม่',
  'มาด้วยกันกี่ท่าน',
  'ชื่อที่ต้องการให้ปรากฏบนการ์ด',
  'ท่านสะดวกรับการ์ดแบบใด',
  'ที่อยู่จัดส่ง',
  'เบอร์โทร',
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
  var side = String(form.side || '').trim().slice(0, 20);
  var attend = String(form.attend || '').trim().slice(0, 20);

  if (name.length < 2) {
    throw new Error('กรุณากรอกชื่อของท่าน');
  }
  if (['ฝั่งเจ้าบ่าว', 'ฝั่งเจ้าสาว'].indexOf(side) === -1) {
    throw new Error('กรุณาเลือกฝั่งแขก');
  }
  if (['สะดวกร่วมงาน', 'ไม่สะดวกร่วมงาน'].indexOf(attend) === -1) {
    throw new Error('กรุณาเลือกสถานะการเข้าร่วม');
  }

  var count = '';
  var cardname = '';
  var deliver = '';
  var addr = '';
  var tel = '';

  if (attend === 'สะดวกร่วมงาน') {
    cardname = String(form.cardname || '').trim().slice(0, 150);
    if (cardname.length < 2) {
      throw new Error('กรุณากรอกชื่อที่จะให้ปรากฏบนการ์ด');
    }

    deliver = String(form.deliver || '').trim().slice(0, 20);
    if (['ส่งไปรษณีย์', 'รับด้วยมือ', 'ไม่ต้องส่ง'].indexOf(deliver) === -1) {
      throw new Error('กรุณาเลือกวิธีรับการ์ด');
    }

    count = Math.max(1, Math.min(10, parseInt(form.count, 10) || 1));

    if (deliver === 'ส่งไปรษณีย์') {
      addr = String(form.addr || '').trim().slice(0, 500);
      tel = String(form.tel || '').trim().slice(0, 20);
      if (addr.length < 10) {
        throw new Error('กรุณากรอกที่อยู่จัดส่งให้ครบถ้วน');
      }
      if (tel.length < 9) {
        throw new Error('กรุณากรอกเบอร์โทรให้ถูกต้อง');
      }
    }
  }

  var wish = String(form.wish || '').trim().slice(0, 500);

  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var sheet = getSheet_();
    sheet.appendRow([new Date(), name, side, attend, count, cardname, deliver, addr, tel, wish]);
  } finally {
    lock.releaseLock();
  }

  return { ok: true };
}
