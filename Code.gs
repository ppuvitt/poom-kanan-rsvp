var SHEET_NAME = 'RSVP';
var COLUMNS = [
  'เวลาบันทึก',
  'ชื่อของท่าน',
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
    .setTitle('Poom & Kanan RSVP')
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
  var attend = String(form.attend || '').trim().slice(0, 20);
  var cardname = String(form.cardname || '').trim().slice(0, 150);
  var deliver = String(form.deliver || '').trim().slice(0, 20);

  if (name.length < 2) {
    throw new Error('กรุณากรอกชื่อของท่าน');
  }
  if (['มาร่วมงาน', 'ยังไม่แน่ใจ', 'ไม่สะดวกมา'].indexOf(attend) === -1) {
    throw new Error('กรุณาเลือกสถานะการเข้าร่วม');
  }
  if (cardname.length < 2) {
    throw new Error('กรุณากรอกชื่อที่จะให้ปรากฏบนการ์ด');
  }
  if (['ส่งไปรษณีย์', 'รับด้วยมือ', 'ไม่ต้องส่ง'].indexOf(deliver) === -1) {
    throw new Error('กรุณาเลือกวิธีรับการ์ด');
  }

  var count = '';
  if (attend !== 'ไม่สะดวกมา') {
    count = Math.max(1, Math.min(10, parseInt(form.count, 10) || 1));
  }

  var addr = '';
  var tel = '';
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

  var wish = String(form.wish || '').trim().slice(0, 500);

  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var sheet = getSheet_();
    sheet.appendRow([new Date(), name, attend, count, cardname, deliver, addr, tel, wish]);
  } finally {
    lock.releaseLock();
  }

  return { ok: true };
}
