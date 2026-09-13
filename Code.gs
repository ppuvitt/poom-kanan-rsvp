var SHEET_NAME = 'RSVP';
var COLUMNS = [
  'เวลาบันทึก',
  'ชื่อ-นามสกุลของท่าน',
  'ชื่อเล่น',
  'เป็นแขกฝั่งเจ้าบ่าวหรือเจ้าสาว',
  'ท่านจะมาร่วมงานหรือไม่',
  'มาด้วยกันกี่ท่าน',
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
  var deliver = '';
  var addr = '';
  var tel = '';

  if (attend === 'สะดวกร่วมงาน') {
    count = Math.max(1, Math.min(10, parseInt(form.count, 10) || 1));
    deliver = form.mail ? 'ส่งไปรษณีย์' : 'ไม่ต้องส่งไปรษณีย์';
  } else {
    deliver = String(form.deliver || '').trim().slice(0, 20);
    if (['สะดวกรับการ์ด', 'ไม่สะดวกรับการ์ด'].indexOf(deliver) === -1) {
      throw new Error('กรุณาเลือกว่าสะดวกรับการ์ดที่ระลึกไหม');
    }
  }

  // ที่อยู่/เบอร์: ต้องกรอกก็ต่อเมื่อขอให้ส่งไปรษณีย์จริง (ทั้งกลุ่มสะดวกร่วมงานและกลุ่มไม่สะดวกร่วมงานที่ยังอยากได้การ์ด)
  var wantsMail = (attend === 'สะดวกร่วมงาน' && !!form.mail) ||
                  (attend === 'ไม่สะดวกร่วมงาน' && deliver === 'สะดวกรับการ์ด' && !!form.mail);
  if (wantsMail) {
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
    sheet.appendRow([new Date(), name, nickname, side, attend, count, deliver, addr, tel, wish]);
  } finally {
    lock.releaseLock();
  }

  return { ok: true };
}
