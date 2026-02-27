const SHEET_ID = '1OTHlZKil7YYuPcEbIDYbyd3m8qE30YI8iyzrzqUEQeg';
const SHEET_NAME = 'sheet1';

function doGet() {
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle('ระบบจองรถ Car Pool กสฟ.(น2)')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getSheet() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(["ประทับเวลา", "ประเภท", "ชื่อผู้จอง", "เบอร์โทร", "สังกัด", "งาน/สถานที่", "รถยนต์", "วันที่เริ่มต้น", "วันที่สิ้นสุด", "การตรวจสอบ"]);
  }
  return sheet;
}

// 🛡️ ฟังก์ชันรีดเอาเฉพาะ "วันที่" มาเปรียบเทียบ (แปลงเวลาเป็นเที่ยงคืนทั้งหมด 00:00:00)
function getDateOnlyTime(val) {
  if (!val) return 0;
  let d;
  if (val instanceof Date) {
    d = new Date(val.getTime());
  } else {
    // ตัดเอาเฉพาะส่วนแรกก่อนช่องว่างหรือ T
    let str = String(val).trim().split('T')[0].split(' ')[0].split(',')[0];
    if (str.includes('-')) {
      let p = str.split('-');
      d = new Date(parseInt(p[0]), parseInt(p[1])-1, parseInt(p[2]));
    } else if (str.includes('/')) {
      let p = str.split('/');
      if (p[0].length === 4) d = new Date(parseInt(p[0]), parseInt(p[1])-1, parseInt(p[2])); // YYYY/MM/DD
      else d = new Date(parseInt(p[2]), parseInt(p[1])-1, parseInt(p[0])); // DD/MM/YYYY
    } else {
      d = new Date(val);
    }
  }
  
  if (isNaN(d.getTime())) return 0;
  
  let y = d.getFullYear();
  if (y > 2500) d.setFullYear(y - 543);
  else if (y < 100) d.setFullYear(y + 2000);
  
  // บังคับเซ็ตเป็น 00:00:00 เพื่อตัดปัญหาเรื่องเวลา
  d.setHours(0, 0, 0, 0);
  return d.getTime();
}

// ฟังก์ชันจัดรูปแบบเวลาสำหรับส่งไปโชว์ให้สวยงาม (เก็บเวลาไว้)
function formatToIso(val) {
  if (!val) return "";
  let d;
  if (val instanceof Date) { d = val; }
  else {
    let str = String(val).trim().replace(',', ' ').replace(/\s+/g, ' ');
    let parts = str.split(' ');
    let dSplit = parts[0].split(/[\/-]/);
    if(dSplit.length === 3) {
       let y, m, day;
       if(dSplit[0].length === 4) { y = parseInt(dSplit[0]); m = parseInt(dSplit[1])-1; day = parseInt(dSplit[2]); }
       else { day = parseInt(dSplit[0]); m = parseInt(dSplit[1])-1; y = parseInt(dSplit[2]); }
       if(y>2500) y-=543; if(y<100) y+=2000;
       let tSplit = (parts[1] || "00:00").split(':');
       d = new Date(y, m, day, parseInt(tSplit[0]||0), parseInt(tSplit[1]||0));
    } else { d = new Date(val); }
  }
  if (isNaN(d.getTime())) return String(val);
  let pad = (n) => n < 10 ? '0' + n : n;
  return `${d.getFullYear()}-${pad(d.getMonth()+1)}-${pad(d.getDate())}T${pad(d.getHours())}:${pad(d.getMinutes())}:00`;
}

function getData() {
  const sheet = getSheet();
  const values = sheet.getDataRange().getValues(); 
  if (values.length <= 1) return [];

  values.shift(); 
  
  return values.map((row, index) => {
    return {
      rowId: index + 2,
      timestamp: row[0] instanceof Date ? formatToIso(row[0]) : row[0],
      type: row[1],
      name: row[2],
      phone: row[3],
      dept: row[4],
      job: row[5],
      car: row[6],
      start: formatToIso(row[7]), 
      end: formatToIso(row[8]),   
      check: row[9] 
    };
  }).filter(item => item.name !== "" && item.car !== ""); 
}

function saveBooking(obj) {
  const sheet = getSheet();

  // 🛡️ แปลงวันที่ใหม่เป็น 00:00:00 ของวันนั้นๆ
  const newStart = getDateOnlyTime(obj.start);
  const newEnd = getDateOnlyTime(obj.end);

  const values = sheet.getDataRange().getValues();
  let isOverlapped = false;

  // ตรวจสอบข้อมูลทั้งหมดใน Sheet
  for(let i = 1; i < values.length; i++) {
    let rowId = i + 1;
    // ข้ามถ้ากำลังแก้ไขคิวตัวเอง
    if (obj.rowId && String(rowId) === String(obj.rowId)) continue; 

    let existCar = String(values[i][6]).trim();
    let newCar = String(obj.car).trim();

    // ถ้ารถคันเดียวกัน ให้เช็ควันที่
    if (existCar === newCar) {
      let existStart = getDateOnlyTime(values[i][7]);
      let existEnd = getDateOnlyTime(values[i][8]);

      // ถ้าแปลงเวลาในอดีตไม่ได้ให้ข้าม
      if (existStart === 0 || existEnd === 0) continue;

      // 🔴 กฎชนวัน: ถ้าวันที่คาบเกี่ยวกัน = ชน 100% (ไม่สนเวลา)
      if (newStart <= existEnd && newEnd >= existStart) {
        isOverlapped = true;
        break; 
      }
    }
  }

  // ⛔ สกัดกั้นทันทีถ้าทับซ้อน (คืนค่า error ไม่อนุญาตให้บันทึก)
  if (isOverlapped) {
    return { 
      status: "error", 
      message: "ไม่สามารถจองรถคันดังกล่าวซ้ำซ้อนกัน", 
      data: getData() 
    };
  }

  // ✅ ถ้าผ่าน บันทึกปกติ
  const rowData = [
    new Date(), obj.type, obj.name, obj.phone, obj.dept, obj.job, obj.car, new Date(obj.start), new Date(obj.end), obj.check
  ];
  
  if (obj.rowId) {
    sheet.getRange(obj.rowId, 1, 1, 10).setValues([rowData]);
  } else {
    sheet.appendRow(rowData);
  }
  
  return { status: "success", data: getData() };
}

function deleteBooking(rowId) {
  const sheet = getSheet();
  sheet.deleteRow(rowId);
  return getData(); 
}
