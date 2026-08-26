/*
  ====== วิธีใช้งาน ======
  1. เปิด Google Sheets ใหม่ 1 ไฟล์แล้วสร้าง 2 sheet:
  Sheet "Transactions" หัวตาราง (แถวที่ 1):
  id | datetime | type | category | title | detail | amount | imageUrl
  Sheet "Categories" หัวตาราง (แถวที่ 1):
  name

  2. สร้างโฟลเดอร์ใน Google Drive สำหรับเก็บภาพสลิป แล้วคัดลอก Folder ID

  3. เปิด Extensions > Apps Script ในไฟล์ Sheets นี้

  4. แก้ค่า 3 ตัวด้านล่างให้ตรงกับของคุณ:
  - SPREADSHEET_ID
  - DRIVE_FOLDER_ID
  - SECRET_TOKEN

  5. กด Deploy > New deployment > เลือกประเภท "Web app"
  - Execute as: Me
  - Who has access: Anyone
*/

const SPREADSHEET_ID = "YOUR_SPREADSHEET_ID";
const DRIVE_FOLDER_ID = "YOUR_DRIVE_FOLDER_ID";
const SECRET_TOKEN = "CHANGE_THIS_SECRET";
const TRANSACTIONS_SHEET = "Transactions";
const CATEGORIES_SHEET = "Categories";
const BANGKOK_TIMEZONE = "Asia/Bangkok";

// ดึงข้อมูลทั้งหมด (transactions + categories) ให้หน้าเว็บอ่าน
function doGet(e) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const txSheet = ss.getSheetByName(TRANSACTIONS_SHEET);
  const txValues = txSheet.getDataRange().getValues();
  const txHeaders = txValues.shift();

  const transactions = txValues
    .filter(row => row.some(cell => cell !== ""))
    .map(row => {
      const obj = {};
      txHeaders.forEach((h, i) => {
        obj[h] = row[i];
      });

      // สำคัญ: datetime ถูกเก็บเป็นข้อความ local time ของประเทศไทย
      // ไม่แปลงผ่าน Date object เพื่อป้องกัน Google Sheets/Apps Script เลื่อนเวลา
      if (obj.datetime instanceof Date) {
        obj.datetime = Utilities.formatDate(obj.datetime, BANGKOK_TIMEZONE, "yyyy-MM-dd'T'HH:mm:ss");
      } else if (obj.datetime != null) {
        obj.datetime = String(obj.datetime);
      }

      return obj;
    });

  const catSheet = ss.getSheetByName(CATEGORIES_SHEET);
  const catValues = catSheet.getDataRange().getValues();
  catValues.shift();
  const categories = catValues.map(row => row[0]).filter(Boolean);

  return jsonResponse({ transactions, categories });
}

// รับข้อมูลใหม่จาก form.html (เพิ่มรายการ) หรือ index.html (เพิ่มหมวดหมู่)
function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);

    if (body.token !== SECRET_TOKEN) {
      return jsonResponse({ ok: false, error: "invalid token" });
    }

    if (body.action === "addTransaction") {
      return addTransaction(body);
    }

    if (body.action === "addCategory") {
      return addCategory(body);
    }

    return jsonResponse({ ok: false, error: "unknown action" });
  } catch (err) {
    return jsonResponse({ ok: false, error: err.message });
  }
}

function addTransaction(body) {
  if (!body.type || !body.title || !body.amount || !body.datetime) {
    return jsonResponse({ ok: false, error: "missing required fields" });
  }

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(TRANSACTIONS_SHEET);
  let imageUrl = "";

  if (body.imageBase64) {
    imageUrl = saveImageToDrive(body.imageBase64, body.imageName || "slip.jpg");
  }

  // รับ datetime จาก form เป็น local time เช่น 2026-08-26T13:01:00
  // ห้าม new Date() / Utilities.formatDate() กับค่านี้ เพราะจะทำให้เกิด timezone shift ได้
  const datetime = String(body.datetime).trim();

  const id = Date.now();
  const row = sheet.getLastRow() + 1;

  // บังคับคอลัมน์ datetime เป็น Plain Text ก่อนเขียน
  // เพื่อไม่ให้ Google Sheets ตีความ ISO datetime แล้วแปลง timezone เอง
  sheet.getRange(row, 2).setNumberFormat("@");

  sheet.getRange(row, 1, 1, 8).setValues([[
    id,
    datetime,
    body.type,
    body.category || "",
    body.title,
    body.detail || "",
    Number(body.amount),
    imageUrl
  ]]);

  return jsonResponse({ ok: true, id, imageUrl });
}

function addCategory(body) {
  if (!body.name) {
    return jsonResponse({ ok: false, error: "missing category name" });
  }

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(CATEGORIES_SHEET);
  const existing = sheet.getDataRange().getValues().flat();

  if (existing.includes(body.name)) {
    return jsonResponse({ ok: true, alreadyExists: true });
  }

  sheet.appendRow([body.name]);
  return jsonResponse({ ok: true });
}

// แปลง base64 เป็นไฟล์เก็บลง Drive แล้วคืนลิงก์ที่เปิดดูได้แบบสาธารณะ
function saveImageToDrive(base64Data, fileName) {
  const folder = DriveApp.getFolderById(DRIVE_FOLDER_ID);
  const matches = base64Data.match(/^data:(.+);base64,(.+)$/);
  const mimeType = matches ? matches[1] : "image/jpeg";
  const data = matches ? matches[2] : base64Data;
  const blob = Utilities.newBlob(Utilities.base64Decode(data), mimeType, fileName);
  const file = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return "https://drive.google.com/uc?export=view&id=" + file.getId();
}

function jsonResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
