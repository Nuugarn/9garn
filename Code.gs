/*
  ====== วิธีใช้งาน ======
  Sheet "Transactions" หัวตาราง (แถวที่ 1):
  id | datetime | location | type | category | title | detail | amount | imageUrl

  Sheet "Categories" หัวตาราง (แถวที่ 1):
  name

  IMPORTANT:
  - คอลัมน์ datetime เป็น Plain Text เพื่อไม่ให้ timezone ถูกเลื่อน
  - location ใช้แยกรายรับ/รายจ่ายของแต่ละสถานที่
*/

const SPREADSHEET_ID = "YOUR_SPREADSHEET_ID";
const DRIVE_FOLDER_ID = "YOUR_DRIVE_FOLDER_ID";
const SECRET_TOKEN = "CHANGE_THIS_SECRET";
const TRANSACTIONS_SHEET = "Transactions";
const CATEGORIES_SHEET = "Categories";
const BANGKOK_TIMEZONE = "Asia/Bangkok";

function doGet(e) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const txSheet = ss.getSheetByName(TRANSACTIONS_SHEET);
  const txValues = txSheet.getDataRange().getValues();
  const txHeaders = txValues.shift();

  const transactions = txValues
    .filter(row => row.some(cell => cell !== ""))
    .map(row => {
      const obj = {};
      txHeaders.forEach((h, i) => obj[h] = row[i]);

      if (obj.datetime instanceof Date) {
        obj.datetime = Utilities.formatDate(
          obj.datetime,
          BANGKOK_TIMEZONE,
          "yyyy-MM-dd'T'HH:mm:ss"
        );
      } else if (obj.datetime != null) {
        obj.datetime = String(obj.datetime);
      }

      // รองรับข้อมูลเก่าที่อาจยังไม่มี location
      obj.location = obj.location == null ? "" : String(obj.location);
      return obj;
    });

  const catSheet = ss.getSheetByName(CATEGORIES_SHEET);
  const catValues = catSheet.getDataRange().getValues();
  catValues.shift();
  const categories = catValues.map(row => row[0]).filter(Boolean);

  // ดึงรายชื่อสถานที่จาก Transactions โดยไม่ต้องสร้าง Sheet เพิ่ม
  const locations = [...new Set(
    transactions.map(t => t.location).filter(Boolean)
  )].sort((a, b) => a.localeCompare(b, "th"));

  return jsonResponse({ transactions, categories, locations });
}

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

  const location = String(body.location || "").trim();
  if (!location) {
    return jsonResponse({ ok: false, error: "missing location" });
  }

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(TRANSACTIONS_SHEET);
  let imageUrl = "";

  if (body.imageBase64) {
    imageUrl = saveImageToDrive(body.imageBase64, body.imageName || "slip.jpg");
  }

  // เก็บเวลาที่ผู้ใช้กรอกเป็น local time ของไทยตรง ๆ
  const datetime = String(body.datetime).trim();

  const id = Date.now();
  const row = sheet.getLastRow() + 1;

  // คอลัมน์ B = datetime และ C = location
  sheet.getRange(row, 2).setNumberFormat("@");
  sheet.getRange(row, 3).setNumberFormat("@");

  sheet.getRange(row, 1, 1, 9).setValues([[
    id,
    datetime,
    location,
    body.type,
    body.category || "",
    body.title,
    body.detail || "",
    Number(body.amount),
    imageUrl
  ]]);

  return jsonResponse({ ok: true, id, imageUrl, location });
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

function saveImageToDrive(base64Data, fileName) {
  const folder = DriveApp.getFolderById(DRIVE_FOLDER_ID);
  const matches = base64Data.match(/^data:(.+);base64,(.+)$/);
  const mimeType = matches ? matches[1] : "image/jpeg";
  const data = matches ? matches[2] : base64Data;
  const blob = Utilities.newBlob(
    Utilities.base64Decode(data),
    mimeType,
    fileName
  );
  const file = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return "https://drive.google.com/uc?export=view&id=" + file.getId();
}

function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
