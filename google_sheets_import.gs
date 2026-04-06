/**
 * Google Apps Script - Import "Success" rows from Source Sheet to Destination Sheet
 *
 * วิธีใช้:
 * 1. เปิด Google Sheet ปลายทาง
 * 2. ไปที่ Extensions > Apps Script
 * 3. วาง script นี้ทั้งหมด
 * 4. แก้ค่า CONFIG ด้านล่างให้ตรงกับ Sheet ต้นทางของคุณ
 * 5. รัน setupTrigger() ครั้งเดียวเพื่อตั้งเวลาทำงานอัตโนมัติ
 *    หรือรัน importSuccessRows() ด้วยตนเอง
 */

// ==================== CONFIG ====================
const CONFIG = {
  // Spreadsheet ID ของ Sheet ต้นทาง (ดูจาก URL: https://docs.google.com/spreadsheets/d/{SPREADSHEET_ID}/edit)
  SOURCE_SPREADSHEET_ID: 'YOUR_SOURCE_SPREADSHEET_ID_HERE',

  // ชื่อ sheet (tab) ใน Sheet ต้นทาง
  SOURCE_SHEET_NAME: 'Sheet1',

  // ชื่อ sheet (tab) ใน Sheet ปลายทาง (Sheet ที่ script นี้อยู่)
  DEST_SHEET_NAME: 'Sheet1',

  // คอลัมน์ที่เก็บสถานะ — G = 7 (A=1, B=2, ..., G=7)
  STATUS_COLUMN_INDEX: 7,

  // ค่าสถานะที่ต้องการคัดลอก (ตัวพิมพ์ใหญ่ตามต้นทาง)
  STATUS_VALUE: 'SUCCESS',

  // คอลัมน์ที่ใช้เป็น unique key เพื่อป้องกันข้อมูลซ้ำ — A = 1 (Doc ID)
  // ตั้งเป็น 0 ถ้าไม่ต้องการตรวจสอบซ้ำ (จะใช้การเทียบแถวทั้งแถวแทน)
  UNIQUE_KEY_COLUMN_INDEX: 1,

  // จำนวนแถว header ใน Sheet ต้นทาง (แถว 1 = ชื่อ Process, แถว 2 = ชื่อคอลัมน์)
  HEADER_ROWS: 2,

  // คัดลอก header จากต้นทางมาปลายทางด้วยหรือไม่ (เฉพาะครั้งแรกที่ปลายทางว่าง)
  COPY_HEADER: true,
};
// ================================================

/**
 * ฟังก์ชันหลัก - คัดลอกแถวที่มีสถานะ "Success" จาก Sheet ต้นทาง มายัง Sheet ปลายทาง
 * ไม่แก้ไขอะไรใน Sheet ต้นทางเลย
 */
function importSuccessRows() {
  const source = getSourceSheet_();
  const dest = getDestSheet_();

  if (!source || !dest) return;

  const sourceData = source.getDataRange().getValues();
  if (sourceData.length <= CONFIG.HEADER_ROWS) {
    Logger.log('ไม่มีข้อมูลใน Sheet ต้นทาง');
    return;
  }

  // คัดลอก header ถ้าปลายทางยังว่าง
  if (CONFIG.COPY_HEADER && CONFIG.HEADER_ROWS > 0) {
    const destData = dest.getDataRange().getValues();
    if (destData.length === 1 && destData[0].join('') === '') {
      const headers = sourceData.slice(0, CONFIG.HEADER_ROWS);
      dest.getRange(1, 1, headers.length, headers[0].length).setValues(headers);
      Logger.log('คัดลอก header เรียบร้อย');
    }
  }

  // สร้าง set ของข้อมูลที่มีอยู่แล้วในปลายทาง เพื่อป้องกันซ้ำ
  const existingKeys = getExistingKeys_(dest);

  // กรองแถวที่สถานะเป็น "Success" และยังไม่มีในปลายทาง
  const statusColIdx = CONFIG.STATUS_COLUMN_INDEX - 1; // 0-based
  const dataRows = sourceData.slice(CONFIG.HEADER_ROWS);
  const newRows = [];

  for (const row of dataRows) {
    const status = String(row[statusColIdx]).trim();
    if (status !== CONFIG.STATUS_VALUE) continue;

    const key = buildRowKey_(row);
    if (existingKeys.has(key)) continue;

    newRows.push(row);
    existingKeys.add(key);
  }

  if (newRows.length === 0) {
    Logger.log('ไม่มีแถวใหม่ที่ต้องนำเข้า');
    return;
  }

  // เขียนแถวใหม่ต่อท้าย Sheet ปลายทาง
  const lastRow = getLastDataRow_(dest);
  dest.getRange(lastRow + 1, 1, newRows.length, newRows[0].length).setValues(newRows);

  Logger.log('นำเข้าข้อมูลใหม่ ' + newRows.length + ' แถว เรียบร้อย');
}

/**
 * สร้าง key สำหรับแต่ละแถว เพื่อตรวจสอบว่าข้อมูลซ้ำหรือไม่
 */
function buildRowKey_(row) {
  if (CONFIG.UNIQUE_KEY_COLUMN_INDEX > 0) {
    return String(row[CONFIG.UNIQUE_KEY_COLUMN_INDEX - 1]).trim();
  }
  // ถ้าไม่ได้กำหนด unique key ใช้ข้อมูลทั้งแถว
  return row.map(function(cell) { return String(cell).trim(); }).join('||');
}

/**
 * ดึง set ของ key ที่มีอยู่แล้วใน Sheet ปลายทาง
 */
function getExistingKeys_(destSheet) {
  const keys = new Set();
  const data = destSheet.getDataRange().getValues();
  const startRow = CONFIG.HEADER_ROWS;

  for (let i = startRow; i < data.length; i++) {
    const key = buildRowKey_(data[i]);
    if (key && key !== '') {
      keys.add(key);
    }
  }
  return keys;
}

/**
 * หาแถวสุดท้ายที่มีข้อมูลจริง
 */
function getLastDataRow_(sheet) {
  const lastRow = sheet.getLastRow();
  return lastRow < CONFIG.HEADER_ROWS ? CONFIG.HEADER_ROWS : lastRow;
}

/**
 * เปิด Sheet ต้นทาง (read-only)
 */
function getSourceSheet_() {
  try {
    const ss = SpreadsheetApp.openById(CONFIG.SOURCE_SPREADSHEET_ID);
    const sheet = ss.getSheetByName(CONFIG.SOURCE_SHEET_NAME);
    if (!sheet) {
      Logger.log('ไม่พบ sheet ชื่อ "' + CONFIG.SOURCE_SHEET_NAME + '" ใน Sheet ต้นทาง');
      return null;
    }
    return sheet;
  } catch (e) {
    Logger.log('ไม่สามารถเปิด Sheet ต้นทางได้: ' + e.message);
    Logger.log('ตรวจสอบว่า SOURCE_SPREADSHEET_ID ถูกต้อง และ Sheet ต้นทางได้แชร์ให้กับ account นี้แล้ว');
    return null;
  }
}

/**
 * เปิด Sheet ปลายทาง (Sheet ที่ script นี้อยู่)
 */
function getDestSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.DEST_SHEET_NAME);
  if (!sheet) {
    Logger.log('ไม่พบ sheet ชื่อ "' + CONFIG.DEST_SHEET_NAME + '" ใน Sheet ปลายทาง');
    return null;
  }
  return sheet;
}

// ==================== TRIGGER SETUP ====================

/**
 * ตั้ง trigger ให้ทำงานอัตโนมัติทุก 5 นาที
 * รัน function นี้ครั้งเดียวจากเมนู Apps Script
 */
function setupTrigger() {
  // ลบ trigger เก่าก่อน (ถ้ามี)
  removeTriggers_();

  ScriptApp.newTrigger('importSuccessRows')
    .timeBased()
    .everyMinutes(5)
    .create();

  Logger.log('ตั้ง trigger ทุก 5 นาที เรียบร้อย');
}

/**
 * ตั้ง trigger ให้ทำงานอัตโนมัติทุก 1 ชั่วโมง
 */
function setupHourlyTrigger() {
  removeTriggers_();

  ScriptApp.newTrigger('importSuccessRows')
    .timeBased()
    .everyHours(1)
    .create();

  Logger.log('ตั้ง trigger ทุก 1 ชั่วโมง เรียบร้อย');
}

/**
 * ลบ trigger ทั้งหมดของ script นี้
 */
function removeTriggers_() {
  const triggers = ScriptApp.getProjectTriggers();
  for (const trigger of triggers) {
    if (trigger.getHandlerFunction() === 'importSuccessRows') {
      ScriptApp.deleteTrigger(trigger);
    }
  }
  Logger.log('ลบ trigger เก่าเรียบร้อย');
}

// ==================== MANUAL MENU ====================

/**
 * เพิ่มเมนูใน Google Sheets เมื่อเปิดไฟล์
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('📥 Import Data')
    .addItem('นำเข้าข้อมูล Success ตอนนี้', 'importSuccessRows')
    .addSeparator()
    .addItem('ตั้ง Trigger ทุก 5 นาที', 'setupTrigger')
    .addItem('ตั้ง Trigger ทุก 1 ชั่วโมง', 'setupHourlyTrigger')
    .addToUi();
}
