/**
 * PR Local Purchase - Excel to CSV Converter
 * แปลงไฟล์ Excel (.xlsm) เป็น CSV สำหรับระบบ PR Local Purchase
 */

// ===================== CONFIGURATION =====================
const CONFIG = {
  // Folder IDs
  DESTINATION_FOLDER_ID: "13rrgnBkhACRWAq15e2Pr49cQbKZJ5YFQ",
  ASIA_CSV_FOLDER_ID: "1EzvuKYvXUOncf-mq3DgJnSd1t3xmH3B8",
  SIAM_CSV_FOLDER_ID: "1R7gBJUCHL7-yTLETLafD2aItzCyENYK0",

  // Sheet names
  PR_FORM_SHEET: "PR Form",
  SELECTED_DATA_SHEET: "selected_data",

  // Data structure
  HEADER_ROWS: 22,

  // Company codes that require CSV export
  VALID_COMPANY_CODES: ["ANGK", "SNGK"]
};

// ===================== MAIN FUNCTIONS =====================

/**
 * แปลงไฟล์ Excel จาก Google Drive link เป็น CSV
 * @param {string} fileLink - URL ของไฟล์ Excel ใน Google Drive
 * @returns {Object} ผลลัพธ์การแปลงไฟล์
 */
function downloadExcel(fileLink) {
  try {
    // 1. Extract file ID from URL
    const fileId = extractFileId(fileLink);
    if (!fileId) {
      return createResult(false, "ไม่สามารถดึง File ID จาก URL ได้");
    }

    // 2. Get Excel file
    const excelFile = DriveApp.getFileById(fileId);
    const newSheetName = excelFile.getName().replace(/\.xlsm$/i, "");

    // 3. Convert Excel to Google Sheets
    const resource = {
      title: newSheetName,
      mimeType: MimeType.GOOGLE_SHEETS,
      parents: [{ id: CONFIG.DESTINATION_FOLDER_ID }]
    };

    const convertedFile = Drive.Files.copy(resource, fileId, {
      supportsAllDrives: true
    });

    const gsheetId = convertedFile.id;
    Logger.log("Created Google Sheet ID: " + gsheetId);

    // 4. Export to CSV
    const csvResult = exportToCSV(gsheetId, fileId);

    // 5. Clean up - delete temporary Google Sheet
    const tempSheet = DriveApp.getFileById(gsheetId);
    tempSheet.setTrashed(true);

    // 6. Return result
    if (csvResult.success) {
      return createResult(true, "แปลงไฟล์สำเร็จ", {
        csvId: csvResult.csvId,
        csvName: csvResult.csvName
      });
    } else {
      return createResult(false, csvResult.error);
    }

  } catch (error) {
    Logger.log("Error in downloadExcel: " + error.message);
    return createResult(false, "เกิดข้อผิดพลาด: " + error.message);
  }
}

/**
 * Export ข้อมูลจาก Google Sheet เป็น CSV
 * @param {string} gsheetId - ID ของ Google Sheet
 * @param {string} excelId - ID ของไฟล์ Excel ต้นฉบับ
 * @returns {Object} ผลลัพธ์การ export
 */
function exportToCSV(gsheetId, excelId) {
  try {
    const ss = SpreadsheetApp.openById(gsheetId);

    // 1. Get PR Form sheet
    const prFormSheet = ss.getSheetByName(CONFIG.PR_FORM_SHEET);
    if (!prFormSheet) {
      return createResult(false, "ไม่พบชีต '" + CONFIG.PR_FORM_SHEET + "'");
    }

    // 2. Get or create selected_data sheet
    let selectedDataSheet = ss.getSheetByName(CONFIG.SELECTED_DATA_SHEET);
    if (!selectedDataSheet) {
      selectedDataSheet = ss.insertSheet(CONFIG.SELECTED_DATA_SHEET);
      Logger.log("สร้างชีต '" + CONFIG.SELECTED_DATA_SHEET + "' ใหม่แล้ว");
    }
    selectedDataSheet.getRange("A:Z").clearContent();

    // 3. Get data from PR Form
    const prData = prFormSheet.getDataRange().getValues();

    // 4. Find last row with data in column D
    const lastRowInD = findLastRowInColumn(prData, 3); // Column D = index 3

    if (lastRowInD <= CONFIG.HEADER_ROWS) {
      return createResult(false, "ไม่มีข้อมูลในชีต (แถวข้อมูล <= " + CONFIG.HEADER_ROWS + ")");
    }

    const dataRowCount = lastRowInD - CONFIG.HEADER_ROWS;

    // 5. Get header values
    const companyCode = prData[9][8];   // I10
    const branchCode = prData[7][8];    // I8
    const docNumber = prData[8][12];    // M9

    // 6. Process data
    const outputData = processFormData(prData, dataRowCount, companyCode, branchCode, docNumber);

    if (outputData.length === 0) {
      return createResult(false, "ไม่พบข้อมูลที่ต้องการ export");
    }

    // 7. Write to selected_data sheet
    selectedDataSheet.getRange(1, 1, outputData.length, outputData[0].length).setValues(outputData);

    // 8. Check if company code requires CSV export
    if (!CONFIG.VALID_COMPANY_CODES.includes(branchCode)) {
      return createResult(false, "รหัสบริษัท '" + branchCode + "' ไม่ต้องการ export CSV");
    }

    // 9. Create CSV file
    const csvContent = createCSVContent(outputData);
    const folderId = branchCode === "ANGK" ? CONFIG.ASIA_CSV_FOLDER_ID : CONFIG.SIAM_CSV_FOLDER_ID;
    const folder = DriveApp.getFolderById(folderId);

    const timestamp = Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), "ddMMyyyyHHmmss");
    const fileName = "PR_for_Local_Purchase_" + timestamp + "_" + docNumber + ";" + excelId + ".csv";

    const csvFile = folder.createFile(fileName, csvContent, MimeType.CSV);

    return createResult(true, "สร้าง CSV สำเร็จ", {
      csvId: csvFile.getId(),
      csvName: fileName
    });

  } catch (error) {
    Logger.log("Error in exportToCSV: " + error.message);
    return createResult(false, "เกิดข้อผิดพลาดในการ export: " + error.message);
  }
}

// ===================== HELPER FUNCTIONS =====================

/**
 * Extract File ID จาก Google Drive URL
 * รองรับหลาย format เช่น:
 * - https://drive.google.com/file/d/FILE_ID/view
 * - https://drive.google.com/open?id=FILE_ID
 * - https://docs.google.com/spreadsheets/d/FILE_ID/edit
 * @param {string} url - URL ของไฟล์
 * @returns {string|null} File ID หรือ null ถ้าไม่พบ
 */
function extractFileId(url) {
  if (!url || typeof url !== "string") {
    return null;
  }

  // Pattern สำหรับจับ File ID (25-50 ตัวอักษร)
  const patterns = [
    /\/d\/([a-zA-Z0-9_-]{25,})/,      // /d/FILE_ID/
    /id=([a-zA-Z0-9_-]{25,})/,         // ?id=FILE_ID
    /folders\/([a-zA-Z0-9_-]{25,})/,   // /folders/FILE_ID
    /([a-zA-Z0-9_-]{25,})/             // fallback: any long alphanumeric string
  ];

  for (const pattern of patterns) {
    const match = url.match(pattern);
    if (match && match[1]) {
      return match[1];
    }
  }

  return null;
}

/**
 * หาแถวสุดท้ายที่มีข้อมูลในคอลัมน์ที่กำหนด
 * @param {Array[]} data - ข้อมูลทั้งหมด
 * @param {number} columnIndex - index ของคอลัมน์ (0-based)
 * @returns {number} หมายเลขแถวสุดท้าย (1-based)
 */
function findLastRowInColumn(data, columnIndex) {
  for (let i = data.length - 1; i >= 0; i--) {
    if (data[i][columnIndex] !== "" && data[i][columnIndex] !== null && data[i][columnIndex] !== undefined) {
      return i + 1;
    }
  }
  return 0;
}

/**
 * ประมวลผลข้อมูลจาก PR Form
 * @param {Array[]} prData - ข้อมูลจาก PR Form sheet
 * @param {number} dataRowCount - จำนวนแถวข้อมูล
 * @param {string} companyCode - รหัสบริษัท (I10)
 * @param {string} branchCode - รหัสสาขา (I8)
 * @param {string} docNumber - เลขที่เอกสาร (M9)
 * @returns {Array[]} ข้อมูลที่ประมวลผลแล้ว
 */
function processFormData(prData, dataRowCount, companyCode, branchCode, docNumber) {
  const outputData = [];

  // Loop through 31 columns (days)
  for (let dayIndex = 1; dayIndex <= 31; dayIndex++) {
    const colOffset = dayIndex * 2;
    const headerColIndex = 5 + colOffset;  // Column for date header
    const dataColIndex = 4 + colOffset;    // Column for data

    // Check if columns exist
    if (headerColIndex >= prData[0].length || dataColIndex >= prData[0].length) {
      break;
    }

    const dateHeader = prData[17][headerColIndex]; // Row 18
    if (dateHeader === "" || dateHeader === null) {
      continue;
    }

    // Check header in row 22
    const headerE22 = prData[21][dataColIndex] ? String(prData[21][dataColIndex]).trim() : "";
    if (headerE22 === "-" || headerE22 === "0" || headerE22 === "") {
      continue;
    }

    // Process each data row
    for (let rowNum = 1; rowNum <= dataRowCount; rowNum++) {
      const rowIndex = rowNum + CONFIG.HEADER_ROWS - 1;
      const cellValue = prData[rowIndex][dataColIndex] ? String(prData[rowIndex][dataColIndex]).trim() : "";

      if (cellValue === "0" || cellValue === "-" || cellValue === "") {
        continue;
      }

      // Get values
      const itemCode = prData[rowIndex][3];                                    // Column D
      const extraValue = prData[rowIndex].length > 73 ? prData[rowIndex][73] : ""; // Column BV

      // Format date
      const formattedDate = formatDateValue(dateHeader);

      outputData.push([
        companyCode,                              // A: Company code (I10)
        itemCode,                                 // B: Item code (Column D)
        cellValue,                                // C: Value
        formattedDate,                            // D: Date (yyyyMMdd)
        extraValue ? String(extraValue).trim() : "", // E: Extra value (Column BV)
        branchCode,                               // F: Branch code (I8)
        "",                                       // G: (empty)
        docNumber,                                // H: Document number (M9)
        "",                                       // I: (empty)
        ""                                        // J: (empty)
      ]);
    }
  }

  return outputData;
}

/**
 * Format วันที่เป็น yyyyMMdd
 * @param {Date|string} dateValue - ค่าวันที่
 * @returns {string} วันที่ในรูปแบบ yyyyMMdd
 */
function formatDateValue(dateValue) {
  if (!dateValue) {
    return "";
  }

  try {
    if (dateValue instanceof Date) {
      return Utilities.formatDate(dateValue, Session.getScriptTimeZone(), "yyyyMMdd");
    }

    // Try to parse as date
    const parsedDate = new Date(dateValue);
    if (!isNaN(parsedDate.getTime())) {
      return Utilities.formatDate(parsedDate, Session.getScriptTimeZone(), "yyyyMMdd");
    }

    return String(dateValue);
  } catch (error) {
    return String(dateValue);
  }
}

/**
 * สร้าง CSV content จากข้อมูล
 * @param {Array[]} data - ข้อมูลที่จะแปลงเป็น CSV
 * @returns {string} CSV content
 */
function createCSVContent(data) {
  return data.map(function(row) {
    return row.map(function(cell) {
      let cellString = (cell === null || cell === undefined) ? "" : String(cell);

      // Escape special characters
      if (cellString.includes(",") || cellString.includes("\"") || cellString.includes("\n")) {
        cellString = "\"" + cellString.replace(/"/g, "\"\"") + "\"";
      }

      return cellString;
    }).join(",");
  }).join("\r\n");
}

/**
 * สร้าง result object มาตรฐาน
 * @param {boolean} success - สถานะสำเร็จ/ไม่สำเร็จ
 * @param {string} message - ข้อความ
 * @param {Object} data - ข้อมูลเพิ่มเติม (optional)
 * @returns {Object} result object
 */
function createResult(success, message, data) {
  const result = {
    success: success,
    message: message,
    timestamp: new Date().toISOString()
  };

  if (data) {
    result.data = data;
  }

  return result;
}

// ===================== LEGACY SUPPORT =====================
// ฟังก์ชันเดิมสำหรับ backward compatibility

/**
 * @deprecated ใช้ downloadExcel() แทน
 */
function download_exc(file_lnk) {
  return downloadExcel(file_lnk);
}

/**
 * @deprecated ใช้ exportToCSV() แทน
 */
function exportCSV(gid, excel_id) {
  return exportToCSV(gid, excel_id);
}
