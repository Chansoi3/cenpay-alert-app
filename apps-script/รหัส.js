// ✅ Hardcode ID ของ Spreadsheet ปลายทาง
//    (ไม่ใช้ getActiveSpreadsheet() เพราะ Web App ไม่มี active sheet)
const SPREADSHEET_ID = "1qHh8Xd9pDpNCKeP8d7iWi0iHOTppxiP_vgrKttByNo4";

// ✅ เก็บอายุข้อมูลที่จะลบอัตโนมัติ (วัน)
const RETENTION_DAYS = 30;

function doPost(e) {
  let result = { status: "error", message: "ไม่พบคำสั่ง" };
  try {
    let params = JSON.parse(e.postData.contents);
    let action = params.action;

    if (action === "getAMList") result = { status: "success", data: getAMList() };
    else if (action === "getAMInfo") result = { status: "success", data: getAMInfo() };
    else if (action === "getSummary") result = { status: "success", data: getUncheckedSummary(params.amCode) };
    else if (action === "getTransactions") result = { status: "success", data: getTransactionsByDate(params.date, params.amCode) };
    else if (action === "saveResult") result = { status: "success", data: saveCheckResult(params.results, params.amLogin) };
    else if (action === "updateCheckedRow") result = { status: "success", data: updateCheckedRow(params.rowId, params.fields) };
    else if (action === "getReport") result = { status: "success", data: getReportSummary(params.date, params.amLogin) };
    else if (action === "getHistory") result = { status: "success", data: getHistory(params.amCode) };
    else if (action === "getHistoryDetail") result = { status: "success", data: getHistoryDetail(params.date, params.amCode) };
    else if (action === "getLastUpload") result = { status: "success", data: getLastUploadTime() };
    else if (action === "uploadExcelDirect") result = { status: "success", data: uploadExcelDirect(params.arrayData) };
  } catch (err) {
    result = { status: "error", message: err.toString() };
  }
  return ContentService.createTextOutput(JSON.stringify(result)).setMimeType(ContentService.MimeType.JSON);
}

function getAMList() {
  const rawSheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('RAW_Data');
  const checkedSheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('Checked_Data');
  const amSet = new Set();

  if (rawSheet && rawSheet.getLastRow() > 1) {
    const rawData = rawSheet.getDataRange().getValues();
    const areaIdx = rawData[0].indexOf('AreaCode');
    if (areaIdx !== -1) {
      for (let i = 1; i < rawData.length; i++) {
        let am = rawData[i][areaIdx] ? rawData[i][areaIdx].toString().trim().toUpperCase() : '';
        if (am && am !== 'AM_LOGIN' && am !== 'AREACODE') amSet.add(am);
      }
    }
  }

  if (checkedSheet && checkedSheet.getLastRow() > 1) {
    const checkedData = checkedSheet.getDataRange().getValues();
    for (let i = 1; i < checkedData.length; i++) {
      let am = checkedData[i][8] ? checkedData[i][8].toString().trim().toUpperCase() : '';
      if (am && am !== 'AM_LOGIN' && am !== 'AREACODE') amSet.add(am);
    }
  }

  return Array.from(amSet).sort();
}

/**
 * ดึง mapping AM code → ชื่อ-นามสกุลเต็ม จากคอลัมน์ DM
 * คืน: { AM404: "Wirakorn Pingboon", AM101: "...", ... }
 */
function getAMInfo() {
  const rawSheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('RAW_Data');
  const checkedSheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('Checked_Data');
  const amInfo = {};

  const extract = (data, dmIdx, areaIdx) => {
    if (dmIdx === -1) return;
    for (let i = 1; i < data.length; i++) {
      const dm = data[i][dmIdx] ? data[i][dmIdx].toString().trim() : '';
      if (!dm) continue;
      const parts = dm.split(/\s+/);
      if (parts.length < 2) continue;
      // parts[0] = "AM404" ส่วนที่เหลือ = ชื่อ
      const code = parts[0].toUpperCase();
      const fullName = parts.slice(1).join(' ').trim();
      // เก็บเฉพาะที่เริ่มด้วย AM และมีชื่อตาม
      if (/^AM\d+/i.test(code) && fullName && !amInfo[code]) {
        amInfo[code] = fullName;
      }
    }
  };

  if (rawSheet && rawSheet.getLastRow() > 1) {
    const rawData = rawSheet.getDataRange().getValues();
    extract(rawData, rawData[0].indexOf('DM'), rawData[0].indexOf('AreaCode'));
  }

  if (checkedSheet && checkedSheet.getLastRow() > 1) {
    const checkedData = checkedSheet.getDataRange().getValues();
    // ใน Checked_Data: DM อยู่คอลัมน์ 1 (index 1)
    extract(checkedData, 1, 8);
  }

  return amInfo;
}

/**
 * อัปเดต row เดิมใน Checked_Data ตาม rowId (row ใน sheet, เริ่มจาก 2)
 * fields = { Type, Cash_Exist, Is_Abnormal, Receiver_Emp_ID, Depositor_Emp_ID, Depositor_Name, Depositor_Position, Remark }
 * คอลัมน์ใน Checked_Data (index): 9=Type 10=Cash_Exist 11=Is_Abnormal 13=Receiver_Emp_ID 14=Depositor_Emp_ID 15=Depositor_Name 16=Depositor_Position 17=Remark
 */
function updateCheckedRow(rowId, fields) {
  if (!rowId || !fields) return { ok: false, message: 'missing rowId or fields' };
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('Checked_Data');
  const lastRow = sheet.getLastRow();
  if (rowId < 2 || rowId > lastRow) return { ok: false, message: 'rowId out of range' };

  // map field → column index (1-based สำหรับ getRange)
  const colMap = {
    Type: 10, Cash_Exist: 11, Is_Abnormal: 12,
    Receiver_Emp_ID: 14, Depositor_Emp_ID: 15,
    Depositor_Name: 16, Depositor_Position: 17, Remark: 18
  };

  Object.keys(fields).forEach(function (key) {
    if (colMap[key] !== undefined) {
      sheet.getRange(rowId, colMap[key]).setValue(fields[key]);
    }
  });

  // อัปเดต Check_Time ด้วย (col 13)
  const ts = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yy HH:mm:ss");
  sheet.getRange(rowId, 13).setValue(ts);

  return { ok: true, rowId: rowId, updatedAt: ts };
}

function formatDateString(val) {
  if (!val) return "";
  if (val instanceof Date) {
    let day = val.getDate().toString().padStart(2, '0');
    let month = (val.getMonth() + 1).toString().padStart(2, '0');
    let year = val.getFullYear().toString().slice(-2);
    return `${day}/${month}/${year}`;
  }
  let str = val.toString().trim();
  if (str.includes('GMT') || str.includes('Time') || str.length > 20) {
    let d = new Date(str);
    if (!isNaN(d.getTime())) {
      let day = d.getDate().toString().padStart(2, '0');
      let month = (d.getMonth() + 1).toString().padStart(2, '0');
      let year = d.getFullYear().toString().slice(-2);
      return `${day}/${month}/${year}`;
    }
  }
  if (/^\d{4}-\d{2}-\d{2}/.test(str)) {
    let parts = str.split('T')[0].split('-');
    let year = parts[0].slice(-2);
    return `${parts[2]}/${parts[1]}/${year}`; 
  }
  if (str.includes('/')) {
     let parts = str.split('/');
     if(parts.length === 3) {
        let p1 = parts[0].padStart(2, '0'); 
        let p2 = parts[1].padStart(2, '0'); 
        let p3 = parts[2]; 
        if (p3.length === 4) p3 = p3.slice(-2);
        return `${p1}/${p2}/${p3}`; 
     }
  }
  return str;
}

function getUncheckedSummary(amCode) {
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('RAW_Data');
  const checkedSheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('Checked_Data');
  if (sheet.getLastRow() <= 1) return [];
  const rawData = sheet.getDataRange().getValues();
  const checkedData = checkedSheet.getLastRow() > 1 ? checkedSheet.getDataRange().getValues() : [];
  const headers = rawData[0];
  const dateIdx = headers.indexOf('date');
  const areaCodeIdx = headers.indexOf('AreaCode');
  const refIdx = headers.indexOf('Ref1');
  
  const checkedKeys = new Set();
  for (let i = 1; i < checkedData.length; i++) {
    checkedKeys.add(checkedData[i][4] + '_' + formatDateString(checkedData[i][0]));
  }
  
  const summary = {};
  for (let i = 1; i < rawData.length; i++) {
    const row = rawData[i];
    const rowAreaCode = row[areaCodeIdx] ? row[areaCodeIdx].toString().trim() : '';
    const loginCode = amCode ? amCode.trim() : '';
    if (loginCode && rowAreaCode !== loginCode) continue;
    
    let dateStr = formatDateString(row[dateIdx]);
    const key = row[refIdx] + '_' + dateStr;
    
    if (dateStr && !checkedKeys.has(key)) {
      if (!summary[dateStr]) summary[dateStr] = 0;
      summary[dateStr]++;
    }
  }
  return Object.keys(summary).map(date => ({ date: date, count: summary[date] }));
}

function getTransactionsByDate(dateStr, amCode) {
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('RAW_Data');
  const checkedSheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('Checked_Data');
  const rawData = sheet.getDataRange().getValues();
  const checkedData = checkedSheet.getLastRow() > 1 ? checkedSheet.getDataRange().getValues() : [];
  const headers = rawData[0];
  const areaCodeIdx = headers.indexOf('AreaCode');
  const dateIdx = headers.indexOf('date');
  const checkedKeys = new Set();
  
  for (let i = 1; i < checkedData.length; i++) {
    checkedKeys.add(checkedData[i][4] + '_' + formatDateString(checkedData[i][0]));
  }
  
  const results = [];
  for (let i = 1; i < rawData.length; i++) {
    const row = rawData[i];
    let rowDateStr = formatDateString(row[dateIdx]);
    const rowAreaCode = row[areaCodeIdx] ? row[areaCodeIdx].toString().trim() : '';
    
    if (rowDateStr === dateStr && !checkedKeys.has(row[headers.indexOf('Ref1')] + '_' + rowDateStr) && rowAreaCode === amCode.trim()) {
      results.push({
        date: rowDateStr, DM: row[headers.indexOf('DM')], StoreCode: row[headers.indexOf('StoreCode')],
        StoreNameE: row[headers.indexOf('StoreNameE')], Ref1: row[headers.indexOf('Ref1')],
        Amount: row[headers.indexOf('Amount')], Txn: row[headers.indexOf('Txn')], BillType: row[headers.indexOf('BillType')],
        Receiver_Emp_ID: headers.indexOf('Receiver Employee ID') > -1 ? row[headers.indexOf('Receiver Employee ID')] : '',
        Depositor_Emp_ID: headers.indexOf('Depositor Employee ID') > -1 ? row[headers.indexOf('Depositor Employee ID')] : '',
        Depositor_Name: headers.indexOf('Depositor Name') > -1 ? row[headers.indexOf('Depositor Name')] : '',
        Depositor_Position: headers.indexOf('Depositor Position') > -1 ? row[headers.indexOf('Depositor Position')] : ''
      });
    }
  }
  return results;
}

function saveCheckResult(results, amLogin) {
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('Checked_Data');
  const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yy HH:mm:ss");

  // ✅ Batch write ครั้งเดียว (เร็วกว่า appendRow ทีละแถว 10x+)
  const rows = results.map(res => [
    res.date, res.DM, res.StoreCode, res.StoreNameE, res.Ref1, res.Amount, res.Txn, res.BillType, amLogin,
    res.Type, res.Cash_Exist, res.Is_Abnormal, timestamp,
    res.Receiver_Emp_ID, res.Depositor_Emp_ID, res.Depositor_Name, res.Depositor_Position,
    res.Remark
  ]);

  const lastRow = sheet.getLastRow();
  if (rows.length > 0) {
    sheet.getRange(lastRow + 1, 1, rows.length, rows[0].length).setValues(rows);
  }
  return "Success";
}

function getReportSummary(dateStr, amLogin) {
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('Checked_Data');
  if (sheet.getLastRow() <= 1) return { totalTxn: 0, abnormalCount: 0, employeeCount: 0 };
  const data = sheet.getDataRange().getValues();
  let totalTxn = 0, abnormalCount = 0, employeeCount = 0;
  for (let i = 1; i < data.length; i++) {
    let d = formatDateString(data[i][0]);
    if (d === dateStr && data[i][8] === amLogin) {
      totalTxn++;
      if (data[i][11] === 'ใช่') abnormalCount++;
      if (data[i][9] === 'พนักงาน') employeeCount++;
    }
  }
  return { totalTxn, abnormalCount, employeeCount };
}

function getHistory(amCode) {
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('Checked_Data');
  if (sheet.getLastRow() <= 1) return [];
  const data = sheet.getDataRange().getValues();
  const historySummary = {};
  for (let i = 1; i < data.length; i++) {
    if (data[i][8] === amCode) {
      let d = formatDateString(data[i][0]);
      if (!historySummary[d]) historySummary[d] = { date: d, total: 0, abnormal: 0, _sort: data[i][0] };
      historySummary[d].total++;
      if (data[i][11] === 'ใช่') historySummary[d].abnormal++;
    }
  }
  // ✅ เรียงตามวันที่จริง (ใช้ค่า Date ดิบแทน string) — ใหม่สุดก่อน
  return Object.values(historySummary)
    .sort((a, b) => {
      const da = parseRowDate_(a._sort) || new Date(0);
      const db = parseRowDate_(b._sort) || new Date(0);
      return db - da;
    })
    .map(h => ({ date: h.date, total: h.total, abnormal: h.abnormal }));
}

function getHistoryDetail(dateStr, amCode) {
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('Checked_Data');
  if (sheet.getLastRow() <= 1) return [];
  const data = sheet.getDataRange().getValues();
  const results = [];
  for (let i = 1; i < data.length; i++) {
    let d = formatDateString(data[i][0]);
    if (d === dateStr && data[i][8] === amCode) {
      results.push({
        // ✅ rowId = เลขแถวจริงใน sheet (เริ่ม 2 เพราะแถว 1 = header) — ใช้ตอน update
        rowId: i + 1,
        date: d, DM: data[i][1], StoreCode: data[i][2], StoreNameE: data[i][3], Ref1: data[i][4],
        Amount: data[i][5], Txn: data[i][6], BillType: data[i][7], AM_Login: data[i][8],
        Type: data[i][9], Cash_Exist: data[i][10], Is_Abnormal: data[i][11],
        Check_Time: data[i][12] instanceof Date ? Utilities.formatDate(data[i][12], Session.getScriptTimeZone(), "dd/MM/yy HH:mm:ss") : data[i][12],
        Receiver_Emp_ID: data[i][13] || '',
        Depositor_Emp_ID: data[i][14] || '',
        Depositor_Name: data[i][15] || '',
        Depositor_Position: data[i][16] || '',
        Remark: data[i][17] || ''
      });
    }
  }
  // ✅ เรียงตาม StoreCode + Ref1 (ภายในวันเดียวกัน)
  return results.sort((a, b) => {
    const s = String(a.StoreCode).localeCompare(String(b.StoreCode));
    if (s !== 0) return s;
    return String(a.Ref1).localeCompare(String(b.Ref1), 'th');
  });
}

function getLastUploadTime() {
  return PropertiesService.getScriptProperties().getProperty('lastUpload') || 'ยังไม่มีข้อมูลการอัปโหลด';
}

function uploadExcelDirect(arrayData) {
  try {
    const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('RAW_Data');
    if (!sheet) return "Error: ไม่พบแผ่นงานชื่อ 'RAW_Data'";

    if(arrayData && arrayData.length > 0) {
      const lastRow = sheet.getLastRow();
      if (lastRow === 0) {
        sheet.getRange(1, 1, arrayData.length, arrayData[0].length).setValues(arrayData);
      } else {
        const dataToAppend = arrayData.slice(1);
        if (dataToAppend.length > 0) {
          sheet.getRange(lastRow + 1, 1, dataToAppend.length, dataToAppend[0].length).setValues(dataToAppend);
        }
      }
    }
    
    const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yy HH:mm:ss");
    PropertiesService.getScriptProperties().setProperty('lastUpload', "อัปโหลด Manual: " + timestamp);
    return timestamp;
  } catch(e) { 
    return "Error: " + e.toString(); 
  }
}

function processDriveFiles() {
  var folderId = '1hm6ZNcHh4g0i_Nft-yDLXW9-vPtq-XyW';
  var folder = DriveApp.getFolderById(folderId);
  var files = folder.getFiles();
  var mainSheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('RAW_Data');
  if (!mainSheet) return;

  while (files.hasNext()) {
    var file = files.next();
    var fileName = file.getName().toLowerCase();

    if (fileName.indexOf('.xlsx') > -1 || fileName.indexOf('.xls') > -1) {
      try {
        var blob = file.getBlob();
        var resource = { title: file.getName() + "_temp", mimeType: MimeType.GOOGLE_SHEETS };
        var tempFile = Drive.Files.insert(resource, blob);

        var tempSpreadsheet = SpreadsheetApp.openById(tempFile.id);
        var tempSheet = tempSpreadsheet.getSheetByName('RAW') || tempSpreadsheet.getSheets()[0];
        var data = tempSheet.getDataRange().getValues();

        if (data.length > 1) {
          var maxCols = 0;
          for (var r = 0; r < data.length; r++) { if (data[r].length > maxCols) maxCols = data[r].length; }
          var cleanData = [];
          for (var r = 0; r < data.length; r++) {
            var newRow = [];
            for (var c = 0; c < maxCols; c++) { newRow.push(data[r][c] !== undefined ? data[r][c] : ""); }
            cleanData.push(newRow);
          }

          var dataToAppend = [];
          for (var rowIdx = 1; rowIdx < cleanData.length; rowIdx++) {
            if (cleanData[rowIdx].join("").trim() !== "") { dataToAppend.push(cleanData[rowIdx]); }
          }

          var lastRow = mainSheet.getLastRow();
          if (lastRow === 0 && cleanData.length > 0) {
            mainSheet.getRange(1, 1, cleanData.length, cleanData[0].length).setValues(cleanData);
          } else if (dataToAppend.length > 0) {
            mainSheet.getRange(lastRow + 1, 1, dataToAppend.length, dataToAppend[0].length).setValues(dataToAppend);
          }

          var timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yy HH:mm:ss");
          PropertiesService.getScriptProperties().setProperty('lastUpload', "ดึงจาก Google Drive อัตโนมัติ: " + timestamp);
        }

        Drive.Files.remove(tempFile.id);
        file.setTrashed(true);
      } catch (e) {
        console.error("Error processing Drive file: " + e.toString());
      }
    }
  }
}

// =====================================================
//  AUTO PURGE — ลบข้อมูลเก่ากว่า RETENTION_DAYS วัน
// =====================================================

/**
 * ลบแถวที่เก่ากว่า RETENTION_DAYS วัน ใน RAW_Data และ Checked_Data
 * อ้างอิงจากคอลัมน์ 'date' (ค้นหา header จริง ไม่ตัดสินใจจากตำแหน่งคอลัมน์)
 */
function purgeOldRecords() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const cutoff = new Date();
  cutoff.setDate(cutoff.getDate() - RETENTION_DAYS);
  cutoff.setHours(0, 0, 0, 0);

  const report = { ranAt: new Date(), cutoff: cutoff, sheets: {} };

  ['RAW_Data', 'Checked_Data'].forEach(function (sheetName) {
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) { report.sheets[sheetName] = { error: 'sheet not found' }; return; }

    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();
    if (lastRow <= 1 || lastCol === 0) {
      report.sheets[sheetName] = { kept: 0, deleted: 0, note: 'empty or header only' };
      return;
    }

    // อ่านข้อมูลทั้งหมด (header + data)
    const allData = sheet.getDataRange().getValues();
    const header = allData[0];

    // หา index ของคอลัมน์ date — ลองหลายชื่อที่อาจเป็นไปได้
    let dateColIdx = -1;
    ['date', 'Date', 'DATE', 'วันที่', 'transaction_date', 'TxnDate'].forEach(function (name) {
      if (dateColIdx === -1) {
        const idx = header.findIndex(function (h) {
          return h !== null && h !== undefined && h.toString().trim().toLowerCase() === name.toLowerCase();
        });
        if (idx !== -1) dateColIdx = idx;
      }
    });

    // fallback: ถ้าไม่เจอ ใช้คอลัมน์แรก (index 0)
    if (dateColIdx === -1) {
      dateColIdx = 0;
      console.log(sheetName + ': header ไม่มี date — ใช้คอลัมน์แรกแทน (header:', header[0], ')');
    }

    const rowsToKeep = [];  // index ใน data (เริ่มจาก 1 = data แถวแรก)
    let deletedCount = 0;
    let unparseableCount = 0;

    for (let i = 1; i < allData.length; i++) {
      const cellValue = allData[i][dateColIdx];
      const parsed = parseRowDate_(cellValue);

      if (!parsed) {
        // ถ้า parse ไม่ได้ → เก็บไว้ (ไม่ลบ) แต่นับเก็บ log
        unparseableCount++;
        rowsToKeep.push(i);
        continue;
      }

      if (parsed < cutoff) {
        deletedCount++;
      } else {
        rowsToKeep.push(i);
      }
    }

    report.sheets[sheetName] = {
      total: lastRow - 1,
      kept: rowsToKeep.length,
      deleted: deletedCount,
      unparseable: unparseableCount,
      dateCol: header[dateColIdx]
    };

    // ถ้าไม่มีอะไรจะลบ → ข้ามการเขียน
    if (deletedCount === 0) {
      console.log(sheetName + ': ไม่มี row ที่จะลบ');
      return;
    }

    // เก็บ header + rows ที่จะเก็บไว้
    const keptData = [allData[0]];
    rowsToKeep.forEach(function (idx) { keptData.push(allData[idx]); });

    // กรณีพิเศษ: ถ้า keptData มี header หลายแถว (เช่น RAW_Data ที่มี header ของ Excel ซ้อนกัน)
    // เราจะเก็บไว้ตามเดิม — ไม่ตัดสินใจแทนผู้ใช้

    // เคลียร์ sheet แล้วเขียนใหม่ทั้งหมด
    sheet.clearContents();
    sheet.getRange(1, 1, keptData.length, keptData[0].length).setValues(keptData);

    console.log(sheetName + ': deleted ' + deletedCount + ', kept ' + (keptData.length - 1));
  });

  PropertiesService.getScriptProperties().setProperty(
    'lastPurge',
    Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yy HH:mm:ss") +
      ' → ' + JSON.stringify(report.sheets)
  );

  console.log('purgeOldRecords result:', JSON.stringify(report));
  return report;
}

/**
 * แปลงค่า date ในเซลล์ (Date object หรือ string dd/MM/yy) เป็น Date object
 * คืน null ถ้า parse ไม่ได้
 */
function parseRowDate_(val) {
  if (val instanceof Date) return val;
  if (!val) return null;
  const str = val.toString().trim();
  // รูปแบบ dd/MM/yy หรือ dd/MM/yyyy
  const m = str.match(/^(\d{1,2})\/(\d{1,2})\/(\d{2,4})/);
  if (!m) {
    const d = new Date(str);
    return isNaN(d.getTime()) ? null : d;
  }
  let day = parseInt(m[1], 10);
  let month = parseInt(m[2], 10) - 1;
  let year = parseInt(m[3], 10);
  if (year < 100) year += 2000;
  const d2 = new Date(year, month, day);
  return isNaN(d2.getTime()) ? null : d2;
}

/**
 * ติดตั้ง trigger ให้ purgeOldRecords รันทุกวัน เวลา 03:00
 * รันฟังก์ชันนี้ 1 ครั้งใน Apps Script editor (หรือ setup ทั้งหมด)
 */
function setupAutoPurgeTrigger() {
  // ลบ trigger เดิมก่อน (กันซ้ำ)
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'purgeOldRecords') {
      ScriptApp.deleteTrigger(t);
    }
  });

  // สร้าง trigger รายวัน เวลา 03:00–04:00
  ScriptApp.newTrigger('purgeOldRecords')
    .timeBased()
    .everyDays(1)
    .atHour(3)
    .create();

  return 'Trigger installed: purgeOldRecords จะรันทุกวัน เวลา ~03:00';
}

/**
 * ฟังก์ชัน setup รวม — รันครั้งแรกใน editor เพื่อติดตั้ง trigger
 */
function setup() {
  const triggerMsg = setupAutoPurgeTrigger();
  console.log(triggerMsg);

  // ทดสอบ purge ครั้งแรก (dry run ไม่ได้ เพราะตัดจริง — แต่จะไม่ตัดอะไรถ้าข้อมูลทุกอย่างใหม่)
  const purgeResult = purgeOldRecords();
  console.log('Initial purge:', JSON.stringify(purgeResult));

  return { trigger: triggerMsg, purge: purgeResult };
}