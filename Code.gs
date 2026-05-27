// ================================================================
//  ASK External Data — Google Apps Script API
//  วิธีใช้:
//  1. เปิด Google Sheets → Extensions → Apps Script
//  2. วางโค้ดนี้ทับโค้ดเดิมทั้งหมด
//  3. ถ้าเปิด Script จากใน Sheets โดยตรง → ไม่ต้องแก้อะไร
//     ถ้าเปิดเป็น Standalone Script → ใส่ ID ใน SPREADSHEET_ID
//  4. Deploy → New deployment → Web app
//     - Execute as: Me
//     - Who has access: Anyone
//  5. Copy URL ไปใส่ใน index.html
// ================================================================

const SPREADSHEET_ID = ''; // ← ถ้าเป็น Standalone Script ใส่ ID ที่นี่
                            //   ถ้าเปิดจาก Extensions → Apps Script ใน Sheets ปล่อยว่างได้เลย

// Sheet ที่ไม่ต้องการแสดงใน Web App (ถ้ามี)
const SKIP_SHEETS = ['Photos - Graphic Team'];

// คอลัมน์ใน Sheet (1-indexed)
// A=1: External ID   B=2: Project Name   C=3: Storage Availability
// D=4: Used (size)   E=5: Remark         F=6: Status   G=7: Total Capacity
const COL = {
  id: 1, project: 2, storageAvailability: 3,
  size: 4, remark: 5, status: 6, totalCapacity: 7
};

// ================================================================

function doGet(e) {
  const action = e.parameter.action || 'getAllData';
  let result;

  try {
    switch (action) {
      case 'getCategories': result = getCategories(); break;
      case 'getData':       result = getData(e.parameter.sheet); break;
      case 'getAllData':    result = getAllData(); break;
      default:             result = { success: false, error: 'Unknown action: ' + action };
    }
  } catch (err) {
    result = { success: false, error: err.message };
  }

  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  let result;

  try {
    const body = JSON.parse(e.postData.contents);
    switch (body.action) {
      case 'addExternal':    result = addExternal(body); break;
      case 'addCategory':    result = addCategory(body.name); break;
      case 'addSubProject':  result = addSubProject(body); break;
      case 'updateExternal': result = updateExternal(body); break;
      default:               result = { success: false, error: 'Unknown action: ' + body.action };
    }
  } catch (err) {
    result = { success: false, error: err.message };
  }

  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

// ================================================================
// HELPERS
// ================================================================

function getSpreadsheet() {
  if (SPREADSHEET_ID) {
    return SpreadsheetApp.openById(SPREADSHEET_ID);
  }
  return SpreadsheetApp.getActiveSpreadsheet();
}

function getCategories() {
  const ss = getSpreadsheet();
  const sheets = ss.getSheets()
    .map(s => s.getName())
    .filter(name => !SKIP_SHEETS.includes(name));
  return { success: true, data: sheets };
}

function getData(sheetName) {
  const ss = getSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) return { success: false, error: 'ไม่พบ Sheet: ' + sheetName };

  const rows = sheet.getDataRange().getValues();
  const externals = [];
  let current = null;

  // Row 0 = header, skip it
  for (let i = 1; i < rows.length; i++) {
    const [id, project, storageAvail, size, remark, status, totalCapacity] = rows[i];
    const idStr      = String(id      || '').trim();
    const projectStr = String(project || '').trim();

    if (!idStr && !projectStr) continue;

    if (idStr) {
      current = {
        id:                  idStr,
        mainProject:         projectStr,
        storageAvailability: String(storageAvail   || '').trim(),
        size:                String(size           || '').trim(),
        remark:              String(remark         || '').trim(),
        status:              String(status         || '').trim(),
        totalCapacity:       String(totalCapacity  || '').trim(),
        subProjects:         [],
        category:            sheetName
      };
      externals.push(current);
    } else if (current) {
      if (projectStr) current.subProjects.push(projectStr);
      if (!current.size && String(size || '').trim()) {
        current.size = String(size).trim();
      }
    }
  }

  return { success: true, data: externals };
}

function getAllData() {
  const cats = getCategories();
  if (!cats.success) return cats;

  const result = {};
  for (const cat of cats.data) {
    const r = getData(cat);
    result[cat] = r.success ? r.data : [];
  }

  return { success: true, data: result, categories: cats.data };
}

function addExternal(payload) {
  const { category, id, mainProject, storageAvailability, size, totalCapacity, remark, status, subProjects } = payload;
  if (!category || !id || !mainProject) {
    return { success: false, error: 'ข้อมูลไม่ครบ: ต้องมี category, id, mainProject' };
  }

  const ss = getSpreadsheet();
  const sheet = ss.getSheetByName(category);
  if (!sheet) return { success: false, error: 'ไม่พบ Sheet: ' + category };

  sheet.appendRow([id, mainProject, storageAvailability || '', size || '', remark || '', status || '', totalCapacity || '']);

  if (Array.isArray(subProjects)) {
    for (const sub of subProjects) {
      if (sub && sub.trim()) {
        sheet.appendRow(['', sub.trim(), '', '', '', '', '']);
      }
    }
  }

  return { success: true, message: 'เพิ่ม External "' + id + '" สำเร็จ' };
}

function addSubProject(payload) {
  const { category, externalId, projectName } = payload;
  if (!category || !externalId || !projectName) {
    return { success: false, error: 'ข้อมูลไม่ครบ' };
  }

  const ss = getSpreadsheet();
  const sheet = ss.getSheetByName(category);
  if (!sheet) return { success: false, error: 'ไม่พบ Sheet: ' + category };

  const rows = sheet.getDataRange().getValues();
  let extRowIndex = -1;
  let insertBefore = -1; // 1-indexed row number in Sheets

  for (let i = 1; i < rows.length; i++) {
    const id = String(rows[i][0] || '').trim();
    if (id === externalId) {
      extRowIndex = i;
    } else if (extRowIndex >= 0 && id) {
      // Found the next external drive — insert sub-project before this row
      insertBefore = i + 1; // convert to 1-indexed
      break;
    }
  }

  if (extRowIndex < 0) return { success: false, error: 'ไม่พบ External: ' + externalId };

  if (insertBefore > 0) {
    sheet.insertRowBefore(insertBefore);
    sheet.getRange(insertBefore, 1, 1, 7).setValues([['', projectName.trim(), '', '', '', '', '']]);
  } else {
    // This external is the last one — append at end
    sheet.appendRow(['', projectName.trim(), '', '', '', '', '']);
  }

  return { success: true, message: 'เพิ่ม sub-project "' + projectName + '" สำเร็จ' };
}

function updateExternal(payload) {
  const { category, id, field, value } = payload;
  if (!category || !id || !field) {
    return { success: false, error: 'ข้อมูลไม่ครบ: ต้องมี category, id, field' };
  }

  const col = COL[field];
  if (!col) return { success: false, error: 'Field ไม่ถูกต้อง: ' + field };

  const ss = getSpreadsheet();
  const sheet = ss.getSheetByName(category);
  if (!sheet) return { success: false, error: 'ไม่พบ Sheet: ' + category };

  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0] || '').trim() === id) {
      sheet.getRange(i + 1, col).setValue(value || '');
      return { success: true, message: 'อัพเดต ' + field + ' ของ "' + id + '" สำเร็จ' };
    }
  }

  return { success: false, error: 'ไม่พบ External: ' + id };
}

function addCategory(name) {
  if (!name || !name.trim()) return { success: false, error: 'ชื่อหมวดหมู่ไม่ถูกต้อง' };
  name = name.trim();

  const ss = getSpreadsheet();
  if (ss.getSheetByName(name)) return { success: false, error: 'มีหมวดหมู่ "' + name + '" อยู่แล้ว' };

  const newSheet = ss.insertSheet(name);
  newSheet.appendRow(['ID', 'Projects', 'Storage Availability', 'Used', 'Remark', 'Status', 'Total Capacity']);
  newSheet.getRange(1, 1, 1, 7).setFontWeight('bold');

  return { success: true, message: 'สร้างหมวดหมู่ "' + name + '" สำเร็จ' };
}
