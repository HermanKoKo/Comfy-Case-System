/**
 * ==========================================
 * 核心邏輯層 (Api.gs)
 * ==========================================
 */

// 1. 搜尋功能 (優化：使用 getDisplayValues)
function searchClient(keyword) {
  try {
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const sheet = ss.getSheetByName(CONFIG.SHEETS.CLIENT);
    if (!sheet) return [];
    const data = sheet.getDataRange().getDisplayValues(); 
    const results = [];
    const query = String(keyword).trim().toLowerCase();
    if (!query) return [];
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const id = String(row[0]).replace(/^'/, '').trim().toLowerCase();
      const name = String(row[1]).trim().toLowerCase();
      const phone = String(row[4]).replace(/^'/, '').trim().toLowerCase();
      if (id.includes(query) || name.includes(query) || phone.includes(query)) {
        results.push({
          '個案編號': row[0], '姓名': row[1], '生日': row[2], '身分證字號': row[3],
          '電話': row[4], '性別': row[5], '慢性病或特殊疾病': row[6],
          'GoogleDrive資料夾連結': row[7], '建立日期': row[8]
        });
      }
    }
    return results;
  } catch (e) { throw new Error(e.message); }
}

// 2. 通用資料儲存功能
function saveData(sheetName, dataObj) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000); 
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const targetSheetName = sheetName || CONFIG.SHEETS.CLIENT;
    const sheet = ss.getSheetByName(targetSheetName);
    if (!sheet) throw new Error("找不到工作表 [" + targetSheetName + "]");

    const rawHeaders = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const headerMap = {}; 
    rawHeaders.forEach((h, i) => headerMap[String(h).replace(/\s+/g, '').toLowerCase()] = i);
    
    let rowIndexToUpdate = -1;
    let existingRecordId = dataObj['紀錄ID'];
    
    if (existingRecordId && targetSheetName !== CONFIG.SHEETS.CLIENT) {
       const idIdx = headerMap['紀錄id'];
       if (idIdx !== undefined && sheet.getLastRow() > 1) {
         const allIds = sheet.getRange(2, idIdx + 1, sheet.getLastRow() - 1, 1).getValues().flat();
         const matchIndex = allIds.indexOf(existingRecordId);
         if (matchIndex > -1) rowIndexToUpdate = matchIndex + 2;
       }
    }

    const rowData = rawHeaders.map(rawH => {
        const cleanH = String(rawH).replace(/\s+/g, '').toLowerCase();
        let val = '';
        for (let key in dataObj) {
            if (key.replace(/\s+/g, '').toLowerCase() === cleanH) { val = dataObj[key]; break; }
        }
        if (cleanH === '紀錄id') return val || 'R' + Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), 'yyyyMMddHHmmss') + Math.floor(Math.random()*900+100);
        if (cleanH.includes('時間') || cleanH.includes('日期')) {
            if ((cleanH === '建立時間' || cleanH === '建立日期') && rowIndexToUpdate > -1) return sheet.getRange(rowIndexToUpdate, headerMap[cleanH]+1).getValue();
            return val || Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), 'yyyy-MM-dd HH:mm:ss');
        }
        if (['電話', '身分證字號', '個案編號'].includes(cleanH)) return "'" + String(val || "");
        return val || '';
    });

    if (rowIndexToUpdate > -1) {
       sheet.getRange(rowIndexToUpdate, 1, 1, rowData.length).setValues([rowData]);
       return { success: true, message: "資料已更新" };
    } else {
       sheet.appendRow(rowData);
       return { success: true, message: "資料已新增" };
    }
  } catch (e) { throw new Error(e.message); } finally { lock.releaseLock(); }
}

function saveDoctorConsultation(formObj) { return saveData(CONFIG.SHEETS.DOCTOR, formObj); }
function saveCaseTracking(formObj) { return saveData(CONFIG.SHEETS.TRACKING, formObj); }

// 3. 系統資料讀取 (醫師:A, 護理師:B, 治療師:C)
function getSystemStaff() {
  try {
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const sheet = ss.getSheetByName(CONFIG.SHEETS.SYSTEM); 
    if (!sheet) return { doctors: [], nurses: [], therapists: [] };
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return { doctors: [], nurses: [], therapists: [] };
    const data = sheet.getRange(2, 1, lastRow - 1, 3).getValues();
    return {
      doctors: data.map(r => r[0]).filter(v => v !== ""),
      nurses: data.map(r => r[1]).filter(v => v !== ""),
      therapists: data.map(r => r[2]).filter(v => v !== "")
    };
  } catch (e) { return { doctors: [], nurses: [], therapists: [] }; }
}

// 4. 歷史紀錄讀取 (通用)
function getClientHistory(clientId, sheetName) {
  try {
    if (!clientId) return [];
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) return [];
    const data = sheet.getDataRange().getDisplayValues();
    if (data.length < 2) return [];
    const headers = data[0].map(h => String(h).replace(/\s+/g, '').toLowerCase());
    const idxCaseId = headers.indexOf('個案編號');
    if (idxCaseId === -1) return [];
    const targetId = String(clientId).replace(/^'/, '').trim().toLowerCase();
    const result = [];
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][idxCaseId]).replace(/^'/, '').trim().toLowerCase() === targetId) {
        let obj = {};
        data[0].forEach((header, index) => { obj[header] = data[i][index]; });
        result.push(obj);
      }
    }
    return result.sort((a,b) => new Date(b['治療日期'] || b['日期']) - new Date(a['治療日期'] || a['日期']));
  } catch (e) { return []; }
}