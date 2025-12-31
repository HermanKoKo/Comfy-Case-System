/**
 * ==========================================
 * 核心邏輯層 (Controller)
 * ==========================================
 */

// ------------------------------------------
// 1. 搜尋功能
// ------------------------------------------
function searchClient(keyword) {
  try {
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const sheet = ss.getSheetByName(CONFIG.SHEETS.CLIENT);
    if (!sheet) return [];

    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) return [];

    // 搜尋建議用 getDisplayValues 以確保電話號碼 0 開頭不消失
    const data = sheet.getDataRange().getDisplayValues(); 
    const results = [];
    const query = String(keyword).trim().toLowerCase();

    if (!query) return [];

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const id = String(row[0]).replace(/^'/, '').trim().toLowerCase();
      const name = String(row[1]).trim().toLowerCase();
      const phoneRaw = String(row[4]).replace(/^'/, '').trim().toLowerCase();
      const phoneClean = phoneRaw.replace(/-/g, '');

      if (id.includes(query) || name.includes(query) || phoneRaw.includes(query) || phoneClean.includes(query)) {
        results.push({
          '個案編號': row[0], 
          '姓名': row[1], 
          '生日': row[2], 
          '身分證字號': row[3],
          '電話': row[4], 
          '性別': row[5], 
          '慢性病或特殊疾病': row[6],
          'GoogleDrive資料夾連結': row[7], 
          '建立日期': row[8]
        });
      }
    }
    return results;
  } catch (e) {
    console.error("搜尋錯誤:", e);
    throw new Error(e.message);
  }
}

// ------------------------------------------
// 2. 通用資料儲存功能
// ------------------------------------------
function saveData(sheetName, dataObj) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000); 
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const targetSheetName = sheetName || CONFIG.SHEETS.CLIENT;
    const sheet = ss.getSheetByName(targetSheetName);
    if (!sheet) throw new Error("找不到工作表 [" + targetSheetName + "]");

    const lastCol = sheet.getLastColumn() || 1;
    const rawHeaders = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    const headerMap = {}; 
    rawHeaders.forEach((h, i) => {
        const clean = String(h).replace(/\s+/g, '').toLowerCase();
        headerMap[clean] = i;
    });
    
    let rowIndexToUpdate = -1;
    let existingRecordId = dataObj['紀錄ID'];
    
    if (existingRecordId && targetSheetName !== CONFIG.SHEETS.CLIENT) {
       const idIdx = headerMap['紀錄id'];
       if (idIdx !== undefined) {
         const lastRow = sheet.getLastRow();
         if (lastRow > 1) {
           const allIds = sheet.getRange(2, idIdx + 1, lastRow - 1, 1).getValues().flat();
           const matchIndex = allIds.indexOf(existingRecordId);
           if (matchIndex > -1) rowIndexToUpdate = matchIndex + 2;
         }
       }
    }

    let generatedFolderUrl = '';
    if (targetSheetName === CONFIG.SHEETS.CLIENT && !dataObj['GoogleDrive資料夾連結'] && dataObj['個案編號'] && dataObj['姓名']) {
         try { generatedFolderUrl = createOrGetCaseFolder(CONFIG.PARENT_FOLDER_ID, dataObj['個案編號'], dataObj['姓名']); } catch (e) {}
    }

    const rowData = rawHeaders.map(rawH => {
        const cleanH = String(rawH).replace(/\s+/g, '').toLowerCase();
        let val = '';
        for (let key in dataObj) {
            if (key.replace(/\s+/g, '').toLowerCase() === cleanH) {
                val = dataObj[key];
                break;
            }
        }

        if (cleanH === '紀錄id') return val || 'R' + Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), 'yyyyMMddHHmmss') + Math.floor(Math.random()*900+100);
        
        if (cleanH.includes('時間') || cleanH.includes('日期')) {
            if ((cleanH === '建立時間' || cleanH === '建立日期') && rowIndexToUpdate > -1) {
                return sheet.getRange(rowIndexToUpdate, headerMap[cleanH]+1).getValue();
            }
            if (val) return val; 
            if (cleanH === '建立時間' || cleanH === '建立日期') return Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), 'yyyy-MM-dd HH:mm:ss');
        }

        if (cleanH === 'googledrive資料夾連結') return generatedFolderUrl || val;
        
        if (cleanH === '電話' || cleanH === '身分證字號' || cleanH === '個案編號') {
            let sVal = String(val || "");
            return sVal.startsWith("'") ? sVal : ("'" + sVal);
        }
        
        return val === undefined ? '' : val;
    });

    if (rowIndexToUpdate > -1) {
       sheet.getRange(rowIndexToUpdate, 1, 1, rowData.length).setValues([rowData]);
       return { success: true, message: "資料已更新", folderUrl: generatedFolderUrl };
    } else {
       sheet.appendRow(rowData);
       return { success: true, message: "資料已新增", folderUrl: generatedFolderUrl };
    }
  } catch (e) {
    console.error("儲存錯誤:", e);
    throw new Error(e.message);
  } finally {
    lock.releaseLock();
  }
}

function saveDoctorConsultation(formObj) { return saveData(CONFIG.SHEETS.DOCTOR, formObj); }
function saveCaseTracking(formObj) { return saveData(CONFIG.SHEETS.TRACKING, formObj); }

// ------------------------------------------
// 3. 輔助資料讀取
// ------------------------------------------

function getSystemTherapists() {
  try {
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const sheet = ss.getSheetByName(CONFIG.SHEETS.SYSTEM); 
    if (!sheet) return [];
    
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return []; 
    
    const values = sheet.getRange(2, 3, lastRow - 1, 1).getValues().flat();
    return values.filter(v => v && String(v).trim() !== "");
  } catch (e) {
    console.error(e);
    return [];
  }
}

function getClientHistory(clientId) {
  try {
    if (!clientId) return [];
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const sheet = ss.getSheetByName(CONFIG.SHEETS.TREATMENT);
    if (!sheet) return [];
    
    // 關鍵修復：使用 getDisplayValues 避免 Date 物件傳輸失敗
    const data = sheet.getDataRange().getDisplayValues();
    if (data.length < 2) return [];

    const rawHeaders = data[0];
    const map = {};
    rawHeaders.forEach((h, i) => {
        map[String(h).replace(/\s+/g, '').toLowerCase()] = i;
    });
    
    const idxId = map['紀錄id'];
    const idxCaseId = map['個案編號'];
    const idxDate = map['治療日期'];
    const idxTherapist = map['執行治療師'];
    const idxComplaint = map['當日主訴'];
    const idxContent = map['治療內容'];
    
    let idxNote = map['備註/下次治療'];
    if (idxNote === undefined) idxNote = map['備註'];

    if (idxCaseId === undefined || idxDate === undefined) return [];

    const logs = [];
    const targetId = String(clientId).replace(/^'/, '').trim().toLowerCase();

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const rowId = String(row[idxCaseId]).replace(/^'/, '').trim().toLowerCase();

      if (rowId === targetId) {
        logs.push({
          '紀錄ID': idxId !== undefined ? row[idxId] : '',
          '個案編號': row[idxCaseId],
          '治療日期': row[idxDate], // 已是字串
          '執行治療師': idxTherapist !== undefined ? row[idxTherapist] : '',
          '當日主訴': idxComplaint !== undefined ? row[idxComplaint] : '',
          '治療內容': idxContent !== undefined ? row[idxContent] : '',
          '備註/下次治療': idxNote !== undefined ? row[idxNote] : ''
        });
      }
    }
    
    // 排序
    return logs.sort((a, b) => {
       return new Date(b['治療日期']) - new Date(a['治療日期']);
    });
    
  } catch (e) {
    console.error("讀取歷史紀錄錯誤:", e);
    return [];
  }
}

// ------------------------------------------
// 4. Drive & Image 
// ------------------------------------------
function createOrGetCaseFolder(parentId, caseId, name) {
  const folderName = `${caseId}_${name}`;
  const parentFolder = DriveApp.getFolderById(parentId);
  const folders = parentFolder.getFoldersByName(folderName);
  if (folders.hasNext()) return folders.next().getUrl();
  return parentFolder.createFolder(folderName).getUrl();
}

function uploadCaseImage(data, type, name, folderUrl) {
  try {
    const folderId = extractIdFromUrl(folderUrl);
    if (!folderId) throw new Error("無效連結");
    const folder = DriveApp.getFolderById(folderId);
    const blob = Utilities.newBlob(Utilities.base64Decode(data), type, name);
    const file = folder.createFile(blob);
    return { success: true, fileUrl: file.getUrl(), thumbnailUrl: file.getThumbnail() };
  } catch (e) { return { success: false, message: e.toString() }; }
}

function getCaseImages(folderUrl) {
  try {
    const folderId = extractIdFromUrl(folderUrl);
    if (!folderId) return [];
    const folder = DriveApp.getFolderById(folderId);
    const files = folder.getFiles();
    const images = [];
    while (files.hasNext()) {
      const file = files.next();
      if (file.getMimeType().indexOf('image') > -1) {
        images.push({
          id: file.getId(), name: file.getName(), url: file.getUrl(),
          thumbnail: "https://drive.google.com/thumbnail?sz=w400&id=" + file.getId(),
          created: file.getDateCreated().getTime()
        });
      }
    }
    return images.sort((a, b) => b.created - a.created);
  } catch (e) { return []; }
}

function extractIdFromUrl(url) {
  if (!url) return null;
  const match = String(url).match(/[-\w]{25,}/);
  return match ? match[0] : null;
}