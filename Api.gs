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

    const data = sheet.getDataRange().getValues();
    const results = [];
    const query = String(keyword).trim().toLowerCase();

    if (!query) return [];

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      // 資料清洗
      const id = String(row[0]).replace(/^'/, '').trim().toLowerCase();
      const name = String(row[1]).trim().toLowerCase();
      const phoneRaw = String(row[4]).replace(/^'/, '').trim().toLowerCase();
      const phoneClean = phoneRaw.replace(/-/g, '');

      if (id.includes(query) || name.includes(query) || phoneRaw.includes(query) || phoneClean.includes(query)) {
        let dob = row[2];
        if (dob instanceof Date) dob = Utilities.formatDate(dob, ss.getSpreadsheetTimeZone(), 'yyyy-MM-dd');
        
        let createdDate = row[8];
        if (createdDate instanceof Date) createdDate = Utilities.formatDate(createdDate, ss.getSpreadsheetTimeZone(), 'yyyy-MM-dd HH:mm:ss');

        results.push({
          '個案編號': row[0], '姓名': row[1], '生日': dob, '身分證字號': row[3],
          '電話': row[4], '性別': row[5], '慢性病或特殊疾病': row[6],
          'GoogleDrive資料夾連結': row[7], '建立日期': createdDate
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

    // 取得標題列並建立 "標準化標題 -> 欄位索引" 的對照表
    const rawHeaders = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const headerMap = {}; // Key: CleanHeader, Value: OriginalHeaderIndex
    rawHeaders.forEach((h, i) => {
        // 標準化：轉字串 -> 去空白 -> 轉小寫
        const clean = String(h).replace(/\s+/g, '').toLowerCase();
        headerMap[clean] = i;
    });
    
    // 編輯模式檢查
    let rowIndexToUpdate = -1;
    let existingRecordId = dataObj['紀錄ID'];
    
    if (existingRecordId && targetSheetName !== CONFIG.SHEETS.CLIENT) {
       const idIdx = headerMap['紀錄id'];
       if (idIdx !== undefined) {
         const allIds = sheet.getRange(2, idIdx + 1, sheet.getLastRow() - 1 || 1, 1).getValues().flat();
         const matchIndex = allIds.indexOf(existingRecordId);
         if (matchIndex > -1) rowIndexToUpdate = matchIndex + 2;
       }
    }

    // 新增個案建立資料夾
    let generatedFolderUrl = '';
    if (targetSheetName === CONFIG.SHEETS.CLIENT && !dataObj['GoogleDrive資料夾連結'] && dataObj['個案編號'] && dataObj['姓名']) {
         try { generatedFolderUrl = createOrGetCaseFolder(CONFIG.PARENT_FOLDER_ID, dataObj['個案編號'], dataObj['姓名']); } catch (e) {}
    }

    // 準備寫入資料
    const rowData = rawHeaders.map(rawH => {
        const cleanH = String(rawH).replace(/\s+/g, '').toLowerCase();
        
        // 從 dataObj 找對應的值 (也要把 dataObj 的 key 標準化來比對)
        let val = '';
        for (let key in dataObj) {
            if (key.replace(/\s+/g, '').toLowerCase() === cleanH) {
                val = dataObj[key];
                break;
            }
        }

        if (cleanH === '紀錄id') return val || 'R' + Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), 'yyyyMMddHHmmss') + Math.floor(Math.random()*900+100);
        
        if (cleanH.includes('時間') || cleanH.includes('日期')) {
            // 如果是建立時間且是更新模式，讀取原值
            if ((cleanH === '建立時間' || cleanH === '建立日期') && rowIndexToUpdate > -1) {
                return sheet.getRange(rowIndexToUpdate, headerMap[cleanH]+1).getValue();
            }
            // 如果 dataObj 有傳日期 (如 治療日期)，直接使用；否則填入當下時間
            if (val) return val; 
            if (cleanH === '建立時間' || cleanH === '建立日期') return Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), 'yyyy-MM-dd HH:mm:ss');
        }

        if (cleanH === 'googledrive資料夾連結') return generatedFolderUrl || val;
        
        if (cleanH === '電話' || cleanH === '身分證字號' || cleanH === '個案編號') {
            return (val && String(val).startsWith("'")) ? val : ("'" + (val || ""));
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
    throw new Error(e.message);
  } finally {
    lock.releaseLock();
  }
}

// Wrapper Functions
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
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const sheet = ss.getSheetByName(CONFIG.SHEETS.TREATMENT);
    if (!sheet) return [];
    
    const data = sheet.getDataRange().getValues();
    if (data.length < 2) return [];

    // 1. 建立標題索引 (標準化：去空白、轉小寫)
    const rawHeaders = data[0];
    const map = {};
    rawHeaders.forEach((h, i) => {
        map[String(h).replace(/\s+/g, '').toLowerCase()] = i;
    });
    
    // 2. 取得關鍵欄位 Index
    const idxId = map['紀錄id'];
    const idxCaseId = map['個案編號'];
    const idxDate = map['治療日期'];
    const idxTherapist = map['執行治療師'];
    const idxComplaint = map['當日主訴'];
    const idxContent = map['治療內容'];
    
    // 容錯：備註欄位可能叫 "備註" 或 "備註/下次治療"
    let idxNote = map['備註/下次治療'];
    if (idxNote === undefined) idxNote = map['備註'];

    // 檢查必要欄位
    if (idxCaseId === undefined || idxDate === undefined) {
        // 若找不到欄位，回傳空陣列 (避免報錯)
        return [];
    }

    const logs = [];
    const targetId = String(clientId).replace(/^'/, '').trim().toLowerCase();

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      // ID 比對：去單引號、去空白、轉小寫
      const rowId = String(row[idxCaseId]).replace(/^'/, '').trim().toLowerCase();

      if (rowId === targetId) {
        
        let dateStr = row[idxDate];
        if (dateStr instanceof Date) {
          dateStr = Utilities.formatDate(dateStr, ss.getSpreadsheetTimeZone(), 'yyyy-MM-dd');
        } else {
          dateStr = dateStr ? String(dateStr) : "";
        }

        logs.push({
          '紀錄ID': idxId !== undefined ? row[idxId] : '',
          '個案編號': row[idxCaseId],
          '治療日期': dateStr,
          '執行治療師': idxTherapist !== undefined ? row[idxTherapist] : '',
          '當日主訴': idxComplaint !== undefined ? row[idxComplaint] : '',
          '治療內容': idxContent !== undefined ? row[idxContent] : '',
          '備註/下次治療': idxNote !== undefined ? row[idxNote] : ''
        });
      }
    }
    
    // 日期排序
    return logs.sort((a, b) => {
       const dA = a['治療日期'] ? new Date(a['治療日期']) : new Date(0);
       const dB = b['治療日期'] ? new Date(b['治療日期']) : new Date(0);
       return dB - dA;
    });
    
  } catch (e) {
    console.error(e);
    return [];
  }
}

// ------------------------------------------
// 4. Drive & Image (保持不變)
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