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



/**
 * 儲存醫師看診紀錄 - 包含備註欄位
 * 欄位對應：[A:紀錄ID, B:個案編號, C:看診日期, D:看診醫師, E:護理師, F:S_主訴, G:O_客觀檢查, H:A_診斷, I:P_治療計劃, J:備註, K:影像附件連結, L:建立時間]
 */
function saveDoctorConsultation(formData) {
  try {
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const sheet = ss.getSheetByName("Doctor_Consultation"); 
    
    if (!sheet) throw new Error("找不到工作表: Doctor_Consultation");

    const recordId = "DOC" + new Date().getTime();
    const timestamp = Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), "yyyy/MM/dd HH:mm:ss");
    
    // 嚴格按照您最新的試算表 12 個欄位順序排列
    const rowData = [
      recordId,                   // A: 紀錄ID
      "'" + formData.clientId,    // B: 個案編號
      formData.date,              // C: 看診日期
      formData.doctor,            // D: 看診醫師
      formData.nurse,             // E: 護理師
      formData.complaint,         // F: S_主訴
      formData.objective,         // G: O_客觀檢查
      formData.diagnosis,         // H: A_診斷
      formData.plan,              // I: P_治療計劃
      formData.remark,            // J: 備註 (新增)
      "",                         // K: 影像附件連結
      timestamp                   // L: 建立時間
    ];

    sheet.appendRow(rowData);
    return { success: true, message: "醫師看診紀錄儲存成功" };
  } catch (e) {
    return { success: false, message: "儲存失敗: " + e.toString() };
  }
}

/**
 * 後端：寫入試算表並建立 Google Drive 資料夾
 */
function createNewClient(data) {
  try {
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const sheet = ss.getSheetByName("Client_Basic_Info");
    
    // 1. 自動產生個案編號 (CF + 當天年月日 + 3位流水)
    const now = new Date();
    const datePart = Utilities.formatDate(now, "GMT+8", "yyyyMMdd");
    const lastRow = sheet.getLastRow();
    // 簡單流水號邏輯：直接取行數
    const suffix = (lastRow + 1).toString().padStart(3, '0');
    const clientId = "CF" + datePart + suffix;

    // 2. 自動建立 Google Drive 資料夾
    let folderUrl = "";
    try {
      // 在根目錄建立資料夾 (建議之後可在 CONFIG 指定父資料夾 ID)
      const folder = DriveApp.createFolder(clientId + "_" + data.name);
      folderUrl = folder.getUrl();
    } catch (e) {
      folderUrl = "資料夾建立失敗";
    }

    // 3. 準備寫入的資料列 (順序對應您的圖1)
    // A:編號, B:姓名, C:生日, D:身分證, E:電話, F:性別, G:緊急聯絡人, H:緊急聯絡電話, I:慢性病, J:Drive, K:建立日期
    const newRow = [
      clientId,           // A
      data.name,          // B
      data.dob,           // C
      data.idNo,          // D
      "'" + data.phone,   // E (加 ' 符號防止科學符號)
      data.gender,        // F
      data.emerName,      // G
      "'" + data.emerPhone,// H
      data.chronic,       // I
      folderUrl,          // J
      now                 // K
    ];

    sheet.appendRow(newRow);

    // 4. 回傳完整個案物件供前端渲染
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const fullData = {};
    headers.forEach((h, i) => {
      fullData[h] = newRow[i];
    });

    return { success: true, clientId: clientId, fullData: fullData };

  } catch (e) {
    return { success: false, message: e.toString() };
  }
}

/**
 * 獲取個案歷史紀錄 (通用)
 * 修正：針對不同工作表的日期欄位進行精確排序 (由新到舊)
 */
function getClientHistory(clientId, sheetName) {
  try {
    if (!clientId) return [];
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) return [];
    
    // 使用 getDisplayValues 確保日期格式與試算表一致
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

    // --- 排序邏輯：由新到舊 ---
    result.sort((a, b) => {
      // 兼容不同工作表的日期欄位名稱
      const dateStrA = a['看診日期'] || a['治療日期'] || a['日期'] || '1900-01-01';
      const dateStrB = b['看診日期'] || b['治療日期'] || b['日期'] || '1900-01-01';
      
      const dateA = new Date(dateStrA);
      const dateB = new Date(dateStrB);
      
      // 若日期無效，則放到最後
      if (isNaN(dateA)) return 1;
      if (isNaN(dateB)) return -1;
      
      return dateB - dateA; // 降序排序 (Newest first)
    });

    return result;
  } catch (e) { 
    console.error("getClientHistory Error: " + e.message);
    return []; 
  }
}