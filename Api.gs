/**
 * ==========================================
 * 設定檔 (Config.gs)
 * ==========================================
 */
const CONFIG = {
  SPREADSHEET_ID: '1LMhlQGyXNXq9Teqm0_W0zU9NbQlVCHKLDL0mSOiDomc', 
  PARENT_FOLDER_ID: '1NIsNHALeSSVm60Yfjc9k-u30A42CuZw8',
  SHEETS: {
    CLIENT: 'Client_Basic_Info',
    TREATMENT: 'Treatment_Logs',
    DOCTOR: 'Doctor_Consultation',
    TRACKING: 'Case_Tracking',
    SYSTEM: 'System',
    MAINTENANCE: 'Health_Maintenance', // 補上代碼中用到的保養表單名稱
    IMAGE: 'Image_Gallery'
  }
};

/**
 * ==========================================
 * 網頁應用程式入口 (WebApp.gs)
 * ==========================================
 */
function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('康飛運醫 | 個案管理系統')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * ==========================================
 * 核心邏輯層 (Api.gs)
 * ==========================================
 */

// 1. 搜尋功能
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
            // 若為追蹤日期等指定日期，保留原始輸入；否則填入當下時間
            if (cleanH === '追蹤日期' && val) return val;
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

/**
 * 1. 更新系統人員與項目清單 (從 System 工作表抓取)
 */
function getSystemStaff() {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const sheet = ss.getSheetByName(CONFIG.SHEETS.SYSTEM);
  const data = sheet.getDataRange().getValues();
  
  const rows = data.slice(1); // 移除標題
  
  return {
    doctors: rows.map(r => r[0]).filter(String),    // A欄
    nurses: rows.map(r => r[1]).filter(String),     // B欄
    therapists: rows.map(r => r[2]).filter(String), // C欄
    trackingTypes: rows.map(r => r[3]).filter(String), // D欄：追蹤項目 (個管追蹤用)
    maintItems: rows.map(r => r[4]).filter(String), // E欄：保養項目
    allStaff: rows.map(r => r[5]).filter(String)    // F欄：所有人員 (用於保養與追蹤人員)
  };
}

/**
 * 儲存個管追蹤紀錄 (DB_Tracking)
 */
function saveTrackingRecord(formObj) {
  try {
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const sheet = ss.getSheetByName(CONFIG.SHEETS.TRACKING);
    
    // 1. 產生唯一追蹤ID (TR + 年月日 + 3位隨機)
    const now = new Date();
    const dateStr = Utilities.formatDate(now, ss.getSpreadsheetTimeZone(), "yyyyMMdd");
    const uniqueId = "TR" + dateStr + Math.floor(Math.random() * 1000).toString().padStart(3, '0');
    
    const newRow = [
      uniqueId,                
      "'" + formObj.clientId,  // 加 ' 防止科學記號
      formObj.trackDate,
      formObj.trackStaff,      // 新增欄位：追蹤人員
      formObj.trackType,       
      formObj.content,         
      Utilities.formatDate(now, ss.getSpreadsheetTimeZone(), "yyyy-MM-dd HH:mm:ss")
    ];
    
    sheet.appendRow(newRow);
    return { success: true, message: "追蹤紀錄已新增" };
  } catch (e) {
    return { success: false, message: "儲存失敗: " + e.toString() };
  }
}

/**
 * 取得個管追蹤歷史紀錄
 */
function getTrackingHistory(clientId) {
  try {
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const sheet = ss.getSheetByName(CONFIG.SHEETS.TRACKING);
    if (!sheet) return [];
    
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    
    // 定義欄位索引
    const idxClientId = headers.indexOf("個案編號");
    const idxId = headers.indexOf("追蹤ID");
    const idxDate = headers.indexOf("追蹤日期");
    const idxStaff = headers.indexOf("追蹤人員");
    const idxType = headers.indexOf("追蹤項目");
    const idxContent = headers.indexOf("追蹤內容");
    
    // 過濾資料
    const records = data.slice(1)
      .filter(row => String(row[idxClientId]).replace(/^'/, '') === String(clientId))
      .map(row => {
        let dateDisplay = row[idxDate];
        if (dateDisplay instanceof Date) {
          dateDisplay = Utilities.formatDate(dateDisplay, ss.getSpreadsheetTimeZone(), "yyyy-MM-dd");
        }
        return {
          id: row[idxId],
          date: dateDisplay,
          staff: row[idxStaff],
          type: row[idxType],
          content: row[idxContent]
        };
      });
      
    // 依照日期倒序排列 (新的在前)
    return records.sort((a, b) => new Date(b.date) - new Date(a.date));
  } catch (e) {
    return [];
  }
}

/**
 * 儲存醫師看診紀錄
 */
function saveDoctorConsultation(formData) {
  try {
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const sheet = ss.getSheetByName(CONFIG.SHEETS.DOCTOR); 
    
    const recordId = "DOC" + new Date().getTime();
    const timestamp = Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), "yyyy/MM/dd HH:mm:ss");
    
    const rowData = [
      recordId,                   
      "'" + formData.clientId,    
      formData.date,              
      formData.doctor,            
      formData.nurse,             
      formData.complaint,         
      formData.objective,         
      formData.diagnosis,         
      formData.plan,              
      formData.remark,            
      "",                         
      timestamp                   
    ];

    sheet.appendRow(rowData);
    return { success: true, message: "醫師看診紀錄儲存成功" };
  } catch (e) {
    return { success: false, message: "儲存失敗: " + e.toString() };
  }
}

/**
 * 建立新個案
 */
function createNewClient(data) {
  try {
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const sheet = ss.getSheetByName(CONFIG.SHEETS.CLIENT);
    
    const now = new Date();
    const datePart = Utilities.formatDate(now, "GMT+8", "yyyyMMdd");
    const lastRow = sheet.getLastRow();
    const suffix = (lastRow + 1).toString().padStart(3, '0');
    const clientId = "CF" + datePart + suffix;

    let folderUrl = "";
    try {
      const folder = DriveApp.createFolder(clientId + "_" + data.name);
      folderUrl = folder.getUrl();
    } catch (e) {
      folderUrl = "資料夾建立失敗";
    }

    const newRow = [
      clientId,           
      data.name,          
      data.dob,           
      data.idNo,          
      "'" + data.phone,   
      data.gender,        
      data.emerName,      
      "'" + data.emerPhone,
      data.chronic,       
      folderUrl,          
      now                 
    ];

    sheet.appendRow(newRow);

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
 * 取得保養歷史
 */
function getMaintenanceHistory(clientId) {
  try {
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const sheet = ss.getSheetByName(CONFIG.SHEETS.MAINTENANCE);
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const clientIdx = headers.indexOf("個案編號");
    
    const results = data.slice(1)
      .filter(row => String(row[clientIdx]).replace(/^'/, '') == String(clientId))
      .map(row => {
        let obj = {};
        headers.forEach((h, i) => {
          let val = row[i];
          if (val instanceof Date) {
            val = Utilities.formatDate(val, "GMT+8", "yyyy-MM-dd");
          }
          obj[h] = val;
        });
        return obj;
      });
    return results.reverse();
  } catch (e) {
    return [];
  }
}

/**
 * 通用歷史紀錄 (用於治療紀錄等)
 */
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

    result.sort((a, b) => {
      const dateStrA = a['看診日期'] || a['治療日期'] || a['日期'] || '1900-01-01';
      const dateStrB = b['看診日期'] || b['治療日期'] || b['日期'] || '1900-01-01';
      const dateA = new Date(dateStrA);
      const dateB = new Date(dateStrB);
      if (isNaN(dateA)) return 1;
      if (isNaN(dateB)) return -1;
      return dateB - dateA;
    });

    return result;
  } catch (e) { 
    return []; 
  }
}

/**
 * 儲存保養紀錄
 */
function saveMaintenanceRecord(data) {
  try {
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const sheet = ss.getSheetByName(CONFIG.SHEETS.MAINTENANCE);
    
    const newRow = [
      Utilities.getUuid(),
      data.clientId,
      data.date,
      data.staff,
      data.item,
      data.bp,
      data.spo2,
      data.hr,
      data.temp,
      data.remark,
      new Date()
    ];
    
    sheet.appendRow(newRow);
    
    return { success: true, message: "保養紀錄儲存成功！" };
  } catch (e) {
    return { success: false, message: "儲存失敗：" + e.toString() };
  }
}

/**
 * 取得個案所有歷程資料 (聚合 4 個工作表)
 */
function getCaseOverviewData(clientId) {
  try {
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const result = [];
    
    // 1. 取得醫師看診紀錄 (Doctor_Consultation)
    const docSheet = ss.getSheetByName(CONFIG.SHEETS.DOCTOR);
    if (docSheet) {
      const data = docSheet.getDataRange().getValues();
      data.slice(1).forEach(row => {
        if (String(row[1]).replace(/^'/, '') === String(clientId)) {
          result.push({
            id: row[0],
            date: formatDateForJSON(row[2]),
            category: 'doctor', // 類別標記
            categoryName: '醫師看診',
            title: row[3] + " 醫師", // 標題顯示醫師名
            subtitle: "診斷：" + (row[7] || '--'), // 副標題顯示診斷
            detail: row[8], // 詳細內容顯示治療計畫
            staff: row[4] // 護理師
          });
        }
      });
    }

    // 2. 取得保養項目紀錄 (Health_Maintenance)
    const maintSheet = ss.getSheetByName(CONFIG.SHEETS.MAINTENANCE);
    if (maintSheet) {
      const data = maintSheet.getDataRange().getValues();
      data.slice(1).forEach(row => {
        if (String(row[1]).replace(/^'/, '') === String(clientId)) {
          // 組合生理數值字串
          const vitals = [];
          if(row[5]) vitals.push(`BP:${row[5]}`);
          if(row[6]) vitals.push(`SpO2:${row[6]}%`);
          result.push({
            id: row[0],
            date: formatDateForJSON(row[2]),
            category: 'maintenance',
            categoryName: '保養項目',
            title: row[4], // 保養項目
            subtitle: vitals.join(' | ') || '無生理數值',
            detail: row[9], // 備註
            staff: row[3] // 執行人員
          });
        }
      });
    }

    // 3. 取得個管追蹤紀錄 (DB_Tracking)
    const trackSheet = ss.getSheetByName(CONFIG.SHEETS.TRACKING);
    if (trackSheet) {
      const data = trackSheet.getDataRange().getValues();
      data.slice(1).forEach(row => {
        if (String(row[1]).replace(/^'/, '') === String(clientId)) {
          result.push({
            id: row[0],
            date: formatDateForJSON(row[2]),
            category: 'tracking',
            categoryName: '個管追蹤',
            title: row[4], // 追蹤項目
            subtitle: "人員：" + (row[3] || '--'),
            detail: row[5], // 內容
            staff: row[3]
          });
        }
      });
    }

    // 4. 取得治療紀錄 (Treatment_Logs)
    const treatSheet = ss.getSheetByName(CONFIG.SHEETS.TREATMENT);
    if (treatSheet) {
      const data = treatSheet.getDataRange().getValues();
      const headers = data[0]; // 需要動態對應 Header
      const idxId = headers.indexOf("個案編號");
      const idxDate = headers.indexOf("治療日期");
      const idxStaff = headers.indexOf("執行治療師");
      const idxContent = headers.indexOf("治療內容");
      
      data.slice(1).forEach(row => {
        if (String(row[idxId]).replace(/^'/, '') === String(clientId)) {
          result.push({
            id: 'T-' + formatDateForJSON(row[idxDate]), // 治療紀錄沒 ID，用日期代替
            date: formatDateForJSON(row[idxDate]),
            category: 'treatment',
            categoryName: '治療紀錄',
            title: "物理治療",
            subtitle: "治療師：" + (row[idxStaff] || '--'),
            detail: row[idxContent],
            staff: row[idxStaff]
          });
        }
      });
    }

    // 排序：日期新到舊
    return result.sort((a, b) => new Date(b.date) - new Date(a.date));

  } catch (e) {
    throw new Error("取得總覽資料失敗: " + e.message);
  }
}

// 輔助函式：日期轉字串 yyyy-MM-dd
function formatDateForJSON(dateVal) {
  if (!dateVal) return "";
  if (dateVal instanceof Date) {
    return Utilities.formatDate(dateVal, Session.getScriptTimeZone(), "yyyy-MM-dd");
  }
  return String(dateVal);
}