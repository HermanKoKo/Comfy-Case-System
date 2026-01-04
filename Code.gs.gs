/**
 * ==========================================
 * 設定檔 (Config.gs)
 * ==========================================
 */
const CONFIG = {
  // 原本的系統資料庫 (存放紀錄、設定、影像等)
  SPREADSHEET_ID: '1LMhlQGyXNXq9Teqm0_W0zU9NbQlVCHKLDL0mSOiDomc', 
  
  // ★ 已修復：個案核心資料庫 (移除網址後綴，只保留純 ID)
  CLIENT_DB_ID: '1SPLLacdq9RYV6Jfur-pZQiQwMqGuBFEZ5jaho_0WUK0',

  PARENT_FOLDER_ID: '1NIsNHALeSSVm60Yfjc9k-u30A42CuZw8',
  SHEETS: {
    CLIENT: 'Client_Basic_Info',
    TREATMENT: 'Treatment_Logs',
    DOCTOR: 'Doctor_Consultation',
    TRACKING: 'Case_Tracking',
    SYSTEM: 'System',
    MAINTENANCE: 'Health_Maintenance',
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

// 輔助：正規化標題 (移除空白、轉小寫) 以避免欄位名稱對應錯誤
function normalizeHeader(header) {
  return String(header).replace(/\s+/g, '').trim().toLowerCase();
}

/**
 * ★ 關鍵核心：智慧分頁選取器
 * 根據 Sheet 名稱自動判斷要連線到「原資料庫」還是「新個案資料庫」
 */
function getSheetHelper(sheetName) {
  let ss;
  // 如果請求的是個案基本資料，連線到新資料庫
  if (sheetName === CONFIG.SHEETS.CLIENT) {
    try {
      ss = SpreadsheetApp.openById(CONFIG.CLIENT_DB_ID);
    } catch (e) {
      throw new Error("無法連接個案核心資料庫 (Client DB)，請檢查 ID 是否正確或是否有權限。");
    }
  } else {
    // 其他資料 (治療紀錄、系統設定等) 維持在原資料庫
    ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  }

  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    // 自動建立缺少的 Sheet (選擇性功能，避免報錯)
    if (sheetName !== CONFIG.SHEETS.CLIENT) {
       return ss.insertSheet(sheetName);
    }
    throw new Error("找不到工作表: " + sheetName);
  }
  return sheet;
}

/**
 * 修正一：取得治療師下拉選單選項
 * 來源：Google Sheets 'system' 分頁的 C 欄 (從第 2 列開始)
 */
function getTherapistList() {
  try {
    // 1. 取得試算表與 'system' 分頁
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("system");
    
    if (!sheet) {
      throw new Error("找不到名為 'system' 的分頁");
    }

    // 2. 取得最後一行，避免讀取過多空行
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return []; // 若只有標題或無資料，回傳空陣列

    // 3. 讀取 C 欄資料 (Row 2, Column 3, 讀取 LastRow-1 行)
    const data = sheet.getRange(2, 3, lastRow - 1, 1).getValues();

    // 4. 扁平化陣列並過濾空值與重複值
    const therapists = data
      .flat()
      .filter(function(name) { return name && name.toString().trim() !== ""; });
      
    return therapists;
    
  } catch (e) {
    Logger.log("Error in getTherapistList: " + e.toString());
    return []; // 發生錯誤時回傳空陣列
  }
}

// 1. 搜尋功能 (已修復：支援模糊搜尋與忽略空格)
function searchClient(keyword) {
  try {
    // 使用 helper 自動取得正確的 Sheet (會在外部 DB 尋找)
    const sheet = getSheetHelper(CONFIG.SHEETS.CLIENT);
    
    const data = sheet.getDataRange().getDisplayValues(); 
    const results = [];
    
    // ★ 優化：移除搜尋字詞中的所有空白並轉小寫，提高容錯率
    const query = String(keyword).replace(/\s+/g, '').toLowerCase();
    if (!query) return [];
    
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      // 取得欄位資料並移除空白進行比對
      const id = String(row[0]).replace(/^'/, '').replace(/\s+/g, '').toLowerCase();
      const name = String(row[1]).replace(/\s+/g, '').toLowerCase();
      const phone = String(row[4]).replace(/^'/, '').replace(/\s+/g, '').toLowerCase();
      
      // 比對邏輯：只要 ID、姓名或電話包含關鍵字即可 (支援模糊搜尋)
      if (id.includes(query) || name.includes(query) || phone.includes(query)) {
        results.push({
          '個案編號': row[0], 
          '姓名': row[1], 
          '生日': row[2], 
          '身分證字號': row[3],
          '電話': row[4], 
          '性別': row[5], 
          '緊急聯絡人': row[6], 
          '緊急聯絡人電話': row[7],
          // ★ 修改處：新增「負責治療師」(I欄, Index 8)，後續欄位索引順延
          '負責治療師': row[8],
          '慢性病或特殊疾病': row[9], // 原本 Index 8 -> 變 9
          'GoogleDrive資料夾連結': row[10], // 原本 Index 9 -> 變 10
          '建立日期': row[11] // 原本 Index 10 -> 變 11
        });
      }
    }
    return results;
  } catch (e) { throw new Error(e.message); }
}

// 2. 通用資料儲存功能 (已修改：支援跨資料庫寫入)
function saveData(sheetName, dataObj) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000); 
    
    // ★ 修改：使用 helper 取得 Sheet，無論它是本地還是外部
    const targetSheetName = sheetName || CONFIG.SHEETS.CLIENT;
    const sheet = getSheetHelper(targetSheetName);
    
    // 取得該 Sheet 所屬的 Spreadsheet 物件 (為了取得正確的 TimeZone)
    const ss = sheet.getParent();

    const rawHeaders = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const headerMap = {}; 
    rawHeaders.forEach((h, i) => headerMap[normalizeHeader(h)] = i);
    
    const cleanIdKey = normalizeHeader('個案編號');
    let hasClientId = false;
    for (let key in dataObj) {
        if (normalizeHeader(key) === cleanIdKey && dataObj[key]) hasClientId = true;
    }
    if (targetSheetName !== CONFIG.SHEETS.CLIENT && !hasClientId) {
        throw new Error("系統錯誤：未偵測到個案編號，無法儲存。");
    }

    const rowData = rawHeaders.map(rawH => {
        const cleanH = normalizeHeader(rawH);
        let val = '';
        for (let key in dataObj) {
            if (normalizeHeader(key) === cleanH) { val = dataObj[key]; break; }
        }
        if (cleanH === '紀錄id') return val || 'R' + Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), 'yyyyMMddHHmmss') + Math.floor(Math.random()*900+100);
        if (cleanH.includes('時間') || cleanH.includes('日期')) {
            if (cleanH === '追蹤日期' && val) return val;
            return val || Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), 'yyyy-MM-dd HH:mm:ss');
        }
        if (['電話', '身分證字號', '個案編號'].includes(cleanH)) return "'" + String(val || "");
        return val || '';
    });

    sheet.appendRow(rowData);
    return { success: true, message: "資料已新增" };
    
  } catch (e) { throw new Error(e.message); } finally { lock.releaseLock(); }
}

/**
 * 取得系統人員與項目清單
 */
function getSystemStaff() {
  // 使用 helper 取得 System 分頁 (位於原資料庫)
  const sheet = getSheetHelper(CONFIG.SHEETS.SYSTEM);
  
  const data = sheet.getDataRange().getValues();
  const rows = data.slice(1);
  return {
    doctors: rows.map(r => r[0]).filter(String),
    nurses: rows.map(r => r[1]).filter(String),
    therapists: rows.map(r => r[2]).filter(String), // C欄 = Index 2
    trackingTypes: rows.map(r => r[3]).filter(String),
    maintItems: rows.map(r => r[4]).filter(String),
    allStaff: rows.map(r => r[5]).filter(String),
    treatmentItems: rows.map(r => r[6]).filter(String)
  };
}

/**
 * 儲存個管追蹤紀錄
 */
function saveTrackingRecord(formObj) {
  try {
    // 使用 helper 取得 Tracking 分頁 (位於原資料庫)
    const sheet = getSheetHelper(CONFIG.SHEETS.TRACKING);
    // getSheetHelper 內建自動建立分頁功能，此處可直接使用
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(["追蹤ID", "個案編號", "追蹤日期", "追蹤人員", "追蹤項目", "追蹤內容", "建立時間"]);
    }
    
    // 取得該分頁的 Spreadsheet 物件以獲取時區
    const ss = sheet.getParent();

    if (!formObj.clientId) throw new Error("無個案編號");

    const now = new Date();
    const uniqueId = "TR" + new Date().getTime();
    
    const newRow = [
      uniqueId,                
      "'" + formObj.clientId,
      formObj.trackDate,
      formObj.trackStaff,      
      formObj.trackType,       
      formObj.content,         
      Utilities.formatDate(now, ss.getSpreadsheetTimeZone(), "yyyy-MM-dd HH:mm:ss")
    ];
    
    sheet.appendRow(newRow);
    return { success: true, message: "追蹤紀錄已新增" };
  } catch (e) { return { success: false, message: "儲存失敗: " + e.toString() }; }
}

/**
 * 取得個管追蹤歷史紀錄
 */
function getTrackingHistory(clientId) {
  try {
    if (!clientId) return [];
    const sheet = getSheetHelper(CONFIG.SHEETS.TRACKING);
    const ss = sheet.getParent();
    
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return [];

    const headers = data[0];
    const idxClientId = headers.map(normalizeHeader).indexOf("個案編號");
    const targetColIdx = idxClientId > -1 ? idxClientId : 1;

    const idxId = headers.indexOf("追蹤ID");
    const idxDate = headers.indexOf("追蹤日期");
    const idxStaff = headers.indexOf("追蹤人員");
    const idxType = headers.indexOf("追蹤項目");
    const idxContent = headers.indexOf("追蹤內容");
    
    const targetId = String(clientId).trim();

    const records = data.slice(1)
      .filter(row => String(row[targetColIdx]).replace(/^'/, '').trim() === targetId)
      .map(row => {
        let dateDisplay = row[idxDate];
        if (dateDisplay instanceof Date) dateDisplay = Utilities.formatDate(dateDisplay, ss.getSpreadsheetTimeZone(), "yyyy-MM-dd");
        return {
          id: row[idxId],
          date: dateDisplay,
          staff: row[idxStaff],
          type: row[idxType],
          content: row[idxContent]
        };
      });
      
    return records.sort((a, b) => new Date(b.date) - new Date(a.date));
  } catch (e) { return []; }
}

/**
 * 儲存醫師看診紀錄
 */
function saveDoctorConsultation(formData) {
  try {
    const sheet = getSheetHelper(CONFIG.SHEETS.DOCTOR);
    const ss = sheet.getParent();
    
    if (!formData.clientId) throw new Error("無個案編號");

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
      formData.nursingRecord,     // 護理紀錄 (J 欄)
      formData.remark,            // K 欄
      "",                         // L 欄
      timestamp                   // M 欄
    ];

    sheet.appendRow(rowData);
    return { success: true, message: "醫師看診紀錄儲存成功" };
  } catch (e) { return { success: false, message: "儲存失敗: " + e.toString() }; }
}

/**
 * 建立新個案 (★ 關鍵：寫入新資料庫)
 */
function createNewClient(data) {
  try {
    // 使用 helper 取得 Client 分頁 (位於新資料庫)
    const sheet = getSheetHelper(CONFIG.SHEETS.CLIENT);
    
    const now = new Date();
    const datePart = Utilities.formatDate(now, "GMT+8", "yyyyMMdd");
    
    // 從新資料庫取得最後一行，確保 ID 生成不重複
    const lastRow = sheet.getLastRow();
    const suffix = (lastRow + 1).toString().padStart(3, '0');
    const clientId = "CF" + datePart + suffix;

    let folderUrl = "";
    // 建立資料夾 (保留在原設定的 Drive 資料夾中，不受影響)
    try { 
        const parentFolder = DriveApp.getFolderById(CONFIG.PARENT_FOLDER_ID);
        const folder = parentFolder.createFolder(clientId + "_" + data.name); 
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
      // ★ 修改處：插入「負責治療師」(I欄)，後續欄位順延
      data.therapist || "",
      data.chronic,       // 變 J
      folderUrl,          // 變 K
      now                 // 變 L
    ];

    sheet.appendRow(newRow);

    const fullData = {
        '個案編號': clientId, '姓名': data.name, '生日': data.dob, '身分證字號': data.idNo,
        '電話': data.phone, '性別': data.gender, '緊急聯絡人': data.emerName, 
        '緊急聯絡人電話': data.emerPhone, 
        // ★ 修改處：回傳物件也加上治療師
        '負責治療師': data.therapist,
        '慢性病或特殊疾病': data.chronic
    };

    return { success: true, clientId: clientId, fullData: fullData };
  } catch (e) { return { success: false, message: e.toString() }; }
}

/**
 * 取得保養歷史
 */
function getMaintenanceHistory(clientId) {
  try {
    if (!clientId) return [];
    const sheet = getSheetHelper(CONFIG.SHEETS.MAINTENANCE);
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    
    let clientIdx = headers.map(normalizeHeader).indexOf(normalizeHeader("個案編號"));
    if (clientIdx === -1) clientIdx = 1; 
    
    const targetId = String(clientId).trim();

    const results = data.slice(1)
      .filter(row => String(row[clientIdx]).replace(/^'/, '').trim() === targetId)
      .map(row => {
        let obj = {};
        headers.forEach((h, i) => {
          let val = row[i];
          if (val instanceof Date) val = Utilities.formatDate(val, "GMT+8", "yyyy-MM-dd");
          obj[h] = val;
        });
        return obj;
      });
    return results.reverse();
  } catch (e) { return []; }
}

/**
 * 通用歷史紀錄 (用於治療紀錄等)
 */
function getClientHistory(clientId, sheetName) {
  try {
    if (!clientId) return [];
    // 使用 helper，自動適應資料庫位置
    const sheet = getSheetHelper(sheetName);
    
    const data = sheet.getDataRange().getDisplayValues();
    if (data.length < 2) return [];
    
    const headers = data[0];
    const normHeaders = headers.map(normalizeHeader);
    let idxCaseId = normHeaders.indexOf(normalizeHeader('個案編號'));
    if (idxCaseId === -1) idxCaseId = 1;
    
    const targetId = String(clientId).replace(/^'/, '').trim().toLowerCase();
    const result = [];
    
    for (let i = 1; i < data.length; i++) {
      const rowId = String(data[i][idxCaseId]).replace(/^'/, '').trim().toLowerCase();
      if (rowId === targetId) {
        let obj = {};
        headers.forEach((header, index) => { obj[header] = data[i][index]; });
        result.push(obj);
      }
    }

    result.sort((a, b) => {
      const dateStrA = a['看診日期'] || a['治療日期'] || a['日期'] || '1900-01-01';
      const dateStrB = b['看診日期'] || b['治療日期'] || b['日期'] || '1900-01-01';
      return new Date(dateStrB) - new Date(dateStrA);
    });

    return result;
  } catch (e) { return []; }
}

/**
 * 儲存保養紀錄
 */
function saveMaintenanceRecord(data) {
  try {
    const sheet = getSheetHelper(CONFIG.SHEETS.MAINTENANCE);
    
    if (!data.clientId) throw new Error("無個案編號");

    const newRow = [
      Utilities.getUuid(),
      "'" + data.clientId,
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
  } catch (e) { return { success: false, message: "儲存失敗：" + e.toString() }; }
}

/**
 * 取得個案總覽資料 (已修改：支援跨資料庫聚合)
 */
function getCaseOverviewData(clientId) {
  try {
    if (!clientId) return [];
    
    const result = [];
    const targetId = String(clientId).trim();
    
    // 1. 醫師看診 (Doctor)
    try {
      const docSheet = getSheetHelper(CONFIG.SHEETS.DOCTOR);
      const data = docSheet.getDataRange().getValues();
      const idx = data[0].map(normalizeHeader).indexOf(normalizeHeader("個案編號"));
      const targetCol = idx > -1 ? idx : 1; 

      data.slice(1).forEach(row => {
        if (String(row[targetCol]).replace(/^'/, '').trim() === targetId) {
          result.push({
            id: row[0],
            date: formatDateForJSON(row[2]),
            category: 'doctor', 
            categoryName: '醫師看診',
            doctor: row[3],
            nurse: row[4],
            s: row[5],
            o: row[6],
            a: row[7],
            p: row[8],
            nursingRecord: row[9],
            remark: row[10]
          });
        }
      });
    } catch (e) { console.log("讀取醫師看診失敗或是空表: " + e.toString()); }

    // 2. 保養項目 (Maintenance)
    try {
      const maintSheet = getSheetHelper(CONFIG.SHEETS.MAINTENANCE);
      const data = maintSheet.getDataRange().getValues();
      const idx = data[0].map(normalizeHeader).indexOf(normalizeHeader("個案編號"));
      const targetCol = idx > -1 ? idx : 1;
      
      data.slice(1).forEach(row => {
        if (String(row[targetCol]).replace(/^'/, '').trim() === targetId) {
          result.push({
            id: row[0],
            date: formatDateForJSON(row[2]),
            category: 'maintenance', 
            categoryName: '保養項目',
            staff: row[3],
            item: row[4],
            bp: row[5],
            spo2: row[6],
            hr: row[7],
            temp: row[8],
            remark: row[9]
          });
        }
      });
    } catch (e) { console.log("讀取保養項目失敗: " + e.toString()); }

    // 3. 個管追蹤 (Tracking)
    try {
      const trackSheet = getSheetHelper(CONFIG.SHEETS.TRACKING);
      const data = trackSheet.getDataRange().getValues();
      const idx = data[0].map(normalizeHeader).indexOf(normalizeHeader("個案編號"));
      const targetCol = idx > -1 ? idx : 1;

      data.slice(1).forEach(row => {
        if (String(row[targetCol]).replace(/^'/, '').trim() === targetId) {
          result.push({
            id: row[0],
            date: formatDateForJSON(row[2]),
            category: 'tracking', 
            categoryName: '個管追蹤',
            staff: row[3],
            type: row[4],
            content: row[5]
          });
        }
      });
    } catch (e) { console.log("讀取個管追蹤失敗: " + e.toString()); }

    // 4. 治療紀錄 (Treatment)
    try {
      const treatSheet = getSheetHelper(CONFIG.SHEETS.TREATMENT);
      const data = treatSheet.getDataRange().getValues();
      const headers = data[0].map(normalizeHeader);
      let idxId = headers.indexOf(normalizeHeader("個案編號"));
      if (idxId === -1) idxId = 1; 
      
      const idxDate = headers.indexOf(normalizeHeader("治療日期"));
      const idxStaff = headers.indexOf(normalizeHeader("執行治療師"));
      const idxItem = headers.indexOf(normalizeHeader("治療項目"));
      const idxComplaint = headers.indexOf(normalizeHeader("當日主訴"));
      const idxContent = headers.indexOf(normalizeHeader("治療內容"));
      const idxNext = headers.indexOf(normalizeHeader("備註/下次治療"));
      
      data.slice(1).forEach(row => {
        if (String(row[idxId]).replace(/^'/, '').trim() === targetId) {
          result.push({
            id: 'T-' + formatDateForJSON(row[idxDate]), 
            date: formatDateForJSON(row[idxDate]),
            category: 'treatment', 
            categoryName: '治療紀錄',
            staff: row[idxStaff],
            item: (idxItem > -1) ? row[idxItem] : "",
            complaint: (idxComplaint > -1) ? row[idxComplaint] : "",
            content: (idxContent > -1) ? row[idxContent] : "",
            nextPlan: (idxNext > -1) ? row[idxNext] : "" 
          });
        }
      });
    } catch (e) { console.log("讀取治療紀錄失敗: " + e.toString()); }

    // 排序：新到舊
    return result.sort((a, b) => new Date(b.date) - new Date(a.date));

  } catch (e) { throw new Error("取得總覽資料失敗: " + e.message); }
}

// ★ 影像功能 - 大幅修改版：直接從 Client DB 讀取資料夾連結，掃描 Drive 檔案
function getClientImages(clientId) {
  try {
    // 1. 先去 Client_Basic_Info 找到該個案的資料夾連結
    const clientSheet = getSheetHelper(CONFIG.SHEETS.CLIENT);
    const clientData = clientSheet.getDataRange().getDisplayValues(); // 使用 DisplayValues 以避免格式問題
    
    let folderUrl = "";
    const targetId = String(clientId).replace(/^'/, '').trim();
    
    // 從第二列開始遍歷資料
    for (let i = 1; i < clientData.length; i++) {
      const rowId = String(clientData[i][0]).replace(/^'/, '').trim();
      
      // 比對個案編號
      if (rowId === targetId) {
        // ★ 關鍵：資料夾連結位於第 11 欄 (Index 10)
        // (根據 createNewClient 邏輯：... 治療師(8), 慢性病(9), 連結(10) ...)
        folderUrl = clientData[i][10];
        break;
      }
    }
    
    // 如果找不到連結，回傳空陣列 (前端會顯示無影像)
    if (!folderUrl) {
      return { success: true, images: [] };
    }
    
    // 2. 解析 Drive Folder ID
    const idMatch = folderUrl.match(/[-\w]{25,}/);
    if (!idMatch) {
       // 連結格式錯誤
       return { success: true, images: [] };
    }
    const folderId = idMatch[0];
    
    // 3. 掃描該資料夾內的檔案
    const folder = DriveApp.getFolderById(folderId);
    const files = folder.getFiles();
    const imageList = [];
    
    while (files.hasNext()) {
      const file = files.next();
      const mimeType = file.getMimeType();
      
      // 簡單過濾：只顯示圖片類型的檔案
      if (mimeType.indexOf('image/') === 0) {
        const fileId = file.getId();
        imageList.push({
          id: fileId,
          name: file.getName(),
          url: file.getUrl(),
          // 產生縮圖連結
          thumbnail: "https://lh3.googleusercontent.com/d/" + fileId + "=s400",
          date: Utilities.formatDate(file.getDateCreated(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm"),
          remark: "" // 直接讀取 Drive，沒有額外的備註欄位
        });
      }
    }
    
    // 4. 依照日期排序 (新到舊)
    imageList.sort((a, b) => new Date(b.date) - new Date(a.date));
    
    return { success: true, images: imageList };

  } catch (e) { 
    return { success: false, message: "讀取影像失敗: " + e.toString() }; 
  }
}

// 上傳影像 - 修改版：從新資料庫查 Folder，寫入舊資料庫
function uploadClientImage(clientId, fileData, fileName, mimeType, remark) {
  try {
    // 1. 先取得資料夾位置 (★ 關鍵：必須去新的 Client DB 找)
    const clientSheet = getSheetHelper(CONFIG.SHEETS.CLIENT);
    const clientData = clientSheet.getDataRange().getDisplayValues();
    let folderUrl = "";
    
    for (let i = 1; i < clientData.length; i++) {
      if (String(clientData[i][0]).replace(/^'/, '') === String(clientId)) { 
          folderUrl = clientData[i][9]; // 這裡原本索引是9，因為我們插入了治療師，這裡是否需要調整?
          // ★ 修正：由於前面 searchClient 我們把 Drive 連結改成 10，所以這裡也要改成 10
          // 但原始代碼這邊原本寫 row[9]，現在資料結構變了，治療師(8), 慢性病(9), Drive(10)。
          // 所以這裡要改為 10
          folderUrl = clientData[i][10]; 
          break; 
      }
    }
    
    if (!folderUrl) throw new Error("找不到個案資料夾 (資料表無連結)");
    const folderIdMatch = folderUrl.match(/[-\w]{25,}/);
    if (!folderIdMatch) throw new Error("資料夾 ID 解析失敗: " + folderUrl);
    
    const folder = DriveApp.getFolderById(folderIdMatch[0]);
    
    // 2. 儲存實體檔案到 Drive
    const blob = Utilities.newBlob(Utilities.base64Decode(fileData), mimeType, fileName);
    const file = folder.createFile(blob);
    const fileUrl = file.getUrl();
    
    // 3. 寫入資料到 Image_Gallery Sheet (★ 寫回原資料庫)
    // 雖然現在 getClientImages 已經改為直接讀 Drive，但為了保持上傳紀錄的完整性，我們還是寫入 Sheet
    let imgSheet = getSheetHelper(CONFIG.SHEETS.IMAGE);
    const ss = imgSheet.getParent(); // 取得原資料庫的 Spreadsheet 物件以獲取時區

    if (imgSheet.getLastRow() === 0) {
      imgSheet.appendRow(["影像ID", "個案編號", "上傳日期", "檔案名稱", "GoogleDrive檔案連結", "備註"]);
    }

    const uniqueImgId = "IMG" + new Date().getTime();
    const nowStr = Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), "yyyy-MM-dd HH:mm");
    
    imgSheet.appendRow([
      uniqueImgId,
      "'" + clientId,
      nowStr,
      fileName,
      fileUrl,
      remark || ""
    ]);

    return { success: true, message: "上傳成功" };
  } catch (e) { return { success: false, message: "上傳失敗: " + e.toString() }; }
}

function formatDateForJSON(dateVal) {
  if (!dateVal) return "";
  if (dateVal instanceof Date) return Utilities.formatDate(dateVal, Session.getScriptTimeZone(), "yyyy-MM-dd");
  return String(dateVal);
}

/**
 * 從 "System" 分頁的 G 欄讀取治療項目清單
 * 假設第一列是標題，資料從第二列開始
 */
function getTreatmentItemsFromSystem() {
  // 使用 helper 取得目前綁定的 "System" 工作表
  var sheet = getSheetHelper(CONFIG.SHEETS.SYSTEM);
  
  // 取得最後一列的列號
  var lastRow = sheet.getLastRow();
  
  // 如果只有標題或沒資料，回傳空陣列
  if (lastRow < 2) {
    return [];
  }
  
  // 讀取 G 欄的範圍：從第 2 列開始，第 7 欄 (G欄)，讀取到最後一列
  // getRange(row, column, numRows)
  var range = sheet.getRange(2, 7, lastRow - 1);
  var values = range.getValues();
  
  // 整理資料：扁平化陣列並過濾掉空值
  var items = values.flat().filter(function(item) {
    return item !== "" && item != null;
  });
  
  return items;
}

/**
 * ==========================================
 * 資料更新專用邏輯 (修正版：自動偵測 ID 欄位)
 * ==========================================
 */
function updateRecord(type, formData) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) {
    return { success: false, message: "系統忙碌中，請稍後再試" };
  }

  try {
    let sheetName = "";
    // 取得前端傳來的 ID (轉成字串並去除空白)
    let targetId = String(formData.record_id).trim(); 
    
    // 1. 根據類型決定 Sheet 名稱與預期 ID 標題
    let expectedIdHeader = "紀錄id"; // 預設
    
    if (type === 'treatment') {
       sheetName = CONFIG.SHEETS.TREATMENT;
       expectedIdHeader = "紀錄id";
    }
    else if (type === 'doctor') {
       sheetName = CONFIG.SHEETS.DOCTOR;
       expectedIdHeader = "紀錄id"; // 通常看診紀錄 ID 也在這，若您標題不同請在此修改
    }
    else if (type === 'maintenance') {
       sheetName = CONFIG.SHEETS.MAINTENANCE;
       expectedIdHeader = "紀錄id"; // 或是 "影像ID"、"保養ID"，視您標題而定
    }
    else if (type === 'tracking') {
       sheetName = CONFIG.SHEETS.TRACKING;
       expectedIdHeader = "追蹤ID";
    }
    else {
       throw new Error("未知的編輯類型");
    }

    // 2. 取得 Sheet 與資料
    const sheet = getSheetHelper(sheetName);
    const data = sheet.getDataRange().getValues();
    
    if (data.length < 2) return { success: false, message: "資料表為空" };

    const headers = data[0]; // 第一列標題
    let idColIndex = -1;

    // ★★★ 關鍵修正：自動尋找 ID 在哪一欄 ★★★
    // 先嘗試找標題完全符合的
    for (let c = 0; c < headers.length; c++) {
      if (normalizeHeader(headers[c]) === normalizeHeader(expectedIdHeader)) {
        idColIndex = c;
        break;
      }
    }
    
    // 如果找不到標題，且是醫師/追蹤/保養 (通常 ID 在第一欄)，則預設為 0
    if (idColIndex === -1) {
       if (type !== 'treatment') {
          idColIndex = 0; 
       } else {
          // 治療紀錄通常依賴 header，若找不到很危險，嘗試找包含 "ID" 的欄位
          for (let c = 0; c < headers.length; c++) {
            if (String(headers[c]).toLowerCase().includes("id")) {
               idColIndex = c;
               break;
            }
          }
          // 真的找不到，報錯
          if (idColIndex === -1) throw new Error(`在工作表 ${sheetName} 中找不到「${expectedIdHeader}」欄位，無法定位資料。`);
       }
    }

    // 3. 尋找對應 ID 的列
    let rowIndex = -1;
    for (let i = 1; i < data.length; i++) {
      // 比對時轉字串並 trim，增加容錯
      if (String(data[i][idColIndex]).trim() === targetId) {
        rowIndex = i + 1; // 轉為實際列號 (1-based)
        break;
      }
    }

    if (rowIndex === -1) {
      return { success: false, message: `找不到該筆資料 (ID: ${targetId})，可能已被刪除或 ID 欄位判斷錯誤。` };
    }

    // 4. 執行更新 (欄位名稱對應)
    // 為了精確更新，我們再次利用 header 找欄位索引
    const getCol = (name) => {
       for(let k=0; k<headers.length; k++) {
         if(normalizeHeader(headers[k]) === normalizeHeader(name)) return k + 1;
       }
       return -1;
    };

    if (type === 'treatment') {
       // 更新治療紀錄：使用動態欄位搜尋
       const colDate = getCol("治療日期");
       const colTherapist = getCol("執行治療師");
       const colItem = getCol("治療項目");
       const colComplaint = getCol("當日主訴");
       const colContent = getCol("治療內容");
       const colNext = getCol("備註/下次治療");

       if (colDate > 0) sheet.getRange(rowIndex, colDate).setValue(formData.date);
       if (colTherapist > 0) sheet.getRange(rowIndex, colTherapist).setValue(formData.therapist);
       if (colItem > 0) sheet.getRange(rowIndex, colItem).setValue(formData.item);
       if (colComplaint > 0) sheet.getRange(rowIndex, colComplaint).setValue(formData.complaint);
       if (colContent > 0) sheet.getRange(rowIndex, colContent).setValue(formData.content);
       if (colNext > 0) sheet.getRange(rowIndex, colNext).setValue(formData.nextPlan);
    } 
    else if (type === 'doctor') {
       // 醫師看診通常欄位固定，但也可用動態搜尋保險
       // 這裡維持索引更新，但您可以改成 getCol 方式更安全
       sheet.getRange(rowIndex, 3).setValue(formData.date);
       sheet.getRange(rowIndex, 4).setValue(formData.doctor);
       sheet.getRange(rowIndex, 5).setValue(formData.nurse);
       sheet.getRange(rowIndex, 6).setValue(formData.complaint);
       sheet.getRange(rowIndex, 7).setValue(formData.objective);
       sheet.getRange(rowIndex, 8).setValue(formData.diagnosis);
       sheet.getRange(rowIndex, 9).setValue(formData.plan);
       sheet.getRange(rowIndex, 10).setValue(formData.nursingRecord);
       sheet.getRange(rowIndex, 11).setValue(formData.remark);
    } 
    else if (type === 'maintenance') {
       sheet.getRange(rowIndex, 3).setValue(formData.date);
       sheet.getRange(rowIndex, 4).setValue(formData.staff);
       sheet.getRange(rowIndex, 5).setValue(formData.item);
       sheet.getRange(rowIndex, 6).setValue(formData.bp);
       sheet.getRange(rowIndex, 7).setValue(formData.spo2);
       sheet.getRange(rowIndex, 8).setValue(formData.hr);
       sheet.getRange(rowIndex, 9).setValue(formData.temp);
       sheet.getRange(rowIndex, 10).setValue(formData.remark);
    } 
    else if (type === 'tracking') {
       sheet.getRange(rowIndex, 3).setValue(formData.trackDate);
       sheet.getRange(rowIndex, 4).setValue(formData.trackStaff);
       sheet.getRange(rowIndex, 5).setValue(formData.trackType);
       sheet.getRange(rowIndex, 6).setValue(formData.content);
    }

    return { success: true, message: "資料更新成功！" };

  } catch (e) {
    return { success: false, message: "更新失敗: " + e.toString() };
  } finally {
    lock.releaseLock();
  }
}