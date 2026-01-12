/**
 * ==========================================
 * 設定檔 (Config.gs)
 * ==========================================
 */
const CONFIG = {
  // 紀錄與設定資料庫 (System DB)
  SPREADSHEET_ID: '1LMhlQGyXNXq9Teqm0_W0zU9NbQlVCHKLDL0mSOiDomc', 
  
  // 個案核心資料庫 (Client DB)
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

// 在 CONFIG 下方加入
let CACHED_SS_CLIENT = null;
let CACHED_SS_SYSTEM = null;

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

function normalizeHeader(header) {
  return String(header).replace(/\s+/g, '').trim().toLowerCase();
}

function getSheetHelper(sheetName) {
  let ss;
  // 判斷是否為 Client DB
  if (sheetName === CONFIG.SHEETS.CLIENT) {
    if (!CACHED_SS_CLIENT) {
      try {
        CACHED_SS_CLIENT = SpreadsheetApp.openById(CONFIG.CLIENT_DB_ID);
      } catch (e) { throw new Error("無法連接個案核心資料庫"); }
    }
    ss = CACHED_SS_CLIENT;
  } else {
    // System DB
    if (!CACHED_SS_SYSTEM) {
      CACHED_SS_SYSTEM = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    }
    ss = CACHED_SS_SYSTEM;
  }

  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    if (sheetName !== CONFIG.SHEETS.CLIENT) return ss.insertSheet(sheetName);
    throw new Error("找不到工作表: " + sheetName);
  }
  return sheet;
}

/**
 * 取得治療師下拉選單 (System C欄) - 增加快取
 */
function getTherapistList() {
  // 嘗試從快取讀取
  const cache = CacheService.getScriptCache();
  const cached = cache.get("THERAPIST_LIST");
  if (cached) return JSON.parse(cached);

  try {
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID); // 直接連線 System DB
    const sheet = ss.getSheetByName("system"); // 注意：原代碼是小寫 system
    if (!sheet) return [];

    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return [];

    const data = sheet.getRange(2, 3, lastRow - 1, 1).getValues();
    const therapists = data.flat().filter(n => n && n.toString().trim() !== "");
    
    // 寫入快取 (存活 6 小時)
    cache.put("THERAPIST_LIST", JSON.stringify(therapists), 21600);
    
    return therapists;
  } catch (e) {
    console.error("getTherapistList error: " + e.toString());
    return [];
  }
}

/**
 * 搜尋功能
 */
function searchClient(keyword) {
  try {
    const sheet = getSheetHelper(CONFIG.SHEETS.CLIENT);
    // 使用 DisplayValues 以符合視覺上的搜尋 (如日期格式)
    const data = sheet.getDataRange().getDisplayValues(); 
    const results = [];
    
    const query = String(keyword).replace(/\s+/g, '').toLowerCase();
    if (!query) return [];
    
    // 從第 1 列 (Index 1) 開始，跳過標題
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      // 優化：預先處理比對字串
      const id = String(row[0]||'').replace(/^'/, '').replace(/\s+/g, '').toLowerCase();
      const name = String(row[1]||'').replace(/\s+/g, '').toLowerCase();
      const phone = String(row[4]||'').replace(/^'/, '').replace(/\s+/g, '').toLowerCase();
      
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
          '負責治療師': row[8], // Index 8
          '慢性病或特殊疾病': row[9],
          'GoogleDrive資料夾連結': row[10],
          '建立日期': row[11]
        });
      }
    }
    return results;
  } catch (e) { throw new Error(e.message); }
}

/**
 * 通用資料儲存
 */
function saveData(sheetName, dataObj) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000); 
    
    const targetSheetName = sheetName || CONFIG.SHEETS.CLIENT;
    const sheet = getSheetHelper(targetSheetName);
    const ss = sheet.getParent();

    // 取得標題並建立映射
    const rawHeaders = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const headerMap = {}; 
    rawHeaders.forEach((h, i) => headerMap[normalizeHeader(h)] = i);
    
    // 檢查是否有個案編號
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
        // 尋找對應的 Key
        for (let key in dataObj) {
            if (normalizeHeader(key) === cleanH) { val = dataObj[key]; break; }
        }
        
        // 自動生成欄位邏輯
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
 * 取得系統人員與項目清單 - [已修復] 移除長快取以支援即時更新
 */
function getSystemStaff() {
  // 若希望即時性高，建議暫時移除 Cache 或設為極短時間 (例如 5 秒)
  // 這裡為了效能保留 5 秒快取，避免短時間重複 request，但能確保重整頁面後拿到新資料
  const cache = CacheService.getScriptCache();
  const cached = cache.get("SYSTEM_STAFF_DATA_V2");
  if (cached) return JSON.parse(cached);

  try {
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const sheet = ss.getSheetByName(CONFIG.SHEETS.SYSTEM);
    if(!sheet) throw new Error("無 System 表");

    const lastRow = sheet.getLastRow();
    // 防呆：如果沒資料
    if (lastRow < 2) return {};

    // 一次讀取所有資料範圍 (A2:G)
    const data = sheet.getRange(2, 1, lastRow - 1, 7).getValues();
    
    // 整理資料：過濾掉空字串
    const result = {
      doctors: data.map(r => r[0]).filter(String),        // A欄: 醫師
      nurses: data.map(r => r[1]).filter(String),         // B欄: 護理師
      therapists: data.map(r => r[2]).filter(String),     // C欄: 治療師
      trackingTypes: data.map(r => r[3]).filter(String),  // D欄: 追蹤項目
      maintItems: data.map(r => r[4]).filter(String),     // E欄: 保養項目
      allStaff: data.map(r => r[5]).filter(String),       // F欄: 所有人員
      treatmentItems: data.map(r => r[6]).filter(String)  // G欄: 治療項目
    };

    // 快取設為 5 秒，確保 F5 重整後能拿到 Sheet 的最新變更
    cache.put("SYSTEM_STAFF_DATA_V2", JSON.stringify(result), 5);
    return result;

  } catch (e) {
    console.error(e);
    return {};
  }
}

/**
 * 儲存個管追蹤紀錄
 */
function saveTrackingRecord(formObj) {
  try {
    const sheet = getSheetHelper(CONFIG.SHEETS.TRACKING);
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(["追蹤ID", "個案編號", "追蹤日期", "追蹤人員", "追蹤項目", "追蹤內容", "建立時間"]);
    }
    const ss = sheet.getParent();
    if (!formObj.clientId) throw new Error("無個案編號");

    const now = new Date();
    const uniqueId = "TR" + now.getTime();
    
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
    
    // 固定欄位索引 (若標題變動可改為 map 搜尋)
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
      formData.nursingRecord,     
      formData.remark,            
      "",                         
      timestamp                   
    ];

    sheet.appendRow(rowData);
    return { success: true, message: "醫師看診紀錄儲存成功" };
  } catch (e) { return { success: false, message: "儲存失敗: " + e.toString() }; }
}

/**
 * 建立新個案
 */
function createNewClient(data) {
  try {
    const sheet = getSheetHelper(CONFIG.SHEETS.CLIENT);
    const now = new Date();
    const datePart = Utilities.formatDate(now, "GMT+8", "yyyyMMdd");
    
    const lastRow = sheet.getLastRow();
    const suffix = (lastRow + 1).toString().padStart(3, '0');
    const clientId = "CF" + datePart + suffix;

    let folderUrl = "";
    try { 
        const parentFolder = DriveApp.getFolderById(CONFIG.PARENT_FOLDER_ID);
        const folder = parentFolder.createFolder(clientId + "_" + data.name); 
        folderUrl = folder.getUrl(); 
    } catch (e) { folderUrl = "資料夾建立失敗"; }

    const newRow = [
      clientId,           
      data.name,          
      data.dob,           
      data.idNo,          
      "'" + data.phone,   
      data.gender,        
      data.emerName,      
      "'" + data.emerPhone,
      data.therapist || "",
      data.chronic,       
      folderUrl,          
      now                 
    ];

    sheet.appendRow(newRow);

    const fullData = {
        '個案編號': clientId, '姓名': data.name, '生日': data.dob, '身分證字號': data.idNo,
        '電話': data.phone, '性別': data.gender, '緊急聯絡人': data.emerName, 
        '緊急聯絡人電話': data.emerPhone, 
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
    const sheet = getSheetHelper(sheetName);
    const data = sheet.getDataRange().getDisplayValues(); // 使用 DisplayValues 保持日期格式
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
      data.rr,    // J欄: 呼吸
      data.remark,// K欄: 備註
      new Date()
    ];
    
    sheet.appendRow(newRow);
    return { success: true, message: "保養紀錄儲存成功！" };
  } catch (e) { return { success: false, message: "儲存失敗：" + e.toString() }; }
}

// [新增] 為了前端「秒開搜尋」，一次回傳所有個案的精簡清單
// 這樣前端就不需要每次打字都問後端
function getAllClientDataForCache() {
  const sheet = getSheetHelper(CONFIG.SHEETS.CLIENT);
  const data = sheet.getDataRange().getDisplayValues(); // 使用 DisplayValues 保持日期與電話格式
  
  // 移除標題列
  const headers = data.shift(); 
  
  // 為了縮小傳輸量，我們只回傳前端搜尋和列表顯示需要的欄位
  // 假設欄位順序：0:ID, 1:姓名, 2:生日, 3:身分證, 4:電話, ... 8:治療師, 9:慢性病, 10:Folder, 11:Date
  return data.map(row => ({
    id: row[0],
    name: row[1],
    dob: row[2],
    phone: row[4],
    therapist: row[8],
    // 把整列資料存成 JSON 字串，前端點擊時直接解開，不用再 fetch 一次
    fullJson: JSON.stringify({
      '個案編號': row[0], '姓名': row[1], '生日': row[2], '身分證字號': row[3],
      '電話': row[4], '性別': row[5], '緊急聯絡人': row[6], '緊急聯絡人電話': row[7],
      '負責治療師': row[8], '慢性病或特殊疾病': row[9], 'GoogleDrive資料夾連結': row[10]
    })
  })).reverse(); // 讓最新建立的個案排在最前面
}

// [優化] getCaseOverviewData: 使用 Promise.all 的概念 (在 GAS 裡是順序執行，但減少物件建立開銷)
function getCaseOverviewData(clientId) {
  try {
    if (!clientId) return [];
    const targetId = String(clientId).trim();
    const result = [];
    
    // 技巧：只開啟 Spreadsheet 物件一次
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    
    // 定義要讀取的 Sheets 與對應的 Category
    const configs = [
      { name: CONFIG.SHEETS.DOCTOR, cat: 'doctor', idHeader: '個案編號', mapFunc: (r, h) => ({
          id: r[0], date: formatDateForJSON(r[2]), category: 'doctor', categoryName: '醫師看診',
          doctor: r[3], nurse: r[4], s: r[5], o: r[6], a: r[7], p: r[8], nursingRecord: r[9], remark: r[10]
      })},
      { name: CONFIG.SHEETS.MAINTENANCE, cat: 'maintenance', idHeader: '個案編號', mapFunc: (r, h) => ({
          id: r[0], date: formatDateForJSON(r[2]), category: 'maintenance', categoryName: '保養項目',
          staff: r[3], item: r[4], bp: r[5], spo2: r[6], hr: r[7], temp: r[8], rr: r[9], remark: r[10]
      })},
      { name: CONFIG.SHEETS.TRACKING, cat: 'tracking', idHeader: '個案編號', mapFunc: (r, h) => ({
          id: r[0], date: formatDateForJSON(r[2]), category: 'tracking', categoryName: '個管追蹤',
          staff: r[3], type: r[4], content: r[5]
      })},
      { name: CONFIG.SHEETS.TREATMENT, cat: 'treatment', idHeader: '個案編號', mapFunc: (r, h) => {
          // 治療紀錄欄位較多，動態抓取 Index (優化：可以在外部定義好 Index，這裡簡化處理)
          const idxDate = h.indexOf("治療日期");
          const idxStaff = h.indexOf("執行治療師");
          const idxItem = h.indexOf("治療項目");
          const idxC = h.indexOf("當日主訴");
          const idxCont = h.indexOf("治療內容");
          const idxNext = h.indexOf("備註/下次治療");
          return {
              id: 'T-' + formatDateForJSON(r[idxDate]), date: formatDateForJSON(r[idxDate]),
              category: 'treatment', categoryName: '物理治療',
              staff: r[idxStaff], item: r[idxItem]>-1?r[idxItem]:"", complaint: r[idxC]>-1?r[idxC]:"",
              content: r[idxCont]>-1?r[idxCont]:"", nextPlan: r[idxNext]>-1?r[idxNext]:""
          };
      }}
    ];

    // 迴圈讀取資料
    configs.forEach(cfg => {
      const sheet = ss.getSheetByName(cfg.name);
      if (!sheet) return;
      
      const data = sheet.getDataRange().getValues();
      if (data.length <= 1) return;
      
      const headers = data[0].map(normalizeHeader);
      const rawHeaders = data[0]; // 保留原始 Header 給 Treatment 用
      let idIdx = headers.indexOf(normalizeHeader(cfg.idHeader));
      if (idIdx === -1) idIdx = 1; // Fallback

      // 記憶體內篩選
      for (let i = 1; i < data.length; i++) {
        if (String(data[i][idIdx]).replace(/^'/, '').trim() === targetId) {
           result.push(cfg.mapFunc(data[i], rawHeaders));
        }
      }
    });

    return result.sort((a, b) => new Date(b.date) - new Date(a.date));
  } catch (e) { throw new Error("總覽讀取失敗: " + e.message); }
}

/**
 * 影像功能 (保持原邏輯：直接讀 Drive)
 */
function getClientImages(clientId) {
  try {
    const clientSheet = getSheetHelper(CONFIG.SHEETS.CLIENT);
    const clientData = clientSheet.getDataRange().getDisplayValues();
    let folderUrl = "";
    const targetId = String(clientId).replace(/^'/, '').trim();
    
    for (let i = 1; i < clientData.length; i++) {
      if (String(clientData[i][0]).replace(/^'/, '').trim() === targetId) {
        folderUrl = clientData[i][10]; // Index 10
        break;
      }
    }
    
    if (!folderUrl) return { success: true, images: [] };
    const idMatch = folderUrl.match(/[-\w]{25,}/);
    if (!idMatch) return { success: true, images: [] };
    
    const folder = DriveApp.getFolderById(idMatch[0]);
    const files = folder.getFiles();
    const imageList = [];
    
    while (files.hasNext()) {
      const file = files.next();
      if (file.getMimeType().indexOf('image/') === 0) {
        const fileId = file.getId();
        imageList.push({
          id: fileId,
          name: file.getName(),
          url: file.getUrl(),
          thumbnail: "https://lh3.googleusercontent.com/d/" + fileId + "=s400",
          date: Utilities.formatDate(file.getDateCreated(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm"),
          remark: ""
        });
      }
    }
    return { success: true, images: imageList.sort((a, b) => new Date(b.date) - new Date(a.date)) };
  } catch (e) { return { success: false, message: "讀取影像失敗: " + e.toString() }; }
}

/**
 * 上傳影像
 */
function uploadClientImage(clientId, fileData, fileName, mimeType, remark) {
  try {
    const clientSheet = getSheetHelper(CONFIG.SHEETS.CLIENT);
    const clientData = clientSheet.getDataRange().getDisplayValues();
    let folderUrl = "";
    const targetId = String(clientId).replace(/^'/, '').trim();
    
    for (let i = 1; i < clientData.length; i++) {
      if (String(clientData[i][0]).replace(/^'/, '').trim() === targetId) { 
          folderUrl = clientData[i][10]; 
          break; 
      }
    }
    
    if (!folderUrl) throw new Error("找不到個案資料夾 (資料表無連結)");
    const folderIdMatch = folderUrl.match(/[-\w]{25,}/);
    if (!folderIdMatch) throw new Error("資料夾 ID 解析失敗");
    
    const folder = DriveApp.getFolderById(folderIdMatch[0]);
    const blob = Utilities.newBlob(Utilities.base64Decode(fileData), mimeType, fileName);
    const file = folder.createFile(blob);
    const fileUrl = file.getUrl();
    
    // 寫入上傳紀錄 (回舊DB)
    let imgSheet = getSheetHelper(CONFIG.SHEETS.IMAGE);
    const ss = imgSheet.getParent(); 
    if (imgSheet.getLastRow() === 0) {
      imgSheet.appendRow(["影像ID", "個案編號", "上傳日期", "檔案名稱", "GoogleDrive檔案連結", "備註"]);
    }
    
    imgSheet.appendRow([
      "IMG" + new Date().getTime(),
      "'" + clientId,
      Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), "yyyy-MM-dd HH:mm"),
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
 * 取得治療項目清單 (System G欄) - 增加快取
 */
function getTreatmentItemsFromSystem() {
  const cache = CacheService.getScriptCache();
  const cached = cache.get("TREATMENT_ITEMS");
  if (cached) return JSON.parse(cached);

  try {
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const sheet = ss.getSheetByName(CONFIG.SHEETS.SYSTEM);
    if(!sheet) return [];
    
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return [];
    
    const values = sheet.getRange(2, 7, lastRow - 1).getValues();
    const items = values.flat().filter(item => item !== "" && item != null);
    
    cache.put("TREATMENT_ITEMS", JSON.stringify(items), 21600);
    return items;
  } catch(e) { return []; }
}

function updateRecord(type, formData) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) return { success: false, message: "系統忙碌中" };

  try {
    // 1. 設定 Sheet 與 ID 欄位
    let sheetName = "";
    let expectedIdHeader = "紀錄id";
    if (type === 'treatment') sheetName = CONFIG.SHEETS.TREATMENT;
    else if (type === 'doctor') sheetName = CONFIG.SHEETS.DOCTOR;
    else if (type === 'maintenance') sheetName = CONFIG.SHEETS.MAINTENANCE;
    else if (type === 'tracking') { sheetName = CONFIG.SHEETS.TRACKING; expectedIdHeader = "追蹤ID"; }

    const sheet = getSheetHelper(sheetName);
    
    // 2. 讀取全部資料 (只讀一次！)
    const dataRange = sheet.getDataRange();
    const values = dataRange.getValues(); // 取得原始值
    const headers = values[0].map(normalizeHeader);
    
    // 3. 尋找 ID 欄位 Index
    let idColIndex = headers.indexOf(normalizeHeader(expectedIdHeader));
    if (idColIndex === -1 && type === 'treatment') {
       idColIndex = headers.findIndex(h => h.includes("id"));
    }
    if (idColIndex === -1 && type !== 'treatment') idColIndex = 0;
    if (idColIndex === -1) throw new Error("找不到 ID 欄位");

    // 4. 尋找目標 Row
    const targetId = String(formData.record_id).trim();
    let rowIndex = -1;
    for (let i = 1; i < values.length; i++) {
      if (String(values[i][idColIndex]).trim() === targetId) {
        rowIndex = i; // 這是 Array index，對應 Sheet Row 是 i + 1
        break;
      }
    }

    if (rowIndex === -1) return { success: false, message: "找不到該筆資料" };

    // 5. 準備更新資料 (在記憶體中操作 Array)
    const rowData = values[rowIndex]; // 取得該列舊資料
    
    // 建立簡易映射函式
    const setCol = (headerName, value) => {
      const idx = headers.indexOf(normalizeHeader(headerName));
      if (idx > -1) rowData[idx] = value;
    };

    // 依照類型更新 Array
    if (type === 'treatment') {
       setCol("治療日期", formData.date);
       setCol("執行治療師", formData.therapist);
       setCol("治療項目", formData.item);
       setCol("當日主訴", formData.complaint);
       setCol("治療內容", formData.content);
       setCol("備註/下次治療", formData.nextPlan);
    } 
    else if (type === 'doctor') {
       setCol("看診日期", formData.date);
       setCol("看診醫師", formData.doctor);
       setCol("護理師", formData.nurse);
       setCol("S_主訴", formData.complaint);
       setCol("O_客觀檢查", formData.objective);
       setCol("A_診斷", formData.diagnosis);
       setCol("P_治療計劃", formData.plan);
       setCol("護理紀錄", formData.nursingRecord);
       setCol("備註", formData.remark);
    } 
    else if (type === 'maintenance') {
       setCol("保養日期", formData.date);
       setCol("執行人員", formData.staff);
       setCol("保養項目", formData.item);
       setCol("血壓", formData.bp);
       setCol("血氧", formData.spo2);
       setCol("心律", formData.hr);
       setCol("體溫", formData.temp);
       setCol("呼吸速率", formData.rr);
       setCol("備註", formData.remark);
    } 
    else if (type === 'tracking') {
       setCol("追蹤日期", formData.trackDate);
       setCol("追蹤人員", formData.trackStaff);
       setCol("追蹤項目", formData.trackType);
       setCol("追蹤內容", formData.content);
    }

    // 6. 一次性寫回整列 (只寫一次！)
    // rowIndex + 1 是 sheet 列號, 1 是起始欄, 1 是行數, rowData.length 是總欄數
    sheet.getRange(rowIndex + 1, 1, 1, rowData.length).setValues([rowData]);

    return { success: true, message: "資料更新成功！" };
  } catch (e) { return { success: false, message: "更新失敗: " + e.toString() }; } finally { lock.releaseLock(); }
}

/**
 * 更新個案基本資料
 */
function updateClientBasicInfo(data) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const sheet = getSheetHelper(CONFIG.SHEETS.CLIENT);
    const rows = sheet.getDataRange().getDisplayValues();
    let rowIndex = -1;
    const targetId = String(data.clientId).trim();
    
    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][0]).replace(/^'/, '').trim() === targetId) { rowIndex = i + 1; break; }
    }
    
    if (rowIndex === -1) throw new Error("找不到此個案編號");
    
    sheet.getRange(rowIndex, 2).setValue(data.name);      
    sheet.getRange(rowIndex, 3).setValue(data.dob);       
    sheet.getRange(rowIndex, 4).setValue(data.idNo);      
    sheet.getRange(rowIndex, 5).setValue("'" + data.phone); 
    sheet.getRange(rowIndex, 6).setValue(data.gender);    
    sheet.getRange(rowIndex, 7).setValue(data.emerName);  
    sheet.getRange(rowIndex, 8).setValue("'" + data.emerPhone); 
    sheet.getRange(rowIndex, 9).setValue(data.therapist); 
    sheet.getRange(rowIndex, 10).setValue(data.chronic);  
    
    return {
      success: true, message: "基本資料更新成功",
      updatedData: {
        '個案編號': data.clientId, '姓名': data.name, '生日': data.dob,
        '身分證字號': data.idNo, '電話': data.phone, '性別': data.gender,
        '緊急聯絡人': data.emerName, '緊急聯絡人電話': data.emerPhone,
        '負責治療師': data.therapist, '慢性病或特殊疾病': data.chronic,
        '狀態': 'Active'
      }
    };
  } catch (e) { return { success: false, message: "更新失敗: " + e.toString() }; } finally { lock.releaseLock(); }
}

