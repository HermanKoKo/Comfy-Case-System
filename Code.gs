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

/**
 * 智慧分頁選取器
 */
function getSheetHelper(sheetName) {
  let ss;
  // 路由邏輯：Client基本資料去新DB，其餘去舊DB
  if (sheetName === CONFIG.SHEETS.CLIENT) {
    try {
      ss = SpreadsheetApp.openById(CONFIG.CLIENT_DB_ID);
    } catch (e) {
      throw new Error("無法連接個案核心資料庫 (Client DB)。");
    }
  } else {
    ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  }

  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    if (sheetName !== CONFIG.SHEETS.CLIENT) {
       return ss.insertSheet(sheetName);
    }
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
 * 取得系統人員與項目清單 - ★★★ 已優化：增加快取 ★★★
 */
function getSystemStaff() {
  const cache = CacheService.getScriptCache();
  const cached = cache.get("SYSTEM_STAFF_DATA");
  if (cached) return JSON.parse(cached);

  try {
    // 這裡直接連 System DB，不需透過 helper 路由判斷 (節省一次判斷)
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const sheet = ss.getSheetByName(CONFIG.SHEETS.SYSTEM);
    if(!sheet) throw new Error("無 System 表");

    const data = sheet.getDataRange().getValues();
    const rows = data.slice(1);
    
    const result = {
      doctors: rows.map(r => r[0]).filter(String),
      nurses: rows.map(r => r[1]).filter(String),
      therapists: rows.map(r => r[2]).filter(String),
      trackingTypes: rows.map(r => r[3]).filter(String),
      maintItems: rows.map(r => r[4]).filter(String),
      allStaff: rows.map(r => r[5]).filter(String),
      treatmentItems: rows.map(r => r[6]).filter(String)
    };

    // 寫入快取，保存 6 小時 (21600 秒)
    cache.put("SYSTEM_STAFF_DATA", JSON.stringify(result), 21600);
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

/**
 * 取得個案總覽資料 - ★★★ 已優化：批次開啟 System 資料庫 ★★★
 */
function getCaseOverviewData(clientId) {
  try {
    if (!clientId) return [];
    const result = [];
    const targetId = String(clientId).trim();

    // 優化：只開啟一次 System DB，因為除了基本資料外，所有紀錄都在這裡
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    
    // 預先取得所有需要的 Sheets
    const docSheet = ss.getSheetByName(CONFIG.SHEETS.DOCTOR);
    const maintSheet = ss.getSheetByName(CONFIG.SHEETS.MAINTENANCE);
    const trackSheet = ss.getSheetByName(CONFIG.SHEETS.TRACKING);
    const treatSheet = ss.getSheetByName(CONFIG.SHEETS.TREATMENT);

    // 1. 醫師看診
    if (docSheet) {
      try {
        const data = docSheet.getDataRange().getValues();
        const idx = data[0].map(normalizeHeader).indexOf(normalizeHeader("個案編號"));
        const targetCol = idx > -1 ? idx : 1; 
        data.slice(1).forEach(row => {
          if (String(row[targetCol]).replace(/^'/, '').trim() === targetId) {
            result.push({
              id: row[0],
              date: formatDateForJSON(row[2]),
              category: 'doctor', categoryName: '醫師看診',
              doctor: row[3], nurse: row[4],
              s: row[5], o: row[6], a: row[7], p: row[8],
              nursingRecord: row[9], remark: row[10]
            });
          }
        });
      } catch (e) {}
    }

    // 2. 保養項目
    if (maintSheet) {
      try {
        const data = maintSheet.getDataRange().getValues();
        const idx = data[0].map(normalizeHeader).indexOf(normalizeHeader("個案編號"));
        const targetCol = idx > -1 ? idx : 1;
        data.slice(1).forEach(row => {
          if (String(row[targetCol]).replace(/^'/, '').trim() === targetId) {
            result.push({
              id: row[0],
              date: formatDateForJSON(row[2]),
              category: 'maintenance', categoryName: '保養項目',
              staff: row[3], item: row[4],
              bp: row[5], spo2: row[6], hr: row[7], temp: row[8],
              rr: row[9], remark: row[10]
            });
          }
        });
      } catch (e) {}
    }

    // 3. 個管追蹤
    if (trackSheet) {
      try {
        const data = trackSheet.getDataRange().getValues();
        const idx = data[0].map(normalizeHeader).indexOf(normalizeHeader("個案編號"));
        const targetCol = idx > -1 ? idx : 1;
        data.slice(1).forEach(row => {
          if (String(row[targetCol]).replace(/^'/, '').trim() === targetId) {
            result.push({
              id: row[0],
              date: formatDateForJSON(row[2]),
              category: 'tracking', categoryName: '個管追蹤',
              staff: row[3], type: row[4], content: row[5]
            });
          }
        });
      } catch (e) {}
    }

    // 4. 治療紀錄
    if (treatSheet) {
      try {
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
              category: 'treatment', categoryName: '物理治療',
              staff: row[idxStaff],
              item: (idxItem > -1) ? row[idxItem] : "",
              complaint: (idxComplaint > -1) ? row[idxComplaint] : "",
              content: (idxContent > -1) ? row[idxContent] : "",
              nextPlan: (idxNext > -1) ? row[idxNext] : "" 
            });
          }
        });
      } catch (e) {}
    }

    return result.sort((a, b) => new Date(b.date) - new Date(a.date));

  } catch (e) { throw new Error("取得總覽資料失敗: " + e.message); }
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

/**
 * 更新紀錄
 */
function updateRecord(type, formData) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) return { success: false, message: "系統忙碌中" };

  try {
    let sheetName = "";
    let expectedIdHeader = "紀錄id";
    if (type === 'treatment') { sheetName = CONFIG.SHEETS.TREATMENT; }
    else if (type === 'doctor') { sheetName = CONFIG.SHEETS.DOCTOR; }
    else if (type === 'maintenance') { sheetName = CONFIG.SHEETS.MAINTENANCE; }
    else if (type === 'tracking') { sheetName = CONFIG.SHEETS.TRACKING; expectedIdHeader = "追蹤ID"; }

    const sheet = getSheetHelper(sheetName);
    const data = sheet.getDataRange().getValues();
    if (data.length < 2) return { success: false, message: "資料表為空" };

    const headers = data[0];
    let idColIndex = -1;
    for (let c = 0; c < headers.length; c++) {
      if (normalizeHeader(headers[c]) === normalizeHeader(expectedIdHeader)) { idColIndex = c; break; }
    }
    if (idColIndex === -1 && type === 'treatment') {
       for (let c = 0; c < headers.length; c++) { if (String(headers[c]).toLowerCase().includes("id")) { idColIndex = c; break; } }
    }
    if (idColIndex === -1 && type !== 'treatment') idColIndex = 0;
    if (idColIndex === -1) throw new Error("找不到 ID 欄位");

    let targetId = String(formData.record_id).trim();
    let rowIndex = -1;
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][idColIndex]).trim() === targetId) { rowIndex = i + 1; break; }
    }

    if (rowIndex === -1) return { success: false, message: "找不到該筆資料" };

    const getCol = (name) => {
       for(let k=0; k<headers.length; k++) { if(normalizeHeader(headers[k]) === normalizeHeader(name)) return k + 1; }
       return -1;
    };

    if (type === 'treatment') {
       const cols = { date: getCol("治療日期"), therapist: getCol("執行治療師"), item: getCol("治療項目"), complaint: getCol("當日主訴"), content: getCol("治療內容"), next: getCol("備註/下次治療") };
       if (cols.date > 0) sheet.getRange(rowIndex, cols.date).setValue(formData.date);
       if (cols.therapist > 0) sheet.getRange(rowIndex, cols.therapist).setValue(formData.therapist);
       if (cols.item > 0) sheet.getRange(rowIndex, cols.item).setValue(formData.item);
       if (cols.complaint > 0) sheet.getRange(rowIndex, cols.complaint).setValue(formData.complaint);
       if (cols.content > 0) sheet.getRange(rowIndex, cols.content).setValue(formData.content);
       if (cols.next > 0) sheet.getRange(rowIndex, cols.next).setValue(formData.nextPlan);
    } 
    else if (type === 'doctor') {
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
       sheet.getRange(rowIndex, 10).setValue(formData.rr);
       sheet.getRange(rowIndex, 11).setValue(formData.remark);
    } 
    else if (type === 'tracking') {
       sheet.getRange(rowIndex, 3).setValue(formData.trackDate);
       sheet.getRange(rowIndex, 4).setValue(formData.trackStaff);
       sheet.getRange(rowIndex, 5).setValue(formData.trackType);
       sheet.getRange(rowIndex, 6).setValue(formData.content);
    }

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