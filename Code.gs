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
 * 核心工具層 (Api.gs)
 * ==========================================
 */

// ★★★ API 寫入輔助函式 (速度優化核心) ★★★
function appendWithAPI(spreadsheetId, sheetName, values) {
  const resource = { values: [values] };
  // valueInputOption: 'USER_ENTERED' 確保日期等格式會被 Sheets 自動解析
  Sheets.Spreadsheets.Values.append(resource, spreadsheetId, `${sheetName}!A1`, { valueInputOption: 'USER_ENTERED' });
}

function normalizeHeader(header) {
  return String(header).replace(/\s+/g, '').trim().toLowerCase();
}

// 傳統 Helper (用於 update 或特定讀取，保留相容性)
function getSheetHelper(sheetName) {
  let ss;
  if (sheetName === CONFIG.SHEETS.CLIENT) {
    if (!CACHED_SS_CLIENT) {
      try { CACHED_SS_CLIENT = SpreadsheetApp.openById(CONFIG.CLIENT_DB_ID); } 
      catch (e) { throw new Error("無法連接個案核心資料庫"); }
    }
    ss = CACHED_SS_CLIENT;
  } else {
    if (!CACHED_SS_SYSTEM) { CACHED_SS_SYSTEM = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID); }
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
 * 取得系統人員與項目清單 (已優化快取)
 */
function getSystemStaff() {
  const cache = CacheService.getScriptCache();
  // 設定 5 秒快取，兼顧效能與即時更新
  const cached = cache.get("SYSTEM_STAFF_DATA_V2");
  if (cached) return JSON.parse(cached);

  try {
    // 使用 API 讀取提升速度
    const range = `${CONFIG.SHEETS.SYSTEM}!A2:G`;
    const response = Sheets.Spreadsheets.Values.get(CONFIG.SPREADSHEET_ID, range);
    const data = response.values;
    
    if (!data || data.length === 0) return {};

    const result = {
      doctors: data.map(r => r[0]).filter(String),
      nurses: data.map(r => r[1]).filter(String),
      therapists: data.map(r => r[2]).filter(String),
      trackingTypes: data.map(r => r[3]).filter(String),
      maintItems: data.map(r => r[4]).filter(String),
      allStaff: data.map(r => r[5]).filter(String),
      treatmentItems: data.map(r => r[6]).filter(String)
    };

    cache.put("SYSTEM_STAFF_DATA_V2", JSON.stringify(result), 5);
    return result;

  } catch (e) {
    console.error(e);
    return {};
  }
}

/**
 * 取得治療項目清單 (保留介面，實際上由 getSystemStaff 統一處理)
 */
function getTreatmentItemsFromSystem() {
   const data = getSystemStaff();
   return data.treatmentItems || [];
}

/**
 * 取得治療師名單 (保留介面)
 */
function getTherapistList() {
   const data = getSystemStaff();
   return data.therapists || [];
}

/**
 * ==========================================
 * 資料寫入功能 (全面改用 API)
 * ==========================================
 */

// 通用儲存 (主要用於治療紀錄)
function saveData(sheetName, dataObj) {
  try {
    const targetSheetName = sheetName || CONFIG.SHEETS.CLIENT;
    
    // 為了確保標題順序正確，我們這裡還是得先讀一次 Header (或寫死)
    // 為了彈性，我們先讀第一列 Header (API 讀一行很快)
    const ssId = (targetSheetName === CONFIG.SHEETS.CLIENT) ? CONFIG.CLIENT_DB_ID : CONFIG.SPREADSHEET_ID;
    const headerResp = Sheets.Spreadsheets.Values.get(ssId, `${targetSheetName}!1:1`);
    const rawHeaders = headerResp.values ? headerResp.values[0] : [];
    
    if (rawHeaders.length === 0) throw new Error("資料表無標題列");

    const rowData = rawHeaders.map(rawH => {
        const cleanH = normalizeHeader(rawH);
        let val = '';
        for (let key in dataObj) {
            if (normalizeHeader(key) === cleanH) { val = dataObj[key]; break; }
        }
        
        if (cleanH === '紀錄id') return val || 'R' + new Date().getTime() + Math.floor(Math.random()*100);
        if (cleanH.includes('時間') || cleanH.includes('日期')) {
            if (cleanH === '追蹤日期' && val) return val;
            return val || Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
        }
        if (['電話', '身分證字號', '個案編號'].includes(cleanH)) return "'" + String(val || "");
        return val || '';
    });

    // ★ 使用 API 寫入
    appendWithAPI(ssId, targetSheetName, rowData);
    
    // 如果是新增個案，清除快取
    if (targetSheetName === CONFIG.SHEETS.CLIENT) {
        CacheService.getScriptCache().remove("ALL_CLIENT_CACHE_V2");
    }

    return { success: true, message: "資料已新增" };
  } catch (e) { throw new Error(e.message); }
}

function saveTrackingRecord(formObj) {
  try {
    if (!formObj.clientId) throw new Error("無個案編號");
    const now = new Date();
    
    // 欄位順序需對應 Case_Tracking: 
    // [追蹤ID, 個案編號, 追蹤日期, 追蹤人員, 追蹤項目, 追蹤內容, 建立時間]
    const rowData = [
      "TR" + now.getTime(),                
      "'" + formObj.clientId,
      formObj.trackDate,
      formObj.trackStaff,      
      formObj.trackType,       
      formObj.content,         
      Utilities.formatDate(now, Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss")
    ];
    
    appendWithAPI(CONFIG.SPREADSHEET_ID, CONFIG.SHEETS.TRACKING, rowData);
    return { success: true, message: "追蹤紀錄已新增" };
  } catch (e) { return { success: false, message: "儲存失敗: " + e.toString() }; }
}

function saveDoctorConsultation(formData) {
  try {
    if (!formData.clientId) throw new Error("無個案編號");
    const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy/MM/dd HH:mm:ss");
    
    // 對應 Doctor_Consultation
    const rowData = [
      "DOC" + new Date().getTime(),                   
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

    appendWithAPI(CONFIG.SPREADSHEET_ID, CONFIG.SHEETS.DOCTOR, rowData);
    return { success: true, message: "醫師看診紀錄儲存成功" };
  } catch (e) { return { success: false, message: "儲存失敗: " + e.toString() }; }
}

function saveMaintenanceRecord(data) {
  try {
    if (!data.clientId) throw new Error("無個案編號");
    // 對應 Health_Maintenance
    const rowData = [
      Utilities.getUuid(),
      "'" + data.clientId,
      data.date,
      data.staff,
      data.item,
      data.bp,
      data.spo2,
      data.hr,
      data.temp,
      data.rr,    
      data.remark,
      new Date()
    ];
    
    appendWithAPI(CONFIG.SPREADSHEET_ID, CONFIG.SHEETS.MAINTENANCE, rowData);
    return { success: true, message: "保養紀錄儲存成功！" };
  } catch (e) { return { success: false, message: "儲存失敗：" + e.toString() }; }
}

function createNewClient(data) {
  try {
    const now = new Date();
    const datePart = Utilities.formatDate(now, "GMT+8", "yyyyMMdd");
    
    // 為了取得最新編號，仍需讀取最後一列，這裡用簡單的 API 讀取
    const resp = Sheets.Spreadsheets.Values.get(CONFIG.CLIENT_DB_ID, `${CONFIG.SHEETS.CLIENT}!A:A`);
    const existingRows = resp.values ? resp.values.length : 1; 
    const suffix = existingRows.toString().padStart(3, '0'); // 因為標題佔一列，所以直接用 length 當序號即可 (row index + 1)
    const clientId = "CF" + datePart + suffix;

    let folderUrl = "";
    try { 
        const parentFolder = DriveApp.getFolderById(CONFIG.PARENT_FOLDER_ID);
        const folder = parentFolder.createFolder(clientId + "_" + data.name); 
        folderUrl = folder.getUrl(); 
    } catch (e) { folderUrl = "資料夾建立失敗"; }

    // Client_Basic_Info 欄位順序:
    // [個案編號, 姓名, 生日, 身分證, 電話, 性別, 緊急聯絡人, 緊急電話, 治療師, 慢性病, Folder, 建立日期]
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

    appendWithAPI(CONFIG.CLIENT_DB_ID, CONFIG.SHEETS.CLIENT, newRow);
    
    // ★ 清除全域個案快取，確保搜尋能搜到新個案
    CacheService.getScriptCache().remove("ALL_CLIENT_CACHE_V2");

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
 * ==========================================
 * 資料讀取功能 (快取與 BatchGet 優化)
 * ==========================================
 */

// ★★★ [優化] 取得所有個案 (Server Cache) ★★★
function getAllClientDataForCache() {
  const cache = CacheService.getScriptCache();
  const cachedData = cache.get("ALL_CLIENT_CACHE_V2");
  
  if (cachedData) {
    return JSON.parse(cachedData);
  }

  // 若無快取，則讀取
  try {
    const range = `${CONFIG.SHEETS.CLIENT}!A2:L`; // 讀到 L 欄即可
    const response = Sheets.Spreadsheets.Values.get(CONFIG.CLIENT_DB_ID, range);
    const rows = response.values || [];
    
    // 欄位映射: 0:ID, 1:Name, 2:Dob, 3:IdNo, 4:Phone, 5:Gender, 6:EmerName, 7:EmerPhone, 8:Therapist, 9:Chronic, 10:Folder
    const result = rows.map(row => ({
      id: row[0],
      name: row[1],
      dob: row[2],
      phone: row[4],
      therapist: row[8],
      fullJson: JSON.stringify({
        '個案編號': row[0], '姓名': row[1], '生日': row[2], '身分證字號': row[3],
        '電話': row[4], '性別': row[5], '緊急聯絡人': row[6], '緊急聯絡人電話': row[7],
        '負責治療師': row[8], '慢性病或特殊疾病': row[9], 'GoogleDrive資料夾連結': row[10]
      })
    })).reverse();

    // 存入快取 6 小時 (21600秒)
    // 注意：Cache 有大小限制 (100KB)，若資料過多可能要分段，目前先直接存
    try {
      cache.put("ALL_CLIENT_CACHE_V2", JSON.stringify(result), 21600);
    } catch(e) {
      console.warn("Cache too large, skipping.");
    }

    return result;
  } catch (e) {
    throw new Error("讀取個案資料庫失敗: " + e.message);
  }
}

// ★★★ [優化] 個案總覽 Batch Read ★★★
function getCaseOverviewData(clientId) {
  try {
    if (!clientId) return [];
    const targetId = String(clientId).trim();
    const result = [];
    
    // 定義四張表的 Range (預設讀取 A:Z，確保包含所有欄位)
    // 順序: [Doctor, Maintenance, Tracking, Treatment]
    const ranges = [
      `${CONFIG.SHEETS.DOCTOR}!A:M`,      
      `${CONFIG.SHEETS.MAINTENANCE}!A:K`, 
      `${CONFIG.SHEETS.TRACKING}!A:G`,    
      `${CONFIG.SHEETS.TREATMENT}!A:Z`    
    ];
    
    // API Batch Get (一次請求抓四張表)
    const response = Sheets.Spreadsheets.Values.batchGet(CONFIG.SPREADSHEET_ID, { ranges: ranges });
    const valueRanges = response.valueRanges;
    
    // 定義處理邏輯
    const processors = [
      { // Doctor
        cat: 'doctor', idIdx: 1, 
        process: (r, h) => ({
            id: r[0], date: formatDateForJSON(r[2]), category: 'doctor', categoryName: '醫師看診',
            doctor: r[3], nurse: r[4], s: r[5], o: r[6], a: r[7], p: r[8], nursingRecord: r[9], remark: r[10]
        })
      },
      { // Maintenance
        cat: 'maintenance', idIdx: 1,
        process: (r, h) => ({
            id: r[0], date: formatDateForJSON(r[2]), category: 'maintenance', categoryName: '保養項目',
            staff: r[3], item: r[4], bp: r[5], spo2: r[6], hr: r[7], temp: r[8], rr: r[9], remark: r[10]
        })
      },
      { // Tracking
        cat: 'tracking', idIdx: 1,
        process: (r, h) => ({
            id: r[0], date: formatDateForJSON(r[2]), category: 'tracking', categoryName: '個管追蹤',
            staff: r[3], type: r[4], content: r[5]
        })
      },
      { // Treatment (需動態找 Index)
        cat: 'treatment', idIdx: 1, 
        process: (r, h) => {
             // 簡易 Index 對照 (假設順序未變，若 Sheet 結構常變，建議用 loop find)
             // 這裡使用 headers.indexOf 的邏輯，但需注意 API 回傳的是 values，h 是第一列
             const cleanH = h.map(normalizeHeader);
             const getIdx = (name) => cleanH.indexOf(normalizeHeader(name));
             
             const idxDate = getIdx("治療日期");
             const idxStaff = getIdx("執行治療師");
             const idxItem = getIdx("治療項目");
             const idxC = getIdx("當日主訴");
             const idxCont = getIdx("治療內容");
             const idxNext = getIdx("備註/下次治療");

             return {
                 id: 'T-' + formatDateForJSON(r[idxDate]), 
                 date: formatDateForJSON(r[idxDate]),
                 category: 'treatment', categoryName: '物理治療',
                 // [邏輯修復] 檢查 Index 是否存在 (>-1)
                 staff: idxStaff > -1 ? r[idxStaff] : "",
                 item: idxItem > -1 ? r[idxItem] : "",
                 complaint: idxC > -1 ? r[idxC] : "",
                 content: idxCont > -1 ? r[idxCont] : "",
                 nextPlan: idxNext > -1 ? r[idxNext] : ""
             };
        }
      }
    ];

    // 處理資料
    valueRanges.forEach((sheetData, index) => {
        const values = sheetData.values;
        if (!values || values.length < 2) return; // 沒資料或只有標題

        const headers = values[0]; // 第一列為標題
        const processor = processors[index];
        
        // 從第二列開始遍歷
        for (let i = 1; i < values.length; i++) {
           const row = values[i];
           // 檢查個案編號 (通常在 index 1，從 headers 確認也可，這裡簡化處理)
           const rowClientId = String(row[processor.idIdx] || "").replace(/^'/, '').trim();
           
           if (rowClientId === targetId) {
               result.push(processor.process(row, headers));
           }
        }
    });

    return result.sort((a, b) => new Date(b.date) - new Date(a.date));

  } catch (e) { throw new Error("總覽讀取失敗(API): " + e.message); }
}

// 歷史紀錄相關 (維持原樣或改用 API 讀取單一表，這裡簡單處理，沿用 getSheetHelper)
// 若要極致優化，這三個也可改用 Sheets.Spreadsheets.Values.get
function getTrackingHistory(clientId) { return getHistoryGeneric(CONFIG.SHEETS.TRACKING, clientId, 'tracking'); }
function getMaintenanceHistory(clientId) { return getHistoryGeneric(CONFIG.SHEETS.MAINTENANCE, clientId, 'maintenance'); }

function getHistoryGeneric(sheetName, clientId, type) {
    try {
        const ssId = CONFIG.SPREADSHEET_ID;
        const response = Sheets.Spreadsheets.Values.get(ssId, `${sheetName}!A:Z`);
        const values = response.values;
        if (!values || values.length < 2) return [];

        const headers = values[0];
        const normalizedHeaders = headers.map(normalizeHeader);
        const idIdx = normalizedHeaders.indexOf(normalizeHeader("個案編號"));
        const targetId = String(clientId).trim();
        
        // 簡單 Filter
        const rawData = values.slice(1).filter(r => String(r[idIdx]||"").replace(/^'/, '').trim() === targetId);

        // Map 到物件
        if (type === 'maintenance') {
             return rawData.map(r => {
                 let obj = {};
                 headers.forEach((h, i) => obj[h] = r[i]);
                 return obj;
             }).reverse();
        } else if (type === 'tracking') {
             // Tracking 特定格式
             const idxId = headers.indexOf("追蹤ID");
             const idxDate = headers.indexOf("追蹤日期");
             const idxStaff = headers.indexOf("追蹤人員");
             const idxType = headers.indexOf("追蹤項目");
             const idxContent = headers.indexOf("追蹤內容");
             return rawData.map(r => ({
                 id: r[idxId], date: formatDateForJSON(r[idxDate]),
                 staff: r[idxStaff], type: r[idxType], content: r[idxContent]
             })).sort((a,b)=>new Date(b.date)-new Date(a.date));
        }
        return [];
    } catch(e) { return []; }
}

function getClientHistory(clientId, sheetName) {
  // 通用歷史紀錄 (Doctor, Treatment)
  try {
     const ssId = (sheetName === CONFIG.SHEETS.CLIENT) ? CONFIG.CLIENT_DB_ID : CONFIG.SPREADSHEET_ID;
     const response = Sheets.Spreadsheets.Values.get(ssId, `${sheetName}!A:Z`);
     const values = response.values;
     if (!values || values.length < 2) return [];

     const headers = values[0];
     const idIdx = headers.map(normalizeHeader).indexOf(normalizeHeader("個案編號"));
     const targetId = String(clientId).trim();
     
     const result = [];
     for(let i=1; i<values.length; i++) {
        if (String(values[i][idIdx]||"").replace(/^'/, '').trim() === targetId) {
            let obj = {};
            headers.forEach((h, idx) => obj[h] = values[i][idx]);
            result.push(obj);
        }
     }
     
     result.sort((a, b) => {
        const dA = a['看診日期'] || a['治療日期'] || a['日期'] || '1900-01-01';
        const dB = b['看診日期'] || b['治療日期'] || b['日期'] || '1900-01-01';
        return new Date(dB) - new Date(dA);
     });
     return result;
  } catch(e) { return []; }
}


/**
 * ==========================================
 * 更新功能 (Update) - 仍使用傳統方法較安全 (需精確定位 Row)
 * ==========================================
 */
function updateRecord(type, formData) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) return { success: false, message: "系統忙碌中" };

  try {
    let sheetName = "";
    let expectedIdHeader = "紀錄id";
    if (type === 'treatment') sheetName = CONFIG.SHEETS.TREATMENT;
    else if (type === 'doctor') sheetName = CONFIG.SHEETS.DOCTOR;
    else if (type === 'maintenance') sheetName = CONFIG.SHEETS.MAINTENANCE;
    else if (type === 'tracking') { sheetName = CONFIG.SHEETS.TRACKING; expectedIdHeader = "追蹤ID"; }

    const sheet = getSheetHelper(sheetName);
    const dataRange = sheet.getDataRange();
    const values = dataRange.getValues();
    const headers = values[0].map(normalizeHeader);
    
    let idColIndex = headers.indexOf(normalizeHeader(expectedIdHeader));
    if (idColIndex === -1 && type === 'treatment') idColIndex = headers.findIndex(h => h.includes("id"));
    if (idColIndex === -1 && type !== 'treatment') idColIndex = 0;
    
    const targetId = String(formData.record_id).trim();
    let rowIndex = -1;
    for (let i = 1; i < values.length; i++) {
      if (String(values[i][idColIndex]).trim() === targetId) { rowIndex = i; break; }
    }

    if (rowIndex === -1) return { success: false, message: "找不到該筆資料" };

    const rowData = values[rowIndex];
    const setCol = (headerName, value) => {
      const idx = headers.indexOf(normalizeHeader(headerName));
      if (idx > -1) rowData[idx] = value;
    };

    // 更新邏輯
    if (type === 'treatment') {
       setCol("治療日期", formData.date); setCol("執行治療師", formData.therapist);
       setCol("治療項目", formData.item); setCol("當日主訴", formData.complaint);
       setCol("治療內容", formData.content); setCol("備註/下次治療", formData.nextPlan);
    } else if (type === 'doctor') {
       setCol("看診日期", formData.date); setCol("看診醫師", formData.doctor);
       setCol("護理師", formData.nurse); setCol("S_主訴", formData.complaint);
       setCol("O_客觀檢查", formData.objective); setCol("A_診斷", formData.diagnosis);
       setCol("P_治療計劃", formData.plan); setCol("護理紀錄", formData.nursingRecord);
       setCol("備註", formData.remark);
    } else if (type === 'maintenance') {
       setCol("保養日期", formData.date); setCol("執行人員", formData.staff);
       setCol("保養項目", formData.item); setCol("血壓", formData.bp);
       setCol("血氧", formData.spo2); setCol("心律", formData.hr);
       setCol("體溫", formData.temp); setCol("呼吸速率", formData.rr);
       setCol("備註", formData.remark);
    } else if (type === 'tracking') {
       setCol("追蹤日期", formData.trackDate); setCol("追蹤人員", formData.trackStaff);
       setCol("追蹤項目", formData.trackType); setCol("追蹤內容", formData.content);
    }

    sheet.getRange(rowIndex + 1, 1, 1, rowData.length).setValues([rowData]);
    return { success: true, message: "資料更新成功！" };
  } catch (e) { return { success: false, message: "更新失敗: " + e.toString() }; } finally { lock.releaseLock(); }
}

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
    
    // 更新欄位 (Columns B~J)
    sheet.getRange(rowIndex, 2).setValue(data.name);      
    sheet.getRange(rowIndex, 3).setValue(data.dob);       
    sheet.getRange(rowIndex, 4).setValue(data.idNo);      
    sheet.getRange(rowIndex, 5).setValue("'" + data.phone); 
    sheet.getRange(rowIndex, 6).setValue(data.gender);    
    sheet.getRange(rowIndex, 7).setValue(data.emerName);  
    sheet.getRange(rowIndex, 8).setValue("'" + data.emerPhone); 
    sheet.getRange(rowIndex, 9).setValue(data.therapist); 
    sheet.getRange(rowIndex, 10).setValue(data.chronic);  
    
    // 清除快取
    CacheService.getScriptCache().remove("ALL_CLIENT_CACHE_V2");

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

/**
 * ==========================================
 * 其他輔助功能
 * ==========================================
 */
function getClientImages(clientId) {
  try {
    // 這裡還是需要 Folder URL，用 API 讀取 Client DB 比較快
    const resp = Sheets.Spreadsheets.Values.get(CONFIG.CLIENT_DB_ID, `${CONFIG.SHEETS.CLIENT}!A:K`);
    const rows = resp.values || [];
    let folderUrl = "";
    const targetId = String(clientId).trim();
    
    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][0]).replace(/^'/, '').trim() === targetId) {
        folderUrl = rows[i][10]; // K欄
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

function uploadClientImage(clientId, fileData, fileName, mimeType, remark) {
  try {
    // 1. 取得 Folder ID
    const resp = Sheets.Spreadsheets.Values.get(CONFIG.CLIENT_DB_ID, `${CONFIG.SHEETS.CLIENT}!A:K`);
    const rows = resp.values || [];
    let folderUrl = "";
    const targetId = String(clientId).trim();
    for (let i = 1; i < rows.length; i++) {
        if (String(rows[i][0]).replace(/^'/, '').trim() === targetId) { folderUrl = rows[i][10]; break; }
    }
    if (!folderUrl) throw new Error("找不到個案資料夾");
    const folderIdMatch = folderUrl.match(/[-\w]{25,}/);
    if (!folderIdMatch) throw new Error("資料夾 ID 解析失敗");
    
    // 2. 上傳檔案 (DriveApp)
    const folder = DriveApp.getFolderById(folderIdMatch[0]);
    const blob = Utilities.newBlob(Utilities.base64Decode(fileData), mimeType, fileName);
    const file = folder.createFile(blob);
    const fileUrl = file.getUrl();
    
    // 3. 寫入紀錄 (Image_Gallery) using API
    const rowData = [
      "IMG" + new Date().getTime(),
      "'" + clientId,
      Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm"),
      fileName,
      fileUrl,
      remark || ""
    ];
    appendWithAPI(CONFIG.SPREADSHEET_ID, CONFIG.SHEETS.IMAGE, rowData);

    return { success: true, message: "上傳成功" };
  } catch (e) { return { success: false, message: "上傳失敗: " + e.toString() }; }
}

function formatDateForJSON(dateVal) {
  if (!dateVal) return "";
  // API 讀取的日期有時是字串有時是數字，若格式正確則直接回傳
  if (String(dateVal).match(/^\d{4}-\d{2}-\d{2}/)) return String(dateVal).substring(0, 10);
  if (dateVal instanceof Date) return Utilities.formatDate(dateVal, Session.getScriptTimeZone(), "yyyy-MM-dd");
  return String(dateVal);
}