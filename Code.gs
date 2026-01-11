/**
 * ==========================================
 * 設定檔 (Config.gs)
 * ==========================================
 */
const CONFIG = {
  SPREADSHEET_ID: '1LMhlQGyXNXq9Teqm0_W0zU9NbQlVCHKLDL0mSOiDomc', 
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

// ★ 優化：Spreadsheet 物件快取，避免重複 openById
var _ssCache = {};

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
 * ★ 優化核心：智慧分頁選取器 (含快取機制)
 */
function getSheetHelper(sheetName) {
  const isClientDB = (sheetName === CONFIG.SHEETS.CLIENT);
  const targetId = isClientDB ? CONFIG.CLIENT_DB_ID : CONFIG.SPREADSHEET_ID;
  
  // 檢查快取
  let ss = _ssCache[targetId];
  if (!ss) {
    try {
      ss = SpreadsheetApp.openById(targetId);
      _ssCache[targetId] = ss; // 存入快取
    } catch (e) {
      throw new Error("無法連接資料庫，ID: " + targetId);
    }
  }

  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    if (!isClientDB) return ss.insertSheet(sheetName);
    throw new Error("找不到工作表: " + sheetName);
  }
  return sheet;
}

/**
 * ★ 優化：一次取得所有系統初始化資料 (取代多次 Server Call)
 */
function getSystemInitData() {
  try {
    // 讀取 System 表
    const sheet = getSheetHelper(CONFIG.SHEETS.SYSTEM);
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return { doctors: [], nurses: [], therapists: [], allStaff: [], treatmentItems: [], trackingTypes: [], maintItems: [] };

    // 一次讀取所有需要的欄位 (A2:G)
    // 假設結構: A:醫, B:護, C:治, D:追蹤項, E:保養項, F:全員, G:治療項
    const data = sheet.getRange(2, 1, lastRow - 1, 7).getValues();

    const result = {
      doctors: [], nurses: [], therapists: [], trackingTypes: [], maintItems: [], allStaff: [], treatmentItems: []
    };

    // 使用 JS 迴圈處理，比多次 getRange 快
    for (let i = 0; i < data.length; i++) {
      if (data[i][0]) result.doctors.push(data[i][0]);
      if (data[i][1]) result.nurses.push(data[i][1]);
      if (data[i][2]) result.therapists.push(data[i][2]);
      if (data[i][3]) result.trackingTypes.push(data[i][3]);
      if (data[i][4]) result.maintItems.push(data[i][4]);
      if (data[i][5]) result.allStaff.push(data[i][5]);
      if (data[i][6]) result.treatmentItems.push(data[i][6]);
    }

    return result;
  } catch (e) {
    Logger.log(e);
    return {};
  }
}

// 1. 搜尋功能
function searchClient(keyword) {
  try {
    const sheet = getSheetHelper(CONFIG.SHEETS.CLIENT);
    const data = sheet.getDataRange().getDisplayValues(); // 保持 DisplayValues 以利比對
    const results = [];
    
    const query = String(keyword).replace(/\s+/g, '').toLowerCase();
    if (!query) return [];
    
    // 優化：預先計算索引
    const headers = data[0]; // 假設第一列是標題，雖然後面用固定索引，但保留擴充性
    
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      // 快速檢查：合併字串比對 (效能微調)
      const id = String(row[0]).replace(/^'/, '').replace(/\s+/g, '').toLowerCase();
      const name = String(row[1]).replace(/\s+/g, '').toLowerCase();
      const phone = String(row[4]).replace(/^'/, '').replace(/\s+/g, '').toLowerCase();
      
      if (id.includes(query) || name.includes(query) || phone.includes(query)) {
        results.push({
          '個案編號': row[0], '姓名': row[1], '生日': row[2], '身分證字號': row[3],
          '電話': row[4], '性別': row[5], '緊急聯絡人': row[6], '緊急聯絡人電話': row[7],
          '負責治療師': row[8], '慢性病或特殊疾病': row[9], 'GoogleDrive資料夾連結': row[10], '建立日期': row[11]
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
    lock.waitLock(5000); // 縮短等待時間
    
    const targetSheetName = sheetName || CONFIG.SHEETS.CLIENT;
    const sheet = getSheetHelper(targetSheetName);
    const ss = sheet.getParent();

    const rawHeaders = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    
    // 檢查必填 ID
    if (targetSheetName !== CONFIG.SHEETS.CLIENT) {
        let hasClientId = false;
        const cleanIdKey = normalizeHeader('個案編號');
        for (let key in dataObj) {
            if (normalizeHeader(key) === cleanIdKey && dataObj[key]) hasClientId = true;
        }
        if (!hasClientId) throw new Error("系統錯誤：未偵測到個案編號。");
    }

    const rowData = rawHeaders.map(rawH => {
        const cleanH = normalizeHeader(rawH);
        let val = '';
        // 尋找對應值 (忽略大小寫與空白)
        for (let key in dataObj) {
            if (normalizeHeader(key) === cleanH) { val = dataObj[key]; break; }
        }
        // 自動生成欄位
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

// 3. 特殊儲存功能 (保持邏輯，僅透過 getSheetHelper 優化連線)
function saveTrackingRecord(formObj) {
  try {
    const sheet = getSheetHelper(CONFIG.SHEETS.TRACKING);
    if (sheet.getLastRow() === 0) sheet.appendRow(["追蹤ID", "個案編號", "追蹤日期", "追蹤人員", "追蹤項目", "追蹤內容", "建立時間"]);
    
    const ss = sheet.getParent();
    if (!formObj.clientId) throw new Error("無個案編號");

    const newRow = [
      "TR" + new Date().getTime(), "'" + formObj.clientId, formObj.trackDate,
      formObj.trackStaff, formObj.trackType, formObj.content,
      Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), "yyyy-MM-dd HH:mm:ss")
    ];
    sheet.appendRow(newRow);
    return { success: true, message: "追蹤紀錄已新增" };
  } catch (e) { return { success: false, message: "儲存失敗: " + e.toString() }; }
}

function saveDoctorConsultation(formData) {
  try {
    const sheet = getSheetHelper(CONFIG.SHEETS.DOCTOR);
    const ss = sheet.getParent();
    const rowData = [
      "DOC" + new Date().getTime(), "'" + formData.clientId, formData.date,
      formData.doctor, formData.nurse, formData.complaint, formData.objective,
      formData.diagnosis, formData.plan, formData.nursingRecord, formData.remark, "",
      Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), "yyyy/MM/dd HH:mm:ss")
    ];
    sheet.appendRow(rowData);
    return { success: true, message: "醫師看診紀錄儲存成功" };
  } catch (e) { return { success: false, message: "儲存失敗: " + e.toString() }; }
}

function saveMaintenanceRecord(data) {
  try {
    const sheet = getSheetHelper(CONFIG.SHEETS.MAINTENANCE);
    const newRow = [
      Utilities.getUuid(), "'" + data.clientId, data.date, data.staff, data.item,
      data.bp, data.spo2, data.hr, data.temp, data.rr, data.remark, new Date()
    ];
    sheet.appendRow(newRow);
    return { success: true, message: "保養紀錄儲存成功！" };
  } catch (e) { return { success: false, message: "儲存失敗：" + e.toString() }; }
}

// 4. 建立新個案
function createNewClient(data) {
  try {
    const sheet = getSheetHelper(CONFIG.SHEETS.CLIENT);
    const now = new Date();
    const datePart = Utilities.formatDate(now, "GMT+8", "yyyyMMdd");
    const lastRow = sheet.getLastRow();
    const clientId = "CF" + datePart + (lastRow + 1).toString().padStart(3, '0');

    let folderUrl = "資料夾建立失敗";
    try { 
        const parentFolder = DriveApp.getFolderById(CONFIG.PARENT_FOLDER_ID);
        const folder = parentFolder.createFolder(clientId + "_" + data.name); 
        folderUrl = folder.getUrl(); 
    } catch (e) {}

    const newRow = [
      clientId, data.name, data.dob, data.idNo, "'" + data.phone, data.gender,
      data.emerName, "'" + data.emerPhone, data.therapist || "", data.chronic, folderUrl, now
    ];
    sheet.appendRow(newRow);

    return { 
      success: true, clientId: clientId, 
      fullData: {
        '個案編號': clientId, '姓名': data.name, '生日': data.dob, '身分證字號': data.idNo,
        '電話': data.phone, '性別': data.gender, '緊急聯絡人': data.emerName, 
        '緊急聯絡人電話': data.emerPhone, '負責治療師': data.therapist, '慢性病或特殊疾病': data.chronic
      } 
    };
  } catch (e) { return { success: false, message: e.toString() }; }
}

// 5. 歷史紀錄查詢 (共用 Helper)
function getHistoryCommon(sheetName, clientId, headersMap) {
  try {
    if (!clientId) return [];
    const sheet = getSheetHelper(sheetName);
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return [];

    const headers = data[0];
    const idxClientId = headers.map(normalizeHeader).indexOf("個案編號");
    const targetColIdx = idxClientId > -1 ? idxClientId : 1;
    const targetId = String(clientId).trim();
    
    const results = data.slice(1)
      .filter(row => String(row[targetColIdx]).replace(/^'/, '').trim() === targetId)
      .map(row => headersMap(row, headers));
      
    // 簡單判斷日期欄位進行排序 (假設 key 包含 date)
    return results.reverse(); 
  } catch (e) { return []; }
}

function getTrackingHistory(clientId) {
  return getHistoryCommon(CONFIG.SHEETS.TRACKING, clientId, (row, h) => {
     const ss = SpreadsheetApp.getActive(); // 僅用於 Format Date
     let d = row[h.indexOf("追蹤日期")];
     if (d instanceof Date) d = Utilities.formatDate(d, Session.getScriptTimeZone(), "yyyy-MM-dd");
     return {
       id: row[h.indexOf("追蹤ID")], date: d, staff: row[h.indexOf("追蹤人員")],
       type: row[h.indexOf("追蹤項目")], content: row[h.indexOf("追蹤內容")]
     };
  });
}

function getMaintenanceHistory(clientId) {
    // 保養項目強制指定 Index
    try {
        if (!clientId) return [];
        const sheet = getSheetHelper(CONFIG.SHEETS.MAINTENANCE);
        const data = sheet.getDataRange().getValues();
        const targetId = String(clientId).trim();
        // Index: 0=ID, 1=ClientID, 2=Date, 3=Staff, 4=Item, 5=BP, 6=SpO2, 7=HR, 8=Temp, 9=RR, 10=Remark
        const results = data.slice(1)
          .filter(row => String(row[1]).replace(/^'/, '').trim() === targetId)
          .map(row => ({
            id: row[0],
            date: (row[2] instanceof Date) ? Utilities.formatDate(row[2], "GMT+8", "yyyy-MM-dd") : row[2],
            staff: row[3], item: row[4], bp: row[5], spo2: row[6], hr: row[7], temp: row[8],
            rr: row[9], remark: row[10]
          }));
        return results.reverse();
    } catch(e) { return []; }
}

function getClientHistory(clientId, sheetName) {
  try {
    if (!clientId) return [];
    const sheet = getSheetHelper(sheetName);
    const data = sheet.getDataRange().getDisplayValues();
    const headers = data[0];
    const normHeaders = headers.map(normalizeHeader);
    let idxCaseId = normHeaders.indexOf(normalizeHeader('個案編號'));
    if (idxCaseId === -1) idxCaseId = 1;

    const targetId = String(clientId).replace(/^'/, '').trim().toLowerCase();
    
    const result = data.slice(1)
       .filter(row => String(row[idxCaseId]).replace(/^'/, '').trim().toLowerCase() === targetId)
       .map(row => {
          let obj = {};
          headers.forEach((h, i) => obj[h] = row[i]);
          return obj;
       });

    result.sort((a, b) => {
      const dA = a['看診日期'] || a['治療日期'] || a['日期'] || '1900-01-01';
      const dB = b['看診日期'] || b['治療日期'] || b['日期'] || '1900-01-01';
      return new Date(dB) - new Date(dA);
    });
    return result;
  } catch (e) { return []; }
}

// 6. 個案總覽 (由於 getSheetHelper 已有快取，這裡的連續呼叫效能已改善)
function getCaseOverviewData(clientId) {
  try {
    if (!clientId) return [];
    const result = [];
    const targetId = String(clientId).trim();

    // 簡單的 Helper 來處理重複的讀表邏輯
    const fetchSheetData = (sheetName, type, mapFunc) => {
      try {
        const sheet = getSheetHelper(sheetName);
        const data = sheet.getDataRange().getValues();
        const headers = data[0];
        const idx = headers.map(normalizeHeader).indexOf(normalizeHeader("個案編號"));
        const targetCol = idx > -1 ? idx : 1;
        
        data.slice(1).forEach(row => {
            if (String(row[targetCol]).replace(/^'/, '').trim() === targetId) {
                const item = mapFunc(row, headers);
                if (item) {
                   item.category = type;
                   item.date = formatDateForJSON(item.date); // 統一格式
                   result.push(item);
                }
            }
        });
      } catch(e) { console.log(sheetName + " Error: " + e); }
    };

    fetchSheetData(CONFIG.SHEETS.DOCTOR, 'doctor', (r, h) => ({
        id: r[0], date: r[2], categoryName: '醫師看診', doctor: r[3], nurse: r[4],
        s: r[5], o: r[6], a: r[7], p: r[8], nursingRecord: r[9], remark: r[10]
    }));

    fetchSheetData(CONFIG.SHEETS.MAINTENANCE, 'maintenance', (r, h) => ({
        id: r[0], date: r[2], categoryName: '保養項目', staff: r[3], item: r[4],
        bp: r[5], spo2: r[6], hr: r[7], temp: r[8], rr: r[9], remark: r[10]
    }));

    fetchSheetData(CONFIG.SHEETS.TRACKING, 'tracking', (r, h) => ({
        id: r[0], date: r[2], categoryName: '個管追蹤', staff: r[3], type: r[4], content: r[5]
    }));

    fetchSheetData(CONFIG.SHEETS.TREATMENT, 'treatment', (r, h) => {
        const getVal = (name) => {
             const i = h.map(normalizeHeader).indexOf(normalizeHeader(name));
             return i > -1 ? r[i] : "";
        };
        return {
           id: 'T-' + formatDateForJSON(getVal("治療日期")), date: getVal("治療日期"), categoryName: '物理治療',
           staff: getVal("執行治療師"), item: getVal("治療項目"), complaint: getVal("當日主訴"),
           content: getVal("治療內容"), nextPlan: getVal("備註/下次治療")
        };
    });

    return result.sort((a, b) => new Date(b.date) - new Date(a.date));
  } catch (e) { throw new Error("總覽失敗: " + e.message); }
}

// 7. 影像功能
function getClientImages(clientId) {
  try {
    const clientSheet = getSheetHelper(CONFIG.SHEETS.CLIENT);
    const clientData = clientSheet.getDataRange().getDisplayValues();
    const targetId = String(clientId).replace(/^'/, '').trim();
    let folderUrl = "";
    
    // 使用迴圈找連結 (假設在第 11 欄/Index 10)
    for (let i = 1; i < clientData.length; i++) {
      if (String(clientData[i][0]).replace(/^'/, '').trim() === targetId) {
        folderUrl = clientData[i][10];
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
        imageList.push({
          id: file.getId(), name: file.getName(), url: file.getUrl(),
          thumbnail: "https://lh3.googleusercontent.com/d/" + file.getId() + "=s400",
          date: Utilities.formatDate(file.getDateCreated(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm")
        });
      }
    }
    imageList.sort((a, b) => new Date(b.date) - new Date(a.date));
    return { success: true, images: imageList };
  } catch (e) { return { success: false, message: e.toString() }; }
}

function uploadClientImage(clientId, fileData, fileName, mimeType, remark) {
  try {
    // 找資料夾邏輯同 getClientImages
    const clientSheet = getSheetHelper(CONFIG.SHEETS.CLIENT);
    const clientData = clientSheet.getDataRange().getValues();
    let folderUrl = "";
    const targetId = String(clientId).trim();
    
    for (let i = 1; i < clientData.length; i++) {
      if (String(clientData[i][0]).trim() === targetId) { 
          folderUrl = clientData[i][10]; break; 
      }
    }
    
    if (!folderUrl) throw new Error("找不到個案資料夾");
    const folder = DriveApp.getFolderById(folderUrl.match(/[-\w]{25,}/)[0]);
    const file = folder.createFile(Utilities.newBlob(Utilities.base64Decode(fileData), mimeType, fileName));
    
    // 寫入紀錄
    const imgSheet = getSheetHelper(CONFIG.SHEETS.IMAGE);
    if (imgSheet.getLastRow() === 0) imgSheet.appendRow(["影像ID", "個案編號", "上傳日期", "檔案名稱", "連結", "備註"]);
    
    imgSheet.appendRow([
      "IMG" + new Date().getTime(), "'" + clientId,
      Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm"),
      fileName, file.getUrl(), remark || ""
    ]);

    return { success: true, message: "上傳成功" };
  } catch (e) { return { success: false, message: e.toString() }; }
}

/**
 * ★ 優化核心：批次更新 Record (取代多次 setValue)
 */
function updateRecord(type, formData) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) return { success: false, message: "系統忙碌" };

  try {
    let sheetName = "";
    let expectedIdHeader = "紀錄id";
    
    // 定義欄位映射
    if (type === 'treatment') { sheetName = CONFIG.SHEETS.TREATMENT; }
    else if (type === 'doctor') { sheetName = CONFIG.SHEETS.DOCTOR; }
    else if (type === 'maintenance') { sheetName = CONFIG.SHEETS.MAINTENANCE; }
    else if (type === 'tracking') { sheetName = CONFIG.SHEETS.TRACKING; expectedIdHeader = "追蹤ID"; }

    const sheet = getSheetHelper(sheetName);
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    
    // 尋找 ID 欄位
    let idColIndex = headers.map(normalizeHeader).indexOf(normalizeHeader(expectedIdHeader));
    if (idColIndex === -1 && type !== 'treatment') idColIndex = 0; // 預設第一欄
    
    const targetId = String(formData.record_id).trim();
    let rowIndex = -1;
    
    // 尋找目標列
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][idColIndex]).trim() === targetId) {
        rowIndex = i + 1;
        break;
      }
    }

    if (rowIndex === -1) return { success: false, message: "找不到資料 ID" };

    // ★ 優化：一次寫入整列 (避免多次 setValue)
    // 1. 取得目前該列資料
    const currentRowValues = data[rowIndex - 1]; 
    
    // 2. 根據 header 更新陣列中的值 (In-Memory Operation)
    const updateVal = (headerName, val) => {
        const idx = headers.map(normalizeHeader).indexOf(normalizeHeader(headerName));
        if (idx > -1) currentRowValues[idx] = val;
    };

    if (type === 'treatment') {
       updateVal("治療日期", formData.date);
       updateVal("執行治療師", formData.therapist);
       updateVal("治療項目", formData.item);
       updateVal("當日主訴", formData.complaint);
       updateVal("治療內容", formData.content);
       updateVal("備註/下次治療", formData.nextPlan);
    } 
    else if (type === 'doctor') {
       // 固定索引亦可，但用 header 比較安全
       currentRowValues[2] = formData.date;
       currentRowValues[3] = formData.doctor;
       currentRowValues[4] = formData.nurse;
       currentRowValues[5] = formData.complaint;
       currentRowValues[6] = formData.objective;
       currentRowValues[7] = formData.diagnosis;
       currentRowValues[8] = formData.plan;
       currentRowValues[9] = formData.nursingRecord;
       currentRowValues[10] = formData.remark;
    } 
    else if (type === 'maintenance') {
       currentRowValues[2] = formData.date;
       currentRowValues[3] = formData.staff;
       currentRowValues[4] = formData.item;
       currentRowValues[5] = formData.bp;
       currentRowValues[6] = formData.spo2;
       currentRowValues[7] = formData.hr;
       currentRowValues[8] = formData.temp;
       currentRowValues[9] = formData.rr;     // Index 9
       currentRowValues[10] = formData.remark; // Index 10
    } 
    else if (type === 'tracking') {
       currentRowValues[2] = formData.trackDate;
       currentRowValues[3] = formData.trackStaff;
       currentRowValues[4] = formData.trackType;
       currentRowValues[5] = formData.content;
    }

    // 3. 一次寫回 (Batch Write)
    sheet.getRange(rowIndex, 1, 1, currentRowValues.length).setValues([currentRowValues]);

    return { success: true, message: "更新成功" };

  } catch (e) {
    return { success: false, message: "更新失敗: " + e.toString() };
  } finally {
    lock.releaseLock();
  }
}

function updateClientBasicInfo(data) {
  // 保持原有邏輯，因欄位較固定，使用 setValues 優化
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(5000);
    const sheet = getSheetHelper(CONFIG.SHEETS.CLIENT);
    const dataRange = sheet.getDataRange();
    const rows = dataRange.getDisplayValues(); // DisplayValues 用於 ID 比對
    const targetId = String(data.clientId).trim();
    
    let rowIndex = -1;
    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][0]).replace(/^'/, '').trim() === targetId) {
        rowIndex = i + 1; break;
      }
    }
    
    if (rowIndex === -1) throw new Error("找不到個案");

    // 更新陣列 (注意：原資料庫可能用 Value，這裡用 setValues 寫入原始值)
    // 取得該列原始資料 (getValues)
    const rawRow = sheet.getRange(rowIndex, 1, 1, sheet.getLastColumn()).getValues()[0];
    
    // Mapping: A=0, B=1(Name), C=2(DOB), D=3(ID), E=4(Phone), F=5(Gen), G=6(Emer), H=7(EPhone), I=8(Therapist), J=9(Chronic)
    rawRow[1] = data.name;
    rawRow[2] = data.dob;
    rawRow[3] = data.idNo;
    rawRow[4] = "'" + data.phone;
    rawRow[5] = data.gender;
    rawRow[6] = data.emerName;
    rawRow[7] = "'" + data.emerPhone;
    rawRow[8] = data.therapist;
    rawRow[9] = data.chronic;

    sheet.getRange(rowIndex, 1, 1, rawRow.length).setValues([rawRow]);
    
    return {
      success: true, message: "更新成功",
      updatedData: {
        '個案編號': data.clientId, '姓名': data.name, '生日': data.dob,
        '身分證字號': data.idNo, '電話': data.phone, '性別': data.gender,
        '緊急聯絡人': data.emerName, '緊急聯絡人電話': data.emerPhone,
        '負責治療師': data.therapist, '慢性病或特殊疾病': data.chronic, '狀態': 'Active'
      }
    };
  } catch (e) { return { success: false, message: e.toString() }; } finally { lock.releaseLock(); }
}

function formatDateForJSON(d) {
  if (d instanceof Date) return Utilities.formatDate(d, Session.getScriptTimeZone(), "yyyy-MM-dd");
  return String(d || "");
}