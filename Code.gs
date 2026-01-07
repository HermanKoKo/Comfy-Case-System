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

/**
 * ==========================================
 * 網頁應用程式入口
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
 * 核心邏輯層
 * ==========================================
 */

function normalizeHeader(header) {
  return String(header).replace(/\s+/g, '').trim().toLowerCase();
}

function getSheetHelper(sheetName) {
  let ss;
  if (sheetName === CONFIG.SHEETS.CLIENT) {
    try {
      ss = SpreadsheetApp.openById(CONFIG.CLIENT_DB_ID);
    } catch (e) {
      throw new Error("無法連接個案核心資料庫 (Client DB)，請檢查 ID 是否正確或是否有權限。");
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

function getTherapistList() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("system");
    if (!sheet) throw new Error("找不到名為 'system' 的分頁");
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return []; 
    const data = sheet.getRange(2, 3, lastRow - 1, 1).getValues();
    return data.flat().filter(function(name) { return name && name.toString().trim() !== ""; });
  } catch (e) {
    Logger.log("Error in getTherapistList: " + e.toString());
    return []; 
  }
}

function searchClient(keyword) {
  try {
    const sheet = getSheetHelper(CONFIG.SHEETS.CLIENT);
    const data = sheet.getDataRange().getDisplayValues(); 
    const results = [];
    const query = String(keyword).replace(/\s+/g, '').toLowerCase();
    if (!query) return [];
    
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
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

function getClientById(clientId) {
  try {
    const sheet = getSheetHelper(CONFIG.SHEETS.CLIENT);
    const data = sheet.getDataRange().getDisplayValues(); 
    const targetId = String(clientId).replace(/^'/, '').trim();
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (String(row[0]).replace(/^'/, '').trim() === targetId) {
        return {
          '個案編號': row[0], '姓名': row[1], '生日': row[2], '身分證字號': row[3],
          '電話': row[4], '性別': row[5], '緊急聯絡人': row[6], '緊急聯絡人電話': row[7],
          '負責治療師': row[8], '慢性病或特殊疾病': row[9], 'GoogleDrive資料夾連結': row[10], '建立日期': row[11]
        };
      }
    }
    throw new Error("找不到個案 ID: " + clientId);
  } catch (e) { throw new Error(e.message); }
}

function saveData(sheetName, dataObj) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000); 
    const targetSheetName = sheetName || CONFIG.SHEETS.CLIENT;
    const sheet = getSheetHelper(targetSheetName);
    const ss = sheet.getParent();
    const rawHeaders = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const cleanIdKey = normalizeHeader('個案編號');
    let hasClientId = false;
    for (let key in dataObj) {
        if (normalizeHeader(key) === cleanIdKey && dataObj[key]) hasClientId = true;
    }
    if (targetSheetName !== CONFIG.SHEETS.CLIENT && !hasClientId) throw new Error("系統錯誤：未偵測到個案編號，無法儲存。");

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

function getSystemStaff() {
  const sheet = getSheetHelper(CONFIG.SHEETS.SYSTEM);
  const data = sheet.getDataRange().getValues();
  const rows = data.slice(1);
  return {
    doctors: rows.map(r => r[0]).filter(String),
    nurses: rows.map(r => r[1]).filter(String),
    therapists: rows.map(r => r[2]).filter(String), 
    trackingTypes: rows.map(r => r[3]).filter(String),
    maintItems: rows.map(r => r[4]).filter(String),
    allStaff: rows.map(r => r[5]).filter(String),
    treatmentItems: rows.map(r => r[6]).filter(String)
  };
}

function saveTrackingRecord(formObj) {
  try {
    const sheet = getSheetHelper(CONFIG.SHEETS.TRACKING);
    if (sheet.getLastRow() === 0) sheet.appendRow(["追蹤ID", "個案編號", "追蹤日期", "追蹤人員", "追蹤項目", "追蹤內容", "建立時間"]);
    const ss = sheet.getParent();
    if (!formObj.clientId) throw new Error("無個案編號");
    const now = new Date();
    sheet.appendRow([
      "TR" + new Date().getTime(), "'" + formObj.clientId, formObj.trackDate, formObj.trackStaff, formObj.trackType, formObj.content,
      Utilities.formatDate(now, ss.getSpreadsheetTimeZone(), "yyyy-MM-dd HH:mm:ss")
    ]);
    return { success: true, message: "追蹤紀錄已新增" };
  } catch (e) { return { success: false, message: "儲存失敗: " + e.toString() }; }
}

function getTrackingHistory(clientId) {
  try {
    if (!clientId) return [];
    const sheet = getSheetHelper(CONFIG.SHEETS.TRACKING);
    const ss = sheet.getParent();
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return [];
    const idxClientId = data[0].map(normalizeHeader).indexOf("個案編號");
    const targetColIdx = idxClientId > -1 ? idxClientId : 1;
    const targetId = String(clientId).trim();
    return data.slice(1)
      .filter(row => String(row[targetColIdx]).replace(/^'/, '').trim() === targetId)
      .map(row => ({
          id: row[0], // 假設追蹤ID在第0欄
          date: row[2] instanceof Date ? Utilities.formatDate(row[2], ss.getSpreadsheetTimeZone(), "yyyy-MM-dd") : row[2],
          staff: row[3], type: row[4], content: row[5]
      })).sort((a, b) => new Date(b.date) - new Date(a.date));
  } catch (e) { return []; }
}

function saveDoctorConsultation(formData) {
  try {
    const sheet = getSheetHelper(CONFIG.SHEETS.DOCTOR);
    const ss = sheet.getParent();
    if (!formData.clientId) throw new Error("無個案編號");
    sheet.appendRow([
      "DOC" + new Date().getTime(), "'" + formData.clientId, formData.date, formData.doctor, formData.nurse,
      formData.complaint, formData.objective, formData.diagnosis, formData.plan, formData.nursingRecord, formData.remark, "",
      Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), "yyyy/MM/dd HH:mm:ss")
    ]);
    return { success: true, message: "醫師看診紀錄儲存成功" };
  } catch (e) { return { success: false, message: "儲存失敗: " + e.toString() }; }
}

function createNewClient(data) {
  try {
    const sheet = getSheetHelper(CONFIG.SHEETS.CLIENT);
    const now = new Date();
    const datePart = Utilities.formatDate(now, "GMT+8", "yyyyMMdd");
    const suffix = (sheet.getLastRow() + 1).toString().padStart(3, '0');
    const clientId = "CF" + datePart + suffix;
    let folderUrl = "";
    try { 
        folderUrl = DriveApp.getFolderById(CONFIG.PARENT_FOLDER_ID).createFolder(clientId + "_" + data.name).getUrl(); 
    } catch (e) { folderUrl = "資料夾建立失敗"; }

    sheet.appendRow([
      clientId, data.name, data.dob, data.idNo, "'" + data.phone, data.gender, data.emerName, "'" + data.emerPhone,
      data.therapist || "", data.chronic, folderUrl, now
    ]);
    return { success: true, clientId: clientId, fullData: {
        '個案編號': clientId, '姓名': data.name, '生日': data.dob, '身分證字號': data.idNo, '電話': data.phone,
        '性別': data.gender, '緊急聯絡人': data.emerName, '緊急聯絡人電話': data.emerPhone,
        '負責治療師': data.therapist, '慢性病或特殊疾病': data.chronic
    }};
  } catch (e) { return { success: false, message: e.toString() }; }
}

function getMaintenanceHistory(clientId) {
  try {
    if (!clientId) return [];
    const sheet = getSheetHelper(CONFIG.SHEETS.MAINTENANCE);
    const data = sheet.getDataRange().getValues();
    const idx = data[0].map(normalizeHeader).indexOf(normalizeHeader("個案編號"));
    const clientIdx = idx > -1 ? idx : 1; 
    const targetId = String(clientId).trim();
    return data.slice(1)
      .filter(row => String(row[clientIdx]).replace(/^'/, '').trim() === targetId)
      .map(row => {
        let obj = {};
        data[0].forEach((h, i) => {
          let val = row[i];
          if (val instanceof Date) val = Utilities.formatDate(val, "GMT+8", "yyyy-MM-dd");
          obj[h] = val;
        });
        return obj;
      }).reverse();
  } catch (e) { return []; }
}

function getClientHistory(clientId, sheetName) {
  try {
    if (!clientId) return [];
    const sheet = getSheetHelper(sheetName);
    const data = sheet.getDataRange().getDisplayValues();
    if (data.length < 2) return [];
    const idxCaseId = data[0].map(normalizeHeader).indexOf(normalizeHeader('個案編號')) > -1 ? data[0].map(normalizeHeader).indexOf(normalizeHeader('個案編號')) : 1;
    const targetId = String(clientId).replace(/^'/, '').trim().toLowerCase();
    
    const result = [];
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][idxCaseId]).replace(/^'/, '').trim().toLowerCase() === targetId) {
        let obj = {};
        data[0].forEach((header, index) => { obj[header] = data[i][index]; });
        result.push(obj);
      }
    }
    return result.sort((a, b) => {
      const dateStrA = a['看診日期'] || a['治療日期'] || a['日期'] || '1900-01-01';
      const dateStrB = b['看診日期'] || b['治療日期'] || b['日期'] || '1900-01-01';
      return new Date(dateStrB) - new Date(dateStrA);
    });
  } catch (e) { return []; }
}

function saveMaintenanceRecord(data) {
  try {
    const sheet = getSheetHelper(CONFIG.SHEETS.MAINTENANCE);
    if (!data.clientId) throw new Error("無個案編號");
    sheet.appendRow([
      Utilities.getUuid(), "'" + data.clientId, data.date, data.staff, data.item,
      data.bp, data.spo2, data.hr, data.temp, data.rr, data.remark, new Date()
    ]);
    return { success: true, message: "保養紀錄儲存成功！" };
  } catch (e) { return { success: false, message: "儲存失敗：" + e.toString() }; }
}

// 1. 新增：標記任務為已完成
function markTaskAsCompleted(taskKey) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(5000);
    const sheet = getSheetHelper("Task_Completion_Log");
    if (sheet.getLastRow() === 0) sheet.appendRow(["TaskKey", "完成時間", "操作人員"]);
    sheet.appendRow([taskKey, Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss"), Session.getActiveUser().getEmail() || "Admin"]);
    return { success: true };
  } catch (e) { return { success: false, message: e.toString() }; } finally { lock.releaseLock(); }
}

// 2. 修改：取得每日追蹤清單
function getDailyTasks() {
  const tasks = [];
  const clientMap = {};
  const completedSet = new Set();

  try {
    const logSheet = getSheetHelper("Task_Completion_Log");
    const logData = logSheet.getDataRange().getValues();
    for (let i = 1; i < logData.length; i++) completedSet.add(String(logData[i][0]).trim());
  } catch (e) { console.warn("Task_Completion_Log 讀取失敗", e); }

  try {
    const clientSheet = getSheetHelper(CONFIG.SHEETS.CLIENT);
    const clientData = clientSheet.getDataRange().getDisplayValues(); 
    for (let i = 1; i < clientData.length; i++) clientMap[String(clientData[i][0]).replace(/^'/, '').trim()] = clientData[i][1];
  } catch (e) { console.warn("個案索引失敗", e); }

  try {
    const docSheet = getSheetHelper(CONFIG.SHEETS.DOCTOR);
    const docData = docSheet.getDataRange().getValues();
    for (let i = 1; i < docData.length; i++) {
      const clientId = String(docData[i][1]).replace(/^'/, '').trim();
      const dateVal = docData[i][2];
      if (clientId && dateVal instanceof Date) {
        const sourceDateStr = Utilities.formatDate(dateVal, Session.getScriptTimeZone(), "yyyy-MM-dd");
        const taskKey = `${clientId}_${sourceDateStr}_doctor`;
        if (completedSet.has(taskKey)) continue;
        const dueDate = new Date(dateVal);
        dueDate.setDate(dueDate.getDate() + 3);
        tasks.push({
          taskKey: taskKey, clientId: clientId, clientName: clientMap[clientId] || "未知個案",
          type: '醫師看診追蹤', sourceDate: sourceDateStr,
          dueDate: Utilities.formatDate(dueDate, Session.getScriptTimeZone(), "yyyy-MM-dd"),
          timestamp: dueDate.getTime(), tagColor: 'bg-emerald-100 text-emerald-800'
        });
      }
    }
  } catch (e) { console.warn("醫師任務失敗", e); }

  try {
    const treatSheet = getSheetHelper(CONFIG.SHEETS.TREATMENT);
    const treatData = treatSheet.getDataRange().getValues();
    if (treatData.length > 1) {
      const headers = treatData[0].map(normalizeHeader);
      let idxId = headers.indexOf(normalizeHeader('個案編號')); if (idxId === -1) idxId = 1; 
      const idxDate = headers.indexOf(normalizeHeader('治療日期'));
      const idxItem = headers.indexOf(normalizeHeader('治療項目'));
      if (idxDate > -1 && idxItem > -1) {
        for (let i = 1; i < treatData.length; i++) {
          const item = String(treatData[i][idxItem]);
          if (item && item.includes("初診")) {
            const clientId = String(treatData[i][idxId]).replace(/^'/, '').trim();
            const dateVal = treatData[i][idxDate];
            if (dateVal instanceof Date) {
              const sourceDateStr = Utilities.formatDate(dateVal, Session.getScriptTimeZone(), "yyyy-MM-dd");
              const taskKey = `${clientId}_${sourceDateStr}_treatment`;
              if (completedSet.has(taskKey)) continue;
              const dueDate = new Date(dateVal);
              dueDate.setDate(dueDate.getDate() + 3);
              tasks.push({
                taskKey: taskKey, clientId: clientId, clientName: clientMap[clientId] || "未知個案",
                type: '物理治療初診追蹤', sourceDate: sourceDateStr,
                dueDate: Utilities.formatDate(dueDate, Session.getScriptTimeZone(), "yyyy-MM-dd"),
                timestamp: dueDate.getTime(), tagColor: 'bg-orange-100 text-orange-800'
              });
            }
          }
        }
      }
    }
  } catch (e) { console.warn("治療任務失敗", e); }

  tasks.sort((a, b) => b.timestamp - a.timestamp);
  return tasks;
}

/**
 * 取得個案總覽資料 (已修改：防呆處理)
 */
function getCaseOverviewData(clientId) {
  try {
    if (!clientId) return [];
    
    const result = [];
    const targetId = String(clientId).trim();
    
    // 1. 醫師看診
    try {
      const docSheet = getSheetHelper(CONFIG.SHEETS.DOCTOR);
      const data = docSheet.getDataRange().getValues();
      if (data.length > 1) {
        const idx = data[0].map(normalizeHeader).indexOf(normalizeHeader("個案編號"));
        const targetCol = idx > -1 ? idx : 1; 
        data.slice(1).forEach(row => {
          if (String(row[targetCol]).replace(/^'/, '').trim() === targetId) {
            result.push({
              id: row[0], date: formatDateForJSON(row[2]), category: 'doctor', categoryName: '醫師看診',
              doctor: row[3] || "", nurse: row[4] || "",
              s: row[5] || "", o: row[6] || "", a: row[7] || "", p: row[8] || "",
              nursingRecord: row[9] || "", remark: row[10] || ""
            });
          }
        });
      }
    } catch (e) { console.log("讀取醫師看診失敗: " + e.toString()); }

    // 2. 保養項目
    try {
      const maintSheet = getSheetHelper(CONFIG.SHEETS.MAINTENANCE);
      const data = maintSheet.getDataRange().getValues();
      if (data.length > 1) {
        const idx = data[0].map(normalizeHeader).indexOf(normalizeHeader("個案編號"));
        const targetCol = idx > -1 ? idx : 1;
        data.slice(1).forEach(row => {
          if (String(row[targetCol]).replace(/^'/, '').trim() === targetId) {
            result.push({
              id: row[0], date: formatDateForJSON(row[2]), category: 'maintenance', categoryName: '保養項目',
              staff: row[3] || "", item: row[4] || "",
              bp: row[5] || "", spo2: row[6] || "", hr: row[7] || "", temp: row[8] || "",
              rr: (row.length > 9) ? row[9] : "", remark: (row.length > 10) ? row[10] : ""
            });
          }
        });
      }
    } catch (e) { console.log("讀取保養項目失敗: " + e.toString()); }

    // 3. 個管追蹤
    try {
      const trackSheet = getSheetHelper(CONFIG.SHEETS.TRACKING);
      const data = trackSheet.getDataRange().getValues();
      if (data.length > 1) {
        const idx = data[0].map(normalizeHeader).indexOf(normalizeHeader("個案編號"));
        const targetCol = idx > -1 ? idx : 1;
        data.slice(1).forEach(row => {
          if (String(row[targetCol]).replace(/^'/, '').trim() === targetId) {
            result.push({
              id: row[0], date: formatDateForJSON(row[2]), category: 'tracking', categoryName: '個管追蹤',
              staff: row[3] || "", type: row[4] || "", content: row[5] || ""
            });
          }
        });
      }
    } catch (e) { console.log("讀取個管追蹤失敗: " + e.toString()); }

    // 4. 治療紀錄
    try {
      const treatSheet = getSheetHelper(CONFIG.SHEETS.TREATMENT);
      const data = treatSheet.getDataRange().getValues();
      if (data.length > 1) {
        const headers = data[0].map(normalizeHeader);
        let idxId = headers.indexOf(normalizeHeader("個案編號")); if (idxId === -1) idxId = 1; 
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
              staff: row[idxStaff] || "",
              item: (idxItem > -1) ? row[idxItem] : "",
              complaint: (idxComplaint > -1) ? row[idxComplaint] : "",
              content: (idxContent > -1) ? row[idxContent] : "",
              nextPlan: (idxNext > -1) ? row[idxNext] : "" 
            });
          }
        });
      }
    } catch (e) { console.log("讀取治療紀錄失敗: " + e.toString()); }

    // 排序：安全處理無效日期
    return result.sort((a, b) => {
        const dateA = new Date(a.date);
        const dateB = new Date(b.date);
        if (isNaN(dateA.getTime())) return 1; // 無效日期排後面
        if (isNaN(dateB.getTime())) return -1;
        return dateB - dateA;
    });

  } catch (e) { throw new Error("取得總覽資料失敗: " + e.message); }
}

function getClientImages(clientId) {
  try {
    const clientSheet = getSheetHelper(CONFIG.SHEETS.CLIENT);
    const clientData = clientSheet.getDataRange().getDisplayValues(); 
    let folderUrl = "";
    const targetId = String(clientId).replace(/^'/, '').trim();
    for (let i = 1; i < clientData.length; i++) {
      if (String(clientData[i][0]).replace(/^'/, '').trim() === targetId) {
        folderUrl = clientData[i][10]; break;
      }
    }
    if (!folderUrl || !folderUrl.match(/[-\w]{25,}/)) return { success: true, images: [] };
    const folder = DriveApp.getFolderById(folderUrl.match(/[-\w]{25,}/)[0]);
    const files = folder.getFiles();
    const imageList = [];
    while (files.hasNext()) {
      const file = files.next();
      if (file.getMimeType().indexOf('image/') === 0) {
        imageList.push({
          id: file.getId(), name: file.getName(), url: file.getUrl(),
          thumbnail: "https://lh3.googleusercontent.com/d/" + file.getId() + "=s400",
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
    const clientSheet = getSheetHelper(CONFIG.SHEETS.CLIENT);
    const clientData = clientSheet.getDataRange().getDisplayValues();
    let folderUrl = "";
    for (let i = 1; i < clientData.length; i++) {
      if (String(clientData[i][0]).replace(/^'/, '') === String(clientId)) { folderUrl = clientData[i][10]; break; }
    }
    if (!folderUrl) throw new Error("找不到個案資料夾 (資料表無連結)");
    const folder = DriveApp.getFolderById(folderUrl.match(/[-\w]{25,}/)[0]);
    const file = folder.createFile(Utilities.newBlob(Utilities.base64Decode(fileData), mimeType, fileName));
    const imgSheet = getSheetHelper(CONFIG.SHEETS.IMAGE);
    if (imgSheet.getLastRow() === 0) imgSheet.appendRow(["影像ID", "個案編號", "上傳日期", "檔案名稱", "GoogleDrive檔案連結", "備註"]);
    imgSheet.appendRow(["IMG" + new Date().getTime(), "'" + clientId, Utilities.formatDate(new Date(), imgSheet.getParent().getSpreadsheetTimeZone(), "yyyy-MM-dd HH:mm"), fileName, file.getUrl(), remark || ""]);
    return { success: true, message: "上傳成功" };
  } catch (e) { return { success: false, message: "上傳失敗: " + e.toString() }; }
}

function formatDateForJSON(dateVal) {
  if (!dateVal) return "";
  if (dateVal instanceof Date) return Utilities.formatDate(dateVal, Session.getScriptTimeZone(), "yyyy-MM-dd");
  return String(dateVal);
}

function getTreatmentItemsFromSystem() {
  const sheet = getSheetHelper(CONFIG.SHEETS.SYSTEM);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  return sheet.getRange(2, 7, lastRow - 1).getValues().flat().filter(item => item !== "" && item != null);
}

function updateRecord(type, formData) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) return { success: false, message: "系統忙碌中，請稍後再試" };
  try {
    let sheetName = "";
    let targetId = String(formData.record_id).trim(); 
    let expectedIdHeader = "紀錄id"; 
    
    if (type === 'treatment') { sheetName = CONFIG.SHEETS.TREATMENT; expectedIdHeader = "紀錄id"; }
    else if (type === 'doctor') { sheetName = CONFIG.SHEETS.DOCTOR; expectedIdHeader = "紀錄id"; }
    else if (type === 'maintenance') { sheetName = CONFIG.SHEETS.MAINTENANCE; expectedIdHeader = "紀錄id"; }
    else if (type === 'tracking') { sheetName = CONFIG.SHEETS.TRACKING; expectedIdHeader = "追蹤ID"; }
    else throw new Error("未知的編輯類型");

    const sheet = getSheetHelper(sheetName);
    const data = sheet.getDataRange().getValues();
    if (data.length < 2) return { success: false, message: "資料表為空" };

    const headers = data[0]; 
    let idColIndex = -1;
    for (let c = 0; c < headers.length; c++) { if (normalizeHeader(headers[c]) === normalizeHeader(expectedIdHeader)) { idColIndex = c; break; } }
    if (idColIndex === -1) {
       if (type !== 'treatment') idColIndex = 0; 
       else {
          for (let c = 0; c < headers.length; c++) { if (String(headers[c]).toLowerCase().includes("id")) { idColIndex = c; break; } }
          if (idColIndex === -1) throw new Error(`在工作表 ${sheetName} 中找不到「${expectedIdHeader}」欄位。`);
       }
    }

    let rowIndex = -1;
    for (let i = 1; i < data.length; i++) { if (String(data[i][idColIndex]).trim() === targetId) { rowIndex = i + 1; break; } }
    if (rowIndex === -1) return { success: false, message: `找不到該筆資料 (ID: ${targetId})` };

    const getCol = (name) => { for(let k=0; k<headers.length; k++) { if(normalizeHeader(headers[k]) === normalizeHeader(name)) return k + 1; } return -1; };

    if (type === 'treatment') {
       const colDate = getCol("治療日期"); const colTherapist = getCol("執行治療師"); const colItem = getCol("治療項目");
       const colComplaint = getCol("當日主訴"); const colContent = getCol("治療內容"); const colNext = getCol("備註/下次治療");
       if (colDate > 0) sheet.getRange(rowIndex, colDate).setValue(formData.date);
       if (colTherapist > 0) sheet.getRange(rowIndex, colTherapist).setValue(formData.therapist);
       if (colItem > 0) sheet.getRange(rowIndex, colItem).setValue(formData.item);
       if (colComplaint > 0) sheet.getRange(rowIndex, colComplaint).setValue(formData.complaint);
       if (colContent > 0) sheet.getRange(rowIndex, colContent).setValue(formData.content);
       if (colNext > 0) sheet.getRange(rowIndex, colNext).setValue(formData.nextPlan);
    } 
    else if (type === 'doctor') {
       sheet.getRange(rowIndex, 3).setValue(formData.date); sheet.getRange(rowIndex, 4).setValue(formData.doctor);
       sheet.getRange(rowIndex, 5).setValue(formData.nurse); sheet.getRange(rowIndex, 6).setValue(formData.complaint);
       sheet.getRange(rowIndex, 7).setValue(formData.objective); sheet.getRange(rowIndex, 8).setValue(formData.diagnosis);
       sheet.getRange(rowIndex, 9).setValue(formData.plan); sheet.getRange(rowIndex, 10).setValue(formData.nursingRecord);
       sheet.getRange(rowIndex, 11).setValue(formData.remark);
    } 
    else if (type === 'maintenance') {
       sheet.getRange(rowIndex, 3).setValue(formData.date); sheet.getRange(rowIndex, 4).setValue(formData.staff);
       sheet.getRange(rowIndex, 5).setValue(formData.item); sheet.getRange(rowIndex, 6).setValue(formData.bp);
       sheet.getRange(rowIndex, 7).setValue(formData.spo2); sheet.getRange(rowIndex, 8).setValue(formData.hr);
       sheet.getRange(rowIndex, 9).setValue(formData.temp); sheet.getRange(rowIndex, 10).setValue(formData.rr);
       sheet.getRange(rowIndex, 11).setValue(formData.remark);
    } 
    else if (type === 'tracking') {
       sheet.getRange(rowIndex, 3).setValue(formData.trackDate); sheet.getRange(rowIndex, 4).setValue(formData.trackStaff);
       sheet.getRange(rowIndex, 5).setValue(formData.trackType); sheet.getRange(rowIndex, 6).setValue(formData.content);
    }
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
    for (let i = 1; i < rows.length; i++) { if (String(rows[i][0]).replace(/^'/, '').trim() === targetId) { rowIndex = i + 1; break; } }
    if (rowIndex === -1) throw new Error("找不到此個案編號: " + targetId);
    
    sheet.getRange(rowIndex, 2).setValue(data.name); sheet.getRange(rowIndex, 3).setValue(data.dob);
    sheet.getRange(rowIndex, 4).setValue(data.idNo); sheet.getRange(rowIndex, 5).setValue("'" + data.phone);
    sheet.getRange(rowIndex, 6).setValue(data.gender); sheet.getRange(rowIndex, 7).setValue(data.emerName);
    sheet.getRange(rowIndex, 8).setValue("'" + data.emerPhone); sheet.getRange(rowIndex, 9).setValue(data.therapist);
    sheet.getRange(rowIndex, 10).setValue(data.chronic);  
    
    return {
      success: true, message: "基本資料更新成功",
      updatedData: {
        '個案編號': data.clientId, '姓名': data.name, '生日': data.dob, '身分證字號': data.idNo,
        '電話': data.phone, '性別': data.gender, '緊急聯絡人': data.emerName, '緊急聯絡人電話': data.emerPhone,
        '負責治療師': data.therapist, '慢性病或特殊疾病': data.chronic, '狀態': 'Active' 
      }
    };
  } catch (e) { return { success: false, message: "更新失敗: " + e.toString() }; } finally { lock.releaseLock(); }
}