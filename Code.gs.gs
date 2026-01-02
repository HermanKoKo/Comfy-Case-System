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
      
      // 注意：搜尋結果的物件映射可能需要根據 Index.html 的需求調整
      // 這裡維持您原本的邏輯，但請注意資料夾連結是在 row[9]
      if (id.includes(query) || name.includes(query) || phone.includes(query)) {
        results.push({
          '個案編號': row[0], '姓名': row[1], '生日': row[2], '身分證字號': row[3],
          '電話': row[4], '性別': row[5], '慢性病或特殊疾病': row[8], // 修正索引
          'GoogleDrive資料夾連結': row[9], // ★ 修正：索引從 7 改為 9 (J欄)
          '建立日期': row[10],
          '緊急聯絡人': row[6], 
          '緊急聯絡人電話': row[7] 
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
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const sheet = ss.getSheetByName(CONFIG.SHEETS.SYSTEM);
  if(!sheet) return { doctors:[], nurses:[], therapists:[], trackingTypes:[], maintItems:[], allStaff:[], treatmentItems:[] };
  
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

/**
 * 儲存個管追蹤紀錄
 */
function saveTrackingRecord(formObj) {
  try {
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    let sheet = ss.getSheetByName(CONFIG.SHEETS.TRACKING);
    if (!sheet) {
      sheet = ss.insertSheet(CONFIG.SHEETS.TRACKING);
      sheet.appendRow(["追蹤ID", "個案編號", "追蹤日期", "追蹤人員", "追蹤項目", "追蹤內容", "建立時間"]);
    }
    
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
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const sheet = ss.getSheetByName(CONFIG.SHEETS.TRACKING);
    if (!sheet) return [];
    
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
 * 儲存醫師看診紀錄 (★ 修改：加入護理紀錄欄位)
 */
function saveDoctorConsultation(formData) {
  try {
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const sheet = ss.getSheetByName(CONFIG.SHEETS.DOCTOR); 
    
    if (!formData.clientId) throw new Error("無個案編號");

    const recordId = "DOC" + new Date().getTime();
    const timestamp = Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), "yyyy/MM/dd HH:mm:ss");
    
    // 依據您提供的 Sheet 結構 (J欄為護理紀錄)
    // 順序：ID, ClientID, Date, Doctor, Nurse, S, O, A, P, NursingRecord(J), Remark(K), Link(L), Time(M)
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
      formData.nursingRecord,     // ★ 新增：護理紀錄 (對應 J 欄)
      formData.remark,            // 對應 K 欄
      "",                         // 對應 L 欄
      timestamp                   // 對應 M 欄
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
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const sheet = ss.getSheetByName(CONFIG.SHEETS.CLIENT);
    
    const now = new Date();
    const datePart = Utilities.formatDate(now, "GMT+8", "yyyyMMdd");
    const lastRow = sheet.getLastRow();
    const suffix = (lastRow + 1).toString().padStart(3, '0');
    const clientId = "CF" + datePart + suffix;

    let folderUrl = "";
    // 建立資料夾
    try { 
        const parentFolder = DriveApp.getFolderById(CONFIG.PARENT_FOLDER_ID);
        const folder = parentFolder.createFolder(clientId + "_" + data.name); 
        folderUrl = folder.getUrl(); 
    } catch (e) { 
        folderUrl = "資料夾建立失敗"; 
    }

    // 注意：這裡的欄位順序決定了 folderUrl 是第幾個
    // 0:ID, 1:Name, 2:DoB, 3:IDNo, 4:Phone, 5:Gender, 6:EmerName, 7:EmerPhone, 8:Chronic, 9:FolderUrl, 10:Time
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

    const fullData = {
        '個案編號': clientId, '姓名': data.name, '生日': data.dob, '身分證字號': data.idNo,
        '電話': data.phone, '性別': data.gender, '緊急聯絡人': data.emerName, 
        '緊急聯絡人電話': data.emerPhone, '慢性病或特殊疾病': data.chronic
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
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const sheet = ss.getSheetByName(CONFIG.SHEETS.MAINTENANCE);
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
 * 通用歷史紀錄 (用於治療紀錄)
 */
function getClientHistory(clientId, sheetName) {
  try {
    if (!clientId) return [];
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) return [];
    
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
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const sheet = ss.getSheetByName(CONFIG.SHEETS.MAINTENANCE);
    
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
 * 取得個案總覽資料
 */
function getCaseOverviewData(clientId) {
  try {
    if (!clientId) return [];
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const result = [];
    const targetId = String(clientId).trim();
    
    const docSheet = ss.getSheetByName(CONFIG.SHEETS.DOCTOR);
    if (docSheet) {
      const data = docSheet.getDataRange().getValues();
      const idx = data[0].map(normalizeHeader).indexOf(normalizeHeader("個案編號"));
      const targetCol = idx > -1 ? idx : 1; 

      data.slice(1).forEach(row => {
        if (String(row[targetCol]).replace(/^'/, '').trim() === targetId) {
          result.push({
            id: row[0],
            date: formatDateForJSON(row[2]),
            category: 'doctor', categoryName: '醫師看診',
            title: row[3] + " 醫師", subtitle: "診斷：" + (row[7] || '--'),
            detail: row[8], staff: row[4]
          });
        }
      });
    }

    const maintSheet = ss.getSheetByName(CONFIG.SHEETS.MAINTENANCE);
    if (maintSheet) {
      const data = maintSheet.getDataRange().getValues();
      const idx = data[0].map(normalizeHeader).indexOf(normalizeHeader("個案編號"));
      const targetCol = idx > -1 ? idx : 1;

      data.slice(1).forEach(row => {
        if (String(row[targetCol]).replace(/^'/, '').trim() === targetId) {
          const vitals = [];
          if(row[5]) vitals.push(`BP:${row[5]}`);
          if(row[6]) vitals.push(`SpO2:${row[6]}%`);
          result.push({
            id: row[0],
            date: formatDateForJSON(row[2]),
            category: 'maintenance', categoryName: '保養項目',
            title: row[4], subtitle: vitals.join(' | ') || '無生理數值',
            detail: row[9], staff: row[3]
          });
        }
      });
    }

    const trackSheet = ss.getSheetByName(CONFIG.SHEETS.TRACKING);
    if (trackSheet) {
      const data = trackSheet.getDataRange().getValues();
      const idx = data[0].map(normalizeHeader).indexOf(normalizeHeader("個案編號"));
      const targetCol = idx > -1 ? idx : 1;

      data.slice(1).forEach(row => {
        if (String(row[targetCol]).replace(/^'/, '').trim() === targetId) {
          result.push({
            id: row[0],
            date: formatDateForJSON(row[2]),
            category: 'tracking', categoryName: '個管追蹤',
            title: row[4], subtitle: "人員：" + (row[3] || '--'),
            detail: row[5], staff: row[3]
          });
        }
      });
    }

    const treatSheet = ss.getSheetByName(CONFIG.SHEETS.TREATMENT);
    if (treatSheet) {
      const data = treatSheet.getDataRange().getValues();
      const headers = data[0].map(normalizeHeader);
      let idxId = headers.indexOf(normalizeHeader("個案編號"));
      if (idxId === -1) idxId = 1; 
      const idxDate = headers.indexOf(normalizeHeader("治療日期"));
      const idxStaff = headers.indexOf(normalizeHeader("執行治療師"));
      const idxContent = headers.indexOf(normalizeHeader("治療內容"));
      const idxItem = headers.indexOf(normalizeHeader("治療項目"));
      
      data.slice(1).forEach(row => {
        if (String(row[idxId]).replace(/^'/, '').trim() === targetId) {
          const itemVal = (idxItem > -1 && row[idxItem]) ? row[idxItem] : "";
          result.push({
            id: 'T-' + formatDateForJSON(row[idxDate]), 
            date: formatDateForJSON(row[idxDate]),
            category: 'treatment', categoryName: '治療紀錄',
            title: "物理治療", 
            subtitle: (itemVal ? itemVal + " | " : "") + "治療師：" + (row[idxStaff] || '--'),
            detail: row[idxContent], staff: row[idxStaff]
          });
        }
      });
    }

    return result.sort((a, b) => new Date(b.date) - new Date(a.date));

  } catch (e) { throw new Error("取得總覽資料失敗: " + e.message); }
}

// 影像功能 - 修改版：從 Sheet 讀取，以便顯示備註
function getClientImages(clientId) {
  try {
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const sheet = ss.getSheetByName(CONFIG.SHEETS.IMAGE); // 連接到 Image_Gallery
    if (!sheet) return { success: false, message: "找不到影像資料表" };

    const data = sheet.getDataRange().getDisplayValues();
    const imageList = [];
    
    // 欄位定義: 0:影像ID, 1:個案編號, 2:上傳日期, 3:檔案名稱, 4:GoogleDrive檔案連結, 5:備註
    // 從第 2 行開始 (索引 1)，略過標題
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      // 比對個案編號
      if (String(row[1]).replace(/^'/, '').trim() === String(clientId).trim()) {
        const fileUrl = row[4];
        let fileId = "";
        
        // 從 URL 提取 File ID
        const idMatch = fileUrl.match(/[-\w]{25,}/);
        if (idMatch) fileId = idMatch[0];

        imageList.push({
          id: row[0],      // 影像ID
          name: row[3],    // 檔案名稱
          url: fileUrl,    // Google Drive 連結
          thumbnail: fileId ? "https://lh3.googleusercontent.com/d/" + fileId + "=s400" : "", // 縮圖
          date: row[2],    // 上傳日期
          remark: row[5]   // ★ 新增：備註 (F欄)
        });
      }
    }
    
    // 依日期排序 (新到舊)
    return { success: true, images: imageList.sort((a,b)=>b.date.localeCompare(a.date)) };

  } catch (e) { return { success: false, message: "讀取影像失敗: " + e.toString() }; }
}

// 上傳影像 - 修改版：新增備註參數並寫入 Sheet
function uploadClientImage(clientId, fileData, fileName, mimeType, remark) {
  try {
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    
    // 1. 先取得資料夾位置 (從 Client 資料表)
    const clientSheet = ss.getSheetByName(CONFIG.SHEETS.CLIENT);
    const clientData = clientSheet.getDataRange().getDisplayValues();
    let folderUrl = "";
    
    for (let i = 1; i < clientData.length; i++) {
      if (String(clientData[i][0]).replace(/^'/, '') === String(clientId)) { 
          folderUrl = clientData[i][9]; 
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
    
    // 3. 寫入資料到 Image_Gallery Sheet
    let imgSheet = ss.getSheetByName(CONFIG.SHEETS.IMAGE);
    if (!imgSheet) {
      // 若無工作表則建立
      imgSheet = ss.insertSheet(CONFIG.SHEETS.IMAGE);
      imgSheet.appendRow(["影像ID", "個案編號", "上傳日期", "檔案名稱", "GoogleDrive檔案連結", "備註"]);
    }

    const uniqueImgId = "IMG" + new Date().getTime(); // 簡單的 ID 生成
    const nowStr = Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), "yyyy-MM-dd HH:mm");
    
    imgSheet.appendRow([
      uniqueImgId,
      "'" + clientId,
      nowStr,
      fileName,
      fileUrl,
      remark || "" // ★ 寫入備註
    ]);

    return { success: true, message: "上傳成功" };
  } catch (e) { return { success: false, message: "上傳失敗: " + e.toString() }; }
}

function formatDateForJSON(dateVal) {
  if (!dateVal) return "";
  if (dateVal instanceof Date) return Utilities.formatDate(dateVal, Session.getScriptTimeZone(), "yyyy-MM-dd");
  return String(dateVal);
}