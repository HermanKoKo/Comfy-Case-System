/**
 * 康飛運醫 | 個案管理系統 - 全域配置設定
 */
const CONFIG = {
  // 1. 試算表 ID
  SPREADSHEET_ID: '1LMhlQGyXNXq9Teqm0_W0zU9NbQlVCHKLDL0mSOiDomc', 
  
  // 2. Google Drive 資料夾 ID
  PARENT_FOLDER_ID: '1NIsNHALeSSVm60Yfjc9k-u30A42CuZw8',
  
  // 3. 工作表名稱定義 (務必與 Google Sheets 下方標籤一致)
  SHEETS: {
    CLIENT: 'Client_Basic_Info',      // 基本資料
    TREATMENT: 'Treatment_Logs',      // 治療紀錄
    DOCTOR: 'Doctor_Consultations',   // 醫師看診
    TRACKING: 'Case_Tracking',        // 個管追蹤
    SYSTEM: 'System',                 // 系統設定 (存放治療師名單)
    IMAGE: 'Image_Gallery'            // (備用)
  }
};