const CONFIG = {
  GEMINI_API_KEY:   DUKA_CONFIG.AUTH.GEMINI_API_KEY,
  SHEETS_API_KEY:   DUKA_CONFIG.AUTH.SHEETS_API_KEY,
  SPREADSHEET_ID:   DUKA_CONFIG.RESOURCES.SPREADSHEET.ID,
  TOKEN_SECRET:     DUKA_CONFIG.AUTH.TOKEN_SECRET,
  UPLOAD_FOLDER_ID: DUKA_CONFIG.RESOURCES.DRIVE.UPLOAD_FOLDER_ID,
  MAX_BYTES:        DUKA_CONFIG.SETTINGS.LIMITS.MAX_BYTES,
};

// ==========================================
// 01_配置与菜单.gs
// ==========================================
// 兼容旧代码的全局变量名：不再硬编码，统一由 DUKA_CONFIG 提供
var API_KEY = CONFIG.GEMINI_API_KEY;
var DEFAULT_PENDING_URL = DUKA_CONFIG.RESOURCES.DRIVE.DEFAULT_PENDING_URL;
var FOLDER_ARCHIVED_ID = DUKA_CONFIG.RESOURCES.DRIVE.FOLDER_ARCHIVED_ID;

// 🌟 新增：你刚刚创建的“专属未完结”中转站 ID
var FOLDER_TAGGED_PENDING_ID = DUKA_CONFIG.RESOURCES.DRIVE.FOLDER_TAGGED_PENDING_ID;

var DEDUPE_FOLDERS = DUKA_CONFIG.RESOURCES.DRIVE.DEDUPE_FOLDERS;

function onOpen() {
  initDashboardHeader();
  
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('🤖 AI 记账助手 (Pro)')
      .addItem('🖥️ 1. 批量对账台 (常规快速处理)', 'openBatchMatchingUI')
      .addSeparator()
      .addItem('📊 2. 检查关联性 (按【行内容】排查)', 'showRowCorrelationSidebar')
      .addItem('🖼️ 3. 检查关联性 (按【凭证】排查)', 'showVoucherCorrelationModal')
      .addSeparator()
      .addItem('📝/🎙️ 4. 快捷记账 (一句话文字/语音)', 'processSmartTextRecord')
      .addItem('🚀 5. 批量粘贴补录 (大段文字解析)', 'processBatchText')
      .addSeparator()
      .addItem('📅 6. 一键按日期排序 (整理账本)', 'sortLedgerByDate')
      .addItem('🔍 7. 查杀重复凭证 (批量收割版)', 'openDeduplicationUI')
      .addItem('🧹 8. 精准解绑选中行 (可单选退回)', 'forceUnbindSelectedRow')
      .addItem('⚖️ 9. 独立双屏断案法庭 (审理冲突)', 'openConflictCourtUI')
      .addItem('🌐 一键发布到 Gitee (国内免翻墙)', 'syncToGitee')
      .addToUi();
}

