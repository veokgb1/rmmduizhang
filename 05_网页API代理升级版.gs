// ============================================================
// 05_网页API代理升级版.gs  (v5 - 二代架构完整版)
// 专门处理前端网页发来的所有 HTTP 请求
// 绝对不影响 01~04.gs 的原有 Spreadsheet 内部功能
// ⚠️ CONFIG 已在 01_配置与菜单.gs 中定义，此文件直接引用
//
// 【二代架构数据链路】
//   GitHub Pages (index.html)
//     → fetch("https://rmm.veokgb1.top/?action=xxx&token=yyy")
//     → Cloudflare Worker (原封不动转发，附加跨域头)
//     → Google Apps Script doGet/doPost (本文件)
//     → 业务函数 → Google Sheets
//
// 【doGet / doPost 仅在此文件定义，02_核心引擎.gs 不重复定义，
//   避免 GAS 多文件 doGet 冲突导致部署失效】
// ============================================================

// ════════════════════════════════════════
// Sheet 名称常量
// ════════════════════════════════════════
const SHEETS = {
  LEDGER:   '① 流水明细',
  MONTHLY:  '② 月度看板',
  ANNUAL:   '③ 年度汇总',
  TIME:     '④ 时间看板',
  CATEGORY: '⑤ 分类设置',
  ACCOUNTS: '账号管理',
  LOG:      '系统日志',
};

// ════════════════════════════════════════
// 统一返回工具
// ════════════════════════════════════════
function makeResponse(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}
function ok(data)       { return makeResponse({ ok: true,  ...data }); }
function err(msg, code) { return makeResponse({ ok: false, error: msg, code: code || 400 }); }

// ════════════════════════════════════════
// GET 入口
// ════════════════════════════════════════
function doGet(e) {
  try {
    const action = e.parameter.action || '';
    const token  = e.parameter.token  || '';

    // 健康检查（Cloudflare Worker 握手探针）
    if (action === 'ping') return ok({ msg: 'pong' });

    // 无 action 时返回在线状态（浏览器直接访问 GAS URL 时的友好提示）
    if (!action) {
      return makeResponse({
        status:  'online',
        message: 'DUKA 二代架构 Google 后端响应成功！',
        info:    '请通过 action 参数传递指令。'
      });
    }

    const user = verifyToken(token);
    if (!user) return err('token 无效或已过期，请重新登录', 401);

    switch (action) {
      case 'read_sheet':     return readSheet(e.parameter.sheet, user);
      case 'get_config':     return ok({ proPasswordHash: getProPasswordHash(user.username) });
      case 'verify_pro_pin': return verifyProPin(e.parameter.pin, user);
      case 'image_proxy':    return imageProxy(e.parameter.fileId, user);
      default:               return err('未知 action: ' + action);
    }
  } catch(ex) {
    return err('服务器内部错误: ' + ex.message, 500);
  }
}

// ════════════════════════════════════════
// POST 入口
// ════════════════════════════════════════
function doPost(e) {
  try {
    const body   = JSON.parse(e.postData.contents);
    const action = body.action || '';

    if (action === 'login') return handleLogin(body);

    const user = verifyToken(body.token);
    if (!user) return err('token 无效或已过期，请重新登录', 401);

    switch (action) {
      case 'append_rows':      return appendRows(body, user);
      case 'update_row':       return updateRow(body, user);
      case 'delete_row':       return deleteRow(body, user);
      case 'migrate_vouchers': return migrateVouchers(user);
      case 'gemini_nlp':       return geminiNLP(body, user);
      case 'gemini_ocr':       return geminiOCR(body, user);
      case 'upload_image':     return uploadImage(body, user);
      default:                 return err('未知 action: ' + action);
    }
  } catch(ex) {
    return err('服务器内部错误: ' + ex.message, 500);
  }
}

// ════════════════════════════════════════
// 登录验证
// 账号管理表列：A=用户名 B=密码 C=显示名 D=是否启用
//               E=高级功能8位密码 F=编辑权限(TRUE/FALSE)
// ════════════════════════════════════════
function handleLogin(body) {
  const username = (body.username || '').trim();
  const password = (body.password || '').trim();

  if (!username || !password) {
    writeLog(username, 'LOGIN_FAIL', '用户名或密码为空');
    return err('用户名和密码不能为空', 400);
  }

  const ss    = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEETS.ACCOUNTS);
  if (!sheet) return err('账号管理表不存在，请先建表', 500);

  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    const row      = data[i];
    const uname    = String(row[0] || '').trim();
    const passwd   = String(row[1] || '').trim();
    const dispName = String(row[2] || uname);
    const enabled  = row[3] !== false && String(row[3]).toUpperCase() !== 'FALSE';
    const proPass  = String(row[4] || '').trim();
    const rawEdit  = row[5];
    const canEdit  = (rawEdit === undefined || rawEdit === '' || rawEdit === true
                      || String(rawEdit).toUpperCase() === 'TRUE');

    if (uname === username && passwd === password) {
      if (!enabled) {
        writeLog(username, 'LOGIN_FAIL', '账号已被禁用');
        return err('账号已被禁用，请联系管理员', 403);
      }
      const token = generateToken(username);
      writeLog(username, 'LOGIN_SUCCESS', `登录成功 canEdit=${canEdit}`);
      return ok({
        token:           token,
        username:        username,
        displayName:     dispName,
        proPasswordHash: simpleHash(proPass),
        canEdit:         canEdit,
      });
    }
  }

  writeLog(username, 'LOGIN_FAIL', '用户名或密码错误');
  return err('用户名或密码错误', 401);
}

// ════════════════════════════════════════
// Token 生成与验证
// ════════════════════════════════════════
function generateToken(username) {
  const ts      = new Date().getTime();
  const payload = username + '|' + ts + '|' + CONFIG.TOKEN_SECRET;
  const hash    = Utilities.computeDigest(
    Utilities.DigestAlgorithm.MD5, payload, Utilities.Charset.UTF_8
  ).map(b => ('0' + (b & 0xFF).toString(16)).slice(-2)).join('');
  return Utilities.base64Encode(username) + ':' + ts + ':' + hash.slice(0, 16);
}

function verifyToken(token) {
  if (!token) return null;
  try {
    const parts     = token.split(':');
    if (parts.length !== 3) return null;
    const username  = Utilities.newBlob(Utilities.base64Decode(parts[0])).getDataAsString();
    const ts        = parts[1];
    const givenHash = parts[2];
    const payload   = username + '|' + ts + '|' + CONFIG.TOKEN_SECRET;
    const expected  = Utilities.computeDigest(
      Utilities.DigestAlgorithm.MD5, payload, Utilities.Charset.UTF_8
    ).map(b => ('0' + (b & 0xFF).toString(16)).slice(-2)).join('').slice(0, 16);
    if (givenHash !== expected) return null;
    return { username };
  } catch(e) {
    return null;
  }
}

// ════════════════════════════════════════
// 查询指定用户的 Pro 密码 hash（供 get_config 使用）
// ════════════════════════════════════════
function getProPasswordHash(username) {
  try {
    const sheet = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID).getSheetByName(SHEETS.ACCOUNTS);
    if (!sheet) return '';
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === username) {
        return simpleHash(String(data[i][4] || '').trim());
      }
    }
  } catch(e) {}
  return '';
}

// ════════════════════════════════════════
// Pro 密码 hash
// ════════════════════════════════════════
function simpleHash(str) {
  if (!str) return '';
  return Utilities.computeDigest(
    Utilities.DigestAlgorithm.MD5, str + CONFIG.TOKEN_SECRET, Utilities.Charset.UTF_8
  ).map(b => ('0' + (b & 0xFF).toString(16)).slice(-2)).join('');
}

// ════════════════════════════════════════
// Pro 密码验证
// ════════════════════════════════════════
function verifyProPin(pin, user) {
  if (!pin) return ok({ valid: false });
  const ss    = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEETS.ACCOUNTS);
  if (!sheet) return ok({ valid: false });
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === user.username) {
      return ok({ valid: pin === String(data[i][4] || '').trim() });
    }
  }
  return ok({ valid: false });
}

// ════════════════════════════════════════
// 读取 Sheet 数据
// ════════════════════════════════════════
function readSheet(sheetKey, user) {
  const nameMap = {
    ledger:   SHEETS.LEDGER,
    monthly:  SHEETS.MONTHLY,
    annual:   SHEETS.ANNUAL,
    time:     SHEETS.TIME,
    category: SHEETS.CATEGORY,
  };
  const sheetName = nameMap[sheetKey];
  if (!sheetName) return err('未知 sheet: ' + sheetKey);
  const sheet = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID).getSheetByName(sheetName);
  if (!sheet) return err('Sheet 不存在: ' + sheetName);
  return ok({ values: sheet.getDataRange().getValues(), sheetName });
}

// ════════════════════════════════════════
// 写入操作
// ════════════════════════════════════════
function appendRows(body, user) {
  const rows = body.rows;
  if (!rows || !rows.length) return err('rows 为空');
  const sheet = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID).getSheetByName(SHEETS.LEDGER);
  if (!sheet) return err('流水明细表不存在');
  rows.forEach(row => sheet.appendRow(row));
  writeLog(user.username, 'APPEND_ROWS', `写入 ${rows.length} 行`);
  return ok({ added: rows.length });
}

function updateRow(body, user) {
  const { rowNum, row } = body;
  if (!rowNum || !row) return err('rowNum 或 row 缺失');
  const sheet = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID).getSheetByName(SHEETS.LEDGER);
  if (!sheet) return err('流水明细表不存在');
  sheet.getRange(rowNum, 1, 1, row.length).setValues([row]);
  writeLog(user.username, 'UPDATE_ROW', `更新第 ${rowNum} 行`);
  return ok({});
}

function deleteRow(body, user) {
  const { rowNum } = body;
  if (!rowNum) return err('rowNum 缺失');
  const sheet = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID).getSheetByName(SHEETS.LEDGER);
  if (!sheet) return err('流水明细表不存在');
  sheet.deleteRow(rowNum);
  writeLog(user.username, 'DELETE_ROW', `删除第 ${rowNum} 行`);
  return ok({});
}

// ════════════════════════════════════════
// 历史凭证迁移 H→K
// ════════════════════════════════════════
function migrateVouchers(user) {
  const sheet = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID).getSheetByName(SHEETS.LEDGER);
  if (!sheet) return err('流水明细表不存在');
  const lastRow = sheet.getLastRow();
  if (lastRow < 3) return ok({ migrated: 0, skipped: 0 });

  let migrated = 0, skipped = 0;
  for (let row = 3; row <= lastRow; row++) {
    const rtv  = sheet.getRange(row, LEDGER_COLS.VOUCHER).getRichTextValue();
    if (!rtv) { skipped++; continue; }
    const text = rtv.getText();
    if (!text || text.trim() === '' || text === '无照片') { skipped++; continue; }
    if (String(sheet.getRange(row, LEDGER_COLS.VOUCHER_V2).getValue() || '').trim()) { skipped++; continue; }
    const ids = [];
    rtv.getRuns().forEach(run => {
      const url = run.getLinkUrl();
      if (url) { const m = url.match(/[-\w]{25,}/); if (m) ids.push(m[0]); }
    });
    const uniq = [...new Set(ids)];
    if (uniq.length) { sheet.getRange(row, LEDGER_COLS.VOUCHER_V2).setValue('v:' + uniq.join('|')); migrated++; }
    else skipped++;
  }
  writeLog(user.username, 'MIGRATE_VOUCHERS', `迁移 ${migrated} 行，跳过 ${skipped} 行`);
  return ok({ migrated, skipped });
}

// ════════════════════════════════════════
// 真实图片上传到 Drive
// ════════════════════════════════════════
function uploadImage(body, user) {
  const { base64, mime, filename } = body;
  if (!base64 || !mime) return err('base64 或 mime 缺失');
  if (!CONFIG.UPLOAD_FOLDER_ID || CONFIG.UPLOAD_FOLDER_ID === 'YOUR_UPLOAD_FOLDER_ID_HERE') {
    return err('后端未配置 UPLOAD_FOLDER_ID，请在 CONFIG 中填写 Drive 文件夹 ID', 500);
  }
  try {
    const decoded  = Utilities.base64Decode(base64);
    const safeExt  = mime.split('/')[1] || 'jpg';
    const safeName = filename || ('voucher_' + Utilities.formatDate(new Date(), 'Asia/Shanghai', 'yyyyMMdd_HHmmss') + '.' + safeExt);
    const blob     = Utilities.newBlob(decoded, mime, safeName);
    const folder   = DriveApp.getFolderById(CONFIG.UPLOAD_FOLDER_ID);
    const file     = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    const fileId   = file.getId();
    writeLog(user.username, 'UPLOAD_IMAGE', `文件ID: ${fileId}  文件名: ${safeName}`);
    return ok({
      fileId:       fileId,
      fileName:     file.getName(),
      thumbnailUrl: `https://drive.google.com/thumbnail?id=${fileId}&sz=w400`,
      viewUrl:      `https://drive.google.com/file/d/${fileId}/view`,
    });
  } catch(e) {
    return err('上传到 Drive 失败：' + e.message, 500);
  }
}

// ════════════════════════════════════════
// Gemini NLP
// ════════════════════════════════════════
function geminiNLP(body, user) {
  const text = body.text || '';
  if (!text) return err('text 为空');
  const today = Utilities.formatDate(new Date(), 'Asia/Shanghai', 'yyyy-MM-dd');
  let categories = '餐饮,交通,购物,娱乐,医疗,教育,住房,水电气,办公,工资,投资,奖金,其他支出,其他收入';
  try {
    const sheet = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID).getSheetByName(SHEETS.CATEGORY);
    if (sheet) {
      const vals = sheet.getDataRange().getValues().flat().filter(Boolean);
      if (vals.length) categories = vals.join(',');
    }
  } catch(e) {}

  const prompt = `你是一个专业记账助手。从用户输入中提取所有账目。
今天：${today}  分类列表：${categories}
规则：1.无明确日期用今天 2.负数金额取绝对值类型改"收入" 3.只从列表选分类 4.只输出纯JSON数组
格式：[{"date":"YYYY-MM-DD","type":"支出"|"收入","category":"分类","amount":数字,"summary":"描述"}]
用户输入："""${text}"""`;

  const result = callGemini(prompt);
  if (!result.ok) return err(result.error);
  writeLog(user.username, 'NLP', `解析：${text.slice(0, 30)}`);
  return ok({ items: result.data });
}

// ════════════════════════════════════════
// Gemini OCR
// ════════════════════════════════════════
function geminiOCR(body, user) {
  const base64 = body.base64 || '';
  const mime   = body.mime   || 'image/jpeg';
  if (!base64) return err('base64 为空');
  const prompt = `分析这张发票/收据图片，只输出纯JSON不加任何说明：
{"merchant":"商户名","amount":总金额数字,"date":"YYYY-MM-DD或null","items":["商品1"],"invoice_type":"发票类型","summary":"一句话摘要"}
不是有效票据则amount设null。`;
  const result = callGeminiVision(prompt, base64, mime);
  if (!result.ok) return err(result.error);
  writeLog(user.username, 'OCR', '图片识别');
  return ok({ aiData: result.data });
}

// ════════════════════════════════════════
// Gemini API 工具（统一使用 PRIMARY 版本，见 00_Constants.gs DUKA_CONFIG.RESOURCES.URLS.GEMINI）
// ════════════════════════════════════════
function callGemini(prompt) { return callGeminiVision(prompt, null, null); }

function callGeminiVision(prompt, base64, mime) {
  const url   = DUKA_CONFIG.RESOURCES.URLS.GEMINI.PRIMARY + '?key=' + CONFIG.GEMINI_API_KEY;
  const parts = [];
  if (base64 && mime) parts.push({ inline_data: { mime_type: mime, data: base64 } });
  parts.push({ text: prompt });
  try {
    const response = UrlFetchApp.fetch(url, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify({ contents: [{ parts }] }),
      muteHttpExceptions: true,
    });
    const code = response.getResponseCode();
    const data = JSON.parse(response.getContentText());
    if (code !== 200) return { ok: false, error: 'Gemini API 错误：' + (data.error?.message || 'HTTP ' + code) };
    let raw = data.candidates?.[0]?.content?.parts?.[0]?.text || '';
    raw = raw.replace(/```json\n?/gi, '').replace(/```\n?/gi, '').trim();
    try { return { ok: true, data: JSON.parse(raw) }; }
    catch(e) { return { ok: false, error: 'AI 返回格式无法解析：' + raw.slice(0, 100) }; }
  } catch(e) {
    return { ok: false, error: '调用 Gemini 失败：' + e.message };
  }
}

// ════════════════════════════════════════
// 图片中转代理（缩略图优先策略）
// ════════════════════════════════════════
function imageProxy(fileId, user) {
  if (!fileId) return err('fileId 缺失');

  const MAX_BYTES = 20 * 1024 * 1024;
  const THUMB_URL = 'https://drive.google.com/thumbnail?id=' + fileId + '&sz=w800';

  try {
    const thumbRes = UrlFetchApp.fetch(THUMB_URL, {
      muteHttpExceptions: true,
      headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() }
    });

    if (thumbRes.getResponseCode() === 200) {
      const blob = thumbRes.getBlob();
      const b64  = Utilities.base64Encode(blob.getBytes());
      const mime = blob.getContentType() || 'image/jpeg';
      return ok({ base64: b64, mime: mime, fileId: fileId, source: 'thumb' });
    }

    const file  = DriveApp.getFileById(fileId);
    const blob  = file.getBlob();
    const bytes = blob.getBytes();

    if (bytes.length > MAX_BYTES) {
      return err(
        '原图体积 ' + Math.round(bytes.length / 1024) + 'KB 超过限制，' +
        '请确认该文件是图片且已共享。缩略图接口状态码：' + thumbRes.getResponseCode(),
        413
      );
    }

    const b64  = Utilities.base64Encode(bytes);
    const mime = blob.getContentType() || 'image/jpeg';
    return ok({ base64: b64, mime: mime, fileId: fileId, source: 'original' });

  } catch(e) {
    return err('图片获取失败：' + e.message, 500);
  }
}

// ════════════════════════════════════════
// 系统日志
// ════════════════════════════════════════
function writeLog(username, action, detail) {
  try {
    const sheet = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID).getSheetByName(SHEETS.LOG);
    if (!sheet) return;
    const ts = Utilities.formatDate(new Date(), 'Asia/Shanghai', 'yyyy-MM-dd HH:mm:ss');
    sheet.appendRow([ts, username || '(匿名)', action, detail || '']);
  } catch(e) { console.log('日志写入失败：' + e.message); }
}