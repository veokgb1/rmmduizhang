// ==========================================
// 02_核心引擎.gs - 终极安全版 (含专属待用文件夹投递)
// ==========================================
// 【二代架构说明】
//   doGet / doPost 统一由 05_网页API代理升级版.gs 处理 Web 请求。
//   本文件只保留 Spreadsheet 内部业务逻辑函数，不再重复定义路由入口，
//   避免 GAS 多文件中出现两个 doGet 导致的函数冲突报错。
// ==========================================

function processSmartTextRecord() {
  var ui = SpreadsheetApp.getUi();
  var response = ui.prompt(
    '📝 AI 快捷记账',
    '请粘贴语音转文字，或输入一段包含多笔账单的话\n（例如：今天打车花了30，中午吃饭花了25，昨天卖二手书赚了50）',
    ui.ButtonSet.OK_CANCEL
  );
  if (response.getSelectedButton() !== ui.Button.OK) return;

  var text = response.getResponseText().trim();
  if (!text) return;

  var today = Utilities.formatDate(new Date(), "GMT+8", "yyyy-MM-dd");
  var prompt = "专业财务助手，请从以下自然语言中提取所有记账记录（可能包含多笔，请务必仔细分辨）：\n「" + text + "」\n返回严格的JSON数组：[{\"date\":\"YYYY-MM-DD\",\"type\":\"支出\"或\"收入\",\"amount\":数字,\"category\":\"分类\",\"summary\":\"摘要\"}]。注意：\n1. 无明确日期默认用" + today + "。\n2. 金额统一返回正数。\n3. 必须返回 Array 格式。";

  SpreadsheetApp.getActiveSpreadsheet().toast("🤖 AI 正在疯狂解析中...", "稍候", -1);

  var res = callGeminiAPI({ "contents": [{ "parts": [{ "text": prompt }] }] });
  if (!res.success) return ui.alert("❌ 解析失败", res.detail, ui.ButtonSet.OK);

  var aiData = res.data;
  if (!Array.isArray(aiData)) aiData = [aiData];
  if (aiData.length === 0 || aiData[0].amount === undefined) {
    return ui.alert("⚠️ 提取失败", "未找到明确的金额或账单信息。", ui.ButtonSet.OK);
  }

  var msg = "✅ AI 共识别到 " + aiData.length + " 笔记录：\n\n";
  aiData.forEach(function(item, index) {
    var dateStr = item.date || today;
    var typeStr = item.type || "支出";
    var catStr  = item.category || "未分类";
    msg += (index + 1) + ". " + dateStr + " | 【" + typeStr + "】 | " + catStr + " | ¥" + item.amount + " | " + item.summary + "\n";
  });
  msg += "\n❓ 是否确认将以上记录写入表格？";

  var confirmRes = ui.alert("🤖 识别结果确认", msg, ui.ButtonSet.YES_NO);
  if (confirmRes !== ui.Button.YES) {
    SpreadsheetApp.getActiveSpreadsheet().toast("已取消录入", "提示", 3);
    return;
  }

  var sheet    = getSheetByKeyword("流水");
  var startRow = sheet.getLastRow() + 1;

  aiData.forEach(function(item) {
    var rowDate = item.date || today;
    var amt     = Math.abs(parseFloat(item.amount) || 0);
    var type    = item.type || "支出";
    sheet.appendRow([rowDate, rowDate.substring(0, 7), type, item.category || "其他", amt, item.summary || "", "快捷文本录入", "", "人工关联"]);
  });

  var numRows = aiData.length;
  applyFormatAndValidationToNewRows(sheet, startRow, numRows);
  sheet.getRange(startRow, LEDGER_COLS.DATE, numRows, LEDGER_COLS.TOTAL_COLS).setBackground("#fff2cc");
  SpreadsheetApp.getActiveSpreadsheet().toast("✅ 成功写入 " + numRows + " 笔记录！", "提示", 5);
}

function handleBatchText(text) {
  if (!text || text.trim() === "") return;
  var today  = Utilities.formatDate(new Date(), "GMT+8", "yyyy-MM-dd");
  var prompt = "专业财务助手，从账单文本提取所有付款记录：\n「" + text + "」\n返回严格JSON数组：[{\"date\":\"YYYY-MM-DD\",\"type\":\"支出\"或\"收入\",\"amount\":数字,\"category\":\"分类\",\"summary\":\"摘要\"}]。无日期默认用" + today;

  var res = callGeminiAPI({ "contents": [{ "parts": [{ "text": prompt }] }] });
  if (!res.success) {
    SpreadsheetApp.getUi().alert("❌ 解析失败", res.detail, SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }

  var arr   = res.data;
  if (!Array.isArray(arr)) arr = [arr];
  var sheet    = getSheetByKeyword("流水");
  var count    = 0;
  var startRow = sheet.getLastRow() + 1;

  arr.forEach(function(item) {
    if (item && item.amount !== undefined) {
      var rowDate = item.date || today;
      var amt     = Math.abs(parseFloat(item.amount) || 0);
      var type    = item.type || "支出";
      sheet.appendRow([rowDate, rowDate.substring(0, 7), type, item.category || "其他", amt, item.summary || "", "批量文本录入", "", "人工关联"]);
      count++;
    }
  });

  if (count > 0) {
    applyFormatAndValidationToNewRows(sheet, startRow, count);
    sheet.getRange(startRow, LEDGER_COLS.DATE, count, LEDGER_COLS.TOTAL_COLS).setBackground("#fff2cc");
  }
  SpreadsheetApp.getActiveSpreadsheet().toast("✅ 批量录入成功！共新增 " + count + " 笔记录。", "提示", 5);
}

function insertRecord(item, source, voucherUrl) {
  var sheet = getSheetByKeyword("流水");
  sheet.appendRow([item.date, item.date.substring(0, 7), "支出", item.category, item.amount, item.summary, source, "", "人工关联"]);
  var newRow = sheet.getLastRow();
  setVoucherLink(sheet, newRow, LEDGER_COLS.VOUCHER, voucherUrl);
  addAuditNote(sheet, newRow, LEDGER_COLS.VOUCHER, item);
  applyFormatAndValidationToNewRows(sheet, newRow, 1);
  sheet.getRange(newRow, LEDGER_COLS.DATE, 1, LEDGER_COLS.TOTAL_COLS).setBackground("#e2f0d9");
  return newRow;
}

function executeAction(action, fileId, aiData, targetRow) {
  if (targetRow === -1 && action !== "NEW") return false;
  var file    = DriveApp.getFileById(fileId);
  var newName = "[永久存档]_" + aiData.date + "_" + file.getName().replace(/^\[未完结\]_/, "");

  if (action === "NEW") {
    var newRow = insertRecord(aiData, "可视对账录入", file.getUrl());
    file.moveTo(getArchiveFolder());
    file.setName(newName);
    highlightTargetRow(newRow);
    return true;
  }

  var sheet = getSheetByKeyword("流水");
  if (action === "LINK_UPDATE") {
    sheet.getRange(targetRow, LEDGER_COLS.AMOUNT).setValue(aiData.amount);
    sheet.getRange(targetRow, LEDGER_COLS.AMOUNT).setBackground("#fff2cc");
  }

  setVoucherLink(sheet, targetRow, LEDGER_COLS.VOUCHER, file.getUrl());
  addAuditNote(sheet, targetRow, LEDGER_COLS.VOUCHER, aiData);
  sheet.getRange(targetRow, LEDGER_COLS.STATUS).setValue("智能关联");
  sheet.getRange(targetRow, LEDGER_COLS.DATE, 1, LEDGER_COLS.TOTAL_COLS).setBackground("#e2f0d9");
  file.moveTo(getArchiveFolder());
  file.setName(newName);
  highlightTargetRow(targetRow);
  return true;
}

function executeConflictAction(action, newFileId, aiData, targetRow, oldFileIds) {
  var sheet   = getSheetByKeyword("流水");
  var newFile = DriveApp.getFileById(newFileId);
  var newName = "[永久存档]_" + aiData.date + "_" + newFile.getName().replace(/^\[未完结\]_/, "");

  if (action === "DELETE_NEW") {
    newFile.setTrashed(true);
    return { success: true };
  }

  if (action === "NEW_ROW") {
    var newRow = insertRecord(aiData, "冲突剥离新账", newFile.getUrl());
    newFile.moveTo(getArchiveFolder());
    newFile.setName(newName);
    highlightTargetRow(newRow);
    return { success: true };
  }

  if (action === "APPEND") {
    setVoucherLink(sheet, targetRow, LEDGER_COLS.VOUCHER, newFile.getUrl());
    addAuditNote(sheet, targetRow, LEDGER_COLS.VOUCHER, aiData);
    sheet.getRange(targetRow, LEDGER_COLS.STATUS).setValue("多图合并");
    newFile.moveTo(getArchiveFolder());
    newFile.setName(newName);
    return { success: true };
  }

  if (action === "REPLACE_RETURN" || action === "REPLACE_TRASH") {
    var pendingFolder = DriveApp.getFolderById(extractDriveId(DEFAULT_PENDING_URL));
    oldFileIds.forEach(function(id) {
      try {
        var oldF = DriveApp.getFileById(id);
        if (action === "REPLACE_TRASH") {
          oldF.setTrashed(true);
        } else {
          oldF.moveTo(pendingFolder);
          oldF.setDescription("");
          oldF.setName(oldF.getName().replace(/^\[永久存档\]_.*?_/, "").replace(/^\[未完结\]_/, ""));
        }
      } catch(e) {}
    });
    sheet.getRange(targetRow, LEDGER_COLS.VOUCHER).clearContent().clearNote();
    setVoucherLink(sheet, targetRow, LEDGER_COLS.VOUCHER, newFile.getUrl());
    addAuditNote(sheet, targetRow, LEDGER_COLS.VOUCHER, aiData);
    sheet.getRange(targetRow, LEDGER_COLS.STATUS).setValue("替换关联");
    sheet.getRange(targetRow, LEDGER_COLS.DATE, 1, LEDGER_COLS.TOTAL_COLS).setBackground("#e2f0d9");
    newFile.moveTo(getArchiveFolder());
    newFile.setName(newName);
    return { success: true };
  }

  return { success: false, error: "未知指令" };
}

function executeForceBindSelected(action, fileId, aiData, doUpdateAmount) {
  try {
    var sheet = SpreadsheetApp.getActiveSheet();
    if (!sheet || sheet.getName().indexOf("流水") === -1) {
      sheet = getSheetByKeyword("流水");
    }
    if (!sheet) return { success: false, error: "❌ 找不到流水表！" };

    var range    = sheet.getActiveRange();
    var startRow = range.getRow();
    var numRows  = range.getNumRows();

    if (startRow < 3) {
      return { success: false, error: "❌ 绑定失败！\n原因：未选中有效行。\n👉 请先在表格里【高亮选中】要绑定的行！" };
    }

    var file    = DriveApp.getFileById(fileId);
    var dateStr = aiData.date || Utilities.formatDate(new Date(), "GMT+8", "yyyy-MM-dd");

    for (var i = 0; i < numRows; i++) {
      var targetRow = startRow + i;

      if (doUpdateAmount && aiData.amount !== undefined) {
        sheet.getRange(targetRow, LEDGER_COLS.AMOUNT).setValue(aiData.amount);
      }

      setVoucherLink(sheet, targetRow, LEDGER_COLS.VOUCHER, file.getUrl());
      addAuditNote(sheet, targetRow, LEDGER_COLS.VOUCHER, aiData);

      if (action === 'DONE') {
        sheet.getRange(targetRow, LEDGER_COLS.STATUS).setValue("人工关联");
        sheet.getRange(targetRow, LEDGER_COLS.DATE, 1, LEDGER_COLS.TOTAL_COLS).setBackground("#e2f0d9");
      } else if (action === 'TAG') {
        sheet.getRange(targetRow, LEDGER_COLS.STATUS).setValue("🔖 待续关联");
        sheet.getRange(targetRow, LEDGER_COLS.DATE, 1, LEDGER_COLS.TOTAL_COLS).setBackground("#fff2cc");
      }
    }

    if (action === 'DONE') {
      var newName = "[永久存档]_" + dateStr + "_" + file.getName().replace(/^\[未完结\]_/, "");
      file.moveTo(getArchiveFolder());
      file.setName(newName);
    } else if (action === 'TAG') {
      var tagFolder = DriveApp.getFolderById(FOLDER_TAGGED_PENDING_ID);
      if (!file.getName().startsWith("[未完结]")) {
        file.setName("[未完结]_" + file.getName());
      }
      file.moveTo(tagFolder);
    }

    sheet.getRange(startRow, LEDGER_COLS.DATE, numRows, LEDGER_COLS.TOTAL_COLS).activate();
    return { success: true };

  } catch (e) {
    return { success: false, error: "后台执行出错: " + e.message };
  }
}

function executePartialUnbind(row, idsToUnbind) {
  var sheet       = getSheetByKeyword("流水");
  var voucherCell = sheet.getRange(row, LEDGER_COLS.VOUCHER);
  var rtv         = voucherCell.getRichTextValue();
  var runs        = rtv ? rtv.getRuns() : [];
  var links       = [];

  runs.forEach(function(run) {
    if (run.getLinkUrl()) links.push(run.getLinkUrl());
  });

  var allFileIds = extractFileIds(links.join(" ") + " " + (voucherCell.getFormula() || voucherCell.getValue()));
  var notesData  = [];
  try { notesData = JSON.parse(voucherCell.getNote() || "[]"); } catch(e) {}
  if (!Array.isArray(notesData)) notesData = [notesData];

  var remainingIds   = [];
  var remainingNotes = [];
  var pendingFolder  = DriveApp.getFolderById(extractDriveId(DEFAULT_PENDING_URL));

  for (var i = 0; i < allFileIds.length; i++) {
    var currentId = allFileIds[i];
    if (idsToUnbind.indexOf(currentId) !== -1) {
      try {
        var f = DriveApp.getFileById(currentId);
        f.moveTo(pendingFolder);
        f.setDescription("");
        f.setName(f.getName().replace(/^\[永久存档\]_.*?_/, "").replace(/^\[未完结\]_/, ""));
      } catch(e) {}
    } else {
      remainingIds.push(currentId);
      if (notesData[i]) remainingNotes.push(notesData[i]);
    }
  }

  voucherCell.clearContent().clearNote();
  sheet.getRange(row, LEDGER_COLS.EXTRA).setValue("");

  if (remainingIds.length === 0) {
    sheet.getRange(row, LEDGER_COLS.STATUS).setValue("未关联");
    sheet.getRange(row, LEDGER_COLS.DATE, 1, LEDGER_COLS.TOTAL_COLS).setBackground(null);
  } else {
    var builder     = SpreadsheetApp.newRichTextValue();
    var displayText = "";
    remainingIds.forEach(function(id, idx) {
      if (idx > 0) displayText += " | ";
      displayText += "🖼️ 凭证" + (idx + 1);
    });
    builder.setText(displayText);

    var currentIndex = 0;
    remainingIds.forEach(function(id, idx) {
      var label = "🖼️ 凭证" + (idx + 1);
      builder.setLinkUrl(currentIndex, currentIndex + label.length, "https://drive.google.com/file/d/" + id + "/view");
      currentIndex += label.length + 3;
    });
    voucherCell.setRichTextValue(builder.build());
    voucherCell.setNote(JSON.stringify(remainingNotes));
  }
  return { success: true, count: idsToUnbind.length, remain: remainingIds.length };
}

function buildPrompt() {
  return "专业财务助手，提取独立实付记录。不要划掉的原价。无明确日期date留空。返回JSON数组：[{\"date\":\"YYYY-MM-DD或空\",\"amount\":数字,\"category\":\"分类\",\"summary\":\"商品名\"}]";
}

function callGeminiAPI(payload) {
  var apiUrl = DUKA_CONFIG.RESOURCES.URLS.GEMINI.PRIMARY + "?key=" + API_KEY;
  try {
    if (!API_KEY) {
      return { success: false, detail: "Gemini API Key 未配置，请检查 Script Properties 中的 GEMINI_API_KEY" };
    }
    var response = UrlFetchApp.fetch(apiUrl, {
      "method": "post",
      "contentType": "application/json",
      "payload": JSON.stringify(payload),
      "muteHttpExceptions": true
    });
    var json    = JSON.parse(response.getContentText());
    if (response.getResponseCode() !== 200) {
      return { success: false, detail: "Gemini API 错误: " + ((json.error && json.error.message) || ("HTTP " + response.getResponseCode())) };
    }
    var candidate = json && json.candidates && json.candidates[0];
    var part = candidate && candidate.content && candidate.content.parts && candidate.content.parts[0];
    if (!part || !part.text) {
      return { success: false, detail: "Gemini 返回结构异常，未找到 candidates[0].content.parts[0].text" };
    }
    var resText = part.text.replace(/```json/g, "").replace(/```/g, "").trim();
    return { success: true, data: JSON.parse(resText) };
  } catch (e) {
    return { success: false, detail: e.toString() };
  }
}

function analyzeSingleImage(fileId) {
  try {
    var file   = DriveApp.getFileById(fileId);
    var base64 = Utilities.base64Encode(file.getBlob().getBytes());
    var mime   = file.getMimeType();
    var res    = callGeminiAPI({ "contents": [{ "parts": [{ "text": buildPrompt() }, { "inline_data": { "mime_type": mime, "data": base64 } }] }] });

    if (!res.success) return res;

    var aiResults = res.data;
    if (!Array.isArray(aiResults)) aiResults = [aiResults];

    var sheet        = getSheetByKeyword("流水");
    var data         = sheet.getDataRange().getValues();
    var bestMatchRow = -1;
    var highestScore = 0;
    var matchInfoObj = null;
    var bestAiItem   = aiResults[0] || { date: "", amount: 0, category: "未知", summary: "无金额" };

    aiResults.forEach(function(aiItem) {
      if (aiItem && aiItem.amount !== undefined) aiItem.amount = Math.abs(parseFloat(aiItem.amount) || 0);
      for (var i = 2; i < data.length; i++) {
        var rowData = {
          date:     data[i][LEDGER_COLS.DATE - 1] instanceof Date ? Utilities.formatDate(data[i][LEDGER_COLS.DATE - 1], "GMT+8", "yyyy-MM-dd") : data[i][LEDGER_COLS.DATE - 1],
          amount:   data[i][LEDGER_COLS.AMOUNT - 1],
          category: data[i][LEDGER_COLS.CATEGORY - 1],
          summary:  data[i][LEDGER_COLS.SUMMARY - 1]
        };
        var score = calculateMatchScore(aiItem, rowData);
        if (score >= 60 && score > highestScore) {
          highestScore = score;
          bestMatchRow = i + 1;
          matchInfoObj = rowData;
          bestAiItem   = aiItem;
        }
      }
    });

    if (highestScore < 60 && aiResults.length > 0) {
      var maxAmtItem = aiResults[0];
      for (var j = 1; j < aiResults.length; j++) {
        if ((aiResults[j].amount || 0) > (maxAmtItem.amount || 0)) maxAmtItem = aiResults[j];
      }
      bestAiItem = maxAmtItem;
    }

    var hasConflict   = false;
    var conflictFileIds = [];

    if (highestScore >= 60 && bestMatchRow !== -1) {
      var voucherContent = data[bestMatchRow - 1][LEDGER_COLS.VOUCHER - 1];
      if (voucherContent && voucherContent.toString().trim() !== "") {
        hasConflict = true;
        var rtv      = sheet.getRange(bestMatchRow, LEDGER_COLS.VOUCHER).getRichTextValue();
        var runs     = rtv ? rtv.getRuns() : [];
        var cellLinks = [];
        runs.forEach(function(run) { if (run.getLinkUrl()) cellLinks.push(run.getLinkUrl()); });
        var cellContentStr = cellLinks.join(" ") + " " + sheet.getRange(bestMatchRow, LEDGER_COLS.VOUCHER).getFormula();
        conflictFileIds = extractFileIds(cellContentStr);
      }
    }

    return {
      success:      true,
      imgSrc:       "data:" + mime + ";base64," + base64,
      aiData:       bestAiItem,
      matchRow:     bestMatchRow,
      matchInfo:    matchInfoObj,
      hasConflict:  hasConflict,
      conflictIds:  conflictFileIds
    };
  } catch (e) {
    return { success: false, detail: "读取异常: " + e.toString() };
  }
}

function calculateMatchScore(aiData, rowData) {
  var score   = 0;
  var aiAmt   = Math.abs(parseFloat(aiData.amount) || 0);
  var rowAmt  = Math.abs(parseFloat(rowData.amount) || 0);
  var absDiff = Math.abs(aiAmt - rowAmt);

  if (absDiff <= 0.01 && aiAmt !== 0) {
    score += 60;
  } else if (aiAmt !== 0 && rowAmt !== 0) {
    if (Math.max(aiAmt, rowAmt) <= 100 && absDiff <= 5) {
      score += 42;
    } else if (absDiff / Math.max(aiAmt, rowAmt) <= 0.20) {
      score += 50 * (1 - (absDiff / Math.max(aiAmt, rowAmt) / 0.20));
    }
  }

  if (aiData.date && rowData.date) {
    var daysDiff = Math.abs(new Date(aiData.date) - new Date(rowData.date)) / (1000 * 3600 * 24);
    if (daysDiff === 0)        score += 20;
    else if (daysDiff <= 3)    score += 15;
    else if (daysDiff <= 15)   score += 10;
  }

  var aiText  = (aiData.summary  || "").toString().toLowerCase().replace(/[\s,，.。!！?？、()（）]/g, '');
  var rowText = (rowData.summary || "").toString().toLowerCase().replace(/[\s,，.。!！?？、()（）]/g, '');

  if (aiText.length > 0 && rowText.length > 0) {
    if (aiText.indexOf(rowText) !== -1 || rowText.indexOf(aiText) !== -1) {
      score += 65;
    } else {
      var matchCount = 0;
      for (var i = 0; i < rowText.length; i++) {
        if (aiText.indexOf(rowText[i]) !== -1) matchCount++;
      }
      var matchRate = matchCount / rowText.length;
      if (matchRate >= 0.8)      score += 65;
      else if (matchRate >= 0.5) score += 30;
    }
  }
  return score;
}

function fastScanDuplicates(folderIds) {
  try {
    var tk  = ScriptApp.getOAuthToken();
    var all = [];
    for (var i = 0; i < folderIds.length; i++) {
      var fid = folderIds[i];
      var pt  = null;
      do {
        var url = "https://www.googleapis.com/drive/v3/files?q='" + fid + "'+in+parents+and+trashed=false&fields=nextPageToken,files(id,md5Checksum,mimeType)&pageSize=1000";
        var res = UrlFetchApp.fetch(url, { headers: { "Authorization": "Bearer " + tk }, muteHttpExceptions: true });
        if (res.getResponseCode() === 200) {
          var d = JSON.parse(res.getContentText());
          if (d.files) {
            d.files.forEach(function(f) {
              if (f.mimeType.indexOf("image/") !== -1) all.push(f);
            });
          }
          pt = d.nextPageToken;
        } else {
          break;
        }
      } while (pt);
    }
    var map  = {};
    var dups = [];
    all.forEach(function(f) {
      if (f.md5Checksum) {
        if (!map[f.md5Checksum]) map[f.md5Checksum] = [];
        map[f.md5Checksum].push(f);
      }
    });
    for (var m in map) {
      if (map[m].length > 1) dups.push(map[m]);
    }
    return { success: true, data: dups };
  } catch(e) {
    return { success: false, error: e.toString() };
  }
}

function trashDuplicateFiles(ids) {
  ids.forEach(function(id) {
    try { DriveApp.getFileById(id).setTrashed(true); } catch(e) {}
  });
}

function syncToGitee() {
  var GITEE_TOKEN  = DUKA_CONFIG.AUTH.GITEE.TOKEN;
  var GITEE_OWNER  = DUKA_CONFIG.AUTH.GITEE.OWNER;
  var GITEE_REPO   = DUKA_CONFIG.AUTH.GITEE.REPO;
  var GITEE_BRANCH = DUKA_CONFIG.AUTH.GITEE.BRANCH;

  var sheet = getSheetByKeyword("流水");
  var data  = sheet.getDataRange().getValues();

  var htmlContent = "<html><head><meta charset='utf-8'><title>财迷账本离线版</title>";
  htmlContent += "<style>table{border-collapse:collapse;width:100%} th,td{border:1px solid #ddd;padding:8px} th{background:#f2f2f2}</style></head><body>";
  htmlContent += "<h1>财迷账本明细</h1><table>";
  for (var i = 1; i < data.length; i++) {
    htmlContent += "<tr>";
    for (var j = 0; j < data[i].length; j++) {
      var val = data[i][j];
      if (j === LEDGER_COLS.VOUCHER - 1 && val.toString().includes("http")) val = "<a href='" + val + "'>查看凭证</a>";
      htmlContent += "<td>" + val + "</td>";
    }
    htmlContent += "</tr>";
  }
  htmlContent += "</table></body></html>";

  var base64Content = Utilities.base64Encode(htmlContent);
  var apiUrl        = "https://gitee.com/api/v5/repos/" + GITEE_OWNER + "/" + GITEE_REPO + "/contents/index.html";
  var payload       = { "access_token": GITEE_TOKEN, "content": base64Content, "message": "Update ledger from Google AI", "branch": GITEE_BRANCH };
  var getRes        = UrlFetchApp.fetch(apiUrl + "?access_token=" + GITEE_TOKEN, { muteHttpExceptions: true });
  if (getRes.getResponseCode() === 200) payload.sha = JSON.parse(getRes.getContentText()).sha;
  var options = { "method": "post", "contentType": "application/json", "payload": JSON.stringify(payload), "muteHttpExceptions": true };
  var res     = UrlFetchApp.fetch(apiUrl, options);
  if (res.getResponseCode() === 201 || res.getResponseCode() === 200) {
    SpreadsheetApp.getUi().alert("🎉 同步成功！");
  } else {
    SpreadsheetApp.getUi().alert("❌ 同步失败：" + res.getContentText());
  }
}
