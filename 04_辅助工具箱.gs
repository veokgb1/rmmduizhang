// ==========================================
// 04_辅助工具箱.gs - 终极完整版 (修复打圈Bug)
// ==========================================

/**
 * 🌟 初始化首行看板 (G1/H1 结余监控)
 */
function initDashboardHeader() {
  var sheet = getSheetByKeyword("流水");
  if (!sheet) return;

  var g1 = sheet.getRange("G1");
  g1.setValue("💰 净结余 (收入-支出)")
    .setFontWeight("bold")
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle");

  var h1 = sheet.getRange("H1");
  h1.setFormula("=D1-F1")
    .setFontWeight("bold")
    .setNumberFormat('¥#,##0.00;[Red]¥-#,##0.00') 
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle");

  var rules = sheet.getConditionalFormatRules();
  var greenRule = SpreadsheetApp.newConditionalFormatRule()
    .whenNumberGreaterThan(0).setBackground("#e2f0d9").setFontColor("#38761d").setRanges([h1]).build();
  var redRule = SpreadsheetApp.newConditionalFormatRule()
    .whenNumberLessThan(0).setBackground("#fce8e6").setFontColor("#990000").setRanges([h1]).build();
    
  sheet.setConditionalFormatRules([greenRule, redRule]);
}

/**
 * 格式与下拉验证继承器 (锁定第三行作为母版)
 */
function applyFormatAndValidationToNewRows(sheet, startRow, numRows) {
  if (startRow <= 3) return; 
  var templateRow = 3; 
  var lastCol = sheet.getLastColumn();
  var sourceRange = sheet.getRange(templateRow, 1, 1, lastCol);
  var targetRange = sheet.getRange(startRow, 1, numRows, lastCol);
  
  sourceRange.copyTo(targetRange, SpreadsheetApp.CopyPasteType.PASTE_DATA_VALIDATION, false);
  sourceRange.copyTo(targetRange, SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
}

/**
 * 一键按时间排序功能
 */
function sortLedgerByDate() {
  var sheet = getSheetByKeyword("流水");
  if (!sheet) return;
  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  if (lastRow <= 2) return; 
  var dataRange = sheet.getRange(3, 1, lastRow - 2, lastCol);
  dataRange.sort({column: 1, ascending: true});
  SpreadsheetApp.getActiveSpreadsheet().toast("📅 账单已按时间重新排序！", "整理完毕", 3);
}

/**
 * 🚨 修复打圈的核心：从多个文件夹批量获取待处理图片
 */
function getFilesFromMultipleFolders(folderIds) {
  var allFiles = [];
  folderIds.forEach(function(id) {
    try {
      var folder = DriveApp.getFolderById(id);
      var files = folder.getFiles();
      while (files.hasNext()) {
        var file = files.next();
        var mime = file.getMimeType();
        // 只读取图片，且避开已经存档的文件
        if (mime.indexOf("image/") !== -1 && file.getName().indexOf("[永久存档]") === -1) {
          allFiles.push({
            id: file.getId(),
            name: file.getName()
          });
        }
      }
    } catch (e) {
      console.log("文件夹读取失败 ID: " + id + " 错误: " + e.message);
    }
  });
  return allFiles;
}

// --- 以下为底层辅助方法 ---

function getSheetByKeyword(keyword) { 
  var ss = SpreadsheetApp.getActiveSpreadsheet().getSheets(); 
  for(var i = 0; i < ss.length; i++) {
    if(ss[i].getName().indexOf(keyword) !== -1) return ss[i]; 
  }
  return null; 
}

function extractDriveId(urlStr) { 
  var match = urlStr.match(/[-\w]{25,}/); 
  return match ? match[0] : null; 
}

function extractFileIds(textStr) { 
  var ids = [];
  var regex = /[-\w]{25,}/g;
  var match; 
  while((match = regex.exec(textStr)) !== null) ids.push(match[0]); 
  return ids; 
}

function getAllPendingPhotos(folder, list) { 
  var files = folder.getFiles();
  var folderName = folder.getName(); 
  while(files.hasNext()) { 
    var file = files.next(); 
    if(file.getMimeType().indexOf("image/") !== -1 && file.getName().indexOf("[永久存档]") === -1) {
      list.push({id: file.getId(), name: file.getName(), url: file.getUrl(), folder: folderName}); 
    }
  } 
  var subFolders = folder.getFolders(); 
  while(subFolders.hasNext()) getAllPendingPhotos(subFolders.next(), list); 
}

function highlightTargetRow(rowNum) { 
  var sheet = getSheetByKeyword("流水"); 
  if(sheet && rowNum > 0) sheet.getRange(rowNum, 1, 1, 9).activate(); 
}

function addAuditNote(sheet, row, col, dataObj) { 
  var cell = sheet.getRange(row, col);
  var notes = []; 
  try { if(cell.getNote()) notes = JSON.parse(cell.getNote()); } catch(e) {} 
  notes.push(dataObj); 
  cell.setNote(JSON.stringify(notes)); 
}

function setVoucherLink(sheet, row, col, url) { 
  var range = sheet.getRange(row, col);
  var rtv = range.getRichTextValue();
  var oldText = rtv ? rtv.getText() : ""; 
  if(!oldText || oldText === "无照片") { 
    range.setRichTextValue(SpreadsheetApp.newRichTextValue().setText("🖼️ 凭证1").setLinkUrl(url).build()); 
  } else { 
    var countMatch = oldText.match(/凭证/g);
    var currentCount = countMatch ? countMatch.length : 0;
    var newText = oldText + " | 🖼️ 凭证" + (currentCount + 1); 
    var builder = SpreadsheetApp.newRichTextValue().setText(newText); 
    if (rtv) {
      rtv.getRuns().forEach(function(rn) {
        if(rn.getLinkUrl()) builder.setLinkUrl(rn.getStartIndex(), rn.getEndIndex(), rn.getLinkUrl());
      });
    }
    builder.setLinkUrl(oldText.length + 3, newText.length, url); 
    range.setRichTextValue(builder.build()); 
  } 
}

function getArchiveFolder() { 
  try { return DriveApp.getFolderById(DUKA_CONFIG.RESOURCES.DRIVE.FOLDER_ARCHIVED_ID); } 
  catch(e) { return DriveApp.createFolder("00_存档"); } 
}

