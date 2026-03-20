// ==========================================
// 03_前端UI弹窗.gs - 终极防报错版 (内嵌三选一面板)
// ==========================================

function openBatchMatchingUI() {
  var tpl = HtmlService.createTemplate(getNormalMatchHtmlContent());
  tpl.foldersStr = JSON.stringify(DEDUPE_FOLDERS);
  SpreadsheetApp.getUi().showModelessDialog(tpl.evaluate().setWidth(1100).setHeight(780), '📸 智能对账台 (常规流水线)');
}

function openConflictCourtUI() {
  var tpl = HtmlService.createTemplate(getCourtHtmlContent());
  tpl.foldersStr = JSON.stringify(DEDUPE_FOLDERS);
  SpreadsheetApp.getUi().showModelessDialog(tpl.evaluate().setWidth(1100).setHeight(780), '⚖️ 独立双屏断案法庭 (审理冲突)');
}

function processBatchText() {
  var ui = SpreadsheetApp.getUi();
  var html = [
    '<html><body style="font-family:sans-serif;padding:10px;">',
    '<textarea id="txt" style="width:100%;height:150px;margin-bottom:10px;padding:8px;box-sizing:border-box;" placeholder="请粘贴大段账单文本..."></textarea><br>',
    '<button onclick="this.innerText=\'正在解析中...\';this.disabled=true;google.script.run.withSuccessHandler(google.script.host.close).handleBatchText(document.getElementById(\'txt\').value)" style="width:100%;padding:12px;background:#1a73e8;color:#fff;border:none;border-radius:4px;cursor:pointer;font-weight:bold;">🚀 批量解析并录入</button>',
    '</body></html>'
  ].join('');
  ui.showModalDialog(HtmlService.createHtmlOutput(html).setWidth(500).setHeight(280), '🚀 批量粘贴补录');
}

function forceUnbindSelectedRow() {
  var sheet = SpreadsheetApp.getActiveSheet(), ui = SpreadsheetApp.getUi();
  if (sheet.getName().indexOf("流水") === -1) { return ui.alert("⚠️ 错误", "请切到流水表！", ui.ButtonSet.OK); }
  var row = sheet.getActiveCell().getRow(); if (row < 3) return ui.alert("⚠️ 错误", "选有效行！", ui.ButtonSet.OK);
  var voucherCell = sheet.getRange(row, 8), rtv = voucherCell.getRichTextValue(), runs = rtv ? rtv.getRuns() : [], links = [];
  runs.forEach(function(run) { if(run.getLinkUrl()) links.push(run.getLinkUrl()); });
  var fileIds = extractFileIds(links.join(" ") + " " + (voucherCell.getFormula() || voucherCell.getValue())); 
  if (fileIds.length === 0) return ui.alert("⚠️ 失败", "没找到照片ID。", ui.ButtonSet.OK);
  var tpl = HtmlService.createTemplate(getUnbindHtmlContent()); tpl.fileIdsStr = JSON.stringify(fileIds); tpl.rowNum = row;
  ui.showModelessDialog(tpl.evaluate().setWidth(900).setHeight(650), '🧹 精准解绑操作台');
}

function openDeduplicationUI() {
  var ui = SpreadsheetApp.getUi(), checkboxesHtml = "";
  for (var i = 0; i < DEDUPE_FOLDERS.length; i++) { 
    checkboxesHtml += "<label class='folder-item' style='margin-bottom:8px;display:block;'><input type='checkbox' class='folder-cb' value='" + DEDUPE_FOLDERS[i].id + "' checked> <b>📁 " + DEDUPE_FOLDERS[i].name + "</b></label>"; 
  }
  var rawHtml = [
    '<!DOCTYPE html><html><body style="font-family:sans-serif;padding:20px;background:#f0f2f5;"><h2>🗂️ 勾选仓库</h2><div id="f-list">',
    checkboxesHtml,
    '</div><button onclick="start()" style="padding:15px;background:#1a73e8;color:white;border:none;border-radius:6px;width:100%;font-size:16px;margin-top:20px;cursor:pointer;">🚀 开始扫描</button><div id="res" style="margin-top:20px;"></div>',
    '<script>function start(){ document.getElementById("res").innerHTML="⏳ 扫描中..."; var cbs=document.querySelectorAll(".folder-cb:checked"), ids=[]; cbs.forEach(c=>ids.push(c.value)); google.script.run.withSuccessHandler(r=>{ if(r.data.length===0) document.getElementById("res").innerHTML="🎉 无重复"; else document.getElementById("res").innerHTML="⚠️ 发现 "+r.data.length+" 组重复（为节省性能，请使用原完整去重界面代码，此处精简展示）"; }).fastScanDuplicates(ids); }</script></body></html>'
  ].join('');
  ui.showModelessDialog(HtmlService.createHtmlOutput(rawHtml).setWidth(600).setHeight(500), '🔍 去重雷达');
}

function showVoucherCorrelationModal() {
  var s=getSheetByKeyword("流水"), lr=Math.max(s.getLastRow(),3), f=s.getRange(3,8,lr-2,1).getFormulas(), v=s.getRange(3,8,lr-2,1).getValues(), n=s.getRange(3,8,lr-2,1).getNotes(), rt=s.getRange(3,8,lr-2,1).getRichTextValues(), a=s.getRange(3,5,lr-2,1).getValues(), map={};
  for(var i=0;i<f.length;i++){ var l=[]; if(rt[i][0]) rt[i][0].getRuns().forEach(r=>{if(r.getLinkUrl()) l.push(r.getLinkUrl());}); var ids=extractFileIds(l.join(" ")+" "+f[i][0]+" "+v[i][0]); if(!ids.length)continue; var ad=[]; try{ad=JSON.parse(n[i][0]||"[]");}catch(e){} if(!Array.isArray(ad))ad=[ad]; var ra=parseFloat(a[i][0])||0; ids.forEach((id,x)=>{ if(!map[id]) map[id]={id:id,t:ad[x]?parseFloat(ad[x].amount):0,r:[]}; map[id].r.push({n:i+3,a:ra}); }); }
  var h='<html><body style="font-family:sans-serif;padding:20px;"><div style="display:grid;grid-template-columns:repeat(auto-fill,minmax(300px,1fr));gap:15px;">';
  for(var k in map){ var x=map[k], u=0; x.r.forEach(r=>u+=r.a); var d=Math.abs(x.t-u), c=d<=0.05?"#34a853":d>0.05?"#fbbc04":"#ea4335"; h+='<div style="border-top:5px solid '+c+';padding:15px;box-shadow:0 2px 5px rgba(0,0,0,0.1);"><img src="'+DUKA_CONFIG.RESOURCES.URLS.GOOGLE_DRIVE.THUMBNAIL+'?id='+x.id+'&sz='+DUKA_CONFIG.SETTINGS.LIMITS.THUMB_SIZE_UPLOAD_RETURN+'" style="width:100%;height:120px;object-fit:cover;"/><br>面额:¥'+x.t+' | 分摊:¥'+u+'</div>'; }
  SpreadsheetApp.getUi().showModelessDialog(HtmlService.createHtmlOutput(h+"</div></body></html>").setWidth(1000).setHeight(700), '🔍 凭证雷达');
}

function showRowCorrelationSidebar() {
  var s=getSheetByKeyword("流水"), lr=Math.max(s.getLastRow(),3), d=s.getRange(3,1,lr-2,9).getValues(), n=s.getRange(3,8,lr-2,1).getNotes(), h='<style>body{font-family:sans-serif;font-size:13px;padding:10px;}</style><h3>排查</h3>';
  for(var i=0;i<d.length;i++){ var ra=parseFloat(d[i][4])||0; if(!d[i][7]){ h+='<div style="color:gray;border:1px solid #eee;margin-bottom:5px;padding:5px;">第'+(i+3)+'行 | ¥'+ra+' (未关联)</div>'; continue; } var ad=[]; try{ad=JSON.parse(n[i][0]);}catch(e){} if(!Array.isArray(ad))ad=[ad]; var pa=0; ad.forEach(x=>pa+=(parseFloat(x.amount)||0)); var c=Math.abs(ra-pa)<=0.05?"green":"red"; h+='<div style="color:'+c+';border:1px solid #eee;margin-bottom:5px;padding:5px;">第'+(i+3)+'行 | 表¥'+ra+' | 凭证¥'+pa+'</div>'; }
  SpreadsheetApp.getUi().showSidebar(HtmlService.createHtmlOutput(h));
}

function getNormalMatchHtmlContent() {
  return [
    '<!DOCTYPE html><html><head><style>',
    'body{font-family:sans-serif;background:#f0f2f5;margin:0;padding:15px;height:100vh;box-sizing:border-box;overflow:hidden;}',
    '#config-page{background:#fff;border-radius:8px;padding:30px;max-width:600px;margin:20px auto;overflow-y:auto;max-height:90vh;}',
    '.folder-item{display:flex;align-items:center;padding:12px;background:#f8f9fa;margin-bottom:8px;border-radius:6px;cursor:pointer;border:1px solid #eee;}',
    '.folder-item input{width:20px;height:20px;margin-right:15px;cursor:pointer;}',
    '.custom-input{width:100%;padding:12px;margin-top:15px;border:1px solid:#ccc;border-radius:6px;box-sizing:border-box;}',
    '.btn-start-scan{background:#1a73e8;color:white;padding:15px;border:none;border-radius:6px;width:100%;font-size:16px;font-weight:bold;margin-top:20px;cursor:pointer;}',
    '#app-page{display:none;flex-direction:column;height:100%;position:relative;}',
    '.header{display:flex;justify-content:space-between;background:#fff;padding:10px 20px;border-radius:8px;margin-bottom:15px;font-weight:bold;align-items:center;}',
    '.progress-bar{flex-grow:1;margin:0 20px;height:10px;background:#e0e0e0;border-radius:5px;overflow:hidden}.progress-fill{height:100%;background:#4a86e8;width:0%;transition:width .3s}',
    '.main-area{display:flex;flex-grow:1;gap:15px;overflow:hidden}',
    '.left-panel{flex:1.2;background:#fff;border-radius:8px;display:flex;flex-direction:column;align-items:center;justify-content:center;overflow:hidden;padding:10px;}',
    '.left-panel img{max-width:100%;max-height:100%;object-fit:contain; cursor:zoom-in;}',
    '.right-panel{flex:1;display:flex;flex-direction:column;gap:10px;overflow-y:auto; padding-right:5px;}',
    '.card{background:#fff;padding:15px;border-radius:8px;}.title{font-size:14px;color:#666;border-bottom:1px solid #eee;padding-bottom:5px;margin-bottom:10px}.data-row{display:flex;justify-content:space-between;margin-bottom:5px;font-size:15px}.data-value{font-weight:bold;color:#333}',
    '.btn-group{display:flex;flex-direction:column;gap:10px;margin-top:auto;padding-bottom:20px}',
    'button{padding:12px;font-size:14px;border:none;border-radius:6px;cursor:pointer;font-weight:bold;color:#fff;}',
    '.btn-link{background:#34a853}.btn-new{background:#4a86e8}.btn-skip{background:#5f6368}.btn-update{background:#fbbc04;color:#000}.btn-exit{background:#5f6368}.btn-trash{background:#ea4335;}',
    '.btn-force-done{background:#8e24aa;} .btn-force-tag{background:#f57f17; color:#000;}', 
    '#loading{display:flex;flex-direction:column;align-items:center;justify-content:center;height:100%;color:#666;}',
    '.spinner{border:4px solid #f3f3f3;border-top:4px solid #4a86e8;border-radius:50%;width:40px;height:40px;animation:spin 1s linear infinite;margin-bottom:15px}@keyframes spin{100%{transform:rotate(360deg)}}',
    '.lightbox { display:none; position:fixed; z-index:9999; top:0; left:0; width:100%; height:100%; background:rgba(0,0,0,0.85); align-items:center; justify-content:center; cursor:zoom-out; } .lightbox img { max-width:95vw; max-height:95vh; object-fit:contain; border-radius:8px; } .lightbox-close { position:absolute; top:20px; right:30px; color:white; font-size:40px; cursor:pointer; font-weight:bold;}',
    '</style></head><body>',
    '<div id="lightbox" class="lightbox" onclick="this.style.display=\'none\'"><span class="lightbox-close">&times;</span><img id="lightbox-img" src=""></div>',
    '<div id="config-page">',
    '<h2 style="color:#1a73e8; margin-top:0;">🖥️ 常规流水线启动</h2>',
    '<div id="config-folder-list"></div><div style="margin-top:20px;"><input type="text" id="custom-folder-url" class="custom-input" placeholder="粘贴其他文件夹ID..."></div>',
    '<button id="btnLaunchMatch" class="btn-start-scan">🚀 加载照片进入对账</button>',
    '</div>',
    '<div id="app-page">',
    '',
    '<div id="custom-confirm" style="display:none; position:absolute; top:0; left:0; width:100%; height:100%; background:rgba(0,0,0,0.65); z-index:1000; align-items:center; justify-content:center; border-radius:8px;">',
    '<div style="background:#fff; padding:25px; border-radius:10px; width:320px; text-align:center; box-shadow:0 4px 15px rgba(0,0,0,0.3);">',
    '<h3 style="margin-top:0; color:#1a73e8;">⚙️ 绑定前确认</h3>',
    '<p style="font-size:14px; color:#333; margin-bottom:20px; line-height:1.5;">强制绑定所选行，是否将原金额修改为发票价：<br><b style="color:#d93025;font-size:24px;">¥<span id="confirm-amt"></span></b>？</p>',
    '<div style="display:flex; flex-direction:column; gap:10px;">',
    '<button onclick="confirmAction(true)" style="background:#ea4335; color:#fff; padding:12px; border:none; border-radius:6px; font-weight:bold; cursor:pointer; font-size:15px;">💰 听发票价 (强行覆盖)</button>',
    '<button onclick="confirmAction(false)" style="background:#34a853; color:#fff; padding:12px; border:none; border-radius:6px; font-weight:bold; cursor:pointer; font-size:15px;">📝 听表格价 (保留原金额)</button>',
    '<button onclick="closeConfirm()" style="background:#f1f3f4; color:#5f6368; padding:12px; border:none; border-radius:6px; font-weight:bold; cursor:pointer; margin-top:5px;">❌ 取消，什么都不做</button>',
    '</div></div></div>',
    '',
    '<div class="header"><span id="status-text">进行中...</span><div class="progress-bar"><div class="progress-fill" id="progress-fill"></div></div><span id="counter">0 / 0</span></div>',
    '<div class="main-area" id="main-area">',
    '<div class="left-panel" id="img-container"><div id="loading"><div class="spinner"></div></div></div>',
    '<div class="right-panel">',
    '<div class="card" id="error-card" style="display:none;"><div class="title" style="color:#d93025">提示</div><div id="error-msg"></div></div>',
    '<div class="card" id="ai-card" style="display:none;"><div class="title">📸 AI 提取</div><div class="data-row"><span>日期：</span><span class="data-value" id="ai-date"></span></div><div class="data-row"><span>金额：</span><span class="data-value" style="color:#d93025; font-size:18px;" id="ai-amount"></span></div><div class="data-row"><span>摘要：</span><span class="data-value" id="ai-cat"></span></div></div>',
    '<div class="card" id="match-card" style="display:none;"><div class="title">🎯 筛选匹配</div><div id="match-content"></div></div>',
    '<div class="btn-group" id="btn-group-normal" style="display:none;">',
    '<button class="btn-link" id="btn-link" style="display:none;">✅ 完美匹配</button>',
    '<button class="btn-update" id="btn-link-update" style="display:none;">🔄 听发票价</button>',
    '<button class="btn-link" id="btn-link-keep" style="display:none;border:1px solid #34a853;background:#fff;color:#34a853">✅ 听表格价</button>',
    '<button class="btn-new" id="btn-new">➕ 存为新账</button>',
    '<button class="btn-force-done" id="btn-force-done">🔗 完结强绑选中 (入库)</button>',
    '<button class="btn-force-tag" id="btn-force-tag">🔖 记号强绑选中 (待用)</button>',
    '<button class="btn-trash" id="btn-trash-new" style="margin-top:10px;">🗑️ 彻底删除废图</button>',
    '<button class="btn-skip" id="btn-skip">⏭️ 暂不处理跳过</button>',
    '<button class="btn-exit" id="btn-exit">🛑 退出对账</button>',
    '<button class="btn-update" id="btn-retry" style="display:none; background:#1a73e8; color:white; margin-top:10px;">🔄 重试</button>',
    '</div></div></div></div>',
    '<script>',
    'var initFolders = <?!= foldersStr ?>; var cList = document.getElementById("config-folder-list"); var chkHtml = "";',
    'for(var i=0; i<initFolders.length; i++) { var isChecked = initFolders[i].name.indexOf("2025-12") !== -1 ? "checked" : ""; chkHtml += "<label class=\'folder-item\'><input type=\'checkbox\' class=\'match-cb\' value=\'"+initFolders[i].id+"\' " + isChecked + "> <b>📁 " + initFolders[i].name + "</b></label>"; }',
    'cList.innerHTML = chkHtml;',
    'var files=[], currentIndex=0, total=0, aiD=null, tr=-1;',
    'function openLightbox(url) { document.getElementById("lightbox-img").src = url; document.getElementById("lightbox").style.display = "flex"; }',
    'document.getElementById("btnLaunchMatch").addEventListener("click", function(){',
    '  var btn = document.getElementById("btnLaunchMatch"); btn.innerText = "⏳ 提取中..."; btn.disabled = true;',
    '  var cbs = document.querySelectorAll(".match-cb:checked"); var selectedIds = [];',
    '  for(var i=0; i<cbs.length; i++) selectedIds.push(cbs[i].value);',
    '  var customVal = document.getElementById("custom-folder-url").value.trim(); if(customVal) { var match = customVal.match(/[-\\w]{25,}/); if(match) selectedIds.push(match[0]); }',
    '  if(selectedIds.length === 0) { alert("❌ 请勾选"); btn.innerText = "🚀 开始"; btn.disabled = false; return; }',
    '  document.getElementById("config-page").style.display = "none"; document.getElementById("app-page").style.display = "flex"; document.getElementById("img-container").innerHTML = "<div id=\'loading\'><div class=\'spinner\'></div></div>";',
    '  google.script.run.withSuccessHandler(function(f) { files = f; total = files.length; if(total === 0) { document.getElementById("app-page").innerHTML = "<h2 style=\'text-align:center;padding:50px;color:#34a853\'>✅ 仓库干净！</h2>"; return; } load(); }).withFailureHandler(function(err) { document.getElementById("img-container").innerHTML = "读取失败："+err; }).getFilesFromMultipleFolders(selectedIds);',
    '});',
    'function load(){',
    '  if(currentIndex>=total){document.getElementById("app-page").innerHTML="<h2 style=\'text-align:center;padding:50px;color:#34a853\'>🎉 对账完成！</h2>"; return;}',
    '  document.getElementById("counter").innerText=(currentIndex+1)+" / "+total; document.getElementById("progress-fill").style.width=(currentIndex/total*100)+"%";',
    '  document.getElementById("img-container").innerHTML = "<div id=\'loading\'><div class=\'spinner\'></div></div>";',
    '  ["ai-card","match-card","btn-group-normal","btn-retry","error-card"].forEach(id=>document.getElementById(id).style.display="none");',
    '  google.script.run.withSuccessHandler(res=>{',
    '      if(!res.success){ showErr(res); return; } aiD=res.aiData; tr=res.matchRow; ',
    '      document.getElementById("img-container").innerHTML="<img src=\'"+res.imgSrc+"\' onclick=\'openLightbox(this.src)\' style=\'cursor:zoom-in;\'>";',
    '      document.getElementById("ai-date").innerText=aiD.date||"无"; document.getElementById("ai-amount").innerText="¥"+(aiD.amount||0); document.getElementById("ai-cat").innerText=aiD.summary; ',
    '      document.getElementById("ai-card").style.display="block"; document.getElementById("match-card").style.display="block"; document.getElementById("btn-group-normal").style.display="flex";',
    '      ["btn-link","btn-link-update","btn-link-keep"].forEach(id=>document.getElementById(id).style.display="none");',
    '      if(res.hasConflict) {',
    '         document.getElementById("match-content").innerHTML="推荐: 第 <b style=\'color:#ea4335;\'>"+tr+"</b> 行 | ¥"+res.matchInfo.amount+"<div style=\'color:#ea4335;font-weight:bold;margin-top:10px;background:#fce8e6;padding:8px;\'>⚠️ 该行已占用，建议跳过留给法庭。</div>";',
    '      } else {',
    '         if(tr !== -1){',
    '           google.script.run.highlightTargetRow(tr); var diff = Math.abs((parseFloat(aiD.amount)||0) - parseFloat(res.matchInfo.amount));',
    '           document.getElementById("match-content").innerHTML="推荐: 第 <b style=\'color:#1a73e8;\'>"+tr+"</b> 行 | ¥"+res.matchInfo.amount;',
    '           if(diff > 0.01) { document.getElementById("btn-link-update").style.display="block"; document.getElementById("btn-link-keep").style.display="block"; } else document.getElementById("btn-link").style.display="block";',
    '         } else { document.getElementById("match-content").innerHTML="无未关联项"; }',
    '      }',
    '    }).withFailureHandler(err=>{ showErr({detail: "超时: " + err.message}); }).analyzeSingleImage(files[currentIndex].id);',
    '}',
    'function showErr(res){ document.getElementById("img-container").innerHTML="<h1 style=\'text-align:center\'>⚠️失败</h1>"; document.getElementById("error-msg").innerText=res.detail; document.getElementById("error-card").style.display="block"; document.getElementById("btn-group-normal").style.display="flex"; document.getElementById("btn-retry").style.display="block"; }',
    'document.getElementById("btn-link-update").onclick=()=>submitAction("LINK_UPDATE"); document.getElementById("btn-link-keep").onclick=()=>submitAction("LINK"); document.getElementById("btn-link").onclick=()=>submitAction("LINK"); document.getElementById("btn-new").onclick=()=>submitAction("NEW");',
    'document.getElementById("btn-skip").onclick=()=>{currentIndex++; load();}; document.getElementById("btn-trash-new").onclick=()=>{ if(!confirm("彻底删？"))return; document.getElementById("img-container").innerHTML="删除中..."; google.script.run.withSuccessHandler(()=>{currentIndex++;load();}).trashDuplicateFiles([files[currentIndex].id]); }; document.getElementById("btn-retry").onclick=load; document.getElementById("btn-exit").onclick=()=>google.script.host.close();',
    'function submitAction(a){ document.getElementById("img-container").innerHTML="<div class=\'spinner\'></div>"; google.script.run.withSuccessHandler(()=>{currentIndex++;load();}).executeAction(a, files[currentIndex].id, aiD, tr); }',
    '// 🌟 全新三选一面板逻辑',
    'var pendingForceAction = "";',
    'document.getElementById("btn-force-done").onclick=()=>{ triggerCustomConfirm("DONE"); };',
    'document.getElementById("btn-force-tag").onclick=()=>{ triggerCustomConfirm("TAG"); };',
    'function triggerCustomConfirm(actionType){ ',
    '  if(aiD && aiD.amount){',
    '    pendingForceAction = actionType;',
    '    document.getElementById("confirm-amt").innerText = aiD.amount;',
    '    document.getElementById("custom-confirm").style.display = "flex";',
    '  } else {',
    '    executeForceBind(actionType, false);',
    '  }',
    '}',
    'function closeConfirm(){ document.getElementById("custom-confirm").style.display = "none"; }',
    'function confirmAction(doUpdate){ closeConfirm(); executeForceBind(pendingForceAction, doUpdate); }',
    'function executeForceBind(actionType, doUpdate){ ',
    '  document.getElementById("img-container").innerHTML="<div class=\'spinner\'>正在挂载凭证...</div>"; ',
    '  google.script.run.withSuccessHandler((res)=>{ ',
    '    if(!res.success){ alert(res.error); load(); return; } ',
    '    currentIndex++; load(); ',
    '  }).withFailureHandler((err)=>{ ',
    '    alert("网络异常: " + err.message); load(); ',
    '  }).executeForceBindSelected(actionType, files[currentIndex].id, aiD, doUpdate); ',
    '}',
    '</script></body></html>'
  ].join('\n');
}

function getCourtHtmlContent() {
  return [
    '<!DOCTYPE html><html><head><style>',
    'body{font-family:sans-serif;background:#202124;margin:0;padding:15px;height:100vh;box-sizing:border-box;overflow:hidden;color:#e8eaed;}',
    '#config-page{background:#303134;border-radius:8px;padding:30px;max-width:600px;margin:20px auto;border:1px solid #5f6368;}',
    '.folder-item{display:flex;align-items:center;padding:12px;background:#3c4043;margin-bottom:8px;border-radius:6px;cursor:pointer;border:1px solid #5f6368;}',
    '.folder-item input{width:20px;height:20px;margin-right:15px;cursor:pointer;}',
    '.btn-start-scan{background:#d93025;color:white;padding:15px;border:none;border-radius:6px;width:100%;font-size:16px;font-weight:bold;margin-top:20px;cursor:pointer;}',
    '#app-page{display:none;flex-direction:column;height:100%;}',
    '.header{display:flex;justify-content:space-between;background:#303134;padding:10px 20px;border-radius:8px;margin-bottom:15px;font-weight:bold;align-items:center;border:1px solid #5f6368;}',
    '.progress-bar{flex-grow:1;margin:0 20px;height:10px;background:#5f6368;border-radius:5px;overflow:hidden}.progress-fill{height:100%;background:#d93025;width:0%;transition:width .3s}',
    '.main-area{display:flex;flex-grow:1;gap:15px;overflow:hidden}',
    '.left-panel{flex:2;background:#303134;border-radius:8px;display:flex;flex-direction:column;align-items:center;justify-content:center;overflow:hidden;padding:10px;border:1px solid #5f6368;}',
    '.right-panel{flex:1;display:flex;flex-direction:column;gap:10px;overflow-y:auto; padding-right:5px;}',
    '.btn-group{display:flex;flex-direction:column;gap:10px;margin-top:auto;}',
    '.conflict-screen { display:flex; gap:10px; width:100%; height:100%; }',
    '.half-screen { flex:1; display:flex; flex-direction:column; border:1px solid #5f6368; border-radius:6px; padding:10px; background:#202124; height:100%; box-sizing:border-box;}',
    '.half-title { font-weight:bold; font-size:15px; margin-bottom:10px; text-align:center; padding:8px; border-radius:4px; width:100%; box-sizing:border-box;}',
    '.conflict-btn { width:100%; text-align:left; padding:14px 15px; margin-bottom:8px; font-size:14px; border-radius:6px; border:none; cursor:pointer; font-weight:bold;}',
    '#loading{display:flex;flex-direction:column;align-items:center;justify-content:center;height:100%;color:#9aa0a6;}',
    '.spinner{border:4px solid #5f6368;border-top:4px solid #d93025;border-radius:50%;width:40px;height:40px;animation:spin 1s linear infinite;margin-bottom:15px}@keyframes spin{100%{transform:rotate(360deg)}}',
    '.half-screen::-webkit-scrollbar{width:6px;} .half-screen::-webkit-scrollbar-thumb{background:#5f6368;border-radius:4px;}',
    '.lightbox { display:none; position:fixed; z-index:9999; top:0; left:0; width:100%; height:100%; background:rgba(0,0,0,0.9); alignitems:center; justify-content:center; cursor:zoom-out; } .lightbox img { max-width:95vw; max-height:95vh; object-fit:contain; border-radius:8px;} .lightbox-close { position:absolute; top:20px; right:30px; color:white; font-size:40px; cursor:pointer; font-weight:bold;}',
    '</style></head><body>',
    '<div id="lightbox" class="lightbox" onclick="this.style.display=\'none\'"><span class="lightbox-close">&times;</span><img id="lightbox-img" src=""></div>',
    '<div id="config-page">',
    '<h2 style="color:#f28b82; margin-top:0;">⚖️ 启动独立双屏法庭</h2>',
    '<p style="color:#9aa0a6;font-size:14px;">自动跳过正常照片，仅审理冲突账单。</p><div id="config-folder-list"></div><button id="btnLaunchMatch" class="btn-start-scan">🚀 升堂断案</button>',
    '</div>',
    '<div id="app-page">',
    '<div class="header"><span id="status-text" style="color:#f28b82">排查中...</span><div class="progress-bar"><div class="progress-fill" id="progress-fill"></div></div><span id="counter">0 / 0</span></div>',
    '<div class="main-area" id="main-area">',
    '<div class="left-panel" id="img-container"><div id="loading"><div class="spinner"></div></div></div>',
    '<div class="right-panel">',
    '<div id="conflict-panels" style="display:none; height:100%; flex-direction:column;">',
    '<div style="border:2px solid #ea4335; background:#3c4043; padding:12px; border-radius:8px; margin-bottom:15px;">',
    '<h3 style="color:#f28b82; margin-top:0; font-size:16px;">⚠️ 冲突！</h3><p style="font-size:13px; margin-bottom:0;">撞车第 <b id="c-row" style="color:#8ab4f8;font-size:16px;"></b> 行！</p>',
    '</div>',
    '<div class="btn-group" style="flex-grow:1; justify-content:center;">',
    '<button class="conflict-btn" style="background:#d32f2f; color:#fff;" onclick="submitConflict(\'DELETE_NEW\')">🗑️ 删左侧新图</button>',
    '<button class="conflict-btn" style="background:#ea4335; color:#fff;" onclick="submitConflict(\'REPLACE_TRASH\')">💥 替换并【删右侧老图】</button>',
    '<button class="conflict-btn" style="background:#fbbc04; color:#000;" onclick="submitConflict(\'REPLACE_RETURN\')">🔙 替换，右侧老图退回</button>',
    '<button class="conflict-btn" style="background:#34a853; color:#fff;" onclick="submitConflict(\'APPEND\')">📎 同笔合并追加</button>',
    '<button class="conflict-btn" style="background:#1a73e8; color:#fff;" onclick="submitConflict(\'NEW_ROW\')">➕ 新建一行入账</button>',
    '<button class="conflict-btn" style="background:#5f6368; color:#fff;" onclick="submitConflict(\'SKIP\')">⏭️ 跳过此案</button>',
    '<button class="conflict-btn" style="background:#202124; color:#9aa0a6; border:1px solid #5f6368; text-align:center; margin-top:20px;" onclick="google.script.host.close()">🛑 休庭</button>',
    '</div></div></div></div></div>',
    '<script>',
    'var initFolders = <?!= foldersStr ?>; var cList = document.getElementById("config-folder-list"); var chkHtml = "";',
    'for(var i=0; i<initFolders.length; i++) { var isChecked = initFolders[i].name.indexOf("2025-12") !== -1 ? "checked" : ""; chkHtml += "<label class=\'folder-item\'><input type=\'checkbox\' class=\'match-cb\' value=\'"+initFolders[i].id+"\' " + isChecked + "> <b>📁 " + initFolders[i].name + "</b></label>"; }',
    'cList.innerHTML = chkHtml;',
    'var files=[], currentIndex=0, total=0, aiD=null, tr=-1, conflictIds=[];',
    'function openLightbox(url) { document.getElementById("lightbox-img").src = url; document.getElementById("lightbox").style.display = "flex"; }',
    'document.getElementById("btnLaunchMatch").addEventListener("click", function(){',
    '  var btn = document.getElementById("btnLaunchMatch"); btn.innerText = "⏳..."; btn.disabled = true;',
    '  var cbs = document.querySelectorAll(".match-cb:checked"); var selectedIds = [];',
    '  for(var i=0; i<cbs.length; i++) selectedIds.push(cbs[i].value);',
    '  if(selectedIds.length === 0) { alert("❌ 请勾选！"); btn.innerText = "🚀"; btn.disabled = false; return; }',
    '  document.getElementById("config-page").style.display = "none"; document.getElementById("app-page").style.display = "flex"; document.getElementById("img-container").innerHTML = "<div id=\'loading\'><div class=\'spinner\'></div></div>";',
    '  google.script.run.withSuccessHandler(function(f) { files = f; total = files.length; if(total === 0) { document.getElementById("app-page").innerHTML = "<h2 style=\'text-align:center;padding:50px;color:#81c995\'>✅ 干净！</h2>"; return; } load(); }).withFailureHandler(function(err) { document.getElementById("img-container").innerHTML = "失败："+err; }).getFilesFromMultipleFolders(selectedIds);',
    '});',
    'function load(){',
    '  if(currentIndex>=total){document.getElementById("app-page").innerHTML="<h2 style=\'text-align:center;padding:50px;color:#81c995\'>🎉 审理完毕！正常照片已自动保留。</h2>"; return;}',
    '  document.getElementById("counter").innerText=(currentIndex+1)+" / "+total; document.getElementById("progress-fill").style.width=(currentIndex/total*100)+"%";',
    '  document.getElementById("img-container").innerHTML="<div id=\'loading\'><div class=\'spinner\'></div>过滤中...</div>"; document.getElementById("conflict-panels").style.display = "none";',
    '  google.script.run.withSuccessHandler(res=>{',
    '      if(!res.success){ currentIndex++; load(); return; } aiD=res.aiData; tr=res.matchRow; conflictIds = res.conflictIds || [];',
    '      if(!res.hasConflict) { currentIndex++; load(); return; }',
    '      var oldImagesHtml = "";',
    '      if (conflictIds && conflictIds.length > 0) {',
    '          var THUMB_BASE = "' + DUKA_CONFIG.RESOURCES.URLS.GOOGLE_DRIVE.THUMBNAIL + '";',
    '          for(var k=0; k<conflictIds.length; k++) {',
    '              var bigUrl = THUMB_BASE + "?id=" + conflictIds[k] + "&sz=w1600";',
    '              oldImagesHtml += "<img src=\\"" + THUMB_BASE + "?id=" + conflictIds[k] + "&sz=w800\\" onclick=\\"openLightbox(\'" + bigUrl + "\')\\" style=\\"max-width:100%; border-radius:4px; margin-bottom:15px; border:1px solid #5f6368; cursor:zoom-in;\\" />";',
    '          }',
    '      } else { oldImagesHtml = "<div style=\'color:#9aa0a6;\'>无原图ID</div>"; }',
    '      document.getElementById("img-container").innerHTML = ',
    '         "<div class=\'conflict-screen\'><div class=\'half-screen\' style=\'overflow-y:auto;\'><div class=\'half-title\' style=\'background:#8ab4f8;color:#202124;\'>🆕 新截图 ¥"+aiD.amount+"</div><img src=\'"+res.imgSrc+"\' onclick=\'openLightbox(this.src)\' style=\'max-width:100%;cursor:zoom-in;\' /></div><div class=\'half-screen\' style=\'overflow-y:auto;background:#3c4043;\'><div class=\'half-title\' style=\'background:#f28b82;color:#202124;\'>🔒 老凭证 ("+conflictIds.length+"张)</div>"+oldImagesHtml+"</div></div>";',
    '      document.getElementById("c-row").innerText = tr; document.getElementById("conflict-panels").style.display = "flex";',
    '    }).withFailureHandler(err=>{ currentIndex++; load(); }).analyzeSingleImage(files[currentIndex].id);',
    '}',
    'window.submitConflict = function(actionCode) {',
    '   if (actionCode === "SKIP") { currentIndex++; load(); return; }',
    '   if(!confirm("确定执行？")) return;',
    '   document.getElementById("img-container").innerHTML="<div class=\'spinner\'></div>";',
    '   google.script.run.withSuccessHandler(function(){ currentIndex++; load(); }).executeConflictAction(actionCode, files[currentIndex].id, aiD, tr, conflictIds);',
    '};',
    '</script></body></html>'
  ].join('\n');
}

function getUnbindHtmlContent() {
  return [
    '<!DOCTYPE html><html><head><style>',
    'body{font-family:sans-serif;background:#f0f2f5;padding:20px;overflow-x:hidden;} .title{color:#d93025;text-align:center;margin-top:0;} .grid{display:grid;grid-template-columns:repeat(auto-fill,minmax(220px,1fr));gap:15px;margin-top:20px;max-height:450px;overflow-y:auto;padding:10px;} .img-box{background:#fff;padding:15px;border-radius:8px;text-align:center;position:relative;border:3px solid transparent;box-shadow:0 2px 5px rgba(0,0,0,0.1);} .img-box.selected{border-color:#ea4335;background:#fce8e6;} .img-box img{max-width:100%;height:180px;object-fit:contain;cursor:zoom-in;border-radius:4px;} .check-overlay{position:absolute;top:20px;left:20px;transform:scale(2);cursor:pointer;z-index:10;} .btn-bar{margin-top:30px;display:flex;justify-content:center;gap:20px;} button{padding:12px 25px;border:none;border-radius:6px;font-size:15px;font-weight:bold;cursor:pointer;} .btn-del{background:#ea4335;color:#fff;} .btn-all{background:#d32f2f;color:#fff;} .btn-cancel{background:#5f6368;color:#fff;} .lightbox{display:none;position:fixed;z-index:9999;top:0;left:0;width:100%;height:100%;background:rgba(0,0,0,0.85);align-items:center;justify-content:center;cursor:zoom-out;} .lightbox img{max-width:95vw;max-height:95vh;object-fit:contain;border-radius:8px;} .lightbox-close{position:absolute;top:20px;right:30px;color:white;font-size:40px;cursor:pointer;font-weight:bold;}',
    '</style></head><body>',
    '<div id="lightbox" class="lightbox" onclick="this.style.display=\'none\'"><span class="lightbox-close">&times;</span><img id="lightbox-img" src=""></div>',
    '<div id="main"><h2 class="title">🧹 精准解绑法庭 (第 <?= rowNum ?> 行)</h2><div class="grid" id="grid"></div><div class="btn-bar"><button class="btn-cancel" onclick="google.script.host.close()">取消关闭</button><button class="btn-del" onclick="submitUnbind(false)">✂️ 仅退回打勾</button><button class="btn-all" onclick="submitUnbind(true)">💥 全部退回</button></div></div>',
    '<script>',
    '  var fileIds = <?!= fileIdsStr ?>; var rowNum = <?= rowNum ?>; var html = "";',
    '  var THUMB_BASE = "' + DUKA_CONFIG.RESOURCES.URLS.GOOGLE_DRIVE.THUMBNAIL + '";',
    '  var THUMB_SMALL = "' + DUKA_CONFIG.SETTINGS.LIMITS.THUMB_SIZE_UPLOAD_RETURN + '";',
    '  fileIds.forEach(function(id, idx) { html += "<div class=\'img-box\' id=\'box-"+idx+"\'><input type=\'checkbox\' class=\'cb check-overlay\' value=\'"+id+"\' id=\'cb-"+idx+"\' onchange=\'document.getElementById(\\"box-"+idx+"\\").classList.toggle(\\"selected\\")\'><img src=\\"" + THUMB_BASE + "?id="+id+"&sz=" + THUMB_SMALL + "\\" onclick=\'document.getElementById(\\"lightbox-img\\").src=\\"" + THUMB_BASE + "?id="+id+"&sz=w1600\\";document.getElementById(\\"lightbox\\").style.display=\\"flex\\"\' /><div>🖼️ 凭证 "+(idx+1)+"</div></div>"; });',
    '  document.getElementById("grid").innerHTML = html;',
    '  function submitUnbind(isAll) {',
    '    var ids = []; if (isAll) { if(!confirm("全退回？")) return; ids = fileIds; } else { document.querySelectorAll(".cb:checked").forEach(c=>ids.push(c.value)); if(ids.length===0) return; if(!confirm("确定剔除？")) return; }',
    '    document.getElementById("main").innerHTML = "<h2 style=\'text-align:center;margin-top:50px;\'>⏳ 执行中...</h2>";',
    '    google.script.run.withSuccessHandler(function(){google.script.host.close();}).executePartialUnbind(rowNum, ids);',
    '  }',
    '</script></body></html>'
  ].join('\n');
}

