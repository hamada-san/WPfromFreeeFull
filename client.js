/**
 * 新規クライアント作成
 */
function createNewClient() {
  const ui = SpreadsheetApp.getUi();
  
  if (!checkAuth()) {
    showAuthDialog();
    return;
  }
  
  let companies;
  try {
    companies = fetchCompaniesFromFreee();
  } catch (e) {
    ui.alert("エラー", "事業所リストの取得に失敗しました: " + e.message, ui.ButtonSet.OK);
    return;
  }
  
  if (companies.length === 0) {
    ui.alert("エラー", "freeeに登録されている事業所がありません。", ui.ButtonSet.OK);
    return;
  }
  
  const savedFolderId = getSavedFolderId();
  const savedFolderName = getSavedFolderName();
  const htmlContent = createCompanySelectHtml(companies, savedFolderId, savedFolderName);
  const html = HtmlService.createHtmlOutput(htmlContent)
    .setWidth(500)
    .setHeight(420);
  ui.showModalDialog(html, '事業所を選択してください');
}

/**
 * 事業所選択用HTMLを生成
 */
function createCompanySelectHtml(companies, savedFolderId, savedFolderName) {
  let options = companies.map(c => 
    `<option value="${c.id}">${c.display_name || c.name}</option>`
  ).join('');
  
  const showFolderInfo = savedFolderName ? 'block' : 'none';
  
  return `<!DOCTYPE html>
    <html>
    <head>
      <base target="_top">
      <style>
        * { box-sizing: border-box; }
        body { font-family: 'Segoe UI', Arial, sans-serif; padding: 24px; margin: 0; color: #333; }
        .form-group { margin-bottom: 20px; }
        label { display: block; margin-bottom: 8px; font-weight: 600; font-size: 14px; }
        select, input[type="text"] { width: 100%; padding: 10px 12px; font-size: 14px; border: 1px solid #ddd; border-radius: 6px; background: #fff; }
        select:focus, input[type="text"]:focus { outline: none; border-color: #4285f4; }
        .folder-section { background: #f8f9fa; padding: 16px; border-radius: 8px; border: 1px solid #e0e0e0; }
        .folder-info { font-size: 13px; color: #1a73e8; margin-bottom: 12px; padding: 8px; background: #e8f0fe; border-radius: 4px; display: ${showFolderInfo}; }
        .note { font-size: 12px; color: #666; margin-top: 8px; }
        .button-row { display: flex; gap: 12px; margin-top: 24px; }
        .submit-btn { background-color: #4285f4; color: white; padding: 12px 24px; border: none; cursor: pointer; font-size: 14px; font-weight: 600; border-radius: 6px; flex: 1; }
        .submit-btn:hover { background-color: #3367d6; }
        .submit-btn:disabled { background-color: #ccc; cursor: not-allowed; }
        .cancel-btn { background-color: #fff; color: #666; padding: 12px 24px; border: 1px solid #ddd; cursor: pointer; font-size: 14px; border-radius: 6px; }
        .cancel-btn:hover { background-color: #f5f5f5; }
        .loading { display: none; color: #4285f4; margin-top: 16px; font-size: 14px; }
        .file-note { background: #e8f0fe; padding: 12px; border-radius: 6px; font-size: 13px; color: #1a73e8; margin-top: 20px; }
      </style>
    </head>
    <body>
      <div class="form-group">
        <label>事業所を選択</label>
        <select id="companyId">${options}</select>
      </div>
      <div class="form-group">
        <label>保存先フォルダID</label>
        <div class="folder-section">
          <div class="folder-info" id="folderInfo">📁 現在の保存先: ${savedFolderName}</div>
          <input type="text" id="folderId" value="${savedFolderId}">
          <div class="note">※ 空欄の場合はマイドライブ直下に作成されます</div>
          <div class="note">※ フォルダIDはGoogleドライブのフォルダURLの末尾の文字列です</div>
        </div>
      </div>
      <div class="file-note">📄 ファイル名: WP_事業所名_○○年○月期</div>
      <div class="button-row">
        <button class="cancel-btn" onclick="google.script.host.close()">キャンセル</button>
        <button id="submitBtn" class="submit-btn" onclick="submitForm()">作成して試算表取得</button>
      </div>
      <div class="loading" id="loading">⏳ 処理中です。しばらくお待ちください...</div>
      <script>
        function submitForm() {
          const companyId = document.getElementById('companyId').value;
          const folderId = document.getElementById('folderId').value.trim();
          document.getElementById('submitBtn').disabled = true;
          document.getElementById('loading').style.display = 'block';
          google.script.run
            .withSuccessHandler(function(result) { google.script.host.close(); })
            .withFailureHandler(function(error) {
              alert('エラー: ' + error.message);
              document.getElementById('submitBtn').disabled = false;
              document.getElementById('loading').style.display = 'none';
            })
            .processNewClient(companyId, folderId);
        }
      </script>
    </body>
    </html>`;
}

/**
 * 新規クライアント処理
 */
function processNewClient(companyId, folderId) {
  const accessToken = getService().getAccessToken();
  const companyDetails = getCompanyDetails(companyId, accessToken);
  const workerEmail = Session.getActiveUser().getEmail();
  
  const mainSs = SpreadsheetApp.getActiveSpreadsheet();
  let configSheet = mainSs.getSheetByName("クライアント一覧");
  
  if (!configSheet) {
    configSheet = createConfigSheet(mainSs);
  }
  
  if (folderId) {
    saveFolderId(folderId);
  }
  
  const fileName = `WP_${companyDetails.companyName}_${companyDetails.periodLabel}`;
  const newSs = mainSs.copy(fileName);
  const newSsId = newSs.getId();
  const newSsUrl = newSs.getUrl();
  
  ["クライアント一覧", "事業所リスト"].forEach(sheetName => {
    const sheet = newSs.getSheetByName(sheetName);
    if (sheet) newSs.deleteSheet(sheet);
  });
  
  if (folderId) {
    try {
      DriveApp.getFileById(newSsId).moveTo(DriveApp.getFolderById(folderId));
    } catch (e) {
      Logger.log("フォルダ移動エラー: " + e.message);
    }
  }
  
  const docketSheet = newSs.getSheetByName("管理ドケット");
  if (docketSheet) {
    docketSheet.getRange("D8").setValue(companyDetails.companyName);
    docketSheet.getRange("D9").setValue(companyDetails.periodLabel);
    docketSheet.getRange("D11").setValue(new Date(companyDetails.startDate));
    docketSheet.getRange("D12").setValue(new Date(companyDetails.endDate));
    docketSheet.getRange("D14").setValue(companyDetails.companyId);
  }
  
  const taxSheet = newSs.getSheetByName("税務基本ステータス");
  if (taxSheet) {
    taxSheet.getRange("D16").setValue(companyDetails.address);
    taxSheet.getRange("D18").setValue(companyDetails.headName);
  }
  
  const bsSheet = newSs.getSheetByName("BS");
  const plSheet = newSs.getSheetByName("PL");
  if (bsSheet) bsSheet.getRange("B1").setValue(companyDetails.companyName);
  if (plSheet) plSheet.getRange("B1").setValue(companyDetails.companyName);
  
  try {
    getTrialBalanceAndPLCore(newSs, companyDetails.companyId, companyDetails.startDate, companyDetails.endDate, companyDetails.companyName);
  } catch (e) {
    addToClientList(configSheet, companyDetails, newSsId, newSsUrl, "❌ エラー", workerEmail);
    throw new Error("シートは作成しましたが、試算表取得でエラー: " + e.message);
  }
  
  addToClientList(configSheet, companyDetails, newSsId, newSsUrl, "✅ 完了", workerEmail);
  
  SpreadsheetApp.getActiveSpreadsheet().toast(`「${companyDetails.companyName}」のシートを作成しました。`, "完了", 5);
  
  return "完了";
}

/**
 * 選択したクライアントの試算表を再取得
 */
function refreshSelectedClient() {
  const ui = SpreadsheetApp.getUi();
  
  if (!checkAuth()) {
    showAuthDialog();
    return;
  }
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const configSheet = ss.getSheetByName("クライアント一覧");
  
  if (!configSheet) {
    ui.alert("エラー", "「クライアント一覧」シートがありません。", ui.ButtonSet.OK);
    return;
  }
  
  const activeRow = configSheet.getActiveCell().getRow();
  if (activeRow <= 1) {
    ui.alert("エラー", "クライアント一覧の2行目以降を選択してください。", ui.ButtonSet.OK);
    return;
  }
  
  const companyName = configSheet.getRange(activeRow, 1).getValue();
  const companyId = configSheet.getRange(activeRow, 2).getValue();
  const targetId = configSheet.getRange(activeRow, 5).getValue();
  
  if (!targetId) {
    ui.alert("エラー", "スプレッドシートIDがありません。", ui.ButtonSet.OK);
    return;
  }
  
  try {
    configSheet.getRange(activeRow, 7).setValue("処理中...");
    SpreadsheetApp.flush();
    
    const accessToken = getService().getAccessToken();
    const companyDetails = getCompanyDetails(companyId, accessToken);
    const workerEmail = Session.getActiveUser().getEmail();
    
    const targetSs = SpreadsheetApp.openById(targetId);
    const docketSheet = targetSs.getSheetByName("管理ドケット");
    if (docketSheet) {
      docketSheet.getRange("D11").setValue(new Date(companyDetails.startDate));
      docketSheet.getRange("D12").setValue(new Date(companyDetails.endDate));
    }
    
    configSheet.getRange(activeRow, 3).setValue(companyDetails.periodLabel);
    configSheet.getRange(activeRow, 4).setValue(new Date(companyDetails.endDate));
    
    getTrialBalanceAndPLCore(targetSs, companyDetails.companyId, companyDetails.startDate, companyDetails.endDate, companyDetails.companyName);
    
    configSheet.getRange(activeRow, 7).setValue("✅ 完了");
    configSheet.getRange(activeRow, 8).setValue(getTimestamp());
    configSheet.getRange(activeRow, 9).setValue(workerEmail);
    
    ui.alert("完了", `「${companyName}」の試算表を再取得しました。`, ui.ButtonSet.OK);
  } catch (e) {
    configSheet.getRange(activeRow, 7).setValue("❌ エラー");
    configSheet.getRange(activeRow, 8).setValue(getTimestamp());
    ui.alert("エラー", e.message, ui.ButtonSet.OK);
  }
}

