/**
 * クライアント一覧シートを作成
 */
function createConfigSheet(ss) {
  const configSheet = ss.insertSheet(CONFIG.SHEETS.CLIENT_LIST);
  const headers = ["事業所名", "事業所ID", "期", "期末日", "スプレッドシートID", "URL", "フォルダURL", "ステータス", "最終更新日時", "作業者"];
  configSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  configSheet.getRange(1, 1, 1, headers.length).setFontWeight("bold").setBackground("#4285f4").setFontColor("#ffffff");
  configSheet.setColumnWidth(1, 200);
  configSheet.setColumnWidth(2, 100);
  configSheet.setColumnWidth(3, 120);
  configSheet.setColumnWidth(4, 100);
  configSheet.setColumnWidth(5, 320);
  configSheet.setColumnWidth(6, 80);
  configSheet.setColumnWidth(7, 120);
  configSheet.setColumnWidth(8, 80);
  configSheet.setColumnWidth(9, 150);
  configSheet.setColumnWidth(10, 200);
  return configSheet;
}

/**
 * クライアント一覧に追加
 */
function addToClientList(configSheet, companyDetails, ssId, ssUrl, folderUrl, status, workerEmail) {
  const lastRow = configSheet.getLastRow();
  const timestamp = getTimestamp();
  configSheet.getRange(lastRow + 1, 1, 1, 10).setValues([[
    companyDetails.companyName,
    companyDetails.companyId,
    companyDetails.periodLabel,
    new Date(companyDetails.endDate),
    ssId,
    "",
    "",
    status,
    timestamp,
    workerEmail
  ]]);
  configSheet.getRange(lastRow + 1, 6).setFormula(`=HYPERLINK("${ssUrl}", "開く")`);
  if (folderUrl) {
    configSheet.getRange(lastRow + 1, 7).setFormula(`=HYPERLINK("${folderUrl}", "開く")`);
  }
}
