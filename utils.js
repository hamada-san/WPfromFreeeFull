/**
 * 更新日時のタイムスタンプを生成
 */
function getTimestamp() {
  const now = new Date();
  const year = now.getFullYear();
  const month = String(now.getMonth() + 1).padStart(2, '0');
  const day = String(now.getDate()).padStart(2, '0');
  const hours = String(now.getHours()).padStart(2, '0');
  const minutes = String(now.getMinutes()).padStart(2, '0');
  return `${year}/${month}/${day} ${hours}:${minutes}更新`;
}

/**
 * 保存先フォルダIDを取得
 */
function getSavedFolderId() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const configSheet = ss.getSheetByName("クライアント一覧");
  if (configSheet) {
    return configSheet.getRange("K1").getValue() || "";
  }
  return "";
}

/**
 * 保存先フォルダ名を取得
 */
function getSavedFolderName() {
  const folderId = getSavedFolderId();
  if (!folderId) return "";
  try {
    const folder = DriveApp.getFolderById(folderId);
    return folder.getName();
  } catch (e) {
    return "";
  }
}

/**
 * 保存先フォルダIDを保存
 */
function saveFolderId(folderId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let configSheet = ss.getSheetByName("クライアント一覧");
  if (!configSheet) {
    configSheet = createConfigSheet(ss);
  }
  configSheet.getRange("J1").setValue("保存先フォルダID:");
  configSheet.getRange("K1").setValue(folderId);
}
