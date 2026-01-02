/**
 * シートを開いた時にカスタムメニューを追加します。
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("freee")
    .addItem("🔐 認証", "showAuthDialog")
    .addItem("🔄 認証リセット", "resetAuth")
    .addSeparator()
    .addItem("📄 新規クライアント作成", "createNewClient")
    .addItem("🔄 選択したクライアントの試算表を再取得", "refreshSelectedClient")
    .addSeparator()
    .addItem("試算表取得（このシート）", "getTrialBalanceAndPL")
    .addToUi();
}

