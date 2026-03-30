const Client_ID = '647162074272699';
const Client_Secret = 'wukiXVLCLLzoHF9qHGaNUmwWqOMzHqYfQsHEUn9e2InPrM-15LQaUz-L_25az9omshkssbytj4XvDiUwUAt-JQ'
/**
 * 認証サービスを定義・生成します。
 */
function getService() {
  return OAuth2.createService('freee')
    .setAuthorizationBaseUrl('https://accounts.secure.freee.co.jp/public_api/authorize')
    .setTokenUrl('https://accounts.secure.freee.co.jp/public_api/token')
    .setClientId(Client_ID)
    .setClientSecret(Client_Secret)
    .setCallbackFunction('authCallback')
    .setPropertyStore(PropertiesService.getUserProperties())
    .setScope('read write');
}

/**
 * 認証完了後に呼ばれるコールバック関数。
 */
function authCallback(request) {
  const service = getService();
  const isAuthorized = service.handleCallback(request);
  if (isAuthorized) {
    return HtmlService.createHtmlOutput('✅ 認証に成功しました。このタブは閉じて、スプレッドシートの操作を続けてください。');
  } else {
    return HtmlService.createHtmlOutput('❌ 認証に失敗しました。再度お試しください。');
  }
}

/**
 * 認証が有効かどうかをチェックします。
 */
function checkAuth() {
  const service = getService();
  return service.hasAccess();
}

/**
 * 手動で認証を開始するためのUIを表示します。
 */
function showAuthDialog() {
  const service = getService();
  const authUrl = service.getAuthorizationUrl();
  const html = HtmlService.createHtmlOutput(
    `<p>以下のリンクをクリックしてfreeeにログイン・認証してください。</p>\n     <p><a href="${authUrl}" target="_blank">▶ 認証はこちら</a></p>`
  ).setWidth(400).setHeight(180);
  SpreadsheetApp.getUi().showModalDialog(html, 'freee認証が必要です');
}

/**
 * 認証をリセット（解除）します。
 */
function resetAuth() {
  getService().reset();
  SpreadsheetApp.getUi().alert('認証をリセットしました。');
}

