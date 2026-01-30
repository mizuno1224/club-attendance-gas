/**
 * メイン: Web App / メニュー / トリガー
 */

function doGet(e) {
  var page = (e && e.parameter && e.parameter.page) ? String(e.parameter.page).toLowerCase() : "member";
  var fileName = (page === "admin") ? "admin" : "member";
  
  // 修正: createHtmlOutputFromFile ではなく createTemplateFromFile を使い、evaluate() でスクリプトレットを実行する
  return HtmlService.createTemplateFromFile(fileName)
    .evaluate()
    .setTitle("女子軟式野球部")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover');
}

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu("部活")
    .addItem("部員ページ（新しいタブ）", "openMember")
    .addItem("管理者ページ（新しいタブ）", "openAdmin")
    .addSeparator()
    .addItem("部員ページ（サイドバー表示）", "openMemberSidebar")
    .addItem("管理者ページ（サイドバー表示）", "openAdminSidebar")
    .addToUi();
}

function _getWebAppUrlOrAlert() {
  var url = ScriptApp.getService().getUrl();
  if (!url) {
    SpreadsheetApp.getUi().alert("先にウェブアプリとしてデプロイしてください（デプロイ→新しいデプロイ→ウェブアプリ）。");
    return null;
  }
  return url;
}

function openMember() {
  var url = _getWebAppUrlOrAlert();
  if (!url) return;
  SpreadsheetApp.getUi().showModalDialog(HtmlService.createHtmlOutput('<p><a href="' + url + '" target="_blank">部員ページを開く</a></p>'), "部員ページ");
}

function openAdmin() {
  var url = _getWebAppUrlOrAlert();
  if (!url) return;
  SpreadsheetApp.getUi().showModalDialog(HtmlService.createHtmlOutput('<p><a href="' + url + '?page=admin" target="_blank">管理者ページを開く</a></p>'), "管理者ページ");
}

function openMemberSidebar() {
  // 修正: サイドバーも同様にテンプレート評価を行う
  var html = HtmlService.createTemplateFromFile('member').evaluate().setTitle('部員ページ');
  SpreadsheetApp.getUi().showSidebar(html);
}

function openAdminSidebar() {
  // 修正: サイドバーも同様にテンプレート評価を行う
  var html = HtmlService.createTemplateFromFile('admin').evaluate().setTitle('管理者ページ');
  SpreadsheetApp.getUi().showSidebar(html);
}

// HTMLファイル内で <?!= include('filename'); ?> として呼び出す関数
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}