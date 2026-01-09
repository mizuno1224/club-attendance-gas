/**
 * メイン: Web App / メニュー / トリガー
 */

function doGet(e) {
  var page = (e && e.parameter && e.parameter.page) ? String(e.parameter.page).toLowerCase() : "member";
  var fileName = (page === "admin") ? "admin" : "member";
  return HtmlService.createHtmlOutputFromFile(fileName).setTitle("女子軟式野球部").setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
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
  var html = HtmlService.createHtmlOutputFromFile('member').setTitle('部員ページ');
  SpreadsheetApp.getUi().showSidebar(html);
}

function openAdminSidebar() {
  var html = HtmlService.createHtmlOutputFromFile('admin').setTitle('管理者ページ');
  SpreadsheetApp.getUi().showSidebar(html);
}

// ...既存のコードの末尾...

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}