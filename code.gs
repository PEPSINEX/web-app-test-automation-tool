let hash = {
  key1: 'value1',
}

function convertJsonFromHash(hash) {
  return JSON.stringify(hash)
}

function onOpen() {
  SpreadsheetApp.getUi()
      .createMenu('スクリプト')
      .addItem('サイドバー表示', 'showSidebar')
      .addToUi();
}

function showSidebar() {
  var html = HtmlService
              .createHtmlOutputFromFile('Sidebar')
              .setTitle('GAS取得データダウンロード')
  SpreadsheetApp.getUi().showSidebar(html)
}

function getData() {
  return convertJsonFromHash(hash)
}