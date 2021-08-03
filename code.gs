let domain = 'https://sfc.jp'
let commandsValue = [{
  command: 'open',
  target: '/ie/contact/inquiry/home/form.php'
}, {
  command: 'close'
}]
const suitesValue = [{
  name: '',
  persistSession: false,
  'parallel': false,
  'timeout': 300,
  'tests': []
}]

let testCase = {
  version: '2.0',
  name: 'form_test',
  url: domain,
  tests: [{
    name: 'main',
    commands: commandsValue
  }],
  suites: suitesValue,
  'urls': [],
  'plugins': []
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
  return JSON.stringify(testCase)
}