const sheet = SpreadsheetApp.getActive().getSheetByName('testData')

let testCase = {
  version: '2.0',
  name: 'form_test',
  url: '',
  tests: [{
    name: 'main',
    commands: ''
  }],
  suites: [{
    name: '',
    persistSession: false,
    parallel: false,
    timeout: 300,
    tests: []
  }],
  urls: [],
  plugins: []
}

function getDomain() {
  let url = sheet.getRange(2,2).getValue()
  
  let regexpValue = 'https:\/\/([\\s\\S]*?)\/' 
  let regexp = new RegExp(regexpValue)
  let result = url.match(regexp)[0]

  return result.slice(0, -1)
}

function setUrlToSheet() {
  let url = sheet.getRange(2,2).getValue()
  let domain = getDomain()

  let pass = url.replace(domain, '')
  sheet.getRange(4,2).setValue(pass)
}

function getCommandList() {
  let commandList = []

  let firstRow = 4
  let lastRow = sheet.getRange(4, 1).getNextDataCell(SpreadsheetApp.Direction.DOWN).getRow()
  let rowRange = lastRow - firstRow + 1
  let values = sheet.getRange(4, 1, rowRange, 3).getValues()
  
  for(let i=0;i<values.length;i++) {
    let tmp = {}

    tmp.command = values[i][0]
    tmp.target = values[i][1]
    tmp.value = values[i][2]

    commandList.push(tmp)
  }
  return commandList
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
  testCase.url = getDomain()
  setUrlToSheet()
  testCase.tests[0].commands = getCommandList()

  return JSON.stringify(testCase)
}