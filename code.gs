const inputSheet = SpreadsheetApp.getActive().getSheetByName('入力欄')

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

//  入力情報をhash形式にて返却する
function getInputData() {
  let inputData = {}

  const siteNameCol = 2
  const urlCol = 3
  const mailCol = 4
  const verifyCol = 5  

  let row = getTargetRow()

  let testCaseSheetName = inputSheet.getRange(row, siteNameCol).getValue()
  inputData.testCaseSheet = SpreadsheetApp.getActive().getSheetByName(testCaseSheetName)
  inputData.url = inputSheet.getRange(row, urlCol).getValue()
  inputData.mail = inputSheet.getRange(row, mailCol).getValue()
  inputData.verify = inputSheet.getRange(row, verifyCol).getValue()

  return inputData
}

// inputSheetの「完了チェック」列がfalseである一番上の列を取得
function getTargetRow() {
  let targetRow

  const col = 7
  const firstRow = 2
  let lastRow = inputSheet.getRange(firstRow, col).getNextDataCell(SpreadsheetApp.Direction.DOWN).getRow()

  for(let i=firstRow;i<=lastRow;i++) {
    let isConfirmed = inputSheet.getRange(i, col).getValue()

    if(!isConfirmed) {
      targetRow = i
      break
    }
  }
  
  return targetRow
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