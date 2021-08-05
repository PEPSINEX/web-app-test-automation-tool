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

function setInputDataToTestCaseSheet(inputData) {
  let ss = inputData.testCaseSheet

  const urlRange = 'B1'
  const directoryRange = 'B3'
  const mailRange = 'C4'

  ss.getRange(urlRange).setValue(inputData.url)
  ss.getRange(directoryRange).setValue(inputData.directory)
  ss.getRange(mailRange).setValue(inputData.mail)
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
  inputData.url           = inputSheet.getRange(row, urlCol).getValue()
  inputData.mail          = inputSheet.getRange(row, mailCol).getValue()
  inputData.verify        = inputSheet.getRange(row, verifyCol).getValue()
  inputData.domain        = splitUrl(inputData.url).domain
  inputData.directory     = splitUrl(inputData.url).directory

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

// URLを分割してハッシュで返す。キーは「URL」「ドメイン」「ディレクトリ」の3つ
function splitUrl(url) {
  let urlArr = {}

  let domain
  let regexpValue = 'https:\/\/([\\s\\S]*?)\/' 
  let regexp = new RegExp(regexpValue)
  domain = url.match(regexp)[0].slice(0, -1)

  let directory
  directory = url.replace(domain, '')

  urlArr = {
    url: url,
    domain: domain,
    directory: directory
  }

  return urlArr
}

function getCommandList(ss) {
  let commandList = []

  const commandCol = 1
  const targetCol = 2
  const valueCol = 3
  const commandRange = 3
  const firstRow = 3
  let lastRow = ss.getRange(firstRow, commandCol).getNextDataCell(SpreadsheetApp.Direction.DOWN).getRow()
  let rowRange = lastRow - firstRow + 1
  let values = ss.getRange(firstRow, commandCol, rowRange, commandRange).getValues()
  
  for(let i=0;i<values.length;i++) {
    let tmp = {}

    tmp.command = values[i][commandCol-1]
    tmp.target = values[i][targetCol-1]
    tmp.value = values[i][valueCol-1]

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
  let inputData = getInputData()
  setInputDataToTestCaseSheet(inputData)

  testCase.url = inputData.domain
  testCase.tests[0].commands = getCommandList(inputData.testCaseSheet)

  Logger.log(testCase)

  return JSON.stringify(testCase)
}