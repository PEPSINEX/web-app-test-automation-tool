/**
 * 入力欄のスプレッドシート
 * @type {Object}
 */
const inputSheet = SpreadsheetApp.getActive().getSheetByName('入力欄')

/**
 * Selenium IDE用のテストファイル情報。ハッシュ形式
 * @type {Object}
 */
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

/**
 * 「入力欄」の必要情報を、テストケースシートの該当箇所に記述
 * @param {Object} inputData
 */
function setInputDataToTestCaseSheet(inputData) {
  let ss = inputData.testCaseSheet

  // URL, directory, メールアドレスを反映
  const urlRange = 'B1'
  const directoryRange = 'B3'
  const mailRange = 'C4'
  ss.getRange(urlRange).setValue(inputData.url)
  ss.getRange(directoryRange).setValue(inputData.directory)
  ss.getRange(mailRange).setValue(inputData.mail)

  // 検証内容を反映
  let targetRow = getLastCommandRow(inputData.testCaseSheet) + 1

}

/**
 * 該当セルを基点とし、連続する一番下のデータのあるセルの行数を返却
 * @param {Object} SpreadSheet 対象シート
 * @return {Number} セルの行数
 */
function getLastCommandRow(targetSheet) {
  let lastRow

  const col = 1
  const firstRow = 3
  lastRow = targetSheet.getRange(firstRow, col).getNextDataCell(SpreadsheetApp.Direction.DOWN).getRow()

  return lastRow
}

/**
 * 「入力欄」の必要情報をhash形式にて返却
 * @return {Object} inputData
 */
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

/**
 * 「入力欄」のテスト該当列を返却
 * 「完了チェック」列がfalseである一番上の列を該当列とする
 * @return {Number} 
 */
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

/**
 * URLを分割して返却
 * 「URL」「ドメイン」「ディレクトリ」がキー
 * @return {Object} 
 */
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

/**
 * テストケースを返却
 * 「testCase.tests[0].commands」に代入する値
 * @param {Object} SpreadSheet
 * @return {Object} 
 */
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

/**
 * スプレッドシートを開いた時の動作を定義
 */
function onOpen() {
  SpreadsheetApp.getUi()            
      .createMenu('スクリプト')            
      .addItem('サイドバー表示', 'showSidebar')            
      .addToUi();
}

/**
 * HTML形式のサイドバーを表示
 */
function showSidebar() {
  var html = HtmlService
              .createHtmlOutputFromFile('Sidebar')
              .setTitle('GAS取得データダウンロード')
  SpreadsheetApp.getUi().showSidebar(html)
}

/**
 * Selenium IDE用のテストファイルを出力
 * サイドバーにてダウンロードボタンを押下した際に実行される関数
 * @return {Object}
 */
function getData() {
  let inputData = getInputData()
  setInputDataToTestCaseSheet(inputData)

  testCase.url = inputData.domain
  testCase.tests[0].commands = getCommandList(inputData.testCaseSheet)

  return JSON.stringify(testCase)
}