/**
 * 入力欄と検証内容マスタのスプレッドシート
 * @type {Object}
 */
const inputSheet = SpreadsheetApp.getActive().getSheetByName('入出力')
const verifyMasterSheet = SpreadsheetApp.getActive().getSheetByName('検証内容マスタ')

/**
 * inputSheetシートのセル情報
 * @type {Object} hash
 */
const inputSheetRange = {
  col: {
    siteName: 2,
    url: 3,
    mail: 4,
    verify: 5,
    isFinished: 7,
    downloadTime: 8
  },
  row: {
    data: 2
  }
}

/**
 * testCaseシートのセル情報
 * @type {Object} hash
 */
const testCaseRange = {
  url: 'B1',
  directory: 'B3',
  mail: 'C4',
  row: {
    data: 3,
    range: 1
  },
  col: {
    command: 1,
    target: 2,
    value: 3,
    range: 3
  }
}

/**
 * Selenium IDE用のテストファイル情報。ハッシュ形式
 * @type {Object}
 */
let testCase = {
  version: '2.0',
  name: '',
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
  ss.getRange(testCaseRange.url).setValue(inputData.url)
  ss.getRange(testCaseRange.directory).setValue(inputData.directory)
  ss.getRange(testCaseRange.mail).setValue(inputData.mail)

  // 検証内容を反映
  let command = getVerifyCommand(inputData.verify)

  let row = testCaseRange.row.data
  let col = testCaseRange.col.command
  let rowRange = testCaseRange.row.range
  let colRange = testCaseRange.col.range

  let targetRow = getLastRow(ss, row, col) + 1
  ss.getRange(targetRow, col, rowRange, colRange).setValues(command)
}

/**
 * テストケースシートに反映した入力情報を削除
 */
function deleteInputDataOfTestCaseSheet() {
  let ss = getInputData().testCaseSheet

  // URL, directory, メールアドレスを削除
  ss.getRange(testCaseRange.url).clearContent()
  ss.getRange(testCaseRange.directory).clearContent()
  ss.getRange(testCaseRange.mail).clearContent()

  // 検証内容を削除
  let row = testCaseRange.row.data
  let col = testCaseRange.col.command
  let rowRange = getLastRow(ss, row, col) - row + 1
  let colRange = 1

  let commandList = ss.getRange(row, col, rowRange, colRange).getValues().flat()
  let targetRow
  for(let i=0;i<commandList.length;i++) {
    let isMatch = commandList[i].match(/assert/)
    if( isMatch ) {
      targetRow = i + row
      break
    }
  }

  ss.getRange(targetRow, col, testCaseRange.row.range, testCaseRange.col.range).clearContent()
}

/**
 * 「入力欄」の必要情報を、テストケースシートの該当箇所に記述
 * @param {Object} inputData
 */
function getVerifyCommand(targetVerify) {
  let command = []

  // 入力欄と一致するマスターの行を取得
  const verifyCol = 2
  const firstRow = 2
  let rowRange = getLastRow(verifyMasterSheet, firstRow, verifyCol) - 1

  let verifyKeyList = verifyMasterSheet.getRange(firstRow, verifyCol, rowRange).getValues().flat()
  let index = verifyKeyList.indexOf(targetVerify)
  let targetRow = index + firstRow

  // マスターの行のコマンドを取得
  const commandCol = 3
  const commandColRange = 3
  const commandRow = targetRow
  const commandRowRange = 1

  command = verifyMasterSheet.getRange(commandRow, commandCol, commandRowRange, commandColRange).getValues()

  return command
}

/**
 * 特定セルを基点とし、連続する一番下のデータのあるセルの行数を返却
 * @param {Object} ss 対象シート
 * @param {Number} col 特定セルの列数
 * @param {Number} row 特定セルの行数
 * @return {Number} セルの行数
 */
function getLastRow(ss, row, col) {
  let lastRow
  lastRow = ss.getRange(row, col).getNextDataCell(SpreadsheetApp.Direction.DOWN).getRow()
  return lastRow
}

/**
 * 「入力欄」の必要情報をhash形式にて返却
 * @return {Object} inputData
 */
function getInputData() {
  let inputData = {}

  let row = getTargetRow()

  inputData.siteName      = inputSheet.getRange(row, inputSheetRange.col.siteName).getValue()
  inputData.testCaseSheet = SpreadsheetApp.getActive().getSheetByName(inputData.siteName)
  inputData.url           = inputSheet.getRange(row, inputSheetRange.col.url).getValue()
  inputData.mail          = inputSheet.getRange(row, inputSheetRange.col.mail).getValue()
  inputData.verify        = inputSheet.getRange(row, inputSheetRange.col.verify).getValue()
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

  let lastRow = getLastRow(inputSheet, inputSheetRange.row.data, inputSheetRange.col.isFinished)
  for(let i=inputSheetRange.row.data;i<=lastRow;i++) {
    let isConfirmed = inputSheet.getRange(i, inputSheetRange.col.isFinished).getValue()

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

  let rowRange = getLastRow(ss, testCaseRange.row.data, testCaseRange.col.command) - testCaseRange.row.data + 1
  let values = ss.getRange(testCaseRange.row.data, testCaseRange.col.command, rowRange, testCaseRange.col.range).getValues()

  for(let i=0;i<values.length;i++) {
    let tmp = {}

    tmp.command = values[i][testCaseRange.col.command - 1]
    tmp.target = values[i][testCaseRange.col.target - 1]
    tmp.value = values[i][testCaseRange.col.value - 1]

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
  testCase.name = inputData.siteName

  data = {
    testFile: JSON.stringify(testCase),
    fileName: inputData.siteName + '.side'
  }

  return data
}