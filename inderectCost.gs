function executeIdc() {
  return calcIdc_();}

function calcIdc_() {
  const ids = idcId();
  // 1 入力読み取り
  // 1.1当月情報取得  2025-12
  const targetMonth = getTargetMonth_(ids);
  // 1.2費用入力シート取得  [[2025-12, 間接労務費, 1650.0]] //入力シートで対象月列を「書式なしテキスト」にしておくこと
  const table = loadAggregationTable_(ids, "input");
  //　1.3対象月抽出  [{idx=0.0, row=[2025-12, 間接労務費, 1650.0]}]
  const rows = filterByMonth_(table, targetMonth);  //labor.gsの使いまわし

  // 2 処理 {amount=1650.0, targetMonth=2025-12, status=ready, sendTo=仕掛品}
  let total = monthlySum_ (rows, fromIdcTo_().name);

  // 3 出力
  let print = [[total.targetMonth, total.timestamp, total.status, total.amount]];
  const outputSheet = outputSheet_(ids).ss;
  outputSheet.getRange(outputSheet.getLastRow()+1, 1, print.length, print[0].length).setValues(print);
  //仮; デバッグ用
  // return print;
}

//確認処理用
function executeIDCConfirmation(){
  return idcConfirmation_();
}
function idcConfirmation_() {
  const ids = idcId();
  const targetMonth = getTargetMonth_(ids);
  // ① データ取得
  const table = loadAggregationTable_(ids, "output");
  // ② 対象月抽出
  const rows = filterByMonth_(table, targetMonth);
  // ③ 確定可否判定
  const decision = pickLatestUnconfirmed_(rows);

  if (decision.type === "alreadyConfirmed") {
    alertAlreadyConfirmed_(targetMonth);  // laborから
  } else if (decision.type === "none") {
    alertNone_ (targetMonth); //laborから
  } else if (decision.type === "ok") {
    // ④ 状態遷移
    applyConfirmation_(table, decision);
    // ⑤ 他シート送信
    sendToWIP_(ids,fromIdcTo_(),decision);
    // ⑥ 書き戻し
    saveAggregationTable_(ids, table);
    alertSuccess_(targetMonth);
  }
  // デバッグ用
  // return {targetMonth, table, rows, decision};
}

//サブ関数群
//シートconfig系
function ss_(file, sheet){
  const ss = SpreadsheetApp.openById(file).getSheetById(sheet);
  return ss;
}
function inputSheet_ (obj) {
  const inputSheet = ss_(obj.fileId, obj.inputSheetId);
  return {headerRow: 2, ss: inputSheet};
}
function outputSheet_(obj) {
  const outputSheet = ss_(obj.fileId, obj.outputSheetId);
  return {headerRow: 1, ss: outputSheet};
}
function fromIdcTo_ () {
  return wipId();
}

//集計処理系
function getTargetMonth_(obj){
  const targetMonth = inputSheet_(obj).ss.getRange(1,2).getValue();  //入力箇所指定B1セル
  return targetMonth;
}

function loadAggregationTable_(obj, whichSheet) {
  let headerRow = 0;
  let ss = null;
  if (whichSheet == "input") {
    headerRow = inputSheet_(obj).headerRow;
    ss = inputSheet_(obj).ss;
  } else if (whichSheet == "output") {
    headerRow = outputSheet_(obj).headerRow;
    ss = outputSheet_(obj).ss;
  }
  const table = ss.getRange(headerRow+1, 1, ss.getLastRow()-headerRow, ss.getLastColumn()).getValues();
  return table;
}

function monthlySum_(rows, next){
  //[{idx=0.0, row=[2025-12, 間接労務費, 1650.0]}, {...}]この形式のものを集計
  //とりあえずは1財生産なので製品ごと配賦は考えない
  let sums = {targetMonth:rows[0].row[0],
              timestamp: new Date(),
              sendTo: next,
              amount: 0,
              status: "ready"  
  };
  for (element of rows){
    sums.amount += Number(element.row[2]);
  }
  return sums;
}

//確認処理系
function applyConfirmation_ (table, decision) {
  const idx = Number(decision.target.idx);
  table[idx][2] = "confirmed";
  return table;
}
function sendToWIP_(ids, next, decision) {
    const contentToSend = [[
    decision.target.row[0],
    ids.name,
    decision.target.row[3]
  ]];
  // return contentToSend;
  const ss1 = SpreadsheetApp
    .openById(next.fileId)
    .getSheetById(next.inputSheetId);
  ss1.getRange(ss1.getLastRow()+1,1,contentToSend.length,contentToSend[0].length).setValues(contentToSend);
}
function saveAggregationTable_(obj, table) {
  const ss = outputSheet_(obj).ss;
  const headerRow = outputSheet_(obj).headerRow;
  ss.getRange(headerRow+1, 1, table.length, table[0].length).setValues(table);
}
function alertSuccess_(targetMonth) {
  SpreadsheetApp.getUi().alert(`${targetMonth}分の集計額を次工程へ送信しました。`)
}

function dev () {
  Logger.log();
}
