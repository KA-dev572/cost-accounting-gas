// 集計フロー v1対応 2026-01-01
function executeLaborCalc() {  //ライブラリ呼び出し用ラッパー
  return flowCalc_(labor);
}
// ver1 確定勤怠簿から月末に労務費勘定の整理を行う。出力は一旦労務費集計シートへ→直接/間接まで実装。2財以上製造の場合はver2以降
// 集計個別サブフロー　これらを呼び出す
function labar_loadInput_ (targetMonth) {
  let headerRow = inputSheet_(labor).headerRow;
  let ss = inputSheet_(labor).ss;
  //ヘッダ行整理
  let header = ss.getRange(headerRow, 1, 1, ss.getLastColumn()).getValues()[0];
  Logger.log(header);  //[日付, 職員氏名, 作業時間, 作業内容, 時給, 金額]: 勤怠簿シートに職員とひもづいた時給で日額を計算できる想定
  let dateClm = header.indexOf("日付");
  let qtyClm = header.indexOf("作業時間");
  let taskClm = header.indexOf("作業内容");
  let costClm = header.indexOf("金額");
  Logger.log([dateClm, qtyClm, taskClm, costClm]); //[0.0, 2.0, 3.0, 5.0]
  let rows = ss.getRange(headerRow+1, 1, ss.getLastRow()-headerRow, ss.getLastColumn()).getValues();
  Logger.log(rows);
  //  [[Wed Dec 03 00:00:00 GMT+09:00 2025, 甲, 7.5, 製造, 1100.0, 8250.0],...]
  // 金額欄の関数除け+連想配列づくり
  let table = [];
  for (let i=0; i<rows.length; i++) {
    if (rows[i][qtyClm] != "") {
      let row = {
        date: rows[i][dateClm],
        qty: rows[i][qtyClm],
        task: rows[i][taskClm],
        cost: rows[i][costClm],
      }
      table.push(row); 
    }
  }
  Logger.log(table);
  //	[[Wed Dec 03 00:00:00 GMT+09:00 2025, 甲, 7.5, 製造, 1100.0, 8250.0],...]
  return table;
}
function labor_aggregation_ (table, targetMonth) {
  // 1 初期状態を作る（0,0）
  let aggregation = {
    targetMonth: targetMonth,
    timestamp: new Date(),
    status: "ready",
    directCost: 0,
    indirectCost: 0
  };
  // 2 入力行を走査して直接/間接に割り振り（入力行ごと）
  for (let i = 0; i < table.length; i++) {
    let task = table[i].task;  //作業内容
    let cost = Number(table[i].cost);  //入力シート側で数値以外を弾くよう設計済み
    if (task == "製造") { //変数化してconfigで保持してもよい
      aggregation.directCost += cost;  //直接労務費に加算
    } else if (task == "手待") {
      aggregation.indirectCost += cost;  //間接労務費に加算
    }
  }
  Logger.log(aggregation);  
  // {targetMonth=2025-12, status=ready, timestamp=Thu Jan 01 16:03:52 GMT+09:00 2026, directCost=28250.0, indirectCost=1650.0}
  return aggregation;
}
function labor_fillInSs_ (aggregation) {
  let outcome = [[aggregation.targetMonth, aggregation.timestamp, aggregation.status, aggregation.directCost, aggregation.indirectCost]];
  let outputSheet = outputSheet_(labor).ss;
  outputSheet.getRange(outputSheet.getLastRow()+1, 1, outcome.length, outcome[0].length).setValues(outcome);
  Logger.log("labor cost has been aggregated.");
}



// 確認フロー v0.4対応 2026-01-01
function executeLCComfirmation () {
  return flowCnfm_(labor);
}
// 確認サブフロー
function labor_loadOutput_ (targetMonth) {
  let headerRow = outputSheet_(labor).headerRow;
  let ss = outputSheet_(labor).ss;
  let temp_table = ss.getRange(headerRow+1, 1, ss.getLastRow()-headerRow, ss.getLastColumn()).getValues();
  let table = filterByMonth_(temp_table, targetMonth);
  Logger.log(table);
  // [{idx=0.0, row=[2025-12, Tue Dec 30 00:14:29 GMT+09:00 2025, ready, 28250.0, 1650.0]},...]
  return table;
}
function labor_applyConfirmation_ (table, targetMonth) {
  let decision = pickLatestUnconfirmed_(table);
  // 
  if (decision.type === "ok") {
    decision.target.row[2]="confirmed";
  } else if (decision.type ==="alreadyConfirmed") {
    alertAlreadyConfirmed_(targetMonth);
  } else if (decision.type === "none") {
    alertNone_(targetMonth);
  }
  Logger.log(decision);
  // 	{target={row=[2025-12, Thu Jan 01 17:00:05 GMT+09:00 2026, confirmed, 28250.0, 1650.0], idx=4.0}, type=ok}
  return decision;
}
function labor_contentToSend_ (decision) {
  let content = {};
  if (decision.type === "ok") {
    let row = decision.target.row;
    content.direct = [row[0],labor.dirName,row[3]];
    content.indirect = [row[0], labor.indName, row[4]];
  }
  Logger.log(content);
  //	{indirect=[2025-12, 間接労務費, 1650.0], direct=[2025-12, 直接労務費, 28250.0]}
  return content;
}
function labor_refreshSs_ (decision) {
  if (decision.type === "ok") {
    let ss = outputSheet_(labor).ss;
    let headerRow = outputSheet_(labor).headerRow;
    let target = decision.target;
    let idx = Number(target.idx);
    ss.getRange(headerRow+idx+1, 3).setValue(target.row[2]);
    Logger.log("refreshing output sheet has been done.");
  } else {
    Logger.log("nothing to be refreshed.");
  }
}
function labor_sendToNext_ (content, targetMonth) {
  if (Object.keys(content).length != 0) {
    let dirNext = wip;
    let idrNext = idc;
    let dirSs = inputSheet_(dirNext).ss;
    let idrSs = inputSheet_(idrNext).ss;
    let dc = [content.direct];
    let ic = [content.indirect];
    dirSs.getRange(dirSs.getLastRow()+1,1,dc.length,dc[0].length).setValues(dc);
    idrSs.getRange(idrSs.getLastRow()+1,1,ic.length,ic[0].length).setValues(ic);
    Logger.log("sending content to next step has been done.");
    alertSuccess_(targetMonth);
  } else {
    Logger.log("nothing to be sent.")
  }
}
