// 2026-01-01 ver1 done
//集計フロー起動
function executeIdcCalc() {
  return flowCalc_(idc);
}
//集計系サブ関数群：これらを順次呼び出す
function idc_loadInput_(targetMonth) {
  let headerRow = inputSheet_(idc).headerRow;
  let ss = inputSheet_(idc).ss;

  let rows = ss.getRange(headerRow+1, 1, ss.getLastRow()-headerRow, ss.getLastColumn()).getValues();
  // Logger.log(rows);
  //  [[2025-12, 間接労務費, 1650.0]]
  let table = filterByMonth_(rows,targetMonth);
  // Logger.log(table);
  //	[{row=[2025-12, 間接労務費, 1650.0], idx=0.0}]
  return table;
}
function idc_aggregation_(table, targetMonth) {
  let aggregation = {targetMonth:targetMonth,
              timestamp: new Date(),
              amount: 0,
              status: "ready"  
  };
  for (element of table){
    aggregation.amount += Number(element.row[2]);
  }
  Logger.log(aggregation);
  // {amount=1650.0, targetMonth=2025-12, status=ready, timestamp=Thu Jan 01 12:40:55 GMT+09:00 2026}
  return aggregation;
}
function idc_fillInSs_(aggregation) {
  let outcome = [[aggregation.targetMonth, aggregation.timestamp, aggregation.status, aggregation.amount]];
  let outputSheet = outputSheet_(idc).ss;
  outputSheet.getRange(outputSheet.getLastRow()+1, 1, outcome.length, outcome[0].length).setValues(outcome);
  return;
}


//確認処理フロー実行
function executeIDCConfirmation(){
  return flowCnfm_(idc);
}
//確認フローサブ関数：これらを順次呼び出す
function idc_loadOutput_(targetMonth) {
  let headerRow = outputSheet_(idc).headerRow;
  let ss = outputSheet_(idc).ss;
  let temp_table = ss.getRange(headerRow+1, 1, ss.getLastRow()-headerRow, ss.getLastColumn()).getValues();
  let table = filterByMonth_(temp_table, targetMonth);
  Logger.log(table);
  // [{row=[2025-12, Wed Dec 31 01:57:03 GMT+09:00 2025, ready, 1650.0], idx=0.0}, {row=[2025-12, Thu Jan 01 12:47:34 GMT+09:00 2026, ready, 1650.0], idx=1.0}]
  return table;
}
function idc_applyConfirmation_(table, targetMonth) {
  let decision = pickLatestUnconfirmed_(table);
  // {target={idx=1.0, row=[2025-12, Thu Jan 01 12:47:34 GMT+09:00 2026, ready, 1650.0]}, type=ok}
  if (decision.type === "ok") {
    decision.target.row[2]="confirmed";
  } else if (decision.type ==="alreadyConfirmed") {
    alertAlreadyConfirmed_(targetMonth);
  } else if (decision.type === "none") {
    alertNone_(targetMonth);
  }
  Logger.log(decision);
  //{target={row=[2025-12, Thu Jan 01 12:47:34 GMT+09:00 2026, confirmed, 1650.0], idx=1.0}, type=ok}
  return decision;
}
function idc_contentToSend_ (decision) {
  let content = [];
  if (decision.type === "ok") {
    let row = decision.target.row;
    content.push([row[0],idc.name,row[3]]);
  }
  Logger.log(content);
  //	[[2025-12, 製造間接費, 1650.0]]
  return content;
}
function idc_refreshSs_ (decision) {
  if (decision.type === "ok") {
    let ss = outputSheet_(idc).ss;
    let headerRow = outputSheet_(idc).headerRow;
    let target = decision.target;
    let idx = Number(target.idx);
    ss.getRange(headerRow+idx+1, 3).setValue(target.row[2]);
    Logger.log("refreshing output sheet has been done.");
  } else {
    Logger.log("nothing to be refreshed.");
  }
}
function idc_sendToNext_ (content, targetMonth) {
  if (content.length != 0) {
    let next = wip;
    let ss = inputSheet_(next).ss;
    ss.getRange(ss.getLastRow()+1,1,content.length,content[0].length).setValues(content);
    Logger.log("sending content to next step has been done.");
    alertSuccess_(targetMonth);
  } else {
    Logger.log("nothing to be sent.")
  }
}
