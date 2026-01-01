// 仕掛品勘定　v1　2026-01-01
function executeWIPCalc() {
  return flowCalc_(wip);
}
//集計系サブ関数群：これらを順次呼び出す
function wip_loadInput_(targetMonth) {
  let headerRow = inputSheet_(wip).headerRow;
  let ss = inputSheet_(wip).ss;
  let rows = ss.getRange(headerRow+1, 1, ss.getLastRow()-headerRow, ss.getLastColumn()).getValues();
  Logger.log(rows);
  //  [[2025-12, 間接労務費, 1650.0]]
  let table = filterByMonth_(rows,targetMonth);
  Logger.log(table);
  //	[{row=[2025-12, 間接労務費, 1650.0], idx=0.0}]
  return table;
}
function wip_aggregation_(table, targetMonth) {
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
function wip_fillInSs_(aggregation) {
  let outcome = [[aggregation.targetMonth, aggregation.timestamp, aggregation.status, aggregation.amount]];
  let outputSheet = outputSheet_(wip).ss;
  outputSheet.getRange(outputSheet.getLastRow()+1, 1, outcome.length, outcome[0].length).setValues(outcome);
}

