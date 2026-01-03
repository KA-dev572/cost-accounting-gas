/** v1.2 2026-01-03 
 * とりあえず原価計算自体は終わっていると考えたうえで、
 * 売上シートと原価シートから簡易p/lを作る
 */

function executePLCalc() {
  return flowCalc_(pl);
}
function pl_loadInput_(targetMonth) {
  let headerRow = inputSheet_(pl).headerRow;
  let ss = inputSheet_(pl).ss;
  //ヘッダ行整理
  let header = ss.getRange(headerRow, 1, 1, ss.getLastColumn()).getValues()[0];
  Logger.log(header);
  //[取引ID, 製品名, 日付, 売上個数, 売上単価, 売上額]
  const nameClm = header.indexOf("製品名");
  const dateClm = header.indexOf("日付");
  const qtyClm = header.indexOf("売上個数");
  const salesAmtClm = header.indexOf("売上額");

  //具体的取引を取得→処理→出力
  let rows = ss.getRange(headerRow+1, 1, ss.getLastRow()-headerRow, ss.getLastColumn()).getValues();
  Logger.log(rows);
  // [[1.0, Wed Dec 03 00:00:00 GMT+09:00 2025, 材料A, 100.0, 80.0, 8000.0, ], ...]
  //金額欄の関数除け
  let table = [];
  for (let i=0; i<rows.length; i++) {
    if (rows[i][qtyClm] != "") {
      let row = {
        date: rows[i][dateClm],
        qty: rows[i][qtyClm],
        salesAmount: rows[i][salesAmtClm],
        name: rows[i][nameClm],     
      }
      table.push(row); 
    }
  }
  Logger.log(table);
  // [{name=製品X, date=Wed Dec 17 00:00:00 GMT+09:00 2025, qty=50.0, salesAmount=25000.0}, {qty=100.0, name=製品X, salesAmount=45000.0, date=Thu Dec 18 00:00:00 GMT+09:00 2025}]
  return table;
}
function pl_aggregation_(table, targetMonth) {
  // 製品名取得（後でこの製品名を各費目に反映させる工夫が必要）
  let sConfig = ss_(pl.fileId,pl.configSheetId);
  let aggregations = sConfig.getRange(1,1,sConfig.getLastRow(), 1).getValues();
  Logger.log(aggregations);
  //[[製品X]]

  let aggregation = {};
  // 1 製品ごとの固定状態を記載（各製品でT勘定の開始仕訳＋固定値の計算をする）
  for (let i = 0; i < aggregations.length; i++) {
    let name = aggregations[i][0];
    aggregation[name] = { //変数があるときはこの「ブラケット記法」でないと追加できない
      targetMonth: targetMonth,  //対象月
      timestamp: new Date(),  //
      salesQty: 0, //当月売上量
      salesAmount: 0, //当月売上額
      status: "ready"
    };
  }
  Logger.log(aggregation);
  // 

  // 2 入力行を走査して各製品名に割り振り（入力行ごとに各仕掛品勘定へ割り振り）→総平均法
  for (let j = 0; j < table.length; j++) {
    let name = table[j].name;
    let state = aggregation[name];

    if (!state) continue; //nameがなければスルー
    let qty = Number(table[j].qty);
    state.salesQty += qty;  //現在在庫量に加算: 一旦固定なので無視
    state.salesAmount += Number(table[j].salesAmount); //在庫金額に加算　列は暫定  
  }
  Logger.log(aggregation);
  // 	
  return aggregation;  
}
function pl_fillInSs_(aggregation) {
  let outcome = [];
  for (let key in aggregation) {
    let temp = aggregation[key]
    outcome.push([temp.targetMonth, temp.timestamp, temp.status, key, temp.salesQty, temp.salesAmount, ]);
  }
  let outputSheet = outputSheet_(pl).ss;
  outputSheet.getRange(outputSheet.getLastRow()+1, 1, outcome.length, outcome[0].length).setValues(outcome);
  Logger.log("sales amount has been calculated.");
}


//　確認フロー
function executePLConfirmation () {
  return flowCnfm_(pl);
}
function pl_loadOutput_(targetMonth) {
  let headerRow = outputSheet_(pl).headerRow;
  let ss = outputSheet_(pl).ss;
  let temp_table = ss.getRange(headerRow+1, 1, ss.getLastRow()-headerRow, ss.getLastColumn()).getValues();
  let table = filterByMonth_(temp_table, targetMonth);
  Logger.log(table);
  // [{row=[2025-12, Sat Jan 03 15:23:09 GMT+09:00 2026, ready, 製品X, 150.0, 70000.0], idx=0.0}, {idx=1.0, row=[2025-12, Sat Jan 03 15:38:55 GMT+09:00 2026, ready, 製品X, 150.0, 70000.0]}]
  return table;
}
function pl_applyConfirmation_(table, targetMonth) {
  let decision = pickLatestUnconfirmed_(table); 
  // ↑2財への拡張があるのでここはこのままではいけない。v2では材料ごとにループしてconfirmしなければならない。
  if (decision.type === "ok") {
    decision.target.row[2]="confirmed";
  } else if (decision.type ==="alreadyConfirmed") {
    alertAlreadyConfirmed_(targetMonth);
  } else if (decision.type === "none") {
    alertNone_(targetMonth);
  }
  Logger.log(decision);
  // 	{target={idx=1.0, row=[2025-12, Sat Jan 03 15:38:55 GMT+09:00 2026, confirmed, 製品X, 150.0, 70000.0]}, type=ok}
  return decision;
}
function pl_contentToSend_(decision) {
  let content = [];
  if (decision.type === "ok") {
    let salesAmount = Number(decision.target.row[5]);
    let salesQty = Number(decision.target.row[4]);

    const ucTableSs = ss_(pl.fileId,pl.unitcostSheetId);
    const ucTable = ucTableSs.getRange(2,1,ucTableSs.getLastRow()-1, ucTableSs.getLastColumn()).getValues();
    Logger.log(ucTable);
    let unitcost = 0;
    for (row of ucTable) {
      if (row[0]===getTargetMonth_(pl)){
        unitcost = Number(row[3]);
      }
    }
    content = [[salesAmount],[salesQty*unitcost]];
  }
  Logger.log(content);
  //	[[70000.0], [35225.0]]
  return content;
}
function pl_refreshSs_ (decision) {
  if (decision.type === "ok") {
    let ss = outputSheet_(pl).ss;
    let headerRow = outputSheet_(pl).headerRow;
    let target = decision.target;
    let idx = Number(target.idx);
    ss.getRange(headerRow+idx+1, 3).setValue(target.row[2]);
    Logger.log("refreshing output sheet has been done.");
  } else {
    Logger.log("nothing to be refreshed.");
  }
}
function pl_sendToNext_ (content,targetMonth) {
  if (content.length != 0) {
    let dirNext = pl;
    let dirSs = ss_(dirNext.fileId, dirNext.plSheetId);
    let dc = content;
    dirSs.getRange(2,2,dc.length,dc[0].length).setValues(dc);
    // idrSs.getRange(idrSs.getLastRow()+1,1,ic.length,ic[0].length).setValues(ic);
    Logger.log("sending content to next step has been done.");
    // alertSuccess_(targetMonth);
  } else {
    Logger.log("nothing to be sent.")
  }
}
