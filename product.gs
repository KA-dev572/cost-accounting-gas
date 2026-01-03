/**
 * v1.2 損益計算書まで 2026-01-03
 */

function executeProductCalc() {
  return flowCalc_(product);
}
function product_loadInput_(targetMonth) {
  let headerRow = inputSheet_(product).headerRow;
  let ss = inputSheet_(product).ss;
  let rows = ss.getRange(headerRow+1, 1, ss.getLastRow()-headerRow, ss.getLastColumn()).getValues();
  Logger.log(rows);
  //  [[2025-12, 仕掛品, 70450.0, 300.0, 製品X]]
  let table = filterByMonth_(rows,targetMonth);
  Logger.log(table);
  //		[{row=[2025-12, 仕掛品, 70450.0, 300.0, 製品X], idx=0.0}]
  return table;
}
function product_aggregation_(table, targetMonth) {
  // 製品名取得（後でこの製品名を各費目に反映させる工夫が必要）
  let sConfig = ss_(product.fileId,product.configSheetId);
  let aggregations = sConfig.getRange(1,1,sConfig.getLastRow(), 1).getValues();
  Logger.log(aggregations);
  //[[製品X]]

  //当面の初期値（１財）はここで設定
  let opQty = [100]; // 前月繰越数量
  let opAmount = [20000]; //前月繰越金額

  // 1 製品ごとの固定状態を記載（各製品でT勘定の開始仕訳＋固定値の計算をする） v1は売上を考えず総平均法で原価だけ出す
  let aggregation = {};
  for (let i = 0; i < aggregations.length; i++) {
    let name = aggregations[i][0];
    aggregation[name] = { //変数があるときはこの「ブラケット記法」でないと追加できない
      targetMonth: targetMonth,  //対象月
      timestamp: new Date(),  //
      openingQty: opQty[i],  //前月繰越数量
      openingAmount: opAmount[i], //前月繰越金額
      inflowQty: 0, //当月入庫数量: 一旦固定
      inflowAmount: 0,  //当月入庫金額（仕入額）
      outflowQty: 0, //当月払出量
      outflowAmount: 0, //当月払出金額（仕掛品勘定へ）
      currentQty: 0,  //最終的には翌月繰越→まずは繰越なし
      currentStorageAmount: 0,  //最終的には翌月繰越
      currentUnitCost: opAmount[i] / opQty[i],  //移動平均法と同じだが、当月受入が1回だけなので結果的に総平均法と同じになる
      status: "ready"
    };
  }
  Logger.log(aggregation);
  // {製品X={openingAmount=20000.0, outflowQty=0.0, inflowAmount=0.0, openingQty=100.0, inflowQty=0.0, currentQty=0.0, currentStorageAmount=0.0, currentUnitCost=200.0, targetMonth=2025-12, timestamp=Fri Jan 02 22:01:38 GMT+09:00 2026, outflowAmount=0.0, status=ready}}

  // 2 入力行を走査して各製品名に割り振り（入力行ごとに各仕掛品勘定へ割り振り）→総平均法
  for (let j = 0; j < table.length; j++) {
    let name = table[j].row[4]; //暫定の位置
    let state = aggregation[name];

    if (!state) continue; //nameがなければスルー
    let qty = Number(table[j].row[3]);
    if (qty > 0) {
      state.inflowQty += qty; //入庫量に加算:
      state.inflowAmount += Number(table[j].row[2]);  //入庫金額に加算 列は暫定
      state.currentQty += qty;  //現在在庫量に加算: 一旦固定なので無視
      state.currentStorageAmount += Number(table[j].row[2]); //在庫金額に加算　列は暫定
      state.currentUnitCost = state.currentStorageAmount / state.currentQty;  
    } else {  //現状は払出は売上と同一視、売り上げは別途実装のため一旦無視してよい
      state.outflowQty += qty; //出庫量に加算(負の数として表示：会計上最後は正負逆にする)          
      state.outflowAmount += qty * state.currentUnitCost; //出庫金額に加算（同上）, 単価は現在のものと同じ
      state.currentQty += qty;  //現在在庫量から減算
      state.currentStorageAmount += qty * state.currentUnitCost; //在庫金額から減算
    }
  }
  Logger.log(aggregation);
  // 	{製品X={targetMonth=2025-12, status=ready, inflowAmount=70450.0, inflowQty=300.0, outflowQty=0.0, openingAmount=20000.0, openingQty=100.0, timestamp=Fri Jan 02 22:05:22 GMT+09:00 2026, currentQty=300.0, outflowAmount=0.0, currentUnitCost=234.83333333333334, currentStorageAmount=70450.0}}
  return aggregation;  
}
function product_fillInSs_(aggregation) {
  let outcome = [];
  for (let key in aggregation) {
    let temp = aggregation[key]
    outcome.push([temp.targetMonth, temp.timestamp, temp.status, key, temp.openingQty, temp.openingAmount, temp.inflowQty, temp.inflowAmount, temp.outflowQty, temp.outflowAmount, temp.currentQty, temp.currentStorageAmount, temp.currentUnitCost]);
  }
  let outputSheet = outputSheet_(product).ss;
  outputSheet.getRange(outputSheet.getLastRow()+1, 1, outcome.length, outcome[0].length).setValues(outcome);
  Logger.log("product cost has been calculated.");
}


// 確認フロー v1.2追加
function executeProductConfirmation () {
  return flowCnfm_(product);
}
// 確認サブフロー：これらを順次呼ぶ
function product_loadOutput_ (targetMonth) {
  let headerRow = outputSheet_(product).headerRow;
  let ss = outputSheet_(product).ss;
  let temp_table = ss.getRange(headerRow+1, 1, ss.getLastRow()-headerRow, ss.getLastColumn()).getValues();
  let table = filterByMonth_(temp_table, targetMonth);
  Logger.log(table);
  // [{idx=0.0, row=[2025-12, Mon Dec 29 23:58:40 GMT+09:00 2025, ready, 材料A, 100.0, 10000.0, 100.0, 8000.0, -100.0, -9000.0, 100.0, 9000.0]},...]
  return table;
}
function product_applyConfirmation_ (table, targetMonth) {
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
  // 		{target={idx=1.0, row=[2025-12, Thu Jan 01 20:15:40 GMT+09:00 2026, confirmed, 材料A, 100.0, 10000.0, 100.0, 8000.0, -100.0, -9000.0, 100.0, 9000.0]}, type=ok}
  return decision;
}
function product_contentToSend_ (decision) {
  let content = {};
  if (decision.type === "ok") {
    let row = decision.target.row;
    content.direct = [row[0],row[1], row[3],Number(row[12])]; //[対象月、timestamp,製品, 販売原価]
  }
  Logger.log(content);
  //	{direct=[2025-12, 直接材料, 9000.0]}
  return content;
}
function product_refreshSs_ (decision) {
  if (decision.type === "ok") {
    let ss = outputSheet_(product).ss;
    let headerRow = outputSheet_(product).headerRow;
    let target = decision.target;
    let idx = Number(target.idx);
    ss.getRange(headerRow+idx+1, 3).setValue(target.row[2]);
    Logger.log("refreshing output sheet has been done.");
  } else {
    Logger.log("nothing to be refreshed.");
  }
}
function product_sendToNext_ (content, targetMonth) {
  if (Object.keys(content).length != 0) {
    let dirNext = pl;
    // let idrNext = idc;
    let dirSs = ss_(pl.fileId, pl.unitcostSheetId);
    // let idrSs = inputSheet_(idrNext).ss;
    let dc = [content.direct];
    let ic = [content.indirect];
    dirSs.getRange(dirSs.getLastRow()+1,1,dc.length,dc[0].length).setValues(dc);
    // idrSs.getRange(idrSs.getLastRow()+1,1,ic.length,ic[0].length).setValues(ic);
    Logger.log("sending content to next step has been done.");
    alertSuccess_(targetMonth);
  } else {
    Logger.log("nothing to be sent.")
  }
}
