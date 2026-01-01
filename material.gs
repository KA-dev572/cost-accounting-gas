// 2026-01-01 v1
function executeMaterialCalc() {  //ライブラリ呼び出し用ラッパー
  return flowCalc_(material);
}

// ver1 材料の受払簿から月末に材料勘定の整理を行う（在庫評価方法は移動平均法）。出力は仕掛品勘定ファイルへ（とりあえずは全額直接費として扱う。製造間接費振替部分はver2以降。）
// 共通管理用に修正。計算用はreadyを何度も繰り返し、最新タイムスタンプのみ確認対象にする。
function material_loadInput_ (targetMonth) {
  let headerRow = inputSheet_(material).headerRow;
  let ss = inputSheet_(material).ss;
  //ヘッダ行整理
  let header = ss.getRange(headerRow, 1, 1, ss.getLastColumn()).getValues()[0];
  Logger.log(header);
  //[取引番号, 日付, 材料名, 数量, 単価, 金額, 摘要]
  let nameClm = header.indexOf("材料名");
  let dateClm = header.indexOf("日付");
  let qtyClm = header.indexOf("数量");
  let uCClm = header.indexOf("単価");

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
        unitcost: rows[i][uCClm],
        name: rows[i][nameClm],     
      }
      table.push(row); 
    }
  }
  Logger.log(table);
  // [{name=材料A, unitcost=80.0, date=Wed Dec 03 00:00:00 GMT+09:00 2025, qty=100.0}, ...]
  return table;
}
function material_aggregation_ (table, targetMonth) {
  //材料名取得（入力にもここからの入力規則を導入済み）
  let sConfig = ss_(material.fileId,material.configSheetId);
  let aggregations = sConfig.getRange(1,1,sConfig.getLastRow(), 1).getValues();
  Logger.log(aggregations); //材料名  [[材料A],...]

  //当面の初期値はここで設定
  let opAQ = [100]; // 前月繰越数量
  let opAA = [10000]; //前月繰越金額

  // 1 材料ごとの初期状態を作る（各材料でT勘定の開始仕訳をする）
  let aggregation = {};
  for (let i = 0; i < aggregations.length; i++) {
    let name = aggregations[i][0];
    aggregation[name] = { //変数があるときはこの「ブラケット記法」でないと追加できない
      targetMonth: targetMonth,  //対象月
      timestamp: new Date(),  //
      openingQty: opAQ[i],  //前月繰越数量
      openingAmount: opAA[i], //前月繰越金額
      inflowQty: 0, //当月入庫数量
      inflowAmount: 0,  //当月入庫金額（仕入額）
      outflowQty: 0, //当月払出量
      outflowAmount: 0, //当月払出金額（仕掛品勘定へ）
      currentQty: opAQ[i],  //最終的には翌月繰越
      currentStorageAmount: opAA[i],  //最終的には翌月繰越
      currentUnitCost: opAA[i] / opAQ[i],  //移動平均法で単価を計算
      status: "ready"
    };
  }
  Logger.log(aggregation);

  // 2 入力行を走査して各材料に割り振り（入力行ごとに各材料勘定へ割り振り）
  for (let j = 0; j < table.length; j++) {
    let name = table[j].name;
    let state = aggregation[name];

    if (!state) continue; //nameがなければスルー
    let qty = Number(table[j].qty);
    if (qty > 0) {
      state.inflowQty += qty; //入庫量に加算
      state.inflowAmount += qty * table[j].unitcost;  //入庫金額に加算
      state.currentQty += qty;  //現在在庫量に加算
      state.currentStorageAmount += qty * table[j].unitcost; //在庫金額に加算
      state.currentUnitCost = state.currentStorageAmount / state.currentQty;  //現在在庫単価（移動平均法）
    } else {
      state.outflowQty += qty; //出庫量に加算(負の数として表示：会計上最後は正負逆にする)          
      state.outflowAmount += qty * state.currentUnitCost; //出庫金額に加算（同上）, 単価は現在のものと同じ
      state.currentQty += qty;  //現在在庫量から減算
      state.currentStorageAmount += qty * state.currentUnitCost; //在庫金額から減算
    }
  }
  Logger.log(aggregation);
  /**
   * {材料A={
   *  currentUnitCost=90.0,
   *  timestamp=Thu Jan 01 19:41:13 GMT+09:00 2026,
   *  currentStorageAmount=9000.0,
   *  inflowAmount=8000.0,
   *  outflowAmount=-9000.0,
   *  status=ready,
   *  currentQty=100.0,
   *  inflowQty=100.0,
   *  openingAmount=10000.0,
   *  outflowQty=-100.0,
   *  openingQty=100.0,
   *  targetMonth=2025-12},
   *  ...
   * }
   */
  return aggregation;
}
function material_fillInSs_ (aggregation) {
  let outcome = [];
  for (let key in aggregation) {
    let temp = aggregation[key]
    outcome.push([temp.targetMonth, temp.timestamp, temp.status, key, temp.openingQty, temp.openingAmount, temp.inflowQty, temp.inflowAmount, temp.outflowQty, temp.outflowAmount, temp.currentQty, temp.currentStorageAmount]);
  }
  let outputSheet = outputSheet_(material).ss;
  outputSheet.getRange(outputSheet.getLastRow()+1, 1, outcome.length, outcome[0].length).setValues(outcome);
  Logger.log("material cost has been aggregated.");
}


// 確認フロー
function executeMCComfirmation () {
  return flowCnfm_(material);
}
// 確認サブフロー：これらを順次呼ぶ
function material_loadOutput_ (targetMonth) {
  let headerRow = outputSheet_(material).headerRow;
  let ss = outputSheet_(material).ss;
  let temp_table = ss.getRange(headerRow+1, 1, ss.getLastRow()-headerRow, ss.getLastColumn()).getValues();
  let table = filterByMonth_(temp_table, targetMonth);
  Logger.log(table);
  // [{idx=0.0, row=[2025-12, Mon Dec 29 23:58:40 GMT+09:00 2025, ready, 材料A, 100.0, 10000.0, 100.0, 8000.0, -100.0, -9000.0, 100.0, 9000.0]},...]
  return table;
}
function material_applyConfirmation_ (table, targetMonth) {
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
function material_contentToSend_ (decision) {
  let content = {};
  if (decision.type === "ok") {
    let row = decision.target.row;
    content.direct = [row[0],material.dirName,Number(-row[9])];
    // content.indirect = [row[0], material.indName, row[4]]; //将来的拡張の余地
  }
  Logger.log(content);
  //	{direct=[2025-12, 直接材料, 9000.0]}
  return content;
}
function material_refreshSs_ (decision) {
  if (decision.type === "ok") {
    let ss = outputSheet_(material).ss;
    let headerRow = outputSheet_(material).headerRow;
    let target = decision.target;
    let idx = Number(target.idx);
    ss.getRange(headerRow+idx+1, 3).setValue(target.row[2]);
    Logger.log("refreshing output sheet has been done.");
  } else {
    Logger.log("nothing to be refreshed.");
  }
}
function material_sendToNext_ (content, targetMonth) {
  if (Object.keys(content).length != 0) {
    let dirNext = wip;
    let idrNext = idc;
    let dirSs = inputSheet_(dirNext).ss;
    let idrSs = inputSheet_(idrNext).ss;
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
