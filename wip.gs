// 仕掛品勘定　v1.1　2026-01-02 製品勘定へ送信。月末仕掛品なし
function executeWIPCalc() {
  return flowCalc_(wip);
}
//集計系サブ関数群：これらを順次呼び出す
function wip_loadInput_(targetMonth) {
  let headerRow = inputSheet_(wip).headerRow;
  let ss = inputSheet_(wip).ss;
  let rows = ss.getRange(headerRow+1, 1, ss.getLastRow()-headerRow, ss.getLastColumn()).getValues();
  let sConfig = ss_(wip.fileId,wip.configSheetId);
  let aggregations = sConfig.getRange(1,1,sConfig.getLastRow(), 1).getValues();
  for (row of rows) {
    row.push(aggregations[0][0]); //  1財だからとりあえずこれで
  }
  Logger.log(rows);
  //  [[2025-12, 製造間接費, 1650.0, 製品X], [2025-12, 直接労務費, 28250.0, 製品X], [2025-12, 直接材料, 9000.0, 製品X]]
  let table = filterByMonth_(rows,targetMonth);
  Logger.log(table);
  //	[{row=[2025-12, 製造間接費, 1650.0, 製品X], idx=0.0}, {idx=1.0, row=[2025-12, 直接労務費, 28250.0, 製品X]}, {row=[2025-12, 直接材料, 9000.0, 製品X], idx=2.0}]
  return table;
}
/** 2026-01-02検討 仕掛品勘定自体も複数製品を含む。材料費と同じく最終的には複数行の出力になるが… 
 * あるいは、材料ごとに勘定を分けているならシートも分ける方が簡単に済みはするものの…
 * そもそも現時点v1の集計だけではどの製品に飛んでいるかわからない仕様：複数財導入時にconfigで製品名を明示し、
 * 直接費用の配賦や部門別配賦の時点でどの製品に飛んでいくかを明記しなければならない。
 * とりあえず、現状では材料勘定のロジックを使いまわしやすいようにinpuyシートに製品名を手入力した。どうせv2で拡張する。
 * 
 * 次に、仕掛品勘定は加工進捗を考慮しなければならない。これはいままでになかった要素。移動平均も使わない。
 * とりあえず、繰越額がどれだけあるかだけがわかればよいとする。この段階で単価は出す必要がない。
 * 
 →検討結果：月末仕掛品なし：製造はすぐに終わる製品だと考えればそう不自然でもない。
*/
function wip_aggregation_(table, targetMonth) {
  // 製品名取得（後でこの製品名を各費目に反映させる工夫が必要）
  let sConfig = ss_(wip.fileId,wip.configSheetId);
  let aggregations = sConfig.getRange(1,1,sConfig.getLastRow(), 1).getValues();
  Logger.log(aggregations);
  //[[製品X]]

  //当面の初期値（１財）はここで設定
  let opQty = [0]; // 前月繰越数量
  let opAmount = [0]; //前月繰越金額
  let inflowQty = [300];  //当月取り掛かり数量 

  // 1 製品ごとの固定状態を記載（各製品でT勘定の開始仕訳＋固定値の計算をする）、ただし数量は一旦固定
  let aggregation = {};
  for (let i = 0; i < aggregations.length; i++) {
    let name = aggregations[i][0];
    aggregation[name] = { //変数があるときはこの「ブラケット記法」でないと追加できない
      targetMonth: targetMonth,  //対象月
      timestamp: new Date(),  //
      openingQty: opQty[i],  //前月繰越数量
      openingAmount: opAmount[i], //前月繰越金額
      inflowQty: inflowQty[i], //当月入庫数量: 一旦固定
      inflowAmount: 0,  //当月入庫金額（仕入額）
      outflowQty: opQty[i]+inflowQty[i], //当月払出量
      outflowAmount: opAmount[i], //当月払出金額（仕掛品勘定へ）
      currentQty: 0,  //最終的には翌月繰越→まずは繰越なし
      currentStorageAmount: 0,  //最終的には翌月繰越
      // currentUnitCost: opAmount[i] / opQty[i],  //移動平均法で単価を計算
      status: "ready"
    };
  }
  Logger.log(aggregation);
  // {製品X={timestamp=Fri Jan 02 17:18:06 GMT+09:00 2026, inflowAmount=0.0, openingAmount=15000.0, outflowQty=200.0, currentStorageAmount=-15000.0, inflowQty=300.0, openingQty=50.0, targetMonth=2025-12, outflowAmount=30000.0, currentQty=150.0, status=ready}}

  // 2 入力行を走査して各製品名に割り振り（入力行ごとに各仕掛品勘定へ割り振り）→総平均法
  for (let j = 0; j < table.length; j++) {
    let name = table[j].row[3]; //暫定の位置
    let state = aggregation[name];

    if (!state) continue; //nameがなければスルー
    // let qty = Number(table[j].qty);  //現状数量の期中変化は無視
    let qty = 1; //暫定：必ず在庫増になるよう指定。
    if (qty > 0) {
      // state.inflowQty += qty; //入庫量に加算: 一旦無視
      state.inflowAmount += qty * table[j].row[2];  //入庫金額に加算 列は暫定
      // state.currentQty += qty;  //現在在庫量に加算: 一旦固定なので無視
      // state.currentStorageAmount += qty * table[j].row[2]; //在庫金額に加算　列は暫定
      // state.currentUnitCost = state.currentStorageAmount / state.currentQty;  
    } else {  //現状は払出は月末一括と考えるのでここは一旦無視してよい
      state.outflowQty += qty; //出庫量に加算(負の数として表示：会計上最後は正負逆にする)          
      state.outflowAmount += qty * state.currentUnitCost; //出庫金額に加算（同上）, 単価は現在のものと同じ
      state.currentQty += qty;  //現在在庫量から減算
      state.currentStorageAmount += qty * state.currentUnitCost; //在庫金額から減算
    }
    state.outflowAmount += state.inflowAmount
  }
  Logger.log(aggregation);
  // 	{製品X={timestamp=Fri Jan 02 17:18:06 GMT+09:00 2026, inflowAmount=38900.0, openingAmount=15000.0, outflowQty=200.0, currentStorageAmount=23900.0, inflowQty=300.0, openingQty=50.0, targetMonth=2025-12, outflowAmount=30000.0, currentQty=150.0, status=ready}}
  return aggregation;
}

// 現在列名対応 [対象月	タイムスタンプ	status	製品名	前月繰越数量	前月繰越加工進捗(仮)	前月繰越金額	当月受入数量	当月受入金額	当月払出数量	当月払出金額	翌月繰越数量	翌月繰越加工進捗(仮)	翌月繰越金額]
function wip_fillInSs_(aggregation) {
  let outcome = [];
  for (let key in aggregation) {
    let temp = aggregation[key]
    outcome.push([temp.targetMonth, temp.timestamp, temp.status, key, temp.openingQty, 0, temp.openingAmount, temp.inflowQty, temp.inflowAmount, temp.outflowQty, temp.outflowAmount, temp.currentQty, 0, temp.currentStorageAmount]);
  }
  let outputSheet = outputSheet_(wip).ss;
  outputSheet.getRange(outputSheet.getLastRow()+1, 1, outcome.length, outcome[0].length).setValues(outcome);
  Logger.log("work in progress has been aggregated.");
}


// 確認フロー
function executeWIPConfirmation () {
  return flowCnfm_(wip);
}
// 確認サブフロー：これらを順次呼ぶ
function wip_loadOutput_ (targetMonth) {
  let headerRow = outputSheet_(wip).headerRow;
  let ss = outputSheet_(wip).ss;
  let temp_table = ss.getRange(headerRow+1, 1, ss.getLastRow()-headerRow, ss.getLastColumn()).getValues();
  let table = filterByMonth_(temp_table, targetMonth);
  Logger.log(table);
  // [{idx=0.0, row=[2025-12, Mon Dec 29 23:58:40 GMT+09:00 2025, ready, 材料A, 100.0, 10000.0, 100.0, 8000.0, -100.0, -9000.0, 100.0, 9000.0]},...]
  return table;
}
function wip_applyConfirmation_ (table, targetMonth) {
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
function wip_contentToSend_ (decision) {
  let content = {};
  if (decision.type === "ok") {
    let row = decision.target.row;
    content.direct = [row[0],wip.name,Number(row[10]), Number(row[9]), row[3]]; //現時点での列位置で[対象月, 振替元, 金額, 数量, 製品名]となる
    // content.indirect = [row[0], wip.indName, row[4]]; //将来的拡張の余地
  }
  Logger.log(content);
  //	{direct=[2025-12, 直接材料, 9000.0]}
  return content;
}
function wip_refreshSs_ (decision) {
  if (decision.type === "ok") {
    let ss = outputSheet_(wip).ss;
    let headerRow = outputSheet_(wip).headerRow;
    let target = decision.target;
    let idx = Number(target.idx);
    ss.getRange(headerRow+idx+1, 3).setValue(target.row[2]);
    Logger.log("refreshing output sheet has been done.");
  } else {
    Logger.log("nothing to be refreshed.");
  }
}
function wip_sendToNext_ (content, targetMonth) {
  if (Object.keys(content).length != 0) {
    let dirNext = product;
    // let idrNext = idc;
    let dirSs = inputSheet_(dirNext).ss;
    // let idrSs = inputSheet_(idrNext).ss;
    let dc = [content.direct];
    // let ic = [content.indirect];
    dirSs.getRange(dirSs.getLastRow()+1,1,dc.length,dc[0].length).setValues(dc);
    // idrSs.getRange(idrSs.getLastRow()+1,1,ic.length,ic[0].length).setValues(ic);
    Logger.log("sending content to next step has been done.");
    alertSuccess_(targetMonth);
  } else {
    Logger.log("nothing to be sent.")
  }
}
