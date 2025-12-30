function executeLaborCalc() {  //ライブラリ呼び出し用ラッパー
  return calculateLaborCost_();
}

function calculateLaborCost_() {
  // ver1 確定勤怠簿から月末に労務費勘定の整理を行う。出力は一旦労務費集計シートへ→直接/間接まで実装。2財以上製造の場合はver2以降
  // 
  // ファイル、シート名整理→ここもconfigで整理すればいいのでは？ややこしい？
  let ss = SpreadsheetApp.openById(laborID().fileId);
  let sIn = ss.getSheetById(laborID().inputSheetId);  //入力シート
  let sOut = ss.getSheetById(laborID().outputSheetId); //出力シート
  // let sConfig = ss.getSheetByName("入力管理");
  // Logger.log([ss, sIn, sOut]);

  //  入力シートの整理
  let sInHeaderRow = 2;
  let sInHeader = sIn.getRange(sInHeaderRow, 1, sInHeaderRow, sIn.getLastColumn()).getValues()[0];
  // Logger.log(sInHeader);  //[日付, 職員氏名, 作業時間, 作業内容, 時給, 金額]: 勤怠簿シートに職員とひもづいた時給で日額を計算できる想定
  let sInDateClm = sInHeader.indexOf("日付");
  let sInQtyClm = sInHeader.indexOf("作業時間");
  let sInKindClm = sInHeader.indexOf("作業内容");
  let sInCostClm = sInHeader.indexOf("金額");
  // Logger.log([sInDateClm, sInQtyColumn, sInKindClm, sInCostClm]); //[0, 2, 3, 5]

  // 出力シート管理
  let sOutLastRow = sOut.getLastRow();

  //勤怠を取得→処理→出力
  let contentRaw = sIn.getRange(sInHeaderRow+1, 1, sIn.getLastRow()-sInHeaderRow, sIn.getLastColumn()).getValues();
  // Logger.log(contentRaw);
  // Logger.log(contentRaw[4][sInQtyClm]);
  // Logger.log(contentRaw.length);
  //金額欄の関数除け
  let content = [];
  for (let h=0; h<contentRaw.length; h++) {
    if (contentRaw[h][sInQtyClm] != "") {
      content.push(contentRaw[h]); 
    }
  }
  // Logger.log(content);
  //実験結果[[Wed Dec 03 00:00:00 GMT+09:00 2025, 甲, 7.5, 製造, 1100.0, 8250.0], [Fri Dec 12 00:00:00 GMT+09:00 2025, 甲, 6.0, 製造, 1100.0, 6600.0], [Fri Dec 12 00:00:00 GMT+09:00 2025, 甲, 1.5, 手待, 1100.0, 1650.0], [Wed Dec 24 00:00:00 GMT+09:00 2025, 丙, 6.0, 製造, 900.0, 5400.0], [Thu Dec 25 00:00:00 GMT+09:00 2025, 乙, 8.0, 製造, 1000.0, 8000.0]]

  // ver0.2.0 当面1財の生産を考えるので、直接費/間接費のみを割り振る→とりあえず製造/手待で判別
  //2財以上になれば製品1製造, 製造2製造,... 製造間接費という形に拡張する予定
  //処理本番：1行ごとに走査して、製造なら直接経費、手待ちなら製造間接費へ加算

  let monthlyLog = [];

  // 1 初期状態を作る（0,0）
  let laborCostState = {
    targetMonth:sIn.getRange(1,2).getValue(),
    timestamp: new Date(),
    status: "ready",
    directCost: 0,
    indirectCost: 0
  };

  // 2 入力行を走査して直接/間接に割り振り（入力行ごと）
  for (let i = 0; i < content.length; i++) {
    let kind = content[i][sInKindClm];  //作業内容
    let cost = Number(content[i][sInCostClm]);  //入力シート側で数値以外を弾くよう設計済み
    if (kind == "製造") { //変数化してconfigで保持してもよい
      laborCostState.directCost += cost;  //直接労務費に加算
    } else if (kind == "手待") {
      laborCostState.indirectCost += cost;  //間接労務費に加算
    }
  }
  // Logger.log(laborCostState);  //	{directCost=28250.0, indirectCost=1650.0}

  //3 出力
  monthlyLog.push([laborCostState.targetMonth, laborCostState.timestamp, laborCostState.status, laborCostState.directCost, laborCostState.indirectCost]);
  // Logger.log(monthlyLog); //結果	[[当月直接労務費, 当月製造間接費振替], [28250.0, 1650.0]]

  sOut.getRange(sOutLastRow+1,1,monthlyLog.length, monthlyLog[0].length).setValues(monthlyLog);
}

//以降、確認用関数。シートから呼び出せるよう設定
function executeLCComfirmation () {
  const ss = SpreadsheetApp.openById(laborID().fileId).getSheetById(laborID().inputSheetId);
  const targetMonth = ss.getRange(1,2).getValue();
  return confirmLaborCost_(targetMonth);
}

function confirmLaborCost_ (targetMonth) {
  const ss = SpreadsheetApp.openById(laborID().fileId).getSheetById(laborID().outputSheetId);
  let aggregationTable = ss.getRange(2, 1, ss.getLastRow(), ss.getLastColumn()).getValues();
  // Logger.log(aggregationTable);
  const rows = filterByMonth_ (aggregationTable, targetMonth);
  // Logger.log(rows); //[{row=[2025-12, Tue Dec 30 00:14:29 GMT+09:00 2025, ready, 28250.0, 1650.0], idx=1.0}, {row=[2025-12, Tue Dec 30 02:34:19 GMT+09:00 2025, confirmed, 28250.0, 1650.0], idx=2.0}]
  const decision = pickLatestUnconfirmed_(rows);
  // Logger.log(decision); //{type=alreadyConfirmed}
  if (decision.type === "alreadyConfirmed") {
    alertAlreadyConfirmed_(targetMonth);
    return;
  } else if (decision.type === "none") {
    alertNone_ (targetMonth);
  } else if (decision.type === "ok") {
    confirmAndSend_(aggregationTable, decision);
    saveAggregationTable_(laborID(),aggregationTable);  //製造間接費から流用
    alertSuccess_ (targetMonth);  //製造間接費から流用
  }
}

function filterByMonth_(table, targetMonth) {
  return table
    .map((row, idx) => ({ row, idx }))
    .filter(x => x.row[0] === targetMonth);

    //mapメソッドは通常、コールバック関数に要素 (row) とインデックス (idx) を提供しますが、このコードは、それらをひとまとめにした新しいデータ構造を作成します。
}

function pickLatestUnconfirmed_(rows) {
  //  当月に状態がconfirmedがあればconfirmedというオブジェクトを作り、それが存在すればalreadyConfirmedを返す
  const confirmed = rows.find(x => x.row[2] === "confirmed");
  if (confirmed) return { type: "alreadyConfirmed" };
  //  当月に状態がconfirmedがなければcandidatesというオブジェクトを作るが、長さ0、つまり該当行がなければnoneを返す
  const candidates = rows.filter(x => x.row[2] !== "confirmed");
  if (candidates.length === 0) return { type: "none" };
  // candidatesオブジェクトに要素が存在するならタイムスタンプ列で降順に並べ替え、先頭要素を取り出してokをラベルを貼る。
  candidates.sort((a, b) => b.row[3] - a.row[3]); // timestamp列想定
  return { type: "ok", target: candidates[0] };
}

function alertAlreadyConfirmed_ (targetMonth) {
  SpreadsheetApp.getUi().alert(`${targetMonth}分は確定済です。修正が必要な場合は管理者、製造間接費担当及び仕掛品担当にも連絡のうえ、statusをdraftに戻して集計をやり直してください。操作方法は管理者に確認してください。`);
}

function alertNone_ (targetMonth) {
  SpreadsheetApp.getUi().alert(`${targetMonth}分の集計が済んでいません。まず原価計算→集計を行ってください。`)
}

function confirmAndSend_(table, targetInfo) {
  const idx = targetInfo.idx;
  table[idx][2] = "confirmed";

  const contentToSend = [[
    table[idx][0],
    "間接労務費",
    table[idx][4]
  ]];

  const ss1 = SpreadsheetApp
    .openById(idcId().fileId)
    .getSheetById(idcId().inputSheetId);

  ss1.getRange(ss1.getLastRow()+1,1,1,3).setValues(contentToSend);
}
