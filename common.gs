//共通して使用する関数。 v1 2026-01-01
// 汎用 各勘定でinputssとoutputssが共通していることが前提、そうでなくなれば個別関数へ
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
function getTargetMonth_(obj){
  let targetMonth = inputSheet_(obj).ss.getRange(1,2).getValue();  //入力箇所指定B1セル
  return targetMonth;
}
function filterByMonth_(table, targetMonth) {
  return table
    .map((row, idx) => ({ row, idx }))
    .filter(x => x.row[0] === targetMonth);
    //mapメソッドは通常、コールバック関数に要素 (row) とインデックス (idx) を提供しますが、このコードは、それらをひとまとめにした新しいデータ構造を作成します。
}

// 共通フロー指定: ctx=勘定を引数に
// 集計フロー
function flowCalc_(ctx) {
  let targetMonth = getTargetMonth_(ctx);
  let table = ctx.loadInput(targetMonth);
  let aggregation = ctx.aggregation(table, targetMonth);
  ctx.fllInSs(aggregation);
}
//　確認フロー
function flowCnfm_(ctx) {
  let targetMonth = getTargetMonth_(ctx);
  let table = ctx.loadOutput(targetMonth);
  let decision = ctx.applyConfirmation(table, targetMonth);
  let content = ctx.contentToSend(decision);
  ctx.refreshSs(decision);
  ctx.sendToNext(content, targetMonth);
  // return content;
}


// 確認作業サブフロー内共通項目（製造間接費基準：他でカスタマイズが必要ならこれは一旦製造間接費に逃して）
// 前提：タイムスタンプが[2]要素目にあること。
function pickLatestUnconfirmed_(rows) {
  //  当月に状態がconfirmedがあればconfirmedというオブジェクトを作り、それが存在すればalreadyConfirmedを返す
  const confirmed = rows.find(x => x.row[2] === "confirmed");
  if (confirmed) return { type: "alreadyConfirmed" };
  //  当月に状態がconfirmedがなければcandidatesというオブジェクトを作るが、長さ0、つまり該当行がなければnoneを返す
  const candidates = rows.filter(x => x.row[2] !== "confirmed");
  if (candidates.length === 0) return { type: "none" };
  // candidatesオブジェクトに要素が存在するならタイムスタンプ列で降順に並べ替え、先頭要素を取り出してokをラベルを貼る。
  candidates.sort((a, b) => b.row[1] - a.row[1]); // timestamp列想定
  return { type: "ok", target: candidates[0] };
}
function alertAlreadyConfirmed_ (targetMonth) {
  SpreadsheetApp.getUi().alert(`${targetMonth}分は確定済です。修正が必要な場合は管理者、次工程担当にも連絡のうえ、statusをdraftに戻して集計をやり直してください。操作方法は管理者に確認してください。`);
}
function alertNone_ (targetMonth) {
  SpreadsheetApp.getUi().alert(`${targetMonth}分の集計が済んでいません。まず原価計算→集計を行ってください。`)
}
function alertSuccess_(targetMonth) {
  SpreadsheetApp.getUi().alert(`${targetMonth}分の集計額を次工程へ送信しました。`)
}
