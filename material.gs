//ver 0.1.2 ライブラリ呼び出し用ラッパー: コードが分散しないよう一元管理してUI部分のみコンテナバインドに薄く実装
function executeMaterialCalc() {
  return calculateMaterialCost_();
}

function calculateMaterialCost_() {
  // ver1 材料の受払簿から月末に材料勘定の整理を行う（在庫評価方法は移動平均法）。出力は仕掛品勘定ファイルへ（とりあえずは全額直接費として扱う。製造間接費振替部分はver2以降。）
  // ファイル、シート名整理
  let ss = SpreadsheetApp.openById(materialID());  //configファイルでIDを管理：各自の環境に合わせる
  let sIn = ss.getSheetByName("材料入力");  //入力シート
  let sOut = ss.getSheetByName("材料勘定"); //出力シート
  let sConfig = ss.getSheetByName("材料名管理");  //材料名の管理用シート。入力用シートの入力規則に使用

  //  入力シートの整理
  let sInHeaderRow = 1;
  let sInHeader = sIn.getRange(sInHeaderRow, 1, sInHeaderRow, sIn.getLastColumn()).getValues()[0];
  // Logger.log(sHeader);  //[取引番号, 日付, 材料名, 数量, 単価, 金額, 摘要]
  let sInNameClm = sInHeader.indexOf("材料名");
  let sInDateColumn = sInHeader.indexOf("日付");
  let sInQtyColumn = sInHeader.indexOf("数量");
  let sInUCColumn = sInHeader.indexOf("単価");
  // Logger.log([sInNameClm, sInDateColumn, sInQtyColumn, sInUCColumn]);

  //材料名取得（入力にもここからの入力規則を導入済み）
  let materialNames = sConfig.getRange(1,1,sConfig.getLastRow(), 1).getValues();
  // Logger.log(materialName); //材料名  [[材料A], [材料B], ...]
  // Logger.log(materialName[0][0]);  // 材料A
  // Logger.log(materialName.length);  // 1

  //当面の初期値はここで設定
  let opAQ = [100]; // 前月繰越数量
  let opAA = [10000]; //前月繰越金額

  //具体的取引を取得→処理→出力
  let contentRaw = sIn.getRange(sInHeaderRow+1, 1, sIn.getLastRow()-sInHeaderRow, sIn.getLastColumn()).getValues();
  // Logger.log(contentRaw);
  // Logger.log(contentRaw[4][sInQtyColumn-1]);
  // Logger.log(contentRaw.length);
  //金額欄の関数除け
  let content = [];
  for (let h=0; h<contentRaw.length; h++) {
    if (contentRaw[h][sInQtyColumn] != "") {
      content.push(contentRaw[h]); 
    }
  }
  // Logger.log(content);  //[[1.0, Wed Dec 03 00:00:00 GMT+09:00 2025, 材料A, 20.0, 80.0, 1600.0, ], [2.0, Fri Dec 12 00:00:00 GMT+09:00 2025, 材料A, -50.0, 30.0, -1500.0, ], [3.0, Wed Dec 24 00:00:00 GMT+09:00 2025, 材料A, -20.0, , , ]]

  let monthlyLog = [["材料名", "前月繰越数量", "前月繰越金額", "当月受入数量", "当月受入金額", "当月払出数量", "当月払出金額", "翌月繰越数量", "翌月繰越金額"]];

  //ver0.1.1: ver0.1に対してリファクタリング：materialState[name]ごとに全入力を走査していたが、以下のとおり

  // 1 材料ごとの初期状態を作る（各材料でT勘定の開始仕訳をする）
  let materialState = {};
  for (let i = 0; i < materialNames.length; i++) {
    let name = materialNames[i][0];
    materialState[name] = { //変数があるときはこの「ブラケット記法」でないと追加できない
      openingQty: opAQ[i],  //前月繰越数量
      openingAmount: opAA[i], //前月繰越金額
      inflowQty: 0, //当月入庫数量
      inflowAmount: 0,  //当月入庫金額（仕入額）
      outflowQty: 0, //当月払出量
      outflowAmount: 0, //当月払出金額（仕掛品勘定へ）
      currentQty: opAQ[i],  //最終的には翌月繰越
      currentStorageAmount: opAA[i],  //最終的には翌月繰越
      currentUnitCost: opAA[i] / opAQ[i]  //移動平均法で単価を計算
    };
  }

  // 2 入力行を走査して各材料に割り振り（入力行ごとに各材料勘定へ割り振り）
  for (let j = 0; j < content.length; j++) {
    let name = content[j][sInNameClm];
    let state = materialState[name];

    if (!state) continue; //nameがなければスルー
    let qty = Number(content[j][sInQtyColumn]);
    if (qty > 0) {
      state.inflowQty += qty; //入庫量に加算
      state.inflowAmount += qty * content[j][sInUCColumn];  //入庫金額に加算
      state.currentQty += qty;  //現在在庫量に加算
      state.currentStorageAmount += qty * content[j][sInUCColumn]; //在庫金額に加算
      state.currentUnitCost = state.currentStorageAmount / state.currentQty;  //現在在庫単価（移動平均法）
    } else {
      state.outflowQty += qty; //出庫量に加算(負の数として表示：会計上最後は正負逆にする)          
      state.outflowAmount += qty * state.currentUnitCost; //出庫金額に加算（同上）, 単価は現在のものと同じ
      state.currentQty += qty;  //現在在庫量から減算
      state.currentStorageAmount += qty * state.currentUnitCost; //在庫金額から減算
    }
  }

  Logger.log(materialState);  //{材料A={currentUnitCost=90.0, inflowQty=100.0, currentQty=100.0, openingAmount=10000.0, inflowAmount=8000.0, openingQty=100.0, outflowAmount=-9000.0, outflowQty=-100.0}}

  // monthlyLog.push([name, state.openingQty, state.openingAmount, state.inflowQty, state.inflowAmount, state.outflowQty, state.outflowAmount, state.currentQty, storageAmount]);

  // Logger.log(monthlyLog); //[[材料名, 前月繰越数量, 前月繰越金額, 当月受入数量, 当月受入金額, 当月払出数量, 当月払出金額, 翌月繰越数量, 翌月繰越金額], [材料A, 100.0, 10000.0, 100.0, 8000.0, -100.0, -9000.0, 100.0, 16500.0]]
  //出力
  for (let key in materialState) {
    let keyState = materialState[key];
    monthlyLog.push([key, keyState.openingQty, keyState.openingAmount, keyState.inflowQty, keyState.inflowAmount, keyState.outflowQty, keyState.outflowAmount, keyState.currentQty, keyState.currentStorageAmount]);
  }
  sOut.getRange(1,1,monthlyLog.length, monthlyLog[0].length).setValues(monthlyLog);
}
