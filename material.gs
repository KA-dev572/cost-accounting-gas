function calculateMaterialCost() {
  // ver1 材料の受払簿から月末に材料勘定の整理を行う（在庫評価方法は移動平均法）。出力は仕掛品勘定ファイルへ（とりあえずは全額直接費として扱う。製造間接費振替部分はver2以降。）
  // ファイル、シート名整理: 一旦コンテナバインド。以下設定部分は各自の環境に合わせること
  let s = SpreadsheetApp.getActiveSpreadsheet();
  let sIn = s.getActiveSheet();  //入力シート
  let sOut = s.getSheetByName("材料勘定"); //出力シート
  let sConfig = s.getSheetByName("材料名管理");

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
  let materialState = {};
  let materialName = sConfig.getRange(1,1,sConfig.getLastRow(), 1).getValues();
  // Logger.log(materialName); //材料名  [[材料A], [材料B], ...]
  // Logger.log(materialName[0][0]);  // 材料A
  // Logger.log(materialName.length);  // 1

  //当面の初期値はここで設定
  let opAQ = 100; // 前月繰越数量
  let opAA = 10000; //前月繰越金額

  //具体的取引を取得→処理→出力
  let contentRaw = sIn.getRange(sInHeaderRow+1, 1, sIn.getLastRow()-sInHeaderRow, sIn.getLastColumn()).getValues();
  // Logger.log(contentRaw);
  // Logger.log(contentRaw[4][sInQtyColumn-1]);
  // Logger.log(contentRaw.length);
  //金額欄の関数除け
  let content = [];
  for (let h=0; h<contentRaw.length-1; h++) {
    if (contentRaw[h][sInQtyColumn] != "") {
      content.push(contentRaw[h]); 
    }
  }
  // Logger.log(content);  //[[1.0, Wed Dec 03 00:00:00 GMT+09:00 2025, 材料A, 20.0, 80.0, 1600.0, ], [2.0, Fri Dec 12 00:00:00 GMT+09:00 2025, 材料A, -50.0, 30.0, -1500.0, ], [3.0, Wed Dec 24 00:00:00 GMT+09:00 2025, 材料A, -20.0, , , ]]

  let monthlyLog = [["材料名", "前月繰越数量", "前月繰越金額", "当月受入数量", "当月受入金額", "当月払出数量", "当月払出金額", "翌月繰越数量", "翌月繰越金額"]];

  for (let i=0; i<materialName.length; i++) {
    let name = materialName[i][0]
    materialState[name] = { //変数があるときはこの「ブラケット記法」でないと追加できない
      openingQty:opAQ,  //前月繰越数量
      openingAmount:opAA, //前月繰越金額
      inflowQty:0,  //当月入庫数量
      inflowAmount: 0,  //当月入庫金額（仕入額）
      outflowQty:0, //当月払出量
      outflowAmount: 0, //当月払出金額（仕掛品勘定へ）
      currentQty:opAQ,  //最終的には翌月繰越
      currentUnitCost:opAA/opAQ,  //最終的には翌月繰越
      transactions: []  //取引履歴
    }
    let storageAmount = materialState[name].currentQty * materialState[name].currentUnitCost
    // Logger.log(materialState);  //{材料A={openingQty=100.0, currentUnitCost=100.0, transactions=[], openingAmount=10000.0, currentQty=100.0}}：初期条件
    for (let j=0; j<content.length; j++) {
      // Logger.log(content[j][sInNameClm])

      if (name == content[j][sInNameClm]) {
        //入庫の場合
        if (Number(content[j][sInQtyColumn])>0) {
          materialState[name].transactions.push({
            date:new Date(content[j][sInDateColumn]),
            qty:Number(content[j][sInQtyColumn]),
            unitcost:Number(content[j][sInUCColumn])
          })
          materialState[name].inflowQty = materialState[name].inflowQty + content[j][sInQtyColumn]; //入庫量に加算
          materialState[name].inflowAmount = materialState[name].inflowAmount + content[j][sInQtyColumn] * content[j][sInUCColumn];  //入庫金額に加算
          materialState[name].currentQty = materialState[name].currentQty + content[j][sInQtyColumn];  //現在在庫量に加算
          storageAmount = storageAmount + content[j][sInQtyColumn] * content[j][sInUCColumn]; //在庫金額に加算
          materialState[name].currentUnitCost = storageAmount / materialState[name].currentQty;  //現在在庫単価（移動平均法）
        } else {  //出庫の場合
          materialState[name].transactions.push({
            date:new Date(content[j][sInDateColumn]),
            qty:Number(content[j][sInQtyColumn])
          })
          materialState[name].outflowQty = materialState[name].outflowQty + content[j][sInQtyColumn]; //出庫量に加算(負の数として表示：会計上最後は正負逆にする)          
          materialState[name].outflowAmount = materialState[name].outflowAmount + content[j][sInQtyColumn] * materialState[name].currentUnitCost; //出庫金額に加算（同上）, 単価は現在のものと同じ
          materialState[name].currentQty = materialState[name].currentQty + content[j][sInQtyColumn];  //現在在庫量から減算
          storageAmount = storageAmount + content[j][sInQtyColumn] * materialState[name].currentUnitCost; //在庫金額から減算
        }
      }
    }
    // Logger.log(materialState);  //{材料A={outflowAmount=-1933.3333333333335, openingAmount=10000.0, outflowQty=-70.0, currentQty=120.0, openingQty=100.0, inflowAmount=1600.0, inflowQty=20.0, transactions=[{date=Wed Dec 03 00:00:00 GMT+09:00 2025, qty=20.0, unitcost=NaN}, {date=Fri Dec 12 00:00:00 GMT+09:00 2025, qty=-50.0}, {date=Wed Dec 24 00:00:00 GMT+09:00 2025, qty=-20.0}], currentUnitCost=96.66666666666667}}
    monthlyLog.push([name, materialState[name].openingQty, materialState[name].openingAmount, materialState[name].inflowQty, materialState[name].inflowAmount, materialState[name].outflowQty, materialState[name].outflowAmount, materialState[name].currentQty, storageAmount]);
  }
  Logger.log(monthlyLog); //[[材料名, 前月繰越数量, 前月繰越金額, 当月受入数量, 当月受入金額, 当月払出数量, 当月払出金額, 翌月繰越数量, 翌月繰越金額], [材料A, 100.0, 10000.0, 100.0, 8000.0, -100.0, -9000.0, 100.0, 16500.0]]

  sOut.getRange(1,1,monthlyLog.length, monthlyLog[0].length).setValues(monthlyLog);

}
