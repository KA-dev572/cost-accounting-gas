// 各自の環境に合わせて設定
// v0.3.0時点
function materialID() {
  //材料勘定ファイルのID
  return "YOUR_LABOR_FILE_ID";
}

function laborID() {
  //労務費勘定ファイルのID, シートID
  return {
    fileId: "FILE_ID",
    inputSheetId: ID_NUMBER,
    outputSheetId: ID_NUMBER,
    configSheetId: ID_NUMBER
  }
}

function idcId() {
  //製造間接費ファイルID
  return {
    name: "製造間接費",
    fileId:"FILE_ID",
    inputSheetId:ID_NUMBER,
    outputSheetId:ID_NUMBER
  }
}

function wipId() {
  //仕掛品勘定ファイルID
  return {
    name: "仕掛品",
    fileId: "FILE_ID",
    inputSheetId:ID_NUMBER,
    outputSheetId:ID_NUMBER
  }
}
