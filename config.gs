// 各自の環境に合わせて設定
// v1.3時点 2026-01-09
const material = {
  //材料勘定ファイルのID
  name: "材料",
  dirName: "直接材料",
  indName: "間接材料",
  fileId: "YOUR_ID",
  inputSheetId: NUMBER,
  outputSheetId: NUMBER,
  configSheetId: NUMBER,
  logSheetId: NUMBER,
  //以下個別関数の呼び出し用
  //集計フロー用
  loadInput: material_loadInput_,  //集計シートから読み込み
  aggregation: material_aggregation_,
  fllInSs: material_fillInSs_,
  //確認フロー用
  loadOutput: material_loadOutput_,
  applyConfirmation: material_applyConfirmation_,
  contentToSend: material_contentToSend_,
  refreshSs: material_refreshSs_,
  sendToNext: material_sendToNext_  
};

const labor = {
  //労務費勘定ファイルのID, シートID
  name: "労務費",
  dirName: "直接労務費",
  indName: "間接労務費",
  fileId: "YOUR_ID",
  inputSheetId: NUMBER,
  outputSheetId: NUMBER,
  configSheetId: NUMBER,
  logSheetId: NUMBER,
  //以下個別関数の呼び出し用
  //集計フロー用
  loadInput: labar_loadInput_,  //集計シートから読み込み
  aggregation: labor_aggregation_,
  fllInSs: labor_fillInSs_,
  //確認フロー用
  loadOutput: labor_loadOutput_,
  applyConfirmation: labor_applyConfirmation_,
  contentToSend: labor_contentToSend_,
  refreshSs: labor_refreshSs_,
  sendToNext: labor_sendToNext_
};

const idc = {
  //製造間接費ファイルID  
  name: "製造間接費",
  fileId: "YOUR_ID",
  inputSheetId: NUMBER,
  outputSheetId: NUMBER,
  logSheetId: NUMBER,
  //以下個別関数の呼び出し用
  //集計フロー用
  loadInput: idc_loadInput_,  //集計シートから読み込み
  aggregation: idc_aggregation_,
  fllInSs: idc_fillInSs_,
  //確認フロー用
  loadOutput: idc_loadOutput_,
  applyConfirmation: idc_applyConfirmation_,
  contentToSend: idc_contentToSend_,
  refreshSs: idc_refreshSs_,
  sendToNext: idc_sendToNext_
};

const wip = {
  //仕掛品勘定ファイルID
    name: "仕掛品",
    fileId: "YOUR_ID",
    inputSheetId: NUMBER,
    outputSheetId: NUMBER,
    configSheetId: NUMBER,
    logSheetId: NUMBER,
    //以下個別関数の呼び出し用
    //集計フロー用
    loadInput: wip_loadInput_,  //集計シートから読み込み
    aggregation: wip_aggregation_,
    fllInSs: wip_fillInSs_,
    //確認フロー用: 製品勘定成立後
    // loadOutput: wip_loadOutput_,
    // applyConfirmation: wip_applyConfirmation_,
    // contentToSend: wip_contentToSend_,
    // refreshSs: wip_refreshSs_,
    // sendToNext: wip_sendToNext_
};

const product = {
  //製品勘定ファイル等
  name: "製品",
  fileId: "YOUR_ID",
  inputSheetId: NUMBER,
  outputSheetId: NUMBER,
  configSheetId: NUMBER,
  logSheetId: NUMBER,
  //以下個別関数の呼び出し用
  //集計フロー用
  loadInput: product_loadInput_,  //集計シートから読み込み
  aggregation: product_aggregation_,
  fllInSs: product_fillInSs_,
  //確認フロー用: 簡易P/L用
  loadOutput: product_loadOutput_,
  applyConfirmation: product_applyConfirmation_,
  contentToSend: product_contentToSend_,
  refreshSs: product_refreshSs_,
  sendToNext: product_sendToNext_
};

const pl = {
  //損益勘定ファイル等
  name: "損益計算書",
  fileId: "YOUR_ID",
  inputSheetId: NUMBER,
  outputSheetId: NUMBER,
  configSheetId: NUMBER,
  plSheetId: NUMBER,
  unitcostSheetId: NUMBER,
  logSheetId: NUMBER,
  //以下個別関数の呼び出し用
  //集計フロー用
  loadInput: pl_loadInput_,
  aggregation: pl_aggregation_,
  fllInSs: pl_fillInSs_,
  //確認フロー用:
  loadOutput: pl_loadOutput_,
  applyConfirmation: pl_applyConfirmation_,
  contentToSend: pl_contentToSend_,
  refreshSs: pl_refreshSs_,
  sendToNext: pl_sendToNext_
}
