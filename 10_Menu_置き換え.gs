/***********************************************************************
 * 10_Menu.gs（★既存ファイルをこの内容にまるごと置き換え / 任意）
 * ---------------------------------------------------------------------
 * 変更点は1行だけ:
 *   「見積書PDF自動作成テスト」メニューを追加した（testSaveEstimate 実行用）。
 *   置き換えなくても機能は動く（テストはエディタから testSaveEstimate を
 *   実行すればよい）が、シート上から実行できて便利なので推奨。
 ***********************************************************************/

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu(APP.menuName)
    .addItem('利用開始まで自動設定', 'completeAutomaticSetup')
    .addItem('注文情報を読み込む', 'loadSelectedOrder')
    .addItem('見積書PDFを作成', 'createEstimatePdf')
    .addItem('請求書PDFを作成', 'createInvoicePdf')
    .addItem('領収書PDFを作成', 'createReceiptPdf')
    .addItem('3帳票をまとめて作成', 'createAllPdfs')
    .addItem('見積書PDF自動作成テスト', 'testSaveEstimate') // ★追加
    .addItem('入力内容をクリア', 'clearInputContents')
    .addSeparator()
    .addItem('フォームとトリガーを初期設定', 'setupSystem')
    .addItem('利用開始チェック', 'checkSystemReadiness')
    .addItem('外部公開チェック', 'showExternalAccessChecklist')
    .addItem('保存先フォルダを作成', 'setupPdfFolders')
    .addItem('エラーログを確認', 'openErrorLogSheet')
    .addToUi();
}

function openErrorLogSheet() {
  getSpreadsheet_().setActiveSheet(getSheet_(SHEETS.errorLog));
}
