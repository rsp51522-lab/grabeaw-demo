/***********************************************************************
 * 12_AutoEstimate.gs（★新規追加ファイル）
 * ---------------------------------------------------------------------
 * 目的:
 *   Googleフォーム送信をきっかけに、以下を全自動で実行する。
 *
 *   フォーム送信
 *     → フォーム回答へ追加（既存機能）
 *     → 注文管理・注文明細・顧客管理へ登録（既存機能）
 *     → 帳票入力シートへデータ反映        …… 本ファイル(1)
 *     → 御見積書シートへ自動転記          …… 本ファイル(2)
 *     → 御見積書シートのみPDF化           …… 本ファイル(3)
 *     → Googleドライブへ保存（同名はゴミ箱へ）… 本ファイル(4)
 *     → PDFのURLをフォーム回答・顧客管理・注文管理へ書き込み … 本ファイル(5)
 *     → 顧客へ見積書PDFをメール送信（テストモード付き）……… 本ファイル(6)
 *     → 管理者へ「〇〇様の見積書を保存しました」通知 ………… 本ファイル(7)
 *     → エラー時: エラーログへ記録＋フォーム回答へ「PDF保存エラー」
 *
 * 設計方針:
 *   ・既存コードは一切変更しない（既存関数を呼び出すだけ）
 *     利用する既存関数:
 *       loadOrderToInputSheet_ / assertTemplateConfiguration_ /
 *       readInputSheetData_ / validateDocumentInput_ / writeTemplateSheet_ /
 *       buildSheetExportUrl_ / updateOrderPdfLink_ / getOrderByNumber_ /
 *       getSettingValue_ / getSettingsMap_ / getSheet_ / getSpreadsheet_ /
 *       getHeaderMap_ / getColumnIndexByHeader_ / findRowByValue_ /
 *       sanitizeFileName_ / buildAddressee_ / inferCustomerType_ /
 *       isValidEmail_ / formatYen_ / clearInputContents / logError_ /
 *       showMessage_ / safeGet_ / getInputOrderNumber_
 *   ・PDF作成が失敗しても「注文登録」は成功扱い（受注データを守る）
 *   ・呼び出し元は「無題 2.gs」の processFormSubmit_ に追加した1行のみ
 ***********************************************************************/

/** 設定シートに追加するキー（初回実行時に自動で行が追加される） */
const AUTO_ESTIMATE_SETTING_KEYS = {
  mailMode: 'メール送信モード',     // 「本番」= 顧客へ送信 / それ以外 = 管理者へテスト送信
  customerMail: '顧客メール送信',   // 「ON」= メール機能を使う / 「OFF」= 使わない
  mailSubject: '顧客メール件名',    // 設定シートで自由に修正できる（空欄なら既定の件名）
  mailBody: '顧客メール本文',       // 設定シートで自由に修正できる（{宛名} が宛名に置換される）
};

/** フォーム回答・顧客管理へ追加する列の見出し */
const AUTO_ESTIMATE_HEADER = '見積書PDF';

/** PDF処理失敗時にフォーム回答へ書き込む文言 */
const AUTO_ESTIMATE_ERROR_TEXT = 'PDF保存エラー';

/**
 * 顧客向けメールの既定文面（{宛名} は宛名に置換される）。
 * ※実際の文面は設定シートの「顧客メール件名」「顧客メール本文」が優先される。
 *   設定シートを空欄にするとこの既定文面に戻る。
 */
const AUTO_ESTIMATE_MAIL_TEXT = {
  customerSubject: '御見積書を送付いたします',
  customerBody: [
    '{宛名}',
    '',
    'この度はお問い合わせありがとうございます。',
    '',
    '御見積書を送付いたします。',
    '',
    'ご確認よろしくお願いいたします。',
    '',
    '添付',
    '御見積書PDF',
  ].join('\n'),
};

/***********************************************************************
 * エントリーポイント
 ***********************************************************************/

/**
 * フォーム送信フローから呼ばれる安全版。
 * PDF処理が失敗しても例外を外へ投げない（注文登録を守るため）。
 * 失敗時はエラーログへ記録し、フォーム回答へ「PDF保存エラー」を書き込む。
 * @param {string} orderNumber 注文番号（例: GR-20260714-001）
 * @return {Object|null} 成功時 { orderNumber, fileName, url, fileId, mail } / 失敗時 null
 */
function autoCreateEstimateSafely_(orderNumber) {
  try {
    return autoCreateEstimateForOrder_(orderNumber);
  } catch (error) {
    // エラーログシートへ記録（管理者メール設定済みなら通知も飛ぶ: 既存 logError_）
    safeGet_(function() { logError_('見積書PDF自動作成', orderNumber, error); return true; }, false);
    // フォーム回答へ「PDF保存エラー」を記録
    safeGet_(function() { writeFormResponsePdfError_(orderNumber); return true; }, false);
    return null;
  }
}

/**
 * 見積書PDF自動作成の本体。手動テスト用 testSaveEstimate() からも呼ばれる。
 * 失敗時は例外を投げる（呼び出し側で処理を分ける）。
 * @param {string} orderNumber 注文番号
 * @return {Object} { orderNumber, fileName, url, fileId, mail }
 */
function autoCreateEstimateForOrder_(orderNumber) {
  if (!orderNumber) {
    throw new Error('注文番号が指定されていません。');
  }

  // 設定シートに「メール送信モード」等が無ければ自動追加
  ensureAutoEstimateSettings_();

  // 注文データを取得（存在しなければここで例外）
  const order = getOrderByNumber_(orderNumber);

  // (1) 帳票入力シートへデータ反映（既存機能をそのまま利用）
  loadOrderToInputSheet_(orderNumber);

  // (2) 御見積書シートへ自動転記（既存機能をそのまま利用）
  assertTemplateConfiguration_();
  const inputData = readInputSheetData_();
  validateDocumentInput_(inputData);
  writeTemplateSheet_('estimate', inputData, orderNumber);
  SpreadsheetApp.flush(); // 転記を確定させてからPDF化する（重要）

  // (3) 御見積書シートのみPDF化
  //     既存 buildSheetExportUrl_ が A4縦・余白小(0.25インチ)・グリッド線なし・
  //     シート名なし・ページ番号なしを指定済み
  const ss = getSpreadsheet_();
  const estimateSheet = getSheet_(SHEETS.estimate);
  const exportUrl = buildSheetExportUrl_(ss.getId(), estimateSheet.getSheetId());
  const response = UrlFetchApp.fetch(exportUrl, {
    headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() },
    muteHttpExceptions: true,
  });
  if (response.getResponseCode() !== 200) {
    throw new Error('PDFエクスポートに失敗しました。HTTP ' + response.getResponseCode());
  }
  const blob = response.getBlob();

  // (4) Googleドライブへ保存（同名ファイルはゴミ箱へ移動してから保存）
  const fileName = buildEstimateAutoFileName_(order.data, orderNumber);
  const folder = getEstimateSaveFolder_();
  trashSameNameFiles_(folder, fileName);
  const file = folder.createFile(blob.setName(fileName));

  // (5) PDFのURLを各シートへ書き込み
  updateOrderPdfLink_(orderNumber, 'estimate', file);            // 注文管理（既存機能）
  writeEstimateUrlToFormResponses_(orderNumber, file.getUrl());  // フォーム回答
  writeEstimateUrlToCustomers_(String(order.data['顧客番号'] || ''), file.getUrl()); // 顧客管理

  // (6) 顧客へメール送信（テストモード付き）
  const mailResult = sendEstimateMailToCustomer_(order.data, file);

  // (7) 管理者へ保存完了通知
  sendEstimateAdminNotice_(order.data, orderNumber, file, mailResult);

  // 後片付け: 帳票入力シートをクリア（既存の手動フローと同じ挙動）
  clearInputContents();

  return {
    orderNumber: orderNumber,
    fileName: file.getName(),
    url: file.getUrl(),
    fileId: file.getId(),
    mail: mailResult,
  };
}

/**
 * 手動テスト用関数。Apps Scriptエディタから実行する。
 * 帳票入力シートの注文番号セルに値があればその注文、
 * 無ければ注文管理シートの最新注文でPDF作成を実行する。
 * @return {Object} autoCreateEstimateForOrder_ の戻り値
 */
function testSaveEstimate() {
  ensureAutoEstimateSettings_();

  const inputCellOrder = safeGet_(function() { return getInputOrderNumber_(); }, '');
  const orderNumber = inputCellOrder || getLatestOrderNumber_();
  if (!orderNumber) {
    throw new Error('テスト対象の注文がありません。先にフォームからテスト送信するか、帳票入力シートの注文番号セルに注文番号を入力してください。');
  }

  const result = autoCreateEstimateForOrder_(orderNumber);
  Logger.log('テスト成功: ' + result.fileName);
  Logger.log('PDF URL: ' + result.url);
  Logger.log('メール: ' + JSON.stringify(result.mail));
  showMessage_('見積書PDF自動作成テスト', '保存しました。\n' + result.fileName + '\n' + result.url);
  return result;
}

/***********************************************************************
 * ファイル名・保存先
 ***********************************************************************/

/**
 * 保存ファイル名を組み立てる。
 *   会社名あり: 株式会社〇〇御中_御見積書_注文番号.pdf
 *   会社名なし: 浅野太郎様_御見積書_注文番号.pdf
 * ファイル名に使えない文字は既存 sanitizeFileName_ が自動置換する。
 * @param {Object} orderData 注文管理シートの1行分（ヘッダー名→値）
 * @param {string} orderNumber 注文番号
 * @return {string} PDFファイル名
 */
function buildEstimateAutoFileName_(orderData, orderNumber) {
  const company = String(orderData['会社名・店舗名'] || '').trim();
  const person = String(orderData['お客様名'] || '').trim();
  const base = company
    ? company + '御中_御見積書_' + orderNumber
    : person + '様_御見積書_' + orderNumber;
  return sanitizeFileName_(base) + '.pdf';
}

/**
 * 保存先フォルダを取得する。
 * 設定シートの「見積書保存先フォルダID」を優先し、
 * 未設定なら「PDF保存先フォルダID」を使う。
 * @return {Folder} Googleドライブのフォルダ
 */
function getEstimateSaveFolder_() {
  const folderId = String(
    getSettingValue_(SETTINGS.estimateFolderId, '') ||
    getSettingValue_(SETTINGS.pdfRootFolderId, '') ||
    ''
  ).trim();
  if (!folderId) {
    throw new Error('設定シートの「見積書保存先フォルダID」が未設定です。setupPdfFolders() を実行してください。');
  }
  try {
    return DriveApp.getFolderById(folderId);
  } catch (error) {
    throw new Error('保存先フォルダを開けません。設定シートのフォルダIDを確認してください: ' + folderId);
  }
}

/**
 * フォルダ内の同名ファイルをゴミ箱へ移動する（上書き保存の代わり）。
 * ゴミ箱のファイルは30日間は復元できるため、誤削除にも対応できる。
 * @param {Folder} folder 対象フォルダ
 * @param {string} fileName ファイル名
 */
function trashSameNameFiles_(folder, fileName) {
  const files = folder.getFilesByName(fileName);
  while (files.hasNext()) {
    files.next().setTrashed(true);
  }
}

/***********************************************************************
 * URL書き込み
 ***********************************************************************/

/**
 * フォーム回答シートの該当行に見積書PDFのURLを書き込む。
 * 「見積書PDF」列が無ければ自動で追加する。
 * @param {string} orderNumber 注文番号
 * @param {string} url PDFのURL
 * @return {boolean} 書き込めたら true
 */
function writeEstimateUrlToFormResponses_(orderNumber, url) {
  const sheet = getSheet_(SHEETS.formResponses);
  const row = findFormResponseRowByOrderNumber_(sheet, orderNumber);
  if (!row) return false;
  const col = ensureColumnByHeader_(sheet, AUTO_ESTIMATE_HEADER);
  sheet.getRange(row, col).setValue(url);
  return true;
}

/**
 * フォーム回答シートの該当行に「PDF保存エラー」を書き込む。
 * エラー内容列と見積書PDF列の両方に記録する。
 * @param {string} orderNumber 注文番号
 * @return {boolean} 書き込めたら true
 */
function writeFormResponsePdfError_(orderNumber) {
  const sheet = getSheet_(SHEETS.formResponses);
  const row = findFormResponseRowByOrderNumber_(sheet, orderNumber);
  if (!row) return false;
  const headerMap = getHeaderMap_(sheet);
  if (headerMap['エラー内容']) {
    sheet.getRange(row, headerMap['エラー内容']).setValue(AUTO_ESTIMATE_ERROR_TEXT);
  }
  const col = ensureColumnByHeader_(sheet, AUTO_ESTIMATE_HEADER);
  sheet.getRange(row, col).setValue(AUTO_ESTIMATE_ERROR_TEXT);
  return true;
}

/**
 * 顧客管理シートの該当顧客に見積書PDFのURLを書き込む。
 * 「見積書PDF」列が無ければ自動で追加し、更新日時も更新する。
 * @param {string} customerNumber 顧客番号（例: C-000001）
 * @param {string} url PDFのURL
 * @return {boolean} 書き込めたら true
 */
function writeEstimateUrlToCustomers_(customerNumber, url) {
  if (!customerNumber) return false;
  const sheet = getSheet_(SHEETS.customers);
  const finder = findRowByValue_(sheet, 1, customerNumber);
  if (!finder) return false;
  const col = ensureColumnByHeader_(sheet, AUTO_ESTIMATE_HEADER);
  sheet.getRange(finder.row, col).setValue(url);
  const updatedCol = getColumnIndexByHeader_(sheet, '更新日時');
  if (updatedCol > 0) {
    sheet.getRange(finder.row, updatedCol).setValue(new Date());
  }
  return true;
}

/**
 * フォーム回答シートから注文番号で行を探す（下の行=新しい回答から検索）。
 * @param {Sheet} sheet フォーム回答シート
 * @param {string} orderNumber 注文番号
 * @return {number} 行番号（見つからなければ 0）
 */
function findFormResponseRowByOrderNumber_(sheet, orderNumber) {
  const headerMap = getHeaderMap_(sheet);
  const col = headerMap['注文番号'];
  if (!col || sheet.getLastRow() <= 1) return 0;
  const values = sheet.getRange(2, col, sheet.getLastRow() - 1, 1).getValues();
  for (let i = values.length - 1; i >= 0; i -= 1) {
    if (String(values[i][0] || '').trim() === String(orderNumber).trim()) {
      return i + 2;
    }
  }
  return 0;
}

/**
 * シートに指定見出しの列が無ければ末尾に追加し、列番号を返す。
 * @param {Sheet} sheet 対象シート
 * @param {string} headerName 見出し名
 * @return {number} 列番号（1始まり）
 */
function ensureColumnByHeader_(sheet, headerName) {
  const headerMap = getHeaderMap_(sheet);
  if (headerMap[headerName]) return headerMap[headerName];
  const col = sheet.getLastColumn() + 1;
  sheet.getRange(1, col).setValue(headerName).setFontWeight('bold');
  return col;
}

/***********************************************************************
 * メール送信
 ***********************************************************************/

/**
 * 顧客へ見積書PDF付きメールを送る。
 * 設定シート「メール送信モード」が「本番」のときだけ顧客へ送信し、
 * それ以外（テスト）のときは管理者宛に内容確認用メールを送る。
 * 「顧客メール送信」が「OFF」なら何も送らない。
 * @param {Object} orderData 注文管理シートの1行分
 * @param {File} file 保存済みPDFファイル
 * @return {Object} { sent, mode, to, reason }
 */
function sendEstimateMailToCustomer_(orderData, file) {
  const enabled = String(getSettingValue_(AUTO_ESTIMATE_SETTING_KEYS.customerMail, 'ON')).trim().toUpperCase();
  if (enabled === 'OFF') {
    return { sent: false, mode: 'OFF', to: '', reason: '顧客メール送信がOFF' };
  }

  const mode = String(getSettingValue_(AUTO_ESTIMATE_SETTING_KEYS.mailMode, 'テスト')).trim();
  const adminEmail = String(getSettingValue_(SETTINGS.adminEmail, '')).trim();
  const customerEmail = String(orderData['メールアドレス'] || '').trim();

  // 宛名: 会社名あり→「株式会社〇〇 御中」/ なし→「浅野太郎 様」（既存 buildAddressee_）
  const addressee = buildAddressee_(
    orderData['会社名・店舗名'],
    orderData['お客様名'],
    inferCustomerType_(orderData)
  );

  // 件名・本文は設定シートの値を優先（空欄なら既定文面）。{宛名} を実際の宛名に置換。
  const subject = String(
    getSettingValue_(AUTO_ESTIMATE_SETTING_KEYS.mailSubject, '') || AUTO_ESTIMATE_MAIL_TEXT.customerSubject
  ).trim();
  const template = String(
    getSettingValue_(AUTO_ESTIMATE_SETTING_KEYS.mailBody, '') || AUTO_ESTIMATE_MAIL_TEXT.customerBody
  );
  const body = template.replace(/\{宛名\}/g, addressee).replace(/\{addressee\}/g, addressee);

  const options = { attachments: [file.getBlob()] };
  const senderName = String(getSettingValue_(SETTINGS.companyName, '')).trim();
  if (senderName) options.name = senderName; // 送信者名（設定シートの会社名）

  if (mode === '本番') {
    if (!customerEmail || !isValidEmail_(customerEmail)) {
      throw new Error('顧客メールアドレスが不正のため送信できません: ' + (customerEmail || '(空欄)'));
    }
    MailApp.sendEmail(customerEmail, subject, body, options);
    return { sent: true, mode: '本番', to: customerEmail };
  }

  // テストモード: 管理者宛に送信して文面・添付を確認する
  if (!adminEmail) {
    return { sent: false, mode: 'テスト', to: '', reason: '管理者メールアドレスが未設定' };
  }
  const testBody = [
    '※これはテストモードの確認用メールです（顧客へは送信されていません）。',
    '本来の宛先: ' + (customerEmail || '（メールアドレス未入力）'),
    '設定シートの「メール送信モード」を「本番」にすると顧客へ直接送信されます。',
    '件名・本文は設定シートの「顧客メール件名」「顧客メール本文」で自由に修正できます。',
    '',
    '----- 以下、顧客へ送信される内容 -----',
    '',
    body,
  ].join('\n');
  MailApp.sendEmail(adminEmail, '【テスト送信】' + subject, testBody, options);
  return { sent: true, mode: 'テスト', to: adminEmail };
}

/**
 * 管理者へ「〇〇様の見積書を保存しました」通知を送る。
 * 管理者メールアドレスが未設定なら何もしない。
 * @param {Object} orderData 注文管理シートの1行分
 * @param {string} orderNumber 注文番号
 * @param {File} file 保存済みPDF
 * @param {Object} mailResult 顧客メールの送信結果
 * @return {boolean} 送信したら true
 */
function sendEstimateAdminNotice_(orderData, orderNumber, file, mailResult) {
  const adminEmail = String(getSettingValue_(SETTINGS.adminEmail, '')).trim();
  if (!adminEmail) return false;

  const company = String(orderData['会社名・店舗名'] || '').trim();
  const person = String(orderData['お客様名'] || '').trim();
  const displayName = company || person;

  const subject = displayName + '様の見積書を保存しました';
  const body = [
    '見積書PDFを保存しました。',
    '',
    '注文番号: ' + orderNumber,
    '会社名・店舗名: ' + (company || '（なし）'),
    'お客様名: ' + person,
    '見積金額: ' + formatYen_(orderData['税込合計']),
    'ファイル名: ' + file.getName(),
    'PDF URL: ' + file.getUrl(),
    '顧客メール: ' + describeMailResult_(mailResult),
    '',
    'スプレッドシート: ' + getSpreadsheet_().getUrl(),
  ].join('\n');

  MailApp.sendEmail(adminEmail, subject, body);
  return true;
}

/**
 * 顧客メール送信結果を日本語1行にする（管理者通知用）。
 * @param {Object} mailResult sendEstimateMailToCustomer_ の戻り値
 * @return {string} 説明文
 */
function describeMailResult_(mailResult) {
  if (!mailResult) return '不明';
  if (mailResult.sent) {
    return mailResult.mode + 'モードで送信済み（宛先: ' + mailResult.to + '）';
  }
  return '未送信（' + (mailResult.reason || '理由不明') + '）';
}

/***********************************************************************
 * 設定・補助
 ***********************************************************************/

/**
 * 設定シートに本機能用のキーが無ければ追加する。
 * 既存の値は一切変更しない（無い行だけ追加）。
 */
function ensureAutoEstimateSettings_() {
  const map = getSettingsMap_();
  const sheet = getSheet_(SHEETS.settings);
  const has = function(key) { return Object.prototype.hasOwnProperty.call(map, key); };
  if (!has(AUTO_ESTIMATE_SETTING_KEYS.mailMode)) {
    sheet.appendRow([AUTO_ESTIMATE_SETTING_KEYS.mailMode, 'テスト']);
  }
  if (!has(AUTO_ESTIMATE_SETTING_KEYS.customerMail)) {
    sheet.appendRow([AUTO_ESTIMATE_SETTING_KEYS.customerMail, 'ON']);
  }
  // 件名・本文は設定シートで自由に修正できる（{宛名} が宛名に置換される）
  if (!has(AUTO_ESTIMATE_SETTING_KEYS.mailSubject)) {
    sheet.appendRow([AUTO_ESTIMATE_SETTING_KEYS.mailSubject, AUTO_ESTIMATE_MAIL_TEXT.customerSubject]);
  }
  if (!has(AUTO_ESTIMATE_SETTING_KEYS.mailBody)) {
    sheet.appendRow([AUTO_ESTIMATE_SETTING_KEYS.mailBody, AUTO_ESTIMATE_MAIL_TEXT.customerBody]);
  }
}

/**
 * 注文管理シートの最終行（最新注文）の注文番号を返す。
 * @return {string} 注文番号（データが無ければ空文字）
 */
function getLatestOrderNumber_() {
  const sheet = getSheet_(SHEETS.orders);
  if (sheet.getLastRow() <= 1) return '';
  return String(sheet.getRange(sheet.getLastRow(), 1).getValue() || '').trim();
}
