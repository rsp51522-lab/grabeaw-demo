/***********************************************************************
 * 無題 2.gs（★既存ファイルをこの内容にまるごと置き換え）
 * ---------------------------------------------------------------------
 * 変更点は1か所だけ:
 *   processFormSubmit_ の中で、注文登録完了後に
 *   autoCreateEstimateSafely_(orderNumber) を呼び出す行を追加した。
 *   （★追加 と書いてあるブロック。それ以外は既存コードのまま）
 *
 * 変更前の問題:
 *   フォーム送信後、見積書PDFの作成は手動メニュー操作が必要だった。
 * 変更内容:
 *   注文登録が終わった直後にPDF自動作成処理を呼び出す。
 *   PDF処理が失敗しても注文登録は成功として扱う（受注データを守る）。
 ***********************************************************************/

function onFormSubmit(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const result = processFormSubmit_(e);
    return result;
  } catch (error) {
    const orderNumber = safeGet_(function() { return e && e.namedValues ? getNamedValue_(e.namedValues, '注文番号') : ''; }, '');
    if (e && e.range && e.range.getSheet) {
      markFormResponseError_(e.range.getSheet(), e.range.getRow(), error.message || String(error));
    }
    logError_('onFormSubmit', orderNumber, error);
    throw error;
  } finally {
    lock.releaseLock();
  }
}

function processFormSubmit_(e) {
  validateFormEvent_(e);
  const formSheet = e.range.getSheet();
  const row = e.range.getRow();
  const rowData = getRowObject_(formSheet, row);

  if (rowData['処理状況'] === APP.processingStatus.done && rowData['注文番号']) {
    return { skipped: true, reason: 'already processed', orderNumber: rowData['注文番号'] };
  }

  const namedValues = e.namedValues || rowObjectToNamedValues_(rowData);
  const input = normalizeFormInput_(namedValues);
  validateOrderInput_(input);

  const productMap = getProductMasterMap_();
  const items = buildOrderItemsFromInput_(input, productMap);
  validateOrderItems_(items);

  const orderNumber = generateOrderNumber_();
  const customer = upsertCustomer_(input, orderNumber, items);
  const totals = calculateTotals_(items, getTaxRate_());

  appendOrderRow_(orderNumber, customer.customerNumber, input, items, totals);
  appendOrderItemRows_(orderNumber, customer.customerNumber, items);
  markFormResponseProcessed_(formSheet, row, orderNumber);
  sendOrderNotification_(orderNumber);

  // ★追加: 見積書PDFの自動作成〜ドライブ保存〜メール送信（12_AutoEstimate.gs）
  //   ここで失敗してもエラーは投げない（注文登録は成功扱いにする）。
  //   失敗時はエラーログシートへ記録され、フォーム回答に「PDF保存エラー」が入る。
  const estimateResult = autoCreateEstimateSafely_(orderNumber);

  return {
    orderNumber: orderNumber,
    customerNumber: customer.customerNumber,
    itemCount: items.length,
    totalAmount: totals.total,
    estimatePdfUrl: estimateResult ? estimateResult.url : '', // ★追加
  };
}

function markFormResponseProcessed_(sheet, row, orderNumber) {
  const headerMap = getHeaderMap_(sheet);
  const now = new Date();
  sheet.getRange(row, headerMap['処理状況']).setValue(APP.processingStatus.done);
  sheet.getRange(row, headerMap['注文番号']).setValue(orderNumber);
  sheet.getRange(row, headerMap['処理日時']).setValue(now);
  sheet.getRange(row, headerMap['エラー内容']).setValue('');
}

function markFormResponseError_(sheet, row, message) {
  const headerMap = getHeaderMap_(sheet);
  sheet.getRange(row, headerMap['処理状況']).setValue(APP.processingStatus.error);
  sheet.getRange(row, headerMap['処理日時']).setValue(new Date());
  sheet.getRange(row, headerMap['エラー内容']).setValue(message);
}

function simulateTestOrder() {
  const input = {
    customerType: 'サロン・店舗',
    companyName: 'テスト美容室',
    customerName: '山田花子',
    customerKana: '',
    postalCode: '323-0000',
    address: '栃木県小山市テスト1-2-3',
    phone: '090-0000-0000',
    email: 'test@example.com',
    contactName: '',
    desiredDate: '',
    paymentMethod: '銀行振込',
    note: '',
    productQuantities: {
      CY001: 1,
      CY003: 2,
      CY004: 1,
    },
  };
  validateOrderInput_(input);
  const productMap = getProductMasterMap_();
  const items = buildOrderItemsFromInput_(input, productMap);
  const orderNumber = generateOrderNumber_();
  const customer = upsertCustomer_(input, orderNumber, items);
  const totals = calculateTotals_(items, getTaxRate_());
  appendOrderRow_(orderNumber, customer.customerNumber, input, items, totals);
  appendOrderItemRows_(orderNumber, customer.customerNumber, items);
  return { orderNumber: orderNumber, customerNumber: customer.customerNumber, total: totals.total };
}
