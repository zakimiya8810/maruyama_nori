/**
 * @fileoverview RPA連携機能
 * 社長決裁完了時に基幹システム連携用のデータをRPAシートに書き込みます。
 */

// =============================================
// 定数定義
// =============================================

const RPA_CUSTOMER_SHEET_NAME = 'RPA処理用:顧客情報';
const RPA_PRICE_SHEET_NAME = 'RPA処理用:単価情報';

// 地区コード判定用の都道府県マッピング
const AREA_CODE_MAP = {
  '001': ['東京都', '神奈川県', '千葉県', '埼玉県', '茨城県', '栃木県', '群馬県'],
  '002': ['山形県', '福島県', '新潟県', '長野県', '富山県', '石川県', '福井県', '山梨県', '岐阜県', '静岡県', '愛知県'],
  '003': ['青森県', '岩手県', '宮城県', '秋田県', '三重県', '滋賀県', '京都府', '大阪府', '兵庫県', '奈良県', '和歌山県'],
  '004': ['鳥取県', '島根県', '岡山県', '広島県', '山口県', '徳島県', '香川県', '愛媛県', '高知県'],
  '005': ['北海道', '福岡県', '佐賀県', '長崎県', '熊本県', '大分県', '宮崎県', '鹿児島県'],
  '006': ['沖縄県']
};

// =============================================
// 顧客情報のRPA連携
// =============================================

/**
 * 決裁完了時に顧客情報をRPAシートに書き込む
 * @param {string} applicationId - 申請ID
 * @param {string} customerId - 得意先コード
 * @param {Date} approvalDate - 決裁完了日時
 */
function writeCustomerToRPA(applicationId, customerId, approvalDate) {
  try {
    console.log(`[writeCustomerToRPA] 開始 - 申請ID: ${applicationId}, 得意先コード: ${customerId}`);

    const ss = SpreadsheetApp.getActiveSpreadsheet();

    // RPAシートを取得または作成
    let rpaSheet = ss.getSheetByName(RPA_CUSTOMER_SHEET_NAME);
    if (!rpaSheet) {
      rpaSheet = createCustomerRPASheet(ss);
    }

    // 申請情報を取得
    const appSheet = ss.getSheetByName('申請管理');
    if (!appSheet) throw new Error('申請管理シートが見つかりません。');

    const appValues = appSheet.getDataRange().getValues();
    const appHeader = appValues[0];
    const appRow = appValues.find(row => String(row[appHeader.indexOf('申請ID')]) === String(applicationId));

    if (!appRow) throw new Error('申請が見つかりません。');

    const appType = appRow[appHeader.indexOf('申請種別')] || '';
    const isNew = appType === '顧客新規登録';

    // 得意先マスタから最新情報を取得
    const customerSheet = ss.getSheetByName('得意先マスタ');
    if (!customerSheet) throw new Error('得意先マスタが見つかりません。');

    const customerValues = customerSheet.getDataRange().getValues();
    const customerHeader = customerValues[0];
    const customerRow = customerValues.find(row => {
      const code = String(row[customerHeader.indexOf('得意先コード')] || '').replace(/^'/, '').replace(/'$/, '');
      return code === String(customerId).replace(/^'/, '').replace(/'$/, '');
    });

    if (!customerRow) throw new Error(`得意先コード ${customerId} がマスタに見つかりません。`);

    // データをマッピング
    const rpaData = mapCustomerDataToRPA(customerRow, customerHeader, isNew, approvalDate, ss);

    // RPAシートに書き込み
    const rpaHeader = rpaSheet.getRange(1, 1, 1, rpaSheet.getLastColumn()).getValues()[0];
    const newRow = rpaHeader.map(colName => rpaData[colName] || '');

    rpaSheet.appendRow(newRow);

    console.log(`[writeCustomerToRPA] 完了 - 得意先コード: ${customerId} をRPAシートに書き込みました。`);

    return { success: true, message: 'RPA連携データを作成しました。' };

  } catch (e) {
    console.error('[writeCustomerToRPA] エラー:', e);
    return { success: false, message: 'RPA連携データの作成に失敗しました: ' + e.message };
  }
}

/**
 * 得意先マスタデータをRPAフォーマットにマッピング
 */
function mapCustomerDataToRPA(customerRow, customerHeader, isNew, approvalDate, spreadsheet) {
  // ヘッダーからインデックスを取得するヘルパー
  const getCol = (name) => {
    const idx = customerHeader.indexOf(name);
    return idx !== -1 ? customerRow[idx] : '';
  };

  // 基本情報
  const customerId = String(getCol('得意先コード')).replace(/^'/, '').replace(/'$/, '');
  const customerName = getCol('得意先名称');
  const billingCode = String(getCol('請求先コード')).replace(/^'/, '').replace(/'$/, '').trim();
  const paymentPattern = getCol('締パターン名称') || '';
  const address = getCol('住所_1') || '';

  // 親子区分の判定
  const parentChildClass = billingCode ? '子' : '親';

  // 掛現区分の判定
  const paymentType = paymentPattern === '現金' ? '現金' : '掛';

  // 入金形態の判定
  const depositType = paymentPattern === '現金' ? '現金' : '振込';

  // 宛名の取得（請求先がある場合はその会社名、ない場合は自社名）
  let addresseeName = customerName;
  if (billingCode) {
    const billingCustomer = findCustomerByCode(billingCode, spreadsheet);
    if (billingCustomer) {
      addresseeName = billingCustomer['得意先名称'] || customerName;
    }
  }

  // 海苔取扱の判定
  const noriUsage = getCol('海苔使用本数');
  const noriHandling = (noriUsage && String(noriUsage).trim() !== '' && Number(noriUsage) > 0) ? '扱う' : '扱わない';

  // お茶取扱の判定
  const teaUsage = getCol('お茶使用量');
  const teaHandling = (teaUsage && String(teaUsage).trim() !== '' && Number(teaUsage) > 0) ? '扱う' : '扱わない';

  // その他取扱の判定
  const naturalFood = getCol('自然食品取扱');
  const otherHandling = naturalFood === '有' ? '扱う' : '扱わない';

  // 地区コードの判定
  const areaCode = getAreaCodeFromAddress(address);

  // 送料発生金額の判定
  const shippingFee = getCol('配送料');
  const shippingAmount = shippingFee === '有' ? '10000' : '';

  // 取引開始日付の取得
  let tradeStartDate = getCol('取引開始日付');
  if (!tradeStartDate || String(tradeStartDate).trim() === '') {
    if (isNew) {
      // 新規の場合は決裁日
      tradeStartDate = Utilities.formatDate(approvalDate, Session.getScriptTimeZone(), 'yyyy/MM/dd');
    } else {
      // 修正の場合は登録日
      tradeStartDate = getCol('登録日') || '';
      if (tradeStartDate instanceof Date) {
        tradeStartDate = Utilities.formatDate(tradeStartDate, Session.getScriptTimeZone(), 'yyyy/MM/dd');
      }
    }
  } else if (tradeStartDate instanceof Date) {
    tradeStartDate = Utilities.formatDate(tradeStartDate, Session.getScriptTimeZone(), 'yyyy/MM/dd');
  }

  // 更新日（決裁完了日）
  const updateDate = Utilities.formatDate(approvalDate, Session.getScriptTimeZone(), 'yyyy/MM/dd HH:mm:ss');

  // コード列の0落ち対策用ヘルパー
  const padCode = (value) => {
    if (!value || String(value).trim() === '') return '';
    return "'" + String(value).trim();
  };

  // 営業担当者コードの0落ち対策
  const handlerCode = String(getCol('営業担当者コード')).replace(/^'/, '').replace(/'$/, '').trim();

  // 集計コードの0落ち対策
  const aggregationCode = String(getCol('得意先グループコード')).replace(/^'/, '').replace(/'$/, '').trim();

  // RPAデータオブジェクトを作成
  return {
    '処理列': '',
    '新規/修正': isNew ? '新規' : '修正',
    '更新日': updateDate,
    '得意先コード': padCode(customerId),
    '得意先名称': customerName,
    '略名称': getCol('略名称'),
    'カナ名称': getCol('カナ名称'),
    '郵便番号': getCol('郵便番号'),
    '住所１': getCol('住所_1'),
    '住所２': getCol('住所_2'),
    'TEL番号': getCol('TEL番号'),
    'FAX番号': getCol('FAX番号'),
    '業態': getCol('業態'),
    '営業担当': padCode(handlerCode),
    '配送担当': getCol('配送方法'),
    '親子区分': parentChildClass,
    '課税区分': '課税',
    '請求先': padCode(billingCode),
    '締日パターン': paymentPattern,
    '掛現区分': paymentType,
    '課税方式': getCol('課税方式名称'),
    '課税単位': '伝票単位',
    '入金形態': depositType,
    '宛名': addresseeName,
    '与信限度額': getCol('与信限度額'),
    '海苔取扱': noriHandling,
    'お茶取扱': teaHandling,
    'その他取扱': otherHandling,
    '優先出荷区分': getCol('出荷場所'),
    '集計コード': padCode(aggregationCode),
    '出力順': padCode(customerId),
    '取引開始日付': tradeStartDate,
    '得意先ランク': getCol('得意先ランク区分名称'),
    '代表者氏名': getCol('代表者氏名'),
    '担当者氏名': getCol('得意先担当者'),
    '送料有無区分': shippingFee,
    '地区コード': padCode(areaCode),
    '送料発生金額（未満）': shippingAmount,
    '請：郵便番号': getCol('請求先郵便番号'),
    '請：宛名': getCol('請求先名称'),
    '請：住所①': getCol('請求先住所1'),
    '請：住所②': getCol('請求先住所2'),
    '請：TEL番号': getCol('請求先TEL番号'),
    '請：FAX番号': getCol('請求先FAX番号'),
    'E-Mail': '', // 現在未実装
    '備考１': '', // 現在未実装
    '備考２': ''  // 現在未実装
  };
}

/**
 * 住所から地区コードを判定
 */
function getAreaCodeFromAddress(address) {
  if (!address || address.trim() === '') return '';

  const addressStr = String(address);

  // 各地区コードの都道府県リストをチェック
  for (const [code, prefectures] of Object.entries(AREA_CODE_MAP)) {
    for (const prefecture of prefectures) {
      if (addressStr.includes(prefecture)) {
        return code;
      }
    }
  }

  // 該当なしの場合は空欄
  return '';
}

/**
 * 得意先コードから顧客情報を取得
 */
function findCustomerByCode(customerId, spreadsheet) {
  try {
    const customerSheet = spreadsheet.getSheetByName('得意先マスタ');
    if (!customerSheet) return null;

    const values = customerSheet.getDataRange().getValues();
    const header = values[0];
    const idCol = header.indexOf('得意先コード');
    const nameCol = header.indexOf('得意先名称');

    if (idCol === -1) return null;

    const cleanId = String(customerId).replace(/^'/, '').replace(/'$/, '').trim();
    const row = values.find(r => {
      const rowId = String(r[idCol] || '').replace(/^'/, '').replace(/'$/, '').trim();
      return rowId === cleanId;
    });

    if (!row) return null;

    const result = {};
    header.forEach((key, idx) => {
      result[key] = row[idx];
    });

    return result;

  } catch (e) {
    console.error('[findCustomerByCode] エラー:', e);
    return null;
  }
}

/**
 * RPA顧客情報シートを作成
 */
function createCustomerRPASheet(spreadsheet) {
  const sheet = spreadsheet.insertSheet(RPA_CUSTOMER_SHEET_NAME);

  // ヘッダー行を設定
  const headers = [
    '処理列', '更新日', '新規/修正', '得意先コード', '得意先名称', '略名称', 'カナ名称',
    '郵便番号', '住所１', '住所２', 'TEL番号', 'FAX番号', '業種', '営業担当', '配送担当',
    '親子区分', '課税区分', '請求先', '締日パターン', '掛現区分', '課税方式', '課税単位',
    '入金形態', '宛名', '与信限度額', '海苔取扱', 'お茶取扱', 'その他取扱', '優先出荷区分',
    '集計コード', '出力順', '取引開始日付', '得意先ランク', '代表者氏名', '担当者氏名',
    '送料有無区分', '地区コード', '送料発生金額（未満）', 'E-Mail', '備考１', '備考２'
  ];

  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#e0e0e0');
  sheet.setFrozenRows(1);

  console.log('[createCustomerRPASheet] RPA顧客情報シートを作成しました。');

  return sheet;
}

// =============================================
// 単価情報のRPA連携
// =============================================

/**
 * 決裁完了時に単価情報をRPAシートに書き込む
 * @param {string} applicationId - 申請ID
 * @param {string} customerId - 得意先コード
 * @param {Date} approvalDate - 決裁完了日時
 */
function writePriceToRPA(applicationId, customerId, approvalDate) {
  try {
    console.log(`[writePriceToRPA] 開始 - 申請ID: ${applicationId}, 得意先コード: ${customerId}`);

    const ss = SpreadsheetApp.getActiveSpreadsheet();

    // RPAシートを取得または作成
    let rpaSheet = ss.getSheetByName(RPA_PRICE_SHEET_NAME);
    if (!rpaSheet) {
      rpaSheet = createPriceRPASheet(ss);
    }

    // 申請詳細データを取得
    const appDetails = getApplicationDetails(applicationId);
    if (!appDetails || !appDetails.prices || appDetails.prices.length === 0) {
      console.log('[writePriceToRPA] 単価情報が存在しません。');
      return { success: true, message: '単価情報が存在しないためスキップしました。' };
    }

    // 登録有効日を取得
    const effectiveDate = appDetails['登録有効日'] || '';
    let effectiveDateStr = '';
    if (effectiveDate) {
      if (effectiveDate === '現行' || effectiveDate === 'current') {
        effectiveDateStr = '現行';
      } else if (effectiveDate instanceof Date) {
        effectiveDateStr = Utilities.formatDate(effectiveDate, Session.getScriptTimeZone(), 'yyyy/MM/dd');
      } else {
        try {
          const date = new Date(effectiveDate);
          if (!isNaN(date.getTime())) {
            effectiveDateStr = Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy/MM/dd');
          } else {
            effectiveDateStr = String(effectiveDate);
          }
        } catch (e) {
          effectiveDateStr = String(effectiveDate);
        }
      }
    }

    // 更新日（決裁完了日）
    const updateDate = Utilities.formatDate(approvalDate, Session.getScriptTimeZone(), 'yyyy/MM/dd HH:mm:ss');

    // RPAシートのヘッダーを取得
    const rpaHeader = rpaSheet.getRange(1, 1, 1, rpaSheet.getLastColumn()).getValues()[0];

    // 各価格情報をRPAシートに書き込み
    let writeCount = 0;
    appDetails.prices.forEach(priceItem => {
      const registrationType = priceItem['登録区分'] || '';

      // 新規/修正の判定
      let operationType = '';
      if (registrationType === '追加') {
        operationType = '新規';
      } else if (registrationType === '変更') {
        operationType = '修正';
      } else if (registrationType === '削除') {
        operationType = '削除';
      } else {
        operationType = registrationType; // そのまま使用
      }

      const productCode = priceItem['商品コード'] || '';

      // 実際販売単価（実際販売価格_修正後を使用）
      let actualPrice = '';
      if (priceItem['実際販売価格_修正後'] !== undefined && priceItem['実際販売価格_修正後'] !== '') {
        actualPrice = priceItem['実際販売価格_修正後'];
      } else if (priceItem['実際販売価格_修正前'] !== undefined) {
        // 修正後がない場合は修正前（削除の場合など）
        actualPrice = priceItem['実際販売価格_修正前'];
      }

      // コード列の0落ち対策
      const cleanCustomerId = String(customerId).replace(/^'/, '').replace(/'$/, '').trim();
      const cleanProductCode = String(productCode).replace(/^'/, '').replace(/'$/, '').trim();

      // RPAデータオブジェクトを作成
      const rpaData = {
        '処理列': '',
        '更新日': updateDate,
        '新規/修正': operationType,
        '得意先コード': cleanCustomerId ? "'" + cleanCustomerId : '',
        '商品コード': cleanProductCode ? "'" + cleanProductCode : '',
        '実際販売単価': actualPrice,
        '備考': '',
        '登録有効日': effectiveDateStr
      };

      // RPAシートに書き込み
      const newRow = rpaHeader.map(colName => rpaData[colName] || '');
      rpaSheet.appendRow(newRow);
      writeCount++;
    });

    console.log(`[writePriceToRPA] 完了 - ${writeCount}件の単価情報をRPAシートに書き込みました。`);

    return { success: true, message: `${writeCount}件の単価情報をRPA連携しました。` };

  } catch (e) {
    console.error('[writePriceToRPA] エラー:', e);
    return { success: false, message: '単価情報のRPA連携に失敗しました: ' + e.message };
  }
}

/**
 * RPA単価情報シートを作成
 */
function createPriceRPASheet(spreadsheet) {
  const sheet = spreadsheet.insertSheet(RPA_PRICE_SHEET_NAME);

  // ヘッダー行を設定
  const headers = [
    '処理列',
    '更新日',
    '新規/修正',
    '得意先コード',
    '商品コード',
    '実際販売単価',
    '備考',
    '登録有効日'
  ];

  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#e0e0e0');
  sheet.setFrozenRows(1);

  console.log('[createPriceRPASheet] RPA単価情報シートを作成しました。');

  return sheet;
}
