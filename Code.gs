/**
 * @fileoverview Customer Karte アプリケーションのバックエンドスクリプト
 * スプレッドシートからのデータ取得、更新、Webアプリの配信を担当します。
 * 商談履歴シートや申請管理シートは初回書き込み時に自動生成されます。
 */

// =============================================
// グローバル定数
// =============================================
// スプレッドシートIDを自動取得（このコードが含まれているスプレッドシート）
const SPREADSHEET = SpreadsheetApp.getActiveSpreadsheet();
const SPREADSHEET_ID = SPREADSHEET.getId();
const CUSTOMER_SHEET = SPREADSHEET.getSheetByName('得意先マスタ');
const TANKA_SHEET = SPREADSHEET.getSheetByName('単価マスタ');
const PRODUCT_SHEET = SPREADSHEET.getSheetByName('商品マスタ');
const REVISION_SHEET = SPREADSHEET.getSheetByName('価格改定リスト');
const EMPLOYEE_SHEET = SPREADSHEET.getSheetByName('社員マスタ');
const ALERT_THRESHOLD_MONTH = 3; // Number of months since last visit to show an alert

// キャッシュキー定数（CACHE_DURATIONはMasterData.jsで定義済み）
const CACHE_KEY_CUSTOMERS = 'customers_cache';
const CACHE_KEY_MEETINGS = 'meetings_cache';
const CACHE_KEY_DASHBOARD = 'dashboard_cache';
const CACHE_KEY_MASTER_DATA = 'master_data_cache'; // ★追加

const SAITAN_MASTER = {
  '701': '全形', '702': '半切', '703': '中巻', '704': '小半キザミ', '705': '小半オビ',
  '706': '横1/3', '707': '横1/4', '708': '十字', '709': '6ツ切置かん', '710': '7ツ切置かん',
  '711': '八切', '712': '', '720': 'その他裁断'
};
const FUKURO_MASTER = {
  '751': '平チャック', '752': 'チャックなし袋', '753': '10枚ずつチャック', '754': '20枚ずつチャック', '755': '25枚ずつチャック',
  '756': '50枚ずつチャック', '757': '100枚チャック', '760': 'その他パック'
};

// =============================================
// キャッシュ管理関数
// =============================================

/**
 * 特定のキャッシュをクリアします
 * @param {string} cacheKey - クリアするキャッシュのキー
 */
function clearCache(cacheKey) {
  try {
    const cache = CacheService.getScriptCache();
    cache.remove(cacheKey);
    console.log('Cache cleared:', cacheKey);
  } catch (e) {
    console.error('Failed to clear cache:', e);
  }
}

/**
 * 全てのキャッシュをクリアします
 */
function clearAllCaches() {
  try {
    const cache = CacheService.getScriptCache();
    cache.removeAll([
      CACHE_KEY_CUSTOMERS,
      CACHE_KEY_MEETINGS,
      CACHE_KEY_DASHBOARD,
      CACHE_KEY_MASTER_DATA
    ]);
    console.log('All caches cleared');
  } catch (e) {
    console.error('Failed to clear all caches:', e);
  }
}

/**
 * 渋滞金リストの基準年月を取得します（スクリプトプロパティから）
 * @return {string} "YYYY-MM" 形式の年月文字列、未設定の場合は空文字
 */
function getJutaikinBaseYearMonth() {
  try {
    return PropertiesService.getScriptProperties().getProperty('jutaikin_base_yearmonth') || '';
  } catch (e) {
    console.error('getJutaikinBaseYearMonth error:', e);
    return '';
  }
}

// =============================================
// ユーティリティ関数
// =============================================

/**
 * 文字列の先頭と末尾のシングルクォートを除去します
 * @param {string|number} value - 処理対象の値
 * @return {string} シングルクォートが除去された文字列
 */
function cleanSingleQuotes(value) {
  return String(value).replace(/^'/, '').replace(/'$/, '');
}

// =============================================
// Web Application Main Process
// =============================================
function doGet(e) {
  if (e && e.parameter && e.parameter.page === 'manual') {
    return HtmlService.createHtmlOutputFromFile('Manual')
      .setTitle('顧客カルテ - 操作マニュアル')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1.0')
      .setSandboxMode(HtmlService.SandboxMode.IFRAME);
  }
  const htmlOutput = HtmlService.createHtmlOutputFromFile('index.html')
    .setTitle('Customer Karte')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0')
    .setSandboxMode(HtmlService.SandboxMode.IFRAME); // Enable modern JavaScript
  return htmlOutput;
}

/**
 * マニュアルページのURLを返します。
 * フロントエンドから呼び出されて新しいタブでマニュアルを開くために使用します。
 */
function getManualUrl() {
  return ScriptApp.getService().getUrl() + '?page=manual';
}

// =============================================
// Data Retrieval Functions Called from Frontend
// =============================================
/**
 * ログインユーザー情報を取得します。
 * 認証システムと統合し、カスタム認証情報を返します。
 */
function getLoginUser() {
  try {
    // カスタム認証を使用
    const currentUser = getCurrentUser();
    if (currentUser) {
      return {
        authenticated: true,
        id: currentUser.id,
        name: currentUser.name,
        email: currentUser.email,
        role: currentUser.role,
        department: currentUser.department,
        adminPermission: currentUser.adminPermission, // ★追加
        handlerCode: currentUser.handlerCode // ★追加
      };
    }

    // 未認証の場合
    return {
      authenticated: false,
      email: Session.getActiveUser().getEmail() // Google認証情報はフォールバックとして保持
    };
  } catch (e) {
    console.error('Failed to get login user:', e);
    return {
      authenticated: false,
      error: e.message
    };
  }
}

/**
 * 社員マスタから全社員のリストを取得します。
 * 分析レポートで使用します。
 */
function getEmployees() {
  try {
    if (!EMPLOYEE_SHEET) {
      console.warn('社員マスタシートが見つかりません。');
      return [];
    }

    const values = EMPLOYEE_SHEET.getDataRange().getValues();
    if (values.length <= 1) return [];

    const header = values.shift();
    const idCol = header.indexOf('ID');
    const nameCol = header.indexOf('担当者名');
    const emailCol = header.indexOf('メールアドレス');

    if (idCol === -1 || nameCol === -1) {
      throw new Error('社員マスタのヘッダーが不正です。');
    }

    const employees = values.map(row => ({
      id: row[idCol],
      name: row[nameCol],
      email: emailCol !== -1 ? (row[emailCol] || '') : ''
    })).filter(emp => emp.id && emp.name);

    return employees;
  } catch (e) {
    console.error('Failed to get employees:', e);
    throw new Error('サーバーエラー: 社員リストの取得に失敗しました。');
  }
}

/**
 * 顧客リストを取得します（キャッシュ付き）
 * @return {Array<Object>} 顧客データ配列
 */
function getCustomers() {
  try {
    const cache = CacheService.getScriptCache();
    const cached = cache.get(CACHE_KEY_CUSTOMERS);

    if (cached) {
      console.log('Cache hit: customers');
      return JSON.parse(cached);
    }

    console.log('Cache miss: customers');
    const data = fetchCustomers();
    cache.put(CACHE_KEY_CUSTOMERS, JSON.stringify(data), CACHE_DURATION);
    return data;

  } catch (e) {
    console.error('顧客データ取得エラー:', e);
    // キャッシュエラー時は直接取得
    return fetchCustomers();
  }
}

/**
 * 顧客リストをスプレッドシートから取得します（内部関数）
 * @return {Array<Object>} 顧客データ配列
 */
function fetchCustomers() {
  try {
    const meetingSheet = SPREADSHEET.getSheetByName('商談管理');
    const lastVisitMap = {};

    if (meetingSheet) {
      const meetingValues = meetingSheet.getDataRange().getValues();
      if (meetingValues.length > 1) {
        const header = meetingValues[0];
        const customerIdCol = header.indexOf('得意先コード');
        const timestampCol = header.indexOf('タイムスタンプ');

        if (customerIdCol !== -1 && timestampCol !== -1) {
          meetingValues.slice(1).forEach(row => {
            const customerId = row[customerIdCol];
            const resultCol = header.indexOf('結果');
            // 実績が登録されている（結果が入力されている）商談のみを対象
            if (row[timestampCol] && row[resultCol]) {
              const meetingDate = new Date(row[timestampCol]);
              if (customerId && !isNaN(meetingDate.getTime())) {
                if (!lastVisitMap[customerId] || meetingDate > lastVisitMap[customerId]) {
                  lastVisitMap[customerId] = meetingDate;
                }
              }
            }
          });
        }
      }
    }

    // 渋滞金リストに得意先コードが載っていれば対象とする
    const jutaikinMap = {};
    const jutaikinSheet = SPREADSHEET.getSheetByName('渋滞金リスト');
    if (jutaikinSheet) {
      const jutaikinValues = jutaikinSheet.getDataRange().getValues();
      if (jutaikinValues.length > 1) {
        const jutaikinHeader = jutaikinValues[0];

        // ★修正: getJutaikinByCustomerId関数と同じ方式で列インデックスを取得（.trim()を使用）
        const jutaikinColIndex = jutaikinHeader.reduce((acc, col, i) => {
          if (col) acc[String(col).trim()] = i;
          return acc;
        }, {});

        const jutaikinCustomerIdCol = jutaikinColIndex['得意先コード'];

        if (jutaikinCustomerIdCol !== undefined) {
          let jutaikinCount = 0;
          jutaikinValues.slice(1).forEach(row => {
            let customerId = row[jutaikinCustomerIdCol];
            if (customerId) {
              // シングルクォートを除去して0埋めを統一（7桁）
              const originalId = String(customerId);
              customerId = cleanSingleQuotes(customerId);
              customerId = customerId.padStart(7, '0');
              // 渋滞金リストに載っていれば対象
              jutaikinMap[customerId] = true;
              jutaikinCount++;
            }
          });
          console.log(`渋滞金リスト読み込み完了: ${jutaikinCount}件`);
        } else {
          console.warn('「渋滞金リスト」シートに「得意先コード」列が見つかりません。');
          console.warn('利用可能な列名:', Object.keys(jutaikinColIndex).join(', '));
        }
      } else {
        console.warn('「渋滞金リスト」シートにデータがありません（ヘッダーのみまたは空）。');
      }
    } else {
      console.warn('「渋滞金リスト」シートが見つかりません。');
    }

    const customerValues = CUSTOMER_SHEET.getDataRange().getValues();
    const customerHeader = customerValues.shift();

    const colIndex = {
      id: customerHeader.indexOf('得意先コード'),
      name: customerHeader.indexOf('得意先名称'),
      rank: customerHeader.indexOf('得意先ランク区分名称'),
      department: customerHeader.indexOf('拠点名称'),
      handlerName: customerHeader.indexOf('営業担当者名称'),
      businessType: customerHeader.indexOf('業態'),
      startDate: customerHeader.indexOf('登録日'),
      deleteFlag: customerHeader.indexOf('削除フラグ'),
      address: customerHeader.indexOf('住所_1'),
      tel: customerHeader.indexOf('TEL番号'),
      aggregationCode: customerHeader.indexOf('得意先グループコード'),
      billingPostal: customerHeader.indexOf('請求先郵便番号'),
      billingAddr1: customerHeader.indexOf('請求先住所1'),
      billingAddr2: customerHeader.indexOf('請求先住所2'),
      billingTel: customerHeader.indexOf('請求先TEL番号'),
      billingFax: customerHeader.indexOf('請求先FAX番号')
    };

    // 社員マスタを取得（営業担当者コードから部門名を取得するため）
    const employees = getEmployeesWithSort();
    const employeeMap = {};

    // 複数の形式でマッピングを作成（フォーマット違いに対応）
    employees.forEach(emp => {
      const code = emp.employeeCode;
      // 元のコードで登録
      employeeMap[code] = emp;
      // トリム後のコードでも登録
      const trimmedCode = String(code).trim();
      if (trimmedCode !== code) {
        employeeMap[trimmedCode] = emp;
      }
    });

    console.log(`社員マスタ取得完了: ${employees.length}件`);
    console.log('社員マスタのサンプルキー:', Object.keys(employeeMap).slice(0, 5).join(', '));

    const alertDate = new Date();
    alertDate.setMonth(alertDate.getMonth() - ALERT_THRESHOLD_MONTH);

    // マッチング統計用カウンター
    let matchedCount = 0;
    let unmatchedCount = 0;
    const unmatchedCodes = [];

    const customers = customerValues.map(row => {
      const customerId = row[colIndex.id];
      const customerName = row[colIndex.name] || '';
      const deleteFlag = row[colIndex.deleteFlag];
      const lastVisit = lastVisitMap[customerId];
      const needsAlert = !lastVisit || lastVisit < alertDate;

      const rankValue = colIndex.rank !== -1 ? String(row[colIndex.rank] || '').trim() : '';
      const isHidden = customerName.includes('Ｏ＿') || rankValue === 'ＯＬＤ' || (deleteFlag && Number(deleteFlag) >= 1);

      // 顧客IDを正規化（シングルクォート除去、7桁0埋め）してマップを検索
      let normalizedCustomerId = cleanSingleQuotes(String(customerId));
      normalizedCustomerId = normalizedCustomerId.padStart(7, '0');
      const hasJutaikin = jutaikinMap[normalizedCustomerId] || false;

      // デバッグログ（渋滞金判定）
      if (hasJutaikin) {
        console.log(`渋滞金発生: 得意先コード=${customerId} (正規化後: ${normalizedCustomerId})`);
      }

      // 営業担当者コードを追加（ソート用）
      const handlerCodeCol = customerHeader.indexOf('営業担当者コード');
      const handlerCode = handlerCodeCol !== -1 ? row[handlerCodeCol] : '';

      // 営業担当者コードから社員マスタの部門名を取得
      const cleanedHandlerCode = cleanSingleQuotes(String(handlerCode || '').trim());
      const employee = employeeMap[cleanedHandlerCode];

      // マッチング統計を更新
      if (cleanedHandlerCode) {
        if (employee) {
          matchedCount++;
          // デバッグログ（最初の3件のみ）
          if (matchedCount <= 3) {
            console.log(`✓ 顧客: ${customerName}, コード: [${cleanedHandlerCode}], 部門: ${employee.departmentName}`);
          }
        } else {
          unmatchedCount++;
          if (unmatchedCodes.length < 10) {
            unmatchedCodes.push(cleanedHandlerCode);
          }
          // デバッグログ（最初の3件のみ）
          if (unmatchedCount <= 3) {
            console.log(`✗ 顧客: ${customerName}, コード: [${cleanedHandlerCode}] → マッチせず、空欄に設定`);
          }
        }
      }

      const departmentName = employee ? employee.departmentName : ''; // 見つからない場合は空欄
      const divisionName = employee ? employee.division : ''; // ★追加: 大区分

      return {
        id: customerId,
        name: customerName,
        rank: row[colIndex.rank],
        department: departmentName,
        division: divisionName, // ★追加
        handlerName: row[colIndex.handlerName],
        handlerCode: cleanedHandlerCode, // ★追加: フィルタ用
        businessType: row[colIndex.businessType],
        startDate: row[colIndex.startDate] ? Utilities.formatDate(new Date(row[colIndex.startDate]), "JST", "yyyy/MM/dd") : '',
        isHidden: isHidden,
        address: row[colIndex.address],
        tel: row[colIndex.tel],
        needsAlert: needsAlert,
        hasJutaikin: hasJutaikin,
        '営業担当者コード': handlerCode,
        '得意先ランク区分名称': row[colIndex.rank],
        '得意先グループコード': colIndex.aggregationCode !== -1 ? row[colIndex.aggregationCode] : '',
        '請求先郵便番号': colIndex.billingPostal !== -1 ? row[colIndex.billingPostal] : '',
        '請求先住所1': colIndex.billingAddr1 !== -1 ? row[colIndex.billingAddr1] : '',
        '請求先住所2': colIndex.billingAddr2 !== -1 ? row[colIndex.billingAddr2] : '',
        '請求先TEL番号': colIndex.billingTel !== -1 ? row[colIndex.billingTel] : '',
        '請求先FAX番号': colIndex.billingFax !== -1 ? row[colIndex.billingFax] : ''
      };
    });

    // マッチング統計を出力
    console.log(`\n=== 営業担当者コードマッチング統計 ===`);
    console.log(`マッチ成功: ${matchedCount}件`);
    console.log(`マッチ失敗: ${unmatchedCount}件`);
    if (unmatchedCodes.length > 0) {
      console.log(`マッチしなかったコード（最大10件）: ${unmatchedCodes.join(', ')}`);
    }
    console.log(`======================================\n`);

    // ソート処理を適用
    const sortedCustomers = sortCustomers(customers);

    return sortedCustomers;
  } catch (e) {
    console.error('Failed to get customer list:', e);
    throw new Error('サーバーエラー: 顧客リストの取得に失敗しました。');
  }
}

/**
 * フィルタ用の選択肢をマスタデータから取得します
 * @return {Object} ソート済みフィルタオプション
 */
function getFilterOptions() {
  try {
    return {
      ranks: getFilterRanks(),
      departments: getFilterDepartments(),
      employees: getFilterEmployees(),
      businessTypes: getFilterBusinessTypes()
    };
  } catch (e) {
    console.error('フィルタオプション取得エラー:', e);
    throw new Error('サーバーエラー: フィルタオプションの取得に失敗しました。');
  }
}

/**
 * 顧客詳細データを取得します。basicInfo に課税方式名称を含めます。
 */
function getCustomerDetails(customerId) {
  try {
    console.log('[getCustomerDetails] 開始 - customerId:', customerId);
    const basicInfo = getCustomerBasicInfo(customerId); // ★ 修正された getCustomerBasicInfo を呼び出す
    console.log('[getCustomerDetails] basicInfo取得完了:', basicInfo ? 'データあり' : 'データなし');
    // 顧客情報が見つからない場合はエラーを投げる
    if (!basicInfo || Object.keys(basicInfo).length === 0 || !basicInfo['得意先コード']) {
        throw new Error(`指定された得意先コード [${customerId}] の基本情報が見つかりません。`);
    }

    // 営業担当者コードから社員マスタの部門名を取得
    const handlerCode = basicInfo['営業担当者コード'] || '';
    let departmentName = '';
    let divisionName = '';

    if (handlerCode) {
      const cleanedHandlerCode = cleanSingleQuotes(String(handlerCode).trim());
      const employee = findEmployeeByCode(cleanedHandlerCode);
      if (employee) {
        departmentName = employee.departmentName || '';
        divisionName = employee.division || '';
      }
    }

    // 社員マスタから取得した部門名を追加
    basicInfo['部門名_社員マスタ'] = departmentName;
    basicInfo['大区分_社員マスタ'] = divisionName;

    // 未来の価格変更申請を取得（エラーが発生しても空配列を返す）
    let futurePriceApps = [];
    try {
      futurePriceApps = getPendingFuturePriceApplications(customerId);
    } catch (futureError) {
      console.error('[getCustomerDetails] 未来の価格変更申請の取得に失敗しました:', futureError);
      // エラーが発生しても処理を続行（空配列を使用）
    }

    // 各データ取得関数にもエラーハンドリングを追加
    let meetings = [];
    try {
      const cleanedHandlerCode = handlerCode ? cleanSingleQuotes(String(handlerCode).trim()) : null;
      const result = getMeetingsByCustomerId(customerId, cleanedHandlerCode);
      meetings = Array.isArray(result) ? result : [];
    } catch (meetingsError) {
      console.error('[getCustomerDetails] 商談履歴の取得に失敗しました:', meetingsError);
      meetings = [];
    }

    let prices = [];
    try {
      const result = getPricesByCustomerId(customerId, basicInfo);
      prices = Array.isArray(result) ? result : [];
    } catch (pricesError) {
      console.error('[getCustomerDetails] 単価情報の取得に失敗しました:', pricesError);
      prices = [];
    }

    let jutaikin = null;
    try {
      const result = getJutaikinByCustomerId(customerId);
      jutaikin = result || null;
    } catch (jutaikinError) {
      console.error('[getCustomerDetails] 受託金情報の取得に失敗しました:', jutaikinError);
      jutaikin = null;
    }

    let sales = {};
    try {
      const result = getSalesByCustomerId(customerId);
      sales = result || {};
    } catch (salesError) {
      console.error('[getCustomerDetails] 売上情報の取得に失敗しました:', salesError);
      sales = {};
    }

    const result = {
      basicInfo: basicInfo,
      meetings: meetings,
      prices: prices, // prices の構造は getPricesByCustomerId に依存
      jutaikin: jutaikin,
      sales: sales,
      futurePriceApps: futurePriceApps // ★ 未来の価格変更申請
    };

    console.log('[getCustomerDetails] 返却データ作成完了 - basicInfo:', basicInfo ? 'あり' : 'なし',
                ', meetings:', meetings ? meetings.length : 0, '件',
                ', prices:', prices ? prices.length : 0, '件',
                ', jutaikin:', jutaikin ? jutaikin.length : 0, '件',
                ', sales:', sales ? 'あり' : 'なし',
                ', futurePriceApps:', futurePriceApps ? futurePriceApps.length : 0, '件');

    // シリアライズテスト
    try {
      const testJson = JSON.stringify(result);
      console.log('[getCustomerDetails] シリアライズ成功 - データサイズ:', testJson.length, '文字');
    } catch (serializeError) {
      console.error('[getCustomerDetails] シリアライズエラー:', serializeError);
      console.error('[getCustomerDetails] 問題のあるデータを特定します...');

      // 各プロパティを個別にテスト
      ['basicInfo', 'meetings', 'prices', 'jutaikin', 'sales', 'futurePriceApps'].forEach(key => {
        try {
          JSON.stringify(result[key]);
          console.log('[getCustomerDetails] ' + key + ': OK');
        } catch (e) {
          console.error('[getCustomerDetails] ' + key + ': エラー', e);
        }
      });
    }

    // google.script.run で転送可能な形式にクリーンアップ
    // undefinedをnullに変換し、JSON経由で確実にシリアライズ可能にする
    try {
      const cleanResult = JSON.parse(JSON.stringify(result));
      console.log('[getCustomerDetails] クリーンアップ完了');
      console.log('[getCustomerDetails] 完了 - customerId:', customerId);
      return cleanResult;
    } catch (cleanError) {
      console.error('[getCustomerDetails] クリーンアップエラー:', cleanError);
      console.log('[getCustomerDetails] 完了（元データを返却） - customerId:', customerId);
      return result;
    }
  } catch (e) {
    console.error(`Failed to get customer details for [${customerId}]:`, e);
    // エラーメッセージに元のエラーも含める
    throw new Error('サーバーエラー: 顧客詳細の取得に失敗しました。(' + e.message + ')');
  }
}

/**
 * 指定された顧客の未来日付の保留中申請を取得します
 * @param {string} customerId - 得意先コード
 * @return {Array<Object>} 未来の保留中申請データ配列（有効日でソート済み）
 */
function getPendingFuturePriceApplications(customerId) {
  try {
    const appSheet = SPREADSHEET.getSheetByName('申請管理');
    if (!appSheet) {
      console.warn('申請管理シートが見つかりません。');
      return [];
    }

    const values = appSheet.getDataRange().getValues();
    if (values.length <= 1) return [];

    const header = values[0];
    const dataRows = values.slice(1);

    // 列インデックスを取得
    const colIndex = {
      id: header.indexOf('申請ID'),
      customerId: header.indexOf('得意先コード'),
      status: header.indexOf('承認ステータス'),
      stage: header.indexOf('承認段階'),
      targetMaster: header.indexOf('対象マスタ'),
      effectiveDate: header.indexOf('登録有効日'),
      reflectionStatus: header.indexOf('マスタ反映状態'),
      appType: header.indexOf('申請種別')
    };

    // 必須列のチェック
    if (colIndex.id === -1 || colIndex.customerId === -1) {
      console.warn('申請管理シートに必要な列が見つかりません。');
      return [];
    }

    // 顧客IDをクリーンに
    const cleanCustomerId = cleanSingleQuotes(String(customerId).trim());

    // 今日の日付（時刻なし）
    const today = new Date();
    today.setHours(0, 0, 0, 0);

    // フィルタ条件：
    // 1. 得意先コードが一致
    // 2. 対象マスタに「単価マスタ」を含む
    // 3. 承認ステータスが「決裁完了」または承認段階が「常務承認済」以降
    // 4. マスタ反映状態が「未反映」で始まる
    // 5. 登録有効日が未来日付
    const pendingApps = dataRows
      .filter(row => {
        // 得意先コードチェック
        const rowCustomerId = colIndex.customerId !== -1 ? cleanSingleQuotes(String(row[colIndex.customerId] || '').trim()) : '';
        if (rowCustomerId !== cleanCustomerId) return false;

        // 対象マスタチェック
        const targetMaster = colIndex.targetMaster !== -1 ? String(row[colIndex.targetMaster] || '') : '';
        if (!targetMaster.includes('単価マスタ')) return false;

        // 承認状態チェック（決裁完了 or 常務承認済以降）
        const status = colIndex.status !== -1 ? String(row[colIndex.status] || '') : '';
        const stage = colIndex.stage !== -1 ? String(row[colIndex.stage] || '') : '';
        const appType = colIndex.appType !== -1 ? String(row[colIndex.appType] || '') : '';

        // 新規登録の場合は「決裁完了」のみ、修正の場合は「常務承認済」以降
        let isApproved = false;
        if (appType === '顧客新規登録') {
          isApproved = stage === '決裁完了';
        } else if (appType === '商品登録修正') {
          isApproved = stage === '常務承認済' || stage === '決裁完了';
        }
        if (!isApproved) return false;

        // マスタ反映状態チェック
        const reflectionStatus = colIndex.reflectionStatus !== -1 ? String(row[colIndex.reflectionStatus] || '') : '';
        if (!reflectionStatus.startsWith('未反映')) return false;

        // 登録有効日チェック（未来日付のみ）
        const effectiveDateValue = colIndex.effectiveDate !== -1 ? row[colIndex.effectiveDate] : null;
        if (!effectiveDateValue) return false;

        try {
          let effectiveDate;
          if (effectiveDateValue instanceof Date) {
            effectiveDate = new Date(effectiveDateValue);
          } else {
            effectiveDate = new Date(effectiveDateValue);
          }

          if (isNaN(effectiveDate.getTime())) return false;
          effectiveDate.setHours(0, 0, 0, 0);

          // 未来日付のみ
          return effectiveDate > today;
        } catch (e) {
          return false;
        }
      })
      .map(row => {
        const applicationId = row[colIndex.id];
        const effectiveDateValue = row[colIndex.effectiveDate];

        let effectiveDate;
        if (effectiveDateValue instanceof Date) {
          effectiveDate = new Date(effectiveDateValue);
        } else {
          effectiveDate = new Date(effectiveDateValue);
        }
        effectiveDate.setHours(0, 0, 0, 0);

        return {
          applicationId: applicationId,
          effectiveDate: effectiveDate,
          effectiveDateStr: Utilities.formatDate(effectiveDate, SPREADSHEET.getSpreadsheetTimeZone(), 'yyyy/MM/dd'),
          appType: row[colIndex.appType] || '',
          stage: row[colIndex.stage] || '',
          reflectionStatus: row[colIndex.reflectionStatus] || ''
        };
      })
      // 有効日でソート（昇順）
      .sort((a, b) => a.effectiveDate - b.effectiveDate);

    console.log(`[getPendingFuturePriceApplications] 顧客 ${customerId}: ${pendingApps.length}件の未来申請を検出`);
    return pendingApps;

  } catch (e) {
    console.error('[getPendingFuturePriceApplications] エラー:', e);
    return [];
  }
}

/**
 * 商談リストを取得します（キャッシュ付き）
 * @return {Array<Object>} 商談データ配列
 */
function getMeetings() {
  try {
    const cache = CacheService.getScriptCache();
    const cached = cache.get(CACHE_KEY_MEETINGS);

    if (cached) {
      console.log('Cache hit: meetings');
      return JSON.parse(cached);
    }

    console.log('Cache miss: meetings');
    const data = fetchMeetings();
    cache.put(CACHE_KEY_MEETINGS, JSON.stringify(data), CACHE_DURATION);
    return data;

  } catch (e) {
    console.error('商談データ取得エラー:', e);
    // キャッシュエラー時は直接取得
    return fetchMeetings();
  }
}

/**
 * 商談リストをスプレッドシートから取得します（内部関数）
 * @return {Array<Object>} 商談データ配列
 */
function fetchMeetings() {
  const meetingSheet = SPREADSHEET.getSheetByName('商談管理');
  if (!meetingSheet) return [];
  
  let values = meetingSheet.getDataRange().getValues();
  if (values.length <= 1) return [];

  const header = values.shift();
  const colIndex = header.reduce((acc, col, i) => { acc[col] = i; return acc; }, {});

  const requiredHeaders = ['ID', '商談予定日', '得意先コード', '企業名', '担当者', 'メールアドレス', 'アポイント有無', '商談目的', '結果', '実績備考', 'タイムスタンプ', '商談実施日'];
  for(const h of requiredHeaders) {
    if(colIndex[h] === undefined) {
      throw new Error(`必須ヘッダー "${h}" が「商談管理」シートに見つかりません。`);
    }
  }

  const customerValues = CUSTOMER_SHEET.getDataRange().getValues();
  const customerHeader = customerValues.shift();
  const customerIdCol = customerHeader.indexOf('得意先コード');
  const customerRankCol = customerHeader.indexOf('得意先ランク区分名称');
  const customerRankMap = customerValues.reduce((map, row) => {
    map[row[customerIdCol]] = row[customerRankCol];
    return map;
  }, {});

  const today = new Date();
  today.setHours(0, 0, 0, 0);

  // 社員マスタから部署と大区分の情報を取得
  const employeeSheet = SPREADSHEET.getSheetByName('社員マスタ');
  const employeeMap = {};
  if (employeeSheet) {
    const empValues = employeeSheet.getDataRange().getValues();
    if (empValues.length > 1) {
      const empHeader = empValues.shift();
      const empColIndex = empHeader.reduce((acc, col, i) => { acc[col] = i; return acc; }, {});
      // 部門名は「部門名」または「部署」カラムから取得
      const deptColName = empColIndex['部門名'] !== undefined ? '部門名' : '部署';
      // ★修正: 担当者名を使用してマッピング（「担当者名」または「名前」）
      const nameColName = empColIndex['担当者名'] !== undefined ? '担当者名' : '名前';
      let mappedCount = 0;
      empValues.forEach(row => {
        const name = row[empColIndex[nameColName]];
        if (name) {
          employeeMap[name] = {
            department: row[empColIndex[deptColName]] || '',
            division: row[empColIndex['大区分']] || ''
          };
          mappedCount++;
          // ★デバッグログ（最初の3件のみ）
          if (mappedCount <= 3) {
            console.log(`社員マスタマッピング: 名前=[${name}], 部門=${employeeMap[name].department}, 大区分=${employeeMap[name].division}`);
          }
        }
      });
      console.log(`商談データ用の社員マスタマッピング完了: ${mappedCount}件`);
    }
  }

  // ★デバッグ用カウンター
  let meetingDebugCount = 0;

  const meetings = values
    .filter(row => row[colIndex['ID']])
    .map(row => {
      const scheduleDate = new Date(row[colIndex['商談予定日']]);
      scheduleDate.setHours(0, 0, 0, 0);
      const actualDateStr = row[colIndex['商談実施日']];
      const actualDate = actualDateStr ? new Date(actualDateStr) : null;
      if (actualDate) {
        actualDate.setHours(0, 0, 0, 0);
      }
      const timestampStr = row[colIndex['タイムスタンプ']];
      const timestamp = timestampStr ? new Date(timestampStr) : null;

      let status = 'upcoming';
      if (actualDate) {
        // 商談実施日が入力されている = 完了
        status = 'completed';
        // 商談実施日が予定日より後ならば遅延
        if (actualDate > scheduleDate) {
          status = 'delayed';
        }
      } else if (scheduleDate < today) {
        // 予定日が過去で、まだ実施されていない = 対応遅れ
        status = 'overdue';
      }

      const customerId = row[colIndex['得意先コード']];
      const handler = row[colIndex['担当者']];
      const handlerNormalized = handler ? handler.replace(/　/g, ' ') : handler;
      const employeeInfo = employeeMap[handlerNormalized] || { department: '', division: '' };

      // ★デバッグログ（最初の3件のみ）
      if (meetingDebugCount < 3) {
        console.log(`商談データ: 担当者=[${handler}], 部門=${employeeInfo.department}, 大区分=${employeeInfo.division}`);
        meetingDebugCount++;
      }

      return {
        id: row[colIndex['ID']],
        customerId: customerId,
        scheduleDate: Utilities.formatDate(new Date(row[colIndex['商談予定日']]), "JST", "yyyy-MM-dd"),
        actualDate: actualDate ? Utilities.formatDate(actualDate, "JST", "yyyy-MM-dd") : null,
        customerName: row[colIndex['企業名']],
        rank: customerRankMap[customerId] || '-',
        handler: handler,
        department: employeeInfo.department,
        division: employeeInfo.division,
        email: row[colIndex['メールアドレス']],
        appointment: row[colIndex['アポイント有無']],
        purpose: row[colIndex['商談目的']],
        details: row[colIndex['予定備考']] || '',
        result: row[colIndex['結果']],
        notes: row[colIndex['実績備考']],
        delayReason: row[colIndex['遅延理由']],
        handlerCode: colIndex['担当者コード'] !== undefined ? (row[colIndex['担当者コード']] || '') : '',
        status: status,
        timestamp: timestamp ? Utilities.formatDate(timestamp, "JST", "yyyy-MM-dd HH:mm:ss") : null
      };
    })
    .sort((a, b) => new Date(b.scheduleDate) - new Date(a.scheduleDate));

  return meetings;
}

/**
 * 社員マスタから全社員のデータを取得します
 * @return {Array<Object>} 社員データ配列
 */
function getAllEmployees() {
  try {
    const employeeSheet = SPREADSHEET.getSheetByName('社員マスタ');
    if (!employeeSheet) return [];

    const values = employeeSheet.getDataRange().getValues();
    if (values.length <= 1) return [];

    const header = values.shift();
    const colIndex = header.reduce((acc, col, i) => { acc[col] = i; return acc; }, {});

    // 名前カラムは「担当者名」または「名前」から取得
    const nameColName = colIndex['担当者名'] !== undefined ? '担当者名' : '名前';

    // 必要なカラムの存在確認
    if (colIndex[nameColName] === undefined) {
      console.warn(`警告: 「${nameColName}」カラムが社員マスタに見つかりません。`);
      console.warn('利用可能なカラム:', header);
      return [];
    }

    // 部門名は「部門名」または「部署」カラムから取得
    const deptColName = colIndex['部門名'] !== undefined ? '部門名' : '部署';

    const employees = values
      .filter(row => row[colIndex[nameColName]])
      .filter(row => {
        if (colIndex['退職者'] === undefined) return true;
        return String(row[colIndex['退職者']] || '').trim() !== '退職';
      })
      .map(row => ({
        code: cleanSingleQuotes(row[colIndex['担当者コード']] || ''),
        name: row[colIndex[nameColName]],
        department: row[colIndex[deptColName]] || '',
        division: row[colIndex['大区分']] || ''
      }));

    return employees;
  } catch (e) {
    console.error('社員データ取得エラー:', e);
    return [];
  }
}

function getApplications() {
  try {
    // ★ 追加: ログインユーザーを取得
    const currentUser = getCurrentUser();
    if (!currentUser) {
      throw new Error('ログインが必要です');
    }

    const sheet = SPREADSHEET.getSheetByName('申請管理');
    if (!sheet) return [];

    const values = sheet.getDataRange().getValues();
    if (values.length <= 1) return [];
    const header = values.shift();

    const colIndex = header.reduce((acc, col, i) => (acc[col] = i, acc), {});

    // ★ 追加: ユーザーの役割を判定
    const userRole = determineUserRole(currentUser);

    // ★ 追加: 承認段階列の存在チェック
    const stageCol = header.indexOf('承認段階');
    const applicantIdCol = header.indexOf('申請者ID');
    const statusCol = header.indexOf('ステータス');

    // ★ 追加: データをフィルタリング
    console.log(`[getApplications フィルタ開始] userRole=${userRole}, currentUser.id=${currentUser.id}, currentUser.departmentCode=${currentUser.departmentCode}`);

    const filteredValues = values.filter((row, index) => {
      // 承認段階または（旧システムの）ステータスを取得
      const stage = stageCol !== -1 ? row[stageCol] : row[statusCol];
      const applicantId = applicantIdCol !== -1 ? row[applicantIdCol] : null;
      const appId = row[colIndex['申請ID']];

      console.log(`[getApplications 申請${appId}] stage=${stage}, applicantId=${applicantId}`);

      // ★ 修正: 申請者IDが記録されている場合のみ厳密にチェック
      if (applicantId) {
        // 申請者IDがある場合：本人の申請 OR 承認権限がある
        if (cleanSingleQuotes(String(applicantId)) === cleanSingleQuotes(String(currentUser.id))) {
          console.log(`[getApplications 申請${appId}] 本人の申請 → 表示`);
          return true; // 本人の申請
        }
        // 本人でない場合は承認権限をチェック
        const canView = canViewApplication(userRole, stage, currentUser, row, header);
        console.log(`[getApplications 申請${appId}] 承認権限チェック結果: ${canView}`);
        return canView;
      }

      // ★ 旧データ対応: 申請者IDが空の場合は全て表示（後方互換性）
      console.log(`[getApplications 申請${appId}] 申請者ID空 → 表示（旧データ）`);
      return true;
    });

    console.log(`[getApplications] フィルタ結果: ${filteredValues.length}件`);

    // ★ 追加: 申請グループごとの最新申請のみを取得
    const latestByGroup = {};

    filteredValues.forEach((row) => {
      const appId = row[colIndex['申請ID']];
      const groupIdCol = colIndex['申請グループID'];
      const groupId = groupIdCol !== undefined && row[groupIdCol]
        ? row[groupIdCol]
        : appId; // 後方互換性: 申請グループIDがない場合は申請IDを使用

      if (!latestByGroup[groupId] || appId > latestByGroup[groupId].id) {
        latestByGroup[groupId] = {
          row: row,
          id: appId
        };
      }
    });

    // 最新申請のみをマッピング
    const applications = Object.values(latestByGroup).map((item, index) => {
      const appDateTime = item.row[colIndex['申請日時']];
      let formattedDate = '日時不明';
      let sortKey = 0;

      if (appDateTime) {
        try {
          const dateObj = new Date(appDateTime);
          formattedDate = Utilities.formatDate(dateObj, "JST", "yyyy/MM/dd HH:mm");
          sortKey = dateObj.getTime();
        } catch (e) {
          console.error('[getApplications] 日付変換エラー:', e);
        }
      }

      // ★ 追加: 申請者IDから社員マスタの部門名を取得
      const applicantId = item.row[colIndex['申請者ID']] || '';
      let applicantDepartment = '';
      if (applicantId) {
        const applicantEmployee = findEmployeeById(applicantId);
        if (applicantEmployee) {
          applicantDepartment = applicantEmployee.department || '';
        }
      }

      return {
        id: item.row[colIndex['申請ID']],
        date: formattedDate,
        sortKey: sortKey, // ソート用（数値）
        type: item.row[colIndex['申請種別']],
        customerName: item.row[colIndex['対象顧客名']],
        targetCustomerId: item.row[colIndex['得意先コード']],
        applicant: item.row[colIndex['申請者名']] || item.row[colIndex['申請者']],
        applicantId: applicantId, // タブフィルタリング用
        applicantDepartment: applicantDepartment, // ★ 追加: 申請者の部署
        status: item.row[colIndex['ステータス']],
        stage: item.row[colIndex['承認段階']] || item.row[colIndex['ステータス']],
        再申請回数: item.row[colIndex['再申請回数']] || 0,
        元申請ID: item.row[colIndex['元申請ID']] || null,
        申請グループID: item.row[colIndex['申請グループID']] || item.id
      };
    });

    // ソート用の数値キーでソート
    applications.sort((a, b) => b.sortKey - a.sortKey);

    return applications;

  } catch (e) {
    console.error('申請取得エラー:', e);
    return [];
  }
}

/**
 * すべての申請を取得します（権限チェックなし）
 * 「すべての申請」タブ用
 * @return {Array<Object>} 全申請データ配列
 */
function getAllApplications() {
  try {
    console.log('[getAllApplications] 全申請を取得します（権限チェックなし）');

    const sheet = SPREADSHEET.getSheetByName('申請管理');
    if (!sheet) return [];

    const values = sheet.getDataRange().getValues();
    if (values.length <= 1) return [];
    const header = values.shift();

    const colIndex = header.reduce((acc, col, i) => (acc[col] = i, acc), {});

    // ★ フィルタリングなし：全データを取得
    // 申請グループごとの最新申請のみを取得
    const latestByGroup = {};

    values.forEach((row) => {
      const appId = row[colIndex['申請ID']];
      const groupIdCol = colIndex['申請グループID'];
      const groupId = groupIdCol !== undefined && row[groupIdCol]
        ? row[groupIdCol]
        : appId;

      if (!latestByGroup[groupId] || appId > latestByGroup[groupId].id) {
        latestByGroup[groupId] = {
          row: row,
          id: appId
        };
      }
    });

    // 最新申請のみをマッピング
    const applications = Object.values(latestByGroup).map((item) => {
      const appDateTime = item.row[colIndex['申請日時']];
      let formattedDate = '日時不明';
      let sortKey = 0;

      if (appDateTime) {
        try {
          const dateObj = new Date(appDateTime);
          formattedDate = Utilities.formatDate(dateObj, "JST", "yyyy/MM/dd HH:mm");
          sortKey = dateObj.getTime();
        } catch (e) {
          console.error('[getAllApplications] 日付変換エラー:', e);
        }
      }

      // 申請者IDから社員マスタの部門名を取得
      const applicantId = item.row[colIndex['申請者ID']] || '';
      let applicantDepartment = '';
      if (applicantId) {
        const applicantEmployee = findEmployeeById(applicantId);
        if (applicantEmployee) {
          applicantDepartment = applicantEmployee.department || '';
        }
      }

      return {
        id: item.row[colIndex['申請ID']],
        date: formattedDate,
        sortKey: sortKey,
        type: item.row[colIndex['申請種別']],
        customerName: item.row[colIndex['対象顧客名']],
        targetCustomerId: item.row[colIndex['得意先コード']],
        applicant: item.row[colIndex['申請者名']] || item.row[colIndex['申請者']],
        applicantId: applicantId,
        applicantDepartment: applicantDepartment,
        status: item.row[colIndex['ステータス']],
        stage: item.row[colIndex['承認段階']] || item.row[colIndex['ステータス']],
        再申請回数: item.row[colIndex['再申請回数']] || 0,
        元申請ID: item.row[colIndex['元申請ID']] || null,
        申請グループID: item.row[colIndex['申請グループID']] || item.id
      };
    });

    // ソート用の数値キーでソート
    applications.sort((a, b) => b.sortKey - a.sortKey);

    console.log('[getAllApplications] 全申請取得完了: ' + applications.length + '件');
    return applications;

  } catch (e) {
    console.error('[getAllApplications] エラー:', e);
    return [];
  }
}

/**
 * 申請詳細データを取得します。
 * [修正] 申請管理シートの「対象マスタ」に応じて、「申請データ_顧客」または「申請データ_単価」から読み込むように変更。
 * [修正] 顧客データ読み込み時の「商品コード」列参照を削除。
 * [★今回の修正★] 返り値の details オブジェクトに「得意先コード」を含めます。
 */
function getApplicationDetails(applicationId) {
  console.log('[getApplicationDetails] 開始 - 申請ID:', applicationId, 'タイプ:', typeof applicationId);

  const appSheet = SPREADSHEET.getSheetByName('申請管理');
  if (!appSheet) throw new Error('申請管理シートが見つかりません。');

  // --- 申請基本情報を取得 ---
  const appValues = appSheet.getDataRange().getValues();
  const appHeader = appValues.shift();
  console.log('[getApplicationDetails] ヘッダー:', appHeader.slice(0, 5).join(', ') + '...');

  const appIdCol = appHeader.indexOf('申請ID');
  console.log('[getApplicationDetails] 申請ID列:', appIdCol);

  if (appIdCol === -1) {
    throw new Error('申請管理シートに「申請ID」列が見つかりません。');
  }

  // 申請IDの型を統一して検索
  const searchId = String(applicationId).trim();
  console.log('[getApplicationDetails] 検索ID:', searchId, 'データ行数:', appValues.length);

  const appRow = appValues.find(r => String(r[appIdCol]).trim() === searchId);

  if (!appRow) {
    console.error('[getApplicationDetails] 申請が見つかりません。利用可能なID（最初の5件）:',
                  appValues.slice(0, 5).map(r => String(r[appIdCol])).join(', '));
    throw new Error('指定された申請が見つかりません。申請ID: ' + searchId);
  }

  console.log('[getApplicationDetails] 申請データ発見');
  
  const details = {};
  appHeader.forEach((key, index) => {
    // ★ 修正: キー名をトリムして空白を除去
    const cleanKey = String(key).trim();
    if (!cleanKey) return; // 空のキーはスキップ

    if (appRow[index] instanceof Date) {
        details[cleanKey] = Utilities.formatDate(appRow[index], SPREADSHEET.getSpreadsheetTimeZone(), "yyyy/MM/dd HH:mm:ss");
    } else {
        details[cleanKey] = appRow[index];
    }
  });

  const customerIdCol = appHeader.indexOf('得意先コード');
  if (customerIdCol !== -1 && appRow[customerIdCol]) {
    details['得意先コード'] = appRow[customerIdCol];
  } else {
    // 得意先コード列がない、または空欄の場合のフォールバック (特に古いデータ用)
    // ※ただし、これに頼らない運用が望ましい
    console.warn(`申請ID ${applicationId} に得意先コードが紐付いていません。`);
    if (!details['得意先コード']) details['得意先コード'] = '';
  }

  // --- [修正] 申請種別に応じて読み込むシートを決定 ---
  const targetMasterCol = appHeader.indexOf('対象マスタ');
  const appTypeCol = appHeader.indexOf('申請種別');
  
  // 新しい形式(対象マスタ列)を優先し、なければ古い形式(申請種別列)で判断
  const targetMaster = (targetMasterCol !== -1 && details['対象マスタ']) 
                       ? details['対象マスタ']
                       : (details['申請種別'] === '商品登録修正' ? '単価マスタ' : '顧客マスタ');

  let dataSheet;
  let detailData = [];

  // ★ 修正: 単価マスタが対象に含まれる場合は単価データを読み込む
  if (targetMaster && targetMaster.includes('単価マスタ')) {
    // --- 単価申請または単価を含む申請の場合 ---
    dataSheet = SPREADSHEET.getSheetByName('申請データ_単価');
    if (!dataSheet) throw new Error('申請データ_単価シートが見つかりません。');

    const dataValues = dataSheet.getDataRange().getValues();
    if (dataValues.length > 1) {
      const dataHeader = dataValues.shift();
      const dataIdCol = dataHeader.indexOf('申請ID');

      // 必須ヘッダーの確認
      if (dataIdCol === -1) {
        throw new Error('申請データ_単価シートに「申請ID」列が見つかりません。');
      }

      const detailRows = dataValues.filter(r => String(r[dataIdCol]) == String(applicationId));

      // デバッグログ
      console.log(`単価申請ID ${applicationId}: ${detailRows.length}件のデータを取得`);

      // 単価シートの全ヘッダーを読み込む
      detailData = detailRows.map(r => {
        const rowData = {};
        dataHeader.forEach((headerKey, i) => {
          if(headerKey) { // 空のヘッダーは無視
              let value = r[i];
              // ★ 修正: Date オブジェクトを文字列に変換してクライアントに返せるようにする
              if (value instanceof Date) {
                value = Utilities.formatDate(value, SPREADSHEET.getSpreadsheetTimeZone(), "yyyy/MM/dd HH:mm:ss");
              }
              rowData[headerKey] = value;
          }
        });
        return rowData;
      });
    } else {
      console.warn(`申請データ_単価シートにデータがありません（申請ID: ${applicationId}）`);
    }
  }

  // ★ 修正: 顧客マスタが対象に含まれる場合は顧客データを読み込む
  if (targetMaster && targetMaster.includes('顧客マスタ')) {
    // --- 顧客申請または顧客を含む申請の場合 ---
    dataSheet = SPREADSHEET.getSheetByName('申請データ_顧客');
    if (!dataSheet) throw new Error('申請データ_顧客シートが見つかりません。');

    const dataValues = dataSheet.getDataRange().getValues();
     if(dataValues.length > 1) {
        const dataHeader = dataValues.shift();
        const dataIdCol = dataHeader.indexOf('申請ID');

        // ★ 修正: 必須ヘッダーを確認
        const fieldCol = dataHeader.indexOf('項目名');
        const beforeCol = dataHeader.indexOf('修正前の値');
        const afterCol = dataHeader.indexOf('修正後の値');

        if (dataIdCol === -1 || fieldCol === -1 || beforeCol === -1 || afterCol === -1) {
           throw new Error('申請データ_顧客シートのヘッダー（申請ID, 項目名, 修正前の値, 修正後の値）が不正です。');
        }

        const detailRows = dataValues.filter(r => String(r[dataIdCol]) == String(applicationId));

        // ★ 修正: 単価データと顧客データが両方ある場合は、customerDetailDataに格納
        const customerDetailData = detailRows.map(r => {
          // ★ 修正: Date オブジェクトを文字列に変換してクライアントに返せるようにする
          let beforeVal = r[beforeCol];
          let afterVal = r[afterCol];

          // Date オブジェクトの場合は文字列に変換
          if (beforeVal instanceof Date) {
            beforeVal = Utilities.formatDate(beforeVal, SPREADSHEET.getSpreadsheetTimeZone(), "yyyy/MM/dd HH:mm:ss");
          }
          if (afterVal instanceof Date) {
            afterVal = Utilities.formatDate(afterVal, SPREADSHEET.getSpreadsheetTimeZone(), "yyyy/MM/dd HH:mm:ss");
          }

          return {
            field: r[fieldCol],
            before: beforeVal,
            after: afterVal,
          };
        });

        // ★ 修正: 単価データが既にある場合は別プロパティに格納、なければdetailDataに格納
        if (detailData.length > 0) {
          // 単価データが既にあるので、顧客データは別プロパティに格納
          details.customerDetailData = customerDetailData;
        } else {
          // 顧客データのみの場合は従来通りdetailDataに格納
          detailData = customerDetailData;
        }
     }
  }

  details.detailData = detailData;

  // ★追加: RPA連携用にpricesエイリアスも設定（単価マスタの場合）
  if (targetMaster && targetMaster.includes('単価マスタ')) {
    details.prices = detailData;
  }

  // ★ 修正: details オブジェクト全体をチェックして Date オブジェクトを文字列に変換
  Object.keys(details).forEach(key => {
    if (details[key] instanceof Date) {
      console.log('[getApplicationDetails] Date変換:', key);
      details[key] = Utilities.formatDate(details[key], SPREADSHEET.getSpreadsheetTimeZone(), "yyyy/MM/dd HH:mm:ss");
    }
  });

  console.log('[getApplicationDetails] 完了 - 返却データキー:', Object.keys(details).join(', '));
  console.log('[getApplicationDetails] detailDataの件数:', detailData.length);
  console.log('[getApplicationDetails] pricesの件数:', details.prices ? details.prices.length : 0);
  console.log('[getApplicationDetails] detailDataの内容:', JSON.stringify(detailData).substring(0, 200));

  // 返却前に最終チェック
  if (!details || typeof details !== 'object') {
    console.error('[getApplicationDetails] detailsが不正です:', details);
    throw new Error('申請詳細データが正しく生成されませんでした');
  }

  console.log('[getApplicationDetails] returnを実行します');
  return details;
}

/**
 * 申請を処理（承認または却下）します。
 * [★今回の修正★] 「商品登録修正」が承認された場合、新設した `updatePriceInMaster` を呼び出し、
 * 「単価マスタ」へのデータ反映を実行します。
 */
/**
 * 申請の承認/却下処理（5段階ワークフロー対応版）
 * @param {string} applicationId - 申請ID
 * @param {boolean} isApproved - true: 承認, false: 却下
 * @param {string} approverId - 承認者のID（オプション）
 * @param {string} rejectReason - 却下理由（却下時のみ）
 * @param {string} notifyTo - 通知先メールアドレス（却下時のみ）
 * @return {Object} 処理結果 {status: string, message: string}
 */
function processApplication(applicationId, isApproved, approverId, rejectReason, notifyTo, customerCode, aggregationCode) {
  try {
    console.log('[processApplication] 開始 - applicationId:', applicationId, 'isApproved:', isApproved, 'approverId:', approverId);

    const appSheet = SPREADSHEET.getSheetByName('申請管理');
    if (!appSheet) throw new Error('申請管理シートが見つかりません。');

    const values = appSheet.getDataRange().getValues();
    const header = values.shift();

    console.log('[processApplication] データ行数:', values.length);

    const idCol = header.indexOf('申請ID');
    const stageCol = header.indexOf('承認段階');
    const statusCol = header.indexOf('ステータス');

    console.log('[processApplication] 列インデックス - idCol:', idCol, 'stageCol:', stageCol);

    if (idCol === -1) {
      throw new Error('「申請ID」列が見つかりません。');
    }

    const rowIndex = values.findIndex(r => String(r[idCol]) === String(applicationId));
    console.log('[processApplication] rowIndex:', rowIndex);

    if (rowIndex === -1) throw new Error('指定された申請が見つかりません。');

    const targetRow = values[rowIndex];
    const currentStage = stageCol !== -1 ? targetRow[stageCol] : '申請中';
    const currentStatus = statusCol !== -1 ? targetRow[statusCol] : '申請中';

    console.log('[processApplication] currentStage:', currentStage, 'currentStatus:', currentStatus);

    // 承認段階がない場合は旧システムとして動作（後方互換性）
    if (stageCol === -1) {
      return processApplication_legacy(applicationId, isApproved);
    }

    // 新規登録申請の管理承認時に得意先コード・得意先グループコードを保存
    if (isApproved && currentStage === '上長承認済' && customerCode && aggregationCode) {
      const appTypeCol = header.indexOf('申請種別');
      const appType = appTypeCol !== -1 ? targetRow[appTypeCol] : '';

      if (appType === '顧客新規登録') {
        // 管理部採番コードを申請データに保存
        const customerCodeCol = header.indexOf('得意先コード');
        const aggregationCodeCol = header.indexOf('得意先グループコード_管理部採番');

        if (customerCodeCol !== -1) {
          appSheet.getRange(rowIndex + 2, customerCodeCol + 1).setValue("'" + customerCode);
          console.log(`管理部採番 - 得意先コード: ${customerCode}`);
        }

        // 得意先グループコード_管理部採番列がない場合は作成
        if (aggregationCodeCol === -1) {
          // 新しい列を追加
          const lastCol = header.length;
          appSheet.getRange(1, lastCol + 1).setValue('得意先グループコード_管理部採番');
          appSheet.getRange(rowIndex + 2, lastCol + 1).setValue("'" + aggregationCode);
          console.log(`管理部採番 - 得意先グループコード（新規列作成）: ${aggregationCode}`);
        } else {
          appSheet.getRange(rowIndex + 2, aggregationCodeCol + 1).setValue("'" + aggregationCode);
          console.log(`管理部採番 - 得意先グループコード: ${aggregationCode}`);
        }
      }
    }

    // 承認処理
    if (isApproved) {
      const result = approveApplication(applicationId, approverId || 'system');
      return {
        status: result.success ? 'success' : 'error',
        message: result.message
      };
    } else {
      // 却下処理
      const result = rejectApplication(applicationId, rejectReason || '理由未入力', notifyTo, false, [], approverId || 'system');
      return {
        status: result.success ? 'success' : 'error',
        message: result.message
      };
    }

  } catch (e) {
    console.error('申請処理エラー:', e);
    return {
      status: 'error',
      message: `処理に失敗しました: ${e.message}`
    };
  }
}

/**
 * レガシー承認処理関数（旧システム用・後方互換性のため残す）
 * @param {string} applicationId - 申請ID
 * @param {boolean} isApproved - true: 承認, false: 却下
 */
function processApplication_legacy(applicationId, isApproved) {
  const appSheet = SPREADSHEET.getSheetByName('申請管理');
  if (!appSheet) throw new Error('申請管理シートが見つかりません。');

  const values = appSheet.getDataRange().getValues();
  const header = values.shift();
  const idCol = header.indexOf('申請ID');
  const statusCol = header.indexOf('ステータス');

  const customerIdCol = header.indexOf('得意先コード');
  const appTypeCol = header.indexOf('申請種別'); // 申請種別も取得
  const effectiveDateCol = header.indexOf('登録有効日'); // ★ 登録有効日も取得

  // 得意先コード列の存在チェック
  if (customerIdCol === -1) {
    throw new Error('「申請管理」シートに「得意先コード」列が見つかりません。処理を中断します。');
  }

  const rowIndex = values.findIndex(r => String(r[idCol]) == String(applicationId));
  if (rowIndex === -1) throw new Error('指定された申請が見つかりません。');

  const targetRow = values[rowIndex];
  const currentStatus = targetRow[statusCol];
  if (currentStatus !== '申請中') {
    return { status: 'warn', message: `この申請は既に処理されています (現在のステータス: ${currentStatus})。` };
  }

  const newStatus = isApproved ? '承認済' : '却下';
  appSheet.getRange(rowIndex + 2, statusCol + 1).setValue(newStatus);

  if (isApproved) {
    // 申請詳細を取得する前に、必要な情報をtargetRowから取得
    const customerId = targetRow[customerIdCol]; // シートから得意先コードを取得
    const appType = targetRow[appTypeCol]; // シートから申請種別を取得

    // 新規登録以外で得意先コードが空の場合はエラー
    if (!customerId && appType !== '顧客新規登録') {
      appSheet.getRange(rowIndex + 2, statusCol + 1).setValue('エラー(コード不明)'); // ステータスをエラーに変更
      throw new Error(`承認処理エラー: 申請ID ${applicationId} の得意先コードが「申請管理」シートに記録されていません。`);
    }

    const appDetails = getApplicationDetails(applicationId);

    // 変更差分をオブジェクト (newData) に格納
    const newData = {};
    // ※「顧客修正」または「顧客新規」の場合のみ、この newData を使用する
    if (appType === '顧客情報修正' || appType === '顧客新規登録') {
      appDetails.detailData.forEach(item => {
          newData[item.field] = item.after; // 変更後の値だけを抽出
      });
    }

    // appDetails['申請種別'] ではなく、信頼できる appType を使用
    if (appType === '顧客情報修正') {
        // customerId (シートから取得したコード) を使って元のデータを取得
        const originalData = getCustomerBasicInfo(customerId);

        // 元データが見つからない場合はエラー
        if (!originalData || !originalData['得意先コード']) {
            appSheet.getRange(rowIndex + 2, statusCol + 1).setValue('エラー(元データ不明)');
            throw new Error(`承認処理エラー: 得意先コード ${customerId} の元データが「得意先マスタ」に見つかりません。`);
        }

        // 元データ(originalData) と 変更差分(newData) をマージ
        const mergedData = { ...originalData, ...newData };

        // マージによって得意先コードが変わってしまうことを防ぐ (念のため)
        mergedData['得意先コード'] = customerId;

        updateCustomerInMaster(mergedData);

    } else if (appType === '顧客新規登録') {
        // 新規登録の場合は、newData (変更後の値) のみを使用
        if (!newData['得意先コード']) {
            appSheet.getRange(rowIndex + 2, statusCol + 1).setValue('エラー(新規コード不明)');
            throw new Error(`承認処理エラー: 新規登録申請 ${applicationId} に得意先コードが含まれていません。`);
        }
        addNewCustomerToMaster(newData);

    } else if (appType === '商品登録修正') {
        console.log(`商品登録修正(ID: ${applicationId}, 顧客: ${customerId})の承認処理を開始します。`);
        // 登録有効日を取得 (Dateオブジェクトとして)
        let effectiveDate = null;
        if (effectiveDateCol !== -1 && targetRow[effectiveDateCol]) {
          try {
            effectiveDate = new Date(targetRow[effectiveDateCol]);
            if (isNaN(effectiveDate.getTime())) effectiveDate = null;
          } catch(e) {
            console.warn('登録有効日の解析に失敗しました:', targetRow[effectiveDateCol], e.message);
          }
        }

        // 新設した関数を呼び出す
        updatePriceInMaster(customerId, appDetails, effectiveDate);
    }
  }

  return { status: 'success', message: `申請ID: ${applicationId} を「${newStatus}」として処理しました。` };
}

// =============================================
// Data Registration/Update Functions
// =============================================

function addMeetingSchedule(scheduleData) {
  try {
    const meetingSheet = SPREADSHEET.getSheetByName('商談管理');
    const header = meetingSheet.getRange(1, 1, 1, meetingSheet.getLastColumn()).getValues()[0];

    const idColData = meetingSheet.getRange(2, 1, meetingSheet.getLastRow(), 1).getValues();
    const newId = Math.max(0, ...idColData.flat().map(Number)) + 1;

    const customerInfo = getCustomerBasicInfo(scheduleData.customerId);
    if (!customerInfo['得意先名称']) {
      throw new Error('指定された得意先コードが見つかりません。');
    }

    const rankColIndex = header.indexOf('ランク');
    const rank = rankColIndex !== -1 ? customerInfo['得意先ランク区分名称'] || '' : '';

    const newRow = header.map(colName => {
        switch(colName) {
            case 'ID': return newId;
            case '商談予定日': return scheduleData.scheduleDate;
            case '得意先コード': return "'" + cleanSingleQuotes(scheduleData.customerId); // 0落ち対策
            case '企業名': return scheduleData.customerName || customerInfo['得意先名称'];
            case 'ランク': return rank;
            case '担当者': return scheduleData.handler;
            case '担当者コード': return scheduleData.handlerCode || '';
            case '企業担当者コード': return cleanSingleQuotes(String(customerInfo['営業担当者コード'] || ''));
            case '企業担当者': return customerInfo['営業担当者名称'] || '';
            case 'アポイント有無': return scheduleData.appointment || '';
            case '商談目的': return scheduleData.purpose || '';
            case '予定備考': return scheduleData.details || '';
            case 'タイムスタンプ': return new Date();
            default: return '';
        }
    });

    meetingSheet.appendRow(newRow);

    // キャッシュをクリア
    clearMeetingsCache();
    clearDashboardCache();

    return { status: 'success', message: '商談の予定を登録しました。' };
  } catch (e) {
    console.error('Failed to add meeting schedule:', e);
    throw new Error('サーバーエラー: 予定の登録に失敗しました。 ' + e.message);
  }
}

function updateMeetingResult(resultData) {
  try {
    const meetingSheet = SPREADSHEET.getSheetByName('商談管理');
    const values = meetingSheet.getDataRange().getValues();
    const header = values.shift();
    const idCol = header.indexOf('ID');

    const rowIndex = values.findIndex(row => String(row[idCol]) === String(resultData.meetingId));
    if (rowIndex === -1) {
      throw new Error('更新対象の商談が見つかりません。');
    }

    const resultCol = header.indexOf('結果');
    const notesCol = header.indexOf('実績備考');
    const delayReasonCol = header.indexOf('遅延理由');
    const timestampCol = header.indexOf('タイムスタンプ');
    const actualDateCol = header.indexOf('商談実施日');

    const sheetRowIndex = rowIndex + 2;
    const now = new Date();

    meetingSheet.getRange(sheetRowIndex, resultCol + 1).setValue(resultData.result);
    meetingSheet.getRange(sheetRowIndex, notesCol + 1).setValue(resultData.notes);
    meetingSheet.getRange(sheetRowIndex, delayReasonCol + 1).setValue(resultData.delayReason);
    meetingSheet.getRange(sheetRowIndex, timestampCol + 1).setValue(now);

    // 商談実施日を記録（resultDataで指定されていればその日付、なければ今日）
    const actualDate = resultData.actualDate ? new Date(resultData.actualDate) : now;
    meetingSheet.getRange(sheetRowIndex, actualDateCol + 1).setValue(Utilities.formatDate(actualDate, "JST", "yyyy-MM-dd"));

    // キャッシュをクリア
    clearMeetingsCache();
    clearDashboardCache();

    return { status: 'success', message: '商談の実績を更新しました。' };
  } catch (e) {
    console.error('Failed to update meeting result:', e);
    throw new Error('サーバーエラー: 実績の更新に失敗しました。');
  }
}

/**
 * 既存の予定に実績のみを追加します
 */
function updateMeetingResultOnly(resultData) {
  try {
    const meetingSheet = SPREADSHEET.getSheetByName('商談管理');
    const values = meetingSheet.getDataRange().getValues();
    const header = values.shift();
    const idCol = header.indexOf('ID');

    const rowIndex = values.findIndex(row => String(row[idCol]) === String(resultData.meetingId));
    if (rowIndex === -1) {
      throw new Error('更新対象の商談が見つかりません。');
    }

    const notesCol = header.indexOf('実績備考');
    const timestampCol = header.indexOf('タイムスタンプ');
    const actualDateCol = header.indexOf('商談実施日');

    const sheetRowIndex = rowIndex + 2;
    const now = new Date();

    if (notesCol !== -1) {
      meetingSheet.getRange(sheetRowIndex, notesCol + 1).setValue(resultData.notes || '');
    }
    if (timestampCol !== -1) {
      meetingSheet.getRange(sheetRowIndex, timestampCol + 1).setValue(now);
    }
    if (actualDateCol !== -1 && resultData.actualDate) {
      meetingSheet.getRange(sheetRowIndex, actualDateCol + 1).setValue(Utilities.formatDate(new Date(resultData.actualDate), "JST", "yyyy-MM-dd"));
    }

    // キャッシュをクリア
    clearMeetingsCache();
    clearDashboardCache();

    return { status: 'success', message: '商談の実績を登録しました。' };
  } catch (e) {
    console.error('Failed to update meeting result only:', e);
    throw new Error('サーバーエラー: 実績の登録に失敗しました。 ' + e.message);
  }
}

/**
 * 予定なしで実績のみを新規登録します
 */
function addMeetingResultOnly(resultData) {
  try {
    const meetingSheet = SPREADSHEET.getSheetByName('商談管理');
    const header = meetingSheet.getRange(1, 1, 1, meetingSheet.getLastColumn()).getValues()[0];

    const idColData = meetingSheet.getRange(2, 1, meetingSheet.getLastRow(), 1).getValues();
    const newId = Math.max(0, ...idColData.flat().map(Number)) + 1;

    const customerInfo = getCustomerBasicInfo(resultData.customerId);
    if (!customerInfo['得意先名称']) {
      throw new Error('指定された得意先コードが見つかりません。');
    }

    const rankColIndex = header.indexOf('ランク');
    const rank = rankColIndex !== -1 ? customerInfo['得意先ランク区分名称'] || '' : '';

    const actualDate = resultData.actualDate ? new Date(resultData.actualDate) : new Date();

    const newRow = header.map(colName => {
        switch(colName) {
            case 'ID': return newId;
            case '得意先コード': return "'" + cleanSingleQuotes(resultData.customerId); // 0落ち対策
            case '企業名': return resultData.customerName || customerInfo['得意先名称'];
            case 'ランク': return rank;
            case '担当者': return resultData.handler;
            case '担当者コード': return resultData.handlerCode || '';
            case '企業担当者コード': return cleanSingleQuotes(String(customerInfo['営業担当者コード'] || ''));
            case '企業担当者': return customerInfo['営業担当者名称'] || '';
            case '商談目的': return resultData.purpose || '';
            case '実績備考': return resultData.notes || '';
            case '商談実施日': return Utilities.formatDate(actualDate, "JST", "yyyy-MM-dd");
            case 'タイムスタンプ': return new Date();
            default: return '';
        }
    });

    meetingSheet.appendRow(newRow);

    // キャッシュをクリア
    clearMeetingsCache();
    clearDashboardCache();

    return { status: 'success', message: '商談の実績を登録しました。' };
  } catch (e) {
    console.error('Failed to add meeting result only:', e);
    throw new Error('サーバーエラー: 実績の登録に失敗しました。 ' + e.message);
  }
}

/**
 * 商談情報を更新します（予定+実績の両方）
 */
function updateMeeting(meetingData) {
  try {
    const meetingSheet = SPREADSHEET.getSheetByName('商談管理');
    const values = meetingSheet.getDataRange().getValues();
    const header = values.shift();
    const idCol = header.indexOf('ID');

    const rowIndex = values.findIndex(row => String(row[idCol]) === String(meetingData.meetingId));
    if (rowIndex === -1) {
      throw new Error('更新対象の商談が見つかりません。');
    }

    const sheetRowIndex = rowIndex + 2;
    const now = new Date();

    // 顧客情報を取得
    const customerInfo = getCustomerBasicInfo(meetingData.customerId);
    const rank = customerInfo['得意先ランク区分名称'] || '';

    // 各カラムのインデックスを取得
    const colMap = {};
    header.forEach((colName, index) => {
      colMap[colName] = index + 1; // 1-based index for getRange
    });

    // 予定項目を更新
    if (colMap['商談予定日']) meetingSheet.getRange(sheetRowIndex, colMap['商談予定日']).setValue(Utilities.formatDate(new Date(meetingData.scheduleDate), "JST", "yyyy-MM-dd"));
    if (colMap['得意先コード']) meetingSheet.getRange(sheetRowIndex, colMap['得意先コード']).setValue("'" + cleanSingleQuotes(meetingData.customerId));
    if (colMap['企業名']) meetingSheet.getRange(sheetRowIndex, colMap['企業名']).setValue(meetingData.customerName);
    if (colMap['ランク']) meetingSheet.getRange(sheetRowIndex, colMap['ランク']).setValue(rank);
    if (colMap['担当者']) meetingSheet.getRange(sheetRowIndex, colMap['担当者']).setValue(meetingData.handler);
    if (colMap['担当者コード'] && meetingData.handlerCode !== undefined) meetingSheet.getRange(sheetRowIndex, colMap['担当者コード']).setValue(meetingData.handlerCode);
    if (colMap['アポイント有無']) meetingSheet.getRange(sheetRowIndex, colMap['アポイント有無']).setValue(meetingData.appointment);
    if (colMap['商談目的']) meetingSheet.getRange(sheetRowIndex, colMap['商談目的']).setValue(meetingData.purpose);
    if (colMap['予定備考']) meetingSheet.getRange(sheetRowIndex, colMap['予定備考']).setValue(meetingData.details || '');

    // 実績項目を更新
    if (meetingData.actualDate && colMap['商談実施日']) {
      meetingSheet.getRange(sheetRowIndex, colMap['商談実施日']).setValue(Utilities.formatDate(new Date(meetingData.actualDate), "JST", "yyyy-MM-dd"));
    }
    if (colMap['実績備考']) meetingSheet.getRange(sheetRowIndex, colMap['実績備考']).setValue(meetingData.notes || '');
    if (colMap['遅延理由']) meetingSheet.getRange(sheetRowIndex, colMap['遅延理由']).setValue(meetingData.delayReason || '');
    if (colMap['タイムスタンプ']) meetingSheet.getRange(sheetRowIndex, colMap['タイムスタンプ']).setValue(now);

    // キャッシュをクリア
    clearMeetingsCache();
    clearDashboardCache();

    return { status: 'success', message: '商談情報を更新しました。' };
  } catch (e) {
    console.error('Failed to update meeting:', e);
    throw new Error('サーバーエラー: 商談の更新に失敗しました。 ' + e.message);
  }
}
/**
 * 承認された単価申請の内容を「単価マスタ」に反映します。
 * * @param {string} customerId - 対象の得意先コード
 * @param {object} appDetails - getApplicationDetails で取得した申請詳細オブジェクト
 * @param {Date | null} effectiveDate - 登録有効日 (Dateオブジェクトまたはnull)
 */
function updatePriceInMaster(customerId, appDetails, effectiveDate) {
  try {
    if (!TANKA_SHEET) {
      throw new Error('「単価マスタ」シートが見つかりません。');
    }
    if (!appDetails || !appDetails.detailData || appDetails.detailData.length === 0) {
      throw new Error('申請詳細データ (detailData) が空です。');
    }

    const sheet = TANKA_SHEET;
    const values = sheet.getDataRange().getValues();
    const header = values.shift(); // ヘッダー行

    // --- 1. ヘッダーインデックスのマップを作成 ---
    const colIndex = header.reduce((map, col, i) => {
      if (col) map[col.trim()] = i;
      return map;
    }, {});

    // 必須列の確認
    const requiredCols = ['得意先コード', '商品コード'];
    for (const col of requiredCols) {
      if (colIndex[col] === undefined) {
        throw new Error(`「単価マスタ」シートに必要な列 "${col}" が見つかりません。`);
      }
    }

    // --- 2. 既存の単価データをマップ化 (商品コード -> シート上の行インデックス) ---
    //    (valuesのインデックスは0始まり, +2でシート上の行番号)
    const existingPriceMap = {};
    values.forEach((row, index) => {
      const rowCustomerId = row[colIndex['得意先コード']];
      const rowProductCode = row[colIndex['商品コード']];
      
      // 今回処理する得意先のデータのみをマップ化
      if (String(rowCustomerId) === String(customerId) && rowProductCode) {
        existingPriceMap[String(rowProductCode)] = index; // values 配列上でのインデックス
      }
    });

    // --- 3. 変更を適用 (追加/修正/削除) ---
    const rowsToDelete = []; // 削除対象の「シート上の行番号」
    const rowsToAdd = [];    // 追加対象のデータ配列
    const rowsToUpdate = {}; // 更新対象のデータマップ (シート上の行番号 -> データ配列)

    appDetails.detailData.forEach(item => {
      const regStatus = item['登録区分'];
      const productCode = item['商品コード'];
      if (!productCode) {
        console.warn('商品コードがない申請データはスキップされました:', item);
        return;
      }

      const existingRowIndex = existingPriceMap[String(productCode)]; // values 配列上のインデックス
      const sheetRowNumber = existingRowIndex !== undefined ? existingRowIndex + 2 : null; // シート上の実行番号

      switch (regStatus) {
        case '追加':
          if (existingRowIndex !== undefined) {
            // マスタに既に存在するが「追加」申請が来た場合 (＝実質「修正」扱い)
            console.warn(`単価マスタ反映: ${customerId} / ${productCode} は既に存在するため、追加ではなく更新します。`);
            const newRowData = createPriceRow(header, colIndex, item, customerId, effectiveDate, 'after');
            rowsToUpdate[sheetRowNumber] = newRowData;
          } else {
            // マスタに存在しない「追加」
            const newRowData = createPriceRow(header, colIndex, item, customerId, effectiveDate, 'after');
            rowsToAdd.push(newRowData);
          }
          break;

        case '修正':
          if (sheetRowNumber) {
            // マスタに存在する「修正」
            const newRowData = createPriceRow(header, colIndex, item, customerId, effectiveDate, 'after');
            rowsToUpdate[sheetRowNumber] = newRowData;
          } else {
            // マスタに存在しないが「修正」申請が来た場合 (＝実質「追加」扱い)
            console.warn(`単価マスタ反映: ${customerId} / ${productCode} がマスタにないため、修正ではなく追加します。`);
            const newRowData = createPriceRow(header, colIndex, item, customerId, effectiveDate, 'after');
            rowsToAdd.push(newRowData);
          }
          break;

        case '削除':
          if (sheetRowNumber) {
            // マスタに存在する「削除」
            rowsToDelete.push(sheetRowNumber);
          } else {
            // マスタに存在しない「削除」(何もしない)
            console.warn(`単価マスタ反映: ${customerId} / ${productCode} はマスタにないため、削除をスキップしました。`);
          }
          break;
          
        default:
          console.warn(`不明な登録区分: ${regStatus}`);
      }
    });

    // --- 4. スプレッドシートへの書き込み実行 ---

    // (A) 更新 (Update)
    //  ※ getRangeList/setValues を使うと高速だが、今回は1行ずつsetする
    for (const [rowNum, data] of Object.entries(rowsToUpdate)) {
      sheet.getRange(Number(rowNum), 1, 1, header.length).setValues([data]);
    }

    // (B) 削除 (Delete)
    //   ※行番号が大きい順（降順）にソートしてから削除する（行がずれないように）
    if (rowsToDelete.length > 0) {
      rowsToDelete.sort((a, b) => b - a); // 降順ソート
      rowsToDelete.forEach(rowNum => {
        sheet.deleteRow(rowNum);
      });
    }

    // (C) 追加 (Add)
    if (rowsToAdd.length > 0) {
      sheet.getRange(sheet.getLastRow() + 1, 1, rowsToAdd.length, header.length).setValues(rowsToAdd);
    }
    
    console.log(`単価マスタを更新しました (顧客ID: ${customerId}): 追加 ${rowsToAdd.length}件, 更新 ${Object.keys(rowsToUpdate).length}件, 削除 ${rowsToDelete.length}件`);

  } catch (e) {
    console.error(`単価マスタの反映に失敗しました (顧客ID: ${customerId}):`, e);
    // ここでエラーを投げ直すか、管理者に通知する
    throw new Error(`単価マスタの反映に失敗しました: ${e.message}`);
  }
}

/**
 * 「単価マスタ」のヘッダー配列と申請データ項目から、シートに書き込むための1行分の配列を作成します。
 * * @param {string[]} header - 単価マスタのヘッダー配列 (例: ['得意先コード', '商品コード', ...])
 * @param {object} colIndex - ヘッダー名と列インデックスのマップ (例: {'得意先コード': 0, ...})
 * @param {object} item - 申請データ_単価 の1行分のデータ (例: {'登録区分': '追加', '商品名_修正後': '...', ...})
 * @param {string} customerId - 得意先コード
 * @param {Date | null} effectiveDate - 登録有効日
 * @param {'before' | 'after'} type - 'before' (修正前) または 'after' (修正後) のどちらのデータを参照するか
 * @returns {Array<any>} スプレッドシートに書き込むための1行分の配列
 */
function createPriceRow(header, colIndex, item, customerId, effectiveDate, type) {
  const newRow = new Array(header.length).fill('');
  
  // 'before' (修正前) または 'after' (修正後) のキー接尾辞を決定
  const suffix = (type === 'after') ? '_修正後' : '_修正前';

  // 「申請データ_単価」シートの列名と、「単価マスタ」の列名のマッピング
  // (キー: 単価マスタ列名, 値: 申請データ列名の接頭辞)
  const keyMap = {
    '商品名': '商品名',
    '裁断方法コード': '裁断方法コード',
    '裁断方法名': '裁断方法名',
    '袋詰方法コード': '袋詰方法コード',
    '袋詰方法名': '袋詰方法名',
    '卸価格': '卸価格',
    'バラ単価': '実際販売価格', // 申請データの「実際販売価格」→単価マスタの「バラ単価」
    '掛率': '掛率',
    '粗利率': '粗利率'
  };

  // ヘッダー配列をループして、新しい行の各セルに値を設定
  header.forEach((colName, index) => {
    switch (colName) {
      case '得意先コード':
        // 0落ち対策: シングルクォートを付ける
        newRow[index] = customerId ? "'" + customerId : '';
        break;
      case '商品コード':
        // 0落ち対策: シングルクォートを付ける
        newRow[index] = item['商品コード'] ? "'" + item['商品コード'] : '';
        break;
      case '登録有効日': // 単価マスタにこの列があれば設定
        if (effectiveDate) {
          newRow[index] = effectiveDate;
        }
        break;
      case '裁断方法コード':
      case '袋詰方法コード':
        // コード列の0落ち対策
        const codeKey = keyMap[colName] + suffix;
        const codeValue = item[codeKey];
        newRow[index] = codeValue ? "'" + codeValue : '';
        break;
      default:
        // keyMap に基づいて値を設定
        if (keyMap[colName]) {
          const itemKey = keyMap[colName] + suffix; // 例: '商品名' + '_修正後' -> '商品名_修正後'
          newRow[index] = item[itemKey] !== undefined ? item[itemKey] : '';
        }
        // マップにない列 (例: '備考' など) は空文字のまま
    }
  });

  return newRow;
}

/**
 * 申請データを申請管理シートと申請データシート（顧客or単価）に書き込みます。
 * [修正] 申請管理シートへの書き込み時に、フォームからの申請者名・メールを反映。
 * [修正] ★単価申請の詳細データ書き込みロジックを実装。
 * [★今回の修正★] 「登録有効日」を日付型(yyyy/MM/dd)で書き込みます。
 */
function addApplication(appData) {
  try {
    // --- シート取得とヘッダー確認 ---
    const appSheet = SPREADSHEET.getSheetByName('申請管理');
    if (!appSheet) throw new Error('申請管理シートが見つかりません。');

    const customerDataSheet = getOrCreateSheet('申請データ_顧客', ['詳細ID', '申請ID', '項目名', '修正前の値', '修正後の値']);
    if (!customerDataSheet) {
      throw new Error('申請データ_顧客シートの取得または作成に失敗しました。');
    }

    // ★ 修正: ご提示いただいたヘッダー定義に変更
    const priceDataHeaders = [
      '詳細ID', '申請ID', '登録区分', '商品コード',
      '商品名_修正前', '商品名_修正後', '裁断方法コード_修正前', '裁断方法コード_修正後',
      '裁断方法名_修正前', '裁断方法名_修正後', '袋詰方法コード_修正前', '袋詰方法コード_修正後',
      '袋詰方法名_修正前', '袋詰方法名_修正後', '卸価格_修正前', '卸価格_修正後',
      '実際販売価格_修正前', '実際販売価格_修正後', '掛率_修正前', '掛率_修正後',
      '粗利率_修正前', '粗利率_修正後'
    ];
    const priceDataSheet = getOrCreateSheet('申請データ_単価', priceDataHeaders);
    if (!priceDataSheet) {
      throw new Error('申請データ_単価シートの取得または作成に失敗しました。');
    }

    // --- 申請IDの採番 ---
    const lastAppRow = appSheet.getLastRow();
    const appRange = appSheet.getRange(2, 1, lastAppRow > 1 ? lastAppRow - 1 : 1, 1);
    const existingAppIds = appRange.getValues().flat().map(Number).filter(n => n > 0);
    const newAppId = existingAppIds.length > 0 ? Math.max(...existingAppIds) + 1 : 1;

    // --- 申請管理シートへの書き込み ---
    const lastCol = appSheet.getLastColumn();
    if (lastCol === 0) {
      throw new Error('申請管理シートにヘッダー行が見つかりません。シートを確認してください。');
    }

    const appHeaderRange = appSheet.getRange(1, 1, 1, lastCol).getValues();
    if (!appHeaderRange || appHeaderRange.length === 0 || !appHeaderRange[0]) {
      throw new Error('申請管理シートのヘッダー行が取得できません。シートを確認してください。');
    }

    const appHeader = appHeaderRange[0];
    if (!appHeader || appHeader.length === 0) {
      throw new Error('申請管理シートのヘッダー行が空です。シートを確認してください。');
    }
    const appHeaderMap = appHeader.reduce((map, header, index) => {
      map[header] = index;
      return map;
    }, {});

    const now = new Date();
    const isPriceApp = appData.type === '商品登録修正';
    // ★ 修正: 新規顧客登録+単価の場合を検出
    const hasCustomerAndPrice = appData.type === '顧客新規登録' && appData.payload?.priceData?.length > 0;
    const targetMaster = isPriceApp ? '単価マスタ' :
                         hasCustomerAndPrice ? '顧客マスタ,単価マスタ' : '顧客マスタ';

    // ★ 修正: 申請者情報を payload から取得
    const applicantData = isPriceApp ? (appData.payload || {}) : (appData.payload?.newData || {});

    let targetCustomerId = '';
    if (isPriceApp) {
      // 単価申請の場合 (appData.customerId は submitApplication で設定)
      targetCustomerId = appData.customerId || '';
    } else if (appData.type === '顧客情報修正') {
      // 顧客修正の場合 (originalData から取得)
      targetCustomerId = appData.payload?.originalData?.['得意先コード'] || appData.payload?.newData?.['得意先コード'] || '';
    } else { // 新規登録
      // 新規登録の場合 (newData から取得)
      targetCustomerId = appData.payload?.newData?.['得意先コード'] || '';
    }

    const newAppRowData = new Array(appHeader.length).fill('');

    newAppRowData[appHeaderMap['申請ID']] = newAppId;
    newAppRowData[appHeaderMap['申請日時']] = now;
    newAppRowData[appHeaderMap['対象マスタ']] = targetMaster;
    newAppRowData[appHeaderMap['申請種別']] = appData.type;
    newAppRowData[appHeaderMap['対象顧客名']] = appData.customerName;

    if (appHeaderMap['得意先コード'] !== undefined) {
      // 先頭にシングルクォート(')を付けて、スプレッドシートに強制的に文字列として認識させる
      newAppRowData[appHeaderMap['得意先コード']] = targetCustomerId ? "'" + targetCustomerId : '';
    } else if (targetCustomerId) {
      console.warn('「申請管理」シートに「得意先コード」列が見つからないため、書き込みをスキップしました。');
    }

    newAppRowData[appHeaderMap['申請者名']] = applicantData['applicantName'] || (isPriceApp ? '' : applicantData['申請者名']);
    newAppRowData[appHeaderMap['申請者メール']] = applicantData['applicantEmail'] || (isPriceApp ? '' : applicantData['申請者メール']);

    // ★ 追加: 申請者IDを保存（ログインユーザーのIDを取得）
    const loginUser = getCurrentUser();
    if (loginUser && appHeaderMap['申請者ID'] !== undefined) {
      newAppRowData[appHeaderMap['申請者ID']] = loginUser.id;
    }

    // ★ 修正: 申請者が上長の場合、承認段階を「上長承認済」にする
    const applicantRole = loginUser ? determineUserRole(loginUser) : 'applicant';
    let initialStage = '申請中';
    let initialStatus = '申請中';

    if (applicantRole === 'supervisor') {
      // 上長が自分で申請した場合、上長承認をスキップして管理部門から承認開始
      initialStage = '上長承認済';
      initialStatus = '承認中';
    } else if (applicantRole === 'manager') {
      // 管理部門が自分で申請した場合、管理部門承認をスキップして常務から承認開始
      initialStage = '管理承認済';
      initialStatus = '承認中';
    } else if (applicantRole === 'division_manager') {
      // 常務が自分で申請した場合、常務承認をスキップして決裁者から承認開始
      initialStage = '常務承認済';
      initialStatus = '承認中';
    } else if (applicantRole === 'approver') {
      // 決裁者（社長）が自分で申請した場合、2段階前（管理承認済）からスタート→常務が次の承認者
      initialStage = '管理承認済';
      initialStatus = '承認中';
    }

    newAppRowData[appHeaderMap['ステータス']] = initialStatus;
    newAppRowData[appHeaderMap['承認段階']] = initialStage;

    // ★ 追加: 申請グループIDを設定（初回は申請IDと同じ）
    if (appHeaderMap['申請グループID'] !== undefined) {
      newAppRowData[appHeaderMap['申請グループID']] = newAppId;
    }

    // ★ 追加: 再申請回数を0に設定
    if (appHeaderMap['再申請回数'] !== undefined) {
      newAppRowData[appHeaderMap['再申請回数']] = 0;
    }
    
    // 対象価格 (単価申請時のみ) または 新規顧客登録+単価の場合
    if ((isPriceApp || hasCustomerAndPrice) && (applicantData['targetPrice'] || appData.payload?.targetPrice)) {
        const targetPrice = applicantData['targetPrice'] || appData.payload?.targetPrice;
        if (targetPrice === 'current') {
            // 現行の価格の場合は「現行」と記録
            newAppRowData[appHeaderMap['登録有効日']] = '現行';
        } else {
            try {
               const effectiveDate = new Date(targetPrice);
               if (!isNaN(effectiveDate)) {
                   // yyyy/MM/dd 形式の「日付」として設定 (時刻を含めない)
                   const dateStr = Utilities.formatDate(effectiveDate, SPREADSHEET.getSpreadsheetTimeZone(), "yyyy/MM/dd");
                   newAppRowData[appHeaderMap['登録有効日']] = new Date(dateStr); // 日付オブジェクトとして再設定
               }
            } catch(e) {
              console.warn('対象価格の日付フォーマット変換に失敗しました:', targetPrice, e.message);
            }
        }
    }

    // ★ 承認権限を持つ申請者の場合、スキップされた承認段階をappendRow前にnewAppRowDataへ直接書き込む
    if (loginUser && applicantRole !== 'applicant') {
      const nowStr = Utilities.formatDate(now, Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");
      const approverName = loginUser.name || loginUser.id;
      // approver（社長）は管理部門欄に記録、それ以外は自分の段階の欄に記録
      const stagePrefix = applicantRole === 'supervisor'       ? '上長'
                        : applicantRole === 'manager'          ? '管理'
                        : applicantRole === 'division_manager' ? '常務'
                        : applicantRole === 'approver'         ? '管理'  // 社長は管理部門欄
                        : null;
      if (stagePrefix) {
        const judgeColIdx    = appHeaderMap[`${stagePrefix}判断者名`];
        const timeColIdx     = appHeaderMap[`${stagePrefix}判断時刻`];
        const approvalColIdx = appHeaderMap[`${stagePrefix}承認`];
        if (judgeColIdx    !== undefined) newAppRowData[judgeColIdx]    = approverName;
        if (timeColIdx     !== undefined) newAppRowData[timeColIdx]     = nowStr;
        if (approvalColIdx !== undefined) newAppRowData[approvalColIdx] = '承認済み';
      }
    }

    appSheet.appendRow(newAppRowData);

    // --- 申請データシートへの書き込み ---
    let targetDataSheet = null;
    let dataSheetHeader = null;
    let dataRowsToAdd = [];
    let detailIdCounter = 0;

    // 詳細IDの採番 (両シート共通で使えるように外に出す)
    const getLastDetailId = (sheet) => {
        const lastRow = sheet.getLastRow();
        if (lastRow <= 1) return 0;
        const range = sheet.getRange(2, 1, lastRow - 1, 1);
        const ids = range.getValues().flat().map(Number).filter(n => n > 0);
        return ids.length > 0 ? Math.max(...ids) : 0;
    };

    if (targetMaster === '単価マスタ') {
        targetDataSheet = priceDataSheet;

        // ★ 修正: 実際のシートからヘッダーを取得（顧客マスタと同様の処理）
        const lastCol = targetDataSheet.getLastColumn();
        if (lastCol === 0) {
          throw new Error('申請データ_単価シートにヘッダーが見つかりません。');
        }
        const headerRange = targetDataSheet.getRange(1, 1, 1, lastCol).getValues();
        if (!headerRange || headerRange.length === 0 || !headerRange[0]) {
          throw new Error('申請データ_単価シートのヘッダー行が取得できません。');
        }
        dataSheetHeader = headerRange[0];

        const headerMap = dataSheetHeader.reduce((map, h, i) => { map[h] = i; return map; }, {});

        const pricePayload = appData.payload; // { added: [], modified: [], deleted: [], ... }
        
        detailIdCounter = getLastDetailId(targetDataSheet);

        // ヘルパー関数: コードから名称を取得
        const getSaitanName = (code) => code ? (SAITAN_MASTER[code] || '') : '';
        const getFukuroName = (code) => code ? (FUKURO_MASTER[code] || '') : '';

        // 1. 追加 (added)
        (pricePayload.added || []).forEach(item => {
            detailIdCounter++;
            const newRow = new Array(dataSheetHeader.length).fill('');
            newRow[headerMap['詳細ID']] = detailIdCounter;
            newRow[headerMap['申請ID']] = newAppId;
            newRow[headerMap['登録区分']] = '追加';
            // 0落ち対策: 商品コードに'を付与
            newRow[headerMap['商品コード']] = item.productCode ? "'" + item.productCode : '';

            newRow[headerMap['商品名_修正後']] = item.productName;
            // 0落ち対策: 裁断方法コード/袋詰方法コードに'を付与
            newRow[headerMap['裁断方法コード_修正後']] = item.saitanCode ? "'" + item.saitanCode : '';
            newRow[headerMap['裁断方法名_修正後']] = getSaitanName(item.saitanCode);
            newRow[headerMap['袋詰方法コード_修正後']] = item.fukuroCode ? "'" + item.fukuroCode : '';
            newRow[headerMap['袋詰方法名_修正後']] = getFukuroName(item.fukuroCode);
            newRow[headerMap['卸価格_修正後']] = item.oroshi;
            newRow[headerMap['実際販売価格_修正後']] = item.jissai;
            newRow[headerMap['掛率_修正後']] = item.kakeritsu;
            newRow[headerMap['粗利率_修正後']] = item.arari;

            dataRowsToAdd.push(newRow);
        });

        // 2. 修正 (modified)
        (pricePayload.modified || []).forEach(item => {
            detailIdCounter++;
            const org = item.originalData;
            const mod = item.newData;
            const newRow = new Array(dataSheetHeader.length).fill('');

            newRow[headerMap['詳細ID']] = detailIdCounter;
            newRow[headerMap['申請ID']] = newAppId;
            newRow[headerMap['登録区分']] = '修正';
            // 0落ち対策: 商品コードに'を付与
            newRow[headerMap['商品コード']] = mod.productCode ? "'" + mod.productCode : ''; // (コードは変更不可前提)

            // 修正前 - 0落ち対策
            newRow[headerMap['商品名_修正前']] = org.productName;
            newRow[headerMap['裁断方法コード_修正前']] = org.saitanCode ? "'" + org.saitanCode : '';
            newRow[headerMap['裁断方法名_修正前']] = getSaitanName(org.saitanCode);
            newRow[headerMap['袋詰方法コード_修正前']] = org.fukuroCode ? "'" + org.fukuroCode : '';
            newRow[headerMap['袋詰方法名_修正前']] = getFukuroName(org.fukuroCode);
            newRow[headerMap['卸価格_修正前']] = org.oroshi;
            newRow[headerMap['実際販売価格_修正前']] = org.jissai;
            newRow[headerMap['掛率_修正前']] = org.kakeritsu;
            newRow[headerMap['粗利率_修正前']] = org.arari;

            // 修正後 - 0落ち対策
            newRow[headerMap['商品名_修正後']] = mod.productName;
            newRow[headerMap['裁断方法コード_修正後']] = mod.saitanCode ? "'" + mod.saitanCode : '';
            newRow[headerMap['裁断方法名_修正後']] = getSaitanName(mod.saitanCode);
            newRow[headerMap['袋詰方法コード_修正後']] = mod.fukuroCode ? "'" + mod.fukuroCode : '';
            newRow[headerMap['袋詰方法名_修正後']] = getFukuroName(mod.fukuroCode);
            newRow[headerMap['卸価格_修正後']] = mod.oroshi;
            newRow[headerMap['実際販売価格_修正後']] = mod.jissai;
            newRow[headerMap['掛率_修正後']] = mod.kakeritsu;
            newRow[headerMap['粗利率_修正後']] = mod.arari;

            dataRowsToAdd.push(newRow);
        });

        // 3. 削除 (deleted)
        (pricePayload.deleted || []).forEach(item => {
            detailIdCounter++;
            const newRow = new Array(dataSheetHeader.length).fill('');
            newRow[headerMap['詳細ID']] = detailIdCounter;
            newRow[headerMap['申請ID']] = newAppId;
            newRow[headerMap['登録区分']] = '削除';
            // 0落ち対策: 商品コードに'を付与
            newRow[headerMap['商品コード']] = item.productCode ? "'" + item.productCode : '';

            // 修正前 (削除されるデータ) - 0落ち対策
            newRow[headerMap['商品名_修正前']] = item.productName;
            newRow[headerMap['裁断方法コード_修正前']] = item.saitanCode ? "'" + item.saitanCode : '';
            newRow[headerMap['裁断方法名_修正前']] = getSaitanName(item.saitanCode);
            newRow[headerMap['袋詰方法コード_修正前']] = item.fukuroCode ? "'" + item.fukuroCode : '';
            newRow[headerMap['袋詰方法名_修正前']] = getFukuroName(item.fukuroCode);
            newRow[headerMap['卸価格_修正前']] = item.oroshi;
            newRow[headerMap['実際販売価格_修正前']] = item.jissai;
            newRow[headerMap['掛率_修正前']] = item.kakeritsu;
            newRow[headerMap['粗利率_修正前']] = item.arari;

            dataRowsToAdd.push(newRow);
        });

    } else { // targetMaster === '顧客マスタ'
        targetDataSheet = customerDataSheet;

        // ★ 修正: 安全にヘッダーを取得
        const lastCol = targetDataSheet.getLastColumn();
        if (lastCol === 0) {
          throw new Error('申請データ_顧客シートにヘッダーが見つかりません。');
        }
        const headerRange = targetDataSheet.getRange(1, 1, 1, lastCol).getValues();
        if (!headerRange || headerRange.length === 0 || !headerRange[0]) {
          throw new Error('申請データ_顧客シートのヘッダー行が取得できません。');
        }
        dataSheetHeader = headerRange[0];

        const { newData, originalData } = appData.payload;

        detailIdCounter = getLastDetailId(targetDataSheet);

        // 0落ち対策が必要なコード列のリスト
        const codeColumns = ['得意先コード', '得意先グループコード', '請求先コード', '営業担当者コード'];

        const allKeys = new Set([...Object.keys(originalData || {}), ...Object.keys(newData || {})]);
        allKeys.forEach(key => {
            // ★ 修正: 申請者情報は詳細データには含めない
            if (['formMode', 'originalCustomerId', '申請者名', '申請者メール', '削除フラグ', 'applicantName', 'applicantEmail', 'effectiveDate'].includes(key)) return;

            let oldValue = originalData ? (originalData[key] || '') : '';
            let newValue = newData ? (newData[key] || '') : '';

            // 0落ち対策: コード列の場合は'を付与
            if (codeColumns.includes(key)) {
                if (oldValue && String(oldValue).trim() !== '') {
                    oldValue = "'" + String(oldValue).replace(/^'/, '').trim();
                }
                if (newValue && String(newValue).trim() !== '') {
                    newValue = "'" + String(newValue).replace(/^'/, '').trim();
                }
            }

            // 修正申請でも値がある項目はすべて保存（却下時の指摘項目選択で全項目を表示するため）
            if (String(oldValue) !== '' || String(newValue) !== '') {
                detailIdCounter++;
                // ヘッダー '詳細ID', '申請ID', '項目名', '修正前の値', '修正後の値' に合わせる
                dataRowsToAdd.push([detailIdCounter, newAppId, key, oldValue, newValue]);
            }
        });
    }

    if (dataRowsToAdd.length > 0 && targetDataSheet) {
      if (!targetDataSheet) {
        throw new Error('詳細データシートが見つかりません。');
      }
      if (!dataSheetHeader || dataSheetHeader.length === 0) {
        throw new Error('詳細データシートのヘッダー情報が取得できませんでした。');
      }

      // データ行の列数チェック
      if (dataRowsToAdd[0] && dataRowsToAdd[0].length !== dataSheetHeader.length) {
        throw new Error(`列数が不一致です。期待: ${dataSheetHeader.length}列, 実際: ${dataRowsToAdd[0].length}列。ヘッダー: [${dataSheetHeader.join(', ')}]`);
      }

      targetDataSheet.getRange(targetDataSheet.getLastRow() + 1, 1, dataRowsToAdd.length, dataSheetHeader.length).setValues(dataRowsToAdd);
    }

    // ★ 追加: 新規顧客登録+単価の場合、単価データも書き込む
    if (hasCustomerAndPrice) {
      const priceData = appData.payload.priceData; // フロントエンドから送られてきた単価データ配列

      // 単価データシートのヘッダー取得
      const lastCol = priceDataSheet.getLastColumn();
      if (lastCol === 0) {
        throw new Error('申請データ_単価シートにヘッダーが見つかりません。');
      }
      const priceHeaderRange = priceDataSheet.getRange(1, 1, 1, lastCol).getValues();
      if (!priceHeaderRange || priceHeaderRange.length === 0 || !priceHeaderRange[0]) {
        throw new Error('申請データ_単価シートのヘッダー行が取得できません。');
      }
      const priceHeader = priceHeaderRange[0];
      const priceHeaderMap = priceHeader.reduce((map, h, i) => { map[h] = i; return map; }, {});

      // 詳細IDの採番
      let priceDetailIdCounter = getLastDetailId(priceDataSheet);

      // ヘルパー関数: コードから名称を取得
      const getSaitanName = (code) => code ? (SAITAN_MASTER[code] || '') : '';
      const getFukuroName = (code) => code ? (FUKURO_MASTER[code] || '') : '';

      // 単価データ行を作成
      const priceRowsToAdd = [];
      priceData.forEach(item => {
        priceDetailIdCounter++;
        const newRow = new Array(priceHeader.length).fill('');
        newRow[priceHeaderMap['詳細ID']] = priceDetailIdCounter;
        newRow[priceHeaderMap['申請ID']] = newAppId;
        newRow[priceHeaderMap['登録区分']] = '追加'; // 新規登録なので常に「追加」
        // 0落ち対策: 商品コードに'を付与
        newRow[priceHeaderMap['商品コード']] = item.productCode ? "'" + item.productCode : '';

        // 新規なので「修正後」のみ設定 - 0落ち対策
        newRow[priceHeaderMap['商品名_修正後']] = item.productName;
        newRow[priceHeaderMap['裁断方法コード_修正後']] = item.saitanCode ? "'" + item.saitanCode : '';
        newRow[priceHeaderMap['裁断方法名_修正後']] = getSaitanName(item.saitanCode);
        newRow[priceHeaderMap['袋詰方法コード_修正後']] = item.fukuroCode ? "'" + item.fukuroCode : '';
        newRow[priceHeaderMap['袋詰方法名_修正後']] = getFukuroName(item.fukuroCode);
        newRow[priceHeaderMap['卸価格_修正後']] = item.oroshi;
        newRow[priceHeaderMap['実際販売価格_修正後']] = item.jissai;
        newRow[priceHeaderMap['掛率_修正後']] = item.kakeritsu;
        newRow[priceHeaderMap['粗利率_修正後']] = item.arari;

        priceRowsToAdd.push(newRow);
      });

      // 単価データシートに書き込み
      if (priceRowsToAdd.length > 0) {
        priceDataSheet.getRange(priceDataSheet.getLastRow() + 1, 1, priceRowsToAdd.length, priceHeader.length).setValues(priceRowsToAdd);
      }
    }

    return { status: 'success', message: '申請が送信されました。', newApplicationId: newAppId };
  } catch(e) {
    console.error('Failed to submit application:', e, e.stack); // ★ エラー詳細をログ出力
    throw new Error('サーバーエラー: 申請の送信に失敗しました。' + e.message);
  }
}

/**
 * 顧客基本情報を取得します。課税方式名称も含めるように修正。
 * 日付は yyyy-MM-dd 形式で返すように修正。
 */
function getCustomerBasicInfo(customerId) {
  // 得意先マスタシートの存在確認
  if (!CUSTOMER_SHEET) {
      throw new Error("「得意先マスタ」シートが見つかりません。");
  }
  const values = CUSTOMER_SHEET.getDataRange().getValues();
  if (values.length <= 1) return {}; // ヘッダーのみ or 空

  const header = values.shift(); // ヘッダー行
  const idColIndex = header.indexOf('得意先コード');

  // 得意先コード列の存在確認
  if (idColIndex === -1) {
      console.error("Header '得意先コード' not found in CUSTOMER_SHEET.");
      throw new Error('ヘッダー「得意先コード」が見つかりません。');
  }

  // 該当する顧客データを検索（シングルクォートを除去して比較）
  const cleanCustomerId = cleanSingleQuotes(customerId);
  console.log('[getCustomerBasicInfo] 検索する顧客ID:', cleanCustomerId, '型:', typeof cleanCustomerId);
  const row = values.find(r => {
    const cellId = cleanSingleQuotes(r[idColIndex]);
    // 文字列として比較、または数値として比較（先頭の0を無視）
    return cellId == cleanCustomerId ||
           String(cellId) === String(cleanCustomerId) ||
           parseInt(cellId, 10) === parseInt(cleanCustomerId, 10);
  });
  if (!row) {
    console.log('[getCustomerBasicInfo] 顧客が見つかりません。ID:', cleanCustomerId);
    console.log('[getCustomerBasicInfo] マスタの最初の5件のID:', values.slice(0, 5).map(r => cleanSingleQuotes(r[idColIndex])));
    return {}; // 見つからなければ空オブジェクト
  }
  console.log('[getCustomerBasicInfo] 顧客が見つかりました。ID:', cleanCustomerId);

  const customerData = {};
  // ヘッダーに基づいてデータをオブジェクトに格納
  header.forEach((key, index) => {
    if (!key) return; // 空のヘッダーは無視
    const value = row[index]; // 対応するセルの値

    if (value instanceof Date && !isNaN(value)) {
      // 日付の場合は yyyy-MM-dd 形式で返す (JSの日付入力欄 value に合わせる)
      try {
           customerData[key.trim()] = Utilities.formatDate(value, Session.getScriptTimeZone(), "yyyy-MM-dd");
      } catch(e) {
           console.warn(`日付フォーマットエラー: ${key} = ${value}`);
           customerData[key.trim()] = value; // フォーマットエラー時は元の値
      }
    } else {
      // 日付以外はそのまま格納 (trimで前後の空白除去)
      customerData[key.trim()] = (value !== undefined && value !== null) ? String(value).trim() : '';
    }
  });

   // ★ 課税方式名称をここで取得 (列が存在するか確認)
   const taxTypeColIndex = header.indexOf('課税方式名称');
   if(taxTypeColIndex !== -1) {
       // customerData['課税方式名称'] = row[taxTypeColIndex]; // 上のforEachで既に入っているはず
       // 念のため、値が空でないか確認
       if (!customerData['課税方式名称']) {
           console.warn(`顧客 ${customerId} の '課税方式名称' が空欄です。`);
       }
   } else {
       customerData['課税方式名称'] = ''; // 列がなければ空文字
       console.warn("得意先マスタに '課税方式名称' 列が見つかりません。");
   }

  return customerData;
}

function getMeetingsByCustomerId(customerId, filterHandlerCode) {
  const meetingSheet = SPREADSHEET.getSheetByName('商談管理');
  if (!meetingSheet) return [];

  const values = meetingSheet.getDataRange().getValues();
  if (values.length <= 1) return [];
  const header = values.shift();

  const colIndex = header.reduce((acc, col, i) => (acc[col] = i, acc), {});

  const cleanHandlerCode = filterHandlerCode ? cleanSingleQuotes(String(filterHandlerCode).trim()) : null;

  // ランクを得意先マスタから取得
  let customerRank = '-';
  try {
    if (CUSTOMER_SHEET) {
      const custValues = CUSTOMER_SHEET.getDataRange().getValues();
      const custHeader = custValues.shift();
      const custIdCol = custHeader.indexOf('得意先コード');
      const custRankCol = custHeader.indexOf('得意先ランク区分名称');
      if (custIdCol !== -1 && custRankCol !== -1) {
        const cleanId = cleanSingleQuotes(String(customerId));
        const custRow = custValues.find(r => cleanSingleQuotes(String(r[custIdCol] || '')) === cleanId);
        if (custRow) customerRank = custRow[custRankCol] || '-';
      }
    }
  } catch (e) {
    console.warn('[getMeetingsByCustomerId] ランク取得エラー:', e);
  }

  // シングルクォートを除去して比較
  const cleanCustomerId = cleanSingleQuotes(customerId);
  const today = new Date();
  today.setHours(0, 0, 0, 0);

  const meetings = values
    .filter(row => {
      const cellId = cleanSingleQuotes(String(row[colIndex['得意先コード']] || ''));
      if (cellId != cleanCustomerId) return false;
      if (!(row[colIndex['商談実施日']] || row[colIndex['商談予定日']])) return false;
      if (cleanHandlerCode && colIndex['担当者コード'] !== undefined) {
        const meetingCode = cleanSingleQuotes(String(row[colIndex['担当者コード']] || '').trim());
        return meetingCode !== '' && meetingCode === cleanHandlerCode;
      }
      return true;
    })
    .map(row => {
      const scheduleDateVal = (() => {
        const v = row[colIndex['商談予定日']];
        if (!v) return null;
        const d = new Date(v);
        return (isNaN(d.getTime()) || d.getFullYear() < 1990) ? null : d;
      })();
      const actualDateVal = (() => {
        const v = row[colIndex['商談実施日']];
        if (!v) return null;
        const d = new Date(v);
        return (isNaN(d.getTime()) || d.getFullYear() < 1990) ? null : d;
      })();
      const timestamp = row[colIndex['タイムスタンプ']];

      let status;
      if (actualDateVal) {
        status = 'completed';
      } else if (scheduleDateVal && scheduleDateVal < today) {
        status = 'overdue';
      } else {
        status = 'scheduled';
      }

      return {
        id: row[colIndex['ID']] || '',
        handler: row[colIndex['担当者']] || '',
        handlerCode: colIndex['担当者コード'] !== undefined ? String(row[colIndex['担当者コード']] || '').trim() : '',
        date: timestamp ? Utilities.formatDate(new Date(timestamp), "JST", "yyyy/MM/dd HH:mm") : '',
        scheduleDate: scheduleDateVal ? Utilities.formatDate(scheduleDateVal, "JST", "yyyy-MM-dd") : '',
        actualDate: actualDateVal ? Utilities.formatDate(actualDateVal, "JST", "yyyy-MM-dd") : '',
        customerName: row[colIndex['企業名']] || '',
        rank: customerRank,
        purpose: row[colIndex['商談目的']] || '',
        result: row[colIndex['結果']] || '',
        notes: row[colIndex['実績備考']] || '',
        status: status
      };
    })
    .sort((a, b) => (b.scheduleDate || b.actualDate || '').localeCompare(a.scheduleDate || a.actualDate || ''));

  return meetings;
}

/**
 * 指定された得意先コードの商品登録情報を「単価マスタ」シートから取得します。
 * [修正] 新しい列構成（卸価格, 実際販売価格, 掛率, 粗利率）を読み込むように修正。
 * [★今回の修正★] 顧客詳細画面表示のため、「裁断方法名」「袋詰方法名」も読み込むように修正。
 */
function getPricesByCustomerId(customerId, customerBasicInfo) {
  try {
    // 単価マスタシートの存在確認
    if (!TANKA_SHEET) {
      console.warn('「単価マスタ」シートが見つかりません。');
      return [];
    }

    // 1. 顧客の課税方式を取得
    // basicInfoが渡されていない場合のみ取得
    const customerBasic = customerBasicInfo || getCustomerBasicInfo(customerId);
    if (!customerBasic) {
      console.warn(`顧客情報が見つかりません: ${customerId}`);
      return [];
    }
    const taxType = customerBasic['課税方式名称'] || '内税'; // デフォルトは内税
    const useTaxInPrice = (taxType === '内税'); // 内税なら税込価格を使用

    // 2. 単価マスタから顧客の商品データを取得
    const tankaValues = TANKA_SHEET.getDataRange().getValues();
    if (tankaValues.length <= 1) return [];
    const tankaHeader = tankaValues.shift();
    const tankaColIndex = tankaHeader.reduce((acc, col, i) => {
      if (col) acc[col.trim()] = i;
      return acc;
    }, {});

    // 必須列の確認
    if (tankaColIndex['得意先コード'] === undefined || tankaColIndex['商品コード'] === undefined) {
      console.error('単価マスタに必要な列が見つかりません');
      return [];
    }

    // 顧客の商品をフィルタリング
    const cleanCustomerId = cleanSingleQuotes(customerId);
    const customerProducts = tankaValues.filter(row => {
      const cellId = cleanSingleQuotes(row[tankaColIndex['得意先コード']]);
      return cellId === cleanCustomerId;
    });

    if (customerProducts.length === 0) {
      console.log(`顧客 ${customerId} の商品データが見つかりません`);
      return [];
    }

    // 3. 商品マスタと価格改定リストを取得
    const productMaster = getProductMasterData();
    const revisionList = getRevisionListData();

    // 商品コードでの検索を高速化するマップを作成
    const productMap = {};
    productMaster.forEach(p => {
      productMap[String(p.code)] = p;
    });

    // 価格改定リストを商品コードでグループ化
    const revisionMap = {};
    revisionList.forEach(r => {
      const code = String(r.code);
      if (!revisionMap[code]) {
        revisionMap[code] = [];
      }
      revisionMap[code].push(r);
    });

    // 4. 各商品の価格情報を計算
    const prices = customerProducts.map(row => {
      const productCode = row[tankaColIndex['商品コード']] || '';
      const productName = (tankaColIndex['商品名称'] !== undefined)
        ? (row[tankaColIndex['商品名称']] || '')
        : '';
      const baraPrice = (tankaColIndex['バラ単価'] !== undefined)
        ? (row[tankaColIndex['バラ単価']] || 0)
        : 0;
      const effectiveDate = (tankaColIndex['変更有効日'] !== undefined)
        ? row[tankaColIndex['変更有効日']]
        : null;

      // 商品マスタから基本情報を取得
      const product = productMap[String(productCode)];
      let finalProductName = productName;
      let wholesalePrice = 0;

      if (product) {
        // 商品名が単価マスタに無ければ商品マスタから取得
        if (!finalProductName) {
          finalProductName = product.name || '';
        }

        // 基本卸価格を税区分に応じて取得
        wholesalePrice = useTaxInPrice ? (product.oroshi_tax_in || 0) : (product.oroshi || 0);

        // 価格改定リストで該当商品の改定があるかチェック
        const revisions = revisionMap[String(productCode)];
        if (revisions && revisions.length > 0) {
          // 変更有効日でフィルタリング（有効日が指定されている場合）
          let applicableRevisions = revisions;
          if (effectiveDate) {
            const effectiveDateStr = (effectiveDate instanceof Date)
              ? Utilities.formatDate(effectiveDate, Session.getScriptTimeZone(), "yyyy-MM-dd")
              : String(effectiveDate);

            applicableRevisions = revisions.filter(r => {
              return r.effectiveDate <= effectiveDateStr;
            });
          }

          // 最新の改定価格を使用（有効日が最も新しいもの）
          if (applicableRevisions.length > 0) {
            applicableRevisions.sort((a, b) => b.effectiveDate.localeCompare(a.effectiveDate));
            const latestRevision = applicableRevisions[0];
            wholesalePrice = useTaxInPrice
              ? (latestRevision.oroshi_tax_in || 0)
              : (latestRevision.oroshi || 0);
          }
        }
      }

      // 実際価格（バラ単価）
      const actualPrice = Number(baraPrice) || 0;
      const oroshi = Number(wholesalePrice) || 0;

      // 原価（商品マスタの仕入単価）を取得
      const genka = product
        ? (useTaxInPrice ? (product.shiire_tax_in || 0) : (product.shiire || 0))
        : 0;

      // 掛率と粗利率を計算
      let kakeritsu = 0;
      let arari = 0;

      if (oroshi > 0 && actualPrice > 0) {
        kakeritsu = (actualPrice / oroshi) * 100;
      }

      if (actualPrice > 0) {
        arari = ((actualPrice - genka) / actualPrice) * 100;
      }

      return {
        productCode: productCode,
        productName: finalProductName,
        oroshi: oroshi,
        jissai: actualPrice,
        kakeritsu: Math.round(kakeritsu * 100) / 100, // 小数点以下2桁
        arari: Math.round(arari * 100) / 100, // 小数点以下2桁
        biko1: tankaColIndex['備考１'] !== undefined ? String(row[tankaColIndex['備考１']] || '').trim() : '',
        biko2: tankaColIndex['備考２'] !== undefined ? String(row[tankaColIndex['備考２']] || '').trim() : ''
      };
    }).filter(p => p.productCode); // 商品コードがない行は除外

    console.log(`顧客 ${customerId} の商品データ取得: ${prices.length}件`);
    return prices;

  } catch (e) {
    console.error('getPricesByCustomerId エラー:', e);
    return [];
  }
}

function getOrCreateSheet(sheetName, headers) {
  let sheet = SPREADSHEET.getSheetByName(sheetName);
  if (!sheet) {
    sheet = SPREADSHEET.insertSheet(sheetName);
    sheet.appendRow(headers);
  } else if (sheet.getLastRow() === 0) {
    sheet.appendRow(headers);
  }
  return sheet;
}

function getCustomerMasterMap() {
    const values = CUSTOMER_SHEET.getDataRange().getValues();
    if (values.length <= 1) return {};
    const header = values.shift();
    const idCol = header.indexOf('得意先コード');
    const nameCol = header.indexOf('得意先名称');
    
    const map = {};
    values.forEach(row => {
        map[row[idCol]] = { name: row[nameCol] };
    });
    return map;
}

function addNewCustomerToMaster(newData) {
  const header = CUSTOMER_SHEET.getRange(1, 1, 1, CUSTOMER_SHEET.getLastColumn()).getValues()[0];
  const newRow = header.map(key => {
    // コード列の0落ち対策（シングルクォートを付ける）
    const codeColumns = ['得意先コード', '得意先グループコード', '請求先コード', '営業担当者コード'];
    if (codeColumns.includes(key) && newData[key]) {
      return "'" + newData[key];
    }
    return newData[key] || '';
  });
  CUSTOMER_SHEET.appendRow(newRow);
}

function updateCustomerInMaster(newData) {
  const customerIdToUpdate = newData['得意先コード'];
  const values = CUSTOMER_SHEET.getDataRange().getValues();
  const header = values.shift();
  const idCol = header.indexOf('得意先コード');

  const rowIndexToUpdate = values.findIndex(r => String(r[idCol]) == String(customerIdToUpdate));
  if (rowIndexToUpdate === -1) {
    throw new Error(`更新対象の顧客コードが見つかりません: ${customerIdToUpdate}`);
  }

  const updateDateTime = new Date();
  const newRow = header.map((key, index) => {
    // コード列の0落ち対策（シングルクォートを付ける）
    const codeColumns = ['得意先コード', '得意先グループコード', '請求先コード', '営業担当者コード'];
    if (codeColumns.includes(key) && newData[key] !== undefined) {
      return "'" + newData[key];
    }
    if (key === '最終更新日') {
      return updateDateTime;
    }
    return newData[key] !== undefined ? newData[key] : values[rowIndexToUpdate][index];
  });

  CUSTOMER_SHEET.getRange(rowIndexToUpdate + 2, 1, 1, header.length).setValues([newRow]);
}
/**
 * OLD登録処理
 * 顧客名に「Ｏ＿」を付与し、申請管理に決裁完了レコードを作成し、RPAシートに書き込みます。
 * @param {string} customerId - 得意先コード
 * @param {string} customerName - 得意先名称（表示用）
 * @return {Object} { success: boolean, message: string, appId?: number }
 */
function addOldRegistration(customerId, customerName) {
  try {
    const loginUser = getCurrentUser();
    if (!loginUser) throw new Error('ログイン情報が取得できません。');

    // 1. 得意先マスタの顧客名に「Ｏ＿」を付与
    const custValues = CUSTOMER_SHEET.getDataRange().getValues();
    const custHeader = custValues.shift();
    const custIdCol = custHeader.indexOf('得意先コード');
    const custNameCol = custHeader.indexOf('得意先名称');
    if (custIdCol === -1 || custNameCol === -1) throw new Error('得意先マスタに必要な列が見つかりません。');

    const cleanId = cleanSingleQuotes(String(customerId));
    const custRowIdx = custValues.findIndex(r => cleanSingleQuotes(String(r[custIdCol] || '')) === cleanId);
    if (custRowIdx === -1) throw new Error('得意先コード ' + customerId + ' がマスタに見つかりません。');

    const currentName = String(custValues[custRowIdx][custNameCol] || '');
    const newName = currentName.startsWith('Ｏ＿') ? currentName : 'Ｏ＿' + currentName;
    CUSTOMER_SHEET.getRange(custRowIdx + 2, custNameCol + 1).setValue(newName);
    console.log('[addOldRegistration] 顧客名を更新:', currentName, '->', newName);

    // ランクを「ＯＬＤ」に更新
    const custRankCol = custHeader.indexOf('得意先ランク区分名称');
    if (custRankCol !== -1) {
      CUSTOMER_SHEET.getRange(custRowIdx + 2, custRankCol + 1).setValue('ＯＬＤ');
      console.log('[addOldRegistration] ランクを「ＯＬＤ」に更新しました。');
    }

    // 2. 申請管理シートにレコードを作成（決裁完了）
    const appSheet = SPREADSHEET.getSheetByName('申請管理');
    if (!appSheet) throw new Error('申請管理シートが見つかりません。');

    const appValues = appSheet.getDataRange().getValues();
    const appHeader = appValues[0];
    const appHeaderMap = appHeader.reduce((m, h, i) => (m[h] = i, m), {});

    const existingIds = appValues.slice(1).map(r => Number(r[appHeaderMap['申請ID']])).filter(n => n > 0);
    const newAppId = existingIds.length > 0 ? Math.max(...existingIds) + 1 : 1;

    const now = new Date();
    const nowStr = Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy/MM/dd HH:mm:ss');

    const newRow = new Array(appHeader.length).fill('');
    const setCol = (colName, value) => {
      if (appHeaderMap[colName] !== undefined) newRow[appHeaderMap[colName]] = value;
    };

    setCol('申請ID', newAppId);
    setCol('申請日時', now);
    setCol('対象マスタ', '顧客マスタ');
    setCol('申請種別', 'OLD登録');
    setCol('対象顧客名', newName);
    setCol('得意先コード', "'" + cleanId);
    setCol('申請者名', loginUser.name || '');
    setCol('申請者メール', loginUser.email || '');
    setCol('申請者ID', loginUser.id || '');
    setCol('申請部署', loginUser.department || '');
    setCol('ステータス', '決裁完了');
    setCol('承認段階', '決裁完了');
    setCol('申請グループID', newAppId);
    setCol('再申請回数', 0);
    // 決裁者情報（社長承認列）
    setCol('決裁者判断者名', loginUser.name || '');
    setCol('決裁者判断時刻', nowStr);
    setCol('社長承認', '承認済み');

    appSheet.appendRow(newRow);
    console.log('[addOldRegistration] 申請管理にOLD登録レコードを追加しました。申請ID:', newAppId);

    // 3. RPAシートに書き込み（新規/修正 = 'OLD登録'）
    try {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      let rpaSheet = ss.getSheetByName(RPA_CUSTOMER_SHEET_NAME);
      if (!rpaSheet) rpaSheet = createCustomerRPASheet(ss);

      const updatedCustValues = CUSTOMER_SHEET.getDataRange().getValues();
      const updatedHeader = updatedCustValues[0];
      const updatedRow = updatedCustValues.find(r =>
        cleanSingleQuotes(String(r[updatedHeader.indexOf('得意先コード')] || '')) === cleanId
      );

      if (updatedRow) {
        const rpaData = mapCustomerDataToRPA(updatedRow, updatedHeader, false, now, ss);
        rpaData['新規/修正'] = 'OLD登録'; // 上書き
        const rpaHeader = rpaSheet.getRange(1, 1, 1, rpaSheet.getLastColumn()).getValues()[0];
        const rpaRow = rpaHeader.map(col => rpaData[col] !== undefined ? rpaData[col] : '');
        rpaSheet.appendRow(rpaRow);
        console.log('[addOldRegistration] RPAシートに書き込みました。');
      }
    } catch (rpaErr) {
      console.warn('[addOldRegistration] RPA連携に失敗しましたが、処理は続行します:', rpaErr);
    }

    // キャッシュクリア（顧客リストとダッシュボードを即時更新させる）
    clearCache(CACHE_KEY_CUSTOMERS);
    clearCache(CACHE_KEY_DASHBOARD);

    return { success: true, message: 'OLD登録が完了しました。', appId: newAppId };

  } catch (e) {
    console.error('[addOldRegistration] エラー:', e);
    return { success: false, message: 'OLD登録に失敗しました: ' + e.message };
  }
}

/**
 * 商品マスタ、裁断マスタ、袋詰マスタ、★価格改定リスト★のデータを取得します。
 */
function getMasterData() {
    try {
        // ★修正: データ量が大きいためキャッシュを無効化
        // 商品マスタデータを取得
        const products = getProductMasterData();
        const revisions = getRevisionListData();

        // GAS内に定義されたマスタと合わせて返す
        const masterData = {
            products: products,
            revisions: revisions,
            saitan: SAITAN_MASTER, // グローバル定数
            fukuro: FUKURO_MASTER  // グローバル定数
        };

        return masterData;

    } catch (e) {
        console.error('Failed to get master data:', e);
        // エラーメッセージに詳細を含める
        throw new Error('サーバーエラー: マスタデータの取得に失敗しました。(' + e.message + ')');
    }
}
/**
 * 価格改定リストシートからデータを読み込み、オブジェクトの配列として返します。
 */
function getRevisionListData() {
    // キャッシュをチェック
    const cache = CacheService.getScriptCache();
    const cacheKey = 'revision_list_data';
    const cached = cache.get(cacheKey);

    if (cached) {
      console.log('Cache hit: revision_list_data');
      return JSON.parse(cached);
    }

    console.log('Cache miss: revision_list_data - シートから読み込みます');

    // 価格改定リストシートの存在確認
    if (!REVISION_SHEET) {
        // シートが必須で存在しない場合はエラーを投げる
        throw new Error('「価格改定リスト」シートが見つかりません。');
    }
    // シートから全データを取得 (2行目以降)
    const values = REVISION_SHEET.getDataRange().getValues();
    // ヘッダー行を取得・削除
    if (values.length <= 1) return []; // ヘッダーのみ or 空の場合は空配列を返す
    const header = values.shift();

    // ヘッダー名から列インデックス(0始まり)へのマップを作成
    const colIndex = header.reduce((acc, col, i) => {
        if (col) acc[col.trim()] = i; 
        return acc;
       }, {});

    // 必要な列名リスト
    const requiredCols = [
        '変更有効日', '商品コード', '名称', 
        '店頭販売単価(卸単価)', '税込店頭販売単価(卸単価)', 
        '店頭仕入単価', '税込店頭仕入単価'
    ];
    // 必要な列がすべて存在するか確認
    for (const col of requiredCols) {
        if (colIndex[col] === undefined) {
            console.error('Missing column in 価格改定リスト:', col, 'Available columns:', header);
            throw new Error(`「価格改定リスト」シートに必要な列 "${col}" が見つかりません。`);
        }
    }

    // データ行をオブジェクトの配列に変換
    const revisions = values.map(row => {
        const effectiveDateRaw = row[colIndex['変更有効日']];
        let effectiveDateStr = '';
        if (effectiveDateRaw instanceof Date && !isNaN(effectiveDateRaw)) {
             // 日付の場合は yyyy-MM-dd 形式で返す
            try {
                effectiveDateStr = Utilities.formatDate(effectiveDateRaw, Session.getScriptTimeZone(), "yyyy-MM-dd");
            } catch(e) {
                console.warn('価格改定リストの日付フォーマット変換に失敗しました:', effectiveDateRaw, e.message);
            }
        }

        return {
            effectiveDate: effectiveDateStr, // yyyy-MM-dd 形式
            code: row[colIndex['商品コード']] !== undefined ? row[colIndex['商品コード']] : '',
            name: row[colIndex['名称']] !== undefined ? row[colIndex['名称']] : '',
            oroshi: row[colIndex['店頭販売単価(卸単価)']] !== undefined ? row[colIndex['店頭販売単価(卸単価)']] : 0,
            oroshi_tax_in: row[colIndex['税込店頭販売単価(卸単価)']] !== undefined ? row[colIndex['税込店頭販売単価(卸単価)']] : 0,
            shiire: row[colIndex['店頭仕入単価']] !== undefined ? row[colIndex['店頭仕入単価']] : 0,
            shiire_tax_in: row[colIndex['税込店頭仕入単価']] !== undefined ? row[colIndex['税込店頭仕入単価']] : 0
        };
    })
    .filter(r => r.code && r.effectiveDate); // 商品コードと有効日が存在する行のみ

    // キャッシュに保存（10分間）
    try {
      cache.put(cacheKey, JSON.stringify(revisions), 600);
      console.log('Cache saved: revision_list_data');
    } catch (e) {
      console.warn('キャッシュ保存失敗（データが大きすぎる可能性）:', e);
    }

    return revisions;
}
/**
 * 商品マスタシートからデータを読み込み、オブジェクトの配列として返します。
 * [修正] 店頭仕入単価（shiire, shiire_tax_in）も読み込むように変更。
 */
function getProductMasterData() {
    // キャッシュをチェック
    const cache = CacheService.getScriptCache();
    const cacheKey = 'product_master_data';
    const cached = cache.get(cacheKey);

    if (cached) {
      console.log('Cache hit: product_master_data');
      return JSON.parse(cached);
    }

    console.log('Cache miss: product_master_data - シートから読み込みます');

    // 商品マスタシートの存在確認
    if (!PRODUCT_SHEET) {
        throw new Error('「商品マスタ」シートが見つかりません。');
    }
    // シートから全データを取得 (2行目以降)
    const values = PRODUCT_SHEET.getDataRange().getValues();
    // ヘッダー行を取得・削除
    if (values.length <= 1) return []; // ヘッダーのみ or 空の場合は空配列を返す
    const header = values.shift();

    // ヘッダー名から列インデックス(0始まり)へのマップを作成 (ヘッダー名の前後の空白は除去)
    const colIndex = header.reduce((acc, col, i) => {
        if (col) acc[col.trim()] = i; // 空でないヘッダーのみ登録
        return acc;
       }, {});

    // 必要な列名リスト
    const requiredCols = [
        '商品コード', '名称', '店頭販売単価(卸単価)', '税込店頭販売単価(卸単価)', 
        '小売単価', '税込小売単価', 
        '店頭仕入単価', '税込店頭仕入単価'
    ];
    // 必要な列がすべて存在するか確認
    for (const col of requiredCols) {
        if (colIndex[col] === undefined) {
            console.error('Missing column in 商品マスタ:', col, 'Available columns:', header);
            throw new Error(`「商品マスタ」シートに必要な列 "${col}" が見つかりません。`);
        }
    }

    // データ行をオブジェクトの配列に変換
    const products = values.map(row => ({
        // 各列の値を取得 (空文字チェック追加)
        code: row[colIndex['商品コード']] !== undefined ? row[colIndex['商品コード']] : '',
        name: row[colIndex['名称']] !== undefined ? row[colIndex['名称']] : '',
        oroshi: row[colIndex['店頭販売単価(卸単価)']] !== undefined ? row[colIndex['店頭販売単価(卸単価)']] : 0, // 数値想定、なければ0
        oroshi_tax_in: row[colIndex['税込店頭販売単価(卸単価)']] !== undefined ? row[colIndex['税込店頭販売単価(卸単価)']] : 0, // 数値想定、なければ0
        kouri: row[colIndex['小売単価']] !== undefined ? row[colIndex['小売単価']] : 0, // 数値想定、なければ0
        kouri_tax_in: row[colIndex['税込小売単価']] !== undefined ? row[colIndex['税込小売単価']] : 0, // 数値想定、なければ0
        shiire: row[colIndex['店頭仕入単価']] !== undefined ? row[colIndex['店頭仕入単価']] : 0, // 数値想定、なければ0
        shiire_tax_in: row[colIndex['税込店頭仕入単価']] !== undefined ? row[colIndex['税込店頭仕入単価']] : 0 // 数値想定、なければ0
    }))
    .filter(p => p.code); // 商品コードが存在する行のみをフィルタリング

    // キャッシュに保存（10分間）
    try {
      cache.put(cacheKey, JSON.stringify(products), 600);
      console.log('Cache saved: product_master_data');
    } catch (e) {
      console.warn('キャッシュ保存失敗（データが大きすぎる可能性）:', e);
    }

    return products;
}

/**
 * 得意先マスタインポート前に、進行中の顧客関連申請を確認します。
 * @returns {Object} { hasActive: boolean, applications: Array }
 */
function checkActiveCustomerApplications() {
  try {
    const appSheet = SPREADSHEET.getSheetByName('申請管理');
    if (!appSheet) return { hasActive: false, applications: [] };

    const values = appSheet.getDataRange().getValues();
    if (values.length <= 1) return { hasActive: false, applications: [] };

    const header = values[0];
    const idCol = header.indexOf('申請ID');
    const typeCol = header.indexOf('申請種別');
    const statusCol = header.indexOf('ステータス');
    const stageCol = header.indexOf('承認段階');
    const customerNameCol = header.indexOf('対象顧客名');
    const customerIdCol = header.indexOf('得意先コード');

    if (idCol === -1 || typeCol === -1 || statusCol === -1) {
      return { hasActive: false, applications: [] };
    }

    const customerTypes = new Set(['顧客新規登録', '顧客情報修正', 'OLD登録']);
    const terminalStatuses = new Set(['決裁完了', '却下', '承認済']);

    const active = values.slice(1).filter(row => {
      const appType = String(row[typeCol] || '');
      const status = String(row[statusCol] || '');
      return customerTypes.has(appType) && !terminalStatuses.has(status);
    }).map(row => ({
      id: row[idCol],
      type: row[typeCol],
      status: row[statusCol],
      stage: stageCol !== -1 ? row[stageCol] : '',
      customerName: customerNameCol !== -1 ? row[customerNameCol] : '',
      customerId: customerIdCol !== -1 ? cleanSingleQuotes(String(row[customerIdCol] || '')) : ''
    }));

    return { hasActive: active.length > 0, applications: active };
  } catch (e) {
    console.error('[checkActiveCustomerApplications] エラー:', e);
    return { hasActive: false, applications: [] };
  }
}

/**
 * クライアントから送信されたファイルデータをパースし、スプレッドシートに書き込みます。
 * @param {string} dataType - データ種別 ('jutaikin', 'sales', 'revision')
 * @param {string} base64Content - Base64エンコードされたファイル内容
 * @param {string} fileName - ファイル名
 * @param {boolean} confirmOverwrite - true の場合、同じ変更有効日の既存データを上書き（revision のみ）
 * @returns {object} - {status: 'success'/'error'/'conflict', message: string, ...}
 */
function importData(dataType, base64Content, fileName, confirmOverwrite, jutaikinBaseYearMonth) {
  try {
    // データ種別の定義
    const dataTypeConfig = {
      'jutaikin': {
        sheetName: '渋滞金リスト',
        requiredHeaders: ['得意先コード', '得意先名称', '1回前', '2回前', '3回前']
      },
      'sales': {
        sheetName: '売上データ',
        requiredHeaders: ['得意先コード', '得意先名称', '当期累計売上', '累計前年比', '前期累計売上', '単月売上', '単月前年比', '前年単月売上']
      },
      'revision': {
        sheetName: '価格改定リスト',
        requiredHeaders: ['変更有効日', '商品コード', '名称', '店頭仕入単価', '税込店頭仕入単価', '店頭販売単価(卸単価)', '税込店頭販売単価(卸単価)']
      },
      'product': {
        sheetName: '商品マスタ',
        requiredHeaders: ['商品コード', '名称', '店頭販売単価(卸単価)', '税込店頭販売単価(卸単価)', '小売単価', '税込小売単価', '店頭仕入単価', '税込店頭仕入単価']
      },
      'customer': {
        sheetName: '得意先マスタ',
        requiredHeaders: ['得意先コード', '得意先名称']
      }
    };

    if (!dataTypeConfig[dataType]) {
      throw new Error('不正なデータ種別が指定されました。');
    }

    const config = dataTypeConfig[dataType];

    // Base64デコード
    const blob = Utilities.newBlob(Utilities.base64Decode(base64Content), 'application/octet-stream', fileName);

    // ファイル形式の判定
    let parsedData;
    if (fileName.toLowerCase().endsWith('.csv')) {
      // CSV形式の場合
      parsedData = parseCsvFile(blob);
    } else if (fileName.toLowerCase().endsWith('.xlsx') || fileName.toLowerCase().endsWith('.xls')) {
      // Excel形式の場合
      parsedData = parseExcelFile(blob, fileName);
    } else {
      throw new Error('サポートされていないファイル形式です。CSVまたはExcelファイルを選択してください。');
    }

    if (!parsedData || parsedData.length === 0) {
      throw new Error('ファイルからデータを読み込めませんでした。');
    }

    // ヘッダー検証（空白・BOM除去）
    const fileHeaders = parsedData[0].map(h => String(h || '').trim().replace(/^\uFEFF/, ''));
    const missingHeaders = config.requiredHeaders.filter(h => !fileHeaders.includes(h));

    if (missingHeaders.length > 0) {
      throw new Error(`必要なヘッダーが不足しています: ${missingHeaders.join(', ')}`);
    }

    // ヘッダー行をクリーン済みの値で上書き（BOM・前後空白がシートに書き込まれるのを防止）
    parsedData[0] = fileHeaders;

    // 対象シートを取得または作成
    let targetSheet = SPREADSHEET.getSheetByName(config.sheetName);
    if (!targetSheet) {
      targetSheet = SPREADSHEET.insertSheet(config.sheetName);
    }

    // コード列（得意先コード/商品コード/得意先グループコード）を7桁ゼロ埋め＋'プレフィックスに正規化
    if (parsedData.length > 1) {
      const codeNormalizeCols = ['得意先コード', '商品コード', '得意先グループコード', '請求先コード'];
      const codeColIndices = [];
      fileHeaders.forEach((h, i) => {
        if (codeNormalizeCols.includes(h)) codeColIndices.push(i);
      });
      if (codeColIndices.length > 0) {
        for (let rowIdx = 1; rowIdx < parsedData.length; rowIdx++) {
          const row = parsedData[rowIdx];
          codeColIndices.forEach(colIdx => {
            if (row[colIdx] !== undefined && row[colIdx] !== null && row[colIdx] !== '') {
              let code = String(row[colIdx]).trim();
              code = code.replace(/^['"]|['"]$/g, '');
              if (/^\d+$/.test(code)) {
                code = code.padStart(7, '0');
              }
              row[colIdx] = "'" + code;
            }
          });
        }
      }
    }

    if (dataType === 'revision') {
      // ── ヘルパー: 日付値を 'yyyy-MM-dd' 文字列に正規化（比較用）──
      function dateToYMD(val) {
        if (!val) return null;
        try {
          const d = new Date(val);
          if (isNaN(d.getTime()) || d.getFullYear() < 1990) return null;
          return Utilities.formatDate(d, 'JST', 'yyyy-MM-dd');
        } catch(e) { return null; }
      }

      const today = new Date();
      today.setHours(0, 0, 0, 0);

      // アップロードファイル内の変更有効日セットを収集
      const newHeader = parsedData[0];
      const newDateColIdx = newHeader.findIndex(h => String(h).trim() === '変更有効日');
      const newCodeColIdx = newHeader.findIndex(h => String(h).trim() === '商品コード');
      const newNameColIdx = newHeader.findIndex(h => String(h).trim() === '名称');

      const uploadedDateSet = new Set();
      if (newDateColIdx !== -1) {
        for (let i = 1; i < parsedData.length; i++) {
          const ds = dateToYMD(parsedData[i][newDateColIdx]);
          if (ds) uploadedDateSet.add(ds);
        }
      }

      // 既存データを読み込んで重複チェック
      const existingData = targetSheet.getDataRange().getValues();
      const conflictRows = [];

      if (existingData.length > 1 && uploadedDateSet.size > 0) {
        const exHeader = existingData[0];
        const exDateCol = exHeader.findIndex(h => String(h).trim() === '変更有効日');
        const exCodeCol = exHeader.findIndex(h => String(h).trim() === '商品コード');
        const exNameCol = exHeader.findIndex(h => String(h).trim() === '名称');

        if (exDateCol !== -1) {
          for (let i = 1; i < existingData.length; i++) {
            const ds = dateToYMD(existingData[i][exDateCol]);
            if (ds && uploadedDateSet.has(ds)) {
              conflictRows.push({
                date: ds,
                code: exCodeCol !== -1 ? String(existingData[i][exCodeCol] || '').replace(/^'/, '') : '',
                name: exNameCol !== -1 ? String(existingData[i][exNameCol] || '') : ''
              });
            }
          }
        }
      }

      // 重複あり かつ 未確認 → 確認を求める
      if (conflictRows.length > 0 && !confirmOverwrite) {
        const conflictDates = [...new Set(conflictRows.map(r => r.date))].sort();
        return {
          status: 'conflict',
          conflictDates: conflictDates,
          conflictRows: conflictRows
        };
      }

      // ── 書き込み処理 ──
      if (existingData.length > 1) {
        const exHeader = existingData[0];
        const exDateCol = exHeader.findIndex(h => String(h).trim() === '変更有効日');

        // 保持条件: 過去日付でない AND（上書き確認済みの場合）アップロード日付と一致しない
        let filteredRows = [exHeader];
        for (let i = 1; i < existingData.length; i++) {
          if (exDateCol !== -1) {
            const dateVal = existingData[i][exDateCol];
            if (dateVal) {
              const ds = dateToYMD(dateVal);
              const rowDate = new Date(dateVal);
              rowDate.setHours(0, 0, 0, 0);
              const isPast = rowDate < today;
              const isOverwriteTarget = confirmOverwrite && ds && uploadedDateSet.has(ds);
              if (!isPast && !isOverwriteTarget) filteredRows.push(existingData[i]);
            } else {
              filteredRows.push(existingData[i]); // 日付なし行は保持
            }
          } else {
            filteredRows.push(existingData[i]);
          }
        }

        // シートを再構成
        targetSheet.clearContents();
        if (filteredRows.length > 0) {
          targetSheet.getRange(1, 1, filteredRows.length, filteredRows[0].length).setValues(filteredRows);
        }

        // 新しいデータ行を追加（ヘッダー行はスキップ）
        const newDataRows = parsedData.slice(1);
        if (newDataRows.length > 0) {
          const lastRow = targetSheet.getLastRow();
          targetSheet.getRange(lastRow + 1, 1, newDataRows.length, newDataRows[0].length).setValues(newDataRows);
        }
      } else {
        // 既存データなし → 全書き込み
        targetSheet.clear();
        if (parsedData.length > 0) {
          targetSheet.getRange(1, 1, parsedData.length, parsedData[0].length).setValues(parsedData);
        }
      }
    } else {
      // その他のデータ種別は全書き換え
      targetSheet.clear();
      if (parsedData.length > 0) {
        targetSheet.getRange(1, 1, parsedData.length, parsedData[0].length).setValues(parsedData);
      }
    }

    // マスタ更新履歴に記録
    try {
      const updater = Session.getActiveUser().getEmail() || 'システム';
      const dataCount = parsedData.length - 1; // ヘッダー行を除く
      const changeDescription = `${dataCount}件のデータをインポート（ファイル: ${fileName}）`;
      addMasterUpdateHistory(updater, config.sheetName, changeDescription);
    } catch (historyError) {
      console.error('[importData] マスタ更新履歴の記録に失敗しました:', historyError);
      // 履歴記録の失敗はインポート処理自体には影響させない
    }

    // 渋滞金インポート時：顧客キャッシュをクリア（hasJutaikinフラグ更新のため）
    if (dataType === 'jutaikin') {
      try {
        CacheService.getScriptCache().remove('customers_cache');
        console.log('[importData] 渋滞金インポートのため顧客キャッシュをクリアしました');
      } catch (cacheError) {
        console.error('[importData] 顧客キャッシュクリアに失敗:', cacheError);
      }
    }

    // 渋滞金インポート時：基準年月をスクリプトプロパティに保存
    if (dataType === 'jutaikin' && jutaikinBaseYearMonth) {
      try {
        PropertiesService.getScriptProperties().setProperty('jutaikin_base_yearmonth', jutaikinBaseYearMonth);
        console.log('[importData] 渋滞金基準年月を保存:', jutaikinBaseYearMonth);
      } catch (propError) {
        console.error('[importData] 基準年月の保存に失敗:', propError);
      }
    }

    // 得意先マスタインポート時：関連キャッシュをクリア
    if (dataType === 'customer') {
      try {
        const cache = CacheService.getScriptCache();
        cache.remove('customers_cache');
        cache.remove('master_data_cache');
        console.log('[importData] 得意先マスタキャッシュをクリアしました');
      } catch (cacheError) {
        console.error('[importData] 得意先マスタキャッシュクリアに失敗:', cacheError);
      }
    }

    // 商品マスタインポート時：キャッシュをクリア
    if (dataType === 'product') {
      try {
        CacheService.getScriptCache().remove('product_master_data');
        console.log('[importData] 商品マスタキャッシュをクリアしました');
      } catch (cacheError) {
        console.error('[importData] 商品マスタキャッシュクリアに失敗:', cacheError);
      }
    }

    // 価格改定リストインポート時：キャッシュをクリア
    if (dataType === 'revision') {
      try {
        CacheService.getScriptCache().remove('revision_list_data');
        console.log('[importData] 価格改定リストキャッシュをクリアしました');
      } catch (cacheError) {
        console.error('[importData] 価格改定リストキャッシュクリアに失敗:', cacheError);
      }
    }

    return {
      status: 'success',
      message: `${parsedData.length - 1}件のデータを「${config.sheetName}」にインポートしました。`
    };

  } catch (error) {
    console.error('Import error:', error);
    return {
      status: 'error',
      message: error.message || 'インポート処理中にエラーが発生しました。'
    };
  }
}

/**
 * CSVファイルをパースします（UTF-8/Shift-JIS自動判定・改善版）
 */
function parseCsvFile(blob) {
  // 試す文字コードのリスト（優先順）
  const encodings = ['UTF-8', 'Shift_JIS', 'EUC-JP', 'ISO-2022-JP'];

  let bestResult = null;
  let bestScore = -1;

  for (const encoding of encodings) {
    try {
      const csvText = blob.getDataAsString(encoding);
      const parsed = Utilities.parseCsv(csvText);

      if (!parsed || parsed.length === 0) continue;

      // ヘッダー行を取得
      const header = parsed[0].join('');

      // 文字化けスコアを計算（低いほど良い）
      let score = 0;

      // 制御文字や不正な文字をカウント
      score += (header.match(/[\x00-\x08\x0B-\x0C\x0E-\x1F\x7F]/g) || []).length * 10;

      // 文字化け記号をカウント
      score += (header.match(/[�]/g) || []).length * 5;
      score += (header.match(/[□]/g) || []).length * 5;
      score += (header.match(/[\uFFFD]/g) || []).length * 5;

      // 意味不明な連続文字をカウント（例：繧・蜷・）
      score += (header.match(/[ｦ-ｿ]{3,}/g) || []).length * 3;

      // 日本語文字が正しく含まれているかチェック（含まれていればスコア減算）
      if (header.match(/[ぁ-んァ-ヶー一-龠]/)) {
        score -= 10;
      }

      // より良い結果を記録
      if (bestResult === null || score < bestScore) {
        bestResult = parsed;
        bestScore = score;
      }

      // スコアが十分良ければ（文字化けなし）、すぐに返す
      if (score < 0) {
        return parsed;
      }

    } catch (error) {
      // このエンコーディングでは読めなかった
      continue;
    }
  }

  if (bestResult) {
    return bestResult;
  }

  throw new Error('CSVファイルの読み込みに失敗しました。UTF-8またはShift-JISで保存されたCSVファイルを使用してください。');
}

/**
 * Excelファイルをパースします（一時的にDriveにアップロード）
 */
function parseExcelFile(blob, fileName) {
  let tempFile = null;
  try {
    // 一時ファイルとしてDriveに保存
    tempFile = DriveApp.createFile(blob);
    tempFile.setName('temp_import_' + new Date().getTime() + '_' + fileName);

    // SpreadsheetとしてExcelファイルを開く
    const tempSpreadsheet = SpreadsheetApp.open(tempFile);
    const firstSheet = tempSpreadsheet.getSheets()[0];

    // データを取得
    const data = firstSheet.getDataRange().getValues();

    // 一時ファイルを削除
    DriveApp.getFileById(tempFile.getId()).setTrashed(true);

    return data;
  } catch (error) {
    // エラーが発生した場合も一時ファイルを削除
    if (tempFile) {
      try {
        DriveApp.getFileById(tempFile.getId()).setTrashed(true);
      } catch (e) {
        console.warn('一時ファイルの削除に失敗しました。手動で削除が必要な場合があります。ファイルID:', tempFile.getId(), e.message);
      }
    }
    throw new Error('Excelファイルの読み込みに失敗しました: ' + error.message);
  }
}

/**
 * 指定された得意先コードの渋滞金情報を取得します。
 */
function getJutaikinByCustomerId(customerId) {
  try {
    const jutaikinSheet = SPREADSHEET.getSheetByName('渋滞金リスト');
    if (!jutaikinSheet) {
      console.warn('「渋滞金リスト」シートが見つかりません。');
      return null;
    }

    const values = jutaikinSheet.getDataRange().getValues();
    if (values.length <= 1) return null; // ヘッダーのみ or 空

    const header = values.shift();
    const colIndex = header.reduce((acc, col, i) => {
      if (col) acc[col.trim()] = i;
      return acc;
    }, {});

    // 必要な列の存在確認
    if (colIndex['得意先コード'] === undefined) {
      console.error('「渋滞金リスト」シートに「得意先コード」列が見つかりません。');
      return null;
    }

    // 該当する顧客の渋滞金データを検索（シングルクォートを除去して比較）
    const cleanCustomerId = cleanSingleQuotes(customerId);
    const row = values.find(r => {
      const cellId = cleanSingleQuotes(r[colIndex['得意先コード']]);
      return cellId === cleanCustomerId;
    });
    if (!row) return null;

    // '4回以前'（半角）または'４回以前'（全角）どちらの列名にも対応
    const col4idx = colIndex['4回以前'] !== undefined ? colIndex['4回以前'] : colIndex['４回以前'];
    return {
      '1回前': row[colIndex['1回前']] || 0,
      '2回前': row[colIndex['2回前']] || 0,
      '3回前': row[colIndex['3回前']] || 0,
      '4回以前': (col4idx !== undefined ? row[col4idx] : 0) || 0
    };
  } catch (error) {
    console.error('渋滞金データの取得に失敗しました:', error);
    return null;
  }
}

/**
 * マスタ更新履歴から指定されたマスタの最新更新日時を取得します。
 * @param {string} targetMaster - 対象マスタ名（例：「売上データ」）
 * @returns {Date|null} 最新の更新日時、見つからない場合はnull
 */
function getLatestMasterUpdateDate(targetMaster) {
  try {
    const historySheet = SPREADSHEET.getSheetByName('マスタ更新履歴');
    if (!historySheet) {
      console.warn('マスタ更新履歴シートが見つかりません。');
      return null;
    }

    const values = historySheet.getDataRange().getValues();
    if (values.length <= 1) {
      console.warn('マスタ更新履歴シートにデータがありません。');
      return null;
    }

    // ヘッダーをスキップして、対象マスタに一致する行を抽出
    const targetRows = values.slice(1).filter(row => {
      const masterName = row[2]; // 対象マスタは3列目（インデックス2）
      return masterName === targetMaster;
    });

    if (targetRows.length === 0) {
      console.warn(`対象マスタ「${targetMaster}」の更新履歴が見つかりません。`);
      return null;
    }

    // 更新日時（1列目、インデックス0）で降順ソートして最新を取得
    targetRows.sort((a, b) => {
      const dateA = new Date(a[0]);
      const dateB = new Date(b[0]);
      return dateB - dateA; // 降順
    });

    const latestDate = targetRows[0][0];
    return latestDate instanceof Date ? latestDate : new Date(latestDate);
  } catch (error) {
    console.error('マスタ更新履歴の取得に失敗しました:', error);
    return null;
  }
}

/**
 * 指定された得意先コードの売上情報を取得します。
 */
function getSalesByCustomerId(customerId) {
  try {
    const salesSheet = SPREADSHEET.getSheetByName('売上データ');
    if (!salesSheet) {
      console.warn('「売上データ」シートが見つかりません。');
      return null;
    }

    const values = salesSheet.getDataRange().getValues();
    if (values.length <= 1) {
      console.warn('「売上データ」シートにデータがありません（ヘッダーのみまたは空）。');
      return null;
    }

    const header = values.shift();
    const colIndex = header.reduce((acc, col, i) => {
      if (col) acc[col.trim()] = i;
      return acc;
    }, {});

    // 必要な列の存在確認
    if (colIndex['得意先コード'] === undefined) {
      console.error('「売上データ」シートに「得意先コード」列が見つかりません。');
      console.error('利用可能な列:', Object.keys(colIndex));
      return null;
    }

    // 該当する顧客の売上データを検索（シングルクォートを除去して比較）
    const cleanCustomerId = cleanSingleQuotes(customerId);
    console.log(`売上データ検索: 得意先コード=${cleanCustomerId}`);

    const row = values.find(r => {
      const cellId = cleanSingleQuotes(r[colIndex['得意先コード']]);
      return cellId === cleanCustomerId;
    });

    if (!row) {
      console.warn(`売上データが見つかりません: 得意先コード=${cleanCustomerId}`);
      return null;
    }

    // マスタ更新履歴から売上データの最新更新日時を取得
    const updateDate = getLatestMasterUpdateDate('売上データ');

    return {
      '当期累計売上': row[colIndex['当期累計売上']] || 0,
      '累計前期比': row[colIndex['累計前年比']] || '',
      '前期累計売上': row[colIndex['前期累計売上']] || 0,
      '当期単月売上': row[colIndex['単月売上']] || 0,
      '単月前期比': row[colIndex['単月前年比']] || '',
      '前期単月売上': row[colIndex['前年単月売上']] || 0,
      '更新日': updateDate // ★追加：マスタ更新履歴から取得した更新日時
    };
  } catch (error) {
    console.error('売上データの取得に失敗しました:', error);
    return null;
  }
}

// =============================================
// 再申請機能
// =============================================

/**
 * 却下された申請データを再申請フォーム用に取得
 * @param {string} applicationId - 元の申請ID
 * @return {Object} 申請データと修正指示項目
 */
function getApplicationForResubmit(applicationId) {
  try {
    const appSheet = SPREADSHEET.getSheetByName('申請管理');
    if (!appSheet) throw new Error('申請管理シートが見つかりません。');

    // 申請管理から基本情報を取得
    const appValues = appSheet.getDataRange().getValues();
    const appHeader = appValues.shift();

    const appIdCol = appHeader.indexOf('申請ID');
    const targetMasterCol = appHeader.indexOf('対象マスタ');
    const appTypeCol = appHeader.indexOf('申請種別');
    const customerIdCol = appHeader.indexOf('得意先コード');
    const rejectReasonCol = appHeader.indexOf('却下理由');
    const requiredFieldsCol = appHeader.indexOf('修正指示項目');
    const groupIdCol = appHeader.indexOf('申請グループID');

    const appRow = appValues.find(r => String(r[appIdCol]) === String(applicationId));
    if (!appRow) throw new Error('指定された申請が見つかりません。');

    const targetMaster = targetMasterCol !== -1 ? appRow[targetMasterCol] : '';
    const applicationType = appTypeCol !== -1 ? appRow[appTypeCol] : '';
    const customerId = customerIdCol !== -1 ? cleanSingleQuotes(appRow[customerIdCol]) : '';
    const rejectReason = rejectReasonCol !== -1 ? appRow[rejectReasonCol] : '';
    const groupId = groupIdCol !== -1 ? appRow[groupIdCol] : applicationId;

    // 修正指示項目をJSON配列としてパース
    let requiredFields = [];
    if (requiredFieldsCol !== -1 && appRow[requiredFieldsCol]) {
      try {
        requiredFields = JSON.parse(appRow[requiredFieldsCol]);
      } catch (e) {
        console.warn('修正指示項目のJSON解析失敗:', e);
      }
    }

    // 申請データを取得
    let applicationData = {};
    let originalData = {};

    if (targetMaster === '顧客マスタ') {
      // 顧客マスタの現在のデータを取得（元データ）
      originalData = getCustomerBasicInfo(customerId);

      // 申請データ_顧客から申請時のデータを取得
      const detailSheet = SPREADSHEET.getSheetByName('申請データ_顧客');
      if (detailSheet) {
        const detailValues = detailSheet.getDataRange().getValues();
        const detailHeader = detailValues.shift();

        const detailAppIdCol = detailHeader.indexOf('申請ID');
        const itemNameCol = detailHeader.indexOf('項目名');
        const newValueCol = detailHeader.indexOf('修正後の値');

        const relatedDetails = detailValues.filter(r =>
          String(r[detailAppIdCol]) === String(applicationId)
        );

        relatedDetails.forEach(row => {
          const itemName = row[itemNameCol];
          const newValue = row[newValueCol];
          applicationData[itemName] = newValue;
        });
      }
    } else if (targetMaster === '単価マスタ') {
      // 単価マスタの場合
      // 現在の単価マスタデータを取得（元データ）
      const prices = getPricesByCustomerId(customerId);
      originalData = prices; // 配列で返す

      // 申請データ_単価から申請時のデータを取得
      const detailSheet = SPREADSHEET.getSheetByName('申請データ_単価');
      if (detailSheet) {
        const detailValues = detailSheet.getDataRange().getValues();
        const detailHeader = detailValues.shift();

        const detailAppIdCol = detailHeader.indexOf('申請ID');
        const kubunCol = detailHeader.indexOf('登録区分');
        const productCodeCol = detailHeader.indexOf('商品コード');

        // 修正前・修正後の列インデックス
        const colMap = {};
        ['商品名', '裁断方法コード', '裁断方法名', '袋詰方法コード', '袋詰方法名',
         '卸価格', '実際販売価格', '掛率', '粗利率'].forEach(fieldName => {
          colMap[fieldName + '_修正前'] = detailHeader.indexOf(fieldName + '_修正前');
          colMap[fieldName + '_修正後'] = detailHeader.indexOf(fieldName + '_修正後');
        });

        const relatedDetails = detailValues.filter(r =>
          String(r[detailAppIdCol]) === String(applicationId)
        );

        // 申請データを配列形式で構築（商品ごとに1オブジェクト）
        applicationData = relatedDetails.map(row => {
          const kubun = row[kubunCol];
          const productCode = row[productCodeCol];

          const item = {
            登録区分: kubun,
            商品コード: productCode
          };

          // 修正前・修正後の値を取得
          ['商品名', '裁断方法コード', '裁断方法名', '袋詰方法コード', '袋詰方法名',
           '卸価格', '実際販売価格', '掛率', '粗利率'].forEach(fieldName => {
            if (colMap[fieldName + '_修正前'] !== -1) {
              item[fieldName + '_修正前'] = row[colMap[fieldName + '_修正前']];
            }
            if (colMap[fieldName + '_修正後'] !== -1) {
              item[fieldName + '_修正後'] = row[colMap[fieldName + '_修正後']];
            }
          });

          return item;
        });
      }
    }

    return {
      applicationData: applicationData,
      originalData: originalData,
      requiredFields: requiredFields,
      rejectReason: rejectReason,
      申請グループID: groupId,
      targetMaster: targetMaster,
      applicationType: applicationType,
      customerId: customerId
    };

  } catch (e) {
    console.error('再申請データ取得エラー:', e);
    throw new Error(`再申請データの取得に失敗しました: ${e.message}`);
  }
}

/**
 * 再申請を作成
 * @param {string} originalApplicationId - 元の申請ID
 * @param {Object} modifiedData - 修正されたデータ
 * @param {string} applicantName - 申請者名
 * @param {string} applicantEmail - 申請者メール
 * @return {Object} 結果
 */
function resubmitApplication(originalApplicationId, modifiedData, applicantName, applicantEmail) {
  try {
    // 元の申請データを取得
    const originalApp = getApplicationForResubmit(originalApplicationId);

    // 申請管理シートから元申請の情報を取得
    const appSheet = SPREADSHEET.getSheetByName('申請管理');
    if (!appSheet) throw new Error('申請管理シートが見つかりません。');

    const appValues = appSheet.getDataRange().getValues();
    const appHeader = appValues.shift();

    const appIdCol = appHeader.indexOf('申請ID');
    const resubmitCountCol = appHeader.indexOf('再申請回数');
    const typeCol = appHeader.indexOf('申請種別');
    const customerNameCol = appHeader.indexOf('対象顧客名');

    const originalRow = appValues.find(r => String(r[appIdCol]) === String(originalApplicationId));
    if (!originalRow) throw new Error('元の申請が見つかりません。');

    const originalResubmitCount = resubmitCountCol !== -1 ? (originalRow[resubmitCountCol] || 0) : 0;
    const applicationType = typeCol !== -1 ? originalRow[typeCol] : '顧客情報修正';

    // 元データに修正データを上書き
    const finalData = { ...originalApp.applicationData, ...modifiedData };

    // 申請者情報を追加
    finalData.applicantName = applicantName;
    finalData.applicantEmail = applicantEmail;

    // 顧客名を取得（元申請の対象顧客名を優先）
    const customerName = (customerNameCol !== -1 && originalRow[customerNameCol])
      ? originalRow[customerNameCol]
      : (finalData['得意先名'] || originalApp.originalData['得意先名'] || '');

    // 新しい申請として登録
    const appData = {
      type: applicationType,
      customerName: customerName,
      customerId: originalApp.customerId,
      payload: {
        newData: finalData,
        originalData: originalApp.originalData
      }
    };

    // addApplication を呼び出して新規申請を作成
    const result = addApplication(appData);

    if (result.status === 'success') {
      // 新しく作成された申請IDを取得（最新の申請ID）
      const lastRow = appSheet.getLastRow();
      const newAppId = appSheet.getRange(lastRow, appIdCol + 1).getValue();

      // 再申請情報を書き込み
      const groupIdCol = appHeader.indexOf('申請グループID');
      const parentIdCol = appHeader.indexOf('元申請ID');

      if (groupIdCol !== -1) {
        appSheet.getRange(lastRow, groupIdCol + 1).setValue(originalApp.申請グループID);
      }
      if (parentIdCol !== -1) {
        appSheet.getRange(lastRow, parentIdCol + 1).setValue(originalApplicationId);
      }
      if (resubmitCountCol !== -1) {
        appSheet.getRange(lastRow, resubmitCountCol + 1).setValue(Number(originalResubmitCount) + 1);
      }

      return {
        status: 'success',
        message: '再申請が完了しました。',
        newApplicationId: newAppId
      };
    } else {
      throw new Error(result.message);
    }

  } catch (e) {
    console.error('再申請エラー:', e);
    return {
      status: 'error',
      message: `再申請に失敗しました: ${e.message}`
    };
  }
}

/**
 * 通常の addApplication で作成した申請を再申請として紐付けます。
 * @param {string} originalApplicationId - 元の申請ID
 * @param {number|string} newApplicationId - 新しく作成された申請ID
 * @return {Object} 結果
 */
function markAsResubmit(originalApplicationId, newApplicationId) {
  try {
    const appSheet = SPREADSHEET.getSheetByName('申請管理');
    if (!appSheet) throw new Error('申請管理シートが見つかりません。');

    const appValues = appSheet.getDataRange().getValues();
    const appHeader = appValues.shift();

    const appIdCol = appHeader.indexOf('申請ID');
    const resubmitCountCol = appHeader.indexOf('再申請回数');
    const groupIdCol = appHeader.indexOf('申請グループID');
    const parentIdCol = appHeader.indexOf('元申請ID');

    const originalRowIndex = appValues.findIndex(r => String(r[appIdCol]) === String(originalApplicationId));
    if (originalRowIndex === -1) throw new Error('元の申請が見つかりません: ' + originalApplicationId);
    const originalRow = appValues[originalRowIndex];

    const originalResubmitCount = resubmitCountCol !== -1 ? (originalRow[resubmitCountCol] || 0) : 0;
    const groupId = groupIdCol !== -1 && originalRow[groupIdCol]
      ? originalRow[groupIdCol]
      : originalApplicationId;

    const newRowIndex = appValues.findIndex(r => String(r[appIdCol]) === String(newApplicationId));
    if (newRowIndex === -1) throw new Error('新しい申請が見つかりません: ' + newApplicationId);
    const newSheetRow = newRowIndex + 2;

    if (groupIdCol !== -1) appSheet.getRange(newSheetRow, groupIdCol + 1).setValue(groupId);
    if (parentIdCol !== -1) appSheet.getRange(newSheetRow, parentIdCol + 1).setValue(originalApplicationId);
    if (resubmitCountCol !== -1) appSheet.getRange(newSheetRow, resubmitCountCol + 1).setValue(Number(originalResubmitCount) + 1);

    return { status: 'success' };
  } catch (e) {
    console.error('[markAsResubmit] エラー:', e);
    return { status: 'error', message: e.message };
  }
}

/**
 * 申請グループの全履歴を取得
 * @param {string} groupId - 申請グループID
 * @return {Array} 申請履歴（時系列順）
 */
function getApplicationHistory(groupId) {
  try {
    const appSheet = SPREADSHEET.getSheetByName('申請管理');
    if (!appSheet) return [];

    const values = appSheet.getDataRange().getValues();
    if (values.length <= 1) return [];

    const header = values.shift();
    const colIndex = header.reduce((acc, col, i) => (acc[col] = i, acc), {});

    const groupIdCol = colIndex['申請グループID'];

    // 同じ申請グループIDの申請を全て抽出
    const history = values
      .filter(row => {
        const rowGroupId = groupIdCol !== undefined && row[groupIdCol]
          ? row[groupIdCol]
          : row[colIndex['申請ID']]; // 後方互換性
        return String(rowGroupId) === String(groupId);
      })
      .map(row => ({
        申請ID: row[colIndex['申請ID']],
        申請日時: Utilities.formatDate(new Date(row[colIndex['申請日時']]), "JST", "yyyy/MM/dd HH:mm"),
        申請種別: row[colIndex['申請種別']],
        対象顧客名: row[colIndex['対象顧客名']],
        ステータス: row[colIndex['ステータス']],
        承認段階: row[colIndex['承認段階']] || row[colIndex['ステータス']],
        再申請回数: row[colIndex['再申請回数']] || 0,
        却下理由: row[colIndex['却下理由']] || ''
      }))
      .sort((a, b) => a.申請ID - b.申請ID); // 申請ID順（時系列順）

    return history;

  } catch (e) {
    console.error('申請履歴取得エラー:', e);
    return [];
  }
}


// =============================================
// ダッシュボードデータ取得関数
// =============================================

/**
 * ダッシュボードに表示するデータを取得します（キャッシュ付き）
 * @return {Object} ダッシュボードデータ
 */
function getDashboardData() {
  try {
    const cache = CacheService.getScriptCache();
    const cached = cache.get(CACHE_KEY_DASHBOARD);

    if (cached) {
      console.log('Cache hit: dashboard');
      return JSON.parse(cached);
    }

    console.log('Cache miss: dashboard');
    const data = fetchDashboardData();
    cache.put(CACHE_KEY_DASHBOARD, JSON.stringify(data), CACHE_DURATION);
    return data;

  } catch (e) {
    console.error('ダッシュボードデータ取得エラー:', e);
    // キャッシュエラー時は直接取得
    return fetchDashboardData();
  }
}

/**
 * ダッシュボードに表示するデータをスプレッドシートから取得します（内部関数）
 * @return {Object} ダッシュボードデータ
 */
function fetchDashboardData() {
  try {
    const meetingSheet = SPREADSHEET.getSheetByName('商談管理');
    if (!meetingSheet) throw new Error('商談管理シートが見つかりません。');

    const meetingValues = meetingSheet.getDataRange().getValues();
    if (meetingValues.length <= 1) {
      return getEmptyDashboardData();
    }

    const meetingHeader = meetingValues.shift();
    const meetingColIndex = meetingHeader.reduce((acc, col, i) => (acc[col] = i, acc), {});

    const customerSheet = CUSTOMER_SHEET;
    if (!customerSheet) throw new Error('得意先マスタシートが見つかりません。');

    const customerValues = customerSheet.getDataRange().getValues();
    const customerHeader = customerValues.shift();
    const customerColIndex = customerHeader.reduce((acc, col, i) => (acc[col] = i, acc), {});

    const customerMap = {};
    customerValues.forEach(row => {
      const customerId = cleanSingleQuotes(row[customerColIndex['得意先コード']]);
      const handler = row[customerColIndex['営業担当者名称']];

      // 「閉店・取引停止」の顧客を除外
      if (handler === '閉店・取引停止') return;

      // OLD登録済み顧客を除外
      const customerName = String(row[customerColIndex['得意先名称']] || '');
      const rankValue = String(row[customerColIndex['得意先ランク区分名称']] || '').trim();
      const deleteFlag = row[customerColIndex['削除フラグ']];
      const isHidden = customerName.includes('Ｏ＿') || rankValue === 'ＯＬＤ' || (deleteFlag && Number(deleteFlag) >= 1);
      if (isHidden) return;

      customerMap[customerId] = {
        name: row[customerColIndex['得意先名称']],
        rank: row[customerColIndex['得意先ランク区分名称']],
        handler: handler,
        handlerCode: cleanSingleQuotes(row[customerColIndex['営業担当者コード']]),
        department: row[customerColIndex['拠点名称']]
      };
    });

    const employees = getEmployeesWithSort();

    // ★修正: 担当者名から大区分・部署を引けるようにマップを作成
    const employeeNameToDivisionMap = {};
    const employeeNameToDepartmentMap = {};
    employees.forEach(emp => {
      employeeNameToDivisionMap[emp.employeeName] = emp.division || '';
      employeeNameToDepartmentMap[emp.employeeName] = emp.departmentName || '';
    });

    const allMeetings = meetingValues.map(row => {
      const customerId = cleanSingleQuotes(row[meetingColIndex['得意先コード']]);
      const customer = customerMap[customerId] || {};
      // 企業担当者コード列が存在すれば直接使用（customerMap結合不要）、なければ customerMap にフォールバック
      const handlerCode = meetingColIndex['企業担当者コード'] !== undefined && String(row[meetingColIndex['企業担当者コード']] || '').trim()
        ? cleanSingleQuotes(String(row[meetingColIndex['企業担当者コード']] || '').trim())
        : customer.handlerCode || '';
      const handlerName = (() => { const h = row[meetingColIndex['担当者']]; return h ? h.replace(/　/g, ' ') : h; })(); // 商談管理シートの担当者（全角スペース正規化）

      return {
        customerId: customerId,
        customerName: row[meetingColIndex['企業名']],
        scheduleDate: (() => {
          const v = row[meetingColIndex['商談予定日']];
          if (!v) return null;
          const d = new Date(v);
          if (isNaN(d.getTime()) || d.getFullYear() < 1990) return null;
          return Utilities.formatDate(d, "JST", "yyyy-MM-dd");
        })(),
        actualDate: row[meetingColIndex['商談実施日']]
          ? Utilities.formatDate(new Date(row[meetingColIndex['商談実施日']]), "JST", "yyyy-MM-dd")
          : null,
        handler: handlerName,
        handlerCode: handlerCode,
        meetingHandlerCode: meetingColIndex['担当者コード'] !== undefined ? cleanSingleQuotes(String(row[meetingColIndex['担当者コード']] || '').trim()) : '',
        department: employeeNameToDepartmentMap[handlerName] || customer.department || '', // ★修正: 担当者の部署を優先
        division: employeeNameToDivisionMap[handlerName] || '', // ★修正: 担当者の大区分を使用
        purpose: row[meetingColIndex['商談目的']],
        appointment: row[meetingColIndex['アポイント有無']],
        result: row[meetingColIndex['結果']],
        rank: customer.rank || '-'
      };
    });

    // ★修正: unvisitedCustomersはフロントエンド側で計算、rankAlertCustomersは再度有効化
    // const unvisitedCustomers = calculateUnvisitedCustomers(allMeetings);
    const achievementData = calculateAchievementData(allMeetings, employees);
    const rankAlertCustomers = calculateRankAlertCustomers(customerValues, customerHeader, allMeetings, employees); // ★再有効化
    const appointmentData = calculateAppointmentData(allMeetings, employees);
    const purposeData = calculatePurposeData(allMeetings, employees);

    const result = {
      employees: employees.map(e => ({
        code: e.employeeCode,
        name: e.employeeName,
        department: e.departmentName,
        division: e.division // ★追加
      })),
      allMeetings: allMeetings,
      // unvisitedCustomers: unvisitedCustomers, // ★削除: フロントエンド側で計算
      achievementData: achievementData,
      rankAlertCustomers: rankAlertCustomers, // ★再有効化
      appointmentData: appointmentData,
      purposeData: purposeData
    };

    // ★デバッグ: データサイズを確認
    const jsonString = JSON.stringify(result);
    const sizeInBytes = jsonString.length;
    const sizeInKB = (sizeInBytes / 1024).toFixed(2);
    console.log(`📊 ダッシュボードデータサイズ: ${sizeInKB} KB (${sizeInBytes} bytes)`);
    console.log(`  - employees: ${employees.length}件`);
    console.log(`  - allMeetings: ${allMeetings.length}件`);
    console.log(`  - achievementData.byEmployee: ${achievementData.byEmployee.length}件`);
    console.log(`  - appointmentData.byEmployee: ${appointmentData.byEmployee.length}件`);
    console.log(`  - purposeData.byEmployee: ${purposeData.byEmployee.length}件`);

    return result;

  } catch (e) {
    console.error('ダッシュボードデータ取得エラー:', e);
    throw new Error('ダッシュボードデータの取得に失敗しました: ' + e.message);
  }
}

function calculateUnvisitedCustomers(allMeetings) {
  return allMeetings
    .filter(m => m.scheduleDate && !m.actualDate)
    .map(m => ({
      customerName: m.customerName,
      scheduleDate: m.scheduleDate,
      handler: m.handler,
      handlerCode: m.handlerCode,
      department: m.department,
      division: m.division, // ★追加
      purpose: m.purpose
    }));
}

function calculateAchievementData(allMeetings, employees) {
  // 今年度の期間を計算（6月始まり）
  const now = new Date();
  const currentYear = now.getFullYear();
  const currentMonth = now.getMonth(); // 0-11

  let fiscalYearStart, fiscalYearEnd;
  if (currentMonth >= 5) { // 6月以降（月は0始まりなので5が6月）
    fiscalYearStart = new Date(currentYear, 5, 1); // 今年の6月1日
    fiscalYearEnd = new Date(currentYear + 1, 4, 31); // 来年の5月31日
  } else {
    fiscalYearStart = new Date(currentYear - 1, 5, 1); // 前年の6月1日
    fiscalYearEnd = new Date(currentYear, 4, 31); // 今年の5月31日
  }

  // 累計達成率（今年度のデータ）
  const fiscalYearMeetings = allMeetings.filter(m => {
    if (!m.scheduleDate) return false;
    const date = new Date(m.scheduleDate);
    return date >= fiscalYearStart && date <= fiscalYearEnd;
  });
  const totalScheduled = fiscalYearMeetings.length;
  const totalCompleted = fiscalYearMeetings.filter(m => m.actualDate).length;
  const totalRate = totalScheduled > 0 ? Math.round((totalCompleted / totalScheduled) * 100) : 0;

  // 今月の達成率
  const thisMonthMeetings = allMeetings.filter(m => {
    if (!m.scheduleDate) return false;
    const date = new Date(m.scheduleDate);
    return date.getFullYear() === now.getFullYear() && date.getMonth() === now.getMonth();
  });
  const monthlyScheduled = thisMonthMeetings.length;
  const monthlyCompleted = thisMonthMeetings.filter(m => m.actualDate).length;
  const monthlyRate = monthlyScheduled > 0 ? Math.round((monthlyCompleted / monthlyScheduled) * 100) : 0;

  const byEmployee = employees
    .filter(emp => emp.employeeName !== '閉店・取引停止') // 「閉店・取引停止」を除外
    .map(emp => {
      const empMeetings = allMeetings.filter(m => m.handler === emp.employeeName);
      const empScheduled = empMeetings.filter(m => m.scheduleDate).length; // 予定登録されている件数
      const empCompletedWithSchedule = empMeetings.filter(m => m.scheduleDate && m.actualDate).length; // 予定と実績両方あるもの
      const empTotalActual = empMeetings.filter(m => m.actualDate).length; // 実績が登録されている全件数（予定なしも含む）
      const empUnachieved = empScheduled - empCompletedWithSchedule; // 未達成件数
      const empRate = empScheduled > 0 ? Math.round((empCompletedWithSchedule / empScheduled) * 100) : 0;

      return {
        code: emp.employeeCode,
        name: emp.employeeName,
        department: emp.departmentName,
        division: emp.division,
        scheduled: empScheduled, // 商談予定件数
        completed: empCompletedWithSchedule, // 商談実績件数（予定と実績両方）
        totalActual: empTotalActual, // 述べ商談件数（実績がある全件数）
        unachieved: empUnachieved,
        rate: empRate
      };
    });

  // ★修正: 今期（6月始まり）の全月を表示
  const history = [];
  const startMonthDate = new Date(fiscalYearStart.getFullYear(), fiscalYearStart.getMonth(), 1);
  const endMonthDate = new Date(now.getFullYear(), now.getMonth(), 1);

  let achievementCurrentMonth = new Date(startMonthDate);
  while (achievementCurrentMonth <= endMonthDate) {
    const year = achievementCurrentMonth.getFullYear();
    const month = achievementCurrentMonth.getMonth();

    const monthMeetings = allMeetings.filter(m => {
      if (!m.scheduleDate) return false;
      const date = new Date(m.scheduleDate);
      return date.getFullYear() === year && date.getMonth() === month;
    });

    const monthScheduled = monthMeetings.length;
    const monthCompleted = monthMeetings.filter(m => m.actualDate).length;
    const monthRate = monthScheduled > 0 ? Math.round((monthCompleted / monthScheduled) * 100) : 0;

    history.push({
      month: year + '/' + (month + 1),
      rate: monthRate
    });

    achievementCurrentMonth.setMonth(achievementCurrentMonth.getMonth() + 1);
  }

  return {
    total: totalRate,
    monthly: monthlyRate,
    byEmployee: byEmployee,
    history: history
  };
}

/**
 * 訪問頻度文字列を年間訪問回数に変換します
 * @param {string} frequencyStr - 訪問頻度文字列（例: "6回/年", "2ヶ月に1回"）
 * @return {number|null} 年間訪問回数、解析できない場合はnull
 */
function parseVisitFrequency(frequencyStr) {
  if (!frequencyStr || typeof frequencyStr !== 'string') return null;

  const str = frequencyStr.trim();
  if (!str) return null;

  // 「○回/年」形式（例: "6回/年", "4回/年　以上"）
  const yearMatch = str.match(/(\d+)\s*回\s*[\/／]\s*年/);
  if (yearMatch) {
    return parseInt(yearMatch[1], 10);
  }

  // 「○ヶ月に1回」形式（例: "2ヶ月に1回", "3ヶ月に1回"）
  const monthMatch = str.match(/(\d+)\s*[ヶケか]*月\s*に\s*1\s*回/);
  if (monthMatch) {
    const months = parseInt(monthMatch[1], 10);
    return Math.ceil(12 / months); // 12ヶ月÷○ヶ月 = 年間回数
  }

  // 「毎月○回」形式（例: "毎月1回"）
  const monthlyMatch = str.match(/毎月\s*(\d+)\s*回/);
  if (monthlyMatch) {
    return parseInt(monthlyMatch[1], 10) * 12;
  }

  return null;
}

function calculateRankAlertCustomers(customerValues, customerHeader, allMeetings, employees) {
  const colIndex = customerHeader.reduce((acc, col, i) => (acc[col] = i, acc), {});

  // ★追加: 担当者コードから大区分を引けるようにマップを作成
  const employeeMap = {};
  employees.forEach(emp => {
    employeeMap[emp.employeeCode] = emp.division || '';
  });

  // ★修正: ランクマスタから訪問頻度を取得
  const rankSheet = SPREADSHEET.getSheetByName('ランクマスタ');
  const rankRequirements = {};
  if (rankSheet) {
    const rankValues = rankSheet.getDataRange().getValues();
    const rankHeader = rankValues.shift();
    const rankColIndex = rankHeader.reduce((acc, col, i) => (acc[col] = i, acc), {});

    // ランク名称の列名を特定（複数の列名に対応）
    const rankNameCol = rankColIndex['名称_1'] !== undefined ? '名称_1' :
                        rankColIndex['ランク名称'] !== undefined ? 'ランク名称' :
                        rankColIndex['得意先ランク区分名称'] !== undefined ? '得意先ランク区分名称' : null;

    const frequencyCol = rankColIndex['訪問頻度'];

    if (rankNameCol !== null && frequencyCol !== undefined) {
      rankValues.forEach(row => {
        const rankName = row[rankColIndex[rankNameCol]];
        const frequencyStr = row[frequencyCol];
        if (rankName) {
          const yearlyCount = parseVisitFrequency(frequencyStr);
          if (yearlyCount !== null) {
            rankRequirements[rankName] = yearlyCount;
          }
        }
      });
      console.log('ランクマスタから訪問頻度を読み込みました:', rankRequirements);
    } else {
      console.warn('ランクマスタの列名が見つかりません。rankNameCol:', rankNameCol, 'frequencyCol:', frequencyCol);
    }
  } else {
    console.warn('ランクマスタシートが見つかりません。');
  }

  const now = new Date();
  // 今期（6月始まり）の期間を計算
  const currentYear = now.getFullYear();
  const currentMonth = now.getMonth(); // 0-indexed
  const fiscalYearStart = currentMonth >= 5 ? new Date(currentYear, 5, 1) : new Date(currentYear - 1, 5, 1);
  const fiscalYearEnd = currentMonth >= 5 ? new Date(currentYear + 1, 4, 31, 23, 59, 59) : new Date(currentYear, 4, 31, 23, 59, 59);

  const alertCustomers = [];

  // ★デバッグ: 処理開始ログ
  let totalRankCustomers = 0;
  let checkedCustomers = 0;

  customerValues.forEach(row => {
    const customerId = cleanSingleQuotes(row[colIndex['得意先コード']]);
    const customerName = row[colIndex['得意先名称']];
    const rank = row[colIndex['得意先ランク区分名称']];
    const handler = row[colIndex['営業担当者名称']];
    const handlerCode = cleanSingleQuotes(row[colIndex['営業担当者コード']]);
    const department = row[colIndex['拠点名称']];

    // 「閉店・取引停止」の顧客を除外
    if (handler === '閉店・取引停止') return;

    const requiredCount = rankRequirements[rank];
    if (!requiredCount) return;

    totalRankCustomers++;

    // 今期（6月始まり）かつ担当者コード一致の訪問回数をカウント
    const visitCount = allMeetings.filter(m => {
      if (m.customerId !== customerId || !m.actualDate) return false;
      const date = new Date(m.actualDate);
      if (date < fiscalYearStart || date > fiscalYearEnd) return false;
      const mCode = cleanSingleQuotes(String(m.meetingHandlerCode || '').trim());
      const cCode = cleanSingleQuotes(String(handlerCode || '').trim());
      return mCode !== '' && cCode !== '' && mCode === cCode;
    }).length;

    checkedCustomers++;

    if (visitCount < requiredCount) {
      alertCustomers.push({
        customerName: customerName,
        rank: rank,
        handler: handler,
        handlerCode: handlerCode,
        department: department,
        division: employeeMap[handlerCode] || '', // ★追加
        visitCount: visitCount,
        requiredCount: requiredCount
      });
    }
  });

  // ★デバッグログ
  console.log(`ランクアラート計算完了: 対象顧客=${totalRankCustomers}件, チェック済み=${checkedCustomers}件, アラート顧客=${alertCustomers.length}件`);
  if (alertCustomers.length > 0) {
    console.log(`ランクアラート顧客例: ${alertCustomers.slice(0, 3).map(c => `${c.customerName}(${c.rank}, ${c.visitCount}/${c.requiredCount}回)`).join(', ')}`);
  }

  return alertCustomers;
}

function calculateAppointmentData(allMeetings, employees) {
  // 今年度の期間を計算（6月始まり）
  const now = new Date();
  const currentYear = now.getFullYear();
  const currentMonth = now.getMonth(); // 0-11

  let fiscalYearStart, fiscalYearEnd;
  if (currentMonth >= 5) { // 6月以降（月は0始まりなので5が6月）
    fiscalYearStart = new Date(currentYear, 5, 1); // 今年の6月1日
    fiscalYearEnd = new Date(currentYear + 1, 4, 31); // 来年の5月31日
  } else {
    fiscalYearStart = new Date(currentYear - 1, 5, 1); // 前年の6月1日
    fiscalYearEnd = new Date(currentYear, 4, 31); // 今年の5月31日
  }

  // 累計アポ獲得率（今年度のデータ）
  const fiscalYearCompleted = allMeetings.filter(m => {
    if (!m.actualDate) return false;
    const date = new Date(m.actualDate);
    return date >= fiscalYearStart && date <= fiscalYearEnd;
  }).length;
  const fiscalYearWithAppointment = allMeetings.filter(m => {
    if (!m.actualDate || m.appointment !== '有') return false;
    const date = new Date(m.actualDate);
    return date >= fiscalYearStart && date <= fiscalYearEnd;
  }).length;
  const totalRate = fiscalYearCompleted > 0 ? Math.round((fiscalYearWithAppointment / fiscalYearCompleted) * 100) : 0;

  // 今月のアポ獲得率
  const thisMonthCompleted = allMeetings.filter(m => {
    if (!m.actualDate) return false;
    const date = new Date(m.actualDate);
    return date.getFullYear() === now.getFullYear() && date.getMonth() === now.getMonth();
  }).length;
  const thisMonthWithAppointment = allMeetings.filter(m => {
    if (!m.actualDate || m.appointment !== '有') return false;
    const date = new Date(m.actualDate);
    return date.getFullYear() === now.getFullYear() && date.getMonth() === now.getMonth();
  }).length;
  const monthlyRate = thisMonthCompleted > 0 ? Math.round((thisMonthWithAppointment / thisMonthCompleted) * 100) : 0;

  // 担当者別のアポ獲得率
  const byEmployee = employees.map(emp => {
    const empCompleted = allMeetings.filter(m =>
      m.handlerCode === emp.employeeCode && m.actualDate
    ).length;
    const empWithAppointment = allMeetings.filter(m =>
      m.handlerCode === emp.employeeCode && m.actualDate && m.appointment === '有'
    ).length;
    const empRate = empCompleted > 0 ? Math.round((empWithAppointment / empCompleted) * 100) : 0;

    return {
      code: emp.employeeCode,
      name: emp.employeeName,
      department: emp.departmentName,
      division: emp.division,
      visitCount: empCompleted, // ★追加: 訪問件数
      appointmentCount: empWithAppointment, // ★追加: うちアポ件数
      rate: empRate
    };
  });

  // ★修正: 今期（6月始まり）の全月を表示
  const history = [];
  const apptStartMonth = new Date(fiscalYearStart.getFullYear(), fiscalYearStart.getMonth(), 1);
  const apptEndMonth = new Date(now.getFullYear(), now.getMonth(), 1);

  let appointmentCurrentMonth = new Date(apptStartMonth);
  while (appointmentCurrentMonth <= apptEndMonth) {
    const year = appointmentCurrentMonth.getFullYear();
    const month = appointmentCurrentMonth.getMonth();

    const monthCompleted = allMeetings.filter(m => {
      if (!m.actualDate) return false;
      const date = new Date(m.actualDate);
      return date.getFullYear() === year && date.getMonth() === month;
    }).length;
    const monthWithAppointment = allMeetings.filter(m => {
      if (!m.actualDate || m.appointment !== '有') return false;
      const date = new Date(m.actualDate);
      return date.getFullYear() === year && date.getMonth() === month;
    }).length;
    const monthRate = monthCompleted > 0 ? Math.round((monthWithAppointment / monthCompleted) * 100) : 0;

    history.push({
      month: year + '/' + (month + 1),
      rate: monthRate
    });

    appointmentCurrentMonth.setMonth(appointmentCurrentMonth.getMonth() + 1);
  }

  return {
    total: totalRate,
    monthly: monthlyRate,
    byEmployee: byEmployee,
    history: history
  };
}

/**
 * 商談目的別のデータを集計します
 */
function calculatePurposeData(allMeetings, employees) {
  const purposeTypes = ['定期商談', 'プラスワン', 'クレーム対応', 'その他', '新規営業'];

  // 担当者別の商談目的集計
  const byEmployee = employees.map(emp => {
    const empMeetings = allMeetings.filter(m => m.handler === emp.employeeName);

    const purposes = {};
    purposeTypes.forEach(purpose => {
      purposes[purpose] = {
        schedule: empMeetings.filter(m => m.scheduleDate && m.purpose === purpose).length,
        actual: empMeetings.filter(m => m.actualDate && m.purpose === purpose).length
      };
    });

    return {
      code: emp.employeeCode,
      name: emp.employeeName,
      department: emp.departmentName,
      division: emp.division,
      purposes: purposes
    };
  });

  // 全体の商談目的分布
  const distribution = {};
  purposeTypes.forEach(purpose => {
    distribution[purpose] = {
      schedule: allMeetings.filter(m => m.scheduleDate && m.purpose === purpose).length,
      actual: allMeetings.filter(m => m.actualDate && m.purpose === purpose).length
    };
  });

  return {
    byEmployee: byEmployee,
    distribution: distribution
  };
}

function getEmptyDashboardData() {
  return {
    employees: [],
    allMeetings: [],
    unvisitedCustomers: [],
    achievementData: { total: 0, monthly: 0, byEmployee: [], history: [] },
    rankAlertCustomers: [],
    appointmentData: { total: 0, monthly: 0, byEmployee: [], history: [] },
    purposeData: { byEmployee: [], distribution: {} }
  };
}

// =============================================
// キャッシュクリア関数
// =============================================

/**
 * 顧客データのキャッシュをクリアします
 */
function clearCustomersCache() {
  try {
    const cache = CacheService.getScriptCache();
    cache.remove(CACHE_KEY_CUSTOMERS);
    console.log('Customers cache cleared');
  } catch (e) {
    console.error('顧客キャッシュクリアエラー:', e);
  }
}

/**
 * 商談データのキャッシュをクリアします
 */
function clearMeetingsCache() {
  try {
    const cache = CacheService.getScriptCache();
    cache.remove(CACHE_KEY_MEETINGS);
    console.log('Meetings cache cleared');
  } catch (e) {
    console.error('商談キャッシュクリアエラー:', e);
  }
}

/**
 * ダッシュボードデータのキャッシュをクリアします
 */
function clearDashboardCache() {
  try {
    const cache = CacheService.getScriptCache();
    cache.remove(CACHE_KEY_DASHBOARD);
    console.log('Dashboard cache cleared');
  } catch (e) {
    console.error('ダッシュボードキャッシュクリアエラー:', e);
  }
}

/**
 * すべてのデータキャッシュをクリアします
 */
function clearAllDataCache() {
  clearCustomersCache();
  clearMeetingsCache();
  clearDashboardCache();
  console.log('All data cache cleared');
}

// =============================================
// マスタ管理機能
// =============================================

/**
 * マスタシートのデータを取得します
 * @param {string} sheetName - シート名（社員マスタ、ランクマスタ、業種マスタ）
 * @return {Array<Array>} シートの全データ（ヘッダー含む）
 */
function getMasterSheetData(sheetName) {
  try {
    const sheet = SPREADSHEET.getSheetByName(sheetName);
    if (!sheet) {
      throw new Error(`シート「${sheetName}」が見つかりません。`);
    }

    const data = sheet.getDataRange().getValues();
    return data;

  } catch (e) {
    console.error('getMasterSheetData error:', e);
    throw new Error('マスタデータの取得に失敗しました: ' + e.message);
  }
}

/**
 * マスタシートのデータを保存します
 * @param {string} sheetName - シート名
 * @param {Array<Array>} data - 保存するデータ（ヘッダー含む）
 * @return {Object} 処理結果
 */
function saveMasterSheetData(sheetName, data) {
  try {
    const sheet = SPREADSHEET.getSheetByName(sheetName);
    if (!sheet) {
      throw new Error(`シート「${sheetName}」が見つかりません。`);
    }

    // 変更履歴記録用: 保存前のデータを取得
    let oldData = [];
    try {
      const lastRow = sheet.getLastRow();
      const lastCol = sheet.getLastColumn();
      if (lastRow > 0 && lastCol > 0) {
        oldData = sheet.getRange(1, 1, lastRow, lastCol).getValues();
      }
    } catch (e) {
      console.warn('保存前データの取得に失敗しました:', e);
    }

    // シートをクリア
    sheet.clear();

    // データを書き込み
    if (data && data.length > 0) {
      sheet.getRange(1, 1, data.length, data[0].length).setValues(data);
    }

    // 関連キャッシュをクリア
    clearAllCaches();

    // マスタ更新履歴に記録
    try {
      const updater = Session.getActiveUser().getEmail() || 'システム';
      const changeDescription = generateMasterChangeDescription(sheetName, oldData, data);
      addMasterUpdateHistory(updater, sheetName, changeDescription);
    } catch (historyError) {
      console.error('[saveMasterSheetData] マスタ更新履歴の記録に失敗しました:', historyError);
      // 履歴記録の失敗は保存処理自体には影響させない
    }

    return { status: 'success', message: 'マスタデータを保存しました。' };

  } catch (e) {
    console.error('saveMasterSheetData error:', e);
    throw new Error('マスタデータの保存に失敗しました: ' + e.message);
  }
}

// =============================================
// 申請の自動削除機能
// =============================================

/**
 * 半年（180日）経過した申請を自動削除します
 * この関数は時間駆動型トリガーで毎日実行されることを想定しています
 *
 * 実行方法：
 * 1. Google Apps Scriptエディタで「トリガー」を開く
 * 2. 「トリガーを追加」をクリック
 * 3. 以下の設定を行う：
 *    - 実行する関数: deleteOldApplications
 *    - イベントのソース: 時間主導型
 *    - 時間ベースのトリガー: 日タイマー
 *    - 時刻: 午前2時～3時（深夜の負荷が少ない時間帯を推奨）
 * 4. 「保存」をクリック
 *
 * @return {Object} 削除結果 {success: boolean, deletedCount: number, message: string}
 */
function deleteOldApplications() {
  try {
    console.log('[deleteOldApplications] 半年（180日）経過した申請の削除処理を開始します');

    // 申請管理シートを取得
    const appSheet = SPREADSHEET.getSheetByName('申請管理');
    if (!appSheet) {
      throw new Error('申請管理シートが見つかりません。');
    }

    // 申請データ_顧客シートを取得
    const customerDataSheet = SPREADSHEET.getSheetByName('申請データ_顧客');
    // 申請データ_単価シートを取得
    const priceDataSheet = SPREADSHEET.getSheetByName('申請データ_単価');

    // 現在日時から180日前の日時を計算
    const halfYearAgo = new Date();
    halfYearAgo.setDate(halfYearAgo.getDate() - 180);
    const halfYearAgoStr = Utilities.formatDate(halfYearAgo, 'JST', 'yyyy/MM/dd HH:mm:ss');
    console.log('[deleteOldApplications] 削除対象: ' + halfYearAgoStr + ' より前の申請');

    // 申請管理シートのデータを取得
    const values = appSheet.getDataRange().getValues();
    if (values.length <= 1) {
      console.log('[deleteOldApplications] 削除対象の申請がありません（データなし）');
      return { success: true, deletedCount: 0, message: '削除対象の申請はありませんでした。' };
    }

    const header = values.shift();
    const appIdCol = header.indexOf('申請ID');
    const appDateCol = header.indexOf('申請日時');

    if (appIdCol === -1 || appDateCol === -1) {
      throw new Error('申請管理シートに「申請ID」または「申請日時」列が見つかりません。');
    }

    // 削除対象の申請IDを特定
    const deleteTargetIds = [];
    const deleteTargetRows = [];

    values.forEach(function(row, index) {
      const appDate = row[appDateCol];
      const appId = row[appIdCol];

      if (appDate && appDate instanceof Date) {
        if (appDate < halfYearAgo) {
          deleteTargetIds.push(appId);
          deleteTargetRows.push(index + 2);
          const appDateStr = Utilities.formatDate(appDate, 'JST', 'yyyy/MM/dd HH:mm:ss');
          console.log('[deleteOldApplications] 削除対象: 申請ID=' + appId + ', 申請日時=' + appDateStr);
        }
      }
    });

    if (deleteTargetIds.length === 0) {
      console.log('[deleteOldApplications] 削除対象の申請がありません');
      return { success: true, deletedCount: 0, message: '削除対象の申請はありませんでした。' };
    }

    console.log('[deleteOldApplications] 削除対象: ' + deleteTargetIds.length + '件');

    // 申請管理シートから削除（降順で削除して行番号のズレを防ぐ）
    deleteTargetRows.sort(function(a, b) { return b - a; });
    deleteTargetRows.forEach(function(rowNum) {
      appSheet.deleteRow(rowNum);
    });

    // 申請データ_顧客シートから関連データを削除
    if (customerDataSheet) {
      deleteRelatedApplicationData(customerDataSheet, deleteTargetIds);
    }

    // 申請データ_単価シートから関連データを削除
    if (priceDataSheet) {
      deleteRelatedApplicationData(priceDataSheet, deleteTargetIds);
    }

    const message = '半年（180日）経過した申請を' + deleteTargetIds.length + '件削除しました。';
    console.log('[deleteOldApplications] ' + message);

    return {
      success: true,
      deletedCount: deleteTargetIds.length,
      message: message
    };

  } catch (e) {
    console.error('[deleteOldApplications] エラー:', e);
    return {
      success: false,
      deletedCount: 0,
      message: '削除処理中にエラーが発生しました: ' + e.message
    };
  }
}

/**
 * 申請データシート（顧客・単価）から指定された申請IDに関連する行を削除します
 * @param {Sheet} sheet - 対象シート
 * @param {Array<string>} applicationIds - 削除対象の申請ID配列
 */
function deleteRelatedApplicationData(sheet, applicationIds) {
  try {
    if (!sheet || applicationIds.length === 0) return;

    const values = sheet.getDataRange().getValues();
    if (values.length <= 1) return;

    const header = values.shift();
    const appIdCol = header.indexOf('申請ID');

    if (appIdCol === -1) {
      console.warn(sheet.getName() + 'シートに「申請ID」列が見つかりません。');
      return;
    }

    // 削除対象の行番号を特定（降順）
    const deleteRows = [];
    values.forEach(function(row, index) {
      const appId = String(row[appIdCol]);
      if (applicationIds.indexOf(appId) !== -1) {
        deleteRows.push(index + 2);
      }
    });

    // 降順で削除
    deleteRows.sort(function(a, b) { return b - a; });
    deleteRows.forEach(function(rowNum) {
      sheet.deleteRow(rowNum);
    });

    console.log('[deleteRelatedApplicationData] ' + sheet.getName() + ': ' + deleteRows.length + '行削除');

  } catch (e) {
    console.error('[deleteRelatedApplicationData] ' + sheet.getName() + 'でエラー:', e);
  }
}

/**
 * マスタ更新履歴シートに履歴を記録する共通関数
 * @param {string} updater - 更新者名
 * @param {string} targetMaster - 対象マスタ名（例：「得意先マスタ」「単価マスタ」）
 * @param {string} changes - 変更内容（簡易形式、例：「得意先名称: AAA → BBB, 住所: 東京都 → 大阪府」）
 */
function addMasterUpdateHistory(updater, targetMaster, changes) {
  try {
    const historySheet = SPREADSHEET.getSheetByName('マスタ更新履歴');
    if (!historySheet) {
      console.warn('マスタ更新履歴シートが見つかりません。履歴記録をスキップします。');
      return;
    }

    const now = new Date();
    const newRow = [
      now,           // 更新日時
      updater,       // 更新者
      targetMaster,  // 対象マスタ
      changes        // 変更内容
    ];

    historySheet.appendRow(newRow);
    console.log(`[マスタ更新履歴] ${targetMaster} - ${updater} - ${changes}`);
  } catch (e) {
    console.error('[addMasterUpdateHistory] エラー:', e);
  }
}

/**
 * 2つのオブジェクトを比較して変更内容を文字列で生成
 * @param {object} before - 変更前のデータ
 * @param {object} after - 変更後のデータ
 * @param {Array<string>} keys - 比較する項目のキー配列
 * @returns {string} 変更内容の文字列（例：「得意先名称: AAA → BBB, 住所: 東京都 → 大阪府」）
 */
function generateChangeDescription(before, after, keys) {
  const changes = [];

  keys.forEach(key => {
    const beforeValue = before[key] || '';
    const afterValue = after[key] || '';

    if (String(beforeValue) !== String(afterValue)) {
      changes.push(`${key}: ${beforeValue} → ${afterValue}`);
    }
  });

  return changes.length > 0 ? changes.join(', ') : '変更なし';
}

/**
 * マスタデータの変更内容を生成します（直接編集用）
 * @param {string} sheetName - シート名
 * @param {Array} oldData - 変更前のデータ（2次元配列）
 * @param {Array} newData - 変更後のデータ（2次元配列）
 * @return {string} 変更内容の説明
 */
function generateMasterChangeDescription(sheetName, oldData, newData) {
  try {
    const oldRowCount = oldData.length > 0 ? oldData.length - 1 : 0; // ヘッダー行を除く
    const newRowCount = newData.length > 0 ? newData.length - 1 : 0; // ヘッダー行を除く

    const changes = [];

    // 行数の変化を記録
    if (newRowCount > oldRowCount) {
      const added = newRowCount - oldRowCount;
      changes.push(`${added}行追加`);
    } else if (newRowCount < oldRowCount) {
      const deleted = oldRowCount - newRowCount;
      changes.push(`${deleted}行削除`);
    }

    // 修正された行を検出（簡易比較）
    if (oldData.length > 0 && newData.length > 0) {
      let modifiedCount = 0;
      const minRowCount = Math.min(oldData.length, newData.length);

      for (let i = 1; i < minRowCount; i++) { // ヘッダー行をスキップ
        const oldRow = oldData[i];
        const newRow = newData[i];

        // 行の内容が異なるかチェック
        if (oldRow && newRow) {
          const oldRowStr = oldRow.map(cell => String(cell || '')).join('|');
          const newRowStr = newRow.map(cell => String(cell || '')).join('|');

          if (oldRowStr !== newRowStr) {
            modifiedCount++;
          }
        }
      }

      if (modifiedCount > 0) {
        changes.push(`${modifiedCount}行修正`);
      }
    }

    // 変更内容がない場合
    if (changes.length === 0) {
      return 'マスタデータを保存（変更なし）';
    }

    return `マスタデータを直接編集: ${changes.join(', ')}（合計: ${newRowCount}行）`;

  } catch (e) {
    console.error('[generateMasterChangeDescription] エラー:', e);
    return 'マスタデータを直接編集';
  }
}

/**
 * マスタ更新履歴を取得します（ページング対応）
 * @param {number} page - ページ番号（1から開始）
 * @param {number} pageSize - 1ページあたりの件数
 * @return {Object} { data: Array, totalCount: number, totalPages: number }
 */
function getMasterUpdateHistory(page = 1, pageSize = 15) {
  try {
    const historySheet = SPREADSHEET.getSheetByName('マスタ更新履歴');
    if (!historySheet) {
      console.warn('マスタ更新履歴シートが見つかりません。');
      return { data: [], totalCount: 0, totalPages: 0 };
    }

    const values = historySheet.getDataRange().getValues();
    if (values.length <= 1) {
      return { data: [], totalCount: 0, totalPages: 0 };
    }

    const header = values.shift();

    // 日付の新しい順にソート（更新日時列でソート）
    const sortedData = values
      .filter(row => row[0]) // 更新日時が空でない行のみ
      .sort((a, b) => {
        const dateA = new Date(a[0]);
        const dateB = new Date(b[0]);
        return dateB - dateA; // 降順
      });

    const totalCount = sortedData.length;
    const totalPages = Math.ceil(totalCount / pageSize);

    // ページングで該当範囲を取得
    const startIndex = (page - 1) * pageSize;
    const endIndex = Math.min(startIndex + pageSize, totalCount);
    const pagedData = sortedData.slice(startIndex, endIndex);

    // オブジェクト配列に変換
    const result = pagedData.map(row => {
      const obj = {};
      header.forEach((colName, index) => {
        if (colName) {
          // 日付の場合はフォーマット
          if (index === 0 && row[index] instanceof Date) {
            obj[colName] = Utilities.formatDate(row[index], SPREADSHEET.getSpreadsheetTimeZone(), 'yyyy/MM/dd HH:mm:ss');
          } else {
            obj[colName] = row[index];
          }
        }
      });
      return obj;
    });

    return {
      data: result,
      totalCount: totalCount,
      totalPages: totalPages,
      currentPage: page
    };

  } catch (e) {
    console.error('[getMasterUpdateHistory] エラー:', e);
    return { data: [], totalCount: 0, totalPages: 0 };
  }
}

// =============================================
// カレンダーExcel出力機能
// =============================================

/**
 * カレンダー形式のExcelファイルを生成します
 * @param {string} targetMonth - 対象月（YYYY-MM形式）
 * @param {Object} filters - フィルタ条件 { division, department, employee }
 * @return {Object} { url, fileName, fileId }
 */
function generateCalendarExcel(targetMonth, filters) {
  try {
    console.log('[generateCalendarExcel] 開始 - targetMonth:', targetMonth, 'filters:', filters);

    // 対象月の開始日と終了日を計算
    const [year, month] = targetMonth.split('-').map(Number);
    const startDate = new Date(year, month - 1, 1);
    const endDate = new Date(year, month, 0);
    
    const startDateStr = Utilities.formatDate(startDate, Session.getScriptTimeZone(), 'yyyy-MM-dd');
    const endDateStr = Utilities.formatDate(endDate, Session.getScriptTimeZone(), 'yyyy-MM-dd');

    // 商談データを取得
    const allMeetings = getMeetings();
    
    // 社員マスタを取得（フィルタリング用）
    const employees = getAllEmployees();

    // 担当者をフィルタリング
    const filteredHandlers = new Set();
    const debugInfo = { total: 0, employeeNotFound: 0, divisionMismatch: 0, deptMismatch: 0, passed: 0 };

    allMeetings.forEach(meeting => {
      const handler = meeting.handler;
      debugInfo.total++;

      // 担当者フィルタ
      if (filters.employee && handler !== filters.employee) return;

      // 部署・大区分フィルタ
      if (filters.division || filters.department) {
        const employee = employees.find(e => e.name === handler);
        if (!employee) {
          debugInfo.employeeNotFound++;
          console.log('[generateCalendarExcel] 社員マスタに見つからない担当者:', handler);
          return; // 社員マスタに存在しない担当者は除外
        }

        if (filters.division && employee.division !== filters.division) {
          debugInfo.divisionMismatch++;
          console.log('[generateCalendarExcel] 大区分不一致:', handler, 'フィルタ:', filters.division, '実際:', employee.division);
          return;
        }

        if (filters.department && employee.department !== filters.department) {
          debugInfo.deptMismatch++;
          return;
        }
      }

      debugInfo.passed++;
      filteredHandlers.add(handler);
    });

    console.log('[generateCalendarExcel] デバッグ情報:', debugInfo);
    console.log('[generateCalendarExcel] フィルタ後の担当者数:', filteredHandlers.size);

    // 一時スプレッドシートを作成
    const fileName = `訪問カレンダー_${targetMonth}`;
    const ss = SpreadsheetApp.create(fileName);
    const sheet = ss.getSheets()[0];
    sheet.setName(targetMonth);

    // 対象月の全日付を生成
    const dates = [];
    const currentDate = new Date(startDate);
    while (currentDate <= endDate) {
      dates.push(Utilities.formatDate(currentDate, Session.getScriptTimeZone(), 'yyyy-MM-dd'));
      currentDate.setDate(currentDate.getDate() + 1);
    }

    // 担当者ごとにデータを分類
    const handlerData = {};
    // 社員マスタの表示順でソート
    const employeesOrdered = getEmployeesWithSort();
    const sortedHandlers = employeesOrdered
      .filter(emp => filteredHandlers.has(emp.employeeName))
      .map(emp => emp.employeeName)
      .concat(Array.from(filteredHandlers).filter(h => !employeesOrdered.some(e => e.employeeName === h)));

    sortedHandlers.forEach(handler => {
      handlerData[handler] = {};
      dates.forEach(date => {
        handlerData[handler][date] = {
          scheduled: [],
          actual: []
        };
      });
    });

    // 商談データを分類
    allMeetings.forEach(meeting => {
      const handler = meeting.handler;
      if (!filteredHandlers.has(handler)) return;

      const appointmentText = meeting.appointment || '未定';
      const rank = meeting.rank || '-';
      const customerName = meeting.customerName || '';
      const meetingInfo = `${appointmentText} ${rank} ${customerName}`;

      // 予定日
      if (meeting.scheduleDate && meeting.scheduleDate >= startDateStr && meeting.scheduleDate <= endDateStr) {
        if (handlerData[handler][meeting.scheduleDate]) {
          handlerData[handler][meeting.scheduleDate].scheduled.push(meetingInfo);
        }
      }

      // 実績日
      if (meeting.actualDate && meeting.actualDate >= startDateStr && meeting.actualDate <= endDateStr) {
        if (handlerData[handler][meeting.actualDate]) {
          handlerData[handler][meeting.actualDate].actual.push(meetingInfo);
        }
      }
    });

    // シートにデータを書き込み
    let currentRow = 1;

    sortedHandlers.forEach((handler, handlerIndex) => {
      if (handlerIndex > 0) {
        currentRow += 2; // 担当者間に空行
      }

      // 担当者見出し
      sheet.getRange(currentRow, 1, 1, 14).merge();
      sheet.getRange(currentRow, 1).setValue(`【担当者: ${handler}】`);
      sheet.getRange(currentRow, 1).setFontSize(12).setFontWeight('bold').setBackground('#4A90E2').setFontColor('#FFFFFF');
      currentRow++;

      // 月の最初の日が何曜日か
      const firstDayOfWeek = startDate.getDay();

      // 週ごとにカレンダーを作成
      for (let weekIndex = 0; weekIndex < Math.ceil((dates.length + firstDayOfWeek) / 7); weekIndex++) {
        const weekDates = [];

        for (let dayIndex = 0; dayIndex < 7; dayIndex++) {
          const dateIndex = weekIndex * 7 + dayIndex - firstDayOfWeek;
          if (dateIndex >= 0 && dateIndex < dates.length) {
            weekDates.push(dates[dateIndex]);
          } else {
            weekDates.push(null);
          }
        }

        // 日付ヘッダー行
        const dateHeaderRow = currentRow;
        for (let i = 0; i < 7; i++) {
          const col = i * 2 + 1;
          if (weekDates[i]) {
            const date = new Date(weekDates[i]);
            const dayLabel = `${date.getMonth() + 1}/${date.getDate()}`;
            const dayOfWeek = ['日', '月', '火', '水', '木', '金', '土'][date.getDay()];
            
            sheet.getRange(dateHeaderRow, col, 1, 2).merge();
            sheet.getRange(dateHeaderRow, col).setValue(`${dayLabel}(${dayOfWeek})`);
            sheet.getRange(dateHeaderRow, col).setHorizontalAlignment('center').setFontWeight('bold');
            
            // 土日の背景色
            if (date.getDay() === 0) { // 日曜日
              sheet.getRange(dateHeaderRow, col).setBackground('#FFE5E5');
            } else if (date.getDay() === 6) { // 土曜日
              sheet.getRange(dateHeaderRow, col).setBackground('#E5F2FF');
            } else {
              sheet.getRange(dateHeaderRow, col).setBackground('#F0F0F0');
            }
          } else {
            sheet.getRange(dateHeaderRow, col, 1, 2).merge();
            sheet.getRange(dateHeaderRow, col).setBackground('#FFFFFF');
          }
        }
        currentRow++;

        // 予定/実績ラベル行
        const labelRow = currentRow;
        for (let i = 0; i < 7; i++) {
          const col = i * 2 + 1;
          if (weekDates[i]) {
            sheet.getRange(labelRow, col).setValue('予定').setFontSize(9).setHorizontalAlignment('center').setBackground('#FFF8DC');
            sheet.getRange(labelRow, col + 1).setValue('実績').setFontSize(9).setHorizontalAlignment('center').setBackground('#E0FFE0');
          }
        }
        currentRow++;

        // データ行（最大5行分）
        const maxRows = 5;
        for (let rowIndex = 0; rowIndex < maxRows; rowIndex++) {
          for (let i = 0; i < 7; i++) {
            const col = i * 2 + 1;
            if (weekDates[i]) {
              const dateData = handlerData[handler][weekDates[i]];
              const scheduledValue = dateData.scheduled[rowIndex] || '';
              const actualValue = dateData.actual[rowIndex] || '';
              
              sheet.getRange(currentRow, col).setValue(scheduledValue).setFontSize(9).setWrap(true);
              sheet.getRange(currentRow, col + 1).setValue(actualValue).setFontSize(9).setWrap(true);
            }
          }
          currentRow++;
        }

        currentRow++; // 週間の空行
      }
    });

    // 列幅を設定
    for (let i = 1; i <= 14; i++) {
      sheet.setColumnWidth(i, 100);
    }

    // 罫線を設定
    const lastRow = currentRow - 1;
    sheet.getRange(1, 1, lastRow, 14).setBorder(true, true, true, true, true, true);

    // SpreadsheetAppの変更を確定させてからautoResize（確定前だと1人目の行高が反映されない）
    SpreadsheetApp.flush();

    // 行の高さを自動調整（折り返しテキストに対応）
    sheet.autoResizeRows(1, lastRow);

    const fileId = ss.getId();
    const downloadUrl = `https://docs.google.com/spreadsheets/d/${fileId}/export?format=xlsx`;

    console.log('[generateCalendarExcel] 完了 - fileId:', fileId);

    // 1時間後に削除するトリガーを設定
    scheduleFileDeletion(fileId);

    return {
      url: downloadUrl,
      fileName: `${fileName}.xlsx`,
      fileId: fileId
    };

  } catch (error) {
    console.error('[generateCalendarExcel] エラー:', error);
    throw new Error('Excelファイルの生成に失敗しました: ' + error.message);
  }
}

/**
 * 一時ファイルを1時間後に削除するトリガーを設定
 * @param {string} fileId - 削除するファイルのID
 */
function scheduleFileDeletion(fileId) {
  try {
    // 1時間後に実行されるトリガーを作成
    const triggerDate = new Date();
    triggerDate.setHours(triggerDate.getHours() + 1);
    
    ScriptApp.newTrigger('deleteTemporaryFile')
      .timeBased()
      .at(triggerDate)
      .create();
    
    // ファイルIDをスクリプトプロパティに保存（トリガー実行時に参照）
    const props = PropertiesService.getScriptProperties();
    const pendingDeletions = props.getProperty('pending_deletions') || '[]';
    const deletions = JSON.parse(pendingDeletions);
    deletions.push({ fileId: fileId, deleteAt: triggerDate.getTime() });
    props.setProperty('pending_deletions', JSON.stringify(deletions));
    
    console.log('[scheduleFileDeletion] 削除予定:', fileId, '削除時刻:', triggerDate);
  } catch (error) {
    console.warn('[scheduleFileDeletion] トリガー設定失敗:', error);
    // エラーでも処理は続行（ファイルは手動で削除可能）
  }
}

/**
 * 一時ファイルを削除（トリガーから呼ばれる）
 */
function deleteTemporaryFile() {
  try {
    const props = PropertiesService.getScriptProperties();
    const pendingDeletions = props.getProperty('pending_deletions') || '[]';
    const deletions = JSON.parse(pendingDeletions);
    const now = new Date().getTime();
    const remainingDeletions = [];

    deletions.forEach(item => {
      if (item.deleteAt <= now) {
        try {
          DriveApp.getFileById(item.fileId).setTrashed(true);
          console.log('[deleteTemporaryFile] 削除完了:', item.fileId);
        } catch (e) {
          console.warn('[deleteTemporaryFile] 削除失敗（既に削除済みの可能性）:', item.fileId, e);
        }
      } else {
        remainingDeletions.push(item);
      }
    });

    props.setProperty('pending_deletions', JSON.stringify(remainingDeletions));

    // トリガー自体を削除
    const triggers = ScriptApp.getProjectTriggers();
    triggers.forEach(trigger => {
      if (trigger.getHandlerFunction() === 'deleteTemporaryFile') {
        ScriptApp.deleteTrigger(trigger);
      }
    });

  } catch (error) {
    console.error('[deleteTemporaryFile] エラー:', error);
  }
}

// =============================================
// 既存データ修復（一度だけ実行）
// =============================================

/**
 * 商談管理シートの得意先コードに '' (ダブルクォート) が混入している
 * レコードを '(シングルクォート) + コード に修正します。
 *
 * ★ このスクリプトエディタから一度だけ手動実行してください。
 *   実行後は再度実行しても安全（冪等）です。
 */
function fixMeetingCustomerIdDoubleQuote() {
  try {
    const meetingSheet = SPREADSHEET.getSheetByName('商談管理');
    if (!meetingSheet) throw new Error('商談管理シートが見つかりません。');

    const values = meetingSheet.getDataRange().getValues();
    if (values.length <= 1) {
      console.log('[fixMeetingCustomerIdDoubleQuote] データなし。スキップ。');
      return;
    }

    const header = values[0];
    const idCol = header.indexOf('得意先コード');
    if (idCol === -1) throw new Error('商談管理シートに「得意先コード」列が見つかりません。');

    let fixedCount = 0;
    for (let i = 1; i < values.length; i++) {
      const raw = String(values[i][idCol] || '');
      // 先頭に ' が2つ以上ある場合、正規化した値を書き直す
      if (raw.startsWith("''")) {
        const cleanCode = raw.replace(/^'+/, ''); // 先頭の ' を全て除去
        const fixedValue = "'" + cleanCode;        // 正規の形式（' × 1 + コード）
        meetingSheet.getRange(i + 1, idCol + 1).setValue(fixedValue);
        console.log(`行${i + 1}: [${raw}] → [${fixedValue}]`);
        fixedCount++;
      }
    }

    // キャッシュをクリアして最新データが反映されるようにする
    clearMeetingsCache();
    clearDashboardCache();

    console.log(`[fixMeetingCustomerIdDoubleQuote] 完了: ${fixedCount}件修正しました。`);
    return `修正完了: ${fixedCount}件`;

  } catch (e) {
    console.error('[fixMeetingCustomerIdDoubleQuote] エラー:', e);
    throw e;
  }
}

/**
 * 指定シートのデータをCSV（UTF-8 BOM付き）としてBase64エンコードして返す
 * @param {string} dataType - エクスポート対象のデータ種別
 * @return {Object} { base64, fileName, rowCount }
 */
function exportDataAsCsv(dataType) {
  try {
    const sheetMap = {
      '得意先マスタ': '得意先マスタ',
      '単価マスタ': '単価マスタ',
      '商品マスタ': '商品マスタ',
      '社員マスタ': '社員マスタ',
      '商談管理': '商談管理',
      '申請管理': '申請管理'
    };

    const sheetName = sheetMap[dataType];
    if (!sheetName) throw new Error('不正なデータ種別です: ' + dataType);

    const sheet = SPREADSHEET.getSheetByName(sheetName);
    if (!sheet) throw new Error(`シート「${sheetName}」が見つかりません。`);

    const data = sheet.getDataRange().getValues();
    const rowCount = Math.max(0, data.length - 1);

    const tz = SPREADSHEET.getSpreadsheetTimeZone();

    // 各セルをCSV用に変換
    const csvRows = data.map(row =>
      row.map(cell => {
        let val;
        if (cell instanceof Date) {
          val = Utilities.formatDate(cell, tz, 'yyyy/MM/dd HH:mm:ss');
        } else {
          val = String(cell === null || cell === undefined ? '' : cell);
        }
        // ダブルクォート・カンマ・改行を含む場合はクォートで囲む
        if (val.includes('"') || val.includes(',') || val.includes('\n') || val.includes('\r')) {
          val = '"' + val.replace(/"/g, '""') + '"';
        }
        return val;
      }).join(',')
    ).join('\r\n');

    // UTF-8 BOM付きでエンコード（Excelで開いたときに文字化けしないよう）
    const csvWithBom = '\uFEFF' + csvRows;
    const blob = Utilities.newBlob(csvWithBom, MimeType.CSV, dataType + '.csv');
    const base64 = Utilities.base64Encode(blob.getBytes());

    const dateStr = Utilities.formatDate(new Date(), tz, 'yyyyMMdd_HHmmss');
    const fileName = dataType + '_' + dateStr + '.csv';

    return { base64: base64, fileName: fileName, rowCount: rowCount };

  } catch (e) {
    console.error('exportDataAsCsv error:', e);
    throw new Error('エクスポートに失敗しました: ' + e.message);
  }
}
