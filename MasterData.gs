/**
 * @fileoverview マスタデータ管理モジュール
 * マスタデータの取得、キャッシュ管理、ソート処理を提供します。
 *
 * キャッシュ戦略:
 * - CacheService で10分間キャッシュ
 * - 社員マスタ、ランクマスタ、業態マスタをキャッシュ
 */

// キャッシュ有効期限（秒）
const CACHE_DURATION = 600; // 10分

// =============================================
// キャッシュ管理
// =============================================

/**
 * キャッシュ付きでマスタデータを取得します
 * @param {string} key - キャッシュキー
 * @param {Function} fetchFunction - データ取得関数
 * @return {Array} マスタデータ
 */
function getCachedMasterData(key, fetchFunction) {
  try {
    const cache = CacheService.getScriptCache();
    const cached = cache.get(key);

    if (cached) {
      console.log(`Cache hit: ${key}`);
      return JSON.parse(cached);
    }

    console.log(`Cache miss: ${key}`);
    const data = fetchFunction();
    cache.put(key, JSON.stringify(data), CACHE_DURATION);
    return data;

  } catch (e) {
    console.error('キャッシュ取得エラー:', e);
    // キャッシュエラー時は直接取得
    return fetchFunction();
  }
}

/**
 * すべてのマスタキャッシュをクリアします
 */
function clearAllMasterCache() {
  try {
    const cache = CacheService.getScriptCache();
    cache.removeAll(['employees_master', 'ranks_master', 'business_types_master']);
    console.log('All master cache cleared');
  } catch (e) {
    console.error('キャッシュクリアエラー:', e);
  }
}

// =============================================
// 社員マスタ取得（ソート付き）
// =============================================

/**
 * 社員マスタを表示順でソートして取得します
 * @return {Array<Object>} ソート済み社員データ
 */
function getEmployeesWithSort() {
  return getCachedMasterData('employees_master', fetchEmployeesWithSort);
}

/**
 * 社員マスタを取得してソートします（内部関数）
 * @return {Array<Object>} ソート済み社員データ
 */
function fetchEmployeesWithSort() {
  try {
    const empSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('社員マスタ');
    if (!empSheet) {
      console.warn('社員マスタシートが見つかりません。');
      return [];
    }

    const values = empSheet.getDataRange().getValues();
    const header = values.shift();

    // デバッグ: ヘッダー情報を出力
    console.log('社員マスタのヘッダー:', header.join(' | '));

    // ヘッダーから列インデックスを取得
    const deptCodeCol = header.indexOf('部門コード');
    const deptNameCol = header.indexOf('部門名');
    const divisionCol = header.indexOf('大区分'); // ★追加
    const displayOrderCol = header.indexOf('表示順');
    const employeeCodeCol = header.indexOf('担当者コード');
    const employeeNameCol = header.indexOf('担当者名');
    const emailCol = header.indexOf('メールアドレス');
    const idCol = header.indexOf('ID');
    const roleCol = header.indexOf('役職');
    const retiredCol = header.indexOf('退職者');

    // デバッグ: 列インデックスを出力
    console.log(`列インデックス - 部門名: ${deptNameCol}, 担当者コード: ${employeeCodeCol}, 表示順: ${displayOrderCol}`);

    // データをオブジェクト配列に変換
    const employees = values.map(row => {
      return {
        departmentCode: deptCodeCol !== -1 ? cleanSingleQuotes(row[deptCodeCol]) : '',
        departmentName: deptNameCol !== -1 ? row[deptNameCol] : '',
        division: divisionCol !== -1 ? String(row[divisionCol]).trim() : '', // ★追加
        displayOrder: displayOrderCol !== -1 ? Number(row[displayOrderCol]) || 999 : 999,
        employeeCode: employeeCodeCol !== -1 ? cleanSingleQuotes(row[employeeCodeCol]) : '',
        employeeName: employeeNameCol !== -1 ? row[employeeNameCol] : '',
        email: emailCol !== -1 ? row[emailCol] : '',
        id: idCol !== -1 ? cleanSingleQuotes(row[idCol]) : '',
        role: roleCol !== -1 ? row[roleCol] : '',
        retired: retiredCol !== -1 ? String(row[retiredCol] || '').trim() : ''
      };
    })
    .filter(emp => emp.employeeCode) // 担当者コードが空のものは除外
    .filter(emp => emp.retired !== '退職'); // 退職者を除外

    // ソート: 表示順 → 担当者コード
    employees.sort((a, b) => {
      // 1. 表示順で比較
      if (a.displayOrder !== b.displayOrder) {
        return a.displayOrder - b.displayOrder;
      }
      // 2. 担当者コードで比較
      return String(a.employeeCode).localeCompare(String(b.employeeCode));
    });

    console.log(`社員マスタ取得: ${employees.length}件`);

    // デバッグ: 最初の3件のデータを出力
    if (employees.length > 0) {
      console.log('社員マスタサンプル（最初の3件）:');
      employees.slice(0, 3).forEach((emp, i) => {
        console.log(`  ${i + 1}. コード: [${emp.employeeCode}], 名前: ${emp.employeeName}, 部門: ${emp.departmentName}, 表示順: ${emp.displayOrder}`);
      });
    }

    return employees;

  } catch (e) {
    console.error('社員マスタ取得エラー:', e);
    return [];
  }
}

/**
 * 社員マスタから担当者コードで検索します
 * @param {string} employeeCode - 担当者コード
 * @return {Object|null} 社員情報
 */
function findEmployeeByCode(employeeCode) {
  try {
    const employees = getEmployeesWithSort();
    const cleanCode = cleanSingleQuotes(employeeCode);
    return employees.find(emp => emp.employeeCode === cleanCode) || null;
  } catch (e) {
    console.error('社員検索エラー:', e);
    return null;
  }
}

// =============================================
// ランクマスタ取得（ソート付き）
// =============================================

/**
 * ランクマスタをソートして取得します
 * @return {Array<Object>} ソート済みランクデータ
 */
function getRanksWithSort() {
  return getCachedMasterData('ranks_master', fetchRanksWithSort);
}

/**
 * ランクマスタを取得してソートします（内部関数）
 * @return {Array<Object>} ソート済みランクデータ
 */
function fetchRanksWithSort() {
  try {
    const rankSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('ランクマスタ');
    if (!rankSheet) {
      console.warn('ランクマスタシートが見つかりません。');
      return [];
    }

    const values = rankSheet.getDataRange().getValues();
    const header = values.shift();

    // データをオブジェクト配列に変換（全列データを含む）
    const ranks = values.map((row, index) => {
      const rankObj = { originalOrder: index };
      header.forEach((colName, colIndex) => {
        const cleanColName = String(colName).trim();
        if (cleanColName) {
          rankObj[cleanColName] = row[colIndex];
        }
      });
      // 後方互換性のために旧フィールドも保持
      rankObj.code = rankObj['名称コード'] ? cleanSingleQuotes(rankObj['名称コード']) : '';
      rankObj.name = rankObj['名称_1'] || rankObj['ランク名称'] || rankObj['得意先ランク区分名称'] || '';
      return rankObj;
    }).filter(rank => rank.name); // 名称が空のものは除外

    // ソート: シート上の順序（上から順番）
    ranks.sort((a, b) => a.originalOrder - b.originalOrder);

    console.log(`ランクマスタ取得: ${ranks.length}件`);
    if (ranks.length > 0) {
      console.log('ランクマスタサンプル:', JSON.stringify(ranks[0]));
    }
    return ranks;

  } catch (e) {
    console.error('ランクマスタ取得エラー:', e);
    return [];
  }
}

/**
 * ランク名から順序を取得します
 * @param {string} rankName - ランク名
 * @return {number} 順序（見つからない場合は999）
 */
function getRankOrder(rankName) {
  try {
    const ranks = getRanksWithSort();
    const rank = ranks.find(r => r.name === rankName);
    return rank ? rank.originalOrder : 999;
  } catch (e) {
    console.error('ランク順序取得エラー:', e);
    return 999;
  }
}

// =============================================
// 業態マスタ取得（ソート付き）
// =============================================

/**
 * 業態マスタをソートして取得します
 * @return {Array<Object>} ソート済み業種データ
 */
function getBusinessTypesWithSort() {
  return getCachedMasterData('business_types_master', fetchBusinessTypesWithSort);
}

/**
 * 業態マスタを取得してソートします（内部関数）
 * @return {Array<Object>} ソート済み業種データ
 */
function fetchBusinessTypesWithSort() {
  try {
    const businessSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('業態マスタ');
    if (!businessSheet) {
      console.warn('業態マスタシートが見つかりません。');
      return [];
    }

    const values = businessSheet.getDataRange().getValues();
    const header = values.shift();

    // ヘッダーから列インデックスを取得
    const codeCol = header.indexOf('業態CD');
    const nameCol = header.indexOf('業態名');

    // データをオブジェクト配列に変換
    const businessTypes = values.map(row => {
      return {
        code: codeCol !== -1 ? cleanSingleQuotes(row[codeCol]) : '',
        name: nameCol !== -1 ? row[nameCol] : '',
        originalOrder: values.indexOf(row) // シート上の順序を保持
      };
    }).filter(bt => bt.name); // 名称が空のものは除外

    // ソート: 業態CD昇順
    businessTypes.sort((a, b) => Number(a.code) - Number(b.code));

    console.log(`業態マスタ取得: ${businessTypes.length}件`);
    return businessTypes;

  } catch (e) {
    console.error('業態マスタ取得エラー:', e);
    return [];
  }
}

// =============================================
// 部門一覧取得（ソート付き）
// =============================================

/**
 * 部門一覧を表示順でソートして取得します
 * @return {Array<Object>} ソート済み部門データ
 */
function getDepartmentsWithSort() {
  try {
    const employees = getEmployeesWithSort();

    // 部門ごとに集約（重複排除）
    const deptMap = {};
    employees.forEach(emp => {
      if (emp.departmentCode && !deptMap[emp.departmentCode]) {
        deptMap[emp.departmentCode] = {
          code: emp.departmentCode,
          name: emp.departmentName,
          displayOrder: emp.displayOrder
        };
      }
    });

    // 配列に変換してソート
    const departments = Object.values(deptMap);
    departments.sort((a, b) => a.displayOrder - b.displayOrder);

    console.log(`部門一覧取得: ${departments.length}件`);
    return departments;

  } catch (e) {
    console.error('部門一覧取得エラー:', e);
    return [];
  }
}

// =============================================
// 顧客データのソート処理
// =============================================

/**
 * 顧客データをソートします
 * ソート順: 部門表示順 → 担当者コード → ランク順 → 業態順
 * @param {Array<Object>} customers - 顧客データ配列
 * @return {Array<Object>} ソート済み顧客データ
 */
function sortCustomers(customers) {
  try {
    console.log(`顧客ソート開始: ${customers.length}件`);
    const startTime = new Date().getTime();

    // マスタデータを事前に取得（キャッシュされる）
    const employees = getEmployeesWithSort();
    const ranks = getRanksWithSort();
    const businessTypes = getBusinessTypesWithSort();

    // 社員コードでの検索を高速化するためのマップを作成
    const employeeMap = {};
    employees.forEach(emp => {
      employeeMap[emp.employeeCode] = emp;
    });

    // ランク名での検索を高速化するためのマップを作成
    const rankMap = {};
    ranks.forEach((rank, index) => {
      rankMap[rank.name] = index;
    });

    // 業態（業種）名での検索を高速化するためのマップを作成
    const businessTypeMap = {};
    businessTypes.forEach((bt, index) => {
      businessTypeMap[bt.name] = index;
    });

    // 各顧客にソートキーを追加
    const customersWithKeys = customers.map(customer => {
      const employeeCode = cleanSingleQuotes(customer['営業担当者コード'] || '');
      const employee = employeeMap[employeeCode];

      return {
        ...customer,
        _sortKey1: employee ? employee.displayOrder : 999,
        _sortKey2: employee ? employee.employeeCode : 'zzz',
        _sortKey3: rankMap[customer['得意先ランク区分名称']] !== undefined
          ? rankMap[customer['得意先ランク区分名称']]
          : 999,
        _sortKey4: businessTypeMap[customer['businessType']] !== undefined
          ? businessTypeMap[customer['businessType']]
          : 999
      };
    });

    // ソート実行
    customersWithKeys.sort((a, b) => {
      // 1. 部門の表示順で比較
      if (a._sortKey1 !== b._sortKey1) {
        return a._sortKey1 - b._sortKey1;
      }
      // 2. 担当者コードで比較
      if (a._sortKey2 !== b._sortKey2) {
        return String(a._sortKey2).localeCompare(String(b._sortKey2));
      }
      // 3. ランク順で比較
      if (a._sortKey3 !== b._sortKey3) {
        return a._sortKey3 - b._sortKey3;
      }
      // 4. 業態順で比較
      return a._sortKey4 - b._sortKey4;
    });

    // ソートキーを削除
    const sortedCustomers = customersWithKeys.map(customer => {
      const { _sortKey1, _sortKey2, _sortKey3, _sortKey4, ...cleanCustomer } = customer;
      return cleanCustomer;
    });

    const endTime = new Date().getTime();
    console.log(`顧客ソート完了: ${endTime - startTime}ms`);

    return sortedCustomers;

  } catch (e) {
    console.error('顧客ソートエラー:', e);
    // エラー時は元のデータを返す
    return customers;
  }
}

// =============================================
// フィルタ用ソート済みデータ取得
// =============================================

/**
 * フィルタ用の部門リストを取得します
 * @return {Array<Object>} ソート済み部門リスト
 */
function getFilterDepartments() {
  try {
    const departments = getDepartmentsWithSort();
    return departments.map(dept => ({
      code: dept.code,
      name: dept.name
    }));
  } catch (e) {
    console.error('フィルタ部門取得エラー:', e);
    return [];
  }
}

/**
 * フィルタ用の大区分リストを取得します
 * @return {Array<Object>} 大区分リスト
 */
function getFilterDivisions() {
  try {
    const employees = getEmployeesWithSort();
    const divisions = [...new Set(employees.map(emp => emp.division).filter(Boolean))];
    return divisions.map(div => ({ name: div }));
  } catch (e) {
    console.error('フィルタ大区分取得エラー:', e);
    return [];
  }
}

/**
 * フィルタ用の担当者リストを取得します
 * @param {string} departmentCode - 部門コード（オプション）
 * @param {string} division - 大区分（オプション）
 * @return {Array<Object>} ソート済み担当者リスト
 */
function getFilterEmployees(departmentCode, division) {
  try {
    let employees = getEmployeesWithSort();

    // 大区分が指定されている場合は絞り込み
    if (division) {
      employees = employees.filter(emp => emp.division === division);
    }

    // 部門コードが指定されている場合は絞り込み
    if (departmentCode) {
      const cleanCode = cleanSingleQuotes(departmentCode);
      employees = employees.filter(emp => emp.departmentCode === cleanCode);
    }

    return employees.map(emp => ({
      code: emp.employeeCode,
      name: emp.employeeName,
      department: emp.departmentName,
      division: emp.division
    }));
  } catch (e) {
    console.error('フィルタ担当者取得エラー:', e);
    return [];
  }
}

/**
 * フィルタ用のランクリストを取得します
 * @return {Array<Object>} ソート済みランクリスト
 */
function getFilterRanks() {
  try {
    const ranks = getRanksWithSort();
    return ranks.map(rank => ({
      code: rank.code,
      name: rank.name,
      visibility: String(rank['表示有無'] || '').trim()
    }));
  } catch (e) {
    console.error('フィルタランク取得エラー:', e);
    return [];
  }
}

/**
 * フィルタ用の業種リストを取得します
 * @return {Array<Object>} ソート済み業種リスト
 */
function getFilterBusinessTypes() {
  try {
    const businessTypes = getBusinessTypesWithSort();
    return businessTypes.map(bt => ({
      code: bt.code,
      name: bt.name
    }));
  } catch (e) {
    console.error('フィルタ業種取得エラー:', e);
    return [];
  }
}

// =============================================
// 締日マスタ取得
// =============================================

/**
 * 締日マスタを取得します
 * @return {Array<Object>} 締日パターンデータ
 */
function getPaymentTerms() {
  return getCachedMasterData('payment_terms_master', fetchPaymentTerms);
}

/**
 * 締日マスタを取得します（内部関数）
 * @return {Array<Object>} 締日パターンデータ
 */
function fetchPaymentTerms() {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('締日マスタ');
    if (!sheet) {
      console.warn('締日マスタシートが見つかりません。');
      return [];
    }

    const values = sheet.getDataRange().getValues();
    const header = values.shift();

    // ヘッダーから列インデックスを取得
    const patternCol = header.indexOf('締日パターン');

    // データをオブジェクト配列に変換
    const terms = values.map((row, index) => {
      return {
        pattern: patternCol !== -1 ? row[patternCol] : '',
        originalOrder: index
      };
    }).filter(t => t.pattern); // 締日パターンが空のものは除外

    // ソート: シート上の順序
    terms.sort((a, b) => a.originalOrder - b.originalOrder);

    console.log(`締日マスタ取得: ${terms.length}件`);
    return terms;

  } catch (e) {
    console.error('締日マスタ取得エラー:', e);
    return [];
  }
}
