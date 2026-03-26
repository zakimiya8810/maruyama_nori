/**
 * @fileoverview ワークフロー管理モジュール
 * 5段階承認フロー、バックアップ・ロールバック、メール通知を管理します。
 *
 * 承認フロー:
 * 申請中 → 上長承認済 → 管理承認済 → 横山承認済 → 決裁完了
 *
 * 重要な処理:
 * - 新規申請: 社長承認（決裁完了）後にマスタ反映
 * - 修正申請: 横山承認時にマスタ先行反映 + バックアップ保存
 *            社長却下時にバックアップから復元
 */

// =============================================
// ユーティリティ関数
// =============================================

/**
 * Web AppのURLを取得します（申請詳細ページへの直リンク用）
 * @param {string} applicationId - 申請ID
 * @return {string} 申請詳細ページへの直リンクURL
 */
function getApplicationDetailUrl(applicationId) {
  try {
    const webAppUrl = ScriptApp.getService().getUrl();
    return `${webAppUrl}?appId=${applicationId}`;
  } catch (e) {
    console.error('Web AppのURL取得エラー:', e);
    return '';
  }
}

// =============================================
// 承認ルート判定
// =============================================

/**
 * 次の承認者を特定します
 * @param {string} applicationId - 申請ID
 * @param {string} currentStage - 現在の承認段階
 * @return {Object} 次の承認者情報 {id, name, email, role}
 */
function getNextApprover(applicationId, currentStage) {
  try {
    const appSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('申請管理');
    if (!appSheet) throw new Error('申請管理シートが見つかりません。');

    const values = appSheet.getDataRange().getValues();
    const header = values.shift();

    // ヘッダーから列インデックスを取得
    const idCol = header.indexOf('申請ID');
    const applicantIdCol = header.indexOf('申請者ID');
    const applicantNameCol = header.indexOf('申請者名');

    if (idCol === -1) throw new Error('「申請ID」列が見つかりません。');

    // 申請データを取得
    const rowIndex = values.findIndex(r => String(r[idCol]) === String(applicationId));
    if (rowIndex === -1) throw new Error('指定された申請が見つかりません。');

    const targetRow = values[rowIndex];
    const applicantId = applicantIdCol !== -1 ? cleanSingleQuotes(targetRow[applicantIdCol]) : null;
    const applicantName = applicantNameCol !== -1 ? targetRow[applicantNameCol] : null;

    // 承認段階に応じて次の承認者を決定
    switch (currentStage) {
      case '申請中':
        // 上長を特定
        if (applicantId) {
          return findSupervisor(applicantId);
        } else if (applicantName) {
          // 申請者IDがない場合は名前から探す
          const applicant = findEmployeeByName(applicantName);
          if (applicant && applicant.id) {
            return findSupervisor(applicant.id);
          }
        }
        throw new Error('申請者情報が不足しており、上長を特定できません。');

      case '上長承認済':
        // 管理部門の承認者を特定（役職が「管理部門」の社員）
        return findApproverByRole('管理部門');

      case '管理承認済':
        // 常務を特定（役職が「常務」の社員）
        return findApproverByRole('常務');

      case '常務承認済':
        // 決裁者を特定（役職が「決裁者」の社員）
        return findApproverByRole('決裁者');

      default:
        throw new Error(`不明な承認段階: ${currentStage}`);
    }

  } catch (e) {
    console.error('次の承認者特定エラー:', e);
    throw e;
  }
}

/**
 * 申請者の上長を特定します（複数いる場合は最初の1人を返す）
 * 同じ大区分で役職が「上長」の社員を検索
 * @param {string} employeeId - 社員ID
 * @return {Object} 上長情報 {id, name, email, role, department}
 */
function findSupervisor(employeeId) {
  try {
    const empSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('社員マスタ');
    if (!empSheet) throw new Error('社員マスタシートが見つかりません。');

    const values = empSheet.getDataRange().getValues();
    const header = values.shift();

    // ヘッダーから列インデックスを取得
    const idCol = header.indexOf('ID');
    const nameCol = header.indexOf('担当者名');
    const emailCol = header.indexOf('メールアドレス');
    const roleCol = header.indexOf('役職');
    const divisionCol = header.indexOf('大区分');
    const deptNameCol = header.indexOf('部門名');

    if (idCol === -1 || nameCol === -1 || emailCol === -1) {
      throw new Error('社員マスタに必要な列が見つかりません。');
    }

    // 申請者の情報を取得
    const applicantRow = values.find(r => cleanSingleQuotes(r[idCol]) === cleanSingleQuotes(employeeId));
    if (!applicantRow) throw new Error(`社員ID ${employeeId} が見つかりません。`);

    const applicantDivision = divisionCol !== -1 ? String(applicantRow[divisionCol]).trim() : null;

    // 同じ大区分で役職が「上長」の社員を検索（複数いる場合は最初の1人）
    const supervisor = values.find(r => {
      const rowDivision = divisionCol !== -1 ? String(r[divisionCol]).trim() : null;
      const rowRole = roleCol !== -1 ? String(r[roleCol]).trim() : '';
      return rowDivision === applicantDivision && rowRole === '上長';
    });

    if (!supervisor) {
      throw new Error(`大区分「${applicantDivision}」の上長が見つかりません。`);
    }

    return {
      id: cleanSingleQuotes(supervisor[idCol]),
      name: supervisor[nameCol],
      email: supervisor[emailCol],
      role: roleCol !== -1 ? supervisor[roleCol] : '',
      department: deptNameCol !== -1 ? supervisor[deptNameCol] : ''
    };

  } catch (e) {
    console.error('上長特定エラー:', e);
    throw e;
  }
}

/**
 * 役職で承認者を特定します
 * @param {string} roleKeyword - 役職のキーワード（例: '管理'）
 * @return {Object} 承認者情報 {id, name, email, role, department}
 */
function findApproverByRole(roleKeyword) {
  try {
    const empSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('社員マスタ');
    if (!empSheet) throw new Error('社員マスタシートが見つかりません。');

    const values = empSheet.getDataRange().getValues();
    const header = values.shift();

    // ヘッダーから列インデックスを取得
    const idCol = header.indexOf('ID');
    const nameCol = header.indexOf('担当者名');
    const emailCol = header.indexOf('メールアドレス');
    const roleCol = header.indexOf('役職');
    const deptNameCol = header.indexOf('部門名');

    if (idCol === -1 || nameCol === -1 || emailCol === -1 || roleCol === -1) {
      throw new Error('社員マスタに必要な列が見つかりません。');
    }

    // 役職にキーワードが含まれる社員を検索
    const approver = values.find(r => {
      const rowRole = String(r[roleCol]);
      return rowRole.includes(roleKeyword);
    });

    if (!approver) {
      throw new Error(`役職「${roleKeyword}」の承認者が見つかりません。`);
    }

    return {
      id: cleanSingleQuotes(approver[idCol]),
      name: approver[nameCol],
      email: approver[emailCol],
      role: approver[roleCol],
      department: deptNameCol !== -1 ? approver[deptNameCol] : ''
    };

  } catch (e) {
    console.error('役職による承認者特定エラー:', e);
    throw e;
  }
}

/**
 * 名前で承認者を特定します
 * @param {string} nameKeyword - 名前のキーワード（例: '横山', '丸山'）
 * @return {Object} 承認者情報 {id, name, email, role, department}
 */
function findApproverByName(nameKeyword) {
  try {
    const empSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('社員マスタ');
    if (!empSheet) throw new Error('社員マスタシートが見つかりません。');

    const values = empSheet.getDataRange().getValues();
    const header = values.shift();

    // ヘッダーから列インデックスを取得
    const idCol = header.indexOf('ID');
    const nameCol = header.indexOf('担当者名');
    const emailCol = header.indexOf('メールアドレス');
    const roleCol = header.indexOf('役職');
    const deptNameCol = header.indexOf('部門名');

    if (idCol === -1 || nameCol === -1 || emailCol === -1) {
      throw new Error('社員マスタに必要な列が見つかりません。');
    }

    // 名前にキーワードが含まれる社員を検索
    const approver = values.find(r => {
      const rowName = String(r[nameCol]);
      return rowName.includes(nameKeyword);
    });

    if (!approver) {
      throw new Error(`名前「${nameKeyword}」の承認者が見つかりません。`);
    }

    return {
      id: cleanSingleQuotes(approver[idCol]),
      name: approver[nameCol],
      email: approver[emailCol],
      role: roleCol !== -1 ? approver[roleCol] : '',
      department: deptNameCol !== -1 ? approver[deptNameCol] : ''
    };

  } catch (e) {
    console.error('名前による承認者特定エラー:', e);
    throw e;
  }
}

/**
 * 申請者の上長を全員取得します（複数上長対応・通知用）
 * 同じ大区分で役職が「上長」の社員を全て返す
 * @param {string} employeeId - 社員ID
 * @return {Array<Object>} 上長情報の配列
 */
function findAllSupervisors(employeeId) {
  try {
    const empSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('社員マスタ');
    if (!empSheet) throw new Error('社員マスタシートが見つかりません。');

    const values = empSheet.getDataRange().getValues();
    const header = values.shift();

    // ヘッダーから列インデックスを取得
    const idCol = header.indexOf('ID');
    const nameCol = header.indexOf('担当者名');
    const emailCol = header.indexOf('メールアドレス');
    const roleCol = header.indexOf('役職');
    const divisionCol = header.indexOf('大区分');
    const deptNameCol = header.indexOf('部門名');

    if (idCol === -1 || nameCol === -1 || emailCol === -1) {
      throw new Error('社員マスタに必要な列が見つかりません。');
    }

    // 申請者の情報を取得
    const applicantRow = values.find(r => cleanSingleQuotes(r[idCol]) === cleanSingleQuotes(employeeId));
    if (!applicantRow) throw new Error(`社員ID ${employeeId} が見つかりません。`);

    const applicantDivision = divisionCol !== -1 ? String(applicantRow[divisionCol]).trim() : null;

    // 同じ大区分で役職が「上長」の社員を全て検索
    const supervisors = values.filter(r => {
      const rowDivision = divisionCol !== -1 ? String(r[divisionCol]).trim() : null;
      const rowRole = roleCol !== -1 ? String(r[roleCol]).trim() : '';
      return rowDivision === applicantDivision && rowRole === '上長';
    }).map(supervisor => ({
      id: cleanSingleQuotes(supervisor[idCol]),
      name: supervisor[nameCol],
      email: supervisor[emailCol],
      role: roleCol !== -1 ? supervisor[roleCol] : '',
      department: deptNameCol !== -1 ? supervisor[deptNameCol] : ''
    }));

    if (supervisors.length === 0) {
      throw new Error(`大区分「${applicantDivision}」の上長が見つかりません。`);
    }

    console.log(`上長を${supervisors.length}人発見: ${supervisors.map(s => s.name).join(', ')}`);
    return supervisors;

  } catch (e) {
    console.error('上長全員取得エラー:', e);
    throw e;
  }
}

/**
 * 名前で社員を検索します
 * @param {string} employeeName - 社員名
 * @return {Object|null} 社員情報 {id, name, email, role, department}
 */
function findEmployeeByName(employeeName) {
  try {
    const empSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('社員マスタ');
    if (!empSheet) return null;

    const values = empSheet.getDataRange().getValues();
    const header = values.shift();

    const idCol = header.indexOf('ID');
    const nameCol = header.indexOf('担当者名');
    const emailCol = header.indexOf('メールアドレス');
    const roleCol = header.indexOf('役職');
    const deptNameCol = header.indexOf('部門名');

    if (nameCol === -1) return null;

    const employee = values.find(r => String(r[nameCol]) === String(employeeName));
    if (!employee) return null;

    return {
      id: idCol !== -1 ? cleanSingleQuotes(employee[idCol]) : null,
      name: employee[nameCol],
      email: emailCol !== -1 ? employee[emailCol] : '',
      role: roleCol !== -1 ? employee[roleCol] : '',
      department: deptNameCol !== -1 ? employee[deptNameCol] : ''
    };

  } catch (e) {
    console.error('社員検索エラー:', e);
    return null;
  }
}

// =============================================
// バックアップ・ロールバック機能
// =============================================

/**
 * 横山承認時に顧客データのバックアップを保存します
 * @param {string} applicationId - 申請ID
 * @param {string} customerId - 顧客コード
 * @return {Object} 結果 {success: boolean, message: string}
 */
function saveSnapshot(applicationId, customerId) {
  try {
    const customerSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('得意先マスタ');
    if (!customerSheet) throw new Error('得意先マスタシートが見つかりません。');

    const values = customerSheet.getDataRange().getValues();
    const header = values.shift();

    // 得意先コードの列を取得
    const customerIdCol = header.indexOf('得意先コード');
    if (customerIdCol === -1) throw new Error('「得意先コード」列が見つかりません。');

    // 顧客データを検索
    const customerRow = values.find(r => cleanSingleQuotes(r[customerIdCol]) === cleanSingleQuotes(customerId));
    if (!customerRow) throw new Error(`顧客コード ${customerId} が見つかりません。`);

    // データをオブジェクトに変換
    const snapshot = {};
    header.forEach((col, idx) => {
      snapshot[col] = customerRow[idx];
    });

    // JSON文字列に変換
    const snapshotJson = JSON.stringify(snapshot);

    // 申請管理シートにバックアップを保存
    const appSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('申請管理');
    if (!appSheet) throw new Error('申請管理シートが見つかりません。');

    const appValues = appSheet.getDataRange().getValues();
    const appHeader = appValues.shift();

    const appIdCol = appHeader.indexOf('申請ID');
    const backupCol = appHeader.indexOf('バックアップデータ');

    if (appIdCol === -1) throw new Error('「申請ID」列が見つかりません。');
    if (backupCol === -1) throw new Error('「バックアップデータ」列が見つかりません。');

    // 申請行を検索
    const appRowIndex = appValues.findIndex(r => String(r[appIdCol]) === String(applicationId));
    if (appRowIndex === -1) throw new Error('指定された申請が見つかりません。');

    // バックアップを保存
    appSheet.getRange(appRowIndex + 2, backupCol + 1).setValue(snapshotJson);

    return {
      success: true,
      message: 'バックアップを保存しました。'
    };

  } catch (e) {
    console.error('バックアップ保存エラー:', e);
    return {
      success: false,
      message: `バックアップ保存に失敗しました: ${e.message}`
    };
  }
}

/**
 * 社長却下時にバックアップからマスタを復元します
 * @param {string} applicationId - 申請ID
 * @return {Object} 結果 {success: boolean, message: string}
 */
function rollbackData(applicationId) {
  try {
    // 申請管理シートからバックアップを取得
    const appSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('申請管理');
    if (!appSheet) throw new Error('申請管理シートが見つかりません。');

    const appValues = appSheet.getDataRange().getValues();
    const appHeader = appValues.shift();

    const appIdCol = appHeader.indexOf('申請ID');
    const backupCol = appHeader.indexOf('バックアップデータ');
    const customerIdCol = appHeader.indexOf('得意先コード');

    // ヘッダー名の互換性対応（旧名・新名の両方をサポート）
    if (appIdCol === -1 || backupCol === -1 || customerIdCol === -1) {
      throw new Error('申請管理シートに必要な列が見つかりません。');
    }

    // 申請行を検索
    const appRowIndex = appValues.findIndex(r => String(r[appIdCol]) === String(applicationId));
    if (appRowIndex === -1) throw new Error('指定された申請が見つかりません。');

    const targetRow = appValues[appRowIndex];
    const backupJson = targetRow[backupCol];
    const customerId = cleanSingleQuotes(targetRow[customerIdCol]);

    if (!backupJson) throw new Error('バックアップデータが見つかりません。');

    // JSONをパース
    const snapshot = JSON.parse(backupJson);

    // 得意先マスタを復元
    const customerSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('得意先マスタ');
    if (!customerSheet) throw new Error('得意先マスタシートが見つかりません。');

    const custValues = customerSheet.getDataRange().getValues();
    const custHeader = custValues.shift();

    const custIdCol = custHeader.indexOf('得意先コード');
    if (custIdCol === -1) throw new Error('「得意先コード」列が見つかりません。');

    // 顧客行を検索
    const custRowIndex = custValues.findIndex(r => cleanSingleQuotes(r[custIdCol]) === cleanSingleQuotes(customerId));
    if (custRowIndex === -1) throw new Error(`顧客コード ${customerId} が見つかりません。`);

    // バックアップデータで上書き
    custHeader.forEach((col, idx) => {
      if (snapshot.hasOwnProperty(col)) {
        customerSheet.getRange(custRowIndex + 2, idx + 1).setValue(snapshot[col]);
      }
    });

    return {
      success: true,
      message: 'マスタをロールバックしました。'
    };

  } catch (e) {
    console.error('ロールバックエラー:', e);
    return {
      success: false,
      message: `ロールバックに失敗しました: ${e.message}`
    };
  }
}

// =============================================
// 単価マスタのバックアップ・ロールバック機能
// =============================================

/**
 * 常務承認時に単価マスタのバックアップを保存します
 * @param {string} applicationId - 申請ID
 * @param {string} customerId - 顧客コード
 * @return {Object} 結果 {success: boolean, message: string}
 */
function savePriceSnapshot(applicationId, customerId) {
  try {
    const tankaSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('単価マスタ');
    if (!tankaSheet) throw new Error('単価マスタシートが見つかりません。');

    const values = tankaSheet.getDataRange().getValues();
    const header = values.shift();

    // 得意先コードの列を取得
    const customerIdCol = header.indexOf('得意先コード');
    if (customerIdCol === -1) throw new Error('「得意先コード」列が見つかりません。');

    // 対象顧客の全商品行を取得
    const cleanedCustomerId = cleanSingleQuotes(customerId);
    const priceRows = values.filter(r => cleanSingleQuotes(r[customerIdCol]) === cleanedCustomerId);

    if (priceRows.length === 0) {
      console.log(`顧客コード ${customerId} の単価データが見つかりません（新規顧客の可能性）。空の配列で保存します。`);
    }

    // 各行をオブジェクトに変換して配列に格納
    const snapshot = priceRows.map(row => {
      const rowData = {};
      header.forEach((col, idx) => {
        rowData[col] = row[idx];
      });
      return rowData;
    });

    // JSON文字列に変換
    const snapshotJson = JSON.stringify(snapshot);

    // 申請管理シートにバックアップを保存
    const appSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('申請管理');
    if (!appSheet) throw new Error('申請管理シートが見つかりません。');

    const appValues = appSheet.getDataRange().getValues();
    const appHeader = appValues.shift();

    const appIdCol = appHeader.indexOf('申請ID');
    let backupCol = appHeader.indexOf('単価バックアップデータ');

    if (appIdCol === -1) throw new Error('「申請ID」列が見つかりません。');

    // 「単価バックアップデータ」列がなければ作成
    if (backupCol === -1) {
      const lastCol = appHeader.length;
      appSheet.getRange(1, lastCol + 1).setValue('単価バックアップデータ');
      backupCol = lastCol;
      console.log('「単価バックアップデータ」列を新規作成しました。');
    }

    // 申請行を検索
    const appRowIndex = appValues.findIndex(r => String(r[appIdCol]) === String(applicationId));
    if (appRowIndex === -1) throw new Error('指定された申請が見つかりません。');

    // バックアップを保存
    appSheet.getRange(appRowIndex + 2, backupCol + 1).setValue(snapshotJson);

    console.log(`単価バックアップ保存完了: 顧客=${customerId}, 商品数=${priceRows.length}`);

    return {
      success: true,
      message: `単価バックアップを保存しました（${priceRows.length}商品）。`
    };

  } catch (e) {
    console.error('単価バックアップ保存エラー:', e);
    return {
      success: false,
      message: `単価バックアップ保存に失敗しました: ${e.message}`
    };
  }
}

/**
 * 決裁者却下時にバックアップから単価マスタを復元します
 * @param {string} applicationId - 申請ID
 * @return {Object} 結果 {success: boolean, message: string}
 */
function rollbackPriceData(applicationId) {
  try {
    // 申請管理シートからバックアップを取得
    const appSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('申請管理');
    if (!appSheet) throw new Error('申請管理シートが見つかりません。');

    const appValues = appSheet.getDataRange().getValues();
    const appHeader = appValues.shift();

    const appIdCol = appHeader.indexOf('申請ID');
    const backupCol = appHeader.indexOf('単価バックアップデータ');
    const customerIdCol = appHeader.indexOf('得意先コード');

    if (appIdCol === -1 || backupCol === -1 || customerIdCol === -1) {
      throw new Error('申請管理シートに必要な列が見つかりません。');
    }

    // 申請行を検索
    const appRowIndex = appValues.findIndex(r => String(r[appIdCol]) === String(applicationId));
    if (appRowIndex === -1) throw new Error('指定された申請が見つかりません。');

    const targetRow = appValues[appRowIndex];
    const backupJson = targetRow[backupCol];
    const customerId = cleanSingleQuotes(targetRow[customerIdCol]);

    if (!backupJson) {
      console.warn('単価バックアップデータが見つかりません。新規顧客または単価データがない申請の可能性があります。');
      return {
        success: true,
        message: '単価バックアップデータがありません（処理スキップ）。'
      };
    }

    // JSONをパース
    const snapshot = JSON.parse(backupJson);

    // 単価マスタを復元
    const tankaSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('単価マスタ');
    if (!tankaSheet) throw new Error('単価マスタシートが見つかりません。');

    const tankaValues = tankaSheet.getDataRange().getValues();
    const tankaHeader = tankaValues.shift();

    const customerIdCol_tanka = tankaHeader.indexOf('得意先コード');
    const productCodeCol = tankaHeader.indexOf('商品コード');

    if (customerIdCol_tanka === -1) throw new Error('「得意先コード」列が見つかりません。');
    if (productCodeCol === -1) throw new Error('「商品コード」列が見つかりません。');

    const cleanedCustomerId = cleanSingleQuotes(customerId);

    // 対象顧客の現在の単価データ行を全て削除
    const rowsToDelete = [];
    tankaValues.forEach((row, idx) => {
      if (cleanSingleQuotes(row[customerIdCol_tanka]) === cleanedCustomerId) {
        rowsToDelete.push(idx + 2); // +2 はヘッダー行とインデックスの調整
      }
    });

    // 行を削除（後ろから削除しないとインデックスがずれる）
    rowsToDelete.reverse().forEach(rowNum => {
      tankaSheet.deleteRow(rowNum);
    });

    console.log(`単価マスタから${rowsToDelete.length}行を削除しました（顧客=${customerId}）`);

    // バックアップデータを再挿入
    if (snapshot.length > 0) {
      const newRows = snapshot.map(rowData => {
        return tankaHeader.map(col => rowData[col] !== undefined ? rowData[col] : '');
      });

      // 最終行の次に追加
      const lastRow = tankaSheet.getLastRow();
      tankaSheet.getRange(lastRow + 1, 1, newRows.length, tankaHeader.length).setValues(newRows);

      console.log(`単価マスタに${newRows.length}行を復元しました（顧客=${customerId}）`);
    }

    return {
      success: true,
      message: `単価マスタをロールバックしました（${snapshot.length}商品）。`
    };

  } catch (e) {
    console.error('単価ロールバックエラー:', e);
    return {
      success: false,
      message: `単価ロールバックに失敗しました: ${e.message}`
    };
  }
}

// =============================================
// 承認処理
// =============================================

/**
 * 承認処理を実行します
 * @param {string} applicationId - 申請ID
 * @param {string} approverId - 承認者ID
 * @return {Object} 結果 {success: boolean, message: string, nextStage: string}
 */
function approveApplication(applicationId, approverId) {
  try {
    console.log('[approveApplication] 開始 - applicationId:', applicationId, 'approverId:', approverId);

    const appSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('申請管理');
    if (!appSheet) throw new Error('申請管理シートが見つかりません。');

    const values = appSheet.getDataRange().getValues();
    const header = values.shift();

    console.log('[approveApplication] ヘッダー:', header);

    // ヘッダーから列インデックスを取得
    const idCol = header.indexOf('申請ID');
    const stageCol = header.indexOf('承認段階');
    const statusCol = header.indexOf('ステータス');
    const typeCol = header.indexOf('申請種別');
    const customerIdCol = header.indexOf('得意先コード');
    const targetMasterCol = header.indexOf('対象マスタ');

    console.log('[approveApplication] 列インデックス - idCol:', idCol, 'stageCol:', stageCol, 'statusCol:', statusCol);

    if (idCol === -1 || stageCol === -1) {
      throw new Error('申請管理シートに必要な列が見つかりません。');
    }

    // 申請行を検索
    const rowIndex = values.findIndex(r => String(r[idCol]) === String(applicationId));
    if (rowIndex === -1) throw new Error('指定された申請が見つかりません。');

    const targetRow = values[rowIndex];
    const currentStage = targetRow[stageCol];
    const applicationType = typeCol !== -1 ? targetRow[typeCol] : '';
    const customerId = customerIdCol !== -1 ? cleanSingleQuotes(targetRow[customerIdCol]) : '';
    const targetMaster = targetMasterCol !== -1 ? targetRow[targetMasterCol] : '';

    console.log('[approveApplication] 現在の承認段階:', currentStage, '申請種別:', applicationType, '対象マスタ:', targetMaster);

    // 次の段階を決定
    let nextStage = '';
    let nextStatus = '';

    switch (currentStage) {
      case '申請中':
        nextStage = '上長承認済';
        nextStatus = '承認中';
        break;
      case '上長承認済':
        nextStage = '管理承認済';
        nextStatus = '承認中';
        break;
      case '管理承認済':
        nextStage = '常務承認済';
        nextStatus = '承認中';

        // 常務承認時：修正申請の場合はバックアップ保存→マスタを先行反映
        if (applicationType === '顧客情報修正') {
          // ★ 顧客マスタのバックアップ保存（更新前の最終更新日を含む）
          const backupResult = saveSnapshot(applicationId, customerId);
          if (!backupResult.success) {
            console.warn('顧客バックアップ保存に失敗しました:', backupResult.message);
          }

          // ★ 顧客マスタ反映処理（最終更新日を更新）
          const updateResult = updateCustomerMaster(applicationId);
          if (!updateResult.success) {
            throw new Error(`マスタ反映に失敗: ${updateResult.message}`);
          }
        }

        // ★ 追加: 単価マスタ修正申請の場合もバックアップ→先行反映
        if (targetMaster && targetMaster.includes('単価マスタ')) {
          console.log('[approveApplication] 単価マスタ修正のため、バックアップ→先行反映を実行');

          // 単価マスタのバックアップ保存
          const priceBackupResult = savePriceSnapshot(applicationId, customerId);
          if (!priceBackupResult.success) {
            console.warn('単価バックアップ保存に失敗しました:', priceBackupResult.message);
          }

          // 単価マスタ反映処理
          const updateResult = updateCustomerMaster(applicationId);
          if (!updateResult.success) {
            throw new Error(`単価マスタ反映に失敗: ${updateResult.message}`);
          }
        }
        break;
      case '常務承認済':
        nextStage = '決裁完了';
        nextStatus = '承認済';

        // 決裁者承認時：新規申請の場合はマスタを反映
        if (applicationType === '顧客新規登録') {
          const updateResult = updateCustomerMaster(applicationId);
          if (!updateResult.success) {
            throw new Error(`マスタ反映に失敗: ${updateResult.message}`);
          }

          // ★ RPA連携：顧客マスタが更新された場合、RPAシートに書き込み
          const approvalDate = new Date();
          const rpaResult = writeCustomerToRPA(applicationId, customerId, approvalDate);
          if (!rpaResult.success) {
            console.warn('[approveApplication] RPA連携に失敗しましたが、処理は続行します:', rpaResult.message);
          }
        }

        // ★ RPA連携：顧客情報修正の場合もRPAシートに書き込み
        if (applicationType === '顧客情報修正') {
          const approvalDate = new Date();
          const rpaResult = writeCustomerToRPA(applicationId, customerId, approvalDate);
          if (!rpaResult.success) {
            console.warn('[approveApplication] RPA連携に失敗しましたが、処理は続行します:', rpaResult.message);
          }
        }

        // ★ RPA連携：単価マスタが含まれる場合、単価情報もRPAシートに書き込み
        if (targetMaster && targetMaster.includes('単価マスタ')) {
          const approvalDate = new Date();
          const rpaPriceResult = writePriceToRPA(applicationId, customerId, approvalDate);
          if (!rpaPriceResult.success) {
            console.warn('[approveApplication] 単価RPA連携に失敗しましたが、処理は続行します:', rpaPriceResult.message);
          }
        }
        break;
      default:
        throw new Error(`不明な承認段階: ${currentStage}`);
    }

    // 承認段階とステータスを更新
    console.log('[approveApplication] 承認段階を更新:', nextStage, 'ステータス:', nextStatus);
    appSheet.getRange(rowIndex + 2, stageCol + 1).setValue(nextStage);
    if (statusCol !== -1) {
      appSheet.getRange(rowIndex + 2, statusCol + 1).setValue(nextStatus);
    }

    // 判断者を記録
    console.log('[approveApplication] updateApprovalStage呼び出し');
    updateApprovalStage(applicationId, nextStage, approverId);

    // 承認通知は無効化中
    // if (nextStage !== '決裁完了') {
    //   notifyNextApprover(applicationId);
    // }

    // 決裁完了時に申請者へ完了通知
    if (nextStage === '決裁完了') {
      notifyApplicantOnCompletion(applicationId);
    }

    console.log('[approveApplication] 承認処理完了');
    return {
      success: true,
      message: `承認しました。`,
      nextStage: nextStage
    };

  } catch (e) {
    console.error('承認処理エラー:', e);
    return {
      success: false,
      message: `承認処理に失敗しました: ${e.message}`
    };
  }
}

/**
 * 却下/却下処理を実行します
 * @param {string} applicationId - 申請ID
 * @param {string} reason - 理由
 * @param {string} notifyTo - 通知先メールアドレス
 * @param {boolean} isReject - true: 却下, false: 却下
 * @param {Array} requiredFields - 修正指示項目
 * @param {string} rejectorId - 却下者ID
 * @return {Object} 結果 {success: boolean, message: string}
 */
function rejectApplication(applicationId, reason, notifyTo, isReject = false, requiredFields = [], rejectorId = '') {
  try {
    const appSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('申請管理');
    if (!appSheet) throw new Error('申請管理シートが見つかりません。');

    const values = appSheet.getDataRange().getValues();
    const header = values.shift();

    const idCol = header.indexOf('申請ID');
    const stageCol = header.indexOf('承認段階');
    const statusCol = header.indexOf('ステータス');
    const rejectReasonCol = header.indexOf('却下理由');
    const requiredFieldsCol = header.indexOf('修正指示項目'); // ★ 追加
    const targetMasterCol = header.indexOf('対象マスタ');

    if (idCol === -1) throw new Error('「申請ID」列が見つかりません。');

    // 申請行を検索
    const rowIndex = values.findIndex(r => String(r[idCol]) === String(applicationId));
    if (rowIndex === -1) throw new Error('指定された申請が見つかりません。');

    const targetRow = values[rowIndex];
    const currentStage = stageCol !== -1 ? targetRow[stageCol] : '';
    const targetMaster = targetMasterCol !== -1 ? targetRow[targetMasterCol] : '';

    // 決裁者却下の場合、常務承認済からのロールバック
    if (currentStage === '常務承認済' && !isReject) {
      // 顧客マスタのロールバック
      const rollbackResult = rollbackData(applicationId);
      if (!rollbackResult.success) {
        console.warn('顧客ロールバック失敗:', rollbackResult.message);
      }

      // ★ 追加: 単価マスタのロールバック
      if (targetMaster && targetMaster.includes('単価マスタ')) {
        console.log('[rejectApplication] 単価マスタのロールバックを実行');
        const priceRollbackResult = rollbackPriceData(applicationId);
        if (!priceRollbackResult.success) {
          console.warn('単価ロールバック失敗:', priceRollbackResult.message);
        }
      }
    }

    // ステータスと承認段階を更新
    const newStatus = isReject ? '却下' : '却下';
    console.log('[rejectApplication] currentStage:', currentStage, 'newStatus:', newStatus);
    if (statusCol !== -1) {
      appSheet.getRange(rowIndex + 2, statusCol + 1).setValue(newStatus);
    }
    // ★ 追加: 承認段階も更新
    if (stageCol !== -1) {
      appSheet.getRange(rowIndex + 2, stageCol + 1).setValue(newStatus);
    }

    // 理由を記録
    if (rejectReasonCol !== -1) {
      appSheet.getRange(rowIndex + 2, rejectReasonCol + 1).setValue(reason);
    }

    // ★ 追加: 修正指示項目をJSON形式で保存
    if (requiredFieldsCol !== -1 && requiredFields && requiredFields.length > 0) {
      const requiredFieldsJson = JSON.stringify(requiredFields);
      appSheet.getRange(rowIndex + 2, requiredFieldsCol + 1).setValue(requiredFieldsJson);
      console.log(`修正指示項目を保存: ${requiredFieldsJson}`);
    }

    // ★ 追加: 判断者名と時刻を記録
    console.log('[rejectApplication] rejectorId:', rejectorId, 'isReject:', isReject);
    if (rejectorId) {
      console.log('[rejectApplication] updateRejectionStage呼び出し');
      updateRejectionStage(applicationId, currentStage, rejectorId);
    }

    // 通知先に通知
    Logger.log(`[rejectApplication] 却下通知 notifyTo=${notifyTo || '(未指定)'} applicationId=${applicationId}`);
    if (notifyTo) {
      const subject = '申請が却下されました';
      const appUrl = getApplicationDetailUrl(applicationId);

      // 宛先ごとに個別送信（宛名を付けるため）
      const recipients = notifyTo.split(',').map(e => e.trim()).filter(e => e);
      recipients.forEach(recipientEmail => {
        const recipientName = findEmployeeByEmail(recipientEmail) || '';
        const greeting = recipientName ? `${recipientName} 様\n\n` : '';

        let body = `${greeting}申請が却下されました。\n\n` +
                   `却下理由・コメント:\n${reason}\n`;

        if (requiredFields && requiredFields.length > 0) {
          body += `\n修正が必要な項目:\n` + requiredFields.map(f => `- ${f}`).join('\n') + '\n';
        }

        body += `\nアプリのリンクはこちら:\n${appUrl}`;

        sendNotificationEmail(recipientEmail, subject, body);
      });
    } else {
      Logger.log('[rejectApplication] notifyToが空のため却下通知をスキップ');
    }

    return {
      success: true,
      message: `${newStatus}処理を完了しました。`
    };

  } catch (e) {
    console.error('却下/却下処理エラー:', e);
    return {
      success: false,
      message: `処理に失敗しました: ${e.message}`
    };
  }
}

/**
 * 承認段階と判断者を記録します
 * @param {string} applicationId - 申請ID
 * @param {string} newStage - 新しい承認段階
 * @param {string} approverId - 承認者ID
 */
function updateApprovalStage(applicationId, newStage, approverId) {
  try {
    const appSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('申請管理');
    if (!appSheet) return;

    const values = appSheet.getDataRange().getValues();
    const header = values.shift();

    const idCol = header.indexOf('申請ID');
    const rowIndex = values.findIndex(r => String(r[idCol]) === String(applicationId));
    if (rowIndex === -1) return;

    // 判断者名を取得
    const approver = findEmployeeById(approverId);
    const approverName = approver ? approver.name : approverId;

    // 現在時刻を取得
    const now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");

    // 承認段階に応じて判断者と時刻を記録
    let judgeCol = -1;
    let timestampCol = -1;

    switch (newStage) {
      case '上長承認済':
        judgeCol = header.indexOf('上長判断者名');
        if (judgeCol === -1) judgeCol = header.indexOf('上長承認者名'); // 互換性のため旧名もサポート
        if (judgeCol === -1) judgeCol = header.indexOf('上長承認'); // 互換性のため旧名もサポート
        if (judgeCol === -1) judgeCol = header.indexOf('上長承認者'); // 互換性のため旧名もサポート
        timestampCol = header.indexOf('上長判断時刻');
        if (timestampCol === -1) timestampCol = header.indexOf('上長承認時刻'); // 互換性のため旧名もサポート
        break;
      case '管理承認済':
        judgeCol = header.indexOf('管理判断者名');
        if (judgeCol === -1) judgeCol = header.indexOf('管理承認者名'); // 互換性のため旧名もサポート
        if (judgeCol === -1) judgeCol = header.indexOf('管理承認'); // 互換性のため旧名もサポート
        if (judgeCol === -1) judgeCol = header.indexOf('管理承認者'); // 互換性のため旧名もサポート
        timestampCol = header.indexOf('管理判断時刻');
        if (timestampCol === -1) timestampCol = header.indexOf('管理承認時刻'); // 互換性のため旧名もサポート
        break;
      case '常務承認済':
        judgeCol = header.indexOf('常務判断者名');
        if (judgeCol === -1) judgeCol = header.indexOf('横山判断者名'); // 互換性のため旧名もサポート
        if (judgeCol === -1) judgeCol = header.indexOf('常務承認');
        if (judgeCol === -1) judgeCol = header.indexOf('横山さん承認'); // 互換性のため旧名もサポート
        timestampCol = header.indexOf('常務判断時刻');
        if (timestampCol === -1) timestampCol = header.indexOf('横山判断時刻'); // 互換性のため旧名もサポート
        break;
      case '決裁完了':
        judgeCol = header.indexOf('決裁者判断者名');
        if (judgeCol === -1) judgeCol = header.indexOf('社長判断者名'); // 互換性のため旧名もサポート
        if (judgeCol === -1) judgeCol = header.indexOf('決裁者承認');
        if (judgeCol === -1) judgeCol = header.indexOf('社長承認'); // 互換性のため旧名もサポート
        timestampCol = header.indexOf('決裁者判断時刻');
        if (timestampCol === -1) timestampCol = header.indexOf('社長判断時刻'); // 互換性のため旧名もサポート
        break;
    }

    // 判断者名を記録
    if (judgeCol !== -1) {
      appSheet.getRange(rowIndex + 2, judgeCol + 1).setValue(approverName);
    }

    // 判断時刻を記録
    if (timestampCol !== -1) {
      appSheet.getRange(rowIndex + 2, timestampCol + 1).setValue(now);
    }

    // 旧名の列に「承認済み」を記録
    let legacyCol = -1;
    switch (newStage) {
      case '上長承認済':
        legacyCol = header.indexOf('上長承認');
        break;
      case '管理承認済':
        legacyCol = header.indexOf('管理承認');
        break;
      case '常務承認済':
        legacyCol = header.indexOf('常務承認');
        if (legacyCol === -1) legacyCol = header.indexOf('管理部長承認');
        break;
      case '決裁完了':
        legacyCol = header.indexOf('社長承認');
        break;
    }
    console.log('[updateApprovalStage] newStage:', newStage, 'legacyCol:', legacyCol);
    if (legacyCol !== -1) {
      console.log('[updateApprovalStage] 承認済みを記録: row', rowIndex + 2, 'col', legacyCol + 1);
      appSheet.getRange(rowIndex + 2, legacyCol + 1).setValue('承認済み');
    } else {
      console.log('[updateApprovalStage] 警告: 旧名の列が見つかりません');
    }

  } catch (e) {
    console.error('判断記録エラー:', e);
  }
}

/**
 * 却下時の判断者と時刻を記録します
 * @param {string} applicationId - 申請ID
 * @param {string} currentStage - 現在の承認段階
 * @param {string} rejectorId - 却下者ID
 */
function updateRejectionStage(applicationId, currentStage, rejectorId) {
  try {
    const appSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('申請管理');
    if (!appSheet) return;

    const values = appSheet.getDataRange().getValues();
    const header = values.shift();

    const idCol = header.indexOf('申請ID');
    const rowIndex = values.findIndex(r => String(r[idCol]) === String(applicationId));
    if (rowIndex === -1) return;

    // 判断者名を取得
    const rejector = findEmployeeById(rejectorId);
    const rejectorName = rejector ? rejector.name : rejectorId;

    // 現在時刻を取得
    const now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");

    // 現在の承認段階に応じて判断者と時刻を記録
    let judgeCol = -1;
    let timestampCol = -1;

    switch (currentStage) {
      case '申請中':
        // 上長が却下
        judgeCol = header.indexOf('上長判断者名');
        if (judgeCol === -1) judgeCol = header.indexOf('上長却下者名'); // 互換性のため旧名もサポート
        timestampCol = header.indexOf('上長判断時刻');
        if (timestampCol === -1) timestampCol = header.indexOf('上長却下時刻'); // 互換性のため旧名もサポート
        break;
      case '上長承認済':
        // 管理者が却下
        judgeCol = header.indexOf('管理判断者名');
        if (judgeCol === -1) judgeCol = header.indexOf('管理却下者名'); // 互換性のため旧名もサポート
        timestampCol = header.indexOf('管理判断時刻');
        if (timestampCol === -1) timestampCol = header.indexOf('管理却下時刻'); // 互換性のため旧名もサポート
        break;
      case '管理承認済':
        // 常務が却下
        judgeCol = header.indexOf('常務判断者名');
        if (judgeCol === -1) judgeCol = header.indexOf('横山判断者名'); // 互換性のため旧名もサポート
        if (judgeCol === -1) judgeCol = header.indexOf('常務却下者名');
        timestampCol = header.indexOf('常務判断時刻');
        if (timestampCol === -1) timestampCol = header.indexOf('横山判断時刻'); // 互換性のため旧名もサポート
        break;
      case '常務承認済':
        // 決裁者が却下
        judgeCol = header.indexOf('決裁者判断者名');
        if (judgeCol === -1) judgeCol = header.indexOf('社長判断者名');
        if (judgeCol === -1) judgeCol = header.indexOf('決裁者却下者名');
        timestampCol = header.indexOf('決裁者判断時刻');
        if (timestampCol === -1) timestampCol = header.indexOf('社長判断時刻');
        if (timestampCol === -1) timestampCol = header.indexOf('決裁者承認時刻');
        if (timestampCol === -1) timestampCol = header.indexOf('社長承認時刻');
        break;
    }

    // 判断者名を記録
    if (judgeCol !== -1) {
      appSheet.getRange(rowIndex + 2, judgeCol + 1).setValue(rejectorName);
    }

    // 判断時刻を記録
    if (timestampCol !== -1) {
      appSheet.getRange(rowIndex + 2, timestampCol + 1).setValue(now);
    }

    // 旧名の列に「却下」を記録
    let legacyCol = -1;
    switch (currentStage) {
      case '申請中':
        legacyCol = header.indexOf('上長承認');
        break;
      case '上長承認済':
        legacyCol = header.indexOf('管理承認');
        break;
      case '管理承認済':
        legacyCol = header.indexOf('常務承認');
        if (legacyCol === -1) legacyCol = header.indexOf('管理部長承認');
        break;
      case '常務承認済':
        legacyCol = header.indexOf('社長承認');
        break;
    }
    console.log('[updateRejectionStage] currentStage:', currentStage, 'legacyCol:', legacyCol);
    if (legacyCol !== -1) {
      console.log('[updateRejectionStage] 却下を記録: row', rowIndex + 2, 'col', legacyCol + 1);
      appSheet.getRange(rowIndex + 2, legacyCol + 1).setValue('却下');
    } else {
      console.log('[updateRejectionStage] 警告: 旧名の列が見つかりません');
    }

  } catch (e) {
    console.error('判断記録エラー:', e);
  }
}

/**
 * IDで社員を検索します
 * @param {string} employeeId - 社員ID
 * @return {Object|null} 社員情報
 */
function findEmployeeById(employeeId) {
  try {
    const empSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('社員マスタ');
    if (!empSheet) return null;

    const values = empSheet.getDataRange().getValues();
    const header = values.shift();

    const idCol = header.indexOf('ID');
    const nameCol = header.indexOf('担当者名');
    const deptCodeCol = header.indexOf('部門コード');
    const deptNameCol = header.indexOf('部門名');
    const divisionCol = header.indexOf('大区分');
    const emailCol = header.indexOf('メールアドレス');
    const roleCol = header.indexOf('役職');

    if (idCol === -1 || nameCol === -1) return null;

    const employee = values.find(r => cleanSingleQuotes(r[idCol]) === cleanSingleQuotes(employeeId));
    if (!employee) return null;

    return {
      id: cleanSingleQuotes(employee[idCol]),
      name: employee[nameCol],
      departmentCode: deptCodeCol !== -1 ? employee[deptCodeCol] : '',
      department: deptNameCol !== -1 ? employee[deptNameCol] : '',
      division: divisionCol !== -1 ? String(employee[divisionCol]).trim() : '',
      email: emailCol !== -1 ? employee[emailCol] : '',
      role: roleCol !== -1 ? employee[roleCol] : ''
    };

  } catch (e) {
    console.error('社員検索エラー:', e);
    return null;
  }
}

// =============================================
// メール通知
// =============================================

/**
 * メール通知を送信します
 * @param {string} to - 送信先メールアドレス
 * @param {string} subject - 件名
 * @param {string} body - 本文
 */
function findEmployeeByEmail(email) {
  try {
    const empSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('社員マスタ');
    if (!empSheet) return null;
    const values = empSheet.getDataRange().getValues();
    const header = values.shift();
    const emailCol = header.indexOf('メールアドレス');
    const nameCol = header.indexOf('担当者名');
    if (emailCol === -1 || nameCol === -1) return null;
    const row = values.find(r => String(r[emailCol]).trim() === String(email).trim());
    return row ? row[nameCol] : null;
  } catch (e) {
    return null;
  }
}

function sendNotificationEmail(to, subject, body) {
  try {
    if (!to) {
      Logger.log('[sendNotificationEmail] 送信先メールアドレスが指定されていないためスキップ');
      console.warn('送信先メールアドレスが指定されていません。');
      return;
    }

    Logger.log(`[sendNotificationEmail] 送信開始 to=${to} subject=${subject}`);
    MailApp.sendEmail({
      to: to,
      subject: subject,
      body: body
    });

    Logger.log(`[sendNotificationEmail] 送信成功 to=${to}`);
    console.log(`メール送信完了: ${to}`);

  } catch (e) {
    Logger.log(`[sendNotificationEmail] 送信エラー to=${to} error=${e.message}`);
    console.error('メール送信エラー:', e);
  }
}

/**
 * 次の承認者に通知を送信します（複数上長対応）
 * @param {string} applicationId - 申請ID
 */
function notifyNextApprover(applicationId) {
  try {
    const appSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('申請管理');
    if (!appSheet) return;

    const values = appSheet.getDataRange().getValues();
    const header = values.shift();

    const idCol = header.indexOf('申請ID');
    const stageCol = header.indexOf('承認段階');
    const customerNameCol = header.indexOf('対象顧客名');
    const applicantCol = header.indexOf('申請者名');
    const applicantIdCol = header.indexOf('申請者ID');

    const rowIndex = values.findIndex(r => String(r[idCol]) === String(applicationId));
    if (rowIndex === -1) return;

    const targetRow = values[rowIndex];
    const currentStage = stageCol !== -1 ? targetRow[stageCol] : '';
    const customerName = customerNameCol !== -1 ? targetRow[customerNameCol] : '';
    const applicantName = applicantCol !== -1 ? targetRow[applicantCol] : '';
    const applicantId = applicantIdCol !== -1 ? cleanSingleQuotes(targetRow[applicantIdCol]) : null;

    // メール共通情報
    const subject = '【承認依頼】顧客情報の承認をお願いします';
    const appUrl = getApplicationDetailUrl(applicationId);

    const buildBody = (recipientName) => {
      const greeting = recipientName ? `${recipientName} 様\n\n` : '';
      return `${greeting}承認依頼が届いています。\n\n` +
             `対象顧客: ${customerName}\n` +
             `申請者: ${applicantName}\n` +
             `現在の段階: ${currentStage}\n\n` +
             `アプリのリンクはこちら:\n${appUrl}`;
    };

    // 上長承認待ちの場合は複数上長に通知
    if (currentStage === '申請中' && applicantId) {
      const supervisors = findAllSupervisors(applicantId);
      supervisors.forEach(supervisor => {
        if (supervisor.email) {
          sendNotificationEmail(supervisor.email, subject, buildBody(supervisor.name));
        }
      });
    } else {
      // 上長以外は単一承認者に通知
      const nextApprover = getNextApprover(applicationId, currentStage);
      if (nextApprover && nextApprover.email) {
        sendNotificationEmail(nextApprover.email, subject, buildBody(nextApprover.name));
      }
    }

  } catch (e) {
    console.error('次承認者への通知エラー:', e);
  }
}

/**
 * 決裁完了時に申請者へ完了通知を送信します
 * @param {string} applicationId - 申請ID
 */
function notifyApplicantOnCompletion(applicationId) {
  try {
    Logger.log(`[notifyApplicantOnCompletion] 開始 applicationId=${applicationId}`);

    const appSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('申請管理');
    if (!appSheet) {
      Logger.log('[notifyApplicantOnCompletion] 申請管理シートが見つかりません');
      return;
    }

    const values = appSheet.getDataRange().getValues();
    const header = values.shift();

    const idCol = header.indexOf('申請ID');
    const applicantEmailCol = header.indexOf('申請者メール');
    const applicantNameCol = header.indexOf('申請者名');
    const customerNameCol = header.indexOf('対象顧客名');
    const typeCol = header.indexOf('申請種別');

    Logger.log(`[notifyApplicantOnCompletion] 列インデックス applicantEmailCol=${applicantEmailCol} applicantNameCol=${applicantNameCol}`);

    const rowIndex = values.findIndex(r => String(r[idCol]) === String(applicationId));
    if (rowIndex === -1) {
      Logger.log(`[notifyApplicantOnCompletion] 申請ID ${applicationId} が見つかりません`);
      return;
    }

    const targetRow = values[rowIndex];
    const applicantEmail = applicantEmailCol !== -1 ? targetRow[applicantEmailCol] : '';
    const applicantName = applicantNameCol !== -1 ? targetRow[applicantNameCol] : '';
    const customerName = customerNameCol !== -1 ? targetRow[customerNameCol] : '';
    const appType = typeCol !== -1 ? targetRow[typeCol] : '';

    Logger.log(`[notifyApplicantOnCompletion] 申請者名=${applicantName} メール=${applicantEmail} 顧客=${customerName} 種別=${appType}`);

    if (!applicantEmail) {
      Logger.log(`[notifyApplicantOnCompletion] 申請者メールが空のためスキップ（申請ID=${applicationId}）`);
      console.warn('[notifyApplicantOnCompletion] 申請者メールが未登録のため完了通知をスキップします。申請ID:', applicationId);
      return;
    }

    const appUrl = getApplicationDetailUrl(applicationId);
    const subject = '【決裁完了】申請が承認されました';
    const body = `${applicantName} 様\n\n` +
                 `ご申請いただいた内容が決裁完了しました。\n\n` +
                 `申請種別: ${appType}\n` +
                 `対象顧客: ${customerName}\n\n` +
                 `アプリのリンクはこちら:\n${appUrl}`;

    sendNotificationEmail(applicantEmail, subject, body);
    Logger.log(`[notifyApplicantOnCompletion] 完了`);

  } catch (e) {
    Logger.log(`[notifyApplicantOnCompletion] エラー: ${e.message}`);
    console.error('申請完了通知エラー:', e);
  }
}

/**
 * 上長に特殊通知を送信します（部下の申請が却下/最終承認された際）
 * ※現在未使用
 * @param {string} applicationId - 申請ID
 * @param {string} decision - 決定内容（'approved' or 'rejected'）
 */
function notifySupervisorOnDecision(applicationId, decision) {
  try {
    const appSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('申請管理');
    if (!appSheet) return;

    const values = appSheet.getDataRange().getValues();
    const header = values.shift();

    const idCol = header.indexOf('申請ID');
    const applicantIdCol = header.indexOf('申請者ID');
    const applicantNameCol = header.indexOf('申請者名');
    const customerNameCol = header.indexOf('対象顧客名');

    const rowIndex = values.findIndex(r => String(r[idCol]) === String(applicationId));
    if (rowIndex === -1) return;

    const targetRow = values[rowIndex];
    const applicantId = applicantIdCol !== -1 ? cleanSingleQuotes(targetRow[applicantIdCol]) : null;
    const applicantName = applicantNameCol !== -1 ? targetRow[applicantNameCol] : '';
    const customerName = customerNameCol !== -1 ? targetRow[customerNameCol] : '';

    // 上長を取得
    if (!applicantId) return;

    const supervisor = findSupervisor(applicantId);
    if (!supervisor || !supervisor.email) return;

    const subject = decision === 'approved'
      ? '【通知】部下の申請が承認されました'
      : '【通知】部下の申請が却下されました';

    const detailUrl = getApplicationDetailUrl(applicationId);
    const body = `部下の申請が処理されました。\n\n` +
                 `申請ID: ${applicationId}\n` +
                 `申請者: ${applicantName}\n` +
                 `対象顧客: ${customerName}\n` +
                 `結果: ${decision === 'approved' ? '承認' : '却下'}\n\n` +
                 `詳細はこちら:\n${detailUrl}`;

    sendNotificationEmail(supervisor.email, subject, body);

  } catch (e) {
    console.error('上長への通知エラー:', e);
  }
}

// =============================================
// マスタ更新処理（既存のprocessApplicationと統合するためのヘルパー）
// =============================================

/**
 * 顧客マスタを更新します
 * @param {string} applicationId - 申請ID
 * @return {Object} 結果 {success: boolean, message: string}
 */
function updateCustomerMaster(applicationId) {
  try {
    // 申請管理シートから得意先コードと申請種別、対象マスタを取得
    const appSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('申請管理');
    if (!appSheet) throw new Error('申請管理シートが見つかりません。');

    const values = appSheet.getDataRange().getValues();
    const header = values.shift();

    const idCol = header.indexOf('申請ID');
    const customerIdCol = header.indexOf('得意先コード');
    const appTypeCol = header.indexOf('申請種別');
    const targetMasterCol = header.indexOf('対象マスタ'); // ★ 追加
    const effectiveDateCol = header.indexOf('登録有効日'); // ★ 追加（単価用）

    if (idCol === -1 || customerIdCol === -1) {
      throw new Error('申請管理シートに必要な列が見つかりません。');
    }

    const rowIndex = values.findIndex(r => String(r[idCol]) === String(applicationId));
    if (rowIndex === -1) throw new Error('指定された申請が見つかりません。');

    const targetRow = values[rowIndex];
    let customerId = customerIdCol !== -1 ? cleanSingleQuotes(targetRow[customerIdCol]) : '';
    const appType = appTypeCol !== -1 ? targetRow[appTypeCol] : '';
    const targetMaster = targetMasterCol !== -1 ? targetRow[targetMasterCol] : ''; // ★ 追加

    // 管理部採番の得意先グループコードを取得
    const aggregationCodeCol = header.indexOf('得意先グループコード_管理部採番');
    const managementAggregationCode = aggregationCodeCol !== -1 ? targetRow[aggregationCodeCol] : '';

    console.log(`マスタ更新処理: 申請ID=${applicationId}, 得意先コード=${customerId}, 申請種別=${appType}, 得意先グループコード=${managementAggregationCode}`);

    if (!customerId) {
      throw new Error('得意先コードが見つかりません。');
    }

    // 申請データ_顧客から修正データを取得
    const detailSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('申請データ_顧客');
    if (!detailSheet) throw new Error('申請データ_顧客シートが見つかりません。');

    const detailValues = detailSheet.getDataRange().getValues();
    const detailHeader = detailValues.shift();

    const detailAppIdCol = detailHeader.indexOf('申請ID');
    const itemNameCol = detailHeader.indexOf('項目名');
    const newValueCol = detailHeader.indexOf('修正後の値');

    const relatedDetails = detailValues.filter(r =>
      String(r[detailAppIdCol]) === String(applicationId)
    );

    // 修正データをオブジェクトに変換
    const newData = {};
    relatedDetails.forEach(row => {
      const itemName = row[itemNameCol];
      const newValue = row[newValueCol];
      if (itemName) {
        newData[itemName] = newValue;
      }
    });

    let mergedData = {};

    if (appType === '顧客新規登録') {
      // 新規登録の場合: 申請データのみを使用
      mergedData = { ...newData };
      mergedData['得意先コード'] = customerId; // 管理部採番の得意先コード

      // 管理部採番の得意先グループコードを追加
      if (managementAggregationCode) {
        mergedData['得意先グループコード'] = managementAggregationCode;
        console.log(`得意先グループコード（管理部採番）を設定: ${managementAggregationCode}`);
      }

      // 日付の設定
      const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");
      mergedData['登録日'] = today;
      mergedData['最終更新日'] = today;
      console.log(`新規申請: 登録日と最終更新日を ${today} に設定`);

    } else {
      // 修正の場合: 元データを取得してマージ
      const originalData = getCustomerBasicInfo(customerId);
      if (!originalData || !originalData['得意先コード']) {
        throw new Error(`得意先コード ${customerId} の元データが見つかりません。`);
      }

      // 元データと修正データをマージ
      mergedData = { ...originalData, ...newData };
      mergedData['得意先コード'] = customerId; // 得意先コードを確保

      // 日付の更新
      const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");
      mergedData['最終更新日'] = today;
      console.log(`修正申請: 最終更新日を ${today} に設定`);
    }

    // 顧客マスタを更新
    if (appType === '顧客新規登録') {
      // 新規登録の場合はaddNewCustomerToMasterを使用
      addNewCustomerToMaster(mergedData);
      console.log(`新規顧客をマスタに追加: 得意先コード=${customerId}`);

      // マスタ更新履歴に記録（新規登録）
      try {
        const updater = Session.getActiveUser().getEmail() || 'システム';
        const changeDescription = `新規顧客登録: 得意先コード=${customerId}, 得意先名称=${mergedData['得意先名称'] || ''}（申請ID: ${applicationId}）`;
        addMasterUpdateHistory(updater, '得意先マスタ', changeDescription);
      } catch (historyError) {
        console.error('[updateCustomerMaster] マスタ更新履歴の記録に失敗しました:', historyError);
      }
    } else {
      // 修正の場合はupdateCustomerInMasterを使用
      updateCustomerInMaster(mergedData);
      console.log(`既存顧客をマスタで更新: 得意先コード=${customerId}`);

      // マスタ更新履歴に記録（修正）
      try {
        const updater = Session.getActiveUser().getEmail() || 'システム';
        // 変更内容を生成
        const changeItems = [];
        for (const key in newData) {
          if (newData.hasOwnProperty(key)) {
            const oldValue = originalData[key] || '';
            const newValue = newData[key] || '';
            if (String(oldValue) !== String(newValue)) {
              changeItems.push(`${key}: ${oldValue} → ${newValue}`);
            }
          }
        }
        const changeDescription = `得意先コード=${customerId} - ${changeItems.join(', ')}（申請ID: ${applicationId}）`;
        addMasterUpdateHistory(updater, '得意先マスタ', changeDescription);
      } catch (historyError) {
        console.error('[updateCustomerMaster] マスタ更新履歴の記録に失敗しました:', historyError);
      }
    }

    // ★ 追加: 単価マスタも更新が必要な場合（新規顧客+単価の場合）
    if (targetMaster && targetMaster.includes('単価マスタ')) {
      console.log(`対象マスタに「単価マスタ」が含まれているため、単価マスタも更新します (申請ID: ${applicationId})`);

      // 上代価格を取得 (Dateオブジェクトとして、または null)
      let effectiveDate = null;
      let targetPriceValue = '';
      if (effectiveDateCol !== -1 && targetRow[effectiveDateCol]) {
        targetPriceValue = targetRow[effectiveDateCol];
        if (targetPriceValue === '現行' || targetPriceValue === 'current') {
          // 現行の価格の場合は null として扱う
          effectiveDate = null;
        } else {
          try {
            effectiveDate = new Date(targetPriceValue);
            if (isNaN(effectiveDate.getTime())) effectiveDate = null;
          } catch(e) {
            console.warn('上代価格の日付解析に失敗しました:', targetPriceValue, e.message);
          }
        }
      }

      // ★ 未来日付チェック: 登録有効日が今日より後の場合は即時反映しない
      const today = new Date();
      today.setHours(0, 0, 0, 0); // 時刻をリセットして日付のみで比較

      let isFutureDate = false;
      let reflectionStatus = '反映済';

      if (effectiveDate && effectiveDate > today) {
        isFutureDate = true;
        const formattedDate = Utilities.formatDate(effectiveDate, Session.getScriptTimeZone(), 'yyyy/MM/dd');
        reflectionStatus = `未反映(${formattedDate}予定)`;
        console.log(`登録有効日が未来日付(${formattedDate})のため、単価マスタへの即時反映をスキップします。`);
      }

      // マスタ反映状態列を取得または作成
      const reflectionStatusCol = header.indexOf('マスタ反映状態');
      if (reflectionStatusCol === -1) {
        // 列が存在しない場合は作成
        appSheet.getRange(1, header.length + 1).setValue('マスタ反映状態');
        appSheet.getRange(rowIndex + 2, header.length + 1).setValue(reflectionStatus);
        console.log('「マスタ反映状態」列を新規作成しました。');
      } else {
        // 既存の列に値を設定
        appSheet.getRange(rowIndex + 2, reflectionStatusCol + 1).setValue(reflectionStatus);
      }

      if (!isFutureDate) {
        // 未来日付でない場合のみ、即時にマスタを更新
        const appDetails = getApplicationDetails(applicationId);

        // 単価マスタを更新（Code.jsの関数を呼び出し）
        updatePriceInMaster(customerId, appDetails, effectiveDate);
        console.log(`単価マスタの更新が完了しました (顧客ID: ${customerId})`);

        // マスタ更新履歴に記録（単価マスタ）
        try {
          const updater = Session.getActiveUser().getEmail() || 'システム';
          const priceCount = appDetails.prices ? appDetails.prices.length : 0;
          const effectiveDateStr = effectiveDate ? Utilities.formatDate(effectiveDate, Session.getScriptTimeZone(), 'yyyy/MM/dd') : '現行';
          const changeDescription = `得意先コード=${customerId} - ${priceCount}件の商品登録情報を更新（上代価格: ${effectiveDateStr}）（申請ID: ${applicationId}）`;
          addMasterUpdateHistory(updater, '単価マスタ', changeDescription);
        } catch (historyError) {
          console.error('[updateCustomerMaster] 単価マスタ更新履歴の記録に失敗しました:', historyError);
        }
      } else {
        console.log(`未来日付のため、単価マスタへの反映は ${reflectionStatus} に設定されました。`);
      }
    }

    // キャッシュをクリア（Code.jsの関数を呼び出し）
    clearCustomersCache();

    return {
      success: true,
      message: 'マスタを更新しました。'
    };

  } catch (e) {
    console.error('マスタ更新エラー:', e);
    return {
      success: false,
      message: `マスタ更新に失敗しました: ${e.message}`
    };
  }
}

// =============================================
// 権限判定機能（セキュリティ）
// =============================================

/**
 * ユーザーの役割を判定します
 * @param {Object} user - ユーザー情報
 * @return {string} 役割（'supervisor', 'manager', 'division_manager', 'approver', 'applicant'）
 */
function determineUserRole(user) {
  try {
    const role = user.role || '';
    const name = user.name || '';

    // 決裁者（役職で判定）
    if (role === '決裁者') {
      return 'approver';
    }

    // 常務（役職で判定）
    if (role === '常務') {
      return 'division_manager';
    }

    // 管理部門（役職で判定）
    if (role === '管理部門') {
      return 'manager';
    }

    // 上長（役職で判定）
    if (role === '上長') {
      return 'supervisor';
    }

    // 一般申請者
    return 'applicant';

  } catch (e) {
    console.error('役割判定エラー:', e);
    return 'applicant';
  }
}

/**
 * ユーザーが申請を閲覧できるか判定します
 * @param {string} userRole - ユーザーの役割
 * @param {string} stage - 承認段階
 * @param {Object} currentUser - 現在のユーザー情報
 * @param {Array} applicationRow - 申請データの行
 * @param {Array} header - 申請管理シートのヘッダー
 * @return {boolean} 閲覧可能か
 */
function canViewApplication(userRole, stage, currentUser, applicationRow, header) {
  try {
    console.log(`[canViewApplication] userRole=${userRole}, stage=${stage}, currentUser.id=${currentUser.id}`);

    // 役割と承認段階のマッチング
    switch (userRole) {
      case 'supervisor':
        // 上長：「申請中」の部下の申請 + 過去に自分が承認した申請（上長承認済以降）
        console.log(`[canViewApplication] 上長チェック - stage=${stage}`);
        if (stage === '申請中') {
          // 申請者が自分の部下かチェック
          const isSubord = isSubordinate(currentUser, applicationRow, header);
          console.log(`[canViewApplication] 部下チェック結果: ${isSubord}`);
          return isSubord;
        }
        // 過去に承認した案件も表示（上長承認済、管理承認済、常務承認済、決裁完了）
        if (['上長承認済', '管理承認済', '常務承認済', '決裁完了'].includes(stage)) {
          const isSubord = isSubordinate(currentUser, applicationRow, header);
          console.log(`[canViewApplication] 過去承認案件の部下チェック結果: ${isSubord}`);
          return isSubord;
        }
        return false;

      case 'manager':
        // 管理部門：「上長承認済」の承認待ち + 過去に承認した申請（管理承認済以降）
        if (stage === '上長承認済') {
          return true; // 承認待ち
        }
        // 過去に承認した案件も表示
        if (['管理承認済', '常務承認済', '決裁完了'].includes(stage)) {
          return true;
        }
        return false;

      case 'division_manager':
        // 常務：「管理承認済」の承認待ち + 過去に承認した申請（常務承認済以降）
        if (stage === '管理承認済') {
          return true; // 承認待ち
        }
        // 過去に承認した案件も表示
        if (['常務承認済', '決裁完了'].includes(stage)) {
          return true;
        }
        return false;

      case 'approver':
        // 決裁者：「常務承認済」の承認待ち + 過去に承認した申請（決裁完了）
        if (stage === '常務承認済') {
          return true; // 承認待ち
        }
        // 過去に承認した案件も表示
        if (stage === '決裁完了') {
          return true;
        }
        return false;

      case 'applicant':
      default:
        // 一般申請者：自分の申請のみ（この関数が呼ばれる前にチェック済み）
        return false;
    }

  } catch (e) {
    console.error('閲覧権限チェックエラー:', e);
    return false;
  }
}

/**
 * 申請者が上長の部下かチェックします
 * @param {Object} supervisor - 上長のユーザー情報
 * @param {Array} applicationRow - 申請データの行
 * @param {Array} header - 申請管理シートのヘッダー
 * @return {boolean} 部下かどうか
 */
function isSubordinate(supervisor, applicationRow, header) {
  try {
    const applicantIdCol = header.indexOf('申請者ID');
    console.log(`[isSubordinate] applicantIdCol=${applicantIdCol}`);
    if (applicantIdCol === -1) {
      console.log('[isSubordinate] 申請者ID列が見つからない');
      return false;
    }

    const applicantId = applicationRow[applicantIdCol];
    console.log(`[isSubordinate] applicantId=${applicantId}`);
    if (!applicantId) {
      console.log('[isSubordinate] 申請者IDが空');
      return false;
    }

    // 申請者の情報を取得
    const applicant = findEmployeeById(applicantId);
    console.log(`[isSubordinate] applicant=${JSON.stringify(applicant)}`);
    if (!applicant) {
      console.log('[isSubordinate] 申請者が見つからない');
      return false;
    }

    // 同じ大区分かチェック（文字列として比較）
    const supervisorDivision = String(supervisor.division || '').trim();
    const applicantDivision = String(applicant.division || '').trim();
    console.log(`[isSubordinate] supervisorDivision=${supervisorDivision}, applicantDivision=${applicantDivision}`);

    // 文字列として比較
    const result = supervisorDivision && applicantDivision && supervisorDivision === applicantDivision;
    console.log(`[isSubordinate] 大区分比較: ${supervisorDivision} === ${applicantDivision} → ${result}`);
    return result;

  } catch (e) {
    console.error('[isSubordinate] 部下チェックエラー:', e);
    return false;
  }
}

// =============================================
// 未来日付の単価マスタ反映処理（日次トリガー用）
// =============================================

/**
 * 未来日付で承認済みの単価申請を確認し、有効日が到来したらマスタに反映する
 * この関数は日次トリガー（深夜）で実行することを想定
 */
function processPendingPriceReflections() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const appSheet = ss.getSheetByName('申請管理');
    if (!appSheet) {
      console.error('[processPendingPriceReflections] 申請管理シートが見つかりません。');
      return;
    }

    const values = appSheet.getDataRange().getValues();
    const header = values[0];
    const dataRows = values.slice(1);

    // 必要な列のインデックスを取得
    const idCol = header.indexOf('申請ID');
    const statusCol = header.indexOf('承認ステータス');
    const targetMasterCol = header.indexOf('対象マスタ');
    const effectiveDateCol = header.indexOf('登録有効日');
    const reflectionStatusCol = header.indexOf('マスタ反映状態');
    const customerIdCol = header.indexOf('得意先コード');

    if (idCol === -1 || statusCol === -1 || reflectionStatusCol === -1) {
      console.error('[processPendingPriceReflections] 必要な列が見つかりません。');
      return;
    }

    const today = new Date();
    today.setHours(0, 0, 0, 0);

    let processedCount = 0;

    // データ行を走査
    dataRows.forEach((row, index) => {
      const applicationId = row[idCol];
      const status = row[statusCol];
      const targetMaster = targetMasterCol !== -1 ? row[targetMasterCol] : '';
      const reflectionStatus = row[reflectionStatusCol];
      const customerId = customerIdCol !== -1 ? cleanSingleQuotes(row[customerIdCol]) : '';

      // 条件: 決裁完了 かつ 対象マスタに単価マスタ含む かつ マスタ反映状態が「未反映」で始まる
      if (
        status === '決裁完了' &&
        targetMaster && targetMaster.includes('単価マスタ') &&
        reflectionStatus && String(reflectionStatus).startsWith('未反映')
      ) {
        // 登録有効日を取得
        const effectiveDateValue = effectiveDateCol !== -1 ? row[effectiveDateCol] : '';
        if (!effectiveDateValue) {
          console.log(`[processPendingPriceReflections] 申請ID ${applicationId}: 登録有効日が空のためスキップ`);
          return;
        }

        let effectiveDate = null;
        try {
          effectiveDate = new Date(effectiveDateValue);
          if (isNaN(effectiveDate.getTime())) {
            console.warn(`[processPendingPriceReflections] 申請ID ${applicationId}: 登録有効日の解析に失敗 (${effectiveDateValue})`);
            return;
          }
          effectiveDate.setHours(0, 0, 0, 0);
        } catch (e) {
          console.warn(`[processPendingPriceReflections] 申請ID ${applicationId}: 登録有効日の解析エラー`, e);
          return;
        }

        // 有効日が今日以前かチェック
        if (effectiveDate <= today) {
          console.log(`[processPendingPriceReflections] 申請ID ${applicationId}: 有効日到来 (${Utilities.formatDate(effectiveDate, Session.getScriptTimeZone(), 'yyyy/MM/dd')}) - マスタ反映を開始`);

          try {
            // 申請詳細データを取得
            const appDetails = getApplicationDetails(applicationId);

            // 単価マスタを更新
            updatePriceInMaster(customerId, appDetails, effectiveDate);
            console.log(`[processPendingPriceReflections] 申請ID ${applicationId}: 単価マスタ更新完了`);

            // マスタ反映状態を「反映済」に更新
            appSheet.getRange(index + 2, reflectionStatusCol + 1).setValue('反映済');
            console.log(`[processPendingPriceReflections] 申請ID ${applicationId}: マスタ反映状態を「反映済」に更新`);

            // マスタ更新履歴に記録
            try {
              const updater = 'システム（日次トリガー）';
              const priceCount = appDetails.prices ? appDetails.prices.length : 0;
              const effectiveDateStr = Utilities.formatDate(effectiveDate, Session.getScriptTimeZone(), 'yyyy/MM/dd');
              const changeDescription = `得意先コード=${customerId} - ${priceCount}件の商品登録情報を自動反映（上代価格: ${effectiveDateStr}）（申請ID: ${applicationId}）`;
              addMasterUpdateHistory(updater, '単価マスタ', changeDescription);
            } catch (historyError) {
              console.error('[processPendingPriceReflections] マスタ更新履歴の記録に失敗:', historyError);
            }

            processedCount++;

          } catch (updateError) {
            console.error(`[processPendingPriceReflections] 申請ID ${applicationId}: マスタ反映処理エラー`, updateError);
            // エラーが発生しても処理を続行
          }
        }
      }
    });

    // キャッシュをクリア
    if (processedCount > 0) {
      clearCustomersCache();
      console.log(`[processPendingPriceReflections] 処理完了: ${processedCount}件の申請をマスタに反映しました。`);
    } else {
      console.log('[processPendingPriceReflections] 反映対象の申請はありませんでした。');
    }

  } catch (e) {
    console.error('[processPendingPriceReflections] 全体エラー:', e);
  }
}
