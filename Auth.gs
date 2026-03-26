/**
 * @fileoverview 認証機能
 * 社員マスタベースのログイン認証とセッション管理を提供します。
 */

// =============================================
// セッション管理用の定数
// =============================================
const SESSION_KEY_USER_ID = 'authenticated_user_id';
const SESSION_KEY_USER_NAME = 'authenticated_user_name';
const SESSION_KEY_USER_EMAIL = 'authenticated_user_email';
const SESSION_KEY_USER_ROLE = 'authenticated_user_role';
const SESSION_KEY_USER_DEPT = 'authenticated_user_dept';
const SESSION_KEY_USER_DEPT_CODE = 'authenticated_user_dept_code';
const SESSION_KEY_USER_DIVISION = 'authenticated_user_division'; // ★ 追加
const SESSION_KEY_ADMIN_PERMISSION = 'authenticated_admin_permission'; // ★ 追加
const SESSION_KEY_USER_HANDLER_CODE = 'authenticated_user_handler_code'; // ★ 追加
const SESSION_KEY_TIMESTAMP = 'authenticated_timestamp';
const SESSION_TIMEOUT_MS = 8 * 60 * 60 * 1000; // 8時間

/**
 * ログイン処理
 * 社員マスタからID/パスワードを照合し、認証を実行します。
 *
 * @param {string} id - ユーザーID
 * @param {string} password - パスワード
 * @return {Object} 認証結果 { success: boolean, message: string, user?: Object }
 */
function login(id, password) {
  try {
    // 入力チェック
    if (!id || !password) {
      return { success: false, message: 'IDとパスワードを入力してください。' };
    }

    // 社員マスタシートを取得
    const employeeSheet = SPREADSHEET.getSheetByName('社員マスタ');
    if (!employeeSheet) {
      throw new Error('社員マスタシートが見つかりません。');
    }

    const values = employeeSheet.getDataRange().getValues();
    if (values.length <= 1) {
      throw new Error('社員マスタにデータが存在しません。');
    }

    // ヘッダー行を取得し、必要な列のインデックスを動的に取得
    const header = values[0];
    const colIndex = {
      id: header.indexOf('ID'),
      password: header.indexOf('PW'),
      name: header.indexOf('担当者名'),
      email: header.indexOf('メールアドレス'),
      role: header.indexOf('役職'),
      department: header.indexOf('部門名'),
      deptCode: header.indexOf('部門コード'),
      division: header.indexOf('大区分'),
      adminPermission: header.indexOf('管理権限'), // ★追加
      handlerCode: header.indexOf('担当者コード') // ★追加
    };

    // 必須カラムの存在チェック
    if (colIndex.id === -1 || colIndex.password === -1) {
      throw new Error('社員マスタに「ID」または「PW」列が見つかりません。');
    }
    if (colIndex.name === -1) {
      throw new Error('社員マスタに「担当者名」列が見つかりません。');
    }

    // IDとパスワードを正規化（シングルクォート除去）
    const normalizedId = cleanSingleQuotes(String(id).trim());
    const normalizedPassword = cleanSingleQuotes(String(password).trim());

    // 社員マスタから該当ユーザーを検索
    const dataRows = values.slice(1);
    const matchedRow = dataRows.find(row => {
      const rowId = cleanSingleQuotes(String(row[colIndex.id] || '').trim());
      const rowPassword = cleanSingleQuotes(String(row[colIndex.password] || '').trim());
      return rowId === normalizedId && rowPassword === normalizedPassword;
    });

    if (!matchedRow) {
      return { success: false, message: 'IDまたはパスワードが正しくありません。' };
    }

    // ユーザー情報を取得
    const user = {
      id: matchedRow[colIndex.id],
      name: matchedRow[colIndex.name] || '',
      email: matchedRow[colIndex.email] || '',
      role: matchedRow[colIndex.role] || '',
      department: matchedRow[colIndex.department] || '',
      deptCode: matchedRow[colIndex.deptCode] || '',
      division: colIndex.division !== -1 ? String(matchedRow[colIndex.division]).trim() : '',
      adminPermission: colIndex.adminPermission !== -1 ? String(matchedRow[colIndex.adminPermission]).trim() : '', // ★追加
      handlerCode: colIndex.handlerCode !== -1 ? String(matchedRow[colIndex.handlerCode]).trim() : '' // ★追加
    };

    // セッションに保存
    const userProperties = PropertiesService.getUserProperties();
    userProperties.setProperties({
      [SESSION_KEY_USER_ID]: user.id,
      [SESSION_KEY_USER_NAME]: user.name,
      [SESSION_KEY_USER_EMAIL]: user.email,
      [SESSION_KEY_USER_ROLE]: user.role,
      [SESSION_KEY_USER_DEPT]: user.department,
      [SESSION_KEY_USER_DEPT_CODE]: user.deptCode,
      [SESSION_KEY_USER_DIVISION]: user.division, // ★ 追加
      [SESSION_KEY_ADMIN_PERMISSION]: user.adminPermission, // ★ 追加
      [SESSION_KEY_USER_HANDLER_CODE]: user.handlerCode, // ★ 追加
      [SESSION_KEY_TIMESTAMP]: new Date().getTime().toString()
    });

    return {
      success: true,
      message: 'ログインに成功しました。',
      user: user
    };

  } catch (e) {
    console.error('Login error:', e);
    return {
      success: false,
      message: 'ログイン処理中にエラーが発生しました: ' + e.message
    };
  }
}

/**
 * ログアウト処理
 * セッション情報をクリアします。
 *
 * @return {Object} 処理結果 { success: boolean, message: string }
 */
function logout() {
  try {
    const userProperties = PropertiesService.getUserProperties();
    userProperties.deleteProperty(SESSION_KEY_USER_ID);
    userProperties.deleteProperty(SESSION_KEY_USER_NAME);
    userProperties.deleteProperty(SESSION_KEY_USER_EMAIL);
    userProperties.deleteProperty(SESSION_KEY_USER_ROLE);
    userProperties.deleteProperty(SESSION_KEY_USER_DEPT);
    userProperties.deleteProperty(SESSION_KEY_USER_DEPT_CODE);
    userProperties.deleteProperty(SESSION_KEY_USER_DIVISION); // ★ 追加
    userProperties.deleteProperty(SESSION_KEY_ADMIN_PERMISSION); // ★ 追加
    userProperties.deleteProperty(SESSION_KEY_USER_HANDLER_CODE); // ★ 追加
    userProperties.deleteProperty(SESSION_KEY_TIMESTAMP);

    return { success: true, message: 'ログアウトしました。' };
  } catch (e) {
    console.error('Logout error:', e);
    return { success: false, message: 'ログアウト処理中にエラーが発生しました。' };
  }
}

/**
 * 現在ログインしているユーザー情報を取得します。
 *
 * @return {Object|null} ユーザー情報、またはログインしていない場合はnull
 */
function getCurrentUser() {
  try {
    if (!isAuthenticated()) {
      return null;
    }

    const userProperties = PropertiesService.getUserProperties();
    return {
      id: userProperties.getProperty(SESSION_KEY_USER_ID),
      name: userProperties.getProperty(SESSION_KEY_USER_NAME),
      email: userProperties.getProperty(SESSION_KEY_USER_EMAIL),
      role: userProperties.getProperty(SESSION_KEY_USER_ROLE),
      department: userProperties.getProperty(SESSION_KEY_USER_DEPT),
      departmentCode: userProperties.getProperty(SESSION_KEY_USER_DEPT_CODE),
      division: userProperties.getProperty(SESSION_KEY_USER_DIVISION), // ★ 追加
      adminPermission: userProperties.getProperty(SESSION_KEY_ADMIN_PERMISSION), // ★ 追加
      handlerCode: userProperties.getProperty(SESSION_KEY_USER_HANDLER_CODE) // ★ 追加
    };
  } catch (e) {
    console.error('Get current user error:', e);
    return null;
  }
}

/**
 * ユーザーが認証されているかチェックします。
 * セッションタイムアウトもチェックします。
 *
 * @return {boolean} 認証されている場合true
 */
function isAuthenticated() {
  try {
    const userProperties = PropertiesService.getUserProperties();
    const userId = userProperties.getProperty(SESSION_KEY_USER_ID);
    const timestamp = userProperties.getProperty(SESSION_KEY_TIMESTAMP);

    if (!userId || !timestamp) {
      return false;
    }

    // セッションタイムアウトチェック
    const loginTime = parseInt(timestamp, 10);
    const currentTime = new Date().getTime();

    if (currentTime - loginTime > SESSION_TIMEOUT_MS) {
      // タイムアウトの場合はセッションをクリア
      logout();
      return false;
    }

    return true;
  } catch (e) {
    console.error('Authentication check error:', e);
    return false;
  }
}

/**
 * セッションのタイムスタンプを更新します（アクティビティ検出用）
 *
 * @return {boolean} 更新成功した場合true
 */
function refreshSession() {
  try {
    if (!isAuthenticated()) {
      return false;
    }

    const userProperties = PropertiesService.getUserProperties();
    userProperties.setProperty(SESSION_KEY_TIMESTAMP, new Date().getTime().toString());
    return true;
  } catch (e) {
    console.error('Session refresh error:', e);
    return false;
  }
}
