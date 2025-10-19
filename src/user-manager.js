/**
 * ユーザー権限管理システム
 * 3段階権限制御（管理者・スタンダード・ゲスト）とセキュリティ機能
 */

// ユーザーロール定数
const USER_ROLES = {
  ADMINISTRATOR: 'Administrator',
  STANDARD: 'Standard', 
  GUEST: 'Guest'
};

// セッション管理定数
const SESSION_PROPERTIES = {
  CURRENT_USER: 'CURRENT_USER_ID',
  USER_ROLE: 'CURRENT_USER_ROLE',
  LOGIN_TIME: 'LOGIN_TIMESTAMP',
  SESSION_TIMEOUT: 'SESSION_TIMEOUT'
};

// セッションタイムアウト（8時間）
const SESSION_TIMEOUT_HOURS = 8;

/**
 * ユーザー管理シートの初期化
 */
function initializeUserManagementSheet() {
  try {
    console.log('🔐 ユーザー管理シート初期化開始...');
    
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = spreadsheet.getSheetByName('ユーザー管理');
    
    if (!sheet) {
      sheet = spreadsheet.insertSheet('ユーザー管理');
      console.log('✅ ユーザー管理シート作成完了');
    }
    
    // ヘッダー設定
    const headers = [
      'ユーザーID', 'ユーザー名', 'パスワードハッシュ', 'ロール',
      '作成日時', '最終ログイン', 'アクティブ状態', '作成者'
    ];
    
    if (sheet.getLastRow() === 0) {
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      
      // ヘッダーのスタイル設定
      const headerRange = sheet.getRange(1, 1, 1, headers.length);
      headerRange.setBackground('#4285f4');
      headerRange.setFontColor('white');
      headerRange.setFontWeight('bold');
      
      // 列幅調整
      sheet.setColumnWidth(1, 100); // ユーザーID
      sheet.setColumnWidth(2, 150); // ユーザー名
      sheet.setColumnWidth(3, 200); // パスワードハッシュ
      sheet.setColumnWidth(4, 120); // ロール
      sheet.setColumnWidth(5, 150); // 作成日時
      sheet.setColumnWidth(6, 150); // 最終ログイン
      sheet.setColumnWidth(7, 100); // アクティブ状態
      sheet.setColumnWidth(8, 120); // 作成者
      
      console.log('✅ ユーザー管理シートヘッダー設定完了');
    }
    
    // デフォルト管理者ユーザーの作成
    createDefaultAdminUser(sheet);
    
    return sheet;
    
  } catch (error) {
    console.error('❌ ユーザー管理シート初期化エラー:', error);
    throw new Error(`ユーザー管理シート初期化失敗: ${error.message}`);
  }
}

/**
 * デフォルト管理者ユーザーの作成
 */
function createDefaultAdminUser(sheet) {
  try {
    // 既存の管理者ユーザーがあるかチェック
    const data = sheet.getDataRange().getValues();
    const adminExists = data.some(row => row[3] === USER_ROLES.ADMINISTRATOR);
    
    if (!adminExists) {
      const adminUserId = generateUserId();
      const adminPassword = 'SalesAuto2024!'; // デフォルト管理者パスワード
      const passwordHash = hashPassword(adminPassword);
      
      const adminUser = [
        adminUserId,
        'システム管理者',
        passwordHash,
        USER_ROLES.ADMINISTRATOR,
        new Date(),
        '', // 最終ログイン（空）
        true, // アクティブ
        'システム初期化'
      ];
      
      sheet.appendRow(adminUser);
      console.log('✅ デフォルト管理者ユーザー作成完了');
      console.log(`管理者ID: ${adminUserId}`);
      console.log(`初期パスワード: ${adminPassword}`);
    }
    
  } catch (error) {
    console.error('❌ デフォルト管理者作成エラー:', error);
    throw error;
  }
}

/**
 * ユーザーID生成
 */
function generateUserId() {
  const timestamp = new Date().getTime();
  const randomNum = Math.floor(Math.random() * 1000);
  return `USER_${timestamp}_${randomNum}`;
}

/**
 * パスワードのハッシュ化（SHA-256 + ソルト）
 */
function hashPassword(password) {
  try {
    const salt = generateSalt();
    const saltedPassword = password + salt;
    const hash = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, saltedPassword);
    const hashString = hash.map(byte => (byte + 256).toString(16).substr(-2)).join('');
    return `${salt}:${hashString}`;
  } catch (error) {
    console.error('❌ パスワードハッシュ化エラー:', error);
    throw error;
  }
}

/**
 * ソルト生成
 */
function generateSalt() {
  const chars = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789';
  let salt = '';
  for (let i = 0; i < 16; i++) {
    salt += chars.charAt(Math.floor(Math.random() * chars.length));
  }
  return salt;
}

/**
 * パスワード検証
 */
function verifyPassword(password, hashedPassword) {
  try {
    const [salt, hash] = hashedPassword.split(':');
    const saltedPassword = password + salt;
    const computedHash = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, saltedPassword);
    const computedHashString = computedHash.map(byte => (byte + 256).toString(16).substr(-2)).join('');
    return computedHashString === hash;
  } catch (error) {
    console.error('❌ パスワード検証エラー:', error);
    return false;
  }
}

/**
 * ユーザー認証
 */
function authenticateUser(username, password) {
  try {
    console.log(`🔐 ユーザー認証開始: ${username}`);
    
    let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('ユーザー管理');
    if (!sheet) {
      console.log('ユーザー管理シートが存在しないため、自動初期化を実行...');
      sheet = initializeUserManagementSheet();
    }
    
    const data = sheet.getDataRange().getValues();
    
    // ヘッダー行をスキップして検索
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const [userId, userName, passwordHash, role, , , isActive] = row;
      
      if (userName === username && isActive === true) {
        if (verifyPassword(password, passwordHash)) {
          // 認証成功 - セッション開始
          startUserSession(userId, userName, role);
          updateLastLogin(sheet, i + 1); // 行番号は1ベース
          
          console.log(`✅ ユーザー認証成功: ${username} (${role})`);
          return {
            success: true,
            userId: userId,
            username: userName,
            role: role
          };
        }
      }
    }
    
    console.log(`❌ ユーザー認証失敗: ${username}`);
    return {
      success: false,
      message: 'ユーザー名またはパスワードが正しくありません'
    };
    
  } catch (error) {
    console.error('❌ ユーザー認証エラー:', error);
    return {
      success: false,
      message: `認証エラー: ${error.message}`
    };
  }
}

/**
 * セッション開始
 */
function startUserSession(userId, username, role) {
  try {
    const properties = PropertiesService.getScriptProperties();
    const loginTime = new Date().getTime();
    
    properties.setProperties({
      [SESSION_PROPERTIES.CURRENT_USER]: userId,
      [SESSION_PROPERTIES.USER_ROLE]: role,
      [SESSION_PROPERTIES.LOGIN_TIME]: loginTime.toString(),
      [SESSION_PROPERTIES.SESSION_TIMEOUT]: (loginTime + (SESSION_TIMEOUT_HOURS * 60 * 60 * 1000)).toString()
    });
    
    console.log(`✅ セッション開始: ${username} (${role})`);
    
    // シート可視性を更新
    try {
      if (typeof updateSheetVisibilityOnLogin === 'function') {
        updateSheetVisibilityOnLogin(role);
      }
    } catch (error) {
      console.log('シート可視性更新をスキップ:', error.message);
    }
    
  } catch (error) {
    console.error('❌ セッション開始エラー:', error);
    throw error;
  }
}

/**
 * 最終ログイン日時更新
 */
function updateLastLogin(sheet, rowIndex) {
  try {
    sheet.getRange(rowIndex, 6).setValue(new Date());
  } catch (error) {
    console.error('❌ 最終ログイン更新エラー:', error);
  }
}

/**
 * 現在のユーザー情報取得
 */
function getCurrentUser() {
  try {
    const properties = PropertiesService.getScriptProperties();
    const userId = properties.getProperty(SESSION_PROPERTIES.CURRENT_USER);
    const role = properties.getProperty(SESSION_PROPERTIES.USER_ROLE);
    const loginTime = properties.getProperty(SESSION_PROPERTIES.LOGIN_TIME);
    const sessionTimeout = properties.getProperty(SESSION_PROPERTIES.SESSION_TIMEOUT);
    
    if (!userId || !role) {
      return {
        isLoggedIn: false,
        role: USER_ROLES.GUEST
      };
    }
    
    // セッション有効性チェック
    const currentTime = new Date().getTime();
    if (sessionTimeout && currentTime > parseInt(sessionTimeout)) {
      // セッション期限切れ
      endUserSession();
      return {
        isLoggedIn: false,
        role: USER_ROLES.GUEST,
        message: 'セッションが期限切れです'
      };
    }
    
    // プラン管理統合ポイント - Phase 0
    let planPermissions = null;
    try {
      // plan-manager.jsが利用可能な場合はプラン権限も含める
      if (typeof getPlanDetails === 'function') {
        const planDetails = getPlanDetails();
        planPermissions = {
          planType: planDetails.planType,
          planDisplayName: planDetails.displayName,
          planLimits: planDetails.limits,
          isTemporary: planDetails.isTemporary
        };
      }
    } catch (error) {
      console.log('プラン管理システム未初期化 - ユーザー権限のみ提供');
    }

    return {
      isLoggedIn: true,
      userId: userId,
      role: role,
      loginTime: loginTime ? new Date(parseInt(loginTime)) : null,
      // Phase 0: プラン管理統合
      planPermissions: planPermissions
    };
    
  } catch (error) {
    console.error('❌ 現在ユーザー取得エラー:', error);
    return {
      isLoggedIn: false,
      role: USER_ROLES.GUEST
    };
  }
}

/**
 * 統合された有効権限を取得（ユーザー権限 + プラン権限）
 * Phase 0: 統合ポイント
 */
function getEffectivePermissions() {
  try {
    const currentUser = getCurrentUser();
    
    // ベース権限（ユーザーロールによる）
    const basePermissions = {
      canAccessAdminFeatures: currentUser.role === USER_ROLES.ADMINISTRATOR,
      canManageUsers: currentUser.role === USER_ROLES.ADMINISTRATOR,
      canViewSystemStats: currentUser.role !== USER_ROLES.GUEST,
      canUseBasicFeatures: true
    };
    
    // プラン制限を統合
    if (currentUser.planPermissions) {
      const planLimits = currentUser.planPermissions.planLimits;
      
      return {
        ...basePermissions,
        // プラン制限を反映
        canGenerateKeywords: planLimits.keywordGeneration,
        canUseAiProposals: planLimits.aiProposals,
        maxCompaniesPerDay: planLimits.maxCompaniesPerDay,
        requiresApiKey: planLimits.requiresApiKey,
        // プラン情報
        planType: currentUser.planPermissions.planType,
        planDisplayName: currentUser.planPermissions.planDisplayName,
        isTemporaryPlan: currentUser.planPermissions.isTemporary
      };
    }
    
    // プラン管理が未初期化の場合は基本権限のみ
    return {
      ...basePermissions,
      // 制限的なデフォルト
      canGenerateKeywords: false,
      canUseAiProposals: false,
      maxCompaniesPerDay: 10,
      requiresApiKey: false,
      planType: 'BASIC',
      planDisplayName: '🥉 ベーシック',
      isTemporaryPlan: false
    };
    
  } catch (error) {
    console.error('❌ 有効権限取得エラー:', error);
    return {
      canAccessAdminFeatures: false,
      canManageUsers: false,
      canViewSystemStats: false,
      canUseBasicFeatures: true,
      canGenerateKeywords: false,
      canUseAiProposals: false,
      maxCompaniesPerDay: 10,
      requiresApiKey: false,
      planType: 'BASIC',
      planDisplayName: '🥉 ベーシック',
      isTemporaryPlan: false
    };
  }
}

/**
 * セッション終了
 */
function endUserSession() {
  try {
    const properties = PropertiesService.getScriptProperties();
    properties.deleteProperty(SESSION_PROPERTIES.CURRENT_USER);
    properties.deleteProperty(SESSION_PROPERTIES.USER_ROLE);
    properties.deleteProperty(SESSION_PROPERTIES.LOGIN_TIME);
    properties.deleteProperty(SESSION_PROPERTIES.SESSION_TIMEOUT);
    
    console.log('✅ セッション終了');
    
    // ゲストモードにシート可視性を戻す
    try {
      if (typeof updateSheetVisibilityOnLogin === 'function') {
        updateSheetVisibilityOnLogin('Guest');
      }
    } catch (error) {
      console.log('シート可視性更新をスキップ:', error.message);
    }
    
  } catch (error) {
    console.error('❌ セッション終了エラー:', error);
  }
}

/**
 * 権限チェック
 */
function checkUserPermission(requiredRole) {
  try {
    const currentUser = getCurrentUser();
    
    if (!currentUser.isLoggedIn) {
      return {
        hasPermission: requiredRole === USER_ROLES.GUEST,
        currentRole: USER_ROLES.GUEST
      };
    }
    
    const roleHierarchy = {
      [USER_ROLES.GUEST]: 1,
      [USER_ROLES.STANDARD]: 2,
      [USER_ROLES.ADMINISTRATOR]: 3
    };
    
    const hasPermission = roleHierarchy[currentUser.role] >= roleHierarchy[requiredRole];
    
    return {
      hasPermission: hasPermission,
      currentRole: currentUser.role,
      userId: currentUser.userId
    };
    
  } catch (error) {
    console.error('❌ 権限チェックエラー:', error);
    return {
      hasPermission: false,
      currentRole: USER_ROLES.GUEST
    };
  }
}

/**
 * 新規ユーザー作成
 */
function createUser(username, password, role, createdBy) {
  try {
    console.log(`👤 新規ユーザー作成開始: ${username}`);
    
    // 管理者権限チェック
    const permission = checkUserPermission(USER_ROLES.ADMINISTRATOR);
    if (!permission.hasPermission) {
      throw new Error('ユーザー作成には管理者権限が必要です');
    }
    
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('ユーザー管理');
    if (!sheet) {
      throw new Error('ユーザー管理シートが見つかりません');
    }
    
    // ユーザー名重複チェック
    const data = sheet.getDataRange().getValues();
    const existingUser = data.find(row => row[1] === username);
    if (existingUser) {
      throw new Error('指定されたユーザー名は既に存在します');
    }
    
    // パスワード強度チェック
    if (!validatePassword(password)) {
      throw new Error('パスワードは8文字以上で、英数字と記号を含む必要があります');
    }
    
    // 新規ユーザー作成
    const userId = generateUserId();
    const passwordHash = hashPassword(password);
    
    const newUser = [
      userId,
      username,
      passwordHash,
      role,
      new Date(),
      '', // 最終ログイン（空）
      true, // アクティブ
      createdBy || 'システム管理者'
    ];
    
    sheet.appendRow(newUser);
    
    console.log(`✅ 新規ユーザー作成完了: ${username} (${role})`);
    return {
      success: true,
      userId: userId,
      message: `ユーザー「${username}」を作成しました`
    };
    
  } catch (error) {
    console.error('❌ ユーザー作成エラー:', error);
    return {
      success: false,
      message: error.message
    };
  }
}

/**
 * パスワード強度検証
 */
function validatePassword(password) {
  if (password.length < 8) return false;
  
  const hasUpperCase = /[A-Z]/.test(password);
  const hasLowerCase = /[a-z]/.test(password);
  const hasNumbers = /\d/.test(password);
  const hasSpecialChar = /[!@#$%^&*(),.?":{}|<>]/.test(password);
  
  return hasUpperCase && hasLowerCase && hasNumbers && hasSpecialChar;
}

/**
 * ユーザーリスト取得（管理者のみ）
 */
function getUserList() {
  try {
    // 管理者権限チェック
    const permission = checkUserPermission(USER_ROLES.ADMINISTRATOR);
    if (!permission.hasPermission) {
      throw new Error('ユーザーリスト取得には管理者権限が必要です');
    }
    
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('ユーザー管理');
    if (!sheet) {
      return [];
    }
    
    const data = sheet.getDataRange().getValues();
    const users = [];
    
    // ヘッダー行をスキップ
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      users.push({
        userId: row[0],
        username: row[1],
        role: row[3],
        createdDate: row[4],
        lastLogin: row[5],
        isActive: row[6],
        createdBy: row[7]
      });
    }
    
    return users;
    
  } catch (error) {
    console.error('❌ ユーザーリスト取得エラー:', error);
    return [];
  }
}

/**
 * ユーザーログイン UI
 */
function showUserLoginDialog() {
  try {
    const html = `
      <div style="padding: 20px; font-family: Arial, sans-serif;">
        <h2>🔐 ユーザーログイン</h2>
        <form>
          <div style="margin-bottom: 15px;">
            <label for="username" style="display: block; margin-bottom: 5px;">ユーザー名:</label>
            <input type="text" id="username" style="width: 100%; padding: 8px; border: 1px solid #ddd; border-radius: 4px;">
          </div>
          <div style="margin-bottom: 15px;">
            <label for="password" style="display: block; margin-bottom: 5px;">パスワード:</label>
            <input type="password" id="password" style="width: 100%; padding: 8px; border: 1px solid #ddd; border-radius: 4px;">
          </div>
          <div style="text-align: center;">
            <button type="button" onclick="performLogin()" style="background: #4285f4; color: white; padding: 10px 20px; border: none; border-radius: 4px; cursor: pointer;">ログイン</button>
            <button type="button" onclick="google.script.host.close()" style="background: #ccc; color: black; padding: 10px 20px; border: none; border-radius: 4px; cursor: pointer; margin-left: 10px;">キャンセル</button>
          </div>
        </form>
        <div id="message" style="margin-top: 15px; padding: 10px; display: none;"></div>
      </div>
      
      <script>
        function performLogin() {
          const username = document.getElementById('username').value;
          const password = document.getElementById('password').value;
          
          if (!username || !password) {
            showMessage('ユーザー名とパスワードを入力してください', 'error');
            return;
          }
          
          google.script.run
            .withSuccessHandler(onLoginSuccess)
            .withFailureHandler(onLoginFailure)
            .authenticateUser(username, password);
        }
        
        function onLoginSuccess(result) {
          if (result.success) {
            showMessage('ログイン成功！', 'success');
            setTimeout(() => {
              google.script.host.close();
            }, 1500);
          } else {
            showMessage(result.message || 'ログインに失敗しました', 'error');
          }
        }
        
        function onLoginFailure(error) {
          showMessage('ログイン処理中にエラーが発生しました', 'error');
        }
        
        function showMessage(text, type) {
          const messageDiv = document.getElementById('message');
          messageDiv.textContent = text;
          messageDiv.style.display = 'block';
          messageDiv.style.backgroundColor = type === 'success' ? '#d4edda' : '#f8d7da';
          messageDiv.style.color = type === 'success' ? '#155724' : '#721c24';
          messageDiv.style.border = '1px solid ' + (type === 'success' ? '#c3e6cb' : '#f5c6cb');
          messageDiv.style.borderRadius = '4px';
        }
      </script>
    `;
    
    const htmlOutput = HtmlService.createHtmlOutput(html)
      .setWidth(400)
      .setHeight(300);
    
    SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'ユーザーログイン');
    
  } catch (error) {
    console.error('❌ ログインダイアログ表示エラー:', error);
    SpreadsheetApp.getUi().alert('ログイン画面の表示に失敗しました');
  }
}

/**
 * 現在のユーザー状態表示
 */
function showCurrentUserStatus() {
  try {
    const currentUser = getCurrentUser();
    
    let statusMessage = '';
    if (currentUser.isLoggedIn) {
      statusMessage = `
        ✅ ログイン状態: ${currentUser.isLoggedIn}
        👤 ユーザーID: ${currentUser.userId}
        🔑 権限レベル: ${currentUser.role}
        🕒 ログイン時刻: ${currentUser.loginTime ? currentUser.loginTime.toLocaleString() : '不明'}
      `;
    } else {
      statusMessage = `
        ❌ ログイン状態: ログインしていません
        🔑 現在の権限: ゲストユーザー（閲覧のみ）
        
        ${currentUser.message || ''}
      `;
    }
    
    SpreadsheetApp.getUi().alert('ユーザー状態', statusMessage, SpreadsheetApp.getUi().ButtonSet.OK);
    
  } catch (error) {
    console.error('❌ ユーザー状態表示エラー:', error);
    SpreadsheetApp.getUi().alert('ユーザー状態の取得に失敗しました');
  }
}

/**
 * ログアウト
 */
function logoutUser() {
  try {
    const currentUser = getCurrentUser();
    
    if (currentUser.isLoggedIn) {
      endUserSession();
      SpreadsheetApp.getUi().alert('ログアウト', 'ログアウトしました', SpreadsheetApp.getUi().ButtonSet.OK);
    } else {
      SpreadsheetApp.getUi().alert('ログアウト', '現在ログインしていません', SpreadsheetApp.getUi().ButtonSet.OK);
    }
    
  } catch (error) {
    console.error('❌ ログアウトエラー:', error);
    SpreadsheetApp.getUi().alert('ログアウト処理中にエラーが発生しました');
  }
}

/**
 * シンプルプロンプトログイン（統一ログイン方法）
 */
function simpleLogin() {
  try {
    const ui = SpreadsheetApp.getUi();
    
    const usernameResponse = ui.prompt('🔐 ログイン 1/2', 'ユーザー名:', ui.ButtonSet.OK_CANCEL);
    if (usernameResponse.getSelectedButton() !== ui.Button.OK) return;
    const username = usernameResponse.getResponseText().trim();
    if (!username) { ui.alert('❌ エラー', 'ユーザー名を入力してください', ui.ButtonSet.OK); return; }
    
    const passwordResponse = ui.prompt('🔐 ログイン 2/2', 'パスワード:', ui.ButtonSet.OK_CANCEL);
    if (passwordResponse.getSelectedButton() !== ui.Button.OK) return;
    const password = passwordResponse.getResponseText();
    if (!password) { ui.alert('❌ エラー', 'パスワードを入力してください', ui.ButtonSet.OK); return; }
    
    const authResult = authenticateUser(username, password);
    
    if (authResult.success) {
      ui.alert('✅ ログイン成功', `ようこそ、${authResult.username}さん！\n権限: ${authResult.role}\n\nページを再読み込みしてください。`, ui.ButtonSet.OK);
    } else {
      ui.alert('❌ ログイン失敗', authResult.message || 'ユーザー名またはパスワードが正しくありません', ui.ButtonSet.OK);
    }
    
  } catch (error) {
    console.error('❌ ログインエラー:', error);
    SpreadsheetApp.getUi().alert('❌ エラー', `ログイン処理中にエラーが発生しました: ${error.message}`, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}