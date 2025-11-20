/**
 * Air Coin App - Server Side Logic
 * @fileoverview 大阪支社コミュニケーション活性化アプリ「エアコイン」のサーバーサイドコード
 */

// --- Constants & Configuration ---
// ★★★ ここを修正してください ★★★
// 例: const SS_ID = '1A2b3C4d5E6f...'; (クォーテーション '' を消さないように注意！)
const SS_ID = '1NoUx4-2gMT9I0bf3FYZ_zCfNAgymSykP-6IVxDkndM8'; 

const SHEET_NAMES = {
  USERS: 'Users',
  TRANSACTIONS: 'Transactions',
  CONFIG: 'Config'
};

const DEFAULT_CONFIG = {
  INITIAL_COIN: 100,
  COST_SAME_DEPT: 1,
  COST_DIFF_DEPT: 10,
  LIMIT_WEEKLY: 3,
  LIMIT_DAILY: 1
};

// --- Web App Entry Points ---

function doGet(e) {
  const template = HtmlService.createTemplateFromFile('index');
  // ID未設定やエラー時の安全策
  try {
    const economyState = analyzeEconomyState();
    template.theme = economyState;
  } catch (e) {
    template.theme = 'normal';
  }
  return template.evaluate()
    .setTitle('Air Coin - 大阪支社')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// --- Core Logic ---

/**
 * コイン送信処理
 */
function sendAirCoin(receiverEmail, comment) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000); 
  } catch (e) {
    return { success: false, message: 'サーバーが混み合っています。' };
  }

  try {
    checkSpreadsheetId();

    const ss = SpreadsheetApp.openById(SS_ID);
    const userSheet = getSheetOrDie(ss, SHEET_NAMES.USERS);
    const transSheet = getSheetOrDie(ss, SHEET_NAMES.TRANSACTIONS);
    
    const senderEmail = Session.getActiveUser().getEmail();
    
    if (!receiverEmail || !comment) throw new Error('宛先とコメントは必須です。');
    if (senderEmail === receiverEmail) throw new Error('自分自身には送れません。');
    
    // ユーザーデータを取得 (Object形式)
    const users = getDataAsObject(userSheet);
    const sender = users[senderEmail];
    const receiver = users[receiverEmail];
    
    if (!sender) throw new Error(`送信者のアカウント(${senderEmail})が見つかりません。Usersシートのuser_idを確認してください。`);
    if (!receiver) throw new Error(`受信者のアカウント(${receiverEmail})が見つかりません。Usersシートのuser_idを確認してください。`);
    
    // コスト計算
    const isSameDept = sender.department === receiver.department;
    const config = getConfig(ss);
    const cost = isSameDept ? Number(config.COST_SAME_DEPT) : Number(config.COST_DIFF_DEPT);
    const gain = 1; 
    
    if (Number(sender.wallet_balance) < cost) {
      throw new Error(`コイン不足です。必要: ${cost}枚 (残高: ${sender.wallet_balance}枚)`);
    }
    
    checkFrequencyLimit(transSheet, senderEmail, receiverEmail, config);
    
    const timestamp = new Date();
    const newTransId = Utilities.getUuid();
    
    // トランザクション記録
    transSheet.appendRow([
      newTransId,
      timestamp,
      senderEmail,
      receiverEmail,
      sender.department,
      receiver.department,
      cost,
      gain,
      comment
    ]);
    
    // ユーザー更新
    const userKeys = Object.keys(users);
    const senderRow = userKeys.indexOf(senderEmail) + 2; 
    const receiverRow = userKeys.indexOf(receiverEmail) + 2;
    
    // 送信者: wallet_balance (E列=5)
    const newSenderBalance = Number(sender.wallet_balance) - cost;
    userSheet.getRange(senderRow, 5).setValue(newSenderBalance);
    userSheet.getRange(senderRow, 8).setValue(timestamp); // last_updated (H列=8)
    
    // 受信者: lifetime_received (F列=6), rank (D列=4)
    const newReceiverLifetime = Number(receiver.lifetime_received) + gain;
    const newRank = calculateRank(newReceiverLifetime);
    userSheet.getRange(receiverRow, 6).setValue(newReceiverLifetime);
    userSheet.getRange(receiverRow, 4).setValue(newRank);
    
    return {
      success: true,
      message: `${receiver.name}さんにコインを送りました！（消費: ${cost}枚）`,
      newBalance: newSenderBalance
    };
    
  } catch (error) {
    console.error(`SendCoin Error: ${error.message}`);
    return { success: false, message: error.message };
  } finally {
    lock.releaseLock();
  }
}

function checkFrequencyLimit(sheet, senderId, receiverId, config) {
  const now = new Date();
  const data = sheet.getDataRange().getValues();
  data.shift(); 
  
  const interactions = data.filter(row => row[2] === senderId && row[3] === receiverId);
  if (interactions.length === 0) return;
  
  const todayStr = now.toDateString();
  const todayCount = interactions.filter(row => new Date(row[1]).toDateString() === todayStr).length;
  if (todayCount >= config.LIMIT_DAILY) {
    throw new Error(`同じ相手への送信は1日${config.LIMIT_DAILY}回までです。`);
  }
  
  const oneWeekAgo = new Date(now.getTime() - 7 * 24 * 60 * 60 * 1000);
  const weekCount = interactions.filter(row => new Date(row[1]) > oneWeekAgo).length;
  if (weekCount >= config.LIMIT_WEEKLY) {
    throw new Error(`同じ相手への送信は週${config.LIMIT_WEEKLY}回までです。`);
  }
}

// --- Helper Functions ---

function checkSpreadsheetId() {
  if (!SS_ID || SS_ID.includes('貼り付けて') || SS_ID.includes('YOUR_SPREADSHEET_ID')) {
    throw new Error('スプレッドシートIDが設定されていません。server.gsの SS_ID を書き換えてください。');
  }
}

/**
 * シートが存在するか確認し、なければわかりやすいエラーを出す
 */
function getSheetOrDie(ss, sheetName) {
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    throw new Error(`シート「${sheetName}」が見つかりません。スプレッドシート下部のタブ名を「${sheetName}」に変更してください。`);
  }
  return sheet;
}

function calculateRank(lifetimeCoins) {
  if (lifetimeCoins >= 1000) return '富豪';
  if (lifetimeCoins >= 500) return '一般';
  if (lifetimeCoins >= 100) return '見習い';
  return '新人';
}

function analyzeEconomyState() {
  try {
    if (!SS_ID || SS_ID.includes('貼り付けて')) return 'normal';

    const ss = SpreadsheetApp.openById(SS_ID);
    // ここはエラーを出さずにnormalを返すだけに留める
    const sheet = ss.getSheetByName(SHEET_NAMES.TRANSACTIONS);
    if (!sheet) return 'normal';

    const data = sheet.getDataRange().getValues();
    data.shift();
    
    const now = new Date();
    const sevenDaysAgo = new Date(now.getTime() - 7 * 24 * 60 * 60 * 1000);
    let totalVolume = 0;
    
    data.forEach(row => {
      const date = new Date(row[1]);
      if (date > sevenDaysAgo) {
        totalVolume += Number(row[6] || 0); 
      }
    });
    
    if (totalVolume > 2000) return 'boom';
    if (totalVolume < 500) return 'recession';
    return 'normal';
  } catch (e) {
    return 'normal';
  }
}

function getInitialData() {
  try {
    checkSpreadsheetId();

    const ss = SpreadsheetApp.openById(SS_ID);
    const userSheet = getSheetOrDie(ss, SHEET_NAMES.USERS);
    const email = Session.getActiveUser().getEmail();
    
    console.log(`Login User: ${email}`);

    const users = getDataAsObject(userSheet);
    const currentUser = users[email];
    
    if (!currentUser) {
      console.error('Registered Users:', Object.keys(users));
      // ユーザー一覧が空、またはヘッダーしかない場合のチェック
      if (Object.keys(users).length === 0) {
         throw new Error(`Usersシートにデータがありません。A列に ${email} を追加してください。`);
      }
      throw new Error(`あなたのメールアドレス(${email})がUsersシートに登録されていません。`);
    }
    
    const userList = Object.values(users)
      .filter(u => u.user_id !== email)
      .map(u => ({ 
        email: u.user_id, 
        name: u.name, 
        department: u.department 
      }));
      
    return {
      user: currentUser,
      userList: userList,
      economy: analyzeEconomyState()
    };
  } catch (e) {
    console.error(e);
    throw e;
  }
}

function getConfig(ss) {
  // Configシートだけは無くても動くようにフォールバックする
  const sheet = ss.getSheetByName(SHEET_NAMES.CONFIG);
  if (!sheet) return DEFAULT_CONFIG;

  const data = sheet.getDataRange().getValues();
  data.shift();
  const config = {};
  data.forEach(row => {
    config[row[0]] = row[1];
  });
  return { ...DEFAULT_CONFIG, ...config };
}

function getDataAsObject(sheet) {
  const data = sheet.getDataRange().getValues();
  // データが空（ヘッダーすらない）場合
  if (data.length === 0) return {};

  const header = data.shift();
  const result = {};
  
  data.forEach(row => {
    const obj = {};
    header.forEach((h, i) => {
      obj[h] = row[i];
    });
    const key = row[0]; 
    if(key) {
      result[key] = obj;
    }
  });
  
  return result;
}

function initializeMonthlyCoin() {
  if (!SS_ID || SS_ID.includes('貼り付けて')) return;
  
  const ss = SpreadsheetApp.openById(SS_ID);
  // バッチ処理なのでエラーログを残す
  const userSheet = ss.getSheetByName(SHEET_NAMES.USERS);
  if (!userSheet) {
    console.error("Monthly Reset Failed: Users sheet not found.");
    return;
  }
  
  const transSheet = ss.getSheetByName(SHEET_NAMES.TRANSACTIONS);
  const config = getConfig(ss);
  
  const usersData = userSheet.getDataRange().getValues();
  const header = usersData.shift();
  
  const now = new Date();
  const lastMonth = new Date(now.getFullYear(), now.getMonth() - 1, 1);
  
  // Transactionsが無い場合は履歴チェックをスキップして全員リセット
  const activeSenders = new Set();
  if (transSheet) {
    const transData = transSheet.getDataRange().getValues();
    transData.forEach((row, i) => {
      if (i === 0) return;
      const date = new Date(row[1]);
      if (date.getFullYear() === lastMonth.getFullYear() && date.getMonth() === lastMonth.getMonth()) {
        activeSenders.add(row[2]);
      }
    });
  }
  
  const updates = usersData.map(user => {
    const userId = user[0];
    const isInactive = !activeSenders.has(userId);
    const penaltyFlag = isInactive ? 'INACTIVE_LAST_MONTH' : '';
    const newBalance = config.INITIAL_COIN;
    return [newBalance, penaltyFlag];
  });
  
  if (updates.length > 0) {
    userSheet.getRange(2, 5, updates.length, 1).setValues(updates.map(u => [u[0]]));
    userSheet.getRange(2, 7, updates.length, 1).setValues(updates.map(u => [u[1]]));
  }
}
