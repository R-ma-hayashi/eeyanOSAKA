/**
 * E-yan Coin App - 大阪支社 (高速化対応版 v2.3)
 * 仕様書v2対応 + パフォーマンスチューニング
 */

// --- Configuration ---
// ★★★ スプレッドシートID ★★★
const SS_ID = '1E0qf3XM-W8TM5HZ_SrPPoGAV4kwObvS6FmQdaFR3Bpw';

const SHEET_NAMES = {
  USERS: 'Users',
  TRANSACTIONS: 'Transactions',
  CONFIG: 'Config',
  DEPARTMENTS: 'Departments',
  ARCHIVE: 'Archive_Log',
  MVP_HISTORY: 'MVP_History'
};

const CACHE_DURATION = {
  USER_LIST: 30,
  ECONOMY: 600,
  DEPARTMENTS: 3600,
  RANKINGS: 1800 // ランキングは30分キャッシュ
};

const DEFAULT_CONFIG = {
  INITIAL_COIN: 1000,
  MULTIPLIER_DIFF_DEPT: 10,
  MESSAGE_MAX_LENGTH: 100,
  ECONOMY_THRESHOLD_L2: 10000,
  ECONOMY_THRESHOLD_L3: 50000
};

// --- Web App Entry Points ---

function doGet(e) {
  const template = HtmlService.createTemplateFromFile('index');
  return template.evaluate()
    .setTitle('E-yan Coin - 大阪支社')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// --- Core Logic : 初期データ取得 ---

function getInitialData() {
  const email = Session.getActiveUser().getEmail();
  try {
    const cache = CacheService.getScriptCache();
    
    // ユーザー一覧取得 (キャッシュ活用)
    let usersMap;
    const cachedUsers = cache.get('ALL_USERS_DATA');
    if (cachedUsers) {
      usersMap = JSON.parse(cachedUsers);
    } else {
      usersMap = fetchAndCacheUsersData();
    }

    const departments = getDepartmentsCached();
    const currentUser = usersMap[email];

    if (!currentUser) {
      return {
        error: 'NOT_REGISTERED',
        email: email,
        departments: departments
      };
    }

    // 前月MVP情報を取得
    const lastMonthMVP = getLastMonthMVP();
    currentUser.isMVP = (lastMonthMVP === email);

    // フロントエンド用ユーザーリスト (軽量化のため必要な情報のみ)
    const userList = Object.values(usersMap)
      .filter(u => u.user_id !== email)
      .map(u => ({
        email: u.user_id,
        name: u.name,
        department: u.department
      }));

    // 景気状態
    let economyState = cache.get('ECONOMY_STATE');
    if (!economyState) {
      economyState = analyzeEconomyState();
      cache.put('ECONOMY_STATE', economyState, CACHE_DURATION.ECONOMY);
    }

    return {
      success: true,
      user: currentUser,
      userList: userList,
      economy: economyState,
      departments: departments,
      config: getConfigCached()
    };

  } catch (e) {
    console.error('Error:', e);
    throw new Error('データ読み込みエラー: ' + e.message);
  }
}

// --- Core Logic : 送金処理 ---

function sendAirCoin(receiverEmail, comment, amountInput) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) {
    return { success: false, message: '現在混み合っています。少し待って再送してください。' };
  }

  try {
    const config = getConfigCached();
    const amount = Number(amountInput); 

    if (!Number.isInteger(amount) || amount <= 0) throw new Error('コイン枚数は1以上の整数で指定してください。');
    if (comment.length > config.MESSAGE_MAX_LENGTH) throw new Error(`メッセージは${config.MESSAGE_MAX_LENGTH}文字以内で入力してください。`);

    const ss = getSpreadsheet();
    const userSheet = ss.getSheetByName(SHEET_NAMES.USERS);
    const transSheet = ss.getSheetByName(SHEET_NAMES.TRANSACTIONS);
    const senderEmail = Session.getActiveUser().getEmail();

    // ユーザーデータ取得
    const usersData = userSheet.getDataRange().getValues();
    const header = usersData.shift();
    const emailToRow = {};
    const usersObj = {};

    usersData.forEach((row, i) => {
      const email = row[0]; 
      emailToRow[email] = i + 2; 
      const u = {};
      header.forEach((h, idx) => u[h] = row[idx]);
      usersObj[email] = u;
    });

    const sender = usersObj[senderEmail];
    const receiver = usersObj[receiverEmail];

    if (!sender || !receiver) throw new Error('ユーザーデータが見つかりません。');
    if (senderEmail === receiverEmail) throw new Error('自分には送れません。');

    // 計算
    const costSender = amount; 
    if (Number(sender.wallet_balance) < costSender) throw new Error(`コイン不足です (保有: ${sender.wallet_balance})`);

    const isSameDept = sender.department === receiver.department;
    const multiplier = isSameDept ? 1 : Number(config.MULTIPLIER_DIFF_DEPT || 10);
    const valueGained = amount * multiplier; 

    const now = new Date();
    
    // トランザクション記録
    transSheet.appendRow([
      Utilities.getUuid(),
      now,
      senderEmail,
      receiverEmail,
      sender.department,
      receiver.department,
      amount,
      multiplier,
      costSender,
      valueGained,
      comment
    ]);

    // ユーザー更新
    const sRow = emailToRow[senderEmail];
    const rRow = emailToRow[receiverEmail];
    
    const newBal = Number(sender.wallet_balance) - costSender;
    const newLifetime = Number(receiver.lifetime_received) + valueGained;
    
    userSheet.getRange(sRow, 5).setValue(newBal);      // E列: wallet
    userSheet.getRange(sRow, 8).setValue(now);         // H列: last_updated
    userSheet.getRange(rRow, 6).setValue(newLifetime); // F列: lifetime

    // ランク判定（仕様書v3対応: RPG風ランク）
    let newRank = receiver.rank;
    if (newLifetime >= 10000) newRank = '天下人';
    else if (newLifetime >= 5000) newRank = '豪商';
    else if (newLifetime >= 1000) newRank = '商人';
    else if (newLifetime >= 100) newRank = '丁稚';
    else newRank = '素浪人';
    
    if (newRank !== receiver.rank) {
       userSheet.getRange(rRow, 4).setValue(newRank); // D列
    }

    // キャッシュクリア
    clearAllCaches();

    // 送信者と受信者の履歴キャッシュをクリア
    const cache = CacheService.getScriptCache();
    cache.remove('HISTORY_' + senderEmail);
    cache.remove('HISTORY_' + receiverEmail);

    return {
      success: true,
      message: isSameDept ? `${amount}枚送りました！` : `他部署ボーナス！相手に${valueGained}枚分の価値として届きました！`,
      newBalance: newBal,
      newLifetime: newLifetime,
      gainedValue: valueGained
    };

  } catch (e) {
    return { success: false, message: e.message };
  } finally {
    lock.releaseLock();
  }
}

// --- High Performance Logic : ランキング取得 (高速化版) ---

function getRankings() {
  const cache = CacheService.getScriptCache();
  const cached = cache.get('RANKINGS_v2_');
  if (cached) return { success: true, rankings: JSON.parse(cached) };

  const users = fetchAndCacheUsersData();
  const ss = getSpreadsheet();
  const transSheet = ss.getSheetByName(SHEET_NAMES.TRANSACTIONS);
  
  // 1. MVP (受信) ランキング
  // ★高速化: Transactionsを集計せず、Usersシートの lifetime_received を使う
  const mvpRanking = Object.values(users)
    .sort((a, b) => Number(b.lifetime_received) - Number(a.lifetime_received))
    .slice(0, 10)
    .map(u => ({
      email: u.user_id,
      name: u.name,
      score: Number(u.lifetime_received),
      dept: u.department
    }));

  // 2. 送信ランキング & 部署ランキング
  // ★高速化: Transactionsを「今月分」だけ「後ろから」走査する
  const now = new Date();
  const firstDay = new Date(now.getFullYear(), now.getMonth(), 1);
  
  const sentMap = {};
  const deptScoreMap = {};

  // データ取得 (getLastRowを使って効率的に)
  const lastRow = transSheet.getLastRow();
  const startRow = Math.max(2, lastRow - 2000); // 最大でも直近2000件見れば十分と判断
  
  if (lastRow >= 2) {
    const data = transSheet.getRange(startRow, 1, lastRow - startRow + 1, 10).getValues();
    
    // 後ろからループ (最新→過去)
    for (let i = data.length - 1; i >= 0; i--) {
      const row = data[i];
      const date = new Date(row[1]); // B列: timestamp
      
      if (date < firstDay) break; // 先月以前のデータが出たら終了（高速化の肝）
      
      const sender = row[2];        // C列: sender
      const receiverDept = row[5];  // F列: receiver_dept
      const val = Number(row[9]);   // J列: value_gained

      // 送信回数集計
      sentMap[sender] = (sentMap[sender] || 0) + 1;
      
      // 部署スコア集計
      if (receiverDept) {
        deptScoreMap[receiverDept] = (deptScoreMap[receiverDept] || 0) + val;
      }
    }
  }

  const giverRanking = Object.keys(sentMap)
    .map(email => ({
      email: email,
      name: users[email] ? users[email].name : '不明',
      count: sentMap[email]
    }))
    .sort((a, b) => b.count - a.count)
    .slice(0, 10);

  const deptRanking = Object.keys(deptScoreMap)
    .map(dept => ({
      name: dept,
      score: deptScoreMap[dept]
    }))
    .sort((a, b) => b.score - a.score);

  const rankings = { mvp: mvpRanking, giver: giverRanking, dept: deptRanking };
  cache.put('RANKINGS', JSON.stringify(rankings), CACHE_DURATION.RANKINGS);

  return { success: true, rankings: rankings };
}

// --- High Performance Logic : 履歴取得 (超高速版) ---

function getUserHistory(limit = 20) {
  const email = Session.getActiveUser().getEmail();
  const cache = CacheService.getScriptCache();

  // キャッシュチェック（5分間）
  const cacheKey = 'HISTORY_v2_' + email;
  const cached = cache.get(cacheKey);
  if (cached) {
    return { success: true, history: JSON.parse(cached) };
  }

  const ss = getSpreadsheet();
  const transSheet = ss.getSheetByName(SHEET_NAMES.TRANSACTIONS);
  const lastRow = transSheet.getLastRow();

  // 履歴がない場合
  if (lastRow < 2) return { success: true, history: [] };

  // ★超高速化: チャンクサイズを小さくし、必要な列のみ取得
  const CHUNK_SIZE = 100; // 500 → 100に削減
  const myHistory = [];
  const senderIds = new Set(); // 送信者IDを収集

  let currentRow = lastRow;

  while (currentRow >= 2 && myHistory.length < limit) {
    const startRow = Math.max(2, currentRow - CHUNK_SIZE + 1);
    const numRows = currentRow - startRow + 1;

    // ★必要な列のみ取得（B, C, D, E, G, J, K列 = 7列）
    // B=timestamp, C=sender, D=receiver, E=sender_dept, G=amount, J=value, K=message
    const data = transSheet.getRange(startRow, 2, numRows, 10).getValues();

    // 取得したデータを後ろから走査
    for (let i = data.length - 1; i >= 0; i--) {
      const row = data[i];
      // D列(receiver) = index 2 (範囲がB列から始まるため)
      if (row[2] === email) {
        const senderId = row[1]; // C列(sender) = index 1
        senderIds.add(senderId);

        myHistory.push({
          timestamp: row[0],    // B列
          sender_id: senderId,  // C列
          sender_dept: row[3],  // E列
          amount: row[5],       // G列
          value: row[8],        // J列
          message: row[9]       // K列
        });

        if (myHistory.length >= limit) break;
      }
    }

    currentRow -= CHUNK_SIZE;
    // 安全策: 過去2000件まで（5000 → 2000に削減）
    if (lastRow - currentRow > 2000) break;
  }

  // ★ユーザー名解決の最適化: 必要な送信者のみ取得
  if (myHistory.length > 0) {
    const users = fetchAndCacheUsersData();
    myHistory.forEach(h => {
      h.sender_name = users[h.sender_id] ? users[h.sender_id].name : '退職済ユーザー';
    });
  }

  // キャッシュに保存（5分間）
  cache.put(cacheKey, JSON.stringify(myHistory), 300);

  return { success: true, history: myHistory };
}

// --- サジェスト: 最近送った人 ---
function getRecentRecipients(limit = 5) {
  const email = Session.getActiveUser().getEmail();
  const ss = getSpreadsheet();
  const transSheet = ss.getSheetByName(SHEET_NAMES.TRANSACTIONS);
  const lastRow = transSheet.getLastRow();

  if (lastRow < 2) return { success: true, recipients: [] };

  const recentRecipients = [];
  const seenEmails = new Set();

  // 後ろから走査して、最近送った人を重複なしで取得
  let currentRow = lastRow;
  const CHUNK_SIZE = 200;

  while (currentRow >= 2 && recentRecipients.length < limit) {
    const startRow = Math.max(2, currentRow - CHUNK_SIZE + 1);
    const numRows = currentRow - startRow + 1;

    const data = transSheet.getRange(startRow, 1, numRows, 4).getValues(); // A~D列

    for (let i = data.length - 1; i >= 0; i--) {
      const row = data[i];
      const sender = row[2]; // C列: sender
      const receiver = row[3]; // D列: receiver

      if (sender === email && !seenEmails.has(receiver)) {
        seenEmails.add(receiver);
        recentRecipients.push(receiver);

        if (recentRecipients.length >= limit) break;
      }
    }

    currentRow -= CHUNK_SIZE;
    if (lastRow - currentRow > 2000) break; // 過去2000件まで
  }

  return { success: true, recipients: recentRecipients };
}

// --- 景気分析 (高速化版) ---
function analyzeEconomyState() {
  const ss = getSpreadsheet();
  const transSheet = ss.getSheetByName(SHEET_NAMES.TRANSACTIONS);
  if (!transSheet) return 'normal';
  
  // ここも「直近のデータ」だけ見れば良いので、getLastRowを活用
  const lastRow = transSheet.getLastRow();
  if (lastRow < 2) return 'normal';
  
  const now = new Date();
  const firstDay = new Date(now.getFullYear(), now.getMonth(), 1);
  
  // 最大1000件取得すれば、月初からのデータとしては十分と仮定
  // (厳密にやるならループが必要だが、景気判定なので概算でOK)
  const startRow = Math.max(2, lastRow - 1000);
  const data = transSheet.getRange(startRow, 2, lastRow - startRow + 1, 9).getValues(); // B列(date)〜J列(value)
  
  let totalValueGained = 0;
  for (let i = 0; i < data.length; i++) {
    // B列はindex 0 (getRangeで2列目から取ったので)
    const rowDate = new Date(data[i][0]);
    if (rowDate >= firstDay) {
      // J列はindex 8
      totalValueGained += Number(data[i][8] || 0);
    }
  }

  const config = getConfigCached();
  const thresholdL3 = Number(config.ECONOMY_THRESHOLD_L3 || 50000);
  const thresholdL2 = Number(config.ECONOMY_THRESHOLD_L2 || 10000);

  if (totalValueGained >= thresholdL3) return 'boom'; 
  if (totalValueGained >= thresholdL2) return 'normal'; 
  return 'depression'; 
}


// --- 定期実行: 月次リセット ---
function resetMonthlyData() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) return;

  try {
    const ss = getSpreadsheet();
    const userSheet = ss.getSheetByName(SHEET_NAMES.USERS);
    const config = getConfigCached();
    const defaultCoin = config.INITIAL_COIN || 1000;

    const data = userSheet.getDataRange().getValues();
    const header = data.shift();

    // MVP（今月最高獲得者）を記録
    if (data.length > 0) {
      let mvpEmail = '';
      let maxLifetime = 0;
      data.forEach(row => {
        const lifetime = Number(row[5]); // F列: lifetime_received
        if (lifetime > maxLifetime) {
          maxLifetime = lifetime;
          mvpEmail = row[0]; // A列: user_id
        }
      });

      if (mvpEmail) {
        saveMVPHistory(mvpEmail, maxLifetime);
      }
    }

    const walletColValues = [];
    const lifetimeColValues = [];
    const rankColValues = [];

    data.forEach(() => {
      rankColValues.push(['素浪人']);
      walletColValues.push([defaultCoin]);
      lifetimeColValues.push([0]);
    });

    if (data.length > 0) {
      userSheet.getRange(2, 4, data.length, 1).setValues(rankColValues);
      userSheet.getRange(2, 5, data.length, 1).setValues(walletColValues);
      userSheet.getRange(2, 6, data.length, 1).setValues(lifetimeColValues);
    }

    let archiveSheet = ss.getSheetByName(SHEET_NAMES.ARCHIVE);
    if (!archiveSheet) {
      archiveSheet = ss.insertSheet(SHEET_NAMES.ARCHIVE);
      archiveSheet.appendRow(['Date', 'Action', 'Count']);
    }
    archiveSheet.appendRow([new Date(), 'Monthly Reset Completed', data.length + ' users reset']);

    clearAllCaches();

  } catch (e) {
    console.error('月次リセットエラー:', e);
  } finally {
    lock.releaseLock();
  }
}

// --- MVP履歴管理 ---
function saveMVPHistory(email, score) {
  const ss = getSpreadsheet();
  let mvpSheet = ss.getSheetByName(SHEET_NAMES.MVP_HISTORY);
  if (!mvpSheet) {
    mvpSheet = ss.insertSheet(SHEET_NAMES.MVP_HISTORY);
    mvpSheet.appendRow(['YearMonth', 'Email', 'Score', 'Timestamp']);
  }

  const now = new Date();
  const yearMonth = Utilities.formatDate(now, 'Asia/Tokyo', 'yyyy-MM');
  mvpSheet.appendRow([yearMonth, email, score, now]);
}

function getLastMonthMVP() {
  const ss = getSpreadsheet();
  const mvpSheet = ss.getSheetByName(SHEET_NAMES.MVP_HISTORY);
  if (!mvpSheet) return null;

  const lastRow = mvpSheet.getLastRow();
  if (lastRow < 2) return null;

  // 最新のMVPを取得（最後の行）
  const mvpEmail = mvpSheet.getRange(lastRow, 2).getValue();
  return mvpEmail || null;
}

// --- 定期実行: リマインド ---
function sendReminderEmails() {
  const usersMap = fetchAndCacheUsersData();
  const users = Object.values(usersMap);
  
  users.forEach(user => {
    const balance = Number(user.wallet_balance);
    if (balance >= 300) {
      try {
        GmailApp.sendEmail(
          user.user_id,
          "【E-yan Coin】コインの有効期限が迫っています",
          `${user.name}さん\n\n今月のコイン残高: ${balance}枚\n月末にリセットされます。感謝を伝えましょう！`
        );
        Utilities.sleep(100);
      } catch (e) { console.error(e); }
    }
  });
}

// --- Helpers ---

function registerNewUser(formObject) {
  const lock = LockService.getScriptLock();
  try {
    if (lock.tryLock(10000)) {
      const ss = getSpreadsheet();
      const userSheet = ss.getSheetByName(SHEET_NAMES.USERS);
      const email = Session.getActiveUser().getEmail();
      const usersMap = fetchAndCacheUsersData();
      
      if (usersMap[email]) return { success: false, message: '既に登録済みです。' };

      const config = getConfigCached();
      userSheet.appendRow([
        email, formObject.name, formObject.department, '素浪人',
        config.INITIAL_COIN || 1000, 0, '', new Date()
      ]);
      clearAllCaches();
      return { success: true, message: '登録完了！' };
    } else {
      return { success: false, message: '混雑しています。' };
    }
  } catch (e) {
    return { success: false, message: '登録エラー: ' + e.message };
  } finally {
    lock.releaseLock();
  }
}

function getSpreadsheet() { return SpreadsheetApp.openById(SS_ID); }

function fetchAndCacheUsersData() {
  const ss = getSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAMES.USERS);
  const data = sheet.getDataRange().getValues();
  const header = data.shift();
  const map = {};
  data.forEach(row => {
    const u = {};
    header.forEach((h, i) => u[h] = row[i]);
    if(u.user_id) map[u.user_id] = u;
  });
  CacheService.getScriptCache().put('ALL_USERS_DATA', JSON.stringify(map), CACHE_DURATION.USER_LIST);
  return map;
}

function getDepartmentsCached() {
  const cache = CacheService.getScriptCache();
  const cached = cache.get('DEPT_LIST');
  if (cached) return JSON.parse(cached);
  const ss = getSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAMES.DEPARTMENTS);
  if (!sheet) return [];
  const vals = sheet.getRange(2, 1, sheet.getLastRow()-1 || 1, 1).getValues();
  const list = vals.flat().filter(String);
  cache.put('DEPT_LIST', JSON.stringify(list), CACHE_DURATION.DEPARTMENTS);
  return list;
}

function getConfigCached() {
  const ss = getSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAMES.CONFIG);
  let conf = {...DEFAULT_CONFIG};
  if(sheet) {
    const data = sheet.getDataRange().getValues();
    data.shift(); 
    data.forEach(r => { if(r[0]) conf[r[0]] = r[1]; });
  }
  return conf;
}

function clearAllCaches() {
  const cache = CacheService.getScriptCache();
  cache.remove('ALL_USERS_DATA');
  cache.remove('ECONOMY_STATE');
  cache.remove('DEPT_LIST');
  cache.remove('RANKINGS');
  // 履歴キャッシュは個別にクリアする必要がある（ユーザーごとに異なるため）
  // 送金時に送信者と受信者の履歴キャッシュをクリア
}
