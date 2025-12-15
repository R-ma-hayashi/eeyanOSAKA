/**
 * E-yan Coin App - 大阪支社 (v4.7 Performance Fix)
 * Update: Reverted History Logic to v4.4 (Chunk) for Speed, Kept v4.6 Ranking Logic
 */

// --- ★設定エリア (Config) ---
const APP_CONFIG = {
  INITIAL_COIN: 100,           // 月初の所持コイン
  MULTIPLIER_DIFF_DEPT: 1.2,   // 他部署倍率
  MESSAGE_MAX_LENGTH: 100,     // メッセージ文字数上限
  ECONOMY_THRESHOLD_L2: 6500,  // 景気Lv2閾値
  ECONOMY_THRESHOLD_L3: 13500, // 景気Lv3閾値
  REMINDER_THRESHOLD: 50,      // リマインド閾値
  
  // ID設定
  SS_ID: '1E0qf3XM-W8TM5HZ_SrPPoGAV4kwObvS6FmQdaFR3Bpw', // メインSS
  ARCHIVE_SS_ID: '1Gk3B_yd0q-sqskmQwHBsWk0PfYbSqfD0UdzYYiMhN5w', // アーカイブSS (指定ID)
  
  // ★手順で生成されたJSONファイルIDをここに貼る
  JSON_FILE_ID: '1K-9jVyC8SK9_g8AS87WxuI1Ax_IeiX7X' 
};

// シート名定義
const SHEET_NAMES = {
  USERS: 'Users',
  TRANSACTIONS: 'Transactions',
  DEPARTMENTS: 'Departments',
  ARCHIVE_LOG: 'Archive_Log',
  MVP_HISTORY: 'MVP_History'
};

// --- Web App Entry Points ---

function doGet(e) {
  const template = HtmlService.createTemplateFromFile('index');
  return template.evaluate()
    .setTitle('E-yan Coin - 大阪支社')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// --- API 1: 起動直後の軽量データ取得 ---
function getInitialData() {
  const email = Session.getActiveUser().getEmail();
  try {
    const ss = getSpreadsheet();
    const userSheet = ss.getSheetByName(SHEET_NAMES.USERS);
    
    // ユーザー検索
    const data = userSheet.getDataRange().getValues();
    const header = data.shift();
    const colIdx = {};
    header.forEach((h, i) => colIdx[h] = i);
    
    let myRow = null;
    for(let i=0; i<data.length; i++) {
      if(data[i][colIdx.user_id] === email) {
        myRow = data[i];
        break;
      }
    }

    if (!myRow) {
      return {
        error: 'NOT_REGISTERED',
        departments: getDepartmentsCached()
      };
    }

    const currentUser = {
      user_id: myRow[colIdx.user_id],
      name: myRow[colIdx.name],
      department: myRow[colIdx.department],
      rank: myRow[colIdx.rank],
      wallet_balance: myRow[colIdx.wallet_balance],
      lifetime_received: myRow[colIdx.lifetime_received],
      memo: myRow[colIdx.memo]
    };

    const lastMonthMVP = getLastMonthMVP();
    currentUser.isMVP = (lastMonthMVP === email);

    let dailySent = 0;
    try {
      if (currentUser.memo) {
        const memoObj = JSON.parse(currentUser.memo);
        const todayStr = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd');
        if (memoObj.last_sent_date === todayStr) {
          dailySent = memoObj.daily_total || 0;
        }
      }
    } catch (e) {}
    currentUser.dailySent = dailySent;

    // 景気状態
    const cache = CacheService.getScriptCache();
    let economyState = cache.get('ECONOMY_STATE_v4');
    if (!economyState) {
      economyState = analyzeEconomyState();
      cache.put('ECONOMY_STATE_v4', economyState, 600);
    }
    
    let dataVersion = new Date().getTime().toString();
    if (APP_CONFIG.JSON_FILE_ID) {
      try {
        const file = DriveApp.getFileById(APP_CONFIG.JSON_FILE_ID);
        dataVersion = file.getLastUpdated().getTime().toString();
      } catch(e) { console.warn('JSON File access error', e); }
    }

    return {
      success: true,
      user: currentUser,
      economy: economyState,
      config: APP_CONFIG,
      dataVersion: dataVersion
    };

  } catch (e) {
    console.error('Error:', e);
    throw new Error('起動エラー: ' + e.message);
  }
}

// --- API 2: ユーザーリスト取得 ---
function getUserListData() {
  try {
    let usersArray = [];
    if (APP_CONFIG.JSON_FILE_ID) {
      try {
        const file = DriveApp.getFileById(APP_CONFIG.JSON_FILE_ID);
        const content = file.getBlob().getDataAsString();
        usersArray = JSON.parse(content);
        return { success: true, list: usersArray, from: 'Drive' };
      } catch(e) {
        console.error('JSON read failed', e);
      }
    }
    const ss = getSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_NAMES.USERS);
    const data = sheet.getDataRange().getValues();
    data.shift(); 
    usersArray = data.map(row => [row[0], row[1], row[2]]);
    return { success: true, list: usersArray, from: 'Sheet' };
  } catch(e) {
    return { success: false, message: e.message };
  }
}

// --- 管理用: JSON手動更新 ---
function admin_updateUserJson() {
  const ss = getSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAMES.USERS);
  const data = sheet.getDataRange().getValues();
  data.shift(); 
  const simpleList = data.map(row => [row[0], row[1], row[2]]);
  const jsonString = JSON.stringify(simpleList);
  
  let file;
  if (APP_CONFIG.JSON_FILE_ID) {
    try {
      file = DriveApp.getFileById(APP_CONFIG.JSON_FILE_ID);
      file.setContent(jsonString);
    } catch(e) {
      file = DriveApp.createFile('EyanCoin_UserList.json', jsonString, MimeType.PLAIN_TEXT);
    }
  } else {
    file = DriveApp.createFile('EyanCoin_UserList.json', jsonString, MimeType.PLAIN_TEXT);
  }
  return file.getId();
}

// --- Core Logic : 送金処理 ---
function sendAirCoin(receiverEmail, comment, amountInput, shareFlag) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) return { success: false, message: '混雑中。再試行してください。' };

  try {
    const amount = Number(amountInput); 
    if (amount > 10) throw new Error('1回10枚までです。');

    const ss = getSpreadsheet();
    const userSheet = ss.getSheetByName(SHEET_NAMES.USERS);
    const transSheet = ss.getSheetByName(SHEET_NAMES.TRANSACTIONS);
    const senderEmail = Session.getActiveUser().getEmail();

    const data = userSheet.getDataRange().getValues();
    const header = data.shift();
    const colIdx = {};
    header.forEach((h, i) => colIdx[h] = i);
    
    let senderRow = -1, receiverRow = -1;
    let senderData, receiverData;

    for (let i = 0; i < data.length; i++) {
      if (data[i][colIdx.user_id] === senderEmail) { senderRow = i; senderData = data[i]; }
      if (data[i][colIdx.user_id] === receiverEmail) { receiverRow = i; receiverData = data[i]; }
    }

    if (senderRow === -1 || receiverRow === -1) throw new Error('ユーザーが見つかりません');

    const memoJsonStr = senderData[colIdx.memo] || "{}";
    let memoObj = {};
    try { memoObj = JSON.parse(memoJsonStr); } catch(e) {}

    const todayStr = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd');
    if (memoObj.last_sent_date !== todayStr) {
      memoObj.last_sent_date = todayStr;
      memoObj.daily_total = 0;
    }
    if (!memoObj.monthly_log) memoObj.monthly_log = {};

    if ((memoObj.daily_total + amount) > 20) throw new Error(`1日上限(20枚)を超えます。`);
    const currentTargetCount = memoObj.monthly_log[receiverEmail] || 0;
    if ((currentTargetCount + amount) > 30) throw new Error(`この人への月間上限(30枚)を超えます。`);
    
    const currentBalance = Number(senderData[colIdx.wallet_balance]);
    if (currentBalance < amount) throw new Error('コイン不足');

    const isSameDept = senderData[colIdx.department] === receiverData[colIdx.department];
    const multiplier = isSameDept ? 1 : Number(APP_CONFIG.MULTIPLIER_DIFF_DEPT);
    const valueGained = Math.floor(amount * multiplier);

    const newBal = currentBalance - amount;
    const newLife = Number(receiverData[colIdx.lifetime_received]) + valueGained;
    
    memoObj.daily_total += amount;
    memoObj.monthly_log[receiverEmail] = currentTargetCount + amount;

    // ランク判定ロジック
    let newRank = receiverData[colIdx.rank];
    if (newLife >= 120) newRank = '天下人';
    else if (newLife >= 90) newRank = '豪商';
    else if (newLife >= 45) newRank = '商人';
    else if (newLife >= 12) newRank = '丁稚';

    const now = new Date();
    userSheet.getRange(senderRow + 2, colIdx.wallet_balance + 1).setValue(newBal);
    userSheet.getRange(senderRow + 2, colIdx.memo + 1).setValue(JSON.stringify(memoObj));
    userSheet.getRange(senderRow + 2, colIdx.last_updated + 1).setValue(now);

    userSheet.getRange(receiverRow + 2, colIdx.lifetime_received + 1).setValue(newLife);
    if (newRank !== receiverData[colIdx.rank]) {
      userSheet.getRange(receiverRow + 2, colIdx.rank + 1).setValue(newRank);
    }

    // shareFlagのデフォルトはtrue
    const isShareable = (shareFlag === undefined || shareFlag === null) ? true : shareFlag;

    transSheet.appendRow([
      Utilities.getUuid(), now, senderEmail, receiverEmail,
      senderData[colIdx.department], receiverData[colIdx.department],
      amount, multiplier, amount, valueGained, comment, isShareable
    ]);

    const cache = CacheService.getScriptCache();
    cache.remove('HISTORY_' + senderEmail);
    cache.remove('HISTORY_' + receiverEmail);
    cache.remove('RANKINGS_v6'); // キャッシュクリア（バージョン更新）

    return {
      success: true, message: '送信完了！',
      newBalance: newBal, dailySent: memoObj.daily_total
    };

  } catch (e) {
    return { success: false, message: e.message };
  } finally {
    lock.releaseLock();
  }
}

// --- Ranking Logic (Updated for Best Giver) ---
function getRankings() {
  const cache = CacheService.getScriptCache();
  const cached = cache.get('RANKINGS_v6'); // キャッシュバージョンアップ
  if (cached) return { success: true, rankings: JSON.parse(cached) };

  const ss = getSpreadsheet();
  
  // 1. MVP (Received) & Dept Headcount
  const userSheet = ss.getSheetByName(SHEET_NAMES.USERS);
  const userData = userSheet.getDataRange().getValues();
  userData.shift(); // remove header
  
  // Create Email->Name/Dept Map & Count Dept Population
  const userMap = {};
  const deptHeadcount = {}; // 部署ごとの人数用

  userData.forEach(r => {
    const dept = r[2];
    userMap[r[0]] = { name: r[1], dept: dept };
    
    // 部署人数カウント
    if(dept) {
      deptHeadcount[dept] = (deptHeadcount[dept] || 0) + 1;
    }
  });

  const mvp = userData.map(r => ({name: r[1], dept: r[2], score: Number(r[5])}))
    .sort((a,b) => b.score - a.score)
    .slice(0, 10);
    
  // 2. Department & Best Giver (Scan Transactions)
  const transSheet = ss.getSheetByName(SHEET_NAMES.TRANSACTIONS);
  const lastRow = transSheet.getLastRow();

  const deptSendCountMap = {}; // { deptName: countOfSending }
  const deptReceiveCountMap = {}; // { deptName: countOfReceiving }
  const giverMap = {}; // { email: countOfSending }

  if(lastRow >= 2) {
    // パフォーマンス最適化: 最新1000件のみ取得（今月のデータ想定）
    const start = Math.max(2, lastRow - 999);
    const numRows = lastRow - start + 1;

    // Transactions: Col C(3)=Sender, D(4)=Receiver, E(5)=SenderDept, F(6)=ReceiverDept
    // getRange(start, 3, rows, 4) => C, D, E, F (必要な列のみ取得)
    // Index: 0:Sender, 1:Receiver, 2:SenderDept(E), 3:ReceiverDept(F)
    const tData = transSheet.getRange(start, 3, numRows, 4).getValues();

    // 高速ループ処理
    for(let i = 0; i < tData.length; i++) {
      const r = tData[i];
      const sender = r[0];
      const receiver = r[1];
      const senderDept = r[2];
      const receiverDept = r[3];

      // 部署カウント
      if(senderDept) deptSendCountMap[senderDept] = (deptSendCountMap[senderDept]||0) + 1;
      if(receiverDept) deptReceiveCountMap[receiverDept] = (deptReceiveCountMap[receiverDept]||0) + 1;

      // ギバーカウント
      if(sender) giverMap[sender] = (giverMap[sender]||0) + 1;
    }
  }

  // Format Dept Ranking (Per Capita - Send & Receive)
  const allDepts = new Set([...Object.keys(deptSendCountMap), ...Object.keys(deptReceiveCountMap)]);
  const dept = Array.from(allDepts).map(k => {
    const sendCount = deptSendCountMap[k] || 0;
    const receiveCount = deptReceiveCountMap[k] || 0;
    const headcount = deptHeadcount[k] || 1; // 0除算防止
    const sendPerCapita = parseFloat((sendCount / headcount).toFixed(2));
    const receivePerCapita = parseFloat((receiveCount / headcount).toFixed(2));
    return {
      name: k,
      sendScore: sendPerCapita,
      receiveScore: receivePerCapita,
      totalScore: parseFloat((sendPerCapita + receivePerCapita).toFixed(2))
    };
  }).sort((a,b) => b.totalScore - a.totalScore).slice(0, 5);
  
  // Format Giver Ranking (Count Based)
  const giver = Object.keys(giverMap).map(k => {
    const u = userMap[k] || { name: k.split('@')[0], dept: '不明' };
    return { name: u.name, dept: u.dept, score: giverMap[k] };
  }).sort((a,b) => b.score - a.score).slice(0, 10);
    
  const rankings = { mvp: mvp, dept: dept, giver: giver };
  cache.put('RANKINGS_v6', JSON.stringify(rankings), 900); // 15 min cache
  return { success: true, rankings: rankings };
}

// --- Archive Logic for Past Rankings ---

function getArchiveMonths() {
  try {
    if (!APP_CONFIG.ARCHIVE_SS_ID) return { success: false, message: 'Archive SS Not Configured' };
    const ss = SpreadsheetApp.openById(APP_CONFIG.ARCHIVE_SS_ID);
    const sheets = ss.getSheets();
    const months = sheets
      .map(s => s.getName())
      .filter(name => name.match(/^\d{4}_\d{2}$/))
      .sort()
      .reverse();
    return { success: true, months: months };
  } catch(e) {
    return { success: false, message: e.message };
  }
}

function getArchiveRankingData(sheetName) {
  try {
    const archiveSS = SpreadsheetApp.openById(APP_CONFIG.ARCHIVE_SS_ID);
    const sheet = archiveSS.getSheetByName(sheetName);
    if (!sheet) return { success: false, message: 'Sheet not found' };

    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return { success: true, rankings: { mvp: [], dept: [], giver: [] } };

    const data = sheet.getRange(2, 1, lastRow - 1, 11).getValues();

    const userListRes = getUserListData();
    const userMap = {}; 
    const deptHeadcount = {}; // アーカイブ用にも現在の人数を適用（過去の人数は不明なため近似値として利用）

    if (userListRes.success) {
      userListRes.list.forEach(u => {
        userMap[u[0]] = { name: u[1], dept: u[2] };
        if(u[2]) deptHeadcount[u[2]] = (deptHeadcount[u[2]] || 0) + 1;
      });
    }

    const mvpMap = {};
    const deptCountMap = {}; // 部署回数カウント用
    const giverMap = {};

    data.forEach(row => {
      // Archive: A=0, B=1, C=2(Sender), D=3(Receiver), E=4(SenderDept), F=5(RecDept)...
      const sender = row[2];
      const receiver = row[3];
      const senderDept = row[4]; // E列
      // const recDept = row[5]; // 不要
      // const amount = Number(row[6] || 0);
      const valGained = Number(row[9] || 0);

      // MVP (Received Value - Unchanged)
      if (receiver) mvpMap[receiver] = (mvpMap[receiver] || 0) + valGained;
      
      // Dept (Sender Count per Dept)
      if (senderDept) deptCountMap[senderDept] = (deptCountMap[senderDept] || 0) + 1;
      
      // Giver (Count)
      if (sender) giverMap[sender] = (giverMap[sender] || 0) + 1;
    });

    const mvp = Object.keys(mvpMap).map(k => {
      const u = userMap[k] || { name: k.split('@')[0], dept: '退職/不明' };
      return { name: u.name, dept: u.dept, score: mvpMap[k] };
    }).sort((a, b) => b.score - a.score).slice(0, 10);

    const dept = Object.keys(deptCountMap).map(k => {
      const count = deptCountMap[k];
      const headcount = deptHeadcount[k] || 1;
      const perCapita = parseFloat((count / headcount).toFixed(2));
      return { name: k, score: perCapita };
    }).sort((a, b) => b.score - a.score).slice(0, 5);

    const giver = Object.keys(giverMap).map(k => {
      const u = userMap[k] || { name: k.split('@')[0], dept: '退職/不明' };
      return { name: u.name, dept: u.dept, score: giverMap[k] };
    }).sort((a, b) => b.score - a.score).slice(0, 10);

    return { success: true, rankings: { mvp, dept, giver } };

  } catch(e) {
    return { success: false, message: e.message };
  }
}

// --- Common Helpers ---

function getSpreadsheet() { return SpreadsheetApp.openById(APP_CONFIG.SS_ID); }

function getDepartmentsCached() {
  const cache = CacheService.getScriptCache();
  const cached = cache.get('DEPT_LIST');
  if (cached) return JSON.parse(cached);
  const ss = getSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAMES.DEPARTMENTS);
  if (!sheet) return [];
  const list = sheet.getRange(2, 1, sheet.getLastRow()-1 || 1, 1).getValues().flat().filter(String);
  cache.put('DEPT_LIST', JSON.stringify(list), 3600);
  return list;
}

function registerNewUser(form) {
  const ss = getSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAMES.USERS);
  const email = Session.getActiveUser().getEmail();
  sheet.appendRow([email, form.name, form.department, '素浪人', APP_CONFIG.INITIAL_COIN, 0, '{}', new Date()]);
  admin_updateUserJson();
  return { success: true, message: '登録完了' };
}

// ★Optimized History Loading (Reverted to v4.4 Chunk Logic for Speed)★
function getUserHistory() {
  const email = Session.getActiveUser().getEmail();
  const cache = CacheService.getScriptCache();
  const cached = cache.get('HISTORY_' + email);
  if(cached) return { success: true, history: JSON.parse(cached) };
  
  const ss = getSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAMES.TRANSACTIONS);
  const lastRow = sheet.getLastRow();
  if(lastRow < 2) return {success:true, history:[]};
  
  const history = [];
  const CHUNK = 200; // 一度に読み込む行数（少なめに設定して高速化）
  let curr = lastRow;
  
  // ★v4.4のロジック復活: 200件ずつ後ろから遡って取得し、20件溜まったら即終了
  // これにより、最近履歴がある人は200行しか読まないので爆速になる
  while(curr >= 2 && history.length < 20) {
    const start = Math.max(2, curr - CHUNK + 1);
    const numRows = curr - start + 1;
    
    // A列(1)からK列(11)まで取得
    // Index: 0:ID, 1:Time, 2:Sender, 3:Receiver, 4:S_Dept, 5:R_Dept, 6:Amt, 7:Mult, 8:Cost, 9:Val, 10:Msg
    const data = sheet.getRange(start, 1, numRows, 11).getValues();
    
    // 後ろから走査
    for(let i = data.length - 1; i >= 0; i--) {
      const row = data[i];
      // 自分が送信者(row[2]) または 受信者(row[3])
      if(row[2] === email || row[3] === email) {
        history.push({
          timestamp: row[1],
          sender_id: row[2],
          receiver_id: row[3],
          sender_dept: row[4],
          amount: row[6],
          value: row[9],
          message: row[10],
          type: row[2] === email ? 'sent' : 'received'
        });
        if(history.length >= 20) break;
      }
    }
    
    curr -= CHUNK;
    // 安全策: 最大3000行遡ったら諦める（無限ループ防止）
    if(lastRow - curr > 3000) break;
  }
  
  // キャッシュ時間は6時間に設定
  cache.put('HISTORY_' + email, JSON.stringify(history), 21600);
  return { success: true, history: history };
}

function analyzeEconomyState() {
  const ss = getSpreadsheet();
  const transSheet = ss.getSheetByName(SHEET_NAMES.TRANSACTIONS);
  if (!transSheet) return 'level2';
  
  const lastRow = transSheet.getLastRow();
  if (lastRow < 2) return 'level2'; 
  
  const startRow = Math.max(2, lastRow - 1000);
  const data = transSheet.getRange(startRow, 10, lastRow - startRow + 1, 1).getValues();
  
  let totalValue = 0;
  for (let i = 0; i < data.length; i++) {
    totalValue += Number(data[i][0] || 0);
  }

  const l2 = APP_CONFIG.ECONOMY_THRESHOLD_L2; 
  const l3 = APP_CONFIG.ECONOMY_THRESHOLD_L3; 

  if (totalValue >= l3) return 'level3'; 
  if (totalValue >= l2) return 'level2'; 
  return 'level1'; 
}

// --- Batch Functions ---

/**
 * 共有可能なメッセージ一覧を全ユーザーに配信
 * トリガー推奨: 毎週金曜 17:00など
 */
function sendPublicMessageDigest() {
  try {
    const ss = getSpreadsheet();
    const userSheet = ss.getSheetByName(SHEET_NAMES.USERS);
    const transSheet = ss.getSheetByName(SHEET_NAMES.TRANSACTIONS);

    if (!userSheet || !transSheet) return;

    const now = new Date();
    // 過去7日間のデータを取得
    const sevenDaysAgo = new Date(now.getTime() - 7 * 24 * 60 * 60 * 1000);

    // Transactionsシートから flag=true のデータのみ取得
    const lastRow = transSheet.getLastRow();
    if (lastRow < 2) return; // データがない場合は終了

    // 全データ取得 (A～L列: ID, Time, Sender, Receiver, SenderDept, ReceiverDept, Amount, Mult, Cost, Value, Message, Flag)
    const transData = transSheet.getRange(2, 1, lastRow - 1, 12).getValues();

    // ユーザーマップ作成（email → name）
    const userData = userSheet.getDataRange().getValues();
    userData.shift(); // ヘッダー削除
    const userMap = {};
    const allUserEmails = [];
    userData.forEach(row => {
      userMap[row[0]] = row[1]; // email → name
      allUserEmails.push(row[0]);
    });

    // 共有可能なメッセージを収集
    const publicMessages = [];
    transData.forEach(row => {
      const timestamp = new Date(row[1]);
      const senderEmail = row[2];
      const receiverEmail = row[3];
      const message = row[10];
      const shareFlag = row[11]; // L列

      // flag=true かつ過去7日間のデータのみ
      if (shareFlag === true && timestamp >= sevenDaysAgo) {
        const senderName = userMap[senderEmail] || senderEmail.split('@')[0];
        const receiverName = userMap[receiverEmail] || receiverEmail.split('@')[0];
        const dateStr = Utilities.formatDate(timestamp, 'Asia/Tokyo', 'MM/dd HH:mm');

        publicMessages.push({
          date: dateStr,
          sender: senderName,
          receiver: receiverName,
          message: message
        });
      }
    });

    if (publicMessages.length === 0) {
      console.log('共有可能なメッセージがありません');
      return;
    }

    // メール本文作成
    let emailBody = `お疲れ様です！\n\n`;
    emailBody += `今週（過去7日間）の社内メッセージをお届けします 📬\n`;
    emailBody += `皆さんがシェアしてくれた温かいメッセージをご覧ください！\n\n`;
    emailBody += `━━━━━━━━━━━━━━━━━━━━\n\n`;

    publicMessages.forEach(msg => {
      emailBody += `【${msg.date}】\n`;
      emailBody += `${msg.sender} さん → ${msg.receiver} さん\n`;
      emailBody += `「${msg.message}」\n\n`;
    });

    emailBody += `━━━━━━━━━━━━━━━━━━━━\n\n`;
    emailBody += `今週も素敵なメッセージがたくさん飛び交っていますね！\n`;
    emailBody += `あなたも感謝の気持ちを伝えてみませんか？\n\n`;
    emailBody += `E-yan Coinアプリ: ${ScriptApp.getService().getUrl()}\n\n`;
    emailBody += `※このメッセージは送信時に「全体共有する」を選択した内容のみ掲載しています`;

    // 全ユーザーにメール送信
    allUserEmails.forEach(email => {
      try {
        GmailApp.sendEmail(
          email,
          '【E-yan Coin】今週のメッセージまとめ 📬',
          emailBody
        );
        Utilities.sleep(500); // レート制限対策
      } catch(e) {
        console.error(`Failed to send to ${email}:`, e);
      }
    });

    console.log(`メッセージ一覧を ${allUserEmails.length} 人に配信完了`);

  } catch(e) {
    console.error('sendPublicMessageDigest Error:', e);
  }
}

function sendDailyRecap() {
  const now = new Date();
  const isFirstDayOfMonth = (now.getDate() === 1);
  const targetDate = new Date(now.getFullYear(), now.getMonth(), now.getDate() - 1);
  const targetDateStr = Utilities.formatDate(targetDate, 'Asia/Tokyo', 'yyyy-MM-dd');
  
  let transData = [];

  if (isFirstDayOfMonth) {
    try {
      if (!APP_CONFIG.ARCHIVE_SS_ID) return;
      const archiveSS = SpreadsheetApp.openById(APP_CONFIG.ARCHIVE_SS_ID);
      const archiveSheetName = Utilities.formatDate(targetDate, 'Asia/Tokyo', 'yyyy_MM');
      const sheet = archiveSS.getSheetByName(archiveSheetName);
      if (sheet && sheet.getLastRow() > 1) {
        transData = sheet.getRange(2, 1, sheet.getLastRow() - 1, 11).getValues();
      }
    } catch(e) { console.error('Archive Access Error', e); return; }
  } else {
    const ss = getSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_NAMES.TRANSACTIONS);
    if (sheet && sheet.getLastRow() > 1) {
      transData = sheet.getRange(2, 1, sheet.getLastRow() - 1, 11).getValues();
    }
  }

  const recipientsMap = {}; 
  transData.forEach(row => {
    const rowTime = new Date(row[1]);
    const rowDateStr = Utilities.formatDate(rowTime, 'Asia/Tokyo', 'yyyy-MM-dd');
    if (rowDateStr === targetDateStr) {
      const receiver = row[3];
      const sender = row[2];
      const msg = row[10];
      if (!recipientsMap[receiver]) recipientsMap[receiver] = [];
      recipientsMap[receiver].push({ sender: sender, msg: msg });
    }
  });

  const appUrl = ScriptApp.getService().getUrl();
  Object.keys(recipientsMap).forEach(email => {
    const msgs = recipientsMap[email];
    if (msgs.length === 0) return;
    let body = `お疲れ様です。\n昨日、${msgs.length}件のE-yan Coinメッセージを受け取りました！\n\n`;
    msgs.forEach(m => { body += `■ ${m.sender}さんより\n「${m.msg}」\n\n`; });
    body += `獲得枚数や詳細はアプリで確認してください。\n${appUrl}\n\n今日も良い一日を！`;
    try {
      GmailApp.sendEmail(email, '【E-yan Coin】メッセージが届いています', body);
      Utilities.sleep(500); 
    } catch(e) { console.error(e); }
  });
}

function checkInactivity() {
  const ss = getSpreadsheet();
  const userSheet = ss.getSheetByName(SHEET_NAMES.USERS);
  const data = userSheet.getDataRange().getValues();
  const header = data.shift();
  const colIdx = {};
  header.forEach((h, i) => colIdx[h] = i);
  const today = new Date();
  const appUrl = ScriptApp.getService().getUrl();

  data.forEach(row => {
    const email = row[colIdx.user_id];
    const memoJson = row[colIdx.memo] || "{}";
    let lastSentDateStr = null;
    try { const memo = JSON.parse(memoJson); lastSentDateStr = memo.last_sent_date; } catch(e) {}

    let daysDiff = 999;
    if (lastSentDateStr) {
      const lastSent = new Date(lastSentDateStr);
      const diffTime = Math.abs(today - lastSent);
      daysDiff = Math.ceil(diffTime / (1000 * 60 * 60 * 24));
    }
    if (daysDiff >= 7) {
      try {
        GmailApp.sendEmail(email, '【E-yan Coin】最近どうですか？', `最近コインを送っていないようです。\nアプリを開く: ${appUrl}`);
        Utilities.sleep(500);
      } catch(e) { console.error(e); }
    }
  });
}

function resetMonthlyData() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(600000)) return;
  try {
    const ss = getSpreadsheet();
    const userSheet = ss.getSheetByName(SHEET_NAMES.USERS);
    const transSheet = ss.getSheetByName(SHEET_NAMES.TRANSACTIONS);
    const archiveId = APP_CONFIG.ARCHIVE_SS_ID;
    const initialCoin = APP_CONFIG.INITIAL_COIN;

    const userData = userSheet.getDataRange().getValues();
    const header = userData.shift();
    const colIdx = {};
    header.forEach((h, i) => colIdx[h] = i);

    let mvpEmail = '';
    let maxLifetime = -1;
    userData.forEach(row => {
      const lifetime = Number(row[colIdx.lifetime_received] || 0);
      if (lifetime > maxLifetime) {
        maxLifetime = lifetime;
        mvpEmail = row[colIdx.user_id];
      }
    });
    if (mvpEmail && maxLifetime > 0) saveMVPHistory(mvpEmail, maxLifetime);

    if (transSheet.getLastRow() > 1 && archiveId) {
      try {
        const archiveSS = SpreadsheetApp.openById(archiveId);
        const now = new Date();
        const lastMonth = new Date(now.getFullYear(), now.getMonth() - 1, 1);
        const sheetName = Utilities.formatDate(lastMonth, 'Asia/Tokyo', 'yyyy_MM');
        let targetSheet = archiveSS.getSheetByName(sheetName);
        if (!targetSheet) {
          targetSheet = archiveSS.insertSheet(sheetName);
          const headers = transSheet.getRange(1, 1, 1, transSheet.getLastColumn()).getValues();
          targetSheet.getRange(1, 1, 1, headers[0].length).setValues(headers);
        }
        const transData = transSheet.getRange(2, 1, transSheet.getLastRow() - 1, transSheet.getLastColumn()).getValues();
        if (transData.length > 0) {
          targetSheet.getRange(targetSheet.getLastRow() + 1, 1, transData.length, transData[0].length).setValues(transData);
        }
        transSheet.deleteRows(2, transSheet.getLastRow() - 1);
        let logSheet = ss.getSheetByName(SHEET_NAMES.ARCHIVE_LOG);
        if(!logSheet) logSheet = ss.insertSheet(SHEET_NAMES.ARCHIVE_LOG);
        logSheet.appendRow([new Date(), `Archived to ${sheetName}`, `${transData.length} rows`]);
      } catch (e) { console.error('Archive failed', e); }
    }

    const numRows = userData.length;
    if (numRows > 0) {
      userSheet.getRange(2, colIdx.rank + 1, numRows, 1).setValue('素浪人'); 
      userSheet.getRange(2, colIdx.wallet_balance + 1, numRows, 1).setValue(initialCoin);
      userSheet.getRange(2, colIdx.lifetime_received + 1, numRows, 1).setValue(0);
      userSheet.getRange(2, colIdx.memo + 1, numRows, 1).setValue('{}');
    }
    
    const cache = CacheService.getScriptCache();
    cache.remove('ALL_USERS_DATA_v4');
    cache.remove('ECONOMY_STATE_v4');
    cache.remove('RANKINGS_v6');
  } catch (e) { console.error(e); } finally { lock.releaseLock(); }
}

function sendReminderEmails() {
  const threshold = APP_CONFIG.REMINDER_THRESHOLD;
  const ss = getSpreadsheet();
  const userSheet = ss.getSheetByName(SHEET_NAMES.USERS);
  const data = userSheet.getDataRange().getValues();
  const header = data.shift();
  const colIdx = {};
  header.forEach((h, i) => colIdx[h] = i);
  data.forEach(row => {
    const bal = Number(row[colIdx.wallet_balance]);
    const email = row[colIdx.user_id];
    const name = row[colIdx.name];
    if (bal >= threshold) {
      try {
        GmailApp.sendEmail(email, "【E-yan Coin】コインを使い切りましょう！", `${name}さん\n残高: ${bal}枚\n${ScriptApp.getService().getUrl()}`);
        Utilities.sleep(100);
      } catch(e){}
    }
  });
}

function saveMVPHistory(email, score) {
  const ss = getSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAMES.MVP_HISTORY);
  if (!sheet) { sheet = ss.insertSheet(SHEET_NAMES.MVP_HISTORY); sheet.appendRow(['YM','Email','Score','Time']); }
  const ym = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM');
  sheet.appendRow([ym, email, score, new Date()]);
}

function getLastMonthMVP() {
  const ss = getSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAMES.MVP_HISTORY);
  if (!sheet || sheet.getLastRow() < 2) return null;
  return sheet.getRange(sheet.getLastRow(), 2).getValue();
}

// --- Email Management Functions ---

/**
 * E-yanCoin関連の送信済みメールを削除
 * ただし、ma-hayashi@race-number.co.jp宛ては保護
 */
function deleteEyanCoinSentEmails() {
  try {
    // E-yanCoin関連のメールを検索
    const query = 'subject:【E-yan Coin】 in:sent';
    const threads = GmailApp.search(query, 0, 100); // 最大100件

    let deletedCount = 0;
    let protectedCount = 0;

    threads.forEach(thread => {
      const messages = thread.getMessages();
      if (messages.length === 0) return;

      // 最初のメッセージの宛先を確認
      const firstMessage = messages[0];
      const toAddress = firstMessage.getTo();

      // ma-hayashi宛てかチェック
      if (toAddress.includes('ma-hayashi@race-number.co.jp')) {
        protectedCount++;
        console.log('Protected: ' + toAddress);
      } else {
        // ma-hayashi以外は削除
        thread.moveToTrash();
        deletedCount++;
        console.log('Deleted: ' + toAddress);
      }
    });

    console.log(`処理完了: ${deletedCount}件削除, ${protectedCount}件保護`);
    return { success: true, deleted: deletedCount, protected: protectedCount };

  } catch (e) {
    console.error('メール削除エラー: ' + e.toString());
    return { success: false, error: e.toString() };
  }
}

