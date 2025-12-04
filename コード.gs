/**
 * E-yan Coin App - 大阪支社 (v4.1 Full Implementation)
 * * 【高速化対応】
 * 1. Lazy Loading: ユーザーリストは別便で取得
 * 2. Array-ification: データ通信量を配列化で削減
 * 3. Config: シート読み込みを廃止し定数化
 * 4. JSON Storage: ユーザーリストをJSONファイルとしてDriveにキャッシュ
 */

// --- ★設定エリア (Configシートの代わり) ---
const APP_CONFIG = {
  INITIAL_COIN: 100,           // 月初の所持コイン
  MULTIPLIER_DIFF_DEPT: 1.5,   // 他部署倍率
  MESSAGE_MAX_LENGTH: 100,     // メッセージ文字数上限
  ECONOMY_THRESHOLD_L2: 10000, // 景気Lv2閾値
  ECONOMY_THRESHOLD_L3: 50000, // 景気Lv3閾値
  REMINDER_THRESHOLD: 50,      // リマインド閾値
  
  // ID設定
  SS_ID: '1E0qf3XM-W8TM5HZ_SrPPoGAV4kwObvS6FmQdaFR3Bpw', // メインSS
  ARCHIVE_SS_ID: '1Gk3B_yd0q-sqskmQwHBsWk0PfYbSqfD0UdzYYiMhN5w', // アーカイブSS
  
  // ★手順で生成されたJSONファイルIDをここに貼る
  JSON_FILE_ID: '' 
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

// --- API 1: 起動直後の軽量データ取得 (自分の情報のみ) ---

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

    // 景気状態 (キャッシュ使用)
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

// --- API 2: ユーザーリスト取得 (JSONファイル経由 & 配列化) ---

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
  console.log('JSON file created/updated. ID:', file.getId());
  return file.getId();
}

// --- Core Logic : 送金処理 ---

function sendAirCoin(receiverEmail, comment, amountInput) {
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

    let newRank = receiverData[colIdx.rank];
    if (newLife >= 10000) newRank = '天下人';
    else if (newLife >= 5000) newRank = '豪商';
    else if (newLife >= 1000) newRank = '商人';
    else if (newLife >= 100) newRank = '丁稚';

    const now = new Date();
    userSheet.getRange(senderRow + 2, colIdx.wallet_balance + 1).setValue(newBal);
    userSheet.getRange(senderRow + 2, colIdx.memo + 1).setValue(JSON.stringify(memoObj));
    userSheet.getRange(senderRow + 2, colIdx.last_updated + 1).setValue(now);

    userSheet.getRange(receiverRow + 2, colIdx.lifetime_received + 1).setValue(newLife);
    if (newRank !== receiverData[colIdx.rank]) {
      userSheet.getRange(receiverRow + 2, colIdx.rank + 1).setValue(newRank);
    }

    transSheet.appendRow([
      Utilities.getUuid(), now, senderEmail, receiverEmail,
      senderData[colIdx.department], receiverData[colIdx.department],
      amount, multiplier, amount, valueGained, comment
    ]);

    const cache = CacheService.getScriptCache();
    cache.remove('HISTORY_' + senderEmail);
    cache.remove('HISTORY_' + receiverEmail);

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

// --- ヘルパー & その他 ---

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

function getRankings() {
  const cache = CacheService.getScriptCache();
  const cached = cache.get('RANKINGS_v4');
  if (cached) return { success: true, rankings: JSON.parse(cached) };

  const ss = getSpreadsheet();
  const userSheet = ss.getSheetByName(SHEET_NAMES.USERS);
  const userData = userSheet.getDataRange().getValues();
  userData.shift();
  const mvp = userData.map(r => ({name: r[1], dept: r[2], score: Number(r[5])}))
    .sort((a,b) => b.score - a.score)
    .slice(0, 10);
    
  const transSheet = ss.getSheetByName(SHEET_NAMES.TRANSACTIONS);
  const lastRow = transSheet.getLastRow();
  const deptMap = {};
  if(lastRow >= 2) {
    const start = Math.max(2, lastRow - 1000);
    const tData = transSheet.getRange(start, 6, lastRow - start + 1, 5).getValues(); 
    tData.forEach(r => {
      const d = r[0]; 
      const v = Number(r[4]||0); 
      if(d) deptMap[d] = (deptMap[d]||0) + v;
    });
  }
  const dept = Object.keys(deptMap).map(k => ({name: k, score: deptMap[k]}))
    .sort((a,b) => b.score - a.score).slice(0, 5);
    
  const rankings = { mvp: mvp, dept: dept, giver: [] };
  cache.put('RANKINGS_v4', JSON.stringify(rankings), 900);
  return { success: true, rankings: rankings };
}

function getUserHistory() {
  const email = Session.getActiveUser().getEmail();
  const cache = CacheService.getScriptCache();
  const cached = cache.get('HISTORY_' + email);
  if(cached) return { success: true, history: JSON.parse(cached) };
  
  const ss = getSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAMES.TRANSACTIONS);
  const lastRow = sheet.getLastRow();
  if(lastRow<2) return {success:true, history:[]};
  
  const history = [];
  const CHUNK = 200;
  let curr = lastRow;
  
  while(curr >= 2 && history.length < 20) {
    const start = Math.max(2, curr - CHUNK + 1);
    const data = sheet.getRange(start, 2, curr - start + 1, 10).getValues();
    for(let i=data.length-1; i>=0; i--) {
      if(data[i][2] === email) {
        history.push({
          timestamp: data[i][0],
          sender_id: data[i][1],
          sender_dept: data[i][3],
          amount: data[i][5],
          value: data[i][8],
          message: data[i][9]
        });
        if(history.length >= 20) break;
      }
    }
    curr -= CHUNK;
    if(lastRow - curr > 2000) break;
  }
  
  cache.put('HISTORY_' + email, JSON.stringify(history), 300);
  return { success: true, history: history };
}

// --- 省略されていた関数群の実装 ---

// 景気分析 (Config定数を使用)
function analyzeEconomyState() {
  const ss = getSpreadsheet();
  const transSheet = ss.getSheetByName(SHEET_NAMES.TRANSACTIONS);
  if (!transSheet) return 'level2';
  
  const lastRow = transSheet.getLastRow();
  if (lastRow < 2) return 'level2'; // データが少なければ通常
  
  // 直近1000件の value_gained (J列) を集計
  // J列は10列目
  const startRow = Math.max(2, lastRow - 1000);
  const data = transSheet.getRange(startRow, 10, lastRow - startRow + 1, 1).getValues();
  
  let totalValue = 0;
  for (let i = 0; i < data.length; i++) {
    totalValue += Number(data[i][0] || 0);
  }

  const l2 = APP_CONFIG.ECONOMY_THRESHOLD_L2; // 10000
  const l3 = APP_CONFIG.ECONOMY_THRESHOLD_L3; // 50000

  if (totalValue >= l3) return 'level3'; // 好景気 (Pink)
  if (totalValue >= l2) return 'level2'; // 通常 (Purple)
  return 'level1';                       // 不況 (Blue)
}

// 月次リセット (Config定数を使用)
function resetMonthlyData() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(600000)) return;

  try {
    const ss = getSpreadsheet();
    const userSheet = ss.getSheetByName(SHEET_NAMES.USERS);
    const transSheet = ss.getSheetByName(SHEET_NAMES.TRANSACTIONS);
    
    // Config定数からアーカイブID取得
    const archiveId = APP_CONFIG.ARCHIVE_SS_ID;
    const initialCoin = APP_CONFIG.INITIAL_COIN;

    // 1. MVP保存
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

    // 2. アーカイブ
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

    // 3. ユーザーリセット
    const numRows = userData.length;
    if (numRows > 0) {
      userSheet.getRange(2, colIdx.rank + 1, numRows, 1).setValue('素浪人');
      userSheet.getRange(2, colIdx.wallet_balance + 1, numRows, 1).setValue(initialCoin);
      userSheet.getRange(2, colIdx.lifetime_received + 1, numRows, 1).setValue(0);
      userSheet.getRange(2, colIdx.memo + 1, numRows, 1).setValue('{}');
    }
    
    // キャッシュ全削除
    const cache = CacheService.getScriptCache();
    cache.remove('ALL_USERS_DATA_v4');
    cache.remove('ECONOMY_STATE_v4');
    cache.remove('RANKINGS_v4');

  } catch (e) { console.error(e); } finally { lock.releaseLock(); }
}

// リマインド (Config定数を使用)
function sendReminderEmails() {
  // Configから閾値取得
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
        GmailApp.sendEmail(
          email,
          "【E-yan Coin】コインを使い切りましょう！",
          `${name}さん\n\n今月の残高: ${bal}枚\n月末リセットされます。\n\n${ScriptApp.getService().getUrl()}`
        );
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

function deleteRetiredUsers() {
  const m = new Date().getMonth() + 1;
  if (![1, 4, 7, 10].includes(m)) return;
  console.log('Quarterly cleanup check (Manual)');
}
