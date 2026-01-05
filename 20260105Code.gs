/**
 * E-yan Coin App - 大阪支社 (v5.1 Fixed Edition)
 * 修正点：景気判断ロジックにおける直近1000件制限の撤廃、キャッシュキー更新
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
  ARCHIVE_SS_ID: '1Gk3B_yd0q-sqskmQwHBsWk0PfYbSqfD0UdzYYiMhN5w', // アーカイブSS

  // ★JSONファイルID (高速化用)
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
    // ★修正: キャッシュキーを変更して即時反映させる (v4 -> v5)
    let economyState = cache.get('ECONOMY_STATE_v5');
    if (!economyState) {
      economyState = analyzeEconomyState();
      cache.put('ECONOMY_STATE_v5', economyState, 600);
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

// --- Core Logic : 送金処理 (L列フラグ対応: 1=共有NG, 0=共有OK) ---
function sendAirCoin(receiverEmail, comment, amountInput, isHidden) {
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

    // ランク判定
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

    // ★共有フラグ: 1=共有NG, 0=共有OK
    const shareFlag = isHidden ? 1 : 0;

    // ★L列(12番目の要素)にフラグを追加して保存
    transSheet.appendRow([
      Utilities.getUuid(), now, senderEmail, receiverEmail,
      senderData[colIdx.department], receiverData[colIdx.department],
      amount, multiplier, amount, valueGained, comment,
      shareFlag
    ]);

    const cache = CacheService.getScriptCache();
    cache.remove('HISTORY_' + senderEmail);
    cache.remove('HISTORY_' + receiverEmail);
    cache.remove('RANKINGS_v6');
    // ★送金時にキャッシュを消して、次回の取得で最新の景気を反映させる
    cache.remove('ECONOMY_STATE_v5');

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

// --- Ranking Logic (部署別コイン総数を含む統合版) ---
function getRankings() {
  const cache = CacheService.getScriptCache();
  const cached = cache.get('RANKINGS_v6');
  if (cached) return { success: true, rankings: JSON.parse(cached) };

  const ss = getSpreadsheet();

  // 1. MVP (Received) & Dept Headcount
  const userSheet = ss.getSheetByName(SHEET_NAMES.USERS);
  const userData = userSheet.getDataRange().getValues();
  userData.shift();

  const userMap = {};
  const deptHeadcount = {};

  userData.forEach(r => {
    const dept = r[2];
    userMap[r[0]] = { name: r[1], dept: dept };
    if(dept) deptHeadcount[dept] = (deptHeadcount[dept] || 0) + 1;
  });

  const mvp = userData.map(r => ({name: r[1], dept: r[2], score: Number(r[5])}))
    .sort((a,b) => b.score - a.score)
    .slice(0, 10);

  // 2. Scan Transactions
  const transSheet = ss.getSheetByName(SHEET_NAMES.TRANSACTIONS);
  const lastRow = transSheet.getLastRow();

  const deptCountMap = {}; // 回数カウント用
  const deptCoinMap = {};  // コイン総数集計用
  const giverMap = {};

  if(lastRow >= 2) {
    const start = 2;
    const tData = transSheet.getRange(start, 3, lastRow - start + 1, 8).getValues();

    tData.forEach(r => {
      const sender = r[0];
      const senderDept = r[2];
      const amount = Number(r[4] || 0);

      if(senderDept) {
        // A. 既存の回数ロジック
        deptCountMap[senderDept] = (deptCountMap[senderDept]||0) + 1;
        // B. 新規: コイン総数ロジック
        deptCoinMap[senderDept] = (deptCoinMap[senderDept]||0) + amount;
      }

      if(sender) giverMap[sender] = (giverMap[sender]||0) + 1;
    });
  }

  // A. 部署ランキング (一人当たり平均回数)
  const dept = Object.keys(deptCountMap).map(k => {
    const count = deptCountMap[k];
    const headcount = deptHeadcount[k] || 1;
    const perCapitaScore = parseFloat((count / headcount).toFixed(2));
    return { name: k, score: perCapitaScore };
  }).sort((a,b) => b.score - a.score).slice(0, 5);

  // B. 部署ランキング (総コイン数)
  const deptTotal = Object.keys(deptCoinMap).map(k => {
    return { name: k, score: deptCoinMap[k] };
  }).sort((a,b) => b.score - a.score).slice(0, 5);

  // Giver Ranking
  const giver = Object.keys(giverMap).map(k => {
    const u = userMap[k] || { name: k.split('@')[0], dept: '不明' };
    return { name: u.name, dept: u.dept, score: giverMap[k] };
  }).sort((a,b) => b.score - a.score).slice(0, 10);

  const rankings = { mvp: mvp, dept: dept, deptTotal: deptTotal, giver: giver };
  cache.put('RANKINGS_v6', JSON.stringify(rankings), 900);
  return { success: true, rankings: rankings };
}

// --- Archive Logic ---

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
    const deptHeadcount = {};

    if (userListRes.success) {
      userListRes.list.forEach(u => {
        userMap[u[0]] = { name: u[1], dept: u[2] };
        if(u[2]) deptHeadcount[u[2]] = (deptHeadcount[u[2]] || 0) + 1;
      });
    }

    const mvpMap = {};
    const deptCountMap = {};
    const giverMap = {};

    data.forEach(row => {
      const sender = row[2];
      const receiver = row[3];
      const senderDept = row[4];
      const valGained = Number(row[9] || 0);

      if (receiver) mvpMap[receiver] = (mvpMap[receiver] || 0) + valGained;
      if (senderDept) deptCountMap[senderDept] = (deptCountMap[senderDept] || 0) + 1;
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
  const CHUNK = 200;
  let curr = lastRow;

  while(curr >= 2 && history.length < 20) {
    const start = Math.max(2, curr - CHUNK + 1);
    const numRows = curr - start + 1;

    const data = sheet.getRange(start, 1, numRows, 11).getValues();

    for(let i = data.length - 1; i >= 0; i--) {
      const row = data[i];
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
    if(lastRow - curr > 3000) break;
  }

  cache.put('HISTORY_' + email, JSON.stringify(history), 21600);
  return { success: true, history: history };
}

// --- ★修正箇所: 全期間の流通量を正しく計算する ---
function analyzeEconomyState() {
  const ss = getSpreadsheet();
  const transSheet = ss.getSheetByName(SHEET_NAMES.TRANSACTIONS);
  if (!transSheet) return 'level2';

  const lastRow = transSheet.getLastRow();
  if (lastRow < 2) return 'level2';

  // ★修正前: const startRow = Math.max(2, lastRow - 1000); // 1000行制限により全体の集計が合わない
  // ★修正後: 2行目から全データを取得
  const startRow = 2;
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
 * 日次レポート（昨日受信したメッセージを全ユーザーに配信）
 * トリガー推奨: 毎日午前9時など
 * ★修正版：全ユーザー送信 & shareFlag==0（共有OK）のみ対象
 */
function sendDailyReportEmail() {
  // 1. 日付の定義（昨日）
  const now = new Date();
  const yesterday = new Date(now.getFullYear(), now.getMonth(), now.getDate() - 1);
  const yesterdayStr = Utilities.formatDate(yesterday, 'Asia/Tokyo', 'yyyy/MM/dd');

  // 2. スプレッドシートを開く
  const ss = getSpreadsheet();
  const transSheet = ss.getSheetByName(SHEET_NAMES.TRANSACTIONS);
  const usersSheet = ss.getSheetByName(SHEET_NAMES.USERS);

  // 3. ユーザー名簿を取得
  const userMap = getUserMapForEmail(usersSheet);
  const allUserEmails = Object.keys(userMap);

  // 4. 履歴データを取得
  const lastRow = transSheet.getLastRow();
  if (lastRow < 2) {
    console.log('トランザクションデータがありません。');
    return;
  }

  const data = transSheet.getRange(2, 1, lastRow - 1, 12).getValues();

  // 5. データ抽出
  const reportLinesHtml = []; // HTML用配列
  let count = 0;

  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    const rowDate = new Date(row[1]);
    const rowDateStr = Utilities.formatDate(rowDate, 'Asia/Tokyo', 'yyyy/MM/dd');
    const shareFlag = row[11];

    // ★shareFlag==0（共有OK）のみ対象
    if (rowDateStr === yesterdayStr && shareFlag == 0) {
      const senderEmail = row[2];
      const receiverEmail = row[3];
      const message = row[10];

      const senderName = userMap[senderEmail] || senderEmail;
      const receiverName = userMap[receiverEmail] || receiverEmail;

      // 文字化け対策：数値文字参照に変換
      const safeSender = escapeToEntities(senderName);
      const safeReceiver = escapeToEntities(receiverName);
      const safeMessage = escapeToEntities(message);

      reportLinesHtml.push(
        `<div style="margin-bottom: 15px; padding: 10px; border-left: 4px solid #4f46e5; background-color: #f9fafb;">` +
        `  <strong>${safeSender}</strong> さん <span style="color:#999">➡</span> <strong>${safeReceiver}</strong> さん<br>` +
        `  <span style="color: #374151;">「${safeMessage}」</span>` +
        `</div>`
      );
      count++;
    }
  }

  // 6. メール送信（★全ユーザーに送信）
  if (count > 0) {
    const subject = `【E-yan Coin】昨日の称賛一覧 (${yesterdayStr})`;

    // HTML本文
    const bodyHtml = `
      <!DOCTYPE html>
      <html>
        <head>
          <meta charset="UTF-8">
        </head>
        <body style="font-family: sans-serif; color: #333;">
          <h2 style="color: #4f46e5;">${yesterdayStr} の称賛履歴</h2>
          <p>昨日のE-yan Coinメッセージをお届けします。<br>※共有不可のものは除外されています。</p>
          <hr style="border: 0; border-top: 1px solid #eee; margin: 20px 0;">

          ${reportLinesHtml.join('')}

          <hr style="border: 0; border-top: 1px solid #eee; margin: 20px 0;">
          <p><strong>合計: ${count} 件</strong>の称賛がありました！<br>今日も一日がんばりましょう！</p>
        </body>
      </html>
    `;

    const bodyText = "このメールはHTML形式で表示してください。";

    // ★全ユーザーにメール送信
    allUserEmails.forEach(email => {
      try {
        GmailApp.sendEmail(email, subject, bodyText, {
          htmlBody: bodyHtml
        });
        Utilities.sleep(500); // レート制限対策
      } catch(e) {
        console.error(`Failed to send to ${email}:`, e);
      }
    });

    console.log(`送信完了: ${allUserEmails.length}名に配信 (${count}件の称賛)`);
  } else {
    console.log('昨日の対象データはありませんでした。');
  }
}

/**
 * 個別ユーザーへの受信通知（受信者のみに送信）
 * トリガー推奨: 毎日午前9時～10時
 */
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
        
        // ★修正: 固定されていない行をすべて削除することはできないエラーを回避するため、
        // 強制的に1行目を固定してから削除を実行する
        transSheet.setFrozenRows(1);
        
        const currentLastRow = transSheet.getLastRow();
        if (currentLastRow >= 2) {
          transSheet.deleteRows(2, currentLastRow - 1);
        }
        
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
    cache.remove('ECONOMY_STATE_v5'); // ★修正: v5に対応
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
    const threads = GmailApp.search(query, 0, 500); // 最大500件

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

// --- Helper Functions for Email ---

function getUserMapForEmail(sheet) {
  const map = {};
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return map;
  const data = sheet.getRange(2, 1, lastRow - 1, 2).getValues();
  data.forEach(row => { map[row[0]] = row[1]; });
  return map;
}

/**
 * 文字化け対策：文字列を数値文字参照に変換
 * 日本語や絵文字を &#12345; 形式に変換
 */
function escapeToEntities(text) {
  if (!text) return "";
  return Array.from(text).map(char => {
    const code = char.codePointAt(0);
    // ASCII文字（英数字記号）はそのまま、それ以外（日本語・絵文字）は変換
    return code > 127 ? `&#${code};` : char;
  }).join('');
}
