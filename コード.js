/* コード.gs */

// ==============================
// 基本設定
// ==============================

const SHEET_USERS = 'users';
const SHEET_GROUPS = 'groups';
const SHEET_LOGS = 'logs';
const SHEET_STATUS = 'current_status';
const SHEET_USAGE_LOGS = 'usage_logs';
const SHEET_USAGE_SUMMARY = 'usage_summary';

const TZ = 'Asia/Tokyo';// 通知対象にする経過時間
const OVERDUE_HOURS = 24;
const USAGE_MONITOR_DAYS = 30;
const LOW_USAGE_OPEN_THRESHOLD = 3;
const LOW_USAGE_ACTION_THRESHOLD = 1;

// シート見出し
const HEADERS = {
  USERS: ['名前', '割当端末', '本人メール', 'グループ', '有効/無効'],
  GROUPS: ['グループ', 'リーダー名', 'リーダーメール', '有効/無効'],
  LOGS: ['日時', '名前', '端末', 'アクション', '備考'],
  STATUS: ['名前', '端末', '現在状態', '貸出日時', '返却日時', '最終更新', '未返却'],
  USAGE_LOGS: ['日時', '名前', '本人メール', 'グループ', '端末', '操作', '備考'],
  USAGE_SUMMARY: ['名前', 'グループ', '端末', '最終利用日時', '30日表示回数', '30日操作回数', '判定', '備考']
};

// ==============================
// 画面表示
// ==============================

function doGet(e) {
  const template = HtmlService.createTemplateFromFile('index');

  // URL末尾に ?mode=admin を付けると管理画面
  // 例：https://script.google.com/xxxx/exec?mode=admin
  template.mode = e && e.parameter && e.parameter.mode === 'admin'
    ? 'admin'
    : 'user';

  return template
    .evaluate()
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, user-scalable=no')
    .setTitle('タブレット管理 TabLog');
}

// ==============================
// 初期セットアップ
// ==============================

function setupApp() {
  setupSheets();
  setupActiveStatusDropdowns();
  syncCurrentStatusFromUsers();

  return { success: true, message: '初期セットアップが完了しました。' };
}

function setupSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const userSheet = getOrCreateSheet_(ss, SHEET_USERS);
  const groupSheet = getOrCreateSheet_(ss, SHEET_GROUPS);
  const logSheet = getOrCreateSheet_(ss, SHEET_LOGS);
  const statusSheet = getOrCreateSheet_(ss, SHEET_STATUS);
  const usageLogSheet = getOrCreateSheet_(ss, SHEET_USAGE_LOGS);
  const usageSummarySheet = getOrCreateSheet_(ss, SHEET_USAGE_SUMMARY);

  setHeaderIfEmpty_(userSheet, HEADERS.USERS);
  setHeaderIfEmpty_(groupSheet, HEADERS.GROUPS);
  setHeaderIfEmpty_(logSheet, HEADERS.LOGS);
  setHeaderIfEmpty_(statusSheet, HEADERS.STATUS);
  setHeaderIfEmpty_(usageLogSheet, HEADERS.USAGE_LOGS);
  setHeaderIfEmpty_(usageSummarySheet, HEADERS.USAGE_SUMMARY);

  formatSheet_(userSheet);
  formatSheet_(groupSheet);
  formatSheet_(logSheet);
  formatSheet_(statusSheet);
  formatSheet_(usageLogSheet);
  formatSheet_(usageSummarySheet);
}

function getOrCreateSheet_(ss, sheetName) {
  let sheet = ss.getSheetByName(sheetName);

  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
  }

  return sheet;
}

function setHeaderIfEmpty_(sheet, headers) {
  const firstRow = sheet.getRange(1, 1, 1, headers.length).getValues()[0];
  const hasHeader = firstRow.some(value => value !== '');

  if (!hasHeader) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  }
}

function formatSheet_(sheet) {
  const lastColumn = sheet.getLastColumn();
  if (lastColumn === 0) return;

  sheet.getRange(1, 1, 1, lastColumn)
    .setFontWeight('bold')
    .setBackground('#e3f2fd');

  sheet.setFrozenRows(1);

  // users シート
  if (sheet.getName() === SHEET_USERS) {
    sheet.setColumnWidth(1, 140); // 名前
    sheet.setColumnWidth(2, 120); // 割当端末
    sheet.setColumnWidth(3, 260); // 本人メール
    sheet.setColumnWidth(4, 140); // グループ
    sheet.setColumnWidth(5, 100); // 有効/無効
    return;
  }

  // groups シート
  if (sheet.getName() === SHEET_GROUPS) {
    sheet.setColumnWidth(1, 140); // グループ
    sheet.setColumnWidth(2, 160); // リーダー名
    sheet.setColumnWidth(3, 260); // リーダーメール
    sheet.setColumnWidth(4, 100); // 有効/無効
    return;
  }

  // logs シート
  if (sheet.getName() === SHEET_LOGS) {
    sheet.setColumnWidth(1, 160); // 日時
    sheet.setColumnWidth(2, 140); // 名前
    sheet.setColumnWidth(3, 120); // 端末
    sheet.setColumnWidth(4, 100); // アクション
    sheet.setColumnWidth(5, 200); // 備考
    return;
  }

  // current_status シート
  if (sheet.getName() === SHEET_STATUS) {
    sheet.setColumnWidth(1, 140); // 名前
    sheet.setColumnWidth(2, 120); // 端末
    sheet.setColumnWidth(3, 100); // 現在状態
    sheet.setColumnWidth(4, 160); // 貸出日時
    sheet.setColumnWidth(5, 160); // 返却日時
    sheet.setColumnWidth(6, 160); // 最終更新
    sheet.setColumnWidth(7, 100); // 未返却
    return;
  }

  // usage_logs シート
  if (sheet.getName() === SHEET_USAGE_LOGS) {
    sheet.setColumnWidth(1, 160); // 日時
    sheet.setColumnWidth(2, 140); // 名前
    sheet.setColumnWidth(3, 260); // 本人メール
    sheet.setColumnWidth(4, 140); // グループ
    sheet.setColumnWidth(5, 120); // 端末
    sheet.setColumnWidth(6, 120); // 操作
    sheet.setColumnWidth(7, 240); // 備考
    return;
  }

  // usage_summary シート
  if (sheet.getName() === SHEET_USAGE_SUMMARY) {
    sheet.setColumnWidth(1, 140); // 名前
    sheet.setColumnWidth(2, 140); // グループ
    sheet.setColumnWidth(3, 120); // 端末
    sheet.setColumnWidth(4, 160); // 最終利用日時
    sheet.setColumnWidth(5, 120); // 30日表示回数
    sheet.setColumnWidth(6, 120); // 30日操作回数
    sheet.setColumnWidth(7, 120); // 判定
    sheet.setColumnWidth(8, 260); // 備考
    return;
  }
}

// ==============================
// 営業員側：ユーザー情報取得
// ==============================

function getUserList() {
  const sheet = getSheet_(SHEET_USERS);
  const lastRow = sheet.getLastRow();

  if (lastRow < 2) return [];

  const values = sheet.getRange(2, 1, lastRow - 1, 5).getValues();

  return values
    .filter(row => {
      const name = String(row[0] || '').trim();
      const activeStatus = String(row[4] || '').trim();
      return name && activeStatus !== '無効';
    })
    .map(row => String(row[0]).trim());
}

function getUserInfo(userName) {
  const user = findUser_(userName);

  if (!user) {
    return {
      found: false,
      message: 'ユーザーが見つかりません。管理者に確認してください。'
    };
  }

  const status = getCurrentStatusByName_(user.name);

  // current_status にまだ行がなければ初期行を作成
  if (!status) {
    createInitialStatusRow_(user.name, user.device);
  }

  const latestStatus = getCurrentStatusByName_(user.name);

  const isBorrowing = latestStatus && latestStatus.currentStatus === '貸出中';

  appendUsageLog_(user.name, '画面表示', '営業員画面を表示');

  return {
    found: true,
    name: user.name,
    device: user.device,
    group: user.group,
    currentStatus: isBorrowing ? '貸出中' : '返却済',
    nextAction: isBorrowing ? '返却' : '貸出',
    borrowAt: latestStatus ? formatDateTime_(latestStatus.borrowAt) : '',
    returnAt: latestStatus ? formatDateTime_(latestStatus.returnAt) : '',
    updatedAt: latestStatus ? formatDateTime_(latestStatus.updatedAt) : ''
  };
}

// ==============================
// 営業員側：貸出・返却処理
// ==============================

function recordLog(userName, action) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);

  try {
    const normalizedAction = normalizeAction_(action);

    if (!normalizedAction) {
      throw new Error('不正なアクションです。');
    }

    const user = findUser_(userName);

    if (!user) {
      throw new Error('ユーザーが見つかりません。');
    }

    const now = new Date();

    // 1. logs に履歴を追記
    const logSheet = getSheet_(SHEET_LOGS);
    logSheet.appendRow([
      now,
      user.name,
      user.device,
      normalizedAction,
      ''
    ]);

    // アプリ利用状況の監視ログにも記録する
    appendUsageLog_(user.name, normalizedAction, '貸出・返却ボタン操作');

    // 2. current_status を更新
    updateCurrentStatus_(user.name, user.device, normalizedAction, now);

    const info = getUserInfo(user.name);

    return {
      success: true,
      action: normalizedAction,
      message: normalizedAction === '貸出'
        ? '貸出処理が完了しました。'
        : '返却処理が完了しました。',
      info: info
    };

  } finally {
    lock.releaseLock();
  }
}

function updateCurrentStatus_(name, device, action, now) {
  const sheet = getSheet_(SHEET_STATUS);
  const rowNumber = findRowByName_(sheet, name);

  if (rowNumber === -1) {
    if (action === '貸出') {
      sheet.appendRow([
        name,
        device,
        '貸出中',
        now,
        '',
        now,
        true
      ]);
    } else {
      sheet.appendRow([
        name,
        device,
        '返却済',
        '',
        now,
        now,
        false
      ]);
    }
    return;
  }

  const currentRow = sheet.getRange(rowNumber, 1, 1, 7).getValues()[0];
  const currentBorrowAt = currentRow[3];

  if (action === '貸出') {
    sheet.getRange(rowNumber, 1, 1, 7).setValues([[
      name,
      device,
      '貸出中',
      now,
      '',
      now,
      true
    ]]);
  }

  if (action === '返却') {
    sheet.getRange(rowNumber, 1, 1, 7).setValues([[
      name,
      device,
      '返却済',
      currentBorrowAt || '',
      now,
      now,
      false
    ]]);
  }
}

function normalizeAction_(action) {
  const value = String(action || '').trim();

  if (value === '貸出') return '貸出';
  if (value === '返却') return '返却';

  return '';
}

// ==============================
// 管理画面用
// ==============================

function getAdminDashboard() {
  ensureUsageMonitoringSheets_();
  syncCurrentStatusFromUsers();
  const usageMonitoring = getUsageMonitoringData_();
  rebuildUsageSummarySheet_(usageMonitoring.rows);

  const statusSheet = getSheet_(SHEET_STATUS);
  const lastRow = statusSheet.getLastRow();

  if (lastRow < 2) {
    return {
      borrowing: [],
      returned: [],
      summary: {
        total: 0,
        borrowing: 0,
        returned: 0
      },
      usage: usageMonitoring
    };
  }

  const users = getUsersMap_();
  const values = statusSheet.getRange(2, 1, lastRow - 1, 7).getValues();

  const rows = values
    .filter(row => String(row[0] || '').trim())
    .map(row => {
      const name = String(row[0] || '').trim();
      const user = users[name] || {};

      return {
        name: name,
        device: String(row[1] || '').trim(),
        group: user.group || '',
        currentStatus: String(row[2] || '').trim(),
        borrowAt: formatDateTime_(row[3]),
        returnAt: formatDateTime_(row[4]),
        updatedAt: formatDateTime_(row[5]),
        overdue: row[6] === true || String(row[6]).toUpperCase() === 'TRUE',
        elapsed: calculateElapsedText_(row[3], row[2])
      };
    });

  const borrowing = rows
    .filter(row => row.currentStatus === '貸出中')
    .sort((a, b) => {
      const groupCompare = String(a.group || '').localeCompare(String(b.group || ''), 'ja');
      if (groupCompare !== 0) return groupCompare;

      const aTime = a.borrowAt || '';
      const bTime = b.borrowAt || '';
      return aTime > bTime ? 1 : -1;
    });

  const returned = rows
    .filter(row => row.currentStatus !== '貸出中')
    .sort((a, b) => {
      const aTime = a.updatedAt || '';
      const bTime = b.updatedAt || '';
      return aTime < bTime ? 1 : -1;
    });

  return {
    borrowing: borrowing,
    returned: returned,
    summary: {
      total: rows.length,
      borrowing: borrowing.length,
      returned: returned.length
    },
    usage: usageMonitoring
  };
}

// ==============================
// 利用監視
// ==============================

function getUsageMonitoringData_() {
  const users = Object.values(getUsersMap_());
  const usageMap = buildRecentUsageMap_(USAGE_MONITOR_DAYS);

  const rows = users.map(user => {
    const usage = usageMap[user.name] || {
      lastUsedAt: null,
      openCount: 0,
      actionCount: 0
    };

    const judgement = getUsageJudgement_(usage);

    return {
      name: user.name,
      group: user.group || '',
      device: user.device || '',
      lastUsedAt: formatDateTime_(usage.lastUsedAt),
      openCount: usage.openCount,
      actionCount: usage.actionCount,
      judgement: judgement.label,
      judgementLevel: judgement.level,
      note: judgement.note
    };
  }).sort((a, b) => {
    const levelOrder = { danger: 1, warn: 2, ok: 3 };
    const levelCompare = (levelOrder[a.judgementLevel] || 9) - (levelOrder[b.judgementLevel] || 9);
    if (levelCompare !== 0) return levelCompare;

    const groupCompare = String(a.group || '').localeCompare(String(b.group || ''), 'ja');
    if (groupCompare !== 0) return groupCompare;

    return String(a.name || '').localeCompare(String(b.name || ''), 'ja');
  });

  const noUse = rows.filter(row => row.judgement === '未利用');
  const lowUse = rows.filter(row => row.judgement === '低頻度');
  const noAction = rows.filter(row => row.judgement === '操作なし');

  return {
    days: USAGE_MONITOR_DAYS,
    lowUsageOpenThreshold: LOW_USAGE_OPEN_THRESHOLD,
    lowUsageActionThreshold: LOW_USAGE_ACTION_THRESHOLD,
    summary: {
      targetUsers: rows.length,
      noUse: noUse.length,
      lowUse: lowUse.length,
      noAction: noAction.length
    },
    rows: rows
  };
}

function appendUsageLog_(userName, operation, note) {
  try {
    ensureUsageMonitoringSheets_();
    const user = findUser_(userName);
    if (!user) return;

    const sheet = getSheet_(SHEET_USAGE_LOGS);

    sheet.appendRow([
      new Date(),
      user.name,
      user.personalEmail || '',
      user.group || '',
      user.device || '',
      String(operation || '').trim(),
      String(note || '').trim()
    ]);
  } catch (error) {
    console.error('利用ログの記録に失敗しました: ' + error.message);
  }
}

function ensureUsageMonitoringSheets_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const usageLogSheet = getOrCreateSheet_(ss, SHEET_USAGE_LOGS);
  const usageSummarySheet = getOrCreateSheet_(ss, SHEET_USAGE_SUMMARY);

  setHeaderIfEmpty_(usageLogSheet, HEADERS.USAGE_LOGS);
  setHeaderIfEmpty_(usageSummarySheet, HEADERS.USAGE_SUMMARY);
  formatSheet_(usageLogSheet);
  formatSheet_(usageSummarySheet);
}

function buildRecentUsageMap_(days) {
  const map = {};
  const since = new Date();
  since.setDate(since.getDate() - days);
  since.setHours(0, 0, 0, 0);

  appendRecentScreenViewCounts_(map, since);
  appendRecentActionCountsFromLogs_(map, since);

  return map;
}

function appendRecentScreenViewCounts_(map, since) {
  const sheet = getSheet_(SHEET_USAGE_LOGS);
  const lastRow = sheet.getLastRow();

  if (lastRow < 2) return;

  const values = sheet.getRange(2, 1, lastRow - 1, 7).getValues();

  values.forEach(row => {
    const usedAt = row[0];
    const name = String(row[1] || '').trim();
    const operation = String(row[5] || '').trim();

    if (!name) return;
    if (!(usedAt instanceof Date)) return;
    if (usedAt < since) return;
    if (operation !== '画面表示') return;

    const usage = getOrCreateUsageStats_(map, name);

    if (!usage.lastUsedAt || usedAt > usage.lastUsedAt) {
      usage.lastUsedAt = usedAt;
    }

    usage.openCount += 1;
  });
}

function appendRecentActionCountsFromLogs_(map, since) {
  const sheet = getSheet_(SHEET_LOGS);
  const lastRow = sheet.getLastRow();

  if (lastRow < 2) return;

  const values = sheet.getRange(2, 1, lastRow - 1, 5).getValues();

  values.forEach(row => {
    const usedAt = row[0];
    const name = String(row[1] || '').trim();
    const action = normalizeAction_(row[3]);

    if (!name) return;
    if (!action) return;
    if (!(usedAt instanceof Date)) return;
    if (usedAt < since) return;

    const usage = getOrCreateUsageStats_(map, name);

    if (!usage.lastUsedAt || usedAt > usage.lastUsedAt) {
      usage.lastUsedAt = usedAt;
    }

    usage.actionCount += 1;
  });
}

function getOrCreateUsageStats_(map, name) {
  if (!map[name]) {
    map[name] = {
      lastUsedAt: null,
      openCount: 0,
      actionCount: 0
    };
  }

  return map[name];
}

function getUsageJudgement_(usage) {
  const openCount = Number(usage.openCount || 0);
  const actionCount = Number(usage.actionCount || 0);

  if (openCount === 0 && actionCount === 0) {
    return {
      label: '未利用',
      level: 'danger',
      note: `${USAGE_MONITOR_DAYS}日間、アプリ利用ログがありません。`
    };
  }

  if (actionCount < LOW_USAGE_ACTION_THRESHOLD) {
    return {
      label: '操作なし',
      level: 'warn',
      note: '画面表示はありますが、貸出・返却操作がありません。'
    };
  }

  if (openCount + actionCount < LOW_USAGE_OPEN_THRESHOLD) {
    return {
      label: '低頻度',
      level: 'warn',
      note: `${USAGE_MONITOR_DAYS}日間の表示・操作合計が${LOW_USAGE_OPEN_THRESHOLD}回未満です。`
    };
  }

  return {
    label: '利用あり',
    level: 'ok',
    note: ''
  };
}

function rebuildUsageSummarySheet_(rows) {
  const sheet = getSheet_(SHEET_USAGE_SUMMARY);
  sheet.clear();
  sheet.getRange(1, 1, 1, HEADERS.USAGE_SUMMARY.length).setValues([HEADERS.USAGE_SUMMARY]);

  const values = rows.map(row => [
    row.name,
    row.group,
    row.device,
    row.lastUsedAt,
    row.openCount,
    row.actionCount,
    row.judgement,
    row.note
  ]);

  if (values.length > 0) {
    sheet.getRange(2, 1, values.length, HEADERS.USAGE_SUMMARY.length).setValues(values);
  }

  formatSheet_(sheet);
}

// 手動実行用：usage_logs と logs から直近30日分を再集計する
function rebuildUsageSummaryFromLogs() {
  ensureUsageMonitoringSheets_();
  const monitoring = getUsageMonitoringData_();
  rebuildUsageSummarySheet_(monitoring.rows);

  return {
    success: true,
    message: `直近${monitoring.days}日間の利用状況を usage_summary に再集計しました。`,
    summary: monitoring.summary
  };
}

function syncCurrentStatusFromUsers() {
  const userSheet = getSheet_(SHEET_USERS);
  const statusSheet = getSheet_(SHEET_STATUS);
  const userLastRow = userSheet.getLastRow();

  if (userLastRow < 2) return;

  const users = userSheet.getRange(2, 1, userLastRow - 1, 5).getValues();

  users.forEach(row => {
    const name = String(row[0] || '').trim();
    const device = String(row[1] || '').trim();
    const activeStatus = String(row[4] || '').trim();

    if (!name) return;
    if (activeStatus === '無効') return;

    const existingRow = findRowByName_(statusSheet, name);

    if (existingRow === -1) {
      statusSheet.appendRow([
        name,
        device,
        '返却済',
        '',
        '',
        '',
        false
      ]);
    } else {
      // users 側で端末名が変わった場合、current_status 側も更新
      statusSheet.getRange(existingRow, 2).setValue(device);
    }
  });
}

// ==============================
// 通知
// ==============================

// 毎日22:30ごろに、当日貸出されたままの端末をチェックするトリガーを設定
function setDailyBorrowNoticeTrigger() {
  // 古い通知トリガーを削除
  const triggers = ScriptApp.getProjectTriggers();

  triggers.forEach(trigger => {
    const handler = trigger.getHandlerFunction();

    if (
      handler === 'checkOverdue' ||
      handler === 'checkSameDayBorrowingAt2230'
    ) {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  // 毎日22:30ごろに実行
  ScriptApp.newTrigger('checkSameDayBorrowingAt2230')
    .timeBased()
    .everyDays(1)
    .atHour(22)
    .nearMinute(30)
    .inTimezone(TZ)
    .create();

  console.log('毎日22:30ごろの貸出中チェックトリガーを設定しました。');
}


// 貸出した日の22:30時点で、まだ貸出中ならメール送信
function checkSameDayBorrowingAt2230() {
  syncCurrentStatusFromUsers();

  const users = getUsersMap_();
  const groups = getGroupsMap_();

  const statusSheet = getSheet_(SHEET_STATUS);
  const lastRow = statusSheet.getLastRow();

  if (lastRow < 2) return;

  const values = statusSheet.getRange(2, 1, lastRow - 1, 7).getValues();
  const todayText = Utilities.formatDate(new Date(), TZ, 'yyyy/MM/dd');

  values.forEach(row => {
    const name = String(row[0] || '').trim();
    const device = String(row[1] || '').trim();
    const currentStatus = String(row[2] || '').trim();
    const borrowAt = row[3];

    if (!name) return;
    if (currentStatus !== '貸出中') return;

    // 貸出日時が日付として入っていない場合は通知しない
    if (!(borrowAt instanceof Date)) return;

    // 貸出日が今日ではない場合は通知しない
    const borrowDateText = Utilities.formatDate(borrowAt, TZ, 'yyyy/MM/dd');

    if (borrowDateText !== todayText) return;

    const user = users[name];

    if (!user) return;

    const groupName = String(user.group || '').trim();
    const groupInfo = groups[groupName];

    const personalEmail = String(user.personalEmail || '').trim();
    const leaderEmail = groupInfo ? String(groupInfo.leaderEmail || '').trim() : '';
    const leaderName = groupInfo ? String(groupInfo.leaderName || '').trim() : '';

    // 本人メールもリーダーメールもなければ送信しない
    if (!personalEmail && !leaderEmail) return;

    const elapsedText = calculateElapsedText_(borrowAt, '貸出中');

    const subject = '【要確認】タブレット未返却のお知らせ';

    const body =
      'お疲れ様です。\n\n' +
      '本日貸出されたタブレットが、22:30時点で「貸出中」のままになっています。\n' +
      '返却済みの場合は、WEBアプリから返却処理をお願いします。\n\n' +
      '【対象者】\n' +
      `氏名：${name}\n` +
      `グループ：${groupName || '未設定'}\n` +
      `端末：${device || '未登録'}\n` +
      `貸出日時：${formatDateTime_(borrowAt) || '-'}\n` +
      `経過時間：${elapsedText || '-'}\n\n` +
      '【確認者】\n' +
      `リーダー：${leaderName || '未設定'}\n\n` +
      '※このメールはシステムからの自動送信です。';

    // 本人メールがある場合：本人宛、リーダーCC
    if (personalEmail) {
      const options = {};

      if (leaderEmail) {
        options.cc = leaderEmail;
      }

      MailApp.sendEmail(personalEmail, subject, body, options);
      return;
    }

    // 本人メールがない場合：リーダー宛
    if (leaderEmail) {
      const leaderOnlyBody =
        'お疲れ様です。\n\n' +
        '本日貸出されたタブレットが、22:30時点で「貸出中」のままになっているメンバーがいます。\n' +
        '状況の確認をお願いします。\n\n' +
        '【対象者】\n' +
        `氏名：${name}\n` +
        `グループ：${groupName || '未設定'}\n` +
        `端末：${device || '未登録'}\n` +
        `貸出日時：${formatDateTime_(borrowAt) || '-'}\n` +
        `経過時間：${elapsedText || '-'}\n\n` +
        '※このメールはシステムからの自動送信です。';

      MailApp.sendEmail(leaderEmail, subject, leaderOnlyBody);
    }
  });
}

// 毎週月曜9時ごろに、リーダーへ利用監視レポートを送信するトリガーを設定
function setWeeklyUsageMonitoringTrigger() {
  const triggers = ScriptApp.getProjectTriggers();

  triggers.forEach(trigger => {
    const handler = trigger.getHandlerFunction();

    if (handler === 'sendWeeklyUsageMonitoringReport') {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  ScriptApp.newTrigger('sendWeeklyUsageMonitoringReport')
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.MONDAY)
    .atHour(9)
    .nearMinute(0)
    .inTimezone(TZ)
    .create();

  console.log('毎週月曜9時ごろの利用監視レポートトリガーを設定しました。');
}

// 直近30日の未利用・低頻度者をグループリーダーへ通知する
function sendWeeklyUsageMonitoringReport() {
  ensureUsageMonitoringSheets_();

  const monitoring = getUsageMonitoringData_();
  rebuildUsageSummarySheet_(monitoring.rows);

  const groups = getGroupsMap_();
  const targetRows = monitoring.rows.filter(row => row.judgementLevel !== 'ok');

  if (targetRows.length === 0) return;

  const rowsByGroup = {};

  targetRows.forEach(row => {
    const groupName = row.group || '未設定';

    if (!rowsByGroup[groupName]) {
      rowsByGroup[groupName] = [];
    }

    rowsByGroup[groupName].push(row);
  });

  Object.keys(rowsByGroup).forEach(groupName => {
    const group = groups[groupName];
    const leaderEmail = group ? String(group.leaderEmail || '').trim() : '';

    if (!leaderEmail) return;

    const lines = rowsByGroup[groupName].map(row => {
      return [
        `氏名：${row.name}`,
        `端末：${row.device || '未登録'}`,
        `判定：${row.judgement}`,
        `最終利用：${row.lastUsedAt || '-'}`,
        `表示回数：${row.openCount}`,
        `操作回数：${row.actionCount}`,
        `備考：${row.note || '-'}`
      ].join('\n');
    });

    const subject = '【確認依頼】タブレット管理アプリ 利用状況レポート';
    const body =
      'お疲れ様です。\n\n' +
      `直近${monitoring.days}日間で、タブレット管理アプリの未利用・低頻度利用のメンバーがいます。\n` +
      '必要に応じて、貸出・返却登録の運用確認をお願いします。\n\n' +
      `【グループ】${groupName}\n\n` +
      lines.join('\n\n---\n\n') +
      '\n\n※このメールはシステムからの自動送信です。';

    MailApp.sendEmail(leaderEmail, subject, body);
  });
}


// テスト用：通知メールを手動で確認したいときに実行
function testCheckSameDayBorrowingAt2230() {
  checkSameDayBorrowingAt2230();
}


// ==============================
// 共通関数
// ==============================

function getSheet_(sheetName) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);

  if (!sheet) {
    throw new Error('シートが見つかりません: ' + sheetName);
  }

  return sheet;
}

function findUser_(userName) {
  const targetName = String(userName || '').trim();

  if (!targetName) return null;

  const sheet = getSheet_(SHEET_USERS);
  const lastRow = sheet.getLastRow();

  if (lastRow < 2) return null;

  const values = sheet.getRange(2, 1, lastRow - 1, 5).getValues();

  for (let i = 0; i < values.length; i++) {
    const name = String(values[i][0] || '').trim();
    const device = String(values[i][1] || '').trim();
    const personalEmail = String(values[i][2] || '').trim();
    const group = String(values[i][3] || '').trim();
    const activeStatus = String(values[i][4] || '').trim();

    if (name === targetName && activeStatus !== '無効') {
      return {
        name: name,
        device: device || '未登録',
        personalEmail: personalEmail,
        group: group,
        activeStatus: activeStatus || '有効'
      };
    }
  }

  return null;
}

function getUsersMap_() {
  const sheet = getSheet_(SHEET_USERS);
  const lastRow = sheet.getLastRow();
  const map = {};

  if (lastRow < 2) return map;

  const values = sheet.getRange(2, 1, lastRow - 1, 5).getValues();

  values.forEach(row => {
    const name = String(row[0] || '').trim();
    const device = String(row[1] || '').trim();
    const personalEmail = String(row[2] || '').trim();
    const group = String(row[3] || '').trim();
    const activeStatus = String(row[4] || '').trim();

    if (!name) return;
    if (activeStatus === '無効') return;

    map[name] = {
      name: name,
      device: device,
      personalEmail: personalEmail,
      group: group,
      activeStatus: activeStatus || '有効'
    };
  });

  return map;
}

function getGroupsMap_() {
  const sheet = getSheet_(SHEET_GROUPS);
  const lastRow = sheet.getLastRow();
  const map = {};

  if (lastRow < 2) return map;

  const values = sheet.getRange(2, 1, lastRow - 1, 4).getValues();

  values.forEach(row => {
    const group = String(row[0] || '').trim();
    const leaderName = String(row[1] || '').trim();
    const leaderEmail = String(row[2] || '').trim();
    const activeStatus = String(row[3] || '').trim();

    if (!group) return;
    if (activeStatus === '無効') return;

    map[group] = {
      group: group,
      leaderName: leaderName,
      leaderEmail: leaderEmail,
      activeStatus: activeStatus || '有効'
    };
  });

  return map;
}

function getCurrentStatusByName_(name) {
  const sheet = getSheet_(SHEET_STATUS);
  const rowNumber = findRowByName_(sheet, name);

  if (rowNumber === -1) return null;

  const row = sheet.getRange(rowNumber, 1, 1, 7).getValues()[0];

  return {
    name: row[0],
    device: row[1],
    currentStatus: row[2],
    borrowAt: row[3],
    returnAt: row[4],
    updatedAt: row[5],
    overdue: row[6]
  };
}

function createInitialStatusRow_(name, device) {
  const sheet = getSheet_(SHEET_STATUS);

  sheet.appendRow([
    name,
    device,
    '返却済',
    '',
    '',
    '',
    false
  ]);
}

function findRowByName_(sheet, name) {
  const targetName = String(name || '').trim();
  const lastRow = sheet.getLastRow();

  if (lastRow < 2) return -1;

  const values = sheet.getRange(2, 1, lastRow - 1, 1).getValues();

  for (let i = 0; i < values.length; i++) {
    const rowName = String(values[i][0] || '').trim();

    if (rowName === targetName) {
      return i + 2;
    }
  }

  return -1;
}

function formatDateTime_(value) {
  if (!value) return '';

  if (value instanceof Date) {
    return Utilities.formatDate(value, TZ, 'yyyy/MM/dd HH:mm');
  }

  return String(value);
}

function calculateElapsedText_(borrowAt, currentStatus) {
  if (!(borrowAt instanceof Date)) return '';
  if (currentStatus !== '貸出中') return '';

  const now = new Date();
  const diffMs = now.getTime() - borrowAt.getTime();

  if (diffMs < 0) return '';

  const totalMinutes = Math.floor(diffMs / 1000 / 60);
  const hours = Math.floor(totalMinutes / 60);
  const minutes = totalMinutes % 60;

  if (hours <= 0) {
    return `${minutes}分`;
  }

  return `${hours}時間${minutes}分`;
}

// users シートを正として current_status を修復する関数
function rebuildCurrentStatusFromUsers() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const userSheet = ss.getSheetByName(SHEET_USERS);
  const statusSheet = ss.getSheetByName(SHEET_STATUS);

  if (!userSheet) {
    throw new Error('users シートが見つかりません。');
  }

  if (!statusSheet) {
    throw new Error('current_status シートが見つかりません。');
  }

  // 端末番号が日付に変換されないように、B列をテキスト形式にする
  userSheet.getRange('B:B').setNumberFormat('@');
  statusSheet.getRange('B:B').setNumberFormat('@');

  // users のC列・D列も表示崩れ防止
  userSheet.getRange('C:C').setNumberFormat('@');
  userSheet.getRange('D:D').setNumberFormat('@');

  // 既存の current_status を名前ごとに退避
  const oldStatusMap = {};
  const statusLastRow = statusSheet.getLastRow();

  if (statusLastRow >= 2) {
    const oldValues = statusSheet.getRange(2, 1, statusLastRow - 1, 7).getValues();

    oldValues.forEach(row => {
      const name = String(row[0] || '').trim();

      if (!name) return;

      oldStatusMap[name] = {
        currentStatus: row[2] || '返却済',
        borrowAt: row[3] || '',
        returnAt: row[4] || '',
        updatedAt: row[5] || '',
        overdue: row[6] === true || String(row[6]).toUpperCase() === 'TRUE'
      };
    });
  }

  // current_status を一度クリア
  statusSheet.clear();

  // 見出しを再作成
  statusSheet.getRange(1, 1, 1, 7).setValues([[
    '氏名',
    '割当端末',
    '現在状態',
    '貸出日時',
    '返却日時',
    '最終更新',
    '未返却'
  ]]);

  // users の表示値を取得する
  // getDisplayValues() を使うことで、2001-1-1 のような端末名を見た目のまま取得しやすくする
  const userLastRow = userSheet.getLastRow();

  if (userLastRow < 2) {
    formatSheet_(statusSheet);
    return;
  }

  const userDisplayValues = userSheet.getRange(2, 1, userLastRow - 1, 5).getDisplayValues();

  const output = [];

  userDisplayValues.forEach(row => {
    const name = String(row[0] || '').trim();
    const device = String(row[1] || '').trim();
    const activeStatus = String(row[4] || '').trim();

    if (!name) return;
    if (activeStatus === '無効') return;

    const old = oldStatusMap[name];

    if (old) {
      output.push([
        name,
        device,
        old.currentStatus,
        old.borrowAt,
        old.returnAt,
        old.updatedAt,
        old.currentStatus === '貸出中'
      ]);
    } else {
      output.push([
        name,
        device,
        '返却済',
        '',
        '',
        '',
        false
      ]);
    }
  });

  if (output.length > 0) {
    statusSheet.getRange(2, 1, output.length, 7).setValues(output);
  }

  // 見やすく整える
  statusSheet.getRange('A:A').setNumberFormat('@');
  statusSheet.getRange('B:B').setNumberFormat('@');
  statusSheet.getRange('C:C').setNumberFormat('@');
  statusSheet.getRange('D:F').setNumberFormat('yyyy/MM/dd HH:mm');

  formatSheet_(statusSheet);

  console.log('current_status を users 基準で修復しました。');
}

// users / groups の有効・無効プルダウンを設定する関数
function setupActiveStatusDropdowns() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const userSheet = ss.getSheetByName(SHEET_USERS);
  const groupSheet = ss.getSheetByName(SHEET_GROUPS);

  if (!userSheet) {
    throw new Error('users シートが見つかりません。');
  }

  if (!groupSheet) {
    throw new Error('groups シートが見つかりません。');
  }

  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['有効', '無効'], true)
    .setAllowInvalid(false)
    .build();

  // users：E列「通知設定」
  const userMaxRows = userSheet.getMaxRows();
  userSheet.getRange(2, 5, userMaxRows - 1, 1).setDataValidation(rule);

  // groups：D列「有効/無効」
  const groupMaxRows = groupSheet.getMaxRows();
  groupSheet.getRange(2, 4, groupMaxRows - 1, 1).setDataValidation(rule);

  // 空欄のところは、名前やグループが入っている行だけ「有効」にする
  fillDefaultActiveStatusOnlyUsedRows_(userSheet, 1, 5);  // users：A列に氏名がある行だけE列を有効
  fillDefaultActiveStatusOnlyUsedRows_(groupSheet, 1, 4); // groups：A列にグループがある行だけD列を有効

  // 見た目を統一
  normalizeActiveStatusStyle_(userSheet, 5);
  normalizeActiveStatusStyle_(groupSheet, 4);

  console.log('有効/無効のプルダウンと表示形式を設定しました。');
}

// 名前やグループが入っている行だけ、空欄を「有効」にする
function fillDefaultActiveStatusOnlyUsedRows_(sheet, keyColumnNumber, statusColumnNumber) {
  const lastRow = sheet.getLastRow();

  if (lastRow < 2) return;

  const keyValues = sheet.getRange(2, keyColumnNumber, lastRow - 1, 1).getValues();
  const statusRange = sheet.getRange(2, statusColumnNumber, lastRow - 1, 1);
  const statusValues = statusRange.getValues();

  const newValues = statusValues.map((row, index) => {
    const keyValue = String(keyValues[index][0] || '').trim();
    const currentValue = String(row[0] || '').trim();

    // 氏名やグループ名が空なら、通知設定も空欄にする
    if (!keyValue) {
      return [''];
    }

    // 氏名やグループ名があり、通知設定が空なら有効にする
    if (!currentValue) {
      return ['有効'];
    }

    return [currentValue];
  });

  statusRange.setValues(newValues);
}

// 有効/無効列の見た目を統一する
function normalizeActiveStatusStyle_(sheet, statusColumnNumber) {
  const maxRows = sheet.getMaxRows();

  if (maxRows < 2) return;

  const range = sheet.getRange(2, statusColumnNumber, maxRows - 1, 1);

  range
    .setFontColor('#000000')
    .setFontWeight('normal')
    .setBackground('#ffffff')
    .setHorizontalAlignment('center')
    .setNumberFormat('@');
}
