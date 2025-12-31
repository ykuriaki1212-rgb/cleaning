/**
 * お掃除記録簿 - Google Apps Script
 * 
 * このスクリプトは「お掃除記録」HTMLダッシュボードと
 * Googleスプレッドシートを連携させるためのAPIを提供します。
 */

// ========================================
// 設定 - 以下のIDを実際のスプレッドシートIDに置き換えてください
// ========================================
const SPREADSHEET_ID = 'ここにスプレッドシートIDを入力';

// シート名
const SHEET_RECORDS = '掃除記録';
const SHEET_SETTINGS = '設定';
const SHEET_TIPS = 'コツ';
const SHEET_LOG = '更新ログ';

/**
 * 初期化関数 - 最初に一度実行してシートを作成
 */
function initializeSpreadsheet() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  
  // 掃除記録シートの作成
  let recordsSheet = ss.getSheetByName(SHEET_RECORDS);
  if (!recordsSheet) {
    recordsSheet = ss.insertSheet(SHEET_RECORDS);
    recordsSheet.appendRow(['ID', '日付', '場所', '満足度', '時間（分）', 'メモ', '更新日時']);
    recordsSheet.getRange(1, 1, 1, 7).setFontWeight('bold').setBackground('#e8f5e9');
    recordsSheet.setFrozenRows(1);
    recordsSheet.setColumnWidth(3, 200);
    recordsSheet.setColumnWidth(6, 300);
  }
  
  // 設定シートの作成
  let settingsSheet = ss.getSheetByName(SHEET_SETTINGS);
  if (!settingsSheet) {
    settingsSheet = ss.insertSheet(SHEET_SETTINGS);
    settingsSheet.appendRow(['項目', '値']);
    settingsSheet.appendRow(['週間目標', '5']);
    settingsSheet.getRange(1, 1, 1, 2).setFontWeight('bold').setBackground('#e8f5e9');
    settingsSheet.setFrozenRows(1);
  }
  
  // コツシートの作成
  let tipsSheet = ss.getSheetByName(SHEET_TIPS);
  if (!tipsSheet) {
    tipsSheet = ss.insertSheet(SHEET_TIPS);
    tipsSheet.appendRow(['ID', 'アイコン', 'タイトル', '説明']);
    tipsSheet.getRange(1, 1, 1, 4).setFontWeight('bold').setBackground('#e8f5e9');
    tipsSheet.setFrozenRows(1);
    // デフォルトのコツを追加
    tipsSheet.appendRow([1, '⏰', '朝のルーティンに組み込む', '起床後の5分掃除で1日を気持ちよくスタート']);
    tipsSheet.appendRow([2, '🎵', 'お気に入りの音楽をかけて', '好きな曲1曲分だけ掃除すると決めると楽しい']);
    tipsSheet.appendRow([3, '📦', '「1日1捨て」を習慣に', '毎日1つ不要なものを手放すとスッキリ']);
    tipsSheet.appendRow([4, '✨', '「ついで掃除」で効率アップ', '手を洗ったら洗面台を拭く、など習慣化']);
  }
  
  // ログシートの作成
  let logSheet = ss.getSheetByName(SHEET_LOG);
  if (!logSheet) {
    logSheet = ss.insertSheet(SHEET_LOG);
    logSheet.appendRow(['タイムスタンプ', '操作', 'ユーザー', '詳細']);
    logSheet.getRange(1, 1, 1, 4).setFontWeight('bold').setBackground('#fff8e1');
    logSheet.setFrozenRows(1);
  }
  
  // デフォルトシートを削除
  const defaultSheet = ss.getSheetByName('シート1');
  if (defaultSheet) {
    ss.deleteSheet(defaultSheet);
  }
  
  Logger.log('初期化完了！');
  return '初期化が完了しました。';
}

/**
 * テスト用初期化関数
 */
function testInit() {
  const result = initializeSpreadsheet();
  Logger.log(result);
}

/**
 * ログを記録
 */
function addLog(operation, details) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const logSheet = ss.getSheetByName(SHEET_LOG);
  const timestamp = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm:ss');
  const user = Session.getActiveUser().getEmail() || '匿名';
  logSheet.appendRow([timestamp, operation, user, details]);
}

/**
 * Web APIエンドポイント - GET
 */
function doGet(e) {
  try {
    const action = e.parameter.action;
    let result;
    
    switch (action) {
      case 'load':
        result = loadAllData();
        addLog('読込', 'データ読込成功');
        break;
      default:
        result = { error: '不明なアクション' };
    }
    
    return ContentService
      .createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
      
  } catch (error) {
    return ContentService
      .createTextOutput(JSON.stringify({ error: error.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * Web APIエンドポイント - POST
 */
function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const action = data.action;
    let result;
    
    switch (action) {
      case 'save':
        result = saveAllData(data);
        addLog('保存', `記録${data.records ? data.records.length : 0}件を保存`);
        break;
      case 'addRecord':
        result = addRecord(data.record);
        addLog('追加', `${data.record.date}の記録を追加`);
        break;
      case 'updateRecord':
        result = updateRecord(data.record);
        addLog('更新', `ID:${data.record.id}の記録を更新`);
        break;
      case 'deleteRecord':
        result = deleteRecord(data.id);
        addLog('削除', `ID:${data.id}の記録を削除`);
        break;
      case 'saveSettings':
        result = saveSettings(data);
        addLog('設定保存', '設定を更新');
        break;
      case 'saveTips':
        result = saveTips(data.tips);
        addLog('コツ保存', `コツ${data.tips.length}件を保存`);
        break;
      default:
        result = { error: '不明なアクション' };
    }
    
    return ContentService
      .createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
      
  } catch (error) {
    return ContentService
      .createTextOutput(JSON.stringify({ error: error.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * 全データを読み込み
 */
function loadAllData() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  
  // 掃除記録を読み込み
  const recordsSheet = ss.getSheetByName(SHEET_RECORDS);
  const recordsData = recordsSheet.getDataRange().getValues();
  const records = [];
  
  for (let i = 1; i < recordsData.length; i++) {
    const row = recordsData[i];
    if (row[0]) {
      records.push({
        id: row[0],
        date: row[1],
        places: row[2] ? row[2].split(',') : [],
        rating: row[3],
        time: row[4],
        note: row[5] || ''
      });
    }
  }
  
  // 設定を読み込み
  const settingsSheet = ss.getSheetByName(SHEET_SETTINGS);
  const settingsData = settingsSheet.getDataRange().getValues();
  let goal = 5;
  
  for (let i = 1; i < settingsData.length; i++) {
    if (settingsData[i][0] === '週間目標') {
      goal = parseInt(settingsData[i][1]) || 5;
    }
  }
  
  // コツを読み込み
  const tipsSheet = ss.getSheetByName(SHEET_TIPS);
  const tipsData = tipsSheet.getDataRange().getValues();
  const tips = [];
  
  for (let i = 1; i < tipsData.length; i++) {
    const row = tipsData[i];
    if (row[0]) {
      tips.push({
        id: row[0],
        icon: row[1],
        title: row[2],
        desc: row[3] || ''
      });
    }
  }
  
  return {
    success: true,
    records: records,
    goal: goal,
    tips: tips
  };
}

/**
 * 全データを保存
 */
function saveAllData(data) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  
  // 掃除記録を保存
  if (data.records) {
    const recordsSheet = ss.getSheetByName(SHEET_RECORDS);
    // ヘッダー以外をクリア
    const lastRow = recordsSheet.getLastRow();
    if (lastRow > 1) {
      recordsSheet.getRange(2, 1, lastRow - 1, 7).clearContent();
    }
    
    // データを書き込み
    const timestamp = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm:ss');
    data.records.forEach((record, index) => {
      recordsSheet.getRange(index + 2, 1, 1, 7).setValues([[
        record.id,
        record.date,
        record.places.join(','),
        record.rating,
        record.time,
        record.note || '',
        timestamp
      ]]);
    });
  }
  
  // 設定を保存
  if (data.goal !== undefined) {
    const settingsSheet = ss.getSheetByName(SHEET_SETTINGS);
    const settingsData = settingsSheet.getDataRange().getValues();
    
    for (let i = 1; i < settingsData.length; i++) {
      if (settingsData[i][0] === '週間目標') {
        settingsSheet.getRange(i + 1, 2).setValue(data.goal);
        break;
      }
    }
  }
  
  // コツを保存
  if (data.tips) {
    saveTips(data.tips);
  }
  
  return { success: true, message: '保存完了' };
}

/**
 * 単一の記録を追加
 */
function addRecord(record) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const recordsSheet = ss.getSheetByName(SHEET_RECORDS);
  const timestamp = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm:ss');
  
  recordsSheet.appendRow([
    record.id,
    record.date,
    record.places.join(','),
    record.rating,
    record.time,
    record.note || '',
    timestamp
  ]);
  
  return { success: true, message: '記録を追加しました' };
}

/**
 * 記録を更新
 */
function updateRecord(record) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const recordsSheet = ss.getSheetByName(SHEET_RECORDS);
  const data = recordsSheet.getDataRange().getValues();
  const timestamp = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm:ss');
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] == record.id) {
      recordsSheet.getRange(i + 1, 1, 1, 7).setValues([[
        record.id,
        record.date,
        record.places.join(','),
        record.rating,
        record.time,
        record.note || '',
        timestamp
      ]]);
      return { success: true, message: '記録を更新しました' };
    }
  }
  
  return { success: false, message: '記録が見つかりません' };
}

/**
 * 記録を削除
 */
function deleteRecord(id) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const recordsSheet = ss.getSheetByName(SHEET_RECORDS);
  const data = recordsSheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] == id) {
      recordsSheet.deleteRow(i + 1);
      return { success: true, message: '記録を削除しました' };
    }
  }
  
  return { success: false, message: '記録が見つかりません' };
}

/**
 * 設定を保存
 */
function saveSettings(data) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const settingsSheet = ss.getSheetByName(SHEET_SETTINGS);
  const settingsData = settingsSheet.getDataRange().getValues();
  
  if (data.goal !== undefined) {
    for (let i = 1; i < settingsData.length; i++) {
      if (settingsData[i][0] === '週間目標') {
        settingsSheet.getRange(i + 1, 2).setValue(data.goal);
        break;
      }
    }
  }
  
  return { success: true, message: '設定を保存しました' };
}

/**
 * コツを保存
 */
function saveTips(tips) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const tipsSheet = ss.getSheetByName(SHEET_TIPS);
  
  // ヘッダー以外をクリア
  const lastRow = tipsSheet.getLastRow();
  if (lastRow > 1) {
    tipsSheet.getRange(2, 1, lastRow - 1, 4).clearContent();
  }
  
  // データを書き込み
  tips.forEach((tip, index) => {
    tipsSheet.getRange(index + 2, 1, 1, 4).setValues([[
      tip.id,
      tip.icon,
      tip.title,
      tip.desc || ''
    ]]);
  });
  
  return { success: true, message: 'コツを保存しました' };
}

/**
 * テスト関数
 */
function testLoad() {
  const result = loadAllData();
  Logger.log(JSON.stringify(result, null, 2));
}
