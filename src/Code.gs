/**
 * IT資産管理システム - メインエントリーポイント
 */

// ===== 定数 =====
const SHEETS = {
  ASSETS: '資産台帳',
  CATEGORIES: 'カテゴリ',
  DEPARTMENTS: '部署',
  USERS: 'ユーザー',
  HISTORY: '履歴',
  DATA_SOURCES: '外部データソース'
};

const STATUS_OPTIONS = ['利用中', '在庫', '受取待ち', '返却待ち', '回収連絡済み', '回収済み', 'リース終了', '紛失', '保管'];

const ASSET_HEADERS = [
  '資産番号', '資産名', 'カテゴリ', 'メーカー', 'モデル', 'シリアル番号',
  '購入日', '購入金額', '使用者ID', '使用者名', '使用者メール', '部署', '設置場所', 'ステータス',
  'IPアドレス', 'MACアドレス', 'OS', '備考', '登録日', '更新日'
];

const CATEGORY_HEADERS = ['カテゴリID', 'カテゴリ名'];
const DEPARTMENT_HEADERS = ['部署ID', '部署名'];
const USER_HEADERS = ['ユーザーID', '氏名', 'メールアドレス', '部署', '役職', '電話番号', '在籍状況'];
const HISTORY_HEADERS = ['履歴ID', '資産番号', '変更日時', '変更種別', '変更内容', '変更者'];
const DATA_SOURCE_HEADERS = ['データソースID', 'データソース名', 'シート名', '紐付けキー（資産側）', '紐付けキー（データ側）', '最終更新日時', '更新方法', '備考'];

const DEFAULT_CATEGORIES = [
  ['CAT001', 'ノートPC'],
  ['CAT002', 'デスクトップPC'],
  ['CAT003', 'モニター'],
  ['CAT004', 'サーバー'],
  ['CAT005', 'ネットワーク機器'],
  ['CAT006', 'プリンター'],
  ['CAT007', 'モバイル'],
  ['CAT008', 'タブレット'],
  ['CAT009', '周辺機器'],
  ['CAT010', 'その他']
];

const DEFAULT_DEPARTMENTS = [
  ['DEP001', '情報システム部'],
  ['DEP002', '総務部'],
  ['DEP003', '営業部'],
  ['DEP004', '開発部'],
  ['DEP005', 'マーケティング部'],
  ['DEP006', '人事部'],
  ['DEP007', '経理部'],
  ['DEP008', '企画部'],
  ['DEP009', '品質管理部'],
  ['DEP010', '経営企画室']
];

/**
 * 外部データソースの初期定義
 * 新しい連携先を追加するときはここに行を足し、setup() を再実行するだけでシートが作られる。
 */
const DEFAULT_DATA_SOURCES = [
  {
    id: 'DS001',
    name: 'GWSユーザー',
    sheetName: 'EXT_GWSユーザー',
    assetKey: '使用者メール',
    dataKey: 'メールアドレス',
    method: 'Admin SDK / Apps Script トリガー（日次）',
    note: 'Google Workspace のユーザー一覧。AdminDirectory API で日次取得。',
    headers: ['メールアドレス', '氏名', '組織部門', 'ステータス', '最終ログイン', '作成日', '管理者', '2段階認証', '取得日時']
  },
  {
    id: 'DS002',
    name: 'CrowdStrike Falcon',
    sheetName: 'EXT_Falcon',
    assetKey: 'MACアドレス',
    dataKey: 'MACアドレス',
    method: 'API連携 / CSV手動取込（日次）',
    note: 'Falcon エージェント情報。ホスト名・OS・最終接続で端末の健全性を確認。',
    headers: ['ホストID', 'ホスト名', 'MACアドレス', 'ローカルIP', 'OS', 'OSバージョン', 'エージェントバージョン', '最終接続日時', 'ポリシー', 'グループ', '取得日時']
  }
];

// ===== メニュー・エントリーポイント =====

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('IT資産管理')
    .addItem('アプリを開く', 'openApp')
    .addSeparator()
    .addItem('初期セットアップ', 'setup')
    .addItem('テストデータ生成（200件）', 'generateTestData')
    .addSeparator()
    .addItem('外部データソース一覧', 'showDataSources')
    .addToUi();
}

function doGet() {
  return HtmlService.createTemplateFromFile('html/Index')
    .evaluate()
    .setTitle('IT資産管理システム')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function openApp() {
  const html = HtmlService.createTemplateFromFile('html/Index')
    .evaluate()
    .setWidth(1400)
    .setHeight(900);
  SpreadsheetApp.getUi().showModalDialog(html, 'IT資産管理システム');
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// ===== 初期セットアップ =====

function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // コアシート
  createSheetIfNotExists(ss, SHEETS.ASSETS, ASSET_HEADERS);
  createSheetIfNotExists(ss, SHEETS.CATEGORIES, CATEGORY_HEADERS);
  createSheetIfNotExists(ss, SHEETS.DEPARTMENTS, DEPARTMENT_HEADERS);
  createSheetIfNotExists(ss, SHEETS.USERS, USER_HEADERS);
  createSheetIfNotExists(ss, SHEETS.HISTORY, HISTORY_HEADERS);

  // 外部データソース管理シート
  createSheetIfNotExists(ss, SHEETS.DATA_SOURCES, DATA_SOURCE_HEADERS);

  initializeMasterData(ss);
  initializeDataSources(ss);

  SpreadsheetApp.getUi().alert('初期セットアップが完了しました。');
}

function createSheetIfNotExists(ss, sheetName, headers) {
  let sheet = ss.getSheetByName(sheetName);
  if (sheet) return sheet;

  sheet = ss.insertSheet(sheetName);
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setValues([headers]);
  headerRange.setBackground('#2563EB');
  headerRange.setFontColor('#FFFFFF');
  headerRange.setFontWeight('bold');
  headerRange.setHorizontalAlignment('center');
  sheet.setFrozenRows(1);

  for (let i = 1; i <= headers.length; i++) {
    sheet.autoResizeColumn(i);
  }

  return sheet;
}

function initializeMasterData(ss) {
  const catSheet = ss.getSheetByName(SHEETS.CATEGORIES);
  if (catSheet.getLastRow() <= 1) {
    catSheet.getRange(2, 1, DEFAULT_CATEGORIES.length, 2).setValues(DEFAULT_CATEGORIES);
  }

  const depSheet = ss.getSheetByName(SHEETS.DEPARTMENTS);
  if (depSheet.getLastRow() <= 1) {
    depSheet.getRange(2, 1, DEFAULT_DEPARTMENTS.length, 2).setValues(DEFAULT_DEPARTMENTS);
  }
}

// ===== 外部データソース管理 =====

/**
 * DEFAULT_DATA_SOURCES に定義された外部データソースのシートを自動作成する。
 * 既に存在するシートはスキップ。データソース管理シートにも登録する。
 */
function initializeDataSources(ss) {
  const dsSheet = ss.getSheetByName(SHEETS.DATA_SOURCES);
  const existingIds = new Set();
  if (dsSheet.getLastRow() > 1) {
    const ids = dsSheet.getRange(2, 1, dsSheet.getLastRow() - 1, 1).getValues();
    ids.forEach(r => { if (r[0]) existingIds.add(r[0]); });
  }

  DEFAULT_DATA_SOURCES.forEach(ds => {
    // 外部データシートを作成
    createSheetIfNotExists(ss, ds.sheetName, ds.headers);

    // 外部データシートにはヘッダー色を変えてわかりやすくする
    const extSheet = ss.getSheetByName(ds.sheetName);
    extSheet.getRange(1, 1, 1, ds.headers.length).setBackground('#7c3aed'); // 紫

    // 管理シートに未登録なら追加
    if (!existingIds.has(ds.id)) {
      dsSheet.appendRow([
        ds.id, ds.name, ds.sheetName, ds.assetKey, ds.dataKey, '', ds.method, ds.note
      ]);
      existingIds.add(ds.id);
    }
  });
}

/**
 * 新しい外部データソースを追加するヘルパー関数。
 * スクリプトから呼ぶか、メニューから拡張可能。
 *
 * 使用例:
 *   addExternalDataSource('Jamf', 'EXT_Jamf', ['シリアル番号','デバイス名','OS','最終チェックイン','取得日時'], 'シリアル番号', 'シリアル番号', 'API連携（日次）', 'Jamf Pro のデバイス情報');
 */
function addExternalDataSource(name, sheetName, headers, assetKey, dataKey, method, note) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dsSheet = ss.getSheetByName(SHEETS.DATA_SOURCES);
  if (!dsSheet) throw new Error('外部データソース管理シートが見つかりません。初期セットアップを実行してください。');

  const id = 'DS' + String(dsSheet.getLastRow()).padStart(3, '0');

  // データシート作成
  createSheetIfNotExists(ss, sheetName, headers);
  const extSheet = ss.getSheetByName(sheetName);
  extSheet.getRange(1, 1, 1, headers.length).setBackground('#7c3aed');

  // 管理シートに登録
  dsSheet.appendRow([id, name, sheetName, assetKey, dataKey, '', method || '', note || '']);

  return id;
}

/**
 * 外部データソースシートのデータを全置換する汎用関数。
 * API連携のトリガー関数から呼ぶ想定。
 *
 * 使用例（GWSユーザー取得トリガー）:
 *   function fetchGWSUsers() {
 *     const users = AdminDirectory.Users.list({ customer: 'my_customer', maxResults: 500 }).users;
 *     const rows = users.map(u => [u.primaryEmail, u.name.fullName, u.orgUnitPath, ...]);
 *     refreshExternalData('EXT_GWSユーザー', rows);
 *   }
 */
function refreshExternalData(sheetName, dataRows) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) throw new Error(`シート「${sheetName}」が見つかりません。`);

  const headerCount = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].length;

  // 既存データクリア
  if (sheet.getLastRow() > 1) {
    sheet.getRange(2, 1, sheet.getLastRow() - 1, headerCount).clearContent();
  }

  // 新データ書き込み
  if (dataRows.length > 0) {
    sheet.getRange(2, 1, dataRows.length, dataRows[0].length).setValues(dataRows);
  }

  // 外部データソース管理シートの最終更新日時を更新
  updateDataSourceTimestamp(sheetName);

  return { success: true, count: dataRows.length };
}

/**
 * 外部データソース管理シートの最終更新日時を更新
 */
function updateDataSourceTimestamp(sheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dsSheet = ss.getSheetByName(SHEETS.DATA_SOURCES);
  if (!dsSheet || dsSheet.getLastRow() <= 1) return;

  const data = dsSheet.getRange(2, 1, dsSheet.getLastRow() - 1, DATA_SOURCE_HEADERS.length).getValues();
  for (let i = 0; i < data.length; i++) {
    if (data[i][2] === sheetName) { // シート名列
      dsSheet.getRange(i + 2, 6).setValue(new Date()); // 最終更新日時列
      break;
    }
  }
}

/**
 * 外部データソースから資産に紐づくデータを取得する。
 * Webアプリ側から呼び出し、詳細画面で表示する。
 */
function getLinkedExternalData(assetId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dsSheet = ss.getSheetByName(SHEETS.DATA_SOURCES);
  if (!dsSheet || dsSheet.getLastRow() <= 1) return [];

  // 対象資産を取得
  const assetSheet = ss.getSheetByName(SHEETS.ASSETS);
  if (!assetSheet || assetSheet.getLastRow() <= 1) return [];
  const assetData = assetSheet.getRange(2, 1, assetSheet.getLastRow() - 1, ASSET_HEADERS.length).getValues();
  let asset = null;
  for (let i = 0; i < assetData.length; i++) {
    if (assetData[i][0] === assetId) { asset = assetData[i]; break; }
  }
  if (!asset) return [];

  const assetObj = {};
  ASSET_HEADERS.forEach((h, i) => assetObj[h] = asset[i]);

  // 各データソースを走査
  const sources = dsSheet.getRange(2, 1, dsSheet.getLastRow() - 1, DATA_SOURCE_HEADERS.length).getValues();
  const results = [];

  sources.forEach(src => {
    const sourceName = src[1];
    const extSheetName = src[2];
    const assetKeyName = src[3]; // 資産側の紐付けキー列名
    const dataKeyName = src[4]; // データ側の紐付けキー列名
    const lastUpdated = src[5];

    const assetKeyValue = String(assetObj[assetKeyName] || '').trim();
    if (!assetKeyValue) return; // 紐付けキーが空なら skip

    const extSheet = ss.getSheetByName(extSheetName);
    if (!extSheet || extSheet.getLastRow() <= 1) return;

    const extHeaders = extSheet.getRange(1, 1, 1, extSheet.getLastColumn()).getValues()[0];
    const dataKeyIdx = extHeaders.indexOf(dataKeyName);
    if (dataKeyIdx === -1) return;

    const extData = extSheet.getRange(2, 1, extSheet.getLastRow() - 1, extHeaders.length).getValues();
    const matched = extData.filter(row => String(row[dataKeyIdx]).trim() === assetKeyValue);

    if (matched.length > 0) {
      results.push({
        sourceName: sourceName,
        lastUpdated: lastUpdated ? Utilities.formatDate(new Date(lastUpdated), 'Asia/Tokyo', 'yyyy-MM-dd HH:mm') : '未取得',
        headers: extHeaders,
        rows: matched.map(row => row.map(cell => {
          if (cell instanceof Date) return Utilities.formatDate(cell, 'Asia/Tokyo', 'yyyy-MM-dd HH:mm');
          return String(cell);
        }))
      });
    }
  });

  return results;
}

/**
 * メニューから外部データソース一覧を表示
 */
function showDataSources() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dsSheet = ss.getSheetByName(SHEETS.DATA_SOURCES);
  if (!dsSheet) {
    SpreadsheetApp.getUi().alert('外部データソース管理シートが見つかりません。初期セットアップを実行してください。');
    return;
  }
  ss.setActiveSheet(dsSheet);
}
