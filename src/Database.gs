/**
 * IT資産管理システム - データベース操作
 */

// ===== 読み取り =====

function getAssets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEETS.ASSETS);
  if (!sheet || sheet.getLastRow() <= 1) return [];

  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, ASSET_HEADERS.length).getValues();
  return data.map(row => rowToAssetObject(row));
}

function getCategories() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEETS.CATEGORIES);
  if (!sheet || sheet.getLastRow() <= 1) return [];

  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getValues();
  return data.map(row => ({ id: row[0], name: row[1] }));
}

function getDepartments() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEETS.DEPARTMENTS);
  if (!sheet || sheet.getLastRow() <= 1) return [];

  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getValues();
  return data.map(row => ({ id: row[0], name: row[1] }));
}

function getUsers() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEETS.USERS);
  if (!sheet || sheet.getLastRow() <= 1) return [];

  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, USER_HEADERS.length).getValues();
  return data.map(row => ({
    id: row[0],
    name: row[1],
    email: row[2],
    department: row[3],
    title: row[4],
    phone: row[5],
    status: row[6]
  }));
}

function getHistory(assetId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEETS.HISTORY);
  if (!sheet || sheet.getLastRow() <= 1) return [];

  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, HISTORY_HEADERS.length).getValues();
  let history = data.map(row => ({
    historyId: row[0],
    assetId: row[1],
    dateTime: formatDateTimeValue(row[2]),
    changeType: row[3],
    changeDescription: row[4],
    changedBy: row[5]
  }));

  if (assetId) {
    history = history.filter(h => h.assetId === assetId);
  }

  history.sort((a, b) => b.dateTime.localeCompare(a.dateTime));
  return history;
}

function getDashboardData() {
  const assets = getAssets();
  const total = assets.length;

  const statusCount = {};
  STATUS_OPTIONS.forEach(s => statusCount[s] = 0);
  assets.forEach(a => {
    if (statusCount[a.status] !== undefined) {
      statusCount[a.status]++;
    }
  });

  const categoryCount = {};
  assets.forEach(a => {
    categoryCount[a.category] = (categoryCount[a.category] || 0) + 1;
  });

  const departmentCount = {};
  assets.forEach(a => {
    if (a.department) {
      departmentCount[a.department] = (departmentCount[a.department] || 0) + 1;
    }
  });

  const totalCost = assets.reduce((sum, a) => sum + (Number(a.price) || 0), 0);
  const recentHistory = getHistory().slice(0, 10);

  // 期限切れ間近（リース終了・保証期限の30日以内）
  const today = new Date();
  today.setHours(0, 0, 0, 0);
  const in30 = new Date(today);
  in30.setDate(in30.getDate() + 30);
  const expiringSoon = assets.filter(a => {
    const lease = parseDateSafe(a.leaseEndDate);
    const warranty = parseDateSafe(a.warrantyEndDate);
    return (lease && lease <= in30 && lease >= today) || (warranty && warranty <= in30 && warranty >= today);
  });

  return {
    total,
    statusCount,
    categoryCount,
    departmentCount,
    totalCost,
    recentHistory,
    expiringSoon
  };
}

// ===== 書き込み =====

function addAsset(asset) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEETS.ASSETS);
  const now = new Date();
  const id = generateId('AST');

  const row = [
    id,
    asset.name || '',
    asset.category || '',
    asset.manufacturer || '',
    asset.model || '',
    asset.serial || '',
    asset.purchaseDate || '',
    asset.price || '',
    asset.leaseEndDate || '',
    asset.warrantyEndDate || '',
    asset.returnDueDate || '',
    asset.userId || '',
    asset.userName || '',
    asset.userEmail || '',
    asset.department || '',
    asset.location || '',
    asset.status || '在庫',
    asset.ip || '',
    asset.mac || '',
    asset.os || '',
    asset.notes || '',
    '',
    '',
    now,
    now
  ];

  sheet.appendRow(row);
  addHistoryEntry(id, '新規登録', `資産「${asset.name}」を登録`);

  return { success: true, id: id };
}

function updateAsset(asset) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEETS.ASSETS);
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, ASSET_HEADERS.length).getValues();

  let rowIndex = -1;
  for (let i = 0; i < data.length; i++) {
    if (data[i][0] === asset.id) {
      rowIndex = i + 2;
      break;
    }
  }

  if (rowIndex === -1) {
    return { success: false, message: '資産が見つかりません' };
  }

  const oldRow = data[rowIndex - 2];
  const oldAsset = rowToAssetObject(oldRow);
  const now = new Date();

  const changes = [];
  if (oldAsset.name !== asset.name) changes.push(`資産名: ${oldAsset.name} → ${asset.name}`);
  if (oldAsset.category !== asset.category) changes.push(`カテゴリ: ${oldAsset.category} → ${asset.category}`);
  if (oldAsset.manufacturer !== asset.manufacturer) changes.push(`メーカー: ${oldAsset.manufacturer} → ${asset.manufacturer}`);
  if (oldAsset.model !== asset.model) changes.push(`モデル: ${oldAsset.model} → ${asset.model}`);
  if (oldAsset.serial !== asset.serial) changes.push(`シリアル: ${oldAsset.serial} → ${asset.serial}`);
  if (oldAsset.userName !== asset.userName) changes.push(`使用者: ${oldAsset.userName} → ${asset.userName}`);
  if (oldAsset.department !== asset.department) changes.push(`部署: ${oldAsset.department} → ${asset.department}`);
  if (oldAsset.location !== asset.location) changes.push(`設置場所: ${oldAsset.location} → ${asset.location}`);
  if (oldAsset.status !== asset.status) changes.push(`ステータス: ${oldAsset.status} → ${asset.status}`);
  if (oldAsset.ip !== asset.ip) changes.push(`IP: ${oldAsset.ip} → ${asset.ip}`);
  if (oldAsset.mac !== asset.mac) changes.push(`MAC: ${oldAsset.mac} → ${asset.mac}`);
  if (oldAsset.os !== asset.os) changes.push(`OS: ${oldAsset.os} → ${asset.os}`);

  const newRow = [
    asset.id,
    asset.name || '',
    asset.category || '',
    asset.manufacturer || '',
    asset.model || '',
    asset.serial || '',
    asset.purchaseDate || '',
    asset.price || '',
    asset.leaseEndDate || '',
    asset.warrantyEndDate || '',
    asset.returnDueDate || '',
    asset.userId || '',
    asset.userName || '',
    asset.userEmail || '',
    asset.department || '',
    asset.location || '',
    asset.status || '',
    asset.ip || '',
    asset.mac || '',
    asset.os || '',
    asset.notes || '',
    oldRow[21] || '', // 棚卸し確認日はそのまま
    oldRow[22] || '', // 棚卸しSlackURLはそのまま
    oldRow[23], // 登録日はそのまま
    now
  ];

  sheet.getRange(rowIndex, 1, 1, ASSET_HEADERS.length).setValues([newRow]);

  if (changes.length > 0) {
    addHistoryEntry(asset.id, '更新', changes.join(' / '));
  }

  // ステータス変更時: Slack通知
  if (oldAsset.status !== asset.status) {
    try { notifyStatusChange(asset.id, asset.name, oldAsset.status, asset.status, asset.userName, asset.userEmail); } catch (e) {}
  }

  // 使用者変更時: 貸出履歴に記録（新使用者がいる場合は貸出、前の使用者名を「貸出元」に）
  if (oldAsset.userName !== asset.userName && asset.userName) {
    addLendingRecord(asset.id, asset.userId, asset.userName, asset.userEmail, oldAsset.userName || '', '');
  }
  // 使用者が外れた場合: 未返却の貸出履歴があれば返却日を記入
  if (oldAsset.userName && !asset.userName) {
    const list = getLendingHistory(asset.id);
    const openRecord = list.find(h => !h.returnedAt);
    if (openRecord) closeLendingRecord(openRecord.historyId);
  }

  return { success: true };
}

function deleteAsset(assetId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEETS.ASSETS);
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, ASSET_HEADERS.length).getValues();

  for (let i = data.length - 1; i >= 0; i--) {
    if (data[i][0] === assetId) {
      const assetName = data[i][1];
      sheet.deleteRow(i + 2);
      addHistoryEntry(assetId, '削除', `資産「${assetName}」を削除`);
      return { success: true };
    }
  }

  return { success: false, message: '資産が見つかりません' };
}

// ===== ユーザーCRUD =====

function addUser(user) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEETS.USERS);
  const id = generateId('USR');

  sheet.appendRow([
    id,
    user.name || '',
    user.email || '',
    user.department || '',
    user.title || '',
    user.phone || '',
    user.status || '在籍'
  ]);

  return { success: true, id: id };
}

function updateUser(user) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEETS.USERS);
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, USER_HEADERS.length).getValues();

  for (let i = 0; i < data.length; i++) {
    if (data[i][0] === user.id) {
      const row = [
        user.id,
        user.name || '',
        user.email || '',
        user.department || '',
        user.title || '',
        user.phone || '',
        user.status || '在籍'
      ];
      sheet.getRange(i + 2, 1, 1, USER_HEADERS.length).setValues([row]);
      return { success: true };
    }
  }

  return { success: false, message: 'ユーザーが見つかりません' };
}

function deleteUser(userId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEETS.USERS);
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, USER_HEADERS.length).getValues();

  for (let i = data.length - 1; i >= 0; i--) {
    if (data[i][0] === userId) {
      sheet.deleteRow(i + 2);
      return { success: true };
    }
  }

  return { success: false, message: 'ユーザーが見つかりません' };
}

// ===== 履歴 =====

function addHistoryEntry(assetId, changeType, changeDescription) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEETS.HISTORY);
  const historyId = Utilities.getUuid();
  const now = new Date();
  let email = '';
  try {
    email = Session.getActiveUser().getEmail();
  } catch (e) {
    email = '不明';
  }

  sheet.appendRow([historyId, assetId, now, changeType, changeDescription, email]);
}

// ===== エクスポート =====

function exportAssets() {
  const assets = getAssets();
  const rows = [ASSET_HEADERS];
  assets.forEach(a => {
    rows.push([
      a.id, a.name, a.category, a.manufacturer, a.model, a.serial,
      a.purchaseDate, a.price, a.leaseEndDate, a.warrantyEndDate, a.returnDueDate,
      a.userId, a.userName, a.userEmail, a.department, a.location, a.status,
      a.ip, a.mac, a.os, a.notes,
      a.inventoryCheckedDate, a.inventorySlackUrl,
      a.createdDate, a.updatedDate
    ]);
  });
  return rows;
}

// ===== 一括操作 =====

function bulkUpdateStatus(assetIds, newStatus) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEETS.ASSETS);
  if (!sheet || sheet.getLastRow() <= 1) return { success: false, message: 'データなし' };

  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, ASSET_HEADERS.length).getValues();
  let updated = 0;

  const statusCol = ASSET_HEADERS.indexOf('ステータス') + 1;
  const updatedCol = ASSET_HEADERS.indexOf('更新日') + 1;
  const idSet = new Set(assetIds);
  for (let i = 0; i < data.length; i++) {
    if (idSet.has(data[i][0])) {
      const oldStatus = data[i][statusCol - 1];
      if (oldStatus !== newStatus) {
        sheet.getRange(i + 2, statusCol).setValue(newStatus);
        sheet.getRange(i + 2, updatedCol).setValue(new Date());
        addHistoryEntry(data[i][0], 'ステータス変更', `ステータス: ${oldStatus} → ${newStatus}（一括変更）`);
        const userNameCol = ASSET_HEADERS.indexOf('使用者名');
        const userEmailCol = ASSET_HEADERS.indexOf('使用者メール');
        try { notifyStatusChange(data[i][0], data[i][1], oldStatus, newStatus, data[i][userNameCol], data[i][userEmailCol]); } catch (e) {}
        updated++;
      }
    }
  }

  return { success: true, updated: updated };
}

function bulkDeleteAssets(assetIds) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEETS.ASSETS);
  if (!sheet || sheet.getLastRow() <= 1) return { success: false, message: 'データなし' };

  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, ASSET_HEADERS.length).getValues();
  const idSet = new Set(assetIds);
  let deleted = 0;

  // 下の行から削除（インデックスずれ防止）
  for (let i = data.length - 1; i >= 0; i--) {
    if (idSet.has(data[i][0])) {
      addHistoryEntry(data[i][0], '削除', `資産「${data[i][1]}」を削除（一括削除）`);
      sheet.deleteRow(i + 2);
      deleted++;
    }
  }

  return { success: true, deleted: deleted };
}

// ===== インポート =====

function importAssets(assets, skipDuplicate) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEETS.ASSETS);
  const now = new Date();

  // 既存IDの取得
  let existingIds = new Set();
  if (skipDuplicate && sheet.getLastRow() > 1) {
    const ids = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues();
    ids.forEach(r => { if (r[0]) existingIds.add(String(r[0])); });
  }

  const rows = [];
  let skipped = 0;

  assets.forEach(a => {
    const id = a.id || generateId('AST');

    if (skipDuplicate && a.id && existingIds.has(String(a.id))) {
      skipped++;
      return;
    }

    rows.push([
      id,
      a.name || '',
      a.category || '',
      a.manufacturer || '',
      a.model || '',
      a.serial || '',
      a.purchaseDate || '',
      a.price || '',
      a.leaseEndDate || '',
      a.warrantyEndDate || '',
      a.returnDueDate || '',
      a.userId || '',
      a.userName || '',
      a.userEmail || '',
      a.department || '',
      a.location || '',
      a.status || '在庫',
      a.ip || '',
      a.mac || '',
      a.os || '',
      a.notes || '',
      '',
      '',
      now,
      now
    ]);
  });

  if (rows.length > 0) {
    const startRow = sheet.getLastRow() + 1;
    sheet.getRange(startRow, 1, rows.length, ASSET_HEADERS.length).setValues(rows);
    addHistoryEntry('IMPORT', 'インポート', `${rows.length}件の資産データをインポート`);
  }

  return { success: true, imported: rows.length, skipped: skipped };
}

function importUsers(users, skipDuplicate) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEETS.USERS);

  // 既存IDの取得
  let existingIds = new Set();
  if (skipDuplicate && sheet.getLastRow() > 1) {
    const ids = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues();
    ids.forEach(r => { if (r[0]) existingIds.add(String(r[0])); });
  }

  const rows = [];
  let skipped = 0;

  users.forEach(u => {
    const id = u.id || generateId('USR');

    if (skipDuplicate && u.id && existingIds.has(String(u.id))) {
      skipped++;
      return;
    }

    rows.push([
      id,
      u.name || '',
      u.email || '',
      u.department || '',
      u.title || '',
      u.phone || '',
      u.status || '在籍'
    ]);
  });

  if (rows.length > 0) {
    sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, USER_HEADERS.length).setValues(rows);
  }

  return { success: true, imported: rows.length, skipped: skipped };
}

// ===== 貸出履歴 =====

function getLendingHistory(assetId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEETS.LENDING);
  if (!sheet || sheet.getLastRow() <= 1) return [];

  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, LENDING_HEADERS.length).getValues();
  let list = data.map(row => ({
    historyId: row[0],
    assetId: row[1],
    lentAt: formatDateTimeValue(row[2]),
    returnedAt: formatDateTimeValue(row[3]),
    lentToUserId: row[4],
    lentToName: row[5],
    lentToEmail: row[6],
    lentFromName: row[7],
    notes: row[8],
    createdAt: formatDateTimeValue(row[9])
  }));
  if (assetId) list = list.filter(h => h.assetId === assetId);
  list.sort((a, b) => (b.lentAt || '').localeCompare(a.lentAt || ''));
  return list;
}

function addLendingRecord(assetId, lentToUserId, lentToName, lentToEmail, lentFromName, notes) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEETS.LENDING);
  if (!sheet) return { success: false, message: '貸出履歴シートがありません' };
  const id = Utilities.getUuid();
  const now = new Date();
  sheet.appendRow([id, assetId, now, '', lentToUserId || '', lentToName || '', lentToEmail || '', lentFromName || '', notes || '', now]);
  return { success: true, id };
}

function closeLendingRecord(historyId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEETS.LENDING);
  if (!sheet || sheet.getLastRow() <= 1) return { success: false };
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, LENDING_HEADERS.length).getValues();
  for (let i = 0; i < data.length; i++) {
    if (data[i][0] === historyId) {
      sheet.getRange(i + 2, 4).setValue(new Date()); // 返却日
      return { success: true };
    }
  }
  return { success: false };
}

// ===== 棚卸し =====

function getActiveInventoryEvent() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEETS.INVENTORY_EVENTS);
  if (!sheet || sheet.getLastRow() <= 1) return null;
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, INVENTORY_EVENT_HEADERS.length).getValues();
  const active = data.find(row => row[4] === '実施中');
  if (!active) return null;
  return { id: active[0], name: active[1], startDate: formatDateValue(active[2]), endDate: formatDateValue(active[3]), status: active[4] };
}

function startInventoryEvent(eventName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const evSheet = ss.getSheetByName(SHEETS.INVENTORY_EVENTS);
  const assetSheet = ss.getSheetByName(SHEETS.ASSETS);
  if (!evSheet) return { success: false, message: '棚卸しイベントシートがありません' };
  const id = 'INV-' + Utilities.getUuid().slice(0, 8);
  const now = new Date();
  evSheet.appendRow([id, eventName || '棚卸し', now, '', '実施中', now]);
  // 全資産の棚卸し確認フラグをクリア
  if (assetSheet && assetSheet.getLastRow() > 1) {
    const checkedCol = ASSET_HEADERS.indexOf('棚卸し確認日') + 1;
    const urlCol = ASSET_HEADERS.indexOf('棚卸しSlackURL') + 1;
    const lastRow = assetSheet.getLastRow();
    if (checkedCol && urlCol) {
      assetSheet.getRange(2, checkedCol, lastRow - 1, 1).clearContent();
      assetSheet.getRange(2, urlCol, lastRow - 1, 1).clearContent();
    }
  }
  return { success: true, id };
}

function markAssetInventoryChecked(assetId, slackUrl) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEETS.ASSETS);
  if (!sheet || sheet.getLastRow() <= 1) return { success: false };
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, ASSET_HEADERS.length).getValues();
  const checkedCol = ASSET_HEADERS.indexOf('棚卸し確認日') + 1;
  const urlCol = ASSET_HEADERS.indexOf('棚卸しSlackURL') + 1;
  const now = new Date();
  for (let i = 0; i < data.length; i++) {
    if (data[i][0] === assetId) {
      sheet.getRange(i + 2, checkedCol).setValue(now);
      if (slackUrl) sheet.getRange(i + 2, urlCol).setValue(slackUrl);
      return { success: true };
    }
  }
  return { success: false };
}

function getUncheckedAssetsForInventory() {
  const assets = getAssets();
  const active = getActiveInventoryEvent();
  if (!active) return { list: [], event: null };
  return {
    event: active,
    list: assets.filter(a => !a.inventoryCheckedDate || a.inventoryCheckedDate === '')
  };
}

// ===== ユーティリティ =====

function rowToAssetObject(row) {
  return {
    id: row[0],
    name: row[1],
    category: row[2],
    manufacturer: row[3],
    model: row[4],
    serial: row[5],
    purchaseDate: formatDateValue(row[6]),
    price: row[7],
    leaseEndDate: formatDateValue(row[8]),
    warrantyEndDate: formatDateValue(row[9]),
    returnDueDate: formatDateValue(row[10]),
    userId: row[11],
    userName: row[12],
    userEmail: row[13],
    department: row[14],
    location: row[15],
    status: row[16],
    ip: row[17],
    mac: row[18],
    os: row[19],
    notes: row[20],
    inventoryCheckedDate: formatDateValue(row[21]),
    inventorySlackUrl: row[22],
    createdDate: formatDateValue(row[23]),
    updatedDate: formatDateValue(row[24])
  };
}

function generateId(prefix) {
  const now = new Date();
  const timestamp = Utilities.formatDate(now, 'Asia/Tokyo', 'yyyyMMddHHmmss');
  const random = Math.random().toString(36).substring(2, 6).toUpperCase();
  return `${prefix}-${timestamp}-${random}`;
}

function formatDateValue(value) {
  if (!value) return '';
  if (value instanceof Date) {
    return Utilities.formatDate(value, 'Asia/Tokyo', 'yyyy-MM-dd');
  }
  return String(value);
}

function formatDateTimeValue(value) {
  if (!value) return '';
  if (value instanceof Date) {
    return Utilities.formatDate(value, 'Asia/Tokyo', 'yyyy-MM-dd HH:mm:ss');
  }
  return String(value);
}

/** 日付文字列またはDateをDateに変換。無効ならnull */
function parseDateSafe(value) {
  if (!value) return null;
  if (value instanceof Date) return isNaN(value.getTime()) ? null : value;
  const d = new Date(value);
  return isNaN(d.getTime()) ? null : d;
}
