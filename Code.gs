/**
 * 交通費記録アプリ — Google Apps Script (バックエンド)
 *
 * セットアップ手順:
 * 1. Google スプレッドシートを新規作成
 * 2. 拡張機能 → Apps Script を開く
 * 3. このファイルの内容をすべて貼り付けて保存
 * 4. デプロイ → 新しいデプロイ → ウェブアプリ
 *    - 実行するユーザー: 自分
 *    - アクセスできるユーザー: 全員
 * 5. デプロイ → URL をコピー
 * 6. アプリの設定画面にその URL を貼り付け
 *
 * ※ スプレッドシートの「交通手段マスタ」シートに行を追加すると
 *   アプリ起動時に自動で反映されます。
 */

// スプレッドシートIDを自動取得（Apps Scriptがバインドされたシート）
function getSpreadsheet() {
  return SpreadsheetApp.getActiveSpreadsheet();
}

// シートを取得（なければ作成）
function getOrCreateSheet(name, headers) {
  const ss = getSpreadsheet();
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.appendRow(headers);
    const headerRange = sheet.getRange(1, 1, 1, headers.length);
    headerRange.setFontWeight('bold');
    headerRange.setBackground('#4285f4');
    headerRange.setFontColor('#ffffff');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

// POST リクエスト処理
function doPost(e) {
  try {
    const payload = JSON.parse(e.postData.contents);

    if (payload.action === 'sync') {
      return syncData(payload);
    }

    if (payload.action === 'syncMaster') {
      return syncMaster(payload);
    }

    return jsonResponse({ status: 'error', message: '不明なアクション' });
  } catch (err) {
    return jsonResponse({ status: 'error', message: err.toString() });
  }
}

// GET リクエスト処理（データ読み込み）
function doGet(e) {
  try {
    const action = (e && e.parameter && e.parameter.action) || 'status';

    if (action === 'status') {
      return jsonResponse({ status: 'ok', message: '接続済み' });
    }

    if (action === 'load') {
      return loadData();
    }

    if (action === 'loadMaster') {
      return loadMaster();
    }

    return jsonResponse({ status: 'error', message: '不明なアクション' });
  } catch (err) {
    return jsonResponse({ status: 'error', message: err.toString() });
  }
}

// ===== 交通手段マスタの読み込み =====
function loadMaster() {
  const masterSheet = getOrCreateSheet('交通手段マスタ', ['区分', 'ID', '名称', '金額', 'カテゴリ']);

  const transports = { go: [], return: [] };
  const categoriesSet = {};
  const masterData = masterSheet.getDataRange().getValues();

  for (let i = 1; i < masterData.length; i++) {
    const [dir, id, name, price, category] = masterData[i];
    if (!dir || !name) continue;
    const key = (dir === '行き') ? 'go' : 'return';
    const cat = category || 'その他';
    transports[key].push({ id: id || (key[0] + i), name: String(name), price: Number(price) || 0, category: cat });
    categoriesSet[cat] = true;
  }

  return jsonResponse({
    status: 'ok',
    transports,
    categories: Object.keys(categoriesSet)
  });
}

// ===== 交通手段マスタの書き込み（初期登録用） =====
function syncMaster(payload) {
  const transports = payload.transports || {};
  const categories = payload.categories || [];

  const masterSheet = getOrCreateSheet('交通手段マスタ', ['区分', 'ID', '名称', '金額', 'カテゴリ']);

  // 既存データをクリア
  const lastRow = masterSheet.getLastRow();
  if (lastRow > 1) {
    masterSheet.getRange(2, 1, lastRow - 1, 5).clearContent();
  }

  const masterData = [];
  (transports.go || []).forEach(t => masterData.push(['行き', t.id, t.name, t.price, t.category || 'その他']));
  (transports.return || []).forEach(t => masterData.push(['帰り', t.id, t.name, t.price, t.category || 'その他']));

  if (masterData.length > 0) {
    masterSheet.getRange(2, 1, masterData.length, 5).setValues(masterData);
    masterSheet.getRange(2, 4, masterData.length, 1).setNumberFormat('#,##0"円"');
  }

  // 列幅調整
  masterSheet.autoResizeColumns(1, 5);

  return jsonResponse({ status: 'ok', message: `${masterData.length}件のマスタを登録しました` });
}

// ===== 記録データ同期 =====
function syncData(payload) {
  const rows = payload.rows || [];
  const transports = payload.transports || {};

  // 記録シート
  const recordSheet = getOrCreateSheet('交通費記録', ['日付', '区分', '交通手段', '金額']);

  const lastRow = recordSheet.getLastRow();
  if (lastRow > 1) {
    recordSheet.getRange(2, 1, lastRow - 1, 4).clearContent();
  }

  if (rows.length > 0) {
    const data = rows.map(r => [r.date, r.direction, r.name, r.price]);
    recordSheet.getRange(2, 1, data.length, 4).setValues(data);
    recordSheet.getRange(2, 4, data.length, 1).setNumberFormat('#,##0"円"');
  }

  // マスタも同期（カテゴリ列付き）
  const masterSheet = getOrCreateSheet('交通手段マスタ', ['区分', 'ID', '名称', '金額', 'カテゴリ']);
  const masterLastRow = masterSheet.getLastRow();
  if (masterLastRow > 1) {
    masterSheet.getRange(2, 1, masterLastRow - 1, 5).clearContent();
  }

  const masterData = [];
  (transports.go || []).forEach(t => masterData.push(['行き', t.id, t.name, t.price, t.category || 'その他']));
  (transports.return || []).forEach(t => masterData.push(['帰り', t.id, t.name, t.price, t.category || 'その他']));

  if (masterData.length > 0) {
    masterSheet.getRange(2, 1, masterData.length, 5).setValues(masterData);
    masterSheet.getRange(2, 4, masterData.length, 1).setNumberFormat('#,##0"円"');
  }

  // 月別集計
  updateMonthlySummary(rows);

  return jsonResponse({ status: 'ok', message: `${rows.length}件の記録を同期しました` });
}

// 月別集計を更新
function updateMonthlySummary(rows) {
  const summarySheet = getOrCreateSheet('月別集計', ['年月', '交通手段', '区分', '回数', '小計']);

  const summary = {};
  rows.forEach(r => {
    const ym = r.date.substring(0, 7);
    const key = `${ym}|${r.name}|${r.direction}`;
    if (!summary[key]) summary[key] = { ym, name: r.name, direction: r.direction, count: 0, total: 0 };
    summary[key].count++;
    summary[key].total += r.price;
  });

  const entries = Object.values(summary).sort((a, b) => a.ym.localeCompare(b.ym) || a.direction.localeCompare(b.direction));

  const lastRow = summarySheet.getLastRow();
  if (lastRow > 1) {
    summarySheet.getRange(2, 1, lastRow - 1, 5).clearContent();
  }

  if (entries.length > 0) {
    const data = entries.map(e => [e.ym, e.name, e.direction, e.count, e.total]);
    summarySheet.getRange(2, 1, data.length, 5).setValues(data);
    summarySheet.getRange(2, 5, data.length, 1).setNumberFormat('#,##0"円"');
  }
}

// データ読み込み（全データ）
function loadData() {
  const recordSheet = getOrCreateSheet('交通費記録', ['日付', '区分', '交通手段', '金額']);
  const masterSheet = getOrCreateSheet('交通手段マスタ', ['区分', 'ID', '名称', '金額', 'カテゴリ']);

  const records = {};
  const recordData = recordSheet.getDataRange().getValues();
  for (let i = 1; i < recordData.length; i++) {
    const [date, direction, name, price] = recordData[i];
    if (!date) continue;
    const dateStr = Utilities.formatDate(new Date(date), 'Asia/Tokyo', 'yyyy-MM-dd');
    if (!records[dateStr]) records[dateStr] = { go: [], return: [] };
    const dir = direction === '行き' ? 'go' : 'return';
    records[dateStr][dir].push({ name, price });
  }

  const transports = { go: [], return: [] };
  const categoriesSet = {};
  const masterData = masterSheet.getDataRange().getValues();
  for (let i = 1; i < masterData.length; i++) {
    const [dir, id, name, price, category] = masterData[i];
    if (!dir) continue;
    const key = dir === '行き' ? 'go' : 'return';
    const cat = category || 'その他';
    transports[key].push({ id, name, price, category: cat });
    categoriesSet[cat] = true;
  }

  return jsonResponse({ status: 'ok', records, transports, categories: Object.keys(categoriesSet) });
}

// JSON レスポンス
function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
