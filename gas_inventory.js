// ============================================================
// PANPANTUTU 在庫管理システム — Google Apps Script
// シンプル版：DriveのCSVを読んでそのまま返す
// ============================================================

const _props = PropertiesService.getScriptProperties();
const SPREADSHEET_ID = _props.getProperty('SPREADSHEET_ID');

const FOLDER_IDS = {
  smaregi:     _props.getProperty('FOLDER_SMAREGI'),
  zozo:        _props.getProperty('FOLDER_ZOZO'),
  zozo_reserve:_props.getProperty('FOLDER_ZOZO_RESERVE'),
  rakuten:     _props.getProperty('FOLDER_RAKUTEN'),
  shipments:   _props.getProperty('FOLDER_SHIPMENTS'),
};

// ════════════════════════════════════════════════════════
// Web API
// ════════════════════════════════════════════════════════
function doGet(e) {
  const action = (e.parameter && e.parameter.action) || '';
  try {
    let data;
    switch (action) {
      case 'ping':
        data = { status: 'ok', time: new Date().toISOString() };
        break;
      case 'getData':
        data = getData_(
          parseInt(e.parameter.offset || '0'),
          parseInt(e.parameter.limit  || '99999')
        );
        break;
      case 'getMeta':
        data = getMeta_();
        break;
      default:
        data = { status: 'ok', time: new Date().toISOString() };
    }
    return respond_({ ok: true, data });
  } catch (err) {
    return respond_({ ok: false, error: err.message });
  }
}

function doPost(e) {
  try {
    const body   = JSON.parse(e.postData.contents);
    const action = body.action;
    let data;
    switch (action) {
      case 'saveTransit':
        data = saveSheet_(SHEET_NAMES.transit, body.rows);
        break;
      case 'saveOrders':
        data = saveSheet_(SHEET_NAMES.orders, body.rows);
        break;
      default:
        return respond_({ ok: false, error: '不明なアクション: ' + action });
    }
    return respond_({ ok: true, data });
  } catch (err) {
    return respond_({ ok: false, error: err.message });
  }
}

function respond_(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

// ════════════════════════════════════════════════════════
// データ取得：DriveのCSVを直接読んで返す
// ════════════════════════════════════════════════════════
function getData_(offset, limit) {
  const result = { products: [], meta: [], transit: [], orders: [], total: 0, offset, limit };

  // スマレジCSVを取得（UTF-8）
  const smFile = getLatestFile_(FOLDER_IDS.smaregi);
  if (smFile) {
    const text = readCsvAutoEncoding_(smFile.getBlob());
    const { headers, rows } = parseCSV_(text);

    result.total = rows.length;
    const sliced = rows.slice(offset, offset + limit);

    // スマレジCSVの列を直接マッピング
    const hi = makeHeaderIndex_(headers);
    result.products = sliced.map(row => ({
      itemNo:     getCellStr_(row, hi, '品番'),
      jan:        getCellStr_(row, hi, '商品コード'),
      name:       getCellStr_(row, hi, '商品名'),
      color:      getCellStr_(row, hi, 'カラー'),
      size:       getCellStr_(row, hi, 'サイズ'),
      price:      getCellNum_(row, hi, '商品単価'),
      // 在庫数
      warehouse:  getCellNum_(row, hi, 'ゼネラルサービス盛岡(panpantutu)の在庫数'),
      shopify:    getCellNum_(row, hi, 'PANPANTUTU ONLINEの在庫数'),
      daikanyama: getCellNum_(row, hi, 'パンパンチュチュ代官山の在庫数'),
      nagoya:     getCellNum_(row, hi, 'パンパンチュチュ名古屋の在庫数'),
      kobe:       getCellNum_(row, hi, 'パンパンチュチュ神戸の在庫数'),
      bazar:      getCellNum_(row, hi, 'bazar et panpantutuの在庫数'),
      linegift:   getCellNum_(row, hi, 'LINEギフトの在庫数'),
      zozo:       0,
      amazon:     0,
      rakuten:    0,
      // 在庫金額（税抜・スマレジ計算値をそのまま使用）
      warehouse_amount:  getCellNum_(row, hi, 'ゼネラルサービス盛岡(panpantutu)の在庫金額'),
      shopify_amount:    getCellNum_(row, hi, 'PANPANTUTU ONLINEの在庫金額'),
      daikanyama_amount: getCellNum_(row, hi, 'パンパンチュチュ代官山の在庫金額'),
      nagoya_amount:     getCellNum_(row, hi, 'パンパンチュチュ名古屋の在庫金額'),
      kobe_amount:       getCellNum_(row, hi, 'パンパンチュチュ神戸の在庫金額'),
      bazar_amount:      getCellNum_(row, hi, 'bazar et panpantutuの在庫金額'),
      linegift_amount:   getCellNum_(row, hi, 'LINEギフトの在庫金額'),
    }));
  }

  // ZOZO在庫を取得してマージ（Shift-JIS固定）
  const zozoFile = getLatestFile_(FOLDER_IDS.zozo);
  if (zozoFile && result.products.length > 0) {
    try {
      const text = readShiftJIS_(zozoFile.getBlob());
      const { headers, rows } = parseCSV_(text);
      const hi = makeHeaderIndex_(headers);
      // CS品番 → zozo在庫 のマップを作成
      // 「販売タイプ」が「予約」の場合は在庫数を0として扱う
      const zozoMap = {};
      const zozoReserveMap = {}; // 予約数も別途保持
      rows.forEach(row => {
        const sku         = getCellStr_(row, hi, 'CS品番');
        const stock       = getCellNum_(row, hi, '販売可能数合計');
        const salesType   = getCellStr_(row, hi, '販売タイプ');
        if (!sku) return;
        if (salesType === '予約') {
          // 予約は実在庫0、予約数として別保持
          zozoMap[sku] = zozoMap[sku] || 0;
          zozoReserveMap[sku] = (zozoReserveMap[sku] || 0) + stock;
        } else {
          // 通常は販売可能数合計をそのまま使用
          zozoMap[sku] = (zozoMap[sku] || 0) + stock;
        }
      });
      // 品番でマッチング（実在庫 + 予約数を別々にセット）
      result.products.forEach(p => {
        if (p.itemNo && zozoMap[p.itemNo] !== undefined) {
          p.zozo         = zozoMap[p.itemNo];
          p.zozo_reserve = zozoReserveMap[p.itemNo] || 0;
        }
      });
    } catch(e) { console.log('ZOZO取得エラー:', e.message); }
  }

  // 楽天在庫を取得してマージ（Shift-JIS固定、1行目は説明行なのでスキップ）
  const rakutenFile = getLatestFile_(FOLDER_IDS.rakuten);
  if (rakutenFile && result.products.length > 0) {
    try {
      const text = readShiftJIS_(rakutenFile.getBlob());
      const { headers, rows } = parseCSV_(text, true);
      const hi = makeHeaderIndex_(headers);
      const rakutenMap = {};
      rows.forEach(row => {
        const jan   = getCellStr_(row, hi, 'JANコード');
        const stock = getCellNum_(row, hi, '受注可能倉庫在庫');
        if (jan) rakutenMap[jan] = (rakutenMap[jan] || 0) + stock;
      });
      result.products.forEach(p => {
        if (p.jan && rakutenMap[p.jan] !== undefined) {
          p.rakuten = rakutenMap[p.jan];
        }
      });
    } catch(e) { console.log('楽天取得エラー:', e.message); }
  }

  // 在途在庫・工場発注
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheets = Object.fromEntries(ss.getSheets().map(s => [s.getName(), s]));
  if (sheets[SHEET_NAMES.meta])    result.meta    = sheetToObjects_(sheets[SHEET_NAMES.meta]);
  if (sheets[SHEET_NAMES.transit]) result.transit = sheetToObjects_(sheets[SHEET_NAMES.transit]);
  if (sheets[SHEET_NAMES.orders])  result.orders  = sheetToObjects_(sheets[SHEET_NAMES.orders]);

  // ZOZO予約一覧を取得
  result.reserves = [];
  const zozoReserveFile = getLatestFile_(FOLDER_IDS.zozo_reserve);
  if (zozoReserveFile) {
    try {
      const text = readShiftJIS_(zozoReserveFile.getBlob());
      const { headers, rows } = parseCSV_(text);
      const hi = makeHeaderIndex_(headers);
      rows.forEach(row => {
        const reserveQty = getCellNum_(row, hi, '予約受付数');
        const orderQty   = getCellNum_(row, hi, '注文数');
        if (reserveQty === 0 && orderQty === 0) return; // 予約なし行はスキップ
        result.reserves.push({
          channel:    'ZOZO',
          name:       getCellStr_(row, hi, '商品名'),
          color:      getCellStr_(row, hi, 'カラー'),
          size:       getCellStr_(row, hi, 'サイズ'),
          csSku:      getCellStr_(row, hi, 'CS品番'),
          reserveQty,
          orderQty,
          deliveryDate: getCellStr_(row, hi, 'お届け予定(初回納期)'),
        });
      });
    } catch(e) { console.log('ZOZO予約取得エラー:', e.message); }
  }

  // 楽天予約を取得（在庫帳票から販売形態=予約商品のみ抽出）
  // JANコードでスマレジ品番と紐付け
  const janToItemNo = {};
  result.products.forEach(p => { if (p.jan) janToItemNo[p.jan] = p.itemNo; });

  const rakutenReserveFile = getLatestFile_(FOLDER_IDS.rakuten);
  if (rakutenReserveFile) {
    try {
      const text = readShiftJIS_(rakutenReserveFile.getBlob());
      const { headers, rows } = parseCSV_(text, true); // 1行目説明行スキップ
      const hi = makeHeaderIndex_(headers);
      rows.forEach(row => {
        if (getCellStr_(row, hi, '販売形態') !== '予約商品') return;
        const reserveQty = getCellNum_(row, hi, '未入荷数');
        const orderQty   = getCellNum_(row, hi, '受注残数');
        if (reserveQty === 0 && orderQty === 0) return;
        const jan = getCellStr_(row, hi, 'JANコード');
        result.reserves.push({
          channel:      '楽天',
          name:         getCellStr_(row, hi, '商品名'),
          color:        getCellStr_(row, hi, 'カラー名'),
          size:         getCellStr_(row, hi, 'サイズ名'),
          csSku:        janToItemNo[jan] || getCellStr_(row, hi, '取引先品番'),
          reserveQty,
          orderQty,
          deliveryDate: '', // 楽天は日付なし
        });
      });
    } catch(e) { console.log('楽天予約取得エラー:', e.message); }
  }

  // 工場出荷ファイルを取得（出荷済みフォルダ内の全CSVを読み込む）
  result.shipments = [];
  try {
    const shipFolder = DriveApp.getFolderById(FOLDER_IDS.shipments);
    const shipFiles = shipFolder.getFiles();
    while (shipFiles.hasNext()) {
      const file = shipFiles.next();
      const mime = file.getMimeType();
      if (!mime.includes('csv') && !mime.includes('text')) continue;

      // ファイル名から工場名・輸送手段・出荷日を抽出
      // 例: 20260503達実エアー出荷.csv
      const fname = file.getName().replace(/\.csv$/i, '');
      const dateMatch    = fname.match(/^(\d{8})/);
      const airMatch     = fname.match(/エアー|エア|air/i);
      const shipDate     = dateMatch ? dateMatch[1].replace(/(\d{4})(\d{2})(\d{2})/, '$1/$2/$3') : '';
      const transport    = airMatch ? 'エアー' : '船';
      // 工場名：日付と「出荷」を除いた部分
      const factoryName  = fname.replace(/^\d{8}/, '').replace(/エアー|エア|船|出荷/g, '').trim();

      try {
        const text = readCsvAutoEncoding_(file.getBlob());
        // ヘッダーなし・空行スキップで読み込む
        // 列順固定：品番, 商品名, カラー, サイズ, 数量
        const lines = text.replace(/^\uFEFF/, '').split(/\r?\n/);
        lines.forEach(line => {
          if (!line.trim()) return;
          const cols = line.split(',').map(c => c.trim().replace(/^"|"$/g, ''));
          const itemNo = cols[0] || '';
          const name   = cols[1] || '';
          const color  = cols[2] || '';
          const size   = cols[3] || '';
          const qty    = parseInt(cols[4]) || 0;
          if (!itemNo || qty === 0) return;
          result.shipments.push({
            itemNo, name, color, size, qty,
            shipDate, transport, factory: factoryName,
          });
        });
      } catch(e) { console.log('出荷ファイル読込エラー:', file.getName(), e.message); }
    }
  } catch(e) { console.log('出荷フォルダ取得エラー:', e.message); }

  return result;
}

// ════════════════════════════════════════════════════════
// ユーティリティ
// ════════════════════════════════════════════════════════

// フォルダ内の最新ファイルを取得
function getLatestFile_(folderId) {
  if (!folderId) return null;
  const folder = DriveApp.getFolderById(folderId);
  let latestFile = null, latestDate = new Date(0);
  const files = folder.getFiles();
  while (files.hasNext()) {
    const file = files.next();
    const mime = file.getMimeType();
    if (!mime.includes('csv') && !mime.includes('text') && !mime.includes('excel')) continue;
    const modified = file.getLastUpdated();
    if (modified > latestDate) { latestDate = modified; latestFile = file; }
  }
  return latestFile;
}

// Shift-JIS固定で読む（楽天用）
function readShiftJIS_(blob) {
  return blob.getDataAsString('Shift_JIS');
}

// UTF-8優先で読む（スマレジ・ZOZO等）
function readCsvAutoEncoding_(blob) {
  try {
    const utf8 = blob.getDataAsString('UTF-8');
    if (/[\u3040-\u9FFF]/.test(utf8)) return utf8;
  } catch (e) {}
  try {
    const sjis = blob.getDataAsString('Shift_JIS');
    if (/[\u3040-\u9FFF]/.test(sjis)) return sjis;
  } catch (e) {}
  return blob.getDataAsString('Shift_JIS');
}

// CSVパース（楽天は1行目が説明行なのでskipFirstLine=trueで2行目をヘッダーとして使う）
function parseCSV_(text, skipFirstLine) {
  // BOM・不可視文字を除去
  text = text.replace(/^[\uFEFF\uFFFE\u0000]+/, '');
  text = text.replace(/^\xEF\xBB\xBF/, ''); // UTF-8 BOM
  const lines = text.split(/\r?\n/).filter(l => l.trim());
  if (lines.length < 2) return { headers: [], rows: [] };

  const parseLine = line => {
    const r = []; let cur = ''; let inQ = false;
    for (const c of line) {
      if (c === '"') { inQ = !inQ; }
      else if (c === ',' && !inQ) { r.push(cur.trim()); cur = ''; }
      else cur += c;
    }
    r.push(cur.trim());
    return r;
  };

  const headerLine = skipFirstLine ? lines[1] : lines[0];
  const dataStart  = skipFirstLine ? 2 : 1;
  const headers = parseLine(headerLine).map(h => h.replace(/^[\s\u3000\uFEFF]+|[\s\u3000\uFEFF]+$/g, ''));
  const rows    = lines.slice(dataStart).map(parseLine).filter(r => r.some(c => c));
  return { headers, rows };
}

// ヘッダー名→インデックスのマップを作成（前後スペース・全角スペース・クォートを除去）
function makeHeaderIndex_(headers) {
  const idx = {};
  headers.forEach((h, i) => {
    const clean = h.replace(/^[\s\u3000"']+|[\s\u3000"']+$/g, '');
    idx[clean] = i;
  });
  return idx;
}

// セルの値を文字列で取得
function getCellStr_(row, hi, key) {
  const i = hi[key];
  return (i !== undefined && row[i] !== undefined) ? String(row[i]).trim() : '';
}

// セルの値を数値で取得
function getCellNum_(row, hi, key) {
  const i = hi[key];
  if (i === undefined || row[i] === undefined) return 0;
  return Number(String(row[i]).replace(/[,，￥¥\s]/g, '')) || 0;
}

function getMeta_() {
  const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_NAMES.meta);
  if (!sheet || sheet.getLastRow() < 2) return [];
  const lastRow  = sheet.getLastRow();
  const startRow = Math.max(1, lastRow - 50);
  const [headers, ...rows] = sheet.getRange(startRow, 1, lastRow - startRow + 1, sheet.getLastColumn()).getValues();
  return rows.map(row => Object.fromEntries(headers.map((h, i) => [h, row[i] ?? ''])));
}

function saveSheet_(name, rows) {
  const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(name) || ss.insertSheet(name);
  sheet.clearContents();
  if (!rows || rows.length === 0) return { saved: 0 };
  const headers = Object.keys(rows[0]);
  sheet.appendRow(headers);
  rows.forEach(r => sheet.appendRow(headers.map(h => r[h] ?? '')));
  return { saved: rows.length };
}

function sheetToObjects_(sheet) {
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow < 2 || lastCol < 1) return [];
  const [headers, ...rows] = sheet.getRange(1, 1, lastRow, lastCol).getValues();
  return rows.map(row => Object.fromEntries(headers.map((h, i) => [h, row[i] ?? ''])));
}

// ════════════════════════════════════════════════════════
// 管理用（GASエディタから手動実行）
// ════════════════════════════════════════════════════════
function testConnection() {
  try {
    console.log('✅ スプレッドシート:', SpreadsheetApp.openById(SPREADSHEET_ID).getName());
  } catch (e) { console.log('❌ スプレッドシートエラー:', e.message); return; }
  Object.entries(FOLDER_IDS).forEach(([ch, id]) => {
    try {
      const folder = DriveApp.getFolderById(id);
      const file   = getLatestFile_(id);
      console.log('✅', ch, '→ フォルダ:', folder.getName(), '/ 最新ファイル:', file ? file.getName() : 'なし');
    } catch (e) { console.log('❌', ch, 'エラー:', e.message); }
  });
}

function testGetData() {
  console.log('=== getData テスト ===');
  const result = getData_(0, 3);
  console.log('総件数:', result.total);
  console.log('先頭3件:', JSON.stringify(result.products, null, 2));
}

function debugLinegift() {
  const result = getData_(0, 2000);
  const hasStock = result.products.filter(p => p.linegift > 0);
  console.log('LINEギフト在庫あり件数:', hasStock.length);
  console.log('先頭3件:', JSON.stringify(hasStock.slice(0, 3), null, 2));
}
