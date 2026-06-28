// =============================================
// Google Apps Script - ブランド物販ダッシュボードAPI
//
// 機能:
// 1. ダッシュボードAPI（未販売・販売済のデータ送信）
// 2. 「ブランド物販」メニュー → 「販売済に移動」
//    未販売シートで行を選択 → メニュークリックで販売済シートに移動
//
// 使い方:
// 1. ブランド物販管理スプレッドシートを開く
// 2. 拡張機能 → Apps Script
// 3. このコードを貼り付けて保存
// 4. デプロイ → 新しいデプロイ → ウェブアプリ
//    - 実行するユーザー: 自分
//    - アクセス: 全員
// 5. URLをダッシュボードのAPI設定に貼り付け
// =============================================

// スプレッドシートを開いた時にメニューを追加
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('ブランド物販')
    .addItem('📦→💰 販売済に移動', 'moveToSold')
    .addToUi();
}

// 未販売シートで選択した行を販売済シートに移動
function moveToSold() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var stockSheet = ss.getSheetByName('未販売');
  var soldSheet = ss.getSheetByName('販売済');

  if (!stockSheet || !soldSheet) {
    SpreadsheetApp.getUi().alert('「未販売」と「販売済」シートが必要です');
    return;
  }

  // 現在のシートが未販売か確認
  var activeSheet = ss.getActiveSheet();
  if (activeSheet.getName() !== '未販売') {
    SpreadsheetApp.getUi().alert('「未販売」シートで行を選択してから実行してください');
    return;
  }

  var selection = activeSheet.getActiveRange();
  var startRow = selection.getRow();
  var numRows = selection.getNumRows();

  // ヘッダー行は除外
  if (startRow <= 1) {
    SpreadsheetApp.getUi().alert('ヘッダー行は移動できません。データ行を選択してください');
    return;
  }

  // 未販売の列数
  var stockCols = stockSheet.getLastColumn();

  // 販売済の最終行を取得
  var soldLastRow = soldSheet.getLastRow();

  // 販売済のヘッダーから列マッピングを作成
  var soldHeaders = soldSheet.getRange(1, 1, 1, soldSheet.getLastColumn()).getValues()[0];
  var stockHeaders = stockSheet.getRange(1, 1, 1, stockCols).getValues()[0];

  var moved = 0;

  // 下の行から処理（削除時にずれないように）
  for (var i = numRows - 1; i >= 0; i--) {
    var row = startRow + i;
    var rowData = stockSheet.getRange(row, 1, 1, stockCols).getValues()[0];

    // 番号が空なら飛ばす
    if (!rowData[0]) continue;

    // 販売済シートに行を追加
    var newRow = soldLastRow + moved + 1;

    // 共通列をマッピング（番号、個数、カテゴリ、商品名、仕入れ先、仕入日、仕入金額、掲載日）
    for (var j = 0; j < Math.min(8, stockCols); j++) {
      soldSheet.getRange(newRow, j + 1).setValue(rowData[j]);
    }

    // 未販売シートから行を削除
    stockSheet.deleteRow(row);
    moved++;
  }

  if (moved > 0) {
    SpreadsheetApp.getUi().alert(moved + '件を販売済シートに移動しました。\n販売先・販売金額・購入された日などを入力してください。');
  } else {
    SpreadsheetApp.getUi().alert('移動するデータがありませんでした');
  }
}

function doGet(e) {
  var callback = e.parameter.callback;
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var stockItems = readSheet(ss, '未販売', false);
  var soldItems = readSheet(ss, '販売済', true);

  var json = JSON.stringify({ stock: stockItems, sold: soldItems });

  if (callback) {
    return ContentService.createTextOutput(callback + '(' + json + ')')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(json)
    .setMimeType(ContentService.MimeType.JSON);
}

function readSheet(ss, sheetName, isSold) {
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) return [];

  var data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];

  // ヘッダー行を探す
  var hi = -1;
  for (var i = 0; i < Math.min(data.length, 10); i++) {
    var row = data[i].join(',');
    if (row.indexOf('番号') >= 0 && row.indexOf('仕入') >= 0) {
      hi = i;
      break;
    }
  }
  if (hi === -1) return [];

  // 列マッピング
  var hdr = data[hi];
  var ci = {};
  for (var j = 0; j < hdr.length; j++) {
    var h = String(hdr[j]).replace(/\s/g, '');
    if (h === '番号') ci.num = j;
    if (/個数|仕入個数/.test(h)) ci.qty = j;
    if (/カテゴリ/.test(h)) ci.category = j;
    if (h === '商品名' || h === '商品') ci.name = j;
    if (/仕入れ先|仕入先/.test(h)) ci.supplier = j;
    if (/仕入日/.test(h)) ci.purchaseDate = j;
    if (/仕入金額|仕入れ額/.test(h)) ci.cost = j;
    if (/掲載日/.test(h)) ci.listDate = j;
    if (/販売先/.test(h)) ci.marketplace = j;
    if (/販売金額|販売価格|販売予定価格/.test(h)) ci.salePrice = j;
    if (/購入された日/.test(h)) ci.saleDate = j;
    if (/販売手数料|想定手数料/.test(h)) ci.commission = j;
    if (/送料/.test(h)) ci.shipping = j;
    // 旧形式互換
    if (/メルカリ/.test(h) && /手数料/.test(h)) ci.comM = j;
    if ((/ペイペイ/.test(h) || /PayPay/i.test(h)) && /手数料/.test(h)) ci.comP = j;
    if (/ラクマ/.test(h) && /手数料/.test(h)) ci.comR = j;
    if (/ヤフオク/.test(h) && /手数料/.test(h)) ci.comY = j;
    if (/エコオク/.test(h) && !/仕入/.test(h)) ci.comE = j;
  }

  function parseYen(v) {
    return Number(String(v || '').replace(/[¥￥,]/g, '')) || 0;
  }
  function fmtDate(v) {
    if (!v) return '';
    var d = new Date(v);
    if (isNaN(d.getTime())) return '';
    return d.getFullYear() + '-' + ('0' + (d.getMonth() + 1)).slice(-2) + '-' + ('0' + d.getDate()).slice(-2);
  }

  var items = [];
  for (var i = hi + 1; i < data.length; i++) {
    var r = data[i];
    var nv = r[ci.num];
    if (!nv || String(r.join(',')).indexOf('合計') >= 0) continue;
    if (!/\d/.test(String(nv))) continue;

    var cat = String(r[ci.category] || '').trim();
    if (cat === 'バック') cat = 'バッグ';
    var sup = String(r[ci.supplier] || '').trim();
    if (sup === 'エコオク') sup = 'エコリング';

    var item = {
      id: 'B-' + ('000' + nv).slice(-3),
      category: cat || '財布',
      name: String(r[ci.name] || '').replace(/\n/g, ' ').trim(),
      supplier: sup,
      purchaseDate: fmtDate(r[ci.purchaseDate]),
      cost: parseYen(r[ci.cost]),
      listDate: fmtDate(r[ci.listDate]),
      quantity: Number(r[ci.qty]) || 1
    };

    if (isSold) {
      var mkt = String(r[ci.marketplace] || '').trim();
      if (mkt === 'エコオク') mkt = 'エコリング';
      if (mkt === 'ペイペイ') mkt = 'PayPay';

      var com = 0;
      if (ci.commission !== undefined) {
        com = parseYen(r[ci.commission]);
      } else {
        com = parseYen(r[ci.comM]) + parseYen(r[ci.comP]) + parseYen(r[ci.comR]) + parseYen(r[ci.comY]) + parseYen(r[ci.comE]);
      }

      item.marketplace = mkt;
      item.salePrice = parseYen(r[ci.salePrice]);
      item.saleDate = fmtDate(r[ci.saleDate]);
      item.commission = com;
      item.shipping = parseYen(r[ci.shipping]);
    } else {
      if (ci.marketplace !== undefined) {
        var mk2 = String(r[ci.marketplace] || '').trim();
        if (mk2 === 'ペイペイ') mk2 = 'PayPay';
        item.marketplace = mk2;
      }
      if (ci.salePrice !== undefined && parseYen(r[ci.salePrice])) item.listPrice = parseYen(r[ci.salePrice]);
      if (ci.commission !== undefined && parseYen(r[ci.commission])) item.estCommission = parseYen(r[ci.commission]);
      if (ci.shipping !== undefined && parseYen(r[ci.shipping])) item.estShipping = parseYen(r[ci.shipping]);
    }

    items.push(item);
  }
  return items;
}
