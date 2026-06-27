// =============================================
// Google Apps Script - ブランド物販ダッシュボードAPI
//
// 「未販売」「販売済」2つのシートから読み取り、
// ダッシュボードにJSON形式で返します。
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
    if (/販売金額|販売価格/.test(h)) ci.salePrice = j;
    if (/購入された日/.test(h)) ci.saleDate = j;
    if (/販売手数料/.test(h)) ci.commission = j;
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
    }

    items.push(item);
  }
  return items;
}
