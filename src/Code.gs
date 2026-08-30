const SHEETS = {
  PRODUCTS: 'Products',
  INVENTORY: 'Inventory',
  ORDERS: 'Orders',
  SETTINGS: 'Settings'
};

const APP_VERSION = '6.0.8';
const UPDATER_MARKER = '__FUNFIELDS_SELF_UPDATER_V1__';

function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('消耗品管理・発注')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// 新規導入専用。既に運用中の場合は再実行しないでください。
function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const productRows = [
    ['P001','レジ・決済','レジロール','巻',5,1,1,true,'','','','',10,'セット'],
    ['P002','レジ・決済','PayCas専用レジロール','巻',3,5,1,true,'','','','',1,'巻'],
    ['P003','袋','無料ショップ袋','束',3,5,1,true,'','','','',1,'束'],
    ['P004','袋','有料ショップ袋 Sサイズ','束',2,5,1,true,'','','','',1,'束'],
    ['P005','袋','有料ショップ袋 Mサイズ','束',2,5,1,true,'','','','',1,'束'],
    ['P006','袋','有料ショップ袋 Lサイズ','束',2,5,1,true,'','','','',1,'束'],
    ['P007','事務用品','コピー用紙','箱',1,2,1,true,'','','','',1,'箱'],
    ['P008','カード用品','業務用スリーブ','箱',2,5,1,true,'','','','',1,'箱'],
    ['P009','カード用品','業務用ローダー','箱',2,5,1,true,'','','','',1,'箱'],
    ['P010','ラベル','POS専用ラベル','巻',2,5,1,true,'','','','',1,'巻'],
    ['P011','ラベル','ハンドラベラー用ラベル','巻',2,5,1,true,'','','','',1,'巻'],
    ['P012','清掃用品','モップ','本',1,1,1,true,'','','','',1,'本'],
    ['P013','清掃用品','窓用アルコール','本',1,2,1,true,'','','','',1,'本'],
    ['P014','清掃用品','雑巾','枚',5,10,1,true,'','','','',1,'枚'],
    ['P015','ラベル','マイラベル','巻',2,5,1,true,'','','','',1,'巻'],
    ['P016','プリンター','プリンター用インク','個',1,2,1,true,'','','','',1,'個'],
    ['P017','オリパ用品','オリパ用スリーブ 赤','箱',1,3,1,true,'','','','',1,'箱'],
    ['P018','オリパ用品','オリパ用スリーブ 黒','箱',1,3,1,true,'','','','',1,'箱'],
    ['P019','オリパ用品','オリパ用スリーブ 白','箱',1,3,1,true,'','','','',1,'箱'],
    ['P020','オリパ用品','オリパ用スリーブ 黄色','箱',1,3,1,true,'','','','',1,'箱'],
    ['P021','オリパ用品','オリパ用スリーブ 青','箱',1,3,1,true,'','','','',1,'箱'],
    ['P022','オリパ用品','オリパ用スリーブ 緑','箱',1,3,1,true,'','','','',1,'箱'],
    ['P023','オリパ用品','オリパ用袋 赤','束',1,3,1,true,'','','','',1,'束'],
    ['P024','オリパ用品','オリパ用袋 青','束',1,3,1,true,'','','','',1,'束'],
    ['P025','オリパ用品','オリパ用袋 黒','束',1,3,1,true,'','','','',1,'束'],
    ['P026','オリパ用品','オリパ用袋 緑','束',1,3,1,true,'','','','',1,'束'],
    ['P027','オリパ用品','オリパ用袋 オレンジ','束',1,3,1,true,'','','','',1,'束'],
    ['P028','PSA用品','PSA用オリパ袋','束',1,3,1,true,'','','','',1,'束'],
    ['P029','PSA用品','PSA保護袋','束',1,3,1,true,'','','','',1,'束']
  ];

  const products = getOrCreateSheet_(ss, SHEETS.PRODUCTS);
  products.clear();
  products.getRange(1,1,1,14).setValues([['商品ID','カテゴリ','商品名','在庫単位','最低在庫','推奨発注数（購入単位）','発注数ステップ','有効','仕入先','発注ページURL','商品画像URL','商品メモ','1発注あたり数量','発注単位名']]);
  products.getRange(2,1,productRows.length,14).setValues(productRows);
  products.setFrozenRows(1);

  const settings = getOrCreateSheet_(ss, SHEETS.SETTINGS);
  settings.clear();
  settings.getRange(1,1,1,2).setValues([['設定項目','値']]);
  settings.getRange(2,1,5,2).setValues([
    ['本部発注メール','head-office@example.com'],
    ['店舗一覧','稲沢本店,トレカダンジョンLite大須店'],
    ['メール件名接頭辞','【消耗品発注依頼】'],
    ['管理アプリ名','消耗品管理・発注'],
    ['本部管理PIN','1234']
  ]);
  settings.setFrozenRows(1);

  const inventory = getOrCreateSheet_(ss, SHEETS.INVENTORY);
  inventory.clear();
  inventory.getRange(1,1,1,6).setValues([['更新日時','店舗','商品ID','商品名','現在庫','更新者']]);
  inventory.setFrozenRows(1);

  const orders = getOrCreateSheet_(ss, SHEETS.ORDERS);
  orders.clear();
  orders.getRange(1,1,1,11).setValues([['発注ID','発注日時','店舗','発注者','商品ID','商品名','発注数','単位','状態','メール送信先','備考']]);
  orders.setFrozenRows(1);

  [products, settings, inventory, orders].forEach(s => s.autoResizeColumns(1, s.getLastColumn()));
  return '初期設定が完了しました。';
}

// v2からv3へ更新する時に「1回だけ」実行してください。既存データは消しません。
function upgradeToV3() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const products = ss.getSheetByName(SHEETS.PRODUCTS);
  if (!products) throw new Error('Productsシートが見つかりません。');
  if (products.getLastColumn() < 9) products.getRange(1,9).setValue('仕入先');
  if (products.getLastColumn() < 10) products.getRange(1,10).setValue('発注ページURL');
  if (!products.getRange(1,9).getValue()) products.getRange(1,9).setValue('仕入先');
  if (!products.getRange(1,10).getValue()) products.getRange(1,10).setValue('発注ページURL');

  const settings = ss.getSheetByName(SHEETS.SETTINGS);
  if (!settings) throw new Error('Settingsシートが見つかりません。');
  const vals = settings.getLastRow() > 1 ? settings.getRange(2,1,settings.getLastRow()-1,2).getValues() : [];
  const hasPin = vals.some(r => r[0] === '本部管理PIN');
  if (!hasPin) settings.appendRow(['本部管理PIN','1234']);
  products.autoResizeColumns(1, Math.max(10, products.getLastColumn()));
  return 'v3への更新が完了しました。Settingsシートの「本部管理PIN」を必ず変更してください。';
}

// v3からv4へ更新する時に「1回だけ」実行してください。既存データは消しません。
function upgradeToV4() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const products = ss.getSheetByName(SHEETS.PRODUCTS);
  if (!products) throw new Error('Productsシートが見つかりません。');
  if (products.getLastColumn() < 11) products.getRange(1,11).setValue('商品画像URL');
  if (products.getLastColumn() < 12) products.getRange(1,12).setValue('商品メモ');
  if (!products.getRange(1,11).getValue()) products.getRange(1,11).setValue('商品画像URL');
  if (!products.getRange(1,12).getValue()) products.getRange(1,12).setValue('商品メモ');
  products.autoResizeColumns(1, Math.max(12, products.getLastColumn()));
  return 'v4への更新が完了しました。Productsシートまたは本部画面から商品画像URLと商品メモを登録できます。';
}


// v4からv5へ更新する時に「1回だけ」実行してください。既存データは消しません。
function upgradeToV5() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const settings = ss.getSheetByName(SHEETS.SETTINGS);
  if (!settings) throw new Error('Settingsシートが見つかりません。');
  const vals = settings.getLastRow() > 1 ? settings.getRange(2,1,settings.getLastRow()-1,2).getValues() : [];
  const hasFolder = vals.some(r => r[0] === '商品画像フォルダID');
  if (!hasFolder) settings.appendRow(['商品画像フォルダID','']);
  return 'v5への更新が完了しました。本部画面から商品画像を直接アップロードできます。初回アップロード時にGoogle Driveの権限許可が必要です。';
}


// v5.1からv5.2へ更新する時に「1回だけ」実行してください。
// Webアプリから確実に同じスプレッドシートを開けるよう、スプレッドシートIDを保存します。
function upgradeToV5_2() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error('このスクリプトを対象のスプレッドシートから開いて実行してください。');
  PropertiesService.getScriptProperties().setProperty('SPREADSHEET_ID', ss.getId());
  return 'v5.2への更新が完了しました。スプレッドシートIDを固定登録しました。';
}


// v5.7からv5.8へ更新する時に「1回だけ」実行してください。既存データは消しません。
// 在庫単位と購入単位を分離します。
function upgradeToV5_8() {
  const ss = getSS_();
  const sheet = ss.getSheetByName(SHEETS.PRODUCTS);
  if (!sheet) throw new Error('Productsシートが見つかりません。');
  sheet.getRange(1,4).setValue('在庫単位');
  sheet.getRange(1,6).setValue('推奨発注数（購入単位）');
  sheet.getRange(1,7).setValue('発注数ステップ');
  sheet.getRange(1,13).setValue('1発注あたり数量');
  sheet.getRange(1,14).setValue('発注単位名');
  if (sheet.getLastRow() > 1) {
    const rows = sheet.getRange(2,1,sheet.getLastRow()-1,14).getValues();
    for (let i = 0; i < rows.length; i++) {
      const rowNo = i + 2;
      const id = String(rows[i][0] || '');
      const stockUnit = String(rows[i][3] || '個');
      let packQty = Number(rows[i][12]) || 0;
      let purchaseUnit = String(rows[i][13] || '').trim();
      if (!packQty) packQty = 1;
      if (!purchaseUnit) purchaseUnit = stockUnit;
      // レジロールは例として「1セット=10巻」に初期設定します。
      if (id === 'P001' && (!rows[i][12] || !rows[i][13])) {
        packQty = 10;
        purchaseUnit = 'セット';
        if (Number(rows[i][5]) === 10) sheet.getRange(rowNo,6).setValue(1);
      }
      sheet.getRange(rowNo,13,1,2).setValues([[packQty,purchaseUnit]]);
      if (!Number(rows[i][6])) sheet.getRange(rowNo,7).setValue(1);
    }
  }
  sheet.autoResizeColumns(1,14);
  return 'v5.8への更新が完了しました。レジロールは「1セット=10巻」に設定しました。他の商品は商品マスターから購入単位を設定できます。';
}


// v5.10からv6.0へ更新する時に「1回だけ」実行してください。既存データは消しません。
function upgradeToV6() {
  const ss = getSS_();
  const settings = ss.getSheetByName(SHEETS.SETTINGS);
  if (!settings) throw new Error('Settingsシートが見つかりません。');
  ensureSetting_(settings, 'アプリバージョン', APP_VERSION);
  ensureSetting_(settings, '更新マニフェストURL', '');
  ensureSetting_(settings, '最終更新確認日時', '');
  ensureSetting_(settings, '最終更新バージョン', APP_VERSION);
  return 'v6.0への更新が完了しました。次に新しいバージョンでデプロイし、本部発注→システム更新で更新元URLを設定してください。';
}


// v6.0.4からv6.0.5へ更新する時に「1回だけ」実行してください。
// 現在存在するWEB_APPデプロイのうち、最終更新日時が最も新しいものを固定保存します。
function upgradeToV6_0_5() {
  const ss = getSS_();
  const scriptId = ScriptApp.getScriptId();
  const result = scriptApiRequest_('get', '/v1/projects/' + encodeURIComponent(scriptId) + '/deployments');
  const deployments = (result && result.deployments) ? result.deployments : [];
  const webApps = deployments.filter(function(d) {
    return d && Array.isArray(d.entryPoints) && d.entryPoints.some(function(ep) {
      return ep && ep.entryPointType === 'WEB_APP' && ep.webApp && ep.webApp.url;
    });
  });
  if (!webApps.length) {
    throw new Error('Webアプリのデプロイが見つかりません。先に「新しいデプロイ」→「ウェブアプリ」でデプロイしてください。');
  }
  webApps.sort(function(a,b) {
    return new Date(b.updateTime || 0).getTime() - new Date(a.updateTime || 0).getTime();
  });
  const selected = webApps[0];
  const webEntry = selected.entryPoints.find(function(ep) {
    return ep && ep.entryPointType === 'WEB_APP' && ep.webApp && ep.webApp.url;
  });
  setSetting_(ss, 'WebアプリデプロイID', String(selected.deploymentId || ''));
  setSetting_(ss, 'WebアプリURL', webEntry && webEntry.webApp ? String(webEntry.webApp.url || '') : '');
  setSetting_(ss, 'アプリバージョン', APP_VERSION);
  setSetting_(ss, '最終更新バージョン', APP_VERSION);
  return 'v6.0.5への更新が完了しました。WebアプリデプロイIDを固定登録しました：' + selected.deploymentId;
}

function getInitialData() {
  const ss = getSS_();
  const settings = getSettings_(ss);
  return {
    appName: settings['管理アプリ名'] || '消耗品管理・発注',
    stores: (settings['店舗一覧'] || '').split(',').map(s => s.trim()).filter(Boolean),
    products: getProducts_(ss)
  };
}

function getInventory(store) {
  const ss = getSS_();
  const products = getProducts_(ss);
  const sheet = ss.getSheetByName(SHEETS.INVENTORY);
  const values = sheet && sheet.getLastRow() > 1 ? sheet.getRange(2,1,sheet.getLastRow()-1,6).getValues() : [];
  const latest = {};
  values.forEach(r => { if (r[1] === store) latest[r[2]] = Number(r[4]) || 0; });
  return products.map(p => ({...p, currentStock: latest[p.id] ?? 0}));
}

function saveInventoryItem(payload) {
  if (!payload || !payload.store || !payload.productId) throw new Error('入力内容が不正です。');
  const ss = getSS_();
  const sheet = ss.getSheetByName(SHEETS.INVENTORY);
  if (!sheet) throw new Error('Inventoryシートが見つかりません。');
  sheet.appendRow([new Date(),payload.store,payload.productId,payload.productName||'',Number(payload.currentStock)||0,payload.user||'']);
  return {ok:true};
}

function submitOrder(payload) {
  if (!payload || !payload.store || !Array.isArray(payload.items) || payload.items.length === 0) throw new Error('発注商品がありません。');
  const ss = getSS_();
  const settings = getSettings_(ss);
  const recipient = settings['本部発注メール'];
  if (!recipient || recipient === 'head-office@example.com') throw new Error('Settingsシートの「本部発注メール」を実際のアドレスに変更してください。');
  const orderId = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd-HHmmss') + '-' + Math.floor(Math.random()*900+100);
  const now = new Date();
  const sheet = ss.getSheetByName(SHEETS.ORDERS);
  const rows = payload.items.map(i => [orderId,now,payload.store,payload.user||'',i.productId,i.productName,Number(i.orderQty)||0,i.orderUnitName||i.unit||'','発注依頼',recipient,payload.note||'']);
  sheet.getRange(sheet.getLastRow()+1,1,rows.length,11).setValues(rows);

  const subject = `${settings['メール件名接頭辞'] || '【消耗品発注依頼】'}${payload.store}`;
  const lines = [`${payload.store}から消耗品の発注依頼です。`,'',`発注ID：${orderId}`,`発注者：${payload.user || '未入力'}`,`発注日時：${Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy/MM/dd HH:mm')}`,'','【発注内容】'];
  payload.items.forEach(i => {
    const purchaseUnit = i.orderUnitName || i.unit || '';
    const packQty = Math.max(1, Number(i.packQty) || 1);
    const stockUnit = i.stockUnit || '';
    const total = (Number(i.orderQty) || 0) * packQty;
    const detail = packQty > 1 || purchaseUnit !== stockUnit ? `（合計 ${total}${stockUnit} / 1${purchaseUnit}=${packQty}${stockUnit}）` : '';
    lines.push(`・${i.productName} × ${i.orderQty}${purchaseUnit}${detail}`);
  });
  if (payload.note) lines.push('','【備考】',payload.note);
  MailApp.sendEmail({to:recipient,subject,body:lines.join('\n')});
  return {ok:true,orderId};
}

function getOrderHistory(store) {
  const ss = getSS_();
  const sheet = ss.getSheetByName(SHEETS.ORDERS);
  if (!sheet || sheet.getLastRow() <= 1) return [];
  const values = sheet.getRange(2,1,sheet.getLastRow()-1,11).getValues();
  const normalizedStore = String(store || '').trim();
  return values
    .filter(r => !normalizedStore || String(r[2] || '').trim() === normalizedStore)
    .slice(-100)
    .reverse()
    .map(r => ({
      orderId:r[0],
      date:r[1] instanceof Date ? Utilities.formatDate(r[1], Session.getScriptTimeZone(), 'yyyy/MM/dd HH:mm') : String(r[1] || ''),
      store:String(r[2] || '').trim(),
      user:r[3],
      productName:r[5],
      qty:r[6],
      unit:r[7],
      status:r[8]
    }));
}

function verifyAdminPin(pin) {
  assertAdmin_(pin);
  return {ok:true};
}

function getHeadOfficeOrders(pin) {
  assertAdmin_(pin);
  const ss = getSS_();
  const products = {};
  getProducts_(ss).forEach(p => products[p.id] = p);
  const sheet = ss.getSheetByName(SHEETS.ORDERS);
  if (!sheet || sheet.getLastRow() <= 1) return [];
  const rows = sheet.getRange(2,1,sheet.getLastRow()-1,11).getValues();
  return rows.slice().reverse().map((r,revIdx) => {
    const p = products[r[4]] || {};
    return {
      rowNumber: sheet.getLastRow() - revIdx,
      orderId:r[0], date:r[1], store:r[2], user:r[3], productId:r[4], productName:r[5], qty:r[6], unit:r[7], status:r[8], note:r[10],
      supplier:p.supplier || '', orderUrl:p.orderUrl || '', imageUrl:p.imageUrl || '', productMemo:p.productMemo || '', packQty:p.packQty || 1, stockUnit:p.unit || '', orderUnitName:p.orderUnitName || p.unit || ''
    };
  });
}

function updateOrderLineStatus(payload) {
  assertAdmin_(payload && payload.pin);
  const allowed = ['発注依頼','手配済み','納品済み'];
  if (!payload || !allowed.includes(payload.status)) throw new Error('不正な状態です。');
  const ss = getSS_();
  const sheet = ss.getSheetByName(SHEETS.ORDERS);
  const row = Number(payload.rowNumber);
  if (!row || row < 2 || row > sheet.getLastRow()) throw new Error('対象行が見つかりません。');
  sheet.getRange(row,9).setValue(payload.status);
  return {ok:true};
}


function getProductAdminMaster(pin) {
  assertAdmin_(pin);
  return getAllProducts_(getSS_());
}

function addProduct(payload) {
  assertAdmin_(payload && payload.pin);
  if (!payload) throw new Error('入力内容が不正です。');
  var category = String(payload.category || '').trim();
  var name = String(payload.name || '').trim();
  var unit = String(payload.unit || '').trim();
  if (!category) throw new Error('カテゴリを入力してください。');
  if (!name) throw new Error('商品名を入力してください。');
  if (!unit) throw new Error('単位を入力してください。');
  var minStock = Math.max(0, Number(payload.minStock) || 0);
  var suggestedQty = Math.max(1, Number(payload.suggestedQty) || 1);
  var orderUnit = 1;
  var packQty = Math.max(1, Number(payload.packQty) || 1);
  var orderUnitName = String(payload.orderUnitName || '').trim() || unit;
  var ss = getSS_();
  var sheet = ss.getSheetByName(SHEETS.PRODUCTS);
  if (!sheet) throw new Error('Productsシートが見つかりません。');
  var all = getAllProducts_(ss);
  for (var i = 0; i < all.length; i++) {
    if (String(all[i].name).toLowerCase() === name.toLowerCase()) throw new Error('同じ商品名がすでに登録されています。');
  }
  var maxNo = 0;
  for (var j = 0; j < all.length; j++) {
    var m = String(all[j].id || '').match(/^P(\d+)$/i);
    if (m) maxNo = Math.max(maxNo, Number(m[1]) || 0);
  }
  var id = 'P' + Utilities.formatString('%03d', maxNo + 1);
  sheet.appendRow([id, category, name, unit, minStock, suggestedQty, orderUnit, true, '', '', '', '', packQty, orderUnitName]);
  return {ok:true, productId:id};
}

function updateProductMaster(payload) {
  assertAdmin_(payload && payload.pin);
  if (!payload || !payload.productId) throw new Error('商品が指定されていません。');
  var category = String(payload.category || '').trim();
  var name = String(payload.name || '').trim();
  var unit = String(payload.unit || '').trim();
  if (!category || !name || !unit) throw new Error('カテゴリ・商品名・単位は必須です。');
  var ss = getSS_();
  var sheet = ss.getSheetByName(SHEETS.PRODUCTS);
  if (!sheet) throw new Error('Productsシートが見つかりません。');
  var lastRow = sheet.getLastRow();
  if (lastRow <= 1) throw new Error('商品が見つかりません。');
  var rows = sheet.getRange(2,1,lastRow-1,12).getValues();
  var idx = -1;
  for (var i = 0; i < rows.length; i++) if (String(rows[i][0]) === String(payload.productId)) { idx = i; break; }
  if (idx < 0) throw new Error('商品が見つかりません。');
  var rowNo = idx + 2;
  sheet.getRange(rowNo,2,1,7).setValues([[
    category,
    name,
    unit,
    Math.max(0, Number(payload.minStock) || 0),
    Math.max(1, Number(payload.suggestedQty) || 1),
    1,
    payload.active === false ? false : true
  ]]);
  var packQty = Math.max(1, Number(payload.packQty) || 1);
  var orderUnitName = String(payload.orderUnitName || '').trim() || unit;
  sheet.getRange(rowNo,13,1,2).setValues([[packQty,orderUnitName]]);
  return {ok:true};
}

function getAllProducts_(ss) {
  var sheet = ss.getSheetByName(SHEETS.PRODUCTS);
  if (!sheet || sheet.getLastRow() <= 1) return [];
  var cols = Math.max(14, sheet.getLastColumn());
  var values = sheet.getRange(2,1,sheet.getLastRow()-1,cols).getValues();
  var out = [];
  for (var i = 0; i < values.length; i++) {
    var r = values[i];
    if (!r[0]) continue;
    out.push({
      id:r[0], category:r[1], name:r[2], unit:r[3], minStock:Number(r[4])||0,
      suggestedQty:Number(r[5])||1, orderUnit:Number(r[6])||1,
      active:(r[7]===true || String(r[7]).toUpperCase()==='TRUE'),
      supplier:r[8]||'', orderUrl:r[9]||'', imageUrl:r[10]||'', productMemo:r[11]||'',
      packQty:Number(r[12])||1, orderUnitName:r[13]||r[3]||''
    });
  }
  return out;
}

function getSupplierMaster(pin) {
  assertAdmin_(pin);
  return getProducts_(getSS_());
}

function saveSupplierLink(payload) {
  assertAdmin_(payload && payload.pin);
  if (!payload || !payload.productId) throw new Error('商品が指定されていません。');
  const url = String(payload.orderUrl || '').trim();
  const imageUrl = String(payload.imageUrl || '').trim();
  if (url && !/^https?:\/\//i.test(url)) throw new Error('発注ページURLは http:// または https:// から入力してください。');
  if (imageUrl && !/^https?:\/\//i.test(imageUrl)) throw new Error('商品画像URLは http:// または https:// から入力してください。');
  const ss = getSS_();
  const sheet = ss.getSheetByName(SHEETS.PRODUCTS);
  const rows = sheet.getRange(2,1,sheet.getLastRow()-1,Math.max(12,sheet.getLastColumn())).getValues();
  const idx = rows.findIndex(r => r[0] === payload.productId);
  if (idx < 0) throw new Error('商品が見つかりません。');
  sheet.getRange(idx+2,9,1,4).setValues([[String(payload.supplier||'').trim(),url,imageUrl,String(payload.productMemo||'').trim()]]);
  return {ok:true};
}



function uploadProductImage(payload) {
  assertAdmin_(payload && payload.pin);
  if (!payload || !payload.productId) throw new Error('商品が指定されていません。');
  const dataUrl = String(payload.dataUrl || '');
  const m = dataUrl.match(/^data:(image\/[a-zA-Z0-9.+-]+);base64,(.+)$/);
  if (!m) throw new Error('画像データの形式が不正です。');
  const mimeType = m[1];
  const bytes = Utilities.base64Decode(m[2]);
  if (bytes.length > 5 * 1024 * 1024) throw new Error('画像サイズは5MB以下にしてください。');
  const ss = getSS_();
  const product = getProducts_(ss).find(p => p.id === payload.productId);
  if (!product) throw new Error('商品が見つかりません。');
  const folder = getOrCreateProductImageFolder_(ss);
  const ext = ({'image/jpeg':'jpg','image/png':'png','image/webp':'webp','image/gif':'gif'})[mimeType] || 'bin';
  const safeName = sanitizeFileName_(product.name || payload.productId);
  const fileName = `${payload.productId}_${safeName}_${Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd_HHmmss')}.${ext}`;
  const blob = Utilities.newBlob(bytes, mimeType, fileName);
  const file = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  const imageUrl = `https://drive.google.com/file/d/${file.getId()}/view?usp=sharing`;
  const sheet = ss.getSheetByName(SHEETS.PRODUCTS);
  const rows = sheet.getRange(2,1,sheet.getLastRow()-1,Math.max(12,sheet.getLastColumn())).getValues();
  const idx = rows.findIndex(r => r[0] === payload.productId);
  if (idx < 0) throw new Error('商品が見つかりません。');
  sheet.getRange(idx+2,11).setValue(imageUrl);
  return {ok:true, imageUrl, fileId:file.getId()};
}

function getOrCreateProductImageFolder_(ss) {
  const settingsSheet = ss.getSheetByName(SHEETS.SETTINGS);
  const values = settingsSheet.getRange(2,1,Math.max(0, settingsSheet.getLastRow()-1),2).getValues();
  let rowIndex = values.findIndex(r => r[0] === '商品画像フォルダID');
  let folderId = rowIndex >= 0 ? String(values[rowIndex][1] || '').trim() : '';
  if (folderId) {
    try {
      return DriveApp.getFolderById(folderId);
    } catch (e) {}
  }
  const rootFolderName = '消耗品管理・発注';
  const childFolderName = '商品画像';
  let rootFolder;
  const rootIt = DriveApp.getFoldersByName(rootFolderName);
  rootFolder = rootIt.hasNext() ? rootIt.next() : DriveApp.createFolder(rootFolderName);
  let childFolder;
  const childIt = rootFolder.getFoldersByName(childFolderName);
  childFolder = childIt.hasNext() ? childIt.next() : rootFolder.createFolder(childFolderName);
  childFolder.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  if (rowIndex >= 0) {
    settingsSheet.getRange(rowIndex+2,2).setValue(childFolder.getId());
  } else {
    settingsSheet.appendRow(['商品画像フォルダID', childFolder.getId()]);
  }
  return childFolder;
}

function sanitizeFileName_(name) {
  return String(name || '').replace(/[\\/:*?"<>|#%&{}$!'@+=`~]/g, '_').replace(/\s+/g, '_').slice(0, 80);
}

// __FUNFIELDS_SELF_UPDATER_V1__
function getSystemUpdateInfo(pin) {
  assertAdmin_(pin);
  const ss = getSS_();
  const settings = getSettings_(ss);
  return {
    currentVersion: APP_VERSION,
    manifestUrl: String(settings['更新マニフェストURL'] || ''),
    serviceUrl: String(settings['WebアプリURL'] || ScriptApp.getService().getUrl() || ''),
    deploymentId: String(settings['WebアプリデプロイID'] || ''),
    scriptId: ScriptApp.getScriptId(),
    lastChecked: settings['最終更新確認日時'] || '',
    lastVersion: settings['最終更新バージョン'] || APP_VERSION
  };
}

function saveUpdateManifestUrl(payload) {
  assertAdmin_(payload && payload.pin);
  const url = String((payload && payload.url) || '').trim();
  if (url && !/^https:\/\//i.test(url)) throw new Error('更新マニフェストURLは https:// から始まるURLを設定してください。');
  const ss = getSS_();
  setSetting_(ss, '更新マニフェストURL', url);
  return {ok:true, url:url};
}

function checkForAppUpdate(pin) {
  assertAdmin_(pin);
  const ss = getSS_();
  const settings = getSettings_(ss);
  const url = String(settings['更新マニフェストURL'] || '').trim();
  if (!url) throw new Error('更新マニフェストURLが未設定です。最初に更新元URLを設定してください。');
  const manifest = fetchJson_(url, '更新マニフェスト');
  validateReleaseManifest_(manifest);
  const hasUpdate = compareVersions_(String(manifest.version), APP_VERSION) > 0;
  setSetting_(ss, '最終更新確認日時', new Date());
  setSetting_(ss, '最終更新バージョン', String(manifest.version));
  return {
    ok:true,
    currentVersion:APP_VERSION,
    latestVersion:String(manifest.version),
    hasUpdate:hasUpdate,
    notes:manifest.notes || '',
    publishedAt:manifest.publishedAt || '',
    packageUrl:String(manifest.packageUrl || '')
  };
}

function applyAppUpdate(payload) {
  assertAdmin_(payload && payload.pin);
  const ss = getSS_();
  const settings = getSettings_(ss);
  const manifestUrl = String(settings['更新マニフェストURL'] || '').trim();
  if (!manifestUrl) throw new Error('更新マニフェストURLが未設定です。');

  const manifest = fetchJson_(manifestUrl, '更新マニフェスト');
  validateReleaseManifest_(manifest);
  const targetVersion = String(manifest.version);
  if (compareVersions_(targetVersion, APP_VERSION) <= 0) throw new Error('現在のバージョンが最新です。');

  const packageData = fetchJson_(String(manifest.packageUrl), '更新パッケージ');
  validateUpdatePackage_(packageData, targetVersion);

  const scriptId = ScriptApp.getScriptId();
  const deploymentId = String(settings['WebアプリデプロイID'] || '').trim();
  if (!deploymentId) {
    throw new Error('WebアプリデプロイIDが未登録です。Apps Scriptで upgradeToV6_0_5() を1回実行してください。');
  }

  // 更新対象が本当にWebアプリかを毎回検証し、ライブラリ等への誤更新を防ぎます。
  const targetDeployment = scriptApiRequest_('get',
    '/v1/projects/' + encodeURIComponent(scriptId) + '/deployments/' + encodeURIComponent(deploymentId)
  );
  const isWebApp = targetDeployment && Array.isArray(targetDeployment.entryPoints) &&
    targetDeployment.entryPoints.some(function(ep) {
      return ep && ep.entryPointType === 'WEB_APP' && ep.webApp && ep.webApp.url;
    });
  if (!isWebApp) {
    throw new Error('固定登録されたデプロイIDがWebアプリではありません。安全のため更新を停止しました。upgradeToV6_0_5() を再実行してください。');
  }

  // 現在のコードをGoogle Driveへバックアップしてから更新します。
  const currentContent = scriptApiRequest_('get', '/v1/projects/' + encodeURIComponent(scriptId) + '/content');
  const backupFile = saveUpdateBackup_(currentContent, APP_VERSION);

  try {
    scriptApiRequest_('put', '/v1/projects/' + encodeURIComponent(scriptId) + '/content', {files:packageData.files});
    const version = scriptApiRequest_('post', '/v1/projects/' + encodeURIComponent(scriptId) + '/versions', {
      description:'消耗品管理・発注 v' + targetVersion + ' 自動更新'
    });
    const versionNumber = Number(version.versionNumber);
    if (!versionNumber) throw new Error('新しいApps Scriptバージョン番号を取得できませんでした。');

    scriptApiRequest_('put', '/v1/projects/' + encodeURIComponent(scriptId) + '/deployments/' + encodeURIComponent(deploymentId), {
      deploymentConfig:{
        scriptId:scriptId,
        versionNumber:versionNumber,
        manifestFileName:'appsscript',
        description:'消耗品管理・発注 v' + targetVersion
      }
    });
    // ソース更新後も、ユーザーが設定した更新元URLを確実に維持します。
    setSetting_(ss, '更新マニフェストURL', manifestUrl);
    setSetting_(ss, 'アプリバージョン', targetVersion);
    setSetting_(ss, '最終更新バージョン', targetVersion);
    return {
      ok:true,
      version:targetVersion,
      versionNumber:versionNumber,
      backupFileName:backupFile.getName(),
      message:'v' + targetVersion + 'への更新が完了しました。数秒待ってから画面を再読み込みしてください。'
    };
  } catch (e) {
    throw new Error('自動更新に失敗しました。更新前バックアップ「' + backupFile.getName() + '」は保存済みです。詳細: ' + errorMessage_(e));
  }
}

function scriptApiRequest_(method, path, body) {
  const url = 'https://script.googleapis.com' + path;
  const options = {
    method:method,
    muteHttpExceptions:true,
    headers:{Authorization:'Bearer ' + ScriptApp.getOAuthToken()},
    contentType:'application/json'
  };
  if (body !== undefined) options.payload = JSON.stringify(body);
  const response = UrlFetchApp.fetch(url, options);
  const code = response.getResponseCode();
  const text = response.getContentText();
  let data = {};
  try { data = text ? JSON.parse(text) : {}; } catch (e) { data = {raw:text}; }
  if (code < 200 || code >= 300) {
    let msg = (data && data.error && data.error.message) ? data.error.message : text;
    if (code === 403 && /SERVICE_DISABLED|has not been used|disabled/i.test(msg || '')) {
      msg += '。Google Cloud側で「Apps Script API (script.googleapis.com)」を有効にしてください。';
    }
    throw new Error('Apps Script APIエラー ' + code + ': ' + msg);
  }
  return data;
}

function fetchJson_(url, label) {
  if (!/^https:\/\//i.test(String(url || ''))) throw new Error(label + 'のURLが不正です。');
  const response = UrlFetchApp.fetch(url, {muteHttpExceptions:true, followRedirects:true});
  const code = response.getResponseCode();
  if (code < 200 || code >= 300) throw new Error(label + 'を取得できませんでした（HTTP ' + code + '）。');
  try { return JSON.parse(response.getContentText()); }
  catch (e) { throw new Error(label + 'が正しいJSONではありません。'); }
}

function validateReleaseManifest_(m) {
  if (!m || !m.version || !m.packageUrl) throw new Error('更新マニフェストに version または packageUrl がありません。');
  if (!/^https:\/\//i.test(String(m.packageUrl))) throw new Error('更新パッケージURLは https:// である必要があります。');
}

function validateUpdatePackage_(p, version) {
  if (!p || String(p.version || '') !== String(version) || !Array.isArray(p.files) || !p.files.length) {
    throw new Error('更新パッケージの形式またはバージョンが一致しません。');
  }
  const names = {};
  p.files.forEach(function(f){ if (f && f.name) names[f.name] = true; });
  if (!names.Code || !names.Index || !names.appsscript) throw new Error('更新パッケージに Code / Index / appsscript が揃っていません。');
  const codeFile = p.files.find(function(f){return f.name === 'Code';});
  if (!codeFile || String(codeFile.source || '').indexOf('__FUNFIELDS_SELF_UPDATER_V1__') < 0) {
    throw new Error('安全のため更新を停止しました。更新パッケージに自動更新機能が含まれていません。');
  }
}

function saveUpdateBackup_(content, version) {
  const root = getOrCreateNamedFolder_('消耗品管理・発注');
  const folder = getOrCreateChildFolder_(root, 'システム更新バックアップ');
  const stamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd_HHmmss');
  return folder.createFile('backup_v' + version + '_' + stamp + '.json', JSON.stringify(content, null, 2), MimeType.PLAIN_TEXT);
}

function getOrCreateNamedFolder_(name) {
  const it = DriveApp.getFoldersByName(name);
  return it.hasNext() ? it.next() : DriveApp.createFolder(name);
}

function getOrCreateChildFolder_(parent, name) {
  const it = parent.getFoldersByName(name);
  return it.hasNext() ? it.next() : parent.createFolder(name);
}

function getCurrentDeploymentId_() {
  const url = ScriptApp.getService().getUrl() || '';
  const m = url.match(/\/s\/([^/]+)\/exec/i);
  return m ? m[1] : '';
}

function compareVersions_(a,b) {
  const aa=String(a||'0').split('.').map(function(x){return parseInt(x,10)||0;});
  const bb=String(b||'0').split('.').map(function(x){return parseInt(x,10)||0;});
  const n=Math.max(aa.length,bb.length);
  for(let i=0;i<n;i++){
    const x=aa[i]||0,y=bb[i]||0;
    if(x>y)return 1;
    if(x<y)return -1;
  }
  return 0;
}

function ensureSetting_(sheet, key, value) {
  const n=Math.max(0,sheet.getLastRow()-1);
  const vals=n ? sheet.getRange(2,1,n,2).getValues() : [];
  const idx=vals.findIndex(function(r){return r[0]===key;});
  if(idx<0) sheet.appendRow([key,value]);
}

function setSetting_(ss, key, value) {
  const sheet=ss.getSheetByName(SHEETS.SETTINGS);
  if(!sheet)throw new Error('Settingsシートが見つかりません。');
  const n=Math.max(0,sheet.getLastRow()-1);
  const vals=n ? sheet.getRange(2,1,n,2).getValues() : [];
  const idx=vals.findIndex(function(r){return r[0]===key;});
  if(idx>=0) sheet.getRange(idx+2,2).setValue(value);
  else sheet.appendRow([key,value]);
}

function errorMessage_(e) {
  return e && e.message ? e.message : String(e || '不明なエラー');
}

function assertAdmin_(pin) {
  const settings = getSettings_(getSS_());
  const expected = String(settings['本部管理PIN'] || '');
  if (!expected || String(pin || '') !== expected) throw new Error('本部管理PINが違います。');
}

function getProducts_(ss) {
  const sheet = ss.getSheetByName(SHEETS.PRODUCTS);
  if (!sheet || sheet.getLastRow() <= 1) return [];
  const cols = Math.max(14, sheet.getLastColumn());
  return sheet.getRange(2,1,sheet.getLastRow()-1,cols).getValues()
    .filter(r => r[7] === true || String(r[7]).toUpperCase() === 'TRUE')
    .map(r => ({id:r[0],category:r[1],name:r[2],unit:r[3],minStock:Number(r[4])||0,suggestedQty:Number(r[5])||1,orderUnit:Number(r[6])||1,supplier:r[8]||'',orderUrl:r[9]||'',imageUrl:r[10]||'',productMemo:r[11]||'',packQty:Number(r[12])||1,orderUnitName:r[13]||r[3]||''}));
}


function getSS_() {
  const id = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
  if (id) return SpreadsheetApp.openById(id);
  const active = SpreadsheetApp.getActiveSpreadsheet();
  if (active) return active;
  throw new Error('管理用スプレッドシートが登録されていません。Apps Scriptで upgradeToV5_2() を1回実行してください。');
}

function getSettings_(ss) {
  const sheet = ss.getSheetByName(SHEETS.SETTINGS);
  const out = {};
  if (!sheet || sheet.getLastRow() <= 1) return out;
  sheet.getRange(2,1,sheet.getLastRow()-1,2).getValues().forEach(r => out[r[0]] = r[1]);
  return out;
}
function getOrCreateSheet_(ss,name){return ss.getSheetByName(name)||ss.insertSheet(name);}
