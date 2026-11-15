/**
 * ごはんナビ 記録保存用サーバー（Google Apps Script） v3
 *
 * 保存するもの：
 *  - visits シート … 行った店（訪問履歴は log 列に保存）
 *  - photos シート … 写真の縮小版と、Googleドライブ上の元画像へのリンク
 *  - 保存した場所 … スクリプト プロパティ（SETTINGS）
 * Googleマップの店舗検索：スクリプト プロパティに PLACES_KEY を登録すると有効になります。
 *
 * 更新したら：デプロイ → デプロイを管理 → 鉛筆 → バージョン「新バージョン」→ デプロイ
 * （URLと合言葉は変わりません。setup の再実行は不要です）
 *
 * v3 の変更点：
 *  - 訪問履歴の突き合わせ（別の端末が足した訪問を、あとから保存した端末が消さない）
 *  - 記録の読み込みを POST でも受け付ける（合言葉をURLに載せない）
 *  - 毎回の書式の書き直しをやめ、写真の件数は必要な列だけ読む（速くなる）
 *  - Googleマップ検索：近い順の指定、ジャンルを指定した探し直し、口コミを開く公式リンク
 */

const VERSION = 3;
const SHEET_NAME = 'visits';
const HEAD = ['id', 'name', 'genre', 'area', 'date', 'rating', 'memo', 'url', 'lat', 'lng', 'osmId', 'placeId', 'log', 'createdAt', 'updatedAt'];
const PHOTO_SHEET = 'photos';
const PHOTO_HEAD = ['id', 'shopId', 'date', 'thumb', 'driveUrl', 'fileId', 'createdAt'];
const PHOTO_FOLDER = 'ごはんナビ写真';
// Googleマップ検索の1日の上限回数（無料枠：月1,000回 → 30回×30日＝900回で収まる）
const DAILY_PLACES_LIMIT = 30;

/** 初回に実行：シート準備、合言葉の発行、権限の承認（更新のたびに実行する必要はありません） */
function setup() {
  const props = PropertiesService.getScriptProperties();
  props.deleteProperty('SCHEMA');
  sheet_();
  photoSheet_();
  let token = props.getProperty('TOKEN');
  if (!token) {
    token = Utilities.getUuid().replace(/-/g, '').slice(0, 20);
    props.setProperty('TOKEN', token);
  }
  photoFolder_(); // ドライブの権限確認を兼ねる
  Logger.log('合言葉（アプリの設定画面に入力）： ' + token);
  Logger.log('記録のスプレッドシート： ' + SpreadsheetApp.getActive().getUrl());
  Logger.log('Googleマップ検索： ' + (props.getProperty('PLACES_KEY') ? '有効（鍵が登録済み）' : '未設定（OpenStreetMapを使用）'));
}

/** Googleマップ検索の鍵が正しく動くか確認する（蒲田駅周辺で1回だけ検索） */
function testPlaces() {
  const r = placesNearby_(35.5623, 139.7163, 500);
  Logger.log(r.ok ? ('OK：' + r.places.length + '件取得。例：' + r.places.slice(0, 3).map(p => p.name).join('、')) : ('エラー：' + r.error + ' ' + (r.detail || '')));
}

function resetToken() {
  PropertiesService.getScriptProperties().deleteProperty('TOKEN');
  setup();
}

/* ---------- シート ---------- */
// 見出しと書式（文字列）を整える。時間がかかるので、シートを作ったとき・版が変わったときだけ行う
function ensureHeader_(s, head) {
  const lastCol = Math.max(s.getLastColumn(), 1);
  const cur = s.getRange(1, 1, 1, lastCol).getDisplayValues()[0];
  if (cur.join('') === '') {
    s.getRange(1, 1, 1, head.length).setValues([head]).setFontWeight('bold');
    s.setFrozenRows(1);
  } else {
    const missing = head.filter(h => cur.indexOf(h) < 0);
    if (missing.length) s.getRange(1, cur.length + 1, 1, missing.length).setValues([missing]).setFontWeight('bold');
  }
  s.getRange(1, 1, s.getMaxRows(), s.getLastColumn()).setNumberFormat('@');
}
function ensureSchema_() {
  const props = PropertiesService.getScriptProperties();
  if (props.getProperty('SCHEMA') === String(VERSION)) return;
  const ss = SpreadsheetApp.getActive();
  [[SHEET_NAME, HEAD], [PHOTO_SHEET, PHOTO_HEAD]].forEach(([name, head]) => {
    ensureHeader_(ss.getSheetByName(name) || ss.insertSheet(name), head);
  });
  props.setProperty('SCHEMA', String(VERSION));
}
function getSheet_(name, head) {
  ensureSchema_();
  const ss = SpreadsheetApp.getActive();
  let s = ss.getSheetByName(name);
  if (!s) { s = ss.insertSheet(name); ensureHeader_(s, head); } // 消されていたら作り直す
  return s;
}
function sheet_() { return getSheet_(SHEET_NAME, HEAD); }
function photoSheet_() { return getSheet_(PHOTO_SHEET, PHOTO_HEAD); }
function photoFolder_() {
  const it = DriveApp.getFoldersByName(PHOTO_FOLDER);
  const f = it.hasNext() ? it.next() : DriveApp.createFolder(PHOTO_FOLDER);
  const props = PropertiesService.getScriptProperties();
  if (props.getProperty('PHOTO_FOLDER_URL') !== f.getUrl()) props.setProperty('PHOTO_FOLDER_URL', f.getUrl());
  return f;
}
function header_(s) { return s.getRange(1, 1, 1, Math.max(s.getLastColumn(), 1)).getDisplayValues()[0]; }
function rows_(s) {
  const v = s.getDataRange().getDisplayValues();
  const head = v.shift() || [];
  return { head, rows: v };
}
/** 1列だけ読む（2行目以降）。大きな列（写真の縮小版など）を読まずに済ませる */
function col_(s, head, name) {
  const i = head.indexOf(name), n = s.getLastRow() - 1;
  if (i < 0 || n < 1) return [];
  return s.getRange(2, i + 1, n, 1).getDisplayValues().map(r => r[0]);
}

function out_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
function authorized_(token) {
  const t = PropertiesService.getScriptProperties().getProperty('TOKEN');
  return !!t && String(token || '') === t;
}
function safe_(v) {
  v = (v === null || v === undefined) ? '' : String(v).slice(0, 45000);
  return /^[=+\-@]/.test(v) ? "'" + v : v;
}

/* ---------- 記録 ---------- */
function list_() {
  const { head, rows } = rows_(sheet_());
  // 写真の件数は shopId の列だけを読んで数える
  const ps = photoSheet_();
  const counts = {};
  col_(ps, header_(ps), 'shopId').forEach(id => { if (id) counts[id] = (counts[id] || 0) + 1; });
  return rows.filter(r => r[0]).map(r => {
    const o = {};
    head.forEach((k, i) => { if (k) o[k] = r[i]; });
    o.photoCount = counts[o.id] || 0;
    return o;
  });
}
function parseLog_(v) {
  try {
    const a = typeof v === 'string' ? (v ? JSON.parse(v) : []) : (Array.isArray(v) ? v : []);
    return a.filter(x => x && typeof x === 'object');
  } catch (e) { return []; }
}
/**
 * 訪問履歴の突き合わせ。
 * base … 送ってきた端末が「サーバーにある」と知っていた訪問のID一覧。
 * サーバーにあって base に無い訪問は、別の端末があとから足したものなので残す。
 * base にあって送られてこなかった訪問は、その端末が消したものなので消す。
 */
function mergeLog_(clientLog, serverLog, base) {
  const known = {}, mine = {};
  base.forEach(id => { known[String(id)] = true; });
  clientLog.forEach(x => { if (x.id) mine[String(x.id)] = true; });
  const extra = serverLog.filter(x => x.id && !known[String(x.id)] && !mine[String(x.id)]);
  if (!extra.length) return clientLog;
  return clientLog.concat(extra).sort((a, b) => String(a.d || '').localeCompare(String(b.d || '')));
}
function upsert_(v) {
  if (!v || !v.id || !v.name) throw new Error('id と name は必須です');
  const s = sheet_();
  const head = header_(s);
  const ids = s.getRange(1, 1, Math.max(s.getLastRow(), 1), 1).getDisplayValues().map(r => r[0]);
  const idx = ids.indexOf(String(v.id));
  const r = idx >= 1 ? idx + 1 : s.getLastRow() + 1;
  if (idx >= 1 && Array.isArray(v.baseLog) && Array.isArray(v.log)) {
    const li = head.indexOf('log');
    const cur = li >= 0 ? parseLog_(s.getRange(r, li + 1).getDisplayValue()) : [];
    const merged = mergeLog_(v.log, cur, v.baseLog);
    if (merged !== v.log) {
      v.log = merged;
      const last = merged[merged.length - 1] || {};
      v.date = last.d || ''; v.rating = Number(last.r) || 0; v.memo = last.m || '';
    }
  }
  v.updatedAt = String(Date.now());
  if (!v.createdAt) v.createdAt = v.updatedAt;
  if (Array.isArray(v.log)) v.log = JSON.stringify(v.log);
  const row = head.map(k => safe_(v[k]));
  s.getRange(r, 1, 1, head.length).setNumberFormat('@').setValues([row]);
}
/** 写真の行を（下から）消す。ドライブ上の画像はゴミ箱へ */
function removePhotoRows_(colName, value) {
  const ps = photoSheet_();
  const head = header_(ps);
  const keys = col_(ps, head, colName), files = col_(ps, head, 'fileId');
  for (let i = keys.length - 1; i >= 0; i--) {
    if (keys[i] === String(value)) {
      try { if (files[i]) DriveApp.getFileById(files[i]).setTrashed(true); } catch (e) {}
      ps.deleteRow(i + 2);
    }
  }
}
function delete_(id) {
  const s = sheet_();
  const ids = s.getRange(1, 1, Math.max(s.getLastRow(), 1), 1).getDisplayValues().map(r => r[0]);
  const idx = ids.indexOf(String(id));
  if (idx >= 1) s.deleteRow(idx + 1);
  removePhotoRows_('shopId', id); // その店の写真も削除
}

/* ---------- 写真 ---------- */
function photoAdd_(b) {
  if (!b.shopId || !b.thumb || !b.full) throw new Error('写真データがありません');
  const m = String(b.full).match(/^data:(image\/\w+);base64,(.+)$/);
  if (!m) throw new Error('写真の形式が違います');
  const blob = Utilities.newBlob(Utilities.base64Decode(m[2]), m[1], (b.date || '') + '_' + String(b.name || 'photo').slice(0, 40) + '.jpg');
  const file = photoFolder_().createFile(blob);
  const id = 'p_' + Date.now().toString(36) + Math.random().toString(36).slice(2, 6);
  const ps = photoSheet_();
  const head = header_(ps);
  const o = { id, shopId: b.shopId, date: b.date || '', thumb: String(b.thumb).slice(0, 45000), driveUrl: file.getUrl(), fileId: file.getId(), createdAt: String(Date.now()) };
  ps.getRange(ps.getLastRow() + 1, 1, 1, head.length).setNumberFormat('@').setValues([head.map(k => safe_(o[k]))]);
  return { id, shopId: o.shopId, date: o.date, thumb: o.thumb, driveUrl: o.driveUrl };
}
function photos_(shopId) {
  const ps = photoSheet_();
  const head = header_(ps);
  const ix = k => head.indexOf(k);
  // shopId の列で行を探し、その店の行だけを読む
  const out = [];
  col_(ps, head, 'shopId').forEach((id, i) => {
    if (id !== String(shopId)) return;
    const r = ps.getRange(i + 2, 1, 1, head.length).getDisplayValues()[0];
    out.push({ id: r[ix('id')], date: r[ix('date')], thumb: r[ix('thumb')], driveUrl: r[ix('driveUrl')] });
  });
  return out;
}
function photoDelete_(id) { removePhotoRows_('id', id); }

/* ---------- 設定（保存した場所など） ---------- */
function settings_() {
  try { return JSON.parse(PropertiesService.getScriptProperties().getProperty('SETTINGS') || '{}'); } catch (e) { return {}; }
}
function settingsSet_(st) {
  const s = JSON.stringify(st || {});
  if (s.length > 8500) throw new Error('保存できる量を超えました');
  PropertiesService.getScriptProperties().setProperty('SETTINGS', s);
}

/* ---------- Googleマップ検索（Places API New: Nearby Search） ---------- */
const FIELDS_FULL = 'places.id,places.displayName,places.location,places.primaryType,places.types,places.rating,places.userRatingCount,places.priceLevel,places.currentOpeningHours.openNow,places.currentOpeningHours.nextCloseTime,places.currentOpeningHours.nextOpenTime,places.regularOpeningHours.weekdayDescriptions,places.goodForChildren,places.shortFormattedAddress,places.googleMapsUri,places.googleMapsLinks.reviewsUri';
const FIELDS_BASIC = 'places.id,places.displayName,places.location,places.primaryType,places.types,places.rating,places.userRatingCount,places.priceLevel,places.currentOpeningHours.openNow,places.regularOpeningHours.weekdayDescriptions,places.shortFormattedAddress,places.googleMapsUri';
const TYPES_FULL = ['restaurant', 'cafe', 'bar', 'coffee_shop', 'fast_food_restaurant'];
const TYPES_BASIC = ['restaurant', 'cafe', 'bar'];
/**
 * ジャンルを指定して探し直すときの、Googleの「場所の種類」。
 * full … 細かい種類まで指定（受け付けられなければ safe で再試行）
 * exact … その種類がジャンルとほぼ一致する（true のとき、アプリは結果をそのジャンルとして扱う）
 */
const GENRE_TYPES = {
  'ラーメン': { full: ['ramen_restaurant'], safe: ['ramen_restaurant'], exact: true },
  '居酒屋・バー': { full: ['japanese_izakaya_restaurant', 'yakitori_restaurant', 'bar', 'pub', 'wine_bar'], safe: ['bar', 'pub'], exact: true },
  'カフェ': { full: ['cafe', 'coffee_shop'], safe: ['cafe'], exact: true },
  '焼肉': { full: ['yakiniku_restaurant', 'korean_barbecue_restaurant', 'barbecue_restaurant', 'korean_restaurant'], safe: ['barbecue_restaurant', 'korean_restaurant'], exact: true },
  '寿司': { full: ['sushi_restaurant'], safe: ['sushi_restaurant'], exact: true },
  '中華': { full: ['chinese_restaurant'], safe: ['chinese_restaurant'], exact: true },
  'そば・うどん': { full: ['japanese_restaurant'], safe: ['japanese_restaurant'], exact: false },
  'カレー': { full: ['japanese_curry_restaurant', 'indian_restaurant'], safe: ['indian_restaurant'], exact: true },
  '定食・和食': { full: ['japanese_restaurant', 'tonkatsu_restaurant'], safe: ['japanese_restaurant'], exact: false },
  'イタリアン': { full: ['italian_restaurant', 'pizza_restaurant'], safe: ['italian_restaurant', 'pizza_restaurant'], exact: true },
  'ファストフード': { full: ['fast_food_restaurant', 'hamburger_restaurant'], safe: ['fast_food_restaurant'], exact: true }
};

// Googleの規約により店舗情報は保存（キャッシュ）しない。場所IDのみ記録に残す。
// opt.rank = 'distance' … 近い順（「今いるお店」用）。opt.genre … ジャンルを指定して探す
function placesNearby_(lat, lng, radius, opt) {
  opt = opt || {};
  const props = PropertiesService.getScriptProperties();
  const key = props.getProperty('PLACES_KEY');
  if (!key) return { ok: false, error: 'no_places_key' };
  lat = Number(lat); lng = Number(lng); radius = Math.min(Math.max(Number(radius) || 600, 100), 5000);
  // 1日の上限（日本時間で日付を区切る）
  const day = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd');
  const cKey = 'PLACES_COUNT_' + day;
  const used = Number(props.getProperty(cKey) || 0);
  if (used >= DAILY_PLACES_LIMIT) return { ok: false, error: 'daily_limit', used, limit: DAILY_PLACES_LIMIT };

  const gt = opt.genre ? GENRE_TYPES[opt.genre] : null;
  const full = gt ? gt.full : TYPES_FULL, safe = gt ? gt.safe : TYPES_BASIC;
  const rank = opt.rank === 'distance' ? 'DISTANCE' : 'POPULARITY';
  // 400（項目や種類の指定が受け付けられない）のときだけ、控えめな条件で再試行する
  const attempts = [[FIELDS_FULL, full], [FIELDS_FULL, safe], [FIELDS_BASIC, full], [FIELDS_BASIC, safe]]
    .filter((a, i, all) => all.findIndex(b => b[0] === a[0] && b[1].join() === a[1].join()) === i);
  let res, code, text;
  for (const [fields, types] of attempts) {
    res = UrlFetchApp.fetch('https://places.googleapis.com/v1/places:searchNearby', {
      method: 'post', contentType: 'application/json', muteHttpExceptions: true,
      headers: { 'X-Goog-Api-Key': key, 'X-Goog-FieldMask': fields },
      payload: JSON.stringify({
        includedTypes: types, maxResultCount: 20, rankPreference: rank,
        languageCode: 'ja', regionCode: 'JP',
        locationRestriction: { circle: { center: { latitude: lat, longitude: lng }, radius } }
      })
    });
    code = res.getResponseCode(); text = res.getContentText();
    if (code !== 400) break;
  }
  props.setProperty(cKey, String(used + 1));
  // 古い日付のカウンターを掃除
  Object.keys(props.getProperties()).forEach(k => { if (/^PLACES_COUNT_/.test(k) && k !== cKey) props.deleteProperty(k); });
  if (code !== 200) return { ok: false, error: 'places_error', detail: code + ' ' + text.slice(0, 300) };

  const j = JSON.parse(text || '{}');
  const places = (j.places || []).map(p => {
    const co = p.currentOpeningHours || {};
    return {
      pid: p.id, name: (p.displayName && p.displayName.text) || '', lat: p.location && p.location.latitude, lng: p.location && p.location.longitude,
      type: p.primaryType || '', types: (p.types || []).slice(0, 6), rating: p.rating || null, n: p.userRatingCount || 0,
      price: p.priceLevel || '', open: (co.openNow === undefined ? null : co.openNow), nextClose: co.nextCloseTime || '', nextOpen: co.nextOpenTime || '',
      hours: (p.regularOpeningHours && p.regularOpeningHours.weekdayDescriptions) || [], kids: (p.goodForChildren === undefined ? null : p.goodForChildren),
      addr: p.shortFormattedAddress || '', uri: p.googleMapsUri || '',
      rev: (p.googleMapsLinks && p.googleMapsLinks.reviewsUri) || '' // Googleマップでその店の口コミを開く公式リンク
    };
  }).filter(p => p.name && p.lat != null);
  return { ok: true, places, src: 'google', used: used + 1, limit: DAILY_PLACES_LIMIT, rank: rank.toLowerCase(), genre: gt ? opt.genre : '', exact: !!(gt && gt.exact) };
}

/* ---------- 入口 ---------- */
function listOut_() {
  const props = PropertiesService.getScriptProperties();
  return { ok: true, ver: VERSION, visits: list_(), settings: settings_(), places: !!props.getProperty('PLACES_KEY'),
    sheetUrl: SpreadsheetApp.getActive().getUrl(), photoUrl: props.getProperty('PHOTO_FOLDER_URL') || '' };
}
// 旧版のアプリ（更新前の画面が残っている端末）用。新しいアプリは POST の list を使い、合言葉をURLに載せない
function doGet(e) {
  const p = (e && e.parameter) || {};
  if (!authorized_(p.token)) return out_({ ok: false, error: 'auth' });
  return out_(listOut_());
}

function doPost(e) {
  let b;
  try { b = JSON.parse(e.postData.contents); } catch (err) { return out_({ ok: false, error: 'bad_request' }); }
  if (!authorized_(b.token)) return out_({ ok: false, error: 'auth' });
  try {
    // 読み取り系はロック不要
    if (b.action === 'list') return out_(listOut_());
    if (b.action === 'nearby') return out_(placesNearby_(b.lat, b.lng, b.radius, { rank: b.rank, genre: b.genre }));
    if (b.action === 'photos') return out_({ ok: true, photos: photos_(b.shopId) });
  } catch (err) {
    return out_({ ok: false, error: String(err && err.message || err) });
  }
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
    if (b.action === 'upsert') upsert_(b.visit);
    else if (b.action === 'delete') delete_(b.id);
    else if (b.action === 'photoAdd') return out_({ ok: true, photo: photoAdd_(b) });
    else if (b.action === 'photoDelete') { photoDelete_(b.id); return out_({ ok: true }); }
    else if (b.action === 'settingsSet') { settingsSet_(b.settings); return out_({ ok: true, settings: settings_() }); }
    else return out_({ ok: false, error: 'unknown_action' });
    return out_({ ok: true, ver: VERSION, visits: list_() });
  } catch (err) {
    return out_({ ok: false, error: String(err && err.message || err) });
  } finally {
    try { lock.releaseLock(); } catch (x) {}
  }
}
