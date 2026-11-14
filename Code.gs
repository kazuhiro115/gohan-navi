/**
 * ごはんナビ 記録保存用サーバー（Google Apps Script） v2
 *
 * 保存するもの：
 *  - visits シート … 行った店（訪問履歴は log 列に保存）
 *  - photos シート … 写真の縮小版と、Googleドライブ上の元画像へのリンク
 *  - 保存した場所 … スクリプト プロパティ（SETTINGS）
 * Googleマップの店舗検索：スクリプト プロパティに PLACES_KEY を登録すると有効になります。
 *
 * 更新したら：デプロイ → デプロイを管理 → 鉛筆 → バージョン「新バージョン」→ デプロイ
 * （URLは変わりません）
 */

const VERSION = 2;
const SHEET_NAME = 'visits';
const HEAD = ['id', 'name', 'genre', 'area', 'date', 'rating', 'memo', 'url', 'lat', 'lng', 'osmId', 'placeId', 'log', 'createdAt', 'updatedAt'];
const PHOTO_SHEET = 'photos';
const PHOTO_HEAD = ['id', 'shopId', 'date', 'thumb', 'driveUrl', 'fileId', 'createdAt'];
const PHOTO_FOLDER = 'ごはんナビ写真';
// Googleマップ検索の1日の上限回数（無料枠：月1,000回 → 30回×30日＝900回で収まる）
const DAILY_PLACES_LIMIT = 30;

/** 初回・更新時に実行：シート準備、合言葉の発行、権限の承認 */
function setup() {
  sheet_();
  photoSheet_();
  const props = PropertiesService.getScriptProperties();
  let token = props.getProperty('TOKEN');
  if (!token) {
    token = Utilities.getUuid().replace(/-/g, '').slice(0, 20);
    props.setProperty('TOKEN', token);
  }
  photoFolder_(); // ドライブの権限確認を兼ねる
  Logger.log('合言葉（アプリの設定画面に入力）： ' + token);
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
function sheet_() {
  const ss = SpreadsheetApp.getActive();
  let s = ss.getSheetByName(SHEET_NAME);
  if (!s) s = ss.insertSheet(SHEET_NAME);
  ensureHeader_(s, HEAD);
  return s;
}
function photoSheet_() {
  const ss = SpreadsheetApp.getActive();
  let s = ss.getSheetByName(PHOTO_SHEET);
  if (!s) s = ss.insertSheet(PHOTO_SHEET);
  ensureHeader_(s, PHOTO_HEAD);
  return s;
}
function photoFolder_() {
  const it = DriveApp.getFoldersByName(PHOTO_FOLDER);
  return it.hasNext() ? it.next() : DriveApp.createFolder(PHOTO_FOLDER);
}
function header_(s) { return s.getRange(1, 1, 1, s.getLastColumn()).getDisplayValues()[0]; }
function rows_(s) {
  const v = s.getDataRange().getDisplayValues();
  const head = v.shift() || [];
  return { head, rows: v };
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
  const counts = {};
  const p = rows_(photoSheet_());
  const si = p.head.indexOf('shopId');
  p.rows.forEach(r => { if (r[si]) counts[r[si]] = (counts[r[si]] || 0) + 1; });
  return rows.filter(r => r[0]).map(r => {
    const o = {};
    head.forEach((k, i) => { if (k) o[k] = r[i]; });
    o.photoCount = counts[o.id] || 0;
    return o;
  });
}
function upsert_(v) {
  if (!v || !v.id || !v.name) throw new Error('id と name は必須です');
  const s = sheet_();
  const head = header_(s);
  const ids = s.getRange(1, 1, Math.max(s.getLastRow(), 1), 1).getDisplayValues().map(r => r[0]);
  v.updatedAt = String(Date.now());
  if (!v.createdAt) v.createdAt = v.updatedAt;
  if (Array.isArray(v.log)) v.log = JSON.stringify(v.log);
  const row = head.map(k => safe_(v[k]));
  const idx = ids.indexOf(String(v.id));
  const r = idx >= 1 ? idx + 1 : s.getLastRow() + 1;
  s.getRange(r, 1, 1, head.length).setNumberFormat('@').setValues([row]);
}
function delete_(id) {
  const s = sheet_();
  const ids = s.getRange(1, 1, Math.max(s.getLastRow(), 1), 1).getDisplayValues().map(r => r[0]);
  const idx = ids.indexOf(String(id));
  if (idx >= 1) s.deleteRow(idx + 1);
  // その店の写真も削除（ドライブ上の画像はゴミ箱へ）
  const ps = photoSheet_();
  const { head, rows } = rows_(ps);
  const si = head.indexOf('shopId'), fi = head.indexOf('fileId');
  for (let i = rows.length - 1; i >= 0; i--) {
    if (rows[i][si] === String(id)) {
      try { if (rows[i][fi]) DriveApp.getFileById(rows[i][fi]).setTrashed(true); } catch (e) {}
      ps.deleteRow(i + 2);
    }
  }
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
  const { head, rows } = rows_(photoSheet_());
  const ix = k => head.indexOf(k);
  return rows.filter(r => r[ix('shopId')] === String(shopId)).map(r => ({
    id: r[ix('id')], date: r[ix('date')], thumb: r[ix('thumb')], driveUrl: r[ix('driveUrl')]
  }));
}
function photoDelete_(id) {
  const ps = photoSheet_();
  const { head, rows } = rows_(ps);
  const ii = head.indexOf('id'), fi = head.indexOf('fileId');
  for (let i = rows.length - 1; i >= 0; i--) {
    if (rows[i][ii] === String(id)) {
      try { if (rows[i][fi]) DriveApp.getFileById(rows[i][fi]).setTrashed(true); } catch (e) {}
      ps.deleteRow(i + 2);
    }
  }
}

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
const FIELDS_FULL = 'places.id,places.displayName,places.location,places.primaryType,places.types,places.rating,places.userRatingCount,places.priceLevel,places.currentOpeningHours.openNow,places.currentOpeningHours.nextCloseTime,places.currentOpeningHours.nextOpenTime,places.regularOpeningHours.weekdayDescriptions,places.goodForChildren,places.shortFormattedAddress,places.googleMapsUri';
const FIELDS_BASIC = 'places.id,places.displayName,places.location,places.primaryType,places.types,places.rating,places.userRatingCount,places.priceLevel,places.currentOpeningHours.openNow,places.regularOpeningHours.weekdayDescriptions,places.shortFormattedAddress,places.googleMapsUri';
const TYPES_FULL = ['restaurant', 'cafe', 'bar', 'coffee_shop', 'fast_food_restaurant'];
const TYPES_BASIC = ['restaurant', 'cafe', 'bar'];

// Googleの規約により店舗情報は保存（キャッシュ）しない。場所IDのみ記録に残す。
function placesNearby_(lat, lng, radius) {
  const props = PropertiesService.getScriptProperties();
  const key = props.getProperty('PLACES_KEY');
  if (!key) return { ok: false, error: 'no_places_key' };
  lat = Number(lat); lng = Number(lng); radius = Math.min(Math.max(Number(radius) || 600, 100), 5000);
  // 1日の上限（日本時間で日付を区切る）
  const day = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd');
  const cKey = 'PLACES_COUNT_' + day;
  const used = Number(props.getProperty(cKey) || 0);
  if (used >= DAILY_PLACES_LIMIT) return { ok: false, error: 'daily_limit', used };

  const attempts = [[FIELDS_FULL, TYPES_FULL], [FIELDS_BASIC, TYPES_FULL], [FIELDS_BASIC, TYPES_BASIC]];
  let res, code, text;
  for (const [fields, types] of attempts) {
    res = UrlFetchApp.fetch('https://places.googleapis.com/v1/places:searchNearby', {
      method: 'post', contentType: 'application/json', muteHttpExceptions: true,
      headers: { 'X-Goog-Api-Key': key, 'X-Goog-FieldMask': fields },
      payload: JSON.stringify({
        includedTypes: types, maxResultCount: 20, rankPreference: 'POPULARITY',
        languageCode: 'ja', regionCode: 'JP',
        locationRestriction: { circle: { center: { latitude: lat, longitude: lng }, radius } }
      })
    });
    code = res.getResponseCode(); text = res.getContentText();
    if (code !== 400) break; // 400（項目や種類の指定ミス）のときだけ控えめな条件で再試行
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
      addr: p.shortFormattedAddress || '', uri: p.googleMapsUri || ''
    };
  }).filter(p => p.name && p.lat != null);
  return { ok: true, places, src: 'google', used: used + 1 };
}

/* ---------- 入口 ---------- */
function doGet(e) {
  const p = (e && e.parameter) || {};
  if (!authorized_(p.token)) return out_({ ok: false, error: 'auth' });
  return out_({ ok: true, ver: VERSION, visits: list_(), settings: settings_(), places: !!PropertiesService.getScriptProperties().getProperty('PLACES_KEY') });
}

function doPost(e) {
  let b;
  try { b = JSON.parse(e.postData.contents); } catch (err) { return out_({ ok: false, error: 'bad_request' }); }
  if (!authorized_(b.token)) return out_({ ok: false, error: 'auth' });
  try {
    // 読み取り系はロック不要
    if (b.action === 'nearby') return out_(placesNearby_(b.lat, b.lng, b.radius));
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
