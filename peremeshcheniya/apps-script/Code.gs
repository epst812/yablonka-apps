/**
 * «Перемещения» — заявки на перемещение товара между точками.
 * Google Apps Script. Вставьте этот код в Расширения → Apps Script таблицы,
 * запустите setup() один раз, затем Развернуть → Веб-приложение.
 *
 * Листы:
 *   Точки        — Город | Точка
 *   Товары       — Товар | Ед.
 *   Перемещения  — журнал: одна строка = один товар в заявке (удобно для СУММЕСЛИМН)
 */

const SHEET_POINTS = 'Точки';
const SHEET_PRODUCTS = 'Товары';
const SHEET_MOVES = 'Перемещения';
const SHEET_SUBS = 'Подписки';

const HEAD = [
  'ID', 'Создана', 'Город', 'Откуда', 'Куда', 'Товар', 'Кол-во', 'Ед.',
  'Статус', 'Принята', 'Получена', 'Получено (факт)', 'Расхождение',
  'Комментарий', 'Причина отказа'
];
const C = {}; HEAD.forEach((h, i) => C[h] = i); // индекс колонки по названию

const ST = { NEW: 'Новая', ACCEPTED: 'Принята', RECEIVED: 'Получена', REJECTED: 'Отклонена', CANCELLED: 'Отменена' };
const HISTORY_DAYS = 7; // сколько дней показывать завершённые заявки

/* ---------- первоначальная настройка (запустить один раз) ---------- */
function setup() {
  const ss = SpreadsheetApp.getActive();

  let pts = ss.getSheetByName(SHEET_POINTS);
  if (!pts) {
    pts = ss.insertSheet(SHEET_POINTS);
    pts.getRange(1, 1, 7, 2).setValues([
      ['Город', 'Точка'],
      ['Пенза', 'Точка 1'], ['Пенза', 'Точка 2'], ['Пенза', 'Точка 3'],
      ['Балаково', 'Точка 4'], ['Балаково', 'Точка 5'], ['Балаково', 'Точка 6']
    ]);
    pts.getRange('1:1').setFontWeight('bold');
    pts.setFrozenRows(1);
  }

  let prs = ss.getSheetByName(SHEET_PRODUCTS);
  if (!prs) {
    prs = ss.insertSheet(SHEET_PRODUCTS);
    prs.getRange(1, 1, 4, 2).setValues([
      ['Товар', 'Ед.'], ['Круассан', 'шт'], ['Синнабон', 'шт'], ['Йогурт', 'шт']
    ]);
    prs.getRange('1:1').setFontWeight('bold');
    prs.setFrozenRows(1);
  }

  if (!ss.getSheetByName(SHEET_SUBS)) {
    const sb = ss.insertSheet(SHEET_SUBS);
    sb.getRange(1, 1, 1, 3).setValues([['Точка', 'Токен', 'Обновлено']]).setFontWeight('bold');
    sb.setFrozenRows(1);
  }

  let mv = ss.getSheetByName(SHEET_MOVES);
  if (!mv) {
    mv = ss.insertSheet(SHEET_MOVES);
    mv.getRange(1, 1, 1, HEAD.length).setValues([HEAD]).setFontWeight('bold');
    mv.setFrozenRows(1);
    mv.getRange('B:B').setNumberFormat('dd.MM.yyyy HH:mm');
    mv.getRange('J:K').setNumberFormat('dd.MM.yyyy HH:mm');
  }
}

/* ---------- веб-приложение ---------- */
function doGet(e) { return respond(e.parameter || {}); }
function doPost(e) {
  let p = {};
  try { p = JSON.parse(e.postData.contents || '{}'); } catch (err) {}
  return respond(p);
}

function respond(p) {
  let out;
  try { out = { ok: true, data: route(p) }; }
  catch (err) { out = { ok: false, error: String(err.message || err) }; }
  return ContentService.createTextOutput(JSON.stringify(out))
    .setMimeType(ContentService.MimeType.JSON);
}

function route(p) {
  switch (p.action) {
    case 'config':  return getConfig();
    case 'list':    return listFor(p.point);
    case 'subscribe': return withLock(() => subscribe(p.point, p.token));
    case 'create':  return notifyAfter(withLock(() => createTransfer(p)));
    case 'accept':  return notifyAfter(withLock(() => setStatus(p.id, p.point, 'to', ST.NEW, ST.ACCEPTED)));
    case 'reject':  return notifyAfter(withLock(() => setStatus(p.id, p.point, 'to', ST.NEW, ST.REJECTED, { reason: p.reason })));
    case 'cancel':  return notifyAfter(withLock(() => setStatus(p.id, p.point, 'from', ST.NEW, ST.CANCELLED)));
    case 'receive': return notifyAfter(withLock(() => setStatus(p.id, p.point, 'to', ST.ACCEPTED, ST.RECEIVED, { items: p.items })));
    default: throw new Error('Неизвестное действие');
  }
}

function withLock(fn) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try { return fn(); } finally { lock.releaseLock(); }
}

/* ---------- справочники ---------- */
function getConfig() {
  const ss = SpreadsheetApp.getActive();
  const points = ss.getSheetByName(SHEET_POINTS).getDataRange().getValues().slice(1)
    .filter(r => r[0] && r[1])
    .map(r => ({ city: String(r[0]).trim(), point: String(r[1]).trim() }));
  const products = ss.getSheetByName(SHEET_PRODUCTS).getDataRange().getValues().slice(1)
    .filter(r => r[0])
    .map(r => ({ name: String(r[0]).trim(), unit: String(r[1] || 'шт').trim() }));
  return { points, products };
}

function cityOf(point) {
  const p = getConfig().points.find(x => x.point === point);
  if (!p) throw new Error('Точка «' + point + '» не найдена');
  return p.city;
}

/* ---------- чтение ---------- */
function movesSheet() { return SpreadsheetApp.getActive().getSheetByName(SHEET_MOVES); }

function listFor(point) {
  if (!point) throw new Error('Не выбрана точка');
  const rows = movesSheet().getDataRange().getValues().slice(1);
  const since = Date.now() - HISTORY_DAYS * 864e5;
  const map = {};
  rows.forEach(r => {
    if (r[C['Откуда']] !== point && r[C['Куда']] !== point) return;
    const id = r[C['ID']];
    if (!map[id]) {
      map[id] = {
        id,
        created: iso(r[C['Создана']]),
        city: r[C['Город']],
        from: r[C['Откуда']],
        to: r[C['Куда']],
        status: r[C['Статус']],
        accepted: iso(r[C['Принята']]),
        received: iso(r[C['Получена']]),
        comment: r[C['Комментарий']],
        reason: r[C['Причина отказа']],
        items: []
      };
    }
    map[id].items.push({
      product: r[C['Товар']],
      qty: r[C['Кол-во']],
      unit: r[C['Ед.']],
      got: r[C['Получено (факт)']] === '' ? null : r[C['Получено (факт)']]
    });
  });
  return Object.values(map)
    .filter(t => t.status === ST.NEW || t.status === ST.ACCEPTED || new Date(t.created).getTime() >= since)
    .sort((a, b) => b.created.localeCompare(a.created));
}

function iso(v) { return v instanceof Date ? v.toISOString() : (v || ''); }

/* ---------- запись ---------- */
function createTransfer(p) {
  if (!p.from || !p.to) throw new Error('Не выбрана точка');
  if (p.from === p.to) throw new Error('Нельзя переместить на свою же точку');
  const city = cityOf(p.from);
  if (cityOf(p.to) !== city) throw new Error('Перемещение возможно только внутри города');
  const items = (p.items || []).filter(i => i.product && Number(i.qty) > 0);
  if (!items.length) throw new Error('Добавьте хотя бы один товар');

  const now = new Date();
  const id = 'П-' + Utilities.formatDate(now, 'Europe/Moscow', 'ddMM-HHmmss') + '-' +
    Math.floor(Math.random() * 90 + 10);
  const units = {};
  getConfig().products.forEach(x => units[x.name] = x.unit);

  const rows = items.map(i => {
    const r = new Array(HEAD.length).fill('');
    r[C['ID']] = id; r[C['Создана']] = now; r[C['Город']] = city;
    r[C['Откуда']] = p.from; r[C['Куда']] = p.to;
    r[C['Товар']] = i.product; r[C['Кол-во']] = Number(i.qty);
    r[C['Ед.']] = units[i.product] || 'шт';
    r[C['Статус']] = ST.NEW; r[C['Комментарий']] = p.comment || '';
    return r;
  });
  const sh = movesSheet();
  sh.getRange(sh.getLastRow() + 1, 1, rows.length, HEAD.length).setValues(rows);
  return { id, status: ST.NEW, from: p.from, to: p.to,
    items: rows.map(r => ({ product: r[C['Товар']], qty: r[C['Кол-во']], unit: r[C['Ед.']] })) };
}

/**
 * Меняет статус всех строк заявки.
 * side: какая точка имеет право на действие — 'to' (получатель) или 'from' (отправитель).
 */
function setStatus(id, point, side, fromStatus, toStatus, extra) {
  extra = extra || {};
  const sh = movesSheet();
  const data = sh.getDataRange().getValues();
  const idx = [];
  for (let i = 1; i < data.length; i++) if (data[i][C['ID']] === id) idx.push(i);
  if (!idx.length) throw new Error('Заявка не найдена');

  const first = data[idx[0]];
  const owner = side === 'to' ? first[C['Куда']] : first[C['Откуда']];
  if (owner !== point) throw new Error('Это действие доступно только точке «' + owner + '»');
  if (first[C['Статус']] !== fromStatus) throw new Error('Заявка уже в статусе «' + first[C['Статус']] + '»');

  const now = new Date();
  const got = {};
  (extra.items || []).forEach(i => got[i.product] = Number(i.qty));

  const items = [];
  idx.forEach(i => {
    const r = data[i];
    r[C['Статус']] = toStatus;
    if (toStatus === ST.ACCEPTED) r[C['Принята']] = now;
    if (toStatus === ST.REJECTED) r[C['Причина отказа']] = extra.reason || '';
    if (toStatus === ST.RECEIVED) {
      r[C['Получена']] = now;
      const fact = got.hasOwnProperty(r[C['Товар']]) ? got[r[C['Товар']]] : Number(r[C['Кол-во']]);
      r[C['Получено (факт)']] = fact;
      r[C['Расхождение']] = fact - Number(r[C['Кол-во']]);
    }
    sh.getRange(i + 1, 1, 1, HEAD.length).setValues([r]);
    items.push({ product: r[C['Товар']], qty: r[C['Кол-во']], unit: r[C['Ед.']], got: r[C['Получено (факт)']] });
  });
  return { id, status: toStatus, from: first[C['Откуда']], to: first[C['Куда']], items, reason: extra.reason || '' };
}

/* =====================================================================
 * Пуш-уведомления через Firebase Cloud Messaging (HTTP v1).
 * Настройка: Project Settings (шестерёнка) → Свойства скрипта →
 *   FCM_SERVICE_ACCOUNT = весь JSON ключа сервисного аккаунта Firebase.
 * Если свойства нет — заявки работают, просто без пушей.
 * ===================================================================== */

function subscribe(point, token) {
  if (!point || !token) throw new Error('Нет точки или токена');
  cityOf(point); // проверка, что точка существует
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEET_SUBS);
  const data = sh.getDataRange().getValues();
  // один телефон = одна точка: убираем этот токен у других точек
  for (let i = data.length - 1; i >= 1; i--) if (data[i][1] === token) sh.deleteRow(i + 1);
  sh.appendRow([point, token, new Date()]);
  return { ok: true };
}

function itemsText_(items, withGot) {
  return items.map(i => {
    const got = withGot && i.got !== '' && Number(i.got) !== Number(i.qty) ? ' (пришло ' + i.got + ')' : '';
    return i.product + ' ' + i.qty + ' ' + i.unit + got;
  }).join(', ');
}

/** Решает, кому и что отправить после действия, и не ломает ответ при ошибке пушей. */
function notifyAfter(t) {
  try {
    let point, title, body;
    if (t.status === ST.NEW) { point = t.to; title = 'Новая заявка от ' + t.from; body = itemsText_(t.items) + '. Нажмите, чтобы принять.'; }
    else if (t.status === ST.ACCEPTED) { point = t.from; title = t.to + ' приняла заявку'; body = 'Передайте товар: ' + itemsText_(t.items); }
    else if (t.status === ST.REJECTED) { point = t.from; title = t.to + ' отклонила заявку'; body = t.reason ? 'Причина: ' + t.reason : itemsText_(t.items); }
    else if (t.status === ST.CANCELLED) { point = t.to; title = t.from + ' отменила заявку'; body = itemsText_(t.items); }
    else if (t.status === ST.RECEIVED) {
      point = t.from; title = t.to + ' получила товар';
      const diff = t.items.filter(i => Number(i.got) !== Number(i.qty));
      body = diff.length ? 'Расхождение: ' + itemsText_(diff, true) : 'Всё сошлось.';
    }
    if (point) sendPush_(point, title, body, t.id);
  } catch (err) {
    console.error('Пуш не отправлен: ' + err);
  }
  return { id: t.id, status: t.status };
}

function sendPush_(point, title, body, id) {
  const sa = PropertiesService.getScriptProperties().getProperty('FCM_SERVICE_ACCOUNT');
  if (!sa) return; // пуши не настроены
  const key = JSON.parse(sa);
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEET_SUBS);
  const rows = sh.getDataRange().getValues().slice(1).filter(r => r[0] === point && r[1]);
  if (!rows.length) return;

  const access = fcmAccessToken_(key);
  const url = 'https://fcm.googleapis.com/v1/projects/' + key.project_id + '/messages:send';
  const reqs = rows.map(r => ({
    url, method: 'post', contentType: 'application/json', muteHttpExceptions: true,
    headers: { Authorization: 'Bearer ' + access },
    payload: JSON.stringify({ message: {
      token: r[1],
      data: { title, body, id: String(id || ''), tag: String(id || 'peremeshcheniya') },
      webpush: { headers: { Urgency: 'high', TTL: '86400' } },
      android: { priority: 'high' }
    } })
  }));
  const res = UrlFetchApp.fetchAll(reqs);

  // убираем токены удалённых/переустановленных приложений
  const dead = [];
  res.forEach((r, i) => {
    const code = r.getResponseCode(), txt = r.getContentText();
    if (code === 404 || (code === 400 && /UNREGISTERED|registration token/i.test(txt))) dead.push(rows[i][1]);
    else if (code >= 300) console.warn('FCM ' + code + ': ' + txt);
  });
  if (dead.length) {
    const data = sh.getDataRange().getValues();
    for (let i = data.length - 1; i >= 1; i--) if (dead.indexOf(data[i][1]) >= 0) sh.deleteRow(i + 1);
  }
}

function fcmAccessToken_(key) {
  const cache = CacheService.getScriptCache();
  const cached = cache.get('fcm_access');
  if (cached) return cached;
  const now = Math.floor(Date.now() / 1000);
  const b64 = o => Utilities.base64EncodeWebSafe(Utilities.newBlob(JSON.stringify(o)).getBytes()).replace(/=+$/, '');
  const unsigned = b64({ alg: 'RS256', typ: 'JWT' }) + '.' + b64({
    iss: key.client_email,
    scope: 'https://www.googleapis.com/auth/firebase.messaging',
    aud: 'https://oauth2.googleapis.com/token',
    iat: now, exp: now + 3600
  });
  const sig = Utilities.base64EncodeWebSafe(Utilities.computeRsaSha256Signature(unsigned, key.private_key)).replace(/=+$/, '');
  const r = UrlFetchApp.fetch('https://oauth2.googleapis.com/token', {
    method: 'post',
    payload: { grant_type: 'urn:ietf:params:oauth:grant-type:jwt-bearer', assertion: unsigned + '.' + sig }
  });
  const token = JSON.parse(r.getContentText()).access_token;
  cache.put('fcm_access', token, 3000);
  return token;
}

/** Проверка пушей вручную: впишите свою точку и запустите. */
function testPush() {
  sendPush_('Точка 1', 'Проверка уведомлений', 'Если вы это видите — пуши работают.', 'test');
}
