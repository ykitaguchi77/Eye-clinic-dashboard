/**
 * 入院・病床管理 v0.1
 * Googleスプレッドシート（コンテナバインド）＋ Apps Script ウェブアプリ
 *
 * 初回：node setup.mjs で作成・公開 → URLを開くと初期設定が自動で走る → 各マスタを入力
 */

const SH = {
  BOOK: '予約台帳',
  PROC: '術式マスタ',
  BED: 'ベッドマスタ',
  SLOT: '手術枠マスタ',
  USER: 'ユーザー',
  LOG: '変更履歴',
  CONF: '設定'
};
const BOOK_HEAD = ['予約ID', '患者ID', '術式コード', '入院日', 'ベッド番号', '状態', 'メモ',
  '登録者', '登録日時', '更新者', '更新日時', 'カレンダーイベントID'];
const TZ = Session.getScriptTimeZone() || 'Asia/Tokyo';
const ROLES_ALL = ['管理', '入力', '閲覧'];
const ROLES_EDIT = ['管理', '入力'];

/* ------------------------------------------------------------
 * 初期設定（エディタから1回だけ実行）
 * 既にあるシートは上書きしません
 * ---------------------------------------------------------- */
function setup() {
  const ss = SpreadsheetApp.getActive();
  // 実行者（ウェブアプリ初回アクセス時はデプロイした人）を最初の管理者にする
  const me = Session.getEffectiveUser().getEmail() || Session.getActiveUser().getEmail();

  const book = makeSheet_(ss, SH.BOOK, BOOK_HEAD, []);
  if (book) {
    book.getRange('B:B').setNumberFormat('@');          // 患者IDの先頭0を保持
    book.getRange('D2:D').setNumberFormat('yyyy-mm-dd');
    book.getRange('F2:F').setDataValidation(
      SpreadsheetApp.newDataValidation().requireValueInList(['仮', '確定', '取消'], true).build());
  }

  const proc = makeSheet_(ss, SH.PROC,
    ['術式コード', '表示名', '入院パターン', '泊数', '眼数', 'カテゴリ', '単価(円)', '有効'], [
      ['CAT1', '白内障 片眼', '1泊2日', 1, 1, '白内障', 0, true],
      ['CAT2', '白内障 両眼', '1泊2日', 1, 2, '白内障', 0, true],
      ['CAT3', '白内障 両眼', '2泊3日', 2, 2, '白内障', 0, true],
      ['LID1', '眼瞼下垂 両側', '1泊2日', 1, 2, '眼瞼', 0, true],
      ['DCR1', '涙嚢鼻腔吻合術', '2泊3日', 2, 1, '涙道', 0, true],
      ['VIT1', '硝子体手術', '4泊5日', 4, 1, '網膜硝子体', 0, true]
    ]);
  if (proc) proc.getRange('H2:H7').insertCheckboxes();

  const bed = makeSheet_(ss, SH.BED, ['ベッド番号', '名称', '有効'],
    [1, 2, 3, 4, 5, 6].map(n => [n, 'ベッド' + n, true]));
  if (bed) bed.getRange('C2:C7').insertCheckboxes();

  makeSheet_(ss, SH.SLOT, ['曜日', '手術枠数'],
    ['月', '火', '水', '木', '金', '土', '日'].map(w => [w, 0]));

  const user = makeSheet_(ss, SH.USER, ['メールアドレス', '氏名', '権限'], [[me, '管理者', '管理']]);
  if (user) user.getRange('C2:C').setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(ROLES_ALL, true).build());

  makeSheet_(ss, SH.CONF, ['項目', '値', '説明'], [
    ['施設名', '', '画面上部に表示（任意）'],
    ['原価率', 0.55, '利益見込みの初期値（0〜1）'],
    ['カレンダーID', '', '入れるとGoogleカレンダーへ自動反映（空欄なら反映しない）']
  ]);

  makeSheet_(ss, SH.LOG, ['日時', 'ユーザー', '操作', '予約ID', '内容'], []);

  const def = ss.getSheetByName('シート1') || ss.getSheetByName('Sheet1');
  if (def && def.getLastRow() === 0 && ss.getSheets().length > 1) ss.deleteSheet(def);

  try { ss.toast('初期設定が完了しました。各マスタを入力してください。'); } catch (e) { /* ウェブアプリから実行した場合は表示しない */ }
}

function makeSheet_(ss, name, head, rows) {
  if (ss.getSheetByName(name)) return null;
  const sh = ss.insertSheet(name);
  sh.getRange(1, 1, 1, head.length).setValues([head]).setFontWeight('bold').setBackground('#EEEBE5');
  if (rows.length) sh.getRange(2, 1, rows.length, rows[0].length).setValues(rows);
  sh.setFrozenRows(1);
  return sh;
}

/* ------------------------------------------------------------
 * ウェブアプリ入口
 * ---------------------------------------------------------- */
function doGet() {
  ensureSetup_();
  const u = currentUser_();
  if (!u) {
    const email = Session.getActiveUser().getEmail() || '（取得できませんでした）';
    return HtmlService.createHtmlOutput(
      '<div style="font-family:sans-serif;padding:32px;line-height:1.8">' +
      '<h2>利用権限がありません</h2>' +
      '<p>アカウント ' + escapeHtml_(email) + ' は登録されていません。<br>' +
      '管理者にスプレッドシートの「ユーザー」シートへの追加を依頼してください。</p></div>')
      .setTitle('入院・病床管理');
  }
  return page_()
    .setTitle('入院・病床管理')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/**
 * 画面のHTML。通常は Index.html を使う。
 * 1ファイル版（gas/Code.gs）では INDEX_HTML_ に同じ内容が埋め込まれている。
 */
function page_() {
  if (typeof INDEX_HTML_ !== 'undefined') return HtmlService.createHtmlOutput(INDEX_HTML_);
  return HtmlService.createHtmlOutputFromFile('Index');
}

/** シートがまだ無ければ初期設定を自動実行（初回アクセス時） */
function ensureSetup_() {
  if (SpreadsheetApp.getActive().getSheetByName(SH.USER)) return;
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    if (!SpreadsheetApp.getActive().getSheetByName(SH.USER)) setup();
  } finally {
    lock.releaseLock();
  }
}

/* ------------------------------------------------------------
 * クライアントから呼ぶ関数
 * ---------------------------------------------------------- */
function getBootstrap() {
  const u = require_(ROLES_ALL);
  const c = conf_();
  const admin = u.role === '管理';
  return {
    user: u,
    facility: String(c['施設名'] || ''),
    procs: procs_().map(p => admin ? p : Object.assign({}, p, { price: null })),
    beds: beds_(),
    costRate: Number(c['原価率']) || 0,
    today: ymd_(new Date())
  };
}

/** from〜to（yyyy-MM-dd）に在院日がかかる予約（取消を除く） */
function getBookings(from, to) {
  require_(ROLES_ALL);
  if (!isYmd_(from) || !isYmd_(to)) throw new Error('日付の形式が正しくありません。');
  const pm = procMap_();
  return bookings_()
    .filter(b => b.status !== '取消' && b.date)
    .filter(b => overlap_(b.date, endOf_(b, pm), from, to))
    .map(b => ({ id: b.id, patientId: b.patientId, proc: b.proc, date: b.date,
                 bed: b.bed, status: b.status, memo: b.memo }));
}

/** 新規登録（p.id なし）または変更（p.id あり） */
function saveBooking(p) {
  const u = require_(ROLES_EDIT);
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const pm = procMap_();
    const proc = pm[p.proc];
    const pid = String(p.patientId || '').trim();
    if (!pid) return { ok: false, message: '患者IDを入力してください。' };
    if (!isYmd_(p.date)) return { ok: false, message: '入院日を入力してください。' };
    if (!proc || !proc.active) return { ok: false, message: '術式を選んでください。' };
    const status = p.status === '確定' ? '確定' : '仮';
    const memo = String(p.memo || '').slice(0, 200);

    const all = bookings_();
    const start = p.date;
    const end = addDays_(start, proc.nights);
    const bedNos = beds_().map(b => b.no);
    const others = all.filter(b => b.status !== '取消' && b.id !== p.id);
    const busy = no => others.some(b => b.bed === no && overlap_(start, end, b.date, endOf_(b, pm)));

    let bed = Number(p.bed) || 0;
    if (bed) {
      if (bedNos.indexOf(bed) < 0) return { ok: false, message: 'そのベッドは使用できません。' };
      if (busy(bed)) return { ok: false, message: 'ベッド' + bed + 'はこの期間ふさがっています。' };
    } else {
      bed = bedNos.find(no => !busy(no)) || 0;
      if (!bed) return { ok: false, message: 'この期間に空いているベッドがありません。' };
    }

    const sh = sheet_(SH.BOOK);
    const now = new Date();
    let id, row, eventId = '';
    if (p.id) {
      const cur = all.find(b => b.id === p.id);
      if (!cur) return { ok: false, message: '予約が見つかりません。画面を更新してください。' };
      id = cur.id; row = cur.row; eventId = cur.eventId;
      sh.getRange(row, 2, 1, 6).setValues([[pid, p.proc, parseYmd_(start), bed, status, memo]]);
      sh.getRange(row, 10, 1, 2).setValues([[u.email, now]]);
      log_(u, '変更', id, describe_(pid, proc, start, bed, status));
    } else {
      id = 'B' + Utilities.formatDate(now, TZ, 'yyMMddHHmmss') + Math.floor(Math.random() * 90 + 10);
      sh.appendRow([id, pid, p.proc, parseYmd_(start), bed, status, memo, u.email, now, u.email, now, '']);
      row = sh.getLastRow();
      log_(u, '登録', id, describe_(pid, proc, start, bed, status));
    }
    syncCalendar_(row, eventId, calTitle_(pid, proc, bed, status), start, end, false);
    return { ok: true, id: id, bed: bed,
             message: pid + ' をベッド' + bed + 'に' + (p.id ? '変更しました' : '登録しました') };
  } finally {
    lock.releaseLock();
  }
}

function cancelBooking(id) {
  const u = require_(ROLES_EDIT);
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const b = bookings_().find(x => x.id === id);
    if (!b) return { ok: false, message: '予約が見つかりません。' };
    const sh = sheet_(SH.BOOK);
    sh.getRange(b.row, 6).setValue('取消');
    sh.getRange(b.row, 10, 1, 2).setValues([[u.email, new Date()]]);
    log_(u, '取消', id, b.patientId + ' ' + b.date);
    syncCalendar_(b.row, b.eventId, '', b.date, b.date, true);
    return { ok: true, message: b.patientId + ' の予約を取り消しました' };
  } finally {
    lock.releaseLock();
  }
}

/** 月次の収益見込み材料（管理者のみ）。埋まり率などの試算は画面側で計算 */
function getMonthSummary(ym) {
  require_(['管理']);
  if (!/^\d{4}-\d{2}$/.test(ym || '')) throw new Error('対象月の形式が正しくありません。');
  const y = Number(ym.slice(0, 4)), m = Number(ym.slice(5, 7));
  const nDays = new Date(y, m, 0).getDate();
  const first = ymd_(new Date(y, m - 1, 1)), last = ymd_(new Date(y, m - 1, nDays));
  const today = ymd_(new Date());
  const pm = procMap_();
  const slots = slots_();
  const bedNos = beds_().map(b => b.no);
  const live = bookings_().filter(b => b.status !== '取消' && b.date);
  const inMonth = live.filter(b => b.date >= first && b.date <= last);

  const sum = { confirmed: { rev: 0, count: 0, eyes: 0 }, tentative: { rev: 0, count: 0, eyes: 0 } };
  const byCat = {};
  const missing = {};
  inMonth.forEach(b => {
    const p = pm[b.proc] || { name: b.proc, price: 0, eyes: 0, cat: 'その他' };
    const k = b.status === '確定' ? 'confirmed' : 'tentative';
    sum[k].rev += p.price; sum[k].count++; sum[k].eyes += p.eyes;
    if (!p.price) missing[p.name] = true;
    byCat[p.cat] = byCat[p.cat] || { cat: p.cat, confirmed: 0, tentative: 0, count: 0 };
    byCat[p.cat][k] += p.price; byCat[p.cat].count++;
  });

  // 手術枠：入院日＝手術日として、今日以降の空き枠を数える
  const perDay = {};
  live.forEach(b => { perDay[b.date] = (perDay[b.date] || 0) + 1; });
  let remaining = 0, totalSlots = 0, occupied = 0;
  for (let d = 1; d <= nDays; d++) {
    const dt = new Date(y, m - 1, d);
    const ds = ymd_(dt);
    const s = Number(slots['日月火水木金土'[dt.getDay()]]) || 0;
    totalSlots += s;
    if (ds >= today) remaining += Math.max(0, s - (perDay[ds] || 0));
    occupied += live.filter(b => bedNos.indexOf(b.bed) >= 0 && b.date <= ds && endOf_(b, pm) >= ds).length;
  }

  const monthPrices = inMonth.map(b => (pm[b.proc] || {}).price || 0).filter(x => x > 0);
  const masterPrices = Object.keys(pm).map(k => pm[k]).filter(p => p.active && p.price > 0).map(p => p.price);
  const mean = a => a.length ? a.reduce((x, y) => x + y, 0) / a.length : 0;

  return {
    confirmed: sum.confirmed,
    tentative: sum.tentative,
    remainingSlots: remaining,
    totalSlots: totalSlots,
    avgPrice: monthPrices.length ? mean(monthPrices) : mean(masterPrices),
    occupied: occupied,
    capacity: bedNos.length * nDays,
    byCat: Object.keys(byCat).map(k => byCat[k]),
    missingPrice: Object.keys(missing)
  };
}

/* ------------------------------------------------------------
 * 内部関数
 * ---------------------------------------------------------- */
function currentUser_() {
  const email = (Session.getActiveUser().getEmail() || '').toLowerCase();
  if (!email) return null;
  const r = rows_(SH.USER).find(x => String(x[0]).trim().toLowerCase() === email);
  if (!r) return null;
  const role = String(r[2] || '閲覧').trim();
  return { email: email, name: String(r[1] || email), role: ROLES_ALL.indexOf(role) >= 0 ? role : '閲覧' };
}

function require_(roles) {
  const u = currentUser_();
  if (!u || roles.indexOf(u.role) < 0) throw new Error('この操作の権限がありません。');
  return u;
}

function sheet_(name) {
  const sh = SpreadsheetApp.getActive().getSheetByName(name);
  if (!sh) throw new Error('シート「' + name + '」がありません。setup を実行してください。');
  return sh;
}

function rows_(name) {
  return sheet_(name).getDataRange().getValues().slice(1)
    .filter(r => r.some(c => c !== '' && c !== null));
}

function conf_() {
  const o = {};
  rows_(SH.CONF).forEach(r => { o[String(r[0]).trim()] = r[1]; });
  return o;
}

function truthy_(v) { return v === true || String(v).toUpperCase() === 'TRUE'; }

function procs_() {
  return rows_(SH.PROC).map(r => ({
    code: String(r[0]).trim(), name: String(r[1]), pattern: String(r[2]),
    nights: Number(r[3]) || 0, eyes: Number(r[4]) || 0, cat: String(r[5] || 'その他'),
    price: Number(r[6]) || 0, active: truthy_(r[7])
  })).filter(p => p.code);
}

function procMap_() {
  const m = {};
  procs_().forEach(p => { m[p.code] = p; });
  return m;
}

function beds_() {
  return rows_(SH.BED).filter(r => truthy_(r[2]) && Number(r[0]))
    .map(r => ({ no: Number(r[0]), name: String(r[1] || ('ベッド' + r[0])) }))
    .sort((a, b) => a.no - b.no);
}

function slots_() {
  const o = {};
  rows_(SH.SLOT).forEach(r => { o[String(r[0]).trim()] = Number(r[1]) || 0; });
  return o;
}

function bookings_() {
  const v = sheet_(SH.BOOK).getDataRange().getValues();
  const out = [];
  for (let i = 1; i < v.length; i++) {
    const r = v[i];
    if (!r[0]) continue;
    out.push({
      id: String(r[0]), patientId: String(r[1]), proc: String(r[2]).trim(), date: ymd_(r[3]),
      bed: Number(r[4]) || 0, status: String(r[5] || '仮'), memo: String(r[6] || ''),
      eventId: String(r[11] || ''), row: i + 1
    });
  }
  return out;
}

/** 在院の最終日（退院日）。退院日も同じベッドを使う扱い */
function endOf_(b, pm) { return addDays_(b.date, (pm[b.proc] || { nights: 0 }).nights); }
function overlap_(a1, a2, b1, b2) { return a1 <= b2 && b1 <= a2; }
function isYmd_(s) { return /^\d{4}-\d{2}-\d{2}$/.test(s || ''); }
function ymd_(d) {
  if (d instanceof Date) return Utilities.formatDate(d, TZ, 'yyyy-MM-dd');
  return d ? String(d).slice(0, 10) : '';
}
function parseYmd_(s) { const a = s.split('-').map(Number); return new Date(a[0], a[1] - 1, a[2]); }
function addDays_(s, n) { const d = parseYmd_(s); d.setDate(d.getDate() + n); return ymd_(d); }

function describe_(pid, proc, start, bed, status) {
  return pid + ' ' + proc.name + ' ' + proc.pattern + ' 入院' + start + ' ベッド' + bed + ' ' + status;
}
function calTitle_(pid, proc, bed, status) {
  return 'ベッド' + bed + ' ' + pid + ' ' + proc.name + ' ' + proc.pattern + (status === '仮' ? '［仮］' : '');
}

function log_(u, action, id, detail) {
  sheet_(SH.LOG).appendRow([new Date(), u.email, action, id, detail]);
}

/** 設定シートにカレンダーIDがあれば終日予定として反映 */
function syncCalendar_(row, eventId, title, start, end, remove) {
  const calId = String(conf_()['カレンダーID'] || '').trim();
  if (!calId) return;
  try {
    const cal = CalendarApp.getCalendarById(calId);
    if (!cal) throw new Error('カレンダーが見つかりません: ' + calId);
    const sh = sheet_(SH.BOOK);
    let ev = eventId ? cal.getEventById(eventId) : null;
    if (remove) {
      if (ev) ev.deleteEvent();
      sh.getRange(row, 12).setValue('');
      return;
    }
    const s = parseYmd_(start), e = parseYmd_(addDays_(end, 1));
    if (ev) {
      ev.setTitle(title);
      ev.setAllDayDates(s, e);
    } else {
      ev = cal.createAllDayEvent(title, s, e);
      sh.getRange(row, 12).setValue(ev.getId());
    }
  } catch (err) {
    log_({ email: 'system' }, 'カレンダー反映失敗', '', String(err));
  }
}

function escapeHtml_(s) {
  return String(s).replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
}
