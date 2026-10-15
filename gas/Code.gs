/**
 * ============================================================
 *  入院・病床管理ダッシュボード（1ファイル版）
 *  このファイルは自動生成です。直接編集せず src/ を編集して
 *  node tools/build-gas.mjs で作り直してください。
 *
 *  使い方：スプレッドシートの「拡張機能」→「Apps Script」を開き、
 *  最初からあるコードをすべて消して、このファイルの中身を貼り付けて保存。
 *  あとは「デプロイ」→「新しいデプロイ」→ ウェブアプリ として公開するだけ。
 * ============================================================
 */

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

/* ------------------------------------------------------------
 * 画面のHTML（src/Index.html と同じ内容）
 * ---------------------------------------------------------- */
const INDEX_HTML_ = `<!DOCTYPE html>
<html lang="ja">
<head>
<base target="_top">
<meta charset="utf-8">
<link rel="preconnect" href="https://fonts.googleapis.com">
<link href="https://fonts.googleapis.com/css2?family=Zen+Kaku+Gothic+New:wght@400;500;700&display=swap" rel="stylesheet">
<style>
:root{--ink:#1F2328;--ground:#EEEBE5;--panel:#FFFFFF;--line:#DDD7CC;--muted:#4A4F57;--accent:#B45309;--sel:#1F4E79;--bedcol:84px}
*{box-sizing:border-box}
body{margin:0;background:var(--ground);color:var(--ink);font-family:'Zen Kaku Gothic New','Hiragino Kaku Gothic ProN','Yu Gothic','Meiryo',sans-serif;font-variant-numeric:tabular-nums}
button,input,select{font-family:inherit}
:focus-visible{outline:3px solid var(--sel);outline-offset:2px}
header{display:flex;align-items:center;gap:20px;padding:12px 24px;background:var(--ink);color:#F4F1EA;flex-wrap:wrap}
.brand{display:flex;align-items:baseline;gap:12px}
header h1{margin:0;font-size:20px;letter-spacing:.04em}
.brand span{color:#C9C4BA;font-size:14px}
nav{display:flex;gap:8px}
.tab{height:40px;padding:0 16px;border-radius:8px;border:1px solid #5A5F66;background:transparent;color:#F4F1EA;font-size:15px;cursor:pointer}
.tab[aria-pressed="true"]{background:#F4F1EA;color:var(--ink);border-color:#F4F1EA;font-weight:700}
#who{margin-left:auto;font-size:13px;color:#C9C4BA}
main{padding:20px 24px}
#viewCal{display:grid;grid-template-columns:340px minmax(0,1fr);gap:20px;align-items:start}
#viewCal.readonly{grid-template-columns:minmax(0,1fr)}
@media (max-width:1000px){#viewCal{grid-template-columns:minmax(0,1fr)}}
.panel{background:var(--panel);border:1px solid var(--line);border-radius:12px;padding:18px 20px}
h2{margin:0 0 14px;font-size:17px}
label{display:flex;flex-direction:column;gap:6px;font-size:13px;color:var(--muted);font-weight:500;margin-bottom:12px}
label.inline{flex-direction:row;align-items:center;gap:6px;margin:0;font-size:14px;color:var(--ink)}
input[type=text],input[type=date],input[type=month],select{height:42px;border:1px solid #CFC8BC;border-radius:8px;padding:0 10px;font-size:16px;background:#FBFAF7;color:var(--ink)}
fieldset{border:none;padding:0;margin:0 0 12px}
legend{font-size:13px;color:var(--muted);font-weight:500;margin-bottom:6px;padding:0}
.proc-grid{display:grid;grid-template-columns:repeat(2,minmax(0,1fr));gap:8px}
.proc{min-height:54px;text-align:left;padding:8px 10px;border-radius:8px;border:1px solid #CFC8BC;background:#fff;color:var(--ink);cursor:pointer}
.proc b{display:block;font-size:14px}
.proc small{font-size:12px}
.proc[aria-pressed="true"]{background:var(--sel);border-color:var(--sel);color:#fff}
.radios{display:flex;gap:18px}
.preview{border-radius:10px;padding:12px 14px;font-size:14px;margin:4px 0 12px;line-height:1.6}
.preview:empty{display:none}
.preview.ok{border:2px dashed var(--accent);background:#FFF4E8;color:#5C2D05}
.preview.ng{border:2px solid #B42318;background:#FDECEA;color:#7A1A12}
.preview strong{font-size:18px;display:block}
.actions{display:flex;flex-direction:column;gap:8px}
.primary{height:48px;border:none;border-radius:10px;background:var(--accent);color:#fff;font-size:16px;font-weight:700;cursor:pointer}
.primary:disabled{background:#E3DED5;color:#6B6F75;cursor:not-allowed}
.danger{height:42px;border-radius:8px;border:1px solid #B42318;background:#fff;color:#B42318;font-size:14px;cursor:pointer}
.ghost{height:40px;padding:0 14px;border-radius:8px;border:1px solid #CFC8BC;background:#fff;color:var(--ink);font-size:14px;cursor:pointer}
#msg{min-height:1.4em;font-size:13px;margin:10px 0 0;line-height:1.5}
#msg.err{color:#B42318}
#msg.ok{color:#1D5A2C}
.cal-toolbar{display:flex;align-items:center;gap:8px;flex-wrap:wrap;margin-bottom:12px}
#range{font-weight:700;margin:0 8px}
.legend{margin-left:auto;display:flex;gap:12px;font-size:12px;color:var(--muted);flex-wrap:wrap}
.legend i{display:inline-block;width:12px;height:12px;border-radius:3px;margin-right:4px;vertical-align:-1px;border:1px solid}
.cal-scroll{overflow-x:auto}
#grid{min-width:900px}
.g-row{display:grid;grid-template-columns:var(--bedcol) minmax(0,1fr);border-bottom:1px solid #EEE9E0}
.g-bed{font-weight:700;font-size:14px;display:flex;align-items:center}
.g-track{display:grid}
.g-day{text-align:center;padding:6px 0;font-size:13px;border-left:1px solid #EEE9E0}
.g-day small{display:block;font-size:12px}
.g-day.sat{color:#1F4E79}
.g-day.sun{color:#9A2A1C}
.g-day.today{background:#FFF4E8;font-weight:700}
.g-cell{grid-row:1;height:48px;border:none;border-left:1px solid #EEE9E0;background:transparent;padding:0}
.g-cell.wk{background:#F7F5F0}
button.g-cell{cursor:pointer}
button.g-cell:hover{background:#FFF4E8}
.g-bar{grid-row:1;z-index:1;align-self:center;height:36px;margin:0 3px;border-radius:6px;border:1px solid transparent;font-size:12px;font-weight:700;text-align:left;padding:0 8px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;cursor:pointer}
.g-bar.tent{border-style:dashed;background-image:repeating-linear-gradient(135deg,rgba(255,255,255,0) 0 6px,rgba(255,255,255,.6) 6px 12px)}
.g-bar.cutL{border-top-left-radius:0;border-bottom-left-radius:0;margin-left:0}
.g-bar.cutR{border-top-right-radius:0;border-bottom-right-radius:0;margin-right:0}
.g-bar.pending{background:#FFF4E8;border:2px dashed var(--accent);color:#7A3A06;cursor:default;z-index:2}
.g-bar.editing{outline:3px solid var(--sel);outline-offset:1px}
.g-usage{display:grid;font-size:13px;color:var(--muted);padding-top:6px;text-align:center}
.g-usage .full{color:#9A3412;font-weight:700}
.rev-controls{display:grid;grid-template-columns:repeat(4,minmax(0,1fr));gap:16px;align-items:end;margin-bottom:16px}
@media (max-width:1000px){.rev-controls{grid-template-columns:repeat(2,minmax(0,1fr))}}
.rev-controls label{margin:0}
.rev-controls output{font-size:16px;font-weight:700;color:var(--ink)}
input[type=range]{width:100%;accent-color:var(--sel)}
.warn{background:#FFF4E8;border:1px solid var(--accent);color:#5C2D05;border-radius:10px;padding:10px 14px;font-size:14px;margin-bottom:16px;line-height:1.6}
.kpis{display:grid;grid-template-columns:repeat(4,minmax(0,1fr));gap:14px;margin-bottom:16px}
@media (max-width:1000px){.kpis{grid-template-columns:repeat(2,minmax(0,1fr))}}
.kpi{background:var(--panel);border:1px solid var(--line);border-radius:12px;padding:14px 16px}
.kpi .l{font-size:13px;color:var(--muted);font-weight:500}
.kpi .v{font-size:26px;font-weight:700;margin:4px 0}
.kpi .s{font-size:12px;color:var(--muted)}
.kpi.main{background:var(--ink);color:#F4F1EA;border-color:var(--ink)}
.kpi.main .l,.kpi.main .s{color:#C9C4BA}
.cat{display:grid;grid-template-columns:110px minmax(0,1fr) 140px 60px;gap:12px;align-items:center;margin-bottom:10px;font-size:14px}
.cat .track{height:22px;background:#F3F0EA;border-radius:4px;overflow:hidden;display:flex}
.cat .r{text-align:right}
.note{font-size:13px;color:var(--muted);line-height:1.7;margin:12px 0 0}
#loading{padding:40px;text-align:center;color:var(--muted)}
</style>
</head>
<body>
<header>
  <div class="brand"><h1>入院・病床管理</h1><span id="facility"></span></div>
  <nav aria-label="画面切替">
    <button type="button" class="tab" id="tabCal" aria-pressed="true">病床カレンダー</button>
    <button type="button" class="tab" id="tabRev" aria-pressed="false" hidden>収益見込み</button>
  </nav>
  <div id="who"></div>
</header>

<main>
  <div id="loading">読み込み中…</div>

  <section id="viewCal" hidden>
    <aside id="formPanel" class="panel">
      <h2 id="formTitle">入院予約の登録</h2>
      <label>患者ID<input type="text" id="fPid" autocomplete="off" placeholder="例：0012345"></label>
      <label>入院日（＝手術日）<input type="date" id="fDate"></label>
      <fieldset><legend>術式</legend><div id="procBtns" class="proc-grid"></div></fieldset>
      <label>ベッド<select id="fBed"></select></label>
      <fieldset><legend>状態</legend>
        <div class="radios">
          <label class="inline"><input type="radio" name="st" value="仮" checked>仮押さえ</label>
          <label class="inline"><input type="radio" name="st" value="確定">確定</label>
        </div>
      </fieldset>
      <label>メモ<input type="text" id="fMemo" maxlength="200" autocomplete="off"></label>
      <div id="preview" class="preview"></div>
      <div class="actions">
        <button type="button" id="btnSave" class="primary">登録する</button>
        <button type="button" id="btnCancelBk" class="danger" hidden>この予約を取り消す</button>
        <button type="button" id="btnNew" class="ghost" hidden>新規入力に戻る</button>
      </div>
      <p id="msg" role="status"></p>
    </aside>

    <div class="panel">
      <div class="cal-toolbar">
        <button type="button" class="ghost" id="prev">前の週</button>
        <button type="button" class="ghost" id="thisWeek">今週</button>
        <button type="button" class="ghost" id="next">次の週</button>
        <span id="range"></span>
        <span class="legend" id="legend"></span>
      </div>
      <div class="cal-scroll"><div id="grid"></div></div>
    </div>
  </section>

  <section id="viewRev" hidden>
    <div class="panel rev-controls">
      <label>対象月<input type="month" id="rMonth"></label>
      <label>残り枠の埋まり率 <output id="rFillOut">70%</output><input type="range" id="rFill" min="0" max="100" step="5" value="70"></label>
      <label>原価率 <output id="rCostOut"></output><input type="range" id="rCost" min="0" max="100" step="1"></label>
      <label class="inline"><input type="checkbox" id="rTent" checked>仮押さえを売上に含める</label>
    </div>
    <div id="revWarn" class="warn" hidden></div>
    <div id="kpis" class="kpis"></div>
    <div class="panel">
      <h2>カテゴリ別の売上（確定＋仮押さえ）</h2>
      <div id="cats"></div>
      <p class="note">売上見込み ＝ 確定分 ＋ 仮押さえ分 ＋ 今日以降の空き手術枠 × 埋まり率 × 平均単価。平均単価はその月の予約の平均（予約がなければ術式マスタの平均）です。単価・手術枠・原価率はスプレッドシートのマスタで変更できます。</p>
    </div>
  </section>
</main>

<script>
const DAYS = 14;
const CAT = {
  '白内障': { bg: '#D6E4F0', fg: '#123A5A' },
  '眼瞼': { bg: '#F3DEC9', fg: '#6B3410' },
  '涙道': { bg: '#D9E6D2', fg: '#28451D' },
  '網膜硝子体': { bg: '#E2DAEE', fg: '#3D2A5C' }
};
const CAT_OTHER = { bg: '#E6E2DA', fg: '#33373D' };
const S = { boot: null, procMap: {}, view: null, bookings: [], selProc: null, editingId: null, rev: null };

const $ = id => document.getElementById(id);
const esc = s => String(s == null ? '' : s).replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
const call = (fn, ...a) => new Promise((res, rej) => google.script.run.withSuccessHandler(res).withFailureHandler(rej)[fn](...a));
const pd = s => { const [y, m, d] = s.split('-').map(Number); return new Date(y, m - 1, d); };
const fd = d => d.getFullYear() + '-' + String(d.getMonth() + 1).padStart(2, '0') + '-' + String(d.getDate()).padStart(2, '0');
const add = (s, n) => { const d = pd(s); d.setDate(d.getDate() + n); return fd(d); };
const diff = (a, b) => Math.round((pd(b) - pd(a)) / 864e5);
const md = s => { const d = pd(s); return (d.getMonth() + 1) + '/' + d.getDate() + '(' + '日月火水木金土'[d.getDay()] + ')'; };
const monday = s => { const d = pd(s); d.setDate(d.getDate() - ((d.getDay() + 6) % 7)); return fd(d); };
const yen = n => '¥' + Math.round(n || 0).toLocaleString('ja-JP');
const catColor = c => CAT[c] || CAT_OTHER;
const procOf = code => S.procMap[code];
const endOf = b => add(b.date, (procOf(b.proc) || { nights: 0 }).nights);
const overlap = (a1, a2, b1, b2) => a1 <= b2 && b1 <= a2;
const canEdit = () => ['管理', '入力'].includes(S.boot.user.role);
const bedName = no => (S.boot.beds.find(b => b.no === no) || { name: 'ベッド' + no }).name;

function setMsg(text, kind) { const el = $('msg'); el.textContent = text || ''; el.className = kind || ''; }
function errText(e) { return (e && e.message) ? e.message : String(e); }

async function init() {
  try {
    S.boot = await call('getBootstrap');
  } catch (e) {
    $('loading').textContent = errText(e);
    return;
  }
  S.boot.procs.forEach(p => { S.procMap[p.code] = p; });
  $('facility').textContent = S.boot.facility;
  $('who').textContent = S.boot.user.name + '（' + S.boot.user.role + '）';
  $('tabRev').hidden = S.boot.user.role !== '管理';
  if (!canEdit()) { $('formPanel').hidden = true; $('viewCal').classList.add('readonly'); }

  S.view = monday(S.boot.today);
  renderProcButtons();
  renderBedOptions();
  renderLegend();
  resetForm();
  bind();

  $('loading').hidden = true;
  $('viewCal').hidden = false;
  await loadBookings();
}

function bind() {
  $('tabCal').onclick = () => showView('cal');
  $('tabRev').onclick = () => showView('rev');
  $('prev').onclick = () => moveView(-7);
  $('next').onclick = () => moveView(7);
  $('thisWeek').onclick = () => { S.view = monday(S.boot.today); loadBookings(); };
  $('fDate').onchange = onDateChange;
  $('fBed').onchange = refresh;
  $('fPid').oninput = refreshPreview;
  document.querySelectorAll('input[name=st]').forEach(r => { r.onchange = refresh; });
  $('btnSave').onclick = save;
  $('btnCancelBk').onclick = cancelBk;
  $('btnNew').onclick = () => { resetForm(); refresh(); };
  $('grid').onclick = onGridClick;
  $('rMonth').onchange = loadRev;
  $('rFill').oninput = renderRev;
  $('rCost').oninput = renderRev;
  $('rTent').onchange = renderRev;
}

function showView(v) {
  $('tabCal').setAttribute('aria-pressed', v === 'cal');
  $('tabRev').setAttribute('aria-pressed', v === 'rev');
  $('viewCal').hidden = v !== 'cal';
  $('viewRev').hidden = v !== 'rev';
  if (v === 'rev') {
    if (!$('rMonth').value) {
      $('rMonth').value = S.boot.today.slice(0, 7);
      $('rCost').value = Math.round(S.boot.costRate * 100);
    }
    loadRev();
  }
}

function moveView(n) { S.view = add(S.view, n); loadBookings(); }

async function loadBookings() {
  $('range').textContent = md(S.view) + ' 〜 ' + md(add(S.view, DAYS - 1));
  try {
    S.bookings = await call('getBookings', add(S.view, -14), add(S.view, DAYS + 14));
  } catch (e) {
    setMsg(errText(e), 'err');
    S.bookings = [];
  }
  refresh();
}

/* ---------- 入力フォーム ---------- */
function renderProcButtons() {
  $('procBtns').innerHTML = S.boot.procs.filter(p => p.active).map(p =>
    '<button type="button" class="proc" data-code="' + esc(p.code) + '" aria-pressed="false">' +
    '<b>' + esc(p.name) + '</b><small>' + esc(p.pattern) + '</small></button>').join('');
  $('procBtns').onclick = e => {
    const b = e.target.closest('.proc');
    if (!b) return;
    S.selProc = b.dataset.code;
    refresh();
  };
}

function renderBedOptions() {
  $('fBed').innerHTML = '<option value="0">自動（空いているベッド）</option>' +
    S.boot.beds.map(b => '<option value="' + b.no + '">' + esc(b.name) + '</option>').join('');
}

function renderLegend() {
  const cats = [...new Set(S.boot.procs.filter(p => p.active).map(p => p.cat))];
  $('legend').innerHTML = cats.map(c => {
    const k = catColor(c);
    return '<span><i style="background:' + k.bg + ';border-color:' + k.fg + '"></i>' + esc(c) + '</span>';
  }).join('') + '<span><i style="background:#fff;border-style:dashed;border-color:#4A4F57"></i>斜線＝仮押さえ</span>';
}

function formVals() {
  return {
    id: S.editingId,
    patientId: $('fPid').value.trim(),
    date: $('fDate').value,
    proc: S.selProc,
    bed: Number($('fBed').value) || 0,
    status: document.querySelector('input[name=st]:checked').value,
    memo: $('fMemo').value
  };
}

function resetForm() {
  S.editingId = null;
  S.selProc = null;
  $('fPid').value = '';
  $('fDate').value = '';
  $('fBed').value = '0';
  $('fMemo').value = '';
  document.querySelector('input[name=st][value="仮"]').checked = true;
  $('formTitle').textContent = '入院予約の登録';
  $('btnSave').textContent = '登録する';
  $('btnCancelBk').hidden = true;
  $('btnNew').hidden = true;
}

function onDateChange() {
  const d = $('fDate').value;
  if (d && (d < S.view || d > add(S.view, DAYS - 1))) { S.view = monday(d); loadBookings(); }
  else refresh();
}

function alloc(f) {
  if (!f.date || !f.proc || !procOf(f.proc)) return null;
  const p = procOf(f.proc);
  const s = f.date, e = add(s, p.nights);
  const others = S.bookings.filter(b => b.id !== f.id);
  const busy = no => others.some(b => b.bed === no && overlap(s, e, b.date, endOf(b)));
  if (f.bed) {
    return busy(f.bed) ? { ng: true, msg: bedName(f.bed) + 'はこの期間ふさがっています。自動を選ぶか、日付を変えてください。' }
                       : { bed: f.bed, s, e, p };
  }
  const free = S.boot.beds.find(b => !busy(b.no));
  return free ? { bed: free.no, s, e, p } : { ng: true, msg: 'この期間に空いているベッドがありません。日付を変えてください。' };
}

function refresh() {
  document.querySelectorAll('.proc').forEach(b => b.setAttribute('aria-pressed', b.dataset.code === S.selProc));
  refreshPreview();
  renderGrid();
}

function refreshPreview() {
  const el = $('preview');
  const a = canEdit() ? alloc(formVals()) : null;
  if (!a) { el.className = 'preview'; el.innerHTML = ''; $('btnSave').disabled = false; return; }
  if (a.ng) { el.className = 'preview ng'; el.textContent = a.msg; $('btnSave').disabled = true; return; }
  el.className = 'preview ok';
  el.innerHTML = '<strong>' + esc(bedName(a.bed)) + '</strong>' + md(a.s) + ' 〜 ' + md(a.e) + '（' + esc(a.p.pattern) + '）';
  $('btnSave').disabled = false;
}

async function save() {
  const f = formVals();
  $('btnSave').disabled = true;
  setMsg('保存しています…');
  try {
    const r = await call('saveBooking', f);
    if (!r.ok) { setMsg(r.message, 'err'); $('btnSave').disabled = false; return; }
    setMsg(r.message, 'ok');
    resetForm();
    await loadBookings();
  } catch (e) {
    setMsg(errText(e), 'err');
    $('btnSave').disabled = false;
  }
}

async function cancelBk() {
  if (!S.editingId) return;
  const b = S.bookings.find(x => x.id === S.editingId);
  if (!confirm((b ? b.patientId + ' ' : '') + 'の予約を取り消しますか？')) return;
  try {
    const r = await call('cancelBooking', S.editingId);
    setMsg(r.message, r.ok ? 'ok' : 'err');
    if (r.ok) { resetForm(); await loadBookings(); }
  } catch (e) { setMsg(errText(e), 'err'); }
}

function editBooking(id) {
  const b = S.bookings.find(x => x.id === id);
  if (!b) return;
  const p = procOf(b.proc);
  if (!canEdit()) {
    alert(b.patientId + '\\n' + (p ? p.name + ' ' + p.pattern : b.proc) + '\\n' + md(b.date) + '〜' + md(endOf(b)) + '　' + bedName(b.bed) + '\\n状態：' + b.status + (b.memo ? '\\nメモ：' + b.memo : ''));
    return;
  }
  S.editingId = id;
  S.selProc = b.proc;
  $('fPid').value = b.patientId;
  $('fDate').value = b.date;
  $('fBed').value = String(b.bed);
  $('fMemo').value = b.memo;
  document.querySelector('input[name=st][value="' + (b.status === '確定' ? '確定' : '仮') + '"]').checked = true;
  $('formTitle').textContent = '予約の変更';
  $('btnSave').textContent = '変更を保存';
  $('btnCancelBk').hidden = false;
  $('btnNew').hidden = false;
  setMsg('');
  refresh();
  $('formPanel').scrollIntoView({ behavior: 'smooth', block: 'start' });
}

/* ---------- 病床カレンダー ---------- */
function renderGrid() {
  const vs = S.view, ve = add(vs, DAYS - 1);
  const cols = 'grid-template-columns:repeat(' + DAYS + ',minmax(0,1fr))';
  const days = Array.from({ length: DAYS }, (_, i) => add(vs, i));
  const f = canEdit() ? formVals() : {};
  const a = canEdit() ? alloc(f) : null;
  const usage = new Array(DAYS).fill(0);
  let h = '';

  h += '<div class="g-row"><div class="g-bed" style="font-weight:500;font-size:12px;color:#4A4F57">病床</div><div class="g-track" style="' + cols + '">' +
    days.map(d => {
      const w = pd(d).getDay();
      const cls = 'g-day' + (w === 6 ? ' sat' : '') + (w === 0 ? ' sun' : '') + (d === S.boot.today ? ' today' : '');
      return '<div class="' + cls + '">' + (pd(d).getMonth() + 1) + '/' + pd(d).getDate() + '<small>' + '日月火水木金土'[w] + '</small></div>';
    }).join('') + '</div></div>';

  S.boot.beds.forEach(bed => {
    let t = '';
    days.forEach((d, i) => {
      const w = pd(d).getDay();
      const wk = (w === 0 || w === 6) ? ' wk' : '';
      t += canEdit()
        ? '<button type="button" class="g-cell' + wk + '" style="grid-column:' + (i + 1) + '" data-act="cell" data-bed="' + bed.no + '" data-date="' + d + '" aria-label="' + esc(bed.name) + ' ' + md(d) + 'を入院日にする"></button>'
        : '<div class="g-cell' + wk + '" style="grid-column:' + (i + 1) + '"></div>';
    });
    S.bookings.filter(b => b.bed === bed.no).forEach(b => {
      const e = endOf(b);
      if (e < vs || b.date > ve) return;
      const si = Math.max(0, diff(vs, b.date)), ei = Math.min(DAYS - 1, diff(vs, e));
      for (let i = si; i <= ei; i++) usage[i]++;
      const p = procOf(b.proc);
      const c = catColor(p ? p.cat : '');
      const tent = b.status !== '確定';
      const cls = 'g-bar' + (tent ? ' tent' : '') + (b.date < vs ? ' cutL' : '') + (e > ve ? ' cutR' : '') + (b.id === S.editingId ? ' editing' : '');
      const label = b.patientId + '　' + (p ? p.name : b.proc) + (tent ? '［仮］' : '');
      const title = label + '　' + md(b.date) + '〜' + md(e) + (b.memo ? '　' + b.memo : '');
      t += '<button type="button" class="' + cls + '" data-act="bar" data-id="' + esc(b.id) + '" style="grid-column:' + (si + 1) + ' / span ' + (ei - si + 1) +
        ';background-color:' + c.bg + ';color:' + c.fg + (tent ? ';border-color:' + c.fg : '') + '" title="' + esc(title) + '">' + esc(label) + '</button>';
    });
    if (a && !a.ng && a.bed === bed.no && !(a.e < vs || a.s > ve)) {
      const si = Math.max(0, diff(vs, a.s)), ei = Math.min(DAYS - 1, diff(vs, a.e));
      for (let i = si; i <= ei; i++) usage[i]++;
      t += '<div class="g-bar pending" style="grid-column:' + (si + 1) + ' / span ' + (ei - si + 1) + '">' + (S.editingId ? '変更後' : '新規') + '　' + esc(a.p.name) + '</div>';
    }
    h += '<div class="g-row"><div class="g-bed">' + esc(bed.name) + '</div><div class="g-track" style="' + cols + '">' + t + '</div></div>';
  });

  const nb = S.boot.beds.length;
  h += '<div class="g-row" style="border-bottom:none"><div class="g-bed" style="font-weight:500;font-size:12px;color:#4A4F57">使用数</div><div class="g-usage" style="' + cols + '">' +
    usage.map(u => '<div class="' + (u >= nb ? 'full' : '') + '">' + u + '/' + nb + '</div>').join('') + '</div></div>';
  $('grid').innerHTML = h;
}

function onGridClick(e) {
  const t = e.target.closest('[data-act]');
  if (!t) return;
  if (t.dataset.act === 'bar') editBooking(t.dataset.id);
  if (t.dataset.act === 'cell' && canEdit()) {
    $('fDate').value = t.dataset.date;
    $('fBed').value = t.dataset.bed;
    refresh();
  }
}

/* ---------- 収益見込み ---------- */
async function loadRev() {
  const ym = $('rMonth').value;
  if (!ym) return;
  $('kpis').innerHTML = '<div class="kpi"><div class="l">読み込み中…</div></div>';
  try {
    S.rev = await call('getMonthSummary', ym);
    renderRev();
  } catch (e) {
    $('kpis').innerHTML = '<div class="kpi"><div class="l">' + esc(errText(e)) + '</div></div>';
  }
}

function renderRev() {
  const d = S.rev;
  const fill = Number($('rFill').value) / 100;
  const cost = Number($('rCost').value) / 100;
  $('rFillOut').textContent = Math.round(fill * 100) + '%';
  $('rCostOut').textContent = Math.round(cost * 100) + '%';
  if (!d) return;
  const incl = $('rTent').checked;
  const base = d.confirmed.rev + (incl ? d.tentative.rev : 0);
  const extra = d.remainingSlots * fill * d.avgPrice;
  const total = base + extra;
  const occ = d.capacity ? Math.round(100 * d.occupied / d.capacity) : 0;
  const cnt = d.confirmed.count + d.tentative.count;
  const eyes = d.confirmed.eyes + d.tentative.eyes;

  const warn = [];
  if (d.missingPrice.length) warn.push('単価が0円の術式があります（' + d.missingPrice.map(esc).join('、') + '）。術式マスタで単価を入れると売上に反映されます。');
  if (!d.totalSlots) warn.push('手術枠マスタが0のため、空き枠からの上乗せは計算されません。曜日ごとの手術枠数を入れてください。');
  $('revWarn').hidden = !warn.length;
  $('revWarn').innerHTML = warn.join('<br>');

  const k = [
    { l: '売上見込み（月）', v: yen(total), s: '確定＋' + (incl ? '仮押さえ＋' : '') + '空き枠の上乗せ', main: true },
    { l: '利益見込み（月）', v: yen(total * (1 - cost)), s: '原価率 ' + Math.round(cost * 100) + '% で計算', main: true },
    { l: '確定分の売上', v: yen(d.confirmed.rev), s: d.confirmed.count + '件' },
    { l: '仮押さえ分の売上', v: yen(d.tentative.rev), s: d.tentative.count + '件' + (incl ? '' : '（見込みに含めない）') },
    { l: '残りの手術枠', v: d.remainingSlots + '枠', s: '月間 ' + d.totalSlots + '枠のうち、今日以降の空き' },
    { l: '空き枠からの上乗せ', v: yen(extra), s: '埋まり率 ' + Math.round(fill * 100) + '% × 平均単価 ' + yen(d.avgPrice) },
    { l: '病床稼働率', v: occ + '%', s: '延べ ' + d.occupied + '床日 ／ ' + d.capacity + '床日' },
    { l: '手術件数（予約済み）', v: cnt + '件', s: eyes + '眼' }
  ];
  $('kpis').innerHTML = k.map(x => '<div class="kpi' + (x.main ? ' main' : '') + '"><div class="l">' + x.l + '</div><div class="v">' + x.v + '</div><div class="s">' + x.s + '</div></div>').join('');

  const max = Math.max(1, ...d.byCat.map(c => c.confirmed + c.tentative));
  $('cats').innerHTML = d.byCat.length ? d.byCat.map(c => {
    const col = catColor(c.cat);
    const w1 = 100 * c.confirmed / max, w2 = 100 * c.tentative / max;
    return '<div class="cat"><div><b>' + esc(c.cat) + '</b></div><div class="track">' +
      '<div style="width:' + w1 + '%;background:' + col.fg + '"></div>' +
      '<div style="width:' + w2 + '%;background:' + col.bg + '"></div></div>' +
      '<div class="r">' + yen(c.confirmed + c.tentative) + '</div><div class="r">' + c.count + '件</div></div>';
  }).join('') + '<p class="note">濃い色＝確定、薄い色＝仮押さえ</p>'
    : '<p class="note">この月の予約はまだありません。</p>';
}

init();
</script>
</body>
</html>
`;
