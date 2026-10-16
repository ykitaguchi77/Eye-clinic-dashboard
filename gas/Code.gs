/**
 * ============================================================
 *  入院・手術・発注管理（1ファイル版）
 *  このファイルは自動生成です。直接編集せず src/ を編集して
 *  node tools/build-gas.mjs で作り直してください。
 *
 *  使い方：スプレッドシートの「拡張機能」→「Apps Script」を開き、
 *  最初からあるコードをすべて消して、このファイルの中身を貼り付けて保存。
 *  あとは「デプロイ」→「新しいデプロイ」→ ウェブアプリ として公開するだけ。
 * ============================================================
 */

/**
 * 入院・手術・発注管理 v0.2
 * Googleスプレッドシート（コンテナバインド）＋ Apps Script ウェブアプリ
 *
 * ・入院／外来手術の登録と病床カレンダー
 * ・薬品・医療材料の発注入力
 * ・収益表、支出合計、発注集計、発注書の Excel 出力
 *
 * 初回：ウェブアプリを公開してURLを開くと、必要なシートが自動で作られる → 各マスタを入力
 */

const SH = {
  BOOK: '手術・入院台帳',
  PROC: '術式マスタ',
  CATG: '収益区分マスタ',
  DOC: '医師マスタ',
  BED: 'ベッドマスタ',
  SLOT: '手術枠マスタ',
  VEND: '発注先マスタ',
  ITEM: '品目マスタ',
  ORDER: '発注台帳',
  USER: 'ユーザー',
  CONF: '設定',
  LOG: '変更履歴'
};

const HEAD = {
  BOOK: ['予約ID', '患者ID', '術式コード', '区分', '手術日', '眼', '執刀医', '入院日', '泊数', 'ベッド番号',
    '状態', 'メモ', '登録者', '登録日時', '更新者', '更新日時', 'カレンダーイベントID'],
  PROC: ['術式コード', '表示名', '収益区分', '標準泊数', '請求額(片眼)', '請求額(両眼)', 'コスト(片眼)', 'コスト(両眼)', '色分け', '有効'],
  CATG: ['区分名', '1眼あたり収入(円)', 'トリプル', '入院表に常に表示', '外来表に常に表示', '有効'],
  DOC: ['氏名', '有効'],
  BED: ['ベッド番号', '名称', '有効'],
  SLOT: ['曜日', '手術枠数'],
  VEND: ['発注先', '発注書の宛名', '入力方法', '手術コストに含める', '有効'],
  ITEM: ['品目ID', '発注先', '区分', '小区分', '商品コード', '品名', 'メーカー', '規格', '入数', '薬価', '納入価', '後発', '有効'],
  ORDER: ['発注ID', '発注日', '発注先', '品目ID', '品名', '区分', '小区分', '規格', '入数', '単価', '数量', '金額', 'メモ',
    '状態', '登録者', '登録日時', '更新者', '更新日時'],
  USER: ['メールアドレス', '氏名', '権限'],
  CONF: ['項目', '値', '説明'],
  LOG: ['日時', 'ユーザー', '操作', 'ID', '内容']
};

// 台帳の列番号（0始まり）
const B = { id: 0, pid: 1, proc: 2, kind: 3, op: 4, eye: 5, doc: 6, admit: 7, nights: 8, bed: 9, status: 10, memo: 11,
  cBy: 12, cAt: 13, uBy: 14, uAt: 15, ev: 16 };
const O = { id: 0, date: 1, vendor: 2, item: 3, name: 4, group: 5, sub: 6, spec: 7, pack: 8, price: 9, qty: 10, amount: 11,
  memo: 12, status: 13, cBy: 14, cAt: 15, uBy: 16, uAt: 17 };

const TZ = Session.getScriptTimeZone() || 'Asia/Tokyo';
const ROLES_ALL = ['管理', '入力', '閲覧'];
const ROLES_EDIT = ['管理', '入力'];
const KINDS = ['入院', '外来'];
const EYES = ['右', '左', '両'];
const WEEKDAYS = '日月火水木金土';

/* ------------------------------------------------------------
 * 初期設定（シートがなければ作る。既にあるシートは上書きしない）
 * ---------------------------------------------------------- */
function setup() {
  const ss = SpreadsheetApp.getActive();
  // 実行者（ウェブアプリ初回アクセス時はデプロイした人）を最初の管理者にする
  const me = Session.getEffectiveUser().getEmail() || Session.getActiveUser().getEmail();

  const book = makeSheet_(ss, SH.BOOK, HEAD.BOOK, []);
  if (book) {
    book.getRange('B:B').setNumberFormat('@');          // 患者IDの先頭0を保持
    book.getRange('E2:E').setNumberFormat('yyyy-mm-dd');
    book.getRange('H2:H').setNumberFormat('yyyy-mm-dd');
  }

  const proc = makeSheet_(ss, SH.PROC, HEAD.PROC, [
    ['CAT', '白内障', '白内障', 1, 0, '', 0, '', '白内障', true],
    ['MF', '多焦点眼内レンズ', '多焦点', 1, 0, '', 0, '', '白内障', true],
    ['VIT', '硝子体手術', '硝子体シングル', 2, 0, '', 0, '', '網膜硝子体', true],
    ['VIT3', '硝子体手術（白内障同時）', '硝子体トリプル', 2, 0, '', 0, '', '網膜硝子体', true],
    ['PTO', '眼瞼下垂', '下垂片眼', 0, 0, '', 0, '', '眼瞼', true],
    ['DCR', '涙嚢鼻腔吻合術', '涙道', 2, 0, '', 0, '', '涙道', true]
  ]);
  if (proc) checkboxes_(proc, 10, 300);

  const catg = makeSheet_(ss, SH.CATG, HEAD.CATG, [
    ['多焦点', 0, false, true, true, true],
    ['白内障', 0, false, true, true, true],
    ['硝子体シングル', 0, false, true, true, true],
    ['硝子体トリプル', 0, true, true, true, true],
    ['下垂片眼', 0, false, false, true, true],
    ['涙道', 0, false, false, true, true]
  ]);
  if (catg) { [3, 4, 5, 6].forEach(c => checkboxes_(catg, c, 100)); }

  const doc = makeSheet_(ss, SH.DOC, HEAD.DOC, []);
  if (doc) checkboxes_(doc, 2, 50);

  const bed = makeSheet_(ss, SH.BED, HEAD.BED, [1, 2, 3, 4, 5, 6].map(n => [n, 'ベッド' + n, true]));
  if (bed) checkboxes_(bed, 3, 50);

  makeSheet_(ss, SH.SLOT, HEAD.SLOT, ['月', '火', '水', '木', '金', '土', '日'].map(w => [w, 0]));

  const vend = makeSheet_(ss, SH.VEND, HEAD.VEND, [
    ['薬品卸', '', '品目', false, true],
    ['医療材料卸', '', '品目', false, true],
    ['レンズ', '', '金額', true, true]
  ]);
  if (vend) {
    vend.getRange('C2:C').setDataValidation(
      SpreadsheetApp.newDataValidation().requireValueInList(['品目', '金額'], true).build());
    [4, 5].forEach(c => checkboxes_(vend, c, 50));
  }

  const item = makeSheet_(ss, SH.ITEM, HEAD.ITEM, []);
  if (item) {
    item.getRange('E:E').setNumberFormat('@');          // 商品コードの先頭0を保持
    [12, 13].forEach(c => checkboxes_(item, c, 1000));
  }

  const order = makeSheet_(ss, SH.ORDER, HEAD.ORDER, []);
  if (order) order.getRange('B2:B').setNumberFormat('yyyy-mm-dd');

  const user = makeSheet_(ss, SH.USER, HEAD.USER, [[me, '管理者', '管理']]);
  if (user) user.getRange('C2:C').setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(ROLES_ALL, true).build());

  makeSheet_(ss, SH.CONF, HEAD.CONF, [
    ['施設名', '', '画面上部と出力ファイルに表示（任意）'],
    ['入院1人あたり収入(円)', 0, '収益表の「入院」の1人あたり収入'],
    ['カレンダーID', '', '入れると Googleカレンダーへ自動反映（空欄なら反映しない）']
  ]);

  makeSheet_(ss, SH.LOG, HEAD.LOG, []);

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

function checkboxes_(sh, col, lastRow) {
  sh.getRange(2, col, lastRow - 1, 1).insertCheckboxes();
}

/** シートが1つでも欠けていれば初期設定を実行（初回アクセス時・版の更新時） */
function ensureSetup_() {
  const ss = SpreadsheetApp.getActive();
  const missing = () => Object.keys(SH).some(k => !ss.getSheetByName(SH[k]));
  if (!missing()) return;
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    if (missing()) setup();
  } finally {
    lock.releaseLock();
  }
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
      .setTitle('入院・手術・発注管理');
  }
  return page_()
    .setTitle('入院・手術・発注管理')
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

/* ------------------------------------------------------------
 * 画面から呼ぶ関数：共通
 * ---------------------------------------------------------- */
function getBootstrap() {
  const u = require_(ROLES_ALL);
  checkHeaders_();
  const admin = u.role === '管理';
  const edit = ROLES_EDIT.indexOf(u.role) >= 0;
  if (edit) ensureItemIds_();
  const c = conf_();
  const procs = procs_().map(p => admin ? p : { code: p.code, name: p.name, cat: p.cat, nights: p.nights, color: p.color, active: p.active });
  return {
    user: u,
    facility: str_(c['施設名']),
    procs: procs,
    cats: admin ? cats_() : cats_().map(x => ({ name: x.name, active: x.active })),
    doctors: doctors_(),
    beds: beds_(),
    vendors: edit ? vendors_() : [],
    items: edit ? items_() : [],
    perInpatient: admin ? num_(c['入院1人あたり収入(円)']) : null,
    today: ymd_(new Date())
  };
}

/* ------------------------------------------------------------
 * 入院・手術
 * ---------------------------------------------------------- */

/** from〜to（yyyy-MM-dd）に在院日または手術日がかかる記録（取消を除く） */
function getBookings(from, to) {
  require_(ROLES_ALL);
  if (!isYmd_(from) || !isYmd_(to)) throw new Error('日付の形式が正しくありません。');
  return bookings_()
    .filter(b => b.status !== '取消')
    .filter(b => (b.kind === '入院' && overlap_(b.admit, endOf_(b), from, to)) || (b.opDate >= from && b.opDate <= to))
    .map(publicBooking_);
}

function publicBooking_(b) {
  return { id: b.id, patientId: b.patientId, proc: b.proc, kind: b.kind, opDate: b.opDate, eye: b.eye, doctor: b.doctor,
    admit: b.admit, nights: b.nights, bed: b.bed, status: b.status, memo: b.memo };
}

/** 新規登録（p.id なし）または変更（p.id あり） */
function saveBooking(p) {
  const u = require_(ROLES_EDIT);
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const v = validateBooking_(p || {});
    if (v.error) return { ok: false, message: v.error };
    const all = bookings_();
    if (p.id && !all.some(b => b.id === p.id)) return { ok: false, message: '記録が見つかりません。画面を更新してください。' };

    let bed = 0;
    if (v.kind === '入院') {
      const bedNos = beds_().map(b => b.no);
      const end = addDays_(v.admit, v.nights);
      const others = all.filter(b => b.status !== '取消' && b.kind === '入院' && b.id !== p.id);
      const busy = no => others.some(b => b.bed === no && overlap_(v.admit, end, b.admit, endOf_(b)));
      bed = Number(p.bed) || 0;
      if (bed) {
        if (bedNos.indexOf(bed) < 0) return { ok: false, message: 'そのベッドは使用できません。' };
        if (busy(bed)) return { ok: false, message: 'ベッド' + bed + 'はこの期間ふさがっています。' };
      } else {
        bed = bedNos.find(no => !busy(no)) || 0;
        if (!bed) return { ok: false, message: 'この期間に空いているベッドがありません。' };
      }
    }

    const sh = sheet_(SH.BOOK);
    const now = new Date();
    const rowVals = [v.patientId, v.proc.code, v.kind, parseYmd_(v.opDate), v.eye, v.doctor,
      v.kind === '入院' ? parseYmd_(v.admit) : '', v.kind === '入院' ? v.nights : '', v.kind === '入院' ? bed : '',
      v.status, v.memo];
    let id, row, eventId = '';
    if (p.id) {
      const cur = all.find(b => b.id === p.id);
      id = cur.id; row = cur.row; eventId = cur.eventId;
      sh.getRange(row, B.pid + 1, 1, rowVals.length).setValues([rowVals]);
      sh.getRange(row, B.uBy + 1, 1, 2).setValues([[u.email, now]]);
      log_(u, '変更', id, describeBooking_(v, bed));
    } else {
      id = newId_('B');
      sh.appendRow([id].concat(rowVals, [u.email, now, u.email, now, '']));
      row = sh.getLastRow();
      log_(u, '登録', id, describeBooking_(v, bed));
    }
    const start = v.kind === '入院' ? v.admit : v.opDate;
    const end = v.kind === '入院' ? addDays_(v.admit, v.nights) : v.opDate;
    syncCalendar_(row, eventId, calTitle_(v, bed), start, end, false);
    const where = v.kind === '入院' ? 'ベッド' + bed + 'に' : '外来手術として';
    return { ok: true, id: id, bed: bed, message: v.patientId + ' を' + where + (p.id ? '変更しました' : '登録しました') };
  } finally {
    lock.releaseLock();
  }
}

function validateBooking_(p) {
  const pm = procMap_();
  const patientId = str_(p.patientId);
  const proc = pm[str_(p.proc)];
  const kind = KINDS.indexOf(p.kind) >= 0 ? p.kind : '';
  const eye = EYES.indexOf(p.eye) >= 0 ? p.eye : '';
  const opDate = str_(p.opDate);
  if (!patientId) return { error: '患者IDを入力してください。' };
  if (!kind) return { error: '入院か外来かを選んでください。' };
  if (!isYmd_(opDate)) return { error: '手術日を入力してください。' };
  if (!proc || !proc.active) return { error: '術式を選んでください。' };
  if (!eye) return { error: '右・左・両眼のどれかを選んでください。' };
  const doctor = str_(p.doctor).slice(0, 50);
  let admit = '', nights = 0;
  if (kind === '入院') {
    admit = str_(p.admit) || opDate;
    if (!isYmd_(admit)) return { error: '入院日の形式が正しくありません。' };
    if (admit > opDate) return { error: '入院日が手術日より後になっています。' };
    nights = Math.round(Number(p.nights));
    if (!isFinite(nights) || nights < 0 || nights > 60) return { error: '泊数を選んでください。' };
    if (addDays_(admit, nights) < opDate) return { error: '手術日が退院日より後になっています。泊数を確認してください。' };
  }
  return {
    patientId: patientId.slice(0, 40), proc: proc, kind: kind, eye: eye, doctor: doctor, opDate: opDate,
    admit: admit, nights: nights, status: p.status === '確定' ? '確定' : '仮', memo: str_(p.memo).slice(0, 200)
  };
}

function cancelBooking(id) {
  const u = require_(ROLES_EDIT);
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const b = bookings_().find(x => x.id === id);
    if (!b) return { ok: false, message: '記録が見つかりません。' };
    const sh = sheet_(SH.BOOK);
    sh.getRange(b.row, B.status + 1).setValue('取消');
    sh.getRange(b.row, B.uBy + 1, 1, 2).setValues([[u.email, new Date()]]);
    log_(u, '取消', id, b.patientId + ' 手術日' + b.opDate);
    syncCalendar_(b.row, b.eventId, '', b.opDate, b.opDate, true);
    return { ok: true, message: b.patientId + ' の記録を取り消しました' };
  } finally {
    lock.releaseLock();
  }
}

/* ------------------------------------------------------------
 * 発注
 * ---------------------------------------------------------- */

/** from〜to の発注（取消を除く）を発注IDごとにまとめて返す */
function getOrders(from, to) {
  require_(ROLES_EDIT);
  if (!isYmd_(from) || !isYmd_(to)) throw new Error('日付の形式が正しくありません。');
  return groupOrders_(orders_().filter(o => o.status !== '取消' && o.date >= from && o.date <= to))
    .sort((a, b) => (b.date + b.id).localeCompare(a.date + a.id));
}

function groupOrders_(lines) {
  const map = {};
  lines.forEach(l => {
    const g = map[l.orderId] || (map[l.orderId] = { id: l.orderId, date: l.date, vendor: l.vendor, memo: l.memo, by: l.cBy, lines: [], total: 0 });
    g.lines.push({ itemId: l.itemId, name: l.name, group: l.group, sub: l.sub, spec: l.spec, pack: l.pack, price: l.price, qty: l.qty, amount: l.amount });
    g.total += l.amount;
  });
  return Object.keys(map).map(k => map[k]);
}

/**
 * 発注を登録（p.id があれば置き換え）
 * p = { id?, date, vendor, memo, lines: [{ itemId, qty, price? }] }   入力方法「品目」の発注先
 * p = { id?, date, vendor, memo, amount }                              入力方法「金額」の発注先
 */
function saveOrder(p) {
  const u = require_(ROLES_EDIT);
  p = p || {};
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const date = str_(p.date);
    if (!isYmd_(date)) return { ok: false, message: '発注日を入力してください。' };
    const vendor = vendors_().find(v => v.name === str_(p.vendor) && v.active);
    if (!vendor) return { ok: false, message: '発注先を選んでください。' };
    const memo = str_(p.memo).slice(0, 200);

    const lines = [];
    if (vendor.mode === '金額') {
      const amount = Math.round(Number(p.amount));
      if (!isFinite(amount) || amount <= 0) return { ok: false, message: '金額を入力してください。' };
      lines.push({ itemId: '', name: memo || vendor.name, group: '', sub: '', spec: '', pack: '', price: amount, qty: 1, amount: amount });
    } else {
      const im = {};
      items_().forEach(it => { if (it.vendor === vendor.name) im[it.id] = it; });
      (p.lines || []).forEach(l => {
        const it = im[str_(l.itemId)];
        const qty = Math.round(Number(l.qty));
        if (!it || !isFinite(qty) || qty <= 0) return;
        const pr = Number(l.price);
        const price = isFinite(pr) && pr >= 0 && str_(l.price) !== '' ? pr : it.price;
        lines.push({ itemId: it.id, name: it.name, group: it.group, sub: it.sub, spec: it.spec, pack: it.pack,
          price: price, qty: qty, amount: Math.round(price * qty) });
      });
      if (!lines.length) return { ok: false, message: '数量が入っている品目がありません。' };
      if (lines.length > 500) return { ok: false, message: '1回の発注は500品目までです。' };
    }

    const sh = sheet_(SH.ORDER);
    const now = new Date();
    let id = str_(p.id);
    if (id) {
      const old = orders_().filter(o => o.orderId === id && o.status !== '取消');
      if (!old.length) return { ok: false, message: '発注が見つかりません。画面を更新してください。' };
      old.forEach(o => {
        sh.getRange(o.row, O.status + 1).setValue('取消');
        sh.getRange(o.row, O.uBy + 1, 1, 2).setValues([[u.email, now]]);
      });
    } else {
      id = newId_('O');
    }
    const rows = lines.map(l => [id, parseYmd_(date), vendor.name, l.itemId, l.name, l.group, l.sub, l.spec, l.pack,
      l.price, l.qty, l.amount, memo, '有効', u.email, now, u.email, now]);
    appendRows_(sh, rows);
    const total = lines.reduce((s, l) => s + l.amount, 0);
    log_(u, p.id ? '発注変更' : '発注', id, vendor.name + ' ' + date + ' ' + lines.length + '品目 ' + total + '円');
    return { ok: true, id: id, message: vendor.name + ' への発注を' + (p.id ? '変更しました' : '登録しました') + '（' + lines.length + '品目）' };
  } finally {
    lock.releaseLock();
  }
}

function cancelOrder(id) {
  const u = require_(ROLES_EDIT);
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const lines = orders_().filter(o => o.orderId === id && o.status !== '取消');
    if (!lines.length) return { ok: false, message: '発注が見つかりません。' };
    const sh = sheet_(SH.ORDER);
    const now = new Date();
    lines.forEach(o => {
      sh.getRange(o.row, O.status + 1).setValue('取消');
      sh.getRange(o.row, O.uBy + 1, 1, 2).setValues([[u.email, now]]);
    });
    log_(u, '発注取消', id, lines[0].vendor + ' ' + lines[0].date);
    return { ok: true, message: lines[0].vendor + ' への発注（' + lines[0].date + '）を取り消しました' };
  } finally {
    lock.releaseLock();
  }
}

/* ------------------------------------------------------------
 * 月次の集計（収益表・支出・発注集計の元データ）
 * ---------------------------------------------------------- */
function getMonthSummary(ym) {
  require_(['管理']);
  return monthData_(ym);
}

function monthData_(ym) {
  if (!/^\d{4}-\d{2}$/.test(ym || '')) throw new Error('対象月の形式が正しくありません。');
  const y = Number(ym.slice(0, 4)), m = Number(ym.slice(5, 7));
  const nDays = new Date(y, m, 0).getDate();
  const days = [];
  for (let d = 1; d <= nDays; d++) days.push(ymd_(new Date(y, m - 1, d)));
  const first = days[0], last = days[nDays - 1];
  const today = ymd_(new Date());
  const c = conf_();
  const pm = procMap_();
  const catList = cats_();
  const catMap = {};
  catList.forEach(x => { catMap[x.name] = x; });
  const live = bookings_().filter(b => b.status !== '取消');

  // 医師（マスタの順＋記録にだけある名前）
  const doctors = doctors_().map(d => d.name);
  const addDoc = n => { if (doctors.indexOf(n) < 0) doctors.push(n); };

  // 入院（人）：入院日がこの月
  const admits = live.filter(b => b.kind === '入院' && b.admit >= first && b.admit <= last);
  const peopleByDoc = {};
  admits.forEach(b => { const d = b.doctor || '未入力'; addDoc(d); peopleByDoc[d] = (peopleByDoc[d] || 0) + 1; });

  // 手術：手術日がこの月
  const ops = live.filter(b => b.opDate >= first && b.opDate <= last);
  const unknownCats = [];
  const table = kind => {
    const rows = {};
    ops.filter(b => b.kind === kind).forEach(b => {
      const p = pm[b.proc];
      const cat = p && p.cat ? p.cat : 'その他';
      if (!catMap[cat] && unknownCats.indexOf(cat) < 0) unknownCats.push(cat);
      const r = rows[cat] || (rows[cat] = { cat: cat, eyes: 0, byDoc: {} });
      const e = eyes_(b.eye);
      r.eyes += e;
      const d = b.doctor || '未入力';
      addDoc(d);
      r.byDoc[d] = (r.byDoc[d] || 0) + e;
    });
    const out = [];
    catList.forEach(x => {
      const shown = kind === '入院' ? x.showIn : x.showOut;
      if (rows[x.name] || (x.active && shown)) {
        const r = rows[x.name] || { cat: x.name, eyes: 0, byDoc: {} };
        out.push({ cat: x.name, unit: x.unit, triple: x.triple, eyes: r.eyes, byDoc: r.byDoc });
      }
    });
    Object.keys(rows).forEach(k => { if (!catMap[k]) out.push({ cat: k, unit: 0, triple: false, eyes: rows[k].eyes, byDoc: rows[k].byDoc }); });
    const eyes = out.reduce((s, r) => s + r.eyes, 0);
    const income = out.reduce((s, r) => s + r.eyes * r.unit, 0);
    return { rows: out, eyes: eyes, income: income };
  };
  const tin = table('入院'), tout = table('外来');
  const tripleEyes = tin.rows.concat(tout.rows).filter(r => r.triple).reduce((s, r) => s + r.eyes, 0);
  const perInpatient = num_(c['入院1人あたり収入(円)']);

  // 術式ごとの件数（コスト計算用）と、日ごとの手術ディスポ代
  const procRows = {};
  const dispoByDay = {};
  let charges = 0;
  ops.forEach(b => {
    const p = pm[b.proc] || { code: b.proc, name: b.proc, charge1: 0, charge2: 0, cost1: 0, cost2: 0 };
    const two = b.eye === '両';
    const r = procRows[p.code] || (procRows[p.code] = { code: p.code, name: p.name, charge1: p.charge1, charge2: p.charge2, cost1: p.cost1, cost2: p.cost2, n1: 0, n2: 0 });
    if (two) r.n2++; else r.n1++;
    dispoByDay[b.opDate] = (dispoByDay[b.opDate] || 0) + (two ? p.cost2 : p.cost1);
    charges += two ? p.charge2 : p.charge1;
  });
  const procList = procs_().filter(p => p.active || procRows[p.code]).map(p => procRows[p.code] ||
    { code: p.code, name: p.name, charge1: p.charge1, charge2: p.charge2, cost1: p.cost1, cost2: p.cost2, n1: 0, n2: 0 });
  Object.keys(procRows).forEach(k => { if (!pm[k]) procList.push(procRows[k]); });

  // 発注：支出（日ごと・発注先ごと）と発注集計（品目×週）
  const vlist = vendors_();
  const monthOrders = orders_().filter(o => o.status !== '取消' && o.date >= first && o.date <= last);
  const vendorNames = vlist.filter(v => v.active).map(v => v.name);
  monthOrders.forEach(o => { if (vendorNames.indexOf(o.vendor) < 0) vendorNames.push(o.vendor); });
  const vendors = vendorNames.map(n => {
    const v = vlist.find(x => x.name === n) || { name: n, mode: '品目', surgeryCost: false };
    return { name: n, mode: v.mode, surgeryCost: v.surgeryCost };
  });
  const spendByDay = {};
  monthOrders.forEach(o => {
    const d = spendByDay[o.date] || (spendByDay[o.date] = {});
    d[o.vendor] = (d[o.vendor] || 0) + o.amount;
  });

  const allItems = items_();
  const orderSummary = vendors.filter(v => v.mode === '品目').map(v => {
    const vItems = allItems.filter(it => it.vendor === v.name);
    const byItem = {};
    const adhoc = {};
    monthOrders.filter(o => o.vendor === v.name).forEach(o => {
      const w = Math.min(5, Math.floor((Number(o.date.slice(8, 10)) - 1) / 7) + 1);
      const known = vItems.find(it => it.id === o.itemId);
      const key = known ? o.itemId : (o.itemId || o.name);
      const bucket = known ? byItem : adhoc;
      const r = bucket[key] || (bucket[key] = { weeks: [0, 0, 0, 0, 0], qty: 0, name: o.name, group: o.group, sub: o.sub, spec: o.spec, pack: o.pack, price: o.price });
      r.weeks[w - 1] += o.amount;
      r.qty += o.qty;
    });
    const rows = vItems.filter(it => it.active || byItem[it.id]).map(it => {
      const r = byItem[it.id] || { weeks: [0, 0, 0, 0, 0], qty: 0 };
      return { id: it.id, group: it.group || 'その他', code: it.code, name: it.name, maker: it.maker, spec: it.spec, pack: it.pack,
        listPrice: it.listPrice, price: it.price, generic: it.generic, weeks: r.weeks, qty: r.qty };
    });
    Object.keys(adhoc).forEach(k => {
      const r = adhoc[k];
      rows.push({ id: '', group: r.group || 'その他', code: '', name: r.name, maker: '', spec: r.spec, pack: r.pack,
        listPrice: '', price: r.price, generic: false, weeks: r.weeks, qty: r.qty });
    });
    return { vendor: v.name, rows: rows };
  });

  // シミュレーション用：手術枠と病床
  const slots = slots_();
  const bedNos = beds_().map(b => b.no);
  const perDay = {};
  live.forEach(b => { perDay[b.opDate] = (perDay[b.opDate] || 0) + 1; });
  let remaining = 0, totalSlots = 0, occupied = 0;
  days.forEach(ds => {
    const s = Number(slots[WEEKDAYS[parseYmd_(ds).getDay()]]) || 0;
    totalSlots += s;
    if (ds >= today) remaining += Math.max(0, s - (perDay[ds] || 0));
    occupied += live.filter(b => b.kind === '入院' && bedNos.indexOf(b.bed) >= 0 && b.admit <= ds && endOf_(b) >= ds).length;
  });
  const cases = ops.length;
  const surgeryIncome = tin.income + tout.income;
  const tentative = ops.filter(b => b.status !== '確定').length;

  return {
    ym: ym, year: y, month: m, days: days, facility: str_(c['施設名']),
    doctors: doctors,
    inpatient: { people: admits.length, peopleByDoc: peopleByDoc, unit: perInpatient, income: admits.length * perInpatient },
    surgery: { in: tin, out: tout, tripleEyes: tripleEyes, cases: cases, tentative: tentative },
    unknownCats: unknownCats,
    charges: charges,
    procRows: procList,
    vendors: vendors,
    spendByDay: spendByDay,
    dispoByDay: dispoByDay,
    orderSummary: orderSummary,
    sim: {
      remainingSlots: remaining, totalSlots: totalSlots, occupied: occupied, capacity: bedNos.length * nDays,
      avgIncomePerCase: cases ? surgeryIncome / cases : avgCatIncome_(catList)
    }
  };
}

function avgCatIncome_(list) {
  const a = list.filter(x => x.active && x.unit > 0).map(x => x.unit);
  return a.length ? a.reduce((s, x) => s + x, 0) / a.length : 0;
}

/* ------------------------------------------------------------
 * Excel 出力
 * ---------------------------------------------------------- */

/** 発注書（1回の発注分）。includeAll なら注文数0の品目も載せる */
function exportOrderSheet(orderId, includeAll) {
  require_(ROLES_EDIT);
  const lines = orders_().filter(o => o.orderId === orderId && o.status !== '取消');
  if (!lines.length) throw new Error('発注が見つかりません。');
  const order = groupOrders_(lines)[0];
  const vendor = vendors_().find(v => v.name === order.vendor) || { name: order.vendor, addressee: '' };
  const vendorItems = items_().filter(it => it.vendor === vendor.name);
  const model = buildOrderModel_(order, vendor, vendorItems, !!includeAll, str_(conf_()['施設名']));
  return writeXlsx_('発注書_' + safeName_(vendor.name) + '_' + order.date.replace(/-/g, ''), [model]);
}

/** 月次レポート（収益表・コスト計算用・支出合計・発注先ごとの発注集計） */
function exportMonthly(ym) {
  require_(['管理']);
  const md = monthData_(ym);
  const models = [buildRevenueModel_(md), buildCostModel_(md), buildExpenseModel_(md)]
    .concat(md.orderSummary.filter(s => s.rows.length).map(s => buildOrderSummaryModel_(md, s)));
  return writeXlsx_('月次レポート_' + ym.replace('-', ''), models);
}

/* ---- 出力の中身（シートの設計図）を作る：Apps Script に依存しない純粋な関数 ---- */

function SheetModel_(name) {
  this.name = name;
  this.cells = {};
  this.maxR = 0;
  this.maxC = 0;
  this.widths = {};
  this.colFmt = {};
  this.borders = [];
  this.frozen = 0;
}
/** v: 値（'=' で始まる文字列は数式として書き込む）, st: { b: 太字, bg: 背景色, al: 寄せ } */
SheetModel_.prototype.set = function (r, c, v, st) {
  this.cells[r + ',' + c] = { v: v, st: st || null };
  if (r > this.maxR) this.maxR = r;
  if (c > this.maxC) this.maxC = c;
  return this;
};
SheetModel_.prototype.row = function (r, c, vals, st) {
  vals.forEach((v, i) => { if (v !== undefined && v !== null) this.set(r, c + i, v, st); });
  return this;
};
SheetModel_.prototype.box = function (r, c, nr, nc) {
  if (nr > 0 && nc > 0) this.borders.push([r, c, nr, nc]);
  return this;
};

function colName_(c) {
  let s = '';
  while (c > 0) { const m = (c - 1) % 26; s = String.fromCharCode(65 + m) + s; c = Math.floor((c - 1) / 26); }
  return s;
}
function a1_(r, c) { return colName_(c) + r; }

const HEAD_ST = { b: true, bg: '#EEEBE5' };
const SUB_ST = { b: true, bg: '#F6F4EF' };
const TOTAL_ST = { b: true, bg: '#FFF4E8' };

/** 収益表：区分の集計、入院手術・外来手術の内訳（医師別）、入院、術式の単価表 */
function buildRevenueModel_(md) {
  const m = new SheetModel_('収益表');
  const docs = md.doctors;
  const dc = 7;                         // 医師別の列の開始（G列）
  m.set(1, 2, md.year + '年' + md.month + '月　収益表' + (md.facility ? '　' + md.facility : ''), { b: true });
  m.set(2, 2, '請求額合計（術式マスタの請求額から計算）');
  m.set(2, 5, md.charges);

  // 1) 区分の集計（4〜9行目）
  m.row(4, 2, ['区分', '件数', '利益（収入）'], HEAD_ST);
  docs.forEach((d, i) => m.set(4, dc + i, d, HEAD_ST));
  const sumRow = { people: 5, inOps: 6, outOps: 7, total: 8, triple: 9 };
  m.set(sumRow.people, 2, '入院（人）');
  m.set(sumRow.inOps, 2, '入院手術（眼）');
  m.set(sumRow.outOps, 2, '外来手術（眼）');
  m.set(sumRow.total, 2, '合計', TOTAL_ST);
  m.set(sumRow.triple, 2, '手術合計（トリプル分解分）');

  // 2) 入院手術の内訳
  let r = 11;
  const section = (label, t, startRow) => {
    let rr = startRow;
    m.row(rr, 2, [label, '眼数', '1眼あたり収入（円）', '収入（円）'], HEAD_ST);
    docs.forEach((d, i) => m.set(rr, dc + i, d, HEAD_ST));
    const first = rr + 1;
    t.rows.forEach(x => {
      rr++;
      m.set(rr, 2, x.cat);
      m.set(rr, 3, x.eyes);
      m.set(rr, 4, x.unit);
      m.set(rr, 5, '=' + a1_(rr, 3) + '*' + a1_(rr, 4));
      docs.forEach((d, i) => { if (x.byDoc[d]) m.set(rr, dc + i, x.byDoc[d]); });
    });
    const last = rr;
    rr++;
    m.set(rr, 2, '合計', TOTAL_ST);
    if (last >= first) {
      m.set(rr, 3, '=SUM(' + a1_(first, 3) + ':' + a1_(last, 3) + ')', TOTAL_ST);
      m.set(rr, 5, '=SUM(' + a1_(first, 5) + ':' + a1_(last, 5) + ')', TOTAL_ST);
      docs.forEach((d, i) => m.set(rr, dc + i, '=SUM(' + a1_(first, dc + i) + ':' + a1_(last, dc + i) + ')', TOTAL_ST));
    } else {
      m.set(rr, 3, 0, TOTAL_ST);
      m.set(rr, 5, 0, TOTAL_ST);
    }
    m.box(startRow, 2, rr - startRow + 1, 4);
    if (docs.length) m.box(startRow, dc, rr - startRow + 1, docs.length);
    const tripleCells = [];
    t.rows.forEach((x, i) => { if (x.triple) tripleCells.push(a1_(first + i, 3)); });
    return { totalRow: rr, tripleCells: tripleCells };
  };
  const sin = section('入院手術', md.surgery.in, r);
  r = sin.totalRow + 2;
  const sout = section('外来手術', md.surgery.out, r);
  r = sout.totalRow + 2;

  // 3) 入院
  m.row(r, 2, ['入院', '人数', '1人あたり収入（円）', '収入（円）'], HEAD_ST);
  docs.forEach((d, i) => m.set(r, dc + i, d, HEAD_ST));
  const pr = r + 1;
  m.set(pr, 2, '入院');
  m.set(pr, 3, md.inpatient.people);
  m.set(pr, 4, md.inpatient.unit);
  m.set(pr, 5, '=' + a1_(pr, 3) + '*' + a1_(pr, 4));
  docs.forEach((d, i) => { if (md.inpatient.peopleByDoc[d]) m.set(pr, dc + i, md.inpatient.peopleByDoc[d]); });
  m.box(r, 2, 2, 4);
  if (docs.length) m.box(r, dc, 2, docs.length);

  // 区分の集計に参照式を入れる
  m.set(sumRow.people, 3, '=' + a1_(pr, 3));
  m.set(sumRow.people, 4, '=' + a1_(pr, 5));
  m.set(sumRow.inOps, 3, '=' + a1_(sin.totalRow, 3));
  m.set(sumRow.inOps, 4, '=' + a1_(sin.totalRow, 5));
  m.set(sumRow.outOps, 3, '=' + a1_(sout.totalRow, 3));
  m.set(sumRow.outOps, 4, '=' + a1_(sout.totalRow, 5));
  m.set(sumRow.total, 3, '=' + a1_(sumRow.inOps, 3) + '+' + a1_(sumRow.outOps, 3), TOTAL_ST);
  m.set(sumRow.total, 4, '=SUM(' + a1_(sumRow.people, 4) + ':' + a1_(sumRow.outOps, 4) + ')', TOTAL_ST);
  const trip = sin.tripleCells.concat(sout.tripleCells);
  m.set(sumRow.triple, 3, '=' + a1_(sumRow.total, 3) + (trip.length ? '+' + trip.join('+') : ''));
  docs.forEach((d, i) => {
    const col = dc + i;
    m.set(sumRow.people, col, '=' + a1_(pr, col));
    m.set(sumRow.inOps, col, sin.totalRow ? '=' + a1_(sin.totalRow, col) : 0);
    m.set(sumRow.outOps, col, '=' + a1_(sout.totalRow, col));
    m.set(sumRow.total, col, '=' + a1_(sumRow.inOps, col) + '+' + a1_(sumRow.outOps, col), TOTAL_ST);
  });
  m.box(4, 2, 6, 3);
  if (docs.length) m.box(4, dc, 5, docs.length);

  // 4) 術式の単価表（右側）
  const pc = dc + docs.length + 1;
  m.row(4, pc, ['手術名', '合計請求額（円）', 'コスト（円）', '収入（円）'], HEAD_ST);
  let rr = 4;
  md.procRows.forEach(p => {
    rr++;
    m.set(rr, pc, p.name);
    if (p.charge1) m.set(rr, pc + 1, p.charge1);
    if (p.cost1) m.set(rr, pc + 2, p.cost1);
    if (p.charge1) m.set(rr, pc + 3, '=' + a1_(rr, pc + 1) + '-' + a1_(rr, pc + 2));
  });
  m.box(4, pc, rr - 3, 4);
  m.set(rr + 1, pc, '※片眼の金額。収入＝請求額−コスト');

  m.widths = { 1: 16, 2: 190, 3: 70, 4: 130, 5: 120 };
  docs.forEach((d, i) => { m.widths[dc + i] = 80; });
  m.widths[dc - 1] = 16;
  m.widths[pc] = 180; m.widths[pc + 1] = 120; m.widths[pc + 2] = 100; m.widths[pc + 3] = 100;
  m.colFmt = { 3: '#,##0', 4: '#,##0', 5: '#,##0' };
  [pc + 1, pc + 2, pc + 3].forEach(c => { m.colFmt[c] = '#,##0'; });
  return m;
}

/** コスト計算用：術式ごとの件数とコスト合計（手術ディスポ代の内訳） */
function buildCostModel_(md) {
  const m = new SheetModel_('コスト計算用');
  m.set(1, 2, md.year + '年' + md.month + '月　手術コスト計算', { b: true });
  m.row(3, 2, ['手術名', '請求額(片眼)', 'コスト(片眼)', '件数(片眼)', '請求額(両眼)', 'コスト(両眼)', '件数(両眼)', 'コスト合計', '請求額合計'], HEAD_ST);
  let r = 3;
  md.procRows.forEach(p => {
    r++;
    m.row(r, 2, [p.name, p.charge1, p.cost1, p.n1, p.charge2, p.cost2, p.n2]);
    m.set(r, 9, '=' + a1_(r, 4) + '*' + a1_(r, 5) + '+' + a1_(r, 7) + '*' + a1_(r, 8));
    m.set(r, 10, '=' + a1_(r, 3) + '*' + a1_(r, 5) + '+' + a1_(r, 6) + '*' + a1_(r, 8));
  });
  r++;
  m.set(r, 2, '合計', TOTAL_ST);
  [5, 8, 9, 10].forEach(c => m.set(r, c, r > 4 ? '=SUM(' + a1_(4, c) + ':' + a1_(r - 1, c) + ')' : 0, TOTAL_ST));
  m.box(3, 2, r - 2, 9);
  m.widths = { 1: 16, 2: 200 };
  for (let c = 3; c <= 10; c++) { m.widths[c] = 95; m.colFmt[c] = '#,##0'; }
  m.frozen = 3;
  return m;
}

/** 支出合計：日ごと・発注先ごとの金額と、手術ディスポ代 */
function buildExpenseModel_(md) {
  const m = new SheetModel_('支出合計');
  m.set(1, 2, md.year + '年' + md.month + '月　支出合計', { b: true });
  const vs = md.vendors;
  const head = ['日付', '曜日'].concat(vs.map(v => v.name), ['手術ディスポ代', '支出合計', '手術コスト合計']);
  m.row(3, 2, head, HEAD_ST);
  const vc0 = 4;                              // 発注先の最初の列
  const dispoC = vc0 + vs.length;
  const totalC = dispoC + 1, surgC = dispoC + 2;
  let r = 3;
  md.days.forEach(ds => {
    r++;
    const dt = parseYmd_(ds);
    m.set(r, 2, dt);
    m.set(r, 3, WEEKDAYS[dt.getDay()], dt.getDay() === 0 ? { color: '#9A2A1C' } : dt.getDay() === 6 ? { color: '#1F4E79' } : null);
    const sp = md.spendByDay[ds] || {};
    vs.forEach((v, i) => { if (sp[v.name]) m.set(r, vc0 + i, sp[v.name]); });
    if (md.dispoByDay[ds]) m.set(r, dispoC, md.dispoByDay[ds]);
    m.set(r, totalC, '=SUM(' + a1_(r, vc0) + ':' + a1_(r, dispoC) + ')');
    const sc = vs.map((v, i) => v.surgeryCost ? a1_(r, vc0 + i) : null).filter(Boolean).concat([a1_(r, dispoC)]);
    m.set(r, surgC, '=' + sc.join('+'));
  });
  r++;
  m.set(r, 2, '合計', TOTAL_ST);
  m.set(r, 3, '', TOTAL_ST);
  for (let c = vc0; c <= surgC; c++) m.set(r, c, '=SUM(' + a1_(4, c) + ':' + a1_(r - 1, c) + ')', TOTAL_ST);
  m.box(3, 2, r - 2, surgC - 1);
  m.widths = { 1: 16, 2: 70, 3: 45 };
  for (let c = vc0; c <= surgC; c++) { m.widths[c] = 115; m.colFmt[c] = '#,##0'; }
  m.colFmt[2] = 'm/d';
  m.frozen = 3;
  m.set(r + 2, 2, '※手術ディスポ代は手術の登録と術式マスタのコストから自動計算。手術コスト合計＝「手術コストに含める」発注先＋手術ディスポ代');
  return m;
}

/** 発注集計：品目×週（1週目＝1〜7日、2週目＝8〜14日、…5週目＝29日〜末日）の金額 */
function buildOrderSummaryModel_(md, s) {
  const m = new SheetModel_(safeName_('発注集計_' + s.vendor).slice(0, 90));
  m.set(1, 2, s.vendor + '　' + md.year + '年' + md.month + '月　発注集計（金額）', { b: true });
  const weekHeads = [1, 2, 3, 4, 5].map(w => md.month + '月 ' + w + '週目');
  m.row(3, 1, ['後発', '商品コード', '品名', 'メーカー', '区分', '規格', '入数', '薬価', '納入価'].concat(weekHeads, ['合計', '数量合計']), HEAD_ST);
  const wc0 = 10, sumC = 15, qtyC = 16;
  const groups = [];
  s.rows.forEach(x => { if (groups.indexOf(x.group) < 0) groups.push(x.group); });
  let r = 3;
  const subtotalRows = [];
  groups.forEach(g => {
    const first = r + 1;
    s.rows.filter(x => x.group === g).forEach(x => {
      r++;
      m.row(r, 1, [x.generic ? '後発' : '', x.code, x.name, x.maker, x.group, x.spec, x.pack, x.listPrice, x.price]);
      x.weeks.forEach((a, i) => { if (a) m.set(r, wc0 + i, a); });
      m.set(r, sumC, '=SUM(' + a1_(r, wc0) + ':' + a1_(r, wc0 + 4) + ')');
      if (x.qty) m.set(r, qtyC, x.qty);
    });
    r++;
    m.set(r, 3, g + '　小計', SUB_ST);
    for (let c = wc0; c <= qtyC; c++) m.set(r, c, '=SUM(' + a1_(first, c) + ':' + a1_(r - 1, c) + ')', SUB_ST);
    subtotalRows.push(r);
  });
  r++;
  m.set(r, 3, '合計', TOTAL_ST);
  for (let c = wc0; c <= qtyC; c++) {
    m.set(r, c, subtotalRows.length ? '=' + subtotalRows.map(x => a1_(x, c)).join('+') : 0, TOTAL_ST);
  }
  m.box(3, 1, r - 2, qtyC);
  m.widths = { 1: 45, 2: 85, 3: 300, 4: 140, 5: 110, 6: 110, 7: 110, 8: 80, 9: 80, 15: 100, 16: 70 };
  for (let c = wc0; c < wc0 + 5; c++) m.widths[c] = 95;
  [8, 9, 10, 11, 12, 13, 14, 15, 16].forEach(c => { m.colFmt[c] = '#,##0'; });
  m.frozen = 3;
  return m;
}

/**
 * 発注書（FAX・メール用）
 * vendorItems: その発注先の品目マスタ（メーカー名の補完と、includeAll のときの全品目一覧に使う）
 */
function buildOrderModel_(order, vendor, vendorItems, includeAll, facility) {
  const m = new SheetModel_('発注書');
  const dt = parseYmd_(order.date);
  const dateText = dt.getFullYear() + '年' + (dt.getMonth() + 1) + '月' + dt.getDate() + '日';
  m.set(1, 2, (vendor.addressee || vendor.name) + '　御中', { b: true });
  m.set(2, 2, (facility ? facility + '　' : '') + '発注書', { b: true });
  m.set(2, 6, '発注日　' + dateText);
  if (order.memo) m.set(3, 2, 'メモ：' + order.memo);
  m.row(5, 2, ['メーカー', '品名', '規格', '入数', '納入価', '注文数', '合計'], HEAD_ST);

  const byId = {};
  vendorItems.forEach(it => { byId[it.id] = it; });
  const ordered = {};
  order.lines.forEach(l => { if (l.itemId) ordered[l.itemId] = l; });
  const lines = [];
  if (includeAll) {
    vendorItems.filter(it => it.active || ordered[it.id]).forEach(it => {
      const l = ordered[it.id];
      lines.push({ group: it.group, sub: it.sub, maker: it.maker, name: it.name, spec: it.spec, pack: it.pack,
        price: l ? l.price : it.price, qty: l ? l.qty : '' });
    });
    order.lines.filter(l => !l.itemId || !byId[l.itemId]).forEach(l => {
      lines.push({ group: l.group, sub: l.sub, maker: '', name: l.name, spec: l.spec, pack: l.pack, price: l.price, qty: l.qty });
    });
  } else {
    order.lines.forEach(l => {
      const it = byId[l.itemId];
      lines.push({ group: l.group, sub: l.sub, maker: it ? it.maker : '', name: l.name, spec: l.spec, pack: l.pack, price: l.price, qty: l.qty });
    });
  }

  let r = 5;
  let lastHead = null;
  const firstData = r + 1;
  lines.forEach(l => {
    const head = [l.group, l.sub].filter(Boolean).join('　');
    if (head && head !== lastHead) {
      r++;
      m.set(r, 2, head, SUB_ST);
      for (let c = 3; c <= 8; c++) m.set(r, c, '', SUB_ST);
      lastHead = head;
    }
    r++;
    m.row(r, 2, [l.maker, l.name, l.spec, l.pack, l.price, l.qty]);
    m.set(r, 8, '=' + a1_(r, 6) + '*' + a1_(r, 7));
  });
  r++;
  m.set(r, 2, '合計', TOTAL_ST);
  for (let c = 3; c <= 6; c++) m.set(r, c, '', TOTAL_ST);
  m.set(r, 7, '=SUM(' + a1_(firstData, 7) + ':' + a1_(r - 1, 7) + ')', TOTAL_ST);
  m.set(r, 8, '=SUM(' + a1_(firstData, 8) + ':' + a1_(r - 1, 8) + ')', TOTAL_ST);
  m.box(5, 2, r - 4, 7);
  m.widths = { 1: 16, 2: 130, 3: 300, 4: 120, 5: 110, 6: 85, 7: 65, 8: 100 };
  m.colFmt = { 6: '#,##0', 7: '#,##0', 8: '#,##0' };
  m.frozen = 5;
  return m;
}

/* ---- 設計図をスプレッドシートに書き、xlsx にして返す ---- */
function writeXlsx_(baseName, models) {
  const ss = SpreadsheetApp.create(baseName);
  try {
    const firstSheet = ss.getSheets()[0];
    models.forEach((m, i) => {
      const sh = i === 0 ? firstSheet.setName(m.name) : ss.insertSheet(m.name);
      renderModel_(sh, m);
    });
    SpreadsheetApp.flush();
    const url = 'https://docs.google.com/spreadsheets/d/' + ss.getId() + '/export?format=xlsx';
    const res = UrlFetchApp.fetch(url, { headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() }, muteHttpExceptions: true });
    if (res.getResponseCode() !== 200) throw new Error('Excelファイルの作成に失敗しました（' + res.getResponseCode() + '）。');
    return { name: baseName + '.xlsx', b64: Utilities.base64Encode(res.getContent()) };
  } finally {
    try { DriveApp.getFileById(ss.getId()).setTrashed(true); } catch (e) { /* 一時ファイルの削除に失敗しても出力は返す */ }
  }
}

function renderModel_(sh, m) {
  const R = Math.max(1, m.maxR), C = Math.max(1, m.maxC);
  if (sh.getMaxRows() < R) sh.insertRowsAfter(sh.getMaxRows(), R - sh.getMaxRows());
  if (sh.getMaxColumns() < C) sh.insertColumnsAfter(sh.getMaxColumns(), C - sh.getMaxColumns());
  const grid = v => Array.from({ length: R }, () => Array.from({ length: C }, () => v));
  const vals = grid(''), weights = grid('normal'), bgs = grid('#ffffff'), colors = grid('#000000');
  Object.keys(m.cells).forEach(k => {
    const [r, c] = k.split(',').map(Number);
    const cell = m.cells[k];
    let v = cell.v;
    if (typeof v === 'string' && v.charAt(0) !== '=' && /^[+\-@]/.test(v)) v = "'" + v;
    vals[r - 1][c - 1] = v;
    if (cell.st) {
      if (cell.st.b) weights[r - 1][c - 1] = 'bold';
      if (cell.st.bg) bgs[r - 1][c - 1] = cell.st.bg;
      if (cell.st.color) colors[r - 1][c - 1] = cell.st.color;
    }
  });
  const range = sh.getRange(1, 1, R, C);
  range.setValues(vals).setFontWeights(weights).setBackgrounds(bgs).setFontColors(colors);
  Object.keys(m.colFmt).forEach(c => sh.getRange(1, Number(c), R, 1).setNumberFormat(m.colFmt[c]));
  m.borders.forEach(b => sh.getRange(b[0], b[1], b[2], b[3])
    .setBorder(true, true, true, true, true, true, '#9A958C', SpreadsheetApp.BorderStyle.SOLID));
  Object.keys(m.widths).forEach(c => sh.setColumnWidth(Number(c), m.widths[c]));
  if (m.frozen) sh.setFrozenRows(m.frozen);
}

function safeName_(s) {
  return String(s).replace(/[\\\/\?\*\[\]:'"<>|]/g, '_').slice(0, 80) || 'sheet';
}

/** 見出しが今の版と違うシートがあれば、分かりやすいエラーにする（古い版のスプレッドシートを使った場合など） */
function checkHeaders_() {
  ['BOOK', 'PROC', 'CATG', 'VEND', 'ITEM', 'ORDER'].forEach(k => {
    const sh = sheet_(SH[k]);
    const head = sh.getRange(1, 1, 1, HEAD[k].length).getValues()[0].map(str_);
    const bad = HEAD[k].findIndex((h, i) => head[i] !== h);
    if (bad >= 0) {
      throw new Error('シート「' + SH[k] + '」の' + (bad + 1) + '列目の見出しが「' + (head[bad] || '空欄') + '」になっています（正しくは「' +
        HEAD[k][bad] + '」）。1行目の見出しは変えずに、2行目以降にデータを入れてください。');
    }
  });
}

/* ------------------------------------------------------------
 * 内部関数：読み取り
 * ---------------------------------------------------------- */
function currentUser_() {
  const email = (Session.getActiveUser().getEmail() || '').toLowerCase();
  if (!email) return null;
  const r = rows_(SH.USER).find(x => str_(x[0]).toLowerCase() === email);
  if (!r) return null;
  const role = str_(r[2]) || '閲覧';
  return { email: email, name: str_(r[1]) || email, role: ROLES_ALL.indexOf(role) >= 0 ? role : '閲覧' };
}

function require_(roles) {
  const u = currentUser_();
  if (!u || roles.indexOf(u.role) < 0) throw new Error('この操作の権限がありません。');
  return u;
}

function sheet_(name) {
  const sh = SpreadsheetApp.getActive().getSheetByName(name);
  if (!sh) throw new Error('シート「' + name + '」がありません。ページを再読み込みしてください。');
  return sh;
}

function rows_(name) {
  return sheet_(name).getDataRange().getValues().slice(1);
}

function conf_() {
  const o = {};
  rows_(SH.CONF).forEach(r => { if (str_(r[0])) o[str_(r[0])] = r[1]; });
  return o;
}

function str_(v) {
  if (v === null || v === undefined) return '';
  if (v instanceof Date) return ymd_(v);
  return String(v).trim();
}
function num_(v) {
  if (typeof v === 'number') return isFinite(v) ? v : 0;
  const n = Number(String(v === null || v === undefined ? '' : v).replace(/[,¥円\s]/g, ''));
  return isFinite(n) ? n : 0;
}
function truthy_(v) { return v === true || String(v).toUpperCase() === 'TRUE'; }

function procs_() {
  return rows_(SH.PROC).filter(r => str_(r[0])).map(r => {
    const c1 = num_(r[4]), k1 = num_(r[6]);
    return {
      code: str_(r[0]), name: str_(r[1]) || str_(r[0]), cat: str_(r[2]),
      nights: Math.max(0, Math.round(num_(r[3]))),
      charge1: c1, charge2: str_(r[5]) === '' ? c1 * 2 : num_(r[5]),
      cost1: k1, cost2: str_(r[7]) === '' ? k1 * 2 : num_(r[7]),
      color: str_(r[8]), active: truthy_(r[9])
    };
  });
}

function procMap_() {
  const m = {};
  procs_().forEach(p => { m[p.code] = p; });
  return m;
}

function cats_() {
  return rows_(SH.CATG).filter(r => str_(r[0])).map(r => ({
    name: str_(r[0]), unit: num_(r[1]), triple: truthy_(r[2]), showIn: truthy_(r[3]), showOut: truthy_(r[4]), active: truthy_(r[5])
  }));
}

function doctors_() {
  return rows_(SH.DOC).filter(r => str_(r[0]) && truthy_(r[1])).map(r => ({ name: str_(r[0]) }));
}

function beds_() {
  return rows_(SH.BED).filter(r => truthy_(r[2]) && Number(r[0]))
    .map(r => ({ no: Number(r[0]), name: str_(r[1]) || ('ベッド' + r[0]) }))
    .sort((a, b) => a.no - b.no);
}

function slots_() {
  const o = {};
  rows_(SH.SLOT).forEach(r => { o[str_(r[0])] = num_(r[1]); });
  return o;
}

function vendors_() {
  return rows_(SH.VEND).filter(r => str_(r[0])).map(r => ({
    name: str_(r[0]), addressee: str_(r[1]), mode: str_(r[2]) === '金額' ? '金額' : '品目', surgeryCost: truthy_(r[3]), active: truthy_(r[4])
  }));
}

function items_() {
  return rows_(SH.ITEM).filter(r => str_(r[0]) && str_(r[5])).map(r => ({
    id: str_(r[0]), vendor: str_(r[1]), group: str_(r[2]), sub: str_(r[3]), code: str_(r[4]), name: str_(r[5]),
    maker: str_(r[6]), spec: str_(r[7]), pack: str_(r[8]), listPrice: str_(r[9]) === '' ? '' : num_(r[9]), price: num_(r[10]),
    generic: truthy_(r[11]), active: truthy_(r[12])
  }));
}

/** 品目IDが空の行に ID を振る（品名が入っている行のみ） */
function ensureItemIds_() {
  const sh = sheet_(SH.ITEM);
  const vals = sh.getDataRange().getValues();
  const need = [];
  let max = 0;
  for (let i = 1; i < vals.length; i++) {
    const id = str_(vals[i][0]);
    const m = /^I(\d+)$/.exec(id);
    if (m) max = Math.max(max, Number(m[1]));
    if (!id && str_(vals[i][5])) need.push(i + 1);
  }
  if (!need.length) return;
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    need.forEach(row => {
      if (str_(sh.getRange(row, 1).getValue())) return;
      max++;
      sh.getRange(row, 1).setValue('I' + ('0000' + max).slice(-4));
    });
  } finally {
    lock.releaseLock();
  }
}

function bookings_() {
  const v = sheet_(SH.BOOK).getDataRange().getValues();
  const out = [];
  for (let i = 1; i < v.length; i++) {
    const r = v[i];
    if (!str_(r[B.id])) continue;
    out.push({
      id: str_(r[B.id]), patientId: str_(r[B.pid]), proc: str_(r[B.proc]), kind: str_(r[B.kind]) === '外来' ? '外来' : '入院',
      opDate: ymd_(r[B.op]), eye: EYES.indexOf(str_(r[B.eye])) >= 0 ? str_(r[B.eye]) : '右', doctor: str_(r[B.doc]),
      admit: ymd_(r[B.admit]), nights: Math.max(0, Math.round(num_(r[B.nights]))), bed: Number(r[B.bed]) || 0,
      status: str_(r[B.status]) || '仮', memo: str_(r[B.memo]), eventId: str_(r[B.ev]), row: i + 1
    });
  }
  return out;
}

function orders_() {
  const v = sheet_(SH.ORDER).getDataRange().getValues();
  const out = [];
  for (let i = 1; i < v.length; i++) {
    const r = v[i];
    if (!str_(r[O.id])) continue;
    out.push({
      orderId: str_(r[O.id]), date: ymd_(r[O.date]), vendor: str_(r[O.vendor]), itemId: str_(r[O.item]), name: str_(r[O.name]),
      group: str_(r[O.group]), sub: str_(r[O.sub]), spec: str_(r[O.spec]), pack: str_(r[O.pack]), price: num_(r[O.price]),
      qty: num_(r[O.qty]), amount: num_(r[O.amount]), memo: str_(r[O.memo]), status: str_(r[O.status]) || '有効',
      cBy: str_(r[O.cBy]), row: i + 1
    });
  }
  return out;
}

/* ---- 日付・その他 ---- */
function eyes_(eye) { return eye === '両' ? 2 : 1; }
/** 入院の最終日（退院日）。退院日も同じベッドを使う扱い */
function endOf_(b) { return b.kind === '入院' ? addDays_(b.admit, b.nights) : b.opDate; }
function overlap_(a1, a2, b1, b2) { return a1 <= b2 && b1 <= a2; }
function isYmd_(s) { return /^\d{4}-\d{2}-\d{2}$/.test(s || ''); }
function ymd_(d) {
  if (d instanceof Date) return Utilities.formatDate(d, TZ, 'yyyy-MM-dd');
  return d ? String(d).slice(0, 10) : '';
}
function parseYmd_(s) { const a = s.split('-').map(Number); return new Date(a[0], a[1] - 1, a[2]); }
function addDays_(s, n) { const d = parseYmd_(s); d.setDate(d.getDate() + n); return ymd_(d); }
function newId_(prefix) {
  return prefix + Utilities.formatDate(new Date(), TZ, 'yyMMddHHmmss') + Math.floor(Math.random() * 900 + 100);
}

function describeBooking_(v, bed) {
  return v.patientId + ' ' + v.kind + ' ' + v.proc.name + '（' + v.eye + '）手術' + v.opDate +
    (v.kind === '入院' ? ' 入院' + v.admit + ' ' + v.nights + '泊 ベッド' + bed : '') + ' ' + (v.doctor || '') + ' ' + v.status;
}
function calTitle_(v, bed) {
  return (v.kind === '入院' ? 'ベッド' + bed + ' ' : '外来手術 ') + v.patientId + ' ' + v.proc.name + '（' + v.eye + '）' +
    (v.doctor ? ' ' + v.doctor : '') + (v.status === '仮' ? '［仮］' : '');
}

/** 複数行をまとめて末尾に追加（行が足りなければ増やす） */
function appendRows_(sh, rows) {
  const start = sh.getLastRow() + 1;
  const need = start + rows.length - 1 - sh.getMaxRows();
  if (need > 0) sh.insertRowsAfter(sh.getMaxRows(), need);
  sh.getRange(start, 1, rows.length, rows[0].length).setValues(rows);
}

function log_(u, action, id, detail) {
  sheet_(SH.LOG).appendRow([new Date(), u.email, action, id, detail]);
}

/** 設定シートにカレンダーIDがあれば終日予定として反映 */
function syncCalendar_(row, eventId, title, start, end, remove) {
  const calId = str_(conf_()['カレンダーID']);
  if (!calId) return;
  try {
    const cal = CalendarApp.getCalendarById(calId);
    if (!cal) throw new Error('カレンダーが見つかりません: ' + calId);
    const sh = sheet_(SH.BOOK);
    let ev = eventId ? cal.getEventById(eventId) : null;
    if (remove) {
      if (ev) ev.deleteEvent();
      sh.getRange(row, B.ev + 1).setValue('');
      return;
    }
    const s = parseYmd_(start), e = parseYmd_(addDays_(end, 1));
    if (ev) {
      ev.setTitle(title);
      ev.setAllDayDates(s, e);
    } else {
      ev = cal.createAllDayEvent(title, s, e);
      sh.getRange(row, B.ev + 1).setValue(ev.getId());
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
:root{--ink:#1F2328;--ground:#EEEBE5;--panel:#FFFFFF;--line:#DDD7CC;--soft:#F6F4EF;--muted:#4A4F57;--accent:#B45309;--accent-bg:#FFF4E8;--sel:#1F4E79;--bad:#B42318;--good:#1D5A2C}
*{box-sizing:border-box}
body{margin:0;background:var(--ground);color:var(--ink);font-family:'Zen Kaku Gothic New','Hiragino Kaku Gothic ProN','Yu Gothic','Meiryo',sans-serif;font-variant-numeric:tabular-nums;-webkit-text-size-adjust:100%}
button,input,select,textarea{font-family:inherit;color:inherit}
:focus-visible{outline:3px solid var(--sel);outline-offset:2px}
[hidden]{display:none !important}
header{display:flex;align-items:center;gap:16px;padding:10px 20px;background:var(--ink);color:#F4F1EA;flex-wrap:wrap;position:sticky;top:0;z-index:20}
.brand{display:flex;align-items:baseline;gap:10px}
header h1{margin:0;font-size:19px;letter-spacing:.04em}
.brand span{color:#C9C4BA;font-size:13px}
nav{display:flex;gap:6px;flex-wrap:wrap}
.tab{min-height:44px;padding:0 16px;border-radius:8px;border:1px solid #5A5F66;background:transparent;color:#F4F1EA;font-size:15px;cursor:pointer}
.tab[aria-pressed="true"]{background:#F4F1EA;color:var(--ink);border-color:#F4F1EA;font-weight:700}
#who{margin-left:auto;font-size:13px;color:#C9C4BA}
main{padding:16px 20px 40px}
.panel{background:var(--panel);border:1px solid var(--line);border-radius:12px;padding:16px 18px}
h2{margin:0 0 12px;font-size:17px}
h3{margin:14px 0 8px;font-size:14px;color:var(--muted)}
.two{display:grid;grid-template-columns:400px minmax(0,1fr);gap:16px;align-items:start}
@media (max-width:1050px){.two{grid-template-columns:minmax(0,1fr)}}
.orderGrid{display:grid;grid-template-columns:minmax(0,1fr) 380px;grid-template-areas:"items cart" "hist cart";gap:16px;align-items:start}
.orderGrid>.a-items{grid-area:items;min-width:0}.orderGrid>.a-cart{grid-area:cart}.orderGrid>.a-hist{grid-area:hist;min-width:0}
@media (max-width:1050px){.orderGrid{grid-template-columns:minmax(0,1fr);grid-template-areas:"items" "cart" "hist"}}
#cartBar{display:none}
@media (max-width:1050px){#vorder{padding-bottom:72px}}
@media (max-width:1050px){#cartBar.show{display:flex;position:fixed;left:0;right:0;bottom:0;z-index:30;align-items:center;gap:12px;padding:10px 16px;background:var(--ink);color:#F4F1EA;font-size:16px}#cartBar .grow{flex:1}#cartBar button{min-height:44px;border-radius:8px;border:none;background:var(--accent);color:#fff;font-weight:700;padding:0 16px;font-size:15px}}
.stack{display:flex;flex-direction:column;gap:16px;min-width:0}
.field{margin-bottom:12px}
.label{display:block;font-size:13px;color:var(--muted);font-weight:500;margin-bottom:6px}
input[type=text],input[type=date],input[type=month],input[type=number],input[type=search],select{height:44px;border:1px solid #CFC8BC;border-radius:8px;padding:0 10px;font-size:16px;background:#FBFAF7;width:100%}
input[type=number]{text-align:right}
.chips{display:flex;flex-wrap:wrap;gap:6px}
.chip{min-height:44px;padding:6px 12px;border-radius:8px;border:1px solid #CFC8BC;background:#fff;font-size:15px;cursor:pointer;text-align:left;line-height:1.25}
.chip small{display:block;font-size:11px;color:var(--muted)}
.chip[aria-pressed="true"]{background:var(--sel);border-color:var(--sel);color:#fff}
.chip[aria-pressed="true"] small{color:#D8E2EC}
.chips.grid3{display:grid;grid-template-columns:repeat(3,minmax(0,1fr))}
.chips.grid3 .chip{font-size:14px;padding:6px 8px}
.chips.seg .chip{flex:1 1 0;text-align:center}
.grouphead{font-size:12px;color:var(--muted);margin:10px 0 4px;display:flex;align-items:center;gap:6px}
.grouphead i{display:inline-block;width:10px;height:10px;border-radius:3px}
.preview{border-radius:10px;padding:10px 12px;font-size:14px;margin:4px 0 12px;line-height:1.6}
.preview:empty{display:none}
.preview.ok{border:2px dashed var(--accent);background:var(--accent-bg);color:#5C2D05}
.preview.ng{border:2px solid var(--bad);background:#FDECEA;color:#7A1A12}
.preview.info{border:1px solid var(--line);background:var(--soft)}
.preview strong{font-size:17px}
.actions{display:flex;flex-direction:column;gap:8px}
.primary{min-height:50px;border:none;border-radius:10px;background:var(--accent);color:#fff;font-size:16px;font-weight:700;cursor:pointer;padding:0 18px}
.primary:disabled{background:#E3DED5;color:#6B6F75;cursor:not-allowed}
.danger{min-height:44px;border-radius:8px;border:1px solid var(--bad);background:#fff;color:var(--bad);font-size:14px;cursor:pointer;padding:0 14px}
.ghost{min-height:44px;padding:0 14px;border-radius:8px;border:1px solid #CFC8BC;background:#fff;font-size:14px;cursor:pointer}
.ghost.sm{min-height:36px;font-size:13px;padding:0 10px}
.msg{min-height:1.4em;font-size:13px;margin:10px 0 0;line-height:1.5}
.msg.err{color:var(--bad)}
.msg.ok{color:var(--good)}
.toolbar{display:flex;align-items:center;gap:8px;flex-wrap:wrap;margin-bottom:12px}
.toolbar .grow{flex:1}
#range{font-weight:700;margin:0 6px}
.legend{display:flex;gap:10px;font-size:12px;color:var(--muted);flex-wrap:wrap}
.legend i{display:inline-block;width:12px;height:12px;border-radius:3px;margin-right:4px;vertical-align:-1px;border:1px solid}
.scroll{overflow-x:auto}
#grid{min-width:680px}
.opwrap{max-height:540px;overflow:auto}
.g-row{display:grid;grid-template-columns:64px minmax(0,1fr);border-bottom:1px solid #EEE9E0}
.g-bed{font-weight:700;font-size:13px;display:flex;align-items:center}
.g-track{display:grid}
.g-day{text-align:center;padding:5px 0;font-size:13px;border-left:1px solid #EEE9E0}
.g-day small{display:block;font-size:11px}
.g-day.sat{color:#1F4E79}
.g-day.sun{color:#9A2A1C}
.g-day.today{background:var(--accent-bg);font-weight:700}
.g-cell{grid-row:1;height:46px;border:none;border-left:1px solid #EEE9E0;background:transparent;padding:0}
.g-cell.wk{background:#F7F5F0}
button.g-cell{cursor:pointer}
button.g-cell:hover{background:var(--accent-bg)}
.g-bar{grid-row:1;z-index:1;align-self:center;height:34px;margin:0 3px;border-radius:6px;border:1px solid transparent;font-size:12px;font-weight:700;text-align:left;padding:0 7px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;cursor:pointer}
.g-bar.tent{border-style:dashed;background-image:repeating-linear-gradient(135deg,rgba(255,255,255,0) 0 6px,rgba(255,255,255,.6) 6px 12px)}
.g-bar.cutL{border-top-left-radius:0;border-bottom-left-radius:0;margin-left:0}
.g-bar.cutR{border-top-right-radius:0;border-bottom-right-radius:0;margin-right:0}
.g-bar.pending{background:var(--accent-bg);border:2px dashed var(--accent);color:#7A3A06;cursor:default;z-index:2}
.g-bar.editing{outline:3px solid var(--sel);outline-offset:1px}
.g-usage{display:grid;font-size:12px;color:var(--muted);padding-top:5px;text-align:center}
.g-usage .full{color:#9A3412;font-weight:700}
table.list{width:100%;border-collapse:collapse;font-size:14px}
table.list th{font-size:12px;color:var(--muted);font-weight:500;text-align:left;padding:6px 8px;border-bottom:1px solid var(--line);white-space:nowrap}
table.list td{padding:8px;border-bottom:1px solid #EEE9E0;vertical-align:middle}
table.list td.num,table.list th.num{text-align:right}
table.list tr.clickable{cursor:pointer}
table.list tr.clickable:hover{background:var(--soft)}
table.list tr.total td{font-weight:700;background:var(--accent-bg)}
.badge{display:inline-block;font-size:11px;font-weight:700;padding:2px 7px;border-radius:5px;border:1px solid;white-space:nowrap}
.badge.in{color:#123A5A;border-color:#4F7FA8;background:#EAF1F8}
.badge.out{color:#5C2D05;border-color:#C47A3A;background:#FBEFE3}
.badge.tent{color:#4A4F57;border-style:dashed;border-color:#8C9198}
.empty{color:var(--muted);font-size:14px;padding:12px 0}
/* 発注 */
.items{display:flex;flex-direction:column}
.item{display:grid;grid-template-columns:minmax(0,1fr) 92px 148px;gap:10px;align-items:center;padding:8px 2px;border-bottom:1px solid #EEE9E0}
.item.on{background:#FFFBF5}
.item .nm{font-size:15px;line-height:1.35;word-break:break-all}
.item .sp{font-size:12px;color:var(--muted)}
.item .pr{text-align:right;font-size:14px}
.stepper{display:grid;grid-template-columns:44px minmax(0,1fr) 44px;gap:4px}
.stepper button{height:44px;border-radius:8px;border:1px solid #CFC8BC;background:#fff;font-size:20px;cursor:pointer;line-height:1}
.stepper input{height:44px;padding:0 6px;text-align:center}
.subhead{font-size:13px;font-weight:700;color:var(--sel);padding:12px 2px 4px;border-bottom:2px solid #D6E4F0}
.cart{position:sticky;top:78px}
.cart .line{display:grid;grid-template-columns:minmax(0,1fr) 54px 92px;gap:6px;align-items:center;padding:6px 0;border-bottom:1px solid #EEE9E0;font-size:14px}
.cart .line input{height:36px;font-size:14px}
.cart .sum{display:flex;justify-content:space-between;font-size:18px;font-weight:700;padding:10px 0}
@media (max-width:700px){.item{grid-template-columns:minmax(0,1fr) 140px}.item .pr{display:none}}
/* 収益 */
.kpis{display:grid;grid-template-columns:repeat(3,minmax(0,1fr));gap:12px;margin-bottom:16px}
@media (max-width:900px){.kpis{grid-template-columns:repeat(2,minmax(0,1fr))}}
.kpi{background:var(--panel);border:1px solid var(--line);border-radius:12px;padding:12px 14px}
.kpi .l{font-size:13px;color:var(--muted);font-weight:500}
.kpi .v{font-size:24px;font-weight:700;margin:4px 0}
.kpi .s{font-size:12px;color:var(--muted)}
.kpi.main{background:var(--ink);color:#F4F1EA;border-color:var(--ink)}
.kpi.main .l,.kpi.main .s{color:#C9C4BA}
.warn{background:var(--accent-bg);border:1px solid var(--accent);color:#5C2D05;border-radius:10px;padding:10px 14px;font-size:14px;margin-bottom:16px;line-height:1.6}
.sim{display:grid;grid-template-columns:repeat(3,minmax(0,1fr));gap:16px;align-items:end}
@media (max-width:700px){.sim{grid-template-columns:minmax(0,1fr)}}
input[type=range]{width:100%;accent-color:var(--sel);height:32px}
.note{font-size:13px;color:var(--muted);line-height:1.7;margin:10px 0 0}
#loading{padding:40px;text-align:center;color:var(--muted)}
.dl{margin-top:10px;font-size:14px}
.dl a{color:var(--sel)}
</style>
</head>
<body>
<header>
  <div class="brand"><h1>入院・手術・発注管理</h1><span id="facility"></span></div>
  <nav aria-label="画面切替">
    <button type="button" class="tab" data-view="book" aria-pressed="true">入院・手術</button>
    <button type="button" class="tab" data-view="order" aria-pressed="false" hidden>発注</button>
    <button type="button" class="tab" data-view="rev" aria-pressed="false" hidden>収益・支出</button>
    <button type="button" class="tab" data-view="out" aria-pressed="false" hidden>出力</button>
  </nav>
  <div id="who"></div>
</header>

<main>
  <div id="loading">読み込み中…</div>

  <!-- 入院・手術 -->
  <section id="vbook" hidden>
    <div class="two" id="bookLayout">
      <aside id="formPanel" class="panel">
        <h2 id="formTitle">手術・入院の登録</h2>
        <div class="field"><label class="label" for="fPid">患者ID</label><input type="text" id="fPid" autocomplete="off" placeholder="例：0012345"></div>
        <div class="field"><span class="label">区分</span><div class="chips seg" id="fKind"></div></div>
        <div class="field"><label class="label" for="fOp">手術日</label><input type="date" id="fOp"></div>
        <div class="field"><span class="label">術式</span><div id="fProc"></div></div>
        <div class="field"><span class="label">眼</span><div class="chips seg" id="fEye"></div></div>
        <div class="field"><span class="label">執刀医</span><div class="chips" id="fDoc"></div></div>
        <div id="inOnly">
          <div class="field"><span class="label">入院日</span>
            <div class="chips seg" id="fAdmitMode"></div>
            <input type="date" id="fAdmit" style="margin-top:6px" hidden>
          </div>
          <div class="field"><span class="label">泊数</span><div class="chips" id="fNights"></div></div>
          <div class="field"><label class="label" for="fBed">ベッド</label><select id="fBed"></select></div>
        </div>
        <div class="field"><span class="label">状態</span><div class="chips seg" id="fStatus"></div></div>
        <div class="field"><label class="label" for="fMemo">メモ</label><input type="text" id="fMemo" maxlength="200" autocomplete="off"></div>
        <div id="preview" class="preview"></div>
        <div class="actions">
          <button type="button" id="btnSave" class="primary">登録する</button>
          <button type="button" id="btnCancelBk" class="danger" hidden>この記録を取り消す</button>
          <button type="button" id="btnNew" class="ghost" hidden>新規入力に戻る</button>
        </div>
        <p id="msgBook" class="msg" role="status"></p>
      </aside>

      <div class="stack">
        <div class="panel">
          <div class="toolbar">
            <button type="button" class="ghost" id="prev">前の週</button>
            <button type="button" class="ghost" id="thisWeek">今週</button>
            <button type="button" class="ghost" id="next">次の週</button>
            <span id="range"></span>
            <span class="grow"></span>
            <span class="legend" id="legend"></span>
          </div>
          <div class="scroll"><div id="grid"></div></div>
        </div>
        <div class="panel">
          <h2>この2週間の手術</h2>
          <div class="scroll opwrap"><table class="list" id="opList"></table></div>
        </div>
      </div>
    </div>
  </section>

  <!-- 発注 -->
  <section id="vorder" hidden>
    <div class="orderGrid">
      <div class="panel a-items">
        <div class="toolbar">
          <div style="width:180px"><label class="label" for="oDate">発注日</label><input type="date" id="oDate"></div>
          <div class="grow"><span class="label">発注先</span><div class="chips" id="oVendor"></div></div>
        </div>
        <div id="oItemsBox">
          <div class="toolbar">
            <div class="chips" id="oGroup"></div>
          </div>
          <div class="toolbar">
            <div class="grow"><input type="search" id="oSearch" placeholder="品名・メーカー・規格で絞り込み" aria-label="品目を絞り込み"></div>
            <button type="button" class="chip" id="oOnly" aria-pressed="false">数量を入れた品目だけ</button>
          </div>
          <div class="items" id="oItems"></div>
        </div>
        <div id="oAmountBox" hidden>
          <div class="field"><label class="label" for="oAmount">金額（円）</label><input type="number" id="oAmount" inputmode="numeric" min="0" step="1"></div>
          <p class="note">品目を登録しない発注先です（例：レンズ）。納品書の金額をそのまま入れてください。</p>
        </div>
      </div>
      <aside class="panel cart a-cart" id="cartPanel">
        <h2 id="cartTitle">発注内容</h2>
        <div id="cartLines"></div>
        <div class="sum"><span>合計</span><span id="cartTotal">¥0</span></div>
        <div class="field"><label class="label" for="oMemo">メモ</label><input type="text" id="oMemo" maxlength="200" autocomplete="off"></div>
        <div class="actions">
          <button type="button" class="primary" id="oSave">この内容で発注を登録</button>
          <button type="button" class="ghost" id="oClear">入力をクリア</button>
        </div>
        <p id="msgOrder" class="msg" role="status"></p>
      </aside>
      <div class="panel a-hist">
        <div class="toolbar"><h2 style="margin:0">発注履歴</h2><span class="grow"></span><input type="month" id="oMonth" style="width:180px" aria-label="表示する月"></div>
        <div class="scroll"><table class="list" id="oHist"></table></div>
        <div id="oDl" class="dl"></div>
      </div>
    </div>
    <div id="cartBar"><span id="cartBarText"></span><span class="grow"></span><button type="button" id="cartBarBtn">発注内容へ</button></div>
  </section>

  <!-- 収益・支出 -->
  <section id="vrev" hidden>
    <div class="panel" style="margin-bottom:16px">
      <div class="toolbar" style="margin:0">
        <div style="width:200px"><label class="label" for="rMonth">対象月</label><input type="month" id="rMonth"></div>
        <span class="grow"></span>
        <button type="button" class="primary" id="rExport">この月の月次レポートをExcelで出力</button>
      </div>
      <div id="rDl" class="dl"></div>
    </div>
    <div id="revWarn" class="warn" hidden></div>
    <div id="kpis" class="kpis"></div>
    <div class="two" style="grid-template-columns:minmax(0,1fr) minmax(0,1fr)">
      <div class="panel"><h2>入院手術</h2><div class="scroll"><table class="list" id="rIn"></table></div></div>
      <div class="panel"><h2>外来手術</h2><div class="scroll"><table class="list" id="rOut"></table></div></div>
    </div>
    <div class="panel" style="margin-top:16px"><h2>執刀医別（眼数）</h2><div class="scroll"><table class="list" id="rDoc"></table></div></div>
    <div class="panel" style="margin-top:16px">
      <h2>見込み（残りの手術枠が埋まった場合）</h2>
      <div class="sim">
        <div><span class="label">残りの手術枠</span><div id="simSlots" style="font-size:20px;font-weight:700"></div></div>
        <div><label class="label" for="simFill">埋まり率 <output id="simFillOut">70%</output></label><input type="range" id="simFill" min="0" max="100" step="5" value="70"></div>
        <div><span class="label">月末の利益（収入）見込み</span><div id="simTotal" style="font-size:20px;font-weight:700"></div></div>
      </div>
      <p class="note" id="simNote"></p>
    </div>
    <div class="panel" style="margin-top:16px"><h2>支出（発注先ごと）</h2><div class="scroll"><table class="list" id="rSpend"></table></div></div>
  </section>

  <!-- 出力 -->
  <section id="vout" hidden>
    <div class="two" style="grid-template-columns:minmax(0,1fr) minmax(0,1fr)">
      <div class="panel">
        <h2>発注書</h2>
        <div class="toolbar"><input type="month" id="xMonth" style="width:180px" aria-label="表示する月"><span class="grow"></span>
          <button type="button" class="chip" id="xAll" aria-pressed="false">注文数0の品目も載せる</button></div>
        <div class="scroll"><table class="list" id="xOrders"></table></div>
        <div id="xDl" class="dl"></div>
      </div>
      <div class="panel" id="xMonthlyPanel">
        <h2>月次レポート</h2>
        <p class="note" style="margin-top:0">収益表・コスト計算用・支出合計・発注先ごとの発注集計を、1つのExcelファイルにまとめて出力します。</p>
        <div class="toolbar"><input type="month" id="xRevMonth" style="width:180px" aria-label="対象月"><button type="button" class="primary" id="xRevBtn">Excelで出力</button></div>
        <div id="xRevDl" class="dl"></div>
      </div>
    </div>
  </section>
</main>

<script>
const DAYS = 14;
const COLORS = {
  '白内障': { bg: '#D6E4F0', fg: '#123A5A' },
  '網膜硝子体': { bg: '#E2DAEE', fg: '#3D2A5C' },
  '緑内障': { bg: '#D5ECE8', fg: '#134E48' },
  '眼瞼': { bg: '#F3DEC9', fg: '#6B3410' },
  '涙道': { bg: '#D9E6D2', fg: '#28451D' }
};
const COLOR_OTHER = { bg: '#E6E2DA', fg: '#33373D' };
const EYES = [['右', '右眼'], ['左', '左眼'], ['両', '両眼']];

const S = {
  boot: null, procMap: {}, view: null, bookings: [], editingId: null,
  F: null,
  order: { vendor: '', group: '', search: '', only: false, cart: {}, editingId: null },
  orders: [], rev: null
};

const $ = id => document.getElementById(id);
const esc = s => String(s == null ? '' : s).replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
const call = (fn, ...a) => new Promise((res, rej) => google.script.run.withSuccessHandler(res).withFailureHandler(rej)[fn](...a));
const pd = s => { const [y, m, d] = s.split('-').map(Number); return new Date(y, m - 1, d); };
const fd = d => d.getFullYear() + '-' + String(d.getMonth() + 1).padStart(2, '0') + '-' + String(d.getDate()).padStart(2, '0');
const add = (s, n) => { const d = pd(s); d.setDate(d.getDate() + n); return fd(d); };
const diff = (a, b) => Math.round((pd(b) - pd(a)) / 864e5);
const md = s => { const d = pd(s); return (d.getMonth() + 1) + '/' + d.getDate() + '(' + '日月火水木金土'[d.getDay()] + ')'; };
const monday = s => { const d = pd(s); d.setDate(d.getDate() - ((d.getDay() + 6) % 7)); return fd(d); };
const monthRange = ym => { const [y, m] = ym.split('-').map(Number); return [ym + '-01', fd(new Date(y, m, 0))]; };
const yen = n => '¥' + Math.round(n || 0).toLocaleString('ja-JP');
const num = n => Math.round(n || 0).toLocaleString('ja-JP');
const colorOf = c => COLORS[c] || COLOR_OTHER;
const procOf = code => S.procMap[code];
const endOf = b => b.kind === '入院' ? add(b.admit, b.nights) : b.opDate;
const overlap = (a1, a2, b1, b2) => a1 <= b2 && b1 <= a2;
const role = () => S.boot.user.role;
const canEdit = () => role() === '管理' || role() === '入力';
const isAdmin = () => role() === '管理';
const bedName = no => (S.boot.beds.find(b => b.no === no) || { name: 'ベッド' + no }).name;
const eyeLabel = e => e === '両' ? '両眼' : e === '左' ? '左眼' : '右眼';
const errText = e => (e && e.message) ? e.message : String(e);
function setMsg(id, text, kind) { const el = $(id); el.textContent = text || ''; el.className = 'msg' + (kind ? ' ' + kind : ''); }
function chip(label, pressed, attrs, small) {
  return '<button type="button" class="chip" aria-pressed="' + (!!pressed) + '" ' + (attrs || '') + '>' + esc(label) + (small ? '<small>' + esc(small) + '</small>' : '') + '</button>';
}

/* ================= 起動 ================= */
async function init() {
  try {
    S.boot = await call('getBootstrap');
  } catch (e) {
    $('loading').textContent = errText(e);
    return;
  }
  S.boot.procs.forEach(p => { S.procMap[p.code] = p; });
  $('facility').textContent = S.boot.facility;
  $('who').textContent = S.boot.user.name + '（' + role() + '）';
  document.querySelector('[data-view="order"]').hidden = !canEdit();
  document.querySelector('[data-view="out"]').hidden = !canEdit();
  document.querySelector('[data-view="rev"]').hidden = !isAdmin();
  $('xMonthlyPanel').hidden = !isAdmin();
  if (!canEdit()) { $('formPanel').hidden = true; $('bookLayout').style.gridTemplateColumns = 'minmax(0,1fr)'; }
  document.querySelectorAll('.tab').forEach(t => { t.onclick = () => showView(t.dataset.view); });

  S.view = monday(S.boot.today);
  const ym = S.boot.today.slice(0, 7);
  ['oMonth', 'rMonth', 'xMonth', 'xRevMonth'].forEach(id => { $(id).value = ym; });
  $('oDate').value = S.boot.today;
  resetForm();
  bindBook();
  bindOrder();
  bindRev();
  bindOut();
  renderLegend();
  $('loading').hidden = true;
  $('vbook').hidden = false;
  await loadBookings();
}

function showView(v) {
  document.querySelectorAll('.tab').forEach(t => t.setAttribute('aria-pressed', t.dataset.view === v));
  ['book', 'order', 'rev', 'out'].forEach(k => { $('v' + k).hidden = k !== v; });
  window.scrollTo(0, 0);
  if (v === 'order') { if (!S.order.vendor) pickVendor(firstVendor()); loadOrders(); }
  if (v === 'rev') loadRev();
  if (v === 'out') loadOutOrders();
}

/* ================= 入院・手術 ================= */
function newForm() {
  return { patientId: '', kind: '', kindTouched: false, opDate: S.boot.today, proc: '', eye: '', doctor: '',
    admitMode: 'same', admit: '', nights: 1, nightsTouched: false, bed: 0, status: '仮', memo: '' };
}

function resetForm() {
  S.editingId = null;
  S.F = newForm();
  $('formTitle').textContent = '手術・入院の登録';
  $('btnSave').textContent = '登録する';
  $('btnCancelBk').hidden = true;
  $('btnNew').hidden = true;
  syncFormInputs();
  renderForm();
}

function syncFormInputs() {
  $('fPid').value = S.F.patientId;
  $('fOp').value = S.F.opDate;
  $('fAdmit').value = S.F.admit;
  $('fMemo').value = S.F.memo;
}

function bindBook() {
  $('prev').onclick = () => { S.view = add(S.view, -7); loadBookings(); };
  $('next').onclick = () => { S.view = add(S.view, 7); loadBookings(); };
  $('thisWeek').onclick = () => { S.view = monday(S.boot.today); loadBookings(); };
  $('fPid').oninput = e => { S.F.patientId = e.target.value; renderPreview(); };
  $('fMemo').oninput = e => { S.F.memo = e.target.value; };
  $('fOp').onchange = e => { S.F.opDate = e.target.value; followDates(); };
  $('fAdmit').onchange = e => { S.F.admit = e.target.value; S.F.admitMode = 'custom'; refreshBook(); };
  $('fBed').onchange = e => { S.F.bed = Number(e.target.value) || 0; refreshBook(); };
  $('formPanel').addEventListener('click', onFormClick);
  $('btnSave').onclick = saveBooking;
  $('btnCancelBk').onclick = cancelBooking;
  $('btnNew').onclick = () => { resetForm(); setMsg('msgBook', ''); renderGrid(); };
  $('grid').onclick = onGridClick;
  $('opList').onclick = e => { const tr = e.target.closest('tr[data-id]'); if (tr) editBooking(tr.dataset.id); };
}

function onFormClick(e) {
  const b = e.target.closest('button[data-f]');
  if (!b) return;
  const f = b.dataset.f, v = b.dataset.v;
  const F = S.F;
  if (f === 'kind') { F.kind = v; F.kindTouched = true; }
  if (f === 'proc') {
    F.proc = v;
    const p = procOf(v);
    if (p && !F.nightsTouched) F.nights = p.nights || 1;
    if (p && !F.kindTouched) F.kind = p.nights > 0 ? '入院' : '外来';
  }
  if (f === 'eye') F.eye = v;
  if (f === 'doc') F.doctor = v;
  if (f === 'admit') { F.admitMode = v; }
  if (f === 'nights') { F.nights = Number(v); F.nightsTouched = true; }
  if (f === 'status') F.status = v;
  followDates();
}

function followDates() {
  const F = S.F;
  if (F.admitMode === 'same') F.admit = F.opDate;
  if (F.admitMode === 'prev' && F.opDate) F.admit = add(F.opDate, -1);
  $('fAdmit').value = F.admit;
  const target = F.kind === '入院' ? F.admit : F.opDate;
  if (target && (target < S.view || target > add(S.view, DAYS - 1))) { S.view = monday(target); loadBookings(); return; }
  refreshBook();
}

function refreshBook() { renderForm(); renderGrid(); }

function renderForm() {
  const F = S.F;
  $('fKind').innerHTML = chip('入院', F.kind === '入院', 'data-f="kind" data-v="入院"') + chip('外来（日帰り）', F.kind === '外来', 'data-f="kind" data-v="外来"');
  // 術式（色分けごと）
  const groups = [];
  S.boot.procs.filter(p => p.active).forEach(p => { const g = p.color || 'その他'; if (groups.indexOf(g) < 0) groups.push(g); });
  $('fProc').innerHTML = groups.map(g => {
    const c = colorOf(g);
    return '<div class="grouphead"><i style="background:' + c.bg + ';border:1px solid ' + c.fg + '"></i>' + esc(g) + '</div><div class="chips grid3">' +
      S.boot.procs.filter(p => p.active && (p.color || 'その他') === g).map(p =>
        chip(p.name, F.proc === p.code, 'data-f="proc" data-v="' + esc(p.code) + '"', p.nights > 0 ? '標準 ' + p.nights + '泊' : '日帰り')).join('') + '</div>';
  }).join('');
  $('fEye').innerHTML = EYES.map(e => chip(e[1], F.eye === e[0], 'data-f="eye" data-v="' + e[0] + '"')).join('');
  $('fDoc').innerHTML = S.boot.doctors.map(d => chip(d.name, F.doctor === d.name, 'data-f="doc" data-v="' + esc(d.name) + '"')).join('') +
    chip('未定', F.doctor === '', 'data-f="doc" data-v=""');
  $('inOnly').hidden = F.kind !== '入院';
  $('fAdmitMode').innerHTML = chip('手術当日', F.admitMode === 'same', 'data-f="admit" data-v="same"') + chip('前日', F.admitMode === 'prev', 'data-f="admit" data-v="prev"') +
    chip('日付を指定', F.admitMode === 'custom', 'data-f="admit" data-v="custom"');
  $('fAdmit').hidden = F.admitMode !== 'custom';
  const nightOpts = [0, 1, 2, 3, 4, 5, 6, 7];
  if (nightOpts.indexOf(F.nights) < 0) nightOpts.push(F.nights);
  $('fNights').innerHTML = nightOpts.map(n => chip(n === 0 ? '0泊（日帰り入院）' : n + '泊' + (n + 1) + '日', F.nights === n, 'data-f="nights" data-v="' + n + '"')).join('');
  const bedOpts = '<option value="0">自動（空いているベッド）</option>' + S.boot.beds.map(b => '<option value="' + b.no + '">' + esc(b.name) + '</option>').join('');
  if ($('fBed').innerHTML !== bedOpts) $('fBed').innerHTML = bedOpts;
  $('fBed').value = String(F.bed);
  $('fStatus').innerHTML = chip('仮押さえ', F.status === '仮', 'data-f="status" data-v="仮"') + chip('確定', F.status === '確定', 'data-f="status" data-v="確定"');
  renderPreview();
}

function alloc() {
  const F = S.F;
  if (F.kind !== '入院' || !F.admit) return null;
  const s = F.admit, e = add(s, F.nights);
  const others = S.bookings.filter(b => b.kind === '入院' && b.id !== S.editingId);
  const busy = no => others.some(b => b.bed === no && overlap(s, e, b.admit, endOf(b)));
  if (F.bed) return busy(F.bed) ? { ng: bedName(F.bed) + 'はこの期間ふさがっています。自動を選ぶか、日付を変えてください。' } : { bed: F.bed, s, e };
  const free = S.boot.beds.find(b => !busy(b.no));
  return free ? { bed: free.no, s, e } : { ng: 'この期間に空いているベッドがありません。日付を変えてください。' };
}

function missing() {
  const F = S.F, m = [];
  if (!F.patientId.trim()) m.push('患者ID');
  if (!F.kind) m.push('区分');
  if (!F.opDate) m.push('手術日');
  if (!F.proc) m.push('術式');
  if (!F.eye) m.push('眼');
  return m;
}

function renderPreview() {
  const el = $('preview'), F = S.F;
  const m = missing();
  let ng = false;
  if (F.kind === '入院') {
    const a = alloc();
    if (a && a.ng) { el.className = 'preview ng'; el.textContent = a.ng; ng = true; }
    else if (a) {
      el.className = 'preview ok';
      el.innerHTML = '<strong>' + esc(bedName(a.bed)) + '</strong>　' + md(a.s) + ' 〜 ' + md(a.e) + '（' + F.nights + '泊）' +
        (F.opDate && F.opDate !== a.s ? '<br>手術日 ' + md(F.opDate) : '');
      if (F.opDate && (F.opDate < a.s || F.opDate > a.e)) { el.className = 'preview ng'; el.textContent = '手術日が入院期間の外になっています。入院日か泊数を確認してください。'; ng = true; }
    } else { el.className = 'preview'; el.innerHTML = ''; }
  } else if (F.kind === '外来') {
    el.className = 'preview info';
    el.textContent = '外来手術として登録します（ベッドは使いません）。';
  } else { el.className = 'preview'; el.innerHTML = ''; }
  $('btnSave').disabled = ng || m.length > 0;
  $('btnSave').textContent = m.length ? '未入力：' + m.join('・') : (S.editingId ? '変更を保存' : '登録する');
}

async function loadBookings() {
  $('range').textContent = md(S.view) + ' 〜 ' + md(add(S.view, DAYS - 1));
  try {
    S.bookings = await call('getBookings', add(S.view, -21), add(S.view, DAYS + 14));
  } catch (e) {
    setMsg('msgBook', errText(e), 'err');
    S.bookings = [];
  }
  refreshBook();
  renderOpList();
}

async function saveBooking() {
  const F = S.F;
  const p = { id: S.editingId, patientId: F.patientId.trim(), kind: F.kind, opDate: F.opDate, proc: F.proc, eye: F.eye,
    doctor: F.doctor, admit: F.admit, nights: F.nights, bed: F.bed, status: F.status, memo: F.memo };
  $('btnSave').disabled = true;
  setMsg('msgBook', '保存しています…');
  try {
    const r = await call('saveBooking', p);
    if (!r.ok) { setMsg('msgBook', r.message, 'err'); $('btnSave').disabled = false; return; }
    const keep = { opDate: F.opDate, doctor: F.doctor, kind: F.kind, kindTouched: F.kindTouched };
    resetForm();
    if (!p.id) Object.assign(S.F, keep);         // 続けて同じ日・同じ医師の入力がしやすいように
    syncFormInputs();
    setMsg('msgBook', r.message, 'ok');
    await loadBookings();
  } catch (e) {
    setMsg('msgBook', errText(e), 'err');
    $('btnSave').disabled = false;
  }
}

async function cancelBooking() {
  const b = S.bookings.find(x => x.id === S.editingId);
  if (!b || !confirm(b.patientId + ' の記録を取り消しますか？')) return;
  try {
    const r = await call('cancelBooking', b.id);
    setMsg('msgBook', r.message, r.ok ? 'ok' : 'err');
    if (r.ok) { resetForm(); await loadBookings(); }
  } catch (e) { setMsg('msgBook', errText(e), 'err'); }
}

function editBooking(id) {
  const b = S.bookings.find(x => x.id === id);
  if (!b) return;
  if (!canEdit()) {
    const p = procOf(b.proc);
    alert(b.patientId + '\\n' + b.kind + '　' + (p ? p.name : b.proc) + '（' + eyeLabel(b.eye) + '）\\n手術日 ' + md(b.opDate) +
      (b.kind === '入院' ? '\\n入院 ' + md(b.admit) + '〜' + md(endOf(b)) + '　' + bedName(b.bed) : '') + '\\n執刀医 ' + (b.doctor || '未定') + '\\n状態 ' + b.status +
      (b.memo ? '\\nメモ ' + b.memo : ''));
    return;
  }
  S.editingId = id;
  S.F = { patientId: b.patientId, kind: b.kind, kindTouched: true, opDate: b.opDate, proc: b.proc, eye: b.eye, doctor: b.doctor,
    admitMode: b.kind !== '入院' ? 'same' : b.admit === b.opDate ? 'same' : b.admit === add(b.opDate, -1) ? 'prev' : 'custom',
    admit: b.kind === '入院' ? b.admit : b.opDate, nights: b.kind === '入院' ? b.nights : 1, nightsTouched: true,
    bed: b.kind === '入院' ? b.bed : 0, status: b.status === '確定' ? '確定' : '仮', memo: b.memo };
  $('formTitle').textContent = '記録の変更';
  $('btnCancelBk').hidden = false;
  $('btnNew').hidden = false;
  setMsg('msgBook', '');
  syncFormInputs();
  refreshBook();
  $('formPanel').scrollIntoView({ behavior: 'smooth', block: 'start' });
}

function renderLegend() {
  const groups = [];
  S.boot.procs.filter(p => p.active).forEach(p => { const g = p.color || 'その他'; if (groups.indexOf(g) < 0) groups.push(g); });
  $('legend').innerHTML = groups.map(g => { const c = colorOf(g); return '<span><i style="background:' + c.bg + ';border-color:' + c.fg + '"></i>' + esc(g) + '</span>'; }).join('') +
    '<span><i style="background:#fff;border-style:dashed;border-color:#4A4F57"></i>斜線＝仮押さえ</span>';
}

function renderGrid() {
  const vs = S.view, ve = add(vs, DAYS - 1);
  const cols = 'grid-template-columns:repeat(' + DAYS + ',minmax(0,1fr))';
  const days = Array.from({ length: DAYS }, (_, i) => add(vs, i));
  const a = canEdit() ? alloc() : null;
  const usage = new Array(DAYS).fill(0);
  const ins = S.bookings.filter(b => b.kind === '入院');
  let h = '<div class="g-row"><div class="g-bed" style="font-weight:500;font-size:12px;color:#4A4F57">病床</div><div class="g-track" style="' + cols + '">' +
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
    ins.filter(b => b.bed === bed.no).forEach(b => {
      const e = endOf(b);
      if (e < vs || b.admit > ve) return;
      const si = Math.max(0, diff(vs, b.admit)), ei = Math.min(DAYS - 1, diff(vs, e));
      for (let i = si; i <= ei; i++) usage[i]++;
      const p = procOf(b.proc);
      const c = colorOf(p ? p.color : '');
      const tent = b.status !== '確定';
      const cls = 'g-bar' + (tent ? ' tent' : '') + (b.admit < vs ? ' cutL' : '') + (e > ve ? ' cutR' : '') + (b.id === S.editingId ? ' editing' : '');
      const label = b.patientId + '　' + (p ? p.name : b.proc) + '（' + b.eye + '）' + (tent ? '［仮］' : '');
      const title = label + '　入院' + md(b.admit) + '〜' + md(e) + '　手術' + md(b.opDate) + (b.doctor ? '　' + b.doctor : '') + (b.memo ? '　' + b.memo : '');
      t += '<button type="button" class="' + cls + '" data-act="bar" data-id="' + esc(b.id) + '" style="grid-column:' + (si + 1) + ' / span ' + (ei - si + 1) +
        ';background-color:' + c.bg + ';color:' + c.fg + (tent ? ';border-color:' + c.fg : '') + '" title="' + esc(title) + '">' + esc(label) + '</button>';
    });
    if (a && !a.ng && a.bed === bed.no && !(a.e < vs || a.s > ve)) {
      const si = Math.max(0, diff(vs, a.s)), ei = Math.min(DAYS - 1, diff(vs, a.e));
      for (let i = si; i <= ei; i++) usage[i]++;
      const p = procOf(S.F.proc);
      t += '<div class="g-bar pending" style="grid-column:' + (si + 1) + ' / span ' + (ei - si + 1) + '">' + (S.editingId ? '変更後' : '新規') + '　' + esc(p ? p.name : '') + '</div>';
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
  if (t.dataset.act === 'bar') { editBooking(t.dataset.id); return; }
  if (t.dataset.act === 'cell' && canEdit()) {
    const F = S.F;
    F.kind = '入院'; F.kindTouched = true;
    F.admit = t.dataset.date; F.bed = Number(t.dataset.bed);
    if (!F.opDate || F.opDate < F.admit || F.opDate > add(F.admit, F.nights)) F.opDate = F.admit;
    F.admitMode = F.admit === F.opDate ? 'same' : F.admit === add(F.opDate, -1) ? 'prev' : 'custom';
    syncFormInputs();
    refreshBook();
  }
}

function renderOpList() {
  const vs = S.view, ve = add(vs, DAYS - 1);
  const ops = S.bookings.filter(b => b.opDate >= vs && b.opDate <= ve).sort((a, b) => (a.opDate + a.patientId).localeCompare(b.opDate + b.patientId));
  if (!ops.length) { $('opList').innerHTML = '<tr><td class="empty">この期間の手術はありません。</td></tr>'; return; }
  $('opList').innerHTML = '<tr><th>手術日</th><th>区分</th><th>患者ID</th><th>術式</th><th>執刀医</th><th>状態</th></tr>' + ops.map(b => {
    const p = procOf(b.proc);
    return '<tr class="clickable" data-id="' + esc(b.id) + '"><td>' + md(b.opDate) + '</td><td><span class="badge ' + (b.kind === '入院' ? 'in' : 'out') + '">' + esc(b.kind) + '</span></td><td>' +
      esc(b.patientId) + '</td><td>' + esc(p ? p.name : b.proc) + '（' + esc(eyeLabel(b.eye)) + '）</td><td>' + esc(b.doctor || '—') + '</td><td>' +
      (b.status === '確定' ? '確定' : '<span class="badge tent">仮</span>') + '</td></tr>';
  }).join('');
}

/* ================= 発注 ================= */
function firstVendor() { const v = S.boot.vendors.find(x => x.active); return v ? v.name : ''; }
function vendorOf(name) { return S.boot.vendors.find(v => v.name === name); }
function vendorItems() { return S.boot.items.filter(it => it.vendor === S.order.vendor && it.active); }

function bindOrder() {
  $('oVendor').onclick = e => {
    const b = e.target.closest('button[data-v]'); if (!b) return;
    if (Object.keys(S.order.cart).length && b.dataset.v !== S.order.vendor && !confirm('入力中の発注内容を消して、発注先を切り替えますか？')) return;
    pickVendor(b.dataset.v);
  };
  $('oGroup').onclick = e => { const b = e.target.closest('button[data-g]'); if (!b) return; S.order.group = b.dataset.g; renderItems(); };
  $('oSearch').oninput = e => { S.order.search = e.target.value.trim(); renderItems(); };
  $('oOnly').onclick = () => { S.order.only = !S.order.only; $('oOnly').setAttribute('aria-pressed', S.order.only); renderItems(); };
  $('oItems').addEventListener('click', e => {
    const b = e.target.closest('button[data-step]'); if (!b) return;
    const id = b.dataset.id;
    const cur = S.order.cart[id] ? S.order.cart[id].qty : 0;
    setQty(id, Math.max(0, cur + Number(b.dataset.step)));
    const inp = document.querySelector('#oItems input[data-id="' + CSS.escape(id) + '"]');
    if (inp) { inp.value = S.order.cart[id] ? S.order.cart[id].qty : ''; inp.closest('.item').classList.toggle('on', !!S.order.cart[id]); }
    renderCart();
  });
  $('oItems').addEventListener('change', e => {
    const inp = e.target.closest('input[data-id]'); if (!inp) return;
    setQty(inp.dataset.id, Math.max(0, Math.round(Number(inp.value) || 0)));
    inp.value = S.order.cart[inp.dataset.id] ? S.order.cart[inp.dataset.id].qty : '';
    inp.closest('.item').classList.toggle('on', !!S.order.cart[inp.dataset.id]);
    renderCart();
  });
  $('cartLines').addEventListener('change', e => {
    const inp = e.target.closest('input[data-id]'); if (!inp) return;
    const l = S.order.cart[inp.dataset.id]; if (!l) return;
    if (inp.dataset.k === 'qty') { setQty(inp.dataset.id, Math.max(0, Math.round(Number(inp.value) || 0))); renderItems(); }
    if (inp.dataset.k === 'price') { const v = Number(inp.value); l.price = isFinite(v) && v >= 0 ? v : l.price; }
    renderCart();
  });
  $('oAmount').oninput = renderCart;
  $('oSave').onclick = saveOrder;
  $('oClear').onclick = () => { clearCart(); setMsg('msgOrder', ''); };
  $('oMonth').onchange = loadOrders;
  $('oHist').onclick = onHistClick;
  $('cartBarBtn').onclick = () => $('cartPanel').scrollIntoView({ behavior: 'smooth', block: 'start' });
}

function setQty(id, qty) {
  const it = S.boot.items.find(x => x.id === id);
  if (!it) return;
  if (qty > 0) S.order.cart[id] = { qty: qty, price: S.order.cart[id] ? S.order.cart[id].price : it.price };
  else delete S.order.cart[id];
}

function clearCart() {
  S.order.cart = {};
  S.order.editingId = null;
  $('oMemo').value = '';
  $('oAmount').value = '';
  $('cartTitle').textContent = '発注内容';
  $('oSave').textContent = 'この内容で発注を登録';
  renderItems();
  renderCart();
}

function pickVendor(name) {
  if (name !== S.order.vendor) setMsg('msgOrder', '');
  S.order.vendor = name;
  S.order.group = '';
  S.order.cart = {};
  S.order.editingId = null;
  $('cartTitle').textContent = '発注内容';
  $('oSave').textContent = 'この内容で発注を登録';
  $('oVendor').innerHTML = S.boot.vendors.filter(v => v.active).map(v => chip(v.name, v.name === name, 'data-v="' + esc(v.name) + '"', v.mode === '金額' ? '金額を入力' : '')).join('');
  const v = vendorOf(name);
  const amountMode = v && v.mode === '金額';
  $('oItemsBox').hidden = amountMode;
  $('oAmountBox').hidden = !amountMode;
  if (!amountMode) {
    const groups = [];
    vendorItems().forEach(it => { if (groups.indexOf(it.group) < 0) groups.push(it.group); });
    S.order.group = groups[0] || '';
    $('oGroup').innerHTML = groups.map(g => chip(g || 'その他', S.order.group === g, 'data-g="' + esc(g) + '"')).join('') +
      chip('すべて', false, 'data-g="*"');
  }
  renderItems();
  renderCart();
}

function renderItems() {
  const v = vendorOf(S.order.vendor);
  if (!v || v.mode === '金額') { $('oItems').innerHTML = ''; return; }
  document.querySelectorAll('#oGroup button').forEach(b => b.setAttribute('aria-pressed', b.dataset.g === S.order.group));
  const q = S.order.search.toLowerCase();
  let list = vendorItems();
  if (S.order.group !== '*' && !q) list = list.filter(it => it.group === S.order.group);
  if (q) list = list.filter(it => (it.name + ' ' + it.maker + ' ' + it.spec + ' ' + it.code).toLowerCase().indexOf(q) >= 0);
  if (S.order.only) list = list.filter(it => S.order.cart[it.id]);
  if (!list.length) { $('oItems').innerHTML = '<p class="empty">該当する品目がありません。品目は「品目マスタ」シートで追加できます。</p>'; return; }
  let h = '', last = null;
  list.forEach(it => {
    const head = (S.order.group === '*' || q || S.order.only ? (it.group || 'その他') + (it.sub ? '　' + it.sub : '') : it.sub);
    if (head && head !== last) { h += '<div class="subhead">' + esc(head) + '</div>'; last = head; }
    const c = S.order.cart[it.id];
    const spec = [it.maker, it.spec, it.pack].filter(Boolean).join('　');
    h += '<div class="item' + (c ? ' on' : '') + '"><div><div class="nm">' + (it.generic ? '<span class="badge tent">後発</span> ' : '') + esc(it.name) + '</div>' +
      (spec ? '<div class="sp">' + esc(spec) + '</div>' : '') + '</div><div class="pr">' + yen(it.price) + '</div>' +
      '<div class="stepper"><button type="button" data-step="-1" data-id="' + esc(it.id) + '" aria-label="' + esc(it.name) + ' を1減らす">−</button>' +
      '<input type="number" inputmode="numeric" min="0" step="1" data-id="' + esc(it.id) + '" value="' + (c ? c.qty : '') + '" aria-label="' + esc(it.name) + ' の数量">' +
      '<button type="button" data-step="1" data-id="' + esc(it.id) + '" aria-label="' + esc(it.name) + ' を1増やす">＋</button></div></div>';
  });
  $('oItems').innerHTML = h;
}

function renderCart() {
  const v = vendorOf(S.order.vendor);
  if (v && v.mode === '金額') {
    const a = Math.round(Number($('oAmount').value) || 0);
    $('cartLines').innerHTML = '<p class="note" style="margin:0">' + esc(v.name) + '：' + (a > 0 ? yen(a) : '金額を入力してください') + '</p>';
    $('cartTotal').textContent = yen(a);
    $('oSave').disabled = !(a > 0);
    $('cartBar').classList.remove('show');
    return;
  }
  const ids = Object.keys(S.order.cart);
  let total = 0;
  $('cartLines').innerHTML = ids.length ? ids.map(id => {
    const it = S.boot.items.find(x => x.id === id);
    const l = S.order.cart[id];
    total += Math.round(l.qty * l.price);
    return '<div class="line"><div>' + esc(it ? it.name : id) + '<div class="sp" style="font-size:12px;color:var(--muted)">単価 <input type="number" inputmode="numeric" min="0" step="1" data-id="' + esc(id) + '" data-k="price" value="' + l.price + '" style="width:96px;display:inline-block" aria-label="単価"></div></div>' +
      '<input type="number" inputmode="numeric" min="0" step="1" data-id="' + esc(id) + '" data-k="qty" value="' + l.qty + '" aria-label="数量">' +
      '<div style="text-align:right">' + yen(l.qty * l.price) + '</div></div>';
  }).join('') : '<p class="note" style="margin:0">品目の「＋」で数量を入れると、ここに並びます。</p>';
  $('cartTotal').textContent = yen(total);
  $('oSave').disabled = !ids.length;
  $('cartBar').classList.toggle('show', ids.length > 0);
  $('cartBarText').textContent = ids.length + '品目　' + yen(total);
}

async function saveOrder() {
  const v = vendorOf(S.order.vendor);
  if (!v) return;
  const p = { id: S.order.editingId, date: $('oDate').value, vendor: v.name, memo: $('oMemo').value };
  if (v.mode === '金額') p.amount = Number($('oAmount').value);
  else p.lines = Object.keys(S.order.cart).map(id => ({ itemId: id, qty: S.order.cart[id].qty, price: S.order.cart[id].price }));
  $('oSave').disabled = true;
  setMsg('msgOrder', '保存しています…');
  try {
    const r = await call('saveOrder', p);
    if (!r.ok) { setMsg('msgOrder', r.message, 'err'); $('oSave').disabled = false; return; }
    clearCart();
    setMsg('msgOrder', r.message, 'ok');
    $('oMonth').value = p.date.slice(0, 7);
    await loadOrders();
  } catch (e) {
    setMsg('msgOrder', errText(e), 'err');
    $('oSave').disabled = false;
  }
}

async function loadOrders() {
  const [from, to] = monthRange($('oMonth').value || S.boot.today.slice(0, 7));
  $('oHist').innerHTML = '<tr><td class="empty">読み込み中…</td></tr>';
  try { S.orders = await call('getOrders', from, to); } catch (e) { $('oHist').innerHTML = '<tr><td class="empty">' + esc(errText(e)) + '</td></tr>'; return; }
  if (!S.orders.length) { $('oHist').innerHTML = '<tr><td class="empty">この月の発注はありません。</td></tr>'; return; }
  $('oHist').innerHTML = '<tr><th>発注日</th><th>発注先</th><th class="num">品目数</th><th class="num">金額</th><th>メモ</th><th></th></tr>' + S.orders.map(o =>
    '<tr><td>' + md(o.date) + '</td><td>' + esc(o.vendor) + '</td><td class="num">' + o.lines.length + '</td><td class="num">' + yen(o.total) + '</td><td>' + esc(o.memo) + '</td>' +
    '<td style="white-space:nowrap"><button type="button" class="ghost sm" data-a="xls" data-id="' + esc(o.id) + '">発注書</button> ' +
    '<button type="button" class="ghost sm" data-a="edit" data-id="' + esc(o.id) + '">修正</button> ' +
    '<button type="button" class="ghost sm" data-a="del" data-id="' + esc(o.id) + '">取消</button></td></tr>').join('');
}

async function onHistClick(e) {
  const b = e.target.closest('button[data-a]'); if (!b) return;
  const o = S.orders.find(x => x.id === b.dataset.id); if (!o) return;
  if (b.dataset.a === 'xls') return exportOrder(o.id, false, 'oDl', b);
  if (b.dataset.a === 'del') {
    if (!confirm(md(o.date) + ' の ' + o.vendor + ' への発注（' + yen(o.total) + '）を取り消しますか？')) return;
    try { const r = await call('cancelOrder', o.id); setMsg('msgOrder', r.message, r.ok ? 'ok' : 'err'); await loadOrders(); }
    catch (err) { setMsg('msgOrder', errText(err), 'err'); }
    return;
  }
  if (b.dataset.a === 'edit') {
    pickVendor(o.vendor);
    S.order.editingId = o.id;
    $('oDate').value = o.date;
    $('oMemo').value = o.memo;
    const v = vendorOf(o.vendor);
    if (v && v.mode === '金額') $('oAmount').value = o.total;
    else o.lines.forEach(l => { if (l.itemId) S.order.cart[l.itemId] = { qty: l.qty, price: l.price }; });
    $('cartTitle').textContent = '発注の修正（' + md(o.date) + '）';
    $('oSave').textContent = '修正を保存';
    S.order.only = true; $('oOnly').setAttribute('aria-pressed', 'true');
    renderItems(); renderCart();
    window.scrollTo({ top: 0, behavior: 'smooth' });
  }
}

/* ================= 出力（共通） ================= */
function download(res, boxId) {
  const bin = atob(res.b64);
  const bytes = new Uint8Array(bin.length);
  for (let i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i);
  const blob = new Blob([bytes], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url; a.download = res.name;
  document.body.appendChild(a); a.click(); a.remove();
  $(boxId).innerHTML = '作成しました：<a href="' + url + '" download="' + esc(res.name) + '">' + esc(res.name) + '</a>（ダウンロードが始まらないときはこのリンクを押してください）';
}

async function exportOrder(id, all, boxId, btn) {
  const label = btn ? btn.textContent : '';
  if (btn) { btn.disabled = true; btn.textContent = '作成中…'; }
  $(boxId).textContent = 'Excelファイルを作成しています…';
  try { download(await call('exportOrderSheet', id, all), boxId); }
  catch (e) { $(boxId).textContent = errText(e); }
  finally { if (btn) { btn.disabled = false; btn.textContent = label; } }
}

async function exportMonthly(ym, boxId, btn) {
  btn.disabled = true;
  const label = btn.textContent;
  btn.textContent = '作成中…';
  $(boxId).textContent = 'Excelファイルを作成しています（数十秒かかることがあります）…';
  try { download(await call('exportMonthly', ym), boxId); }
  catch (e) { $(boxId).textContent = errText(e); }
  finally { btn.disabled = false; btn.textContent = label; }
}

function bindOut() {
  $('xMonth').onchange = loadOutOrders;
  $('xAll').onclick = () => $('xAll').setAttribute('aria-pressed', $('xAll').getAttribute('aria-pressed') !== 'true');
  $('xOrders').onclick = e => {
    const b = e.target.closest('button[data-id]'); if (!b) return;
    exportOrder(b.dataset.id, $('xAll').getAttribute('aria-pressed') === 'true', 'xDl', b);
  };
  $('xRevBtn').onclick = () => exportMonthly($('xRevMonth').value, 'xRevDl', $('xRevBtn'));
}

async function loadOutOrders() {
  const [from, to] = monthRange($('xMonth').value || S.boot.today.slice(0, 7));
  $('xOrders').innerHTML = '<tr><td class="empty">読み込み中…</td></tr>';
  let list;
  try { list = await call('getOrders', from, to); } catch (e) { $('xOrders').innerHTML = '<tr><td class="empty">' + esc(errText(e)) + '</td></tr>'; return; }
  if (!list.length) { $('xOrders').innerHTML = '<tr><td class="empty">この月の発注はありません。</td></tr>'; return; }
  $('xOrders').innerHTML = '<tr><th>発注日</th><th>発注先</th><th class="num">金額</th><th></th></tr>' + list.map(o =>
    '<tr><td>' + md(o.date) + '</td><td>' + esc(o.vendor) + '</td><td class="num">' + yen(o.total) + '</td><td><button type="button" class="ghost sm" data-id="' + esc(o.id) + '">発注書をダウンロード</button></td></tr>').join('');
}

/* ================= 収益・支出 ================= */
function bindRev() {
  $('rMonth').onchange = loadRev;
  $('simFill').oninput = renderSim;
  $('rExport').onclick = () => exportMonthly($('rMonth').value, 'rDl', $('rExport'));
}

async function loadRev() {
  const ym = $('rMonth').value;
  if (!ym) return;
  $('kpis').innerHTML = '<div class="kpi"><div class="l">読み込み中…</div></div>';
  try { S.rev = await call('getMonthSummary', ym); renderRev(); }
  catch (e) { $('kpis').innerHTML = '<div class="kpi"><div class="l">' + esc(errText(e)) + '</div></div>'; }
}

function revTable(t) {
  if (!t.rows.length) return '<tr><td class="empty">表示する区分がありません。</td></tr>';
  return '<tr><th>区分</th><th class="num">眼数</th><th class="num">1眼あたり収入</th><th class="num">収入</th></tr>' +
    t.rows.map(r => '<tr><td>' + esc(r.cat) + (r.triple ? ' <span class="badge tent">トリプル</span>' : '') + '</td><td class="num">' + r.eyes + '</td><td class="num">' + num(r.unit) + '</td><td class="num">' + num(r.eyes * r.unit) + '</td></tr>').join('') +
    '<tr class="total"><td>合計</td><td class="num">' + t.eyes + '</td><td></td><td class="num">' + num(t.income) + '</td></tr>';
}

function renderRev() {
  const d = S.rev;
  const sIn = d.surgery.in, sOut = d.surgery.out;
  const income = d.inpatient.income + sIn.income + sOut.income;
  let spend = 0;
  const byVendor = {};
  Object.keys(d.spendByDay).forEach(day => Object.keys(d.spendByDay[day]).forEach(v => { byVendor[v] = (byVendor[v] || 0) + d.spendByDay[day][v]; spend += d.spendByDay[day][v]; }));
  const dispo = Object.keys(d.dispoByDay).reduce((s, k) => s + d.dispoByDay[k], 0);

  const warn = [];
  if (d.unknownCats.length) warn.push('収益区分マスタにない区分があります（' + d.unknownCats.map(esc).join('、') + '）。収入0で集計しています。');
  const zero = sIn.rows.concat(sOut.rows).filter(r => r.eyes > 0 && !r.unit).map(r => r.cat);
  if (zero.length) warn.push('1眼あたり収入が0の区分に手術があります（' + zero.map(esc).join('、') + '）。');
  if (!d.inpatient.unit && d.inpatient.people) warn.push('設定シートの「入院1人あたり収入(円)」が0です。');
  $('revWarn').hidden = !warn.length;
  $('revWarn').innerHTML = warn.join('<br>');

  const k = [
    { l: '利益（収入）合計', v: yen(income), s: '入院 ' + yen(d.inpatient.income) + '＋手術 ' + yen(sIn.income + sOut.income), main: true },
    { l: '支出合計', v: yen(spend + dispo), s: '発注 ' + yen(spend) + '＋手術ディスポ代 ' + yen(dispo), main: true },
    { l: '請求額合計', v: yen(d.charges), s: '術式マスタの請求額から計算' },
    { l: '入院（人）', v: d.inpatient.people + '人', s: '入院日がこの月の人数' },
    { l: '入院手術', v: sIn.eyes + '眼', s: '外来手術 ' + sOut.eyes + '眼' },
    { l: '手術合計（トリプル分解分）', v: (sIn.eyes + sOut.eyes + d.surgery.tripleEyes) + '眼', s: 'うち仮押さえ ' + d.surgery.tentative + '件' }
  ];
  $('kpis').innerHTML = k.map(x => '<div class="kpi' + (x.main ? ' main' : '') + '"><div class="l">' + x.l + '</div><div class="v">' + x.v + '</div><div class="s">' + x.s + '</div></div>').join('');
  $('rIn').innerHTML = revTable(sIn);
  $('rOut').innerHTML = revTable(sOut);

  const docs = d.doctors;
  const sumDoc = (rows, doc) => rows.reduce((s, r) => s + (r.byDoc[doc] || 0), 0);
  $('rDoc').innerHTML = docs.length ? '<tr><th></th>' + docs.map(x => '<th class="num">' + esc(x) + '</th>').join('') + '</tr>' +
    '<tr><td>入院（人）</td>' + docs.map(x => '<td class="num">' + (d.inpatient.peopleByDoc[x] || 0) + '</td>').join('') + '</tr>' +
    '<tr><td>入院手術</td>' + docs.map(x => '<td class="num">' + sumDoc(sIn.rows, x) + '</td>').join('') + '</tr>' +
    '<tr><td>外来手術</td>' + docs.map(x => '<td class="num">' + sumDoc(sOut.rows, x) + '</td>').join('') + '</tr>' +
    '<tr class="total"><td>合計</td>' + docs.map(x => '<td class="num">' + (sumDoc(sIn.rows, x) + sumDoc(sOut.rows, x)) + '</td>').join('') + '</tr>'
    : '<tr><td class="empty">医師マスタに医師が登録されていません。</td></tr>';

  const vs = d.vendors.map(v => v.name);
  $('rSpend').innerHTML = '<tr><th>発注先</th><th class="num">金額</th></tr>' + vs.map(v => '<tr><td>' + esc(v) + '</td><td class="num">' + yen(byVendor[v] || 0) + '</td></tr>').join('') +
    '<tr><td>手術ディスポ代（自動計算）</td><td class="num">' + yen(dispo) + '</td></tr><tr class="total"><td>合計</td><td class="num">' + yen(spend + dispo) + '</td></tr>';
  renderSim();
}

function renderSim() {
  const d = S.rev;
  if (!d) return;
  const fill = Number($('simFill').value) / 100;
  $('simFillOut').textContent = Math.round(fill * 100) + '%';
  const income = d.inpatient.income + d.surgery.in.income + d.surgery.out.income;
  const extra = d.sim.remainingSlots * fill * d.sim.avgIncomePerCase;
  $('simSlots').textContent = d.sim.remainingSlots + '枠（月間 ' + d.sim.totalSlots + '枠）';
  $('simTotal').textContent = yen(income + extra);
  $('simNote').textContent = '見込み＝現在の利益（収入）＋残りの手術枠×埋まり率×1件あたり平均収入（' + yen(d.sim.avgIncomePerCase) + '）。手術枠は「手術枠マスタ」で設定します。病床稼働率 ' +
    (d.sim.capacity ? Math.round(100 * d.sim.occupied / d.sim.capacity) : 0) + '%。';
}

init();
</script>
</body>
</html>
`;
