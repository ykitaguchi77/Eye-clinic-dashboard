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
