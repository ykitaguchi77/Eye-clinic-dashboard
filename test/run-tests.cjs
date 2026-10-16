/**
 * サーバー側の動作確認（Apps Script のモック上で src/Code.gs を動かす）
 *   TZ=Asia/Tokyo node test/run-tests.cjs
 */
const load = require('./load.cjs');
const { paste } = load;
let failed = 0;
const ok = (cond, name) => { console.log((cond ? 'ok   ' : 'FAIL ') + name); if (!cond) failed++; };
const throwsLike = (fn, re, name) => { try { fn(); ok(false, name); } catch (e) { ok(re.test(e.message), name + '（' + e.message + '）'); } };

const { ctx, mock } = load();
const ss = mock.state.active;

// --- 初期設定（持ち主が最初に開く） ---
const OWNER = 'admin@example.com', STAFF = 'staff@example.com', VIEWER = 'viewer@example.com', MANAGER = 'manager@example.com';
mock.as(STAFF);
ok(/持ち主/.test(JSON.stringify(ctx.doGet && (() => { try { ctx.doGet(); return 'opened'; } catch (e) { return e.message; } })())) || !ss.getSheetByName('設定'), '持ち主より先にスタッフが開いても初期設定は走らない');
mock.as(OWNER);
ctx.doGet();
ok(Object.values(ctx.__T.SH).every(n => ss.getSheetByName(n)), '持ち主が開くと全シートが作られる');
const adminSs = mock.state.all.find(x => x.id === mock.state.props[ctx.__T.ADMIN_PROP]);
ok(adminSs && Object.values(ctx.__T.ASH).every(n => adminSs.getSheetByName(n)), '管理者用スプレッドシートが別に作られる');
ok(ss.getSheetByName('術式マスタ').getProtections().length === 1 && ss.getSheetByName('手術・入院台帳').getProtections().length === 0, 'マスタは保護、台帳は保護しない');

// --- テスト用マスタ（架空のデータ） ---
paste(mock, '術式マスタ', [
  ['A', '術式A', '区分A', 1, '白内障', true],
  ['B', '術式B', '区分B', 0, '眼瞼', true],
  ['T', '術式T', '区分T', 2, '網膜硝子体', true]
]);
paste(mock, '術式単価', [['A', 100000, '', 30000, ''], ['B', 50000, 90000, 5000, 8000], ['T', 400000, '', 150000, '']], true);
paste(mock, '収益区分マスタ', [
  ['区分A', 70000, false, true, true, true],
  ['区分B', 45000, false, false, true, true],
  ['区分T', 250000, true, true, false, true]
], true);
paste(mock, '医師マスタ', [['医師1', true], ['医師2', true]]);
paste(mock, 'ベッドマスタ', [[1, 'ベッド1', true], [2, 'ベッド2', true]]);
paste(mock, '発注先マスタ', [['卸X', '卸X株式会社', '品目', false, true], ['レンズ', '', '金額', true, true]]);
paste(mock, '品目マスタ', [
  ['', '卸X', '区分1', '棚A', '001', '品目1', 'メーカーP', '規格1', '10個', 1200, 1000, false, true],
  ['', '卸X', '区分1', '棚A', '002', '品目2', 'メーカーQ', '', '1本', '', 2500, true, true],
  ['', '卸X', '区分2', '', '003', '品目3', '', '', '', '', 300, false, false]
]);
adminSs.getSheetByName('管理者設定').getRange(2, 2).setValue(50000);

const boot = ctx.getBootstrap();
ok(boot.user.role === '管理', '持ち主は「管理」');
ok(boot.items.length === 3 && boot.items.every(i => /^I\d{4}$/.test(i.id)), '品目IDが空なら自動で振られる');
ok(boot.procs[0].charge1 === 100000 && boot.cats[0].unit === 70000, '管理者には単価が見える');

// --- 入院・外来の登録 ---
const save = p => ctx.saveBooking(p);
let r = save({ patientId: '001', kind: '入院', opDate: '2026-11-02', admit: '2026-11-01', nights: 2, proc: 'A', eye: '右', doctor: '医師1', status: '確定' });
ok(r.ok && r.bed === 1, '入院：空いているベッド1に入る');
r = save({ patientId: '002', kind: '入院', opDate: '2026-11-03', nights: 1, proc: 'T', eye: '左', doctor: '医師2' });
ok(r.ok && r.bed === 2, '入院：重なる期間は次のベッドに入る（入院日の省略は手術日）');
r = save({ patientId: '003', kind: '入院', opDate: '2026-11-03', nights: 1, proc: 'A', eye: '右' });
ok(!r.ok && /空いているベッドがありません/.test(r.message), '入院：満床なら登録できない');
r = save({ patientId: '003', kind: '入院', opDate: '2026-11-03', nights: 1, proc: 'A', eye: '右', bed: 1 });
ok(!r.ok && /ふさがっています/.test(r.message), '入院：指定したベッドがふさがっていれば登録できない');
r = save({ patientId: '004', kind: '外来', opDate: '2026-11-03', proc: 'B', eye: '両', doctor: '医師1', status: '確定' });
ok(r.ok && r.bed === 0, '外来：ベッドを使わずに登録できる');
r = save({ patientId: '005', kind: '入院', opDate: '2026-11-01', admit: '2026-11-03', nights: 1, proc: 'A', eye: '右' });
ok(!r.ok && /入院日が手術日より後/.test(r.message), '入院日が手術日より後ならエラー');
r = save({ patientId: '006', kind: '外来', opDate: '2026-11-04', proc: 'A' });
ok(!r.ok && /右・左・両眼/.test(r.message), '眼が未選択ならエラー');
const b1 = ctx.getBookings('2026-11-01', '2026-11-30');
ok(b1.length === 3, '期間の記録が取れる（取消・失敗分は含まない）');
const outId = b1.find(b => b.patientId === '004').id;
r = save({ id: outId, patientId: '004', kind: '外来', opDate: '2026-11-03', proc: 'B', eye: '右', doctor: '医師2', status: '確定' });
ok(r.ok, '記録を変更できる');
r = ctx.cancelBooking(b1.find(b => b.patientId === '002').id);
ok(r.ok && ctx.getBookings('2026-11-01', '2026-11-30').length === 2, '記録を取り消せる');
save({ patientId: '007', kind: '入院', opDate: '2026-11-10', nights: 3, proc: 'T', eye: '両', doctor: '医師2', status: '仮' });

// --- 発注 ---
const items = ctx.getBootstrap().items;
const i1 = items.find(i => i.name === '品目1').id, i2 = items.find(i => i.name === '品目2').id, i3 = items.find(i => i.name === '品目3').id;
r = ctx.saveOrder({ date: '2026-11-02', vendor: '卸X', lines: [{ itemId: i1, qty: 2 }, { itemId: i2, qty: 1, price: 2000 }, { itemId: i3, qty: 0 }] });
ok(r.ok, '品目の発注を登録できる（数量0は無視）');
const oid = r.id;
r = ctx.saveOrder({ date: '2026-11-09', vendor: '卸X', lines: [{ itemId: i1, qty: 1 }] });
r = ctx.saveOrder({ date: '2026-11-30', vendor: 'レンズ', amount: 120000, memo: 'IOL' });
ok(r.ok, '金額だけの発注を登録できる');
r = ctx.saveOrder({ date: '2026-11-02', vendor: '卸X', lines: [] });
ok(!r.ok, '数量のない発注は登録できない');
let orders = ctx.getOrders('2026-11-01', '2026-11-30');
const o1 = orders.find(o => o.id === oid);
ok(o1.total === 2 * 1000 + 2000, '単価を変えた品目はその単価で計算する');
r = ctx.saveOrder({ id: oid, date: '2026-11-02', vendor: '卸X', lines: [{ itemId: i1, qty: 5 }] });
orders = ctx.getOrders('2026-11-01', '2026-11-30');
ok(r.ok && orders.length === 3 && orders.find(o => o.id === oid).total === 5000, '発注を修正すると置き換わる');

// --- 月次の集計 ---
const md = ctx.getMonthSummary('2026-11');
ok(md.inpatient.people === 2 && md.inpatient.income === 100000, '入院人数×1人あたり収入');
ok(md.surgery.in.eyes === 3 && md.surgery.in.income === 70000 + 2 * 250000, '入院手術：眼数×1眼あたり収入（両眼は2眼）');
ok(md.surgery.out.eyes === 1 && md.surgery.out.income === 45000, '外来手術は外来の記録だけで集計');
ok(md.surgery.tripleEyes === 2, 'トリプル分解分');
ok(md.dispoByDay['2026-11-02'] === 30000 && md.dispoByDay['2026-11-10'] === 300000, '手術ディスポ代＝術式のコスト（両眼は2倍）');
ok(md.spendByDay['2026-11-02']['卸X'] === 5000 && md.spendByDay['2026-11-30']['レンズ'] === 120000, '支出は発注日・発注先ごと');
const os = md.orderSummary.find(s => s.vendor === '卸X');
const row1 = os.rows.find(x => x.id === i1);
ok(row1.weeks[0] === 5000 && row1.weeks[1] === 1000 && row1.qty === 6, '発注集計：1〜7日が1週目、8〜14日が2週目');
ok(os.rows.some(x => x.id === i3) === false, '使っていない品目（有効でない・発注なし）は発注集計に載せない');

// --- 出力の設計図 ---
const T = ctx.__T;
const rev = T.buildRevenueModel_(md);
const cellsAt = (m, col) => Object.keys(m.cells).filter(k => k.endsWith(',' + col)).map(k => [Number(k.split(',')[0]), m.cells[k].v]);
const totals = cellsAt(rev, 2).filter(([, v]) => v === '合計').map(([r]) => r).sort((a, b) => a - b);
const inTotal = rev.cells[totals[1] + ',5'].v, outTotal = rev.cells[totals[2] + ',5'].v;
const rangeStart = f => Number(/\(E(\d+):/.exec(f)[1]);
ok(rangeStart(outTotal) > totals[1], '収益表：外来手術の合計は入院手術の行を含まない（' + outTotal + '）');
ok(/^=SUM\(E\d+:E\d+\)$/.test(inTotal), '収益表：入院手術の合計式');
const exp = T.buildExpenseModel_(md);
ok(exp.maxR >= 4 + 30 && exp.cells['3,2'].v === '日付', '支出合計：月の全日の行がある');
const xl = ctx.exportMonthly('2026-11');
const sheets = JSON.parse(Buffer.from(xl.b64, 'base64').toString('utf8'));
ok(xl.name === '月次レポート_202611.xlsx' && sheets.map(s => s.name).join() === '収益表,コスト計算用,支出合計,発注集計_卸X', '月次レポートのシート構成');
ok(mock.state.created.filter(s => s !== adminSs).every(s => s.trashed), '出力に使った一時スプレッドシートはゴミ箱に移す');
const fax = JSON.parse(Buffer.from(ctx.exportOrderSheet(oid, false).b64, 'base64').toString('utf8'))[0];
ok(fax.data[0][1] === '卸X株式会社　御中' && fax.data.some(row => row[2] === '品目1' && row[6] === 5), '発注書：宛名と注文した品目');

// --- 権限（Googleの共有設定で決まる） ---
mock.share(ss, STAFF, 'edit');
mock.share(ss, VIEWER, 'view');
mock.share(ss, MANAGER, 'edit');
mock.share(adminSs, MANAGER, 'edit');
mock.as(STAFF);
const sb = ctx.getBootstrap();
ok(sb.user.role === '入力', '編集者として共有された人は「入力」');
ok(sb.procs[0].charge1 === undefined && sb.cats.length === 0 && sb.perInpatient === null && sb.items.length === 3, '入力：単価・収入は受け取らないが品目は見える');
throwsLike(() => ctx.getMonthSummary('2026-11'), /権限/, '入力：収益は見られない');
ok(ctx.saveBooking({ patientId: '101', kind: '外来', opDate: '2026-11-20', proc: 'A', eye: '右', doctor: '医師1' }).ok, '入力：記録を登録できる（台帳は保護していない）');
ok(ctx.saveOrder({ date: '2026-11-20', vendor: '卸X', lines: [{ itemId: i1, qty: 1 }] }).ok, '入力：発注を登録できる');
mock.as(VIEWER);
ok(ctx.getBootstrap().user.role === '閲覧', '閲覧者として共有された人は「閲覧」');
throwsLike(() => ctx.saveBooking({}), /権限/, '閲覧：登録できない');
mock.as('stranger@example.com');
throwsLike(() => ctx.getBootstrap(), /権限/, '共有されていない人は使えない');
mock.as(MANAGER);
const mb = ctx.getBootstrap();
ok(mb.user.role === '管理' && mb.procs[0].charge1 === 100000 && mb.perInpatient === 50000, '管理者用スプレッドシートを共有された人は「管理」');
ok(ctx.getMonthSummary('2026-11').surgery.out.eyes === 2, '管理：収益を見られる');
mock.as(OWNER);
ss.getSheetByName('術式マスタ').protection = null;
ss.getSheetByName('術式マスタ').getRange(1, 3).setValue('入院パターン');
mock.state.cache = {};
throwsLike(() => ctx.getBootstrap(), /術式マスタ.*3列目/, '見出しが違えば分かりやすいエラーになる');

console.log(failed ? '\n' + failed + ' 件失敗' : '\nすべて成功');
process.exitCode = failed ? 1 : 0;
