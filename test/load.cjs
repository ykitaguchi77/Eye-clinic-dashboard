/**
 * src/Code.gs をモックの Apps Script 環境に読み込む（Node 用）
 *   const { ctx, mock } = require('./load.cjs')();
 *   ctx.doGet();  // 初期設定が走る
 */
const fs = require('fs');
const path = require('path');
const vm = require('vm');

module.exports = function load(codePath) {
  if (process.env.TZ !== 'Asia/Tokyo') throw new Error('TZ=Asia/Tokyo で実行してください');
  const sandbox = { console, Buffer, Date, Math, JSON, TextEncoder };
  vm.createContext(sandbox);
  vm.runInContext(fs.readFileSync(path.join(__dirname, 'gas-mock.js'), 'utf8'), sandbox);
  Object.assign(sandbox, sandbox.GasMock.globals);
  sandbox.GasMock.reset();
  const code = fs.readFileSync(codePath || path.join(__dirname, '..', 'src', 'Code.gs'), 'utf8');
  // const/let のトップレベル宣言も外から参照できるように、最後に関数群をまとめて公開する
  vm.runInContext(code + '\n;globalThis.__T = { SH, HEAD, B, O, SheetModel_, monthData_, buildRevenueModel_, buildExpenseModel_, buildCostModel_, buildOrderSummaryModel_, buildOrderModel_, procs_, cats_, items_, vendors_, bookings_, orders_ };', sandbox);
  return { ctx: sandbox, mock: sandbox.GasMock, T: sandbox.__T };
};

/** シートの2行目以降を rows で置き換える（初期マスタの貼り付けを真似る） */
module.exports.paste = function paste(mock, sheetName, rows) {
  const sh = mock.state.active.getSheetByName(sheetName);
  sh.data = sh.data.slice(0, 1);
  if (rows.length) sh.getRange(2, 1, rows.length, rows[0].length).setValues(rows);
};
