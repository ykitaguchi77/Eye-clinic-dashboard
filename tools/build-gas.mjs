#!/usr/bin/env node
/**
 * src/Code.gs と src/Index.html を1つにまとめた gas/Code.gs を作る。
 * Node.js を使えないPCでは、この gas/Code.gs を Apps Script エディタに貼り付けるだけで動く。
 *
 *   node tools/build-gas.mjs          生成する
 *   node tools/build-gas.mjs --check  src と gas/Code.gs がずれていないか確認する（ずれていれば終了コード1）
 */
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const code = fs.readFileSync(path.join(ROOT, 'src/Code.gs'), 'utf8');
const html = fs.readFileSync(path.join(ROOT, 'src/Index.html'), 'utf8');

// テンプレートリテラルに入れるため \ ` ${ をエスケープ
const escaped = html.replace(/\\/g, '\\\\').replace(/`/g, '\\`').replace(/\$\{/g, '\\${');

const out =
`/**
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

${code.trimEnd()}

/* ------------------------------------------------------------
 * 画面のHTML（src/Index.html と同じ内容）
 * ---------------------------------------------------------- */
const INDEX_HTML_ = \`${escaped}\`;
`;

const target = path.join(ROOT, 'gas/Code.gs');
if (process.argv.includes('--check')) {
  const cur = fs.existsSync(target) ? fs.readFileSync(target, 'utf8') : '';
  if (cur !== out) {
    console.error('gas/Code.gs が src/ と一致しません。node tools/build-gas.mjs を実行してください。');
    process.exit(1);
  }
  console.log('gas/Code.gs は最新です。');
} else {
  fs.mkdirSync(path.dirname(target), { recursive: true });
  fs.writeFileSync(target, out);
  console.log('gas/Code.gs を作成しました（' + Buffer.byteLength(out) + ' バイト）。');
}
