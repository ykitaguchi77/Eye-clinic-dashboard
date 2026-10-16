#!/usr/bin/env node
/**
 * 入院・手術・発注管理 一括セットアップ／更新スクリプト
 *
 *   node setup.mjs            初回：Googleにログイン → スプレッドシート作成 → コード転送 → ウェブアプリ公開
 *                             2回目以降：コードを転送して、同じURLのまま新しい版に更新
 *   node setup.mjs --title "名前"   作成するスプレッドシートの名前を指定（初回のみ）
 *
 * 必要なもの：Node.js 20以上、Apps Script API の有効化
 *   （https://script.google.com/home/usersettings で「オン」にする。1回だけ）
 */
import { spawnSync } from 'node:child_process';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const ROOT = path.dirname(fileURLToPath(import.meta.url));
const CLASP = ['-y', '@google/clasp@3'];
const CLASP_JSON = path.join(ROOT, '.clasp.json');
const DEPLOY_JSON = path.join(ROOT, '.deployment.json');
const PULL_DIR = path.join(ROOT, '.clasp-pull');
const isWin = process.platform === 'win32';

process.chdir(ROOT);

const titleIdx = process.argv.indexOf('--title');
const TITLE = titleIdx > 0 && process.argv[titleIdx + 1] ? process.argv[titleIdx + 1] : '入院・手術・発注管理';

function fail(msg) {
  console.error('\n[エラー] ' + msg);
  process.exit(1);
}

function quote(a) {
  return isWin && /[\s"&|<>^]/.test(a) ? '"' + a.replace(/"/g, '\\"') + '"' : a;
}

/** clasp を実行。capture=true なら出力を返す（画面にも表示） */
function clasp(args, { capture = false } = {}) {
  const full = [...CLASP, ...args];
  const r = spawnSync(isWin ? 'npx.cmd' : 'npx', isWin ? full.map(quote) : full, {
    cwd: ROOT,
    shell: isWin,
    stdio: capture ? ['inherit', 'pipe', 'inherit'] : 'inherit',
    encoding: 'utf8'
  });
  if (r.error) fail('clasp を起動できませんでした: ' + r.error.message);
  if (r.status !== 0) {
    if (capture && r.stdout) process.stdout.write(r.stdout);
    fail('clasp ' + args[0] + ' が失敗しました。上のメッセージを確認してください。\n' +
      '  ・「User has not enabled the Apps Script API」と出た場合は\n' +
      '    https://script.google.com/home/usersettings で Apps Script API をオンにして、もう一度実行してください。');
  }
  return r.stdout || '';
}

/** 出力の中から最後の JSON オブジェクトを取り出す */
function lastJson(text) {
  const end = text.lastIndexOf('}');
  for (let start = text.lastIndexOf('{', end); start >= 0; start = text.lastIndexOf('{', start - 1)) {
    try { return JSON.parse(text.slice(start, end + 1)); } catch { /* 外側の { を探す */ }
  }
  fail('clasp の出力を読み取れませんでした:\n' + text);
}

function webAppUrl(id) { return 'https://script.google.com/macros/s/' + id + '/exec'; }

// ---- 前提チェック ----
const major = Number(process.versions.node.split('.')[0]);
if (major < 20) fail('Node.js 20 以上が必要です（現在 ' + process.versions.node + '）。https://nodejs.org から LTS 版を入れてください。');
for (const f of ['src/Code.gs', 'src/Index.html', 'src/appsscript.json']) {
  if (!fs.existsSync(path.join(ROOT, f))) fail(f + ' が見つかりません。リポジトリのフォルダで実行してください。');
}

// ---- ログイン（未ログイン時のみ） ----
if (!fs.existsSync(path.join(os.homedir(), '.clasprc.json'))) {
  console.log('\n▶ Googleにログインします。ブラウザが開いたら、クリニックの Google Workspace アカウントで許可してください。\n');
  clasp(['login']);
}

const today = new Date().toLocaleDateString('sv-SE'); // yyyy-mm-dd（ローカル時刻）

if (!fs.existsSync(CLASP_JSON)) {
  // ---- 初回：作成 → 転送 → 公開 ----
  console.log('\n▶ スプレッドシート「' + TITLE + '」とスクリプトを作成します…');
  fs.rmSync(PULL_DIR, { recursive: true, force: true });
  const created = lastJson(clasp(['create-script', '--type', 'sheets', '--title', TITLE, '--rootDir', '.clasp-pull', '--json'], { capture: true }));
  // 作成時に取り込まれる初期ファイルは使わない。転送元を src/ に切り替える
  const conf = JSON.parse(fs.readFileSync(CLASP_JSON, 'utf8'));
  conf.rootDir = 'src';
  fs.writeFileSync(CLASP_JSON, JSON.stringify(conf, null, 2));
  fs.rmSync(PULL_DIR, { recursive: true, force: true });

  console.log('\n▶ コードを転送します…');
  clasp(['push', '--force']);

  console.log('\n▶ ウェブアプリとして公開します…');
  const dep = lastJson(clasp(['create-deployment', '--description', 'v0.2 ' + today, '--json'], { capture: true }));
  fs.writeFileSync(DEPLOY_JSON, JSON.stringify({ deploymentId: dep.deploymentId, spreadsheetId: created.parentId, scriptId: created.scriptId }, null, 2));

  console.log('\n==================== 完了 ====================');
  console.log('スプレッドシート : https://docs.google.com/spreadsheets/d/' + created.parentId + '/edit');
  console.log('ダッシュボードURL : ' + webAppUrl(dep.deploymentId));
  console.log('\n次にやること：');
  console.log(' 1. ダッシュボードURLを「自分で」最初に開き、表示される権限の確認を許可する');
  console.log('    （初回アクセス時にシートが自動で作られ、あなたが管理者として登録されます）');
  console.log(' 2. スプレッドシートの各マスタ（単価・床数・手術枠・ユーザー）を入力する');
  console.log(' 3. ダッシュボードURLを職員に共有する');
  console.log('==============================================\n');
} else {
  // ---- 2回目以降：転送 → 同じURLで更新 ----
  if (!fs.existsSync(DEPLOY_JSON)) fail('.deployment.json がありません。初回セットアップをしたフォルダで実行してください。');
  const info = JSON.parse(fs.readFileSync(DEPLOY_JSON, 'utf8'));

  console.log('\n▶ コードを転送します…');
  clasp(['push', '--force']);

  console.log('\n▶ 公開中のウェブアプリを新しい版に更新します…');
  clasp(['create-deployment', '--deploymentId', info.deploymentId, '--description', 'update ' + today, '--json'], { capture: true });

  console.log('\n更新しました。URLは変わりません：' + webAppUrl(info.deploymentId) + '\n');
}
