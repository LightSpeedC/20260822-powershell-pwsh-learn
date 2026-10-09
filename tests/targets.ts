/**
 * 処理対象のファイル一覧。
 *
 * capture-figures / export-outline の2つから読み込む。
 * 資料を追加したときの追記箇所をここ1つに集約している。
 */

import * as path from 'path';

// プロジェクトのルート。このファイル（tests/）の 1 つ上。
// ドライブレターを書かないため、置き場から組み立てる。
// Playwright は spec を CommonJS として読むので import.meta.dirname は使えない（SyntaxError になる）。
export const ROOT = path.resolve(__dirname, '..').replace(/\\/g, '/');

export type Target = {
	/** ROOT からの相対パス */
	file: string;
	/** 図の連番や JSON のファイル名に使う接頭辞 */
	prefix: string;
};

export const targets: Target[] = [
	{ file: 'README.html', prefix: 'readme' },
	{ file: 'docs/01-背景.html', prefix: '01' },
	{ file: 'docs/02-ファイル形式.html', prefix: '02' },
	{ file: 'docs/03-実行環境.html', prefix: '03' },
	{ file: 'docs/04-変数と型.html', prefix: '04' },
	{ file: 'docs/05-パイプライン.html', prefix: '05' },
	{ file: 'docs/06-制御構文と関数.html', prefix: '06' },
	{ file: 'docs/07-ファイル操作.html', prefix: '07' },
	{ file: 'docs/08-エラー処理.html', prefix: '08' },
	{ file: 'docs/09-実務パターン.html', prefix: '09' },
	{ file: 'docs/10-ハンズオン.html', prefix: '10' },
	{ file: 'docs/11-Add-Type.html', prefix: '11' },
	{ file: 'docs/12-COMとOffice.html', prefix: '12' },
	{ file: 'docs/A1-逆引き.html', prefix: 'A1' },
	{ file: 'docs/A2-テンプレ集.html', prefix: 'A2' },
	{ file: 'docs/A3-モジュール.html', prefix: 'A3' },
	{ file: 'docs/ZZ-構成.html', prefix: 'ZZ' },
];

/** file:/// 形式の URL に変換する */
export const fileUrl = (t: Target) => encodeURI(`file:///${ROOT}/${t.file}`);
