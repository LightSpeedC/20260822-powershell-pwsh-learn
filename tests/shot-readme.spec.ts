import { test } from '@playwright/test';
import * as path from 'path';

/**
 * README の見映えを確認するためのスクリーンショット。
 *
 * HTML（ローカルの原本）と Markdown（GitHub がレンダリングしたもの）を
 * 同じ幅で撮り、並べて比べられるようにする。
 * 図を PNG から SVG に切り替えたため、表示が崩れていないかを見るのが目的。
 */

import { ROOT } from './targets';

const OUT = path.join(ROOT, 'tmp/shot');
const MD_URL =
	'https://github.com/LightSpeedC/20260822-powershell-pwsh-learn/blob/develop/README.md';

test.use({ viewport: { width: 1700, height: 1200 } });

test('README.html を撮る', async ({ page }) => {
	await page.goto('file:///' + ROOT + '/README.html');
	await page.waitForLoadState('networkidle');
	await page.screenshot({ path: path.join(OUT, 'readme-html-full.png'), fullPage: true });

	// 章一覧のカード（span.no を part に変えた箇所）
	const cards = page.locator('ul.chapters');
	await cards.screenshot({ path: path.join(OUT, 'readme-html-cards.png') });

	// 最初の図（インライン SVG のまま表示される）
	const fig = page.locator('figure').first();
	if (await fig.count()) {
		await fig.screenshot({ path: path.join(OUT, 'readme-html-fig.png') });
	}
});

test('README.md を GitHub 上で撮る', async ({ page }) => {
	await page.goto(MD_URL);
	await page.waitForLoadState('networkidle');

	// 本文だけを撮る。ヘッダやサイドバーは対象外
	const article = page.locator('article').first();
	await article.screenshot({ path: path.join(OUT, 'readme-md-full.png') });

	// 章一覧の表
	const tables = article.locator('table');
	if (await tables.count()) {
		await tables.first().screenshot({ path: path.join(OUT, 'readme-md-table.png') });
	}

	// 切り出した SVG が表示されているか
	const img = article.locator('img[src*=".svg"]').first();
	if (await img.count()) {
		await img.screenshot({ path: path.join(OUT, 'readme-md-fig.png') });
	}
});
