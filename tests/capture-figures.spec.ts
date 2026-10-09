import { test, expect } from '@playwright/test';
import * as fs from 'fs';
import * as path from 'path';
import { targets, fileUrl, ROOT } from './targets';

/**
 * 資料の HTML から SVG 図を PNG として書き出す。
 *
 * ブラウザで描画してから撮るため、CSS 変数（--accent 等）が解決済みの状態になり、
 * フォントも本文と完全に一致する。pptx と Markdown の両方がこの PNG を使う。
 */

const OUT = path.join(ROOT, 'tmp/figures');

// 高 DPI で撮る（プロジェクタ投影でも粗くならない）
test.use({ deviceScaleFactor: 3 });

for (const t of targets) {
	test(`図の書き出し: ${t.file}`, async ({ page }) => {
		fs.mkdirSync(OUT, { recursive: true });

		await page.goto(fileUrl(t));

		const figures = page.locator('figure');
		const count = await figures.count();

		for (let i = 0; i < count; i++) {
			const svg = figures.nth(i).locator('svg');
			const name = `${t.prefix}-fig-${String(i + 1).padStart(2, '0')}.png`;
			await svg.screenshot({ path: path.join(OUT, name) });
		}

		expect(count, `${t.file} の図の枚数`).toBeGreaterThanOrEqual(0);
		console.log(`${t.file}: ${count} 枚を書き出し`);
	});
}