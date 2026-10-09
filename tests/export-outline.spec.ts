import { test, expect } from '@playwright/test';
import * as fs from 'fs';
import * as path from 'path';
import { targets, fileUrl, ROOT } from './targets';

/**
 * PowerShell 学習資料の HTML から、pptx 生成用の構造 (JSON) を書き出す。
 *
 * ブラウザ上で DOM を辿るため、HTML を自前でパースする必要がない。
 * 章のアクセント色は getComputedStyle で解決済みの RGB として取り出すので、
 * pptx 側でも HTML と同じ配色を再現できる。
 */

const OUT = path.join(ROOT, 'tmp/outline');

// 対象は targets.ts に集約している。
// README は目次ページなのでスライド化の対象から外す。
const outlineTargets = targets.filter((t) => t.prefix !== 'readme');

for (const t of outlineTargets) {
	test(`構造の書き出し: ${t.file}`, async ({ page }) => {
		fs.mkdirSync(OUT, { recursive: true });

		await page.goto(fileUrl(t));

		const outline = await page.evaluate(() => {
			const norm = (s: string | null) => (s ?? '').replace(/\s+/g, ' ').trim();

			// CSS の色指定 (hsl 等) を rgb(r, g, b) 文字列に解決する
			const toRgb = (value: string) => {
				const probe = document.createElement('span');
				probe.style.color = value.trim();
				document.body.appendChild(probe);
				const c = getComputedStyle(probe).color;
				probe.remove();
				return c;
			};

			// figure は文書順に通し番号を振る（PNG のファイル名と一致させるため）
			const allFigures = Array.from(document.querySelectorAll('figure'));

			// バッジは文言と種類（b-good / b-bad / b-ps7 ...）の両方を返す。
			// pptx 側で同じ配色の図形として描くため。
			const badgeOf = (el: Element) => {
				const b = el.querySelector('.badge');
				if (!b) return null;
				const cls = Array.from(b.classList).find((c) => c.startsWith('b-')) ?? '';
				return { text: norm(b.textContent), cls };
			};

			const readBlocks = (root: Element) => {
				const blocks: any[] = [];

				for (const el of Array.from(root.children)) {
					const tag = el.tagName.toLowerCase();

					if (tag === 'h1') continue; // 章見出しは別途扱う

					if (tag === 'h2' || tag === 'h3') {
						blocks.push({ type: tag, text: norm(el.textContent) });
						continue;
					}

					if (tag === 'p') {
						const badge = badgeOf(el);
						const clone = el.cloneNode(true) as Element;
						clone.querySelectorAll('.badge').forEach((b) => b.remove());
						const text = norm(clone.textContent);
						if (text) blocks.push({ type: 'p', text, badge });
						continue;
					}

					if (tag === 'ul' || tag === 'ol') {
						const items = Array.from(el.querySelectorAll(':scope > li')).map((li) => {
							const badge = badgeOf(li);
							const clone = li.cloneNode(true) as Element;
							clone.querySelectorAll('.badge').forEach((b) => b.remove());
							return { text: norm(clone.textContent), badge };
						});
						if (items.length) blocks.push({ type: 'list', ordered: tag === 'ol', items });
						continue;
					}

					if (tag === 'pre') {
						blocks.push({ type: 'pre', text: (el.textContent ?? '').replace(/\s+$/, '') });
						continue;
					}

					if (tag === 'figure') {
						const idx = allFigures.indexOf(el) + 1;
						const cap = el.querySelector('figcaption');
						blocks.push({ type: 'figure', index: idx, caption: cap ? norm(cap.textContent) : '' });
						continue;
					}

					if (tag === 'div' && el.classList.contains('callout')) {
						const strong = el.querySelector('strong');
						const clone = el.cloneNode(true) as Element;
						const s = clone.querySelector('strong');
						if (s) s.remove();
						blocks.push({
							type: 'callout',
							title: strong ? norm(strong.textContent) : '',
							text: norm(clone.textContent),
						});
						continue;
					}

					if (tag === 'div' && el.classList.contains('table-wrap')) {
						const table = el.querySelector('table');
						if (!table) continue;
						const rows = Array.from(table.querySelectorAll('tr'));
						if (!rows.length) continue;
						const header = Array.from(rows[0].querySelectorAll('th')).map((c) => norm(c.textContent));
						const bodyRows = rows
							.slice(header.length ? 1 : 0)
							.map((r) => Array.from(r.querySelectorAll('td')).map((c) => norm(c.textContent)))
							.filter((r) => r.length > 0);
						blocks.push({ type: 'table', header, rows: bodyRows });
						continue;
					}
				}

				return blocks;
			};

			const sections = Array.from(document.querySelectorAll('section')).map((sec) => {
				const cs = getComputedStyle(sec);
				return {
					id: sec.id,
					chapter: sec.className,
					h1: norm(sec.querySelector('h1')?.textContent ?? ''),
					accent: toRgb(cs.getPropertyValue('--accent')),
					accent2: toRgb(cs.getPropertyValue('--accent2')),
					blocks: readBlocks(sec),
				};
			});

			const tb = document.querySelector('.titlebar');
			const ps = tb ? Array.from(tb.querySelectorAll('p')) : [];

			return {
				title: norm(tb?.querySelector('h1')?.textContent ?? document.title),
				subtitle: ps.length > 0 ? norm(ps[0].textContent) : '',
				meta: ps.length > 1 ? norm(ps[1].textContent) : '',
				figureCount: allFigures.length,
				sections,
			};
		});

		expect(outline.sections.length, '章が1つ以上あること').toBeGreaterThan(0);

		const outFile = path.join(OUT, `${t.prefix}.json`);
		fs.writeFileSync(outFile, JSON.stringify({ file: t.file, prefix: t.prefix, ...outline }, null, '\t'), 'utf8');

		console.log(`${t.file}: ${outline.sections.length} 章 / 図 ${outline.figureCount} 枚 -> ${outFile}`);
	});
}
