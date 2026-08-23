---
title: PowerShell 学習資料 - Local Project Rules
date: 2026-08-22
type: local-rule
---

# PowerShell 学習資料作成時の Local Rules

このプロジェクト固有の制作方針。

## ファイル形式

**資料は HTML で直接作成する**（`docs/NN-タイトル.html`）。中間の Markdown は作らない。

- 構成案・検討メモは `docs/plan/`
- ルールは `docs/rules/`

---

## スクリプトの置き場所

**ps1 / cmd スクリプトは `src/scripts/<用途>/` に置く。** ここはコミット対象。

```
src/scripts/
├── figures/    # docs の SVG 図を PNG に書き出す
└── pptx/       # PNG と HTML から pptx を生成する
```

- **用途ごとにフォルダを分ける**。ルート直下にスクリプトを並べない
- ps1 には**必ず同名の cmd ランチャーを添える**（`~/.claude/CLAUDE.md` の規則）
- 出力（PNG・pptx 等のビルド成果物）は `tmp/` に置く。コミットしない
- `etc/` はセッションログと `c.bat` 用。**スクリプトを置かない**（gitignore 対象のため）

**Why:** スクリプトは資料と同じく再現に必要な資産なので、履歴に残す必要がある。
`etc/` は gitignore 対象なので置くと失われる。

---

## 資料を追加したら計画の状態を更新する

**`docs/NN-*.html` を新規作成・完成させたら、同じ作業の中で `docs/plan/構成案.html` の状態を更新する。**

更新する箇所は2つ:

1. 「ファイル一覧」表の該当行のバッジ（`<span class="badge b-none">未着手</span>` → `<span class="badge b-good">作成済</span>`）
2. 「基礎編／実践編の詳細」の該当 h2 見出しの末尾バッジ

**Why:** 計画と実体がずれると、次にどこから再開すればよいか分からなくなる。
状態の更新を別作業にすると必ず忘れる。

---

## 見出しへの番号付け

見出しテキストの先頭に階層番号を付ける：

- ファイルレベル（header）：`01. PowerShell の生まれた経緯`
- セクション h1：`01.1 Windows Command Shell の限界`、`01.2 ...` など

**Why:** 複数ファイルがある場合、参照が明確になる。「01.2 を参照」で誰もが同じセクションを指せる。

---

## SVG 図の追加

**できるだけ SVG 図をいっぱい付ける。**

各ファイルにフロー図・比較図・概念図を複数枚埋め込む。描き方は `docs/rules/html-design.md` および
`~/.claude/CLAUDE.md` の「SVG図」節に従う。

---

## 等幅フォントは BIZ UDゴシックを使う

**`code` / `pre` / SVG 内のコード表記は、バックスラッシュを `¥` で描くフォントを使う。**

```css
font-family: "BIZ UDゴシック", "ＭＳ ゴシック", Consolas, "Cascadia Mono", monospace;
```

- **文字そのものは U+005C（バックスラッシュ）のまま**。表示だけが `¥` になるので、
  コピペすれば正しく `\` として貼り付けられる
- `¥`（U+00A5）を直接書いてはいけない。コピペしたパスが動かなくなる
- pptx 生成側（`build-pptx.ps1` の `$FontCode`）も同じフォントに揃える

実測での比較:

| フォント | `\` の表示 | 等幅 |
|---|---|---|
| Consolas / Cascadia Mono | `\` | ✅ |
| **BIZ UDゴシック** | **`¥`** | ✅ 採用 |
| ＭＳ ゴシック | `¥` | ✅ 見た目が古い |
| MS UI Gothic | `¥` | ❌ プロポーショナル |

**Why:** 日本語 Windows のコンソールは `\` を `¥` で表示する。
資料の表示を実機に合わせたほうが、読者が画面と資料を照合しやすい。

---

## バージョン差はバッジで明示する

**資料は Windows PowerShell 5.1 と PowerShell 7 (pwsh) の両方を対象にする。**
どちらか一方でしか動かない、または挙動が異なる内容には、その場でバッジを付ける。

```css
.b-ps7  { background: linear-gradient(100deg, #1a5fa8, #4d9ae0); }  /* pwsh 7 以降 */
.b-ps51 { background: linear-gradient(100deg, #5b2d82, #9a6bc4); }  /* 5.1 固有 */
```

```html
<span class="badge b-ps7">pwsh 7</span>三項演算子 <code>$a ? $b : $c</code> が使える
<span class="badge b-ps51">5.1</span>BOM がないと日本語が壊れる
```

- **両方で動く内容にはバッジを付けない**。付けるのは差がある箇所だけ
- 良い/悪い/未実施/警告のバッジ（緑・赤・灰・橙）とは**別の色系統（青・紫）**にする。意味の軸が違うことを色で示すため

**Why:** 読者の実行環境が 5.1 か 7 か分からない。コピペして動かなかったとき、
原因がバージョン差だと即座に判別できる必要がある。
