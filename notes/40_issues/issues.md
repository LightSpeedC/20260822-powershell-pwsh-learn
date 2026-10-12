# PowerShell 学習資料 課題

これから解決すること

> 📅 作成: 2026-10-12 / 更新: 2026-10-12

[README](../../README.md)

## 1. 資料の体裁（1 件）

### i261012-01 ✅ **済** コードブロックに言語の指定が無い

`docs/` の HTML にあるコードブロック（`<pre><code>`）283 個に、`class="language-xxx"` が無い。共通ルール「HTMLデザインルール」＞「クラス名」では、コードの言語を正式名で書く。無いと、GitHub 上の Markdown でシンタックスハイライトが効かない。

| やること | 状態 |
|---|---|
| 中身から言語を決めて `class="language-xxx"` を付ける（PowerShell・実行結果は `text`・cmd は `batch` 等）。283 個に付けた（powershell 201・text 67・batch 7・csharp 2・json 2・shell 1・javascript 1・html 1・ini 1） | ✅ **済** |
| html2md で Markdown を作り直し、言語の無いフェンスが 0 件になったことを数える。開きフェンス 279 個のうち 0 個 | ✅ **済** |

[README](../../README.md)
