<#
.SYNOPSIS
	資料の HTML から GitHub 用の Markdown を生成する。

.DESCRIPTION
	HTML が原本で、Markdown はそこから作る派生物。
	GitHub は .html をレンダリングしないため、リポジトリ上で読める形として用意する。

	処理は2段階:
	  1) 共有 Playwright 環境で HTML を読み、Markdown を書き出す
	     - リンクは .html → .md に書き換わる
	     - #chNN のアンカーは見出しのスラッグに差し替わる
	  2) SVG 図の PNG を docs/images/ にコピーする
	     （Markdown からは画像参照になるため、コミット対象に置く必要がある）

	さきに src/scripts/figures/capture-figures.ps1 を実行して PNG を用意しておくこと。
#>

$ErrorActionPreference = 'Stop'

$PlaywrightRoot = 'N:\2026\PlayWright'
$ProjectName    = '20260822-powershell-pwsh-learn'
$Root           = (Resolve-Path (Join-Path $PSScriptRoot '..\..\..')).Path
$FigureDir      = Join-Path $Root 'tmp\figures'
$ImageDir       = Join-Path $Root 'docs\images'

Write-Host '=== Markdown の生成 ===' -ForegroundColor Cyan

if (-not (Test-Path -LiteralPath $PlaywrightRoot)) {
	Write-Host "共有 Playwright 環境が見つかりません: $PlaywrightRoot" -ForegroundColor Red
	exit 1
}

if (-not (Test-Path -LiteralPath $FigureDir)) {
	Write-Host "図の PNG がありません: $FigureDir" -ForegroundColor Red
	Write-Host 'さきに src\scripts\figures\capture-figures.ps1 を実行してください。'
	exit 1
}

# ---- 1) Markdown を書き出す ----
Push-Location -LiteralPath $PlaywrightRoot
try {
	$env:PW_PROJECT = $ProjectName
	npx playwright test --config=playwright.projects.config.ts export-markdown
	$exitCode = $LASTEXITCODE
}
finally {
	Pop-Location
}

if ($exitCode -ne 0) {
	Write-Host "Markdown の生成に失敗しました (終了コード: $exitCode)" -ForegroundColor Red
	exit $exitCode
}

# ---- 2) 図をコミット対象へコピー ----
New-Item -ItemType Directory -Force -Path $ImageDir | Out-Null
Copy-Item (Join-Path $FigureDir '*.png') $ImageDir -Force

$imgs = @(Get-ChildItem -LiteralPath $ImageDir -Filter '*.png' -File)
$mb   = [math]::Round(($imgs | Measure-Object Length -Sum).Sum / 1MB, 1)

Write-Host ''
Write-Host "=== 結果 ===" -ForegroundColor Green
$mds = @(Get-ChildItem -LiteralPath $Root -Filter '*.md' -File) +
       @(Get-ChildItem -LiteralPath (Join-Path $Root 'docs') -Filter '*.md' -File)
foreach ($m in ($mds | Sort-Object Name)) {
	Write-Host ('  {0,-28} {1,5} 行' -f $m.Name, (Get-Content -LiteralPath $m.FullName).Count)
}
Write-Host ''
Write-Host ("  図: {0} 枚 / {1} MB -> docs\images" -f $imgs.Count, $mb)
Write-Host ''
Write-Host '完了しました。' -ForegroundColor Green
