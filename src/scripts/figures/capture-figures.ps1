<#
.SYNOPSIS
	docs/*.html に埋め込まれた SVG 図を PNG として書き出す。

.DESCRIPTION
	N:\2026\PlayWright の共有 Playwright 環境を呼び出し、各 HTML の <figure> 内の
	SVG を高 DPI (3x) の PNG として tmp/figures/ に出力する。

	ブラウザで描画してから撮るため、CSS 変数 (--accent 等) が解決済みの状態になり、
	フォントも本文と完全に一致する。出力した PNG は pptx 生成が利用する。

	書き出す対象ファイルの一覧は、共有環境側の targets.ts に集約されている:
	  N:\2026\PlayWright\projects\20260822-powershell-pwsh-learn\targets.ts
	capture-figures / export-outline / export-markdown の3つが同じ一覧を読むため、
	資料を追加したときの追記先はこの1ファイルだけでよい。
#>

$ErrorActionPreference = 'Stop'

$PlaywrightRoot = 'N:\2026\PlayWright'
$ProjectName    = '20260822-powershell-pwsh-learn'
$OutDir         = Join-Path $PSScriptRoot '..\..\..\tmp\figures'

Write-Host '=== SVG 図の書き出し ===' -ForegroundColor Cyan

# 共有環境の存在確認
if (-not (Test-Path -LiteralPath $PlaywrightRoot)) {
	Write-Host "共有 Playwright 環境が見つかりません: $PlaywrightRoot" -ForegroundColor Red
	Write-Host '~/.claude/CLAUDE.md の「PlayWright テストの実行」を確認してください。'
	exit 1
}

$SpecDir = Join-Path $PlaywrightRoot "projects\$ProjectName"
if (-not (Test-Path -LiteralPath $SpecDir)) {
	Write-Host "spec フォルダが見つかりません: $SpecDir" -ForegroundColor Red
	exit 1
}

# 撮り直す前に出力先を空にする。
# 上書きしかしないと、図を削除・改名したときに古い PNG が残り続け、
# build-markdown.ps1 のコピーで docs\images に復活してしまう。
if (Test-Path -LiteralPath $OutDir) {
	$stale = @(Get-ChildItem -LiteralPath $OutDir -Filter '*.png' -File)
	if ($stale.Count -gt 0) {
		$stale | Remove-Item -Force
		Write-Host ("  古い PNG {0} 枚を消しました" -f $stale.Count)
	}
}

# 実行（共有環境のルートで npm を叩く）
Write-Host "共有環境: $PlaywrightRoot"
Write-Host "プロジェクト: $ProjectName"
Write-Host ''

Push-Location -LiteralPath $PlaywrightRoot
try {
	$env:PW_PROJECT = $ProjectName
	npm run test:projects -- $ProjectName
	$exitCode = $LASTEXITCODE
}
finally {
	Pop-Location
}

Write-Host ''

if ($exitCode -ne 0) {
	Write-Host "書き出しに失敗しました (終了コード: $exitCode)" -ForegroundColor Red
	exit $exitCode
}

# 結果の確認
if (-not (Test-Path -LiteralPath $OutDir)) {
	Write-Host "出力フォルダが作成されませんでした: $OutDir" -ForegroundColor Red
	exit 1
}

$files = @(Get-ChildItem -LiteralPath $OutDir -Filter '*.png' -File | Sort-Object Name)

Write-Host "=== 書き出し結果: $($files.Count) 枚 ===" -ForegroundColor Green
Write-Host "出力先: $((Resolve-Path -LiteralPath $OutDir).Path)"
Write-Host ''

Add-Type -AssemblyName System.Drawing
foreach ($f in $files) {
	$img = [System.Drawing.Image]::FromFile($f.FullName)
	Write-Host ('  {0,-18} {1,5} x {2,-5} px   {3,7} KB' -f $f.Name, $img.Width, $img.Height, [math]::Round($f.Length / 1KB, 1))
	$img.Dispose()
}

Write-Host ''
Write-Host '完了しました。' -ForegroundColor Green
