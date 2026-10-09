<#
.SYNOPSIS
	docs/*.html に埋め込まれた SVG 図を PNG として書き出す。

.DESCRIPTION
	PlayWright 共有環境のランナー (playwright-test) を呼び出し、各 HTML の <figure> 内の
	SVG を高 DPI (3x) の PNG として tmp/figures/ に出力する。

	ブラウザで描画してから撮るため、CSS 変数 (--accent 等) が解決済みの状態になり、
	フォントも本文と完全に一致する。出力した PNG は pptx 生成が利用する。

	書き出す対象ファイルの一覧は tests/targets.ts に集約されている。
	capture-figures / export-outline の2つが同じ一覧を読むため、
	資料を追加したときの追記先はこの1ファイルだけでよい。
#>

$ErrorActionPreference = 'Stop'

$Root    = (Resolve-Path (Join-Path $PSScriptRoot '../../..')).Path
$TestDir = Join-Path $Root 'tests'
$TmpDir  = Join-Path $Root 'tmp'
$OutDir  = Join-Path $TmpDir 'figures'

Write-Host '=== SVG 図の書き出し ===' -ForegroundColor Cyan

# ランナーの存在確認
if (-not (Get-Command playwright-test -ErrorAction SilentlyContinue)) {
	Write-Host 'playwright-test が見つかりません。' -ForegroundColor Red
	Write-Host '共通ルールの「PlayWright テストの実行」を確認してください。'
	exit 1
}

if (-not (Test-Path -LiteralPath $TestDir)) {
	Write-Host "テストの置き場が見つかりません: $TestDir" -ForegroundColor Red
	exit 1
}

# 撮り直す前に出力先を空にする。
# 上書きしかしないと、図を削除・改名したときに古い PNG が残り続け、
# build-pptx.ps1 が消したはずの図をスライドに載せてしまう。
if (Test-Path -LiteralPath $OutDir) {
	$stale = @(Get-ChildItem -LiteralPath $OutDir -Filter '*.png' -File)
	if ($stale.Count -gt 0) {
		$stale | Remove-Item -Force
		Write-Host ("  古い PNG {0} 枚を消しました" -f $stale.Count)
	}
}

# 実行
# 走らせる spec を名指しする。渡さないと tests 配下が全部動く。
# ブラウザは chromium だけにする。3 つで走ると同じ PNG を 3 回撮り直すことになる。
Write-Host "テスト: $TestDir"
Write-Host ''

playwright-test --test-dir "$TestDir" --out-dir "$TmpDir" --project=chromium capture-figures export-outline
$exitCode = $LASTEXITCODE

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
