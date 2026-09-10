<#
.SYNOPSIS
	各章のフッターの前後リンクを自動生成する。

.DESCRIPTION
	docs/NN-*.html のフッターにある下記マーカー間を書き換える。

	    <!-- AUTO:nav -->  ...  <!-- /AUTO:nav -->

	並び順はファイル名（01〜12）で決まり、章タイトルは各ファイルの
	タイトルバーの h1 から取る。そのため章を追加・改題しても、
	このスクリプトを実行するだけで全章のリンクが揃う。

	付録（A1〜A3）は順に読むものではないため対象外。
	フッターの2行目以降（10 章の付録リストなど）はそのまま残る。

	手書きだと章を1つ足すたびに前後2ファイルを直す必要があり、
	実際にタイトルの食い違いと README リンクの欠落が発生したため自動化した。
#>

$ErrorActionPreference = 'Stop'

$Root    = (Resolve-Path (Join-Path $PSScriptRoot '..\..\..')).Path
$DocsDir = Join-Path $Root 'docs'

Write-Host '=== フッターの前後リンク生成 ===' -ForegroundColor Cyan

$enc   = [System.Text.UTF8Encoding]::new($false)
$files = @(Get-ChildItem -LiteralPath $DocsDir -Filter '*.html' -File |
	Where-Object Name -match '^\d\d-' | Sort-Object Name)

if ($files.Count -eq 0) {
	Write-Host '対象のファイルがありません' -ForegroundColor Red
	exit 1
}

# 先に全ファイルのタイトルを集める
$titles = @{}
foreach ($f in $files) {
	$text = [System.IO.File]::ReadAllText($f.FullName, $enc)
	$m = [regex]::Match($text, '(?s)<div class="titlebar">.*?<h1>(.*?)</h1>')
	if (-not $m.Success) {
		Write-Host ("  タイトルが取れません: {0}" -f $f.Name) -ForegroundColor Red
		exit 1
	}
	$titles[$f.Name] = ($m.Groups[1].Value -replace '<[^>]+>', '').Trim()
}

Write-Host ("  対象: {0} ファイル" -f $files.Count)

$updated = 0
for ($i = 0; $i -lt $files.Count; $i++) {
	$f    = $files[$i]
	$text = [System.IO.File]::ReadAllText($f.FullName, $enc)

	$parts = @()
	if ($i -gt 0) {
		$prev = $files[$i - 1]
		$parts += ('前章: <a href="{0}">{1}</a>' -f $prev.Name, $titles[$prev.Name])
	}
	if ($i -lt $files.Count - 1) {
		$next = $files[$i + 1]
		$parts += ('次章: <a href="{0}">{1}</a>' -f $next.Name, $titles[$next.Name])
	}
	$parts += '<a href="../README.html">目次に戻る</a>'
	$nav = $parts -join ' ／ '

	$pattern = '(?s)(<!-- AUTO:nav -->).*?(<!-- /AUTO:nav -->)'
	if ($text -notmatch $pattern) {
		Write-Host ("  マーカーがありません: {0}" -f $f.Name) -ForegroundColor Yellow
		continue
	}
	$new = [regex]::Replace($text, $pattern, { param($m) $m.Groups[1].Value + $nav + $m.Groups[2].Value })

	# 大小文字だけの差でも書き込むよう -cne で比較する
	if ($new -cne $text) {
		[System.IO.File]::WriteAllText($f.FullName, $new, $enc)
		$updated++
	}
	Write-Host ('  {0,-24} {1}' -f $f.Name, ($nav -replace '<[^>]+>', ''))
}

Write-Host ''
Write-Host ("{0} ファイルを更新しました。" -f $updated) -ForegroundColor Green
Write-Host 'Markdown への反映は html2md --dir docs を実行してください。'
