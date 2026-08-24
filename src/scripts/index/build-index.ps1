<#
.SYNOPSIS
	docs/A1-逆引き.html のコマンド索引を自動生成する。

.DESCRIPTION
	docs/NN-*.html を走査して、登場する Cmdlet と出現章の対応表を作り、
	A1-逆引き.html の下記マーカー間を置き換える。

	    <!-- AUTO:cmdlets -->  ...  <!-- /AUTO:cmdlets -->

	用途の説明は本スクリプト内の $Purpose に持つ。資料に新しい Cmdlet が
	登場すると、表には出るが用途が空欄になり、実行時に警告が出る。

	手書きの表だと章を追加するたびに陳腐化するため、この形にしている。
#>

$ErrorActionPreference = 'Stop'

$Root    = (Resolve-Path (Join-Path $PSScriptRoot '..\..\..')).Path
$DocsDir = Join-Path $Root 'docs'
$Target  = Join-Path $DocsDir 'A1-逆引き.html'

# 資料内で定義している例示用の関数（実在の Cmdlet ではないので除外する）
$Examples = @(
	'Get-FileReport','Get-User','Write-Log','Add-Tag','Get-Report','Get-One','Get-OneFixed',
	'Get-LogSummary','Test-Return','Show-Args','Test-Param','Test-Out','Test-Write','Test-Scope',
	'Set-Counter','Rename-Bulk','Get-Empty','Add-Prefix','Get-Foo','Get-MyFunction'
)

# 用途の一言説明。新しい Cmdlet が資料に入ったらここに追記する
$Purpose = @{
	'Add-Content'              = 'ファイルの末尾に追記する'
	'Convert-Path'             = '相対パスを絶対パスの文字列にする'
	'Copy-Item'                = 'ファイル・フォルダをコピーする'
	'Export-Clixml'            = 'オブジェクトを型ごと保存する（PowerShell 専用形式）'
	'Export-Csv'               = 'CSV に書き出す。5.1 では -NoTypeInformation が要る'
	'Find-Module'              = 'PowerShell Gallery のモジュールを検索する'
	'Format-List'              = '全プロパティを縦に並べて表示する。行き止まり'
	'Format-Table'             = '表形式で表示する。パイプラインの最後にだけ置く'
	'Get-Alias'                = 'エイリアスの一覧・逆引き'
	'Get-ChildItem'            = 'ファイル・フォルダを列挙する'
	'Get-CimInstance'          = 'WMI 情報を取得する（7 では Get-WmiObject の代替）'
	'Get-Command'              = 'コマンドを探す。動詞・名詞・モジュールで絞れる'
	'Get-Content'              = 'ファイルを読む。既定は行の配列、-Raw で 1 文字列'
	'Get-Date'                 = '現在日時の取得と書式化'
	'Get-Error'                = '直前のエラーを全プロパティ展開して表示する'
	'Get-ExecutionPolicy'      = '実行ポリシーを確認する。-List でスコープ別'
	'Get-Help'                 = 'コマンドの使い方。-Examples が最も実用的'
	'Get-Item'                 = '1 つの項目を取得する'
	'Get-Location'             = 'カレントディレクトリを取得する'
	'Get-Member'               = 'オブジェクトの型・プロパティ・メソッドを調べる'
	'Get-Module'               = '読み込み済み・利用可能なモジュールを見る'
	'Get-PSBreakpoint'         = '設定済みのブレークポイントを一覧する'
	'Get-Process'              = '実行中のプロセスを取得する'
	'Get-Random'               = '乱数・ランダムな要素を得る'
	'Get-Service'              = 'Windows サービスを取得する'
	'Get-Verb'                 = '承認された動詞の一覧（100 個）'
	'Get-WmiObject'            = '7 で削除済み。Get-CimInstance に置き換える'
	'Group-Object'             = 'キーでまとめる。Name / Count / Group を返す'
	'Import-Clixml'            = 'Export-Clixml で保存したオブジェクトを復元する'
	'Import-Csv'               = 'CSV を読む。値はすべて文字列になる'
	'Import-Module'            = 'モジュールを明示的に読み込む'
	'Install-Module'           = 'モジュールを導入する。-Scope CurrentUser 推奨'
	'Invoke-Command'           = 'リモートやスクリプトブロックを実行する'
	'Invoke-WebRequest'        = 'HTTP 要求。5.1 では curl がこれのエイリアスだった'
	'Join-Path'                = 'パスを連結する。区切りの重複を吸収する'
	'Measure-Object'           = '件数・合計・平均・最大最小を求める'
	'Move-Item'                = 'ファイル・フォルダを移動する'
	'New-Item'                 = 'ファイル・フォルダを作る。-Force で親ごと'
	'New-ScheduledTaskAction'  = 'タスクスケジューラの実行内容を定義する'
	'New-ScheduledTaskTrigger' = 'タスクスケジューラの起動条件を定義する'
	'Out-File'                 = 'ファイルに書き出す。5.1 の既定は UTF-16LE'
	'Out-Null'                 = '出力を捨てる。ループ内では $null = のほうが速い'
	'Register-ScheduledTask'   = 'タスクを登録する'
	'Remove-Item'              = '削除する。-WhatIf で事前確認できる'
	'Remove-PSBreakpoint'      = 'ブレークポイントを削除する'
	'Rename-Item'              = '名前を変更する。-WhatIf で事前確認できる'
	'Resolve-Path'             = '絶対パスに解決する（存在しないとエラー）'
	'Select-Object'            = 'プロパティを選ぶ。-ExpandProperty で値そのもの'
	'Set-Alias'                = 'エイリアスを定義する'
	'Set-Content'              = 'ファイルに書く。-Encoding を必ず明示する'
	'Set-ExecutionPolicy'      = '実行ポリシーを変更する。-Scope CurrentUser 推奨'
	'Set-PSBreakpoint'         = '行・変数・コマンドにブレークポイントを置く'
	'Set-PSDebug'              = '実行行のトレースを有効にする'
	'Set-StrictMode'           = '未定義変数などを即エラーにする。先頭に書く'
	'Split-Path'               = 'パスを親・末尾に分解する'
	'Start-Sleep'              = '指定秒だけ待つ'
	'Stop-Process'             = 'プロセスを終了する'
	'Stop-Service'             = 'サービスを停止する'
	'Test-Path'                = '存在を確認する。-PathType でファイル／フォルダを区別'
	'Unblock-File'             = 'ダウンロード由来の印（Zone.Identifier）を外す'
	'Update-Help'              = 'ヘルプを取得・更新する'
	'Wait-Debugger'            = 'その場でデバッガに入る'
	'Write-Debug'              = '-Debug 指定時だけ出るメッセージ'
	'Write-Error'              = '非終了エラーを出す。catch には入らない'
	'Write-Host'               = '画面に出す。戻り値には混ざらない'
	'Write-Output'             = '本来の出力。戻り値に混ざる'
	'Write-Progress'           = '進捗を表示する'
	'Write-Verbose'            = '-Verbose 指定時だけ出る。消さずに残せる'
	'Write-Warning'            = '警告を出す'
}

Write-Host '=== コマンド索引の生成 ===' -ForegroundColor Cyan

if (-not (Test-Path -LiteralPath $Target)) {
	Write-Host "対象が見つかりません: $Target" -ForegroundColor Red
	exit 1
}

# ---- 走査 ----
$verbs = (Get-Verb).Verb
$enc   = [System.Text.UTF8Encoding]::new($false)
$found = @{}

foreach ($f in (Get-ChildItem -LiteralPath $DocsDir -Filter '*.html' | Where-Object Name -match '^\d\d-')) {
	$text = ([System.IO.File]::ReadAllText($f.FullName, $enc)) -replace '<[^>]+>', ' '
	$ch   = $f.BaseName.Substring(0, 2)
	foreach ($m in [regex]::Matches($text, '\b([A-Z][a-z]+)-([A-Z][A-Za-z]+)\b')) {
		$name = $m.Groups[1].Value + '-' + $m.Groups[2].Value
		if ($verbs -notcontains $m.Groups[1].Value) { continue }
		if ($Examples -contains $name) { continue }
		if (-not $found.ContainsKey($name)) { $found[$name] = [System.Collections.Generic.HashSet[string]]::new() }
		[void]$found[$name].Add($ch)
	}
}

Write-Host ("  検出した Cmdlet: {0} 種" -f $found.Count)

# 用途が未登録のものを警告する
$missing = @($found.Keys | Where-Object { -not $Purpose.ContainsKey($_) } | Sort-Object)
if ($missing.Count -gt 0) {
	Write-Host ("  用途が未登録: {0} 種 — build-index.ps1 の `$Purpose に追記してください" -f $missing.Count) -ForegroundColor Yellow
	$missing | ForEach-Object { Write-Host "    $_" -ForegroundColor Yellow }
}

# ---- 表を組み立てる ----
$rows = foreach ($k in ($found.Keys | Sort-Object)) {
	$chs = ($found[$k] | Sort-Object) -join ' / '
	$use = if ($Purpose.ContainsKey($k)) { $Purpose[$k] } else { '—' }
	"<tr><td class=`"mono`">$k</td><td>$use</td><td class=`"nowrap`">$chs</td></tr>"
}

$table = @(
	'<div class="table-wrap">'
	'<table>'
	'<tr><th>Cmdlet</th><th>用途</th><th>登場する章</th></tr>'
	$rows
	'</table>'
	'</div>'
) -join "`n"

# ---- マーカー間を置き換える ----
$html = [System.IO.File]::ReadAllText($Target, $enc)
$pattern = '(?s)(<!-- AUTO:cmdlets -->).*?(<!-- /AUTO:cmdlets -->)'
if ($html -notmatch $pattern) {
	Write-Host '  マーカー <!-- AUTO:cmdlets --> が見つかりません' -ForegroundColor Red
	exit 1
}
$html = [regex]::Replace($html, $pattern, { param($m) $m.Groups[1].Value + "`n" + $table + "`n" + $m.Groups[2].Value })
[System.IO.File]::WriteAllText($Target, $html, $enc)

Write-Host ("  {0} 行の表を A1-逆引き.html に埋め込みました" -f $rows.Count) -ForegroundColor Green
Write-Host ''
Write-Host '完了しました。図と pptx・Markdown の再生成は別途行ってください。' -ForegroundColor Green
