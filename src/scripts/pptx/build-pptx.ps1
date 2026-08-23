<#
.SYNOPSIS
	資料の HTML から PowerPoint (pptx) を生成する。

.DESCRIPTION
	tmp/outline/NN.json（構造）と tmp/figures/NN-fig-XX.png（図）を材料に、
	PowerPoint の COM 自動化でスライドを組み立てる。

	テキスト・表・バッジはネイティブのオブジェクトとして作るので、生成後に
	PowerPoint 上でそのまま編集できる。図だけは PNG として貼り込む。

	高さは AutoSize で実測してから配置を決めるため、無駄な「（続き）」が出ない。

	事前に src/scripts/figures/capture-figures.ps1 を実行しておくこと。

.PARAMETER Prefix
	対象のファイル番号。省略時は tmp/outline 配下のすべてを処理する。

.PARAMETER IncludeSupplements
	「PowerShell が初めての人へ」の補足枠もスライドに含める。
	既定では読み物向けの補足とみなして除外し、スライド枚数を抑える。

.EXAMPLE
	.\build-pptx.ps1 -Prefix 01
#>

param(
	[string]$Prefix,
	[switch]$IncludeSupplements,
	[switch]$Pdf,
	[switch]$Images
)

$ErrorActionPreference = 'Stop'

$Root       = (Resolve-Path (Join-Path $PSScriptRoot '..\..\..')).Path
$OutlineDir = Join-Path $Root 'tmp\outline'
$FigureDir  = Join-Path $Root 'tmp\figures'
$PptxDir    = Join-Path $Root 'tmp\pptx'
$SlideImgDir = Join-Path $Root 'tmp\slides'

# ---- スライドの寸法（16:9 / ポイント単位） ----
$SlideW   = 960
$SlideH   = 540
$MarginX  = 48
$BandH    = 70
$BodyTop  = $BandH + 22
$BodyW    = $SlideW - $MarginX * 2
$BodyMaxY = $SlideH - 30

# ---- PowerPoint の定数 ----
$msoTrue        = -1
$msoFalse       = 0
$msoTextHoriz   = 1
$msoShapeRect   = 1
$msoShapeRound  = 5
$ppLayoutBlank  = 12
$ppSaveAsPptx   = 24
$ppSaveAsPdf    = 32
$ppAutoSizeFit  = 1

$FontUI   = 'Yu Gothic UI'
$FontCode = 'BIZ UDゴシック'   # バックスラッシュを ¥ で描く等幅フォント

# ---- 共通色（PowerPoint は BGR 順） ----
$NavyDark  = 0x4D2212   # #12224d
$NavyLight = 0xBF5F2F   # #2f5fbf
$InkColor  = 0x30231C   # #1c2330
$SoftInk   = 0x68554A   # #4a5568
$CodeBg    = 0xFAF6F4   # #f4f6fa
$LineColor = 0xE8DFD9   # #d9dfe8

# HTML のバッジと同じ配色
$BadgeColors = @{
	'b-good' = @{ From = 0x3C7A1D; To = 0x6CB545 }   # #1d7a3c -> #45b56c
	'b-bad'  = @{ From = 0x221FA5; To = 0x5C5AE0 }   # #a51f22 -> #e05a5c
	'b-none' = @{ From = 0x72645A; To = 0xAEA197 }   # #5a6472 -> #97a1ae
	'b-warn' = @{ From = 0x0A62B8; To = 0x48A5F0 }   # #b8620a -> #f0a548
	'b-ps7'  = @{ From = 0xA85F1A; To = 0xE09A4D }   # #1a5fa8 -> #4d9ae0
	'b-ps51' = @{ From = 0x822D5B; To = 0xC46B9A }   # #5b2d82 -> #9a6bc4
}

# 補足枠（読み物向け）と判定する見出し
$SupplementTitle = 'PowerShell が初めての人へ'

function ConvertTo-BgrColor {
	<# "rgb(80, 14, 113)" を PowerPoint の BGR 整数へ #>
	param([string]$Rgb)
	if ($Rgb -match 'rgba?\(\s*(\d+)\s*,\s*(\d+)\s*,\s*(\d+)') {
		return [int]$matches[1] + ([int]$matches[2] -shl 8) + ([int]$matches[3] -shl 16)
	}
	return 0x000000
}

function Get-TextWidth {
	<# 全角=1.0em、半角=0.55em として概算幅（ポイント）を返す #>
	param([string]$Text, [double]$FontSize)
	$w = 0.0
	foreach ($ch in $Text.ToCharArray()) {
		$w += if ([int]$ch -lt 128) { 0.55 } else { 1.0 }
	}
	return $w * $FontSize
}

# =====================================================================
#  部品
# =====================================================================

function New-Slide {
	param($Presentation)
	return $Presentation.Slides.Add($Presentation.Slides.Count + 1, $ppLayoutBlank)
}

function Set-ShapeGradient {
	param($Shape, [int]$ColorFrom, [int]$ColorTo)
	try {
		$Shape.Fill.TwoColorGradient(1, 1)
		$Shape.Fill.ForeColor.RGB = $ColorFrom
		$Shape.Fill.BackColor.RGB = $ColorTo
	} catch {
		$Shape.Fill.Solid()
		$Shape.Fill.ForeColor.RGB = $ColorFrom
	}
	$Shape.Line.Visible = $msoFalse
}

function Add-TitleBand {
	param($Slide, [string]$Text, [int]$ColorFrom, [int]$ColorTo)
	$band = $Slide.Shapes.AddShape($msoShapeRect, 0, 0, $SlideW, $BandH)
	Set-ShapeGradient -Shape $band -ColorFrom $ColorFrom -ColorTo $ColorTo
	$tf = $band.TextFrame
	$tf.MarginLeft = $MarginX; $tf.MarginRight = 24; $tf.MarginTop = 6; $tf.MarginBottom = 6
	$tf.WordWrap = $msoTrue
	$tr = $tf.TextRange
	$tr.Text = $Text
	$tr.Font.Name = $FontUI; $tr.Font.Size = 24; $tr.Font.Bold = $msoTrue
	$tr.Font.Color.RGB = 0xFFFFFF
	$tr.ParagraphFormat.Alignment = 1
	$tf.VerticalAnchor = 3
	return $band
}

function New-ContentSlide {
	param($Presentation, [string]$Title, [int]$Accent, [int]$Accent2)
	$s = New-Slide -Presentation $Presentation
	Add-TitleBand -Slide $s -Text $Title -ColorFrom $Accent -ColorTo $Accent2 | Out-Null
	return $s
}

function Add-BodyText {
	<# 本文のテキストボックス。AutoSize で実測した高さを返す #>
	param(
		$Slide, [string]$Text, [double]$Y,
		[int]$Size = 15, [int]$Color = $InkColor,
		[bool]$Bold = $false, [bool]$Bullet = $false, [double]$IndentX = 0
	)
	$tb = $Slide.Shapes.AddTextbox($msoTextHoriz, $MarginX + $IndentX, $Y, $BodyW - $IndentX, 24)
	$tf = $tb.TextFrame
	$tf.WordWrap = $msoTrue
	$tf.MarginLeft = 0; $tf.MarginRight = 0; $tf.MarginTop = 0; $tf.MarginBottom = 0
	$tr = $tf.TextRange
	$tr.Text = $Text
	$tr.Font.Name = $FontUI
	$tr.Font.Size = $Size
	$tr.Font.Bold = $(if ($Bold) { $msoTrue } else { $msoFalse })
	$tr.Font.Color.RGB = $Color
	$tr.ParagraphFormat.SpaceWithin = 1.05
	if ($Bullet) {
		$tr.ParagraphFormat.Bullet.Visible = $msoTrue
		$tr.ParagraphFormat.Bullet.Character = 8226
	}
	$tf.AutoSize = $ppAutoSizeFit
	# AutoSize は図形の中心を保って伸縮するため Top がずれる。指定位置へ戻す。
	$tb.Top = $Y
	return @{ Shapes = @($tb); Height = [double]$tb.Height }
}

function Add-BadgedText {
	<# 色付きバッジ図形 ＋ 本文。HTML の .badge と同じ配色にする #>
	param(
		$Slide, [string]$BadgeText, [string]$BadgeClass, [string]$Text,
		[double]$Y, [int]$Size = 15, [double]$IndentX = 0
	)
	$c = $BadgeColors[$BadgeClass]
	if (-not $c) { $c = $BadgeColors['b-none'] }

	$bSize = 11
	$bw = [Math]::Max(44, (Get-TextWidth -Text $BadgeText -FontSize $bSize) + 18)
	$bh = 19

	$badge = $Slide.Shapes.AddShape($msoShapeRound, $MarginX + $IndentX, $Y + 2, $bw, $bh)
	Set-ShapeGradient -Shape $badge -ColorFrom $c.From -ColorTo $c.To
	try { $badge.Adjustments.Item(1) = 0.22 } catch {}
	$btf = $badge.TextFrame
	$btf.MarginLeft = 2; $btf.MarginRight = 2; $btf.MarginTop = 0; $btf.MarginBottom = 0
	$btf.WordWrap = $msoFalse
	$btr = $btf.TextRange
	$btr.Text = $BadgeText
	$btr.Font.Name = $FontUI; $btr.Font.Size = $bSize; $btr.Font.Bold = $msoTrue
	$btr.Font.Color.RGB = 0xFFFFFF
	$btr.ParagraphFormat.Alignment = 2

	$textX = $MarginX + $IndentX + $bw + 10
	$tb = $Slide.Shapes.AddTextbox($msoTextHoriz, $textX, $Y, $SlideW - $MarginX - $textX, 24)
	$tf = $tb.TextFrame
	$tf.WordWrap = $msoTrue
	$tf.MarginLeft = 0; $tf.MarginRight = 0; $tf.MarginTop = 0; $tf.MarginBottom = 0
	$tr = $tf.TextRange
	$tr.Text = $Text
	$tr.Font.Name = $FontUI; $tr.Font.Size = $Size
	$tr.Font.Color.RGB = $InkColor
	$tr.ParagraphFormat.SpaceWithin = 1.05
	$tf.AutoSize = $ppAutoSizeFit
	$tb.Top = $Y

	$h = [Math]::Max([double]$tb.Height, $bh + 4)
	return @{ Shapes = @($tb, $badge); Height = $h }
}

function Add-CodeBox {
	param($Slide, [string]$Text, [double]$Y, [int]$Accent)
	$lines = ($Text -split "`n").Count
	$h = $lines * 16.5 + 16

	$box = $Slide.Shapes.AddShape($msoShapeRect, $MarginX, $Y, $BodyW, $h)
	$box.Fill.Solid(); $box.Fill.ForeColor.RGB = $CodeBg
	$box.Line.Visible = $msoFalse
	$tf = $box.TextFrame
	$tf.MarginLeft = 16; $tf.MarginRight = 8; $tf.MarginTop = 6; $tf.MarginBottom = 6
	$tf.WordWrap = $msoFalse
	$tr = $tf.TextRange
	$tr.Text = ($Text -replace "`n", "`r")
	$tr.Font.Name = $FontCode; $tr.Font.Size = 12
	$tr.Font.Color.RGB = 0x462B1B
	$tr.ParagraphFormat.Alignment = 1
	$tf.VerticalAnchor = 1

	$bar = $Slide.Shapes.AddShape($msoShapeRect, $MarginX, $Y, 5, $h)
	$bar.Fill.Solid(); $bar.Fill.ForeColor.RGB = $Accent
	$bar.Line.Visible = $msoFalse

	return @{ Shapes = @($box, $bar); Height = $h }
}

function Add-CalloutBox {
	param($Slide, [string]$Title, [string]$Text, [double]$Y, [int]$Accent)
	$full = if ($Title) { "$Title`r$Text" } else { $Text }

	$box = $Slide.Shapes.AddShape($msoShapeRect, $MarginX, $Y, $BodyW, 40)
	$box.Fill.Solid(); $box.Fill.ForeColor.RGB = 0xFAF7F4
	$box.Line.ForeColor.RGB = $LineColor; $box.Line.Weight = 1
	$tf = $box.TextFrame
	$tf.MarginLeft = 20; $tf.MarginRight = 14; $tf.MarginTop = 7; $tf.MarginBottom = 7
	$tf.WordWrap = $msoTrue
	$tr = $tf.TextRange
	$tr.Text = $full
	$tr.Font.Name = $FontUI; $tr.Font.Size = 14
	$tr.Font.Color.RGB = $InkColor
	$tr.ParagraphFormat.SpaceWithin = 1.05
	$tr.ParagraphFormat.Alignment = 1   # AutoShape は既定が中央揃えなので左に戻す
	$tf.VerticalAnchor = 1
	if ($Title) {
		$p1 = $tr.Paragraphs(1)
		$p1.Font.Bold = $msoTrue
		$p1.Font.Color.RGB = $Accent
	}
	$tf.AutoSize = $ppAutoSizeFit
	# AutoSize は中心を保って伸縮するので Top がずれる。バーを作る前に戻す。
	$box.Top = $Y
	$h = [double]$box.Height

	$bar = $Slide.Shapes.AddShape($msoShapeRect, $MarginX, $Y, 6, $h)
	$bar.Fill.Solid(); $bar.Fill.ForeColor.RGB = $Accent
	$bar.Line.Visible = $msoFalse

	return @{ Shapes = @($box, $bar); Height = $h }
}

function Add-FigureSlide {
	param($Presentation, [string]$ImagePath, [string]$Caption, [string]$SectionTitle, [int]$Accent, [int]$Accent2)
	$slide = New-ContentSlide -Presentation $Presentation -Title $SectionTitle -Accent $Accent -Accent2 $Accent2

	Add-Type -AssemblyName System.Drawing
	$img = [System.Drawing.Image]::FromFile($ImagePath)
	$iw = $img.Width; $ih = $img.Height
	$img.Dispose()

	$availH = $BodyMaxY - $BodyTop - 30
	$scale = [Math]::Min($BodyW / $iw, $availH / $ih)
	$w = $iw * $scale; $h = $ih * $scale

	$slide.Shapes.AddPicture($ImagePath, $msoFalse, $msoTrue, ($SlideW - $w) / 2, $BodyTop, $w, $h) | Out-Null

	if ($Caption) {
		$tb = $slide.Shapes.AddTextbox($msoTextHoriz, $MarginX, $BodyTop + $h + 6, $BodyW, 22)
		$tr = $tb.TextFrame.TextRange
		$tr.Text = $Caption
		$tr.Font.Name = $FontUI; $tr.Font.Size = 12
		$tr.Font.Color.RGB = $SoftInk
		$tr.ParagraphFormat.Alignment = 2
	}
	return $slide
}

function Add-TableSlide {
	param($Presentation, $Block, [string]$SectionTitle, [int]$Accent, [int]$Accent2)
	$slide = New-ContentSlide -Presentation $Presentation -Title $SectionTitle -Accent $Accent -Accent2 $Accent2

	$hasHeader = $Block.header -and $Block.header.Count -gt 0
	$cols = if ($hasHeader) { $Block.header.Count } else { $Block.rows[0].Count }
	$rows = $Block.rows.Count + $(if ($hasHeader) { 1 } else { 0 })

	$availH = $BodyMaxY - $BodyTop
	$tblH = [Math]::Min($availH, $rows * 32)

	$tbl = $slide.Shapes.AddTable($rows, $cols, $MarginX, $BodyTop, $BodyW, $tblH).Table

	$r = 1
	if ($hasHeader) {
		for ($c = 1; $c -le $cols; $c++) {
			$cell = $tbl.Cell($r, $c)
			$cell.Shape.Fill.Solid(); $cell.Shape.Fill.ForeColor.RGB = $Accent
			$tr = $cell.Shape.TextFrame.TextRange
			$tr.Text = [string]$Block.header[$c - 1]
			$tr.Font.Name = $FontUI; $tr.Font.Size = 12.5; $tr.Font.Bold = $msoTrue
			$tr.Font.Color.RGB = 0xFFFFFF
		}
		$r++
	}
	foreach ($row in $Block.rows) {
		for ($c = 1; $c -le $cols; $c++) {
			$cell = $tbl.Cell($r, $c)
			$cell.Shape.Fill.Solid()
			$cell.Shape.Fill.ForeColor.RGB = $(if ($r % 2 -eq 0) { 0xFFFFFF } else { 0xFEFBFA })
			$tr = $cell.Shape.TextFrame.TextRange
			$tr.Text = if ($c -le $row.Count) { [string]$row[$c - 1] } else { '' }
			$tr.Font.Name = $FontUI; $tr.Font.Size = 11.5
			$tr.Font.Color.RGB = $InkColor
		}
		$r++
	}
	return $slide
}

# =====================================================================
#  本体
# =====================================================================

function Build-Presentation {
	param($App, [string]$JsonPath, [bool]$WithSupplements)

	$data = Get-Content -LiteralPath $JsonPath -Encoding UTF8 -Raw | ConvertFrom-Json
	$pres = $App.Presentations.Add($msoFalse)
	$pres.PageSetup.SlideWidth = $SlideW
	$pres.PageSetup.SlideHeight = $SlideH

	$skipped = 0

	# ---- 表紙 ----
	$slide = New-Slide -Presentation $pres
	$bg = $slide.Shapes.AddShape($msoShapeRect, 0, 0, $SlideW, $SlideH)
	Set-ShapeGradient -Shape $bg -ColorFrom $NavyDark -ColorTo $NavyLight

	$tb = $slide.Shapes.AddTextbox($msoTextHoriz, 72, 180, $SlideW - 144, 90)
	$tr = $tb.TextFrame.TextRange
	$tr.Text = $data.title
	$tr.Font.Name = $FontUI; $tr.Font.Size = 40; $tr.Font.Bold = $msoTrue
	$tr.Font.Color.RGB = 0xFFFFFF

	if ($data.subtitle) {
		$tb = $slide.Shapes.AddTextbox($msoTextHoriz, 72, 278, $SlideW - 144, 50)
		$tr = $tb.TextFrame.TextRange
		$tr.Text = $data.subtitle
		$tr.Font.Name = $FontUI; $tr.Font.Size = 17
		$tr.Font.Color.RGB = 0xFFE4D8
	}
	if ($data.meta) {
		$tb = $slide.Shapes.AddTextbox($msoTextHoriz, 72, 400, $SlideW - 144, 30)
		$tr = $tb.TextFrame.TextRange
		$tr.Text = $data.meta
		$tr.Font.Name = $FontUI; $tr.Font.Size = 13
		$tr.Font.Color.RGB = 0xF5CDB9
	}

	foreach ($sec in $data.sections) {
		$accent  = ConvertTo-BgrColor -Rgb $sec.accent
		$accent2 = ConvertTo-BgrColor -Rgb $sec.accent2

		# ---- 章の扉 ----
		$slide = New-Slide -Presentation $pres
		$band = $slide.Shapes.AddShape($msoShapeRect, 0, 190, $SlideW, 160)
		Set-ShapeGradient -Shape $band -ColorFrom $accent -ColorTo $accent2
		$tf = $band.TextFrame
		$tf.MarginLeft = 72; $tf.MarginRight = 48
		$tr = $tf.TextRange
		$tr.Text = $sec.h1
		$tr.Font.Name = $FontUI; $tr.Font.Size = 34; $tr.Font.Bold = $msoTrue
		$tr.Font.Color.RGB = 0xFFFFFF
		$tr.ParagraphFormat.Alignment = 1
		$tf.VerticalAnchor = 3

		# ---- 本文 ----
		# PowerShell の入れ子スコープでは親の変数へ代入できないので、状態は参照型で持つ
		# PendingH3 は「直前に置いた h3」。次のブロックが次スライドへ送られるとき、
		# 見出しだけが取り残されないように一緒に連れて行く。
		$st = @{ Slide = $null; Y = $BodyTop; Title = $sec.h1; PendingH3 = $null }

		foreach ($b in $sec.blocks) {

			if ($b.type -eq 'h2') {
				$st.Title = $b.text
				$st.Slide = New-ContentSlide -Presentation $pres -Title $st.Title -Accent $accent -Accent2 $accent2
				$st.Y = $BodyTop
				$st.PendingH3 = $null
				continue
			}

			if ($b.type -eq 'figure') {
				$png = Join-Path $FigureDir ('{0}-fig-{1:D2}.png' -f $data.prefix, $b.index)
				if (Test-Path -LiteralPath $png) {
					Add-FigureSlide -Presentation $pres -ImagePath $png -Caption $b.caption `
						-SectionTitle $st.Title -Accent $accent -Accent2 $accent2 | Out-Null
					$st.Slide = $null
				}
				continue
			}

			if ($b.type -eq 'table') {
				Add-TableSlide -Presentation $pres -Block $b -SectionTitle $st.Title `
					-Accent $accent -Accent2 $accent2 | Out-Null
				$st.Slide = $null
				continue
			}

			# 読み物向けの補足枠はスライドに載せない（枚数を抑える）
			if ($b.type -eq 'callout' -and -not $WithSupplements -and $b.title -like "$SupplementTitle*") {
				$skipped++
				continue
			}

			if (-not $st.Slide) {
				$st.Slide = New-ContentSlide -Presentation $pres -Title $st.Title -Accent $accent -Accent2 $accent2
				$st.Y = $BodyTop
			}

			# 「置いてから実測し、はみ出したら次スライドへ作り直す」
			$place = {
				param($sl, $yy)
				switch ($b.type) {
					'h3'      { Add-BodyText -Slide $sl -Text $b.text -Y $yy -Size 17 -Color $accent -Bold $true }
					'pre'     { Add-CodeBox  -Slide $sl -Text $b.text -Y $yy -Accent $accent2 }
					'callout' { Add-CalloutBox -Slide $sl -Title $b.title -Text $b.text -Y $yy -Accent $accent }
					'p' {
						if ($b.badge) {
							Add-BadgedText -Slide $sl -BadgeText $b.badge.text -BadgeClass $b.badge.cls -Text $b.text -Y $yy
						} else {
							Add-BodyText -Slide $sl -Text $b.text -Y $yy
						}
					}
				}
			}

			if ($b.type -eq 'list') {
				foreach ($it in $b.items) {
					$placeItem = {
						param($sl, $yy)
						if ($it.badge) {
							Add-BadgedText -Slide $sl -BadgeText $it.badge.text -BadgeClass $it.badge.cls -Text $it.text -Y $yy -IndentX 16
						} else {
							Add-BodyText -Slide $sl -Text $it.text -Y $yy -Bullet $true -IndentX 16
						}
					}
					$res = & $placeItem $st.Slide $st.Y
					if (($st.Y + $res.Height) -gt $BodyMaxY -and $st.Y -gt $BodyTop) {
						foreach ($s in $res.Shapes) { $s.Delete() }

						$carry = $null
						if ($st.PendingH3) {
							foreach ($s in $st.PendingH3.Shapes) { $s.Delete() }
							$carry = $st.PendingH3.Text
							$st.PendingH3 = $null
						}

						$st.Slide = New-ContentSlide -Presentation $pres -Title ($st.Title + '（続き）') -Accent $accent -Accent2 $accent2
						$st.Y = $BodyTop

						if ($carry) {
							$h3res = Add-BodyText -Slide $st.Slide -Text $carry -Y $st.Y -Size 17 -Color $accent -Bold $true
							$st.Y = $st.Y + $h3res.Height + 4
						}
						$res = & $placeItem $st.Slide $st.Y
					}
					$st.PendingH3 = $null
					$st.Y = $st.Y + $res.Height + 5
				}
				continue
			}

			$res = & $place $st.Slide $st.Y
			if (-not $res) { continue }
			if (($st.Y + $res.Height) -gt $BodyMaxY -and $st.Y -gt $BodyTop) {
				foreach ($s in $res.Shapes) { $s.Delete() }

				# 直前の h3 が孤立するので一緒に次スライドへ移す
				$carry = $null
				if ($st.PendingH3) {
					foreach ($s in $st.PendingH3.Shapes) { $s.Delete() }
					$carry = $st.PendingH3.Text
					$st.PendingH3 = $null
				}

				$st.Slide = New-ContentSlide -Presentation $pres -Title ($st.Title + '（続き）') -Accent $accent -Accent2 $accent2
				$st.Y = $BodyTop

				if ($carry) {
					$h3res = Add-BodyText -Slide $st.Slide -Text $carry -Y $st.Y -Size 17 -Color $accent -Bold $true
					$st.PendingH3 = @{ Shapes = $h3res.Shapes; Text = $carry }
					$st.Y = $st.Y + $h3res.Height + 4
				}
				$res = & $place $st.Slide $st.Y
			}

			if ($b.type -eq 'h3') {
				$st.PendingH3 = @{ Shapes = $res.Shapes; Text = $b.text }
			} else {
				$st.PendingH3 = $null
			}

			$gap = if ($b.type -eq 'h3') { 4 } elseif ($b.type -eq 'pre' -or $b.type -eq 'callout') { 10 } else { 7 }
			$st.Y = $st.Y + $res.Height + $gap
		}
	}

	return @{ Presentation = $pres; SlideCount = $pres.Slides.Count; Skipped = $skipped }
}

# ---------------------------------------------------------------------

Write-Host '=== pptx の生成 ===' -ForegroundColor Cyan

if (-not (Test-Path -LiteralPath $OutlineDir)) {
	Write-Host "構造 JSON がありません: $OutlineDir" -ForegroundColor Red
	Write-Host 'さきに src\scripts\figures\capture-figures.ps1 を実行してください。'
	exit 1
}

$jsonFiles = @(Get-ChildItem -LiteralPath $OutlineDir -Filter '*.json' -File | Sort-Object Name)
if ($Prefix) {
	$jsonFiles = @($jsonFiles | Where-Object { $_.BaseName -eq $Prefix })
	if ($jsonFiles.Count -eq 0) {
		Write-Host "指定された番号の JSON がありません: $Prefix" -ForegroundColor Red
		exit 1
	}
}

New-Item -ItemType Directory -Force -Path $PptxDir | Out-Null

if ($IncludeSupplements) {
	Write-Host '補足枠（PowerShell が初めての人へ）も含めます。'
} else {
	Write-Host '補足枠（PowerShell が初めての人へ）は除外します。含めるには -IncludeSupplements を付けてください。'
}

$app = $null
try {
	$app = New-Object -ComObject PowerPoint.Application
	foreach ($jf in $jsonFiles) {
		Write-Host ''
		Write-Host ("[{0}] 生成中..." -f $jf.BaseName)

		$data = Get-Content -LiteralPath $jf.FullName -Encoding UTF8 -Raw | ConvertFrom-Json
		$result = Build-Presentation -App $app -JsonPath $jf.FullName -WithSupplements ([bool]$IncludeSupplements)

		$name = [System.IO.Path]::GetFileNameWithoutExtension($data.file) + '.pptx'
		$outPath = Join-Path $PptxDir $name

		# PowerPoint の SaveAs は既存ファイルを上書きせず「名前 2.pptx」を作る。
		# 再生成のたびに増えてしまうので、先に消しておく。
		if (Test-Path -LiteralPath $outPath) { Remove-Item -LiteralPath $outPath -Force }

		$result.Presentation.SaveAs($outPath, $ppSaveAsPptx)

		$kb = [math]::Round((Get-Item -LiteralPath $outPath).Length / 1KB, 1)
		Write-Host ("  {0}  スライド {1} 枚  {2} KB  （補足枠 {3} 個を除外）" -f $name, $result.SlideCount, $kb, $result.Skipped) -ForegroundColor Green

		if ($Pdf) {
			# ExportAsFixedFormat は PowerShell の COM バインドで引数が解決できないため
			# SaveAs の PDF 形式（ppSaveAsPDF = 32）を使う
			$pdfPath = [System.IO.Path]::ChangeExtension($outPath, '.pdf')
			if (Test-Path -LiteralPath $pdfPath) { Remove-Item -LiteralPath $pdfPath -Force }
			$result.Presentation.SaveAs($pdfPath, $ppSaveAsPdf)
			$pkb = [math]::Round((Get-Item -LiteralPath $pdfPath).Length / 1KB, 1)
			Write-Host ("  {0}  {1} KB" -f (Split-Path $pdfPath -Leaf), $pkb) -ForegroundColor Green
		}

		if ($Images) {
			$imgDir = Join-Path $SlideImgDir $jf.BaseName
			New-Item -ItemType Directory -Force -Path $imgDir | Out-Null
			$result.Presentation.Export($imgDir, 'PNG', 1600, 900)
			$n = @(Get-ChildItem -LiteralPath $imgDir -Filter '*.PNG' -File).Count
			Write-Host ("  スライド画像 {0} 枚 -> {1}" -f $n, $imgDir) -ForegroundColor Green
		}

		$result.Presentation.Close()
	}
}
finally {
	if ($app) { try { $app.Quit() } catch {} }
}

Write-Host ''
Write-Host "出力先: $PptxDir"
Write-Host '完了しました。' -ForegroundColor Green
