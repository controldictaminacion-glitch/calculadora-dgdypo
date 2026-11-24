param(
  [string]$InputPath = 'tmp/preview-workbook.json',
  [string]$OutputPath = 'output/previos/Excel_previo.png'
)
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Drawing
$sheet = Get-Content -LiteralPath $InputPath -Raw -Encoding UTF8 | ConvertFrom-Json
$scale = 1.5
$gutter = 44 * $scale
$top = 26 * $scale
$x = @(0.0)
$y = @(0.0)
foreach ($column in $sheet.columns) { $x += $x[-1] + [Math]::Floor($column.width * 7 + 5) * $scale }
foreach ($row in $sheet.rows) { $y += $y[-1] + $row.height * 96 / 72 * $scale }
$bitmap = New-Object System.Drawing.Bitmap ([int][Math]::Ceiling($x[-1] + $gutter + 4)), ([int][Math]::Ceiling($y[-1] + $top + 4))
$graphics = [System.Drawing.Graphics]::FromImage($bitmap)
$graphics.Clear([System.Drawing.Color]::White)
$graphics.TextRenderingHint = [System.Drawing.Text.TextRenderingHint]::ClearTypeGridFit
$graphics.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias

function Color-From($Value, $Fallback = '#202733') {
  if ($Value.argb) {
    $hex = $Value.argb
    if ($hex.Length -eq 8) { $hex = $hex.Substring(2) }
    return [System.Drawing.ColorTranslator]::FromHtml('#' + $hex)
  }
  return [System.Drawing.ColorTranslator]::FromHtml($Fallback)
}
function Column-Index($Letters) {
  $number = 0
  foreach ($character in $Letters.ToCharArray()) { $number = $number * 26 + ([int]$character - 64) }
  return $number
}
function Display-Value($Cell) {
  $value = $Cell.value
  if ($null -ne $value.formula) { $value = $value.result }
  if ($Cell.numFmt -match 'dd.*mm.*yyyy') {
    if ($value -is [double] -or $value -is [int] -or $value -is [long] -or $value -is [decimal]) { return [datetime]::FromOADate([double]$value).ToString('dd/MM/yyyy') }
    if ($value -match '^\d{4}-\d{2}-\d{2}') { return ([datetimeoffset]::Parse([string]$value, [Globalization.CultureInfo]::InvariantCulture)).UtcDateTime.ToString('dd/MM/yyyy') }
  }
  if ($value -is [double] -or $value -is [int] -or $value -is [long] -or $value -is [decimal]) {
    if ($Cell.numFmt -match '\$') { return '$' + ([double]$value).ToString('N2', [Globalization.CultureInfo]::GetCultureInfo('es-MX')) }
    return ([double]$value).ToString('N0', [Globalization.CultureInfo]::GetCultureInfo('es-MX'))
  }
  return [string]$Cell.text
}

$headerBrush = New-Object System.Drawing.SolidBrush ([System.Drawing.ColorTranslator]::FromHtml('#F1F3F5'))
$headerPen = New-Object System.Drawing.Pen ([System.Drawing.ColorTranslator]::FromHtml('#D2D7DC')), 1
$headerFont = New-Object System.Drawing.Font 'Arial', (9 * $scale), ([System.Drawing.FontStyle]::Regular), ([System.Drawing.GraphicsUnit]::Pixel)
$headerText = New-Object System.Drawing.SolidBrush ([System.Drawing.ColorTranslator]::FromHtml('#687380'))
$center = New-Object System.Drawing.StringFormat
$center.Alignment = [System.Drawing.StringAlignment]::Center
$center.LineAlignment = [System.Drawing.StringAlignment]::Center
$graphics.FillRectangle($headerBrush, 0, 0, $bitmap.Width, $top)
$graphics.FillRectangle($headerBrush, 0, 0, $gutter, $bitmap.Height)
foreach ($column in $sheet.columns) {
  $rectangle = New-Object System.Drawing.RectangleF ([single]($gutter + $x[$column.number - 1])), 0, ([single]($x[$column.number] - $x[$column.number - 1])), ([single]$top)
  $graphics.DrawRectangle($headerPen, $rectangle.X, $rectangle.Y, $rectangle.Width, $rectangle.Height)
  $graphics.DrawString($column.letter, $headerFont, $headerText, $rectangle, $center)
}
foreach ($row in $sheet.rows) {
  $rectangle = New-Object System.Drawing.RectangleF 0, ([single]($top + $y[$row.number - 1])), ([single]$gutter), ([single]($y[$row.number] - $y[$row.number - 1]))
  $graphics.DrawRectangle($headerPen, $rectangle.X, $rectangle.Y, $rectangle.Width, $rectangle.Height)
  $graphics.DrawString([string]$row.number, $headerFont, $headerText, $rectangle, $center)
}

foreach ($cell in $sheet.cells) {
  if ($cell.isMerged -and $cell.master -ne $cell.address) { continue }
  $lastColumn = $cell.col
  $lastRow = $cell.row
  foreach ($merge in $sheet.merges) {
    if ($merge -match '^([A-Z]+)(\d+):([A-Z]+)(\d+)$') {
      if ((Column-Index $Matches[1]) -eq $cell.col -and [int]$Matches[2] -eq $cell.row) {
        $lastColumn = Column-Index $Matches[3]
        $lastRow = [int]$Matches[4]
        break
      }
    }
  }
  $rectangle = New-Object System.Drawing.RectangleF ([single]($gutter + $x[$cell.col - 1])), ([single]($top + $y[$cell.row - 1])), ([single]($x[$lastColumn] - $x[$cell.col - 1])), ([single]($y[$lastRow] - $y[$cell.row - 1]))
  if ($cell.fill.fgColor -and $cell.fill.pattern -eq 'solid') {
    $brush = New-Object System.Drawing.SolidBrush (Color-From $cell.fill.fgColor '#FFFFFF')
    $graphics.FillRectangle($brush, $rectangle)
    $brush.Dispose()
  }
  foreach ($edge in @('top', 'bottom', 'left', 'right')) {
    $border = $cell.border.$edge
    if ($border.style) {
      $pen = New-Object System.Drawing.Pen (Color-From $border.color '#D9DEE5'), ([single]$scale)
      switch ($edge) {
        'top' { $graphics.DrawLine($pen, $rectangle.Left, $rectangle.Top, $rectangle.Right, $rectangle.Top) }
        'bottom' { $graphics.DrawLine($pen, $rectangle.Left, $rectangle.Bottom, $rectangle.Right, $rectangle.Bottom) }
        'left' { $graphics.DrawLine($pen, $rectangle.Left, $rectangle.Top, $rectangle.Left, $rectangle.Bottom) }
        'right' { $graphics.DrawLine($pen, $rectangle.Right, $rectangle.Top, $rectangle.Right, $rectangle.Bottom) }
      }
      $pen.Dispose()
    }
  }
  $text = Display-Value $cell
  if (-not $text) { continue }
  $fontStyle = [System.Drawing.FontStyle]::Regular
  if ($cell.font.bold) { $fontStyle = $fontStyle -bor [System.Drawing.FontStyle]::Bold }
  if ($cell.font.italic) { $fontStyle = $fontStyle -bor [System.Drawing.FontStyle]::Italic }
  $size = 11
  if ($cell.font.size) { $size = $cell.font.size }
  $font = New-Object System.Drawing.Font 'Arial', ([single]($size * 96 / 72 * $scale)), $fontStyle, ([System.Drawing.GraphicsUnit]::Pixel)
  $brush = New-Object System.Drawing.SolidBrush (Color-From $cell.font.color)
  $format = New-Object System.Drawing.StringFormat
  $format.LineAlignment = [System.Drawing.StringAlignment]::Center
  $format.Trimming = [System.Drawing.StringTrimming]::None
  if (-not $cell.alignment.wrapText) { $format.FormatFlags = [System.Drawing.StringFormatFlags]::NoWrap }
  switch ($cell.alignment.horizontal) {
    'right' { $format.Alignment = [System.Drawing.StringAlignment]::Far }
    'center' { $format.Alignment = [System.Drawing.StringAlignment]::Center }
    default { $format.Alignment = [System.Drawing.StringAlignment]::Near }
  }
  $rectangle.Inflate(-5 * $scale, -2 * $scale)
  $graphics.DrawString($text, $font, $brush, $rectangle, $format)
  $format.Dispose()
  $brush.Dispose()
  $font.Dispose()
}
$bitmap.Save([System.IO.Path]::GetFullPath($OutputPath), [System.Drawing.Imaging.ImageFormat]::Png)
$graphics.Dispose()
$bitmap.Dispose()
Write-Output ([System.IO.Path]::GetFullPath($OutputPath))
