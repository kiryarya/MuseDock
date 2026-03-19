$ErrorActionPreference = "Stop"

Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName WindowsBase

$root = Split-Path -Parent $PSScriptRoot
$assetsDir = Join-Path $root "src\MuseDock.Desktop\Assets"
$pngPath = Join-Path $assetsDir "AppIcon.png"
$icoPath = Join-Path $assetsDir "AppIcon.ico"

if (-not (Test-Path $assetsDir)) {
    New-Item -ItemType Directory -Path $assetsDir | Out-Null
}

$size = 256
$visual = [System.Windows.Media.DrawingVisual]::new()
$drawingContext = $visual.RenderOpen()

$cardBrush = [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.Color]::FromArgb(255, 250, 250, 252))
$cardBorder = [System.Windows.Media.Pen]::new(
    [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.Color]::FromArgb(255, 228, 232, 238)),
    1.5)
$gradient = [System.Windows.Media.LinearGradientBrush]::new()
$gradient.StartPoint = [System.Windows.Point]::new(0, 0.5)
$gradient.EndPoint = [System.Windows.Point]::new(1, 0.5)
$gradient.GradientStops.Add([System.Windows.Media.GradientStop]::new([System.Windows.Media.Color]::FromRgb(152, 79, 244), 0.0))
$gradient.GradientStops.Add([System.Windows.Media.GradientStop]::new([System.Windows.Media.Color]::FromRgb(54, 106, 246), 1.0))

$cardRect = [System.Windows.Rect]::new(40, 40, 176, 176)
$drawingContext.DrawRoundedRectangle($cardBrush, $cardBorder, $cardRect, 34, 34)

$topGeometry = [System.Windows.Media.StreamGeometry]::new()
$topContext = $topGeometry.Open()
$topContext.BeginFigure([System.Windows.Point]::new(78, 136), $false, $false)
$topContext.LineTo([System.Windows.Point]::new(102, 89), $true, $false)
$topContext.LineTo([System.Windows.Point]::new(128, 117), $true, $false)
$topContext.LineTo([System.Windows.Point]::new(154, 89), $true, $false)
$topContext.LineTo([System.Windows.Point]::new(178, 136), $true, $false)
$topContext.Close()
$topGeometry.Freeze()

$topPen = [System.Windows.Media.Pen]::new($gradient, 22)
$topPen.StartLineCap = [System.Windows.Media.PenLineCap]::Round
$topPen.EndLineCap = [System.Windows.Media.PenLineCap]::Round
$topPen.LineJoin = [System.Windows.Media.PenLineJoin]::Round
$drawingContext.DrawGeometry($null, $topPen, $topGeometry)

$trayRect = [System.Windows.Rect]::new(66, 138, 124, 42)
$drawingContext.DrawRoundedRectangle($gradient, $null, $trayRect, 18, 18)

$notchRect = [System.Windows.Rect]::new(92, 138, 72, 20)
$drawingContext.DrawRoundedRectangle($cardBrush, $null, $notchRect, 12, 12)

$drawingContext.Close()

$bitmap = [System.Windows.Media.Imaging.RenderTargetBitmap]::new(
    $size,
    $size,
    96,
    96,
    [System.Windows.Media.PixelFormats]::Pbgra32)
$bitmap.Render($visual)

$pngEncoder = [System.Windows.Media.Imaging.PngBitmapEncoder]::new()
$pngEncoder.Frames.Add([System.Windows.Media.Imaging.BitmapFrame]::Create($bitmap))
$pngStream = [System.IO.MemoryStream]::new()
$pngEncoder.Save($pngStream)
$pngBytes = $pngStream.ToArray()
[System.IO.File]::WriteAllBytes($pngPath, $pngBytes)

$iconStream = [System.IO.MemoryStream]::new()
$writer = [System.IO.BinaryWriter]::new($iconStream)
$writer.Write([UInt16]0)
$writer.Write([UInt16]1)
$writer.Write([UInt16]1)
$writer.Write([Byte]0)
$writer.Write([Byte]0)
$writer.Write([Byte]0)
$writer.Write([Byte]0)
$writer.Write([UInt16]1)
$writer.Write([UInt16]32)
$writer.Write([UInt32]$pngBytes.Length)
$writer.Write([UInt32]22)
$writer.Write($pngBytes)
$writer.Flush()
[System.IO.File]::WriteAllBytes($icoPath, $iconStream.ToArray())

Write-Host "Generated: $icoPath"
