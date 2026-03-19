Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName WindowsBase

if ($args.Count -lt 2) {
    throw "Usage: run.ps1 <contextPath> <grayscale|resize-half>"
}

$ContextPath = [string]$args[0]
$Mode = [string]$args[1]

if ($Mode -notin @("grayscale", "resize-half")) {
    throw "Unsupported mode: $Mode"
}

function Get-Encoder($extension) {
    switch ($extension.ToLowerInvariant()) {
        ".png" { return [System.Windows.Media.Imaging.PngBitmapEncoder]::new() }
        ".jpg" { $encoder = [System.Windows.Media.Imaging.JpegBitmapEncoder]::new(); $encoder.QualityLevel = 92; return $encoder }
        ".jpeg" { $encoder = [System.Windows.Media.Imaging.JpegBitmapEncoder]::new(); $encoder.QualityLevel = 92; return $encoder }
        ".bmp" { return [System.Windows.Media.Imaging.BmpBitmapEncoder]::new() }
        ".gif" { return [System.Windows.Media.Imaging.GifBitmapEncoder]::new() }
        ".tif" { return [System.Windows.Media.Imaging.TiffBitmapEncoder]::new() }
        ".tiff" { return [System.Windows.Media.Imaging.TiffBitmapEncoder]::new() }
        default { throw "Unsupported image format: $extension" }
    }
}

function Get-DestinationPath($sourcePath, $mode) {
    $directory = [System.IO.Path]::GetDirectoryName($sourcePath)
    $name = [System.IO.Path]::GetFileNameWithoutExtension($sourcePath)
    $extension = [System.IO.Path]::GetExtension($sourcePath)
    $suffix = if ($mode -eq "grayscale") { "-gray" } else { "-half" }
    return Join-Path $directory ($name + $suffix + $extension)
}

function Load-Frame($sourcePath) {
    $stream = [System.IO.File]::Open($sourcePath, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::ReadWrite)
    try {
        $decoder = [System.Windows.Media.Imaging.BitmapDecoder]::Create(
            $stream,
            [System.Windows.Media.Imaging.BitmapCreateOptions]::PreservePixelFormat,
            [System.Windows.Media.Imaging.BitmapCacheOption]::OnLoad)
        return $decoder.Frames[0]
    }
    finally {
        $stream.Dispose()
    }
}

function Convert-ToGray($frame) {
    $converted = [System.Windows.Media.Imaging.FormatConvertedBitmap]::new()
    $converted.BeginInit()
    $converted.Source = $frame
    $converted.DestinationFormat = [System.Windows.Media.PixelFormats]::Gray8
    $converted.EndInit()
    $converted.Freeze()
    return $converted
}

function Resize-Half($frame) {
    $scale = [System.Windows.Media.ScaleTransform]::new(0.5, 0.5)
    $bitmap = [System.Windows.Media.Imaging.TransformedBitmap]::new($frame, $scale)
    $bitmap.Freeze()
    return $bitmap
}

$context = Get-Content -LiteralPath $ContextPath -Raw | ConvertFrom-Json
if (-not $context.selectedItems -or $context.selectedItems.Count -eq 0) {
    throw "No selected file."
}

$target = $context.selectedItems[0]
if ($target.isDirectory) {
    throw "Directories are not supported."
}

$sourcePath = [string]$target.filePath
if (-not (Test-Path -LiteralPath $sourcePath)) {
    throw "File not found: $sourcePath"
}

$extension = [System.IO.Path]::GetExtension($sourcePath)
$frame = Load-Frame $sourcePath
$outputBitmap = if ($Mode -eq "grayscale") { Convert-ToGray $frame } else { Resize-Half $frame }
$destinationPath = Get-DestinationPath $sourcePath $Mode
$encoder = Get-Encoder $extension
$encoder.Frames.Add([System.Windows.Media.Imaging.BitmapFrame]::Create($outputBitmap))

$output = [System.IO.File]::Create($destinationPath)
try {
    $encoder.Save($output)
}
finally {
    $output.Dispose()
}

Write-Host "MuseDock image plugin created: $destinationPath"
