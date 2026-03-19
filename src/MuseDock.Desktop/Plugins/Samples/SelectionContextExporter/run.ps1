param(
    [Parameter(Mandatory = $true)]
    [string]$ContextPath
)

$outputDir = Join-Path $env:TEMP "MuseDock\SamplePlugin"
New-Item -ItemType Directory -Force -Path $outputDir | Out-Null

$timestamp = Get-Date -Format "yyyyMMdd-HHmmss"
$outputPath = Join-Path $outputDir "selection-context-$timestamp.json"

Copy-Item -LiteralPath $ContextPath -Destination $outputPath -Force
Write-Host "MuseDock sample plugin exported context to $outputPath"
