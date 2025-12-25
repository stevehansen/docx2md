# Publish script for docx2md
# Creates a self-contained single-file executable for Windows x64

param(
    [switch]$Clean,
    [switch]$RegisterContextMenu
)

$ErrorActionPreference = "Stop"
$ProjectRoot = $PSScriptRoot
$PublishDir = Join-Path $ProjectRoot "publish\win-x64"
$UIProject = Join-Path $ProjectRoot "src\Docx2Md.UI\Docx2Md.UI.csproj"

Write-Host "=== docx2md Publish Script ===" -ForegroundColor Cyan

# Clean if requested
if ($Clean -and (Test-Path $PublishDir)) {
    Write-Host "Cleaning publish directory..." -ForegroundColor Yellow
    Remove-Item -Recurse -Force $PublishDir
}

# Build and publish
Write-Host "Publishing self-contained executable..." -ForegroundColor Green
dotnet publish $UIProject `
    -c Release `
    -r win-x64 `
    --self-contained true `
    -p:PublishSingleFile=true `
    -p:PublishReadyToRun=true `
    -p:IncludeNativeLibrariesForSelfExtract=true `
    -p:EnableCompressionInSingleFile=true `
    -o $PublishDir

if ($LASTEXITCODE -ne 0) {
    Write-Host "Publish failed!" -ForegroundColor Red
    exit 1
}

# Get the exe path
$ExePath = Join-Path $PublishDir "Docx2Md.UI.exe"

if (Test-Path $ExePath) {
    $FileInfo = Get-Item $ExePath
    $SizeMB = [math]::Round($FileInfo.Length / 1MB, 2)
    $Version = (Get-Item $ExePath).VersionInfo.FileVersion

    Write-Host ""
    Write-Host "=== Publish Complete ===" -ForegroundColor Green
    Write-Host "Executable: $ExePath"
    Write-Host "Size: $SizeMB MB"
    Write-Host "Version: $Version"
    Write-Host ""

    # Offer to register context menu
    if ($RegisterContextMenu) {
        Write-Host "Registering Windows Explorer context menu..." -ForegroundColor Yellow
        $RegisterScript = Join-Path $ProjectRoot "scripts\register-context-menu.ps1"
        & $RegisterScript -ExePath $ExePath
    }
} else {
    Write-Host "Error: Expected executable not found at $ExePath" -ForegroundColor Red
    exit 1
}
