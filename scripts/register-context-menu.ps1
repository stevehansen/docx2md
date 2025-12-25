# Register docx2md in Windows Explorer context menu for .docx files
# Run as Administrator for system-wide registration, or as user for current user only

param(
    [string]$ExePath,
    [switch]$AllUsers
)

$ErrorActionPreference = "Stop"

# Find the executable if not specified
if (-not $ExePath) {
    $ScriptDir = Split-Path -Parent $PSScriptRoot
    $ExePath = Join-Path $ScriptDir "publish\win-x64\Docx2Md.UI.exe"

    if (-not (Test-Path $ExePath)) {
        Write-Host "Error: Could not find Docx2Md.UI.exe" -ForegroundColor Red
        Write-Host "Please specify the path with -ExePath or run publish.ps1 first" -ForegroundColor Yellow
        exit 1
    }
}

$ExePath = (Resolve-Path $ExePath).Path
Write-Host "Using executable: $ExePath" -ForegroundColor Cyan

# Determine registry root based on scope
if ($AllUsers) {
    $RegRoot = "HKLM:"
    $ScopeDesc = "all users (requires Administrator)"
} else {
    $RegRoot = "HKCU:"
    $ScopeDesc = "current user"
}

Write-Host "Registering context menu for $ScopeDesc..." -ForegroundColor Green

# Registry paths
$DocxShellPath = "$RegRoot\Software\Classes\.docx\shell\docx2md"
$DocxCommandPath = "$DocxShellPath\command"

# Also register for the SystemFileAssociations (works even when .docx is associated with Word)
$SystemShellPath = "$RegRoot\Software\Classes\SystemFileAssociations\.docx\shell\docx2md"
$SystemCommandPath = "$SystemShellPath\command"

try {
    # Register under .docx directly
    Write-Host "Creating registry keys..." -ForegroundColor Yellow

    # Create shell key
    if (-not (Test-Path $DocxShellPath)) {
        New-Item -Path $DocxShellPath -Force | Out-Null
    }
    Set-ItemProperty -Path $DocxShellPath -Name "(Default)" -Value "Open with docx2md"
    Set-ItemProperty -Path $DocxShellPath -Name "Icon" -Value "`"$ExePath`",0"

    # Create command key
    if (-not (Test-Path $DocxCommandPath)) {
        New-Item -Path $DocxCommandPath -Force | Out-Null
    }
    Set-ItemProperty -Path $DocxCommandPath -Name "(Default)" -Value "`"$ExePath`" `"%1`""

    # Also register under SystemFileAssociations for better compatibility
    if (-not (Test-Path $SystemShellPath)) {
        New-Item -Path $SystemShellPath -Force | Out-Null
    }
    Set-ItemProperty -Path $SystemShellPath -Name "(Default)" -Value "Open with docx2md"
    Set-ItemProperty -Path $SystemShellPath -Name "Icon" -Value "`"$ExePath`",0"

    if (-not (Test-Path $SystemCommandPath)) {
        New-Item -Path $SystemCommandPath -Force | Out-Null
    }
    Set-ItemProperty -Path $SystemCommandPath -Name "(Default)" -Value "`"$ExePath`" `"%1`""

    Write-Host ""
    Write-Host "=== Registration Complete ===" -ForegroundColor Green
    Write-Host "You can now right-click any .docx file and select 'Open with docx2md'"
    Write-Host ""
    Write-Host "To unregister, run: .\scripts\unregister-context-menu.ps1" -ForegroundColor Cyan

} catch {
    Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Red
    if ($AllUsers) {
        Write-Host "Try running PowerShell as Administrator for system-wide registration" -ForegroundColor Yellow
    }
    exit 1
}
