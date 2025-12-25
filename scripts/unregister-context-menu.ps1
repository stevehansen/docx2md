# Unregister docx2md from Windows Explorer context menu
# Run with same privileges used during registration

param(
    [switch]$AllUsers
)

$ErrorActionPreference = "Stop"

# Determine registry root based on scope
if ($AllUsers) {
    $RegRoot = "HKLM:"
    $ScopeDesc = "all users (requires Administrator)"
} else {
    $RegRoot = "HKCU:"
    $ScopeDesc = "current user"
}

Write-Host "Unregistering context menu for $ScopeDesc..." -ForegroundColor Yellow

# Registry paths to remove
$DocxShellPath = "$RegRoot\Software\Classes\.docx\shell\docx2md"
$SystemShellPath = "$RegRoot\Software\Classes\SystemFileAssociations\.docx\shell\docx2md"

$Removed = $false

try {
    if (Test-Path $DocxShellPath) {
        Remove-Item -Path $DocxShellPath -Recurse -Force
        Write-Host "Removed: $DocxShellPath" -ForegroundColor Green
        $Removed = $true
    }

    if (Test-Path $SystemShellPath) {
        Remove-Item -Path $SystemShellPath -Recurse -Force
        Write-Host "Removed: $SystemShellPath" -ForegroundColor Green
        $Removed = $true
    }

    if ($Removed) {
        Write-Host ""
        Write-Host "=== Unregistration Complete ===" -ForegroundColor Green
        Write-Host "Context menu entry has been removed."
    } else {
        Write-Host "No context menu entries found to remove." -ForegroundColor Yellow
    }

} catch {
    Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Red
    if ($AllUsers) {
        Write-Host "Try running PowerShell as Administrator" -ForegroundColor Yellow
    }
    exit 1
}
