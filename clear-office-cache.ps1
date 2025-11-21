# Clear Office Add-in Cache Script
# This script clears the Office Web Extension Framework (WEF) cache
# Run this script as Administrator for best results

param(
    [switch]$Force = $false
)

Write-Host "Office Add-in Cache Clearing Script" -ForegroundColor Green
Write-Host "===================================" -ForegroundColor Green

# Define cache paths
$wefPath = "$env:LOCALAPPDATA\Microsoft\Office\16.0\Wef"
$tempPath = "$env:TEMP\OfficeAddins*"

# Check if Outlook is running
$outlookProcesses = Get-Process -Name "OUTLOOK" -ErrorAction SilentlyContinue

if ($outlookProcesses) {
    Write-Host "WARNING: Outlook is currently running!" -ForegroundColor Yellow
    Write-Host "For best results, close Outlook before clearing cache." -ForegroundColor Yellow
    
    if (-not $Force) {
        $response = Read-Host "Do you want to continue anyway? (y/n)"
        if ($response -ne 'y' -and $response -ne 'Y') {
            Write-Host "Script cancelled by user." -ForegroundColor Red
            exit 1
        }
    }
}

# Clear WEF cache
Write-Host "`nClearing Office WEF cache..." -ForegroundColor Cyan
if (Test-Path $wefPath) {
    try {
        Remove-Item -Path "$wefPath\*" -Recurse -Force -ErrorAction Stop
        Write-Host "✓ WEF cache cleared successfully: $wefPath" -ForegroundColor Green
    }
    catch {
        Write-Host "✗ Failed to clear WEF cache: $($_.Exception.Message)" -ForegroundColor Red
        Write-Host "  Try running as Administrator or close all Office applications" -ForegroundColor Yellow
    }
}
else {
    Write-Host "✓ WEF cache directory not found (nothing to clear): $wefPath" -ForegroundColor Green
}

# Clear temporary Office files
Write-Host "`nClearing temporary Office add-in files..." -ForegroundColor Cyan
try {
    $tempFiles = Get-ChildItem -Path "$env:TEMP" -Name "OfficeAddins*" -ErrorAction SilentlyContinue
    if ($tempFiles) {
        Remove-Item -Path "$env:TEMP\OfficeAddins*" -Recurse -Force -ErrorAction Stop
        Write-Host "✓ Temporary files cleared successfully" -ForegroundColor Green
    }
    else {
        Write-Host "✓ No temporary Office add-in files found" -ForegroundColor Green
    }
}
catch {
    Write-Host "✗ Failed to clear temporary files: $($_.Exception.Message)" -ForegroundColor Red
}

# Clear browser cache for Office
Write-Host "`nClearing Edge WebView2 cache (used by Office)..." -ForegroundColor Cyan
$edgeCachePath = "$env:LOCALAPPDATA\Microsoft\Office\16.0\Addins\WebCache"
if (Test-Path $edgeCachePath) {
    try {
        Remove-Item -Path "$edgeCachePath\*" -Recurse -Force -ErrorAction Stop
        Write-Host "✓ Edge WebView2 cache cleared successfully" -ForegroundColor Green
    }
    catch {
        Write-Host "✗ Failed to clear Edge WebView2 cache: $($_.Exception.Message)" -ForegroundColor Red
    }
}
else {
    Write-Host "✓ Edge WebView2 cache directory not found" -ForegroundColor Green
}

Write-Host "`n===================================" -ForegroundColor Green
Write-Host "Cache clearing process completed!" -ForegroundColor Green
Write-Host "Restart Outlook to ensure changes take effect." -ForegroundColor Yellow

# Optional: Restart Outlook if it was running
if ($outlookProcesses -and $Force) {
    Write-Host "`nRestarting Outlook..." -ForegroundColor Cyan
    try {
        Stop-Process -Name "OUTLOOK" -Force -ErrorAction Stop
        Start-Sleep -Seconds 3
        Start-Process "outlook.exe" -ErrorAction Stop
        Write-Host "✓ Outlook restarted successfully" -ForegroundColor Green
    }
    catch {
        Write-Host "✗ Failed to restart Outlook: $($_.Exception.Message)" -ForegroundColor Red
        Write-Host "  Please restart Outlook manually" -ForegroundColor Yellow
    }
}