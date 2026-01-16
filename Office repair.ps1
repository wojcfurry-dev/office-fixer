# ==========================================
#   Wojc's office fixer
# ==========================================

# 1. SELF-ELEVATE (Run as Admin automatically)
if (!([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
    Write-Host "Requesting Admin rights..." -ForegroundColor Yellow
    Start-Process powershell.exe "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`"" -Verb RunAs
    Exit
}

$ErrorActionPreference = "SilentlyContinue"

# 2. KILL OFFICE APPS
Write-Host "Closing Office Applications..." -ForegroundColor Cyan
$Apps = "WINWORD", "EXCEL", "OUTLOOK", "POWERPNT", "ONENOTE", "TEAMS", "ONEDRIVE", "SystemSettings"
$Apps | ForEach-Object { Get-Process $_ | Stop-Process -Force }

# 3. OPEN SETTINGS & PAUSE
Write-Host "`n>>> ACTION REQUIRED <<<" -ForegroundColor Red
Write-Host "1. I am opening Settings > Access work or school."
Write-Host "2. Find the stuck account and click DISCONNECT."
Write-Host "3. When the account is gone, come back here and PRESS ENTER."
Start-Sleep -Seconds 2
Start-Process "ms-settings:workplace"
Pause

# 4. licence fix (Registry Lock)
$Policy = "HKLM:\SOFTWARE\Policies\Microsoft\Windows\WorkplaceJoin"
if (!(Test-Path $Policy)) { New-Item -Path $Policy -Force | Out-Null }
New-ItemProperty -Path $Policy -Name "BlockAADWorkplaceJoin" -Value 1 -PropertyType DWORD -Force | Out-Null



# 6. RESTART SERVICES & LAUNCH OFFICE
Write-Host "Restarting Auth Services..." -ForegroundColor Green
Stop-Service -Name "TokenBroker" -Force
Start-Service -Name "TokenBroker"

Write-Host "Launching Word..." -ForegroundColor Green
Start-Process "winword.exe"