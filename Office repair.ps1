# ==========================================
#   Wojc's office fixer
# ==========================================

# 1. UNIVERSAL SELF-ELEVATION (File + Web Support)
$WebUrl = "https://raw.githubusercontent.com/wojcfurry-dev/office-fixer/main/Office%20repair.ps1"

if (!([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
    if ($PSCommandPath) {
        Write-Host "Restarting Local File as Admin..." -ForegroundColor Yellow
        Start-Process powershell.exe "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`"" -Verb RunAs
    } else {
        Write-Host "Restarting Web Script as Admin..." -ForegroundColor Yellow
        Start-Process powershell.exe "-NoProfile -ExecutionPolicy Bypass -Command `"irm '$WebUrl' | iex`"" -Verb RunAs
    }
    Exit
}

$ErrorActionPreference = "SilentlyContinue"

# 2. CLOSE OFFICE APPLICATIONS
Write-Host "Closing Office Applications..." -ForegroundColor Cyan
$Apps = "WINWORD", "EXCEL", "OUTLOOK", "POWERPNT", "ONENOTE", "TEAMS", "ONEDRIVE", "SystemSettings"
$Apps | ForEach-Object { Get-Process $_ | Stop-Process -Force }

# 3. MANUAL REMOVAL STEP
Write-Host "`n>>> ACTION REQUIRED <<<" -ForegroundColor Red
Write-Host "1. I am opening the 'Access work or school' settings."
Write-Host "2. Please find the stuck account and click DISCONNECT."
Write-Host "3. When the account is gone, come back here and PRESS ENTER."
Start-Sleep -Seconds 2
Start-Process "ms-settings:workplace"
Pause

# 4. FIX activation 
Write-Host "`nFixing office activation..." -ForegroundColor Yellow
$Policy = "HKLM:\SOFTWARE\Policies\Microsoft\Windows\WorkplaceJoin"
if (!(Test-Path $Policy)) { New-Item -Path $Policy -Force | Out-Null }
New-ItemProperty -Path $Policy -Name "BlockAADWorkplaceJoin" -Value 1 -PropertyType DWORD -Force | Out-Null
Write-Host "Prevention policy set." -ForegroundColor Green

# 5. WORD LAUNCH LOOP
Write-Host "`nOpening Word (Phase 1)..." -ForegroundColor Cyan
Start-Process "winword.exe"

Write-Host "Keeping Word open for 10 seconds..." -ForegroundColor Yellow
Start-Sleep -Seconds 10

Write-Host "Closing Word..." -ForegroundColor Cyan
Stop-Process -Name "WINWORD" -Force
Start-Sleep -Seconds 2

Write-Host "Re-opening Word (Final)..." -ForegroundColor Green
Start-Process "winword.exe"


Write-Host "`nDONE. Made with love in Hertfordshire :3"
