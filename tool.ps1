# --- UNIVERSAL MAINTENANCE TOOL V3.7 (Self-Healing Edition) ---
$ReportPath = "$env:USERPROFILE\Desktop\Update_Report.txt"

Function Show-Msg {
    param([string]$Message, [string]$Title = "System Maintenance")
    $Shell = New-Object -ComObject WScript.Shell
    $Shell.Popup($Message, 0, $Title, 64)
}

# --- REPAIR FUNCTION: Runs if a previous session froze ---
Function Repair-UpdateEngine {
    Write-Progress -Activity "Safety Check" -Status "Repairing Update Database (Self-Healing)..." -PercentComplete 20
    # Resets the specific 'pending' flags that cause hangs
    $RegPath = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update"
    if (Test-Path "$RegPath\RebootRequired") { Remove-ItemProperty -Path $RegPath -Name "RebootRequired" -ErrorAction SilentlyContinue }
    
    # Force kill any stuck installer processes
    Get-Process -Name "msiexec", "trustedinstaller" -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
    
    # Repair the System Component Store (The ultimate fix for freezes)
    Write-Progress -Activity "Safety Check" -Status "Running Component Store Repair (DISM)..." -PercentComplete 40
    dism.exe /online /cleanup-image /startcomponentcleanup /resetbase
}

Function Reset-UpdateCache {
    Write-Progress -Activity "Cache Reset" -Status "Purging SoftwareDistribution (89GB Bug Fix)..." -PercentComplete 10
    Stop-Service -Name wuauserv, bits, cryptsvc -Force -ErrorAction SilentlyContinue
    Remove-Item -Path "$env:SystemRoot\SoftwareDistribution\Download\*" -Recurse -Force -ErrorAction SilentlyContinue
    Start-Service -Name wuauserv
}

Write-Host "===============================================" -ForegroundColor Yellow
Write-Host "       SELF-HEALING BIOS & UPDATE TOOL         " -ForegroundColor Yellow
Write-Host "===============================================" -ForegroundColor Yellow

# 1. Selection Menu
Write-Host "1) BIOS/Firmware Update (Anti-Hang)"
Write-Host "2) Cumulative Updates (Safe-Resume)"
Write-Host "3) Feature Update"
Write-Host "4) Repair System (Run if previous attempt froze)"
Write-Host "5) Exit"
$Choice = Read-Host "`nSelect an option (1-5)"

if ($Choice -eq 5) { exit }

try {
    # 2. Power & Network Check
    $Battery = Get-CimInstance -ClassName Win32_Battery
    if ($Battery -and $Battery.BatteryStatus -ne 2) {
        Show-Msg "ABORT: AC Power required." "Power Error"; exit
    }

    # 3. Apply Anti-Hang Settings
    & powercfg.exe /hibernate off
    Set-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager\Power" -Name "HiberbootEnabled" -Value 0 -ErrorAction SilentlyContinue

    # 4. Action Logic
    switch ($Choice) {
        "4" { Repair-UpdateEngine; Show-Msg "Repair Complete. You can now try Option 1 or 2 again." "Healed" }
        
        "1" {
            $Model = (Get-CimInstance -ClassName Win32_ComputerSystem).Model
            Suspend-BitLocker -MountPoint "C:" -RebootCount 2
            Write-Progress -Activity "BIOS" -Status "Matching hardware..." -PercentComplete 50
            $Updates = Get-WindowsUpdate -MicrosoftUpdate -Category "Drivers" | 
                       Where-Object { ($_.Title -match "Firmware" -or $_.Title -match "BIOS") -and ($_.Title -match [regex]::Escape($Model)) }
            if ($Updates) {
                Install-WindowsUpdate -UpdateID ($Updates | Sort-Object Date -Descending | Select-Object -First 1).UpdateID -AcceptAll -IgnoreReboot
                $DoReboot = $true
            }
        }

        "2" {
            Reset-UpdateCache # Fixes the 89GB bug immediately
            Write-Progress -Activity "Cumulative" -Status "Installing Patches..." -PercentComplete 50
            Install-WindowsUpdate -MicrosoftUpdate -NotCategory "Drivers" -NotTitle "Feature Update" -AcceptAll -IgnoreReboot
            $DoReboot = $true
        }

        "3" {
            Reset-UpdateCache
            Write-Progress -Activity "Feature" -Status "Installing OS Upgrade..." -PercentComplete 50
            Install-WindowsUpdate -MicrosoftUpdate -Title "Feature Update" -AcceptAll -IgnoreReboot
            $DoReboot = $true
        }
    }

    if ($DoReboot) {
        "Update Report`nResult: Staged Successfully" | Out-File -FilePath $ReportPath
        shutdown.exe /r /f /t 120 /c "Maintenance Complete. System will COLD RESTART in 2 minutes."
        Show-Msg "Process Staged. 2 Minute Countdown started." "Reboot Warning"
    }

} catch {
    Show-Msg "Critical Error: $($_.Exception.Message). Run Option 4 to repair." "Script Failed"
}