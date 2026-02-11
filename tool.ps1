# --- UNIVERSAL MAINTENANCE TOOL V3.5 (The "Bulletproof" Edition) ---
$ReportPath = "$env:USERPROFILE\Desktop\Update_Report.txt"

Function Show-Msg {
    param([string]$Message, [string]$Title = "System Maintenance")
    $Shell = New-Object -ComObject WScript.Shell
    $Shell.Popup($Message, 0, $Title, 64)
}

Write-Host "===============================================" -ForegroundColor Yellow
Write-Host "       BULLETPROOF BIOS & UPDATE TOOL          " -ForegroundColor Yellow
Write-Host "===============================================" -ForegroundColor Yellow

# 1. Selection Menu
Write-Host "1) BIOS/Firmware Update (Full Anti-Hang Suite)"
Write-Host "2) Cumulative Updates"
Write-Host "3) Feature Update"
Write-Host "4) Exit"
$Choice = Read-Host "`nSelect an option (1-4)"

if ($Choice -eq 4) { exit }

try {
    # 2. Initial Safety Checks
    Write-Progress -Activity "Initializing" -Status "Checking Power..." -PercentComplete 5
    $Battery = Get-CimInstance -ClassName Win32_Battery
    if ($Battery -and $Battery.BatteryStatus -ne 2) {
        Show-Msg "ABORT: AC Power is required to prevent BIOS corruption." "Power Error"
        exit
    }

    # 3. PRE-FLIGHT: Kill Fast Startup & Hibernate
    # This ensures a "Cold Boot" which is the only way to safely update BIOS with docks.
    Write-Progress -Activity "Pre-Flight" -Status "Disabling Fast Startup to prevent hangs..." -PercentComplete 15
    Set-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager\Power" -Name "HiberbootEnabled" -Value 0 -ErrorAction SilentlyContinue
    & powercfg.exe /hibernate off

    # 4. Module Prep
    if (-not (Get-Module -ListAvailable PSWindowsUpdate)) {
        Install-Module -Name PSWindowsUpdate -Force -Confirm:$false -Scope CurrentUser
    }
    Import-Module PSWindowsUpdate

    switch ($Choice) {
        "1" {
            $Model = (Get-CimInstance -ClassName Win32_ComputerSystem).Model
            
            # Suspend BitLocker for 2 reboots (Handles the Flash + Verification reboots)
            Write-Host "Suspending BitLocker (2 Reboots)..."
            Suspend-BitLocker -MountPoint "C:" -RebootCount 2 -ErrorAction SilentlyContinue

            Write-Progress -Activity "BIOS Update" -Status "Matching Firmware for $Model..." -PercentComplete 30
            $Updates = Get-WindowsUpdate -MicrosoftUpdate -Category "Drivers" | 
                       Where-Object { ($_.Title -match "Firmware" -or $_.Title -match "BIOS") -and ($_.Title -match [regex]::Escape($Model)) }
            
            if ($Updates) {
                $Target = $Updates | Sort-Object Date -Descending | Select-Object -First 1
                Write-Progress -Activity "BIOS Update" -Status "Staging Payload..." -PercentComplete 70
                Install-WindowsUpdate -UpdateID $Target.UpdateID -AcceptAll -IgnoreReboot
                $StatusMessage = "SUCCESS: BIOS staged for $Model. Fast Startup Disabled."
                $DoReboot = $true
            } else { Show-Msg "No matching BIOS found." }
        }

        "2" {
            Write-Progress -Activity "Cumulative" -Status "Installing Patches..." -PercentComplete 50
            Install-WindowsUpdate -MicrosoftUpdate -NotCategory "Drivers" -AcceptAll -IgnoreReboot
            $DoReboot = $true
        }

        "3" {
            Write-Progress -Activity "Feature" -Status "Installing OS Upgrade..." -PercentComplete 50
            Install-WindowsUpdate -MicrosoftUpdate -Title "Feature Update" -AcceptAll -IgnoreReboot
            $DoReboot = $true
        }
    }

    # 5. The "Deep Cold Boot" Shutdown
    if ($DoReboot) {
        "Update Report`nResult: $StatusMessage" | Out-File -FilePath $ReportPath
        
        Write-Progress -Activity "Finalizing" -Status "Triggering Cold Reboot..." -PercentComplete 100
        
        # /r = Restart, /f = Force close apps, /t 120 = 2 min timer
        # The inclusion of 'powercfg /h off' earlier ensures this is a TRUE cold restart.
        shutdown.exe /r /f /t 120 /c "BIOS/Updates Staged. SYSTEM WILL COLD BOOT IN 2 MINUTES. If using a Dock and the screen stays black, do not power off manually."
        
        Show-Msg "Process Staged. Computer will perform a COLD RESTART in 2 minutes." "Reboot Warning"
    }

} catch {
    Show-Msg "Critical Error: $($_.Exception.Message)" "Script Failed"
}