# --- UNIVERSAL MAINTENANCE TOOL V3.3 (Network Aware + Full Progress) ---
$ReportPath = "$env:USERPROFILE\Desktop\Update_Report.txt"

Function Show-Msg {
    param([string]$Message, [string]$Title = "System Maintenance")
    $Shell = New-Object -ComObject WScript.Shell
    $Shell.Popup($Message, 0, $Title, 64)
}

Write-Host "===============================================" -ForegroundColor Yellow
Write-Host "       PRECISION SYSTEM MAINTENANCE TOOL       " -ForegroundColor Yellow
Write-Host "===============================================" -ForegroundColor Yellow

# 1. Selection Menu
Write-Host "1) BIOS/Firmware Update (~5-10 mins)"
Write-Host "2) Cumulative Updates (~10-30 mins)"
Write-Host "3) Feature Update (~45-90 mins)"
Write-Host "4) Exit"
$Choice = Read-Host "`nSelect an option (1-4)"

if ($Choice -eq 4) { exit }

try {
    # 2. Safety Gates (Power & Network)
    Write-Progress -Activity "Initializing" -Status "Checking AC Power..." -PercentComplete 5
    $Battery = Get-CimInstance -ClassName Win32_Battery
    if ($Battery -and $Battery.BatteryStatus -ne 2) {
        Show-Msg "ABORT: System is on Battery. Please plug in AC Power." "Power Error"
        exit
    }

    Write-Progress -Activity "Initializing" -Status "Testing Network Latency..." -PercentComplete 10
    $Ping = Test-Connection -ComputerName "8.8.8.8" -Count 2 -Quiet
    if (-not $Ping) {
        Show-Msg "ABORT: No stable internet connection detected." "Network Error"
        exit
    }

    Write-Progress -Activity "Setup" -Status "Initializing Update Module..." -PercentComplete 15
    if (-not (Get-Module -ListAvailable PSWindowsUpdate)) {
        Install-Module -Name PSWindowsUpdate -Force -Confirm:$false -Scope CurrentUser
    }
    Import-Module PSWindowsUpdate

    $StatusMessage = ""
    $DoReboot = $false

    switch ($Choice) {
        "1" {
            $Model = (Get-CimInstance -ClassName Win32_ComputerSystem).Model
            Write-Progress -Activity "BIOS Update" -Status "Scanning for $Model Firmware..." -PercentComplete 30
            Suspend-BitLocker -MountPoint "C:" -RebootCount 1 -ErrorAction SilentlyContinue
            
            $Updates = Get-WindowsUpdate -MicrosoftUpdate -Category "Drivers" | 
                       Where-Object { ($_.Title -match "Firmware" -or $_.Title -match "BIOS") -and ($_.Title -match [regex]::Escape($Model)) }
            
            if ($Updates) {
                $Target = $Updates | Sort-Object Date -Descending | Select-Object -First 1
                Write-Progress -Activity "BIOS Update" -Status "Downloading: $($Target.Title)" -PercentComplete 70
                Install-WindowsUpdate -UpdateID $Target.UpdateID -AcceptAll -IgnoreReboot
                $StatusMessage = "SUCCESS: BIOS Update staged."
                $DoReboot = $true
            } else { $StatusMessage = "No matching BIOS found."; Show-Msg $StatusMessage }
        }

        "2" {
            Write-Progress -Activity "Cumulative Updates" -Status "Querying Microsoft Catalog..." -PercentComplete 30
            $Updates = Get-WindowsUpdate -MicrosoftUpdate -NotCategory "Drivers" -NotTitle "Feature Update"
            
            if ($Updates) {
                Write-Progress -Activity "Cumulative Updates" -Status "Downloading & Installing patches..." -PercentComplete 60
                Install-WindowsUpdate -MicrosoftUpdate -AcceptAll -IgnoreReboot
                $StatusMessage = "SUCCESS: Cumulative Updates installed."
                $DoReboot = $true
            } else { $StatusMessage = "System is up to date."; Show-Msg $StatusMessage }
        }

        "3" {
            Write-Progress -Activity "Feature Update" -Status "Checking for Major OS Upgrade..." -PercentComplete 30
            $Updates = Get-WindowsUpdate -MicrosoftUpdate -Title "Feature Update"
            
            if ($Updates) {
                Write-Progress -Activity "Feature Update" -Status "Downloading Large OS Files..." -PercentComplete 50
                Install-WindowsUpdate -MicrosoftUpdate -Title "Feature Update" -AcceptAll -IgnoreReboot
                $StatusMessage = "SUCCESS: Feature Update staged."
                $DoReboot = $true
            } else { $StatusMessage = "No Feature Updates available."; Show-Msg $StatusMessage }
        }
    }

    # 3. Finalization
    if ($DoReboot) {
        "Update Report - $(Get-Date)`nResult: $StatusMessage" | Out-File -FilePath $ReportPath
        Write-Progress -Activity "Finalizing" -Status "Preparing Reboot..." -PercentComplete 100
        shutdown.exe /r /t 120 /c "Updates staged. System will restart in 2 minutes. Save your work!"
        Show-Msg "Maintenance Staged. Computer will restart in 2 minutes." "Reboot Warning"
    }

} catch {
    Show-Msg "Critical Error: $($_.Exception.Message)" "Script Failed"
} finally {
    Write-Progress -Activity "Cleaning up" -Completed
}