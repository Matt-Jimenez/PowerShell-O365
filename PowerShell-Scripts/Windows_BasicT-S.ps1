# Script to provide a menu for common troubleshooting commands

# --- Helper Function to Run Commands ---
function Start-AdminCommand {
    param(
        [string]$Command,
        [string]$WindowTitle = "Command Output"
    )
    Write-Host "Running: $Command (New window will open)"
    try {
        # Using /k to keep the window open after command execution
        Start-Process cmd.exe -ArgumentList "/k $Command" -Verb RunAs -Wait
        Write-Host "Command window for '$WindowTitle' closed. Press Enter to return to menu."
        Read-Host
    } catch {
        Write-Warning "Failed to start process for '$WindowTitle'. Error: $($_.Exception.Message)"
        Write-Host "This might happen if you cancel the UAC prompt or if there's an issue with cmd.exe."
        Read-Host "Press Enter to continue..."
    }
}

# --- Network Troubleshooting Functions ---
function Show-NetworkMenu {
    Clear-Host
    Write-Host "==============================================="
    Write-Host " Network Troubleshooting Command Menu "
    Write-Host "==============================================="
    Write-Host "Please select an option:"
    Write-Host "1. ipconfig /all (Display all TCP/IP configuration)"
    Write-Host "2. Find and Ping Default Gateway"
    Write-Host "3. ipconfig /release (Release current IP address)"
    Write-Host "4. ipconfig /renew (Renew IP address)"
    Write-Host "5. ipconfig /flushdns (Flush DNS resolver cache)"
    Write-Host "6. netsh int tcp set global autotuninglevel=normal (Set TCP auto-tuning to normal)"
    Write-Host "7. netsh interface tcp show heuristics (Show TCP heuristics state)"
    Write-Host "8. netsh interface tcp set heuristics disabled (Disable TCP heuristics)"
    Write-Host "B. Back to Main Menu"
    Write-Host "Q. Quit"
    Write-Host "==============================================="
}

function Find-And-Ping-Gateway {
    Write-Host "Attempting to find and ping your Default Gateway..."
    try {
        $ipconfigOutput = (ipconfig | Out-String) -split [System.Environment]::NewLine
        $gatewayIP = $null

        foreach ($line in $ipconfigOutput) {
            if ($line -match "Default Gateway.*?: ((?:\d{1,3}\.){3}\d{1,3})(?!\d)") {
                $gatewayIP = $Matches[1].Trim()
                if ($gatewayIP -ne "0.0.0.0") { break } else { $gatewayIP = $null }
            }
        }

        if ($gatewayIP) {
            Write-Host "Default Gateway found: $gatewayIP"
            Start-AdminCommand -Command "ping $gatewayIP -t" -WindowTitle "Ping Gateway"
        } else {
            Write-Warning "Could not automatically determine a valid IPv4 Default Gateway."
            Write-Host "You can run 'ipconfig /all' (Option 1) to check manually."
            Read-Host "Press Enter to continue..."
        }
    } catch {
        Write-Error "An error occurred while trying to find the gateway: $($_.Exception.Message)"
        Read-Host "Press Enter to continue..."
    }
}

function Handle-NetworkMenu {
    do {
        Show-NetworkMenu
        $selection = Read-Host "Enter your choice"
        $goBack = $false

        switch ($selection) {
            "1" { Start-AdminCommand -Command "ipconfig /all" -WindowTitle "ipconfig /all" }
            "2" { Find-And-Ping-Gateway }
            "3" { Start-AdminCommand -Command "ipconfig /release" -WindowTitle "ipconfig /release" }
            "4" { Start-AdminCommand -Command "ipconfig /renew" -WindowTitle "ipconfig /renew" }
            "5" { Start-AdminCommand -Command "ipconfig /flushdns" -WindowTitle "ipconfig /flushdns" }
            "6" { Start-AdminCommand -Command "netsh int tcp set global autotuninglevel=normal" -WindowTitle "Set TCP Autotuning" }
            "7" { Start-AdminCommand -Command "netsh interface tcp show heuristics" -WindowTitle "Show TCP Heuristics" }
            "8" { Start-AdminCommand -Command "netsh interface tcp set heuristics disabled" -WindowTitle "Disable TCP Heuristics" }
            "B" { $goBack = $true }
            "Q" { Write-Host "Exiting script."; exit }
            default {
                Write-Warning "Invalid selection. Please try again."
                Read-Host "Press Enter to continue..."
            }
        }
    } while (-not $goBack)
}

# --- System (Non-Network) Troubleshooting Functions ---
function Show-SystemMenu {
    Clear-Host
    Write-Host "==============================================="
    Write-Host " System (Non-Network) Troubleshooting Menu "
    Write-Host "==============================================="
    Write-Host "Please select an option:"
    Write-Host "1. sfc /scannow (Scan and repair system files)"
    Write-Host "2. DISM /Online /Cleanup-Image /RestoreHealth (Repair Windows image)"
    Write-Host "3. chkdsk /f /r (Check disk for errors - requires reboot if issues found on C:)"
    Write-Host "4. Manage Windows Update Service (wuauserv)"
    Write-Host "5. Manage Background Intelligent Transfer Service (BITS)"
    Write-Host "6. Delete Windows Update Cache (SoftwareDistribution)"
    Write-Host "B. Back to Main Menu"
    Write-Host "Q. Quit"
    Write-Host "==============================================="
}

function Run-Chkdsk {
    $driveLetter = Read-Host "Enter drive letter to check (e.g., C). Default is C"
    if ([string]::IsNullOrWhiteSpace($driveLetter)) {
        $driveLetter = "C"
    }
    $drive = $driveLetter.ToUpper() + ":"
    Write-Host "Note: If '$drive' is the system drive or in use, a reboot might be required to complete the scan." -ForegroundColor Yellow
    $confirm = Read-Host "Proceed with chkdsk $drive /f /r? (yes/no)"
    if ($confirm -eq 'yes') {
        Start-AdminCommand -Command "chkdsk $drive /f /r" -WindowTitle "chkdsk $drive"
    } else {
        Write-Host "chkdsk aborted."
        Read-Host "Press Enter to continue..."
    }
}

function Manage-ServiceAction {
    param(
        [string]$ServiceName,
        [string]$ServiceDisplayName
    )
    $action = Read-Host "Do you want to (S)tart or (T)op the '$ServiceDisplayName' service? (S/T)"
    switch ($action.ToUpper()) {
        "S" { Start-AdminCommand -Command "net start $ServiceName" -WindowTitle "Start $ServiceDisplayName" }
        "T" { Start-AdminCommand -Command "net stop $ServiceName" -WindowTitle "Stop $ServiceDisplayName" }
        default { Write-Warning "Invalid action. Please choose S or T." ; Read-Host "Press Enter..." }
    }
}

function Clear-SoftwareDistributionCache {
    Write-Warning "This will delete the contents of the Windows Update cache (%windir%\SoftwareDistribution)."
    $confirm = Read-Host "Are you sure you want to proceed? (yes/no)"
    if ($confirm -eq 'yes') {
        # Stopping services first is often recommended
        Write-Host "Attempting to stop Windows Update and BITS services..."
        Start-AdminCommand -Command "net stop wuauserv && net stop bits" -WindowTitle "Stopping Services"
        Write-Host "Services stop command issued. Proceeding with cache deletion..."
        Start-AdminCommand -Command "del %windir%\SoftwareDistribution\*.* /s /q" -WindowTitle "Delete SoftwareDistribution"
        Write-Host "Cache deletion command issued. It's recommended to restart the services or reboot."
        # Optionally, offer to restart services
        $restartServices = Read-Host "Do you want to attempt to restart Windows Update and BITS services now? (yes/no)"
        if ($restartServices -eq 'yes') {
            Start-AdminCommand -Command "net start wuauserv && net start bits" -WindowTitle "Starting Services"
        }
    } else {
        Write-Host "SoftwareDistribution cache deletion aborted."
        Read-Host "Press Enter to continue..."
    }
}

function Handle-SystemMenu {
    do {
        Show-SystemMenu
        $selection = Read-Host "Enter your choice"
        $goBack = $false

        switch ($selection) {
            "1" { Start-AdminCommand -Command "sfc /scannow" -WindowTitle "SFC Scan" }
            "2" { Start-AdminCommand -Command "DISM /Online /Cleanup-Image /RestoreHealth" -WindowTitle "DISM RestoreHealth" }
            "3" { Run-Chkdsk }
            "4" { Manage-ServiceAction -ServiceName "wuauserv" -ServiceDisplayName "Windows Update" }
            "5" { Manage-ServiceAction -ServiceName "bits" -ServiceDisplayName "Background Intelligent Transfer Service" }
            "6" { Clear-SoftwareDistributionCache }
            "B" { $goBack = $true }
            "Q" { Write-Host "Exiting script."; exit }
            default {
                Write-Warning "Invalid selection. Please try again."
                Read-Host "Press Enter to continue..."
            }
        }
    } while (-not $goBack)
}

# --- Main Menu Function ---
function Show-MainMenu {
    Clear-Host
    Write-Host "==============================================="
    Write-Host " Windows Troubleshooting Utility "
    Write-Host "==============================================="
    Write-Host "Please select a category:"
    Write-Host "1. Network Troubleshooting"
    Write-Host "2. System (Non-Network) Troubleshooting"
    Write-Host "Q. Quit"
    Write-Host "==============================================="
}

# --- Main Script Loop ---
do {
    Show-MainMenu
    $mainSelection = Read-Host "Enter your choice"

    switch ($mainSelection) {
        "1" { Handle-NetworkMenu }
        "2" { Handle-SystemMenu }
        "Q" { Write-Host "Exiting script."; $exitScript = $true }
        default {
            Write-Warning "Invalid selection. Please try again."
            Read-Host "Press Enter to continue..."
        }
    }
} while (-not $exitScript)