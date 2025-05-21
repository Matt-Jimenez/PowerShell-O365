<#
.SYNOPSIS
    Prompts for computer names, searches Active Directory for each one,
    prompts to disable the computer accounts, and captures separation ticket
    information for clipboard use.

.DESCRIPTION
    This script allows you to search for and disable multiple computer objects in
    Active Directory. For each computer, it asks for confirmation before
    disabling the account and prompts for a separation ticket number which is
    placed in the clipboard for easy pasting into documentation.

.NOTES
    Requires the Active Directory PowerShell module to be installed.
    Run this script with administrative privileges.
#>

# Define color functions
function Write-ColoredOutput {
    param(
        [Parameter(Mandatory=$true)]
        [string]$Message,
        [Parameter(Mandatory=$true)]
        [ValidateSet("Red", "Yellow", "Green", "White")]
        [string]$Color
    )
    switch ($Color) {
        "Red"    { Write-Host -ForegroundColor Red $Message }
        "Yellow" { Write-Host -ForegroundColor Yellow $Message }
        "Green"  { Write-Host -ForegroundColor Green $Message }
        "White"  { Write-Host $Message } # Default color
    }
}

# Check if the Active Directory module is installed
if (-not (Get-Module -ListAvailable -Name ActiveDirectory)) {
    Write-ColoredOutput -Message "The Active Directory PowerShell module is not installed. Please install it before running this script." -Color Red
    exit 1
}

# Import the Active Directory module
Import-Module ActiveDirectory

# Add clipboard functionality
Add-Type -AssemblyName System.Windows.Forms

# Main processing loop
$continueProcessing = $true

while ($continueProcessing) {
    # Prompt the user for the computer name
    $ComputerName = Read-Host "Enter the computer name to search for (or type 'exit' to quit)"
    
    # Check if user wants to exit
    if ($ComputerName -eq 'exit') {
        $continueProcessing = $false
        break
    }
    
    # Search Active Directory for the computer
    Write-ColoredOutput -Message "Searching Active Directory for computer '$ComputerName'..." -Color White
    $Computer = Get-ADComputer -Filter "Name -eq '$ComputerName'"
    
    # Check if the computer was found
    if (-not $Computer) {
        Write-ColoredOutput -Message "Computer '$ComputerName' not found in Active Directory." -Color Yellow
        continue
    }
    
    # Display the found computer information
    Write-ColoredOutput -Message "Found the following computer:" -Color Green
    Write-Host "  Name          : $($Computer.Name)"
    Write-Host "  Enabled       : $($Computer.Enabled)"
    
    # Prompt the user to disable the computer if it's enabled
    if ($Computer.Enabled) {
        $DisableConfirmation = Read-Host "Do you want to disable the computer account '$ComputerName'? (yes/no)"
        if ($DisableConfirmation -ceq "yes") {
            try {
                Disable-ADAccount -Identity $Computer.SamAccountName
                Write-ColoredOutput -Message "Computer account '$ComputerName' has been disabled." -Color Green
                
                # Prompt for separation ticket number
                $TicketNumber = Read-Host "Enter the separation ticket number"
                
                # Create the formatted text for clipboard
                $ClipboardText = "Separation:$TicketNumber"
                
                # Copy to clipboard
                [System.Windows.Forms.Clipboard]::SetText($ClipboardText)
                
                Write-ColoredOutput -Message "'$ClipboardText' has been copied to clipboard. You can now paste it into the computer description field in AD." -Color Green
            } catch {
                Write-ColoredOutput -Message "Error disabling computer account '$ComputerName': $($_.Exception.Message)" -Color Red
            }
        } else {
            Write-ColoredOutput -Message "Skipping disabling the computer account." -Color Yellow
        }
    } else {
        Write-ColoredOutput -Message "Computer account '$ComputerName' is already disabled." -Color Yellow
        
        # Still prompt for ticket number even if already disabled
        $TicketNumber = Read-Host "Enter the separation ticket number"
        
        # Create the formatted text for clipboard
        $ClipboardText = "Separation:$TicketNumber"
        
        # Copy to clipboard
        [System.Windows.Forms.Clipboard]::SetText($ClipboardText)
        
        Write-ColoredOutput -Message "'$ClipboardText' has been copied to clipboard. You can now paste it into the computer description field in AD." -Color Green
    }
    
    Write-ColoredOutput -Message "Processing completed for computer '$ComputerName'." -Color Green
    Write-Host ""
}