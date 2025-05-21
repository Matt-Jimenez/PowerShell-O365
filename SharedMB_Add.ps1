# Prompt for the target user
$UserToGrant = Read-Host "Enter the user you want to grant permissions to (e.g., user@aus.com)"

# Ask if a CSV will be used
$useCSV = Read-Host "Do you want to use a CSV file with shared mailboxes? (Y/N)"

# Initialize mailbox list
$SharedMailboxes = @()

if ($useCSV.ToUpper() -eq "Y") {
    $csvPath = Read-Host "Enter the full path to the CSV file (e.g., C:\Users\YourName\Documents\mailboxes.csv)"
    if (Test-Path $csvPath) {
        try {
            $csvData = Import-Csv -Path $csvPath
            foreach ($entry in $csvData) {
                if ($entry.SharedMailbox -and -not [string]::IsNullOrWhiteSpace($entry.SharedMailbox)) {
                    $SharedMailboxes += $entry.SharedMailbox.Trim()
                }
                elseif ($entry.SharedMailbox)
                {
                    Write-Warning "Empty SharedMailbox entry found in CSV. Skipping."
                }
            }
        } catch {
            Write-Host "Error reading CSV file: $_" -ForegroundColor Red
            exit
        }
    } else {
        Write-Host "CSV file not found at: $csvPath" -ForegroundColor Red
        exit
    }
} else {
    $manualInput = Read-Host "Enter the shared mailboxes, separated by commas (e.g., shared1@aus.com,shared2@aus.com)"
    $SharedMailboxes = $manualInput -split "," | ForEach-Object { $_.Trim() }
}

# Ask which permissions to assign
$fullAccess = Read-Host "Grant Full Access permission? (Y/N)"
$sendAs = Read-Host "Grant Send As permission? (Y/N)"

# Apply permissions
foreach ($Shared in $SharedMailboxes) {
    if ($fullAccess.ToUpper() -eq "Y") {
        try {
            Add-MailboxPermission -Identity $Shared -User $UserToGrant -AccessRights FullAccess -InheritanceType All -AutoMapping:$true
            Write-Host "Granted Full Access to $UserToGrant for $Shared" -ForegroundColor Green
        } catch {
            Write-Host "Failed to grant Full Access to $Shared. Error: $_" -ForegroundColor Red
        }
    }

    if ($sendAs.ToUpper() -eq "Y") {
        try {
            Add-RecipientPermission -Identity $Shared -Trustee $UserToGrant -AccessRights SendAs -Confirm:$false
            Write-Host "Granted Send As to $UserToGrant for $Shared" -ForegroundColor Green
        } catch {
            Write-Host "Failed to grant Send As to $Shared. Error: $_" -ForegroundColor Red
        }
    }
}
