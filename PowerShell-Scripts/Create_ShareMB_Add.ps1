<#
.SYNOPSIS
Creates a shared mailbox and grants users full access and Send As permissions.

.DESCRIPTION
This script creates a new shared mailbox, prompts for user input, and grants
specified users both FullAccess and SendAs permissions.  It continues to prompt
for users until you indicate you are finished.

.PARAMETER SharedMailboxName
The name of the shared mailbox to create.

.PARAMETER SharedMailboxEmailAddress
The email address for the new shared mailbox.

.NOTES
Requires the Exchange Online PowerShell module.  Ensure you are connected with
appropriate permissions before running this script.
#>
[CmdletBinding()]
param (
    [Parameter(Mandatory = $true)]
    [string]$SharedMailboxName,

    [Parameter(Mandatory = $true)]
    [string]$SharedMailboxEmailAddress
)

# Connect to Exchange Online (if not already connected)
# This part assumes you have already connected, if not, you can add the Connect-ExchangeOnline here.
# Example (you might need to adjust based on your environment):
try {
    Connect-ExchangeOnline
} catch {
    Write-Error "Failed to connect to Exchange Online: $($_.Exception.Message)"
    exit 1
}

# Create the shared mailbox
try {
    New-Mailbox -Shared -Name $SharedMailboxName -DisplayName $SharedMailboxName -PrimarySmtpAddress $SharedMailboxEmailAddress
    Write-Host -ForegroundColor Green "Successfully created shared mailbox: '$SharedMailboxName' with email address '$SharedMailboxEmailAddress'."
} catch {
    Write-Error "Error creating shared mailbox: $($_.Exception.Message)"
    exit 1
}

# Prompt to add users
$AddUsers = Read-Host "Do you want to add users to this shared mailbox? (yes/no)"

if ($AddUsers -eq "yes") {
    do {
        $UserToAdd = Read-Host "Enter the user's email address"

        # Check for empty input
        if ([string]::IsNullOrEmpty($UserToAdd)) {
            Write-Warning "Warning: User email address cannot be empty. Please enter a valid email address or 'done'."
            continue
        }

        try {
            # Grant FullAccess permission
            Add-MailboxPermission -Identity $SharedMailboxEmailAddress -User $UserToAdd -AccessRights FullAccess -InheritanceType All
            Write-Host -ForegroundColor Green "Granted FullAccess to '$UserToAdd' for shared mailbox '$SharedMailboxName'."

            # Grant SendAs permission
            Add-RecipientPermission -Identity $SharedMailboxEmailAddress -Trustee $UserToAdd -AccessRights SendAs
            Write-Host -ForegroundColor Green "Granted SendAs permission to '$UserToAdd' for shared mailbox '$SharedMailboxName'."

        } catch {
            Write-Error "Error granting permissions to '$UserToAdd' for shared mailbox '$SharedMailboxName': $($_.Exception.Message)"
        }

        $AddMoreUsers = Read-Host "Do you want to add another user? (yes/no)"
    } while ($AddMoreUsers -eq "yes")
} else {
    Write-Host -ForegroundColor Yellow "No users will be added to the shared mailbox at this time."
}

Write-Host -ForegroundColor Yellow "Script completed."
