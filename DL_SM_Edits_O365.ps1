# Prompt the host to enter the username
$username = Read-Host "Enter the username"

# Verify if the username exists
try {
    $user = Get-ADUser -Identity $username -ErrorAction Stop
    Write-Host "User '$username' found." -ForegroundColor Green
} catch {
    Write-Host "User '$username' not found: $($_.Exception.Message)" -ForegroundColor Red
    exit
}

# Ask whether to add to a distribution list or a shared mailbox
$choice = Read-Host "Enter 'DL' for Distribution List or 'SM' for Shared Mailbox"

switch ($choice.ToUpper()) {
    'DL' {
        # Prompt for distribution list name
        $distributionList = Read-Host "Enter the Distribution List name"
        
        # Ask if adding or removing
        $action = Read-Host "Enter 'Add' to add users or 'Remove' to remove users"
        
        if ($action.ToUpper() -eq 'REMOVE') {
            # Check if there is a CSV file
            $csvChoice = Read-Host "Is there a CSV file of users to remove? (Yes/No)"
            if ($csvChoice.ToUpper() -eq 'YES') {
                $csvPath = Read-Host "Enter the path to the CSV file"
                $users = Import-Csv -Path $csvPath | Select-Object -ExpandProperty UserName
            } else {
                $userList = Read-Host "Enter the list of users separated by comma"
                $users = $userList -split ','
            }
            
            foreach ($user in $users) {
                try {
                    Remove-DistributionGroupMember -Identity $distributionList -Member $user -ErrorAction Stop
                    Write-Host "Removed user '$user' from distribution list $distributionList." -ForegroundColor Green
                } catch {
                    Write-Host "Failed to remove user '$user': $($_.Exception.Message)" -ForegroundColor Red
                }
            }
            exit
        }
        try {
            # Check if distribution list exists
            $dlExists = Get-DistributionGroup -Identity $distributionList -ErrorAction Stop
            Write-Host "Distribution list '$distributionList' found." -ForegroundColor Green
            
            # Check if the user is already a member of the distribution list
            $member = Get-DistributionGroupMember -Identity $distributionList | Where-Object { $_.SamAccountName -eq $username }
            if ($member) {
                Write-Host "User '$username' is already in the distribution list $distributionList." -ForegroundColor Yellow
            } else {
                # Add the user to the distribution list
                Add-DistributionGroupMember -Identity $distributionList -Member $username -ErrorAction Stop
                Write-Host "Added user '$username' to distribution list $distributionList." -ForegroundColor Green
            }
        } catch {
            Write-Host "Failed to process distribution list operation: $($_.Exception.Message)" -ForegroundColor Red
        }
    }
    'SM' {
        # Prompt for shared mailbox name
        $sharedMailbox = Read-Host "Enter the Shared Mailbox name"
        
        # Ask if adding or removing
        $action = Read-Host "Enter 'Add' to add users or 'Remove' to remove users"
        
        if ($action.ToUpper() -eq 'REMOVE') {
            # Check if there is a CSV file
            $csvChoice = Read-Host "Is there a CSV file of users to remove? (Yes/No)"
            if ($csvChoice.ToUpper() -eq 'YES') {
                $csvPath = Read-Host "Enter the path to the CSV file"
                $users = Import-Csv -Path $csvPath | Select-Object -ExpandProperty UserName
            } else {
                $userList = Read-Host "Enter the list of users separated by comma"
                $users = $userList -split ','
            }
            
            foreach ($user in $users) {
                try {
                    Remove-MailboxPermission -Identity $sharedMailbox -User $user -AccessRights FullAccess -Confirm:$false -ErrorAction Stop
                    Write-Host "Removed Full Access permissions for user '$user' from shared mailbox $sharedMailbox." -ForegroundColor Green
                } catch {
                    Write-Host "Failed to remove permissions for user '$user': $($_.Exception.Message)" -ForegroundColor Red
                }
            }
            exit
        }
        try {
            # Check if shared mailbox exists
            $mbExists = Get-Mailbox -Identity $sharedMailbox -ErrorAction Stop
            Write-Host "Shared mailbox '$sharedMailbox' found." -ForegroundColor Green

            # Check if user already has permissions
            $existingPermissions = Get-MailboxPermission -Identity $sharedMailbox | Where-Object { $_.User -like "*$username*" -and $_.AccessRights -contains "FullAccess" }
            if ($existingPermissions) {
                Write-Host "User '$username' already has Full Access permissions to shared mailbox $sharedMailbox." -ForegroundColor Yellow
            } else {
                # Add the user to the shared mailbox with appropriate rights
                Add-MailboxPermission -Identity $sharedMailbox -User $username -AccessRights FullAccess -AutoMapping $true -ErrorAction Stop
                Write-Host "Added user '$username' with Full Access permissions to shared mailbox $sharedMailbox." -ForegroundColor Green
            }
            
            # Check if user already has Send As permissions
            $existingSendAs = Get-RecipientPermission -Identity $sharedMailbox | Where-Object { $_.Trustee -like "*$username*" -and $_.AccessRights -contains "SendAs" }
            if ($existingSendAs) {
                Write-Host "User '$username' already has Send As permissions to shared mailbox $sharedMailbox." -ForegroundColor Yellow
            } else {
                Add-RecipientPermission -Identity $sharedMailbox -Trustee $username -AccessRights SendAs -Confirm:$false -ErrorAction Stop
                Write-Host "Added Send As permissions for user '$username' to shared mailbox $sharedMailbox." -ForegroundColor Green
            }
        } catch {
            Write-Host "Failed to process shared mailbox operation: $($_.Exception.Message)" -ForegroundColor Red
        }
    }
    default {
        Write-Host "Invalid option selected. Please enter 'DL' for Distribution List or 'SM' for Shared Mailbox." -ForegroundColor Red
    }
}

Write-Host "Script execution completed." -ForegroundColor Cyan