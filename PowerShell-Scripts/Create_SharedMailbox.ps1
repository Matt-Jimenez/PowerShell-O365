# Define shared mailbox parameters
$mailboxName = Read-Host "Enter the name for the shared mailbox"
$mailboxAlias = Read-Host "Enter the alias for the shared mailbox"
$mailboxDisplayName = Read-Host "Enter the display name for the shared mailbox"
$domain = Read-Host "Enter your domain (e.g., yourdomain.com)"

# Validate inputs
if ([string]::IsNullOrWhiteSpace($mailboxName) -or [string]::IsNullOrWhiteSpace($mailboxAlias) -or [string]::IsNullOrWhiteSpace($mailboxDisplayName) -or [string]::IsNullOrWhiteSpace($domain)) {
    Write-Host "All fields are required. Please run the script again." -ForegroundColor Red
    exit
}

# Define default user list file path
$defaultUserListFile = Join-Path -Path $PSScriptRoot -ChildPath "Users-To-Add_SharedMailbox.csv"

# Check if user wants to use a custom file path
$useDefaultFile = Read-Host "Use default file path ($defaultUserListFile)? (Y/N)"

if ($useDefaultFile.ToUpper() -eq 'Y') {
    $userListFile = $defaultUserListFile
} else {
    $userListFile = Read-Host "Enter the path to the text file containing user list (e.g., C:\Temp\users.txt)"
}

# Check if file exists
if (-not (Test-Path $userListFile)) {
    Write-Host "File not found: $userListFile" -ForegroundColor Yellow
    $createFile = Read-Host "Would you like to create this file? (Y/N)"
    
    if ($createFile.ToUpper() -eq 'Y') {
        try {
            # Create directory if it doesn't exist
            $directory = Split-Path -Path $userListFile -Parent
            if (-not (Test-Path $directory) -and $directory) {
                New-Item -ItemType Directory -Path $directory -Force | Out-Null
            }
            
            # Create CSV file with header
            "UserPrincipalName" | Out-File -FilePath $userListFile -Encoding UTF8
            Write-Host "File created: $userListFile" -ForegroundColor Green
            
            # Check if Excel is installed
            $excelInstalled = $false
            try {
                $excelApp = New-Object -ComObject Excel.Application
                $excelInstalled = $true
                $excelApp.Quit()
                [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excelApp) | Out-Null
            } catch {
                $excelInstalled = $false
            }
            
            # Open file in Excel if installed, otherwise notify user
            if ($excelInstalled) {
                Write-Host "Opening file in Excel. Please add users (one per line) and save the file." -ForegroundColor Cyan
                Start-Process -FilePath $userListFile
                
                $fileReady = Read-Host "Press Enter when you have saved and closed the file"
            } else {
                Write-Host "Excel is not installed. Please edit the file manually and add one user principal name per line." -ForegroundColor Yellow
                Write-Host "File location: $userListFile" -ForegroundColor Yellow
                
                $fileReady = Read-Host "Press Enter when you have edited and saved the file"
            }
        } catch {
            Write-Host "Error creating file: $($_.Exception.Message)" -ForegroundColor Red
            exit
        }
    } else {
        Write-Host "Operation cancelled. Please create the user list file and run the script again." -ForegroundColor Yellow
        exit
    }
}

try {
    # Create the shared mailbox
    Write-Host "Creating shared mailbox '$mailboxDisplayName'..." -ForegroundColor Cyan
    New-Mailbox -Shared -Name $mailboxName -Alias $mailboxAlias -DisplayName $mailboxDisplayName -PrimarySmtpAddress "$mailboxAlias@$domain" -ErrorAction Stop
    Write-Host "Shared mailbox '$mailboxDisplayName' created successfully." -ForegroundColor Green
    
    # Get users from file
    $users = Get-Content $userListFile | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Select-Object -Skip 1
    
    # Validate users list
    if ($users.Count -eq 0) {
        Write-Host "No users found in the file." -ForegroundColor Yellow
        exit
    }
    
    Write-Host "Found $($users.Count) users in the file." -ForegroundColor Cyan
    
    # Add permissions for each user
    foreach ($user in $users) {
        try {
            # Trim any whitespace
            $user = $user.Trim()
            
            if ([string]::IsNullOrWhiteSpace($user)) {
                continue
            }
            
            # Add Full Access permission
            Add-MailboxPermission -Identity $mailboxName -User $user -AccessRights FullAccess -AutoMapping $true -ErrorAction Stop
            
            # Add Send As permission
            Add-RecipientPermission -Identity $mailboxName -Trustee $user -AccessRights SendAs -Confirm:$false -ErrorAction Stop
            
            Write-Host "User $user added to the shared mailbox '$mailboxName' with Full Access and Send As permissions." -ForegroundColor Green
        }
        catch {
            Write-Host "Failed to add user $user to the shared mailbox: $($_.Exception.Message)" -ForegroundColor Red
        }
    }
    
    Write-Host "All users have been processed." -ForegroundColor Yellow
}
catch {
    Write-Host "Error creating shared mailbox: $($_.Exception.Message)" -ForegroundColor Red
}

Write-Host "Script execution completed." -ForegroundColor Cyan