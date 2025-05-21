# Function to add user to an Active Directory group and handle success/failure messages
function AddUserToADGroup {
    param(
        [string]$UsernameParameter,
        [string]$GroupNameParameter
    )
    try {
        # Check if the group exists first
        Get-ADGroup -Identity $GroupNameParameter -ErrorAction Stop | Out-Null
        
        # If group exists, try to add the user
        Add-ADGroupMember -Identity $GroupNameParameter -Members $UsernameParameter -ErrorAction Stop
        Write-Host "User '$UsernameParameter' successfully added to AD group '$GroupNameParameter'." -ForegroundColor Green
    }
    catch [Microsoft.ActiveDirectory.Management.ADIdentityNotFoundException] {
        # Specific catch for when the group is not found
        Write-Host "Failed to add user '$UsernameParameter' to AD group '$GroupNameParameter': Group '$GroupNameParameter' not found." -ForegroundColor Red
    }
    catch {
        # General catch for other errors, like user already being a member or other permission issues
        Write-Host "Failed to add user '$UsernameParameter' to AD group '$GroupNameParameter': $($_.Exception.Message)" -ForegroundColor Red
    }
}

# Main script starts here
# Prompt the host to enter the username
$scriptUsername = Read-Host "Enter the username"

# Verify if the username exists (using Get-ADUser for Active Directory)
try {
    $adUserObject = Get-ADUser -Identity $scriptUsername -ErrorAction Stop
    Write-Host "User '$($adUserObject.SamAccountName)' found and verified in Active Directory." -ForegroundColor Cyan
} catch {
    Write-Host "User '$scriptUsername' not found in Active Directory. $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "Please ensure the username is correct and the Active Directory module is available." -ForegroundColor Yellow
    exit # Exit the script if the user is not found
}

# Loop to handle adding to multiple Active Directory groups
do {
    # Prompting to enter the Active Directory group name to add the user to
    $adGroupName = Read-Host "Enter the Active Directory group name to add '$scriptUsername' to (or press Enter to skip)"
    
    if (-not ([string]::IsNullOrWhiteSpace($adGroupName))) {
        AddUserToADGroup -UsernameParameter $scriptUsername -GroupNameParameter $adGroupName
    } elseif ($adGroupName -eq "") {
        Write-Host "No group name entered, skipping this step." -ForegroundColor Yellow
    }

    # Ask if there are any other Active Directory groups to add the user to
    $addMoreGroups = Read-Host "Do you want to add '$scriptUsername' to another AD group? (Y/N)"
} while ($addMoreGroups -ieq "Y") # Use -ieq for case-insensitive comparison (Y or y)

Write-Host "Operation completed." -ForegroundColor Green