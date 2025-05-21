# Ensure the Active Directory module is imported
Import-Module ActiveDirectory

# Prompt for user details
$FirstName = Read-Host "Enter First Name"
$LastName = Read-Host "Enter Last Name"
$UserLogon = Read-Host "Enter Username (User Logon)"

# Use default password for new users
$DefaultPasswordPlain = "8en=sWORuc1oTohu"
$Password = ConvertTo-SecureString $DefaultPasswordPlain -AsPlainText -Force

# Construct values
$DisplayName = "$FirstName $LastName"
$Name = "$FirstName $LastName"
$SamAccountName = $UserLogon
$UserPrincipalName = "$UserLogon@aus.com"
$OU = "OU=Mexico,OU=AUS,DC=aus,DC=com"

# Groups to add user to
$Groups = @(
    "Kiosk License Assignment",
    "Mimecast Awareness Training",
    "MXGoogleChromebookProvisioning"
)

try {
    # Check if user already exists
    if (Get-ADUser -Filter {SamAccountName -eq $SamAccountName}) {
        Write-Host "User '$SamAccountName' already exists in AD." -ForegroundColor Yellow
    } else {
        # Create the AD User
        New-ADUser `
            -Name $Name `
            -GivenName $FirstName `
            -Surname $LastName `
            -DisplayName $DisplayName `
            -SamAccountName $SamAccountName `
            -UserPrincipalName $UserPrincipalName `
            -AccountPassword $Password `
            -PasswordNeverExpires $false `
            -ChangePasswordAtLogon $true `
            -Enabled $true `
            -Path $OU

        Write-Host "User '$Name' created successfully." -ForegroundColor Green

        # Add user to groups
        foreach ($Group in $Groups) {
            try {
                Add-ADGroupMember -Identity $Group -Members $SamAccountName
                Write-Host "Added '$SamAccountName' to group '$Group'." -ForegroundColor Cyan
            } catch {
                Write-Host "Failed to add '$SamAccountName' to group '$Group'. Error: $_" -ForegroundColor Red
            }
        }
    }
}
catch {
    Write-Host "Failed to create user '$Name'. Error: $_" -ForegroundColor Red
}
