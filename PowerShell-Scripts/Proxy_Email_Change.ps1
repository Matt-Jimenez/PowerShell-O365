# Function to output text with color for Windows PowerShell ISE
function Write-IseColoredText {
    param(
        [Parameter(Mandatory=$true)]
        [string]$Text,
        [Parameter(Mandatory=$true)]
        [string]$Color # Valid values: Red, Green
    )
    switch ($Color.ToLower()) {
        "red"   { Write-Host $Text -ForegroundColor Red }
        "green" { Write-Host $Text -ForegroundColor Green }
        default { Write-Host $Text }
    }
}

# Prompt the user for the username or SAMAccountName
Write-Host "Enter the username: " -NoNewline
$username = Read-Host

# Get the current proxy addresses of the user
$user = Get-ADUser -Identity $username -Properties proxyAddresses

# Display the existing proxy addresses
Write-Host "Current proxy addresses for $($user.SamAccountName):"
$foundProxy = $false
if ($user.proxyAddresses) {
    foreach ($proxy in $user.proxyAddresses) {
        if ($proxy -notlike "X500:*") {
            Write-IseColoredText "$proxy" "green"
            $foundProxy = $true
        }
    }
    if (-not $foundProxy) {
        Write-Host "No non-X500 proxy addresses found for this user."
    }
} else {
    Write-Host "No proxy addresses found for this user."
}

# Prompt the user for the local part of the new SMTP address
Write-Host "Enter the local part of the new SMTP address (e.g., newuser): " -NoNewline
$localPart = Read-Host

# Determine the SMTP prefix (SMTP or smtp)
$smtpPrefix = ""
while ($true) {
    Write-Host "Should the prefix be 'SMTP' or 'smtp'? (Enter SMTP or smtp): " -NoNewline
    $prefixInput = Read-Host
    if ($prefixInput -ceq "SMTP" -or $prefixInput -ceq "smtp") {
        $smtpPrefix = $prefixInput + ":"
        break
    } else {
        Write-IseColoredText "Invalid input. Please enter 'SMTP' or 'smtp'." "red"
    }
}

# Construct the new SMTP proxy address
$newSmtpAddress = $smtpPrefix + $localPart

# Confirmation prompt
Write-Host "You are about to add the proxy address: '$newSmtpAddress'. Confirm? (Y/N): " -NoNewline
$confirmation = Read-Host
if ($confirmation.ToLower() -ne "y") {
    Write-Host "Operation cancelled by user."
    exit
}

# Create an array to hold the updated proxy addresses
$updatedProxyAddresses = @()

# Add the existing proxy addresses (excluding x500) to the new array
if ($user.proxyAddresses) {
    foreach ($proxy in $user.proxyAddresses) {
        if ($proxy -notlike "X500:*") {
            $updatedProxyAddresses += $proxy
        }
    }
}

# Add the new SMTP proxy address to the array
$updatedProxyAddresses += $newSmtpAddress

# Set the updated proxyAddresses attribute for the user
$setResult = Set-ADUser -Identity $username -Replace @{proxyAddresses = $updatedProxyAddresses}

# Output after attempting to set the proxy address
Write-Host "`nAttempting to add proxy address '$newSmtpAddress'..."
if ($?) {
    Write-IseColoredText "Successfully added proxy address '$newSmtpAddress'." "green"
} else {
    Write-IseColoredText "Failed to add proxy address '$newSmtpAddress'." "red"
}

# Display the new list of proxy addresses
$updatedUser = Get-ADUser -Identity $username -Properties proxyAddresses
Write-Host "`nNew proxy addresses for $($updatedUser.SamAccountName):"
$foundNewSmtpInOutput = $false
if ($updatedUser.proxyAddresses) {
    foreach ($proxy in $updatedUser.proxyAddresses) {
        if ($proxy -notlike "X500:*") {
            if ($proxy -ceq $newSmtpAddress) {
                Write-IseColoredText "$proxy" "green"
                $foundNewSmtpInOutput = $true
            } else {
                Write-Host "$proxy"
            }
        }
    }
    if ($foundNewSmtp -and -not $foundNewSmtpInOutput) {
        Write-Host "Warning: The newly added SMTP address was not found in the updated list." -ForegroundColor Red
    }
} else {
    Write-Host "No proxy addresses found after the update." -ForegroundColor Red
}