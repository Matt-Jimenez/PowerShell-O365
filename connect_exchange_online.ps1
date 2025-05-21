# Prompt the user for credentials
$UserCredential = Get-Credential -Message "Enter your Office 365 administrator credentials"

# Connect to Exchange Online using the provided credentials
Connect-ExchangeOnline -Credential $UserCredential

# Optional: Confirm connection
if ($?) {
    Write-Host "Connected to Exchange Online successfully."
} else {
    Write-Host "Failed to connect to Exchange Online."
}