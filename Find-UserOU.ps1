# Script to find current user's OU and search for other OUs.
# Requires the Active Directory module (part of Remote Server Administration Tools - RSAT).

# Function to convert an AD distinguished name path (or part of it) to a slash-separated path style
function Convert-ADPathToSlashPath {
    param(
        [string]$ADPath # e.g., OU=Sales,OU=Departments,DC=example,DC=com or just OU=Sales,OU=Departments
    )

    $components = $ADPath -split ','
    $ouNames = @()

    foreach ($component in $components) {
        if ($component -match '^OU=') {
            $ouNames += $component -replace '^OU=', ''
        }
        # We are focusing on OU paths as per the example.
        # If you also wanted to include CNs for containers like CN=Users, you could add:
        # elseif ($component -match '^CN=') {
        #    $ouNames += $component -replace '^CN=', '' # This would make it a generic path converter
        # }
    }

    if ($ouNames.Count -gt 0) {
        [array]::Reverse($ouNames) # Reverse to get top-down path
        return $ouNames -join '\'
    } else {
        return $null # Return null if no OU components found (e.g. for DC=, CN= paths)
    }
}

# Attempt to import the Active Directory module
Import-Module ActiveDirectory -ErrorAction SilentlyContinue
if (-not (Get-Module ActiveDirectory)) {
    Write-Error "Active Directory module is not available. Please ensure Remote Server Administration Tools (RSAT) are installed and the AD DS Tools feature is enabled."
    Write-Error "You can typically install RSAT via 'Settings > Apps > Optional features' or using DISM."
    exit 1
}

# Get current logged-in user's information
try {
    Write-Host "Fetching your Active Directory user information..."
    $currentUser = Get-ADUser -Identity $env:USERNAME -Properties DistinguishedName -ErrorAction Stop
} catch {
    Write-Error "Failed to retrieve current user information from Active Directory: $($_.Exception.Message)"
    Write-Error "Please ensure you are running this script on a domain-joined machine and have permissions to query Active Directory."
    exit 1
}

# Extract the parent container (OU or other container) path from the DistinguishedName
$userDN = $currentUser.DistinguishedName
$parentContainerDN = ($userDN -split ',', 2)[1] # Splits at the first comma and takes the rest

if ($parentContainerDN) {
    $slashPath = Convert-ADPathToSlashPath -ADPath $parentContainerDN
    if ($parentContainerDN -match '^OU=') {
        Write-Host "You are currently in the Organizational Unit (OU):"
        Write-Host "  DistinguishedName: $parentContainerDN"
        if ($slashPath) {
            Write-Host "  Path style:        $slashPath"
        }
    } else {
        Write-Host "You are currently in the container:"
        Write-Host "  DistinguishedName: $parentContainerDN"
        # Optionally, if Convert-ADPathToSlashPath was modified to handle CNs, you could show a path here too.
    }
} else {
    Write-Warning "Could not determine your current organizational location from DistinguishedName: $userDN"
}

Write-Host "" # Add a blank line for readability

# Ask if the user wants to search for a specific OU
$response = Read-Host "Do you need to search for a specific OU? (Enter 'yes' or 'no')"

if ($response -eq 'yes') {
    $targetOUName = Read-Host "Enter the name (or part of the name) of the OU you are looking for (e.g., 'Sales' or 'IT Support')"

    if (-not ([string]::IsNullOrWhiteSpace($targetOUName))) {
        Write-Host "Searching for OUs with name containing '$targetOUName'..."
        try {
            # Search for OUs where the Name property is like the input. Wildcards are added for partial matches.
            $foundOUs = Get-ADOrganizationalUnit -Filter "Name -like '*$targetOUName*'" -Properties DistinguishedName -ErrorAction Stop

            if ($foundOUs) {
                Write-Host "Found the following OU(s):"
                $foundOUs | ForEach-Object {
                    Write-Host "- Name: $($_.Name)"
                    Write-Host "  DistinguishedName: $($_.DistinguishedName)"
                    $ouSlashPath = Convert-ADPathToSlashPath -ADPath $_.DistinguishedName
                    if ($ouSlashPath) {
                        Write-Host "  Path style:        $ouSlashPath"
                    }
                    Write-Host "" # Blank line between OUs
                }
            } else {
                Write-Host "No OUs found matching '*$targetOUName*'."
            }
        } catch {
            Write-Warning "An error occurred while searching for OUs: $($_.Exception.Message)"
            Write-Host "This could be due to insufficient permissions or if the OU doesn't exist."
        }
    } else {
        Write-Host "No OU name entered. Skipping search."
    }
} elseif ($response -eq 'no') {
    Write-Host "OU search not requested."
} else {
    Write-Host "Invalid response: '$response'. Please answer 'yes' or 'no' next time. Exiting."
}

Write-Host ""
Write-Host "Script finished."