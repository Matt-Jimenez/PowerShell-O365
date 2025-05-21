# Requires the Exchange Online PowerShell module to be installed.
# You'll need to ensure you are logged into Exchange Online PowerShell
# before running this script.

do {
    # Prompt the user for the email address to check
    $EmailToCheck = Read-Host "Enter the email address to check for forwarding (or type 'exit' to quit)"

    if ($EmailToCheck -ceq "exit") {
        break # Exit the loop if the user types 'exit'
    }

    try {
        # Get the mailbox object for the specified email address
        $Mailbox = Get-Mailbox -Identity $EmailToCheck -ErrorAction Stop

        if ($Mailbox) {
            # Check if forwarding is enabled and get the forwarding address
            if ($Mailbox.ForwardingSmtpAddress) {
                Write-Host "Email address '$EmailToCheck' is currently being forwarded to: $($Mailbox.ForwardingSmtpAddress)"
                $CurrentForwarding = $Mailbox.ForwardingSmtpAddress
            } elseif ($Mailbox.ForwardingAddress) {
                # ForwardingAddress usually points to another mailbox object within the same organization
                $ForwardingMailbox = Get-Mailbox -Identity $Mailbox.ForwardingAddress -ErrorAction SilentlyContinue # Use SilentlyContinue to avoid errors if target mailbox doesn't exist
                if ($ForwardingMailbox) {
                    Write-Host "Email address '$EmailToCheck' is currently being forwarded to the internal mailbox: $($ForwardingMailbox.PrimarySmtpAddress)"
                    $CurrentForwarding = $ForwardingMailbox.PrimarySmtpAddress
                } else {
                    Write-Warning "Forwarding is configured, but the target mailbox '$($Mailbox.ForwardingAddress)' could not be found."
                    $CurrentForwarding = $null
                }
            } else {
                Write-Host "Email address '$EmailToCheck' is not currently configured for forwarding."
                $CurrentForwarding = $null
            }

            # Ask if the user wants to change the forwarding
            $ChangeForwarding = Read-Host "Do you want to change the forwarding for '$EmailToCheck'? (yes/no)"

            if ($ChangeForwarding -ceq "yes") {
                $NewForwardingAddress = Read-Host "Enter the new email address to forward to (leave blank to disable forwarding)"

                if (-not [string]::IsNullOrEmpty($NewForwardingAddress)) {
                    # Attempt to set forwarding to the new address
                    try {
                        Set-Mailbox -Identity $EmailToCheck -ForwardingSmtpAddress $NewForwardingAddress -DeliverToMailboxAndForward $true # You might want to adjust DeliverToMailboxAndForward
                        Write-Host "Forwarding for '$EmailToCheck' has been successfully changed to '$NewForwardingAddress'."
                    }
                    catch {
                        Write-Error "An error occurred while trying to set forwarding: $($_.Exception.Message)"
                    }
                } else {
                    # Disable forwarding
                    try {
                        Set-Mailbox -Identity $EmailToCheck -ForwardingSmtpAddress $null -ForwardingAddress $null -DeliverToMailboxAndForward $false
                        Write-Host "Forwarding for '$EmailToCheck' has been successfully disabled."
                    }
                    catch {
                        Write-Error "An error occurred while trying to disable forwarding: $($_.Exception.Message)"
                    }
                }
            } elseif ($ChangeForwarding -ceq "no") {
                Write-Host "No changes to forwarding were made."
            } else {
                Write-Warning "Invalid input. Please enter 'yes' or 'no'."
            }
        } else {
            Write-Warning "The email address '$EmailToCheck' was not found."
        }
    }
    catch {
        Write-Error "An error occurred: $($_.Exception.Message)"
        Write-Error "Make sure you are connected to Exchange Online PowerShell with the necessary permissions."
    }

} while ($true) # Loop indefinitely until the 'break' statement is encountered

Write-Host "Script finished."