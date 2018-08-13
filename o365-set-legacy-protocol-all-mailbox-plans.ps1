$ErrorActionPreference = "Stop";
# Set IMAP, POP3, and ActiveSync for all mailboxes plans on a Office 365 tenant
# This will set any FUTURE accounts created under these mailboxes with the same settings
# Note you can't disable MAPI on mailbox plans. The documentation states "This parameter is reserved for internal Microsoft use."

# For accounts with MFA you need to install the Exchange Online Remote PowerShell Module
# https://docs.microsoft.com/en-us/powershell/exchange/exchange-online/connect-to-exchange-online-powershell/mfa-connect-to-exchange-online-powershell?view=exchange-ps

### Script config ###
# Multi-factor authentication
# Set to $true if you're using MFA for your administrator accounts (you should be)
$mfa=$true

# Set your administrator account below
$userPrincipalName="shane@mallia.onmicrosoft.com"

# Enable/Disable IMAP
$imapEnabled=$false

# Enable/Disable POP
$popEnabled=$false

# Enable/Disable ActiveSync
$activeSyncEnabled=$false

### End of script config ###

if ($mfa) {

    Try {
        #Connect to O365 Exchange with MFA
        $Session = Connect-EXOPSSession -UserPrincipalName $userPrincipalName -ConnectionUri “https://ps.outlook.com/powershell"
    } 
    Catch [System.Management.Automation.CommandNotFoundException] {
        Write-Output "Please ensure you have the Exchange Online Remote PowerShell Module installed"
        Write-Output "Visit https://docs.microsoft.com/en-us/powershell/exchange/exchange-online/connect-to-exchange-online-powershell/mfa-connect-to-exchange-online-powershell?view=exchange-ps"
        Exit
    }
    Catch {
        $ErrorMessage = $_.Exception.Message
        $FailedItem = $_.Exception.ItemName
        Write-Output "Failed to connect to O365 Exchange"
        Write-Output "Error Message : $ErrorMessage"
        throw $_
    }

} else {
    
    Try {
        # Connect to O365 without MFA
        $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri “https://ps.outlook.com/powershell" -Credential $userPrincipalName -Authentication Basic -AllowRedirection
    }
    Catch [System.Management.Automation.CommandNotFoundException] {
        Write-Output "Please ensure you have the MSOnline module installed"
        Write-Ouptut "Run \"Install-Module MSOnline\" in elevated PowerShell"
        Exit
    }
    Catch {
        $ErrorMessage = $_.Exception.Message
        $FailedItem = $_.Exception.ItemName
        Write-Output "Failed to connect to O365 Exchange"
        Write-Output "Error Message : $ErrorMessage"
        throw $_
    }

}

$mailboxPlans =  Get-CASMailboxPlan -ResultSize Unlimited 

Write-Output "Setting all $($mailboxPlans.Count) mailbox plans with the following settings:"
Write-Output $(if($imapEnabled) {"Enabling IMAP"} else {"Disabling IMAP"})
Write-Output $(if($popEnabled) {"Enabling POP"} else {"Disabling POP"})
Write-Output $(if($activeSyncEnabled) {"Enabling ActiveSync"} else {"Disabling ActiveSync"})

Read-Host -Prompt "Press any key to continue or CTRL+C to quit" 

ForEach ($mailboxPlan in $mailboxPlans) {
    Write-Output "Setting mailbox plan $($mailboxPlan.Identity))"
    $mailboxPlan |  Set-CASMailboxPlan -ImapEnabled $imapEnabled -PopEnabled $popEnabled -ActiveSyncEnabled $activeSyncEnabled
}
