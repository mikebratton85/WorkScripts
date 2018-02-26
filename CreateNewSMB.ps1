
# Variables
$ad01 = "ad01.worldventures.local"
$ad02 = "ad02.worldventures.local"
$cas02 = "CAS02.WorldVentures.local"
$SMBOU = "WorldVentures.local/WorldVentures/Shared Mailboxes"
$SGOU = "OU=Security Groups,OU=WorldVentures,DC=WorldVentures,DC=local"

# Changing Execution Policy
Set-ExecutionPolicy RemoteSigned

# Add Exchange On-Prem Snapin
Write-Host "Adding Exchange On-Prem Snapin"-ForegroundColor Cyan
Add-PSSnapin Microsoft.Exchange.Management.PowerShell.SnapIn;

# Set Exch to CAS02 and DC to AD01
Write-Host "Setting Domain Controller to AD01" -ForegroundColor Cyan
Set-ExchangeServer -Identity $cas02 -DomainController $ad01

# Enter Credentials
Write-Host 'Prompting for Office 365 Credentials'-ForegroundColor Cyan
$Credential = Get-Credential

# Add Alias, Domain, and Display Name for SMB (Spaces are ok in display name)
$SMBAlias = Read-Host -Prompt 'Shared Mailbox Alias (part before the @)'
$SMBDomain = Read-Host -Prompt 'Shared Mailbox Domain (e.g. wvholdings.com)'
$SMBDisplayName = Read-Host -Prompt 'Shared Mailbox Display Name'

# Creates AD/Exch Remote Mailbox
Write-Host "Creating AD User/Remote Mailbox" -ForegroundColor Cyan
New-RemoteMailbox -Name $SMBDisplayName -UserPrincipalName "$SMBAlias@$SMBDomain".ToLower() -Alias "$SMBAlias".ToLower() -OnPremisesOrganizationalUnit $SMBOU -Equipment

# Creates AD/Exch SG
Write-Host "Creating Exchange Security Group for SMB Delegation Permissions" -ForegroundColor Cyan
New-DistributionGroup -Name "SG Exchange Mailbox $SMBAlias" -Type Security -OrganizationalUnit $SGOU

# Allow time for SG to create, then hide from GAL
Start-Sleep -s 15
Set-DistributionGroup -Identity "SG Exchange Mailbox $SMBAlias" -HiddenFromAddressListsEnabled:$true

# Replicating changes to all DCs
Write-Host "Replicating DCs and waiting 1 minute" -ForegroundColor Cyan
repadmin /replicate $ad01 $ad02 dc=worldventures,dc=local
repadmin /replicate $ad02 $ad01 dc=worldventures,dc=local
Start-Sleep -s 60

# Running Azure sync to update cloud
Write-Host "Starting Azure AD Sync Cycle and waiting 5 minutes" -ForegroundColor Cyan
Start-ADSyncSyncCycle -PolicyType Delta
Start-Sleep -s 300

# Remove Exchange On-Prem Snapin
Write-Host "Removing Exchange On-Prem Snapin" -ForegroundColor Cyan
Remove-PSSnapin Microsoft.Exchange.Management.PowerShell.SnapIn;

# Connect to EOL
Write-Host "Connecting to Exchange Online" -ForegroundColor Cyan
$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $Credential -Authentication Basic -AllowRedirection
Import-PSSession $Session

# Convert to SMB from Remote and Delegate Access
Write-Host "Converting Mailbox to Shared & Delegating Access" -ForegroundColor Cyan
Set-Mailbox -Identity $SMBAlias@$SMBDomain -Type Shared
Add-MailboxPermission -AccessRights FullAccess -Identity $SMBAlias@$SMBDomain -AutoMapping $true -User "SG Exchange Mailbox $SMBAlias"
Add-RecipientPermission -Identity $SMBAlias@$SMBDomain -AccessRights SendAs -Trustee "SG Exchange Mailbox $SMBAlias" -Confirm:$false

# Disconnect from EOL
Write-Host "Removing Exchange Online PowerShell Session" -ForegroundColor Cyan
Remove-PSSession $Session


Write-Host "Setting On-Prem Mailbox attributes to be that of a Remote Shared Mailbox" -ForegroundColor Cyan
# This must be done AFTER the mailbox is provisioned/synced in Exchange Online
Set-ADUser -Identity "$SMBAlias"  -Replace @{msExchRemoteRecipientType="100";msExchRecipientTypeDetails="34359738368"}

# Success Message
Write-Host "Completed. Add users to SG Exchange Mailbox $SMBAlias in Active Directory to grant access to the mailbox. It can take a couple of hours before the user will see the mailbox appear in Outlook." -ForegroundColor Cyan