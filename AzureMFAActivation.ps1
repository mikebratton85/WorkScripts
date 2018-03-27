# Written by: Mike Bratton
# Date: March 13, 2018

# Description: Azure Multi-Factor Authentication Activation and Enforcement/Removal - will pull anyone 
# from the "SG AAD Multi-Factor Authentication" Security Group and Enforce MFA on their account.
# Likewise, it will check against AD and anyone not in the group will have Multi-Factor 
# Authentication disabled.

#region Variables
$groupname = Get-ADGroup "SG AAD Multi-Factor Authentication"
$username = "sa.userimac@wvholdings.onmicrosoft.com"
$password = "Et2OFw4ptlr&"
$secureStringPwd = $password | ConvertTo-SecureString -AsPlainText -Force 
$Creds = New-Object System.Management.Automation.PSCredential -ArgumentList $username, $secureStringPwd
#endregion Variables

#region Setup
Import-Module -Force $PSScriptRoot\..\Library\log4ps -ArgumentList ".\AzureMFAActivation.ps1.config"
Set-ExecutionPolicy Unrestricted –Scope Process
Write-Verbose "Initializing Azure MFA Activation Scan"
Write-Verbose "Stored credentials for $($Creds.username)"
Import-Module ActiveDirectory -ErrorAction 'Stop'
Write-Verbose "Imported Active Directory Module"
Import-Module MSOnline -ErrorAction 'Stop'
Write-Verbose "Imported Microsoft Online Module"
Connect-MsolService -Credential $Creds -WarningAction SilentlyContinue -ErrorAction 'Stop'
Write-Verbose "Connected to MSOL services"
#endregion Setup

#region Main
# Sets MFA Authentication to enforced for any devices issued before the current date/time
$auth = New-Object -TypeName Microsoft.Online.Administration.StrongAuthenticationRequirement
$auth.RelyingParty = "*"
$auth.State = "Enforced"
$auth.RememberDevicesNotIssuedBefore = (Get-Date)
# Pull all users from the "SG AAD Multi-Factor Authentication" Security Group in AD
Write-Verbose "Pulling users from 'SG AAD Multi-Factor Authentication' Security Group"
$NotEnforced = @()
$MFAGroupMembers = Get-ADGroupMember $groupname -Recursive
if($MFAGroupMembers){
    $MFAUsers = $MFAGroupMembers.samaccountname
    $MFAUsers | ForEach-Object {
        $ADUser = Get-ADUser $_
        $MSOLUser = Get-MsolUser -UserPrincipalName $ADUser.UserPrincipalName -ErrorAction SilentlyContinue
        if($MSOLUser){
            if($MSOLUser.StrongAuthenticationRequirements.state -ne "Enforced"){                
                # Enables MFA for each user in the "SG AAD Multi-Factor Authentication" Security Group in AD
                try{
                    Set-MsolUser -UserPrincipalName $ADUser.UserPrincipalName -StrongAuthenticationRequirements $auth -ErrorAction Stop
                    Write-Verbose "Successfully enforced Azure Multi-Factor Authentication for $($ADUser.name)"
                    $NotEnforced += $MSOLUser
                }
                catch{Write-Warning "Failed to enforce Azure Multi-Factor Authentication for $($ADUser.name)"}
            }
        }
        else{Write-Warning "MSOL user $($ADUser.UserPrincipalName) not found"}
    }
}
else{Write-Verbose "There are no members of the 'SG AAD Multi-Factor Authentication' Security Group in AD"}
if($NotEnforced){Write-Verbose "$($NotEnforced.count) new users have had Azure Multi-Factor Authentication enforced"}
else {Write-Verbose "No new users have been added to the 'SG AAD Multi-Factor Authentication' Security Group in AD"}
# Remove MFA for any user not in the "SG AAD Multi-Factor Authentication" Security Group in AD
Write-Verbose "Checking for users removed from 'SG AAD Multi-Factor Authentication' Security Group"
$RemovedEnforce = @()
$msolusers = Get-MsolUser -EnabledFilter EnabledOnly -All
$msolusers | ForEach-Object {
    $msolcheck = $_
    if($msolcheck.LastDirSyncTime -ne $null){
        if(($msolcheck.StrongAuthenticationRequirements.state -eq "Enabled") -or ($msolcheck.StrongAuthenticationRequirements.state -eq "Enforced")){
            $ADcheck = Get-ADUser -f{UserPrincipalName -like $msolcheck.UserPrincipalName} -Properties *
            if($ADcheck){
                if($ADcheck.memberof -notcontains $groupname){
                    # Sets MFA Authentication to Disabled
                    $auth = @()
                    try{
                        Set-MsolUser -UserPrincipalName $msolcheck.UserPrincipalName -StrongAuthenticationRequirements $auth -ErrorAction stop
                        Write-Verbose "Successfully removed Azure Multi-Factor Authentication for $($msolcheck.Displayname)"
                        $RemovedEnforce += $msolcheck
                    }
                    catch{Write-Warning "Failed to remove Azure Multi-Factor Authentication for $($msolcheck.name)"}
                }
            }
            else{
                try{
                    Set-MsolUser -UserPrincipalName $msolcheck.UserPrincipalName -StrongAuthenticationRequirements $auth -ErrorAction stop
                    Write-Verbose "$($msolcheck.UserPrincipalName) does not exist in AD! Successfully removed Azure Multi-Factor Authentication"
                }
                catch{Write-Warning "$($msolcheck.UserPrincipalName) does not exist in AD! Failed to remove Azure Multi-Factor Authentication"}
            }
        }
    }
}
if($RemovedEnforce){Write-Verbose "$($RemovedEnforce.count) users have had Azure Multi-Factor Authentication removed"}
else {Write-Verbose "No users have have had Azure Multi-Factor Authentication removed"}
Write-Verbose "End Azure MFA Activation Scan"
#endregion Main