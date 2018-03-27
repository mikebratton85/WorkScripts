# Written by: Mike Bratton
# Date: March 20, 2018
# Description: Azure Multi-Factor Authentication Reset - will reset MFA or allow enabling of it for a specified user's account 

#region Variables

$groupname = Get-ADGroup "SG AAD Multi-Factor Authentication"
$username = "sa.userimac@wvholdings.onmicrosoft.com"
$SecurePassword = Get-Content C:\pscred\sa.userimac.techutil01.txt | ConvertTo-SecureString
$Creds = New-Object System.Management.Automation.PSCredential -ArgumentList $Username,$SecurePassword

#endregion Variables

#region Setup

$ErrorActionPreference = "Stop"
Clear-Host
Write-Host -BackgroundColor DarkCyan "Reset Azure Multi-Factor Authentication"
Import-Module -Force $PSScriptRoot\..\Library\log4ps -ArgumentList ".\AzureMFAReset.ps1.config"
Set-ExecutionPolicy Unrestricted –Scope Process
Write-Verbose "Importing Active Directory Module"
Import-Module ActiveDirectory
Write-Verbose "Importing Microsoft Online Module"
Import-Module MSOnline
Write-Verbose "Connecting to MSOL services"
Connect-MsolService -Credential $Creds

#endregion Setup

#region Main

#Sets MFA Authentication to enforced for any devices issued before the current date/time
$auth = New-Object -TypeName Microsoft.Online.Administration.StrongAuthenticationRequirement
$auth.RelyingParty = "*"
$auth.State = "Enforced"
$auth.RememberDevicesNotIssuedBefore = (Get-Date)

#Remove MFA from user
$UserReset = Read-Host "Please enter the user's primary SMTP address. This is not case sensitive but spelling must be accurate"
$MsolUser = Get-MsolUser -UserPrincipalName $UserReset
while($MsolUser -eq $Null){
    $UserReset = Read-Host "Invalid UPN, please check your spelling and try again"
    $MsolUser = Get-MsolUser -UserPrincipalName $UserReset
}
if($MsolUser){
    if(($msoluser.StrongAuthenticationRequirements.state -eq "Enforced") -or ($msoluser.StrongAuthenticationRequirements.state -eq "Enabled")){
        try{
            Write-Verbose "Removing Azure Multi-Factor Authentication for $($MsolUser.Displayname)"
            Set-MsolUser -UserPrincipalName $MsolUser.UserPrincipalName -StrongAuthenticationRequirements @()
        }
        catch{
            Write-Warning "Failed to reset Azure Multi-Factor Authentication for $($MsolUser.Displayname)"
            Write-Host -ForegroundColor DarkCyan "End Azure MFA User Reset Script"
            pause
            exit
        }        
    }
    else{
        Write-Warning "$($MsolUser.Displayname) has not had Azure Multi-Factor Authentication enabled!"
        $answer = Read-Host "Enable? y/n"
        while($answer -notlike "y*" -and $answer -notlike "n*"){    
            $answer = Read-Host "Invalid Selection, please type 'y' or 'n'"
        }
        if($answer -eq "n"){
            Write-Host -ForegroundColor DarkCyan "End Azure MFA User Reset Script"
            pause
            exit           
        }
        if($answer -eq "y"){
            $ADUser = Get-ADUser -f{userprincipalname -like $MsolUser.UserPrincipalName} -Properties *
            if($ADUser){
                try{
                    Write-Verbose "Attempting to add $($ADUser.name) to the 'SG AAD Multi-Factor Authentication' Security Group in AD..."                    
                    $groupname | Add-ADGroupMember -Members $ADUser -Credential $Creds
                }
                catch{
                    Write-Warning "Failed to add $($ADUser.name) to the 'SG AAD Multi-Factor Authentication' Security Group, please add manually. Error: '$($_.exception.message)'"
                }
            }
            else{Write-Warning "$($MsolUser.displayname) does not exist in AD!!"}
        }
    }   
}

#Enforce MFA for user
try{
    Write-Verbose "Enforcing Azure Multi-Factor Authentication for $($MsolUser.Displayname)"
    Set-MsolUser -UserPrincipalName $MsolUser.UserPrincipalName -StrongAuthenticationRequirements $auth
}
catch{
    Write-Warning "Failed to enforce Azure Multi-Factor Authentication for $($MsolUser.Displayname)"
}
if($MsolUser.StrongAuthenticationRequirements.state -eq "Enforced"){
    Write-Host -ForegroundColor Green "Successfully reset Azure Multi-Factor Authentication for $($MsolUser.Displayname)"
}
Write-Host -BackgroundColor DarkCyan "End Azure MFA User Reset Script"
pause

#endregion Main




