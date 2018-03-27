# Script Written by Mike Bratton
# January 15, 2018
#
# CRM REMOVAL TOOL
#
# This script will import a prepared CSV report of users that need to be removed from CRM Licensing
# The report must contain fullname, samaccountname, or userprincipalname for each user
# The users will be filtered through Office 365 and the CRM licenses will be removed
# The associated CRM AD Groups will also be removed
# Fail/Success will be logged on TECHUTIL01
#
#  ***Consider running FilterOutDisabled.ps1 first to obtain a CSV report of only disabled users from original report***
#
# --------------------------------------------------------------------------------------------------------------------------------------------------






$dateFormat = "$($(Get-Date).ToString('yyyy-MM-dd_hh-mm-ss'))"
$logfile = "\\techutil01\c$\logs\RemovedCRMLicenses-$($dateFormat).txt"

function Log-It ($Message){

Write-Host -ForegroundColor Cyan $Message
"$($(Get-Date).ToString('yyyy-MM-dd::hh:mm:ss')) | $Message" | Out-File $logfile -Append


}

function Log-Success ($message){

    Write-Host -ForegroundColor Green $Message 
    "$($(Get-Date).ToString('yyyy-MM-dd::hh:mm:ss')) | $Message" | Out-File $logfile -Append

}

function Log-Failure ($message){

    Write-Host -ForegroundColor Red $Message
    "$($(Get-Date).ToString('yyyy-MM-dd::hh:mm:ss')) | $Message" | Out-File $logfile -Append

}

Function Get-FileName($initialDirectory){  
        [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") |
        Out-Null
        $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
        $OpenFileDialog.initialDirectory = $initialDirectory
        $OpenFileDialog.filter = "All files (*.*)| *.*"
        $OpenFileDialog.ShowDialog() | Out-Null
        $OpenFileDialog.filename
}

function Remove-CRMLicensing {
    [CmdletBinding()]
    Param(
        [Parameter(
            Mandatory = $true,
            Position = 0,
            ValueFromPipeline=$true
        )]
        [ValidateNotNullOrEmpty()]
        $User
    )

    begin{
        $FN = "Remove-CRMLicensing"
        $NotRemoved = @()
        $Removed = @()
    }

    process{
        $Dynamics = "SG 365 License Dynamics"
        $Basic = "SG 365 License CRM Basic"
        $Pro = "SG 365 License CRM Pro"
        $SKU = "wvholdings:DYN365_ENTERPRISE_PLAN1"
        if(($user.islicensed -eq $true) -and ($user.licenses.accountskuid.contains($SKU))){
            Log-It "[OK] $FN | Trying to remove CRM licensing for $($User.displayname)"            
            try{ 
                $user | Set-MsolUserLicense -RemoveLicense $SKU -erroraction stop
                Log-Success "[SUCCESS] $FN | Removed CRM license from $($user.displayname)"
                $Removed += $user
            }
            catch{ 
                Log-Failure "[ERROR] $FN | Failed to remove CRM license from $($user.displayname)"
                Write-Host -ForegroundColor DarkYellow $error[0].exception.message

                $NotRemoved += $user
            }
        }
        else{
            Log-It "[OK] $FN | $($User.displayname) is not licensed with CRM"
        }
        $aduser = Get-ADUser -f "name -like '$($user.displayname)*'" -Properties *
        $CRMDynamics = Get-ADGroupMember -Identity $Dynamics
        if($CRMDynamics.name -contains $aduser.name){
            Log-It "[OK] $FN | Trying to remove $($aduser.name) from the CRM Dynamics AD Group"
            try{
                Remove-ADGroupMember -Identity $Dynamics -members $aduser -Confirm:$false -ErrorAction stop
                Log-Success "[SUCCESS] $FN | Removed $($aduser.name) from CRM Dynamics AD group"
            }
            catch{
                Log-Failure "[ERROR] $FN | Failed to remove $($aduser.name) from CRM Dynamics AD Group"
                Write-Host -ForegroundColor DarkYellow $error[0].exception.message
            }
        }
        $CRMBasic = Get-ADGroupMember -Identity $Basic        
        if($CRMBasic.name -contains $aduser.name){
            Log-It "[OK] $FN | Trying to remove $($aduser.name) from the CRM Basic AD Group"
            try{
                Remove-ADGroupMember -Identity $Basic -members $aduser -Confirm:$false -ErrorAction stop
                Log-Success "[SUCCESS] $FN | Removed $($aduser.name) from CRM Basic AD group"
            }
            catch{
                Log-Failure "[ERROR] $FN | Failed to remove $($aduser.name) from CRM Basic AD Group"
                Write-Host -ForegroundColor DarkYellow $error[0].exception.message
            }
        }
        $CRMPro = Get-ADGroupMember -Identity $Pro
        if($CRMPro.name -contains $aduser.name){
            Log-It "[OK] $FN | Trying to remove $($aduser.name) from the CRM Pro AD Group"
            try{
                Remove-ADGroupMember -Identity $Pro -members $aduser -Confirm:$false -ErrorAction stop
                Log-Success "[SUCCESS] $FN | Removed $($aduser.name) from CRM Pro AD group"
            }
            catch{
                Log-Failure "[ERROR] $FN | Failed to remove $($aduser.name) from CRM Pro AD Group"
                Write-Host -ForegroundColor DarkYellow $error[0].exception.message
            }
        }
    }

    end{ 
        return @{Removed=$Removed.count;NotRemoved=$NotRemoved.count} 
    }



}

Read-Host "Please select the CSV file to import"

$filename = Get-FileName

$CRMUsers = Import-CSV $filename

$username = "sa.userimac@wvholdings.onmicrosoft.com"
$password = "Et2OFw4ptlr&"
$secureStringPwd = $password | ConvertTo-SecureString -AsPlainText -Force 
$creds = New-Object System.Management.Automation.PSCredential -ArgumentList $username, $secureStringPwd

Connect-MsolService -Credential $Creds

$ToUnlicense = @()

foreach($user in $CRMUsers){

    $aduser = Get-Aduser -f "userprincipalname -like '*$($user.email)*'" -Properties *

    $aduser.name

    $UserToUnlicense = Get-msoluser -UserPrincipalName $($aduser.userprincipalname)

    $ToUnlicense += $UserToUnlicense
}

$ToUnlicense | Remove-CRMLicensing









