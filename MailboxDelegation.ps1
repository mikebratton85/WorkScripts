# ---------------------------------------------------------
# Mailbox_Delegation.ps1
# ---------------------------------------------------------
# Author: Mike Bratton
# ---------------------------------------------------------
#region Setup Variables
# ---------------------------------------------------------

$logfile = "\\techutil01\c$\testlog\mailbox_delegate.txt"

$Global:ErrorActionPreference = 'Stop'


# ---------------------------------------------------------
#endregion Setup Variables
# ---------------------------------------------------------
#region Setup Function
# ---------------------------------------------------------


Function Test-ADPassword($Creds){
    $Username = $Creds.username
    $Password = $Creds.GetNetworkCredential().password 
     #Get current domain using logged-on user's credentials
    $CurrentDomain = "LDAP://" + ([ADSI]"").distinguishedName
    $Domain = New-Object System.DirectoryServices.DirectoryEntry($CurrentDomain,$Username,$Password)
    $DomainName = $Domain.name
    if ($Domain.name -eq $null){
        Log-Failure "[FATAL] | Authentication failed for $Username"
        Exit
    }
    else{
        Log-Success "[SUCCESS] | Successfully authenticated $username"
        return $True
    }
    $Password = $null
}

function Get-MainMenu {

cls

echo ""
echo ""
echo ""
echo ""
echo ""
echo ""
echo ""
echo ""
echo "                             _________________________________________________________"
echo "                             |-------------------------------------------------------|" 
echo "                             |                   MAILBOX DELEGATION                  |"
echo "                             |-------------------------------------------------------|"
echo "                             |                                                       |"
echo "                             |                                                       |"
echo "                             |                                                       |"
echo "                             |   1. Get User Mailbox Delegate Info                   |"
echo "                             |   2. Delegate Full Access to a Mailbox                |"
echo "                             |   3. Remove Full Access to a Mailbox                  |"
echo "                             |   4. Exit                                             |"
echo "                             |                                                       |"
echo "                             |                                                       |"
echo "                             |                                                       |"
echo "                             |-------------------------------------------------------|"
echo ""
echo ""
echo ""

$answer = Read-Host "                                            Please Make a Selection"

if($answer -eq 1){Get-DelegateSearchMenu}
if($answer -eq 2){Get-AddSearchMenu}
if($answer -eq 3){Get-RemoveSearchMenu}
if($answer -eq 4){Get-AreYouSure}

else{Write-Host -ForegroundColor Red "Invalid Selection, Try Again." 
    sleep 2
    Get-MainMenu
}

}

function Get-AddSearchMenu {

    cls

    $searchName = Read-Host "Who's mailbox do you want to add a delegate to? Type 'q' to return to main menu.
"

    if($searchName -eq "q"){Get-MainMenu}

    elseif($searchName){Find-User -search $searchName}

    else{

        cls
        
        Write-Host -ForegroundColor DarkYellow "You didn't type anything weirdo"

        Sleep 2

        Get-AddSearchMenu
    
    }

}

function Get-RemoveSearchMenu {

    cls

    $searchName = Read-Host "Who's mailbox do you want to remove a delegate from? Type 'q' to return to main menu.
"

    if($searchName -eq "q"){Get-MainMenu}

    elseif($searchName){Find-RemoveUser -search $searchName}

    else{

        cls
        
        Write-Host -ForegroundColor DarkYellow "You didn't type anything weirdo"

        Sleep 2

        Get-RemoveSearchMenu
    
    }

}

function Get-DelegateSearchMenu {

    cls

    $searchName = Read-Host "Who's mailbox do you want to find delegates for? Type 'q' to return to main menu.
"

    if($searchName -eq "q"){Get-MainMenu}

    elseif($searchName){Find-UserDelegates -search $searchName}

    else{

        cls
        
        Write-Host -ForegroundColor DarkYellow "You didn't type anything weirdo"

        Sleep 2

        Get-AddSearchMenu
    
    }  

}

function Find-User ($search){

    $ADUsers = Get-ADUser -f "name -like '*$search*'" -Properties * | sort -Unique

    if($ADUsers.count -gt 1){
    
        $ADUsers.ForEach({

            "$($ADUsers.IndexOf($_) + 1) | $($_.name)"

        })

        $answer = $null
 
        $answer = Read-Host "Select a name by typing a number"
        
        while((($answer -match [regex]"^\d*$") -eq $False) -or ([int]$answer -gt $ADUsers.Count) -or ([int]$answer -eq 0) -or ($answer -eq "")){

            Write-Host -ForegroundColor Red "Invalid Selection, Try Again"
            
            $answer = Read-Host -Prompt "Select a name by typing a number"
        
        }     

        $answer = [int]$answer

        if($answer){
           
            $selectedName = $ADUsers[$answer - 1]

            $nameAnswer = Read-Host "You have selected $($selectedName.name), is this correct? y/n"

            while($nameAnswer -notlike "y*" -and $nameAnswer -notlike "n*"){
    
                Write-Host -ForegroundColor Red "Invalid Selection, Try Again"

                $nameAnswer = Read-Host "You have selected $($selectedName.name), is this correct? y/n"

            }

            if($nameAnswer -eq "y"){

                Log-Only "[INFO] | Selected user $($selectedName.name)"
            
                Add-MailboxDelegate -user $selectedName
                
            }
    
            elseif($nameAnswer -eq "n"){Get-AddSearchMenu}

        }
       
    }

    elseif($ADUsers.count -eq 0){

        Write-Host -ForegroundColor Red "USER NOT FOUND!"

        Sleep 2

        Get-AddSearchMenu

    }

    else{
    
        $answer = Read-Host "$($ADUsers.name) is the only user matching your search, is this the correct user? y/n"

        while($answer -notlike "y*" -and $answer -notlike "n*"){
    
        Write-Host -ForegroundColor Red "Invalid Selection, Try Again"

        sleep 2

        $answer = Read-Host "$($ADUsers.name) is the only user matching your search, is this the correct user? y/n"

        }

        if($answer -eq "y"){
        
            Log-Only "[INFO] | Selected user $($selectedName.name)"
            
            Add-MailboxDelegate -user $ADUsers
            
        }
    
        elseif($answer -eq "n"){Get-AddSearchMenu}
    
    }

}

function Find-RemoveUser ($search){

    $ADUsers = Get-ADUser -f "name -like '*$search*'" -Properties * | sort -Unique

    if($ADUsers.count -gt 1){
    
        $ADUsers.ForEach({

            "$($ADUsers.IndexOf($_) + 1) | $($_.name)"

        })

        $answer = $null
 
        $answer = Read-Host "Select a name by typing a number"
        
        while((($answer -match [regex]"^\d*$") -eq $False) -or ([int]$answer -gt $ADUsers.Count) -or ([int]$answer -eq 0) -or ($answer -eq "")){

            Write-Host -ForegroundColor Red "Invalid Selection, Try Again"
            
            $answer = Read-Host -Prompt "Select a name by typing a number"
        
        }     

        $answer = [int]$answer

        if($answer){
           
            $selectedName = $ADUsers[$answer - 1]

            $nameAnswer = Read-Host "You have selected $($selectedName.name), is this correct? y/n"

            while($nameAnswer -notlike "y*" -and $nameAnswer -notlike "n*"){
    
                Write-Host -ForegroundColor Red "Invalid Selection, Try Again"

                $nameAnswer = Read-Host "You have selected $($selectedName.name), is this correct? y/n"

            }

            if($nameAnswer -eq "y"){
            
                Log-Only "[INFO] | Selected user $($selectedName.name)"
            
                Remove-MailboxDelegate -user $selectedName
                
            }
    
            elseif($nameAnswer -eq "n"){Get-RemoveSearchMenu}

        }
       
    }

    elseif($ADUsers.count -eq 0){

        Write-Host -ForegroundColor Red "USER NOT FOUND!"

        Sleep 2

        Get-RemoveSearchMenu

    }

    else{
    
        $answer = Read-Host "$($ADUsers.name) is the only user matching your search, is this the correct user? y/n"

        while($answer -notlike "y*" -and $answer -notlike "n*"){
    
        Write-Host -ForegroundColor Red "Invalid Selection, Try Again"

        sleep 2

        $answer = Read-Host "$($ADUsers.name) is the only user matching your search, is this the correct user? y/n"

        }

        if($answer -eq "y"){
        
            Log-Only "[INFO] | Selected user $($ADUsers.name)"
            
            Remove-MailboxDelegate -user $ADUsers
            
            }
    
        elseif($answer -eq "n"){Get-RemoveSearchMenu}
    
    }

}

function Find-UserDelegates ($search){

    $ADUsers = Get-ADUser -f "name -like '*$search*'" -Properties * | sort -Unique

    if($ADUsers.count -gt 1){
    
        $ADUsers.ForEach({

            "$($ADUsers.IndexOf($_) + 1) | $($_.name)"

        })

        $answer = $null
 
        $answer = Read-Host "Select a name by typing a number"
        
        while((($answer -match [regex]"^\d*$") -eq $False) -or ([int]$answer -gt $ADUsers.Count) -or ([int]$answer -eq 0) -or ($answer -eq "")){

            Write-Host -ForegroundColor Red "Invalid Selection, Try Again"
            
            $answer = Read-Host -Prompt "Select a name by typing a number"
        
        }     

        $answer = [int]$answer

        if($answer){
           
            $selectedName = $ADUsers[$answer - 1]

            $nameAnswer = Read-Host "You have selected $($selectedName.name), is this correct? y/n"

            while($nameAnswer -notlike "y*" -and $nameAnswer -notlike "n*"){
    
                Write-Host -ForegroundColor Red "Invalid Selection, Try Again"

                $nameAnswer = Read-Host "You have selected $($selectedName.name), is this correct? y/n"

            }

            if($nameAnswer -eq "y"){

                cls
            
                Log-Only "[INFO] | Selected user $($selectedName.name)"

                Write-Host -ForegroundColor Cyan "
    
_____________________________________________

Mailbox Delegates for $($selectedName.name)
_____________________________________________



"

                Log-This "[INFO] | Searching for $($selectedName.name)'s Mailbox..."    

                try{

                    $Mailbox = Get-Mailbox -Identity $selectedName.UserPrincipalName | select -Property * -ErrorAction Stop

                    Log-Success "[SUCCESS] | Found Mailbox for $($mailbox.name)!"

                }

                catch{

                    Log-Failure "[FATAL] | MAILBOX NOT FOUND"

                    sleep 2

                    Get-DelegateSearchMenu
    
                }

                $FullAccessUsers = Get-MailboxPermission -identity $($selectedName.samaccountname)
            
                $Delegates = @()

                foreach($result in $FullAccessUsers){

                    $Account = $result.user.Contains("@")

                    if($Account -eq $true){
    
                        $FullAccessAccount = $result.user.split("@")[0]

                        $Delegate = Get-ADUser -identity $FullAccessAccount

                        $Delegates += $Delegate
    
                    }                     
 
                }

                while($Delegates.count -eq 0){

                    Log-Failure "[FATAL] | $($selectedName.name) HAS NO EMAIL DELEGATES!"

                    Pause

                    Get-DelegateSearchMenu

                }

                if($Delegates){
                
                    Write-Host -ForegroundColor Cyan "
Current Mailbox Delegates for $($selectedName.name)
"

                    $Delegates.name

                    Echo ""

                    pause

                    Get-Continue

                }
                
            }
    
            elseif($nameAnswer -eq "n"){Get-DelegateSearchMenu}

        }
       
    }

    elseif($ADUsers.count -eq 0){

        Write-Host -ForegroundColor Red "USER NOT FOUND!"

        Sleep 2

        Get-DelegateSearchMenu

    }

    else{
    
        $answer = Read-Host "$($ADUsers.name) is the only user matching your search, is this the correct user? y/n"

        while($answer -notlike "y*" -and $answer -notlike "n*"){
    
        Write-Host -ForegroundColor Red "Invalid Selection, Try Again"

        sleep 2

        $answer = Read-Host "$($ADUsers.name) is the only user matching your search, is this the correct user? y/n"

        }

        if($answer -eq "y"){
        
            cls
            
            Log-Only "[INFO] | Selected user $($ADUsers.name)"

            Write-Host -ForegroundColor Cyan "
    
_____________________________________________

Mailbox Delegates for $($ADUsers.name)
_____________________________________________



"

            Log-This "[INFO] | Searching for $($ADUsers.name)'s Mailbox..."    

            try{

                $Mailbox = Get-Mailbox -Identity $ADUsers.UserPrincipalName | select -Property * -ErrorAction Stop

                Log-Success "[SUCCESS] | Found Mailbox for $($mailbox.name)!"

            }

            catch{

                Log-Failure "[FATAL] | MAILBOX NOT FOUND"

                sleep 2

                Get-DelegateSearchMenu
    
            }

            $FullAccessUsers = Get-MailboxPermission -identity $($ADUsers.samaccountname)
            
            $Delegates = @()

            foreach($result in $FullAccessUsers){

                $Account = $result.user.Contains("@")

                if($Account -eq $true){
    
                    $FullAccessAccount = $result.user.split("@")[0]

                    $Delegate = Get-ADUser -identity $FullAccessAccount

                    $Delegates += $Delegate
    
                }  
 
            }

            while($Delegates.count -eq 0){

                Log-Failure "[FATAL] | $($ADUsers.name) HAS NO EMAIL DELEGATES!"

                Pause

                Get-DelegateSearchMenu

            }

            if($Delegates){
                
                Write-Host -ForegroundColor Cyan "
Current Mailbox Delegates for $($ADUsers.name)
"

                $Delegates.name

                Echo ""

                pause

                Get-Continue

            }

        }
    
        elseif($answer -eq "n"){Get-DelegateSearchMenu}
    
    }

}

function Add-MailboxDelegate ($user) {

    cls

    Write-Host -ForegroundColor Cyan "
    
_____________________________________________

Add Mailbox Delegate to $($user.name)
_____________________________________________



"

    Log-This "[INFO] | Searching for $($user.name)'s Mailbox..."    

    try{

        $Mailbox = Get-Mailbox -Identity $user.UserPrincipalName | select -Property * -ErrorAction Stop

        Log-Success "[SUCCESS] | Found Mailbox for $($mailbox.name)!"

    }

    catch{

        Log-Failure "[FATAL] | MAILBOX NOT FOUND"

        sleep 2

        Get-AddSearchMenu
    
    }

    $searchName = Read-Host "
    
Please type the name of the person you would like to set as mailbox delegate for $($user.name). 
Press 'q' to return to the user search.
    
"

    if($searchName -eq "q"){Get-AddSearchMenu}

    elseif($searchName){
    
        $ADUsers = Get-ADUser -f "name -like '*$searchName*'" -Properties * | sort -Unique

        while($ADUsers.count -eq 0){

            Write-Host -ForegroundColor Red "USER NOT FOUND!"

            Sleep 2

            Add-MailboxDelegate -user $user

        }

        if($ADUsers.count -gt 1){

            $ADUsers.foreach({
            
                "$($ADUsers.indexof($_) + 1) | $($_.name)"
            
            })

            $answer = Read-Host "
            
Select a name by typing a number"

            while((($answer -match [regex]"^\d*$") -eq $False) -or ([int]$answer -gt $ADUsers.Count) -or ([int]$answer -eq 0) -or ($answer -eq "")){

                Write-Host -ForegroundColor Red "Invalid Selection, Try Again"
            
                $answer = Read-Host -Prompt "Select a name by typing a number"
        
            }

            $answer = [int]$answer
           
            if($answer){
            
                $selectedUser = $ADUsers[$answer - 1]

                $answer2 = Read-Host "
                
You have select user $($selectedUser.name), is this correct? y/n"

                while(($answer2 -ne "y") -and ($answer2 -ne "n")){
                
                    Write-Host -ForegroundColor Red "Invalid Selection, Try Again"

                    $answer2 = Read-Host "You have select user $($selectedUser.name), is this correct? y/n"
                
                }

                if($answer2 -eq "y"){

                    Log-Only "[INFO] | Selected user $($selectedUser.name)"

                    try{

                        Log-This "[INFO] | Verifying that $($selectedUser.name) has a Mailbox"
                    
                        $selectedUserMailbox = Get-Mailbox -Identity $selectedUser.UserPrincipalName | select -Property * -ErrorAction 'Stop'

                        Log-Success "[SUCCESS] | Verified!"
                    
                    }

                    catch{
                    
                        Log-Failure "[FATAL] | MAILBOX NOT FOUND!"

                        Sleep 2

                        Add-MailboxDelegate -user $user
                    
                    }
                
                    try{                      

                        Add-MailboxPermission -Identity $mailbox.UserPrincipalName -user $selectedUsermailbox.UserPrincipalName -AccessRights FullAccess -InheritanceType all -AutoMapping $true -WarningAction SilentlyContinue -ErrorAction 'Stop' 

                        Log-Success "[SUCCESS] | $($selectedUsermailbox.name) now has full access to $($mailbox.name)'s mailbox."

                    }

                    catch{

                        Log-Failure "[FATAL] | Failed to provide $($selectedUsermailbox.name) with full access to $($mailbox.name)'s mailbox."
                    
                    }

                    Pause

                    cls

                    Get-MainMenu
                
                }

                if($answer2 -eq "n"){Add-MailboxDelegate -user $user}
            
            }     

        }

        elseif($ADusers){
    
            $answer = Read-Host "$($ADUsers.name) is the only user matching your search, is this the correct user? y/n"

            while($answer -notlike "y*" -and $answer -notlike "n*"){
    
                Write-Host -ForegroundColor Red "Invalid Selection, Try Again"

                sleep 2

                $answer = Read-Host "$($ADUsers.name) is the only user matching your search, is this the correct user? y/n"

            }

            if($answer -eq "y"){
        
                Log-Only "[INFO] | Selected user $($ADUsers.name)"
            
                try{

                    Log-This "[INFO] | Verifying that $($ADUsers.name) has a Mailbox"
                    
                    $selectedUserMailbox = Get-Mailbox -Identity $ADUsers.UserPrincipalName | select -Property * -ErrorAction 'Stop'

                    Log-Success "[SUCCESS] | Verified!"
                    
                }

                catch{
                    
                    Log-Failure "[FATAL] | MAILBOX NOT FOUND!"

                    Sleep 2

                    Add-MailboxDelegate -user $user
                    
                }
                
                try{                      

                    Add-MailboxPermission -Identity $mailbox.UserPrincipalName -user $selectedUsermailbox.UserPrincipalName -AccessRights FullAccess -InheritanceType all -AutoMapping $true -WarningAction SilentlyContinue -ErrorAction 'Stop' 

                    Log-Success "[SUCCESS] | $($selectedUsermailbox.name) now has full access to $($mailbox.name)'s mailbox."

                    }

                catch{
                    
                    Log-Failure "[FATAL] | Failed to provide $($selectedUsermailbox.name) with full access to $($mailbox.name)'s mailbox."
                    
                }

                Pause

                cls

                Get-MainMenu
            
            }
    
            elseif($answer -eq "n"){Add-MailboxDelegate -user $user}

        }

    }

    else{

        cls
        
        Write-Host -ForegroundColor DarkYellow "You didn't type anything weirdo"

        Sleep 2

        Add-MailboxDelegate -user $user
    
    }


}

function Remove-MailboxDelegate ($user) {

    cls

    Write-Host -ForegroundColor Cyan "
    
_________________________________________________

Remove Mailbox Delegate from $($user.name)
_________________________________________________



"

    Log-This "[INFO] | Searching for $($user.name)'s Mailbox..."    

    try{

        $Mailbox = Get-Mailbox -Identity $user.UserPrincipalName | select -Property * -ErrorAction Stop

        Log-Success "[SUCCESS] | Found Mailbox for $($mailbox.name)!"

    }

    catch{

        Log-Failure "[FATAL] | MAILBOX NOT FOUND"

        sleep 2

        Get-RemoveSearchMenu
    
    }

    Log-This "[INFO] | Finding Delegates for $($mailbox.name)..."

    $FullAccessUsers = Get-MailboxPermission -identity $mailbox.name

    $ADUsers = @()

    foreach($result in $FullAccessUsers){

        $Account = $result.user.Contains("@")

        if($Account -eq $true){
    
            $FullAccessAccount = $result.user.split("@")[0]

            $ADUser = Get-ADUser -identity $FullAccessAccount

            $ADUsers += $ADUser
    
        }  
 
    }

    while($ADUsers.count -eq 0){

        Log-Failure "[FATAL] | $($mailbox.name) HAS NO EMAIL DELEGATES!"

        Sleep 2

        Get-RemoveSearchMenu

        }

    if($ADUsers.count -gt 1){

        $ADUsers.foreach({
            
            "$($ADUsers.indexof($_) + 1) | $($_.name)"
            
        })

        $answer = Read-Host "
            
Select a name by typing a number. Type 'q' to return to search menu."

        if($answer -eq "q"){Get-RemoveSearchMenu}

        while((($answer -match [regex]"^\d*$") -eq $False) -or ([int]$answer -gt $ADUsers.Count) -or ([int]$answer -eq 0) -or ($answer -eq "")){

            Write-Host -ForegroundColor Red "Invalid Selection, Try Again"
            
            $answer = Read-Host -Prompt "Select a name by typing a number"
        
        }

        $answer = [int]$answer
           
        if($answer){
            
            $selectedUser = $ADUsers[$answer - 1]

            $answer2 = Read-Host "
                
You have select user $($selectedUser.name), is this correct? y/n"

            while(($answer2 -ne "y") -and ($answer2 -ne "n")){
                
                    Write-Host -ForegroundColor Red "Invalid Selection, Try Again"

                    $answer2 = Read-Host "You have select user $($selectedUser.name), is this correct? y/n"
                
                }

            if($answer2 -eq "y"){

                try{

                    Log-This "[INFO] | Verifying that $($selectedUser.name) has a Mailbox"
                    
                    $selectedUserMailbox = Get-Mailbox -Identity $selectedUser.UserPrincipalName | select -Property * -ErrorAction 'Stop'

                    Log-Success "[SUCCESS] | Verified!"
                    
                }

                catch{
                    
                    Log-Failure "[FATAL] | MAILBOX NOT FOUND!"

                    Sleep 2

                    Remove-MailboxDelegate -user $user
                    
                }
                
                try{                      

                    Remove-MailboxPermission -Identity $mailbox.UserPrincipalName -user $selectedUsermailbox.UserPrincipalName -AccessRights FullAccess -Confirm:$false -WarningAction SilentlyContinue -ErrorAction 'Stop' 

                    Log-Success "[SUCCESS] | $($selectedUsermailbox.name) no longer has access to $($mailbox.name)'s mailbox."

                }

                catch{
                                           
                    Log-Failure "[FATAL] | Failed to remove $($selectedUsermailbox.name)'s full access rights to $($mailbox.name)'s mailbox."
                    
                }

                Pause

                cls

                Get-MainMenu
                
            }

            if($answer2 -eq "n"){Remove-MailboxDelegate -user $user}
            
        }     
    
        pause

        Get-Continue

    }

    elseif($ADusers){
    
        $answer = Read-Host "$($ADUsers.name) is the only user matching your search, is this the correct user? y/n"

        while($answer -notlike "y*" -and $answer -notlike "n*"){
    
                Write-Host -ForegroundColor Red "Invalid Selection, Try Again"

                sleep 2

                $answer = Read-Host "$($ADUsers.name) is the only user matching your search, is this the correct user? y/n"

            }

        if($answer -eq "y"){
        
                Log-Only "[INFO] | Selected user $($ADUsers.name)"
            
                try{

                    Log-This "[INFO] | Verifying that $($ADUsers.name) has a Mailbox"
                    
                    $selectedUserMailbox = Get-Mailbox -Identity $ADUsers.UserPrincipalName | select -Property * -ErrorAction 'Stop'

                    Log-Success "[SUCCESS] | Verified!"
                    
                }

                catch{
                    
                    Log-Failure "[FATAL] | MAILBOX NOT FOUND!"

                    Sleep 2

                    Get-RemoveSearchMenu
                    
                }
                
                try{                      

                    Remove-MailboxPermission -Identity $mailbox.UserPrincipalName -user $selectedUsermailbox.UserPrincipalName -AccessRights FullAccess -Confirm:$false -WarningAction SilentlyContinue -ErrorAction 'Stop' 

                    Log-Success "[SUCCESS] | $($selectedUsermailbox.name) no longer has access to $($mailbox.name)'s mailbox."

                }

                catch{
                                           
                    Log-Failure "[FATAL] | Failed to remove $($selectedUsermailbox.name)'s full access rights to $($mailbox.name)'s mailbox."
                    
                }

                Pause

                cls

                Get-MainMenu
            
            }
    
        elseif($answer -eq "n"){Get-RemoveSearchMenu}

    }

    else{

        cls
        
        Write-Host -ForegroundColor DarkYellow "You didn't type anything weirdo"

        Sleep 2

        Remove-MailboxDelegate -user $user
    
    }

}

function Log-Success ($message){

    Write-Host -ForegroundColor Green $Message 
    "$($(Get-Date).ToString('yyyy-MM-dd::hh:mm:ss')) | $Message" | Out-File $logfile -Append

}

function Log-Failure ($message){

    Write-Host -ForegroundColor Red $Message
    "$($(Get-Date).ToString('yyyy-MM-dd::hh:mm:ss')) | $Message" | Out-File $logfile -Append

}

function Log-This ($message){

    Write-Something -words $Message 
    "$($(Get-Date).ToString('yyyy-MM-dd::hh:mm:ss')) | $Message" | Out-File $logfile -Append

}

function Log-Only ($message){

    "$($(Get-Date).ToString('yyyy-MM-dd::hh:mm:ss')) | $Message" | Out-File $logfile -Append

}

function Write-Something($words){

    Write-Host -ForegroundColor Cyan $words

}

function Get-Continue{

    cls

    $continue = Read-Host -prompt "Would you like to continue? y/n"

    if($continue -eq "y"){Get-MainMenu}

    if($continue -eq "n"){Get-AreYouSure}

    else{

        Write-Host -ForegroundColor Red "Invalid Selection, Try Again"

        Sleep 3

        Get-Continue

}

}

function Get-AreYouSure{

cls

echo ""
echo ""
echo ""
echo ""
echo ""
echo ""
echo ""
echo ""
echo "                             _________________________________________________________"
echo "                             |-------------------------------------------------------|" 
echo "                             |                         EXIT??                        |"
echo "                             |-------------------------------------------------------|"
echo ""
echo ""
echo ""

$areyousure = Read-Host "                                        Are you sure you want to exit? y/n"

if($areyousure -eq "y"){

    cls

    Log-This -message "[INFO] | End Script"
    
    sleep 3

    exit
    
}

if($areyousure -eq "n"){Get-MainMenu}

else{

    write-host -foregroundcolor red "Invalid Selection, Try Again."   
    
    sleep 5
     
    Get-AreYouSure  
     
    }

}


# ---------------------------------------------------------
#endregion Setup Function
# ---------------------------------------------------------
#region Setup
# ---------------------------------------------------------

cls

# Self-elevate the script if required
if (-Not ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] 'Administrator')) {
    if ([int](Get-CimInstance -Class Win32_OperatingSystem | Select-Object -ExpandProperty BuildNumber) -ge 6000) {
        $CommandLine = "-File `"" + $MyInvocation.MyCommand.Path + "`" " + $MyInvocation.UnboundArguments
        Start-Process -FilePath PowerShell.exe -Verb Runas -ArgumentList $CommandLine
        Exit
    }
}
Log-This -message "[INFO] | Mailbox_Delegation"
Write-Host -ForegroundColor green "Running Admin PowerShell Session."
$username = "sa.userimac@wvholdings.onmicrosoft.com"
$password = "9fwsk1(Yvv#S"
$secureStringPwd = $password | ConvertTo-SecureString -AsPlainText -Force 
$Creds = New-Object System.Management.Automation.PSCredential -ArgumentList $username, $secureStringPwd
Log-This "[INFO] SETUP | Stored credentials for $($Creds.username)"
If(Test-ADPassword -Creds $Creds){

    Log-This "[INFO] SETUP | Importing Active Directory Module"
    Import-Module ActiveDirectory

    Log-This "[INFO] SETUP | Importing Microsoft Online Module"
    Import-Module MSOnline

    Log-This "[INFO] SETUP | Connecting to MSOL services"
    try{
        Connect-MsolService -Credential $Creds -WarningAction SilentlyContinue -ErrorAction 'Stop'
        Log-Success "[SUCCESS] | Connected to MSOL services"
    }
    catch{
        Log-Failure "[FATAL] | Failed to connect to MSOL services"
        exit
    }

    Log-This "[INFO] SETUP | Initializing Exchange Online Session"
    try{
        $EOLSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $Creds -Authentication Basic -AllowRedirection 
        Import-PSSession $EOLSession -AllowClobber -WarningVariable ignore
        Log-Success "[SUCCESS] | Established session with Exchange Online"
    }
    catch{
        Log-Failure "[FATAL] | Failed to establish session with Exchange Online"
        Log-This "[INFO] SETUP | Attempting to Manually Initializing Exchange Online Session"
        try{
            $EOLSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $(Get-Credential) -Authentication Basic -AllowRedirection 
            Import-PSSession $EOLSession -AllowClobber -WarningVariable ignore
            Log-Success "[SUCCESS] | Established session with Exchange Online"
        }
        catch{
            Log-Failure "[FATAL] | Failed to manually establish session with Exchange Online"
            exit
        }    
    }
}
else{
    Log-Failure "[FATAL] SETUP | Invalid credentials"
    pause
    exit
}

sleep 2


# ---------------------------------------------------------
#endregion Setup
# ---------------------------------------------------------
#region Main
# ---------------------------------------------------------

Get-MainMenu

# ---------------------------------------------------------
#endregion Main
# ---------------------------------------------------------





