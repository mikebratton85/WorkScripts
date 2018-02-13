#Offboarding_Tasks
#
#Script written by Mike Bratton
#
#
#
#
#region Variables
#---------------------------------

$Licenses = @(
    [PSCUSTOMOBJECT]@{ Title="e1"; Sku="STANDARDPACK"; Group="SG 365 License E1"; Disabled=@("TEAMS1")}

    [PSCUSTOMOBJECT]@{ Title="e3"; Sku="ENTERPRISEPACK"; Group="SG 365 License E3"; Disabled=@("TEAMS1")}

    [PSCUSTOMOBJECT]@{ Title="e5"; Sku="ENTERPRISEPREMIUM"; Group="SG 365 License E5"; Disabled=@("TEAMS1")}
    
    [PSCUSTOMOBJECT]@{ Title="dynamics"; Sku="DYN365_ENTERPRISE_PLAN1"; Group="SG 365 License Dynamics"; Disabled=@("PROJECT_CLIENT_SUBSCRIPTION", "FLOW_DYN_P2", "SHAREPOINT_PROJECT", "SHAREPOINTENTERPRISE", "NBENTERPRISE", "SHAREPOINTWAC")}
    
    [PSCUSTOMOBJECT]@{ Title="SFBPSTNDom"; sku="MCOPSTN1";Group= "SG 365 License SFB PSTN Domestic Calling"; Disabled=@()}

    [PSCUSTOMOBJECT]@{ Title="visio"; Sku="VISIOCLIENT"; Group="SG 365 License Visio"; Disabled=@()}

    [PSCUSTOMOBJECT]@{ Title="project"; Sku="PROJECTPROFESSIONAL"; Group="SG 365 License Project"; Disabled=@()}

    [PSCUSTOMOBJECT]@{ Title="SFBPSTNConf"; sku="MCOEV";Group="SG 365 License SFB PSTN Conferencing"; Disabled=@()}

    #[PSCUSTOMOBJECT]@{ Title="SFBPSTNIntl"; sku="MCOPSTN2";Group= "SG 365 License PSTN Domestic and International Calling"; Disabled=@()}
)

$Testing = $False

$dateFormat = "$($(Get-Date).ToString('yyyy-MM-dd_hh-mm-ss'))"
$CSVOffboarded = "\\techutil01\c$\Scripts\Task\Offboarding-Reports\Reports\Offboarded_Users\Offboarded_Users_Report_$($env:USERNAME)_$($dateFormat).csv"
$CSVE1 = "\\techutil01\c$\Scripts\Task\Offboarding-Reports\Reports\Offboarded_Users_with_E1_Licenses\Offboarded_Users_With_E1_Report_$($env:USERNAME)_$($dateFormat).csv"
$CSVLicenses = "\\techutil01\c$\Scripts\Task\Offboarding-Reports\Reports\All_Licensing\Licenses_Held_By_Offboarded_Users_Report_$($env:USERNAME)_$($dateFormat).csv"
$CSVRemovals = "\\techutil01\c$\Scripts\Task\Offboarding-Reports\Reports\Purged_Accounts\90_Day_Purged_Accounts_Report_$($env:USERNAME)_$($dateFormat).csv"
$CSVE1Removals = "\\techutil01\c$\Scripts\Task\Offboarding-Reports\Reports\E1_Licensing\30_Day_E1_License_Removals_Report_$($env:USERNAME)_$($dateFormat).csv"
$CSVMovedtoPurge = "\\techutil01\c$\Scripts\Task\Offboarding-Reports\Reports\Ready_to_Purge_Accounts\90_Day_Accounts_Report_$($env:USERNAME)_$($dateFormat).csv"
$ErrorActionPreference = 'SilentlyContinue'
$E1full = "wvholdings:STANDARDPACK"
$E1 = $E1full.replace("wvholdings","E1")
$TechOps = "team-techops@wvholdings.com"
$Infra = "infra@worldventures.com"
$MailToCC = "mbratton@worldventures.com"

if($Testing -eq $False){$logfile = "\\techutil01\c$\Logs\Offboarding_Tasks_Log_$($env:USERNAME)_$($dateFormat).txt"}
elseif($Testing -eq $True){$logfile = "\\techutil01\c$\Scripts\Task\Offboarding-Reports\Logs\Offboarding_Tasks_Log_$($env:USERNAME)_$($dateFormat).txt"}
else{

    Write-Host -ForegroundColor Yellow 'Is this a Test?"'

    $Testing = Read-Host 'y/n'

    while(($testing -ne "y") -and ($testing -ne "n")){
    
        Write-Host -ForegroundColor Red "Invalid Selection, Please Try Again"

        $Testing = Read-Host '$Testing = '
    
    }

    if($testing -eq "y"){$logfile = "\\techutil01\c$\Scripts\Task\Offboarding-Reports\Logs\Offboarding_Tasks_Log_$($env:USERNAME)_$($dateFormat).txt"}
    elseif($testing -eq "n"){$logfile = "\\techutil01\c$\Logs\Offboarding_Tasks_Log_$($env:USERNAME)_$($dateFormat).txt"}

}


#---------------------------------
#endregion
#---------------------------------
#region Functions
#---------------------------------
#region Logging
#---------------------------------



function Write-Something($words){

    Write-Host -foregroundcolor Cyan $words

}

function Log-This ($message){

    Write-Something -words $message
    "$($(Get-Date).ToString('yyyy-MM-dd::hh:mm:ss'))|$Message" | Out-File $logfile -Append

}

function Log-Only ($message){

    "$($(Get-Date).ToString('yyyy-MM-dd::hh:mm:ss'))|$Message" | Out-File $logfile -Append

}



#---------------------------------
#endregion
#---------------------------------
#region Menus
#---------------------------------



function Get-MainMenu{

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
echo "                             |                   OFFBOARDING TASKS                   |"
echo "                             |-------------------------------------------------------|"
echo "                             |                                                       |"
echo "                             |                                                       |"
echo "                             |                                                       |"
echo "                             |   1. Search for Users in Offboarding OU               |"
echo "                             |   2. All Users in Offboarding OU                      |"
echo "                             |   3. Find All Disabled Users with Licensing in 365    |"
echo "                             |   4. Find All Disabled Users with an E1 License       |"
echo "                             |   5. Remove E1 Licensing From Offboarded Users        |"
echo "                             |   6. Check For Offboarded Users over 90 Days Old      |"
echo "                             |   7. Purge 90 Day or Older Accounts                   |"
echo "                             |   8. Exit                                             |"
echo "                             |                                                       |"
echo "                             |                                                       |"
echo "                             |                                                       |"
echo "                             |-------------------------------------------------------|"
echo ""
echo ""
echo ""

$answer = Read-Host "                                            Please Make a Selection"

if($answer -eq 1){Get-SearchUserMenu}
if($answer -eq 2){Get-AllOffboardedUsers}
if($answer -eq 3){Get-AllDisabledandLicensed}
if($answer -eq 4){Get-E1}
if($answer -eq 5){Remove-E1Licenses}
if($answer -eq 6){Check-90DayUsers}
if($answer -eq 7){Purge-90DayUsers}
if($answer -eq 8){Get-AreYouSure}

else{Write-Host -ForegroundColor Red "Invalid Selection, Try Again." 
    sleep 2
    Get-MainMenu
}

}

function Get-SearchUserMenu {

    cls

    $searchName = Read-Host "Please type the name of the Offboarded user you would like to find. Press 'q' to return to the main menu"

    if($searchName -eq "q"){Get-MainMenu}

    elseif($searchName){Find-User -search $searchName}

    else{

        cls
        
        Write-Host -ForegroundColor DarkYellow "You didn't type anything weirdo"

        Sleep 2

        Get-SearchUserMenu
    
    }

}

function Get-UserOptionsMenu ($account){

cls

echo ""
echo ""
echo ""
echo ""
echo ""
echo ""
echo ""
Write-Host -ForegroundColor Cyan "                                        You have chosen user $($account.name)"
echo "                             _________________________________________________________"
echo "                             |-------------------------------------------------------|" 
echo "                             |                 OFFBOARDED USER TASKS                 |"
echo "                             |-------------------------------------------------------|"
echo "                             |                                                       |"
echo "                             |                                                       |"
echo "                             |                                                       |"
echo "                             |   1. Get User Info                                    |"
echo "                             |   2. Modify Employee Type                             |"
echo "                             |   3. Modify Employee Location                         |"
echo "                             |   4. Modify AD Groups                                 |"
echo "                             |   5. Enable/Disable User                              |"
echo "                             |   6. Modify Office 365 Licensing                      |"
echo "                             |   7. Add Calendar Delegate                            |"
echo "                             |   8. Add Mailbox Delegate                             |"
echo "                             |   9. Purge 90 Day or Older Account                    |"
echo "                             |   10. Select Different User                           |"
echo "                             |   11. Return to Main Menu                             |"
echo "                             |   12. Exit                                            |"
echo "                             |                                                       |"
echo "                             |                                                       |"
echo "                             |                                                       |"
echo "                             |-------------------------------------------------------|"
echo ""
echo ""
echo ""

#Clear-Variable answer, user

$answer = Read-Host "                                            Please Make a Selection"

if($answer -eq 1){Get-UserInfo -user $account}
if($answer -eq 2){Modify-Attribute15 -user $account}
if($answer -eq 3){Modify-Attribute14 -user $account}
#if($answer -eq 4){Modify-UserADGroups}
if($answer -eq 5){Modify-UserStatus -user $account}
#if($answer -eq 6){Modify-UserLicensing}
if($answer -eq 7){
    $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $plainCred -Authentication Basic -AllowRedirection
    Import-PSSession $Session
    cls
    Add-CalendarDelegate -user $account
}
if($answer -eq 8){
    $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $plainCred -Authentication Basic -AllowRedirection
    Import-PSSession $Session
    cls
    Add-MailboxDelegate -user $account
}
#if($answer -eq 9){Purge-90DayUser}
if($answer -eq 10){Get-SearchUserMenu}
if($answer -eq 11){Get-MainMenu}
if($answer -eq 12){Get-AreYouSure}


else{Write-Host -ForegroundColor Red "Invalid Selection, Try Again." 
    sleep 2
    Get-UserOptionsMenu -account $account
    }

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

    Log-This -message "INFO  End Script"
    
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



#---------------------------------
#endregion
#---------------------------------
#region All-User Tasks
#---------------------------------



function Get-AllOffboardedUsers{

    cls

    Log-This -message "
    
___________________________________________

Offboarded Users Status in Active Directory
___________________________________________



"

    $aduser = Get-ADUser -f "*" -SearchBase "OU=Disabled,OU=WorldVentures,DC=WorldVentures,DC=local" -SearchScope OneLevel -Properties description, name, manager, title, department | sort -Property description

    $table = @()

    $Rows = @()

    foreach($user in $aduser){

        if($user.enabled -eq $true){

            Write-Host -BackgroundColor Green -ForegroundColor Black "$($user.name) is Enabled"
            Log-Only -message "$($user.name) is Enabled"

        }

        else{            

            if($user.manager -ne $null){
            
                $manager = $($user.manager.split(",")[0].replace("CN=",""))

            }

            else{$manager = "Not Specified"}

            $table += [pscustomobject] @{
    
                Offboarded_User = $user.name;
                Position = $user.title;
                Department = $user.department
                Offboarded_Data = $user.description;
                Manager = $manager;
                
    
            }

            Log-This -message "$($user.name) is Disabled and was Offboarded on $($user.description.split(" ")[0])"            

            $OU = $user.distinguishedname.Split(",")[1].Replace("OU=","")

            $Row = @"

            <tr>
                <td>$($user.name)</td>
                <td>$($user.samaccountname)</td>
                <td>$($user.title)</td>
                <td>$($user.department)</td>
                <td>$($user.Description)</td>
                <td>$($manager)</td>
                <td>$($OU)</td>
            </tr>
"@

        } 

        $Rows += $Row

    }

    #$table | sort -Property Offboarded_Data | ft

    Write-Host -ForeGroundColor Cyan "
    
Total Offboarded Users in Disabled OU: $($table.Count)
    
    "

    pause

    cls

    $HTML = @"

        <html>
            <head>
                <style>
                    table{
                        font-size: 1.3em;
                        text-allign: center;
                        width: 80%;
                        margin: 0 auto;
                    }
                    th{
                        background-color: black;
                        color: yellow;
                    }
                    body{
                        font-family: Sans-Comic;
                    }
                </style>
            </head>
            <body>
                <h1>
                    Offboarded Accounts that are in the Disabled OU
                    $(Get-Date)
                </h1>
                <p>These accounts have been offboarded and are sitting in the disabled OU in Active Directory</p>
                <p>Total accounts in disabled OU: $($table.Count)</p>
                <table>
                    <thead>
                        <th>Name</th>
                        <th>Username</th>
                        <th>Position</th>
                        <th>Department</th>
                        <th>Description</th>
                        <th>Manager</th>
                        <th>OU</th>
                    </thead>
                    <tbody>
                        $Rows
                    </tbody>
                </table>
            </body>    
        </html>
"@

    $save = $null

    $save = Read-Host -prompt "Would you like to save to CSV? y/n"

    while($save -notlike "y*" -and $save -notlike "n*"){
    
        Write-Host -ForegroundColor Red "Invalid Selection, Try Again"

        sleep 2
        
        $Save = Read-Host -prompt "Would you like to save to CSV? y/n"
        
    }  

    if($save -eq "y"){
    
        $table | sort -Property Offboarded_Data | Export-Csv $CSVOffboarded

        Send-MailMessage -From "Offboarding Tasks Report no-reply@worldventures.com" -To $techops -Cc $MailToCC -Subject "Offboarded Users - $((get-date).ToShortDateString())" -Priority High -Body $HTML -BodyAsHtml -SmtpServer "relay.worldventures.com" -Attachments $CSVOffboarded
        
        Get-Continue
        
        }

    if($save -eq "n"){        
    
        Send-MailMessage -From "Offboarding Tasks Report no-reply@worldventures.com" -To $techops -Cc $MailToCC -Subject "Offboarded Users - $((get-date).ToShortDateString())" -Priority High -Body $HTML -BodyAsHtml -SmtpServer "relay.worldventures.com"
        
        Get-Continue

    }

}

function Get-DisabledUsers{

    cls

    $disabled = @()

    $aduser = $null

    $aduser = Get-ADUser -f "*" -SearchBase "OU=Disabled,OU=WorldVentures,DC=WorldVentures,DC=local" -SearchScope OneLevel -Properties name, description, samaccountname, title, department, manager

    foreach($user in $aduser){

        if($user.enabled -eq $true){continue}

        else{$disabled += $user}

    }

    return $disabled | sort -Property description

}

function Check-90DayUsers{
    
    cls

    $disabledusers = Get-DisabledUsers

    Log-This -message "
    
_________________________________________________

Check Disabled OU for Accounts Older than 90 Days
_________________________________________________



"

    $90Daydisabled = $disabledusers | where { (Get-Date($_.description.split(" ")[0])) -lt (Get-Date).AddDays(-91) -and ($_.enabled -eq $False) }

    $90Daydisabled | select -Property name, description | ft

    Log-This "There are $($90Daydisabled.count) offboarded accounts over 90 days old in the disabled OU
    "

    $Purge = Read-Host -prompt "Would you like to move these accounts to the 'Ready to Purge' OU? y/n"

    while($Purge -notlike "y*" -and $Purge -notlike "n*"){
    
        Write-Host -ForegroundColor Red "Invalid Selection, Try Again"

        sleep 2
        
        $Purge = Read-Host "Would you like to move these accounts to the 'Ready to Purge' OU? y/n"

    }

    if($Purge -eq "y"){
        
        cls

        $MovedUsers = @()

        foreach($user in $90Daydisabled){

            try{

                $user | Move-ADObject -targetpath "OU=should be gone already,OU=Disabled,OU=WorldVentures,DC=WorldVentures,DC=local" -ErrorAction Stop

                Log-This -message "$($user.name) has been successfully moved to the 'Ready to Purge' OU"

                $MovedUsers += [pscustomobject] @{
        
                    Name = $user.Name;
                    Username = $user.samaccountname;
                    Email = $user.UserPrincipalName;
                    Offboarded_Data = $user.description;       
    
                }

                $Rows += @"

                    <tr>
                        <td>$($user.Name)</td>
                        <td>$($user.samaccountname)</td>
                        <td>$($user.UserPrincipalName)</td>
                        <td>$($user.description)</td>
                    </tr>

"@        
        
            }

            catch{
                
                Write-Host -ForegroundColor Red "$($user.name) failed to move to the 'Ready to Purge' OU"

                Log-Only -message "$($user.name) failed to move to the 'Ready to Purge' OU"
            
            }

        }

        Write-Host -ForegroundColor Cyan "
    
Total users moved to the 'Ready to Purge' OU: $($MovedUsers.Count)
    
        "

        pause

        cls

        $HTML = @"

        <html>
            <head>
                <style>
                    table{
                        font-size: 1.3em;
                        text-allign: center;
                        width: 80%;
                        margin: 0 auto;
                    }
                    th{
                        background-color: black;
                        color: yellow;
                    }
                    body{
                        font-family: Sans-Comic;
                    }
                </style>
            </head>
            <body>
                <h1>
                    New 90 Day Old Offboarded Accounts
                    $(Get-Date)
                </h1>
                <p>These accounts have hit 90 days since offboarding and have been moved to the 'Ready to Purge' OU in Active Directory</p>
                <p>Total accounts moved: $($MovedUsers.Count)</p>
                <table>
                    <thead>
                        <th>Name</th>
                        <th>Username</th>
                        <th>Email</th>
                        <th>Offboarded Data</th>
                    </thead>
                    <tbody>
                        $Rows
                    </tbody>
                </table>
            </body>    
        </html>
"@

        $save = $null

        $save = Read-Host -prompt "Would you like to save to CSV? y/n"

        if($save -eq "y"){
    
            $MovedUsers | Export-Csv "$CSVMovedtoPurge"
        
            Send-MailMessage -From "Offboarding Tasks Report no-reply@worldventures.com" -To $techops -Cc $MailToCC -Subject "New 90 Day Old Offboarded Accounts - $((get-date).ToShortDateString())" -Priority High -Body $HTML -BodyAsHtml -SmtpServer "relay.worldventures.com" -Attachments $CSVMovedtoPurge

            Get-Continue
        
        }

        if($save -eq "n"){

            Send-MailMessage -From "Offboarded Users Report no-reply@worldventures.com" -To $techops -Cc $MailToCC -Subject "New 90 Day Old Offboarded Accounts - $((get-date).ToShortDateString())" -Priority High -Body $HTML -BodyAsHtml -SmtpServer "relay.worldventures.com"
    
            Get-Continue
            
        }
        
    }

    if($Purge -eq "n"){        
        
        Get-Continue

    }

}

function Get-90DayUsers{

    cls

    $90Day = @()

    $aduser = $null

    $aduser = Get-ADUser -f "*" -SearchBase "OU=should be gone already,OU=Disabled,OU=WorldVentures,DC=WorldVentures,DC=local" -SearchScope OneLevel -Properties name, description, samaccountname, title, department, manager

    foreach($user in $aduser){

        if($user.enabled -eq $true){continue}

        else{$90Day += $user}

    }

    return $90Day | sort -Property description

}

function Get-AllDisabledandLicensed{

    cls

    $disabledusers = Get-DisabledUsers

    Log-This -message "
    
______________________________________________________

All Active Licenses For Offboarded Users in Office 365
______________________________________________________



"

    $UserLicenses = @()

    $LicenseCount = @()

    foreach($name in $disabledusers){

        $msoluser = Get-MsolUser -UserPrincipalName $($name.userprincipalname) -ErrorAction SilentlyContinue

        if($msoluser -eq $null){continue}

        elseif($msoluser.IsLicensed -eq $false){continue}

        else{
           
            $License = $msoluser.licenses.accountskuid.replace("wvholdings:","")
            
            $LicensesStr = ""
            
            $License.ForEach({ $LicensesStr += $(if($License[-1] -eq $_){"$_"} else{"$_, "}) })

            Write-Something -words "$($msoluser.displayname) has these licenses"

            Log-Only -message "$($msoluser.displayname) has the following licenses: $LicensesStr"

            $License

            if($License.count -gt 0){$LicenseCount += $License}

        }
    
        $UserLicenses += [pscustomobject] @{
        
            Name = $msoluser.DisplayName;
            Username = $($msoluser.UserPrincipalName.Split("@")[0]);
            Email = $msoluser.UserPrincipalName;
            Offboarded_Data = $name.description;
            Active_Licenses = $LicensesStr;        
    
        }

        $Rows += @"

           <tr>
               <td>$($msoluser.DisplayName)</td>
               <td>$($msoluser.UserPrincipalName.Split("@")[0])</td>
               <td>$($msoluser.UserPrincipalName)</td>
               <td>$($name.description)</td>
               <td>$($LicensesStr)</td>
           </tr>

"@        

    }
    
    Write-Host -ForegroundColor Cyan "
    
Total Licenses Held by Offboarded Users: $($LicenseCount.Count)
    
    "

    pause

    cls

    $HTML = @"

        <html>
            <head>
                <style>
                    table{
                        font-size: 1.3em;
                        text-allign: center;
                        width: 80%;
                        margin: 0 auto;
                    }
                    th{
                        background-color: black;
                        color: yellow;
                    }
                    body{
                        font-family: Sans-Comic;
                    }
                </style>
            </head>
            <body>
                <h1>
                    All Licenses Held By Offboarded Users
                    $(Get-Date)
                </h1>
                <p>These are all the licenses held by Offboarded users in Office 365</p>
                <p>Total licenses that can be made available: $($LicenseCount.Count)</p>
                <table>
                    <thead>
                        <th>Name</th>
                        <th>Username</th>
                        <th>Email</th>
                        <th>Offboarded Data</th>
                        <th>Active Licenses</th>
                    </thead>
                    <tbody>
                        $Rows
                    </tbody>
                </table>
            </body>    
        </html>
"@

    $save = $null

    $save = Read-Host -prompt "Would you like to save to CSV? y/n"

    while($save -notlike "y*" -and $save -notlike "n*"){
    
        Write-Host -ForegroundColor Red "Invalid Selection, Try Again"

        sleep 2
        
        $Save = Read-Host -prompt "Would you like to save to CSV? y/n"
        
    }  

    if($save -eq "y"){
    
        $UserLicenses | Export-Csv "$CSVLicenses"
        
        Send-MailMessage -From "Offboarding Tasks Report no-reply@worldventures.com" -To $techops -Cc $MailToCC -Subject "All Licenses Held By Offboarded Users - $((get-date).ToShortDateString())" -Priority High -Body $HTML -BodyAsHtml -SmtpServer "relay.worldventures.com" -Attachments $CSVLicenses

        Get-Continue
        
        }

    if($save -eq "n"){

        Send-MailMessage -From "Offboarded Users Report no-reply@worldventures.com" -To $techops -Cc $MailToCC -Subject "All License Held By Offboarded Users - $((get-date).ToShortDateString())" -Priority High -Body $HTML -BodyAsHtml -SmtpServer "relay.worldventures.com"
    
        Get-Continue
            
    }

}

function Get-E1{

    cls

    $disabledusers = Get-DisabledUsers

    Log-This -message "
    
________________________________________________________

Offboarded Users With an Active E1 License in Office 365
________________________________________________________



"

    $users = @()

    $E1UserList = @()

    foreach($disableduser in $disabledusers){

       $exchangeuserE1 = Get-MsolUser -UserPrincipalName $($disableduser.userprincipalname) -ErrorAction SilentlyContinue

       if($exchangeuserE1 -eq $null){Log-Only -message "Disabled user $($disableduser.name) NOT FOUND in Office 365!"}

       elseif($exchangeuserE1.IsLicensed -eq $false){Log-Only "Disabled user $($exchangeuserE1.displayname) NOT LICENSED!"}

       elseif($exchangeuserE1.licenses.accountskuid -eq $E1full){

           $users += $exchangeuserE1

           Log-This -message "Disabled user $($exchangeuserE1.displayname) has an $E1 license"

       }
           
       else{Log-Only -message "Disabled user $($exchangeuserE1.displayname) is licensed, but not with an $E1 license!"}  
       
    }
           
    foreach($account in $users){           
           
       $adaccount = Get-ADUser -f{UserPrincipalName -like $account.UserPrincipalName} -Properties name, manager, samaccountname, title, department, description
           
           if($adaccount.manager -ne $null){
           
               $manager = $($adaccount.manager.split(",")[0].replace("CN=",""))

           }

           else{$manager = "Not Specified"}

       $E1UserList += [pscustomobject] @{
   
           Offboarded_User = $adaccount.name;
           Username = $adaccount.samaccountname;
           Position = $adaccount.title;
           Department = $adaccount.department;
           Offboarded_Data = $adaccount.Description;
           Manager = $manager;
           License = $E1;  
           
       }                 

       $OU = $adaccount.distinguishedname.Split(",")[1].Replace("OU=","")
          
       $Rows += @"

           <tr>
               <td>$($adaccount.name)</td>
               <td>$($adaccount.samaccountname)</td>
               <td>$($adaccount.title)</td>
               <td>$($adaccount.department)</td>
               <td>$($adaccount.Description)</td>
               <td>$($manager)</td>
               <td>$($OU)</td>
               <td>$($E1)</td>
           </tr>

"@        
                   
    }

    Write-Host -ForegroundColor Cyan "
    
Total Offboarded Users Still E1 Licensed: $($E1UserList.Count)
    
    "

    pause

    cls

    $HTML = @"

        <html>
            <head>
                <style>
                    table{
                        font-size: 1.3em;
                        text-allign: center;
                        width: 80%;
                        margin: 0 auto;
                    }
                    th{
                        background-color: black;
                        color: yellow;
                    }
                    body{
                        font-family: Sans-Comic;
                    }
                </style>
            </head>
            <body>
                <h1>
                    Offboarded Accounts with an E1 License
                    $(Get-Date)
                </h1>
                <p>These accounts have been offboarded and are still taking up an E1 license in Office 365</p>
                <p>Total E1 Licenses: $($E1UserList.Count)</p>
                <table>
                    <thead>
                        <th>Name</th>
                        <th>Username</th>
                        <th>Position</th>
                        <th>Department</th>
                        <th>Offboarded Data</th>
                        <th>Manager</th>
                        <th>OU</th>
                        <th>License</th>
                    </thead>
                    <tbody>
                        $Rows
                    </tbody>
                </table>
            </body>    
        </html>
"@

    $save = $null

    $save = Read-Host -prompt "Would you like to save to CSV? y/n"

    while($save -notlike "y*" -and $save -notlike "n*"){
    
        Write-Host -ForegroundColor Red "Invalid Selection, Try Again"

        sleep 2
        
        $Save = Read-Host -prompt "Would you like to save to CSV? y/n"
        
    }  

    if($save -eq "y"){
    
        $E1UserList | sort -property Offboarded_Data | Export-Csv $CSVE1
        
        Send-MailMessage -From "Offboarding Tasks Report no-reply@worldventures.com" -To $TechOps -Cc $MailToCC -Subject "Offboarded User Licenses: E1 - $((get-date).ToShortDateString())" -Priority High -Body $HTML -BodyAsHtml -SmtpServer "relay.worldventures.com" -Attachments $CSVE1

        Get-Continue
        
        }

    if($save -eq "n"){

        Send-MailMessage -From "Offboarded Users Report no-reply@worldventures.com" -To $TechOps -Cc $MailToCC -Subject "Offboarded User Licenses: E1 - $((get-date).ToShortDateString())" -Priority High -Body $HTML -BodyAsHtml -SmtpServer "relay.worldventures.com"
    
        Get-Continue
            
    }

}

function Remove-E1Licenses{

    cls

    $disabledusers = Get-DisabledUsers

    Log-This -message "
    
_________________________________________________________

Remove E1 Licenses from Offboarded Users over 30 days old
_________________________________________________________



"

    Write-Host -ForegroundColor Yellow "
    
Verifying users were offboarded at least 30 days ago and have an E1 license

    "

    $30DaydisabledE1 = @()

    foreach($user in $disabledusers){

        $30dayuser = $user | where { (Get-Date($_.description.split(" ")[0])) -lt (Get-Date).AddDays(-31) -and ($_.enabled -eq $False) }

        $msoluserE1 = Get-MsolUser -UserPrincipalName $30dayuser.UserPrincipalName -ErrorAction SilentlyContinue | Where-Object {$_.licenses.accountskuid -eq $E1full}

        if($msoluserE1){ 
        
            $30DaydisabledE1 += $msoluserE1           

            Write-Host -ForegroundColor Cyan "Verified $($msoluserE1.displayname) is ready to have E1 license removed"
        }

        else{Continue}

    }
    
    Write-Host -ForegroundColor Cyan "
    
Total E1 licenses that can be made available: $($30DaydisabledE1.Count)

    "

    $Warning = Read-Host -prompt "Are you sure you want to remove E1 licensing? This cannot be undone. y/n"

    while($Warning -notlike "y*" -and $Warning -notlike "n*"){
    
        Write-Host -ForegroundColor Red "Invalid Selection, Try Again"

        sleep 2
        
        $Warning = Read-Host -prompt "Are you sure you want to remove E1 licensing? This cannot be undone. y/n"
        
    }  
    
    if($Warning -eq "n"){Get-MainMenu}
    
    if($Warning -eq "y"){}  

    $RemovedE1 = @()

    foreach($name in $30DaydisabledE1){

        $aduser = Get-ADUser -f {UserPrincipalName -eq $name.UserPrincipalName}

        Write-Host -ForegroundColor Yellow "Removing E1 licensing from user $($name.displayname)"

        Try{
                
            Set-MsolUserLicense $name -RemoveLicenses $E1full -ErrorAction Stop

            $RemovedE1 += $name

            Log-This -message "Successfully removed E1 licensing from user $($name.displayname)"

        }

        Catch{

            Write-Host -ForegroundColor Red "Failed to remove E1 licensing from $($name.displayname)"
                
            Log-Only -message "Failed to remove E1 licensing from $($name.displayname)"               
            
        }

        Write-Host -ForegroundColor Yellow "Removing user $($aduser.name) from the E1 licensing group in Active Directory"
    
        Try{
        
            Remove-ADGroupMember -Identity "SG 365 License E1" -members $aduser.samaccountname

            Log-This -message "Successfully removed $($aduser.name) from the E1 licensing group in Active Directory
            "

        }

        Catch{
                
            Write-Host -ForegroundColor Red "Failed to remove $($aduser.name) from the E1 licensing group in Active Directory
            "
                
            Log-Only -message "Failed to remove $($aduser.name) from the E1 licensing group in Active Directory
            "                               
                
        }
    
    } 
    
    $RemovedUser = @()

    foreach($user in $RemovedE1){
        
        $aduser = Get-ADUser -f {UserPrincipalName -eq $user.UserPrincipalName} -Properties description
           
        $RemovedUser += [pscustomobject] @{
        
            Name = $user.displayname;
            Username = $($user.UserPrincipalName.split("@")[0]);
            Email = $user.UserPrincipalName;
            Offboarded_Data = $aduser.description;
            Removed_Licenses = $E1;
            Total_Licenses_Removed = $user.Count;     
        
        }

        $Rows += @"

           <tr>
               <td>$($user.DisplayName)</td>
               <td>$($user.UserPrincipalName.Split("@")[0])</td>
               <td>$($user.UserPrincipalName)</td>
               <td>$($aduser.description)</td>
               <td>$($E1)</td>
               <td>$($user.count)</td>
           </tr>

"@        


    }   

    Write-Host -ForegroundColor Cyan "
    
Total E1 licenses made available: $($RemovedE1.Count)

    "

    Pause

    cls    

    $HTML = @"

        <html>
            <head>
                <style>
                    table{
                        font-size: 1.3em;
                        text-allign: center;
                        width: 80%;
                        margin: 0 auto;
                    }
                    th{
                        background-color: black;
                        color: yellow;
                    }
                    body{
                        font-family: Sans-Comic;
                    }
                </style>
            </head>
            <body>
                <h1>
                    90 Day E1 License Removal Report
                    $(Get-Date)
                </h1>
                <p>These users have had their E1 licenses removed in Office 365 and they have been removed from the 'SG 365 License E1' Security Group</p>
                <p>Total licenses that have been made available: $($RemovedE1.Count)</p>
                <table>
                    <thead>
                        <th>Name</th>
                        <th>Username</th>
                        <th>Email</th>
                        <th>Offboarded Data</th>
                        <th>Removed Licenses</th>
                        <th>Total Licenses Removed</th>
                    </thead>
                    <tbody>
                        $Rows
                    </tbody>
                </table>
            </body>    
        </html>
"@

    $save = $null

    $save = Read-Host -prompt "Would you like to save to CSV? y/n"

    while($save -notlike "y*" -and $save -notlike "n*"){
    
        Write-Host -ForegroundColor Red "Invalid Selection, Try Again"

        sleep 2
        
        $Save = Read-Host -prompt "Would you like to save to CSV? y/n"
        
    }  

    if($save -eq "y"){
    
        $RemovedUser | Export-Csv $CSVE1Removals
        
        Send-MailMessage -From "Offboarding Tasks Report no-reply@worldventures.com" -To $techops -Cc $MailToCC -Subject "30 Day E1 License Removals - $((get-date).ToShortDateString())" -Priority High -Body $HTML -BodyAsHtml -SmtpServer "relay.worldventures.com" -Attachments $CSVE1Removals

        Get-Continue
        
        }

    if($save -eq "n"){

        Send-MailMessage -From "Offboarded Users Report no-reply@worldventures.com" -To $techops -Cc $MailToCC -Subject "30 Day E1 License Removals - $((get-date).ToShortDateString())" -Priority High -Body $HTML -BodyAsHtml -SmtpServer "relay.worldventures.com"
    
        Get-Continue
            
    }

}

function Purge-90DayUsers{

    cls

    $90DayUsers = Get-90DayUsers

    Log-This -message "
    
________________________________________________________________________________________

Purge All 90 Day or Older Users From Active Directory and Remove Licensing in Office 365
________________________________________________________________________________________



"

    foreach($name in $90DayUsers){
    
        Write-Host -ForegroundColor Cyan "$($name.name) is ready to purge
        "
    
    }

    $Warning = Read-Host -prompt "Are you sure you want to remove licensing and delete accounts? This cannot be undone. y/n"

    while($Warning -notlike "y*" -and $Warning -notlike "n*"){
    
        Write-Host -ForegroundColor Red "Invalid Selection, Try Again"

        sleep 2
        
        $Warning = Read-Host -prompt "Are you sure you want to remove licensing and delete accounts? This cannot be undone. y/n"
        
    }  

    if($Warning -eq "n"){Get-MainMenu}
    
    elseif($Warning -eq "y"){

        $RemovedLicense = @()

        $Removed90DayUser = @()

        foreach($name in $90DayUsers){

            $90Daymsoluser = Get-MsolUser -UserPrincipalName $($name.userprincipalname)

            $90DayLicense = $90Daymsoluser.licenses.accountskuid.replace("wvholdings:","")
            
            $90DayLicensesStr = ""
        
            $90DayLicense.ForEach({ $90DayLicensesStr += $(if($90DayLicense[-1] -eq $_){"$_"} else{"$_, "}) })

            try{

                $name | Remove-ADUser -Confirm:$false -erroraction Stop

                Log-This -message "$($name.name) was successfully removed from Active Directory"

            }
                
            catch{
                
                Write-Host -ForegroundColor Red "Failed to remove $($name.name) from Active Directory"
                
                Log-Only -message "Failed to remove $($name.name) from Active Directory"

            } 
                
            if($90Daymsoluser -eq $null){Continue}

            else{
            
                if(($90Daymsoluser.IsLicensed -eq $true) -and ($90DayLicense.count -gt 0)){
                
                    Write-Host -foregroundcolor Yellow "Removing Licenses from user $($90Daymsoluser.displayname)"

                    Log-Only -message "Removing Licenses from user $($90Daymsoluser.displayname)"
                    
                }

                if($90Daymsoluser.IsLicensed -eq $false){Continue}

                else{
                
                    try{
                
                        $90DayMSOLUser | Set-MSOLUserLicense -removelicense $90DayMSOLUser.licenses.accountskuid -ErrorAction Stop
                    
                        Write-Host -ForegroundColor DarkYellow "Successfully removed licenses $90DayLicense"

                        Log-Only -message "Successfully Removed Licenses $90DayLicensesStr"

                        if($90DayLicense.count -gt 0){$RemovedLicense += $90DayLicense}

                        Log-This -message "All Licensing has been removed from user $($90Daymsoluser.displayname)"

                    }

                    catch{
                
                        Write-Host -ForegroundColor Red "Failed to remove licenses from $($90Daymsoluser.displayname)"

                        Log-Only -message "Failed to remove licenses from $($90Daymsoluser.displayname)"
                    
                    }

                }
                
            }

            $Removed90DayUser += [pscustomobject] @{
        
                Name = $90Daymsoluser.DisplayName;
                Username = $($90Daymsoluser.UserPrincipalName.split("@")[0]);
                Email = $90Daymsoluser.UserPrincipalName;
                Offboarded_Data = $name.description;
                Removed_Licenses = $90DayLicensesStr;
                Total_Licenses_Removed = $90DayLicense.Count;       
    
            }

            $Rows += @"

                <tr>
                    <td>$($90Daymsoluser.DisplayName)</td>
                    <td>$($($90Daymsoluser.UserPrincipalName.split("@")[0]))</td>
                    <td>$($90Daymsoluser.UserPrincipalName)</td>
                    <td>$($name.description)</td>
                    <td>$($90DayLicensesStr)</td>
                    <td>$($90DayLicense.Count)</td>
                </tr>

"@        



        } 
           
    }

    Write-Host -ForegroundColor Cyan "
    
Total licenses made available: $($RemovedLicense.Count)

    "

    Pause

    cls

    $HTML = @"

        <html>
            <head>
                <style>
                    table{
                        font-size: 1.3em;
                        text-allign: center;
                        width: 80%;
                        margin: 0 auto;
                    }
                    th{
                        background-color: black;
                        color: yellow;
                    }
                    body{
                        font-family: Sans-Comic;
                    }
                </style>
            </head>
            <body>
                <h1>
                    90 Day User Purge Report
                    $(Get-Date)
                </h1>
                <p>These users have been removed from Active Directory and their licenses have been removed in Office 365</p>
                <p>Total licenses that have been made available: $($RemovedLicense.Count)</p>
                <table>
                    <thead>
                        <th>Name</th>
                        <th>Username</th>
                        <th>Email</th>
                        <th>Offboarded Data</th>
                        <th>Removed Licenses</th>
                        <th>Total Licenses Removed</th>
                    </thead>
                    <tbody>
                        $Rows
                    </tbody>
                </table>
            </body>    
        </html>
"@

    $save = $null

    $save = Read-Host -prompt "Would you like to save to CSV? y/n"

    while($save -notlike "y*" -and $save -notlike "n*"){
    
        Write-Host -ForegroundColor Red "Invalid Selection, Try Again"

        sleep 2
        
        $Save = Read-Host -prompt "Would you like to save to CSV? y/n"
        
    }  

    if($save -eq "y"){
    
        $Removed90DayUser | Export-Csv $CSVRemovals
        
        Send-MailMessage -From "Offboarding Tasks Report no-reply@worldventures.com" -To $techops -Cc $MailToCC -Subject "90 Day User Purge Report - $((get-date).ToShortDateString())" -Priority High -Body $HTML -BodyAsHtml -SmtpServer "relay.worldventures.com" -Attachments $CSVRemovals

        Get-Continue
        
        }

    if($save -eq "n"){

        Send-MailMessage -From "Offboarded Users Report no-reply@worldventures.com" -To $techops -Cc $MailToCC -Subject "90 Day User Purge Report - $((get-date).ToShortDateString())" -Priority High -Body $HTML -BodyAsHtml -SmtpServer "relay.worldventures.com"
    
        Get-Continue
            
    }

}



#---------------------------------
#endregion
#---------------------------------
#region Single-User Tasks
#---------------------------------



function Find-User ($search){

    $ADUsers = Get-ADUser -f "name -like '*$search*'" -searchbase (Get-ADOrganizationalUnit -f {name -like "*disabled*"}).distinguishedname -SearchScope OneLevel -Properties * | sort -Unique

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

            if($nameAnswer -eq "y"){Get-UserOptionsMenu -account $selectedName}
    
            elseif($nameAnswer -eq "n"){Get-SearchUserMenu}

        }
       
    }

    elseif($ADUsers.count -eq 0){

        Write-Host -ForegroundColor Red "USER NOT FOUND!"

        Sleep 2

        Get-SearchUserMenu

    }

    else{
    
        $answer = Read-Host "$($ADUsers.name) is the only user matching your search, is this the correct user? y/n"

        while($answer -notlike "y*" -and $answer -notlike "n*"){
    
        Write-Host -ForegroundColor Red "Invalid Selection, Try Again"

        sleep 2

        $answer = Read-Host "$($ADUsers.name) is the only user matching your search, is this the correct user? y/n"

        }

        if($answer -eq "y"){Get-UserOptionsMenu -account $ADUsers}
    
        elseif($answer -eq "n"){Get-SearchUserMenu}
    
    }

}

function Get-UserInfo ($user){
    
    cls

    Log-This -message "
    
______________________________________________________

Offboarded User $($user.name) Account Information
______________________________________________________



"

    $aduser = $user | select -Property name, samaccountname, title, department, company, mail, telephonenumber, extensionattribute14, extensionattribute15, extensionattribute12, enabled, description, manager

    if($aduser.manager -ne $null){
            
        $manager = $($aduser.manager.split(",")[0].replace("CN=",""))

    }

    else{$manager = "Not Specified"}

    if($aduser.telephonenumber -ne $null){
    
        $phone = $($aduser.telephonenumber.Replace("+ 1 ","")) 
    
    }

    else{$phone = "Not Specified"}

    if($aduser.description -like "*Delegate:*"){

        $Table = [pscustomobject] @{

            Name = $aduser.Name;
            Username = $aduser.SamAccountName
            Title = $aduser.title;
            Department = $aduser.department;
            Manager = $manager;
            Company = $aduser.company;
            Email = $aduser.mail;
            Phone = $phone;
            Location = $aduser.extensionattribute14;
            Employement_Type = $aduser.extensionattribute15;
            Employement_Status = $aduser.extensionattribute12;
            Account_Status = $aduser.Enabled.tostring().replace("True","Enabled").replace("False","Disabled");
            Offboarding_Date = $aduser.description.split(" ")[0];
            Offboard_Submitter = $aduser.description.Split(":")[2].trim();
            Delegate = $aduser.description.split(":")[1].split("|")[0].replace("Submitter", "").trim();
 
        }

        $Table

        pause

    }

    elseif($aduser.description -like "*Okta*"){

        $Table = [pscustomobject] @{

            Name = $aduser.Name;
            Username = $aduser.SamAccountName
            Title = $aduser.title;
            Department = $aduser.department;
            Manager = $manager;
            Company = $aduser.company;
            Email = $aduser.mail;
            Phone = $phone;
            Location = $aduser.extensionattribute14;
            Employement_Type = $aduser.extensionattribute15;
            Employement_Status = $aduser.extensionattribute12;
            Account_Status = $aduser.Enabled.tostring().replace("True","Enabled").replace("False","Disabled");
            Offboarding_Date = $aduser.description.split(" ")[0];
            Offboard_Reason = $aduser.description.Replace("2017-09-22 | ","");
 
        }

        $Table

        pause

    }


    else{
    
        $Table = [pscustomobject] @{

            Name = $aduser.Name;
            Username = $aduser.SamAccountName
            Title = $aduser.title;
            Department = $aduser.department;
            Manager = $manager;
            Company = $aduser.company;
            Email = $aduser.mail;
            Phone = $phone;
            Location = $aduser.extensionattribute14;
            Employement_Type = $aduser.extensionattribute15;
            Employement_Status = $aduser.extensionattribute12;
            Account_Status = $aduser.Enabled.tostring().replace("True","Enabled").replace("False","Disabled");
            Description = $aduser.description;
    
        }

        $Table

        pause

    }

    Get-UserOptionsMenu -account $user
}

function Modify-Attribute15 ($user) {

cls

echo ""
echo ""
echo ""
echo ""
echo ""
echo ""
echo ""
Write-Host -ForegroundColor Cyan "                                $($user.name) currently has the employee type of '$($user.extensionattribute15)'"
Write-Host -ForegroundColor Cyan "                                Please select from the list below to make a change"
echo "                             _________________________________________________________"
echo "                             |-------------------------------------------------------|" 
echo "                             |                 CHANGE EMPLOYEE TYPE                  |"
echo "                             |-------------------------------------------------------|"
echo "                             |                                                       |"
echo "                             |                                                       |"
echo "                             |                                                       |"
echo "                             |   1. Staff                                            |"
echo "                             |   2. Partner                                          |"
echo "                             |   3. Agency                                           |"
echo "                             |   4. 1099                                             |"
echo "                             |   5. Service                                          |"
echo "                             |   6. Admin                                            |"
echo "                             |   7. Resource                                         |"
echo "                             |   8. Shared                                           |"
echo "                             |   9. Test                                             |"
echo "                             |   10. Others                                          |"
echo "                             |   11. Return to User Menu                             |"
echo "                             |   12. Return to Main Menu                             |"
echo "                             |                                                       |"
echo "                             |                                                       |"
echo "                             |                                                       |"
echo "                             |-------------------------------------------------------|"
echo ""
echo ""
echo ""

$answer = Read-Host "                                            Please Make a Selection"

if($answer -eq 1){

    while($user.extensionattribute15 -eq "Staff"){

        Write-Host -ForegroundColor Red "$($user.name) already has an employee type of Staff!"

        Sleep 2

        Modify-Attribute15 -user $user

    }
    
    $answer2 = Read-Host "Are you sure you want to change $($user.name) to the employee type of Staff? y/n"

    while(($answer2 -ne "y") -and ($answer2 -ne "n")){
    
        Write-Host -ForegroundColor Red "Invalid Selection, Try Again!"

        $answer2 = Read-Host "Are you sure you want to change $($user.name) to the employee type of Staff? y/n"
    
    }

    if($answer2 -eq "y"){
    
        try{
        
            Set-ADUser $user -replace @{extensionattribute15="Staff"} -ErrorAction Stop

            Log-This "Successfully changed $($user.name) to the employee type of Staff"

            Pause
        
        }
        
        catch{

            Write-Host -ForegroundColor Red "Failed to change $($user.name) to the employee type of Staff"
        
            Log-Only "Failed to change $($user.name) to the employee type of Staff"

            Pause
        
        }
        
    Get-UserOptionsMenu -account $user      

    }

    if($answer2 -eq "n"){Get-UserOptionsMenu -account $user}

}
if($answer -eq 2){

    while($user.extensionattribute15 -eq "Partner"){

        Write-Host -ForegroundColor Red "$($user.name) already has an employee type of Partner!"

        Sleep 2

        Modify-Attribute15 -user $user

    }
    
    $answer2 = Read-Host "Are you sure you want to change $($user.name) to the employee type of Partner? y/n"

    while(($answer2 -ne "y") -and ($answer2 -ne "n")){
    
        Write-Host -ForegroundColor Red "Invalid Selection, Try Again!"

        $answer2 = Read-Host "Are you sure you want to change $($user.name) to the employee type of Partner? y/n"
    
    }

    if($answer2 -eq "y"){
    
        try{
        
            Set-ADUser $user -replace @{extensionattribute15="Partner"} -ErrorAction Stop

            Log-This "Successfully changed $($user.name) to the employee type of Partner"

            Pause
        
        }
        
        catch{

            Write-Host -ForegroundColor Red "Failed to change $($user.name) to the employee type of Partner"
        
            Log-Only "Failed to change $($user.name) to the employee type of Partner"

            Pause
        
        }
        
    Get-UserOptionsMenu -account $user      

    }

    if($answer2 -eq "n"){Get-UserOptionsMenu -account $user}

}
if($answer -eq 3){

    while($user.extensionattribute15 -eq "Agency"){

        Write-Host -ForegroundColor Red "$($user.name) already has an employee type of Agency!"

        Sleep 2

        Modify-Attribute15 -user $user

    }
    
    $answer2 = Read-Host "Are you sure you want to change $($user.name) to the employee type of Agency? y/n"

    while(($answer2 -ne "y") -and ($answer2 -ne "n")){
    
        Write-Host -ForegroundColor Red "Invalid Selection, Try Again!"

        $answer2 = Read-Host "Are you sure you want to change $($user.name) to the employee type of Agency? y/n"
    }

    if($answer2 -eq "y"){
    
        try{
        
            Set-ADUser $user -replace @{extensionattribute15="Agency"} -ErrorAction Stop

            Log-This "Successfully changed $($user.name) to the employee type of Agency"

            Pause
        
        }
        
        catch{

            Write-Host -ForegroundColor Red "Failed to change $($user.name) to the employee type of Agency"
        
            Log-Only "Failed to change $($user.name) to the employee type of Agency"

            Pause
        
        }
        
    Get-UserOptionsMenu -account $user      

    }

    if($answer2 -eq "n"){Get-UserOptionsMenu -account $user}

}
if($answer -eq 4){

    while($user.extensionattribute15 -eq "1099"){

        Write-Host -ForegroundColor Red "$($user.name) already has an employee type of 1099!"

        Sleep 2

        Modify-Attribute15 -user $user

    }
    
    $answer2 = Read-Host "Are you sure you want to change $($user.name) to the employee type of 1099? y/n"

    while(($answer2 -ne "y") -and ($answer2 -ne "n")){
    
        Write-Host -ForegroundColor Red "Invalid Selection, Try Again!"

        $answer2 = Read-Host "Are you sure you want to change $($user.name) to the employee type of 1099? y/n"
    }

    if($answer2 -eq "y"){
    
        try{
        
            Set-ADUser $user -replace @{extensionattribute15="1099"} -ErrorAction Stop

            Log-This "Successfully changed $($user.name) to the employee type of 1099"

            Pause
        
        }
        
        catch{

            Write-Host -ForegroundColor Red "Failed to change $($user.name) to the employee type of 1099"
        
            Log-Only "Failed to change $($user.name) to the employee type of 1099"

            Pause
        
        }
        
    Get-UserOptionsMenu -account $user      

    }

    if($answer2 -eq "n"){Get-UserOptionsMenu -account $user}

}
if($answer -eq 5){

    while($user.extensionattribute15 -eq "Service"){

        Write-Host -ForegroundColor Red "$($user.name) already has an employee type of Service!"

        Sleep 2

        Modify-Attribute15 -user $user

    }
    
    $answer2 = Read-Host "Are you sure you want to change $($user.name) to the employee type of Service? y/n"

    while(($answer2 -ne "y") -and ($answer2 -ne "n")){
    
        Write-Host -ForegroundColor Red "Invalid Selection, Try Again!"

        $answer2 = Read-Host "Are you sure you want to change $($user.name) to the employee type of Service? y/n"
    }

    if($answer2 -eq "y"){
    
        try{
        
            Set-ADUser $user -replace @{extensionattribute15="Service"} -ErrorAction Stop

            Log-This "Successfully changed $($user.name) to the employee type of Service"

            Pause
        
        }
        
        catch{

            Write-Host -ForegroundColor Red "Failed to change $($user.name) to the employee type of Service"
        
            Log-Only "Failed to change $($user.name) to the employee type of Service"

            Pause
        
        }
        
    Get-UserOptionsMenu -account $user      

    }

    if($answer2 -eq "n"){Get-UserOptionsMenu -account $user}

}
if($answer -eq 6){

    while($user.extensionattribute15 -eq "Admin"){

        Write-Host -ForegroundColor Red "$($user.name) already has an employee type of Admin!"

        Sleep 2

        Modify-Attribute15 -user $user

    }
    
    $answer2 = Read-Host "Are you sure you want to change $($user.name) to the employee type of Admin? y/n"

    while(($answer2 -ne "y") -and ($answer2 -ne "n")){
    
        Write-Host -ForegroundColor Red "Invalid Selection, Try Again!"

        $answer2 = Read-Host "Are you sure you want to change $($user.name) to the employee type of Admin? y/n"
    }

    if($answer2 -eq "y"){
    
        try{
        
            Set-ADUser $user -replace @{extensionattribute15="Admin"} -ErrorAction Stop

            Log-This "Successfully changed $($user.name) to the employee type of Admin"

            Pause
        
        }
        
        catch{

            Write-Host -ForegroundColor Red "Failed to change $($user.name) to the employee type of Admin"
        
            Log-Only "Failed to change $($user.name) to the employee type of Admin"

            Pause
        
        }
        
    Get-UserOptionsMenu -account $user      

    }

    if($answer2 -eq "n"){Get-UserOptionsMenu -account $user}

}
if($answer -eq 7){

    while($user.extensionattribute15 -eq "Resource"){

        Write-Host -ForegroundColor Red "$($user.name) already has an employee type of Resource!"

        Sleep 2

        Modify-Attribute15 -user $user

    }
    
    $answer2 = Read-Host "Are you sure you want to change $($user.name) to the employee type of Resource? y/n"

    while(($answer2 -ne "y") -and ($answer2 -ne "n")){
    
        Write-Host -ForegroundColor Red "Invalid Selection, Try Again!"

        $answer2 = Read-Host "Are you sure you want to change $($user.name) to the employee type of Resource? y/n"
    }

    if($answer2 -eq "y"){
    
        try{
        
            Set-ADUser $user -replace @{extensionattribute15="Resource"} -ErrorAction Stop

            Log-This "Successfully changed $($user.name) to the employee type of Resource"

            Pause
        
        }
        
        catch{

            Write-Host -ForegroundColor Red "Failed to change $($user.name) to the employee type of Resource"
        
            Log-Only "Failed to change $($user.name) to the employee type of Resource"

            Pause
        
        }
        
    Get-UserOptionsMenu -account $user      

    }

    if($answer2 -eq "n"){Get-UserOptionsMenu -account $user}

}
if($answer -eq 8){

    while($user.extensionattribute15 -eq "Shared"){

        Write-Host -ForegroundColor Red "$($user.name) already has an employee type of Shared!"

        Sleep 2

        Modify-Attribute15 -user $user

    }
    
    $answer2 = Read-Host "Are you sure you want to change $($user.name) to the employee type of Shared? y/n"

    while(($answer2 -ne "y") -and ($answer2 -ne "n")){
    
        Write-Host -ForegroundColor Red "Invalid Selection, Try Again!"

        $answer2 = Read-Host "Are you sure you want to change $($user.name) to the employee type of Shared? y/n"
    }

    if($answer2 -eq "y"){
    
        try{
        
            Set-ADUser $user -replace @{extensionattribute15="Shared"} -ErrorAction Stop

            Log-This "Successfully changed $($user.name) to the employee type of Shared"

            Pause
        
        }
        
        catch{

            Write-Host -ForegroundColor Red "Failed to change $($user.name) to the employee type of Shared"
        
            Log-Only "Failed to change $($user.name) to the employee type of Shared"

            Pause
        
        }
        
    Get-UserOptionsMenu -account $user      

    }

    if($answer2 -eq "n"){Get-UserOptionsMenu -account $user}

}
if($answer -eq 9){

    while($user.extensionattribute15 -eq "Test"){

        Write-Host -ForegroundColor Red "$($user.name) already has an employee type of Test!"

        Sleep 2

        Modify-Attribute15 -user $user

    }
    
    $answer2 = Read-Host "Are you sure you want to change $($user.name) to the employee type of Test? y/n"

    while(($answer2 -ne "y") -and ($answer2 -ne "n")){
    
        Write-Host -ForegroundColor Red "Invalid Selection, Try Again!"

        $answer2 = Read-Host "Are you sure you want to change $($user.name) to the employee type of Test? y/n"
    }

    if($answer2 -eq "y"){
    
        try{
        
            Set-ADUser $user -replace @{extensionattribute15="Test"} -ErrorAction Stop

            Log-This "Successfully changed $($user.name) to the employee type of Test"

            Pause
        
        }
        
        catch{

            Write-Host -ForegroundColor Red "Failed to change $($user.name) to the employee type of Test"
        
            Log-Only "Failed to change $($user.name) to the employee type of Test"

            Pause
        
        }
        
    Get-UserOptionsMenu -account $user      

    }

    if($answer2 -eq "n"){Get-UserOptionsMenu -account $user}

}
if($answer -eq 10){

    while($user.extensionattribute15 -eq "Others"){

        Write-Host -ForegroundColor Red "$($user.name) already has an employee type of Others!"

        Sleep 2

        Modify-Attribute15 -user $user

    }
    
    $answer2 = Read-Host "Are you sure you want to change $($user.name) to the employee type of Others? y/n"

    while(($answer2 -ne "y") -and ($answer2 -ne "n")){
    
        Write-Host -ForegroundColor Red "Invalid Selection, Try Again!"

        $answer2 = Read-Host "Are you sure you want to change $($user.name) to the employee type of Others? y/n"
    }

    if($answer2 -eq "y"){
    
        try{
        
            Set-ADUser $user -replace @{extensionattribute15="Others"} -ErrorAction Stop

            Log-This "Successfully changed $($user.name) to the employee type of Others"

            Pause
        
        }
        
        catch{

            Write-Host -ForegroundColor Red "Failed to change $($user.name) to the employee type of Others"
        
            Log-Only "Failed to change $($user.name) to the employee type of Others"

            Pause
        
        }
        
    Get-UserOptionsMenu -account $user      

    }

    if($answer2 -eq "n"){Get-UserOptionsMenu -account $user}

}
if($answer -eq 11){Get-UserOptionsMenu -account $user}
if($answer -eq 12){Get-MainMenu}


else{Write-Host -ForegroundColor Red "Invalid Selection, Try Again." 
    sleep 2
    Get-UserOptionsMenu -account $account
    }

}

function Modify-Attribute14 ($user) {

cls

echo ""
echo ""
echo ""
echo ""
echo ""
echo ""
echo ""
Write-Host -ForegroundColor Cyan "                            $($user.name) currently has the employee location of '$($user.extensionattribute14)'"
Write-Host -ForegroundColor Cyan "                               Please select from the list below to make a change"
echo "                             _________________________________________________________"
echo "                             |-------------------------------------------------------|" 
echo "                             |               CHANGE EMPLOYEE LOCATION                |"
echo "                             |-------------------------------------------------------|"
echo "                             |                                                       |"
echo "                             |                                                       |"
echo "                             |                                                       |"
echo "                             |   1. Plano1                                           |"
echo "                             |   2. Plano2                                           |"
echo "                             |   3. Greenville1                                      |"
echo "                             |   4. HongKong1                                        |"
echo "                             |   5. Singapore1                                       |"
echo "                             |   6. Malaysia1                                        |"
echo "                             |   7. Manila1                                          |"
echo "                             |   8. Taipei1                                          |"
echo "                             |   9. Australia1                                       |"
echo "                             |   10. Remote                                          |"
echo "                             |   11. Home                                            |"
echo "                             |   12. Return to User Menu                             |"
echo "                             |   13. Return to Main Menu                             |"
echo "                             |                                                       |"
echo "                             |                                                       |"
echo "                             |                                                       |"
echo "                             |-------------------------------------------------------|"
echo ""
echo ""
echo ""

$answer = Read-Host "                                            Please Make a Selection"

if($answer -eq 1){

    while($user.extensionattribute14 -eq "Plano1"){

        Write-Host -ForegroundColor Red "$($user.name) already has an employee location of Plano1!"

        Sleep 2

        Modify-Attribute14 -user $user

    }
    
    $answer2 = Read-Host "Are you sure you want to change $($user.name) to the employee location of Plano1? y/n"

    while(($answer2 -ne "y") -and ($answer2 -ne "n")){
    
        Write-Host -ForegroundColor Red "Invalid Selection, Try Again!"

        $answer2 = Read-Host "Are you sure you want to change $($user.name) to the employee location of Plano1? y/n"
    
    }

    if($answer2 -eq "y"){
    
        try{
        
            Set-ADUser $user -replace @{extensionattribute14="Plano1"} -ErrorAction Stop

            Log-This "Successfully changed $($user.name) to the employee location of Plano1"

            Pause
        
        }
        
        catch{

            Write-Host -ForegroundColor Red "Failed to change $($user.name) to the employee location of Plano1"
        
            Log-Only "Failed to change $($user.name) to the employee location of Plano1"

            Pause
        
        }
        
    Get-UserOptionsMenu -account $user      

    }

    if($answer2 -eq "n"){Get-UserOptionsMenu -account $user}

}
if($answer -eq 2){

    while($user.extensionattribute14 -eq "Plano2"){

        Write-Host -ForegroundColor Red "$($user.name) already has an employee location of Plano2!"

        Sleep 2

        Modify-Attribute14 -user $user

    }
    
    $answer2 = Read-Host "Are you sure you want to change $($user.name) to the employee location of Plano2? y/n"

    while(($answer2 -ne "y") -and ($answer2 -ne "n")){
    
        Write-Host -ForegroundColor Red "Invalid Selection, Try Again!"

        $answer2 = Read-Host "Are you sure you want to change $($user.name) to the employee location of Plano2? y/n"
    
    }

    if($answer2 -eq "y"){
    
        try{
        
            Set-ADUser $user -replace @{extensionattribute14="Plano2"} -ErrorAction Stop

            Log-This "Successfully changed $($user.name) to the employee location of Plano2"

            Pause
        
        }
        
        catch{

            Write-Host -ForegroundColor Red "Failed to change $($user.name) to the employee location of Plano2"
        
            Log-Only "Failed to change $($user.name) to the employee location of Plano2"

            Pause
        
        }
        
    Get-UserOptionsMenu -account $user      

    }

    if($answer2 -eq "n"){Get-UserOptionsMenu -account $user}

}
if($answer -eq 3){

    while($user.extensionattribute14 -eq "Greenville1"){

        Write-Host -ForegroundColor Red "$($user.name) already has an employee location of Greenville1!"

        Sleep 2

        Modify-Attribute14 -user $user

    }
    
    $answer2 = Read-Host "Are you sure you want to change $($user.name) to the employee location of Greenville1? y/n"

    while(($answer2 -ne "y") -and ($answer2 -ne "n")){
    
        Write-Host -ForegroundColor Red "Invalid Selection, Try Again!"

        $answer2 = Read-Host "Are you sure you want to change $($user.name) to the employee location of Greenville1? y/n"
    }

    if($answer2 -eq "y"){
    
        try{
        
            Set-ADUser $user -replace @{extensionattribute14="Greenville1"} -ErrorAction Stop

            Log-This "Successfully changed $($user.name) to the employee location of Greenville1"

            Pause
        
        }
        
        catch{

            Write-Host -ForegroundColor Red "Failed to change $($user.name) to the employee location of Greenville1"
        
            Log-Only "Failed to change $($user.name) to the employee location of Greenville1"

            Pause
        
        }
        
    Get-UserOptionsMenu -account $user      

    }

    if($answer2 -eq "n"){Get-UserOptionsMenu -account $user}

}
if($answer -eq 4){

    while($user.extensionattribute14 -eq "HongKong1"){

        Write-Host -ForegroundColor Red "$($user.name) already has an employee location of HongKong1!"

        Sleep 2

        Modify-Attribute14 -user $user

    }
    
    $answer2 = Read-Host "Are you sure you want to change $($user.name) to the employee location of HongKong1? y/n"

    while(($answer2 -ne "y") -and ($answer2 -ne "n")){
    
        Write-Host -ForegroundColor Red "Invalid Selection, Try Again!"

        $answer2 = Read-Host "Are you sure you want to change $($user.name) to the employee location of HongKong1? y/n"
    }

    if($answer2 -eq "y"){
    
        try{
        
            Set-ADUser $user -replace @{extensionattribute14="HongKong1"} -ErrorAction Stop

            Log-This "Successfully changed $($user.name) to the employee location of HongKong1"

            Pause
        
        }
        
        catch{

            Write-Host -ForegroundColor Red "Failed to change $($user.name) to the employee location of HongKong1"
        
            Log-Only "Failed to change $($user.name) to the employee location of HongKong1"

            Pause
        
        }
        
    Get-UserOptionsMenu -account $user      

    }

    if($answer2 -eq "n"){Get-UserOptionsMenu -account $user}

}
if($answer -eq 5){

    while($user.extensionattribute14 -eq "Singapore1"){

        Write-Host -ForegroundColor Red "$($user.name) already has an employee location of Singapore1!"

        Sleep 2

        Modify-Attribute14 -user $user

    }
    
    $answer2 = Read-Host "Are you sure you want to change $($user.name) to the employee location of Singapore1? y/n"

    while(($answer2 -ne "y") -and ($answer2 -ne "n")){
    
        Write-Host -ForegroundColor Red "Invalid Selection, Try Again!"

        $answer2 = Read-Host "Are you sure you want to change $($user.name) to the employee location of Singapore1? y/n"
    }

    if($answer2 -eq "y"){
    
        try{
        
            Set-ADUser $user -replace @{extensionattribute14="Singapore1"} -ErrorAction Stop

            Log-This "Successfully changed $($user.name) to the employee location of Singapore1"

            Pause
        
        }
        
        catch{

            Write-Host -ForegroundColor Red "Failed to change $($user.name) to the employee location of Singapore1"
        
            Log-Only "Failed to change $($user.name) to the employee location of Singapore1"

            Pause
        
        }
        
    Get-UserOptionsMenu -account $user      

    }

    if($answer2 -eq "n"){Get-UserOptionsMenu -account $user}

}
if($answer -eq 6){

    while($user.extensionattribute14 -eq "Malaysia1"){

        Write-Host -ForegroundColor Red "$($user.name) already has an employee location of Malaysia1!"

        Sleep 2

        Modify-Attribute14 -user $user

    }
    
    $answer2 = Read-Host "Are you sure you want to change $($user.name) to the employee location of Malaysia1? y/n"

    while(($answer2 -ne "y") -and ($answer2 -ne "n")){
    
        Write-Host -ForegroundColor Red "Invalid Selection, Try Again!"

        $answer2 = Read-Host "Are you sure you want to change $($user.name) to the employee location of Malaysia1? y/n"
    }

    if($answer2 -eq "y"){
    
        try{
        
            Set-ADUser $user -replace @{extensionattribute14="Malaysia1"} -ErrorAction Stop

            Log-This "Successfully changed $($user.name) to the employee location of Malaysia1"

            Pause
        
        }
        
        catch{

            Write-Host -ForegroundColor Red "Failed to change $($user.name) to the employee location of Malaysia1"
        
            Log-Only "Failed to change $($user.name) to the employee location of Malaysia1"

            Pause
        
        }
        
    Get-UserOptionsMenu -account $user      

    }

    if($answer2 -eq "n"){Get-UserOptionsMenu -account $user}

}
if($answer -eq 7){

    while($user.extensionattribute14 -eq "Manila1"){

        Write-Host -ForegroundColor Red "$($user.name) already has an employee location of Manila1!"

        Sleep 2

        Modify-Attribute14 -user $user

    }
    
    $answer2 = Read-Host "Are you sure you want to change $($user.name) to the employee location of Manila1? y/n"

    while(($answer2 -ne "y") -and ($answer2 -ne "n")){
    
        Write-Host -ForegroundColor Red "Invalid Selection, Try Again!"

        $answer2 = Read-Host "Are you sure you want to change $($user.name) to the employee location of Manila1? y/n"
    }

    if($answer2 -eq "y"){
    
        try{
        
            Set-ADUser $user -replace @{extensionattribute14="Manila1"} -ErrorAction Stop

            Log-This "Successfully changed $($user.name) to the employee location of Manila1"

            Pause
        
        }
        
        catch{

            Write-Host -ForegroundColor Red "Failed to change $($user.name) to the employee location of Manila1"
        
            Log-Only "Failed to change $($user.name) to the employee location of Manila1"

            Pause
        
        }
        
    Get-UserOptionsMenu -account $user      

    }

    if($answer2 -eq "n"){Get-UserOptionsMenu -account $user}

}
if($answer -eq 8){

    while($user.extensionattribute14 -eq "Taipei1"){

        Write-Host -ForegroundColor Red "$($user.name) already has an employee location of Taipei1!"

        Sleep 2

        Modify-Attribute14 -user $user

    }
    
    $answer2 = Read-Host "Are you sure you want to change $($user.name) to the employee location of Taipei1? y/n"

    while(($answer2 -ne "y") -and ($answer2 -ne "n")){
    
        Write-Host -ForegroundColor Red "Invalid Selection, Try Again!"

        $answer2 = Read-Host "Are you sure you want to change $($user.name) to the employee location of Taipei1? y/n"
    }

    if($answer2 -eq "y"){
    
        try{
        
            Set-ADUser $user -replace @{extensionattribute14="Taipei1"} -ErrorAction Stop

            Log-This "Successfully changed $($user.name) to the employee location of Taipei1"

            Pause
        
        }
        
        catch{

            Write-Host -ForegroundColor Red "Failed to change $($user.name) to the employee location of Taipei1"
        
            Log-Only "Failed to change $($user.name) to the employee location of Taipei1"

            Pause
        
        }
        
    Get-UserOptionsMenu -account $user      

    }

    if($answer2 -eq "n"){Get-UserOptionsMenu -account $user}

}
if($answer -eq 9){

    while($user.extensionattribute14 -eq "Australia1"){

        Write-Host -ForegroundColor Red "$($user.name) already has an employee location of Australia1!"

        Sleep 2

        Modify-Attribute14 -user $user

    }
    
    $answer2 = Read-Host "Are you sure you want to change $($user.name) to the employee location of Australia1? y/n"

    while(($answer2 -ne "y") -and ($answer2 -ne "n")){
    
        Write-Host -ForegroundColor Red "Invalid Selection, Try Again!"

        $answer2 = Read-Host "Are you sure you want to change $($user.name) to the employee location of Australia1? y/n"
    }

    if($answer2 -eq "y"){
    
        try{
        
            Set-ADUser $user -replace @{extensionattribute14="Australia1"} -ErrorAction Stop

            Log-This "Successfully changed $($user.name) to the employee location of Australia1"

            Pause
        
        }
        
        catch{

            Write-Host -ForegroundColor Red "Failed to change $($user.name) to the employee location of Australia1"
        
            Log-Only "Failed to change $($user.name) to the employee location of Australia1"

            Pause
        
        }
        
    Get-UserOptionsMenu -account $user      

    }

    if($answer2 -eq "n"){Get-UserOptionsMenu -account $user}

}
if($answer -eq 10){

    while($user.extensionattribute14 -eq "Remote"){

        Write-Host -ForegroundColor Red "$($user.name) already has an employee location of Remote!"

        Sleep 2

        Modify-Attribute14 -user $user

    }
    
    $answer2 = Read-Host "Are you sure you want to change $($user.name) to the employee location of Remote? y/n"

    while(($answer2 -ne "y") -and ($answer2 -ne "n")){
    
        Write-Host -ForegroundColor Red "Invalid Selection, Try Again!"

        $answer2 = Read-Host "Are you sure you want to change $($user.name) to the employee location of Remote? y/n"
    }

    if($answer2 -eq "y"){
    
        try{
        
            Set-ADUser $user -replace @{extensionattribute14="Remote"} -ErrorAction Stop

            Log-This "Successfully changed $($user.name) to the employee location of Remote"

            Pause
        
        }
        
        catch{

            Write-Host -ForegroundColor Red "Failed to change $($user.name) to the employee location of Remote"
        
            Log-Only "Failed to change $($user.name) to the employee location of Remote"

            Pause
        
        }
        
    Get-UserOptionsMenu -account $user      

    }

    if($answer2 -eq "n"){Get-UserOptionsMenu -account $user}

}
if($answer -eq 11){

    while($user.extensionattribute14 -eq "Home"){

        Write-Host -ForegroundColor Red "$($user.name) already has an employee location of Home!"

        Sleep 2

        Modify-Attribute14 -user $user

    }
    
    $answer2 = Read-Host "Are you sure you want to change $($user.name) to the employee location of Home? y/n"

    while(($answer2 -ne "y") -and ($answer2 -ne "n")){
    
        Write-Host -ForegroundColor Red "Invalid Selection, Try Again!"

        $answer2 = Read-Host "Are you sure you want to change $($user.name) to the employee location of Home? y/n"
    }

    if($answer2 -eq "y"){
    
        try{
        
            Set-ADUser $user -replace @{extensionattribute14="Home"} -ErrorAction Stop

            Log-This "Successfully changed $($user.name) to the employee location of Home"

            Pause
        
        }
        
        catch{

            Write-Host -ForegroundColor Red "Failed to change $($user.name) to the employee location of Home"
        
            Log-Only "Failed to change $($user.name) to the employee location of Home"

            Pause
        
        }
        
    Get-UserOptionsMenu -account $user      

    }

    if($answer2 -eq "n"){Get-UserOptionsMenu -account $user}

}
if($answer -eq 12){Get-UserOptionsMenu -account $user}
if($answer -eq 13){Get-MainMenu}


else{Write-Host -ForegroundColor Red "Invalid Selection, Try Again." 
    sleep 2
    Get-UserOptionsMenu -account $account
    }

}

function Modify-UserADGroups ($user) {




}

function Modify-UserStatus ($user) {

    if($user.enabled -eq $true){

        Log-This "$($user.name) is Enabled!"
    
        $answer = Read-Host "Would you like to set to Disabled? y/n"

        while(($answer -ne "y") -and ($answer -ne "n")){
        
            Write-Host -ForegroundColor Red "Invalid Selection, Please Try Again"
        
            $answer = Read-Host "Would you like to set to Disabled? y/n"
        
        }

        if($answer -eq "y"){
        
            try{
            
                Set-ADUser $user -Enabled $false -ErrorAction Stop

                Log-This "Successfully Disabled the account $($user.name)"
            
            }

            catch{
            
                Write-Host -ForegroundColor Red "Failed to Disable the account $($user.name)"

                Log-Only "Failed to Disable the account $($user.name)"
            
            }
            
            Pause

            Get-UserOptionsMenu -account $user        
        
        }

        if($answer -eq "n"){Get-UserOptionsMenu -account $user}
    
    }

    if($user.enabled -eq $false){

        Log-This "$($user.name) is Disabled!"
    
        $answer = Read-Host "Would you like to set to Enabled? y/n"

        while(($answer -ne "y") -and ($answer -ne "n")){
        
            Write-Host -ForegroundColor Red "Invalid Selection, Please Try Again"
        
            $answer = Read-Host "Would you like to set to Enabled? y/n"
        
        }

        if($answer -eq "y"){
        
            try{
            
                Set-ADUser $user -Enabled $True -ErrorAction Stop

                Log-This "Successfully Enabled the account $($user.name)"
            
            }

            catch{
            
                Write-Host -ForegroundColor Red "Failed to Enable the account $($user.name)"

                Log-Only "Failed to Enable the account $($user.name)"
            
            }
            
            Pause

            Get-UserOptionsMenu -account $user        
        
        }

        if($answer -eq "n"){Get-UserOptionsMenu -account $user}
    
    }


}

function Add-CalendarDelegate ($user) {

    cls

    Log-This -message "
    
________________________________________

Add Calendar Delegate to Offboarded User
________________________________________



"

    Log-This "Searching for $($user.name)'s Mailbox..."    

    try{

        $Mailbox = Get-Mailbox -Identity $user.UserPrincipalName | select -Property * -ErrorAction Stop

        Log-This "Found Mailbox for $($mailbox.name)!"

    }

    catch{
    
        Write-Host -ForegroundColor Red "MAILBOX NOT FOUND"

        Log-Only "MAILBOX NOT FOUND"

        Get-UserOptionsMenu -account $user
    
    }

    Write-Host -ForegroundColor Cyan "
    
Current Delegates for this Calendar    
    "
    
    Get-MailboxFolderPermission -Identity "$($mailbox.UserPrincipalName):\Calendar" | select user, AccessRights

    $searchName = Read-Host "
    
Please type the name of the person you would like to set as calendar delegate for $($user.name). 
Press 'q' to return to the User Tasks Menu
    
"

    if($searchName -eq "q"){Get-UserOptionsMenu -account $user}

    elseif($searchName){
    
        $ADUsers = Get-ADUser -f "name -like '*$searchName*'" -Properties * | sort -Unique

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

                    try{

                        Log-This "Verifying that $($selectedUser.name) has a Mailbox"

                        $Global:ErrorActionPreference = 'Stop'
                    
                        $selectedUserMailbox = Get-Mailbox -Identity $selectedUser.UserPrincipalName | select -Property * -ErrorAction 'Stop'

                        Log-This "Verified!"
                    
                    }

                    catch{
                    
                        Write-Host -ForegroundColor Red "MAILBOX NOT FOUND!"

                        Sleep 2

                        Add-CalendarDelegate -user $user
                    
                    }
                
                    try{

                        $Global:ErrorActionPreference = 'Stop'
                    
                        Add-MailboxFolderPermission -Identity "$($mailbox.UserPrincipalName):\Calendar" -User $selectedUsermailbox.UserPrincipalName -AccessRights owner -WarningAction SilentlyContinue -ErrorAction 'Stop'

                        Log-This "$($selectedUsermailbox.name) is now a delegate for $($mailbox.name)'s calendar."

                    }

                    catch{
                    
                        Write-Host -ForegroundColor Red "Failed to add $($selectedUsermailbox.name) as calendar delegate for $($mailbox.name)."

                        Log-Only "Failed to add $($selectedUsermailbox.name) as calendar delegate for $($mailbox.name)."
                    
                    }

                    Pause

                    Get-UserOptionsMenu -account $user
                
                }

                if($answer2 -eq "n"){Add-CalendarDelegate -user $user}
            
            }     




        }

        pause

        Get-UserOptionsMenu -account $user
    
    }

    else{

        cls
        
        Write-Host -ForegroundColor DarkYellow "You didn't type anything weirdo"

        Sleep 2

        Get-SearchUserMenu
    
    }


}

function Add-MailboxDelegate ($user) {

    cls

    Log-This -message "
    
_______________________________________

Add Mailbox Delegate to Offboarded User
_______________________________________



"

    Log-This "Searching for $($user.name)'s Mailbox..."    

    try{

        $Mailbox = Get-Mailbox -Identity $user.UserPrincipalName | select -Property * -ErrorAction Stop

        Log-This "Found Mailbox for $($mailbox.name)!"

    }

    catch{
    
        Write-Host -ForegroundColor Red "MAILBOX NOT FOUND"

        Log-Only "MAILBOX NOT FOUND"

        Get-UserOptionsMenu -account $user
    
    }

    $searchName = Read-Host "
    
Please type the name of the person you would like to set as mailbox delegate for $($user.name). 
Press 'q' to return to the User Tasks Menu
    
"

    if($searchName -eq "q"){Get-UserOptionsMenu -account $user}

    elseif($searchName){
    
        $ADUsers = Get-ADUser -f "name -like '*$searchName*'" -Properties * | sort -Unique

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

                    try{

                        Log-This "Verifying that $($selectedUser.name) has a Mailbox"

                        $Global:ErrorActionPreference = 'Stop'
                    
                        $selectedUserMailbox = Get-Mailbox -Identity $selectedUser.UserPrincipalName | select -Property * -ErrorAction 'Stop'

                        Log-This "Verified!"
                    
                    }

                    catch{
                    
                        Write-Host -ForegroundColor Red "MAILBOX NOT FOUND!"

                        Sleep 2

                        Add-MailboxDelegate -user $user
                    
                    }
                
                    try{

                        $Global:ErrorActionPreference = 'Stop'                      

                        Add-MailboxPermission -Identity $mailbox.UserPrincipalName -user $selectedUsermailbox.UserPrincipalName -AccessRights FullAccess -InheritanceType all -AutoMapping $true -WarningAction SilentlyContinue -ErrorAction 'Stop' 

                        Log-This "$($selectedUsermailbox.name) now has full access to $($mailbox.name)'s mailbox."

                    }

                    catch{
                        
                        Write-Host -ForegroundColor Red "Failed to provide $($selectedUsermailbox.name) with full access to $($mailbox.name)'s mailbox."

                        Log-Only "Failed to provide $($selectedUsermailbox.name) with full access to $($mailbox.name)'s mailbox."
                    
                    }

                    Pause

                    Get-UserOptionsMenu -account $user
                
                }

                if($answer2 -eq "n"){Add-MailboxDelegate -user $user}
            
            }     




        }

        pause

        Get-UserOptionsMenu -account $user
    
    }

    else{

        cls
        
        Write-Host -ForegroundColor DarkYellow "You didn't type anything weirdo"

        Sleep 2

        Get-SearchUserMenu
    
    }


}


#---------------------------------
#endregion
#---------------------------------
#endregion
#---------------------------------
#region Setup
#---------------------------------

cls

#if (-Not ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] 'Administrator')) {
#    if ([int](Get-CimInstance -Class Win32_OperatingSystem | Select-Object -ExpandProperty BuildNumber) -ge 6000) {
#        $CommandLine = "-File `"" + $MyInvocation.MyCommand.Path + "`" " + $MyInvocation.UnboundArguments
#        Start-Process -FilePath PowerShell.exe -Verb Runas -ArgumentList $CommandLine
#        Exit
#    }
#}

"


____________________________________________

Offboarding Tasks Log | $($(Get-Date).ToString('yyyy-MM-dd::hh:mm:ss'))
____________________________________________



" | Out-File $logfile

Log-This -message "[INFO] | Offboarding Tasks"
Log-This -message "[INFO] | Script Ran by: $env:USERNAME"

$username = "sa.userimac@wvholdings.onmicrosoft.com"
$password = "9fwsk1(Yvv#S"
$sp = $password | ConvertTo-SecureString -AsPlainText -force
$plainCred = New-Object System.Management.Automation.PSCredential -ArgumentList $username, $sp

Connect-MsolService -Credential $plainCred

sleep 2

cls

echo "                                                                                           "
echo "                                                                                           "
echo "                                                                                           "
echo "                                                                                           "
echo "                                                                                           "
echo "                                                                                           "
echo "                                                                                           "
echo "                                                                                           "
echo "                                                                                           "
echo "                                                                                           "
echo "                                                                                           "
echo "                                                                                           "
echo "                                                                                           "
echo "                                                                                           "
echo "                                                                                           "
echo "                     ____________________________________________________________________  "
echo "                     |************|****************************************|************|  "
echo "                     |************|****************************************|************|  "
echo "                     |************|-----------OFFBOARDING--TASKS-----------|************|  "
echo "                     |************|****************************************|************|  "
echo "                     |************|****************************************|************|  "
echo "                     |************|--------WRITTEN-BY:-MIKE-BRATTON--------|************|  "
echo "                     |************|****************************************|************|  "
echo "                     |************|****************************************|************|  "
echo "                     |************|********--------WELCOME!--------********|************|  "
echo "                     |************|****************************************|************|  "
echo "                     |************|****************************************|************|  "
echo "                     |__________________________________________________________________|  "
echo "                      <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>   "
echo "                                                                                           "
echo "                                                                                           "
echo "                                                                                           "
echo "                                                                                           "
echo "                                                                                           "
echo "                                                                                           "
echo "                                                                                           "
echo "                                                                                           "
echo "                                                                                           "
echo "                                                                                           "
echo "                                                                                           "
echo "                                                                                           "
echo "                                                                                           "
echo "                                                                                           "
echo "                                                                                           "

sleep 2

#---------------------------------
#endregion
#---------------------------------
#region Main
#---------------------------------


Get-MainMenu



#---------------------------------
#endregion
#---------------------------------