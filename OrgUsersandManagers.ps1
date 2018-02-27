#Org_Users_and_Managers.ps1
#
#Script written by Mike Bratton
#December 21, 2017
#
#
#Variables
$dateFormat = "$($(Get-Date).ToString('yyyy-MM-dd_hh-mm-ss'))"
$CSV = "\\techutil01\c$\CSV\Active_Organization_User_List_$($dateFormat).csv"
$Rows = @()
$table = @()
$NoManager = @()
$NoTitle = @()
$NoCompany = @()
$staff = @()
$partner = @()
$1099 = @()
$agency = @()
$Others = @()

#Pull all enabled users and their managers
$ExtAtr15 = Get-ADUser -f{name -like "*"} -Properties extensionattribute15, enabled, manager, userprincipalname, title, department, company | where {$_.enabled -eq $true} | sort -Unique

#Loop through users to build tables
foreach($user in $ExtAtr15){

    #Exclude the non-user employee types
    if(($user.extensionattribute15 -ne "Staff") -and ($user.extensionattribute15 -ne "1099") -and ($user.extensionattribute15 -ne "Agency") -and ($user.extensionattribute15 -ne "Partner") -and ($user.extensionattribute15 -ne "Others")){Continue}

    #Include the rest
    else{
    
        #Correct manager naming Scheme
        if($user.Manager -ne $null){

            $manager = $user.manager.Split(",")[0].replace("CN=","")

        }
    
        #Notifies there is no manager specified
        else{

            $manager = "Not Specified"

            #Adds to the total count on unspecified managers
            $NoManager += $manager

        }
        
        #Finds manager
        $ManagerAccount = Get-ADUser -f{name -like $manager} -Properties mail

        #Selects manager's email
        $ManagerEmail = $ManagerAccount.mail

        #Notifies there is no manager email
        if($ManagerAccount -eq $null){
        
            $ManagerEmail = "Not Specified"
        
        }

        #Create $title variable
        $title = $user.title

        #Notifies there is no title specified
        if($title -eq $null){

            $title = "Not Specified"

            #Adds to the total count on unspecified titles
            $NoTitle += $title

        }

        #Create $company variable
        $company = $user.company

        #Notifies there is no company specified
        if($company -eq $null){
        
            $company = "Not Specified"

            #Adds to the total count on unspecified company
            $NoCompany += $company
        
        }

        #Add name and manager properties to the table
        $table += [pscustomobject] @{
            
            Name = $user.name;
            Email = $user.UserPrincipalName;
            Title = $title;
            Department = $user.department
            Manager = $manager;
            Manager_Email = $ManagerEmail;
            Employee_Type = $user.extensionattribute15;
            Company = $company;
            
        }

        #Add name and manager row
        $Row = @"

            <tr>
                <td>$($user.name)</td>
                <td>$($user.UserPrincipalName)</td>
                <td>$($title)</td>
                <td>$($user.department)</td>                           
                <td>$($manager)</td>
                <td>$($managerEmail)</td>
                <td>$($user.extensionattribute15)</td>
                <td>$($company)</td>             
            </tr>

"@
    
        #Add row to the $Rows table
        $Rows += $Row

        #Adds each user's employee type to an array
        switch($user.extensionattribute15){

            "Staff" {$staff += $user}
            "Partner" {$partner += $user}
            "1099" {$1099 += $user}
            "Agency" {$agency += $user}
            "Others" {$others += $user}
            default {continue}
    
        }

    }

}

#Body of the report
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
                    Active Employees
                    $(Get-Date)
                </h1>
                <p><u>Full List of Active Employee Accounts in Active Directory</u></p>
                <p><strong>Total Accounts:</strong> $($table.Count)</p>
                <p>
                    <strong>Unspecified Manager Count:</strong> $($NoManager.count)<br>
                    <strong>Unspecified Title Count:</strong> $($NoTitle.count)<br>
                    <strong>Unspecified Company Count:</strong> $($NoCompany.count)
                </p>
                <p>
                    <strong>Total Staff Employees:</strong> $($staff.count)<br>
                    <strong>Total Partners:</strong> $($partner.count)<br>
                    <strong>Total 1099s:</strong> $($1099.count)<br>
                    <strong>Total Agency Contractors:</strong> $($agency.count)<br>
                    <strong>Employees with Type "Other":</strong> $($Others.count)
                </p>
                <table>
                    <thead>
                        <th>Name</th>
                        <th>Email</th>
                        <th>Title</th>
                        <th>Department</th>
                        <th>Manager</th>
                        <th>Manager Email</th>
                        <th>Employee Type</th>
                        <th>Company</th>                     
                    </thead>
                    <tbody>
                        $Rows
                    </tbody>
                </table>
            </body>    
        </html>
"@

#Create the CSV file containing the full table
$table | Export-CSV $CSV

#Email report with attached CSV file
Send-MailMessage -From "Active Employee Report no-reply@worldventures.com" -To infra@worldventures.com -Subject "Active Employees - $((get-date).ToShortDateString())" -Priority High -Body $HTML -BodyAsHtml -SmtpServer "relay.worldventures.com" -Attachments $CSV 



#schatfield@wvholdings.com, sbradshaw@worldventures.com, jmathews@worldventures.com,  