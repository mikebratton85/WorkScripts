$Licenses = @(
    [PSCUSTOMOBJECT]@{ Title="e1"; Sku="STANDARDPACK"; Group="SG 365 License E1"; Enable=@("TEAMS1")}

    [PSCUSTOMOBJECT]@{ Title="e3"; Sku="ENTERPRISEPACK"; Group="SG 365 License E3"; Enable=@("TEAMS1")}

    [PSCUSTOMOBJECT]@{ Title="e5"; Sku="ENTERPRISEPREMIUM"; Group="SG 365 License E5"; Enable=@("TEAMS1")}
)

$logfile = "C:\logs\ApplyTeamsLicense-$($(Get-Date).ToString('MM-dd-yyyy')).txt"

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

Function Try-Enabling(){
    [CmdletBinding()]
    Param(
        [Parameter(
            Mandatory = $true,
            Position = 0,
            ValueFromPipeline=$true
        )]
        [ValidateNotNullOrEmpty()]
        $User,
        [Parameter(
            Mandatory = $true,
            ValueFromPipelineByPropertyName=$true,
            Position = 1
        )]
        [ValidateNotNullOrEmpty()]
        $Sku,
        [Parameter(
            Mandatory = $true,
            Position = 2
        )]
        $Enable,
        [Parameter(
            Mandatory = $true,
            Position = 3
        )]
        $Title
    )

    begin{
        $FN = "Try-Enabling"
        $Enabled = @()
        $NoEnabled = @()
    }

    Process{
    
        if(($User.IsLicensed) -and ($user.Licenses.AccountSkuID.Contains($Sku))){
            
            Log-It "[OK] $FN | Trying to enable any disabled sub-licensing for $($User.userprincipalname) under $($Title)"
            
            $MSOLLO = New-MsolLicenseOptions -AccountSkuId $Sku -DisabledPlans $null
            
            Try{
                $User | Set-MsolUserLicense -LicenseOptions $MSOLLO -ErrorAction Stop
                $Enabled += $user
                Log-Success "[SUCCESS] $FN | Licensed $($User.userprincipalname) with $($enable)"
            }
            
            Catch{
                Log-Failure -Message "[ERROR] $FN | Failed to enable $($Enable) for $($user.userprincipalname) | $($_.Exception.Message)"
                $NoEnabled += $user
            }

        }
        
        else{
            Log-Failure -Message "[ERROR] $FN | $($user.userprincipalname) is not licensed with E1, E3, or E5"
        }

    }

    end{ 
        return @{enabled=$($enabled.count);noenabled=$($NoEnabled.count)} 
    }

}

$username = "sa.userimac@wvholdings.onmicrosoft.com"
$password = "9fwsk1(Yvv#S"
$secureStringPwd = $password | ConvertTo-SecureString -AsPlainText -Force 
$creds = New-Object System.Management.Automation.PSCredential -ArgumentList $username, $secureStringPwd

Connect-MsolService -Credential $Creds

$365Users = Get-msoluser -All

Log-It -message "[OK] Setup | Retrieved $($365Users.count) total users"
    
$Licensed365Users = $365Users | where{$_.isLicensed -eq $True}

Log-It -message "[OK] Setup | Retrieved $($Licensed365Users.count) licensed users"

ForEach($License in $Licenses){
    
    # Gets the matching sku for this license
    $MSOLSKU = Get-MsolAccountSku | where {$_.accountskuid -like "*$($License.Sku)*"} | % accountskuid 

    # Stores each AD group's members AD objects with country
    $Group = $License.Group
    $ADGroupMembers = Get-ADGroup -f {Name -like $Group} | Get-ADGroupMember | Get-ADUser -Properties userprincipalname, country | Where {$_.ObjectClass -EQ "user"}

    # Stores each 365 account that is licensed for a specific license as a key pair in $CurrentlyLicensed
    $CurrentlyLicensed = $Licensed365Users | where{$_.licenses.accountskuid.contains($MSOLSKU)} 

    If(($ADGroupMembers.count -gt 0) -and ($CurrentlyLicensed.Count -gt 0)){
        # Finds the difference between the users already licensed and those a memberof the AD group
        $LicenseResults = Compare-Object -ReferenceObject $ADGroupMembers.userprincipalname -DifferenceObject $CurrentlyLicensed.userprincipalname
    }
        
    # Processes each users in the comparison list for licensing operations
    $RecentlyLicensed[$License.Title] = @()
    $NotLicensed[$License.Title] = @()
    $LicenseResults | Try-Enabling -Sku $MSOLSKU -Enable $License.Enable -Title $License.Title
}

