$Script = "C:\Users\mbratton\Documents\WindowsPowerShell\PowerShell\Working_Scripts\CompareADProperties.ps1"

Function Compare-ObjectProperties {
    Param(
        [PSObject]$ReferenceObject,
        [PSObject]$DifferenceObject 
    )
    $objprops = $ReferenceObject | Get-Member -MemberType Property,NoteProperty | % Name
    $objprops += $DifferenceObject | Get-Member -MemberType Property,NoteProperty | % Name
    $objprops = $objprops | Sort | Select -Unique
    $diffs = @()
    foreach ($objprop in $objprops) {
        $diff = Compare-Object $ReferenceObject $DifferenceObject -Property $objprop
        if ($diff) {            
            $diffprops = @{
                PropertyName=$objprop
                Reference=($diff | ? {$_.SideIndicator -eq '<='} | % $($objprop))
                Difference=($diff | ? {$_.SideIndicator -eq '=>'} | % $($objprop))
            }
            $diffs += New-Object PSObject -Property $diffprops
        }        
    }
    if ($diffs) {return ($diffs | Select PropertyName, Reference, Difference)}     
}


$Ref = Read-Host "Type the name of the AD Object you would like to reference"

if($Ref -eq "q"){exit}

$Reference = Get-ADUser -f "name -like '*$Ref*'" -Properties *

$RefObjects = @()

While($Reference -eq $null){

    Write-Host -ForegroundColor Red "There is no AD Object that matches your search, try again"

    $Ref = Read-Host "Type the name of the AD Object you would like to reference"

    $Reference = Get-ADUser -f "name -like '*$Ref*'" -Properties *

}

if($Reference.count -gt 1){

    foreach($object in $Reference){
    
       $RefObjects += $object
    
    }

    $RefObjects.samaccountname

    $Reference = Read-Host "Which of the above AD Objects would you like to reference"

    $Reference = Get-ADUser $Reference -Properties *

}


else{

    $answer = Read-Host "$($Reference.samaccountname) is the only account matching your input. Is this the correct account? y/n"

    while($answer -notlike "y*" -and $answer -notlike "n*"){
    
        Write-Host -ForegroundColor Red "Invalid Selection, Try Again"

        $answer = Read-Host "$($Reference.samaccountname) is the only account matching your input. Is this the correct account? y/n"

        }

        if($answer -eq "y"){}
    
        elseif($answer -eq "n"){Invoke-Expression -command $Script}

}

$Dif = Read-Host "Type the name of the AD Object you would like to see the comparison differences of"

if($Dif -eq "q"){exit}

$Difference = Get-ADUser -f "name -like '*$Dif*'" -Properties *

$DifObjects = @()

While($Difference.count -eq 0){

    Write-Host -ForegroundColor Red "There is no AD Object that matches your search, try again"

    $Dif = Read-Host "Type the name of the AD Object you would like to reference"

    $Difference = Get-ADUser -f "name -like '*$Dif*'" -Properties *

}

if($Difference.count -gt 1){

    foreach($object in $Difference){
    
       $DifObjects += $object
    
    }

    $DifObjects.samaccountname

    $Difference = Read-Host "Which of the above AD Objects would you like to reference"

    $Difference = Get-ADUser $Difference -Properties *

}

else{

    $answer = Read-Host "$($Difference.samaccountname) is the only account matching your input. Is this the correct account? y/n"

    while($answer -notlike "y*" -and $answer -notlike "n*"){
    
        Write-Host -ForegroundColor Red "Invalid Selection, Try Again"

        $answer = Read-Host "$($Difference.samaccountname) is the only account matching your input. Is this the correct account? y/n"

        }

        if($answer -eq "y"){Continue}
    
        elseif($answer -eq "n"){Invoke-Expression -command $Script}

}
Compare-ObjectProperties $Reference $Difference