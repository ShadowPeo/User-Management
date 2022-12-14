$classYear = ""
$classCreatorCSV = ""
$adStaffBaseOU = ""
$fileOutput = "$PSScriptRoot\ClassCreator.csv"

$adStaff = Get-ADUser -Properties enabled, name -SearchBase $adStaffBaseOU -Filter *


$newClasses = Import-CSV -path $classCreatorCSV | Select-Object *, @{Name='STAFF_SIS_ID';Expression={''}},@{Name='Class_Year';Expression={''}},@{Name='Class_School_Year';Expression={$classYear}}

for ($i=0; $i -lt $newClasses.Count; $i++)
{
    $newClasses[$i].'Class_Year' = (($newClasses[$i].'New Class').Substring(0,1)).PadLeft(2,'0')
    if (![string]::IsNullOrWhiteSpace($newClasses[$i].'New Teacher'))
    {
        $tempSTAFF_SIS_ID = $null
        $tempSTAFF_SIS_ID = ($adStaff | Where-Object Name -like "$($newClasses[$i].'New Teacher')*" | Select-Object SamAccountName).SamAccountName
        if($tempSTAFF_SIS_ID.Count -gt 1)
        {
            $tempSTAFF_SIS_ID = ($adStaff | Where-Object {($_.Name -like "$($newClasses[$i].'New Teacher')*") -and ($_.Enabled -eq $true)} | Select-Object SamAccountName).SamAccountName
        }
        $newClasses[$i].'STAFF_SIS_ID' = $tempSTAFF_SIS_ID
    }
}

$newClasses | Export-CSV $fileOutput -Encoding ASCII  -NoTypeInformation