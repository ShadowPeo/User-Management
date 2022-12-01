$newClasses = Import-CSV -path ''

$essentialAssessment = $newClasses | Select-Object -Property 'First Name', 'Last Name', 'Student ID', @{label="Rollover Class 1";expression={$($_."New Class")}}

$essentialAssessment | Export-CSV "$PSScriptRoot\EssentialAssessment.csv" -Encoding ASCII  -NoTypeInformation


$adUsers = Get-ADUser -Properties mail -SearchBase '' -Filter *

$matific = @()

$matificStudents = $newClasses | Select-Object -Property @{label="Student First Name*";expression={$($_."First Name")}},@{label="Student Last Name*";expression={$($_."Last Name")}},@{label="2023 Class Name*";expression={$($_."New Class")}},'2023 Grade/Year Level*', "2023 Grade/Year Level*",'Teacher Email*', 'Teacher First Name*', 'Teacher Last Name*', 'New Teacher'
$matificOutput = @()
for ($i = 0; $i -lt ($matificStudents.Count); $i++)
{
    $yearLevel = $null
    $yearLevel = ($matificStudents[$i].'2023 Class Name*').Substring(0,1)

    if ($yearLevel -ieq 'F' -or ($yearLevel -ge 1 -and $yearLevel -le 2))
    {
        Write-Host $matificStudents[$i].'Student First Name*'
        $matificStudents[$i].'2023 Grade/Year Level*' = $yearLevel

        
        $matificOutput += $matificStudents[$i]

    }
}

