$edvalsFilestoProcess = @(
("$PSScriptRoot\Input\Edvals\Sem 1,2","2022","FY"),
("$PSScriptRoot\Input\Edvals\Sem 1","2022","S01")
<#,("$PSScriptRoot\Input\Edvals\Sem 2","2022","S02")#>
)

function Read-edvalsData
{
    Param 
    (
        [Parameter(Mandatory=$true)][string]$dataPath,
        [Parameter(Mandatory=$true)][string]$schoolYear,
        [Parameter(Mandatory=$true)][string]$schoolSession
    )
    Write-Host "$dataPath\Section.csv"
    # Import Section Data
    #$sectionData = $null
    $sectionData =  Import-Csv "$dataPath\Section.csv" | SELECT-Object 'SIS ID','Course Subject','Course Name','YearLevel','CorrectedSubject','academicSession','schoolYear','classTitle'

    # Process Section Data (Add Leading 0's, Add School Year, Add Academic Session etc)
    foreach ($subject in $sectionData)
    {
        <#$subject.'SIS ID' = $subject.'SIS ID' -replace '-and-','&'
        $subject.'SIS ID' = $subject.'SIS ID' -replace '-slash-','/'
        $subject.'SIS ID' = $subject.'SIS ID' -replace '-plus-','+'
        $subject.'SIS ID' = $subject.'SIS ID' -replace '  ',' '#>

        $subject.'Course Name' = $subject.'Course Name' -replace '-and-',' & '
        $subject.'Course Name' = $subject.'Course Name' -replace '-slash-','/'
        $subject.'Course Name' = $subject.'Course Name' -replace ' -plus-','+'
        $subject.'Course Name' = $subject.'Course Name' -replace '  ',' '

        $subject.'Course Subject' = $subject.'Course Subject' -replace '-and-','&'
        $subject.'Course Subject' = $subject.'Course Subject' -replace '-slash-','/'
        $subject.'Course Subject' = $subject.'Course Subject' -replace '-plus-','+'
        $subject.'Course Subject' = $subject.'Course Subject' -replace '  ',' '

        if( $subject.'Course Subject'.Substring(0,1) -match '[7-9]')
        {
            $subject.'CorrectedSubject' = "0$($subject.'Course Subject')"
        }
        else
        {
            $subject.'CorrectedSubject' = $subject.'Course Subject'
        }

        if ($subject.'CorrectedSubject'.Substring(0,2) -le 12)
        {
            $subject.'YearLevel' = $subject.'CorrectedSubject'.Substring(0,2)
        }

        if((($subject.'SIS ID').Substring(0,1) -match "^\d+$"))
        {
            if(($subject.'SIS ID').Substring(($subject.'SIS ID').IndexOf('.')+1,(($subject.'SIS ID').LastIndexOf('-') - ($subject.'SIS ID').IndexOf('.') - 1)) -match "^\d+$")
            {
                # Do something
                $tempCourseName = ("$($subject.YearLevel).")+(($subject.'SIS ID').Substring(($subject.'SIS ID').IndexOf('.')+1))
                $tempSubject = $subject.'Course Name'
                $subject.classTitle = $tempCourseName -replace "-"," $tempSubject - "
            }
            else
            {
                $tempCourseName = $subject.YearLevel+(($subject.'SIS ID').Substring(($subject.'SIS ID').IndexOf('.')+1))
                $tempSubject = $subject.'Course Name'
                $subject.classTitle = $tempCourseName -replace "-"," $tempSubject - "
            }
        }
        else
        {
                $tempCourseName = ($subject.'Course Name')+" "+(($subject.'SIS ID').Substring(($subject.'SIS ID').IndexOf('.')+1))
                $tempSubject = $subject.'Course Name'
                $subject.classTitle = ($tempCourseName -replace "-"," - ")
        }

        $subject.classTitle = $subject.classTitle -replace "Year 07 ",""
        $subject.classTitle = $subject.classTitle -replace "Year 08 ",""
        $subject.classTitle = $subject.classTitle -replace "Year 09 ",""
        $subject.classTitle = $subject.classTitle -replace "Year 10 ",""
        $subject.classTitle = $subject.classTitle -replace "Year 11 ",""
        $subject.classTitle = $subject.classTitle -replace "Year 12 ",""
        $subject.'Course Name' = $subject.'Course Name' -replace "Year 07 ",""
        $subject.'Course Name' = $subject.'Course Name' -replace "Year 08 ",""
        $subject.'Course Name' = $subject.'Course Name' -replace "Year 09 ",""
        $subject.'Course Name' = $subject.'Course Name' -replace "Year 10 ",""
        $subject.'Course Name' = $subject.'Course Name' -replace "Year 11 ",""
        $subject.'Course Name' = $subject.'Course Name' -replace "Year 12 ",""

        if ($subject.YearLevel -eq "" -or $null -eq $subject.YearLevel)
        {
            $subject.YearLevel = "Other"
        }


        $subject.schoolYear = $schoolYear
        $subject.academicSession = $schoolSession

    }

    $sectionData = $sectionData  | Sort-Object -Property "YearLevel","CorrectedSubject"

    
    # Student Class Enrollments Transform

    $edvalsStudentEnrollments = Import-Csv "$dataPath\StudentEnrollment.csv" | Sort-Object -Property 'Section SIS ID'

    foreach ($studentSession in $edvalsStudentEnrollments)
    {
        <#$studentSession.'Section SIS ID' = $studentSession.'Section SIS ID' -replace '-and-','&'
        $studentSession.'Section SIS ID' = $studentSession.'Section SIS ID' -replace '-slash-','/'
        $studentSession.'Section SIS ID' = $studentSession.'Section SIS ID' -replace '-plus-','+'
        $studentSession.'Section SIS ID' = $studentSession.'Section SIS ID' -replace '  ',' '#>

        if( $studentSession.'Section SIS ID'.Substring(0,1) -match '[7-9]')
        {
            $studentSession.'Section SIS ID' = "0$($studentSession.'Section SIS ID')"
        }
    }

    $edvalsStaffRoster = Import-Csv "$dataPath\TeacherRoster.csv" | Sort-Object -Property 'Section SIS ID'

    # Staff Class Roster Transform
    foreach ($staffSession in $edvalsStaffRoster)
    {
        <#$staffSession.'Section SIS ID' = $staffSession.'Section SIS ID' -replace '-and-','&'
        $staffSession.'Section SIS ID' = $staffSession.'Section SIS ID' -replace '-slash-','/'
        $staffSession.'Section SIS ID' = $staffSession.'Section SIS ID' -replace '-plus-','+'
        $staffSession.'Section SIS ID' = $staffSession.'Section SIS ID' -replace '  ',' '#>
        
        if( $staffSession.'Section SIS ID'.Substring(0,1) -match '[7-9]')
        {
            $staffSession.'Section SIS ID' = "0$($staffSession.'Section SIS ID')"
        }
    
    }
    
    return $sectionData, $edvalsStudentEnrollments, $edvalsStaffRoster
}


###################### COURSES OUTPUT ###########################

$edvalsCourses = $sectionData | Sort-Object -Unique -Property "YearLevel","CorrectedSubject"
$tempCourse = @()

foreach ($edvalsCourse in $edvalsCourses)
{
    $tempOutput = New-Object PSObject
    $tempOutput | Add-Member -MemberType NoteProperty -Name "sourceId" -Value ("$($edvalsCourse.CorrectedSubject)-$year")
    $tempOutput | Add-Member -MemberType NoteProperty -Name "orgSourcedId" -Value ("789301")
    if ($edvalsCourse.YearLevel  -match "^\d+$")
    {
        $tempOutput | Add-Member -MemberType NoteProperty -Name "title" -Value ("Year $($edvalsCourse.YearLevel) $($edvalsCourse.'Course Name')")
    }
    else
    {
        $tempOutput | Add-Member -MemberType NoteProperty -Name "title" -Value ($edvalsCourse.'Course Name')
    }
    
    $tempOutput | Add-Member -MemberType NoteProperty -Name "grade" -Value ($edvalsCourse.YearLevel)
        
    $tempCourse += $tempOutput
}
####################### CLASSES OUTPUT ############################
$edvalsClasses = $sectionData
$tempClasses = @()

foreach ($edvalsClass in $edvalsClasses)
{
    $tempOutput = New-Object PSObject
    $tempOutput | Add-Member -MemberType NoteProperty -Name "sourceId" -Value ($edvalsClass.'SIS ID')
    $tempOutput | Add-Member -MemberType NoteProperty -Name "orgSourcedId" -Value ("789301")
    $tempOutput | Add-Member -MemberType NoteProperty -Name "title" -Value ($edvalsClass.classTitle)
    #$tempOutput | Add-Member -MemberType NoteProperty -Name "sessionSourcedId" -Value ($edvalsCourse.YearLevel)
    $tempOutput | Add-Member -MemberType NoteProperty -Name "courseSourcedId" -Value ("$($edvalsClass.CorrectedSubject)-$year")    
    $tempClasses += $tempOutput
}


#$edvalsClasses, $edvalsStudentEnrollment, $endvalsStaffRoster = Read-edvalsData -dataPath ".\SDS\Sem 1,2\" -schoolYear "2021" -schoolSession "2021FY"
$edvalsClasses = $null
$edvalsStudentEnrollment = $null
$endvalsStaffRoster = $null

foreach ($file in $edvalsFilestoProcess)
{
    $outputClassesTemp = $null
    $outputEnrollmentTemp = $null
    $outputRosterTemp = $null
    $outputClassesTemp, $outputEnrollmentTemp, $outputRosterTemp = Read-edvalsData -dataPath $file[0] -schoolYear $file[1] -schoolSession "$($file[1])$($file[2])"
    $edvalsClasses += $outputClassesTemp
    $edvalsStudentEnrollment += $outputEnrollmentTemp
    $endvalsStaffRoster += $outputRosterTemp

}