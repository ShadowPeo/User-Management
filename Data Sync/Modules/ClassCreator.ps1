function Sync-ClassCreatorStudents
{
    param 
    (
        [Parameter( Mandatory = $true)][System.Array]$students,
        [Parameter( Mandatory = $true)][System.Array]$classCreator
    )

    $classCreator = $classCreator | Select-Object 'Student ID', 'New Year Level', 'New Class' | Sort-Object 'Student ID'
    foreach ($student in $students)
    {
        $studentID = $student.STKEY
        $classCreatorStudent = $classCreator | Where-Object { $_.'Student ID' -eq $studentID }
        if ($classCreatorStudent)
        {
            if ($classCreatorStudent.'New Year Level' -ieq "P")
            {
                $classCreatorStudent.'New Year Level' = "00"
            }
            elseif (($classCreatorStudent.'New Year Level').Length -eq 1)
            {
                $classCreatorStudent.'New Year Level' = "0$($classCreatorStudent.'New Year Level')"
            }
            $student.SCHOOL_YEAR = $classCreatorStudent.'New Year Level'
            
            if ($classCreatorStudent.'New Class' -notmatch '([01][0-9F])[A-Z]*')
            {
                $classCreatorStudent.'New Class' = "0$($classCreatorStudent.'New Class')"
            }

            $student.HOME_GROUP = $classCreatorStudent.'New Class'
        }
    }
    return $students
}

function Sync-ClassCreatorClasses
{
    param 
    (
        [Parameter( Mandatory = $true)][System.Array]$classes,
        [Parameter( Mandatory = $true)][System.Array]$classCreator,
        [Parameter( Mandatory = $true)][System.Array]$staff
    )
    $classCreator = $classCreator | Select-Object 'New Year Level','New Class','New Teacher' -Unique | Sort-Object 'New Year Level','New Class'
    $classes | Where-Object KGCKEY -ne "ZZZ" | ForEach-Object { $_.ACTIVE = "N"}
    $classes | ForEach-Object { $_.ROOM = ""}
    $classes | ForEach-Object { $_.TEACHER_B = ""}
    $classCreator | Where-Object 'New Year Level' -ieq "P" | ForEach-Object { $_.'New Year Level' = "P"}
    
    foreach ($class in $classCreator)
    {
        $classYear = $class.'New Year Level'
        if ($classYear.'New Class' -notmatch '([01][0-9F])')
        {
            $classYear = "0$classYear"
        }

        $classClass = $class.'New Class'
        if ($classClass -notmatch '([01][0-9F])[A-Z]*')
        {
            $classClass = "0$classClass"
        }
        $classTeacher = $class.'New Teacher'
        $eduHubClass = $classes | Where-Object { $_.KGCKEY -eq $classClass }
        if ($eduHubClass)
        {
            $eduHubClass.ACTIVE = "Y"
            if (($staff | Where-Object {$_.SURNAME -contains ($classTeacher.SubString(($classTeacher.LastIndexOf(" ")+1)))}).Count -gt 1)
            {
                $eduHubClass.TEACHER = ($staff | Where-Object {$_.SURNAME -contains ($classTeacher.SubString(($classTeacher.LastIndexOf(" ")+1))) -and $_.STAFF_STATUS -eq "ACTV"}).SFKEY
            }
            else
            {
                $eduHubClass.TEACHER = ($staff | Where-Object {$_.SURNAME -contains ($classTeacher.SubString(($classTeacher.LastIndexOf(" ")+1)))}).SFKEY
            }
            
            $eduHubClass.DESCRIPTION = "Class $classClass"
        }
        else 
        {
            $tempClass = [PSCustomObject]@{
                KGCKEY = $classClass
                TEACHER = ($staff | Where-Object {$_.SURNAME -contains ($classTeacher.SubString(($classTeacher.LastIndexOf(" ")+1))) -and $_.STAFF_STATUS -eq "ACTV"}).SFKEY
                DESCRIPTION = "Class $classClass"
                ACTIVE = "Y"
            }
            $classes += $tempClass
        }
    }
    return $classes

}

