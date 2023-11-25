param 
    (
        
        #School Details
        [string]$schoolNumber = "<<SCHOOL_ID>>", # Used for export and for import if using CASES File Names, TODO: Add code to pull from server name/bbportal

        #File Locations
        [string]$fileImportLocation = "$PSScriptRoot\..\eduHub Export\Output",
        [string]$fileOutputLocation = "$PSScriptRoot\Output",

        #Processing Handling Varialbles
        [float]$handlingStudentNoExportAfter = 365, #How long to export the data after the staff member or student has left. this is calculated based upon Exit Date, if it does not exist but marked as left they will be exported until exit date is established; 0 Disables export of left students, -1 will always export them
        [string]$exportFormat = "utf8", #Formats supported for output ascii, unicode, utf8, utf32
        [bool]$nextYear = $false,
        [bool]$usePrefNameStudents = $false,
        [bool]$usePrefNameStaff = $true,
        [bool]$validateEmailInAD = $true,

        #Log File Info
        [string]$logPath = "$PSScriptRoot\Logs",
        [string]$logName = "$(Get-Date -UFormat '+%Y-%m-%d-%H-%M')-$(if($dryRun -eq $true){"DRYRUN-"})$([io.path]::GetFileNameWithoutExtension($MyInvocation.MyCommand.Name)).log",
        [string]$logLevel = "Information",

        #Program Varialbles
        [bool]$dryRun = $false

    )

<#
.SYNOPSIS
  Takes eduhub data drops the un-needed data for privacy and exports the files to Curric Servers

.DESCRIPTION
  
.PARAMETER <Parameter_Name>
  <Brief description of parameter input required. Repeat this attribute if required>

.INPUTS
  
  
.OUTPUTS
  CSV Files in output directory as specified in paramters
  Student Data (ST File)                ST_<SCHOOL_NUMBER>.csv

  Logfile as per path                   <YEAR>-<MONTH>-<DAY>-<HOUR>-<MINUTE> - Oliver.log

.NOTES
  Version:        1.0
  Author:         Justin Simmonds
  Creation Date:  2023-11-21
  Purpose/Change: Initial script development
  
.EXAMPLE
  <Example goes here. Repeat this attribute for more than one example>
#>

#---------------------------------------------------------[Initialisations]--------------------------------------------------------

#Set Error Action to Silently Continue
$ErrorActionPreference = "SilentlyContinue"

#----------------------------------------------------------[Declarations]----------------------------------------------------------


#Script Variables - Declared to stop it being generated multiple times per run

#Variables for Output Blanking - The Keys are always exported. All fields are imported and processed for ease of programming, its only at the output the the fields will be dropped
#Comment out by putting # at the start of the line which fields you do not want to export, do this carefully as it may have unintended concequences

$fieldsST = @(
                    'STKEY'
                    'SURNAME'
                    'FIRST_NAME'
                    'PREF_NAME'
                    'SCHOOL_YEAR'
                    'HOME_GROUP'
                    'NEXT_HG'
                    'SIGN IN MODE'
                    'E_MAIL'
                    'PASSWORD'
                    'TAGS'
)

$fieldsSF = @(
                    'SFKEY'
                    'PAYROLL_REC_NO'
                    'FIRST_NAME'
                    'SURNAME'
                    'PREF_NAME'
                    'E_MAIL'
                )

$fieldsKGC = @(
                    'KGCKEY'
                    'DESCRIPTION'
                    'TEACHER'
                    'TEACHER_B'
                )

#-----------------------------------------------------------[Functions]------------------------------------------------------------

function Write-Log 
{
    Param 
    (
        [Parameter(Mandatory=$true)][string]$logMessage, 
        [System.ConsoleColor]$ForegroundColor
    )

    if ($null -eq $ForegroundColor)
    {
        Write-Host "$(Get-Date -UFormat '+%Y-%m-%d %H:%M:%S') - $logMessage"
    }
    else {
        Write-Host "$(Get-Date -UFormat '+%Y-%m-%d %H:%M:%S') - $logMessage" -ForegroundColor $ForegroundColor
    }
    
}

#-----------------------------------------------------------[Execution]------------------------------------------------------------

#Log-Start -LogPath $sLogPath -LogName $sLogName -ScriptVersion $sScriptVersion

if (-not (Test-Path -Path  $logPath))
{
    New-Item $logPath -ItemType Directory | Out-Null
}

$classes = $null

if (Test-Path -Path ("$fileImportLocation\KGC_$schoolNumber.csv"))
{
    Write-Log "KGC File Exists, Importing"
    $classes = Import-CSV -Path "$fileImportLocation\KGC_$schoolNumber.csv" -Delimiter "," | Where-Object ACTIVE -eq "Y" | Select-Object $fieldsKGC | Sort-Object KGCKEY

}
else 
{
    Write-Log "KGC file does not exist, no export will be processed"
    Exit
}

$staff = $null
$adUsers = $null
if ($validateEmailInAD)
{
    $adUsers = Get-ADUser -Filter * -Properties mail, employeeID
}

if (Test-Path -Path ("$fileImportLocation\SF_$schoolNumber.csv"))
{
    Write-Log "SF File Exists, Importing"
    $staff = Import-CSV -Path "$fileImportLocation\SF_$schoolNumber.csv" -Delimiter "," | Where-Object { $_.STAFF_STATUS -ne "LEFT" -and $_.STAFF_STATUS -ne "INAC"} | Select-Object $fieldsSF | Sort-Object SFKEY
    foreach ($staffMember in $staff)
    {
        #Handle setting user to prefered name if this is what is desired
        if ($usePrefNameStaff -and (-not ([string]::IsNullOrWhiteSpace($staffMember.PREF_NAME)) -and ($staffMember.FIRST_NAME -ne $staffMember.PREF_NAME)))
        {
            $staffMember.FIRST_NAME = $staffMember.PREF_NAME
            Write-Host "Setting user $($staffMember.SFKEY) to use prefered name of $($staffMember.PREF_NAME)"
        }
        if ($validateEmailInAD)
        {
            $staffAD = $null
            $staffAD = $adUsers | Where-Object { $_.samAccountName -eq $staffMember.PAYROLL_REC_NO -or $_.employeeID -eq $staffMember.SFKEY }
            $staffMember.E_MAIL = $staffAD.mail
        }
    }
}
else 
{
    Write-Log "SF file does not exist, no export will be processed"
    Exit
}

#Get Year to Output
$outputYear = (Get-Date -Format "yyyy")

#Blank Current Teacher Variable
$currentTeacher = $null

#Create array for output Data
$outputData = @()

if (Test-Path -Path ("$fileImportLocation\ST_$schoolNumber.csv"))
{
    Write-Log "ST File Exists, Importing"
    $students = Import-CSV -Path "$fileImportLocation\ST_$schoolNumber.csv" -Delimiter "," | Where-Object{$_.STATUS -ne "LEFT" -and $_.STATUS -ne "INAC"} | Select-Object $fieldsST | Sort-Object STKEY

    foreach ($student in ($students | Sort-Object HOME_GROUP,SCHOOL_YEAR,STKEY))
    {
        #Handle setting user to prefered name if this is what is desired
        if ($usePrefName -and (-not ([string]::IsNullOrWhiteSpace($student.PREF_NAME)) -and ($student.FIRST_NAME -ne $student.PREF_NAME)))
        {
            $student.FIRST_NAME = $student.PREF_NAME
            Write-Host "Setting user $($student.STKEY) to use prefered name of $($student.PREF_NAME)"
        }

        if ($validateEmailInAD)
        {
            $studentAD = $null
            $studentAD = $adUsers | Where-Object { $_.samAccountName -eq $student.STKEY }
            $student.E_MAIL = $studentAD.mail
        }

        if ([string]::IsNullOrWhiteSpace($currentTeacher.SFKEY) -or $currentTeacher.SFKEY -ne ($classes | Where-Object KGCKEY -eq $student.HOME_GROUP).TEACHER)
        {
            $currentTeacher = $null
            $currentTeacher = $staff | Where-Object SFKEY -eq ($classes | Where-Object KGCKEY -eq $student.HOME_GROUP).TEACHER
        }
        

        #Create Object with the required details to output to the system
        $tempObject = $null
        $tempObject = [PSCustomObject]@{
            'Teacher Email'	= $currentTeacher.E_MAIL
            'Teacher First Name' = $currentTeacher.FIRST_NAME
            'Teacher Last Name' = $currentTeacher.SURNAME
            'Class Name' = "$outputYear - Class $($student.HOME_GROUP)"
            'Grade Level' = $student.SCHOOL_YEAR
            'Sign In Mode' = "Google"
            'Student Name' = "$($student.FIRST_NAME) $($student.SURNAME)"
            'Student ID' = $student.STKEY
            'Student Email' = $student.E_MAIL
            'Student Password' = ""
            'Co-Teacher 1 Email' = ""
            'Co-Teacher 1 First Name' = ""
            'Co-Teacher 1 Last Name' = ""
         }
         
        # Add secondary Teacher if they exist
        if (-not [string]::IsNullOrWhiteSpace(($classes | Where-Object KGCKEY -eq $student.HOME_GROUP).TEACHER_B))
         {
            $tempTeach = $null
            $tempTeach = $staff | Where-Object SFKEY -eq ($classes | Where-Object KGCKEY -eq $student.HOME_GROUP).TEACHER_B
            $tempObject."Co-Teacher 1 Email" = $tempTeach.E_MAIL
            $tempObject."Co-Teacher 1 First Name" = $tempTeach.FIRST_NAME
            $tempObject."Co-Teacher 1 Last Name" = $tempTeach.SURNAME
            $tempTeach = $null
         }

         #add the temporary object to the output array and clear the temporary object
         $outputData += $tempObject
         $tempObject = $null

        
    }
    
}
else 
{
    Write-Log "ST file does not exist, no export will be processed"
    Exit
}



<#
#Drop blank homegroups and leaving students if the output is set for next year
if ($nextYear)
{
    $students = $students | Where-Object{$_.NEXT_HG -ne "" -and $_.NEXT_HG -ne "LVNG"}
}
#>
if (-not (Test-Path $fileOutputLocation))
{
    New-Item -Path $fileOutputLocation -ItemType Directory | Out-Null
}


if ($outputData.Count -eq 0)
{
    Write-Log "No Students to Export - Perhaps no homegroups set and trying to do a next year export"
    Exit
}
else 
{
    $outputData | Export-Csv -Path "$fileOutputLocation\SEESAW-$(if ($nextYear -eq $false) { Get-Date -Format "yyyy" } else {[int](Get-Date -Format "yyyy") +1 }).csv" -Force -Encoding $exportFormat -NoTypeInformation
}