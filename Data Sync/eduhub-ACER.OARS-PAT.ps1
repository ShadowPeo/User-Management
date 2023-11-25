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
        [bool]$usePrefName = $true,

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
                    'BIRTHDATE'
                    'STATUS'
                    'EXIT_DATE'
                    'SCHOOL_YEAR'
                    'HOME_GROUP'
                    'NEXT_HG'
                    'GENDER'
                    'LW_DATE'
                    'PASSWORD'
                    'TAGS'
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

if (Test-Path -Path ("$fileImportLocation\ST_$schoolNumber.csv"))
{
    Write-Log "ST File Exists, Importing"
    $students = Import-CSV -Path "$fileImportLocation\ST_$schoolNumber.csv" -Delimiter "," | Where-Object{$_.STATUS -ne "LEFT" -and $_.STATUS -ne "INAC"} | Select-Object $fieldsST | Sort-Object STKEY

    foreach ($student in $students)
    {
        #Set Birthdate to match format required explicitly
        $student.BIRTHDATE = Get-Date $student.BIRTHDATE -Format "dd/MM/yyyy"
        
        #Handle setting user to prefered name if this is what is desired
        if ($usePrefName -and (-not ([string]::IsNullOrWhiteSpace($student.PREF_NAME)) -and ($student.FIRST_NAME -ne $student.PREF_NAME)))
        {
            $student.FIRST_NAME = $student.PREF_NAME
            Write-Host "Setting user $($student.STKEY) to use prefered name of $($student.PREF_NAME)"
        }

        if ($nextYear)
        {
            #Set Students homegroup to next years homegroup
            $student.HOME_GROUP = $student.NEXT_HG

            #Set Students School Year to next years school year
            if ($student.SCHOOL_YEAR -eq 12 -or $student.SCHOOL_YEAR -eq 13)
            {
                $student.SCHOOL_YEAR = 12
            }
            else {
                $student.SCHOOL_YEAR = $student.SCHOOL_YEAR + 1
            }
        }

        #TODO: Process: Tags, Password, First name vs PrefName
    }
    
}
else 
{
    Write-Log "ST file does not exist, no export will be processed"
    Exit
}

if (-not (Test-Path $fileOutputLocation))
{
    New-Item -Path $fileOutputLocation -ItemType Directory | Out-Null
}

#Drop blank homegroups and leaving students if the output is set for next year
if ($nextYear)
{
    $students = $students | Where-Object{$_.NEXT_HG -ne "" -and $_.NEXT_HG -ne "LVNG"}
}

if ($students.Count -eq 0)
{
    Write-Log "No Students to Export - Perhaps no homegroups set and trying to do a next year export"
    Exit
}
else 
{
    $students | Select-Object SURNAME,FIRST_NAME,STKEY,PASSWORD,BIRTHDATE,GENDER,TAGS,UNIQUE_ID,SCHOOL_YEAR | Sort-Object STKEY | Export-Csv -Path "$fileOutputLocation\ACER-OARS-$(if ($nextYear -eq $false) { Get-Date -Format "yyyy" } else {[int](Get-Date -Format "yyyy") +1 }).csv" -Force -Encoding $exportFormat -NoTypeInformation
}

