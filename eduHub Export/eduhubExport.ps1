param 
    (
        
        #School Details
        [string]$schoolNumber = "<<SCHOOL_ID>>", # Used for export and for import if using CASES File Names, TODO: Add code to pull from server name/bbportal

        #File Settings
        [boolean]$includeDeltas = $true, #Include eduHub Delta File

        #File Locations
        [string]$fileImportLocation = "$PSScriptRoot\eduHub",
        [string]$fileOutputLocation = "$PSScriptRoot\Output",
        [string]$fileStafManualMatch = "$PSScriptRoot\Config\manualMatch.csv",
        [string]$filesToProcess = "ST|SF|DF|KGC|KCY|UM",

        #Processing Handling Varialbles
        [float]$handlingStudentNoExportAfter = 365, #How long to export the data after the staff member or student has left. this is calculated based upon Exit Date, if it does not exist but marked as left they will be exported until exit date is established; 0 Disables export of left students, -1 will always export them
        [float]$handlingStaffNoExportAfter = -1, #How long to export the data after the staff member or student has left. this is calculated based upon Exit Date, if it does not exist but marked as left they will be exported until exit date is established; 0 Disables export of left staff, -1 will always export them
        [string]$exportFormat = "utf8", #Formats supported for output ascii, unicode, utf8, utf32

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
  Year Level Descriptions (KCY File)    KCY_<SCHOOL_NUMBER>.csv
  Student Data (ST File)                ST_<SCHOOL_NUMBER>.csv
  Staff Data (SF File)                  SF_<SCHOOL_NUMBER>.csv
  Family Data (DF File)                 DF_<SCHOOL_NUMBER>.csv
  Address Data (UM File)                UM_<SCHOOL_NUMBER>.csv
  Homegroup Data (KGC File)             KGC_<SCHOOL_NUMBER>.csv

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

#Dot Source required Function Libraries
"$PSScriptRoot\Modules\Logging.ps1"

#----------------------------------------------------------[Declarations]----------------------------------------------------------


#Script Variables - Declared to stop it being generated multiple times per run

#Variables for Output Blanking - The Keys are always exported. All fields are imported and processed for ease of programming, its only at the output the the fields will be dropped
#Comment out by putting # at the start of the line which fields you do not want to export, do this carefully as it may have unintended concequences

$fieldsST = @(
                    'STKEY'
                    'SURNAME'
                    'FIRST_NAME'
                    'SECOND_NAME'
                    'PREF_NAME'
                    'BIRTHDATE'
                    'E_MAIL'
                    'HOUSE'
                    'CAMPUS'
                    'STATUS'
                    'ENTRY'
                    'EXIT_DATE'
                    'SCHOOL_YEAR'
                    'HOME_GROUP'
                    'NEXT_HG'
                    'FAMILY' #Used to lookup family details, to get parent and address details
                    'CONTACT_A'
                    'GENDER'
                    'MOBILE'
                    'KOORIE'
                    'DISABILITY'
                    'ED_ALLOW'
                    'LOTE_HOME_CODE'
                    'ENG_SPEAK'
                    'LW_DATE'
)

$fieldsSF = @(
                    'SFKEY'
                    'SURNAME'
                    'FIRST_NAME'
                    'SECOND_NAME'
                    'PREF_NAME'
                    'BIRTHDATE'
                    'E_MAIL'
                    'HOUSE'
                    'CAMPUS'
                    'STAFF_STATUS'
                    'START'
                    'FINISH'
                    'TITLE'
                    'MOBILE'
                    'WORK_PHONE'
                    'PAYROLL_REC_NO'
                    'FTE'
                    'GENDER'
                    'HOMEKEY'           #Used to lookup address details, to get home address details
                    'MAILKEY'           #Used to lookup address details, to get mailing address details
                    'LW_DATE'
)

$fieldsDF = @(
                    'DFKEY'
                    'E_MAIL_A'
                    'MOBILE_A'
                    'E_MAIL_B'
                    'MOBILE_B'
                    'HOMEKEY'
                    'LOTE_HOME_CODE_A'
                    'LOTE_HOME_CODE_B'
                    'LW_DATE'
)

$fieldsUM = @(
                    'UMKEY'
                    'ADDRESS01'
                    'ADDRESS02'
                    'ADDRESS03'
                    'STATE'
                    'POSTCODE'
                    'TELEPHONE'
                    'MOBILE'
                    'LW_DATE'
)

$fieldsKCY = @(
                    'KCYKEY'
                    'DESCRIPTION'
                    'NUM_EQVT'
                    'LW_DATE'
)

$fieldsKGC = @(
                    'KGCKEY'
                    'DESCRIPTION'
                    'CAMPUS'
                    'TEACHER_A'
                    'TEACHER_B'
                    'ACTIVE'
                    'ROOM'
                    'LW_DATE'
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

$filesToProcess = $filesToProcess.Split("|")
$currentDate = Get-Date
$validFamilies = @()
$validAddresses = @()

foreach ($file in $filesToProcess)
{

    $workingFile = $null
    if (Test-Path -Path ("$fileImportLocation\$file`_$schoolNumber.csv"))
    {
        Write-Log "Full $file file Exists, Importing" 
        $workingFile = Import-Csv -Path ("$fileImportLocation\$file`_$schoolNumber.csv")

        if ($includeDeltas -and ((Test-Path -Path ("$fileImportLocation\$file`_$schoolNumber`_D.csv")) -and (((Get-Item ("$fileImportLocation\$file`_$schoolNumber`_D.csv")).LastWriteTime) -gt ((Get-Item ("$fileImportLocation\$file`_$schoolNumber.csv")).LastWriteTime))))
        {
            Write-Log "$file Delta processing is enabled, the file exists, and is newer than the full output file, Merging" 

            foreach ($record in (Import-Csv -Path ("$fileImportLocation\$file`_$schoolNumber`_D.csv")))
            {
                if ($workingFile.$("$file`KEY") -contains $record.$("$file`KEY"))
                {
                    Write-Log "Record ($($record.$("$file`KEY"))) Matches Existing Record, Merging" 
                    
                    foreach ($row in $workingFile)
                    {
                        if ($row.$("$file`KEY") -eq $record.$("$file`KEY"))
                        {
                            $row = $record
                        }
                    }
                }
                else
                {
                    Write-Log "New Record Found, Inserting" 
                    $workingFile += $record
                }
            }
        }
        else 
        {
            if (-not $includeDeltas)
            {
                Write-Log "$file Delta processing is disabled, using only the full file" 
            }
            elseif (-not (Test-Path -Path ("$fileImportLocation\$file`_$schoolNumber`_D.csv")))
            {
                Write-Log "$file Delta file does not exist, using only the full file" 
            }
            else 
            {
                Write-Log "$file Delta is not newer than the full output file, using only the full file" 
            }
        }
        
        
        $fileHeaders = $null
        $fileHeaders = (($workingFile | Select-Object -First 1).psobject.properties).Name
        $workingFile = ($workingFile | Select-Object ((Get-Variable -Name "fields$file").Value) | Sort-Object ("$file`KEY"))
        $workingFile = [System.Collections.ArrayList]$workingFile
        
        switch ($file)
        {
            "ST" 
            { 
                switch ($handlingStudentNoExportAfter) {
                    -1
                    { 
                        Write-Log "Student export filtering Disabled, exporting all" 
                    }

                    0
                    {
                        Write-Log "Student export filtering set to export none after left, exporting Future, Active and Leaving only" 
                        $workingFile = $workingFile | Where-Object {$_.STAFF_STATUS -ne 'LEFT' -and $_.STAFF_STATUS -ne 'INAC'}
                        $validAddresses += $workingFile | Where-Object {$_.STAFF_STATUS -ne 'LEFT' -and $_.STAFF_STATUS -ne 'INAC'} | Select-Object HOMEKEY
                        $validAddresses += $workingFile | Where-Object {$_.STAFF_STATUS -ne 'LEFT' -and $_.STAFF_STATUS -ne 'INAC'} | Select-Object MAILKEY

                    }

                    Default 
                    {
                        Write-Log "Student export filtering set to restrict export to $handlingStudentNoExportAfter days after left" 
                        $tempExited = $workingFile | Where-Object {$_.STATUS -eq 'LEFT'  -or $_.STATUS -eq 'INAC'}
                        $workingFile = $workingFile | Where-Object {$_.STATUS -ne 'LEFT' -and $_.STATUS -ne 'INAC'}
                        $validFamilies += $workingFile.FAMILY
                        
                        foreach ($student in $tempExited)
                        {
                            if ((($currentDate - (Get-Date ($student.EXIT_DATE))).Days) -le $handlingStudentNoExportAfter)
                            {
                                $workingFile += $student
                                $validFamilies += $studnet.FAMILY
                            }
                        }
                        $validFamilies = ($validFamilies | Get-Unique | Sort-Object)
                    }
                }
            }

            "SF"
            {

                switch ($handlingStaffNoExportAfter) {
                    -1
                    { 
                        Write-Log "Staff export filtering Disabled, exporting all" 
                        $validAddresses += $workingFile | Select-Object HOMEKEY
                        $validAddresses += $workingFile | Select-Object MAILKEY
                    }

                    0
                    {
                        Write-Log "Staff export filtering set to export none after left, exporting Active only" 
                        $workingFile = $workingFile | Where-Object {$_.STAFF_STATUS -ne 'LEFT' -and $_.STAFF_STATUS -ne 'INAC'}
                        $validAddresses += $workingFile | Where-Object {$_.STAFF_STATUS -ne 'LEFT' -and $_.STAFF_STATUS -ne 'INAC'} | Select-Object HOMEKEY
                        $validAddresses += $workingFile | Where-Object {$_.STAFF_STATUS -ne 'LEFT' -and $_.STAFF_STATUS -ne 'INAC'} | Select-Object MAILKEY

                    }

                    Default 
                    {
                        Write-Log "Staff export filtering set to restrict export to $handlingStaffNoExportAfter days after left" 
                        $tempExited = $workingFile | Where-Object {$_.STAFF_STATUS -eq 'LEFT' -or $_.STAFF_STATUS -eq 'INAC'}
                        $workingFile = $workingFile | Where-Object {$_.STAFF_STATUS -ne 'LEFT' -and $_.STAFF_STATUS -ne 'INAC'}
                        
                        foreach ($staffMember in $tempExited)
                        {
                            if ((($currentDate - (Get-Date ($staffMember.FINISH))).Days) -le $handlingStaffNoExportAfter)
                            {
                                $workingFile += $staffMember
                                $validAddresses += $staffMember.HOMEKEY
                                $validAddresses += $staffMember.MAILKEY
                            }
                        }
                    }
                }
                
                if (Test-Path -Path $fileStafManualMatch)
                {
                    $manualMatchUser = Import-CSV -Path $fileStafManualMatch
                    foreach($staffMember in $workingFile)
                    {
                        if($manualMatchUser.CASESID -contains $staffMember.SFKEY)
                        {
                            Write-Log "$($staffMember.SFKEY) has a manual PAYROLL_REC_NO configured, changing to $(($manualMatchUser | where-object CASESID -eq $staffMember.SFKEY).T0NUMBER)"
                            $staffMember.PAYROLL_REC_NO = ($manualMatchUser | where-object CASESID -eq $staffMember.SFKEY).T0NUMBER
                        }
                    }
                }
                else
                {
                    Write-Log "Manual Match Path does not exist, ignoring"
                }
                
                
            }

            "DF"
            {
                $tempData = @()
                foreach ($family in $workingFile)
                {
                    
                    if ($validFamilies -contains $family.("$file`KEY"))
                    {
                        $tempData += $family
                        $validAddresses += $family.HOMEKEY
                    }
                }
                $workingFile = $tempData
            }

            "UM"
            {
                $tempData = @()
                $validAddresses = $validAddresses | Get-Unique
                foreach ($address in $workingFile)
                {
                    if ($validAddresses -contains $address.("$file`KEY"))
                    {
                        $tempData += $address
                    }
                }
                $workingFile = $tempData

            }
        }

        $workingFile | Select-Object $fileHeaders | Sort-Object ("$file`KEY") | Export-Csv -Path "$fileOutputLocation\$file`_$schoolNumber.csv" -Force -Encoding $exportFormat
    }
    else
    {
        Write-Log "$file file does not exist, no export will be processed, continuing to next file" 
        $filesToProcess.Remove($file)
    }
}