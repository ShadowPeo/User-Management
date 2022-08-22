#requires -version 2
<#
.SYNOPSIS
  Takes eduhub data drops the un-needed data for privacy and exports the files for upload to the Oliver servers.

.DESCRIPTION
  Script reads the five needed incoming CSV's dropping fields that are not needed
  
  School Year names are imported and processed based upon the settings configured

  Users (staff and students) are iterated keeping those only as specified by the import paramters and dropping the rest. The working set then have their usernames 
  and potentially emails updated based upon settings
  
  Student Family data is iterated, dropping those not referenced in the students data, replacing the alt contacts (the _A values) with the _B values where contact B is
  marked as the primary contact
  
  Addresses are iterrated, dropping those not referenced in the staff or family data

  Data is output to CSV's dropping data only for filtering and processing

  CSV's uploaded to Oliver servers

.PARAMETER <Parameter_Name>
  <Brief description of parameter input required. Repeat this attribute if required>

.INPUTS
  Year Level Descriptions (KCY File)
  Student Data (ST File)
  Staff Data (SF File)
  Family Data (DF File)
  Address Data (UM File)
  
.OUTPUTS
  CSV Files for upload in output directory as specified in paramters
  Year Level Descriptions (KCY File)    KCY_<SCHOOL_NUMBER>.csv
  Student Data (ST File)                ST_<SCHOOL_NUMBER>.csv
  Staff Data (SF File)                  SF_<SCHOOL_NUMBER>.csv
  Family Data (DF File)                 DF_<SCHOOL_NUMBER>.csv
  Address Data (UM File)                UM_<SCHOOL_NUMBER>.csv

  Logfile as per path                   <YEAR>-<MONTH>-<DAY>-<HOUR>-<MINUTE> - Oliver.log

.NOTES
  Version:        1.0
  Author:         Justin Simmonds
  Creation Date:  2022-08-19
  Purpose/Change: Initial script development
  
.EXAMPLE
  <Example goes here. Repeat this attribute for more than one example>
#>

#---------------------------------------------------------[Initialisations]--------------------------------------------------------

#Set Error Action to Silently Continue
$ErrorActionPreference = "SilentlyContinue"

#Dot Source required Function Libraries
. "$PSScriptRoot\Modules\Logging.ps1"

#----------------------------------------------------------[Declarations]----------------------------------------------------------

#Script Version
$sScriptVersion = "1.0"

#Config File Decleration - if used it will overwrite the default parameters below
$fileConfig = $null #Null or Blank for ignore, any other value and the script will attempt import

#School Details
$schoolID = "schoolno" # Used for export and for import if using CASES File Names
#$schoolID = [system.environment]::MachineName.Trim().Substring(0,4)

$schoolEmailDomain = "schooldom" #Only used if processing emails or users from CASES Data

#File Settings
$modifiedHeaders = $false #Use Modified Export Headers (from export script in this Repo), if not it will look for standard eduHub headers
$includeDeltas = $false #Include eduHub Delta File

#File Locations
$fileLocation = "$PSSCriptRoot/Import"
$importFileStudents = "ST_$SchoolID.csv"
$importFileStudentsDelta = "ST_$SchoolID_D.csv"
$importFileStaff = "SF_$SchoolID.csv"
$importFileStaffDelta = "SF_$SchoolID_D.csv"
$importFileYearLevels = "KCY_$SchoolID.csv"
$importFileFamilies = "DF_$SchoolID.csv"
$importFileFamiliesDelta = "DF_$SchoolID_D.csv"
$importFileAddresses = "UM_$SchoolID.csv"
$importFileAddressesDelta = "UM_$SchoolID_D.csv"

#Processing Handling Varialbles
$handlingStudentExitAfter = 365 #How long to export the data after the staff member or student has left. this is calculated based upon Exit Date, if it does not exist but marked as left they will be exported until exit date is established; 0 Disables export of left students
$handlingStaffExitAfter = 365 #How long to export the data after the staff member or student has left. this is calculated based upon Exit Date, if it does not exist but marked as left they will be exported until exit date is established; 0 Disables export of left staff, -1 will always export them
$handlingFileYearLevel = 1 # 1 = Static (use the one from cache, if not exist cache copy and us as literal) 2 = Use Literal, description will e exported exactly as is. 3 = Pad the year numbers (if they exist) in the description field
$handlingIncludeFutures = $true #Include Future Students
$handlingStudentEmail = 1 #1 = Use eduHub Email, 2 = Calculate from eduHub Data (STKEY/SIS_ID)@domain, 3 = pull from AD UPN, 4 = Pull from AD Mail, 5 = Pull from AD ProxyAddresses looking for primary (Capital SMTP)
$handlingStaffEmail = 1 #1 = Use eduHub Email, 2 = Calculate from eduHub Data (SFKEY/SIS_ID)@domain, 3 = pull from AD UPN, 4 = Pull from AD Mail, 5 = Pull from AD ProxyAddresses looking for primary (Capital SMTP),  6 = Use employeeID (PAYROLL_REC_NO/EmployeeNumber) from AD, fall back to SFKEY, 7 = Use employeeID (PAYROLL_REC_NO/EmployeeNumber) from eduHub Data, fall back to SFKEY
$handlingStudentUsername = 1 #-1 = Exclude from Export, #0 = Blank, 1 = use eduHub Data (STKEY/SIS_ID), 2 = Calculate from eduHub Data (STKEY/SIS_ID)@domain, 3 = pull from AD UPN, 4 = Pull from AD Mail, 5 Use samAccountName
$handlingStaffUsername = 1 #-1 = Exclude from Export, #0 = Blank, 1 = use eduHub Data (SFKEY/SIS_ID), 2 = Calculate from eduHub Data (SFKEY/SIS_ID)@domain, 3 = pull from AD UPN, 4 = Pull from AD Mail, 5 Use samAccountName, 6 = Use employeeID (PAYROLL_REC_NO/EmployeeNumber) from AD, fall back to SFKEY, 7 = Use employeeID (PAYROLL_REC_NO/EmployeeNumber) from eduHub Data, fall back to SFKEY
$handlingStudentAlias = 4 #1 = SFKEY, 2= use samAccountName, 3 = Use employeeID from Active Directory - Fall back to STKEY
$handlingStaffAlias = 4 #1 = SFKEY, 2= use samAccountName, 3 = Use employeeID from Active Directory - Fall back to SFKEY, 4 = Use employeeID (PAYROLL_REC_NO/EmployeeNumber) from eduHub Data - Fall back to SFKEY
$handlingValidateLicencing = $true #Validate the licencing for Oliver, this will drop accounts where it is explictly disabled
$handlingCreateNonEduhub = $true #Create accounts for users where licencing is explicitly enabled but not in eduHub data samAccountName becomes Alias other attributes handled as per settings (where available) or defaults
$handlingLicencingValue = "licencingOliver" #The attribute name for the licencing Data
$handlingADStaffType = "employeeType" #The attribute name for stating whether its a staff user or not for imports, only important if $handlingCreateNonEduhub is true, needs to be "Staff" or "15" (as in UserCreator) otherwise will assume student
$handlingExportNoUser = $true #Export user if there is no matching username in AD, if AD lookup is in use

#Active Directory Settings (Only required if using AD lookups - Active Directory lookups rely on the samAccountName being either the Key (SFKEY/STKEY) or in the case of staff members PAYROLL_REC_NO Matches will also be based upon email matching UPN
$runAsLoggedIn = $true
$activeDirectoryUser = $null #Username to connect to AD as, will prompt for password if credentials do not exist or are incorrect, not used if not running as logged in user
$activeDirectoryServer = $null #DNS Name or IP of AD Server

#Log File Info
$sLogPath = "C:\Windows\Temp"
$sLogName = "<script_name>.log"
$sLogFile = Join-Path -Path $sLogPath -ChildPath $sLogName

#-----------------------------------------------------------[Functions]------------------------------------------------------------

function Get-SchoolDetails 
{
    $path = "C:\Windows\Temp\whoami.xml"
    Invoke-WebRequest -Uri http://broadband.doe.wan/ipsearch/showresult.php -Method POST -Body @{mode='whoami'}| Select-Object -Expand Content | Out-File -Encoding "UTF8" $path
    $oXMLDocument=New-Object System.XML.XMLDocument  
    $oXMLDocument.Load($path)
    return $oXMLDocument.resultset.site
}

function Join-eduHubDelta
{
    Param 
    (
        [Parameter(Mandatory=$true)][string]$file1, 
        [Parameter(Mandatory=$true)][string]$file2, 
        [Parameter(Mandatory=$true)][string]$matchAttribute
    )

    $file1Import = Import-Csv -Path $file1
    
    ## Merge the two files if the second file is newer than the second
    if ((Get-Item $file1).LastWriteTime -lt (Get-Item $file2).LastWriteTime)
    {
        $file2Import = Import-Csv -Path $file2

        foreach ($record in $file2Import)
        {
            if ($file1Import.$matchAttribute -contains $record.$matchAttribute)
            {
                $file1Import[([array]::IndexOf( $file1Import.$matchAttribute , $record.$matchAttribute ))] = $record
            }
        }

        return $file1Import
    }
    else
    { 
        return $file1Import
    }
}



Function Merge-User
{
    Param 
    (
        [Parameter(Mandatory=$true)]$workingUser, 
        [Parameter(Mandatory=$true)]$exitAfter, 
        [Parameter(Mandatory=$true)]$handlingEmail,
        [Parameter(Mandatory=$true)]$handlingUsername,
        [Parameter(Mandatory=$true)]$handlingAlias,
        [Parameter(Mandatory=$true)][boolean]$handlingNoUser
    )
  
  Begin{
    #Log-Write -LogPath $sLogFile -LineValue "<description of what is going on>..."
  }
  
  Process{
    Try
    {
        #Set whether this is a staff or student user
        $userStaff = $false
        if ($null -eq $workingUser.SFKEY)
        {
            $userStaff = $true
        }

        #If user is marked as left, run exit check
        if ($workingUser.STATUS -eq "LEFT" -or $workingUser.STAFF_STATUS -eq "LEFT")
        {
            #Handle Exited Users if they are not meant to be exported, return null, else continue
            if ($exitAfter -gt 1)
            {
                #check if current date is more than $exitAfter days after the users exited date
                if((-not $userStaff -and (($currentDate - (Get-Date $workingUser.EXIT_DATE)).Days) -gt $exitAfter) -or ($userStaff -and (($currentDate - (Get-Date $workingUser.FINISH)).Days) -gt $exitAfter))
                {
                    return $null
                }
            }
            elseif ($exitedAfter -eq 0)
            {
                return $null
            }
        }

        #Email Handling

        switch ($handlingEmail)
        {
            #1 = Use eduHub Email
            1 
                {
                    #Do nothing, using eduHub Email
                }
            #2 = Calculate from eduHub Data (STKEY/SIS_ID)@domain
            2   {
                    if ([string]::IsNullOrWhiteSpace($schoolEmailDomain))
                    {
                        Write-Host "Email Domain Blank but told to use in settings, exiting"
                        exit
                    }

                    if (-not $userStaff)
                    {
                        $workingUser.E_MAIL = "$(($workingUser.STKEY).ToLower())@$schoolEmailDomain"
                    }
                    else
                    {
                        $workingUser.E_MAIL = "$(($workingUser.SFKEY).ToLower())@$schoolEmailDomain"
                    }
                }
            #3 = pull from Active Directory UPN
            3   {
                    
                }
            #4 = Pull from Active Directory Mail
            4   {
                    
                }
            #5 = Pull from Active Directory ProxyAddresses looking for primary (Capital SMTP)
            5   {
                    
                }
            #6 = Use employeeID (PAYROLL_REC_NO/EmployeeNumber) from Active Directory, fall back to SFKEY
            6   {
                    
                }
            #7 = Use employeeID (PAYROLL_REC_NO/EmployeeNumber) from eduHub, fall back to SFKEY
            {7 -and $userStaff -eq $true}

                {
                    if (-not [string]::IsNullOrWhiteSpace($workingUser.PAYROLL_REC_NO))
                    {
                        $workingUser.E_MAIL = "$(($workingUser.PAYROLL_REC_NO).ToLower())@$schoolEmailDomain"
                    }
                    else
                    {
                        $workingUser.E_MAIL = "$(($workingUser.SFKEY).ToLower())@$schoolEmailDomain"
                    }
                }
            #default = Use eduHub Email
            default 
                {
                    #Do nothing, using eduHub Email
                }
        }

        #Username Handling

        switch ($handlingUsername)
        {
            #-1 = Excluded Column on Export
            -1 
                {
                    $workingUser.USERNAME = "EXCLUDED"
                }
            
            #0 = Blank the field ""
            0 
                {
                    $workingUser.USERNAME = ""
                }
            
            #1 = Use eduHub Key
            1 
                {
                    if (-not $userStaff)
                    {
                        $workingUser.USERNAME = ($workingUser.STKEY).ToUpper()
                    }
                    else
                    {
                        $workingUser.USERNAME = ($workingUser.SFKEY).ToUpper()
                    }
                }
            #2 = Calculate from eduHub Data (STKEY/SIS_ID)@domain
            2   {
                    if ([string]::IsNullOrWhiteSpace($schoolEmailDomain))
                    {
                        Write-Host "Email Domain Blank but told to use in settings, exiting"
                        exit
                    }

                    if (-not $userStaff)
                    {
                        $workingUser.USERNAME = "$(($workingUser.STKEY).ToLower())@$schoolEmailDomain"
                    }
                    else
                    {
                        $workingUser.USERNAME = "$(($workingUser.SFKEY).ToLower())@$schoolEmailDomain"
                    }
                }
            #3 = pull from Active Directory UPN
            3   {
                    
                }
            #4 = Pull from Active Directory Mail
            4   {
                    
                }
            #5 = Pull from Active Directory ProxyAddresses looking for primary (Capital SMTP)
            5   {
                    
                }
            #6 = Use employeeID (PAYROLL_REC_NO/EmployeeNumber) from Active Directory, fall back to SFKEY
            6   {
                    
                }
            #7 = Use employeeID (PAYROLL_REC_NO/EmployeeNumber) from eduHub, fall back to SFKEY
            {7 -and $userStaff -eq $true}

                {
                    if (-not [string]::IsNullOrWhiteSpace($workingUser.PAYROLL_REC_NO))
                    {
                        $workingUser.USERNAME = $workingUser.PAYROLL_REC_NO
                    }
                    else
                    {
                        $workingUser.USERNAME = $workingUser.SFKEY
                    }
                }
            #Default = Use eduHub Key (STKEY/SFKEY)
            default 
                {
                    if (-not $userStaff)
                    {
                        $workingUser.USERNAME = ($workingUser.STKEY).ToUpper()
                    }
                    else
                    {
                        $workingUser.USERNAME = ($workingUser.SFKEY).ToUpper()
                    }
                }
        }

        #Alias Handling

        switch ($handlingAlias)
        {

            #1 = Use eduHub Key (STKEY/SFKEY)
            1 
                {
                    if (-not $userStaff)
                    {
                        $workingUser.ALIAS = ($workingUser.STKEY).ToUpper()
                    }
                    else
                    {
                        $workingUser.ALIAS = ($workingUser.SFKEY).ToUpper()
                    }
                }
            #2 = Use samAccountName from Active Directory
            2   {
                    
                }
            #3 = Use employeeID from Active Directory
            3   {
                    
                }
            #4 = Use employeeID (PAYROLL_REC_NO/EmployeeNumber) from eduHub Data
            {4 -and $userStaff -eq $true}

                {
                    if (-not [string]::IsNullOrWhiteSpace($workingUser.PAYROLL_REC_NO))
                    {
                        $workingUser.ALIAS = $workingUser.PAYROLL_REC_NO
                    }
                    else
                    {
                        $workingUser.ALIAS = $workingUser.SFKEY
                    }
                }
            #Default = Use eduHub Key (STKEY/SFKEY)
            default 
                {
                    if (-not $userStaff)
                    {
                        $workingUser.ALIAS = ($workingUser.STKEY).ToUpper()
                    }
                    else
                    {
                        $workingUser.ALIAS = ($workingUser.SFKEY).ToUpper()
                    }
                }
        }


        return $workingUser

    }
    
    Catch{
      #Log-Error -LogPath $sLogFile -ErrorDesc $_.Exception -ExitGracefully $True
      Break
    }
  }
  
  End{
    If($?){
      #Log-Write -LogPath $sLogFile -LineValue "Completed Successfully."
      #Log-Write -LogPath $sLogFile -LineValue " "
    }
  }
}


#-----------------------------------------------------------[Execution]------------------------------------------------------------

#Log-Start -LogPath $sLogPath -LogName $sLogName -ScriptVersion $sScriptVersion
#Script Execution goes here
#Log-Finish -LogPath $sLogFile

if ($runAsLoggedIn -eq $false)
{
    Import-Module "$PSScriptRoot\Modules\Authentication.ps1"
    $schoolServiceCreds = Get-SavedCredentials_WithRequest "$PSScriptRoot\Credentials\schoolDC-$([Environment]::MachineName)-$([Environment]::UserName).crd" $activeDirectoryUser
    $schoolServiceCreds = new-object -typename System.Management.Automation.PSCredential -argumentlist $schoolServiceCreds.Username,$schoolServiceCreds.Password
    $userCredentails = $null
}


######################Import and Process Students######################

$importedStudents = $null
$workingStudents = @()

#Get The Date, Put here so only done once for the run
$currentDate = Get-Date

#Import Students from CSV(s) based upon settings

if ($includeDeltas -eq $true -and $modifiedHeaders -eq $false) #Only do Delta join if not using files from exporter as exporter joins the files
{
    <#
    $importedStudents = (Join-eduHubDelta $fileStudent $fileStudentDelta "STKEY") | 
	select -Property @{label="SIS_ID";expression={$($_."STKEY")}},SURNAME,FIRST_NAME,SECOND_NAME,PREF_NAME,BIRTHDATE,@{label="SIS_EMAIL";expression={$($_."E_MAIL")}},HOUSE,CAMPUS,STATUS,@{label="START";expression={$($_."ENTRY")}},@{label="FINISH";expression={$($_."EXIT_DATE")}},SCHOOL_YEAR,HOME_GROUP,NEXT_HG
	Sort-Object -property SIS_ID 
    #>
}
elseif ($includeDeltas -eq $false -and $modifiedHeaders -eq $false) #Only Run import if not using modified headers from exporter
{
    $importedStudents = Import-CSV (Join-Path -Path $fileLocation -ChildPath $importFileStudents) | Select-Object -Property STKEY,PREF_NAME,FIRST_NAME,SURNAME,BIRTHDATE,GENDER,EXIT_DATE,HOME_GROUP,SCHOOL_YEAR,FAMILY,USERNAME,E_MAIL,CONTACT_A,STATUS,ALIAS,EXPORT | Sort-Object -property STATUS, STKEY
}
elseif ($modifiedHeaders -eq $true)
{
    <#
    $importedStudents = (Join-eduHubDelta $fileStudent $fileStudentDelta "STKEY") | 
	select -Property @{label="STKEY";expression={$($_."SIS_ID")}},SURNAME,FIRST_NAME,SECOND_NAME,PREF_NAME,BIRTHDATE,@{label="SIS_EMAIL";expression={$($_."E_MAIL")}},HOUSE,CAMPUS,STATUS,@{label="START";expression={$($_."ENTRY")}},@{label="FINISH";expression={$($_."EXIT_DATE")}},SCHOOL_YEAR,HOME_GROUP,NEXT_HG
	Sort-Object -property SIS_ID 
    #>
}
else
{
    throw "Cannot Import Error with locating or processing files"
}

#Include Active, and Leaving students as we know they do not need date validation, and Future Students if set for that as well
if ($handlingIncludeFutures -eq $true)
{
    $workingStudents = $importedStudents | Where-Object {$_.STATUS -eq "ACTV" -or $_.STATUS -eq "LVNG"  -or $_.STATUS -eq "FUT" }
}
else
{
    $workingStudents = $importedStudents | Where-Object {$_.STATUS -eq "ACTV" -or $_.STATUS -eq "LVNG" }
}

foreach ($student in $workingStudents)
{
    if ($null -ne ($tempUser = (Merge-User -workingUser $student -exitAfter $handlingStudentExitAfter -handlingEmail $handlingStudentEmail -handlingUsername $handlingStudentUsername -handlingAlias $handlingStudentAlias -handlingNoUser $handlingExportNoUser)))
    {
        $workingStudents += $tempUser
    }
}


##################### $importedStudents | Where-Object {$_.STATUS -eq "LEFT" -and ($currentDate - ((Get-Date $_.EXIT_DATE).Days) -lt $handlingStudentExitAfter) }
foreach ($student in ($importedStudents | Where-Object {$_.STATUS -eq "LEFT" }))
{
    if ($null -ne ($tempUser = (Merge-User -workingUser $student -exitAfter $handlingStudentExitAfter -handlingEmail $handlingStudentEmail -handlingUsername $handlingStudentUsername -handlingAlias $handlingStudentAlias -handlingNoUser $handlingExportNoUser)))
    {
        $workingStudents += $tempUser
    }
}


######################Import and Process Staff######################

$importedStaff = $null
$workingStaff = @()

#Get The Date, Put here so only done once for the run
$currentDate = Get-Date

#Import Staff from CSV(s) based upon settings

if ($includeDeltas -eq $true -and $modifiedHeaders -eq $false) #Only do Delta join if not using files from exporter as exporter joins the files
{
    <#
    $importedStaff = (Join-eduHubDelta $fileStaff $fileStaffDelta "STKEY") | 
	select -Property @{label="SIS_ID";expression={$($_."STKEY")}},SURNAME,FIRST_NAME,SECOND_NAME,PREF_NAME,BIRTHDATE,@{label="SIS_EMAIL";expression={$($_."E_MAIL")}},HOUSE,CAMPUS,STATUS,@{label="START";expression={$($_."ENTRY")}},@{label="FINISH";expression={$($_."EXIT_DATE")}},SCHOOL_YEAR,HOME_GROUP,NEXT_HG
	Sort-Object -property SIS_ID 
    #>
}
elseif ($includeDeltas -eq $false -and $modifiedHeaders -eq $false) #Only Run import if not using modified headers from exporter
{
    $importedStaff = Import-CSV (Join-Path -Path $fileLocation -ChildPath $importFileStaff) | Select-Object -Property SFKEY,PREF_NAME,FIRST_NAME,SURNAME,BIRTHDATE,GENDER,FINISH,HOMEKEY,USERNAME,E_MAIL,STAFF_STATUS,PAYROLL_REC_NO,ALIAS,EXPORT | Sort-Object -property STAFF_STATUS, SFKEY
}
elseif ($modifiedHeaders -eq $true)
{
    <#
    $importedStaff = (Join-eduHubDelta $fileStaff $fileStaffDelta "STKEY") | 
	select -Property @{label="STKEY";expression={$($_."SIS_ID")}},SURNAME,FIRST_NAME,SECOND_NAME,PREF_NAME,BIRTHDATE,@{label="SIS_EMAIL";expression={$($_."E_MAIL")}},HOUSE,CAMPUS,STATUS,@{label="START";expression={$($_."ENTRY")}},@{label="FINISH";expression={$($_."EXIT_DATE")}},SCHOOL_YEAR,HOME_GROUP,NEXT_HG
	Sort-Object -property SIS_ID 
    #>
}
else
{
    throw "Cannot Import Error with locating or processing files"
}

##################### $importedStaff | Where-Object {$_.STATUS -eq "LEFT" -and ($currentDate - ((Get-Date $_.EXIT_DATE).Days) -lt $handlingStaffExitAfter) }
foreach ($staff in $importedStaff)
{
    if ($null -ne ($tempUser = (Merge-User -workingUser $staff -exitAfter $handlingStaffExitAfter -handlingEmail $handlingStaffEmail -handlingUsername $handlingStaffUsername  -handlingAlias $handlingStaffAlias -handlingNoUser $handlingExportNoUser)))
    {
        $workingStaff += $tempUser
    }
}