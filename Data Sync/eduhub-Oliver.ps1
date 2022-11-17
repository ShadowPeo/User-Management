

param 
    (
        
        #School Details
        [string]$schoolID = "3432", # Used for export and for import if using CASES File Names
        [string]$schoolEmailDomain = "mwps.vic.edu.au", #Only used if processing emails or users from CASES Data

        #File Settings
        [boolean]$includeDeltas = $true, #Include eduHub Delta File

        #File Locations
        [string]$fileImportLocation = "$PSSCriptRoot\Import",
        [string]$fileOutputLocation = "$PSSCriptRoot\Output",
        [string]$importFileStudents = "ST_$($SchoolID).csv",
        [string]$importFileStudentsDelta = "ST_$($SchoolID)_D.csv",
        [string]$importFileStaff = "SF_$($SchoolID).csv",
        [string]$importFileStaffDelta = "SF_$($SchoolID)_D.csv",
        [string]$importFileYearLevels = "KCY_$($SchoolID).csv",
        [string]$importFileFamilies = "DF_$($SchoolID).csv",
        [string]$importFileFamiliesDelta = "DF_$($SchoolID)_D.csv",
        [string]$importFileAddresses = "UM_$($SchoolID).csv",
        [string]$importFileAddressesDelta = "UM_$($SchoolID)_D.csv",

        #Processing Handling Varialbles
        [float]$handlingStudentExitAfter = 365, #How long to export the data after the staff member or student has left. this is calculated based upon Exit Date, if it does not exist but marked as left they will be exported until exit date is established; 0 Disables export of left students, -1 will always export them
        [float]$handlingStaffExitAfter = 365, #How long to export the data after the staff member or student has left. this is calculated based upon Exit Date, if it does not exist but marked as left they will be exported until exit date is established; 0 Disables export of left staff, -1 will always export them
        [int]$handlingFileYearLevel = 2, # 1 = Use Literal, description will e exported exactly as is. 2 = Pad the year numbers (if they exist) in the description field
        [boolean]$handlingIncludeFutures = $true, #Include Future Students
        [int]$handlingStudentEmail = 4, #1 = Use eduHub Email, 2 = Calculate from eduHub Data (SIS_ID)@domain, 3 = pull from AD UPN, 4 = Pull from AD Mail, 5 = Pull from AD ProxyAddresses looking for primary (Capital SMTP),  6 = Use employeeID (PAYROLL_REC_NO/SIS_EMPNO/EmployeeNumber) from AD, fall back to SIS_ID
        [int]$handlingStaffEmail = 4, #1 = Use eduHub Email, 2 = Calculate from eduHub Data (SIS_ID)@domain, 3 = pull from AD UPN, 4 = Pull from AD Mail, 5 = Pull from AD ProxyAddresses looking for primary (Capital SMTP),  6 = Use employeeID (PAYROLL_REC_NO/SIS_EMPNO/EmployeeNumber) from AD, fall back to SIS_ID, 7 = Use employeeID (PAYROLL_REC_NO/SIS_EMPNO/EmployeeNumber) from eduHub Data, fall back to SIS_ID
        [float]$handlingStudentUsername = 5, #0 = Blank, 1 = use eduHub Data (SIS_ID), 2 = Calculate from eduHub Data (SIS_ID)@domain, 3 = pull from AD UPN, 4 = Pull from AD Mail, 5 = Pull from AD ProxyAddresses looking for primary (Capital SMTP), 6 = Use samAccountName, 7 = Use employeeID (PAYROLL_REC_NO/SIS_EMPNO/EmployeeNumber) from AD, fall back to SIS_ID
        [float]$handlingStaffUsername = 5, #0 = Blank, 1 = use eduHub Data (SIS_ID), 2 = Calculate from eduHub Data (SIS_ID)@domain, 3 = pull from AD UPN, 4 = Pull from AD Mail, 5 = Pull from AD ProxyAddresses looking for primary (Capital SMTP), 6 = Use samAccountName, 7 = Use employeeID (PAYROLL_REC_NO/SIS_EMPNO/EmployeeNumber) from AD, fall back to SIS_ID, 8 = Use employeeID (PAYROLL_REC_NO/SIS_EMPNO/EmployeeNumber) from eduHub Data, fall back to SIS_ID
        [int]$handlingStudentAlias = 1, #1 = SIS_ID, 2= use samAccountName - Fall back to SIS_ID, 3 = Use employeeID from Active Directory - Fall back to SIS_ID
        [int]$handlingStaffAlias = 1, #1 = SIS_ID, 2= use samAccountName, 3 = Use employeeID from Active Directory - Fall back to SIS_ID, 4 = Use employeeID (PAYROLL_REC_NO/SIS_EMPNO/EmployeeNumber) from eduHub Data - Fall back to SIS_ID
        [boolean]$handlingValidateLicencing = $false, #Validate the licencing for Oliver, this will drop accounts where it is explictly disabled or where no user exists NOTE: This does not  validating the licencing value, only that the field is not blank - this is meant for use with something like Azure AD where access rights can be assigned based upon dynamic groups based upon AD fields
        [string]$handlingLicencingValue = "licencingLibrary", #The attribute name for the licencing Data NOTE: Ensure the AD schema value exists before running or you will get a silent error
        [boolean]$handlingExportNoUser = $false, #Export user if there is no matching username in AD, if AD lookup is in use
        [boolean]$exportFull = $false, #Include all columns from eduhub in export, blanking those not required
        [boolean]$exportCustom = $true, #Include all columns from eduhub in export, blanking those not required

        #Active Directory Settings (Only required if using AD lookups - Active Directory lookups rely on the samAccountName being either the Key (SIS_ID) or in the case of staff members PAYROLL_REC_NO/SIS_EMPNO Matches will also be based upon email matching UPN
        [boolean]$runAsLoggedIn = $false,
        [string]$activeDirectoryUser = "CURRIC\da.st00605", #Username to connect to AD as, will prompt for password if credentials do not exist or are incorrect, not used if not running as logged in user
        [string]$activeDirectoryServer = "10.128.136.35", #DNS Name or IP of AD Server
        [string]$activeDirectorySearchBase = "OU=User Accounts,OU=Accounts,OU=3432 - Mount Waverley PS,DC=curric,DC=mount-waverley-ps,DC=wan", #DNS Name or IP of AD Server

        #Log File Info
        [string]$logPath = "$PSScriptRoot\Logs",
        [string]$logName = "<script_name>.log",
        [string]$logLevel = "Information"
    )

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

#$sLogFile = Join-Path -Path $sLogPath -ChildPath $sLogName

#Script Variables - Declared to stop it being generated multiple times per run

#Date
$currentDate = Get-Date
$adCheck = $false #Changes to true if one of the settings requires an AD Check

#Variables for Output Blanking - The SIS_ID (STKEY/SFKEY) are always exported. All fields are imported and processed for ease of programming, its only at the output the the fields will be dropped
#Comment out by putting # at the start of the line which fields you do not want to export, do this carefully as it may have unintended concequences

$fieldsStudent = @(
                    'STKEY'
                    'PREF_NAME'
                    'FIRST_NAME'
                    'SURNAME'
                    'BIRTHDATE'
                    'GENDER'
                    'EXIT_DATE'
                    'HOME_GROUP'
                    'SCHOOL_YEAR'
                    'FAMILY'            #Used to lookup family details, to get parent and address details
                    'E_MAIL'
                    'MOBILE'
                    'SECOND_NAME'
                    'STATUS'    
                    'USERNAME'          #Allow this if you want to use the username field
                    'ALIAS'
)

$fieldsStaff = @(
                    'SFKEY'
                    'PREF_NAME'
                    'FIRST_NAME'
                    'SURNAME'
                    'BIRTHDATE'
                    'GENDER'
                    'FINISH'
                    'HOMEKEY'           #Used to lookup address details, to get home address details
                    'E_MAIL'
                    'MAILKEY'           #Used to lookup address details, to get mailing address details
                    'PAYROLL_REC_NO'
                    'SECOND_NAME'
                    'TITLE'
                    'MOBILE'
                    'WORK_PHONE'
                    'STAFF_STATUS'
                    'USERNAME'          #Allow this if you want to use the username field
                    'ALIAS'
)

$fieldsFamily = @(
                    'DFKEY'
                    'E_MAIL_A'
                    'MOBILE_A'
                    'HOMEKEY'
)

$fieldsAddress = @(
                    'UMKEY'
                    'ADDRESS01'
                    'ADDRESS02'
                    'ADDRESS03'
                    'STATE'
                    'POSTCODE'
                    'TELEPHONE'
                    'MOBILE'
)

$fieldsYearLevel = @(
                    'KCYKEY'
                    'DESCRIPTION'
)

#-----------------------------------------------------------[Functions]------------------------------------------------------------

function Join-eduHubDelta
{
    Param 
    (
        [Parameter(Mandatory=$true)][string]$file1, 
        [Parameter(Mandatory=$true)][string]$file2, 
        [Parameter(Mandatory=$true)][string]$outputPath,
        [Parameter(Mandatory=$true)][string]$matchAttribute,
        [boolean]$force
    )
    
    ## Merte the two files if the second file is newer than the second
    if (((Get-Item $file1).LastWriteTime -lt (Get-Item $file2).LastWriteTime) -or $force -eq $true)
    {
        LogWrite "Delta File is newer than Base File, Merging" -noOutput:$true
        
        ### Set the output file location
        
        ### Test to ensure output path is vaild, if not create it
        if(!(test-path $outputPath))
        {
              New-Item -ItemType Directory -Force -Path $outputPath | Out-Null
        }
        
        $outputFile = "$outputPath\$((Get-Item $file1).Name)"

        $file1Import = Import-Csv -Path $file1
        $file2Import = Import-Csv -Path $file2

        foreach ($record in $file2Import)
        {
            if ($file1Import.$matchAttribute -contains $record.$matchAttribute)
            {
                LogWrite "Record ($($record.$matchAttribute)) Matches Existing Record, Merging" -noOutput:$true
                
                foreach ($row in $file1Import)
                {
                    if ($row.$matchAttribute -eq $record.$matchAttribute)
                    {
                        $row = $record
                    }
                }
            }
            else
            {
                LogWrite "New Record Found, Inserting" -noOutput:$true
                
                return $file1
                $file1Import += $record
            }
        }

        $file1Import | Export-CSV $outputFile -Encoding ASCII  -NoTypeInformation

        return $outputFile
    }
    else
    { 
        LogWrite "Newer file not detected, skipping" -noOutput:$true
        return $file1
    }
}

Function Merge-User
{
    Param 
    (
        [Parameter(Mandatory=$true)]$workingUser, 
        [Parameter(Mandatory=$true)][float]$exitAfter, 
        [Parameter(Mandatory=$true)][int]$handlingEmail,
        [Parameter(Mandatory=$true)][float]$handlingUsername,
        [Parameter(Mandatory=$true)][int]$handlingAlias,
        [switch]$userStaff = $false
    )
  
  Begin{
    #Log-Write -LogPath $sLogFile -LineValue "<description of what is going on>..."
  }
  
  Process{
    Try
    {   
        #If user is marked as left, run exit check
        if ($workingUser.STATUS -eq "INAC" -or $workingUser.STATUS -eq "DEL")
        {
            return $null
        }
        elseif ($workingUser.STATUS -eq "LEFT")
        {
            #Handle Exited Users if they are not meant to be exported, return null, else continue
            if ($exitAfter -ge 1)
            {
                #check if current date is more than $exitAfter days after the users exited date
                if(((($currentDate - (Get-Date $workingUser.FINISH)).Days) -gt $exitAfter))
                {
                    return $null
                }
            }
            elseif ($exitedAfter -eq 0)
            {
                return $null
            }
        } 

        Write-Host "Processing $($workingUser.SIS_ID)"
        
        if ($adCheck)
        {
            #Check AD User existence if $handingExportNoUser = $false or $handlingValidateLicencing is true and there is no user that has the ID in the samAccountName, UserPrincipalName or EmployeeID Fields
            if (
                (-not $handlingExportNoUser -or $handlingValidateLicencing) -and
                    (
                        (
                            (-not $handlingExportNoUser -or $handlingValidateLicencing) -and
                            (
                                (
                                    ($ADUsers.samAccountName -notcontains $workingUser.SIS_ID) -and 
                                    ($ADUsers.employeeID -notcontains $workingUser.SIS_ID) -and
                                    (@($ADUsers.UserPrincipalName -like "$($workingUser.SIS_ID)@*").count -eq 0 )
                                ) -and 
                                ( 
                                    -not $userStaff -or 
                                    (
                                        $userStaff -and
                                        (
                                            ($ADUsers.samAccountName -notcontains $workingUser.SIS_EMPNO) -and
                                            (@($ADUsers.UserPrincipalName -like "$($workingUser.SIS_EMPNO)@*").Count -eq 0 ) -and
                                            ($ADUsers.employeeID -notcontains $workingUser.SIS_EMPNO)
                                        )
                                    )
                                )
                            )
                        )
                    )
                )
            {
                Write-Host "Dropping user $($workingUser.SIS_ID)$(if ( -not [string]::IsNullOrWhiteSpace($workingUser.SIS_EMPNO)){"|$($workingUser.SIS_EMPNO)"}) because there is no AD User and either licencing validation is in effect or exporting of users without an AD account is disabled"
                return $null
                
            }
            
            $AD_User = $null
            $AD_User = $ADusers | where-object { 
                $_.samAccountName -eq $workingUser.SIS_ID -or 
                 ($_.employeeID -contains $workingUser.SIS_ID -and 
                     (
                         ((@($ADUsers | Where-Object -Property employeeID -eq  $workingUser.SIS_ID).Count) -eq 1) -or 
                         ((@($ADUsers | Where-Object -Property employeeID -eq  $workingUser.SIS_ID).Count) -gt 1 -and $_.Enabled -eq $true)
                     )
                 ) -or 
                 ($_.UserPrincipalName -like "$($workingUser.SIS_ID)@*") -or
                 (
                     (
                         ($userStaff -and $null -ne [string]::IsNullOrWhiteSpace($workingUser.SIS_EMPNO)) -and 
                         (
                             ($_.samAccountName -contains $workingUser.SIS_EMPNO) -or
                             ($_.employeeID -contains $workingUser.SIS_EMPNO -and 
                                 (
                                     ((@($ADUsers | Where-Object -Property employeeID -eq  $workingUser.SIS_EMPNO).Count) -eq 1) -or 
                                     ((@($ADUsers | Where-Object -Property employeeID -eq  $workingUser.SIS_EMPNO).Count) -gt 1 -and $_.Enabled -eq $true)
                                 )
                             ) -or 
                             ($_.UserPrincipalName -like "$($workingUser.SIS_EMPNO)@*")
                         )
                     )
             
                 ) 
             }
            
            if ($null -eq $AD_User)
            {
                Write-Host "NULL AD: $($AD_User.samAccountName) | $($workingUser.SIS_ID)"
                Pause
            }

            #Validate the licencing if required
            if ($handlingValidateLicencing -eq $true -and [string]::IsNullOrWhiteSpace($AD_User.$handlingLicencingValue))
            {
                Write-Host "Dropping User $($workingUser.SIS_ID) as licencing check fails"
                return $null
            }
        }

       #Email Handling
       switch ($handlingEmail)
        {
            #1 = Use eduHub Email
            1   {
                    #Do nothing, using eduHub Email
                }
            #2 = Calculate from eduHub Data (SIS_ID)@domain
            2   {
                    if ([string]::IsNullOrWhiteSpace($schoolEmailDomain))
                    {
                        Write-Host "Email Domain Blank but told to use in settings, exiting"
                        exit
                    }
                    $workingUser.E_MAIL = "$(($workingUser.SIS_ID).ToLower())@$schoolEmailDomain"
                }
            #3 = pull from Active Directory UPN - else fall back to mail - else fallback to eduhub
            3   {
                    if (-not [string]::IsNullOrWhiteSpace($AD_User.UserPrincipalName))
                    {
                        $workingUser.E_MAIL = $AD_User.UserPrincipalName
                    }
                }
            #4 = Pull from Active Directory Mail - else fallback to eduhub
            4   {
                    if (-not [string]::IsNullOrWhiteSpace($AD_User.Mail))
                    {
                        $workingUser.E_MAIL = $AD_User.Mail
                    }
                }
            #5 = Pull from Active Directory ProxyAddresses looking for primary (Capital SMTP) - else fall back to mail - else fallback to eduhub
            5   {
                    if (-not [string]::IsNullOrWhiteSpace($AD_User.proxyAddresses -clike "SMTP:*"))
                    {
                        $workingUser.E_MAIL = ((($AD_User.proxyAddresses -clike "SMTP:*")[0]).SubString(5)).ToLower()
                    }
                }
            #6 = Use employeeID (PAYROLL_REC_NO/SIS_EMPNO/EmployeeNumber) from Active Directory, fall back to eduHub
            6   {
                    if (-not [string]::IsNullOrWhiteSpace($AD_User.EmployeeID))
                    {
                        $workingUser.E_MAIL = "$($AD_User.EmployeeID)@$schoolEmailDomain"
                    }
                }
            #7 = Use employeeID (PAYROLL_REC_NO/SIS_EMPNO/EmployeeNumber) from eduHub, fall back to SIS_ID
            {$_ -eq 7 -and $userStaff}
                {
                    if (-not [string]::IsNullOrWhiteSpace($workingUser.SIS_EMPNO))
                    {
                        $workingUser.E_MAIL = "$(($workingUser.SIS_EMPNO).ToLower())@$schoolEmailDomain"
                    }
                    else
                    {
                        $workingUser.E_MAIL = "$(($workingUser.SIS_ID).ToLower())@$schoolEmailDomain"
                    }
                }
            #default = Use eduHub Email
            default 
                {
                    #Do nothing, using eduHub Email
                }
        }

        if(($userStaff -eq $false -and $fieldsStudent.Contains("USERNAME")) -or ($userStaff -eq $true -and $fieldsStaff.Contains("USERNAME")))
        {
            #Username Handling
            switch ($handlingUsername)
            {
                #0 = Blank the field ""
                0 
                    {
                        $workingUser.USERNAME = ""
                    }
                #1 = Use eduHub Key
                1 
                    {
                       $workingUser.USERNAME = ($workingUser.SIS_ID).ToUpper()
                    }
                #2 = Calculate from eduHub Data (SIS_ID)@domain
                2   {
                        if ([string]::IsNullOrWhiteSpace($schoolEmailDomain))
                        {
                            Write-Host "Email Domain Blank but told to use in settings, exiting"
                            exit
                        }

                        $workingUser.USERNAME = "$(($workingUser.SIS_ID).ToLower())@$schoolEmailDomain"

                    }
                #3 = pull from Active Directory UPN
                3   {
                        if (-not [string]::IsNullOrWhiteSpace($AD_User.UserPrincipalName))
                        {
                            $workingUser.USERNAME = $AD_User.UserPrincipalName
                        }
                    }
                #4 = Pull from Active Directory Mail
                4   {
                        if (-not [string]::IsNullOrWhiteSpace($AD_User.Mail))
                        {
                            $workingUser.USERNAME = $AD_User.Mail
                        }
                    }
                #5 = Pull from Active Directory ProxyAddresses looking for primary (Capital SMTP)
                5   {
                        if (-not [string]::IsNullOrWhiteSpace($AD_User.proxyAddresses -clike "SMTP:*"))
                        {
                            $workingUser.USERNAME = ((($AD_User.proxyAddresses -clike "SMTP:*")[0]).SubString(5)).ToLower()
                        }
                        elseif(-not [string]::IsNullOrWhiteSpace($AD_User.Mail))
                        {
                            $workingUser.USERNAME = $AD_User.Mail
                        }
                        else
                        {
                            $workingUser.USERNAME = ($workingUser.SIS_ID).ToUpper()
                        }
                    }
                #6 = samAccountName
                6   {
                        if (-not [string]::IsNullOrWhiteSpace($AD_User.SamAccountName))
                        {
                            $workingUser.USERNAME = $AD_User.SamAccountName
                        }
                    }
                #7 = Use employeeID (PAYROLL_REC_NO/SIS_EMPNO/EmployeeNumber) from Active Directory, fall back to SIS_ID
                7   {
                        if (-not [string]::IsNullOrWhiteSpace($AD_User.EmployeeID))
                        {
                            $workingUser.USERNAME = "$($AD_User.EmployeeID)"
                        }
                    }
                #8 = Use employeeID (PAYROLL_REC_NO/SIS_EMPNO/EmployeeNumber) from eduHub, fall back to SIS_ID
                {$_ -eq 8 -and $userStaff -eq $true}

                    {
                        if (-not [string]::IsNullOrWhiteSpace($workingUser.SIS_EMPNO))
                        {
                            $workingUser.USERNAME = $workingUser.SIS_EMPNO
                        }
                        else
                        {
                            $workingUser.USERNAME = $workingUser.SIS_ID
                        }
                    }
                #Default = Use eduHub Key (SIS_ID)
                default 
                    {
                        if (-not $userStaff)
                        {
                            $workingUser.USERNAME = ($workingUser.SIS_ID).ToUpper()
                        }
                        else
                        {
                            $workingUser.USERNAME = ($workingUser.SIS_ID).ToUpper()
                        }
                    }
            }
        }

        #Alias Handling
        if(($userStaff -eq $false -and $fieldsStudent.Contains("ALIAS")) -or ($userStaff -eq $true -and $fieldsStaff.Contains("ALIAS")))
        {
            switch ($handlingAlias)
            {

                #1 = Use eduHub Key (SIS_ID)
                1 
                    {
                        if (-not $userStaff)
                        {
                            $workingUser.ALIAS = ($workingUser.SIS_ID).ToUpper()
                        }
                        else
                        {
                            $workingUser.ALIAS = ($workingUser.SIS_ID).ToUpper()
                        }
                    }
                #2 = Use samAccountName from Active Directory
                2   {
                        if (-not [string]::IsNullOrWhiteSpace($AD_User.samAccountName))
                        {
                            $workingUser.ALIAS = $AD_User.SamAccountName
                        }
                    }
                #3 = Use employeeID from Active Directory
                3   {
                        if (-not [string]::IsNullOrWhiteSpace($AD_User.EmployeeID))
                        {
                            $workingUser.ALIAS = $AD_User.EmployeeID
                        }
                    }
                #4 = Use employeeID (PAYROLL_REC_NO/SIS_EMPNO/EmployeeNumber) from eduHub Data
                {$_ -eq 4 -and $userStaff -eq $true}

                    {
                        if (-not [string]::IsNullOrWhiteSpace($workingUser.SIS_EMPNO))
                        {
                            $workingUser.ALIAS = $workingUser.SIS_EMPNO
                        }
                        else
                        {
                            $workingUser.ALIAS = $workingUser.SIS_ID
                        }
                    }
                #Default = Use eduHub Key (SIS_ID)
                default 
                    {
                        if (-not $userStaff)
                        {
                            $workingUser.ALIAS = ($workingUser.SIS_ID).ToUpper()
                        }
                        else
                        {
                            $workingUser.ALIAS = ($workingUser.SIS_ID).ToUpper()
                        }
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

###################### Retrieve AD Users if Required ######################

$ADUsers = $null

if ($handlingValidateLicencing -or -not $handlingExportNoUser -or (($handlingStudentEmail -ge 3) -and ($handlingStudentEmail -le 5)) -or (($handlingStudentUsername -ge 3) -and ($handlingStudentUsername -le 5)) -or (($handlingStudentAlias -ge 2) -and ($handlingStudentAlias -le 3)) -or (($handlingStaffEmail -ge 3) -and ($handlingStaffEmail -le 6)) -or (($handlingStaffUsername -ge 3) -and ($handlingStaffUsername -le 6)) -or (($handlingStaffAlias -ge 2) -and ($handlingStaffAlias -le 3)))
{
    $adCheck = $true
    try 
    {
        Import-Module ActiveDirectory
        Write-Host "Activating Active Directory Module"
    }
    catch
    {
        throw "Cannot Load Active Directory Module"
    }
    
    if ([string]::IsNullOrWhiteSpace($activeDirectoryServer))
    {
        Write-Host "Active Directory use is required, but no server is specfied"
        exit
    }

    try 
    {
        #Create Splat
        $activeDirectorySplat = @{
            Filter = '*'
            Server = $activeDirectoryServer
        }

        $adPropertiesList = @(
            "employeeID"
            "Mail"
            "ProxyAddresses"
        )

        #If licencing check is turned on and the value for the licencing variable is not blank then add it to the properties array
        if ($handlingValidateLicencing -and -not [string]::IsNullOrWhiteSpace($handlingLicencingValue))
        {
            $adPropertiesList += $handlingLicencingValue
        }
        #If licencing check is turned on and the value for the licencing variable is blank then error out
        elseif ($handlingValidateLicencing -and [string]::IsNullOrWhiteSpace($handlingLicencingValue))
        {
            Write-Host "Told to Validate Licencing but no AD field with licencing Value specified, Exiting"
            exit
        }
        
        #Add Properties Array to Splat now that it is calculated        
        $activeDirectorySplat.Properties = $adPropertiesList

        #Check if there is a Search Base set, if so add it to the splat

        if (-not [string]::IsNullOrWhiteSpace($activeDirectorySearchBase))
        {
            $activeDirectorySplat.SearchBase = $activeDirectorySearchBase
        }

        if ($runAsLoggedIn -eq $false)
        {
            try
            {
                Write-Host "Attempting to import Authentication Module"
                Import-Module "$PSScriptRoot\Modules\Authentication.psm1"
            }
            catch
            {
                Write-Host "Cannot Import Authentication Module"
            }
            
            if ([string]::IsNullOrWhiteSpace($activeDirectoryUser))
            {
                Write-Host "Active Directory use is required, but no credentials are specfied and running as logged in user is disabled"
                exit
            }

            #$schoolServiceCreds = Get-SavedCredentials_WithRequest "$PSScriptRoot\Credentials\schoolDC-$([Environment]::MachineName)-$([Environment]::UserName).crd" $activeDirectoryUser
            #$schoolServiceCreds = new-object -typename System.Management.Automation.PSCredential -argumentlist $schoolServiceCreds.Username,$schoolServiceCreds.Password
            #$activeDirectorySplat.Credentials = $schoolServiceCreds

        }

        $ADUsers = Get-ADUser @activeDirectorySplat | Sort-Object samAccountName
    }
    catch 
    {
        Write-Output "Error with AD Query, Exiting"
        exit
    }

}


######################Import and Process Students######################

$importedStudents = $null
$workingStudents = @()

#Import Students from CSV(s) based upon settings

if ($includeDeltas -eq $true) #Only do Delta join if not using files from exporter as exporter joins the files
{
    $importedStudents = Import-CSV (Join-eduHubDelta (Join-Path -Path $fileImportLocation -ChildPath $importFileStudents) (Join-Path -Path $fileImportLocation -ChildPath $importFileStudentsDelta) "$PSScriptRoot\Cache\" "STKEY")  | Sort-Object -property STATUS, STKEY
}
elseif ($includeDeltas -eq $false) #Only Run import if not using modified headers from exporter
{
    $importedStudents = Import-CSV (Join-Path -Path $fileImportLocation -ChildPath $importFileStudents) | Sort-Object -property STATUS, STKEY
}
else
{
    throw "Cannot Import Error with locating or processing files"
}



#Handle eduHub headers vs required headers
$headersStudent = $null
$headersStudent = (($importedStudents | Select-Object -First 1).psobject.properties).Name
$importedStudents = $importedStudents | Select-Object ($fieldsStudent + @("CONTACT_A")) #Selecting only required fields and adding Contact_A for family/address processing

#Alias keys for processing
$importedStudents = $importedStudents | Add-Member -MemberType AliasProperty -Name SIS_ID -Value STKEY -PassThru | Add-Member -MemberType AliasProperty -Name FINISH -Value EXIT_DATE -PassThru

#Process Students
foreach ($student in $importedStudents)
{
    $tempUser = $null
    if ($null -ne ($tempUser = (Merge-User -workingUser $student -exitAfter $handlingStudentExitAfter -handlingEmail $handlingStudentEmail -handlingUsername $handlingStudentUsername -handlingAlias $handlingStudentAlias)))
    {
        $workingStudents += $tempUser
    }
}

#Merge non-eduhub fields into end of eduhub headers if they are required
foreach ($field in $fieldsStudent)
{
    if ($headersStudent -notcontains $field)
    {
        $headersStudent += $field
    }
}

$importedStudents = $null #Explicitly destroy data to clear up resources

###################### Import and Process Staff ######################

$importedStaff = $null
$workingStaff = @()

#Import Staff from CSV(s) based upon settings

if ($includeDeltas -eq $true) #Only do Delta join if not using files from exporter as exporter joins the files
{
    $importedStaff = Import-CSV (Join-eduHubDelta (Join-Path -Path $fileImportLocation -ChildPath $importFileStaff)(Join-Path -Path $fileImportLocation -ChildPath $importFileStaffDelta) "$PSScriptRoot\Cache\" "SFKEY")  | Sort-Object -property STAFF_STATUS, SFKEY
}
elseif ($includeDeltas -eq $false) #Only Run import if not using modified headers from exporter
{
    $importedStaff = Import-CSV (Join-Path -Path $fileImportLocation -ChildPath $importFileStaff) | Sort-Object -property STAFF_STATUS, SFKEY
}
else
{
    throw "Cannot Import Error with locating or processing files"
}

#Handle eduHub headers vs required headers
$headersStaff = $null
$headersStaff = (($importedStaff |Select-Object -First 1).psobject.properties).Name
$importedStaff = $importedStaff | Select-Object $fieldsStaff #Selecting only required fields 

#Alias keys for data processing
$importedStaff = $importedStaff | Add-Member -MemberType AliasProperty -Name SIS_ID -Value SFKEY -PassThru | Add-Member -MemberType AliasProperty -Name STATUS -Value STAFF_STATUS -PassThru | Add-Member -MemberType AliasProperty -Name SIS_EMPNO -Value PAYROLL_REC_NO -PassThru

#Process Staff
foreach ($staff in $importedStaff)
{
    if ($null -ne ($tempUser = (Merge-User -workingUser $staff -exitAfter $handlingStaffExitAfter -handlingEmail $handlingStaffEmail -handlingUsername $handlingStaffUsername  -handlingAlias $handlingStaffAlias -userStaff )))
    {
        $workingStaff += $tempUser
    }
}

#Merge non-eduhub fields into end of eduhub headers if they are required
foreach ($field in $fieldsStaff)
{
    if ($headersStaff -notcontains $field)
    {
        $headersStaff += $field
    }
}

$importedStaff = $null #Explicitly destroy data to clear up resources


#Process Staff for Export here so that futher processing has a clean set of data, specifically addresses are cleared if not wanted

###################### Import and Process Families ######################

$importedFamilies = $null
$workingFamilies = @()

#Import Families from CSV(s) based upon settings - No modified headers here as there is no need due to their only being the one table of this type

if ($includeDeltas -eq $true) #Only do Delta join if not using files from exporter as exporter joins the files
{
    $importedFamilies = Import-CSV (Join-eduHubDelta (Join-Path -Path $fileImportLocation -ChildPath $importFileFamilies) (Join-Path -Path $fileImportLocation -ChildPath $importFileFamiliesDelta) "$PSScriptRoot\Cache\" "DFKEY") | Sort-Object -property DFKEY
}
else
{
    $importedFamilies = Import-CSV (Join-Path -Path $fileImportLocation -ChildPath $importFileFamilies) | Sort-Object -property DFKEY
}

#Handle eduHub headers vs required headers
$headersFamily = $null
$headersFamily = (($importedFamilies |Select-Object -First 1).psobject.properties).Name
$importedFamilies = $importedFamilies | Select-Object ($fieldsFamily + @("E_MAIL_B") + @("MOBILE_B")) #Selecting only required fields and fields for processing only data

#Sort families so that only families where there is an active student are kept and that are due to be exported, then with an active family check to see if primary contact is contact B (A and C are left as A), if so change the details, Contact B is dropped on export. If there use only the first record (usally the oldest student) to calculate this

foreach ($family in $importedFamilies)
{
    if ($workingStudents.FAMILY -match $family.DFKEY)
    {
        if ((($workingStudents | Where-Object {$_.FAMILY -eq $family.DFKEY} | Sort-Object -Property SIS_ID | select-object -First 1).CONTACT_A) -eq "B")
        {
            $family.E_MAIL_A = $family.E_MAIL_B
            $family.MOBILE_A = $family.MOBILE_B
            Write-Host "Changing Contacts for $($family.DFKEY)"
        }
        elseif (((($workingStudents | Where-Object {$_.FAMILY -eq $family.DFKEY} | Sort-Object -Property SIS_ID | select-object -First 1).CONTACT_A) -eq "C") -and -not [string]::IsNullOrWhiteSpace($family.E_MAIL_B) -and ($family.E_MAIL_B -ne $family.E_MAIL_A))
        {
            $family.E_MAIL_A += ";$($family.E_MAIL_B)"
            Write-Host "Adding Secondary Email for $($family.DFKEY)"
        }
        $workingFamilies += $family
    }
    
}

$importedFamilies = $null #Explicitly destroy data to clear up resources

###################### Import and Process Addresses ######################

$importedAddresses = $null
$workingAddresses = @()

#Import Addresses from CSV(s) based upon settings - No modified headers here as there is no need due to their only being the one table of this type

if ($includeDeltas -eq $true) #Only do Delta join if not using files from exporter as exporter joins the files
{
    $importedAddresses = Import-CSV (Join-eduHubDelta (Join-Path -Path $fileImportLocation -ChildPath $importFileAddresses) (Join-Path -Path $fileImportLocation -ChildPath $importFileAddressesDelta) "$PSScriptRoot\Cache\" "UMKEY") | Sort-Object UMKEY
}
else
{
    $importedAddresses = Import-CSV (Join-Path -Path $fileImportLocation -ChildPath $importFileAddresses)  | Sort-Object UMKEY
}

#Handle eduHub headers vs required headers
$headersAddress = $null
$headersAddress = (($importedAddresses |Select-Object -First 1).psobject.properties).Name
$importedAddresses = $importedAddresses | Select-Object $fieldsAddress #Selecting only required fields

#Sort through addresses, only keeping those where there is a family or staff member that are due to be exported associated with the address
foreach ($address in $importedAddresses)
{
    if (($workingFamilies.HOMEKEY -match $address.UMKEY) -or ($workingStaff.HOMEKEY -match $address.UMKEY) -or ($workingStaff.MAILKEY -match $address.UMKEY))
    {
        $workingAddresses += $address
    }
    
}

$importedAddresses = $null #Explicitly destroy data to clear up resources


###################### Import and Process Year Level Descriptions ######################

$importedYearLevels = $null
$workingYearLevels = @()

#Import Year Levels from CSV(s) based upon settings - No modified headers here as there is no need due to their only being the one table of this type - No Delta's for this file

$importedYearLevels = Import-CSV (Join-Path -Path $fileImportLocation -ChildPath $importFileYearLevels)  | Sort-Object KCYKEY

#Handle eduHub headers vs required headers
$headersYearLevel = $null
$headersYearLevel = (($importedYearLevels |Select-Object -First 1).psobject.properties).Name
$importedYearLevels = $importedYearLevels | Select-Object $fieldsYearLevel #Selecting only required fields

#Sort through YearLevels, only keeping those where there is a family or staff member that are due to be exported associated with the Year Level
foreach ($YearLevel in $importedYearLevels)
{
    if ($YearLevel.DESCRIPTION -ne "DO NOT USE")
    {
        if (($YearLevel.DESCRIPTION -match "Year [0-9]" -and $YearLevel.DESCRIPTION -notmatch "Year [0-1][0-9]") -and $handlingFileYearLevel -eq 2)
        {
            $YearLevel.DESCRIPTION = "Year 0$(($YearLevel.DESCRIPTION.Trim()).SubString(5,1))"
        }

        $workingYearLevels += $YearLevel
    }
    
}

$importedYearLevels = $null #Explicitly destroy data to clear up resources

###################### Process Data for Export ######################

if ($exportFull -and !$exportCustom)
{
    $workingStudents = $workingStudents | Select-Object $fieldsStudent | Select-Object $headersStudent #Double Conversion to clear processing data and then re-instate eduhub fields in the correct order
    $workingStaff = $workingStaff | Select-Object $fieldsStaff | Select-Object $headersStaff #Double Conversion to clear processing data and then re-instate eduhub fields in the correct order
    $workingFamilies = $workingFamilies | Select-Object $fieldsFamily | Select-Object $headersFamily #Double Conversion to clear processing data and then re-instate eduhub fields in the correct order
    $workingAddresses = $workingAddresses | Select-Object $fieldsAddress | Select-Object $headersAddress #Double Conversion to clear processing data and then re-instate eduhub fields in the correct order
    $workingYearLevels = $workingYearLevels | Select-Object $fieldsYearLevel | Select-Object $headersYearLevel #Double Conversion to clear processing data and then re-instate eduhub fields in the correct order
}
elseif (!$exportfull -and !$exportCustom)
{
    $workingStudents = $workingStudents | Select-Object $fieldsStudent #Single Conversion to clear processing data
    $workingStaff = $workingStaff | Select-Object $fieldsStaff #Single Conversion to clear processing data
    $workingFamilies = $workingFamilies | Select-Object $fieldsFamily #Single Conversion to clear processing data
    $workingAddresses = $workingAddresses | Select-Object $fieldsAddress #Single Conversion to clear processing data
    $workingYearLevels = $workingYearLevels | Select-Object $fieldsYearLevel #Single Conversion to clear processing data
}

elseif($exportCustom)
{
    #Iterate and Process Students
    foreach ($student in $workingStudents)
    {
        $studentFamily = $null
        $studentAddress = $null
        $studentFamily = $workingFamilies | WHERE-OBJECT DFKEY -eq $student.FAMILY
        $studentAddress = $workingAddresses | WHERE-OBJECT UMKEY -eq $studentFamily.HOMEKEY
        

        if($studentFamily.E_MAIL_A -notmatch ";" -and ![string]::IsNullOrWhiteSpace($studentFamily.E_MAIL_A))
        {
            $student | Add-Member -Type NoteProperty -Name "PARENT_E_MAIL_A" -Value ($studentFamily.E_MAIL_A)
            $student | Add-Member -Type NoteProperty -Name "PARENT_E_MAIL_B" -Value ""
        }
        elseif ($studentFamily.E_MAIL_A -match ";")
        {
            $split = $null
            $split = ($studentFamily.E_MAIL_A) -split ";"
            $student | Add-Member -Type NoteProperty -Name "PARENT_E_MAIL_A" -Value $split[0] 
            $student | Add-Member -Type NoteProperty -Name "PARENT_E_MAIL_B" -Value $split[1] 
        }

        if($studentFamily.MOBILE_A -notmatch ";" -and ![string]::IsNullOrWhiteSpace($studentFamily.MOBILE_A))
        {
            $student | Add-Member -Type NoteProperty -Name "PARENT_MOBILE_A" -Value ($studentFamily.MOBILE_A)
            $student | Add-Member -Type NoteProperty -Name "PARENT_MOBILE_B" -Value ""
        }
        elseif ($studentFamily.MOBILE_A -match ";")
        {
            $split = $null
            $split = ($studentFamily.MOBILE_A) -split ";"
            $student | Add-Member -Type NoteProperty -Name "PARENT_MOBILE_A" -Value $split[0] 
            $student | Add-Member -Type NoteProperty -Name "PARENT_MOBILE_B" -Value $split[1] 
        }

        $student | Add-Member -Type NoteProperty -Name "ADDRESS01" -Value ($studentAddress.ADDRESS01)
        $student | Add-Member -Type NoteProperty -Name "ADDRESS02" -Value ($studentAddress.ADDRESS02)
        $student | Add-Member -Type NoteProperty -Name "ADDRESS03" -Value ($studentAddress.ADDRESS03)
        $student | Add-Member -Type NoteProperty -Name "STATE" -Value ($studentAddress.STATE)
        $student | Add-Member -Type NoteProperty -Name "POSTCODE" -Value ($studentAddress.POSTCODE)

        if($StudentAddress.MOBILE -ne $student.PARENT_MOBILE_A -and $StudentAddress.MOBILE -ne $student.PARENT_MOBILE_B)
        {
            if([string]::IsNullOrWhiteSpace($student.PARENT_MOBILE_A))
            {
                $student | Add-Member -Type NoteProperty -Name "PARENT_MOBILE_A" -Value ($studentAddress.MOBILE)
            }
            elseif([string]::IsNullOrWhiteSpace($student.PARENT_MOBILE_A))
            {
                $student | Add-Member -Type NoteProperty -Name "PARENT_MOBILE_B" -Value ($studentAddress.MOBILE)
            }
        }

        if($StudentAddress.TELEPHONE -ne $student.PARENT_MOBILE_A -and $StudentAddress.TELEPHONE -ne $student.PARENT_MOBILE_B)
        {
                $student | Add-Member -Type NoteProperty -Name "TELEPHONE" -Value ($studentAddress.TELEPHONE)
        }
        else 
        {
            $student | Add-Member -Type NoteProperty -Name "TELEPHONE" -Value ""
        }

    }

    #Iterate and process Staff
    foreach ($staffMember in $workingStaff)
    {
        $staffMemberAddress = $null
        $staffMemberAddress = $workingAddresses | WHERE-OBJECT UMKEY -eq $staffMember.MAILKEY
        
        $staffMember | Add-Member -Type NoteProperty -Name "ADDRESS01" -Value ($staffMemberAddress.ADDRESS01)
        $staffMember | Add-Member -Type NoteProperty -Name "ADDRESS02" -Value ($staffMemberAddress.ADDRESS02)
        $staffMember | Add-Member -Type NoteProperty -Name "ADDRESS03" -Value ($staffMemberAddress.ADDRESS03)
        $staffMember | Add-Member -Type NoteProperty -Name "STATE" -Value ($staffMemberAddress.STATE)
        $staffMember | Add-Member -Type NoteProperty -Name "POSTCODE" -Value ($staffMemberAddress.POSTCODE)

        if(!([string]::IsNullOrWhiteSpace($staffMemberAddress.MOBILE)) -and $staffMemberAddress.MOBILE -ne $staffMember.MOBILE)
        {
                $staffMember.MOBILE = $staffMemberAddress.MOBILE
        }

        if($staffMemberAddress.TELEPHONE -ne $staffMember.MOBILE)
        {
                $staffMember | Add-Member -Type NoteProperty -Name "TELEPHONE" -Value ($staffMemberAddress.TELEPHONE)
        }
        else 
        {
            $staffMember | Add-Member -Type NoteProperty -Name "TELEPHONE" -Value ""
        }
    }
}

###################### Export Data ######################
Get-ChildItem -Path $fileOutputLocation -Include *.* -File -Recurse | ForEach-Object { $_.Delete()}

if(!(Test-Path($fileOutputLocation)))
{
    New-Item -Path $fileOutputLocation -ItemType Directory | Out-Null
}

if (!$exportCustom)
{
    $workingStudents | ConvertTo-Csv -NoTypeInformation | Out-File (Join-Path -Path $fileOutputLocation -ChildPath $importFileStudents) -encoding ascii
    $workingStaff | ConvertTo-Csv -NoTypeInformation | Out-File (Join-Path -Path $fileOutputLocation -ChildPath $importFileStaff) -encoding ascii
    $workingFamilies | ConvertTo-Csv -NoTypeInformation | Out-File (Join-Path -Path $fileOutputLocation -ChildPath $importFileFamilies) -encoding ascii
    $workingAddresses | ConvertTo-Csv -NoTypeInformation | Out-File (Join-Path -Path $fileOutputLocation -ChildPath $importFileAddresses) -encoding ascii
    $workingYearLevels | ConvertTo-Csv -NoTypeInformation | Out-File (Join-Path -Path $fileOutputLocation -ChildPath $importFileYearLevels) -encoding ascii
}
else 
{
    $workingStudents | ConvertTo-Csv -NoTypeInformation | Out-File (Join-Path -Path $fileOutputLocation -ChildPath $importFileStudents) -encoding ascii
    $workingStaff | ConvertTo-Csv -NoTypeInformation | Out-File (Join-Path -Path $fileOutputLocation -ChildPath $importFileStaff) -encoding ascii
}

#Log-Finish -LogPath $sLogFile