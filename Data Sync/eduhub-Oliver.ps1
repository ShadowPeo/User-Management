

param 
    (
        
        #School Details
        [string]$schoolID = "3432", # Used for export and for import if using CASES File Names
        #$schoolID = [system.environment]::MachineName.Trim().Substring(0,4)

        [string]$schoolEmailDomain = "mwps.vic.edu.au", #Only used if processing emails or users from CASES Data

        #File Settings
        [boolean]$modifiedHeaders = $false, #Use Modified Export Headers (from export script in this Repo), if not it will look for standard eduHub headers
        [boolean]$includeDeltas = $true, #Include eduHub Delta File

        #File Locations
        [string]$fileLocation = "$PSSCriptRoot\Import",
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
        [int]$handlingFileYearLevel = 1, # 1 = Static (use the one from cache, if not exist cache copy and us as literal) 2 = Use Literal, description will e exported exactly as is. 3 = Pad the year numbers (if they exist) in the description field
        [boolean]$handlingIncludeFutures = $true, #Include Future Students
        [int]$handlingStudentEmail = 1, #1 = Use eduHub Email, 2 = Calculate from eduHub Data (SIS_ID)@domain, 3 = pull from AD UPN, 4 = Pull from AD Mail, 5 = Pull from AD ProxyAddresses looking for primary (Capital SMTP)
        [int]$handlingStaffEmail = 1, #1 = Use eduHub Email, 2 = Calculate from eduHub Data (SIS_ID)@domain, 3 = pull from AD UPN, 4 = Pull from AD Mail, 5 = Pull from AD ProxyAddresses looking for primary (Capital SMTP),  6 = Use employeeID (PAYROLL_REC_NO/SIS_EMPNO/EmployeeNumber) from AD, fall back to SIS_ID, 7 = Use employeeID (PAYROLL_REC_NO/SIS_EMPNO/EmployeeNumber) from eduHub Data, fall back to SIS_ID
        [float]$handlingStudentUsername = 1, #-1 = Exclude from Export, #0 = Blank, 1 = use eduHub Data (SIS_ID), 2 = Calculate from eduHub Data (SIS_ID)@domain, 3 = pull from AD UPN, 4 = Pull from AD Mail, 5 Use samAccountName
        [float]$handlingStaffUsername = 1, #-1 = Exclude from Export, #0 = Blank, 1 = use eduHub Data (SIS_ID), 2 = Calculate from eduHub Data (SIS_ID)@domain, 3 = pull from AD UPN, 4 = Pull from AD Mail, 5 Use samAccountName, 6 = Use employeeID (PAYROLL_REC_NO/SIS_EMPNO/EmployeeNumber) from AD, fall back to SIS_ID, 7 = Use employeeID (PAYROLL_REC_NO/SIS_EMPNO/EmployeeNumber) from eduHub Data, fall back to SIS_ID
        [int]$handlingStudentAlias = 1, #1 = SIS_ID, 2= use samAccountName - Fall back to SIS_ID, 3 = Use employeeID from Active Directory - Fall back to SIS_ID
        [int]$handlingStaffAlias = 1, #1 = SIS_ID, 2= use samAccountName, 3 = Use employeeID from Active Directory - Fall back to SIS_ID, 4 = Use employeeID (PAYROLL_REC_NO/SIS_EMPNO/EmployeeNumber) from eduHub Data - Fall back to SIS_ID
        [boolean]$handlingValidateLicencing = $false, #Validate the licencing for Oliver, this will drop accounts where it is explictly disabled or where no user exists 
        [string]$handlingLicencingValue = "licencingOliver", #The attribute name for the licencing Data
        [boolean]$handlingExportNoUser = $false, #Export user if there is no matching username in AD, if AD lookup is in use

        #Active Directory Settings (Only required if using AD lookups - Active Directory lookups rely on the samAccountName being either the Key (SIS_ID) or in the case of staff members PAYROLL_REC_NO/SIS_EMPNO Matches will also be based upon email matching UPN
        [boolean]$runAsLoggedIn = $true,
        [string]$activeDirectoryUser = $null, #Username to connect to AD as, will prompt for password if credentials do not exist or are incorrect, not used if not running as logged in user
        [string]$activeDirectoryServer = "10.128.136.35", #DNS Name or IP of AD Server
        [string]$activeDirectorySearchBase = $null, #DNS Name or IP of AD Server

        #Log File Info
        [string]$sLogPath = "C:\Windows\Temp",
        [string]$sLogName = "<script_name>.log"
    )

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

$sLogFile = Join-Path -Path $sLogPath -ChildPath $sLogName

#Script Variables - Declared to stop it being generated multiple times per run

#Date
$currentDate = Get-Date
$adCheck = $false #Changes to true if one of the settings requires an AD Check

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
            
            Write-Host "$($AD_User.samAccountName) | $($workingUser.SIS_ID)"

            if ($null -eq $AD_User)
            {
                Write-Host "NULL AD: $($AD_User.samAccountName) | $($workingUser.SIS_ID)"
                Pause
            }
            #Validate the licencing if required
            <#if ($handlingValidateLicencing -eq $true)
            {
                return $null
            }#>
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
                    
                }
            #4 = Pull from Active Directory Mail - else fallback to eduhub
            4   {
                    
                }
            #5 = Pull from Active Directory ProxyAddresses looking for primary (Capital SMTP) - else fall back to mail - else fallback to eduhub
            5   {
                    
                }
            #6 = Use employeeID (PAYROLL_REC_NO/SIS_EMPNO/EmployeeNumber) from Active Directory, fall back to SIS_ID
            6   {
                    
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
                    
                }
            #4 = Pull from Active Directory Mail
            4   {
                    
                }
            #5 = Pull from Active Directory ProxyAddresses looking for primary (Capital SMTP)
            5   {
                    
                }
            #6 = Use employeeID (PAYROLL_REC_NO/SIS_EMPNO/EmployeeNumber) from Active Directory, fall back to SIS_ID
            6   {
                    
                }
            #7 = Use employeeID (PAYROLL_REC_NO/SIS_EMPNO/EmployeeNumber) from eduHub, fall back to SIS_ID
            {$_ -eq 7 -and $userStaff -eq $true}

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

        #Alias Handling

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
                    
                }
            #3 = Use employeeID from Active Directory
            3   {
                    
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
    if ([string]::IsNullOrWhiteSpace($activeDirectoryUser) -and $runAsLoggedIn -eq $false)
    {
        Write-Host "Active Directory use is required, but no credentials are specfied and running as logged in user is disabled"
        exit
    }

    try 
    {

       
        if ($runAsLoggedIn -eq $true)
        {
            $ADUsers = Get-ADUser -Server $activeDirectoryServer -Properties employeeID -Filter * | Sort-Object employeeID
        }
        <#if ($runAsLoggedIn -eq $false)
        {
            try
            {
                Import-Module "$PSScriptRoot\Modules\Authentication.ps1"
            }
            catch
            {
                throw "Cannot Import Authentication Module"
            }
            $schoolServiceCreds = Get-SavedCredentials_WithRequest "$PSScriptRoot\Credentials\schoolDC-$([Environment]::MachineName)-$([Environment]::UserName).crd" $activeDirectoryUser
            $schoolServiceCreds = new-object -typename System.Management.Automation.PSCredential -argumentlist $schoolServiceCreds.Username,$schoolServiceCreds.Password
            $ADUsers = Get-ADUser -Server $activeDirectoryServer -Properties employeeID -SearchBase "OU=Users,OU=Western Port Secondary College,DC=Curric,DC=Western-Port-SC,DC=wan" -Filter * -Credential $schoolServiceCreds | Sort-Object employeeID
        }#>
    }
    catch 
    {
        
    }

}
######################Import and Process Students######################

$importedStudents = $null
$workingStudents = @()

#Import Students from CSV(s) based upon settings

if ($includeDeltas -eq $true -and $modifiedHeaders -eq $false) #Only do Delta join if not using files from exporter as exporter joins the files
{
    $importedStudents = Import-CSV (Join-eduHubDelta (Join-Path -Path $fileLocation -ChildPath $importFileStudents) (Join-Path -Path $fileLocation -ChildPath $importFileStudentsDelta) "$PSScriptRoot\Cache\" "STKEY")  | Select-Object -Property  @{Name="SIS_ID";Expression={$_."STKEY"}},PREF_NAME,FIRST_NAME,SURNAME,BIRTHDATE,GENDER,@{Name="FINISH";Expression={$_."EXIT_DATE"}},HOME_GROUP,SCHOOL_YEAR,FAMILY,USERNAME,E_MAIL,CONTACT_A,STATUS,ALIAS,EXPORT | Sort-Object -property STATUS, SIS_ID
}
elseif ($includeDeltas -eq $false -and $modifiedHeaders -eq $false) #Only Run import if not using modified headers from exporter
{
    $importedStudents = Import-CSV (Join-Path -Path $fileLocation -ChildPath $importFileStudents) | Select-Object -Property  @{Name="SIS_ID";Expression={$_."STKEY"}},PREF_NAME,FIRST_NAME,SURNAME,BIRTHDATE,GENDER,@{Name="FINISH";Expression={$_."EXIT_DATE"}},HOME_GROUP,SCHOOL_YEAR,FAMILY,USERNAME,E_MAIL,CONTACT_A,STATUS,ALIAS,EXPORT | Sort-Object -property STATUS, SIS_ID
}
elseif ($modifiedHeaders -eq $true)
{
    $importedStudents = (Join-Path -Path $fileLocation -ChildPath $importFileStudents) | Select-Object -Property  SIS_ID,PREF_NAME,FIRST_NAME,SURNAME,BIRTHDATE,GENDER,FINISH,HOME_GROUP,SCHOOL_YEAR,FAMILY,USERNAME,E_MAIL,CONTACT_A,STATUS,ALIAS,EXPORT | Sort-Object -property STATUS, SIS_ID
}
else
{
    throw "Cannot Import Error with locating or processing files"
}

#Process Students

foreach ($student in $importedStudents)
{
    $tempUser = $null
    if ($null -ne ($tempUser = (Merge-User -workingUser $student -exitAfter $handlingStudentExitAfter -handlingEmail $handlingStudentEmail -handlingUsername $handlingStudentUsername -handlingAlias $handlingStudentAlias)))
    {
        $workingStudents += $tempUser
    }
}

#$importedStudents = $null #Explicitly destroy data to clear up resources

###################### Import and Process Staff ######################

$importedStaff = $null
$workingStaff = @()

#Import Staff from CSV(s) based upon settings

if ($includeDeltas -eq $true -and $modifiedHeaders -eq $false) #Only do Delta join if not using files from exporter as exporter joins the files
{
    $importedStaff = Import-CSV (Join-eduHubDelta (Join-Path -Path $fileLocation -ChildPath $importFileStaff)(Join-Path -Path $fileLocation -ChildPath $importFileStaffDelta) "$PSScriptRoot\Cache\" "SFKEY")  | Select-Object -Property @{Name="SIS_ID";Expression={$_."SFKEY"}},PREF_NAME,FIRST_NAME,SURNAME,BIRTHDATE,GENDER,FINISH,HOMEKEY,USERNAME,E_MAIL,@{Name="STATUS";Expression={$_."STAFF_STATUS"}},@{Name="SIS_EMPNO";Expression={$_."PAYROLL_REC_NO"}},ALIAS,EXPORT | Sort-Object -property STATUS, SIS_ID
}
elseif ($includeDeltas -eq $false -and $modifiedHeaders -eq $false) #Only Run import if not using modified headers from exporter
{
    $importedStaff = Import-CSV (Join-Path -Path $fileLocation -ChildPath $importFileStaff) | Select-Object -Property @{Name="SIS_ID";Expression={$_."SFKEY"}},PREF_NAME,FIRST_NAME,SURNAME,BIRTHDATE,GENDER,FINISH,HOMEKEY,USERNAME,E_MAIL,@{Name="STATUS";Expression={$_."STAFF_STATUS"}},@{Name="SIS_EMPNO";Expression={$_."PAYROLL_REC_NO"}},ALIAS,EXPORT | Sort-Object -property STATUS, SIS_ID
}
elseif ($modifiedHeaders -eq $true)
{
    $importedStaff = (Join-eduHubDelta $fileStaff $fileStaffDelta "SIS_ID") | Select-Object -Property SIS_ID,PREF_NAME,FIRST_NAME,SURNAME,BIRTHDATE,GENDER,FINISH,HOMEKEY,USERNAME,E_MAIL,STATUS,SIS_EMPNO,ALIAS,EXPORT | Sort-Object -property STATUS, SIS_ID
}
else
{
    throw "Cannot Import Error with locating or processing files"
}

#Process Staff
foreach ($staff in $importedStaff)
{
    if ($null -ne ($tempUser = (Merge-User -workingUser $staff -exitAfter $handlingStaffExitAfter -handlingEmail $handlingStaffEmail -handlingUsername $handlingStaffUsername  -handlingAlias $handlingStaffAlias -userStaff )))
    {
        $workingStaff += $tempUser
    }
}

#$importedStaff = $null #Explicitly destroy data to clear up resources

###################### Import and Process Families ######################

$importedFamilies = $null
$workingFamilies = @()

#Import Families from CSV(s) based upon settings - No modified headers here as there is no need due to their only being the one table of this type

if ($includeDeltas -eq $true) #Only do Delta join if not using files from exporter as exporter joins the files
{
    $importedFamilies = Import-CSV (Join-eduHubDelta (Join-Path -Path $fileLocation -ChildPath $importFileFamilies) (Join-Path -Path $fileLocation -ChildPath $importFileFamiliesDelta) "$PSScriptRoot\Cache\" "DFKEY") | Select-Object -Property DFKEY,E_MAIL_A,MOBILE_A,E_MAIL_B,MOBILE_B,HOMEKEY | Sort-Object -property DFKEY
}
else
{
    $importedFamilies = Import-CSV (Join-Path -Path $fileLocation -ChildPath $importFileFamilies) | Select-Object -Property DFKEY,E_MAIL_A,MOBILE_A,E_MAIL_B,MOBILE_B,HOMEKEY | Sort-Object -property DFKEY
}


#Sort families so that only families where there is an active student are kept and that are due to be exported, then with an active family check to see if primary contact is contact B (A and C are left as A), if so change the details, Contact B is dropped on export. If there use only the first record (usally the oldest student) to calculate this

foreach ($family in $importedFamilies)
{
    if ($workingStudents.FAMILY -match $family.DFKEY)
    {
        $workingFamilies += $family
        
        if ((($workingStudents | Where-Object {$_.FAMILY -eq $family.DFKEY} | Sort-Object -Property SIS_ID | select-object -First 1).CONTACT_A) -eq "B")
        {
            $family.E_MAIL_A = $family.E_MAIL_B
            $family.MOBILE_A = $family.MOBILE_B
            Write-Host "Changing Contacts for $($family.DFKEY)"
        }
        elseif (((($workingStudents | Where-Object {$_.FAMILY -eq $family.DFKEY} | Sort-Object -Property SIS_ID | select-object -First 1).CONTACT_A) -eq "C") -and -not [string]::IsNullOrWhiteSpace($family.E_MAIL_B) -and $family.E_MAIL_B -ne $family.E_MAIL_A)
        {
            $family.E_MAIL_A += ";$($family.E_MAIL_B)"
            Write-Host "Adding Secondary Email for $($family.DFKEY)"
        }
        
    }
    
}

$importedFamilies = $null #Explicitly destroy data to clear up resources

###################### Import and Process Addresses ######################

$importedAddresses = $null
$workingAddresses = @()

#Import Addresses from CSV(s) based upon settings - No modified headers here as there is no need due to their only being the one table of this type

if ($includeDeltas -eq $true) #Only do Delta join if not using files from exporter as exporter joins the files
{
    $importedAddresses = Import-CSV (Join-eduHubDelta (Join-Path -Path $fileLocation -ChildPath $importFileAddresses) (Join-Path -Path $fileLocation -ChildPath $importFileAddressesDelta) "$PSScriptRoot\Cache\" "UMKEY")  | Select-Object -Property UMKEY,ADDRESS01,ADDRESS02,ADDRESS03,STATE,POSTCODE,TELEPHONE,MOBILE | Sort-Object -property UMKEY
}
else
{
    $importedAddresses = Import-CSV (Join-Path -Path $fileLocation -ChildPath $importFileAddresses) | Select-Object -Property UMKEY,ADDRESS01,ADDRESS02,ADDRESS03,STATE,POSTCODE,TELEPHONE,MOBILE | Sort-Object -property UMKEY
}

#Sort through addresses, only keeping those where there is a family or staff member that are due to be exported associated with the address
foreach ($address in $importedAddresses)
{
    if (($workingFamilies.HOMEKEY -match $address.UMKEY) -or ($workingStaff.HOMEKEY -match $address.UMKEY))
    {
        $workingAddresses += $address
    }
    
}

$importedAddresses = $null #Explicitly destroy data to clear up resources