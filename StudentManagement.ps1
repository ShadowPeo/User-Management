## Import Config Files

Import-Module "$PSScriptRoot\Config\GenericConfig.ps1" -Force
Import-Module "$PSScriptRoot\Config\StudentConfig.ps1" -Force

##################################################
#Modules to Import
Import-Module ActiveDirectory -Force
Import-Module $PSScriptRoot/Modules/Logging.psm1 -Force
#Import-Module $PSScriptRoot/Modules/ActiveDirectory.psm1 -Force
Import-Module $PSScriptRoot/Modules/eduPASS.psm1 -Force
Import-Module $PSScriptRoot/Modules/Attributes.psm1 -Force
Import-Module $PSScriptRoot/Modules/Authentication.psm1 -Force
Import-Module $PSScriptRoot/Modules/UserFunctions.psm1 -Force
Import-Module $PSScriptRoot/Modules/TextFunctions.psm1 -Force
Import-Module $PSScriptRoot/Modules/eduHubFiles.psm1 -Force
Import-Module $PSScriptRoot/Modules/Emails.psm1 -Force
Import-Module $PSScriptRoot/Modules/SnipeitPS.psm1 -Force

######################################################

#LogFile Parameters
$LogFile = "$LogPath\$(Get-Date -UFormat '+%Y-%m-%d-%H-%M')-$(if($dryRun -eq $true){"DRYRUN-"})$([io.path]::GetFileNameWithoutExtension($MyInvocation.MyCommand.Name)).log"

##Service Account Setup

#Variable Reset
$detServiceCreds = $null
$schoolServiceCreds = $null
$userCredentails = $null

$detServiceCreds = Get-SavedCredentials_WithRequest "$PSScriptRoot\Credentials\edu002-$([Environment]::MachineName)-$([Environment]::UserName).crd" $edu002DC_User
$detServiceCreds = new-object -typename System.Management.Automation.PSCredential -argumentlist $detServiceCreds.Username,$detServiceCreds.Password

##School DC Credential Check

if ($schoolRunAsLoggedIn -eq $false)
{
    $schoolServiceCreds = Get-SavedCredentials_WithRequest "$PSScriptRoot\Credentials\schoolDC-$([Environment]::MachineName)-$([Environment]::UserName).crd" $schoolDC_User
    $schoolServiceCreds = new-object -typename System.Management.Automation.PSCredential -argumentlist $schoolServiceCreds.Username,$schoolServiceCreds.Password
    $userCredentails = $null
}

##Attribute Array Setup

#Variable Reset
$attributeFiles = $null
$customSettings = @{}
$customAuthChanges = @{}
$customAttributeNames = @()
$AttributeName = $null
$aupCheck = $null

$attributeFiles = Get-ChildItem -Path "$PSScriptRoot\Attributes" -Recurse -Filter "$attributeFilePrefix*"

Set-AttributeFiles


#Import CASES Users File
$Users = Import-Csv -Path  (Join-eduHubDelta $fileCASES $fileCASESDelta $temporaryDirectory "STKEY") 
#$Users = (Import-Csv -Path (Join-eduHubDelta $fileCASES $fileCASESDelta $temporaryDirectory "STKEY") | where-object {$_.STKEY -eq "ALE0007"})

#Import Manual Match File
$bannedUser = @(Import-Csv -Path $fileBanned)

#Email Config
$emailConfig = @(Import-Csv -Path $fileEmailConfig)

#Group Config
$groupConfig = @(Import-Csv -Path $fileGroupConfig)

#Import Ignored Accounts File
$ignoredAccounts = ""
#$ignoredAccounts = @(Import-Csv -Path $fileIgnoredAccounts)

#Import Ignored UPN PREF_NAME File
#$ignoredPREFUPN = @(Import-Csv -Path $fileIgnoredUPN)

if ($enableAUPCheck -eq $true)
{
    $aupCheck = Import-Csv -Path $fileAUP
}

Start-Log

Send-HealthCheck $healthchecksCASESStart $healthchecksEnabled $healthchecksDryRun $DryRun "Starting CASES File Processing"

# Year 07 Password Change Check

if ($Y07InitExclude -eq $true)
{
    LogWrite "Begining Year 07 Initialization Count"
    $csvList = ((Import-Csv -Path (Join-eduHubDelta $fileCASES $fileCASESDelta $temporaryDirectory "STKEY") | where-object {$_.SCHOOL_YEAR -eq "07" -and $_.STATUS -eq "ACTV"}))
    $csvCount = 0
    $adList = Get-ADUser -Server $SchoolDC -ErrorAction Stop -filter * -Properties userCASESStatus,department | where-object {$_.userCASESStatus -eq "FUT" -and $_.department -eq "Future - Year 07"} | Sort-Object samAccountName | SELECT samAccountName
    foreach ($Y07Student in $csvList)
    {
       
        if ($adList.samAccountName -contains $Y07Student.STKEY)
        {
            
            $csvCount++
        }
    }

    if (($adList.Count - $csvCount) -lt $allowedY07Changes)
    {
        LogWrite "Locking Y07 Initialization as there is more than $allowedY07Changes scheduled"
        $Y07InitializationLock = $true
    }
}

#Complete Check Loop
ForEach ($User in $Users) 
{
    #CSV File Variables
    $CASESID = $User.STKEY
    $FirstName = $User.FIRST_NAME
    $PrefName = $User.PREF_NAME
    $Surname = $User.SURNAME
    $calcUPN = "$($CASESID.ToLower())@westernportsc.vic.edu.au"
    $userCommonName = "$($User.FIRST_NAME) $($User.SURNAME.ToUpper()) ($($User.STKEY))"
    $accountCreated = $false
    $userCurrentGroups = $null
    
    #Variable Cleanup (Set to null to ensure a clean run)
    $AD_User = $null
    $DET_User = $null
    $nameChange = $false
    $commonNameChange = $false
    $changeYearLevel = $false
    $prevStatus = $null
    $changeAUP = $false

    #Write-Output "$userCommonName Processing"

    #Main User Loop
    if((($User.STATUS -eq "ACTV") -or ($User.STATUS -eq "LVNG") -or ($User.STATUS -eq "FUT")) -and ($ignoredAccounts.CASESID -notcontains $User.STKEY))
    {

        ### Change Status to ACTV if today is equal to or greater than the entry date, but only if they are marked as Future
        if ((Get-Date) -ge (Get-Date $User.ENTRY) -and $User.STATUS -eq "FUT")
        {
            $User.STATUS = "ACTV"
        }

        ### Try to retrieve AD User based upon CASES Code, if not try to create user
        if ($schoolRunAsLoggedIn -eq $true)
        {
            try 
            {
                $AD_User = Get-ADUser $CASESID -Server $SchoolDC -Properties * -ErrorAction Stop
            }
            catch
            {
                LogWrite "A User with the UserID of $CASESID does not exist, creating user"

                if ($User.STATUS -eq "FUT")
                {
                    New-ADUser -Name "$($User.FIRST_NAME) $(($User.Surname).ToUpper()) ($($User.STKEY))" -GivenName "$($User.FIRST_NAME)" -Surname "$($User.SURNAME)" -SamAccountName "$($User.STKEY)" -UserPrincipalName "$($User.STKEY)@westernportsc.vic.edu.au" -Path "$ouNew,$studentBaseOU" -AccountPassword(ConvertTo-SecureString (Get-Password -pwType 'Complex' -userBirthdate $User.BIRTHDATE) -AsPlainText -Force) -Enabled $false -Server $schoolDC
                }
                else
                {
                    New-ADUser -Name "$($User.FIRST_NAME) $(($User.Surname).ToUpper()) ($($User.STKEY))" -GivenName "$($User.FIRST_NAME)" -Surname "$($User.SURNAME)" -SamAccountName "$($User.STKEY)" -UserPrincipalName "$($User.STKEY)@westernportsc.vic.edu.au" -Path "OU=Year $($User.SCHOOL_YEAR),$studentBaseOU" -AccountPassword(ConvertTo-SecureString (Get-Password -pwType 'Complex' -userBirthdate $User.BIRTHDATE) -AsPlainText -Force) -Enabled $false -Server $schoolDC
                }
                
                $AD_User = Get-ADUser $User.STKEY -Server $SchoolDC -Properties * -ErrorAction Stop
                $accountCreated = $true
            }
        }
        else
        {
            try 
            {
                $AD_User = Get-ADUser $CASESID -Server $SchoolDC -Credential $schoolServiceCreds -Properties * -ErrorAction Stop
            }
            catch
            {
                LogWrite "A User with the UserID of $CASESID does not exist, creating user"
                
                if ($User.STATUS -eq "FUT")
                {
                    New-ADUser -Name "$($User.FIRST_NAME) $(($User.Surname).ToUpper()) ($($User.STKEY))" -GivenName "$($User.FIRST_NAME)" -Surname "$($User.SURNAME)" -SamAccountName "$($User.STKEY)" -UserPrincipalName "$($User.STKEY)@westernportsc.vic.edu.au" -Path "$ouNew,$studentBaseOU" -AccountPassword(ConvertTo-SecureString (Get-Password -pwType 'Complex' -userBirthdate $User.BIRTHDATE) -AsPlainText -Force) -Enabled $false -Server $schoolDC -Credential $schoolServiceCreds
                }
                else
                {
                    New-ADUser -Name "$($User.FIRST_NAME) $(($User.Surname).ToUpper()) ($($User.STKEY))" -GivenName "$($User.FIRST_NAME)" -Surname "$($User.SURNAME)" -SamAccountName "$($User.STKEY)" -UserPrincipalName "$($User.STKEY)@westernportsc.vic.edu.au" -Path "OU=Year $($User.SCHOOL_YEAR),$studentBaseOU" -AccountPassword(ConvertTo-SecureString (Get-Password -pwType 'Complex' -userBirthdate $User.BIRTHDATE) -AsPlainText -Force) -Enabled $false -Server $schoolDC -Credential $schoolServiceCreds
                }
                
                $AD_User = Get-ADUser $User.STKEY -Server $SchoolDC -Credential $schoolServiceCreds -Properties * -ErrorAction Stop
                $accountCreated = $true
            }
        }

        if ($accountCreated -eq $true)
        {
            Start-Sleep -Seconds 10
        }

        $userCurrentGroups = (Get-ADPrincipalGroupMembership $AD_User.SamAccountName).name

        ### Update UPN
        updateADValue "UPN" "UserPrincipalName" $calcUPN.ToLower() 


        ### Update Name
        updateADValue "Given Name" "givenName" $(correctToTitle $User.FIRST_NAME)
        
        ##Put conversion for middle initials here
        <#
        if ($AD_User.sn -ne $User.SURNAME)
        {
            $AD_User.sn = $User.SURNAME
            LogWrite "$userCommonName's Given Name set to $($User.SURNAME)"
            $nameChange = $true
        }#>
        
        ### Update Surname
        updateADValue "Surname" "sn" $(correctToTitle $User.SURNAME)

        if ($AD_User.cn -ne $userCommonName)
        {
            $commonNameChange = $true
            LogWrite "$userCommonName's Common Name set to $userCommonName"
        }
        
        ##Uncomment this section for Prefered name as Display Name
        <#
        if ((($User.FIRST_NAME -eq $User.PREF_NAME) -or ($User.PREF_NAME -eq $null) -or ($User.PREF_NAME -eq "")) -and ($AD_User.displayName -ne "$($User.FIRST_NAME) $($User.SURNAME.ToUpper())"))
        {
            $AD_User.displayName = "$($User.FIRST_NAME) $($User.SURNAME.ToUpper())"
            LogWrite "$userCommonName's Display Name set to $($User.FIRST_NAME) $($User.SURNAME.ToUpper())"
        }
        elseif ((($User.FIRST_NAME -ne $User.PREF_NAME) -and (($User.PREF_NAME -ne $null) -and ($User.PREF_NAME -ne ""))) -and ($AD_User.displayName -ne "$($User.FIRST_NAME) $($User.SURNAME.ToUpper())"))
        {
            $AD_User.displayName = "$($User.PREF_NAME) $($User.SURNAME.ToUpper())"
            LogWrite "$userCommonName's Display Name set to $($User.PREF_NAME) $($User.SURNAME.ToUpper())"
        }
        #>
        
        ##Uncomment this section for first name as Display Name
        updateADValue "Display Name" "displayName" "$($User.FIRST_NAME) $($User.SURNAME.ToUpper())"

        if ($User.E_MAIL -ne $null -and $User.E_MAIL -ne "" -and $User.E_MAIL -notlike "*$upnDomain")
        {
            updateADArray "Other Mailboxes" "otherMailbox" $User.E_MAIL.ToLower()
        }
        
        
        ## Future vs Current students

        if ($User.STATUS -eq "FUT")
        {
            ##Job Title to Future Student
            updateADValue "Title" "Title" "Future Student"
            ##SchoolYear to Department
            updateADValue "Department" "department" "Future - Year $($User.SCHOOL_YEAR)"
        }
        else
        {
            ##Job Title to Student
            updateADValue "Title" "Title" "Student"
            ##SchoolYear to Department
            updateADValue "Department" "department" "Year $($User.SCHOOL_YEAR)"
            ## Year Level Change
            if ($User.SCHOOL_YEAR -ne "Year $($User.SCHOOL_YEAR)")
            {
                $changeYearLevel = $true
            }
        }

        #Description
        if (($User.STATUS -eq "FUT") -and ($User.ENTRY -ne "" -and $User.ENTRY -ne $null) -and $AD_User.Description -ne "Starting on $(Get-Date $User.ENTRY -UFormat '+%Y-%m-%d')")
        {
            updateADValue "Description" "Description" "Starting on $(Get-Date $User.ENTRY -UFormat '+%Y-%m-%d')"
        }
        elseif ($User.STATUS -eq "ACTV" -or $User.STATUS -eq "LVNG")
        {
            updateADValue "Description" "Description" $USER.HOME_GROUP
        }

        ##Office to Homegroup
        updateADValue "Office" "physicalDeliveryOfficeName" $User.HOME_GROUP
        
        ##Add/Update School House
        updateADValue "House" "schoolHouse" $(correctToTitle $User.HOUSE)
      
        ##Employee Type to Staff or Student
        updateADValue "Employee Type" "employeeType" "Student"

        ##CASES Status
        if ($AD_User.userCASESStatus -ne $User.STATUS)
        {
            $prevStatus = $AD_User.userCASESStatus
            updateADValue "CASES Status" "userCASESStatus" $User.STATUS
        }

        ##Employee Number (CASESID)
        updateADValue "Employee Number" "EmployeeNumber" $User.STKEY

        ##Employee ID (CASESID)
        updateADValue "Employee ID" "EmployeeID" $User.STKEY.ToUpper()

        ##Email Address
        updateADValue "Email Address" "mail" $calcUPN

        ##Secondary Email
        $tempTo=$calcUPN.Split("@")[0]
        updateADArray "proxyAddresses" "proxyAddresses" "smtp:$tempTo@$secondaryDomain"

        ### Ensure that the UPN/Primary Email is in the SMTP list as primary
        if ($AD_User.proxyAddresses -notcontains "SMTP:$calcUPN")
        {
            updateADArray "proxyAddresses" "proxyAddresses" "SMTP:$calcUPN"
        }

        ### Find Address that need to be changed, add them to a temporary array
        foreach ($address in $AD_User.proxyAddresses)
        {
            if($address -ne "SMTP:$calcUPN" -and $address.Substring(0,4) -ceq "SMTP")
            {
                LogWrite "$userCommonName - Found Non-Primary email ($address) set as primary, correcting" -foregroundColour "Magenta"
                $changedProxyAddresses += $address
            }
        }

        ##Enable Student if Required
        if (($AD_User.Enabled -ne $true) -and (($bannedUser -notcontains $AD_User.SamAccountName) -and ($User.STATUS -ne "FUT")))
        {
            $AD_User.Enabled = $true
            LogWrite "$userCommonName's account inactive, making active"
            UpdateADArray "Inactive Date" "userInactive" "SETNULL"
        }

        ## Change AD User Based Upon DET Use

        ### Search Users OtherMailbox field for a valid DET email address
                        
        foreach ($otherEmail in $AD_User.otherMailbox)
        {
            if ($otherEmail -like "*schools.vic.edu.au")
            {
                ##### Retrieve user by DET (@schools.vic.edu.au) email id already known
                try 
                {
                    $DET_User = Get-edu002_User_ByEmail $SchoolID $edu002DC $detServiceCreds $otherEmail
                }
                catch
                {
                    LogWrite "A DET User with the UserID of $CASESID does not exist with email $otherEmail"
                    $DET_User = $null
                    continue
                }
                break;
            }
        }
        
        ### Retrieve DET User
        if ($DET_User -eq $null)
        {
            ##### Try to retrieve user based upon Display Name and scope limiting
            try 
            {
                $DET_User = Get-edu002_User_ByName $SchoolID $edu002DC $detServiceCreds $AD_User.DisplayName
            }
            catch
            {
                LogWrite "A DET User with the UserID of $CASESID does not exist with a name fof $($AD_User.DisplayName) within the schools group structure"
                $DET_User = $null
                continue
            }
        }
        ### Change Values
        if ($DET_User -ne $null)
        {        
            #### Company to Match DET Office
            updateADValue "Company" "Company" $(correctToTitle $DET_User.Office)
        
            #### Process changes to email if a valid email was found
            updateADArray "Other Mailboxes" "otherMailbox" $DET_User.mail.ToLower()
        }

        ## Licencing ACLs - Uses Attributes Module
        Set-Licencing

        #Check against AUP if the check is enabled
        if ($enableAUPCheck = $true)
        {
            $tempAUP = $null
            $tempAUP = $aupCheck | WHERE CASES -eq $CASESID
                
            if ($tempAUP.returned -eq "Yes" -and $AD_User.userCASESStatus-ne "FUT")
            {

                if ($AD_User.userAUPStatus -ne "Returned")
                {
                    $changeAUP = $true
                }
                
                #Change AUP status to returned
                updateADValue "AUP Status" "userAUPStatus" "Returned"

                #If Member of no AUP group remove
                    
                if ($userCurrentGroups -contains $groupNoAUP)
                {
                    Remove-ADGroupMember -Identity $groupNoAUP -Members $AD_User -Confirm:$false
                    LogWrite "$userCommonName's has been removed from $groupNoAUP"
                }
                
            }
            elseif (($AD_User.userCASESStatus -eq "FUT" -and $tempAUP.returned -ne "Yes") -and $userCurrentGroups -contains $groupNoAUP)
            {
                #Ensure not a member of no AUP as they are a future student
                Remove-ADGroupMember -Identity $groupNoAUP -Members $AD_User -Confirm:$false
            }
            elseif ($AD_User.userCASESStatus -ne "FUT")
            {
                #Change AUP status to not returned
                updateADValue "AUP Status" "userAUPStatus" "Unreturned"
                
                #If not Member of no AUP group add
                
                if ($userCurrentGroups -notcontains $groupNoAUP)
                {
                    Add-ADGroupMember -Identity $groupNoAUP -Members $AD_User.SamAccountName
                    LogWrite "$userCommonName's has been added to $groupNoAUP"
                }
                
            }
        }

        #Write User back to AD if not a dry run
        if ($DryRun -eq $false)
        {
            Set-ADUser -Instance $AD_User

            ### Set Password if the account is required to be newly active

            if ($accountCreated -eq $false -and $AD_User.userCasesStatus -eq "FUT" -and (($User.STATUS -eq "ACTV" -or (NEW-TIMESPAN –Start ((Get-Date $User.ENTRY).AddDays($activeBefore-(2*$activeBefore))) –End (Get-Date)).Days -ge 0 ) -and (([datetime]::fromfiletime($AD_User.pwdLastSet)) -ge ((Get-Date $User.ENTRY).AddDays($activeBefore-(2*$activeBefore))))))  #- Timespan less than +2 (less than 48 hours in the future) or already past)
            {
                Write-Host "Change User Password and send Welcome Email - IMPLEMENT ME"
                Initialize-User
                
            }
            elseif($accountCreated -eq $true -and $User.STATUS -eq "ACTV")
            {
                Initialize-User
            }
            elseif($changeAUP -eq $true)
            {
                Initialize-User
            }
            elseif ($User.STATUS -eq "ACTV" -or $User.STATUS -eq "LVNG")
            {
                Set-StudentGroups
            }
            else
            {
                Write-Output "$($AD_User.samAccountName) - Future Student - Do Nothing"
            }


            #Rename the user if Common Name has changed and not a dry run
            if ($commonNameChange -eq $true)
            {
                #Uncomment this to use First Name for Common Name
                Rename-ADObject -Identity $AD_User -NewName "$($User.FIRST_NAME) $($User.SURNAME) ($($User.STKEY))"
            }
            
            ##Move user to Correct OU    
            if ($AD_User.DistinguishedName -ne "CN=$($AD_User.Name),OU=Year $($User.SCHOOL_YEAR),$studentBaseOU" -and $User.STATUS -ne "FUT")
            {
                Move-ADObject $AD_User.DistinguishedName -TargetPath "OU=Year $($User.SCHOOL_YEAR),$studentBaseOU"
                LogWrite "$userCommonName's - Moving to new OU - OU=Year $($User.SCHOOL_YEAR),$studentBaseOU"
            }
            elseif ($AD_User.DistinguishedName -ne "CN=$($AD_User.Name),$ouNew,$studentBaseOU" -and $User.STATUS -eq "FUT")
            {
                Move-ADObject $AD_User.DistinguishedName -TargetPath "$ouNew,$studentBaseOU"
                LogWrite "$userCommonName's - Moving to new OU - $ouNew,$studentBaseOU"
            }
        }
        elseif ($DryRun -eq $true)
        {
            #pause
        }
    }
    elseif ($User.STATUS -eq "LEFT" -or $User.STATUS -eq "INAC")
    {
        
        ### Set Variable to remove user to false, to trigger removal of user, set to true programmatically
        $removeUser = $false

        try 
        {
            if ($schoolRunAsLoggedIn -eq $true)
            {
                $AD_User = Get-ADUser $CASESID -Server $SchoolDC -Properties * -ErrorAction Stop
            }
            else
            {
                $AD_User = Get-ADUser $CASESID -Server $SchoolDC -Credential $schoolServiceCreds -Properties * -ErrorAction Stop
            }
        }
        catch
        {
            continue #Move to next iteration if no user exists
        }

        ### Disable user as they should not be needed or dodd to Inactive users group to Deny Logon only on local machines?
        if ( $AD_User.enabled -eq $true)
        {
            $AD_User.enabled = $false
            LogWrite "$userCommonName's - Account is being disabled"

        }

        ### Job Title to Student
        updateADValue "Title" "Title" "Exiting Student"

        ### CASES Status
        updateADValue "CASES Status" "userCASESStatus" $User.STATUS

        ### Set Inactive date to today if not already set
        if (($AD_User.userInactive.Count -eq $null -or $AD_User.userInactive.Count -eq "") -or (($User.EXIT_DATE -ne $null -and $User.EXIT_DATE -ne "") -and $AD_User.userInactive -ne (Get-Date $User.EXIT_DATE).ToFileTime()))
        {
            updateADValue "Inactive Date" "userInactive" (Get-Date $User.EXIT_DATE).ToFileTime()
        }
        elseif ($AD_User.userInactive.Count -eq $null -or $AD_User.userInactive.Count -eq "")
        {
            updateADValue "Inactive Date" "userInactive" (Get-Date).ToFileTime().ToString()
        }

        ## Set Licencing
        if ($AD_User.UserPrincipalName -like "*@westernportsc.vic.edu.au")
        {
            Set-Licencing
        }

        #Package and Remove home fodler if exists
        if ($AD_User.HomeDirectory -ne $null)
        {
            if (Remove-HomeFolder $AD_User.HomeDirectory $directoryHomeDriveArchive $AD_User.SamAccountName)
            {
                $AD_User.HomeDirectory = $null
                $AD_User.HomeDrive = $null
            }
        }


        #SEND EMAIL

        #Description
        if (($User.EXIT_DATE -ne "" -and $User.EXIT_DATE -ne $null) -and $AD_User.Description -ne "Marked as Exiting on $(Get-Date ([datetime]::fromfiletime($AD_User.userInactive)) -UFormat '+%Y-%m-%d')")
        {
            updateADValue "Description" "Description" "Exited on $(Get-Date ([datetime]::fromfiletime($AD_User.userInactive)) -UFormat '+%Y-%m-%d')"
        }
        elseif ($AD_User.Description -notlike "Marked as Exiting*")
        {
            updateADValue "Description" "Description" "Marked as Exiting on $(Get-Date -UFormat '+%Y-%m-%d')"
        }
        
        ###Switch to do days after made inactive tasks
        switch ((NEW-TIMESPAN –Start $([datetime]::fromfiletime($AD_User.userInactive)) –End (Get-Date)).Days)
        {
            {$_ -gt $exitedAfter}
            {
                $removeUser = $true   
            }
        }
      
        #Write User back to AD if not a dry run
        if ($DryRun -eq $false)
        {
            #Strip Groups
            foreach($group in $AD_User.MemberOf)
            {

                if (($group).Substring(3,($aclDenyLocalLogon.Length)) -ne $aclDenyLocalLogon)
                {
                    Remove-ADGroupMember -Identity $group -Members $AD_User -Confirm:$false
                    LogWrite "$userCommonName's has been removed from $group"
                }
            }
            
            Set-ADUser -Instance $AD_User
            
            ##Move user to Correct OU
            if ($AD_User.DistinguishedName -ne "CN=$($AD_User.Name),$ouExiting,$studentBaseOU")
            {
                Move-ADObject $AD_User.DistinguishedName -TargetPath "$ouExiting,$studentBaseOU"
                LogWrite "$userCommonName's account is being moved to the Exiting OU"
            }

           ### Add to Deny Local Logon Group if less than number of exited days
            if ((Get-ADGroupMember -Identity $aclDenyLocalLogon | Select -ExpandProperty Name) -notcontains $AD_User.Name)
            {
                Add-ADGroupMember -Identity $aclDenyLocalLogon -Members $AD_User -Confirm:$false
                LogWrite "$userCommonName's has been added to $aclDenyLocalLogon"
            }

            ## Remove User
            if ($removeUser -eq $true)
            {
                Remove-ADUser -Identity ($AD_User.samAccountName) -Confirm:$false -Server $schoolDC
                LogWrite "$userCommonName's account is being removed"
            }

        }

    }
}
Send-HealthCheck $healthchecksCASESComplete $healthchecksEnabled $healthchecksDryRun $DryRun "Completed CASES File Processing"

#Clean up Temporary Directory
if(test-path $temporaryDirectory)
{
    LogWrite "Removing Temporary Directory and Files" "Verbose"

    Get-ChildItem -Path $temporaryDirectory -Recurse | Remove-Item -force -recurse
    #Remove-Item $temporaryDirectory -Force 
}