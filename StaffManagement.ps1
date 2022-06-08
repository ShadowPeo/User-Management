
Import-Module "$PSScriptRoot\Config\GenericConfig.ps1" -Force
Import-Module "$PSScriptRoot\Config\StaffConfig.ps1" -Force


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
Import-Module $PSScriptRoot/Modules/SnipeitPS.psm1 -Force

######################################################
#Local Functions for testing


function Exit-StaffUser 
{
        ### CASES Status
        updateADValue "CASES Status" "userCASESStatus" $User.STAFF_STATUS

        ### Do certain tasks based upon the current status
        switch ($User.STAFF_STATUS)
        {
            "LEFT"
            {
                #### Update Staff Title
                updateADValue "Title" "Title" "Exited Staff"
                
                #### Set Inactive date to today if not already set in CASES
                if (($AD_User.userInactive -ne  (Get-Date $User.FINISH).ToFileTime()) -and ($User.FINISH -ne $null -and $User.FINISH -ne ""))
                {
                    updateADValue "Inactive Date" "userInactive" (Get-Date $User.FINISH).ToFileTime()
                }
                elseif (($AD_User.userInactive.Count -eq $null -or $AD_User.userInactive.Count -eq "") -and ($User.FINISH -eq $null -or $User.FINISH -eq ""))
                {
                    updateADValue "Inactive Date" "userInactive" (Get-Date (Get-Date -Format "dd/MM/yyyy")).ToFileTime()
                }

                updateADValue "Description" "Description" "Staff exited on $(Get-Date ([datetime]::FromFileTime($AD_User.userInactive)) -UFormat '%Y-%m-%d')"

            }

            "INAC"
            {
                #### Update Staff Title
                updateADValue "Title" "Title" "Inactive Staff"

                #### Set Inactive date to today if not already set
                if (($AD_User.userInactive -eq $null -or ($AD_User.userInactive -eq "" ) -and ($User.FINISH -eq $null -or $User.FINISH -eq "")))
                {
                    updateADValue "Inactive Date" "userInactive" (Get-Date (Get-Date -Format "dd/MM/yyyy")).ToFileTime()
                }
                updateADValue "Description" "Description" "Staff made inactive on $(Get-Date ([datetime]::FromFileTime($AD_User.userInactive)) -UFormat '%Y-%m-%d')"
            }

            "INACLOGIN"
            {
                #### Update Staff Title
                updateADValue "Title" "Title" "Inactive Staff - Due to No Login"

                #### Set Inactive date to today if not already set
                if (($AD_User.userInactive -eq $null) -or ($AD_User.userInactive -eq "" ) -and ($User.FINISH -eq $null -or $User.FINISH -eq ""))
                {
                    updateADValue "Inactive Date" "userInactive" (Get-Date (Get-Date -Format "dd/MM/yyyy")).ToFileTime()
                }
                updateADValue "Description" "Description" "User Disabled for account inactivity ($(Get-Date ([datetime]::FromFileTime($AD_User.userInactive)) -UFormat '%Y-%m-%d'))" 
            }
        }
        
        ### Update licencing only if they were here when we implemented AzureAD

        if ($AD_User.UserPrincipalName -like "*@$upnDomain")
        {
            Set-Licencing
        }


        ###Switch to do days after made inactive tasks
        switch ((NEW-TIMESPAN –Start $([datetime]::fromfiletime($AD_User.userInactive.Value)) –End (Get-Date)).Days)
        {
            "77"
            {
                #Move to Exited OU and Disable
            }
        }

        if (((NEW-TIMESPAN –Start $([datetime]::fromfiletime($AD_User.userInactive)) –End (Get-Date)).Days) -le $exitedAfter)
        {
            #### Set OU to Inactive OU
            $targetDN="$ouInactive,$staffBaseOU"
        }
        else
        {
            #### Allow Destructive Tasks to run
            $disableUser = $true

            #### Set OU to the Exited OU
            $targetDN= "$ouExited,$staffBaseOU"
        }

        ## SEND EMAIL


        ## Run Destructive Tasks
        if ($DryRun -eq $false)
        {

            #Remove all compass access as soon as they are marked as left
            if ($User.STAFF_STATUS -eq "LEFT")
            {
                foreach($group in $AD_User.MemberOf)
                {
                    
                    if (($group.Substring(3,7)) -eq "Compass")
                    {
                        Remove-ADGroupMember -Identity $group -Members $AD_User -Confirm:$false
                        LogWrite "$userCommonName's has been removed from $group"
                    }
                }
            }
            
            if ($disableUser -eq $true)
            {
                #Disable Account as it is now longer than required - Resets password as well due to Compass not respecting the enabled status
                if ($AD_User.enabled -eq $true)
                {
                    #### Disable Account
                    $AD_User.enabled = $false
                    LogWrite "$userCommonName's Account Disabled"

                }
            
                #Strip Groups
                foreach($group in $AD_User.MemberOf)
                {
                    Remove-ADGroupMember -Identity $group -Members $AD_User -Confirm:$false
                    LogWrite "$userCommonName's has been removed from $group"
                }

                #Package and Remove home fodler if exists and they are not inactive due to no login
                if ($AD_User.HomeDirectory -ne $null -and $User.STAFF_STATUS -ne "INACLOGIN")
                {
                    if (Remove-HomeFolder $AD_User.HomeDirectory $directoryHomeDriveArchive $AD_User.SamAccountName)
                    {
                        $AD_User.HomeDirectory = $null
                        $AD_User.HomeDrive = $null
                    }
                }
            }
            else
            {
                ### Add to Deny Local Logon Group if less than number of exited days
                if ((Get-ADGroupMember -Identity $aclDenyLocalLogon | Select -ExpandProperty Name) -notcontains $AD_User.Name)
                {
                    Add-ADGroupMember -Identity $aclDenyLocalLogon -Members $AD_User -Confirm:$false
                    LogWrite "$userCommonName's has been added to $aclDenyLocalLogon"
                }
            }

            Set-ADUser -Instance $AD_User

            ### Change Password
            if ($AD_User.pwdLastSet -lt $AD_User.userInactive -and $disableUser -eq $true)
            {
                LogWrite "$userCommonName's Password has been reset for security purposes"
                if ($User.BIRTHDATE -eq "" -or $User.BIRTHDATE -eq $null)
                {
                    $User.BIRTHDATE = Get-Date -UFormat '%d/%m/%Y'
                }
                Set-ADAccountPassword -Identity $AD_User.DistinguishedName -Reset -NewPassword (ConvertTo-SecureString -AsPlainText (Get-Password $User.BIRTHDATE "Complex") -Force)
            }
            
            ### Move the OU the account is in if required
            if ($AD_User.DistinguishedName -ne "CN=$($AD_User.Name),$targetDN" -and $targetDN -ne $null)
            {
                LogWrite "$userCommonName's account is being moved to $targetDN"
                Move-ADObject $AD_User.DistinguishedName -TargetPath $targetDN
            }
        }
        else
        {
            if ($AD_User.DistinguishedName -ne "CN=$($AD_User.Name),$targetDN" -and $targetDN -ne $null)
            {
                LogWrite "$userCommonName's account is being moved to $targetDN"
            }
        }

        if ($DryRunPause -eq $true)
        {
            pause
        }
}

function Get-userCommonName-Staff
{
    
    if ($User.PAYROLL_REC_NO -eq "" -or $User.PAYROLL_REC_NO -eq $null)
    {
        if ($usePrefName -eq $false -or $usePrefCommonName -eq $false)
        {
            return "$($User.FIRST_NAME) $($User.SURNAME.ToUpper()) ($($User.SFKEY.ToUpper()))"
        }
        else
        {
            return "$($User.PREF_NAME) $($User.SURNAME.ToUpper()) ($($User.SFKEY.ToUpper()))"
        }
    }
    else
    {
        if ($usePrefName -eq $false -or $usePrefCommonName -eq $false)
        {
            return "$($User.FIRST_NAME) $($User.SURNAME.ToUpper()) ($($User.PAYROLL_REC_NO))"
        }
        else
        {
            return "$($User.PREF_NAME) $($User.SURNAME.ToUpper()) ($($User.PAYROLL_REC_NO))"
        }
    }
}

#####################################################

#LogFile Parameters
$LogFile = "$LogPath\$(Get-Date -UFormat '+%Y-%m-%d-%H-%M')-$(if($dryRun -eq $true){"DRYRUN-"})$([io.path]::GetFileNameWithoutExtension($MyInvocation.MyCommand.Name)).log"

##Service Account Setup

#Variable Reset
$detServiceCreds = $null
$schoolServiceCreds = $null
$userCredentails = $null

$detServiceCreds = Get-SavedCredentials_WithRequest "$PSScriptRoot\Credentials\edu001-$([Environment]::MachineName)-$([Environment]::UserName).crd" $edu001DC_User
$detServiceCreds = new-object -typename System.Management.Automation.PSCredential -argumentlist $detServiceCreds.Username,$detServiceCreds.Password
$userCredentails = $null

##School DC Credential Check

if ($schoolRunAsLoggedIn -eq $false)
{
    $schoolServiceCreds = Get-SavedCredentials_WithRequest "$PSScriptRoot\Credentials\schoolDC-$([Environment]::MachineName)-$([Environment]::UserName).crd" $schoolDC_User
    $schoolServiceCreds=new-object -typename System.Management.Automation.PSCredential -argumentlist $schoolServiceCreds.Username,$schoolServiceCreds.Password
    $userCredentails = $null
}


#Make sure $LogPath exists
If(!(test-path $LogPath))
{
      New-Item -ItemType Directory -Force -Path $LogPath
}

##Attribute Array Setup

#Variable Reset
$attributeFiles = $null
$customSettings = @{}
$customAuthChanges = @{}
$customAttributeNames = @()
$AttributeName = $null

$attributeFiles = Get-ChildItem -Path "$PSScriptRoot\Attributes" -Recurse -Filter "$attributeFilePrefix*"

Set-AttributeFiles

#Import CASES Users File
$Users = Import-Csv -Path $fileCASES
#$Users = (Import-Csv -Path (Join-eduHubDelta $fileCASES $fileCASESDelta $temporaryDirectory "SFKEY") | where-object {$_.PAYROLL_REC_NO -eq "09100590"})
#$Users = Import-Csv -Path $fileCASES | where-object {$_.SFKEY -eq "RUS"}

#Import Ignored Accounts File
$ignoredAccounts = Import-Csv -Path $fileIgnoredAccounts

#Import Ignored UPN PREF_NAME File
$ignoredPREFUPN = Import-Csv -Path $fileIgnoredUPN

#Import Ignored UPN PREF_NAME File
$ignoredPrefDisplay = Import-Csv -Path $fileIgnoredDisplay

#Import Ignored Inactive Users File
$ignoredInactiveUser = Import-Csv -Path $fileIgnoredInactive

LogWrite "*****Processing of CASES Data Starting*****"

Send-HealthCheck $healthchecksCASESStart $healthchecksEnabled $healthchecksDryRun $DryRun "Starting CASES File Processing"

#Complete Check Loop
ForEach ($User in $Users) 
{
    #Variable Cleanup (Set to null to ensure a clean run)
    $AD_User = $null
    $DET_User = $null
    $calcUPN = $null
    $T0NUM = $null
    $nameChange = $false
    $commonNameChange = $false
    $inactiveUser = $null
    $usePrefName = $false

    #Build Log ID Reference
    $userCommonName = $null
    
    #Decide whether PrefName is valid, set to true if it is
    if (((($User.FIRST_NAME -ne $User.PREF_NAME) -and ($ignoredPREFUPN.CASESID -notcontains $User.SFKEY)) -and (($User.PREF_NAME -ne $null) -and ($User.PREF_NAME -ne ""))))
    {
        $usePrefName = $true
    }

    #Manually Correct T0Number (PAYROLL_REC_NUMBER)

        $userCommonName = Get-userCommonName-Staff
        LogWrite "$userCommonName Processing" -logLevel "VERBOSE"

    #Manually Set inactive if not logged in more than $inactiveAfter days
    if (($ignoredInactiveUser.USERID -notcontains $User.SFKEY -and $ignoredInactiveUser.USERID -notcontains $User.PAYROLL_REC_NO ) -and ($User.STAFF_STATUS -ne "INAC" -and $User.STAFF_STATUS -ne "LEFT"))
    {
        try
        {
            $time = (Get-Date).Adddays(-($inactiveAfter)).ToFileTime()
            $TONUM = $User.PAYROLL_REC_NO
            $inactiveUser = Get-ADUser -Filter {(LastLogonTimeStamp -lt $time) -and (samAccountName -eq $TONUM )}
            if ($inactiveUser -ne $null)
            {
                $User.STAFF_STATUS = "INACLOGIN"

                if ($inactiveUser.Enabled -eq $true)
                {
                    LogWrite "$userCommonName is Inactive for more than $inactiveAfter days, disabling"
                }
            }

        }
        catch
        {

        }
    }
    
    #Main User Loop
    if(($User.STAFF_STATUS -eq "ACTV") -and ($ignoredAccounts.CASESID -notcontains $User.SFKEY))# -and (($User.PAYROLL_REC_NO -ne $null) -and ($User.PAYROLL_REC_NO -ne "")))
    {
        #Variable Cleanup (Set to null to ensure a clean run)
        $nameChange = $false
        $commonNameChange = $false
        $userInactive = $false
        $accountCreated = $false

        #Calculated Variables
        $CASESID = $User.SFKEY
        $T0NUM = $User.PAYROLL_REC_NO

        #AD Variables
        if ($schoolRunAsLoggedIn -eq $true)
            {
                try 
                {
                    $AD_User = Get-ADUser $T0NUM -Server $SchoolDC -Properties * -ErrorAction Stop
                }
                catch
                {
            
                    LogWrite "$userCommonName does not have account utilising T0 Number does not exist, trying CASES ID" "Verbose"
            
                    ### Try SFKEY to see if its a local account
                    try
                    {
                        $AD_User = Get-ADUser ($User.SFKEY) -Server $SchoolDC -Properties * -ErrorAction Stop
                        LogWrite "$userCommonName account found as a local account using their CASES ID" "Verbose"

                        #Correct samAccountName if needed
                        if (($TONUM -ne $null -and $TONUM -ne "") -and $User.SFKEY -ne $User.PAYROLL_REC_NO)
                        {
                            LogWrite "$userCommonName now has a T0 Number (PAYROLL_REC_NO) renamiging the account to use it" -foregroundColour "Red"
                            Set-ADUser $AD_User -Replace @{samaccountname = $User.PAYROLL_REC_NO}
                            $AD_User = Get-ADUser $User.PAYROLL_REC_NO -Server $SchoolDC -Properties * -ErrorAction Stop
                        }
                    }
                    catch
                    {
                        #Create user utilising the correct ID

                        LogWrite "No existing user found with either T0 number or CASES ID, Creating User" "Verbose"

                        if ($T0NUM -ne $null -and $T0NUM -ne "")
                        {
                            LogWrite "$userCommonName user account being created utilising T0 number"
                            New-ADUser -Name "$($User.FIRST_NAME) $(($User.SURNAME).ToUpper()) ($($User.PAYROLL_REC_NO))" -GivenName "$($User.FIRST_NAME)" -Surname "$($User.SURNAME)" -SamAccountName "$($User.PAYROLL_REC_NO)" -UserPrincipalName "$($User.PAYROLL_REC_NO)@westernportsc.vic.edu.au" -Path "$ouNew,$staffBaseOU" -AccountPassword(ConvertTo-SecureString (Get-Password -pwType 'Complex' -userBirthdate $User.BIRTHDATE) -AsPlainText -Force) -Enabled $false -Server $schoolDC
                            $AD_User = Get-ADUser $T0NUM -Server $SchoolDC -Properties * -ErrorAction Stop
                            $accountCreated = $true
                        }
                        else
                        {
                            LogWrite "$userCommonName user account being created utilising CASES ID"
                            New-ADUser -Name "$($User.FIRST_NAME) $(($User.SURNAME).ToUpper()) ($($User.SFKEY))" -GivenName "$($User.FIRST_NAME)" -Surname "$($User.SURNAME)" -SamAccountName "$($User.SFKEY)" -UserPrincipalName "$($User.SFKEY)@westernportsc.vic.edu.au" -Path "OU=$ouNew,$staffBaseOU" -AccountPassword(ConvertTo-SecureString (Get-Password -pwType 'Complex' -userBirthdate $User.BIRTHDATE) -AsPlainText -Force) -Enabled $false -Server $schoolDC
                            $AD_User = Get-ADUser ($User.SFKEY) -Server $SchoolDC -Properties * -ErrorAction Stop
                            $accountCreated = $true
                        }
                    }
                }
            }
            else
            {
                try 
                {
                    $AD_User = Get-ADUser $T0NUM -Server $SchoolDC -Credential $schoolServiceCreds -Properties * -ErrorAction Stop
                }
                catch
                {
            
                    LogWrite "$userCommonName does not have account utilising T0 Number does not exist, trying CASES ID" "Verbose"
            
                    ### Try SFKEY to see if its a local account
                    try
                    {
                        $AD_User = Get-ADUser ($User.SFKEY) -Server $SchoolDC -Credential $schoolServiceCreds -Properties * -ErrorAction Stop
                        LogWrite "$userCommonName account found as a local account using their CASES ID" "Verbose"

                        #Correct samAccountName if needed
                        if (($TONUM -ne $null -and $TONUM -ne "") -and $User.SFKEY -ne $User.PAYROLL_REC_NO)
                        {
                            LogWrite "$userCommonName now has a T0 Number (PAYROLL_REC_NO) renamiging the account to use it" -foregroundColour "Red"
                            Set-ADUser $AD_User -Replace @{samaccountname = $User.PAYROLL_REC_NO} -Credential $schoolServiceCreds
                            $AD_User = Get-ADUser $User.PAYROLL_REC_NO -Server $SchoolDC -Credential $schoolServiceCreds -Properties * -ErrorAction Stop
                        }

                    }
                    catch
                    {
                        #Create user utilising the correct ID
                        LogWrite "No existing user found with either T0 number or CASES ID, Creating User" "Verbose"

                        if ($T0NUM -ne $null -and $T0NUM -ne "")
                        {
                            LogWrite "$userCommonName user account being created utilising T0 number"
                            New-ADUser -Name "$($User.FIRST_NAME) $(($User.SURNAME).ToUpper()) ($($User.PAYROLL_REC_NO))" -GivenName "$($User.FIRST_NAME)" -Surname "$($User.SURNAME)" -SamAccountName "$($User.PAYROLL_REC_NO)" -UserPrincipalName "$($User.PAYROLL_REC_NO)@westernportsc.vic.edu.au" -Path "OU=$ouNew,$staffBaseOU" -AccountPassword(ConvertTo-SecureString (Get-Password -pwType 'Complex' -userBirthdate $User.BIRTHDATE) -AsPlainText -Force) -Enabled $false -Server $schoolDC -Credential $schoolServiceCreds
                            $AD_User = Get-ADUser $T0NUM -Server $SchoolDC -Credential $schoolServiceCreds -Properties * -ErrorAction Stop
                            $accountCreated = $true
                        }
                        else
                        {
                            LogWrite "$userCommonName user account being created utilising CASES ID"
                            New-ADUser -Name "$($User.FIRST_NAME) $(($User.SURNAME).ToUpper()) ($($User.SFKEY))" -GivenName "$($User.FIRST_NAME)" -Surname "$($User.SURNAME)" -SamAccountName "$($User.SFKEY)" -UserPrincipalName "$($User.SFKEY)@westernportsc.vic.edu.au" -Path "OU=$ouNew,$staffBaseOU" -AccountPassword(ConvertTo-SecureString (Get-Password -pwType 'Complex' -userBirthdate $User.BIRTHDATE) -AsPlainText -Force) -Enabled $false -Server $schoolDC -Credential $schoolServiceCreds
                            $AD_User = Get-ADUser ($User.SFKEY) -Server $SchoolDC -Credential $schoolServiceCreds -Properties * -ErrorAction Stop
                            $accountCreated = $true
                        }
                    }
                }
            }
        
        #DET User Variables
        try 
        {
            $DET_User = Get-ADUser $T0NUM -Server $edu001DC -Credential $detServiceCreds -Properties * -ErrorAction Stop
        }
        catch
        {
            LogWrite "A User with the UserID of $TONUM does not exist"
            continue
        }
        

        ##Caclulated and Reference Variables

        if ($usePrefName -eq $false)
        {        
            $calcUPN =  Get-validUPN $User.FIRST_NAME $User.SURNAME $User.PAYROLL_REC_NO $upnDomain
            $tempTo=$calcUPN.Split("@")[0]
            updateADArray "proxyAddresses" "proxyAddresses" "smtp:$tempTo@$secondaryDomain"
        }
        else
        {
            LogWrite "$userCommonName's UPN Changed to Prefered Name" -logLevel "Verbose"
            $tempFirstNameUPN = Get-validUPN $User.FIRST_NAME $User.SURNAME $User.PAYROLL_REC_NO $upnDomain
            updateADArray "proxyAddresses" "proxyAddresses" "smtp:$tempFirstNameUPN"
            $tempTo=$tempFirstNameUPN.Split("@")[0]
            updateADArray "proxyAddresses" "proxyAddresses" "smtp:$tempTo@$secondaryDomain"
            $calcUPN = Get-validUPN $User.PREF_NAME $User.SURNAME $User.PAYROLL_REC_NO $upnDomain
            $tempTo=$calcUPN.Split("@")[0]
            updateADArray "proxyAddresses" "proxyAddresses" "smtp:$tempTo@$secondaryDomain"
        }
        
        #Correct UPN
        if ($calcUPN -ne $AD_User.UserPrincipalName)
        {
            updateADValue "UPN" "UserPrincipalName" $calcUPN.ToLower()
            updateADArray "proxyAddresses" "proxyAddresses" "SMTP:$calcUPN"
        }

        ### Clean Proxy addresses to allow for correct Primary Allocation
        $changedProxyAddresses = @()
        
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

        #Correct Name
        if ($AD_User.givenName -ne $User.FIRST_NAME)
        {
            updateADValue "Given Name" "givenName" $(correctToTitle $User.FIRST_NAME)
            $nameChange = $true
        }

        if ($AD_User.sn -ne $User.SURNAME)
        {
            updateADValue "Surname" "sn" $(correctToTitle $User.SURNAME)
            $nameChange = $true
        }

        if ($AD_User.cn -ne $userCommonName)
        {
            $commonNameChange = $true
            LogWrite "$userCommonName's Common Name set to $userCommonName"
        }
        
        ##Uncomment this section for Prefered name as Display Name
        if ($usePrefDisplayName -eq $true -and $usePrefName -eq $true)
        {
            updateADValue "Display Name" "displayName" "$($User.PREF_NAME) $($User.SURNAME.ToUpper())"
        }
        else
        {
            updateADValue "Display Name" "displayName" "$($User.FIRST_NAME) $($User.SURNAME.ToUpper())"
        }


        #Process changes to email if a valid email was found
        updateADArray "Other Mailboxes" "otherMailbox" $DET_User.mail.ToLower()
        
        if ($User.E_MAIL -ne $null -and $User.E_MAIL -ne "")
        {
            updateADArray "Other Mailboxes" "otherMailbox" $User.E_MAIL.ToLower()
        }
        
        if ($DET_User.Office -eq $schoolName)
        {
            ##Company to Match DET Office
            updateADValue "Company" "Company" $(correctToTitle $DET_User.Office)

            ##Job Title to Match DET
            updateADValue "Title" "Title" $(correctToTitle $DET_User.Title)
        }

        ##Add/Update School House
        updateADValue "House" "schoolHouse" $(correctToTitle $User.HOUSE)
      
        ##Employee Type to Staff
        updateADValue "Employee Type" "employeeType" "Staff"

        ### CASES Status

        if ($User.STAFF_STATUS -eq "ACTV" -and ($AD_User.userCASESStatus -ne "ACTV" -and $AD_User.userCASESStatus -ne "" -and $AD_User.userCASESStatus -ne $null))
        {
            LogWrite "$userCommonName was previously inactive or left, user has activated and password has been reset"
            
            $tempPassword = Get-Password $User.BIRTHDATE "Complex"
            Set-ADAccountPassword -Identity $AD_User.DistinguishedName -Reset -NewPassword (ConvertTo-SecureString -AsPlainText $tempPassword -Force)
            Set-ADUser -Identity $AD_User.DistinguishedName -ChangePasswordAtLogon:$true

            #Put in Code to move to new OU as they are re-enabled
        }

        ##Remove the user of inactivity
        if (($AD_User.Enabled -ne $true) -and $ignoredInactiveUser.USERID -contains $AD_User.SamAccountName)
        {
            $AD_User.Enabled = $true
            LogWrite "$userCommonName's account disabled, reactivating"
            UpdateADArray "Inactive Date" "userInactive" "SETNULL"
        }

        updateADValue "CASES Status" "userCASESStatus" $User.STAFF_STATUS

        ##Do Staff Specific Variable Assignment

        ##Employee Number (PAYROLL_REC/T0 Number)
        updateADValue "Employee Number" "EmployeeNumber" $User.PAYROLL_REC_NO

        ##Employee ID (CASESID)
        updateADValue "Employee ID" "EmployeeID" $User.SFKEY.ToUpper()

        ##Email Address
        if ($calcUPN -ne $AD_User.mail)
        {
            if ($AD_User.mail  -like "*$upnDomain")
            {
                updateADArray "Other Mailboxes" "otherMailbox" $AD_User.mail.ToLower()
            }
            
            updateADValue "Email Address" "mail" $calcUPN
        }


        ##Licencing ACLs

        Set-Licencing

        #Write User back to AD if not a dry run
        if ($DryRun -eq $false)
        {
            ### Change proxy addresses where required
            foreach ($removedAddress in $changedProxyAddresses)
            {
                $AD_User.proxyAddresses.Remove($removedAddress)
                updateADArray "proxyAddresses" "proxyAddresses" "smtp:$($removedAddress.Substring(5,($removedAddress.Length-5)))"
            }
            
            #Save changes to AD User
            Set-ADUser -Instance $AD_User
        
            #Rename the user if Common Name has changed and not a dry run
            if ($commonNameChange -eq $true)
            {
                Rename-ADObject -Identity $AD_User -NewName $userCommonName
            }
        
        }

        if ($DryRunPause -eq $true)
        {
            pause
        }
    }
    elseif (($User.STAFF_STATUS -eq "LEFT" -or $User.STAFF_STATUS -eq "INAC"  -or $User.STAFF_STATUS -eq "INACLOGIN") -and ($ignoredAccounts.CASESID -notcontains $User.SFKEY)) #-and ($User.PAYROLL_REC_NO -ne $null -and  $User.PAYROLL_REC_NO -ne ""))
    {
        
        ### Set Variable to remove user to false, to trigger removal of user, set to true programmatically
        $AD_User = $null
        $DET_User = $null
        $calcUPN = $null
        $T0NUM = $null
        $disableUser = $false
        $targetDN = $null

        #Calculated Variables
        $CASESID = $User.SFKEY
        $T0NUM = $User.PAYROLL_REC_NO
        
        if ($User.PAYROLL_REC_NO -ne "" -and $user.PAYROLL_REC_NO -ne $null)
        {
            $duplicateCheck = (Import-Csv -Path $fileCASES | where-object {$_.PAYROLL_REC_NO -eq $T0NUM})

            if (($duplicateCheck | Measure-Object).Count -gt 1)
            {
                LogWrite "Duplicate T0 Number for $($User.SFKEY)" "Error"
            }

        }

        #AD Variables
        try 
        {
            if ($schoolRunAsLoggedIn -eq $true)
            {
                $AD_User = Get-ADUser $T0NUM -Server $SchoolDC -Properties * -ErrorAction Stop
            }
            else
            {
                $AD_User = Get-ADUser $T0NUM -Server $SchoolDC -Credential $schoolServiceCreds -Properties * -ErrorAction Stop
            }
        }
        catch
        {
            
            LogWrite "$userCommonName does not have account utilising T0 Number does not exist, trying CASES ID" "Verbose"
            
            ### Try SFKEY to see if its a local account
            try
            {
                if ($schoolRunAsLoggedIn -eq $true)
                {
                    $AD_User = Get-ADUser ($User.SFKEY) -Server $SchoolDC -Properties * -ErrorAction Stop
                }
                else
                {
                    $AD_User = Get-ADUser ($User.SFKEY) -Server $SchoolDC -Credential $schoolServiceCreds -Properties * -ErrorAction Stop
                }

                LogWrite "$userCommonName account found as a local account using their CASES ID" "Verbose"
            }
            catch
            {
                 LogWrite "$userCommonName - No Active Directory Account found" "Verbose"
                continue #Move to next iteration if no user exists
            }
        }
        
        Exit-StaffUser

    }
}

Send-HealthCheck $healthchecksCASESComplete $healthchecksEnabled $healthchecksDryRun $DryRun "Completed CASES File Processing"

LogWrite "*****Processing of CASES Data Complete*****"



# Process Pre-Staged CRT Users

LogWrite "*****Processing CRT Users*****"

Reset-crtAccounts -crtOU "$ouCRT,$staffBaseOU" -crtPrefix "CRT"
$CRTs = Import-Csv -Path $fileCRTUsers #| Where-Object 

foreach ($crt in $CRTs)
{
    if ((Get-Date $crt.StartDate) -le (Get-Date))
    {
        New-CRTAccount -userFirstName $crt.GivenName -userSurname $crt.Surname -userExit $crt.EndDate -crtID (Find-availableCRT -crtOU "$ouCRT,$staffBaseOU" -crtPrefix "CRT").samAccountName
    }
}

# Process Users that do not exist in CASES for Inactivity

<# LogWrite "*****Processing Users that do not exist in CASES for Inactivity*****"

Send-HealthCheck $healthchecksLocalStart $healthchecksEnabled $healthchecksDryRun $DryRun "Starting Local User Processing"

# Get time as FileTime for easier comparison
$time = (Get-Date).Adddays(-($inactiveAfter)).ToFileTime()

# Get users who are enabled, but have not logged in in more than $inctiveDays and have not had the password set in less than that time (to take into account new accounts) limit the search to the Staff OU
$inactiveUsers = Get-ADUser -Filter {((-not(LastLogonTimestamp -like "*") -or (LastLogonTimeStamp -lt $time)) -and (pwdLastSet -lt $time) ) -and (enabled -eq $true)} -SearchBase $staffBaseOU -Properties *

#Clear CSV User data and set status to inactive login for processing

$User = $null
$User = @{}
$User.STAFF_STATUS ="INACLOGIN"

$excludedPrefixes = @(
        'pst-'
        'sa.'
        'crt'
    )

# Cycle through the accounts, but ignoring the ones in the ignored users file and those whom are in the CASES CSV

foreach ($AD_User in $inactiveUsers)
{
    $userCommonName = "$($AD_User.GivenName) $($AD_User.Surname) ($($AD_User.SamAccountName))"
    
    $excludedByPrefix = $false

    #Continue to next itereation if the account prefix is excluded from checking
    foreach ($prefix in $excludedPrefixes)
    {
        if ($AD_User.SamAccountName.Substring(0,$prefix.Length) -eq $prefix)
        {
            $excludedByPrefix = $true
            break
        }
    }

    #Proceed with Exit process if it is not excludeds in any way
    if ($ignoredInactiveUser.USERID -notcontains $AD_User.SamAccountName -and ($Users.PAYROLL_REC_NO -notcontains $AD_User.SamAccountName -and $Users.SFKEY -notcontains $AD_User.SamAccountName) -and $excludedByPrefix -eq $false)
    {
        if ($AD_User.userInactive -eq $null -or $AD_User.userInactive -eq "")
        {
            LogWrite "$userCommonName is being marked as inactive - Local Search"
        }
        Exit-StaffUser
    }
}

Send-HealthCheck $healthchecksLocalComplete $healthchecksEnabled $healthchecksDryRun $DryRun "Completed Local User Processing"

LogWrite "*****Processing Users that do not exist in CASES for Inactivity Complete*****"#>

# Clean up temporary folder
if(test-path $temporaryDirectory)
{
    LogWrite "Removing Temporary Directory and Files" "Verbose"

    Get-ChildItem -Path $temporaryDirectory -Recurse | Remove-Item -force -recurse
    #Remove-Item $temporaryDirectory -Force 
}