

#Get Password from Dinopass
function Get-Password
{
    Param(
            [string]$userBirthdate,
            [string]$pwType
         )

    try 
    {
        if ( $pwType -ieq "Simple")
        {
            return Invoke-RestMethod -UseBasicParsing "http://www.dinopass.com/password"
        }
        else
        {
            return Invoke-RestMethod  -UseBasicParsing "http://www.dinopass.com/password/strong"
        }
    }
    catch
    {
        if ($userBirthdate -ne $null -and $userBirthdate -ne "")
        {
            return "Western@" + (Get-Date -date $userBirthdate -format ddMM)
        }
        else
        {
            return "Western@" + (Get-Random -Minimum 1000 -Maximum 9999)
        }
        
    }
}

function Get-validUPNName 
{
    Param
    (
        [Parameter(Mandatory=$true)]
        [string]$userName
    )

    $userName = ($userName -replace '[^a-zA-Z0-9-]', '').ToLower()

    return $userName

}

#Get a Valid UPN/email given the firstname and surname of the account, check that it does not exist allocated to another user, if it does then increment by 1
function Get-validUPN
{
    Param
    (
        [Parameter(Mandatory=$true)][string]$userFirstName,
        [Parameter(Mandatory=$true)][string]$userSurname,
        [Parameter(Mandatory=$true)][string]$userID,
        [Parameter(Mandatory=$true)][string]$userUPNDomain
    )

    $firstname = Get-validUPNName $userFirstName
    $surname = Get-validUPNName $userSurname
    $i = $null

    
    while ($outputUPN -eq $null)
    {
        $checkUPN = "$firstname.$surname$i@$userUPNDomain"
        if (Get-ADUser -Filter ("samAccountName -ne '$userID' -and (proxyAddresses -eq 'SMTP:$checkUPN' -or proxyAddresses -eq 'smtp:$checkUPN' -or mail -eq '$checkUPN' -or otherMailbox -eq '$checkUPN')"))
        {
            $i = $i + 1
            if ($i -gt 100)
            {
                $outputUPN = "ERROR"
            }
            
        }
        else
        {
            $outputUPN =  $checkUPN
        }
    }

    if ($outputUPN -ne "ERROR")
    {
        return $outputUPN
    }

}


# Reset CRT accounts to disabled if they have expired, and clear the expiry date

function Reset-crtAccounts
{

    Param
    (
        [string]$crtOU,
        [string]$crtPrefix
    )

    $currentDate = (Get-Date).ToFileTime()

    if (($crtOU -ne $null -and $crtOU -ne "") -and ($crtPrefix -ne "" -and $crtPrefix -ne $null))
    {
        try 
        {
            $crtPrefix = "$crtPrefix*"
            $CRT_userList = Get-ADUser -Filter {((accountExpires -gt 0 ) -and (accountExpires -le $currentDate)) -and (samAccountName -like $crtPrefix)} -SearchBase "$crtOU" -Properties accountExpires, displayName
        }
        catch
        {
             LogWrite "No User Accounts require resetting, continuing"
        }

    }
    elseif(($crtOU -ne $null -and $crtOU -ne ""))
    {
        try 
        {
            $CRT_userList = Get-ADUser -Filter {((accountExpires -gt 0 ) -and (accountExpires -le $currentDate))} -SearchBase "$crtOU"  -Properties accountExpires, displayName
        }
        catch
        {
            LogWrite "No User Accounts require resetting, continuing"
        }
    }
    elseif(($crtPrefix -ne "" -and $crtPrefix -ne $null))
    {
        $crtPrefix = "$crtPrefix*"

        try 
        {
            $CRT_userList = Get-ADUser -Filter {((accountExpires -gt 0 ) -and (accountExpires -le $currentDate)) -and (samAccountName -like $crtPrefix)}  -Properties accountExpires, displayName
        }
        catch
        {
            LogWrite "No User Accounts require resetting, continuing"
        }
    }
    else
    {
        LogWrite "No valid search paramters provided, not continuing"
        break
    }


    foreach ($crtAccount in $CRT_userList)
    {
        $crtAccount.enabled = $false
        $crtAccount.DisplayName = $crtAccount.samAccountName
        $crtAccount.GivenName = "Unused"
        $crtAccount.Surname = "CRT Account"
        $crtAccount.DisplayName = "Unused CRT Account"

        Set-ADUser -Instance $crtAccount
        
        Clear-ADAccountExpiration -Identity $crtAccount.samAccountName

        Set-ADAccountPassword -Identity $crtAccount.DistinguishedName -Reset -NewPassword (ConvertTo-SecureString -AsPlainText (Get-Password -pwType "Complex") -Force)
        
        Rename-ADObject -Identity $crtAccount -NewName "Unused CRT Account $($crtAccount.samAccountName)"

        LogWrite "$($crtAccount.samAccountName) is an expired account, resetting"
    }
}

# If a preferred account is specified try that, if not specified or available, return first available account (as sorted by SAM Account Name)
function Find-availableCRT
{
    Param
    (
        [string]$crtOU,
        [string]$crtPrefix,
        [string]$crtID
    )

    $CRT_userList = $null
    
    if ($crtID -ne $null -and $crtID -ne "")
    {
        try
        {
            $CRT_userList = Get-ADUser -Filter {(enabled -eq $false) -and (samAccountName -eq $crtID)} -Properties * | sort-object samAccountName
            
            return $CRT_userList
        }
        catch
        {
            LogWrite -logString "Account with an identifier of $crtID is not found or already in use, continuing on" -logLevel "Verbose"
        }
    }


    if (($crtOU -ne $null -and $crtOU -ne "") -and ($crtPrefix -ne "" -and $crtPrefix -ne $null))
    {
        try 
        {
            $crtPrefix = "$crtPrefix*"
            $CRT_userList = @(Get-ADUser -Filter {(enabled -eq $false) -and (samAccountName -like $crtPrefix)} -SearchBase "$crtOU" -Properties * | sort-object samAccountName)
            return $CRT_userList[0]
            
        }
        catch
        {
            return "-2"
        }

    }
    elseif(($crtOU -ne $null -and $crtOU -ne ""))
    {
        try 
        {
            $CRT_userList = @(Get-ADUser -Filter {(enabled -eq $false)} -SearchBase "$crtOU" -Properties * | sort-object samAccountName)
            return $CRT_userList[0]
        }
        catch
        {
            return "-2"
        }
    }
    elseif(($crtPrefix -ne "" -and $crtPrefix -ne $null))
    {
        $crtPrefix = "$crtPrefix*"
        try 
        {
            $CRT_userList = @(Get-ADUser -Filter {(enabled -eq $false) -and (samAccountName -like $crtPrefix)} -Properties * | sort-object samAccountName)
            return $CRT_userList[0]
        }
        catch
        {
            return "-2"
        }
    }
    else
    {
        return "-1"
    }
}


#Remove Homedrive
function Remove-HomeFolder ($directoryToCheck, $archivePath, $userSAM)
{
        if (Test-Path $directoryToCheck) 
        {
            try 
            {
                #Compress-Archive -Path $directoryToCheck -DestinationPath ($archivePath + "\" + $userSAM) -Force
                LogWrite "$userCommonName's Home Folder has been transfered to Archive"

                try 
                {
                    #Remove-Item $directoryToCheck -Force -Recurse
                    LogWrite "$userCommonName's Home Folder has been removed"
                    return $true
                }
                catch 
                {
                    LogWrite "$userCommonName's Home Folder removal failed"
                    return $false
                }
            }
            catch 
            {
                LogWrite "$userCommonName's Home Folder compression failed, please check"
                return $false
            }
        }
        else
        {
            LogWrite "$userCommonName's Home Folder does not exist, removing reference"
            return $true
        }
}

function New-CRTAccount
{
    Param
    (
        [Parameter(Mandatory=$true)][string]$userFirstName,
        [Parameter(Mandatory=$true)][string]$userSurname,
        [Parameter(Mandatory=$true)][string]$userExit,
        [Parameter(Mandatory=$true)][string]$crtID
    )

    $crtUser = Get-ADUser -Identity $crtID -Properties *
    $crtUser.GivenName = $crt.GivenName
    $crtUser.Surname = $crt.Surname
    $crtUser.DisplayName = "$($crt.GivenName) $($crt.Surname)"
    $crtUser.AccountExpirationDate = ((Get-Date (Get-Date $crt.EndDate)))
    $crtUser.Enabled = $true
    Set-ADUser -Instance $crtUser
    
    Set-ADAccountPassword -Identity $crtUser.DistinguishedName -Reset -NewPassword (ConvertTo-SecureString -AsPlainText (Get-Password -pwType "Complex") -Force)
    
    Rename-ADObject -Identity $crtUser -NewName "$($crt.GivenName) $($crt.Surname)"

    LogWrite "Allocating CRT Account $($crtUser.SamAccountName) to $($crt.GivenName) $($crt.Surname) with an expiry of $(Get-Date $crt.EndDate -UFormat '%d/%m/%Y')"
}



###############################################
function Get-StaffAccount 
{
        Param
    (
        [Parameter(Mandatory=$true)] [string]$serverSchoolDC,
        [string]$userSurname,
        [string]$userID,
        [string]$userUPNDomain
    )
        
        #AD Variables
        try 
        {
            if ($schoolRunAsLoggedIn -eq $true)
            {
                $AD_User = Get-ADUser $T0NUM -Server $serverSchoolDC -Properties * -ErrorAction Stop
            }
            else
            {
                $AD_User = Get-ADUser $T0NUM -Server $serverSchoolDC -Credential $schoolServiceCreds -Properties * -ErrorAction Stop
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
                    $AD_User = Get-ADUser ($User.SFKEY) -Server $serverSchoolDC -Properties * -ErrorAction Stop
                }
                else
                {
                    $AD_User = Get-ADUser ($User.SFKEY) -Server $serverSchoolDC -Credential $schoolServiceCreds -Properties * -ErrorAction Stop
                }

                LogWrite "$userCommonName account found as a local account using their CASES ID" "Verbose"

                if ($AD_User -ne $null -and $TONUM -ne $null -and $TONUM -ne "")
                {
                    LogWrite "$userCommonName now has a T0 Number (PAYROLL_REC_NO) renamiging the account to use it"
                    #### Insert code to change samAccountName
                }
            }
            catch
            {
                #Create user utilising the correct ID

                LogWrite "No existing user found with either T0 number or CASES ID, Creating User" "Verbose"

                if ($T0NUM -ne $null -and $T0NUM -ne "")
                {
                    LogWrite "$userCommonName user account being created utilising T0 number"
                }
                else
                {
                    LogWrite "$userCommonName user account being created utilising CASES ID"
                }
                
                continue #Temporary until user creation implemented
            }
        }
}

<#function Exit-StaffUser ($User)
{
    
}#>



##Update an Active Directory Array based Value
function updateADArray ($attributeDisplay, $attributeToCheck, $attributeCheckValue)
{


    if ($attributeCheckValue -ne "SETNULL")
    {
        if ($AD_User.$attributeToCheck -notcontains $attributeCheckValue)
        {
            $AD_User.$attributeToCheck.Add($attributeCheckValue) | Out-Null
            LogWrite "Adding $attributeCheckValue to $userCommonName's $attributeDisplay"  -foregroundColour:"Magenta"
        }
    }
    else
    {
        if ($AD_User.$attributeToCheck -ne $null)
        {
            Set-ADUser $AD_User -Clear $attributeToCheck
            LogWrite "Clearing $attributeCheckValue on $userCommonName's account"  -foregroundColour:"Magenta"
        }
    }
}

##Update an Active Directory (non-Array) Value
function updateADValue ($attributeDisplay, $attributeToCheck, $attributeCheckValue)
{
    if (($attributeCheckValue -ne $null) -and ($attributeCheckValue -ne ""))
    {
        ####Set Attribute to null if commanded
        if ($attributeCheckValue -eq "SETNULL")
        {
            $attributeCheckValue = $null
        }
        if($AD_User.$attributeToCheck -eq $null) 
        {
            $AD_User.$attributeToCheck = $attributeCheckValue
            LogWrite "Setting $userCommonName's $attributeDisplay to $attributeCheckValue"  -foregroundColour:"Magenta"
        }
        ElseIf ($AD_User.$attributeToCheck -ne $attributeCheckValue)
        {
            $AD_User."$attributeToCheck" = $attributeCheckValue
            LogWrite "Updating $userCommonName's $attributeDisplay to $attributeCheckValue"  -foregroundColour:"Magenta"
        }
    }
}

##Update an Active Directory (non-Array) Value, return boolean if it has been completed
function updateADValueReturnBool ($attributeDisplay, $attributeToCheck, $attributeCheckValue)
{
    if($AD_User.$attributeToCheck -eq $null) 
    {
        $AD_User.$attributeToCheck = $attributeCheckValue
        LogWrite "Setting $userCommonName's $attributeDisplay to $attributeCheckValue"  -foregroundColour:"Magenta"
        return $true
    }
    ElseIf ($AD_User.$attributeToCheck -ne $attributeCheckValue)
    {
        $AD_User.$attributeToCheck = $attributeCheckValue
        LogWrite "Updating $userCommonName's $attributeDisplay to $attributeCheckValue"  -foregroundColour:"Magenta"
        return $true
    }
    else
    {
        return $false
    }
}

function Get-Student_WithCreate
{
    Param 
    (
        [Parameter(Mandatory=$true)][string]$primaryUserID, #UserID to Search
        [Parameter(Mandatory=$true)][string]$domainController,
        [Parameter(Mandatory=$true)][string]$FirstName,
        [Parameter(Mandatory=$true)][string]$Surname,
        [Parameter(Mandatory=$true)][string][string]$userYear,
        [Parameter(Mandatory=$true)][string][string]$userStatus,
        [Parameter(Mandatory=$true)][string][string]$ouCreation,
        $domainCredentials
    )

    $Surname = $Surname.ToUpper()

    if ($domainCredentials -eq $null -and $domainCredentials -eq "")
    {
        ### Try to retrieve AD User based upon CASES Code, if not try to create user

        try 
        {
            $AD_User = Get-ADUser $primaryUserID -Server $domainController -Properties * -ErrorAction Stop
        }
        catch
        {
            LogWrite "A User with the UserID of $primaryUserID does not exist, creating user"

            if ($User.STATUS -eq "FUT")
            {
                New-ADUser -Name "$FirstName $Surname ($PrimaryUserID)" -GivenName "$FirstName" -Surname "$Surname" -SamAccountName "$PrimaryUserID" -UserPrincipalName "$PrimaryUserID@westernportsc.vic.edu.au" -Path "OU=$ouNew,$studentBaseOU" -AccountPassword(ConvertTo-SecureString (Get-Password -pwType 'Complex' -userBirthdate $User.BIRTHDATE) -AsPlainText -Force) -Enabled $false -Server $domainController
            }
            else
            {
                New-ADUser -Name "$FirstName $Surname ($PrimaryUserID)" -GivenName "$FirstName" -Surname "$Surname" -SamAccountName "$PrimaryUserID" -UserPrincipalName "$PrimaryUserID@westernportsc.vic.edu.au" -Path "OU=Year $($User.SCHOOL_YEAR),$studentBaseOU" -AccountPassword(ConvertTo-SecureString (Get-Password -pwType 'Complex' -userBirthdate $User.BIRTHDATE) -AsPlainText -Force) -Enabled $false -Server $domainController
            }

            $AD_User = Get-ADUser $User.STKEY -Server $domainController -Properties * -ErrorAction Stop
        }

    }
    else
    {
        
        try 
        {
            $AD_User = Get-ADUser $primaryUserID -Server $domainController -Credential $domainCredentials -Properties * -ErrorAction Stop
        }
        catch
        {
            LogWrite "A User with the UserID of $primaryUserID does not exist, creating user"
        
            if ($User.STATUS -eq "FUT")
            {
                New-ADUser -Name "$FirstName $Surname ($PrimaryUserID)" -GivenName "$FirstName" -Surname "$Surname" -SamAccountName "$PrimaryUserID" -UserPrincipalName "$PrimaryUserID@westernportsc.vic.edu.au" -Path "OU=$ouNew,$studentBaseOU" -AccountPassword(ConvertTo-SecureString (Get-Password -pwType 'Complex' -userBirthdate $User.BIRTHDATE) -AsPlainText -Force) -Enabled $false -Server $domainController -Credential $domainCredentials
            }
            else
            {
                New-ADUser -Name "$FirstName $Surname ($PrimaryUserID)" -GivenName "$FirstName" -Surname "$Surname" -SamAccountName "$PrimaryUserID" -UserPrincipalName "$PrimaryUserID@westernportsc.vic.edu.au" -Path "OU=Year $($User.SCHOOL_YEAR),$studentBaseOU" -AccountPassword(ConvertTo-SecureString (Get-Password -pwType 'Complex' -userBirthdate $User.BIRTHDATE) -AsPlainText -Force) -Enabled $false -Server $domainController -Credential $domainCredentials
            }
        }
        $AD_User = Get-ADUser $User.STKEY -Server $domainController -Credential $domainCredentials -Properties * -ErrorAction Stop
    }

    return $AD_User
}

function Set-StudentGroups
{
    
    $userGroups = (Get-ADPrincipalGroupMembership $AD_User.SamAccountName).name
    
    foreach ($group in $groupConfig)
    {
        $attributeName = $null
        $attributeValue = $null

        $attributeName = $group.attributeName
        $attributeValue = $group.attributeValue

        if ($attributeName -eq $null -or $attributeName -eq $null -or $attributeName.ToUpper() -eq "ALL")
        {
            if ($userGroups -notcontains $group.group)
            {
                Add-ADGroupMember -Identity $group.group -Members $AD_User.SamAccountName
                LogWrite "$userCommonName's has been added to $($group.group)"
            }
        }
        elseif ($AD_User.$attributeName -eq $attributeValue)
        {
            if ($userGroups -notcontains $group.group)
            {
                Add-ADGroupMember -Identity $group.group -Members $AD_User.SamAccountName
                LogWrite "$userCommonName's has been added to $($group.group)"
            }

        }
        else
        {
            if ($userGroups -contains $group.group)
            {
                Remove-ADGroupMember -Identity $group.group -Members $AD_User -Confirm:$false
                LogWrite "$userCommonName's has been removed from $($group.group)"
            }
        }
    }
}

function Initialize-User
{
    Set-StudentGroups

    $welcomeLetter = $null
    
    $emails = $emailConfig | WHERE-OBJECT {($_.emailTemplate -eq "ALL") -or ($_.emailTemplate -eq "") -or ($_.emailTemplate -eq $null) -or ($_.emailTemplate -eq "StudentWelcome")}
    
    if ($changeAUP -eq $true -and $AD_User.userAUPStatus -eq "Returned" -and $AD_User.lastLogon -eq 0  -and ($Y07InitializationLock -eq $false -and $AD_User.Department -eq "Year 07"))
    {
        LogWrite "Change User Password and send Welcome Email"
        $Password = Get-Password -pwType 'Complex' -userBirthdate $User.BIRTHDATE
        Set-ADAccountPassword -PassThru -Identity $AD_User -NewPassword ($Password  | ConvertTo-SecureString -AsPlainText -Force) -Reset | Out-Null   
        
        $welcomeLetter = Write-StudentWelcomeLetter -DisplayName $AD_User.DisplayName -CASESID $AD_User.SamAccountName -UPN $AD_User.UserPrincipalName -userPassword $Password -TempDirectory $temporaryDirectory
        $emailBody = Send-StudentWelcome -DisplayName $AD_User.DisplayName -CASESID $AD_User.SamAccountName -UPN $AD_User.UserPrincipalName -HomeGroup $AD_User.physicalDeliveryOfficeName
    }
    elseif ($AD_User.userAUPStatus -eq "Unreturned")
    {
        $emailBody = Send-StudentWelcomeNoAUP -DisplayName $AD_User.DisplayName -CASESID $AD_User.SamAccountName -HomeGroup $AD_User.physicalDeliveryOfficeName
    }
    
    

    if ($Y07InitializationLock -eq $false -and $AD_User.Department -eq "Year 07")
    {
        foreach ($email in $emails)
        {
            $attributeName = $null
            $attributeValue = $null

            $attributeName = $email.attributeName
            $attributeValue = $email.attributeValue

            if ($attributeName -eq $null -or $attributeName -eq $null -or $attributeName.ToUpper() -eq "ALL")
            {
                if ($welcomeLetter -ne $null)
                {
                    send-MailMessage -to $email.emailTo -from $EmailFrom -subject "New Student | $($AD_User.Department) | $($AD_User.physicalDeliveryOfficeName) | $($AD_User.DisplayName)"  -SmtpServer $SMTPServer -Attachment $welcomeLetter -Body $emailBody -bodyashtml -verbose
                }
                else
                {
                    send-MailMessage -to $email.emailTo -from $EmailFrom -subject "New Student | $($AD_User.Department) | $($AD_User.physicalDeliveryOfficeName) | $($AD_User.DisplayName)"  -SmtpServer $SMTPServer -Body $emailBody -bodyashtml -verbose
                }
                LogWrite -logString "Sending Welcome email to $($email.emailTo)"
            }
            elseif ($AD_User.$attributeName -eq $attributeValue)
            {
                if ($welcomeLetter -ne $null)
                {
                    send-MailMessage -to $email.emailTo -from $EmailFrom -subject "New Student  | $($AD_User.Department) | $($AD_User.physicalDeliveryOfficeName) | $($AD_User.DisplayName)"  -SmtpServer $SMTPServer -Attachment $welcomeLetter -Body $emailBody -bodyashtml -verbose
                }
                else
                {
                    send-MailMessage -to $email.emailTo -from $EmailFrom -subject "New Student  | $($AD_User.Department) | $($AD_User.physicalDeliveryOfficeName) | $($AD_User.DisplayName)"  -SmtpServer $SMTPServer -Body $emailBody -bodyashtml -verbose
                }
                LogWrite -logString "Sending Welcome email to $($email.emailTo)"
            }

        }
    }
}