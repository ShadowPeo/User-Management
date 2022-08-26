#Check AES Encryption Key
<#
if ( test-path "$PSScriptRoot\Credentials\edu001-$([Environment]::MachineName)-$([Environment]::UserName).crd" )
{

}#>

##Retrieve Credentials from File if exists
function Get-SavedCredentials ($filePath,$credentialUser)
{
    #Check for existing credential file
    if((($filePath -ne $null) -and ($filePath -ne "")) -and (($credentialUser -ne $null) -and ($credentialUser -ne "")))
    {
        if ( test-path $filePath )
        {
            $password = get-content $filePath | convertto-securestring
            [pscredential]$userCredentails = new-object -typename System.Management.Automation.PSCredential -argumentlist $credentialUser,$password
            LogWrite "Credentials Retrieved for user $credentialUser"
            return $userCredentails
        }
        else
        {
            LogWrite "Credential File does not exist"
        }
    }
    else
    {
        LogWrite "Blank Credential Filepath or User supplied"
    }
}

function Get-SavedCredentials_WithRequest ($filePath,$credentialUser)
{
    #Check for existing credential file
    if((($filePath -ne $null) -and ($filePath -ne "")) -and (($credentialUser -ne $null) -and ($credentialUser -ne "")))
    {
        if ( test-path $filePath )
        {
            $password = get-content $filePath | convertto-securestring
            [pscredential]$userCredentails = new-object -typename System.Management.Automation.PSCredential -argumentlist $credentialUser,$password
            LogWrite "Credentials Retrieved for user $credentialUser"
            return $userCredentails
        }
        else
        {
            try
            {
                $userCredentails = Get-Credential -UserName "$credentialUser" -Message "Please Enter Details for Authrorised User"
                $options = '&Yes', '&No' # 0=Yes, 1=No
                $response = $Host.UI.PromptForChoice("Save Credentials?", "Did you want to save the entered credentials?", $options, 1)
                if ($response -eq 0)
                {
                    $userCredentails.Password  | convertfrom-securestring | out-file $filePath
                }
                return $userCredentails
            }
            catch
            {
                LogWrite "Could not Get Valid Credentials"
                break
            }
        }
    }
}

function Get-SavedCredentials_WithRequest_NEW
{
    Param 
    (
        [Parameter(Mandatory=$true)][string]$filePath, #Path to the Credential Folder
        [Parameter(Mandatory=$true)][string]$credentialName, #Name Reference for the Credential
        [string]$credentialUser,
        [switch]$noUser = $false

    )

    if ($noUser -eq $false -and ($credentialUser -eq "" -or $credentialUser -eq $null))
    {
        throw "No User Supplied and user not excluded"
    }
    else
    {
            $fileName = "$credentialName-$([Environment]::MachineName)-$([Environment]::UserName).crd"
            Write-Host "$filePath\$fileName"
            pause
            #Check for existing credential file
            if ( test-path "$filePath\$fileName" )
            {
                $password = get-content "$filePath\$fileName" | convertto-securestring
                [pscredential]$userCredentails = new-object -typename System.Management.Automation.PSCredential -argumentlist $credentialUser,$password
                LogWrite "Credentials Retrieved for user $credentialUser"
                return $userCredentails
            }
            else
            {
                try
                {
                    $userCredentails = Get-Credential -UserName "$credentialUser" -Message "Please Enter Details for Authrorised User"
                    $options = '&Yes', '&No' # 0=Yes, 1=No
                    $response = $Host.UI.PromptForChoice("Save Credentials?", "Did you want to save the entered credentials?", $options, 1)
                    if ($response -eq 0)
                    {
                        if (Test-Path $filePath)
                        {                      
                            $userCredentails.Password  | convertfrom-securestring | out-file "$filePath\$fileName"
                        }
                        else
                        {
                            New-Item -ItemType Directory -Force -Path $filePath
                            $userCredentails.Password  | convertfrom-securestring | out-file "$filePath\$fileName"
                        }
                    }
                    return $userCredentails
                }
                catch
                {
                    LogWrite "Could not Get Valid Credentials" -logLevel "Error"
                    throw "Could not Get Valid Credentials"
                }
            }
    }
}