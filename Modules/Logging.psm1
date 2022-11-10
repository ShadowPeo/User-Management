Function Start-Log ()
{
    #Make sure $LogPath exists
    If(!(test-path $LogPath))
    {
          New-Item -ItemType Directory -Force -Path $LogPath
    }

}

# Write to the logfile and other procceses depending on the arguments supplied
Function LogWrite
{
    Param 
    (
    
        [Parameter(Mandatory=$true)][string]$logString, 
        [string]$logLevel,
        [string]$TrasactionCommand,
        [string]$foregroundColour,
        [boolean]$noOutput
    
    )
    #Write-Host "$foregroundColour"
    $foregroundSet = $null

    if ($logLevel.ToLower() -eq "warning")
    {
        $logString = "WARNING: $logString"
        $foregroundSet = "Yellow"
    }
   
    if ($logLevel.ToLower() -eq "critical")
    {
        $logString = "CRITICAL: $logString"
        $foregroundSet = "Red"
    }

    if ($logLevel.ToLower() -eq "verbose")
    {

    }
    elseif ($logLevel.ToLower() -eq "warning" -and ($Flag.ToLower() -eq "warning" -or $Flag.ToLower() -eq "critical"))
    {

    }
    elseif ($logLevel.ToLower() -eq "critical" -and $Flag.ToLower() -eq "critical")
    {

    }
    elseif ($logLevel.ToLower() -ne "verbose")
    {
        if ($foregroundColour -eq $null -or $foregroundColour -eq "")
        {
            $foregroundSet = "Green"
        }
        else
        {
            $foregroundSet = $foregroundColour
        }
    }
    
    Add-content $Logfile -value "$(Get-Date -UFormat '+%Y-%m-%d %H:%M:%S') - $logString"

    #Write to output no matter the log level
    if ($noOutput -ne $true)
    {
        if ($foregroundSet -ne "" -and $null -ne $foregroundSet)
        {
            Write-Host "$(Get-Date -UFormat '+%Y-%m-%d %H:%M:%S') - $logString" -ForegroundColor $foregroundSet
        }
        else
        {
            Write-Host "$(Get-Date -UFormat '+%Y-%m-%d %H:%M:%S') - $logString"
        }

    }
}

# Ping a healthchecks server at the given URI but only if the system is allowed to

function Send-HealthCheck
{
    Param 
    (
        [Parameter(Mandatory=$true)][string]$healthcheckURI, 
        [Parameter(Mandatory=$true)][boolean]$healthchecksEnabled, 
        [Parameter(Mandatory=$true)][boolean]$healthchecksDryRun,
        [Parameter(Mandatory=$true)][boolean]$DryRun,
        [Parameter(Mandatory=$true)][string]$healthcheckName 
    )
    
    if ($healthchecksEnabled -eq $true -and (($healthchecksDryRun -eq $true -and $DryRun -eq $true) -or $DryRun -eq $false))
    {
        if ($healthcheckName -ne "" -and $healthcheckName -ne $null)
        {
            LogWrite "Sending Healthcheck ping for $healthcheckName" "VERBOSE"
        }
        else
        {
            LogWrite "Sending Healthcheck ping" "VERBOSE"
        }
        
        Invoke-RestMethod $healthcheckURI | Out-Null
    }
}
