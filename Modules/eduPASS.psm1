Function Get-edu001_Mail($DC,$Credentials, $T0Number){
    try
    {
        $serviceUser = Get-ADUser $T0Number -Server $DC -Credential $Credentials -Properties mail
        return $serviceUser.mail
    }
    catch 
    {
        "$($_.Exception.Message)"
    }
}

Function Get-edu001_User($DC,$Credentials,$T0Number){
    try 
    {
        $serviceUser = Get-ADUser $T0Number -Server $DC -Credential $Credentials -Properties *
        
        return $serviceUser
    }
    catch 
    {
        "$($_.Exception.Message)"
    }
}

Function Get-edu001_User_ByEmail($DC,$Credentials,$DET_Mail){
    try
    {
        $serviceUser = Get-ADUser -Server $DC -Credential $Credentials -Filter {mail -eq $DET_Mail} -SearchBase "OU=Users,OU=Schools,DC=education,DC=vic,DC=gov,DC=au" -Properties *
        
        return $serviceUser
    }
    catch 
    {
        "$($_.Exception.Message)"
    }
}

Function Get-edu001_Mail_ByName($School,$DC,$Credentials, $DisplayName){
    try
    {
        $serviceUser = Get-ADUser -Server $DC -Credential $Credentials -Filter {DisplayName -eq $DisplayName} -SearchBase "OU=Users,OU=Schools,DC=education,DC=vic,DC=gov,DC=au" -Properties * |Where-Object {$_.memberof -contains "CN=$($School)-gs-All Staff,OU=School Groups,OU=Central,DC=services,DC=education,DC=vic,DC=gov,DC=au"}
        return $serviceUser.mail
    }
    catch 
    {
        "$($_.Exception.Message)"
    }
}

Function Get-edu002_Mail($School,$DC,$Credentials, $DisplayName){
    try
    {
        $edu002_user = Get-ADUser -Server $DC -Credential $Credentials -Filter {DisplayName -eq $DisplayName} -SearchBase "OU=Accounts,DC=services,DC=education,DC=vic,DC=gov,DC=au " -Properties * |Where-Object {$_.memberof -contains "CN=$($School)-gs-All Students,OU=School Groups,OU=Central,DC=services,DC=education,DC=vic,DC=gov,DC=au"}
        return $edu002_user.mail
    }
    catch 
    {
        "$($_.Exception.Message)"
    }
}

Function Get-edu002_User_ByName($School,$DC,$Credentials, $DisplayName){
    try
    {
        $serviceUser =  Get-ADUser -Server $DC -Credential $Credentials -Filter {DisplayName -eq $DisplayName} -SearchBase "OU=Accounts,DC=services,DC=education,DC=vic,DC=gov,DC=au " -Properties * |Where-Object {$_.memberof -contains "CN=$($School)-gs-All Students,OU=School Groups,OU=Central,DC=services,DC=education,DC=vic,DC=gov,DC=au"}
        return $serviceUser
        
    }
    catch 
    {
        "$($_.Exception.Message)"
    }
}

Function Get-edu002_User_ByEmail($School,$DC,$Credentials, $email){
    try
    {
        $serviceUser =  Get-ADUser -Server $DC -Credential $Credentials -Filter {mail -eq $email} -SearchBase "OU=Accounts,DC=services,DC=education,DC=vic,DC=gov,DC=au " -Properties * -ErrorAction Stop | Where-Object {$_.memberof -contains "CN=$($School)-gs-All Students,OU=School Groups,OU=Central,DC=services,DC=education,DC=vic,DC=gov,DC=au"}
        return $serviceUser
    }
    catch 
    {
        "$($_.Exception.Message)"
    }
}

