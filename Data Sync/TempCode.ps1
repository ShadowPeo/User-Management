(($ADUsers.samAccountName -notcontains $workingUser.SIS_ID) -and ($ADUsers.UserPrincipalName -like "$($workingUser.SIS_ID)@*" ) -and ($ADUsers.employeeID -notcontains $workingUser.SIS_ID))


 -and (([string]::IsNullOrWhiteSpace($workingUser.SIS_EMPNO) -and ($ADUsers.samAccountName -contains $workingUser.SIS_EMPNO)))
 -and $ADUsers.UserPrincipalName -like "$($workingUser.SIS_EMPNO)@*"
 -and ([string]::IsNullOrWhiteSpace($workingUser.SIS_EMPNO) -and ($ADUsers.employeeID -notcontains $workingUser.SIS_EMPNO)))




 $workingUser = $importedStaff | Where-Object -Property SIS_ID -eq "SIM"
 Write-Host "$($AD_User.samAccountName) | $($workingUser.SIS_ID)"