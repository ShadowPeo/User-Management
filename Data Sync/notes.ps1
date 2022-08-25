$handlingStudentEmail = 2 #3 = pull from AD UPN, 4 = Pull from AD Mail, 5 = Pull from AD ProxyAddresses looking for primary (Capital SMTP)
$handlingStaffEmail = 2 #3 = pull from AD UPN, 4 = Pull from AD Mail, 5 = Pull from AD ProxyAddresses looking for primary (Capital SMTP), 6 = Use employeeID (PAYROLL_REC_NO/EmployeeNumber) from AD, fall back to SFKEY
$handlingStudentUsername = 1 #3 = pull from AD UPN, 4 = Pull from AD Mail, 5 Use samAccountName
$handlingStaffUsername = 1 #3 = pull from AD UPN, 4 = Pull from AD Mail, 5 Use samAccountName, 6 = Use employeeID (PAYROLL_REC_NO/EmployeeNumber) from AD, fall back to SFKEY
$handlingStudentAlias = 1 #2= use samAccountName, 3 = Use employeeID from Active Directory - Fall back to STKEY
$handlingStaffAlias = 1 #2= use samAccountName - Fall back to SFKEY, 3 = Use employeeID from Active Directory - Fall back to SFKEY
$handlingValidateLicencing = $true #Validate the licencing for Oliver, this will drop accounts where it is explictly disabled or where no user exists 
$handlingLicencingValue = "licencingOliver" #The attribute name for the licencing Data

samAccountName
upn
mail
proxyAddresses
employeeID

SearchBase
Proxy Address Lookup
Properties Array for AD lookup if Licencing lookup turned on and value is present

VALIDATE ALL AD LOOKUPS BAR SMTP AS ITS NOT IMPLEMENTED