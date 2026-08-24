# Import the Active Directory module
Import-Module ActiveDirectory
 
# Specify the user's UPN (User Principal Name) and the alias you want to add
$userUPN = "kkrauss"
$alias = "kmares@loubachrodt.com"
 
# Get the user object
$user = Get-ADUser -Filter {UserPrincipalName -eq $userUPN}
 
# Add the alias to the proxyAddresses attribute
$user | Set-ADUser -Add @{proxyAddresses="smtp:$alias"}
 
# Verify the changes
Get-ADUser -Identity $userUPN -Properties proxyAddresses | Select-Object -ExpandProperty proxyAddresses