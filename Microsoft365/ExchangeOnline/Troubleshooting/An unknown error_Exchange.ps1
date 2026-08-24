# Exchange: An unknown error has occurred. Refer to correlation ID: da9c79cc-5a62-498c-8169-7f7a13f7f707. (Windows PowerShell).

# Connect
Connect-MsolService
# Get the actual error message
Get-MsolUser -UserPrincipalName sample@sample.com | fl DisplayName, UserPrincipalName, @{Name="Error";Expression={($_.errors[0].ErrorDetail.objecterrors.errorrecord.ErrorDescription)}}

<#
SAMPLE OUTPUT after running the cmdlet above.

DisplayName         : sample account
UserPrincipalName   : sample@sample.com
Error               : {The value "96f0022a-8e83-4199-8a5c-be21af121b4d" of property "ExchangeGuid" is used by anohter recipient object.}
#>

# Then by using ExchangeOnlineManagement Module. PowerShell 7
Import-Module ExchangeOnlineManagement
# Connect
Connect-ExchangeOnline
# Get the users who have the ExchangeGuid Value (Sample Value "96f0022a-8e83-4199-8a5c-be21af121b4d")
# Option 1
Get-Mailbox | Select-Object UserPrincipalName, ExchangeGuid | Where-Object {$_.ExchangeGuid -eq "96f0022a-8e83-4199-8a5c-be21af121b4d"}
# Option 2
Get-Mailbox -ResutlSize Unlimited | Select-Object UserPrincipalName, ExchangeGuid | Where-Object {$_.ExchangeGuid -eq "96f0022a-8e83-4199-8a5c-be21af121b4d"}
# Disconnect
Disconnect-ExchangeOnline
# End