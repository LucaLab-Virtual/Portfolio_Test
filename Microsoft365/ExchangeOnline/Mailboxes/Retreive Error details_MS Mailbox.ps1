# Retrieve error details from any Mailbox (MsolService Windows PowerShell)

# Install the Microsoft Online Services
Install-Module MSOnline
# Import Module
Import-Module MSOnline
# Connect
Connect-MsolService
# Retreive error details
(Get-MsolUser -UserPrincipalName sample@sample.com).errors[0].ErrorDetail.objecterrors.errorrecord.ErrorDescription
# End