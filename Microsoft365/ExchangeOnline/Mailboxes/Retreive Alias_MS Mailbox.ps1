# Retrieve the mailboxes that have a specific email address as an alias (ExchangeOnlline PowerShell 7)

# Install Module
Install-Module Microsoft.Online.SharePoint.PowerShell
# Import Module
Import-Module Microsoft.Online.SharePoint.PowerShell -UseWindowsPowerShell
# Connect
Connect-ExchangeOnline
# Retreive Alias
Get-Mailbox -ResutlSize Unlimited | Where-Object {$_.EmailAddress -Like "*sample@sample.com*"} 
| Select-Object DisplayName,@{Name="EmailAddresses";Expression={$_.EmailAddress | Where-Object {$_ -Like "SMTP:*"}}}

# Retrieve the mailboxes that have a specific email address as an alias by ADDC

# Open Windows PowerShell in the Domain Controller
Get-ADUser -filter *EmailAddress -like '*sample@sample.com*' -Properties DisplayName, EmailAddress, | Select-Object DisplayName, EmailAddress
# End