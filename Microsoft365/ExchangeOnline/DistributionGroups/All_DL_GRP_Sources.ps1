# Get all the distribution lists, groups, and resources an account belongs to in MS 365 - WindowsPowerShell

# Install the Microsoft Online Services
Install-Module MSOnline
# Import Module
Import-Module MSOnline
# Connect
Connect-MsolService
# You can use the following PowerShell command to get all the distribution lists, groups, and resources an account belongs to in MS 365:
Get-UnifiedGroupLinks -Identity sample@sample.com -LinkType Member | Select-Object DisplayName | Format-Table

# By AzureAD Module

# Install the AzureAD Module
Install-Module AzureAD
# Connect
Connect-AzureAD
# Get the info
Get-AzureADUserMembership -ObjectId "sample@sample.com" | Select-Object DisplayName

# Get the shared mailboxes a user belongs to by ExchangeOnline Module

# Import
Import-Module ExchangeOnlineManagement
# Connect
Connect-ExchangeOnline
# First cmdlet
Get-Mailbox -RecipientTypeDetails SharedMailbox -ResultSize Unlimited | Get-MailboxPermission | Where-Object { ($_.User -eq 'sample@sample.com') }
# Second cmdlet
Get-Mailbox -RecipientTypeDetails SharedMailbox | Get-MailboxPermission -User "sample@sample.com" | Where-Object {$_.AccessRights -eq "FullAccess"} | Select-Object User, Identity
# End