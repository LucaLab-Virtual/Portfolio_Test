# Removing email forwarding (When Exchange fails) POWERSHELL 7

# Import
Import-Module ExchangeOnlineManagement
# Connect
Connect-ExchangeOnline
# Get Mailbox forwarding info
Get-Mailbox -Identity <Email address> | select UserPrincipalName, ForwardingSmtpAddress, DeliveryToMailBoxAndForward
# Set email forwarding
Get-Mailbox -Identity <Email address> | Set-Mailbox -ForwardingSmtpAddress $Null
# End