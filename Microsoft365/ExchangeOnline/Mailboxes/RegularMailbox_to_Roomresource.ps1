# Steps to convert a regular mailbox to <Regular | Room | Equipment | Shared> (Via PowerShell) POWERSHELL 7
<#
https://learn.microsoft.com/en-us/exchange/address-books/address-lists/manage-address-lists
#>

# Install Exchange Online Managment module by running the command below.
Install-Module -Name ExchangeOnlineManagement
# Import the module.
Import-Module ExchangeOnlineManagement
# Connect to Exchange Online (o1.help@aman.com) by the next command (a new brower tab will open prompting you the credentials).
Connect-ExchangeOnline
# Command line to turn the regular mailbox into the different options available (This case room resource).
Set-Mailbox -Identity sample@sample.com -Type Room
# Verify It's done.
Get-Mailbox -Identity sample@sample.com | Format-List RecipientTypeDetail
# Disconnect from Echange Online.
Disconnect-ExchangeOnline

# CREATING A ROOM LISTS (to make the room resource visible on outlook ROOM FINDER)

# Create a new room list
New-Distributiongroup -Name 'sample room' -RoomList
# Add a resource room as a member of the room list
Add-DistributionGroupMember -Identity 'sample room' -Member sample@sample.com
# Verify room list status "Identity = roomlist ID"
Get-DistributionGroup -Identity sample@samplegroupsarl.onmicrosoft.com
# Verify members that belong to the room list
Get-DistributionGroupMember -Identity sample@samplegroupsarl.onmicrosoft.com

# Get a room list by name 
Get-DistributionGroup -Anr 'sample room'
# End