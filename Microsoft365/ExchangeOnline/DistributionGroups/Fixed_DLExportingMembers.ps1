# Export Members of a DL to a CSV file (PowerShell 7)

# Import
Import-Module -Name ExchangeOnlineManagement
# Connect
Connect-ExchangeOnline
# Create a variable that contains the distribution list name
$dlName = "Sample DL"
# Variable that gets the DL
$dl = Get-DistributionGroup -Identity $dlName
# Variable that gets members
$members = Get-DistributionGroupMember -Identity $dl.Identity
# Export Members to a CSV file
$members | Select-Object DisplayName, PrimarySmtpAddress, ObjectId | Export-Csv -Path "C:User\LuisRamirez\Desktop\Members_All.csv" -NoTypeInformation
# Disconnect
Disconnect-ExchangeOnline
# End