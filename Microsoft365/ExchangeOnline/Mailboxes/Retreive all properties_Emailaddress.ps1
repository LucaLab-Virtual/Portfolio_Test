# Retrieves all the properties of an email address (ExchangeOnline PowerShell 7)

# Install Exchange Online Managment module by running the command below.
Install-Module -Name ExchangeOnlineManagement
# Import the module.
Import-Module ExchangeOnlineManagement
# Connect to Exchange Online (admin@sample.com) by the next command (a new browser tab will open prompting you the credentials).
Connect-ExchangeOnline
# Retreive the properties
Get-Recipient -Identity sample@sample.com | Select-Object * | Out-File -FilePath "C:\Users\LuisRamirez\Documents\emailpropterties.txt"
# End