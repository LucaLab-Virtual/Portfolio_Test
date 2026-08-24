# Microsoft Sharepoint Online WindowsPowerShell

# Install Module
Install-Module Microsoft.Online.SharePoint.PowerShell
# Update Module
Update-Module Microsoft.Online.SharePoint.PowerShell
# Import Module
Import-Module Microsoft.Online.SharePoint.PowerShell -UseWindowsPowerShell
# Get the Module version
Get-Module -Name Microsoft.Online.SharePoint.PowerShell -ListAvailable | Select Name, version
# Uninstall Module
Uninstall-Module -Name Microsoft.Online.SharePoint.PowerShell -force
# Connect
Connect-SPOService https://sampletenant-admin-sharepoint.com
# Get all sites
Get-SPOSite
# Disconnect
Disconnect-SPOService
# End