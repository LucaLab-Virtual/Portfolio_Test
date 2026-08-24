# Retrieve the OneDrive link by email address (Microsoft Online SharePoint - SPOService PS 5.1)

# Install Module
Install-Module Microsoft.Online.SharePoint.PowerShell
# Update Module
Update-Module Microsoft.Online.SharePoint.PowerShell
# Import Module
Import-Module Microsoft.Online.SharePoint.PowerShell -UseWindowsPowerShell
# Connect
Connect-SPOService https://amangroupsarl-admin.sharepoint.com
# Get URL
Get-SPOSite -IncludePersonalSite $true -Limit All | Where {$_.Owner -eq 'sample@sample.com'} | Select Url
# Disconnect
Disconnect-SPOService
# End