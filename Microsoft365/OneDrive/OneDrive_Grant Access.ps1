# Grant Access to a User's OneDrive using WindowsPowerShell

# Import Module
Import-Module Microsoft.OneDrive.SharePoint.PowerShell
# Connect
Connect-SPOService https://amangroupsarl-admin-sharepoint.com
# Grant access
Set-SPOUser -Site https://sampletenant-my.shapoint.com/personal/sample_site_com -IsSiteCollectionAdmin $true LoginName sample2@sample.com
# End