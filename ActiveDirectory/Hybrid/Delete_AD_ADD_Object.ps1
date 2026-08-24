# Delete AD and AAD users (Module AzureAD PowerShell 5.1)

# Install Azure PowerShell Module
Install-Module -Name Az -Force -AllowClobber
# Install AzureAD module
Install-Module -Name AzureAD
# Connect to Azure AD tenant
Connect-AzureAD
# Verify your connnection
Get-AzureADTenantDetail
# Disconnect from the Azure AD tenant
Disconnect-AzureAD

# The license must be removed before running the next two cmdlets
# Delete a user
Remove-AzureADUser -ObjectId #<userObjectId>

# Delete a group or distribution list
Remove-AzureADGroup -ObjectId #<ObjectId>

# Force sync
# Make sure the Microsoft Azure AD Sync service (services.msc) is running before using the next cmdlets
# Import ADSync module
Import-Module ADSync
# Force a full synchronization
Start-ADSyncSyncCycle -PolicyType Initial
# Or force an incremental synchronization
Start-ADSyncSyncCycle -PolicyType Delta