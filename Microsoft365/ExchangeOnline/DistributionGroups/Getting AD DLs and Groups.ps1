# Import the Active Directory module if not already loaded
Import-Module ActiveDirectory

# Get all distribution groups
$distGroups = Get-ADGroup -Filter {GroupCategory -eq 'Distribution'} -Properties Name

# Get all security groups
$securityGroups = Get-ADGroup -Filter {GroupCategory -eq 'Security'} -Properties Name

# Display the results
Write-Host "Distribution Groups:"
$distGroups | Select-Object Name | Format-Table -AutoSize

Write-Host "Security Groups:"
$securityGroups | Select-Object Name | Format-Table -AutoSize


#--------------------------------------------------------------------------------------------

# Import the Active Directory module if not already loaded
Import-Module ActiveDirectory

# Get all distribution groups
$distGroups = Get-ADGroup -Filter {GroupCategory -eq 'Distribution'} -Properties Name

# Get all security groups
$securityGroups = Get-ADGroup -Filter {GroupCategory -eq 'Security'} -Properties Name

# Define the output file path
$outputFilePath = "C:\Users\orbis\Downloads\DLsAndGroups"

# Create an array to store the group names
$outputArray = @()

# Add distribution group names to the output array
$outputArray += "Distribution Groups:"
$outputArray += ($distGroups | Select-Object -ExpandProperty Name)

# Add security group names to the output array
$outputArray += "Security Groups:"
$outputArray += ($securityGroups | Select-Object -ExpandProperty Name)

# Export the array to a text file
$outputArray | Out-File -FilePath $outputFilePath -Encoding UTF8

Write-Host "Groups list exported to $outputFilePath"