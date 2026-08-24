# First, install the QRCodeGenerator module if you haven't already
Install-Module -Name QRCodeGenerator

# Import the module
Import-Module QRCodeGenerator

# Define the URL you want to encode
$url = "https://1dps0q-my.sharepoint.com/:v:/g/personal/admin_online_1dps0q_onmicrosoft_com/EUZxcwXTsAFMuvCq4OmDjyQBgma4jC4lVA-q_MDKDBsVbg?e=3rRg5K"

# Generate the QR code
New-QRCodeURI -URI $url -OutPath "C:\Users\Lucarez\Documents\VideoPresentationQRCode.png"

# Display the path to the generated QR code
Write-Host "QR code generated: VideoPresentationQRCode.png"

#---------------------------------------------------------------------------------------------------------------------------------------------------#

# Check Module Version
Get-Module QRCodeGenerator -ListAvailable

# Verify Module Content
Get-Command -Module QRCodeGenerator

# Reinstall Module
Uninstall-Module -Name QRCodeGenerator
Install-Module -Name QRCodeGenerator