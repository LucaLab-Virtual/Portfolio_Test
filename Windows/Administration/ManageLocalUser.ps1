# Getting Local User
Get-LocalUser
# Creating a new Local User
$password = Read-host "Create a password" -AsSecureString
New-LocalUser -Name "TestUser" -Password $password -Fullname "TestUser" -Description "Test User"
# Adding a local User to the Administrators group
Add-LocalGroupMember -Group "Administrators" -Member "TestUser"
# Get member in a local group
Get-LocalGroupMember -Group "Administrators"


# Change password of a Local User
Set-LocalUser -Name "Admin07" -Description "Description of this account."
$Password = Read-Host -AsSecureString
$UserAccount = Get-LocalUser -Name "User02"
$UserAccount | Set-LocalUser -Password $Password

Enable-LocalUser -Name "Administrator"

# Lock computer
rundll32.exe user32.dll,LockWorkStation


# SuperStaff
# IT Support Account
$password = "101%Frog1!"
$SecureStringPassword = ConvertTo-SecureString -String $password -AsPlainText -Force
New-LocalUser -Name "SuperStaff" -Password $SecureStringPassword -Fullname "SuperStaff" -Description "IT Support Account"
Add-LocalGroupMember -Group "Administrators" -Member "SuperStaff"

# User Account
$password = "Medellin2023*"
$SecureStringPassword = ConvertTo-SecureString -String $password -AsPlainText -Force
New-LocalUser -Name "SuperStaff User" -Password $SecureStringPassword -Fullname "User"
Add-LocalGroupMember -Group "Users" -Member "SuperStaff User"

<#
Verify the LocalUser is enabled GUI

Press Windows Key + R, type lusrmgr.msc, and click OK.
In the Local Users and Groups window, navigate to the Users section.
Right-click on the newly created user account and select Properties.
Make sure the Account is disabled checkbox is unchecked.


Registry Modification:

Sometimes, a registry tweak can resolve this issue.

Press Windows Key + R, type regedit, and click OK.
Navigate to the key: 
HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Authentication\LogonUI\UserSwitch.
Double-click on the option named “Enabled” on the right side.
Change the value to 1 and click OK to save the change.
Close the registry editor and reboot your PC.

The OTHER USER options is missing in the WELCOME SCREEN

Press Windows Key + R, type gpedit.msc
Navigate to:
Local Computer Policy
/Computer Configuration
/Windows Settings/
Security Settings/
Local Policy/
Security Options/
Interactive logon: Don't display last signed-in - Enabled

<<<<<<< HEAD
Press Windows Key + R, type netplwiz

Go to the "Advanced" tab and check "Require user to press Ctrl+Alt+Delete"

#>
=======
#>
>>>>>>> 2c4e4a0b108228ca67faa4399b7aef9746140847
