# Task - Rename Local User Account
# Category WEFIX

<#
 Renames the local user account on the computer and optionally restarts it to take effect.
 #>

 Import-Module $env:SyncroModule -WarningAction SilentlyContinue

 # Chanage the local user name in Windows
 Rename-LocalUser -name "$CurrentUserName" -newname "$NewUserName"
 
 # Restart the computer for rename to take effect
 if ($ToRestartTypeYes -eq 'yes') {
     Restart-Computer -Force
 } 

 # End