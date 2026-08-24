# Task - Rename Computer

<#
Renames the computer and optionally restarts it to take effect. 
You will have to change the asset name manually if you want it to match.
#>

Import-Module $env:SyncroModule -WarningAction SilentlyContinue

# Chanage the computer name in Windows
Rename-Computer -newname "$NewComputerName"

# Restart the computer for rename to take effect
if ($ToRestartTypeYes -eq 'yes') {
    Restart-Computer -Force
}

# End