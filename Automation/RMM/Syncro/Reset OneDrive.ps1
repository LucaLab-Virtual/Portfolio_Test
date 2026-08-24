# Reset OneDrive
# Category WEFIX

<#
Runs OneDrive reset via command-line and then restarts the service.
#>

Import-Module $env:SyncroModule
Log-Activity -Message "OneDrive Reset script ran" -EventName "OneDrive Reset"

# This resets the hidden OneDrive service and restarts it after a 10 second pause.
C:\Program Files\Microsoft OneDrive\onedrive.exe /reset

# End