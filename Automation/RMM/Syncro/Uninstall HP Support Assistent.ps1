# Uninstall HP Support Assistent

<#
Attempts to uninstall HP Support Assistant after a security report found it had a hacker security flaw.
#>

Import-Module $env:SyncroModule

# This displays a popup alert on the desktop.
Display-Alert -Message "SECURITY ALERT:  We need to unisntall $programname.  Please click YES to any prompts that may popup to compelete removal process"

#Uninstalls software
Get-WMIObject -class win32_product -Filter "Name like 'HP SupportAssistant'" | ForEach-Object { $_.Uninstall()}

# End