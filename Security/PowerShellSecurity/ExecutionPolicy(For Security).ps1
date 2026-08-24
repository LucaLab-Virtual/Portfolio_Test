# Disable These 3 Windows settings (For Security)

# Remove Windows PowerShell 2.0 (Turn Windows features on or off)

# Constrained Language mode (Windows PowerShell)
$ExecutionContext.SessionState.LanguageMode
<#
Create a New System Variable (System Properties/Advanced/Environment Variables)
Variable name: PSLockDownPolicy
Variable value: 4
#> 
# Execution Policy
Get-ExecutionPolicy -List
Set-ExecutionPolicy -ExecutionPolicy Bypass -Scope Process
<#
Setting up Execution Policy 
Windows 10 Pro
(Local Group Policy Editor/Computer Configuration/
Administrative Templates/Windows Components/Windows PowerShell/
Turn on Script Execution)

Windows home edition
(Policy Plus)

PowerShell 7 and latest
Download the latest version, sample "PowerShell-7.3.6-win-x64.zip"
PowerShellCoreExecutionPolicy.adml Copy to C:\Windows\PolicyDefinitions\en-US
PowerShellCoreExecutionPolicy.admx copy to C:\Windows\PolicyDefinitions
(/Administrative Templates/PowerShell Core/Turn on Script Execution)

Disabled = it'll block any user with admin right to run Powershell Scripts
#>
