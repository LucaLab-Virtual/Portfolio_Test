# Retrieves the OS version of a Windows machine
# Get ComputerInfo
Get-ComputerInfo | Select WindowsProductName, WindowsVerion, OsHardwareAbstractionLayer

# Get-ItemProperty
(Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion").ReleaseId

# systeminfo
Systeminfo

# System.Environment
[System.Environment]::OSVersion.Version

# Get-CimInstance
(Get-CimInstance Win32_OperationSystem).Version
