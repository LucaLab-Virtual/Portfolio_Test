# RDP-Enable
# Category WEFIX

<#
Powershell to enable RDP and allow it through the firewall.
#>

#Enables RDP via regkey change.
Set-ItemProperty -Path 'HKLM:\System\CurrentControlSet\Control\Terminal Server' -name "fDenyTSConnections" -value 0

#Adds optional Windows Firewall Rule option.
Enable-NetFirewallRule -DisplayGroup "Remote Desktop"

# End