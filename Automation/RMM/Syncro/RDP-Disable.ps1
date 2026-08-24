# RDP-Disable
# Category WEFIX

<#
Powershell to disable RDP and remove the allow firewall rule.
#>

#Disables RDP via regkey change.
Set-ItemProperty -Path 'HKLM:\System\CurrentControlSet\Control\Terminal Server' -name "fDenyTSConnections" -value 1

#Removes optional Windows Firewall Rule option.
Disable-NetFirewallRule -DisplayGroup "Remote Desktop"

# End