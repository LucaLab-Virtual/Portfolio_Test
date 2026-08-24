# Release IP cmdlets
ipconfig /release
ipconfig /renew
ipconfig /flushdns


# Disable or enable Network Adapter
# Get existing Network adapters
Get-NetAdapter | Format-List
# Disable
Disable-NetAdapter -Name Ethernet -Confirm:$false
# Enable
Enable-NetAdapter -Name Ethernet

# Add a route to the IPV4 routing table (Sample 192.172.3.0/24 /12 /32)
netsh interface ipv4 add route <pkc_ip>/24 "PKC"

# Verify the route was added successfully
netsh interface ipv4 show route | findstr "<pkc_ip>"
<#
The command netsh interface ipv4 add route <pkc_ip>/24 "PKC" 
is a Windows command-line instruction that adds a route to the IPv4 routing table. 
Here's what each part of the command does:
netsh: This is the command used to interact with network settings in Windows. It stands for "network shell".
interface ipv4: Specifies that we are working with IPv4 settings.
add route: This part of the command instructs the netsh utility to add a new route to the routing table.
<pkc_ip>/24: This is the destination network address and subnet mask. <pkc_ip> should be replaced with the actual IP address, and /24 denotes the subnet mask, indicating that the route is for all addresses within the same /24 subnet.
"PKC": This is the name or description given to the route.
#>
