<#
.SYNOPSIS
    Check and modify network interface priority (Interface Metric).

.DESCRIPTION
    Windows prefers the network interface with the LOWEST metric.
    This is useful when a Tailscale adapter is preferred over the physical Ethernet adapter.
#>

# Step 1 - Display interface priority (lower metric = higher priority)
Get-NetIPInterface |
    Sort-Object InterfaceMetric |
    Format-Table InterfaceAlias, InterfaceMetric, AddressFamily -AutoSize

# Step 2 - Display detailed interface information
Get-NetIPInterface |
    Sort-Object InterfaceMetric

# Step 3 - Set Ethernet as the preferred interface
Set-NetIPInterface -InterfaceAlias "Ethernet" -InterfaceMetric 10

# Step 4 - Lower the priority of the Tailscale interface
Set-NetIPInterface -InterfaceAlias "Tailscale" -InterfaceMetric 100

# Step 5 - Verify the changes
Get-NetIPInterface |
    Sort-Object InterfaceMetric |
    Format-Table InterfaceAlias, InterfaceMetric, AddressFamily -AutoSize
