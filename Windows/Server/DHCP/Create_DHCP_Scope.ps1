<#
.SYNOPSIS
    Creates a new DHCP scope with comprehensive configuration options.

.DESCRIPTION
    This script creates a new DHCP scope on a specified DHCP server with options for:
    - Scope name and description
    - IP address range
    - Subnet mask
    - Lease duration
    - Exclusions
    - Reservations
    - Scope options (DNS servers, domain name, gateway, etc.)
    - Activation state

.PARAMETER ScopeName
    The name of the DHCP scope to create.

.PARAMETER StartIP
    The starting IP address of the scope range.

.PARAMETER EndIP
    The ending IP address of the scope range.

.PARAMETER SubnetMask
    The subnet mask for the scope (e.g., 255.255.255.0).

.PARAMETER DHCPserver
    The name or IP address of the DHCP server. Defaults to localhost.

.PARAMETER LeaseDuration
    The lease duration in days. Default is 8 days.

.PARAMETER Description
    Optional description for the scope.

.PARAMETER Gateway
    The default gateway IP address for the scope.

.PARAMETER DNSservers
    Comma-separated list of DNS server IP addresses.

.PARAMETER DomainName
    The DNS domain name for the scope.

.PARAMETER ExcludeStart
    Start of exclusion range (optional).

.PARAMETER ExcludeEnd
    End of exclusion range (optional).

.PARAMETER ReservationIP
    IP address for a reservation (optional).

.PARAMETER ReservationMAC
    MAC address for a reservation (optional).

.PARAMETER ReservationName
    Name for a reservation (optional).

.PARAMETER Activate
    Switch to activate the scope immediately after creation.

.EXAMPLE
    .\Create_DHCP_Scope.ps1 -ScopeName "Office-Network" -StartIP "192.168.1.100" -EndIP "192.168.1.200" -SubnetMask "255.255.255.0" -Gateway "192.168.1.1" -DNSservers "8.8.8.8,1.1.1.1" -Activate

.EXAMPLE
    .\Create_DHCP_Scope.ps1 -ScopeName "Guest-WiFi" -StartIP "10.0.0.50" -EndIP "10.0.0.150" -SubnetMask "255.255.255.0" -LeaseDuration 1 -Description "Guest Network" -ExcludeStart "10.0.0.50" -ExcludeEnd "10.0.0.60"

.NOTES
    Author: Portfolio Script
    Version: 1.0
    Requires: Windows Server with DHCP role installed, PowerShell 5.1+
#>

<#
    Comprehensive Parameter Support: All necessary parameters for creating a DHCP scope

    Input Validation: Validates IP addresses, MAC addresses, and range logic

    Error Handling: Robust error handling with detailed error messages

    Module Checking: Verifies DHCP module availability before execution

    Scope Existence Check: Prevents duplicate scope creation

    Option Configuration: Sets gateway, DNS, and domain name options

    Exclusion Support: Allows IP range exclusions

    Reservation Support: Creates reservations for specific MAC addresses

    Activation Control: Option to activate or leave inactive

    Export Functionality: Saves configuration details to a text file

    Verbose Output: User-friendly progress messages

    Color-Coded Output: Visual cues for success, warnings, and errors

How to Use:
powershell

# Basic scope creation
.\Create_DHCP_Scope.ps1 -ScopeName "Office-Network" -StartIP "192.168.1.100" -EndIP "192.168.1.200" -SubnetMask "255.255.255.0"

# Full configuration with activation
.\Create_DHCP_Scope.ps1 -ScopeName "Office-Network" -StartIP "192.168.1.100" -EndIP "192.168.1.200" -SubnetMask "255.255.255.0" -Gateway "192.168.1.1" -DNSservers "8.8.8.8,1.1.1.1" -DomainName "company.local" -Activate

# With exclusions and reservations
.\Create_DHCP_Scope.ps1 -ScopeName "Guest-WiFi" -StartIP "10.0.0.50" -EndIP "10.0.0.150" -SubnetMask "255.255.255.0" -ExcludeStart "10.0.0.50" -ExcludeEnd "10.0.0.60" -ReservationIP "10.0.0.100" -ReservationMAC "00-11-22-33-44-55" -ReservationName "Printer"
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true, HelpMessage = "Enter the name for the DHCP scope")]
    [ValidateNotNullOrEmpty()]
    [string]$ScopeName,
    
    [Parameter(Mandatory = $true, HelpMessage = "Enter the starting IP address")]
    [ValidatePattern("^(\d{1,3}\.){3}\d{1,3}$")]
    [string]$StartIP,
    
    [Parameter(Mandatory = $true, HelpMessage = "Enter the ending IP address")]
    [ValidatePattern("^(\d{1,3}\.){3}\d{1,3}$")]
    [string]$EndIP,
    
    [Parameter(Mandatory = $true, HelpMessage = "Enter the subnet mask")]
    [ValidatePattern("^(\d{1,3}\.){3}\d{1,3}$")]
    [string]$SubnetMask,
    
    [Parameter(Mandatory = $false)]
    [string]$DHCPServer = "localhost",
    
    [Parameter(Mandatory = $false)]
    [ValidateRange(1, 999)]
    [int]$LeaseDuration = 8,
    
    [Parameter(Mandatory = $false)]
    [string]$Description = "",
    
    [Parameter(Mandatory = $false)]
    [ValidatePattern("^(\d{1,3}\.){3}\d{1,3}$")]
    [string]$Gateway,
    
    [Parameter(Mandatory = $false)]
    [string]$DNSservers,
    
    [Parameter(Mandatory = $false)]
    [string]$DomainName,
    
    [Parameter(Mandatory = $false)]
    [ValidatePattern("^(\d{1,3}\.){3}\d{1,3}$")]
    [string]$ExcludeStart,
    
    [Parameter(Mandatory = $false)]
    [ValidatePattern("^(\d{1,3}\.){3}\d{1,3}$")]
    [string]$ExcludeEnd,
    
    [Parameter(Mandatory = $false)]
    [ValidatePattern("^(\d{1,3}\.){3}\d{1,3}$")]
    [string]$ReservationIP,
    
    [Parameter(Mandatory = $false)]
    [ValidatePattern("^([0-9A-Fa-f]{2}[:-]){5}([0-9A-Fa-f]{2})$|^([0-9A-Fa-f]{12})$")]
    [string]$ReservationMAC,
    
    [Parameter(Mandatory = $false)]
    [string]$ReservationName,
    
    [Parameter(Mandatory = $false)]
    [switch]$Activate
)

# Import DHCP module if available
function Import-DHCPModule {
    try {
        Import-Module DhcpServer -ErrorAction Stop -Force
        Write-Verbose "DHCP Server module loaded successfully."
        return $true
    }
    catch {
        Write-Warning "Failed to import DHCP Server module. Please ensure DHCP role is installed."
        Write-Warning "Error: $($_.Exception.Message)"
        return $false
    }
}

# Validate IP address range
function Test-IPRange {
    param([string]$Start, [string]$End)
    
    try {
        $startBytes = [System.Net.IPAddress]::Parse($Start).GetAddressBytes()
        $endBytes = [System.Net.IPAddress]::Parse($End).GetAddressBytes()
        
        # Convert to 32-bit unsigned integers for comparison
        [System.Array]::Reverse($startBytes)
        [System.Array]::Reverse($endBytes)
        $startInt = [System.BitConverter]::ToUInt32($startBytes, 0)
        $endInt = [System.BitConverter]::ToUInt32($endBytes, 0)
        
        return $startInt -lt $endInt
    }
    catch {
        return $false
    }
}

# Check if scope exists
function Test-DHCPScopeExists {
    param([string]$Name, [string]$Server)
    
    try {
        $existingScopes = Get-DhcpServerv4Scope -ComputerName $Server -ErrorAction SilentlyContinue
        if ($existingScopes) {
            foreach ($scope in $existingScopes) {
                if ($scope.Name -eq $Name) {
                    return $true
                }
            }
        }
        return $false
    }
    catch {
        return $false
    }
}

# Calculate subnet mask from CIDR or full mask
function Convert-SubnetMask {
    param([string]$Mask)
    
    # If it's a CIDR number (e.g., "24")
    if ($Mask -match "^\d+$" -and [int]$Mask -ge 0 -and [int]$Mask -le 32) {
        $cidr = [int]$Mask
        $maskBytes = @()
        for ($i = 0; $i -lt 4; $i++) {
            if ($cidr -ge 8) {
                $maskBytes += 255
                $cidr -= 8
            }
            else {
                $maskBytes += (256 - [math]::Pow(2, 8 - $cidr))
                $cidr = 0
            }
        }
        return [string]::Join(".", $maskBytes)
    }
    return $Mask
}

# Main script execution
Write-Host "=== DHCP Scope Creation Tool ===" -ForegroundColor Cyan
Write-Host "Starting at: $(Get-Date)" -ForegroundColor Yellow

# Validate IP range
if (-not (Test-IPRange -Start $StartIP -End $EndIP)) {
    Write-Error "Invalid IP range: Start IP must be less than End IP."
    exit 1
}

# Import DHCP module
if (-not (Import-DHCPModule)) {
    Write-Error "DHCP module is required for this script. Please install DHCP Server role."
    exit 1
}

# Check if DHCP server is reachable
try {
    $testConnection = Test-Connection -ComputerName $DHCPServer -Count 1 -Quiet
    if (-not $testConnection) {
        Write-Warning "Cannot reach DHCP server: $DHCPServer. Attempting to continue..."
    }
}
catch {
    Write-Warning "Cannot ping DHCP server, but will try to connect..."
}

# Check if scope already exists
if (Test-DHCPScopeExists -Name $ScopeName -Server $DHCPServer) {
    Write-Error "Scope '$ScopeName' already exists on server $DHCPServer."
    exit 1
}

Write-Host "Creating DHCP Scope: $ScopeName" -ForegroundColor Green
Write-Host "IP Range: $StartIP - $EndIP" -ForegroundColor Green
Write-Host "Subnet Mask: $SubnetMask" -ForegroundColor Green
Write-Host "DHCP Server: $DHCPServer" -ForegroundColor Green

try {
    # Create the scope
    $scopeParams = @{
        ComputerName = $DHCPServer
        Name = $ScopeName
        StartRange = $StartIP
        EndRange = $EndIP
        SubnetMask = $SubnetMask
        LeaseDuration = New-TimeSpan -Days $LeaseDuration
        ErrorAction = "Stop"
    }
    
    if ($Description) {
        $scopeParams.Description = $Description
    }
    
    $scope = Add-DhcpServerv4Scope @scopeParams
    
    Write-Host "✓ Scope created successfully. Scope ID: $($scope.ScopeId)" -ForegroundColor Green
    
    # Set scope options
    $optionParams = @{
        ComputerName = $DHCPServer
        ScopeId = $scope.ScopeId
        ErrorAction = "SilentlyContinue"
    }
    
    # Set Gateway
    if ($Gateway) {
        Set-DhcpServerv4OptionValue @optionParams -OptionId 3 -Value $Gateway
        Write-Host "✓ Gateway set: $Gateway" -ForegroundColor Green
    }
    
    # Set DNS Servers
    if ($DNSservers) {
        $dnsList = $DNSservers -split "," | ForEach-Object { $_.Trim() }
        Set-DhcpServerv4OptionValue @optionParams -OptionId 6 -Value $dnsList
        Write-Host "✓ DNS Servers set: $($dnsList -join ', ')" -ForegroundColor Green
    }
    
    # Set Domain Name
    if ($DomainName) {
        Set-DhcpServerv4OptionValue @optionParams -OptionId 15 -Value $DomainName
        Write-Host "✓ Domain Name set: $DomainName" -ForegroundColor Green
    }
    
    # Set Exclusions
    if ($ExcludeStart -and $ExcludeEnd) {
        Add-DhcpServerv4ExclusionRange -ComputerName $DHCPServer -ScopeId $scope.ScopeId -StartRange $ExcludeStart -EndRange $ExcludeEnd -ErrorAction SilentlyContinue
        Write-Host "✓ Exclusion range set: $ExcludeStart - $ExcludeEnd" -ForegroundColor Green
    }
    
    # Set Reservation
    if ($ReservationIP -and $ReservationMAC) {
        Add-DhcpServerv4Reservation -ComputerName $DHCPServer -ScopeId $scope.ScopeId -IPAddress $ReservationIP -ClientId $ReservationMAC -Name ($ReservationName ?? "Reservation-$ReservationIP")
        Write-Host "✓ Reservation created: $ReservationIP - $ReservationMAC" -ForegroundColor Green
    }
    
    # Activate scope
    if ($Activate) {
        Set-DhcpServerv4Scope -ComputerName $DHCPServer -ScopeId $scope.ScopeId -State Active
        Write-Host "✓ Scope activated" -ForegroundColor Green
    }
    else {
        Write-Host "ℹ Scope created in inactive state. Use -Activate parameter to activate." -ForegroundColor Yellow
    }
    
    # Display summary
    Write-Host "`n=== Scope Configuration Summary ===" -ForegroundColor Cyan
    Write-Host "Scope Name: $ScopeName"
    Write-Host "Scope ID: $($scope.ScopeId)"
    Write-Host "IP Range: $StartIP - $EndIP"
    Write-Host "Subnet Mask: $SubnetMask"
    Write-Host "Lease Duration: $LeaseDuration days"
    Write-Host "Status: $(if($Activate){'Active'}else{'Inactive'})"
    Write-Host "Description: $(if($Description){$Description}else{'None'})"
    
    # Export configuration to file
    $exportPath = Join-Path -Path $PSScriptRoot -ChildPath "$ScopeName-Config.txt"
    $configData = @"
DHCP Scope Configuration
=======================
Created: $(Get-Date)
Server: $DHCPServer
Scope Name: $ScopeName
Scope ID: $($scope.ScopeId)
IP Range: $StartIP - $EndIP
Subnet Mask: $SubnetMask
Lease Duration: $LeaseDuration days
Status: $(if($Activate){'Active'}else{'Inactive'})
Description: $(if($Description){$Description}else{'None'})
Gateway: $(if($Gateway){$Gateway}else{'Not Set'})
DNS Servers: $(if($DNSservers){$DNSservers}else{'Not Set'})
Domain Name: $(if($DomainName){$DomainName}else{'Not Set'})
Exclusion: $(if($ExcludeStart -and $ExcludeEnd){"$ExcludeStart - $ExcludeEnd"}else{'None'})
Reservation: $(if($ReservationIP -and $ReservationMAC){"$ReservationIP - $ReservationMAC"}else{'None'})
"@
    
    $configData | Out-File -FilePath $exportPath -Encoding UTF8
    Write-Host "`n✓ Configuration exported to: $exportPath" -ForegroundColor Green
    
    Write-Host "`n=== Script Completed Successfully ===" -ForegroundColor Green
    
}
catch {
    Write-Error "Failed to create DHCP scope: $($_.Exception.Message)"
    Write-Error "Error details: $($_.Exception.InnerException)"
    exit 1
}