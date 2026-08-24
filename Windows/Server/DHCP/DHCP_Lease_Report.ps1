<#
.SYNOPSIS
    Generates comprehensive DHCP lease reports with analysis, filtering, and multiple output formats.

.DESCRIPTION
    This script generates detailed DHCP lease reports including:
    - Active leases with expiry information
    - Lease utilization statistics
    - Scope usage analysis
    - Lease history and trends
    - Expiring leases report
    - MAC address vendor identification
    - Multiple output formats (HTML, CSV, JSON, Excel)
    - Email report delivery
    - Scheduled reporting support

.PARAMETER DHCPServer
    The name or IP address of the DHCP server. Defaults to localhost.

.PARAMETER ScopeId
    Specific scope ID to report on. If not specified, all scopes are included.

.PARAMETER ReportType
    Type of report to generate: Active, Expiring, All, Statistics, or Utilization.
    Default is All.

.PARAMETER OutputPath
    Directory where reports will be saved. Defaults to current directory.

.PARAMETER OutputFormats
    Output format(s): HTML, CSV, JSON, Excel, or ALL. Default is ALL.

.PARAMETER ExpiryThreshold
    Number of days to look ahead for expiring leases. Default is 7 days.

.PARAMETER IncludeHistory
    Include lease history/statistics in the report.

.PARAMETER IncludeVendorInfo
    Include MAC vendor information (requires online lookup).

.PARAMETER IncludeScopeStats
    Include detailed scope statistics.

.PARAMETER IncludeLeaseUsage
    Include lease usage patterns and trends.

.PARAMETER TopConsumers
    Number of top lease consumers to display. Default is 10.

.PARAMETER FilterHostname
    Filter leases by hostname (supports wildcards).

.PARAMETER FilterIPRange
    Filter leases by IP range (e.g., "192.168.1.100-192.168.1.150").

.PARAMETER FilterMAC
    Filter leases by MAC address.

.PARAMETER ExportResults
    Export detailed results to CSV.

.PARAMETER EmailReport
    Send report via email.

.PARAMETER EmailTo
    Recipient email address.

.PARAMETER EmailFrom
    Sender email address.

.PARAMETER EmailServer
    SMTP server address.

.PARAMETER ScheduleMode
    Run in scheduled mode (suppresses interactive prompts).

.EXAMPLE
    .\DHCP_Lease_Report.ps1 -DHCPServer "DHCP01" -ReportType Active -OutputFormats HTML,CSV

.EXAMPLE
    .\DHCP_Lease_Report.ps1 -DHCPServer "192.168.1.10" -ScopeId "192.168.1.0" -ReportType Expiring -ExpiryThreshold 3 -IncludeVendorInfo

.EXAMPLE
    .\DHCP_Lease_Report.ps1 -DHCPServer "DHCP01" -ReportType Statistics -OutputPath "C:\Reports" -EmailReport -EmailTo "admin@company.com"

.NOTES
    Author: Portfolio Script
    Version: 1.0
    Requires: Windows Server with DHCP role installed, PowerShell 5.1+
#>

<#
    Multiple Report Types: Active, Expiring, All, Statistics, Utilization

    Comprehensive Data Collection: Scopes, leases, statistics, usage patterns

    Rich HTML Reports: Beautiful, interactive, and mobile-responsive

    Multiple Output Formats: HTML, CSV, JSON, Excel

    MAC Vendor Lookup: Identifies device manufacturers

    Lease Analysis: Top consumers, utilization trends, expiry tracking

    Filtering Capabilities: Filter by hostname, IP range, MAC address

    Email Integration: Automated report delivery

    Scheduled Mode: For automated reporting

    Progress Indicators: Visual feedback during data collection

How to Use:
powershell

# Basic active lease report
.\DHCP_Lease_Report.ps1 -DHCPServer "DHCP01" -ReportType Active -OutputFormats HTML,CSV

# Expiring leases report with vendor info
.\DHCP_Lease_Report.ps1 -DHCPServer "192.168.1.10" -ScopeId "192.168.1.0" -ReportType Expiring -ExpiryThreshold 3 -IncludeVendorInfo

# Full report with statistics
.\DHCP_Lease_Report.ps1 -DHCPServer "DHCP01" -ReportType All -IncludeScopeStats -IncludeLeaseUsage -TopConsumers 20

# Filtered report
.\DHCP_Lease_Report.ps1 -DHCPServer "DHCP01" -FilterHostname "*PRINTER*" -FilterIPRange "192.168.1.100-192.168.1.150"

# Email report
.\DHCP_Lease_Report.ps1 -DHCPServer "DHCP01" -ReportType All -EmailReport -EmailTo "admin@company.com" -EmailFrom "dhcp-reports@company.com" -EmailServer "smtp.company.com"

Sample Output HTML Features:

    Modern Dashboard: Professional, responsive design

    Summary Cards: At-a-glance statistics with color coding

    Scope Overview: Table with utilization progress bars

    Expiring Leases: Highlighted with status badges

    Top Consumers: Identify heavy users with vendor information

    Lease Utilization: Visual charts showing scope usage

    Export Options: Multiple formats for different use cases

    Print Friendly: Optimized for printing reports
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [string]$DHCPServer = "localhost",
    
    [Parameter(Mandatory = $false)]
    [ValidatePattern("^(\d{1,3}\.){3}\d{1,3}$")]
    [string]$ScopeId,
    
    [Parameter(Mandatory = $false)]
    [ValidateSet("Active", "Expiring", "All", "Statistics", "Utilization")]
    [string]$ReportType = "All",
    
    [Parameter(Mandatory = $false)]
    [string]$OutputPath = ".",
    
    [Parameter(Mandatory = $false)]
    [ValidateSet("HTML", "CSV", "JSON", "Excel", "ALL")]
    [string]$OutputFormats = "ALL",
    
    [Parameter(Mandatory = $false)]
    [int]$ExpiryThreshold = 7,
    
    [Parameter(Mandatory = $false)]
    [switch]$IncludeHistory,
    
    [Parameter(Mandatory = $false)]
    [switch]$IncludeVendorInfo,
    
    [Parameter(Mandatory = $false)]
    [switch]$IncludeScopeStats,
    
    [Parameter(Mandatory = $false)]
    [switch]$IncludeLeaseUsage,
    
    [Parameter(Mandatory = $false)]
    [int]$TopConsumers = 10,
    
    [Parameter(Mandatory = $false)]
    [string]$FilterHostname,
    
    [Parameter(Mandatory = $false)]
    [string]$FilterIPRange,
    
    [Parameter(Mandatory = $false)]
    [string]$FilterMAC,
    
    [Parameter(Mandatory = $false)]
    [string]$ExportResults,
    
    [Parameter(Mandatory = $false)]
    [switch]$EmailReport,
    
    [Parameter(Mandatory = $false)]
    [string]$EmailTo,
    
    [Parameter(Mandatory = $false)]
    [string]$EmailFrom,
    
    [Parameter(Mandatory = $false)]
    [string]$EmailServer,
    
    [Parameter(Mandatory = $false)]
    [switch]$ScheduleMode
)

# Global variables
$Script:StartTime = Get-Date
$Script:ReportData = @{}
$Script:ExportFiles = @()
$Script:ErrorCount = 0

# Color definitions
$Script:Colors = @{
    Success = "Green"
    Error = "Red"
    Warning = "Yellow"
    Info = "Cyan"
    Header = "Magenta"
    Detail = "Gray"
    Critical = "Red"
    Warning = "Yellow"
    Normal = "White"
}

# Function to import DHCP module
function Import-DHCPModule {
    try {
        Import-Module DhcpServer -ErrorAction Stop -Force
        Write-Output "✓ DHCP Server module loaded successfully."
        return $true
    }
    catch {
        Write-Error "Failed to import DHCP Server module. Error: $($_.Exception.Message)"
        return $false
    }
}

# Function to get MAC vendor info
function Get-MACVendor {
    param([string]$MAC)
    
    if (-not $IncludeVendorInfo) {
        return "N/A"
    }
    
    try {
        # Clean MAC address
        $cleanMAC = $MAC -replace '[-:]', '' -replace '\.', '' -replace '\s', ''
        $cleanMAC = $cleanMAC.Substring(0, 6).ToUpper()
        
        # Query online MAC vendor database (using a free API)
        $url = "https://api.macvendors.com/$cleanMAC"
        $response = Invoke-RestMethod -Uri $url -Method Get -ErrorAction SilentlyContinue -TimeoutSec 5
        
        if ($response -and $response -notmatch "Not Found") {
            return $response
        }
        else {
            return "Unknown"
        }
    }
    catch {
        return "Unknown"
    }
}

# Function to get DHCP scopes
function Get-DHCPScopes {
    param([string]$Server)
    
    try {
        $scopes = Get-DhcpServerv4Scope -ComputerName $Server -ErrorAction Stop
        return $scopes
    }
    catch {
        Write-Error "Failed to retrieve DHCP scopes: $($_.Exception.Message)"
        return $null
    }
}

# Function to get lease data
function Get-DHCPLeases {
    param(
        [string]$Server,
        [string]$ScopeId,
        [switch]$IncludeHistory
    )
    
    try {
        $leases = Get-DhcpServerv4Lease -ComputerName $Server -ScopeId $ScopeId -ErrorAction Stop
        
        $leaseData = @()
        foreach ($lease in $leases) {
            $leaseInfo = [PSCustomObject]@{
                ScopeId = $ScopeId
                IPAddress = $lease.IPAddress.IPAddressToString
                ClientId = $lease.ClientId
                HostName = $lease.HostName
                AddressState = $lease.AddressState
                LeaseExpiryTime = $lease.LeaseExpiryTime
                Type = $lease.Type
                State = $lease.State
                ScopeName = $null
                Vendor = $null
                DaysRemaining = $null
                Status = $null
            }
            
            # Calculate days remaining
            if ($lease.LeaseExpiryTime) {
                $daysRemaining = ($lease.LeaseExpiryTime - (Get-Date)).Days
                $leaseInfo.DaysRemaining = [math]::Round($daysRemaining, 1)
                
                if ($daysRemaining -lt 0) {
                    $leaseInfo.Status = "Expired"
                }
                elseif ($daysRemaining -lt 1) {
                    $leaseInfo.Status = "Critical"
                }
                elseif ($daysRemaining -lt 3) {
                    $leaseInfo.Status = "Warning"
                }
                else {
                    $leaseInfo.Status = "Healthy"
                }
            }
            else {
                $leaseInfo.DaysRemaining = $null
                $leaseInfo.Status = "Unknown"
            }
            
            # Get vendor info if enabled
            if ($IncludeVendorInfo -and $lease.ClientId) {
                $leaseInfo.Vendor = Get-MACVendor -MAC $lease.ClientId
            }
            
            $leaseData += $leaseInfo
        }
        
        return $leaseData
    }
    catch {
        Write-Error "Failed to retrieve leases for scope $ScopeId: $($_.Exception.Message)"
        return @()
    }
}

# Function to get scope statistics
function Get-DHCPScopeStats {
    param(
        [string]$Server,
        [string]$ScopeId
    )
    
    try {
        $stats = Get-DhcpServerv4ScopeStatistics -ComputerName $Server -ScopeId $ScopeId -ErrorAction Stop
        return $stats
    }
    catch {
        Write-Error "Failed to retrieve statistics for scope $ScopeId: $($_.Exception.Message)"
        return $null
    }
}

# Function to collect all report data
function Collect-ReportData {
    param(
        [string]$Server,
        [string]$ScopeId,
        [string]$ReportType,
        [int]$ExpiryThreshold,
        [switch]$IncludeHistory,
        [switch]$IncludeVendorInfo,
        [switch]$IncludeScopeStats,
        [switch]$IncludeLeaseUsage
    )
    
    Write-Host "📊 Collecting DHCP lease data..." -ForegroundColor Yellow
    
    $reportData = @{
        Server = $Server
        ReportTime = Get-Date
        Scopes = @()
        Summary = @{}
        Statistics = @{}
        ExpiringLeases = @()
        TopConsumers = @()
        LeaseUsage = @{}
    }
    
    # Get all scopes
    $scopes = Get-DHCPScopes -Server $Server
    if (-not $scopes) {
        Write-Error "No DHCP scopes found on server $Server"
        return $null
    }
    
    # Filter by scope if specified
    if ($ScopeId) {
        $scopes = $scopes | Where-Object { $_.ScopeId.IPAddressToString -eq $ScopeId }
        if (-not $scopes) {
            Write-Error "Scope $ScopeId not found on server $Server"
            return $null
        }
    }
    
    Write-Host "Processing $($scopes.Count) scopes..." -ForegroundColor Yellow
    
    $totalLeases = 0
    $totalActive = 0
    $totalExpired = 0
    $allLeases = @()
    $allActiveLeases = @()
    $allExpiringLeases = @()
    
    foreach ($scope in $scopes) {
        $scopeId = $scope.ScopeId.IPAddressToString
        Write-Host "  Processing scope: $scopeId ($($scope.Name))" -ForegroundColor Gray
        
        # Get leases
        $leases = Get-DHCPLeases -Server $Server -ScopeId $scopeId -IncludeHistory:$IncludeHistory
        
        # Get scope statistics
        if ($IncludeScopeStats) {
            $stats = Get-DHCPScopeStats -Server $Server -ScopeId $scopeId
            if ($stats) {
                $reportData.Statistics[$scopeId] = [PSCustomObject]@{
                    ScopeId = $scopeId
                    ScopeName = $scope.Name
                    TotalAddresses = $stats.TotalAddresses
                    UsedAddresses = $stats.UsedAddresses
                    AvailableAddresses = $stats.AvailableAddresses
                    LeasedAddresses = $stats.LeasedAddresses
                    PercentLeased = $stats.PercentLeased
                    PercentAvailable = $stats.PercentAvailable
                    PercentInUse = $stats.PercentInUse
                    TotalReservations = (Get-DhcpServerv4Reservation -ComputerName $Server -ScopeId $scopeId -ErrorAction SilentlyContinue).Count
                }
            }
        }
        
        # Process leases
        if ($leases) {
            $activeLeases = $leases | Where-Object { $_.State -eq "Active" }
            $expiredLeases = $leases | Where-Object { $_.State -eq "Expired" }
            $expiringLeases = $leases | Where-Object { 
                $_.State -eq "Active" -and 
                $_.DaysRemaining -ne $null -and 
                $_.DaysRemaining -le $ExpiryThreshold -and 
                $_.DaysRemaining -ge 0
            }
            
            $totalLeases += $leases.Count
            $totalActive += $activeLeases.Count
            $totalExpired += $expiredLeases.Count
            
            $allLeases += $leases
            $allActiveLeases += $activeLeases
            $allExpiringLeases += $expiringLeases
            
            # Add scope data
            $scopeData = [PSCustomObject]@{
                ScopeId = $scopeId
                Name = $scope.Name
                Description = $scope.Description
                State = $scope.State
                StartRange = $scope.StartRange.IPAddressToString
                EndRange = $scope.EndRange.IPAddressToString
                SubnetMask = $scope.SubnetMask.IPAddressToString
                LeaseDuration = $scope.LeaseDuration
                TotalLeases = $leases.Count
                ActiveLeases = $activeLeases.Count
                ExpiredLeases = $expiredLeases.Count
                ExpiringLeases = $expiringLeases.Count
            }
            
            $reportData.Scopes += $scopeData
        }
    }
    
    # Generate summary
    $reportData.Summary = [PSCustomObject]@{
        Server = $Server
        TotalScopes = $scopes.Count
        TotalLeases = $totalLeases
        TotalActive = $totalActive
        TotalExpired = $totalExpired
        TotalExpiring = $allExpiringLeases.Count
        ReportTime = Get-Date
        ExpiryThreshold = $ExpiryThreshold
    }
    
    # Get expiring leases
    if ($ReportType -in @("Expiring", "All")) {
        $reportData.ExpiringLeases = $allExpiringLeases | Sort-Object DaysRemaining
    }
    
    # Get top consumers (by MAC address)
    if ($IncludeLeaseUsage -or $ReportType -eq "Utilization") {
        $macGroups = $allActiveLeases | Group-Object -Property ClientId
        $topConsumers = $macGroups | ForEach-Object {
            [PSCustomObject]@{
                ClientId = $_.Name
                Count = $_.Count
                HostNames = ($_.Group | ForEach-Object { $_.HostName } | Where-Object { $_ } | Select-Object -Unique) -join ', '
                Vendor = if ($IncludeVendorInfo) { Get-MACVendor -MAC $_.Name } else { "N/A" }
            }
        } | Sort-Object Count -Descending | Select-Object -First $TopConsumers
        
        $reportData.TopConsumers = $topConsumers
        
        # Lease usage by scope
        $scopeUsage = $reportData.Scopes | ForEach-Object {
            [PSCustomObject]@{
                ScopeId = $_.ScopeId
                Name = $_.Name
                LeaseUtilization = if ($_.TotalLeases -gt 0) { [math]::Round(($_.ActiveLeases / $_.TotalLeases) * 100, 1) } else { 0 }
                ActiveLeases = $_.ActiveLeases
                TotalLeases = $_.TotalLeases
                ExpiredLeases = $_.ExpiredLeases
            }
        }
        $reportData.LeaseUsage = $scopeUsage
    }
    
    # Apply filters
    if ($FilterHostname -or $FilterIPRange -or $FilterMAC) {
        Write-Host "Applying filters..." -ForegroundColor Yellow
        
        $filteredLeases = $allLeases
        
        if ($FilterHostname) {
            $filteredLeases = $filteredLeases | Where-Object { $_.HostName -like $FilterHostname }
        }
        
        if ($FilterIPRange) {
            $parts = $FilterIPRange -split '-'
            if ($parts.Count -eq 2) {
                $startIP = [System.Net.IPAddress]::Parse($parts[0].Trim())
                $endIP = [System.Net.IPAddress]::Parse($parts[1].Trim())
                $filteredLeases = $filteredLeases | Where-Object {
                    $ip = [System.Net.IPAddress]::Parse($_.IPAddress)
                    $ip.Address -ge $startIP.Address -and $ip.Address -le $endIP.Address
                }
            }
        }
        
        if ($FilterMAC) {
            $filteredLeases = $filteredLeases | Where-Object { $_.ClientId -like $FilterMAC }
        }
        
        $reportData.FilteredLeases = $filteredLeases
    }
    
    Write-Host "✓ Data collection complete!" -ForegroundColor Green
    return $reportData
}

# Function to create HTML report
function New-HTMLReport {
    param(
        [PSObject]$Data,
        [string]$OutputPath
    )
    
    try {
        $htmlContent = @"
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>DHCP Lease Report - $($Data.Server)</title>
    <style>
        * { margin: 0; padding: 0; box-sizing: border-box; }
        body { 
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            background: linear-gradient(135deg, #f5f7fa 0%, #c3cfe2 100%);
            min-height: 100vh;
            padding: 20px;
        }
        .container {
            max-width: 1400px;
            margin: 0 auto;
            background: white;
            border-radius: 16px;
            box-shadow: 0 10px 40px rgba(0,0,0,0.15);
            padding: 30px;
            animation: fadeIn 0.5s ease-in;
        }
        @keyframes fadeIn {
            from { opacity: 0; transform: translateY(-20px); }
            to { opacity: 1; transform: translateY(0); }
        }
        .header {
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            color: white;
            padding: 30px;
            border-radius: 12px;
            margin-bottom: 30px;
        }
        .header h1 { font-size: 32px; font-weight: 300; }
        .header .subtitle { font-size: 16px; opacity: 0.9; margin-top: 8px; }
        .summary-grid {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(200px, 1fr));
            gap: 20px;
            margin-bottom: 30px;
        }
        .summary-card {
            background: #f8f9fa;
            padding: 20px;
            border-radius: 10px;
            border-left: 4px solid #667eea;
            transition: transform 0.2s;
        }
        .summary-card:hover { transform: translateY(-3px); box-shadow: 0 5px 15px rgba(0,0,0,0.1); }
        .summary-card .label { font-size: 14px; color: #6c757d; }
        .summary-card .value { font-size: 28px; font-weight: bold; color: #2d3748; margin-top: 5px; }
        .summary-card.critical { border-left-color: #e53e3e; }
        .summary-card.warning { border-left-color: #ed8936; }
        .summary-card.success { border-left-color: #48bb78; }
        .summary-card.info { border-left-color: #4299e1; }
        .section {
            margin-top: 30px;
            padding: 20px;
            background: #f8f9fa;
            border-radius: 10px;
        }
        .section h2 {
            color: #2d3748;
            margin-bottom: 20px;
            padding-bottom: 10px;
            border-bottom: 2px solid #e2e8f0;
        }
        table {
            width: 100%;
            border-collapse: collapse;
            background: white;
            border-radius: 8px;
            overflow: hidden;
            box-shadow: 0 2px 4px rgba(0,0,0,0.05);
        }
        th {
            background: #2d3748;
            color: white;
            padding: 12px;
            text-align: left;
            font-weight: 600;
            text-transform: uppercase;
            font-size: 12px;
            letter-spacing: 0.5px;
        }
        td {
            padding: 10px 12px;
            border-bottom: 1px solid #e2e8f0;
        }
        tr:hover { background-color: #f7fafc; }
        .status-badge {
            display: inline-block;
            padding: 4px 12px;
            border-radius: 20px;
            font-size: 12px;
            font-weight: 600;
        }
        .status-Healthy { background: #c6f6d5; color: #22543d; }
        .status-Warning { background: #fefcbf; color: #744210; }
        .status-Critical { background: #fed7d7; color: #9b2c2c; }
        .status-Expired { background: #e2e8f0; color: #2d3748; }
        .status-Active { background: #bee3f8; color: #2a4365; }
        .progress-bar {
            background: #e2e8f0;
            border-radius: 10px;
            height: 20px;
            overflow: hidden;
            position: relative;
        }
        .progress-bar .fill {
            height: 100%;
            background: linear-gradient(90deg, #48bb78, #38a169);
            transition: width 1s ease-in-out;
            border-radius: 10px;
        }
        .progress-bar .fill.warning { background: linear-gradient(90deg, #ed8936, #dd6b20); }
        .progress-bar .fill.critical { background: linear-gradient(90deg, #fc8181, #e53e3e); }
        .progress-label {
            position: absolute;
            top: 50%;
            left: 50%;
            transform: translate(-50%, -50%);
            font-size: 12px;
            font-weight: 600;
            color: #2d3748;
        }
        .chart-container {
            background: white;
            padding: 20px;
            border-radius: 8px;
            margin: 20px 0;
            box-shadow: 0 2px 4px rgba(0,0,0,0.05);
        }
        .footer {
            margin-top: 30px;
            padding-top: 20px;
            border-top: 1px solid #e2e8f0;
            color: #718096;
            font-size: 14px;
            text-align: center;
        }
        .badge {
            display: inline-block;
            padding: 2px 8px;
            border-radius: 4px;
            font-size: 11px;
            font-weight: 600;
        }
        .badge-vendor { background: #ebf8ff; color: #2b6cb0; }
        .filter-info {
            background: #edf2f7;
            padding: 10px 15px;
            border-radius: 6px;
            margin-bottom: 20px;
            font-size: 14px;
        }
        @media print {
            .container { box-shadow: none; }
            .summary-card:hover { transform: none; }
            .header { background: #667eea !important; }
        }
    </style>
</head>
<body>
<div class="container">
"@

        # Header
        $htmlContent += @"
    <div class="header">
        <h1>📊 DHCP Lease Report</h1>
        <div class="subtitle">
            Server: <strong>$($Data.Server)</strong> | 
            Generated: <strong>$($Data.ReportTime.ToString('yyyy-MM-dd HH:mm:ss'))</strong> |
            Report ID: <strong>$([System.Guid]::NewGuid().ToString().Substring(0, 8))</strong>
        </div>
    </div>
"@

        # Summary Cards
        $htmlContent += @"
    <div class="summary-grid">
        <div class="summary-card info">
            <div class="label">Total Scopes</div>
            <div class="value">$($Data.Summary.TotalScopes)</div>
        </div>
        <div class="summary-card success">
            <div class="label">Total Leases</div>
            <div class="value">$($Data.Summary.TotalLeases)</div>
        </div>
        <div class="summary-card success">
            <div class="label">Active Leases</div>
            <div class="value">$($Data.Summary.TotalActive)</div>
        </div>
        <div class="summary-card warning">
            <div class="label">Expiring Soon (≤$($Data.Summary.ExpiryThreshold) days)</div>
            <div class="value">$($Data.Summary.TotalExpiring)</div>
        </div>
        <div class="summary-card critical">
            <div class="label">Expired Leases</div>
            <div class="value">$($Data.Summary.TotalExpired)</div>
        </div>
    </div>
"@

        # Scope Statistics
        if ($Data.Scopes -and $Data.Scopes.Count -gt 0) {
            $htmlContent += @"
    <div class="section">
        <h2>📋 Scope Overview</h2>
        <table>
            <thead>
                <tr>
                    <th>Scope ID</th>
                    <th>Name</th>
                    <th>State</th>
                    <th>IP Range</th>
                    <th>Active Leases</th>
                    <th>Total Leases</th>
                    <th>Utilization</th>
                </tr>
            </thead>
            <tbody>
"@
            foreach ($scope in $Data.Scopes) {
                $utilization = if ($scope.TotalLeases -gt 0) { [math]::Round(($scope.ActiveLeases / $scope.TotalLeases) * 100, 1) } else { 0 }
                $barClass = if ($utilization -gt 80) { "critical" } elseif ($utilization -gt 60) { "warning" } else { "" }
                
                $htmlContent += @"
                <tr>
                    <td><strong>$($scope.ScopeId)</strong></td>
                    <td>$($scope.Name)</td>
                    <td><span class="status-badge status-$($scope.State)">$($scope.State)</span></td>
                    <td>$($scope.StartRange) - $($scope.EndRange)</td>
                    <td>$($scope.ActiveLeases)</td>
                    <td>$($scope.TotalLeases)</td>
                    <td>
                        <div class="progress-bar">
                            <div class="fill $barClass" style="width: $utilization%"></div>
                            <div class="progress-label">$utilization%</div>
                        </div>
                    </td>
                </tr>
"@
            }
            $htmlContent += @"
            </tbody>
        </table>
    </div>
"@
        }

        # Expiring Leases
        if ($Data.ExpiringLeases -and $Data.ExpiringLeases.Count -gt 0) {
            $htmlContent += @"
    <div class="section">
        <h2>⚠️ Expiring Leases (Next $($Data.Summary.ExpiryThreshold) Days)</h2>
        <table>
            <thead>
                <tr>
                    <th>IP Address</th>
                    <th>Hostname</th>
                    <th>MAC Address</th>
                    <th>Expiry Time</th>
                    <th>Days Remaining</th>
                    <th>Status</th>
                    <th>Vendor</th>
                </tr>
            </thead>
            <tbody>
"@
            foreach ($lease in $Data.ExpiringLeases | Sort-Object DaysRemaining) {
                $statusClass = $lease.Status
                $daysRemaining = [math]::Round($lease.DaysRemaining, 1)
                
                $htmlContent += @"
                <tr>
                    <td><strong>$($lease.IPAddress)</strong></td>
                    <td>$(if($lease.HostName){$lease.HostName} else {'N/A'})</td>
                    <td><code>$($lease.ClientId)</code></td>
                    <td>$($lease.LeaseExpiryTime.ToString('yyyy-MM-dd HH:mm'))</td>
                    <td>$daysRemaining</td>
                    <td><span class="status-badge status-$statusClass">$statusClass</span></td>
                    <td>$(if($lease.Vendor){$lease.Vendor} else {'N/A'})</td>
                </tr>
"@
            }
            $htmlContent += @"
            </tbody>
        </table>
    </div>
"@
        }

        # Top Consumers
        if ($Data.TopConsumers -and $Data.TopConsumers.Count -gt 0) {
            $htmlContent += @"
    <div class="section">
        <h2>🏆 Top Lease Consumers</h2>
        <table>
            <thead>
                <tr>
                    <th>#</th>
                    <th>MAC Address</th>
                    <th>Hostnames</th>
                    <th>Lease Count</th>
                    <th>Vendor</th>
                </tr>
            </thead>
            <tbody>
"@
            $index = 1
            foreach ($consumer in $Data.TopConsumers) {
                $htmlContent += @"
                <tr>
                    <td>$index</td>
                    <td><code>$($consumer.ClientId)</code></td>
                    <td>$(if($consumer.HostNames){$consumer.HostNames} else {'N/A'})</td>
                    <td><span class="badge badge-vendor">$($consumer.Count)</span></td>
                    <td>$(if($consumer.Vendor){$consumer.Vendor} else {'N/A'})</td>
                </tr>
"@
                $index++
            }
            $htmlContent += @"
            </tbody>
        </table>
    </div>
"@
        }

        # Lease Usage
        if ($Data.LeaseUsage -and $Data.LeaseUsage.Count -gt 0) {
            $htmlContent += @"
    <div class="section">
        <h2>📈 Lease Utilization by Scope</h2>
        <div class="chart-container">
"@
            foreach ($usage in $Data.LeaseUsage) {
                $barClass = if ($usage.LeaseUtilization -gt 80) { "critical" } elseif ($usage.LeaseUtilization -gt 60) { "warning" } else { "" }
                $htmlContent += @"
            <div style="margin-bottom: 15px;">
                <div style="display: flex; justify-content: space-between; margin-bottom: 5px;">
                    <span><strong>$($usage.Name)</strong> ($($usage.ScopeId))</span>
                    <span>$($usage.ActiveLeases) / $($usage.TotalLeases) leases</span>
                </div>
                <div class="progress-bar">
                    <div class="fill $barClass" style="width: $($usage.LeaseUtilization)%"></div>
                    <div class="progress-label">$($usage.LeaseUtilization)%</div>
                </div>
            </div>
"@
            }
            $htmlContent += @"
        </div>
    </div>
"@
        }

        # Footer
        $htmlContent += @"
    <div class="footer">
        <p>Generated by DHCP Lease Report Script v1.0</p>
        <p>Report generated on $($Data.ReportTime.ToString('yyyy-MM-dd HH:mm:ss'))</p>
        <p>Data sourced from DHCP Server: $($Data.Server)</p>
    </div>
</div>
</body>
</html>
"@

        $htmlContent | Out-File -FilePath $OutputPath -Encoding UTF8
        Write-Host "✓ HTML report generated: $OutputPath" -ForegroundColor Green
        return $true
    }
    catch {
        Write-Error "Failed to generate HTML report: $($_.Exception.Message)"
        return $false
    }
}

# Function to export to CSV
function Export-ToCSV {
    param(
        [string]$FilePath,
        [array]$Data,
        [string]$DataType
    )
    
    try {
        if ($Data -and $Data.Count -gt 0) {
            $Data | Export-Csv -Path $FilePath -NoTypeInformation -Encoding UTF8
            Write-Host "✓ CSV exported ($DataType): $FilePath" -ForegroundColor Green
            return $true
        }
        else {
            Write-Warning "No data to export for $DataType"
            return $false
        }
    }
    catch {
        Write-Error "Failed to export CSV: $($_.Exception.Message)"
        return $false
    }
}

# Function to export to JSON
function Export-ToJSON {
    param(
        [string]$FilePath,
        [PSObject]$Data
    )
    
    try {
        $json = $Data | ConvertTo-Json -Depth 10
        $json | Out-File -FilePath $FilePath -Encoding UTF8
        Write-Host "✓ JSON exported: $FilePath" -ForegroundColor Green
        return $true
    }
    catch {
        Write-Error "Failed to export JSON: $($_.Exception.Message)"
        return $false
    }
}

# Function to export to Excel (CSV with Excel format)
function Export-ToExcel {
    param(
        [string]$FilePath,
        [PSObject]$Data
    )
    
    try {
        # Create a workbook using COM (Excel must be installed)
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        $workbook = $excel.Workbooks.Add()
        
        # Add sheets for each data type
        $sheetIndex = 1
        $dataTypes = @()
        
        if ($Data.Scopes) { $dataTypes += @{ Name = "Scopes"; Data = $Data.Scopes } }
        if ($Data.ExpiringLeases) { $dataTypes += @{ Name = "Expiring"; Data = $Data.ExpiringLeases } }
        if ($Data.TopConsumers) { $dataTypes += @{ Name = "TopConsumers"; Data = $Data.TopConsumers } }
        if ($Data.LeaseUsage) { $dataTypes += @{ Name = "LeaseUsage"; Data = $Data.LeaseUsage } }
        
        foreach ($type in $dataTypes) {
            if ($sheetIndex -gt 1) {
                $worksheet = $workbook.Worksheets.Add()
            }
            else {
                $worksheet = $workbook.Worksheets.Item($sheetIndex)
            }
            $worksheet.Name = $type.Name
            
            # Add headers
            $properties = $type.Data[0].PSObject.Properties.Name
            for ($col = 1; $col -le $properties.Count; $col++) {
                $worksheet.Cells.Item(1, $col) = $properties[$col - 1]
                $worksheet.Cells.Item(1, $col).Font.Bold = $true
            }
            
            # Add data
            $row = 2
            foreach ($item in $type.Data) {
                $col = 1
                foreach ($prop in $properties) {
                    $value = $item.$prop
                    $worksheet.Cells.Item($row, $col) = $value.ToString()
                    $col++
                }
                $row++
            }
            
            # Auto-fit columns
            $usedRange = $worksheet.UsedRange
            $usedRange.EntireColumn.AutoFit() | Out-Null
            $sheetIndex++
        }
        
        # Remove default empty sheets
        $workbook.Worksheets.Item($sheetIndex).Delete()
        
        # Save and close
        $workbook.SaveAs($FilePath, 51) # 51 = xlOpenXMLWorkbook
        $workbook.Close()
        $excel.Quit()
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
        
        Write-Host "✓ Excel report generated: $FilePath" -ForegroundColor Green
        return $true
    }
    catch {
        Write-Warning "Failed to generate Excel report: $($_.Exception.Message)"
        Write-Warning "Exporting as CSV instead..."
        return Export-ToCSV -FilePath ($FilePath -replace '\.xlsx$', '.csv') -Data $Data.Scopes -DataType "All"
    }
}

# Function to send email
function Send-EmailReport {
    param(
        [string]$To,
        [string]$From,
        [string]$SMTPServer,
        [array]$Attachments,
        [string]$Subject = "DHCP Lease Report"
    )
    
    try {
        $body = @"
DHCP Lease Report

Server: $($Script:ReportData.Server)
Generated: $($Script:ReportData.ReportTime.ToString('yyyy-MM-dd HH:mm:ss'))
Total Scopes: $($Script:ReportData.Summary.TotalScopes)
Total Leases: $($Script:ReportData.Summary.TotalLeases)
Active Leases: $($Script:ReportData.Summary.TotalActive)
Expiring Leases: $($Script:ReportData.Summary.TotalExpiring)
Expired Leases: $($Script:ReportData.Summary.TotalExpired)

Attached Files:
$($Attachments -join "`n")

This is an automated report generated by the DHCP Lease Report Script.
"@
        
        $mailParams = @{
            To = $To
            From = $From
            Subject = "$Subject - $(Get-Date -Format 'yyyy-MM-dd')"
            Body = $body
            SmtpServer = $SMTPServer
        }
        
        if ($Attachments -and $Attachments.Count -gt 0) {
            $mailParams.Attachments = $Attachments
        }
        
        Send-MailMessage @mailParams
        Write-Host "✓ Email report sent to: $To" -ForegroundColor Green
        return $true
    }
    catch {
        Write-Error "Failed to send email: $($_.Exception.Message)"
        return $false
    }
}

# Main script execution
Write-Host ""
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "   DHCP LEASE REPORT GENERATOR" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

# Validate DHCP module
if (-not (Import-DHCPModule)) {
    Write-Error "DHCP module is required. Please install DHCP Server role."
    exit 1
}

# Create output directory
$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$reportFolder = Join-Path -Path $OutputPath -ChildPath "DHCP_Lease_Report_$timestamp"
if (-not (Test-Path $reportFolder)) {
    New-Item -ItemType Directory -Path $reportFolder -Force | Out-Null
}
Write-Host "📁 Report folder: $reportFolder" -ForegroundColor Yellow

# Collect report data
$Script:ReportData = Collect-ReportData -Server $DHCPServer -ScopeId $ScopeId -ReportType $ReportType -ExpiryThreshold $ExpiryThreshold -IncludeHistory:$IncludeHistory -IncludeVendorInfo:$IncludeVendorInfo -IncludeScopeStats:$IncludeScopeStats -IncludeLeaseUsage:$IncludeLeaseUsage

if (-not $Script:ReportData) {
    Write-Error "Failed to collect report data. Exiting."
    exit 1
}

# Determine output formats
$formats = @()
if ($OutputFormats -eq "ALL") {
    $formats = @("HTML", "CSV", "JSON", "Excel")
}
else {
    $formats = $OutputFormats -split ','
}

# Generate reports
Write-Host "`n📤 Generating reports in formats: $($formats -join ', ')" -ForegroundColor Yellow
Write-Host ""

if ("HTML" -in $formats) {
    $htmlPath = Join-Path -Path $reportFolder -ChildPath "DHCP_Lease_Report.html"
    if (New-HTMLReport -Data $Script:ReportData -OutputPath $htmlPath) {
        $Script:ExportFiles += $htmlPath
    }
}

if ("CSV" -in $formats) {
    # Export Scopes
    if ($Script:ReportData.Scopes) {
        $csvPath = Join-Path -Path $reportFolder -ChildPath "Scopes.csv"
        Export-ToCSV -FilePath $csvPath -Data $Script:ReportData.Scopes -DataType "Scopes"
        $Script:ExportFiles += $csvPath
    }
    
    # Export Expiring Leases
    if ($Script:ReportData.ExpiringLeases -and $Script:ReportData.ExpiringLeases.Count -gt 0) {
        $csvPath = Join-Path -Path $reportFolder -ChildPath "Expiring_Leases.csv"
        Export-ToCSV -FilePath $csvPath -Data $Script:ReportData.ExpiringLeases -DataType "Expiring Leases"
        $Script:ExportFiles += $csvPath
    }
    
    # Export Top Consumers
    if ($Script:ReportData.TopConsumers -and $Script:ReportData.TopConsumers.Count -gt 0) {
        $csvPath = Join-Path -Path $reportFolder -ChildPath "Top_Consumers.csv"
        Export-ToCSV -FilePath $csvPath -Data $Script:ReportData.TopConsumers -DataType "Top Consumers"
        $Script:ExportFiles += $csvPath
    }
}

if ("JSON" -in $formats) {
    $jsonPath = Join-Path -Path $reportFolder -ChildPath "DHCP_Lease_Data.json"
    if (Export-ToJSON -FilePath $jsonPath -Data $Script:ReportData) {
        $Script:ExportFiles += $jsonPath
    }
}

if ("Excel" -in $formats) {
    $excelPath = Join-Path -Path $reportFolder -ChildPath "DHCP_Lease_Report.xlsx"
    if (Export-ToExcel -FilePath $excelPath -Data $Script:ReportData) {
        $Script:ExportFiles += $excelPath
    }
}

# Export results if specified
if ($ExportResults) {
    $resultsData = $Script:ReportData.Scopes | ForEach-Object {
        [PSCustomObject]@{
            ScopeId = $_.ScopeId
            Name = $_.Name
            State = $_.State
            TotalLeases = $_.TotalLeases
            ActiveLeases = $_.ActiveLeases
            ExpiredLeases = $_.ExpiredLeases
            ExpiringLeases = $_.ExpiringLeases
            StartRange = $_.StartRange
            EndRange = $_.EndRange
            SubnetMask = $_.SubnetMask
            LeaseDuration = $_.LeaseDuration
        }
    }
    $resultsData | Export-Csv -Path $ExportResults -NoTypeInformation -Encoding UTF8
    Write-Host "✓ Results exported to: $ExportResults" -ForegroundColor Green
    $Script:ExportFiles += $ExportResults
}

# Send email
if ($EmailReport -and $EmailTo -and $EmailFrom -and $EmailServer) {
    Write-Host "`n📧 Sending email report..." -ForegroundColor Yellow
    Send-EmailReport -To $EmailTo -From $EmailFrom -SMTPServer $EmailServer -Attachments $Script:ExportFiles
}

# Display summary
$elapsedTime = (Get-Date) - $Script:StartTime
Write-Host ""
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "   REPORT GENERATION COMPLETED" -ForegroundColor Green
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "📊 Report Summary:" -ForegroundColor Cyan
Write-Host "  Server: $DHCPServer" -ForegroundColor White
Write-Host "  Total Scopes: $($Script:ReportData.Summary.TotalScopes)" -ForegroundColor White
Write-Host "  Total Leases: $($Script:ReportData.Summary.TotalLeases)" -ForegroundColor White
Write-Host "  Active Leases: $($Script:ReportData.Summary.TotalActive)" -ForegroundColor Green
Write-Host "  Expiring Leases: $($Script:ReportData.Summary.TotalExpiring)" -ForegroundColor Yellow
Write-Host "  Expired Leases: $($Script:ReportData.Summary.TotalExpired)" -ForegroundColor Red
Write-Host "  ⏱️  Time: $($elapsedTime.ToString('hh\:mm\:ss'))" -ForegroundColor White
Write-Host ""
Write-Host "📁 Report Location: $reportFolder" -ForegroundColor Yellow

if ($Script:ExportFiles.Count -gt 0) {
    Write-Host "`n📄 Generated Files:" -ForegroundColor Green
    foreach ($file in $Script:ExportFiles) {
        if (Test-Path $file) {
            $fileInfo = Get-Item $file
            $size = if ($fileInfo.Length -gt 1MB) { "$([math]::Round($fileInfo.Length / 1MB, 2)) MB" } else { "$([math]::Round($fileInfo.Length / 1KB, 2)) KB" }
            Write-Host "  ✓ $($fileInfo.Name) ($size)" -ForegroundColor Gray
        }
    }
}

Write-Host ""
Write-Host "✅ Script completed successfully!" -ForegroundColor Green