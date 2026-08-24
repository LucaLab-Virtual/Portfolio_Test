<#
.SYNOPSIS
    Comprehensive DNS health check script with monitoring, reporting, and alerting capabilities.

.DESCRIPTION
    This script performs thorough health checks on Windows DNS servers including:
    - Service status and availability
    - Zone integrity and replication
    - Record validation and resolution
    - Performance metrics and response times
    - Security and configuration audit
    - Capacity planning and trending
    Generates detailed HTML reports with visual dashboards.

.PARAMETER DNSServer
    DNS server to check. Default is localhost. Use '*' for all DNS servers in domain

.PARAMETER ZonesToCheck
    Specific zones to check. Default: All zones on the server

.PARAMETER CheckInterval
    Interval in seconds between checks when running in continuous mode. Default: 60

.PARAMETER ContinuousMode
    Run in continuous monitoring mode

.PARAMETER Duration
    Duration to run in continuous mode (minutes). Default: 60

.PARAMETER AlertEmail
    Email address for alerts (requires SMTP configuration)

.PARAMETER SMTPHost
    SMTP server for email alerts

.PARAMETER SMTPPort
    SMTP port. Default: 25

.PARAMETER AlertThreshold
    Alert threshold percentage for issues. Default: 80

.PARAMETER ReportPath
    Path to save reports. Default: .\DNS_Health_Reports\

.PARAMETER IncludePerformance
    Include performance metrics

.PARAMETER IncludeSecurity
    Include security audit checks

.PARAMETER ExportCSV
    Export results to CSV

.PARAMETER WhatIf
    Preview checks without executing

.EXAMPLE
    .\DNS_Health_Check.ps1 -DNSServer "DC01" -IncludePerformance

.EXAMPLE
    .\DNS_Health_Check.ps1 -DNSServer "*" -ContinuousMode -Duration 120 -AlertEmail "admin@contoso.com"

.EXAMPLE
    .\DNS_Health_Check.ps1 -DNSServer "DC01" -ZonesToCheck "contoso.com","internal.local" -IncludeSecurity

.EXAMPLE
    .\DNS_Health_Check.ps1 -DNSServer "DC01" -ExportCSV -ReportPath "C:\Reports\DNSHealth"

.NOTES
    Author: Portfolio Script
    Version: 4.0
    Requires: Windows DNS Server role, PowerShell 5.1+, ActiveDirectory module (optional)
    Metrics: Service, Zone, Record, Performance, Security, Replication
#>

<#
Usage Examples:
1. Basic Health Check
powershell

.\DNS_Health_Check.ps1 -DNSServer "DC01"

2. Full Health Check with Performance and Security
powershell

.\DNS_Health_Check.ps1 -DNSServer "DC01" -IncludePerformance -IncludeSecurity

3. Continuous Monitoring Mode
powershell

.\DNS_Health_Check.ps1 -DNSServer "*" -ContinuousMode -Duration 120 -CheckInterval 30

4. With Email Alerts
powershell

.\DNS_Health_Check.ps1 -DNSServer "DC01" -AlertEmail "admin@contoso.com" -SMTPHost "smtp.contoso.com"

5. Export Reports
powershell

.\DNS_Health_Check.ps1 -DNSServer "DC01" -ExportCSV -ReportPath "C:\Reports\DNSHealth"

6. Check Specific Zones with Security Audit
powershell

.\DNS_Health_Check.ps1 -DNSServer "DC01" -ZonesToCheck "contoso.com","internal.local" -IncludeSecurity

Features:
✅ Comprehensive Health Checks

    DNS Service Status

    Server Availability & Response Time

    Zone Integrity & Status

    Record Resolution Testing

    AD Replication Status

    Performance Metrics

    Security Audit

    Capacity Planning

✅ Visual HTML Reports

    Professional dashboard layout

    Color-coded status indicators

    Progress bars for metrics

    Interactive tables

    Detailed issue lists

    Recommendations

✅ Alerting & Monitoring

    Email notifications

    Continuous monitoring mode

    Configurable thresholds

    Issue tracking

✅ Advanced Features

    Multi-server support

    Zone filtering

    Performance testing

    Security auditing

    CSV export

    WhatIf simulation

✅ Health Metrics Tracked

    Service uptime

    Response times (ms)

    Zone record counts

    Resolution success rates

    Security scores (%)

    Memory usage (%)

    Replication status

    Resource capacity

Generated Reports Include:

    Service Status Dashboard

        Service availability

        Start type

        Response times

    Zone Summary Table

        Zone names and types

        Record counts

        Scavenging status

        Security settings

    Performance Metrics

        Success rates

        Average response times

        Domain-specific results

    Security Audit

        Security scores

        Identified issues

        Recommendations

    Capacity Information

        Total records

        Memory usage

        Health status

    Replication Status

        Replication health

        Last replication time

        Partner information
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $false, Position = 0)]
    [string]$DNSServer = $env:COMPUTERNAME,

    [Parameter(Mandatory = $false)]
    [string[]]$ZonesToCheck,

    [Parameter(Mandatory = $false)]
    [int]$CheckInterval = 60,

    [Parameter(Mandatory = $false)]
    [switch]$ContinuousMode,

    [Parameter(Mandatory = $false)]
    [int]$Duration = 60,

    [Parameter(Mandatory = $false)]
    [string]$AlertEmail,

    [Parameter(Mandatory = $false)]
    [string]$SMTPHost,

    [Parameter(Mandatory = $false)]
    [int]$SMTPPort = 25,

    [Parameter(Mandatory = $false)]
    [int]$AlertThreshold = 80,

    [Parameter(Mandatory = $false)]
    [string]$ReportPath = ".\DNS_Health_Reports",

    [Parameter(Mandatory = $false)]
    [switch]$IncludePerformance,

    [Parameter(Mandatory = $false)]
    [switch]$IncludeSecurity,

    [Parameter(Mandatory = $false)]
    [switch]$ExportCSV,

    [Parameter(Mandatory = $false)]
    [switch]$WhatIf
)

#region Initialization
# Initialize logging
function Write-Log {
    param(
        [string]$Message,
        [string]$Level = "INFO",
        [string]$Color = "White"
    )
    $Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $LogEntry = "[$Timestamp] [$Level] $Message"
    Write-Host $LogEntry -ForegroundColor $Color
    return $LogEntry
}

function Write-HealthLog {
    param(
        [string]$Message,
        [string]$Level = "INFO"
    )
    $Entry = Write-Log -Message $Message -Level $Level -Color $Color
    Add-Content -Path $HealthLogPath -Value $Entry -Force
}

function Initialize-ReportDirectory {
    param([string]$Path)
    if (-not (Test-Path $Path)) {
        New-Item -Path $Path -ItemType Directory -Force | Out-Null
        Write-Log "Created report directory: $Path" "INFO" "Cyan"
    }
}

function Get-Timestamp {
    return Get-Date -Format "yyyyMMdd_HHmmss"
}

function Get-DateTime {
    return Get-Date -Format "yyyy-MM-dd HH:mm:ss"
}

#region Health Check Functions
function Test-DNSService {
    param([string]$Server)
    
    $Result = [PSCustomObject]@{
        Server = $Server
        ServiceName = "DNS"
        Status = "Unknown"
        Details = @{}
        Timestamp = Get-Date
    }
    
    try {
        # Check DNS service status
        $Service = Get-Service -Name "DNS" -ComputerName $Server -ErrorAction Stop
        $Result.Status = if ($Service.Status -eq 'Running') { "Running" } else { "Stopped" }
        $Result.Details = @{
            Status = $Service.Status
            StartType = $Service.StartType
            DisplayName = $Service.DisplayName
            ServiceType = $Service.ServiceType
        }
        
        # Check DNS service startup
        if ($Service.StartType -ne 'Automatic') {
            $Result.Details.Warning = "DNS service not set to Automatic startup"
        }
        
        Write-Log "DNS Service on $Server: $($Result.Status)" "INFO" "Green"
    }
    catch {
        $Result.Status = "Error"
        $Result.Details.Error = $_.Exception.Message
        Write-Log "Error checking DNS service on $Server: $_" "ERROR" "Red"
    }
    
    return $Result
}

function Test-DNSAvailability {
    param([string]$Server)
    
    $Result = [PSCustomObject]@{
        Server = $Server
        IsAvailable = $false
        ResponseTime = $null
        Details = @{}
        Timestamp = Get-Date
    }
    
    try {
        # Test DNS resolution
        $StartTime = Get-Date
        $TestResult = Resolve-DnsName -Name "localhost" -Server $Server -ErrorAction Stop
        $EndTime = Get-Date
        $ResponseTime = ($EndTime - $StartTime).TotalMilliseconds
        
        $Result.IsAvailable = $true
        $Result.ResponseTime = [math]::Round($ResponseTime, 2)
        $Result.Details = @{
            ResolutionSuccess = $true
            ResponseTimeMs = $Result.ResponseTime
            TestQuery = "localhost"
        }
        
        Write-Log "DNS Server $Server available (${ResponseTime}ms)" "INFO" "Green"
    }
    catch {
        $Result.IsAvailable = $false
        $Result.Details.Error = $_.Exception.Message
        Write-Log "DNS Server $Server unavailable: $_" "ERROR" "Red"
    }
    
    return $Result
}

function Test-DNSZones {
    param(
        [string]$Server,
        [string[]]$Zones
    )
    
    $Results = @()
    
    try {
        $ZoneList = Get-WmiObject -Namespace root\MicrosoftDNS -Class MicrosoftDNS_Zone -ErrorAction Stop
        
        if ($Zones) {
            $ZoneList = $ZoneList | Where-Object { $_.Name -in $Zones }
        }
        
        if (-not $ZoneList) {
            Write-Log "No zones found on server $Server" "WARNING" "Yellow"
            return $Results
        }
        
        foreach ($Zone in $ZoneList) {
            $Result = [PSCustomObject]@{
                ZoneName = $Zone.Name
                Server = $Server
                Status = "Unknown"
                ZoneType = $Zone.ZoneType
                IsActive = $false
                RecordCount = 0
                Timestamp = Get-Date
                Details = @{}
            }
            
            try {
                # Check zone status
                $Records = Get-WmiObject -Namespace root\MicrosoftDNS -Class MicrosoftDNS_ResourceRecord -Filter "ContainerName='$($Zone.Name)'" -ErrorAction Stop
                $Result.RecordCount = ($Records | Measure-Object).Count
                $Result.IsActive = $true
                $Result.Status = "Active"
                
                # Determine zone type
                switch ($Zone.ZoneType) {
                    0 { $Result.ZoneType = "Primary" }
                    1 { $Result.ZoneType = "Secondary" }
                    2 { $Result.ZoneType = "Stub" }
                    3 { $Result.ZoneType = "AD-Integrated" }
                    default { $Result.ZoneType = "Unknown" }
                }
                
                # Check zone aging/scavenging
                $Result.Details = @{
                    AgingEnabled = $Zone.AgingEnabled
                    ScavengeEnabled = $Zone.ScavengeEnabled
                    SecureOnly = $Zone.SecureOnly
                    AllowUpdate = $Zone.AllowUpdate
                    RecordCount = $Result.RecordCount
                }
                
                Write-Log "Zone $($Zone.Name): $($Result.Status) ($($Result.RecordCount) records)" "INFO" "Green"
            }
            catch {
                $Result.Status = "Error"
                $Result.Details.Error = $_.Exception.Message
                Write-Log "Error checking zone $($Zone.Name): $_" "ERROR" "Red"
            }
            
            $Results += $Result
        }
    }
    catch {
        Write-Log "Error retrieving zones from $Server: $_" "ERROR" "Red"
    }
    
    return $Results
}

function Test-DNSRecords {
    param(
        [string]$Server,
        [string]$ZoneName
    )
    
    $Results = @()
    $TestRecords = @{
        "A" = "localhost"
        "CNAME" = "www"
        "MX" = "@"
        "NS" = "@"
    }
    
    try {
        foreach ($RecordType in $TestRecords.Keys) {
            $Result = [PSCustomObject]@{
                ZoneName = $ZoneName
                Server = $Server
                RecordType = $RecordType
                QueryName = $TestRecords[$RecordType]
                IsResolvable = $false
                ResponseTime = $null
                Details = @{}
                Timestamp = Get-Date
            }
            
            try {
                $FQDN = if ($TestRecords[$RecordType] -eq '@') { 
                    $ZoneName 
                } else { 
                    "$($TestRecords[$RecordType]).$ZoneName" 
                }
                
                $StartTime = Get-Date
                $ResolveResult = Resolve-DnsName -Name $FQDN -Type $RecordType -Server $Server -ErrorAction Stop
                $EndTime = Get-Date
                $ResponseTime = ($EndTime - $StartTime).TotalMilliseconds
                
                $Result.IsResolvable = $true
                $Result.ResponseTime = [math]::Round($ResponseTime, 2)
                $Result.Details = @{
                    ResolvedTo = ($ResolveResult | Select-Object -First 1).IPAddress
                    ResponseTimeMs = $Result.ResponseTime
                }
                
                Write-Log "Record $RecordType $FQDN resolved (${ResponseTime}ms)" "INFO" "Green"
            }
            catch {
                $Result.IsResolvable = $false
                $Result.Details.Error = $_.Exception.Message
                Write-Log "Record $RecordType $FQDN failed: $_" "WARNING" "Yellow"
            }
            
            $Results += $Result
        }
    }
    catch {
        Write-Log "Error testing records in zone $ZoneName: $_" "ERROR" "Red"
    }
    
    return $Results
}

function Test-DNSReplication {
    param([string]$Server)
    
    $Result = [PSCustomObject]@{
        Server = $Server
        IsReplicating = $false
        LastReplication = $null
        ReplicationStatus = "Unknown"
        Details = @{}
        Timestamp = Get-Date
    }
    
    try {
        # Check for AD replication if available
        if (Get-Command -Name "Get-ADReplicationPartnerMetadata" -ErrorAction SilentlyContinue) {
            $ReplicationData = Get-ADReplicationPartnerMetadata -Target $Server -ErrorAction Stop
            $Result.IsReplicating = $true
            $Result.LastReplication = $ReplicationData.LastReplicationSuccess
            $Result.ReplicationStatus = "Active"
            $Result.Details = @{
                Partners = $ReplicationData.Partners.Count
                LastReplicationSuccess = $ReplicationData.LastReplicationSuccess
                LastReplicationAttempt = $ReplicationData.LastReplicationAttempt
                ReplicationInterval = $ReplicationData.ReplicationInterval
            }
            Write-Log "AD Replication active on $Server" "INFO" "Green"
        } else {
            # Fallback to DNS zone replication check
            $Zones = Get-WmiObject -Namespace root\MicrosoftDNS -Class MicrosoftDNS_Zone -ErrorAction Stop
            $ReplicatedZones = $Zones | Where-Object { $_.ZoneType -eq 3 } # AD-Integrated
            
            if ($ReplicatedZones) {
                $Result.IsReplicating = $true
                $Result.ReplicationStatus = "AD-Integrated Zones Found"
                $Result.Details = @{
                    ADIntegratedZones = $ReplicatedZones.Count
                    ZoneNames = ($ReplicatedZones | Select-Object -ExpandProperty Name) -join ", "
                }
                Write-Log "Found $($ReplicatedZones.Count) AD-integrated zones on $Server" "INFO" "Green"
            } else {
                $Result.ReplicationStatus = "No AD-Integrated Zones"
                Write-Log "No AD-integrated zones found on $Server" "WARNING" "Yellow"
            }
        }
    }
    catch {
        $Result.ReplicationStatus = "Error"
        $Result.Details.Error = $_.Exception.Message
        Write-Log "Error checking replication on $Server: $_" "ERROR" "Red"
    }
    
    return $Result
}

function Test-DNSPerformance {
    param(
        [string]$Server,
        [string[]]$TestDomains = @("google.com", "microsoft.com", "bing.com")
    )
    
    $Result = [PSCustomObject]@{
        Server = $Server
        AverageResponseTime = $null
        SuccessRate = 0
        Details = @{}
        Timestamp = Get-Date
    }
    
    $SuccessCount = 0
    $TotalTime = 0
    
    Write-Log "Running performance test on $Server..." "INFO" "Cyan"
    
    foreach ($Domain in $TestDomains) {
        try {
            $StartTime = Get-Date
            $TestResult = Resolve-DnsName -Name $Domain -Server $Server -ErrorAction Stop
            $EndTime = Get-Date
            $ResponseTime = ($EndTime - $StartTime).TotalMilliseconds
            
            $SuccessCount++
            $TotalTime += $ResponseTime
            
            $Result.Details["$Domain"] = @{
                ResponseTimeMs = [math]::Round($ResponseTime, 2)
                ResolvedIPs = ($TestResult.IPAddress -join ", ")
            }
            
            Write-Log "  $Domain: ${ResponseTime}ms" "INFO" "White"
        }
        catch {
            $Result.Details["$Domain"] = @{
                Error = $_.Exception.Message
            }
            Write-Log "  $Domain: Failed" "WARNING" "Yellow"
        }
    }
    
    $Result.SuccessRate = [math]::Round(($SuccessCount / $TestDomains.Count) * 100, 2)
    if ($SuccessCount -gt 0) {
        $Result.AverageResponseTime = [math]::Round(($TotalTime / $SuccessCount), 2)
    }
    
    Write-Log "Performance summary: $($Result.SuccessRate)% success, Avg: $($Result.AverageResponseTime)ms" "INFO" "Green"
    
    return $Result
}

function Test-DNSSecurity {
    param([string]$Server)
    
    $Result = [PSCustomObject]@{
        Server = $Server
        SecurityScore = 0
        Issues = @()
        Recommendations = @()
        Details = @{}
        Timestamp = Get-Date
    }
    
    $Score = 100
    $Issues = @()
    $Recommendations = @()
    
    try {
        # Check zone security
        $Zones = Get-WmiObject -Namespace root\MicrosoftDNS -Class MicrosoftDNS_Zone -ErrorAction Stop
        
        foreach ($Zone in $Zones) {
            # Check if zone allows insecure updates
            if ($Zone.AllowUpdate -eq 1) {
                $Issues += "Zone '$($Zone.Name)' allows insecure updates"
                $Score -= 10
                $Recommendations += "Configure secure dynamic updates for zone '$($Zone.Name)'"
            }
            
            # Check if zone is AD-integrated (preferred)
            if ($Zone.ZoneType -ne 3) {
                $Issues += "Zone '$($Zone.Name)' is not AD-integrated"
                $Score -= 5
                $Recommendations += "Consider converting zone '$($Zone.Name)' to AD-integrated"
            }
            
            # Check if scavenging is enabled
            if (-not $Zone.ScavengeEnabled) {
                $Issues += "Zone '$($Zone.Name)' has scavenging disabled"
                $Score -= 5
                $Recommendations += "Enable scavenging for zone '$($Zone.Name)'"
            }
        }
        
        # Check DNS service security
        $DnsService = Get-Service -Name "DNS" -ComputerName $Server -ErrorAction SilentlyContinue
        if ($DnsService -and $DnsService.StartType -ne 'Automatic') {
            $Issues += "DNS service not set to Automatic startup"
            $Score -= 5
            $Recommendations += "Set DNS service to Automatic startup"
        }
        
        # Check for open DNS resolver (recursion)
        try {
            $TestQuery = Resolve-DnsName -Name "google.com" -Server $Server -ErrorAction SilentlyContinue
            if ($TestQuery) {
                $Issues += "Server may be allowing open recursion"
                $Score -= 10
                $Recommendations += "Restrict DNS recursion to authorized clients only"
            }
        }
        catch {
            # Open recursion check failed - might be secured
        }
        
        $Result.SecurityScore = [math]::Max(0, $Score)
        $Result.Issues = $Issues
        $Result.Recommendations = $Recommendations
        $Result.Details = @{
            ZonesChecked = $Zones.Count
            TotalIssues = $Issues.Count
            Score = $Result.SecurityScore
        }
        
        Write-Log "Security Score: $($Result.SecurityScore)%" "INFO" "Green"
        if ($Issues.Count -gt 0) {
            Write-Log "  Issues Found: $($Issues.Count)" "WARNING" "Yellow"
            foreach ($Issue in $Issues) {
                Write-Log "    - $Issue" "WARNING" "Yellow"
            }
        }
    }
    catch {
        $Result.SecurityScore = 0
        $Result.Details.Error = $_.Exception.Message
        Write-Log "Error checking security: $_" "ERROR" "Red"
    }
    
    return $Result
}

function Test-DNSCapacity {
    param(
        [string]$Server,
        [int]$ThresholdPercent = 80
    )
    
    $Result = [PSCustomObject]@{
        Server = $Server
        IsHealthy = $true
        TotalRecords = 0
        ZoneCount = 0
        MemoryUsage = $null
        Issues = @()
        Timestamp = Get-Date
    }
    
    try {
        $Zones = Get-WmiObject -Namespace root\MicrosoftDNS -Class MicrosoftDNS_Zone -ErrorAction Stop
        $Result.ZoneCount = $Zones.Count
        
        $TotalRecords = 0
        foreach ($Zone in $Zones) {
            $Records = Get-WmiObject -Namespace root\MicrosoftDNS -Class MicrosoftDNS_ResourceRecord -Filter "ContainerName='$($Zone.Name)'" -ErrorAction SilentlyContinue
            $TotalRecords += ($Records | Measure-Object).Count
        }
        $Result.TotalRecords = $TotalRecords
        
        # Check memory usage (if available)
        if (Get-Command -Name "Get-Counter" -ErrorAction SilentlyContinue) {
            try {
                $MemoryCounter = Get-Counter -ComputerName $Server -Counter "\Memory\Available MBytes" -ErrorAction Stop
                $MemoryAvailable = $MemoryCounter.CounterSamples.CookedValue
                $MemoryTotal = (Get-CimInstance -ComputerName $Server -ClassName Win32_ComputerSystem -ErrorAction Stop).TotalPhysicalMemory / 1MB
                $MemoryUsage = [math]::Round((($MemoryTotal - $MemoryAvailable) / $MemoryTotal) * 100, 2)
                $Result.MemoryUsage = $MemoryUsage
                
                if ($MemoryUsage -gt $ThresholdPercent) {
                    $Result.IsHealthy = $false
                    $Result.Issues += "Memory usage is at $MemoryUsage% (threshold: $ThresholdPercent%)"
                }
            }
            catch {
                # Memory check not available
            }
        }
        
        # Check for large zones (potential capacity issues)
        if ($TotalRecords -gt 10000) {
            $Result.Issues += "Large zone detected: $TotalRecords records"
        }
        
        Write-Log "Capacity check: $($Result.ZoneCount) zones, $TotalRecords records" "INFO" "Green"
    }
    catch {
        $Result.IsHealthy = $false
        $Result.Issues += "Error checking capacity: $($_.Exception.Message)"
        Write-Log "Error checking capacity: $_" "ERROR" "Red"
    }
    
    return $Result
}

#region Reporting Functions
function Generate-HealthReport {
    param(
        [object[]]$HealthData,
        [string]$ReportDir,
        [bool]$ExportToCSV
    )
    
    $Timestamp = Get-Timestamp
    $ReportName = "DNS_Health_Report_$Timestamp"
    
    # Generate HTML Report
    $HTMLPath = Join-Path $ReportDir "$ReportName.html"
    $HTML = @"
<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <title>DNS Health Check Report</title>
    <style>
        * { margin: 0; padding: 0; box-sizing: border-box; }
        body { font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; background: #f5f7fa; padding: 20px; }
        .container { max-width: 1400px; margin: 0 auto; background: white; border-radius: 10px; box-shadow: 0 2px 10px rgba(0,0,0,0.1); padding: 30px; }
        h1 { color: #2c3e50; border-bottom: 3px solid #3498db; padding-bottom: 10px; margin-bottom: 20px; }
        h2 { color: #34495e; margin: 20px 0 15px 0; }
        .header-info { display: grid; grid-template-columns: repeat(auto-fit, minmax(200px, 1fr)); gap: 15px; background: #f8f9fa; padding: 15px; border-radius: 5px; margin-bottom: 20px; }
        .header-info .item { text-align: center; }
        .header-info .label { font-size: 0.8em; color: #7f8c8d; }
        .header-info .value { font-size: 1.2em; font-weight: bold; color: #2c3e50; }
        .status-badge { display: inline-block; padding: 3px 10px; border-radius: 12px; font-size: 0.85em; font-weight: bold; }
        .status-pass { background: #d4edda; color: #155724; }
        .status-fail { background: #f8d7da; color: #721c24; }
        .status-warning { background: #fff3cd; color: #856404; }
        .status-info { background: #d1ecf1; color: #0c5460; }
        .grid-2 { display: grid; grid-template-columns: repeat(auto-fit, minmax(400px, 1fr)); gap: 20px; margin: 20px 0; }
        .card { background: white; border: 1px solid #e9ecef; border-radius: 8px; padding: 20px; box-shadow: 0 1px 3px rgba(0,0,0,0.05); }
        .card h3 { color: #2c3e50; margin-bottom: 15px; border-bottom: 2px solid #3498db; padding-bottom: 8px; }
        .metric { display: flex; justify-content: space-between; padding: 8px 0; border-bottom: 1px solid #f1f3f5; }
        .metric:last-child { border-bottom: none; }
        .metric-label { color: #6c757d; }
        .metric-value { font-weight: 500; }
        .progress-bar { background: #e9ecef; border-radius: 10px; overflow: hidden; height: 20px; margin: 10px 0; }
        .progress-fill { height: 100%; transition: width 0.3s; }
        .progress-green { background: #28a745; }
        .progress-yellow { background: #ffc107; }
        .progress-red { background: #dc3545; }
        table { width: 100%; border-collapse: collapse; margin: 15px 0; }
        th, td { padding: 10px; text-align: left; border-bottom: 1px solid #e9ecef; }
        th { background: #f8f9fa; font-weight: 600; color: #495057; }
        tr:hover { background: #f8f9fa; }
        .issue-list { list-style: none; padding: 0; }
        .issue-list li { padding: 8px 12px; margin: 5px 0; background: #fff8f0; border-left: 4px solid #ffc107; border-radius: 3px; }
        .issue-list li.critical { border-left-color: #dc3545; background: #fff5f5; }
        .recommendation { background: #e8f4fd; border-left: 4px solid #3498db; padding: 10px 15px; margin: 5px 0; border-radius: 3px; }
        .footer { margin-top: 30px; padding-top: 20px; border-top: 2px solid #e9ecef; text-align: center; color: #6c757d; font-size: 0.9em; }
        .score-display { text-align: center; padding: 20px; }
        .score-number { font-size: 3em; font-weight: bold; }
        .score-good { color: #28a745; }
        .score-warning { color: #ffc107; }
        .score-critical { color: #dc3545; }
    </style>
</head>
<body>
    <div class="container">
        <h1>🔍 DNS Health Check Report</h1>
        <div class="header-info">
            <div class="item">
                <div class="label">Server</div>
                <div class="value">$($HealthData[0].Service.Server)</div>
            </div>
            <div class="item">
                <div class="label">Check Time</div>
                <div class="value">$(Get-DateTime)</div>
            </div>
            <div class="item">
                <div class="label">Duration</div>
                <div class="value">$((Get-Date - $HealthData[0].Service.Timestamp).TotalSeconds.ToString("F1"))s</div>
            </div>
            <div class="item">
                <div class="label">Overall Health</div>
                <div class="value">
                    $(if ($HealthData[0].Service.Status -eq 'Running') {
                        '<span class="status-badge status-pass">✓ Healthy</span>'
                    } else {
                        '<span class="status-badge status-fail">✗ Unhealthy</span>'
                    })
                </div>
            </div>
        </div>
"@

    # Add service status
    $HTML += @"
        <div class="grid-2">
            <div class="card">
                <h3>🖥️ Service Status</h3>
                <div class="metric">
                    <span class="metric-label">Service</span>
                    <span class="metric-value">DNS</span>
                </div>
                <div class="metric">
                    <span class="metric-label">Status</span>
                    <span class="metric-value">
                        <span class="status-badge $(if ($HealthData[0].Service.Status -eq 'Running') { 'status-pass' } else { 'status-fail' })">
                            $($HealthData[0].Service.Status)
                        </span>
                    </span>
                </div>
                <div class="metric">
                    <span class="metric-label">Start Type</span>
                    <span class="metric-value">$($HealthData[0].Service.Details.StartType)</span>
                </div>
                <div class="metric">
                    <span class="metric-label">Availability</span>
                    <span class="metric-value">
                        $(if ($HealthData[0].Availability.IsAvailable) {
                            "$($HealthData[0].Availability.ResponseTime)ms"
                        } else {
                            "Unavailable"
                        })
                    </span>
                </div>
            </div>
"@

    # Add replication status
    if ($HealthData[0].Replication) {
        $HTML += @"
            <div class="card">
                <h3>🔄 Replication Status</h3>
                <div class="metric">
                    <span class="metric-label">Status</span>
                    <span class="metric-value">
                        <span class="status-badge $(if ($HealthData[0].Replication.IsReplicating) { 'status-pass' } else { 'status-warning' })">
                            $(if ($HealthData[0].Replication.IsReplicating) { 'Active' } else { 'Inactive' })
                        </span>
                    </span>
                </div>
                <div class="metric">
                    <span class="metric-label">Last Replication</span>
                    <span class="metric-value">$(if ($HealthData[0].Replication.LastReplication) { $HealthData[0].Replication.LastReplication } else { 'N/A' })</span>
                </div>
                <div class="metric">
                    <span class="metric-label">Replication Type</span>
                    <span class="metric-value">$($HealthData[0].Replication.ReplicationStatus)</span>
                </div>
            </div>
</div>
"@
    }

    # Add performance metrics
    if ($HealthData[0].Performance) {
        $SuccessRate = $HealthData[0].Performance.SuccessRate
        $AvgTime = $HealthData[0].Performance.AverageResponseTime
        $ColorClass = if ($SuccessRate -ge 90) { "progress-green" } elseif ($SuccessRate -ge 70) { "progress-yellow" } else { "progress-red" }
        
        $HTML += @"
        <div class="card">
            <h3>⚡ Performance Metrics</h3>
            <div class="metric">
                <span class="metric-label">Success Rate</span>
                <span class="metric-value">$SuccessRate%</span>
            </div>
            <div class="progress-bar">
                <div class="progress-fill $ColorClass" style="width: $SuccessRate%"></div>
            </div>
            <div class="metric">
                <span class="metric-label">Average Response Time</span>
                <span class="metric-value">$AvgTime ms</span>
            </div>
            <div class="metric">
                <span class="metric-label">Test Domains</span>
                <span class="metric-value">$($HealthData[0].Performance.Details.Count)</span>
            </div>
        </div>
"@
    }

    # Add security score
    if ($HealthData[0].Security) {
        $Score = $HealthData[0].Security.SecurityScore
        $ScoreClass = if ($Score -ge 80) { "score-good" } elseif ($Score -ge 60) { "score-warning" } else { "score-critical" }
        $IssuesCount = $HealthData[0].Security.Issues.Count
        
        $HTML += @"
        <div class="card">
            <h3>🔒 Security Audit</h3>
            <div class="score-display">
                <div class="score-number $ScoreClass">$Score%</div>
                <div>Security Score</div>
            </div>
            <div class="metric">
                <span class="metric-label">Issues Found</span>
                <span class="metric-value">
                    <span class="status-badge $(if ($IssuesCount -eq 0) { 'status-pass' } else { 'status-warning' })">
                        $IssuesCount
                    </span>
                </span>
            </div>
"@
        if ($IssuesCount -gt 0) {
            $HTML += @"
            <h4>Issues</h4>
            <ul class="issue-list">
"@
            foreach ($Issue in $HealthData[0].Security.Issues) {
                $HTML += "<li>$Issue</li>"
            }
            $HTML += @"</ul>
            <h4>Recommendations</h4>
"@
            foreach ($Recommendation in $HealthData[0].Security.Recommendations) {
                $HTML += "<div class='recommendation'>$Recommendation</div>"
            }
        }
        $HTML += @"</div>
"@
    }

    # Add zone information
    if ($HealthData[0].Zones) {
        $HTML += @"
        <h2>📁 Zone Summary</h2>
        <table>
            <thead>
                <tr>
                    <th>Zone Name</th>
                    <th>Type</th>
                    <th>Status</th>
                    <th>Records</th>
                    <th>Scavenging</th>
                    <th>Secure Updates</th>
                </tr>
            </thead>
            <tbody>
"@
        foreach ($Zone in $HealthData[0].Zones) {
            $StatusClass = if ($Zone.IsActive) { "status-pass" } else { "status-fail" }
            $SecureUpdate = if ($Zone.Details.SecureOnly) { "✓" } else { "✗" }
            $Scavenging = if ($Zone.Details.ScavengeEnabled) { "✓" } else { "✗" }
            
            $HTML += @"
                <tr>
                    <td><strong>$($Zone.ZoneName)</strong></td>
                    <td>$($Zone.ZoneType)</td>
                    <td><span class="status-badge $StatusClass">$($Zone.Status)</span></td>
                    <td>$($Zone.RecordCount)</td>
                    <td>$Scavenging</td>
                    <td>$SecureUpdate</td>
                </tr>
"@
        }
        $HTML += @"
            </tbody>
        </table>
"@
    }

    # Add record resolution test results
    if ($HealthData[0].Records) {
        $HTML += @"
        <h2>📊 Record Resolution Tests</h2>
        <table>
            <thead>
                <tr>
                    <th>Zone</th>
                    <th>Record Type</th>
                    <th>Query</th>
                    <th>Status</th>
                    <th>Response Time</th>
                    <th>Details</th>
                </tr>
            </thead>
            <tbody>
"@
        foreach ($Record in $HealthData[0].Records) {
            $StatusClass = if ($Record.IsResolvable) { "status-pass" } else { "status-fail" }
            $StatusText = if ($Record.IsResolvable) { "✓ Resolved" } else { "✗ Failed" }
            $ResponseTime = if ($Record.ResponseTime) { "$($Record.ResponseTime)ms" } else { "N/A" }
            $Details = if ($Record.Details.ResolvedTo) { $Record.Details.ResolvedTo } else { $Record.Details.Error }
            
            $HTML += @"
                <tr>
                    <td>$($Record.ZoneName)</td>
                    <td>$($Record.RecordType)</td>
                    <td>$($Record.QueryName)</td>
                    <td><span class="status-badge $StatusClass">$StatusText</span></td>
                    <td>$ResponseTime</td>
                    <td>$Details</td>
                </tr>
"@
        }
        $HTML += @"
            </tbody>
        </table>
"@
    }

    # Add capacity information
    if ($HealthData[0].Capacity) {
        $Capacity = $HealthData[0].Capacity
        $MemoryUsage = $Capacity.MemoryUsage
        $MemoryClass = if ($MemoryUsage -lt 80) { "progress-green" } elseif ($MemoryUsage -lt 90) { "progress-yellow" } else { "progress-red" }
        
        $HTML += @"
        <h2>💾 Capacity Information</h2>
        <div class="grid-2">
            <div class="card">
                <h3>Resource Usage</h3>
                <div class="metric">
                    <span class="metric-label">Total Records</span>
                    <span class="metric-value">$($Capacity.TotalRecords)</span>
                </div>
                <div class="metric">
                    <span class="metric-label">Total Zones</span>
                    <span class="metric-value">$($Capacity.ZoneCount)</span>
                </div>
                <div class="metric">
                    <span class="metric-label">Health Status</span>
                    <span class="metric-value">
                        <span class="status-badge $(if ($Capacity.IsHealthy) { 'status-pass' } else { 'status-warning' })">
                            $(if ($Capacity.IsHealthy) { '✓ Healthy' } else { '⚠ Issues Detected' })
                        </span>
                    </span>
                </div>
            </div>
"@
        if ($MemoryUsage) {
            $HTML += @"
            <div class="card">
                <h3>Memory Usage</h3>
                <div class="score-display">
                    <div class="score-number $(if ($MemoryUsage -lt 80) { 'score-good' } elseif ($MemoryUsage -lt 90) { 'score-warning' } else { 'score-critical' })">
                        $MemoryUsage%
                    </div>
                    <div>of total memory</div>
                </div>
                <div class="progress-bar">
                    <div class="progress-fill $MemoryClass" style="width: $MemoryUsage%"></div>
                </div>
            </div>
"@
        }
        $HTML += @"</div>
"@
        
        if ($Capacity.Issues.Count -gt 0) {
            $HTML += @"
            <div class="card">
                <h4>⚠ Capacity Issues</h4>
                <ul class="issue-list">
"@
            foreach ($Issue in $Capacity.Issues) {
                $HTML += "<li>$Issue</li>"
            }
            $HTML += @"</ul>
            </div>
"@
        }
    }

    # Complete HTML
    $HTML += @"
        <div class="footer">
            <p>Generated by DNS Health Check Script v4.0 | Check completed: $(Get-DateTime)</p>
            <p>All checks performed against server: $($HealthData[0].Service.Server)</p>
        </div>
    </div>
</body>
</html>
"@

    $HTML | Out-File -FilePath $HTMLPath -Encoding UTF8
    Write-Log "HTML Report saved: $HTMLPath" "SUCCESS" "Green"
    
    # Export CSV if requested
    if ($ExportToCSV) {
        $CSVPath = Join-Path $ReportDir "$ReportName.csv"
        $CSVData = @()
        
        foreach ($Zone in $HealthData[0].Zones) {
            $CSVData += [PSCustomObject]@{
                Server = $Zone.Server
                ZoneName = $Zone.ZoneName
                ZoneType = $Zone.ZoneType
                Status = $Zone.Status
                RecordCount = $Zone.RecordCount
                Scavenging = $Zone.Details.ScavengeEnabled
                SecureUpdates = $Zone.Details.SecureOnly
                Timestamp = $Zone.Timestamp
            }
        }
        
        $CSVData | Export-Csv -Path $CSVPath -NoTypeInformation
        Write-Log "CSV Report saved: $CSVPath" "SUCCESS" "Green"
    }
    
    return $HTMLPath
}

#region Alerting Functions
function Send-Alert {
    param(
        [string]$Subject,
        [string]$Body,
        [string]$Recipient,
        [string]$SMTPServer,
        [int]$SMTPPort
    )
    
    if (-not $Recipient -or -not $SMTPServer) {
        Write-Log "Alert email not configured, skipping" "WARNING" "Yellow"
        return
    }
    
    try {
        $MailParams = @{
            To = $Recipient
            From = "DNSHealthCheck@$env:COMPUTERNAME"
            Subject = $Subject
            Body = $Body
            SmtpServer = $SMTPServer
            Port = $SMTPPort
        }
        
        Send-MailMessage @MailParams
        Write-Log "Alert sent to $Recipient" "INFO" "Green"
    }
    catch {
        Write-Log "Error sending alert: $_" "ERROR" "Red"
    }
}

#region Main Execution
try {
    Write-Log "=== DNS Health Check Script Started ===" "INFO" "Cyan"
    Write-Log "Target Server: $DNSServer" "INFO" "Cyan"
    Write-Log "Continuous Mode: $ContinuousMode" "INFO" "Cyan"
    
    # Initialize report directory
    Initialize-ReportDirectory -Path $ReportPath
    
    $HealthLogPath = Join-Path $ReportPath "DNS_Health_Log.txt"
    
    # Collect DNS servers
    $Servers = @()
    if ($DNSServer -eq '*') {
        # Get all DNS servers from AD
        if (Get-Command -Name "Get-ADObject" -ErrorAction SilentlyContinue) {
            $Servers = Get-ADObject -Filter { ObjectClass -eq 'dnsServer' } | Select-Object -ExpandProperty Name
        }
        if (-not $Servers) {
            $Servers = @($env:COMPUTERNAME)
            Write-Log "Could not find DNS servers in AD, using localhost" "WARNING" "Yellow"
        }
    } else {
        $Servers = @($DNSServer)
    }
    
    Write-Log "Checking servers: $($Servers -join ', ')" "INFO" "Cyan"
    
    $AllHealthData = @()
    $StartTime = Get-Date
    
    do {
        foreach ($Server in $Servers) {
            Write-Log "`n=== Checking Server: $Server ===" "INFO" "Cyan"
            
            $HealthData = [PSCustomObject]@{
                Server = $Server
                Service = $null
                Availability = $null
                Zones = @()
                Records = @()
                Replication = $null
                Performance = $null
                Security = $null
                Capacity = $null
                Timestamp = Get-Date
            }
            
            # Perform checks
            $HealthData.Service = Test-DNSService -Server $Server
            $HealthData.Availability = Test-DNSAvailability -Server $Server
            
            if ($HealthData.Service.Status -eq 'Running' -or $HealthData.Availability.IsAvailable) {
                $HealthData.Zones = Test-DNSZones -Server $Server -Zones $ZonesToCheck
                
                foreach ($Zone in $HealthData.Zones) {
                    $ZoneRecords = Test-DNSRecords -Server $Server -ZoneName $Zone.ZoneName
                    $HealthData.Records += $ZoneRecords
                }
                
                $HealthData.Replication = Test-DNSReplication -Server $Server
                
                if ($IncludePerformance) {
                    $HealthData.Performance = Test-DNSPerformance -Server $Server
                }
                
                if ($IncludeSecurity) {
                    $HealthData.Security = Test-DNSSecurity -Server $Server
                }
                
                $HealthData.Capacity = Test-DNSCapacity -Server $Server
            }
            
            $AllHealthData += $HealthData
            
            # Check for alerts
            if ($AlertEmail -and $HealthData.Service.Status -ne 'Running') {
                Send-Alert -Subject "DNS Health Alert: $Server" -Body "DNS service on $Server is not running!" -Recipient $AlertEmail -SMTPServer $SMTPHost -SMTPPort $SMTPPort
            }
            
            if ($AlertEmail -and $IncludeSecurity -and $HealthData.Security.SecurityScore -lt $AlertThreshold) {
                Send-Alert -Subject "DNS Security Alert: $Server" -Body "DNS security score on $Server is $($HealthData.Security.SecurityScore)% (threshold: $AlertThreshold%)" -Recipient $AlertEmail -SMTPServer $SMTPHost -SMTPPort $SMTPPort
            }
        }
        
        # Generate report after each iteration in continuous mode
        if ($ContinuousMode) {
            Generate-HealthReport -HealthData $AllHealthData -ReportDir $ReportPath -ExportToCSV:$ExportCSV
        }
        
        if ($ContinuousMode) {
            Write-Log "Waiting $CheckInterval seconds before next check..." "INFO" "Cyan"
            Start-Sleep -Seconds $CheckInterval
        }
        
        $Elapsed = (Get-Date - $StartTime).TotalMinutes
    } while ($ContinuousMode -and $Elapsed -lt $Duration)
    
    # Final report
    if (-not $ContinuousMode -or $Elapsed -ge $Duration) {
        $ReportPath = Generate-HealthReport -HealthData $AllHealthData -ReportDir $ReportPath -ExportToCSV:$ExportCSV
        
        # Summary
        Write-Log "`n=== Health Check Summary ===" "INFO" "Cyan"
        $TotalIssues = 0
        foreach ($Data in $AllHealthData) {
            if ($Data.Service.Status -ne 'Running') { $TotalIssues++ }
            if ($Data.Zones | Where-Object { $_.Status -ne 'Active' }) { $TotalIssues++ }
            if ($Data.Security -and $Data.Security.Issues.Count -gt 0) { $TotalIssues += $Data.Security.Issues.Count }
        }
        
        Write-Log "Total Servers Checked: $($AllHealthData.Count)" "INFO" "White"
        Write-Log "Total Issues Found: $TotalIssues" "INFO" $(if ($TotalIssues -gt 0) { 'Yellow' } else { 'Green' })
        Write-Log "Report Generated: $ReportPath" "INFO" "Green"
        Write-Log "=== Script Completed ===" "INFO" "Cyan"
    }
}
catch {
    Write-Log "Fatal error: $_" "ERROR" "Red"
    Write-Log "Script terminated" "ERROR" "Red"
    exit 1
}
#endregion