<#
.SYNOPSIS
    Comprehensive Domain Controller Health Report.
.DESCRIPTION
    This script performs a thorough health check on all Domain Controllers in the domain,
    checking services, disk space, replication, time synchronization, event logs, and more.
    Results are exported to CSV and HTML formats for easy review.
.PARAMETER ExportPath
    Path where the report files will be saved. Defaults to current directory.
.PARAMETER GenerateHTML
    Switch to generate an HTML report in addition to CSV.
.PARAMETER IncludeDetailedEvents
    Switch to include recent critical events in the report (may increase runtime).
.PARAMETER SendEmail
    Switch to send report via email (requires -SmtpServer, -To, -From).
.PARAMETER SmtpServer
    SMTP server for email notifications.
.PARAMETER To
    Recipient email address.
.PARAMETER From
    Sender email address.
.PARAMETER CriticalOnly
    Switch to only report DCs with critical issues.
.EXAMPLE
    .\Domain_Controllers_Health_Report.ps1 -ExportPath "C:\Reports" -GenerateHTML
.EXAMPLE
    .\Domain_Controllers_Health_Report.ps1 -CriticalOnly -SendEmail -SmtpServer "smtp.domain.com" -To "admin@domain.com" -From "reports@domain.com"
.NOTES
    Author: PowerShell Portfolio
    Version: 1.0
    Requires: ActiveDirectory module, Admin privileges
#>

<#
    Network Connectivity: Ping tests with response time

    Critical Services: NTDS, KDC, NetLogon, DNS, W32Time

    Disk Space: System drive free space with warnings/critical alerts

    Time Synchronization: Time offset checking with NTP source

    Replication Status: AD replication health and errors

    Event Logs: Critical AD, DNS, and Kerberos events (optional)

    System Information: OS version, uptime, last boot time

    Comprehensive Output: CSV and optional HTML reports

    Email Notifications: Send reports via email

    Filtering: Critical-only mode for focused reporting

Export Strategy:

    CSV file with all health metrics (saved with timestamp)

    Optional HTML report with color-coded status and summary dashboard

    Console output shows only summary information and top issues

    Both exports avoid cluttering the console with excessive data

The script generates files like:

    DCHealth_Report_20260122_143025.csv

    DCHealth_Report_20260122_143025.html (if requested)
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [string]$ExportPath = ".",
    
    [Parameter(Mandatory = $false)]
    [switch]$GenerateHTML,
    
    [Parameter(Mandatory = $false)]
    [switch]$IncludeDetailedEvents,
    
    [Parameter(Mandatory = $false)]
    [switch]$SendEmail,
    
    [Parameter(Mandatory = $false)]
    [string]$SmtpServer,
    
    [Parameter(Mandatory = $false)]
    [string]$To,
    
    [Parameter(Mandatory = $false)]
    [string]$From,
    
    [Parameter(Mandatory = $false)]
    [switch]$CriticalOnly,
    
    [Parameter(Mandatory = $false)]
    [int]$DaysBack = 7
)

# Function to test if running as administrator
function Test-Administrator {
    $CurrentUser = [Security.Principal.WindowsIdentity]::GetCurrent()
    $Principal = New-Object Security.Principal.WindowsPrincipal($CurrentUser)
    return $Principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

# Check for admin privileges
if (-not (Test-Administrator)) {
    Write-Warning "This script requires administrator privileges for some checks. Some data may be limited."
}

# Ensure the ActiveDirectory module is loaded
try {
    Import-Module ActiveDirectory -ErrorAction Stop
    Write-Host "ActiveDirectory module loaded successfully." -ForegroundColor Green
}
catch {
    Write-Error "Failed to load ActiveDirectory module. Please install RSAT-AD-PowerShell."
    exit 1
}

# Create export directory if it doesn't exist
if (!(Test-Path -Path $ExportPath)) {
    New-Item -ItemType Directory -Path $ExportPath -Force | Out-Null
    Write-Host "Created export directory: $ExportPath" -ForegroundColor Yellow
}

# Get the domain
try {
    $Domain = Get-ADDomain
    $DomainName = $Domain.DNSRoot
    Write-Host "Domain: $DomainName" -ForegroundColor Cyan
}
catch {
    Write-Error "Failed to get domain information: $_"
    exit 1
}

# Get all Domain Controllers
try {
    $DCs = Get-ADDomainController -Filter * | Sort-Object Name
    Write-Host "Found $($DCs.Count) Domain Controllers" -ForegroundColor Green
}
catch {
    Write-Error "Failed to retrieve Domain Controllers: $_"
    exit 1
}

# Array to store results
$Results = @()
$TotalDCs = $DCs.Count
$CurrentDC = 0

Write-Host "`nStarting health check on Domain Controllers..." -ForegroundColor Yellow
Write-Host "============================================`n" -ForegroundColor Gray

# Function to check DC health
function Check-DCHealth {
    param(
        [Microsoft.ActiveDirectory.Management.ADDomainController]$DC,
        [string]$ExportPath,
        [bool]$IncludeDetailedEvents,
        [int]$DaysBack
    )
    
    $DCName = $DC.Name
    $DCFQDN = $DC.HostName
    $Site = $DC.Site
    
    Write-Host "Checking: $DCName (Site: $Site)" -ForegroundColor Cyan
    
    # Initialize health status
    $HealthStatus = @{
        DCName = $DCName
        FQDN = $DCFQDN
        Site = $Site
        IPv4Address = $DC.IPv4Address
        OperatingSystem = $null
        LastLogonReplication = $null
        IsGlobalCatalog = $DC.IsGlobalCatalog
        IsReadOnly = $DC.IsReadOnly
        ComputerGUID = $DC.ComputerGUID
        ServerObjectGUID = $DC.ServerObjectGUID
        
        # Health checks
        PingStatus = $false
        PingResponseTime = $null
        ServicesRunning = $true
        FailedServices = @()
        DiskSpaceOK = $true
        DiskSpaceWarning = $false
        DiskSpaceCritical = $false
        FreeSpaceGB = $null
        TotalSpaceGB = $null
        FreeSpacePercent = $null
        TimeSyncStatus = $false
        TimeOffsetSeconds = $null
        NTPStatus = $null
        ReplicationStatus = $true
        ReplicationErrors = @()
        NTDSStatus = $true
        KDCStatus = $true
        NetLogonStatus = $true
        DNSStatus = $true
        W32TimeStatus = $true
        CriticalEvents = @()
        EventLogStatus = $true
        SystemUptime = $null
        LastBootTime = $null
        OverallHealth = "Healthy"
        HealthIssues = @()
        Recommendations = @()
        Notes = @()
    }
    
    # Test network connectivity (ICMP ping)
    try {
        $Ping = Test-Connection -ComputerName $DCFQDN -Count 2 -ErrorAction Stop -Quiet
        if ($Ping) {
            $PingResponse = Test-Connection -ComputerName $DCFQDN -Count 1 -ErrorAction Stop
            $HealthStatus.PingStatus = $true
            $HealthStatus.PingResponseTime = [math]::Round($PingResponse.ResponseTime, 1)
            Write-Host "  ✓ Ping successful ($($HealthStatus.PingResponseTime)ms)" -ForegroundColor Green
        }
        else {
            Write-Host "  ✗ Ping failed" -ForegroundColor Red
            $HealthStatus.HealthIssues += "Ping failed - DC unreachable"
            $HealthStatus.OverallHealth = "Critical"
        }
    }
    catch {
        Write-Host "  ✗ Ping failed: $_" -ForegroundColor Red
        $HealthStatus.PingStatus = $false
        $HealthStatus.HealthIssues += "Ping failed - DC unreachable"
        $HealthStatus.OverallHealth = "Critical"
    }
    
    # Only proceed with remote checks if DC is reachable
    if ($HealthStatus.PingStatus) {
        # Get operating system information
        try {
            $OSInfo = Get-WmiObject -Class Win32_OperatingSystem -ComputerName $DCFQDN -ErrorAction Stop
            $HealthStatus.OperatingSystem = $OSInfo.Caption
            $HealthStatus.SystemUptime = (Get-Date) - $OSInfo.LastBootUpTime
            $HealthStatus.LastBootTime = $OSInfo.LastBootUpTime
            Write-Host "  ✓ OS: $($HealthStatus.OperatingSystem)" -ForegroundColor Green
        }
        catch {
            $HealthStatus.OperatingSystem = "Unknown"
            $HealthStatus.Notes += "Unable to retrieve OS information"
            Write-Host "  ⚠ Could not retrieve OS information" -ForegroundColor Yellow
        }
        
        # Check critical services
        Write-Host "  Checking critical services..." -ForegroundColor Gray
        $Services = @{
            "NTDS" = "Active Directory Domain Services"
            "KDC" = "Kerberos Key Distribution Center"
            "NetLogon" = "Net Logon"
            "DNS" = "DNS Server"
            "W32Time" = "Windows Time"
        }
        
        foreach ($Service in $Services.Keys) {
            try {
                $ServiceObj = Get-Service -Name $Service -ComputerName $DCFQDN -ErrorAction Stop
                if ($ServiceObj.Status -eq 'Running') {
                    Write-Host "    ✓ $($Services[$Service]) is running" -ForegroundColor Green
                    switch ($Service) {
                        "NTDS" { $HealthStatus.NTDSStatus = $true }
                        "KDC" { $HealthStatus.KDCStatus = $true }
                        "NetLogon" { $HealthStatus.NetLogonStatus = $true }
                        "DNS" { $HealthStatus.DNSStatus = $true }
                        "W32Time" { $HealthStatus.W32TimeStatus = $true }
                    }
                }
                else {
                    Write-Host "    ✗ $($Services[$Service]) is $($ServiceObj.Status)" -ForegroundColor Red
                    $HealthStatus.ServicesRunning = $false
                    $HealthStatus.FailedServices += "$($Services[$Service]) ($($ServiceObj.Status))"
                    $HealthStatus.HealthIssues += "$($Services[$Service]) is not running"
                    $HealthStatus.OverallHealth = "Critical"
                }
            }
            catch {
                Write-Host "    ⚠ Could not check $($Services[$Service])" -ForegroundColor Yellow
                $HealthStatus.Notes += "Unable to check $($Services[$Service]) service"
            }
        }
        
        # Check disk space on system drive
        try {
            $DiskInfo = Get-WmiObject -Class Win32_LogicalDisk -ComputerName $DCFQDN -Filter "DeviceID='C:'" -ErrorAction Stop
            $HealthStatus.TotalSpaceGB = [math]::Round($DiskInfo.Size / 1GB, 1)
            $HealthStatus.FreeSpaceGB = [math]::Round($DiskInfo.FreeSpace / 1GB, 1)
            $HealthStatus.FreeSpacePercent = [math]::Round(($DiskInfo.FreeSpace / $DiskInfo.Size) * 100, 1)
            
            if ($HealthStatus.FreeSpacePercent -lt 10) {
                Write-Host "  ✗ Critical disk space: $($HealthStatus.FreeSpacePercent)% free" -ForegroundColor Red
                $HealthStatus.DiskSpaceOK = $false
                $HealthStatus.DiskSpaceCritical = $true
                $HealthStatus.HealthIssues += "Critical disk space: $($HealthStatus.FreeSpacePercent)% free on C:"
                $HealthStatus.OverallHealth = "Critical"
            }
            elseif ($HealthStatus.FreeSpacePercent -lt 20) {
                Write-Host "  ⚠ Warning disk space: $($HealthStatus.FreeSpacePercent)% free" -ForegroundColor Yellow
                $HealthStatus.DiskSpaceOK = $false
                $HealthStatus.DiskSpaceWarning = $true
                $HealthStatus.HealthIssues += "Low disk space: $($HealthStatus.FreeSpacePercent)% free on C:"
                if ($HealthStatus.OverallHealth -ne "Critical") {
                    $HealthStatus.OverallHealth = "Warning"
                }
            }
            else {
                Write-Host "  ✓ Disk space OK: $($HealthStatus.FreeSpacePercent)% free" -ForegroundColor Green
            }
        }
        catch {
            Write-Host "  ⚠ Could not check disk space" -ForegroundColor Yellow
            $HealthStatus.Notes += "Unable to check disk space"
        }
        
        # Check time synchronization
        try {
            $TimeInfo = Get-WmiObject -Class Win32_ComputerSystem -ComputerName $DCFQDN -ErrorAction Stop
            $LocalTime = Get-Date
            $RemoteTime = $TimeInfo.CurrentTime
            $TimeOffset = ($LocalTime - $RemoteTime).TotalSeconds
            
            $HealthStatus.TimeOffsetSeconds = [math]::Round($TimeOffset, 2)
            
            if ([Math]::Abs($TimeOffset) -gt 300) {
                Write-Host "  ✗ Time sync issue: $($HealthStatus.TimeOffsetSeconds) seconds offset" -ForegroundColor Red
                $HealthStatus.TimeSyncStatus = $false
                $HealthStatus.HealthIssues += "Time synchronization issue: $($HealthStatus.TimeOffsetSeconds) seconds offset"
                $HealthStatus.OverallHealth = "Critical"
            }
            elseif ([Math]::Abs($TimeOffset) -gt 60) {
                Write-Host "  ⚠ Time sync warning: $($HealthStatus.TimeOffsetSeconds) seconds offset" -ForegroundColor Yellow
                $HealthStatus.TimeSyncStatus = $false
                $HealthStatus.HealthIssues += "Time synchronization warning: $($HealthStatus.TimeOffsetSeconds) seconds offset"
                if ($HealthStatus.OverallHealth -ne "Critical") {
                    $HealthStatus.OverallHealth = "Warning"
                }
            }
            else {
                Write-Host "  ✓ Time sync OK: $($HealthStatus.TimeOffsetSeconds) seconds offset" -ForegroundColor Green
                $HealthStatus.TimeSyncStatus = $true
            }
            
            # Check NTP configuration
            try {
                $NTPQuery = w32tm /query /computer:$DCFQDN /source 2>$null
                if ($LASTEXITCODE -eq 0) {
                    $HealthStatus.NTPStatus = $NTPQuery
                    Write-Host "  ✓ NTP Source: $NTPQuery" -ForegroundColor Green
                }
                else {
                    $HealthStatus.NTPStatus = "Unable to query"
                    Write-Host "  ⚠ Could not query NTP source" -ForegroundColor Yellow
                }
            }
            catch {
                $HealthStatus.NTPStatus = "Query failed"
            }
        }
        catch {
            Write-Host "  ⚠ Could not check time synchronization" -ForegroundColor Yellow
            $HealthStatus.Notes += "Unable to check time synchronization"
        }
        
        # Check replication status
        try {
            Write-Host "  Checking replication..." -ForegroundColor Gray
            $ReplSummary = Get-ADReplicationFailure -Target $DCFQDN -ErrorAction Stop
            
            if ($ReplSummary) {
                foreach ($Error in $ReplSummary) {
                    $HealthStatus.ReplicationErrors += "$($Error.SourceDC) - $($Error.LastFailureMessage)"
                    Write-Host "    ✗ Replication error from $($Error.SourceDC): $($Error.LastFailureMessage)" -ForegroundColor Red
                }
                $HealthStatus.ReplicationStatus = $false
                $HealthStatus.HealthIssues += "Replication errors detected"
                $HealthStatus.OverallHealth = "Warning"
            }
            else {
                Write-Host "    ✓ Replication status OK" -ForegroundColor Green
                $HealthStatus.ReplicationStatus = $true
            }
            
            # Get last successful replication
            try {
                $ReplMeta = Get-ADReplicationAttributeMetadata -Target $DCFQDN -ObjectClass server -ErrorAction SilentlyContinue | 
                            Where-Object { $_.LastOriginatingChangeTime } | 
                            Sort-Object LastOriginatingChangeTime -Descending | 
                            Select-Object -First 1
                if ($ReplMeta) {
                    $HealthStatus.LastLogonReplication = $ReplMeta.LastOriginatingChangeTime
                }
            }
            catch {
                # Ignore - not critical
            }
        }
        catch {
            Write-Host "  ⚠ Could not check replication status" -ForegroundColor Yellow
            $HealthStatus.Notes += "Unable to check replication status"
        }
        
        # Check critical events
        if ($IncludeDetailedEvents) {
            Write-Host "  Checking recent critical events..." -ForegroundColor Gray
            try {
                $EventLogFilter = @{
                    LogName = 'System'
                    StartTime = (Get-Date).AddDays(-$DaysBack)
                    ProviderName = @('Microsoft-Windows-ActiveDirectory_DomainService', 
                                   'Microsoft-Windows-DirectoryServices-DS', 
                                   'Microsoft-Windows-Kerberos-Key-Distribution-Center',
                                   'Microsoft-Windows-DNS-Server-Service')
                    Level = 1, 2, 3  # Critical, Error, Warning
                    MaxEvents = 100
                }
                
                $Events = Get-WinEvent -ComputerName $DCFQDN -FilterHashtable $EventLogFilter -ErrorAction SilentlyContinue
                
                if ($Events) {
                    $EventCount = $Events.Count
                    $HealthStatus.EventLogStatus = $true
                    
                    # Group and summarize events
                    $Events | Group-Object ProviderName, LevelDisplayName | ForEach-Object {
                        $HealthStatus.CriticalEvents += "$($_.Name) - $($_.Count) events"
                    }
                    
                    Write-Host "    Found $EventCount critical events in the last $DaysBack days" -ForegroundColor Yellow
                }
                else {
                    Write-Host "    ✓ No critical events found in the last $DaysBack days" -ForegroundColor Green
                }
            }
            catch {
                Write-Host "    ⚠ Could not check event logs" -ForegroundColor Yellow
                $HealthStatus.Notes += "Unable to check event logs"
            }
        }
        
        # Recommendations based on health status
        if ($HealthStatus.OverallHealth -eq "Critical") {
            $HealthStatus.Recommendations += "Immediate investigation required"
            if (-not $HealthStatus.PingStatus) { $HealthStatus.Recommendations += "Check network connectivity and DC availability" }
            if ($HealthStatus.FailedServices.Count -gt 0) { $HealthStatus.Recommendations += "Restart failed services: $($HealthStatus.FailedServices -join ', ')" }
            if (-not $HealthStatus.DiskSpaceOK) { $HealthStatus.Recommendations += "Free up disk space on C: drive" }
            if (-not $HealthStatus.TimeSyncStatus) { $HealthStatus.Recommendations += "Investigate time synchronization" }
        }
        elseif ($HealthStatus.OverallHealth -eq "Warning") {
            $HealthStatus.Recommendations += "Review warnings within 24 hours"
            if ($HealthStatus.DiskSpaceWarning) { $HealthStatus.Recommendations += "Monitor disk space, plan cleanup" }
            if (-not $HealthStatus.TimeSyncStatus) { $HealthStatus.Recommendations += "Check time synchronization settings" }
            if (-not $HealthStatus.ReplicationStatus) { $HealthStatus.Recommendations += "Address replication errors" }
        }
        else {
            $HealthStatus.Recommendations += "All systems operational - continue monitoring"
        }
    }
    else {
        # DC is unreachable
        $HealthStatus.OverallHealth = "Critical"
        $HealthStatus.HealthIssues += "DC is unreachable via network"
        $HealthStatus.Recommendations += "Immediate investigation required - DC appears offline"
    }
    
    Write-Host "  Overall Health: $($HealthStatus.OverallHealth)" -ForegroundColor $(if ($HealthStatus.OverallHealth -eq "Healthy") { "Green" } elseif ($HealthStatus.OverallHealth -eq "Warning") { "Yellow" } else { "Red" })
    Write-Host ""
    
    return $HealthStatus
}

# Process each DC
foreach ($DC in $DCs) {
    $CurrentDC++
    Write-Progress -Activity "Checking Domain Controller Health" -Status "Processing $($DC.Name)" -PercentComplete (($CurrentDC / $TotalDCs) * 100)
    
    # Skip if CriticalOnly and previous DCs are healthy (implemented after processing)
    $Result = Check-DCHealth -DC $DC -ExportPath $ExportPath -IncludeDetailedEvents $IncludeDetailedEvents -DaysBack $DaysBack
    
    # Add to results
    $Results += [PSCustomObject]@{
        DCName = $Result.DCName
        FQDN = $Result.FQDN
        Site = $Result.Site
        IPv4Address = $Result.IPv4Address
        OperatingSystem = $Result.OperatingSystem
        OverallHealth = $Result.OverallHealth
        PingStatus = $Result.PingStatus
        PingResponseTime = $Result.PingResponseTime
        ServicesRunning = $Result.ServicesRunning
        FailedServices = ($Result.FailedServices -join "; ")
        DiskSpaceOK = $Result.DiskSpaceOK
        FreeSpaceGB = $Result.FreeSpaceGB
        TotalSpaceGB = $Result.TotalSpaceGB
        FreeSpacePercent = $Result.FreeSpacePercent
        TimeSyncStatus = $Result.TimeSyncStatus
        TimeOffsetSeconds = $Result.TimeOffsetSeconds
        NTPStatus = $Result.NTPStatus
        ReplicationStatus = $Result.ReplicationStatus
        ReplicationErrors = ($Result.ReplicationErrors -join "; ")
        NTDSStatus = $Result.NTDSStatus
        KDCStatus = $Result.KDCStatus
        NetLogonStatus = $Result.NetLogonStatus
        DNSStatus = $Result.DNSStatus
        W32TimeStatus = $Result.W32TimeStatus
        IsGlobalCatalog = $Result.IsGlobalCatalog
        IsReadOnly = $Result.IsReadOnly
        LastLogonReplication = $Result.LastLogonReplication
        SystemUptime = if ($Result.SystemUptime) { "$([math]::Round($Result.SystemUptime.TotalDays, 1)) days" } else { $null }
        LastBootTime = $Result.LastBootTime
        CriticalEvents = ($Result.CriticalEvents -join "; ")
        EventLogStatus = $Result.EventLogStatus
        HealthIssues = ($Result.HealthIssues -join "; ")
        Recommendations = ($Result.Recommendations -join "; ")
        Notes = ($Result.Notes -join "; ")
    }
    
    Write-Progress -Activity "Checking Domain Controller Health" -Completed
}

# If CriticalOnly flag is set, filter results
if ($CriticalOnly) {
    $Results = $Results | Where-Object { $_.OverallHealth -ne "Healthy" }
    Write-Host "CriticalOnly mode: Showing only DCs with issues" -ForegroundColor Yellow
}

# Generate timestamp
$Timestamp = Get-Date -Format "yyyyMMdd_HHmmss"

# Generate CSV report
$CSVFileName = "DCHealth_Report_$Timestamp.csv"
$CSVPath = Join-Path -Path $ExportPath -ChildPath $CSVFileName
$Results | Export-Csv -Path $CSVPath -NoTypeInformation -Encoding UTF8

Write-Host "`nCSV Report saved to: $CSVPath" -ForegroundColor Green

# Generate summary
$TotalHealthy = ($Results | Where-Object { $_.OverallHealth -eq "Healthy" }).Count
$TotalWarning = ($Results | Where-Object { $_.OverallHealth -eq "Warning" }).Count
$TotalCritical = ($Results | Where-Object { $_.OverallHealth -eq "Critical" }).Count

Write-Host "`n===== DC HEALTH REPORT SUMMARY =====" -ForegroundColor Cyan
Write-Host "Total Domain Controllers: $TotalDCs" -ForegroundColor White
Write-Host "Healthy: $TotalHealthy" -ForegroundColor Green
Write-Host "Warning: $TotalWarning" -ForegroundColor Yellow
Write-Host "Critical: $TotalCritical" -ForegroundColor Red
Write-Host "====================================`n" -ForegroundColor Cyan

# Generate HTML report if requested
if ($GenerateHTML -and $Results.Count -gt 0) {
    $HTMLFileName = "DCHealth_Report_$Timestamp.html"
    $HTMLPath = Join-Path -Path $ExportPath -ChildPath $HTMLFileName
    
    $HTMLHeader = @"
<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <title>Domain Controller Health Report - $DomainName</title>
    <style>
        body { font-family: Arial, sans-serif; margin: 20px; }
        h1 { color: #2c3e50; border-bottom: 2px solid #3498db; padding-bottom: 10px; }
        h2 { color: #34495e; margin-top: 20px; }
        .summary { background-color: #ecf0f1; padding: 15px; border-radius: 5px; margin-bottom: 20px; display: flex; justify-content: space-around; }
        .summary-item { text-align: center; }
        .summary-item .number { font-size: 24px; font-weight: bold; }
        .summary-item .label { font-size: 14px; color: #7f8c8d; }
        .healthy { color: #27ae60; }
        .warning { color: #f39c12; }
        .critical { color: #e74c3c; }
        table { border-collapse: collapse; width: 100%; margin-top: 20px; font-size: 14px; }
        th { background-color: #34495e; color: white; padding: 12px; text-align: left; position: sticky; top: 0; }
        td { padding: 10px; border-bottom: 1px solid #ddd; }
        tr:hover { background-color: #f5f5f5; }
        tr.healthy { background-color: #d5f5e3; }
        tr.warning { background-color: #fdebd0; }
        tr.critical { background-color: #fadbd8; }
        .status-badge { display: inline-block; padding: 3px 10px; border-radius: 15px; font-weight: bold; color: white; }
        .status-badge.healthy { background-color: #27ae60; }
        .status-badge.warning { background-color: #f39c12; }
        .status-badge.critical { background-color: #e74c3c; }
        .footer { margin-top: 30px; font-size: 0.9em; color: #7f8c8d; text-align: center; }
        .issues-list { font-size: 12px; color: #e74c3c; }
        .recommendations-list { font-size: 12px; color: #2980b9; }
    </style>
</head>
<body>
"@

    $HTMLSummary = @"
    <h1>Domain Controller Health Report</h1>
    <p><strong>Domain:</strong> $DomainName</p>
    <p><strong>Generated:</strong> $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')</p>
    
    <div class="summary">
        <div class="summary-item">
            <div class="number">$TotalDCs</div>
            <div class="label">Total DCs</div>
        </div>
        <div class="summary-item">
            <div class="number healthy">$TotalHealthy</div>
            <div class="label">Healthy</div>
        </div>
        <div class="summary-item">
            <div class="number warning">$TotalWarning</div>
            <div class="label">Warning</div>
        </div>
        <div class="summary-item">
            <div class="number critical">$TotalCritical</div>
            <div class="label">Critical</div>
        </div>
    </div>
"@

    # Build HTML table rows
    $HTMLRows = ""
    foreach ($DC in $Results) {
        $RowClass = $DC.OverallHealth.ToLower()
        $StatusBadge = "<span class='status-badge $RowClass'>$($DC.OverallHealth)</span>"
        
        $HTMLRows += @"
        <tr class="$RowClass">
            <td><strong>$($DC.DCName)</strong></td>
            <td>$($DC.Site)</td>
            <td>$($DC.IPv4Address)</td>
            <td>$($DC.OperatingSystem)</td>
            <td>$StatusBadge</td>
            <td>$($DC.PingResponseTime)ms</td>
            <td>$($DC.FreeSpaceGB)/$($DC.TotalSpaceGB) GB<br><small>($($DC.FreeSpacePercent)%)</small></td>
            <td>$($DC.TimeOffsetSeconds)s</td>
            <td>$($DC.ReplicationStatus)</td>
            <td>$($DC.SystemUptime)</td>
            <td class="issues-list">$(($DC.HealthIssues -split "; " | ForEach-Object { "<div>• $_</div>" }) -join '')</td>
            <td class="recommendations-list">$(($DC.Recommendations -split "; " | ForEach-Object { "<div>• $_</div>" }) -join '')</td>
        </tr>
"@
    }

    $HTMLTable = @"
    <h2>Detailed Health Status</h2>
    <div style="overflow-x: auto;">
        <table>
            <thead>
                <tr>
                    <th>DC Name</th>
                    <th>Site</th>
                    <th>IPv4</th>
                    <th>OS</th>
                    <th>Health</th>
                    <th>Ping</th>
                    <th>Disk Space</th>
                    <th>Time Sync</th>
                    <th>Replication</th>
                    <th>Uptime</th>
                    <th>Issues</th>
                    <th>Recommendations</th>
                </tr>
            </thead>
            <tbody>
                $HTMLRows
            </tbody>
        </table>
    </div>
"@

    $HTMLFooter = @"
    <div class="footer">
        <p>Report generated by Domain_Controllers_Health_Report.ps1 | $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')</p>
        <p style="font-size: 0.8em; color: #95a5a6;">* Critical events include: Error, Warning, and Critical level events from AD, DNS, and Kerberos services</p>
    </div>
</body>
</html>
"@

    $HTMLContent = $HTMLHeader + $HTMLSummary + $HTMLTable + $HTMLFooter
    $HTMLContent | Out-File -FilePath $HTMLPath -Encoding UTF8
    
    Write-Host "HTML Report saved to: $HTMLPath" -ForegroundColor Green
}

# Display quick summary in console
Write-Host "`nQuick Overview:" -ForegroundColor Yellow
$Results | Select-Object DCName, OverallHealth, @{N='Issues';E={$_.HealthIssues.Split(';')[0]}} | Format-Table -AutoSize

# Send email if requested
if ($SendEmail -and $SmtpServer -and $To -and $From) {
    try {
        $Subject = "Domain Controller Health Report - $DomainName - $(Get-Date -Format 'yyyy-MM-dd')"
        $Body = @"
Domain Controller Health Report

Total DCs: $TotalDCs
Healthy: $TotalHealthy
Warning: $TotalWarning  
Critical: $TotalCritical

Report files:
- CSV: $CSVPath
$($GenerateHTML ? "- HTML: $HTMLPath" : "")

Please review the attached report for detailed information.
"@

        $Attachment = if ($GenerateHTML) { $CSVPath, $HTMLPath } else { $CSVPath }
        
        Send-MailMessage -SmtpServer $SmtpServer -To $To -From $From -Subject $Subject -Body $Body -Attachments $Attachment -ErrorAction Stop
        
        Write-Host "Email report sent successfully to $To" -ForegroundColor Green
    }
    catch {
        Write-Warning "Failed to send email: $_"
    }
}

Write-Host "`nDomain Controller Health Report completed!" -ForegroundColor Green