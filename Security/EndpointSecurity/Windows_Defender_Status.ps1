<#
.SYNOPSIS
    Windows Defender Security Status and Health Report Script
.DESCRIPTION
    This script performs a comprehensive audit of Windows Defender security status,
    including protection settings, threat history, signature updates, real-time
    protection, and overall system security health. It generates detailed reports
    with actionable recommendations.
.NOTES
    Author: Your Name
    Date: 2026-08-22
    Version: 1.0
    Requires: Administrative privileges
    Features: Health check, signature status, threat detection, configuration audit
.LINK
    https://docs.microsoft.com/en-us/windows/security/threat-protection/windows-defender-antivirus/
#>

<#
Basic Usage:
powershell

# Run as Administrator
.\Windows_Defender_Status.ps1

Advanced Usage Examples:
powershell

# Comprehensive audit with all features
.\Windows_Defender_Status.ps1 -CheckThreatHistory -CheckExclusions -CheckPerformance

# Generate specific report format
.\Windows_Defender_Status.ps1 -OutputFormat HTML

# Custom output location
.\Windows_Defender_Status.ps1 -OutputPath "D:\SecurityReports\Defender" -CheckThreatHistory

# Full security assessment
.\Windows_Defender_Status.ps1 -CheckThreatHistory -CheckExclusions -CheckPerformance -OutputFormat All

# Quick status check
.\Windows_Defender_Status.ps1 -OutputFormat HTML -CheckPerformance

Key Features:
🛡️ Status Monitoring:

    ✅ Real-time protection status

    ✅ Signature update age and version

    ✅ Cloud protection status

    ✅ Network protection

    ✅ Tamper protection

    ✅ Scan history (quick and full)

🔍 Security Analysis:

    ✅ Threat detection history

    ✅ Quarantined items

    ✅ Exclusion lists (paths, extensions, processes)

    ✅ Performance metrics (CPU, memory)

    ✅ Risk score calculation

    ✅ Overall health assessment

📊 Report Generation:

    HTML: Comprehensive visual dashboard

    CSV: Detailed data exports

    JSON: Programmatic processing

    All: Complete reporting suite

🎯 Health Categories:

    Excellent: All protections enabled, updated

    Good: Minor issues, generally secure

    Fair: Some protections disabled or outdated

    Poor: Multiple critical issues

    Critical: Major security gaps

Sample Use Cases:
1. Security Compliance Audit
powershell

# Check security posture
.\Windows_Defender_Status.ps1 -CheckThreatHistory -CheckExclusions

2. Incident Response
powershell

# Check for recent threats
.\Windows_Defender_Status.ps1 -CheckThreatHistory -OutputPath "C:\IR\Defender_$(Get-Date -Format 'yyyyMMdd_HHmm')"

3. Performance Monitoring
powershell

# Check performance impact
.\Windows_Defender_Status.ps1 -CheckPerformance -OutputFormat HTML

4. Regular Health Monitoring
powershell

# Scheduled task friendly
.\Windows_Defender_Status.ps1 -CheckThreatHistory -OutputFormat HTML -OutputPath "D:\Monitoring\Defender"

Output Examples:
HTML Report Includes:

    Overall health dashboard

    Risk score visualization

    Detailed configuration table

    Threat history (if enabled)

    Exclusion lists (if enabled)

    Performance metrics (if enabled)

    Actionable recommendations

    Best practices section

Status Checks Performed:

    Real-time Protection - Is it enabled?

    Signature Updates - Are definitions current?

    Scan History - When was last scan?

    Threat Status - Any active threats?

    Cloud Protection - Is cloud protection enabled?

    Sample Submission - Automatic sample submission?

    Exclusions - Too many exclusions?

Risk Score Breakdown:

    Real-time protection off: +30

    Outdated signatures: +10-20

    Active threats: +25

    Cloud protection off: +10

    Sample submission off: +5

    Excessive exclusions: +5

Security Recommendations:

    Enable real-time protection

    Update signatures

    Perform system scans

    Enable cloud protection

    Review exclusions

    Check tamper protection

    Monitor threat detection
#>

#Requires -RunAsAdministrator

[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [string]$OutputPath = "$env:USERPROFILE\Desktop\Windows_Defender_Report_$(Get-Date -Format 'yyyyMMdd_HHmmss')",
    
    [Parameter(Mandatory = $false)]
    [ValidateSet('CSV', 'HTML', 'JSON', 'All')]
    [string]$OutputFormat = 'All',
    
    [Parameter(Mandatory = $false)]
    [switch]$CheckThreatHistory,
    
    [Parameter(Mandatory = $false)]
    [switch]$CheckExclusions,
    
    [Parameter(Mandatory = $false)]
    [switch]$CheckPerformance,
    
    [Parameter(Mandatory = $false)]
    [string]$LogPath = "$env:USERPROFILE\Desktop\Windows_Defender_Log_$(Get-Date -Format 'yyyyMMdd_HHmmss').txt"
)

# Script configuration
$ErrorActionPreference = "Continue"
$ScriptName = "Windows_Defender_Status"
$ScriptVersion = "1.0"
$ExecutionTime = Get-Date

# Initialize logging
function Write-Log {
    param(
        [string]$Message,
        [string]$Level = "INFO"
    )
    $TimeStamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $LogEntry = "[$TimeStamp] [$Level] $Message"
    Write-Host $LogEntry
    Add-Content -Path $LogPath -Value $LogEntry -ErrorAction SilentlyContinue
}

# Display header
Write-Host "=" * 80 -ForegroundColor Cyan
Write-Host "Windows Defender Security Status v$ScriptVersion" -ForegroundColor Cyan
Write-Host "Generated on: $ExecutionTime" -ForegroundColor Cyan
Write-Host "=" * 80 -ForegroundColor Cyan
Write-Host ""

# Check if Windows Defender is installed and available
function Test-DefenderAvailable {
    try {
        $DefenderStatus = Get-MpComputerStatus -ErrorAction SilentlyContinue
        if ($DefenderStatus) {
            Write-Log "Windows Defender is available" "INFO"
            return $true
        }
        
        # Alternative check for older Windows versions
        $DefenderService = Get-Service -Name 'WinDefend' -ErrorAction SilentlyContinue
        if ($DefenderService) {
            Write-Log "Windows Defender service found" "INFO"
            return $true
        }
        
        Write-Log "Windows Defender not found on this system" "WARNING"
        return $false
    }
    catch {
        Write-Log "Error checking Windows Defender availability: $_" "WARNING"
        return $false
    }
}

# Function to get comprehensive Defender status
function Get-DefenderStatus {
    try {
        Write-Log "Retrieving Windows Defender status..." "INFO"
        $DefenderStatus = Get-MpComputerStatus -ErrorAction Stop
        
        # Check if required module is available
        if (-not $DefenderStatus) {
            Write-Log "Unable to retrieve Defender status. Try importing SecurityCenter module." "WARNING"
            try {
                Import-Module SecurityCenter -ErrorAction SilentlyContinue
                $DefenderStatus = Get-MpComputerStatus -ErrorAction Stop
            }
            catch {
                Write-Log "Failed to import SecurityCenter module" "ERROR"
                return $null
            }
        }
        
        return $DefenderStatus
    }
    catch {
        Write-Log "Error getting Defender status: $_" "ERROR"
        return $null
    }
}

# Function to get threat history
function Get-ThreatHistory {
    $Threats = @()
    
    if (-not $CheckThreatHistory) {
        return $Threats
    }
    
    try {
        Write-Log "Retrieving threat history..." "INFO"
        
        # Try to get threats from Windows Defender
        $Threats = Get-MpThreatDetection -ErrorAction SilentlyContinue
        
        if ($Threats) {
            Write-Log "Found $($Threats.Count) threat detections" "INFO"
        }
        
        # Also check Windows Event Log for additional threat information
        try {
            $Events = Get-WinEvent -LogName "Microsoft-Windows-Windows Defender/Operational" -MaxEvents 100 -ErrorAction SilentlyContinue | Where-Object {
                $_.Id -in @(1000, 1001, 1002, 1116, 1117, 1118, 1119, 1120, 1121, 1122, 1123, 1124, 1125, 1126, 1127, 1128, 1129, 1130)
            }
            
            if ($Events) {
                foreach ($Event in $Events) {
                    $Threats += [PSCustomObject]@{
                        ComputerName = $env:COMPUTERNAME
                        ThreatName = (Get-EventLogEntryMessage -Event $Event -Property "Threat Name")
                        ThreatPath = (Get-EventLogEntryMessage -Event $Event -Property "Path")
                        DetectionTime = $Event.TimeCreated
                        Category = (Get-EventLogEntryMessage -Event $Event -Property "Category")
                        EventID = $Event.Id
                        Action = (Get-EventLogEntryMessage -Event $Event -Property "Action")
                    }
                }
            }
        }
        catch {
            Write-Log "Error retrieving threat events from Event Log: $_" "WARNING"
        }
    }
    catch {
        Write-Log "Error retrieving threat history: $_" "ERROR"
    }
    
    return $Threats
}

# Helper function to extract message properties
function Get-EventLogEntryMessage {
    param(
        [System.Diagnostics.Eventing.Reader.EventLogRecord]$Event,
        [string]$Property
    )
    
    try {
        $Message = $Event.Message
        if ($Message) {
            $Lines = $Message -split "`r`n"
            foreach ($Line in $Lines) {
                if ($Line -match "$Property\s*:\s*(.+)$") {
                    return $Matches[1].Trim()
                }
            }
        }
    }
    catch {
        return "N/A"
    }
    return "N/A"
}

# Function to get exclusions
function Get-DefenderExclusions {
    $Exclusions = @()
    
    if (-not $CheckExclusions) {
        return $Exclusions
    }
    
    try {
        Write-Log "Retrieving exclusion configurations..." "INFO"
        
        $PathExclusions = Get-MpPreference -ExclusionPath -ErrorAction SilentlyContinue
        $ExtensionExclusions = Get-MpPreference -ExclusionExtension -ErrorAction SilentlyContinue
        $ProcessExclusions = Get-MpPreference -ExclusionProcess -ErrorAction SilentlyContinue
        
        foreach ($Path in $PathExclusions) {
            $Exclusions += [PSCustomObject]@{
                Type = "Path"
                Value = $Path
                Description = "File/Path exclusion"
            }
        }
        
        foreach ($Extension in $ExtensionExclusions) {
            $Exclusions += [PSCustomObject]@{
                Type = "Extension"
                Value = $Extension
                Description = "File extension exclusion"
            }
        }
        
        foreach ($Process in $ProcessExclusions) {
            $Exclusions += [PSCustomObject]@{
                Type = "Process"
                Value = $Process
                Description = "Process exclusion"
            }
        }
        
        Write-Log "Found $($Exclusions.Count) exclusions" "INFO"
    }
    catch {
        Write-Log "Error retrieving exclusions: $_" "ERROR"
    }
    
    return $Exclusions
}

# Function to check performance
function Get-DefenderPerformance {
    $Performance = @()
    
    if (-not $CheckPerformance) {
        return $Performance
    }
    
    try {
        Write-Log "Checking Defender performance..." "INFO"
        
        # Check CPU usage of Defender processes
        $DefenderProcesses = @('MsMpEng', 'NisSrv', 'SecurityHealthService')
        foreach ($ProcessName in $DefenderProcesses) {
            $Process = Get-Process -Name $ProcessName -ErrorAction SilentlyContinue
            if ($Process) {
                $Performance += [PSCustomObject]@{
                    ProcessName = $ProcessName
                    CPU = [math]::Round($Process.CPU, 2)
                    MemoryMB = [math]::Round($Process.WorkingSet / 1MB, 2)
                    Handles = $Process.HandleCount
                    Threads = $Process.Threads.Count
                    Status = "Running"
                }
            }
            else {
                $Performance += [PSCustomObject]@{
                    ProcessName = $ProcessName
                    CPU = 0
                    MemoryMB = 0
                    Handles = 0
                    Threads = 0
                    Status = "Not Running"
                }
            }
        }
        
        # Check disk usage of Defender
        $DefenderPath = "$env:ProgramData\Microsoft\Windows Defender"
        if (Test-Path $DefenderPath) {
            $FolderSize = Get-ChildItem -Path $DefenderPath -Recurse -ErrorAction SilentlyContinue | Measure-Object -Property Length -Sum
            $Performance += [PSCustomObject]@{
                ProcessName = "Defender Storage"
                CPU = 0
                MemoryMB = [math]::Round($FolderSize.Sum / 1MB, 2)
                Handles = 0
                Threads = 0
                Status = "Active"
                Note = "Storage size on disk"
            }
        }
        
        Write-Log "Performance check completed" "INFO"
    }
    catch {
        Write-Log "Error checking performance: $_" "ERROR"
    }
    
    return $Performance
}

# Function to evaluate security health
function Evaluate-SecurityHealth {
    param(
        $DefenderStatus,
        [array]$Threats,
        [array]$Exclusions
    )
    
    $HealthChecks = @()
    $RiskScore = 0
    
    # Check real-time protection
    if ($DefenderStatus.RealTimeProtectionEnabled) {
        $HealthChecks += [PSCustomObject]@{
            Check = "Real-time Protection"
            Status = "PASS"
            Details = "Real-time protection is enabled"
            Recommendation = "Maintain current configuration"
        }
    } else {
        $HealthChecks += [PSCustomObject]@{
            Check = "Real-time Protection"
            Status = "FAIL"
            Details = "Real-time protection is DISABLED"
            Recommendation = "CRITICAL: Enable real-time protection immediately"
        }
        $RiskScore += 30
    }
    
    # Check signature update
    $SignatureAge = (Get-Date) - $DefenderStatus.AntivirusSignatureLastUpdated
    if ($SignatureAge.TotalHours -le 24) {
        $HealthChecks += [PSCustomObject]@{
            Check = "Signature Updates"
            Status = "PASS"
            Details = "Definitions updated within last 24 hours (Updated: $($SignatureAge.TotalHours.ToString('0.0')) hours ago)"
            Recommendation = "Continue regular updates"
        }
    } elseif ($SignatureAge.TotalHours -le 72) {
        $HealthChecks += [PSCustomObject]@{
            Check = "Signature Updates"
            Status = "WARNING"
            Details = "Definitions are $($SignatureAge.TotalHours.ToString('0.0')) hours old"
            Recommendation = "Update definitions within next 24 hours"
        }
        $RiskScore += 10
    } else {
        $HealthChecks += [PSCustomObject]@{
            Check = "Signature Updates"
            Status = "FAIL"
            Details = "Definitions are OUTDATED ($($SignatureAge.TotalHours.ToString('0.0')) hours old)"
            Recommendation = "CRITICAL: Update definitions immediately"
        }
        $RiskScore += 20
    }
    
    # Check scan schedule
    $LastScanDate = $DefenderStatus.QuickScanLastEndTime
    $LastFullScan = $DefenderStatus.FullScanLastEndTime
    
    if ($LastScanDate) {
        $ScanAge = (Get-Date) - $LastScanDate
        if ($ScanAge.TotalDays -le 7) {
            $HealthChecks += [PSCustomObject]@{
                Check = "Recent Scan"
                Status = "PASS"
                Details = "Last quick scan completed $($ScanAge.TotalDays.ToString('0')) days ago"
                Recommendation = "Maintain regular scanning schedule"
            }
        } else {
            $HealthChecks += [PSCustomObject]@{
                Check = "Recent Scan"
                Status = "WARNING"
                Details = "No quick scan in $($ScanAge.TotalDays.ToString('0')) days"
                Recommendation = "Perform a quick scan soon"
            }
            $RiskScore += 5
        }
    } else {
        $HealthChecks += [PSCustomObject]@{
            Check = "Recent Scan"
            Status = "WARNING"
            Details = "No scan history found"
            Recommendation = "Perform initial system scan"
        }
        $RiskScore += 5
    }
    
    # Check for active threats
    if ($DefenderStatus.ThreatsFound -gt 0) {
        $HealthChecks += [PSCustomObject]@{
            Check = "Threat Status"
            Status = "FAIL"
            Details = "$($DefenderStatus.ThreatsFound) threats detected"
            Recommendation = "Review and remediate detected threats immediately"
        }
        $RiskScore += 25
    } else {
        $HealthChecks += [PSCustomObject]@{
            Check = "Threat Status"
            Status = "PASS"
            Details = "No active threats detected"
            Recommendation = "Continue monitoring"
        }
    }
    
    # Check cloud protection
    if ($DefenderStatus.CloudProtectionEnabled) {
        $HealthChecks += [PSCustomObject]@{
            Check = "Cloud Protection"
            Status = "PASS"
            Details = "Cloud-delivered protection is enabled"
            Recommendation = "Maintain cloud protection"
        }
    } else {
        $HealthChecks += [PSCustomObject]@{
            Check = "Cloud Protection"
            Status = "WARNING"
            Details = "Cloud-delivered protection is DISABLED"
            Recommendation = "Enable cloud protection for better detection"
        }
        $RiskScore += 10
    }
    
    # Check sample submission
    if ($DefenderStatus.SampleSubmissionEnabled) {
        $HealthChecks += [PSCustomObject]@{
            Check = "Sample Submission"
            Status = "PASS"
            Details = "Automatic sample submission is enabled"
            Recommendation = "Maintain current configuration"
        }
    } else {
        $HealthChecks += [PSCustomObject]@{
            Check = "Sample Submission"
            Status = "WARNING"
            Details = "Automatic sample submission is DISABLED"
            Recommendation = "Enable sample submission for better protection"
        }
        $RiskScore += 5
    }
    
    # Check exclusions
    if ($Exclusions.Count -gt 10) {
        $HealthChecks += [PSCustomObject]@{
            Check = "Exclusions"
            Status = "WARNING"
            Details = "$($Exclusions.Count) exclusions configured"
            Recommendation = "Review exclusions for security impact"
        }
        $RiskScore += 5
    }
    
    # Determine overall health
    $OverallHealth = switch ($RiskScore) {
        { $_ -eq 0 } { "Excellent" }
        { $_ -le 10 } { "Good" }
        { $_ -le 25 } { "Fair" }
        { $_ -le 50 } { "Poor" }
        default { "Critical" }
    }
    
    return @{
        HealthChecks = $HealthChecks
        RiskScore = $RiskScore
        OverallHealth = $OverallHealth
    }
}

# Function to generate HTML report
function Export-HTMLReport {
    param(
        $DefenderStatus,
        [array]$Threats,
        [array]$Exclusions,
        [array]$Performance,
        $HealthAnalysis,
        [string]$OutputPath
    )
    
    $HTMLPath = "$OutputPath.html"
    
    # Calculate additional stats
    $DefenderVersion = $DefenderStatus.AntivirusVersion
    $SignatureVersion = $DefenderStatus.AntivirusSignatureVersion
    $LastUpdate = $DefenderStatus.AntivirusSignatureLastUpdated
    $RealTimeEnabled = $DefenderStatus.RealTimeProtectionEnabled
    $CloudEnabled = $DefenderStatus.CloudProtectionEnabled
    
    $HTML = @"
<!DOCTYPE html>
<html>
<head>
    <title>Windows Defender Security Report - $($env:COMPUTERNAME)</title>
    <style>
        body { font-family: Arial, sans-serif; margin: 20px; background-color: #f5f5f5; }
        .header { background: linear-gradient(135deg, #0078d4 0%, #004a8c 100%); color: white; padding: 20px; border-radius: 5px; }
        .header h1 { margin: 0; }
        .summary { display: grid; grid-template-columns: repeat(auto-fit, minmax(150px, 1fr)); gap: 15px; margin: 20px 0; }
        .summary-card { background: white; padding: 15px; border-radius: 5px; box-shadow: 0 2px 4px rgba(0,0,0,0.1); text-align: center; }
        .summary-card .number { font-size: 28px; font-weight: bold; margin: 10px 0; }
        .summary-card .label { color: #666; font-size: 14px; }
        .status-pass { color: #4CAF50; }
        .status-fail { color: #f44336; }
        .status-warning { color: #ff9800; }
        .section { background: white; padding: 20px; border-radius: 5px; margin: 20px 0; box-shadow: 0 2px 4px rgba(0,0,0,0.1); }
        table { width: 100%; border-collapse: collapse; }
        th { background: #0078d4; color: white; padding: 10px; text-align: left; }
        td { padding: 8px; border-bottom: 1px solid #ddd; }
        tr:hover { background: #f5f5f5; }
        .badge { padding: 3px 8px; border-radius: 3px; font-weight: bold; font-size: 12px; }
        .badge-pass { background: #4CAF50; color: white; }
        .badge-fail { background: #f44336; color: white; }
        .badge-warning { background: #ff9800; color: white; }
        .badge-info { background: #2196F3; color: white; }
        .health-excellent { color: #4CAF50; }
        .health-good { color: #8BC34A; }
        .health-fair { color: #ff9800; }
        .health-poor { color: #f44336; }
        .health-critical { color: #d32f2f; font-weight: bold; }
        .footer { margin-top: 20px; padding: 10px; background: #333; color: white; border-radius: 5px; text-align: center; }
        .recommendation-box { background: #fff3cd; border-left: 4px solid #ff9800; padding: 10px; margin: 10px 0; }
        .critical-box { background: #ffebee; border-left: 4px solid #f44336; padding: 10px; margin: 10px 0; }
        .success-box { background: #e8f5e9; border-left: 4px solid #4CAF50; padding: 10px; margin: 10px 0; }
        .info-box { background: #e3f2fd; border-left: 4px solid #2196F3; padding: 10px; margin: 10px 0; }
    </style>
</head>
<body>
    <div class="header">
        <h1>🛡️ Windows Defender Security Status Report</h1>
        <p>Generated: $ExecutionTime</p>
        <p>Computer: $($env:COMPUTERNAME)</p>
        <p>Script Version: $ScriptVersion</p>
    </div>
    
    <div class="summary">
        <div class="summary-card">
            <div class="label">Overall Health</div>
            <div class="number health-$($HealthAnalysis.OverallHealth.ToLower())">$($HealthAnalysis.OverallHealth)</div>
        </div>
        <div class="summary-card">
            <div class="label">Risk Score</div>
            <div class="number $(if ($HealthAnalysis.RiskScore -gt 25) { 'status-fail' } elseif ($HealthAnalysis.RiskScore -gt 10) { 'status-warning' } else { 'status-pass' })">$($HealthAnalysis.RiskScore)/100</div>
        </div>
        <div class="summary-card">
            <div class="label">Real-time Protection</div>
            <div class="number $($RealTimeEnabled ? 'status-pass' : 'status-fail')">$($RealTimeEnabled ? '✅ On' : '❌ Off')</div>
        </div>
        <div class="summary-card">
            <div class="label">Cloud Protection</div>
            <div class="number $($CloudEnabled ? 'status-pass' : 'status-warning')">$($CloudEnabled ? '✅ On' : '❌ Off')</div>
        </div>
        <div class="summary-card">
            <div class="label">Threats Detected</div>
            <div class="number $(if ($DefenderStatus.ThreatsFound -gt 0) { 'status-fail' } else { 'status-pass' })">$($DefenderStatus.ThreatsFound)</div>
        </div>
    </div>
    
    <div class="section">
        <h2>📊 Security Health Checks</h2>
        <table>
            <thead>
                <tr>
                    <th>Check</th>
                    <th>Status</th>
                    <th>Details</th>
                    <th>Recommendation</th>
                </tr>
            </thead>
            <tbody>
"@
    
    foreach ($Check in $HealthAnalysis.HealthChecks) {
        $StatusClass = switch ($Check.Status) {
            'PASS' { 'badge-pass' }
            'FAIL' { 'badge-fail' }
            'WARNING' { 'badge-warning' }
            default { 'badge-info' }
        }
        $HTML += @"
                <tr>
                    <td>$($Check.Check)</td>
                    <td><span class="badge $StatusClass">$($Check.Status)</span></td>
                    <td>$($Check.Details)</td>
                    <td>$($Check.Recommendation)</td>
                </tr>
"@
    }
    
    $HTML += @"
            </tbody>
        </table>
    </div>
    
    <div class="section">
        <h2>⚙️ Configuration Details</h2>
        <table>
            <tr><td><strong>Product Version</strong></td><td>$DefenderVersion</td></tr>
            <tr><td><strong>Signature Version</strong></td><td>$SignatureVersion</td></tr>
            <tr><td><strong>Last Update</strong></td><td>$($LastUpdate.ToString('yyyy-MM-dd HH:mm:ss'))</td></tr>
            <tr><td><strong>Real-time Protection</strong></td><td>$($RealTimeEnabled ? 'Enabled' : 'Disabled')</td></tr>
            <tr><td><strong>Cloud Protection</strong></td><td>$($CloudEnabled ? 'Enabled' : 'Disabled')</td></tr>
            <tr><td><strong>Network Protection</strong></td><td>$($DefenderStatus.NetworkProtectionEnabled ? 'Enabled' : 'Disabled')</td></tr>
            <tr><td><strong>Tamper Protection</strong></td><td>$($DefenderStatus.TamperProtectionEnabled ? 'Enabled' : 'Disabled')</td></tr>
            <tr><td><strong>Quarantine Items</strong></td><td>$($DefenderStatus.QuarantinedThreats)</td></tr>
            <tr><td><strong>Last Quick Scan</strong></td><td>$($DefenderStatus.QuickScanLastEndTime.ToString('yyyy-MM-dd HH:mm:ss') ?? 'Never')</td></tr>
            <tr><td><strong>Last Full Scan</strong></td><td>$($DefenderStatus.FullScanLastEndTime.ToString('yyyy-MM-dd HH:mm:ss') ?? 'Never')</td></tr>
        </table>
    </div>
"@
    
    if ($Exclusions.Count -gt 0) {
        $HTML += @"
    <div class="section">
        <h2>🚫 Exclusions</h2>
        <table>
            <thead>
                <tr>
                    <th>Type</th>
                    <th>Value</th>
                    <th>Description</th>
                </tr>
            </thead>
            <tbody>
"@
        foreach ($Exclusion in $Exclusions) {
            $HTML += @"
                <tr>
                    <td>$($Exclusion.Type)</td>
                    <td>$($Exclusion.Value)</td>
                    <td>$($Exclusion.Description)</td>
                </tr>
"@
        }
        $HTML += @"
            </tbody>
        </table>
    </div>
"@
    }
    
    if ($Performance.Count -gt 0) {
        $HTML += @"
    <div class="section">
        <h2>⚡ Performance Metrics</h2>
        <table>
            <thead>
                <tr>
                    <th>Process</th>
                    <th>Status</th>
                    <th>CPU Usage</th>
                    <th>Memory (MB)</th>
                    <th>Handles</th>
                    <th>Threads</th>
                </tr>
            </thead>
            <tbody>
"@
        foreach ($Item in $Performance) {
            $StatusClass = if ($Item.Status -eq 'Running') { 'badge-pass' } else { 'badge-warning' }
            $HTML += @"
                <tr>
                    <td>$($Item.ProcessName)</td>
                    <td><span class="badge $StatusClass">$($Item.Status)</span></td>
                    <td>$($Item.CPU)%</td>
                    <td>$($Item.MemoryMB)</td>
                    <td>$($Item.Handles)</td>
                    <td>$($Item.Threads)</td>
                </tr>
"@
        }
        $HTML += @"
            </tbody>
        </table>
    </div>
"@
    }
    
    if ($Threats.Count -gt 0) {
        $HTML += @"
    <div class="section">
        <h2>⚠️ Threat History</h2>
        <table>
            <thead>
                <tr>
                    <th>Date</th>
                    <th>Threat Name</th>
                    <th>Category</th>
                    <th>Path</th>
                    <th>Action</th>
                </tr>
            </thead>
            <tbody>
"@
        foreach ($Threat in $Threats | Sort-Object DetectionTime -Descending | Select-Object -First 20) {
            $HTML += @"
                <tr>
                    <td>$($Threat.DetectionTime.ToString('yyyy-MM-dd HH:mm'))</td>
                    <td>$($Threat.ThreatName)</td>
                    <td>$($Threat.Category)</td>
                    <td>$($Threat.ThreatPath)</td>
                    <td>$($Threat.Action)</td>
                </tr>
"@
        }
        $HTML += @"
            </tbody>
        </table>
    </div>
"@
    }
    
    # Recommendations section
    $HTML += @"
    <div class="section">
        <h2>💡 Security Recommendations</h2>
"@
    
    $CriticalIssues = $HealthAnalysis.HealthChecks | Where-Object { $_.Status -eq 'FAIL' }
    $WarningIssues = $HealthAnalysis.HealthChecks | Where-Object { $_.Status -eq 'WARNING' }
    
    if ($CriticalIssues.Count -gt 0) {
        $HTML += @"
        <div class="critical-box">
            <h3>🔴 Critical Issues to Address</h3>
            <ul>
"@
        foreach ($Issue in $CriticalIssues) {
            $HTML += @"
                <li><strong>$($Issue.Check):</strong> $($Issue.Recommendation)</li>
"@
        }
        $HTML += @"
            </ul>
        </div>
"@
    }
    
    if ($WarningIssues.Count -gt 0) {
        $HTML += @"
        <div class="recommendation-box">
            <h3>🟡 Warning Issues to Review</h3>
            <ul>
"@
        foreach ($Issue in $WarningIssues) {
            $HTML += @"
                <li><strong>$($Issue.Check):</strong> $($Issue.Recommendation)</li>
"@
        }
        $HTML += @"
            </ul>
        </div>
"@
    }
    
    if ($CriticalIssues.Count -eq 0 -and $WarningIssues.Count -eq 0) {
        $HTML += @"
        <div class="success-box">
            <h3>✅ No Critical Issues Found</h3>
            <p>Windows Defender is properly configured and up to date. Continue regular maintenance and monitoring.</p>
        </div>
"@
    }
    
    $HTML += @"
        <div class="info-box">
            <h3>📋 Best Practices</h3>
            <ul>
                <li>Keep Windows Defender and signatures updated</li>
                <li>Schedule regular full system scans</li>
                <li>Enable real-time and cloud-delivered protection</li>
                <li>Review exclusions periodically</li>
                <li>Monitor threat detection logs</li>
                <li>Use group policies for enterprise management</li>
            </ul>
        </div>
    </div>
    
    <div class="footer">
        <p>Generated by $ScriptName v$ScriptVersion | Windows Defender Security Audit Tool</p>
        <p>⚠ All findings should be verified and validated before action</p>
    </div>
</body>
</html>
"@
    
    $HTML | Out-File -FilePath $HTMLPath -Encoding UTF8
    Write-Log "HTML report exported to: $HTMLPath" "INFO"
    return $HTMLPath
}

# Main function
function Invoke-DefenderAudit {
    Write-Log "Starting Windows Defender Security Audit..." "INFO"
    
    # Check if Defender is available
    if (-not (Test-DefenderAvailable)) {
        Write-Log "Windows Defender is not available on this system" "ERROR"
        Write-Host "⚠ Windows Defender not found on this system" -ForegroundColor Yellow
        Write-Host "This script requires Windows Defender Antivirus" -ForegroundColor Yellow
        return $false
    }
    
    # Get Defender status
    $DefenderStatus = Get-DefenderStatus
    if (-not $DefenderStatus) {
        Write-Log "Failed to retrieve Windows Defender status" "ERROR"
        return $false
    }
    
    Write-Host ""
    Write-Host "=" * 60 -ForegroundColor Cyan
    Write-Host "Windows Defender Status" -ForegroundColor Cyan
    Write-Host "=" * 60 -ForegroundColor Cyan
    
    # Display basic status
    Write-Host "Product Version: $($DefenderStatus.AntivirusVersion)" -ForegroundColor White
    Write-Host "Signature Version: $($DefenderStatus.AntivirusSignatureVersion)" -ForegroundColor White
    $LastUpdate = $DefenderStatus.AntivirusSignatureLastUpdated
    $UpdateAge = (Get-Date) - $LastUpdate
    Write-Host "Last Update: $($LastUpdate.ToString('yyyy-MM-dd HH:mm:ss')) ($($UpdateAge.TotalHours.ToString('0.0')) hours ago)" -ForegroundColor White
    Write-Host "Real-time Protection: $($DefenderStatus.RealTimeProtectionEnabled ? '✅ Enabled' : '❌ Disabled')" -ForegroundColor $(if ($DefenderStatus.RealTimeProtectionEnabled) { 'Green' } else { 'Red' })
    Write-Host "Cloud Protection: $($DefenderStatus.CloudProtectionEnabled ? '✅ Enabled' : '❌ Disabled')" -ForegroundColor $(if ($DefenderStatus.CloudProtectionEnabled) { 'Green' } else { 'Yellow' })
    Write-Host "Threats Found: $($DefenderStatus.ThreatsFound)" -ForegroundColor $(if ($DefenderStatus.ThreatsFound -gt 0) { 'Red' } else { 'Green' })
    Write-Host "Quarantined Items: $($DefenderStatus.QuarantinedThreats)" -ForegroundColor White
    Write-Host ""
    
    # Get additional data
    $Threats = Get-ThreatHistory
    $Exclusions = Get-DefenderExclusions
    $Performance = Get-DefenderPerformance
    
    # Evaluate security health
    $HealthAnalysis = Evaluate-SecurityHealth -DefenderStatus $DefenderStatus -Threats $Threats -Exclusions $Exclusions
    
    # Export results
    $ExportedFiles = @()
    
    # Export to CSV
    if ($OutputFormat -in @('CSV', 'All')) {
        $CSVPath = "$OutputPath`_Status.csv"
        $DefenderStatus.PSObject.Properties | Select-Object Name, Value | Export-Csv -Path $CSVPath -NoTypeInformation -Encoding UTF8
        Write-Log "CSV report exported to: $CSVPath" "INFO"
        $ExportedFiles += $CSVPath
        
        if ($Threats.Count -gt 0) {
            $ThreatCSV = "$OutputPath`_Threats.csv"
            $Threats | Export-Csv -Path $ThreatCSV -NoTypeInformation -Encoding UTF8
            Write-Log "Threat history exported to: $ThreatCSV" "INFO"
            $ExportedFiles += $ThreatCSV
        }
        
        if ($Exclusions.Count -gt 0) {
            $ExclCSV = "$OutputPath`_Exclusions.csv"
            $Exclusions | Export-Csv -Path $ExclCSV -NoTypeInformation -Encoding UTF8
            Write-Log "Exclusions exported to: $ExclCSV" "INFO"
            $ExportedFiles += $ExclCSV
        }
    }
    
    # Export to HTML
    if ($OutputFormat -in @('HTML', 'All')) {
        $HTMLPath = Export-HTMLReport -DefenderStatus $DefenderStatus -Threats $Threats -Exclusions $Exclusions -Performance $Performance -HealthAnalysis $HealthAnalysis -OutputPath $OutputPath
        $ExportedFiles += $HTMLPath
    }
    
    # Export to JSON
    if ($OutputFormat -in @('JSON', 'All')) {
        $JSONPath = "$OutputPath`_Defender.json"
        $DefenderStatus | ConvertTo-Json -Depth 10 | Out-File -FilePath $JSONPath -Encoding UTF8
        Write-Log "JSON report exported to: $JSONPath" "INFO"
        $ExportedFiles += $JSONPath
    }
    
    # Display summary
    Write-Host ""
    Write-Host "=" * 80 -ForegroundColor Cyan
    Write-Host "AUDIT SUMMARY" -ForegroundColor Cyan
    Write-Host "=" * 80 -ForegroundColor Cyan
    
    Write-Host "Overall Health: $($HealthAnalysis.OverallHealth)" -ForegroundColor $(
        if ($HealthAnalysis.OverallHealth -eq 'Excellent') { 'Green' } 
        elseif ($HealthAnalysis.OverallHealth -eq 'Good') { 'Cyan' }
        elseif ($HealthAnalysis.OverallHealth -eq 'Fair') { 'Yellow' }
        elseif ($HealthAnalysis.OverallHealth -eq 'Poor') { 'Red' }
        else { 'Red' }
    )
    Write-Host "Risk Score: $($HealthAnalysis.RiskScore)/100" -ForegroundColor $(
        if ($HealthAnalysis.RiskScore -le 10) { 'Green' }
        elseif ($HealthAnalysis.RiskScore -le 25) { 'Yellow' }
        else { 'Red' }
    )
    Write-Host ""
    
    Write-Host "Health Checks:" -ForegroundColor Cyan
    foreach ($Check in $HealthAnalysis.HealthChecks) {
        $StatusColor = switch ($Check.Status) {
            'PASS' { 'Green' }
            'FAIL' { 'Red' }
            'WARNING' { 'Yellow' }
            default { 'Gray' }
        }
        Write-Host "  [$($Check.Status)] $($Check.Check): $($Check.Details)" -ForegroundColor $StatusColor
    }
    Write-Host ""
    
    Write-Host "Reports Generated:" -ForegroundColor Cyan
    foreach ($File in $ExportedFiles) {
        Write-Host "  - $File" -ForegroundColor White
    }
    Write-Host ""
    Write-Host "Log File: $LogPath" -ForegroundColor Gray
    
    Write-Log "Windows Defender audit completed successfully" "INFO"
    return $true
}

# Main execution
try {
    $Result = Invoke-DefenderAudit
    if ($Result) {
        Write-Host "`n✓ Windows Defender Security Audit completed successfully!" -ForegroundColor Green
        exit 0
    } else {
        Write-Host "`n⚠ Audit completed with errors. Check the log for details." -ForegroundColor Yellow
        exit 1
    }
}
catch {
    Write-Log "Fatal error: $_" "ERROR"
    Write-Log "Error details: $($_.Exception.Message)" "ERROR"
    Write-Host "`n✗ Script failed with fatal error. Check log for details." -ForegroundColor Red
    exit 1
}
finally {
    Write-Host ""
    Write-Host "Press any key to exit..." -ForegroundColor Gray
    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
}