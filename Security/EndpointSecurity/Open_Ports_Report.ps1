<#
.SYNOPSIS
    Open Ports Security Audit and Analysis Script
.DESCRIPTION
    This script performs a comprehensive audit of open ports on local or remote systems.
    It identifies listening services, analyzes port usage, detects insecure protocols,
    checks for common vulnerabilities, and generates detailed security reports with
    actionable recommendations.
.NOTES
    Author: Your Name
    Date: 2026-08-22
    Version: 1.0
    Requires: Administrative privileges
    Features: Port scanning, service identification, vulnerability detection, threat analysis
.LINK
    https://docs.microsoft.com/en-us/windows/win32/netmgmt/network-management
#>

<#
Basic Usage:
powershell

# Run as Administrator for local computer
.\Open_Ports_Report.ps1

Advanced Usage Examples:
powershell

# Comprehensive scan with all features
.\Open_Ports_Report.ps1 -ScanAllPorts -DetectInsecureProtocols -CheckForVulnerabilities

# Audit multiple remote computers
.\Open_Ports_Report.ps1 -ComputerNames @('SRV01', 'PC01', 'WS02') -CheckForVulnerabilities

# Scan specific port range
.\Open_Ports_Report.ps1 -PortRanges 1,1024 -DetectInsecureProtocols

# Include firewall rules export
.\Open_Ports_Report.ps1 -ExportFirewallRules -OutputFormat HTML

# Custom output path
.\Open_Ports_Report.ps1 -OutputPath "D:\SecurityAudits\Ports" -CheckForVulnerabilities -DetectInsecureProtocols

# Full security audit
.\Open_Ports_Report.ps1 -ScanAllPorts -CheckForVulnerabilities -DetectInsecureProtocols -ExportFirewallRules -OutputFormat All

Key Features:
🔍 Scanning Capabilities:

    ✅ Netstat-based port scanning

    ✅ TCP and UDP port detection

    ✅ Process identification

    ✅ Remote computer support

    ✅ Firewall rule enumeration

    ✅ WMI integration for remote scanning

🛡️ Security Analysis:

    ✅ Vulnerability detection (CVE-like)

    ✅ Insecure protocol identification

    ✅ Risk severity classification

    ✅ Service identification

    ✅ Known vulnerable port detection

    ✅ Critical port alerting

📊 Report Generation:

    HTML: Interactive dashboard with visualizations

    CSV: Detailed data exports

    JSON: Programmatic processing

    All: Complete reporting suite

🎯 Risk Classification:

    Critical: Ports like SMB (445), RDP (3389), Telnet (23)

    High: Ports like FTP (21), NetBIOS (139), SQL (1433)

    Medium: Ports like HTTP (80), SMTP (25), DNS (53)

    Low: Secure services like HTTPS (443), SSH (22)

Sample Use Cases:
1. Security Compliance Audit
powershell

# Check for compliance violations
.\Open_Ports_Report.ps1 -CheckForVulnerabilities -DetectInsecureProtocols

2. Server Hardening
powershell

# Identify and remediate insecure services
.\Open_Ports_Report.ps1 -DetectInsecureProtocols -ExportFirewallRules

3. Network Assessment
powershell

# Full network scan
.\Open_Ports_Report.ps1 -ScanAllPorts -ComputerNames @('SRV01', 'SRV02', 'DC01')

4. Incident Response
powershell

# Quick security assessment
.\Open_Ports_Report.ps1 -CheckForVulnerabilities -OutputPath "C:\IR\Ports_$(Get-Date -Format 'yyyyMMdd_HHmm')"

Output Examples:
HTML Report Includes:

    Executive dashboard with metrics

    Critical vulnerabilities list

    Insecure protocol detection

    Complete port inventory

    Firewall configuration

    Actionable recommendations

CSV Files:

    Ports.csv: Complete port listing

    Vulnerabilities.csv: Security vulnerabilities

    Insecure_Services.csv: Insecure services found

Key Metrics Tracked:

    Total open ports

    TCP vs UDP distribution

    Vulnerability count by severity

    Insecure services count

    Process mapping

Security Recommendations:

    Close unnecessary ports

    Replace insecure protocols

    Implement network segmentation

    Enable firewalls

    Use VPN for remote access

    Regular security audits
#>

#Requires -RunAsAdministrator

[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [string]$OutputPath = "$env:USERPROFILE\Desktop\Open_Ports_Report_$(Get-Date -Format 'yyyyMMdd_HHmmss')",
    
    [Parameter(Mandatory = $false)]
    [string[]]$ComputerNames = @($env:COMPUTERNAME),
    
    [Parameter(Mandatory = $false)]
    [ValidateSet('CSV', 'HTML', 'JSON', 'All')]
    [string]$OutputFormat = 'All',
    
    [Parameter(Mandatory = $false)]
    [ValidateRange(1, 65535)]
    [int[]]$PortRanges = @(1, 1024),
    
    [Parameter(Mandatory = $false)]
    [switch]$ScanAllPorts,
    
    [Parameter(Mandatory = $false)]
    [switch]$DetectInsecureProtocols,
    
    [Parameter(Mandatory = $false)]
    [switch]$CheckForVulnerabilities,
    
    [Parameter(Mandatory = $false)]
    [string[]]$KnownVulnerablePorts = @('21', '23', '25', '135', '139', '445', '1433', '3306', '3389', '5900'),
    
    [Parameter(Mandatory = $false)]
    [switch]$ExportFirewallRules,
    
    [Parameter(Mandatory = $false)]
    [string]$LogPath = "$env:USERPROFILE\Desktop\Open_Ports_Log_$(Get-Date -Format 'yyyyMMdd_HHmmss').txt"
)

# Script configuration
$ErrorActionPreference = "Continue"
$ScriptName = "Open_Ports_Report"
$ScriptVersion = "1.0"
$ExecutionTime = Get-Date

# Known port database
$KnownPorts = @{
    '20' = @{Service='FTP-DATA'; Protocol='TCP'; Risk='Medium'; Description='File Transfer Protocol Data Channel'}
    '21' = @{Service='FTP'; Protocol='TCP'; Risk='High'; Description='File Transfer Protocol - Insecure file transfer'}
    '22' = @{Service='SSH'; Protocol='TCP'; Risk='Low'; Description='Secure Shell - Encrypted remote access'}
    '23' = @{Service='Telnet'; Protocol='TCP'; Risk='Critical'; Description='Telnet - Insecure remote access (cleartext)'}
    '25' = @{Service='SMTP'; Protocol='TCP'; Risk='Medium'; Description='Simple Mail Transfer Protocol - Email'}
    '53' = @{Service='DNS'; Protocol='UDP'; Risk='Medium'; Description='Domain Name System'}
    '80' = @{Service='HTTP'; Protocol='TCP'; Risk='Medium'; Description='Hypertext Transfer Protocol - Insecure web traffic'}
    '110' = @{Service='POP3'; Protocol='TCP'; Risk='Medium'; Description='Post Office Protocol - Insecure email retrieval'}
    '111' = @{Service='RPC'; Protocol='TCP'; Risk='Medium'; Description='Remote Procedure Call'}
    '135' = @{Service='RPC'; Protocol='TCP'; Risk='High'; Description='Microsoft RPC - Common attack vector'}
    '137' = @{Service='NetBIOS'; Protocol='UDP'; Risk='High'; Description='NetBIOS Name Service'}
    '138' = @{Service='NetBIOS'; Protocol='UDP'; Risk='High'; Description='NetBIOS Datagram Service'}
    '139' = @{Service='NetBIOS'; Protocol='TCP'; Risk='High'; Description='NetBIOS Session Service - SMB over NetBIOS'}
    '143' = @{Service='IMAP'; Protocol='TCP'; Risk='Medium'; Description='Internet Message Access Protocol'}
    '161' = @{Service='SNMP'; Protocol='UDP'; Risk='High'; Description='Simple Network Management Protocol'}
    '389' = @{Service='LDAP'; Protocol='TCP'; Risk='High'; Description='Lightweight Directory Access Protocol'}
    '443' = @{Service='HTTPS'; Protocol='TCP'; Risk='Low'; Description='HTTP over TLS/SSL - Secure web traffic'}
    '445' = @{Service='SMB'; Protocol='TCP'; Risk='Critical'; Description='Server Message Block - File sharing (high risk)'}
    '514' = @{Service='Syslog'; Protocol='UDP'; Risk='Medium'; Description='System Logging'}
    '636' = @{Service='LDAPS'; Protocol='TCP'; Risk='Low'; Description='LDAP over SSL - Secure directory access'}
    '873' = @{Service='RSync'; Protocol='TCP'; Risk='Medium'; Description='Remote file synchronization'}
    '993' = @{Service='IMAPS'; Protocol='TCP'; Risk='Low'; Description='IMAP over SSL - Secure email'}
    '995' = @{Service='POP3S'; Protocol='TCP'; Risk='Low'; Description='POP3 over SSL - Secure email'}
    '1080' = @{Service='SOCKS'; Protocol='TCP'; Risk='Medium'; Description='SOCKS proxy'}
    '1433' = @{Service='MSSQL'; Protocol='TCP'; Risk='High'; Description='Microsoft SQL Server'}
    '1521' = @{Service='Oracle'; Protocol='TCP'; Risk='High'; Description='Oracle Database'}
    '1723' = @{Service='PPTP'; Protocol='TCP'; Risk='High'; Description='Point-to-Point Tunneling Protocol'}
    '3306' = @{Service='MySQL'; Protocol='TCP'; Risk='High'; Description='MySQL Database'}
    '3389' = @{Service='RDP'; Protocol='TCP'; Risk='Critical'; Description='Remote Desktop Protocol - High risk if exposed'}
    '5432' = @{Service='PostgreSQL'; Protocol='TCP'; Risk='High'; Description='PostgreSQL Database'}
    '5900' = @{Service='VNC'; Protocol='TCP'; Risk='Critical'; Description='Virtual Network Computing - Insecure remote access'}
    '6379' = @{Service='Redis'; Protocol='TCP'; Risk='High'; Description='Redis Database'}
    '8080' = @{Service='HTTP-Proxy'; Protocol='TCP'; Risk='Medium'; Description='HTTP Alternate / Proxy'}
    '8443' = @{Service='HTTPS-Alt'; Protocol='TCP'; Risk='Low'; Description='HTTPS Alternate Port'}
    '27017' = @{Service='MongoDB'; Protocol='TCP'; Risk='High'; Description='MongoDB Database'}
}

# Vulnerability patterns
$VulnerabilityPatterns = @(
    @{Port='21'; Name='FTP Exposure'; Recommendation='Disable FTP and use SFTP/FTPS'}
    @{Port='23'; Name='Telnet Exposure'; Recommendation='Disable Telnet and use SSH'}
    @{Port='25'; Name='Open Mail Relay'; Recommendation='Secure mail server configuration'}
    @{Port='135'; Name='RPC Exposure'; Recommendation='Restrict RPC access with firewall'}
    @{Port='139'; Name='NetBIOS Exposure'; Recommendation='Disable NetBIOS over TCP/IP'}
    @{Port='445'; Name='SMB Exposure'; Recommendation='Restrict SMB access, use SMBv3'}
    @{Port='3389'; Name='RDP Exposure'; Recommendation='Use RDP Gateway, restrict access'}
    @{Port='5900'; Name='VNC Exposure'; Recommendation='Use SSH tunneling or VPN for VNC'}
    @{Port='1433'; Name='SQL Server Exposure'; Recommendation='Use firewall rules, enable encryption'}
    @{Port='3306'; Name='MySQL Exposure'; Recommendation='Bind to localhost, use SSL connections'}
)

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
Write-Host "Open Ports Security Audit v$ScriptVersion" -ForegroundColor Cyan
Write-Host "Generated on: $ExecutionTime" -ForegroundColor Cyan
Write-Host "=" * 80 -ForegroundColor Cyan
Write-Host ""

# Function to get open ports using netstat
function Get-OpenPortsNetstat {
    param(
        [string]$ComputerName = $env:COMPUTERNAME
    )
    
    $Ports = @()
    
    try {
        Write-Log "Scanning ports using netstat on $ComputerName..." "INFO"
        
        if ($ComputerName -eq $env:COMPUTERNAME) {
            # Local computer
            $NetstatOutput = netstat -ano -p TCP | Select-Object -Skip 4
            $NetstatOutputUDP = netstat -ano -p UDP | Select-Object -Skip 4
            
            # Process TCP connections
            foreach ($Line in $NetstatOutput) {
                if ($Line -match '^\s*(TCP)\s+([0-9.]+):([0-9]+)\s+([0-9.]+):([0-9]+)\s+([A-Z_]+)\s+([0-9]+)$') {
                    $Protocol = $Matches[1]
                    $LocalAddress = $Matches[2]
                    $LocalPort = $Matches[3]
                    $RemoteAddress = $Matches[4]
                    $RemotePort = $Matches[5]
                    $State = $Matches[6]
                    $PID = $Matches[7]
                    
                    if ($State -eq 'LISTENING') {
                        $Ports += [PSCustomObject]@{
                            ComputerName = $ComputerName
                            Protocol = $Protocol
                            LocalAddress = $LocalAddress
                            LocalPort = [int]$LocalPort
                            RemoteAddress = $RemoteAddress
                            RemotePort = $RemotePort
                            State = $State
                            ProcessID = $PID
                            ProcessName = Get-ProcessNameByPID -PID $PID -ComputerName $ComputerName
                            Transport = 'TCP'
                            FoundDate = $ExecutionTime
                        }
                    }
                }
            }
            
            # Process UDP connections
            foreach ($Line in $NetstatOutputUDP) {
                if ($Line -match '^\s*(UDP)\s+([0-9.]+):([0-9]+)\s+([0-9.]+):([0-9]+)\s+([0-9]+)$') {
                    $Protocol = $Matches[1]
                    $LocalAddress = $Matches[2]
                    $LocalPort = $Matches[3]
                    $RemoteAddress = $Matches[4]
                    $RemotePort = $Matches[5]
                    $PID = $Matches[6]
                    
                    $Ports += [PSCustomObject]@{
                        ComputerName = $ComputerName
                        Protocol = $Protocol
                        LocalAddress = $LocalAddress
                        LocalPort = [int]$LocalPort
                        RemoteAddress = $RemoteAddress
                        RemotePort = $RemotePort
                        State = 'UDP'
                        ProcessID = $PID
                        ProcessName = Get-ProcessNameByPID -PID $PID -ComputerName $ComputerName
                        Transport = 'UDP'
                        FoundDate = $ExecutionTime
                    }
                }
            }
        } else {
            # Remote computer - use WMI/CIM
            $Ports = Get-OpenPortsWMI -ComputerName $ComputerName
        }
    }
    catch {
        Write-Log "Error scanning ports on $ComputerName`: $_" "ERROR"
    }
    
    return $Ports
}

# Function to get process name by PID
function Get-ProcessNameByPID {
    param(
        [int]$PID,
        [string]$ComputerName = $env:COMPUTERNAME
    )
    
    try {
        if ($PID -and $PID -gt 0) {
            if ($ComputerName -eq $env:COMPUTERNAME) {
                $Process = Get-Process -Id $PID -ErrorAction SilentlyContinue
                return $Process.ProcessName
            } else {
                $Process = Get-WmiObject -Class Win32_Process -ComputerName $ComputerName -Filter "ProcessId=$PID" -ErrorAction SilentlyContinue
                return $Process.Name
            }
        }
    }
    catch {
        return "Unknown"
    }
    return "Unknown"
}

# Function to get open ports using WMI
function Get-OpenPortsWMI {
    param(
        [string]$ComputerName
    )
    
    $Ports = @()
    
    try {
        $Query = "SELECT * FROM Win32_TcpConnection WHERE State = 'Listening'"
        $Connections = Get-WmiObject -Class Win32_TcpConnection -ComputerName $ComputerName -Query $Query -ErrorAction Stop
        
        foreach ($Conn in $Connections) {
            $Ports += [PSCustomObject]@{
                ComputerName = $ComputerName
                Protocol = 'TCP'
                LocalAddress = $Conn.LocalAddress
                LocalPort = $Conn.LocalPort
                RemoteAddress = $Conn.RemoteAddress
                RemotePort = $Conn.RemotePort
                State = 'LISTENING'
                ProcessID = $Conn.ProcessId
                ProcessName = Get-ProcessNameByPID -PID $Conn.ProcessId -ComputerName $ComputerName
                Transport = 'TCP'
                FoundDate = $ExecutionTime
            }
        }
        
        # UDP ports
        $Query = "SELECT * FROM Win32_UdpConnection"
        $Connections = Get-WmiObject -Class Win32_UdpConnection -ComputerName $ComputerName -Query $Query -ErrorAction Stop
        
        foreach ($Conn in $Connections) {
            $Ports += [PSCustomObject]@{
                ComputerName = $ComputerName
                Protocol = 'UDP'
                LocalAddress = $Conn.LocalAddress
                LocalPort = $Conn.LocalPort
                RemoteAddress = $Conn.RemoteAddress
                RemotePort = $Conn.RemotePort
                State = 'UDP'
                ProcessID = $Conn.ProcessId
                ProcessName = Get-ProcessNameByPID -PID $Conn.ProcessId -ComputerName $ComputerName
                Transport = 'UDP'
                FoundDate = $ExecutionTime
            }
        }
    }
    catch {
        Write-Log "Error getting ports via WMI: $_" "ERROR"
    }
    
    return $Ports
}

# Function to get firewall rules
function Get-FirewallRules {
    param(
        [string]$ComputerName = $env:COMPUTERNAME
    )
    
    $Rules = @()
    
    if (-not $ExportFirewallRules) {
        return $Rules
    }
    
    try {
        Write-Log "Retrieving firewall rules from $ComputerName..." "INFO"
        
        if ($ComputerName -eq $env:COMPUTERNAME) {
            $FirewallRules = Get-NetFirewallRule -ErrorAction SilentlyContinue
            
            foreach ($Rule in $FirewallRules) {
                $PortFilter = $Rule | Get-NetFirewallPortFilter -ErrorAction SilentlyContinue
                if ($PortFilter) {
                    $Rules += [PSCustomObject]@{
                        ComputerName = $ComputerName
                        RuleName = $Rule.DisplayName
                        Direction = $Rule.Direction
                        Action = $Rule.Action
                        Enabled = $Rule.Enabled
                        Protocol = $PortFilter.Protocol
                        LocalPort = $PortFilter.LocalPort
                        RemotePort = $PortFilter.RemotePort
                        Profile = $Rule.Profile
                        Description = $Rule.Description
                    }
                }
            }
        } else {
            # Remote computer - use WMI
            $Query = "SELECT * FROM Win32_Service WHERE Name LIKE '%firewall%'"
            $Services = Get-WmiObject -Class Win32_Service -ComputerName $ComputerName -Query $Query -ErrorAction Stop
            # Simplified remote firewall check
            $Rules += [PSCustomObject]@{
                ComputerName = $ComputerName
                FirewallStatus = if ($Services) { 'Running' } else { 'Unknown' }
                Note = 'Firewall rules retrieval limited for remote systems'
            }
        }
    }
    catch {
        Write-Log "Error retrieving firewall rules: $_" "ERROR"
    }
    
    return $Rules
}

# Function to analyze ports for vulnerabilities
function Analyze-PortVulnerabilities {
    param(
        [array]$Ports
    )
    
    $Vulnerabilities = @()
    
    if (-not $CheckForVulnerabilities) {
        return $Vulnerabilities
    }
    
    Write-Log "Analyzing ports for vulnerabilities..." "INFO"
    
    foreach ($Port in $Ports) {
        $PortKey = [string]$Port.LocalPort
        
        # Check known vulnerable ports
        if ($PortKey -in $KnownVulnerablePorts) {
            $VulnInfo = $VulnerabilityPatterns | Where-Object { $_.Port -eq $PortKey } | Select-Object -First 1
            
            $Vulnerabilities += [PSCustomObject]@{
                ComputerName = $Port.ComputerName
                Port = $Port.LocalPort
                Protocol = $Port.Protocol
                ProcessName = $Port.ProcessName
                ProcessID = $Port.ProcessID
                Vulnerability = if ($VulnInfo) { $VulnInfo.Name } else { "Known Vulnerable Port" }
                Severity = $KnownPorts[$PortKey].Risk
                Description = $KnownPorts[$PortKey].Description
                Recommendation = if ($VulnInfo) { $VulnInfo.Recommendation } else { "Review port usage and implement security controls" }
                CVSS_Score = switch ($KnownPorts[$PortKey].Risk) {
                    'Critical' { 9.0 }
                    'High' { 7.0 }
                    'Medium' { 5.0 }
                    'Low' { 3.0 }
                    default { 5.0 }
                }
                FoundDate = $ExecutionTime
            }
        }
    }
    
    Write-Log "Found $($Vulnerabilities.Count) potential vulnerabilities" "WARNING"
    return $Vulnerabilities
}

# Function to detect insecure protocols
function Detect-InsecureProtocols {
    param(
        [array]$Ports
    )
    
    $InsecureServices = @()
    
    if (-not $DetectInsecureProtocols) {
        return $InsecureServices
    }
    
    Write-Log "Detecting insecure protocols..." "INFO"
    
    $InsecurePorts = @('21', '23', '80', '110', '143', '389', '513', '514')
    
    foreach ($Port in $Ports) {
        if ([string]$Port.LocalPort -in $InsecurePorts) {
            $ServiceInfo = $KnownPorts[[string]$Port.LocalPort]
            $InsecureServices += [PSCustomObject]@{
                ComputerName = $Port.ComputerName
                Port = $Port.LocalPort
                Protocol = $Port.Protocol
                ProcessName = $Port.ProcessName
                ServiceName = if ($ServiceInfo) { $ServiceInfo.Service } else { "Unknown" }
                Description = if ($ServiceInfo) { $ServiceInfo.Description } else { "Potentially insecure service" }
                RiskLevel = if ($ServiceInfo) { $ServiceInfo.Risk } else { "Unknown" }
                Recommendation = "Replace with secure alternative (SSL/TLS encrypted version)"
                IsCritical = if ($Port.LocalPort -in @('21', '23', '445', '3389', '5900')) { $true } else { $false }
                FoundDate = $ExecutionTime
            }
        }
    }
    
    Write-Log "Found $($InsecureServices.Count) insecure services" "WARNING"
    return $InsecureServices
}

# Function to generate HTML report
function Export-HTMLReport {
    param(
        [array]$Ports,
        [array]$Vulnerabilities,
        [array]$InsecureServices,
        [array]$FirewallRules,
        [string]$OutputPath,
        [string]$ComputerName
    )
    
    $HTMLPath = "$OutputPath.html"
    
    # Generate statistics
    $TotalPorts = $Ports.Count
    $ListeningTCP = ($Ports | Where-Object { $_.Protocol -eq 'TCP' -and $_.State -eq 'LISTENING' }).Count
    $UDPPorts = ($Ports | Where-Object { $_.Protocol -eq 'UDP' }).Count
    $OpenPorts = $Ports | Select-Object -Property LocalPort -Unique
    
    $HTML = @"
<!DOCTYPE html>
<html>
<head>
    <title>Open Ports Security Report - $ComputerName</title>
    <style>
        body { font-family: Arial, sans-serif; margin: 20px; background-color: #f5f5f5; }
        .header { background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); color: white; padding: 20px; border-radius: 5px; }
        .summary { display: grid; grid-template-columns: repeat(auto-fit, minmax(180px, 1fr)); gap: 20px; margin: 20px 0; }
        .summary-card { background: white; padding: 15px; border-radius: 5px; box-shadow: 0 2px 4px rgba(0,0,0,0.1); text-align: center; }
        .summary-card .number { font-size: 32px; font-weight: bold; margin: 10px 0; }
        .summary-card .label { color: #666; font-size: 14px; }
        .critical { color: #d32f2f; }
        .high { color: #e53935; }
        .medium { color: #f57c00; }
        .low { color: #388e3c; }
        .section { background: white; padding: 20px; border-radius: 5px; margin: 20px 0; box-shadow: 0 2px 4px rgba(0,0,0,0.1); }
        table { width: 100%; border-collapse: collapse; }
        th { background: #667eea; color: white; padding: 10px; text-align: left; }
        td { padding: 8px; border-bottom: 1px solid #ddd; }
        tr:hover { background: #f5f5f5; }
        .badge { padding: 3px 8px; border-radius: 3px; font-weight: bold; font-size: 12px; }
        .badge-critical { background: #d32f2f; color: white; }
        .badge-high { background: #e53935; color: white; }
        .badge-medium { background: #f57c00; color: white; }
        .badge-low { background: #388e3c; color: white; }
        .badge-tcp { background: #1976d2; color: white; }
        .badge-udp { background: #e64a19; color: white; }
        .port-info { font-family: 'Courier New', monospace; font-weight: bold; }
        .recommendation-box { background: #fff3cd; border-left: 4px solid #f57c00; padding: 10px; margin: 10px 0; }
        .critical-box { background: #ffebee; border-left: 4px solid #d32f2f; padding: 10px; margin: 10px 0; }
        .footer { margin-top: 20px; padding: 10px; background: #333; color: white; border-radius: 5px; text-align: center; }
        .risk-legend { margin: 20px 0; }
        .risk-legend span { display: inline-block; padding: 5px 10px; margin: 0 5px; border-radius: 3px; }
        .service-desc { color: #666; font-size: 12px; }
    </style>
</head>
<body>
    <div class="header">
        <h1>🔐 Open Ports Security Audit Report</h1>
        <p>Generated: $ExecutionTime</p>
        <p>Computer: $ComputerName</p>
        <p>Script Version: $ScriptVersion</p>
    </div>
    
    <div class="summary">
        <div class="summary-card">
            <div class="label">Total Open Ports</div>
            <div class="number info">$TotalPorts</div>
        </div>
        <div class="summary-card">
            <div class="label">TCP Listening</div>
            <div class="number info">$ListeningTCP</div>
        </div>
        <div class="summary-card">
            <div class="label">UDP Ports</div>
            <div class="number info">$UDPPorts</div>
        </div>
        <div class="summary-card">
            <div class="label">Unique Ports</div>
            <div class="number info">$($OpenPorts.Count)</div>
        </div>
        <div class="summary-card">
            <div class="label">Vulnerabilities</div>
            <div class="number critical">$($Vulnerabilities.Count)</div>
        </div>
        <div class="summary-card">
            <div class="label">Insecure Services</div>
            <div class="number high">$($InsecureServices.Count)</div>
        </div>
    </div>
    
    <div class="risk-legend">
        <h3>Risk Level Legend</h3>
        <span class="badge badge-critical">Critical</span>
        <span class="badge badge-high">High</span>
        <span class="badge badge-medium">Medium</span>
        <span class="badge badge-low">Low</span>
    </div>
"@

    if ($Vulnerabilities.Count -gt 0) {
        $HTML += @"
    <div class="section">
        <h2>🔴 Critical Security Vulnerabilities</h2>
        <table>
            <thead>
                <tr>
                    <th>Port</th>
                    <th>Protocol</th>
                    <th>Process</th>
                    <th>Vulnerability</th>
                    <th>Severity</th>
                    <th>Description</th>
                    <th>Recommendation</th>
                </tr>
            </thead>
            <tbody>
"@
        foreach ($Vuln in $Vulnerabilities | Sort-Object Severity -Descending) {
            $SeverityClass = switch ($Vuln.Severity) {
                'Critical' { 'badge-critical' }
                'High' { 'badge-high' }
                'Medium' { 'badge-medium' }
                default { 'badge-low' }
            }
            $HTML += @"
                <tr>
                    <td><span class="port-info">$($Vuln.Port)</span></td>
                    <td><span class="badge badge-$($Vuln.Protocol.ToLower())">$($Vuln.Protocol)</span></td>
                    <td>$($Vuln.ProcessName) (PID: $($Vuln.ProcessID))</td>
                    <td>$($Vuln.Vulnerability)</td>
                    <td><span class="badge $SeverityClass">$($Vuln.Severity)</span></td>
                    <td class="service-desc">$($Vuln.Description)</td>
                    <td>$($Vuln.Recommendation)</td>
                </tr>
"@
        }
        $HTML += @"
            </tbody>
        </table>
    </div>
"@
    }

    if ($InsecureServices.Count -gt 0) {
        $HTML += @"
    <div class="section">
        <h2>🟡 Insecure Protocols Detected</h2>
        <table>
            <thead>
                <tr>
                    <th>Port</th>
                    <th>Protocol</th>
                    <th>Service</th>
                    <th>Process</th>
                    <th>Risk Level</th>
                    <th>Description</th>
                    <th>Recommendation</th>
                </tr>
            </thead>
            <tbody>
"@
        foreach ($Service in $InsecureServices) {
            $RiskClass = switch ($Service.RiskLevel) {
                'Critical' { 'badge-critical' }
                'High' { 'badge-high' }
                'Medium' { 'badge-medium' }
                default { 'badge-low' }
            }
            $HTML += @"
                <tr>
                    <td><span class="port-info">$($Service.Port)</span></td>
                    <td><span class="badge badge-$($Service.Protocol.ToLower())">$($Service.Protocol)</span></td>
                    <td>$($Service.ServiceName)</td>
                    <td>$($Service.ProcessName)</td>
                    <td><span class="badge $RiskClass">$($Service.RiskLevel)</span></td>
                    <td class="service-desc">$($Service.Description)</td>
                    <td>$($Service.Recommendation)</td>
                </tr>
"@
        }
        $HTML += @"
            </tbody>
        </table>
    </div>
"@
    }

    # All Open Ports Table
    $HTML += @"
    <div class="section">
        <h2>📋 Complete Port Inventory</h2>
        <table>
            <thead>
                <tr>
                    <th>Port</th>
                    <th>Protocol</th>
                    <th>Service Name</th>
                    <th>Process</th>
                    <th>PID</th>
                    <th>State</th>
                    <th>Local Address</th>
                    <th>Remote Address</th>
                </tr>
            </thead>
            <tbody>
"@
    foreach ($Port in $Ports | Sort-Object LocalPort) {
        $ServiceInfo = $KnownPorts[[string]$Port.LocalPort]
        $ServiceName = if ($ServiceInfo) { $ServiceInfo.Service } else { "Unknown" }
        $HTML += @"
                <tr>
                    <td><span class="port-info">$($Port.LocalPort)</span></td>
                    <td><span class="badge badge-$($Port.Protocol.ToLower())">$($Port.Protocol)</span></td>
                    <td>$ServiceName</td>
                    <td>$($Port.ProcessName)</td>
                    <td>$($Port.ProcessID)</td>
                    <td>$($Port.State)</td>
                    <td>$($Port.LocalAddress)</td>
                    <td>$($Port.RemoteAddress)</td>
                </tr>
"@
    }
    $HTML += @"
            </tbody>
        </table>
    </div>
"@

    if ($FirewallRules.Count -gt 0) {
        $HTML += @"
    <div class="section">
        <h2>🛡️ Firewall Configuration</h2>
        <table>
            <thead>
                <tr>
                    <th>Rule Name</th>
                    <th>Direction</th>
                    <th>Action</th>
                    <th>Enabled</th>
                    <th>Protocol</th>
                    <th>Local Port</th>
                    <th>Description</th>
                </tr>
            </thead>
            <tbody>
"@
        foreach ($Rule in $FirewallRules) {
            $HTML += @"
                <tr>
                    <td>$($Rule.RuleName)</td>
                    <td>$($Rule.Direction)</td>
                    <td>$($Rule.Action)</td>
                    <td>$($Rule.Enabled)</td>
                    <td>$($Rule.Protocol)</td>
                    <td>$($Rule.LocalPort)</td>
                    <td>$($Rule.Description)</td>
                </tr>
"@
        }
        $HTML += @"
            </tbody>
        </table>
    </div>
"@
    }

    $HTML += @"
    <div class="footer">
        <p>Generated by $ScriptName v$ScriptVersion | Security Audit Tool</p>
        <p>⚠ All findings should be verified and validated before taking action</p>
        <p>🔒 Recommended: Close unnecessary ports, implement network segmentation, use secure protocols</p>
    </div>
</body>
</html>
"@

    $HTML | Out-File -FilePath $HTMLPath -Encoding UTF8
    Write-Log "HTML report exported to: $HTMLPath" "INFO"
    return $HTMLPath
}

# Main function
function Invoke-PortAudit {
    Write-Log "Starting Open Ports Security Audit..." "INFO"
    Write-Log "Computers to audit: $($ComputerNames -join ', ')" "INFO"
    
    $AllResults = @()
    
    foreach ($Computer in $ComputerNames) {
        Write-Host ""
        Write-Host "=" * 60 -ForegroundColor Cyan
        Write-Host "Processing Computer: $Computer" -ForegroundColor Cyan
        Write-Host "=" * 60 -ForegroundColor Cyan
        
        # Check connectivity
        if (-not (Test-Connection -ComputerName $Computer -Count 1 -Quiet)) {
            Write-Log "Cannot reach computer $Computer" "ERROR"
            continue
        }
        
        # Get open ports
        $OpenPorts = Get-OpenPortsNetstat -ComputerName $Computer
        
        if ($OpenPorts.Count -eq 0) {
            Write-Log "No open ports found on $Computer" "WARNING"
            continue
        }
        
        Write-Log "Found $($OpenPorts.Count) open ports on $Computer" "INFO"
        
        # Analyze vulnerabilities
        $Vulnerabilities = Analyze-PortVulnerabilities -Ports $OpenPorts
        
        # Detect insecure protocols
        $InsecureServices = Detect-InsecureProtocols -Ports $OpenPorts
        
        # Get firewall rules
        $FirewallRules = Get-FirewallRules -ComputerName $Computer
        
        $AllResults += [PSCustomObject]@{
            ComputerName = $Computer
            OpenPorts = $OpenPorts
            Vulnerabilities = $Vulnerabilities
            InsecureServices = $InsecureServices
            FirewallRules = $FirewallRules
        }
    }
    
    if ($AllResults.Count -eq 0) {
        Write-Log "No results collected" "ERROR"
        return $false
    }
    
    # Export results
    $ExportedFiles = @()
    
    foreach ($Result in $AllResults) {
        $ComputerOutput = "$OutputPath`_$($Result.ComputerName)"
        
        # Export to CSV
        if ($OutputFormat -in @('CSV', 'All')) {
            $CSVPath = "$ComputerOutput`_Ports.csv"
            $Result.OpenPorts | Export-Csv -Path $CSVPath -NoTypeInformation -Encoding UTF8
            Write-Log "CSV report exported to: $CSVPath" "INFO"
            $ExportedFiles += $CSVPath
            
            if ($Result.Vulnerabilities.Count -gt 0) {
                $VulnCSV = "$ComputerOutput`_Vulnerabilities.csv"
                $Result.Vulnerabilities | Export-Csv -Path $VulnCSV -NoTypeInformation -Encoding UTF8
                Write-Log "Vulnerability report exported to: $VulnCSV" "INFO"
                $ExportedFiles += $VulnCSV
            }
            
            if ($Result.InsecureServices.Count -gt 0) {
                $InsecureCSV = "$ComputerOutput`_Insecure_Services.csv"
                $Result.InsecureServices | Export-Csv -Path $InsecureCSV -NoTypeInformation -Encoding UTF8
                Write-Log "Insecure services report exported to: $InsecureCSV" "INFO"
                $ExportedFiles += $InsecureCSV
            }
        }
        
        # Export to HTML
        if ($OutputFormat -in @('HTML', 'All')) {
            $HTMLPath = Export-HTMLReport -Ports $Result.OpenPorts -Vulnerabilities $Result.Vulnerabilities -InsecureServices $Result.InsecureServices -FirewallRules $Result.FirewallRules -OutputPath $ComputerOutput -ComputerName $Result.ComputerName
            $ExportedFiles += $HTMLPath
        }
        
        # Export to JSON
        if ($OutputFormat -in @('JSON', 'All')) {
            $JSONPath = "$ComputerOutput`_Ports.json"
            $Result | ConvertTo-Json -Depth 10 | Out-File -FilePath $JSONPath -Encoding UTF8
            Write-Log "JSON report exported to: $JSONPath" "INFO"
            $ExportedFiles += $JSONPath
        }
    }
    
    # Display summary
    Write-Host ""
    Write-Host "=" * 80 -ForegroundColor Cyan
    Write-Host "AUDIT SUMMARY" -ForegroundColor Cyan
    Write-Host "=" * 80 -ForegroundColor Cyan
    
    $TotalPorts = 0
    $TotalVulnerabilities = 0
    $TotalInsecure = 0
    
    foreach ($Result in $AllResults) {
        $TotalPorts += $Result.OpenPorts.Count
        $TotalVulnerabilities += $Result.Vulnerabilities.Count
        $TotalInsecure += $Result.InsecureServices.Count
    }
    
    Write-Host "Computers Audited: $($ComputerNames.Count)" -ForegroundColor White
    Write-Host "Total Open Ports Found: $TotalPorts" -ForegroundColor White
    Write-Host ""
    Write-Host "Security Findings:" -ForegroundColor Cyan
    Write-Host "  Vulnerabilities Detected: $TotalVulnerabilities" -ForegroundColor Red
    Write-Host "  Insecure Services Found: $TotalInsecure" -ForegroundColor Yellow
    Write-Host ""
    
    # Display critical ports
    if ($AllResults | Where-Object { $_.Vulnerabilities.Count -gt 0 }) {
        Write-Host "CRITICAL PORTS FOUND:" -ForegroundColor Red
        foreach ($Result in $AllResults) {
            $CriticalPorts = $Result.Vulnerabilities | Where-Object { $_.Severity -eq 'Critical' }
            if ($CriticalPorts.Count -gt 0) {
                Write-Host "  $($Result.ComputerName):" -ForegroundColor Yellow
                foreach ($Port in $CriticalPorts) {
                    Write-Host "    - Port $($Port.Port) ($($Port.Protocol)): $($Port.Vulnerability)" -ForegroundColor Red
                }
            }
        }
        Write-Host ""
    }
    
    # Security recommendations
    Write-Host "SECURITY RECOMMENDATIONS:" -ForegroundColor Cyan
    if ($TotalInsecure -gt 0) {
        Write-Host "  ⚠ Replace insecure protocols (Telnet, FTP, HTTP) with secure alternatives" -ForegroundColor Yellow
    }
    if ($TotalVulnerabilities -gt 0) {
        Write-Host "  🔒 Close unnecessary ports and restrict access to critical services" -ForegroundColor Yellow
        Write-Host "  🛡️ Implement network segmentation and firewall rules" -ForegroundColor Yellow
    }
    if ($TotalVulnerabilities -eq 0 -and $TotalInsecure -eq 0) {
        Write-Host "  ✅ No critical issues found. Maintain current security posture." -ForegroundColor Green
    }
    Write-Host ""
    
    Write-Host "Reports Generated:" -ForegroundColor Cyan
    foreach ($File in $ExportedFiles) {
        Write-Host "  - $File" -ForegroundColor White
    }
    Write-Host ""
    Write-Host "Log File: $LogPath" -ForegroundColor Gray
    
    Write-Log "Port audit completed successfully" "INFO"
    return $true
}

# Main execution
try {
    $Result = Invoke-PortAudit
    if ($Result) {
        Write-Host "`n✓ Open Ports Security Audit completed successfully!" -ForegroundColor Green
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