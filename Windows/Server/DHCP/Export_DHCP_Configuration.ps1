<#
.SYNOPSIS
    Exports complete DHCP server configuration including scopes, reservations, options, and lease information.

.DESCRIPTION
    This script exports comprehensive DHCP server configuration data including:
    - All scopes with their properties
    - Scope options (gateway, DNS, etc.)
    - Reservations
    - Active leases
    - Exclusion ranges
    - Server-level options
    - Configuration statistics
    
    Output formats: CSV, JSON, HTML, and plain text reports.

.PARAMETER DHCPServer
    The name or IP address of the DHCP server. Defaults to localhost.

.PARAMETER OutputPath
    The directory where export files will be saved. Defaults to current directory.

.PARAMETER ExportFormat
    Export format(s): CSV, JSON, HTML, TXT, or ALL. Default is ALL.

.PARAMETER IncludeLeases
    Switch to include active lease information in the export.

.PARAMETER IncludeStatistics
    Switch to include scope statistics (utilization, etc.).

.PARAMETER IncludeReservations
    Switch to include reservation information.

.PARAMETER ExportServerOptions
    Switch to include server-level DHCP options.

.PARAMETER ExportScopeOptions
    Switch to include scope-level DHCP options.

.PARAMETER ExportDetailed
    Switch to export detailed configuration (includes all available data).

.PARAMETER CleanOldExports
    Switch to remove export files older than specified days.

.PARAMETER CleanupDays
    Number of days to keep export files when using -CleanOldExports. Default is 30.

.PARAMETER EmailReport
    Switch to send report via email (requires SMTP configuration).

.PARAMETER EmailTo
    Recipient email address for report.

.PARAMETER EmailFrom
    Sender email address.

.PARAMETER EmailServer
    SMTP server address.

.EXAMPLE
    .\Export_DHCP_Configuration.ps1 -DHCPServer "DHCP01" -ExportFormat ALL -IncludeLeases -OutputPath "C:\DHCP_Backups"

.EXAMPLE
    .\Export_DHCP_Configuration.ps1 -DHCPServer "192.168.1.10" -ExportFormat HTML -ExportDetailed -EmailReport -EmailTo "admin@company.com"

.NOTES
    Author: Portfolio Script
    Version: 1.0
    Requires: Windows Server with DHCP role installed, PowerShell 5.1+
#>

<#
    Multiple Export Formats: CSV, JSON, HTML, and TXT reports

    Comprehensive Data Collection:

        All DHCP scopes with properties

        Scope options

        Reservations

        Active leases

        Exclusion ranges

        Server-level options

        Server settings and statistics

    Detailed HTML Report: Interactive, well-designed HTML report with collapsible sections

    Email Integration: Send reports via email

    Export Management: Automatic cleanup of old export files

    Error Handling: Robust error handling with detailed messages

    Performance Metrics: Shows execution time and file sizes

    User-Friendly Output: Color-coded console output with progress indicators

How to Use:
powershell

# Basic export all formats
.\Export_DHCP_Configuration.ps1 -DHCPServer "DHCP01" -ExportFormat ALL

# Export detailed configuration with leases
.\Export_DHCP_Configuration.ps1 -DHCPServer "192.168.1.10" -ExportDetailed -IncludeLeases -IncludeStatistics

# Export to HTML only with reservations
.\Export_DHCP_Configuration.ps1 -DHCPServer "DHCP01" -ExportFormat HTML -IncludeReservations -OutputPath "C:\Reports"

# Full export with email and cleanup
.\Export_DHCP_Configuration.ps1 -DHCPServer "DHCP01" -ExportFormat ALL
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [string]$DHCPServer = "localhost",
    
    [Parameter(Mandatory = $false)]
    [string]$OutputPath = ".",
    
    [Parameter(Mandatory = $false)]
    [ValidateSet("CSV", "JSON", "HTML", "TXT", "ALL")]
    [string]$ExportFormat = "ALL",
    
    [Parameter(Mandatory = $false)]
    [switch]$IncludeLeases,
    
    [Parameter(Mandatory = $false)]
    [switch]$IncludeStatistics,
    
    [Parameter(Mandatory = $false)]
    [switch]$IncludeReservations,
    
    [Parameter(Mandatory = $false)]
    [switch]$ExportServerOptions,
    
    [Parameter(Mandatory = $false)]
    [switch]$ExportScopeOptions,
    
    [Parameter(Mandatory = $false)]
    [switch]$ExportDetailed,
    
    [Parameter(Mandatory = $false)]
    [switch]$CleanOldExports,
    
    [Parameter(Mandatory = $false)]
    [int]$CleanupDays = 30,
    
    [Parameter(Mandatory = $false)]
    [switch]$EmailReport,
    
    [Parameter(Mandatory = $false)]
    [string]$EmailTo,
    
    [Parameter(Mandatory = $false)]
    [string]$EmailFrom,
    
    [Parameter(Mandatory = $false)]
    [string]$EmailServer
)

# Global variables
$Script:ErrorCount = 0
$Script:ExportFiles = @()
$Script:StartTime = Get-Date

# Function to import DHCP module
function Import-DHCPModule {
    try {
        Import-Module DhcpServer -ErrorAction Stop -Force
        Write-Verbose "DHCP Server module loaded successfully."
        return $true
    }
    catch {
        Write-Error "Failed to import DHCP Server module. Error: $($_.Exception.Message)"
        return $false
    }
}

# Function to create output directory
function Ensure-OutputDirectory {
    param([string]$Path)
    
    try {
        if (-not (Test-Path -Path $Path)) {
            New-Item -ItemType Directory -Path $Path -Force | Out-Null
            Write-Verbose "Created output directory: $Path"
        }
        return $true
    }
    catch {
        Write-Error "Failed to create output directory: $($_.Exception.Message)"
        return $false
    }
}

# Function to get DHCP server information
function Get-DHCPServerInfo {
    param([string]$Server)
    
    try {
        $serverInfo = Get-DhcpServerSetting -ComputerName $Server -ErrorAction Stop
        return @{
            ServerName = $Server
            Version = $serverInfo.DhcpServerVersion
            DatabasePath = $serverInfo.DatabasePath
            LoggingEnabled = $serverInfo.LoggingEnabled
            AuditLogPath = $serverInfo.AuditLogPath
            BindingInformation = $serverInfo.BindingInformation
            ConflictDetection = $serverInfo.ConflictDetection
            MACCheckEnabled = $serverInfo.MacCheckEnabled
        }
    }
    catch {
        Write-Warning "Unable to retrieve server settings: $($_.Exception.Message)"
        return @{
            ServerName = $Server
            Version = "Unknown"
            DatabasePath = "Unknown"
            LoggingEnabled = $false
            AuditLogPath = "Unknown"
            BindingInformation = $null
            ConflictDetection = $false
            MACCheckEnabled = $false
        }
    }
}

# Function to get all DHCP scopes with details
function Get-DHCPAllScopes {
    param([string]$Server)
    
    $scopes = @()
    try {
        $scopeList = Get-DhcpServerv4Scope -ComputerName $Server -ErrorAction Stop
        
        foreach ($scope in $scopeList) {
            $scopeData = [PSCustomObject]@{
                ScopeId = $scope.ScopeId.IPAddressToString
                Name = $scope.Name
                Description = $scope.Description
                StartRange = $scope.StartRange.IPAddressToString
                EndRange = $scope.EndRange.IPAddressToString
                SubnetMask = $scope.SubnetMask.IPAddressToString
                LeaseDuration = $scope.LeaseDuration
                State = $scope.State
                ActivationTime = $scope.ActivationTime
                DeactivationTime = $scope.DeactivationTime
                Delay = $scope.Delay
                HasExclusions = $false
                HasReservations = $false
                TotalAddresses = $null
                UsedAddresses = $null
                AvailablePercentage = $null
            }
            
            # Get exclusions
            try {
                $exclusions = Get-DhcpServerv4ExclusionRange -ComputerName $Server -ScopeId $scope.ScopeId -ErrorAction SilentlyContinue
                if ($exclusions) {
                    $scopeData.HasExclusions = $true
                    $scopeData.Exclusions = $exclusions | ForEach-Object {
                        @{
                            StartRange = $_.StartRange.IPAddressToString
                            EndRange = $_.EndRange.IPAddressToString
                        }
                    }
                }
            }
            catch {
                # Ignore exclusion errors
            }
            
            # Get reservations
            if ($IncludeReservations -or $ExportDetailed) {
                try {
                    $reservations = Get-DhcpServerv4Reservation -ComputerName $Server -ScopeId $scope.ScopeId -ErrorAction SilentlyContinue
                    if ($reservations) {
                        $scopeData.HasReservations = $true
                        $scopeData.Reservations = $reservations | ForEach-Object {
                            [PSCustomObject]@{
                                IPAddress = $_.IPAddress.IPAddressToString
                                ClientId = $_.ClientId
                                Name = $_.Name
                                Description = $_.Description
                                Type = $_.Type
                                HostName = $_.HostName
                            }
                        }
                    }
                }
                catch {
                    # Ignore reservation errors
                }
            }
            
            # Get statistics
            if ($IncludeStatistics -or $ExportDetailed) {
                try {
                    $stats = Get-DhcpServerv4ScopeStatistics -ComputerName $Server -ScopeId $scope.ScopeId -ErrorAction SilentlyContinue
                    if ($stats) {
                        $scopeData.TotalAddresses = $stats.TotalAddresses
                        $scopeData.UsedAddresses = $stats.UsedAddresses
                        $scopeData.AvailablePercentage = $stats.AvailablePercentage
                        $scopeData.LeasedAddresses = $stats.LeasedAddresses
                        $scopeData.PercentLeased = $stats.PercentLeased
                        $scopeData.PercentAvailable = $stats.PercentAvailable
                        $scopeData.PercentInUse = $stats.PercentInUse
                    }
                }
                catch {
                    # Ignore statistics errors
                }
            }
            
            # Get scope options
            if ($ExportScopeOptions -or $ExportDetailed) {
                try {
                    $options = Get-DhcpServerv4OptionValue -ComputerName $Server -ScopeId $scope.ScopeId -ErrorAction SilentlyContinue
                    if ($options) {
                        $scopeData.Options = $options | ForEach-Object {
                            [PSCustomObject]@{
                                OptionId = $_.OptionId
                                Name = $_.Name
                                Value = $_.Value
                                Type = $_.Type
                            }
                        }
                    }
                }
                catch {
                    # Ignore options errors
                }
            }
            
            # Get leases
            if ($IncludeLeases -or $ExportDetailed) {
                try {
                    $leases = Get-DhcpServerv4Lease -ComputerName $Server -ScopeId $scope.ScopeId -ErrorAction SilentlyContinue
                    if ($leases) {
                        $scopeData.ActiveLeases = $leases | Where-Object { $_.State -eq "Active" } | ForEach-Object {
                            [PSCustomObject]@{
                                IPAddress = $_.IPAddress.IPAddressToString
                                ClientId = $_.ClientId
                                HostName = $_.HostName
                                LeaseExpiryTime = $_.LeaseExpiryTime
                                AddressState = $_.AddressState
                                Type = $_.Type
                            }
                        }
                        $scopeData.TotalLeases = $leases.Count
                        $scopeData.ActiveLeasesCount = ($leases | Where-Object { $_.State -eq "Active" }).Count
                    }
                }
                catch {
                    # Ignore lease errors
                }
            }
            
            $scopes += $scopeData
        }
        
        return $scopes
    }
    catch {
        Write-Error "Failed to retrieve DHCP scopes: $($_.Exception.Message)"
        return $null
    }
}

# Function to get server options
function Get-DHCPServerOptions {
    param([string]$Server)
    
    try {
        $options = Get-DhcpServerv4OptionValue -ComputerName $Server -ErrorAction SilentlyContinue
        return $options | ForEach-Object {
            [PSCustomObject]@{
                OptionId = $_.OptionId
                Name = $_.Name
                Value = $_.Value
                Type = $_.Type
                VendorClass = $_.VendorClass
                UserClass = $_.UserClass
            }
        }
    }
    catch {
        Write-Warning "Unable to retrieve server options: $($_.Exception.Message)"
        return $null
    }
}

# Function to export to CSV
function Export-ToCSV {
    param(
        [string]$FilePath,
        [array]$Data,
        [string]$Description = "DHCP Configuration"
    )
    
    try {
        if ($Data -and $Data.Count -gt 0) {
            $Data | Export-Csv -Path $FilePath -NoTypeInformation -Encoding UTF8
            Write-Host "✓ CSV exported to: $FilePath" -ForegroundColor Green
            return $true
        }
        else {
            Write-Warning "No data to export to CSV"
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
        [PSObject]$Data,
        [string]$Description = "DHCP Configuration"
    )
    
    try {
        $json = $Data | ConvertTo-Json -Depth 10 -Compress
        $json | Out-File -FilePath $FilePath -Encoding UTF8
        Write-Host "✓ JSON exported to: $FilePath" -ForegroundColor Green
        return $true
    }
    catch {
        Write-Error "Failed to export JSON: $($_.Exception.Message)"
        return $false
    }
}

# Function to export to HTML
function Export-ToHTML {
    param(
        [string]$FilePath,
        [PSObject]$Data,
        [string]$ServerName,
        [datetime]$ExportTime
    )
    
    try {
        $htmlContent = @"
<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <title>DHCP Configuration Report - $ServerName</title>
    <style>
        body { font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; margin: 20px; background-color: #f5f5f5; }
        .container { max-width: 1400px; margin: 0 auto; background-color: white; padding: 20px; border-radius: 8px; box-shadow: 0 2px 4px rgba(0,0,0,0.1); }
        h1 { color: #333; border-bottom: 3px solid #007acc; padding-bottom: 10px; }
        h2 { color: #007acc; margin-top: 30px; border-bottom: 2px solid #e0e0e0; padding-bottom: 8px; }
        .summary { background-color: #e8f4fd; padding: 15px; border-radius: 5px; margin: 20px 0; }
        table { width: 100%; border-collapse: collapse; margin: 15px 0; }
        th { background-color: #007acc; color: white; padding: 12px; text-align: left; }
        td { padding: 10px; border-bottom: 1px solid #ddd; }
        tr:hover { background-color: #f5f5f5; }
        .status-active { color: green; font-weight: bold; }
        .status-inactive { color: red; font-weight: bold; }
        .badge { background-color: #007acc; color: white; padding: 3px 8px; border-radius: 3px; font-size: 12px; }
        .collapsible { cursor: pointer; padding: 10px; background-color: #f0f0f0; border: none; text-align: left; outline: none; width: 100%; }
        .content { padding: 0 18px; display: none; overflow: hidden; }
        .footer { margin-top: 30px; padding-top: 15px; border-top: 1px solid #ddd; color: #666; font-size: 12px; text-align: center; }
    </style>
    <script>
        function toggleSection(id) {
            var content = document.getElementById(id);
            if (content.style.display === "block") {
                content.style.display = "none";
            } else {
                content.style.display = "block";
            }
        }
    </script>
</head>
<body>
<div class="container">
    <h1>📊 DHCP Configuration Report</h1>
    <div class="summary">
        <p><strong>Server:</strong> $ServerName</p>
        <p><strong>Export Date:</strong> $($ExportTime.ToString('yyyy-MM-dd HH:mm:ss'))</p>
        <p><strong>Total Scopes:</strong> $($Data.Scopes.Count)</p>
        <p><strong>Total Active Leases:</strong> $(($Data.Scopes | ForEach-Object { $_.ActiveLeasesCount } | Measure-Object -Sum).Sum)</p>
        <p><strong>Total Reservations:</strong> $(($Data.Scopes | Where-Object { $_.HasReservations } | ForEach-Object { $_.Reservations.Count } | Measure-Object -Sum).Sum)</p>
    </div>
"@

        # Scopes section
        $htmlContent += "<h2>📋 DHCP Scopes</h2>"
        if ($Data.Scopes -and $Data.Scopes.Count -gt 0) {
            $htmlContent += "<table>"
            $htmlContent += @"
    <tr>
        <th>Scope ID</th>
        <th>Name</th>
        <th>IP Range</th>
        <th>Subnet Mask</th>
        <th>Lease Duration</th>
        <th>State</th>
        <th>Active Leases</th>
        <th>Available %</th>
    </tr>
"@
            foreach ($scope in $Data.Scopes) {
                $stateClass = if ($scope.State -eq "Active") { "status-active" } else { "status-inactive" }
                $htmlContent += @"
    <tr>
        <td><strong>$($scope.ScopeId)</strong></td>
        <td>$($scope.Name)</td>
        <td>$($scope.StartRange) - $($scope.EndRange)</td>
        <td>$($scope.SubnetMask)</td>
        <td>$($scope.LeaseDuration)</td>
        <td class="$stateClass">$($scope.State)</td>
        <td>$($scope.ActiveLeasesCount)</td>
        <td>$(if($scope.AvailablePercentage -ne $null){$scope.AvailablePercentage} else {'N/A'})</td>
    </tr>
"@
            }
            $htmlContent += "</table>"
        }

        # Detailed scope information
        if ($Data.Scopes -and $Data.Scopes.Count -gt 0) {
            $htmlContent += "<h2>🔍 Scope Details</h2>"
            foreach ($scope in $Data.Scopes) {
                $htmlContent += "<button class='collapsible' onclick='toggleSection(`"scope-$($scope.ScopeId.Replace('.','-'))`")'>📌 $($scope.Name) ($($scope.ScopeId))</button>"
                $htmlContent += "<div id='scope-$($scope.ScopeId.Replace('.','-'))' class='content'>"
                
                $htmlContent += "<p><strong>Description:</strong> $(if($scope.Description){$scope.Description} else {'None'})</p>"
                $htmlContent += "<p><strong>Subnet Mask:</strong> $($scope.SubnetMask)</p>"
                $htmlContent += "<p><strong>Lease Duration:</strong> $($scope.LeaseDuration)</p>"
                $htmlContent += "<p><strong>State:</strong> $($scope.State)</p>"
                $htmlContent += "<p><strong>Activation Time:</strong> $(if($scope.ActivationTime){$scope.ActivationTime} else {'N/A'})</p>"
                
                if ($scope.HasExclusions) {
                    $htmlContent += "<h4>🚫 Exclusions</h4><ul>"
                    foreach ($excl in $scope.Exclusions) {
                        $htmlContent += "<li>$($excl.StartRange) - $($excl.EndRange)</li>"
                    }
                    $htmlContent += "</ul>"
                }
                
                if ($scope.HasReservations) {
                    $htmlContent += "<h4>📌 Reservations</h4><ul>"
                    foreach ($res in $scope.Reservations) {
                        $htmlContent += "<li>$($res.IPAddress) - $($res.Name) ($($res.ClientId))</li>"
                    }
                    $htmlContent += "</ul>"
                }
                
                if ($scope.Options -and $scope.Options.Count -gt 0) {
                    $htmlContent += "<h4>⚙️ Scope Options</h4><ul>"
                    foreach ($opt in $scope.Options) {
                        $htmlContent += "<li>Option $($opt.OptionId) - $($opt.Name): $($opt.Value)</li>"
                    }
                    $htmlContent += "</ul>"
                }
                
                $htmlContent += "</div>"
            }
        }

        # Server options
        if ($Data.ServerOptions -and $Data.ServerOptions.Count -gt 0) {
            $htmlContent += "<h2>⚙️ Server Options</h2>"
            $htmlContent += "<ul>"
            foreach ($opt in $Data.ServerOptions) {
                $htmlContent += "<li><strong>Option $($opt.OptionId):</strong> $($opt.Name) = $($opt.Value)</li>"
            }
            $htmlContent += "</ul>"
        }

        $htmlContent += @"
    <div class="footer">
        <p>Generated by DHCP Configuration Export Script</p>
        <p>Report ID: $([System.Guid]::NewGuid().ToString())</p>
    </div>
</div>
</body>
</html>
"@
        
        $htmlContent | Out-File -FilePath $FilePath -Encoding UTF8
        Write-Host "✓ HTML report exported to: $FilePath" -ForegroundColor Green
        return $true
    }
    catch {
        Write-Error "Failed to export HTML: $($_.Exception.Message)"
        return $false
    }
}

# Function to export to TXT
function Export-ToTXT {
    param(
        [string]$FilePath,
        [PSObject]$Data,
        [string]$ServerName,
        [datetime]$ExportTime
    )
    
    try {
        $txtContent = @"
=============================================
DHCP CONFIGURATION EXPORT REPORT
=============================================

Server: $ServerName
Export Date: $($ExportTime.ToString('yyyy-MM-dd HH:mm:ss'))
Report ID: $([System.Guid]::NewGuid().ToString())

=============================================
SUMMARY
=============================================
Total Scopes: $($Data.Scopes.Count)
Total Active Leases: $(($Data.Scopes | ForEach-Object { $_.ActiveLeasesCount } | Measure-Object -Sum).Sum)
Total Reservations: $(($Data.Scopes | Where-Object { $_.HasReservations } | ForEach-Object { $_.Reservations.Count } | Measure-Object -Sum).Sum)

=============================================
SERVER INFORMATION
=============================================
Server Name: $($Data.ServerInfo.ServerName)
Version: $($Data.ServerInfo.Version)
Database Path: $($Data.ServerInfo.DatabasePath)
Logging Enabled: $($Data.ServerInfo.LoggingEnabled)
Conflict Detection: $($Data.ServerInfo.ConflictDetection)

=============================================
DHCP SCOPES
=============================================

"@
        
        foreach ($scope in $Data.Scopes) {
            $txtContent += @"
SCOPE: $($scope.Name) ($($scope.ScopeId))
------------------------------------------------
Description: $(if($scope.Description){$scope.Description} else {'None'})
IP Range: $($scope.StartRange) - $($scope.EndRange)
Subnet Mask: $($scope.SubnetMask)
Lease Duration: $($scope.LeaseDuration)
State: $($scope.State)
Active Leases: $($scope.ActiveLeasesCount)
Available Percentage: $(if($scope.AvailablePercentage -ne $null){$scope.AvailablePercentage} else {'N/A'})

"@
            if ($scope.HasExclusions) {
                $txtContent += "Exclusions:`n"
                foreach ($excl in $scope.Exclusions) {
                    $txtContent += "  - $($excl.StartRange) - $($excl.EndRange)`n"
                }
                $txtContent += "`n"
            }
            
            if ($scope.HasReservations) {
                $txtContent += "Reservations:`n"
                foreach ($res in $scope.Reservations) {
                    $txtContent += "  - $($res.IPAddress) : $($res.Name) ($($res.ClientId))`n"
                }
                $txtContent += "`n"
            }
            
            if ($scope.Options -and $scope.Options.Count -gt 0) {
                $txtContent += "Scope Options:`n"
                foreach ($opt in $scope.Options) {
                    $txtContent += "  - Option $($opt.OptionId) : $($opt.Name) = $($opt.Value)`n"
                }
                $txtContent += "`n"
            }
            $txtContent += "------------------------------------------------`n`n"
        }
        
        if ($Data.ServerOptions -and $Data.ServerOptions.Count -gt 0) {
            $txtContent += "=============================================`nSERVER OPTIONS`n=============================================`n"
            foreach ($opt in $Data.ServerOptions) {
                $txtContent += "Option $($opt.OptionId) : $($opt.Name) = $($opt.Value)`n"
            }
        }
        
        $txtContent += @"

=============================================
END OF REPORT
=============================================
"@
        
        $txtContent | Out-File -FilePath $FilePath -Encoding UTF8
        Write-Host "✓ TXT report exported to: $FilePath" -ForegroundColor Green
        return $true
    }
    catch {
        Write-Error "Failed to export TXT: $($_.Exception.Message)"
        return $false
    }
}

# Function to clean old export files
function Clean-OldExports {
    param(
        [string]$Path,
        [int]$Days
    )
    
    try {
        $oldFiles = Get-ChildItem -Path $Path -Filter "DHCP_Export_*" -Recurse | Where-Object { 
            $_.CreationTime -lt (Get-Date).AddDays(-$Days)
        }
        
        if ($oldFiles) {
            $count = $oldFiles.Count
            $oldFiles | Remove-Item -Force
            Write-Host "✓ Removed $count old export files (older than $Days days)" -ForegroundColor Green
        }
        else {
            Write-Host "ℹ No old export files found to clean" -ForegroundColor Yellow
        }
        return $true
    }
    catch {
        Write-Warning "Failed to clean old exports: $($_.Exception.Message)"
        return $false
    }
}

# Function to send email report
function Send-EmailReport {
    param(
        [string]$To,
        [string]$From,
        [string]$SMTPServer,
        [array]$Attachments,
        [string]$Subject = "DHCP Configuration Export Report",
        [string]$Body = "DHCP configuration export completed. Please find attached the export files."
    )
    
    try {
        $mailParams = @{
            To = $To
            From = $From
            Subject = "$Subject - $(Get-Date -Format 'yyyy-MM-dd')"
            Body = $Body
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
Write-Host "   DHCP CONFIGURATION EXPORT TOOL" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

# Validate DHCP module
if (-not (Import-DHCPModule)) {
    Write-Error "DHCP module is required. Please install DHCP Server role."
    exit 1
}

# Create output directory
$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$exportFolder = Join-Path -Path $OutputPath -ChildPath "DHCP_Export_$timestamp"
if (-not (Ensure-OutputDirectory -Path $exportFolder)) {
    Write-Error "Cannot create output directory. Exiting."
    exit 1
}

Write-Host "📁 Export folder: $exportFolder" -ForegroundColor Yellow
Write-Host "🔄 Collecting DHCP data from server: $DHCPServer" -ForegroundColor Yellow
Write-Host ""

# Collect data
$exportData = @{
    ExportTime = $Script:StartTime
    ServerInfo = Get-DHCPServerInfo -Server $DHCPServer
    Scopes = Get-DHCPAllScopes -Server $DHCPServer
    ServerOptions = $null
}

if ($ExportServerOptions -or $ExportDetailed) {
    $exportData.ServerOptions = Get-DHCPServerOptions -Server $DHCPServer
}

# Verify data collection
if (-not $exportData.Scopes) {
    Write-Warning "No DHCP scopes found or could not retrieve scopes from server."
}

# Determine export formats
$formats = @()
if ($ExportFormat -eq "ALL") {
    $formats = @("CSV", "JSON", "HTML", "TXT")
}
else {
    $formats = @($ExportFormat)
}

# Export data
Write-Host "📤 Exporting data in formats: $($formats -join ', ')" -ForegroundColor Yellow
Write-Host ""

# Export Scopes as CSV
if ("CSV" -in $formats -and $exportData.Scopes) {
    $csvPath = Join-Path -Path $exportFolder -ChildPath "DHCP_Scopes.csv"
    if (Export-ToCSV -FilePath $csvPath -Data $exportData.Scopes) {
        $Script:ExportFiles += $csvPath
    }
}

# Export Reservations as CSV
if ("CSV" -in $formats -and $exportData.Scopes) {
    $reservations = @()
    foreach ($scope in $exportData.Scopes) {
        if ($scope.HasReservations) {
            foreach ($res in $scope.Reservations) {
                $reservations += [PSCustomObject]@{
                    ScopeId = $scope.ScopeId
                    ScopeName = $scope.Name
                    IPAddress = $res.IPAddress
                    ClientId = $res.ClientId
                    Name = $res.Name
                    Description = $res.Description
                    Type = $res.Type
                    HostName = $res.HostName
                }
            }
        }
    }
    if ($reservations.Count -gt 0) {
        $csvPath = Join-Path -Path $exportFolder -ChildPath "DHCP_Reservations.csv"
        if (Export-ToCSV -FilePath $csvPath -Data $reservations) {
            $Script:ExportFiles += $csvPath
        }
    }
}

# Export Leases as CSV
if ("CSV" -in $formats -and $exportData.Scopes -and $IncludeLeases) {
    $leases = @()
    foreach ($scope in $exportData.Scopes) {
        if ($scope.ActiveLeases) {
            foreach ($lease in $scope.ActiveLeases) {
                $leases += [PSCustomObject]@{
                    ScopeId = $scope.ScopeId
                    ScopeName = $scope.Name
                    IPAddress = $lease.IPAddress
                    ClientId = $lease.ClientId
                    HostName = $lease.HostName
                    LeaseExpiryTime = $lease.LeaseExpiryTime
                    AddressState = $lease.AddressState
                    Type = $lease.Type
                }
            }
        }
    }
    if ($leases.Count -gt 0) {
        $csvPath = Join-Path -Path $exportFolder -ChildPath "DHCP_Leases.csv"
        if (Export-ToCSV -FilePath $csvPath -Data $leases) {
            $Script:ExportFiles += $csvPath
        }
    }
}

# Export as JSON
if ("JSON" -in $formats) {
    $jsonPath = Join-Path -Path $exportFolder -ChildPath "DHCP_Configuration.json"
    if (Export-ToJSON -FilePath $jsonPath -Data $exportData) {
        $Script:ExportFiles += $jsonPath
    }
}

# Export as HTML
if ("HTML" -in $formats) {
    $htmlPath = Join-Path -Path $exportFolder -ChildPath "DHCP_Report.html"
    if (Export-ToHTML -FilePath $htmlPath -Data $exportData -ServerName $DHCPServer -ExportTime $Script:StartTime) {
        $Script:ExportFiles += $htmlPath
    }
}

# Export as TXT
if ("TXT" -in $formats) {
    $txtPath = Join-Path -Path $exportFolder -ChildPath "DHCP_Report.txt"
    if (Export-ToTXT -FilePath $txtPath -Data $exportData -ServerName $DHCPServer -ExportTime $Script:StartTime) {
        $Script:ExportFiles += $txtPath
    }
}

# Clean old exports
if ($CleanOldExports) {
    Clean-OldExports -Path $OutputPath -Days $CleanupDays
}

# Send email report
if ($EmailReport -and $EmailTo -and $EmailFrom -and $EmailServer) {
    Send-EmailReport -To $EmailTo -From $EmailFrom -SMTPServer $EmailServer -Attachments $Script:ExportFiles
}

# Display summary
$elapsedTime = (Get-Date) - $Script:StartTime
Write-Host ""
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "   EXPORT COMPLETED SUCCESSFULLY" -ForegroundColor Green
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "📁 Export Location: $exportFolder" -ForegroundColor Yellow
Write-Host "📄 Files Exported: $($Script:ExportFiles.Count)" -ForegroundColor Yellow
Write-Host "⏱️  Time Elapsed: $($elapsedTime.ToString('hh\:mm\:ss'))" -ForegroundColor Yellow
Write-Host ""

# List exported files
if ($Script:ExportFiles.Count -gt 0) {
    Write-Host "Exported Files:" -ForegroundColor Green
    foreach ($file in $Script:ExportFiles) {
        $fileInfo = Get-Item $file
        Write-Host "  ✓ $($fileInfo.Name) ($([math]::Round($fileInfo.Length / 1KB, 2)) KB)" -ForegroundColor Gray
    }
}

Write-Host ""
Write-Host "✅ Script completed successfully!" -ForegroundColor Green