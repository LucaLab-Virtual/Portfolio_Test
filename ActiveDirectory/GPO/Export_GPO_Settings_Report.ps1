<#
.SYNOPSIS
    Exports comprehensive GPO settings reports in multiple formats (HTML, CSV, XML, JSON).
.DESCRIPTION
    This script generates detailed reports of GPO settings including all policy configurations,
    security settings, registry settings, and deployment information. Supports single GPO,
    multiple GPOs, or all GPOs with filtering capabilities.
.PARAMETER GPOName
    Name of the specific GPO to export (supports wildcards)
.PARAMETER GPOList
    Array of GPO names to export
.PARAMETER ExportAll
    Exports all GPOs in the domain
.PARAMETER OutputPath
    Directory path for export files (default: current directory)
.PARAMETER OutputFormat
    Export format: HTML, CSV, XML, JSON, or All (default: HTML)
.PARAMETER ReportType
    Type of report: Full, Summary, Security, Registry, Deployment (default: Full)
.PARAMETER IncludeLinks
    Includes GPO link information
.PARAMETER IncludeSecurity
    Includes security filtering information
.PARAMETER IncludeWMI
    Includes WMI filter information
.PARAMETER IncludeStatus
    Includes GPO status and version information
.PARAMETER IncludeComments
    Includes GPO comments and descriptions
.PARAMETER ExportDetailed
    Exports detailed policy settings (requires GP module)
.PARAMETER FilterGPO
    Filter by GPO status: All, Enabled, Disabled, NotLinked (default: All)
.PARAMETER FilterByOU
    Filter by linked OU (supports wildcard)
.PARAMETER CompareGPOs
    Compares differences between GPOs (specify at least 2 GPO names)
.PARAMETER EmailReport
    Sends report via email (requires SMTP configuration)
.PARAMETER CompressOutput
    Creates a ZIP archive of the report
.PARAMETER Domain
    Target domain
.PARAMETER Force
    Overwrites existing output files
.PARAMETER WhatIf
    Shows what would happen without performing the export
.EXAMPLE
    .\Export_GPO_Settings_Report.ps1 -GPOName "Default Domain Policy" -OutputPath "C:\Reports"
    Exports a single GPO to HTML report
.EXAMPLE
    .\Export_GPO_Settings_Report.ps1 -ExportAll -OutputFormat "All" -IncludeLinks -IncludeSecurity
    Exports all GPOs in all formats with detailed information
.EXAMPLE
    .\Export_GPO_Settings_Report.ps1 -GPOList "Policy1","Policy2" -OutputFormat "HTML" -CompareGPOs
    Exports two GPOs and generates a comparison report
.EXAMPLE
    .\Export_GPO_Settings_Report.ps1 -ExportAll -FilterGPO "Enabled" -FilterByOU "OU=Workstations"
    Exports only enabled GPOs linked to Workstations OU
.NOTES
    Requires Group Policy Management Console (GPMC) module
    Requires HTML to PDF conversion for PDF reports (optional)
    Must be run with administrative privileges
    Author: Portfolio Script
    Date: $(Get-Date -Format "yyyy-MM-dd")
#>

<#
    Multiple Export Formats:

        HTML (interactive, styled reports)

        CSV (for Excel/analysis)

        XML (for programmatic processing)

        JSON (for API/integration)

    Flexible Reporting:

        Full report with all details

        Summary report

        Security settings only

        Registry settings only

        Deployment information

    Filtering Capabilities:

        By GPO status (Enabled/Disabled)

        By linked OU

        By name (wildcards supported)

    Comprehensive Data Collection:

        GPO basic information

        Security filtering

        WMI filters

        Link information

        Detailed policy settings

        Registry settings

        Security settings

    Advanced Features:

        GPO comparison

        Email reports

        Compressed output (ZIP)

        Interactive HTML with accordion sections

        Progress tracking

Usage Examples:
powershell

# Export single GPO to HTML
.\Export_GPO_Settings_Report.ps1 -GPOName "Default Domain Policy" -OutputPath "C:\Reports"

# Export all GPOs in all formats
.\Export_GPO_Settings_Report.ps1 -ExportAll -OutputFormat "All" -OutputPath "C:\Reports"

# Export with filtering
.\Export_GPO_Settings_Report.ps1 -ExportAll -FilterGPO "Enabled" -FilterByOU "OU=Workstations" -OutputFormat "HTML"

# Export with detailed settings
.\Export_GPO_Settings_Report.ps1 -GPOName "Security Policy" -ExportDetailed -IncludeLinks -IncludeSecurity

# Compare two GPOs
.\Export_GPO_Settings_Report.ps1 -GPOName "Policy1" -GPOList "Policy2" -CompareGPOs -OutputFormat "HTML"

# Export and email report
.\Export_GPO_Settings_Report.ps1 -ExportAll -OutputFormat "HTML" -CompressOutput -EmailReport @{
    SMTPServer="smtp.domain.com"
    SMTPPort=587
    From="reports@domain.com"
    To="admin@domain.com"
}

# Preview what would be exported
.\Export_GPO_Settings_Report.ps1 -ExportAll -WhatIf

Generated Reports:
HTML Report Features:

    Interactive accordion sections

    Color-coded status badges

    Statistics dashboard

    Expandable detailed settings

    Print-friendly layout

CSV Report:

    Flat structure for Excel import

    All major GPO properties

    Link count and enforcement stats

XML/JSON Reports:

    Complete hierarchical structure

    All settings preserved

    Suitable for automation
#>

[CmdletBinding(DefaultParameterSetName = 'ExportAll')]
param(
    [Parameter(ParameterSetName = 'Single')]
    [Parameter(ParameterSetName = 'Compare')]
    [string]$GPOName,
    
    [Parameter(ParameterSetName = 'Multiple')]
    [string[]]$GPOList,
    
    [Parameter(ParameterSetName = 'ExportAll')]
    [switch]$ExportAll,
    
    [Parameter()]
    [string]$OutputPath = ".\GPO_Reports",
    
    [Parameter()]
    [ValidateSet('HTML', 'CSV', 'XML', 'JSON', 'All')]
    [string]$OutputFormat = "HTML",
    
    [Parameter()]
    [ValidateSet('Full', 'Summary', 'Security', 'Registry', 'Deployment')]
    [string]$ReportType = "Full",
    
    [Parameter()]
    [switch]$IncludeLinks,
    
    [Parameter()]
    [switch]$IncludeSecurity,
    
    [Parameter()]
    [switch]$IncludeWMI,
    
    [Parameter()]
    [switch]$IncludeStatus,
    
    [Parameter()]
    [switch]$IncludeComments,
    
    [Parameter()]
    [switch]$ExportDetailed,
    
    [Parameter()]
    [ValidateSet('All', 'Enabled', 'Disabled', 'NotLinked')]
    [string]$FilterGPO = "All",
    
    [Parameter()]
    [string]$FilterByOU,
    
    [Parameter(ParameterSetName = 'Compare')]
    [switch]$CompareGPOs,
    
    [Parameter()]
    [PSCustomObject]$EmailConfig,
    
    [Parameter()]
    [switch]$CompressOutput,
    
    [Parameter()]
    [string]$Domain,
    
    [Parameter()]
    [switch]$Force,
    
    [Parameter()]
    [switch]$WhatIf
)

# Set error handling
$ErrorActionPreference = "Stop"
$script:StartTime = Get-Date
$script:ExportedCount = 0
$script:FailedCount = 0
$script:SkippedCount = 0
$script:ReportFiles = @()

# Function to write colored output
function Write-OutputColor {
    param(
        [string]$Message,
        [string]$Color = "White"
    )
    Write-Host "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] " -NoNewline -ForegroundColor Gray
    Write-Host $Message -ForegroundColor $Color
}

# Function to check administrator privileges
function Test-Administrator {
    $identity = [System.Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object System.Security.Principal.WindowsPrincipal($identity)
    return $principal.IsInRole([System.Security.Principal.WindowsBuiltInRole]::Administrator)
}

# Function to check GPMC module
function Test-GPModule {
    try {
        Import-Module GroupPolicy -ErrorAction SilentlyContinue
        return (Get-Module -Name GroupPolicy -ListAvailable) -ne $null
    } catch {
        return $false
    }
}

# Function to check AD module
function Test-ADModule {
    try {
        Import-Module ActiveDirectory -ErrorAction SilentlyContinue
        return (Get-Module -Name ActiveDirectory -ListAvailable) -ne $null
    } catch {
        return $false
    }
}

# Function to get GPOs with filters
function Get-GPOsWithFilters {
    param(
        [string]$FilterStatus,
        [string]$FilterOU
    )
    
    $allGPOs = Get-GPO -All -ErrorAction Stop
    $filteredGPOs = @()
    
    foreach ($gpo in $allGPOs) {
        $include = $true
        
        # Filter by status
        if ($FilterStatus -ne "All") {
            $gpoStatus = $gpo.GpoStatus
            if ($FilterStatus -eq "Enabled" -and $gpoStatus -ne "AllSettingsEnabled") {
                $include = $false
            } elseif ($FilterStatus -eq "Disabled" -and $gpoStatus -eq "AllSettingsEnabled") {
                $include = $false
            }
        }
        
        # Filter by OU link
        if ($FilterOU -and $include) {
            try {
                $links = Get-GPOReport -Name $gpo.DisplayName -ReportType XML -ErrorAction SilentlyContinue
                if ($links -match "<LinksTo>") {
                    $linksMatch = [regex]::Matches($links, '<LinksTo[^>]+>')
                    $hasMatchingOU = $false
                    foreach ($link in $linksMatch) {
                        if ($link.Value -match $FilterOU) {
                            $hasMatchingOU = $true
                            break
                        }
                    }
                    if (-not $hasMatchingOU) {
                        $include = $false
                    }
                } else {
                    if ($FilterStatus -eq "NotLinked") {
                        $include = $true
                    } else {
                        $include = $false
                    }
                }
            } catch {
                $include = $false
            }
        }
        
        if ($include) {
            $filteredGPOs += $gpo
        }
    }
    
    return $filteredGPOs
}

# Function to get GPO detailed settings
function Get-GPODetailedSettings {
    param(
        [string]$GPOName,
        [string]$ReportType = "Full"
    )
    
    $settings = @{
        Name = $GPOName
        Id = $null
        Domain = $null
        Description = ""
        Comment = ""
        GpoStatus = ""
        CreationTime = $null
        ModificationTime = $null
        WmiFilter = ""
        SecurityFilter = @()
        Links = @()
        ComputerSettings = @()
        UserSettings = @()
        RegistrySettings = @()
        SecuritySettings = @()
        DetailedPolicy = $null
    }
    
    try {
        # Get GPO basic info
        $gpo = Get-GPO -Name $GPOName -ErrorAction Stop
        $settings.Id = $gpo.Id
        $settings.Domain = $gpo.DomainName
        $settings.Description = $gpo.Description
        $settings.Comment = $gpo.Comment
        $settings.GpoStatus = $gpo.GpoStatus
        $settings.CreationTime = $gpo.CreationTime
        $settings.ModificationTime = $gpo.ModificationTime
        
        # Get WMI filter
        if ($IncludeWMI -or $ReportType -eq "Full") {
            try {
                $wmiFilter = Get-GPWMIFilter -Name $GPOName -ErrorAction SilentlyContinue
                if ($wmiFilter) {
                    $settings.WmiFilter = $wmiFilter.Name
                }
            } catch {
                # WMI filter not found
            }
        }
        
        # Get security filtering
        if ($IncludeSecurity -or $ReportType -eq "Full" -or $ReportType -eq "Security") {
            try {
                $security = Get-GPPermission -Name $GPOName -ErrorAction SilentlyContinue
                $settings.SecurityFilter = $security | Where-Object { 
                    $_.Permission -eq "GpoApply" -or $_.Permission -eq "GpoRead"
                } | ForEach-Object {
                    $_.Trustee.Name
                }
            } catch {
                # Security filtering not available
            }
        }
        
        # Get links
        if ($IncludeLinks -or $ReportType -eq "Full" -or $ReportType -eq "Deployment") {
            try {
                $report = Get-GPOReport -Name $GPOName -ReportType XML -ErrorAction SilentlyContinue
                if ($report -match "<LinksTo>") {
                    $linksMatch = [regex]::Matches($report, '<LinksTo[^>]+>')
                    foreach ($link in $linksMatch) {
                        $linkData = @{
                            Target = $null
                            Enforced = $false
                            Order = 0
                            Disabled = $false
                        }
                        
                        if ($link.Value -match 'Target="([^"]+)"') {
                            $linkData.Target = $matches[1]
                        }
                        if ($link.Value -match 'Enforced="([^"]+)"') {
                            $linkData.Enforced = $matches[1] -eq "1"
                        }
                        if ($link.Value -match 'Order="([^"]+)"') {
                            $linkData.Order = [int]$matches[1]
                        }
                        if ($link.Value -match 'Disabled="([^"]+)"') {
                            $linkData.Disabled = $matches[1] -eq "1"
                        }
                        
                        $settings.Links += $linkData
                    }
                }
            } catch {
                # Links not available
            }
        }
        
        # Get detailed policy settings
        if ($ExportDetailed -or $ReportType -eq "Full") {
            try {
                $report = Get-GPOReport -Name $GPOName -ReportType XML -ErrorAction Stop
                $xml = [xml]$report
                
                # Parse computer settings
                if ($xml.GPO.Computer -and $xml.GPO.Computer.ExtensionData) {
                    foreach ($extension in $xml.GPO.Computer.ExtensionData) {
                        $extensionData = @{
                            Name = $extension.extensionId
                            Settings = $extension.InnerXml
                        }
                        $settings.ComputerSettings += $extensionData
                    }
                }
                
                # Parse user settings
                if ($xml.GPO.User -and $xml.GPO.User.ExtensionData) {
                    foreach ($extension in $xml.GPO.User.ExtensionData) {
                        $extensionData = @{
                            Name = $extension.extensionId
                            Settings = $extension.InnerXml
                        }
                        $settings.UserSettings += $extensionData
                    }
                }
                
                # Parse registry settings
                if ($xml.GPO.Computer -and $xml.GPO.Computer.ExtensionData -and 
                    $xml.GPO.Computer.ExtensionData.extensionId -contains "Registry") {
                    # Extract registry settings
                    $regNodes = $xml.SelectNodes("//Registry")
                    foreach ($node in $regNodes) {
                        $regSetting = @{
                            Key = $node.Key
                            Value = $node.Value
                            Data = $node.Data
                            Type = $node.Type
                        }
                        $settings.RegistrySettings += $regSetting
                    }
                }
                
                # Parse security settings
                if ($ReportType -eq "Security" -or $ReportType -eq "Full") {
                    # Extract security settings from the report
                    $securityNodes = $xml.SelectNodes("//SecurityOption")
                    foreach ($node in $securityNodes) {
                        $securitySetting = @{
                            Name = $node.Name
                            Value = $node.Value
                            Type = $node.Type
                        }
                        $settings.SecuritySettings += $securitySetting
                    }
                }
                
                $settings.DetailedPolicy = $xml
                
            } catch {
                Write-OutputColor "  → Could not get detailed settings: $($_.Exception.Message)" "Yellow"
            }
        }
        
        return $settings
        
    } catch {
        Write-OutputColor "  → Failed to get GPO settings: $($_.Exception.Message)" "Red"
        return $null
    }
}

# Function to generate HTML report
function New-GPOHTMLReport {
    param(
        [array]$GPOData,
        [string]$OutputPath,
        [string]$ReportType,
        [bool]$CompareMode = $false
    )
    
    $htmlContent = @"
<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <title>GPO Settings Report - $(Get-Date -Format 'yyyy-MM-dd HH:mm')</title>
    <style>
        body { font-family: Segoe UI, Tahoma, Geneva, Verdana, sans-serif; margin: 20px; background-color: #f5f5f5; }
        .container { max-width: 1400px; margin: 0 auto; background-color: white; padding: 20px; border-radius: 8px; box-shadow: 0 2px 4px rgba(0,0,0,0.1); }
        h1 { color: #2c3e50; border-bottom: 3px solid #3498db; padding-bottom: 10px; }
        h2 { color: #34495e; margin-top: 25px; border-left: 4px solid #3498db; padding-left: 10px; }
        h3 { color: #555; margin-top: 20px; }
        .header-info { background: #ecf0f1; padding: 15px; border-radius: 5px; margin-bottom: 20px; }
        .gpo-card { 
            border: 1px solid #ddd; 
            border-radius: 5px; 
            margin-bottom: 20px; 
            padding: 15px;
            background: #fafafa;
            box-shadow: 0 1px 3px rgba(0,0,0,0.1);
        }
        .gpo-header { 
            background: #3498db; 
            color: white; 
            padding: 10px; 
            margin: -15px -15px 15px -15px;
            border-radius: 5px 5px 0 0;
        }
        .gpo-header.enabled { background: #27ae60; }
        .gpo-header.disabled { background: #e74c3c; }
        .table { 
            width: 100%; 
            border-collapse: collapse; 
            margin: 10px 0;
            font-size: 13px;
        }
        .table th { 
            background: #34495e; 
            color: white; 
            padding: 8px; 
            text-align: left;
            border: 1px solid #2c3e50;
        }
        .table td { 
            padding: 8px; 
            border: 1px solid #ddd; 
            vertical-align: top;
        }
        .table tr:nth-child(even) { background: #f9f9f9; }
        .table tr:hover { background: #f1f1f1; }
        .badge {
            display: inline-block;
            padding: 3px 8px;
            border-radius: 3px;
            font-size: 11px;
            font-weight: bold;
        }
        .badge-success { background: #27ae60; color: white; }
        .badge-danger { background: #e74c3c; color: white; }
        .badge-warning { background: #f39c12; color: white; }
        .badge-info { background: #3498db; color: white; }
        .badge-secondary { background: #95a5a6; color: white; }
        .filter-section { 
            background: #f8f9fa; 
            padding: 15px; 
            border-radius: 5px;
            margin-bottom: 20px;
            border: 1px solid #dee2e6;
        }
        .stats-grid {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(200px, 1fr));
            gap: 15px;
            margin: 20px 0;
        }
        .stat-card {
            background: #f8f9fa;
            padding: 15px;
            border-radius: 5px;
            text-align: center;
            border: 1px solid #dee2e6;
        }
        .stat-number {
            font-size: 28px;
            font-weight: bold;
            color: #2c3e50;
        }
        .stat-label {
            color: #7f8c8d;
            font-size: 14px;
            margin-top: 5px;
        }
        .accordion { 
            cursor: pointer; 
            padding: 10px; 
            border: none; 
            text-align: left; 
            outline: none;
            font-size: 15px;
            font-weight: bold;
            transition: 0.4s;
            background: #ecf0f1;
            width: 100%;
            margin: 5px 0;
        }
        .accordion.active, .accordion:hover { background: #d5dbdb; }
        .panel { 
            padding: 0 18px; 
            display: none; 
            background-color: white; 
            overflow: hidden;
            border: 1px solid #ddd;
            border-top: none;
        }
        .panel.show { display: block; }
        .settings-panel {
            max-height: 400px;
            overflow-y: auto;
            background: #f8f9fa;
            padding: 10px;
            border-radius: 3px;
            font-family: Consolas, monospace;
            font-size: 12px;
            white-space: pre-wrap;
            word-wrap: break-word;
        }
        .comparison-table td.different { background: #ffeaa7; }
        .comparison-table td.missing { background: #ff7675; color: white; }
        .comparison-table td.new { background: #55efc4; }
        @media print {
            .no-print { display: none; }
            body { background-color: white; }
            .container { box-shadow: none; }
        }
    </style>
    <script>
        function toggleAccordion(element) {
            element.classList.toggle("active");
            var panel = element.nextElementSibling;
            if (panel.classList.contains("show")) {
                panel.classList.remove("show");
            } else {
                panel.classList.add("show");
            }
        }
    </script>
</head>
<body>
    <div class="container">
        <h1>Group Policy Objects Settings Report</h1>
        <div class="header-info">
            <p><strong>Generated:</strong> $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')</p>
            <p><strong>Domain:</strong> $((Get-ADDomain).DNSRoot)</p>
            <p><strong>Report Type:</strong> $ReportType</p>
            <p><strong>Total GPOs:</strong> $($GPOData.Count)</p>
        </div>
"@

    # Add statistics
    $htmlContent += @"
        <div class="stats-grid">
            <div class="stat-card">
                <div class="stat-number">$($GPOData.Count)</div>
                <div class="stat-label">Total GPOs</div>
            </div>
            <div class="stat-card">
                <div class="stat-number">$(($GPOData | Where-Object { $_.GpoStatus -eq "AllSettingsEnabled" }).Count)</div>
                <div class="stat-label">Enabled</div>
            </div>
            <div class="stat-card">
                <div class="stat-number">$(($GPOData | Where-Object { $_.GpoStatus -ne "AllSettingsEnabled" }).Count)</div>
                <div class="stat-label">Disabled/Partially Disabled</div>
            </div>
            <div class="stat-card">
                <div class="stat-number">$(($GPOData | Where-Object { $_.Links.Count -gt 0 }).Count)</div>
                <div class="stat-label">Linked</div>
            </div>
        </div>
"@

    # Generate GPO cards
    foreach ($gpo in $GPOData) {
        $statusClass = if ($gpo.GpoStatus -eq "AllSettingsEnabled") { "enabled" } else { "disabled" }
        $statusText = if ($gpo.GpoStatus -eq "AllSettingsEnabled") { "Enabled" } else { $gpo.GpoStatus }
        
        $htmlContent += @"
        <div class="gpo-card">
            <div class="gpo-header $statusClass">
                <h3 style="margin: 0; color: white;">$($gpo.Name)</h3>
            </div>
            
            <h4>GPO Information</h4>
            <table class="table">
                <tr>
                    <td><strong>ID:</strong></td>
                    <td>$($gpo.Id)</td>
                    <td><strong>Domain:</strong></td>
                    <td>$($gpo.Domain)</td>
                </tr>
                <tr>
                    <td><strong>Status:</strong></td>
                    <td><span class="badge badge-$($statusClass)">$statusText</span></td>
                    <td><strong>Created:</strong></td>
                    <td>$($gpo.CreationTime)</td>
                </tr>
                <tr>
                    <td><strong>Modified:</strong></td>
                    <td>$($gpo.ModificationTime)</td>
                    <td><strong>WMI Filter:</strong></td>
                    <td>$($gpo.WmiFilter)</td>
                </tr>
"@

        if ($IncludeComments) {
            $htmlContent += @"
                <tr>
                    <td><strong>Description:</strong></td>
                    <td colspan="3">$($gpo.Description)</td>
                </tr>
                <tr>
                    <td><strong>Comment:</strong></td>
                    <td colspan="3">$($gpo.Comment)</td>
                </tr>
"@
        }

        $htmlContent += @"
            </table>
            
            <h4>Security Filtering</h4>
            <table class="table">
                <tr>
                    <th>Security Group</th>
                </tr>
"@

        if ($gpo.SecurityFilter -and $gpo.SecurityFilter.Count -gt 0) {
            foreach ($filter in $gpo.SecurityFilter) {
                $htmlContent += @"
                <tr>
                    <td>$filter</td>
                </tr>
"@
            }
        } else {
            $htmlContent += @"
                <tr>
                    <td><em>No security filtering</em></td>
                </tr>
"@
        }

        $htmlContent += @"
            </table>
            
            <h4>Links</h4>
            <table class="table">
                <tr>
                    <th>Target OU</th>
                    <th>Order</th>
                    <th>Enforced</th>
                    <th>Disabled</th>
                </tr>
"@

        if ($gpo.Links -and $gpo.Links.Count -gt 0) {
            foreach ($link in $gpo.Links) {
                $htmlContent += @"
                <tr>
                    <td>$($link.Target)</td>
                    <td>$($link.Order)</td>
                    <td>$(if ($link.Enforced) { '<span class="badge badge-success">Yes</span>' } else { '<span class="badge badge-secondary">No</span>' })</td>
                    <td>$(if ($link.Disabled) { '<span class="badge badge-danger">Yes</span>' } else { '<span class="badge badge-success">No</span>' })</td>
                </tr>
"@
            }
        } else {
            $htmlContent += @"
                <tr>
                    <td colspan="4"><em>Not linked</em></td>
                </tr>
"@
        }

        $htmlContent += @"
            </table>
"@

        # Add detailed settings if available
        if ($ExportDetailed) {
            $htmlContent += @"
            <h4>Detailed Settings</h4>
            <button class="accordion" onclick="toggleAccordion(this)">Computer Settings</button>
            <div class="panel show">
                <div class="settings-panel">
                    <pre>$($gpo.ComputerSettings | ConvertTo-Json -Depth 3)</pre>
                </div>
            </div>
            
            <button class="accordion" onclick="toggleAccordion(this)">User Settings</button>
            <div class="panel show">
                <div class="settings-panel">
                    <pre>$($gpo.UserSettings | ConvertTo-Json -Depth 3)</pre>
                </div>
            </div>
            
            <button class="accordion" onclick="toggleAccordion(this)">Registry Settings</button>
            <div class="panel show">
                <div class="settings-panel">
                    <pre>$($gpo.RegistrySettings | ConvertTo-Json -Depth 3)</pre>
                </div>
            </div>
            
            <button class="accordion" onclick="toggleAccordion(this)">Security Settings</button>
            <div class="panel show">
                <div class="settings-panel">
                    <pre>$($gpo.SecuritySettings | ConvertTo-Json -Depth 3)</pre>
                </div>
            </div>
"@
        }

        $htmlContent += @"
        </div>
"@
    }

    # Close HTML
    $htmlContent += @"
        <div class="footer no-print" style="margin-top: 30px; padding-top: 20px; border-top: 1px solid #ddd; text-align: center; color: #7f8c8d; font-size: 12px;">
            <p>Report generated by GPO Settings Export Script v1.0</p>
            <p>Generated on $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')</p>
        </div>
    </div>
</body>
</html>
"@

    $htmlPath = Join-Path $OutputPath "GPO_Report_$(Get-Date -Format 'yyyyMMdd_HHmmss').html"
    $htmlContent | Out-File -FilePath $htmlPath -Encoding UTF8 -Force
    $script:ReportFiles += $htmlPath
    
    return $htmlPath
}

# Function to generate CSV report
function New-GPOCSVReport {
    param(
        [array]$GPOData,
        [string]$OutputPath
    )
    
    $csvData = @()
    
    foreach ($gpo in $GPOData) {
        $row = [PSCustomObject]@{
            GPO_Name = $gpo.Name
            GPO_ID = $gpo.Id
            Domain = $gpo.Domain
            Description = $gpo.Description
            Status = $gpo.GpoStatus
            Created = $gpo.CreationTime
            Modified = $gpo.ModificationTime
            WMI_Filter = $gpo.WmiFilter
            Security_Filter = ($gpo.SecurityFilter -join "; ")
            Links = ($gpo.Links | ForEach-Object { $_.Target }) -join "; "
            Link_Count = $gpo.Links.Count
            Enforced_Links = ($gpo.Links | Where-Object { $_.Enforced }).Count
        }
        $csvData += $row
    }
    
    $csvPath = Join-Path $OutputPath "GPO_Report_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
    $csvData | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8 -Force
    $script:ReportFiles += $csvPath
    
    return $csvPath
}

# Function to generate XML report
function New-GPOXMLReport {
    param(
        [array]$GPOData,
        [string]$OutputPath
    )
    
    $xmlPath = Join-Path $OutputPath "GPO_Report_$(Get-Date -Format 'yyyyMMdd_HHmmss').xml"
    $gpoData | Export-Clixml -Path $xmlPath -Depth 5 -Force
    $script:ReportFiles += $xmlPath
    
    return $xmlPath
}

# Function to generate JSON report
function New-GPOJSONReport {
    param(
        [array]$GPOData,
        [string]$OutputPath
    )
    
    $jsonPath = Join-Path $OutputPath "GPO_Report_$(Get-Date -Format 'yyyyMMdd_HHmmss').json"
    $gpoData | ConvertTo-Json -Depth 5 | Out-File -FilePath $jsonPath -Encoding UTF8 -Force
    $script:ReportFiles += $jsonPath
    
    return $jsonPath
}

# Function to generate comparison report
function New-GPOComparisonReport {
    param(
        [array]$GPOData,
        [string]$OutputPath
    )
    
    if ($GPOData.Count -lt 2) {
        Write-OutputColor "Need at least 2 GPOs for comparison" "Red"
        return $null
    }
    
    $htmlContent = @"
<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <title>GPO Comparison Report - $(Get-Date -Format 'yyyy-MM-dd HH:mm')</title>
    <style>
        body { font-family: Segoe UI, Tahoma, Geneva, Verdana, sans-serif; margin: 20px; background-color: #f5f5f5; }
        .container { max-width: 1400px; margin: 0 auto; background-color: white; padding: 20px; border-radius: 8px; box-shadow: 0 2px 4px rgba(0,0,0,0.1); }
        h1 { color: #2c3e50; border-bottom: 3px solid #3498db; padding-bottom: 10px; }
        .table { width: 100%; border-collapse: collapse; margin: 10px 0; font-size: 13px; }
        .table th { background: #34495e; color: white; padding: 8px; text-align: left; border: 1px solid #2c3e50; }
        .table td { padding: 8px; border: 1px solid #ddd; vertical-align: top; }
        .table tr:nth-child(even) { background: #f9f9f9; }
        .table tr:hover { background: #f1f1f1; }
        .different { background: #ffeaa7; }
        .missing { background: #ff7675; color: white; }
        .new { background: #55efc4; }
    </style>
</head>
<body>
    <div class="container">
        <h1>GPO Comparison Report</h1>
        <p><strong>Generated:</strong> $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')</p>
        <p><strong>GPOs Compared:</strong> $(($GPOData | ForEach-Object { $_.Name }) -join ', ')</p>
        
        <table class="table">
            <tr>
                <th>Property</th>
"@

    foreach ($gpo in $GPOData) {
        $htmlContent += "<th>$($gpo.Name)</th>"
    }
    
    $htmlContent += @"
            </tr>
"@

    # Compare properties
    $properties = @('Description', 'GpoStatus', 'WmiFilter', 'SecurityFilter', 'Links')
    
    foreach ($prop in $properties) {
        $htmlContent += "<tr><td><strong>$prop</strong></td>"
        $values = @()
        foreach ($gpo in $GPOData) {
            $value = $gpo.$prop
            if ($value -is [array]) {
                $value = $value -join ", "
            }
            $values += $value
        }
        
        foreach ($value in $values) {
            $class = ""
            if ($values -and $value -ne $values[0]) {
                $class = "different"
            }
            $htmlContent += "<td class='$class'>$value</td>"
        }
        $htmlContent += "</tr>"
    }
    
    $htmlContent += @"
        </table>
    </div>
</body>
</html>
"@

    $comparePath = Join-Path $OutputPath "GPO_Comparison_$(Get-Date -Format 'yyyyMMdd_HHmmss').html"
    $htmlContent | Out-File -FilePath $comparePath -Encoding UTF8 -Force
    $script:ReportFiles += $comparePath
    
    return $comparePath
}

# Function to compress output
function Compress-ReportOutput {
    param(
        [string]$OutputPath,
        [array]$Files
    )
    
    try {
        $zipPath = Join-Path $OutputPath "GPO_Reports_$(Get-Date -Format 'yyyyMMdd_HHmmss').zip"
        
        if (Get-Command Compress-Archive -ErrorAction SilentlyContinue) {
            Compress-Archive -Path $Files -DestinationPath $zipPath -Force
        } else {
            # Use .NET for compression
            Add-Type -AssemblyName System.IO.Compression.FileSystem
            $zip = [System.IO.Compression.ZipFile]::Open($zipPath, 'Create')
            foreach ($file in $Files) {
                $entryName = [System.IO.Path]::GetFileName($file)
                [System.IO.Compression.ZipFileExtensions]::CreateEntryFromFile($zip, $file, $entryName)
            }
            $zip.Dispose()
        }
        
        Write-OutputColor "  → Compressed report: $zipPath" "Green"
        $script:ReportFiles += $zipPath
        
    } catch {
        Write-OutputColor "  → Failed to compress output: $($_.Exception.Message)" "Yellow"
    }
}

# Function to send email
function Send-GPOReportEmail {
    param(
        [PSCustomObject]$Config,
        [array]$Files
    )
    
    try {
        $smtpServer = $Config.SMTPServer
        $smtpPort = $Config.SMTPPort
        $username = $Config.Username
        $password = $Config.Password
        $from = $Config.From
        $to = $Config.To
        $subject = "GPO Settings Report - $(Get-Date -Format 'yyyy-MM-dd')"
        $body = @"
GPO Settings Report

Generated: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
Domain: $(Get-ADDomain).DNSRoot
Number of GPOs reported: $($script:ExportedCount)
Report files attached: $(($Files | ForEach-Object { Split-Path $_ -Leaf }) -join ', ')

This report was automatically generated by the GPO Settings Export Script.
"@
        
        $mailParams = @{
            SmtpServer = $smtpServer
            Port = $smtpPort
            From = $from
            To = $to
            Subject = $subject
            Body = $body
            Attachments = $Files
        }
        
        if ($username -and $password) {
            $securePassword = ConvertTo-SecureString $password -AsPlainText -Force
            $credential = New-Object System.Management.Automation.PSCredential($username, $securePassword)
            $mailParams.Credential = $credential
            $mailParams.UseSsl = $true
        }
        
        Send-MailMessage @mailParams
        Write-OutputColor "  → Email report sent to: $to" "Green"
        
    } catch {
        Write-OutputColor "  → Failed to send email: $($_.Exception.Message)" "Yellow"
    }
}

# Main execution
try {
    Write-OutputColor "===== GPO Settings Export Script Started =====" "Cyan"
    
    # Check administrator rights
    if (-not (Test-Administrator)) {
        Write-OutputColor "This script requires administrative privileges." "Red"
        Write-OutputColor "Please run PowerShell as Administrator." "Red"
        exit 1
    }
    
    # Check Group Policy module
    if (-not (Test-GPModule)) {
        Write-OutputColor "Group Policy Management Console (GPMC) module is not available." "Red"
        Write-OutputColor "Please install RSAT-AD-PowerShell feature." "Red"
        Write-OutputColor "Run: Install-WindowsFeature RSAT-AD-PowerShell" "Yellow"
        exit 1
    }
    
    # Check AD module
    if (-not (Test-ADModule)) {
        Write-OutputColor "Active Directory module is not available." "Red"
        Write-OutputColor "Please install RSAT-AD-PowerShell feature." "Red"
        Write-OutputColor "Run: Install-WindowsFeature RSAT-AD-PowerShell" "Yellow"
        exit 1
    }
    
    # Import required modules
    Import-Module GroupPolicy -Force
    Import-Module ActiveDirectory -Force
    
    # Verify domain connectivity
    try {
        $domain = Get-ADDomain -ErrorAction Stop
        Write-OutputColor "Connected to domain: $($domain.DNSRoot)" "Green"
    } catch {
        Write-OutputColor "Cannot connect to Active Directory domain." "Red"
        Write-OutputColor "Error: $_" "Red"
        exit 1
    }
    
    # Create output directory
    $OutputPath = [System.IO.Path]::GetFullPath($OutputPath)
    if (-not (Test-Path $OutputPath)) {
        New-Item -Path $OutputPath -ItemType Directory -Force | Out-Null
        Write-OutputColor "Created output directory: $OutputPath" "Green"
    }
    
    # Get GPO list based on parameters
    $gpoList = @()
    
    if ($PSCmdlet.ParameterSetName -eq 'Single') {
        if ($GPOName -match '\*') {
            $gpoList = Get-GPO -All | Where-Object { $_.DisplayName -like $GPOName }
        } else {
            $gpo = Get-GPO -Name $GPOName -ErrorAction SilentlyContinue
            if ($gpo) { $gpoList = @($gpo) }
        }
    } elseif ($PSCmdlet.ParameterSetName -eq 'Multiple') {
        foreach ($name in $GPOList) {
            $gpo = Get-GPO -Name $name -ErrorAction SilentlyContinue
            if ($gpo) { $gpoList += $gpo }
        }
    } elseif ($PSCmdlet.ParameterSetName -eq 'Compare') {
        $gpoList = @()
        if ($GPOName) {
            $gpo = Get-GPO -Name $GPOName -ErrorAction SilentlyContinue
            if ($gpo) { $gpoList += $gpo }
        }
        if ($GPOList) {
            foreach ($name in $GPOList) {
                $gpo = Get-GPO -Name $name -ErrorAction SilentlyContinue
                if ($gpo) { $gpoList += $gpo }
            }
        }
        if ($gpoList.Count -lt 2) {
            Write-OutputColor "Comparison requires at least 2 GPOs" "Red"
            exit 1
        }
    } else {
        # Export all with filters
        $gpoList = Get-GPOsWithFilters -FilterStatus $FilterGPO -FilterOU $FilterByOU
    }
    
    if ($gpoList.Count -eq 0) {
        Write-OutputColor "No GPOs found matching criteria" "Yellow"
        exit 0
    }
    
    Write-OutputColor "Found $($gpoList.Count) GPO(s) to export" "Green"
    
    # Process each GPO
    $gpoData = @()
    $total = $gpoList.Count
    $current = 0
    
    foreach ($gpo in $gpoList) {
        $current++
        Write-ProgressBar -Current $current -Total $total -Activity "Exporting GPO settings"
        
        Write-OutputColor "`nProcessing: $($gpo.DisplayName)" "White"
        
        if ($WhatIf) {
            Write-OutputColor "[WHATIF] Would export: $($gpo.DisplayName)" "Gray"
            $script:ExportedCount++
            continue
        }
        
        $gpoSettings = Get-GPODetailedSettings -GPOName $gpo.DisplayName -ReportType $ReportType
        
        if ($gpoSettings) {
            $gpoData += $gpoSettings
            $script:ExportedCount++
            Write-OutputColor "  ✓ Collected settings for: $($gpo.DisplayName)" "Green"
        } else {
            $script:FailedCount++
            Write-OutputColor "  ✗ Failed to collect settings for: $($gpo.DisplayName)" "Red"
        }
    }
    
    if ($gpoData.Count -eq 0) {
        Write-OutputColor "No GPO data collected" "Red"
        exit 1
    }
    
    # Generate reports
    Write-OutputColor "`nGenerating reports..." "Yellow"
    
    $formats = if ($OutputFormat -eq "All") { @('HTML', 'CSV', 'XML', 'JSON') } else { @($OutputFormat) }
    
    foreach ($format in $formats) {
        Write-OutputColor "  → Generating $format report..." "Cyan"
        
        try {
            switch ($format) {
                "HTML" {
                    if ($CompareGPOs) {
                        New-GPOComparisonReport -GPOData $gpoData -OutputPath $OutputPath
                    } else {
                        New-GPOHTMLReport -GPOData $gpoData -OutputPath $OutputPath -ReportType $ReportType
                    }
                }
                "CSV" { New-GPOCSVReport -GPOData $gpoData -OutputPath $OutputPath }
                "XML" { New-GPOXMLReport -GPOData $gpoData -OutputPath $OutputPath }
                "JSON" { New-GPOJSONReport -GPOData $gpoData -OutputPath $OutputPath }
            }
        } catch {
            Write-OutputColor "  → Failed to generate $format report: $($_.Exception.Message)" "Red"
        }
    }
    
    # Compress output if requested
    if ($CompressOutput) {
        Compress-ReportOutput -OutputPath $OutputPath -Files $script:ReportFiles
    }
    
    # Send email if configured
    if ($EmailConfig) {
        Send-GPOReportEmail -Config $EmailConfig -Files $script:ReportFiles
    }
    
    # Display summary
    $totalDuration = (Get-Date) - $script:StartTime
    
    Write-OutputColor "`n===== GPO Export Summary =====" "Cyan"
    Write-OutputColor "Total GPOs found: $($gpoList.Count)" "White"
    Write-OutputColor "Successfully exported: $script:ExportedCount" "Green"
    Write-OutputColor "Failed: $script:FailedCount" $(if ($script:FailedCount -gt 0) { "Red" } else { "Green" })
    Write-OutputColor "Total duration: $($totalDuration.ToString('hh\:mm\:ss'))" "White"
    Write-OutputColor "Reports saved to: $OutputPath" "Cyan"
    
    if ($script:ReportFiles.Count -gt 0) {
        Write-OutputColor "`nGenerated reports:" "Cyan"
        foreach ($file in $script:ReportFiles) {
            Write-OutputColor "  ✓ $(Split-Path $file -Leaf)" "Green"
        }
    }
    
    # Exit with appropriate code
    if ($script:FailedCount -gt 0 -and $script:ExportedCount -gt 0) {
        exit 1  # Partial success
    } elseif ($script:FailedCount -gt 0) {
        exit 2  # Complete failure
    } else {
        exit 0  # Complete success
    }
    
} catch {
    Write-OutputColor "Unexpected error: $_" "Red"
    Write-OutputColor "Error details: $($_.Exception.Message)" "Red"
    exit 4
}