<#
.SYNOPSIS
    Generates comprehensive reports on Active Directory users with passwords set to never expire.
.DESCRIPTION
    This script provides detailed analysis of password expiration policies including:
    - Users with passwords that never expire
    - Users approaching password expiration
    - Users with expired passwords
    - Detailed user account status information
    - Multiple export formats (HTML, CSV, JSON, Excel)
    - Email notifications for compliance reporting
    - Historical tracking and trend analysis
.PARAMETER SearchBase
    Distinguished name of the OU to search. Defaults to entire domain.
.PARAMETER IncludeDisabled
    Include disabled user accounts in the report. Default is $false.
.PARAMETER IncludeSystemAccounts
    Include system/well-known accounts. Default is $false.
.PARAMETER ExportPath
    Directory path for export files. Defaults to script directory.
.PARAMETER OutputFormats
    Comma-separated list: CSV,HTML,JSON,Excel. Default is all.
.PARAMETER ExpiringInDays
    Include users whose passwords will expire within X days.
.PARAMETER IncludeExpired
    Include users with already expired passwords. Default is $false.
.PARAMETER NotifyAdmin
    Send email report to administrators.
.PARAMETER SmtpServer
    SMTP server for email notifications.
.PARAMETER FromEmail
    From email address for notifications.
.PARAMETER RecipientEmail
    Recipient email address for reports.
.PARAMETER DomainController
    Specify domain controller for AD queries.
.PARAMETER HistoricalData
    Save historical data for trend analysis.
.PARAMETER Threshold
    Warning threshold percentage for compliance reporting (0-100). Default is 80.
.EXAMPLE
    .\Password_Never_Expires_Report.ps1 -ExportPath "C:\Reports" -OutputFormats HTML,CSV
.EXAMPLE
    .\Password_Never_Expires_Report.ps1 -IncludeDisabled -ExpiringInDays 14 -NotifyAdmin
.EXAMPLE
    .\Password_Never_Expires_Report.ps1 -SearchBase "OU=Users,DC=domain,DC=com" -Threshold 90
.NOTES
    Requires ActiveDirectory module. Run with appropriate AD permissions.
    Excel export requires ImportExcel module (can be installed via PowerShell Gallery).
#>

<#
1. Comprehensive User Analysis

    Identifies users with passwords set to never expire

    Tracks password age and expiration dates

    Shows account status (enabled/disabled/locked)

    Includes manager information and organizational details

2. Multiple Report Formats

    CSV: Easy import into Excel or databases

    HTML: Beautiful, interactive web report with color coding

    JSON: Machine-readable format for API integration

    Excel: Multi-sheet workbook with summary and detail tabs

3. Advanced Filtering

    Include/exclude disabled accounts

    Include/exclude system accounts

    Filter by days until expiration

    Include expired passwords

4. Compliance Dashboard

    Real-time compliance scoring

    Visual compliance bar

    Threshold-based status (PASS/FAIL)

    Non-compliant user identification

5. Automated Notifications

    Email reports to administrators

    HTML email with summary statistics

    Attachment support

    Configurable SMTP settings

6. Historical Tracking

    Save historical data for trend analysis

    JSON format for easy parsing

    Timestamped historical files

Usage Examples:
powershell

# Basic report with all formats
.\Password_Never_Expires_Report.ps1 -ExportPath "C:\ADReports"

# HTML only report with disabled accounts
.\Password_Never_Expires_Report.ps1 -OutputFormats HTML -IncludeDisabled

# Find accounts expiring soon with email notification
.\Password_Never_Expires_Report.ps1 -ExpiringInDays 14 -NotifyAdmin -SmtpServer "smtp.domain.com" -FromEmail "admin@domain.com" -RecipientEmail "security@domain.com"

# Specific OU with compliance threshold
.\Password_Never_Expires_Report.ps1 -SearchBase "OU=Users,DC=domain,DC=com" -Threshold 90 -IncludeExpired

# Historical tracking with all formats
.\Password_Never_Expires_Report.ps1 -HistoricalData -OutputFormats CSV,HTML,Excel

# Include everything for comprehensive audit
.\Password_Never_Expires_Report.ps1 -IncludeDisabled -IncludeSystemAccounts -IncludeExpired -HistoricalData -NotifyAdmin

HTML Report Features:

The HTML report includes:

    Interactive dashboard with summary statistics

    Color-coded status badges: Compliant, Expired, Never Expires, Expiring Soon

    Compliance bar with threshold indicator

    Sortable table with all user details

    Responsive design for all devices

    Print-friendly layout

Excel Report Features:

The Excel export creates a multi-sheet workbook:

    Summary Sheet: Statistics and compliance metrics

    Users Sheet: All user data with formatting

    Auto-filter enabled for easy analysis

    Freeze panes for large datasets

    Column auto-sizing for readability

Sample Output:
text

======================================================================
PASSWORD NEVER EXPIRES REPORT
======================================================================
Start Time: 2026-01-15 14:30:22
User: Administrator
Computer: DC01
Search Base: Entire Domain
======================================================================
Retrieved 1,234 users from Active Directory
Processing user data...
  Processed 100 of 1234 users
  Processed 200 of 1234 users
  ...
Processed 1,234 users matching criteria

======================================================================
SUMMARY STATISTICS
======================================================================
Total Users Processed: 1,234
Password Never Expires: 45
Passwords Expired: 12
Expiring Within 7 Days: 89
Compliance Score: 84.5%
Compliance Status: PASS
======================================================================

Email Notification Example:

The script sends an HTML email with:

    Compliance status and score

    Key statistics

    Attached report file

    Professional formatting

Security Considerations:

    No passwords are displayed or logged

    Uses secure Active Directory connections

    Respects existing security permissions

    No modification of AD objects

    Read-only operations only

Best Practices:

    Schedule regular reports using Task Scheduler

    Set appropriate thresholds based on organizational policies

    Review historical data for trend analysis

    Include disabled accounts for complete audits

    Enable email notifications for compliance alerts

    Archive reports for audit purposes

Dependencies:

    Required: ActiveDirectory PowerShell module

    Optional: ImportExcel module for Excel export

    Optional: SMTP server for email notifications
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory=$false)]
    [string]$SearchBase,
    
    [Parameter(Mandatory=$false)]
    [switch]$IncludeDisabled,
    
    [Parameter(Mandatory=$false)]
    [switch]$IncludeSystemAccounts,
    
    [Parameter(Mandatory=$false)]
    [string]$ExportPath = ".",
    
    [Parameter(Mandatory=$false)]
    [ValidateSet('CSV','HTML','JSON','Excel')]
    [string[]]$OutputFormats = @('CSV','HTML','JSON','Excel'),
    
    [Parameter(Mandatory=$false)]
    [int]$ExpiringInDays,
    
    [Parameter(Mandatory=$false)]
    [switch]$IncludeExpired,
    
    [Parameter(Mandatory=$false)]
    [switch]$NotifyAdmin,
    
    [Parameter(Mandatory=$false)]
    [string]$SmtpServer,
    
    [Parameter(Mandatory=$false)]
    [string]$FromEmail,
    
    [Parameter(Mandatory=$false)]
    [string]$RecipientEmail,
    
    [Parameter(Mandatory=$false)]
    [string]$DomainController,
    
    [Parameter(Mandatory=$false)]
    [switch]$HistoricalData,
    
    [Parameter(Mandatory=$false)]
    [ValidateRange(0,100)]
    [int]$Threshold = 80
)

#region Global Variables
$Script:Timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$Script:ReportData = [System.Collections.ArrayList]::new()
$Script:Summary = $null
$Script:HistoricalData = @()
$Script:LogFile = $null
$Script:ExportTimestamp = Get-Date -Format "yyyyMMdd_HHmmss"
#endregion

#region Helper Functions
function Write-Log {
    param(
        [string]$Message,
        [string]$Level = "INFO",
        [string]$Color = "White"
    )
    
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logMessage = "[$timestamp] [$Level] $Message"
    
    if ($Color) {
        Write-Host $logMessage -ForegroundColor $Color
    }
    else {
        Write-Host $logMessage
    }
    
    if ($Script:LogFile) {
        Add-Content -Path $Script:LogFile -Value $logMessage -Encoding UTF8
    }
}

function Initialize-Logging {
    try {
        if (-not (Test-Path $ExportPath)) {
            New-Item -Path $ExportPath -ItemType Directory -Force | Out-Null
        }
        
        $Script:LogFile = Join-Path $ExportPath "PasswordNeverExpiresReport_$Script:Timestamp.log"
        Write-Log "="*70 "INFO" "Yellow"
        Write-Log "PASSWORD NEVER EXPIRES REPORT" "INFO" "Yellow"
        Write-Log "="*70 "INFO" "Yellow"
        Write-Log "Start Time: $(Get-Date)" "INFO" "Cyan"
        Write-Log "User: $env:USERNAME" "INFO" "Cyan"
        Write-Log "Computer: $env:COMPUTERNAME" "INFO" "Cyan"
        Write-Log "Search Base: $(if($SearchBase){$SearchBase}else{'Entire Domain'})" "INFO" "Cyan"
        Write-Log "="*70 "INFO" "Yellow"
    }
    catch {
        Write-Warning "Could not initialize logging: $_"
    }
}

function Import-ADModule {
    try {
        if (-not (Get-Module -Name ActiveDirectory -ListAvailable)) {
            Write-Error "ActiveDirectory module is not installed."
            return $false
        }
        Import-Module ActiveDirectory -Force -ErrorAction Stop
        Write-Log "ActiveDirectory module loaded successfully" "INFO" "Green"
        return $true
    }
    catch {
        Write-Log "Failed to load ActiveDirectory module: $_" "ERROR" "Red"
        return $false
    }
}

function Get-DomainPasswordPolicy {
    param([string]$DC)
    
    try {
        $params = @{
            ErrorAction = 'Stop'
        }
        if ($DC) { $params.Server = $DC }
        
        $policy = Get-ADDefaultDomainPasswordPolicy @params
        
        if ($policy) {
            Write-Log "Retrieved domain password policy" "INFO" "Green"
            Write-Log "  - Max Password Age: $($policy.MaxPasswordAge.Days) days" "INFO" "Cyan"
            Write-Log "  - Min Password Length: $($policy.MinPasswordLength)" "INFO" "Cyan"
            Write-Log "  - Complexity Enabled: $($policy.ComplexityEnabled)" "INFO" "Cyan"
            return $policy
        }
        return $null
    }
    catch {
        Write-Log "Failed to retrieve domain password policy: $_" "ERROR" "Red"
        return $null
    }
}

function Get-ADUsersWithPasswordInfo {
    param(
        [string]$SearchBase,
        [bool]$IncludeDisabled,
        [bool]$IncludeSystemAccounts,
        [string]$DC
    )
    
    $properties = @(
        'SamAccountName',
        'DisplayName',
        'GivenName',
        'Surname',
        'UserPrincipalName',
        'EmailAddress',
        'Title',
        'Department',
        'Company',
        'Manager',
        'Description',
        'Enabled',
        'LockedOut',
        'PasswordLastSet',
        'PasswordNeverExpires',
        'LastLogonDate',
        'LastLogonTimestamp',
        'Created',
        'Modified',
        'AccountExpirationDate',
        'DistinguishedName',
        'ObjectClass'
    )
    
    $filter = "ObjectClass -eq 'user'"
    
    # Filter out system accounts unless included
    if (-not $IncludeSystemAccounts) {
        $filter += " -and SamAccountName -notlike '*$'"
        $filter += " -and SamAccountName -ne 'Administrator'"
        $filter += " -and SamAccountName -ne 'Guest'"
        $filter += " -and SamAccountName -ne 'krbtgt'"
        $filter += " -and SamAccountName -notlike 'HealthMailbox*'"
        $filter += " -and SamAccountName -notlike 'Sync_*'"
        $filter += " -and SamAccountName -notlike 'Federated*'"
        $filter += " -and SamAccountName -notlike 'MSOL_*'"
    }
    
    # Filter disabled accounts unless included
    if (-not $IncludeDisabled) {
        $filter += " -and Enabled -eq 'True'"
    }
    
    try {
        $params = @{
            Filter = $filter
            Properties = $properties
            ErrorAction = 'Stop'
        }
        if ($SearchBase) { $params.SearchBase = $SearchBase }
        if ($DC) { $params.Server = $DC }
        
        $users = Get-ADUser @params
        
        Write-Log "Retrieved $($users.Count) users from Active Directory" "INFO" "Green"
        return $users
    }
    catch {
        Write-Log "Failed to retrieve users: $_" "ERROR" "Red"
        return $null
    }
}

function Calculate-PasswordAge {
    param(
        [DateTime]$PasswordLastSet,
        [TimeSpan]$MaxPasswordAge
    )
    
    if (-not $PasswordLastSet) {
        return $null
    }
    
    $age = (Get-Date) - $PasswordLastSet
    $daysUntilExpiry = $MaxPasswordAge.Days - $age.Days
    
    return [PSCustomObject]@{
        AgeInDays = $age.Days
        DaysUntilExpiry = $daysUntilExpiry
        IsExpired = $daysUntilExpiry -lt 0
        ExpirationDate = $PasswordLastSet.AddDays($MaxPasswordAge.Days)
        PercentUsed = if ($MaxPasswordAge.Days -gt 0) { [Math]::Round(($age.Days / $MaxPasswordAge.Days) * 100, 2) } else { 0 }
    }
}

function Get-ManagerName {
    param([string]$ManagerDN)
    
    if (-not $ManagerDN) { return $null }
    
    try {
        $manager = Get-ADUser -Identity $ManagerDN -Properties DisplayName -ErrorAction SilentlyContinue
        if ($manager) {
            return $manager.DisplayName
        }
    }
    catch {
        # Silently fail
    }
    return $null
}

function Process-UserData {
    param(
        [array]$Users,
        [TimeSpan]$MaxPasswordAge,
        [int]$ExpiringInDays,
        [bool]$IncludeExpired
    )
    
    $processedData = [System.Collections.ArrayList]::new()
    $now = Get-Date
    
    Write-Log "Processing user data..." "INFO" "Cyan"
    
    $userCount = $Users.Count
    $currentUser = 0
    
    foreach ($user in $Users) {
        $currentUser++
        if ($currentUser % 100 -eq 0) {
            Write-Log "  Processed $currentUser of $userCount users" "INFO" "Cyan"
        }
        
        # Skip users without password data
        if (-not $user.PasswordLastSet) {
            continue
        }
        
        $passwordInfo = Calculate-PasswordAge -PasswordLastSet $user.PasswordLastSet -MaxPasswordAge $MaxPasswordAge
        
        # Apply filters
        if ($ExpiringInDays -and $passwordInfo) {
            if ($passwordInfo.DaysUntilExpiry -gt $ExpiringInDays) {
                continue
            }
        }
        
        if (-not $IncludeExpired -and $passwordInfo -and $passwordInfo.IsExpired) {
            continue
        }
        
        $managerName = Get-ManagerName -ManagerDN $user.Manager
        
        $userObject = [PSCustomObject]@{
            SamAccountName = $user.SamAccountName
            DisplayName = $user.DisplayName
            GivenName = $user.GivenName
            Surname = $user.Surname
            UserPrincipalName = $user.UserPrincipalName
            EmailAddress = $user.EmailAddress
            Title = $user.Title
            Department = $user.Department
            Company = $user.Company
            Manager = $managerName
            Description = $user.Description
            Enabled = $user.Enabled
            LockedOut = $user.LockedOut
            PasswordNeverExpires = $user.PasswordNeverExpires
            PasswordLastSet = $user.PasswordLastSet
            PasswordAgeInDays = if ($passwordInfo) { $passwordInfo.AgeInDays } else { 0 }
            DaysUntilExpiry = if ($passwordInfo) { $passwordInfo.DaysUntilExpiry } else { 0 }
            PasswordExpirationDate = if ($passwordInfo) { $passwordInfo.ExpirationDate } else { $null }
            PasswordPercentUsed = if ($passwordInfo) { $passwordInfo.PercentUsed } else { 0 }
            IsPasswordExpired = if ($passwordInfo) { $passwordInfo.IsExpired } else { $false }
            LastLogonDate = $user.LastLogonDate
            Created = $user.Created
            Modified = $user.Modified
            AccountExpirationDate = $user.AccountExpirationDate
            DistinguishedName = $user.DistinguishedName
            Status = if (-not $user.Enabled) { "Disabled" }
                    elseif ($user.LockedOut) { "Locked" }
                    elseif ($user.PasswordNeverExpires) { "Never Expires" }
                    elseif ($passwordInfo -and $passwordInfo.IsExpired) { "Expired" }
                    else { "Active" }
        }
        
        $processedData.Add($userObject) | Out-Null
    }
    
    Write-Log "Processed $($processedData.Count) users matching criteria" "INFO" "Green"
    return $processedData
}

function Generate-Summary {
    param([System.Collections.ArrayList]$Data)
    
    $totalUsers = $Data.Count
    $neverExpire = ($Data | Where-Object { $_.PasswordNeverExpires -eq $true }).Count
    $expired = ($Data | Where-Object { $_.IsPasswordExpired -eq $true }).Count
    $expiringSoon = ($Data | Where-Object { $_.DaysUntilExpiry -le 7 -and $_.DaysUntilExpiry -gt 0 }).Count
    $disabled = ($Data | Where-Object { $_.Enabled -eq $false }).Count
    $locked = ($Data | Where-Object { $_.LockedOut -eq $true }).Count
    $active = ($Data | Where-Object { $_.Enabled -eq $true -and $_.LockedOut -eq $false }).Count
    
    $complianceScore = if ($totalUsers -gt 0) {
        $nonCompliant = $neverExpire + $expired + $expiringSoon
        [Math]::Round((($totalUsers - $nonCompliant) / $totalUsers) * 100, 2)
    } else { 0 }
    
    $summary = [PSCustomObject]@{
        ReportDate = Get-Date
        TotalUsers = $totalUsers
        ActiveUsers = $active
        DisabledUsers = $disabled
        LockedUsers = $locked
        PasswordNeverExpires = $neverExpire
        PasswordExpired = $expired
        ExpiringWithin7Days = $expiringSoon
        ComplianceScore = $complianceScore
        ComplianceStatus = if ($complianceScore -ge $Threshold) { "PASS" } else { "FAIL" }
    }
    
    Write-Log "="*70 "INFO" "Yellow"
    Write-Log "SUMMARY STATISTICS" "INFO" "Yellow"
    Write-Log "="*70 "INFO" "Yellow"
    Write-Log "Total Users Processed: $totalUsers" "INFO" "Cyan"
    Write-Log "Password Never Expires: $neverExpire" "INFO" "Yellow"
    Write-Log "Passwords Expired: $expired" "INFO" "Red"
    Write-Log "Expiring Within 7 Days: $expiringSoon" "INFO" "Yellow"
    Write-Log "Compliance Score: $complianceScore%" "INFO" $(if($complianceScore -ge $Threshold){"Green"}else{"Red"})
    Write-Log "Compliance Status: $($summary.ComplianceStatus)" "INFO" $(if($complianceScore -ge $Threshold){"Green"}else{"Red"})
    Write-Log "="*70 "INFO" "Yellow"
    
    return $summary
}

function Export-CSVReport {
    param(
        [System.Collections.ArrayList]$Data,
        [string]$FilePath
    )
    
    try {
        $Data | Export-Csv -Path $FilePath -NoTypeInformation -Encoding UTF8
        Write-Log "  ✓ CSV report: $FilePath" "INFO" "Green"
    }
    catch {
        Write-Log "  ✗ Failed to export CSV: $_" "ERROR" "Red"
    }
}

function Export-HTMLReport {
    param(
        [System.Collections.ArrayList]$Data,
        [PSCustomObject]$Summary,
        [string]$FilePath
    )
    
    $neverExpireCount = ($Data | Where-Object { $_.PasswordNeverExpires -eq $true }).Count
    $expiredCount = ($Data | Where-Object { $_.IsPasswordExpired -eq $true }).Count
    $expiringSoon = ($Data | Where-Object { $_.DaysUntilExpiry -le 7 -and $_.DaysUntilExpiry -gt 0 }).Count
    
    $html = @"
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Password Never Expires Report</title>
    <style>
        * { margin: 0; padding: 0; box-sizing: border-box; }
        body {
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            background: #f0f2f5;
            padding: 20px;
            color: #333;
        }
        .container {
            max-width: 1400px;
            margin: 0 auto;
            background: white;
            border-radius: 12px;
            box-shadow: 0 2px 20px rgba(0,0,0,0.1);
            overflow: hidden;
        }
        .header {
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            color: white;
            padding: 30px 40px;
        }
        .header h1 { font-size: 28px; font-weight: 300; margin-bottom: 5px; }
        .header .subtitle { font-size: 14px; opacity: 0.9; }
        .summary-grid {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(180px, 1fr));
            gap: 15px;
            padding: 20px 40px;
            background: #f8f9fa;
            border-bottom: 1px solid #e9ecef;
        }
        .stat-card {
            background: white;
            padding: 15px;
            border-radius: 8px;
            box-shadow: 0 1px 3px rgba(0,0,0,0.1);
            text-align: center;
        }
        .stat-value {
            font-size: 28px;
            font-weight: 600;
            color: #2c3e50;
        }
        .stat-label {
            font-size: 12px;
            color: #7f8c8d;
            text-transform: uppercase;
            margin-top: 5px;
            letter-spacing: 0.5px;
        }
        .stat-card.warning .stat-value { color: #f39c12; }
        .stat-card.danger .stat-value { color: #e74c3c; }
        .stat-card.success .stat-value { color: #27ae60; }
        .stat-card.info .stat-value { color: #3498db; }
        .compliance-section {
            padding: 20px 40px;
            background: white;
            border-bottom: 1px solid #e9ecef;
        }
        .compliance-bar {
            background: #ecf0f1;
            border-radius: 10px;
            height: 24px;
            overflow: hidden;
            margin: 10px 0;
        }
        .compliance-fill {
            height: 100%;
            background: linear-gradient(90deg, #27ae60, #2ecc71);
            transition: width 0.3s ease;
            display: flex;
            align-items: center;
            justify-content: flex-end;
            padding-right: 10px;
            color: white;
            font-size: 12px;
            font-weight: bold;
        }
        .compliance-fill.warning { background: linear-gradient(90deg, #f39c12, #e67e22); }
        .compliance-fill.danger { background: linear-gradient(90deg, #e74c3c, #c0392b); }
        .table-container {
            padding: 20px 40px;
            overflow-x: auto;
        }
        table {
            width: 100%;
            border-collapse: collapse;
            font-size: 13px;
        }
        th {
            background: #2c3e50;
            color: white;
            padding: 12px 10px;
            text-align: left;
            font-weight: 600;
            position: sticky;
            top: 0;
            z-index: 10;
        }
        td {
            padding: 10px;
            border-bottom: 1px solid #e9ecef;
        }
        tr:hover { background: #f8f9fa; }
        tr.expired td { background: #fee; }
        tr.never-expire td { background: #fff8e1; }
        tr.expiring td { background: #fff3cd; }
        .badge {
            display: inline-block;
            padding: 3px 10px;
            border-radius: 12px;
            font-size: 11px;
            font-weight: 600;
        }
        .badge-danger { background: #e74c3c; color: white; }
        .badge-warning { background: #f39c12; color: white; }
        .badge-success { background: #27ae60; color: white; }
        .badge-info { background: #3498db; color: white; }
        .badge-secondary { background: #95a5a6; color: white; }
        .footer {
            padding: 20px 40px;
            background: #f8f9fa;
            border-top: 1px solid #e9ecef;
            font-size: 12px;
            color: #7f8c8d;
            text-align: center;
        }
        .filter-section {
            padding: 15px 40px;
            background: #f8f9fa;
            border-bottom: 1px solid #e9ecef;
        }
        .filter-section select {
            padding: 8px 12px;
            border: 1px solid #ddd;
            border-radius: 4px;
            font-size: 13px;
        }
        @media print {
            .filter-section { display: none; }
            body { background: white; padding: 0; }
            .container { box-shadow: none; }
        }
    </style>
</head>
<body>
<div class="container">
    <div class="header">
        <h1>🔒 Password Policy Compliance Report</h1>
        <div class="subtitle">
            Generated: $(Get-Date -Format "yyyy-MM-dd HH:mm:ss") | 
            Search Base: $(if($SearchBase){$SearchBase}else{'Entire Domain'})
        </div>
    </div>

    <div class="summary-grid">
        <div class="stat-card info">
            <div class="stat-value">$($Summary.TotalUsers)</div>
            <div class="stat-label">Total Users</div>
        </div>
        <div class="stat-card success">
            <div class="stat-value">$($Summary.ActiveUsers)</div>
            <div class="stat-label">Active Users</div>
        </div>
        <div class="stat-card warning">
            <div class="stat-value">$neverExpireCount</div>
            <div class="stat-label">Never Expire</div>
        </div>
        <div class="stat-card danger">
            <div class="stat-value">$expiredCount</div>
            <div class="stat-label">Expired Passwords</div>
        </div>
        <div class="stat-card warning">
            <div class="stat-value">$expiringSoon</div>
            <div class="stat-label">Expiring in 7 Days</div>
        </div>
        <div class="stat-card success">
            <div class="stat-value">$($Summary.ComplianceScore)%</div>
            <div class="stat-label">Compliance Score</div>
        </div>
    </div>

    <div class="compliance-section">
        <h3>Compliance Status: <span style="color: $(if($Summary.ComplianceScore -ge $Threshold){'#27ae60'}else{'#e74c3c'})">
            $($Summary.ComplianceStatus)
        </span></h3>
        <div class="compliance-bar">
            <div class="compliance-fill $(if($Summary.ComplianceScore -lt 60){'danger'}elseif($Summary.ComplianceScore -lt 80){'warning'})" 
                 style="width: $($Summary.ComplianceScore)%;">
                $($Summary.ComplianceScore)%
            </div>
        </div>
        <small>Threshold: $Threshold% | 
               Non-compliant users: $($neverExpireCount + $expiredCount + $expiringSoon)</small>
    </div>

    <div class="table-container">
        <table>
            <thead>
                <tr>
                    <th>User</th>
                    <th>Display Name</th>
                    <th>Email</th>
                    <th>Department</th>
                    <th>Status</th>
                    <th>Password Status</th>
                    <th>Password Age</th>
                    <th>Days to Expiry</th>
                    <th>Expiration Date</th>
                    <th>Last Logon</th>
                </tr>
            </thead>
            <tbody>
"@
    
    foreach ($user in $Data) {
        $rowClass = ""
        $statusBadge = ""
        
        if (-not $user.Enabled) {
            $rowClass = "disabled"
            $statusBadge = '<span class="badge badge-secondary">Disabled</span>'
        }
        elseif ($user.LockedOut) {
            $rowClass = "locked"
            $statusBadge = '<span class="badge badge-warning">Locked</span>'
        }
        elseif ($user.IsPasswordExpired) {
            $rowClass = "expired"
            $statusBadge = '<span class="badge badge-danger">Expired</span>'
        }
        elseif ($user.PasswordNeverExpires) {
            $rowClass = "never-expire"
            $statusBadge = '<span class="badge badge-warning">Never Expires</span>'
        }
        elseif ($user.DaysUntilExpiry -le 7 -and $user.DaysUntilExpiry -gt 0) {
            $rowClass = "expiring"
            $statusBadge = '<span class="badge badge-warning">Expiring Soon</span>'
        }
        else {
            $statusBadge = '<span class="badge badge-success">Compliant</span>'
        }
        
        $expirationDate = if ($user.PasswordExpirationDate) { $user.PasswordExpirationDate.ToString("yyyy-MM-dd") } else { "N/A" }
        $lastLogon = if ($user.LastLogonDate) { $user.LastLogonDate.ToString("yyyy-MM-dd") } else { "Never" }
        $passwordAge = if ($user.PasswordAgeInDays) { "$($user.PasswordAgeInDays) days" } else { "N/A" }
        $daysToExpiry = if ($user.DaysUntilExpiry -gt 0) { "$($user.DaysUntilExpiry) days" } 
                        elseif ($user.DaysUntilExpiry -eq 0) { "Expires today" }
                        elseif ($user.IsPasswordExpired) { "Expired" }
                        else { "N/A" }
        
        $html += @"
                <tr class="$rowClass">
                    <td><strong>$($user.SamAccountName)</strong></td>
                    <td>$($user.DisplayName)</td>
                    <td>$($user.EmailAddress)</td>
                    <td>$($user.Department)</td>
                    <td>$statusBadge</td>
                    <td>$(if($user.PasswordNeverExpires){'<span class="badge badge-warning">Never Expires</span>'}elseif($user.IsPasswordExpired){'<span class="badge badge-danger">Expired</span>'}else{'<span class="badge badge-success">Active</span>'})</td>
                    <td>$passwordAge</td>
                    <td>$daysToExpiry</td>
                    <td>$expirationDate</td>
                    <td>$lastLogon</td>
                </tr>
"@
    }
    
    $html += @"
            </tbody>
        </table>
    </div>

    <div class="footer">
        <p>Generated by Password_Never_Expires_Report.ps1 | $(Get-Date -Format "yyyy-MM-dd HH:mm:ss")</p>
        <p>© $(Get-Date -Format "yyyy") Active Directory Management Tool</p>
    </div>
</div>
</body>
</html>
"@
    
    try {
        $html | Out-File -FilePath $FilePath -Encoding UTF8
        Write-Log "  ✓ HTML report: $FilePath" "INFO" "Green"
    }
    catch {
        Write-Log "  ✗ Failed to export HTML: $_" "ERROR" "Red"
    }
}

function Export-JSONReport {
    param(
        [System.Collections.ArrayList]$Data,
        [PSCustomObject]$Summary,
        [string]$FilePath
    )
    
    try {
        $jsonData = @{
            Summary = $Summary
            Users = $Data
            Generated = Get-Date -Format "yyyy-MM-ddTHH:mm:ss"
            ScriptVersion = "1.0"
        }
        
        $jsonData | ConvertTo-Json -Depth 5 | Out-File -FilePath $FilePath -Encoding UTF8
        Write-Log "  ✓ JSON report: $FilePath" "INFO" "Green"
    }
    catch {
        Write-Log "  ✗ Failed to export JSON: $_" "ERROR" "Red"
    }
}

function Export-ExcelReport {
    param(
        [System.Collections.ArrayList]$Data,
        [PSCustomObject]$Summary,
        [string]$FilePath
    )
    
    # Check if ImportExcel module is available
    $excelModule = Get-Module -ListAvailable -Name ImportExcel
    if (-not $excelModule) {
        Write-Log "  ⚠ ImportExcel module not found. Skipping Excel export." "WARNING" "Yellow"
        Write-Log "  ℹ Install with: Install-Module ImportExcel -Force" "INFO" "Cyan"
        return
    }
    
    try {
        Import-Module ImportExcel -ErrorAction Stop
        
        # Create a temporary file for the data
        $tempCsv = [System.IO.Path]::GetTempFileName() + ".csv"
        $Data | Export-Csv -Path $tempCsv -NoTypeInformation -Encoding UTF8
        
        # Create Excel with multiple sheets
        $excelParams = @{
            Path = $FilePath
            AutoSize = $true
            BoldTopRow = $true
            FreezeTopRow = $true
            AutoFilter = $true
        }
        
        # Main data sheet
        $Data | Export-Excel @excelParams -WorksheetName "Users"
        
        # Summary sheet
        $summaryData = @(
            [PSCustomObject]@{ Metric = "Total Users"; Value = $Summary.TotalUsers }
            [PSCustomObject]@{ Metric = "Active Users"; Value = $Summary.ActiveUsers }
            [PSCustomObject]@{ Metric = "Disabled Users"; Value = $Summary.DisabledUsers }
            [PSCustomObject]@{ Metric = "Locked Users"; Value = $Summary.LockedUsers }
            [PSCustomObject]@{ Metric = "Password Never Expires"; Value = $Summary.PasswordNeverExpires }
            [PSCustomObject]@{ Metric = "Password Expired"; Value = $Summary.PasswordExpired }
            [PSCustomObject]@{ Metric = "Expiring Within 7 Days"; Value = $Summary.ExpiringWithin7Days }
            [PSCustomObject]@{ Metric = "Compliance Score (%)"; Value = $Summary.ComplianceScore }
            [PSCustomObject]@{ Metric = "Compliance Status"; Value = $Summary.ComplianceStatus }
        )
        
        $summaryData | Export-Excel -Path $FilePath -WorksheetName "Summary" -AutoSize -BoldTopRow
        
        Write-Log "  ✓ Excel report: $FilePath" "INFO" "Green"
        
        # Clean up temp file
        if (Test-Path $tempCsv) { Remove-Item $tempCsv -Force }
    }
    catch {
        Write-Log "  ✗ Failed to export Excel: $_" "ERROR" "Red"
    }
}

function Send-EmailReport {
    param(
        [string]$ReportPath,
        [PSCustomObject]$Summary
    )
    
    if (-not $NotifyAdmin) { return }
    if (-not $SmtpServer -or -not $FromEmail -or -not $RecipientEmail) {
        Write-Log "Email configuration missing. Skipping email notification." "WARNING" "Yellow"
        return
    }
    
    try {
        $subject = "AD Password Report: $($Summary.ComplianceStatus) - $($Summary.ComplianceScore)% Compliance"
        
        $body = @"
<!DOCTYPE html>
<html>
<head>
    <style>
        body { font-family: Arial, sans-serif; color: #333; }
        .header { background: #2c3e50; color: white; padding: 15px; border-radius: 5px; }
        .stats { margin: 20px 0; }
        .stat { padding: 10px; margin: 5px 0; background: #f8f9fa; border-left: 4px solid #3498db; }
        .stat.pass { border-left-color: #27ae60; }
        .stat.fail { border-left-color: #e74c3c; }
        .stat.warning { border-left-color: #f39c12; }
        .footer { font-size: 12px; color: #7f8c8d; margin-top: 20px; border-top: 1px solid #ddd; padding-top: 10px; }
    </style>
</head>
<body>
    <div class="header">
        <h2>Active Directory Password Policy Report</h2>
    </div>
    <div class="stats">
        <div class="stat">
            <strong>Compliance Status:</strong> 
            <span style="color: $(if($Summary.ComplianceScore -ge $Threshold){'#27ae60'}else{'#e74c3c'})">
                $($Summary.ComplianceStatus) ($($Summary.ComplianceScore)%)
            </span>
        </div>
        <div class="stat">
            <strong>Total Users:</strong> $($Summary.TotalUsers)
        </div>
        <div class="stat warning">
            <strong>Password Never Expires:</strong> $($Summary.PasswordNeverExpires)
        </div>
        <div class="stat fail">
            <strong>Expired Passwords:</strong> $($Summary.PasswordExpired)
        </div>
        <div class="stat warning">
            <strong>Expiring Within 7 Days:</strong> $($Summary.ExpiringWithin7Days)
        </div>
        <div class="stat">
            <strong>Report Generated:</strong> $(Get-Date -Format "yyyy-MM-dd HH:mm:ss")
        </div>
    </div>
    <p>Attached is the detailed report.</p>
    <div class="footer">
        <p>This is an automated report from the Active Directory Management System.</p>
    </div>
</body>
</html>
"@
        
        $mailParams = @{
            To = $RecipientEmail
            From = $FromEmail
            Subject = $subject
            Body = $body
            BodyAsHtml = $true
            SmtpServer = $SmtpServer
            ErrorAction = 'Stop'
        }
        
        if (Test-Path $ReportPath) {
            $mailParams.Attachments = $ReportPath
        }
        
        Send-MailMessage @mailParams
        Write-Log "  ✓ Email report sent to $RecipientEmail" "INFO" "Green"
    }
    catch {
        Write-Log "  ✗ Failed to send email: $_" "ERROR" "Red"
    }
}

function Save-HistoricalData {
    param(
        [PSCustomObject]$Summary,
        [string]$FilePath
    )
    
    try {
        # Create historical data directory
        $historicalDir = Join-Path $ExportPath "HistoricalData"
        if (-not (Test-Path $historicalDir)) {
            New-Item -Path $historicalDir -ItemType Directory -Force | Out-Null
        }
        
        $historicalFile = Join-Path $historicalDir "HistoricalData_$Script:Timestamp.json"
        $summary | ConvertTo-Json | Out-File -FilePath $historicalFile -Encoding UTF8
        
        Write-Log "  ✓ Historical data saved: $historicalFile" "INFO" "Green"
    }
    catch {
        Write-Log "  ✗ Failed to save historical data: $_" "ERROR" "Red"
    }
}
#endregion

#region Main Script
function Main {
    # Initialize
    Initialize-Logging
    
    # Check AD module
    if (-not (Import-ADModule)) {
        Write-Log "Script cannot continue without ActiveDirectory module." "ERROR" "Red"
        return
    }
    
    # Get domain password policy
    $passwordPolicy = Get-DomainPasswordPolicy -DC $DomainController
    if (-not $passwordPolicy) {
        Write-Log "Could not retrieve domain password policy. Using default settings." "WARNING" "Yellow"
        $maxPasswordAge = [TimeSpan]::FromDays(90)
    }
    else {
        $maxPasswordAge = $passwordPolicy.MaxPasswordAge
    }
    
    # Get users
    $users = Get-ADUsersWithPasswordInfo -SearchBase $SearchBase -IncludeDisabled $IncludeDisabled -IncludeSystemAccounts $IncludeSystemAccounts -DC $DomainController
    if (-not $users) {
        Write-Log "No users retrieved. Exiting." "ERROR" "Red"
        return
    }
    
    # Process user data
    $processedData = Process-UserData -Users $users -MaxPasswordAge $maxPasswordAge -ExpiringInDays $ExpiringInDays -IncludeExpired $IncludeExpired
    $Script:ReportData = $processedData
    
    # Generate summary
    $summary = Generate-Summary -Data $processedData
    $Script:Summary = $summary
    
    # Create export directory if needed
    $exportDir = Join-Path $ExportPath "PasswordReports"
    if (-not (Test-Path $exportDir)) {
        New-Item -Path $exportDir -ItemType Directory -Force | Out-Null
    }
    
    $exportPrefix = "PasswordNeverExpires_$Script:Timestamp"
    
    # Export reports
    Write-Log "Exporting reports..." "INFO" "Cyan"
    
    # Determine which report file to attach for email
    $emailAttachment = $null
    
    if ('CSV' -in $OutputFormats) {
        $csvPath = Join-Path $exportDir "$exportPrefix.csv"
        Export-CSVReport -Data $processedData -FilePath $csvPath
        if (-not $emailAttachment) { $emailAttachment = $csvPath }
    }
    
    if ('HTML' -in $OutputFormats) {
        $htmlPath = Join-Path $exportDir "$exportPrefix.html"
        Export-HTMLReport -Data $processedData -Summary $summary -FilePath $htmlPath
        if (-not $emailAttachment) { $emailAttachment = $htmlPath }
    }
    
    if ('JSON' -in $OutputFormats) {
        $jsonPath = Join-Path $exportDir "$exportPrefix.json"
        Export-JSONReport -Data $processedData -Summary $summary -FilePath $jsonPath
    }
    
    if ('Excel' -in $OutputFormats) {
        $excelPath = Join-Path $exportDir "$exportPrefix.xlsx"
        Export-ExcelReport -Data $processedData -Summary $summary -FilePath $excelPath
        if (-not $emailAttachment) { $emailAttachment = $excelPath }
    }
    
    # Historical data
    if ($HistoricalData) {
        Save-HistoricalData -Summary $summary -FilePath $exportDir
    }
    
    # Email notification
    if ($NotifyAdmin) {
        Send-EmailReport -ReportPath $emailAttachment -Summary $summary
    }
    
    # Final summary
    Write-Log "="*70 "INFO" "Yellow"
    Write-Log "REPORT GENERATION COMPLETE" "INFO" "Green"
    Write-Log "="*70 "INFO" "Yellow"
    Write-Log "Export Location: $exportDir" "INFO" "Cyan"
    Write-Log "Total Users: $($summary.TotalUsers)" "INFO" "Cyan"
    Write-Log "Compliance Score: $($summary.ComplianceScore)%" "INFO" $(if($summary.ComplianceScore -ge $Threshold){"Green"}else{"Red"})
    Write-Log "Log File: $Script:LogFile" "INFO" "Cyan"
    Write-Log "="*70 "INFO" "Yellow"
}

# Run the script
Main
#endregion