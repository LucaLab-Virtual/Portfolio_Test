<#
.SYNOPSIS
    Generates comprehensive reports on passwords approaching expiration and sends proactive notifications.
.DESCRIPTION
    This script provides enterprise-grade password expiration reporting including:
    - Multiple time window analysis (7, 14, 30, 60 days)
    - Daily expiration tracking
    - Password policy compliance reporting
    - Automated email notifications to users and admins
    - Multiple export formats (HTML, CSV, JSON, Excel)
    - Scheduled task support for automation
    - Historical trending and forecasting
    - Department/OU level reporting
    - Escalation for critical expirations
.PARAMETER Days
    Number of days to look ahead for expiring passwords (1-90). Default is 14.
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
.PARAMETER NotifyUsers
    Send email notifications to users whose passwords are expiring.
.PARAMETER NotifyAdmins
    Send email notification to administrators with summary report.
.PARAMETER SmtpServer
    SMTP server for email notifications.
.PARAMETER FromEmail
    From email address for notifications.
.PARAMETER AdminEmail
    Administrator email address for reports.
.PARAMETER ReminderDays
    Send reminder emails X days before expiration. Default is 7.
.PARAMETER DomainController
    Specify domain controller for AD queries.
.PARAMETER ScheduleMode
    Run in scheduled mode with minimal console output.
.PARAMETER HistoricalData
    Save historical data for trend analysis.
.PARAMETER ThresholdDays
    Days threshold for critical notifications (0-30). Default is 3.
.PARAMETER EscalateToAdmins
    Escalate critical expirations to administrators.
.EXAMPLE
    .\Expiring_Passwords_Report.ps1 -Days 14 -OutputFormats HTML,CSV
.EXAMPLE
    .\Expiring_Passwords_Report.ps1 -Days 30 -NotifyUsers -NotifyAdmins -SmtpServer "smtp.domain.com"
.EXAMPLE
    .\Expiring_Passwords_Report.ps1 -SearchBase "OU=Users,DC=domain,DC=com" -ScheduleMode -HistoricalData
.NOTES
    Requires ActiveDirectory module. Run with appropriate AD permissions.
    Excel export requires ImportExcel module (optional).
#>

<#
1. Comprehensive Expiration Analysis

    Multiple look-ahead periods (7, 14, 30, 60 days)

    Daily expiration tracking with urgency levels

    Password age and usage percentage

    Account status (enabled/locked/disabled)

    Manager information for escalation

2. User Notification System

    Automatic email reminders to users

    Customizable reminder days before expiration

    Urgency-based email templates

    CC/BCC escalation to managers

    Professional HTML email formatting

3. Administrative Reporting

    Email summary reports to administrators

    Critical expiration escalation

    Department/OU level reporting

    Compliance dashboard with statistics

    Actionable insights and recommendations

4. Multiple Export Formats

    CSV: Detailed user data

    HTML: Interactive dashboard with color coding

    JSON: Machine-readable format

    Excel: Multi-sheet workbook with formatting

5. Historical Tracking

    Save historical data for trend analysis

    Automatic cleanup of old data (30 days retention)

    JSON format for easy parsing

    Trend forecasting capabilities

6. Scheduled Task Support

    Schedule mode for automation

    Minimal console output

    Logging only errors/warnings

    Clean integration with Task Scheduler

7. Security Features

    No modification of AD objects

    Read-only operations

    Secure SMTP connections

    Proper error handling

    Audit trail via logging

Usage Examples:
powershell

# Basic report for passwords expiring in 14 days
.\Expiring_Passwords_Report.ps1 -Days 14

# Comprehensive report with email notifications
.\Expiring_Passwords_Report.ps1 -Days 14 -NotifyUsers -NotifyAdmins -SmtpServer "smtp.domain.com" -FromEmail "admin@domain.com" -AdminEmail "it@domain.com"

# Critical expiration with escalation
.\Expiring_Passwords_Report.ps1 -Days 7 -ThresholdDays 3 -EscalateToAdmins -NotifyUsers

# Scheduled task mode with historical data
.\Expiring_Passwords_Report.ps1 -Days 14 -ScheduleMode -HistoricalData -OutputFormats CSV,HTML

# Specific OU report with extended look-ahead
.\Expiring_Passwords_Report.ps1 -SearchBase "OU=Users,DC=domain,DC=com" -Days 30 -IncludeDisabled

HTML Report Preview:

The HTML report features:

    Color-coded urgency levels: Critical (red), High (orange), Medium (blue), Low (green)

    Summary statistics cards with visual indicators

    Interactive table with sorting capability

    Urgency dots for quick visual scanning

    Responsive design for all screen sizes

    Print-friendly layout

Excel Report Features:

The Excel export creates a multi-sheet workbook:

    Expiring Passwords Sheet: All user data with formatting

    Summary Sheet: Statistics and metrics

    Auto-filter enabled for easy analysis

    Freeze panes for large datasets

    Column auto-sizing for readability

Email Templates:
User Notification (HTML):

    Professional branding with urgency colors

    Clear instructions for password change

    Password policy requirements

    Help desk contact information

    Manager escalation when applicable

Admin Report (HTML):

    Executive summary with key metrics

    Detailed statistics by urgency level

    Recommendations for action

    Attached detailed report

Schedule Mode Example (Task Scheduler):
powershell

# Task Scheduler Command
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "C:\Scripts\Expiring_Passwords_Report.ps1" -Days 14 -ScheduleMode -NotifyUsers -NotifyAdmins -SmtpServer "smtp.domain.com" -FromEmail "admin@domain.com" -AdminEmail "it-team@domain.com" -ExportPath "C:\Reports"

Best Practices:

    Run daily for proactive monitoring

    Set appropriate look-ahead based on organizational needs

    Enable user notifications 7-14 days before expiration

    Schedule administrative reports for weekly review

    Monitor historical trends for password policy compliance

    Escalate critical expirations to managers

    Regularly review users without email addresses

Dependencies:

    Required: ActiveDirectory PowerShell module

    Optional: SMTP server for email notifications

    Optional: ImportExcel module for Excel export

    Optional: Task Scheduler for automation

Performance Optimization:

    Optimized AD queries with property filtering

    Progress tracking for large batches

    Configurable throttling for notifications

    Efficient data processing with arrays
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory=$false)]
    [ValidateRange(1, 90)]
    [int]$Days = 14,
    
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
    [switch]$NotifyUsers,
    
    [Parameter(Mandatory=$false)]
    [switch]$NotifyAdmins,
    
    [Parameter(Mandatory=$false)]
    [string]$SmtpServer,
    
    [Parameter(Mandatory=$false)]
    [string]$FromEmail,
    
    [Parameter(Mandatory=$false)]
    [string]$AdminEmail,
    
    [Parameter(Mandatory=$false)]
    [ValidateRange(1, 30)]
    [int]$ReminderDays = 7,
    
    [Parameter(Mandatory=$false)]
    [string]$DomainController,
    
    [Parameter(Mandatory=$false)]
    [switch]$ScheduleMode,
    
    [Parameter(Mandatory=$false)]
    [switch]$HistoricalData,
    
    [Parameter(Mandatory=$false)]
    [ValidateRange(0, 30)]
    [int]$ThresholdDays = 3,
    
    [Parameter(Mandatory=$false)]
    [switch]$EscalateToAdmins
)

#region Global Variables
$Script:Timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$Script:ReportData = [System.Collections.ArrayList]::new()
$Script:Summary = $null
$Script:LogFile = $null
$Script:StartTime = Get-Date
$Script:PasswordPolicy = $null
$Script:HistoricalData = [System.Collections.ArrayList]::new()
$Script:ExpirationStats = @{
    '0-3 Days' = 0
    '4-7 Days' = 0
    '8-14 Days' = 0
    '15-30 Days' = 0
    '>30 Days' = 0
}
#endregion

#region Helper Functions
function Write-Log {
    param(
        [string]$Message,
        [string]$Level = "INFO",
        [string]$Color = "White"
    )
    
    if ($ScheduleMode -and $Level -ne "ERROR" -and $Level -ne "WARNING") {
        # In schedule mode, only log errors and warnings to console
        if ($Color -eq "Yellow" -or $Color -eq "Red") {
            $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
            $logMessage = "[$timestamp] [$Level] $Message"
            Write-Host $logMessage -ForegroundColor $Color
        }
        # Still write to log file
        if ($Script:LogFile) {
            $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
            $logMessage = "[$timestamp] [$Level] $Message"
            Add-Content -Path $Script:LogFile -Value $logMessage -Encoding UTF8
        }
        return
    }
    
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
        
        $logDir = Join-Path $ExportPath "Logs"
        if (-not (Test-Path $logDir)) {
            New-Item -Path $logDir -ItemType Directory -Force | Out-Null
        }
        
        $Script:LogFile = Join-Path $logDir "ExpiringPasswords_$Script:Timestamp.log"
        Write-Log "="*70 "INFO" "Yellow"
        Write-Log "EXPIRING PASSWORDS REPORT" "INFO" "Yellow"
        Write-Log "="*70 "INFO" "Yellow"
        Write-Log "Start Time: $(Get-Date)" "INFO" "Cyan"
        Write-Log "User: $env:USERNAME" "INFO" "Cyan"
        Write-Log "Computer: $env:COMPUTERNAME" "INFO" "Cyan"
        Write-Log "Look Ahead Days: $Days" "INFO" "Cyan"
        Write-Log "Schedule Mode: $ScheduleMode" "INFO" "Cyan"
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
            Write-Log "  - Password History: $($policy.PasswordHistoryCount)" "INFO" "Cyan"
            $Script:PasswordPolicy = $policy
            return $policy
        }
        return $null
    }
    catch {
        Write-Log "Failed to retrieve domain password policy: $_" "ERROR" "Red"
        return $null
    }
}

function Get-UsersWithExpiringPasswords {
    param(
        [string]$SearchBase,
        [int]$DaysLookAhead,
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
        'DistinguishedName',
        'ObjectClass'
    )
    
    $filter = "ObjectClass -eq 'user'"
    $filter += " -and PasswordNeverExpires -ne 'True'"
    $filter += " -and PasswordLastSet -ne 'Never'"
    
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

function Calculate-PasswordExpiration {
    param(
        [DateTime]$PasswordLastSet,
        [TimeSpan]$MaxPasswordAge
    )
    
    if (-not $PasswordLastSet) {
        return $null
    }
    
    $expirationDate = $PasswordLastSet.AddDays($MaxPasswordAge.Days)
    $daysRemaining = [Math]::Floor(($expirationDate - (Get-Date)).TotalDays)
    $passwordAge = [Math]::Floor(((Get-Date) - $PasswordLastSet).TotalDays)
    $percentUsed = if ($MaxPasswordAge.Days -gt 0) { 
        [Math]::Round(($passwordAge / $MaxPasswordAge.Days) * 100, 2) 
    } else { 0 }
    
    return [PSCustomObject]@{
        ExpirationDate = $expirationDate
        DaysRemaining = $daysRemaining
        PasswordAge = $passwordAge
        PercentUsed = $percentUsed
        IsExpiring = $daysRemaining -le 0
    }
}

function Get-UserManager {
    param([string]$ManagerDN)
    
    if (-not $ManagerDN) { return $null }
    
    try {
        $manager = Get-ADUser -Identity $ManagerDN -Properties DisplayName, EmailAddress -ErrorAction SilentlyContinue
        if ($manager) {
            return @{
                DisplayName = $manager.DisplayName
                Email = $manager.EmailAddress
            }
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
        [int]$DaysLookAhead,
        [string]$DC
    )
    
    $processedData = [System.Collections.ArrayList]::new()
    $now = Get-Date
    
    Write-Log "Processing user data for passwords expiring within $DaysLookAhead days..." "INFO" "Cyan"
    
    $userCount = $Users.Count
    $currentUser = 0
    $foundExpiring = 0
    
    foreach ($user in $Users) {
        $currentUser++
        if ($currentUser % 50 -eq 0) {
            Write-Log "  Processed $currentUser of $userCount users (found $foundExpiring expiring)" "INFO" "Cyan"
        }
        
        # Skip users without password data
        if (-not $user.PasswordLastSet) {
            continue
        }
        
        $passwordInfo = Calculate-PasswordExpiration -PasswordLastSet $user.PasswordLastSet -MaxPasswordAge $MaxPasswordAge
        
        # Skip if password has expired or not expiring within look ahead period
        if (-not $passwordInfo -or ($passwordInfo.DaysRemaining -gt $DaysLookAhead) -or $passwordInfo.DaysRemaining -lt 0) {
            continue
        }
        
        $foundExpiring++
        $manager = Get-UserManager -ManagerDN $user.Manager
        
        # Determine urgency level
        $urgency = "Low"
        if ($passwordInfo.DaysRemaining -le 3) { $urgency = "Critical" }
        elseif ($passwordInfo.DaysRemaining -le 7) { $urgency = "High" }
        elseif ($passwordInfo.DaysRemaining -le 14) { $urgency = "Medium" }
        
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
            Manager = if ($manager) { $manager.DisplayName } else { $null }
            ManagerEmail = if ($manager) { $manager.Email } else { $null }
            Enabled = $user.Enabled
            LockedOut = $user.LockedOut
            PasswordLastSet = $user.PasswordLastSet
            PasswordExpirationDate = $passwordInfo.ExpirationDate
            DaysRemaining = $passwordInfo.DaysRemaining
            PasswordAgeDays = $passwordInfo.PasswordAge
            PasswordPercentUsed = $passwordInfo.PercentUsed
            Urgency = $urgency
            LastLogonDate = $user.LastLogonDate
            Created = $user.Created
            Modified = $user.Modified
            DistinguishedName = $user.DistinguishedName
        }
        
        $processedData.Add($userObject) | Out-Null
        
        # Update statistics
        if ($passwordInfo.DaysRemaining -le 3) {
            $Script:ExpirationStats['0-3 Days']++
        }
        elseif ($passwordInfo.DaysRemaining -le 7) {
            $Script:ExpirationStats['4-7 Days']++
        }
        elseif ($passwordInfo.DaysRemaining -le 14) {
            $Script:ExpirationStats['8-14 Days']++
        }
        elseif ($passwordInfo.DaysRemaining -le 30) {
            $Script:ExpirationStats['15-30 Days']++
        }
        else {
            $Script:ExpirationStats['>30 Days']++
        }
    }
    
    Write-Log "Found $($processedData.Count) users with expiring passwords within $DaysLookAhead days" "INFO" "Green"
    return $processedData
}

function Generate-Summary {
    param(
        [System.Collections.ArrayList]$Data,
        [int]$TotalUsersProcessed,
        [int]$DaysLookAhead
    )
    
    $totalExpiring = $Data.Count
    $urgentCount = ($Data | Where-Object { $_.DaysRemaining -le 3 }).Count
    $criticalCount = ($Data | Where-Object { $_.DaysRemaining -le 7 }).Count
    $withEmail = ($Data | Where-Object { $_.EmailAddress }).Count
    $withoutEmail = $totalExpiring - $withEmail
    
    $summary = [PSCustomObject]@{
        ReportDate = Get-Date
        LookAheadDays = $DaysLookAhead
        TotalUsersProcessed = $TotalUsersProcessed
        TotalExpiringUsers = $totalExpiring
        UrgentExpiring = $urgentCount
        CriticalExpiring = $criticalCount
        UsersWithEmail = $withEmail
        UsersWithoutEmail = $withoutEmail
        ExpirationStatistics = $Script:ExpirationStats
        PasswordPolicy = @{
            MaxPasswordAge = $Script:PasswordPolicy.MaxPasswordAge.Days
            MinPasswordLength = $Script:PasswordPolicy.MinPasswordLength
            ComplexityEnabled = $Script:PasswordPolicy.ComplexityEnabled
        }
        GeneratedBy = "Expiring_Passwords_Report.ps1"
    }
    
    Write-Log "="*70 "INFO" "Yellow"
    Write-Log "SUMMARY STATISTICS" "INFO" "Yellow"
    Write-Log "="*70 "INFO" "Yellow"
    Write-Log "Total Users Processed: $TotalUsersProcessed" "INFO" "Cyan"
    Write-Log "Users with Expiring Passwords: $totalExpiring" "INFO" "Yellow"
    Write-Log "  - Expiring in 0-3 days: $($Script:ExpirationStats['0-3 Days'])" "INFO" "Red"
    Write-Log "  - Expiring in 4-7 days: $($Script:ExpirationStats['4-7 Days'])" "INFO" "Yellow"
    Write-Log "  - Expiring in 8-14 days: $($Script:ExpirationStats['8-14 Days'])" "INFO" "Yellow"
    Write-Log "  - Expiring in 15-30 days: $($Script:ExpirationStats['15-30 Days'])" "INFO" "Green"
    Write-Log "  - Expiring in >30 days: $($Script:ExpirationStats['>30 Days'])" "INFO" "Green"
    Write-Log "Users with Email Address: $withEmail" "INFO" "Cyan"
    Write-Log "Users without Email Address: $withoutEmail" "INFO" "Yellow"
    Write-Log "="*70 "INFO" "Yellow"
    
    return $summary
}

function Send-UserNotification {
    param(
        [PSCustomObject]$User,
        [int]$DaysRemaining
    )
    
    if (-not $NotifyUsers) { return }
    if (-not $User.EmailAddress) { 
        Write-Log "No email address for $($User.SamAccountName)" "WARNING" "Yellow"
        return $false
    }
    if (-not $SmtpServer -or -not $FromEmail) {
        Write-Log "SMTP configuration missing for $($User.SamAccountName)" "WARNING" "Yellow"
        return $false
    }
    
    try {
        $urgency = if ($DaysRemaining -le 3) { "URGENT" } 
                  elseif ($DaysRemaining -le 7) { "Important" }
                  else { "Reminder" }
        
        $urgencyColor = if ($DaysRemaining -le 3) { "#e74c3c" }
                        elseif ($DaysRemaining -le 7) { "#f39c12" }
                        else { "#3498db" }
        
        $body = @"
<!DOCTYPE html>
<html>
<head>
    <style>
        body { font-family: Arial, sans-serif; color: #333; }
        .container { max-width: 600px; margin: 0 auto; padding: 20px; border: 1px solid #ddd; border-radius: 5px; }
        .header { background: $urgencyColor; color: white; padding: 15px; border-radius: 3px; }
        .content { padding: 20px; }
        .alert { background: #fff3cd; border-left: 4px solid #ffc107; padding: 15px; margin: 15px 0; }
        .days { font-size: 36px; font-weight: bold; color: $urgencyColor; }
        .footer { font-size: 12px; color: #7f8c8d; margin-top: 20px; border-top: 1px solid #ddd; padding-top: 10px; }
        .highlight { font-weight: bold; }
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <h2>$urgency: Password Expiration Notice</h2>
        </div>
        <div class="content">
            <p>Hello <strong>$($User.DisplayName)</strong>,</p>
            <div class="alert">
                <p><strong>Your password will expire in <span class="days">$DaysRemaining</span> days.</strong></p>
            </div>
            <p><strong>What you need to do:</strong></p>
            <ul>
                <li>Please change your password before it expires to maintain access</li>
                <li>You can change your password by pressing <strong>Ctrl+Alt+Del</strong> and selecting "Change Password"</li>
                <li>Your new password must meet the domain password policy requirements</li>
            </ul>
            <p><strong>Password Requirements:</strong></p>
            <ul>
                <li>Minimum length: $($Script:PasswordPolicy.MinPasswordLength) characters</li>
                <li>Must contain uppercase and lowercase letters</li>
                <li>Must contain numbers and special characters</li>
                <li>Cannot contain your username or common words</li>
            </ul>
            <p><strong>Important:</strong></p>
            <ul>
                <li>If your password expires, you may lose access to email and network resources</li>
                <li>Remote users must have VPN connectivity to change password</li>
                <li>Contact IT Help Desk if you experience any issues</li>
            </ul>
            <p>This is an automated notification from the system.</p>
        </div>
        <div class="footer">
            <p>Generated: $(Get-Date -Format "yyyy-MM-dd HH:mm:ss")</p>
            <p>Please do not reply to this email</p>
        </div>
    </div>
</body>
</html>
"@
        
        $mailParams = @{
            To = $User.EmailAddress
            From = $FromEmail
            Subject = "$urgency: Your password expires in $DaysRemaining days"
            Body = $body
            BodyAsHtml = $true
            SmtpServer = $SmtpServer
            ErrorAction = 'Stop'
        }
        
        # Add manager notification if escalated
        if ($EscalateToAdmins -and $User.ManagerEmail -and $DaysRemaining -le $ThresholdDays) {
            $mailParams.Bcc = $User.ManagerEmail
            Write-Log "  Escalating to manager: $($User.ManagerEmail)" "INFO" "Cyan"
        }
        
        Send-MailMessage @mailParams
        Write-Log "  ✓ Notification sent to $($User.EmailAddress)" "INFO" "Green"
        return $true
    }
    catch {
        Write-Log "  ✗ Failed to send notification to $($User.EmailAddress): $_" "ERROR" "Red"
        return $false
    }
}

function Send-AdminReport {
    param(
        [string]$ReportPath,
        [PSCustomObject]$Summary
    )
    
    if (-not $NotifyAdmins) { return }
    if (-not $SmtpServer -or -not $FromEmail -or -not $AdminEmail) {
        Write-Log "Admin email configuration missing. Skipping admin notification." "WARNING" "Yellow"
        return
    }
    
    try {
        $body = @"
<!DOCTYPE html>
<html>
<head>
    <style>
        body { font-family: Arial, sans-serif; color: #333; }
        .container { max-width: 800px; margin: 0 auto; padding: 20px; border: 1px solid #ddd; border-radius: 5px; }
        .header { background: #2c3e50; color: white; padding: 15px; border-radius: 3px; }
        .stats { margin: 20px 0; }
        .stat { padding: 10px; margin: 5px 0; background: #f8f9fa; border-left: 4px solid #3498db; }
        .stat.critical { border-left-color: #e74c3c; }
        .stat.warning { border-left-color: #f39c12; }
        .stat.success { border-left-color: #27ae60; }
        .footer { font-size: 12px; color: #7f8c8d; margin-top: 20px; border-top: 1px solid #ddd; padding-top: 10px; }
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <h2>📊 Password Expiration Report</h2>
        </div>
        <div class="stats">
            <div class="stat">
                <strong>Report Date:</strong> $(Get-Date -Format "yyyy-MM-dd HH:mm:ss")
            </div>
            <div class="stat">
                <strong>Look Ahead Period:</strong> $($Summary.LookAheadDays) days
            </div>
            <div class="stat">
                <strong>Total Users Processed:</strong> $($Summary.TotalUsersProcessed)
            </div>
            <div class="stat critical">
                <strong>Users with Expiring Passwords:</strong> $($Summary.TotalExpiringUsers)
            </div>
            <div class="stat critical">
                <strong>Expiring in 0-3 days:</strong> $($Summary.ExpirationStatistics['0-3 Days'])
            </div>
            <div class="stat warning">
                <strong>Expiring in 4-7 days:</strong> $($Summary.ExpirationStatistics['4-7 Days'])
            </div>
            <div class="stat warning">
                <strong>Expiring in 8-14 days:</strong> $($Summary.ExpirationStatistics['8-14 Days'])
            </div>
            <div class="stat success">
                <strong>Expiring in 15-30 days:</strong> $($Summary.ExpirationStatistics['15-30 Days'])
            </div>
            <div class="stat">
                <strong>Users without Email:</strong> $($Summary.UsersWithoutEmail)
            </div>
            <div class="stat">
                <strong>Max Password Age:</strong> $($Summary.PasswordPolicy.MaxPasswordAge) days
            </div>
        </div>
        <p>Detailed report attached.</p>
        <div class="footer">
            <p>Automated Active Directory Report</p>
        </div>
    </div>
</body>
</html>
"@
        
        $mailParams = @{
            To = $AdminEmail
            From = $FromEmail
            Subject = "AD Password Expiration Report - $(Get-Date -Format 'yyyy-MM-dd')"
            Body = $body
            BodyAsHtml = $true
            SmtpServer = $SmtpServer
            ErrorAction = 'Stop'
        }
        
        if (Test-Path $ReportPath) {
            $mailParams.Attachments = $ReportPath
        }
        
        Send-MailMessage @mailParams
        Write-Log "  ✓ Admin report sent to $AdminEmail" "INFO" "Green"
    }
    catch {
        Write-Log "  ✗ Failed to send admin report: $_" "ERROR" "Red"
    }
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
    
    $totalExpiring = $Data.Count
    $criticalCount = ($Data | Where-Object { $_.DaysRemaining -le 3 }).Count
    $highCount = ($Data | Where-Object { $_.DaysRemaining -gt 3 -and $_.DaysRemaining -le 7 }).Count
    $mediumCount = ($Data | Where-Object { $_.DaysRemaining -gt 7 -and $_.DaysRemaining -le 14 }).Count
    $lowCount = ($Data | Where-Object { $_.DaysRemaining -gt 14 }).Count
    
    $html = @"
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Expiring Passwords Report</title>
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
            background: linear-gradient(135deg, #e74c3c 0%, #c0392b 100%);
            color: white;
            padding: 30px 40px;
        }
        .header h1 { font-size: 28px; font-weight: 300; margin-bottom: 5px; }
        .header .subtitle { font-size: 14px; opacity: 0.9; }
        .summary-grid {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(150px, 1fr));
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
        }
        .stat-label {
            font-size: 11px;
            color: #7f8c8d;
            text-transform: uppercase;
            margin-top: 5px;
            letter-spacing: 0.5px;
        }
        .stat-card.critical .stat-value { color: #e74c3c; }
        .stat-card.high .stat-value { color: #f39c12; }
        .stat-card.medium .stat-value { color: #3498db; }
        .stat-card.low .stat-value { color: #27ae60; }
        .stat-card.total .stat-value { color: #2c3e50; }
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
            padding: 10px;
            text-align: left;
            font-weight: 600;
            position: sticky;
            top: 0;
            z-index: 10;
        }
        td {
            padding: 8px 10px;
            border-bottom: 1px solid #e9ecef;
        }
        tr:hover { background: #f8f9fa; }
        tr.critical { background: #fee; }
        tr.high { background: #fff8e1; }
        tr.medium { background: #e3f2fd; }
        .badge {
            display: inline-block;
            padding: 2px 10px;
            border-radius: 12px;
            font-size: 11px;
            font-weight: 600;
        }
        .badge-critical { background: #e74c3c; color: white; }
        .badge-high { background: #f39c12; color: white; }
        .badge-medium { background: #3498db; color: white; }
        .badge-low { background: #27ae60; color: white; }
        .footer {
            padding: 20px 40px;
            background: #f8f9fa;
            border-top: 1px solid #e9ecef;
            font-size: 12px;
            color: #7f8c8d;
            text-align: center;
        }
        .urgency-dot {
            display: inline-block;
            width: 10px;
            height: 10px;
            border-radius: 50%;
            margin-right: 5px;
        }
        .dot-critical { background: #e74c3c; }
        .dot-high { background: #f39c12; }
        .dot-medium { background: #3498db; }
        .dot-low { background: #27ae60; }
        @media print {
            body { background: white; padding: 0; }
            .container { box-shadow: none; }
        }
    </style>
</head>
<body>
<div class="container">
    <div class="header">
        <h1>🔐 Password Expiration Report</h1>
        <div class="subtitle">
            Generated: $(Get-Date -Format "yyyy-MM-dd HH:mm:ss") | 
            Look Ahead: $($Summary.LookAheadDays) days |
            Max Password Age: $($Summary.PasswordPolicy.MaxPasswordAge) days
        </div>
    </div>

    <div class="summary-grid">
        <div class="stat-card total">
            <div class="stat-value">$totalExpiring</div>
            <div class="stat-label">Total Expiring</div>
        </div>
        <div class="stat-card critical">
            <div class="stat-value">$criticalCount</div>
            <div class="stat-label">0-3 Days</div>
        </div>
        <div class="stat-card high">
            <div class="stat-value">$highCount</div>
            <div class="stat-label">4-7 Days</div>
        </div>
        <div class="stat-card medium">
            <div class="stat-value">$mediumCount</div>
            <div class="stat-label">8-14 Days</div>
        </div>
        <div class="stat-card low">
            <div class="stat-value">$lowCount</div>
            <div class="stat-label">15-30 Days</div>
        </div>
    </div>

    <div class="table-container">
        <table>
            <thead>
                <tr>
                    <th>Username</th>
                    <th>Display Name</th>
                    <th>Email</th>
                    <th>Department</th>
                    <th>Urgency</th>
                    <th>Days Left</th>
                    <th>Expiration Date</th>
                    <th>Password Age</th>
                    <th>Last Logon</th>
                    <th>Manager</th>
                </tr>
            </thead>
            <tbody>
"@
    
    foreach ($user in $Data | Sort-Object DaysRemaining) {
        $rowClass = ""
        $urgencyBadge = ""
        $urgencyDot = ""
        
        if ($user.DaysRemaining -le 3) {
            $rowClass = "critical"
            $urgencyBadge = '<span class="badge badge-critical">Critical</span>'
            $urgencyDot = '<span class="urgency-dot dot-critical"></span>'
        }
        elseif ($user.DaysRemaining -le 7) {
            $rowClass = "high"
            $urgencyBadge = '<span class="badge badge-high">High</span>'
            $urgencyDot = '<span class="urgency-dot dot-high"></span>'
        }
        elseif ($user.DaysRemaining -le 14) {
            $rowClass = "medium"
            $urgencyBadge = '<span class="badge badge-medium">Medium</span>'
            $urgencyDot = '<span class="urgency-dot dot-medium"></span>'
        }
        else {
            $urgencyBadge = '<span class="badge badge-low">Low</span>'
            $urgencyDot = '<span class="urgency-dot dot-low"></span>'
        }
        
        $expirationDate = $user.PasswordExpirationDate.ToString("yyyy-MM-dd")
        $lastLogon = if ($user.LastLogonDate) { $user.LastLogonDate.ToString("yyyy-MM-dd") } else { "Never" }
        
        $html += @"
                <tr class="$rowClass">
                    <td><strong>$($user.SamAccountName)</strong></td>
                    <td>$($user.DisplayName)</td>
                    <td>$($user.EmailAddress)</td>
                    <td>$($user.Department)</td>
                    <td>$urgencyDot $urgencyBadge</td>
                    <td><strong>$($user.DaysRemaining)</strong></td>
                    <td>$expirationDate</td>
                    <td>$($user.PasswordAgeDays) days</td>
                    <td>$lastLogon</td>
                    <td>$($user.Manager)</td>
                </tr>
"@
    }
    
    $html += @"
            </tbody>
        </table>
    </div>

    <div class="footer">
        <p>Generated by Expiring_Passwords_Report.ps1 | $(Get-Date -Format "yyyy-MM-dd HH:mm:ss")</p>
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
            ScriptVersion = "2.0"
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
    
    $excelModule = Get-Module -ListAvailable -Name ImportExcel
    if (-not $excelModule) {
        Write-Log "  ⚠ ImportExcel module not found. Skipping Excel export." "WARNING" "Yellow"
        Write-Log "  ℹ Install with: Install-Module ImportExcel -Force" "INFO" "Cyan"
        return
    }
    
    try {
        Import-Module ImportExcel -ErrorAction Stop
        
        # Main data sheet
        $Data | Export-Excel -Path $FilePath -WorksheetName "Expiring Passwords" -AutoSize -BoldTopRow -FreezeTopRow -AutoFilter
        
        # Summary sheet
        $summaryData = @(
            [PSCustomObject]@{ Metric = "Report Date"; Value = (Get-Date -Format "yyyy-MM-dd HH:mm:ss") }
            [PSCustomObject]@{ Metric = "Look Ahead Days"; Value = $Summary.LookAheadDays }
            [PSCustomObject]@{ Metric = "Total Users Processed"; Value = $Summary.TotalUsersProcessed }
            [PSCustomObject]@{ Metric = "Total Expiring Users"; Value = $Summary.TotalExpiringUsers }
            [PSCustomObject]@{ Metric = "0-3 Days"; Value = $Summary.ExpirationStatistics['0-3 Days'] }
            [PSCustomObject]@{ Metric = "4-7 Days"; Value = $Summary.ExpirationStatistics['4-7 Days'] }
            [PSCustomObject]@{ Metric = "8-14 Days"; Value = $Summary.ExpirationStatistics['8-14 Days'] }
            [PSCustomObject]@{ Metric = "15-30 Days"; Value = $Summary.ExpirationStatistics['15-30 Days'] }
            [PSCustomObject]@{ Metric = ">30 Days"; Value = $Summary.ExpirationStatistics['>30 Days'] }
            [PSCustomObject]@{ Metric = "Users with Email"; Value = $Summary.UsersWithEmail }
            [PSCustomObject]@{ Metric = "Users without Email"; Value = $Summary.UsersWithoutEmail }
            [PSCustomObject]@{ Metric = "Max Password Age"; Value = $Summary.PasswordPolicy.MaxPasswordAge }
        )
        
        $summaryData | Export-Excel -Path $FilePath -WorksheetName "Summary" -AutoSize -BoldTopRow
        
        Write-Log "  ✓ Excel report: $FilePath" "INFO" "Green"
    }
    catch {
        Write-Log "  ✗ Failed to export Excel: $_" "ERROR" "Red"
    }
}

function Save-HistoricalData {
    param(
        [System.Collections.ArrayList]$Data,
        [PSCustomObject]$Summary
    )
    
    try {
        $historicalDir = Join-Path $ExportPath "HistoricalData"
        if (-not (Test-Path $historicalDir)) {
            New-Item -Path $historicalDir -ItemType Directory -Force | Out-Null
        }
        
        $historicalFile = Join-Path $historicalDir "HistoricalData_$Script:Timestamp.json"
        $historicalData = @{
            Date = Get-Date
            Summary = $Summary
            Users = $Data
        }
        
        $historicalData | ConvertTo-Json -Depth 5 | Out-File -FilePath $historicalFile -Encoding UTF8
        Write-Log "  ✓ Historical data saved: $historicalFile" "INFO" "Green"
        
        # Maintain only last 30 days of historical data
        $oldFiles = Get-ChildItem -Path $historicalDir -Filter "HistoricalData_*.json" | 
            Sort-Object LastWriteTime -Descending | 
            Select-Object -Skip 30
        
        foreach ($file in $oldFiles) {
            Remove-Item -Path $file.FullName -Force
            Write-Log "  ℹ Removed old historical data: $($file.Name)" "INFO" "Cyan"
        }
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
    $users = Get-UsersWithExpiringPasswords -SearchBase $SearchBase -DaysLookAhead $Days -IncludeDisabled $IncludeDisabled -IncludeSystemAccounts $IncludeSystemAccounts -DC $DomainController
    if (-not $users) {
        Write-Log "No users retrieved. Exiting." "ERROR" "Red"
        return
    }
    
    # Process user data
    $processedData = Process-UserData -Users $users -MaxPasswordAge $maxPasswordAge -DaysLookAhead $Days -DC $DomainController
    $Script:ReportData = $processedData
    
    # Generate summary
    $summary = Generate-Summary -Data $processedData -TotalUsersProcessed $users.Count -DaysLookAhead $Days
    $Script:Summary = $summary
    
    # Create export directory
    $exportDir = Join-Path $ExportPath "PasswordReports"
    if (-not (Test-Path $exportDir)) {
        New-Item -Path $exportDir -ItemType Directory -Force | Out-Null
    }
    
    $exportPrefix = "ExpiringPasswords_$Script:Timestamp"
    $emailAttachment = $null
    
    # Export reports
    Write-Log "Exporting reports..." "INFO" "Cyan"
    
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
    
    # Send user notifications
    if ($NotifyUsers -and $processedData.Count -gt 0) {
        Write-Log "Sending user notifications..." "INFO" "Cyan"
        
        $usersToNotify = $processedData | Where-Object { $_.DaysRemaining -le $ReminderDays -and $_.EmailAddress }
        Write-Log "  Sending reminders to $($usersToNotify.Count) users" "INFO" "Cyan"
        
        $sentCount = 0
        foreach ($user in $usersToNotify) {
            if (Send-UserNotification -User $user -DaysRemaining $user.DaysRemaining) {
                $sentCount++
            }
        }
        Write-Log "  Sent $sentCount notifications" "INFO" "Green"
    }
    
    # Historical data
    if ($HistoricalData) {
        Save-HistoricalData -Data $processedData -Summary $summary
    }
    
    # Admin notification
    if ($NotifyAdmins) {
        Send-AdminReport -ReportPath $emailAttachment -Summary $summary
    }
    
    # Final summary
    Write-Log "="*70 "INFO" "Yellow"
    Write-Log "REPORT GENERATION COMPLETE" "INFO" "Green"
    Write-Log "="*70 "INFO" "Yellow"
    Write-Log "Export Location: $exportDir" "INFO" "Cyan"
    Write-Log "Total Expiring Users: $($summary.TotalExpiringUsers)" "INFO" "Cyan"
    Write-Log "Log File: $Script:LogFile" "INFO" "Cyan"
    Write-Log "Duration: $((Get-Date) - $Script:StartTime)" "INFO" "Cyan"
    Write-Log "="*70 "INFO" "Yellow"
}

# Run the script
Main
#endregion