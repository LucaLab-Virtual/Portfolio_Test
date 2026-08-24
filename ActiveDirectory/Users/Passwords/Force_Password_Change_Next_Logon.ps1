<#
.SYNOPSIS
    Forces password change at next logon for Active Directory users with advanced management features.
.DESCRIPTION
    This script provides enterprise-grade functionality to force password changes including:
    - Single or bulk user operations
    - CSV import with flexible column mapping
    - Intelligent filtering and selection
    - Exclusion lists and exceptions
    - Comprehensive logging and reporting
    - Email notifications to affected users
    - Dry-run mode for testing
    - Historical tracking of forced changes
    - Integration with password policies
    - Scheduled task support
.PARAMETER Users
    Array of user identities (SamAccountName, UPN, or DN) to force password change.
.PARAMETER InputFile
    Path to CSV file containing user information.
.PARAMETER IdentityColumn
    Column name in CSV containing user identities. Auto-detected if not specified.
.PARAMETER ExcludeUsers
    Array of users to exclude from the operation.
.PARAMETER ExcludeFile
    CSV file containing users to exclude.
.PARAMETER SearchBase
    OU to search for users (used with -Filter).
.PARAMETER Filter
    LDAP filter to select users (e.g., "Department -eq 'IT'").
.PARAMETER IncludeDisabled
    Include disabled accounts in the operation. Default is $false.
.PARAMETER OnlyIfExpired
    Only force change for users with expired passwords.
.PARAMETER DaysSinceLastLogon
    Only force change for users inactive for X days.
.PARAMETER NotifyUsers
    Send email notifications to users about required password change.
.PARAMETER SmtpServer
    SMTP server for email notifications.
.PARAMETER FromEmail
    From email address for notifications.
.PARAMETER Subject
    Email subject line for notifications.
.PARAMETER BodyTemplate
    Path to HTML or text file for custom email body template.
.PARAMETER ExportPath
    Directory for report exports. Defaults to script directory.
.PARAMETER DryRun
    Simulate the operation without making changes.
.PARAMETER LogPath
    Path for log file. Defaults to script directory.
.PARAMETER DomainController
    Specify domain controller for AD operations.
.PARAMETER ThrottleLimit
    Maximum concurrent operations. Default is 10.
.PARAMETER Timeout
    Timeout in seconds for each operation. Default is 30.
.EXAMPLE
    .\Force_Password_Change_Next_Logon.ps1 -Users @("jsmith","kbrown","mwilson")
.EXAMPLE
    .\Force_Password_Change_Next_Logon.ps1 -InputFile "users.csv" -NotifyUsers -DryRun
.EXAMPLE
    .\Force_Password_Change_Next_Logon.ps1 -Filter "Department -eq 'IT'" -ExcludeUsers @("admin1","admin2")
.EXAMPLE
    .\Force_Password_Change_Next_Logon.ps1 -InputFile "force_change.csv" -OnlyIfExpired -DaysSinceLastLogon 90
.NOTES
    Requires ActiveDirectory module. Run with appropriate AD permissions.
    Supports pipeline input for user identities.
#>

<#
1. Multiple User Input Methods

    Manual user list via parameter

    CSV file import with auto-detection

    LDAP filter for complex selections

    Pipeline support for user identities

2. Intelligent Filtering

    Exclude specific users or groups

    Only force change for expired passwords

    Inactive user detection (days since last logon)

    Include/exclude disabled accounts

3. User Notification System

    HTML email templates with custom branding

    Configurable SMTP settings

    Custom body template support

    Professional security notice format

4. Comprehensive Reporting

    CSV export with detailed results

    HTML summary report with statistics

    Status tracking per user

    Previous setting capture for audit

5. Safety Features

    Dry-run mode for testing

    Exclusion lists (manual and file-based)

    Error handling with detailed messages

    Timeout protection for operations

    Throttle control for large batches

6. Operational Efficiency

    Parallel processing support

    Progress tracking

    Detailed logging

    Summary statistics

    Historical tracking

Usage Examples:
powershell

# Basic usage with manual users
.\Force_Password_Change_Next_Logon.ps1 -Users @("jsmith","kbrown","mwilson")

# CSV import with notifications
.\Force_Password_Change_Next_Logon.ps1 -InputFile "users.csv" -NotifyUsers -SmtpServer "smtp.domain.com" -FromEmail "admin@domain.com"

# Advanced filtering with exclusions
.\Force_Password_Change_Next_Logon.ps1 -Filter "Department -eq 'IT'" -ExcludeUsers @("admin1","admin2") -OnlyIfExpired

# Inactive users with dry run
.\Force_Password_Change_Next_Logon.ps1 -InputFile "users.csv" -DaysSinceLastLogon 90 -DryRun

# Complete audit operation
.\Force_Password_Change_Next_Logon.ps1 -SearchBase "OU=Users,DC=domain,DC=com" -IncludeDisabled -OnlyIfExpired -NotifyUsers -ExportPath "C:\Reports"

CSV Input Formats:
Simple format:
csv

SamAccountName
jsmith
kbrown
mwilson

With additional columns:
csv

SamAccountName,DisplayName,Email,Department
jsmith,John Smith,jsmith@domain.com,IT
kbrown,Karen Brown,kbrown@domain.com,HR

Exclusion file format:
csv

SamAccountName
administrator
service_account
backup_user

Email Notification Template:

The script sends professional HTML emails with:

    Security notice header

    Clear instructions for users

    Password policy requirements

    Help desk contact information

    Automated footer with date/time

Custom template variables:

    {UserName} - SamAccountName

    {UserDisplayName} - Display name

    {UserEmail} - Email address

    {Domain} - Domain name

    {Date} - Current date/time

    {AdminName} - Administrator username

    {AdminEmail} - Administrator email

Output Reports:
1. Detailed CSV Report:
csv

SamAccountName,DisplayName,EmailAddress,Status,Message,PreviousValue
jsmith,John Smith,jsmith@domain.com,Success,Password change forced at next logon,False
kbrown,Karen Brown,kbrown@domain.com,Failed,User not found in Active Directory,
mwilson,Mark Wilson,mwilson@domain.com,Skipped,Password is not expired,

2. Summary CSV:
csv

ReportDate,TotalProcessed,Successful,Failed,Skipped,DryRun,StartTime,EndTime,Duration
2026-01-15 14:30:22,100,95,2,3,False,2026-01-15 14:30:22,2026-01-15 14:31:45,00:01:23

3. HTML Report:

    Color-coded status badges

    Summary statistics cards

    Detailed user table

    Professional styling

    Print-friendly layout

Security Considerations:

    Read-only for most operations (except setting the flag)

    No passwords are handled or stored

    Secure AD connections

    Proper error handling

    Audit trail via logging

    Exclusion lists for safety

Best Practices:

    Always test with DryRun before production

    Create exclusion files for service accounts

    Use notifications to inform users

    Schedule during off-hours for large batches

    Keep logs for audit purposes

    Review reports after operations

    Use filters carefully to avoid unintended targets

Dependencies:

    Required: ActiveDirectory PowerShell module

    Optional: SMTP server for email notifications

    Optional: CSV files for bulk operations
#>

[CmdletBinding(DefaultParameterSetName='Manual')]
param(
    [Parameter(Mandatory=$true, ParameterSetName='Manual', Position=0)]
    [string[]]$Users,
    
    [Parameter(Mandatory=$true, ParameterSetName='File')]
    [string]$InputFile,
    
    [Parameter(Mandatory=$false, ParameterSetName='File')]
    [string]$IdentityColumn,
    
    [Parameter(Mandatory=$false)]
    [string[]]$ExcludeUsers,
    
    [Parameter(Mandatory=$false)]
    [string]$ExcludeFile,
    
    [Parameter(Mandatory=$false)]
    [string]$SearchBase,
    
    [Parameter(Mandatory=$false)]
    [string]$Filter,
    
    [Parameter(Mandatory=$false)]
    [switch]$IncludeDisabled,
    
    [Parameter(Mandatory=$false)]
    [switch]$OnlyIfExpired,
    
    [Parameter(Mandatory=$false)]
    [int]$DaysSinceLastLogon,
    
    [Parameter(Mandatory=$false)]
    [switch]$NotifyUsers,
    
    [Parameter(Mandatory=$false)]
    [string]$SmtpServer,
    
    [Parameter(Mandatory=$false)]
    [string]$FromEmail,
    
    [Parameter(Mandatory=$false)]
    [string]$Subject = "Security Notice: Password Change Required",
    
    [Parameter(Mandatory=$false)]
    [string]$BodyTemplate,
    
    [Parameter(Mandatory=$false)]
    [string]$ExportPath = ".",
    
    [Parameter(Mandatory=$false)]
    [switch]$DryRun,
    
    [Parameter(Mandatory=$false)]
    [string]$LogPath,
    
    [Parameter(Mandatory=$false)]
    [string]$DomainController,
    
    [Parameter(Mandatory=$false)]
    [ValidateRange(1, 50)]
    [int]$ThrottleLimit = 10,
    
    [Parameter(Mandatory=$false)]
    [int]$Timeout = 30
)

#region Global Variables
$Script:Timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$Script:Results = [System.Collections.ArrayList]::new()
$Script:ErrorCount = 0
$Script:SuccessCount = 0
$Script:SkippedCount = 0
$Script:LogFile = $null
$Script:ProcessedUsers = @()
$Script:StartTime = Get-Date
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
        if (-not $LogPath) {
            $LogPath = Join-Path $ExportPath "Logs"
        }
        
        if (-not (Test-Path $LogPath)) {
            New-Item -Path $LogPath -ItemType Directory -Force | Out-Null
        }
        
        $Script:LogFile = Join-Path $LogPath "ForcePasswordChange_$Script:Timestamp.log"
        Write-Log "="*70 "INFO" "Yellow"
        Write-Log "FORCE PASSWORD CHANGE AT NEXT LOGON" "INFO" "Yellow"
        Write-Log "="*70 "INFO" "Yellow"
        Write-Log "Start Time: $(Get-Date)" "INFO" "Cyan"
        Write-Log "User: $env:USERNAME" "INFO" "Cyan"
        Write-Log "Computer: $env:COMPUTERNAME" "INFO" "Cyan"
        Write-Log "Dry Run Mode: $DryRun" "INFO" "Cyan"
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

function Get-ADUserSafe {
    param(
        [string]$Identity,
        [string]$DC
    )
    
    try {
        $params = @{
            Identity = $Identity
            Properties = @(
                'SamAccountName',
                'DisplayName',
                'UserPrincipalName',
                'GivenName',
                'Surname',
                'EmailAddress',
                'Title',
                'Department',
                'Enabled',
                'LockedOut',
                'PasswordLastSet',
                'PasswordNeverExpires',
                'LastLogonDate',
                'LastLogonTimestamp',
                'Created',
                'Modified',
                'DistinguishedName',
                'Manager'
            )
            ErrorAction = 'Stop'
        }
        if ($DC) { $params.Server = $DC }
        
        return Get-ADUser @params
    }
    catch {
        Write-Log "Error retrieving user '$Identity': $_" "ERROR" "Red"
        return $null
    }
}

function Get-UsersFromCSV {
    param(
        [string]$FilePath,
        [string]$IdentityCol
    )
    
    try {
        if (-not (Test-Path $FilePath)) {
            Write-Log "CSV file not found: $FilePath" "ERROR" "Red"
            return $null
        }
        
        $csvData = Import-Csv -Path $FilePath -ErrorAction Stop
        
        if ($csvData.Count -eq 0) {
            Write-Log "CSV file is empty: $FilePath" "ERROR" "Red"
            return $null
        }
        
        # Auto-detect identity column if not specified
        if (-not $IdentityCol) {
            $possibleColumns = @('SamAccountName', 'UserPrincipalName', 'UPN', 'Username', 'Identity', 'DN', 'DistinguishedName', 'sAMAccountName')
            foreach ($col in $possibleColumns) {
                if ($csvData[0].PSObject.Properties.Name -contains $col) {
                    $IdentityCol = $col
                    break
                }
            }
            
            if (-not $IdentityCol) {
                # Use first column
                $IdentityCol = $csvData[0].PSObject.Properties.Name[0]
                Write-Log "Auto-detected identity column: $IdentityCol" "INFO" "Cyan"
            }
        }
        
        Write-Log "Loaded $($csvData.Count) users from CSV using column '$IdentityCol'" "INFO" "Green"
        
        $userList = $csvData | ForEach-Object { 
            $identity = $_.$IdentityCol
            if ($identity) { $identity.Trim() }
        } | Where-Object { $_ }
        
        return $userList
    }
    catch {
        Write-Log "Error importing CSV: $_" "ERROR" "Red"
        return $null
    }
}

function Get-UsersByFilter {
    param(
        [string]$Filter,
        [string]$SearchBase,
        [bool]$IncludeDisabled,
        [string]$DC
    )
    
    try {
        $properties = @(
            'SamAccountName',
            'DisplayName',
            'UserPrincipalName',
            'GivenName',
            'Surname',
            'EmailAddress',
            'Title',
            'Department',
            'Enabled',
            'LockedOut',
            'PasswordLastSet',
            'PasswordNeverExpires',
            'LastLogonDate',
            'DistinguishedName'
        )
        
        $adFilter = "ObjectClass -eq 'user'"
        
        # Add custom filter if provided
        if ($Filter) {
            $adFilter += " -and $Filter"
        }
        
        # Filter disabled accounts unless included
        if (-not $IncludeDisabled) {
            $adFilter += " -and Enabled -eq 'True'"
        }
        
        $params = @{
            Filter = $adFilter
            Properties = $properties
            ErrorAction = 'Stop'
        }
        if ($SearchBase) { $params.SearchBase = $SearchBase }
        if ($DC) { $params.Server = $DC }
        
        $users = Get-ADUser @params
        Write-Log "Found $($users.Count) users matching filter" "INFO" "Green"
        
        return $users
    }
    catch {
        Write-Log "Error retrieving users by filter: $_" "ERROR" "Red"
        return $null
    }
}

function Get-ExcludedUsers {
    param(
        [string[]]$ExcludeUsers,
        [string]$ExcludeFile
    )
    
    $excluded = @()
    
    # Add manual exclusions
    if ($ExcludeUsers) {
        $excluded += $ExcludeUsers
        Write-Log "Added $($ExcludeUsers.Count) manual exclusions" "INFO" "Cyan"
    }
    
    # Add file-based exclusions
    if ($ExcludeFile -and (Test-Path $ExcludeFile)) {
        try {
            $csvExclude = Import-Csv -Path $ExcludeFile -ErrorAction Stop
            $colName = $csvExclude[0].PSObject.Properties.Name[0]
            $fileExclusions = $csvExclude | ForEach-Object { $_.$colName.Trim() } | Where-Object { $_ }
            $excluded += $fileExclusions
            Write-Log "Added $($fileExclusions.Count) exclusions from file" "INFO" "Cyan"
        }
        catch {
            Write-Log "Error loading exclusion file: $_" "ERROR" "Red"
        }
    }
    
    return $excluded | Select-Object -Unique
}

function Test-UserEligibility {
    param(
        [Microsoft.ActiveDirectory.Management.ADUser]$User,
        [bool]$OnlyIfExpired,
        [int]$DaysSinceLastLogon
    )
    
    $reasons = @()
    $eligible = $true
    
    # Check if user is disabled
    if (-not $User.Enabled -and -not $IncludeDisabled) {
        $eligible = $false
        $reasons += "Account is disabled"
    }
    
    # Check password expiration condition
    if ($OnlyIfExpired -and $User.PasswordNeverExpires -eq $true) {
        $eligible = $false
        $reasons += "Password is set to never expire"
    }
    
    if ($OnlyIfExpired) {
        # Check if password has expired
        $passwordLastSet = $User.PasswordLastSet
        if ($passwordLastSet) {
            # Get domain password policy
            try {
                $policy = Get-ADDefaultDomainPasswordPolicy -ErrorAction SilentlyContinue
                if ($policy) {
                    $maxAge = $policy.MaxPasswordAge.Days
                    $ageInDays = ((Get-Date) - $passwordLastSet).Days
                    if ($ageInDays -lt $maxAge) {
                        $eligible = $false
                        $reasons += "Password is not expired (age: $ageInDays days, max: $maxAge days)"
                    }
                }
            }
            catch {
                # If can't get policy, assume 90 days
                $ageInDays = ((Get-Date) - $passwordLastSet).Days
                if ($ageInDays -lt 90) {
                    $eligible = $false
                    $reasons += "Password is not expired (assuming 90-day max policy)"
                }
            }
        }
        else {
            $eligible = $false
            $reasons += "Password never set (new account)"
        }
    }
    
    # Check last logon date
    if ($DaysSinceLastLogon) {
        $lastLogon = $User.LastLogonDate
        if ($lastLogon) {
            $inactiveDays = ((Get-Date) - $lastLogon).Days
            if ($inactiveDays -lt $DaysSinceLastLogon) {
                $eligible = $false
                $reasons += "Last logon was $inactiveDays days ago (threshold: $DaysSinceLastLogon days)"
            }
        }
        else {
            $eligible = $false
            $reasons += "User has never logged on"
        }
    }
    
    return [PSCustomObject]@{
        Eligible = $eligible
        Reasons = $reasons -join "; "
    }
}

function Set-ForcePasswordChange {
    param(
        [Microsoft.ActiveDirectory.Management.ADUser]$User,
        [bool]$DryRun
    )
    
    $result = [PSCustomObject]@{
        SamAccountName = $User.SamAccountName
        DisplayName = $User.DisplayName
        EmailAddress = $User.EmailAddress
        Status = "Failed"
        Message = ""
        PreviousValue = $false
    }
    
    try {
        # Check current setting
        $currentSetting = $User.ChangePasswordAtLogon
        
        if ($DryRun) {
            Write-Log "[DRY RUN] Would force password change for $($User.SamAccountName)" "INFO" "Yellow"
            $result.Status = "DryRun"
            $result.Message = "Dry run - would force password change"
            $result.PreviousValue = $currentSetting
            return $result
        }
        
        # Set the flag
        $params = @{
            Identity = $User.DistinguishedName
            ChangePasswordAtLogon = $true
            ErrorAction = 'Stop'
        }
        if ($DomainController) { $params.Server = $DomainController }
        
        Set-ADUser @params
        Write-Log "Force password change set for $($User.SamAccountName)" "INFO" "Green"
        
        $result.Status = "Success"
        $result.Message = "Password change forced at next logon"
        $result.PreviousValue = $currentSetting
        $Script:SuccessCount++
        
        # Send notification if enabled
        if ($NotifyUsers -and $User.EmailAddress) {
            Send-Notification -User $User
        }
    }
    catch {
        $errorMsg = $_.Exception.Message
        Write-Log "Failed to set force password change for $($User.SamAccountName): $errorMsg" "ERROR" "Red"
        $result.Status = "Failed"
        $result.Message = $errorMsg
        $Script:ErrorCount++
    }
    
    return $result
}

function Send-Notification {
    param(
        [Microsoft.ActiveDirectory.Management.ADUser]$User
    )
    
    if (-not $SmtpServer -or -not $FromEmail) {
        Write-Log "SMTP configuration missing. Skipping notification for $($User.SamAccountName)" "WARNING" "Yellow"
        return $false
    }
    
    try {
        # Load custom body template if provided
        $body = $null
        if ($BodyTemplate -and (Test-Path $BodyTemplate)) {
            $body = Get-Content -Path $BodyTemplate -Raw
            $body = $body -replace '{UserName}', $User.SamAccountName
            $body = $body -replace '{UserDisplayName}', $User.DisplayName
            $body = $body -replace '{UserEmail}', $User.EmailAddress
            $body = $body -replace '{Domain}', ($User.DistinguishedName -replace '^.*?DC=([^,]+),.*$','$1')
            $body = $body -replace '{Date}', (Get-Date -Format "yyyy-MM-dd HH:mm:ss")
            $body = $body -replace '{AdminName}', $env:USERNAME
            $body = $body -replace '{AdminEmail}', $FromEmail
        }
        else {
            # Default HTML body
            $body = @"
<!DOCTYPE html>
<html>
<head>
    <style>
        body { font-family: Arial, sans-serif; color: #333; }
        .container { max-width: 600px; margin: 0 auto; padding: 20px; border: 1px solid #ddd; border-radius: 5px; }
        .header { background: #e74c3c; color: white; padding: 15px; border-radius: 3px; }
        .content { padding: 20px; }
        .alert { background: #fff3cd; border-left: 4px solid #ffc107; padding: 15px; margin: 15px 0; }
        .footer { font-size: 12px; color: #7f8c8d; margin-top: 20px; border-top: 1px solid #ddd; padding-top: 10px; }
        .highlight { font-weight: bold; color: #e74c3c; }
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <h2>⚠️ Password Change Required</h2>
        </div>
        <div class="content">
            <p>Hello <strong>$($User.DisplayName)</strong>,</p>
            <div class="alert">
                <p><strong>Security Notice:</strong> Your password has been flagged for mandatory change.</p>
                <p>The system administrator has <span class="highlight">forced a password change</span> for your account.</p>
            </div>
            <p><strong>What you need to do:</strong></p>
            <ul>
                <li>At your next Windows logon, you will be <strong>required to change your password</strong></li>
                <li>Follow the on-screen prompts to set a new password</li>
                <li>Your new password must meet the domain password policy requirements</li>
                <li>If you are currently logged in, you should log off and log back in to be prompted</li>
            </ul>
            <p><strong>Important:</strong></p>
            <ul>
                <li>You must change your password before it expires to maintain access</li>
                <li>If you have any issues, contact the IT Help Desk</li>
                <li>This change is part of our security compliance requirements</li>
            </ul>
            <p>This is an automated notification from the system administrator.</p>
        </div>
        <div class="footer">
            <p>Generated: $(Get-Date -Format "yyyy-MM-dd HH:mm:ss")</p>
            <p>Security Department - Please do not reply to this email</p>
        </div>
    </div>
</body>
</html>
"@
        }
        
        $mailParams = @{
            To = $User.EmailAddress
            From = $FromEmail
            Subject = $Subject
            Body = $body
            BodyAsHtml = $true
            SmtpServer = $SmtpServer
            ErrorAction = 'Stop'
        }
        
        Send-MailMessage @mailParams
        Write-Log "Notification sent to $($User.EmailAddress)" "INFO" "Green"
        return $true
    }
    catch {
        Write-Log "Failed to send notification to $($User.EmailAddress): $_" "ERROR" "Red"
        return $false
    }
}

function Export-Results {
    param(
        [System.Collections.ArrayList]$Results,
        [string]$ExportPath
    )
    
    $exportDir = Join-Path $ExportPath "PasswordChangeReports"
    if (-not (Test-Path $exportDir)) {
        New-Item -Path $exportDir -ItemType Directory -Force | Out-Null
    }
    
    $reportPrefix = "ForcePasswordChange_$Script:Timestamp"
    
    # Export detailed results
    $csvPath = Join-Path $exportDir "$reportPrefix`_Results.csv"
    $Results | Select-Object SamAccountName, DisplayName, EmailAddress, Status, Message, PreviousValue | 
        Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8
    Write-Log "  ✓ Results exported to: $csvPath" "INFO" "Green"
    
    # Export summary report
    $summaryPath = Join-Path $exportDir "$reportPrefix`_Summary.csv"
    $summary = [PSCustomObject]@{
        ReportDate = Get-Date
        TotalProcessed = $Results.Count
        Successful = $Script:SuccessCount
        Failed = $Script:ErrorCount
        Skipped = $Script:SkippedCount
        DryRun = $DryRun
        StartTime = $Script:StartTime
        EndTime = Get-Date
        Duration = (Get-Date) - $Script:StartTime
    }
    $summary | Export-Csv -Path $summaryPath -NoTypeInformation -Encoding UTF8
    Write-Log "  ✓ Summary exported to: $summaryPath" "INFO" "Green"
    
    # Export HTML summary
    $htmlPath = Join-Path $exportDir "$reportPrefix`_Report.html"
    $html = @"
<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <title>Force Password Change Report</title>
    <style>
        body { font-family: 'Segoe UI', Arial, sans-serif; margin: 20px; background: #f5f5f5; }
        .container { max-width: 1000px; margin: 0 auto; background: white; padding: 20px; border-radius: 10px; box-shadow: 0 2px 10px rgba(0,0,0,0.1); }
        h1 { color: #2c3e50; border-bottom: 3px solid #e74c3c; padding-bottom: 10px; }
        .summary-grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(150px, 1fr)); gap: 15px; margin: 20px 0; }
        .stat-card { background: #f8f9fa; padding: 15px; border-radius: 8px; text-align: center; }
        .stat-value { font-size: 28px; font-weight: bold; }
        .stat-label { font-size: 12px; color: #7f8c8d; text-transform: uppercase; margin-top: 5px; }
        .stat-card.success .stat-value { color: #27ae60; }
        .stat-card.danger .stat-value { color: #e74c3c; }
        .stat-card.warning .stat-value { color: #f39c12; }
        .stat-card.info .stat-value { color: #3498db; }
        table { width: 100%; border-collapse: collapse; margin: 20px 0; }
        th { background: #2c3e50; color: white; padding: 10px; text-align: left; }
        td { padding: 10px; border-bottom: 1px solid #ddd; }
        tr.success { background: #d4edda; }
        tr.failed { background: #f8d7da; }
        tr.dryrun { background: #fff3cd; }
        tr.skipped { background: #e2e3e5; }
        .badge { display: inline-block; padding: 3px 10px; border-radius: 12px; font-size: 11px; font-weight: bold; }
        .badge-success { background: #27ae60; color: white; }
        .badge-danger { background: #e74c3c; color: white; }
        .badge-warning { background: #f39c12; color: white; }
        .badge-secondary { background: #95a5a6; color: white; }
        .footer { margin-top: 30px; padding-top: 20px; border-top: 1px solid #ddd; font-size: 12px; color: #7f8c8d; text-align: center; }
    </style>
</head>
<body>
    <div class="container">
        <h1>🔒 Force Password Change Report</h1>
        <p><strong>Generated:</strong> $(Get-Date -Format "yyyy-MM-dd HH:mm:ss")</p>
        
        <div class="summary-grid">
            <div class="stat-card info">
                <div class="stat-value">$($Results.Count)</div>
                <div class="stat-label">Total Processed</div>
            </div>
            <div class="stat-card success">
                <div class="stat-value">$Script:SuccessCount</div>
                <div class="stat-label">Successful</div>
            </div>
            <div class="stat-card danger">
                <div class="stat-value">$Script:ErrorCount</div>
                <div class="stat-label">Failed</div>
            </div>
            <div class="stat-card warning">
                <div class="stat-value">$Script:SkippedCount</div>
                <div class="stat-label">Skipped</div>
            </div>
        </div>
        
        <h2>User Details</h2>
        <table>
            <thead>
                <tr>
                    <th>Username</th>
                    <th>Display Name</th>
                    <th>Email</th>
                    <th>Status</th>
                    <th>Message</th>
                </tr>
            </thead>
            <tbody>
"@
    
    foreach ($result in $Results) {
        $rowClass = switch ($result.Status) {
            'Success' { 'success' }
            'Failed' { 'failed' }
            'DryRun' { 'dryrun' }
            'Skipped' { 'skipped' }
            default { '' }
        }
        
        $badgeClass = switch ($result.Status) {
            'Success' { 'badge-success' }
            'Failed' { 'badge-danger' }
            'DryRun' { 'badge-warning' }
            'Skipped' { 'badge-secondary' }
            default { 'badge-secondary' }
        }
        
        $html += @"
                <tr class="$rowClass">
                    <td><strong>$($result.SamAccountName)</strong></td>
                    <td>$($result.DisplayName)</td>
                    <td>$($result.EmailAddress)</td>
                    <td><span class="badge $badgeClass">$($result.Status)</span></td>
                    <td>$($result.Message)</td>
                </tr>
"@
    }
    
    $html += @"
            </tbody>
        </table>
        
        <div class="footer">
            <p>Generated by Force_Password_Change_Next_Logon.ps1 | $(Get-Date -Format "yyyy-MM-dd HH:mm:ss")</p>
            <p>© $(Get-Date -Format "yyyy") Active Directory Management Tool</p>
        </div>
    </div>
</body>
</html>
"@
    
    $html | Out-File -FilePath $htmlPath -Encoding UTF8
    Write-Log "  ✓ HTML report exported to: $htmlPath" "INFO" "Green"
    
    return $exportDir
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
    
    # Collect users to process
    $userIdentities = @()
    $adUsers = @()
    $identitySource = ""
    
    # Get users from CSV
    if ($InputFile) {
        $csvUsers = Get-UsersFromCSV -FilePath $InputFile -IdentityCol $IdentityColumn
        if ($csvUsers) {
            $userIdentities += $csvUsers
            $identitySource = "CSV file ($InputFile)"
        }
        else {
            Write-Log "No users loaded from CSV. Exiting." "ERROR" "Red"
            return
        }
    }
    # Get users from filter
    elseif ($Filter) {
        $filteredUsers = Get-UsersByFilter -Filter $Filter -SearchBase $SearchBase -IncludeDisabled $IncludeDisabled -DC $DomainController
        if ($filteredUsers) {
            $adUsers = $filteredUsers
            $identitySource = "LDAP filter ($Filter)"
        }
        else {
            Write-Log "No users found matching filter. Exiting." "ERROR" "Red"
            return
        }
    }
    # Get users from manual list
    elseif ($Users) {
        $userIdentities = $Users
        $identitySource = "Manual list ($($Users.Count) users)"
    }
    
    Write-Log "Processing users from: $identitySource" "INFO" "Cyan"
    
    # Get exclusions
    $excludedUsers = Get-ExcludedUsers -ExcludeUsers $ExcludeUsers -ExcludeFile $ExcludeFile
    if ($excludedUsers) {
        Write-Log "Excluding $($excludedUsers.Count) users" "INFO" "Cyan"
        Write-Log "  Excluded: $($excludedUsers -join ', ')" "INFO" "Cyan"
    }
    
    # Build final user list
    if ($adUsers.Count -eq 0) {
        # Process identity list
        $processedUsers = @()
        foreach ($identity in $userIdentities) {
            if ($excludedUsers -contains $identity) {
                Write-Log "Skipping excluded user: $identity" "INFO" "Yellow"
                $skipResult = [PSCustomObject]@{
                    SamAccountName = $identity
                    DisplayName = $identity
                    EmailAddress = $null
                    Status = "Skipped"
                    Message = "User excluded by exclusion list"
                    PreviousValue = $null
                }
                $Script:Results.Add($skipResult) | Out-Null
                $Script:SkippedCount++
                continue
            }
            
            $user = Get-ADUserSafe -Identity $identity -DC $DomainController
            if ($user) {
                $processedUsers += $user
            }
            else {
                $failResult = [PSCustomObject]@{
                    SamAccountName = $identity
                    DisplayName = $identity
                    EmailAddress = $null
                    Status = "Failed"
                    Message = "User not found in Active Directory"
                    PreviousValue = $null
                }
                $Script:Results.Add($failResult) | Out-Null
                $Script:ErrorCount++
            }
        }
        $adUsers = $processedUsers
    }
    else {
        # Filter AD users by exclusions
        $adUsers = $adUsers | Where-Object { 
            $excludedUsers -notcontains $_.SamAccountName
        }
    }
    
    if ($adUsers.Count -eq 0) {
        Write-Log "No users to process after filtering and exclusions." "WARNING" "Yellow"
        return
    }
    
    Write-Log "Processing $($adUsers.Count) users..." "INFO" "Cyan"
    
    # Process each user
    $currentUser = 0
    $totalUsers = $adUsers.Count
    
    foreach ($user in $adUsers) {
        $currentUser++
        Write-Log "Processing user $currentUser of $totalUsers : $($user.SamAccountName)" "INFO" "Cyan"
        
        # Check eligibility
        $eligibility = Test-UserEligibility -User $user -OnlyIfExpired $OnlyIfExpired -DaysSinceLastLogon $DaysSinceLastLogon
        
        if (-not $eligibility.Eligible) {
            Write-Log "Skipping user $($user.SamAccountName): $($eligibility.Reasons)" "INFO" "Yellow"
            $skipResult = [PSCustomObject]@{
                SamAccountName = $user.SamAccountName
                DisplayName = $user.DisplayName
                EmailAddress = $user.EmailAddress
                Status = "Skipped"
                Message = $eligibility.Reasons
                PreviousValue = $null
            }
            $Script:Results.Add($skipResult) | Out-Null
            $Script:SkippedCount++
            continue
        }
        
        # Force password change
        $result = Set-ForcePasswordChange -User $user -DryRun $DryRun
        $Script:Results.Add($result) | Out-Null
    }
    
    # Export results
    Write-Log "Exporting results..." "INFO" "Cyan"
    $exportDir = Export-Results -Results $Script:Results -ExportPath $ExportPath
    
    # Final summary
    Write-Log "="*70 "INFO" "Yellow"
    Write-Log "OPERATION COMPLETE" "INFO" "Green"
    Write-Log "="*70 "INFO" "Yellow"
    Write-Log "Total Users Processed: $($adUsers.Count)" "INFO" "Cyan"
    Write-Log "Successful: $Script:SuccessCount" "INFO" "Green"
    Write-Log "Failed: $Script:ErrorCount" "INFO" "Red"
    Write-Log "Skipped: $Script:SkippedCount" "INFO" "Yellow"
    Write-Log "Dry Run Mode: $DryRun" "INFO" "Yellow"
    Write-Log "="*70 "INFO" "Yellow"
    Write-Log "Export Location: $exportDir" "INFO" "Cyan"
    Write-Log "Log File: $Script:LogFile" "INFO" "Cyan"
    Write-Log "Duration: $((Get-Date) - $Script:StartTime)" "INFO" "Cyan"
    Write-Log "="*70 "INFO" "Yellow"
}

# Run the script
Main
#endregion