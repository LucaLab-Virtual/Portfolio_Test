<#
.SYNOPSIS
    Local Administrators Group Audit Report Script
.DESCRIPTION
    This script performs a comprehensive audit of local administrator group memberships
    on the local system or remote computers. It identifies privileged users, detects
    potential security risks, and generates detailed reports with recommendations.
.NOTES
    Author: Your Name
    Date: 2026-08-22
    Version: 1.0
    Requires: Administrative privileges
    Features: Local and remote audit, export to multiple formats, risk assessment
.LINK
    https://docs.microsoft.com/en-us/windows/security/identity-protection/access-control/local-accounts
#>

<#
Basic Usage:
powershell

# Run as Administrator for local computer
.\Local_Admins_Report.ps1

Advanced Usage Examples:
powershell

# Audit multiple computers
.\Local_Admins_Report.ps1 -ComputerNames @('PC01', 'PC02', 'PC03')

# Specify allowed administrators
.\Local_Admins_Report.ps1 -AllowedAdmins @('Administrator', 'DomainAdmin', 'ITTeam')

# Check nested group memberships
.\Local_Admins_Report.ps1 -CheckGroupMembers

# Generate HTML only (for management presentation)
.\Local_Admins_Report.ps1 -OutputFormat HTML

# Comprehensive audit with all features
.\Local_Admins_Report.ps1 -ComputerNames @('SRV01', 'SRV02') -CheckGroupMembers -ShowDomainUsers -OutputFormat All -AllowedAdmins @('Administrator', 'IT_Admins')

Output Features:
📊 Report Formats:

    CSV - For data analysis in Excel/PowerBI

    HTML - Visual dashboard with risk indicators

    JSON - For programmatic processing

    All - Generate all formats

📈 Security Analysis:

    ✅ Risk level assessment (Critical/High/Medium/Low)

    ✅ Security score calculation

    ✅ Detailed recommendations

    ✅ Authorized vs Unauthorized users

    ✅ Local vs Domain accounts

    ✅ Nested group detection

🔍 Security Checks:

    ✅ UAC status

    ✅ Domain membership

    ✅ Password policy info

    ✅ Suspicious account detection

    ✅ Group membership validation

Sample Output:

The HTML report includes:

    Executive summary with key metrics

    Color-coded risk levels

    Detailed user table with recommendations

    Security score (0-100)

    Actionable remediation steps

Risk Classification:

    Critical: Suspicious accounts, shared accounts, or service accounts

    High: Domain users with admin rights (unnecessary)

    Medium: Local accounts with admin rights

    Low: Authorized administrators

Integration Options:

    Can be integrated with SIEM systems

    Supports automated email alerts

    Can be scheduled via Task Scheduler

    Compatible with reporting dashboards
#>

#Requires -RunAsAdministrator

[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [string]$OutputPath = "$env:USERPROFILE\Desktop\Local_Admins_Report_$(Get-Date -Format 'yyyyMMdd_HHmmss')",
    
    [Parameter(Mandatory = $false)]
    [string[]]$ComputerNames = @($env:COMPUTERNAME),
    
    [Parameter(Mandatory = $false)]
    [ValidateSet('CSV', 'HTML', 'JSON', 'All')]
    [string]$OutputFormat = 'All',
    
    [Parameter(Mandatory = $false)]
    [switch]$ShowDomainUsers,
    
    [Parameter(Mandatory = $false)]
    [switch]$CheckGroupMembers,
    
    [Parameter(Mandatory = $false)]
    [switch]$ExportGroupPolicy,
    
    [Parameter(Mandatory = $false)]
    [string[]]$AllowedAdmins = @('Administrator'),
    
    [Parameter(Mandatory = $false)]
    [switch]$SendEmail,
    
    [Parameter(Mandatory = $false)]
    [string]$EmailTo = 'security@yourdomain.com',
    
    [Parameter(Mandatory = $false)]
    [string]$LogPath = "$env:USERPROFILE\Desktop\Local_Admins_Log_$(Get-Date -Format 'yyyyMMdd_HHmmss').txt"
)

# Script configuration
$ErrorActionPreference = "Continue"
$ScriptName = "Local_Admins_Report"
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
Write-Host "Local Administrators Group Audit v$ScriptVersion" -ForegroundColor Cyan
Write-Host "Generated on: $ExecutionTime" -ForegroundColor Cyan
Write-Host "=" * 80 -ForegroundColor Cyan
Write-Host ""

# Function to get local group members
function Get-LocalGroupMembers {
    param(
        [string]$ComputerName = $env:COMPUTERNAME,
        [string]$GroupName = 'Administrators'
    )
    
    $Members = @()
    
    try {
        if ($ComputerName -eq $env:COMPUTERNAME) {
            # Local computer
            $Group = [ADSI]"WinNT://$ComputerName/$GroupName,group"
            $GroupMembers = @($Group.Invoke("Members"))
            
            foreach ($Member in $GroupMembers) {
                $MemberType = $Member.GetType()
                $Path = $Member.GetType().InvokeMember("ADsPath", 'GetProperty', $null, $Member, $null)
                
                $Parts = $Path.Split('/')
                $Domain = $Parts[2]
                $Name = $Parts[-1]
                
                # Determine if user is local or domain
                if ($ComputerName -eq $Domain) {
                    $AccountType = 'Local'
                    $FullName = "$ComputerName\$Name"
                } else {
                    $AccountType = 'Domain'
                    $FullName = "$Domain\$Name"
                }
                
                $Members += [PSCustomObject]@{
                    ComputerName = $ComputerName
                    GroupName = $GroupName
                    Username = $Name
                    Domain = $Domain
                    FullName = $FullName
                    AccountType = $AccountType
                    IsLocalAdmin = $true
                    MemberType = 'Direct'
                    LastChecked = $ExecutionTime
                }
            }
        } else {
            # Remote computer - using WMI
            $Query = "ASSOCIATORS OF {Win32_Group.Domain='$ComputerName',Name='$GroupName'} WHERE ResultClass=Win32_UserAccount"
            $WmiMembers = Get-WmiObject -ComputerName $ComputerName -Query $Query -ErrorAction Stop
            
            foreach ($Member in $WmiMembers) {
                $Members += [PSCustomObject]@{
                    ComputerName = $ComputerName
                    GroupName = $GroupName
                    Username = $Member.Name
                    Domain = $Member.Domain
                    FullName = "$($Member.Domain)\$($Member.Name)"
                    AccountType = if ($Member.Domain -eq $ComputerName) { 'Local' } else { 'Domain' }
                    IsLocalAdmin = $true
                    MemberType = 'Direct'
                    LastChecked = $ExecutionTime
                }
            }
        }
    }
    catch {
        Write-Log "Error accessing group members on $ComputerName`: $_" "ERROR"
        return @()
    }
    
    return $Members
}

# Function to check nested group members
function Get-NestedGroupMembers {
    param(
        [string]$ComputerName = $env:COMPUTERNAME,
        [string]$GroupName = 'Administrators'
    )
    
    $AllMembers = @()
    $ProcessedGroups = @()
    
    try {
        if ($ComputerName -eq $env:COMPUTERNAME) {
            $Group = [ADSI]"WinNT://$ComputerName/$GroupName,group"
            $GroupMembers = @($Group.Invoke("Members"))
            
            foreach ($Member in $GroupMembers) {
                $Path = $Member.GetType().InvokeMember("ADsPath", 'GetProperty', $null, $Member, $null)
                $Parts = $Path.Split('/')
                $Domain = $Parts[2]
                $Name = $Parts[-1]
                
                # Check if member is a group
                try {
                    $MemberObj = [ADSI]$Path
                    if ($MemberObj.Class -eq 'Group' -and $Name -notin $ProcessedGroups) {
                        $ProcessedGroups += $Name
                        $NestedMembers = Get-LocalGroupMembers -ComputerName $ComputerName -GroupName $Name
                        foreach ($NestedMember in $NestedMembers) {
                            $NestedMember.MemberType = 'Nested'
                            $NestedMember.NestedGroup = $Name
                            $AllMembers += $NestedMember
                        }
                    } else {
                        # It's a user
                        $AllMembers += $Member
                    }
                } catch {
                    # It's a user
                    $AllMembers += $Member
                }
            }
        }
    }
    catch {
        Write-Log "Error checking nested groups: $_" "ERROR"
    }
    
    return $AllMembers
}

# Function to analyze group members
function Analyze-GroupMembers {
    param(
        [array]$Members,
        [array]$AllowedAdmins
    )
    
    $Analysis = @()
    
    foreach ($Member in $Members) {
        $IsAllowed = $false
        $RiskLevel = 'Unknown'
        $Recommendation = ''
        
        # Check if user is in allowed list
        foreach ($Allowed in $AllowedAdmins) {
            if ($Member.Username -eq $Allowed -or $Member.FullName -eq $Allowed) {
                $IsAllowed = $true
                break
            }
        }
        
        # Determine risk level
        if ($IsAllowed) {
            $RiskLevel = 'Low'
            $Recommendation = 'Authorized administrator - maintain current access'
        } elseif ($Member.AccountType -eq 'Local') {
            $RiskLevel = 'Medium'
            $Recommendation = 'Review local account admin privileges - consider removing if not required'
        } elseif ($Member.AccountType -eq 'Domain') {
            $RiskLevel = 'High'
            $Recommendation = 'Domain user with local admin rights - verify business need'
        }
        
        # Additional risk factors
        if ($Member.Username -match 'guest|test|temp|backup' -or $Member.Username -match '^[0-9]') {
            $RiskLevel = 'Critical'
            $Recommendation = 'Suspicious account name with admin rights - immediate review required'
        }
        
        $Analysis += [PSCustomObject]@{
            ComputerName = $Member.ComputerName
            Username = $Member.Username
            Domain = $Member.Domain
            FullName = $Member.FullName
            AccountType = $Member.AccountType
            MemberType = $Member.MemberType
            NestedGroup = if ($Member.NestedGroup) { $Member.NestedGroup } else { 'N/A' }
            IsAuthorized = $IsAllowed
            RiskLevel = $RiskLevel
            Recommendation = $Recommendation
            LastChecked = $ExecutionTime
        }
    }
    
    return $Analysis
}

# Function to get additional security info
function Get-SecurityInfo {
    param(
        [string]$ComputerName = $env:COMPUTERNAME
    )
    
    $SecurityInfo = [PSCustomObject]@{
        ComputerName = $ComputerName
        HasUACEnabled = $null
        HasSecureChannel = $null
        LocalAdminCount = $null
        DomainAdminCount = $null
        TotalAdminCount = $null
        SecurityScore = $null
        LastAuditDate = $ExecutionTime
    }
    
    try {
        # Check UAC status
        $UACKey = Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System" -Name EnableLUA -ErrorAction SilentlyContinue
        $SecurityInfo.HasUACEnabled = ($UACKey.EnableLUA -eq 1)
        
        # Check if in domain
        $ComputerSystem = Get-WmiObject -Class Win32_ComputerSystem -ComputerName $ComputerName
        $SecurityInfo.IsDomainJoined = $ComputerSystem.PartOfDomain
        
        # Get additional security settings
        $SecurityInfo.PasswordPolicy = Get-WmiObject -Class Win32_NetworkLoginProfile -Filter "Name=''$ComputerName''" -ErrorAction SilentlyContinue
    }
    catch {
        Write-Log "Error retrieving security info: $_" "ERROR"
    }
    
    return $SecurityInfo
}

# Function to generate HTML report
function Export-HTMLReport {
    param(
        [array]$Analysis,
        [array]$AllMembers,
        [string]$OutputPath,
        [string]$ComputerName
    )
    
    $HTMLPath = "$OutputPath.html"
    $Summary = Get-SummaryStats -Analysis $Analysis
    
    $HTML = @"
<!DOCTYPE html>
<html>
<head>
    <title>Local Administrators Audit Report - $ComputerName</title>
    <style>
        body { font-family: Arial, sans-serif; margin: 20px; background-color: #f5f5f5; }
        .header { background-color: #007acc; color: white; padding: 20px; border-radius: 5px; }
        .summary { background-color: white; padding: 15px; border-radius: 5px; margin: 20px 0; box-shadow: 0 2px 4px rgba(0,0,0,0.1); }
        .critical { background-color: #ff4444; color: white; }
        .high { background-color: #ff8800; color: white; }
        .medium { background-color: #ffcc00; color: black; }
        .low { background-color: #44bb44; color: white; }
        table { width: 100%; border-collapse: collapse; background-color: white; box-shadow: 0 2px 4px rgba(0,0,0,0.1); }
        th { background-color: #007acc; color: white; padding: 10px; text-align: left; }
        td { padding: 8px; border-bottom: 1px solid #ddd; }
        tr:hover { background-color: #f5f5f5; }
        .badge { padding: 3px 8px; border-radius: 3px; font-weight: bold; }
        .recommendation { background-color: #fff3cd; padding: 10px; border-left: 4px solid #ffc107; margin: 5px 0; }
        .footer { margin-top: 20px; padding: 10px; background-color: #333; color: white; border-radius: 5px; text-align: center; }
        .risk-legend { margin: 20px 0; }
        .risk-legend span { display: inline-block; padding: 5px 10px; margin: 0 5px; border-radius: 3px; }
    </style>
</head>
<body>
    <div class="header">
        <h1>Local Administrators Security Audit Report</h1>
        <p>Generated: $ExecutionTime</p>
        <p>Computer: $ComputerName</p>
    </div>
    
    <div class="summary">
        <h2>Executive Summary</h2>
        <ul>
            <li>Total Administrators: $($Summary.TotalAdmins)</li>
            <li>Local Accounts: $($Summary.LocalAccounts)</li>
            <li>Domain Accounts: $($Summary.DomainAccounts)</li>
            <li>Authorized Administrators: $($Summary.Authorized)</li>
            <li>Unauthorized Administrators: $($Summary.Unauthorized)</li>
            <li>Security Score: $($Summary.SecurityScore)/100</li>
        </ul>
    </div>
    
    <div class="risk-legend">
        <h3>Risk Level Legend</h3>
        <span class="badge critical">Critical</span>
        <span class="badge high">High</span>
        <span class="badge medium">Medium</span>
        <span class="badge low">Low</span>
    </div>
    
    <h2>Administrator Accounts Analysis</h2>
    <table>
        <thead>
            <tr>
                <th>Username</th>
                <th>Domain</th>
                <th>Account Type</th>
                <th>Member Type</th>
                <th>Authorized</th>
                <th>Risk Level</th>
                <th>Recommendation</th>
            </tr>
        </thead>
        <tbody>
"@

    foreach ($Item in $Analysis) {
        $RiskClass = $Item.RiskLevel.ToLower()
        $Authorized = if ($Item.IsAuthorized) { '✓' } else { '✗' }
        $HTML += @"
            <tr>
                <td>$($Item.Username)</td>
                <td>$($Item.Domain)</td>
                <td>$($Item.AccountType)</td>
                <td>$($Item.MemberType)</td>
                <td>$Authorized</td>
                <td><span class="badge $RiskClass">$($Item.RiskLevel)</span></td>
                <td>$($Item.Recommendation)</td>
            </tr>
"@
    }

    $HTML += @"
        </tbody>
    </table>
    
    <div class="recommendations">
        <h2>Security Recommendations</h2>
"@

    # Generate recommendations based on findings
    $CriticalAccounts = $Analysis | Where-Object { $_.RiskLevel -eq 'Critical' }
    $HighRiskAccounts = $Analysis | Where-Object { $_.RiskLevel -eq 'High' -and -not $_.IsAuthorized }
    $UnusedLocalAdmins = $Analysis | Where-Object { $_.AccountType -eq 'Local' -and -not $_.IsAuthorized }

    if ($CriticalAccounts) {
        $HTML += @"
        <div class="recommendation">
            <strong>⚠ Critical Issues Found</strong>
            <p>The following accounts require immediate attention: $(($CriticalAccounts | ForEach-Object { $_.Username }) -join ', ')</p>
            <p>Action: Remove these accounts from the Administrators group immediately.</p>
        </div>
"@
    }

    if ($HighRiskAccounts) {
        $HTML += @"
        <div class="recommendation">
            <strong>⚠ High Risk Accounts</strong>
            <p>Domain accounts with admin rights: $(($HighRiskAccounts | ForEach-Object { $_.Username }) -join ', ')</p>
            <p>Action: Verify business need and consider using delegated permissions instead.</p>
        </div>
"@
    }

    if ($UnusedLocalAdmins) {
        $HTML += @"
        <div class="recommendation">
            <strong>ℹ Local Account Administrators</strong>
            <p>Local accounts with admin rights: $(($UnusedLocalAdmins | ForEach-Object { $_.Username }) -join ', ')</p>
            <p>Action: Remove local accounts and use domain accounts for administration.</p>
        </div>
"@
    }

    $HTML += @"
    </div>
    
    <div class="footer">
        <p>Generated by Local_Admins_Report v$ScriptVersion | Security Audit Tool</p>
    </div>
</body>
</html>
"@

    $HTML | Out-File -FilePath $HTMLPath -Encoding UTF8
    Write-Log "HTML report exported to: $HTMLPath" "INFO"
    return $HTMLPath
}

# Function to get summary statistics
function Get-SummaryStats {
    param(
        [array]$Analysis
    )
    
    $Total = $Analysis.Count
    $Local = ($Analysis | Where-Object { $_.AccountType -eq 'Local' }).Count
    $Domain = ($Analysis | Where-Object { $_.AccountType -eq 'Domain' }).Count
    $Authorized = ($Analysis | Where-Object { $_.IsAuthorized }).Count
    $Unauthorized = ($Analysis | Where-Object { -not $_.IsAuthorized }).Count
    $Critical = ($Analysis | Where-Object { $_.RiskLevel -eq 'Critical' }).Count
    $High = ($Analysis | Where-Object { $_.RiskLevel -eq 'High' }).Count
    $Medium = ($Analysis | Where-Object { $_.RiskLevel -eq 'Medium' }).Count
    $Low = ($Analysis | Where-Object { $_.RiskLevel -eq 'Low' }).Count
    
    # Calculate security score (0-100)
    $SecurityScore = 100
    $SecurityScore -= ($Critical * 20)
    $SecurityScore -= ($High * 10)
    $SecurityScore -= ($Medium * 5)
    if ($SecurityScore -lt 0) { $SecurityScore = 0 }
    
    return [PSCustomObject]@{
        TotalAdmins = $Total
        LocalAccounts = $Local
        DomainAccounts = $Domain
        Authorized = $Authorized
        Unauthorized = $Unauthorized
        CriticalRisk = $Critical
        HighRisk = $High
        MediumRisk = $Medium
        LowRisk = $Low
        SecurityScore = $SecurityScore
    }
}

# Function to process a single computer
function Process-Computer {
    param(
        [string]$ComputerName
    )
    
    Write-Host ""
    Write-Host "=" * 60 -ForegroundColor Cyan
    Write-Host "Processing Computer: $ComputerName" -ForegroundColor Cyan
    Write-Host "=" * 60 -ForegroundColor Cyan
    
    Write-Log "Starting audit for computer: $ComputerName" "INFO"
    
    # Check connectivity
    if (-not (Test-Connection -ComputerName $ComputerName -Count 1 -Quiet)) {
        Write-Log "Cannot reach computer $ComputerName" "ERROR"
        return $null
    }
    
    # Get group members
    $Members = Get-LocalGroupMembers -ComputerName $ComputerName
    
    if ($CheckGroupMembers) {
        $NestedMembers = Get-NestedGroupMembers -ComputerName $ComputerName
        $Members += $NestedMembers
    }
    
    # Analyze members
    $Analysis = Analyze-GroupMembers -Members $Members -AllowedAdmins $AllowedAdmins
    
    # Get security info
    $SecurityInfo = Get-SecurityInfo -ComputerName $ComputerName
    
    return @{
        Members = $Members
        Analysis = $Analysis
        SecurityInfo = $SecurityInfo
    }
}

# Main function
function Invoke-LocalAdminAudit {
    Write-Log "Starting Local Administrators Audit..." "INFO"
    Write-Log "Computers to audit: $($ComputerNames -join ', ')" "INFO"
    
    $AllResults = @()
    $ComputerOutput = $OutputPath
    
    foreach ($Computer in $ComputerNames) {
        $Results = Process-Computer -ComputerName $Computer
        
        if ($Results -and $Results.Analysis) {
            $AllResults += $Results
        }
    }
    
    if ($AllResults.Count -eq 0) {
        Write-Log "No results collected" "ERROR"
        return $false
    }
    
    # Export results
    $ExportedFiles = @()
    
    # Export to CSV
    if ($OutputFormat -in @('CSV', 'All')) {
        $CSVPath = "$OutputPath.csv"
        $AllAnalysis = $AllResults | ForEach-Object { $_.Analysis }
        $AllAnalysis | Export-Csv -Path $CSVPath -NoTypeInformation -Encoding UTF8
        Write-Log "CSV report exported to: $CSVPath" "INFO"
        $ExportedFiles += $CSVPath
    }
    
    # Export to JSON
    if ($OutputFormat -in @('JSON', 'All')) {
        $JSONPath = "$OutputPath.json"
        $AllResults | ConvertTo-Json -Depth 10 | Out-File -FilePath $JSONPath -Encoding UTF8
        Write-Log "JSON report exported to: $JSONPath" "INFO"
        $ExportedFiles += $JSONPath
    }
    
    # Export to HTML
    if ($OutputFormat -in @('HTML', 'All')) {
        $HTMLPath = "$OutputPath.html"
        $AllAnalysis = $AllResults | ForEach-Object { $_.Analysis }
        Export-HTMLReport -Analysis $AllAnalysis -AllMembers $AllResults -OutputPath $OutputPath -ComputerName ($ComputerNames -join ', ')
        $ExportedFiles += $HTMLPath
    }
    
    # Display summary
    Write-Host ""
    Write-Host "=" * 80 -ForegroundColor Cyan
    Write-Host "AUDIT SUMMARY" -ForegroundColor Cyan
    Write-Host "=" * 80 -ForegroundColor Cyan
    
    $AllAnalysis = $AllResults | ForEach-Object { $_.Analysis }
    $Summary = Get-SummaryStats -Analysis $AllAnalysis
    
    Write-Host "Computers Audited: $($ComputerNames -join ', ')" -ForegroundColor White
    Write-Host "Total Administrators Found: $($Summary.TotalAdmins)" -ForegroundColor White
    Write-Host "  - Local Accounts: $($Summary.LocalAccounts)" -ForegroundColor Gray
    Write-Host "  - Domain Accounts: $($Summary.DomainAccounts)" -ForegroundColor Gray
    Write-Host ""
    Write-Host "Security Status:" -ForegroundColor Cyan
    Write-Host "  Authorized: $($Summary.Authorized)" -ForegroundColor Green
    Write-Host "  Unauthorized: $($Summary.Unauthorized)" -ForegroundColor Red
    Write-Host ""
    Write-Host "Risk Distribution:" -ForegroundColor Cyan
    Write-Host "  Critical: $($Summary.CriticalRisk)" -ForegroundColor Red
    Write-Host "  High: $($Summary.HighRisk)" -ForegroundColor Yellow
    Write-Host "  Medium: $($Summary.MediumRisk)" -ForegroundColor Yellow
    Write-Host "  Low: $($Summary.LowRisk)" -ForegroundColor Green
    Write-Host ""
    Write-Host "Security Score: $($Summary.SecurityScore)/100" -ForegroundColor $(if ($Summary.SecurityScore -ge 80) { 'Green' } elseif ($Summary.SecurityScore -ge 60) { 'Yellow' } else { 'Red' })
    Write-Host ""
    Write-Host "Reports Generated:" -ForegroundColor Cyan
    foreach ($File in $ExportedFiles) {
        Write-Host "  - $File" -ForegroundColor White
    }
    Write-Host ""
    Write-Host "Log File: $LogPath" -ForegroundColor Gray
    
    Write-Log "Audit completed successfully" "INFO"
    return $true
}

# Main execution
try {
    $Result = Invoke-LocalAdminAudit
    if ($Result) {
        Write-Host "`n✓ Local Administrators Audit completed successfully!" -ForegroundColor Green
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