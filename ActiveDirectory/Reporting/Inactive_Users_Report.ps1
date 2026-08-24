<#
.SYNOPSIS
    Generates a report of inactive Active Directory users based on last logon timestamp.
.DESCRIPTION
    This script identifies users who haven't logged in for a specified number of days.
    It checks both lastLogon and lastLogonTimestamp attributes for accuracy.
    Results are exported to CSV and optionally HTML format.
.PARAMETER InactiveDays
    Number of days since last logon to consider a user inactive. Default is 90 days.
.PARAMETER SearchBase
    Distinguished Name of the OU to search. Defaults to entire domain.
.PARAMETER ExportPath
    Path where the report files will be saved. Defaults to current directory.
.PARAMETER IncludeServiceAccounts
    Switch to include service accounts in the report (accounts with $ in name).
.PARAMETER GenerateHTML
    Switch to generate an HTML report in addition to CSV.
.EXAMPLE
    .\Inactive_Users_Report.ps1 -InactiveDays 60 -ExportPath "C:\Reports"
.EXAMPLE
    .\Inactive_Users_Report.ps1 -InactiveDays 30 -SearchBase "OU=Users,DC=domain,DC=com" -GenerateHTML
.NOTES
    Author: PowerShell Portfolio
    Version: 1.0
    Requires: ActiveDirectory module
#>

<#
Exports to CSV (and optionally HTML) to avoid large console output

Checks both lastLogon and lastLogonTimestamp for accuracy across domain controllers

Identifies users who have never logged in based on creation date

Includes comprehensive filtering options (inactive days, OU scope, service accounts)

Provides summary statistics in console without flooding it with user details

Generates timestamped report files to prevent overwriting

Includes progress indicators during processing

Formats the CSV with all relevant user attributes

Optional HTML report for better readability

Displays top 10 inactive users in console for quick reference
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [int]$InactiveDays = 90,
    
    [Parameter(Mandatory = $false)]
    [string]$SearchBase,
    
    [Parameter(Mandatory = $false)]
    [string]$ExportPath = ".",
    
    [Parameter(Mandatory = $false)]
    [switch]$IncludeServiceAccounts,
    
    [Parameter(Mandatory = $false)]
    [switch]$GenerateHTML
)

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

# Calculate the cutoff date
$CutoffDate = (Get-Date).AddDays(-$InactiveDays)
Write-Host "Generating report for users inactive since: $CutoffDate" -ForegroundColor Cyan

# Build the filter
$Filter = "Enabled -eq 'True'"
if (!$IncludeServiceAccounts) {
    $Filter += " -and (Name -notlike '*$')"
}

Write-Host "Retrieving Active Directory users..." -ForegroundColor Yellow

# Get all enabled users with their logon attributes
$Params = @{
    Filter = $Filter
    Properties = @(
        'lastLogon',
        'lastLogonTimestamp',
        'whenCreated',
        'DisplayName',
        'Title',
        'Department',
        'Manager',
        'Mail',
        'EmployeeID'
    )
    ErrorAction = 'Stop'
}

if ($SearchBase) {
    $Params.SearchBase = $SearchBase
}

try {
    $Users = Get-ADUser @Params
    Write-Host "Found $($Users.Count) enabled users to evaluate." -ForegroundColor Green
}
catch {
    Write-Error "Failed to retrieve users: $_"
    exit 1
}

# Create an array to store results
$Results = @()
$TotalUsers = $Users.Count
$Current = 0

Write-Host "Processing users..." -ForegroundColor Yellow

foreach ($User in $Users) {
    $Current++
    $Progress = [math]::Round(($Current / $TotalUsers) * 100, 0)
    Write-Progress -Activity "Analyzing User Logon Activity" -Status "Processing $($User.Name)" -PercentComplete $Progress
    
    # Convert lastLogon (logon to specific DC) to DateTime
    $LastLogon = $null
    if ($User.lastLogon -and $User.lastLogon -gt 0) {
        $LastLogon = [datetime]::FromFileTime($User.lastLogon)
    }
    
    # Convert lastLogonTimestamp (replicated across DCs) to DateTime
    $LastLogonTimestamp = $null
    if ($User.lastLogonTimestamp -and $User.lastLogonTimestamp -gt 0) {
        $LastLogonTimestamp = [datetime]::FromFileTime($User.lastLogonTimestamp)
    }
    
    # Determine the most recent logon
    $MostRecentLogon = $null
    if ($LastLogon -and $LastLogonTimestamp) {
        $MostRecentLogon = [datetime]::Compare($LastLogon, $LastLogonTimestamp) -ge 0 ? $LastLogon : $LastLogonTimestamp
    }
    elseif ($LastLogon) {
        $MostRecentLogon = $LastLogon
    }
    elseif ($LastLogonTimestamp) {
        $MostRecentLogon = $LastLogonTimestamp
    }
    
    # Check if user has never logged in
    $NeverLoggedIn = $false
    if (-not $MostRecentLogon) {
        $NeverLoggedIn = $true
    }
    
    # Check if user is inactive
    $IsInactive = $false
    $DaysSinceLogon = $null
    
    if ($MostRecentLogon) {
        $DaysSinceLogon = (Get-Date) - $MostRecentLogon
        if ($DaysSinceLogon.TotalDays -gt $InactiveDays) {
            $IsInactive = $true
        }
    }
    else {
        # User has never logged in; check creation date as fallback
        $DaysSinceCreation = (Get-Date) - $User.whenCreated
        if ($DaysSinceCreation.TotalDays -gt $InactiveDays) {
            $IsInactive = $true
            $NeverLoggedIn = $true
        }
    }
    
    # Add to results if inactive
    if ($IsInactive) {
        $Results += [PSCustomObject]@{
            Name               = $User.Name
            SamAccountName     = $User.SamAccountName
            DisplayName        = $User.DisplayName
            UserPrincipalName  = $User.UserPrincipalName
            Email              = $User.Mail
            Title              = $User.Title
            Department         = $User.Department
            EmployeeID         = $User.EmployeeID
            Manager            = if ($User.Manager) { (Get-ADUser -Identity $User.Manager -ErrorAction SilentlyContinue).Name } else { $null }
            Enabled            = $User.Enabled
            Created            = $User.whenCreated
            LastLogon          = $MostRecentLogon
            DaysSinceLogon     = if ($DaysSinceLogon) { [math]::Round($DaysSinceLogon.TotalDays, 1) } else { "Never logged in" }
            NeverLoggedIn      = $NeverLoggedIn
            DistinguishedName  = $User.DistinguishedName
        }
    }
}

Write-Progress -Activity "Analyzing User Logon Activity" -Completed

# Sort results by days since logon (descending - most inactive first)
$Results = $Results | Sort-Object -Property DaysSinceLogon -Descending

# Generate CSV report
$Timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$CSVFileName = "InactiveUsers_Report_$Timestamp.csv"
$CSVPath = Join-Path -Path $ExportPath -ChildPath $CSVFileName
$Results | Export-Csv -Path $CSVPath -NoTypeInformation -Encoding UTF8

Write-Host "`nCSV Report saved to: $CSVPath" -ForegroundColor Green

# Generate summary statistics
$TotalInactive = $Results.Count
$NeverLoggedInCount = ($Results | Where-Object { $_.NeverLoggedIn -eq $true }).Count
$OldestInactive = if ($Results.Count -gt 0) { ($Results | Select-Object -First 1).Name } else { "N/A" }

# Display summary
Write-Host "`n===== REPORT SUMMARY =====" -ForegroundColor Cyan
Write-Host "Total enabled users evaluated: $TotalUsers" -ForegroundColor White
Write-Host "Total inactive users found: $TotalInactive" -ForegroundColor Yellow
Write-Host "Users who never logged in: $NeverLoggedInCount" -ForegroundColor Yellow
Write-Host "Oldest inactive user: $OldestInactive" -ForegroundColor Yellow
Write-Host "Inactive threshold: $InactiveDays days" -ForegroundColor White
Write-Host "=========================`n" -ForegroundColor Cyan

# Generate HTML report if requested
if ($GenerateHTML -and $Results.Count -gt 0) {
    $HTMLFileName = "InactiveUsers_Report_$Timestamp.html"
    $HTMLPath = Join-Path -Path $ExportPath -ChildPath $HTMLFileName
    
    $HTMLHeader = @"
<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <title>Inactive Active Directory Users Report</title>
    <style>
        body { font-family: Arial, sans-serif; margin: 20px; }
        h1 { color: #2c3e50; border-bottom: 2px solid #3498db; padding-bottom: 10px; }
        h2 { color: #34495e; margin-top: 20px; }
        .summary { background-color: #ecf0f1; padding: 15px; border-radius: 5px; margin-bottom: 20px; }
        table { border-collapse: collapse; width: 100%; margin-top: 20px; }
        th { background-color: #3498db; color: white; padding: 12px; text-align: left; }
        td { padding: 10px; border-bottom: 1px solid #ddd; }
        tr:hover { background-color: #f5f5f5; }
        .inactive { background-color: #ffe6e6; }
        .never-logged { background-color: #fff3cd; }
        .critical { background-color: #f8d7da; }
        .footer { margin-top: 30px; font-size: 0.9em; color: #7f8c8d; }
    </style>
</head>
<body>
"@

    $HTMLSummary = @"
    <h1>Inactive Active Directory Users Report</h1>
    <div class="summary">
        <p><strong>Generated:</strong> $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')</p>
        <p><strong>Inactive Threshold:</strong> $InactiveDays days</p>
        <p><strong>Total Inactive Users:</strong> $TotalInactive</p>
        <p><strong>Never Logged In:</strong> $NeverLoggedInCount</p>
        <p><strong>Domain:</strong> $(Get-ADDomain | Select-Object -ExpandProperty DNSRoot)</p>
    </div>
"@

    # Build HTML table rows
    $HTMLRows = ""
    foreach ($User in $Results) {
        $RowClass = if ($User.NeverLoggedIn) { "never-logged" } elseif ($User.DaysSinceLogon -gt 180) { "critical" } else { "inactive" }
        $HTMLRows += @"
        <tr class="$RowClass">
            <td>$($User.Name)</td>
            <td>$($User.SamAccountName)</td>
            <td>$($User.Title)</td>
            <td>$($User.Department)</td>
            <td>$($User.Email)</td>
            <td>$($User.LastLogon)</td>
            <td>$($User.DaysSinceLogon)</td>
            <td>$($User.NeverLoggedIn)</td>
        </tr>
"@
    }

    $HTMLTable = @"
    <h2>Inactive Users List</h2>
    <table>
        <thead>
            <tr>
                <th>Name</th>
                <th>SamAccountName</th>
                <th>Title</th>
                <th>Department</th>
                <th>Email</th>
                <th>Last Logon</th>
                <th>Days Since Logon</th>
                <th>Never Logged In</th>
            </tr>
        </thead>
        <tbody>
            $HTMLRows
        </tbody>
    </table>
"@

    $HTMLFooter = @"
    <div class="footer">
        <p>Report generated by Inactive_Users_Report.ps1 | $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')</p>
    </div>
</body>
</html>
"@

    $HTMLContent = $HTMLHeader + $HTMLSummary + $HTMLTable + $HTMLFooter
    $HTMLContent | Out-File -FilePath $HTMLPath -Encoding UTF8
    
    Write-Host "HTML Report saved to: $HTMLPath" -ForegroundColor Green
}

# Display the top 10 inactive users if any exist
if ($Results.Count -gt 0) {
    Write-Host "Top 10 Most Inactive Users:" -ForegroundColor Yellow
    $Results | Select-Object -First 10 | Format-Table -Property Name, SamAccountName, DaysSinceLogon, NeverLoggedIn -AutoSize
}

Write-Host "Script execution completed successfully!" -ForegroundColor Green