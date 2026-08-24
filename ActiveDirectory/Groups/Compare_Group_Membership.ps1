<#
Usage Examples:
powershell

# Basic comparison
.\Compare_Group_Membership.ps1 -GroupA "Domain Admins" -GroupB "Enterprise Admins"

# With custom export path
.\Compare_Group_Membership.ps1 -GroupA "Sales" -GroupB "Marketing" -ExportPath "C:\Reports"

# Using a specific domain controller
.\Compare_Group_Membership.ps1 -GroupA "IT Team" -GroupB "Dev Team" -DomainController "DC01.domain.com"

Sample Output Files:

The script generates 5 CSV files:

    GroupCompare_20260115_143022_Summary.csv - Statistics summary

    GroupCompare_20260115_143022_OnlyInGroupA.csv - Users only in first group

    GroupCompare_20260115_143022_OnlyInGroupB.csv - Users only in second group

    GroupCompare_20260115_143022_InBothGroups.csv - Users in both groups

    GroupCompare_20260115_143022_DetailedComparison.csv - Complete membership matrix
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory=$true, Position=0)]
    [string]$GroupA,
    
    [Parameter(Mandatory=$true, Position=1)]
    [string]$GroupB,
    
    [Parameter(Mandatory=$false)]
    [string]$ExportPath = ".",
    
    [Parameter(Mandatory=$false)]
    [string]$DomainController
)

# Function to import AD module with error handling
function Import-ADModule {
    try {
        if (-not (Get-Module -Name ActiveDirectory -ListAvailable)) {
            Write-Error "ActiveDirectory module is not installed. Please install RSAT-AD-PowerShell."
            return $false
        }
        
        Import-Module ActiveDirectory -Force -ErrorAction Stop
        Write-Host "ActiveDirectory module loaded successfully." -ForegroundColor Green
        return $true
    }
    catch {
        Write-Error "Failed to import ActiveDirectory module: $_"
        return $false
    }
}

# Function to get group members with error handling
function Get-GroupMembers {
    param(
        [string]$GroupName,
        [string]$DC
    )
    
    try {
        $params = @{
            Identity = $GroupName
            Properties = @('member', 'members')
            ErrorAction = 'Stop'
        }
        
        if ($DC) {
            $params.Server = $DC
        }
        
        $group = Get-ADGroup @params
        
        if (-not $group) {
            Write-Error "Group '$GroupName' not found."
            return $null
        }
        
        Write-Host "Getting members for group: $($group.Name)" -ForegroundColor Cyan
        
        # Get members including nested groups (recursive)
        $members = Get-ADGroupMember -Identity $group.DistinguishedName -Recursive -ErrorAction Stop
        
        # Filter only user objects (optional, but common use case)
        $users = $members | Where-Object { $_.objectClass -eq 'user' }
        
        # Get additional user details
        $userDetails = @()
        foreach ($user in $users) {
            try {
                $userParams = @{
                    Identity = $user.DistinguishedName
                    Properties = @('SamAccountName', 'DisplayName', 'UserPrincipalName', 'Enabled')
                }
                if ($DC) { $userParams.Server = $DC }
                
                $userObj = Get-ADUser @userParams
                $userDetails += [PSCustomObject]@{
                    SamAccountName = $userObj.SamAccountName
                    DisplayName = $userObj.DisplayName
                    UserPrincipalName = $userObj.UserPrincipalName
                    DistinguishedName = $userObj.DistinguishedName
                    Enabled = $userObj.Enabled
                }
            }
            catch {
                Write-Warning "Could not retrieve details for user: $($user.DistinguishedName)"
            }
        }
        
        Write-Host "Found $($userDetails.Count) members in group: $($group.Name)" -ForegroundColor Green
        return $userDetails
        
    }
    catch {
        Write-Error "Error retrieving group '$GroupName': $_"
        return $null
    }
}

# Main script execution
function Main {
    Write-Host "="*60 -ForegroundColor Yellow
    Write-Host "AD Group Membership Comparison Tool" -ForegroundColor Yellow
    Write-Host "="*60 -ForegroundColor Yellow
    Write-Host ""
    
    # Check AD module
    if (-not (Import-ADModule)) {
        Write-Error "Cannot proceed without ActiveDirectory module."
        return
    }
    
    # Validate export path
    if (-not (Test-Path $ExportPath)) {
        try {
            New-Item -Path $ExportPath -ItemType Directory -Force -ErrorAction Stop | Out-Null
            Write-Host "Created export directory: $ExportPath" -ForegroundColor Green
        }
        catch {
            Write-Error "Cannot create export directory: $_"
            return
        }
    }
    
    # Get members of both groups
    Write-Host "Retrieving group memberships..." -ForegroundColor Cyan
    Write-Host ""
    
    $membersA = Get-GroupMembers -GroupName $GroupA -DC $DomainController
    if ($null -eq $membersA) { 
        Write-Error "Failed to retrieve members for Group A. Exiting."
        return 
    }
    
    $membersB = Get-GroupMembers -GroupName $GroupB -DC $DomainController
    if ($null -eq $membersB) { 
        Write-Error "Failed to retrieve members for Group B. Exiting."
        return 
    }
    
    Write-Host ""
    
    # Create hashsets for comparison
    $samA = [System.Collections.Generic.HashSet[string]]::new([string[]]($membersA.SamAccountName))
    $samB = [System.Collections.Generic.HashSet[string]]::new([string[]]($membersB.SamAccountName))
    
    # Find differences
    $onlyInA = $membersA | Where-Object { -not $samB.Contains($_.SamAccountName) }
    $onlyInB = $membersB | Where-Object { -not $samA.Contains($_.SamAccountName) }
    
    # Find intersection
    $inBoth = $membersA | Where-Object { $samB.Contains($_.SamAccountName) }
    
    # Prepare timestamp for filename
    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $exportPrefix = "GroupCompare_$timestamp"
    
    # Create results objects for export
    $exportResults = @{
        "OnlyInGroupA" = $onlyInA | Select-Object SamAccountName, DisplayName, UserPrincipalName, Enabled
        "OnlyInGroupB" = $onlyInB | Select-Object SamAccountName, DisplayName, UserPrincipalName, Enabled
        "InBothGroups" = $inBoth | Select-Object SamAccountName, DisplayName, UserPrincipalName, Enabled
    }
    
    # Generate summary
    $summary = [PSCustomObject]@{
        ComparisonDate = Get-Date
        GroupA = $GroupA
        GroupB = $GroupB
        MembersInGroupA = $membersA.Count
        MembersInGroupB = $membersB.Count
        MembersInBoth = $inBoth.Count
        MembersOnlyInGroupA = $onlyInA.Count
        MembersOnlyInGroupB = $onlyInB.Count
        DomainController = if ($DomainController) { $DomainController } else { "Default" }
        ExportPath = $ExportPath
    }
    
    # Export to CSV files
    Write-Host "Exporting comparison results..." -ForegroundColor Cyan
    
    # Export summary
    $summaryPath = Join-Path $ExportPath "$exportPrefix`_Summary.csv"
    $summary | Export-Csv -Path $summaryPath -NoTypeInformation
    Write-Host "  ✓ Summary: $summaryPath" -ForegroundColor Green
    
    # Export only in Group A
    if ($onlyInA.Count -gt 0) {
        $pathA = Join-Path $ExportPath "$exportPrefix`_OnlyInGroupA.csv"
        $exportResults.OnlyInGroupA | Export-Csv -Path $pathA -NoTypeInformation
        Write-Host "  ✓ Only in Group A ($($onlyInA.Count) users): $pathA" -ForegroundColor Green
    }
    else {
        Write-Host "  ℹ No users only in Group A" -ForegroundColor Yellow
    }
    
    # Export only in Group B
    if ($onlyInB.Count -gt 0) {
        $pathB = Join-Path $ExportPath "$exportPrefix`_OnlyInGroupB.csv"
        $exportResults.OnlyInGroupB | Export-Csv -Path $pathB -NoTypeInformation
        Write-Host "  ✓ Only in Group B ($($onlyInB.Count) users): $pathB" -ForegroundColor Green
    }
    else {
        Write-Host "  ℹ No users only in Group B" -ForegroundColor Yellow
    }
    
    # Export in both
    if ($inBoth.Count -gt 0) {
        $pathBoth = Join-Path $ExportPath "$exportPrefix`_InBothGroups.csv"
        $exportResults.InBothGroups | Export-Csv -Path $pathBoth -NoTypeInformation
        Write-Host "  ✓ In both groups ($($inBoth.Count) users): $pathBoth" -ForegroundColor Green
    }
    else {
        Write-Host "  ℹ No users in both groups" -ForegroundColor Yellow
    }
    
    # Export detailed comparison report
    $detailPath = Join-Path $ExportPath "$exportPrefix`_DetailedComparison.csv"
    $detailReport = @()
    
    # Add all users with membership status
    $allUsers = @{}
    foreach ($user in $membersA) {
        $allUsers[$user.SamAccountName] = @{
            User = $user
            InGroupA = $true
            InGroupB = $false
        }
    }
    foreach ($user in $membersB) {
        if ($allUsers.ContainsKey($user.SamAccountName)) {
            $allUsers[$user.SamAccountName].InGroupB = $true
        }
        else {
            $allUsers[$user.SamAccountName] = @{
                User = $user
                InGroupA = $false
                InGroupB = $true
            }
        }
    }
    
    foreach ($entry in $allUsers.Values) {
        $status = if ($entry.InGroupA -and $entry.InGroupB) { "In Both" }
                  elseif ($entry.InGroupA) { "Only in A" }
                  else { "Only in B" }
        
        $detailReport += [PSCustomObject]@{
            SamAccountName = $entry.User.SamAccountName
            DisplayName = $entry.User.DisplayName
            UserPrincipalName = $entry.User.UserPrincipalName
            Enabled = $entry.User.Enabled
            InGroupA = $entry.InGroupA
            InGroupB = $entry.InGroupB
            MembershipStatus = $status
        }
    }
    
    $detailReport | Export-Csv -Path $detailPath -NoTypeInformation
    Write-Host "  ✓ Detailed comparison: $detailPath" -ForegroundColor Green
    
    # Display summary on console
    Write-Host ""
    Write-Host "="*60 -ForegroundColor Yellow
    Write-Host "COMPARISON SUMMARY" -ForegroundColor Yellow
    Write-Host "="*60 -ForegroundColor Yellow
    Write-Host "Group A: $GroupA - $($membersA.Count) members" -ForegroundColor Cyan
    Write-Host "Group B: $GroupB - $($membersB.Count) members" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "Users in both groups: $($inBoth.Count)" -ForegroundColor Green
    Write-Host "Users only in Group A: $($onlyInA.Count)" -ForegroundColor Yellow
    Write-Host "Users only in Group B: $($onlyInB.Count)" -ForegroundColor Yellow
    Write-Host ""
    Write-Host "Export files saved to: $ExportPath" -ForegroundColor Green
    Write-Host "Completed at: $(Get-Date)" -ForegroundColor Green
    Write-Host "="*60 -ForegroundColor Yellow
}

# Run the main function
Main