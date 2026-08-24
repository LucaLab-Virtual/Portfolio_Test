<#
It supports:

Exporting members of one or multiple groups
Reading groups from a CSV
Direct membership export
-Recursive nested-group expansion
Detecting circular nesting to avoid infinite loops
Optional -Server for a specific domain controller
CSV output with:
Group name
Group SamAccountName
Member name
Member SamAccountName
Member type
Distinguished Name
Membership path
Error handling per group
Summary information
Proper PowerShell comment-based help
Basic usage
.\Export_Group_Members.ps1 `
    -Group "IT-Helpdesk" `
    -OutputPath .\IT-Helpdesk-Members.csv

For nested groups:

.\Export_Group_Members.ps1 `
    -Group "IT-Helpdesk" `
    -Recursive `
    -OutputPath .\IT-Helpdesk-Members.csv

For several groups:

.\Export_Group_Members.ps1 `
    -Group "IT-Helpdesk","HR-Team","Finance-Team" `
    -OutputPath .\Group-Members.csv
#>

[CmdletBinding()]
param (
    [Parameter(Mandatory = $false)]
    [string[]]$Group,

    [Parameter(Mandatory = $false)]
    [string]$CsvPath,

    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string]$OutputPath,

    [Parameter(Mandatory = $false)]
    [switch]$Recursive,

    [Parameter(Mandatory = $false)]
    [string]$Server
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Get-ADParameters {
    param (
        [hashtable]$Parameters
    )

    if ($Server) {
        $Parameters['Server'] = $Server
    }

    return $Parameters
}

function Get-GroupMembersRecursive {
    param (
        [Parameter(Mandatory = $true)]
        [Microsoft.ActiveDirectory.Management.ADGroup]$ParentGroup,

        [Parameter(Mandatory = $true)]
        [string]$RootGroupName,

        [Parameter(Mandatory = $false)]
        [string]$CurrentPath = "",

        [Parameter(Mandatory = $false)]
        [System.Collections.Generic.HashSet[string]]$VisitedGroups
    )

    $results = @()

    if (-not $VisitedGroups) {
        $VisitedGroups = [System.Collections.Generic.HashSet[string]]::new()
    }

    # Prevent infinite loops caused by circular group nesting.
    if (-not $VisitedGroups.Add($ParentGroup.DistinguishedName)) {
        return $results
    }

    $memberParameters = @{
        Identity    = $ParentGroup.DistinguishedName
        ErrorAction = 'Stop'
    }

    $memberParameters = Get-ADParameters -Parameters $memberParameters
    $members = Get-ADGroupMember @memberParameters

    foreach ($member in $members) {
        $memberPath = if ($CurrentPath) {
            "$CurrentPath -> $($member.Name)"
        }
        else {
            $member.Name
        }

        $results += [PSCustomObject]@{
            GroupName              = $RootGroupName
            GroupSamAccountName    = $ParentGroup.SamAccountName
            MemberName             = $member.Name
            MemberSamAccountName   = $member.SamAccountName
            MemberType             = $member.objectClass
            MemberDistinguishedName = $member.DistinguishedName
            MembershipPath         = $memberPath
        }

        if ($member.objectClass -eq 'group') {
            $groupParameters = @{
                Identity    = $member.DistinguishedName
                ErrorAction = 'Stop'
            }

            $groupParameters = Get-ADParameters -Parameters $groupParameters
            $nestedGroup = Get-ADGroup @groupParameters

            $results += Get-GroupMembersRecursive `
                -ParentGroup $nestedGroup `
                -RootGroupName $RootGroupName `
                -CurrentPath $memberPath `
                -VisitedGroups $VisitedGroups
        }
    }

    return $results
}

try {
    # Verify the Active Directory module.
    if (-not (Get-Module -ListAvailable -Name ActiveDirectory)) {
        throw "The ActiveDirectory PowerShell module is not installed or available."
    }

    Import-Module ActiveDirectory -ErrorAction Stop

    # Require either -Group or -CsvPath, but not both.
    if (-not $Group -and -not $CsvPath) {
        throw "Specify either -Group or -CsvPath."
    }

    if ($Group -and $CsvPath) {
        throw "Use either -Group or -CsvPath, not both."
    }

    # Build the list of groups to process.
    $groupsToProcess = @()

    if ($Group) {
        foreach ($groupName in $Group) {
            if (-not [string]::IsNullOrWhiteSpace($groupName)) {
                $groupsToProcess += $groupName.Trim()
            }
        }
    }
    else {
        $resolvedCsvPath = (Resolve-Path -LiteralPath $CsvPath -ErrorAction Stop).Path

        if (-not (Test-Path -LiteralPath $resolvedCsvPath -PathType Leaf)) {
            throw "CSV file was not found: $resolvedCsvPath"
        }

        $csvGroups = @(Import-Csv -LiteralPath $resolvedCsvPath)

        if (-not $csvGroups) {
            throw "The CSV file contains no group records."
        }

        $columns = @($csvGroups[0].PSObject.Properties.Name)

        $groupColumn = @('Group', 'GroupName', 'SamAccountName') |
            Where-Object { $_ -in $columns } |
            Select-Object -First 1

        if (-not $groupColumn) {
            throw "CSV must contain one of these columns: Group, GroupName, SamAccountName."
        }

        foreach ($row in $csvGroups) {
            if (-not [string]::IsNullOrWhiteSpace([string]$row.$groupColumn)) {
                $groupsToProcess += ([string]$row.$groupColumn).Trim()
            }
        }
    }

    if (-not $groupsToProcess) {
        throw "No valid groups were supplied."
    }

    $results = @()
    $failed = 0

    foreach ($groupIdentity in $groupsToProcess) {
        try {
            $groupParameters = @{
                Identity    = $groupIdentity
                ErrorAction = 'Stop'
            }

            $groupParameters = Get-ADParameters -Parameters $groupParameters
            $adGroup = Get-ADGroup @groupParameters

            if ($Recursive) {
                $visitedGroups = [System.Collections.Generic.HashSet[string]]::new()

                $groupResults = Get-GroupMembersRecursive `
                    -ParentGroup $adGroup `
                    -RootGroupName $adGroup.Name `
                    -VisitedGroups $visitedGroups

                $results += $groupResults
            }
            else {
                $memberParameters = @{
                    Identity    = $adGroup.DistinguishedName
                    ErrorAction = 'Stop'
                }

                $memberParameters = Get-ADParameters -Parameters $memberParameters
                $members = @(Get-ADGroupMember @memberParameters)

                foreach ($member in $members) {
                    $results += [PSCustomObject]@{
                        GroupName               = $adGroup.Name
                        GroupSamAccountName     = $adGroup.SamAccountName
                        MemberName              = $member.Name
                        MemberSamAccountName    = $member.SamAccountName
                        MemberType              = $member.objectClass
                        MemberDistinguishedName = $member.DistinguishedName
                        MembershipPath          = $member.Name
                    }
                }
            }

            Write-Host "[EXPORTED] $($adGroup.Name)" -ForegroundColor Green
        }
        catch {
            Write-Host "[FAILED] $groupIdentity : $($_.Exception.Message)" -ForegroundColor Red
            $failed++
        }
    }

    # Ensure the output directory exists.
    $outputDirectory = Split-Path -Path $OutputPath -Parent

    if ($outputDirectory -and -not (Test-Path -LiteralPath $outputDirectory)) {
        New-Item -Path $outputDirectory -ItemType Directory -Force | Out-Null
    }

    if ($results.Count -gt 0) {
        $results |
            Sort-Object GroupName, MembershipPath |
            Export-Csv -LiteralPath $OutputPath -NoTypeInformation -Encoding UTF8

        Write-Host ""
        Write-Host "Export completed: $OutputPath" -ForegroundColor Cyan
        Write-Host "Members exported : $($results.Count)"
        Write-Host "Groups processed : $($groupsToProcess.Count)"
        Write-Host "Groups failed    : $failed"
        Write-Host "Recursive mode   : $Recursive"
    }
    else {
        # Still create a valid CSV with headers when no members were found.
        @(
            [PSCustomObject]@{
                GroupName               = ''
                GroupSamAccountName     = ''
                MemberName              = ''
                MemberSamAccountName    = ''
                MemberType              = ''
                MemberDistinguishedName = ''
                MembershipPath          = ''
            }
        ) |
            Select-Object GroupName, GroupSamAccountName, MemberName,
                MemberSamAccountName, MemberType,
                MemberDistinguishedName, MembershipPath |
            Export-Csv -LiteralPath $OutputPath -NoTypeInformation -Encoding UTF8

        Write-Warning "No group members were found. An empty CSV was created."
    }

    if ($failed -gt 0) {
        exit 1
    }
}
catch {
    Write-Error $_.Exception.Message
    exit 1
}
