<#
It includes:

CSV-driven bulk group creation
Name, SamAccountName, GroupScope, GroupCategory, Path, and Description
Existing-group detection
OU/container validation
-WhatIf dry-run support
Optional -Server parameter for a specific domain controller
Per-group error handling
Creation/skip/failure summary
Comment-based help and examples

Example usage:

.\Bulk_Create_Groups_From_CSV.ps1 -CsvPath .\groups.csv -WhatIf

Then, once you're satisfied with the dry run:

.\Bulk_Create_Groups_From_CSV.ps1 -CsvPath .\groups.csv
#>

[CmdletBinding(SupportsShouldProcess)]
param (
    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string]$CsvPath,

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

try {
    # Verify the Active Directory module is available.
    if (-not (Get-Module -ListAvailable -Name ActiveDirectory)) {
        throw "The ActiveDirectory PowerShell module is not installed or available."
    }

    Import-Module ActiveDirectory -ErrorAction Stop

    # Resolve and validate the CSV path.
    $resolvedCsvPath = (Resolve-Path -LiteralPath $CsvPath -ErrorAction Stop).Path

    if (-not (Test-Path -LiteralPath $resolvedCsvPath -PathType Leaf)) {
        throw "CSV file was not found: $resolvedCsvPath"
    }

    $groups = Import-Csv -LiteralPath $resolvedCsvPath

    if (-not $groups) {
        Write-Warning "The CSV file contains no group records."
        return
    }

    # Validate the required CSV columns.
    $requiredColumns = @(
        'Name',
        'SamAccountName',
        'GroupScope',
        'GroupCategory',
        'Path',
        'Description'
    )

    $csvColumns = @($groups[0].PSObject.Properties.Name)
    $missingColumns = $requiredColumns | Where-Object { $_ -notin $csvColumns }

    if ($missingColumns) {
        throw "Missing required CSV column(s): $($missingColumns -join ', ')"
    }

    $created = 0
    $skipped = 0
    $failed = 0

    foreach ($group in $groups) {
        try {
            # Validate required values for this row.
            foreach ($property in @('Name', 'SamAccountName', 'GroupScope', 'GroupCategory')) {
                if ([string]::IsNullOrWhiteSpace([string]$group.$property)) {
                    throw "Required value '$property' is empty."
                }
            }

            $name = [string]$group.Name
            $samAccountName = [string]$group.SamAccountName
            $groupScope = [string]$group.GroupScope
            $groupCategory = [string]$group.GroupCategory
            $path = [string]$group.Path
            $description = [string]$group.Description

            # Validate allowed AD group values.
            if ($groupScope -notin @('DomainLocal', 'Global', 'Universal')) {
                throw "Invalid GroupScope '$groupScope'. Use DomainLocal, Global, or Universal."
            }

            if ($groupCategory -notin @('Security', 'Distribution')) {
                throw "Invalid GroupCategory '$groupCategory'. Use Security or Distribution."
            }

            # Build lookup parameters.
            $lookupParameters = @{
                Identity    = $samAccountName
                ErrorAction = 'SilentlyContinue'
            }

            $lookupParameters = Get-ADParameters -Parameters $lookupParameters
            $existingGroup = Get-ADGroup @lookupParameters

            if ($existingGroup) {
                Write-Host "[SKIP] Group already exists: $samAccountName" -ForegroundColor Yellow
                $skipped++
                continue
            }

            # Validate the destination path when supplied.
            if (-not [string]::IsNullOrWhiteSpace($path)) {
                $ouParameters = @{
                    Identity    = $path
                    ErrorAction = 'Stop'
                }

                $ouParameters = Get-ADParameters -Parameters $ouParameters
                $null = Get-ADObject @ouParameters
            }

            # Prepare New-ADGroup parameters.
            $newGroupParameters = @{
                Name           = $name
                SamAccountName = $samAccountName
                GroupScope     = $groupScope
                GroupCategory  = $groupCategory
                ErrorAction    = 'Stop'
            }

            if (-not [string]::IsNullOrWhiteSpace($path)) {
                $newGroupParameters['Path'] = $path
            }

            if (-not [string]::IsNullOrWhiteSpace($description)) {
                $newGroupParameters['Description'] = $description
            }

            $newGroupParameters = Get-ADParameters -Parameters $newGroupParameters

            if ($PSCmdlet.ShouldProcess(
                $samAccountName,
                "Create Active Directory group '$name'"
            )) {
                New-ADGroup @newGroupParameters

                Write-Host "[CREATED] $samAccountName" -ForegroundColor Green
                $created++
            }
        }
        catch {
            Write-Host "[FAILED] $($group.SamAccountName): $($_.Exception.Message)" -ForegroundColor Red
            $failed++
        }
    }

    Write-Host ""
    Write-Host "========== Summary ==========" -ForegroundColor Cyan
    Write-Host "Created : $created"
    Write-Host "Skipped : $skipped"
    Write-Host "Failed  : $failed"
    Write-Host "Total   : $($groups.Count)"
}
catch {
    Write-Error $_.Exception.Message
    exit 1
}
