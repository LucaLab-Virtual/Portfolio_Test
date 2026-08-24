<#
.SYNOPSIS
    Links Group Policy Objects (GPOs) to Organizational Units (OUs) with advanced management options.
.DESCRIPTION
    This script creates, manages, and removes GPO links to OUs. Supports bulk operations,
    order management, enforcement, and comprehensive reporting.
.PARAMETER GPOName
    Name of the GPO to link (supports wildcards)
.PARAMETER GPOList
    Array of GPO names to link
.PARAMETER TargetOU
    OU path to link the GPO(s) to
.PARAMETER TargetOUList
    Array of OU paths to link GPOs to
.PARAMETER CSVFile
    Path to CSV file containing link configurations
.PARAMETER LinkOrder
    Link order priority (1 = highest priority)
.PARAMETER Enforced
    Enforces the GPO link (cannot be blocked at lower levels)
.PARAMETER Disabled
    Creates the link in disabled state
.PARAMETER RemoveLink
    Removes the GPO link instead of creating it
.PARAMETER ReplaceLinks
    Replaces all existing links on the OU with the new ones
.PARAMETER UpdateExisting
    Updates existing link properties (order, enforced status)
.PARAMETER NoReplacement
    Doesn't replace existing links (adds to existing ones)
.PARAMETER BackupBeforeChange
    Creates backup of GPOs before making changes
.PARAMETER Domain
    Target domain for operations
.PARAMETER Force
    Skips confirmation prompts
.PARAMETER WhatIf
    Shows what would happen without performing the operation
.EXAMPLE
    .\Link_GPO_To_OU.ps1 -GPOName "Workstation Policy" -TargetOU "OU=Workstations,DC=domain,DC=com"
    Links a single GPO to an OU
.EXAMPLE
    .\Link_GPO_To_OU.ps1 -GPOList "Policy1","Policy2" -TargetOU "OU=Computers,DC=domain,DC=com" -LinkOrder 1 -Enforced
    Links multiple GPOs to an OU with enforcement and order
.EXAMPLE
    .\Link_GPO_To_OU.ps1 -TargetOU "OU=Workstations,DC=domain,DC=com" -RemoveLink -GPOName "Old Policy"
    Removes a GPO link from an OU
.EXAMPLE
    .\Link_GPO_To_OU.ps1 -CSVFile "C:\GPOS\link_config.csv" -ReplaceLinks
    Mass configure GPO links from CSV file
.NOTES
    Requires Group Policy Management Console (GPMC) module
    Must be run with administrative privileges
    Author: Portfolio Script
    Date: $(Get-Date -Format "yyyy-MM-dd")
#>

<#
Sample CSV File Template

Create a CSV file with the following format for bulk link management:
csv

GPOName,TargetOU,LinkOrder,Enforced,Disabled,Operation
"Workstation Policy","OU=Workstations,DC=domain,DC=com",1,TRUE,FALSE,Link
"Server Policy","OU=Servers,DC=domain,DC=com",1,TRUE,FALSE,Link
"User Policy","OU=Users,DC=domain,DC=com",1,FALSE,FALSE,Link
"Old Policy","OU=Workstations,DC=domain,DC=com",,,,Remove
"AppLocker Policy","OU=Computers,DC=domain,DC=com",2,FALSE,FALSE,Update
"Security Policy","OU=Workstations,DC=domain,DC=com",2,TRUE,FALSE,Link

Key Features:

    Multiple Link Methods:

        Single GPO to single OU

        Multiple GPOs to single OU

        Single GPO to multiple OUs

        CSV import for bulk operations

    Link Management:

        Create new links

        Update existing links

        Remove links

        Replace all links on an OU

    Advanced Options:

        Link order priority management

        Enforcement (prevent blocking)

        Disabled links

        Backup before changes

    Comprehensive Reporting:

        Shows current links before changes

        Detailed operation logs

        Success/failure tracking

        Color-coded output

    Safety Features:

        OU validation

        GPO existence checks

        Confirmation prompts

        WhatIf preview mode

        Backup capability

Usage Examples:
powershell

# Link single GPO to OU
.\Link_GPO_To_OU.ps1 -GPOName "Workstation Policy" -TargetOU "OU=Workstations,DC=domain,DC=com"

# Link multiple GPOs with enforcement and order
.\Link_GPO_To_OU.ps1 -GPOList "Policy1","Policy2","Policy3" -TargetOU "OU=Computers,DC=domain,DC=com" -LinkOrder 1 -Enforced

# Remove GPO link
.\Link_GPO_To_OU.ps1 -GPOName "Old Policy" -TargetOU "OU=Workstations,DC=domain,DC=com" -RemoveLink

# Replace all links on an OU
.\Link_GPO_To_OU.ps1 -GPOList "NewPolicy1","NewPolicy2" -TargetOU "OU=Servers,DC=domain,DC=com" -ReplaceLinks -Enforced

# Update existing link properties
.\Link_GPO_To_OU.ps1 -GPOName "Existing Policy" -TargetOU "OU=Computers,DC=domain,DC=com" -UpdateExisting -Enforced -LinkOrder 1

# Bulk configure from CSV
.\Link_GPO_To_OU.ps1 -CSVFile "C:\GPOS\link_config.csv" -Force

# Preview changes
.\Link_GPO_To_OU.ps1 -GPOName "New Policy" -TargetOU "OU=Computers,DC=domain,DC=com" -WhatIf

CSV Column Descriptions:
Column		Description			Values
GPOName		Name of the GPO			String
TargetOU	OU Distinguished Name		String
LinkOrder	Link priority (1=highest)	Integer
Enforced	Enforce link			TRUE/FALSE
Disabled	Disable link			TRUE/FALSE
Operation	Type of operation		Link/Remove/Update
#>

[CmdletBinding(DefaultParameterSetName = 'Single')]
param(
    [Parameter(ParameterSetName = 'Single', Mandatory = $true)]
    [Parameter(ParameterSetName = 'Remove', Mandatory = $true)]
    [string]$GPOName,
    
    [Parameter(ParameterSetName = 'Multiple')]
    [string[]]$GPOList,
    
    [Parameter(ParameterSetName = 'Single', Mandatory = $true)]
    [Parameter(ParameterSetName = 'Multiple', Mandatory = $true)]
    [Parameter(ParameterSetName = 'Remove', Mandatory = $true)]
    [string]$TargetOU,
    
    [Parameter(ParameterSetName = 'BulkMultiple')]
    [string[]]$TargetOUList,
    
    [Parameter(ParameterSetName = 'CSVImport')]
    [ValidateScript({Test-Path $_})]
    [string]$CSVFile,
    
    [Parameter()]
    [int]$LinkOrder = 1,
    
    [Parameter()]
    [switch]$Enforced,
    
    [Parameter()]
    [switch]$Disabled,
    
    [Parameter(ParameterSetName = 'Remove')]
    [switch]$RemoveLink,
    
    [Parameter()]
    [switch]$ReplaceLinks,
    
    [Parameter()]
    [switch]$UpdateExisting,
    
    [Parameter()]
    [switch]$NoReplacement,
    
    [Parameter()]
    [switch]$BackupBeforeChange,
    
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
$script:LinkedCount = 0
$script:UpdatedCount = 0
$script:RemovedCount = 0
$script:FailedCount = 0
$script:SkippedCount = 0
$script:Operations = @()

# Function to write colored output
function Write-OutputColor {
    param(
        [string]$Message,
        [string]$Color = "White"
    )
    Write-Host "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] " -NoNewline -ForegroundColor Gray
    Write-Host $Message -ForegroundColor $Color
}

# Function to create a progress bar
function Write-ProgressBar {
    param(
        [int]$Current,
        [int]$Total,
        [string]$Activity = "Processing"
    )
    
    $percent = if ($Total -gt 0) { [math]::Round(($Current / $Total) * 100) } else { 0 }
    Write-Progress -Activity $Activity -Status "Progress: $percent%" -PercentComplete $percent
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

# Function to validate OU path
function Test-OUPath {
    param([string]$OUPath)
    
    try {
        $ou = Get-ADOrganizationalUnit -Identity $OUPath -ErrorAction SilentlyContinue
        return $ou -ne $null
    } catch {
        return $false
    }
}

# Function to get OU name from DN
function Get-OUSimpleName {
    param([string]$OUPath)
    
    $parts = $OUPath -split ','
    $ouPart = $parts | Where-Object { $_ -match '^OU=' }
    if ($ouPart) {
        return $ouPart -replace '^OU=', ''
    }
    return $OUPath
}

# Function to check if GPO exists
function Test-GPOExists {
    param([string]$Name)
    
    try {
        $gpo = Get-GPO -Name $Name -ErrorAction SilentlyContinue
        return $gpo -ne $null
    } catch {
        return $false
    }
}

# Function to get current GPO links on OU
function Get-GPOLinksForOU {
    param([string]$OUPath)
    
    try {
        $ou = Get-ADOrganizationalUnit -Identity $OUPath -Properties gPLink, gPOptions -ErrorAction Stop
        
        $links = @()
        if ($ou.gPLink) {
            $linkString = $ou.gPLink
            $linkParts = $linkString -split ';' | Where-Object { $_ -match '^LDAP://' }
            
            foreach ($part in $linkParts) {
                $match = [regex]::Match($part, 'LDAP://(?:[^,]+),?([^,]+),?([^,]+)?')
                if ($match.Success) {
                    $gpoDN = $match.Value -replace '^LDAP://', ''
                    $options = $part -split ';' | Where-Object { $_ -match '^[0-9]+$' }
                    $enforced = ($options -contains '1')
                    
                    try {
                        $gpo = Get-GPO -Name $gpoDN -ErrorAction SilentlyContinue
                        if ($gpo) {
                            $links += @{
                                GPOName = $gpo.DisplayName
                                GPOId = $gpo.Id
                                GPOADPath = $gpoDN
                                Enforced = $enforced
                                Order = $links.Count + 1
                            }
                        }
                    } catch {
                        # GPO might not exist or be accessible
                        $links += @{
                            GPOName = $gpoDN
                            GPOId = $null
                            GPOADPath = $gpoDN
                            Enforced = $enforced
                            Order = $links.Count + 1
                        }
                    }
                }
            }
        }
        
        return $links
        
    } catch {
        Write-OutputColor "Failed to get GPO links for OU '$OUPath': $($_.Exception.Message)" "Yellow"
        return @()
    }
}

# Function to display current links
function Show-CurrentLinks {
    param(
        [string]$OUPath,
        $Links
    )
    
    Write-OutputColor "`nCurrent GPO Links for OU: $(Get-OUSimpleName -OUPath $OUPath)" "Cyan"
    Write-OutputColor ("=" * 80) "Gray"
    
    if ($Links.Count -eq 0) {
        Write-OutputColor "  No GPOs linked to this OU" "Yellow"
    } else {
        Write-OutputColor ("{0,-5} {1,-40} {2,-12} {3,-10}" -f "Order", "GPO Name", "Enforced", "Status") "Yellow"
        Write-OutputColor ("-" * 80) "Gray"
        
        foreach ($link in $Links) {
            $status = "Active"
            $order = $link.Order
            $name = if ($link.GPOName.Length -gt 38) { $link.GPOName.Substring(0, 38) + "..." } else { $link.GPOName }
            $enforced = if ($link.Enforced) { "Yes" } else { "No" }
            Write-OutputColor ("{0,-5} {1,-40} {2,-12} {3,-10}" -f $order, $name, $enforced, $status) "White"
        }
    }
    
    Write-OutputColor ("=" * 80) "Gray"
}

# Function to backup GPO
function Backup-GPOForLink {
    param([string]$GPOName)
    
    try {
        $backupPath = Join-Path $env:TEMP "GPO_Backup_$(Get-Date -Format 'yyyyMMdd_HHmmss')"
        New-Item -Path $backupPath -ItemType Directory -Force | Out-Null
        
        Backup-GPO -Name $GPOName -Path $backupPath -ErrorAction Stop
        Write-OutputColor "  → Backup created: $backupPath" "Gray"
        return $backupPath
    } catch {
        Write-OutputColor "  → Failed to backup GPO: $($_.Exception.Message)" "Yellow"
        return $null
    }
}

# Function to link GPO to OU
function Add-GPOLink {
    param(
        [string]$GPOName,
        [string]$OUPath,
        [int]$Order = 1,
        [bool]$Enforced = $false,
        [bool]$Disabled = $false,
        [bool]$ReplaceExisting = $false,
        [bool]$UpdateExistingLink = $false
    )
    
    try {
        # Validate GPO exists
        if (-not (Test-GPOExists -Name $GPOName)) {
            Write-OutputColor "  → GPO not found: $GPOName" "Red"
            return $false
        }
        
        # Validate OU exists
        if (-not (Test-OUPath -OUPath $OUPath)) {
            Write-OutputColor "  → OU not found: $OUPath" "Red"
            return $false
        }
        
        # Get current links
        $currentLinks = Get-GPOLinksForOU -OUPath $OUPath
        
        # Check if GPO is already linked
        $existingLink = $currentLinks | Where-Object { $_.GPOName -eq $GPOName }
        
        if ($existingLink) {
            if ($UpdateExistingLink) {
                # Update existing link
                $linkParams = @{
                    Name = $GPOName
                    Target = $OUPath
                    ErrorAction = "Stop"
                }
                
                if ($Order -gt 0) {
                    $linkParams.Order = $Order
                }
                
                if ($Enforced) {
                    $linkParams.Enforced = $true
                }
                
                if ($Disabled) {
                    $linkParams.GPOStatus = "Disabled"
                }
                
                Set-GPLink @linkParams
                Write-OutputColor "  → Updated existing link: $GPOName" "Yellow"
                $script:UpdatedCount++
                return $true
            }
            
            Write-OutputColor "  → GPO already linked to this OU: $GPOName" "Yellow"
            $script:SkippedCount++
            return $false
        }
        
        # Create new link
        $linkParams = @{
            Name = $GPOName
            Target = $OUPath
            ErrorAction = "Stop"
        }
        
        if ($Order -gt 0) {
            $linkParams.Order = $Order
        }
        
        if ($Enforced) {
            $linkParams.Enforced = $true
        }
        
        if ($Disabled) {
            $linkParams.GPOStatus = "Disabled"
        }
        
        New-GPLink @linkParams
        Write-OutputColor "  ✓ Successfully linked GPO: $GPOName" "Green"
        $script:LinkedCount++
        
        # Log operation
        $script:Operations += @{
            Operation = "Link"
            GPOName = $GPOName
            OUPath = $OUPath
            Order = $Order
            Enforced = $Enforced
            Disabled = $Disabled
        }
        
        return $true
        
    } catch {
        Write-OutputColor "  ✗ Failed to link GPO: $GPOName" "Red"
        Write-OutputColor "  → Error: $($_.Exception.Message)" "Red"
        $script:FailedCount++
        return $false
    }
}

# Function to remove GPO link
function Remove-GPOLinkFromOU {
    param(
        [string]$GPOName,
        [string]$OUPath
    )
    
    try {
        # Validate OU exists
        if (-not (Test-OUPath -OUPath $OUPath)) {
            Write-OutputColor "  → OU not found: $OUPath" "Red"
            return $false
        }
        
        # Get current links
        $currentLinks = Get-GPOLinksForOU -OUPath $OUPath
        
        # Check if GPO is linked
        $existingLink = $currentLinks | Where-Object { $_.GPOName -eq $GPOName }
        
        if (-not $existingLink) {
            Write-OutputColor "  → GPO not linked to this OU: $GPOName" "Yellow"
            $script:SkippedCount++
            return $false
        }
        
        # Backup if requested
        if ($BackupBeforeChange) {
            Backup-GPOForLink -GPOName $GPOName
        }
        
        # Remove link
        Remove-GPLink -Name $GPOName -Target $OUPath -ErrorAction Stop
        Write-OutputColor "  ✓ Successfully removed link: $GPOName" "Green"
        $script:RemovedCount++
        
        # Log operation
        $script:Operations += @{
            Operation = "Remove"
            GPOName = $GPOName
            OUPath = $OUPath
        }
        
        return $true
        
    } catch {
        Write-OutputColor "  ✗ Failed to remove link: $GPOName" "Red"
        Write-OutputColor "  → Error: $($_.Exception.Message)" "Red"
        $script:FailedCount++
        return $false
    }
}

# Function to replace all links on OU
function Replace-GPOLinks {
    param(
        [string[]]$GPONames,
        [string]$OUPath,
        [bool]$Enforced = $false,
        [bool]$Disabled = $false
    )
    
    try {
        # Get current links
        $currentLinks = Get-GPOLinksForOU -OUPath $OUPath
        
        if ($currentLinks.Count -gt 0) {
            Write-OutputColor "  → Removing existing links ($($currentLinks.Count) found)" "Yellow"
            
            foreach ($link in $currentLinks) {
                try {
                    Remove-GPLink -Name $link.GPOName -Target $OUPath -ErrorAction Stop
                    Write-OutputColor "    → Removed link: $($link.GPOName)" "Gray"
                } catch {
                    Write-OutputColor "    → Failed to remove link: $($link.GPOName)" "Yellow"
                }
            }
        }
        
        # Add new links
        Write-OutputColor "  → Adding new links ($($GPONames.Count) GPOs)" "Cyan"
        $order = 1
        
        foreach ($name in $GPONames) {
            Add-GPOLink -GPOName $name -OUPath $OUPath -Order $order -Enforced $Enforced -Disabled $Disabled
            $order++
        }
        
        return $true
        
    } catch {
        Write-OutputColor "  ✗ Failed to replace links: $($_.Exception.Message)" "Red"
        return $false
    }
}

# Function to process CSV file
function Process-GPOLinkCSV {
    param([string]$CSVPath)
    
    try {
        $configs = Import-Csv -Path $CSVPath -ErrorAction Stop
        
        if ($configs.Count -eq 0) {
            Write-OutputColor "CSV file is empty or invalid" "Red"
            return
        }
        
        Write-OutputColor "Found $($configs.Count) GPO link configurations" "Green"
        
        $total = $configs.Count
        $current = 0
        
        foreach ($config in $configs) {
            $current++
            Write-ProgressBar -Current $current -Total $total -Activity "Processing CSV entries"
            
            $gpoName = $config.GPOName
            $ouPath = $config.TargetOU
            $order = if ($config.LinkOrder) { [int]$config.LinkOrder } else { $LinkOrder }
            $enforced = if ($config.Enforced -eq 'TRUE') { $true } else { $Enforced }
            $disabled = if ($config.Disabled -eq 'TRUE') { $true } else { $Disabled }
            $operation = if ($config.Operation) { $config.Operation } else { "Link" }
            
            if ([string]::IsNullOrEmpty($gpoName) -or [string]::IsNullOrEmpty($ouPath)) {
                Write-OutputColor "Skipping incomplete entry (missing GPO name or OU)" "Yellow"
                $script:SkippedCount++
                continue
            }
            
            Write-OutputColor "`nProcessing: $gpoName -> $(Get-OUSimpleName -OUPath $ouPath)" "White"
            
            switch ($operation.ToUpper()) {
                "LINK" {
                    Add-GPOLink -GPOName $gpoName -OUPath $ouPath -Order $order -Enforced $enforced -Disabled $disabled
                }
                "REMOVE" {
                    Remove-GPOLinkFromOU -GPOName $gpoName -OUPath $ouPath
                }
                "UPDATE" {
                    Add-GPOLink -GPOName $gpoName -OUPath $ouPath -Order $order -Enforced $enforced -Disabled $disabled -UpdateExistingLink
                }
                default {
                    Write-OutputColor "Unknown operation: $operation" "Yellow"
                    $script:SkippedCount++
                }
            }
        }
        
        Write-ProgressBar -Current $total -Total $total -Activity "Complete"
        
    } catch {
        Write-OutputColor "Failed to process CSV file: $($_.Exception.Message)" "Red"
        throw
    }
}

# Main execution
try {
    Write-OutputColor "===== GPO Link Management Script Started =====" "Cyan"
    
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
    
    # Import Group Policy module
    Import-Module GroupPolicy -Force
    
    # Verify domain connectivity
    try {
        $domain = Get-ADDomain -ErrorAction Stop
        Write-OutputColor "Connected to domain: $($domain.DNSRoot)" "Green"
    } catch {
        Write-OutputColor "Cannot connect to Active Directory domain." "Red"
        Write-OutputColor "Error: $_" "Red"
        exit 1
    }
    
    # Prepare for operations
    $gpoList = @()
    $ouList = @()
    
    # Build GPO list
    if ($PSCmdlet.ParameterSetName -eq 'Single' -or $PSCmdlet.ParameterSetName -eq 'Remove') {
        $gpoList = @($GPOName)
    } elseif ($PSCmdlet.ParameterSetName -eq 'Multiple') {
        $gpoList = $GPOList
    } elseif ($PSCmdlet.ParameterSetName -eq 'BulkMultiple') {
        $gpoList = $GPOList
        $ouList = $TargetOUList
    }
    
    # Build OU list
    if ($PSCmdlet.ParameterSetName -eq 'Single' -or $PSCmdlet.ParameterSetName -eq 'Remove' -or $PSCmdlet.ParameterSetName -eq 'Multiple') {
        $ouList = @($TargetOU)
    }
    
    # Display operation summary
    Write-OutputColor "`nOperation Summary:" "Cyan"
    Write-OutputColor ("=" * 60) "Gray"
    
    if ($RemoveLink) {
        Write-OutputColor "Operation: Remove GPO Links" "Yellow"
    } else {
        Write-OutputColor "Operation: Create/Update GPO Links" "Yellow"
    }
    
    Write-OutputColor "GPO(s): $($gpoList -join ', ')" "White"
    Write-OutputColor "OU(s): $($ouList -join ', ')" "White"
    Write-OutputColor "Link Order: $LinkOrder" "White"
    Write-OutputColor "Enforced: $(if ($Enforced) { 'Yes' } else { 'No' })" "White"
    Write-OutputColor "Disabled: $(if ($Disabled) { 'Yes' } else { 'No' })" "White"
    Write-OutputColor "Replace Existing: $(if ($ReplaceLinks) { 'Yes' } else { 'No' })" "White"
    Write-OutputColor ("=" * 60) "Gray"
    
    # Show current links for each OU
    if (-not $RemoveLink -and -not $WhatIf -and ($Force -or $PSCmdlet.ParameterSetName -eq 'CSVImport')) {
        foreach ($ou in $ouList) {
            if (Test-OUPath -OUPath $ou) {
                $currentLinks = Get-GPOLinksForOU -OUPath $ou
                Show-CurrentLinks -OUPath $ou -Links $currentLinks
            }
        }
    }
    
    # Confirm before proceeding
    if (-not $Force -and -not $WhatIf -and $PSCmdlet.ParameterSetName -ne 'CSVImport') {
        Write-OutputColor "`nDo you want to proceed with these changes?" "Yellow"
        $confirm = Read-Host "Enter Y to continue, N to cancel"
        if ($confirm -ne 'Y' -and $confirm -ne 'y') {
            Write-OutputColor "Operation cancelled by user" "Yellow"
            exit 0
        }
    }
    
    # Perform operations
    Write-OutputColor "`nStarting GPO link operations..." "Yellow"
    Write-OutputColor "" "White"
    
    if ($PSCmdlet.ParameterSetName -eq 'CSVImport') {
        # Process CSV file
        Process-GPOLinkCSV -CSVPath $CSVFile
    } else {
        # Process standard operations
        $totalOUs = $ouList.Count
        $currentOU = 0
        
        foreach ($ou in $ouList) {
            $currentOU++
            Write-ProgressBar -Current $currentOU -Total $totalOUs -Activity "Processing OUs"
            
            # Validate OU
            if (-not (Test-OUPath -OUPath $ou)) {
                Write-OutputColor "OU not found: $ou" "Red"
                $script:FailedCount += $gpoList.Count
                continue
            }
            
            $ouName = Get-OUSimpleName -OUPath $ou
            Write-OutputColor "`nProcessing OU: $ouName" "Cyan"
            
            if ($RemoveLink) {
                # Remove links
                foreach ($gpo in $gpoList) {
                    if ($WhatIf) {
                        Write-OutputColor "[WHATIF] Would remove link: $gpo from $ouName" "Gray"
                        $script:RemovedCount++
                        continue
                    }
                    
                    Remove-GPOLinkFromOU -GPOName $gpo -OUPath $ou
                }
            } elseif ($ReplaceLinks) {
                # Replace all links on OU
                if ($WhatIf) {
                    Write-OutputColor "[WHATIF] Would replace all links on $ouName with: $($gpoList -join ', ')" "Gray"
                    continue
                }
                
                Replace-GPOLinks -GPONames $gpoList -OUPath $ou -Enforced $Enforced -Disabled $Disabled
            } else {
                # Add links
                $order = $LinkOrder
                foreach ($gpo in $gpoList) {
                    if ($WhatIf) {
                        Write-OutputColor "[WHATIF] Would link: $gpo to $ouName (Order: $order)" "Gray"
                        $order++
                        continue
                    }
                    
                    Add-GPOLink -GPOName $gpo -OUPath $ou -Order $order -Enforced $Enforced -Disabled $Disabled -UpdateExistingLink:$UpdateExisting
                    $order++
                }
            }
        }
    }
    
    # Display summary
    $totalDuration = (Get-Date) - $script:StartTime
    
    Write-OutputColor "`n===== GPO Link Operation Summary =====" "Cyan"
    Write-OutputColor "Total operations: $($script:LinkedCount + $script:UpdatedCount + $script:RemovedCount + $script:FailedCount + $script:SkippedCount)" "White"
    Write-OutputColor "Links created: $script:LinkedCount" "Green"
    Write-OutputColor "Links updated: $script:UpdatedCount" "Yellow"
    Write-OutputColor "Links removed: $script:RemovedCount" "Red"
    Write-OutputColor "Failed: $script:FailedCount" $(if ($script:FailedCount -gt 0) { "Red" } else { "Green" })
    Write-OutputColor "Skipped: $script:SkippedCount" "Yellow"
    Write-OutputColor "Total duration: $($totalDuration.ToString('hh\:mm\:ss'))" "White"
    
    # Show operations log
    if ($script:Operations.Count -gt 0 -and -not $WhatIf) {
        Write-OutputColor "`nDetailed Operations:" "Cyan"
        foreach ($op in $script:Operations) {
            $color = if ($op.Operation -eq 'Remove') { "Red" } else { "Green" }
            $details = "  $($op.Operation): $($op.GPOName) -> $(Get-OUSimpleName -OUPath $op.OUPath)"
            if ($op.Order) { $details += " (Order: $($op.Order))" }
            if ($op.Enforced) { $details += " [Enforced]" }
            if ($op.Disabled) { $details += " [Disabled]" }
            Write-OutputColor $details $color
        }
    }
    
    # Write log file
    $logPath = Join-Path $env:TEMP "GPO_Link_Log_$(Get-Date -Format 'yyyyMMdd_HHmmss').txt"
    $logContent = @"
GPO Link Management Report
==========================
Date: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
Domain: $($domain.DNSRoot)
Operation: $(if ($RemoveLink) { 'Remove Links' } else { 'Create/Update Links' })
Total Operations: $($script:LinkedCount + $script:UpdatedCount + $script:RemovedCount + $script:FailedCount + $script:SkippedCount)
Links Created: $script:LinkedCount
Links Updated: $script:UpdatedCount
Links Removed: $script:RemovedCount
Failed: $script:FailedCount
Skipped: $script:SkippedCount
Duration: $($totalDuration.ToString('hh\:mm\:ss'))

Detailed Operations:
-------------------
"@

    foreach ($op in $script:Operations) {
        $logContent += "`n$($op.Operation): $($op.GPOName) -> $($op.OUPath)"
        if ($op.Order) { $logContent += " (Order: $($op.Order))" }
        if ($op.Enforced) { $logContent += " [Enforced]" }
        if ($op.Disabled) { $logContent += " [Disabled]" }
    }

    try {
        $logContent | Out-File -FilePath $logPath -Encoding UTF8 -Force
        Write-OutputColor "`nOperation log saved to: $logPath" "Cyan"
    } catch {
        Write-OutputColor "Failed to write log file: $($_.Exception.Message)" "Yellow"
    }
    
    # Exit with appropriate code
    if ($script:FailedCount -gt 0 -and ($script:LinkedCount -gt 0 -or $script:UpdatedCount -gt 0 -or $script:RemovedCount -gt 0)) {
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