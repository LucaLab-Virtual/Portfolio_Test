<#
.SYNOPSIS
    Creates new Group Policy Objects (GPOs) with configurable settings.
.DESCRIPTION
    This script creates new GPOs with optional settings including security filtering,
    WMI filtering, and initial configuration. Supports creating single or multiple GPOs
    from various input sources including CSV files.
.PARAMETER GPOName
    Name of the GPO to create (supports multiple names as array)
.PARAMETER GPOList
    Array of GPO names to create
.PARAMETER CSVFile
    Path to CSV file containing GPO configuration (columns: Name, Description, Filter, etc.)
.PARAMETER Description
    Description for the new GPO
.PARAMETER SecurityFilter
    Security group(s) to apply for GPO filtering
.PARAMETER WMIFilter
    WMI filter name or path to apply to the GPO
.PARAMETER Comment
    Additional comment for the GPO
.PARAMETER TargetOU
    OU path to link the GPO after creation
.PARAMETER LinkOrder
    Link order for the GPO (1 = highest priority)
.PARAMETER Enforced
    Enforces the GPO link
.PARAMETER Disabled
    Creates the GPO in disabled state
.PARAMETER Domain
    Target domain for GPO creation
.PARAMETER TemplateGPO
    Name of existing GPO to use as template
.PARAMETER BackupBeforeCreate
    Creates backup of existing GPO if it exists
.PARAMETER Force
    Overwrites existing GPO without confirmation
.PARAMETER WhatIf
    Shows what would happen without performing the creation
.EXAMPLE
    .\Create_GPO.ps1 -GPOName "New Security Policy" -Description "Company Security Settings"
    Creates a single GPO with description
.EXAMPLE
    .\Create_GPO.ps1 -GPOList "Policy1","Policy2","Policy3" -SecurityFilter "Domain Admins"
    Creates multiple GPOs with security filtering
.EXAMPLE
    .\Create_GPO.ps1 -CSVFile "C:\GPOS\gpo_list.csv" -Enforced
    Creates GPOs from CSV file with enforcement
.EXAMPLE
    .\Create_GPO.ps1 -GPOName "New Policy" -TemplateGPO "Existing Policy" -TargetOU "OU=Computers,DC=domain,DC=com"
    Creates a new GPO from template and links to OU
.NOTES
    Requires Group Policy Management Console (GPMC) module
    Must be run with administrative privileges
    Author: Portfolio Script
    Date: $(Get-Date -Format "yyyy-MM-dd")
#>

<#
Sample CSV File Template

Create a CSV file with the following format for bulk GPO creation:
csv

Name,Description,SecurityFilter,WMIFilter,Comment,TargetOU,TemplateGPO,Enforced,Disabled
"Workstation Security Policy","Security settings for workstations","Domain Computers","","Workstation policy","OU=Workstations,DC=domain,DC=com","","FALSE","FALSE"
"Server Security Policy","Security settings for servers","Domain Servers","","Server policy","OU=Servers,DC=domain,DC=com","Base Server Policy","TRUE","FALSE"
"AppLocker Policy","Application control policy","Domain Users","","AppLocker settings","OU=Computers,DC=domain,DC=com","AppLocker Template","FALSE","FALSE"
"User Desktop Policy","Desktop configuration","Domain Users","","Desktop settings","OU=Users,DC=domain,DC=com","","FALSE","TRUE"

Key Features:

    Multiple Creation Methods:

        Single GPO creation

        Bulk creation from array

        CSV file import for mass deployment

    Template Support:

        Create GPOs based on existing templates

        Copy settings from template GPOs

    Advanced Configuration:

        Security filtering (security groups)

        WMI filters

        Comments and descriptions

        Disabled state option

    OU Integration:

        Automatic linking to OUs

        Link order configuration

        Enforced links

    Safety Features:

        Checks for existing GPOs

        Backup option before overwriting

        Force mode for automation

        WhatIf preview mode

    Comprehensive Logging:

        Detailed creation logs

        Success/failure tracking

        GPO ID and details captured

Usage Examples:
powershell

# Create single GPO with description
.\Create_GPO.ps1 -GPOName "Workstation Security" -Description "Security settings for all workstations"

# Create multiple GPOs with security filtering
.\Create_GPO.ps1 -GPOList "Policy1","Policy2","Policy3" -SecurityFilter "Domain Admins","Domain Users"

# Create GPO from template and link to OU
.\Create_GPO.ps1 -GPOName "New Workstation Policy" -TemplateGPO "Base Workstation Policy" -TargetOU "OU=Workstations,DC=domain,DC=com"

# Create GPO with all options
.\Create_GPO.ps1 -GPOName "Advanced Policy" -Description "Advanced security settings" -SecurityFilter "IT Security Team" -WMIFilter "All Windows 10" -Comment "Created by script" -TargetOU "OU=Computers,DC=domain,DC=com" -Enforced -BackupBeforeCreate

# Bulk create from CSV
.\Create_GPO.ps1 -CSVFile "C:\GPOS\gpo_config.csv" -Force

# Preview what would be created
.\Create_GPO.ps1 -GPOName "Test Policy" -WhatIf

CSV Column Descriptions:
Column	Description	Required
Name	GPO display name	Yes
Description	GPO description	No
SecurityFilter	Semicolon-separated security groups	No
WMIFilter	WMI filter name	No
Comment	Additional comment	No
TargetOU	OU path for linking	No
TemplateGPO	Name of template GPO	No
Enforced	TRUE/FALSE for link enforcement	No
Disabled	TRUE/FALSE for disabled state	No
#>

[CmdletBinding(DefaultParameterSetName = 'Single')]
param(
    [Parameter(ParameterSetName = 'Single', Mandatory = $true, Position = 0)]
    [Parameter(ParameterSetName = 'Template', Mandatory = $true, Position = 0)]
    [string]$GPOName,
    
    [Parameter(ParameterSetName = 'Multiple')]
    [string[]]$GPOList,
    
    [Parameter(ParameterSetName = 'CSVImport')]
    [ValidateScript({Test-Path $_})]
    [string]$CSVFile,
    
    [Parameter()]
    [string]$Description,
    
    [Parameter()]
    [string[]]$SecurityFilter,
    
    [Parameter()]
    [string]$WMIFilter,
    
    [Parameter()]
    [string]$Comment,
    
    [Parameter()]
    [string]$TargetOU,
    
    [Parameter()]
    [int]$LinkOrder = 0,
    
    [Parameter()]
    [switch]$Enforced,
    
    [Parameter()]
    [switch]$Disabled,
    
    [Parameter()]
    [string]$Domain,
    
    [Parameter(ParameterSetName = 'Template')]
    [string]$TemplateGPO,
    
    [Parameter()]
    [switch]$BackupBeforeCreate,
    
    [Parameter()]
    [switch]$Force,
    
    [Parameter()]
    [switch]$WhatIf
)

# Set error handling
$ErrorActionPreference = "Stop"
$script:StartTime = Get-Date
$script:CreatedCount = 0
$script:FailedCount = 0
$script:SkippedCount = 0
$script:CreatedGPOs = @()
$script:FailedGPOs = @()

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

# Function to backup existing GPO
function Backup-ExistingGPO {
    param([string]$Name)
    
    try {
        $backupPath = Join-Path $env:TEMP "GPO_Backup_$(Get-Date -Format 'yyyyMMdd_HHmmss')"
        New-Item -Path $backupPath -ItemType Directory -Force | Out-Null
        
        Backup-GPO -Name $Name -Path $backupPath -ErrorAction Stop
        Write-OutputColor "  → Backup created: $backupPath" "Gray"
        return $backupPath
    } catch {
        Write-OutputColor "  → Failed to backup existing GPO: $($_.Exception.Message)" "Yellow"
        return $null
    }
}

# Function to create GPO from template
function New-GPOFromTemplate {
    param(
        [string]$Name,
        [string]$TemplateName,
        [string]$Description,
        [string[]]$SecurityFilter,
        [string]$WMIFilter,
        [string]$Comment,
        [bool]$Disabled
    )
    
    try {
        # Create new GPO
        $createParams = @{
            Name = $Name
            ErrorAction = "Stop"
        }
        
        if ($Description) {
            $createParams.Description = $Description
        }
        
        if ($Comment) {
            $createParams.Comment = $Comment
        }
        
        if ($Domain) {
            $createParams.Domain = $Domain
        }
        
        if ($Disabled) {
            $createParams.GPOStatus = "AllDisabled"
        }
        
        $newGPO = New-GPO @createParams
        
        # Copy settings from template
        if ($TemplateName) {
            Write-OutputColor "  → Copying settings from template: $TemplateName" "Cyan"
            Copy-GPO -SourceName $TemplateName -TargetName $Name -ErrorAction Stop
        }
        
        # Apply security filtering
        if ($SecurityFilter) {
            foreach ($filter in $SecurityFilter) {
                try {
                    Set-GPPermission -Name $Name -PermissionLevel GpoApply -TargetName $filter -TargetType Group -ErrorAction Stop
                    Write-OutputColor "  → Added security filter: $filter" "Cyan"
                } catch {
                    Write-OutputColor "  → Failed to add security filter '$filter': $($_.Exception.Message)" "Yellow"
                }
            }
        }
        
        # Apply WMI filter
        if ($WMIFilter) {
            try {
                Set-GPWMIFilter -Name $Name -WmiFilter $WMIFilter -ErrorAction Stop
                Write-OutputColor "  → Applied WMI filter: $WMIFilter" "Cyan"
            } catch {
                Write-OutputColor "  → Failed to apply WMI filter '$WMIFilter': $($_.Exception.Message)" "Yellow"
            }
        }
        
        return $newGPO
        
    } catch {
        Write-OutputColor "  → Failed to create GPO from template: $($_.Exception.Message)" "Red"
        throw
    }
}

# Function to create GPO
function New-GPOWithSettings {
    param(
        [string]$Name,
        [string]$Description,
        [string[]]$SecurityFilter,
        [string]$WMIFilter,
        [string]$Comment,
        [bool]$Disabled,
        [string]$TemplateGPO = $null
    )
    
    try {
        # Check if GPO exists
        $exists = Test-GPOExists -Name $Name
        
        if ($exists) {
            if (-not $Force) {
                $response = Read-Host "GPO '$Name' already exists. Overwrite? (Y/N/Skip)"
                if ($response -eq 'S' -or $response -eq 's') {
                    Write-OutputColor "Skipping GPO: $Name" "Yellow"
                    $script:SkippedCount++
                    return $null
                }
                if ($response -ne 'Y' -and $response -ne 'y') {
                    Write-OutputColor "Skipping GPO: $Name" "Yellow"
                    $script:SkippedCount++
                    return $null
                }
            }
            
            # Backup existing GPO if requested
            if ($BackupBeforeCreate) {
                Backup-ExistingGPO -Name $Name
            }
            
            if (-not $WhatIf) {
                # Remove existing GPO
                Remove-GPO -Name $Name -ErrorAction Stop
                Write-OutputColor "  → Removed existing GPO: $Name" "Yellow"
            }
        }
        
        if ($WhatIf) {
            Write-OutputColor "[WHATIF] Would create GPO: $Name" "Gray"
            $script:CreatedCount++
            return $null
        }
        
        # Create GPO
        $newGPO = if ($TemplateGPO) {
            New-GPOFromTemplate -Name $Name -TemplateName $TemplateGPO -Description $Description -SecurityFilter $SecurityFilter -WMIFilter $WMIFilter -Comment $Comment -Disabled $Disabled
        } else {
            $createParams = @{
                Name = $Name
                ErrorAction = "Stop"
            }
            
            if ($Description) {
                $createParams.Description = $Description
            }
            
            if ($Comment) {
                $createParams.Comment = $Comment
            }
            
            if ($Domain) {
                $createParams.Domain = $Domain
            }
            
            if ($Disabled) {
                $createParams.GPOStatus = "AllDisabled"
            }
            
            $newGPO = New-GPO @createParams
            
            # Apply security filtering
            if ($SecurityFilter) {
                foreach ($filter in $SecurityFilter) {
                    try {
                        Set-GPPermission -Name $Name -PermissionLevel GpoApply -TargetName $filter -TargetType Group -ErrorAction Stop
                        Write-OutputColor "  → Added security filter: $filter" "Cyan"
                    } catch {
                        Write-OutputColor "  → Failed to add security filter '$filter': $($_.Exception.Message)" "Yellow"
                    }
                }
            }
            
            # Apply WMI filter
            if ($WMIFilter) {
                try {
                    Set-GPWMIFilter -Name $Name -WmiFilter $WMIFilter -ErrorAction Stop
                    Write-OutputColor "  → Applied WMI filter: $WMIFilter" "Cyan"
                } catch {
                    Write-OutputColor "  → Failed to apply WMI filter '$WMIFilter': $($_.Exception.Message)" "Yellow"
                }
            }
            
            $newGPO
        }
        
        if ($newGPO) {
            $script:CreatedCount++
            $script:CreatedGPOs += $newGPO
            
            Write-OutputColor "  ✓ Successfully created GPO: $Name" "Green"
            Write-OutputColor "  → GPO ID: $($newGPO.Id)" "Gray"
            
            # Link to OU if specified
            if ($TargetOU) {
                Link-GPOToOU -GPOName $Name -TargetOU $TargetOU -Order $LinkOrder -Enforced:$Enforced
            }
            
            return $newGPO
        }
        
    } catch {
        $script:FailedCount++
        $script:FailedGPOs += @{
            Name = $Name
            Error = $_.Exception.Message
        }
        Write-OutputColor "  ✗ Failed to create GPO: $Name" "Red"
        Write-OutputColor "  → Error: $($_.Exception.Message)" "Red"
        return $null
    }
}

# Function to link GPO to OU
function Link-GPOToOU {
    param(
        [string]$GPOName,
        [string]$TargetOU,
        [int]$Order = 0,
        [bool]$Enforced = $false
    )
    
    try {
        # Verify OU exists
        if (-not (Test-OUPath -OUPath $TargetOU)) {
            Write-OutputColor "  → OU not found: $TargetOU" "Yellow"
            return $false
        }
        
        # Create GPO link
        $linkParams = @{
            Name = $GPOName
            Target = $TargetOU
            ErrorAction = "Stop"
        }
        
        if ($Order -gt 0) {
            $linkParams.Order = $Order
        }
        
        if ($Enforced) {
            $linkParams.Enforced = $true
        }
        
        New-GPLink @linkParams
        Write-OutputColor "  → Linked to OU: $TargetOU" "Green"
        
        if ($Order -gt 0) {
            Write-OutputColor "  → Link Order: $Order" "Gray"
        }
        if ($Enforced) {
            Write-OutputColor "  → Enforced: Yes" "Yellow"
        }
        
        return $true
        
    } catch {
        Write-OutputColor "  → Failed to link GPO to OU: $($_.Exception.Message)" "Yellow"
        return $false
    }
}

# Function to create GPOs from CSV
function New-GPOFromCSV {
    param([string]$CSVPath)
    
    try {
        $gpoConfigs = Import-Csv -Path $CSVPath -ErrorAction Stop
        
        if ($gpoConfigs.Count -eq 0) {
            Write-OutputColor "CSV file is empty or invalid" "Red"
            return
        }
        
        Write-OutputColor "Found $($gpoConfigs.Count) GPO configurations in CSV" "Green"
        
        foreach ($config in $gpoConfigs) {
            $name = $config.Name
            if ([string]::IsNullOrEmpty($name)) {
                Write-OutputColor "Skipping entry with missing Name field" "Yellow"
                continue
            }
            
            Write-OutputColor "`nProcessing: $name" "White"
            
            $description = if ($config.Description) { $config.Description } else { $Description }
            $securityFilter = if ($config.SecurityFilter) { $config.SecurityFilter -split ';' } else { $SecurityFilter }
            $wmiFilter = if ($config.WMIFilter) { $config.WMIFilter } else { $WMIFilter }
            $comment = if ($config.Comment) { $config.Comment } else { $Comment }
            $targetOU = if ($config.TargetOU) { $config.TargetOU } else { $TargetOU }
            $templateGPO = if ($config.TemplateGPO) { $config.TemplateGPO } else { $TemplateGPO }
            $enforced = if ($config.Enforced -eq 'TRUE') { $true } else { $Enforced }
            $disabled = if ($config.Disabled -eq 'TRUE') { $true } else { $Disabled }
            
            New-GPOWithSettings -Name $name -Description $description -SecurityFilter $securityFilter -WMIFilter $wmiFilter -Comment $comment -Disabled $disabled -TemplateGPO $templateGPO
        }
        
    } catch {
        Write-OutputColor "Failed to process CSV file: $($_.Exception.Message)" "Red"
        throw
    }
}

# Main execution
try {
    Write-OutputColor "===== GPO Creation Script Started =====" "Cyan"
    
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
    
    # Validate Template GPO if specified
    if ($TemplateGPO) {
        if (-not (Test-GPOExists -Name $TemplateGPO)) {
            Write-OutputColor "Template GPO not found: $TemplateGPO" "Red"
            exit 1
        }
        Write-OutputColor "Using template GPO: $TemplateGPO" "Cyan"
    }
    
    # Validate OU if specified
    if ($TargetOU -and (-not (Test-OUPath -OUPath $TargetOU))) {
        Write-OutputColor "Target OU not found: $TargetOU" "Red"
        Write-OutputColor "Please verify the OU path is correct." "Yellow"
        exit 1
    }
    
    # Start GPO creation
    Write-OutputColor "`nStarting GPO creation process..." "Yellow"
    Write-OutputColor "" "White"
    
    switch ($PSCmdlet.ParameterSetName) {
        'Single' {
            Write-OutputColor "Creating single GPO: $GPOName" "Cyan"
            New-GPOWithSettings -Name $GPOName -Description $Description -SecurityFilter $SecurityFilter -WMIFilter $WMIFilter -Comment $Comment -Disabled $Disabled -TemplateGPO $TemplateGPO
        }
        
        'Multiple' {
            Write-OutputColor "Creating $($GPOList.Count) GPOs" "Cyan"
            foreach ($name in $GPOList) {
                Write-OutputColor "`nProcessing: $name" "White"
                New-GPOWithSettings -Name $name -Description $Description -SecurityFilter $SecurityFilter -WMIFilter $WMIFilter -Comment $Comment -Disabled $Disabled -TemplateGPO $TemplateGPO
            }
        }
        
        'CSVImport' {
            Write-OutputColor "Creating GPOs from CSV file: $CSVFile" "Cyan"
            New-GPOFromCSV -CSVPath $CSVFile
        }
        
        'Template' {
            Write-OutputColor "Creating GPO from template: $GPOName (Template: $TemplateGPO)" "Cyan"
            New-GPOWithSettings -Name $GPOName -Description $Description -SecurityFilter $SecurityFilter -WMIFilter $WMIFilter -Comment $Comment -Disabled $Disabled -TemplateGPO $TemplateGPO
        }
    }
    
    # Display summary
    $totalDuration = (Get-Date) - $script:StartTime
    
    Write-OutputColor "`n===== GPO Creation Summary =====" "Cyan"
    Write-OutputColor "Total GPOs processed: $($script:CreatedCount + $script:FailedCount + $script:SkippedCount)" "White"
    Write-OutputColor "Successfully created: $script:CreatedCount" "Green"
    Write-OutputColor "Failed: $script:FailedCount" $(if ($script:FailedCount -gt 0) { "Red" } else { "Green" })
    Write-OutputColor "Skipped: $script:SkippedCount" "Yellow"
    Write-OutputColor "Total duration: $($totalDuration.ToString('hh\:mm\:ss'))" "White"
    
    if ($script:CreatedGPOs.Count -gt 0) {
        Write-OutputColor "`nCreated GPOs:" "Green"
        foreach ($gpo in $script:CreatedGPOs) {
            Write-OutputColor "  ✓ $($gpo.DisplayName) (ID: $($gpo.Id))" "Green"
        }
    }
    
    if ($script:FailedGPOs.Count -gt 0) {
        Write-OutputColor "`nFailed Items:" "Red"
        foreach ($item in $script:FailedGPOs) {
            Write-OutputColor "  ✗ $($item.Name) - $($item.Error)" "Red"
        }
    }
    
    # Write log file
    $logPath = Join-Path $env:TEMP "GPO_Creation_Log_$(Get-Date -Format 'yyyyMMdd_HHmmss').txt"
    $logContent = @"
GPO Creation Report
===================
Creation Date: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
Domain: $($domain.DNSRoot)
Parameter Set: $($PSCmdlet.ParameterSetName)
Total Processed: $($script:CreatedCount + $script:FailedCount + $script:SkippedCount)
Successful: $script:CreatedCount
Failed: $script:FailedCount
Skipped: $script:SkippedCount
Duration: $($totalDuration.ToString('hh\:mm\:ss'))

Created GPOs:
-------------
"@

    foreach ($gpo in $script:CreatedGPOs) {
        $logContent += "`n$($gpo.DisplayName) - ID: $($gpo.Id)"
    }

    if ($script:FailedGPOs.Count -gt 0) {
        $logContent += "`n`nFailed GPOs:"
        $logContent += "`n-------------"
        foreach ($item in $script:FailedGPOs) {
            $logContent += "`n$($item.Name) - $($item.Error)"
        }
    }

    try {
        $logContent | Out-File -FilePath $logPath -Encoding UTF8 -Force
        Write-OutputColor "`nCreation log saved to: $logPath" "Cyan"
    } catch {
        Write-OutputColor "Failed to write log file: $($_.Exception.Message)" "Yellow"
    }
    
    # Exit with appropriate code
    if ($script:FailedCount -gt 0 -and $script:CreatedCount -gt 0) {
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