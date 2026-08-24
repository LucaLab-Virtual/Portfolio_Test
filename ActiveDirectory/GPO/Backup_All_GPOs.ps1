<#
.SYNOPSIS
    Backs up all Group Policy Objects (GPOs) in the domain to a specified location.
.DESCRIPTION
    This script creates a complete backup of all GPOs in the Active Directory domain.
    The backup includes all GPO settings, security filtering, and WMI filters.
    Each GPO is saved in a separate folder with a timestamp and GPO ID.
.PARAMETER BackupPath
    The path where GPO backups will be stored. Default is C:\GPO_Backups
.PARAMETER Description
    A description to include with the backup. Default is "Scheduled Backup"
.PARAMETER CreateTimestampFolder
    Creates a subfolder with current date/time stamp. Default is $true
.PARAMETER WhatIf
    Shows what would happen without actually performing the backup
.EXAMPLE
    .\Backup_All_GPOs.ps1
    Backs up all GPOs to C:\GPO_Backups\ with timestamp folder
.EXAMPLE
    .\Backup_All_GPOs.ps1 -BackupPath "D:\GPO_Archive" -Description "Monthly Full Backup"
    Backs up all GPOs to D:\GPO_Archive with custom description
.NOTES
    Requires Group Policy Management Console (GPMC) module
    Must be run with administrative privileges
    Author: Portfolio Script
    Date: $(Get-Date -Format "yyyy-MM-dd")
#>

<#
    Comprehensive Backup: Backs up all GPOs with their complete settings

    Timestamp Organization: Creates folders with timestamps for easy versioning

    Detailed Logging: Generates a backup report with success/failure details

    Error Handling: Robust error handling with informative messages

    Administrator Check: Verifies admin privileges before execution

    WhatIf Support: Preview mode to see what would be backed up

    Color Output: Visual feedback with colored console output

    Progress Tracking: Shows real-time progress during backup

    Domain Verification: Confirms domain connectivity before proceeding

Usage Examples:
powershell

# Basic backup (creates timestamp folder in C:\GPO_Backups)
.\Backup_All_GPOs.ps1

# Custom backup location without timestamp
.\Backup_All_GPOs.ps1 -BackupPath "D:\GPO_Archive" -CreateTimestampFolder:$false

# With custom description and preview mode
.\Backup_All_GPOs.ps1 -Description "Monthly Full Backup" -WhatIf

# Custom path with description
.\Backup_All_GPOs.ps1 -BackupPath "E:\Backups\GPO" -Description "Pre-Upgrade Backup"

Requirements:

    Windows Server or Windows 10/11 with RSAT

    Group Policy Management Console (GPMC) feature

    Active Directory module for PowerShell

    Administrative privileges

    PowerShell 5.1 or higher

Output Structure:
text

C:\GPO_Backups\
└── GPO_Backup_2024-01-15_14-30-45/
    ├── Backup_Report.txt
    └── {GPO-ID-1}/
    │   ├── backup.xml
    │   ├── gpo.inf
    │   └── ... (policy settings)
    ├── {GPO-ID-2}/
    │   └── ... (backup files)
    └── ...
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [string]$BackupPath = "C:\GPO_Backups",
    
    [Parameter(Mandatory = $false)]
    [string]$Description = "Scheduled Backup",
    
    [Parameter(Mandatory = $false)]
    [bool]$CreateTimestampFolder = $true,
    
    [Parameter(Mandatory = $false)]
    [switch]$WhatIf
)

# Set error handling
$ErrorActionPreference = "Stop"
$script:StartTime = Get-Date

# Function to write colored output
function Write-OutputColor {
    param(
        [string]$Message,
        [string]$Color = "White"
    )
    Write-Host "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] " -NoNewline -ForegroundColor Gray
    Write-Host $Message -ForegroundColor $Color
}

# Function to check if running as administrator
function Test-Administrator {
    $identity = [System.Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object System.Security.Principal.WindowsPrincipal($identity)
    return $principal.IsInRole([System.Security.Principal.WindowsBuiltInRole]::Administrator)
}

# Function to check GPMC module availability
function Test-GPModule {
    try {
        Import-Module GroupPolicy -ErrorAction SilentlyContinue
        return (Get-Module -Name GroupPolicy -ListAvailable) -ne $null
    } catch {
        return $false
    }
}

# Function to create backup directory
function New-BackupDirectory {
    param([string]$Path)
    
    if (-not (Test-Path $Path)) {
        try {
            New-Item -Path $Path -ItemType Directory -Force -ErrorAction Stop | Out-Null
            Write-OutputColor "Created backup directory: $Path" "Green"
        } catch {
            Write-OutputColor "Failed to create backup directory: $Path" "Red"
            Write-OutputColor "Error: $_" "Red"
            exit 1
        }
    }
    return $Path
}

# Function to get backup folder path
function Get-BackupFolderPath {
    param(
        [string]$BasePath,
        [string]$Description
    )
    
    $timestamp = Get-Date -Format "yyyy-MM-dd_HH-mm-ss"
    $folderName = "GPO_Backup_$timestamp"
    
    if ($Description -ne "Scheduled Backup") {
        $safeDescription = $Description -replace '[^a-zA-Z0-9]', '_'
        $folderName = "GPO_Backup_${timestamp}_$safeDescription"
    }
    
    return Join-Path $BasePath $folderName
}

# Main script execution
try {
    # Pre-flight checks
    Write-OutputColor "===== GPO Backup Script Started =====" "Cyan"
    
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
    
    # Prepare backup location
    $BackupPath = $BackupPath -replace '/', '\' -replace '\\$', ''
    $BackupPath = [System.IO.Path]::GetFullPath($BackupPath)
    
    if ($CreateTimestampFolder) {
        $BackupFolder = Get-BackupFolderPath -BasePath $BackupPath -Description $Description
    } else {
        $BackupFolder = $BackupPath
    }
    
    # Create backup directory
    if (-not $WhatIf) {
        New-BackupDirectory -Path $BackupFolder
    }
    
    # Get all GPOs in the domain
    Write-OutputColor "Retrieving all GPOs from the domain..." "Yellow"
    
    try {
        $allGPOs = Get-GPO -All -ErrorAction Stop
        $gpoCount = $allGPOs.Count
        
        if ($gpoCount -eq 0) {
            Write-OutputColor "No GPOs found in the domain." "Yellow"
            exit 0
        }
        
        Write-OutputColor "Found $gpoCount GPO(s) to backup." "Green"
        
        # Initialize counters
        $successCount = 0
        $failCount = 0
        $failedGPOs = @()
        
        # Create log file
        $logFile = Join-Path $BackupFolder "Backup_Report.txt"
        
        # Start backup process
        Write-OutputColor "Starting GPO backup process..." "Yellow"
        Write-OutputColor "Backup location: $BackupFolder" "Cyan"
        Write-OutputColor "Description: $Description" "Cyan"
        Write-OutputColor "" "White"
        
        # Process each GPO
        foreach ($gpo in $allGPOs) {
            try {
                $gpoId = $gpo.Id
                $gpoName = $gpo.DisplayName
                $gpoDomain = $gpo.DomainName
                
                Write-OutputColor "Backing up GPO: $gpoName" "White"
                
                if ($WhatIf) {
                    Write-OutputColor "[WHATIF] Would backup GPO: $gpoName (ID: $gpoId)" "Gray"
                    $successCount++
                    continue
                }
                
                # Perform backup
                $backupResult = Backup-GPO -Name $gpoName -Path $BackupFolder -Description "$Description - $gpoName" -ErrorAction Stop
                
                if ($backupResult) {
                    $successCount++
                    Write-OutputColor "  ✓ Successfully backed up: $gpoName" "Green"
                    Write-OutputColor "  → Backup ID: $($backupResult.BackupId)" "Gray"
                    Write-OutputColor "  → Location: $($backupResult.BackupPath)" "Gray"
                }
                
            } catch {
                $failCount++
                $failedGPOs += @{
                    Name = $gpoName
                    ID = $gpoId
                    Error = $_.Exception.Message
                }
                Write-OutputColor "  ✗ Failed to backup: $gpoName" "Red"
                Write-OutputColor "  → Error: $($_.Exception.Message)" "Red"
            }
            
            Write-OutputColor "" "White"
        }
        
        # Generate backup report
        $totalDuration = (Get-Date) - $script:StartTime
        
        Write-OutputColor "===== Backup Summary =====" "Cyan"
        Write-OutputColor "Total GPOs processed: $gpoCount" "White"
        Write-OutputColor "Successfully backed up: $successCount" "Green"
        Write-OutputColor "Failed backups: $failCount" $(if ($failCount -gt 0) { "Red" } else { "Green" })
        Write-OutputColor "Total duration: $($totalDuration.ToString('hh\:mm\:ss'))" "White"
        Write-OutputColor "Backup location: $BackupFolder" "Cyan"
        
        # Write detailed log
        if (-not $WhatIf) {
            $logContent = @"
GPO Backup Report
==================
Backup Date: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
Domain: $($domain.DNSRoot)
Description: $Description
Backup Location: $BackupFolder
Total GPOs: $gpoCount
Successful: $successCount
Failed: $failCount
Duration: $($totalDuration.ToString('hh\:mm\:ss'))

Detailed Results:
----------------
"@
            
            foreach ($gpo in $allGPOs) {
                $gpoName = $gpo.DisplayName
                $gpoId = $gpo.Id
                $status = if ($failedGPOs | Where-Object { $_.Name -eq $gpoName }) { "FAILED" } else { "SUCCESS" }
                $logContent += "`n$gpoName ($gpoId) - $status"
            }
            
            if ($failedGPOs.Count -gt 0) {
                $logContent += "`n`nFailed GPOs Details:"
                $logContent += "`n-------------------"
                foreach ($failed in $failedGPOs) {
                    $logContent += "`n$($failed.Name) - Error: $($failed.Error)"
                }
            }
            
            try {
                $logContent | Out-File -FilePath $logFile -Encoding UTF8 -Force
                Write-OutputColor "Report saved to: $logFile" "Cyan"
            } catch {
                Write-OutputColor "Failed to write log file: $($_.Exception.Message)" "Yellow"
            }
            
            # Display backup location in Explorer if successful
            if ($successCount -gt 0) {
                Write-OutputColor "`nBackup completed successfully!" "Green"
                Write-OutputColor "You can view the backups in: $BackupFolder" "Cyan"
                
                # Ask if user wants to open the backup folder
                $openFolder = Read-Host "`nDo you want to open the backup folder? (Y/N)"
                if ($openFolder -eq 'Y' -or $openFolder -eq 'y') {
                    Start-Process explorer.exe $BackupFolder
                }
            }
        }
        
        # Exit with appropriate code
        if ($failCount -gt 0 -and $successCount -gt 0) {
            exit 1  # Partial success
        } elseif ($failCount -gt 0) {
            exit 2  # Complete failure
        } else {
            exit 0  # Complete success
        }
        
    } catch {
        Write-OutputColor "Error during GPO retrieval: $_" "Red"
        exit 3
    }
    
} catch {
    Write-OutputColor "Unexpected error: $_" "Red"
    Write-OutputColor "Error details: $($_.Exception.Message)" "Red"
    exit 4
}