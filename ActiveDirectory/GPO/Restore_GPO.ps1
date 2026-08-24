<#
.SYNOPSIS
    Restores Group Policy Objects (GPOs) from backup files.
.DESCRIPTION
    This script restores one or multiple GPOs from backup folders created by the GPO backup process.
    Supports restoring individual GPOs, all GPOs from a backup, or restoring to new GPOs.
.PARAMETER BackupPath
    Path to the backup folder containing GPO backups
.PARAMETER RestoreAll
    Restores all GPOs found in the backup location
.PARAMETER GPOName
    Name of the GPO to restore (supports wildcards)
.PARAMETER GPOId
    ID of the GPO to restore (GUID format)
.PARAMETER RestoreToNew
    Creates a new GPO instead of overwriting existing ones
.PARAMETER NewName
    New name for the restored GPO (only when restoring a single GPO)
.PARAMETER TargetDomain
    Target domain for restoration (defaults to current domain)
.PARAMETER Force
    Overwrites existing GPOs without confirmation
.PARAMETER WhatIf
    Shows what would happen without performing the restore
.EXAMPLE
    .\Restore_GPO.ps1 -BackupPath "C:\GPO_Backups\GPO_Backup_2024-01-15_14-30-45"
    Restores all GPOs from the specified backup with interactive prompts
.EXAMPLE
    .\Restore_GPO.ps1 -BackupPath "C:\GPO_Backups\latest" -RestoreAll -Force
    Restores all GPOs without confirmation prompts
.EXAMPLE
    .\Restore_GPO.ps1 -BackupPath "C:\GPO_Backups\latest" -GPOName "Default Domain Policy" -RestoreToNew -NewName "Default Domain Policy_Restored"
    Restores a specific GPO as a new GPO with a different name
.NOTES
    Requires Group Policy Management Console (GPMC) module
    Must be run with administrative privileges
    Author: Portfolio Script
    Date: $(Get-Date -Format "yyyy-MM-dd")
#>

<#
    Multiple Restore Options:

        Restore all GPOs from backup

        Restore specific GPO by name (with wildcards)

        Restore specific GPO by ID

        Interactive mode for selection

    Safety Features:

        Checks for existing GPOs with warning/confirmation

        Option to create new GPOs instead of overwriting

        Force mode for automated restores

        WhatIf preview mode

    Comprehensive Logging:

        Detailed restore log file

        Color-coded console output

        Success/failure tracking

    Flexible Restore Options:

        Restore to new GPO with custom name

        Restore to original name (overwrite)

        Domain target specification

    User-Friendly Interface:

        Interactive mode with backup selection

        Clear progress indicators

        Detailed summary at completion

Usage Examples:
powershell

# Restore all GPOs from a backup (interactive)
.\Restore_GPO.ps1 -BackupPath "C:\GPO_Backups\GPO_Backup_2024-01-15_14-30-45"

# Force restore all GPOs without prompts
.\Restore_GPO.ps1 -BackupPath "C:\GPO_Backups\latest" -RestoreAll -Force

# Restore specific GPO as new
.\Restore_GPO.ps1 -BackupPath "C:\GPO_Backups\latest" -GPOName "Default Domain Policy" -RestoreToNew -NewName "Default Domain Policy_Restored"

# Restore GPO by ID
.\Restore_GPO.ps1 -BackupPath "C:\GPO_Backups\latest" -GPOId "12345678-1234-1234-1234-123456789012"

# Preview what would be restored
.\Restore_GPO.ps1 -BackupPath "C:\GPO_Backups\latest" -RestoreAll -WhatIf

Output Structure:

The script will:

    Validate the backup folder

    Display available GPO backups

    Let you choose which GPOs to restore

    Perform the restore with proper error handling

    Generate a detailed restore report

Safety Notes:

    The script will not overwrite existing GPOs unless explicitly told to do so

    Use -Force only when you're sure about overwriting

    The -RestoreToNew option creates a new GPO with a timestamp suffix

    Always test with -WhatIf first to verify what will be restored
#>

[CmdletBinding(DefaultParameterSetName = 'RestoreAll')]
param(
    [Parameter(Mandatory = $true, Position = 0)]
    [string]$BackupPath,
    
    [Parameter(ParameterSetName = 'RestoreAll')]
    [switch]$RestoreAll,
    
    [Parameter(ParameterSetName = 'SpecificGPO')]
    [string]$GPOName,
    
    [Parameter(ParameterSetName = 'SpecificGPO')]
    [guid]$GPOId,
    
    [Parameter()]
    [switch]$RestoreToNew,
    
    [Parameter(ParameterSetName = 'SpecificGPO')]
    [string]$NewName,
    
    [Parameter()]
    [string]$TargetDomain,
    
    [Parameter()]
    [switch]$Force,
    
    [Parameter()]
    [switch]$WhatIf
)

# Set error handling
$ErrorActionPreference = "Stop"
$script:StartTime = Get-Date
$script:RestoredCount = 0
$script:FailedCount = 0
$script:SkippedCount = 0
$script:FailedItems = @()

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

# Function to validate backup folder
function Test-BackupFolder {
    param([string]$Path)
    
    if (-not (Test-Path $Path)) {
        Write-OutputColor "Backup folder not found: $Path" "Red"
        return $false
    }
    
    # Check for GPO backup folders (they have GUID format)
    $backupItems = Get-ChildItem -Path $Path -Directory | Where-Object {
        $_.Name -match '^\{?[0-9a-fA-F]{8}-([0-9a-fA-F]{4}-){3}[0-9a-fA-F]{12}\}?$' -or
        (Test-Path (Join-Path $_.FullName "backup.xml"))
    }
    
    if ($backupItems.Count -eq 0) {
        Write-OutputColor "No valid GPO backup folders found in: $Path" "Yellow"
        return $false
    }
    
    return $true
}

# Function to get GPO backups
function Get-GPOBackups {
    param([string]$Path)
    
    $backups = @()
    $items = Get-ChildItem -Path $Path -Directory
    
    foreach ($item in $items) {
        # Check if it's a GPO backup folder
        if ($item.Name -match '^\{?[0-9a-fA-F]{8}-([0-9a-fA-F]{4}-){3}[0-9a-fA-F]{12}\}?$' -or
            (Test-Path (Join-Path $item.FullName "backup.xml"))) {
            
            try {
                # Try to load backup info
                $backupInfo = Get-GPOBackup -Path $item.FullName -ErrorAction SilentlyContinue
                if ($backupInfo) {
                    $backups += $backupInfo
                }
            } catch {
                Write-OutputColor "Skipping invalid backup folder: $($item.Name)" "Yellow"
            }
        }
    }
    
    return $backups
}

# Function to display backup details
function Show-BackupDetails {
    param($Backups)
    
    Write-OutputColor "`nFound GPO Backups:" "Cyan"
    Write-OutputColor ("=" * 80) "Gray"
    Write-OutputColor ("{0,-15} {1,-40} {2,-25}" -f "Type", "Name", "Backup Date") "Yellow"
    Write-OutputColor ("-" * 80) "Gray"
    
    foreach ($backup in $Backups) {
        $type = if ($RestoreToNew) { "NEW" } else { "RESTORE" }
        $name = if ($backup.DisplayName.Length -gt 38) { $backup.DisplayName.Substring(0, 38) + "..." } else { $backup.DisplayName }
        $date = if ($backup.CreationTime) { $backup.CreationTime.ToString('yyyy-MM-dd HH:mm') } else { "Unknown" }
        Write-OutputColor ("{0,-15} {1,-40} {2,-25}" -f $type, $name, $date) "White"
    }
    
    Write-OutputColor ("=" * 80) "Gray"
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

# Function to restore a single GPO
function Restore-GPOSingle {
    param(
        $BackupInfo,
        [bool]$CreateNew = $false,
        [string]$NewName = $null,
        [bool]$Force = $false
    )
    
    $backupPath = $BackupInfo.BackupPath
    $backupId = $BackupInfo.Id
    $originalName = $BackupInfo.DisplayName
    
    try {
        # Determine target name
        $targetName = if ($NewName) { $NewName } else { $originalName }
        
        # Check if GPO already exists
        $exists = Test-GPOExists -Name $targetName
        
        if ($exists) {
            if ($CreateNew) {
                Write-OutputColor "GPO '$targetName' already exists. Cannot create new GPO with same name." "Yellow"
                $script:SkippedCount++
                return
            }
            
            if (-not $Force) {
                $response = Read-Host "GPO '$targetName' already exists. Overwrite? (Y/N)"
                if ($response -ne 'Y' -and $response -ne 'y') {
                    Write-OutputColor "Skipping GPO: $originalName" "Yellow"
                    $script:SkippedCount++
                    return
                }
            }
            
            if ($WhatIf) {
                Write-OutputColor "[WHATIF] Would overwrite existing GPO: $targetName" "Gray"
                return
            }
            
            # Remove existing GPO
            try {
                Remove-GPO -Name $targetName -ErrorAction Stop
                Write-OutputColor "Removed existing GPO: $targetName" "Yellow"
            } catch {
                Write-OutputColor "Failed to remove existing GPO: $targetName" "Red"
                throw
            }
        }
        
        # Perform restore
        $restoreParams = @{
            BackupId = $backupId
            Path = (Split-Path $backupPath -Parent)
            TargetName = $targetName
            ErrorAction = "Stop"
        }
        
        if ($WhatIf) {
            Write-OutputColor "[WHATIF] Would restore GPO: $originalName as $targetName" "Gray"
            $script:RestoredCount++
            return
        }
        
        Write-OutputColor "Restoring GPO: $originalName" "White"
        if ($targetName -ne $originalName) {
            Write-OutputColor "  → New Name: $targetName" "Cyan"
        }
        
        $result = Restore-GPO @restoreParams
        
        if ($result) {
            $script:RestoredCount++
            Write-OutputColor "  ✓ Successfully restored: $targetName" "Green"
            Write-OutputColor "  → GPO ID: $($result.Id)" "Gray"
            return $result
        }
        
    } catch {
        $script:FailedCount++
        $script:FailedItems += @{
            Name = $originalName
            Error = $_.Exception.Message
        }
        Write-OutputColor "  ✗ Failed to restore: $originalName" "Red"
        Write-OutputColor "  → Error: $($_.Exception.Message)" "Red"
        return $null
    }
}

# Main execution
try {
    Write-OutputColor "===== GPO Restore Script Started =====" "Cyan"
    
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
    
    # Validate backup path
    $BackupPath = [System.IO.Path]::GetFullPath($BackupPath)
    if (-not (Test-BackupFolder -Path $BackupPath)) {
        exit 1
    }
    
    # Get available backups
    $backups = Get-GPOBackups -Path $BackupPath
    
    if ($backups.Count -eq 0) {
        Write-OutputColor "No valid GPO backups found in: $BackupPath" "Yellow"
        exit 0
    }
    
    Write-OutputColor "Found $($backups.Count) GPO backup(s)" "Green"
    
    # Filter backups based on parameters
    $selectedBackups = @()
    
    if ($RestoreAll) {
        $selectedBackups = $backups
        Write-OutputColor "Restoring all GPOs" "Cyan"
    } elseif ($GPOId) {
        $selectedBackups = $backups | Where-Object { $_.Id -eq $GPOId }
        if ($selectedBackups.Count -eq 0) {
            Write-OutputColor "No backup found with ID: $GPOId" "Red"
            exit 1
        }
        Write-OutputColor "Restoring GPO with ID: $GPOId" "Cyan"
    } elseif ($GPOName) {
        $selectedBackups = $backups | Where-Object { $_.DisplayName -like $GPOName }
        if ($selectedBackups.Count -eq 0) {
            Write-OutputColor "No backup found with name matching: $GPOName" "Red"
            exit 1
        }
        Write-OutputColor "Restoring GPOs matching: $GPOName ($($selectedBackups.Count) found)" "Cyan"
    } else {
        # Interactive mode - show available backups and let user choose
        Show-BackupDetails -Backups $backups
        
        Write-OutputColor "`nSelect restore option:" "Yellow"
        Write-OutputColor "  1) Restore all GPOs" "White"
        Write-OutputColor "  2) Restore specific GPO by number" "White"
        Write-OutputColor "  3) Cancel" "White"
        
        $choice = Read-Host "`nEnter your choice (1-3)"
        
        switch ($choice) {
            "1" { $selectedBackups = $backups }
            "2" {
                $number = Read-Host "Enter the backup number (1-$($backups.Count))"
                $index = [int]$number - 1
                if ($index -ge 0 -and $index -lt $backups.Count) {
                    $selectedBackups = @($backups[$index])
                } else {
                    Write-OutputColor "Invalid selection" "Red"
                    exit 1
                }
            }
            default {
                Write-OutputColor "Restore cancelled" "Yellow"
                exit 0
            }
        }
    }
    
    if ($selectedBackups.Count -eq 0) {
        Write-OutputColor "No backups selected for restore" "Yellow"
        exit 0
    }
    
    # Display restore summary
    Write-OutputColor "`nRestore Summary:" "Cyan"
    Write-OutputColor ("=" * 60) "Gray"
    Write-OutputColor "Backup Location: $BackupPath" "White"
    Write-OutputColor "Number of GPOs to restore: $($selectedBackups.Count)" "White"
    Write-OutputColor "Mode: $(if ($RestoreToNew) { 'Create New GPOs' } else { 'Restore (Overwrite)' })" "White"
    Write-OutputColor ("=" * 60) "Gray"
    
    # Confirm before proceeding
    if (-not $Force -and -not $WhatIf) {
        Write-OutputColor "`nDo you want to proceed with the restore?" "Yellow"
        $confirm = Read-Host "Enter Y to continue, N to cancel"
        if ($confirm -ne 'Y' -and $confirm -ne 'y') {
            Write-OutputColor "Restore cancelled by user" "Yellow"
            exit 0
        }
    }
    
    # Perform restore
    Write-OutputColor "`nStarting restore process..." "Yellow"
    Write-OutputColor "" "White"
    
    foreach ($backup in $selectedBackups) {
        $targetName = if ($RestoreToNew -and $NewName) {
            $NewName
        } elseif ($RestoreToNew) {
            "$($backup.DisplayName)_Restored_$(Get-Date -Format 'yyyyMMdd_HHmm')"
        } else {
            $null
        }
        
        Restore-GPOSingle -BackupInfo $backup -CreateNew $RestoreToNew -NewName $targetName -Force $Force
    }
    
    # Display summary
    $totalDuration = (Get-Date) - $script:StartTime
    
    Write-OutputColor "`n===== Restore Summary =====" "Cyan"
    Write-OutputColor "Total GPOs processed: $($selectedBackups.Count)" "White"
    Write-OutputColor "Successfully restored: $script:RestoredCount" "Green"
    Write-OutputColor "Failed: $script:FailedCount" $(if ($script:FailedCount -gt 0) { "Red" } else { "Green" })
    Write-OutputColor "Skipped: $script:SkippedCount" "Yellow"
    Write-OutputColor "Total duration: $($totalDuration.ToString('hh\:mm\:ss'))" "White"
    
    if ($script:FailedItems.Count -gt 0) {
        Write-OutputColor "`nFailed Items:" "Red"
        foreach ($item in $script:FailedItems) {
            Write-OutputColor "  $($item.Name) - $($item.Error)" "Red"
        }
    }
    
    # Write log file
    $logPath = Join-Path (Split-Path $BackupPath -Parent) "Restore_Log_$(Get-Date -Format 'yyyyMMdd_HHmmss').txt"
    $logContent = @"
GPO Restore Report
==================
Restore Date: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
Domain: $($domain.DNSRoot)
Backup Source: $BackupPath
Mode: $(if ($RestoreToNew) { 'Create New GPOs' } else { 'Restore (Overwrite)' })
Total Processed: $($selectedBackups.Count)
Successful: $script:RestoredCount
Failed: $script:FailedCount
Skipped: $script:SkippedCount
Duration: $($totalDuration.ToString('hh\:mm\:ss'))

Detailed Results:
----------------
"@
    
    foreach ($backup in $selectedBackups) {
        $status = "SUCCESS"
        $errorMsg = ""
        $failedItem = $script:FailedItems | Where-Object { $_.Name -eq $backup.DisplayName }
        if ($failedItem) {
            $status = "FAILED"
            $errorMsg = " - $($failedItem.Error)"
        }
        $logContent += "`n$($backup.DisplayName) ($($backup.Id)) - $status$errorMsg"
    }
    
    try {
        $logContent | Out-File -FilePath $logPath -Encoding UTF8 -Force
        Write-OutputColor "`nRestore log saved to: $logPath" "Cyan"
    } catch {
        Write-OutputColor "Failed to write log file: $($_.Exception.Message)" "Yellow"
    }
    
    # Exit with appropriate code
    if ($script:FailedCount -gt 0 -and $script:RestoredCount -gt 0) {
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