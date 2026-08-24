<#
.SYNOPSIS
    BitLocker Status Report Script
.DESCRIPTION
    This script generates a comprehensive report of BitLocker encryption status for all fixed drives
    on the local system. It exports the results to a CSV file with detailed drive information.
.NOTES
    Author: Your Name
    Date: 2026-08-22
    Version: 1.0
    Requires: Administrative privileges
.LINK
    https://docs.microsoft.com/en-us/powershell/module/bitlocker/
#>

<#
Basic Usage:

    Run as Administrator (required for BitLocker operations)

    Double-click the script or run in PowerShell:

powershell

.\BitLocker_Status_Report.ps1

Advanced Usage:
powershell

# Custom output path
.\BitLocker_Status_Report.ps1 -OutputPath "C:\Reports\BitLocker.csv"

# Include recovery key information
.\BitLocker_Status_Report.ps1 -IncludeRecoveryKeys

# Both options together
.\BitLocker_Status_Report.ps1 -OutputPath "C:\Reports\BitLocker_Recovery.csv" -IncludeRecoveryKeys

Features:

    ✅ Scans all fixed drives

    ✅ Shows encryption status and percentage

    ✅ Displays protection status (On/Off)

    ✅ Exports to CSV with timestamp

    ✅ Option to include recovery key IDs

    ✅ Color-coded console output

    ✅ Error handling

    ✅ Administrative privileges check

Sample Output:

The CSV will contain:

    DriveLetter (e.g., C:, D:)

    ProtectionStatus (On/Off)

    EncryptionStatus (FullyEncrypted, PartiallyEncrypted, etc.)

    EncryptionPercentage

    SizeGB and FreeSpaceGB

    KeyProtectors (list of protectors)

    RecoveryKeyID (if enabled)

    LastCheck timestamp
#>

#Requires -RunAsAdministrator

[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [string]$OutputPath = "$env:USERPROFILE\Desktop\BitLocker_Report_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv",
    
    [Parameter(Mandatory = $false)]
    [switch]$IncludeRecoveryKeys
)

# Script configuration
$ErrorActionPreference = "Stop"
$ScriptName = "BitLocker_Status_Report"
$ScriptVersion = "1.0"
$ExecutionTime = Get-Date

Write-Host "=" * 60 -ForegroundColor Cyan
Write-Host "BitLocker Status Report v$ScriptVersion" -ForegroundColor Cyan
Write-Host "Generated on: $ExecutionTime" -ForegroundColor Cyan
Write-Host "=" * 60 -ForegroundColor Cyan
Write-Host ""

# Function to get BitLocker status for all drives
function Get-BitLockerStatusReport {
    [CmdletBinding()]
    param(
        [bool]$IncludeRecoveryKeys = $false
    )
    
    $DriveStatusReport = @()
    
    try {
        # Get all fixed drives
        $FixedDrives = Get-WmiObject -Class Win32_LogicalDisk | Where-Object { $_.DriveType -eq 3 }
        
        if (-not $FixedDrives) {
            Write-Warning "No fixed drives found on this system."
            return $DriveStatusReport
        }
        
        Write-Host "Scanning BitLocker status for $($FixedDrives.Count) fixed drives..." -ForegroundColor Yellow
        
        foreach ($Drive in $FixedDrives) {
            $DriveLetter = $Drive.DeviceID
            $Volume = Get-BitLockerVolume -MountPoint $DriveLetter -ErrorAction SilentlyContinue
            
            if (-not $Volume) {
                # Drive doesn't have BitLocker configuration
                $DriveStatusReport += [PSCustomObject]@{
                    DriveLetter     = $DriveLetter
                    VolumeType      = $Drive.VolumeName
                    FileSystem      = $Drive.FileSystem
                    SizeGB          = [math]::Round($Drive.Size / 1GB, 2)
                    FreeSpaceGB     = [math]::Round($Drive.FreeSpace / 1GB, 2)
                    ProtectionStatus = "Not Configured"
                    EncryptionStatus = "Not Encrypted"
                    EncryptionPercentage = "N/A"
                    KeyProtectors  = "N/A"
                    RecoveryKeyID  = "N/A"
                    LastCheck      = $ExecutionTime
                }
                continue
            }
            
            # Process BitLocker enabled drives
            $ProtectionStatus = $Volume.ProtectionStatus
            $EncryptionStatus = $Volume.VolumeStatus
            $EncryptionPercentage = $Volume.EncryptionPercentage
            
            # Get key protectors if available
            $KeyProtectors = ""
            $RecoveryKeyID = ""
            
            if ($IncludeRecoveryKeys) {
                try {
                    $Protectors = $Volume.KeyProtector
                    if ($Protectors) {
                        $RecoveryProtectors = $Protectors | Where-Object { $_.KeyProtectorType -eq 'RecoveryPassword' }
                        if ($RecoveryProtectors) {
                            $RecoveryKeyID = ($RecoveryProtectors | Select-Object -First 1).KeyProtectorId
                            $KeyProtectors = ($Protectors | ForEach-Object { $_.KeyProtectorType }) -join ", "
                        }
                    }
                }
                catch {
                    Write-Warning "Could not retrieve key protectors for drive $DriveLetter"
                    $KeyProtectors = "Error retrieving"
                }
            }
            
            # Create report object
            $DriveStatusReport += [PSCustomObject]@{
                DriveLetter     = $DriveLetter
                VolumeType      = $Drive.VolumeName
                FileSystem      = $Drive.FileSystem
                SizeGB          = [math]::Round($Drive.Size / 1GB, 2)
                FreeSpaceGB     = [math]::Round($Drive.FreeSpace / 1GB, 2)
                ProtectionStatus = if ($ProtectionStatus) { "On" } else { "Off" }
                EncryptionStatus = $EncryptionStatus.ToString().Replace('_', ' ')
                EncryptionPercentage = "$EncryptionPercentage%"
                KeyProtectors  = if ($KeyProtectors) { $KeyProtectors } else { "None" }
                RecoveryKeyID  = if ($RecoveryKeyID) { $RecoveryKeyID } else { "N/A" }
                LastCheck      = $ExecutionTime
            }
        }
    }
    catch {
        Write-Error "Error retrieving BitLocker status: $_"
        throw
    }
    
    return $DriveStatusReport
}

# Main execution
try {
    # Import BitLocker module if available
    if (-not (Get-Module -ListAvailable -Name BitLocker)) {
        Write-Warning "BitLocker module not found. Attempting to import..."
        try {
            Import-Module BitLocker -ErrorAction Stop
        }
        catch {
            Write-Error "Failed to import BitLocker module. Some features may not work."
        }
    }
    
    # Generate report
    Write-Host "Generating BitLocker status report..." -ForegroundColor Yellow
    $ReportData = Get-BitLockerStatusReport -IncludeRecoveryKeys $IncludeRecoveryKeys
    
    if ($ReportData.Count -eq 0) {
        Write-Warning "No data collected. Please ensure BitLocker module is available."
        exit 1
    }
    
    # Export to CSV
    try {
        $ReportData | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8
        Write-Host "✓ Report exported successfully to:" -ForegroundColor Green
        Write-Host "  $OutputPath" -ForegroundColor Cyan
    }
    catch {
        Write-Error "Failed to export CSV: $_"
        exit 1
    }
    
    # Display summary
    Write-Host ""
    Write-Host "=" * 60 -ForegroundColor Cyan
    Write-Host "SUMMARY REPORT" -ForegroundColor Cyan
    Write-Host "=" * 60 -ForegroundColor Cyan
    
    $TotalDrives = $ReportData.Count
    $EncryptedDrives = ($ReportData | Where-Object { $_.ProtectionStatus -eq "On" }).Count
    $UnencryptedDrives = ($ReportData | Where-Object { $_.ProtectionStatus -ne "On" }).Count
    
    Write-Host "Total Drives Examined: $TotalDrives" -ForegroundColor White
    Write-Host "Encrypted Drives: $EncryptedDrives" -ForegroundColor Green
    Write-Host "Unencrypted Drives: $UnencryptedDrives" -ForegroundColor Yellow
    
    Write-Host ""
    Write-Host "Drive Status Details:" -ForegroundColor Cyan
    $ReportData | Format-Table -Property DriveLetter, ProtectionStatus, EncryptionStatus, EncryptionPercentage, @{
        Name='Size(GB)'; Expression={$_.SizeGB}
    } -AutoSize
    
    Write-Host ""
    Write-Host "Report saved to: $OutputPath" -ForegroundColor Green
    Write-Host "Script execution completed successfully!" -ForegroundColor Green
    
}
catch {
    Write-Error "Script execution failed: $_"
    Write-Error "Error details: $($_.Exception.Message)"
    exit 1
}
finally {
    Write-Host ""
    Write-Host "Press any key to exit..." -ForegroundColor Gray
    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
}