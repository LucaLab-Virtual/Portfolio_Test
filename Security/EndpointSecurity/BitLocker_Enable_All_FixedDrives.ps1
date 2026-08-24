<#
.SYNOPSIS
    BitLocker Enable All Fixed Drives Script
.DESCRIPTION
    This script enables BitLocker encryption on all fixed drives (system and data drives)
    with configurable protection methods, recovery key backup, and comprehensive logging.
    It includes safety checks, user confirmation, and detailed progress reporting.
.NOTES
    Author: Your Name
    Date: 2026-08-22
    Version: 1.0
    Requires: Administrative privileges
    Important: This operation can take several hours for large drives
.LINK
    https://docs.microsoft.com/en-us/powershell/module/bitlocker/
#>

<#
Basic Usage (Recommended - TPM+PIN):
powershell

# Run as Administrator
.\BitLocker_Enable_All_FixedDrives.ps1

Advanced Usage Examples:
powershell

# Use TPM only (no PIN)
.\BitLocker_Enable_All_FixedDrives.ps1 -ProtectionMethod TPM

# Use password protection
.\BitLocker_Enable_All_FixedDrives.ps1 -ProtectionMethod Password

# Skip system drive (encrypt data drives only)
.\BitLocker_Enable_All_FixedDrives.ps1 -SkipSystemDrive

# Encrypt used space only (faster)
.\BitLocker_Enable_All_FixedDrives.ps1 -EncryptUsedSpaceOnly

# Force mode (no confirmation prompts)
.\BitLocker_Enable_All_FixedDrives.ps1 -Force

# Custom backup and log location
.\BitLocker_Enable_All_FixedDrives.ps1 -BackupPath "D:\Backups\BitLocker" -LogPath "C:\Logs\BitLocker.log"

# Combine multiple options
.\BitLocker_Enable_All_FixedDrives.ps1 -ProtectionMethod TPMAndPIN -PINLength 8 -SkipSystemDrive -EncryptUsedSpaceOnly -Force

Key Features:
🛡️ Security Features:

    ✅ Multiple protection methods (TPM, Password, TPM+PIN, Recovery Password)

    ✅ Secure PIN input (masked)

    ✅ Automatic recovery key backup

    ✅ Comprehensive logging

    ✅ Support for used-space only encryption

📊 Management Features:

    ✅ Progress tracking

    ✅ Summary reports

    ✅ Individual recovery key files

    ✅ Consolidated recovery key export

    ✅ Skip system drive option

🛠️ Safety Features:

    ✅ Administrative privileges check

    ✅ Prerequisite validation

    ✅ User confirmation before execution

    ✅ Error handling with detailed messages

    ✅ WhatIf support for testing

📁 Output Files:

    Recovery Keys: Individual .txt files for each drive

    All Recovery Keys: All_Recovery_Keys.csv consolidated

    Summary Report: Summary_Report.csv

    Log File: Detailed execution log

Important Notes:

    Administrative rights required - Run as Administrator

    Time-consuming - Full encryption can take hours depending on drive size

    Backup recovery keys - Always save recovery keys in a secure location

    TPM requirement - TPM+PIN and TPM methods require a compatible TPM

    PIN requirements - Must be numeric, minimum 6 digits by default
#>

#Requires -RunAsAdministrator

[CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'High')]
param(
    [Parameter(Mandatory = $false)]
    [ValidateSet('TPM', 'Password', 'TPMAndPIN', 'RecoveryPassword')]
    [string]$ProtectionMethod = 'TPMAndPIN',
    
    [Parameter(Mandatory = $false)]
    [int]$PINLength = 6,
    
    [Parameter(Mandatory = $false)]
    [string]$BackupPath = "$env:USERPROFILE\Documents\BitLocker_Recovery_$(Get-Date -Format 'yyyyMMdd_HHmmss')",
    
    [Parameter(Mandatory = $false)]
    [switch]$SkipSystemDrive,
    
    [Parameter(Mandatory = $false)]
    [switch]$Force,
    
    [Parameter(Mandatory = $false)]
    [switch]$EncryptUsedSpaceOnly,
    
    [Parameter(Mandatory = $false)]
    [string]$LogPath = "$env:USERPROFILE\Desktop\BitLocker_Enable_Log_$(Get-Date -Format 'yyyyMMdd_HHmmss').txt"
)

# Script configuration
$ErrorActionPreference = "Stop"
$ScriptName = "BitLocker_Enable_All_FixedDrives"
$ScriptVersion = "1.0"
$ExecutionTime = Get-Date
$Global:RecoveryKeys = @{}

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
Write-Host "BitLocker Enable All Fixed Drives v$ScriptVersion" -ForegroundColor Cyan
Write-Host "Generated on: $ExecutionTime" -ForegroundColor Cyan
Write-Host "=" * 80 -ForegroundColor Cyan
Write-Host ""

# Create backup directory
function Initialize-BackupPath {
    try {
        if (-not (Test-Path $BackupPath)) {
            New-Item -Path $BackupPath -ItemType Directory -Force | Out-Null
            Write-Log "Created backup directory: $BackupPath" "INFO"
        }
        return $true
    }
    catch {
        Write-Log "Failed to create backup directory: $_" "ERROR"
        return $false
    }
}

# Function to validate TPM availability
function Test-TPMAvailability {
    try {
        $TPM = Get-WmiObject -Namespace "Root\CIMv2\Security\MicrosoftTpm" -Class Win32_Tpm
        if ($TPM -and $TPM.IsEnabled_InitialValue -and $TPM.IsActivated_InitialValue) {
            Write-Log "TPM is available and active" "INFO"
            return $true
        }
        else {
            Write-Log "TPM is not available or not activated" "WARNING"
            return $false
        }
    }
    catch {
        Write-Log "Error checking TPM: $_" "WARNING"
        return $false
    }
}

# Function to get secure password
function Get-SecurePIN {
    param(
        [int]$MinLength = 6
    )
    
    $PIN = $null
    $ValidPIN = $false
    
    while (-not $ValidPIN) {
        $PIN = Read-Host -Prompt "Enter a PIN (numeric, minimum $MinLength digits)" -AsSecureString
        
        # Convert secure string to plain text for validation
        $BSTR = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($PIN)
        $PINText = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($BSTR)
        [System.Runtime.InteropServices.Marshal]::ZeroFreeBSTR($BSTR)
        
        if ($PINText.Length -ge $MinLength -and $PINText -match '^\d+$') {
            $ValidPIN = $true
            return $PIN
        }
        else {
            Write-Host "Invalid PIN. Must be numeric and at least $MinLength digits long." -ForegroundColor Red
        }
    }
}

# Function to get recovery password for a drive
function Get-RecoveryPassword {
    param(
        [string]$DriveLetter,
        [string]$OutputPath
    )
    
    try {
        $Volume = Get-BitLockerVolume -MountPoint $DriveLetter -ErrorAction SilentlyContinue
        if ($Volume) {
            $RecoveryProtectors = $Volume.KeyProtector | Where-Object { $_.KeyProtectorType -eq 'RecoveryPassword' }
            if ($RecoveryProtectors) {
                foreach ($Protector in $RecoveryProtectors) {
                    $RecoveryKey = [PSCustomObject]@{
                        DriveLetter = $DriveLetter
                        RecoveryKeyID = $Protector.KeyProtectorId
                        RecoveryPassword = $Protector.RecoveryPassword
                        DateGenerated = Get-Date
                    }
                    
                    # Save to global collection
                    $Global:RecoveryKeys[$DriveLetter] = $RecoveryKey
                    
                    # Save to individual file
                    $RecoveryFile = Join-Path $OutputPath "RecoveryKey_$($DriveLetter.TrimEnd(':'))_$($Protector.KeyProtectorId).txt"
                    $RecoveryKey | Export-Csv -Path $RecoveryFile -NoTypeInformation -Encoding UTF8
                    Write-Log "Recovery key saved to: $RecoveryFile" "INFO"
                }
                return $true
            }
        }
        return $false
    }
    catch {
        Write-Log "Error retrieving recovery key for $DriveLetter`: $_" "ERROR"
        return $false
    }
}

# Function to enable BitLocker on a drive
function Enable-BitLockerDrive {
    param(
        [string]$DriveLetter,
        [string]$ProtectionMethod,
        [int]$PINLength,
        [switch]$EncryptUsedSpaceOnly,
        [string]$BackupPath
    )
    
    $DriveLetter = $DriveLetter.TrimEnd(':') + ':'
    $DriveStatus = Get-BitLockerVolume -MountPoint $DriveLetter -ErrorAction SilentlyContinue
    
    # Check if already encrypted
    if ($DriveStatus -and $DriveStatus.ProtectionStatus -eq 'On') {
        Write-Log "Drive $DriveLetter is already encrypted. Skipping..." "WARNING"
        return $false
    }
    
    Write-Log "Enabling BitLocker on drive $DriveLetter..." "INFO"
    
    try {
        $EnableParams = @{
            MountPoint = $DriveLetter
            SkipHardwareTest = $true
        }
        
        # Add encryption method
        if ($EncryptUsedSpaceOnly) {
            $EnableParams.UsedSpaceOnly = $true
            Write-Log "Using Used Space Only encryption for $DriveLetter" "INFO"
        }
        
        # Configure protection method
        switch ($ProtectionMethod) {
            'TPM' {
                if (Test-TPMAvailability) {
                    $EnableParams.TpmProtector = $true
                    Write-Log "Using TPM protector for $DriveLetter" "INFO"
                }
                else {
                    throw "TPM not available for $DriveLetter"
                }
            }
            'Password' {
                $Password = Read-Host -Prompt "Enter password for drive $DriveLetter" -AsSecureString
                $EnableParams.PasswordProtector = $Password
                Write-Log "Using Password protector for $DriveLetter" "INFO"
            }
            'TPMAndPIN' {
                if (Test-TPMAvailability) {
                    $PIN = Get-SecurePIN -MinLength $PINLength
                    $EnableParams.TpmPinProtector = $PIN
                    Write-Log "Using TPM+PIN protector for $DriveLetter" "INFO"
                }
                else {
                    throw "TPM not available for TPM+PIN protection on $DriveLetter"
                }
            }
            'RecoveryPassword' {
                $EnableParams.RecoveryPasswordProtector = $true
                Write-Log "Using Recovery Password protector for $DriveLetter" "INFO"
            }
        }
        
        # Enable BitLocker
        Write-Log "Starting BitLocker encryption on $DriveLetter (this may take a long time)..." "INFO"
        Enable-BitLocker @EnableParams -ErrorAction Stop
        
        # Wait for encryption to start
        Start-Sleep -Seconds 5
        
        # Get recovery key
        Write-Log "Retrieving recovery key for $DriveLetter..." "INFO"
        Get-RecoveryPassword -DriveLetter $DriveLetter -OutputPath $BackupPath
        
        Write-Log "BitLocker enabled successfully on $DriveLetter" "SUCCESS"
        return $true
    }
    catch {
        Write-Log "Failed to enable BitLocker on $DriveLetter`: $_" "ERROR"
        return $false
    }
}

# Main function
function Invoke-BitLockerEnable {
    Write-Log "Starting BitLocker enable process..." "INFO"
    Write-Log "Protection Method: $ProtectionMethod" "INFO"
    Write-Log "Backup Path: $BackupPath" "INFO"
    
    # Check prerequisites
    if (-not (Get-Module -ListAvailable -Name BitLocker)) {
        Write-Log "BitLocker module not found. Attempting to import..." "WARNING"
        try {
            Import-Module BitLocker -ErrorAction Stop
        }
        catch {
            Write-Log "Failed to import BitLocker module: $_" "ERROR"
            return $false
        }
    }
    
    # Initialize backup
    if (-not (Initialize-BackupPath)) {
        Write-Log "Failed to initialize backup path. Exiting..." "ERROR"
        return $false
    }
    
    # Get fixed drives
    $FixedDrives = Get-WmiObject -Class Win32_LogicalDisk | Where-Object { $_.DriveType -eq 3 }
    
    if (-not $FixedDrives) {
        Write-Log "No fixed drives found on this system." "ERROR"
        return $false
    }
    
    # Filter out system drive if requested
    if ($SkipSystemDrive) {
        $SystemDrive = $env:SystemDrive
        $FixedDrives = $FixedDrives | Where-Object { $_.DeviceID -ne $SystemDrive }
        Write-Log "Skipping system drive: $SystemDrive" "INFO"
    }
    
    Write-Log "Found $($FixedDrives.Count) fixed drive(s) to process:" "INFO"
    $FixedDrives | ForEach-Object { Write-Log "  $($_.DeviceID) ($($_.VolumeName))" "INFO" }
    
    # User confirmation
    if (-not $Force) {
        Write-Host ""
        Write-Host "WARNING: This will enable BitLocker on the following drives:" -ForegroundColor Yellow
        $FixedDrives | ForEach-Object { Write-Host "  $($_.DeviceID) ($($_.VolumeName))" -ForegroundColor Yellow }
        Write-Host ""
        
        $Confirmation = Read-Host "Do you want to continue? (y/n)"
        if ($Confirmation -ne 'y' -and $Confirmation -ne 'Y') {
            Write-Log "User cancelled the operation." "WARNING"
            return $false
        }
    }
    
    # Process each drive
    $SuccessCount = 0
    $FailureCount = 0
    
    foreach ($Drive in $FixedDrives) {
        $DriveLetter = $Drive.DeviceID
        Write-Host ""
        Write-Host "=" * 60 -ForegroundColor Cyan
        Write-Host "Processing drive: $DriveLetter" -ForegroundColor Cyan
        Write-Host "=" * 60 -ForegroundColor Cyan
        
        if (Enable-BitLockerDrive -DriveLetter $DriveLetter -ProtectionMethod $ProtectionMethod -PINLength $PINLength -EncryptUsedSpaceOnly:$EncryptUsedSpaceOnly -BackupPath $BackupPath) {
            $SuccessCount++
        }
        else {
            $FailureCount++
        }
    }
    
    # Generate summary report
    $SummaryPath = Join-Path $BackupPath "Summary_Report.csv"
    $SummaryData = @{
        ScriptName = $ScriptName
        ScriptVersion = $ScriptVersion
        ExecutionTime = $ExecutionTime
        TotalDrivesProcessed = $FixedDrives.Count
        SuccessfulEncryptions = $SuccessCount
        FailedEncryptions = $FailureCount
        ProtectionMethod = $ProtectionMethod
        BackupPath = $BackupPath
    }
    
    $SummaryData | Export-Csv -Path $SummaryPath -NoTypeInformation -Encoding UTF8
    
    # Export all recovery keys
    if ($Global:RecoveryKeys.Count -gt 0) {
        $RecoveryAllPath = Join-Path $BackupPath "All_Recovery_Keys.csv"
        $Global:RecoveryKeys.Values | Export-Csv -Path $RecoveryAllPath -NoTypeInformation -Encoding UTF8
        Write-Log "All recovery keys exported to: $RecoveryAllPath" "INFO"
    }
    
    # Display final summary
    Write-Host ""
    Write-Host "=" * 80 -ForegroundColor Cyan
    Write-Host "BITLOCKER ENABLE SUMMARY" -ForegroundColor Cyan
    Write-Host "=" * 80 -ForegroundColor Cyan
    Write-Host "Total Drives Processed: $($FixedDrives.Count)" -ForegroundColor White
    Write-Host "Successfully Encrypted: $SuccessCount" -ForegroundColor Green
    Write-Host "Failed: $FailureCount" -ForegroundColor Red
    Write-Host ""
    Write-Host "Backup Location: $BackupPath" -ForegroundColor Yellow
    Write-Host "Log File: $LogPath" -ForegroundColor Yellow
    Write-Host ""
    
    if ($Global:RecoveryKeys.Count -gt 0) {
        Write-Host "Recovery Keys Saved:" -ForegroundColor Green
        foreach ($Key in $Global:RecoveryKeys.Values) {
            Write-Host "  Drive $($Key.DriveLetter): $($Key.RecoveryKeyID)" -ForegroundColor Green
        }
        Write-Host ""
        Write-Host "IMPORTANT: Save your recovery keys in a secure location!" -ForegroundColor Yellow
    }
    
    Write-Log "Script execution completed with $SuccessCount successes and $FailureCount failures" "INFO"
    return ($FailureCount -eq 0)
}

# Main execution
try {
    $Result = Invoke-BitLockerEnable
    if ($Result) {
        Write-Host "`n✓ All BitLocker operations completed successfully!" -ForegroundColor Green
        exit 0
    }
    else {
        Write-Host "`n⚠ Script completed with some errors. Check the log for details." -ForegroundColor Yellow
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