<#
.SYNOPSIS
    Creates DHCP reservations with comprehensive validation, bulk import, and management capabilities.

.DESCRIPTION
    This script creates DHCP reservations with advanced features including:
    - Individual or bulk reservation creation
    - CSV import for mass reservations
    - MAC address validation (multiple formats)
    - IP address availability checking
    - Reservation conflict detection
    - Optional description and hostname support
    - Export of created reservations
    - Rollback capability for failed operations

.PARAMETER DHCPServer
    The name or IP address of the DHCP server. Defaults to localhost.

.PARAMETER ScopeId
    The Scope ID (network address) where the reservation will be created.

.PARAMETER IPAddress
    The IP address to reserve.

.PARAMETER ClientId
    The MAC address of the client (formats: AA-BB-CC-DD-EE-FF, AA:BB:CC:DD:EE:FF, or AABBCCDDEEFF).

.PARAMETER Name
    The friendly name for the reservation.

.PARAMETER Description
    Optional description for the reservation.

.PARAMETER HostName
    Optional hostname for the reservation.

.PARAMETER Type
    Reservation type: DHCP (default) or BOOTP.

.PARAMETER CSVFilePath
    Path to CSV file for bulk import (requires headers: ScopeId, IPAddress, ClientId, Name, [Description], [HostName], [Type]).

.PARAMETER ExportResults
    Path to export results summary (CSV format).

.PARAMETER SkipConflictCheck
    Skip checking for IP or MAC conflicts.

.PARAMETER AllowDuplicateIP
    Allow duplicate IP reservations (use with caution).

.PARAMETER ValidateMAC
    Validate MAC address format (default: true).

.PARAMETER WhatIf
    Simulate the operation without actually creating reservations.

.PARAMETER Confirm
    Prompt for confirmation before creating reservations.

.PARAMETER RollbackOnError
    Rollback all reservations if any creation fails.

.EXAMPLE
    .\Create_DHCP_Reservation.ps1 -ScopeId "192.168.1.0" -IPAddress "192.168.1.100" -ClientId "00-11-22-33-44-55" -Name "Printer-HP" -Description "HP LaserJet Office"

.EXAMPLE
    .\Create_DHCP_Reservation.ps1 -CSVFilePath "C:\Reservations\bulk_reservations.csv" -DHCPServer "DHCP01" -ExportResults "C:\Reservations\results.csv"

.EXAMPLE
    .\Create_DHCP_Reservation.ps1 -ScopeId "10.0.0.0" -IPAddress "10.0.0.50" -ClientId "AABBCCDDEEFF" -Name "Server-DC01" -HostName "dc01.domain.com" -Type DHCP -WhatIf

.NOTES
    Author: Portfolio Script
    Version: 1.0
    Requires: Windows Server with DHCP role installed, PowerShell 5.1+
#>

<#
Sample CSV File for Bulk Import

Create a CSV file with the following format:
csv

ScopeId,IPAddress,ClientId,Name,Description,HostName,Type
192.168.1.0,192.168.1.100,00-11-22-33-44-55,Printer-Office,HP LaserJet MFP,printer.office.local,DHCP
192.168.1.0,192.168.1.101,AA-BB-CC-DD-EE-FF,Scanner-Office,Epson Scanner,scanner.office.local,DHCP
10.0.0.0,10.0.0.50,11:22:33:44:55:66,Server-DC01,Primary Domain Controller,dc01.domain.com,DHCP
10.0.0.0,10.0.0.51,112233445566,Server-Web01,Web Server,iis01.domain.com,DHCP
192.168.2.0,192.168.2.200,FF-EE-DD-CC-BB-AA,VoIP-Phone01,Cisco Phone,phone01.voip.local,BOOTP

Key Features:

    Multiple Input Methods: Single reservation or bulk CSV import

    MAC Address Validation: Supports multiple MAC formats (AA-BB-CC-DD-EE-FF, AA:BB:CC:DD:EE:FF, AABBCCDDEEFF)

    Conflict Detection: Checks for IP and MAC conflicts

    Scope Validation: Verifies IP is within scope range

    Backup & Rollback: Creates backups and supports rollback on failure

    Progress Indicators: Visual progress for bulk operations

    Export Results: Saves detailed results to CSV

    WhatIf Support: Test before creating

    Error Handling: Comprehensive error handling with detailed messages

    Color-Coded Output: Easy-to-read console output

How to Use:
powershell

# Single reservation
.\Create_DHCP_Reservation.ps1 -ScopeId "192.168.1.0" -IPAddress "192.168.1.100" -ClientId "00-11-22-33-44-55" -Name "Printer" -Description "Office Printer"

# Single reservation with hostname
.\Create_DHCP_Reservation.ps1 -ScopeId "10.0.0.0" -IPAddress "10.0.0.50" -ClientId "AABBCCDDEEFF" -Name "Server-DC01" -HostName "dc01.domain.com"

# Bulk import from CSV
.\Create_DHCP_Reservation.ps1 -CSVFilePath "C:\Reservations\bulk_reservations.csv" -DHCPServer "DHCP01"

# With WhatIf to test
.\Create_DHCP_Reservation.ps1 -CSVFilePath "C:\Reservations\bulk_reservations.csv" -WhatIf

# With rollback on error
.\Create_DHCP_Reservation.ps1 -CSVFilePath "C:\Reservations\bulk_reservations.csv" -RollbackOnError -ExportResults "C:\Reservations\results.csv"

# Skip conflict checks and allow duplicates
.\Create_DHCP_Reservation.ps1 -ScopeId "192.168.1.0" -IPAddress "192.168.1.100" -ClientId "00-11-22-33-44-55" -Name "Printer" -SkipConflictCheck -AllowDuplicateIP
#>

[CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'Medium')]
param(
    [Parameter(Mandatory = $false)]
    [string]$DHCPServer = "localhost",
    
    [Parameter(Mandatory = $true, ParameterSetName = "Single")]
    [ValidatePattern("^(\d{1,3}\.){3}\d{1,3}$")]
    [string]$ScopeId,
    
    [Parameter(Mandatory = $true, ParameterSetName = "Single")]
    [ValidatePattern("^(\d{1,3}\.){3}\d{1,3}$")]
    [string]$IPAddress,
    
    [Parameter(Mandatory = $true, ParameterSetName = "Single")]
    [string]$ClientId,
    
    [Parameter(Mandatory = $true, ParameterSetName = "Single")]
    [string]$Name,
    
    [Parameter(Mandatory = $false, ParameterSetName = "Single")]
    [string]$Description,
    
    [Parameter(Mandatory = $false, ParameterSetName = "Single")]
    [string]$HostName,
    
    [Parameter(Mandatory = $false, ParameterSetName = "Single")]
    [ValidateSet("DHCP", "BOOTP")]
    [string]$Type = "DHCP",
    
    [Parameter(Mandatory = $true, ParameterSetName = "Bulk")]
    [ValidateScript({
        if (-not (Test-Path $_)) {
            throw "CSV file not found: $_"
        }
        return $true
    })]
    [string]$CSVFilePath,
    
    [Parameter(Mandatory = $false)]
    [string]$ExportResults,
    
    [Parameter(Mandatory = $false)]
    [switch]$SkipConflictCheck,
    
    [Parameter(Mandatory = $false)]
    [switch]$AllowDuplicateIP,
    
    [Parameter(Mandatory = $false)]
    [switch]$ValidateMAC = $true,
    
    [Parameter(Mandatory = $false)]
    [switch]$RollbackOnError,
    
    [Parameter(Mandatory = $false)]
    [switch]$WhatIf,
    
    [Parameter(Mandatory = $false)]
    [switch]$Confirm
)

# Global variables
$Script:ReservationsCreated = 0
$Script:ReservationsFailed = 0
$Script:ReservationsSkipped = 0
$Script:CreatedReservations = @()
$Script:FailedReservations = @()
$Script:ErrorMessages = @()
$Script:RollbackList = @()
$Script:StartTime = Get-Date

# Color definitions for console output
$Script:Colors = @{
    Success = "Green"
    Error = "Red"
    Warning = "Yellow"
    Info = "Cyan"
    Header = "Magenta"
    Detail = "Gray"
}

# Function to import DHCP module
function Import-DHCPModule {
    try {
        Import-Module DhcpServer -ErrorAction Stop -Force
        Write-Output "✓ DHCP Server module loaded successfully."
        return $true
    }
    catch {
        Write-Error "Failed to import DHCP Server module. Error: $($_.Exception.Message)"
        return $false
    }
}

# Function to normalize MAC address
function Normalize-MACAddress {
    param([string]$MAC)
    
    # Remove all separators and convert to uppercase
    $cleanMAC = $MAC -replace '[-:]', '' -replace '\.', '' -replace '\s', '' -replace '^0x', '' -replace '^0X', ''
    $cleanMAC = $cleanMAC.ToUpper()
    
    # Check length (should be 12 characters for a MAC address)
    if ($cleanMAC.Length -ne 12) {
        Write-Error "Invalid MAC address length: $MAC (expected 12 hex characters)"
        return $null
    }
    
    # Validate hex characters
    if ($cleanMAC -notmatch '^[0-9A-F]{12}$') {
        Write-Error "Invalid MAC address format: $MAC (contains non-hex characters)"
        return $null
    }
    
    # Format as XX-XX-XX-XX-XX-XX
    $formattedMAC = $cleanMAC -replace '(.{2})(.{2})(.{2})(.{2})(.{2})(.{2})', '$1-$2-$3-$4-$5-$6'
    return $formattedMAC
}

# Function to validate IP address in scope
function Test-IPInScope {
    param(
        [string]$IP,
        [string]$ScopeId,
        [string]$Server
    )
    
    try {
        $scope = Get-DhcpServerv4Scope -ComputerName $Server -ScopeId $ScopeId -ErrorAction Stop
        if (-not $scope) {
            return $false
        }
        
        $ipBytes = [System.Net.IPAddress]::Parse($IP).GetAddressBytes()
        $startBytes = $scope.StartRange.GetAddressBytes()
        $endBytes = $scope.EndRange.GetAddressBytes()
        
        # Convert to 32-bit integers for comparison
        [System.Array]::Reverse($ipBytes)
        [System.Array]::Reverse($startBytes)
        [System.Array]::Reverse($endBytes)
        $ipInt = [System.BitConverter]::ToUInt32($ipBytes, 0)
        $startInt = [System.BitConverter]::ToUInt32($startBytes, 0)
        $endInt = [System.BitConverter]::ToUInt32($endBytes, 0)
        
        return ($ipInt -ge $startInt -and $ipInt -le $endInt)
    }
    catch {
        return $false
    }
}

# Function to check if IP is already reserved
function Test-IPReserved {
    param(
        [string]$IP,
        [string]$ScopeId,
        [string]$Server
    )
    
    try {
        $reservations = Get-DhcpServerv4Reservation -ComputerName $Server -ScopeId $ScopeId -ErrorAction SilentlyContinue
        if ($reservations) {
            foreach ($res in $reservations) {
                if ($res.IPAddress.IPAddressToString -eq $IP) {
                    return $true
                }
            }
        }
        return $false
    }
    catch {
        return $false
    }
}

# Function to check if MAC is already reserved
function Test-MACReserved {
    param(
        [string]$MAC,
        [string]$ScopeId,
        [string]$Server
    )
    
    try {
        $reservations = Get-DhcpServerv4Reservation -ComputerName $Server -ScopeId $ScopeId -ErrorAction SilentlyContinue
        if ($reservations) {
            $normalizedMAC = Normalize-MACAddress -MAC $MAC
            if (-not $normalizedMAC) {
                return $false
            }
            foreach ($res in $reservations) {
                $resMAC = Normalize-MACAddress -MAC $res.ClientId
                if ($resMAC -eq $normalizedMAC) {
                    return $true
                }
            }
        }
        return $false
    }
    catch {
        return $false
    }
}

# Function to create a single reservation
function New-DHCPReservation {
    param(
        [string]$Server,
        [string]$ScopeId,
        [string]$IPAddress,
        [string]$ClientId,
        [string]$Name,
        [string]$Description,
        [string]$HostName,
        [string]$Type,
        [switch]$SkipConflictCheck,
        [switch]$AllowDuplicateIP,
        [switch]$WhatIf
    )
    
    $reservationResult = [PSCustomObject]@{
        ScopeId = $ScopeId
        IPAddress = $IPAddress
        ClientId = $ClientId
        Name = $Name
        Description = $Description
        HostName = $HostName
        Type = $Type
        Success = $false
        Message = ""
        Created = $null
    }
    
    try {
        # Validate MAC address
        if ($ValidateMAC) {
            $normalizedMAC = Normalize-MACAddress -MAC $ClientId
            if (-not $normalizedMAC) {
                $reservationResult.Message = "Invalid MAC address format"
                return $reservationResult
            }
            $reservationResult.ClientId = $normalizedMAC
        }
        
        # Validate scope exists
        $scopeExists = Get-DhcpServerv4Scope -ComputerName $Server -ScopeId $ScopeId -ErrorAction SilentlyContinue
        if (-not $scopeExists) {
            $reservationResult.Message = "Scope $ScopeId not found on server $Server"
            return $reservationResult
        }
        
        # Validate IP in scope range
        if (-not (Test-IPInScope -IP $IPAddress -ScopeId $ScopeId -Server $Server)) {
            $reservationResult.Message = "IP $IPAddress is not in the scope range"
            return $reservationResult
        }
        
        # Check for conflicts
        if (-not $SkipConflictCheck) {
            # Check IP conflict
            if (Test-IPReserved -IP $IPAddress -ScopeId $ScopeId -Server $Server) {
                if ($AllowDuplicateIP) {
                    Write-Warning "IP $IPAddress is already reserved, but allowing duplicate as specified"
                }
                else {
                    $reservationResult.Message = "IP $IPAddress is already reserved in this scope"
                    return $reservationResult
                }
            }
            
            # Check MAC conflict (skip if duplicate IP is allowed as it might be intentional)
            if (-not $AllowDuplicateIP) {
                if (Test-MACReserved -MAC $ClientId -ScopeId $ScopeId -Server $Server) {
                    $reservationResult.Message = "MAC $ClientId is already reserved in this scope"
                    return $reservationResult
                }
            }
        }
        
        # Create the reservation
        if ($WhatIf) {
            $reservationResult.Message = "[WHATIF] Would create reservation for $Name ($IPAddress) with MAC $ClientId"
            $reservationResult.Success = $true
            return $reservationResult
        }
        
        $createParams = @{
            ComputerName = $Server
            ScopeId = $ScopeId
            IPAddress = $IPAddress
            ClientId = $ClientId
            Name = $Name
            ErrorAction = "Stop"
        }
        
        if ($Description) {
            $createParams.Description = $Description
        }
        
        if ($HostName) {
            $createParams.HostName = $HostName
        }
        
        if ($Type -eq "BOOTP") {
            $createParams.Type = "Both"  # DHCP and BOOTP
        }
        
        $reservation = Add-DhcpServerv4Reservation @createParams
        $reservationResult.Success = $true
        $reservationResult.Message = "Reservation created successfully"
        $reservationResult.Created = $reservation
        $reservationResult.CreatedIP = $reservation.IPAddress.IPAddressToString
        $reservationResult.CreatedMAC = $reservation.ClientId
        
        Write-Host "  ✓ Created: $Name ($IPAddress)" -ForegroundColor Green
        
    }
    catch {
        $reservationResult.Message = "Error: $($_.Exception.Message)"
        if ($_.Exception.InnerException) {
            $reservationResult.Message += " - $($_.Exception.InnerException.Message)"
        }
    }
    
    return $reservationResult
}

# Function to read CSV file
function Read-CSVReservations {
    param([string]$FilePath)
    
    try {
        $csvData = Import-Csv -Path $FilePath -ErrorAction Stop
        
        # Validate required columns
        $requiredColumns = @("ScopeId", "IPAddress", "ClientId", "Name")
        $missingColumns = $requiredColumns | Where-Object { $_ -notin $csvData[0].PSObject.Properties.Name }
        
        if ($missingColumns) {
            Write-Error "CSV file missing required columns: $($missingColumns -join ', ')"
            return $null
        }
        
        return $csvData
    }
    catch {
        Write-Error "Failed to read CSV file: $($_.Exception.Message)"
        return $null
    }
}

# Function to create backup of current reservations
function Backup-Reservations {
    param(
        [string]$Server,
        [string]$ScopeId,
        [string]$BackupPath
    )
    
    try {
        $reservations = Get-DhcpServerv4Reservation -ComputerName $Server -ScopeId $ScopeId -ErrorAction SilentlyContinue
        if ($reservations) {
            $backup = $reservations | ForEach-Object {
                [PSCustomObject]@{
                    IPAddress = $_.IPAddress.IPAddressToString
                    ClientId = $_.ClientId
                    Name = $_.Name
                    Description = $_.Description
                    Type = $_.Type
                    HostName = $_.HostName
                }
            }
            $backup | Export-Csv -Path $BackupPath -NoTypeInformation -Encoding UTF8
            Write-Host "  ✓ Backup created at: $BackupPath" -ForegroundColor Gray
            return $true
        }
        else {
            Write-Host "  ℹ No existing reservations to backup" -ForegroundColor Yellow
            return $true
        }
    }
    catch {
        Write-Warning "Failed to create backup: $($_.Exception.Message)"
        return $false
    }
}

# Function to rollback reservations
function Rollback-Reservations {
    param(
        [string]$Server,
        [array]$Reservations
    )
    
    $rollbackCount = 0
    $rollbackFailed = 0
    
    Write-Host "`n🔄 Rolling back reservations..." -ForegroundColor Yellow
    
    foreach ($res in $Reservations) {
        try {
            if ($res.Success -and $res.Created) {
                Remove-DhcpServerv4Reservation -ComputerName $Server -ScopeId $res.ScopeId -IPAddress $res.IPAddress -Force -ErrorAction Stop
                Write-Host "  ✓ Rollback: $($res.Name) ($($res.IPAddress))" -ForegroundColor Green
                $rollbackCount++
            }
        }
        catch {
            Write-Host "  ✗ Failed to rollback: $($res.Name) ($($res.IPAddress)) - $($_.Exception.Message)" -ForegroundColor Red
            $rollbackFailed++
        }
    }
    
    Write-Host "Rollback completed: $rollbackCount succeeded, $rollbackFailed failed" -ForegroundColor Yellow
    return @{ Succeeded = $rollbackCount; Failed = $rollbackFailed }
}

# Function to export results
function Export-Results {
    param(
        [array]$Results,
        [string]$Path
    )
    
    try {
        $exportData = $Results | ForEach-Object {
            [PSCustomObject]@{
                ScopeId = $_.ScopeId
                IPAddress = $_.IPAddress
                ClientId = $_.ClientId
                Name = $_.Name
                Description = $_.Description
                HostName = $_.HostName
                Type = $_.Type
                Success = $_.Success
                Message = $_.Message
                Created = if ($_.Created) { $_.CreatedIP } else { "" }
            }
        }
        $exportData | Export-Csv -Path $Path -NoTypeInformation -Encoding UTF8
        Write-Host "✓ Results exported to: $Path" -ForegroundColor Green
        return $true
    }
    catch {
        Write-Error "Failed to export results: $($_.Exception.Message)"
        return $false
    }
}

# Function to display progress
function Show-Progress {
    param(
        [int]$Current,
        [int]$Total,
        [string]$Message = "Processing"
    )
    
    $percent = if ($Total -gt 0) { [math]::Round(($Current / $Total) * 100, 0) } else { 0 }
    Write-Progress -Activity "Creating DHCP Reservations" -Status "$Message" -PercentComplete $percent -CurrentOperation "Processing $Current of $Total"
}

# Function to validate scope exists
function Validate-Scope {
    param(
        [string]$ScopeId,
        [string]$Server
    )
    
    try {
        $scope = Get-DhcpServerv4Scope -ComputerName $Server -ScopeId $ScopeId -ErrorAction SilentlyContinue
        if (-not $scope) {
            Write-Error "Scope $ScopeId not found on server $Server"
            return $false
        }
        return $true
    }
    catch {
        Write-Error "Error validating scope: $($_.Exception.Message)"
        return $false
    }
}

# Main script execution
Write-Host ""
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "   DHCP RESERVATION CREATION TOOL" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

# Validate DHCP module
if (-not (Import-DHCPModule)) {
    Write-Error "DHCP module is required. Please install DHCP Server role."
    exit 1
}

# Check server connectivity
try {
    $testConnection = Test-Connection -ComputerName $DHCPServer -Count 1 -Quiet -ErrorAction SilentlyContinue
    if (-not $testConnection) {
        Write-Warning "Cannot reach DHCP server: $DHCPServer. Will attempt to continue..."
    }
}
catch {
    Write-Warning "Cannot ping DHCP server, but will try to connect..."
}

# Process single reservation or bulk
$reservationsToProcess = @()

if ($PSCmdlet.ParameterSetName -eq "Single") {
    # Single reservation
    Write-Host "📋 Creating single reservation..." -ForegroundColor Yellow
    
    # Validate scope
    if (-not (Validate-Scope -ScopeId $ScopeId -Server $DHCPServer)) {
        exit 1
    }
    
    $reservationsToProcess += [PSCustomObject]@{
        ScopeId = $ScopeId
        IPAddress = $IPAddress
        ClientId = $ClientId
        Name = $Name
        Description = $Description
        HostName = $HostName
        Type = $Type
    }
}
elseif ($PSCmdlet.ParameterSetName -eq "Bulk") {
    # Bulk import from CSV
    Write-Host "📋 Importing reservations from CSV: $CSVFilePath" -ForegroundColor Yellow
    
    $csvData = Read-CSVReservations -FilePath $CSVFilePath
    if (-not $csvData) {
        Write-Error "Failed to read CSV file. Exiting."
        exit 1
    }
    
    Write-Host "📊 Found $($csvData.Count) reservations to process" -ForegroundColor Yellow
    Write-Host ""
    
    # Group by ScopeId for validation
    $scopeGroups = $csvData | Group-Object -Property ScopeId
    $invalidScopes = @()
    
    foreach ($group in $scopeGroups) {
        if (-not (Validate-Scope -ScopeId $group.Name -Server $DHCPServer)) {
            $invalidScopes += $group.Name
        }
    }
    
    if ($invalidScopes) {
        Write-Error "Invalid scopes found: $($invalidScopes -join ', ')"
        if (-not $Confirm) {
            $response = Read-Host "Continue with valid scopes only? (Y/N)"
            if ($response -ne "Y") {
                Write-Host "Operation cancelled." -ForegroundColor Yellow
                exit 1
            }
        }
    }
    
    # Build reservation objects
    foreach ($row in $csvData) {
        $res = [PSCustomObject]@{
            ScopeId = $row.ScopeId
            IPAddress = $row.IPAddress
            ClientId = $row.ClientId
            Name = $row.Name
            Description = if ($row.Description) { $row.Description } else { "" }
            HostName = if ($row.HostName) { $row.HostName } else { "" }
            Type = if ($row.Type) { $row.Type } else { "DHCP" }
        }
        $reservationsToProcess += $res
    }
}

# Confirm operation
if ($Confirm -and -not $WhatIf) {
    Write-Host ""
    Write-Host "📋 Summary:" -ForegroundColor Cyan
    Write-Host "  Server: $DHCPServer"
    Write-Host "  Total Reservations: $($reservationsToProcess.Count)"
    Write-Host "  First reservation: $($reservationsToProcess[0].Name) ($($reservationsToProcess[0].IPAddress))"
    Write-Host ""
    
    $response = Read-Host "Proceed with creating these reservations? (Y/N)"
    if ($response -ne "Y") {
        Write-Host "Operation cancelled." -ForegroundColor Yellow
        exit 0
    }
}

# Create backup before processing (if not WhatIf)
if (-not $WhatIf) {
    Write-Host "`n💾 Creating backup of current reservations..." -ForegroundColor Yellow
    
    # Group by ScopeId for backup
    $scopeGroups = $reservationsToProcess | Group-Object -Property ScopeId
    $backupPath = Join-Path -Path $env:TEMP -ChildPath "DHCP_Reservations_Backup_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
    
    foreach ($group in $scopeGroups) {
        $backupFile = Join-Path -Path (Split-Path $backupPath -Parent) -ChildPath "DHCP_Backup_$($group.Name)_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
        Backup-Reservations -Server $DHCPServer -ScopeId $group.Name -BackupPath $backupFile
    }
}

# Process reservations
Write-Host "`n📝 Creating reservations..." -ForegroundColor Yellow
Write-Host ""

$totalCount = $reservationsToProcess.Count
$currentIndex = 0
$batchResults = @()

foreach ($res in $reservationsToProcess) {
    $currentIndex++
    Show-Progress -Current $currentIndex -Total $totalCount -Message "Creating $($res.Name)"
    
    # Process the reservation
    $result = New-DHCPReservation -Server $DHCPServer -ScopeId $res.ScopeId -IPAddress $res.IPAddress -ClientId $res.ClientId -Name $res.Name -Description $res.Description -HostName $res.HostName -Type $res.Type -SkipConflictCheck:$SkipConflictCheck -AllowDuplicateIP:$AllowDuplicateIP -WhatIf:$WhatIf
    
    $batchResults += $result
    
    if ($result.Success) {
        $Script:ReservationsCreated++
        $Script:CreatedReservations += $result
    }
    elseif ($result.Message -like "*already reserved*" -or $result.Message -like "*already exists*") {
        $Script:ReservationsSkipped++
        Write-Host "  ℹ Skipped: $($res.Name) ($($res.IPAddress)) - Already exists" -ForegroundColor Yellow
    }
    else {
        $Script:ReservationsFailed++
        $Script:FailedReservations += $result
        Write-Host "  ✗ Failed: $($res.Name) ($($res.IPAddress)) - $($result.Message)" -ForegroundColor Red
    }
    
    # Check if we should rollback on error
    if ($RollbackOnError -and -not $result.Success -and -not $WhatIf) {
        Write-Host "`n⚠️ Error detected, rolling back all reservations..." -ForegroundColor Red
        $rollbackResult = Rollback-Reservations -Server $DHCPServer -Reservations $Script:CreatedReservations
        Write-Host "Rollback completed: $($rollbackResult.Succeeded) succeeded, $($rollbackResult.Failed) failed" -ForegroundColor Yellow
        Write-Host "Operation halted due to error." -ForegroundColor Red
        exit 1
    }
}

# Clear progress
Write-Progress -Activity "Creating DHCP Reservations" -Completed

# Export results
if ($ExportResults) {
    Export-Results -Results $batchResults -Path $ExportResults
}

# Display summary
$elapsedTime = (Get-Date) - $Script:StartTime
Write-Host ""
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "   RESERVATION CREATION COMPLETED" -ForegroundColor Green
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "📊 Summary:" -ForegroundColor Cyan
Write-Host "  Total Requests: $totalCount" -ForegroundColor White
Write-Host "  ✅ Created: $Script:ReservationsCreated" -ForegroundColor Green
Write-Host "  ⚠️  Skipped: $Script:ReservationsSkipped" -ForegroundColor Yellow
Write-Host "  ❌ Failed: $Script:ReservationsFailed" -ForegroundColor Red
Write-Host "  ⏱️  Time: $($elapsedTime.ToString('hh\:mm\:ss'))" -ForegroundColor White
Write-Host ""

# Show detailed results
if ($Script:CreatedReservations.Count -gt 0) {
    Write-Host "Created Reservations:" -ForegroundColor Green
    foreach ($res in $Script:CreatedReservations) {
        Write-Host "  ✓ $($res.Name) - $($res.IPAddress) ($($res.ClientId))" -ForegroundColor Gray
    }
    Write-Host ""
}

if ($Script:FailedReservations.Count -gt 0) {
    Write-Host "Failed Reservations:" -ForegroundColor Red
    foreach ($res in $Script:FailedReservations) {
        Write-Host "  ✗ $($res.Name) - $($res.IPAddress): $($res.Message)" -ForegroundColor Gray
    }
    Write-Host ""
}

# Final status
if ($Script:ReservationsFailed -eq 0) {
    Write-Host "✅ All reservations processed successfully!" -ForegroundColor Green
    exit 0
}
else {
    Write-Host "⚠️  Some reservations failed. Review the errors above." -ForegroundColor Yellow
    if ($ExportResults) {
        Write-Host "📄 Check the results file for details: $ExportResults" -ForegroundColor Yellow
    }
    exit 1
}