<#
.SYNOPSIS
    Creates A records in Windows DNS Server with comprehensive validation and logging.

.DESCRIPTION
    This script creates one or more A records in a specified DNS zone on a Windows DNS server.
    It includes validation for IP addresses, duplicate record checking, and detailed logging.
    Supports both single record creation and bulk creation from CSV file.

.PARAMETER ZoneName
    The DNS zone where the A record will be created (e.g., "contoso.com")

.PARAMETER HostName
    The hostname for the A record (e.g., "webserver" or "www")

.PARAMETER IPAddress
    The IPv4 address for the A record (e.g., "192.168.1.100")

.PARAMETER TTL
    Time-To-Live for the DNS record in seconds. Default is 3600 (1 hour)

.PARAMETER DNSServer
    The DNS server to create the record on. Default is localhost

.PARAMETER CSVPath
    Path to CSV file for bulk creation. Format: HostName,IPAddress,TTL (TTL optional)

.PARAMETER LogPath
    Path to log file. Default: .\Create_A_Record_Log.txt

.PARAMETER WhatIf
    Shows what would happen without actually creating the record

.EXAMPLE
    .\Create_A_Record.ps1 -ZoneName "contoso.com" -HostName "webserver" -IPAddress "192.168.1.100"

.EXAMPLE
    .\Create_A_Record.ps1 -ZoneName "contoso.com" -HostName "mail" -IPAddress "10.0.0.50" -TTL 7200

.EXAMPLE
    .\Create_A_Record.ps1 -CSVPath "C:\Records\NewRecords.csv" -DNSServer "DC01"

.EXAMPLE
    .\Create_A_Record.ps1 -ZoneName "contoso.com" -HostName "test" -IPAddress "192.168.1.200" -WhatIf

.NOTES
    Author: Portfolio Script
    Version: 1.0
    Requires: Windows DNS Server role, PowerShell 5.1+
#>

<#
Usage Examples:
1. Single Record Creation
powershell

.\Create_A_Record.ps1 -ZoneName "contoso.com" -HostName "webserver" -IPAddress "192.168.1.100"

2. With Custom TTL
powershell

.\Create_A_Record.ps1 -ZoneName "contoso.com" -HostName "mail" -IPAddress "10.0.0.50" -TTL 7200

3. Bulk Creation from CSV
powershell

.\Create_A_Record.ps1 -ZoneName "contoso.com" -CSVPath "C:\Records\NewRecords.csv"

4. What-If Mode (Test Run)
powershell

.\Create_A_Record.ps1 -ZoneName "contoso.com" -HostName "test" -IPAddress "192.168.1.200" -WhatIf

5. Specify DNS Server
powershell

.\Create_A_Record.ps1 -ZoneName "contoso.com" -HostName "app" -IPAddress "192.168.1.150" -DNSServer "DC01"

CSV Template for Bulk Creation:

Create a CSV file with these columns:
csv

HostName,IPAddress,TTL
www,192.168.1.100,3600
mail,192.168.1.101,7200
ftp,192.168.1.102,1800
api,192.168.1.103,3600

Features:

    ✅ Full validation (IP format, hostname format)

    ✅ Duplicate record checking

    ✅ Detailed logging with timestamps

    ✅ Bulk creation from CSV

    ✅ WhatIf simulation mode

    ✅ Error handling and recovery

    ✅ Progress tracking

    ✅ Support for custom TTL

    ✅ Multiple DNS server support

    ✅ Interactive duplicate handling
#>

[CmdletBinding(DefaultParameterSetName = 'Single')]
param(
    [Parameter(Mandatory = $true, ParameterSetName = 'Single')]
    [Parameter(Mandatory = $true, ParameterSetName = 'SingleWithTTL')]
    [string]$ZoneName,

    [Parameter(Mandatory = $true, ParameterSetName = 'Single')]
    [Parameter(Mandatory = $true, ParameterSetName = 'SingleWithTTL')]
    [string]$HostName,

    [Parameter(Mandatory = $true, ParameterSetName = 'Single')]
    [Parameter(Mandatory = $true, ParameterSetName = 'SingleWithTTL')]
    [string]$IPAddress,

    [Parameter(Mandatory = $false, ParameterSetName = 'SingleWithTTL')]
    [int]$TTL = 3600,

    [Parameter(Mandatory = $false)]
    [string]$DNSServer = $env:COMPUTERNAME,

    [Parameter(Mandatory = $true, ParameterSetName = 'Bulk')]
    [string]$CSVPath,

    [Parameter(Mandatory = $false)]
    [string]$LogPath = ".\Create_A_Record_Log_$(Get-Date -Format 'yyyyMMdd_HHmmss').txt",

    [Parameter(Mandatory = $false)]
    [switch]$WhatIf
)

# Initialize logging
function Write-Log {
    param(
        [string]$Message,
        [string]$Level = "INFO"
    )
    $Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $LogEntry = "[$Timestamp] [$Level] $Message"
    Write-Host $LogEntry
    Add-Content -Path $LogPath -Value $LogEntry -Force
}

# Validate IP address format
function Test-ValidIP {
    param([string]$IP)
    $IPRegex = '^(25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\.(25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\.(25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\.(25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)$'
    return ($IP -match $IPRegex)
}

# Validate hostname format
function Test-ValidHostName {
    param([string]$HostName)
    $HostNameRegex = '^[a-zA-Z0-9]([a-zA-Z0-9-]*[a-zA-Z0-9])?$'
    return ($HostName -match $HostNameRegex)
}

# Check if record already exists
function Test-DNSRecordExists {
    param(
        [string]$Zone,
        [string]$Host,
        [string]$Server
    )
    try {
        $FQDN = if ($Host -eq '@') { $Zone } else { "$Host.$Zone" }
        $Result = Resolve-DnsName -Name $FQDN -Type A -Server $Server -ErrorAction SilentlyContinue
        return ($Result -ne $null)
    }
    catch {
        return $false
    }
}

# Create DNS A Record
function New-DNSARecord {
    param(
        [string]$Zone,
        [string]$Host,
        [string]$IP,
        [int]$TTLValue,
        [string]$Server,
        [bool]$Simulate
    )

    $FQDN = if ($Host -eq '@') { $Zone } else { "$Host.$Zone" }
    
    if ($Simulate) {
        Write-Log "WHATIF: Would create A record: $FQDN -> $IP (TTL: $TTLValue) on $Server" "INFO"
        return $true
    }

    try {
        # Using WMI to create DNS record
        $DnsServer = Get-WmiObject -Namespace root\MicrosoftDNS -Class MicrosoftDNS_Server -Filter "Name='$Server'" -ErrorAction Stop
        
        if (-not $DnsServer) {
            Write-Log "DNS server $Server not found" "ERROR"
            return $false
        }

        $ZoneObj = Get-WmiObject -Namespace root\MicrosoftDNS -Class MicrosoftDNS_Zone -Filter "Name='$Zone'" -ErrorAction Stop
        
        if (-not $ZoneObj) {
            Write-Log "DNS zone $Zone not found on server $Server" "ERROR"
            return $false
        }

        # Create the record using WMI
        $Record = Get-WmiObject -Namespace root\MicrosoftDNS -List -Class MicrosoftDNS_AType
        $Result = $Record.CreateInstanceFromPropertyData(
            $Server,
            $Zone,
            $Host,
            $IP,
            $TTLValue
        )
        
        if ($Result) {
            Write-Log "Successfully created A record: $FQDN -> $IP (TTL: $TTLValue)" "SUCCESS"
            return $true
        }
        else {
            Write-Log "Failed to create A record: $FQDN" "ERROR"
            return $false
        }
    }
    catch {
        Write-Log "Error creating record: $_" "ERROR"
        return $false
    }
}

# Main script execution
try {
    Write-Log "=== DNS A Record Creation Script Started ===" "INFO"
    Write-Log "Target DNS Server: $DNSServer" "INFO"
    Write-Log "WhatIf Mode: $WhatIf" "INFO"
    
    $RecordsCreated = 0
    $RecordsFailed = 0

    # Single record creation
    if ($PSCmdlet.ParameterSetName -eq 'Single' -or $PSCmdlet.ParameterSetName -eq 'SingleWithTTL') {
        Write-Log "Processing single record creation..." "INFO"
        
        # Validate parameters
        if (-not (Test-ValidIP -IP $IPAddress)) {
            Write-Log "Invalid IP address format: $IPAddress" "ERROR"
            throw "Invalid IP address format"
        }
        
        if (-not (Test-ValidHostName -HostName $HostName)) {
            Write-Log "Invalid hostname format: $HostName" "ERROR"
            throw "Invalid hostname format"
        }
        
        # Check for duplicate
        if (Test-DNSRecordExists -Zone $ZoneName -Host $HostName -Server $DNSServer) {
            Write-Log "WARNING: A record for $HostName.$ZoneName already exists" "WARNING"
            $Continue = Read-Host "Do you want to continue? (Y/N)"
            if ($Continue -ne 'Y' -and $Continue -ne 'y') {
                Write-Log "User canceled operation" "INFO"
                exit
            }
        }
        
        # Create the record
        if (New-DNSARecord -Zone $ZoneName -Host $HostName -IP $IPAddress -TTLValue $TTL -Server $DNSServer -Simulate $WhatIf) {
            $RecordsCreated++
        }
        else {
            $RecordsFailed++
        }
    }
    
    # Bulk creation from CSV
    if ($PSCmdlet.ParameterSetName -eq 'Bulk') {
        Write-Log "Processing bulk creation from CSV: $CSVPath" "INFO"
        
        if (-not (Test-Path $CSVPath)) {
            Write-Log "CSV file not found: $CSVPath" "ERROR"
            throw "CSV file not found"
        }
        
        try {
            $Records = Import-Csv -Path $CSVPath
            Write-Log "Found $($Records.Count) records to process" "INFO"
            
            foreach ($Record in $Records) {
                $HostName = $Record.HostName
                $IPAddress = $Record.IPAddress
                $TTLValue = if ($Record.TTL) { [int]$Record.TTL } else { 3600 }
                
                Write-Log "Processing: $HostName -> $IPAddress" "INFO"
                
                # Validate
                if (-not (Test-ValidIP -IP $IPAddress)) {
                    Write-Log "Invalid IP for $HostName: $IPAddress" "ERROR"
                    $RecordsFailed++
                    continue
                }
                
                if (-not (Test-ValidHostName -HostName $HostName)) {
                    Write-Log "Invalid hostname: $HostName" "ERROR"
                    $RecordsFailed++
                    continue
                }
                
                # Check for duplicate
                if (Test-DNSRecordExists -Zone $ZoneName -Host $HostName -Server $DNSServer) {
                    Write-Log "Record exists: $HostName.$ZoneName" "WARNING"
                    $RecordsFailed++
                    continue
                }
                
                # Create record
                if (New-DNSARecord -Zone $ZoneName -Host $HostName -IP $IPAddress -TTLValue $TTLValue -Server $DNSServer -Simulate $WhatIf) {
                    $RecordsCreated++
                }
                else {
                    $RecordsFailed++
                }
            }
        }
        catch {
            Write-Log "Error processing CSV: $_" "ERROR"
            throw
        }
    }
    
    # Summary
    Write-Log "=== Summary ===" "INFO"
    Write-Log "Records Created: $RecordsCreated" "INFO"
    Write-Log "Records Failed: $RecordsFailed" "INFO"
    Write-Log "Total Processed: $($RecordsCreated + $RecordsFailed)" "INFO"
    Write-Log "Log saved to: $LogPath" "INFO"
    Write-Log "=== Script Completed ===" "INFO"
}
catch {
    Write-Log "Fatal error: $_" "ERROR"
    Write-Log "Script terminated" "ERROR"
    exit 1
}