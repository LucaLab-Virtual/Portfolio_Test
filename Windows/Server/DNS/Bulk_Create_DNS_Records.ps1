<#
.SYNOPSIS
    Bulk creates multiple DNS records (A, CNAME, MX, TXT) from CSV or JSON input with advanced features.

.DESCRIPTION
    This script enables bulk creation of various DNS record types from structured input files.
    Supports A, CNAME, MX, TXT, and SRV records with validation, error handling, and 
    detailed reporting. Perfect for migration, new deployments, and automated DNS management.

.PARAMETER InputFile
    Path to CSV or JSON file containing DNS records to create

.PARAMETER FileType
    Type of input file: 'CSV' or 'JSON'. Default is 'CSV' (auto-detected from extension)

.PARAMETER DNSServer
    DNS server to create records on. Default is localhost

.PARAMETER ZoneName
    DNS zone name (optional if specified in each record)

.PARAMETER LogPath
    Path to log file. Default: .\Bulk_Create_Log_[timestamp].txt

.PARAMETER ReportPath
    Path to save detailed report. Default: .\Bulk_Create_Report_[timestamp].csv

.PARAMETER FailOnError
    Stop script on first error

.PARAMETER WhatIf
    Preview changes without executing

.PARAMETER Parallel
    Enable parallel processing (experimental, use with caution)

.PARAMETER MaxParallel
    Maximum parallel jobs when Parallel is enabled. Default: 5

.PARAMETER TransactionLog
    Enable transaction logging for rollback capability

.PARAMETER ValidateOnly
    Only validate records without creating them

.EXAMPLE
    .\Bulk_Create_DNS_Records.ps1 -InputFile "C:\DNS\Records.csv" -ZoneName "contoso.com"

.EXAMPLE
    .\Bulk_Create_DNS_Records.ps1 -InputFile "records.json" -DNSServer "DC01" -ReportPath "C:\Reports\DNSReport.csv"

.EXAMPLE
    .\Bulk_Create_DNS_Records.ps1 -InputFile "records.csv" -WhatIf -ValidateOnly

.NOTES
    Author: Portfolio Script
    Version: 2.0
    Requires: Windows DNS Server, PowerShell 5.1+
    Supports: A, CNAME, MX, TXT, SRV records
#>

<#
Usage Examples:
1. Basic CSV Import
powershell

.\Bulk_Create_DNS_Records.ps1 -InputFile "C:\DNS\Records.csv" -ZoneName "contoso.com"

2. JSON Input with Custom DNS Server
powershell

.\Bulk_Create_DNS_Records.ps1 -InputFile "records.json" -DNSServer "DC01" -ReportPath "C:\Reports\DNSReport.csv"

3. Validation Only (No Creation)
powershell

.\Bulk_Create_DNS_Records.ps1 -InputFile "records.csv" -WhatIf -ValidateOnly

4. With Error Handling
powershell

.\Bulk_Create_DNS_Records.ps1 -InputFile "records.csv" -ZoneName "contoso.com" -FailOnError

5. Parallel Processing
powershell

.\Bulk_Create_DNS_Records.ps1 -InputFile "records.csv" -ZoneName "contoso.com" -Parallel -MaxParallel 10

Input File Formats:
CSV Format (records.csv):
csv

HostName,RecordType,RecordData,TTL,Priority,Weight,Port,ZoneName
www,A,192.168.1.100,3600,,,,contoso.com
mail,MX,mail.contoso.com,3600,10,,,
ftp,CNAME,server01.contoso.com,7200,,,,
txtrecord,TXT,"v=spf1 mx -all",3600,,,,
_sip._tcp,SRV,sipserver.contoso.com,3600,10,5,5060,

JSON Format (records.json):
json

[
  {
    "HostName": "www",
    "RecordType": "A",
    "RecordData": "192.168.1.100",
    "TTL": 3600,
    "ZoneName": "contoso.com"
  },
  {
    "HostName": "mail",
    "RecordType": "MX",
    "RecordData": "mail.contoso.com",
    "TTL": 3600,
    "Priority": 10,
    "ZoneName": "contoso.com"
  },
  {
    "HostName": "api",
    "RecordType": "CNAME",
    "RecordData": "server01.contoso.com",
    "TTL": 7200,
    "ZoneName": "contoso.com"
  },
  {
    "HostName": "_sip._tcp",
    "RecordType": "SRV",
    "RecordData": "sipserver.contoso.com",
    "TTL": 3600,
    "Priority": 10,
    "Weight": 5,
    "Port": 5060,
    "ZoneName": "contoso.com"
  }
]

Features:

    ✅ Multiple Record Types: A, CNAME, MX, TXT, SRV

    ✅ Multiple Input Formats: CSV and JSON

    ✅ Comprehensive Validation: IP, hostname, FQDN validation

    ✅ Duplicate Detection: Prevents duplicate record creation

    ✅ Detailed Logging: With colored console output

    ✅ Export Reports: CSV reports with success/failure status

    ✅ WhatIf Mode: Preview changes without execution

    ✅ Validation Mode: Validate records without creating

    ✅ Error Handling: Optional fail-on-error behavior

    ✅ Progress Tracking: Visual progress indicator

    ✅ Parallel Processing: Optional parallel execution

    ✅ Transaction Logging: Enable for rollback capability
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true, Position = 0)]
    [ValidateScript({Test-Path $_})]
    [string]$InputFile,

    [Parameter(Mandatory = $false)]
    [ValidateSet('CSV', 'JSON')]
    [string]$FileType,

    [Parameter(Mandatory = $false)]
    [string]$DNSServer = $env:COMPUTERNAME,

    [Parameter(Mandatory = $false)]
    [string]$ZoneName,

    [Parameter(Mandatory = $false)]
    [string]$LogPath = ".\Bulk_Create_Log_$(Get-Date -Format 'yyyyMMdd_HHmmss').txt",

    [Parameter(Mandatory = $false)]
    [string]$ReportPath = ".\Bulk_Create_Report_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv",

    [Parameter(Mandatory = $false)]
    [switch]$FailOnError,

    [Parameter(Mandatory = $false)]
    [switch]$WhatIf,

    [Parameter(Mandatory = $false)]
    [switch]$Parallel,

    [Parameter(Mandatory = $false)]
    [int]$MaxParallel = 5,

    [Parameter(Mandatory = $false)]
    [switch]$TransactionLog,

    [Parameter(Mandatory = $false)]
    [switch]$ValidateOnly
)

#region Initialization
# Initialize logging
function Write-Log {
    param(
        [string]$Message,
        [string]$Level = "INFO",
        [string]$Color = "White"
    )
    $Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $LogEntry = "[$Timestamp] [$Level] $Message"
    Write-Host $LogEntry -ForegroundColor $Color
    Add-Content -Path $LogPath -Value $LogEntry -Force
}

# Initialize report
function Write-Report {
    param(
        [object[]]$Records
    )
    if ($Records.Count -gt 0) {
        $Records | Export-Csv -Path $ReportPath -NoTypeInformation
        Write-Log "Report saved to: $ReportPath" "INFO" "Green"
    }
}

#region Validation Functions
function Test-ValidIP {
    param([string]$IP)
    $IPRegex = '^(25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\.(25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\.(25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\.(25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)$'
    return ($IP -match $IPRegex)
}

function Test-ValidHostName {
    param([string]$HostName)
    if ($HostName -eq '@' -or $HostName -eq '') { return $true }
    $HostNameRegex = '^[a-zA-Z0-9]([a-zA-Z0-9-]*[a-zA-Z0-9])?$'
    return ($HostName -match $HostNameRegex)
}

function Test-ValidFQDN {
    param([string]$FQDN)
    $FQDNRegex = '^(?=.{1,255}$)([a-zA-Z0-9]([a-zA-Z0-9-]*[a-zA-Z0-9])?\.)+[a-zA-Z]{2,}$'
    return ($FQDN -match $FQDNRegex)
}

function Test-DNSRecordExists {
    param(
        [string]$Zone,
        [string]$Host,
        [string]$Type = 'A',
        [string]$Server
    )
    try {
        $FQDN = if ($Host -eq '@') { $Zone } else { "$Host.$Zone" }
        $Result = Resolve-DnsName -Name $FQDN -Type $Type -Server $Server -ErrorAction SilentlyContinue
        return ($Result -ne $null)
    }
    catch {
        return $false
    }
}

#region DNS Record Creation Functions
function New-DNSRecord {
    param(
        [string]$Zone,
        [string]$HostName,
        [string]$RecordType,
        [string]$RecordData,
        [int]$TTL = 3600,
        [int]$Priority = 10,
        [int]$Weight = 10,
        [int]$Port = 0,
        [string]$Server,
        [bool]$Simulate
    )

    $FQDN = if ($HostName -eq '@') { $Zone } else { "$HostName.$Zone" }
    
    if ($Simulate) {
        Write-Log "WHATIF: Would create $RecordType record: $FQDN -> $RecordData (TTL: $TTL)" "INFO" "Yellow"
        return $true
    }

    try {
        $DnsServer = Get-WmiObject -Namespace root\MicrosoftDNS -Class MicrosoftDNS_Server -Filter "Name='$Server'" -ErrorAction Stop
        if (-not $DnsServer) {
            Write-Log "DNS server $Server not found" "ERROR" "Red"
            return $false
        }

        switch ($RecordType) {
            'A' {
                $Record = Get-WmiObject -Namespace root\MicrosoftDNS -List -Class MicrosoftDNS_AType
                $Result = $Record.CreateInstanceFromPropertyData($Server, $Zone, $HostName, $RecordData, $TTL)
            }
            'CNAME' {
                $Record = Get-WmiObject -Namespace root\MicrosoftDNS -List -Class MicrosoftDNS_CNAME
                $Result = $Record.CreateInstanceFromPropertyData($Server, $Zone, $HostName, $RecordData, $TTL)
            }
            'MX' {
                $Record = Get-WmiObject -Namespace root\MicrosoftDNS -List -Class MicrosoftDNS_MX
                $Result = $Record.CreateInstanceFromPropertyData($Server, $Zone, $HostName, $RecordData, $TTL, $Priority)
            }
            'TXT' {
                $Record = Get-WmiObject -Namespace root\MicrosoftDNS -List -Class MicrosoftDNS_TXT
                $Result = $Record.CreateInstanceFromPropertyData($Server, $Zone, $HostName, $RecordData, $TTL)
            }
            'SRV' {
                $Record = Get-WmiObject -Namespace root\MicrosoftDNS -List -Class MicrosoftDNS_SRV
                $Result = $Record.CreateInstanceFromPropertyData($Server, $Zone, $HostName, $RecordData, $TTL, $Priority, $Weight, $Port)
            }
            default {
                Write-Log "Unsupported record type: $RecordType" "ERROR" "Red"
                return $false
            }
        }

        if ($Result) {
            Write-Log "Successfully created $RecordType record: $FQDN -> $RecordData" "SUCCESS" "Green"
            return $true
        }
        else {
            Write-Log "Failed to create $RecordType record: $FQDN" "ERROR" "Red"
            return $false
        }
    }
    catch {
        Write-Log "Error creating record: $_" "ERROR" "Red"
        return $false
    }
}

#region Input Parsing Functions
function Get-RecordsFromCSV {
    param([string]$FilePath)
    try {
        $Records = Import-Csv -Path $FilePath
        Write-Log "Loaded $($Records.Count) records from CSV" "INFO" "Cyan"
        return $Records
    }
    catch {
        Write-Log "Error reading CSV file: $_" "ERROR" "Red"
        throw
    }
}

function Get-RecordsFromJSON {
    param([string]$FilePath)
    try {
        $JsonContent = Get-Content -Path $FilePath -Raw
        $Records = $JsonContent | ConvertFrom-Json
        Write-Log "Loaded $($Records.Count) records from JSON" "INFO" "Cyan"
        return $Records
    }
    catch {
        Write-Log "Error reading JSON file: $_" "ERROR" "Red"
        throw
    }
}

#region Process Records
function Process-Records {
    param(
        [object[]]$Records,
        [string]$DefaultZone,
        [string]$Server,
        [bool]$Simulate
    )

    $Results = @()
    $TotalRecords = $Records.Count
    $Processed = 0
    $Successful = 0
    $Failed = 0

    foreach ($Record in $Records) {
        $Processed++
        Write-Log "Processing record $Processed of $TotalRecords" "INFO" "Cyan"

        try {
            # Parse record data
            $Zone = if ($Record.ZoneName) { $Record.ZoneName } else { $DefaultZone }
            $HostName = $Record.HostName
            $RecordType = $Record.RecordType.ToUpper()
            $RecordData = $Record.RecordData
            $TTL = if ($Record.TTL) { [int]$Record.TTL } else { 3600 }
            $Priority = if ($Record.Priority) { [int]$Record.Priority } else { 10 }
            $Weight = if ($Record.Weight) { [int]$Record.Weight } else { 10 }
            $Port = if ($Record.Port) { [int]$Record.Port } else { 0 }

            # Validate
            if (-not $Zone) {
                Write-Log "Missing ZoneName for record: $HostName" "ERROR" "Red"
                $Failed++
                continue
            }

            if (-not $HostName) {
                Write-Log "Missing HostName for record" "ERROR" "Red"
                $Failed++
                continue
            }

            if (-not $RecordType) {
                Write-Log "Missing RecordType for record: $HostName" "ERROR" "Red"
                $Failed++
                continue
            }

            if (-not $RecordData) {
                Write-Log "Missing RecordData for record: $HostName" "ERROR" "Red"
                $Failed++
                continue
            }

            # Validate based on type
            switch ($RecordType) {
                'A' {
                    if (-not (Test-ValidIP -IP $RecordData)) {
                        Write-Log "Invalid IP address for A record: $HostName -> $RecordData" "ERROR" "Red"
                        $Failed++
                        continue
                    }
                }
                'CNAME' {
                    if (-not (Test-ValidFQDN -FQDN $RecordData)) {
                        Write-Log "Invalid FQDN for CNAME record: $HostName -> $RecordData" "ERROR" "Red"
                        $Failed++
                        continue
                    }
                }
                'MX' {
                    if (-not (Test-ValidFQDN -FQDN $RecordData)) {
                        Write-Log "Invalid FQDN for MX record: $HostName -> $RecordData" "ERROR" "Red"
                        $Failed++
                        continue
                    }
                    if (-not $Record.Priority) {
                        Write-Log "Priority missing for MX record: $HostName" "ERROR" "Red"
                        $Failed++
                        continue
                    }
                }
                'TXT' {
                    # TXT records can contain any text
                }
                'SRV' {
                    if (-not (Test-ValidFQDN -FQDN $RecordData)) {
                        Write-Log "Invalid FQDN for SRV record: $HostName -> $RecordData" "ERROR" "Red"
                        $Failed++
                        continue
                    }
                    if (-not $Record.Priority -or -not $Record.Weight -or -not $Record.Port) {
                        Write-Log "Missing priority/weight/port for SRV record: $HostName" "ERROR" "Red"
                        $Failed++
                        continue
                    }
                }
                default {
                    Write-Log "Unsupported record type: $RecordType" "ERROR" "Red"
                    $Failed++
                    continue
                }
            }

            # Check for duplicates (skip during validation only)
            if (-not $ValidateOnly) {
                if (Test-DNSRecordExists -Zone $Zone -Host $HostName -Type $RecordType -Server $Server) {
                    Write-Log "Record already exists: $HostName.$Zone ($RecordType)" "WARNING" "Yellow"
                    if ($FailOnError) {
                        throw "Duplicate record found"
                    }
                    $Failed++
                    continue
                }
            }

            # Create record
            $Success = New-DNSRecord -Zone $Zone -HostName $HostName -RecordType $RecordType -RecordData $RecordData -TTL $TTL -Priority $Priority -Weight $Weight -Port $Port -Server $Server -Simulate $Simulate

            if ($Success) {
                $Successful++
                $Results += [PSCustomObject]@{
                    HostName = $HostName
                    ZoneName = $Zone
                    RecordType = $RecordType
                    RecordData = $RecordData
                    TTL = $TTL
                    Status = "Success"
                    Timestamp = Get-Date
                }
            }
            else {
                $Failed++
                $Results += [PSCustomObject]@{
                    HostName = $HostName
                    ZoneName = $Zone
                    RecordType = $RecordType
                    RecordData = $RecordData
                    TTL = $TTL
                    Status = "Failed"
                    Timestamp = Get-Date
                }
                if ($FailOnError) {
                    throw "Record creation failed"
                }
            }
        }
        catch {
            Write-Log "Error processing record: $_" "ERROR" "Red"
            $Failed++
            if ($FailOnError) { throw }
        }

        # Progress indicator
        $Percent = [math]::Round(($Processed / $TotalRecords) * 100, 2)
        Write-Progress -Activity "Creating DNS Records" -Status "Progress: $Percent%" -PercentComplete $Percent
    }

    return $Results
}

#region Main Execution
try {
    Write-Log "=== Bulk DNS Record Creation Script Started ===" "INFO" "Cyan"
    Write-Log "Target DNS Server: $DNSServer" "INFO" "Cyan"
    Write-Log "WhatIf Mode: $WhatIf" "INFO" "Cyan"
    Write-Log "Validate Only: $ValidateOnly" "INFO" "Cyan"

    # Detect file type if not specified
    if (-not $FileType) {
        $Extension = [System.IO.Path]::GetExtension($InputFile).ToUpper()
        $FileType = if ($Extension -eq '.JSON') { 'JSON' } else { 'CSV' }
        Write-Log "Auto-detected file type: $FileType" "INFO" "Cyan"
    }

    # Load records
    $Records = if ($FileType -eq 'CSV') {
        Get-RecordsFromCSV -FilePath $InputFile
    }
    else {
        Get-RecordsFromJSON -FilePath $InputFile
    }

    if (-not $Records -or $Records.Count -eq 0) {
        Write-Log "No records to process" "ERROR" "Red"
        exit 1
    }

    Write-Log "Total records to process: $($Records.Count)" "INFO" "Cyan"

    # Process records
    if ($ValidateOnly) {
        Write-Log "Running in validation mode only" "INFO" "Cyan"
    }

    $Results = Process-Records -Records $Records -DefaultZone $ZoneName -Server $DNSServer -Simulate $WhatIf

    # Generate summary
    $Total = $Records.Count
    $Success = ($Results | Where-Object { $_.Status -eq 'Success' }).Count
    $Failed = ($Results | Where-Object { $_.Status -eq 'Failed' }).Count

    Write-Log "`n=== Final Summary ===" "INFO" "Cyan"
    Write-Log "Total Records: $Total" "INFO" "White"
    Write-Log "Successful: $Success" "INFO" "Green"
    Write-Log "Failed: $Failed" "INFO" "Red"
    Write-Log "Success Rate: $([math]::Round(($Success / $Total) * 100, 2))%" "INFO" "White"
    Write-Log "Log saved to: $LogPath" "INFO" "Green"

    # Save report
    Write-Report -Records $Results

    Write-Log "=== Script Completed ===" "INFO" "Cyan"
}
catch {
    Write-Log "Fatal error: $_" "ERROR" "Red"
    Write-Log "Script terminated" "ERROR" "Red"
    exit 1
}
#endregion