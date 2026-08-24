<#
.SYNOPSIS
    Exports Windows DNS zones to multiple formats with advanced filtering and reporting.

.DESCRIPTION
    This script exports DNS zone data to various formats including CSV, JSON, XML, and HTML.
    Supports filtering by record type, wildcard patterns, and includes zone statistics.
    Perfect for backups, documentation, migrations, and DNS analysis.

.PARAMETER ZoneName
    Name of the DNS zone to export (e.g., "contoso.com"). Use '*' for all zones

.PARAMETER DNSServer
    DNS server to export from. Default is localhost

.PARAMETER OutputPath
    Directory path for output files. Default: .\DNS_Exports\

.PARAMETER OutputFormat
    Output format: CSV, JSON, XML, HTML, or ALL. Default: CSV

.PARAMETER RecordTypes
    Filter by record types (A, CNAME, MX, TXT, SRV, etc.). Default: All types

.PARAMETER IncludeTimestamp
    Include timestamp in output filename

.PARAMETER IncludeStatistics
    Generate statistical report

.PARAMETER FilterPattern
    Wildcard pattern to filter records (e.g., "*.contoso.com")

.PARAMETER ExcludePattern
    Wildcard pattern to exclude records

.PARAMETER ExportSOA
    Include SOA records in export

.PARAMETER ExportNS
    Include NS records in export

.PARAMETER CompressOutput
    Create compressed ZIP archive of exported files

.PARAMETER WhatIf
    Preview export without creating files

.EXAMPLE
    .\Export_DNS_Zone.ps1 -ZoneName "contoso.com" -OutputFormat ALL

.EXAMPLE
    .\Export_DNS_Zone.ps1 -ZoneName "*" -RecordTypes A,CNAME -FilterPattern "*.contoso.com"

.EXAMPLE
    .\Export_DNS_Zone.ps1 -ZoneName "contoso.com" -OutputFormat JSON -IncludeStatistics -CompressOutput

.EXAMPLE
    .\Export_DNS_Zone.ps1 -ZoneName "internal.local" -DNSServer "DC01" -ExcludePattern "*.test.*"

.NOTES
    Author: Portfolio Script
    Version: 3.0
    Requires: Windows DNS Server role, PowerShell 5.1+
    Output Formats: CSV, JSON, XML, HTML, ALL
#>

<#
Usage Examples:
1. Basic Zone Export (CSV)
powershell

.\Export_DNS_Zone.ps1 -ZoneName "contoso.com"

2. Export All Zones with Timestamp
powershell

.\Export_DNS_Zone.ps1 -ZoneName "*" -IncludeTimestamp

3. Multi-Format Export with Statistics
powershell

.\Export_DNS_Zone.ps1 -ZoneName "contoso.com" -OutputFormat ALL -IncludeStatistics

4. Filtered Export (Specific Record Types)
powershell

.\Export_DNS_Zone.ps1 -ZoneName "contoso.com" -RecordTypes A,CNAME,MX -FilterPattern "*.contoso.com"

5. Export with Compression and Exclusions
powershell

.\Export_DNS_Zone.ps1 -ZoneName "internal.local" -OutputFormat JSON -ExcludePattern "*.test.*" -CompressOutput

6. Export with SOA and NS Records
powershell

.\Export_DNS_Zone.ps1 -ZoneName "contoso.com" -ExportSOA -ExportNS -OutputFormat HTML

Output Examples:
CSV Output (DNS_Export_contoso.com_CSV.csv):
csv

ZoneName,RecordName,RecordType,RecordData,TTL,TimeStamp,Server
contoso.com,@,SOA,ns1.contoso.com,3600,2024-01-15 10:00:00,DC01
contoso.com,www,A,192.168.1.100,3600,2024-01-15 10:00:00,DC01
contoso.com,mail,MX,mail.contoso.com,3600,2024-01-15 10:00:00,DC01

JSON Output (DNS_Export_contoso.com_JSON.json):
json

[
  {
    "ZoneName": "contoso.com",
    "RecordName": "www",
    "RecordType": "A",
    "RecordData": "192.168.1.100",
    "TTL": 3600,
    "Server": "DC01",
    "AdditionalData": {
      "IPAddress": "192.168.1.100"
    }
  }
]

HTML Output with Statistics:

    Beautiful formatted table with colored record types

    Statistics dashboard with key metrics

    Record type breakdown with percentages

    Responsive design for viewing in any browser

Features:

    ✅ Multiple Export Formats: CSV, JSON, XML, HTML

    ✅ Zone Filtering: Single zone, multiple zones, or all zones

    ✅ Record Filtering: By type, wildcard patterns, exclusions

    ✅ Statistics Generation: Detailed zone analytics

    ✅ Compression: ZIP archive of exported files

    ✅ Timestamp Support: Unique filenames

    ✅ SOA/NS Record Support: Optional inclusion

    ✅ Colorized Console Output: Easy to read logging

    ✅ WhatIf Mode: Preview without exporting

    ✅ Error Handling: Graceful error recovery

    ✅ Progress Tracking: Visual feedback during export

Statistics Generated:

    Total records count

    Record type breakdown with percentages

    Unique hosts and IPs

    Average, max, min TTL values

    Top domains by record count

    Export timestamp and server information
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true, Position = 0)]
    [string]$ZoneName,

    [Parameter(Mandatory = $false)]
    [string]$DNSServer = $env:COMPUTERNAME,

    [Parameter(Mandatory = $false)]
    [string]$OutputPath = ".\DNS_Exports",

    [Parameter(Mandatory = $false)]
    [ValidateSet('CSV', 'JSON', 'XML', 'HTML', 'ALL')]
    [string]$OutputFormat = 'CSV',

    [Parameter(Mandatory = $false)]
    [string[]]$RecordTypes,

    [Parameter(Mandatory = $false)]
    [switch]$IncludeTimestamp,

    [Parameter(Mandatory = $false)]
    [switch]$IncludeStatistics,

    [Parameter(Mandatory = $false)]
    [string]$FilterPattern,

    [Parameter(Mandatory = $false)]
    [string]$ExcludePattern,

    [Parameter(Mandatory = $false)]
    [switch]$ExportSOA,

    [Parameter(Mandatory = $false)]
    [switch]$ExportNS,

    [Parameter(Mandatory = $false)]
    [switch]$CompressOutput,

    [Parameter(Mandatory = $false)]
    [switch]$WhatIf
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
}

function Initialize-OutputDirectory {
    param([string]$Path)
    if (-not (Test-Path $Path)) {
        New-Item -Path $Path -ItemType Directory -Force | Out-Null
        Write-Log "Created output directory: $Path" "INFO" "Cyan"
    }
}

function Get-Timestamp {
    return Get-Date -Format "yyyyMMdd_HHmmss"
}

#region DNS Zone Retrieval Functions
function Get-DNSZones {
    param(
        [string]$Server,
        [string]$ZonePattern
    )
    try {
        $Zones = Get-WmiObject -Namespace root\MicrosoftDNS -Class MicrosoftDNS_Zone -Filter "Name LIKE '%$ZonePattern%'" -ErrorAction Stop
        
        if (-not $Zones) {
            Write-Log "No zones found matching pattern: $ZonePattern" "WARNING" "Yellow"
            return @()
        }
        
        Write-Log "Found $($Zones.Count) DNS zones on server $Server" "INFO" "Cyan"
        return $Zones
    }
    catch {
        Write-Log "Error retrieving DNS zones: $_" "ERROR" "Red"
        throw
    }
}

function Get-DNSRecords {
    param(
        [string]$ZoneName,
        [string]$Server,
        [string[]]$Types,
        [bool]$IncludeSOA,
        [bool]$IncludeNS
    )
    
    try {
        $Records = @()
        
        # Get all records from the zone
        $DnsRecords = Get-WmiObject -Namespace root\MicrosoftDNS -Class MicrosoftDNS_ResourceRecord -Filter "ContainerName='$ZoneName'" -ErrorAction Stop
        
        if (-not $DnsRecords) {
            Write-Log "No records found in zone: $ZoneName" "WARNING" "Yellow"
            return @()
        }
        
        # Filter records by type if specified
        if ($Types -and $Types.Count -gt 0) {
            $DnsRecords = $DnsRecords | Where-Object { $_.__CLASS -replace 'MicrosoftDNS_', '' -in $Types }
        }
        
        # Filter SOA records
        if (-not $IncludeSOA) {
            $DnsRecords = $DnsRecords | Where-Object { $_.__CLASS -ne 'MicrosoftDNS_SOA' }
        }
        
        # Filter NS records
        if (-not $IncludeNS) {
            $DnsRecords = $DnsRecords | Where-Object { $_.__CLASS -ne 'MicrosoftDNS_NS' }
        }
        
        # Process each record
        foreach ($Record in $DnsRecords) {
            try {
                $RecordType = $Record.__CLASS -replace 'MicrosoftDNS_', ''
                $RecordData = $null
                $AdditionalData = @{}
                
                # Extract data based on record type
                switch ($RecordType) {
                    'AType' {
                        $RecordType = 'A'
                        $RecordData = $Record.IPAddress
                        $AdditionalData = @{
                            'IPAddress' = $Record.IPAddress
                        }
                    }
                    'CNAME' {
                        $RecordType = 'CNAME'
                        $RecordData = $Record.PrimaryName
                        $AdditionalData = @{
                            'PrimaryName' = $Record.PrimaryName
                        }
                    }
                    'MX' {
                        $RecordType = 'MX'
                        $RecordData = $Record.MXExchange
                        $AdditionalData = @{
                            'Priority' = $Record.Preference
                            'Exchange' = $Record.MXExchange
                        }
                    }
                    'TXT' {
                        $RecordType = 'TXT'
                        $RecordData = $Record.TextualData
                        $AdditionalData = @{
                            'TextData' = $Record.TextualData
                        }
                    }
                    'SRV' {
                        $RecordType = 'SRV'
                        $RecordData = "$($Record.SRVDomainName):$($Record.SRVPort)"
                        $AdditionalData = @{
                            'Priority' = $Record.SRVPriority
                            'Weight' = $Record.SRVWeight
                            'Port' = $Record.SRVPort
                            'DomainName' = $Record.SRVDomainName
                        }
                    }
                    'SOA' {
                        $RecordType = 'SOA'
                        $RecordData = $Record.PrimaryServer
                        $AdditionalData = @{
                            'PrimaryServer' = $Record.PrimaryServer
                            'ResponsiblePerson' = $Record.ResponsiblePerson
                            'SerialNumber' = $Record.SerialNumber
                            'RefreshInterval' = $Record.RefreshInterval
                            'RetryInterval' = $Record.RetryInterval
                            'ExpireLimit' = $Record.ExpireLimit
                            'MinimumTTL' = $Record.MinimumTTL
                        }
                    }
                    'NS' {
                        $RecordType = 'NS'
                        $RecordData = $Record.NSHost
                        $AdditionalData = @{
                            'NSHost' = $Record.NSHost
                        }
                    }
                    'PTR' {
                        $RecordType = 'PTR'
                        $RecordData = $Record.PTRDomainName
                        $AdditionalData = @{
                            'DomainName' = $Record.PTRDomainName
                        }
                    }
                    default {
                        $RecordData = $Record.__PATH
                        $AdditionalData = @{
                            'RawData' = $Record.__PATH
                        }
                    }
                }
                
                # Create record object
                $RecordObject = [PSCustomObject]@{
                    ZoneName = $ZoneName
                    RecordName = $Record.OwnerName
                    RecordType = $RecordType
                    RecordData = $RecordData
                    TTL = $Record.TTL
                    TimeStamp = $Record.TimeStamp
                    Server = $Server
                    AdditionalData = $AdditionalData
                }
                
                $Records += $RecordObject
            }
            catch {
                Write-Log "Error processing record: $_" "WARNING" "Yellow"
            }
        }
        
        Write-Log "Retrieved $($Records.Count) records from zone: $ZoneName" "INFO" "Cyan"
        return $Records
    }
    catch {
        Write-Log "Error retrieving DNS records: $_" "ERROR" "Red"
        throw
    }
}

#region Filtering Functions
function Apply-Filters {
    param(
        [object[]]$Records,
        [string]$FilterPattern,
        [string]$ExcludePattern
    )
    
    $FilteredRecords = $Records
    
    # Apply include filter
    if ($FilterPattern) {
        $FilteredRecords = $FilteredRecords | Where-Object { 
            $_.RecordName -like $FilterPattern -or 
            $_.RecordData -like $FilterPattern
        }
        Write-Log "Applied include filter: $FilterPattern ($($FilteredRecords.Count) records retained)" "INFO" "Cyan"
    }
    
    # Apply exclude filter
    if ($ExcludePattern) {
        $FilteredRecords = $FilteredRecords | Where-Object { 
            $_.RecordName -notlike $ExcludePattern -and 
            $_.RecordData -notlike $ExcludePattern
        }
        Write-Log "Applied exclude filter: $ExcludePattern ($($FilteredRecords.Count) records retained)" "INFO" "Cyan"
    }
    
    return $FilteredRecords
}

#region Statistics Functions
function Get-DNSStatistics {
    param(
        [object[]]$Records,
        [string]$ZoneName
    )
    
    $Stats = @{
        ZoneName = $ZoneName
        TotalRecords = $Records.Count
        RecordTypes = $Records | Group-Object RecordType | Select-Object Name, Count
        TopDomains = $Records | 
            ForEach-Object { ($_.RecordName -split '\.')[-2..-1] -join '.' } | 
            Group-Object | 
            Sort-Object Count -Descending | 
            Select-Object -First 5 Name, Count
        UniqueHosts = ($Records | Where-Object { $_.RecordType -eq 'A' } | Select-Object -ExpandProperty RecordName -Unique).Count
        UniqueIPs = ($Records | Where-Object { $_.RecordType -eq 'A' } | Select-Object -ExpandProperty RecordData -Unique).Count
        TotalTTL = [math]::Round(($Records | Measure-Object -Property TTL -Average).Average, 0)
        MaxTTL = ($Records | Measure-Object -Property TTL -Maximum).Maximum
        MinTTL = ($Records | Measure-Object -Property TTL -Minimum).Minimum
        Timestamp = Get-Date
    }
    
    return $Stats
}

#region Export Functions
function Export-ToCSV {
    param(
        [object[]]$Records,
        [string]$FilePath
    )
    try {
        $Records | Export-Csv -Path $FilePath -NoTypeInformation -Encoding UTF8
        Write-Log "Exported CSV to: $FilePath" "SUCCESS" "Green"
        return $true
    }
    catch {
        Write-Log "Error exporting CSV: $_" "ERROR" "Red"
        return $false
    }
}

function Export-ToJSON {
    param(
        [object[]]$Records,
        [string]$FilePath
    )
    try {
        $Records | ConvertTo-Json -Depth 10 | Out-File -FilePath $FilePath -Encoding UTF8
        Write-Log "Exported JSON to: $FilePath" "SUCCESS" "Green"
        return $true
    }
    catch {
        Write-Log "Error exporting JSON: $_" "ERROR" "Red"
        return $false
    }
}

function Export-ToXML {
    param(
        [object[]]$Records,
        [string]$FilePath
    )
    try {
        $Records | Export-Clixml -Path $FilePath -Depth 10
        Write-Log "Exported XML to: $FilePath" "SUCCESS" "Green"
        return $true
    }
    catch {
        Write-Log "Error exporting XML: $_" "ERROR" "Red"
        return $false
    }
}

function Export-ToHTML {
    param(
        [object[]]$Records,
        [string]$FilePath,
        [hashtable]$Statistics
    )
    try {
        $HTML = @"
<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <title>DNS Zone Export: $(if($Records[0]){$Records[0].ZoneName})</title>
    <style>
        body { font-family: Arial, sans-serif; margin: 20px; }
        h1 { color: #333; border-bottom: 2px solid #333; }
        .stats { background: #f4f4f4; padding: 15px; margin: 15px 0; border-radius: 5px; }
        .stats-grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(200px, 1fr)); gap: 10px; }
        .stat-item { background: white; padding: 10px; border-radius: 3px; box-shadow: 0 2px 4px rgba(0,0,0,0.1); }
        .stat-label { font-weight: bold; color: #666; font-size: 0.9em; }
        .stat-value { font-size: 1.2em; margin-top: 5px; }
        table { border-collapse: collapse; width: 100%; margin-top: 20px; }
        th, td { border: 1px solid #ddd; padding: 8px; text-align: left; }
        th { background-color: #f2f2f2; font-weight: bold; }
        tr:nth-child(even) { background-color: #f9f9f9; }
        tr:hover { background-color: #f5f5f5; }
        .type-A { color: #2196F3; }
        .type-CNAME { color: #4CAF50; }
        .type-MX { color: #FF9800; }
        .type-TXT { color: #9C27B0; }
        .type-SRV { color: #F44336; }
        .type-SOA { color: #795548; }
        .type-NS { color: #607D8B; }
        .type-PTR { color: #E91E63; }
        .footer { margin-top: 30px; color: #999; font-size: 0.8em; text-align: center; }
    </style>
</head>
<body>
    <h1>DNS Zone Export: $(if($Records[0]){$Records[0].ZoneName})</h1>
    <p>Exported from server: $(if($Records[0]){$Records[0].Server})</p>
    <p>Export date: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')</p>
"@

        # Add statistics
        if ($Statistics) {
            $HTML += @"
    <div class="stats">
        <h2>Zone Statistics</h2>
        <div class="stats-grid">
            <div class="stat-item">
                <div class="stat-label">Total Records</div>
                <div class="stat-value">$($Statistics.TotalRecords)</div>
            </div>
            <div class="stat-item">
                <div class="stat-label">Unique Hosts</div>
                <div class="stat-value">$($Statistics.UniqueHosts)</div>
            </div>
            <div class="stat-item">
                <div class="stat-label">Unique IPs</div>
                <div class="stat-value">$($Statistics.UniqueIPs)</div>
            </div>
            <div class="stat-item">
                <div class="stat-label">Average TTL</div>
                <div class="stat-value">$($Statistics.TotalTTL) sec</div>
            </div>
            <div class="stat-item">
                <div class="stat-label">Record Types</div>
                <div class="stat-value">$($Statistics.RecordTypes.Count)</div>
            </div>
        </div>
    </div>
"@
            
            # Record type breakdown
            $HTML += @"
    <div class="stats">
        <h3>Record Type Breakdown</h3>
        <table>
            <thead>
                <tr>
                    <th>Record Type</th>
                    <th>Count</th>
                    <th>Percentage</th>
                </tr>
            </thead>
            <tbody>
"@
            foreach ($Type in $Statistics.RecordTypes) {
                $Percent = [math]::Round(($Type.Count / $Statistics.TotalRecords) * 100, 2)
                $HTML += @"
                <tr>
                    <td>$($Type.Name)</td>
                    <td>$($Type.Count)</td>
                    <td>$Percent%</td>
                </tr>
"@
            }
            $HTML += @"
            </tbody>
        </table>
    </div>
"@
        }

        # Records table
        $HTML += @"
    <h2>DNS Records</h2>
    <table>
        <thead>
            <tr>
                <th>Record Name</th>
                <th>Type</th>
                <th>Data</th>
                <th>TTL</th>
                <th>Additional Info</th>
            </tr>
        </thead>
        <tbody>
"@
        
        foreach ($Record in $Records) {
            $TypeClass = "type-$($Record.RecordType)"
            $AdditionalInfo = ''
            if ($Record.AdditionalData) {
                $AdditionalInfo = ($Record.AdditionalData.Keys | ForEach-Object { "$_=$($Record.AdditionalData[$_])" }) -join ', '
            }
            
            $HTML += @"
            <tr>
                <td>$($Record.RecordName)</td>
                <td class="$TypeClass">$($Record.RecordType)</td>
                <td>$($Record.RecordData)</td>
                <td>$($Record.TTL)</td>
                <td>$AdditionalInfo</td>
            </tr>
"@
        }
        
        $HTML += @"
        </tbody>
    </table>
    <div class="footer">
        <p>Generated by DNS Export Script v3.0 | Total Records: $($Records.Count)</p>
    </div>
</body>
</html>
"@
        
        $HTML | Out-File -FilePath $FilePath -Encoding UTF8
        Write-Log "Exported HTML to: $FilePath" "SUCCESS" "Green"
        return $true
    }
    catch {
        Write-Log "Error exporting HTML: $_" "ERROR" "Red"
        return $false
    }
}

function Compress-ExportFiles {
    param(
        [string]$Directory,
        [string]$ArchiveName
    )
    try {
        $ArchivePath = Join-Path $Directory $ArchiveName
        Compress-Archive -Path "$Directory\*" -DestinationPath $ArchivePath -Force
        Write-Log "Created compressed archive: $ArchivePath" "SUCCESS" "Green"
        return $true
    }
    catch {
        Write-Log "Error creating archive: $_" "ERROR" "Red"
        return $false
    }
}

#region Main Execution
try {
    Write-Log "=== DNS Zone Export Script Started ===" "INFO" "Cyan"
    Write-Log "Target DNS Server: $DNSServer" "INFO" "Cyan"
    Write-Log "Zone Name: $ZoneName" "INFO" "Cyan"
    Write-Log "WhatIf Mode: $WhatIf" "INFO" "Cyan"

    # Create output directory
    Initialize-OutputDirectory -Path $OutputPath
    
    # Determine timestamp for filenames
    $Timestamp = Get-Timestamp
    $FileSuffix = if ($IncludeTimestamp) { "_$Timestamp" } else { "" }
    
    # Get zones
    $Zones = Get-DNSZones -Server $DNSServer -ZonePattern $ZoneName
    
    if (-not $Zones) {
        Write-Log "No zones found to export" "ERROR" "Red"
        exit 1
    }

    $TotalExported = 0
    $ExportFiles = @()
    
    foreach ($Zone in $Zones) {
        $ZoneName = $Zone.Name
        Write-Log "`nProcessing zone: $ZoneName" "INFO" "Cyan"
        
        # Get records
        $Records = Get-DNSRecords -ZoneName $ZoneName -Server $DNSServer -Types $RecordTypes -IncludeSOA:$ExportSOA -IncludeNS:$ExportNS
        
        if (-not $Records) {
            Write-Log "No records found in zone: $ZoneName" "WARNING" "Yellow"
            continue
        }
        
        # Apply filters
        $FilteredRecords = Apply-Filters -Records $Records -FilterPattern $FilterPattern -ExcludePattern $ExcludePattern
        
        if (-not $FilteredRecords) {
            Write-Log "No records after filtering for zone: $ZoneName" "WARNING" "Yellow"
            continue
        }
        
        # Generate statistics
        if ($IncludeStatistics) {
            $Statistics = Get-DNSStatistics -Records $FilteredRecords -ZoneName $ZoneName
            Write-Log "Zone Statistics:" "INFO" "Cyan"
            Write-Log "  Total Records: $($Statistics.TotalRecords)" "INFO" "White"
            Write-Log "  Record Types: $($Statistics.RecordTypes.Count)" "INFO" "White"
            Write-Log "  Unique Hosts: $($Statistics.UniqueHosts)" "INFO" "White"
            Write-Log "  Unique IPs: $($Statistics.UniqueIPs)" "INFO" "White"
            Write-Log "  Average TTL: $($Statistics.TotalTTL) seconds" "INFO" "White"
        }
        
        if ($WhatIf) {
            Write-Log "WHATIF: Would export $($FilteredRecords.Count) records from zone $ZoneName" "INFO" "Yellow"
            continue
        }
        
        # Determine export formats
        $Formats = if ($OutputFormat -eq 'ALL') { @('CSV', 'JSON', 'XML', 'HTML') } else { @($OutputFormat) }
        
        foreach ($Format in $Formats) {
            $BaseFileName = "DNS_Export_${ZoneName}_${Format}"
            $FileName = "${BaseFileName}${FileSuffix}.$(if($Format -eq 'HTML'){'html'}else{$Format.ToLower()})"
            $FilePath = Join-Path $OutputPath $FileName
            
            $Success = switch ($Format) {
                'CSV' { Export-ToCSV -Records $FilteredRecords -FilePath $FilePath }
                'JSON' { Export-ToJSON -Records $FilteredRecords -FilePath $FilePath }
                'XML' { Export-ToXML -Records $FilteredRecords -FilePath $FilePath }
                'HTML' { Export-ToHTML -Records $FilteredRecords -FilePath $FilePath -Statistics $Statistics }
            }
            
            if ($Success) {
                $ExportFiles += $FilePath
                $TotalExported += $FilteredRecords.Count
            }
        }
    }
    
    # Compress output if requested
    if ($CompressOutput -and $ExportFiles.Count -gt 0) {
        $ArchiveName = "DNS_Export_${Timestamp}.zip"
        Compress-ExportFiles -Directory $OutputPath -ArchiveName $ArchiveName
    }
    
    # Summary
    Write-Log "`n=== Export Summary ===" "INFO" "Cyan"
    Write-Log "Total Zones Exported: $($Zones.Count)" "INFO" "White"
    Write-Log "Total Records Exported: $TotalExported" "INFO" "White"
    Write-Log "Export Files Created: $($ExportFiles.Count)" "INFO" "White"
    Write-Log "Output Directory: $OutputPath" "INFO" "Green"
    
    Write-Log "=== Script Completed Successfully ===" "INFO" "Cyan"
}
catch {
    Write-Log "Fatal error: $_" "ERROR" "Red"
    Write-Log "Script terminated" "ERROR" "Red"
    exit 1
}
#endregion