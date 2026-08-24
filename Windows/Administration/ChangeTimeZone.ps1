# Retrives the current time zone
Get-TimeZone
# Lists all the time zones
$allTimeZones = [System.TimeZoneInfo]::GetSystemTimeZones()

foreach ($timeZone in $allTimeZones) {
    Write-Host "Time Zone ID: $($timeZone.Id)"
    Write-Host "Display Name: $($timeZone.DisplayName)"
    Write-Host ""
}
# Sets the new time zone
Set-TimeZone -Id "Eastern Standard Time"