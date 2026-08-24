#Requires -Modules ActiveDirectory

[CmdletBinding()]
param(
    [string]$CsvPath = "D:\Delete-ADUsers.csv",
    [string]$LogFolder = "D:\Logs"
)

Import-Module ActiveDirectory

New-Item -ItemType Directory -Force -Path $LogFolder | Out-Null

$DateStamp = Get-Date -Format "yyyyMMdd_HHmmss"

$BackupCSV = Join-Path $LogFolder "DeletedUsers_Backup_$DateStamp.csv"
$LogFile   = Join-Path $LogFolder "DeleteUsers_$DateStamp.txt"

$DisabledOU = "OU=Dis_Users,OU=_Disabled Objects,DC=lucalab,DC=local"

$Backup  = @()
$Success = @()
$Failed  = @()

$Users = Import-Csv $CsvPath

foreach ($User in $Users)
{
    try
    {
        $ADUser = Get-ADUser `
            -Identity $User.SamAccountName `
            -Properties * `
            -ErrorAction Stop

        # Safety Check #1
        if ($ADUser.Enabled)
        {
            throw "User account is still enabled."
        }

        # Safety Check #2
        if ($ADUser.DistinguishedName -notlike "*$DisabledOU")
        {
            throw "User is not located inside the Disabled Users OU."
        }

        # Backup before deletion
        $Groups = Get-ADPrincipalGroupMembership $ADUser |
                  Select-Object -ExpandProperty Name

        $Backup += [PSCustomObject]@{
            Name              = $ADUser.Name
            DisplayName       = $ADUser.DisplayName
            SamAccountName    = $ADUser.SamAccountName
            UserPrincipalName = $ADUser.UserPrincipalName
            DistinguishedName = $ADUser.DistinguishedName
            Department        = $ADUser.Department
            Title             = $ADUser.Title
            Company           = $ADUser.Company
            Manager           = $ADUser.Manager
            Groups            = ($Groups -join ";")
            DeletedOn         = Get-Date
        }

        # Delete the account
        Remove-ADUser `
            -Identity $ADUser `
            -Confirm:$false

        $Success += $ADUser.SamAccountName
    }
    catch
    {
        $Failed += [PSCustomObject]@{
            User  = $User.SamAccountName
            Error = $_.Exception.Message
        }
    }
}

# Export backup
$Backup | Export-Csv $BackupCSV -NoTypeInformation -Encoding UTF8

# Write log
"Delete AD Users Log" | Out-File $LogFile
"===================" | Add-Content $LogFile
"" | Add-Content $LogFile

Add-Content $LogFile "Successfully Deleted:"
foreach ($User in $Success)
{
    Add-Content $LogFile " - $User"
}

if ($Failed.Count -gt 0)
{
    "" | Add-Content $LogFile
    Add-Content $LogFile "Failures:"

    foreach ($Item in $Failed)
    {
        Add-Content $LogFile " - $($Item.User): $($Item.Error)"
    }
}

Write-Host ""
Write-Host "Deletion completed."
Write-Host "Deleted: $($Success.Count)"
Write-Host "Failed : $($Failed.Count)"
Write-Host ""
Write-Host "Backup CSV: $BackupCSV"
Write-Host "Log File  : $LogFile"