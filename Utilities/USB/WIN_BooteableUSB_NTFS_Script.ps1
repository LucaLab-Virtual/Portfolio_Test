# ================================

# Windows Server 2022 USB Creator

# ================================

# --- Self-elevate ---

if (-not ([Security.Principal.WindowsPrincipal] `     [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole(`
[Security.Principal.WindowsBuiltInRole]::Administrator)) {

```
Start-Process powershell "-ExecutionPolicy Bypass -File `"$PSCommandPath`"" -Verb RunAs
exit
```

}

# --- VARIABLES ---

$isoPath = "D:\Users\Lucarez\Downloads\WinSer22.iso"
$diskNumber = 1   # Adjust if needed

Write-Host "=== Preparing USB Disk ===" -ForegroundColor Cyan

# --- Disk Preparation ---

Get-Disk $diskNumber | Set-Disk -IsReadOnly $false
Get-Disk $diskNumber | Clear-Disk -RemoveData -Confirm:$false

Initialize-Disk -Number $diskNumber -PartitionStyle MBR -ErrorAction SilentlyContinue

$partition = New-Partition -DiskNumber $diskNumber -UseMaximumSize -IsActive -AssignDriveLetter

$usbDrive = ($partition | Get-Volume).DriveLetter + ":"

Format-Volume -DriveLetter $partition.DriveLetter -FileSystem NTFS -NewFileSystemLabel "WS2022" -Confirm:$false

Write-Host "USB Drive: $usbDrive"

# --- Mount ISO ---
Mount-DiskImage -ImagePath $isoPath
Start-Sleep -Seconds 2

# --- Detect ISO Drive (FIXED) ---
$isoDrive = (Get-DiskImage -ImagePath $isoPath | Get-Volume).DriveLetter + ":"

Write-Host "ISO Drive: $isoDrive"

# --- Validation ---
if (-not $isoDrive -or -not (Test-Path $isoDrive)) {
    Write-Error "ISO mount failed. Cannot continue."
    exit
}

# --- Copy Files ---
robocopy $isoDrive $usbDrive /E /COPY:DAT /R:0 /W:0 /V /ETA

# --- Make Bootable ---

Write-Host "=== Making USB Bootable ===" -ForegroundColor Cyan
$bootSect = "$isoDrive\boot\bootsect.exe"
& $bootSect /nt60 $usbDrive

# --- Validation ---

Write-Host "=== Final Validation ===" -ForegroundColor Green
Get-ChildItem $usbDrive

# --- Unmount ISO ---

Write-Host "=== Cleaning Up ===" -ForegroundColor Yellow
Dismount-DiskImage -ImagePath $isoPath

Write-Host "=== DONE ===" -ForegroundColor Green
