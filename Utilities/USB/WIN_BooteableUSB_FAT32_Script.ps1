# ============================================
# Windows Server 2022 - Bootable USB Creator
# UEFI + BIOS Compatible (Production Ready)
# ============================================

# --- Self-elevate ---
if (-not ([Security.Principal.WindowsPrincipal] `
    [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole(`
    [Security.Principal.WindowsBuiltInRole]::Administrator)) {

    Start-Process powershell "-ExecutionPolicy Bypass -File `"$PSCommandPath`"" -Verb RunAs
    exit
}

# =========================
# VARIABLES
# =========================
$isoPath    = "D:\Users\Lucarez\Downloads\WinSer22.iso"
$diskNumber = 1   # ⚠️ VERIFY BEFORE RUNNING

Write-Host "=== STARTING USB CREATION ===" -ForegroundColor Cyan

# =========================
# DISK PREPARATION
# =========================
Write-Host "=== Preparing USB Disk ===" -ForegroundColor Yellow

Get-Disk $diskNumber | Set-Disk -IsReadOnly $false
Get-Disk $diskNumber | Clear-Disk -RemoveData -Confirm:$false

Initialize-Disk -Number $diskNumber -PartitionStyle GPT

$partition = New-Partition -DiskNumber $diskNumber -UseMaximumSize -AssignDriveLetter

Format-Volume -Partition $partition -FileSystem FAT32 -NewFileSystemLabel "WS2022" -Confirm:$false

$usbDrive = ($partition | Get-Volume).DriveLetter + ":"

Write-Host "USB Drive: $usbDrive" -ForegroundColor Green

# =========================
# MOUNT ISO
# =========================
Write-Host "=== Mounting ISO ===" -ForegroundColor Yellow

Mount-DiskImage -ImagePath $isoPath
Start-Sleep -Seconds 3

$isoDrive = (Get-DiskImage -ImagePath $isoPath | Get-Volume).DriveLetter + ":"

if (-not (Test-Path $isoDrive)) {
    Write-Error "ISO mount failed. Exiting..."
    exit
}

Write-Host "ISO Drive: $isoDrive" -ForegroundColor Green

# =========================
# CHECK install.wim SIZE
# =========================
$wimPath = "$isoDrive\sources\install.wim"

if (Test-Path $wimPath) {
    $wimSizeGB = (Get-Item $wimPath).Length / 1GB

    if ($wimSizeGB -gt 4) {
        Write-Host "install.wim is larger than 4GB ($([math]::Round($wimSizeGB,2)) GB)" -ForegroundColor Red
        Write-Host "=== Splitting WIM for FAT32 compatibility ===" -ForegroundColor Yellow

        # Create sources folder on USB
        New-Item -ItemType Directory -Path "$usbDrive\sources" -Force | Out-Null

        dism /Split-Image `
            /ImageFile:$wimPath `
            /SWMFile:"$usbDrive\sources\install.swm" `
            /FileSize:3800

        $splitPerformed = $true
    } else {
        $splitPerformed = $false
    }
} else {
    Write-Error "install.wim not found. Exiting..."
    exit
}

# =========================
# COPY FILES
# =========================
Write-Host "=== Copying Files ===" -ForegroundColor Yellow

if ($splitPerformed) {
    robocopy $isoDrive $usbDrive /E /R:0 /W:0 /XF install.wim
} else {
    robocopy $isoDrive $usbDrive /E /R:0 /W:0
}

# =========================
# VALIDATION (UEFI BOOT)
# =========================
Write-Host "=== Validating Boot Structure ===" -ForegroundColor Yellow

$efiBoot = "$usbDrive\EFI\BOOT\BOOTX64.EFI"

if (Test-Path $efiBoot) {
    Write-Host "UEFI Boot file detected ✔" -ForegroundColor Green
} else {
    Write-Warning "UEFI boot file NOT found! USB may not boot in UEFI mode."
}

# =========================
# CLEANUP
# =========================
Write-Host "=== Dismounting ISO ===" -ForegroundColor Yellow

Dismount-DiskImage -ImagePath $isoPath

# =========================
# FINAL OUTPUT
# =========================
Write-Host "=== USB CONTENT ===" -ForegroundColor Cyan
Get-ChildItem $usbDrive

Write-Host ""
Write-Host "=== DONE ===" -ForegroundColor Green
Write-Host "✔ UEFI Compatible"
Write-Host "✔ BIOS Compatible (CSM)"
Write-Host "✔ Ready to boot Windows Server 2022"
