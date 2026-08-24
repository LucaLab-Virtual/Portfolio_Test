# 💾 USB Utilities Toolkit (PowerShell)

![PowerShell](https://img.shields.io/badge/PowerShell-Automation-blue)
![Bootable USB](https://img.shields.io/badge/USB-Bootable%20Tools-green)
![Platform](https://img.shields.io/badge/Platform-Windows%20%7C%20Linux-lightgrey)
![Status](https://img.shields.io/badge/Status-Production%20Ready-success)

---

## 🎯 Overview

This project provides a **collection of PowerShell scripts** designed to create and manage **bootable USB drives** for both Windows and Linux environments.

It is intended for IT professionals who need a **fast, repeatable, and reliable way** to prepare USB media for installations, recovery, and deployment tasks.

---

## 🧠 Architecture

```text
        +----------------------+
        |   IT Administrator   |
        |   (PowerShell CLI)   |
        +----------+-----------+
                   |
                   | Execute Script
                   v
        +----------------------+
        | PowerShell Scripts   |
        | USB Preparation Tool |
        +----------+-----------+
                   |
        +----------+-----------+
        | Disk Management      |
        | (Format / Partition) |
        +----------+-----------+
                   |
                   v
        +----------------------+
        | Bootable USB Device  |
        | FAT32 / NTFS         |
        +----------+-----------+
                   |
                   v
        +----------------------+
        | Target System        |
        | BIOS / UEFI Boot     |
        +----------------------+
```

---

## 📂 Project Structure

```text
USB-Utilities/
│
├── LINUX_BooteableUSB_BIOS-UEFI.ps1
├── LINUX_BooteableUSB_UEFIAdaptedVersion.ps1
├── WIN_BooteableUSB_FAT32.ps1
├── WIN_BooteableUSB_FAT32_Script.ps1
├── WIN_BooteableUSB_NTFS.ps1
├── WIN_BooteableUSB_NTFS_Script.ps1
│
└── README.md
```

---

## 🧠 What is This Toolkit?

This toolkit provides **automated USB preparation scripts** that allow you to:

* Create bootable USB drives for Windows and Linux
* Support both BIOS and UEFI systems
* Choose between FAT32 and NTFS file systems
* Standardize USB deployment procedures
* Reduce manual errors during setup

---

## 🚀 Features

### 💻 Windows Bootable USB

* FAT32 format (UEFI compatible)
* NTFS format (large file support)
* Automated disk preparation
* Scripted deployment workflow

### 🐧 Linux Bootable USB

* BIOS + UEFI hybrid support
* UEFI-adapted version available
* Simplified ISO-to-USB process

### ⚙️ Automation

* Repeatable execution
* Reduced manual intervention
* Faster deployment across multiple machines

---

## ⚙️ Script Overview

### 🐧 Linux Scripts

* `LINUX_BooteableUSB_BIOS-UEFI.ps1`
  → Creates a bootable USB compatible with both BIOS and UEFI systems

* `LINUX_BooteableUSB_UEFIAdaptedVersion.ps1`
  → Optimized specifically for UEFI environments

---

### 💻 Windows Scripts

* `WIN_BooteableUSB_FAT32.ps1`
  → Prepares USB using FAT32 (UEFI-safe)

* `WIN_BooteableUSB_FAT32_Script.ps1`
  → Scripted/automated version of FAT32 setup

* `WIN_BooteableUSB_NTFS.ps1`
  → Prepares USB using NTFS (supports large install files)

* `WIN_BooteableUSB_NTFS_Script.ps1`
  → Automated NTFS deployment version

---

## ▶️ How to Run

### Requirements:

* Windows 10/11
* PowerShell 5.1+
* Administrator privileges
* USB drive (recommended ≥ 8GB)

---

### Run a Script:

```powershell
powershell.exe -ExecutionPolicy Bypass -File .\WIN_BooteableUSB_FAT32.ps1
```

---

## ⚙️ How It Works

1. Script initializes and checks permissions
2. USB disk is selected and prepared
3. Disk is formatted (FAT32 or NTFS)
4. Boot files are copied to USB
5. USB becomes bootable
6. Ready for system installation

---

## 📁 Example Use Cases

```text
Create Windows 11 USB → FAT32 → UEFI laptop
Create Windows Server USB → NTFS → large ISO
Create Ubuntu USB → BIOS/UEFI hybrid
```

---

## ⚠️ Notes

* ⚠️ **ALL DATA ON USB WILL BE ERASED**
* FAT32 has a **4GB file size limit**
* NTFS may not boot on all UEFI systems
* Always verify disk selection before execution

---

## 🧪 Best Practices

* Test scripts on non-critical USB devices
* Confirm disk number before formatting
* Use FAT32 for maximum compatibility
* Use NTFS only when required (large files)

---

## 🚀 Future Improvements

* Multi-ISO USB support
* GUI interface (optional)
* Integration with deployment tools (MDT / SCCM)
* Logging system for audit tracking
* Automatic ISO detection

---

## 👨‍💻 Author

Luis Carlos Ramirez
IT Support Specialist / Cloud & Systems Administrator

---
