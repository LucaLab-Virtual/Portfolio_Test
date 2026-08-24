<#
.SYNOPSIS
    Enables Nested Virtualization for a Hyper-V virtual machine.

.DESCRIPTION
    By default, Hyper-V does not expose the host CPU's hardware virtualization
    extensions (Intel VT-x or AMD-V) to guest virtual machines.

    As a result, applications and features inside the guest operating system
    that require hardware virtualization will not work unless Nested
    Virtualization is enabled.

    This script enables virtualization extensions for a specific VM and
    provides commands to verify the configuration from both the host and
    the guest operating system.

    Common scenarios requiring Nested Virtualization include:

        • Running Hyper-V inside a virtual machine
        • Using Windows Sandbox inside the guest
        • Installing Docker Desktop with Hyper-V or WSL2
        • Running Android Studio Emulator
        • Running VMware Workstation or VirtualBox inside the VM
        • Performing virtualization and cloud labs

.PREREQUISITES

    1. The virtual machine must be powered OFF.
    2. Run this script from an elevated PowerShell session on the Hyper-V host.
    3. Replace "VMName" with the name of your virtual machine.

.NOTES
    Tested on:
        - Windows 11 Pro
        - Hyper-V

    Requires:
        - Hyper-V PowerShell Module
        - Administrator privileges

.AUTHOR
    Luis Carlos Ramirez IT Support Specialist / Cloud & Systems Administrator
#>

# --------------------------------------------------------------------
# STEP 1
# Enable Nested Virtualization
# --------------------------------------------------------------------
#
# Exposes the host CPU virtualization extensions (Intel VT-x / AMD-V)
# to the specified virtual machine.
#
# Replace "VMName" with the name of your VM.
#
Set-VMProcessor -VMName "VMName" -ExposeVirtualizationExtensions $true


# --------------------------------------------------------------------
# STEP 2
# Verify from inside the Guest VM
# --------------------------------------------------------------------
#
# Run the following command INSIDE the virtual machine.
#
# Expected result:
#
# HyperVisorPresent               : True
# VirtualizationFirmwareEnabled   : True
#
# If both values are True, the guest operating system can access
# the virtualization extensions provided by the Hyper-V host.
#
Get-ComputerInfo |
    Select-Object HyperVisorPresent, VirtualizationFirmwareEnabled


# --------------------------------------------------------------------
# STEP 3
# Verify from the Hyper-V Host
# --------------------------------------------------------------------
#
# Run the following command on the HOST to verify that Nested
# Virtualization is enabled for the VM.
#
# Expected result:
#
# VMName                          : MyVirtualMachine
# ExposeVirtualizationExtensions  : True
#
Get-VMProcessor -VMName "VMName" |
    Select-Object VMName, ExposeVirtualizationExtensions
