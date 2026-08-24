# Enable Bitlocker on system volume

<#
Enables Bitlocker using AES 256 encryption. 
#>

Get-BitLockerVolume | Enable-BitLocker -EncryptionMethod Aes256 -UsedSpaceOnly -RecoveryKeyPath "E:\Recovery\" -RecoveryKeyProtector

# End