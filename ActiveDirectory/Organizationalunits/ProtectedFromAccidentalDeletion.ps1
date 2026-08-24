# =========================================================================================
# Enable or Disable for OUs the "Object" property - Protec object from accidental deletion
# =========================================================================================

$OUNames = @("Staff", "Admin", "_IT", "_HR")

foreach ($Name in $OUNames) {
    Get-ADOrganizationalUnit -Filter {Name -eq $Name} |
        Set-ADOrganizationalUnit -ProtectedFromAccidentalDeletion $false # $true
}
