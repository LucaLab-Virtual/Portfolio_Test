Import-Module ActiveDirectory

$DomainDN = (Get-ADDomain).DistinguishedName
$DomainDNS = (Get-ADDomain).DNSRoot

Write-Host $DomainDNS

function Show-OUTree {
    param(
        [string]$SearchBase,
        [string]$Indent = ""
    )

    $ChildOUs = Get-ADOrganizationalUnit `
        -SearchBase $SearchBase `
        -SearchScope OneLevel `
        -Filter * |
        Sort-Object Name

    foreach ($OU in $ChildOUs)
    {
        Write-Host "$Indent├── OU=$($OU.Name)"
        Show-OUTree -SearchBase $OU.DistinguishedName -Indent "$Indent│   "
    }
}

Show-OUTree -SearchBase $DomainDN
