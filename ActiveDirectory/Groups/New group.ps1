# Create a new group

$Name = Read-Host "Enter the name for the new group"
$Security = "Security"
$Distribution = "Distribution"
$Category = Read-Host "Enter the group category (Security or Distribution)"
while (($Category -ne $Security) -and ($Category -ne $Distribution)) {
    Write-Output "The Category is invalid."
    $Category = Read-Host "Enter the group category (Security or Distribution)"
}

$Global = "Global"
$DomainLocal = "Domain local"
$Universal = "Universal"
$Scope = Read-Host "Enter the group scope (Global, Domain local or Universal)"
while (($Scope -ne $Global) -and ($Scope -ne $DomainLocal) -and ($Scope -ne $Universal)) {
    Write-Output "The Scope is invalid."
    $Scope = Read-Host "What's the scope?"
}


$OU_List = Get-ADOrganizationalUnit -Filter * | Select-Object -ExpandProperty Name
$Selected_OU = $OU_List | Out-GridView -Title "Select OU to store the Group" -OutputMode Single

if ($Category -eq "Distribution"){
    New-ADGroup -Name $Name `
            -GroupCategory $Category `
            -GroupScope $Scope `
            -Path "OU=$Selected_OU,DC=lcrtest,DC=com"
            
            $listName = $Name
            $emailaddress = $Name + '@lcrtest.com'
    Write-Host "The '$Name' Distribution list has been created successfully"
}
else{New-ADGroup -Name $Name `
            -GroupCategory $Category `
            -GroupScope $Scope `
            -Path "OU=$Selected_OU,DC=lcrtest,DC=com"
    Write-Host "The '$Name' Security Group has been created successfully"
}
