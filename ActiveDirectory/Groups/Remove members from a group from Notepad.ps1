# Remove members of a group from a .txt file

$groupName = "MarvelWorld-LSG"
$membersFile = "C:\Users\lramirez\Documents\EmailA.txt"

if (Get-ADGroup -Filter "Name -eq '$groupName'") {
    $members = Get-Content $membersFile
    
    foreach ($member in $members) {
        $trimmedMember = $member.Trim()
        Remove-ADGroupMember -Identity $groupName -Members $trimmedMember -Confirm:$false
    }
    
    Write-Host "Members removed from group successfully."
}
else {
    Write-Host "Group not found."
}