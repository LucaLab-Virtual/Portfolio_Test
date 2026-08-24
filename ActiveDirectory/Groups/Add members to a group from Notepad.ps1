# Add members to a group from a .txt file using the UserLogonName
$groupName = "MarvelWorld-LSG"
$membersFile = "C:\Users\lramirez\Documents\EmailA.txt"

if (Get-ADGroup -Filter "Name -eq '$groupName'") {
    
    $members = Get-Content $membersFile
    
    
    foreach ($member in $members) {
        $trimedMember = $members.Trim()
        Add-ADGroupMember -Identity $groupName -Members $member
    }
    Write-Host "Members added to group successfully."
}
else {
    Write-Host "Group not found."
}