$UserList = Get-Content C:\Users\lramirez\Documents\EmailA.txt
$GroupName = Read-Host "Introduce the Group Name: "
$AddedCount = 0
$ExistingCount = 0

foreach ($User in $UserList) {
    $UserObj = Get-ADUser -Filter {EmailAddress -eq $User} -Properties MemberOf
    $GroupObj = Get-ADGroup -Identity $GroupName -Properties Members

    if ($UserObj.MemberOf -notcontains $GroupObj.DistinguishedName) 
    {Add-ADGroupMember -Identity $GroupName -Members (Get-ADUser -Filter {EmailAddress -eq $User})
    Write-Host "User '$User' has been added as a member of '$GroupName'" -BackgroundColor Black -ForegroundColor Cyan 
    $AddedCount++}
    if ($UserObj.MemberOf -contains $GroupObj.DistinguishedName)
    {Write-Host "User '$User' is already a member of '$GroupName'" -BackgroundColor Black -ForegroundColor Red 
    $ExistingCount++}

}
Write-Host "SUCCESSFULLY!!" -BackgroundColor Black -ForegroundColor Yellow
Write-Host "Members added: $AddedCount - Existing members: $ExistingCount" -BackgroundColor Black -ForegroundColor Yellow