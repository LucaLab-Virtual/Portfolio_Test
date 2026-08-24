# Read usernames from text file
$users = Get-Content -Path "C:\Users\Administrator\Documents\RemoveUsers2.txt"

foreach ($termuser in $users) {
    # Script halt 10 seconds to enter USERNAME
    Start-Sleep -Seconds 10

    # Exports Group Memberships to txt
    $target = "C:\Users\Administrator\Documents\DisabledUsersReport\" + $termuser + ".txt"
    Get-ADPrincipalGroupMembership $termuser | Select-Object -ExpandProperty Name | Export-Csv -Path $target
    Write-Host "* Group Memberships archived to" $target

    # Move to "Disabled Users" OU
    Get-ADUser $termuser | Move-ADObject -TargetPath "OU=_DisabledUsers,DC=lcrtest,DC=com"
    Write-Host "* " $termuser "moved to Disabled Users"

    # Change Description to "Terminated YYYY.MM.DD - CURRENT USER"
    $terminatedby = $env:username
    $termDate = Get-Date -UFormat "%Y.%m.%d"
    $termUserDesc = "Disabled by " + $terminatedby + " - " + $termDate
    Set-ADUser $termuser -Description $termUserDesc
    Write-Host "* " $termuser "description set to" $termUserDesc

    # Removes from all distribution groups
    $dlists = (Get-ADUser $termuser -Properties memberof).memberof
    foreach ($dlist in $dlists) {
        Remove-ADGroupMember -Identity $dlist -Members $termuser -Confirm:$False
    }
    Write-Host "* Removed from all distribution and security groups"

    # Disable user
    Disable-ADAccount -Identity $termuser
    Get-ADUser -Identity $termuser -Properties * |
        Set-ADUser -ChangePasswordAtLogon:$true -Clear telephoneNumber, homeDrive, homeDirectory, manager, mail, company, scriptpath
}
Write-Host "All the users were disabled Successfully!" -BackgroundColor Black -ForegroundColor Cyan

# End
