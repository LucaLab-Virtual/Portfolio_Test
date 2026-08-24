#The script will create a new OU and users taken from a .txt file named Users.txt
$Password_For_User = "Wow456@1" 
$User_List = Get-Content .\Users.txt
$ou = Read-Host "Enter the OU name (sample: _Admins)"

$password = ConvertTo-SecureString $Password_For_User -AsPlainText -Force
New-ADOrganizationalUnit -Name $ou -ProtectedFromAccidentalDeletion $false

Foreach ($n in $User_List) {
    $first = $n.split(" ")[0]
    $last = $n.split(" ")[1]
    $username = "$($first.Substring(0,1))$($last)".ToLower()
    $displayname = "$($first + " " + $last)"
    $description = "Created by IT on $(get-date -Format "MM-dd-yyyy")"
    $office = Read-Host "Enter the office for $($displayname)"
    $emailaddress = $username + "@lcrtest.com"
    $userlogonname = $username + "@lcrtest.com"
    $jobtitle = Read-Host "Enter the job title for $($displayname)"
    $dep = "IT"
    $comp = "LCRTEST"
    $manager = "CN=Lucarez Ramirez,OU=_Admins,DC=lcrtest,DC=com"
    Write-Host "Creating User: $($displayname)" -BackgroundColor Black -ForegroundColor Cyan

    New-ADUser -AccountPassword $password `
               -GivenName $first `
               -Surname $last `
               -DisplayName $displayname `
               -Description $description `
               -Office $office `
               -EmailAddress $emailaddress `
               -UserPrincipalName $userlogonname `
               -Title $jobtitle `
               -Department: $dep `
               -Company: $comp `
               -Manager $manager `
               -Name $username `
               -EmployeeID $username `
               -PasswordNeverExpires $true `
               -Path "ou=$ou,$(([ADSI]`"").distinguishedName)" `
               -Enabled $true
}