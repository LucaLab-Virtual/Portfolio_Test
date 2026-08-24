# Define a regular expression pattern to validate user input for first and last names.
$validPattern = "^[A-Za-z0-9.\- ]+$"

# Prompt the user for their first name, validate it, and keep prompting until it's valid.
$FnameConfirmed = $false
do {
    $Fname = Read-Host "Type first name?"
    $FnameConfirmation = Read-Host "Is the first name you introduced $Fname, correct? (Yes/No)"
    if ($FnameConfirmation -eq "Yes") {
        $FnameConfirmed = $true
    }
    elseif ($Fname -notmatch $validPattern) {
        Write-Output "The first name '$Fname' contains invalid characters. Please enter a valid name."
    }
} while (-not ($Fname -match $validPattern) -or -not $FnameConfirmed)

# Prompt the user for their last name, validate it, and keep prompting until it's valid.
$LnameConfirmed = $false
do {
    $Lname = Read-Host "Type Last name"
    $LnameConfirmation = Read-Host "Is the last name you introduced $Lname, correct? (Yes/No)"
    if ($LnameConfirmation -eq "Yes") {
        $LnameConfirmed = $true
    }
    if ($Lname -notmatch $validPattern) {
        Write-Output "The last name '$Lname' contains invalid characters. Please enter a valid name."
    }
} while (-not ($Lname -match $validPattern) -or -not $LnameConfirmed)

# Combine the first name and last name to create a display name for the user.
$Displayname = "$Fname $Lname"

# Prompt the user for the ticket number, and confirm it with "Yes" or "No"
$TicketConfirmed = $false
do {
    $Ticket = Read-Host "Type the ticket number"
    $TicketConfirmation = Read-Host "The ticket number you introduced is $Ticket. Is this the right ticket? (Yes/No)"
    if ($TicketConfirmation -eq "Yes") {
        $TicketConfirmed = $true
    }
} While (-not $TicketConfirmed)

# Set various user attributes.
$Descrip = "OrbisOne $(Get-Date -Format "MM-dd-yyyy") T#:$Ticket"

# Prompt the user for the office name, and confirm it with "Yes" or "No."
$OfficeConfirmed = $false
do {
    $Office = Read-Host "Type the office name (Sample: Pompano Beach/Coconut Creek)"
    $OfficeConfirmation = Read-Host "The office name you introduced is '$Office'. Is this the right office? (Yes/No)"
    if ($OfficeConfirmation -eq "Yes") {
        $OfficeConfirmed = $true
    }
} while (-not $OfficeConfirmed)

# Prompt the user for the title name, and confirm with "Yes" or "No"
$TitleConfirmed = $false
do {
    $Title = Read-Host "Type the title name"
    $TitleConfirmation = Read-Host "The title name you introduced is '$Title'. Is this the right title? (Yes/No)"
    if ($TitleConfirmation -eq "Yes") {
        $TitleConfirmed = $true
    }
} While (-not $TitleConfirmed)

# Prompt the user for the company name, and confirm with "Yes" or "No"
$CompanyConfirmed = $false
do {
    $Company =  Read-Host "Type the company name (Sample: Mazda/Chevrolet)"
    $CompanyConfirmation = Read-Host "The company name you introduced is '$Company'. Is this the right company? (Yes/No)"
    if ($CompanyConfirmation -eq "Yes") {
        $CompanyConfirmed = $true
    }
} While (-not $CompanyConfirmed)

# Prompt the user for the department name, and confirm with "Yes" or "No"
$DepartmentConfirmed = $false
do {
    $Department = Read-Host "Type the department name"
    $DepartmentConfirmation = Read-Host "The department name you introduced is $Department. Is this the right department? (Yes/No)"
    if ($DepartmentConfirmation -eq "Yes") {
        $DepartmentConfirmed = $true
    }
} While (-not $DepartmentConfirmed)

# Prompt the user for the email address domain. and confirm with "Yes" or "No"
$DomainConfirmed = $false
do {
    $Domain = Read-Host "Type the domain (Sample: @loubachrodt.com)"
    $DomainConfirmation = Read-Host "The domain you introduced is $Domain. Is this the right domain? (Yes/No)"
    if ($DomainConfirmation -eq "Yes") {
        $DomainConfirmed = $true
    }
} While (-not $DomainConfirmed)

# Prompt the user for the manager's logon name.
$ManagerQuestion = Read-Host "Who's the Manager (Introduce logon name (Sample: lramirez))? "

# Retrieve a list of Organizational Units (OUs) and let the user select one.
$OU_List = Get-ADOrganizationalUnit -Filter * | Select-Object -ExpandProperty Name
$Selected_OU = $OU_List | Out-GridView -Title "Select an Organizational Unit" -OutputMode Single
$UserPath = (Get-ADOrganizationalUnit -Filter "Name -eq '$Selected_OU'").DistinguishedName

# Generate a random password for the user.
$PasswordLength = 10
$UpperCaseLetters = (65..90)
$LowerCaseLetters = (97..122)
$NumbersZeroThroughNine = (48..57)
$SpecialCharacters = (33..47)
$MoreSpecialCharacters = (58..64)
$CurlyBrackets = (123..126)
$Password = -join (
    $UpperCaseLetters + $LowerCaseLetters + $NumbersZeroThroughNine + $SpecialCharacters + $MoreSpecialCharacters + $CurlyBrackets |
    Get-Random -Count $PasswordLength | ForEach-Object {[char]$_}
)

# Convert the password to a secure string and retrieve it as plain text.
$SecureStringPassword = ConvertTo-SecureString -String $Password -AsPlainText -Force
$PasswordPlainText = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto([System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($SecureStringPassword))

# Generate a username, email address, and user logon name.
$UserName = "$($Fname.Substring(0,1))$($Lname)".ToLower()
$EmailAddress = $UserName + $Domain
$UserLogonName = $UserName + $Domain

# Retrieve the manager's information if available.
$Manager = Get-ADUser -Filter "SamAccountName -eq '$ManagerQuestion'" -Properties EmailAddress, DisplayName
if ($Manager) {
    $ManagerName = $Manager.DisplayName
    $ManagerEmail = $Manager.EmailAddress
} else {
    $ManagerName = $ManagerQuestion
    $ManagerEmail = "N/A"
}

# Create a new Active Directory user with the specified attributes.
New-ADUser -Name $Displayname `
           -GivenName $Fname `
           -Surname $Lname `
           -DisplayName $Displayname `
           -Description $Descrip `
           -Office $Office `
           -Title $Title `
           -Department $Department `
           -Company $Company `
           -Manager $Manager `
           -EmailAddress $EmailAddress `
           -UserPrincipalName $UserLogonName `
           -AccountPassword $SecureStringPassword `
           -Path $UserPath `
           -Enabled $true

# Set the user's SamAccountName.
Set-ADUser -Identity $Displayname -SamAccountName $UserName

# Prompt the user to add the user to one or more groups.
do {
    do {
        $GroupQuestion = Read-Host "Would you like to add the user to a group? YES/NO "
    } while ($GroupQuestion -notmatch "^(Yes|No)$")

    if ($GroupQuestion -eq "Yes") {
        $GroupName = Read-Host "Introduce Group Name: "
    }

    if (Get-ADGroup -Filter "Name -eq '$GroupName'") {
        Add-ADGroupMember -Identity $GroupName -Members $UserName
        Write-Host "User added successfully to '$GroupName'"
    } else {
        Write-host "Group not found"
    }
} While ($GroupQuestion -eq "Yes")

# Retrieve the groups the user is a member of.
$MemberOf = Get-ADPrincipalGroupMembership -Identity $UserName | Select-Object Name

# Display user details to the console.
Write-Host "We have created a new user. The details are below.

Display name: $($Displayname)
Email address: $($EmailAddress)
Password:  $($PasswordPlainText)  (We recommend changing it as soon as possible).
License:  (License location - ).
Title: $($Title)
Department: $($Department)
Manager: $ManagerName ($ManagerEmail)
Office: $($Office)
MemberOf: $($MemberOf)

Regards!" -BackgroundColor Black -ForegroundColor Cyan