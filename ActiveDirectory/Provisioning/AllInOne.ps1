function Invoke-AdminMenu {
    [CmdletBinding(DefaultParameterSetName='Menu')]
    param()

    $previousCommandStatus = $null

    $Menu = @"
Choose the option you need:

1. Create a new ADUser
2. Create a new Distribution List or Group
3. Add a member to a Distribution List or Group
4. Remove a member from a Distribution List or Group
5. Add an Alias
6. Change an ADUser password
7. Modify an ADUser feature
8. Modify a Distribution List or Group
9. Delete and ADUser
10. Delete a Distribution List or Group
0. Exit
"@

    $choice = -1

    while ($choice -ne 0) {
        Clear-Host # Clear the console before displaying the menu
        if ($previousCommandStatus) {
            Write-Host $previousCommandStatus
            $previousCommandStatus = $null
        }
        Write-Host $Menu
        $choice = Read-Host "Enter your choice"

        switch ($choice) {
            1{
                # Create a new ADUser
                Write-Host "Creating a new ADUser..."
                Start-Sleep -Seconds 3
                $previousCommandStatus = "1. The new ADUser has been created successfully!"
                # Perform actions for option 1 here
                do {
                    Clear-Host
                    $validPattern = "^[A-Za-z0-9.\- ]+$"
                    do {
                    do {
                        $Fname = Read-Host "New User First Name"
                        if ($Fname -notmatch $validPattern) {
                            Write-Host "The first name '$Fname' contains invalid characters. Please enter a valid name: " -BackgroundColor Black -ForegroundColor Red
                        }
                    } while ($Fname -notmatch $validPattern)
                    
                    do {
                        $Lname = Read-Host "New User Last Name"
                        if ($Lname -notmatch $validPattern) {
                            Write-Host "The last name '$Lname' contains invalid characters. Please enter a valid name: " -BackgroundColor Black -ForegroundColor Red
                        }
                    } while ($Lname -notmatch $validPattern)
                        $RightNames = Read-Host "Is the new user's full name correct? (Yes/No)"
                    } while ($RightNames -eq "No")
                    $Displayname = "$Fname $Lname"
                    $Descrip = "Created by IT on $(Get-Date -Format "MM-dd-yyyy")"
                    $Office = Read-Host "What will be $($Displayname)'s OFFICE?"
                    $Title = Read-Host "What will be $($Displayname)'s JOB TITLE?"
                    $Department = Read-Host "What will be $($Displayname)'s DEPARTMENT?"
                    $Company = Read-Host "What COMPANY does $($Displayname) work for?"
                    $ManagerQuestion = Read-Host "Who's $($Displayname)'s MANAGER? (Format sample: lramirez)"
                    $ValidDomain = "^@[\w.-]+\.[a-zA-Z]{2,}$"
                    
                    do {
                        $Domain = Read-Host "Introduce a domain (Format sample: @lcrtest.com)"
                        if ($Domain -notmatch $ValidDomain) {
                            Write-Host "The domain must start with '@'. Please enter a valid domain (Format sample: @lcrtest.com): " -BackgroundColor Black -ForegroundColor Red
                        }
                    } While ($Domain -notmatch $ValidDomain)
                
                    # Select an existing OU
                    $OU_List = Get-ADOrganizationalUnit -Filter * | Select-Object -ExpandProperty Name
                    $Selected_OU = $OU_List | Out-GridView -Title "SELECT AN ORGANIZATIONAL UNIT" -OutputMode Single
                    $UserPath = (Get-ADOrganizationalUnit -Filter "Name -eq '$Selected_OU'").DistinguishedName
                
                    # Create a random password
                    $PasswordLength = 8 # Max 8
                    $UpperCaseLetters = (65..90) # Upper case letters A-Z (SBNCKZRFWJPOIVEMQUGD)
                    $LowerCaseLetters = (97..122) # Lower case letters a-z (mnwtocbkjgxpraudhzly)
                    $NumbersZeroThroughNine = (48..57) # Numbers 0-9 (7402163589)
                    $SpecialCharacters = (33..47) # Special characters ()+$,%./*#'-(&!")
                    $MoreSpecialCharacters = (58..64) # More special characters (>@;?<:=)
                    $CurlyBrackets = (123..126) # Curly brackets (}{~|)
                
                    $Password = -join ($UpperCaseLetters +
                                $LowerCaseLetters +
                                $NumbersZeroThroughNine +
                                $SpecialCharacters +
                                $MoreSpecialCharacters +
                                $CurlyBrackets |
                                Get-Random -Count $PasswordLength | ForEach-Object {[char]$_})
                
                    $SecureStringPassword = ConvertTo-SecureString -String $Password -AsPlainText -Force
                    $PasswordPlainText = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto([System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($SecureStringPassword))
                
                    function CheckUserNameAvailability {
                        param (
                            [string]$UserName
                        )
                        $UserExists = Get-ADUser -Filter "SamAccountName -eq '$UserName'"
                        if ($UserExists) {
                            return $true
                        } else {
                            return $false
                        }
                    }
                    # Generate the initial UserName
                    $UserName = "$($Fname.Substring(0,1))$($Lname)".ToLower()
                
                    # Check if the initial UserName already exists
                    if (CheckUserNameAvailability -UserName $UserName) {
                        do {
                            Write-Host "The username '$UserName' already exists." -BackgroundColor Black -ForegroundColor Red
                            $UserName = Read-Host " Please enter a different username" 
                        } While (CheckUserNameAvailability -UserName $UserName)
                    }
                
                    $EmailAddress = "$UserName$Domain"
                    $UserLogonName = "$UserName$Domain"
                
                    $Manager = Get-ADUser -Filter "SamAccountName -eq '$ManagerQuestion'" -Properties EmailAddress, DisplayName
                
                    if ($Manager) {
                        $ManagerName = $Manager.DisplayName
                        $ManagerEmail = $Manager.EmailAddress
                    } else {
                        $ManagerName = $ManagerQuestion
                        $ManagerEmail = "N/A"
                    }
                
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
                
                    Set-ADUser -Identity $Displayname -SamAccountName $UserName
                
                    do {
                        do {
                            $GroupQuestion = Read-Host "Would you like to add the user to a group? YES/NO "
                        } while ($GroupQuestion -notmatch "^(Yes|No)$")
                        
                        if ($GroupQuestion -eq "Yes") {
                            $GroupName = Read-Host "Introduce Group Name: "
                            
                            if (-not [string]::IsNullOrWhiteSpace($GroupName)) {
                                $Group = Get-ADGroup -Filter "Name -eq '$GroupName'"
                                if ($Group) {
                                    Add-ADGroupMember -Identity $Group -Members $UserName
                                    Write-Host "User added successfully to '$GroupName'"
                                } else {
                                    Write-host "Group '$GroupName' not found" -BackgroundColor Black -ForegroundColor Red
                                }
                            } else {
                                Write-Host "No group name provided" -BackgroundColor Black -ForegroundColor Red
                            }
                        }
                    } While ($GroupQuestion -eq "Yes")
                    
                    $MemberOf = Get-ADPrincipalGroupMembership -Identity $UserName | Select-Object Name
                
                    Write-Host "We have created a new user. The details are below.
                
                Display name: $($Displayname)
                Email address: $($EmailAddress)
                Password:  $($PasswordPlainText)  (We recommend changing it as soon as possible).
                License:  (License location - ).
                Title: $($Title)
                Department: $($Department)
                Manager: $ManagerName "($ManagerEmail)"
                Office: $($Office)
                Member of:" -BackgroundColor Black -ForegroundColor Cyan
                
                    $MemberOf | ForEach-Object { 
                        Write-Host "$($_.Name)" -BackgroundColor Black -ForegroundColor Cyan
                    }
                
                    Write-Host "`nRegards!" -BackgroundColor Black -ForegroundColor Cyan
                
                    $Continue = Read-Host "Do you want to create another user? (Yes/No)"
                } while ($Continue -eq "Yes")
                break
            }
            2{
                # Create a new Distribution List or Group
                Write-Host "Creating a new Distribution List or Group..."
                Start-Sleep -Seconds 3
                $previousCommandStatus = "2. The new Distribution List or Group has been created successfully!"
                # Perform actions for option 2 here
                break
            }
            3{
                # Add a member to a Distribution List or Group
                Write-Host "Adding a member to a Distribution List or Group..."
                Start-Sleep -Seconds 3
                $previousCommandStatus = "3. Member has been added to the Distribution List or Group successfully!"
                # Perform actions for option 3 here
                break
            }
            4{
                # Remove a member from a Distribution List or Group
                Write-Host "Removing a member from a Distribution List or Group..."
                Start-Sleep -Seconds 3
                $previousCommandStatus = "4. Member has been removed from the Distribution List or Group successfully!"
                # Perfom actions for option 4 here
                break
            }
            5{
                # Add an Alias
                Write-Host "Adding an Alias..."
                Start-Sleep -Seconds 3
                $previousCommandStatus = "5. Alias has been added successfully!"
                # Perform actions for option 5 here
                break
            }
            6{
                # Change an ADUser password
                Write-Host "Changing and ADUser password..."
                Start-Sleep -Seconds 3
                $previousCommandStatus = "6. ADUser password has been changed successfully!"
                # Perfom actions for option 6 here
                break
            }
            7{
                # Modify an ADUser feature
                Write-Host "Modifying an ADUser feature..."
                Start-Sleep -Seconds 3
                $previousCommandStatus = "7. ADUser feature has been modified successfully!"
                # Perfom actions for option 7 here
                break
            }
            8{
                # Modify a Distribution List or Group
                Write-Host "Modifying a Distribution List or Group..."
                Start-Sleep -Seconds 3
                $previousCommandStatus = "8. Distribution List or Group has been modified successfully!"
                # Perfom actions for option 8 here
                break
            }
            9{
                # Delete an ADUser
                Write-Host "Deleting an ADUser..."
                Start-Sleep -Seconds 3
                $previousCommandStatus = "9. ADUser has been deleted successfully!"
                # Perform actions for option 9 here
                break
            }
            10{
                # Delete a Distribution List or Group
                Write-Host "Deleling a Distribution List or Group..."
                Start-Sleep -Seconds 3
                $previousCommandStatus = "10. Distribution List or Group has been deleted successfully!"
                # Perform action for option 10 here
                break
            }
            0{
                # Exit
                Write-Host "Exiting..."
                break
            }
        }
    }
}

# Hide the cmdlet while running
$function:Visibility = 'Private'
Invoke-AdminMenu