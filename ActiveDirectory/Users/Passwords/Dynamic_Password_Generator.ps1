# Title: Dynamic Random Password Generator
# Description: Generates a random password based on the specified criteria

Clear-Host
Set-Location C:\Users\LuisRamirez\OneDrive - PG Solutions\Documents\PowerShell Study\powershelldemo

$PasswordLength = 20               # Max 8

$UpperCaseLetters = (65..90)        # Upper case letters A-Z  (SBNCKZRFWJPOIVEMQUGD)
$LowerCaseLetters = (97..122)       # Lower case letters a-z  (mnwtocbkjgxpraudhzly)
$NumbersZeroThroughNine = (48..57)  # Numbers 0-9             (7402163589)
$SpecialCharacters = (33..47)       # Special characters      ()+$,%./*#'-(&!")
$MoreSpecialCharacters = (58..64)   # More special characters (>@;?<:=)
$CurlyBrackets = (123..126)         # Curly brackets          (}{~|)

$a = -join ($UpperCaseLetters +
            $LowerCaseLetters +
            $NumbersZeroThroughNine +
            $SpecialCharacters +
            $MoreSpecialCharacters +
            $CurlyBrackets |
Get-Random -Count $PasswordLength | % {[char]$_})
$a
$a | clip