<#
.SYNOPSIS
    Performs bulk password resets for Active Directory users with advanced features.
.DESCRIPTION
    This script provides enterprise-grade bulk password reset functionality including:
    - CSV import with multiple user identification methods
    - Configurable password generation (random, dictionary, or custom)
    - Password policy validation
    - Force password change at next logon
    - Email notification to users
    - Secure password handling and encryption
    - Comprehensive logging and error handling
    - Dry-run mode for testing
.PARAMETER InputFile
    Path to CSV file containing user information. Required if -Users not specified.
.PARAMETER Users
    Array of user identities (SamAccountName, UPN, or DN). Can be piped.
.PARAMETER PasswordGeneration
    Method for password generation: 'Random', 'Dictionary', 'Custom'. Default is 'Random'.
.PARAMETER PasswordLength
    Length of generated passwords (12-20). Default is 14.
.PARAMETER CustomPassword
    Specific password to use for all users (not recommended for bulk operations).
.PARAMETER ForceChangeAtNextLogon
    Force users to change password at next logon. Default is $true.
.PARAMETER NotifyUsers
    Send email notifications to users with their new passwords.
.PARAMETER SmtpServer
    SMTP server for email notifications.
.PARAMETER FromEmail
    From email address for notifications.
.PARAMETER Subject
    Email subject line for notifications.
.PARAMETER BodyTemplate
    Path to HTML or text file for custom email body template.
.PARAMETER ExportPasswords
    Export passwords to encrypted CSV file (requires -SecureExport).
.PARAMETER SecureExport
    Export passwords with encryption using provided key or certificate.
.PARAMETER DryRun
    Simulate the operation without making changes.
.PARAMETER LogPath
    Path for log file. Defaults to script directory.
.PARAMETER UnlockUsers
    Unlock users after password reset. Default is $true.
.PARAMETER DisableUser
    Disable user account after password reset (emergency use).
.PARAMETER DomainController
    Specify domain controller for AD operations.
.EXAMPLE
    .\Bulk_Reset_Passwords.ps1 -InputFile "users.csv" -PasswordGeneration Random -NotifyUsers
.EXAMPLE
    .\Bulk_Reset_Passwords.ps1 -Users @("user1","user2","user3") -PasswordLength 16 -DryRun
.EXAMPLE
    .\Bulk_Reset_Passwords.ps1 -InputFile "reset.csv" -ExportPasswords -SecureExport -SmtpServer "smtp.domain.com"
.NOTES
    Requires ActiveDirectory module. Run with elevated privileges.
    For secure password export, use Windows DPAPI or certificate encryption.
#>

<#
CSV Input File Format:
csv

SamAccountName,DisplayName,Email,Department
jsmith,John Smith,jsmith@domain.com,IT
kbrown,Karen Brown,kbrown@domain.com,HR
mwilson,Mark Wilson,mwilson@domain.com,Sales

Alternative column names supported:

    SamAccountName, UPN, Username, Identity, UserPrincipalName, DN, DistinguishedName

Usage Examples:
powershell

# Basic CSV reset with random passwords
.\Bulk_Reset_Passwords.ps1 -InputFile "users.csv" -PasswordGeneration Random

# Dictionary passwords with notifications
.\Bulk_Reset_Passwords.ps1 -InputFile "users.csv" -PasswordGeneration Dictionary -NotifyUsers -SmtpServer "smtp.domain.com" -FromEmail "admin@domain.com"

# Test mode with specific users
.\Bulk_Reset_Passwords.ps1 -Users @("jsmith","kbrown","mwilson") -DryRun

# Secure export with forced password change
.\Bulk_Reset_Passwords.ps1 -InputFile "users.csv" -ExportPasswords -SecureExport -ForceChangeAtNextLogon

# Custom email template and disabled accounts
.\Bulk_Reset_Passwords.ps1 -InputFile "users.csv" -NotifyUsers -BodyTemplate "email_template.html" -DisableUser -UnlockUsers

Email Template Example:

Create a custom HTML template with these placeholders:

    {UserName} - SamAccountName

    {UserDisplayName} - Display name

    {NewPassword} - The generated password

    {Domain} - Domain name

    {Date} - Current date/time

Log File Structure:
text

[Bulk Password Reset Log]
Start: 2026-01-15 14:30:22
User: Administrator
Computer: DC01
====================================
[2026-01-15 14:30:25] [INFO] Loaded 15 users from CSV
[2026-01-15 14:30:26] [INFO] Processing user 1 of 15 : jsmith
[2026-01-15 14:30:27] [INFO] Generated random password for jsmith
[2026-01-15 14:30:28] [INFO] Password reset successful for jsmith
...
====================================
Total Users Processed: 15
Successful Resets: 14
Failed Resets: 1
====================================
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory=$false, ParameterSetName="File")]
    [string]$InputFile,
    
    [Parameter(Mandatory=$false, ParameterSetName="Array")]
    [string[]]$Users,
    
    [Parameter(Mandatory=$false)]
    [ValidateSet('Random', 'Dictionary', 'Custom')]
    [string]$PasswordGeneration = 'Random',
    
    [Parameter(Mandatory=$false)]
    [ValidateRange(12, 20)]
    [int]$PasswordLength = 14,
    
    [Parameter(Mandatory=$false)]
    [string]$CustomPassword,
    
    [Parameter(Mandatory=$false)]
    [bool]$ForceChangeAtNextLogon = $true,
    
    [Parameter(Mandatory=$false)]
    [switch]$NotifyUsers,
    
    [Parameter(Mandatory=$false)]
    [string]$SmtpServer,
    
    [Parameter(Mandatory=$false)]
    [string]$FromEmail,
    
    [Parameter(Mandatory=$false)]
    [string]$Subject = "Your Password Has Been Reset",
    
    [Parameter(Mandatory=$false)]
    [string]$BodyTemplate,
    
    [Parameter(Mandatory=$false)]
    [switch]$ExportPasswords,
    
    [Parameter(Mandatory=$false)]
    [switch]$SecureExport,
    
    [Parameter(Mandatory=$false)]
    [switch]$DryRun,
    
    [Parameter(Mandatory=$false)]
    [string]$LogPath,
    
    [Parameter(Mandatory=$false)]
    [switch]$UnlockUsers = $true,
    
    [Parameter(Mandatory=$false)]
    [switch]$DisableUser,
    
    [Parameter(Mandatory=$false)]
    [string]$DomainController
)

#region Global Variables and Initialization
$Script:LogFile = $null
$Script:Results = [System.Collections.ArrayList]::new()
$Script:ErrorCount = 0
$Script:SuccessCount = 0
$Script:Timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$Script:DefaultPassword = $null

# Set default log path if not specified
if (-not $LogPath) {
    $LogPath = Join-Path $PSScriptRoot "Logs"
}
#endregion

#region Helper Functions
function Write-Log {
    param(
        [string]$Message,
        [string]$Level = "INFO",
        [string]$Color = "White"
    )
    
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logMessage = "[$timestamp] [$Level] $Message"
    
    # Write to console with color if not in file mode
    if ($Color) {
        Write-Host $logMessage -ForegroundColor $Color
    }
    else {
        Write-Host $logMessage
    }
    
    # Write to log file
    if ($Script:LogFile) {
        Add-Content -Path $Script:LogFile -Value $logMessage -Encoding UTF8
    }
}

function Initialize-Logging {
    try {
        if (-not (Test-Path $LogPath)) {
            New-Item -Path $LogPath -ItemType Directory -Force | Out-Null
        }
        
        $Script:LogFile = Join-Path $LogPath "BulkPasswordReset_$Script:Timestamp.log"
        Write-Log "="*70 "INFO" "Yellow"
        Write-Log "BULK PASSWORD RESET OPERATION STARTED" "INFO" "Yellow"
        Write-Log "="*70 "INFO" "Yellow"
        Write-Log "Start Time: $(Get-Date)" "INFO" "Cyan"
        Write-Log "User: $env:USERNAME" "INFO" "Cyan"
        Write-Log "Computer: $env:COMPUTERNAME" "INFO" "Cyan"
        Write-Log "Dry Run Mode: $DryRun" "INFO" "Cyan"
        Write-Log "="*70 "INFO" "Yellow"
    }
    catch {
        Write-Warning "Could not initialize logging: $_"
    }
}

function Import-ADModule {
    try {
        if (-not (Get-Module -Name ActiveDirectory -ListAvailable)) {
            Write-Error "ActiveDirectory module is not installed."
            return $false
        }
        Import-Module ActiveDirectory -Force -ErrorAction Stop
        Write-Log "ActiveDirectory module loaded successfully" "INFO" "Green"
        return $true
    }
    catch {
        Write-Log "Failed to load ActiveDirectory module: $_" "ERROR" "Red"
        return $false
    }
}

function Import-CSVUsers {
    param([string]$FilePath)
    
    try {
        if (-not (Test-Path $FilePath)) {
            Write-Log "CSV file not found: $FilePath" "ERROR" "Red"
            return $null
        }
        
        $csvData = Import-Csv -Path $FilePath -ErrorAction Stop
        Write-Log "Loaded $($csvData.Count) users from CSV" "INFO" "Green"
        return $csvData
    }
    catch {
        Write-Log "Error importing CSV: $_" "ERROR" "Red"
        return $null
    }
}

function Generate-RandomPassword {
    param([int]$Length = 14)
    
    # Password requirements: at least one uppercase, lowercase, digit, special character
    $charSets = @(
        'ABCDEFGHIJKLMNOPQRSTUVWXYZ'  # Uppercase
        'abcdefghijklmnopqrstuvwxyz'  # Lowercase
        '0123456789'                  # Digits
        '!@#$%^&*()-_=+[]{};:,.<>?'  # Special characters
    )
    
    # Ensure at least one from each set
    $password = @(
        $charSets[0][(Get-Random -Maximum $charSets[0].Length)]
        $charSets[1][(Get-Random -Maximum $charSets[1].Length)]
        $charSets[2][(Get-Random -Maximum $charSets[2].Length)]
        $charSets[3][(Get-Random -Maximum $charSets[3].Length)]
    )
    
    # Fill remaining length with random characters from all sets
    $allChars = $charSets -join ''
    for ($i = $password.Count; $i -lt $Length; $i++) {
        $password += $allChars[(Get-Random -Maximum $allChars.Length)]
    }
    
    # Shuffle the password
    $shuffled = $password | Sort-Object { Get-Random }
    return -join $shuffled
}

function Generate-DictionaryPassword {
    param([int]$Length = 14)
    
    # Common dictionary words with substitutions
    $dictionary = @(
        'Summer', 'Winter', 'Spring', 'Autumn', 'Fire', 'Storm', 'Thunder',
        'Lightning', 'Rainbow', 'Sunset', 'Dragon', 'Phoenix', 'Eagle', 'Tiger',
        'Lion', 'Wolf', 'Star', 'Moon', 'Galaxy', 'Cosmos', 'Serenity', 'Harmony',
        'Freedom', 'Justice', 'Knight', 'Warrior', 'Guardian', 'Phoenix'
    )
    
    # Add numbers and special characters
    $numbers = '0123456789'
    $specials = '!@#$%^&*'
    
    $word = $dictionary | Get-Random
    $number = -join (1..3 | ForEach-Object { $numbers[(Get-Random -Maximum $numbers.Length)] })
    $special = $specials[(Get-Random -Maximum $specials.Length)]
    
    $password = "$word$number$special"
    
    # Ensure minimum length
    while ($password.Length -lt $Length) {
        $password += $numbers[(Get-Random -Maximum $numbers.Length)]
    }
    
    return $password
}

function Test-PasswordComplexity {
    param([string]$Password)
    
    $hasUpper = $Password -cmatch '[A-Z]'
    $hasLower = $Password -cmatch '[a-z]'
    $hasDigit = $Password -match '\d'
    $hasSpecial = $Password -match '[!@#$%^&*()\-_=+\[\]{};:,.<>?]'
    $isLongEnough = $Password.Length -ge 12
    
    return ($hasUpper -and $hasLower -and $hasDigit -and $hasSpecial -and $isLongEnough)
}

function Get-UserObject {
    param(
        [string]$Identity,
        [string]$DC
    )
    
    try {
        $params = @{
            Identity = $Identity
            Properties = @(
                'SamAccountName',
                'DisplayName',
                'UserPrincipalName',
                'GivenName',
                'Surname',
                'EmailAddress',
                'Title',
                'Department',
                'Enabled',
                'LockedOut',
                'PasswordNeverExpires',
                'PasswordLastSet',
                'DistinguishedName'
            )
            ErrorAction = 'Stop'
        }
        if ($DC) { $params.Server = $DC }
        
        return Get-ADUser @params
    }
    catch {
        Write-Log "Error retrieving user '$Identity': $_" "ERROR" "Red"
        return $null
    }
}

function Send-PasswordNotification {
    param(
        [string]$UserEmail,
        [string]$UserName,
        [string]$UserDisplayName,
        [string]$NewPassword,
        [string]$Domain
    )
    
    if (-not $NotifyUsers) { return }
    if (-not $SmtpServer -or -not $FromEmail) {
        Write-Log "SMTP configuration missing. Skipping notification for $UserName" "WARNING" "Yellow"
        return
    }
    
    try {
        # Load custom body template if provided
        $body = $null
        if ($BodyTemplate -and (Test-Path $BodyTemplate)) {
            $body = Get-Content -Path $BodyTemplate -Raw
            $body = $body -replace '{UserName}', $UserName
            $body = $body -replace '{UserDisplayName}', $UserDisplayName
            $body = $body -replace '{NewPassword}', $NewPassword
            $body = $body -replace '{Domain}', $Domain
            $body = $body -replace '{Date}', (Get-Date -Format "yyyy-MM-dd HH:mm:ss")
        }
        else {
            # Default HTML body
            $body = @"
<!DOCTYPE html>
<html>
<head>
    <style>
        body { font-family: Arial, sans-serif; }
        .container { max-width: 600px; margin: 0 auto; padding: 20px; border: 1px solid #ddd; border-radius: 5px; }
        .header { background: #2c3e50; color: white; padding: 10px; border-radius: 3px; }
        .content { padding: 20px; }
        .password { background: #f8f9fa; padding: 10px; border-left: 4px solid #3498db; font-family: monospace; font-size: 16px; }
        .footer { font-size: 12px; color: #7f8c8d; margin-top: 20px; border-top: 1px solid #ddd; padding-top: 10px; }
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <h2>Password Reset Notification</h2>
        </div>
        <div class="content">
            <p>Hello <strong>$UserDisplayName</strong>,</p>
            <p>Your password for the domain <strong>$Domain</strong> has been reset.</p>
            <p><strong>New Password:</strong></p>
            <div class="password">$NewPassword</div>
            <p><strong>Important:</strong> You will be required to change this password at your next logon.</p>
            <p>If you did not request this password reset, please contact your system administrator immediately.</p>
            <p>This is an automated notification. Please do not reply to this email.</p>
        </div>
        <div class="footer">
            <p>Generated: $(Get-Date -Format "yyyy-MM-dd HH:mm:ss")</p>
            <p>Security Notice: Please keep your password confidential.</p>
        </div>
    </div>
</body>
</html>
"@
        }
        
        $mailParams = @{
            To = $UserEmail
            From = $FromEmail
            Subject = $Subject
            Body = $body
            BodyAsHtml = $true
            SmtpServer = $SmtpServer
            ErrorAction = 'Stop'
        }
        
        Send-MailMessage @mailParams
        Write-Log "Notification sent to $UserEmail" "INFO" "Green"
        return $true
    }
    catch {
        Write-Log "Failed to send notification to $UserEmail: $_" "ERROR" "Red"
        return $false
    }
}

function Reset-UserPassword {
    param(
        [Microsoft.ActiveDirectory.Management.ADUser]$User,
        [string]$NewPassword,
        [hashtable]$Options
    )
    
    $result = [PSCustomObject]@{
        SamAccountName = $User.SamAccountName
        DisplayName = $User.DisplayName
        Email = $User.EmailAddress
        Status = "Failed"
        Error = $null
        NewPassword = if ($Options.ExportPasswords) { $NewPassword } else { "***" }
    }
    
    try {
        if ($DryRun) {
            Write-Log "[DRY RUN] Would reset password for $($User.SamAccountName)" "INFO" "Yellow"
            $result.Status = "DryRun"
            return $result
        }
        
        $params = @{
            Identity = $User.DistinguishedName
            NewPassword = (ConvertTo-SecureString -String $NewPassword -AsPlainText -Force)
            ErrorAction = 'Stop'
        }
        
        if ($DomainController) { $params.Server = $DomainController }
        
        # Set the password
        Set-ADAccountPassword @params
        Write-Log "Password reset successful for $($User.SamAccountName)" "INFO" "Green"
        
        # Force password change at next logon
        if ($Options.ForceChangeAtNextLogon) {
            $changeParams = @{
                Identity = $User.DistinguishedName
                ErrorAction = 'Stop'
            }
            if ($DomainController) { $changeParams.Server = $DomainController }
            Set-ADUser @changeParams -ChangePasswordAtLogon $true
            Write-Log "Password change required at next logon for $($User.SamAccountName)" "INFO" "Green"
        }
        
        # Unlock user if locked out
        if ($Options.UnlockUsers -and $User.LockedOut) {
            $unlockParams = @{
                Identity = $User.DistinguishedName
                ErrorAction = 'Stop'
            }
            if ($DomainController) { $unlockParams.Server = $DomainController }
            Unlock-ADAccount @unlockParams
            Write-Log "Account unlocked for $($User.SamAccountName)" "INFO" "Green"
        }
        
        # Disable user if requested
        if ($Options.DisableUser) {
            $disableParams = @{
                Identity = $User.DistinguishedName
                ErrorAction = 'Stop'
            }
            if ($DomainController) { $disableParams.Server = $DomainController }
            Disable-ADAccount @disableParams
            Write-Log "Account disabled for $($User.SamAccountName)" "INFO" "Yellow"
        }
        
        $result.Status = "Success"
        $Script:SuccessCount++
        
        # Send notification if enabled
        if ($Options.NotifyUsers -and $User.EmailAddress) {
            $domain = $User.DistinguishedName -replace '^.*?DC=([^,]+),.*$','$1'
            Send-PasswordNotification -UserEmail $User.EmailAddress -UserName $User.SamAccountName -UserDisplayName $User.DisplayName -NewPassword $NewPassword -Domain $domain
        }
        
    }
    catch {
        $errorMsg = $_.Exception.Message
        Write-Log "Failed to reset password for $($User.SamAccountName): $errorMsg" "ERROR" "Red"
        $result.Status = "Failed"
        $result.Error = $errorMsg
        $Script:ErrorCount++
    }
    
    return $result
}

function Export-Results {
    param(
        [System.Collections.ArrayList]$Results,
        [string]$ExportPassword = $null
    )
    
    $exportTime = Get-Date -Format "yyyyMMdd_HHmmss"
    
    # Export summary report
    $summaryPath = Join-Path $LogPath "PasswordReset_Summary_$exportTime.csv"
    
    $summaryData = $Results | Select-Object SamAccountName, DisplayName, Email, Status, Error, NewPassword
    $summaryData | Export-Csv -Path $summaryPath -NoTypeInformation
    Write-Log "Summary report exported to: $summaryPath" "INFO" "Green"
    
    # Export passwords if requested
    if ($ExportPasswords) {
        $passwordPath = Join-Path $LogPath "PasswordReset_Export_$exportTime.csv"
        
        # Filter only successful resets
        $successful = $Results | Where-Object { $_.Status -eq "Success" }
        
        if ($SecureExport) {
            # Export with password protection (simplified DPAPI encryption)
            try {
                $tempFile = [System.IO.Path]::GetTempFileName()
                $successful | Select-Object SamAccountName, DisplayName, NewPassword | Export-Csv -Path $tempFile -NoTypeInformation
                
                # Encrypt using Windows DPAPI
                $content = Get-Content -Path $tempFile -Raw
                $secureBytes = [System.Security.Cryptography.ProtectedData]::Protect(
                    [System.Text.Encoding]::UTF8.GetBytes($content),
                    $null,
                    [System.Security.Cryptography.DataProtectionScope]::CurrentUser
                )
                
                $secureContent = [System.Convert]::ToBase64String($secureBytes)
                $secureContent | Set-Content -Path $passwordPath
                
                Write-Log "Encrypted passwords exported to: $passwordPath (DPAPI Protected)" "INFO" "Green"
                Write-Log "⚠️  Password file is encrypted using Windows DPAPI - only your account can decrypt it" "WARNING" "Yellow"
                Write-Log "To decrypt: Use [System.Security.Cryptography.ProtectedData]::Unprotect()" "INFO" "Cyan"
                
                Remove-Item -Path $tempFile -Force
            }
            catch {
                Write-Log "Failed to encrypt password export: $_" "ERROR" "Red"
            }
        }
        else {
            $successful | Select-Object SamAccountName, DisplayName, NewPassword | Export-Csv -Path $passwordPath -NoTypeInformation
            Write-Log "⚠️  Passwords exported in plain text to: $passwordPath" "WARNING" "Yellow"
            Write-Log "   Please secure this file and delete after use!" "WARNING" "Yellow"
        }
    }
}
#endregion

#region Main Script
function Main {
    # Initialize logging
    Initialize-Logging
    
    # Check AD module
    if (-not (Import-ADModule)) {
        Write-Log "Script cannot continue without ActiveDirectory module." "ERROR" "Red"
        return
    }
    
    # Load user list
    $userList = @()
    
    if ($InputFile) {
        $csvData = Import-CSVUsers -FilePath $InputFile
        if (-not $csvData) {
            Write-Log "No users loaded from CSV. Exiting." "ERROR" "Red"
            return
        }
        
        # Determine the column to use for identity
        $identityColumn = $null
        $possibleColumns = @('SamAccountName', 'UserPrincipalName', 'UPN', 'Username', 'Identity', 'DN', 'DistinguishedName')
        foreach ($col in $possibleColumns) {
            if ($csvData[0].PSObject.Properties.Name -contains $col) {
                $identityColumn = $col
                break
            }
        }
        
        if (-not $identityColumn) {
            Write-Log "CSV must contain one of: SamAccountName, UserPrincipalName, UPN, Username, Identity" "ERROR" "Red"
            Write-Log "Available columns: $($csvData[0].PSObject.Properties.Name -join ', ')" "ERROR" "Red"
            return
        }
        
        Write-Log "Using '$identityColumn' as identity column" "INFO" "Cyan"
        $userList = $csvData | ForEach-Object { $_.$identityColumn }
    }
    elseif ($Users) {
        $userList = $Users
    }
    else {
        Write-Log "No users specified. Use -InputFile or -Users parameter." "ERROR" "Red"
        return
    }
    
    if ($userList.Count -eq 0) {
        Write-Log "No users found to process." "ERROR" "Red"
        return
    }
    
    Write-Log "Processing $($userList.Count) users..." "INFO" "Cyan"
    
    # Setup options
    $options = @{
        ForceChangeAtNextLogon = $ForceChangeAtNextLogon
        NotifyUsers = $NotifyUsers
        ExportPasswords = $ExportPasswords
        UnlockUsers = $UnlockUsers
        DisableUser = $DisableUser
    }
    
    # Process each user
    $totalUsers = $userList.Count
    $currentUser = 0
    
    foreach ($identity in $userList) {
        $currentUser++
        Write-Log "Processing user $currentUser of $totalUsers : $identity" "INFO" "Cyan"
        
        # Get user object
        $user = Get-UserObject -Identity $identity -DC $DomainController
        if (-not $user) {
            $result = [PSCustomObject]@{
                SamAccountName = $identity
                DisplayName = $identity
                Email = $null
                Status = "Failed"
                Error = "User not found"
                NewPassword = "***"
            }
            $Script:Results.Add($result) | Out-Null
            $Script:ErrorCount++
            continue
        }
        
        # Check if account is enabled
        if (-not $user.Enabled) {
            Write-Log "User $($user.SamAccountName) is disabled. Skipping." "WARNING" "Yellow"
            $result = [PSCustomObject]@{
                SamAccountName = $user.SamAccountName
                DisplayName = $user.DisplayName
                Email = $user.EmailAddress
                Status = "Skipped"
                Error = "Account disabled"
                NewPassword = "***"
            }
            $Script:Results.Add($result) | Out-Null
            continue
        }
        
        # Generate password
        $password = $null
        if ($PasswordGeneration -eq 'Custom' -and $CustomPassword) {
            $password = $CustomPassword
            Write-Log "Using custom password for $($user.SamAccountName)" "INFO" "Cyan"
        }
        elseif ($PasswordGeneration -eq 'Dictionary') {
            $password = Generate-DictionaryPassword -Length $PasswordLength
            Write-Log "Generated dictionary password for $($user.SamAccountName)" "INFO" "Cyan"
        }
        else {
            $password = Generate-RandomPassword -Length $PasswordLength
            Write-Log "Generated random password for $($user.SamAccountName)" "INFO" "Cyan"
        }
        
        # Validate password complexity
        if (-not (Test-PasswordComplexity -Password $password)) {
            Write-Log "Generated password may not meet complexity requirements for $($user.SamAccountName)" "WARNING" "Yellow"
        }
        
        # Reset password
        $result = Reset-UserPassword -User $user -NewPassword $password -Options $options
        $Script:Results.Add($result) | Out-Null
    }
    
    # Generate final summary
    Write-Log "="*70 "INFO" "Yellow"
    Write-Log "OPERATION COMPLETE" "INFO" "Yellow"
    Write-Log "="*70 "INFO" "Yellow"
    Write-Log "Total Users Processed: $totalUsers" "INFO" "Cyan"
    Write-Log "Successful Resets: $Script:SuccessCount" "INFO" "Green"
    Write-Log "Failed Resets: $Script:ErrorCount" "INFO" "Red"
    Write-Log "Dry Run Mode: $DryRun" "INFO" "Yellow"
    Write-Log "="*70 "INFO" "Yellow"
    
    # Export results
    Export-Results -Results $Script:Results
    
    Write-Log "Log file: $Script:LogFile" "INFO" "Cyan"
    Write-Log "End Time: $(Get-Date)" "INFO" "Cyan"
    Write-Log "="*70 "INFO" "Yellow"
}

# Run the script
Main
#endregion