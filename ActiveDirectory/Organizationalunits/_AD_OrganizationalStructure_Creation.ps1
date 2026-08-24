Import-Module ActiveDirectory

# Shown at the end of the tool.
$AdminName    = "Luis Ramirez"
$AdminTitle   = "IT Support Specialist | Cloud & Systems Administrator"
$ScriptEnd = "$AdminName - $AdminTitle"

$DomainDN = (Get-ADDomain).DistinguishedName

$OUDescriptions = @{
    "_Administration"="Stores privileged administrative accounts, service accounts, and administrative workstations."
    "Admin_Workstations"="Dedicated computers used by IT administrators for privileged administration tasks."
    "Privileged_Accounts"="Administrative user accounts with elevated permissions separate from standard user accounts."
    "Service_Accounts"="Accounts used by applications, services, scheduled tasks, and automated processes."
    "_Departments"="Parent OU containing all company departments and their associated users, computers, and administrators."
    "Finance"="Organizational Unit for the Finance department."
    "Fin_Admins"="Administrative accounts responsible for managing Finance department resources."
    "Fin_Computers"="Workstations assigned to Finance department employees."
    "Fin_Users"="Standard user accounts belonging to the Finance department."
    "Human_Resources"="Organizational Unit for the Human Resources department."
    "HR_Admins"="Administrative accounts responsible for HR department management."
    "HR_Computers"="Workstations assigned to Human Resources personnel."
    "HR_Users"="Standard user accounts belonging to Human Resources."
    "Information_Technology"="Organizational Unit for the Information Technology department."
    "IT_Admins"="Administrative accounts for IT personnel with delegated administrative rights."
    "IT_Computers"="Workstations assigned to Information Technology staff."
    "IT_Users"="Standard user accounts belonging to the Information Technology department."
    "Operations"="Organizational Unit for the Operations department."
    "Ops_Admins"="Administrative accounts responsible for Operations department resources."
    "Ops_Computers"="Workstations assigned to Operations personnel."
    "Ops_Users"="Standard user accounts belonging to the Operations department."
    "Talent_Acquisition"="Organizational Unit for the Talent Acquisition (Recruitment) department."
    "TA_Admins"="Administrative accounts responsible for Talent Acquisition resources."
    "TA_Computers"="Workstations assigned to Talent Acquisition staff."
    "TA_Users"="Standard user accounts belonging to the Talent Acquisition department."
    "_Disabled Objects"="Temporary holding OU for disabled user and computer accounts pending deletion or archival."
    "Dis_Computers"="Disabled computer accounts removed from production."
    "Dis_Users"="Disabled user accounts awaiting retention expiration or permanent removal."
    "_Groups"="Stores all Active Directory security and distribution groups."
    "Applications"="Security groups used to assign permissions to enterprise applications."
    "Distribution"="Distribution groups used exclusively for email communications."
    "Files_Shares"="Security groups controlling access to shared folders and file servers."
    "Microsoft_365"="Groups used for Microsoft 365 licensing, cloud services, and Microsoft Entra ID synchronization."
    "Security"="General-purpose security groups for delegated permissions and access control."
    "_Servers"="Contains all domain member servers for centralized management and Group Policy application."
    "Application_Servers"="Servers hosting enterprise applications and business services."
    "Database_Servers"="Servers hosting SQL and other database management systems."
    "File_Servers"="Servers providing centralized file storage and shared folders."
    "Infrastructure"="Servers providing core infrastructure services such as DHCP, DNS, PKI, monitoring, and management."
    "Management"="Servers dedicated to systems management, automation, backup, patch management, and remote administration."
}

$script:CreatedCount=0
$script:ExistingCount=0

function New-OUIfMissing {
    param([string]$Name,[string]$ParentDN)
    $dn="OU=$Name,$ParentDN"
    if(-not(Get-ADOrganizationalUnit -LDAPFilter "(ou=$Name)" -SearchBase $ParentDN -ErrorAction SilentlyContinue)){
        New-ADOrganizationalUnit -Name $Name -Path $ParentDN -Description $OUDescriptions[$Name] -ProtectedFromAccidentalDeletion $true
        $script:CreatedCount++
        Write-Host "Created: $dn"
    } else {
        $script:ExistingCount++
        Write-Host "Already exists: $dn"
    }
}

$Tree=@{
"_Administration"=@("Admin_Workstations","Privileged_Accounts","Service_Accounts");
"_Departments"=@{
"Finance"=@("Fin_Admins","Fin_Computers","Fin_Users");
"Human_Resources"=@("HR_Admins","HR_Computers","HR_Users");
"Information_Technology"=@("IT_Admins","IT_Computers","IT_Users");
"Operations"=@("Ops_Admins","Ops_Computers","Ops_Users");
"Talent_Acquisition"=@("TA_Admins","TA_Computers","TA_Users")};
"_Disabled Objects"=@("Dis_Computers","Dis_Users");
"_Groups"=@("Applications","Distribution","Files_Shares","Microsoft_365","Security");
"_Servers"=@("Application_Servers","Database_Servers","File_Servers","Infrastructure","Management")}

Clear-Host
Write-Host "The following Organizational Units will be created:`n" -ForegroundColor Yellow
foreach($Root in $Tree.Keys){
 Write-Host $Root -ForegroundColor Cyan
 if($Tree[$Root] -is [hashtable]){
  foreach($Child in $Tree[$Root].Keys){
   Write-Host "├── $Child" -ForegroundColor Cyan
   foreach($OU in $Tree[$Root][$Child]){Write-Host "│   └── $OU" -ForegroundColor Cyan}
  }
 } else {
  foreach($OU in $Tree[$Root]){Write-Host "└── $OU" -ForegroundColor Cyan}
 }
 Write-Host
}

do {
    $choice = (Read-Host "Type 'continue' to create the OUs or 'exit' to cancel").Trim()

    if ($choice -ceq "continue") {
        break
    }

    if ($choice -ceq "exit") {
        Clear-Host
        Start-Sleep -Milliseconds 50
        Write-Host ""
        Write-Host "Operation cancelled." -ForegroundColor Yellow
        Write-Host ""
        Write-Host $ScriptEnd -ForegroundColor DarkGray
        Write-Host ""
        return
    }

    Write-Host ""
    Write-Host "Invalid input." -ForegroundColor Red
    Write-Host "Please type exactly:" -ForegroundColor Yellow
    Write-Host "  continue" -ForegroundColor Green
    Write-Host "  exit" -ForegroundColor Green
    Write-Host ""

} while ($true)

foreach($k in $Tree.Keys){
 New-OUIfMissing $k $DomainDN
 $v=$Tree[$k]
 if($v -is [hashtable]){
  foreach($c in $v.Keys){
   New-OUIfMissing $c "OU=$k,$DomainDN"
   foreach($s in $v[$c]){New-OUIfMissing $s "OU=$c,OU=$k,$DomainDN"}
  }
 } else {
  foreach($s in $v){New-OUIfMissing $s "OU=$k,$DomainDN"}
 }
}

Clear-Host
Write-Host ""
if($script:CreatedCount -eq 0){
Write-Host "All requested Organizational Units already exist. No new OUs were created." -ForegroundColor Yellow
} elseif($script:ExistingCount -gt 0){
Write-Host "Organizational Units created. Existing OUs were skipped." -ForegroundColor Green
} else {
Write-Host "All Organizational Units have been created successfully." -ForegroundColor Green
}
Write-Host ""
Write-Host $ScriptEnd -ForegroundColor DarkCyan
Start-Sleep -Seconds 2
