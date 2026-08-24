$ouList = Get-ADOrganizationalUnit -Filter * | Select-Object -ExpandProperty Name
$ouView = $ouList | Out-GridView -Title "Select an Organizational Unit" -OutputMode Single
$ouPath = (Get-ADOrganizationalUnit -Filter "Name -eq '$ouView'").DistinguishedName
Get-ADUser -Filter * -Properties EmailAddress,Name,Enabled,SamAccountName -SearchBase $ouPath | Select-Object EmailAddress,Name,Enabled,SamAccountName | Out-File C:\Users\lramirez\Documents\EmailA.txt

#$ou = "OU=StarWars,OU=_Users,DC=lcrtest,DC=com"
#Get-ADUser -Filter * -Properties EmailAddress -SearchBase $ou | Select-Object EmailAddress | Out-File C:\Users\lramirez\Documents\EmailA.txt
 
Get-ADUser -Filter * -Properties EmailAddress,Name,Enabled,SamAccountName -SearchBase $ouPath | Select-Object EmailAddress,Name,Enabled,SamAccountName | Export-Csv C:\Users\lramirez\Documents\EmailA.csv -NoTypeInformation