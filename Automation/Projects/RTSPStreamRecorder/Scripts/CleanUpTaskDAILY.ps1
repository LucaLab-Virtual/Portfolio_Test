$action = New-ScheduledTaskAction -Execute "powershell.exe" `
-Argument "-ExecutionPolicy Bypass -File D:\OutdoorCamera\Scripts\cleanup.ps1"

$trigger = New-ScheduledTaskTrigger -Daily -At 3pm

Register-ScheduledTask `
-TaskName "OutdoorCameraCleanup" `
-Action $action `
-Trigger $trigger `
-RunLevel Highest `
-User "SYSTEM"
