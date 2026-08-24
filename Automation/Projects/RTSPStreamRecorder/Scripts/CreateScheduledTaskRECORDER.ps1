$action = New-ScheduledTaskAction -Execute "powershell.exe" `
-Argument "-WindowStyle Hidden -ExecutionPolicy Bypass -File D:\OutdoorCamera\Scripts\record.ps1"

$trigger = New-ScheduledTaskTrigger -AtStartup

Register-ScheduledTask `
-TaskName "OutdoorCameraRecorder" `
-Action $action `
-Trigger $trigger `
-RunLevel Highest `
-User "SYSTEM"
