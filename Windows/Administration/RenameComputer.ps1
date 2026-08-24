# Rename a local computer
Local computer
Rename-Computer -NewName "YourNewComputerName" -Restart

# Rename a network computer
Remote computer
Rename-Computer -ComputerName "RemoteComputerName" -NewName "NewComputerName" -DomainCredential Domain01\Admin01 -Force