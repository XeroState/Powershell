
#Sets Variables
$PC = Read-Host "What is the name of the PC?"
$RemotePC = "\\" + $PC
$FilePath = $RemotePC + "\c$\Update.ps1"
$Exec = "C:\Windows\System32\PsExec.exe"

#Copies script to local machine
Write-Host "Copying Update Script"
Copy-Item "\\skynet01\dfrohlick\home\Stuff\Scripts\PowerShell\Update.ps1"  $FilePath

#Changes policy, runs script, than changes policy back
Write-Host "Setting ExecutionPolicy to unrestricted"
(& $Exec $RemotePC -s -h -i powershell -inputformat none "set-executionpolicy unrestricted -force")  2>&1 | Out-Null
Write-Host "Running Update Script"
(& $Exec $RemotePC -s -h -i cmd /c powershell -file c:\RemoteUpdate.ps1)  2>&1 | Out-Null
Write-Host "Setting ExecutionPolicy to restricted"
(& $Exec $RemotePC -s -h -i cmd /c powershell -inputformat none "set-executionpolicy restricted -force")  2>&1 | Out-Null

#Deletes script
Write-Host "Removing Update Script"
Remove-Item $FilePath