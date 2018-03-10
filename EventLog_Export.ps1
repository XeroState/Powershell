###########################################################
# Script Name: EventLog_Export.ps1
# Created On: Sept 26, 2016
# Author: David Frohlick
# 
# Purpose: Exports EventLog on Remote PC
#
#Version : 1.0 - Initial Script  
###########################################################

$PC = Read-Host "What PC?"
$Computer = "\\" + $PC

$Type = Read-Host "Which log? [ie Application, Security, System, etc)"

Write-Host "EventLog will be saved to C:\$Type.evtx on the remote machine"

psexec $Computer cmd /c "wevtutil epl $Type C:\$Type.evtx"