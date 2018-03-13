###########################################################
# Script Name: VM_Snapshot.ps1
# Created On: June 20, 2016
# Author: David Frohlick
# 
# Purpose: Automates the shutting down and snapshot process
#  
# Version: 2.0 - Updated for new PC
#                 1.5 - Attempted to do both old PC and new PC, failed due to VMWare
#                 1.1 - Cleaned up
#                 1.0 - Initial Script
###########################################################


Function Vari {
    #Sets variables

     $Date = Get-Date -format M
    $VMFolder = "V:\VMs"
       
    Work
}

Function Work {
    #Finds VMs, finds snapshots, asks to delete them and asks to create new
    
    Set-Location "C:\Program Files (x86)\VMWare\VMWare Workstation"

    #Gets a list of all the VMs in that folder
    Get-ChildItem $VMFolder -recurse | Where-Object {$_ -like '*.vmx'} | % {$_.FullName} | Out-File "\\skynet01\dfrohlick\home\stuff\scripts\powershell\temp\VMs.txt"
    $VMs = Get-Content "\\skynet01\dfrohlick\home\stuff\scripts\powershell\temp\VMs.txt" | Where-Object {($_ -notlike 'FullName') -or ($_ -notlike '*-*')}

    #Stops each VM
    ForEach ($VM in $VMs) {
        .\VMrun -T ws stop `"$VM`" soft
        
        #Gets a list of the snapshots
        .\VMRun -T ws listsnapshots "$VM" | Where-Object {$_ -notlike '*Total snapshot*'} | Out-File "\\skynet01\dfrohlick\home\stuff\scripts\powershell\temp\Snaps.txt"
        $Snapshots = Get-Content "\\skynet01\dfrohlick\home\stuff\scripts\powershell\temp\Snaps.txt"

        #For each snapshot, asks if it should be deleted
        ForEach ($Snap in $Snapshots) {
            $Choice = Read-Host "Delete $Snap on $VM snapshot? [Y/N]"
                If ($Choice -like 'y') {
                    Write-Host "Deleting... This may take awhile"
                    .\VMRun -T ws deleteSnapshot `"$VM`" $Snap
                }
        }

        #Asks if a new snapshot should be made
        $Que = Read-Host "Take Snapshot of $VM now? [Y/N]"
        If ($Que -like 'y') {
            .\VMRun -T ws snapshot `"$VM`" "$Date - Off"
        }
       
        #Removes snapshot list
        Remove-Item \\skynet01\dfrohlick\home\stuff\scripts\powershell\temp\Snaps.txt
    }
    
    #Removes vm list
    Remove-Item \\skynet01\dfrohlick\home\stuff\scripts\powershell\temp\VMs.txt
    
}

Vari