###########################################################
# Script Name: AD_Tools.ps1
# Created On: Feb 3, 2017
# Author: David Frohlick
# 
# Purpose: Grouping of Powershell AD Tools
#  
# Script Version: 1.0 - Intial Script
###########################################################

Add-Type -AssemblyName Microsoft.VisualBasic

Function AddSCCMGroup {
    $DN = 'OU=SCCM,OU=Groups,DC=saskenergy,DC=net'
    Get-ADObject -Filter * -SearchBase $DN -Properties description | Select-Object name | Sort-Object -Property name | Out-File "\\skynet01\dfrohlick\Home\Stuff\Scripts\Powershell\Lists\SCCMList.txt" -Force
    $SCCMList = Get-Content "\\skynet01\dfrohlick\Home\Stuff\Scripts\Powershell\Lists\SCCMList.txt"
    $SCCMList | ForEach {$_.TrimEnd()} | Set-Content "\\skynet01\dfrohlick\Home\Stuff\Scripts\Powershell\Lists\SCCM_List.txt" -Force


    $Msg = "What is the name of the PC?"
    $PC= [Microsoft.VisualBasic.Interaction]::InputBox($Msg,"PC Name","PC")
    $PCSAM = $PC + "$"

    $Msg = "What is the software? (ie Visio, Acrobat Standard, etc)"
    $Group = [Microsoft.VisualBasic.Interaction]::InputBox($Msg,"Group Name","")

    ForEach ($Item in Get-Content "\\skynet01\dfrohlick\Home\Stuff\Scripts\Powershell\Lists\SCCM_List.txt") {
        If ($Item -like "*$Group*") {
            $Msg = "Confirm group (yes/no) = $Item"
            $Confirm = [Microsoft.VisualBasic.Interaction]::InputBox($Msg,"Confirm","Yes")
            If (($Confirm -match 'y') -or ($Confirm -match 'Yes')) {
                Add-ADGroupMember "$Item" -Members $PCSAM
                Exit
            }
        }
    }
}

Function CopyADGroup {
    #Gets PC names and creates the samaccountname for the new one
    $Msg = "What is the old PC Name?"
    $OldPC =  [Microsoft.VisualBasic.Interaction]::InputBox($Msg,"Old PC Name","PC")
    $Msg = "What is the new PC Name?"
    $NewPC = [Microsoft.VisualBasic.Interaction]::InputBox($Msg,"New PC Name","PC")
    $NewPCSAM = $NewPC + "$"
    $ErrorActionPreference = "SilentlyContinue"

    #Grabs list of groups for the old pc
    $Groups = Get-ADComputer $OldPC -Properties MemberOf | Select-Object -ExpandProperty MemberOf

    ForEach ($Group in $Groups) {
    
        #Shortens it down to just the group name
        $Item = $Group.TrimStart("CN=")
        $Item = $Item.Split(',')[0]
    
        #If it's a SCCM group, adds the new pc to the group
            If ($Item -like '*SCCM*') {
                Add-ADGroupMember "$Item" -Members $NewPCSAM 
                Write-Host "$NewPC has been added to $Item`n" -ForegroundColor Green
            }

        #If it's not a SCCM group, prompts on whether it should add it or not
        If ($Item -notlike '*SCCM*') {  
            Write-Host "`n$OldPC was also a member of $Item"
            $Other = Read-Host -Prompt "Would you like to add $NewPC to this group now? [Y/N]"
                If ($Other -like 'y') {
                    Add-ADGroupMember "$Item" -Members $NewPCSAM
                    Write-Host "$NewPC has been added to $Item`n" -ForegroundColor Green
                }
        }
    }
}

Function Choice {
    $Msg = "What would you like to do?`n 1. Add PC to SCCM Group`n 2. Copy AD Groups from Old PC to New"
    $Choice = [Microsoft.VisualBasic.Interaction]::InputBox($Msg,"AD Tool Selection","")

    If ($Choice -eq '1') {AddSCCMGroup}
    If ($Choice -eq '2') {CopyADGroup}
}

Choice