###########################################################
# Script Name: Copy_PC_AD_Groups.ps1
# Created On: May 19, 2016
# Author: David Frohlick
# 
# Purpose: Automate adding SCCM Groups
#  
# Version: 2.1 - The SCCM loop now avoids adding OpenText and Remove groups
#                 2.0 - Added ability to loop for more PC's
#                 1.1 - Added the ability to add non sccm groups
#                 1.0 - Initial Script
###########################################################

#Requires -version 3

Function CopyGroups {

        #Gets PC names and creates the samaccountname for the new one
        $OldPC = Read-Host -Prompt "What is the old PC name?"
        $NewPC = Read-Host -Prompt "What is the new PC name?"
        $NewPCSAM = $NewPC + "$"
        $ErrorActionPreference = "SilentlyContinue"

        #Grabs list of groups for the old pc
        $Groups = Get-ADComputer $OldPC -Properties MemberOf | Select-Object -ExpandProperty MemberOf

        ForEach ($Group in $Groups) {
    
            #Shortens it down to just the group name
            $Item = $Group.TrimStart("CN=")
            $Item = $Item.Split(',')[0]
    
            #If it's a SCCM group, adds the new pc to the group unless it's OpenText
            If ($Item -like '*SCCM*') {
                If($Item -notlike "*Remove*") {
                    If($Item -notlike "*OpenText*") {
                        Add-ADGroupMember "$Item" -Members $NewPCSAM 
                        Write-Host "$NewPC has been added to $Item`n" -ForegroundColor Green
                    }
                }
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

    $Choice = Read-Host "Another computer? [y/n]"
    If ($Choice -eq 'y') {CopyGroups}
    If ($Choice -eq 'n') {Exit}
}
CopyGroups