###########################################################
# Script Name: Ping_Test.ps1
# Created On: Mar 9, 2016
# Author: David Frohlick
# 
# Purpose: Ping Test
#  
###########################################################

Function Info {
    Read-Host "Enter Computer Name" >> \\skynet01\dfrohlick\Home\Stuff\Scripts\PowerShell\Temp\PCList.txt
    Write-Host
    Write-Host "1. Add Additional PC |  2. Start Test"
    [Int]$MenuChoice = Read-Host "Enter Option"

    Switch($MenuChoice) {
        1{Info}
        2{Ping}
    }
}


Function Ping {
    $Repeat = 1
    While ($Repeat -eq '1')
    {
        ForEach($System in Get-Content "\\skynet01\dfrohlick\Home\Stuff\Scripts\PowerShell\Temp\PCList.txt") {
                If(Test-Connection $System -count 1 -quiet) {
                    $Path = "\\" + $System + "\c$"
                    If (Test-Path $Path) {
                        Write-Host "$System is Online" -foregroundcolor "Green"
                    }
                    Else {
                        Write-Host "$System is Offline" -ForegroundColor "Red"
                   }
                }
                Else {
                    Write-Host "$System is Offline" -foregroundcolor "Red"
                } 
        }
    
        Write-Host "Repeat test?"
        Write-Host "1. Yes | 2. No"
        $Repeat = Read-Host -prompt "Enter Option"
    }
    Finish
}  

Function Finish {
    Remove-Item \\skynet01\dfrohlick\Home\Stuff\Scripts\PowerShell\Temp\PCList.txt
}
  
Info
      