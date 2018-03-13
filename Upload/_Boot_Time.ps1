###########################################################
# Script Name: _Boot_Time.ps1
# Created On: May 17, 2016
# Author: David Frohlick
# 
# Purpose: Quick Script to Output Boot Time
#  
###########################################################


Try {    
    #Gets PC Name
    $ComputerName = Read-Host -prompt "What is the PC Name?"
    
    #Set's WMI Variable
    $OS = Get-WMIObject -Class Win32_OperatingSystem -Namespace "root\cimv2" -ComputerName $ComputerName
    
    #Gets the boot time and converts it
    $Boot = $OS.ConvertToDateTime($OS.LastBootUpTime)
    $Days = (New-TimeSpan -Start $Boot).TotalDays

    #Based on time, outputs text in different colours
    If ($Days -lt '5') {Write-Host ("Last boot time was " + $Boot) -foregroundcolor Green}
    If ($Days -ge '5' -and $Days -lt '10') {Write-Host ("Last boot time was " + $Boot) -foregroundcolor Yellow}
    If ($Days -ge '10') {Write-Host ("Last boot time was " + $Boot) -foregroundcolor Red}

    Pause
}
Catch {
    Write-Host "Operation Failed"
    Pause
}