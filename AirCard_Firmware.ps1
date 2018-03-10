####################################################################
# Script Name: AirCard_Firmware.ps1
# Created On: Sept 22nd, 2016
# Author: David Frohlick
# 
# Purpose: Grabs AirCard Firmware Info
#
####################################################################

    $PC = Read-Host "What is the PC name?"
    $Computer = "\\" + $PC
    $Temp = $Computer + "\C$\Temp\Temp\"

    If (Test-Path "$Computer\C$") {
        
        Copy-Item -Path \\skynet01\dfrohlick\Home\Stuff\Scripts\Powershell\Do Not Use\Firmware.ps1 -Destination (New-Item $Temp -Type Container -Force) -container -Force
        ($Info = psexec.exe $Computer /accepteula cmd /c "powershell -executionpolicy bypass -file C:\temp\temp\Firmware.ps1") | Out-Null
        Get-Content "$Temp\FW.txt"
        Remove-Item $Temp -Force -Recurse -ErrorAction 'SilentlyContinue'
        Pause
    }
    Else {
        Write-Host "PC is offline"
    }

