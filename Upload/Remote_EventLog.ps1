###########################################################
# Script Name: Remote_Registry.ps1
# Created On: Mar 14, 2016
# Author: David Frohlick
# 
# Purpose: Scan remote registry for specific event ID
#  
###########################################################

#Requires -version 3

Function Info {
    #Get information to be used
    $ComputerName = Read-Host "What is the PC Name?"
    $Days = Read-Host -Prompt 'How many days back to check the Event Log? [Numbers Only]'
    $Time = (get-date).AddDays("-" + $Days)
    
    HDD
}

Function HDD {
    #Looks for HDD errors in the System log
    Write-Host "System log for HDD Errors" -foregroundcolor "Red"
    $HDDEvents = Get-WinEvent -ComputerName $ComputerName -MaxEvents 25 -FilterHashtable @{LogName="System"; ProviderName="disk"; ID=7,11,51,52; StartTime=$Time} -ErrorAction SilentlyContinue
    $HDDEvents | Select-Object LogName, Message, TimeCreated | Export-CSV \\skynet01\dfrohlick\Home\Stuff\Scripts\Powershell\Temp\Event.csv -Append
    
    $HDDEvents2 = Get-WinEvent -ComputerName $ComputerName -MaxEvents 25 -FilterHashtable @{LogName="System"; ProviderName="NTFS"; ID=55; StartTime=$Time} -ErrorAction SilentlyContinue
    $HDDEvents2 | Select-Object LogName, Message, TimeCreated | Export-CSV \\skynet01\dfrohlick\Home\Stuff\Scripts\Powershell\Temp\Event.csv -Append
    
    Power
}

Function Power {
    #Looks for HDD errors in the System log
    Write-Host "System log for Power Errors" -foregroundcolor "Red"
    $HDDEvents = Get-WinEvent -ComputerName $ComputerName -MaxEvents 25 -FilterHashtable @{LogName="System"; ID=41; StartTime=$Time} -ErrorAction SilentlyContinue
    $HDDEvents | Select-Object LogName, Message, TimeCreated | Export-CSV \\skynet01\dfrohlick\Home\Stuff\Scripts\Powershell\Temp\Event.csv -Append
    
    $HDDEvents2 = Get-WinEvent -ComputerName $ComputerName -MaxEvents 25 -FilterHashtable @{LogName="System"; ProviderName="EventLog"; ID=6008; StartTime=$Time} -ErrorAction SilentlyContinue
    $HDDEvents2 | Select-Object LogName, Message, TimeCreated | Export-CSV \\skynet01\dfrohlick\Home\Stuff\Scripts\Powershell\Temp\Event.csv -Append
    
    AppLocker
}

Function AppLocker {
    #Look for exe errors in applocker
    Write-Host "AppLocker EXE and DLL Events" -foregroundcolor "Red"
    $EXEEvents = Get-WinEvent -ComputerName $ComputerName -MaxEvents 10 -FilterHashtable @{LogName="Microsoft-Windows-AppLocker/EXE and DLL"; ID=8004; StartTime=$Time} -ErrorAction SilentlyContinue
    $EXEEvents | Select-Object LogName, Message, TimeCreated | Export-CSV \\skynet01\dfrohlick\Home\Stuff\Scripts\Powershell\Temp\Event.csv -Append

    #Look for msi errors in applocker   
    Write-Host "AppLocker MSI and Script Events" -foregroundcolor "Red"
    $MSIEvents = Get-WinEvent -ComputerName $ComputerName -MaxEvents 10 -FilterHashtable @{LogName="Microsoft-Windows-AppLocker/MSI and Script"; ID=8007; StartTime=$Time}  -ErrorAction SilentlyContinue | Where-Object {$_.Message -notlike "*USERS.BAT*" -and $_.Message -notlike "*TOP.BAT*" -and $_.Message -notlike "*TS_BROKENSHORTCUTS.PS1*" -and $_.Message -notlike "*TS_UNUSEDDESKTOPICONS.PS1*"}
    $MSIEvents | Select-Object LogName, Message, TimeCreated | Export-CSV \\skynet01\dfrohlick\Home\Stuff\Scripts\Powershell\Temp\Event.csv -Append
    
    Printer
}

Function Printer {   
    #Look for printer driver errors
    Write-Host "Printer - Failed to Install Driver" -foregroundcolor "Red"
    $DriverEvents = Get-WinEvent -ComputerName $ComputerName -MaxEvents 10 -FilterHashtable @{LogName="Microsoft-Windows-PrintService/Admin"; ID=215; StartTime=$Time} -ErrorAction SilentlyContinue
    $DriverEvents | Select-Object LogName, Message, TimeCreated | Export-CSV \\skynet01\dfrohlick\Home\Stuff\Scripts\Powershell\Temp\Event.csv -Append
    
    #Look for printer adding errors
    Write-Host "Printer - Fail to Add Printer" -foregroundcolor "Red"
    $GPEvents = Get-WinEvent -ComputerName $ComputerName -MaxEvents 10 -FilterHashtable @{LogName="Microsoft-Windows-PrintService/Admin"; ID=513; StartTime=$Time} -ErrorAction SilentlyContinue
    $GPEvents | Select-Object LogName, Message, TimeCreated | Export-CSV \\skynet01\dfrohlick\Home\Stuff\Scripts\Powershell\Temp\Event.csv -Append

    #Look for print proccessor errors
    Write-Host "Printer - Print Processor fail" -foregroundcolor "Red"
    $GPEvents = Get-WinEvent -ComputerName $ComputerName -MaxEvents 10 -FilterHashtable @{LogName="Microsoft-Windows-PrintService/Admin"; ID=512; StartTime=$Time} -ErrorAction SilentlyContinue
    $GPEvents | Select-Object LogName, Message, TimeCreated | Export-CSV \\skynet01\dfrohlick\Home\Stuff\Scripts\Powershell\Temp\Event.csv -Append

    Output
}

Function Output {
    #Formats the csv and outputs it in a nice xml window
    Import-CSV -path \\skynet01\dfrohlick\Home\Stuff\Scripts\PowerShell\Temp\Event.csv | Sort TimeCreated | Export-Csv \\skynet01\dfrohlick\Home\Stuff\Scripts\Powershell\Temp\combined.csv -NoTypeInformation
    Import-CSV \\skynet01\dfrohlick\Home\Stuff\Scripts\Powershell\Temp\combined.csv | Out-GridView
    Remove-Item \\skynet01\dfrohlick\Home\Stuff\Scripts\Powershell\Temp\*.*
    Pause
}

Info