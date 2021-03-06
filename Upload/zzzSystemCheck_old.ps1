###########################################################
# Script Name: SystemCheck.ps1
# Created On: Mar 16, 2016
# Author: David Frohlick
# 
# Purpose: Get system info from WMI
#  
###########################################################

#Requires -version 3

Function Info {
    #Get Computer and test to see if it's online
    $ComputerName = Read-Host -Prompt "What is the computer name?"
    

    If(Test-Connection $ComputerName -count 1 -quiet) {
        ComputerSystem
    }
    Else {
        Write-Host $ComputerName" is offline"
        Pause
    }
}

Function ComputerSystem {
    #Grab physical memory and the currently logged on username
    $CompSystem = Get-WMIObject -Class Win32_ComputerSystem -Namespace "root\cimv2" -ComputerName $ComputerName
    ForEach ($Item in $CompSystem) {
        
        $RAM = [Math]::Round(($Item.TotalPhysicalMemory/1GB),2)
        $User = $Item.UserName
    }
    $User = $User.TrimStart("SKENERGY\")
    $Person = Get-ADUser $User -Properties *
    PhysicalMem
}

Function PhysicalMem {
    #Get the readable installed memory number
    $InstalledRam = (Get-WMIObject -Class Win32_PhysicalMemory -Namespace "root\cimv2" -Computername $ComputerName | Measure-Object -Property Capacity -Sum | % {[Math]::Round(($_.sum/1GB),2)})
   
    Processor
}

Function Processor {
    #Get CPU information 
    $Processor = Get-WMIObject -Class Win32_Processor -Namespace "root\cimv2" -ComputerName $ComputerName
    ForEach ($Item in $Processor) {
        $CPUName = $Item.Name
        $CPULogic = $Item.NumberOfLogicalProcessors
        $CPUPhys = $Item.NumberofCores
        $CPULoad = $Item.LoadPercentage
    }
    SystemProduct
}

Function SystemProduct {
    #Get machine information  
    $SystemProduct = Get-WMIObject -Class Win32_ComputerSystemProduct -Namespace "root\cimv2" -ComputerName $ComputerName
    ForEach ($Item in $SystemProduct) {
        $CompType = "Lenovo " + $Item.Version + " (" + $Item.Name + "), S/N " + $Item.IdentifyingNumber
    }
    OS
}

  
Function OS {
    #Get last boot time and free ram  
    $OS = Get-WMIObject -Class Win32_OperatingSystem -Namespace "root\cimv2" -ComputerName $ComputerName
    ForEach ($Item in $OS) {
            $Boot = $OS.ConvertToDateTime($OS.LastBootUpTime)
            $FreeRAM = [Math]::Round(($OS.FreePhysicalMemory/1MB),2)
    }
    $Days = (New-TimeSpan -Start $Boot).TotalDays
    HDD
}

Function HDD {
    #Get C:\ and D:\ hdd values
    $HDD1 = Get-WMIObject -Class Win32_LogicalDisk -Namespace "root\cimv2" -ComputerName $ComputerName -Filter "DeviceID='C:'"
    ForEach ($Item in $HDD1) {
        $CSize = [Math]::Round(($HDD1.Size/1Gb),2)
        $CFreeSpace = [Math]::Round(($HDD1.Freespace/1GB),2)
    }
    $HDD2 = Get-WMIObject -Class Win32_LogicalDisk -Namespace "root\cimv2" -ComputerName $ComputerName -Filter "DeviceID='D:'"
    ForEach ($Item in $HDD2) {
        $DSize = [Math]::Round(($HDD2.Size/1Gb),2)
        $DFreeSpace = [Math]::Round(($HDD2.Freespace/1GB),2)
    }
    BIOS
}

Function BIOS {
    #Get BIOS Version
    $BIOS = Get-WMIObject -Class Win32_BIOS -Namespace "root\cimv2" -ComputerName $ComputerName
    ForEach ($Item in $BIOS) {
        $BIOSv = $BIOS.SMBIOSBIOSVersion
        $BIOSDate = $BIOS.ConvertToDateTime($BIOS.ReleaseDate)
        
    }
    Processes
}

Function Processes {
    #Get RAM use for Internet Explorer, VMWare products, Lotus Notes and MS Office
    $Process = Get-WmiObject -Class Win32_Process -Namespace "root\cimv2" -ComputerName $ComputerName   
    $IERam = $Process | Where-Object {$_.Description -like 'iexplore.exe'} | Measure-Object -Property WorkingSetSize -Sum | % {[Math]::Round(($_.sum/1GB),2)}
    $NotesRam = $Process | Where-Object {$_.Description -like '*notes*'} | Measure-Object -Property WorkingSetSize -Sum | % {[Math]::Round(($_.sum/1GB),2)}
    $VMRam = $Process | Where-Object {$_.Description -like '*vmware*' -or $_.Description -like '*vpx*'} | Measure-Object -Property WorkingSetSize -Sum | % {[Math]::Round(($_.sum/1GB),2)}
    $OfficeRam = $Process | Where-Object {$_.Description -like '*excel*' -or $_.Description -like '*word*' -or $_.Description -like '*powerpnt*' -or $_.Description -like '*onenote*'} | Measure-Object -Property WorkingSetSize -Sum | % {[Math]::Round(($_.sum/1GB),2)}
    Output
}
    
Function Output {
    #Writes it all to the screen

    Clear-Host
    Write-Host
    Write-Host "----------------- For " $ComputerName " -----------------"
    Write-Host
    Write-Host ("Currently logged on User is: " + $Person.displayName)
    If ($Days -lt '5') {Write-Host ("Last boot time was " + $Boot) -foregroundcolor Green}
    If ($Days -ge '5' -and $Days -lt '10') {Write-Host ("Last boot time was " + $Boot) -foregroundcolor Yellow}
    If ($Days -ge '10') {Write-Host ("Last boot time was " + $Boot) -foregroundcolor Red}
    Write-Host
    Write-Host ($CompType)
    Write-Host ("It has an " + $CPUName + ", with " + $CPUPhys + " cores (" + $CPULogic + " logical).")
    Write-Host ("CPU is at " + $CPULoad + "% load.")
    Write-Host
    Write-Host ("There is " + $FreeRam + "GB of " + $RAM + "GB RAM is free. " + $InstalledRAM + "GB is installed.")
    Write-Host ("Primary disk has " + $CFreeSpace + "GB free of " + $CSize + "GB.")
    If ($DSize) {Write-Host ("Secondary disk has " + $DFreeSpace + "GB free of " + $DSize + "GB.")}
    Write-Host
    If ($IERam) {Write-Host ("Internet Explorer is using "+ $IERAM + "GB of RAM.")}
    If ($NotesRam) {Write-Host ("Lotus Notes is using " + $NotesRam + "GB of RAM.")}
    If ($OfficeRam) {Write-Host ("Microsoft Office is using " + $OfficeRam + "GB of RAM.")}
    If ($VMRAM) {Write-Host ("VMware is using " + $VMRam + "GB of RAM.")}
    Write-Host
    Write-Host ("BIOS version is " + $BIOSv + ", with a release date of " + $BIOSDate)
    Write-Host
    
    Pause
}

Info