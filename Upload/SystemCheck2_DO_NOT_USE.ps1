###########################################################
# Script Name: SystemCheck2.ps1
# Created On: Apr 26, 2016
# Author: David Frohlick
# 
# Purpose: New approach using CMI
#          Requires WinRM which doesn't work for us until Win10 perhaps
#  
###########################################################


Function Info {
    #Get Computer and test to see if it's online
    $ComputerName = Read-Host -Prompt "What is the computer name?"
    $Today = Get-Date

    If(Test-Connection $ComputerName -count 1 -quiet) {
        OS
    }
    Else {
        Write-Host $ComputerName" is offline"
        Pause
    }
}


Function OS {
    $OS = Get-CimInstance Win32_OperatingSystem -Computer $ComputerName
    ForEach ($Item in $OS) {
        $FreeRam = $Item.FreePhysicalMemory
        $Boot = $Item.LastBootUpTime
        }
          
    $FreeRamValue = [Math]::Round(($FreeRam/1048576),2)
    $Days = (New-TimeSpan -Start $Boot).TotalDays

    ComputerSystem
}

Function ComputerSystem {

    $CS = Get-CimInstance Win32_ComputerSystem -Computer $ComputerName
    ForEach ($Item in $CS) {
        $TotalRam = $Item.TotalPhysicalMemory
        $User = $Item.UserName
    }
    $TotalRamValue = [Math]::Round(($TotalRam/1GB),2)
    $User = $User.TrimStart("SKENERGY\")
    $Person = Get-ADUser $User -Properties *


    Processor
}


Function Processor {
    
    $Processor = Get-CimInstance Win32_Processor -Computer $ComputerName
    ForEach ($Item in $Processor) {
        $CPULoad = $Item.LoadPercentage
        $CPUName = $Item.Name
        $CPUCores = $Item.NumberofCores
        $CPULogic = $Item.NumberofLogicalProcessors
    }

    SystemProduct
}


Function SystemProduct {
    
    $SystemProduct = Get-CimInstance Win32_ComputerSystemProduct -Computer $ComputerName
    ForEach ($Item in $SystemProduct) {
        $PCModel = $Item.Version
        $PCType = $Item.Name
        $PCSN = $Item.IdentifyingNumber
    }
    #Write-Host ("Lenovo " + $PCModel + "(" + $PCType + "), S/N " + $PCSN)
    HDD
}


Function HDD {

    $HDD1 = Get-CimInstance Win32_LogicalDisk -Computer $ComputerName -Filter "DeviceID='C:'"
    ForEach ($Item in $HDD1) {
        $CSize = [Math]::Round(($HDD1.Size/1Gb),2)
        $CFreeSpace = [Math]::Round(($HDD1.Freespace/1GB),2)
    }
    
    $HDD2 = Get-CimInstance Win32_LogicalDisk -Computer $ComputerName -Filter "DeviceID='D:'"
    ForEach ($Item in $HDD2) {
        $DSize = [Math]::Round(($HDD2.Size/1Gb),2)
        $DFreeSpace = [Math]::Round(($HDD2.Freespace/1GB),2)
    }

    BIOS
}

Function BIOS {

    
    $BIOS = Get-CimInstance Win32_BIOS -Computer $ComputerName
    ForEach ($Item in $BIOS) {
        $BIOSv = $Item.SMBIOSBIOSVersion
        $BIOSDate = $Item.ReleaseDate
    }

    Process
}


Function Process {

    $Process = Get-CimInstance Win32_Process -Computer PC12081
    $IERam = $Process | Where-Object {$_.Description -like '*iexplore*'} | Measure-Object -Property WorkingSetSize -Sum | % {[Math]::Round(($_.sum/1GB),2)}
    $NotesRam = $Process | Where-Object {$_.Description -like '*notes*'} | Measure-Object -Property WorkingSetSize -Sum | % {[Math]::Round(($_.sum/1GB),2)}
    $VMRam = $Process | Where-Object {$_.Description -like '*vmware*' -or $_.Description -like '*vpx*'} | Measure-Object -Property WorkingSetSize -Sum | % {[Math]::Round(($_.sum/1GB),2)}
    $OfficeRam = $Process | Where-Object {$_.Description -like '*excel*' -or $_.Description -like '*word*' -or $_.Description -like '*powerpnt*' -or $_.Description -like '*onenote*'} | Measure-Object -Property WorkingSetSize -Sum | % {[Math]::Round(($_.sum/1GB),2)}

    Output
}



Function Output {
    
    Clear-Host
    Write-Host
    Write-Host "----------------- For " $ComputerName " -----------------"
    Write-Host
    Write-Host ("Currently logged on User is: " + $Person.displayName)
    If ($Days -lt '5') {Write-Host ("Last boot time was " + $Boot) -foregroundcolor Green}
    If ($Days -gt '5') {Write-Host ("Last boot time was " + $Boot) -foregroundcolor Red}
    Write-Host
    Write-Host ($CompType)
    Write-Host ("It has a " + $CPUName + ", with " + $CPUCores + " cores (" + $CPULogic + " logical).")
    Write-Host ("CPU is at " + $CPULoad + "% load.")
    Write-Host
    Write-Host ("There is " + $FreeRamValue + "GB of " + $TotalRamValue + "GB RAM is free. ")
    Write-Host ("Primary disk has " + $CFreeSpace + "GB free of " + $CSize + "GB.")
    If ($DSize) {Write-Host ("Secondary disk has " + $DFreeSpace + "GB free of " + $DSize + "GB.")}
    Write-Host
    If ($IERam) {Write-Host ("Internet Explorer is using "+ $IERAM + "GB of RAM.")}
    If ($NotesRam) {Write-Host ("Lotus Notes is using " + $NotesRam + "GB of RAM.")}
    If ($OfficeRam) {Write-Host ("Microsoft Office is using " + $OfficeRam + "GB of RAM.")}
    If ($VMRAM) {Write-Host ("VMware is using " + $VMRam + "GB of RAM.")}
    Write-Host
    Write-Host ("BIOS version is " + $BIOSv + " with a release date of " + $BIOSDate)
    Write-Host
    
    Pause
}

Info