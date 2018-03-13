###########################################################################
# Name: Printer_Queue_Creation.ps1
#
# Version : 6.4 - Checks for BranchOffice and set's the property accordingly
#                 6.3 - Now correctly applies spooler attributes and restarts spooler
#                 6.2 - PrinterPort now checks to see if the port exists
#                 6.1 - Added HP reg key to disable status updates on the server
#                 6.0 - Added SNMP function to check the hp snmp value and add/update it if incorrect
#                 5.1 - Created a log file
#                 5.0 - Discovered various settings do not get set in auto-config. Manually forcing them
#                       - Note: It was due to SNMP being wrong so auto-config wouldn't work.
#                 4.3 - CSV now contains driver names so to work with PCL, PS and others
#                 4.2 - Added Requires parameter for the PrintManagement module
#                 4.1 - Updated comments to be clearer, added more where it was missing
#                 4.0 - Moved into functions and import csv function
#                 3.0 - Alteration of print queue settings (duplex/color/auto config) working
#                 2.0 - Queue and port creation working
#                 1.1 - Added pnputil to add the driver
#                 1.0 - Beginning
#
##########################################################################

#Requires -Modules PrintManagement

# Can only be ran on Windows 8.1+ or Server 2012+
# Older versions do not contain the PrintManagement module which is why the Requires parameter is set
# From what I can tell, this module is just a front end for WMI calls, simplifying the commands




# Before a port can be created, the driver has to be in the driver store
# You can manually install it or use pnputil (built into Windows)
#     pnputil -i -a "path to inf"



# Creating Port
# We use the ip for the name, but they can be independent if desired
#     Add-PrinterPort -Name "portname" -PrinterHostAddress "port_ip"

# Creating Printer
# Can create a different share name, but just using -Shared will use the queue name as the shared name
#    Add-Printer -Name name -DriverName "Driver name" -PortName "portname" -Comment "comment" -Shared -Published


# Can add -ComputerName $Server to everything if you wanted to do it remotely
# Current plan is to do this locally and not remotely
#$Server = "skyprt01"




# This is the import file containing the print queue information
# Change this as required
$File = "\\skynet01\dfrohlick\Home\Stuff\Printers\4th floor.csv"

# Location for the log file
$Log = $env:userprofile + "\desktop\Print_Queue_Creation.txt"


Function SNMP {
# Checks the SNMP registry and updates it if required
# This is only for the HP UPD Auto-Config
    
    $Path = 'HKLM:\SOFTWARE\Hewlett-Packard\SNMP'
    $Value = "Printer_Get"

    Write-Output "Testing for HP SNMP Path" | Tee-Object $Log -Append
    $TestPath = Test-Path $Path

    If ($TestPath -ne $True) {
        # If the key doesn't exit, create it and the entry
        
        Write-Output "HP SNMP Key not found." | Tee-Object $Log -Append
        Write-Output "Creating $Path" | Tee-Object $Log -Append
        New-Item -Path HKLM:\SOFTWARE\Hewlett-Packard -Name SNMP | Out-Null
        
        Write-Output "Creating SNMP entry MystSNMPCommunityName" | Tee-Object $Log -Append
        New-ItemProperty -Path $Path -Name 'MystSNMPCommunityName' -Value $Value | Out-Null
    }

    If ($TestPath -eq $True) {
        # If the key does exist, check the value of the entry
        
        Write-Output "HP SNMP Key found. Checking value" | Tee-Object $Log -Append
        $Val = Get-ItemProperty -Path $Path -Name "MystSNMPCommunityName"
            
        If ($Val.MystSNMPCommunityName -ne "Printer_Get") {
            # If the value doesn't match the correct SNMP value, update it
                
            Write-Output "SNMP Value incorrect. Updating to proper SNMP value" | Tee-Object $Log -Append
            New-ItemProperty -Path $Path -Name 'MystSNMPCommunityName' -Value $Value -Force | Out-Null
        }
        If ($Val.MystSNMPCommunityName -eq "Printer_Get") {
            Write-Output "SNMP Value is correct." | Tee-Object $Log -Append
        }
    }
}

Function CreatePort {
# Creates the printer port
# SNMP 1 turns it on and we add in our custom community string

    # Checks to see if the port exists
    $Exist = Test-Path "HKLM:\SYSTEM\CurrentControlSet\Control\Print\Monitors\Standard TCP/IP Port\Ports\$($Printer.Port)"

    # If the port doesn't exist, create it
    If ($Exist -eq $False) {    
        Write-Output "`r`nAdding Printer Port $($Printer.Name)" | Tee-Object $Log -Append
        Add-PrinterPort -Name $Printer.Port -PrinterHostAddress $Printer.Port -SNMP 1 -SNMPCommunity "Printer_Get"
    }
    # If the port does exist (more than one queue going to the same printer)
    If ($Exist -eq $True) {
        Write-Output "`r`nPrinter Port $($Printer.Name) already exists, adding print queue to existing port" | Tee-Object $Log -Append
    }
}

Function CreatePrinter {
# Creates the Print Queue

    # Adds the print queue
    # It checks the csv file to see if the old queue was published or not and does so accordingly
    Write-Output "Adding Print Queue $($Printer.Name)" | Tee-Object $Log -Append
    If ($Printer.Published -eq "TRUE") {
            Add-Printer -Name $Printer.Name -DriverName $Printer.Driver -PortName $Printer.Port -Comment $Printer.Comment -Location $Printer.Location -Shared #-Published
    }
    If ($Printer.Published -eq "FALSE") {
        Add-Printer -Name $Printer.Name -DriverName $Printer.Driver -PortName $Printer.Port -Comment $Printer.Comment -Location $Printer.Location -Shared
    }

    # Sets the Duplex and Collate mode to what it was previously
    Write-Output "Setting Duplex Mode on $($Printer.Name)" | Tee-Object $Log -Append
    Set-PrintConfiguration -PrinterName $Printer.Name -Duplex $Printer.DuplexingMode
    If ($Printer.Collate -eq "TRUE") {Write-Output "Setting Collate Mode on $Name" | Tee-Object $Log -Append}
    If ($Printer.Collate -eq "TRUE") {Set-PrintConfiguration -PrinterName $Printer.Name -Collate $True}

    # Sets the type of printer (Auto or Color).  Found that sometimes Auto would change a color printer back to b&w
    Write-Output "Setting Printer Type on $($Printer.Name)" | Tee-Object $Log -Append
    Set-PrinterProperty -PrinterName $Printer.Name -PropertyName Config:DynamicRender -Value $Printer.'Config:DynamicRender'
    
    # Setting remaining printer properties
    Write-Output "Setting Printer Properties on $($Printer.Name)" | Tee-Object $Log -Append
    Set-PrinterProperty -PrinterName $Printer.Name -PropertyName Config:AccessoryOutputBins -Value $Printer.'Config:AccessoryOutputBins'
    Set-PrinterProperty -PrinterName $Printer.Name -PropertyName Config:DeviceIsMopier -Value $Printer.'Config:DeviceIsMopier'
    Set-PrinterProperty -PrinterName $Printer.Name -PropertyName Config:DuplexUnit -Value $Printer.'Config:DuplexUnit'
    Set-PrinterProperty -PrinterName $Printer.Name -PropertyName Config:HPInstallableHCO -Value $Printer.'Config:HPInstallableHCO'
    Set-PrinterProperty -PrinterName $Printer.Name -PropertyName Config:HPMOutputBinHCOMap -Value $Printer.'Config:HPMOutputBinHCOMap'
    Set-PrinterProperty -PrinterName $Printer.Name -PropertyName Config:Memory -Value $Printer.'Config:Memory'
    Set-PrinterProperty -PrinterName $Printer.Name -PropertyName Config:PrinterHardDisk -Value $Printer.'Config:PrinterHardDisk'
    Set-PrinterProperty -PrinterName $Printer.Name -PropertyName Config:SecurePrinting -Value $Printer.'Config:SecurePrinting'
    # This doesn't seem to exist on UPD 6.5.0.  Not sure what it did in the first place. Can't find any info on it
    #Set-PrinterProperty -PrinterName $Printer.Name -PropertyName Config:SNPEnableDisable -Value $Printer.'Config:SNPEnableDisable'
    
    #Configures installed trays
    Write-Output "Setting Installed Trays on $($Printer.Name)" | Tee-Object $Log -Append
    Set-PrinterProperty -PrinterName $Printer.Name -PropertyName Config:Tray1_install -Value $Printer.'Config:Tray1_install'
    Set-PrinterProperty -PrinterName $Printer.Name -PropertyName Config:Tray2_install -Value $Printer.'Config:Tray2_install'
    Set-PrinterProperty -PrinterName $Printer.Name -PropertyName Config:Tray3_install -Value $Printer.'Config:Tray3_install'
    Set-PrinterProperty -PrinterName $Printer.Name -PropertyName Config:Tray4_install -Value $Printer.'Config:Tray4_install'
    Set-PrinterProperty -PrinterName $Printer.Name -PropertyName Config:Tray5_install -Value $Printer.'Config:Tray5_install'
    Set-PrinterProperty -PrinterName $Printer.Name -PropertyName Config:Tray6_install -Value $Printer.'Config:Tray6_install'
    Set-PrinterProperty -PrinterName $Printer.Name -PropertyName Config:Tray7_install -Value $Printer.'Config:Tray7_install'
    Set-PrinterProperty -PrinterName $Printer.Name -PropertyName Config:Tray8_install -Value $Printer.'Config:Tray8_install'
    Set-PrinterProperty -PrinterName $Printer.Name -PropertyName Config:Tray9_install -Value $Printer.'Config:Tray9_install'
    Set-PrinterProperty -PrinterName $Printer.Name -PropertyName Config:Tray10_install -Value $Printer.'Config:Tray10_install'
    Set-PrinterProperty -PrinterName $Printer.Name -PropertyName Config:TrayExt1_install -Value $Printer.'Config:TrayExt1_install'
    Set-PrinterProperty -PrinterName $Printer.Name -PropertyName Config:TrayExt2_install -Value $Printer.'Config:TrayExt2_install'
    Set-PrinterProperty -PrinterName $Printer.Name -PropertyName Config:TrayExt3_install -Value $Printer.'Config:TrayExt3_install'
    Set-PrinterProperty -PrinterName $Printer.Name -PropertyName Config:TrayExt4_install -Value $Printer.'Config:TrayExt4_install'
    Set-PrinterProperty -PrinterName $Printer.Name -PropertyName Config:TrayExt5_install -Value $Printer.'Config:TrayExt5_install'
    Set-PrinterProperty -PrinterName $Printer.Name -PropertyName Config:TrayExt6_install -Value $Printer.'Config:TrayExt6_install'

    # This regkey turns off the super annoying HP status popups on the server which are completely unneeded
    # Because HP
    Write-Output "Turning off GUI Status Notifications for $($Printer.Name)" | Tee-Object $Log -Append
    New-ItemProperty -Path HKLM:\SYSTEM\CurrentControlSet\Control\Print\Printers\$($Printer.Name)\PrinterDriverData -Name 'SSNPNotifyEventSettings' -PropertyType  DWORD -Value 0 -Force | Out-Null
    New-ItemProperty -Path HKLM:\SYSTEM\CurrentControlSet\Control\Print\Printers\$($Printer.Name)\PrinterDriverData -Name 'SSNPDriverUISetting' -PropertyType  DWORD -Value 0 -Force | Out-Null

    # This enables Branch Office Direct Printing, if to be used (only works on Server 2012+ and Win 8+)
    # Current plan is to have the printers outside of Head Office/Saskatoon/RSC use this
    If ($Printer.BranchOffice -eq "Yes") {
        Write-Output "Branch Office Cache was enabled on old queue. Enabling on new queue" | Tee-Object $Log -Append
        Set-Printer -Name $Printer.Name -RenderingMode BranchOffice
    }
}

Function Spooler {
# Restores the spooler settings to what they were previously
# However since the new queue may not be shared or published, checks the new queue and alters the value accordingly
# Values for the attributes can be found here http://www.undocprint.org/winspool/registry

# This covers things like printing while spooling or waiting to spool then print

    # Gets the spooler value from the CSV and the new print queue
    $Spool = $Printer.Spooler
    $New_Queue = Get-Printer -Name $Printer.Name

    # If the printer is not shared, subtracts 8 from the spooler value
    If ($New_Queue.Shared -eq $False) {
    Write-Output "$($Printer.Name) is not currently shared.  Removing shared value from spooler attribute" | Tee-Object $Log -Append
    Add-Content $Log "Value of 8 removed"
        $Spool -= 8
    }

    # If the printer is not published and the old queue was published, subtracts 8192 from the spooler value
    If (($New_Queue.Published -eq $False) -and ($Printer.Published -eq "TRUE")) {
    Write-Output "$($Printer.Name) is not currently published. Removing published value from spooler attribute" | Tee-Object $Log -Append
    Add-Content $Log "Value of 8192 removed"
        $Spool -= 8192
    }

    # Forces a new attribute value for the spooler
    New-ItemProperty -Path HKLM:\SYSTEM\CurrentControlSet\Control\Print\Printers\$($Printer.Name) -Name 'Attributes' -PropertyType  DWORD -Value $Spool -Force | Out-Null
    Add-Content $Log "Original spooler value of $($Printer.Spooler) converted to new value of $($Spool)"
}





# Grabs the printers
$Printers = Import-Csv $File    

# Checks server wide SNMP registry setting
SNMP

# For each printer in the CSV, create the port (if required) and then the queue and it's settings
ForEach ($Printer in $Printers) {
    CreatePort
    CreatePrinter
    Spooler
}    

# Forces a restart of the print spooler so the new spooler attributes will be active
Write-Output "`r`nRestarting Print Spooler to enforce spooler attribute values" | Tee-Object $Log -Append
Restart-Service Spooler -Force | Out-File $Log -Append