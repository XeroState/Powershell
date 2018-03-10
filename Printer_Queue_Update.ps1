###########################################################################
# Name: Printer_Queue_Update.ps1
#
# Author: David Frohlick
#
# Purpose: Creates print queues from a CSV file
#
# Version : 4.1 - Added logging to driver update
#                 4.0 - Driver update now contains a settings revert for all the settings
#                 3.2 - Added Requires parameter for the PrintManagement module
#                 3.1 - Updated comments
#                 3.0 - Added publish function
#                 2.0 - Added share function
#                 1.0 - Beginning
#
# TODO:
#     Make old and new driver names a prompt.  Maybe make it a selectable list from installed drivers?
#     That'd be cool.
#
##########################################################################

#Requires -Modules PrintManagement

# Can only be ran on Windows 8.1+ or Server 2012+
# Older versions do not contain the PrintManagement module which is why the Requires parameter is set


Function Update_Driver {
# Update driver used by printer
# Found in testing that using PowerShell/WMI to change the driver resets the print queue to defaults
# Which is a good thing, should prevent any odd issues with like paper type

# Function will go through each queue and any queue with the old driver will be updated


    # Type in the old driver name and the new driver name
    $NewDriver = "HP Universal Printing PCL 6 (v6.2.1)"
    $OldDriver = "HP Universal Printing PCL 6 (v5.8.0)"

    $Log = $env:userprofile + "\desktop\update_driver.txt"

    ForEach ($Printer in $Printers) {
        If ($Printer.DriverName -eq $OldDriver) {
        # If the old driver is used by the printer, runs the update portion
                       
            # Due to the default reset on driver change, it grabs the current basic configuration of the printer for color/duplexing
            Write-Output "`r`nGrabbing $($Printer.Name) config and properties" | Tee-Object $Log -Append
            $Config = Get-PrintConfiguration -ComputerName $Server -PrinterName $Printer.Name
            $Property = Get-PrinterProperty -ComputerName $Server -PrinterName $Printer.Name
            
            # Changes the driver used by the queue, which does a reset on all settings
            Write-Output "Updating $($Printer.Name) to $NewDriver" | Tee-Object $Log -Append
            Set-Printer -ComputerName $Server -Name $Printer.Name -DriverName $NewDriver

            # Sets the Duplex Mode
            Write-Output "Setting Duplex on $($Printer.Name)" | Tee-Object $Log -Append
            Set-PrintConfiguration -ComputerName $Server -PrinterName $Printer.Name -Duplex $Config.DuplexingMode
            
            # Sets the Collate Mode
            If ($Config.Collate -eq "TRUE") {Write-Output "Setting Collate on $($Printer.Name)" | Tee-Object $Log -Append}
            If ($Config.Collate -eq "TRUE") {Set-PrintConfiguration -ComputerName $Server -PrinterName $Printer.Name -Collate $True}
            
            # Sets remaining Printer Properties
            Write-Output "Setting Remaining Printer Properties on $($Printer.Name)" | Tee-Object $Log -Append
            Set-PrinterProperty -ComputerName $Server -PrinterName $Printer.Name -PropertyName Config:DynamicRender -Value $Printer.'Config:DynamicRender'
            Set-PrinterProperty -ComputerName $Server -PrinterName $Printer.Name -PropertyName Config:AccessoryOutputBins -Value $Property.'Config:AccessoryOutputBins'
            Set-PrinterProperty -ComputerName $Server -PrinterName $Printer.Name -PropertyName Config:DeviceIsMopier -Value $Property.'Config:DeviceIsMopier'
            Set-PrinterProperty -ComputerName $Server -PrinterName $Printer.Name -PropertyName Config:DuplexUnit -Value $Property.'Config:DuplexUnit'
            Set-PrinterProperty -ComputerName $Server -PrinterName $Printer.Name -PropertyName Config:HPInstallableHCO -Value $Property.'Config:HPInstallableHCO'
            Set-PrinterProperty -ComputerName $Server -PrinterName $Printer.Name -PropertyName Config:HPMOutputBinHCOMap -Value $Property.'Config:HPMOutputBinHCOMap'
            Set-PrinterProperty -ComputerName $Server -PrinterName $Printer.Name -PropertyName Config:Memory -Value $Property.'Config:Memory'
            Set-PrinterProperty -ComputerName $Server -PrinterName $Printer.Name -PropertyName Config:PrinterHardDisk -Value $Property.'Config:PrinterHardDisk'
            Set-PrinterProperty -ComputerName $Server -PrinterName $Printer.Name -PropertyName Config:SNPEnableDisable -Value $Property.'Config:SNPEnableDisable'
            Set-PrinterProperty -ComputerName $Server -PrinterName $Printer.Name -PropertyName Config:SecurePrinting -Value $Property.'Config:SecurePrinting'

            # Sets Installed Trays
            Write-Output "Setting Installed Trays on $($Printer.Name)" | Tee-Object $Log -Append
            Set-PrinterProperty -ComputerName $Server -PrinterName $Printer.Name -PropertyName Config:Tray1_install -Value $Property.'Config:Tray1_install'
            Set-PrinterProperty -ComputerName $Server -PrinterName $Printer.Name -PropertyName Config:Tray2_install -Value $Property.'Config:Tray2_install'
            Set-PrinterProperty -ComputerName $Server -PrinterName $Printer.Name -PropertyName Config:Tray3_install -Value $Property.'Config:Tray3_install'
            Set-PrinterProperty -ComputerName $Server -PrinterName $Printer.Name -PropertyName Config:Tray4_install -Value $Property.'Config:Tray4_install'
            Set-PrinterProperty -ComputerName $Server -PrinterName $Printer.Name -PropertyName Config:Tray5_install -Value $Property.'Config:Tray5_install'
            Set-PrinterProperty -ComputerName $Server -PrinterName $Printer.Name -PropertyName Config:Tray6_install -Value $Property.'Config:Tray6_install'
            Set-PrinterProperty -ComputerName $Server -PrinterName $Printer.Name -PropertyName Config:Tray7_install -Value $Property.'Config:Tray7_install'
            Set-PrinterProperty -ComputerName $Server -PrinterName $Printer.Name -PropertyName Config:Tray8_install -Value $Property.'Config:Tray8_install'
            Set-PrinterProperty -ComputerName $Server -PrinterName $Printer.Name -PropertyName Config:Tray9_install -Value $Property.'Config:Tray9_install'
            Set-PrinterProperty -ComputerName $Server -PrinterName $Printer.Name -PropertyName Config:Tray10_install -Value $Property.'Config:Tray10_install'
            Set-PrinterProperty -ComputerName $Server -PrinterName $Printer.Name -PropertyName Config:TrayExt1_install -Value $Property.'Config:TrayExt1_install'
            Set-PrinterProperty -ComputerName $Server -PrinterName $Printer.Name -PropertyName Config:TrayExt2_install -Value $Property.'Config:TrayExt2_install'
            Set-PrinterProperty -ComputerName $Server -PrinterName $Printer.Name -PropertyName Config:TrayExt3_install -Value $Property.'Config:TrayExt3_install'
            Set-PrinterProperty -ComputerName $Server -PrinterName $Printer.Name -PropertyName Config:TrayExt4_install -Value $Property.'Config:TrayExt4_install'
            Set-PrinterProperty -ComputerName $Server -PrinterName $Printer.Name -PropertyName Config:TrayExt5_install -Value $Property.'Config:TrayExt5_install'
            Set-PrinterProperty -ComputerName $Server -PrinterName $Printer.Name -PropertyName Config:TrayExt6_install -Value $Property.'Config:TrayExt6_install'
        }
    }

}

Function Share_Printer {
# Shares any print queue not already shared

    $Log = $env:userprofile + "\desktop\Share_Printer.txt"

    ForEach ($Printer in $Printers) {
        If ($Printer.Shared = $False) {
            Write-Output "Sharing $($Printer.Name)" | Tee-Object $Log -Append
            Set-Printer -ComputerName $Server -Name $Printer.Name -Shared $True
        }
    }
}

Function Publish_Printer {
# Publishes any print queue not already published

    $Log = $env:userprofile + "\desktop\Publish_Printer.txt"

    ForEach ($Printer in $Printers) {
        If ($Printer.Published = $False) {
            Write-Output "Publishing $($Printer.Name)" | Tee-Object $Log -Append
            Set-Printer -ComputerName $Server -Name $Printer.Name -Published $True
        }
    }
}

# Gets print server name and what function to run
$Choice = Read-Host "What do you want to run? Update Driver [UD], Share Printers [SP] or Publish Printers [PP]?"
$Server = Read-Host "What is the printer server name?"

# Grabs all the print queues for the functions to use that are not the XPS printer
$Printers = Get-Printer -ComputerName $Server -Full | Where-Object {$_.Name -notlike "*XPS*"}
$Printers = $Printers | Sort-Object Name -Unique

If ($Choice -eq "UD") {Update_Driver}
If ($Choice -eq "SP") {Share_Printer}
If ($Choice -eq "PP") {Publish_Printer}