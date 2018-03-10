###########################################################################
# Name: Printer_Queue_Output.ps1
#
# Author: David Frohlick
#
# Purpose: Creates a CSV of the print queues
#
# Version : 4.1 - Added Requires parameter
#                 4.0 - Updated comments to be clearer, added more where it was missing
#                 3.0 - Gets duplex mode and type
#                 2.0 - Got the values out of get-printer exporting
#                 1.1 - Added -Full to the get-printer to get all the details
#                 1.0 - Beginning
#
##########################################################################

#Requires -Modules PrintManagement

# Can only be ran on Windows 10 or Server 2012+
# Older versions do not contain the PrintManagement module which is why the Requires parameter is set

$Server = Read-Host "What is the print server name?"
$Choice = Read-Host -Prompt "Default export is $env:userprofile\desktop\$server printers.csv.  Change? [y/n]"

If ($Choice -eq "y") {$Export = Read-Host "What should the output file be? (Full location path required)"}
If ($Choice -eq "n") {$Export = "$env:userprofile\Desktop\$Server Printers.csv"}


# Gathers a list of all the print queues
# In testing, from both a script and within the console this seems to create duplicates
# Has something to do with there being a 'print' queue and a 'print3d' version. No idea why it is that way
# Either way, sorting unqiue names seems to fix it

# DESPLT also throws a bunch of errors, no idea why.. old crappy driver maybe?
# Will have to add more exclusions for things like Fax/Microsoft to PDF/etc if they are present on the server
$List = Get-Printer -ComputerName $Server -Full | Where-Object {($_.Name -notlike "*XPS*") -and ($_.Name -notlike "*DESPLT*") -and ($_.DeviceType -ne "Print3D")}
$List = $List | Sort-Object Name -Unique
$Report = @()

#Loops for each device
ForEach ($Item in $List) {
    $i = New-Object PSObject

    # Grabs the spooler type.  This is like spooler before printing, spool and printer, etc
    $Path = "SYSTEM\CurrentControlSet\Control\Print\Printers\$($Item.Name)"
    $Hive = [Microsoft.Win32.RegistryHive]::LocalMachine
    $Reg = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey($Hive, $Server)
    $Key = $Reg.OpenSubKey($Path, $True)
    $Spool = $Key.GetValue("Attributes")

    # Grabs the Printer Config and Printer Properties, aside from some xml and wmi info
    $Config = Get-PrintConfiguration -ComputerName $Server  -PrinterName $Item.Name | Select-Object * -ExcludeProperty PrintCapabilitiesXML, PrintTicketXML, CimClass, CimInstanceProperties, CimSystemProperties
    $Property = Get-PrinterProperty -ComputerName $Server  -PrinterName $Item.Name
    
    # Adds Printer Info and Config into array
    $i | Add-Member -MemberType NoteProperty -Name "Name" -Value $Item.Name
    $i | Add-Member -MemberType NoteProperty -Name "Driver" -Value $Item.DriverName
    $i | Add-Member -MemberType NoteProperty -Name "Comment" -Value $Item.Comment
    $i | Add-Member -MemberType NoteProperty -Name "Location" -Value $Item.Location
    $i | Add-Member -MemberType NoteProperty -Name "Port" -Value $Item.PortName
    $i | Add-Member -MemberType NoteProperty -Name "Collate" -Value $Config.Collate
    $i | Add-Member -MemberType NoteProperty -Name "Color" -Value $Config.Color
    $i | Add-Member -MemberType NoteProperty -Name "DuplexingMode" -Value $Config.DuplexingMode
    $i | Add-Member -MemberType NoteProperty -Name "Published" -Value $Item.Published
    #$i | Add-Member -MemberType NoteProperty -Name "RenderingMode" -Value $Item.RenderingMode  # This will grab the Branch Office Direct Printing value, only useful if the print server is 2012+

    # Adds spooler attribute value
    $i | Add-Member -MemberType NoteProperty -Name "Spooler" -Value $Spool

    # Loops through the properties outputting the info into the array
    ForEach ($Thing in $Property) {
        $i | Add-Member -MemberType NoteProperty -Name $Thing.PropertyName -Value $Thing.Value
    }

    # Adds all the info in the object to the array
    $Report += $i

    }

    # Exports the csv
    $Report | Export-CSV $Export