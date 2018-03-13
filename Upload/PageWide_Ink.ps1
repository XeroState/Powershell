####################################################################
# Script Name: PageWide_Ink.ps1
# Created On: Nov 8th, 2016
# Author: David Frohlick
# 
# Purpose: Grabs the current ink level of all PageWide Pro devices
#  
####################################################################

# PageWide Pro Ink Levels
$Logfile = "C:\Users\dfrohlick\Desktop\Ink.txt"

Function LogWrite
{
   Param ([string]$logstring)
   Add-content $Logfile -value $logstring
}

Function SNMP {
        #IP of all PageWide Pro's
        $Printer_Ip = "10.2.3.7", "10.34.0.102", "10.23.0.106", "10.22.0.102", "10.25.0.103", "10.33.0.103", "10.33.0.103", "10.31.0.105", "10.31.0.118", "10.1.3.10", "10.27.0.103", "10.32.0.101"

        $SNMP = New-Object -ComObject olePrn.OleSNMP

        ForEach ($IP in $Printer_IP) {            
            #Reset's each variable
            $Yellow_Max = $null
            $Yellow_Current = $null
            $Magenta_Max = $null
            $Magenta_Current = $null
            $Cyan_Max = $null
            $Cyan_Current = $null
            $Black_Max = $null
            $Black_Current = $null

            #Opens SNMP to the device
            $SNMP.Open($IP,"Printer_Get",2,3000)
            
            #Gets info from SNMP
            $Yellow_Max = $SNMP.Get(".1.3.6.1.2.1.43.11.1.1.8.1.1")
            $Yellow_Current = $SNMP.Get(".1.3.6.1.2.1.43.11.1.1.9.1.1")
            $Magenta_Max = $SNMP.Get(".1.3.6.1.2.1.43.11.1.1.8.1.2")
            $Magenta_Current = $SNMP.Get(".1.3.6.1.2.1.43.11.1.1.9.1.2")
            $Cyan_Max = $SNMP.Get(".1.3.6.1.2.1.43.11.1.1.8.1.3")
            $Cyan_Current = $SNMP.Get(".1.3.6.1.2.1.43.11.1.1.9.1.3")
            $Black_Max = $SNMP.Get(".1.3.6.1.2.1.43.11.1.1.8.1.4")
            $Black_Current = $SNMP.Get(".1.3.6.1.2.1.43.11.1.1.9.1.4")
            $Name = $SNMP.Get(".1.3.6.1.2.1.1.5.0")
            
            #Writes info to the file.  Everything is measured in "15 tenths of millilitre".
            #That should be 1.5ml, but the volumes seem way out of whack for that.   Likely mean 0.15ml 
            LogWrite ("For Printer " + $Name + ", at " + $IP)
            If ($Yellow_Max -eq $null) {
                LogWrite "Device did not respond"
            }
            Else {
                LogWrite "      Yellow:"
                LogWrite ("      There is " + ($Yellow_Current * 0.15) + "ml remaining of " + ($Yellow_Max * 0.15) + "ml. Which is " + ("{0:N2}" -f (($Yellow_Current / $Yellow_Max) * 100)) + "% left.")
                LogWrite "      Magenta:"
                LogWrite ("      There is " + ($Magenta_Current * 0.15) + "ml remaining of " + ($Magenta_Max * 0.15) + "ml. Which is " + ("{0:N2}" -f (($Magenta_Current / $Magenta_Max) * 100)) + "% left.")
                LogWrite "      Cyan:"
                LogWrite ("      There is " + ($Cyan_Current * 0.15) + "ml remaining of " + ($Cyan_Max * 0.15) + "ml. Which is " + ("{0:N2}" -f (($Cyan_Current / $Cyan_Max) * 100)) + "% left.")
                LogWrite "      Black:"
                LogWrite ("      There is " + ($Black_Current * 0.15) + "ml remaining of " + ($Black_Max * 0.15) + "ml. Which is " + ("{0:N2}" -f (($Black_Current / $Black_Max) * 100)) + "% left.")
                LogWrite ""
                }
            }
    }

SNMP