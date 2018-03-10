Set-ExecutionPolicy ByPass -Force

#Grabs the COM Port that the aircard is on
$comPortNumber = Get-WMIObject win32_potsmodem | where-object {($_.DeviceID -like "*250F*") -or ($_.DeviceID -like '*9041*')} | foreach {$_.AttachedTo} 

#Creates a connection to the COM Port
$port = New-Object System.IO.Ports.SerialPort $comPortNumber, "9600", "None", "8", "1"
$port.DtrEnable = "true"

$port.Open()
$port.Write("AT!PRIID?`r")
$1 = $port.readline()
$2 = $port.readline()
$3 = $port.readline()
$4 = $port.readline()
$5 = $port.readline()
$6 = $port.readline()
$7 = $port.readline()

$Out = "$1`n$2`n$3`n$4`n$5`n$6`n$7"
$Out | Out-File "C:\Temp\Temp\PRI.txt" -force