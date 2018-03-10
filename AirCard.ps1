Set-ExecutionPolicy ByPass -Force

#Grabs the COM Port that the aircard is on
$comPortNumber = Get-WMIObject win32_potsmodem | where-object {($_.DeviceID -like "*250F*") -or ($_.DeviceID -like '*9041*')} | foreach {$_.AttachedTo} 

#Creates a connection to the COM Port
$port = New-Object System.IO.Ports.SerialPort $comPortNumber, "9600", "None", "8", "1"
$port.DtrEnable = "true"

#Opens the serial connection, writes the command and reads the output for the SIM Card number, then closes the connection
$port.Open()
$port.Write("AT+ICCID`r")
$nonsense = $port.ReadLine()
start-sleep -m 50
$ICCID = $port.ReadLine()
$port.Close()

#Modifies the text output to just be the SIM Card number
$ICCID = $ICCID -replace "ICCID: ", ""
$ICCID = $ICCID.Substring(0,$ICCID.Length-1)

#Gets Interface Name
$Interfaces = cmd.exe /c "netsh mbn show interfaces"
$Interface = ($Interfaces -split ':')[5]
$Interface = $Interface.trim()

#Sets the aircard run reg key
$HKLM = [Microsoft.Win32.RegistryKey]::OpenBaseKey('LocalMachine', 'Registry64')
$Path = "SOFTWARE\Microsoft\Windows\CurrentVersion\Run"
$Key = $HKLM.OpenSubKey($Path, $true)
$Key.SetValue("Sasktel","%comspec% /c netsh mbn connect interface=`"$Interface`" connmode=name name=`"SaskTel`"")

#Gets the current script location, sets it as the current location
$Location = Get-Location
Set-Location $Location
$xml = 'mynewprofile.xml'

#Gets the Subscriber ID and modifies the text output to be just the number
$Sub_ID = cmd.exe /c "netsh mbn show ready interface=*"
$Sub_ID = ($Sub_ID -split ':')[9]
$Sub_ID = $Sub_ID.trim()

#Replaces two terms in the template profile with the SIM Card and Subscriber ID numbers
(Get-Content $xml) | ForEach-Object {
    $_ -replace 'subid', "$Sub_ID" `
       -replace 'simid', "$ICCID"
       } | Set-Content 'c:\profile.xml'

#Imports the new profile, connects, then deletes the file
cmd.exe /c "netsh mbn add profile interface=`"$Interface`" name=c:\profile.xml" | Out-Null
#cmd.exe /c "netsh mbn connect profile interface=`"$Interface`" name=Sasktel" | Out-Null