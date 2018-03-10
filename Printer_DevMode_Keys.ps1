#####################################################################
#    Name: Printer_DevMode_Keys.ps1
#
#    Purpose: Removes DevMode Keys from Print Queues
#
#    Version: 1.4 - Added Write-Host and comments
#                   1.3 - Finalized HKLM
#                   1.2 - Added HKLM
#                   1.1 - Finalized HKCU
#                   1.0 - Initial Script
#####################################################################

Function Info {
#Gathers information and generates variables

    $User = Read-Host "What is the Username (actual username, not full name)?"
    Try {
        $UserObj = New-Object System.Security.Principal.NTAccount("SKENERGY", $User)
        $UserSID = $UserObj.Translate([System.Security.Principal.SecurityIdentifier])
    }
    Catch {
        Write-Host "Invalid Username"
        Info
    }

    $PC = Read-Host "What is the PC Name?"
    $Printer = Read-Host "What is the print queue name?"
    $Server = Read-Host "What server is the print queue on?"

    $Computer = "\\" + $PC
    $PrintQueue = "\\" + $Server + "\" + $Printer
    $PrintKey = ",," + $Server + "," + $Printer
 
    UserKeys
}

Function UserKeys {
#There are 3 spots that DevMode Keys are can live in the HKCU path
#Goes to all 3 and removes them if they exist

    #Generates variables
    $HKCU = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey('Users',$Computer)
    $Connections = $UserSID.ToString() + "\Printers\Connections\" + $PrintKey
    $DevModePerUser = $UserSID.ToString() + "\Printers\DevModePerUser"
    $DevModes2 = $UserSID.ToString() + "\Printers\DevModes2"
    
    #Location 1. Looks to find DevMode and removes if exists
    $ConnectionsSubKeys = $HKCU.OpenSubKey($Connections, $True)
    $ConnectionsValues = $ConnectionsSubKeys.GetValueNames()
    ForEach ($Value in $ConnectionsValues) {
        If ($Value -like 'DevMode') {
            $ConnectionsSubKeys.DeleteValue($Value)
            Write-Host "DevMode Key removed from $Connections"
            Write-Host `n
        }
    }
  
    #Location 2. Looks to find DevMode and removes if exists
    $DevModePerUserSubKeys = $HKCU.OpenSubKey($DevModePerUser, $True)
    $DevModePerUserValues = $DevModePerUserSubKeys.GetValueNames()
    ForEach ($Value in $DevModePerUserValues) {
        If ($Value -notlike 'Default') {
            $DevModePerUserSubKeys.DeleteValue($Value)
            Write-Host "DevMode Key removed from $DevModePerUser"
            Write-Host `n
        }
    }

    #Location 3. Looks to find DevMode and removes if exists
    $DevModes2SubKeys = $HKCU.OpenSubKey($DevModes2, $True)
    $DevModes2Values = $DevModes2SubKeys.GetValueNames()
    ForEach ($Value in $DevModes2Values) {
        If ($Value -like $PrintQueue) {
            $DevModes2SubKeys.DeleteValue($Value)
            Write-Host "DevMode Key removed from $DevModes2"
            Write-Host `n
        }
    }

    LocalMachineKey
}

Function LocalMachineKey {
#There is 1 location in the HKLM key for DevMode Keys.
#These are the defaults that load when it doesn't have anything else
#If this is broken, no matter what else you do, the print queue will always be broken

    #Generates variables
    $HKLM = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey('LocalMachine',$Computer)
    $Connection = "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Print\Providers\Client Side Rendering Print Provider\" + $UserSID.ToString() + "\Printers\Connections\" + $PrintKey
    
    #Looks to find DefaultDevMode and deletes it
    $ConnectionSubKeys = $HKLM.OpenSubKey($Connection, $True)
    $ConnectionValues = $ConnectionSubKeys.GetValueNames()
    ForEach ($Value in $ConnectionValues) {
        If ($Value -like 'DefaultDevMode') {
            $ConnectionSubKeys.DeleteValue($Value)
            Write-Host "DevMode Key removed from $Connection"
            Write-Host `n
        }
    }
    Write-Host `n
    Pause
}

Function Test {
    #Checks to see if Powershell has Admin rights
    ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")
}


If (!(Test)) {
    # Create a new process object that starts PowerShell as Admin
    # This is needed so that it can remove reg keys
    Start-Process Powershell.exe -PassThru -Verb Runas "\\skynet01\dfrohlick\Home\Stuff\Scripts\Powershell\Printer_DevMode_Keys.ps1" | Out-Null
    exit
}
Else {
Info
}