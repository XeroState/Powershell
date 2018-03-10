#####################################################################
#    Name: Remove_Printer.ps1
#
#    Purpose: Removes GP added printer and all associated reg keys
#             Then does a GPUpdate to re-add it
#####################################################################

Function Info {
    #Gets variable info
    $Printer = Read-Host -Prompt "What is the printer name? (ie. SEPRT03)"
    $Printer = "*" + $Printer + "*"
    $User = Read-Host -Prompt "What is the username? (ie. dfrohlick)"
    $PC = Read-Host -Prompt "What is the PC Name?"

    #Converts username to SID
    $UserObj = New-Object System.Security.Principal.NTAccount("SKENERGY", $User)
    $UserSID = $UserObj.Translate([System.Security.Principal.SecurityIdentifier])

    LocalMachine
}

Function LocalMachine {
    #Function to search HKLM

    #Opens HKLM and the top key, getting all subkey names
    $HKLM = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey('LocalMachine',$PC)
    $LMKey = "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Print"
    $LMSub = $HKLM.OpenSubKey($LMKey, $True)
    $LMSubNames = $LMSub.GetSubKeyNames()

    #Loops through each subkey and then pumps it into the recursive search
    ForEach ($LMName in $LMSubNames){
        $LMTempKey = $LMKey + "\" + $LMName
        LMSearch "$LMTempKey"
    }
    CurrentUser
}


Function CurrentUser {
    #Function to search HKCU

    #Opens HKCU and the top key, getting all subkey names
    $HKCU = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey('Users',$PC)
    $CUKey = $UserSID.ToString() + "\Printers"
    $CUSub = $HKCU.OpenSubKey($CUKey, $True)
    $CUSubNames = $CUSub.GetSubKeyNames()

    #Loops through each subkey and then pumps it into the recursive search
    ForEach ($CUName in $CUSubNames){
        $CUTempKey = $CUKey + "\" + $CUName
        CUSearch "$CUTempKey"
    }
    Service
}

Function LMSearch ($NewKey) {
    #Recursive search function for HKLM

    #Looks at each subfolder, value and valuename inside the top key
    $NewSub = $HKLM.OpenSubKey($NewKey, $True)
    $NewSubNames = $NewSub.GetSubKeyNames()
    $Val = $NewSub.GetValueNames()

    #If it finds the printer it deletes the key
    ForEach ($Value in $Val) {
        If ($Value -like $Printer) {
            $NewSub.DeleteValue($Value)
        }
        $SubVal = $NewSub.GetValue("$Value")
        If ($SubVal -like $Printer) {
            $HKLM.DeleteSubKey($NewKey)
        }
    }

    ForEach ($Item in $NewSubNames){
        If ($Item -like $Printer) {
            $NewSub.DeleteSubKey($Item)
        }
        Else {
            #Continues searching through until all subkeys have been searched
            $NewTempKey = $NewKey + "\" + $Item
            LMSearch "$NewTempKey"
        }
    }
}

Function CUSearch ($NewKey) {
    #Recursive search function for HKLM

    #Looks at each subfolder, value and valuename inside the top key
    $NewSub = $HKCU.OpenSubKey($NewKey, $True)
    $NewSubNames = $NewSub.GetSubKeyNames()
    $Val = $NewSub.GetValueNames()

    #If it finds the printer it deletes the key
    ForEach ($Value in $Val) {
        If ($Value -like $Printer) {
            $NewSub.DeleteValue($Value)
        }
        $SubVal = $NewSub.GetValue("$Value")
        If ($SubVal -like $Printer) {
            $HKCU.DeleteSubKey($NewKey)
        }
    }

    ForEach ($Item in $NewSubNames){
        If ($Item -like $Printer) {
            $NewSub.DeleteSubKey($Item)
        }
        Else {
            #Continues searching through until all subkeys have been searched
            $NewTempKey = $NewKey + "\" + $Item
            CUSearch "$NewTempKey"
        }
    }
}

Function Service {
    #Sets the variable and then stops the service
    Write-Host "Stopping Print Spooler"
    $Service = Get-WMIObject -ComputerName $PC -Class Win32_Service -Filter "Name='Spooler'"
    $Service.StopService() | Out-Null
    Sleep 1
    
    #Resets the variable, then loops until it's state is Stopped
    $Service = Get-WMIObject -ComputerName $PC -Class Win32_Service -Filter "Name='Spooler'"
    While ($Service.State -ne "Stopped") {
        Write-Host "Still Stopping"
        Sleep 2
    }
    Write-Host "Print Spooler is stopped"
    
    #Once stopped, it starts the service again
    Write-Host "Starting Print Spooler"
    $Service.StartService() | Out-Null
    Sleep 1
    
    #Resets the variable, then loops until it's state is Running
    $Service = Get-WMIObject -ComputerName $PC -Class Win32_Service -Filter "Name='Spooler'"
    While ($Service.State -ne "Running") {
        Write-Host "Still Starting"
        Sleep 3
    }
    Write-Host "Print Spooler is running"
    $PCName = "\\" + $PC
    psexec $PCName gpupdate /force
    Pause
}

Function Test {
    #Checks to see if Powershell has Admin rights
    ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")
}


If (!(Test)) {
    # Create a new process object that starts PowerShell as Admin
    # This is needed so that it can run takeown.exe
    Start-Process Powershell.exe -PassThru -Verb Runas "\\skynet01\dfrohlick\Home\Stuff\Scripts\Powershell\Remove_Printer.ps1" | Out-Null
    exit
}
Else {
Info
}