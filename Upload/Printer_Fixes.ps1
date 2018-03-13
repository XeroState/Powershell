###########################################################
# Script Name: Printer_Fixes.ps1
# Created On: Feb 17, 2017
# Author: David Frohlick
# 
# Purpose: Combine multiple printer fixes into 1 script
#
#Version : 6.4 - Local Machine keys now loop thru as well
#              - No longer have to type printer/server names
#          6.3 - Made the user name prompt before the choice so you don't have to keep typing it
#          6.2 - Improved devmode loop to not care about printer name
#              - And delete anything found for all installed printers
#          6.1 - Now loops back to choice
#          6.0 - Added Event Log check for recommend fix
#          5.1 - Driver cache rebuild working with Copy (bits not)
#          5.0 - Added driver cache rebuild, not working
#          4.0 - Added Registry check after reg delete
#          3.3 - Added in step to delete local driver cache
#          3.2 - Fixed formatting of code, added test path
#          3.1 - Added comments
#          3.0 - Added bits transfer
#          2.0 - Added devmodekey fix
#          1.2 - Added commenting
#          1.1 - Improved choice function for future fixes
#          1.0 - Initial Script  
###########################################################
Import-Module BITSTransfer

Function Info {
    # Gets the PC name and creates the variables
    $Global:PC = Read-Host "What is the PC Name?"
    $Global:Computer = "\\" + $Global:PC
    $TestPath = $Global:Computer + "\c$\Windows"
    $ErrorActionPreference = 'SilentlyContinue'

    # Tests if the remote path is available
    If (!(Test-Path $TestPath)) {
        Write-Host "PC is offline." -ForegroundColor Red
        Info
    }
    # Gets the username
    $Global:User = Read-Host "What is the username?"
    ErrorCodes
}

Function ErrorCodes {
# Grabs the admin printservice log and parses it for known errors
    
    Write-Host "Grabbing Print Admin Event Logs." -ForegroundColor Yellow
    Write-Host "This may take awhile if it's a slow connection." -ForegroundColor Yellow
    
    # Creates variables and nulls out certain ones
    $Time = (Get-Date).AddDays("-1")
    $CacheEvents = $null
    $DriverEvents = $null

    # Grabs the event log for the past day
    $Events = Get-WinEvent -ComputerName $Global:PC -MaxEvents 20 -FilterHashtable @{LogName="Microsoft-Windows-PrintService/Admin";StartTime=$Time}

    # Looks at each entry and grabs it's EventID and Message
    ForEach ($Event in $Events) {
        $Id = $Event | Select ID
        $Msg = $Event | Select Message
        
        # If it's EventID 808 and has 0x7e error code message, adds message to variable
        If (($Id -match "808") -and ($Msg -like "*error code 0x7e*")) {
            $CacheEvents = "Yes"
        }
        # If it's EventID 319 and mentions no driver, adds message to variable
        If (($Id -match "319") -and ($Msg -like "*driver could not be found*")) {
            $DriverEvents = "Yes"
        }
    }
    
    # If the specific variable contains anything (null to start with) outputs recommended fix 
    If ($CacheEvents) {
        Write-Host "`nPC has EventID 808 with error code 0x7e.  Reccommend using fix #2`n" -ForegroundColor Cyan
    }
    If ($DriverEvents) {
        Write-Host "`nPC has EventID 319 with message driver could not be found.  Recommend fix #1." -ForegroundColor Cyan
        Write-Host "If it still fails, try #2 and finally #4 if it still fails`n" -ForegroundColor Cyan
    }

    Choice
}

Function Choice {
# Gets the choice of what fix to run

    Write-Host "`n`nWhat fix do you want to apply?"
    [Int]$Choice = 0
    
    # Loops until a valid answer is gotten
    While ($Choice -lt 1 -or $Choice -gt 5) {
        Write-Host "1. Delete Main Registry Key" -ForegroundColor Green
        Write-Host "       -Fixes missing printers, slow printing, errors when trying to print" -ForegroundColor Yellow
        Write-Host "2. Rebuild Print Driver Cache folder" -ForegroundColor Green
        Write-Host "       -Used to fix various issues, does involve copying 140MB of files" -ForegroundColor Yellow
        Write-Host "3. Delete DevModeKeys (general 1st attempt fix for everything)" -ForegroundColor Green
        Write-Host "       -Fixes broken printing preferences, slow printing" -ForegroundColor Yellow
        Write-Host "       -Depending on the issue, a reboot is required to resolve after deleting the key" -ForegroundColor Yellow
        Write-Host "4. Re-copy drivers" -ForegroundColor Green
        Write-Host "       -Fixes invalid driver errors, last known fix for slow printing" -ForegroundColor Yellow
        Write-Host "5. Exit`n" -ForegroundColor Green
        [Int]$Choice = Read-Host "Enter Option [1-5]"
    }
    Write-Host "`n"

    # Based on choice, runs a collection of functions
    Switch($Choice) {
        1{Main_Driver_Key;NetSpool;Policy;RegCheck;Choice}
        2{DriverCache;NetSpool;Policy;Choice}
        3{DevModeKey;UserKeys;LocalMachineKey;;NetSpool;Policy;Choice}
        4{CopyDriver;BITS;NetSpool;Policy;Choice}
        5{Exit}
    }    
}

Function Main_Driver_Key {
    # Deletes the main driver registry key.  It forces the PC to verify against the print server all it's print settings
    Write-Host "Searching for Driver Keys, deleting when found" -ForegroundColor Yellow

    $Global:v580 = "No"
    $Global:v621 = "No"
    $Global:RegChk = "Yes"

    # Connects to remote registry and opens up the key
    $HKLM = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey('LocalMachine',$Global:PC)
    $LMKey = "SYSTEM\CurrentControlSet\Control\Print\Environments\Windows x64\Drivers\Version-3"
    $Subs = $HKLM.OpenSubKey($LMKey, $True)
    $SubNames = $Subs.GetSubKeyNames()
    
    # Loops through the keys and deletes the two in-prod drivers we currently use
    ForEach ($Sub in $SubNames) {
        If ($Sub -eq "HP Universal Printing PCL 6 (v5.8.0)") {
            $Subs.DeleteSubKey($Sub) | Out-Null
            Write-Host "Deleted Reg Key for v5.8.0" -ForegroundColor Green
            $Global:v580 = "Yes"
        }
        If ($Sub -eq "HP Universal Printing PCL 6 (v6.2.1)") {
            $Subs.DeleteSubKey($Sub) | Out-Null
            Write-Host "Deleted Reg Key for v6.2.1" -ForegroundColor Green
            $Global:v621 = "Yes"
        }
    }
}

Function DriverCache {
    # Deletes all the files in the current driver folder as this seems to be a common problem
    # There are 8 locked files, but they don't seem to matter
    
    Write-Host "Removing driver cache" -ForegroundColor Yellow
    $Folder = $Global:Computer + "\c$\Windows\System32\spool\drivers\x64\3\"
    Get-ChildItem -Path $Folder -Include *.* -File -Recurse | ForEach {$_.Delete()}
    Write-Host "Finished removing driver cache" -ForegroundColor Green

    Write-Host "Starting file copy" -ForegroundColor Yellow
    Copy "C:\Windows\System32\spool\drivers\x64\3" "$Global:Computer\C$\Windows\System32\spool\drivers\x64" -Recurse
    Write-Host "Finished file copy" -ForegroundColor Green
}

Function DevModeKey {
    # Gets the info required to delete all the devmodekeys for a specific printer

  
    # Attempts to get the SID, if it fails it says so
    Try {
        $UserObj = New-Object System.Security.Principal.NTAccount("SKENERGY", $Global:User)
        $Global:UserSID = $UserObj.Translate([System.Security.Principal.SecurityIdentifier])
    }
    Catch {
        Write-Host "Invalid Username" -ForegroundColor Red
        DevModeKey
    }

}

Function UserKeys {
    # There are 3 spots that DevMode Keys are can live in the HKCU path
    # Goes to all 3 and removes them if they exist
    Write-Host "`nChecking for DevModeKeys, deleting when found" -ForegroundColor Yellow


    # Sets variables
    $HKCU = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey('Users',$Global:Computer)
    $Connections = $Global:UserSID.ToString() + "\Printers\Connections\"
    $DevModePerUser = $Global:UserSID.ToString() + "\Printers\DevModePerUser"
    $DevModes2 = $Global:UserSID.ToString() + "\Printers\DevModes2"

    
    # Location #1. Deletes any devmode keys found in any printer added
    $ConnectionsSubKeys = $HKCU.OpenSubKey($Connections, $True)
    ForEach ($Key in $ConnectionsSubKeys.GetSubKeyNames()) {
        $TempKey = $Connections + $Key
        $SubKey = $HKCU.OpenSubKey($TempKey, $True)
        $SubKeyValues = $SubKey.GetValueNames()
        ForEach ($Value in $SubKeyValues) {
            If ($Value -like 'DevMode') {
                $SubKey.DeleteValue($Value)
                Write-Host "DevMode Key removed from $SubKey" -ForegroundColor Green
            }
        }
    }
    
    # Location 2. Looks to find DevMode and removes if exists
    $DevModePerUserSubKeys = $HKCU.OpenSubKey($DevModePerUser, $True)
    $DevModePerUserValues = $DevModePerUserSubKeys.GetValueNames()
    ForEach ($Value in $DevModePerUserValues) {
        If ($Value -notlike 'Default') {
            $DevModePerUserSubKeys.DeleteValue($Value)
            Write-Host "DevMode Key removed from $DevModePerUser" -ForegroundColor Green
        }
    }

    # Location 3. Looks to find DevMode and removes if exists
    $DevModes2SubKeys = $HKCU.OpenSubKey($DevModes2, $True)
    $DevModes2Values = $DevModes2SubKeys.GetValueNames()
    ForEach ($Value in $DevModes2Values) {
        If ($Value -like $PrintQueue) {
            $DevModes2SubKeys.DeleteValue($Value)
            Write-Host "DevMode Key removed from $DevModes2" -ForegroundColor Green
        }
    }


}

Function LocalMachineKey {
    # There is 1 location in the HKLM key for DevMode Keys.
    # These are the defaults that load when it doesn't have anything else
    # If this is broken, no matter what else you do, the print queue will always be broken

    # Generates variables
    $HKLM = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey('LocalMachine',$Global:Computer)
    $Connection = "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Print\Providers\Client Side Rendering Print Provider\" + $Global:UserSID.ToString() + "\Printers\Connections\"

       
    $ConnectionSubKeys = $HKLM.OpenSubKey($Connection, $True)
    ForEach ($Key in $ConnectionSubKeys.GetSubKeyNames()) {
        $TempKey = $Connection + $Key
        $SubKey = $HKLM.OpenSubKey($TempKey, $True)
        $SubKeyValues = $SubKey.GetValueNames()
        ForEach ($Value in $SubKeyValues) {
            If ($Value -like 'DefaultDevMode') {
                $SubKey.DeleteValue($Value)
                Write-Host "DevMode Key removed from $SubKey" -ForegroundColor Green
            }
        }
    }
}

Function CopyDriver {
    # Gets which version it is to copy, takes ownership and creates variables for the actual file copy

    # Prompts to get the driver version
    [Int]$Version = 0
    While ($Version -lt 1 -or $Version -gt 2) {
        Write-Host "`nWhich driver version do you want to copy?" -ForegroundColor Green
        Write-Host "1. v5.8" -ForegroundColor Yellow
        Write-Host "2. v6.2" -ForegroundColor Yellow
        [Int]$Version = Read-Host "Enter Option [1-2]"
    }
    
    # Creates variables for the remote machine
    $Global:TopFolder = $Global:Computer + "\c$\Windows\System32\DriverStore\FileRepository"
    If ($Version -eq '1') {$Global:DriverFolder = $Global:TopFolder + "\hpcu160u.inf_amd64_neutral_aa0dc619ec5ee3fd"}
    If ($Version -eq '2') {$Global:DriverFolder = $Global:TopFolder + "\hpcu186u.inf_amd64_neutral_6c2bcb2a67636da6"}
    $Global:StupidFolder = $Global:DriverFolder +"\drivers\dot4\AMD64\winxp"

    # Creates variables for the local machine
    If ($Version -eq '1') {$Global:Source = "C:\Windows\System32\DriverStore\FileRepository\hpcu160u.inf_amd64_neutral_aa0dc619ec5ee3fd\*.*"}
    If ($Version -eq '2') {$Global:Source = "C:\Windows\System32\DriverStore\FileRepository\hpcu186u.inf_amd64_neutral_6c2bcb2a67636da6\*.*"}
    If ($Version -eq '1') {$Global:StupidFile = "C:\Windows\System32\DriverStore\FileRepository\hpcu160u.inf_amd64_neutral_aa0dc619ec5ee3fd\drivers\dot4\AMD64\winxp\difxapi.dll"}
    If ($Version -eq '2') {$Global:StupidFile = "C:\Windows\System32\DriverStore\FileRepository\hpcu186u.inf_amd64_neutral_6c2bcb2a67636da6\drivers\dot4\AMD64\winxp\difxapi.dll"}
    
    # Creates variable for the access rules
    $Rule1 = New-Object System.Security.AccessControl.FileSystemAccessRule("BUILTIN\Administrators", 'FullControl','ObjectInherit', 'None', 'Allow')
    $Global:Rule2 = New-Object System.Security.AccessControl.FileSystemAccessRule("BUILTIN\Administrators", 'FullControl','ContainerInherit,ObjectInherit', 'None', 'Allow')       

    # Takes ownership of the file repository folder
    Takeown.exe /A /F $Global:TopFolder | Out-Null
    If ($? -eq 'True') {
        Write-Host "Ownership taken of File Repository" -ForegroundColor Green
    }
    Else {
        Write-Host "Failed to take ownership of driverstore\filerepository, rest of script will fail"
        Pause
    }

    # Gets the current ACL for the file repository
    $ACL = (Get-Item $Global:TopFolder).GetAccessControl('Access')
    # Updates it to the new rule
    $ACL.SetAccessRule($Rule1)
    # Sets the new ACL
    Set-ACL -Path $Global:TopFolder -AclObject $ACL

    # Takes ownership of the actual driver folder
    Takeown.exe /A /R /F $Global:DriverFolder | Out-Null
    If ($? -eq 'True') {
        Write-Host "Ownership taken of Driver folder" -ForegroundColor Green
    }
    Else {
        Write-Host "Failed to take ownership of driver folder, rest of script will fail"
        Pause
    }
    
    # Gets the currnet ACL for the driver folder
    $Global:ACL2 = (Get-Item $Global:DriverFolder).GetAccessControl('Access')
    # Updates it to the new rule
    $Global:ACL2.SetAccessRule($Rule2)
    # Sets the new ACL
    Set-ACL -Path $Global:DriverFolder -AclObject $Global:ACL2

    # Deletes everything in the driver folder
    Remove-Item  -Recurse -Force $Global:DriverFolder| Out-Null
}

Function BITS {
    # Uses BITS to transfer the driver from your PC to theirs    

    # Recreates the driver folder
    New-Item $Global:DriverFolder -Type Directory | Out-Null
    # Sets the ACL to be correct
    Set-ACL -Path $Global:DriverFolder -AclObject $Global:ACL2
    # Copies the directory structure from your PC to theirs
    XCopy.exe /T /E $Global:Source $Global:DriverFolder /Y | Out-Null
        
    # Copies the 1 stupid file that is nested
    Copy-Item $Global:StupidFile -destination $Global:StupidFolder | Out-Null
      
    # Starts BITS Transfer of driver files
    $BitsJob = Start-BITSTransfer -Source $Global:Source -Destination $Global:DriverFolder -Asynchronous
        
    # While it's copying, display a percentage every 5 seconds
    While( ($BitsJob.JobState.ToString() -eq 'Transferring') -or ($BitsJob.JobState.ToString() -eq 'Connecting') ) {
        Write-Host ("BITS is " + $BitsJob.JobState.ToString()) -ForegroundColor Yellow
        $Amount = [Math]::Round(($BitsJob.BytesTransferred/$BitsJob.BytesTotal),2)*100
        Write-Host $Amount "%" -ForegroundColor Yellow
    
        Sleep 3
    }

    # Once completed, finish the job
    Complete-BitsTransfer -BitsJob $BitsJob
        
    Write-Host "Finished BITS Copy" -ForegroundColor Green
}

Function NetSpool {
    # Restarts the remote printer spooler
    Write-Host "`nRestarting Spooler" -ForegroundColor Yellow
   
    # Sets the variable and then stops the service
    $Service = Get-WMIObject -ComputerName $Global:PC -Class Win32_Service -Filter "Name='Spooler'"
    $Service.StopService() | Out-Null
    Sleep 1
    
    # Resets the variable, then loops until it's state is Stopped
    $Service = Get-WMIObject -ComputerName $Global:PC -Class Win32_Service -Filter "Name='Spooler'"
    While ($Service.State -ne "Stopped") {
        Sleep 2
    }
           
    # Once stopped, it starts the service again
    $Service.StartService() | Out-Null
    Sleep 1
    
    # Resets the variable, then loops until it's state is Running
    $Service = Get-WMIObject -ComputerName $Global:PC -Class Win32_Service -Filter "Name='Spooler'"
    While ($Service.State -ne "Running") {
        Sleep 3
    }
    Write-Host "Spooler restarted" -ForegroundColor Green
}

Function Policy {        
    # Does a gpupdate on the remote machine
    
    Start-Sleep -s 10
    
    Write-Host "`nDoing GPUpdate" -ForegroundColor Yellow

    ($Policy = psexec -h $Global:Computer gpupdate) 2>&1 | Out-Null
    Write-Host "Group Policy Finished" -ForegroundColor Green

    If ($Global:RegChk -ne "Yes") {Pause}
}

Function RegCheck {

    If ($Global:v580 -eq "Yes") {$New58 = "No"}
    If ($Global:v621 -eq "Yes") {$New621 = "No"}

    $HKLM = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey('LocalMachine',$Global:PC)
    $LMKey = "SYSTEM\CurrentControlSet\Control\Print\Environments\Windows x64\Drivers\Version-3"
    $Subs = $HKLM.OpenSubKey($LMKey, $True)
    $SubNames = $Subs.GetSubKeyNames()
    
    #Loops through the keys and deletes the two in-prod drivers we currently use
    ForEach ($Sub in $SubNames) {
        If ($Sub -eq "HP Universal Printing PCL 6 (v5.8.0)") {
            If ($Global:58 = "Yes") {
                $New58 = "Yes"
            }
        }
        
        If ($Sub -eq "HP Universal Printing PCL 6 (v6.2.1)") {
            If ($Global:621 = "Yes") {
                $New621 = "Yes"
            }
        }

    }
    
    If (($New58 -eq "No") -or ($New621 -eq "No")) {
        If ($New58 -eq "No") {Copy "\\skynet01\dfrohlick\Home\Stuff\Scripts\Files\5.8.reg" "$Global:Computer\C$"}
        If ($New621 -eq "No") {Copy "\\skynet01\dfrohlick\Home\Stuff\Scripts\Files\6.2.reg" "$Global:Computer\C$"}
    
        ($RegImport = psexec $Global:Computer -e cmd.exe /c reg import c:\5.8.reg "&" reg import c:\6.2.reg) 2>&1 | Out-Null
        Remove-Item "$Global:Computer\C$\5.8.reg" -Force
        Remove-Item "$Global:Computer\C$\6.2.reg" -Force
    
        Write-Host "Reg Keys did not automatically rebuild. Key(s) have been re-added" -ForegroundColor Magenta
        Write-Host "Print spooler will restart again and update policy." -ForegroundColor Magenta
        Write-Host "If Printer issues still exist, try rebuilding the print driver cache (option 2)" -ForegroundColor Magenta
        NetSpool
        Policy
    }
    Pause
}     

Function Test {
    # Checks to see if Powershell has Admin rights
    ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")
}

If (!(Test)) {
    # Create a new process object that starts PowerShell as Admin
    Start-Process Powershell.exe -PassThru -Verb Runas "\\skynet01\dfrohlick\Home\Stuff\Scripts\Powershell\Printer_Fixes.ps1" | Out-Null
    exit
}
Else {
    Info
}