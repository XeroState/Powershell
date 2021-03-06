###########################################################
# Script Name: Print_Driver_Fix.ps1
# Created On: Apr 11, 2016
# Author: David Frohlick
# 
# Purpose: Transfer of PCL6 via BITS for slow connections
#          Script needs to run as Admin to run Takeown.exe
#  
###########################################################


Function Success ($SuccessMsg) {
#Function to create a log of the successes

    $global:SuccessMsg += ($SuccessMsg)
}

Function Error ($ErrorMsg) {
#Function to create a log of the errors

    $global:ErrorMsg += ($ErrorMsg)
}

Function Output {
#Outputs successes and errors

    If ($global:SuccessMsg -ne $Null) {
        Write-Host `n
        Write-Host "$global:SuccessMsg" -ForegroundColor Green
    }
    If ($global:ErrorMsg -ne $Null) {
        Write-Host `n
        Write-Host "$global:ErrorMsg" -ForegroundColor Red
    }
    Pause
}


Function Info {
#Gets info and creates most variables required for the rest of the script    
    
    #Import the module to use BITS
    Import-Module BITSTransfer

    #Gets Info
    $PC = Read-Host -prompt "What is the PC Name?"
    Write-Host `n
    $Version = Read-Host -prompt "Which driver version? `n[1] 5.8 or [2] 6.2? `n[Type 1 or 2 and Press Enter]"
    Write-Host `n

    $PCName = "\\" + $PC
    $global:SuccessMsg = $Null
    $global:ErrorMsg = $Null

    #Creates variables for the remote machine
    $TopFolder = $PCName + "\c$\Windows\System32\DriverStore\FileRepository"
    If ($Version -eq '1') {$DriverFolder = $TopFolder + "\hpcu160u.inf_amd64_neutral_aa0dc619ec5ee3fd"}
    If ($Version -eq '2') {$DriverFolder = $TopFolder + "\hpcu186u.inf_amd64_neutral_6c2bcb2a67636da6"}
    $StupidFolder = $DriverFolder +"\drivers\dot4\AMD64\winxp"

    #Creates variables for the local machine
    If ($Version -eq '1') {$Source = "C:\Windows\System32\DriverStore\FileRepository\hpcu160u.inf_amd64_neutral_aa0dc619ec5ee3fd\*.*"}
    If ($Version -eq '2') {$Source = "C:\Windows\System32\DriverStore\FileRepository\hpcu186u.inf_amd64_neutral_6c2bcb2a67636da6\*.*"}
    If ($Version -eq '1') {$StupidFile = "C:\Windows\System32\DriverStore\FileRepository\hpcu160u.inf_amd64_neutral_aa0dc619ec5ee3fd\drivers\dot4\AMD64\winxp\difxapi.dll"}
    If ($Version -eq '2') {$StupidFile = "C:\Windows\System32\DriverStore\FileRepository\hpcu186u.inf_amd64_neutral_6c2bcb2a67636da6\drivers\dot4\AMD64\winxp\difxapi.dll"}
    
    #Creates variable for the access rules
    $Rule1 = New-Object System.Security.AccessControl.FileSystemAccessRule("BUILTIN\Administrators", 'FullControl','ObjectInherit', 'None', 'Allow')
    $Rule2 = New-Object System.Security.AccessControl.FileSystemAccessRule("BUILTIN\Administrators", 'FullControl','ContainerInherit,ObjectInherit', 'None', 'Allow')

    Reg
}

Function Reg {
#Deletes the reg key for the driver

    $HKLM = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey('LocalMachine',$PC)
    $LMKey = "SYSTEM\CurrentControlSet\Control\Print\Environments\Windows x64\Drivers\Version-3"
    $Subs = $HKLM.OpenSubKey($LMKey, $True)
    $SubNames = $Subs.GetSubKeyNames()

    ForEach ($Sub in $SubNames) {    
        If ($Version -eq '1') {
            Try {
                If ($Sub -eq "HP Universal Printing PCL 6 (v5.8.0)") {
                    $Subs.DeleteSubKey($Sub) | Out-Null                
                    Success("Registry Key Deletion: Success`n")
                }
            }
            Catch {Error("Registry Key Deletion: Failed`n")}
        }
        If ($Version -eq '2') {
            Try {
                If ($Sub -eq "HP Universal Printing PCL 6 (v6.2.1)") {
                    $Subs.DeleteSubKey($Sub) | Out-Null
                    Success("Registry Key Deletion: Success`n")
                }          
            }
            Catch {Error("Registry Key Deletion: Failed`n")}
        }
    }
    Permissions
}



Function Permissions {
#Takes control of the folder and then deletes the files

    
    Try {
        #Adds permission to FileRepository folder
        Takeown.exe /A /F $TopFolder | Out-Null
        
        $ACL = (Get-Item $TopFolder).GetAccessControl('Access')
        $ACL.SetAccessRule($Rule1)
        Set-ACL -Path $TopFolder -AclObject $ACL
        
        Success("FileRepository Permissions: Success`n")
    }
    Catch {Error("FileRepository Permissions: Failed`n")}


    Try {
        #Checks if driver folder is there
        If (Test-Path $DriverFolder) {
                    
            #Take ownership, then add permissions on all the files
            Takeown.exe /A /R /F $DriverFolder | Out-Null
            
            #$ACL2 = Get-ACL -Path $DriverFolder | Out-Null
            #$ACL2.AddAccessRule($Rule) | Out-Null
            #Set-ACL $DriverFolder $ACL2 | Out-Null

            $ACL2 = (Get-Item $DriverFolder).GetAccessControl('Access')
            $ACL2.SetAccessRule($Rule2)
            Set-ACL -Path $DriverFolder -AclObject $ACL2

            
            #Removes all the files and folder
            Remove-Item  -Recurse -Force $DriverFolder| Out-Null
            
            Success("Ownership Taken: Success`nRemoval of old Files: Success`n")
        }
    }
    Catch {Error("Ownership Taken/Removal of Files: Failed`n")} 
    BITS
}


Function BITS {
#Creates the folders and copies the files

    Try {
        #Creates the folder and the folder structure
        New-Item $DriverFolder -Type Directory | Out-Null
        XCopy.exe /T /E $Source $DriverFolder /Y | Out-Null
        
        #Copies the 1 stupid file that is nested
        Copy-Item $StupidFile -destination $StupidFolder | Out-Null
        
        Success("New Directory Structure: Success`nStupid File Copied: Success`n")
    }
    Catch {Error("New Directory Structure/Stupid File Copied: Failed`n")}

    
    Try {
        #Starts BITS Transfer of driver files
        $Job = Start-BITSTransfer -Source "C:\Windows\System32\DriverStore\FileRepository\hpcu160u.inf_amd64_neutral_aa0dc619ec5ee3fd\*.*" -Destination $DriverFolder -Asynchronous

        #While it's copying, display a percentage every 5 seconds
        While( ($Job.JobState.ToString() -eq 'Transferring') -or ($Job.JobState.ToString() -eq 'Connecting') ) {
            Write-Host ("BITS is " + $Job.JobState.ToString())
            $Amount = [Math]::Round(($Job.BytesTransferred/$Job.BytesTotal),2)*100
            Write-Host $Amount "%"
    
            Sleep 3
        }

        #Once completed, finish the job
        Complete-BitsTransfer -BitsJob $Job
        
        Success("BITS File Copy: Success`n")
    }
    Catch {Error("BITS File Copy: Failed`n")}
    Service
}


Function Service {
#Restarts the print spooler

    Try {
        #Sets the variable and then stops the service
        $Service = Get-WMIObject -ComputerName $PC -Class Win32_Service -Filter "Name='Spooler'"
        $Service.StopService() | Out-Null
        Sleep 1
    
        #Resets the variable, then loops until it's state is Stopped
        $Service = Get-WMIObject -ComputerName $PC -Class Win32_Service -Filter "Name='Spooler'"
        While ($Service.State -ne "Stopped") {
            Sleep 2
        }
           
        #Once stopped, it starts the service again
        $Service.StartService() | Out-Null
        Sleep 1
    
        #Resets the variable, then loops until it's state is Running
        $Service = Get-WMIObject -ComputerName $PC -Class Win32_Service -Filter "Name='Spooler'"
        While ($Service.State -ne "Running") {
            Sleep 3
        }
        
        Success("Print Spooler Restart: Success`n")
    }
    Catch {Error("Print Spooler Restart: Failed`n")}

    
    Try {
        
        If ($PCName -like '*tl*') {Write-Host "PC is a truck laptop, no need to update policy" -ForegroundColor Blue}
        Else {
            #Updates computer policy... maybe? Only command I can find is v3 on Win8+
            ($Restart = psexec -h $PCName gpupdate) 2>&1 | Out-Null

            Success("Group Policy Update: Success")
        }
    }
    Catch {Error("Group Policy Update: Failed")}
    Output
}


Function Test {
    #Checks to see if Powershell has Admin rights
    ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")
}




If (!(Test)) {
    # Create a new process object that starts PowerShell as Admin
    # This is needed so that it can run takeown.exe
    Start-Process Powershell.exe -PassThru -Verb Runas "\\skynet01\dfrohlick\Home\Stuff\Scripts\Powershell\Print_Driver_Fix.ps1" | Out-Null
    exit
}
Else {
    Info
}