###########################################################
# Script Name: Remote_AutoDesk_AppLocker_Fix.ps1
# Created On: Mar 9, 2016
# Author: David Frohlick
# 
# Purpose: Add keys to allow autodesk to be uninstalled/modified
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


Function Info {
    #Get Info
    $ComputerName = Read-Host "What is the computer name?"

    #Variables
    $global:ErrorMsg = $Null
    $global:SuccessMsg = $Null
    
    $Key1 = "SOFTWARE\Policies\Microsoft\Windows\SrpV2\Script"
    $Key2 = "SOFTWARE\WOW6432Node\Policies\Microsoft\Windows\SrpV2\Script"
    $Key3 = "SYSTEM\CurrentControlSet\Control\Srp\Gp\Msi"
    $Type = [Microsoft.Win32.RegistryHive]::LocalMachine
    
    RegEdit
}

Function RegEdit {
    
    Try {
        #Open Hive on remote machine and then opens the sub key, adds new key
        $RegKey = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey($Type, $ComputerName)
        $TempKey = $RegKey.OpenSubKey($Key1, $True)
        $TempKey.SetValue('EnforcementMode', '0', 'DWORD') | Out-Null
        Success("SOFTWARE\Policies\Microsoft\Windows\SrpV2\Script\EnformentMode Update: Success`n")
    }
    Catch {Success("SrpV2\Script\EnformentMode Update: Failed`n")}

    Try {
        $RegKey = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey($Type, $ComputerName)
        $TempKey = $RegKey.OpenSubKey($Key2, $True)
        $TempKey.SetValue('EnforcementMode', '0', 'DWORD') | Out-Null
        Success("SOFTWARE\WOW6432Node\Policies\Microsoft\Windows\SrpV2\Script\EnforcementMode Update: Success`n")
    }
    Catch {Error("SOFTWARE\WOW6432Node\Policies\Microsoft\Windows\SrpV2\Script\EnforcementMode Update: Failed`n")}

    Try {
        $RegKey = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey($Type, $ComputerName)
        $TempKey = $RegKey.OpenSubKey($Key3, $True)
        $TempKey.SetValue('EnforcementMode', '0', 'DWORD') | Out-Null
        $TempKey.SetValue('EnabledAttributes', '0', 'DWORD') | Out-Null
        Success("SYSTEM\CurrentControlSet\Control\Srp\Gp\Msi\EnforcementMode & EnabledAttribe Update: Success`n")
    }
    Catch {Error("SYSTEM\CurrentControlSet\Control\Srp\Gp\Msi\EnforcementMode & EnabledAttribe Update: Failed`n")}
    
    Output
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

Info
