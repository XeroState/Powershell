###########################################################
# Script Name: UserChoice_Fix.ps1
# Created On: Mar 8, 2016
# Author: David Frohlick
# 
# Purpose: Remote fix of users changing UserChoice keys
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
    #Get Info
    $User = Read-Host "What is the username?"
    $ComputerName = Read-Host "What is the computer name?"

        #Get user SID for registry user
    $UserObj = New-Object System.Security.Principal.NTAccount("SKENERGY", $User)
    $UserSID = $UserObj.Translate([System.Security.Principal.SecurityIdentifier])

    #Variables
    $Key = $UserSID.ToString() + "\Software\Microsoft\Windows\CurrentVersion\Explorer\FileExts"
    $Type = [Microsoft.Win32.RegistryHive]::Users
    $global:SuccessMsg = $Null
    $global:ErrorMsg = $Null
    
    Reg
}

Function Reg {
    #Open Hive on remote machine and then opens the sub key
    $RegKey = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey($Type, $ComputerName)
    $RegKey1 = $RegKey.OpenSubKey($Key, $True)

    #Loop for every key found in subkey
    ForEach($SubKeyName in $RegKey1.GetSubKeyNames()){
        $TempKey = "$key\$SubKeyName"
    
        #Open each subkey found in the subkey
        $SubKey = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey($Type, $ComputerName)
        $SubKey1 = $SubKey.OpenSubKey($TempKey, $True)

        #Loop for each of these subsub keys found, if named UserChoice, delete
        ForEach($SubSub in $SubKey1.GetSubKeyNames()){
            Try {
                If($SubSub -like 'UserChoice'){
                    $SubKey1.DeleteSubKey($SubSub)
                    Success("UserChoice Key found in $SubKey1 and was deleted`n")
                }
            }
            Catch {Error("UserChoice Key found in $SubKey1 but failed to delete`n")}
        }
    }
    Output
}

Info

