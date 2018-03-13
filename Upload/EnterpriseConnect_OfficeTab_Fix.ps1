#####################################################################
#    Name: EnterpriseConnect_OfficeTab_Fix.ps1
#
#    Purpose: Removes regkey which is preventing the tab from appearing
#
#    Version: 1.4 - Added Write-Host and comments
#                   1.3 - Finalized Excel
#                   1.2 - Added Excel
#                   1.1 - Finalized Word
#                   1.0 - Initial Script
#####################################################################

Function Info {
#Gets information and creates variables
    $User = Read-Host "What is the username (not full name)?"
        Try {
        $UserObj = New-Object System.Security.Principal.NTAccount("SKENERGY", $User)
        $UserSID = $UserObj.Translate([System.Security.Principal.SecurityIdentifier])
    }
    Catch {
        Write-Host "Invalid Username"
        Info
    }

    $PC = Read-Host "What is the PC Number?"
    $Computer = "\\" + $PC

    Removal
}

Function Removal {
#Removes the keys

    #Generates variables
    $HKCU = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey('Users',$Computer)
    $Word = $UserSID.ToString() + "\Software\Microsoft\Office\14.0\Word"
    $Excel = $UserSID.ToString() + "\Software\Microsoft\Office\14.0\Excel"

    #Opens Word keys, if Resilicency is found, moves to DisabledItems
    #Deletes everything aside from Default in the DisabledItems
    $WordSubKeys = $HKCU.OpenSubKey($Word, $True)
    $WordSubKeyNames = $WordSubKeys.GetSubKeyNames()    
    ForEach ($Key in $WordSubKeyNames) {
        If ($Key -like 'Resiliency') {
            $Disable = $Word + "\Resiliency\DisabledItems"
            $DisableSubKeys = $HKCU.OpenSubKey($Disable, $True)
            $DisableValues = $DisableSubKeys.GetValueNames()
            ForEach ($Value in $DisableValues) {
                If ($Value -notlike 'Default') {
                    $DisableSubKeys.DeleteValue($Value)
                    Write-Host "Deleted $Value in $Disable"
                    Write-Host `n
                }
           }
        }
    }
     
     #Opens Excel keys, if Resilicency is found, moves to DisabledItems
    #Deletes everything aside from Default in the DisabledItems
    $ExcelSubKeys = $HKCU.OpenSubKey($Excel, $True)
    $ExcelSubKeyNames = $ExcelSubKeys.GetSubKeyNames()    
    ForEach ($Key in $ExcelSubKeyNames) {
        If ($Key -like 'Resiliency') {
            $Disable = $Excel + "\Resiliency\DisabledItems"
            $DisableSubKeys = $HKCU.OpenSubKey($Disable, $True)
            $DisableValues = $DisableSubKeys.GetValueNames()
            ForEach ($Value in $DisableValues) {
                If ($Value -notlike 'Default') {
                    $DisableSubKeys.DeleteValue($Value)
                    Write-Host "Deleted $Value in $Disable"
                    Write-Host `n
                }
           }
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
    Start-Process Powershell.exe -PassThru -Verb Runas "\\skynet01\dfrohlick\Home\Stuff\Scripts\Powershell\EnterpriseConnect_OfficeTab_Fix.ps1" | Out-Null
    exit
}
Else {
Info
}