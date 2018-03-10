###########################################################
# Script Name: Compare_Software.ps1
# Created On: May 24, 2016
# Author: David Frohlick
# 
# Purpose: Compares installed apps from old pc to new (registry only)
#  
# Script Version: 1.0
#
# Script history
#        1.0 DF - May 24, 2016
###########################################################



Function Info {

    Write-Host "****************WARNING!!!******************" -ForegroundColor Red
    Write-Host "*  THIS ONLY LOOKS FOR ITEMS THAT  *" -ForegroundColor Red
    Write-Host "*       CONTAIN A REGISTRY ENTRY         *" -ForegroundColor Red
    Write-Host "*************************************************" -ForegroundColor Red

    $PC1 = Read-Host -Prompt "`nWhat is the Old PC?"
    $PC2 = Read-Host -Prompt "What is the New PC?"
    $PCs = @($PC1, $PC2)
    
    $Path = "\\skynet01\dfrohlick\home\stuff\scripts\powershell\temp\"
    $Path1 = $Path + $PC1 + ".csv"
    $Path2 = $Path + $PC2 + ".csv"
    
    $UninstallKeys = @("SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Uninstall",
                      "SOFTWARE\\Wow6432Node\\Microsoft\\Windows\\CurrentVersion\\Uninstall")

    Gather
}

Function Gather {

    ForEach ($PC in $PCs) {
        Write-Host "Working on $PC"
        If (Test-Connection -ComputerName $PC -Count 1 -ea 0) {
            ForEach ($UninstallKey in $UninstallKeys) {
                $HKLM = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey('LocalMachine',$PC)
                $UninstallRef = $HKLM.OpenSubKey($UninstallKey)
                $Applications = $UninstallRef.GetSubKeyNames()

                ForEach ($App in $Applications) {
                    $AppRegistryKey = $UninstallKey + "\\" + $App            
                    $AppDetails = $HKLM.OpenSubKey($AppRegistryKey)            
                    $AppDisplayName = $($AppDetails.GetValue("DisplayName"))                                
                
                    If (($AppDisplayName -notlike '*Update for*') -and ($AppDisplayName -notlike '*MUI*') -and ($AppDisplayName -notlike '*Microsoft Office*')) {
                        New-Object -TypeName PSCustomObject -Property @{
                            DisplayName = $AppDisplayName
                        } | Export-CSV -Path "\\skynet01\dfrohlick\home\stuff\scripts\powershell\temp\$PC.csv" -NoTypeInformation -Append
                    }
                }
            }
        }
    }
    Final
}

Function Final {
    $File1 = Import-CSV -Path $Path1
    $File2 = Import-CSV -Path $Path2

    Write-Host "`nMissing from $PC2" -ForegroundColor Green
    Compare-Object $File1 $File2 -Property DisplayName | Where-Object {$_.SideIndicator -eq "<="} | Out-Host
    Remove-Item \\skynet01\dfrohlick\Home\Stuff\Scripts\Powershell\Temp\*.csv
    Pause
}

Info