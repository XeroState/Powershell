###########################################################
# Script Name: Remote_Cleanup.ps1
# Created On: Jan 25, 2016
# Author: David Frohlick
# 
# Purpose: Automate the cleanup of "infected" machines to 
##         speed up the scan process
#  
# Script Version: 2.1
#
# Script history
#        1.0 DF - Jan 25, 2016
#                 Copied TG's script to modify
#        1.1 DF - Jan 26, 2016
#                 Changed psexec command to be silent
#        1.2 DF - Jan 27, 2016
#                 Added AppLocker output
#        1.3 DF - Feb 23, 2016
#                 Fixed the press any key function
#        1.4 DF - Mar 11, 2016
#                 Added ErrorActions
#        2.0 DF - Mar 18, 2016
#                 Added functions to open ePO and being scanning profile
#                 Changed AppLocker output to be xml viewer
#        2.1 DF - Mar 21, 2016
#                 Updated script to have a choice for epo/scanning
#
###########################################################

#Requires -version 5


Function Info {
    Add-Type -AssemblyName System.Windows.Forms   
    
    # Get Computer & User & Time Frame Info
    $You = $env:username
    $ComputerName = Read-Host -Prompt 'What is the computer name?'
    $UserName = Read-Host -Prompt 'What is the profile name (usually AD Username)?'
    $Days = Read-Host -Prompt 'How many days back to check the Event Log? [Numbers Only]'
    

    # Create paths, creates a start time and add's a type so the script can type key strokes later
    $RemotePath = "\\" + $ComputerName + "\C$\Users\" + $UserName
    $ProfilePath = "C:\Users\" + $Username
    $Computer = "\\" + $ComputerName    
    $IECachePath = $ProfilePath + "\AppData\Local\Microsoft\Windows\Temporary Internet Files"    
    $CookiesPath = $ProfilePath + "\AppData\Roaming\Microsoft\Windows\Cookies"   
    $TempPath = $ProfilePath + "\AppData\Local\Temp"
    $Time = (Get-Date).AddDays("-" + $Days)    
    
    Clean
}


Function Clean {
    # Checks for correct profile
    If (Test-Path $RemotePath) {
   
        # PSEXEC's to computer and removes temp files  
        Write-Host "Cleaning"$UserName"'s profile, please wait `n" -foregroundcolor "Blue"
        ($Clean = psexec $Computer -e cmd.exe /c rd /S /Q $TempPath "&" rd /S /Q $IECachePath "&" rd /S /Q $CookiesPath) 2>&1 | Out-Null
        Write-Host "Done cleaning user profile!" -foregroundcolor "Green"
    
        # Show Applocker event logs   
        Write-Host `n 
        Write-Host "Checking Applocker Logs" -foregroundcolor "Blue"

        If (Test-Path \\skynet01\dfrohlick\Home\Stuff\Scripts\Powershell\Temp\Cleanup.csv) {Remove-Item \\skynet01\dfrohlick\Home\Stuff\Scripts\Powershell\Temp\Cleanup.csv}
        $EXEEvents = Get-WinEvent -ComputerName $ComputerName -MaxEvents 10 -FilterHashtable @{LogName="Microsoft-Windows-AppLocker/EXE and DLL"; ID=8004; StartTime=$Time} -ErrorAction SilentlyContinue
        $EXEEvents | Select-Object Message, TimeCreated | Export-CSV \\skynet01\dfrohlick\Home\Stuff\Scripts\Powershell\Temp\Cleanup.csv -Append
      
        $MSIEvents = Get-WinEvent -ComputerName $ComputerName -MaxEvents 10 -FilterHashtable @{LogName="Microsoft-Windows-AppLocker/MSI and Script"; ID=8007; StartTime=$Time} -ErrorAction SilentlyContinue  |
        Where-Object {$_.Message -notlike "*USERS.BAT*" -and $_.Message -notlike "*TOP.BAT*" -and $_.Message -notlike "*TS_BROKENSHORTCUTS.PS1*" -and $_.Message -notlike "*TS_UNUSEDDESKTOPICONS.PS1*"}
        $MSIEvents | Select-Object Message, TimeCreated | Export-CSV \\skynet01\dfrohlick\Home\Stuff\Scripts\Powershell\Temp\Cleanup.csv -Append
        
        Output
    }
    Else {
   
        # For if Profile path doesn't exist
        Write-Host $RemotePath" is not accessible."
        If (Test-Connection -ComputerName $ComputerName -Quiet) {
            $UsersPath =  "\\" + $ComputerName + "\C$\Users"
                If (Test-Path -Path $UsersPath) {
                    Write-Host $ComputerName" is accessible, check that you are specifying a profile that exists below"
                    Get-ChildItem -Path $UsersPath -exclude Public
                }
                Else {
                    Write-Host "We can ping "$ComputerName" but are unable to enumerate user profiles"
                }     
        }
        Else {
              Write-Host "Unable to ping $ComputerName"
        }
    Pause
    }
}

Function Output { 
    #Formats the csv and outputs it in a nice xml window    
    Import-CSV -path \\skynet01\dfrohlick\Home\Stuff\Scripts\Powershell\Temp\Cleanup.csv | Sort TimeCreated | Export-Csv \\skynet01\dfrohlick\Home\Stuff\Scripts\Powershell\Temp\Cleanup2.csv -NoTypeInformation
    Import-CSV -path \\skynet01\dfrohlick\Home\Stuff\Scripts\Powershell\Temp\Cleanup2.csv | Out-GridView
    Remove-Item \\skynet01\dfrohlick\Home\Stuff\Scripts\Powershell\Temp\*.csv
    Remove-Item \\skynet01\dfrohlick\Home\Stuff\Scripts\Powershell\Temp\*.xml
    
    Other
}

Function Other {
    #Asks what to do now
    Write-Host "What do you want to do now?"
    Write-Host
    
    [Int]$MenuChoice = 0
    While ($MenuChoice -lt 1 -or $MenuChoice -gt 2) {
        Write-Host "1 - Start scanning"
        Write-Host "2 - Exit"
        [Int]$MenuChoice = Read-Host "Enter Option [1-2]"
    }
    
    #Based on choice, moves on
    Switch($MenuChoice) {
        1{Scan}
        2{Exit}
    }
}      
    
    

Function ePO {
    #Won't work with -adm accounts. Removed from the script
    #Prompts for password
    $MePass = Read-Host -Prompt 'Enter your password' -AsSecureString
    
    #Convert password to plain text
    $NewPass = [Runtime.InteropServices.Marshal]::PtrToStringAuto(
               [Runtime.InteropServices.Marshal]::SecureStringtoBSTR($MePass))
    
    #Opens IE
    $url="https://skyepoprd02:8443/console/orionDashboard.do?id=11&orion.user.security.token=IYKTEhv6HddZd1LT"
    $ie = New-Object -ComObject InternetExplorer.Application
    $ie.Navigate($url)
    $ie.visible=$false
    Start-Sleep -seconds 1
    
    #Checks to see if it's the invalid cert page, if it is, searches for the link to bypass and clicks it
    If ($ie.document.url -Match "invalidcert") {
        $Link = $ie.document.getelementsbytagname("A") | Where-Object {$_.href -like "*skyepoprd02*"}
        $Link.click()
    }
    Start-Sleep -seconds 1
    
    #Finds the username/password fields and pastes in the info
    $user = $ie.document.getelementbyid("j_username")
    $user.value = $you
    $pass = $ie.document.getelementbyid("j_password")
    $pass.value = $Newpass
    $ie.visible=$true
    Start-Sleep -seconds 1
    
    #Presses enter to log in
    [System.Windows.Forms.SendKeys]::SendWait('~')
    
    #Blanks the password variables as I discovered it keeps them even after the script ends
    $Newpass = $null
    $Mepass = $null
    
    #Moves on if the choice was to do so
    If ($MenuChoice -eq "2") {Exit}
    If ($MenuChoice -eq "1") {Scan}
}

Function Scan {       
    #Opens Explorer to the users profile, then moves down the context menu to scan with malware [if it's not running]
    #May have to make custom scripts if our context menus are different, this works for mine
    Invoke-Expression "explorer '/select,$remotepath'"
    
    $bytes = Get-Process mbam -ErrorAction SilentlyContinue
    If (!$bytes) {
        Start-Sleep -seconds 1
        [System.Windows.Forms.SendKeys]::SendWait("+{F10}")
        [System.Windows.Forms.SendKeys]::SendWait('{DOWN}')
        [System.Windows.Forms.SendKeys]::SendWait('{DOWN}')
        [System.Windows.Forms.SendKeys]::SendWait('{DOWN}')
        [System.Windows.Forms.SendKeys]::SendWait('{DOWN}')
        [System.Windows.Forms.SendKeys]::SendWait('{DOWN}')
        [System.Windows.Forms.SendKeys]::SendWait('{DOWN}')
        [System.Windows.Forms.SendKeys]::SendWait('{DOWN}')
        [System.Windows.Forms.SendKeys]::SendWait('{DOWN}')
        [System.Windows.Forms.SendKeys]::SendWait('{DOWN}')
        [System.Windows.Forms.SendKeys]::SendWait('~')
    }
    Else {
        Write-Host "MalwareBytes' is already running and cannot be opened again."
    } 
}

Function Test {
    #Checks to see if Powershell has Admin rights
    #Without going to epO this isn't required
    #([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")
}


Info