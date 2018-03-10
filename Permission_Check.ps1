##############################################################################################################
# Based on: AppLocker Bypass Checker (Default Rules) v1.0 
# 
# One of the Default Rules in AppLocker allows everything in the folder C:\Windows to be executed. 
# A normal user shouln't have write permission in that folder, but that is not always the case. 
# This script tries to copy an executable to every folder in Windows and (if the copy succeeds) 
# it will try to execute it. 
# Read more at https://mssec.wordpress.com/2015/10/22/applocker-bypass-checker/ 
# 
# // Tom Aafloen, 2015-10-22 
#
#
# Modified to do all of C:\
# Original script would launch all the apps it could, and leave them running. Then check for running apps at the end
# Depending on how bad permissions are, that can become a massive resource hog
# Script now launches the exe/script when it checks to see if it exists. Logs it, closes it, and moves on to the next one
#
# Uses 800 milisecond pauses. While testing in my VM it behaved oddly without it
# It almost seemed like the script would be at the next line before putty would launch
# Tried with a 300 milisecond pause, which was better, but still occassionally gave odd results
# //DF, Oct-2017
#############################################################################################################

    
    # Creates log file and adds header info the result document
     $Log = "C:\Temp\App_Verification_Results.txt"
     
    Add-Content $Log "`r`n====================================================="
    Add-Content $Log "The following folders have more open permissions:`r`n"

    $ErrorActionPreference = "SilentlyContinue"

    Write-Host "Checking all folders on C:\. This will take several minutes."
    Write-Host "You will likely get Access Denied and Blocked by Group Policy messages."
    Write-Host "This is expected and normal.`n`n"

    # Loop through C:\, trying to copy the exe and the batch file
    # Some folders that allow copying but not executing will throw an Access Denied error 
    ForEach($_ in (Get-ChildItem C:\ -Recurse)) { 
    
        If($_.PSIsContainer) {
         
            # Moves locations, attempts to copy the two files
            Set-Location $_.FullName
            Copy-Item "C:\temp\putty.exe" .\ABCtestfile.exe
            Copy-Item "C:\temp\test.bat" .\scripttest.bat
            $Path = $_.FullName
            
           
            # If the EXE is found, logs that it can write
            If (Test-Path -Path .\ABCtestfile.exe) { 
                Add-Content $Log "Current Path: $Path"
                Add-Content $Log "Writing of EXE's = Success"
                # Attempts to run the EXE. If it can, logs that it can, otherwise logs that it can't
                Start-Process  .\ABCtestfile.exe
                Start-Sleep -Milliseconds 800
                $Check = $null
                $Check = Get-Process ABCtestfile | Where-Object {$_.Path -like "$Path*"}
                If ($Check) {Add-Content $Log "*** Executing of EXE's = Success ***"}
                If (!$Check) {Add-Content $Log "Executing of EXE's = Failed"}
                # Kills the process
                Stop-Process -Name ABCtestfile -Force 
            }

            # If the batch file is found, logs that it can write
            If (Test-Path -Path .\scripttest.bat) { 
                Add-Content $Log "Writing of script's = Success"
                # Attempts to run the batch file
                cmd.exe /c ".\scripttest.bat"
                Start-Sleep -Milliseconds 800
                # Batch file creates a txt file on the desktop. If it's found, the script can be executed
                # Logs that and then removes the txt file
                $TestPath = "C:\users\win10test\desktop\scripttest.txt"
                If (Test-Path -Path $TestPath) {Add-Content $Log "*** Executing of scripts = Success ***`r`n"}
                If (!(Test-Path -Path $TestPath)) {Add-Content $Log "Executing of scripts = Failed`r`n"}
                Remove-Item $TestPath -Force
            } 
        } 
    } 
    
    # Removes all the exe's copied.
    Get-ChildItem C:\ -Include "ABCtestfile.exe", "scripttest.bat", "scripttest.txt", "putty.exe", "test.bat"  -Recurse | Remove-Item
    