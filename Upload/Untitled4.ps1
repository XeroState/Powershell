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
# Also does the start process while checking for the existence of the file. Otherwise it was just chewing memory

    # Adds header info the result document
    Add-Content $Log "`r`n====================================================="
    Add-Content $Log "The following have more open permissions:`r`n"

    $User = "SKENERGY\win10test"
    $Pass = "CocaCola46"
    $PSS = ConvertTo-SecureString $Pass -AsPlainText -Force
    $Cred = New-Object System.Management.Automation.PSCredential $User,$PSS
    $Job = Start-Job -ScriptBlock {
    $ErrorActionPreference = "silentlycontinue"
 $Log = "C:\Users\win10test\desktop\App_Verification_Results.txt"
 New-Item $Log -ItemType File | Out-Null
    # Loop through C:\, try to copy executable and - if successful - try to execute it. 
    # Some folders that allow copying but not executing will throw an Access Denied error 
    ForEach($_ in (Get-ChildItem C:\ -Recurse)) { 
    
        If($_.PSIsContainer) {
         
            Set-Location $_.FullName
            Copy-Item "C:\temp\putty.exe" .\ABCtestfile.exe
            Copy-Item "C:\temp\test.vbs" .\scripttest.vbs
            $Path = (Get-Item -Path ".\" -Verbose).FullName
           
            # If the EXE is found, logs that it can write
            If (Test-Path -Path .\ABCtestfile.exe) { 
                Add-Content $Log "`r`nCurrent Path: $Path"
                Add-Content $Log "Writing of EXE's = Success"
                # Attempts to run the EXE. If it can, logs that it can
                Start-Process  .\ABCtestfile.exe -WindowStyle Minimized
                Sleep -Milliseconds 100
                $Check = $null
                $Check = Get-Process ABCtestfile | Where-Object {$_.Path -match "$Path\ABCtestfile.exe"}
                If ($Check) {Add-Content $Log "Executing of EXE's = Success"}
                If (!$Check) {Add-Content $Log "Executing of EXE's = Failed"}
                # Kills the process
                Stop-Process -Name ABCtestfile -Force 
            }

            If (Test-Path -Path .\scripttest.vbs) { 
                Add-Content $Log "Writing of script's = Success"
                # Attempts to run the EXE. If it can, logs that it can
                & wscript.exe ".\scripttest.vbs" -WindowStyle Minimized
                $Check2 = $null
                Sleep -Milliseconds 200
                $Check = Get-Process wscript
                If ($Check2) {Add-Content $Log "Executing of scripts = Success"}
                If (!$Check2) {Add-Content $Log "Executing of scripts = Failed"}
                # Kills the process
                Stop-Process -Name wscript -Force 
            } 
        } 
    } 
        } -Credential $Cred
    
    # Removes all the exe's copied.
    Get-ChildItem C:\ -Include "ABCtestfile.exe", "scripttest.vbs", "scripttest.txt" -Recurse | Remove-Item
    Set-Location C:\