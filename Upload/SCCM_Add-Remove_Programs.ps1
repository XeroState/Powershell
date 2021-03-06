###########################################################
# Script Name: SCCM_Add-Remove_Programs.ps1
# Created On: Mar 21, 2016
# Author: David Frohlick
# 
# Purpose: Input values into IE for SCCM reports
#  
###########################################################
    
 Function Info {
    #Get information from user
    Add-Type -Assembly "Microsoft.VisualBasic"
    Add-Type -AssemblyName System.Windows.Forms
    
    $App = Read-Host -Prompt "What App are you searching for? Do not include wild cards."
    Write-Host "What type of machine to search?"
    Write-Host  
    [Int]$MenuChoice = 0
    While ($MenuChoice -lt 1 -or $MenuChoice -gt 2) {
        Write-Host "1 - All Physical Workstations"
        Write-Host "2 - All Windows 7 Virtual Machines"
        [Int]$MenuChoice = Read-Host "Enter Option [1-2]"
    }
    
    #Based on choice
    Switch($MenuChoice) {
        1{$Type = "1"}
        2{$Type = "2"}
    }

    #Set values
    $Search = "%" + $App + "%"
    $User = [Environment]::UserName
    If ($MenuChoice -eq "1" ) {$File = $App + " in physical machines.csv"}
    If ($MenuChoice -eq "2" ) {$File = $App + " in virtual machines.csv"}
    Execute
}

Function Execute {    
    #Opens IE
    $url="http://skycmsqlprd01:8080/Reports/Pages/Report.aspx?ItemPath=%2fConfigMgr_SE1%2fSoftware+-+Companies+and+Products%2fComputers+with+specific+software+registered+in+Add+Remove+Programs&ViewMode=Detail"
    $ie = New-Object -ComObject InternetExplorer.Application
    $ie.Navigate($url)
    $ie.visible=$false
    Start-Sleep -seconds 1
   
    #Searches for the value field at the top and puts in search info
    $Application = $ie.document.getelementbyid('ctl32$ctl04$ctl03$txtValue')
    $Application.value = $Search
    
    #Searches for the button to click
    $Button = $ie.document.getelementbyid('ctl32$ctl04$ctl00')
    $Button.Click()
    
    #Waits for page to reload, then searches for dropdown and selects the option based on user input
    Start-Sleep -seconds 2
    $Collection =$ie.document.getelementbyid('ctl32$ctl04$ctl05$ddValue').Focus
    If ($Type -eq "1") {$Collection =$ie.document.getelementbyid('ctl32$ctl04$ctl05$ddValue').SelectedIndex = 1}
    If ($Type -eq "2") {$Collection =$ie.document.getelementbyid('ctl32$ctl04$ctl05$ddValue').SelectedIndex = 2}
    Start-Sleep -seconds 2
    
    #Clicks the button again
    $Button2 = $ie.document.getelementbyid('ctl32$ctl04$ctl00')
    $Button2.Click()
    Start-Sleep -seconds 5
    
    #Searches for the Save button, clicks the drop down then clicks on the CSV button
    $Save1 = $ie.document.getelementbyid("ctl32_ctl05_ctl04_ctl00_ButtonImg")
    $Save1.Click()
    $Link = $ie.document.getelementsbytagname("A") | Where-Object {$_.title -like "*CSV*"}
    $Link.Click()
    
    #Gets the IE process and forces it to be the active window
    #Sends key commands to complete the save
    $IEProc = Get-Process | ? {$_.MainWindowHandle -eq $ie.HWMD}
    [Microsoft.VisualBasic.Interaction]::AppActivate($IEProc.ID)
    Start-Sleep -seconds 2    
    [System.Windows.Forms.SendKeys]::SendWait('{TAB}')
    [System.Windows.Forms.SendKeys]::SendWait('{TAB}')
    [System.Windows.Forms.SendKeys]::SendWait('{TAB}')
    [System.Windows.Forms.SendKeys]::SendWait('{TAB}')
    [System.Windows.Forms.SendKeys]::SendWait('{TAB}')
    [System.Windows.Forms.SendKeys]::SendWait('{TAB}')
    [System.Windows.Forms.SendKeys]::SendWait('{ENTER}')
    Start-Sleep -seconds 1
    
    #Renames file, opens Explorer and closes the IE session
    Rename-Item "C:\Users\$User\Downloads\Computers with specific software registered in Add Remove Programs.csv" "$File"
    Invoke-Item C:\Users\$User\Downloads
    $ie.quit()
}
    
    
Function Test {
    #Checks to see if Powershell has Admin rights
    ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")
}


If (!(Test)) {
    # Create a new process object that starts PowerShell as Admin
    # This is needed to launch IE as an Admin
    Start-Process Powershell.exe -PassThru -Verb Runas "`"\\skynet01\dfrohlick\Home\Stuff\Scripts\Powershell\SCCM_Add-Remove_Programs.ps1`"" | Out-Null
    exit
}
Else {
    Info
}
    
    


    
