###########################################################
# Script Name: SCCM_Username_to_Computer.ps1
# Created On: Mar 22, 2016
# Author: David Frohlick
# 
# Purpose: Input values into IE for SCCM reports
#  
###########################################################

Function Info {
    #Gets the info
    $Username = Read-Host -prompt "What username are you looking for?"
    $url = 'http://skycmsqlprd01:8080/Reports/Pages/Report.aspx?ItemPath=%2fConfigMgr_SE1%2f_SEI+-+Custom+Reports%2f_SEI+-+Computers+for+a+Specific+Username&ViewMode=Detail'
    Execute
}

Function Execute {
    #Runs it all
    $ie = New-Object -ComObject InternetExplorer.Application
    $ie.navigate($url)
    $ie.visible=$false
    Start-Sleep -seconds 1
    $User = $ie.document.getelementbyid('ctl32_ctl04_ctl05_txtValue')
    $User.value = $Username
    $Button = $ie.document.getelementbyid('ctl32_ctl04_ctl00')
    $Button.Click()
    $ie.visible=$true
}

Function Test {
    #Checks to see if Powershell has Admin rights
    ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")
}

If (!(Test)) {
    # Create a new process object that starts PowerShell as Admin
    # This is needed to launch IE as an Admin
    Start-Process Powershell.exe -PassThru -Verb Runas "\\skynet01\dfrohlick\Home\Stuff\Scripts\Powershell\SCCM_Username_to_Computer.ps1" | Out-Null
    exit
}
Else {
Info
}
    