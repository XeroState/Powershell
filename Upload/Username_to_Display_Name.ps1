###########################################################
# Script Name: Username_to_Display_Name.ps1
# Created On: Mar 22, 2016
# Author: David Frohlick
# 
# Take a list of usernames and convert them to Full Names
#  
###########################################################

#Import-Module ActiveDirectory

Function Info {
    $Loop="1"
    While ($Loop -eq "1") {
        $User = Read-Host -Prompt "Enter Username" | Out-File "\\skynet01\dfrohlick\Home\stuff\scripts\powershell\temp\users.txt" -Force -Append
        Write-Host
        Write-Host "1. Additional Username  |  2. Done" -ForegroundColor Green
        $Loop = Read-Host -Prompt "Enter Option"
        If ($Loop -lt 1 -or $Loop -gt 2) {
            Write-Host "Options are 1 or 2 dummy"
            Write-Host "1. Additional Username  |  2. Done" -ForegroundColor Green
            $Loop = Read-Host -Prompt "Enter Option"
            }
        }
    Collection
}     

Function Collection {
    #Loops for each person in the list and grabs the properities
    $Users = ForEach ($user in $(Get-Content \\skynet01\dfrohlick\Home\stuff\scripts\powershell\temp\users.txt)){
        If (Get-ADUser -Filter {SAMAccountName -eq $User}) {
            Get-ADUser $user -Properties *
            }
        Else {
            Write-Host ($user + " doesn't exist")
            }
    }
    Output
}

Function Output {
    #Outputs the username and displayname
    $users | Select-Object SAMAccountName,DisplayName
    pause
    Remove-Item \\skynet01\dfrohlick\Home\Stuff\Scripts\Powershell\Temp\users.txt
}

Info
