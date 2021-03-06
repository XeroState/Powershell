###########################################################
# Script Name: Net-Send.ps1
# Created On: April 7, 2016
# Author: David Frohlick
# 
# Purpose: Mainly to send messages to Kyle for reasons
#  
###########################################################

Function Send {
    $PC = Read-Host "Enter computer name "
    $MSG = Read-Host "Enter your message "
    Invoke-WmiMethod -Path Win32_Process -Name Create -ArgumentList "msg * $MSG" -ComputerName $PC | Out-Null

    $Choice= Read-Host "`n[1] Another Message [2] Exit"
    If ($Choice -eq '1') {Send}
    Else {Pause}
}

Function Kyle {
    $PC = "PC13050"
    $MSG = Read-Host "Enter your message "
    Invoke-WmiMethod -Path Win32_Process -Name Create -ArgumentList "msg * $MSG" -ComputerName $PC | Out-Null

    $Choice= Read-Host "`n[1] Another Message [2] Exit"
    If ($Choice -eq '1') {Kyle}
    Else {Pause}
}

$Who = Read-Host "Is this to Kyle?"
If ($Who -like 'y') {Kyle}
Else {Send}