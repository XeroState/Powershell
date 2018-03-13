####################################################################
# Script Name: SCCM_Group_Report.ps1
# Created On: Dec 6th, 2016
# Author: David Frohlick
# 
# Purpose: Goes through my folder of SCCM notifications
#  
# Script Version: 1.5 - Fixed spelling mistakes
#                            1.4 - Fixed PC output
#                            1.3 - Fixed user and group output
#                            1.2 - Added more to the output
#                            1.1 - Connecting to notes properly
#                            1.0 - Initial Script
####################################################################

Function Check {
#Powershell must be run in 32-bit mode to interact with Notes

    If ($env:PROCESSOR_ARCHITECTURE -ne 'x86') {
        &"C:\Windows\SysWOW64\WindowsPowerShell\v1.0\powershell.exe"  -File "\\skynet01\dfrohlick\Home\Stuff\Scripts\Powershell\SCCM_Group_Report.ps1" | Out-Null
    }
    Else {
        Report
    }
}

Function Report {
#Generates a report based on the emails

    $Mail = "mail\dfrohlic.nsf"
    $Report = @()

    #Creates Notes COM Session, connecting to the mail database
    $NotesSession = New-Object -ComObject Lotus.NotesSession
	$NotesSession.Initialize()
	$NotesDatabase = $NotesSession.GetDatabase( "skydmlprd01",$Mail, 1 )
	
    #Opens View
    $View = $NotesDatabase.GetView('SCCM Changes\Added\Current Month')
    $ViewNav = $View.CreateViewNav()
    $Docs = $ViewNav.GetFirstDocument()
    
    #Loops while it still has an active doc
    While ($Docs -ne $null) {
        $a = New-Object PSObject

        #Grabs the body info and breaks it up at every return
        $Document = $Docs.Document
        $Values = $Document.Items | Select *
        $Data = ($Values.text -split '\n')
        
        #Finds the user
        $Who = $Data -match "Caller user name"
        $Who = $Who.Trim("Caller user name	  ")
        
        #Finds the PC
        $PC = $Data -match "Member Name"
        $PC = $PC.Trim("Member Name	  ")
        $PCName = ($PC -split "=")[1]
        $PCName = $PCName.Trim(",OU")

        #Finds whether it was added or removed
        $Type = $Data -match "Remarks"
        If ($Type -like '*added*') {$Func = "Added"}
        Else {$Func = "Removed"}
        
        #Finds which group
        $Group = $Data -match "Group Name"
        $Group = $Group.Trim("Old Group Name	  -")
        $Group = $Group.Trim("Group Name	  ")
        $Group = $Group.Trim(", ")
        
        #Adds it to the report
        $a | Add-Member -MemberType NoteProperty -Name PC -Value $PCName
        $a | Add-Member -MemberType NoteProperty -Name Group -Value $Group
        $a | Add-Member -MemberType NoteProperty -Name Type -Value $Func
        $a | Add-Member -MemberType NoteProperty -Name Who -Value $Who
        
        $Report += $a
        
        #Gets the next email      
        $Docs = $ViewNav.GetNextDocument($Docs)
    }
    #Outputs
    $Report | Out-GridView -PassThru
}


Check