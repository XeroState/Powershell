####################################################################
# Script Name: Printer_History.ps1
# Created On: Sept 30th, 2016
# Author: David Frohlick
# 
# Purpose: Gets print queue history
#  
# Version: 3.0 - Added calendar picker for dates
#                 2.0 - Fixed the spool to print timer, though if there has been 255 print jobs you'll get more than 1
#                 1.1 - Added passthru on the first out-view to enable looking at time from spool to print
#                 1.0 - Initial Script
####################################################################

Function Info {
#Gathers Info
    
    $Export = Read-Host "Export to CSV? [y or n]"
    $Server = Read-Host "What server are we searching on? [ie skyprt03 or rsback01]"
    $Choice = Read-Host "Searching for [P]rinter or [U]ser or Pa[g]es?"
   
   #If it's by Printer, gets the printer name and how many hours
   #Sets the search parameter field
    If (($Choice -like 'Printer') -or ($Choice -like 'p')) {
        $Value = Read-Host "What is the printer name?"
        $Search = "Param5"
        Search
    }
    #If it's by User, gets the username and how many hours
    #Sets the search paramter field
    If (($Choice -like 'User') -or ($Choice -like 'u')) {
        $Value = Read-Host "What is the username?"
        $Search = "Param3"
        Search
    }
    If (($Choice -like 'Page') -or ($Choice -like 'g')) {
        [int]$Value = Read-Host "Minimum page threshold?"
        $Search = "Param8"
        Search
    }
    #If you suck at typing
    Else {
        Write-Host "Please type printer, user, pages, p, u or g"
        Info
    }
}

Function Search {
#Function to search the event log
    
    #Report array, start time and the event log path
    $Report = @()
    
    $Time = Read-Host "More than 1 day? [y/n]"
    If ($Time -like 'y') {
        Write-Host "Pick Start Date"
        $Start = Calendar
        Write-Host "Pick End Date"
        $End = Calendar
        Write-Host "Generating report"
        $EventLog = Get-WinEvent -ComputerName $Server -FilterHashtable @{LogName='Microsoft-Windows-PrintService/Operational';ID=307;StartTime=$Start;EndTime=$End}
    }
    Else {
        $Start = (Get-Date).AddDays(-1)
        Write-Host "Generating report"
        $EventLog = Get-WinEvent -ComputerName $Server -FilterHashtable @{LogName='Microsoft-Windows-PrintService/Operational';ID=307;StartTime=$Start}
    }

    #For every event log item gathered
    ForEach ($Event in $EventLog) {
        #If the event log item matches the search paramters from the Info function
        $Temp = [xml]$Event.ToXml()
        If ($Search -match "Param8") {
            [int]$TempValue = $Temp.Event.UserData.DocumentPrinted.Param8
            If ($TempValue -ge $Value) {
           
                #GridView/CSV must have data properties, must create a member with a name and value
                $i = New-Object PSObject
        
                #Time of the print job
                $Time = ($Event.TimeCreated).ToShortDateString() + " " + ($Event.TimeCreated).ToShortTimeString()
                $i | Add-Member -MemberType NoteProperty -Name "Time" -Value $Time
                #Who printed it
                $Username = $Temp.Event.UserData.DocumentPrinted.Param3
                $i | Add-Member -MemberType NoteProperty -Name "User" -Value $Username
                #Name of the print job
                $DocName = $Temp.Event.UserData.DocumentPrinted.Param2
                $i | Add-Member -MemberType NoteProperty -Name "Document" -Value $DocName          
                #Where it came from
                $Computer = $Temp.Event.UserData.DocumentPrinted.Param4
                $i | Add-Member -MemberType NoteProperty -Name "Computer" -Value $Computer
                #To what printer
               $Printer = $Temp.Event.UserData.DocumentPrinted.Param5
                $i | Add-Member -MemberType NoteProperty -Name "Printer" -Value $Printer
                #How many pages it was
                $Pages = $Temp.Event.UserData.DocumentPrinted.Param8
                $i | Add-Member -MemberType NoteProperty -Name "Pages" -Value $Pages
                #Size of the job
                $JobSize = $Temp.Event.UserData.DocumentPrinted.Param7
                $JobSize = (($JobSize / 1024)/1024)
                $JobSize = [Math]::Round($JobSize,3)
                $i | Add-Member -MemberType NoteProperty -Name "FileSize" -Value $JobSize
                #How many pages it was
                $JobNumber = $Temp.Event.UserData.DocumentPrinted.Param1
                $i | Add-Member -MemberType NoteProperty -Name "JobNumber" -Value $JobNumber
                
                #Adds fields to the report
                $Report += $i
            }
        }
        Else {
            If ($Temp.Event.UserData.DocumentPrinted.$Search -eq $Value) {
           
                #GridView/CSV must have data properties, must create a member with a name and value
                $i = New-Object PSObject
        
                #Time of the print job
                $Time = ($Event.TimeCreated).ToShortDateString() + " " + ($Event.TimeCreated).ToShortTimeString()
                $i | Add-Member -MemberType NoteProperty -Name "Time" -Value $Time
                #Who printed it
                $Username = $Temp.Event.UserData.DocumentPrinted.Param3
                $i | Add-Member -MemberType NoteProperty -Name "User" -Value $Username
                #Name of the print job
                $DocName = $Temp.Event.UserData.DocumentPrinted.Param2
                $i | Add-Member -MemberType NoteProperty -Name "Document" -Value $DocName          
                #Where it came from
                $Computer = $Temp.Event.UserData.DocumentPrinted.Param4
                $i | Add-Member -MemberType NoteProperty -Name "Computer" -Value $Computer
                #To what printer
               $Printer = $Temp.Event.UserData.DocumentPrinted.Param5
                $i | Add-Member -MemberType NoteProperty -Name "Printer" -Value $Printer
                #How many pages it was
                $Pages = $Temp.Event.UserData.DocumentPrinted.Param8
                $i | Add-Member -MemberType NoteProperty -Name "Pages" -Value $Pages
                #Size of the job
                $JobSize = $Temp.Event.UserData.DocumentPrinted.Param7
                $JobSize = (($JobSize / 1024)/1024)
                $JobSize = [Math]::Round($JobSize,3)
                $i | Add-Member -MemberType NoteProperty -Name "FileSize" -Value $JobSize
                #How many pages it was
                $JobNumber = $Temp.Event.UserData.DocumentPrinted.Param1
                $i | Add-Member -MemberType NoteProperty -Name "JobNumber" -Value $JobNumber
                
                #Adds fields to the report
                $Report += $i
            }
        }
    }

    If ($Report -eq $null) {
        #If nothing found
        Write-Host "No print jobs found on $Server for those search parameters"
        Pause
    }
    Else {
        #Displays the report
        If ($Export -like 'y') {
            $Report | Export-CSV C:\Temp\$Value.csv
            Exit
        }
        Else {
            $Choice = $null
            $Choice = $Report | Out-GridView -PassThru
            Pause
            If ($Choice -ne $null) {
                $Info = $Choice -Split "="
                $TimeOf = $Info[1]
                $TimeOf = $TimeOf.TrimEnd("; User")
                $TimeOf = $TimeOf.Trim()

                $Job = $Info[8]
                $Job = $Job.TrimEnd("}")
                $Job = $Job.Trim()
                JobInfo
            }
            Exit
        }
    }
}

Function JobInfo {
    #Takes Time and Job Number from previous report and reruns it
    #Won't always be singular in results and job numbers overrun every 30ish minutes on SKYPRT03
    #But should be good enough to see length of time of a print job

    $NewStart = (Get-Date $TimeOf) - (New-TimeSpan -Minutes 30)
    $End = (Get-Date $TimeOf) + (New-TimeSpan -Minutes 1)
    $EventLog = Get-WinEvent -ComputerName $Server -FilterHashtable @{LogName='Microsoft-Windows-PrintService/Operational';ID=307;StartTime=$NewStart;EndTime=$End}
    $SecondEventLog = Get-WinEvent -ComputerName $Server -FilterHashtable @{LogName='Microsoft-Windows-PrintService/Operational';ID=800;StartTime=$NewStart;EndTime=$End}
    $Report2 = @()

    ForEach ($Event in $EventLog) {
         #If the event log item matches the search paramters from the Info function
        $Temp = [xml]$Event.ToXml()
        If ($Temp.Event.UserData.DocumentPrinted.Param1 -eq $Job) {
            #GridView/CSV must have data properties, must create a member with a name and value
           $b = New-Object PSObject
            
            #Time of the print job
            $Time = ($Event.TimeCreated).ToShortDateString() + " " + ($Event.TimeCreated).ToLongTimeString()
            $b | Add-Member -MemberType NoteProperty -Name "Time" -Value $Time -Force

            #Name of the print job
            $DocName = $Temp.Event.UserData.DocumentPrinted.Param2
            $b | Add-Member -MemberType NoteProperty -Name "Document" -Value $DocName -Force
        
            #Who printed it
            $Username = $Temp.Event.UserData.DocumentPrinted.Param3
            $b | Add-Member -MemberType NoteProperty -Name "User" -Value $Username -Force

            #Where it came from
            $Computer = $Temp.Event.UserData.DocumentPrinted.Param4
            $b | Add-Member -MemberType NoteProperty -Name "Computer" -Value $Computer -Force
        
            #To what printer
            $Printer = $Temp.Event.UserData.DocumentPrinted.Param5
            $b | Add-Member -MemberType NoteProperty -Name "Printer" -Value $Printer -Force

            #How many pages it was
            $Pages = $Temp.Event.UserData.DocumentPrinted.Param8
            $b | Add-Member -MemberType NoteProperty -Name "Pages" -Value $Pages -Force

            #Size of the job
            $JobSize = $Temp.Event.UserData.DocumentPrinted.Param7
            $JobSize = (($JobSize / 1024)/1024)
            $JobSize = [Math]::Round($JobSize,3)
            $b | Add-Member -MemberType NoteProperty -Name "FileSize" -Value $JobSize

            #Job Number
            $JobNumber = $Temp.Event.UserData.DocumentPrinted.Param1
            $b | Add-Member -MemberType NoteProperty -Name "JobNumber" -Value $JobNumber -Force

            #Adds fields to the report
            $Report2 += $b
        }
    }

    ForEach ($Event in $SecondEventLog) {
    $Temp = [xml]$Event.ToXml()
    If ($Temp.Event.UserData.JobDiag.JobId -eq $Job) {
        #GridView/CSV must have data properties, must create a member with a name and value
           $a = New-Object PSObject
        
            #Time of the print job
            $Time = ($Event.TimeCreated).ToShortDateString() + " " + ($Event.TimeCreated).ToLongTimeString()
            $a | Add-Member -MemberType NoteProperty -Name "Time" -Value $Time
                        
            #Job Number
            $JobNum = $Temp.Event.UserData.JobDiag.JobId
            $a | Add-Member -MemberType NoteProperty -Name "JobNumber" -Value $JobNum
            
            #Adds fields to the report
            $Report2 += $a
        }
    }

    $Report2 | Out-GridView -PassThru
    Exit
}

Function Calendar {
#Pop up calendar to select start and end dates

    #Required assemblies
    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing

    #Creates basic form
    $Form = New-Object Windows.Forms.Form 
    $Form.Text = "Select a Date" 
    $Form.Size = New-Object Drawing.Size @(243,230) 
    $Form.StartPosition = "CenterScreen"

    #Creates calendar
    $Calendar = New-Object System.Windows.Forms.MonthCalendar 
    $Calendar.ShowTodayCircle = $False
    $Calendar.MaxSelectionCount = 1
    $Form.Controls.Add($Calendar) 

    #Creates OK button
    $OKButton = New-Object System.Windows.Forms.Button
    $OKButton.Location = New-Object System.Drawing.Point(38,165)
    $OKButton.Size = New-Object System.Drawing.Size(75,23)
    $OKButton.Text = "OK"
    $OKButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $Form.AcceptButton = $OKButton
    $Form.Controls.Add($OKButton)
    
    #Creates cancel button
    $CancelButton = New-Object System.Windows.Forms.Button
    $CancelButton.Location = New-Object System.Drawing.Point(113,165)
    $CancelButton.Size = New-Object System.Drawing.Size(75,23)
    $CancelButton.Text = "Cancel"
    $CancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $Form.CancelButton = $CancelButton
    $Form.Controls.Add($CancelButton)

    $Form.Topmost = $True

    $Result = $Form.ShowDialog() 

    #If you select OK, sets the date to be the date selected at midnight
    If ($result -eq [System.Windows.Forms.DialogResult]::OK) {
        $Date = $Calendar.SelectionStart
        $Date = $Date.ToLongDateString()
        $Date = $Date + " 12:00:01 AM"
    }
    Return $Date
}

Info