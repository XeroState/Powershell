####################################################################
# Script Name: Printer_Report.ps1
# Created On: Sept 30th, 2016
# Author: David Frohlick
# 
# Purpose: Updates Sharepoint Printer Usage Report
#  
# Script Version: 5.2.1 - Fixed spelling mistakes
#                            5.2 - Removed some unneeded leftover code
#                            5.1 - Added 32-bit version check
#                            5.0 - Now includes MFP usage report
#                            4.0 - Added a Note export function
#                            3.0 - Added an automatic update to the report csv
#                            2.2 - Added warning function [removed in 4.0]
#                            2.1 - Got check-in to work
#                            2.0 - Broke it into individual functions in prep for future updates
#                            1.2 - Got updating the excel sheet working
#                            1.1 - Check-Out and Open working
#                            1.0 - Initial Script
####################################################################

Function Check {
#Powershell must be run in 32-bit mode to interact with Notes

    If ($env:PROCESSOR_ARCHITECTURE -ne 'x86') {
        &"C:\Windows\SysWOW64\WindowsPowerShell\v1.0\powershell.exe"  -File "\\skynet01\dfrohlick\Home\Stuff\Scripts\Powershell\Printer_Report.ps1" | Out-Null
    }
    Else {
        Export
    }
}

Function Export{
#Goes through mail file and exports any webjet export in the last 5 days found
#Value can be easily changed by editing line 66	
	
    #Gets the name of the mail file
    $Name = Read-Host "What is your mail file name? (Do not incl .nsf)"
    $Mail = "mail\" + $Name + ".nsf"

    #Creates Notes COM Session, connecting to the mail database
    $NotesSession = New-Object -ComObject Lotus.NotesSession
	$NotesSession.Initialize()
	$NotesDatabase = $NotesSession.GetDatabase( "skydmlprd01",$Mail, 1 )
	
    #Gets the view that is by person.  Just chose one
    $AllViews = $NotesDatabase.Views | Select-Object -ExpandProperty Name
	$DBView = $AllViews | Select-String -Pattern "(By Person)"

	#Grabs the navigation of the view and then grabs the first document
    $View = $NotesDatabase.GetView($DBView)
	$ViewNav = $View.CreateViewNav()
	$Docs = $ViewNav.GetFirstDocument()
         
    #So long as there is a document selected, it'll loop
    While ($Docs -ne $null){
        #Gets the emails info    
        $Document = $Docs.Document
        $Values = $Document.Items | Select *
        
        #Checks to see if it has the word "Jet" in it.  Can't imagine there are too many emails with "Jet" and an attachment
        If ($Values.Text -like "*Jet*") { 
            #Gets the date created and looks to see if it's less than 5 days old
            $DocDate = $Document.Created
			$Date = Get-Date
            $Timespan = New-Timespan -days 5
			If (($Date - $DocDate) -lt $Timespan){
                #If it has an attachment
                If ($Document.HasEmbedded){
				    #For every attachment, export it
                    ForEach ($Item in $Document.Items){
                        If ($Item.type -eq 1084){
						    $Attach = $Document.GetItemValue($Item.Name)
							$Attachment = $Document.GetAttachment($Attach)
                            #Depending on the report, sets the name
                            If ($Values.text -like "*MFP Usage*") {
                                $Save = "C:\temp\mfpreport.csv"
                            }
                            ElseIf($Values.text -like "*Page Usage*") {
                                $Save = "C:\temp\pagereport.csv"
                            }
							$Attachment.ExtractFile($Save)
						}
                    }
                }
			}
        }
        #Grabs the next document
		$Docs = $ViewNav.GetNextDocument($Docs)
    }
    PageUpdate
}

Function PageUpdate {
#Removes first 12 rows from CSV file and makes a new one    

    $Input = Get-Item C:\Temp\pagereport.csv
    
    #Creates Excel Object and sets it to be hidden and hide prompts
    $Excel = New-Object -ComObject Excel.Application
    $Excel.Visible = $False
    $Excel.DisplayAlerts = $False

    #Opens the file and the sheet
    $Workbook = $Excel.Workbooks.Open($Input)
    $WorkSheet = $Workbook.Sheets.Item(1)
    
    #Sets counter and loops 12 times to delete the first 12 rows
    #If you delete row 1, row 2 has now become row 1 which is why it loops to delete row 1
    For ($i=1; $i -le '12'; $i++) {
        [void]$Worksheet.Cells.Item(1,1).EntireRow.Delete()
    }

    #Saves it as Report.csv
    $Worksheet.SaveAs("C:\temp\" + "Page_Report" + ".csv", 6)
    
    #Exits and kills the process
    $Excel.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Excel) | Out-Null
    Remove-Variable Excel | Out-Null

   MFPUpdate
}

Function MFPUpdate {
#Removes first 12 rows from CSV file and makes a new one    

    $Input = Get-Item C:\Temp\mfpreport.csv
    
    #Creates Excel Object and sets it to be hidden and hide prompts
    $Excel = New-Object -ComObject Excel.Application
    $Excel.Visible = $False
    $Excel.DisplayAlerts = $False

    #Opens the file and the sheet
    $Workbook = $Excel.Workbooks.Open($Input)
    $WorkSheet = $Workbook.Sheets.Item(1)
    
    #Sets counter and loops 12 times to delete the first 12 rows
    #If you delete row 1, row 2 has now become row 1 which is why it loops to delete row 1
    For ($i=1; $i -le '12'; $i++) {
        [void]$Worksheet.Cells.Item(1,1).EntireRow.Delete()
    }

    #Saves it as Report.csv
    $Worksheet.SaveAs("C:\temp\" + "MFP_Report" + ".csv", 6)
    
    #Exits and kills the process
    $Excel.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Excel) | Out-Null
    Remove-Variable Excel | Out-Null

   PageReport
}

Function PageReport {
#Function to checkout, open, update and save/checkin document

    #Paths for the files
    $Path = "https://is.saskenergy.net/tierii/Procedures/PrinterPageUsage.xlsx"
    $Input = "C:\temp\page_report.csv"

    #Creates Excel Object and sets it to be hidden and hide prompts
    $Excel = New-Object -ComObject Excel.Application
    $Excel.Visible = $False
    $Excel.DisplayAlerts = $False

    #If it can be checked out, then proceed
    If ($Excel.Workbooks.CanCheckOut($Path)) {
        #Not sure why you have to Open / CheckOut / Open, but it works. May be a more efficient way
        $Workbook = $Excel.Workbooks.Open($Path)
        $Workbook = $Excel.Workbooks.CheckOut($Path)
        $Workbook = $Excel.Workbooks.Open($Path)
        Start-Sleep -s 5

        #Does a value refresh and waits for it to finish
        $Workbook.RefreshAll()
        $Excel.CalculateUntilAsyncQueriesDone()
        Start-Sleep -s 5

        #Gets previous month name
 ##   #Doesn't work for Dec, puts the new year... Will have to figure that out someday
        $Today = Get-Date
        $Year = Get-Date -Format yyy
        $LastMonth = $Today.AddMonths(-1) | Get-Date -Format MMM
        $Title = $LastMonth + " " + $Year

        #Gets the new report and adds a new worksheet to the workbook and renames it
        $Report = Import-Csv -Path $Input
        $WorkSheet = $Workbook.Worksheets.Add()
        $WorkSheet.Name = "$Title"

        #Creates headers for the new worksheet
        $WorkSheet.Cells.Item(1,1) = "Printer"
        $WorkSheet.Cells.Item(1,2) = "IP Address"
        $WorkSheet.Cells.Item(1,3) = "Letter"
        $WorkSheet.Cells.Item(1,4) = "Legal"
        $WorkSheet.Cells.Item(1,5) = "11x17"
        $WorkSheet.Cells.Item(1,6) = "Envelope #10"
        $WorkSheet.Cells.Item(1,7) = "Total"

        #Sets the counter for row
        $i = 2

        #For each line in the report, add the values in to it and increment the row by 1
        ForEach ($printer in $Report) {
            $WorkSheet.Cells.Item($i,1) = $Printer."IP Hostname"
            $WorkSheet.Cells.Item($i,2) = $Printer."IP Address"
            $WorkSheet.Cells.Item($i,3) = $Printer."Letter (8.5x11 in)"
            $WorkSheet.Cells.Item($i,4) = $Printer."Legal (8.5x14 in)"
            $WorkSheet.Cells.Item($i,5) = $Printer."11x17 in"
            $WorkSheet.Cells.Item($i,6) = $Printer."Envelope #10"
            $WorkSheet.Cells.Item($i,7) = $Printer.Total
            $i++
        }
        #Checks the workbook back in and saves it
        $Workbook.CheckInWithVersion()
        $Excel.Quit()
    }
    Else {
        Write-Host "Page Report workbook is already Checked-Out" -ForegroundColor Red
        Write-Host "Verify who has it Checked-Out and have it Checked-In so you can update it" -ForegroundColor Red
    }
    #Clears the variables
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Excel) | Out-Null
    Remove-Variable Excel | Out-Null

    MFPReport
}

Function MFPReport {
#Function to checkout, open, update and save/checkin document

    #Paths for the files
    $Path = "https://is.saskenergy.net/tierii/Procedures/PrinterMFPUsage.xlsx"
    $Input = "C:\temp\mfp_report.csv"

    #Creates Excel Object and sets it to be hidden and hide prompts
    $Excel = New-Object -ComObject Excel.Application
    $Excel.Visible = $False
    $Excel.DisplayAlerts = $False

    #If it can be checked out, then proceed
    If ($Excel.Workbooks.CanCheckOut($Path)) {
        #Not sure why you have to Open / CheckOut / Open, but it works. May be a more efficient way
        $Workbook = $Excel.Workbooks.Open($Path)
        $Workbook = $Excel.Workbooks.CheckOut($Path)
        $Workbook = $Excel.Workbooks.Open($Path)
        Start-Sleep -s 5

        #Does a value refresh and waits for it to finish
        $Workbook.RefreshAll()
        $Excel.CalculateUntilAsyncQueriesDone()
        Start-Sleep -s 5

        #Gets previous month name
##   #Doesn't work for Dec, puts the new year... Will have to figure that out someday
        $Today = Get-Date
        $Year = Get-Date -Format yyy
        $LastMonth = $Today.AddMonths(-1) | Get-Date -Format MMM
        $Title = $LastMonth + " " + $Year

        #Gets the new report and adds a new worksheet to the workbook and renames it
        $Report = Import-Csv -Path $Input
        $WorkSheet = $Workbook.Worksheets.Add()
        $WorkSheet.Name = "$Title"

        #Creates headers for the new worksheet
        $WorkSheet.Cells.Item(1,1) = "Printer"
        $WorkSheet.Cells.Item(1,2) = "IP Address"
        $WorkSheet.Cells.Item(1,3) = "Copy"
        $WorkSheet.Cells.Item(1,4) = "Digital Send"
        $WorkSheet.Cells.Item(1,5) = "Incoming Fax"
        $WorkSheet.Cells.Item(1,6) = "Outgoing Fax"
        $WorkSheet.Cells.Item(1,7) = "Scan Count"
        $WorkSheet.Cells.Item(1,8) = "Total"

        #Sets the counter for row
        $i = 2

        #For each line in the report, add the values in to it and increment the row by 1
        ForEach ($printer in $Report) {
            $WorkSheet.Cells.Item($i,1) = $Printer."IP Hostname"
            $WorkSheet.Cells.Item($i,2) = $Printer."IP Address"
            $WorkSheet.Cells.Item($i,3) = $Printer."Copy"
            $WorkSheet.Cells.Item($i,4) = $Printer."Digital Send"
            $WorkSheet.Cells.Item($i,5) =  $Printer."Incoming Fax"
            $WorkSheet.Cells.Item($i,6) =  $Printer."Outgoing Fax"
            $WorkSheet.Cells.Item($i,7) =  $Printer."Scan Count"
            $WorkSheet.Cells.Item($i,8) =  $Printer."Total"
            $i++
        }
        #Checks the workbook back in and saves it
        $Workbook.CheckInWithVersion()
        $Excel.Quit()
    }
    Else {
        Write-Host "MFP Report workbook is already Checked-Out" -ForegroundColor Red
        Write-Host "Verify who has it Checked-Out and have it Checked-In so you can update it" -ForegroundColor Red
        Pause
    }
    #Clears the variables
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Excel) | Out-Null
    Remove-Variable Excel | Out-Null

    Cleanup
}

Function CleanUp {
#Removes files
    
    Remove-Item C:\temp\page*
    Remove-Item C:\temp\mfp*
}

Check