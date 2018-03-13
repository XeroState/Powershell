###########################################################
# Script Name: Copy.ps1
# Created On: Aug 30, 2016
# Author: David Frohlick
# 
# Purpose: Copy file/folder to various computers
#  
# Script history
#        1.0 - Created Script
###########################################################

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName Microsoft.VisualBasic

Function Info {
#What are you copying
    
    #Finds out whether it's just a file or a folder
    $Msg = "Are you wanting to copy a File (i) or a Folder (o)?"
    $Choice = [Microsoft.VisualBasic.Interaction]::InputBox($Msg,"What are we copying?")

    #If you can't read it loops
    While($Choice -ne 'File' -And $Choice -ne 'i' -And $Choice -ne 'Folder' -And $Choice -ne 'o') {	
		$Choice = Read-Host -Prompt "Invalid entry, are you copying a File (i) or a Folder (o)?"        
	}
    
    #If it's a file, opens a file dialog box
    If ($Choice -like 'File' -Or $Choice -like 'i') {
        $FileSelect = New-Object System.Windows.Forms.OpenFileDialog
        $FileSelect.Title = “Please select a file”
        $FileSelect.InitialDirectory = “H:\”
        $FileSelect.Filter = “All Files (*.*)|*.*”
        $Result = $FileSelect.ShowDialog()

        If ($Result -eq “OK”) {
            $File = $FileSelect.FileName
            WhereTo
        }
        Else {
            Write-Host “Cancelled by user”
            Pause
        }
    }
    
    #If it's a folder, opens the folder dialog box
    Else {
        $FolderSelect = New-Object System.Windows.Forms.FolderBrowserDialog
        $Result = $FolderSelect.ShowDialog()        
        If ($Result -eq "OK") {
            $Folder = $FolderSelect.SelectedPath
            WhereTo
        }
        Else {
            Write-Host "Cancelled by user"
            Pause
        }
    }
}

Function WhereTo {
#Where are you copying to

    #Finds out what folder it's going to
    $Msg = "Where do you want to copy the file? Do not inclue the PC name or C$ or quotations (ie Users\Public\Public Desktop\)"
    $WhereFolder = [Microsoft.VisualBasic.Interaction]::InputBox($Msg,"What folder location?")

    #Finds out which PCs it's going to
    $Msg = "What PC's do you want to copy to? Surround each with quotations(`") and separated by a comma(,). (ie `"PC12081`",`"PC12041`")"
    $WherePC = [Microsoft.VisualBasic.Interaction]::InputBox($Msg,"What PCs?")

    #Splits pc list
    $PCs = $WherePC.Split('"', [StringSplitOptions]::RemoveEmptyEntries)
    $PCList = $PCs -ne ","
    
    #Forms paths
    ForEach ($Item in $PCList) {
        $PCList2 += "\\" + $Item + "\C$\" + $WhereFolder
        $PCList2 += ","
    }

    #Re-splits it because coding is weird like that
    $PCList3 = $PCList2.Split(',', [StringSplitOptions]::RemoveEmptyEntries)

    #Goes through and copies if the path is available
    ForEach ($PC in $PCList3) {
        If (!(Test-Path $PC)) {
            Write-Host "$Path is not reachable"
        }
        ElseIf ($Choice -like 'File' -Or $Choice -like 'i') {
            Copy-Item $File $PC -Recurse -Force
        } 
        Else {
            Copy-Item $Folder $PC -Recurse -Force
        }
    }
}

Info