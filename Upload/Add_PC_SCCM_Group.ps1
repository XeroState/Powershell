####################################################################
# Script Name: Add_PC_SCCM_Group.ps1
# Created On: Dec 21st, 2016
# Author: David Frohlick
# 
# Purpose: Goes through my folder of SCCM notifications
#  
# Script Version: 1.2 - Sorted SCCM groups
#                            1.1 - Added PC input box
#                            1.0 - Initial Script
####################################################################

[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing") 
Add-Type -AssemblyName Microsoft.VisualBasic


Function List {
    
    $DN = 'OU=SCCM,OU=Groups,DC=saskenergy,DC=net'
     Get-ADObject -Filter * -SearchBase $DN -Properties description | Select-Object name | Sort-Object -Property name | Out-File "\\skynet01\dfrohlick\Home\Stuff\Scripts\Powershell\Lists\SCCMList.txt" -Force
    $a = Get-Content "\\skynet01\dfrohlick\Home\Stuff\Scripts\Powershell\Lists\SCCMList.txt" 
    $a | ForEach {$_.TrimEnd()} | Set-Content "\\skynet01\dfrohlick\Home\Stuff\Scripts\Powershell\Lists\SCCM_List.txt" -Force
       
    GUI
}

Function GUI {
    
    $Msg = "What is the name of the PC?"
    $PC= [Microsoft.VisualBasic.Interaction]::InputBox($Msg,"PC Name","PC")
    $PCSAM = $PC + "$"


    $selection = $null
    $DialogResult = $null

    $Font = New-Object System.Drawing.Font("Calibri",12,[System.Drawing.FontStyle]::Regular)

    #Builds Window
    $objForm = New-Object System.Windows.Forms.Form 
    $objForm.Text = "SCCM Group Add Tool"
    $objForm.Size = New-Object System.Drawing.Size(400,400) 
    $objForm.StartPosition = "CenterScreen"

    #Builds OK button
    $OKButton = New-Object System.Windows.Forms.Button
    $OKButton.Location = New-Object System.Drawing.Size(97,325)
    $OKButton.Size = New-Object System.Drawing.Size(100,25)
    $OKButton.Text = "OK"
    $OKButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $objForm.Controls.Add($OKButton)
    $objForm.AcceptButton = $OKButton

    #Builds Cancel button
    $CancelButton = New-Object System.Windows.Forms.Button
    $CancelButton.Location = New-Object System.Drawing.Size(203,325)
    $CancelButton.Size = New-Object System.Drawing.Size(100,25)
    $CancelButton.Text = "Cancel"
    $CancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $objForm.Controls.Add($CancelButton)
    $objForm.CancelButton = $CancelButton

    #Window text
    $objLabel = New-Object System.Windows.Forms.Label
    $objLabel.Location = New-Object System.Drawing.Size(10,20) 
    $objLabel.Size = New-Object System.Drawing.Size(280,20) 
    $objLabel.Text = "Select a script to run:"
    $objForm.Controls.Add($objLabel) 

    #Builds list box
    $objListBox = New-Object System.Windows.Forms.ListBox 
    $objListBox.Font=$Font
    $objListBox.Location = New-Object System.Drawing.Size(5,40) 
    $objListBox.Size = New-Object System.Drawing.Size(373,20) 
    $objListBox.Height = 275

    #Goes through the folder, grabs the sripts and adds them to the list box
    Get-Content "\\skynet01\dfrohlick\Home\Stuff\Scripts\Powershell\Lists\SCCM_List.txt" | ForEach-Object {[void] $objListBox.Items.Add($_)}

    #Adds the listbox to the list, makes it visible and active
    $objForm.Controls.Add($objListBox) 
    $objForm.Topmost = $True

    $DialogResult = $objForm.ShowDialog()

    If ($DialogResult -eq [System.Windows.Forms.DialogResult]::OK) {
        $Selection=($objListBox.SelectedItem)
        Add-ADGroupMember "$Selection" -Members $PCSAM 
        }
    Else {
        [System.Windows.Forms.Application]::Exit($null)
        }

}

List