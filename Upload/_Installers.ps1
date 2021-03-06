###########################################################
# Script Name: _Installers.ps1
# Created On: Feb 05, 2016
# Author: David Frohlick
# 
# Purpose: Automate the remote running of batch files
#  
# Script Version: 2.2
#
# Script history
#        1.0 DF - Feb 05, 2016
#                 Created script to grab all variables
#        1.1 DF - Feb 08, 2016
#                 Modified to input encrypted password and decypher
#                 plain text
#        1.2 DF - Feb 22, 2016
#                 Cleaned up psexec command and comments
#        1.3 DF - Feb 23, 2016
#                 Fixed press any key function
#        2.0 DF - Changed the code to pull a list from \Software using a depth check (version 5 required)
#                 Also use Start-Process for psexec which pops up a cmd window and gets rid of the errors
#        2.1 DF - Made it an input box so everything is in a GUI
#        2.2 DF - Fixed the input box gui to stop putting 'False' into the variables
#
###########################################################

#Requires -version 5


Function List {

    Try {
        Get-ChildItem "\\skynet02\public\Software" -Name -Recurse -Depth 1| Where-Object {($_ -like "*.bat" -or "*.cmd") -and ($_ -like "*install*")}  | Sort-Object $_.Name | Out-File "\\skynet01\dfrohlick\Home\Stuff\Scripts\Powershell\Lists\Install.txt" -Force
        Get-Content "\\skynet01\dfrohlick\Home\Stuff\Scripts\Powershell\Lists\install.txt" | Where-Object {($_ -notlike "*_archive*") -and ($_ -notlike "*advantex*")} | Set-Content "\\skynet01\dfrohlick\Home\Stuff\Scripts\Powershell\Lists\Install_List.txt" -Force


        $List = Get-Item \\skynet01\dfrohlick\Home\Stuff\Scripts\PowerShell\Lists\install.txt
        $Days = (New-TimeSpan -Start $List.LastWriteTime).TotalDays
    }
    Catch {
        Write-Host "Generating script list failed"
        Pause
    }   
    
    Info
}

Function Info {
    #Gets the computer name and the user to install it as
    $ComputerName = InputBox "Computer Name" "What is the PC Name?"
    $User = InputBox "Username" "What is your username? [Do not include SKENERGY]"
 
    GUI
}

Function GUI {
    [void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
    [void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing") 

    $selection = $null
    $DialogResult = $null

    #Builds Window
    $objForm = New-Object System.Windows.Forms.Form 
    $objForm.Text = "Installer Selection Tool"
    $objForm.Size = New-Object System.Drawing.Size(400,600) 
    $objForm.StartPosition = "CenterScreen"    

    #Builds OK button
    $OKButton = New-Object System.Windows.Forms.Button
    $OKButton.Location = New-Object System.Drawing.Size(97,525)
    $OKButton.Size = New-Object System.Drawing.Size(100,25)
    $OKButton.Text = "Run this shit"
    $OKButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $objForm.Controls.Add($OKButton)
    $objForm.AcceptButton = $OKButton

    #Builds Cancel button
    $CancelButton = New-Object System.Windows.Forms.Button
    $CancelButton.Location = New-Object System.Drawing.Size(203,525)
    $CancelButton.Size = New-Object System.Drawing.Size(100,25)
    $CancelButton.Text = "Nah I'm good"
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
    $objListBox.Location = New-Object System.Drawing.Size(5,40) 
    $objListBox.Size = New-Object System.Drawing.Size(373,20) 
    $objListBox.Height = 475

    #Grabs the list from the text file
    Get-Content "\\skynet01\dfrohlick\Home\Stuff\Scripts\Powershell\Lists\install_list.txt" | ForEach-Object {[void] $objListBox.Items.Add($_)}

    #Adds the listbox to the list, makes it visible and active
    $objForm.Controls.Add($objListBox) 
    $objForm.Topmost = $True

    $DialogResult = $objForm.ShowDialog()

    If ($DialogResult -eq [System.Windows.Forms.DialogResult]::OK) {
        $Selection=($objListBox.SelectedItem)
        $Installer = "`"\\skynet02\public\software\$Selection`""
        $Command = "-u SKENERGY\" + $User + " -h \\" + $ComputerName + " cmd.exe /c $Installer"
        
        Start-Process psexec $Command
        }
    Else {
        [System.Windows.Forms.Application]::Exit($null)
        }
}

Function InputBox ([string] $Title, [string] $Message) {
    
    [void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing")
    [void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")

    $Form = New-Object System.Windows.Forms.Form
    $Form.Text = "$Title"
    $Form.Size = New-Object System.Drawing.Size(350,150)
    $Form.StartPosition = "CenterScreen"

    $OK_Button = New-Object System.Windows.Forms.Button
    $OK_Button.Location = New-Object System.Drawing.Size(125,80)
    $OK_Button.Size = New-Object System.Drawing.Size(75,23)
    $OK_Button.Text = "OK"
    $OK_Button.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $Form.Controls.Add($OK_Button)
    $Form.AcceptButton = $OK_Button

    $Label = New-Object System.Windows.Forms.Label
    $Label.Location = New-Object System.Drawing.Size(10,20)
    $Label.Size = New-Object System.Drawing.Size(280,20)
    $Label.Text = "$Message"
    $Form.Controls.Add($Label)

    $TextBox = New-Object System.Windows.Forms.TextBox
    $TextBox.Location = New-Object System.Drawing.Size(10,40)
    $TextBox.Size = New-Object System.Drawing.Size(310,20)
    $TextBox.Text = ""
    $Form.Controls.Add($TextBox)

    $Form.Topmost = $True

    $Form.Add_Shown({$Form.Activate(); $TextBox.focus()})
    $TextBox.Focus() | Out-Null
    [void] $Form.ShowDialog()

    $Value = $TextBox.Text
    Return $Value
        
}

List