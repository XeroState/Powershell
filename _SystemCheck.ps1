###########################################################
# Script Name: _SystemCheck.ps1
# Created On: May 5, 2016
# Author: David Frohlick
# 
# Purpose: 
#  
###########################################################


Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
$RegularFont = New-Object System.Drawing.Font("Calibri",10,[System.Drawing.FontStyle]::Regular)
$HeaderFont = New-Object System.Drawing.Font("Calibri",10,[System.Drawing.FontStyle]::Bold)
$TitleFont = New-Object System.Drawing.Font("Calibri",18,[System.Drawing.FontStyle]::Bold)
$LabelFont = New-Object System.Drawing.Font("Calibri",12,[System.Drawing.FontStyle]::Bold)



Function Window {


    #Builds Main Form Window
    $Window = New-Object System.Windows.Forms.Form 
    $Window.Text = "System Check Tool"
    $Window.Size = New-Object System.Drawing.Size(800,800) 
    $Window.StartPosition = "CenterScreen"
    $Window.AutoSize = $True

    #Makes Title
    $Title = New-Object System.Windows.Forms.Label
    $Title.Location = New-Object System.Drawing.Size(10,20) 
    $Title.Size = New-Object System.Drawing.Size(280,30) 
    $Title.Text = "System Check Tool"
    $Title.Font = $LabelFont

    #Makes Get Info Button
    $GetInfo = New-Object System.Windows.Forms.Button
    $GetInfo.Location = New-Object System.Drawing.Size(115,70)
    $GetInfo.Size = New-Object System.Drawing.Size(75,25)
    $GetInfo.Text = "Get Info"
    $GetInfo.Font = $LabelFont
    $GetInfo.Add_Click({Info})
    $Window.AcceptButton = $GetInfo
           
    #Makes Input Box
    $TextBox = New-Object System.Windows.Forms.TextBox
    $TextBox.Location = New-Object System.Drawing.Size(10,70)
    $TextBox.Size = New-Object System.Drawing.Size(100,20)
    $TextBox.Text = ""
    $TextBox.Font = $RegularFont

    #Makes Update Box
    $Update = New-Object System.Windows.Forms.Label
    $Update.Location = New-Object System.Drawing.Size(200,75)
    $Update.Size = New-Object System.Drawing.Size(250,20)
    $Update.Font = $RegularFont
    $Update.Visible = $False

    #Makes User Info Title
    $User_Title = New-Object System.Windows.Forms.Label
    $User_Title.Location = New-Object System.Drawing.Size(10,110)
    $User_Title.Size = New-Object System.Drawing.Size(150,20)
    $User_Title.Text = "User Information:"
    $User_Title.Font = $LabelFont

    #Makes Network Info Title
    $Network_Title = New-Object System.Windows.Forms.Label
    $Network_Title.Location = New-Object System.Drawing.Size(10,170)
    $Network_Title.Size = New-Object System.Drawing.Size(150,20)
    $Network_Title.Text = "Networking:"
    $Network_Title.Font = $LabelFont

    #Makes Software Info Title
    $Software_Title = New-Object System.Windows.Forms.Label
    $Software_Title.Location = New-Object System.Drawing.Size(10,245)
    $Software_Title.Size = New-Object System.Drawing.Size(200,20)
    $Software_Title.Text = "Software Information:"
    $Software_Title.Font = $LabelFont

    #Makes Hardware Info Title
    $Hardware_Title = New-Object System.Windows.Forms.Label
    $Hardware_Title.Location = New-Object System.Drawing.Size(10,320)
    $Hardware_Title.Size = New-Object System.Drawing.Size(200,20)
    $Hardware_Title.Text = "Hardware Information:"
    $Hardware_Title.Font = $LabelFont

    #Makes HW Usage Title
    $HWUsage_Title = New-Object System.Windows.Forms.Label
    $HWUsage_Title.Location = New-Object System.Drawing.Size(10,475)
    $HWUsage_Title.Size = New-Object System.Drawing.Size(200,20)
    $HWUsage_Title.Text = "Hardware Usage:"
    $HWUsage_Title.Font = $LabelFont

    #Makes Additional Info Title
    $Additional_Title = New-Object System.Windows.Forms.Label
    $Additional_Title.Location = New-Object System.Drawing.Size(400,110)
    $Additional_Title.Size = New-Object System.Drawing.Size(300,20)
    $Additional_Title.Text = "Additional Information:"
    $Additional_Title.Font = $LabelFont

    #Adds all the Titles
    $Window.Controls.Add($GetInfo)
    $Window.Controls.Add($Title)
    $Window.Controls.Add($TextBox)
    $Window.Controls.Add($Update)
    $Window.Controls.Add($User_Title)
    $Window.Controls.Add($Network_Title)
    $Window.Controls.Add($Software_Title)
    $Window.Controls.Add($Hardware_Title)
    $Window.Controls.Add($HWUsage_Title)
    $Window.Controls.Add($Additional_Title)
    [void] $Window.ShowDialog()
    

}

Function Info {

    #Updates update field
    $Update.Text = "Generating WMI Queries..."
    $Update.Visible = $True

    #Generates basic info
    $ComputerName=$TextBox.Text
    $RemotePath = "\\" + $ComputerName + "\c$\"
    $CompSys = Get-WMIObject Win32_ComputerSystem -ComputerName $ComputerName
    $OpSys = Get-WMIObject Win32_OperatingSystem -ComputerName $ComputerName
    $Net = Get-WmiObject Win32_NetworkAdapterConfiguration -Computer $ComputerName | Where-Object {$_.DNSDomain -like 'saskenergy.net'}
    $SystemProduct = Get-WMIObject Win32_ComputerSystemProduct -ComputerName $ComputerName
    $Processor = Get-WMIObject Win32_Processor -ComputerName $ComputerName
    $BIOS = Get-WMIObject Win32_BIOS -ComputerName $ComputerName
    $Process = Get-WmiObject -Class Win32_Process -Namespace "root\cimv2" -ComputerName $ComputerName
    $Printers = Get-WmiObject Win32_Printer -ComputerName $ComputerName


    UserInfo
}

Function UserInfo {

    #Updates update field
    $Update.Text = "Getting User Info..."
    
    #Gets user info
    $User = $CompSys.UserName
    $User = $User.TrimStart("SKENERGY\")
    $Person = Get-ADUser $User -Properties *

    #Gets Boot time
    $Boot = $OpSys.ConvertToDateTime($OpSys.LastBootUpTime)

    $Ping = Test-Connection $ComputerName -Count 2

    #Makes labels
    $User_Header1 = New-Object System.Windows.Forms.Label
    $User_Header1.Location = New-Object System.Drawing.Size(10,130)
    $User_Header1.Size = New-Object System.Drawing.Size(40,15)
    $User_Header1.Font = $HeaderFont
    $User_Header1.Text = "User:"

    $User_Info1 = New-Object System.Windows.Forms.Label
    $User_Info1.Location = New-Object System.Drawing.Size(50,130)
    $User_Info1.Size = New-Object System.Drawing.Size(100,15)
    $User_Info1.Font = $RegularFont
    $User_Info1.Text = $Person.displayName

    $User_Header2 = New-Object System.Windows.Forms.Label
    $User_Header2.Location = New-Object System.Drawing.Size(10,145)
    $User_Header2.Size = New-Object System.Drawing.Size(100,15)
    $User_Header2.Font = $HeaderFont
    $User_Header2.Text = "Last Boot Time:"
    
    $User_Info2 = New-Object System.Windows.Forms.Label
    $User_Info2.Location = New-Object System.Drawing.Size(110,145)
    $User_Info2.Size = New-Object System.Drawing.Size(150,15)
    $User_Info2.Font = $RegularFont
    $User_Info2.Text = $Boot

    Network   

}

Function Network {

    #Updates update field
    $Update.Text = "Getting Network..."
    
    #Gets network info    
    $IP = $Net.IPAddress
    $MAC = $Net.MACAddress
    $PC = $Net.DNSHostName

    #Generates labels
    $Net_Header1 = New-Object System.Windows.Forms.Label
    $Net_Header1.Location = New-Object System.Drawing.Size(10,190)
    $Net_Header1.Size = New-Object System.Drawing.Size(70,15)
    $Net_Header1.Font = $HeaderFont
    $Net_Header1.Text = "Hostname:"

    $Net_Info1 = New-Object System.Windows.Forms.Label
    $Net_Info1.Location = New-Object System.Drawing.Size(80,190)
    $Net_Info1.Size = New-Object System.Drawing.Size(100,15)
    $Net_Info1.Font = $RegularFont
    $Net_Info1.Text = $PC

    $Net_Header2 = New-Object System.Windows.Forms.Label
    $Net_Header2.Location = New-Object System.Drawing.Size(10,205)
    $Net_Header2.Size = New-Object System.Drawing.Size(70,15)
    $Net_Header2.Font = $HeaderFont
    $Net_Header2.Text = "IP Address:"

    $Net_Info2 = New-Object System.Windows.Forms.Label
    $Net_Info2.Location = New-Object System.Drawing.Size(80,205)
    $Net_Info2.Size = New-Object System.Drawing.Size(200,15)
    $Net_Info2.Font = $RegularFont
    $Net_Info2.Text = $IP

    $Net_Header3 = New-Object System.Windows.Forms.Label
    $Net_Header3.Location = New-Object System.Drawing.Size(10,220)
    $Net_Header3.Size = New-Object System.Drawing.Size(40,15)
    $Net_Header3.Font = $HeaderFont
    $Net_Header3.Text = "MAC:"

    $Net_Info3 = New-Object System.Windows.Forms.Label
    $Net_Info3.Location = New-Object System.Drawing.Size(50,220)
    $Net_Info3.Size = New-Object System.Drawing.Size(200,15)
    $Net_Info3.Font = $RegularFont
    $Net_Info3.Text = $MAC

    Software
}

Function Software {

    #Updates update field
    $Update.Text = "Getting Software..."
    
    #Links to paths/registry
    $Path = $RemotePath + "Program Files (x86)\Microsoft Application Virtualization Client\sftmime.exe"
    $Reg = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey("LocalMachine",$ComputerName)
    $Key = $Reg.OpenSubKey('SOFTWARE\Microsoft\Internet Explorer')
    
    #Gets values
    $IEVersion = $Key.GetValue('Version')    
    $AppV = [System.Diagnostics.FileVersionInfo]::GetVersionInfo("$Path").FileVersion
    $OS = $OpSys.Caption

    #Generates labels
    $Soft_Header1 = New-Object System.Windows.Forms.Label
    $Soft_Header1.Location = New-Object System.Drawing.Size(10,265)
    $Soft_Header1.Size = New-Object System.Drawing.Size(40,15)
    $Soft_Header1.Font = $HeaderFont
    $Soft_Header1.Text = "O/S:"

    $Soft_Info1 = New-Object System.Windows.Forms.Label
    $Soft_Info1.Location = New-Object System.Drawing.Size(50,265)
    $Soft_Info1.Size = New-Object System.Drawing.Size(300,15)
    $Soft_Info1.Font = $RegularFont
    $Soft_Info1.Text = $OS

    $Soft_Header2 = New-Object System.Windows.Forms.Label
    $Soft_Header2.Location = New-Object System.Drawing.Size(10,280)
    $Soft_Header2.Size = New-Object System.Drawing.Size(110,15)
    $Soft_Header2.Font = $HeaderFont
    $Soft_Header2.Text = "Internet Explorer:"

    $Soft_Info2 = New-Object System.Windows.Forms.Label
    $Soft_Info2.Location = New-Object System.Drawing.Size(120,280)
    $Soft_Info2.Size = New-Object System.Drawing.Size(200,15)
    $Soft_Info2.Font = $RegularFont
    $Soft_Info2.Text = $IEVersion

    $Soft_Header3 = New-Object System.Windows.Forms.Label
    $Soft_Header3.Location = New-Object System.Drawing.Size(10,295)
    $Soft_Header3.Size = New-Object System.Drawing.Size(50,15)
    $Soft_Header3.Font = $HeaderFont
    $Soft_Header3.Text = "App-V:"

    $Soft_Info3 = New-Object System.Windows.Forms.Label
    $Soft_Info3.Location = New-Object System.Drawing.Size(60,295)
    $Soft_Info3.Size = New-Object System.Drawing.Size(200,15)
    $Soft_Info3.Font = $RegularFont
    $Soft_Info3.Text = $AppV

    Hardware
}

Function Hardware {

    #Udpates update field
    $Update.Text = "Getting Hardware..."

    #Gets PC info
    $Model = $SystemProduct.Name
    $Computer = "Lenovo " + $SystemProduct.Version
    $Serial = $SystemProduct.IdentifyingNumber

    #Gets CPU Info
    $CPUName = $Processor.Name
    
    #Gets ram info
    $InstalledRam = (Get-WMIObject Win32_PhysicalMemory -Computername $ComputerName | Measure-Object -Property Capacity -Sum | % {[Math]::Round(($_.sum/1GB),2)})
    $TotalRam = "" + $InstalledRam + "GB"

    #Gets bios info
    $BIOSv = $BIOS.SMBIOSBIOSVersion
    $BIOSDate = $BIOS.ConvertToDateTime($BIOS.ReleaseDate)
    
    #Generates labels
    $Hard_Header1 = New-Object System.Windows.Forms.Label
    $Hard_Header1.Location = New-Object System.Drawing.Size(10,345)
    $Hard_Header1.Size = New-Object System.Drawing.Size(70,15)
    $Hard_Header1.Font = $HeaderFont
    $Hard_Header1.Text = "Computer:"

    $Hard_Info1 = New-Object System.Windows.Forms.Label
    $Hard_Info1.Location = New-Object System.Drawing.Size(80,345)
    $Hard_Info1.Size = New-Object System.Drawing.Size(300,15)
    $Hard_Info1.Font = $RegularFont
    $Hard_Info1.Text = $Computer

    $Hard_Header2 = New-Object System.Windows.Forms.Label
    $Hard_Header2.Location = New-Object System.Drawing.Size(10,360)
    $Hard_Header2.Size = New-Object System.Drawing.Size(50,15)
    $Hard_Header2.Font = $HeaderFont
    $Hard_Header2.Text = "Model:"
    
    $Hard_Info2 = New-Object System.Windows.Forms.Label
    $Hard_Info2.Location = New-Object System.Drawing.Size(60,360)
    $Hard_Info2.Size = New-Object System.Drawing.Size(80,15)
    $Hard_Info2.Font = $RegularFont
    $Hard_Info2.Text = $Model

    $Hard_Header3 = New-Object System.Windows.Forms.Label
    $Hard_Header3.Location = New-Object System.Drawing.Size(10,375)
    $Hard_Header3.Size = New-Object System.Drawing.Size(50,15)
    $Hard_Header3.Font = $HeaderFont
    $Hard_Header3.Text = "Serial:"

    $Hard_Info3 = New-Object System.Windows.Forms.Label
    $Hard_Info3.Location = New-Object System.Drawing.Size(60,375)
    $Hard_Info3.Size = New-Object System.Drawing.Size(100,15)
    $Hard_Info3.Font = $RegularFont
    $Hard_Info3.Text = $Serial

    $Hard_Header4 = New-Object System.Windows.Forms.Label
    $Hard_Header4.Location = New-Object System.Drawing.Size(10,390)
    $Hard_Header4.Size = New-Object System.Drawing.Size(70,15)
    $Hard_Header4.Font = $HeaderFont
    $Hard_Header4.Text = "Processor:"

    $Hard_Info4 = New-Object System.Windows.Forms.Label
    $Hard_Info4.Location = New-Object System.Drawing.Size(80,390)
    $Hard_Info4.Size = New-Object System.Drawing.Size(150,30)
    $Hard_Info4.Font = $RegularFont
    $Hard_Info4.Text = $CPUName

    $Hard_Header5 = New-Object System.Windows.Forms.Label
    $Hard_Header5.Location = New-Object System.Drawing.Size(10,420)
    $Hard_Header5.Size = New-Object System.Drawing.Size(85,15)
    $Hard_Header5.Font = $HeaderFont
    $Hard_Header5.Text = "Memory Size:"

    $Hard_Info5 = New-Object System.Windows.Forms.Label
    $Hard_Info5.Location = New-Object System.Drawing.Size(100,420)
    $Hard_Info5.Size = New-Object System.Drawing.Size(100,15)
    $Hard_Info5.Font = $RegularFont
    $Hard_Info5.Text = $TotalRAM

    $Hard_Header6 = New-Object System.Windows.Forms.Label
    $Hard_Header6.Location = New-Object System.Drawing.Size(10,435)
    $Hard_Header6.Size = New-Object System.Drawing.Size(85,15)
    $Hard_Header6.Font = $HeaderFont
    $Hard_Header6.Text = "BIOS Version:"

    $Hard_Info6 = New-Object System.Windows.Forms.Label
    $Hard_Info6.Location = New-Object System.Drawing.Size(95,435)
    $Hard_Info6.Size = New-Object System.Drawing.Size(100,15)
    $Hard_Info6.Font = $RegularFont
    $Hard_Info6.Text = $BIOSv

    $Hard_Header7 = New-Object System.Windows.Forms.Label
    $Hard_Header7.Location = New-Object System.Drawing.Size(10,450)
    $Hard_Header7.Size = New-Object System.Drawing.Size(70,15)
    $Hard_Header7.Font = $HeaderFont
    $Hard_Header7.Text = "BIOS Date:"

    $Hard_Info7 = New-Object System.Windows.Forms.Label
    $Hard_Info7.Location = New-Object System.Drawing.Size(80,450)
    $Hard_Info7.Size = New-Object System.Drawing.Size(200,15)
    $Hard_Info7.Font = $RegularFont
    $Hard_Info7.Text = $BIOSDate

    HWUsage
}

Function HWUsage {

    #Updates update field
    $Update.Text = "Getting HW Usage..."

    #Gets cpu load
    $CPULoad = "" + $Processor.LoadPercentage + "%"

    #Gets free ram
    $FreeRAM = "" + [Math]::Round(($OpSys.FreePhysicalMemory/1MB),2) + "GB"

    #Gets HDD info
    $HDD1 = Get-WMIObject -Class Win32_LogicalDisk -Namespace "root\cimv2" -ComputerName $ComputerName -Filter "DeviceID='C:'"
    ForEach ($Item in $HDD1) {
        $CSize = [Math]::Round(($HDD1.Size/1Gb),2)
        $CFreeSpace = [Math]::Round(($HDD1.Freespace/1GB),2)
    }
    $HDD2 = Get-WMIObject -Class Win32_LogicalDisk -Namespace "root\cimv2" -ComputerName $ComputerName -Filter "DeviceID='D:'"
    ForEach ($Item in $HDD2) {
        $DSize = [Math]::Round(($HDD2.Size/1Gb),2)
        $DFreeSpace = [Math]::Round(($HDD2.Freespace/1GB),2)
    }

    #Creates hdd labels
    $COut = "" + $CFreeSpace + "GB of " + $CSize + "GB free"
    If ($DSize) {$DOut = "" + $DFreeSpace + "GB of " + $DSize + "GB free"}

    #Generates labels
    $HW_Header1 = New-Object System.Windows.Forms.Label
    $HW_Header1.Location = New-Object System.Drawing.Size(10,505)
    $HW_Header1.Size = New-Object System.Drawing.Size(70,15)
    $HW_Header1.Font = $HeaderFont
    $HW_Header1.Text = "CPU Load:"

    $HW_Info1 = New-Object System.Windows.Forms.Label
    $HW_Info1.Location = New-Object System.Drawing.Size(80,505)
    $HW_Info1.Size = New-Object System.Drawing.Size(70,15)
    $HW_Info1.Font = $RegularFont
    $HW_Info1.Text = $CPULoad

    $HW_Header2 = New-Object System.Windows.Forms.Label
    $HW_Header2.Location = New-Object System.Drawing.Size(10,520)
    $HW_Header2.Size = New-Object System.Drawing.Size(70,15)
    $HW_Header2.Font = $HeaderFont
    $HW_Header2.Text = "Free RAM:"

    $HW_Info2 = New-Object System.Windows.Forms.Label
    $HW_Info2.Location = New-Object System.Drawing.Size(80,520)
    $HW_Info2.Size = New-Object System.Drawing.Size(70,15)
    $HW_Info2.Font = $RegularFont
    $HW_Info2.Text = $FreeRAM

    $HW_Header3 = New-Object System.Windows.Forms.Label
    $HW_Header3.Location = New-Object System.Drawing.Size(10,535)
    $HW_Header3.Size = New-Object System.Drawing.Size(100,15)
    $HW_Header3.Font = $HeaderFont
    $HW_Header3.Text = "Main Partition:"

    $HW_Info3 = New-Object System.Windows.Forms.Label
    $HW_Info3.Location = New-Object System.Drawing.Size(110,535)
    $HW_Info3.Size = New-Object System.Drawing.Size(200,15)
    $HW_Info3.Font = $RegularFont
    $HW_Info3.Text = $COut

    $HW_Header4 = New-Object System.Windows.Forms.Label
    $HW_Header4.Location = New-Object System.Drawing.Size(10,550)
    $HW_Header4.Size = New-Object System.Drawing.Size(100,15)
    $HW_Header4.Font = $HeaderFont
    $HW_Header4.Text = "2nd Partition:"

    $HW_Info4 = New-Object System.Windows.Forms.Label
    $HW_Info4.Location = New-Object System.Drawing.Size(110,550)
    $HW_Info4.Size = New-Object System.Drawing.Size(200,15)
    $HW_Info4.Font = $RegularFont
    $HW_Info4.Text = $DOut

    Additional
}

Function Additional {

    #Updates update field
    $Update.Text = "Getting Add. Info..."

    #Gets ram usage for processes   
    $IERam = $Process | Where-Object {$_.Description -like 'iexplore.exe'} | Measure-Object -Property WorkingSetSize -Sum | % {[Math]::Round(($_.sum/1GB),2)}
    $NotesRam = $Process | Where-Object {$_.Description -like '*notes*'} | Measure-Object -Property WorkingSetSize -Sum | % {[Math]::Round(($_.sum/1GB),2)}
    $VMRam = $Process | Where-Object {$_.Description -like '*vmware*' -or $_.Description -like '*vpx*'} | Measure-Object -Property WorkingSetSize -Sum | % {[Math]::Round(($_.sum/1GB),2)}
    $OfficeRam = $Process | Where-Object {$_.Description -like '*excel*' -or $_.Description -like '*word*' -or $_.Description -like '*powerpnt*' -or $_.Description -like '*onenote*'} | Measure-Object -Property WorkingSetSize -Sum | % {[Math]::Round(($_.sum/1GB),2)}
    
    #Generates text labels for ram usage
    If ($IERam) {$IERamD = "`n" + "IE is using: " + $IERam + "GB."}
    If ($NotesRam) {$NotesRamD = "Notes is using: " + $NotesRam + "GB."}
    If ($VMRam) {$VMRamD = "`n" + "VMWare is using: " + $VMRam + "GB."}
    If ($OfficeRam) {$OfficeRamD = "`n" + "MS Office is using: " + $OfficeRam + "GB."}

    #Generates final text output
    $RAMUse = "" + $NotesRamD + $IERamD + $OfficeRamD + $VMRamD

    #Gets User SID
    $UserObj = New-Object System.Security.Principal.NTAccount($User)
    $UserSID = $UserObj.Translate([System.Security.Principal.SecurityIdentifier])

    #Sets printer keys
    $PrinterKey1 = $UserSID.ToString() + "\Printers\Connections"
    $PrinterKey2 = $UserSID.ToString() + "\Software\Microsoft\Windows NT\CurrentVersion\Windows"
    
    #Sets reg info
    $Type = [Microsoft.Win32.RegistryHive]::Users
    $Reg2 = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey($Type, $ComputerName)

    #Opens first key
    $RegKey = $Reg2.OpenSubKey("$PrinterKey1", $True)
    
    #Creates array and fills it with network printers
    $NetworkPrinters = @()
    ForEach ($SubKey in $RegKey.GetSubKeyNames())
        {$NetworkPrinters = $NetworkPrinters + $SubKey.TrimStart(",,")}
    
    #Opens 2nd key and gets the default printer
    $RegKey2 = $Reg2.OpenSubKey("$PrinterKey2", $True)
    $DefaultPrinter = $RegKey2.GetValue('Device')

    #Creates array and fills it with local printers
    $LocalPrinter = @()
    $LocalPrinter = $LocalPrinter + ""
    ForEach ($Printer in $Printers) {$LocalPrinter = $LocalPrinter + $Printer.Name + "`n"}
        
    #Generates labels
    $Add_Header1 = New-Object System.Windows.Forms.Label
    $Add_Header1.Location = New-Object System.Drawing.Size(400,130)
    $Add_Header1.Size = New-Object System.Drawing.Size(90,15)
    $Add_Header1.Font = $HeaderFont
    $Add_Header1.Text = "RAM Usage:"

    $Add_Info1 = New-Object System.Windows.Forms.Label
    $Add_Info1.Location = New-Object System.Drawing.Size(490,130)
    $Add_Info1.Size = New-Object System.Drawing.Size(200,70)
    $Add_Info1.Font = $RegularFont
    $Add_Info1.Text = $RAMUse

    $Add_Header2 = New-Object System.Windows.Forms.Label
    $Add_Header2.Location = New-Object System.Drawing.Size(400,200)
    $Add_Header2.Size = New-Object System.Drawing.Size(100,15)
    $Add_Header2.Font = $HeaderFont
    $Add_Header2.Text = "Default Printer:"

    $Add_Info2 = New-Object System.Windows.Forms.Label
    $Add_Info2.Location = New-Object System.Drawing.Size(520,200)
    $Add_Info2.Size = New-Object System.Drawing.Size(130,15)
    $Add_Info2.Font = $RegularFont
    $Add_Info2.Text = $DefaultPrinter

    $Add_Header3 = New-Object System.Windows.Forms.Label
    $Add_Header3.Location = New-Object System.Drawing.Size(400,215)
    $Add_Header3.Size = New-Object System.Drawing.Size(100,15)
    $Add_Header3.Font = $HeaderFont
    $Add_Header3.Text = "Local Printers:"

    $Add_Info3 = New-Object System.Windows.Forms.Label
    $Add_Info3.Location = New-Object System.Drawing.Size(520,215)
    $Add_Info3.Size = New-Object System.Drawing.Size(200,70)
    $Add_Info3.Font = $RegularFont
    $Add_Info3.Text = $LocalPrinter
    
    $Add_Header4 = New-Object System.Windows.Forms.Label
    $Add_Header4.Location = New-Object System.Drawing.Size(400,285)
    $Add_Header4.Size = New-Object System.Drawing.Size(120,150)
    $Add_Header4.Font = $HeaderFont
    $Add_Header4.Text = "Network Printers:"

    $Add_Info4 = New-Object System.Windows.Forms.Label
    $Add_Info4.Location = New-Object System.Drawing.Size(520,285)
    $Add_Info4.Size = New-Object System.Drawing.Size(150,300)
    $Add_Info4.Font = $RegularFont
    $Add_Info4.Text = $NetworkPrinters

    $Img = [System.Drawing.Image]::FromFile('\\skynet01\tgardiner\trogdor.png')
    $PicBox = New-Object Windows.Forms.PictureBox
    $PicBox.Image = $Img
    $PicBox.Location = New-Object System.Drawing.Size(600,600)
    $PicBox.Width = $Img.Size.Width
    $PicBox.Height = $Img.Size.Height
        
    Output
}

Function Output {

    #Updates update field
    $Update.Text = "Finsihed"
    
    #Adds all generated headers
    $Window.Controls.Add($User_Header1)
    $Window.Controls.Add($User_Header2)
    $Window.Controls.Add($Net_Header1)
    $Window.Controls.Add($Net_Header2)
    $Window.Controls.Add($Net_Header3)
    $Window.Controls.Add($Soft_Header1)
    $Window.Controls.Add($Soft_Header2)
    $Window.Controls.Add($Soft_Header3)
    $Window.Controls.Add($Hard_Header1)
    $Window.Controls.Add($Hard_Header2)
    $Window.Controls.Add($Hard_Header3)
    $Window.Controls.Add($Hard_Header4)
    $Window.Controls.Add($Hard_Header5)
    $Window.Controls.Add($Hard_Header6)
    $Window.Controls.Add($Hard_Header7)
    $Window.Controls.Add($HW_Header1)
    $Window.Controls.Add($HW_Header2)
    $Window.Controls.Add($HW_Header3)
    If ($DSize) {$Window.Controls.Add($HW_Header4)}
    $Window.Controls.Add($Add_Header1)
    $Window.Controls.Add($Add_Header2)
    $Window.Controls.Add($Add_Header3)
    $Window.Controls.Add($Add_Header4)

    #Adds all generated info
    $Window.Controls.Add($User_Info1)
    $Window.Controls.Add($User_Info2)
    $Window.Controls.Add($Net_Info1)
    $Window.Controls.Add($Net_Info2)
    $Window.Controls.Add($Net_Info3)
    $Window.Controls.Add($Soft_Info1)
    $Window.Controls.Add($Soft_Info2)
    $Window.Controls.Add($Soft_Info3)
    $Window.Controls.Add($Hard_Info1)
    $Window.Controls.Add($Hard_Info2)
    $Window.Controls.Add($Hard_Info3)
    $Window.Controls.Add($Hard_Info4)
    $Window.Controls.Add($Hard_Info5)
    $Window.Controls.Add($Hard_Info6)
    $Window.Controls.Add($Hard_Info7)
    $Window.Controls.Add($HW_Info1)
    $Window.Controls.Add($HW_Info2)
    $Window.Controls.Add($HW_Info3)
    If ($DSize) {$Window.Controls.Add($HW_Info4)}
    $Window.Controls.Add($Add_Info1)
    $Window.Controls.Add($Add_Info2)
    $Window.Controls.Add($Add_Info3)
    $Window.Controls.Add($Add_Info4)

    $Window.Controls.Add($PicBox)
}

Window