###########################################################
# Script Name: App_Uninstall.ps1
# Created On: Feb 19, 2016
# Author: David Frohlick
# 
# Purpose: Automate the remote removal of applications
#  
###########################################################


Function Info {
    #Display choices; Loops until a valid choice is made
    Write-Host "`n"
    Write-Host "                  What do you want to uninstall?"
    Write-Host "`n"

    [Int]$MenuChoice = 0

    While ($MenuChoice -lt 1 -or $MenuChoice -gt 17) {
        Write-Host "   1. Help & Manual 6.0        |          2. Microsoft Visio 2010"
        Write-Host "   3. AutoCAD LT 2015          |          4. Google Earth Pro"
        Write-Host "   5. Articulate               |          6. SnagIt 11"
        Write-Host "   7. Skype                    |          8. Adobe Acrobat Standard"
        Write-Host "   9. Intrepid                 |         10. Mobilink [547/679/760/998]"
        Write-Host "  11. Adobe Acrobat Pro        |         12. Adobe CS6 Design Std"
        Write-Host "  13. Adobe Indesign           |         14. Adobe Illustrator"
        Write-Host "  15. Adobe LiveCycle          |         16. Adobe Photoshop"
        Write-Host "                       17. **Exit**"
        Write-Host "`n"
        [Int]$MenuChoice = Read-Host "Enter Option [1-17]"
        If ($MenuChoice -eq '17') {Exit}
        $ComputerName = Read-Host "What is the computer name?"
    }
    
    Commands
}


Function Commands {
    #Variables
    $Computer = "\\" + $ComputerName
    $Temp = $Computer + "\C$\Temp\Temp\"
    $TestPath = $Computer + "\C$\Windows\"
    $TempRemoval = "Remove-Item $Temp -Force -Recurse -ErrorAction 'SilentlyContinue'"

    #Help and Manual 6.0; 2 uninstallers
    $HelpAndManual1 = "`"C:\Program Files (x86)\EC Software\HelpAndManual6\unins001.exe`" /SILENT"
    $HelpAndManual2 = "`"C:\Program Files (x86)\EC Software\HelpAndManual6\unins000.exe`" /SILENT"

    #Visio 2010; copy required xml and bat files
    $Visio1 = 'Copy-Item \\skynet01\dfrohlick\Home\Stuff\Scripts\RemoveVisio -Destination $Temp -Recurse'
    $Visio2 = "`"C:\Temp\Temp\RemoveVisio.bat`""
    $Visio3 = 'Copy-Item \\skynet01\dfrohlick\Home\Stuff\Scripts\RepairOffice -Destination $Temp -Recurse'
    $Visio4 = "`"C:\Temp\Temp\RepairOffice\RepairOffice.bat`""

    #AutoCAD LT 2015; copy bat file to run msi
    $CADLT1 = 'Copy-Item -Path \\skynet01\dfrohlick\Home\Stuff\Scripts\Uninstall\AutoCADLT2015.bat -Destination (New-Item $Temp -Type Container -Force) -Container -Force'
    $CADLT2 = "`"C:\Temp\Temp\AutoCADLT2015.bat`""

    #Google Earth Pro; copy bat file to run msi
    $GE1 = 'Copy-Item -Path \\skynet01\dfrohlick\Home\Stuff\Scripts\Uninstall\GoogleEarth.bat -Destination (New-Item $Temp -Type Container -force) -Container -force'
    $GE2 = "`"C:\Temp\Temp\GoogleEarth.bat`""

    #Articulate `13; copy silent iss files
    $Art1 = 'Copy-Item -Recurse -Filter Articulate* -Path \\skynet01\dfrohlick\Home\Stuff\Scripts\Uninstall\ -Destination (New-Item $Temp -Type Container -Force) -Container -Force'
    $Art2 = "`"C:\Program Files (x86)\InstallShield Installation Information\{6F39B67B-5A19-46EC-BB9F-A88C2F2D9730}\setup.exe`" -s -f1C:\Temp\Temp\Uninstall\ArticulateCharRemove.iss -sms"
    $Art3 = "`"C:\Program Files (x86)\InstallShield Installation Information\{3E5131E9-1241-4E43-8036-E870C0DE3012}\setup.exe`" -s -f1C:\Temp\Temp\Uninstall\ArticulateReplayRemove.iss -sms"
    $Art4 = "`"C:\Program Files (x86)\InstallShield Installation Information\{3E5131E9-1241-4E43-8036-E870C0DE2012}\setup.exe`" -s -f1C:\Temp\Temp\Uninstall\ArticulateRemove.iss -sms"

    #SnagIt 11; copy bat file to run msi
    $Snag1 = 'Copy-Item -Path \\skynet01\dfrohlick\Home\Stuff\Scripts\Uninstall\SnagIt.bat -Destination (New-Item $Temp -Type Container -Force) -Container -Force'
    $Snag2 = "`"C:\Temp\Temp\SnagIt.bat`""

    #Skype 7.37; copy bat file to run msi
    $Skype1 = 'Copy-Item -Path \\skynet01\dfrohlick\Home\Stuff\Scripts\Uninstall\Skype.bat -Destination (New-Item $Temp -Type Container -Force) -Container -Force'
    $Skype2 = "`"C:\Temp\Temp\Skype.bat`""

    #Adobe Acrobat Std 11; copy bat file to run msi
    $AcroStd1 = 'Copy-Item -Path \\skynet01\dfrohlick\Home\Stuff\Scripts\Uninstall\AcroStd.bat -Destination (New-Item $Temp -Type Container -Force) -container -Force'
    $AcroStd2 = "`"C:\Temp\Temp\AcroStd.bat`""

    #Intrepid; copy bat file to run msi
    $Intrepid1 = 'Copy-Item -Path \\skynet01\dfrohlick\Home\Stuff\Scripts\Uninstall\Intrepid.bat -Destination (New-Item $Temp -Type Container -Force) -Container -Force'
    $Intrepid2 = "`"C:\Temp\Temp\Intrepid.bat`""

    #Mobilink; copy bat file which searches registry and runs the correct msi for mobilink and drivers
    $Mobi1 = 'Copy-Item -Path \\skynet01\dfrohlick\Home\Stuff\Scripts\Uninstall\Mobilink.bat -Destination (New-Item $Temp -Type Container -Force) -Container -Force'
    $Mobi2 = "`"C:\Temp\Temp\Mobilink.bat`""

    #CS6 Design Std; copy bat file to run msi
    $CS6a = 'Copy-Item -Path \\skynet01\dfrohlick\Home\Stuff\Scripts\Uninstall\CS6DStd.bat -Destination (New-Item $Temp -Type Container -Force) -Container -Force'
    $CS6b = "`"C:\Temp\Temp\CS6Dstd.bat`""

    #InDesign; copy bat file to run msi
    $InDesign1 = 'Copy-Item -Path \\skynet01\dfrohlick\Home\Stuff\Scripts\Uninstall\InDesign.bat -Destination (New-Item $Temp -Type Container -Force) -Container -Force'
    $InDesign2 = "`"C:\Temp\Temp\InDesign.bat`""

    #Photoshop; copy bat file to run msi
    $Photo1 = 'Copy-Item -Path \\skynet01\dfrohlick\Home\Stuff\Scripts\Uninstall\Photoshop.bat -Destination (New-Item $Temp -Type Container -Force) -Container -Force'
    $Photo2 = "`"C:\Temp\Temp\Photoshop.bat`""

    #Illustrator; copy bat file to run msi
    $Illus1 = 'Copy-Item -Path \\skynet01\dfrohlick\Home\Stuff\Scripts\Uninstall\Illustrator.bat -Destination (New-Item $Temp -Type Container -Force) -Container -Force'
    $Illus2 = "`"C:\Temp\Temp\Illustrator.bat`""

    #LiveCycle Designer; copy bat file to run msi
    $Live1 = 'Copy-Item -Path \\skynet01\dfrohlick\Home\Stuff\Scripts\Uninstall\LiveCycle.bat -Destination (New-Item $Temp -Type Container -Force) -Container -Force'
    $Live2 = "`"C:\Temp\Temp\LiveCycle.bat`""

    #Adobe Acrobat Pro 11; copy bat file to run msi
    $AcroPro1 = 'Copy-Item -Path \\skynet01\dfrohlick\Home\Stuff\Scripts\Uninstall\AcroPro.bat -Destination (New-Item $Temp -Type Container -Force) -container -Force'
    $AcroPro2 = "`"C:\Temp\Temp\AcroPro.bat`""
    
    Execution
}

Function Execution {
    #Runs the command based on choice
    If (Test-Path $TestPath){
    
        #Most start by emptying Temp/Temp as it shouldn't exist and it ensures pathing is correct for the rest of the commands; end with emptying it as well
        Switch($MenuChoice) {
            1{psexec $Computer $HelpAndManual1; psexec $Computer $HelpAndManual2}
            2{Invoke-Expression $TempRemoval; Invoke-Expression $Visio1; psexec $Computer $Visio2; Invoke-Expression $Visio3; psexec $Computer $Visio4;Invoke-Expression $TempRemoval}
            3{Invoke-Expression $TempRemoval; Invoke-Expression $CADLT1; psexec $Computer $CADLT2; Invoke-Expression $TempRemoval}
            4{Invoke-Expression $TempRemoval; Invoke-Expression $GE1; psexec $Computer $GE2; Invoke-Expression $TempRemoval}
            5{Invoke-Expression $TempRemoval; Invoke-Expression $Art1; psexec $Computer $Art4; psexec $Computer $Art3; psexec $Computer $Art2; Invoke-Expression $TempRemoval}
            6{Invoke-Expression $TempRemoval; Invoke-Expression $Snag1; psexec $Computer $Snag2; Invoke-Expression $TempRemoval}
            7{Invoke-Expression $TempRemoval; Invoke-Expression $Skype1; psexec $Computer $Skype2; Invoke-Expression $TempRemoval}
            8{Invoke-Expression $TempRemoval; Invoke-Expression $AcroStd1; psexec $Computer $AcroStd2; Invoke-Expression $TempRemoval}
            9{Invoke-Expression $TempRemoval; Invoke-Expression $Intrepid1; psexec $Computer $Intrepid2; Invoke-Expression $TempRemoval}
            10{Invoke-Expression $TempRemoval; Invoke-Expression $Mobi1; psexec $Computer $Mobi2; Invoke-Expression $TempRemoval}
            11{Invoke-Expression $TempRemoval; Invoke-Expression $AcroPro1; psexec $Computer $AcroPro2; Invoke-Expression $TempRemoval}
            12{Invoke-Expression $TempRemoval; Invoke-Expression $CS6a; psexec $Computer $CS6b; Invoke-Expression $TempRemoval}
            13{Invoke-Expression $TempRemoval; Invoke-Expression $InDesign1; psexec $Computer $InDesign2; Invoke-Expression $TempRemoval}
            14{Invoke-Expression $TempRemoval; Invoke-Expression $Illus1; psexec $Computer $Illus2; Invoke-Expression $TempRemoval}
            15{Invoke-Expression $TempRemoval; Invoke-Expression $Live1; psexec $Computer $Live2; Invoke-Expression $TempRemoval}
            16{Invoke-Expression $TempRemoval; Invoke-Expression $Photo1; psexec $Computer $Photo2; Invoke-Expression $TempRemoval}
        }
    Pause        
    }
    Else{
        #CatchAll to pause
        Write-Host "Unable to reach"$ComputerName
        pause
    }
}

Info