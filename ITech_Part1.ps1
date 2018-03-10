###########################################################################
#
# Name: ITech_Part1.ps1
#
# Author: David Frohlick
#
# Purpose: Combine several installers into 1 script to make building easier
#
# Version : 3.3 - Forgot to copy the hyperterminal shortcut
#                 3.2 - Log file header information
#                 3.1 - Logging has been cleaned up
#                 3.0 - All software is now added
#                 2.0 - Reg key is added for post-reboot installs
#                 1.0 - Intial Script
#
##########################################################################

Function Info {
# Collection of things

    $Log = "C:\Program Files\SaskEnergy\Logs\ITech_Install.log"
    $RepGroup = Read-Host "What is the users Station Tracker Replication Group number?"

    $Title = "______________Instrument Technician Imaging Script______________"
    $Msg = "`r`n`r`nThis log file will log what the script is doing, not what the installers are doing."
    $Msg2 = "`r`nFor errors involving the actual setups look at their respective log file."
    $Msg3 = "`r`nIf this script errors, the error will be logged in the section it occured in for easier troubleshooting.`r`n"

    Add-Content $Log $Title
    Add-Content $Log $Msg
    Add-Content $Log $Msg2
    Add-Content $Log $Msg3

    Station_Tracker
}

Function Station_Tracker {
# Copies the Station Tracker files
# The user configuration is still required at time of swap

    # Clears global error variables
    $Error.Clear()

    Write-Output "`r`n**********************************************************`r`n" >> $Log
    Write-Output "Copying required Station Tracker Files" | Tee $Log -Append
    Write-Output "Remember, user configuration is still required!" | Tee $Log -Append

    New-Item -Path C:\ -Name "Tracker" -Type Directory >> $Log
    Copy-Item "\\skynet02\public\Tracker Admin\Installation\Version2016.01.05\Tracker.mde" "C:\Tracker\Tracker.mde" >> $Log
    Copy-Item "\\skynet02\public\Tracker Admin\Installation\tracker.ico" "C:\Tracker\tracker.ico" >> $Log
    Copy-Item "\\skynet02\public\Tracker Admin\RepGroup$RepGroup\Base\statdata.mdb" "C:\Tracker\statdata.mdb" >> $Log
    Copy-Item "\\skynet02\public\Tracker Admin\Installation\Shortcut\Tracker.lnk" "C:\Users\Public\Desktop\Tracker.lnk" >> $Log
    Copy-Item "\\skynet02\public\Tracker Admin\Installation\Shortcut\Virtual 2003 Replication Manager\Replication Manager.lnk" >> $Log
    
    If ($Error) {
        Write-Output $Error >> $Log
    }

    OpenBSI
}

Function OpenBSI {
# This is part 1 of the OpenBSI installer. A reboot is required before part 2

    # Clears global error variables
    $Error.Clear()
    
    Write-Output "`r`n**********************************************************`r`n" >> $Log
    Write-Output "`n`nInstalling OpenBSI, Part 1" | Tee $Log -Append
    Write-Output "Part 2 will install after a reboot" | Tee $Log -Append

    Start-Process -FilePath "https://is.saskenergy.net/tierii/Lists/softwareinstalls/DispForm.aspx?ID=146&RootFolder=%2A"

    Write-Output "`n`nDo up to Step 11, but do not restart the computer!!!"
    Pause

    cmd.exe /c "\\skynet02\public\Software\Bristol\OpenBSI_Install.bat"
        
    If ($Error) {
        Write-Output $Error >> $Log
    }

    475Easy
}

Function 475Easy {
# Runs the Emerson 475 Easy Upgrade installer

    # Clears global error variables
    $Error.Clear()

    Write-Output "`r`n**********************************************************`r`n" >> $Log
    Write-Output "`n`nInstalling Emerson 475 Easy Upgrade" | Tee $Log -Append
    Write-Output "`nClick OK on the 3 prompts"

    cmd.exe /c "\\skynet02\public\Software\Emerson\475_Easy_Upgrade.bat"
        
    If ($Error) {
        Write-Output $Error >> $Log
    }

    PSI_Modem
}

Function PSI_Modem {
# Runs the Phoenix PSI Modem Install

    # Clears global error variables
    $Error.Clear()

    Write-Output "`r`n**********************************************************`r`n" >> $Log
    Write-Output "`n`nInstalling Emerson 475 Easy Upgrade" | Tee $Log -Append

    cmd.exe /c "\\skynet02\public\Software\Phoenix\PSI_Modem_Install.bat"
        
    If ($Error) {
        Write-Output $Error >> $Log
    }

    SCADAPack_Vision
}

Function SCADAPack_Vision {
# Calls the SCADAPack Vision installer

    # Clears global error variables
    $Error.Clear()

    Write-Output "`r`n**********************************************************`r`n" >> $Log
    Write-Output "`n`nInstalling Control Microsystems SCADAPack Vision" | Tee $Log -Append

    Start-Process -FilePath "https://is.saskenergy.net/tierii/Lists/softwareinstalls/DispForm.aspx?ID=241&RootFolder=%2A"

    Start-Process -FilePath "\\skynet02\public\Software\Control Microsystems\SCADAPack Vision\setup.exe" -Wait

        
    If ($Error) {
        Write-Output $Error >> $Log
    }

    902M
}

Function 902M {
# Calls the 902M Config Software

    # Clears global error variables
    $Error.Clear()

    Write-Output "`r`n**********************************************************`r`n" >> $Log
    Write-Output "`n`nInstalling Galvanic 902M Configuration Software" | Tee $Log -Append

    Start-Process -FilePath "https://is.saskenergy.net/tierii/Lists/softwareinstalls/DispForm.aspx?ID=237&RootFolder=%2A"

    Start-Process -FilePath "\\skynet02\public\Software\Galvanic\902M Configuration Software\v6.07.07\Setup.exe" -Wait
        
    If ($Error) {
        Write-Output $Error >> $Log
    }

    Meterlink
}

Function Meterlink {
# Calls the Daniel Meterlink installer

    # Clears global error variables
    $Error.Clear()

    Write-Output "`r`n**********************************************************`r`n" >> $Log
    Write-Output "`n`nInstalling Daniel Meterlink" | Tee $Log -Append

    Start-Process -FilePath "https://is.saskenergy.net/tierii/Lists/softwareinstalls/DispForm.aspx?ID=120&RootFolder=%2A"

    Start-Process -FilePath "\\skynet02\public\Software\Emerson\Daniel Meterlink\Daniel MeterLink 1.31.009 full.exe" -Wait

    Write-Output "`n`nOnce finished the configuration steps, press Enter to continue"
    Pause
        
    If ($Error) {
        Write-Output $Error >> $Log
    }

    Radlink
}

Function Radlink {
# Runs the RAD-Link script

    # Clears global error variables
    $Error.Clear()

    Write-Output "`r`n**********************************************************`r`n" >> $Log
    Write-Output "`n`nInstalling Phoenix RAD-Link" | Tee $Log -Append

    cmd.exe /c "\\skynet02\public\Software\Phoenix\RAD-Link_Install.bat"
        
    If ($Error) {
        Write-Output $Error >> $Log
    }

    903Analyzer
}

Function 903Analyzer {
# Runs the 903 Analyzer script

    # Clears global error variables
    $Error.Clear()

    Write-Output "`r`n**********************************************************`r`n" >> $Log
    Write-Output "`n`nInstalling Galvanic 903 Analyzer" | Tee $Log -Append

    Start-Process -FilePath "https://is.saskenergy.net/tierii/Lists/softwareinstalls/DispForm.aspx?ID=234&RootFolder=%2A"

    cmd.exe /c "\\skynet02\public\Software\Galvanic\903_Analyzer_Install.bat"

    Write-Output "`n`nPress Enter to continue the script after the installer is finished"
    Pause
        
    If ($Error) {
        Write-Output $Error >> $Log
    }

    USR5637
}

Function USR5637 {
# Runs the US Robotics 5637 script

    # Clears global error variables
    $Error.Clear()

    Write-Output "`r`n**********************************************************`r`n" >> $Log
    Write-Output "`n`nInstalling US Robotics 5637 Modem" | Tee $Log -Append

    cmd.exe /c "\\skynet02\public\Software\US Robotics\USR5637_Install.bat"
        
    If ($Error) {
        Write-Output $Error >> $Log
    }

    SICK
}

Function SICK {
# Runs the SICK Mepaflow script
    
    # Clears global error variables
    $Error.Clear()

    Write-Output "`r`n**********************************************************`r`n" >> $Log
    Write-Output "`n`nInstalling SICK MEPAFLOW600 CBM" | Tee $Log -Append

    cmd.exe /c "\\skynet02\public\Software\SICK\MEPAFLOW600_CBM_Install.bat"

    If ($Error) {
        Write-Output $Error >> $Log
    }

    Roclink
}

Function Roclink {
# Runs the Roclink script
    
    # Clears global error variables
    $Error.Clear()

    Write-Output "`r`n**********************************************************`r`n" >> $Log
    Write-Output "`n`nInstalling Fisher Controls Roclink" | Tee $Log -Append

    cmd.exe /c "\\skynet02\public\Software\Fisher Controls\Roclink_Install.bat"
        
    If ($Error) {
        Write-Output $Error >> $Log
    }

    Blackbox
}

Function Blackbox {
# Runs the Blackbox Pocket Modem script
    
    # Clears global error variables
    $Error.Clear()

    Write-Output "`r`n**********************************************************`r`n" >> $Log
    Write-Output "`n`nInstalling Blackbox Pocket Modem" | Tee $Log -Append

    cmd.exe /c "\\skynet02\public\Software\Blackbox\Pocket_Modem_Install.bat"
        
    If ($Error) {
        Write-Output $Error >> $Log
    }

    WinHelp
}

Function WinHelp {
# Runs the old Windows Help script
    
    # Clears global error variables
    $Error.Clear()

    Write-Output "`r`n**********************************************************`r`n" >> $Log
    Write-Output "`n`nInstalling Windows Help" | Tee $Log -Append

    cmd.exe /c "\\skynet02\public\Software\Microsoft\WinHelp_Install.cmd"
        
    If ($Error) {
        Write-Output $Error >> $Log
    }

    Flowcheck
}

Function Flowcheck {
# Runs the old Flowcheck script
    
    # Clears global error variables
    $Error.Clear()

    Write-Output "`r`n**********************************************************`r`n" >> $Log
    Write-Output "`n`nInstalling Emerson FlowCheck" | Tee $Log -Append

    cmd.exe /c "\\skynet02\public\Software\Emerson\FlowCheck_Install.bat"
        
    If ($Error) {
        Write-Output $Error >> $Log
    }

    Hyperterminal
}

Function Hyperterminal {
# Calls the HyperTerminal installer

    # Clears global error variables
    $Error.Clear()

    Write-Output "`r`n**********************************************************`r`n" >> $Log
    Write-Output "`n`nInstalling Hilgraeve HyperTerminal" | Tee $Log -Append

    Start-Process -FilePath "https://is.saskenergy.net/tierii/Lists/softwareinstalls/DispForm.aspx?ID=320&RootFolder=%2A"

    Start-Process -FilePath "\\skynet02\public\Software\Hilgraeve\Hyperterminal\htpe7.exe" -Wait
    Copy "\\skynet02\public\Software\Hilgraeve\Hyperterminal\Shortcut\Hyper Terminal.lnk" "C:\Users\Public\Desktop\Hyper Terminal.lnk"

    Write-Output "`n`nPress Enter to continue after licensing the application"
    Pause
        
    If ($Error) {
        Write-Output $Error >> $Log
    }

    Startalk
}

Function Startalk {
# Runs the Startalk XP script
    
    # Clears global error variables
    $Error.Clear()

    Write-Output "`r`n**********************************************************`r`n" >> $Log
    Write-Output "`n`nInstalling FlowServe StarTalk XP" | Tee $Log -Append

    cmd.exe /c "\\skynet02\public\Software\FlowServe\StarTalk_XP_Install.bat"
        
    If ($Error) {
        Write-Output $Error >> $Log
    }

    Prolink3
}

Function Prolink3 {
# Runs the Prolink III script

    # Clears global error variables
    $Error.Clear()

    Write-Output "`r`n**********************************************************`r`n" >> $Log
    Write-Output "`n`nInstalling Emerson Prolink III" | Tee $Log -Append

    Start-Process -FilePath "https://is.saskenergy.net/tierii/Lists/softwareinstalls/DispForm.aspx?ID=334&RootFolder=%2A"

    cmd.exe /c "\\skynet02\public\Software\Emerson\ProlinkIII_Install.bat"
        
    If ($Error) {
        Write-Output $Error >> $Log
    }

    RTU
}

Function RTU {
# Runs the RTU Maintenance script
    
    # Clears global error variables
    $Error.Clear()

    Write-Output "`r`n**********************************************************`r`n" >> $Log
    Write-Output "`n`nInstalling RTU Maintenance" | Tee $Log -Append

    cmd.exe /c "\\skynet02\public\Software\SaskEnergy\RTUMaintenance_Install.bat"
        
    If ($Error) {
        Write-Output $Error >> $Log
    }

    MasterLink
}

Function MasterLink {
# Runs the MasterLink SQL script

    # Clears global error variables
    $Error.Clear()

    Write-Output "`r`n**********************************************************`r`n" >> $Log
    Write-Output "`n`nInstalling Honeywell MasterLink SQL" | Tee $Log -Append

    Start-Process -FilePath "https://is.saskenergy.net/tierii/Lists/softwareinstalls/DispForm.aspx?ID=264&RootFolder=%2A"

    cmd.exe /c "\\skynet02\public\Software\Honeywell\MasterLinkSQL.bat"
        
    If ($Error) {
        Write-Output $Error >> $Log
    }

    BWTech
}

Function BWTech {
# Calls the BW Technologies FleetManager installer

    # Clears global error variables
    $Error.Clear()

    Write-Output "`r`n**********************************************************`r`n" >> $Log
    Write-Output "`n`nInstalling BW Technologies FleetManager" | Tee $Log -Append

    Start-Process -FilePath "https://is.saskenergy.net/tierii/Lists/softwareinstalls/DispForm.aspx?ID=81&RootFolder=%2A"

    Start-Process -FilePath "\\skynet02\public\Software\Honeywell\BW Technologies\Fleet Manager II - V2.6.3\BWFleetManager-II_V2.6.3.exe" -Wait

    Copy "\\skynet02\public\Software\Honeywell\BW Technologies\Fleet Manager II - V2.6.3\IRLinkDriver\GA-USB1-IR.inf" "C:\Windows\inf\GA-USB1-IR.inf" >> $Log
        
    If ($Error) {
        Write-Output $Error >> $Log
    }

    PSIConf
}

Function PSIConf {
# Runs the PSI-Conf script

    # Clears global error variables
    $Error.Clear()

    Write-Output "`r`n**********************************************************`r`n" >> $Log
    Write-Output "`n`nInstalling Phoenix PSI-Conf" | Tee $Log -Append

    Start-Process -FilePath "https://is.saskenergy.net/tierii/Lists/softwareinstalls/DispForm.aspx?ID=339&RootFolder=%2A"

    Start-Process -FilePath "\\skynet02\public\Software\Phoenix\PSI-Conf\PSI-CONF_Setup_v2.20.exe" -Wait

    Copy "\\skynet02\public\Software\Phoenix\PSI-Conf\ProjectData.SAV" "C:\ProgramData\Phoenix Contact\PSIConfSoftware\ProjectData.SAV"
        
    If ($Error) {
        Write-Output $Error >> $Log
    }

    COMIP
}

Function COMIP {
# Runs the COM-IP script

    # Clears global error variables
    $Error.Clear()

    Write-Output "`r`n**********************************************************`r`n" >> $Log
    Write-Output "`n`nInstalling Tactical COM-IP" | Tee $Log -Append

    Start-Process -FilePath "https://is.saskenergy.net/tierii/Lists/softwareinstalls/DispForm.aspx?ID=171&RootFolder=%2A"

    Start-Process -FilePath "\\skynet02\public\Software\Tactical\COM IP\v4.9.3\COMIP493.exe" -Wait

    Write-Output "`n`nPress Enter once configuration of COM-IP has been completed"
    Pause
        
    If ($Error) {
        Write-Output $Error >> $Log
    }

    ModScan
}

Function ModScan {
# Runs the ModScan script
    
    # Clears global error variables
    $Error.Clear()

    Write-Output "`r`n**********************************************************`r`n" >> $Log
    Write-Output "`n`nInstalling ModScan/ModSim" | Tee $Log -Append

    cmd.exe /c "\\skynet02\public\Software\WinTech\ModScan_Install.bat"
        
    If ($Error) {
        Write-Output $Error >> $Log
    }

    3095USB
}

Function 3095USB {
# Runs the 3095 USB Modem script

    # Clears global error variables
    $Error.Clear()

    Write-Output "`r`n**********************************************************`r`n" >> $Log
    Write-Output "`n`nInstalling ZoomTel Zoom 3095 Modem" | Tee $Log -Append

    Start-Process -FilePath "https://is.saskenergy.net/tierii/Lists/softwareinstalls/DispForm.aspx?ID=319&RootFolder=%2A"

    Write-Output "`n`nModem must be plugged in before proceeding"
    Pause
    Start-Process -FilePath "\\skynet02\public\Software\ZoomTel\Zoom_3095_Win7-64\Setup64.exe" -Wait

    Write-Output "Paused in case -wait does not work correctly on the installer.  Press Enter to continue"
    Pause
        
    If ($Error) {
        Write-Output $Error >> $Log
    }

    Final
}

Function Final {
# End of stuff

    # Adding RegKey to run the final 2 stuff after reboot
    Push-Location
    Set-Location HKCU:
    New-ItemProperty -Path .\SOFTWARE\Microsoft\Windows\CurrentVersion\RunOnce -Name Part2 -Type string -Value "\\skynet02\public\Software\Tools and Utilities\Imaging\Instrument Tech\ITech_2.lnk" | Out-File $Log -Append
    Pop-Location

    Write-Output "Computer needs a reboot, has to complete the Masterlink and OpenBSI installs after the reboot" | Tee $Log -Append
    Pause
    Exit
}


Info