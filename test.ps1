###################################################
#  Script to test different install types with PowerShell 
#                                                                                             
#
###################################################


# Variable list for each Installer

# MSI = True/False.  If it's an MSI installer, powershell needs to call msiexec
# Name = Name of the application, used for output
# Installed = Name of the file to check to see if the application is already installed or not
# InstallFile = Actual installer to be called
# Log = Path/name of the log file
# Arguments = List of installer arguments to be passed to the installer
# Values = just the arguments variable converted to a string


Function Alber_Battery_Analysis {
    $MSI = "True"
    $Name = "`"Alber Battery Analysis`""
    $Installed = "C:\Program Files (x86)\Alber\Alber Battery Analysis\BatteryAnalysis.exe"
    $InstallFile = "`"\\skynet02\public\Software\Alber\Battery Analysis\Alber Battery Analysis.msi`""
    $Log = "`"C:\Program Files\SaskEnergy\Logs\Alber Battery Analysis Install.log`""
    $Arguments = @(
        "/i"
        $InstallFile
        "/qb"
        "/l*v+ $Log"
        "RESTART=REALLYSUPPRESS"
    )
    $Values = [string]$Arguments
    Check $Name $Installed $InstallFile $Values $MSI

}

Function Blackbox_Pocket_Modem {
    $MSI = "False"
    $Name = "`"Blackbox Pocket Modem`""
    $Installed = "C:\Program Files (x86)\ModemWiz\ModemWiz.exe"
    $InstallFile = "`"\\skynet02\public\Software\Blackbox\Pocket Modem\Setup.exe`""
    $Arguments = @(
        "/s"
        "/sms"
        "/f1`"\\skynet02\public\Software\Blackbox\Pocket Modem\setup.iss`""
        "/f2`"C:\Program Files\SaskEnergy\Logs\Blackbox Pocket Modem Install.log`""
        "/verbose"
    )
    $Values = [string]$Arguments
    Check $Name $Installed $InstallFile $Values $MSI
}

Function Check ($Name, $Installed, $InstallFile, $Values, $MSI){
    If (Test-Path $Installed) {
        Write-Output "$Name is already installed."
        $Choice = Read-host "Install again anyways [y/n]?"

        If ($Choice -match "n") {Exit}
        If (($Choice -match "y") -and ($MSI -like "True")) {MSIInstall $Name $Values}
        If (($Choice -match "y") -and ($MSI -like "False")) {Install $Name $InstallFile $Values}
    } Else {
        If ($MSI -like "True") {MSIInstall $Name $Values}
        If ($MSI -like "False") {Install $Name $InstallFile $Values}
    }
 }

Function Install ($Name, $InstallFile, $Values) {   
    Write-Output "Installing $Name"
    Start-Process -FilePath $InstallFile -ArgumentList $Arguments -Wait
}

Function MSIInstall ($Name, $Values){
    Write-Output "Installing $Name"
    Start-Process "msiexec.exe" -ArgumentList $Values -Wait
}

Function GUI {
}