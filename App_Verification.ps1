###################################################################################
# Used to see if any shares or services were created and if any permissions are incorrect
# In the Permissions function it copies 3 files (putty.exe, test.bat and permissions_check.ps1)
# If the base version of these files are moved, update script accordingly
#
# Need to know the Win10Test password, as it will prompt to run the permissions_check script
# Win10Test must also be in the Policy_AppLocker_Config_Powershell group
# No reboot is required, just add the account to the group, run the script, then remove the group
#
# Output is saved in the file "C:\Temp\App_Verification_Results.txt"
# //DF, Oct-2017
###################################################################################


Function Shares {
# Checks for non-standard shares. Exports their names and paths
# Also adds the removal command
    
    # Adds header info the result document
    Add-Content $Log "`r`n====================================================="
    Add-Content $Log "The following are non-standard shares:`r`n"
    
    # Grabs the current shares
    $Shares = Get-WmiObject -Class Win32_Share
    
    # Checks all the current shares against our standard shares
    ForEach ($Share in $Shares) {
        # For anything non-standard, outputs the share name and share path to the reult document
        If (($Share.Name -notlike "ADMIN$") -and ($Share.Name -notlike "C$") -and ($Share.Name -notlike "IPC$") -and ($Share.Name -notlike "print$")) {
            $Name = $Share.Name
            $Path = $Share.Path
            Add-Content $Log "Share Name: $Name"
            Add-Content $Log "Share Path: $Path"
            Add-Content $Log "To remove, have the installer run:   net share $Name /delete`r`n"
        }
    }
    Services
}


Function Services {
# Checks for new services created in the last day. Exports name, account, etc
# Also contains a check for SQL Server Browser and the required steps to disable it
    
    # Adds header info the result document
    Add-Content $Log "`r`n====================================================="
    Add-Content $Log "The following are new services:`r`n"

    # Items within the Event Log's to grab as they are unnamed XML events
    # Without running the script as an Admin, there is no other way of grabbing the info (Admin can convert the event to XML)
    $Select = @{
        'Name' = 0
        'File' = 1
        'Service' = 2
        'Type' = 3
        'Account' = 4
    }

    # Gets the event log, filtering to the EventID that is for a new service creation
    $Events = Get-WinEvent -FilterHashTable @{LogName='System';ID=7045;StartTime=$Time}

    # Goes through the event log, finding the data and adding a named property for them
    ForEach ($Item in $Events) {
        $Selected_Events = ForEach ($Event in $Item) {
        $New_Event = $Event
        ForEach($Key in $Select.Keys) {
            $New_Event = $New_Event | 
            Select-Object *,@{
                'Name' = $Key
                'Expression' = { $_.Properties[$Select[$Key]].Value }
            }
        }
        $New_Event
        }
        
        # Adds the data to the result document
        # Includes the required disable commands for the SQL Server Browser service
        $Service_Name = $Selected_Events.Name
        $Service_File = $Selected_Events.File
        $Service_Type = $Selected_Events.Service
        $Start_Type = $Selected_Events.Type
        $Service_Account = $Selected_Events.Account
        Add-Content $Log "`r`nService Name: $Service_Name"
        Add-Content $Log "Service File Name: $Service_File"
        Add-Content $Log "Service Type: $Service_Type"
        Add-Content $Log "Service Start Type: $Start_Type"
        Add-Content $Log "Service Account: $Service_Account"
        If ($Service_Name -like "*SQL Server Browser*") {
            Add-Content $Log "`r`n***************THIS SERVICE MUST BE TURNED OFF***************"
            Add-Content $Log "**  Have installer do the following:"
            Add-Content $Log "**      sc config SQLBrowser start= disabled"
            Add-Content $Log "**      wmic /NAMESPACE:\\root\Microsoft\SqlServer\ComputerManagement PATH ServerNetworkProtocol WHERE ProtocolName!='Sm' CALL SetDisable"
            Add-Content $Log "**************************************************************"
        }
    }
    Permissions
}


Function Permissions {
# Checks for write and execute permissions
# Exports if you can write and/or execute an exe and a bat

    # If RSAT is installed, automatically puts Win10Test into the Powershell group
    # If RSAT is not installed, pauses the script and tells you to confirm it's in the group
    $RSAT = "C:\Windows\System32\adsiedit.msc"
    $Group = "Policy_AppLocker_Config_Powershell"
    $User = "Win10Test"

    If (!($RSAT|Test-Path)) {
        Write-Output "Microsoft RSAT is not installed on this machine."
        Write-Output "Cannot verify if Win10Test is in the correct group."
        Write-Output "******Confirm that Win10Test is in Policy_AppLocker_Config_Powershell before continuing!!******"
        Pause
    }
    Else {
        # Checks to see if it's already in the group. If it's not, add's it
        $Members = Get-ADGroupMember -Identity $Group -Recursive | Select -ExpandProperty Name
        If ($Members -contains $User) {$Present = "Yes"}
        Else {
            Add-ADGroupMember -Identity $Group -Member $User -Confirm:$false
            $Added = "Yes"
        }
    }

    # Copies Permission Check script and required files
    Copy-Item "\\skynet02\public\Software\Tools and Utilities\App Verification\Permission_Check.ps1" "C:\temp\Permission_Check.ps1" -Force
    Copy-Item "\\skynet02\public\Software\Tools and Utilities\App Verification\putty.exe" "C:\temp\putty.exe" -Force
    Copy-Item "\\skynet02\public\Software\Tools and Utilities\App Verification\test.bat" "C:\temp\test.bat" -Force

    # Runs it as win10test and pauses this script
    Start-Process powershell.exe -Credential "SKENERGY\win10test" -NoNewWindow -ArgumentList " -File C:\temp\Permission_Check.ps1"
    Write-Host "Press Enter once other script has finished"
    Pause
    
    # Removes the permission powershell script (it removes the putty and script file)
    Remove-Item "C:\Temp\Permission_Check.ps1" -Force
}



# Creates a log file and sets the date variable to the previous day for the event log search
$Log = "C:\Temp\App_Verification_Results.txt"
If(!(Test-Path "C:\Temp")) {New-Item "C:\Temp" -ItemType Directory | Out-Null}
New-Item $Log -ItemType File | Out-Null
$Time = (Get-Date).AddDays(-1)

# Launching Powershell from within Powershell keeps the working directory.  By default that's H:\
# Win10Test doesn't have access to your H:\ so script sets itself to a general location
Set-Location C:\Temp

Shares