###########################################################
# Script Name: Registry_Search.ps1
# Created On: May 18, 2016
# Author: David Frohlick
# 
# Purpose: Searches through registry for thinga-ma-jigs
#  
###########################################################

Function Info {
#Gets basic info

    $PC = Read-Host -Prompt "What is the Computer Name?"
    $Search = Read-Host -Prompt "What are you searching for?"
    $KeyType = Read-Host -Prompt "Which Hive? [LocalMachine // Users]"
    $List = Read-Host -Prompt "Everything in Hive or just 1 Key?`n'Hive' or 'Key'"
    $Classes = Read-Host -Prompt "Include Classes key? [y / n]"

#Moves on based on if it's the User hive or LocalMachine hive    
    If ($KeyType -like 'Users') {User}
    If ($KeyType -like 'Localmachine') {Machine}
}


Function User {
#Gets further info if User hive

    $User = Read-Host -Prompt "What is the username?"
    $UserObj = New-Object System.Security.Principal.NTAccount("SKENERGY", $User)
    $UserSID = $UserObj.Translate([System.Security.Principal.SecurityIdentifier])

#If searching just 1 key
    If ($List -like 'key') {
        $Key = Read-Host -Prompt "Which Key to search? `nie. SOFTWARE or System\CurrentControlSet"
        Write-Host "`nSearching... `nThis may take awhile..."
        Search "$Key"
        Write-Host "`nFinished Searching"
        Pause
        }
#If searching the entire hive, pulls from txt file with all key names
    If ($List -like 'hive') {
        Write-Host "`nSearching... `nThis may take awhile..."
        Get-Content "\\skynet01\dfrohlick\home\stuff\scripts\powershell\lists\hkcu.txt" | ForEach-Object{$Key = $UserSID.ToString() + "\" + $_; Search "$Key"}
        Write-Host "`nFinished Searching"
        Pause
        }
}


Function Machine {
#If doing the localmachine hive

#If searching just 1 key
    If ($List -like 'key') {
        $Key = Read-Host -Prompt "Which Key to search? `nie. SOFTWARE or System\CurrentControlSet"
        Write-Host "`nSearching... `nThis may take awhile..."
        Search "$Key"
        Write-Host "`nFinished Searching"
        Pause
        }
#If searching entire hive
    If ($List -like 'hive') {
        Write-Host "`nSearching... `nThis may take awhile..."
        Get-Content "\\skynet01\dfrohlick\home\stuff\scripts\powershell\lists\hklm.txt" | ForEach-Object{$Key = $_; Search "$Key"}
        Write-Host "`nFinished Searching"
        Pause
        }
}


Function Search ($CurrentKey) {
#First level key search function, pulling in key from either the User or Machine function
    
    $Search = "*" + $Search +"*"
    $Type = [Microsoft.Win32.RegistryHive]::$KeyType
    $RegKey = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey($Type, $PC)

    Try {
    
        $SubKey = $RegKey.OpenSubKey($CurrentKey)

#Looks at each name. If it matches it outputs the location
#Otherwise it passes the keyname down to the recurse search function
        ForEach ($SubKeyName in $SubKey.GetSubKeyNames()){
            If ($SubKeyName -like $Search) {Write-Host "Found: $CurrentKey\$Subkeyname" -ForegroundColor Green}
            $TempKey = $CurrentKey + "\" + $SubKeyName
            Write-Host "Searching Top Key: $TempKey" -ForegroundColor Gray
            If ($Classes -like 'y') {RegSearch_Classes "$TempKey"}
            If ($Classes -like 'n') {RegSearch "$TempKey"}
        }
    }
    Catch [System.Security.SecurityException] {
        Out-Null
    }
}



Function RegSearch_Classes ($NewKey) {
#Recursive search function with classes hive

    Try {
#If we are searching a key including the massive Classes hive        

            If($NewKey){
                $SubSubKey = $RegKey.OpenSubKey($NewKey)
#Loops through each looking for a match
#Passes the new key name back to this function all over again to recursively look
                ForEach ($Item in $SubSubKey.GetSubKeyNames()) {
                    $NewTempKey = $NewKey + "\" + $Item
                    If ($Item -like $Search) {Write-Host "Found: $NewTempKey" -ForegroundColor Green}
                    RegSearch_Classes $NewTempKey
                }
            }
    }
    Catch [System.Security.SecurityException] {
    Out-Null
    }
}

Function RegSearch {
#Recursive search function without classes hive
    Try {        
        If ($NewKey -notlike '*Classes*') {
            $SubSubKey = $RegKey.OpenSubKey($NewKey)
#Loops through each looking for a match
#Passes the new key name back to this function all over again                
            ForEach ($Item in $SubSubKey.GetSubKeyNames()) {
                $NewTempKey = $NewKey + "\" + $Item
                If ($Item -like $Search) {Write-Host "Found: $NewTempKey" -ForegroundColor Green}
                RegSearch $NewTempKey
            }
        }

    }
    Catch [System.Security.SecurityException] {
        Out-Null
    }
}



Info