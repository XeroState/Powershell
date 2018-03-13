#####################################################################
# Script Name: Undelete.ps1
# Created On: August 26th, 2016
# Author: David Frohlick
# Based on Alan Kaplan's Undelete-User.ps1 script, 9-8-14 v 2.0
# Used a lot of his code, however there was a massive amount of cleanup to be done
# 
# Purpose: Automate the searching and restoral of deleted users and computers
#
# Version: 1.0 - Created script
#  
#####################################################################

#Requires -Version 3
#Requires -Module ActiveDirectory

#Add assembly for VB message and input boxes
Add-Type -AssemblyName Microsoft.VisualBasic

Function Bail {
#Exit function

   "Done." 
    If ($Log){Stop-Transcript}
    $VerbosePreference = $oldverbose
    Pause
}

Function Test {
#Prompt about using test mode or live

    $TestMsg ="1) Commit Changes to AD`n2) Run in Test Mode`n0) Quit"

    [int]$TestChoice= [Microsoft.VisualBasic.Interaction]::InputBox($TestMsg,"David's UnDeleter",2)

    Switch($TestChoice) {
        0 {Bail}
        1 {$Test = $False ; Break}
        2 {$Test = $True ;Break}
    }
    If ($Test -eq $True) {
        $VerbosePreference = "Continue"
        Write-Verbose "Running in Test mode"
    }
    Else {
        $VerbosePreference = "SilentlyContinue"
    }
    Log
}

Function Log {
#Prompt about whether to log or not, if outside of ISE

    If (($Host.Name).Contains("ISE") -eq $False) {
        $Window = [System.Windows.Forms.MessageBox]::Show("Would you like a transaction log of script activity?", "Log", "YesNoCancel” , "Question” , "Button2")
        Switch ($Window) {
            Cancel {Bail}
            Yes {
                #prompt for name
                $Log = "$ENV:UserProfile\Desktop\"+$(Get-Date).ToString("yyyyMMdd_HHmm")+"_UndeleteLog.txt"
                $Log = [Microsoft.VisualBasic.Interaction]::InputBox("Log Path",  "Path",$Log)
                Start-Transcript -Path $Log 
            }
            Default {}
        }
    }
    Domain
}

Function Domain {
#Grabs Domain information

    $Domain = [Microsoft.VisualBasic.Interaction]::InputBox("Search for deleted user account(s) in what AD domain", "Domain Name", $MyDomain)
    If ($Domain.Length -eq 0){Bail}

    If ($MyDomain -inotlike $Domain) {
        #Get list of all GC
        $DomainGC = Get-ADDomainController -server $Domain -Filter {(Enabled -eq $True) -and (IsGlobalCatalog -eq $True)}
        $Server =  $DomainGC| Select -Property Hostname | Out-Gridview -Title "Select closest GC" -PassThru 

        #add GC Port to first GC in array
        $Server = $Server.Hostname +":3268"
    }
    Else {
        $Server = $Domain
    }
    Choice
}

Function Choice {
#Prompt for a choice between restoring a computer object or a user

    $ChoiceMsg = "1) Search for Users `n2) Search for Computers `n0) Quit"
    [int]$ChoiceBox = [Microsoft.VisualBasic.Interaction]::InputBox($ChoiceMsg,"Object Type",1)

    Switch($ChoiceBox) {
        0 {Bail}
        1 {User}
        2 {Comp}
    }
}

Function User {
#Function to search for the user object and then restore if not in test mode
    
    #Default is search for all user names
    $User = "*"
    $UserMsg = "You may limit the search to a single NT style SamAccountName, or return all deleted user accounts.`nDo you want to search for a single account?"
    $Retval = [System.Windows.Forms.MessageBox]::Show($UserMsg,"Single User", "YesNoCancel” , "Question” , "Button2")
    
    Switch ($Retval)
    {
        Cancel {Bail}
        Yes {
            #Prompt for name
            $myName= ((Get-ADUser $env:USERNAME).SamAccountName).tostring().ToLower()
            $Msg = "Do not include domain, for example: $myName. `n`nSearch for this SAMAccountName:"
            $User = [Microsoft.VisualBasic.Interaction]::InputBox($Msg,"SAMAccountName")
        }
        Default {}
    }

    If($User -eq ""){Bail}

    #Ask go back how far?
    $Msg = "Limiting the days speeds your search.`nLook for objects going back how many days?"
    [int]$SearchDays= [Microsoft.VisualBasic.Interaction]::InputBox($Msg,"Search Past Number of Days",5)
    if ($SearchDays -eq 0){Bail}

    #subtract seach days to get starting date
    $ChangedDate = (Get-Date).AddDays(-$SearchDays)
    
    #get object representing $domain -- which really is a string
    $DomainObj = Get-ADDomain -Identity $Domain
    Write-Verbose "Getting deleted items path for $Domain" 
    $SearchBase = $DomainObj.DeletedObjectsContainer

    #Write-Verbose 'Getting default restore path'
    #$RestorePath = "OU=Managed,OU=Accounts,DC=saskenergy,DC=net"
    #$RestorePath = [Microsoft.VisualBasic.Interaction]::InputBox("Restore users to this path:", "Restore Path", $RestorePath)
    #if ($RestorePath.Length -eq 0){Bail}

    Write-Output "`nSearching for deleted users in $Domain going back $Searchdays days, please wait  .... `n`t(Users without sufficent rights may get no results from this search`)"

    Try {
        #Do the search.  
        $Users = Get-ADObject -searchbase $Searchbase -Server $Server -resultpagesize 100 -Filter {(whenChanged -ge $ChangedDate) -and (Deleted -eq $True) -and (SamAccountName -like $User) } -includeDeletedObjects -properties DisplayName, Description, SamAccountName, userprincipalname, DistinguishedName, WhenChanged 
    }
    Catch {
        $ErrorMessage = $Error[0].Exception.Message
        Write-Error $Errormessage
        Bail
    }

    #Use where-object to remove computers in the list
    $Users = $Users | Where objectclass -ne "computer"


    #Use select with expressions to show nicer labels
    $Users = $Users | Select @{Name="Display Name";Expression={$_."DisplayName"}},Description, @{Name="NT Account";Expression={$_."SamAccountName"}}, UserPrincipalName,@{Name="Date Deleted";Expression={$_."WhenChanged"}}, DistinguishedName | Sort-Object -Property "Display name" 
    
    #Bail after notification if zero user objects returned
    If ($Users.count -eq 0){
        $Msg= "No users found to undelete in $Domain going back $Searchdays days with filter `"SamAccountName like $User`"."
        $Retval = [System.Windows.Forms.MessageBox]::Show($Msg,"No User(s) Found", "ok” , "Exclamation” , "Button1")
        Bail
    }

    $Msg = "Select User Account(s) to Restore, and click [OK]"
    If ($Test -eq $True) {$Msg = "[Test Mode - Restore is Simulated] $Msg"}

    #Send result list to Out-Gridview.  PassThru allows selection to pass through to restore
    $Restore = $Users| Out-GridView -Title $Msg -Passthru

    ForEach ($User in $Restore){
        "`nRestoring " +$User.DistinguishedName.tostring()
        Restore-ADObject -Server $Domain -identity $User.DistinguishedName -whatif:$Test 
    }

    If ($Log){Stop-Transcript }
    $VerbosePreference = $oldVerbosePreference
}

Function Comp {
#Function to search for the computer object and then restore if not in test mode
    
    #Default is search for all user names
    $Comp = "*"
    $CompMsg = "You may limit the search to a single PC name, or return all deleted computer accounts.`nDo you want to search for a single account?"
    $Retval = [System.Windows.Forms.MessageBox]::Show($CompMsg,"Single User", "YesNoCancel” , "Question” , "Button2")
    
    Switch ($Retval)
    {
        Cancel {Bail}
        Yes {
            #Prompt for name
            $myComp = $env:COMPUTERNAME
            $Msg = "Do not include domain, for example: $myComp. `n`nSearch for this Computer Name:"
            $Comp = [Microsoft.VisualBasic.Interaction]::InputBox($Msg,"CN")
        }
        Default {}
    }

    If($Comp -eq ""){Bail}

    #Ask go back how far?
    $Msg = "Limiting the days speeds your search.`nLook for objects going back how many days?"
    [int]$SearchDays= [Microsoft.VisualBasic.Interaction]::InputBox($Msg,"Search Past Number of Days",5)
    if ($SearchDays -eq 0){Bail}

    #subtract seach days to get starting date
    $ChangedDate = (Get-Date).AddDays(-$SearchDays)
    
    #get object representing $domain -- which really is a string
    $DomainObj = Get-ADDomain -Identity $Domain
    Write-Verbose "Getting deleted items path for $Domain" 
    $SearchBase = $DomainObj.DeletedObjectsContainer

    #Write-Verbose 'Getting default restore path'
    #$RestorePath = "OU=Managed,OU=Accounts,DC=saskenergy,DC=net"
    #$RestorePath = [Microsoft.VisualBasic.Interaction]::InputBox("Restore users to this path:", "Restore Path", $RestorePath)
    #if ($RestorePath.Length -eq 0){Bail}

    Write-Output "`nSearching for deleted computers in $Domain going back $Searchdays days, please wait  .... `n`t(Users without sufficent rights may get no results from this search`)"

    Try {
        #Do the search.  
        $Comps = Get-ADObject -searchbase $Searchbase -Server $Server -resultpagesize 100 -Filter {(whenChanged -ge $ChangedDate) -and (Deleted -eq $True) -and (cn -like $Comp) } -includeDeletedObjects -properties cn, WhenChanged, DistinguishedName 
    }
    Catch {
        $ErrorMessage = $Error[0].Exception.Message
        Write-Error $Errormessage
        Bail
    }

    #Use where-object to remove computers in the list
    $Comps = $Comps | Where objectclass -Like "computer"


    #Use select with expressions to show nicer labels
    $Comps = $Comps | Select @{Name="cn";Expression={$_."cn"}},@{Name="Date Deleted";Expression={$_."WhenChanged"}}, DistinguishedName | Sort-Object -Property "cn"
        
    #Bail after notification if zero user objects returned
    If ($Comps.count -eq 0){
        $Msg= "No computers found to undelete in $Domain going back $Searchdays days with filter `"cn like $Comp`"."
        $Retval = [System.Windows.Forms.MessageBox]::Show($Msg,"No Computer(s) Found", "ok” , "Exclamation” , "Button1")
        Bail
    }

    $Msg = "Select Computer Account(s) to Restore, and click [OK]"
    If ($Test -eq $True) {$Msg = "[Test Mode - Restore is Simulated] $Msg"}

    #Send result list to Out-Gridview.  PassThru allows selection to pass through to restore
    $Restore = $Comps| Out-GridView -Title $Msg -Passthru

    ForEach ($Comp in $Restore){
        "`nRestoring " +$Comp.DistinguishedName.tostring()
        Restore-ADObject -Server $Domain -identity $Comp.DistinguishedName -whatif:$Test 
    }

    If ($Log){Stop-Transcript }
    $VerbosePreference = $oldVerbosePreference
}

cls
$Error.Clear()
Test
