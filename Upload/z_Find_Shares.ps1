

Function Info {
#Gathers Info
    $Server = Read-Host "What server are we working on? [ie skynet01 or rsback01]"
    Search
}

Function Search {
#Function to search the event log
    
    #Grab Event Log
    $Start = (Get-Date) - (New-TimeSpan -Days 30)
    $EventLog = Get-WinEvent -ComputerName $Server -FilterHashtable @{LogName='System';ID=2511;StartTime=$Start}

    #For every event log item gathered
    ForEach ($Event in $EventLog) {
        #Gather the EventData Data field, split it at the space.  [0] is the sharename and [1] is the physical folder path
        $Temp = [xml]$Event.ToXml()
        $Values = $Temp.Event.EventData.Data
        $Value = $Values.split(" ")
        Add-content C:\Users\dfrohlick\desktop\$server.txt -value $Value[0]
        }
}

Info