$Servers = Get-Content \\skynet01\dfrohlick\Home\Stuff\Scripts\PowerShell\Lists\PrintServers.txt
$Time = (Get-Date).AddDays(-7)

ForEach ($Server in $Servers) {
    Write-Host "Generating $Server"
    $i = $null
    $EventLog = $null
    $Report = $null
    $Report = @()

    $EventLog = Get-WinEvent -ComputerName $Server -FilterHashtable @{LogName='Microsoft-Windows-PrintService/Operational';ID=307;StartTime=$Time}
    ForEach ($Event in $EventLog) {
        #If the event log item matches the search paramters from the Info function
        $Temp = [xml]$Event.ToXml()
        [Int]$TempValue = $Temp.Event.UserData.DocumentPrinted.Param7
        $TempValue = (($TempValue /1024)/1024)
        $TempValue = [Math]::Round($TempValue,3)
            If ($TempValue -ge 30) {
           
                #GridView/CSV must have data properties, must create a member with a name and value
                $i = New-Object PSObject
        
                #Time of the print job
                $Time = ($Event.TimeCreated).ToShortDateString() + " " + ($Event.TimeCreated).ToShortTimeString()
                $i | Add-Member -MemberType NoteProperty -Name "Time" -Value $Time
                #Who printed it
                $Username = $Temp.Event.UserData.DocumentPrinted.Param3
                $i | Add-Member -MemberType NoteProperty -Name "User" -Value $Username
                #Name of the print job
                $DocName = $Temp.Event.UserData.DocumentPrinted.Param2
                $i | Add-Member -MemberType NoteProperty -Name "Document" -Value $DocName          
                #Where it came from
                $Computer = $Temp.Event.UserData.DocumentPrinted.Param4
                $i | Add-Member -MemberType NoteProperty -Name "Computer" -Value $Computer
                #To what printer
               $Printer = $Temp.Event.UserData.DocumentPrinted.Param5
                $i | Add-Member -MemberType NoteProperty -Name "Printer" -Value $Printer
                #How many pages it was
                $Pages = $Temp.Event.UserData.DocumentPrinted.Param8
                $i | Add-Member -MemberType NoteProperty -Name "Pages" -Value $Pages
                #Size of the job
                $JobSize = $Temp.Event.UserData.DocumentPrinted.Param7
                $JobSize = (($JobSize / 1024)/1024)
                $JobSize = [Math]::Round($JobSize,3)
                $i | Add-Member -MemberType NoteProperty -Name "FileSize" -Value $JobSize
                #How many pages it was
                $JobNumber = $Temp.Event.UserData.DocumentPrinted.Param1
                $i | Add-Member -MemberType NoteProperty -Name "JobNumber" -Value $JobNumber
                
                #Adds fields to the report
                $Report += $i
            }
        }
        $Report | Export-CSV \\skynet01\dfrohlick\Home\$Server.csv
    }