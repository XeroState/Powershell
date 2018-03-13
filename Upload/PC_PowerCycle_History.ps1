# Gets power cycle history of PC
# Shows restart vs shutdown, who did it, date and time

$PC = Read-Host "What is the PC name?"
Get-WinEvent -ComputerName $PC -FilterHashtable @{logname='System';id=1074} | ForEach-Object {
    $rv = $null
    $rv = New-Object PSObject | Select-Object Date, User, Action
    $rv.Date = $_.TimeCreated
    $rv.User = $_.Properties[6].Value
    $rv.Action = $_.Properties[4].Value
    $rv
    } | Select-Object Date, Action, User