
Function EST {
    #Eastern Time Zone
    $EST = [System.TimeZoneInfo]::ConvertTimeBySystemTimeZoneId([DateTime]::Now,"US Eastern Standard Time")
    Write-Host "`nUS Eastern Standard Time:`n$EST"
}

Function PST {
    #Pacific Standard Time
    $PST = [System.TimeZoneInfo]::ConvertTimeBySystemTimeZoneId([DateTime]::Now,"Pacific Standard Time")
    Write-Host "`nPacific Standard Time:`n$PST"
}

Function AST {
    #Atlantic Standard Time
    $AST = [System.TimeZoneInfo]::ConvertTimeBySystemTimeZoneId([DateTime]::Now,"Atlantic Standard Time")
    Write-Host "`nAtlantic Standard Time:`n$AST"
}

EST
PST
AST