Add-Type -AssemblyName PresentationFramework

Function Test {
    ForEach ($PC in $PCList) {
        $Test = Test-Connection -Count 1 -ComputerName $PC -Quiet
        If ($Test -like "True") {
            $a = [System.Windows.MessageBox]::Show("$PC is Online")
            $PCList = $PCList | Where-Object {$_ -ne "$PC"}
        }
    }
    If ($PCList -gt 0) {
        Start-Sleep -seconds 120
        Test
    }
}

$PCList = Read-Host "Type PC names, separated by only a comma"
$PCList = $PCList.Split(',')

Test