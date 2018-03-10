Function Tea {
    $Rand = Get-Random -input 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12 -count 12

    $Tea = Get-Content H:\Home\Stuff\Scripts\PowerShell\Lists\Tea.txt
    $TeaHash = @{}

    $i = 1

    Foreach ($Flavour in $Tea) {
        $TeaHash.Add($Flavour,($Rand[$i]))
        $i++
    }

    $Choice = Read-Host "Pick a number between 1 and 11, inclusive"

    $Pick = $TeaHash.GetEnumerator() | Where-Object -Property Value -EQ $Choice
    Write-Host ("Your tea is " + $Pick.Name) -ForegroundColor Green

    $Repeat = Read-Host "Again?"
    If ($Repeat -like 'y') {
        Tea
    }
}

Tea