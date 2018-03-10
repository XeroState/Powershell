Function RandomFood {
	$food = @{"BLT" = 0; "CF" = 0; "FC" = 0; "Pizza"=0; "Chicken Burger" = 0; "Greek Style Chicken" = 0}
        #$food = @{"Pizza" = 0; "CF" = 0}

	For($i = 0; $i -le 50000; $i++)	{		
        $item = $food.GetEnumerator() | Get-Random
        $food.Item($item.Name)++
    }
    
    $List = @()
    $List = $food.GetEnumerator() | Sort -Descending Value

    $Winner = $List[0].Name

    If ($Winner-like "*CF*") {
            $Sauce = $true
            $Sauces = @{"Plum" = 0; "BBQ" = 0}
            For($i=0; $i -le 5000; $i++) {
                    $Item = $Sauces.GetEnumerator() | Get-Random
                    $Sauces.Item($Item.Name)++
            }
            $List2 = @()
            $List2 = $Sauces.GetEnumerator() | Sort-Object -Descending Value
            $WinSauce = $List2[0].Name
    }




    Write-Output "Winning Food is: $Winner"
    If ($Sauce -eq $true) {Write-Output "Winning Sauce is $WinSauce"}
    Write-Output ""
    
    $Food.GetEnumerator() | Sort -Descending Value
    If ($Sauce -eq $true) {$Sauces.GetEnumerator() | Sort -Descending Value}
}

RandomFood