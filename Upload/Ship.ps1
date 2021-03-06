############################################################
# Script Name: Ship.ps1
# Created On: Mar 9, 2016
# Author: David Frohlick
# 
# Purpose: 
#  
###########################################################


Function ShipTo {
    #Get Info to put into doc
    $Name = Read-Host "Who is it going to? (Full name)"
    $Company = Read-Host "SaskEnergy or TransGas Compressor Station?"

    $User = Get-ADUser -Filter {displayName -like $Name} -Properties streetAddress, postalCode, l, telephoneNumber


    $Address = $User.streetAddress
    $PostalCode = $User.postalCode
    $City = $User.l
    $Phone = $User.telephoneNumber

    #Type the info into the doc
    $Word=new-object -ComObject "Word.Application"
    $Doc=$Word.documents.Add()
    $Word.Visible=$True
    $Selection=$Word.Selection
    $Word.Selection.paragraphFormat.SpaceBefore = 0
    $Word.Selection.paragraphFormat.SpaceAfter = 0
    $Selection.Font.Name="Calibri"
    $Selection.Font.Size=16
    $Selection.TypeText("SaskEnergy Inc.")
    $Selection.TypeParagraph()
    $Selection.Font.Size=7
    $Selection.TypeText("Information Systems")
    $Selection.TypeParagraph()
    $Selection.TypeText("1777 Victoria Ave")
    $Selection.TypeParagraph()
    $Selection.TypeText("4th Floor")
    $Selection.TypeParagraph()
    $Selection.TypeText("Regina, SK")
    $Selection.TypeParagraph()
    $Selection.TypeText("S4P 4K5")
    $Selection.TypeParagraph()

    $Selection.Font.Size=16
    $Selection.TypeText("                                $Company")
    $Selection.TypeParagraph()
    $Selection.TypeText("                                $Address")
    $Selection.TypeParagraph()
    $Selection.TypeText("                                $City, Saskatchewan")
    $Selection.TypeParagraph()
    $Selection.TypeText("                                $PostalCode")
    $Selection.TypeParagraph()
    $Selection.TypeText("                                Attention: $Name")
    $Selection.TypeParagraph()
    $Selection.TypeText("                                $Phone")


    $null = [System.Runtime.InteropServices.Marshal]::ReleaseComObject([System.__ComObject]$Word)
    [gc]::Collect()
    [gc]::WaitForPendingFinalizers()
    Remove-Variable word 

    $Repeat = Read-Host "Make another lable? [y/n]"
    If ($Repeat -like 'y') {
        ShipTo
    }
    Else {
        Exit
    }
}

Function ShipFrom {
    #Get Info to put into doc
    $Name = Read-Host "Who is it coming from? (Full Name)"
    $Company = "SaskEnergy"

    $User = Get-ADUser -Filter {displayName -like $Name} -Properties streetAddress, postalCode, l, telephoneNumber


    $Address = $User.streetAddress
    $PostalCode = $User.postalCode
    $City = $User.l
    $Phone = $User.telephoneNumber

    #Type the info into the doc
    $Word=new-object -ComObject "Word.Application"
    $Doc=$Word.documents.Add()
    $Word.Visible=$True
    $Selection=$Word.Selection
    $Word.Selection.paragraphFormat.SpaceBefore = 0
    $Word.Selection.paragraphFormat.SpaceAfter = 0
    $Selection.Font.Name="Calibri"
    $Selection.Font.Size=16
    $Selection.TypeText("SaskEnergy Inc.")
    $Selection.TypeParagraph()
    $Selection.Font.Size=7
    #$Selection.TypeText("Information Systems")
    #$Selection.TypeParagraph()
    $Selection.TypeText("$Address")
    $Selection.TypeParagraph()
    #$Selection.TypeText("4th Floor")
    #$Selection.TypeParagraph()
    $Selection.TypeText("$City, SK")
    $Selection.TypeParagraph()
    $Selection.TypeText("$PostalCode")
    $Selection.TypeParagraph()

    $Selection.Font.Size=16
    $Selection.TypeText("                                SaskEnergy")
    $Selection.TypeParagraph()
    $Selection.TypeText("                                400-1777 Victoria Ave")
    $Selection.TypeParagraph()
    $Selection.TypeText("                                Regina, Saskatchewan")
    $Selection.TypeParagraph()
    $Selection.TypeText("                                S4P 4K5")
    $Selection.TypeParagraph()
    $Selection.TypeText("                                Attention: Information Systems")
    #$Selection.TypeParagraph()
    #$Selection.TypeText("                                $Phone")


    $null = [System.Runtime.InteropServices.Marshal]::ReleaseComObject([System.__ComObject]$Word)
    [gc]::Collect()
    [gc]::WaitForPendingFinalizers()
    Remove-Variable word 

    $Repeat = Read-Host "Make another lable? [y/n]"
    If ($Repeat -like 'y') {
        ShipFrom
    }
    Else {
        Exit
    }
}

Write-Host "Label to ship to someone, or a return label to IS?"
$Choice = Read-Host "To or From?"

If (($Choice -like 'to') -or ($Choice -like 't')) {
    ShipTo
}
Else {
    ShipFrom
}