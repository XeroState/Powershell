$Computer = "skyprt03"		
$Results = qwinsta /server:$Computer
$ObjMembers = ($Results[0].trim(" ") -replace ("\b *\B")).split(" ")
$Results = $Results[1..$($Results.Count - 1)]
$RDPArray = @()
    ForEach ($Result in $Results) {
	    $RDPMember = New-Object Object
		Add-Member -InputObject $RDPMember -MemberType NoteProperty -Name $ObjMembers[1] -Value $Result.Substring(19,22).Trim()
		Add-Member -InputObject $RDPMember -MemberType NoteProperty -Name $ObjMembers[2] -Value $Result.Substring(41,7).Trim()
		Add-Member -InputObject $RDPMember -MemberType NoteProperty -Name $ObjMembers[3] -Value $Result.Substring(48,8).Trim()
		If ($full) {
		    Add-Member -InputObject $RDPMember -MemberType NoteProperty -Name $ObjMembers[0] -Value $Result.Substring(1,18).Trim()
			Add-Member -InputObject $RDPMember -MemberType NoteProperty -Name $ObjMembers[4] -Value $Result.Substring(56,12).Trim()
			Add-Member -InputObject $RDPMember -MemberType NoteProperty -Name $ObjMembers[5] -Value $Result.Substring(68,8).Trim()
		}
		$RDPArray += $RDPMember
	}

ForEach ($Item in $RDPArray.GetEnumerator()) {
    If ($Item.Username -like 'dfrohlick-adm') {
        $id = $Item.id
        write-host $id
    }
    
}