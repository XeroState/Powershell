$VMs = "vmw7dfroh-1","vmw7dfroh-2"

ForEach ($VM in $VMs) {
    $VMProperties = Get-ADComputer $VM -Properties *
    $PwdTime = $VMProperties | select PasswordLastSet
    $PwdTime = ($PwdTime -split'=')[1]
    $PwdTime = $PwdTime.TrimEnd('}')

    Write-Host "For $VM" -ForegroundColor Green
    Write-Host "$PwdTime`n"
}