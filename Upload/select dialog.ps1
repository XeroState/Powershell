$FileSelect = New-Object System.Windows.Forms.OpenFileDialog
$FileSelect.Title = “Please select a file”
$FileSelect.InitialDirectory = “H:\”
$FileSelect.Filter = “All Files (*.*)|*.*”
$result = $FileSelect.ShowDialog()

If($result -eq “OK”) {
$inputFile = $FileSelect.FileName
write-host $inputfile
}
Else {
Write-Host “Cancelled by user”
}
