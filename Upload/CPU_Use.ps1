###########################################################################
# Name: CPU_Use.ps1
# Get CPU % for each process
###########################################################################

$Name = Read-Host "What is the PC Name?"
$Date = Get-Date -format "MM-dd-yyyy_hh-mm-ss"
$Log = "C:\Temp\$Date.log"

Get-WmiObject -ComputerName $Name Win32_PerfFormattedData_PerfProc_Process | Select-Object Name,PercentProcessorTime | Sort-Object PercentProcessorTime -desc | Out-File $Log