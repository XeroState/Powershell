# Gets the TPM version

$TPM = Get-WMIObject –class Win32_Tpm –Namespace root\cimv2\Security\MicrosoftTpm
$Version = $TPM.ManufacturerVersion
$Version | Out-File "C:\TPM_Version.txt"