#Generate Cert
New-SelfSignedCertificate -CertStoreLocation cert:\currentuser\my `
-Subject "CN=Test Code Signing" `
-KeyAlgorithm RSA `
-KeyLength 2048 `
-Provider "Microsoft Enhanced RSA and AES Cryptographic Provider" `
-KeyExportPolicy Exportable `
-KeyUsage DigitalSignature `
-Type CodeSigningCert

#Grabs Cert
$Cert = Get-ChildItem -Path Cert:\LocalMachine\My -DNS "Test Code Signing"
#Variable for exported Cert
$ExCert = "C:\users\dfrohlick\desktop\cert.cer"
#Export Cert
Export-Certificate -cert $Cert -filepath $ExCert
#Import Cert
Import-Certificate -filepath $ExCert -certstorelocation Cert:\LocalMachine\Root
#Grabs new Root Cert
$NewCert = Get-ChildItem -path Cert:\LocalMachine\Root -DNS "Test Code Signing"

#Inf2Cat - If need to generate cat files
#Inf2Cat /drv:c:\whatever /os:10_x64

#Sign Cat files
Set-AuthenticodeSignature -FilePath "x" -Certificate $NewCert -TimestampServer "http://timestamp.digicert.com"

