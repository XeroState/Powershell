Function Hosts {

    # Gets the IP address of my two VMs
    $W7 = [System.Net.Dns]::GetHostAddresses("vmw7dfroh-1") | Select-Object -ExpandProperty IPAddressToString
    $W10 = [System.Net.Dns]::GetHostAddresses("vmw10dfroh-1") | Select-Object -ExpandProperty IPAddressToString

    # Creates the HOSTS line for the VMs
    $Line1 = $W7 + "	vm7"
    $Line2 = $W10 + "	vm1"

    # Sets the file locatoins
    $File = "C:\Users\dfrohlick\Desktop\Other\hosts.bak"
    $Output = "C:\Users\dfrohlick\Desktop\Other\hosts"
    $Hosts = "C:\Windows\System32\drivers\etc\hosts"

    # Creates a new HOSTS file
    Copy-Item $File $Output

    # Adds both the vm lines
    $Line1 | Add-Content $Output
    $Line2 | Add-Content $Output

    # Copies it into the correct folder
    Copy $Output $Hosts
}

Function Test {
    #Checks to see if Powershell has Admin rights
    ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")
}


If (!(Test)) {
    # Create a new process object that starts PowerShell as Admin
    # This is needed to save a hosts file into the correct folder
    Start-Process Powershell.exe -PassThru -Verb Runas "`"\\skynet01\dfrohlick\Home\Stuff\Scripts\Powershell\zHosts.ps1`"" | Out-Null
    exit
}
Else {
    Hosts
}