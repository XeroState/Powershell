Function Get-LenovoBIOSSetting
    {
        [cmdletbinding()]
       # Param
        #    (
         #       [Parameter(Mandatory=$false)]
          #          [string[]]$ComputerName = $env:computername,
 
           #     [Parameter(Mandatory=$false)]
            #        [string[]]$Setting
            #)
        $Computer = "PC14228"
        Begin { [System.Collections.ArrayList]$BIOSSetting = @() }
 
        Process
            {
                foreach($Computer in $ComputerName)
                    {
                        if($Setting)
                            {
                                Foreach($Item in $Setting)
                                    {
                                        Try
                                            {
                                                $arrCurrentBIOSSetting = gwmi -class Lenovo_BiosSetting -namespace root\wmi -Filter "CurrentSetting like '$Item,%'" -ComputerName $Computer -ErrorAction Stop | ? { $_.CurrentSetting -ne "" } | Select -ExpandProperty CurrentSetting
                                                Foreach($CurrentBIOSSetting in $arrCurrentBIOSSetting)
                                                    {
                                                        [string]$CurrentItem = $CurrentBIOSSetting.SubString(0,$($CurrentBIOSSetting.IndexOf(',')))
                                                 
                                                        if($CurrentBIOSSetting.IndexOf(';') -gt 0) { [string]$CurrentValue = $CurrentBIOSSetting.SubString($($CurrentBIOSSetting.IndexOf(',')+1),$CurrentBIOSSetting.IndexOf(';')-$($CurrentBIOSSetting.IndexOf(',')+1)) }
                                                        else { [string]$CurrentValue = $CurrentBIOSSetting.SubString($($CurrentBIOSSetting.IndexOf(',')+1)) }
                                                 
                                                        if($CurrentBIOSSetting.IndexOf(';') -gt 0)
                                                            {
                                                                [string]$OptionalValue = $CurrentBIOSSetting.SubString($($CurrentBIOSSetting.IndexOf(';')+1))
                                                                [string]$OptionalValue = $OptionalValue.Replace('[','').Replace(']','').Replace('Optional:','').Replace('Excluded from boot order:','')
                                                            }
                                                        Else { [string]$OptionalValue = 'N/A' }
                                                                                 
                                                        $BIOSSetting += [pscustomobject]@{ComputerName=$Computer;Setting=$CurrentItem;CurrentValue=$CurrentValue;OptionalValue=$OptionalValue;}
                                                 
                                                        Remove-Variable CurrentItem,CurrentValue,OptionalValue -ErrorAction SilentlyContinue -WhatIf:$false
                                                    }
                                                Remove-Variable arrCurrentBIOSSetting -ErrorAction SilentlyContinue -WhatIf:$false
                                            }
                                        Catch { Write-Output "ERROR: UNABLE TO QUERY THE BIOS VIA CLASS [Lenovo_BiosSeting] ON [$Computer] - POSSIBLE INVALID SETTING SPECIFIED [$Item][$CurrentItem]: $_"; throw }
                                    }
                            }
                        Else
                            {
                                Try
                                    {
                                        $arrCurrentBIOSSetting = gwmi -class Lenovo_BiosSetting -namespace root\wmi -ComputerName $Computer -ErrorAction Stop | ? { $_.CurrentSetting -ne "" } | Select -ExpandProperty CurrentSetting
                                        Foreach($CurrentBIOSSetting in $arrCurrentBIOSSetting)
                                            {
                                                [string]$CurrentItem = $CurrentBIOSSetting.SubString(0,$($CurrentBIOSSetting.IndexOf(',')))
                                         
                                                if($CurrentBIOSSetting.IndexOf(';') -gt 0) { [string]$CurrentValue = $CurrentBIOSSetting.SubString($($CurrentBIOSSetting.IndexOf(',')+1),$CurrentBIOSSetting.IndexOf(';')-$($CurrentBIOSSetting.IndexOf(',')+1)) }
                                                else { [string]$CurrentValue = $CurrentBIOSSetting.SubString($($CurrentBIOSSetting.IndexOf(',')+1)) }
                                                                         
                                                if($CurrentBIOSSetting.IndexOf(';') -gt 0)
                                                    {
                                                        [string]$OptionalValue = $CurrentBIOSSetting.SubString($($CurrentBIOSSetting.IndexOf(';')+1))
                                                        [string]$OptionalValue = $OptionalValue.Replace('[','').Replace(']','').Replace('Optional:','').Replace('Excluded from boot order:','')
                                                    }
                                                Else { [string]$OptionalValue = 'N/A' }
                                                                                                                         
                                                $BIOSSetting += [pscustomobject]@{ComputerName=$Computer;Setting=$CurrentItem;CurrentValue=$CurrentValue;OptionalValue=$OptionalValue;}
                                         
                                                Remove-Variable CurrentItem,CurrentValue,OptionalValue -ErrorAction SilentlyContinue -WhatIf:$false
                                            }
                                        Remove-Variable arrCurrentBIOSSetting -ErrorAction SilentlyContinue -WhatIf:$false
                                    }
                                Catch { Write-Output "ERROR: UNABLE TO QUERY THE BIOS VIA CLASS [Lenovo_BiosSeting] ON [$Computer]: $_"; throw }
                            }
                    }
            }
         
        End { $BIOSSetting }
    }
    
Get-LenovoBIOSSetting
Get-LenovoBIOSSetting -Setting Wi%
Get-LenovoBIOSSetting -ComputerName Leno-T470 -Setting 'Fingerprint%'
Get-LenovoBIOSSetting -ComputerName Leno-M900,Leno-M910 -Setting 'USB%','C State Support','Intel(R) Virtualization Technology'
Full code below
1
2
3
4
5
6
7
8
9
10
11
12
13
14
15
16
17
18
19
20
21
22
23
24
25
26
27
28
29
30
31
32
33
34
35
36
37
38
39
40
41
42
43
44
45
46
47
48
49
50
51
52
53
54
55
56
57
58
59
60
61
62
63
64
65
66
67
68
69
70
71
72
73
74
75
76
77
78
79
80
81
82
#region Function Get-LenovoBIOSSetting

endregion Function Get-LenovoBIOSSetting