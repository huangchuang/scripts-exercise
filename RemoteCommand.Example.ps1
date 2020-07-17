# How to call powershell script?
#     1. D:\huangchuang\scripts-exercise\test.ps1
#     2. powershell.exe -ExecutionPolicy "Unrestricted" "D:\huangchuang\scripts-exercise\test.ps1"
#     3. invoke-command -computername localhost -FilePath "D:\huangchuang\scripts-exercise\test.ps1" -ArgumentList {0}
#     4. invoke-command -computername localhost -Scriptblock {Get-WmiObject -class win32_processor | select-object -ExpandProperty MaxClockSpeed | format-table HideTableHeaders | out-string}
#     5. invoke-command -computername $Hostname -credential $Cred -scriptblock {Get-WmiObject -class win32_processor | select-object -ExpandProperty MaxClockSpeed | format-table HideTableHeaders | out-string}

# working with domain accounts only on in-domain machines
$Hostname = 'cnshg-server1'
$Username = 'ap\huangcf'
$Password = 'Doudou12@gtcc'
$pass = ConvertTo-SecureString -AsPlainText $Password -Force
$Cred = New-Object System.Management.Automation.PSCredential -ArgumentList $Username,$pass
invoke-command -computername $Hostname -credential $Cred -scriptblock {systeminfo}

$Hostname = 'USTR-R1T10S23'
$Username = 'ap\huangcf'
$Password = 'Doudou12@gtcc'
$pass = ConvertTo-SecureString -AsPlainText $Password -Force
$Cred = New-Object System.Management.Automation.PSCredential -ArgumentList $Username,$pass
invoke-command -computername $Hostname -credential $Cred -scriptblock {systeminfo}
