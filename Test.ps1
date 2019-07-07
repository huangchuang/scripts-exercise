# How to call powershell script?
#     1. D:\huangchuang\scripts-exercise\test.ps1
#     2. powershell.exe -ExecutionPolicy "Unrestricted" "D:\huangchuang\scripts-exercise\test.ps1"
#     3. invoke-command -computername localhost -FilePath "D:\huangchuang\scripts-exercise\test.ps1" -ArgumentList {0}
#     4. invoke-command -computername localhost -Scriptblock {Get-WmiObject -class win32_processor | select-object -ExpandProperty MaxClockSpeed | format-table HideTableHeaders | out-string}
#     5. invoke-command -computername $Hostname -credential $Cred -scriptblock {Get-WmiObject -class win32_processor | select-object -ExpandProperty MaxClockSpeed | format-table HideTableHeaders | out-string}

# working
$Hostname = 'cnshg-server1'
$Username = 'ap\huangcf'
$Password = 'Doudou3@gtcc'
$pass = ConvertTo-SecureString -AsPlainText $Password -Force
$Cred = New-Object System.Management.Automation.PSCredential -ArgumentList $Username,$pass
invoke-command -computername $Hostname -credential $Cred -scriptblock {systeminfo}


$Hostname = 'USTR-R3T27S22.ap.uis.unisys.com'
$Username = 'USTR-R3T27S22\ABSuite'
$Password = 'Unisys*2012'
$pass = ConvertTo-SecureString -AsPlainText $Password -Force
$Cred = New-Object System.Management.Automation.PSCredential -ArgumentList $Username,$pass
invoke-command -computername $Hostname -credential $Cred -FilePath "D:\OneDrive - Unisys\GetMachineInfo.ps1" -ArgumentList {0}
invoke-command -computername $Hostname -credential $Cred -scriptblock {Get-WmiObject -class win32_processor | select-object -ExpandProperty MaxClockSpeed | format-table HideTableHeaders | out-string}


# Testing
$Hostname = 'USTR-R2T9-27
$Username = 'USTR-R2T9-27\ABSuite'
$Password = 'Administer4Me'
$pass = ConvertTo-SecureString -AsPlainText $Password -Force
$Cred = New-Object System.Management.Automation.PSCredential -ArgumentList $Username,$pass
invoke-command -computername $Hostname -credential $Cred -scriptblock {systeminfo}

