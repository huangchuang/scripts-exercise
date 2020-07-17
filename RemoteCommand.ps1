param(
	[string]$Hostname,
	[string]$Username,
	[string]$Password,
	[string]$FilePath
)

function connect ($Hostname, $Username, $Password, $FilePath)
{
	$pass = ConvertTo-SecureString -AsPlainText $Password -Force
	$Cred = New-Object System.Management.Automation.PSCredential -ArgumentList $Username,$pass
	invoke-command -computername $Hostname -credential $Cred -FilePath $FilePath
}

# RemoteCommand.ps1 -Hostname USTR-R3T27S22.ap.uis.unisys.com -Username USTR-R3T27S22\ABSuite -Password Unisys*1 -FilePath "D:\huangchuang\scripts-exercise\GetMachineInfo.ps1"
# connect "USTR-R3T27S22.ap.uis.unisys.com" "USTR-R3T27S22\ABSuite" "Unisys*2012" "D:\huangchuang\scripts-exercise\GetMachineInfo.ps1"

connect $Hostname $Username $Password $FilePath