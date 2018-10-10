Function ListPrograms
{  
	param($SearchPaths)

	
	$Applications = @()
	
	foreach($RegPath in $SearchPaths)
	{
		$QueryPath = dir $RegPath -Name
		
		foreach($Name in $QueryPath)
		{
			$Application = (Get-ItemProperty -Path $RegPath$Name).DisplayName
			$Applications += $Application
		}
	}
	
	foreach($Application in $Applications | sort)
	{
		Write-Host $Application
	}
}

if ([IntPtr]::Size -eq 8)
{
	# Write-Host "[*] OS: x64"
	
	$SearchPaths = (
	"Registry::HKEY_CURRENT_USER\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\",
	"Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\",
	"Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\")
	ListPrograms -SearchPaths $SearchPaths
 }
else
{
	# Write-Host "[*] OS: x86"
	
	$SearchPaths = ("Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\")
	ListPrograms -RegPath $SearchPaths
}
