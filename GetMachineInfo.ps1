param(
	[int]$option
)

function GetCPUNumberOfCores
{
	[int64]$result = 0
	
	$rows = Get-WmiObject -class win32_processor | select-object -ExpandProperty NumberOfCores | format-table HideTableHeaders
	foreach ($row in $rows)
	{
		[int64]$rowValue = $row | out-string
		
		$result += $rowValue
	}
	
	return $result
}

function GetCPUSpeed
{
	[int64]$result = 0
	
	$rows = Get-WmiObject -class win32_processor | select-object -ExpandProperty MaxClockSpeed | sort-object -Property MaxClockSpeed -Descending | format-table HideTableHeaders
	foreach ($row in $rows)
	{
		[int64]$rowValue = $row | out-string
		
		$result = $rowValue
		
		break
	}
	
	return $result
}

function GetMemoryPhysicalTotal
{
	[int64]$result = 0
	
	$rows = Get-WmiObject -class "Win32_PhysicalMemory" | select-object -ExpandProperty Capacity | format-table HideTableHeaders
	foreach ($row in $rows)
	{
		[int64]$rowValue = $row | out-string
		$rowValue = ($rowValue / 1024 / 1024 / 1024)
		
		$result += $rowValue
	}
	
	return $result
}

function GetDiskSizeTotal
{
	[int64]$result = 0
	
	$rows = get-disk | select-object -ExpandProperty Size | format-table HideTableHeaders
	foreach ($row in $rows)
	{
		[int64]$rowValue = $row | out-string
		$rowValue = ($rowValue / 1024 / 1024 / 1024)
		
		$result += $rowValue
	}
	
	return $result
}

function GetDiskSizeRemaining
{
	[int64]$result = 0
	
	$rows = get-volume | where-object {$_.DriveLetter -ne $null} | select-object -ExpandProperty SizeRemaining | format-table HideTableHeaders
	foreach ($row in $rows)
	{
		[int64]$rowValue = $row | out-string
		$rowValue = ($rowValue / 1024 / 1024 / 1024)
		
		$result += $rowValue
	}
	
	return $result
}

function Get-SystemInfo
{
  param($ComputerName = $env:ComputerName)
 
      $header='Hostname','OSName','OSVersion','OSManufacturer','OSConfig','Buildtype','RegisteredOwner','RegisteredOrganization','ProductID','InstallDate','StartTime','Manufacturer','Model','Type','Pr$iocessor','BIOSVersion','WindowsFolder','SystemFolder','StartDevice','Culture','UICulture','TimeZone','PhysicalMemory','AvailablePhysicalMemory','MaxVirtualMemory','AvailableVirtualMemory','UsedVirtualMemory','PagingFile','Domain','LogonServer','Hotfix','NetworkAdapter'
      systeminfo.exe /FO CSV /S $ComputerName | Select-Object -Skip 1 | ConvertFrom-CSV -Header $header
}

function Get-VMOSDetail
{
    Param(
        [Parameter()]
        $ComputerName = $Env:ComputerName,
         
        [Parameter()]
        $VMName
         
    )
     
    # Creating HASH Table for object creation
    $MyObj = @{}
     
    # Getting VM Object
    $Vm = Get-WmiObject -Namespace root\virtualization -Query "Select * From Msvm_ComputerSystem Where ElementName='$VMName'" -ComputerName $ComputerName
     
    # Getting VM Details
    $Kvp = Get-WmiObject -Namespace root\virtualization -Query "Associators of {$Vm} Where AssocClass=Msvm_SystemDevice ResultClass=Msvm_KvpExchangeComponent" -ComputerName $ComputerName
     
    # Converting XML to Object
    foreach($CimXml in $Kvp.GuestIntrinsicExchangeItems)
    {
 
        $XML = [XML]$CimXml
 
        if($XML)
        {
            foreach ($CimProperty in $XML.SelectNodes("/INSTANCE/PROPERTY"))
            {
                switch -exact ($CimProperty.Name)
                {
                    "Data"      { $Value = $CimProperty.VALUE }
                    "Name"      { $Name  = $CimProperty.VALUE }
                }
            }
            $MyObj.add($Name,$Value)
        }
    }
     
    # Outputting Object
    New-Object -TypeName PSCustomObject -Property $MyObj
     
}

if ($option -eq 1)
{
	GetCPUNumberOfCores
}
elseif ($option -eq 2)
{
	GetCPUSpeed
}
elseif ($option -eq 3)
{
	GetMemoryPhysicalTotal
}
elseif ($option -eq 4)
{
	GetDiskSizeTotal
}
elseif ($option -eq 5)
{
	GetDiskSizeRemaining
}
else
{
	$results = @()
	$results += GetCPUNumberOfCores
	$results += GetCPUSpeed
	$results += GetMemoryPhysicalTotal
	$results += GetDiskSizeTotal
	$results += GetDiskSizeRemaining
	
	Write-Host $results
}