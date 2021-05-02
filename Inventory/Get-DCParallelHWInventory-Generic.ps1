<#
	.NOTES
	THIS CODE IS MADE AVAILABLE AS IS, WITHOUT WARRANTY OF ANY KIND. THE
	RISK OF THE USE OR THE RESULTS FROM THE USE OF THIS CODE REMAINS WITH
	THE USER.

	.SYNOPSIS
	Script to create DC hardware inventory
	
	.DESCRIPTION
	This script will create a hardware inventory of all domain controllers
	in an AD forest using the Invoke-Parallel.ps1 runspace script to create
	parallel connections and gather inventories in parallel improving per-
	formance.
	
	.PARAMETER -Passthru
	Creates an object from Query results
	
	.PARAMETER -Credential
	Allows admin to pass a credential object into the script for an authenticated
	connection to remote computr objects
	
	.OUTPUTS
	CSV file of hardware inventory with date/time stamp
	
	.EXAMPLE
	PS C:\> Get-DCParallelHWInventory-Generic.ps1 -Passthru -Credential $creds

#>


[CmdletBinding()]
param
(
	[Parameter(Mandatory = $false)]
	[Switch]$PassThru,
	[Parameter(Mandatory = $false)]
	[ValidateNotNullOrEmpty()]
	[String]$ForestName,
	[Parameter(Mandatory = $false)]
	[ValidateNotNullOrEmpty()]
	[System.Management.Automation.PsCredential]$Credential
)


#Region Variables
$driveRoot = (Get-Location).Drive.Root
$logFolder = "{0}{1}" -f $driveRoot, "Logs"
$reportFldr = "{0}{1}" -f $driveRoot, "Reports"
$ns = 'root\CIMv2'
$hive = [uint32]'0x80000002' # HKLM
$subkey = 'SOFTWARE\Microsoft\Windows NT\CurrentVersion'
$value = 'ReleaseID'
$osProps = "BuildNumber", "Caption", "FreePhysicalMemory", "Name", "TotalVisibleMemorySize", "Version"
$compProps = "Domain", "Model", "Manufacturer", "Name"
$procProps = "Caption", "SystemName", "Name", "MaxClockSpeed", "AddressWidth", "NumberOfCores", "NumberOfLogicalProcessors"
$Headings = "Computer Name", "Network Adapter", "Description", "MAC Address", "IP Address", "IP Subnet", "IP Gateway", "DNS Server Search Order",`
"DNS Domain Suffixes", "Dynamic DNS Registration Enabled", "Primary WINS Server", "Secondary WINS Server", "NetBIOS Option"
# need invoke-parallel from same folder as this script
. "$PSScriptRoot\Invoke-Parallel.ps1"
#EndRegion

#Region Functions


Function Test-PathExists
{
	#Begin function to check path variable and return results
	[CmdletBinding()]
	Param
	(
		[Parameter(Mandatory, Position = 0)]
		[string]$Path,
		[Parameter(Mandatory, Position = 1)]
		$PathType
	)
	
	#Define variables
	$VerbosePreference = "Continue"
	
	Switch ($PathType)
	{
		File	{
			If ((Test-Path -Path $Path -PathType Leaf) -eq $true)
			{
				Write-Verbose -Message "File: $Path already exists..." -Verbose
			}
			Else
			{
				New-Item -Path $Path -ItemType File -Force
				Write-Verbose -Message "File: $Path not present, creating new file..." -Verbose
			}
		}
		Folder
		{
			If ((Test-Path -Path $Path -PathType Container) -eq $true)
			{
				Write-Verbose -Message "Folder: $Path already exists..." -Verbose
			}
			Else
			{
				New-Item -Path $Path -ItemType Directory -Force
				Write-Verbose -Message "Folder: $Path not present, creating new folder" -Verbose
			}
		}
	}
} #end function Test-PathExists


Function Write-Logfile
{
	#Begin function to log entries to file
	Param (
		[Parameter(Mandatory = $true)]
		$logEntry,
		[Parameter(Mandatory = $true)]
		$logFile,
		[Parameter(Mandatory = $false)]
		$level = 1
	)
	
	Switch ($level)
	{
		1 { $loglevel = "INFO" }
		2 { $loglevel = "WARN" }
		3 { $loglevel = "ERROR" }
	}
	
	Write-Verbose -Message $logentry -Verbose
	$now = [DateTime]::UtcNow
	$timeStamp = Get-Date $now -DisplayHint Time
	("{0} [{1}] - {2}" -f $timeStamp, $logLevel, $logEntry) | Out-File $logFile -Append
	
} #End function Write-LogFile


Function Get-MyInvocation
{
	return $MyInvocation
} #end function Get-MyInvocation


Function Get-AnADObject
{
	[CmdletBinding()]
	param
	(
		[Parameter(Mandatory = $true)]
		[string]$DomainController,
		[Parameter(Mandatory = $true)]
		[string]$SearchRoot,
		[Parameter(Mandatory = $false)]
		[string]$SearchScope,
		[Parameter(Mandatory = $false)]
		[string]$Filter,
		[Parameter(Mandatory = $false)]
		[string[]]$Properties,
		[Parameter(Mandatory = $false)]
		[ValidateSet(389, 3268, 636, 3269)]
		[int32]$Port,
		[Parameter(Mandatory = $false)]
		[System.Management.Automation.PsCredential]$Credential
	)
	
	$AuthenticationType = [System.DirectoryServices.AuthenticationTypes]::Signing -bor [System.DirectoryServices.AuthenticationTypes]::Sealing -bor [System.DirectoryServices.AuthenticationTypes]::Secure
	
	if ($Port -eq 389 -or $Port -eq 636) { $SearchRoot = "LDAP://{0}:{1}/{2}" -f ($DomainController, $Port, $SearchRoot) }
	elseif ($Port -eq 3268 -or $Port -eq 3269) { $SearchRoot = "GC://{0}:{1}/{2}" -f ($DomainController, $Port, $SearchRoot) }
	else { $SearchRoot = "LDAP://{0}/{1}" -f ($DomainController, $SearchRoot) }
	
	if ($Credential)
	{
		$DirectoryEntryUserName = $Credential.UserName.ToString()
		$DirectoryEntryPassword = $Credential.GetNetworkCredential().Password.ToString()
		$objDirectoryEntry = New-Object System.DirectoryServices.DirectoryEntry($SearchRoot, $DirectoryEntryUserName, $DirectoryEntryPassword, $AuthenticationType)
	}
	else
	{
		$objDirectoryEntry = New-Object System.DirectoryServices.DirectoryEntry($SearchRoot)
		$objDirectoryEntry.psbase.AuthenticationType = $AuthenticationType
	}
	
	$objDirectorySearcher = New-Object System.DirectoryServices.DirectorySearcher
	$objDirectorySearcher.SearchRoot = $objDirectoryEntry
	if ($SearchScope) { $objDirectorySearcher.SearchScope = $SearchScope }
	$objDirectorySearcher.PageSize = 1000
	$objDirectorySearcher.ReferralChasing = "All"
	$objDirectorySearcher.CacheResults = $true
	foreach ($property in $Properties) { [void]$objDirectorySearcher.PropertiesToLoad.Add($property) }
	if ($Filter) { $objDirectorySearcher.Filter = $Filter }
	$colADObject = $objDirectorySearcher.FindAll()
	
	return $colADObject
} #end function Get-AnADObject


Function Search-DnsNames
{
	[CmdletBinding()]
	param
	(
		[Parameter(Mandatory = $true)]
		[ValidateNotNullOrEmpty()]
		[string]$HostName,
		[Parameter(Mandatory = $true)]
		[ValidateNotNullOrEmpty()]
		[string]$Type
	)
	
	try
	{
		$queryResult = Resolve-DnsName $HostName -Type $Type -DnsOnly -ErrorAction Stop
		$objProperties = [PSCustomObject]@{
			HostName   = $HostName;
			RecordData = $queryResult.Name.ToString();
			Type	      = $queryResult.Type.ToString();
			Section    = $queryResult.Section.ToString();
			IPAddress  = $queryResult.IPAddress.ToString();
			Status     = "Resolved";
		}
	}
	catch
	{
		$objProperties = [PSCustomObject]@{
			HostName   = $HostName;
			RecordData = $null;
			Type	      = $null;
			Section    = $null;
			IPAddress  = $null;
			Status     = $Error[0].Exception.Message;
		}
	}
	return $objProperties
} #end function Search-DnsNames


Function Get-FqdnFromDN
{
	[CmdletBinding()]
	param
	(
		[Parameter(Mandatory = $true)]
		[string]$DistinguishedName
	)
	
	if ([string]::IsNullOrEmpty($DistinguishedName) -eq $true) { return $null }
	$domainComponents = $DistinguishedName.ToString().ToLower().Substring($DistinguishedName.ToString().ToLower().IndexOf("dc=")).Split(",")
	for ($i = 0; $i -lt $domainComponents.count; $i++)
	{
		$domainComponents[$i] = $domainComponents[$i].Substring($domainComponents[$i].IndexOf("=") + 1)
	}
	$fqdn = [string]::Join(".", $domainComponents)
	
	return $fqdn
} #end function Get-FqdnFromDN


Function Find-WritableDomainController
{
	[CmdletBinding()]
	param
	(
		[Parameter(Mandatory = $true, HelpMessage = "Enter the FQDN for the target Active Directory domain.")]
		[ValidateNotNullOrEmpty()]
		[string]$Domain,
		[Parameter(Mandatory = $false, HelpMessage = "Enter the target Active Directory site name.")]
		[string]$ADSite,
		[Parameter(Mandatory = $false)]
		[System.Management.Automation.PsCredential]$Credential
	)
	
	$locatorOptions = [System.DirectoryServices.ActiveDirectory.LocatorOptions]::WriteableRequired
	
	if ($Credential)
	{
		$DirectoryEntryUserName = [string]$Credential.UserName
		$DirectoryEntryPassword = [string]$Credential.GetNetworkCredential().Password
		$directoryContext = New-Object System.DirectoryServices.ActiveDirectory.DirectoryContext([System.DirectoryServices.ActiveDirectory.DirectoryContextType]::Domain, $Domain, $DirectoryEntryUserName, $DirectoryEntryPassword)
		
		try
		{
			$objDc = [System.DirectoryServices.ActiveDirectory.DomainController]::FindOne($directoryContext, $ADSite, $locatorOptions)
		}
		catch
		{
			$objDc = [System.DirectoryServices.ActiveDirectory.DomainController]::FindOne($directoryContext, $locatorOptions)
		}
	}
	else
	{
		$directoryContext = New-Object System.DirectoryServices.ActiveDirectory.DirectoryContext([System.DirectoryServices.ActiveDirectory.DirectoryContextType]::Domain, $Domain)
		
		try
		{
			$objDc = [System.DirectoryServices.ActiveDirectory.DomainController]::FindOne($directoryContext, $ADSite, $locatorOptions)
		}
		catch
		{
			$objDc = [System.DirectoryServices.ActiveDirectory.DomainController]::FindOne($directoryContext, $locatorOptions)
		}
		
	}
	return $objDc
} #end function Find-WriteableDomainController


Function Get-TodaysDate
{
	#Begin function set Todays date format
	Get-Date -Format "dd-MM-yyyy"
} #End function fnGet-TodaysDate


Function Get-UtcTime
{
	#Begin function to get current date and time in UTC format
	[System.DateTime]::UtcNow
} #End function Get-UtcTime


Function Get-SMTPServer
{
	#Begin function to get SMTP server for AD forest
	[CmdletBinding()]
	Param
	(
		[Parameter(Mandatory = $true)]
		[string]$Domain
	)
	
	Begin { }
	Process
	{
		Switch -Wildcard ($Domain)
		{
			'*childdomain2.com' { $smtpServer = "smtprelay.childdomain2.com" }
			'*childdomain1.com' { $smtpServer = "smtprelay.childdomain1.com" }
			
			default { $smtpserver = "smtprelay@domain.com" }
		}
	}
	End
	{
		$out = [PSCustomObject] @{
			SmtpServer = $smtpServer
			Port	      = '25'
		}
		Return $out
	}
} #end function Get-SMTPServer


Function Get-MyNewCimSession
{
	[CmdletBinding()]
	Param
	(
		[Parameter(Mandatory = $true, ValueFromPipeline = $true, Position = 0)]
		[ValidateNotNullorEmpty()]
		[String]$serverName,
		[Parameter(Mandatory = $false, ValueFromPipeline = $true, Position = 1)]
		[System.Management.Automation.Credential()]
		$Credential = [System.Management.Automation.PSCredential]::Empty
		
	)
	
	BEGIN
	{
		$so = New-CimSessionOption -Protocol Dcom
		
		$sessionParams = @{
			ErrorAction    = 'Stop'
			Authentication = 'Negotiate'
		}
		
		If ($PSBoundParameters['Credential'])
		{
			$sessionParams.Credential = $Credential
		}
		
	}
	PROCESS
	{
		ForEach ($server in $serverName)
		{
			$sessionParams.ComputerName = $server
			If ((Test-NetConnection -ComputerName $server -Port 5985).TcpTestSucceeded = $true)
			{
				Try
				{
					[String]$wsManTest = "Attempting to connect to $($server) using WSMAN protocol."
					Write-Verbose -Message "Attempting to connect to $($server) using WSMAN protocol." -Verbose
					New-CimSession @sessionParams
				}
				Catch
				{
					[String]$errorMsg = $_.Exception.Message
					[String]$Message2 = "Failed to establish CIM-Session on $($serverName)"
				}
			}
			ElseIf ((Test-NetConnection -ComputerName $server -Port 445).TcpTestSucceeded = $true)
			{
				$sessionParams.SessionOption = $so
				Try
				{
					[String]$dcomTest = "Attempting to connect to $($server) using DCOM protocol."
					Write-Verbose -Message "Attempting to connect to $($server) using DCOM protocol." -Verbose
					New-CimSession @sessionParams
				}
				Catch
				{
					[String]$errorMsg = $_.Exception.Message
					[String]$Message2 = "Failed to establish CIM-Session on $($serverName)"
				}
				$sessionParams.Remove('SessionOption')
			}
		}
		
	}
	
} #end Get-MyNewCimSession


#EndRegion



#region Script
#Begin Script
$Error.Clear()
$dtmFormatString = "yyyy-MM-dd HH:mm:ss"
$dtmScriptStartTimeUTC = Get-UtcTime
$myInv = Get-MyInvocation
$scriptDir = $myInv.PSScriptRoot
$scriptName = $myInv.ScriptName

#Start Function timer, to display elapsed time for function. Uses System.Diagnostics.Stopwatch class - see here: https://msdn.microsoft.com/en-us/library/system.diagnostics.stopwatch(v=vs.110).aspx 
$stopWatch = [System.Diagnostics.Stopwatch]::StartNew()
$dtmScriptStartTimeUTC = Get-UtcTime
$dtmFormatString = "yyyy-MM-dd HH:mm:ss"
$transcriptFileName = "{0}-{1}-Transcript.txt" -f $dtmScriptStartTimeUTC.ToString("yyyy-MM-dd_HH-mm-ss"), "DCInventory"

$workingDir = "{0}\{1}" -f $scriptDir, "workingDir"
Test-PathExists -Path $workingDir -PathType Folder
$transcriptDir = "{0}\{1}" -f $logFolder, "Transcripts"
Test-PathExists -Path $transcriptDir -PathType Folder
Test-PathExists -Path $reportFldr -PathType Folder
Test-PathExists -Path $logFolder -PathType Folder
$DCRptFldr = "{0}{1}\{2}" -f $driveRoot, "Reports", "DCInventories"
Test-PathExists -Path $DCRptFldr -PathType Folder

# Start transcript file
Start-Transcript ("{0}\{1}" -f $transcriptDir, $transcriptFileName)
Write-Verbose -Message ("[{0} UTC] [SCRIPT] Beginning execution of script." -f $dtmScriptStartTimeUTC.ToString($dtmFormatString)) -Verbose
Write-Verbose -Message ("[{0} UTC] [SCRIPT] Script Name:  {1}" -f $(Get-UtcTime).ToString($dtmFormatString), $scriptName) -Verbose
Write-Verbose -Message ("[{0} UTC] [SCRIPT] Log directory path:  {1}" -f $(Get-UtcTime).ToString($dtmFormatString), $logFolder) -Verbose
Write-Verbose -Message ("[{0} UTC] [SCRIPT] Report directory path:  {1}" -f $(Get-UtcTime).ToString($dtmFormatString), $DCRptFldr) -Verbose
Write-Verbose -Message ("[{0} UTC] [SCRIPT] Working directory path:  {1}" -f $(Get-UtcTime).ToString($dtmFormatString), $workingDir) -Verbose
Write-Verbose -Message ("[{0} UTC] [SCRIPT] Transcript directory path:  {1}" -f $(Get-UtcTime).ToString($dtmFormatString), $transcriptDir) -Verbose

$objComputerDomain = [System.DirectoryServices.ActiveDirectory.Domain]::GetComputerDomain()
$AuthenticationType = [System.DirectoryServices.AuthenticationTypes]::Signing -bor [System.DirectoryServices.AuthenticationTypes]::Sealing -bor [System.DirectoryServices.AuthenticationTypes]::Secure
$forestName = $objComputerDomain.Forest.Name


Write-Verbose ("[{0} UTC] [SCRIPT] Beginning execution of script." -f $dtmScriptStartTimeUTC.ToString($dtmFormatString)) -Verbose
Write-Verbose ("[{0} UTC] [SCRIPT] Script Name:  {1}" -f $(Get-UtcTime).ToString($dtmFormatString), $scriptName) -Verbose


$totalDCs = 0


$searchScript = {
	
	$domain = $_
	
	try
	{
		$objDc = $domain.FindDomainController([System.DirectoryServices.ActiveDirectory.ActiveDirectorySite]::GetComputerSite().Name.ToString(), [System.DirectoryServices.ActiveDirectory.LocatorOptions]::WriteableRequired)
	}
	catch
	{
		$objDc = Find-WritableDomainController -Domain $domain.Name.ToString() -Credential $Credential
	}
	
	Write-Verbose ("[{0} UTC] Searching for nTDSDSA objects using domain controller:  {1}" -f $(Get-UtcTime).ToString($dtmFormatString), $objDc.Name.ToString()) -Verbose
	$rootDsePath = "LDAP://{0}/rootDSE" -f $objDc.Name.ToString()
	$rootDse = New-Object System.DirectoryServices.DirectoryEntry($rootDsePath)
	$rootDse.psbase.AuthenticationType = $AuthenticationType
	
	Write-Verbose ("[{0} UTC] Searching root:  {1}" -f $(Get-UtcTime).ToString($dtmFormatString), $SearchRoot) -Verbose
	$getAdObjectParams = @{
		DomainController = $objDc.Name.ToString()
		SearchRoot	  = "CN=Sites,{0}" -f $rootDse.configurationNamingContext.ToString()
		SearchScope	  = "subtree"
		Filter		  = "(&(objectClass={0})(msDS-HasDomainNCs={1}))" -f "ntdsdsa", $rootDse.defaultNamingContext.ToString()
		Properties	  = "*"
		Port		       = 389
		Credential	  = $Credential
	}
	Get-AnADObject @getAdObjectParams
}

Write-Verbose ("[{0} UTC] [SCRIPT] Searching for domain controllers in each domain, please wait..." -f $(Get-UtcTime).ToString($dtmFormatString)) -Verbose
$Params = @{
	InputObject = $objComputerDomain.Forest.Domains
	ImportFunctions = $true
	ImportVariables = $true
	ScriptBlock = $searchScript
	Throttle    = 75
	RunspaceTimeout = 3600
}

$colADObjects = @(Invoke-Parallel @Params)

$Script = {
	$adObject = $_
	
	$serverDirectoryEntry = New-Object System.DirectoryServices.DirectoryEntry(($adObject.GetDirectoryEntry()).Parent)
	$serverDirectoryEntry.psbase.AuthenticationType = $AuthenticationType
	
	$dnsResult = Search-DnsNames -HostName $serverDirectoryEntry.dnsHostName.ToString() -Type A
	$srv = $serverDirectoryEntry.dnsHostName
	
	[bool]$hasCIMSession = $false
	Clear-DnsClientCache
	
	$cimParams = @{
		ServerName = $srv
	}
	
	if ($PsBoundParameters.containskey('Credential'))
	{
		$cimParams.add('Credential', $Credential)
	}
	
	try
	{
		$s = Get-MyNewCimSession @cimParams
		if ($s)
		{
			[bool]$hasCIMSession = $true
			$operatingSystem = Get-CimInstance -ClassName Win32_OperatingSystem -Namespace $ns -CimSession $s | Select-Object $osProps
			
			$buildNumber = Invoke-CimMethod -ClassName 'StdRegProv' -Namespace $ns -MethodName 'GetStringValue' -CimSession $s -Arguments @{
				'hDefKey'     = $hive
				'sSubKeyName' = $subkey
				'sValueName'  = $value
			}
			$buildNumber = $buildNumber.sValue
			
			$computerSystem = Get-CimInstance -ClassName Win32_ComputerSystem -Namespace $ns -CimSession $s | Select-Object $compProps
			
			$procInfo = Get-CimInstance -ClassName Win32_Processor -Namespace $ns -CimSession $s | Select-Object $procProps
			
			$netAdapter = Get-CimInstance -ClassName Win32_NetworkAdapterConfiguration -Namespace $ns -CimSession $s | Select-Object -Property [a-z]*
			
			Remove-CimSession $s
			$Error.Clear()
		}
		
	}
	catch
	{
		
		if ($PsBoundParameters.containskey('Credential'))
		{
			$operatingSystem = Get-WmiObject -Class Win32_OperatingSystem -Namespace $ns -ComputerName $srv -Credential $Credential | Select-Object $osProps
			
			$buildNumber = Invoke-Command -ComputerName $srv -ScriptBlock { (Get-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion').ReleaseId } -Credential $Credential
			
			$computerSystem = Get-WMIObject -Class Win32_ComputerSystem -Namespace $ns -ComputerName $srv -Credential $Credential | Select-Object $compProps
			
			$procInfo = Get-WmiObject -Class Win32_Processor -Property $procProps -ComputerName $srv -Credential $Credential | Select-Object $procProps
			
			$netAdapter = Get-WmiObject -Class Win32_NetworkAdapterConfiguration -Namespace $ns -ComputerName $srv -Credential $Credential | Select-Object -Property [a-z]*
			$Error.Clear()
		}
		else
		{
			$operatingSystem = Get-CimInstance -ClassName Win32_OperatingSystem -Namespace $ns -ComputerName $srv | Select-Object $osProps
			
			$buildNumber = Invoke-CimMethod -ClassName 'StdRegProv' -Namespace $ns -MethodName 'GetStringValue' -ComputerName $srv -Arguments @{
				'hDefKey'     = $hive
				'sSubKeyName' = $subkey
				'sValueName'  = $value
			}
			$buildNumber = $buildNumber.sValue
			
			$computerSystem = Get-CimInstance -ClassName Win32_ComputerSystem -Namespace $ns -ComputerName $srv | Select-Object $compProps
			
			$procInfo = Get-CimInstance -ClassName Win32_Processor -Namespace $ns -ComputerName $srv | Select-Object $procProps
			
			$netAdapter = Get-CimInstance -ClassName Win32_NetworkAdapterConfiguration -Namespace $ns -ComputerName $srv | Select-Object -Property [a-z]*
			
			$Error.Clear()
			
		}
		
	}
	
	$cpuType = $procInfo.Name | Select -Unique
	$cpuFamily = $procInfo.Caption | Select -Unique
	$cpuClockSpeed = $procInfo.MaxClockSpeed | Select -Unique
	$cpuAddressWidth = $procInfo.AddressWidth | Select -Unique
	$cpuPhysicalProcs = ($procInfo).Count
	$cpuPhysicalCores = ($procInfo | Measure -Property NumberOfLogicalProcessors -Sum).sum
	$TotalRAM = $operatingSystem.TotalVisibleMemorySize/1MB
	
	[String]$IPAddress = $netAdapter.IPAddress -join " "
	[String]$IPSubnet = $netAdapter.IPSubnet
	[String]$DefGW = $netAdapter.DefaultIPGateway
	# "" | Select @{n='TotalPhysicalProcessors';e={(,( gwmi Win32_Processor)).count}}, @{n='TotalPhysicalProcessorCores'; e={ (gwmi Win32_Processor | measure -Property NumberOfLogicalProcessors -sum).sum}}, @{n='TotalVirtualCPUs'; e={ (Get-VM | Get-VMProcessor | measure -Property Count -sum).sum }}, @{n='TotalVirtualCPUsInUse'; e={ (Get-VM | Where { $_.State -eq 'Running'} | Get-VMProcessor | measure -Property Count -sum).sum }}, @{n='TotalMSVMProcessors'; e={ (gwmi -ns root\virtualization\v2 MSVM_Processor).count }}, @{n='TotalMSVMProcessorsForVMs'; e={ (gwmi -ns root\virtualization\v2 MSVM_Processor -Filter "Description='Microsoft Virtual Processor'").count }}
	Write-Verbose ("[{0} UTC] Getting entry for server reference:  {1}" -f $(Get-UtcTime).ToString($dtmFormatString), $serverDirectoryEntry.ServerReference.ToString()) -Verbose
	$computerDirectoryEntry = New-Object System.DirectoryServices.DirectoryEntry("LDAP://" + $serverDirectoryEntry.ServerReference.ToString())
	
	Write-Verbose ("[{0} UTC] Getting properties:  {1}" -f $(Get-UtcTime).ToString($dtmFormatString), $computerDirectoryEntry.DistinguishedName.ToString())
	$objProperties = [PSCustomObject] @{
		DomainName = Get-FqdnFromDN -DistinguishedName $computerDirectoryEntry.DistinguishedName.ToString()
		ADSite     = (New-Object System.DirectoryServices.DirectoryEntry(New-Object System.DirectoryServices.DirectoryEntry($serverDirectoryEntry.Parent)).Parent).Name.ToString()
		ServerName = $serverDirectoryEntry.Name.ToString()
		DnsHostName = $serverDirectoryEntry.dnsHostName.ToString()
		#IPAddress = $(if ($dnsResult.Status -eq "Resolved") {$dnsResult.IPAddress.ToString()})
		IPAddress  = $IPAddress
		IPSubnet   = $IPSubnet
		DefaultGateway = $DefGW
		DnsServerSearchOrder = $netAdapter.DNSServerSearchOrder -join " "
		OperatingSystem = [String]($operatingSystem).Caption
		OSVersion  = [String]($operatingSystem).Version
		OSBuildNumber = [String]($operatingSystem).BuildNumber
		ReleaseNumber = [String]$buildNumber
		Manufacturer = [String]($computerSystem).Manufacturer
		Model	 = [String]$computerSystem.Model
		BuildNumber = [String]$buildNumber
		cpuType    = [String]$cpuType
		cpuFamily  = [String]$cpuFamily
		cpuClockSpeed = [String]$cpuClockSpeed
		cpuAddressWidth = [String]$cpuAddressWidth
		cpuLogicalProcessors = [String]$cpuPhysicalProcs
		cpuCores   = [String]$cpuPhysicalCores
		TotalRAMinGB = [Math]::Round($TotalRAM, 0)
		WhenCreatedUTC = $serverDirectoryEntry.whenCreated.ToUniversalTime().ToString($dtmFormatString)
		FromCIMSession = [bool]$hasCIMSession
	}
	
	$objProperties
}

Write-Verbose ("[{0} UTC] [SCRIPT] Iterating through collection of nTDSDSA objects, please wait..." -f $(Get-UtcTime).ToString($dtmFormatString)) -Verbose
$objParams = @{
	InputObject     = $colADObjects
	ImportFunctions = $true
	ImportVariables = $true
	ScriptBlock     = $Script
	Throttle	      = 75
	RunspaceTimeout = 3600
}

$colResults = @(Invoke-Parallel @objParams)

$stopWatch.Stop
Stop-Transcript

$outputFile = "{0}\{1}" -f $DCRptFldr, "{0}_{1}_{2}.csv" -f $dtmScriptStartTimeUTC.ToString("yyyy-MM-dd_HH-mm-ss"), "DCHardwareInventoryList", $forestName
$colResults | Sort-Object -Property DomainName, ADSite, ServerName | Export-Csv $outputFile -Append -NoTypeInformation
$totalDCs = $colADObjects.Count
Write-Host ("[{0} UTC] [SCRIPT] Total # of Domains             :  {1}" -f $(Get-UtcTime).ToString($dtmFormatString), $objComputerDomain.Forest.Domains.Count)
Write-Host ("[{0} UTC] [SCRIPT] Total # of Domain Controllers  :  {1}" -f $(Get-UtcTime).ToString($dtmFormatString), $totalDCs)

$runTime = $stopWatch.Elapsed.ToString('dd\.hh\:mm\:ss')


$Body = @"
	<p>$(Get-TodaysDate)</p>
	
	<p>IAM Infrastructure Team,</p>

	<p>Attached to this message is the latest hardware inventory of domain controllers for the $($forestName) forest.</p>

	<p>Thank you.</p>

	<p>*** This is an automatically generated email. Please do not reply. ***</p>
"@

#Generate report for admin consumption
$Admins = "me@domain.com" # List of users to email your report to (separate by comma)
$FromEmail = "noreply@domain.com"
$smtpSettings = Get-SmtpServer -Domain ([System.DirectoryServices.ActiveDirectory.Domain]::GetCurrentDomain().Name) #enter your own SMTP server DNS name / IP address here
$smtpServer = $smtpSettings.smtpServer

# Email our report out
#Send-MailMessage -from $FromEmail -to $Admins -cc $CC -subject "$($forestName) DC Hardware Inventory" -Attachments $outputFile -BodyAsHTML -Body $Body -priority Normal -smtpServer $smtpServer
Send-MailMessage -from $FromEmail -to $Admins -subject "$($forestName) DC Hardware Inventory" -Attachments $outputFile -BodyAsHTML -Body $Body -priority Normal -smtpServer $smtpServer


# script is complete
$dtmScriptStopTimeUTC = Get-UtcTime
$elapsedTime = New-TimeSpan -Start $dtmScriptStartTimeUTC -End $dtmScriptStopTimeUTC
Write-Host ("[{0} UTC] [SCRIPT] Script Complete" -f $(Get-UtcTime).ToString($dtmFormatString))
Write-Host ("[{0} UTC] [SCRIPT] Script Start Time :  {1}" -f $(Get-UtcTime).ToString($dtmFormatString), $dtmScriptStartTimeUTC.ToString($dtmFormatString))
Write-Host ("[{0} UTC] [SCRIPT] Script Stop Time  :  {1}" -f $(Get-UtcTime).ToString($dtmFormatString), $dtmScriptStopTimeUTC.ToString($dtmFormatString))
Write-Host ("[{0} UTC] [SCRIPT] Elapsed Time: {1:N0}.{2:N0}:{3:N0}:{4:N1}  (Days.Hours:Minutes:Seconds)" -f $(Get-UtcTime).ToString($dtmFormatString), $elapsedTime.Days, $elapsedTime.Hours, $elapsedTime.Minutes, $elapsedTime.Seconds)
Write-Host ("[{0} UTC] [SCRIPT] Output File:  {1}" -f $(Get-UtcTime).ToString($dtmFormatString), $outputFile)

#endregion