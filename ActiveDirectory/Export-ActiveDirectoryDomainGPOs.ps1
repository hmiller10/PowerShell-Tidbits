#Requires -Module ActiveDirectory, ImportExcel, GroupPolicy, HelperFunctions
#Requires -Version 5
#Requires -RunAsAdministrator
<#

	.SYNOPSIS
		Create GPO report

	.DESCRIPTION
		This script leverages external PowerShell modules to create a set of
		reposts on domain group policies for the specified domain(s)
		
	.PARAMETER DomainName
		Fully qualified domain name of domain where GPOs should be inventoried
		
	.PARAMETER Credential
		PSCredential
	
	.OUTPUTS
		OfficeOpenXml.ExcelPackage
	
	.EXAMPLE
	PS C:> .\Export-ActiveDirectoryDomainGPOs.ps1
	
	.EXAMPLE
	PS C:> .\Export-ActiveDirectoryDomainGPOs.ps1 -DomainName my.domain.com -Credential PSCredential

	.NOTES
		THIS CODE IS MADE AVAILABLE AS IS, WITHOUT WARRANTY OF ANY KIND. THE
		ENTIRE RISK OF THE USE OR THE RESULTS FROM THE USE OF THIS CODE REMAINS
		WITH THE USER.
	
	.LINK
		https://github.com/dfinke/ImportExcel
	
	.LINK
		https://www.powershellgallery.com/packages/HelperFunctions/
#>

###########################################################################
#
#
# AUTHOR:  Heather Miller
#
# VERSION HISTORY: 5.0 - Reformatted Excel output to provide cleaner report
# presentation
# 
############################################################################
[CmdletBinding()]
param
(
[Parameter(Mandatory = $false)]
[ValidateNotNullOrEmpty()]
[string[]]$DomainName,
[Parameter(Mandatory = $false,
		HelpMessage = 'Enter PS credential to connect to AD domain with.')]
[ValidateNotNull()]
[System.Management.Automation.PsCredential][System.Management.Automation.Credential()]
$Credential = [System.Management.Automation.PSCredential]::Empty
)

#Region Modules
#Check if required module is loaded, if not load import it
try
{
	Import-Module -Name ActiveDirectory -Force -ErrorAction Stop
}
catch
{
	try
	{
		Import-Module C:\Windows\System32\WindowsPowerShell\v1.0\Modules\ActiveDirectory\ActiveDirectory.psd1 -ErrorAction Stop
	}
	catch
	{
		throw "Active Directory module could not be loaded. $($_.Exception.Message)"
	}
	
}

try
{
	Import-Module GroupPolicy -SkipEditionCheck -ErrorAction Stop
}
catch
{
	try
	{
		Import-Module C:\Windows\System32\WindowsPowerShell\v1.0\Modules\GroupPolicy\GroupPolicy.psd1 -ErrorAction Stop
	}
	catch
	{
		throw "Group Policy module could not be loaded. $($_.Exception.Message)"
	}
}


try
{
	Import-Module -Name ImportExcel -Force -ErrorAction Stop
}
catch
{
	try
	{
		$moduleName = 'ImportExcel'
		$ErrorActionPreference = 'Stop';
		$module = Get-Module -ListAvailable -Name $moduleName;
		$ErrorActionPreference = 'Continue';
		$modulePath = Split-Path $module.Path;
		$psdPath = "{0}\{1}" -f $modulePath, "ImportExcel.psd1"
		Import-Module $psdPath -ErrorAction Stop
	}
	catch
	{
		Write-Error "ImportExcel PS module could not be loaded. $($_.Exception.Message)" -ErrorAction Stop
	}
}

try
{
	Import-Module -Name HelperFunctions -Force -ErrorAction Stop
}
catch
{
	try
	{
		$moduleName = 'HelperFunctions'
		$ErrorActionPreference = 'Stop';
		$module = Get-Module -ListAvailable -Name $moduleName;
		$ErrorActionPreference = 'Continue';
		$modulePath = Split-Path $module.Path;
		$psdPath = "{0}\{1}" -f $modulePath, "HelperFunctions.psd1"
		Import-Module $psdPath -ErrorAction Stop
	}
	catch
	{
		Write-Error "HelperFunctions PS module could not be loaded. $($_.Exception.Message)" -ErrorAction Stop
	}
}

#EndRegion

#region Variables

$domGpoHeadersCSV = @"
	ColumnName,DataType
	"Domain Name",string
	"GPO Name",string
	"GPO GUID",string
	"GPO Creation Time",string
	"GPO Modification Date",string
	"GPO WMIFilter",string
	"GPO Computer Configuration",string
	"GPO Computer Version Directory",string
	"GPO Computer Sysvol Version",string
	"GPO Computer Extensions",string
	"GPO User Configuration",string
	"GPO User Version",string
	"GPO User Sysvol Version",string
	"GPO User Extension",string
	"GPO Links",string
	"GPO Link Enabled",string
	"GPO Link Override",string
	"GPO Owner",string
	"GPO Inherits",string
	"GPO Groups",string
	"GPO Permission Type",string
	"GPO Permissions",string
"@
#endregion

#Region Functions

function Find-DomainController
{
<#
	.SYNOPSIS
		Returns an object representing the directory entry of an object.

    .PARAMETER Domain
		Specifies the AD domain to target.

    .PARAMETER ADSite
        Specifies the AD site to target.

	.PARAMETER Credential
		Specifies the username and password of an account with access to target domain or site.

	.EXAMPLE
        PS> Find-DomainController -Domain domain.com -Credential (Get-Credential)

	.EXAMPLE
        PS> Find-DomainController -ADSite "Default-First-Site" -Credential (Get-Credential)

#>
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
	
	try
	{
		if (($PSBoundParameters.ContainsKey("Credential") -eq $true) -and ($null -ne $PSBoundParameters["Credential"]))
		{
			$directoryContext = New-Object System.DirectoryServices.ActiveDirectory.DirectoryContext([System.DirectoryServices.ActiveDirectory.DirectoryContextType]::Domain, $Domain, $Credential.UserName.ToString(), $Credential.GetNetworkCredential().Password.ToString())
			
			if ([string]::IsNullOrEmpty($ADSite) -ne $true)
			{
				$objDc = [System.DirectoryServices.ActiveDirectory.DomainController]::FindOne($directoryContext, $ADSite, $locatorOptions)
			}
			else
			{
				$objDc = [System.DirectoryServices.ActiveDirectory.DomainController]::FindOne($directoryContext, $locatorOptions)
			}
		}
		else
		{
			$directoryContext = New-Object System.DirectoryServices.ActiveDirectory.DirectoryContext([System.DirectoryServices.ActiveDirectory.DirectoryContextType]::Domain, $Domain)
			
			if ([string]::IsNullOrEmpty($ADSite) -ne $true)
			{
				$objDc = [System.DirectoryServices.ActiveDirectory.DomainController]::FindOne($directoryContext, $ADSite, $locatorOptions)
			}
			else
			{
				$objDc = [System.DirectoryServices.ActiveDirectory.DomainController]::FindOne($directoryContext, $locatorOptions)
			}
		}
	}
	catch
	{
		throw $Error[0]
	}
	
	return $objDc
} #End function Find-DomainController

function Get-GPOInfo
{
<#
	.SYNOPSIS
		Function to return GPO properties
	
	.DESCRIPTION
		This function returns the properties of a group policy object passed into the function as a parameter
		and uses it to populate specific data into an array
	
	.PARAMETER DomainFQDN
		Fully qualified domain name EG: my.domain.com
	
	.PARAMETER gpoGUID
		Group policy GUID value
	
	.EXAMPLE
		PS C:\> Get-GPOInfo -DomainFQDN $value1 -gpoGUID $value2
	
	.NOTES
		THIS CODE IS MADE AVAILABLE AS IS, WITHOUT WARRANTY OF ANY KIND. THE ENTIRE RISK OF 
		THE USE OR THE RESULTS FROM THE USE OF THIS CODE REMAINS WITH THE USER.
#>
	
	#Begin function to get GPO properties
	[CmdletBinding()]
	param
	(
	[Parameter(Mandatory = $true,
			 ValueFromPipeline = $true,
			 Position = 0)]
	[ValidateNotNullOrEmpty()]
	$DomainFQDN,
	[Parameter(Mandatory = $true,
			 ValueFromPipeline = $true,
			 Position = 1)]
	[ValidateNotNullOrEmpty()]
	$gpoGUID
	)
	
	begin
	{
		#Gets the XML version of the GPO Report
		$GPOReport = Get-GPOReport -GUID $gpoGUID -ReportType XML -Domain $DomainFQDN
	}
	process
	{
		#Converts it to an XML variable for manipulation
		$GPOXML = [xml]$GPOReport
		
		#Create array to store info
		$GPOInfo = @()
		
		#Get's info from XML and adds to array
		#General Information
		
		#$GPODomain = $GPOXML.GPO.Domain
		$GPOInfo += , $DomainFQDN
		#$GPODomain = $GPOXML.Identifier.Domain
		#$GPOInfo += , $GPODomain
		
		$Name = $GPOXML.GPO.Name
		Write-Verbose -Message "Working on GPO $($Name)."
		$GPOInfo += , $Name
		
		$GPOGUID = $GPOXML.GPO.Identifier.Identifier.'#text'
		$GPOInfo += , $GPOGUID
		
		if (!([string]::IsNullOrEmpty($GPOXML.GPO.CreatedTime)))
		{
			[DateTime]$Created = $GPOXML.GPO.CreatedTime
			$GPOInfo += , $Created.ToString("G")
		}
		else
		{
			$Created = "No Creation Date Available."
			$GPOInfo += , $Created
		}
		
		
		[DateTime]$Modified = $GPOXML.GPO.ModifiedTime
		$GPOInfo += , $Modified.ToString("G")
		
		#WMI Filter
		if ($GPOXML.GPO.FilterName)
		{
			$WMIFilter = $GPOXML.GPO.FilterName
		}
		else
		{
			$WMIFilter = "<none>"
		}
		$GPOInfo += , $WMIFilter
		
		#Computer Configuration
		$ComputerEnabled = $GPOXML.GPO.Computer.Enabled
		$GPOInfo += , $ComputerEnabled
		
		$ComputerVerDir = $GPOXML.GPO.Computer.VersionDirectory
		$GPOInfo += , $ComputerVerDir
		
		$ComputerVerSys = $GPOXML.GPO.Computer.VersionSysvol
		$GPOInfo += , $ComputerVerSys
		
		if ($GPOXML.GPO.Computer.ExtensionData)
		{
			$ComputerExtensions = $GPOXML.GPO.Computer.ExtensionData | ForEach-Object { $_.Name }
			$ComputerExtensions = [String]::join("`n", $ComputerExtensions)
		}
		else
		{
			$ComputerExtensions = "<none>"
		}
		$GPOInfo += , $ComputerExtensions
		
		#User Configuration
		$UserEnabled = $GPOXML.GPO.User.Enabled
		$GPOInfo += , $UserEnabled
		
		$UserVerDir = $GPOXML.GPO.User.VersionDirectory
		$GPOInfo += , $UserVerDir
		
		$UserVerSys = $GPOXML.GPO.User.VersionSysvol
		$GPOInfo += , $UserVerSys
		
		if ($GPOXML.GPO.User.ExtensionData)
		{
			$UserExtensions = $GPOXML.GPO.User.ExtensionData | ForEach-Object { $_.Name }
			$UserExtensions = [string]::join("`n", $UserExtensions)
		}
		else
		{
			$UserExtensions = "<none>"
		}
		$GPOInfo += , $UserExtensions
		
		#Links
		if ($GPOXML.GPO.LinksTo)
		{
			$Links = $GPOXML.GPO.LinksTo | ForEach-Object { $_.SOMPath }
			$Links = [string]::join("`n", $Links)
			$LinksEnabled = $GPOXML.GPO.LinksTo | ForEach-Object { $_.Enabled }
			$LinksEnabled = [string]::join("`n", $LinksEnabled)
			$LinksNoOverride = $GPOXML.GPO.LinksTo | ForEach-Object { $_.NoOverride }
			$LinksNoOverride = [string]::join("`n", $LinksNoOverride)
		}
		else
		{
			$Links = "<none>"
			$LinksEnabled = "<none>"
			$LinksNoOverride = "<none>"
		}
		$GPOInfo += , $Links
		$GPOInfo += , $LinksEnabled
		$GPOInfo += , $LinksNoOverride
		
		#Security Info
		$Owner = $GPOXML.GPO.SecurityDescriptor.Owner.Name.'#text'
		$GPOInfo += , $Owner
		
		$SecurityInherits = $GPOXML.GPO.SecurityDescriptor.Permissions.InheritsFromParent
		$SecurityInherits = [string]::join("`n", $SecurityInherits)
		$GPOInfo += , $SecurityInherits
		
		$SecurityGroups = $GPOXML.GPO.SecurityDescriptor.Permissions.TrusteePermissions | ForEach-Object { $_.Trustee.Name.'#text' }
		if ($null -ne $SecurityGroups)
		{
			$SecurityGroups = [string]::join("`n", $SecurityGroups)
			$GPOInfo += , $SecurityGroups
		}
		else
		{
			$SecurityGroups = "None."
			$GPOInfo += , $SecurityGroups
		}
		
		$SecurityType = $GPOXML.GPO.SecurityDescriptor.Permissions.TrusteePermissions | ForEach-Object { $_.Type.PermissionType }
		$SecurityType = [string]::join("`n", $SecurityType)
		$GPOInfo += , $SecurityType
		
		$SecurityPerms = $GPOXML.GPO.SecurityDescriptor.Permissions.TrusteePermissions | ForEach-Object { $_.Standard.GPOGroupedAccessEnum }
		$SecurityPerms = [string]::join("`n", $SecurityPerms)
		$GPOInfo += , $SecurityPerms
	}
	end
	{
		return $GPOInfo
	}
	
} #End function Get-GPOInfo

function Get-TimeStamp
{
<#
	.SYNOPSIS
		Retrun date and time in long format
	
	.DESCRIPTION
		This function returns the current date and time in long format as a string.
	
	.EXAMPLE
		PS C:\> Get-TimeStamp
	
	.NOTES
		THIS CODE IS MADE AVAILABLE AS IS, WITHOUT WARRANTY OF ANY KIND. THE ENTIRE RISK OF 
		THE USE OR THE RESULTS FROM THE USE OF THIS CODE REMAINS WITH THE USER.
#>
	
	
	(Get-Date).ToString('yyyy-MM-dd_hh-mm-ss')
} #End function Get-TimeStamp

#EndRegion




#Region Script
$Error.Clear()
try
{
	# Enable TLS 1.2 and 1.3
	try {
		#https://docs.microsoft.com/en-us/dotnet/api/system.net.securityprotocoltype?view=netcore-2.0#System_Net_SecurityProtocolType_SystemDefault
		if ($PSVersionTable.PSVersion.Major -lt 6 -and [Net.ServicePointManager]::SecurityProtocol -notmatch 'Tls12') {
			Write-Verbose -Message 'Adding support for TLS 1.2'
			[Net.ServicePointManager]::SecurityProtocol += [Net.SecurityProtocolType]::Tls12
		}
	}
	catch {
		Write-Warning -Message 'Adding TLS 1.2 to supported security protocols was unsuccessful.'
	}

	try
	{
		$localComputer = Get-CimInstance -ClassName CIM_ComputerSystem -Namespace 'root\CIMv2' -ErrorAction Stop
	}
	catch
	{
		$errorMessage = "{0}: {1}" -f $Error[0], $Error[0].InvocationInfo.PositionMessage
		Write-Error $errorMessage -ErrorAction Continue
	}
	   
	if ($null -ne $localComputer.Name)   
	{
		if (($localComputer.Caption -match "Windows 11") -eq $true) {
			try {
				#https://docs.microsoft.com/en-us/dotnet/api/system.net.securityprotocoltype?view=netcore-2.0#System_Net_SecurityProtocolType_SystemDefault
				if ($PSVersionTable.PSVersion.Major -lt 6 -and [Net.ServicePointManager]::SecurityProtocol -notmatch 'Tls13') {
					Write-Verbose -Message 'Adding support for TLS 1.3'
					[Net.ServicePointManager]::SecurityProtocol += [Net.SecurityProtocolType]::Tls13
				}
			}
			catch {
				Write-Warning -Message 'Adding TLS 1.3 to supported security protocols was unsuccessful.'
			}
		}
		elseif (($localComputer.Caption -match "Server 2022") -eq $true) {
			try {
				#https://docs.microsoft.com/en-us/dotnet/api/system.net.securityprotocoltype?view=netcore-2.0#System_Net_SecurityProtocolType_SystemDefault
				if ($PSVersionTable.PSVersion.Major -lt 6 -and [Net.ServicePointManager]::SecurityProtocol -notmatch 'Tls13') {
					Write-Verbose -Message 'Adding support for TLS 1.3'
					[Net.ServicePointManager]::SecurityProtocol += [Net.SecurityProtocolType]::Tls13
				}
			}
			catch {
				Write-Warning -Message 'Adding TLS 1.3 to supported security protocols was unsuccessful.'
			}
		}
	}
	
	if ($null -eq ($PSBoundParameters["DomainName"]))
	{
		$DomainName = ([System.DirectoryServices.ActiveDirectory.Domain]::GetCurrentDomain().Name.ToString())
	}
	
	$dCount = 1
	foreach ($Domain in $DomainName)
	{
		$ActivityMessage = "Gathering {0} domain GPO information, please wait..." -f $Domain
		$ProcessingStatus = "Processing domain {0} of {1}: {2}" -f $dCount, $DomainName.count, $Domain
		$percentComplete = ($dCount / $DomainName.Count * 100)
		Write-Progress -Activity $ActivityMessage -Status $ProcessingStatus -PercentComplete $percentComplete -Id 1
		
		$GPOs = @()
		$domainParams = @{
			Identity    = $Domain
			ErrorAction = 'Stop'
		}
		
		if (($PSBoundParameters.ContainsKey('Credential')) -and ($null -ne $PSBoundParameters["Credential"]))
		{
			$domainParams.Add('AuthType', 'Negotiate')
			$domainParams.Add('Credential', $Credential)
		}
		
		Write-Output ("Domains to query are: {0}" -f $Domain)
		if ($Domain -ne ([System.DirectoryServices.ActiveDirectory.Domain]::GetCurrentDomain().Name.ToString()))
		{
			$dcParams = @{
				Domain = $Domain
				ErrorAction = 'Stop'
			}
			
			if (($PSBoundParameters.ContainsKey('Credential')) -and ($null -ne $PSBoundParameters["Credential"]))
			{
				$dcParams.Add('Credential', $Credential)
			}
			
			$objDC = Find-DomainController @dcParams | Select-Object -ExpandProperty Name
			
			$domainParams.Add('Server',$objDC)
			
			$cmdParams = @{
				ComputerName = $objDC
				ErrorAction = 'Stop'
			}
			
			if (($PSBoundParameters.ContainsKey('Credential')) -and ($null -ne $PSBoundParameters["Credential"]))
			{
				$cmdParams.Add('Authentication', 'Negotiate')
				$cmdParams.Add('Credential', $Credential)
			}
			
			$getGpoInfoDef = "function Get-GpoInfo { ${function:Get-GpoInfo} }"
			$domTable = Invoke-Command @cmdParams -ScriptBlock {
				param ([hashtable]$Hash, $getGpoInfoDef)
				
				. ([ScriptBlock]::Create($getGpoInfoDef));
				
				try
				{
					$domainInfo = Get-ADDomain -Identity $Hash.Identity -Server $Hash.Server -ErrorAction Stop | Select-Object -Property distinguishedName, DnsRoot, Name, pdcEmulator
				}
				catch
				{
					$errorMessage = "{0}: {1}" -f $Error[0], $Error[0].InvocationInfo.PositionMessage
					Write-Error $errorMessage -ErrorAction Stop
				}
				
				if ($null -ne $domainInfo.DistinguishedName)
				{
					$domDNS = $domainInfo.dnsRoot
					$pdcFSMO = $domainInfo.pdcEmulator
				}
				
				$GPOs = Get-GPO -Domain $domDNS -Server $pdcFSMO -All
				if ($? -eq $false)
				{
					try
					{
						$GPOs = Get-GPO -Domain $domDNS -Server $domDNS -All
					}
					catch
					{
						$errorMessage = "{0}: {1}" -f $Error[0], $Error[0].InvocationInfo.PositionMessage
						Write-Error $errorMessage -ErrorAction Continue
					}
				}
				
				if ($GPOs.Count -ge 1)
				{
					[psobject[]]$gpObjects = @()

					foreach ($gpo in $GPOs)
					{
						$currentGPO = Get-GPOInfo -DomainFQDN $domDNS -gpoGUID $gpo.ID
						
						if ($currentGPO[6] -eq 'true') { $computerConfig = "Enabled" }
						else { $computerConfig = "Disabled" }
						
						if ($currentGPO[10] -eq 'true') { $userConfig = "Enabled" }
						else { $userConfig = "Disabled" }
						
						$gpObjects += New-Object -TypeName PSCustomObject -Property ([ordered] @{
								
								"Domain Name"			    = $currentGPO[0]
								"GPO Name"			    = $currentGPO[1]
								"GPO GUID"			    = $currentGPO[2]
								"GPO Creation Time"	         = $currentGPO[3]
								"GPO Modification Date"	    = $currentGPO[4]
								"GPO WMIFilter"		    = $currentGPO[5]
								"GPO Computer Configuration" = [String]$computerConfig
								"GPO Computer Version Directory" = $currentGPO[7]
								"GPO Computer Sysvol Version" = $currentGPO[8]
								"GPO Computer Extensions"    = $currentGPO[9]
								"GPO User Configuration"     = [String]$userConfig
								"GPO User Version"		    = $currentGPO[11]
								"GPO User Sysvol Version"    = $currentGPO[12]
								"GPO User Extension"	    = $currentGPO[13]
								"GPO Links"			    = $currentGPO[14]
								"GPO Link Enabled"		    = $currentGPO[15]
								"GPO Link Override"	         = $currentGPO[16]
								"GPO Owner"			    = $currentGPO[17]
								"GPO Inherits"		         = $currentGPO[18]
								"GPO Groups"			    = $currentGPO[19]
								"GPO Permission Type"	    = $currentGPO[20]
								"GPO Permissions"		    = $currentGPO[21]
							})
						
						Write-Output $gpObjects
						$null = $currentGPO = $computerConfig = $userConfig
						
						[System.GC]::GetTotalMemory('ForceFullCollection') | Out-Null
					} #end foreach
				} #end if gpo count
				
				return $gpObjects
			} -ArgumentList ([hashtable]$domainParams, $getGpoInfoDef)
		}
		else
		{
			try
			{
				$domainInfo = Get-ADDomain @domainParams | Select-Object -Property distinguishedName, DnsRoot, Name, pdcEmulator
			}
			catch
			{
				$errorMessage = "{0}: {1}" -f $Error[0], $Error[0].InvocationInfo.PositionMessage
				Write-Error $errorMessage -ErrorAction Stop
			}
			
			if ($null -ne $domainInfo.DistinguishedName)
			{
				$domDNS = $domainInfo.dnsRoot
				$pdcFSMO = $domainInfo.pdcEmulator
			}
			
			$domTblName = "tblADDomainGpos"
			$domHeaders = ConvertFrom-Csv -InputObject $domGpoHeadersCsv
			
			try
			{
				$domTable = Add-DataTable -TableName $domTblName -ColumnArray $domHeaders -ErrorAction Stop
			}
			catch
			{
				$errorMessage = "{0}: {1}" -f $Error[0], $Error[0].InvocationInfo.PositionMessage
				Write-Error $errorMessage -ErrorAction Continue
			}
			
			$GPOs = Get-GPO -Domain $domDNS -Server $pdcFSMO -All
			if ($? -eq $false)
			{
				try
				{
					$GPOs = Get-GPO -Domain $domDNS -Server $domDNS -All
				}
				catch
				{
					$errorMessage = "{0}: {1}" -f $Error[0], $Error[0].InvocationInfo.PositionMessage
					Write-Error $errorMessage -ErrorAction Continue
				}
			}
			
			if ($GPOs.Count -ge 1)
			{
				$gpoCounter = 1
				foreach ($gpo in $GPOs)
				{
					$currentGPO = Get-GPOInfo -DomainFQDN $domDNS -gpoGUID $gpo.ID
					
					Write-Verbose -Message "Processing AD GPO $($currentGPO[1]) for domain $($domDNS)."
					$gpoActivityMessage = "Gathering domain GPO information, please wait..."
					$gpoStatus = "Processing group policy {0} of {1}: {2}" -f $gpoCounter, $GPOs.count, $currentGPO[1]
					$gpoPercentComplete = ($gpoCounter / $GPOs.count * 100)
					Write-Progress -Activity $gpoActivityMessage -Status $gpoStatus -PercentComplete $gpoPercentComplete -Id 2
					
					if ($currentGPO[6] -eq 'true') { $computerConfig = "Enabled" }
					else { $computerConfig = "Disabled" }
					
					if ($currentGPO[10] -eq 'true') { $userConfig = "Enabled" }
					else { $userConfig = "Disabled" }
					
					$dtGpoRow = $domTable.NewRow()
					$dtGpoRow."Domain Name" = $currentGPO[0]
					$dtGpoRow."GPO Name" = $currentGPO[1]
					$dtGpoRow."GPO GUID" = $currentGPO[2]
					$dtGpoRow."GPO Creation Time" = $currentGPO[3]
					$dtGpoRow."GPO Modification Date" = $currentGPO[4]
					$dtGpoRow."GPO WMIFilter" = $currentGPO[5]
					$dtGpoRow."GPO Computer Configuration" = [String]$computerConfig
					$dtGpoRow."GPO Computer Version Directory" = $currentGPO[7]
					$dtGpoRow."GPO Computer Sysvol Version" = $currentGPO[8]
					$dtGpoRow."GPO Computer Extensions" = $currentGPO[9]
					$dtGpoRow."GPO User Configuration" = [String]$userConfig
					$dtGpoRow."GPO User Version" = $currentGPO[11]
					$dtGpoRow."GPO User Sysvol Version" = $currentGPO[12]
					$dtGpoRow."GPO User Extension" = $currentGPO[13]
					$dtGpoRow."GPO Links" = $currentGPO[14]
					$dtGpoRow."GPO Link Enabled" = $currentGPO[15]
					$dtGpoRow."GPO Link Override" = $currentGPO[16]
					$dtGpoRow."GPO Owner" = $currentGPO[17]
					$dtGpoRow."GPO Inherits" = $currentGPO[18]
					$dtGpoRow."GPO Groups" = $currentGPO[19]
					$dtGpoRow."GPO Permission Type" = $currentGPO[20]
					$dtGpoRow."GPO Permissions" = $currentGPO[21]
					
					$domTable.Rows.Add($dtGpoRow)
					
					Write-Progress -Activity "Done gathering GPO info. for $domain" -Status "Ready" -Completed
					$null = $currentGPO = $computerConfig = $userConfig
					
					$gpoCounter++
					[System.GC]::GetTotalMemory('ForceFullCollection') | Out-Null
				}
			} #end if row count
		}#end else
		$dCount++
	}#end foreach $Domain
	Write-Progress -Activity "Done gathering AD domain GPO information" -Status "Ready" -Completed
}
catch
{
	$errorMessage = "{0}: {1}" -f $Error[0], $Error[0].InvocationInfo.PositionMessage
	Write-Error $errorMessage -ErrorAction Continue
}
finally
{
	#Save output
	$driveRoot = (Get-Location).Drive.Root
	$rptFolder = "{0}{1}" -f $driveRoot, "Reports"
	
	Test-PathExists -Path $rptFolder -PathType Folder
	
	if ($domTable.Rows.Count -gt 1)
	{
		Write-Verbose -Message "Exporting data tables to Excel spreadsheet tabs."
		
		$ColToExport = $domHeaders.ColumnName
		$outputFile = "{0}\{1}_Domain_GPO_Configuration.csv" -f $rptFolder, (Get-TimeStamp)
		$xlOutput = $OutputFile.ToString().Replace([System.IO.Path]::GetExtension($OutputFile), ".xlsx")
		$domTable | Select-Object -Property $ColToExport -ExcludeProperty PSComputerName, RunspaceID, PSShowComputerName | Export-Csv -Path $outputFile -NoTypeInformation
		$wsName = "AD Group Policies"
		$xlParams = @{
			Path	         = $xlOutput
			WorkSheetName = $wsName
			TableStyle    = 'Medium15'
			StartRow	    = 2
			StartColumn   = 1
			AutoSize	    = $true
			AutoFilter    = $true
			BoldTopRow    = $true
			PassThru	    = $true
		}
		
		$headerParams1 = @{
			Bold			     = $true
			VerticalAlignment   = 'Center'
			HorizontalAlignment = 'Center'
		}
		
		$headerParams2 = @{
			Bold			     = $true
			VerticalAlignment   = 'Center'
			HorizontalAlignment = 'Left'
		}
		
		$setParams = @{
			Wraptext		     = $true
			VerticalAlignment   = 'Bottom'
			HorizontalAlignment = 'Left'
		}
		
		$titleParams = @{
			FontColor         = 'White'
			FontSize	        = 16
			Bold		        = $true
			BackgroundColor   = 'Black'
			BackgroundPattern = 'Solid'
		}
		
		$xl = $domTable |  Select-Object -Property $ColToExport -ExcludeProperty PSComputerName, RunspaceID, PSShowComputerName  | Export-Excel @xlParams
		$Sheet = $xl.Workbook.Worksheets[$wsName]
		$lastRow = $siteSheet.Dimension.End.Row
	
		Set-ExcelRange -Range $Sheet.Cells["A1"] -Value "Active Directory Domain Group Policy Configuration" @titleParams
		Set-ExcelRange -Range $Sheet.Cells["A2"] @headerParams1
		Set-ExcelRange -Range $Sheet.Cells["B2:Z2"] @headerParams2
		Set-ExcelRange -Range $Sheet.Cells["A3:V$($lastRow)"] @setParams
		
		Export-Excel -ExcelPackage $xl -AutoSize -FreezePane 3, 0 -WorksheetName $wsName
	}
	
}

#endregion