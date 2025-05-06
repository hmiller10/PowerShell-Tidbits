#Requires -Module ActiveDirectory, GroupPolicy, HelperFunctions, ImportExcel
#Requires -Version 5
#Requires -RunAsAdministrator
<#
	.SYNOPSIS
		Export AD Site Info to Excel. Requires PowerShell module ImportExcel
	
	.DESCRIPTION
		This script is desigend to gather and report information on all Active Directory sites
		in a given forest.
	
	.PARAMETER ForestName
		Enter AD forest name to gather info. on.
	
	.PARAMETER Credential
		Enter PS credential to connecct to AD forest with.
	
	.PARAMETER OutputFormat
		A description of the OutputFormat parameter.
	
	.EXAMPLE
		.\Export-ActiveDirectorySiteInfo.ps1
	
	.EXAMPLE
		.\Export-ActiveDirectorySiteInfo.ps1 -ForestName myForest.com -Credential (Get-Credential)
	
	.OUTPUTS
		OfficeOpenXml.ExcelPackage
	
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
	[Parameter(Position = 0,
			 HelpMessage = 'Enter AD forest name to gather info. on.')]
	[ValidateNotNullOrEmpty()]
	[string]$ForestName,
	[Parameter(Position = 1,
			 HelpMessage = 'Enter credential for remote forest.')]
	[ValidateNotNull()]
	[System.Management.Automation.PsCredential][System.Management.Automation.Credential()]
	$Credential = [System.Management.Automation.PSCredential]::Empty,
	[Parameter(Mandatory = $true,
			 Position = 2)]
	[ValidateNotNullOrEmpty()]
	[ValidateSet('CSV', 'Excel', IgnoreCase = $true)]
	[string]$OutputFormat
)

#Region Execution Policy
#Set-ExecutionPolicy -ExecutionPolicy RemoteSigned
#EndRegion

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
	Import-Module GroupPolicy -ErrorAction Stop
}
catch
{
	throw "Group Policy module could not be loaded. $($_.Exception.Message)"
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

#EndRegion

#Region Global Variables

$gpoNames = @()
$Sites = @()
$dtSiteHeadersCSV =
@"
ColumnName,DataType
"Site Name",string
"Site Location",string
"Site Links",string
"Adjacent Sites",string
"Subnets in Site",string
"Domains in Site",string
"Servers in Site",string
"Bridgehead Servers",string
"GPOs linked to Site",string
"Notes",string
"@
$dtmFormatString = "yyyy-MM-dd HH:mm:ss"
$dtmFileFormatString = "yyyy-MM-dd_HH-mm-ss"

#EndRegion

#Region Functions

function Get-GPSiteLink
{
	[CmdletBinding()]
	param
	(
		[Parameter(Position = 0, ValueFromPipeline = $True)]
		[string]$SiteName = "Default-First-Site-Name",
		[Parameter(Position = 1)]
		[string]$Domain = "myDomain.com",
		[Parameter(Position = 2)]
		[string]$Forest = "MyForest.com"
	)
	
	begin
	{
		Write-Verbose "Starting Function"
		#define the permission constants hash table
		$GPPerms = @{
			"permGPOApply"			      = 65536;
			"permGPORead"			      = 65792;
			"permGPOEdit"			      = 65793;
			"permGPOEditSecurityAndDelete" = 65794;
			"permGPOCustom"			 = 65795;
			"permWMIFilterEdit"		      = 131072;
			"permWMIFilterFullControl"     = 131073;
			"permWMIFilterCustom"	      = 131074;
			"permSOMLink"			      = 1835008;
			"permSOMLogging"		      = 1573120;
			"permSOMPlanning"		      = 1573376;
			"permSOMGPOCreate"		      = 1049600;
			"permSOMWMICreate"		      = 1049344;
			"permSOMWMIFullControl"	      = 1049345;
			"permStarterGPORead"		 = 197888;
			"permStarterGPOEdit"		 = 197889;
			"permStarterGPOFullControl"    = 197890;
			"permStarterGPOCustom"	      = 197891;
		}
		
		#define the GPMC COM Objects
		$gpm = New-Object -ComObject "GPMGMT.GPM"
		$gpmConstants = $gpm.GetConstants()
		$gpmDomain = $gpm.GetDomain($domain, "", $gpmConstants.UseAnyDC)
	} #Begin
	process
	{
		foreach ($item in $siteName)
		{
			#connect to site container
			$SiteContainer = $gpm.GetSitesContainer($forest, $domain, $null, $gpmConstants.UseAnyDC)
			Write-Verbose "Connected to site container on $($SiteContainer.domainController)"
			#get sites
			Write-Verbose "Getting $item"
			$site = $SiteContainer.GetSite($item)
			if ($site)
			{
				Write-Verbose "Getting site GPO links"
				$links = $Site.GetGPOLinks()
				if ($links)
				{
					#add the GPO name
					Write-Verbose ("Found {0} GPO links" -f ($links | measure-object).count)
					$links | Select-Object @{ Name = "Name"; Expression = { ($gpmDomain.GetGPO($_.GPOID)).DisplayName } },
									   @{ Name = "Description"; Expression = { ($gpmDomain.GetGPO($_.GPOID)).Description } }, GPOID, Enabled, Enforced, GPODomain, SOMLinkOrder, @{ Name = "SOM"; Expression = { $_.SOM.Path } }
				} #if $links
			} #if $site
		} #foreach site  
		
	} #process
	end
	{
		Write-Verbose "Finished"
	} #end
} #End function Get-GPSiteLink

function Get-Forest
{
	#Begin function to get an AD forest
<#
	.SYNOPSIS
		Returns an object representing either the current forest or specified target forest.

    .PARAMETER ForestName
		Specifies the fully qualified domain name of the target forest. If not specified, the current forest is returned.

	.PARAMETER Credential
		Specifies the username and password of an account with access to specified forest.

    .EXAMPLE
        PS> Get-Forest

    .EXAMPLE
        PS> Get-Forest -ForestName "example.com"

    .EXAMPLE
        PS> Get-Forest -ForestName "example.com" -Credential (Get-Credential)
#>
	[CmdletBinding()]
	[OutputType([System.DirectoryServices.ActiveDirectory.Forest])]
	param
	(
		[Parameter(Mandatory = $false, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true, HelpMessage = "Enter the FQDN for the target forest.")]
		[ValidateNotNullOrEmpty()]
		[string]$ForestName,
		[Parameter(Mandatory = $false)]
		[System.Management.Automation.PsCredential]$Credential
	)
	
	process
	{
		try
		{
			if (($PSBoundParameters.ContainsKey("ForestName") -eq $true) -and ($PSBoundParameters.ContainsKey("Credential") -eq $true))
			{
				$directoryContext = New-Object System.DirectoryServices.ActiveDirectory.DirectoryContext([System.DirectoryServices.ActiveDirectory.DirectoryContextType]::Forest, $PSBoundParameters["ForestName"], $PSBoundParameters["Credential"].UserName.ToString(), $PSBoundParameters["Credential"].GetNetworkCredential().Password.ToString())
				return ([System.DirectoryServices.ActiveDirectory.Forest]::GetForest($directoryContext))
			}
			elseif (($PSBoundParameters.ContainsKey("ForestName") -eq $true) -and ($PSBoundParameters.ContainsKey("Credential") -eq $false))
			{
				$directoryContext = New-Object System.DirectoryServices.ActiveDirectory.DirectoryContext([System.DirectoryServices.ActiveDirectory.DirectoryContextType]::Forest, $PSBoundParameters["ForestName"])
				return ([System.DirectoryServices.ActiveDirectory.Forest]::GetForest($directoryContext))
			}
			elseif (($PSBoundParameters.ContainsKey("ForestName") -eq $false))
			{
				return [System.DirectoryServices.ActiveDirectory.Forest]::GetCurrentForest()
			}
		}
		catch
		{
			throw $Error[0]
		}
	}
	
} #End function Get-Forest

#EndRegion






#Region Script
$Error.Clear()
try
{
	# Enable TLS 1.2 and 1.3
	try
	{
		#https://docs.microsoft.com/en-us/dotnet/api/system.net.securityprotocoltype?view=netcore-2.0#System_Net_SecurityProtocolType_SystemDefault
		if ($PSVersionTable.PSVersion.Major -lt 6 -and [Net.ServicePointManager]::SecurityProtocol -notmatch 'Tls12')
		{
			Write-Verbose -Message 'Adding support for TLS 1.2'
			[Net.ServicePointManager]::SecurityProtocol += [Net.SecurityProtocolType]::Tls12
		}
	}
	catch
	{
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
	
	#Region SiteConfig
	if ($PSBoundParameters.ContainsKey('ForestName'))
	{
		if (($PSBoundParameters.ContainsKey('Credential')) -and ($null -ne $PSBoundParameters["Credential"]))
		{
			$DSForest = Get-Forest -ForestName $ForestName -Credential $Credential
			$DSForestName = ($DSForest).Name.ToString().ToUpper()
			$Sites = $DSForest.Sites | Sort-Object -Property Name
		}
		else
		{
			$DSForest = Get-Forest -ForestName $ForestName
			$DSForestName = ($DSForest).Name.ToString().ToUpper()
			$Sites = $DSForest.Sites | Sort-Object -Property Name
		}
	}
	else
	{
		$DSForest = [System.DirectoryServices.ActiveDirectory.Forest]::GetCurrentForest()
		$DSForestName = ($DSForest).Name.ToString().ToUpper()
		$Sites = $DSForest.Sites | Sort-Object -Property Name
	}
	
	$schemaFSMO = ($DSForest).SchemaRoleOwner
	
	$dseParams = @{
		Server	  = $schemaFSMO
		ErrorAction = 'Stop'
	}
	
	if (($PSBoundParameters.ContainsKey('Credential')) -and ($null -ne $PSBoundParameters["Credential"]))
	{
		$dseParams.Add('AuthType', 'Negotiate')
		$dseParams.Add('Credential', $Credential)
	}
	
	try
	{
		$rootDSE = Get-ADRootDse @dseParams
		$rootCNC = $rootDSE.ConfigurationNamingContext
	}
	catch
	{
		$errorMessage = "{0}: {1}" -f $Error[0], $Error[0].InvocationInfo.PositionMessage
		Write-Error $errorMessage -ErrorAction Continue
	}
	
	#Create data table and add columns
	$sitesTblName = "tbl$($DSForestName)ADSites"
	$dtSiteHeaders = ConvertFrom-Csv -InputObject $dtSiteHeadersCsv
	try
	{
		$dtSites = Add-DataTable -TableName $sitesTblName -ColumnArray $dtSiteHeaders -ErrorAction Stop
	}
	catch
	{
		$errorMessage = "{0}: {1}" -f $Error[0], $Error[0].InvocationInfo.PositionMessage
		Write-Error $errorMessage -ErrorAction Continue
	}
	
	$adSiteProps = @("distinguishedName", "gPLink", "gPOptions", "Name")
	$sitesCount = 1
	foreach ($Site in $Sites)
	{
		Write-Verbose -Message "Working on AD site $($Site.Name)..."
		$SiteName = [String]$Site.Name
		$SiteLocation = [String]($Site).Location
		$sitesActivityMessage = "Gathering AD site information, please wait..."
		$sitesProcessingStatus = "Processing site {0} of {1}: {2}" -f $sitesCount, $Sites.count, $SiteName.ToString()
		$percentSitesComplete = ($sitesCount / $Sites.count * 100)
		Write-Progress -Activity $sitesActivityMessage -Status $sitesProcessingStatus -PercentComplete $percentSitesComplete -Id 1
		
		if ($PSBoundParameters.ContainsValue("CSV"))
		{
			$SiteSubnets = [String]($Site.Subnets -join " ")
			$SiteLinks = [String]($Site.SiteLinks -join " ")
			$AdjacentSites = [String]($Site.AdjacentSites -join " ")
			$SiteDomains = [String]($Site.Domains -join " ")
			$SiteServers = [String]($Site.Servers.Name -join " ")
			$BridgeHeads = [String]($Site.BridgeHeadServers -join " ")
		}
		elseif ($PSBoundParameters.ContainsValue("Excel"))
		{
			$SiteSubnets = [String]($Site.Subnets -join "`n")
			$SiteLinks = [String]($Site.SiteLinks -join "`n")
			$AdjacentSites = [String]($Site.AdjacentSites -join "`n")
			$SiteDomains = [String]($Site.Domains -join "`n")
			$SiteServers = [String]($Site.Servers.Name -join "`n")
			$BridgeHeads = [String]($Site.BridgeHeadServers -join "`n")
		}
		
		
		try
		{
			if (($PSBoundParameters.ContainsKey('Credential')) -and ($null -ne $PSBoundParameters["Credential"]))
			{
				$adSite = Get-ADObject -Filter '(objectClass -eq "site") -and (Name -eq $SiteName)' -SearchBase "CN=Sites,$($rootCNC)" -SearchScope OneLevel -ResultSetSize $null -Properties $adSiteProps -Server $schemaFSMO -AuthType Negotiate -Credential $Credential -ErrorAction SilentlyContinue | Select-Object -Property $adSiteProps
			}
			else
			{
				$adSite = Get-ADObject -Filter '(objectClass -eq "site") -and (Name -eq $SiteName)' -SearchBase "CN=Sites,$($rootCNC)" -SearchScope OneLevel -ResultSetSize $null -Properties $adSiteProps -Server $schemaFSMO -ErrorAction SilentlyContinue | Select-Object -Property $adSiteProps
			}
		}
		catch
		{
			$errorMessage = "{0}: {1}" -f $Error[0], $Error[0].InvocationInfo.PositionMessage
			Write-Error $errorMessage -ErrorAction Continue
		}
		
		$siteGpoNames = @()
		$siteGPOS = @()
		
		if ($null -ne $adSite.gpLink)
		{
			foreach ($siteDomain in $site.Domains)
			{
				$siteGPOS += Get-GPSiteLink -SiteName $SiteName -Domain $siteDomain -Forest $DSForest.Name.ToString()
			}
			
			if ($siteGPOS.Count -ge 1)
			{
				foreach ($siteGPO in $siteGPOS)
				{
					$id = ($siteGPO).GPOID
					$gpoDom = ($siteGPO).GPODomain
					try
					{
						if ($DSForestName -eq (Get-ADForest -Current LocalComputer).Name)
						{
							$gpoInfo = Get-GPO -Guid $id -Domain $gpoDom -Server $gpoDom -ErrorAction Stop
						}
						else
						{
							if (($PSBoundParameters.ContainsKey('Credential')) -and ($null -ne $PSBoundParameters["Credential"]))
							{
								$gpoInfo = Invoke-Command -ComputerName $schemaFSMO -Credential $Credential -Authentication Negotiate -ScriptBlock { Get-GPO -Guid $using:id -Domain $using:gpoDom -Server $using:gpoDom -ErrorAction Stop }
							}
							else
							{
								$gpoInfo = Invoke-Command -ComputerName $schemaFSMO -ScriptBlock { Get-GPO -Guid $using:id -Domain $using:gpoDom -Server $using:gpoDom -ErrorAction Stop }
							}
						}
						
					}
					catch
					{
						Write-Output $Site.Name
						Write-Output $id
						Write-Output $gpoDom
						$errorMessage = "{0}: {1}" -f $Error[0], $Error[0].InvocationInfo.PositionMessage
						Write-Error $errorMessage -ErrorAction Continue
					}
					
					if ([String]::IsNullOrEmpty($gpoInfo.DisplayName) -eq $false)
					{
						$gpoName = $gpoInfo.DisplayName
						$siteGpoNames += $gpoName
					}
					else
					{
						$siteGpoNames += "None"
					}
					
					$null = $siteGPO = $id = $gpoDom = $gpoInfo = $gpoName
				}
			}
			else
			{
				$siteGpoNames = "None."
			}
			
		}
		elseif ($null -eq $adSite.gpLink)
		{
			$siteGpoNames = "None."
		}
		
		$siteGpoNames = $siteGpoNames | Select-Object -Unique
		
		$siteRow = $dtSites.NewRow()
		$siteRow."Site Name" = $SiteName
		$siteRow."Site Location" = $SiteLocation
		$siteRow."Site Links" = $SiteLinks
		$siteRow."Adjacent Sites" = $AdjacentSites
		$siteRow."Subnets in Site" = $SiteSubnets
		$siteRow."Domains in Site" = $SiteDomains
		$siteRow."Servers in Site" = $SiteServers
		$siteRow."Bridgehead Servers" = $BridgeHeads
		if (($PSBoundParameters.ContainsKey('OutputFormat')) -and ($PSBoundParameters.ContainsValue("CSV")))
		{
			$siteRow."GPOs linked to Site" = $siteGpoNames -join " " | Out-String
		}
		elseif (($PSBoundParameters.ContainsKey('OutputFormat')) -and ($PSBoundParameters.ContainsValue("Excel")))
		{
			$siteRow."GPOs linked to Site" = $siteGpoNames -join "`n" | Out-String
		}
		
		$siteRow."Notes" = $null | Out-String
		
		$dtSites.Rows.Add($siteRow)
		
		$null = $Site = $SiteLocation = $siteGPOS = $SiteLinks = $SiteName = $SiteSubnets = $AdjacentSites = $SiteDomains = $SiteServers = $BridgeHeads
		$null = $adSite = $SiteGpoNames
		[GC]::Collect()
		$sitesCount++
	}
	
	Write-Progress -Activity "Done gathering AD site information for $($DSForestName)" -Status "Ready" -Completed
	
	#EndRegion
	
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
	
	if ($dtSites.Rows.Count -gt 1)
	{
		$colToExport = $dtSiteHeaders.ColumnName
		$outputFile = "{0}\{1}_{2}_Active_Directory_Site_Info.csv" -f $rptFolder, (Get-UTCTime).ToString($dtmFileFormatString), $DSForestName
		
		switch ($PSBoundParameters["OutputFormat"])
		{
			"CSV" {
				Write-Verbose ("[{0} UTC] Exporting results data to CSV, please wait..." -f (Get-UTCTime).ToString($dtmFormatString))
				$dtSites | Select-Object $colToExport | Export-Csv -Path $outputFile -NoTypeInformation
			}
			"Excel" {
				Write-Verbose -Message ("[{0} UTC] Exporting data tables to Excel spreadsheet tabs." -f (Get-UTCTime).ToString($dtmFormatString))
				$xlOutput = $OutputFile.ToString().Replace([System.IO.Path]::GetExtension($OutputFile), ".xlsx")
				$wsName = "AD Site Configuration"
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
					VerticalAlignment   = 'Bottom'
					HorizontalAlignment = 'Left'
					ErrorAction         = 'SilentlyContinue'
				}
				
				$titleParams = @{
					FontColor	        = 'White'
					FontSize	        = 16
					Bold		        = $true
					BackgroundColor   = 'Black'
					BackgroundPattern = 'Solid'
				}
				
				$xl = $dtSites | Select-Object $colToExport | Sort-Object -Property "Site Name" | Export-Excel @xlParams
				$Sheet = $xl.Workbook.Worksheets[$wsName]
				$Sheet.Cells["A1"].Value = 
				$lastRow = $Sheet.Dimension.End.Row
				
				Set-ExcelRange -Range $Sheet.Cells["A1"] -Value "$($DSForestName) Active Directory Site(s) Configuration" @titleParams
				Set-ExcelRange -Range $Sheet.Cells["A2"] @headerParams1
				Set-ExcelRange -Range $Sheet.Cells["B2:Z2"] @headerParams2
				Set-ExcelRange -Range $Sheet.Cells["A3:J$($lastRow)"] @setParams
				
				Export-Excel -ExcelPackage $xl -AutoSize -FreezePane 3, 0 -WorksheetName $wsName
			}
		} #end Switch
	} #end if $dtSites
	
}

#EndRegion