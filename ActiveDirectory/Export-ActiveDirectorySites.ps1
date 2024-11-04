#Requires -Module ActiveDirectory, GroupPolicy, HelperFunctions, ImportExcel
#Requires -Version 5
#Requires -RunAsAdministrator
<#
	.SYNOPSIS
		Export AD Site Info to Excel. Requires PowerShell module ImportExcel
	
	.DESCRIPTION
		This script is desigend to gather and report information on all Active Directory sites
		in a given forest.
	
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
# VERSION HISTORY: 7.0 added error handling and credential support
# 
###########################################################################

param
(
	[Parameter(Position = 0,
			 HelpMessage = 'Enter AD forest name to gather info. on.')]
	[ValidateNotNullOrEmpty()]
	[string]$ForestName,
	[Parameter(Position = 1,
			 HelpMessage = 'Enter PS credential to connecct to AD forest with.')]
	[ValidateNotNull()]
	[System.Management.Automation.PsCredential]
	[System.Management.Automation.Credential()]
	$Credential = [System.Management.Automation.PSCredential]::Empty,
	[Parameter(Mandatory = $true,
			 Position = 2)]
	[ValidateSet('CSV', 'Excel', IgnoreCase = $true)]
	[ValidateNotNullOrEmpty()]
	[string]$OutputFormat
)

#Region Execution Policy
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned
#EndRegion

#Region Modules
#Check if required module is loaded, if not load import it
try
{
	Import-Module ActiveDirectory -SkipEditionCheck -ErrorAction Stop
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
	Import-Module -Name ImportExcel -Force  -ErrorAction Stop
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
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
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
	Param
	(
		[Parameter(Position = 0, ValueFromPipeline = $True)]
		[string]$SiteName = "Default-First-Site-Name",
		[Parameter(Position = 1)]
		[string]$Domain = "myDomain.com",
		[Parameter(Position = 2)]
		[string]$Forest = "MyForest.com"
	)
	
	Begin
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
	Process
	{
		ForEach ($item in $siteName)
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
	End
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
	$sitesTblName = "$($DSForestName)_AD_Sites"
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
		Write-Verbose -Message "Working on AD site $($Site.Name)..." -Verbose
		$SiteName = [String]$Site.Name
		$sitesActivityMessage = "Gathering AD site information, please wait..."
		$sitesProcessingStatus = "Processing site {0} of {1}: {2}" -f $sitesCount, $Sites.count, $SiteName.ToString()
		$percentSitesComplete = ($sitesCount / $Sites.count * 100)
		Write-Progress -Activity $sitesActivityMessage -Status $sitesProcessingStatus -PercentComplete $percentSitesComplete -Id 1
		
		switch ($PSBoundParameters["OutputFormat"])
		{
			"CSV" {
				$SCSubnets = [String]($Site.Subnets -join " ")
				$SiteLinks = [String]($Site.SiteLinks -join " ")
				$AdjacentSites = [String]($Site.AdjacentSites -join " ")
				$SiteDomains = [String]($Site.Domains -join " ")
				$SiteServers = [String]($Site.Servers -join " ")
				$BridgeHeads = [String]($Site.BridgeHeadServers -join " ")
			}
			"Excel" {
				$SCSubnets = [String]($Site.Subnets -join "`n")
				$SiteLinks = [String]($Site.SiteLinks -join "`n")
				$AdjacentSites = [String]($Site.AdjacentSites -join "`n")
				$SiteDomains = [String]($Site.Domains -join "`n")
				$SiteServers = [String]($Site.Servers -join "`n")
				$BridgeHeads = [String]($Site.BridgeHeadServers -join "`n")
			}
		}
		
		try
		{
			$adSite = Get-ADObject -Filter '(objectClass -eq "site") -and (Name -eq $SiteName)' -SearchBase "CN=Sites,$($rootCNC)" -SearchScope OneLevel -Properties $adSiteProps -ErrorAction SilentlyContinue | Select-Object -Property $adSiteProps
		}
		catch
		{
			$errorMessage = "{0}: {1}" -f $Error[0], $Error[0].InvocationInfo.PositionMessage
			Write-Error $errorMessage -ErrorAction Continue
		}

		$gpoNames = @()
		$siteGPOS = @()
		
		If ($null -ne $adSite.gpLink)
		{
			ForEach ($siteDomain in $site.Domains)
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
						$gpoInfo = Get-GPO -Guid $id -Domain $gpoDom -Server $gpoDom -ErrorAction Stop
					}
					catch
					{
						$errorMessage = "{0}: {1}" -f $Error[0], $Error[0].InvocationInfo.PositionMessage
						Write-Error $errorMessage -ErrorAction Continue
					}
					
					if ($null -ne $gpoInfo.DisplayName)
					{
						$gpoName = $gpoInfo.DisplayName.ToString()
						$gpoNames += $gpoName
					}
					
					
					$siteGPO = $id = $gpoDom = $gpoInfo = $gpoName = $null
				}
			}
			else
			{
				$gpoNames = "None."
			}
			
		}
		elseif ($null -eq $adSite.gpLink)
		{
			$gpoNames = "None."
		}
		$gpoNames = $gpoNames | Select-Object -Unique
		
		$siteRow = $dtSites.NewRow()
		$siteRow."Site Name" = $SiteName | Out-String
		$siteRow."Site Location" = $SiteLocation | Out-String
		$siteRow."Site Links" = $SiteLinks | Out-String
		$siteRow."Adjacent Sites" = $AdjacentSites | Out-String
		$siteRow."Subnets in Site" = $SCSubnets | Out-String
		$siteRow."Domains in Site" = $SiteDomains | Out-String
		$siteRow."Servers in Site" = $SiteServers | Out-String
		$siteRow."Bridgehead Servers" = $BridgeHeads | Out-String
		switch ($PSBoundParameters["OutputFormat"])
		{
			"CSV" {
				$siteRow."GPOs linked to Site" = $gpoNames -join " " | Out-String
			}
			"Excel" {
				$siteRow."GPOs linked to Site" = $gpoNames -join "`n" | Out-String
			}
		}
		$siteRow."Notes" = $null | Out-String
		
		$dtSites.Rows.Add($siteRow)
		
		$Site = $SiteLocation = $siteGPOS = $SiteLinks = $SiteName = $SCSubnets = $AdjacentSites = $SiteDomains = $SiteServers = $BridgeHeads = $null
		$adSite = $gpoNames = $null
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
		
		switch ($PSBoundParameters["OutputFormat"])
		{
			"CSV" {
				Write-Verbose ("[{0} UTC] Exporting results data to CSV, please wait..." -f (Get-UTCTime).ToString($dtmFormatString))
				$outputFile = "{0}\{1}_{2}_Active_Directory_Site_Info.csv" -f $rptFolder, (Get-UTCTime).ToString($dtmFileFormatString), $DSForestName
				$dtSites | Select-Object $colToExport | Export-Csv -Path $outputFile -NoTypeInformation
			}
			"Excel" {
				Write-Verbose -Message ("[{0} UTC] Exporting data tables to Excel spreadsheet tabs." -f (Get-UTCTime).ToString($dtmFormatString))
				$wsName = "AD Site Configuration"
				$xlParams = @{
					Path	        = $outputFile = "{0}\{1}_{2}_Active_Directory_Site_Info.xlsx" -f $rptFolder, (Get-UTCTime).ToString($dtmFileFormatString), $DSForestName
					WorkSheetName = $wsName
					TableStyle = 'Medium15'
					StartRow     = 2
					StartColumn  = 1
					AutoSize   = $true
					AutoFilter   = $true
					BoldTopRow   = $true
					PassThru = $true
				}
				
				$xl = $dtSites | Select-Object $colToExport | Sort-Object -Property "Site Name" | Export-Excel @xlParams
				$Sheet = $xl.Workbook.Worksheets["AD Site Configuration"]
                	Set-ExcelRange -Range $Sheet.Cells["A2:Z2"] -WrapText -HorizontalAlignment Center -VerticalAlignment Center -AutoFit
                	$cols = $Sheet.Dimension.Columns				
                	Set-ExcelRange -Range $Sheet.Cells["A3:Z$($cols)"] -Wraptext -HorizontalAlignment Left -VerticalAlignment Bottom
				Export-Excel -ExcelPackage $xl -WorksheetName $wsName -FreezePane 3, 0 -Title "$($DSForestName.ToUpper()) Active Directory Site Configuration" -TitleBold -TitleSize 16
			}
		} #end Switch
	} #end if $dtSites
	
}

#EndRegion