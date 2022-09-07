﻿#Requires -Version 7
#Requires -RunAsAdministrator
<#

.NOTES
	THIS CODE IS MADE AVAILABLE AS IS, WITHOUT WARRANTY OF ANY KIND. THE
	ENTIRE RISK OF THE USE OR THE RESULTS FROM THE USE OF THIS CODE REMAINS
	WITH THE USER.

.SYNOPSIS
	Export AD Site Info to Excel. Requires PowerShell module ImportExcel

.DESCRIPTION
	This script is desigend to gather and report information on all Active Directory sites
	in a given forest.

.LINK
	https://github.com/dfinke/ImportExcel

.OUTPUTS
	Excel file containing relevant site information

.EXAMPLE 
	.\Export-ActiveDirectorySiteInfo.ps1

#>
###########################################################################
#
#
# AUTHOR:  Heather Miller
#
# VERSION HISTORY: 2.0 - Added export to .CSV and updated output file naming convention.
# 
###########################################################################

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
	Import-Module ImportExcel -Force
}
catch
{
	try
	{
		$module = Get-Module -Name ImportExcel;
		$modulePath = Split-Path $module.Path;
		$psdPath = "{0}\{1}" -f $modulePath, "ImportExcel.psd1"
		Import-Module $psdPath -ErrorAction Stop
	}
	catch
	{
		throw "ImportExcel PS module could not be loaded. $($_.Exception.Message)"
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

#EndRegion

#Region Global Variables
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

$forestName = (Get-ADForest).Name.ToString().ToUpper()
$rootCNC = (Get-ADRootDSE).ConfigurationNamingContext
$rptFolder = 'E:\Reports'
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

[int32]$throttleLimit = 50
#EndRegion

#Region Functions
function Add-DataTable
{
<#
	.SYNOPSIS
		Creates PS data table with assigned name and column data
	
	.DESCRIPTION
		This function creates a [System.Data.DataTable] to store script output for reporting.
	
	.PARAMETER TableName
		A brief description to reference the data table by
	
	.PARAMETER ColumnArray
		List of column headers including ColumnName and DataType
	
	.EXAMPLE
		PS C:\> Add-DataTable -TableName <TableName> -ColumnArray <DataColumnDefinitions>
	
	.NOTES
		THIS CODE IS MADE AVAILABLE AS IS, WITHOUT WARRANTY OF ANY KIND. THE ENTIRE RISK OF 
		THE USE OR THE RESULTS FROM THE USE OF THIS CODE REMAINS WITH THE USER.
#>
	
	[CmdletBinding()]
	[OutputType([System.Data.DataTable])]
	param
	(
		[Parameter(Mandatory = $true,
				 Position = 0)]
		[ValidateNotNullOrEmpty()]
		[String]$TableName,
		#'TableName'
		[Parameter(Mandatory = $true,
				 Position = 1)]
		[ValidateNotNullOrEmpty()]
		$ColumnArray #'DataColumnDefinitions'
	)
	
	
	begin
	{
		$dt = $null
		$dt = New-Object System.Data.DataTable("$TableName")
	}
	process
	{
		foreach ($col in $ColumnArray)
		{
			[void]$dt.Columns.Add([System.Data.DataColumn]$col.ColumnName.ToString(), $col.DataType)
		}
	}
	end
	{
		Write-Output @( ,$dt)
	}
} #end function Add-DataTable

function Test-PathExists
{
<#
.SYNOPSIS
Checks if a path to a file or folder exists, and creates it if it does not exist.

.DESCRIPTION
Checks if a path to a file or folder exists, and creates it if it does not exist.

.PARAMETER Path
Full path to the file or folder to be checked

.PARAMETER PathType
Valid options are "File" and "Folder", depending on which to check.

.OUTPUTS
None

.EXAMPLE
Test-PathExists -Path "C:\temp\SomeFile.txt" -PathType File
	
.EXAMPLE
Test-PathExists -Path "C:\temp" -PathFype Folder

#>
	
[CmdletBinding(SupportsShouldProcess = $true)]
	param
	(
		[Parameter( Mandatory = $true,
				 Position = 0,
				 HelpMessage = 'Type the file system where the folder or file to check should be verified.')]
		[string]$Path,
		[Parameter(Mandatory = $true,
				 Position = 1,
				 HelpMessage = 'Specify path content as file or folder')]
		[string]$PathType
	)
	
	begin
	{
		$VerbosePreference = 'Continue';
	}
	
	process
	{
		switch ($PathType)
		{
			File
			{
				if ((Test-Path -Path $Path -PathType Leaf) -eq $true)
				{
					Write-Output ("File: {0} already exists..." -f $Path)
				}
				else
				{
					Write-Verbose -Message ("File: {0} not present, creating new file..." -f $Path)
					if ($PSCmdlet.ShouldProcess($Path, "Create file"))
					{
						[System.IO.File]::Create($Path)
					}
				}
			}
			Folder
			{
				if ((Test-Path -Path $Path -PathType Container) -eq $true)
				{
					Write-Output ("Folder: {0} already exists..." -f $Path)
				}
				else
				{
					Write-Verbose -Message ("Folder: {0} not present, creating new folder..." -f $Path)
					if ($PSCmdlet.ShouldProcess($Path, "Create folder"))
					{
						[System.IO.Directory]::CreateDirectory($Path)
					}
					
					
				}
			}
		}
	}
	
	end { }
	
}#end function Test-PathExists

function Get-UTCTime
{
<#
	.SYNOPSIS
		Get UTC Time
	
	.DESCRIPTION
		This functions returns the Universal Coordinated Date and Time. 
	
	.EXAMPLE
		PS C:\> Get-UTCTime
	
	.NOTES
		THIS CODE IS MADE AVAILABLE AS IS, WITHOUT WARRANTY OF ANY KIND. THE ENTIRE RISK OF 
		THE USE OR THE RESULTS FROM THE USE OF THIS CODE REMAINS WITH THE USER.
#>
	
	#Begin function to get current date and time in UTC format
	[System.DateTime]::UtcNow
} #end function Get-UTCTime

function Get-GPSiteLink
{
<#
	.SYNOPSIS
		function to get GPOs linked to an AD site
	
	.DESCRIPTION
		This function will return all group policy objects linked to an AD site.
	
	.PARAMETER SiteName
		Active Directory site name
	
	.PARAMETER Domain
		Active Directory Domain
	
	.PARAMETER Forest
		Active Directory Forest
	
	.EXAMPLE
		PS C:\> Get-GPSiteLink -SiteName "Default-First-Site-Name"
	
	.NOTES
		THIS CODE IS MADE AVAILABLE AS IS, WITHOUT WARRANTY OF ANY KIND. THE ENTIRE RISK OF
		THE USE OR THE RESULTS FROM THE USE OF THIS CODE REMAINS WITH THE USER.
#>
	
	[CmdletBinding()]
	param
	(
		[Parameter(Mandatory = $true,
				 Position = 0)]
		[ValidateNotNullOrEmpty()]
		[String]$SiteName,
		[Parameter(Position = 1)]
		[String]$Domain,
		[Parameter(Position = 2)]
		[String]$Forest
	)
	
	begin
	{
		$VerbosePreference = 'Continue'
		Write-Verbose "Starting function to get gpos linked to an AD site."
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
			Write-Verbose ("Found {0} sites" -f ($sites | Measure-Object).count)
			if ($site)
			{
				Write-Verbose "Getting site GPO links"
				$links = $Site.GetGPOLinks()
				if ($links)
				{
					#add the GPO name
					Write-Verbose ("Found {0} GPO links" -f ($links | Measure-Object).count)
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
} #end function Get-GPSiteLink

function Get-FQDNfromDN
{
<#
	.SYNOPSIS
		Convert DN to FQDN
	
	.DESCRIPTION
		This function converts an Active Directory distinguished name to a fully qualified domain name.
	
	.PARAMETER DistinguishedName
		AD distinguishedName
	
	.EXAMPLE
		PS C:\> Get-FQDNfromDN -DistinguishedName <ADDistinguishedName>
	
	.NOTES
		THIS CODE IS MADE AVAILABLE AS IS, WITHOUT WARRANTY OF ANY KIND. THE ENTIRE RISK OF 
		THE USE OR THE RESULTS FROM THE USE OF THIS CODE REMAINS WITH THE USER.
#>
	
	[CmdletBinding()]
	[OutputType([String])]
	param
	(
		[Parameter(Mandatory = $true)]
		[ValidateNotNullOrEmpty()]
		[string]$DistinguishedName
	)
	
	begin { }
	process
	{
		if ([string]::IsNullOrEmpty($DistinguishedName) -eq $true) { return $null }
		$domainComponents = $DistinguishedName.ToString().ToLower().Substring($DistinguishedName.ToString().ToLower().IndexOf("dc=")).Split(",")
		for ($i = 0; $i -lt $domainComponents.count; $i++)
		{
			$domainComponents[$i] = $domainComponents[$i].Substring($domainComponents[$i].IndexOf("=") + 1)
		}
		$fqdn = [string]::Join(".", $domainComponents)
	}
	end
	{
		return [string]$fqdn
	}
	
} #end function Get-FQDNfromDN

#EndRegion



#Region Script
$Error.Clear()

$dtmFileFormatString = "yyyy-MM-dd_HH-mm-ss"

#Create data table and add columns
$dtSiteHeaders = ConvertFrom-Csv -InputObject $dtSiteHeadersCsv
$sitesTblName = "$($forestName)_AD_Sites_Info"
$dtSites = Add-DataTable -TableName $sitesTblName -ColumnArray $dtSiteHeaders

#Region SiteConfig
#Begin collecting AD Site Configuration info.
$Sites = [System.DirectoryServices.ActiveDirectory.Forest]::GetCurrentForest().Sites | Sort-Object -Property Name

$getGpSiteLinkDef = ${function:Get-GPSiteLink}.ToString()

$Sites | ForEach-Object -Parallel {
	
	${function:Get-GPSiteLink} = $using:getGpSiteLinkDef
	
	$SiteName = [String]$_.Name
	$SiteLocation = [String]$_.Location
	$SCSubnets = [String]($_.Subnets -join "`n")
	$SiteLinks = [String]($_.SiteLinks -join "`n")
	$AdjacentSites = [String]($_.AdjacentSites -join "`n")
	$SiteDomains = [String]($_.Domains -join "`n")
	$SiteServers = [String]($_.Servers -join "`n")
	$BridgeHeads = [String]($_.BridgeHeadServers -join "`n")
	
	$adSite += Get-ADObject -Filter '(objectClass -eq "site") -and (Name -eq $_.Name)' -SearchBase "CN=Sites,$($using:rootCNC)" -SearchScope OneLevel -Properties name, distinguishedName, gPLink, gPOptions -ErrorAction SilentlyContinue
	$gpoNames = @()
	$siteGPOS = @()
	
	if (($adSite).gpLink -eq $null)
	{
		$gpoNames = "None."
	}
	else
	{
		foreach ($siteDomain in $_.Domains)
		{
			$siteGPOS += Get-GPSiteLink -SiteName $_.Name -Domain $siteDomain -Forest $forestName
		}
		
		foreach ($siteGPO in $siteGPOS)
		{
			$id = ($siteGPO).GPOID
			$gpoDom = ($siteGPO).GPODomain
			$gpoInfo = Get-GPO -Guid $id -Domain $gpoDom -Server $gpoDom -ErrorAction SilentlyContinue
			$gpoName = $gpoInfo.DisplayName.ToString()
			
			$gpoNames += $gpoName
			
			$siteGPO = $id = $gpoDom = $gpoInfo = $gpoName = $null
		}
	}
	
	$table = $using:dtSites
	$siteRow = $table.NewRow()
	$siteRow."Site Name" = $SiteName | Out-String
	$siteRow."Site Location" = $SiteLocation | Out-String
	$siteRow."Site Links" = $SiteLinks | Out-String
	$siteRow."Adjacent Sites" = $AdjacentSites | Out-String
	$siteRow."Subnets in Site" = $SCSubnets | Out-String
	$siteRow."Domains in Site" = $SiteDomains | Out-String
	$siteRow."Servers in Site" = $SiteServers | Out-String
	$siteRow."Bridgehead Servers" = $BridgeHeads | Out-String
	$siteRow."GPOs linked to Site" = $gpoNames -join "`n" | Out-String
	$siteRow."Notes" = $null | Out-String
	
	$table.Rows.Add($siteRow)
	
	$null = $SiteLocation = $siteGPOS = $SiteLinks = $SiteName = $SCSubnets = $AdjacentSites = $SiteDomains = $SiteServers = $BridgeHeads
	$null = $adSite = $gpoNames
	[System.GC]::GetTotalMemory('ForceFullCollection') | Out-Null
} -ThrottleLimit $throttleLimit

#EndRegion

#Save output
$driveRoot = (Get-Location).Drive.Root
$rptFolder = "{0}{1}" -f $driveRoot, "Reports"

Test-PathExists -Path $rptFolder -PathType Folder

$colToExport = $dtSiteHeaders.ColumnName

Write-Verbose ("[{0} UTC] Exporting results data to CSV, please wait..." -f $(Get-UTCTime).ToString($dtmFormatString))
$outputCSV = "{0}\{1}_{2}_Active_Directory_Site_Info.csv" -f $rptFolder, (Get-UTCTime).ToString($dtmFileFormatString), $forestName
$dtSites | Select-Object $colToExport | Export-Csv -Path $outputCSV -NoTypeInformation

Write-Verbose ("[{0} UTC] Exporting results data in Excel format, please wait..." -f $(Get-UTCTime).ToString($dtmFormatString))
$wsName = "AD Site Configuration"
$outputFile = "{0}\{1}_{2}_Active_Directory_Site_Info.xlsx" -f $rptFolder, (Get-UTCTime).ToString($dtmFileFormatString), $forestName
$ExcelParams = @{
	Path	        = $outputFile
	StartRow     = 2
	StartColumn  = 1
	AutoSize     = $true
	AutoFilter   = $true
	FreezeTopRow = $true
}

$Excel = $dtSites | Select-Object $colToExport | Sort-Object -Property "Site Name" | Export-Excel @ExcelParams -WorkSheetname $wsName -PassThru
$Sheet = $Excel.Workbook.Worksheets["AD Site Configuration"]
$totalRows = $Sheet.Dimension.Rows
Set-Format -Address $Sheet.Cells["A2:Z$($totalRows)"] -Wraptext -VerticalAlignment Bottom -HorizontalAlignment Left
Export-Excel -ExcelPackage $Excel -WorksheetName $wsName -Title "$($forestName) Active Directory Site Configuration" -TitleSize 18 -TitleBackgroundColor LightBlue -TitleFillPattern Solid

#EndRegion