﻿#Requires -Module ActiveDirectory, ImportExcel
#Requires -Version 7
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
# VERSION HISTORY: 1.0
# 
###########################################################################

#Region Execution Policy
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned
#EndRegion

#Region Modules
#Check if required module is loaded, if not load import it
Try 
{
	Import-Module ActiveDirectory -ErrorAction Stop
}
Catch
{
	Try
	{
	    Import-Module C:\Windows\System32\WindowsPowerShell\v1.0\Modules\ActiveDirectory\ActiveDirectory.psd1 -ErrorAction Stop
	}
	Catch
	{
	   Throw "Active Directory module could not be loaded. $($_.Exception.Message)"
	}
	
}

Try
{
	Import-Module ImportExcel -Force
}
Catch
{
	Try
	{
		$module = Get-Module -Name ImportExcel;
		 $modulePath = Split-Path $module.Path;
		 $psdPath = "{0}\{1}" -f $modulePath, "ImportExcel.psd1"
		Import-Module $psdPath -ErrorAction Stop
	}
	Catch
	{
		Throw "ImportExcel PS module could not be loaded. $($_.Exception.Message)"
	}
}
   
Try 
{
	Import-Module GroupPolicy -ErrorAction Stop
}
Catch
{
	Throw "Group Policy module could not be loaded. $($_.Exception.Message)"
}
#EndRegion

#Region Global Variables
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

$forestName = (Get-ADForest).Name.ToString().ToUpper()
$rootCNC = (Get-ADRootDSE).ConfigurationNamingContext
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

[int32]$throttleLimit = 100
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
		[String]$TableName,  #'TableName'
		[Parameter(Mandatory = $true,
				 Position = 1)]
		[ValidateNotNullOrEmpty()]
		$ColumnArray  #'DataColumnDefinitions'
	)
	
	
	Begin
	{
		$dt = $null
		$dt = New-Object System.Data.DataTable("$TableName")
	}
	Process
	{
		ForEach ($col in $ColumnArray)
		{
			[void]$dt.Columns.Add([System.Data.DataColumn]$col.ColumnName.ToString(), $col.DataType)
		}
	}
	End
	{
		Write-Output @(,$dt)
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
	[CmdletBinding()]
	param
	(
		[Parameter(Mandatory = $true,
				 Position = 0)]
		[String]$Path,
		[Parameter(Mandatory = $true,
				 Position = 1)]
		[Object]$PathType
	)
	
	Begin { $VerbosePreference = 'Continue' }
	
	Process
	{
		Switch ($PathType)
		{
			File
			{
				If ((Test-Path -Path $Path -PathType Leaf) -eq $true)
				{
					Write-Information -MessageData "File: $Path already exists..."
				}
				Else
				{
					New-Item -Path $Path -ItemType File -Force
					Write-Verbose -Message "File: $Path not present, creating new file..."
				}
			}
			Folder
			{
				If ((Test-Path -Path $Path -PathType Container) -eq $true)
				{
					Write-Information -MessageData "Folder: $Path already exists..."
				}
				Else
				{
					New-Item -Path $Path -ItemType Directory -Force
					Write-Verbose -Message "Folder: $Path not present, creating new folder"
				}
			}
		}
	}
	
	End { }
	
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
} #End function Get-UTCTime

#EndRegion




#Region Script
$Error.Clear()

$dtmFormatString = "yyyy-MM-dd HH:mm:ss"
$dtmFileFormatString = "yyyy-MM-dd_HH-mm-ss"

#Create data table and add columns
$dtSiteHeaders = ConvertFrom-Csv -InputObject $dtSiteHeadersCsv
$sitesTblName = "$($forestName)_AD_Sites_Info"
$dtSites = Add-DataTable -TableName $sitesTblName -ColumnArray $dtSiteHeaders

#Region SiteConfig
#Begin collecting AD Site Configuration info.
$Sites = [System.DirectoryServices.ActiveDirectory.Forest]::GetCurrentForest().Sites | Sort-Object -Property Name

$Sites | ForEach-Object -Parallel {
	
#	function Get-GPSiteLink
#	{
#	<#
#		.SYNOPSIS
#			function to get GPOs linked to an AD site
#		
#		.DESCRIPTION
#			This function will return all group policy objects linked to an AD site.
#		
#		.PARAMETER SiteName
#			Active Directory site name
#		
#		.PARAMETER Domain
#			Active Directory Domain
#		
#		.PARAMETER Forest
#			Active Directory Forest
#		
#		.EXAMPLE
#			PS C:\> Get-GPSiteLink -SiteName "Default-First-Site-Name"
#		
#		.NOTES
#			THIS CODE IS MADE AVAILABLE AS IS, WITHOUT WARRANTY OF ANY KIND. THE ENTIRE RISK OF
#			THE USE OR THE RESULTS FROM THE USE OF THIS CODE REMAINS WITH THE USER.
#	#>
#		
#		[CmdletBinding()]
#		param
#		(
#			[Parameter(Mandatory = $true,
#					 Position = 0)]
#			[ValidateNotNullOrEmpty()]
#			[String]$SiteName,
#			[Parameter(Position = 1)]
#			[String]$Domain,
#			[Parameter(Position = 2)]
#			[String]$Forest
#		)
#	
#		Begin
#		{
#			Write-Verbose "Starting function to get gpos linked to an AD site."
#			#define the permission constants hash table
#			$GPPerms = @{
#				"permGPOApply"			      = 65536;
#				"permGPORead"			      = 65792;
#				"permGPOEdit"			      = 65793;
#				"permGPOEditSecurityAndDelete" = 65794;
#				"permGPOCustom"			 = 65795;
#				"permWMIFilterEdit"		      = 131072;
#				"permWMIFilterFullControl"     = 131073;
#				"permWMIFilterCustom"	      = 131074;
#				"permSOMLink"			      = 1835008;
#				"permSOMLogging"		      = 1573120;
#				"permSOMPlanning"		      = 1573376;
#				"permSOMGPOCreate"		      = 1049600;
#				"permSOMWMICreate"		      = 1049344;
#				"permSOMWMIFullControl"	      = 1049345;
#				"permStarterGPORead"		 = 197888;
#				"permStarterGPOEdit"		 = 197889;
#				"permStarterGPOFullControl"    = 197890;
#				"permStarterGPOCustom"	      = 197891;
#			}
#			
#			#define the GPMC COM Objects
#			$gpm = New-Object -ComObject "GPMGMT.GPM"
#			$gpmConstants = $gpm.GetConstants()
#			$gpmDomain = $gpm.GetDomain($domain, "", $gpmConstants.UseAnyDC)
#		} #Begin
#		Process
#		{
#			ForEach ($item in $siteName)
#			{
#				#connect to site container
#				$SiteContainer = $gpm.GetSitesContainer($forest, $domain, $null, $gpmConstants.UseAnyDC)
#				Write-Verbose "Connected to site container on $($SiteContainer.domainController)"
#				#get sites
#				Write-Verbose "Getting $item"
#				$site = $SiteContainer.GetSite($item)
#				Write-Verbose ("Found {0} sites" -f ($sites | Measure-Object).count)
#				if ($site)
#				{
#					Write-Verbose "Getting site GPO links"
#					$links = $Site.GetGPOLinks()
#					if ($links)
#					{
#						#add the GPO name
#						Write-Verbose ("Found {0} GPO links" -f ($links | Measure-Object).count)
#						$links | Select-Object @{ Name = "Name"; Expression = { ($gpmDomain.GetGPO($_.GPOID)).DisplayName } },
#									 @{ Name = "Description"; Expression = { ($gpmDomain.GetGPO($_.GPOID)).Description } }, GPOID, Enabled, Enforced, GPODomain, SOMLinkOrder, @{ Name = "SOM"; Expression = { $_.SOM.Path } }
#					} #if $links
#				} #if $site
#			} #foreach site  
#			
#		} #process
#		End
#		{
#			Write-Verbose "Finished"
#		} #end
#	} #End function Get-GPSiteLink
	
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
	$siteGPODisplayNames = @()
	
	if ($null -ne $adSite.gpLink)
	{
#		foreach ($siteDomain in $_.Domains)
#		{
#			$siteGPOS += Get-GPSiteLink -SiteName $_.Name -Domain $siteDomain -Forest $using:forestName
#		}
#		
#		foreach ($siteGPO in $siteGPOS)
#		{
#			$id = ($siteGPO).GPOID
#			$gpoDom = ($siteGPO).GPODomain
#			$gpoInfo = Get-GPO -Guid $id -Domain $gpoDom -Server $gpoDom -ErrorAction SilentlyContinue
#			$gpoName = $gpoInfo.DisplayName.ToString()
#			
#			$gpoNames += $gpoName
#			
#			$siteGPO = $id = $gpoDom = $gpoInfo = $gpoName = $null
#		}
		
		try
		{
			foreach ($SiteDomain in $_.Domains)
			{
				$siteGPONames = $adSite | Select-Object -Property *, @{
					Name	      = 'GPODisplayName'
					Expression = {
						$_.gpLink | ForEach-Object {
							-join ([adsi]"LDAP://$_").displayName
						}
					}
				}
				
				if ($? -eq $true)
				{
					$siteGPODisplayNames += $siteGPONames.GPODisplayName -join "`n"
				}
				else
				{
					$siteGPODisplayNames += (Get-GPInheritance -Target $adSite -Domain $siteDomain | `
						Select-Object -Property GpoLinks).GpoLinks | Select-Object -ExpandProperty DisplayName
				}
			}
			
			
		}
		catch
		{
			$errorMessage = "{0}: {1}" -f $Error[0], $Error[0].InvocationInfo.PositionMessage
			Write-Error $errorMessage -ErrorAction Continue
		}
		
	}
	else
	{
		$gpoNames = "None."
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
	
	$SiteLocation = $siteGPOS = $SiteLinks = $SiteName = $SCSubnets = $AdjacentSites = $SiteDomains = $SiteServers = $BridgeHeads = $null
	$adSite = $gpoNames = $null
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


Write-Verbose -Message ("[{0} UTC] Exporting data tables to Excel spreadsheet tabs." -f $(Get-UTCTime).ToString($dtmFormatString))
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
Export-Excel -ExcelPackage $Excel -WorksheetName $wsName -Title "$($forestName) Active Directory Site Configuration" -TitleSize 16 -TitleBackgroundColor LightBlue -TitleFillPattern Solid

#EndRegion