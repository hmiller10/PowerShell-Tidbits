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
	.\Export-ADSiteInfo.ps1

#>
###########################################################################
#
#
# AUTHOR:  Heather Miller
#
# VERSION HISTORY: 6.0 - Improved object filtering for better performance
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
	Throw "Active Directory module could not be loaded. $($_.Exception.Message)";
	exit
}

Try
{
	Import-Module ImportExcel -ErrorAction Stop
}
Catch
{
	Throw "PowerShell ImportExcel module could not be loaded. $($_.Exception.Message)";
	exit
}

Try
{
	Import-Module GroupPolicy -ErrorAction Stop
}
Catch
{
	Throw "PowerShell Group Policy module could not be loaded. $($_.Exception.Message)";
	exit
}
#EndRegion

#Region Global Variables
$adRootDSE = Get-ADRootDSE
$forestName = (Get-ADForest).Name.ToString().ToUpper()
$rootCNC = ($adRootDSE).ConfigurationNamingContext
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
	
	Switch ($PathType)
	{
		File
		{
			If ((Test-Path -Path $Path -PathType Leaf) -eq $true)
			{
				#Write-Host "File: $Path already exists..." -BackgroundColor White -ForegroundColor Red
				Write-Verbose -Message "File: $Path already exists.." -Verbose
			}
			Else
			{
				New-Item -Path $Path -ItemType File -Force
				#Write-Host "File: $Path not present, creating new file..." -BackgroundColor Black -ForegroundColor Yellow
				Write-Verbose -Message "File: $Path not present, creating new file..." -Verbose
			}
		}
		Folder
		{
			If ((Test-Path -Path $Path -PathType Container) -eq $true)
			{
				#Write-Host "Folder: $Path already exists..." -BackgroundColor White -ForegroundColor Red
				Write-Verbose -Message "Folder: $Path already exists..." -Verbose
			}
			Else
			{
				New-Item -Path $Path -ItemType Directory -Force
				#Write-Host "Folder: $Path not present, creating new folder"
				Write-Verbose -Message "Folder: $Path not present, creating new folder" -Verbose
			}
		}
	}
} #end function Test-PathExists

Function Get-ReportDate
{
	#Begin function set report date format
	Get-Date -Format "yyyy-MM-dd"
} #End function Get-ReportDate

Function Get-GPSiteLink
{
	
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
		Write-Verbose "Starting Function" -Verbose
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
			Write-Verbose "Connected to site container on $($SiteContainer.domainController)" -Verbose
			#get sites
			Write-Verbose "Getting $item" -Verbose
			$site = $SiteContainer.GetSite($item)
			Write-Verbose ("Found {0} sites" -f ($sites | measure-object).count) -Verbose
			if ($site)
			{
				Write-Verbose "Getting site GPO links"
				$links = $Site.GetGPOLinks()
				if ($links)
				{
					#add the GPO name
					Write-Verbose ("Found {0} GPO links" -f ($links | measure-object).count) -Verbose
					$links | Select @{ Name = "Name"; Expression = { ($gpmDomain.GetGPO($_.GPOID)).DisplayName } },
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

Function Get-FqdnFromDN
{
	[CmdletBinding()]
	Param
	(
		[Parameter(Mandatory = $true)]
		[string]$DistinguishedName
	)
	
	If ([string]::IsNullOrEmpty($DistinguishedName) -eq $true) { return $null }
	$domainComponents = $DistinguishedName.ToString().ToLower().Substring($DistinguishedName.ToString().ToLower().IndexOf("dc=")).Split(",")
	For ($i = 0; $i -lt $domainComponents.count; $i++)
	{
		$domainComponents[$i] = $domainComponents[$i].Substring($domainComponents[$i].IndexOf("=") + 1)
	}
	$fqdn = [string]::Join(".", $domainComponents)
	
	Return $fqdn
} #End function Get-FqdnFromDN  

#EndRegion








#Region Script
$Error.Clear()
#Create data table and add columns
$dtSiteHeaders = ConvertFrom-Csv -InputObject $dtSiteHeadersCsv
$dtSites = New-Object System.Data.DataTable "$forestName Site Properties"

ForEach ($siteHeader in $dtSiteHeaders)
{
	[void]$dtSites.Columns.Add([System.Data.DataColumn]$siteHeader.ColumnName.ToString(), $siteHeader.DataType)
}

#Region SiteConfig
#Begin collecting AD Site Configuration info.
$Sites = [System.DirectoryServices.ActiveDirectory.Forest]::GetCurrentForest().Sites | Sort-Object -Property Name

$sitesCount = 1
ForEach ($Site in $Sites)
{
	#Write-Verbose -Message "Working on AD site $Site..." -Verbose
	$SiteName = [String]$Site.Name
	$sitesActivityMessage = "Gathering AD site information, please wait..."
	$sitesProcessingStatus = "Processing site {0} of {1}: {2}" -f $sitesCount, $Sites.count, $SiteName.ToString()
	$percentSitesComplete = ($sitesCount / $Sites.count * 100)
	Write-Progress -Activity $sitesActivityMessage -Status $sitesProcessingStatus -PercentComplete $PercentComplete -Id 1
	
	$SiteLocation = [String]($Site).Location
	$SCSubnets = [String]($Site.Subnets -join "`n")
	$SiteLinks = [String]($Site.SiteLinks -join "`n")
	$AdjacentSites = [String]($Site.AdjacentSites -join "`n")
	$SiteDomains = [String]($Site.Domains -join "`n")
	$SiteServers = [String]($Site.Servers -join "`n")
	$BridgeHeads = [String]($Site.BridgeHeadServers -join "`n")
	
	$adSite += Get-ADObject -Filter '( objectClass -eq "site") -and (Name -eq $SiteName)' -SearchBase "CN=Sites,$($rootCNC)" -SearchScope OneLevel -Properties name, distinguishedName, gPLink, gPOptions -ErrorAction SilentlyContinue
	$gpoCount = ($adSite).gpLink.count
	$gpoNames = @()
	$siteGPOS = @()
	
	If (($adSite).gpLink -eq $null)
	{
		$gpoNames = "None."
	}
	Else
	{
		ForEach ($siteDomain in ($site).Domains)
		{
			$siteGPOS += Get-GPSiteLink -SiteName $SiteName -Domain $siteDomain -Forest $forestName
		}
		
		ForEach ($siteGPO in $siteGPOS)
		{
			$id = ($siteGPO).GPOID
			$gpoDom = ($siteGPO).GPODomain
			$gpoInfo = Get-GPO -Guid $id -Domain $gpoDom -Server $gpoDom -ErrorAction SilentlyContinue
			$gpoName = $gpoInfo.DisplayName.ToString()
			
			$gpoNames += $gpoName
			
			$siteGPO = $id = $gpoDom = $gpoInfo = $gpoGUID = $gpoName = $null
		}
	}
	
	
	$siteRow = $dtSites.NewRow()
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
	
	$dtSites.Rows.Add($siteRow)
	
	$Site = $SiteLocation = $siteGPOS = $SiteLinks = $SiteName = $SCSubnets = $AdjacentSites = $SiteDomains = $SiteServers = $BridgeHeads = $null
	$adSite = $gpoNames = $null
	[GC]::Collect()
	$sitesCount++
}

Write-Progress -Activity "Done gathering AD site information for $($forestName)" -Status "Ready" -Completed
#EndRegion

#Save output
Test-PathExists -Path $rptFolder -PathType Folder

$wsName = "AD Site Configuration"
$outputFile = "{0}\{1}" -f $rptFolder, "$($forestName)_Active_Directory_Site_Info_as_of_$(Get-ReportDate).xlsx"
$ExcelParams = @{
	Path	        = $outputFile
	StartRow     = 2
	StartColumn  = 1
	AutoSize     = $true
	AutoFilter   = $true
	FreezeTopRow = $true
}

$colToExport = $dtSiteHeaders.ColumnName
$Excel = $dtSites | Select-Object $colToExport | Sort-Object -Property "Site Name" | Export-Excel @ExcelParams -WorkSheetname $wsName -PassThru
$Sheet = $Excel.Workbook.Worksheets["AD Site Configuration"]
$totalRows = $Sheet.Dimension.Rows
Set-Format -Address $Sheet.Cells["A2:Z$($totalRows)"] -Wraptext -VerticalAlignment Bottom -HorizontalAlignment Left
Export-Excel -ExcelPackage $Excel -WorksheetName $wsName -Title "$($forestName) Active Directory Site Configuration" -TitleSize 16 -TitleBackgroundColor LightBlue -TitleFillPattern Solid
#endregion