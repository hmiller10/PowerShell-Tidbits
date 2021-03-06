<#

.NOTES
	THIS CODE IS MADE AVAILABLE AS IS, WITHOUT WARRANTY OF ANY KIND. THE
	ENTIRE RISK OF THE USE OR THE RESULTS FROM THE USE OF THIS CODE REMAINS
	WITH THE USER.

.SYNOPSIS
	Export AD Site Link Info to Excel. Requires PowerShell module ImportExcel

.DESCRIPTION
	This script is desigend to gather and report information on all Active Directory site links
	in a given forest.

.LINK
	https://github.com/dfinke/ImportExcel

.OUTPUTS
	Excel file containing relevant site link information

.EXAMPLE 
	.\Export-ADSiteLinkInfo.ps1

#>
###########################################################################
#
#
# AUTHOR:  Heather Miller
#
# VERSION HISTORY: 4.0
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
#EndRegion

#Region Global Variables
$adRootDSE = Get-ADRootDSE
$forestName = (Get-ADForest).Name.ToString().ToUpper()
$rootCNC = ($adRootDSE).ConfigurationNamingContext
$rptFolder = 'E:\Reports'
$dtSLHeadersCSV =
@"
ColumnName,DataType
"Site Link Name",string
"Site Link Type",string
"Site Link Cost",string
"Site Link Replication Frequency",string
"Site Link Replication Schedule",string
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

Function Add-DataTable
{
	#Begin function to dynamically build data table and add columns
	[CmdletBinding()]
	Param
	(
		[Parameter(Mandatory, Position = 0)]
		[String]$TableName,
		[Parameter(Mandatory, Position = 1)]
		$ColumnArray
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
		Return, $dt
	}
} #end function Add-DataTable

#EndRegion




#Region Script
#Begin Script

#Create data table and add columns
$dtSLTblName = "$($forestName)_AD_Site_Link_Info"
$dtSLHeaders = ConvertFrom-Csv -InputObject $dtSLHeadersCsv
$dtSL = Add-DataTable -TableName $dtSLTblName -ColumnArray $dtSLHeaders

#Region SiteLinkConfig
$siteLinkProps = @("cost", "distinguishedName", "Name", "replInterval", "Schedule")
$siteLinkInfo = Get-ADObject -Filter 'objectClass -eq "siteLink"' -Searchbase $rootCNC -Property $siteLinkProps | `
Select-Object Cost, distinguishedName, Description, Name, ReplInterval, @{ Name = "Schedule"; Expression = { If ($_.Schedule) { If (($_.Schedule -Join "`n").Contains("240")) { "NonDefault" }
		Else { "24x7" } }
		Else { "24x7" } } } | `
Sort-Object -Property Name

#Create data table and add columns
$dtSLTblName = "$($forestName)_AD_Site_Link_Info"
$dtSLHeaders = ConvertFrom-Csv -InputObject $dtSLHeadersCsv
$dtSL = Add-DataTable -TableName $dtSLTblName -ColumnArray $dtSLHeaders

$siteLinkCount = 1
ForEach ($siteLink in $siteLinkInfo)
{
	$siteLinkName = $siteLink.Name
	$siteLinkType = ($siteLink).distinguishedName
	$siteLinkActivityMessage = "Gathering AD site link information, please wait..."
	$siteLinkProcessingStatus = "Processing site link {0} of {1}: {2}" -f $siteLinkCount, $siteLinkInfo.count, $siteLinkName.ToString()
	$percentSiteLinksComplete = ($siteLinkCount / $siteLinkInfo.count * 100)
	Write-Progress -Activity $siteLinkActivityMessage -Status $siteLinkProcessingStatus -PercentComplete $percentSiteLinksComplete -Id 1
	
	If ($siteLinkType -like "*CN=IP*")
	{
		$linkType = "IP"
	}
	Else
	{
		$linkType = "SMTP"
	}
	$siteLinkCost = $siteLink.Cost
	$siteLinkFreq = $siteLink.replInterval
	$siteLinkSchedule = $siteLink.Schedule
	
	$slpRow = $dtSL.NewRow()
	$slpRow."Site Link Name" = $siteLinkName | Out-String
	$slpRow."Site Link Type" = $linkType | Out-String
	$slpRow."Site Link Cost" = $siteLinkCost | Out-String
	$slpRow."Site Link Replication Frequency" = $siteLinkFreq | Out-String
	$slpRow."Site Link Replication Schedule" = $siteLinkSchedule | Out-String
	
	$dtSL.Rows.Add($slpRow)
	
	$siteLink = $siteLinkName = $siteLinkType = $linkType = $siteLinkCost = $siteLinkCost = $siteLinkFreq = $siteLinkSchedule = $null
	[GC]::Collect()
	
	$siteLinkCount++
}

Write-Progress -Activity "Done gathering AD site link information for $($forestName)" -Status "Ready" -Completed
$siteLinkProps = $siteLinkInfo = $null
#EndRegion

#Save output
Test-PathExists -Path $rptFolder -PathType Folder

$wsName = "AD Site-Link Configuration"
$outputFile = "{0}\{1}" -f $rptFolder, "$($forestName).Active_Directory_SiteLink_Info_as_of_$(Get-ReportDate).xlsx"
$ExcelParams = @{
	Path	        = $outputFile
	StartRow     = 2
	StartColumn  = 1
	AutoSize     = $true
	AutoFilter   = $true
	BoldTopRow   = $true
	FreezeTopRow = $true
}

$colToExport = $dtSLHeaders.ColumnName
$Excel = $dtSL | Select-Object $colToExport | Sort-Object -Property "Site Link Name" | Export-Excel @ExcelParams -WorkSheetname $wsName -PassThru
$Sheet = $Excel.Workbook.Worksheets["AD Site-Link Configuration"]
$totalRows = $Sheet.Dimension.Rows
Set-Format -Address $Sheet.Cells["A2:Z$($totalRows)"] -Wraptext -VerticalAlignment Center -HorizontalAlignment Center
Export-Excel -ExcelPackage $Excel -WorksheetName $wsName -Title "$($forestName) Active Directory Site-Link Configuration" -TitleSize 18 -TitleBackgroundColor LightBlue -TitleFillPattern Solid


#EndRegion