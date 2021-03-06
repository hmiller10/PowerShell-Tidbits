<#

.NOTES
	THIS CODE IS MADE AVAILABLE AS IS, WITHOUT WARRANTY OF ANY KIND. THE
	ENTIRE RISK OF THE USE OR THE RESULTS FROM THE USE OF THIS CODE REMAINS
	WITH THE USER.

.SYNOPSIS
	Export AD Site Link Bridge Info to Excel. Requires PowerShell module ImportExcel

.DESCRIPTION
	This script is desigend to gather and report information on all Active Directory site link bridges
	in a given forest.

.LINK
	https://github.com/dfinke/ImportExcel

.OUTPUTS
	Excel file containing relevant site link bridge information

.EXAMPLE 
	.\Export-ADSiteLinkBridgeInfo.ps1

#>
###########################################################################
#
#
# AUTHOR:  Heather Miller
#
# VERSION HISTORY: 3.0
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
$ADRootDSE = Get-ADRootDSE
$forestName = (Get-ADForest).Name
$rptFolder = 'E:\Reports'
$dtSLBHeadersCSV =
@"
ColumnName,DataType
"Site Link Bridge Name", string
"Site Link Bridge DN", string
"Site Links in Bridge", string
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

#EndRegion






#Region Script
#Begin Script

#Region SiteLinkBridgeConfig

#Create data table and add columns
$dtSLBHeaders = ConvertFrom-Csv -InputObject $dtSLBHeadersCsv
$dtSLB = New-Object System.Data.DataTable "$($forestName) Site Link Bridge Info"

ForEach ($Header in $dtSLBHeaders)
{
	[void]$dtSLB.Columns.Add([System.Data.DataColumn]$Header.ColumnName.ToString(), $Header.DataType)
}

#Begin collecting AD Site Link Bridge Configuration info.
$SiteLinkBridges = Get-ADReplicationSiteLinkBridge -Filter * | Sort-Object -Property Name

$slbCount = 1
ForEach ($slb in $SiteLinkBridges)
{
	$slbName = [String]$slb.Name
	$slbDN = [String]$slb.distinguishedName
	$slbActivityMessage = "Gathering AD site link bridge information, please wait..."
	$slbProcessingStatus = "Processing site link bridge {0} of {1}: {2}" -f $slbCount, $SiteLinkBridges.count, $slbName.ToString()
	$percentSLBComplete = ($sLBCount / $siteLinkBridges.count * 100)
	Write-Progress -Activity $slbActivityMessage -Status $slbProcessingStatus -PercentComplete $percentSLBComplete -Id 1
	
	$slbLinksIncluded = [String]($slb.SiteLinksIncluded -join "`n")
	
	$slbRow = $dtSLB.NewRow()
	$slbRow."Site Link Bridge Name" = $slbName
	$slbRow."Site Link Bridge DN" = $slbDN
	$slbRow."Site Links In Bridge" = $slbLinksIncluded
	
	
	$dtSLB.Rows.Add($slbRow)
	
	$slbName = $slbDN = $slbLinksIncluded = $null
	[GC]::Collect()
	
	$slbCount++
}

Write-Progress -Activity "Done gathering AD site bridge information for $($forestName)" -Status "Ready" -Completed
$SiteLinkBridges = $null
#EndRegion

#Save output
Test-PathExists -Path $rptFolder -PathType Folder

$wsName = "AD Site-Link Bridge Config"
$outputFile = "{0}\{1}" -f $rptFolder, "$($forestName)_Active_Directory_Site_Link_Bridge_Info_as_of_$(Get-ReportDate).xlsx"
$ExcelParams = @{
	Path	        = $outputFile
	StartRow     = 2
	StartColumn  = 1
	AutoSize     = $true
	AutoFilter   = $true
	BoldTopRow   = $true
	FreezeTopRow = $true
}

$colToExport = $dtSLBHeaders.ColumnName
$Excel = $dtSLB | Select-Object $colToExport | Sort-Object -Property "Site Link Bridge Name" | Export-Excel @ExcelParams -WorkSheetname $wsName -PassThru
$Sheet = $Excel.Workbook.Worksheets["AD Site-Link Bridge Config"]
$totalRows = $Sheet.Dimension.Rows
Set-Format -Address $Sheet.Cells["A2:Z$($totalRows)"] -Wraptext -VerticalAlignment Center -HorizontalAlignment Center
Export-Excel -ExcelPackage $Excel -WorksheetName $wsName -Title "$($forestName) Active Directory Site-Link Bridge Configuration" -TitleSize 18 -TitleBackgroundColor LightBlue -TitleFillPattern Solid
#EndRegion
