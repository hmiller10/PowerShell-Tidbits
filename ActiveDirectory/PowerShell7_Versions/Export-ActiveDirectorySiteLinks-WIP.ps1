#Requires -Version 7
#Requires -RunAsAdministrator
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
	.\Export-ActiveDirectorySiteLinks.ps1

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
   
   
#EndRegion

#Region Global Variables
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

$forestName = (Get-ADForest).Name.ToString().ToUpper()
$rootCNC = (Get-ADRootDSE).ConfigurationNamingContext

$dtSLHeadersCSV =
@"
ColumnName,DataType
"Site Link Name",string
"Site Link Type",string
"Site Link Cost",string
"Site Link Replication Frequency",string
"Site Link Replication Schedule",string
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

#EndRegion







#Region Script
$Error.Clear()

$dtmFormatString = "yyyy-MM-dd HH:mm:ss"
$dtmFileFormatString = "yyyy-MM-dd_HH-mm-ss"

#Create data table and add columns
$dtSLHeaders = ConvertFrom-Csv -InputObject $dtSLHeadersCsv
$slTblName = "$($forestName)_AD_Site_Link_Info"
$dtSL = Add-DataTable -TableName $slTblName -ColumnArray $dtSLHeaders

#Region SiteLinkConfig
$siteLinkProps = @("cost", "distinguishedName", "Name", "replInterval", "Schedule")
$siteLinks = Get-ADObject -Filter 'objectClass -eq "siteLink"' -Searchbase $rootCNC -Property $siteLinkProps | `
Select-Object Cost, distinguishedName, Description, Name, ReplInterval, @{ Name = "Schedule"; Expression = { If ($_.Schedule) { If (($_.Schedule -Join "`n").Contains("240")) { "NonDefault" }
			Else { "24x7" } }
		Else { "24x7" } } } | `
Sort-Object -Property Name

$siteLinks | ForEach-Object -Parallel {
	$siteLinkName = $_.Name
	$siteLinkType = $_.distinguishedName
	
	If ($siteLinkType -like "*CN=IP*")
	{
		$linkType = "IP"
	}
	Else
	{
		$linkType = "SMTP"
	}
	$siteLinkCost = $_.Cost
	$siteLinkFreq = $_.replInterval
	$siteLinkSchedule = $_.Schedule
	
	$table = $using:dtSL
	$slpRow = $table.NewRow()
	$slpRow."Site Link Name" = $siteLinkName | Out-String
	$slpRow."Site Link Type" = $linkType | Out-String
	$slpRow."Site Link Cost" = $siteLinkCost | Out-String
	$slpRow."Site Link Replication Frequency" = $siteLinkFreq | Out-String
	$slpRow."Site Link Replication Schedule" = $siteLinkSchedule | Out-String
	
	$table.Rows.Add($slpRow)
	
	$siteLinkName = $siteLinkType = $linkType = $siteLinkCost = $siteLinkCost = $siteLinkFreq = $siteLinkSchedule = $null
	[System.GC]::GetTotalMemory('ForceFullCollection') | Out-Null

} -ThrottleLimit $throttleLimit

$siteLinkProps = $null
#EndRegion

#Save output

$driveRoot = (Get-Location).Drive.Root
$rptFolder = "{0}{1}" -f $driveRoot, "Reports"

Test-PathExists -Path $rptFolder -PathType Folder

$colToExport =  $dtSLHeaders.ColumnName

Write-Verbose ("[{0} UTC] Exporting results data to CSV, please wait..." -f $(Get-UTCTime).ToString($dtmFormatString))
$outputCSV = "{0}\{1}_{2}_Active_Directory_SiteLink_Report.csv" -f $rptFolder, (Get-UTCTime).ToString($dtmFileFormatString), $forestName
$dtSL | Select-Object $colToExport | Export-Csv -Path $outputCSV -NoTypeInformation

Write-Verbose ("[{0} UTC] Exporting results data in Excel format, please wait..." -f $(Get-UTCTime).ToString($dtmFormatString))
$wsName = "AD Site-Link Configuration"
$outputFile = "{0}\{1}_{2}_Active_Directory_SiteLink_Report.xlsx" -f $rptFolder, (Get-UTCTime).ToString($dtmFileFormatString), $forestName
$ExcelParams = @{
	Path	        = $outputFile
	StartRow     = 2
	StartColumn  = 1
	AutoSize     = $true
	AutoFilter   = $true
	BoldTopRow   = $true
	FreezeTopRow = $true
}

$Excel = $dtSL | Select-Object $colToExport | Sort-Object -Property "Site Link Name" | Export-Excel @ExcelParams -WorkSheetname $wsName -PassThru
$Sheet = $Excel.Workbook.Worksheets["AD Site-Link Configuration"]
$totalRows = $Sheet.Dimension.Rows
Set-Format -Address $Sheet.Cells["A2:Z$($totalRows)"] -Wraptext -VerticalAlignment Bottom -HorizontalAlignment Left
Export-Excel -ExcelPackage $Excel -WorksheetName $wsName -Title "$($forestName) Active Directory Site-Link Configuration" -TitleSize 18 -TitleBackgroundColor LightBlue -TitleFillPattern Solid


#EndRegion