#Requires -Module ActiveDirectory, ImportExcel
#Requires -Version 7
#Requires -RunAsAdministrator
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
	.\Export-ActiveDirectorySiteLinkBridges.ps1

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

#EndRegion

#Region Global Variables
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

$forestName = (Get-ADForest).Name.ToString().ToUpper()

$dtSLBHeadersCSV =
@"
ColumnName,DataType
"Site Link Bridge Name", string
"Site Link Bridge DN", string
"Site Links in Bridge", string
"@
[int32]$throttleLimit = 20
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

function Get-ReportDate
{
<#
	.SYNOPSIS
		function to get date in format yyyy-MM-dd
	
	.DESCRIPTION
		function to get date using the Get-Date cmdlet in the format yyyy-MM-dd
	
	.EXAMPLE
		PS C:\> $rptDate = Get-ReportDate
	
	.NOTES
		THIS CODE IS MADE AVAILABLE AS IS, WITHOUT WARRANTY OF ANY KIND. THE ENTIRE RISK OF 
		THE USE OR THE RESULTS FROM THE USE OF THIS CODE REMAINS WITH THE USER.
#>
	
	#Begin function get report execution date
	Get-Date -Format "yyyy-MM-dd"
} #End function Get-ReportDate

#EndRegion




#Region Script
try
{
	$Error.Clear()
	
	$dtmFormatString = "yyyy-MM-dd HH:mm:ss"
	$dtmFileFormatString = "yyyy-MM-dd_HH-mm-ss"
	
	#Create data table and add columns
	$dtSLBHeaders = ConvertFrom-Csv -InputObject $dtSLBHeadersCsv
	$slbTableName = "$($forestName)_AD_SiteLinkBridges"
	$dtSLB = Add-DataTable -TableName $slbTableName -ColumnArray $dtSLBHeaders
	
	#Region SiteLinkBridgeConfig
	
	#Begin collecting AD Site Link Bridge Configuration info.
	$SiteLinkBridges = Get-ADReplicationSiteLinkBridge -Filter * -Properties * | Sort-Object -Property Name
	
	$SiteLinkBridges | ForEach-Object -Parallel {
		$slbName = [String]$_.Name
		$slbDN = [String]$_.distinguishedName
		$slbLinksIncluded = [String]($_.SiteLinksIncluded -join "`n")
		
		$table = $using:dtSLB
		$slbRow = $table.NewRow()
		$slbRow."Site Link Bridge Name" = $slbName
		$slbRow."Site Link Bridge DN" = $slbDN
		$slbRow."Site Links In Bridge" = $slbLinksIncluded
		
		
		$table.Rows.Add($slbRow)
		
		$slbName = $slbDN = $slbLinksIncluded = $null
		[System.GC]::GetTotalMemory('ForceFullCollection') | Out-Null
		
	} -ThrottleLimit $throttleLimit
	
	$SiteLinkBridges = $null
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
	
	$colToExport = $dtSLBHeaders.ColumnName
	
	if ($dtSLB.Rows.Count -gt 1)
	{
		Write-Verbose ("[{0} UTC] Exporting results data to CSV, please wait..." -f $(Get-UTCTime).ToString($dtmFormatString))
		$outputCsv = "{0}\{1}-{2}_Active_Directory_Site_Link_Bridge_Info.csv" -f $rptFolder, $(Get-UTCTime).ToString("yyyy-MM-dd_HH-mm-ss"), $forestName
		$dtSL | Select-Object $ttColToExport | Export-Csv -Path $outputCsv -NoTypeInformation
		
		Write-Verbose ("[{0} UTC] Exporting results data in Excel format, please wait..." -f $(Get-UTCTime).ToString($dtmFormatString))
		$wsName = "AD Site-Link Bridge Config"
		$outputFile = "{0}\{1}-{2}_Active_Directory_Site_Link_Bridge_Info.xlsx" -f $rptFolder, $(Get-UTCTime).ToString("yyyy-MM-dd_HH-mm-ss"), $forestName
		
		$ExcelParams = @{
			Path	        = $outputFile
			StartRow     = 2
			StartColumn  = 1
			AutoSize     = $true
			AutoFilter   = $true
			BoldTopRow   = $true
			FreezeTopRow = $true
		}
		
		$Excel = $dtSLB | Select-Object $colToExport | Sort-Object -Property "Site Link Bridge Name" | Export-Excel @ExcelParams -WorkSheetname $wsName -PassThru
		$Sheet = $Excel.Workbook.Worksheets["AD Site-Link Bridge Config"]
		$totalRows = $Sheet.Dimension.Rows
		Set-Format -Address $Sheet.Cells["A2:Z$($totalRows)"] -Wraptext -VerticalAlignment Center -HorizontalAlignment Center
		Export-Excel -ExcelPackage $Excel -WorksheetName $wsName -Title "$($forestName) Active Directory Site-Link Bridge Configuration" -TitleSize 18 -TitleBackgroundColor LightBlue -TitleFillPattern Solid
	}
	
}

#EndRegion