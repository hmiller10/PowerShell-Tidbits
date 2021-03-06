<#

.NOTES
	THIS CODE IS MADE AVAILABLE AS IS, WITHOUT WARRANTY OF ANY KIND. THE
	ENTIRE RISK OF THE USE OR THE RESULTS FROM THE USE OF THIS CODE REMAINS
	WITH THE USER.

.SYNOPSIS
	Export AD Forest Info to Excel

.DESCRIPTION
	This script is designed to gather and report information on an Active Directory forest.

.OUTPUTS
	Excel file containing relevant forest information

.EXAMPLE 
	.\Export-ADForestInfo.ps1
#>

###########################################################################
#
#
# AUTHOR:  Heather Miller
#
# VERSION HISTORY: 6.0 - Modified function names and set all variables to
# clear
# 
###########################################################################

#region Execution Policy
#Set-ExecutionPolicy -ExecutionPolicy RemoteSigned
#endregion

#region Modules
#Check if required module is loaded, if not load import it
Try
{
	Import-Module  ActiveDirectory -ErrorAction Stop
}
Catch
{
	Throw "Active Directory module could not be loaded. $($_.Exception.Message)";
	exit
}

Try
{
	Import-Module  ImportExcel -ErrorAction Stop
}
Catch
{
	Throw "PowerShell ImportExcel module could not be loaded. $($_.Exception.Message)";
	exit
}
#endregion

#Region Global Variables
$adRootDSE = Get-ADRootDSE
$rootCNC = ($adRootDSE).ConfigurationNamingContext

$forestHeadersCsv = 
@"
ColumnName,DataType
"Forest Name" ,string
"Forest Functional Level",string
"Forest Root Domain",string
"Domains in Forest",string
"UPN Suffixes",string
"Forest Partitions Container",string
"Forest Application Partitions",string
"Replicated Naming Contexts",string
"Schema Master FSMO Holder" ,string
"Domain Naming Master FSMO Holder" ,string
"Recycle Bin Enabled",string
"Recycle Bin Scope",string
"Recycle Bin Object Lifetime in Days",string
"@

$rptFolder = 'E:\Reports'
#EndRegion

#Region Functions

Function Get-ReportDate {#Begin function set report date format
	Get-Date -Format "yyyy-MM-dd"
} #End function Get-ReportDate

Function Test-PathExists {#Begin function to check path variable and return results
	[CmdletBinding()]
	Param
	(
	   [Parameter(Mandatory,Position=0)]
	   [String]$Path,
	   [Parameter(Mandatory,Position=1)]
	   $PathType
	)

    Switch ( $PathType )
    {
    		File
			{
		   		If ( ( Test-Path -Path $Path -PathType Leaf ) -eq $true )
				{
					#Write-Host "File: $Path already exists..." -BackgroundColor White -ForegroundColor Red
					Write-Verbose -Message "File: $Path already exists.."-Verbose
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
				If ( ( Test-Path -Path $Path -PathType Container ) -eq $true )
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
}#end function Test-PathExists

Function Add-DataTable { #Begin function to dynamically build data table and add columns
  	[CmdletBinding()]
    Param
    (
        [Parameter(Mandatory,Position=0)]
        [String]$TableName,
        [Parameter(Mandatory,Position=1)]
        $ColumnArray
    )
	
	Begin {
		$dt = $null
		$dt = New-Object System.Data.DataTable("$TableName")
	}
	Process {
		ForEach ($col in $ColumnArray)
	 	 {
			[void]$dt.Columns.Add([System.Data.DataColumn]$col.ColumnName.ToString(), $col.DataType)
		 }
	}
	End {
		Return ,$dt
	}
}#end function Add-DataTable

#EndRegion








#Region Script
#Begin Script
$Error.Clear()

	#Region Forest Config
	$Domains = @()
	
	#Get AD Forest Basic Information
	$forestProperties = @("ApplicationPartitions", "Domains", "DomainNamingMaster", "ForestMode", "Name", "RootDomain", "PartitionsContainer", "SchemaMaster", "SPNSuffixes", "UPNSuffixes")
	$Forest = Get-ADForest | Select-Object -Property $forestProperties
	$Partitions = (Get-ADReplicationConnection).ReplicatedNamingContexts | Select-Object -Unique
	$replParts = ($Partitions -join "`n")
	$forestName = ($Forest).Name.ToString().ToUpper()
	$forestFunctionalLevel = ($Forest).ForestMode.ToString().ToUpper()
	$forestRootDomain = ($Forest).RootDomain.ToString().ToUpper()
	$Domains = ($Forest).Domains
	$Domains = ($Domains -join "`n")
	$forestReplCntxt = ($Forest).PartitionsContainer
	$upnSuffixes = ($Forest).UPNSuffixes
	$upnSuffixes = $upnSuffixes -join "`n"
	$appPartitions = ($Forest).ApplicationPartitions | Select-Object -Unique
	$appPartitions -join "`n"
	$schemaFSMO = ($Forest).SchemaMaster
	$dnmFSMO = ($Forest).DomainNamingMaster
	
	$objRecBin = Get-ADOptionalFeature -Filter 'Name -like "Recycle Bin Feature"' -Properties Name, FeatureScope | Select-Object -Property Name, FeatureScope
	If ( $objRecBin.Name -ne $null ) { [bool]$recBinEnabled = $true } Else { [bool]$recBinEnabled = $false }
	
	$recBinDN = "CN=Directory Service,CN=Windows NT,CN=Services," + $rootCNC
	$recBinLifeTime = (Get-ADObject -Identity $recBinDN -Properties msDS-DeletedObjectLifeTime –Partition $rootCNC).'msDS-DeletedObjectLifeTime'
	If ( $recBinLifeTime -eq $null )
	{
		$recBinLifeTime = "Default"
	}
	
	#Create data table and add columns
	$forestTblName = "$($forestName)_Information"
	$forestHeaders = ConvertFrom-Csv -InputObject $forestHeadersCsv
	$forestTable = Add-DataTable -TableName $forestTblName -ColumnArray $forestHeaders

	$forestRow = $forestTable.NewRow()
	$forestRow."Forest Name" = $forestName
	$forestRow."Forest Functional Level" = $forestFunctionalLevel
	$forestRow."Forest Root Domain" = $forestRootDomain
	$forestRow."Domains in Forest" = [String]$Domains
	$forestRow."UPN Suffixes" = [String]$upnSuffixes
	$forestRow."Forest Partitions Container" = [String]$forestReplCntxt
	$forestRow."Forest Application Partitions" = $appPartitions | Out-String
	$forestRow."Replicated Naming Contexts" = [String]$replParts
	$forestRow."Schema Master FSMO Holder" = [String]$schemaFSMO
	$forestRow."Domain Naming Master FSMO Holder" = [String]$dnmFSMO
	$forestRow."Recycle Bin Enabled" = $recBinEnabled
	$forestRow."Recycle Bin Scope" = [String]$objRecBin.FeatureScope
	$forestRow."Recycle Bin Object Lifetime in Days" = $recBinLifeTime

	$forestTable.Rows.Add($forestRow)

	$forestFunctionalLevel = $forestRootDomain = $upnSuffixes = $forestReplCntxt = $appPartitions = $Partitions = $replParts = $null
	$schemaFSMO = $dnmFSMO = $upnSuffixes = $objRecBin = $recBinDN = $recBinEnabled = $recBinLifeTime = $null
	[GC]::Collect()
	
#EndRegion

$wsName = "AD Forest Configuration"
$outputFileName = "$($forestName)_Active_Directory_Forest_Info_as_of_$(Get-ReportDate).xlsx"
Test-PathExists -Path $rptFolder -PathType Folder
$outputFile = "{0}\{1}" -f $rptFolder, $outputFileName

$ExcelParams = @{
	Path = $outputFile
	StartRow = 2
	StartColumn = 1
	AutoSize = $true
	AutoFilter = $true
	BoldTopRow = $true
	FreezeTopRow = $true
}
$colToExport = $dtForestHeaders.ColumnName
$Excel = $forestTable | Select-Object $colToExport | Sort-Object -Property "Forest Name"  | Export-Excel @ExcelParams -WorkSheetname $wsName -PassThru
$Sheet = $Excel.Workbook.Worksheets["AD Forest Configuration"]
$totalRows = $Sheet.Dimension.Rows
Set-Format -Address $Sheet.Cells["A2:Z$($totalRows)"] -Wraptext -VerticalAlignment Center -HorizontalAlignment Center
Export-Excel -ExcelPackage $Excel -WorksheetName $wsName -Title "$($forestName) Active Directory Forest Configuration" -TitleSize 18 -TitleBackgroundColor LightBlue -TitleFillPattern Solid
#EndRegion