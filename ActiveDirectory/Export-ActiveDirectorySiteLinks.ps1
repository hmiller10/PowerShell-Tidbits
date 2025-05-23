﻿#Requires -Module  ActiveDirectory, ImportExcel, HelperFunctions
#Requires -Version 5
#Requires -RunAsAdministrator
<#
	.SYNOPSIS
		Export AD Site Link Info to Excel. Requires PowerShell module ImportExcel
	
	.DESCRIPTION
		This script is desigend to gather and report information on all Active Directory site links in a given forest.
	
	.PARAMETER ForestName
		Enter AD forest name to gather info. on.
	
	.PARAMETER Credential
		Enter PS credential to connecct to AD forest with.
	
	.EXAMPLE
		.\Export-ActiveDirectorySiteLinks.ps1
	
	.EXAMPLE
		.\Export-ActiveDirectorySiteLinks.ps1 -ForestName myForest.com -Credential (Get-Credential)
	
	.OUTPUTS
		CSV file containing relevant site link information
		Excel file containing relevant site link information
	
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
	[string[]]
	$ForestName,
	[Parameter(Position = 1,
	           HelpMessage = 'Enter credential for remote forest.')]
	[ValidateNotNull()]
	[System.Management.Automation.PsCredential][System.Management.Automation.Credential()]
	$Credential = [System.Management.Automation.PSCredential]::Empty,
	[Parameter(Mandatory = $true,
		HelpMessage = 'Specify the file output format you desire.')]
	[ValidateSet('CSV', 'Excel', IgnoreCase = $true)]
	[ValidateNotNullOrEmpty()]
	[string]
	$OutputFormat
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
	Import-Module -Name HelperFunctions -Force  -ErrorAction Stop
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
#EndRegion

#Region Global Variables

$dtSLHeadersCSV =
@"
ColumnName,DataType
"Forest Name", string
"Site Link Name",string
"Site Link DistinguishedName", string
"Site Link Transport Protocol",string
"Site Link Cost",string
"Site Link Replication Frequency",string
"Sites Included In Sitelink",string
"@
$dtmFormatString = "yyyy-MM-dd HH:mm:ss"
$dtmFileFormatString = "yyyy-MM-dd_HH-mm-ss"
#EndRegion

#Region Functions

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
	
	$ForestParams = @{
			ErrorAction = 'Stop'
		}
		
	if (($PSBoundParameters.ContainsKey('ForestName')) -and ($null -ne $PSBoundParameters["ForestName"]))
	{
		$ForestName = $ForestName -split (",")
	}
	else
	{
		$ForestName = Get-ADForest -Current LocalComputer | Select-Object -ExpandProperty Name
	}
	
	foreach ($Forest in $ForestName)
	{
		$ForestParams = @{
			Identity = $Forest
			Server = $Forest
			ErrorAction = 'Stop'
		}
		
		if (($PSBoundParameters.ContainsKey('Credential')) -and ($null -ne $PSBoundParameters["Credential"]))
		{
			$ForestParams.Add('AuthType', 'Negotiate')
			$ForestParams.Add('Credential', $Credential)
		}
		
		try
		{
			$DSForest = Get-ADForest @ForestParams
		}
		catch
		{
			$errorMessage = "{0}: {1}" -f $Error[0], $Error[0].InvocationInfo.PositionMessage
			Write-Error $errorMessage -ErrorAction Continue
		}
		
		$DSForestName = $DSForest.Name.ToString().ToUpper()
		$schemaMaster = $DSForest.schemaMaster
		
		#Create data table and add columns
		$dtSLHeaders = ConvertFrom-Csv -InputObject $dtSLHeadersCsv
		$slTblName = "tblADSiteLinkInfo"
		try
		{
			$dtSL = Add-DataTable -TableName $slTblName -ColumnArray $dtSLHeaders
		}
		catch
		{
			$errorMessage = "{0}: {1}" -f $Error[0], $Error[0].InvocationInfo.PositionMessage
			Write-Error $errorMessage -ErrorAction Continue
		}
		
		$siteLinkProps = @("Cost", "distinguishedName", "IntersiteTransportProtocol", "Name", "ReplicationFrequencyInMinutes", "SitesIncluded")
		
		$slParams = @{
			Server	    = $schemaMaster
			ErrorAction   = 'Stop'
		}
		
		if (($PSBoundParameters.ContainsKey('Credential')) -and ($null -ne $PSBoundParameters["Credential"]))
		{
			$slParams.Add('AuthType', 'Negotiate')
			$slParams.Add('Credential', $Credential)
		}
		
		try
		{
			$siteLinks = Get-ADReplicationSiteLink -Filter * -Properties $siteLinkProps @slParams
		}
		catch
		{
			$errorMessage = "{0}: {1}" -f $Error[0], $Error[0].InvocationInfo.PositionMessage
			Write-Error $errorMessage -ErrorAction Continue
		}
		
		$siteLinks.ForEach({
			$siteLinkName = [string]$_.Name
			$sitelinkDN = $_.distinguishedName
			$siteLinkTransportProtocol = [string]$_.InterSiteTransportProtocol
			$siteLinkCost = [string]$_.Cost
			$siteLinkFreq = $_.ReplicationFrequencyInMinutes
			if ($PSBoundParameters.ContainsValue('Excel'))
			{
				$sitesIncluded = [string]($_.SitesIncluded -join "`n")
			}
			elseif ($PSBoundParameters.ContainsValue('CSV'))
			{
				$sitesIncluded = [string]($_.SitesIncluded -join ";")
			}
			
			$slRow = $dtSL.NewRow()
			$slRow."Forest Name" = $DSForestName
			$slRow."Site Link Name" = $siteLinkName
			$slRow."Site Link DistinguishedName" = $sitelinkDN
			$slRow."Site Link Transport Protocol" = $siteLinkTransportProtocol
			$slRow."Site Link Cost" = $siteLinkCost
			$slRow."Site Link Replication Frequency" = $siteLinkFreq
			$slRow."Sites Included In Sitelink" = $sitesIncluded
			
			$dtSL.Rows.Add($slRow)
			
			$siteLinkName = $sitelinkDN = $siteLinkTransportProtocol = $siteLinkCost = $siteLinkCost = $siteLinkFreq = $sitesIncluded = $null
			[System.GC]::GetTotalMemory('ForceFullCollection') | Out-Null
			
		})
		
		#Save output
		$driveRoot = (Get-Location).Drive.Root
		$rptFolder = "{0}{1}" -f $driveRoot, "Reports"
		
		Test-PathExists -Path $rptFolder -PathType Folder
		
		$colToExport = $dtSLHeaders.ColumnName
		
		if ($dtSL.Rows.Count -gt 1)
		{
			Write-Verbose ("[{0} UTC] Exporting results data to CSV, please wait..." -f (Get-UTCTime).ToString($dtmFormatString))
			$outputFile = "{0}\{1}-{2}_Active_Directory_SiteLink_Info.csv" -f $rptFolder, (Get-UTCTime).ToString($dtmFileFormatString), $DSForestName
			$xlOutput = $OutputFile.ToString().Replace([System.IO.Path]::GetExtension($OutputFile), ".xlsx")
			$dtSL | Select-Object $colToExport | Export-Csv -Path $outputFile -NoTypeInformation
			
			Write-Verbose ("[{0} UTC] Exporting results data in Excel format, please wait..." -f $(Get-UTCTime).ToString($dtmFormatString))
			$wsName = "AD Site-Link Configuration"
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
				WrapText            = $true
				VerticalAlignment   = 'Bottom'
				HorizontalAlignment = 'Left'
			}
			
			$titleParams = @{
				FontColor         = 'White'
				FontSize	        = 16
				Bold		        = $true
				BackgroundColor   = 'Black'
				BackgroundPattern = 'Solid'
			}
			
			$xl = $dtSL | Select-Object $colToExport | Sort-Object -Property "Site Link Name" | Export-Excel @xlParams
			$Sheet = $xl.Workbook.Worksheets[$wsName]
			$lastRow = $Sheet.Dimension.End.Row
			
			Set-ExcelRange -Range $Sheet.Cells["A1"] -Value "$($DSForestName) Active Directory Site-Link Configuration" @titleParams
			Set-ExcelRange -Range $Sheet.Cells["A2"] @headerParams1
			Set-ExcelRange -Range $Sheet.Cells["B2:Z2"] @headerParams2
			Set-ExcelRange -Range $Sheet.Cells["A3:G$($lastRow)"] @setParams
			
			Export-Excel -ExcelPackage $xl -AutoSize -FreezePane 3, 0 -WorksheetName $wsName
		}
		else
		{
			Write-Warning -Message ("There are no Active Directory site links present in: {0}" -f $DSForestName)
		}
	}
}
catch
{
	$errorMessage = "{0}: {1}" -f $Error[0], $Error[0].InvocationInfo.PositionMessage
	Write-Error $errorMessage -ErrorAction Continue
}

#EndRegion