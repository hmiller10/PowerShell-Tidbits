#Requires -Module  ActiveDirectory, ImportExcel, HelperFunctions
#Requires -Version 5
#Requires -RunAsAdministrator
<#
	.SYNOPSIS
		Export AD Site Link Bridge Info to Excel. Requires PowerShell module ImportExcel
	
	.DESCRIPTION
		This script is desigend to gather and report information on all Active Directory site link bridges
		in a given forest.
	
	.PARAMETER ForestName
		Active Directory forest name
	
	.PARAMETER Credential
		PowerShell credential object
	
	.EXAMPLE
		.\Export-ActiveDirectorySiteLinkBridges.ps1
		
	.EXAMPLE
		.\Export-ActiveDirectorySiteLinkBridges.ps1 -ForestName myTestForest.com -Credential (Get-Credential)
	
	.OUTPUTS
		OfficeOpenXml.ExcelPackage
	
	.NOTES
		THIS CODE IS MADE AVAILABLE AS IS, WITHOUT WARRANTY OF ANY KIND. THE
		ENTIRE RISK OF THE USE OR THE RESULTS FROM THE USE OF THIS CODE REMAINS
		WITH THE USER.
	
	.LINK
		https://www.powershellgallery.com/packages/ImportExcel/
		
	.LINK
		https://www.powershellgallery.com/packages/HelperFunctions/
		
#>

[CmdletBinding()]
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
	$Credential = [System.Management.Automation.PSCredential]::Empty
)

#Region Execution Policy
#Set-ExecutionPolicy -ExecutionPolicy RemoteSigned
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

$dtSLBHeadersCSV =
@"
ColumnName,DataType
"Site Link Bridge Name", string
"Site Link Bridge DN", string
"Site Link Transport Protocol", string
"Site Links in Bridge", string
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

	$ForestParams = @{
		ErrorAction = 'Stop'
	}
	
	if (($PSBoundParameters.ContainsKey('ForestName')) -and ($null -ne $PSBoundParameters["ForestName"]))
	{
		$ForestParams.Add('Identity', $ForestName)
		$ForestParams.Add('Server', $ForestName)
	}
	else
	{
		$ForestParams.Add('Current', 'LocalComputer')
	}
	
	if (($PSBoundParameters.ContainsKey('Credential')) -and ($null -ne $PSBoundParameters["Credential"]))
	{
		$ForestParams.Add('AuthType', 'Negotiate')
		$ForestParams.Add('Credential', $Credential)
	}
	
	try
	{
		$DSForest = Get-ADForest @ForestParams
		$DSForestName = ($DSForest).Name.ToString().ToUpper()
	}
	catch
	{
		$errorMessage = "{0}: {1}" -f $Error[0], $Error[0].InvocationInfo.PositionMessage
		Write-Error $errorMessage -ErrorAction Continue
	}

	#Create data table and add columns
	$dtSLBHeaders = ConvertFrom-Csv -InputObject $dtSLBHeadersCsv
	$slbTableName = "$($DSForestName)_AD_SiteLinkBridges"
	try
	{
		$dtSLB = Add-DataTable -TableName $slbTableName -ColumnArray $dtSLBHeaders -ErrorAction Stop
	}
	catch
	{
		$errorMessage = "{0}: {1}" -f $Error[0], $Error[0].InvocationInfo.PositionMessage
		Write-Error $errorMessage -ErrorAction Continue
	}
	
	#Region SiteLinkBridgeConfig
	$slbParams = @{
		Filter = '*'
		Properties = '*'
		Server = $DSForest.domainNamingMaster.ToString()
		ErrorAction = 'Stop'
	}
	
	#Begin collecting AD Site Link Bridge Configuration info.
	if (($PSBoundParameters.ContainsKey('Credential')) -and ($null -ne $PSBoundParameters["Credential"]))
	{
		$slbParams.Add('AuthType','Negotiate')
		$slbParams.Add('Credential',$Credential)
	}
	
	try
	{
		$SiteLinkBridges = Get-ADReplicationSiteLinkBridge @slbParams | Sort-Object -Property Name
	}
	catch
	{
		$errorMessage = "{0}: {1}" -f $Error[0], $Error[0].InvocationInfo.PositionMessage
		Write-Error $errorMessage -ErrorAction Continue
	}

	$SiteLinkBridges.ForEach({
		$slbName = [String]$_.Name
		$slbDN = [String]$_.distinguishedName
		$slbLinksIncluded = [String]($_.SiteLinksIncluded -join "`n")
		$slbProtocol = [string]$_.IntersiteTransportProtocol

		$slbRow = $dtSLB.NewRow()
		$slbRow."Site Link Bridge Name" = $slbName
		$slbRow."Site Link Bridge DN" = $slbDN
		$slbRow."Site Link Transport Protocol" = $slbProtocol
		$slbRow."Site Links In Bridge" = $slbLinksIncluded
		
		
		$dtSLB.Rows.Add($slbRow)
		
		$slbName = $slbDN = $slbLinksIncluded = $slbProtocol = $null
		[System.GC]::GetTotalMemory('ForceFullCollection') | Out-Null
		
	})
	
	$null = $SiteLinkBridges
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
		Write-Verbose ("[{0} UTC] Exporting results data to CSV, please wait..." -f (Get-UTCTime).ToString($dtmFormatString))
		$outputFile = "{0}\{1}-{2}_Active_Directory_Site_Link_Bridge_Info.csv" -f $rptFolder, (Get-UTCTime).ToString($dtmFileFormatString), $DSForestName
		$dtSLB | Select-Object $colToExport | Export-Csv -Path $outputFile -NoTypeInformation
		
		Write-Verbose ("[{0} UTC] Exporting results data in Excel format, please wait..." -f (Get-UTCTime).ToString($dtmFormatString))
		$wsName = "AD Site-Link Bridge Config"
		
		$xlParams = @{
			Path	        = $outputFile.ToString().Replace([System.IO.Path]::GetExtension($outputFile), ".xlsx")
			WorkSheetName = $wsName
			TableStyle    = 'Medium15'
			StartRow	    = 2
			StartColumn   = 1
			AutoSize	    = $true
			AutoFilter    = $true
			BoldTopRow    = $true
			PassThru	    = $true
		}
	}
	
	$xl = $dtSLB | Select-Object $colToExport | Sort-Object -Property "Site Link Bridge Name" | Export-Excel @xlParams
	$Sheet = $xl.Workbook.Worksheets["AD Site-Link Bridge Config"]
	Set-ExcelRange -Range $Sheet.Cells["A2:Z2"] -WrapText -HorizontalAlignment Center -VerticalAlignment Center -AutoFit
	$cols = $Sheet.Dimension.Columns
	Set-ExcelRange -Range $Sheet.Cells["A3:Z$($cols)"] -Wraptext -HorizontalAlignment Left -VerticalAlignment Bottom
	Export-Excel -ExcelPackage $xl -WorksheetName $wsName -FreezePane 3, 0 -Title "$($DSForestName) Active Directory Site-Link Bridge Configuration" -TitleBold -TitleSize 16
}

#EndRegion