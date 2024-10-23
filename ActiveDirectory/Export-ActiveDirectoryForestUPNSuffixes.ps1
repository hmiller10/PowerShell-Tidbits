#Requires -Module  ActiveDirectory, ImportExcel, HelperFunctions
#Requires -Version 5
#Requires -RunAsAdministrator
<#
	.SYNOPSIS
		Export AD Forest Info to Excel or CSV
	
	.DESCRIPTION
		This script is designed to gather and report information on an Active Directory forest.
	
	.PARAMETER ForestName
		Active Directory Forest Name
	
	.PARAMETER Credential
		PS credential object
	
	.PARAMETER OutputFormat
		Specify the format the data should be exported to. Choices are CSV or Excel
	
	.EXAMPLE
		.\Export-ADForestInfo.ps1 -ForestName example.com -Credential (Get-Credential)
	
	.EXAMPLE
		.\Export-ADForestInfo.ps1 -CSV
	
	.EXAMPLE
		.\Export-ADForestInfo.ps1 -Excel
	
	.EXAMPLE
		.\Export-ADForestInfo.ps1 -ForestName myForest.com -Credential (Get-Credential) -Excel
	
	.OUTPUTS
		Excel file containing relevant forest information
	
	.NOTES
		THIS CODE IS MADE AVAILABLE AS IS, WITHOUT WARRANTY OF ANY KIND. THE
		ENTIRE RISK OF THE USE OR THE RESULTS FROM THE USE OF THIS CODE REMAINS
		WITH THE USER.
#>
[CmdletBinding(PositionalBinding = $false)]
param
(
	[Parameter(Position = 0,
	           HelpMessage = 'Enter AD forest name to gather info. on.')]
	[ValidateNotNullOrEmpty()]
	[string]
	$ForestName,
	[Parameter(Position = 1,
	           HelpMessage = 'Enter PS credential to connecct to AD forest with.')]
	[ValidateNotNullOrEmpty()]
	[pscredential]
	$Credential,
	[Parameter(Mandatory = $true,
	           Position = 2)]
	[ValidateNotNullOrEmpty()]
	[ValidateSet('CSV', 'Excel', IgnoreCase = $true)]
	[string]
	$OutputFormat
)

#region Execution Policy
#Set-ExecutionPolicy -ExecutionPolicy RemoteSigned
#endregion

#region Modules
#Check if required module is loaded, if not load import it
Try
{
	Import-Module -Name ActiveDirectory -ErrorAction Stop
}
Catch
{
	Throw "Active Directory module could not be loaded. $($_.Exception.Message)";
}

try
{
	Import-Module -Name ImportExcel -ErrorAction Stop
}
catch
{
	throw "PowerShell ImportExcel module could not be loaded. $($_.Exception.Message)";

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

#endregion

#Region Global Variables

$upnHeadersCsv =
@"
ColumnName,DataType
"Forest Name",string
"UPN Suffix",string
"@
$dtmFormatString = "yyyy-MM-dd HH:mm:ss"
$dtmFileFormatString = "yyyy-MM-dd_HH-mm-ss"
$forestProperties = @("Name", "UPNSuffixes")
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
	
	
	#Region Forest Config
	foreach ($Forest in $ForestName)
	{
		$ForestParams = @{
			ErrorAction = 'Stop'
		}
		
		if (($PSBoundParameters.ContainsKey('ForestName')) -and ($null -ne $PSBoundParameters["ForestName"]))
		{
			$ForestParams.Add('Identity', $Forest)
			$ForestParams.Add('Server', $Forest)
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
			$DSForest = Get-ADForest @ForestParams | Select-Object -Property $forestProperties
			$DSForestName = ($DSForest).Name.ToString().ToUpper()
			$upnSuffixes = $DSForest.UPNSuffixes
			
		}
		catch
		{
			$errorMessage = "{0}: {1}" -f $Error[0], $Error[0].InvocationInfo.PositionMessage
			Write-Error $errorMessage -ErrorAction Continue
		}
		
		if ($DSForest.UPNSuffixes.Count -ge 1)
		{
			#Create data table and add columns
			$upnTblName = "$($DSForestName)_Information"
			$upnHeaders = ConvertFrom-Csv -InputObject $upnHeadersCsv
			
			try
			{
				$upnTable = Add-DataTable -TableName $upnTblName -ColumnArray $upnHeaders -ErrorAction Stop
			}
			catch
			{
				$errorMessage = "{0}: {1}" -f $Error[0], $Error[0].InvocationInfo.PositionMessage
				Write-Error $errorMessage -ErrorAction Continue
			}
			
			$UpnCount = 1
			foreach ($upn in $upnSuffixes)
			{
				Write-Verbose -Message ("Processing AD Forest {0} UPN Suffix {1}" -f $DSForestName, $upn)
				
				$UpnActivityMessage = "Gathering AD UPN suffix information, please wait..."
				$UpnProcessingStatus = "Processing UPN {0} of {1}: {2}" -f $UpnCount, $upnSuffixes.count, $upn
				$percentUpnComplete = ($UpnCount / $upnSuffixes.count * 100)
				Write-Progress -Activity $UpnActivityMessage -Status $UpnProcessingStatus -PercentComplete $percentUpnComplete -Id 1
				
				$upnRow = $upnTable.NewRow()
				$upnRow."Forest Name" = $DSForestName
				$upnRow."UPN Suffix" = [String]$upn
				
				$upnTable.Rows.Add($upnRow)
				$null = $upn
				
				$UpnCount++
			}
			
			Write-Progress -Activity "Done gathering AD UPN Suffix information for $($DSForestName)" -Status "Ready" -Completed
			$null = $upnSuffixes
			[System.GC]::GetTotalMemory('forcefullcollection') | Out-Null
		}
		else
		{
			throw "There are no UPN suffixes assigned to this AD forest."
		}

	}#end $Forest
	
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
	
	$colToExport = $upnHeaders.ColumnName
	if ($upnTable.Rows.Count -ge 1)
	{
		switch ($PSBoundParameters["OutputFormat"])
		{
			"CSV" {
				Write-Verbose ("[{0} UTC] Exporting results data to CSV, please wait..." -f (Get-UTCTime).ToString($dtmFormatString))
				$outputFile = "{0}\{1}_{2}_Active_Directory_UPNSuffix_List.csv" -f $rptFolder, (Get-UTCTime).ToString($dtmFileFormatString), $DSForestName
				$upnTable | Select-Object $colToExport | Export-Csv -Path $outputFile -NoTypeInformation
			}
			"Excel" {
				Write-Verbose -Message ("[{0} UTC] Exporting data tables to Excel spreadsheet tabs." -f $(Get-UTCTime).ToString($dtmFormatString))
				$outputFile = "{0}\{1}_{2}_Active_Directory_UPNSuffix_List.xlsx" -f $rptFolder, (Get-UTCTime).ToString($dtmFileFormatString), $DSForestName
				[string]$wsName = "AD Forest UPN List"
				$xlParams = @{
					Path	        = $outputFile
					WorkSheetName = $wsName
					TableStyle = 'Medium15'
					StartRow     = 2
					StartColumn  = 1
					AutoSize   = $true
					AutoFilter   = $true
					BoldTopRow   = $true
					FreezeTopRow = $true
					PassThru = $true
				}

				$xl = $upnTable | Select-Object $colToExport | Sort-Object -Property "Forest Name" | Export-Excel @xlParams
				$Sheet = $xl.Workbook.Worksheets["AD Forest UPN List"]
				Set-ExcelRange -Range $Sheet.Cells["A2:Z2"] -WrapText -HorizontalAlignment Center -VerticalAlignment Center -AutoFit
				$cols = $Sheet.Dimension.Columns
				Set-ExcelRange -Range $Sheet.Cells["A3:Z$($cols)"] -Wraptext -HorizontalAlignment Left -VerticalAlignment Bottom
				Export-Excel -ExcelPackage $xl -WorksheetName $wsName -Title "$($DSForestName) Active Directory UPN Suffix List" -TitleBold -TitleSize 16
			}
		} #end Switch
		
	} #end if
	
}

#endregion