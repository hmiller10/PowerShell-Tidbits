#Requires -Module ActiveDirectory, GroupPolicy, HelperFunctions, ImportExcel
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
[CmdletBinding(PositionalBinding = $false)]
param
(
	[Parameter(Position = 0,
			 HelpMessage = 'Enter AD forest name to gather info. on.')]
	[ValidateNotNullOrEmpty()]
	[string]$ForestName,
	[Parameter(Position = 1,
			 HelpMessage = 'Enter credential for remote forest.')]
	[ValidateNotNull()]
	[System.Management.Automation.PsCredential][System.Management.Automation.Credential()]
	$Credential = [System.Management.Automation.PSCredential]::Empty,
	[Parameter(Mandatory = $true,
			 Position = 2)]
	[ValidateNotNullOrEmpty()]
	[ValidateSet('CSV', 'Excel', IgnoreCase = $true)]
	[string]$OutputFormat
)

#region Execution Policy
#Set-ExecutionPolicy -ExecutionPolicy RemoteSigned
#endregion

#region Modules
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
	Import-Module -Name ImportExcel -ErrorAction Stop
}
catch
{
	throw "PowerShell ImportExcel module could not be loaded. $($_.Exception.Message)";
	
}

try
{
	Import-Module -Name HelperFunctions -Force -ErrorAction Stop
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

$forestHeadersCsv =
@"
ColumnName,DataType
"Forest Name",string
"Forest Functional Level",string
"Forest Root Domain",string
"Domains in Forest",string
"Forest Partitions Container",string
"Forest Application Partitions",string
"Replicated Naming Contexts",string
"Schema Master FSMO Holder",string
"Domain Naming Master FSMO Holder",string
"Recycle Bin Enabled",string
"Recycle Bin Scope",string
"Recycle Bin Deleted Object Lifetime in Days",string
"Recycle Bin Tombstone Object Lifetime in Days",string
"@

#EndRegion

#Region Functions

function Set-FreezePane
{
    <#
    .SYNOPSIS
        Set FreezePanes on a specified worksheet
 
    .DESCRIPTION
        Set FreezePanes on a specified worksheet
     
    .PARAMETER Worksheet
        Worksheet to add FreezePanes to
     
    .PARAMETER Row
        The first row with live data.
 
        Examples and outcomes:
            -Row 2 Freeze row 1
            -Row 5 Freeze rows 1 through 4
 
    .PARAMETER Column
        Examples and outcomes:
            -Column 2 Freeze column 1
            -Column 5 Freeze columns 1 through 4
 
    .PARAMETER Passthru
        If specified, pass the Worksheet back
 
    .EXAMPLE
        $WorkSheet | Set-FreezePane
 
        #Freeze the top row of $Worksheet (default parameter values handle this)
 
    .EXAMPLE
        $WorkSheet | Set-FreezePane -Row 2 -Column 4
 
        # Freeze the top row and top 3 columns of $Worksheet
 
    .NOTES
        Thanks to Doug Finke for his example:
            https://github.com/dfinke/ImportExcel/blob/master/ImportExcel.psm1
 
        Thanks to Philip Thompson for an expansive set of examples on working with EPPlus in PowerShell:
            https://excelpslib.codeplex.com/
 
    .LINK
        https://github.com/RamblingCookieMonster/PSExcel
 
    .FUNCTIONALITY
        Excel
    #>
	[cmdletbinding()]
	param (
		[parameter(Mandatory = $true,
				 ValueFromPipeline = $true,
				 ValueFromPipelineByPropertyName = $true)]
		[OfficeOpenXml.ExcelWorksheet]$WorkSheet,
		[int]$Row = 2,
		[int]$Column = 1,
		[switch]$Passthru
		
	)
	process
	{
		$WorkSheet.View.FreezePanes($Row, $Column)
		if ($Passthru)
		{
			$WorkSheet
		}
	}
}

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
		if (($localComputer.Caption -match "Windows 11") -eq $true)
		{
			try
			{
				#https://docs.microsoft.com/en-us/dotnet/api/system.net.securityprotocoltype?view=netcore-2.0#System_Net_SecurityProtocolType_SystemDefault
				if ($PSVersionTable.PSVersion.Major -lt 6 -and [Net.ServicePointManager]::SecurityProtocol -notmatch 'Tls13')
				{
					Write-Verbose -Message 'Adding support for TLS 1.3'
					[Net.ServicePointManager]::SecurityProtocol += [Net.SecurityProtocolType]::Tls13
				}
			}
			catch
			{
				Write-Warning -Message 'Adding TLS 1.3 to supported security protocols was unsuccessful.'
			}
		}
		elseif (($localComputer.Caption -match "Server 2022") -eq $true)
		{
			try
			{
				#https://docs.microsoft.com/en-us/dotnet/api/system.net.securityprotocoltype?view=netcore-2.0#System_Net_SecurityProtocolType_SystemDefault
				if ($PSVersionTable.PSVersion.Major -lt 6 -and [Net.ServicePointManager]::SecurityProtocol -notmatch 'Tls13')
				{
					Write-Verbose -Message 'Adding support for TLS 1.3'
					[Net.ServicePointManager]::SecurityProtocol += [Net.SecurityProtocolType]::Tls13
				}
			}
			catch
			{
				Write-Warning -Message 'Adding TLS 1.3 to supported security protocols was unsuccessful.'
			}
		}
	}
	
	$dtmFormatString = "yyyy-MM-dd HH:mm:ss"
	$dtmFileFormatString = "yyyy-MM-dd_HH-mm-ss"
	
	#Region Forest Config
	#Get AD Forest Basic Information
	
	Write-Verbose -Message ("Working on {0}." -f $Forest)
	$forestProperties = @("ApplicationPartitions", "Domains", "DomainNamingMaster", "ForestMode", "Name", "RootDomain", "PartitionsContainer", "SchemaMaster", "SPNSuffixes")
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
		$ForestParams = @{
			Current = 'LocalComputer'
		}
	}
	
	if (($PSBoundParameters.ContainsKey('Credential')) -and ($null -ne $PSBoundParameters["Credential"]))
	{
		$ForestParams.Add('AuthType', 'Negotiate')
		$ForestParams.Add('Credential', $Credential)
	}
	
	try
	{
		$DSForest = Get-ADForest @ForestParams | Select-Object -Property $forestProperties
		$DSForestName = $DSForest.Name.ToString().ToUpper()
		$schemaFSMO = ($DSForest).SchemaMaster
	}
	catch
	{
		$errorMessage = "{0}: {1}" -f $Error[0], $Error[0].InvocationInfo.PositionMessage
		Write-Error $errorMessage -ErrorAction Continue
	}
	
	#Get RootDSE
	$dseParams = @{
		Server	  = $schemaFSMO
		ErrorAction = 'Stop'
	}
	
	if (($PSBoundParameters.ContainsKey('Credential')) -and ($null -ne $PSBoundParameters["Credential"]))
	{
		$dseParams.Add('AuthType', 'Negotiate')
		$dseParams.Add('Credential', $Credential)
	}
	
	try
	{
		$rootDSE = Get-ADRootDse @dseParams
		$rootCNC = $rootDSE.ConfigurationNamingContext
	}
	catch
	{
		$errorMessage = "{0}: {1}" -f $Error[0], $Error[0].InvocationInfo.PositionMessage
		Write-Error $errorMessage -ErrorAction Continue
	}
	
	$DSForestName = ($DSForest).Name.ToUpper()
	$forestFunctionalLevel = ($DSForest).ForestMode.ToString().ToUpper()
	$forestRootDomain = ($DSForest).RootDomain.ToString().ToUpper()
	$Domains = ($DSForest).Domains
	$forestReplCntxt = ($DSForest).PartitionsContainer
	$appPartitions = ($DSForest).ApplicationPartitions | Select-Object -Unique
	$dnmFSMO = ($DSForest).DomainNamingMaster
	
	#Get replicated connections
	if (($PSBoundParameters.ContainsKey('Credential')) -and ($null -ne $PSBoundParameters["Credential"]))
	{
		$Partitions = (Get-ADReplicationConnection -AuthType 0 -Credential $Credential -Properties ReplicatedNamingContexts -Server $schemaFSMO).ReplicatedNamingContexts | Select-Object -Unique
	}
	else
	{
		$Partitions = (Get-ADReplicationConnection -Properties ReplicatedNamingContexts -Server $schemaFSMO).ReplicatedNamingContexts | Select-Object -Unique
	}
	
	#Detect AD recycle bin status
	if (($PSBoundParameters.ContainsKey('Credential')) -and ($null -ne $PSBoundParameters["Credential"]))
	{
		$objRecBin = Get-ADOptionalFeature -Filter 'Name -eq "Recycle Bin Feature"' -AuthType 0 -Credential $Credential -Properties Name, FeatureScope -Server $schemaFSMO | Select-Object -Property Name, FeatureScope
	}
	else
	{
		$objRecBin = Get-ADOptionalFeature -Filter 'Name -eq "Recycle Bin Feature"' -Properties Name, FeatureScope -Server $schemaFSMO | Select-Object -Property Name, FeatureScope
	}
	
	if ($null -ne $objRecBin.Name) { $recBinEnabled = $true }
	else { $recBinEnabled = $false }
	
	$recBinDN = "CN=Directory Service,CN=Windows NT,CN=Services,{0}" -f $rootCNC
	
	#Get AD deleted object lifetime
	if (($PSBoundParameters.ContainsKey('Credential')) -and ($null -ne $PSBoundParameters["Credential"]))
	{
		$objDelLifeTime = (Get-ADObject -Identity $recBinDN -AuthType 0 -Credential $Credential -Properties msDS-DeletedObjectLifeTime -Partition $rootCNC -Server $schemaFSMO).'msDS-DeletedObjectLifeTime'
	}
	else
	{
		$objDelLifeTime = (Get-ADObject -Identity $recBinDN -Properties msDS-DeletedObjectLifeTime -Partition $rootCNC -Server $schemaFSMO).'msDS-DeletedObjectLifeTime'
	}
	
	if ($null -eq $objDelLifeTime)
	{
		$objDelLifeTime = "Default"
	}
	
	#Get Config. partition
	$dsConfigDN = "CN=Directory Service,CN=Windows NT,CN=Services," + $rootCNC
	
	try
	{
		if (($PSBoundParameters.ContainsKey('Credential')) -and ($null -ne $PSBoundParameters["Credential"]))
		{
			$configPartition = Get-ADObject -Identity $dsConfigDN -AuthType 0 -Credential $Credential -Properties * -Server $schemaFSMO
		}
		else
		{
			$configPartition = Get-ADObject -Identity $dsConfigDN -Properties * -Server $schemaFSMO
		}
	}
	catch
	{
		$errorMessage = "{0}: {1}" -f $Error[0], $Error[0].InvocationInfo.PositionMessage
		Write-Error $errorMessage -ErrorAction Continue
	}
	
	#Get Tombstone Object Lifetime
	try
	{
		if (($PSBoundParameters.ContainsKey('Credential')) -and ($null -ne $PSBoundParameters["Credential"]))
		{
			[string]$objTSLifeTime = (Get-ADObject -Identity $configPartition.distinguishedName -AuthType 0 -Credential $Credential -Properties tombstoneLifeTime -Partition $rootCNC -Server $schemaFSMO).tombstoneLifeTime
		}
		else
		{
			[string]$objTSLifeTime = (Get-ADObject -Identity $configPartition.distinguishedName -Properties tombstoneLifeTime -Partition $rootCNC -Server $schemaFSMO).tombstoneLifeTime
		}
		
	}
	catch
	{
		$errorMessage = "{0}: {1}" -f $Error[0], $Error[0].InvocationInfo.PositionMessage
		Write-Error $errorMessage -ErrorAction Continue
	}
	
	#Get AD replication connections
	try
	{
		if (($PSBoundParameters.ContainsKey('Credential')) -and ($null -ne $PSBoundParameters["Credential"]))
		{
			$Partitions = (Get-ADReplicationConnection -AuthType 0 -Credential $Credential -Properties ReplicatedNamingContexts -Server $schemaFSMO -ErrorAction Stop).ReplicatedNamingContexts | Select-Object -Unique
		}
		else
		{
			$Partitions = (Get-ADReplicationConnection -Properties ReplicatedNamingContexts -Server $schemaFSMO -ErrorAction Stop).ReplicatedNamingContexts | Select-Object -Unique
		}
		
	}
	catch
	{
		$errorMessage = "{0}: {1}" -f $Error[0], $Error[0].InvocationInfo.PositionMessage
		Write-Error $errorMessage -ErrorAction Continue
	}
	
	if ($Partitions.Count -ge 1)
	{
		if (($PSBoundParameters.ContainsKey('OutputFormat')) -and ($PSBoundParameters["OutputFormat"] -eq 'CSV'))
		{
			$replParts = $Partitions -join " "
		}
		
		if (($PSBoundParameters.ContainsKey('OutputFormat')) -and ($PSBoundParameters["OutputFormat"] -eq 'Excel'))
		{
			$replParts = $Partitions -join "`n"
		}
	}
	else
	{
		$replParts = "Nothing replicated."
	}
	
	if ($DSForest.ApplicationPartitions.Count -ge 1)
	{
		$appPartitions = ($DSForest).ApplicationPartitions | Select-Object -Unique
		if (($PSBoundParameters.ContainsKey('OutputFormat')) -and ($PSBoundParameters["OutputFormat"] -eq 'CSV'))
		{
			$appPartitions = $appPartitions -join " "
		}
		
		if (($PSBoundParameters.ContainsKey('OutputFormat')) -and ($PSBoundParameters["OutputFormat"] -eq 'Excel'))
		{
			$appPartitions = $appPartitions -join "`n"
		}
	}
	else
	{
		$appPartitions = 'None'
	}
	
	if (($PSBoundParameters.ContainsKey('OutputFormat')) -and ($PSBoundParameters["OutputFormat"] -eq 'CSV'))
	{
		$domList = $Domains -join " "
	}
	
	if (($PSBoundParameters.ContainsKey('OutputFormat')) -and ($PSBoundParameters["OutputFormat"] -eq 'Excel'))
	{
		$domList = $Domains -join "`n"
	}
	
	#Create data table and add columns
	$forestTblName = "$($DSForestName)_Information"
	$forestHeaders = ConvertFrom-Csv -InputObject $forestHeadersCsv
	try
	{
		$forestTable = Add-DataTable -TableName $forestTblName -ColumnArray $forestHeaders -ErrorAction Stop
	}
	catch
	{
		$errorMessage = "{0}: {1}" -f $Error[0], $Error[0].InvocationInfo.PositionMessage
		Write-Error $errorMessage -ErrorAction Continue
	}
	
	$forestRow = $forestTable.NewRow()
	$forestRow."Forest Name" = $DSForestName
	$forestRow."Forest Functional Level" = $forestFunctionalLevel
	$forestRow."Forest Root Domain" = $forestRootDomain
	$forestRow."Domains in Forest" = [String]$domList
	$forestRow."Forest Partitions Container" = [String]$forestReplCntxt
	$forestRow."Forest Application Partitions" = $appPartitions | Out-String
	$forestRow."Replicated Naming Contexts" = [String]$replParts
	$forestRow."Schema Master FSMO Holder" = [String]$schemaFSMO
	$forestRow."Domain Naming Master FSMO Holder" = [String]$dnmFSMO
	#To-Do Add rows dependent upon Recycle Bin state
	$forestRow."Recycle Bin Enabled" = $recBinEnabled
	$forestRow."Recycle Bin Scope" = [String]$objRecBin.FeatureScope
	$forestRow."Recycle Bin Deleted Object Lifetime in Days" = $objDelLifeTime
	$forestRow."Recycle Bin Tombstone Object Lifetime in Days" = $objTSLifeTime
	
	$forestTable.Rows.Add($forestRow)
	
	$null = $forestFunctionalLevel = $forestRootDomain = $domList = $forestReplCntxt = $appPartitions = $Partitions = $replParts
	$null = $schemaFSMO = $dnmFSMO = $objRecBin = $recBinEnabled
	$null = $objDelLifeTime = $objTSLifeTime
	[System.GC]::GetTotalMemory('forcefullcollection') | Out-Null
	
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
	
	$colToExport = $forestHeaders.ColumnName
	$outputFile = "{0}\{1}_{2}_Active_Directory_Forest_Info.csv" -f $rptFolder, (Get-UTCTime).ToString($dtmFileFormatString), $DSForestName
	if ($forestTable.Rows.Count -ge 1)
	{
		switch ($PSBoundParameters["OutputFormat"])
		{
			"CSV" {
				Write-Verbose ("[{0} UTC] Exporting results data to CSV, please wait..." -f (Get-UTCTime).ToString($dtmFormatString))
				$forestTable | Select-Object $colToExport | Export-Csv -Path $outputFile -NoTypeInformation
			}
			"Excel" {
				Write-Verbose -Message ("[{0} UTC] Exporting data tables to Excel spreadsheet tabs." -f $(Get-UTCTime).ToString($dtmFormatString))
				$xlOutput = $outputFile.ToString().Replace([System.IO.Path]::GetExtension($OutputFile), ".xlsx")
				[string]$wsName = "AD Forest Configuration"
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
					VerticalAlignment   = 'Bottom'
					HorizontalAlignment = 'Left'
					ErrorAction         = 'SilentlyContinue'
				}
				
				$titleParams = @{
					AutoSize	        = $true
					FontColor	        = 'White'
					FontSize	        = 16
					Bold		        = $true
					BackgroundColor   = 'Black'
					BackgroundPattern = 'Solid'
				}
				
				$xl = $forestTable | Select-Object $colToExport | Sort-Object -Property "Forest Name" | Export-Excel @xlParams
				$Sheet = $xl.Workbook.Worksheets[$wsName]
				$lastRow = $Sheet.Dimension.End.Row
				
				Set-ExcelRange -Range $Sheet.Cells["A1"] -Value "$($DSForestName) Active Directory Forest Configuration" @titleParams
				Set-ExcelRange -Range $Sheet.Cells["A2"] @headerParams1
				Set-ExcelRange -Range $Sheet.Cells["B2:Z2"] @headerParams2
				Set-ExcelRange -Range $Sheet.Cells["A3:M$($lastRow)"] @setParams
				
				Export-Excel -ExcelPackage $xl -WorksheetName $wsName -FreezePane 3, 0
				
			}
		} #end Switch
		
	} #end if
	
}

#endregion