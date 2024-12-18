﻿#Requires -Module ActiveDirectory, ImportExcel, HelperFunctions
#Requires -Version 5
#Requires -RunAsAdministrator
<#

	.SYNOPSIS
	Export trust information for all trusts in an AD forest
	
	.DESCRIPTION
	This script gathers information on Active Directory trusts within the AD
	forest in parallel from which the script is run. 	The information is
	written to a datatableand then exported to a spreadsheet for artifact collection.
	
	.OUTPUTS
	CSV file containing domain trust information
	Excel spreasheet containing domain trust information

	.EXAMPLE 
	PS C:\>.\Export-ActiveDirectoryTrusts.ps1

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
param
(
[Parameter(Position = 0,
		 HelpMessage = 'Enter AD forest name to gather info. on.')]
[ValidateNotNullOrEmpty()]
[string[]]$DomainName,
[Parameter(Position = 1,
		 HelpMessage = 'Enter credential for remote domain.')]
[ValidateNotNull()]
[System.Management.Automation.PsCredential][System.Management.Automation.Credential()]
$Credential = [System.Management.Automation.PSCredential]::Empty
)

#Region Modules
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
		$modulePath = $modulePath | Sort-Object -Descending
		if ($modulePath.Count -gt 1)
		{
			$modulePath = $modulePath[0]
		}
		$psdPath = "{0}\{1}" -f $modulePath, "HelperFunctions.psd1"
		Import-Module $psdPath -ErrorAction Stop
	}
	catch
	{
		Write-Error "HelperFunctions PS module could not be loaded. $($_.Exception.Message)" -ErrorAction Stop
	}
}

#EndRegion


#Region Variables
$domainProperties = @("DistinguishedName", "DNSRoot", "Forest", "Name", "NetBIOSName", "ParentDomain", "PDCEmulator")
$ns = 'root\MicrosoftActiveDirectory'
$trustHeadersCsv =
@"
	ColumnName,DataType
	"Source Name",string
	"Target Name",string
	"Forest Transitive Trust",string
	"IntraForest Trust",string
	"Trust Direction",string
	"Trust Type",string
	"Trust Attributes",string
	"SID History",string
	"SID Filtering",string
	"Selective Authentication",string
	"UsesAESKeys",string
	"UsesRC4Encryption",string
	"CIMPartnerDCName",string
	"CIMTrustIsOK",string
	"CIMTrustStatus",string
	"AD Trust whenCreated",string
	"AD Trust whenChanged",string
"@
$dtmFormatString = "yyyy-MM-dd HH:mm:ss"
$dtmFileFormatString = "yyyy-MM-dd_HH-mm-ss"
$thisHost = $env:COMPUTERNAME
#EndRegion


#Region Functions

   
#EndRegion





#region Scripts
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
	
	#Create data table and add columns
	$trustTblName = "tblADDomainTrusts"
	$trustHeaders = ConvertFrom-Csv -InputObject $trustHeadersCsv
	$trustTable = Add-DataTable -TableName $trustTblName -ColumnArray $trustHeaders
	
	if (($PSBoundParameters.ContainsKey('DomainName')) -and ($null -ne $PSBoundParameters["DomainName"]))
	{
		$Domains = $DomainName -split (",")
	}
	else
	{
		try
		{
			$Domains = Get-ADDomain -Current LocalComputer -ErrorAction Stop
		}
		catch
		{
			$errorMessage = "{0}: {1}" -f $Error[0], $Error[0].InvocationInfo.PositionMessage
			Write-Error $errorMessage -ErrorAction Stop
		}
	}
	
	foreach ($Domain in $Domains)
	{
		
		# List of properties of a trust relationship
		$trusts = @()
		$trustStatus = @()
		
		$domainParams = @{
			Identity = $Domain
		}
		
		if (($PSBoundParameters.ContainsKey('Credential')) -and ($null -ne $PSBoundParameters["Credential"]))
		{
			$domainParams.Add('AuthType', 'Negotiate')
			$domainParams.Add('Credential', $Credential)
		}
		
		try
		{
			$domainInfo = Get-ADDomain @domainParams -Server (Get-ADDomain @domainParams).pdcEmulator -ErrorAction SilentlyContinue | Select-Object -Property $DomainProperties
			if ($? -eq $false)
			{
				$domainInfo = Get-ADDomain @domainParams -Server $Domain -ErrorAction Stop | Select-Object -Property $DomainProperties
			}
		}
		catch
		{
			$errorMessage = "{0}: {1}" -f $Error[0], $Error[0].InvocationInfo.PositionMessage
			Write-Error $errorMessage -ErrorAction Continue
			break
		}
		
		$pdcFSMO = ($domainInfo).PDCEmulator
		$domDNS = ($domainInfo).DNSRoot
		
		$trustParams = @{
			Filter     = '*'
			Properties = '*'
		}
		
		if (($PSBoundParameters.ContainsKey('Credential')) -and ($null -ne $PSBoundParameters["Credential"]))
		{
			$trustParams.Add('AuthType', 'Negotiate')
			$trustParams.Add('Credential', $Credential)
		}
		
		try
		{
			$trusts = @(Get-ADTrust @trustParams -Server $pdcFSMO -ErrorAction SilentlyContinue | Select-Object -Property *)
			if ($? -eq $false)
			{
				$trusts = @(Get-ADTrust @trustParams -Server $domDNS -ErrorAction Stop | Select-Object -Property *)
			}
		}
		catch
		{
			$errorMessage = "{0}: {1}" -f $Error[0], $Error[0].InvocationInfo.PositionMessage
			Write-Error $errorMessage -ErrorAction Continue
		}
		
		if ($trusts.Count -ge 1)
		{
			if (($localComputer.Name -eq $thisHost) -and ($localComputer.DomainRole -gt 3))
			{
				try
				{
					$trustStatus = Get-CimInstance -Namespace $ns -Query "Select * from Microsoft_DomainTrustStatus" -ErrorAction Stop

				}
				catch
				{
					$errorMessage = "{0}: {1}" -f $Error[0], $Error[0].InvocationInfo.PositionMessage
					Write-Error $errorMessage -ErrorAction Stop
				}
			}
			else
			{
				if (($PSBoundParameters.ContainsKey('Credential')) -and ($null -ne $PSBoundParameters["Credential"]))
				{
					try
					{
						$cimS = Get-MyNewCimSession -ServerName $pdcFSMO -Credential $Credential
					}
					catch
					{
						$errorMessage = "{0}: {1}" -f $Error[0], $Error[0].InvocationInfo.PositionMessage
						Write-Error $errorMessage -ErrorAction Stop
					}
				}
				else
				{
					try
					{
						$cimS = Get-MyNewCimSession -ServerName $pdcFSMO
					}
					catch
					{
						$errorMessage = "{0}: {1}" -f $Error[0], $Error[0].InvocationInfo.PositionMessage
						Write-Error $errorMessage -ErrorAction Stop
					}					
				}
				
				if ($null -ne $cimS.Name)
				{
					try
					{
						$trustStatus = Get-CimInstance -CimSession $cimS -Namespace $ns -Query "Select * from Microsoft_DomainTrustStatus" -ErrorAction Stop
					}
					catch
					{
						$errorMessage = "{0}: {1}" -f $Error[0], $Error[0].InvocationInfo.PositionMessage
						Write-Error $errorMessage -ErrorAction Continue
					}
				}
				
			}
			
			
			foreach ($t in $trusts)
			{
				$trustSource = Get-FQDNfromDN ($t).Source
				$trustTarget = ($t).Target
				$trustType = ($t).TrustType
				$forestTrust = ($t).ForestTransitive
				$intraForest = ($t).IntraForest
				$intTrustDirection = ($t).TrustDirection
				$usesAESKeys = ($t).UsesAESKeys
				$usesRC4Encryption = ($t).UsesRC4Encryption
				switch ($intTrustDirection)
				{
					0 { $trustDirection = "Disabled (The relationship exists but has been disabled)" }
					1 { $trustDirection = "Inbound (TrustING domain)" }
					2 { $trustDirection = "Outbound (TrustED domain)" }
					3 { $trustDirection = "Bidirectional (Two-Way Trust)" }
					Default{ $trustDirection = $intTrustDirection }
				}
				
				$TrustAttributesNumber = ($t).TrustAttributes
				switch ($TrustAttributesNumber)
				{
					
					1 { $trustAttributes = "Non-Transitive" }
					2 { $trustAttributes = "Uplevel clients only (Windows 2000 or newer" }
					4 { $trustAttributes = "Quarantined Domain (External)" }
					8 { $trustAttributes = "Forest Trust" }
					16 { $trustAttributes = "Cross-Organizational Trust (Selective Authentication)" }
					20 { $trustAttributes = "Intra-Forest Trust (trust within the forest)" }
					32 { $trustAttributes = "Intra-Forest Trust (trust within the forest)" }
					64 { $trustAttributes = "Inter-Forest Trust (trust with another forest)" }
					68 { $trustAttributes = "Quarantined Domain (External)" }
					Default { $trustAttributes = $TrustAttributesNumber }
					
				}
				
				if (-not ($trustAttributes)) { $trustAttributes = $TrustAttributesNumber }
				
				# Check if SID History is Enabled
				if ($TrustAttributesNumber -band 64) { $sidHistory = "Enabled" }
				else { $sidHistory = "Disabled" }
				
				# Check if SID Filtering is Enabled
				if ((($t.SIDFilteringQuarantined) -eq $false) -or (($t.SIDFilteringForestAware) -eq $false)) { $sidFiltering = "None" }
				else { $sidFiltering = "Quarantine Enabled" }
				
				if (($trustStatus).Count -ge 1)
				{
					$trustStatus | ForEach-Object {
						$trustPartnerDC = $_.TrustedDCName
						$partnerDC = $trustPartnerDC.TrimStart("\\")
						if ($_.TrustIsOk -eq $true) { $trustOK = "Yes" }
						else { $trustOK = "No - remediate" }
						$Status = ($_).TrustStatusString
					}
				}
				
				$trustSelectiveAuth = ($t).SelectiveAuthentication
				$whenCreated = ($t).Created -f "mm/dd/yyyy hh:mm:ss"
				$whenTrustChanged = ($t).modifyTimeStamp -f "mm/dd/yyyy hh:mm:ss"
				
				$trustRow = $trustTable.NewRow()
				$trustRow."Source Name" = $trustSource
				$trustRow."Target Name" = $trustTarget
				$trustRow."Forest Transitive Trust" = $forestTrust
				$trustRow."IntraForest Trust" = $intraForest
				$trustRow."Trust Direction" = $trustDirection
				$trustRow."Trust Type" = $trustType
				$trustRow."Trust Attributes" = $trustAttributes
				$trustRow."SID History" = $sidHistory
				$trustRow."SID Filtering" = $sidFiltering
				$trustRow."Selective Authentication" = $trustSelectiveAuth
				$trustRow."UsesAESKeys" = $usesAESKeys
				$trustRow."UsesRC4Encryption" = $usesRC4Encryption
				$trustRow."CIMPartnerDCName" = $partnerDC
				$trustRow."CIMTrustIsOK" = $trustOK
				$trustRow."CIMTrustStatus" = $Status
				$trustRow."AD Trust whenCreated" = $whenCreated
				$trustRow."AD Trust whenChanged" = $whenTrustChanged
				
				$trustTable.Rows.Add($trustRow)
				[GC]::Collect()
			}
		} #end $Trusts.Count
		
	}
}
catch
{
	$errorMessage = "{0}: {1}" -f $Error[0], $Error[0].InvocationInfo.PositionMessage
	Write-Error $errorMessage -ErrorAction Continue
}
finally
{
	#Save output
	#Check required folders and files exist, create if needed
	$rptFolder = 'E:\Reports'
	if ((Test-Path -Path $rptFolder -PathType Container) -eq $false) { New-Item -Path $rptFolder -ItemType Directory -Force }
	Test-PathExists -Path $rptFolder -PathType Folder
	
	if ($trustTable)
	{
		$ttColToExport = $trustHeaders.ColumnName
		
		Write-Verbose ("[{0} UTC] Exporting results data to CSV, please wait..." -f (Get-UTCTime).ToString($dtmFormatString))
		$outputFile = "{0}\{1}-{2}_Active_Directory_Domain_Trust_Info.csv" -f $rptFolder, (Get-UTCTime).ToString($dtmFileFormatString), $domDNS.ToString().ToUpper()
		$xlOutput = $OutputFile.ToString().Replace([System.IO.Path]::GetExtension($OutputFile), ".xlsx")
		$trustTable | Select-Object $ttColToExport | Export-Csv -Path $outputFile -NoTypeInformation
		
		Write-Verbose ("[{0} UTC] Exporting results data in Excel format, please wait..." -f $(Get-UTCTime).ToString($dtmFormatString))
		$wsName = "AD Trust Configuration"
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
		}
		
		$titleParams = @{
			FontColor         = 'White'
			FontSize	        = 16
			Bold		        = $true
			BackgroundColor   = 'Black'
			BackgroundPattern = 'Solid'
		}
		
		$xl = $trustTable | Select-Object $ttColToExport | Export-Excel @xlParams
		$Sheet = $xl.Workbook.Worksheets[$wsName]
		$lastRow = $siteSheet.Dimension.End.Row
		
		Set-ExcelRange -Range $Sheet.Cells["A1"] -Value "Active Directory Domain Trust Configuration" @titleParams
		Set-ExcelRange -Range $Sheet.Cells["A2"] @headerParams1
		Set-ExcelRange -Range $Sheet.Cells["B2:Z2"] @headerParams2
		Set-ExcelRange -Range $Sheet.Cells["A3:Z$($lastRow)"] @setParams
		
		Export-Excel -ExcelPackage $xl -AutoSize -FreezePane 3, 0 -WorksheetName $wsName
	} #end If
	
	if ($null -ne $cimS.Name)
	{
		Remove-CimSession -Id $cimS.Id
	}
}
#EndRegion