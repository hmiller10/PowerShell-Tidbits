#Requires -Module  ActiveDirectory, ImportExcel, HelperFunctions
#Requires -Version 5
#Requires -RunAsAdministrator
<#

.NOTES
#------------------------------------------------------------------------------
#
# THIS CODE IS MADE AVAILABLE AS IS, WITHOUT WARRANTY OF ANY KIND. THE
# ENTIRE RISK OF THE USE OR THE RESULTS FROM THE USE OF THIS CODE REMAINS
# WITH THE USER.
#
#------------------------------------------------------------------------------
.SYNOPSIS
	Export trust information for all trusts in an AD forest
	
.DESCRIPTION
	This script gathers information on Active Directory trusts within the AD
	forest in parallel from which the script is run. 	The information is
	written to a datatableand then exported to a spreadsheet for artifact collection.
	
.OUTPUTS
	Excel spreasheet containing forest/domain trust information
	
.EXAMPLE 
	PS C:\>.\Export-ActiveDirectoryTrusts.ps1
	
.EXAMPLE 
	PS C:\>.\Export-ActiveDirectoryTrusts.ps1 -ForestName myForest.com -Credential (Get-Credential)
#>

###########################################################################
#
#
# AUTHOR:  Heather Miller
#
# VERSION HISTORY:
# 3.0 11/20/2023 - Added error handling and credential support
#
# 
###########################################################################

param
(
	[Parameter(Position = 0,
			 HelpMessage = 'Enter AD forest name to gather info. on.')]
	[ValidateNotNullOrEmpty()]
	[string[]]$ForestName,
	[Parameter(Position = 1,
			 HelpMessage = 'Enter PS credential to connecct to AD forest with.')]
	[ValidateNotNullOrEmpty()]
	[pscredential]$Credential
)

#Region Modules
try 
{
	Import-Module -Name ActiveDirectory -Force -ErrorAction Stop
}
catch
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


#Region Variables
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
#EndRegion


#Region Functions


#EndRegion



#region Scripts
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
	
	#Create data table and add columns
	$trustTblName = "Forest_Domains_Trust_Info"
	$trustHeaders = ConvertFrom-Csv -InputObject $trustHeadersCsv
	try
	{
		$trustTable = Add-DataTable -TableName $trustTblName -ColumnArray $trustHeaders -ErrorAction Stop
	}
	catch
	{
		$errorMessage = "{0}: {1}" -f $Error[0], $Error[0].InvocationInfo.PositionMessage
		Write-Error $errorMessage -ErrorAction Stop
	}
	
	if (($PSBoundParameters.ContainsKey('ForestName')) -and ($null -ne $PSBoundParameters["ForestName"]))
	{
		Write-Output ("Forests to query are: {0}" -f $ForestName)
		$ForestName = $ForestName -split (",")
	}
	else
	{
		$ForestName = [System.DirectoryServices.ActiveDirectory.Forest]::GetCurrentForest().Name.ToString()
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
			$DSForestName = $DSForest.Name.ToString().ToUpper()
		}
		catch
		{
			$errorMessage = "{0}: {1}" -f $Error[0], $Error[0].InvocationInfo.PositionMessage
			Write-Error $errorMessage -ErrorAction Continue
		}
		
		foreach ($Domain in $DSForest.Domains)
		{
			$DomainParams = @{
				Identity = $Domain
				Server   = $Domain
			}
			
			if (($PSBoundParameters.ContainsKey('Credential')) -and ($null -ne $PSBoundParameters["Credential"]))
			{
				$DomainParams.Add('AuthType', 'Negotiate')
				$DomainParams.Add('Credential', $Credential)
			}
			
			# List of properties of a trust relationship
			$trusts = @()
			$trustStatus = @()
			
			try
			{
				$domainInfo = Get-ADDomain @DomainParams -ErrorAction Stop
			}
			catch
			{
				$errorMessage = "{0}: {1}" -f $Error[0], $Error[0].InvocationInfo.PositionMessage
				Write-Error $errorMessage -ErrorAction Continue
				break;
			}
			
			if ($null -ne $domainInfo)
			{
				$pdcFSMO = ($domainInfo).PDCEmulator
				$domDNS = ($domainInfo).DNSRoot
			}
			else
			{
				$pdcFSMO = $Domain
				$domDNS = $Domain
			}
			
			$trustParams = @{
				Filter     = '*'
				Properties = '*'
			}
			
			if (($PSBoundParameters.ContainsKey('Credential')) -and ($null -ne $PSBoundParameters["Credential"]))
			{
				$trustParams.Add('AuthType', 'Negotitate')
				$trustParams.Add('Credential', $Credential)
			}
			
			try
			{
				$trusts = Get-ADTrust @trustParams -Server $pdcFSMO -ErrorAction SilentlyContinue
				if ($? -eq $false)
				{
					$trusts = Get-ADTrust @trustParams -Server $domDNS -ErrorAction Stop
				}
			}
			catch
			{
				$errorMessage = "{0}: {1}" -f $Error[0], $Error[0].InvocationInfo.PositionMessage
				Write-Error $errorMessage -ErrorAction Continue
				break;
			}
			
			if (($trusts.Count -ge 1) -and ($null -ne $trusts))
			{
				if ($domainInfo.Name -ne ([System.DirectoryServices.ActiveDirectory.Domain]::GetCurrentDomain().Name.ToString()))
				{
					try
					{
						$s = New-CimSession -Authentication Negotiate -ComputerName $pdcFsmo -Credential $Credential -ErrorAction SilentlyContinue
						if ($? -eq $false)
						{
							$s = New-CimSession -Authentication Negotiate -ComputerName $domDNS -Credential $Credential -ErrorAction Stop
						}
					}
					catch
					{
						$errorMessage = "{0}: {1}" -f $Error[0], $Error[0].InvocationInfo.PositionMessage
						Write-Error $errorMessage -ErrorAction Continue
					}
					
					try
					{
						if ($null -ne $s.Name)
						{
							$trustStatus = Get-CimInstance -CimSession $s -Namespace 'root\MicrosoftActiveDirectory' -Query "Select * from Microsoft_DomainTrustStatus" -ErrorAction SilentlyContinue -ErrorVariable CIMError
						}
						
					}
					catch
					{
						$errorMessage = "{0}: {1}" -f $CIMError[0], $CIMError[0].InvocationInfo.PositionMessage
						Write-Error $errorMessage -ErrorAction Continue
					}
				}
				else
				{
					try
					{
						
						$trustStatus = Get-CimInstance -ComputerName $pdcFSMO -Namespace 'root\MicrosoftActiveDirectory' -Query "Select * from Microsoft_DomainTrustStatus" -ErrorAction SilentlyContinue -ErrorVariable CIMError
						if ($? -eq $false)
						{
							$trustStatus = Get-CimInstance -ComputerName $domDNS -Namespace 'root\MicrosoftActiveDirectory' -Query "Select * from Microsoft_DomainTrustStatus" -ErrorAction SilentlyContinue -ErrorVariable CIMError
						}
						
					}
					catch
					{
						$errorMessage = "{0}: {1}" -f $CIMError[0], $CIMError[0].InvocationInfo.PositionMessage
						Write-Error $errorMessage -ErrorAction Continue
					}
				}
				
				foreach ($t in $trusts)
				{
					$trustSource = Get-FQDNfromDN -DistinguishedName ($t).Source
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
					$whenCreated = ($t).Created -f $dtmFileFormatString
					$whenTrustChanged = ($t).modifyTimeStamp -f $dtmFileFormatString
					
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
		}#end foreach $Domain
	}#end foreach $Forest
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

	if ($trustTable.Rows.Count -gt 1)
	{
		$ttColToExport = $trustHeaders.ColumnName
		
		Write-Verbose ("[{0} UTC] Exporting results data to CSV, please wait..." -f (Get-UTCTime).ToString($dtmFormatString))
		$outputFile = "{0}\{1}-{2}_Active_Directory_Forest_Trust_Info.csv" -f $rptFolder, (Get-UTCTime).ToString($dtmFileFormatString), $DSForestName
		$trustTable | Select-Object $ttColToExport | Export-Csv -Path $outputFile -NoTypeInformation
		
		Write-Verbose ("[{0} UTC] Exporting results data in Excel format, please wait..." -f $(Get-UTCTime).ToString($dtmFormatString))
		$wsName = "AD Trust Configuration"
		$xlParams = @{
			Path	        = $outputFile = "{0}\{1}_{2}_Active_Directory_Forest_Trust_Info.xlsx" -f $rptFolder, $(Get-UTCTime).ToString($dtmFileFormatString), $DSForestName
			WorkSheetName = $wsName
			TableStyle = 'Medium15'
			StartRow = 2
			StartColumn = 1
			AutoSize = $true
			AutoFilter = $true
			BoldTopRow = $true
			FreezeTopRow = $true
			PassThru = $true
		}
		
		$xl = $trustTable | Select-Object $ttColToExport | Export-Excel @xlParams
		$Sheet = $xl.Workbook.Worksheets["AD Trust Configuration"]
		Set-ExcelRange -Range $Sheet.Cells["A2:Z2"] -WrapText -HorizontalAlignment Center -VerticalAlignment Center -AutoFit
		$cols = $Sheet.Dimension.Columns
		Set-ExcelRange -Range $Sheet.Cells["A3:Z$($cols)"] -Wraptext -HorizontalAlignment Left -VerticalAlignment Bottom
		Export-Excel -ExcelPackage $xl -WorksheetName $wsName -Title "Active Directory Trust Configuration" -TitleBold -TitleSize 16
	} #end If
	
}
#EndRegion