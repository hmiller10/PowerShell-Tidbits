#Requires -Module  ActiveDirectory, ImportExcel, HelperFunctions
#Requires -Version 5
#Requires -RunAsAdministrator
<#
	.SYNOPSIS
		Get AD password policies
	
	.DESCRIPTION
		This script executes an AD PowerShell cmdlet to gather the default domain
		password policies and exports the results to an Excel spreadsheet.
	
	.PARAMETER ForestName
		The name of the Active Directory Forest to gather default password policies for.
	
	.PARAMETER DomainName
		Enter the name of the Active Directory domain to gather the default password policy on.
	
	.PARAMETER Credential
		Enter the Credential Object.
	
	.EXAMPLE
		.\Export-ADDomainPasswordPolicies.ps1 -ForestName exampleforest.com -Credential (Get-Credential)
		
	.EXAMPLE
		.\Export-ADDomainPasswordPolicies.ps1 -DomainName exampledomain.com -Credential (Get-Credential)
		
	.EXAMPLE
		.\Export-ADDomainPasswordPolicies.ps1
	
	.OUTPUTS
		Excel spreadsheet with the default password policy information settings
	
	.NOTES
		THIS CODE IS MADE AVAILABLE AS IS, WITHOUT WARRANTY OF ANY KIND. THE
		ENTIRE RISK OF THE USE OR THE RESULTS FROM THE USE OF THIS CODE REMAINS
		WITH THE USER.
#>

###########################################################################
#
#
# AUTHOR:  Heather Miller
#
# VERSION HISTORY: 3.0 - added Error handling and credential support
# 
############################################################################

[CmdletBinding(DefaultParameterSetName = 'ForestParamSet',
			SupportsShouldProcess = $true)]
param
(
	[Parameter(ParameterSetName = 'ForestParamSet',
			 Position = 0,
			 HelpMessage = 'Enter the name of AD forest.')]
	[ValidateNotNullOrEmpty()]
	[String]$ForestName,
	[Parameter(ParameterSetName = 'DomainParamSet',
			 Position = 0,
			 HelpMessage = 'Enter the AD domain name.')]
	[ValidateNotNullOrEmpty()]
	[String]$DomainName,
	[Parameter(ParameterSetName = 'ForestParamSet',
			 Position = 1,
			 HelpMessageBaseName = 'Add the credential object variable name.')]
	[Parameter(ParameterSetName = 'DomainParamSet',
			 Position = 1,
			 HelpMessage = 'Add the credential object variable name.')]
	[pscredential]$Credential
)

#Region Modules
Try 
{
	Import-Module -Name ActiveDirectory -SkipEditionCheck -Force -ErrorAction Stop
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
	Import-Module -Name ImportExcel -Force
}
catch
{
	try
	{
		$moduleName = 'ImportExcel'
		$ErrorActionPreference = 'Stop';
		$module = Get-Module -ListAvailable | Where-Object { $_.Name -eq $moduleName };
		$ErrorActionPreference = 'Continue';
		$modulePath = Split-Path $module.Path;
		$psdPath = "{0}\{1}" -f $modulePath, "ImportExcel.psd1"
		Import-Module $psdPath -ErrorAction Stop
	}
	catch
	{
		Throw "ImportExcel PS module could not be loaded. $($_.Exception.Message)"
	}
}
   
#EndRegion


#Region Variables
   
$dtPPHeadersCsv = 
@"
ColumnName,DataType
"Domain Name",string
"Complexity Enabled",string
"Lockout Duration",string
"Lockout Window",string
"Lockout Threshold",string
"Max Password Age",string
"Min Password Age",string
"Min Password Length",string
"Password History Count",string
"Reversible Encryption Enabled",string
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

	$dtPPHeaders = ConvertFrom-Csv -InputObject $dtPPHeadersCsv
	
	$tblName = "Domain_Password_Policies"
	$domPPTable = Add-DataTable -TableName $tblName -ColumnArray $dtPPHeaders
	
	$domainProperties = @("DistinguishedName", "DNSRoot", "Forest", "InfrastructureMaster", "Name", "PDCEmulator")
	
	Switch ($PSCmdlet.ParameterSetName)
	{
		"ForestParamSet"
		{
			$Domains = @()
			
			#Get AD Forest Basic Information
			$forestProperties = @("ApplicationPartitions", "Domains", "DomainNamingMaster", "ForestMode", "Name", "RootDomain", "PartitionsContainer", "SchemaMaster", "SPNSuffixes", "UPNSuffixes")
			
			$forestParams = @{
				ErrorAction = 'Stop'
			}
			
			if (($PSBoundParameters.ContainsKey('ForestName')) -and ($null -ne ($PSBoundParameters["ForestName"])))
			{
				$forestParams.Add('Identity',$ForestName)
				
				if (($PSBoundParameters.ContainsKey('Credential')) -and ($null -ne ($PSBoundParameters["Credential"])))
				{
					$forestParams.Add('Credential',$Credential)
				}
			}
			
			$DSForest = Get-ADForest @forestParams -Server (Get-ADForest @forestParams).SchemaMaster | Select-Object -Property $forestProperties	
			$DSForestName = ($DSForest).Name.ToString().ToUpper()
			$Domains = ($DSForest).Domains
				
			foreach ($Domain in $Domains)
			{
				$domainParams = @{
					Identity = $Domain
					ErrorAction = 'Continue'
				}
				
				if (($PSBoundParameters.ContainsKey('Credential')) -and ($null -ne ($PSBoundParameters["Credential"])))
				{
					$domainParams.Add('Credential',$Credential)
				}
				
				$domainInfo = Get-ADDomain @domainParams -Server (Get-ADDomain @domainParams).pdcEmulator | Select-Object -Property $domainProperties
				if ($null -ne $domainInfo.distinguishedName)
				{
					$domainDN = ($domainInfo).distinguishedName
					$domDns= ($domainInfo).DNSRoot
					$pdcFSMO = ($domainInfo).PDCEmulator
				}
				
				#Region Domain Password Policies
				try
				{
					$defPP = Get-ADDefaultDomainPasswordPolicy -Identity $domainDN -Server $pdcFSMO -ErrorAction 'Stop'
					if ($? -eq $false)
					{
						$defPP = Get-ADDefaultDomainPasswordPolicy -Identity $domainDN -Server $domDns -ErrorAction 'Continue'
					}
				}
				catch
				{
					$errorMessage = "{0}: {1}" -f $Error[0], $Error[0].InvocationInfo.PositionMessage
					Write-Error $errorMessage -ErrorAction Stop
				}
				
				[String]$domDN = ($defPP).distinguishedName
				[String]$complexityEnabled = ($defPP).ComplexityEnabled
				[String]$lockoutDuration = ($defPP).LockoutDuration
				[String]$lockoutWindow = ($defPP).LockoutObservationWindow
				[String]$lockoutThreshold = ($defPP).LockoutThreshold
				[String]$maxPWAge = ($defPP).MaxPasswordAge
				[String]$minPWAge = ($defPP).MinPasswordAge
				[String]$minPWLength = ($defPP).MinPasswordLength
				[String]$pwHistoryCount = ($defPP).PasswordHistoryCount
				[String]$encryptionEnabled = ($defPP).ReversibleEncryptionEnabled
				
				$dtRow = $domPPTable.NewRow()
				$dtRow."Domain Name" = $domDns
				$dtRow."Complexity Enabled" = $complexityEnabled
				$dtRow."Lockout Duration" = $lockoutDuration
				$dtRow."Lockout Window" = $lockoutWindow
				$dtRow."Lockout Threshold" = $lockoutThreshold
				$dtRow."Max Password Age" = $maxPWAge
				$dtRow."Min Password Age" = $minPWAge
				$dtRow."Min Password Length" = $minPWLength
				$dtRow."Password History Count" = $pwHistoryCount
				$dtRow."Reversible Encryption Enabled" = $encryptionEnabled
				
				$domPPTable.Rows.Add($dtRow)
				
				$null = $defPP = $domDN = $complexityEnabled = $lockoutDuration = $lockoutThreshold = $lockoutWindow = $maxPWAge = $minPWAge = $minPWLength = $pwHistoryCount = $encryptionEnabled
				#EndRegion
				
				$null = $Domain = $domainInfo = $domainDN = $domDns = $pdcFSMO
				[System.GC]::GetTotalMemory('ForceFullCollection') | Out-Null
			}
			$ScopeVariable = $DSForestName
			
		}
		"DomainParamSet"
		{
			$domainParams = @{
				Identity    = $DomainName
				ErrorAction = 'Continue'
			}
			
			if (($PSBoundParameters.ContainsKey('Credential')) -and ($null -ne ($PSBoundParameters["Credential"])))
			{
				$domainParams.Add('Credential', $Credential)
			}
			
			$domainInfo = Get-ADDomain @domainParams -Server (Get-ADDomain @domainParams).pdcEmulator | Select-Object -Property $domainProperties
			if ($null -ne $domainInfo.distinguishedName)
			{
				$domainDN = ($domainInfo).distinguishedName
				$domDns = ($domainInfo).DNSRoot
				$pdcFSMO = ($domainInfo).PDCEmulator
			}
			
			#Region Domain Password Policies
			try
			{
				$defPP = Get-ADDefaultDomainPasswordPolicy -Identity $domainDN -Server $pdcFSMO -ErrorAction 'Stop'
				if ($? -eq $false)
				{
					$defPP = Get-ADDefaultDomainPasswordPolicy -Identity $domainDN -Server $domDns -ErrorAction 'Continue'
				}
			}
			catch
			{
				$errorMessage = "{0}: {1}" -f $Error[0], $Error[0].InvocationInfo.PositionMessage
				Write-Error $errorMessage -ErrorAction Stop
			}
			
			[String]$domDN = ($defPP).distinguishedName
			[String]$complexityEnabled = ($defPP).ComplexityEnabled
			[String]$lockoutDuration = ($defPP).LockoutDuration
			[String]$lockoutWindow = ($defPP).LockoutObservationWindow
			[String]$lockoutThreshold = ($defPP).LockoutThreshold
			[String]$maxPWAge = ($defPP).MaxPasswordAge
			[String]$minPWAge = ($defPP).MinPasswordAge
			[String]$minPWLength = ($defPP).MinPasswordLength
			[String]$pwHistoryCount = ($defPP).PasswordHistoryCount
			[String]$encryptionEnabled = ($defPP).ReversibleEncryptionEnabled
			
			$dtRow = $domPPTable.NewRow()
			$dtRow."Domain Name" = $domDns
			$dtRow."Complexity Enabled" = $complexityEnabled
			$dtRow."Lockout Duration" = $lockoutDuration
			$dtRow."Lockout Window" = $lockoutWindow
			$dtRow."Lockout Threshold" = $lockoutThreshold
			$dtRow."Max Password Age" = $maxPWAge
			$dtRow."Min Password Age" = $minPWAge
			$dtRow."Min Password Length" = $minPWLength
			$dtRow."Password History Count" = $pwHistoryCount
			$dtRow."Reversible Encryption Enabled" = $encryptionEnabled
			
			$domPPTable.Rows.Add($dtRow)
			
			$ScopeVariable = $domDns.ToString().ToUpper()
			
			$null = $domainDN = $domainInfo = $domDns = $pdcFSMO
			$null = $defPP = $domDN = $complexityEnabled = $lockoutDuration = $lockoutThreshold = $lockoutWindow = $maxPWAge = $minPWAge = $minPWLength = $pwHistoryCount = $encryptionEnabled
			#EndRegion
			
			$null = $DomainName = $domainInfo = $domainDN = $domDns = $pdcFSMO
			[System.GC]::GetTotalMemory('ForceFullCollection') | Out-Null
		}
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
	$driveRoot = (Get-Location).Drive.Root
	$rptFolder = "{0}{1}" -f $driveRoot, "Reports"
	
	Test-PathExists -Path $rptFolder -PathType Folder
	
	$colToExport = $dtPPHeaders.ColumnName
	
	Write-Verbose ("[{0} UTC] Exporting results data to CSV, please wait..." -f $(Get-UTCTime).ToString($dtmFormatString))
	$outputFile = "{0}\{1}-{2}-Domain-Password-Policies.csv" -f $rptFolder, $(Get-UTCTime).ToString($dtmFileFormatString), $ScopeVariable
	$domPPTable | Select-Object $colToExport | Export-Csv -Path $outputFile -NoTypeInformation
	
	Write-Verbose -Message ("[{0} UTC] Exporting data tables to Excel spreadsheet tabs." -f $(Get-UTCTime).ToString($dtmFormatString))
	[String]$wsName = "AD Domains PP Config"
	$xlParams = @{
		Path	        = $outputFile.ToString().Replace([System.IO.Path]::GetExtension($outputFile), ".xlsx")
		WorkSheetName = $wsName
		TableStyle    = 'Medium15'
		StartRow	    = 2
		StartColumn   = 1
		AutoSize	    = $true
		AutoFilter    = $true
		BoldTopRow    = $true
		FreezeTopRow  = $true
		PassThru	    = $true
	}

	$xl = $domPPTable | Select-Object $colToExport | Sort-Object -Property "Domain Name" | Export-Excel @xlParams
	$Sheet = $xl.Workbook.Worksheets["AD Domains PP Config"]
	Set-ExcelRange -Range $Sheet.Cells["A2:Z2"] -WrapText -HorizontalAlignment Center -VerticalAlignment Center -AutoFit
	$cols = $Sheet.Dimension.Columns
	Set-ExcelRange -Range $Sheet.Cells["A3:Z$($cols)"] -Wraptext -HorizontalAlignment Left -VerticalAlignment Bottom
	Export-Excel -ExcelPackage $xl -WorksheetName $wsName -Title "Active Directory Domain Password Policies" -TitleBold -TitleSize 16
}

#EndRegion