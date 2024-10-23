﻿#Requires -Module ActiveDirectory, ImportExcel, HelperFunctions
#Requires -Version 5
#Requires -RunAsAdministrator
<#
	.SYNOPSIS
		Export all fine-grained password policies from AD
	
	.DESCRIPTION
		This script will query all domains in the specified AD forest or local AD forest (if one is not specified) for any fine-grained password policies in each domain and report back on the status of that policy.
	
	.PARAMETER ForestName
		AD forest name
	
	.PARAMETER Credential
		ACredential for authenticating to remote AD forest.
	
	.EXAMPLE
		PS> Export-ActiveDirectoryDefaultPassPolicies.ps1
		
	.EXAMPLE
		PS> Export-ActiveDirectoryDefaultPassPolicies.ps1 -ForestName myforest.com -Credential (Get-Credential)
	
	.OUTPUTS
		OfficeOpenXml.ExcelPackage
	
	.NOTES
		THIS CODE IS MADE AVAILABLE AS IS, WITHOUT WARRANTY OF ANY KIND. THE
		RISK OF THE USE OR THE RESULTS FROM THE USE OF THIS CODE REMAINS WITH
		THE USER.
	
	.LINK
		https://www.powershellgallery.com/packages/ImportExcel/
		
	.LINK
		https://www.powershellgallery.com/packages/HelperFunctions/
		
#>


[CmdletBinding()]
param
(
[Parameter(Mandatory = $false,
		 HelpMessage = 'Enter AD forest name.')]
[ValidateNotNullOrEmpty()]
[ValidateNotNull()]
[string]$ForestName,
[Parameter(Mandatory = $false,
		 HelpMessage = 'Enter credential for remote forest.')]
[ValidateNotNull()]
[System.Management.Automation.PsCredential]
[System.Management.Automation.Credential()]
$Credential = [System.Management.Automation.PSCredential]::Empty
)


#region Modules
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

#endregion

#region Variables
$dtmFormatString = "yyyy-MM-dd HH:mm:ss"
$dtmFileFormatString = "yyyy-MM-dd_HH-mm-ss"
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

$rptFolder = 'E:\Reports'
#endregion



#region Script
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

	try
	{
		$ForestParams = @{
			ErrorAction = 'Stop'
		}
	
		if (($PSBoundParameters.ContainsKey('ForestName')) -and ($null -ne $PSBoundParameters["ForestName"]))
		{
			$ForestParams.Add('Identity',$PSBoundParameters["ForestName"])
		}
		else
		{
			$ForestName = [System.DirectoryServices.ActiveDirectory.Forest]::GetCurrentForest()
			$ForestParams.Add('Current','LocalComputer')
		}
	
		if (($PSBoundParameters.ContainsKey('Credential')) -and ($null -ne $PSBoundParameters["Credential"]))
		{
			$ForestParams.Add('AuthType', 'Negotiate')
			$ForestParams.Add('Credential', $Credential)
		}
	
		$DSForest = Get-ADForest @ForestParams
		$DSForestName = $DSForest.Name.ToString().ToUpper()
	}
	catch
	{
		$errorMessage = "{0}: {1}" -f $Error[0], $Error[0].InvocationInfo.PositionMessage
		Write-Error $errorMessage -ErrorAction Stop
	}
	
	$domPPTblName = "$($DSForestName)_Domain_Password_Policies"
	$dtPPHeaders = ConvertFrom-Csv -InputObject $dtPPHeadersCsv
	
	try
	{
		$domPPTable = Add-DataTable -TableName $domPPTblName -ColumnArray $dtPPHeaders -ErrorAction Stop
	}
	catch
	{
		$errorMessage = "{0}: {1}" -f $Error[0], $Error[0].InvocationInfo.PositionMessage
		Write-Error $errorMessage -ErrorAction Stop
	}
	
	$Domains = ($DSForest).Domains
	
	$dCount = 1
	foreach ($DomainName in $Domains)
	{
		$ActivityMessage = "Gathering AD Domain Default Password Policy information, please wait..."
		$ProcessingStatus = "Processing Domain {0} of {1}: {2}" -f $dCount, $Domains.count, $DomainName.ToString()
		$percentComplete = ($dCount / $Domains.Count * 100)
		Write-Progress -Activity $ActivityMessage -Status $ProcessingStatus -PercentComplete $percentComplete -Id 1
		
		$domainParams = @{
			Identity = $DomainName
			Server = $DomainName
		}
		
		if (($PSBoundParameters.ContainsKey('Credential')) -and ($null -ne $PSBoundParameters["Credential"]))
		{
			$domainParams.Add('AuthType', 'Negotiate')
			$domainParams.Add('Credential', $Credential)
		}
		
		try
		{
			$Domain = Get-ADDomain @domainParams -ErrorAction SilentlyContinue
			if ($? -eq $false)
			{
				try
				{
					$Domain = Get-ADDomain @domainParams -ErrorAction Stop
				}
				catch
				{
					$errorMessage = "{0}: {1}" -f $Error[0], $Error[0].InvocationInfo.PositionMessage
					Write-Error $errorMessage -ErrorAction Stop
				}   
			}
		}
		catch
		{
			$errorMessage = "{0}: {1}" -f $Error[0], $Error[0].InvocationInfo.PositionMessage
			Write-Error $errorMessage -ErrorAction Stop
		}
		
		if ($null -ne $Domain.Name)
		{
			$pdcFSMO = $Domain.pdcEmulator
			$domDNS = $Domain.DNSRoot
			$domainDN = $Domain.DistinguishedName
			
			#Region Domain Password Policies
			try
			{
				$ppParams = @{
					Identity = $domainDN
				}
				
				if (($PSBoundParameters.ContainsKey('Credential')) -and ($null -ne $PSBoundParameters["Credential"]))
				{
					$ppParams.Add('AuthType', 'Negotiate')
					$ppParams.Add('Credential', $Credential)
				}
				
				$defPP = Get-ADDefaultDomainPasswordPolicy @ppParams -Server $pdcFSMO -ErrorAction SilentlyContinue
				if ($? -eq $false)
				{
					try
					{
						$defPP = Get-ADDefaultDomainPasswordPolicy @ppParams -Server $domDns -ErrorAction Stop
					}
					catch
					{
						$errorMessage = "{0}: {1}" -f $Error[0], $Error[0].InvocationInfo.PositionMessage
						Write-Error $errorMessage -ErrorAction Continue
					}
				}
			}
			catch
			{
				$errorMessage = "{0}: {1}" -f $Error[0], $Error[0].InvocationInfo.PositionMessage
				Write-Error $errorMessage -ErrorAction Continue
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
			
			$null = $domainDN = $domain = $domDns = $pdcFSMO
			$null = $defPP = $domDN = $complexityEnabled = $lockoutDuration = $lockoutThreshold = $lockoutWindow = $maxPWAge = $minPWAge = $minPWLength = $pwHistoryCount = $encryptionEnabled
			#EndRegion
			
		}#end if $Domain.Name
		$dCount++
	}
	
	Write-Progress -Activity "Done gathering AD domain password policy information for $($DSForestName)" -Status "Ready" -Completed
	
	$ColToExport = $dtPPHeaders.ColumnName

	$outputFile = "{0}\{1}-{2}-Default-Domain-Password-Policies.csv" -f $rptFolder, (Get-UTCTime).ToString($dtmFileFormatString), $DSForestName
	$domPPTable | Select-Object $ColToExport | Export-Csv -Path $outputFile -NoTypeInformation

	Write-Verbose -Message ("[{0} UTC] Exporting data tables to Excel spreadsheet tabs." -f (Get-UTCTime).ToString($dtmFormatString))
	[String]$wsName = "AD Domain PP Config"
	$ExcelParams = @{
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

	$xl = $domPPTable | Select-Object $ColToExport | Sort-Object -Property "Domain Name" | Export-Excel @ExcelParams
	$Sheet = $xl.Workbook.Worksheets["AD Domain PP Config"]
	Set-ExcelRange -Range $Sheet.Cells["A2:Z2"] -WrapText -HorizontalAlignment Center -VerticalAlignment Center -AutoFit
	$cols = $Sheet.Dimension.Columns
	Set-ExcelRange -Range $Sheet.Cells["A3:Z$($cols)"] -Wraptext -HorizontalAlignment Left -VerticalAlignment Bottom
	Export-Excel -ExcelPackage $xl -WorksheetName $wsName -Title "Active Directory Domain Password Policies" -TitleBold -TitleSize 16
	[System.GC]::GetTotalMemory('ForceFullCollection') | Out-Null	
}
catch
{
	$errorMessage = "{0}: {1}" -f $Error[0], $Error[0].InvocationInfo.PositionMessage
	Write-Error $errorMessage -ErrorAction Stop
}
#endregion