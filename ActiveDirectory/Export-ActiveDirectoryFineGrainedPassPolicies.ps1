#Requires -Module  ActiveDirectory, ImportExcel, HelperFunctions
#Requires -Version 5
#Requires -RunAsAdministrator
<#
	.SYNOPSIS
		Export all fine-grained password policies from AD
	
	.DESCRIPTION
		This script will query all domains in the specified AD forest or local AD forest (if one is not specified) for any fine-grained password policies in each domain and report back on the status of that policy.
	
	.PARAMETER ForestName
		Fully qualified domain name of AD forest where domain fine-grained password policies should be documented.
		
	.PARAMETER Credential
		PSCredential

	.EXAMPLE
		PS> Export-ActiveDirectoryFineGrainedPasswordPolicies.ps1
		
	.EXAMPLE
		PS> Export-ActiveDirectoryFineGrainedPasswordPolicies.ps1 -ForestName myForest.com -Credential (Get-Credential)	
	
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
	[Parameter(Mandatory = $false)]
	[ValidateNotNullOrEmpty()]
	[string]
	$ForestName,
	[Parameter(Mandatory = $false)]
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
$dtfgPPHeadersCsv =
@"
ColumnName,DataType
"Domain Name",string
"Applied To",string
"Policy Name",string
"Complexity Enabled",string
"Lockout Duration",string
"Lockout Window",string
"Lockout Threshold",string
"Max Password Age",string
"Min Password Age",string
"Min Password Length",string
"Password History Count",string
"Reversible Encryption Enabled",string
"WhenCreated",string
"WhenChanged",string
"@

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
	   
	$domfgPPTblName = "{0}_Domain_Fine-Grained_Password_Policies" -f $Forest.Name
	$dtfgPPHeaders = ConvertFrom-Csv -InputObject $dtfgPPHeadersCsv
	try
	{
		$domfgPPTable = Add-DataTable -TableName $domfgPPTblName -ColumnArray $dtfgPPHeaders -ErrorAction Continue
	}
	catch
	{
		$errorMessage = "{0}: {1}" -f $Error[0], $Error[0].InvocationInfo.PositionMessage
		Write-Error $errorMessage -ErrorAction Continue
	}
	
	try
	{
		if ($null -eq ($PSBoundParameters["ForestName"]))
		{
			$ForestName = [System.DirectoryServices.ActiveDirectory.Forest]::GetCurrentForest()
		}
		else
		{
			$ForestName = $PSBoundParameters["ForestName"]
		}
	}
	catch
	{
		$errorMessage = "{0}: {1}" -f $Error[0], $Error[0].InvocationInfo.PositionMessage
		Write-Error $errorMessage -ErrorAction Stop
	}
	
	foreach ($Forest in $ForestName)
	{
		$ForestParams = @{
			ErrorAction = 'Stop'
		}

		if (($PSBoundParameters.ContainsKey('Credential')) -and ($null -ne $PSBoundParameters["Credential"]))
		{
			$ForestParams.Add('AuthType', 'Negotiate')
			$ForestParams.Add('Credential', $Credential)
		}
		
		$DSForest = Get-ADForest -Identity $Forest @ForestParams
		$DSForestName = $DSForest.Name.ToString().ToUpper()
		
		$Domains = ($DSForest).Domains
		$dCount = 1
		foreach ($DomainName in $Domains)
		{
			$ActivityMessage = "Gathering AD fine-grained password policy information, please wait..."
			$ProcessingStatus = "Processing Domain {0} of {1}: {2}" -f $dCount, $Domains.count, $DomainName.ToString()
			$percentComplete = ($dCount / $Domains.Count * 100)
			Write-Progress -Activity $ActivityMessage -Status $ProcessingStatus -PercentComplete $percentComplete -Id 1
			
			$domainParams = @{
				Identity = $DomainName
				Server   = $DomainName
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
				
				#Region Domain FineGrained Password Policies
				Write-Verbose -Message ("Working on fine-grained password policies for domain: {0}" -f $domDNS)
				
				try
				{
					$fgParams = @{
						Filter	    = '*'
						Properties    = '*'
						SearchBase    = $domainDN
						SearchScope   = 'Subtree'
						ResultSetSize = $null
					}
					
					if (($PSBoundParameters.ContainsKey('Credential')) -and ($null -ne $PSBoundParameters["Credential"]))
					{
						$fgParams.Add('AuthType', 'Negotiate')
						$fgParams.Add('Credential', $Credential)
					}
					
					$fgPP = Get-ADFineGrainedPasswordPolicy @fgParams -Server $pdcFSMO -ErrorAction SilentlyContinue
					if ($? -eq $false)
					{
						try
						{
							$fgPP = Get-ADFineGrainedPasswordPolicy @fgParams -Server $domDNS -ErrorAction Stop
						}
						catch
						{
							$fgPP = "Unable to get password policy settings for {0}" -f $domDNS
							$errorMessage = "{0}: {1}" -f $Error[0], $Error[0].InvocationInfo.PositionMessage
							Write-Error $fgPP -ErrorAction Continue
							Write-Error $errorMessage -ErrorAction Continue
						}
					}
				}
				catch
				{
					$fgPP = "Unable to get password policy settings for {0}" -f $domDNS
					$errorMessage = "{0}: {1}" -f $Error[0], $Error[0].InvocationInfo.PositionMessage
					Write-Error $fgPP -ErrorAction Continue
					Write-Error $errorMessage -ErrorAction Continue
				}
				
				#Write fine-grained password policy to data table
				if ($null -ne $fgPP.CN)
				{
					foreach ($policy in $fgPP)
					{
						$dtfgPPRow = $domfgPPTable.NewRow()
						$dtfgPPRow."Domain Name" = $domDNS
						$dtfgPPRow."Policy Name" = $policy.Name
						$dtfgPPRow."Applied To" = $policy.AppliesTo | Out-String
						$dtfgPPRow."Complexity Enabled" = $policy.ComplexityEnabled.ToString()
						$dtfgPPRow."Lockout Duration" = $policy.LockoutDuration.ToString()
						$dtfgPPRow."Lockout Window" = $policy.LockoutObservationWindow.ToString()
						$dtfgPPRow."Lockout Threshold" = $policy.LockoutThreshold.ToString()
						$dtfgPPRow."Max Password Age" = $policy.MaxPasswordAge.ToString()
						$dtfgPPRow."Min Password Age" = $policy.MinPasswordAge.ToString()
						$dtfgPPRow."Min Password Length" = $policy.MinPasswordLength.ToString()
						$dtfgPPRow."Password History Count" = $policy.PasswordHistoryCount.ToString()
						$dtfgPPRow."Reversible Encryption Enabled" = $policy.ReversibleEncryptionEnabled.ToString()
						$dtfgPPRow."WhenCreated" = $policy.WhenCreated.ToString($dtmFileFormatString)
						$dtfgPPRow."WhenChanged" = $policy.WhenChanged.ToString($dtmFileFormatString)
						
						$domfgPPTable.Rows.Add($dtfgPPRow)
					}
					
				}
				else
				{
					$dtfgPPRow = $domfgPPTable.NewRow()
					$dtfgPPRow."Domain Name" = $domDNS
					$dtfgPPRow."Policy Name" = "None"
					$dtfgPPRow."Applied To" = "There are no fine-grained password policies configured for this domain: {0}" -f $domDNS
					$dtfgPPRow."Complexity Enabled" = "None"
					$dtfgPPRow."Lockout Duration" = "None"
					$dtfgPPRow."Lockout Window" = "None"
					$dtfgPPRow."Lockout Threshold" = "None"
					$dtfgPPRow."Max Password Age" = "None"
					$dtfgPPRow."Min Password Age" = "None"
					$dtfgPPRow."Min Password Length" = "None"
					$dtfgPPRow."Password History Count" = "None"
					$dtfgPPRow."Reversible Encryption Enabled" = "None"
					$dtfgPPRow."WhenCreated" = "None"
					$dtfgPPRow."WhenChanged" = "None"
					
					$domfgPPTable.Rows.Add($dtfgPPRow)
					
				}
				
				$null = $fgPP
			}
			$dCount++
		}
	}
	
	Write-Progress -Activity "Done gathering AD fine-grained password policy information" -Status "Ready" -Completed
}
catch
{
	$errorMessage = "{0}: {1}" -f $Error[0], $Error[0].InvocationInfo.PositionMessage
	Write-Error $errorMessage -ErrorAction Continue
}
finally
{
	$driveRoot = (Get-Location).Drive.Root
	$rptFolder = "{0}{1}" -f $driveRoot, "Reports"
	
	Test-PathExists -Path $rptFolder -PathType Folder
	
	$ColToExport = $dtfgPPHeaders.ColumnName
	
	$outputFile = "{0}\{1}-{2}-Finegrained-Password-Policies.csv" -f $rptFolder, (Get-UTCTime).ToString($dtmFileFormatString), $DSForestName
	$domfgPPTable | Select-Object $ColToExport | Export-Csv -Path $outputFile -NoTypeInformation
	
	Write-Verbose -Message ("[{0} UTC] Exporting data tables to Excel spreadsheet tabs." -f (Get-UTCTime).ToString($dtmFormatString))
	[String]$wsName = "AD Fine-Grained PP Config"
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
	
	$xl = $domfgPPTable | Select-Object $ColToExport | Sort-Object -Property "Domain Name" | Export-Excel @xlParams
	$Sheet = $xl.Workbook.Worksheets["AD Fine-Grained PP Config"]
	Set-ExcelRange -Range $Sheet.Cells["A2:Z2"] -WrapText -HorizontalAlignment Center -VerticalAlignment Center -AutoFit
	$cols = $Sheet.Dimension.Columns
	Set-ExcelRange -Range $Sheet.Cells["A3:Z$($cols)"] -Wraptext -HorizontalAlignment Left -VerticalAlignment Bottom
	Export-Excel -ExcelPackage $xl -WorksheetName $wsName -Title "Active Directory Fine-Grained Password Policies" -TitleBold -TitleSize 16
	[System.GC]::GetTotalMemory('ForceFullCollection') | Out-Null
}

#endregion