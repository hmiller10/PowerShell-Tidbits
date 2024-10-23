#Requires -Module ActiveDirectory, ImportExcel, GroupPolicy
#Requires -Version 5
#Requires -RunAsAdministrator
<#
	.NOTES
	THIS CODE IS MADE AVAILABLE AS IS, WITHOUT WARRANTY OF ANY KIND. THE
	RISK OF THE USE OR THE RESULTS FROM THE USE OF THIS CODE REMAINS WITH
	THE USER.

	.SYNOPSIS
		Create OU report

	.DESCRIPTION
		This script leverages the parallel processing functionality in PowerShell 7
		to process and report on the OU structure of the domain named piped to the script parameter
		
	.PARAMETER DomainName
		Fully qualified domain name of domain where OU report should be created from
		
	.PARAMETER Credential
		PSCredential
	
	.OUTPUTS
	Excel spreadsheet with OU configuration for named AD domain
	
	.EXAMPLE
	PS C:> .\Export-ActiveDirectoryOUStructures.ps1 -DomainName my.domain.com
	
	.EXAMPLE
	PS C:> .\Export-ActiveDirectoryOUStructures.ps1 -DomainName my.domain.com -Credential PSCredential

	.LINK
	https://github.com/dfinke/ImportExcel
#>

[CmdletBinding()]
param
(
[Parameter(Mandatory = $true)]
[ValidateNotNullOrEmpty()]
[string]$DomainName,
[Parameter(Mandatory = $false)]
[ValidateNotNullOrEmpty()]
[System.Management.Automation.PsCredential]$Credential
)

#Region Modules
#Check if required module is loaded, if not load import it
try
{
	Import-Module ActiveDirectory -ErrorAction Stop
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
	Import-Module -Name ImportExcel -Force -ErrorAction Stop
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
	Import-Module GroupPolicy -SkipEditionCheck -ErrorAction Stop
}
catch
{
	try
	{
		Import-Module C:\Windows\System32\WindowsPowerShell\v1.0\Modules\GroupPolicy\GroupPolicy.psd1 -ErrorAction Stop
	}
	catch
	{
		throw "Group Policy module could not be loaded. $($_.Exception.Message)"
	}
}

#EndRegion

#region Variables
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

$rptFolder = 'E:\Reports'

[PSObject[]]$ouObject = @()
$colResults = @()

#endregion

#Region Functions

function Get-ADOUPerms
{
<#
	.SYNOPSIS
		Retrieves delegation information for Organizational Units (OUs) in Active Directory.
	
	.DESCRIPTION
		The Get-OUDelegations function fetches and displays delegation details for OUs in Active Directory.
		It allows filtering by OU name, specific Active Directory rights, and includes an option for verbose output.
	
	.PARAMETER OUDistinguishedName
		The AD object distinguishedName attribute value.
	
	.PARAMETER DomainController
		The name of the domain controller to connect to in order to find the OU properties in AD.
	
	.PARAMETER RightsFilter
		An array of strings to filter the results by specific Active Directory rights.
		Valid options include GenericAll, GenericRead, GenericWrite, CreateChild, DeleteChild, ListChildren,
		Self, ReadProperty, WriteProperty, DeleteTree, ListObject, ExtendedRight, Delete, ReadControl,
		WriteDacl, and WriteOwner. If not specified, no filtering on rights is applied.
	
	.PARAMETER VerboseOutput
		A switch parameter that enables verbose output. When used, additional details about the operation's progress are displayed.
	
	.PARAMETER Credential
		If available, add PSCredential variable.
	
	.EXAMPLE
		Get-OUDelegations -OUDistinguishedName "ou=MyOU,dc=domain,dc=com" -VerboseOutput
		This example retrieves delegation information for OUs that start with "Sales" and displays verbose output.
	
	.EXAMPLE
		Get-OUDelegations -OUDistinguishedName "ou=MyOU,dc=domain,dc=com" -RightsFilter GenericRead,GenericWrite
		This example retrieves delegation information for OUs where the delegations include either GenericRead or GenericWrite permissions.
	
	.NOTES
		Requires the Active Directory module to be installed and available.
		The user running this command must have permissions to read Active Directory and OU objects.
	
	.LINK
		Get-ADOrganizationalUnit
		Get-Acl
#>
	
	[CmdletBinding(DefaultParameterSetName = 'System.Management.Automation.PSCustomObject')]
	param
	(
		[Parameter(Mandatory = $true,
		           Position = 0,
		           HelpMessage = 'Specify OU distinguishedName.')]
		[string]
		$OUDistinguishedName,
		[Parameter(Mandatory = $true,
		           Position = 1,
		           HelpMessage = 'DC to use to search OUs')]
		[string]
		$DomainController,
		[Parameter(Mandatory = $false,
		           Position = 2,
		           HelpMessage = 'Filter by specific Active Directory rights.')]
		[ValidateSet('GenericAll', 'GenericRead', 'GenericWrite', 'CreateChild', 'DeleteChild', 'ListChildren', 'Self', 'ReadProperty', 'WriteProperty', 'DeleteTree', 'ListObject', 'ExtendedRight', 'Delete', 'ReadControl', 'WriteDacl', 'WriteOwner')]
		[string[]]
		$RightsFilter,
		[Parameter(Mandatory = $false,
		           Position = 3,
		           HelpMessage = 'Enable verbose output.')]
		[switch]
		$VerboseOutput,
		[Parameter(Position = 4,
		           HelpMessage = 'If available, add PSCredential variable.')]
		[ValidateNotNullOrEmpty()]
		[pscredential]
		$Credential
	)
	
begin
	{
		Import-Module -Name ActiveDirectory -Force -ErrorAction Stop
		
		# Initialize result array
		$Result = @()
		
	}
	Process
	{
		$ouParams = @{
			Identity    = $OUDistinguishedName
			Server	  = $DomainController
			ErrorAction = 'Stop'
		}
		
		if (($PSBoundParameters.ContainsKey('Credential')) -and ($null -ne $PSBoundParameters["Credential"]))
		{
			$ouParams.Add('AuthType', 'Negotiate')
			$ouParams.Add('Credential', $Credential)
		}
		
		try
		{
			$OU = Get-ADOrganizationalUnit @ouParams
		}
		catch
		{
			$errorMessage = "{0}: {1}" -f $Error[0], $Error[0].InvocationInfo.PositionMessage
			Write-Error $errorMessage -ErrorAction Continue
		}
		
		$Path = "AD:\$OUDistinguishedName"
		$ACLs = (Get-Acl -Path $Path).Access
		
		# Process each ACL
		ForEach ($ACL in $ACLs)
		{
			If ($ACL.IsInherited -eq $False)
			{
				$Rights = $ACL.ActiveDirectoryRights.ToString().Split(", ")
				if (-not $RightsFilter -or ($RightsFilter | ForEach-Object { $_ -in $Rights }))
				{
					# Create custom PSObject
					$IdentityReference = try
					{
						(New-Object System.Security.Principal.SecurityIdentifier($ACL.IdentityReference.Value)).Translate([System.Security.Principal.NTAccount]).Value
					}
					catch
					{
						$ACL.IdentityReference.Value
					}
					
					$Delegation = [PSCustomObject]@{
						OU			       = $OU.DistinguishedName
						IdentityReference     = $IdentityReference
						ActiveDirectoryRights = $ACL.ActiveDirectoryRights
						AccessControlType     = $ACL.AccessControlType
					}
					$Result += $Delegation
				}
			}
		}
		
		if ($VerboseOutput)
		{
			Write-Verbose "Processed OU: $($OU.DistinguishedName)"
		}
	}
	End
	{
		# Return results as PSObjects
		return $Result	
	}
}#end function Get-ADOUPerms

#EndRegion



#Region Script
$Error.Clear()

$domainParams = @{
	Identity    = $DomainName
	Server	  = $DomainName
	ErrorAction = 'Stop'
}

if (($PSBoundParameters.ContainsKey('Credential')) -and ($null -ne $PSBoundParameters["Credential"]))
{
	$domainParams.Add('AuthType', 'Negotiate')
	$domainParams.Add('Credential', $Credential)
}

try
{
	$Domain = Get-ADDomain @domainParams | Select-Object -Property distinguishedName, DnsRoot, Name, pdcEmulator
}
catch
{
	$errorMessage = "{0}: {1}" -f $Error[0], $Error[0].InvocationInfo.PositionMessage
	Write-Error $errorMessage -ErrorAction Stop
}


if ($null -ne $Domain.DistinguishedName)
{
	$domainDN = $Domain.DistinguishedName
	$domDNS = $Domain.dnsRoot
	$pdcFSMO = $Domain.pdcEmulator
}



#List properties to be collected into array for writing to OU tab
$OUs = @()
$ouProps = @("distinguishedName", "gpLink", "LinkedGroupPolicyObjects", "ManagedBy", "Name", "ntSecurityDescriptor", "objectCategory", "objectClass", "ParentGUID", "sDRightsEffective", "whenCreated", "whenChanged")

Write-Verbose -Message ("Gathering collection of AD Organizational Units for {0}" -f $Domain.Name)
try
{
	$ouParams = @{
		Filter	    = '*'
		Properties    = $ouProps
		SearchBase    = $domainDN
		SearchScope   = 'Subtree'
		ResultSetSize = $null
	}
	
	if (($PSBoundParameters.ContainsKey('Credential')) -and ($null -ne ($PSBoundParameters["Credential"])))
	{
		$ouParams.Add('AuthType', 'Negotiate')
		$ouParams.Add('Credential', $Credential)
	}
	
	try
	{
		$OUs = Get-ADOrganizationalUnit @ouParams -Server $pdcFSMO -ErrorAction SilentlyContinue
	}
	catch
	{
		$OUs = Get-ADOrganizationalUnit @ouParams -Server $domDNS -ErrorAction Stop
	}
	
	$OUs.ForEach({
		
		$OU = $_
		$ouGPOs = @()
		$ouChildNames = @()
		
		$ouDN = ($OU).distinguishedName
		$ouCreated = ($OU).whenCreated
		$ouLastModified = ($OU).whenChanged
		
		try
		{
			Write-Verbose -Message ("Working on Organizational Unit: {0}" -f $ouDN)
			#Convert the parentGUID attribute (stored as a byte array) into a proper-job GUID
			$ParentGuid = ([GUID]$Ou.ParentGuid).Guid
			
			#Attempt to retrieve the object referenced by the parent GUID
			$objParams = @{
				Identity = $ParentGuid
			}
			
			if (($PSBoundParameters.ContainsKey('Credential')) -and ($null -ne ($PSBoundParameters["Credential"])))
			{
				$objParams.Add('AuthType', 'Negotiate')
				$objParams.Add('Credential', $Credential)
			}
			
			$ParentObject = Get-ADObject @objParams -Server $pdcFSMO -ErrorAction SilentlyContinue
			if ($? -eq $False)
			{
				$ParentObject = Get-ADObject @objParams -Server $domDNS -ErrorAction SilentlyContinue
			}
		}
		catch
		{
			Write-Warning ("Error occurred geting parent OU information for: {0}" -f $ouDN)
			$errorMessage = "{0}: {1}" -f $Error[0], $Error[0].InvocationInfo.PositionMessage
			Write-Error $errorMessage -ErrorAction Continue
		}
		
		try
		{
			Write-Verbose -Message ("Examining Sub-OUs of {0}" -f $ouDN)
			$childOUParams = @{
				LDAPFilter    = '(objectClass=organizationalUnit)'
				Properties    = 'DistinguishedName'
				SearchBase    = $ouDN
				SearchScope   = 'OneLevel'
				ResultSetSize = $null
			}
			
			if (($PSBoundParameters.ContainsKey('Credential')) -and ($null -ne ($PSBoundParameters["Credential"])))
			{
				$childOUParams.Add('AuthType', 'Negotiate')
				$childOUParams.Add('Credential', $Credential)
			}
			
			[Array]$ouChildNames = (Get-ADOrganizationalUnit @childOUParams -Server $pdcFSMO -ErrorAction SilentlyContinue).DistinguishedName
			if ($? -eq $false)
			{
				[Array]$ouChildNames = (Get-ADOrganizationalUnit @childOUParams -Server $domDNS -ErrorAction Stop).DistinguishedName
			}
			
		}
		catch
		{
			Write-Warning -Message ("Error occurred get list of child OUs for {0}." -f $ouDN)
			$errorMessage = "{0}: {1}" -f $Error[0], $Error[0].InvocationInfo.PositionMessage
			Write-Error $errorMessage -ErrorAction Continue
		}
		
		
		if ($ouChildNames.Count -ge 1)
		{
			$ChildOUs = [String]($ouChildNames -join "`n")
		}
		else
		{
			$ChildOUs = "None"
		}
		
		if ($null -ne $OU.ManagedBy)
		{
			$ouMgr = ($OU).ManagedBy
		}
		else
		{
			$ouMgr = "None listed for this OU."
		}
		
		Write-Verbose -Message ("Gathering list of group policies linked to {0}." -f $ouDN)
		try
		{
			$ouGPOs = $OU | Select-Object -ExpandProperty LinkedGroupPolicyObjects
			if ($ouGPOs.Count -ge 1)
			{
				try
				{
					$ouGPONames = $OU | Select-Object -Property *, @{
						Name	      = 'GPODisplayName'
						Expression = {
							$_.LinkedGroupPolicyObjects | ForEach-Object {
								-join ([adsi]"LDAP://$_").displayName
							}
						}
					}
					
					if ($? -eq $true)
					{
						$ouGPODisplayNames = $ouGPONames.GPODisplayName -join "`n"
					}
					else
					{
						$ouGPODisplayNames = (Get-GPInheritance -Target $ouDN -Domain $domDNS -Server $domDNS).GpoLinks | `
						Foreach-Object { Get-GPO -Name ($_.DisplayName) -Domain $domDNS -Server $domDNS }
					}
					
				}
				catch
				{
					$errorMessage = "{0}: {1}" -f $Error[0], $Error[0].InvocationInfo.PositionMessage
					Write-Error $errorMessage -ErrorAction Continue
				}
				
			}
			else
			{
				$ouGPODisplayNames = "None"
			}
		}
		catch
		{
			$errorMessage = "{0}: {1}" -f $Error[0], $Error[0].InvocationInfo.PositionMessage
			Write-Error $errorMessage -ErrorAction Continue
		}
		
		#$gc = $pdcFSMO + ":3268"
		$permsHash = @{
			OUDistinguishedName = $OU.DistinguishedName
			DomainController = $pdcFSMO
			VerboseOutput = $true
			ErrorAction = 'Stop'
			}
			
		if (($PSBoundParameters.ContainsKey('Credential')) -and ($null -ne ($PSBoundParameters["Credential"])))
		{
			$permsHash.Add('AuthType', 'Negotiate')
			$permsHash.Add('Credential', $Credential)
		}
		
		$ouPerms = Get-ADOUPerms @permsHash
		$ouPerms = $ouPerms | Select-Object -Property IdentityReference, ActiveDirectoryRights, AccessControlType
		#$ouPerms = [string]($ouPerms -join "`n")
		
		$ouObject += New-Object -TypeName PSCustomObject -Property ([ordered] @{
			"Domain"	     = $domDNS
			"OU Name"	     = $ouDN
			"Parent OU"    = $ParentObject
			"Child OUs"    = $ChildOUs
			"Managed By"   = $ouMgr
			"Linked GPOs"  = $ouGPODisplayNames
			"Permissions"  = $ouPerms | Out-String
			"When Created" = $ouCreated
			"When Changed" = $ouLastModified
		})
		
		$null = $OU = $ouDN = $ChildOUs = $OUParent = $ouParentName = $ouChildNames = $ouGPODisplayNames = $ouPerms
		
		$colResults += $ouObject
	})
	
	$null = $OUs
	[System.GC]::GetTotalMemory('ForceFullCollection') | Out-Null
}
catch
{
	$errorMessage = "{0}: {1}" -f $Error[0], $Error[0].InvocationInfo.PositionMessage
	Write-Error $errorMessage -ErrorAction Continue
}
finally
{
	#Save output
	Test-PathExists -Path $rptFolder -PathType Folder
	
	Write-Verbose -Message "Exporting data tables to Excel spreadsheet tabs."
	$strDomain = $DomainName.ToString().ToUpper()
	$outputCSV = "{0}\{1}" -f $rptFolder, "$($strDomain)_OU_Structure_as_of_$(Get-ReportDate).csv"
	$outputFile = "{0}\{1}" -f $rptFolder, "$($strDomain)_OU_Structure_as_of_$(Get-ReportDate).xlsx"
	
	$ExcelParams = @{
		Path	        = $outputFile
		StartRow     = 2
		StartColumn  = 1
		AutoSize     = $true
		AutoFilter   = $true
		BoldTopRow   = $true
		FreezeTopRow = $true
	}
	
	$setParams = @{
		Wraptext		      = $true
		VerticalAlignment    = 'Bottom'
		HorizontalAlignment = 'Left'
	}
	
	$colResults | Export-Csv -Path $outputCSV -NoTypeInformation
	$Excel = $colResults | Export-Excel @ExcelParams -WorkSheetname "AD Organizational Units" -PassThru
	$Sheet = $Excel.Workbook.Worksheets["AD Organizational Units"]
	$totalRows = $Sheet.Dimension.Rows
	Set-ExcelRange -Range $Sheet.Cells["A2:Z$($totalRows)"] @setParams
	Export-Excel -ExcelPackage $Excel -WorksheetName "AD Organizational Units" -Title "$($strDomain) Active Directory OU Configuration" -TitleFillPattern Solid -TitleSize 18 -TitleBackgroundColor LightBlue
}

#endregion


# SIG # Begin signature block
# MIIxzgYJKoZIhvcNAQcCoIIxvzCCMbsCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCBL/wOODSYWQRB3
# A8n86+5qcgeLcGGMh4DwGTzghDzoLqCCLAMwggV/MIIDZ6ADAgECAhAYtcKEQ5AS
# l0GsCYozZaYQMA0GCSqGSIb3DQEBCwUAMFIxEzARBgoJkiaJk/IsZAEZFgNjb20x
# GDAWBgoJkiaJk/IsZAEZFghEZWxvaXR0ZTEhMB8GA1UEAxMYRGVsb2l0dGUgU0hB
# MiBMZXZlbCAxIENBMB4XDTE1MDkwMTE1MDcyNVoXDTM1MDkwMTE1MDcyNVowUjET
# MBEGCgmSJomT8ixkARkWA2NvbTEYMBYGCgmSJomT8ixkARkWCERlbG9pdHRlMSEw
# HwYDVQQDExhEZWxvaXR0ZSBTSEEyIExldmVsIDEgQ0EwggIiMA0GCSqGSIb3DQEB
# AQUAA4ICDwAwggIKAoICAQCU+o2qpWkTjV2nWzX6d4z5e/nN9QApOsPXREAX16kU
# WYggwfrZWAxc5hhYGvI1BpQBg9mW+/9O3RwIpxrlcBYqngNsF5uUKbF8eyoTPdH+
# TOf8IdEedDdgxlEyisBxyrzYN3EqLCfD2jRblIYPkD7M1eH0ONwLHQblE4BqqK/u
# bcdjPYesS+p0i4yQyiPtjaJ4yLD8+4iNVTzCah2W2QGYagB445xVhpYFlOkrLQ0L
# /Fgvt4d8wp2BFprfSkVV5mWI3wMyI397Ft+vqK7TR9ACw3GJmV4f/oIse44N2H7s
# orTRanZaL2rP8h44AaSPyPSNcWedfk8dBZxVW/wTywgrVL3nEPGaE6yswaQYDobc
# XtrWp8jVae29gFP3x5SBlDfCEOqPKmPrbaONcuRTGV+5R64EcHb5P2RJtqlAciGM
# ATd/3sD8U67MbRp6uvp40Ll2g2M5ffUF+nlbRaSFtfwPQB7F/u/46xuioWQCRo3N
# JLahpfv4bdcDDav6yUObKKav5uhosSji8z6gXMVR4FA5e+MYESca3FeZIg/Cwu74
# eWE6+O1cg0tuxzOXZBbMLZqtE/HHocuvWE+PR6LLPtsoM9D435T2Xk560vo00WCm
# lUOipJWjkVTCHjvBkpqpdnxvWM+P9JL4hGxsqcInfv5hd+QDYyJ33HygzJF8F/B5
# 3wIDAQABo1EwTzALBgNVHQ8EBAMCAYYwDwYDVR0TAQH/BAUwAwEB/zAdBgNVHQ4E
# FgQUvougK2ZZWfqQcVGlp6gGQk6QPO0wEAYJKwYBBAGCNxUBBAMCAQAwDQYJKoZI
# hvcNAQELBQADggIBAINO9HIsbvrfvlRKl7Ep/QEYK5h9b6j/obXwZGpl5iWzuU7b
# nXGfuF9dpx8RkgcI+iLNqfTK9km1hShxdEyJ6jvrpnCfyDg6ARmMlezDzWnQaWTK
# WZ1aGo76xG4bJi8ZjfYxZnXqjPrH8Ib6ux6i7Yewsu2VJXvoPZc+cO60gAcj9LiK
# zgDsXRX6g1fThoelRcxvRjUGQ8o0XWAiWZ/it21GHr7avfvhA/5G7D6cmI83AXtM
# XbzCTPWJRdYr0VKYzVDr7+5fGnH2pEVHC/6kJ3/Tsid5ncabHFM0nFLPMQnEh/Cg
# kifocv01g+W0BnsK5ZiyblMVj7HlNnmpL87r6cAd3fRPsr+r4fmGOn22KxhqHdGI
# E45T6Jm+1SyvVdotPvOLPGeaniywcevIuE2Ri7a6H91PX+KrxJszd8oYSUwv6YTt
# k8CDarYHvqSBtj4asS3M2fPA18jCSH8gMpgn2folTbJDihrmwmNN/m4dSHFX1l3z
# 8FkrwvcflLsjLfIJUvKr3zG5RncKEvgWgp5Lz24yuAoAM/bGcB0jDWed3rtlI+tN
# WTexk7+4lUpPOOlTRXXSxzbtWZXH1gHJzijjEOhDr54CVRs2m6gWlyHN0vZ3HklR
# qUuXZPiODhQjMpyJ0+ESbJTUCkpRcVBNvHQ3y/VYLZRdE7C5zRgMtHu/pEIuMIIF
# jTCCBHWgAwIBAgIQDpsYjvnQLefv21DiCEAYWjANBgkqhkiG9w0BAQwFADBlMQsw
# CQYDVQQGEwJVUzEVMBMGA1UEChMMRGlnaUNlcnQgSW5jMRkwFwYDVQQLExB3d3cu
# ZGlnaWNlcnQuY29tMSQwIgYDVQQDExtEaWdpQ2VydCBBc3N1cmVkIElEIFJvb3Qg
# Q0EwHhcNMjIwODAxMDAwMDAwWhcNMzExMTA5MjM1OTU5WjBiMQswCQYDVQQGEwJV
# UzEVMBMGA1UEChMMRGlnaUNlcnQgSW5jMRkwFwYDVQQLExB3d3cuZGlnaWNlcnQu
# Y29tMSEwHwYDVQQDExhEaWdpQ2VydCBUcnVzdGVkIFJvb3QgRzQwggIiMA0GCSqG
# SIb3DQEBAQUAA4ICDwAwggIKAoICAQC/5pBzaN675F1KPDAiMGkz7MKnJS7JIT3y
# ithZwuEppz1Yq3aaza57G4QNxDAf8xukOBbrVsaXbR2rsnnyyhHS5F/WBTxSD1If
# xp4VpX6+n6lXFllVcq9ok3DCsrp1mWpzMpTREEQQLt+C8weE5nQ7bXHiLQwb7iDV
# ySAdYyktzuxeTsiT+CFhmzTrBcZe7FsavOvJz82sNEBfsXpm7nfISKhmV1efVFiO
# DCu3T6cw2Vbuyntd463JT17lNecxy9qTXtyOj4DatpGYQJB5w3jHtrHEtWoYOAMQ
# jdjUN6QuBX2I9YI+EJFwq1WCQTLX2wRzKm6RAXwhTNS8rhsDdV14Ztk6MUSaM0C/
# CNdaSaTC5qmgZ92kJ7yhTzm1EVgX9yRcRo9k98FpiHaYdj1ZXUJ2h4mXaXpI8OCi
# EhtmmnTK3kse5w5jrubU75KSOp493ADkRSWJtppEGSt+wJS00mFt6zPZxd9LBADM
# fRyVw4/3IbKyEbe7f/LVjHAsQWCqsWMYRJUadmJ+9oCw++hkpjPRiQfhvbfmQ6QY
# uKZ3AeEPlAwhHbJUKSWJbOUOUlFHdL4mrLZBdd56rF+NP8m800ERElvlEFDrMcXK
# chYiCd98THU/Y+whX8QgUWtvsauGi0/C1kVfnSD8oR7FwI+isX4KJpn15GkvmB0t
# 9dmpsh3lGwIDAQABo4IBOjCCATYwDwYDVR0TAQH/BAUwAwEB/zAdBgNVHQ4EFgQU
# 7NfjgtJxXWRM3y5nP+e6mK4cD08wHwYDVR0jBBgwFoAUReuir/SSy4IxLVGLp6ch
# nfNtyA8wDgYDVR0PAQH/BAQDAgGGMHkGCCsGAQUFBwEBBG0wazAkBggrBgEFBQcw
# AYYYaHR0cDovL29jc3AuZGlnaWNlcnQuY29tMEMGCCsGAQUFBzAChjdodHRwOi8v
# Y2FjZXJ0cy5kaWdpY2VydC5jb20vRGlnaUNlcnRBc3N1cmVkSURSb290Q0EuY3J0
# MEUGA1UdHwQ+MDwwOqA4oDaGNGh0dHA6Ly9jcmwzLmRpZ2ljZXJ0LmNvbS9EaWdp
# Q2VydEFzc3VyZWRJRFJvb3RDQS5jcmwwEQYDVR0gBAowCDAGBgRVHSAAMA0GCSqG
# SIb3DQEBDAUAA4IBAQBwoL9DXFXnOF+go3QbPbYW1/e/Vwe9mqyhhyzshV6pGrsi
# +IcaaVQi7aSId229GhT0E0p6Ly23OO/0/4C5+KH38nLeJLxSA8hO0Cre+i1Wz/n0
# 96wwepqLsl7Uz9FDRJtDIeuWcqFItJnLnU+nBgMTdydE1Od/6Fmo8L8vC6bp8jQ8
# 7PcDx4eo0kxAGTVGamlUsLihVo7spNU96LHc/RzY9HdaXFSMb++hUD38dglohJ9v
# ytsgjTVgHAIDyyCwrFigDkBjxZgiwbJZ9VVrzyerbHbObyMt9H5xaiNrIv8SuFQt
# J37YOtnwtoeW/VvRXKwYw02fc7cBqZ9Xql4o4rmUMIIF3jCCA8agAwIBAgITPgAA
# AAp01W3Jvy6VAgACAAAACjANBgkqhkiG9w0BAQsFADBUMRMwEQYKCZImiZPyLGQB
# GRYDY29tMRgwFgYKCZImiZPyLGQBGRYIRGVsb2l0dGUxIzAhBgNVBAMTGkRlbG9p
# dHRlIFNIQTIgTGV2ZWwgMiBDQSAyMB4XDTIxMDYyOTE5MzUwMVoXDTI2MDYyOTE5
# NDUwMVowbDETMBEGCgmSJomT8ixkARkWA2NvbTEYMBYGCgmSJomT8ixkARkWCGRl
# bG9pdHRlMRYwFAYKCZImiZPyLGQBGRYGYXRyYW1lMSMwIQYDVQQDExpEZWxvaXR0
# ZSBTSEEyIExldmVsIDMgQ0EgMjCCASIwDQYJKoZIhvcNAQEBBQADggEPADCCAQoC
# ggEBAKMErt43dTBJZGDAXh2pNncSX8kiKriuGlr08U/3ZtI1k6YKEtUpns34twsN
# 6Rwmuq8Z2FTxlfFlCKHitdWkr6ES+gC/uh0MAPix1XZmErACC2j2rVDX1ELXzwtd
# zCIrzpBaXXxD+lCw0eou0CEnSQXAfYLEZ3+Eoj6HjejDLwuBAhTisC4mwEyIoTVU
# sgkZns4l3X0rXyvZfsxN7lGLV9wIDzP73qAl+AJ6W3vShFbNb7Gzzhln5qvho/y5
# 542rzi+SwcAtCLbmL+nrxSyNjc+p1w3qHV+ZmknT7Vtz30738mln8F9ne0ZvWo9M
# Ba9Mtu3H/FmFcyW/m9hlsYnV0pECAwEAAaOCAY8wggGLMBIGCSsGAQQBgjcVAQQF
# AgMCAAIwIwYJKwYBBAGCNxUCBBYEFLy+iTDVUCgkot0nGQ6PZed3WpwRMB0GA1Ud
# DgQWBBQ4oakuFXDiR2EUBm025muPDi7DYTAZBgkrBgEEAYI3FAIEDB4KAFMAdQBi
# AEMAQTALBgNVHQ8EBAMCAYYwEgYDVR0TAQH/BAgwBgEB/wIBADAfBgNVHSMEGDAW
# gBRHLjbutJz/XF4YfLgT4b6pIB4UszBcBgNVHR8EVTBTMFGgT6BNhktodHRwOi8v
# cGtpLmRlbG9pdHRlLmNvbS9DZXJ0RW5yb2xsL0RlbG9pdHRlJTIwU0hBMiUyMExl
# dmVsJTIwMiUyMENBJTIwMi5jcmwwdgYIKwYBBQUHAQEEajBoMGYGCCsGAQUFBzAC
# hlpodHRwOi8vcGtpLmRlbG9pdHRlLmNvbS9DZXJ0RW5yb2xsL1NIQTJMVkwyQ0Ey
# X0RlbG9pdHRlJTIwU0hBMiUyMExldmVsJTIwMiUyMENBJTIwMigyKS5jcnQwDQYJ
# KoZIhvcNAQELBQADggIBAGb01UEleTvuxAzRCWf39e2Ksfpsk8hfLVoWHKvXJ1M6
# 5ndOqA5ZjFlhd3muKeyRastbEQ16n1RV760y70Npp2L8Zmp/u0FmlvdzTtnWcc4m
# ny9FO0hFOHShoDy+ZGvKdsikSnod01D0dc5OCHGUEMse3xJvOobzXy02yVlo98Ec
# AuyPgWP21LbSOPAPU9OJPtNmbBSi9Tgcazl+204X+FpGrT+eBlh5p4sR5hSY2HYo
# ZYplGGvT5OABwS3U/eMXw3oSHgnMwtj6MmUJH/M/RZaeyxPsETZ9itakLVI1JnYb
# wJuR6DdlXQDgQ5KKulVHT3LDRbf/+GmJn56dGk2kWUQzsbqYVfWB6WY6JDndX3eL
# jeKE+7ukWmSE4rCHk0h9M9waCWaZjnuTqEqDOim91L/UoFJJ4KnpPDrGe5dZj6FX
# VdOVPZb8AO+ZP0QgjyxSwssAQpfUJbUFJI2Y91Qz7dUEyWyunHI1g/CPVUWWL1UV
# +VqwXq7C0d8RAdi21aFRaMS7leHS7zzPIEduMQgEIrvClZ0rRwuMjH2TeiC9t4Qb
# NEJFuLCFwLTC3sTi2p4tWd4nVUv5oQtYSqMUhub4p+yXIoX5je0Oqb5s+T5cdamI
# sw4WI4S11ZlosGZ/XryOU8MdrGIyIBoL/CGhDHWvmlRM7rOQ4aRfvw4+IOvBRTSE
# MIIGrjCCBJagAwIBAgIQBzY3tyRUfNhHrP0oZipeWzANBgkqhkiG9w0BAQsFADBi
# MQswCQYDVQQGEwJVUzEVMBMGA1UEChMMRGlnaUNlcnQgSW5jMRkwFwYDVQQLExB3
# d3cuZGlnaWNlcnQuY29tMSEwHwYDVQQDExhEaWdpQ2VydCBUcnVzdGVkIFJvb3Qg
# RzQwHhcNMjIwMzIzMDAwMDAwWhcNMzcwMzIyMjM1OTU5WjBjMQswCQYDVQQGEwJV
# UzEXMBUGA1UEChMORGlnaUNlcnQsIEluYy4xOzA5BgNVBAMTMkRpZ2lDZXJ0IFRy
# dXN0ZWQgRzQgUlNBNDA5NiBTSEEyNTYgVGltZVN0YW1waW5nIENBMIICIjANBgkq
# hkiG9w0BAQEFAAOCAg8AMIICCgKCAgEAxoY1BkmzwT1ySVFVxyUDxPKRN6mXUaHW
# 0oPRnkyibaCwzIP5WvYRoUQVQl+kiPNo+n3znIkLf50fng8zH1ATCyZzlm34V6gC
# ff1DtITaEfFzsbPuK4CEiiIY3+vaPcQXf6sZKz5C3GeO6lE98NZW1OcoLevTsbV1
# 5x8GZY2UKdPZ7Gnf2ZCHRgB720RBidx8ald68Dd5n12sy+iEZLRS8nZH92GDGd1f
# tFQLIWhuNyG7QKxfst5Kfc71ORJn7w6lY2zkpsUdzTYNXNXmG6jBZHRAp8ByxbpO
# H7G1WE15/tePc5OsLDnipUjW8LAxE6lXKZYnLvWHpo9OdhVVJnCYJn+gGkcgQ+ND
# Y4B7dW4nJZCYOjgRs/b2nuY7W+yB3iIU2YIqx5K/oN7jPqJz+ucfWmyU8lKVEStY
# dEAoq3NDzt9KoRxrOMUp88qqlnNCaJ+2RrOdOqPVA+C/8KI8ykLcGEh/FDTP0kyr
# 75s9/g64ZCr6dSgkQe1CvwWcZklSUPRR8zZJTYsg0ixXNXkrqPNFYLwjjVj33GHe
# k/45wPmyMKVM1+mYSlg+0wOI/rOP015LdhJRk8mMDDtbiiKowSYI+RQQEgN9XyO7
# ZONj4KbhPvbCdLI/Hgl27KtdRnXiYKNYCQEoAA6EVO7O6V3IXjASvUaetdN2udIO
# a5kM0jO0zbECAwEAAaOCAV0wggFZMBIGA1UdEwEB/wQIMAYBAf8CAQAwHQYDVR0O
# BBYEFLoW2W1NhS9zKXaaL3WMaiCPnshvMB8GA1UdIwQYMBaAFOzX44LScV1kTN8u
# Zz/nupiuHA9PMA4GA1UdDwEB/wQEAwIBhjATBgNVHSUEDDAKBggrBgEFBQcDCDB3
# BggrBgEFBQcBAQRrMGkwJAYIKwYBBQUHMAGGGGh0dHA6Ly9vY3NwLmRpZ2ljZXJ0
# LmNvbTBBBggrBgEFBQcwAoY1aHR0cDovL2NhY2VydHMuZGlnaWNlcnQuY29tL0Rp
# Z2lDZXJ0VHJ1c3RlZFJvb3RHNC5jcnQwQwYDVR0fBDwwOjA4oDagNIYyaHR0cDov
# L2NybDMuZGlnaWNlcnQuY29tL0RpZ2lDZXJ0VHJ1c3RlZFJvb3RHNC5jcmwwIAYD
# VR0gBBkwFzAIBgZngQwBBAIwCwYJYIZIAYb9bAcBMA0GCSqGSIb3DQEBCwUAA4IC
# AQB9WY7Ak7ZvmKlEIgF+ZtbYIULhsBguEE0TzzBTzr8Y+8dQXeJLKftwig2qKWn8
# acHPHQfpPmDI2AvlXFvXbYf6hCAlNDFnzbYSlm/EUExiHQwIgqgWvalWzxVzjQEi
# Jc6VaT9Hd/tydBTX/6tPiix6q4XNQ1/tYLaqT5Fmniye4Iqs5f2MvGQmh2ySvZ18
# 0HAKfO+ovHVPulr3qRCyXen/KFSJ8NWKcXZl2szwcqMj+sAngkSumScbqyQeJsG3
# 3irr9p6xeZmBo1aGqwpFyd/EjaDnmPv7pp1yr8THwcFqcdnGE4AJxLafzYeHJLtP
# o0m5d2aR8XKc6UsCUqc3fpNTrDsdCEkPlM05et3/JWOZJyw9P2un8WbDQc1PtkCb
# ISFA0LcTJM3cHXg65J6t5TRxktcma+Q4c6umAU+9Pzt4rUyt+8SVe+0KXzM5h0F4
# ejjpnOHdI/0dKNPH+ejxmF/7K9h+8kaddSweJywm228Vex4Ziza4k9Tm8heZWcpw
# 8De/mADfIBZPJ/tgZxahZrrdVcA6KYawmKAr7ZVBtzrVFZgxtGIJDwq9gdkT/r+k
# 0fNX2bwE+oLeMt8EifAAzV3C+dAjfwAL5HYCJtnwZXZCpimHCUcr5n8apIUP/JiW
# 9lVUKx+A+sDyDivl1vupL0QVSucTDh3bNzgaoSv27dZ8/DCCBrwwggSkoAMCAQIC
# EAuuZrxaun+Vh8b56QTjMwQwDQYJKoZIhvcNAQELBQAwYzELMAkGA1UEBhMCVVMx
# FzAVBgNVBAoTDkRpZ2lDZXJ0LCBJbmMuMTswOQYDVQQDEzJEaWdpQ2VydCBUcnVz
# dGVkIEc0IFJTQTQwOTYgU0hBMjU2IFRpbWVTdGFtcGluZyBDQTAeFw0yNDA5MjYw
# MDAwMDBaFw0zNTExMjUyMzU5NTlaMEIxCzAJBgNVBAYTAlVTMREwDwYDVQQKEwhE
# aWdpQ2VydDEgMB4GA1UEAxMXRGlnaUNlcnQgVGltZXN0YW1wIDIwMjQwggIiMA0G
# CSqGSIb3DQEBAQUAA4ICDwAwggIKAoICAQC+anOf9pUhq5Ywultt5lmjtej9kR8Y
# xIg7apnjpcH9CjAgQxK+CMR0Rne/i+utMeV5bUlYYSuuM4vQngvQepVHVzNLO9RD
# nEXvPghCaft0djvKKO+hDu6ObS7rJcXa/UKvNminKQPTv/1+kBPgHGlP28mgmoCw
# /xi6FG9+Un1h4eN6zh926SxMe6We2r1Z6VFZj75MU/HNmtsgtFjKfITLutLWUdAo
# Wle+jYZ49+wxGE1/UXjWfISDmHuI5e/6+NfQrxGFSKx+rDdNMsePW6FLrphfYtk/
# FLihp/feun0eV+pIF496OVh4R1TvjQYpAztJpVIfdNsEvxHofBf1BWkadc+Up0Th
# 8EifkEEWdX4rA/FE1Q0rqViTbLVZIqi6viEk3RIySho1XyHLIAOJfXG5PEppc3XY
# eBH7xa6VTZ3rOHNeiYnY+V4j1XbJ+Z9dI8ZhqcaDHOoj5KGg4YuiYx3eYm33aebs
# yF6eD9MF5IDbPgjvwmnAalNEeJPvIeoGJXaeBQjIK13SlnzODdLtuThALhGtycon
# cVuPI8AaiCaiJnfdzUcb3dWnqUnjXkRFwLtsVAxFvGqsxUA2Jq/WTjbnNjIUzIs3
# ITVC6VBKAOlb2u29Vwgfta8b2ypi6n2PzP0nVepsFk8nlcuWfyZLzBaZ0MucEdeB
# iXL+nUOGhCjl+QIDAQABo4IBizCCAYcwDgYDVR0PAQH/BAQDAgeAMAwGA1UdEwEB
# /wQCMAAwFgYDVR0lAQH/BAwwCgYIKwYBBQUHAwgwIAYDVR0gBBkwFzAIBgZngQwB
# BAIwCwYJYIZIAYb9bAcBMB8GA1UdIwQYMBaAFLoW2W1NhS9zKXaaL3WMaiCPnshv
# MB0GA1UdDgQWBBSfVywDdw4oFZBmpWNe7k+SH3agWzBaBgNVHR8EUzBRME+gTaBL
# hklodHRwOi8vY3JsMy5kaWdpY2VydC5jb20vRGlnaUNlcnRUcnVzdGVkRzRSU0E0
# MDk2U0hBMjU2VGltZVN0YW1waW5nQ0EuY3JsMIGQBggrBgEFBQcBAQSBgzCBgDAk
# BggrBgEFBQcwAYYYaHR0cDovL29jc3AuZGlnaWNlcnQuY29tMFgGCCsGAQUFBzAC
# hkxodHRwOi8vY2FjZXJ0cy5kaWdpY2VydC5jb20vRGlnaUNlcnRUcnVzdGVkRzRS
# U0E0MDk2U0hBMjU2VGltZVN0YW1waW5nQ0EuY3J0MA0GCSqGSIb3DQEBCwUAA4IC
# AQA9rR4fdplb4ziEEkfZQ5H2EdubTggd0ShPz9Pce4FLJl6reNKLkZd5Y/vEIqFW
# Kt4oKcKz7wZmXa5VgW9B76k9NJxUl4JlKwyjUkKhk3aYx7D8vi2mpU1tKlY71AYX
# B8wTLrQeh83pXnWwwsxc1Mt+FWqz57yFq6laICtKjPICYYf/qgxACHTvypGHrC8k
# 1TqCeHk6u4I/VBQC9VK7iSpU5wlWjNlHlFFv/M93748YTeoXU/fFa9hWJQkuzG2+
# B7+bMDvmgF8VlJt1qQcl7YFUMYgZU1WM6nyw23vT6QSgwX5Pq2m0xQ2V6FJHu8z4
# LXe/371k5QrN9FQBhLLISZi2yemW0P8ZZfx4zvSWzVXpAb9k4Hpvpi6bUe8iK6Wo
# nUSV6yPlMwerwJZP/Gtbu3CKldMnn+LmmRTkTXpFIEB06nXZrDwhCGED+8RsWQSI
# XZpuG4WLFQOhtloDRWGoCwwc6ZpPddOFkM2LlTbMcqFSzm4cd0boGhBq7vkqI1uH
# Rz6Fq1IX7TaRQuR+0BGOzISkcqwXu7nMpFu3mgrlgbAW+BzikRVQ3K2YHcGkiKjA
# 4gi4OA/kz1YCsdhIBHXqBzR0/Zd2QwQ/l4Gxftt/8wY3grcc/nS//TVkej9nmUYu
# 83BDtccHHXKibMs/yXHhDXNkoPIdynhVAku7aRZOwqw6pDCCBskwggSxoAMCAQIC
# EzQAAAAHiSF1iXPNJ/IAAAAAAAcwDQYJKoZIhvcNAQELBQAwUjETMBEGCgmSJomT
# 8ixkARkWA2NvbTEYMBYGCgmSJomT8ixkARkWCERlbG9pdHRlMSEwHwYDVQQDExhE
# ZWxvaXR0ZSBTSEEyIExldmVsIDEgQ0EwHhcNMjAwODA1MTczMjU2WhcNMzAwODA1
# MTc0MjU2WjBUMRMwEQYKCZImiZPyLGQBGRYDY29tMRgwFgYKCZImiZPyLGQBGRYI
# RGVsb2l0dGUxIzAhBgNVBAMTGkRlbG9pdHRlIFNIQTIgTGV2ZWwgMiBDQSAyMIIC
# IjANBgkqhkiG9w0BAQEFAAOCAg8AMIICCgKCAgEAmPb6sHLB25JD286NfyR2RfuN
# gmSXaR2dLojx7rPDqiEWKM01mSdquzeXj7Qu/VQsiQLV/9oxwMArSvjJHRjeQ2L7
# orPGytxWiO6nNHkKbPUCkBTmRALVcXK0iYmXhQjaypjx5y8bi3K13AR7axTbNlPE
# Fy3z9TwFGftmeJOIvle3dBvOCxJre1mxmf544tkzq+Df0ENP8sA41WeQbA5ZyDa2
# C8PWm8XL59X00UgtMJcOq4fCG+xkjl7nnbQ4/AP7lGHGkl0bnYE5Xd/nVA86+wO+
# uTUcmbs0fJ9fKO3bq3wgiUaRyyBbUQ2NzGlgaffxqge2lM3WCmiQeHKyfKsOkfg4
# 1+6h7qUFywDoDkvnVBjJs2+tgImqqD6iwmgZWHt6PeIiwJA/IIKBf0t1O16G39ui
# m6NSiesSK+wfOMxyxZio/BzKGPOtv4PwosBlPKlhK5bbvMWY2RFsWQJ6LPiRXlE5
# NIYbh/CTyngIdM6Drwr57sIZGWbKCJc9nORteVgx3pgciFAxOFGn1k3zmxM83qYx
# xgKi6fql8KCgbo+l6luROLa5rsRfkGPtRXy1HWJ7xwcf8/JxLJGlp1rtnGnZljvb
# 0Tbtwo8GwDoihSMSh9MoGrJTrtk8tnYf4UpLgGKjGyGOUBFGrRGQcEhWbzDTK5qZ
# P/0f31d3CndzQORYAb8CAwEAAaOCAZQwggGQMBAGCSsGAQQBgjcVAQQDAgECMCMG
# CSsGAQQBgjcVAgQWBBQV4b/ii/DtWsyFdE+p2v+xjwi3MzAdBgNVHQ4EFgQURy42
# 7rSc/1xeGHy4E+G+qSAeFLMwEQYDVR0gBAowCDAGBgRVHSAAMBkGCSsGAQQBgjcU
# AgQMHgoAUwB1AGIAQwBBMAsGA1UdDwQEAwIBhjASBgNVHRMBAf8ECDAGAQH/AgEB
# MB8GA1UdIwQYMBaAFL6LoCtmWVn6kHFRpaeoBkJOkDztMFgGA1UdHwRRME8wTaBL
# oEmGR2h0dHA6Ly9wa2kuZGVsb2l0dGUuY29tL0NlcnRFbnJvbGwvRGVsb2l0dGUl
# MjBTSEEyJTIwTGV2ZWwlMjAxJTIwQ0EuY3JsMG4GCCsGAQUFBwEBBGIwYDBeBggr
# BgEFBQcwAoZSaHR0cDovL3BraS5kZWxvaXR0ZS5jb20vQ2VydEVucm9sbC9TSEEy
# TFZMMUNBX0RlbG9pdHRlJTIwU0hBMiUyMExldmVsJTIwMSUyMENBLmNydDANBgkq
# hkiG9w0BAQsFAAOCAgEAh56DUZ5xeRvT/JdAM8biqLOI4PLlIhIPufRfMPJlmo64
# dY9G/JVkF7qh8SHWm7umjSTvb357kayJCks5Y3VwS11A9HMsRK11083exB27HUBd
# 2W3IyRv2KBZT+SsAsnhtb2slEuPqqrpFZC3u2RZa8XonKVVcX3wfFN0qxE+yXkjY
# MUNxr3kYuclb2kt/4/RggkfV06dL0X2lHktLMYILmr8Tb2/eU2S7//hrdcH/tcWZ
# 29hiIzL0qayp0j2MBuXACV/ZDNheEBvD659p14ae23CrgTpXSLL68RwHjaQqFVf2
# EWPXjR1MVJSvjB3QKiGdXTltUu1MBsrRHbFwj83xhiS1nTSWfIxSM+NG0u+tj9SJ
# 5fOQSEMlCe0achdoXPvF50uDwaTLxUOoBoDK2DKd8nJFa/x8/Gj35jn7RNp//Uuz
# bmIhOr7YZqdfiBwnGffm4rS577EnBSsQhuOjzujrJbJd3NP2ar293Zupr8d+QYUb
# U51ny+mUYbGQ8VQeZgo72XkFAnzx1vZw9UK5VU7pC0zlBZL/FNV6hbgcnxQ0K/qR
# AudJtx03GpNF5sqhyEC0ndvCSdljKsf4mvgNwrEDTa6HLtEKONisnLSg56IrcPx5
# W/eD8Ksodlwpfg8UM/A942V8JRZJLgFrc+nqysPi+cINMTd/n40h2wzGRlcjDE8w
# ggbKMIIFsqADAgECAhNlAJU42fb9El3zjUwgAAIAlTjZMA0GCSqGSIb3DQEBCwUA
# MGwxEzARBgoJkiaJk/IsZAEZFgNjb20xGDAWBgoJkiaJk/IsZAEZFghkZWxvaXR0
# ZTEWMBQGCgmSJomT8ixkARkWBmF0cmFtZTEjMCEGA1UEAxMaRGVsb2l0dGUgU0hB
# MiBMZXZlbCAzIENBIDIwHhcNMjMwNDI2MTgwNDUwWhcNMjUwNDI1MTgwNDUwWjBu
# MQswCQYDVQQGEwJVUzELMAkGA1UECBMCUEExEzARBgNVBAcTCkdsZW4gTWlsbHMx
# ETAPBgNVBAoTCERlbG9pdHRlMQwwCgYDVQQLEwNEVFMxHDAaBgNVBAMTE0lBTUlu
# ZnJhQ29kZVNpZ25pbmcwggEiMA0GCSqGSIb3DQEBAQUAA4IBDwAwggEKAoIBAQCX
# +Dk7TU4HXEhaUyGdGsgMjg5yed6BbkA+YP+pydr3VA1gf78cO4wSTadA27HzEfnQ
# fAGI7BADObnRYgIRK9CCHMWyzkBnv7qEXKli3dbLd+30ZiJBxPdCwvN49CCsltse
# sxyWZUH0gHdHG05y2BGzwPHaOqaQuLHnpmTDXcCcHhgfIaIJfX0DUondq484hh2F
# m6Ne0x3kcISftDK4mfZ8VaZWMrZ6dE5iDWlZXBeGZgLgs4CpBetTEkqc8odMm8nx
# cwCQk8eGMmm2QvYlhcqNuv039yr2mb4iN5Cs9GQ4EAKEH3HgoPAarflMrcKSwKL9
# vHgAbwZtn/FsxZgEN06zAgMBAAGjggNhMIIDXTAdBgNVHQ4EFgQU09Ll9rti/7M/
# YbA4pLtq/002ccIwHwYDVR0jBBgwFoAUOKGpLhVw4kdhFAZtNuZrjw4uw2EwggFB
# BgNVHR8EggE4MIIBNDCCATCgggEsoIIBKIaB1WxkYXA6Ly8vQ049RGVsb2l0dGUl
# MjBTSEEyJTIwTGV2ZWwlMjAzJTIwQ0ElMjAyKDIpLENOPXVzYXRyYW1lZW0wMDQs
# Q049Q0RQLENOPVB1YmxpYyUyMEtleSUyMFNlcnZpY2VzLENOPVNlcnZpY2VzLENO
# PUNvbmZpZ3VyYXRpb24sREM9ZGVsb2l0dGUsREM9Y29tP2NlcnRpZmljYXRlUmV2
# b2NhdGlvbkxpc3Q/YmFzZT9vYmplY3RDbGFzcz1jUkxEaXN0cmlidXRpb25Qb2lu
# dIZOaHR0cDovL3BraS5kZWxvaXR0ZS5jb20vQ2VydGVucm9sbC9EZWxvaXR0ZSUy
# MFNIQTIlMjBMZXZlbCUyMDMlMjBDQSUyMDIoMikuY3JsMIIBVwYIKwYBBQUHAQEE
# ggFJMIIBRTCBxAYIKwYBBQUHMAKGgbdsZGFwOi8vL0NOPURlbG9pdHRlJTIwU0hB
# MiUyMExldmVsJTIwMyUyMENBJTIwMixDTj1BSUEsQ049UHVibGljJTIwS2V5JTIw
# U2VydmljZXMsQ049U2VydmljZXMsQ049Q29uZmlndXJhdGlvbixEQz1kZWxvaXR0
# ZSxEQz1jb20/Y0FDZXJ0aWZpY2F0ZT9iYXNlP29iamVjdENsYXNzPWNlcnRpZmlj
# YXRpb25BdXRob3JpdHkwfAYIKwYBBQUHMAKGcGh0dHA6Ly9wa2kuZGVsb2l0dGUu
# Y29tL0NlcnRlbnJvbGwvdXNhdHJhbWVlbTAwNC5hdHJhbWUuZGVsb2l0dGUuY29t
# X0RlbG9pdHRlJTIwU0hBMiUyMExldmVsJTIwMyUyMENBJTIwMigyKS5jcnQwCwYD
# VR0PBAQDAgeAMDwGCSsGAQQBgjcVBwQvMC0GJSsGAQQBgjcVCIGBvUmFvoUTgtWb
# PIPXjgeG8ckKXIPK9y3C8zICAWQCAR4wEwYDVR0lBAwwCgYIKwYBBQUHAwMwGwYJ
# KwYBBAGCNxUKBA4wDDAKBggrBgEFBQcDAzANBgkqhkiG9w0BAQsFAAOCAQEAX94W
# DLBVQFBHMypTzJWGnOMCXwzvt6041xwGivYARE0aaJ4FaVe0DYcsNiFgybImyVuC
# Z6y6vVnPX7bLmrcU9k8PPSGSGUbAQ9/gPjEspmR1nQBbPf9gE/YDIsJONgW+6quY
# 8qhKl+PBawBrlbbRa4v2JTmwNeC/cnrHPWVqg7Mk+gEBlL0k4HhqMqsXuUxomKZO
# 04/3oJkMBQKXyBkab3JUCT7upI1iJ1g3c7hrXjt7dKcTm+zYijpoGMOZURYmlBJ4
# Xm/TC/xyMH4pH0mmjn8UUu5bp2Duuxl9VwLE+rD2WqSycGtuL00tCvX4dmlZkZtp
# CUCrRGiyUrzxRp93wDGCBSEwggUdAgEBMIGDMGwxEzARBgoJkiaJk/IsZAEZFgNj
# b20xGDAWBgoJkiaJk/IsZAEZFghkZWxvaXR0ZTEWMBQGCgmSJomT8ixkARkWBmF0
# cmFtZTEjMCEGA1UEAxMaRGVsb2l0dGUgU0hBMiBMZXZlbCAzIENBIDICE2UAlTjZ
# 9v0SXfONTCAAAgCVONkwDQYJYIZIAWUDBAIBBQCgTDAZBgkqhkiG9w0BCQMxDAYK
# KwYBBAGCNwIBBDAvBgkqhkiG9w0BCQQxIgQgonTAk29DLms9en1MNnpdPVR/HwmE
# 5+R1o4gIeCl2cl8wDQYJKoZIhvcNAQEBBQAEggEATJsSoquL7gVmTEcVyl2IMnCt
# oS38cX8sGSnE2CZs5TIDG6HDs8ch4UqS9lNeWQT80MBgQMizdqmIrvI/fLXYfkRQ
# 0WoHz6lFncNIPERQAeKZ7JHjydbh+/XaMVMC9CQzEWb10flzHED7+VabFey3sYFS
# NQdSqTr0wKuBzur1bzST8JiO4ktee4BCyzZvosuVjibC/f6/gt82WcFMdgmnube4
# aGzPLzE8QsA0T+xmyKKljP8iiJIf2USuPnSAnAHYy0aKogFBZXG49PQnG3gWvpLB
# tbghOu/XjwD00/AN6xKcFnSk5bTQoiNG3EJRaCTD52GAKoVTVEXwltslsJsWJqGC
# AyAwggMcBgkqhkiG9w0BCQYxggMNMIIDCQIBATB3MGMxCzAJBgNVBAYTAlVTMRcw
# FQYDVQQKEw5EaWdpQ2VydCwgSW5jLjE7MDkGA1UEAxMyRGlnaUNlcnQgVHJ1c3Rl
# ZCBHNCBSU0E0MDk2IFNIQTI1NiBUaW1lU3RhbXBpbmcgQ0ECEAuuZrxaun+Vh8b5
# 6QTjMwQwDQYJYIZIAWUDBAIBBQCgaTAYBgkqhkiG9w0BCQMxCwYJKoZIhvcNAQcB
# MBwGCSqGSIb3DQEJBTEPFw0yNDEwMTAxNTQ3MzNaMC8GCSqGSIb3DQEJBDEiBCDn
# Um+n4A06NcmdS08SE5f0/lWlfMUV6dOH+XNo/YRhCTANBgkqhkiG9w0BAQEFAASC
# AgBulg5csa8LF5JICiM/COJ6ycjU+ONadenX9IPSXlbcuYbxwQmmRUce0O9Y9cdq
# 6h2acuEMuYbuzMJGac1sVyhrPLJ6bSeFwOYa8qTa641M26MiO8K9SKstM4jqACMQ
# t04u4xz0Om2+6En0saI+OmmfEE201uvx2OYfC+j2xmJ2zhE7P9Ba2gztdR5uDtag
# UV8HZCAOQoNXcfiCsrHMIgwbGKGMaHc45FzqqBeiDwJlo5PgfUfLi5Krd3wgjgXL
# QSp3w/z2wMYneTh5vfZmGh+ZkQ8BgIbM0UUSZEJ+w4c3VqcAM7CrpAe3kehJdgLa
# W6kg8hco9ZwQc8X6w/hcIqKWRYHfAWLBVNonbg2w4lvAd7rKL5K6D3wlYMikvjGU
# 5yyma1IDowO7yK5I1eKCin0WgsC/ppYTglKh+Qk4KkP8KRK2wqVCVcy8bCe+154g
# 71uRlWDtBBho63csxkrDmOqtg3fM3eXw/ZmpeioiKBIou+SutC1PhtQlyrltNdbw
# e7un95/I4UlssZLoZ68fPdRS0mbIXa7/dLZOJbTN/MKRnaR+D38wc1Y8gLFz6rQq
# JxxIPkph+8KT8eyQuX3OdYSgH8YTzI1czj6HCT0HzmyDXdKD0MOM/sSBUYF8JUn1
# hJhr/sWDvl7ESXj9oh6o7/515PtVtzW8PpKr3WpzfH/vJA==
# SIG # End signature block
