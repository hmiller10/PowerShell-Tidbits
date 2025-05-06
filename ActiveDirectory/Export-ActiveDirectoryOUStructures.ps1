#Requires -Module ActiveDirectory, ImportExcel, GroupPolicy
#Requires -Version 5
#Requires -RunAsAdministrator
<#
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
###########################################################################
[CmdletBinding()]
param
(
[Parameter(Mandatory = $true)]
[ValidateNotNullOrEmpty()]
[string]$DomainName,
[Parameter(Mandatory = $false,
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
				if ($PSBoundParameters.ContainsValue('Excel'))
				{
					$ChildOUs = [String]($ouChildNames -join "`n")
				}
				elseif ($PSBoundParameters.ContainsValue('CSV'))
				{
					$ChildOUs = [String]($ouChildNames -join ";")
				}
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
							if ($PSBoundParameters.ContainsValue('Excel'))
							{
								$ouGPODisplayNames = $ouGPONames.GPODisplayName -join "`n"
							}
							elseif ($PSBoundParameters.ContainsValue('CSV'))
							{
								$ouGPODisplayNames = $ouGPONames.GPODisplayName -join ";"
							}
							
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
			
			$permsHash = @{
				OUDistinguishedName = $OU.DistinguishedName
				DomainController    = $pdcFSMO
				VerboseOutput	     = $true
				ErrorAction	     = 'Stop'
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
					"Domain"	    = $domDNS
					"OU Name"     = $ouDN
					"Parent OU"   = $ParentObject
					"Child OUs"   = $ChildOUs
					"Managed By"  = $ouMgr
					"Linked GPOs" = $ouGPODisplayNames
					"Permissions" = $ouPerms | Out-String
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
	$outputFile = "{0}\{1}" -f $rptFolder, "$($strDomain)_OU_Structure_as_of_$(Get-ReportDate).csv"
	$xlOutput = $OutputFile.ToString().Replace([System.IO.Path]::GetExtension($OutputFile), ".xlsx")
	$colResults | Export-Csv -Path $outputFile -NoTypeInformation
	$wsName = "AD Organizational Units"
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
		FontColor         = 'White'
		FontSize	        = 16
		Bold		        = $true
		BackgroundColor   = 'Black'
		BackgroundPattern = 'Solid'
	}
	
	$xl = $colResults | Export-Excel @xlParams
	$Sheet = $xl.Workbook.Worksheets[$wsName]
	$lastRow = $Sheet.Dimension.End.Row
	
	Set-ExcelRange -Range $Sheet.Cells["A1"] -Value "$($strDomain) Active Directory OU Configuration" @titleParams
	Set-ExcelRange -Range $Sheet.Cells["A2"] @headerParams1
	Set-ExcelRange -Range $Sheet.Cells["B2:Z2"] @headerParams2
	Set-ExcelRange -Range $Sheet.Cells["A3:I$($lastRow)"] @setParams
	
	Export-Excel -ExcelPackage $xl -WorksheetName $wsName -FreezePane 3, 0
}

#endregion

