#Requires -Module ActiveDirectory, ImportExcel
#Requires -Version 7
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

Try
{
	Import-Module ImportExcel -Force
}
Catch
{
	Try
	{
		$module = Get-Module -Name ImportExcel;
		 $modulePath = Split-Path $module.Path;
		 $psdPath = "{0}\{1}" -f $modulePath, "ImportExcel.psd1"
		Import-Module $psdPath -ErrorAction Stop
	}
	Catch
	{
		Throw "ImportExcel PS module could not be loaded. $($_.Exception.Message)"
	}
}

#EndRegion

#region Variables
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

$domainParams = @{
	Identity = $DomainName
	Server = $DomainName
	ErrorAction = 'Stop'
}

if ($PSBoundParameters.ContainsKey('Credential') -and ($PSBoundParameters["Credential"]))
{
	$domainParams.Add('Credential', $Credential)
}

$Domain = Get-ADDomain @domainParams | Select-Object -Property distinguishedName, DnsRoot, Name, pdcEmulator
$dnsRoot = $Domain.dnsRoot

[int32]$throttleLimit = 100
$rptFolder = 'E:\Reports'

[PSObject[]]$global:ouObject = @()
$colResults = @()
#endregion

#Region Functions

function Test-PathExists
{
<#
.SYNOPSIS
Checks if a path to a file or folder exists, and creates it if it does not exist.

.DESCRIPTION
Checks if a path to a file or folder exists, and creates it if it does not exist.

.PARAMETER Path
Full path to the file or folder to be checked

.PARAMETER PathType
Valid options are "File" and "Folder", depending on which to check.

.OUTPUTS
None

.EXAMPLE
Test-PathExists -Path "C:\temp\SomeFile.txt" -PathType File
	
.EXAMPLE
Test-PathExists -Path "C:\temp" -PathFype Folder

#>
	[CmdletBinding()]
	param
	(
		[Parameter(Mandatory = $true,
				 Position = 0)]
		[String]$Path,
		[Parameter(Mandatory = $true,
				 Position = 1)]
		[Object]$PathType
	)
	
	Begin { $VerbosePreference = 'Continue' }
	
	Process
	{
		Switch ($PathType)
		{
			File
			{
				If ((Test-Path -Path $Path -PathType Leaf) -eq $true)
				{
					Write-Information -MessageData "File: $Path already exists..."
				}
				Else
				{
					New-Item -Path $Path -ItemType File -Force
					Write-Verbose -Message "File: $Path not present, creating new file..."
				}
			}
			Folder
			{
				If ((Test-Path -Path $Path -PathType Container) -eq $true)
				{
					Write-Information -MessageData "Folder: $Path already exists..."
				}
				Else
				{
					New-Item -Path $Path -ItemType Directory -Force
					Write-Verbose -Message "Folder: $Path not present, creating new folder"
				}
			}
		}
	}
	
	End { }
	
}#end function Test-PathExists

function Get-ReportDate
{
<#
	.SYNOPSIS
		function to get date in format yyyy-MM-dd
	
	.DESCRIPTION
		function to get date using the Get-Date cmdlet in the format yyyy-MM-dd
	
	.EXAMPLE
		PS C:\> $rptDate = Get-ReportDate
	
	.NOTES
		THIS CODE IS MADE AVAILABLE AS IS, WITHOUT WARRANTY OF ANY KIND. THE ENTIRE RISK OF 
		THE USE OR THE RESULTS FROM THE USE OF THIS CODE REMAINS WITH THE USER.
#>
	
	#Begin function get report execution date
	Get-Date -Format "yyyy-MM-dd"
} #End function Get-ReportDate



#EndRegion




#Region Script
$Error.Clear()


#List properties to be collected into array for writing to OU tab
$OUs = @()
$ouProps = @("distinguishedName", "gpLink", "LinkedGroupPolicyObjects", "ManagedBy", "Name", "objectCategory", "objectClass", "whenCreated", "whenChanged")
	
Write-Verbose -Message ("Gathering collection of AD Organizational Units for {0}" -f $Domain.Name)
try
{
	$OUs = Get-ADOrganizationalUnit -Filter * -Properties $ouProps -SearchBase $Domain.distinguishedName -SearchScope Subtree -ResultSetSize $null -Server $Domain.pdcEmulator | Select-Object -Property $ouProps
	if (!($?))
	{
		try
		{
			$OUs = Get-ADOrganizationalUnit -Filter * -Properties $ouProps -SearchBase $Domain.distinguishedName -SearchScope Subtree -ResultSetSize $null -Server $Domain.DnsRoot | Select-Object -Property $ouProps
		}
		catch
		{
			Write-Warning ("Error occurred getting list of AD OUs for {0}" -f $Domain.Name)
			$errorMessage = "{0}: {1}" -f $Error[0], $Error[0].InvocationInfo.PositionMessage
			Write-Error $errorMessage -ErrorAction Continue
		}
	}
	
	$colResults = $OUs | ForEach-Object -Parallel {
		
		function Get-GPODisplayNames
		{
		<#
			.SYNOPSIS
				Return GPO display names of all GPOs
			
			.DESCRIPTION
				This functions will return the displayName of a GPO using ADSI
			
			.PARAMETER GPOs
				Collection of GPOs to get display names for
			
			.EXAMPLE
				PS C:\> Get-GPODisplayNames -GPOs <GPOList>
			
			.NOTES
				THIS CODE IS MADE AVAILABLE AS IS, WITHOUT WARRANTY OF ANY KIND. THE ENTIRE RISK OF 
				THE USE OR THE RESULTS FROM THE USE OF THIS CODE REMAINS WITH THE USER.
		#>
			
			[CmdletBinding()]
			param
			(
				[Parameter(Mandatory = $true,
						 ValueFromPipeline = $true,
						 Position = 0)]
				[ValidateNotNullOrEmpty()]
				$GPOs
			)
			
			begin
			{
				#Begin function to get DisplayName attribute of GPO object
				$colGPOs = @()
			}
			process
			{
				$GPOs.foreach({
						
						try
						{
							#$currentGPO = [ADSI]"LDAP://$_" | Select-Object -Property DisplayName
							$currentGPO = [ADSI]"LDAP://$_"
							if ($null -ne $currentGPO.DisplayName)
							{
								$colGPOs += $currentGPO
							}
							
							if ($colGPOs.Count -ge 1)
							{
								$objDisplayNames = $colGPOs.DisplayName
							}
							else
							{
								$objDisplayNames = "GPO display name for the current GPO is not available."
							}
						}
						catch
						{
							$errorMessage = "{0}: {1}" -f $Error[0], $Error[0].InvocationInfo.PositionMessage
							Write-Error $errorMessage -ErrorAction Continue
						}
						
					})
			}
			end
			{
				return $objDisplayNames
			}
			
		} #End function Get-GPODisplayNames
		
		$OU = $_
		$ouGPOs = @()
		$ouChildNames = @()
		
		$ouDN = ($OU).distinguishedName
		$ouCreated = ($OU).whenCreated
		$ouLastModified = ($OU).whenChanged
		
		try
		{
			Write-Verbose -Message ("Working on Organizational Unit: {0}" -f $ouDN)
			$ouParent = [ADSI]"LDAP://$ouDN"
			$ouParentName = ($ouParent).Parent -replace "LDAP://", ""
		}
		catch
		{
			Write-Warning ("Error occurred geting parent OU information for: {0}" -f $ouDN)
			$errorMessage = "{0}: {1}" -f $Error[0], $Error[0].InvocationInfo.PositionMessage
			Write-Error $errorMessage -ErrorAction Continue
		}
		
		try
		{
			Write-Verbose -Message ("Examining Sub-OUs of: {0}" -f $ouDN)
			[Array]$ouChildNames = ($ouParent).psBase.Children | Where-Object { $_.psBase.schemaClassName -eq "OrganizationalUnit" } | Select-Object -ExpandProperty distinguishedName
			if (($ouChildNames).Count -ge 1) { $ChildOUs = [String]($ouChildNames -join "`n") }
			else { $ChildOUs = "None" }
		}
		catch
		{
			Write-Warning ("Error occurred get list of child OUs for: {0}" -f $ouDN)
			$errorMessage = "{0}: {1}" -f $Error[0], $Error[0].InvocationInfo.PositionMessage
			Write-Error $errorMessage -ErrorAction Continue
		}
		
		
		
		if ($null -ne $OU.ManagedBy)
		{
			$ouMgr = ($OU).ManagedBy
		}
		else
		{
			$ouMgr = "None listed for this OU."
		}
		
		Write-Verbose -Message "Gathering list of group policies linked to $($ouDN)."
		try
		{
			$ouGPOs = $OU | Select-Object -ExpandProperty LinkedGroupPolicyObjects
			if ($ouGPOs.Count -ge 1)
			{
				$ouGPODisplayNames = Get-GPODisplayNames -GPOs $ouGPOs
				$ouGPODisplayNames = $ouGPODisplayNames -join "`n"
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
		
		$ouObject += New-Object -TypeName PSCustomObject -Property ([ordered] @{
			"Domain" = $using:dnsRoot
			"OU Name" = $ouDN
			"Parent OU" = $ouParentName
			"Child OUs" = $ChildOUs
			"Managed By" =  $ouMgr
			"Linked GPOs" = $ouGPODisplayNames
			"When Created" = $ouCreated
			"When Changed" = $ouLastModified
			})
		
		$ouDN = $ChildOUs = $OUParent = $ouParentName = $ouChildNames = $ouGPODisplayNames = $null
		
		Write-Output $ouObject
	} -ThrottleLimit $throttleLimit
	
	$null = $OUs
	[System.GC]::GetTotalMemory('forcefullcollection') | Out-Null
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
	
	$colResults | Export-Csv -Path $outputCSV -NoTypeInformation
	$Excel = $colResults | Export-Excel @ExcelParams -WorkSheetname "AD Organizational Units" -PassThru
	$Sheet = $Excel.Workbook.Worksheets["AD Organizational Units"]
	$totalRows = $Sheet.Dimension.Rows
	Set-Format -Address $Sheet.Cells["A2:Z$($totalRows)"] -Wraptext -VerticalAlignment Center -HorizontalAlignment Center
	Export-Excel -ExcelPackage $Excel -WorksheetName "AD Organizational Units" -Title "$($strDomain) Active Directory OU Configuration"  -TitleFillPattern Solid -TitleSize 18 -TitleBackgroundColor LightBlue
}

#endregion