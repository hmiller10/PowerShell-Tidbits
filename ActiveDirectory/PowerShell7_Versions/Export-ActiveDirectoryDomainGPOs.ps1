#Requires -Module ActiveDirectory, ImportExcel, GroupPolicy
#Requires -Version 7
#Requires -RunAsAdministrator
<#
	.NOTES
	THIS CODE IS MADE AVAILABLE AS IS, WITHOUT WARRANTY OF ANY KIND. THE
	RISK OF THE USE OR THE RESULTS FROM THE USE OF THIS CODE REMAINS WITH
	THE USER.

	.SYNOPSIS
		Create GPO report

	.DESCRIPTION
		This script leverages the parallel processing functionality in PowerShell 7
		to process and report on the group policy configuration of the domain named 
		piped to the script parameter
		
	.PARAMETER DomainName
		Fully qualified domain name of domain where GPO report should be created from
		
	.PARAMETER Credential
		PSCredential
	
	.OUTPUTS
	Excel spreadsheet with group policy configuration for named AD domain
	
	.EXAMPLE
	PS C:> .\Export-ActiveDirectoryDomainGPOs.ps1 -DomainName my.domain.com
	
	.EXAMPLE
	PS C:> .\Export-ActiveDirectoryDomainGPOs.ps1 -DomainName my.domain.com -Credential PSCredential

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

Try 
{
	Import-Module GroupPolicy -ErrorAction Stop
}
Catch
{
	Throw "Group Policy module could not be loaded. $($_.Exception.Message)"
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
$pdcE = $Domain.pdcEmulator
$dnsRoot = $Domain.dnsRoot

[int32]$throttleLimit = 100
$rptFolder = 'E:\Reports'

[PSObject[]]$global:gpoObject = @()
$colResults = @()
#endregion

#Region Functions

function Add-DataTable
{
<#
	.SYNOPSIS
		Creates PS data table with assigned name and column data
	
	.DESCRIPTION
		This function creates a [System.Data.DataTable] to store script output for reporting.
	
	.PARAMETER TableName
		A brief description to reference the data table by
	
	.PARAMETER ColumnArray
		List of column headers including ColumnName and DataType
	
	.EXAMPLE
		PS C:\> Add-DataTable -TableName <TableName> -ColumnArray <DataColumnDefinitions>
	
	.NOTES
		THIS CODE IS MADE AVAILABLE AS IS, WITHOUT WARRANTY OF ANY KIND. THE ENTIRE RISK OF 
		THE USE OR THE RESULTS FROM THE USE OF THIS CODE REMAINS WITH THE USER.
#>
	
	[CmdletBinding()]
	[OutputType([System.Data.DataTable])]
	param
	(
		[Parameter(Mandatory = $true,
				 Position = 0)]
		[ValidateNotNullOrEmpty()]
		[String]$TableName,  #'TableName'
		[Parameter(Mandatory = $true,
				 Position = 1)]
		[ValidateNotNullOrEmpty()]
		$ColumnArray  #'DataColumnDefinitions'
	)
	
	
	Begin
	{
		$dt = $null
		$dt = New-Object System.Data.DataTable("$TableName")
	}
	Process
	{
		ForEach ($col in $ColumnArray)
		{
			[void]$dt.Columns.Add([System.Data.DataColumn]$col.ColumnName.ToString(), $col.DataType)
		}
	}
	End
	{
		Write-Output @(,$dt)
	}
} #end function Add-DataTable

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

$GPOs = @()

try
{
	$GPOs = Get-GPO -Domain $DomainName -Server $pdcE -All
	if (!($?))
	{
		try
		{
			$GPOs = Get-GPO -Domain $domDNS -Server $dnsRoot -All
		}
		catch
		{
			$errorMessage = "{0}: {1}" -f $Error[0], $Error[0].InvocationInfo.PositionMessage
			Write-Error $errorMessage -ErrorAction Continue
		}
	}
	
	$colResults = $GPOs | ForEach-Object -Parallel {
		
		function Get-GPOInfo
		{
		<#
			.SYNOPSIS
				Function to return GPO properties
			
			.DESCRIPTION
				This function returns the properties of a group policy object passed into the function as a parameter
				and uses it to populate specific data into an array
			
			.PARAMETER DomainFQDN
				Fully qualified domain name EG: my.domain.com
			
			.PARAMETER gpoGUID
				Group policy GUID value
			
			.EXAMPLE
				PS C:\> Get-GPOInfo -DomainFQDN $value1 -gpoGUID $value2
			
			.NOTES
				THIS CODE IS MADE AVAILABLE AS IS, WITHOUT WARRANTY OF ANY KIND. THE ENTIRE RISK OF 
				THE USE OR THE RESULTS FROM THE USE OF THIS CODE REMAINS WITH THE USER.
		#>
			
			#Begin function to get GPO properties
			[CmdletBinding()]
			param
			(
				[Parameter(Mandatory = $true,
						 ValueFromPipeline = $true,
						 Position = 0)]
				[ValidateNotNullOrEmpty()]
				$DomainFQDN,
				[Parameter(Mandatory = $true,
						 ValueFromPipeline = $true,
						 Position = 1)]
				[ValidateNotNullOrEmpty()]
				$gpoGUID
			)
			
			begin
			{
				#Gets the XML version of the GPO Report
				$GPOReport = Get-GPOReport -GUID $gpoGUID -ReportType XML -Domain $DomainFQDN
			}
			process
			{
				#Converts it to an XML variable for manipulation
				$GPOXML = [xml]$GPOReport
				
				#Create array to store info
				$GPOInfo = @()
				
				#Get's info from XML and adds to array
				#General Information
				
				#$GPODomain = $GPOXML.GPO.Domain
				$GPOInfo += , $DomainFQDN
				#$GPODomain = $GPOXML.Identifier.Domain
				#$GPOInfo += , $GPODomain
				
				$Name = $GPOXML.GPO.Name
				Write-Verbose -Message "Working on GPO $($Name)."
				$GPOInfo += , $Name
				
				$GPOGUID = $GPOXML.GPO.Identifier.Identifier.'#text'
				$GPOInfo += , $GPOGUID
				
				if (!([string]::IsNullOrEmpty($GPOXML.GPO.CreatedTime)))
				{
					[DateTime]$Created = $GPOXML.GPO.CreatedTime
					$GPOInfo += , $Created.ToString("G")
				}
				else
				{
					$Created = "No Creation Date Available."
					$GPOInfo += , $Created
				}
				
				
				[DateTime]$Modified = $GPOXML.GPO.ModifiedTime
				$GPOInfo += , $Modified.ToString("G")
				
				#WMI Filter
				if ($GPOXML.GPO.FilterName)
				{
					$WMIFilter = $GPOXML.GPO.FilterName
				}
				else
				{
					$WMIFilter = "<none>"
				}
				$GPOInfo += , $WMIFilter
				
				#Computer Configuration
				$ComputerEnabled = $GPOXML.GPO.Computer.Enabled
				$GPOInfo += , $ComputerEnabled
				
				$ComputerVerDir = $GPOXML.GPO.Computer.VersionDirectory
				$GPOInfo += , $ComputerVerDir
				
				$ComputerVerSys = $GPOXML.GPO.Computer.VersionSysvol
				$GPOInfo += , $ComputerVerSys
				
				if ($GPOXML.GPO.Computer.ExtensionData)
				{
					$ComputerExtensions = $GPOXML.GPO.Computer.ExtensionData | ForEach-Object { $_.Name }
					$ComputerExtensions = [String]::join("`n", $ComputerExtensions)
				}
				else
				{
					$ComputerExtensions = "<none>"
				}
				$GPOInfo += , $ComputerExtensions
				
				#User Configuration
				$UserEnabled = $GPOXML.GPO.User.Enabled
				$GPOInfo += , $UserEnabled
				
				$UserVerDir = $GPOXML.GPO.User.VersionDirectory
				$GPOInfo += , $UserVerDir
				
				$UserVerSys = $GPOXML.GPO.User.VersionSysvol
				$GPOInfo += , $UserVerSys
				
				if ($GPOXML.GPO.User.ExtensionData)
				{
					$UserExtensions = $GPOXML.GPO.User.ExtensionData | ForEach-Object { $_.Name }
					$UserExtensions = [string]::join("`n", $UserExtensions)
				}
				else
				{
					$UserExtensions = "<none>"
				}
				$GPOInfo += , $UserExtensions
				
				#Links
				if ($GPOXML.GPO.LinksTo)
				{
					$Links = $GPOXML.GPO.LinksTo | ForEach-Object { $_.SOMPath }
					$Links = [string]::join("`n", $Links)
					$LinksEnabled = $GPOXML.GPO.LinksTo | ForEach-Object { $_.Enabled }
					$LinksEnabled = [string]::join("`n", $LinksEnabled)
					$LinksNoOverride = $GPOXML.GPO.LinksTo | ForEach-Object { $_.NoOverride }
					$LinksNoOverride = [string]::join("`n", $LinksNoOverride)
				}
				else
				{
					$Links = "<none>"
					$LinksEnabled = "<none>"
					$LinksNoOverride = "<none>"
				}
				$GPOInfo += , $Links
				$GPOInfo += , $LinksEnabled
				$GPOInfo += , $LinksNoOverride
				
				#Security Info
				$Owner = $GPOXML.GPO.SecurityDescriptor.Owner.Name.'#text'
				$GPOInfo += , $Owner
				
				$SecurityInherits = $GPOXML.GPO.SecurityDescriptor.Permissions.InheritsFromParent
				$SecurityInherits = [string]::join("`n", $SecurityInherits)
				$GPOInfo += , $SecurityInherits
				
				$SecurityGroups = $GPOXML.GPO.SecurityDescriptor.Permissions.TrusteePermissions | ForEach-Object { $_.Trustee.Name.'#text' }
				if ($null -ne $SecurityGroups)
				{
					$SecurityGroups = [string]::join("`n", $SecurityGroups)
					$GPOInfo += , $SecurityGroups
				}
				else
				{
					$SecurityGroups = "None."
					$GPOInfo += , $SecurityGroups
				}
				
				$SecurityType = $GPOXML.GPO.SecurityDescriptor.Permissions.TrusteePermissions | ForEach-Object { $_.Type.PermissionType }
				$SecurityType = [string]::join("`n", $SecurityType)
				$GPOInfo += , $SecurityType
				
				$SecurityPerms = $GPOXML.GPO.SecurityDescriptor.Permissions.TrusteePermissions | ForEach-Object { $_.Standard.GPOGroupedAccessEnum }
				$SecurityPerms = [string]::join("`n", $SecurityPerms)
				$GPOInfo += , $SecurityPerms
			}
			end
			{
				return $GPOInfo
			}
			
		} #End function Get-GPOInfo
		
		
		$currentGPO = Get-GPOInfo -DomainFQDN $using:dnsRoot -gpoGUID $_.ID
		
		Write-Verbose -Message "Processing AD GPO $($currentGPO[1]) for domain $($domDNS)."
		if ($currentGPO[6] -eq 'true') { $computerConfig = "Enabled" }
		else { $computerConfig = "Disabled" }
		
		if ($currentGPO[10] -eq 'true') { $userConfig = "Enabled" }
		else { $userConfig = "Disabled" }
		
		$gpoObject += New-Object -TypeName PSCustomObject -Property ([ordered] @{
			"Domain Name" = $currentGPO[0]
			"GPO Name" = $currentGPO[1]
			"GPO GUID" = $currentGPO[2]
			"GPO Creation Time" = $currentGPO[3]
			"GPO Modification Date" = $currentGPO[4]
			"GPO WMIFilter" = $currentGPO[5]
			"GPO Computer Configuration" = [String]$computerConfig
			"GPO Computer Version Directory" = $currentGPO[7]
			"GPO Computer Sysvol Version" = $currentGPO[8]
			"GPO Computer Extensions" = $currentGPO[9]
			"GPO User Configuration" = [String]$userConfig
			"GPO User Version" = $currentGPO[11]
			"GPO User Sysvol Version" = $currentGPO[12]
			"GPO User Extension" = $currentGPO[13]
			"GPO Links" = $currentGPO[14]
			"GPO Link Enabled" = $currentGPO[15]
			"GPO Link Override" = $currentGPO[16]
			"GPO Owner" = $currentGPO[17]
			"GPO Inherits" = $currentGPO[18]
			"GPO Groups" = $currentGPO[19]
			"GPO Permission Type" = $currentGPO[20]
			"GPO Permissions" = $currentGPO[21]
		
		})
		
		$currentGPO = $computerConfig = $userConfig = $null
		
		Write-Output $gpoObject
			
	} -ThrottleLimit $throttleLimit
	
	[System.GC]::GetTotalMemory('forcefullcollection') | out-null
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
	$outputCSV = "{0}\{1}" -f $rptFolder, "$($strDomain)_GPO_Configuration_as_of_$(Get-ReportDate).csv"
	$outputFile = "{0}\{1}" -f $rptFolder, "$($strDomain)_GPO_Configuration_as_of_$(Get-ReportDate).xlsx"

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
	$Excel = $colResults | Export-Excel @ExcelParams -WorkSheetname "AD Group Policies" -PassThru
	$Sheet = $Excel.Workbook.Worksheets["AD Group Policies"]
	$totalRows = $Sheet.Dimension.Rows
	Set-Format -Address $Sheet.Cells["A2:Z$($totalRows)"] -Wraptext -VerticalAlignment Center -HorizontalAlignment Center
	Export-Excel -ExcelPackage $Excel -WorksheetName "AD Group Policies" -Title "$($strDomain) Active Directory Group Policy Configuration"  -TitleFillPattern Solid -TitleSize 18 -TitleBackgroundColor LightBlue
}

#endregion