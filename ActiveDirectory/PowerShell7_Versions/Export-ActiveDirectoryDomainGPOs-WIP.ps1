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

###########################################################################
#
#
# AUTHOR:  Heather Miller
#
# VERSION HISTORY: 2.0 - Converted output to data table and added CSV file output
# 
###########################################################################

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
	Import-Module ActiveDirectory -SkipEditionCheck -ErrorAction Stop
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
	Import-Module ImportExcel -Force
}
catch
{
	try
	{
		$module = Get-Module -Name ImportExcel;
		$modulePath = Split-Path $module.Path;
		$psdPath = "{0}\{1}" -f $modulePath, "ImportExcel.psd1"
		Import-Module $psdPath -ErrorAction Stop
	}
	catch
	{
		throw "ImportExcel PS module could not be loaded. $($_.Exception.Message)"
	}
}

try
{
	Import-Module GroupPolicy -SkipEditionCheck -ErrorAction Stop
}
catch
{
	throw "Group Policy module could not be loaded. $($_.Exception.Message)"
}
#EndRegion

#Region Variables
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

$domainParams = @{
	Identity    = $DomainName
	Server	  = $DomainName
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

$gpHeadersCsv =
@"
ColumnName,DataType
"Domain Name",string
"GPO Name",string
"GPO GUID",string
"GPO Creation Time",string
"GPO Modification Date",string
"GPO WMIFilter",string
"GPO Computer Configuration",string
"GPO Computer Version Directory",string
"GPO Computer Sysvol Version",string
"GPO Computer Extensions",string
"GPO User Configuration",string
"GPO User Version",string
"GPO User Sysvol Version",string
"GPO User Extension",string
"GPO Links",string
"GPO Link Enabled",string
"GPO Link Override",string
"GPO Owner",string
"GPO Inherits",string
"GPO Groups",string
"GPO Permission Type",string
"GPO Permissions",string
"@

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
		[String]$TableName,
		#'TableName'
		[Parameter(Mandatory = $true,
				 Position = 1)]
		[ValidateNotNullOrEmpty()]
		$ColumnArray #'DataColumnDefinitions'
	)
	
	
	begin
	{
		$dt = $null
		$dt = New-Object System.Data.DataTable("$TableName")
	}
	process
	{
		foreach ($col in $ColumnArray)
		{
			[void]$dt.Columns.Add([System.Data.DataColumn]$col.ColumnName.ToString(), $col.DataType)
		}
	}
	end
	{
		Write-Output @( ,$dt)
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
	
	[CmdletBinding(SupportsShouldProcess = $true)]
	param
	(
		[Parameter(Mandatory = $true,
				 Position = 0,
				 HelpMessage = 'Type the file system where the folder or file to check should be verified.')]
		[string]$Path,
		[Parameter(Mandatory = $true,
				 Position = 1,
				 HelpMessage = 'Specify path content as file or folder')]
		[string]$PathType
	)
	
	begin
	{
		$VerbosePreference = 'Continue';
	}
	
	process
	{
		switch ($PathType)
		{
			File
			{
				if ((Test-Path -Path $Path -PathType Leaf) -eq $true)
				{
					Write-Output ("File: {0} already exists..." -f $Path)
				}
				else
				{
					Write-Verbose -Message ("File: {0} not present, creating new file..." -f $Path)
					if ($PSCmdlet.ShouldProcess($Path, "Create file"))
					{
						[System.IO.File]::Create($Path)
					}
				}
			}
			Folder
			{
				if ((Test-Path -Path $Path -PathType Container) -eq $true)
				{
					Write-Output ("Folder: {0} already exists..." -f $Path)
				}
				else
				{
					Write-Verbose -Message ("Folder: {0} not present, creating new folder..." -f $Path)
					if ($PSCmdlet.ShouldProcess($Path, "Create folder"))
					{
						[System.IO.Directory]::CreateDirectory($Path)
					}
					
					
				}
			}
		}
	}
	
	end { }
	
} #end function Test-PathExists

function Get-UTCTime
{
<#
	.SYNOPSIS
		Get UTC Time
	
	.DESCRIPTION
		This functions returns the Universal Coordinated Date and Time. 
	
	.EXAMPLE
		PS C:\> Get-UTCTime
	
	.NOTES
		THIS CODE IS MADE AVAILABLE AS IS, WITHOUT WARRANTY OF ANY KIND. THE ENTIRE RISK OF 
		THE USE OR THE RESULTS FROM THE USE OF THIS CODE REMAINS WITH THE USER.
#>
	
	#Begin function to get current date and time in UTC format
	[System.DateTime]::UtcNow
} #End function Get-UTCTime

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
		$GpoGUID
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

#EndRegion







#Region Script
$Error.Clear()

$dtmFormatString = "yyyy-MM-dd HH:mm:ss"
$dtmFileFormatString = "yyyy-MM-dd_HH-mm-ss"

#Add data table to hold output results
$gpTblName = "$($dnsRoot)_Domain_GPO_Info"
$gpHeaders = ConvertFrom-Csv -InputObject $gpHeadersCsv
$gpTable = Add-DataTable -TableName $gpTblName -ColumnArray $gpHeaders

$GPOs = @()

try
{
	$GPOs = Get-GPO -Domain $DomainName -Server $pdcE -All
	if ($? -eq $false)
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
	
	$getGpoInfoDef = ${function:Get-GPOInfo}.ToString()
	
	$GPOs | ForEach-Object -Parallel {
		
		${function:Get-GPOInfo} = $using:getGpoInfoDef
		
		$currentGPO = Get-GPOInfo -DomainFQDN $using:dnsRoot -GpoGUID $_.ID
		
		Write-Verbose -Message "Processing AD GPO $($currentGPO[1]) for domain $($using:dnsRoot)."
		if ($currentGPO[6] -eq 'true') { $computerConfig = "Enabled" }
		else { $computerConfig = "Disabled" }
		
		if ($currentGPO[10] -eq 'true') { $userConfig = "Enabled" }
		else { $userConfig = "Disabled" }
		
		$table = $using:gpTable
		$gpRow = $table.NewRow()
		$gpRow."Domain Name" = $currentGPO[0]
		$gpRow."GPO Name" = $currentGPO[1]
		$gpRow."GPO GUID" = $currentGPO[2]
		$gpRow."GPO Creation Time" = $currentGPO[3]
		$gpRow."GPO Modification Date" = $currentGPO[4]
		$gpRow."GPO WMIFilter" = $currentGPO[5]
		$gpRow."GPO Computer Configuration" = [String]$computerConfig
		$gpRow."GPO Computer Version Directory" = $currentGPO[7]
		$gpRow."GPO Computer Sysvol Version" = $currentGPO[8]
		$gpRow."GPO Computer Extensions" = $currentGPO[9]
		$gpRow."GPO User Configuration" = [String]$userConfig
		$gpRow."GPO User Version" = $currentGPO[11]
		$gpRow."GPO User Sysvol Version" = $currentGPO[12]
		$gpRow."GPO User Extension" = $currentGPO[13]
		$gpRow."GPO Links" = $currentGPO[14]
		$gpRow."GPO Link Enabled" = $currentGPO[15]
		$gpRow."GPO Link Override" = $currentGPO[16]
		$gpRow."GPO Owner" = $currentGPO[17]
		$gpRow."GPO Inherits" = $currentGPO[18]
		$gpRow."GPO Groups" = $currentGPO[19]
		$gpRow."GPO Permission Type" = $currentGPO[20]
		$gpRow."GPO Permissions" = $currentGPO[21]
		
		$table.Rows.Add($gpRow)
		
		$currentGPO = $computerConfig = $userConfig = $null

	} -ThrottleLimit $throttleLimit
	
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

	Write-Verbose -Message "Exporting data tables to Excel spreadsheet tabs."
	$strDomain = $DomainName.ToString().ToUpper()
	
	$driveRoot = (Get-Location).Drive.Root
	$rptFolder = "{0}{1}" -f $driveRoot, "Reports"
	
	Test-PathExists -Path $rptFolder -PathType Folder
	
	$colToExport = $gpHeaders.ColumName
	
	Write-Verbose ("[{0} UTC] Exporting results data to CSV, please wait..." -f $(Get-UTCTime).ToString($dtmFormatString))
	$outputCSV = "{0}\{1}_{2}_Active_Directory_Domain_GPOs_Report.csv" -f $rptFolder, (Get-UTCTime).ToString($dtmFileFormatString), $strDomain
	$gpTable | Select-Object $colToExport | Export-Csv -Path $outputCSV -NoTypeInformation
	
	Write-Verbose ("[{0} UTC] Exporting results data in Excel format, please wait..." -f $(Get-UTCTime).ToString($dtmFormatString))
	$outputFile = "{0}\{1}_{2}_Active_Directory_OU_Structure_Report.xlsx" -f $rptFolder, (Get-UTCTime).ToString($dtmFileFormatString), $strDomain
	$ExcelParams = @{
		Path	        = $outputFile
		StartRow     = 2
		StartColumn  = 1
		AutoSize     = $true
		AutoFilter   = $true
		BoldTopRow   = $true
		FreezeTopRow = $true
	}
	
	$Excel = $gpTable | Select-Object $colToExport | Export-Excel @ExcelParams -WorkSheetname "AD Group Policies" -PassThru
	$Sheet = $Excel.Workbook.Worksheets["AD Group Policies"]
	$totalRows = $Sheet.Dimension.Rows
	Set-Format -Address $Sheet.Cells["A2:Z$($totalRows)"] -Wraptext -VerticalAlignment Bottom -HorizontalAlignment Left
	Export-Excel -ExcelPackage $Excel -WorksheetName "AD Group Policies" -Title "$($strDomain) Active Directory Group Policy Configuration" -TitleFillPattern Solid -TitleSize 18 -TitleBackgroundColor LightBlue
}

#endregion