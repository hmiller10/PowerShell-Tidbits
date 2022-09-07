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
###########################################################################
#
#
# AUTHOR:  Heather Miller
#
# VERSION HISTORY: 2.0
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
Try 
{
	Import-Module ActiveDirectory -SkipEditionCheck -ErrorAction Stop
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
$pdcE = $Domain.pdcEmulator
$dnsRoot = $Domain.DnsRoot

[int32]$throttleLimit = 100

$ouHeadersCsv =
@"
ColumnName,DataType
"Domain",string
"OU Name",string
"Parent OU",string
"Child OUs",string
"Managed By",string
"Delegated Objects",string
"Linked GPOs",string
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
	
[CmdletBinding(SupportsShouldProcess = $true)]
	param
	(
		[Parameter( Mandatory = $true,
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
	
}#end function Test-PathExists

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

#EndRegion




#Region Script
$Error.Clear()

$dtmFormatString = "yyyy-MM-dd HH:mm:ss"
$dtmFileFormatString = "yyyy-MM-dd_HH-mm-ss"

#List properties to be collected into array for writing to OU tab
$OUs = @()
$ouProps = @("distinguishedName", "gpLink", "LinkedGroupPolicyObjects", "ManagedBy", "Name", "objectCategory", "objectClass", "whenCreated", "whenChanged")

#Add data table to hold output results
$ouTblName = "$($Domain.DnsRoot)_OU_Info"
$ouHeaders = ConvertFrom-Csv -InputObject $ouHeadersCsv
$ouTable = Add-DataTable -TableName $ouTblName -ColumnArray $ouHeaders

Write-Verbose -Message ("Gathering collection of AD Organizational Units for {0}" -f $Domain.Name)
try
{
	$OUs = Get-ADOrganizationalUnit -Filter * -Properties $ouProps -SearchBase $Domain.distinguishedName -SearchScope Subtree -ResultSetSize $null -Server ($Domain).pdcEmulator | Select-Object -Property $ouProps
	if ($? -eq $false)
	{
		try
		{
			$OUs = Get-ADOrganizationalUnit -Filter * -Properties $ouProps -SearchBase $Domain.distinguishedName -SearchScope Subtree -ResultSetSize $null -Server ($Domain).dnsRoot | Select-Object -Property $ouProps
		}
		catch
		{
			Write-Warning ("Error occurred getting list of AD OUs for {0}" -f $Domain.Name)
			$errorMessage = "{0}: {1}" -f $Error[0], $Error[0].InvocationInfo.PositionMessage
			Write-Error $errorMessage -ErrorAction Continue
		}
	}
	
	$OUs | ForEach-Object -Parallel {

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
			[Array]$ouChildNames = Get-ADOrganizationalUnit -LDAPFilter '(objectClass=organizationalUnit)' -Properties * -SearchBase $ouDN -SearchScope OneLevel -Server $using:pdcE -ResultSetSize $null | Select-Object -ExpandProperty DistinguishedName
			if ($? -eq $false)
			{
				[Array]$ouChildNames = Get-ADOrganizationalUnit -LDAPFilter '(objectClass=organizationalUnit)' -Properties * -SearchBase $ouDN -SearchScope OneLevel -Server $using:dnsRoot -ResultSetSize $null | Select-Object -ExpandProperty DistinguishedName
			}
			
			if (($ouChildNames).Count -ge 1)
			{
				$ChildOUs = [String]($ouChildNames -join "`n")
			}
			else
			{
				$ChildOUs = "None"
			}
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
						$ouGPODisplayNames = (Get-GPInheritance -Target $ouDN | `
							Select-Object -Property GpoLinks).GpoLinks | Select-Object -ExpandProperty DisplayName
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
		
		$table = $using:ouTable
		$ouRow = $table.NewRow()
		$ouRow."Domain" = $domDNS
		$ouRow."OU Name" = $ouDN
		$ouRow."Parent OU" = $ouParentName
		$ouRow."Child OUs" = $ChildOUs
		$ouRow."Managed By" = $ouMgr
		$ouRow."Delegated Objects" = "See separate report" #$OUDelegatePerms -join "`n"
		$ouRow."Linked GPOs" = $ouGPODisplayNames
		
		$table.Rows.Add($ouRow)
		
		$null = $ouDN = $ChildOUs = $OUParent = $ouParentName = $ouChildNames = $ChildOUs = $ouMgr = $ouGPOs = $ouGPONames = $ouGPODisplayNames
		[System.GC]::GetTotalMemory('ForceFullCollection') | Out-Null

	} -ThrottleLimit $throttleLimit
	
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
	
	$driveRoot = (Get-Location).Drive.Root
	$rptFolder = "{0}{1}" -f $driveRoot, "Reports"
	
	Test-PathExists -Path $rptFolder -PathType Folder
	
	$strDomain = $DomainName.ToString().ToUpper()
	
	$colToExport = $ouHeaders.ColumnName
	
	Write-Verbose ("[{0} UTC] Exporting results data to CSV, please wait..." -f $(Get-UTCTime).ToString($dtmFormatString))
	$outputCSV = "{0}\{1}_{2}_Active_Directory_OU_Structure_Report.csv" -f $rptFolder, (Get-UTCTime).ToString($dtmFileFormatString), $strDomain
	$ouTable | Select-Object $colToExport | Export-Csv -Path $outputCSV -NoTypeInformation
	
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
	
	$Excel = $ouTable | Select-Object $colToExport | Export-Excel @ExcelParams -WorkSheetname "AD Organizational Units" -PassThru
	$Sheet = $Excel.Workbook.Worksheets["AD Organizational Units"]
	$totalRows = $Sheet.Dimension.Rows
	Set-Format -Address $Sheet.Cells["A2:Z$($totalRows)"] -Wraptext -VerticalAlignment Bottom -HorizontalAlignment Left
	Export-Excel -ExcelPackage $Excel -WorksheetName "AD Organizational Units" -Title "$($strDomain) Active Directory OU Configuration"  -TitleFillPattern Solid -TitleSize 18 -TitleBackgroundColor LightBlue
}

#endregion