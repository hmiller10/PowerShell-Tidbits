#Requires -Module ActiveDirectory, ImportExcel
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
	CSV file containing domain trust information
	Excel spreasheet containing domain trust information
.EXAMPLE 
	PS C:\>.\Export-ActiveDirectoryTrusts.ps1
#>

###########################################################################
#
#
# AUTHOR:  Heather Miller
#
# VERSION HISTORY:
# 3.0 10/22/2024 - Added additional trust attribute properties to report
#
# 
###########################################################################

param
(
[Parameter(Position = 0,
		 HelpMessage = 'Enter AD forest name to gather info. on.')]
[ValidateNotNullOrEmpty()]
[string[]]$DomainName,
[Parameter(Position = 1,
		 HelpMessage = 'Enter PS credential to connecct to AD forest with.')]
[ValidateNotNullOrEmpty()]
[pscredential]$Credential
)

#Region Modules
Try 
{
	Import-Module -Name ActiveDirectory -SkipEditionCheck -ErrorAction Stop
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

#EndRegion


#Region Variables
$domainProperties = @("DistinguishedName", "DNSRoot", "Forest", "Name", "NetBIOSName", "ParentDomain", "PDCEmulator")
$ns = 'root\MicrosoftActiveDirectory'
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

function Get-FQDNfromDN
{
<#
	.SYNOPSIS
		Convert DN to FQDN
	
	.DESCRIPTION
		This function converts an Active Directory distinguished name to a fully qualified domain name.
	
	.PARAMETER DistinguishedName
		AD distinguishedName
	
	.EXAMPLE
		PS C:\> Get-FQDNfromDN -DistinguishedName <ADDistinguisedName>
	
	.NOTES
		THIS CODE IS MADE AVAILABLE AS IS, WITHOUT WARRANTY OF ANY KIND. THE ENTIRE RISK OF 
		THE USE OR THE RESULTS FROM THE USE OF THIS CODE REMAINS WITH THE USER.
#>
	
	[CmdletBinding()]
	[OutputType([String])]
	param
	(
		[Parameter(Mandatory = $true)]
		[ValidateNotNullOrEmpty()]
		[string]$DistinguishedName
	)
	
	begin { }
	process
	{
		if ([string]::IsNullOrEmpty($DistinguishedName) -eq $true) { return $null }
		$domainComponents = $DistinguishedName.ToString().ToLower().Substring($DistinguishedName.ToString().ToLower().IndexOf("dc=")).Split(",")
		for ($i = 0; $i -lt $domainComponents.count; $i++)
		{
			$domainComponents[$i] = $domainComponents[$i].Substring($domainComponents[$i].IndexOf("=") + 1)
		}
		$fqdn = [string]::Join(".", $domainComponents)
	}
	end
	{
		return [string]$fqdn
	}
	
} #End function Get-FQDNfromDN

function Get-UTCTime {
<#
.SYNOPSIS
Gets current date and time in UTC format

.DESCRIPTION
Gets current date and time in UTC format

.INPUTS
None

.OUTPUTS
SYSTEM.DATETIME in UTC

.EXAMPLE
Get-UtcTime

#>
   [System.DateTime]::UtcNow

}#End function Get-UTCTime 
   
#EndRegion






#region Scripts
$Error.Clear()

#Create data table and add columns
$trustTblName = "tblDomaiTrusts"
$trustHeaders = ConvertFrom-Csv -InputObject $trustHeadersCsv
$trustTable = Add-DataTable -TableName $trustTblName -ColumnArray $trustHeaders

if (($PSBoundParameters.ContainsKey('DomainName')) -and ($null -ne $PSBoundParameters["DomainName"]))
{
	$Domains = $DomainName -split (",")
}
else
{
	try
	{
		$Domains = Get-ADDomain -Current LocalComputer -ErrorAction Stop
	}
	catch
	{
		$errorMessage = "{0}: {1}" -f $Error[0], $Error[0].InvocationInfo.PositionMessage
		Write-Error $errorMessage -ErrorAction Stop
	}
}

foreach ($Domain in $Domains)
{
	
	# List of properties of a trust relationship
	$trusts = @()
	$trustStatus = @()
	
	$domainParams = @{
		Identity = $Domain
	}
	
	if (($PSBoundParameters.ContainsKey('Credential')) -and ($null -ne $PSBoundParameters["Credential"]))
	{
		$domainParams.Add('AuthType', 'Negotitate')
		$domainParams.Add('Credential', $Credential)
	}
	
	try
	{
		$domainInfo = Get-ADDomain @domainParams -Server (Get-ADDomain @domainParams).pdcEmulator -ErrorAction SilentlyContinue | Select-Object -Property $DomainProperties
		if ($? -eq $false)
		{
			$domainInfo = Get-ADDomain @domainParams -Server $Domain -ErrorAction Stop | Select-Object -Property $DomainProperties
		}
	}
	catch
	{
		$errorMessage = "{0}: {1}" -f $Error[0], $Error[0].InvocationInfo.PositionMessage
		Write-Error $errorMessage -ErrorAction Continue
		break
	}
	
	$pdcFSMO = ($domainInfo).PDCEmulator
	$domDNS = ($domainInfo).DNSRoot
	
	$trustParams = @{
		Filter = '*'
		Properties = '*'
	}
	
	if (($PSBoundParameters.ContainsKey('Credential')) -and ($null -ne $PSBoundParameters["Credential"]))
	{
		$trustParams.Add('AuthType', 'Negotitate')
		$trustParams.Add('Credential', $Credential)
	}
	
	try
	{
		$trusts = Get-ADTrust @trustParams -Server $pdcFSMO -ErrorAction SilentlyContinue | Select-Object -Property *
		if ($? -eq $false)
		{
			$trusts = Get-ADTrust @trustParams -Server $domDNS -ErrorAction Stop | Select-Object -Property *
		}
	}
	catch
	{
		$errorMessage = "{0}: {1}" -f $Error[0], $Error[0].InvocationInfo.PositionMessage
		Write-Error $errorMessage -ErrorAction Continue
	}
	
	if ($trusts.Count -ge 1)
	{
		try
		{
			$trustStatus = Get-CimInstance -ComputerName $pdcFSMO -Namespace $ns -Query "Select * from Microsoft_DomainTrustStatus" -ErrorAction SilentlyContinue -ErrorVariable CIMError
			if ($? -eq $false)
			{
				$trustStatus = Get-CimInstance -ComputerName $domDNS -Namespace $ns -Query "Select * from Microsoft_DomainTrustStatus" -ErrorAction Stop -ErrorVariable CIMError
			}
		}
		catch
		{
			$errorMessage = "{0}: {1}" -f $Error[0], $Error[0].InvocationInfo.PositionMessage
			Write-Error $errorMessage -ErrorAction Stop
		}
		
		foreach ($t in $trusts)
		{
			$trustSource = Get-FQDNfromDN ($t).Source
			$trustTarget = ($t).Target
			$trustType = ($t).TrustType
			$forestTrust = ($t).ForestTransitive
			$intraForest = ($t).IntraForest
			$intTrustDirection = ($t).TrustDirection
			$usesAESKeys = ($t).UsesAESKeys
			$usesRC4Encrption = ($t).UsesRC4Encryption
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
			$whenCreated = ($t).Created -f "mm/dd/yyyy hh:mm:ss"
			$whenTrustChanged = ($t).modifyTimeStamp -f "mm/dd/yyyy hh:mm:ss"

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
	}#end $Trusts.Count
	
}

#Save output
#Check required folders and files exist, create if needed
$rptFolder = 'E:\Reports'
if ((Test-Path -Path $rptFolder -PathType Container) -eq $false) { New-Item -Path $rptFolder -ItemType Directory -Force }
Test-PathExists -Path $rptFolder -PathType Folder

if ($trustTable.Rows.Count -gt 1)
{
	$ttColToExport = $trustHeaders.ColumnName
	
	Write-Verbose ("[{0} UTC] Exporting results data to CSV, please wait..." -f (Get-UTCTime).ToString($dtmFormatString))
	$outputFile = "{0}\{1}-{2}_Active_Directory_Domain_Trust_Info.csv" -f $rptFolder, (Get-UTCTime).ToString($dtmFileFormatString), $domDNS.ToString().ToUpper()
	$trustTable | Select-Object $ttColToExport | Export-Csv -Path $outputFile -NoTypeInformation
	
	Write-Verbose ("[{0} UTC] Exporting results data in Excel format, please wait..." -f $(Get-UTCTime).ToString($dtmFormatString))
	$wsName = "AD Trust Configuration"
	$xlParams = @{
		Path = $outputFile = "{0}\{1}_{2}_Active_Directory_Domain_Trust_Info.xlsx" -f $rptFolder, $(Get-UTCTime).ToString($dtmFileFormatString), $domDNS.ToString().ToUpper()
		WorkSheetName = $wsName
		TableStyle = 'Medium15'
		StartRow = 2
		StartColumn = 1
		AutoSize = $true
		AutoFilter = $true
		BoldTopRow = $true
		PassThru = $true
	}
	
	$xl = $trustTable | Select-Object $ttColToExport | Export-Excel @xlParams
	$Sheet = $xl.Workbook.Worksheets["AD Trust Configuration"]
	Set-ExcelRange -Range $Sheet.Cells["A2:Z2"] -WrapText -HorizontalAlignment Center -VerticalAlignment Center -AutoFit
	$cols = $Sheet.Dimension.Columns
	Set-ExcelRange -Range $Sheet.Cells["A3:Z$($cols)"] -Wraptext -HorizontalAlignment Left -VerticalAlignment Bottom
	Export-Excel -ExcelPackage $xl -WorksheetName $wsName -FreezePane 3, 0 -Title "Active Directory Domain Trust Configuration" -TitleBold -TitleSize 16
} #end If
#EndRegion