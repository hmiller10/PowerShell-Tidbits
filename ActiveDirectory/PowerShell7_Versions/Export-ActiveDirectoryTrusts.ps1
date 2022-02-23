﻿#Requires -Module ActiveDirectory, ImportExcel
#Requires -Version 7
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
	Excel spreasheet containing forest/domain trust information
.EXAMPLE 
	PS C:\>.\Export-ActiveDirectoryTrusts.ps1
#>

###########################################################################
#
#
# AUTHOR:  Heather Miller
#
# VERSION HISTORY:
# 1.0 08/18/2021 - Initial release
#
# 
###########################################################################

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

Try
{
	Import-Module -Name ImportExcel -SkipEditionCheck -ErrorAction Stop
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


#Region Variables
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
$sleepDurationSeconds = 5
$Forest = Get-ADForest
$forestName = ($Forest).Name.ToString().ToUpper()
$Domains = ($Forest).Domains
$domainProperties = @("DistinguishedName", "DNSRoot", "Forest", "Name", "NetBIOSName", "ParentDomain", "PDCEmulator")
$ns = 'root\MicrosoftActiveDirectory'
$trustHeadersCsv =
@"
		ColumnName,DataType
		"Source Name",string
		"Target Name",string
		"Forest Trust",string
		"IntraForest Trust",string
		"Trust Direction",string
		"Trust Type",string
		"Trust Attributes",string
		"SID History",string
		"SID Filtering",string
		"Selective Authentication",string
		"CIMPartnerDCName",string
		"CIMTrustIsOK",string
		"CIMTrustStatus",string
		"AD Trust whenCreated",string
		"AD Trust whenChanged",string
"@

[int32]$throttleLimit = 20
$dtmFormatString = "yyyy-MM-dd HH:mm:ss"
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
$trustTblName = "$($forestName)_Domain_Trust_Info"
$trustHeaders = ConvertFrom-Csv -InputObject $trustHeadersCsv
$trustTable = Add-DataTable -TableName $trustTblName -ColumnArray $trustHeaders
$Domains | ForEach-Object -Parallel {

	
	# List of properties of a trust relationship
	$trusts = @()
	$trustStatus = @()
	
	try
	{
		$domainInfo = Get-ADDomain -Identity $_ -Server (Get-ADDomain -Identity $_).pdcEmulator | Select-Object -Property $using:DomainProperties
	}
	catch
	{
		$domainInfo = Get-ADDomain -Identity $_ -Server $_ | Select-Object -Property $using:DomainProperties
	}
	
	$pdcFSMO = ($domainInfo).PDCEmulator
	$domDNS = ($domainInfo).DNSRoot
	
	
	try
	{
		$trusts = Get-ADTrust -Filter * -Properties * -Server $pdcFSMO -ErrorAction Continue | Select-Object -Property *
	}
	catch
	{
		$trusts = Get-ADTrust -Filter * -Properties * -Server $domDNS -ErrorAction Continue | Select-Object -Property *
	}
	
	try
	{
		$trustStatus = Get-CimInstance -ComputerName $pdcFSMO -Namespace $using:ns -Query "Select * from Microsoft_DomainTrustStatus" -ErrorAction Continue -ErrorVariable CIMError
	}
	catch
	{
		$trustStatus = Get-CimInstance -ComputerName $domDNS -Namespace $using:ns -Query "Select * from Microsoft_DomainTrustStatus" -ErrorAction Continue -ErrorVariable CIMError
	}
	
	$trusts | ForEach-Object {
		$t = $_
		$trustSource = Get-FQDNfromDN ($t).Source
		$trustTarget = ($t).Target
		$trustType = ($t).TrustType
		$forestTrust = ($t).ForestTransitive
		$intraForest = ($t).IntraForest
		$intTrustDirection = ($t).TrustDirection
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
			
			0 { $trustAttributes = "Inbound Trust" }
			1 { $trustAttributes = "Non-Transitive" }
			2 { $trustAttributes = "Uplevel clients only (Windows 2000 or newer" }
			4 { $trustAttributes = "Quarantined Domain (External)" }
			8 { $trustAttributes = "Forest Trust" }
			16 { $trustAttributes = "Cross-Organizational Trust (Selective Authentication)" }
			20 { $trustAttributes = "Intra-Forest Trust (trust within the forest)" }
			32 { $trustAttributes = "Intra-Forest Trust (trust within the forest)" }
			64 { $trustAttributes = "Inter-Forest Trust (trust with another forest)" }
			68 { $trustAttributes = "Quarantined Domain (External)" }
			4194304 { $trustAttributes = "Obsolete value combination"}
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
		
		$table = $using:trustTable
		$trustRow = $table.NewRow()
		$trustRow."Source Name" = $trustSource
		$trustRow."Target Name" = $trustTarget
		$trustRow."Forest Trust" = $forestTrust
		$trustRow."IntraForest Trust" = $intraForest
		$trustRow."Trust Direction" = $trustDirection
		$trustRow."Trust Type" = $trustType
		$trustRow."Trust Attributes" = $trustAttributes
		$trustRow."SID History" = $sidHistory
		$trustRow."SID Filtering" = $sidFiltering
		$trustRow."Selective Authentication" = $trustSelectiveAuth
		$trustRow."CIMPartnerDCName" = $partnerDC
		$trustRow."CIMTrustIsOK" = $trustOK
		$trustRow."CIMTrustStatus" = $Status
		$trustRow."AD Trust whenCreated" = $whenCreated
		$trustRow."AD Trust whenChanged" = $whenTrustChanged
		
		
		$table.Rows.Add($trustRow)
		[GC]::Collect()
	}

} -ThrottleLimit $throttleLimit


#Save output
#Check required folders and files exist, create if needed
$rptFolder = 'E:\Reports'
if ((Test-Path -Path $rptFolder -PathType Container) -eq $false) { New-Item -Path $rptFolder -ItemType Directory -Force }
$ttColToExport = $trustHeaders.ColumnName

Write-Verbose ("[{0} UTC] Exporting results data to CSV, please wait..." -f [datetime]::UtcNow.ToString($dtmFormatString))
$outputCsv = "{0}\{1}-{2}" -f $rptFolder, $(Get-UTCTime).ToString("yyyy-MM-dd_HH-mm-ss"), "{0}-Forest_Trust_Info.csv" -f $forestName
$trustTable | Select-Object $ttColToExport | Export-Csv -Path $outputCsv -NoTypeInformation

Write-Verbose ("[{0} UTC] Exporting results data in Excel format, please wait..." -f [datetime]::UtcNow.ToString($dtmFormatString))
$outputFile = "{0}\{1}-{2}" -f $rptFolder, $(Get-UTCTime).ToString("yyyy-MM-dd_HH-mm-ss"), "{0}-Forest_Trust_Info.xlsx" -f $forestName

$wsName = "AD Trust Configuration"
$ExcelParams = @{
	Path	        = $outputFile
	StartRow     = 2
	StartColumn  = 1
	AutoSize     = $true
	AutoFilter   = $true
	BoldTopRow   = $true
	FreezeTopRow = $true
}


$xl = $trustTable | Select-Object $ttColToExport | Export-Excel @ExcelParams -WorkSheetname $wsName -Passthru
$Sheet = $xl.Workbook.Worksheets["AD Trust Configuration"]
$totalRows = $Sheet.Dimension.Rows
Set-ExcelRange -Address $Sheet.Cells["A2:Z$($totalRows)"] -Wraptext -HorizontalAlignment Left -VerticalAlignment Bottom
Export-Excel -ExcelPackage $xl -WorksheetName $wsName -Title "$($forestName) Active Directory Trust Configuration" -TitleFillPattern Solid -TitleSize 14 -TitleBold -TitleBackgroundColor LightBlue


if (($Error).Count -eq 0)
{
	[String]$status = "Success"
}
else
{
	[String]$status = "Failed"
	if ($colJobErrors.Count -gt 0)
	{
		$fileStamp = (Get-Date).ToString("yyyy-MM-dd_HH-mm-ss")
		$errorFileName = "US_MF_CertRpt_Errors_as_of_{0}.csv" -f $fileStamp
		$errorFile = "{0}\{1}" -f $rptFolder, $errorFileName
		$colJobErrors | Export-Csv -Path $errorFile -Append -NoTypeInformation
	}
}

#EndRegion