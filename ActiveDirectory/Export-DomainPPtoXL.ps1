<#

.NOTES
	THIS CODE IS MADE AVAILABLE AS IS, WITHOUT WARRANTY OF ANY KIND. THE
	ENTIRE RISK OF THE USE OR THE RESULTS FROM THE USE OF THIS CODE REMAINS
	WITH THE USER.

.SYNOPSIS
	Get AD password policies

.DESCRIPTION
	This script executes an AD PowerShell cmdlet to gather the default domain
	password policies and exports the results to an Excel spreadsheet.

.OUTPUTS
	Excel spreadsheet with the default password policy information settings

.EXAMPLE 
	.\Get-DomainPPtoXL.ps1

#>
###########################################################################
#
#
# AUTHOR:  Heather Miller
#          
#
# VERSION HISTORY:
# 1.0 10/20/2017 - Initial release
# 2.0 04/20/2018 - Converted to data table use from PSCustomObject, implemented folder
# path check
# 3.0 08/08/2018 - Added ability to create data table and add columns
# 4.0 07/25/2019 - renamed functons and set all variables to clear after use
#
###########################################################################

#Region Modules
Try 
{
	Import-Module ActiveDirectory -ErrorAction Stop
}
Catch
{
	Throw "Active Directory module could not be loaded. $($_.Exception.Message)"
}
Try
{
	Import-Module ImportExcel -ErrorAction Stop
}
Catch
{
	Throw "Import Excel module could not be loaded. $($_.Exception.Message)"
}
#EndRegion


#Region Variables
$Domains = @()

#Get AD Forest Basic Information
$forestProperties = @("ApplicationPartitions", "Domains", "DomainNamingMaster", "ForestMode", "Name", "RootDomain", "PartitionsContainer", "SchemaMaster", "SPNSuffixes", "UPNSuffixes")
$Forest = Get-ADForest | Select-Object -Property $forestProperties
$forestName = ($Forest).Name.ToString().ToUpper()
$Domains = ($Forest).Domains
$rptFolder = "E:\Reports"

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

#EndRegion


#Region Functions

Function Get-ReportDate {#Begin function get report execution date
	Get-Date -Format "yyyy-MM-dd"
}#End function Get-ReportDate

Function Test-PathExists {#Begin function to check path variable and return results
 	[CmdletBinding()]
    Param
    (
        [Parameter(Mandatory,Position=0)]
        [string]$Path,
        [Parameter(Mandatory,Position=1)]
        $PathType
    )
    
    Switch ( $PathType )
    {
    		File	{
		   		If ( ( Test-Path -Path $Path -PathType Leaf ) -eq $true )
				{
					Write-Host "File: $Path already exists..." -BackgroundColor White -ForegroundColor Green
				}
				Else
				{
					New-Item -Path $Path -ItemType File -Force
					Write-Host "File: $Path not present, creating new file..." -BackgroundColor White -ForegroundColor Red
				}
			}
		Folder
			{
				If ( ( Test-Path -Path $Path -PathType Container ) -eq $true )
				{
					Write-Host "Folder: $Path already exists..." -BackgroundColor White -ForegroundColor Green
				}
				Else
				{
					New-Item -Path $Path -ItemType Directory -Force
					Write-Host "Folder: $Path not present, creating new folder" -BackgroundColor White -ForegroundColor Red
				}
			}
	}
}#End function Test-PathExists

#EndRegion









#Region Script
$Error.Clear()
Test-PathExists -Path $rptFolder -PathType Folder
$dtPPHeaders = ConvertFrom-Csv -InputObject $dtPPHeadersCsv

$tblName = "$($forestName)_Domain_Password_Policies"
$domPPTable = New-Object System.Data.DataTable $tblName

ForEach ($col in $dtPPHeaders)
{
    [void]$domPPTable.Columns.Add([System.Data.DataColumn]$col.ColumnName.ToString(), $col.DataType)
	$col = $null
}


$domainProperties = @( "DistinguishedName", "DNSRoot", "Forest", "InfrastructureMaster", "Name","PDCEmulator")
ForEach ( $Domain in $Domains )
{
	$domainInfo = Get-ADDomain -Identity $Domain | Select-Object -Property $domainProperties
	$domainDN = ($domainInfo).distinguishedName
	$domainName = ($domainInfo).DNSRoot
	$pdcFSMO = ($domainInfo).PDCEmulator

	#Region Domain Password Policies
	$Error.Clear()
	Try
	{
	   $defPP = Get-ADDefaultDomainPasswordPolicy -Identity $domainDN -Server $pdcFSMO
	}
	Catch
	{
		$defPP = Get-ADDefaultDomainPasswordPolicy -Identity $domainDN -Server $domainName
		$Error.Clear()
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
	$dtRow."Domain Name" = $domDN
	$dtRow."Complexity Enabled" = $complexityEnabled
	$dtRow."Lockout Duration" = $lockoutDuration
	$dtRow."Lockout Window" = $lockoutWindow
	$dtRow."Lockout Threshold" = $lockoutThreshold
	$dtRow."Max Password Age" =  $maxPWAge
	$dtRow."Min Password Age" =  $minPWAge
	$dtRow."Min Password Length" = $minPWLength
	$dtRow."Password History Count" = $pwHistoryCount
	$dtRow."Reversible Encryption Enabled" = $encryptionEnabled

	$domPPTable.Rows.Add($dtRow)
		
	$defPP = $domDN = $complexityEnabled = $lockoutDuration = $lockoutThreshold = $lockoutWindow = $maxPWAge = $minPWAge = $minPWLength = $pwHistoryCount = $encryptionEnabled = $null
	#EndRegion
	
$Domain = $domainInfo = $domainDN = $domainName = $pdcFSMO = $null
[GC]::Collect()
}

$domainProperties = $null

$outputFile = "$($forestName)_domain_pwd_policies_in_XL_as_of_$(Get-ReportDate).xlsx"
$newOutputFile = "{0}\{1}" -f $rptFolder, $outputFile
$ExcelParams = @{
	Path = $newOutputFile
	StartRow = 2
	StartColumn = 1
	AutoSize = $true
	AutoFilter = $true
	BoldTopRow = $true
	FreezeTopRow = $true
}

$colToExport = $dtPPHeaders.ColumnName
[String]$wsName = "AD Domains PP Configuration"
$Excel = $domPPTable  | Select-Object $colToExport | Sort-Object -Property "Domain Name"  | Export-Excel @ExcelParams -WorkSheetname $wsName -PassThru
$Sheet = $Excel.Workbook.Worksheets["AD Domains PP Configuration"]
$totalRows = $Sheet.Dimension.Rows
Set-Format -Address $Sheet.Cells["A2:Z$($totalRows)"] -Wraptext -VerticalAlignment Center -HorizontalAlignment Center
Export-Excel -ExcelPackage $Excel -WorksheetName $wsName -Title "Active Directory Domain Password Policies" -TitleSize 18 -TitleBackgroundColor LightBlue -TitleFillPattern Solid


#EndRegion