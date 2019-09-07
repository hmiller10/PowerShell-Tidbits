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
	forest from which the script is run. The information is written to a datatable
	and then exported to a spreadsheet for artifact collection.
.OUTPUTS
	Excel spreasheet containing forest/domain trust information
.EXAMPLE 
	PS C:\>.\Export-ADTrustInfor2XL.ps1
#>

###########################################################################
#
#
# AUTHOR:  Heather Miller
#
# VERSION HISTORY:
# 1.0 07/26/2019 - Initial release
#
# 
###########################################################################


#Check if required module is loaded, if not load import it
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

#Global Variables
$Forest = Get-ADForest
$forestName = ($Forest).Name.ToString().ToUpper()
$Domains = ($Forest).Domains
$adRootDSE = Get-ADRootDSE
$rootRDNC = ($adRootDSE).rootDomainNamingContext
$rootCNC = ($adRootDSE).ConfigurationNamingContext
$rootDNC = ($adRootDSE).defaultNamingContext
$NCs = ($adRootDSE).namingContexts
$ns2 = 'root\MicrosoftActiveDirectory'

$trustHeadersCsv =
@"
ColumnName,DataType
"Source Name",string
"Target Name",string
"Forest Trust",string
"Trust Direction",string
"Trust Type",string
"Trust Attributes",string
"SID History",string
"SID Filtering",string
"Selective Authentication",string
"WMIPartnerDCName",string
"WMITrustIsOK",string
"WMITrustStatus",string
"AD whenChanged",string
"@

#Functions
Function Make-Table { #Begin function to dynamically build data table and add columns
  	[CmdletBinding()]
    Param
    (
        [Parameter(Mandatory,Position=0)]
        [String]$TableName,
        [Parameter(Mandatory,Position=1)]
        $ColumnArray
    )
	
	Begin {
		$dt = $null
		$dt = New-Object System.Data.DataTable("$TableName")
	}
	Process {
		ForEach ($col in $ColumnArray)
	 	 {
			[void]$dt.Columns.Add([System.Data.DataColumn]$col.ColumnName.ToString(), $col.DataType)
		 }
	}
	End {
		Return ,$dt
	}
}#end function Make-Table

Function Check-Path {#Begin function to check path variable and return results
 	[CmdletBinding()]
    Param
    (
        [Parameter(Mandatory,Position=0)]
        [String]$Path,
        [Parameter(Mandatory,Position=1)]
        $PathType
    )
    
    Switch ( $PathType )
    {
    	File
			{
		   		If ( ( Test-Path -Path $Path -PathType Leaf ) -eq $true )
				{
					#Write-Host "File: $Path already exists..." -BackgroundColor White -ForegroundColor Red
					Write-Verbose -Message "File: $Path already exists.."-Verbose
				}
				Else
				{
					New-Item -Path $Path -ItemType File -Force
					#Write-Host "File: $Path not present, creating new file..." -BackgroundColor Black -ForegroundColor Yellow
					Write-Verbose -Message "File: $Path not present, creating new file..." -Verbose
				}
			}
		Folder
			{
				If ( ( Test-Path -Path $Path -PathType Container ) -eq $true )
				{
					#Write-Host "Folder: $Path already exists..." -BackgroundColor White -ForegroundColor Red
					Write-Verbose -Message "Folder: $Path already exists..." -Verbose
				}
				Else
				{
					New-Item -Path $Path -ItemType Directory -Force
					#Write-Host "Folder: $Path not present, creating new folder"
					Write-Verbose -Message "Folder: $Path not present, creating new folder" -Verbose
				}
			}
	}
}#end function Check-Path

Function Utc-Now {#Begin function to get current date and time in UTC format
	[System.DateTime]::UtcNow
}#End function Utc-Now

Function Get-FileDate {#Begin function to get date and time in long format
	(Get-Date).ToString('yyyy-MM-dd-hh:mm:ss')
}#End function Get-FileDate

Function Get-MyInvocation {#Begin function to define $MyInvocation
	Return $MyInvocation
}#End function Get-MyInvocation

Function Get-ReportDate {#Begin function get report execution date
	Get-Date -Format "yyyy-MM-dd"
}#End function Get-ReportDate

Function Get-FQDNfromDN {
       [CmdletBinding()]
       Param
       (
           [Parameter(Mandatory = $true)]
           [string]$DistinguishedName
       )

    If ([string]::IsNullOrEmpty($DistinguishedName) -eq $true) {return $null}
    $domainComponents = $DistinguishedName.ToString().ToLower().Substring($DistinguishedName.ToString().ToLower().IndexOf("dc=")).Split(",")
    For($i = 0; $i -lt $domainComponents.count; $i++)
    {
        $domainComponents[$i] = $domainComponents[$i].Substring($domainComponents[$i].IndexOf("=") + 1)
    }
    $fqdn = [string]::Join(".", $domainComponents)

    Return $fqdn
}#End function Get-FQDNfromDN 







#Begin script
$Error.Clear()
$myInv = Get-MyInvocation
$scriptDir = $myInv.PSScriptRoot
$scriptName = $myInv.ScriptName

$trustTblName = "$($forestName)_Domain_Trust_Info"
$trustHeaders = ConvertFrom-Csv -InputObject $trustHeadersCsv
$trustTable = Make-Table -TableName $trustTblName -ColumnArray $trustHeaders

$domCount = 1
ForEach ($Domain in $Domains)
{
	$domActivityMessage = "Gathering domain information, please wait..."
	$domainStatus = "Processing domain {0} of {1}: {2}" -f $domCount, (Get-ADForest).Domains.count, $domain.Name
	$domPercentComplete = ($domCount / (Get-ADForest).Domains.count * 100)
	Write-Progress -Activity $domActivityMessage -Status $domainStatus -PercentComplete $domPercentComplete -Id 1
	
	Try
	{
		$domainInfo = Get-ADDomain -Identity $Domain -Server (Get-ADDomain -Identity $Domain).pdcEmulator | Select-Object -Property $domainProperties
	}
	Catch
	{
		$domainInfo = Get-ADDomain -Identity $Domain -Server $Domain | Select-Object -Property $domainProperties
	}
	
	$pdcFSMO = ($domainInfo).PDCEmulator
	$domainDN = ($domainInfo).distinguishedName
	$domDNS = ($domainInfo).DNSRoot
	
	#Region Trusts
	$Error.Clear()
	# List of properties of a trust relationship
	$trust = @()
	$trustStatus = @()
	Try
	{
		$trust += Get-ADTrust -Filter * -Properties * -Server $pdcFSMO
	}
	Catch
	{
		$trust += Get-ADTrust -Filter * -Properties * -Server $domDNS
		$Error.Clear()
	}
	
	Try
	{
		$trustStatus += Get-CimInstance -ClassName Microsoft_DomainTrustStatus -Namespace $ns2 -ComputerName $pdcFSMO -ErrorAction SilentlyContinue -ErrorVariable WMIError
	}
	Catch
	{
		$trustStatus += Get-CimInstance -ClassName Microsoft_DomainTrustStatus -Namespace $ns2 -ComputerName $domDNS -ErrorAction SilentlyContinue -ErrorVariable WMIError
		$Error.Clear()
	}
	
	If ( ( $trust ).Count -gt 1 )
	{
		$tCount = 1
		ForEach ( $t in $trust)
		{				
		  $trustActivityMessage = "Gathering AD trust information, please wait..."
			$trustProcessingStatus = "Processing trust {0} of {1}: {2}" -f $tCount, $trust.count, $t.flatName
			$trustsPercentComplete = ($tCount / $trust.count * 100)
			Write-Progress -Activity $trustActivityMessage -Status $trustProcessingStatus -PercentComplete $trustsPercentComplete -Id 2
	
			$trustSource = Get-FQDNfromDN ($t).Source
			$trustTarget = ($t).Target
			$trustType = ($t).TrustType
			$forestTrust = ($t).ForestTransitive

			$intTrustType = ($t).TrustDirection
			Switch ($intTrustType)
			{  
				0 {$trustDirection = "Disabled (The relationship exists but has been disabled)"}
				1 {$trustDirection = "Inbound (TrustING domain)"}
				2 {$trustDirection = "Outbound (TrustED domain)"}
				3 {$trustDirection = "Bidirectional (Two-Way Trust)"}
				Default{$intTrustType}

			}  

			$TrustAttributesNumber = ($t).TrustAttributes
			Switch ($TrustAttributesNumber)
			{  

				1 { $trustAttributes = "Non-Transitive"} 
				2 { $trustAttributes = "Uplevel clients only (Windows 2000 or newer"} 
				4 { $trustAttributes = "Quarantined Domain (External)"} 
				8 { $trustAttributes = "Forest Trust"} 
				16 { $trustAttributes = "Cross-Organizational Trust (Selective Authentication)"}
				20 { $trustAttributes = "Intra-Forest Trust (trust within the forest)"}
				32 { $trustAttributes = "Intra-Forest Trust (trust within the forest)"} 
				64 { $trustAttributes = "Inter-Forest Trust (trust with another forest)"}
				68 { $trustAttributes = "Quarantined Domain (External)" }
				Default { $trustAttributes = $TrustAttributesNumber }

			} 

			If (!$trustAttributes) { $trustAttributes = $TrustAttributesNumber }

			# Check if SID History is Enabled
			If ( $TrustAttributesNumber -band 64 ) { $sidHistory = "Enabled" } Else { $sidHistory = "Disabled" }

			# Check if SID Filtering is Enabled
			If ( ( ( $t.SIDFilteringQuarantined ) -eq $false ) -or ( ( $t.SIDFilteringForestAware ) -eq $false ) ) { $sidFiltering = "None" } Else { $sidFiltering = "Quarantine Enabled" }
					
			$trustSelectiveAuth = ($t).SelectiveAuthentication
			$whenTrustChanged = ($t).modifyTimeStamp -f "mm/dd/yyyy hh:mm:ss"
			ForEach ( $tStatus in $trustStatus )
			{
				If ( ( $trustStatus ).Count -gt 1)
				{
					$trustPartnerDC = ($tStatus).TrustedDCName
					$partnerDC = $trustPartnerDC.TrimStart("\\")
					If ( ( ($tStatus).TrustIsOk ) -eq $true ) { $trustOK = "Yes" } Else { $trustOK = "No - remediate" }
					$Status = ($tStatus).TrustStatusString
				}
			
			}
			
			$trustRow = $trustTable.NewRow()
			$trustRow."Source Name" = $trustSource
			$trustRow."Target Name" = $trustTarget
			$trustRow."Forest Trust" = $forestTrust
			$trustRow."Trust Direction" = $trustDirection
			$trustRow."Trust Type" = $trustType
			$trustRow."Trust Attributes" = $trustAttributes
			$trustRow."SID History" = $sidHistory
			$trustRow."SID Filtering" = $sidFiltering
			$trustRow."Selective Authentication" = $trustSelectiveAuth
			$trustRow."WMIPartnerDCName" = $partnerDC
			$trustRow."WMITrustIsOK" = $trustOK
			$trustRow."WMITrustStatus" = $Status
			$trustRow."AD whenChanged" = $whenTrustChanged

			$trustTable.Rows.Add($trustRow)
			
			$tStatus = $trustPartnerDC = $partnerDC = $Status = $trustOK = $null
			$t = $trustSource = $trustTarget = $forestTrust = $trustDirection = $trustType = $trustAttributes = $sidHistory = $sidFiltering = $trustSelectiveAuth = $whenTrustChanged = $null
			[GC]::Collect()
		
			$tCount++

		}#end ForEach $trust
		Write-Progress  -Activity "Done collecting all trust info. for $t.flatName" -Status "Ready" -Completed
	}#End If
	Else
	{
		$trustSource = Get-FQDNfromDN ($trust).Source
		$trustTarget = ($trust).Target
		$trustDirection = ($trust).Direction
		$trustType = ($trust).TrustType

		$intTrustType = ($trust).TrustDirection
		Switch ($intTrustType)
		{  

			1 {$trustDirection = "Inbound"}
			2 {$trustDirection = "Outbound"}
			3 {$trustDirection = "Bidirectional"}
			Default{$intTrustType = "Unknown"}

		} 

		If (!$trustType) { $trustType = $intTrustType }

		$TrustAttributesNumber = ($trust).TrustAttributes
		Switch ($TrustAttributesNumber)
		{  

			1 { $trustAttributes = "Non-Transitive"} 
			2 { $trustAttributes = "Uplevel clients only (Windows 2000 or newer"} 
			4 { $trustAttributes = "Quarantined Domain (External)"} 
			8 { $trustAttributes = "Forest Trust"} 
			16 { $trustAttributes = "Cross-Organizational Trust (Selective Authentication)"}
			20 { $trustAttributes = "Intra-Forest Trust (trust within the forest)"}
			32 { $trustAttributes = "Intra-Forest Trust (trust within the forest)"} 
			64 { $trustAttributes = "Inter-Forest Trust (trust with another forest)"}
			68 { $trustAttributes = "Quarantined Domain (External)" }
			Default { $trustAttributes = $TrustAttributesNumber }

		}  

		If (!$trustAttributes) { $trustAttributes = $TrustAttributesNumber }

		# Check if SID History is Enabled
		If ( $TrustAttributesNumber -band 64 ) { $sidHistory = "Enabled" } Else { $sidHistory = "Disabled" }

		# Check if SID Filtering is Enabled
		If ( ( ( $trust.SIDFilteringQuarantined ) -eq $false ) -or ( ( $trust.SIDFilteringForestAware ) -eq $false ) ) { $sidFiltering = "None" } Else { $sidFiltering = "Quarantine Enabled" }
				
		$trustSelectiveAuth = ($trust).SelectiveAuthentication
		$trustPartnerDC = ($trustStatus).TrustedDCName
		$partnerDC = $trustPartnerDC.TrimStart("\\")
		
		If ( ( ($trustStatus).TrustIsOk ) -eq $true ) { $trustOK = "Yes" } Else { $trustOK = "No - remediate" }
		$Status = ($trustStatus).TrustStatusString
		$whenTrustChanged = ($trust).modifyTimeStamp -f "mm/dd/yyyy hh:mm:ss"

		$trustRow = $trustTable.NewRow()
		$trustRow."Source Name" = $trustSource
		$trustRow."Target Name" = $trustTarget
		$trustRow."Trust Direction" = $trustDirection
		$trustRow."Trust Type" = $trustType
		$trustRow."Trust Attributes" = $trustAttributes
		$trustRow."SID History" = $sidHistory
		$trustRow."SID Filtering" = $sidFiltering
		$trustRow."Selective Authentication" = $trustSelectiveAuth
		$trustRow."WMIPartnerDCName" = $partnerDC
		$trustRow."WMITrustIsOK" = $trustOK
		$trustRow."WMITrustStatus" = $Status
		$trustRow."AD whenChanged" = $whenTrustChanged

		$trustTable.Rows.Add($trustRow)

		$Trust = $trustStatus = $null
		[GC]::Collect()
	$trustSource = $trustTarget = $trustDirection = $trustType = $trustAttributes = $sidHistory = $sidFiltering = $trustSelectiveAuth = $partnerDC = $trustOK = $Status = $whenTrustChanged = $null
	}#End Else
	
	$domCount++
	Write-Progress -Activity "Done collecting all info. for $domain" -Status "Ready" -Completed
}#End ForEach $Domain
	
#Save output
#Check required folders and files exist, create if needed
$rptFolder = 'E:\Reports'
Check-Path -Path $rptFolder -PathType Folder


$wsName = "AD Trust Configuration"
$outputFile = "{0}\{1}" -f $rptFolder, "$($forestName).Active_Directory_Trust_Info_as_of_$(Get-ReportDate).xlsx"
$ExcelParams = @{
    Path = $outputFile
    StartRow = 2
    StartColumn = 1
    AutoSize = $true
    AutoFilter = $true
    BoldTopRow = $true
    FreezeTopRow = $true
}

$ttColToExport = $trustHeaders.ColumnName
$xl = $trustTable | Select-Object $ttColToExport | Export-Excel @ExcelParams -WorkSheetname "AD Trust Configuration" -Passthru
$Sheet = $xl.Workbook.Worksheets["AD Trust Configuration"]
$totalRows = $Sheet.Dimension.Rows
Set-Format -Address $Sheet.Cells["A2:Z$($totalRows)"] -Wraptext -VerticalAlignment Bottom -HorizontalAlignment Left
Export-Excel -ExcelPackage $xl -WorksheetName "AD Trust Configuration" -Title "Active Directory Trust Configuration" -TitleSize 14 -TitleBold -TitleBackgroundColor LightBlue -TitleFillPattern Solid