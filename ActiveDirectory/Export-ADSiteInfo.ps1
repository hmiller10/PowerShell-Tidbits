<#

.NOTES
	THIS CODE IS MADE AVAILABLE AS IS, WITHOUT WARRANTY OF ANY KIND. THE
	ENTIRE RISK OF THE USE OR THE RESULTS FROM THE USE OF THIS CODE REMAINS
	WITH THE USER.

.SYNOPSIS
	Export AD Site Info to Excel. Requires PowerShell module ImportExcel

.DESCRIPTION
	This script is desigend to gather and report information on all Active Directory sites
	in a given forest.

.LINK
	https://github.com/dfinke/ImportExcel

.OUTPUTS
	Excel file containing relevant site information

.EXAMPLE 
	.\Export-ADSiteInfo.ps1

#>
###########################################################################
#
#
# AUTHOR:  Heather Miller
#
# VERSION HISTORY: 6.0 - Improved object filtering for better performance
# 
###########################################################################

#Region Execution Policy
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned
#EndRegion

#Region Modules
#Check if required module is loaded, if not load import it
Try
{
	Import-Module ActiveDirectory -ErrorAction Stop
}
Catch
{
	Throw "Active Directory module could not be loaded. $($_.Exception.Message)";
	exit
}

Try
{
	Import-Module ImportExcel -ErrorAction Stop
}
Catch
{
	Throw "PowerShell ImportExcel module could not be loaded. $($_.Exception.Message)";
	exit
}

Try
{
	Import-Module GroupPolicy -ErrorAction Stop
}
Catch
{
	Throw "PowerShell Group Policy module could not be loaded. $($_.Exception.Message)";
	exit
}
#EndRegion

#Region Global Variables
$adRootDSE = Get-ADRootDSE
$forestName = (Get-ADForest).Name.ToString().ToUpper()
$rootCNC = ($adRootDSE).ConfigurationNamingContext
$rptFolder = 'E:\Reports'
$dtSiteHeadersCSV =
@"
ColumnName,DataType
"Site Name",string
"Site Location",string
"Site Links",string
"Adjacent Sites",string
"Subnets in Site",string
"Domains in Site",string
"Servers in Site",string
"Bridgehead Servers",string
"GPOs linked to Site",string
"Notes",string
"@
#EndRegion

#Region Functions

Function Check-Path
{
	#Begin function to check path variable and return results
	[CmdletBinding()]
	Param
	(
		[Parameter(Mandatory, Position = 0)]
		[string]$Path,
		[Parameter(Mandatory, Position = 1)]
		$PathType
	)
	
	Switch ($PathType)
	{
		File
		{
			If ((Test-Path -Path $Path -PathType Leaf) -eq $true)
			{
				#Write-Host "File: $Path already exists..." -BackgroundColor White -ForegroundColor Red
				Write-Verbose -Message "File: $Path already exists.." -Verbose
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
			If ((Test-Path -Path $Path -PathType Container) -eq $true)
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
} #end function Check-Path

Function Get-ReportDate
{
	#Begin function set report date format
	Get-Date -Format "yyyy-MM-dd"
} #End function Get-ReportDate

Function Get-GPSiteLink
{
	
	Param
	(
		[Parameter(Position = 0, ValueFromPipeline = $True)]
		[string]$SiteName = "Default-First-Site-Name",
		[Parameter(Position = 1)]
		[string]$Domain = "myDomain.com",
		[Parameter(Position = 2)]
		[string]$Forest = "MyForest.com"
	)
	
	Begin
	{
		Write-Verbose "Starting Function" -Verbose
		#define the permission constants hash table
		$GPPerms = @{
			"permGPOApply"			      = 65536;
			"permGPORead"			      = 65792;
			"permGPOEdit"			      = 65793;
			"permGPOEditSecurityAndDelete" = 65794;
			"permGPOCustom"			 = 65795;
			"permWMIFilterEdit"		      = 131072;
			"permWMIFilterFullControl"     = 131073;
			"permWMIFilterCustom"	      = 131074;
			"permSOMLink"			      = 1835008;
			"permSOMLogging"		      = 1573120;
			"permSOMPlanning"		      = 1573376;
			"permSOMGPOCreate"		      = 1049600;
			"permSOMWMICreate"		      = 1049344;
			"permSOMWMIFullControl"	      = 1049345;
			"permStarterGPORead"		 = 197888;
			"permStarterGPOEdit"		 = 197889;
			"permStarterGPOFullControl"    = 197890;
			"permStarterGPOCustom"	      = 197891;
		}
		
		#define the GPMC COM Objects
		$gpm = New-Object -ComObject "GPMGMT.GPM"
		$gpmConstants = $gpm.GetConstants()
		$gpmDomain = $gpm.GetDomain($domain, "", $gpmConstants.UseAnyDC)
	} #Begin
	Process
	{
		ForEach ($item in $siteName)
		{
			#connect to site container
			$SiteContainer = $gpm.GetSitesContainer($forest, $domain, $null, $gpmConstants.UseAnyDC)
			Write-Verbose "Connected to site container on $($SiteContainer.domainController)" -Verbose
			#get sites
			Write-Verbose "Getting $item" -Verbose
			$site = $SiteContainer.GetSite($item)
			Write-Verbose ("Found {0} sites" -f ($sites | measure-object).count) -Verbose
			if ($site)
			{
				Write-Verbose "Getting site GPO links"
				$links = $Site.GetGPOLinks()
				if ($links)
				{
					#add the GPO name
					Write-Verbose ("Found {0} GPO links" -f ($links | measure-object).count) -Verbose
					$links | Select @{ Name = "Name"; Expression = { ($gpmDomain.GetGPO($_.GPOID)).DisplayName } },
								 @{ Name = "Description"; Expression = { ($gpmDomain.GetGPO($_.GPOID)).Description } }, GPOID, Enabled, Enforced, GPODomain, SOMLinkOrder, @{ Name = "SOM"; Expression = { $_.SOM.Path } }
				} #if $links
			} #if $site
		} #foreach site  
		
	} #process
	End
	{
		Write-Verbose "Finished"
	} #end
} #End function Get-GPSiteLink

Function Get-FqdnFromDN
{
	[CmdletBinding()]
	Param
	(
		[Parameter(Mandatory = $true)]
		[string]$DistinguishedName
	)
	
	If ([string]::IsNullOrEmpty($DistinguishedName) -eq $true) { return $null }
	$domainComponents = $DistinguishedName.ToString().ToLower().Substring($DistinguishedName.ToString().ToLower().IndexOf("dc=")).Split(",")
	For ($i = 0; $i -lt $domainComponents.count; $i++)
	{
		$domainComponents[$i] = $domainComponents[$i].Substring($domainComponents[$i].IndexOf("=") + 1)
	}
	$fqdn = [string]::Join(".", $domainComponents)
	
	Return $fqdn
} #End function Get-FqdnFromDN  

#EndRegion








#Region Script
#Begin Script
$Error.Clear()
#Create data table and add columns
$dtSiteHeaders = ConvertFrom-Csv -InputObject $dtSiteHeadersCsv
$dtSites = New-Object System.Data.DataTable "$forestName Site Properties"

ForEach ($siteHeader in $dtSiteHeaders)
{
	[void]$dtSites.Columns.Add([System.Data.DataColumn]$siteHeader.ColumnName.ToString(), $siteHeader.DataType)
}

#Region SiteConfig
#Begin collecting AD Site Configuration info.
$Sites = [System.DirectoryServices.ActiveDirectory.Forest]::GetCurrentForest().Sites | Sort-Object -Property Name

$sitesCount = 1
ForEach ($Site in $Sites)
{
	#Write-Verbose -Message "Working on AD site $Site..." -Verbose
	$SiteName = [String]$Site.Name
	$sitesActivityMessage = "Gathering AD site information, please wait..."
	$sitesProcessingStatus = "Processing site {0} of {1}: {2}" -f $sitesCount, $Sites.count, $SiteName.ToString()
	$percentSitesComplete = ($sitesCount / $Sites.count * 100)
	Write-Progress -Activity $sitesActivityMessage -Status $sitesProcessingStatus -PercentComplete $PercentComplete -Id 1
	
	$SiteLocation = [String]($Site).Location
	$SCSubnets = [String]($Site.Subnets -join "`n")
	$SiteLinks = [String]($Site.SiteLinks -join "`n")
	$AdjacentSites = [String]($Site.AdjacentSites -join "`n")
	$SiteDomains = [String]($Site.Domains -join "`n")
	$SiteServers = [String]($Site.Servers -join "`n")
	$BridgeHeads = [String]($Site.BridgeHeadServers -join "`n")
	
	$adSite += Get-ADObject -Filter '( objectClass -eq "site") -and (Name -eq $SiteName)' -SearchBase "CN=Sites,$($rootCNC)" -SearchScope OneLevel -Properties name, distinguishedName, gPLink, gPOptions -ErrorAction SilentlyContinue
	$gpoCount = ($adSite).gpLink.count
	$gpoNames = @()
	$siteGPOS = @()
	
	If (($adSite).gpLink -eq $null)
	{
		$gpoNames = "None."
	}
	Else
	{
		ForEach ($siteDomain in ($site).Domains)
		{
			$siteGPOS += Get-GPSiteLink -SiteName $SiteName -Domain $siteDomain -Forest $forestName
		}
		
		ForEach ($siteGPO in $siteGPOS)
		{
			$id = ($siteGPO).GPOID
			$gpoDom = ($siteGPO).GPODomain
			$gpoInfo = Get-GPO -Guid $id -Domain $gpoDom -Server $gpoDom -ErrorAction SilentlyContinue
			$gpoName = $gpoInfo.DisplayName.ToString()
			
			$gpoNames += $gpoName
			
			$siteGPO = $id = $gpoDom = $gpoInfo = $gpoGUID = $gpoName = $null
		}
	}
	
	
	$siteRow = $dtSites.NewRow()
	$siteRow."Site Name" = $SiteName | Out-String
	$siteRow."Site Location" = $SiteLocation | Out-String
	$siteRow."Site Links" = $SiteLinks | Out-String
	$siteRow."Adjacent Sites" = $AdjacentSites | Out-String
	$siteRow."Subnets in Site" = $SCSubnets | Out-String
	$siteRow."Domains in Site" = $SiteDomains | Out-String
	$siteRow."Servers in Site" = $SiteServers | Out-String
	$siteRow."Bridgehead Servers" = $BridgeHeads | Out-String
	$siteRow."GPOs linked to Site" = $gpoNames -join "`n" | Out-String
	$siteRow."Notes" = $null | Out-String
	
	$dtSites.Rows.Add($siteRow)
	
	$Site = $SiteLocation = $siteGPOS = $SiteLinks = $SiteName = $SCSubnets = $AdjacentSites = $SiteDomains = $SiteServers = $BridgeHeads = $null
	$adSite = $gpoNames = $null
	[GC]::Collect()
	$sitesCount++
}

Write-Progress -Activity "Done gathering AD site information for $($forestName)" -Status "Ready" -Completed
#EndRegion

#Save output
Check-Path -Path $rptFolder -PathType Folder

$wsName = "AD Site Configuration"
$outputFile = "{0}\{1}" -f $rptFolder, "$($forestName)_Active_Directory_Site_Info_as_of_$(Get-ReportDate).xlsx"
$ExcelParams = @{
	Path	        = $outputFile
	StartRow     = 2
	StartColumn  = 1
	AutoSize     = $true
	AutoFilter   = $true
	FreezeTopRow = $true
}

$colToExport = $dtSiteHeaders.ColumnName
$Excel = $dtSites | Select-Object $colToExport | Sort-Object -Property "Site Name" | Export-Excel @ExcelParams -WorkSheetname $wsName -PassThru
$Sheet = $Excel.Workbook.Worksheets["AD Site Configuration"]
$totalRows = $Sheet.Dimension.Rows
Set-Format -Address $Sheet.Cells["A2:Z$($totalRows)"] -Wraptext -VerticalAlignment Bottom -HorizontalAlignment Left
Export-Excel -ExcelPackage $Excel -WorksheetName $wsName -Title "$($forestName) Active Directory Site Configuration" -TitleSize 16 -TitleBackgroundColor LightBlue -TitleFillPattern Solid

#EndRegion
# SIG # Begin signature block
# MIInCQYJKoZIhvcNAQcCoIIm+jCCJvYCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCAzLQ2u9cX2wH7P
# ZCvC/KCXfE6+iKym/ri7B9AwGf7k9qCCIaUwggQVMIIC/aADAgECAgsEAAAAAAEx
# icZQBDANBgkqhkiG9w0BAQsFADBMMSAwHgYDVQQLExdHbG9iYWxTaWduIFJvb3Qg
# Q0EgLSBSMzETMBEGA1UEChMKR2xvYmFsU2lnbjETMBEGA1UEAxMKR2xvYmFsU2ln
# bjAeFw0xMTA4MDIxMDAwMDBaFw0yOTAzMjkxMDAwMDBaMFsxCzAJBgNVBAYTAkJF
# MRkwFwYDVQQKExBHbG9iYWxTaWduIG52LXNhMTEwLwYDVQQDEyhHbG9iYWxTaWdu
# IFRpbWVzdGFtcGluZyBDQSAtIFNIQTI1NiAtIEcyMIIBIjANBgkqhkiG9w0BAQEF
# AAOCAQ8AMIIBCgKCAQEAqpuOw6sRUSUBtpaU4k/YwQj2RiPZRcWVl1urGr/SbFfJ
# MwYfoA/GPH5TSHq/nYeer+7DjEfhQuzj46FKbAwXxKbBuc1b8R5EiY7+C94hWBPu
# TcjFZwscsrPxNHaRossHbTfFoEcmAhWkkJGpeZ7X61edK3wi2BTX8QceeCI2a3d5
# r6/5f45O4bUIMf3q7UtxYowj8QM5j0R5tnYDV56tLwhG3NKMvPSOdM7IaGlRdhGL
# D10kWxlUPSbMQI2CJxtZIH1Z9pOAjvgqOP1roEBlH1d2zFuOBE8sqNuEUBNPxtyL
# ufjdaUyI65x7MCb8eli7WbwUcpKBV7d2ydiACoBuCQIDAQABo4HoMIHlMA4GA1Ud
# DwEB/wQEAwIBBjASBgNVHRMBAf8ECDAGAQH/AgEAMB0GA1UdDgQWBBSSIadKlV1k
# sJu0HuYAN0fmnUErTDBHBgNVHSAEQDA+MDwGBFUdIAAwNDAyBggrBgEFBQcCARYm
# aHR0cHM6Ly93d3cuZ2xvYmFsc2lnbi5jb20vcmVwb3NpdG9yeS8wNgYDVR0fBC8w
# LTAroCmgJ4YlaHR0cDovL2NybC5nbG9iYWxzaWduLm5ldC9yb290LXIzLmNybDAf
# BgNVHSMEGDAWgBSP8Et/qC5FJK5NUPpjmove4t0bvDANBgkqhkiG9w0BAQsFAAOC
# AQEABFaCSnzQzsm/NmbRvjWek2yX6AbOMRhZ+WxBX4AuwEIluBjH/NSxN8RooM8o
# agN0S2OXhXdhO9cv4/W9M6KSfREfnops7yyw9GKNNnPRFjbxvF7stICYePzSdnno
# 4SGU4B/EouGqZ9uznHPlQCLPOc7b5neVp7uyy/YZhp2fyNSYBbJxb051rvE9ZGo7
# Xk5GpipdCJLxo/MddL9iDSOMXCo4ldLA1c3PiNofKLW6gWlkKrWmotVzr9xG2wSu
# kdduxZi61EfEVnSAR3hYjL7vK/3sbL/RlPe/UOB74JD9IBh4GCJdCC6MHKCX8x2Z
# faOdkdMGRE4EbnocIOM28LZQuTCCBMYwggOuoAMCAQICDCRUuH8eFFOtN/qheDAN
# BgkqhkiG9w0BAQsFADBbMQswCQYDVQQGEwJCRTEZMBcGA1UEChMQR2xvYmFsU2ln
# biBudi1zYTExMC8GA1UEAxMoR2xvYmFsU2lnbiBUaW1lc3RhbXBpbmcgQ0EgLSBT
# SEEyNTYgLSBHMjAeFw0xODAyMTkwMDAwMDBaFw0yOTAzMTgxMDAwMDBaMDsxOTA3
# BgNVBAMMMEdsb2JhbFNpZ24gVFNBIGZvciBNUyBBdXRoZW50aWNvZGUgYWR2YW5j
# ZWQgLSBHMjCCASIwDQYJKoZIhvcNAQEBBQADggEPADCCAQoCggEBANl4YaGWrhL/
# o/8n9kRge2pWLWfjX58xkipI7fkFhA5tTiJWytiZl45pyp97DwjIKito0ShhK5/k
# Ju66uPew7F5qG+JYtbS9HQntzeg91Gb/viIibTYmzxF4l+lVACjD6TdOvRnlF4RI
# shwhrexz0vOop+lf6DXOhROnIpusgun+8V/EElqx9wxA5tKg4E1o0O0MDBAdjwVf
# ZFX5uyhHBgzYBj83wyY2JYx7DyeIXDgxpQH2XmTeg8AUXODn0l7MjeojgBkqs2Iu
# YMeqZ9azQO5Sf1YM79kF15UgXYUVQM9ekZVRnkYaF5G+wcAHdbJL9za6xVRsX4ob
# +w0oYciJ8BUCAwEAAaOCAagwggGkMA4GA1UdDwEB/wQEAwIHgDBMBgNVHSAERTBD
# MEEGCSsGAQQBoDIBHjA0MDIGCCsGAQUFBwIBFiZodHRwczovL3d3dy5nbG9iYWxz
# aWduLmNvbS9yZXBvc2l0b3J5LzAJBgNVHRMEAjAAMBYGA1UdJQEB/wQMMAoGCCsG
# AQUFBwMIMEYGA1UdHwQ/MD0wO6A5oDeGNWh0dHA6Ly9jcmwuZ2xvYmFsc2lnbi5j
# b20vZ3MvZ3N0aW1lc3RhbXBpbmdzaGEyZzIuY3JsMIGYBggrBgEFBQcBAQSBizCB
# iDBIBggrBgEFBQcwAoY8aHR0cDovL3NlY3VyZS5nbG9iYWxzaWduLmNvbS9jYWNl
# cnQvZ3N0aW1lc3RhbXBpbmdzaGEyZzIuY3J0MDwGCCsGAQUFBzABhjBodHRwOi8v
# b2NzcDIuZ2xvYmFsc2lnbi5jb20vZ3N0aW1lc3RhbXBpbmdzaGEyZzIwHQYDVR0O
# BBYEFNSHuI3m5UA8nVoGY8ZFhNnduxzDMB8GA1UdIwQYMBaAFJIhp0qVXWSwm7Qe
# 5gA3R+adQStMMA0GCSqGSIb3DQEBCwUAA4IBAQAkclClDLxACabB9NWCak5BX87H
# iDnT5Hz5Imw4eLj0uvdr4STrnXzNSKyL7LV2TI/cgmkIlue64We28Ka/GAhC4evN
# GVg5pRFhI9YZ1wDpu9L5X0H7BD7+iiBgDNFPI1oZGhjv2Mbe1l9UoXqT4bZ3hcD7
# sUbECa4vU/uVnI4m4krkxOY8Ne+6xtm5xc3NB5tjuz0PYbxVfCMQtYyKo9JoRbFA
# uqDdPBsVQLhJeG/llMBtVks89hIq1IXzSBMF4bswRQpBt3ySbr5OkmCCyltk5lXT
# 0gfenV+boQHtm/DDXbsZ8BgMmqAc6WoICz3pZpendR4PvyjXCSMN4hb6uvM0MIIF
# fzCCA2egAwIBAgIQGLXChEOQEpdBrAmKM2WmEDANBgkqhkiG9w0BAQsFADBSMRMw
# EQYKCZImiZPyLGQBGRYDY29tMRgwFgYKCZImiZPyLGQBGRYIRGVsb2l0dGUxITAf
# BgNVBAMTGERlbG9pdHRlIFNIQTIgTGV2ZWwgMSBDQTAeFw0xNTA5MDExNTA3MjVa
# Fw0zNTA5MDExNTA3MjVaMFIxEzARBgoJkiaJk/IsZAEZFgNjb20xGDAWBgoJkiaJ
# k/IsZAEZFghEZWxvaXR0ZTEhMB8GA1UEAxMYRGVsb2l0dGUgU0hBMiBMZXZlbCAx
# IENBMIICIjANBgkqhkiG9w0BAQEFAAOCAg8AMIICCgKCAgEAlPqNqqVpE41dp1s1
# +neM+Xv5zfUAKTrD10RAF9epFFmIIMH62VgMXOYYWBryNQaUAYPZlvv/Tt0cCKca
# 5XAWKp4DbBeblCmxfHsqEz3R/kzn/CHRHnQ3YMZRMorAccq82DdxKiwnw9o0W5SG
# D5A+zNXh9DjcCx0G5ROAaqiv7m3HYz2HrEvqdIuMkMoj7Y2ieMiw/PuIjVU8wmod
# ltkBmGoAeOOcVYaWBZTpKy0NC/xYL7eHfMKdgRaa30pFVeZliN8DMiN/exbfr6iu
# 00fQAsNxiZleH/6CLHuODdh+7KK00Wp2Wi9qz/IeOAGkj8j0jXFnnX5PHQWcVVv8
# E8sIK1S95xDxmhOsrMGkGA6G3F7a1qfI1WntvYBT98eUgZQ3whDqjypj622jjXLk
# UxlfuUeuBHB2+T9kSbapQHIhjAE3f97A/FOuzG0aerr6eNC5doNjOX31Bfp5W0Wk
# hbX8D0Aexf7v+OsboqFkAkaNzSS2oaX7+G3XAw2r+slDmyimr+boaLEo4vM+oFzF
# UeBQOXvjGBEnGtxXmSIPwsLu+HlhOvjtXINLbsczl2QWzC2arRPxx6HLr1hPj0ei
# yz7bKDPQ+N+U9l5OetL6NNFgppVDoqSVo5FUwh47wZKaqXZ8b1jPj/SS+IRsbKnC
# J37+YXfkA2Mid9x8oMyRfBfwed8CAwEAAaNRME8wCwYDVR0PBAQDAgGGMA8GA1Ud
# EwEB/wQFMAMBAf8wHQYDVR0OBBYEFL6LoCtmWVn6kHFRpaeoBkJOkDztMBAGCSsG
# AQQBgjcVAQQDAgEAMA0GCSqGSIb3DQEBCwUAA4ICAQCDTvRyLG76375USpexKf0B
# GCuYfW+o/6G18GRqZeYls7lO251xn7hfXacfEZIHCPoizan0yvZJtYUocXRMieo7
# 66Zwn8g4OgEZjJXsw81p0GlkylmdWhqO+sRuGyYvGY32MWZ16oz6x/CG+rseou2H
# sLLtlSV76D2XPnDutIAHI/S4is4A7F0V+oNX04aHpUXMb0Y1BkPKNF1gIlmf4rdt
# Rh6+2r374QP+Ruw+nJiPNwF7TF28wkz1iUXWK9FSmM1Q6+/uXxpx9qRFRwv+pCd/
# 07IneZ3GmxxTNJxSzzEJxIfwoJIn6HL9NYPltAZ7CuWYsm5TFY+x5TZ5qS/O6+nA
# Hd30T7K/q+H5hjp9tisYah3RiBOOU+iZvtUsr1XaLT7zizxnmp4ssHHryLhNkYu2
# uh/dT1/iq8SbM3fKGElML+mE7ZPAg2q2B76kgbY+GrEtzNnzwNfIwkh/IDKYJ9n6
# JU2yQ4oa5sJjTf5uHUhxV9Zd8/BZK8L3H5S7Iy3yCVLyq98xuUZ3ChL4FoKeS89u
# MrgKADP2xnAdIw1nnd67ZSPrTVk3sZO/uJVKTzjpU0V10sc27VmVx9YByc4o4xDo
# Q6+eAlUbNpuoFpchzdL2dx5JUalLl2T4jg4UIzKcidPhEmyU1ApKUXFQTbx0N8v1
# WC2UXROwuc0YDLR7v6RCLjCCBdwwggPEoAMCAQICEz4AAAAHOQSYtK/MwjQAAQAA
# AAcwDQYJKoZIhvcNAQELBQAwVDETMBEGCgmSJomT8ixkARkWA2NvbTEYMBYGCgmS
# JomT8ixkARkWCERlbG9pdHRlMSMwIQYDVQQDExpEZWxvaXR0ZSBTSEEyIExldmVs
# IDIgQ0EgMjAeFw0xODEyMDQyMDEyNTBaFw0yMzEyMDQyMDIyNTBaMGwxEzARBgoJ
# kiaJk/IsZAEZFgNjb20xGDAWBgoJkiaJk/IsZAEZFghkZWxvaXR0ZTEWMBQGCgmS
# JomT8ixkARkWBmF0cmFtZTEjMCEGA1UEAxMaRGVsb2l0dGUgU0hBMiBMZXZlbCAz
# IENBIDIwggEiMA0GCSqGSIb3DQEBAQUAA4IBDwAwggEKAoIBAQCvao3WR6CSYsMe
# I4sWzR6nXvczKc7voHTVzi/q3LbOD6j6YQNa/WnJeDITb2yf8BcIUXeLqm9dd64S
# in69YS3gTLT7ZFucodBp11g6IaA1R40tbWW9x2WDxYGMDoN+Hvq78bQMsSFEo1Ad
# mZRS/GGCO69u0ROyFtAgRt3E4jLFuzm1RWiNdEl00qNYnmaN4iLz2dEnKtJm+Cl2
# NH1xlB+m47ovgHlejoqJ/eg9kLmwEZam8o2SzgMrBup85GO8UmV55f3mv7zrRNhe
# oL+rdBAqN3NsA3n2a2JmLZAkcRD5Zk5I46EnJhRZpguRoafd4INeOPYH2iKNKqpe
# HFIbyWKPAgMBAAGjggGNMIIBiTAQBgkrBgEEAYI3FQEEAwIBATAjBgkrBgEEAYI3
# FQIEFgQURHtJJiaF3HfA4va/QlnTnPpod7wwHQYDVR0OBBYEFGmVYfUC8O4CaCIJ
# kuTjxIa0u/lpMBkGCSsGAQQBgjcUAgQMHgoAUwB1AGIAQwBBMAsGA1UdDwQEAwIB
# hjASBgNVHRMBAf8ECDAGAQH/AgEAMB8GA1UdIwQYMBaAFEcuNu60nP9cXhh8uBPh
# vqkgHhSzMFwGA1UdHwRVMFMwUaBPoE2GS2h0dHA6Ly9wa2kuZGVsb2l0dGUuY29t
# L0NlcnRFbnJvbGwvRGVsb2l0dGUlMjBTSEEyJTIwTGV2ZWwlMjAyJTIwQ0ElMjAy
# LmNybDB2BggrBgEFBQcBAQRqMGgwZgYIKwYBBQUHMAKGWmh0dHA6Ly9wa2kuZGVs
# b2l0dGUuY29tL0NlcnRFbnJvbGwvU0hBMkxWTDJDQTJfRGVsb2l0dGUlMjBTSEEy
# JTIwTGV2ZWwlMjAyJTIwQ0ElMjAyKDEpLmNydDANBgkqhkiG9w0BAQsFAAOCAgEA
# T4VkpKHJQHX5pk2FaNiXUHQKkZQXs/uD8lbhSdUgPqZCUaD7rml/aqzusVpA2GML
# zrsUcomq7xt4S9kOKIGQabSUeg681nGvzXrp0P8xOsXYUWqR9PIcEkfdDYs3pNce
# S98TAFl8+hKkMm2XMDaOpBz7AT6xb5ISKEybUWf/Gsdfmha1UzfCtIDVQUdWQcFD
# FQnFfVL4gcKfmwp7fq5bZi5l4/4kMM1OP1s10Og8PaAPhRkaYdQapDbaT82czXZS
# v0dqimBXWImTAJx9PbcWc5iqmNtrUxPsYCt2yGNByO3spCIa96MqfkiQQBISZxRr
# NT6pjMGtdR3Kij/rixmEBy/ITd4Ua4Za4TR09C8Lw/+ukmdV3D0G+3zRwqcwURAV
# Bvxwp62sVe0+yUYnckwmIiwbI9X8VYyCURk0YvKqfsXRZjnWtGOhSjT2EnxO87e4
# hrO4G9akInQvzAL6giL/K4UCzpl4qotDlYK8PzvmsceuGWx23nZaQQ3K21FgNduo
# HIvqVuslCf+u7Z/ZYCwguGb6xKIzDS1vpzkqMuSHa1gxsmLm+PzMyM4i9E9FFnbX
# vKf3P6SXyk0yXi7bB/KcG9t7QsITpZ7X+LA2+gWDY2LE1i7XLsOoOn5KaV70sTB6
# PoL5qaqOJAxswoJ0t2j1itrhsG7y/GUPhcG3kWq9V+EwggaOMIIFdqADAgECAhNl
# ADwUlwDjJaP1Xv9CAAAAPBSXMA0GCSqGSIb3DQEBCwUAMGwxEzARBgoJkiaJk/Is
# ZAEZFgNjb20xGDAWBgoJkiaJk/IsZAEZFghkZWxvaXR0ZTEWMBQGCgmSJomT8ixk
# ARkWBmF0cmFtZTEjMCEGA1UEAxMaRGVsb2l0dGUgU0hBMiBMZXZlbCAzIENBIDIw
# HhcNMTgxMTE0MTQxNjU1WhcNMjAxMDI5MTg1OTI0WjAZMRcwFQYDVQQDEw5IZWF0
# aGVyIE1pbGxlcjCCASIwDQYJKoZIhvcNAQEBBQADggEPADCCAQoCggEBANNqwwoQ
# yJmuG4mVQZhiZ9GyeXKwxRsjzRbeDmi5PkDuNKCF0zm/3nJqRyMSAXkNL9KElsqm
# 8lHwphrJfo/XJgxBRSkSY+4Y+Fh4pCeBQAevNXE2wA3A1sEsmaP4uxKgEUtbJEDS
# 35h9SEDvj+esroKB09wa6qFkaTjaWq6GnhYzHWts2BFTaJ3iHu+mNBdZRfYH0jgg
# HEcGGRZaMmrXhGm0mf9UmZxgZZG4/mu9ZFdLOgV3Spwy897XmjMdpzlBZtvgKn44
# UpXwfw5PxEK4Ygx+VbaPIwJ0sKRZrbYyLeaTleVBm1ckK76t+b2/sITPx3gv1Sv9
# 8ECrBEWRrqD2muECAwEAAaOCA3owggN2MDwGCSsGAQQBgjcVBwQvMC0GJSsGAQQB
# gjcVCIGBvUmFvoUTgtWbPIPXjgeG8ckKXIPK9y3C8zICAWQCAR4wEwYDVR0lBAww
# CgYIKwYBBQUHAwMwCwYDVR0PBAQDAgeAMBsGCSsGAQQBgjcVCgQOMAwwCgYIKwYB
# BQUHAwMwIAYDVR0RBBkwF4EVaGVtaWxsZXJAZGVsb2l0dGUuY29tMB0GA1UdDgQW
# BBTYjzXkhwjhstLpzEior3SlOAA+RDAfBgNVHSMEGDAWgBRplWH1AvDuAmgiCZLk
# 48SGtLv5aTCCATsGA1UdHwSCATIwggEuMIIBKqCCASagggEihoHSbGRhcDovLy9D
# Tj1EZWxvaXR0ZSUyMFNIQTIlMjBMZXZlbCUyMDMlMjBDQSUyMDIsQ049dXNhdHJh
# bWVlbTAwNCxDTj1DRFAsQ049UHVibGljJTIwS2V5JTIwU2VydmljZXMsQ049U2Vy
# dmljZXMsQ049Q29uZmlndXJhdGlvbixEQz1kZWxvaXR0ZSxEQz1jb20/Y2VydGlm
# aWNhdGVSZXZvY2F0aW9uTGlzdD9iYXNlP29iamVjdENsYXNzPWNSTERpc3RyaWJ1
# dGlvblBvaW50hktodHRwOi8vcGtpLmRlbG9pdHRlLmNvbS9DZXJ0ZW5yb2xsL0Rl
# bG9pdHRlJTIwU0hBMiUyMExldmVsJTIwMyUyMENBJTIwMi5jcmwwggFUBggrBgEF
# BQcBAQSCAUYwggFCMIHEBggrBgEFBQcwAoaBt2xkYXA6Ly8vQ049RGVsb2l0dGUl
# MjBTSEEyJTIwTGV2ZWwlMjAzJTIwQ0ElMjAyLENOPUFJQSxDTj1QdWJsaWMlMjBL
# ZXklMjBTZXJ2aWNlcyxDTj1TZXJ2aWNlcyxDTj1Db25maWd1cmF0aW9uLERDPWRl
# bG9pdHRlLERDPWNvbT9jQUNlcnRpZmljYXRlP2Jhc2U/b2JqZWN0Q2xhc3M9Y2Vy
# dGlmaWNhdGlvbkF1dGhvcml0eTB5BggrBgEFBQcwAoZtaHR0cDovL3BraS5kZWxv
# aXR0ZS5jb20vQ2VydGVucm9sbC91c2F0cmFtZWVtMDA0LmF0cmFtZS5kZWxvaXR0
# ZS5jb21fRGVsb2l0dGUlMjBTSEEyJTIwTGV2ZWwlMjAzJTIwQ0ElMjAyLmNydDAN
# BgkqhkiG9w0BAQsFAAOCAQEAqFnnDf3WnhUtTZO7fhCSm1vcLN5H7xh55Fhsrapj
# Ku0aCSvHgWlZ9xlH2DboVFoMd589lU6DQujvfcTTpqY9zQu97QdszH8Wfhk9mW2O
# vVA3hDjahCEt+2vahw3aqsoSZaPYAjaRAMmeq23olHjMnFXvYntZImHjJjcSUpe+
# KkWxpdMd9rgKRUj86EQ0CluNC3ro3yrai/IUiDqboZ0lvI7GZYDnNzJMZHI3CtTn
# eDvfgtMY+xU+5ra53hbp93TYgr32bktk7p3Qp2kENBLYV/D59CghE4wxJW0pZ/Sw
# VXaJx3xzOjeO6PyAZ8vQieiBaf+2IDHXIh62x8UqlT1RDDCCBskwggSxoAMCAQIC
# EzQAAAAFqIzfrA2XWTIAAAAAAAUwDQYJKoZIhvcNAQELBQAwUjETMBEGCgmSJomT
# 8ixkARkWA2NvbTEYMBYGCgmSJomT8ixkARkWCERlbG9pdHRlMSEwHwYDVQQDExhE
# ZWxvaXR0ZSBTSEEyIExldmVsIDEgQ0EwHhcNMTUxMDI5MTcyMTAzWhcNMjUxMDI5
# MTczMTAzWjBUMRMwEQYKCZImiZPyLGQBGRYDY29tMRgwFgYKCZImiZPyLGQBGRYI
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
# P/0f31d3CndzQORYAb8CAwEAAaOCAZQwggGQMBAGCSsGAQQBgjcVAQQDAgEBMCMG
# CSsGAQQBgjcVAgQWBBRF4tTkKKaihh8hZlZ2wn5W1acT+zAdBgNVHQ4EFgQURy42
# 7rSc/1xeGHy4E+G+qSAeFLMwEQYDVR0gBAowCDAGBgRVHSAAMBkGCSsGAQQBgjcU
# AgQMHgoAUwB1AGIAQwBBMAsGA1UdDwQEAwIBhjASBgNVHRMBAf8ECDAGAQH/AgEB
# MB8GA1UdIwQYMBaAFL6LoCtmWVn6kHFRpaeoBkJOkDztMFgGA1UdHwRRME8wTaBL
# oEmGR2h0dHA6Ly9wa2kuZGVsb2l0dGUuY29tL0NlcnRFbnJvbGwvRGVsb2l0dGUl
# MjBTSEEyJTIwTGV2ZWwlMjAxJTIwQ0EuY3JsMG4GCCsGAQUFBwEBBGIwYDBeBggr
# BgEFBQcwAoZSaHR0cDovL3BraS5kZWxvaXR0ZS5jb20vQ2VydEVucm9sbC9TSEEy
# TFZMMUNBX0RlbG9pdHRlJTIwU0hBMiUyMExldmVsJTIwMSUyMENBLmNydDANBgkq
# hkiG9w0BAQsFAAOCAgEAUIIxw2cOQAxpWz1ZyL6PUsJPtdtzaxKmz4Tsw48uWk/l
# TbmWm7bD0WbFIlWwZ5DREGa9G99F0L3f+CO8Bqn+T6Jcw6xQ6Po53cXG4NSgoL6V
# v6CIfKVg9UwgcIj4J49sjTgiY7pn+wav9EKXM99AxPpNqxjLhRvTBk6Mbdg2ifED
# ljdc12PBWrHOE1M72cngFDkdRNboPpLylH8wUC3PojELdMIWC80//HOqLFsM07FM
# JaHHLB95oDuP+7+B0Q8n22MQVKyPihVAVDE6rhiAI7b2dt0C5vweubo0bTTIWhBA
# x5RO6b7/J1shCGb33HBxoAqX40i6AHaX6t+hapLCwYn1jGI0Ba57U0MeoLTrg77O
# KdqxwaJRauS8pORzZIJMEcJztATZaFf9cTKm8rD7EcvEfJib0I/ydR6chS55RWgD
# h8GlPoikRKW8xIomoA/iCKYMrroq5E6rY3ChgoYb3OwvtiTNpYKLsCVjRn4KieEm
# x4wl4h77RFywMjnGISoj56wrrk4jePpxjfiTHQVGx/6nQYx22IYPkMTEcMqVtT0Y
# Omd0rISvbwdSbyuozw923cC3lF86FoZAz1F5muSdE2VeejZYe7eYBxOeHHKk+/LA
# 3La7TCE/j/wWzN31mpOgQq62ct+HdG9o7EX/ITmwN7EDM4Aa4oMZytupX8iO61gx
# ggS6MIIEtgIBATCBgzBsMRMwEQYKCZImiZPyLGQBGRYDY29tMRgwFgYKCZImiZPy
# LGQBGRYIZGVsb2l0dGUxFjAUBgoJkiaJk/IsZAEZFgZhdHJhbWUxIzAhBgNVBAMT
# GkRlbG9pdHRlIFNIQTIgTGV2ZWwgMyBDQSAyAhNlADwUlwDjJaP1Xv9CAAAAPBSX
# MA0GCWCGSAFlAwQCAQUAoEwwGQYJKoZIhvcNAQkDMQwGCisGAQQBgjcCAQQwLwYJ
# KoZIhvcNAQkEMSIEIKyxNaaki575fkaVWAM1ZtbpNrDufCfjsOo3LIjzUcfQMA0G
# CSqGSIb3DQEBAQUABIIBAL1bJOWlc/763lzFWBzXgQ0+83ZLgIWrC4E6BC21PNKi
# EJJFQ5vXxPlUKCJU1f+K+FbSNuzoKnkyWoXAQ4EdNB9EMfmteDiAUyw02mciZNyk
# kZIh3GhHA8lkAdlpUcibFITTWUi7yAwNDq4mDNVadHB34JPhn8YnMxHPxtnBswGa
# yRQPwQYASh0bmZM9uD0Lc+EaB7PPjkXByGj1Kg5s6RsfPoy1oUhP0jIhiVcqQf2O
# nfHchvAqZ3QCeCgBqzmmhPu0RsfvNbSpB3MG4g9mLTDI6QjTGJNj4yxwr7lw69fZ
# zG4kLHHqTioVGZOX8PHKX6VAzc7xxhIm511qfMmQd7KhggK5MIICtQYJKoZIhvcN
# AQkGMYICpjCCAqICAQEwazBbMQswCQYDVQQGEwJCRTEZMBcGA1UEChMQR2xvYmFs
# U2lnbiBudi1zYTExMC8GA1UEAxMoR2xvYmFsU2lnbiBUaW1lc3RhbXBpbmcgQ0Eg
# LSBTSEEyNTYgLSBHMgIMJFS4fx4UU603+qF4MA0GCWCGSAFlAwQCAQUAoIIBDDAY
# BgkqhkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0xOTExMTEx
# NTQ1NDJaMC8GCSqGSIb3DQEJBDEiBCCN5OuplE0YtUUcSMoNVpH6xonFcqTlmR6s
# 0v4FUi5AujCBoAYLKoZIhvcNAQkQAgwxgZAwgY0wgYowgYcEFD7HZtXU1HLiGx8h
# Q1IcMbeQ2UtoMG8wX6RdMFsxCzAJBgNVBAYTAkJFMRkwFwYDVQQKExBHbG9iYWxT
# aWduIG52LXNhMTEwLwYDVQQDEyhHbG9iYWxTaWduIFRpbWVzdGFtcGluZyBDQSAt
# IFNIQTI1NiAtIEcyAgwkVLh/HhRTrTf6oXgwDQYJKoZIhvcNAQEBBQAEggEAKKOk
# HQ9TBXpc0oAbLtRhzTJ1gOWbpEd2+edkNwabTYxaCqL5upxdYX2IugPACqXI4SNt
# jL4wH83zeSLduyxDPbmYtTYQex6CJxSQDiscvW584cHX7fQ3LrBevuux+3oAof6b
# ipA+znDqCCJAfGDBOgNAWj9ETwaL99eMAkUvYtq+vmm9yvMEmKL2XKZMLvM6Qp7Z
# GkvxBtmJi1qy3DrY0O+Hed2d1kaKBFiaF8ygz8lApmPNNZGvVRYV3jhl+vdUSPlH
# zqdM4to/HvTI4wmpuvitMoCaLACuNPo2qLTrFQPA6iLrNE9sxQ/g2E0kCb7G6sOf
# sIZSYYLgaqcfvsbCUg==
# SIG # End signature block
