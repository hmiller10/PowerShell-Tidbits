﻿<#

.NOTES
THIS CODE IS MADE AVAILABLE AS IS, WITHOUT WARRANTY OF ANY KIND. THE
RISK OF THE USE OR THE RESULTS FROM THE USE OF THIS CODE REMAINS	WITH
THE USER.

.SYNOPSIS
    Clean expired certificates from CA database within the defined time period.

.DESCRIPTION
    This script will connect to the Certification Authority server defined in the
	script to locate and remove any expired certificate within the last year,
    except those issued by templates that are filtered out, from the CA database, 
	and if specified will ignore certificates issued from EFS, sMIME and key 
	recovery agent templates. This script will not compact or cleanup 
	white space in the database.

.EXAMPLE 
    PS> Remove-ExpiredCerts-Automated.ps1
	
.LINK
    https://www.sysadmins.lv/blog-en/categoryview/powershellpowershellpkimodule.aspx

.LINK
    https://github.com/Crypt32/PSPKI

#>


###########################################################################
#
#
# AUTHOR:  
#	Heather Miller
#
#
# VERSION HISTORY:
# 	2.0 10/11/2023 - Revised search syntax
#
# 
###########################################################################

[CmdletBinding()]
Param (
	[Parameter(Mandatory = $false, HelpMessage = "Use this switch to filter out certificates issued by specified certificate templates.",
			Position = 0)]
	[Switch]$ApplyFilters
)

#Modules
try
{
	Import-Module -Name PSPKI -Force -ErrorAction Stop
}
catch
{
	try
	{
		$moduleName = 'PSPKI'
		$ErrorActionPreference = 'Stop';
		$module = Get-Module -ListAvailable -Name $moduleName;
		$ErrorActionPreference = 'Continue';
		$modulePath = Split-Path $module.Path;
		$psdPath = "{0}\{1}" -f $modulePath, "PSPKI.psd1"
		Import-Module $psdPath -ErrorAction Stop
	}
	catch
	{
		Write-Error "PSPKI PS module could not be loaded. $($_.Exception.Message)" -ErrorAction Stop
	}
}

#Variables
[String]$sMIME = 'S/MIME'
[String]$efs = 'EFS'
[String]$Recovery = 'Recovery'
$StartDate = [datetime]::UtcNow.AddDays(-7)
$EndDate = [datetime]::UtcNow
$CA = [System.Net.Dns]::GetHostByName("LocalHost").HostName



#Script
$Error.Clear()

try
{
	Connect-CertificationAuthority -ComputerName $CA -ErrorAction Stop
}
catch
{
	$errorMessage = "{0}: {1}" -f $Error[0], $Error[0].InvocationInfo.PositionMessage
	Write-Error $errorMessage -ErrorAction Stop
}

If ( $PSBoundParameters.ContainsKey("ApplyFilters") )
{
	If ( ( Get-Module -Name PSPKI).Version -ge 3.4 )
	{
		try
		{
			$CA | Get-IssuedRequest -Filter "RequestID -gt $LastID", "NotAfter -ge $StartDate", "NotAfter -le $EndDate" -Properties $certProps -Page $pageNumber -PageSize $pageSize -ErrorAction Stop | `
			Where-Object { ((($_.CertificateTemplateOid.FriendlyName) -notlike "*$Recovery*") -and (($_.CertificateTemplateOid.FriendlyName) -notlike "*$efs*") -and (($_.CertificateTemplateOid.FriendlyName) -notlike "*$sMime*")) } | `
			Remove-AdcsDatabaseRow
		}
		catch
		{
			$errorMessage = "{0}: {1}" -f $Error[0], $Error[0].InvocationInfo.PositionMessage
			Write-Error $errorMessage -ErrorAction Continue
		}
	}
	Else
	{
		try
		{
			$CA | Get-IssuedRequest -Filter "RequestID -gt $LastID", "NotAfter -ge $StartDate", "NotAfter -le $EndDate" -Properties $certProps -Page $pageNumber -PageSize $pageSize -ErrorAction Stop | `
			Where-Object { ((($_.CertificateTemplateOid.FriendlyName) -notlike "*$Recovery*") -and (($_.CertificateTemplateOid.FriendlyName) -notlike "*$efs*") -and (($_.CertificateTemplateOid.FriendlyName) -notlike "*$sMime*")) } | `
			Remove-DatabaseRow
		}
		catch
		{
			$errorMessage = "{0}: {1}" -f $Error[0], $Error[0].InvocationInfo.PositionMessage
			Write-Error $errorMessage -ErrorAction Continue
		}
	}
}
Else
{
	If ( ( Get-Module -Name PSPKI).Version -ge 3.4 )
	{
		try
		{
			$CA | Get-IssuedRequest -Filter "NotAfter -ge $StartDate", "NotAfter -le $EndDate" -ErrorAction Stop | Remove-AdcsDatabaseRow
		}
		catch
		{
			$errorMessage = "{0}: {1}" -f $Error[0], $Error[0].InvocationInfo.PositionMessage
			Write-Error $errorMessage -ErrorAction Continue
		}

	}
	Else
	{
		try
		{
			Get-IssuedRequest -CertificationAuthority $CA -Filter "NotAfter -ge $StartDate", "NotAfter -le $EndDate" | Remove-DatabaseRow
		}
		catch
		{
			$errorMessage = "{0}: {1}" -f $Error[0], $Error[0].InvocationInfo.PositionMessage
			Write-Error $errorMessage -ErrorAction Continue
		}
	}
}

#End script


# SIG # Begin signature block
# MIIx1AYJKoZIhvcNAQcCoIIxxTCCMcECAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCAs8E/5XjSdqLvJ
# 4n1CZa8ANmUe9fyPvSi4txRUccOLZ6CCLAkwggV/MIIDZ6ADAgECAhAYtcKEQ5AS
# l0GsCYozZaYQMA0GCSqGSIb3DQEBCwUAMFIxEzARBgoJkiaJk/IsZAEZFgNjb20x
# GDAWBgoJkiaJk/IsZAEZFghEZWxvaXR0ZTEhMB8GA1UEAxMYRGVsb2l0dGUgU0hB
# MiBMZXZlbCAxIENBMB4XDTE1MDkwMTE1MDcyNVoXDTM1MDkwMTE1MDcyNVowUjET
# MBEGCgmSJomT8ixkARkWA2NvbTEYMBYGCgmSJomT8ixkARkWCERlbG9pdHRlMSEw
# HwYDVQQDExhEZWxvaXR0ZSBTSEEyIExldmVsIDEgQ0EwggIiMA0GCSqGSIb3DQEB
# AQUAA4ICDwAwggIKAoICAQCU+o2qpWkTjV2nWzX6d4z5e/nN9QApOsPXREAX16kU
# WYggwfrZWAxc5hhYGvI1BpQBg9mW+/9O3RwIpxrlcBYqngNsF5uUKbF8eyoTPdH+
# TOf8IdEedDdgxlEyisBxyrzYN3EqLCfD2jRblIYPkD7M1eH0ONwLHQblE4BqqK/u
# bcdjPYesS+p0i4yQyiPtjaJ4yLD8+4iNVTzCah2W2QGYagB445xVhpYFlOkrLQ0L
# /Fgvt4d8wp2BFprfSkVV5mWI3wMyI397Ft+vqK7TR9ACw3GJmV4f/oIse44N2H7s
# orTRanZaL2rP8h44AaSPyPSNcWedfk8dBZxVW/wTywgrVL3nEPGaE6yswaQYDobc
# XtrWp8jVae29gFP3x5SBlDfCEOqPKmPrbaONcuRTGV+5R64EcHb5P2RJtqlAciGM
# ATd/3sD8U67MbRp6uvp40Ll2g2M5ffUF+nlbRaSFtfwPQB7F/u/46xuioWQCRo3N
# JLahpfv4bdcDDav6yUObKKav5uhosSji8z6gXMVR4FA5e+MYESca3FeZIg/Cwu74
# eWE6+O1cg0tuxzOXZBbMLZqtE/HHocuvWE+PR6LLPtsoM9D435T2Xk560vo00WCm
# lUOipJWjkVTCHjvBkpqpdnxvWM+P9JL4hGxsqcInfv5hd+QDYyJ33HygzJF8F/B5
# 3wIDAQABo1EwTzALBgNVHQ8EBAMCAYYwDwYDVR0TAQH/BAUwAwEB/zAdBgNVHQ4E
# FgQUvougK2ZZWfqQcVGlp6gGQk6QPO0wEAYJKwYBBAGCNxUBBAMCAQAwDQYJKoZI
# hvcNAQELBQADggIBAINO9HIsbvrfvlRKl7Ep/QEYK5h9b6j/obXwZGpl5iWzuU7b
# nXGfuF9dpx8RkgcI+iLNqfTK9km1hShxdEyJ6jvrpnCfyDg6ARmMlezDzWnQaWTK
# WZ1aGo76xG4bJi8ZjfYxZnXqjPrH8Ib6ux6i7Yewsu2VJXvoPZc+cO60gAcj9LiK
# zgDsXRX6g1fThoelRcxvRjUGQ8o0XWAiWZ/it21GHr7avfvhA/5G7D6cmI83AXtM
# XbzCTPWJRdYr0VKYzVDr7+5fGnH2pEVHC/6kJ3/Tsid5ncabHFM0nFLPMQnEh/Cg
# kifocv01g+W0BnsK5ZiyblMVj7HlNnmpL87r6cAd3fRPsr+r4fmGOn22KxhqHdGI
# E45T6Jm+1SyvVdotPvOLPGeaniywcevIuE2Ri7a6H91PX+KrxJszd8oYSUwv6YTt
# k8CDarYHvqSBtj4asS3M2fPA18jCSH8gMpgn2folTbJDihrmwmNN/m4dSHFX1l3z
# 8FkrwvcflLsjLfIJUvKr3zG5RncKEvgWgp5Lz24yuAoAM/bGcB0jDWed3rtlI+tN
# WTexk7+4lUpPOOlTRXXSxzbtWZXH1gHJzijjEOhDr54CVRs2m6gWlyHN0vZ3HklR
# qUuXZPiODhQjMpyJ0+ESbJTUCkpRcVBNvHQ3y/VYLZRdE7C5zRgMtHu/pEIuMIIF
# jTCCBHWgAwIBAgIQDpsYjvnQLefv21DiCEAYWjANBgkqhkiG9w0BAQwFADBlMQsw
# CQYDVQQGEwJVUzEVMBMGA1UEChMMRGlnaUNlcnQgSW5jMRkwFwYDVQQLExB3d3cu
# ZGlnaWNlcnQuY29tMSQwIgYDVQQDExtEaWdpQ2VydCBBc3N1cmVkIElEIFJvb3Qg
# Q0EwHhcNMjIwODAxMDAwMDAwWhcNMzExMTA5MjM1OTU5WjBiMQswCQYDVQQGEwJV
# UzEVMBMGA1UEChMMRGlnaUNlcnQgSW5jMRkwFwYDVQQLExB3d3cuZGlnaWNlcnQu
# Y29tMSEwHwYDVQQDExhEaWdpQ2VydCBUcnVzdGVkIFJvb3QgRzQwggIiMA0GCSqG
# SIb3DQEBAQUAA4ICDwAwggIKAoICAQC/5pBzaN675F1KPDAiMGkz7MKnJS7JIT3y
# ithZwuEppz1Yq3aaza57G4QNxDAf8xukOBbrVsaXbR2rsnnyyhHS5F/WBTxSD1If
# xp4VpX6+n6lXFllVcq9ok3DCsrp1mWpzMpTREEQQLt+C8weE5nQ7bXHiLQwb7iDV
# ySAdYyktzuxeTsiT+CFhmzTrBcZe7FsavOvJz82sNEBfsXpm7nfISKhmV1efVFiO
# DCu3T6cw2Vbuyntd463JT17lNecxy9qTXtyOj4DatpGYQJB5w3jHtrHEtWoYOAMQ
# jdjUN6QuBX2I9YI+EJFwq1WCQTLX2wRzKm6RAXwhTNS8rhsDdV14Ztk6MUSaM0C/
# CNdaSaTC5qmgZ92kJ7yhTzm1EVgX9yRcRo9k98FpiHaYdj1ZXUJ2h4mXaXpI8OCi
# EhtmmnTK3kse5w5jrubU75KSOp493ADkRSWJtppEGSt+wJS00mFt6zPZxd9LBADM
# fRyVw4/3IbKyEbe7f/LVjHAsQWCqsWMYRJUadmJ+9oCw++hkpjPRiQfhvbfmQ6QY
# uKZ3AeEPlAwhHbJUKSWJbOUOUlFHdL4mrLZBdd56rF+NP8m800ERElvlEFDrMcXK
# chYiCd98THU/Y+whX8QgUWtvsauGi0/C1kVfnSD8oR7FwI+isX4KJpn15GkvmB0t
# 9dmpsh3lGwIDAQABo4IBOjCCATYwDwYDVR0TAQH/BAUwAwEB/zAdBgNVHQ4EFgQU
# 7NfjgtJxXWRM3y5nP+e6mK4cD08wHwYDVR0jBBgwFoAUReuir/SSy4IxLVGLp6ch
# nfNtyA8wDgYDVR0PAQH/BAQDAgGGMHkGCCsGAQUFBwEBBG0wazAkBggrBgEFBQcw
# AYYYaHR0cDovL29jc3AuZGlnaWNlcnQuY29tMEMGCCsGAQUFBzAChjdodHRwOi8v
# Y2FjZXJ0cy5kaWdpY2VydC5jb20vRGlnaUNlcnRBc3N1cmVkSURSb290Q0EuY3J0
# MEUGA1UdHwQ+MDwwOqA4oDaGNGh0dHA6Ly9jcmwzLmRpZ2ljZXJ0LmNvbS9EaWdp
# Q2VydEFzc3VyZWRJRFJvb3RDQS5jcmwwEQYDVR0gBAowCDAGBgRVHSAAMA0GCSqG
# SIb3DQEBDAUAA4IBAQBwoL9DXFXnOF+go3QbPbYW1/e/Vwe9mqyhhyzshV6pGrsi
# +IcaaVQi7aSId229GhT0E0p6Ly23OO/0/4C5+KH38nLeJLxSA8hO0Cre+i1Wz/n0
# 96wwepqLsl7Uz9FDRJtDIeuWcqFItJnLnU+nBgMTdydE1Od/6Fmo8L8vC6bp8jQ8
# 7PcDx4eo0kxAGTVGamlUsLihVo7spNU96LHc/RzY9HdaXFSMb++hUD38dglohJ9v
# ytsgjTVgHAIDyyCwrFigDkBjxZgiwbJZ9VVrzyerbHbObyMt9H5xaiNrIv8SuFQt
# J37YOtnwtoeW/VvRXKwYw02fc7cBqZ9Xql4o4rmUMIIF3jCCA8agAwIBAgITPgAA
# AAp01W3Jvy6VAgACAAAACjANBgkqhkiG9w0BAQsFADBUMRMwEQYKCZImiZPyLGQB
# GRYDY29tMRgwFgYKCZImiZPyLGQBGRYIRGVsb2l0dGUxIzAhBgNVBAMTGkRlbG9p
# dHRlIFNIQTIgTGV2ZWwgMiBDQSAyMB4XDTIxMDYyOTE5MzUwMVoXDTI2MDYyOTE5
# NDUwMVowbDETMBEGCgmSJomT8ixkARkWA2NvbTEYMBYGCgmSJomT8ixkARkWCGRl
# bG9pdHRlMRYwFAYKCZImiZPyLGQBGRYGYXRyYW1lMSMwIQYDVQQDExpEZWxvaXR0
# ZSBTSEEyIExldmVsIDMgQ0EgMjCCASIwDQYJKoZIhvcNAQEBBQADggEPADCCAQoC
# ggEBAKMErt43dTBJZGDAXh2pNncSX8kiKriuGlr08U/3ZtI1k6YKEtUpns34twsN
# 6Rwmuq8Z2FTxlfFlCKHitdWkr6ES+gC/uh0MAPix1XZmErACC2j2rVDX1ELXzwtd
# zCIrzpBaXXxD+lCw0eou0CEnSQXAfYLEZ3+Eoj6HjejDLwuBAhTisC4mwEyIoTVU
# sgkZns4l3X0rXyvZfsxN7lGLV9wIDzP73qAl+AJ6W3vShFbNb7Gzzhln5qvho/y5
# 542rzi+SwcAtCLbmL+nrxSyNjc+p1w3qHV+ZmknT7Vtz30738mln8F9ne0ZvWo9M
# Ba9Mtu3H/FmFcyW/m9hlsYnV0pECAwEAAaOCAY8wggGLMBIGCSsGAQQBgjcVAQQF
# AgMCAAIwIwYJKwYBBAGCNxUCBBYEFLy+iTDVUCgkot0nGQ6PZed3WpwRMB0GA1Ud
# DgQWBBQ4oakuFXDiR2EUBm025muPDi7DYTAZBgkrBgEEAYI3FAIEDB4KAFMAdQBi
# AEMAQTALBgNVHQ8EBAMCAYYwEgYDVR0TAQH/BAgwBgEB/wIBADAfBgNVHSMEGDAW
# gBRHLjbutJz/XF4YfLgT4b6pIB4UszBcBgNVHR8EVTBTMFGgT6BNhktodHRwOi8v
# cGtpLmRlbG9pdHRlLmNvbS9DZXJ0RW5yb2xsL0RlbG9pdHRlJTIwU0hBMiUyMExl
# dmVsJTIwMiUyMENBJTIwMi5jcmwwdgYIKwYBBQUHAQEEajBoMGYGCCsGAQUFBzAC
# hlpodHRwOi8vcGtpLmRlbG9pdHRlLmNvbS9DZXJ0RW5yb2xsL1NIQTJMVkwyQ0Ey
# X0RlbG9pdHRlJTIwU0hBMiUyMExldmVsJTIwMiUyMENBJTIwMigyKS5jcnQwDQYJ
# KoZIhvcNAQELBQADggIBAGb01UEleTvuxAzRCWf39e2Ksfpsk8hfLVoWHKvXJ1M6
# 5ndOqA5ZjFlhd3muKeyRastbEQ16n1RV760y70Npp2L8Zmp/u0FmlvdzTtnWcc4m
# ny9FO0hFOHShoDy+ZGvKdsikSnod01D0dc5OCHGUEMse3xJvOobzXy02yVlo98Ec
# AuyPgWP21LbSOPAPU9OJPtNmbBSi9Tgcazl+204X+FpGrT+eBlh5p4sR5hSY2HYo
# ZYplGGvT5OABwS3U/eMXw3oSHgnMwtj6MmUJH/M/RZaeyxPsETZ9itakLVI1JnYb
# wJuR6DdlXQDgQ5KKulVHT3LDRbf/+GmJn56dGk2kWUQzsbqYVfWB6WY6JDndX3eL
# jeKE+7ukWmSE4rCHk0h9M9waCWaZjnuTqEqDOim91L/UoFJJ4KnpPDrGe5dZj6FX
# VdOVPZb8AO+ZP0QgjyxSwssAQpfUJbUFJI2Y91Qz7dUEyWyunHI1g/CPVUWWL1UV
# +VqwXq7C0d8RAdi21aFRaMS7leHS7zzPIEduMQgEIrvClZ0rRwuMjH2TeiC9t4Qb
# NEJFuLCFwLTC3sTi2p4tWd4nVUv5oQtYSqMUhub4p+yXIoX5je0Oqb5s+T5cdamI
# sw4WI4S11ZlosGZ/XryOU8MdrGIyIBoL/CGhDHWvmlRM7rOQ4aRfvw4+IOvBRTSE
# MIIGrjCCBJagAwIBAgIQBzY3tyRUfNhHrP0oZipeWzANBgkqhkiG9w0BAQsFADBi
# MQswCQYDVQQGEwJVUzEVMBMGA1UEChMMRGlnaUNlcnQgSW5jMRkwFwYDVQQLExB3
# d3cuZGlnaWNlcnQuY29tMSEwHwYDVQQDExhEaWdpQ2VydCBUcnVzdGVkIFJvb3Qg
# RzQwHhcNMjIwMzIzMDAwMDAwWhcNMzcwMzIyMjM1OTU5WjBjMQswCQYDVQQGEwJV
# UzEXMBUGA1UEChMORGlnaUNlcnQsIEluYy4xOzA5BgNVBAMTMkRpZ2lDZXJ0IFRy
# dXN0ZWQgRzQgUlNBNDA5NiBTSEEyNTYgVGltZVN0YW1waW5nIENBMIICIjANBgkq
# hkiG9w0BAQEFAAOCAg8AMIICCgKCAgEAxoY1BkmzwT1ySVFVxyUDxPKRN6mXUaHW
# 0oPRnkyibaCwzIP5WvYRoUQVQl+kiPNo+n3znIkLf50fng8zH1ATCyZzlm34V6gC
# ff1DtITaEfFzsbPuK4CEiiIY3+vaPcQXf6sZKz5C3GeO6lE98NZW1OcoLevTsbV1
# 5x8GZY2UKdPZ7Gnf2ZCHRgB720RBidx8ald68Dd5n12sy+iEZLRS8nZH92GDGd1f
# tFQLIWhuNyG7QKxfst5Kfc71ORJn7w6lY2zkpsUdzTYNXNXmG6jBZHRAp8ByxbpO
# H7G1WE15/tePc5OsLDnipUjW8LAxE6lXKZYnLvWHpo9OdhVVJnCYJn+gGkcgQ+ND
# Y4B7dW4nJZCYOjgRs/b2nuY7W+yB3iIU2YIqx5K/oN7jPqJz+ucfWmyU8lKVEStY
# dEAoq3NDzt9KoRxrOMUp88qqlnNCaJ+2RrOdOqPVA+C/8KI8ykLcGEh/FDTP0kyr
# 75s9/g64ZCr6dSgkQe1CvwWcZklSUPRR8zZJTYsg0ixXNXkrqPNFYLwjjVj33GHe
# k/45wPmyMKVM1+mYSlg+0wOI/rOP015LdhJRk8mMDDtbiiKowSYI+RQQEgN9XyO7
# ZONj4KbhPvbCdLI/Hgl27KtdRnXiYKNYCQEoAA6EVO7O6V3IXjASvUaetdN2udIO
# a5kM0jO0zbECAwEAAaOCAV0wggFZMBIGA1UdEwEB/wQIMAYBAf8CAQAwHQYDVR0O
# BBYEFLoW2W1NhS9zKXaaL3WMaiCPnshvMB8GA1UdIwQYMBaAFOzX44LScV1kTN8u
# Zz/nupiuHA9PMA4GA1UdDwEB/wQEAwIBhjATBgNVHSUEDDAKBggrBgEFBQcDCDB3
# BggrBgEFBQcBAQRrMGkwJAYIKwYBBQUHMAGGGGh0dHA6Ly9vY3NwLmRpZ2ljZXJ0
# LmNvbTBBBggrBgEFBQcwAoY1aHR0cDovL2NhY2VydHMuZGlnaWNlcnQuY29tL0Rp
# Z2lDZXJ0VHJ1c3RlZFJvb3RHNC5jcnQwQwYDVR0fBDwwOjA4oDagNIYyaHR0cDov
# L2NybDMuZGlnaWNlcnQuY29tL0RpZ2lDZXJ0VHJ1c3RlZFJvb3RHNC5jcmwwIAYD
# VR0gBBkwFzAIBgZngQwBBAIwCwYJYIZIAYb9bAcBMA0GCSqGSIb3DQEBCwUAA4IC
# AQB9WY7Ak7ZvmKlEIgF+ZtbYIULhsBguEE0TzzBTzr8Y+8dQXeJLKftwig2qKWn8
# acHPHQfpPmDI2AvlXFvXbYf6hCAlNDFnzbYSlm/EUExiHQwIgqgWvalWzxVzjQEi
# Jc6VaT9Hd/tydBTX/6tPiix6q4XNQ1/tYLaqT5Fmniye4Iqs5f2MvGQmh2ySvZ18
# 0HAKfO+ovHVPulr3qRCyXen/KFSJ8NWKcXZl2szwcqMj+sAngkSumScbqyQeJsG3
# 3irr9p6xeZmBo1aGqwpFyd/EjaDnmPv7pp1yr8THwcFqcdnGE4AJxLafzYeHJLtP
# o0m5d2aR8XKc6UsCUqc3fpNTrDsdCEkPlM05et3/JWOZJyw9P2un8WbDQc1PtkCb
# ISFA0LcTJM3cHXg65J6t5TRxktcma+Q4c6umAU+9Pzt4rUyt+8SVe+0KXzM5h0F4
# ejjpnOHdI/0dKNPH+ejxmF/7K9h+8kaddSweJywm228Vex4Ziza4k9Tm8heZWcpw
# 8De/mADfIBZPJ/tgZxahZrrdVcA6KYawmKAr7ZVBtzrVFZgxtGIJDwq9gdkT/r+k
# 0fNX2bwE+oLeMt8EifAAzV3C+dAjfwAL5HYCJtnwZXZCpimHCUcr5n8apIUP/JiW
# 9lVUKx+A+sDyDivl1vupL0QVSucTDh3bNzgaoSv27dZ8/DCCBsIwggSqoAMCAQIC
# EAVEr/OUnQg5pr/bP1/lYRYwDQYJKoZIhvcNAQELBQAwYzELMAkGA1UEBhMCVVMx
# FzAVBgNVBAoTDkRpZ2lDZXJ0LCBJbmMuMTswOQYDVQQDEzJEaWdpQ2VydCBUcnVz
# dGVkIEc0IFJTQTQwOTYgU0hBMjU2IFRpbWVTdGFtcGluZyBDQTAeFw0yMzA3MTQw
# MDAwMDBaFw0zNDEwMTMyMzU5NTlaMEgxCzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5E
# aWdpQ2VydCwgSW5jLjEgMB4GA1UEAxMXRGlnaUNlcnQgVGltZXN0YW1wIDIwMjMw
# ggIiMA0GCSqGSIb3DQEBAQUAA4ICDwAwggIKAoICAQCjU0WHHYOOW6w+VLMj4M+f
# 1+XS512hDgncL0ijl3o7Kpxn3GIVWMGpkxGnzaqyat0QKYoeYmNp01icNXG/Opfr
# lFCPHCDqx5o7L5Zm42nnaf5bw9YrIBzBl5S0pVCB8s/LB6YwaMqDQtr8fwkklKSC
# Gtpqutg7yl3eGRiF+0XqDWFsnf5xXsQGmjzwxS55DxtmUuPI1j5f2kPThPXQx/ZI
# LV5FdZZ1/t0QoRuDwbjmUpW1R9d4KTlr4HhZl+NEK0rVlc7vCBfqgmRN/yPjyobu
# tKQhZHDr1eWg2mOzLukF7qr2JPUdvJscsrdf3/Dudn0xmWVHVZ1KJC+sK5e+n+T9
# e3M+Mu5SNPvUu+vUoCw0m+PebmQZBzcBkQ8ctVHNqkxmg4hoYru8QRt4GW3k2Q/g
# WEH72LEs4VGvtK0VBhTqYggT02kefGRNnQ/fztFejKqrUBXJs8q818Q7aESjpTtC
# /XN97t0K/3k0EH6mXApYTAA+hWl1x4Nk1nXNjxJ2VqUk+tfEayG66B80mC866msB
# sPf7Kobse1I4qZgJoXGybHGvPrhvltXhEBP+YUcKjP7wtsfVx95sJPC/QoLKoHE9
# nJKTBLRpcCcNT7e1NtHJXwikcKPsCvERLmTgyyIryvEoEyFJUX4GZtM7vvrrkTjY
# UQfKlLfiUKHzOtOKg8tAewIDAQABo4IBizCCAYcwDgYDVR0PAQH/BAQDAgeAMAwG
# A1UdEwEB/wQCMAAwFgYDVR0lAQH/BAwwCgYIKwYBBQUHAwgwIAYDVR0gBBkwFzAI
# BgZngQwBBAIwCwYJYIZIAYb9bAcBMB8GA1UdIwQYMBaAFLoW2W1NhS9zKXaaL3WM
# aiCPnshvMB0GA1UdDgQWBBSltu8T5+/N0GSh1VapZTGj3tXjSTBaBgNVHR8EUzBR
# ME+gTaBLhklodHRwOi8vY3JsMy5kaWdpY2VydC5jb20vRGlnaUNlcnRUcnVzdGVk
# RzRSU0E0MDk2U0hBMjU2VGltZVN0YW1waW5nQ0EuY3JsMIGQBggrBgEFBQcBAQSB
# gzCBgDAkBggrBgEFBQcwAYYYaHR0cDovL29jc3AuZGlnaWNlcnQuY29tMFgGCCsG
# AQUFBzAChkxodHRwOi8vY2FjZXJ0cy5kaWdpY2VydC5jb20vRGlnaUNlcnRUcnVz
# dGVkRzRSU0E0MDk2U0hBMjU2VGltZVN0YW1waW5nQ0EuY3J0MA0GCSqGSIb3DQEB
# CwUAA4ICAQCBGtbeoKm1mBe8cI1PijxonNgl/8ss5M3qXSKS7IwiAqm4z4Co2efj
# xe0mgopxLxjdTrbebNfhYJwr7e09SI64a7p8Xb3CYTdoSXej65CqEtcnhfOOHpLa
# wkA4n13IoC4leCWdKgV6hCmYtld5j9smViuw86e9NwzYmHZPVrlSwradOKmB521B
# XIxp0bkrxMZ7z5z6eOKTGnaiaXXTUOREEr4gDZ6pRND45Ul3CFohxbTPmJUaVLq5
# vMFpGbrPFvKDNzRusEEm3d5al08zjdSNd311RaGlWCZqA0Xe2VC1UIyvVr1MxeFG
# xSjTredDAHDezJieGYkD6tSRN+9NUvPJYCHEVkft2hFLjDLDiOZY4rbbPvlfsELW
# j+MXkdGqwFXjhr+sJyxB0JozSqg21Llyln6XeThIX8rC3D0y33XWNmdaifj2p8fl
# TzU8AL2+nCpseQHc2kTmOt44OwdeOVj0fHMxVaCAEcsUDH6uvP6k63llqmjWIso7
# 65qCNVcoFstp8jKastLYOrixRoZruhf9xHdsFWyuq69zOuhJRrfVf8y2OMDY7Bz1
# tqG4QyzfTkx9HmhwwHcK1ALgXGC7KP845VJa1qwXIiNO9OzTF/tQa/8Hdx9xl0RB
# ybhG02wyfFgvZ0dl5Rtztpn5aywGRu9BHvDwX+Db2a2QgESvgBBBijCCBskwggSx
# oAMCAQICEzQAAAAHiSF1iXPNJ/IAAAAAAAcwDQYJKoZIhvcNAQELBQAwUjETMBEG
# CgmSJomT8ixkARkWA2NvbTEYMBYGCgmSJomT8ixkARkWCERlbG9pdHRlMSEwHwYD
# VQQDExhEZWxvaXR0ZSBTSEEyIExldmVsIDEgQ0EwHhcNMjAwODA1MTczMjU2WhcN
# MzAwODA1MTc0MjU2WjBUMRMwEQYKCZImiZPyLGQBGRYDY29tMRgwFgYKCZImiZPy
# LGQBGRYIRGVsb2l0dGUxIzAhBgNVBAMTGkRlbG9pdHRlIFNIQTIgTGV2ZWwgMiBD
# QSAyMIICIjANBgkqhkiG9w0BAQEFAAOCAg8AMIICCgKCAgEAmPb6sHLB25JD286N
# fyR2RfuNgmSXaR2dLojx7rPDqiEWKM01mSdquzeXj7Qu/VQsiQLV/9oxwMArSvjJ
# HRjeQ2L7orPGytxWiO6nNHkKbPUCkBTmRALVcXK0iYmXhQjaypjx5y8bi3K13AR7
# axTbNlPEFy3z9TwFGftmeJOIvle3dBvOCxJre1mxmf544tkzq+Df0ENP8sA41WeQ
# bA5ZyDa2C8PWm8XL59X00UgtMJcOq4fCG+xkjl7nnbQ4/AP7lGHGkl0bnYE5Xd/n
# VA86+wO+uTUcmbs0fJ9fKO3bq3wgiUaRyyBbUQ2NzGlgaffxqge2lM3WCmiQeHKy
# fKsOkfg41+6h7qUFywDoDkvnVBjJs2+tgImqqD6iwmgZWHt6PeIiwJA/IIKBf0t1
# O16G39uim6NSiesSK+wfOMxyxZio/BzKGPOtv4PwosBlPKlhK5bbvMWY2RFsWQJ6
# LPiRXlE5NIYbh/CTyngIdM6Drwr57sIZGWbKCJc9nORteVgx3pgciFAxOFGn1k3z
# mxM83qYxxgKi6fql8KCgbo+l6luROLa5rsRfkGPtRXy1HWJ7xwcf8/JxLJGlp1rt
# nGnZljvb0Tbtwo8GwDoihSMSh9MoGrJTrtk8tnYf4UpLgGKjGyGOUBFGrRGQcEhW
# bzDTK5qZP/0f31d3CndzQORYAb8CAwEAAaOCAZQwggGQMBAGCSsGAQQBgjcVAQQD
# AgECMCMGCSsGAQQBgjcVAgQWBBQV4b/ii/DtWsyFdE+p2v+xjwi3MzAdBgNVHQ4E
# FgQURy427rSc/1xeGHy4E+G+qSAeFLMwEQYDVR0gBAowCDAGBgRVHSAAMBkGCSsG
# AQQBgjcUAgQMHgoAUwB1AGIAQwBBMAsGA1UdDwQEAwIBhjASBgNVHRMBAf8ECDAG
# AQH/AgEBMB8GA1UdIwQYMBaAFL6LoCtmWVn6kHFRpaeoBkJOkDztMFgGA1UdHwRR
# ME8wTaBLoEmGR2h0dHA6Ly9wa2kuZGVsb2l0dGUuY29tL0NlcnRFbnJvbGwvRGVs
# b2l0dGUlMjBTSEEyJTIwTGV2ZWwlMjAxJTIwQ0EuY3JsMG4GCCsGAQUFBwEBBGIw
# YDBeBggrBgEFBQcwAoZSaHR0cDovL3BraS5kZWxvaXR0ZS5jb20vQ2VydEVucm9s
# bC9TSEEyTFZMMUNBX0RlbG9pdHRlJTIwU0hBMiUyMExldmVsJTIwMSUyMENBLmNy
# dDANBgkqhkiG9w0BAQsFAAOCAgEAh56DUZ5xeRvT/JdAM8biqLOI4PLlIhIPufRf
# MPJlmo64dY9G/JVkF7qh8SHWm7umjSTvb357kayJCks5Y3VwS11A9HMsRK11083e
# xB27HUBd2W3IyRv2KBZT+SsAsnhtb2slEuPqqrpFZC3u2RZa8XonKVVcX3wfFN0q
# xE+yXkjYMUNxr3kYuclb2kt/4/RggkfV06dL0X2lHktLMYILmr8Tb2/eU2S7//hr
# dcH/tcWZ29hiIzL0qayp0j2MBuXACV/ZDNheEBvD659p14ae23CrgTpXSLL68RwH
# jaQqFVf2EWPXjR1MVJSvjB3QKiGdXTltUu1MBsrRHbFwj83xhiS1nTSWfIxSM+NG
# 0u+tj9SJ5fOQSEMlCe0achdoXPvF50uDwaTLxUOoBoDK2DKd8nJFa/x8/Gj35jn7
# RNp//UuzbmIhOr7YZqdfiBwnGffm4rS577EnBSsQhuOjzujrJbJd3NP2ar293Zup
# r8d+QYUbU51ny+mUYbGQ8VQeZgo72XkFAnzx1vZw9UK5VU7pC0zlBZL/FNV6hbgc
# nxQ0K/qRAudJtx03GpNF5sqhyEC0ndvCSdljKsf4mvgNwrEDTa6HLtEKONisnLSg
# 56IrcPx5W/eD8Ksodlwpfg8UM/A942V8JRZJLgFrc+nqysPi+cINMTd/n40h2wzG
# RlcjDE8wggbKMIIFsqADAgECAhNlAJU42fb9El3zjUwgAAIAlTjZMA0GCSqGSIb3
# DQEBCwUAMGwxEzARBgoJkiaJk/IsZAEZFgNjb20xGDAWBgoJkiaJk/IsZAEZFghk
# ZWxvaXR0ZTEWMBQGCgmSJomT8ixkARkWBmF0cmFtZTEjMCEGA1UEAxMaRGVsb2l0
# dGUgU0hBMiBMZXZlbCAzIENBIDIwHhcNMjMwNDI2MTgwNDUwWhcNMjUwNDI1MTgw
# NDUwWjBuMQswCQYDVQQGEwJVUzELMAkGA1UECBMCUEExEzARBgNVBAcTCkdsZW4g
# TWlsbHMxETAPBgNVBAoTCERlbG9pdHRlMQwwCgYDVQQLEwNEVFMxHDAaBgNVBAMT
# E0lBTUluZnJhQ29kZVNpZ25pbmcwggEiMA0GCSqGSIb3DQEBAQUAA4IBDwAwggEK
# AoIBAQCX+Dk7TU4HXEhaUyGdGsgMjg5yed6BbkA+YP+pydr3VA1gf78cO4wSTadA
# 27HzEfnQfAGI7BADObnRYgIRK9CCHMWyzkBnv7qEXKli3dbLd+30ZiJBxPdCwvN4
# 9CCsltsesxyWZUH0gHdHG05y2BGzwPHaOqaQuLHnpmTDXcCcHhgfIaIJfX0DUond
# q484hh2Fm6Ne0x3kcISftDK4mfZ8VaZWMrZ6dE5iDWlZXBeGZgLgs4CpBetTEkqc
# 8odMm8nxcwCQk8eGMmm2QvYlhcqNuv039yr2mb4iN5Cs9GQ4EAKEH3HgoPAarflM
# rcKSwKL9vHgAbwZtn/FsxZgEN06zAgMBAAGjggNhMIIDXTAdBgNVHQ4EFgQU09Ll
# 9rti/7M/YbA4pLtq/002ccIwHwYDVR0jBBgwFoAUOKGpLhVw4kdhFAZtNuZrjw4u
# w2EwggFBBgNVHR8EggE4MIIBNDCCATCgggEsoIIBKIaB1WxkYXA6Ly8vQ049RGVs
# b2l0dGUlMjBTSEEyJTIwTGV2ZWwlMjAzJTIwQ0ElMjAyKDIpLENOPXVzYXRyYW1l
# ZW0wMDQsQ049Q0RQLENOPVB1YmxpYyUyMEtleSUyMFNlcnZpY2VzLENOPVNlcnZp
# Y2VzLENOPUNvbmZpZ3VyYXRpb24sREM9ZGVsb2l0dGUsREM9Y29tP2NlcnRpZmlj
# YXRlUmV2b2NhdGlvbkxpc3Q/YmFzZT9vYmplY3RDbGFzcz1jUkxEaXN0cmlidXRp
# b25Qb2ludIZOaHR0cDovL3BraS5kZWxvaXR0ZS5jb20vQ2VydGVucm9sbC9EZWxv
# aXR0ZSUyMFNIQTIlMjBMZXZlbCUyMDMlMjBDQSUyMDIoMikuY3JsMIIBVwYIKwYB
# BQUHAQEEggFJMIIBRTCBxAYIKwYBBQUHMAKGgbdsZGFwOi8vL0NOPURlbG9pdHRl
# JTIwU0hBMiUyMExldmVsJTIwMyUyMENBJTIwMixDTj1BSUEsQ049UHVibGljJTIw
# S2V5JTIwU2VydmljZXMsQ049U2VydmljZXMsQ049Q29uZmlndXJhdGlvbixEQz1k
# ZWxvaXR0ZSxEQz1jb20/Y0FDZXJ0aWZpY2F0ZT9iYXNlP29iamVjdENsYXNzPWNl
# cnRpZmljYXRpb25BdXRob3JpdHkwfAYIKwYBBQUHMAKGcGh0dHA6Ly9wa2kuZGVs
# b2l0dGUuY29tL0NlcnRlbnJvbGwvdXNhdHJhbWVlbTAwNC5hdHJhbWUuZGVsb2l0
# dGUuY29tX0RlbG9pdHRlJTIwU0hBMiUyMExldmVsJTIwMyUyMENBJTIwMigyKS5j
# cnQwCwYDVR0PBAQDAgeAMDwGCSsGAQQBgjcVBwQvMC0GJSsGAQQBgjcVCIGBvUmF
# voUTgtWbPIPXjgeG8ckKXIPK9y3C8zICAWQCAR4wEwYDVR0lBAwwCgYIKwYBBQUH
# AwMwGwYJKwYBBAGCNxUKBA4wDDAKBggrBgEFBQcDAzANBgkqhkiG9w0BAQsFAAOC
# AQEAX94WDLBVQFBHMypTzJWGnOMCXwzvt6041xwGivYARE0aaJ4FaVe0DYcsNiFg
# ybImyVuCZ6y6vVnPX7bLmrcU9k8PPSGSGUbAQ9/gPjEspmR1nQBbPf9gE/YDIsJO
# NgW+6quY8qhKl+PBawBrlbbRa4v2JTmwNeC/cnrHPWVqg7Mk+gEBlL0k4HhqMqsX
# uUxomKZO04/3oJkMBQKXyBkab3JUCT7upI1iJ1g3c7hrXjt7dKcTm+zYijpoGMOZ
# URYmlBJ4Xm/TC/xyMH4pH0mmjn8UUu5bp2Duuxl9VwLE+rD2WqSycGtuL00tCvX4
# dmlZkZtpCUCrRGiyUrzxRp93wDGCBSEwggUdAgEBMIGDMGwxEzARBgoJkiaJk/Is
# ZAEZFgNjb20xGDAWBgoJkiaJk/IsZAEZFghkZWxvaXR0ZTEWMBQGCgmSJomT8ixk
# ARkWBmF0cmFtZTEjMCEGA1UEAxMaRGVsb2l0dGUgU0hBMiBMZXZlbCAzIENBIDIC
# E2UAlTjZ9v0SXfONTCAAAgCVONkwDQYJYIZIAWUDBAIBBQCgTDAZBgkqhkiG9w0B
# CQMxDAYKKwYBBAGCNwIBBDAvBgkqhkiG9w0BCQQxIgQgAANkXSDalfOxSKSUtYcP
# BdrZtyabSPQrUWdeJX5DgAAwDQYJKoZIhvcNAQEBBQAEggEACfoVTqqpVfMu+ooy
# BYGTf34by0u1pLe3v11BCB/HDui42MeE82uIBN1ZgWGawbeEddqzsujhDZDP5M15
# Z3xRcqkIOgUEV51J7NdG/02oyj+Q0m/GMIQT0w1WwDHtoO011HXziZsowM1fe83N
# 9jltbYT1inFFKWsd9OoV5jxDoEr0xQjdTDVMTn0PGcFjV6T0cGmFjuZs9GXv3E1y
# 4l6iMC6YJWT7Zof9yMlrKvABnDz/UYyBuENlbFceZYdikQZJyzpHqjCkFH6/bPCd
# oduTqWuPsOp8ChsFqKjkRS8nMteMFUmQYZWyPE2KsT0HFP4663vY+X19D3VVQl1l
# NV2KhqGCAyAwggMcBgkqhkiG9w0BCQYxggMNMIIDCQIBATB3MGMxCzAJBgNVBAYT
# AlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjE7MDkGA1UEAxMyRGlnaUNlcnQg
# VHJ1c3RlZCBHNCBSU0E0MDk2IFNIQTI1NiBUaW1lU3RhbXBpbmcgQ0ECEAVEr/OU
# nQg5pr/bP1/lYRYwDQYJYIZIAWUDBAIBBQCgaTAYBgkqhkiG9w0BCQMxCwYJKoZI
# hvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yMzEwMTEyMjQ5MjlaMC8GCSqGSIb3DQEJ
# BDEiBCAnMd73nkhCsHqFh2AezJFtku6kJMZgZPn5yRrHHmwm8zANBgkqhkiG9w0B
# AQEFAASCAgAyfornm11D1S35vw0JA+9DwrzPaeUELBgqDvtYWJp4xcrbt0demH69
# SJcT5D0fC5adWScERmUzTt2TwUgu0LkyhGKvAu9y48PQGF18i+48Sd5Ci6UgRNxu
# spmh36Ch4P04fWxYDcfb9YGCFfkzPHtb7rfw5HSaco+ZhBEHAgFVNdyoxbkJh7O/
# rX9SW/uSnMookFH8aMWFG3SCCEuD/RdjsA9mQvNKKgE+T8U7yonMzFVT2fypUblR
# kGCrg6mkgiPXZG3S52t3hgwCvwUxJnXwnzEb2f1bHgf+cGO3qOw6m0fE8xG62kxG
# jN05YsVfJO6g67u8mLO6IK/G8TX8BQo6oX3VWbjTP9keQDP1CmH9eeTWZJK1V+re
# CfO6ruWwcRnncIiTqeQEI9hc5LcF5Tgpi7/oux6O4HlbRs3TXT9IKD5QawZ+LJpf
# 7zvOZN1mI8tKZUJX8TLeHNYw78Yy50cS1QRRf9CQyXG/JGAl+iDm0Tv3ovkScmAl
# 7wIsntOP0AZHpa7x2v7lrnbttB/UkPRgyaf3/shGO6H3yN+qtUnCaGyIvQTmLFOC
# Su6/0yNyl5aZOBI0Wwd/qgv9r/S0N9NvDdiqcdLxpY4Cb9iomXr7r982Ch6ObU6/
# ISCtMTKa1iSVnf/nqCd7IX26q0h9s5f1APA2AJne0TJnZvJvdsnPCQ==
# SIG # End signature block
