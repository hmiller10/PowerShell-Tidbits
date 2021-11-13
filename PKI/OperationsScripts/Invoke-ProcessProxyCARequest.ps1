#Requires -Module PSPKI
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
	Submit CSR to Proxy CA

.DESCRIPTION

	Submit Subordinate Certificate Authority request to Proxy Certificate Authority

.OUTPUTS
	None

.PARAMETER CertficateRequest
	Path to certificate request file

.PARAMETER CertficateAuthority
	Certificate Authority Fully Qualified Domain Name

.PARAMETER Credential
	<PSCredential>

.PARAMETER Destination
	<Destination path to return proxy certification certificate>

.EXAMPLE 
PS C:\>.\Invoke-ProcessProxyCARequest.ps1 -CertificateRequest <Path\To\CSR/REQ> -CertificateAuthority <CA FQDN> -Credential (Get-Credential) -Destination <Path\To\CSR>


#>

###########################################################################
#
#
# AUTHOR:  Heather Miller, Manager, Identity and Access Management
#          GTS GTI IAM Infrastructure
#
# VERSION HISTORY:
# 1.0 07/07/2020 - Initial release
#
# 
###########################################################################


param (
	[Parameter(Mandatory = $false,
			 HelpMessage = "Enter path to certificate request file. EG: C:\myrequest.csr",
			 ValueFromPipeline = $true)]
	[Alias('CSR')]
	[String]$CertificateRequest,
	[Parameter(Mandatory = $true,
			 HelpMessage = "Enter certificate authority DNS name",
			 ValueFromPipeline = $false)]
	[String]$CertificateAuthority,
	[Parameter(Mandatory = $false,
			 HelpMessage = "Enter certificate authority admin credentials",
			 ValueFromPipeline = $false)]
	[System.Management.Automation.PSCredential]$Credential,
	[Parameter(Mandatory = $false,
			 HelpMessage = "Enter the path to where the CSR file should be archived to once processed.",
			 ValueFromPipeline = $false)]
	[String]$Destination
)

#Import Modules
Try
{
	Import-Module PSPKI -ErrorAction Stop
}
Catch
{
	Try
	{
		$modulePath = "{0}\{1}\{2}\{3}" -f $env:ProgramFiles, "WindowsPowerShell", "Modules", "PSPKI"
		$psdPath = "{0}\{1}\{2}" -f $modulePath, (Get-Module -Name PSPKI).Version, "pspki.psd1"
		Import-Module $psdPath -ErrorAction Stop
	}
	Catch
	{
		Throw "PSPKI module could not be loaded. $($_.Exception.Message)"
	}
	
}

#Variables
$template = "SubCA" # must always use concatenated name format
$CA = Get-CertificationAuthority -ComputerName $CertificateAuthority
$attrib = "CertificateTemplate:$template"
$arvHostName = [System.Net.Dns]::GetHostByName("LocalHost").HostName






#Script

if ($srvHostName -eq $CA)
{
	$caServerService = (Get-Service -Name CertSvc).Status
	if ($caServerService -ne 'Running') { Start-Service -Name Certsvc }
	if ($? -eq $true) { Continue }
	else { exit }
}
else
{
	$caServerService = (Get-Service -Name "CertSvc" -ComputerName $CA).Status
	if ($caServerService -ne 'Running')
	{
		Invoke-Command -ComputerName $CA -ScriptBlock { Start-Service -Name CertSvc; if ($? -eq $true) { Continue }
			else { exit } }
	}
}

if ($PSBoundParameters.ContainsKey('CertificateRequest'))
{
	$csrs = Get-ChildItem -Path $CertificateRequest -Force
}
else
{
	$files = Get-ChildItem -Path $csrdir
	$csrs = $files | Where-Object { ($_.Extension -eq ".csr") -or ($_.Extension -eq ".req") -or ($_.Extension -eq ".txt") }
}

$csrs.foreach({
		$csr = $_
		Write-Verbose -message "Requesting certificate $($csr) ..." -Verbose
		
		# Submit the CSR to be signed by the CA
		$submitParams = @{
			Path = $csr.FullName
			CertificationAuthority = $CA.ComputerName
			Attribute = $attrib
		}
		
		if ($PSBoundParameters.ContainsKey('Credential'))
		{
			$submitParams.Add('Credential', $Credential)
		}
		
		Submit-CertificateRequest @submitParams
		
		Write-Verbose -message "Finished request $csr" -Verbose
		
		if ($PSBoundParameters.ContainsKey('Destination')) { $archiveFolder = $Destination }
		else
		{
			# Specify the location of the request files
			$csrdir = "C:\Proxy Requests"
			$archiveFolder = "{0}\{1}" -f $csrdir, "old requests"
		}
		
		
		if ((Test-Path -Path $archiveFolder -PathType Container) -eq $false) { mkdir $archiveFolder }
		Move-Item -Path $csr.FullName -Destination $archiveFolder -Force -PassThru
	}) # end foreach
# SIG # Begin signature block
# MIInDAYJKoZIhvcNAQcCoIIm/TCCJvkCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCCH3lYxy8xNGztT
# 3ENh8Kh+G+O9d9lEzcIZBzHIHa/ntqCCIagwggQVMIIC/aADAgECAgsEAAAAAAEx
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
# WC2UXROwuc0YDLR7v6RCLjCCBdwwggPEoAMCAQICEz4AAAAGO07hEEosasAAAQAA
# AAYwDQYJKoZIhvcNAQELBQAwVDETMBEGCgmSJomT8ixkARkWA2NvbTEYMBYGCgmS
# JomT8ixkARkWCERlbG9pdHRlMSMwIQYDVQQDExpEZWxvaXR0ZSBTSEEyIExldmVs
# IDIgQ0EgMjAeFw0xODEyMDQyMDEyNTlaFw0yMzEyMDQyMDIyNTlaMGwxEzARBgoJ
# kiaJk/IsZAEZFgNjb20xGDAWBgoJkiaJk/IsZAEZFghkZWxvaXR0ZTEWMBQGCgmS
# JomT8ixkARkWBmF0cmFtZTEjMCEGA1UEAxMaRGVsb2l0dGUgU0hBMiBMZXZlbCAz
# IENBIDQwggEiMA0GCSqGSIb3DQEBAQUAA4IBDwAwggEKAoIBAQCtpGyZiH0yaD0z
# r4MLFTe6w5qehr6kjJ0cXHs0s0SQt7MnVQqWNsdfsxQ2am9O4a8ulC5Khbxb8xg6
# xjFhX/ugP/CFBXdjqZYsBDun40SS+MphLiDwLHcUx/H212aW6bsJDQLfrnzQvESs
# JJTzs8xv+alKob+oG//jXeR9Db2PHRr3/BRkX+ybjLIgCU/ifX/yRClXYeuXdQDr
# J7d2hwuT+au9VjnpkpTcsVs6nYPcFM/IR/ApXTTNSNWuQ1QHvr0y8xKOMoGoR+Lt
# r8lkWpJAQvPtKwIOOf1Otq2p9QQ2YB4QYkI+WbDJgBMyFI7uHW1xSeEIBsG5onBj
# PVAlupqlAgMBAAGjggGNMIIBiTAQBgkrBgEEAYI3FQEEAwIBATAjBgkrBgEEAYI3
# FQIEFgQUKNXmcn9zvywlPduF/DoH4rB80AgwHQYDVR0OBBYEFKnGxgr3tK9rxe5H
# XQC9E29rCiZiMBkGCSsGAQQBgjcUAgQMHgoAUwB1AGIAQwBBMAsGA1UdDwQEAwIB
# hjASBgNVHRMBAf8ECDAGAQH/AgEAMB8GA1UdIwQYMBaAFEcuNu60nP9cXhh8uBPh
# vqkgHhSzMFwGA1UdHwRVMFMwUaBPoE2GS2h0dHA6Ly9wa2kuZGVsb2l0dGUuY29t
# L0NlcnRFbnJvbGwvRGVsb2l0dGUlMjBTSEEyJTIwTGV2ZWwlMjAyJTIwQ0ElMjAy
# LmNybDB2BggrBgEFBQcBAQRqMGgwZgYIKwYBBQUHMAKGWmh0dHA6Ly9wa2kuZGVs
# b2l0dGUuY29tL0NlcnRFbnJvbGwvU0hBMkxWTDJDQTJfRGVsb2l0dGUlMjBTSEEy
# JTIwTGV2ZWwlMjAyJTIwQ0ElMjAyKDEpLmNydDANBgkqhkiG9w0BAQsFAAOCAgEA
# WDtffgO53157nRcfi6HS5BBcrWnWMCdy9rXa0FemqQnhH6p8OcYoHMzuBkGJGlpD
# ZIMb3BYryU5Ljp00RL0LK89S6mxlmdeMvTbCt9zynRK4moFLja3qUTAmI9wguO8m
# olTkvN8bcwMbHivqkzqTBLGDjbK6v6Al8433aRplJMFgmcEWwKJHdET/p2NHksNv
# hmNuDBYG9laY3bToa+AZxWKdDsLNvlZrv+IghHEjKiFt5h3dyVKhbx1mXBgQ3TOP
# B/R3xH0gFq9otGM71kldMJh3jE8Ki3MZd/D8HyM5PhV7RPdnyK9TJKi2So7QN6PK
# K564/EYgwOT0TQhNRh4X9RRxJJdgK1t2MuqIYcm8rS+PCuY+uhlZAVHEEt9wDPMV
# MJSEHbNB8vRb05W5GMfhSLZVBKF/X/u6G+Z+ck1yb7btunS6bJ6st3uVvtpSJ7SF
# GZ7K3IPOLiP3c6ABdyFFUQudoWDe+HoiAD4ZoFpB3bmXVyw7LcNjNykK1oE6iq7q
# q46ISWZCs+F3xJLl8Zz5Cm7BZyiZCoZVrb6xaPhXfj0usaXfiLOHAMXkt/iZxyaK
# 0bU6iH7gCB4xqq9/nzbra68cDW0bVol28qxeh97U8IfhGLs68VICvkbM3avn2a+b
# bdKiVkCm8LpU+theny/RF19Ooc1e+EvZWZXwiF0ITa4wggaRMIIFeaADAgECAhNy
# ABdojjNmx1zWqakdAAEAF2iOMA0GCSqGSIb3DQEBCwUAMGwxEzARBgoJkiaJk/Is
# ZAEZFgNjb20xGDAWBgoJkiaJk/IsZAEZFghkZWxvaXR0ZTEWMBQGCgmSJomT8ixk
# ARkWBmF0cmFtZTEjMCEGA1UEAxMaRGVsb2l0dGUgU0hBMiBMZXZlbCAzIENBIDQw
# HhcNMjAwOTI0MTUyMjA4WhcNMjIwOTI0MTUyMjA4WjAZMRcwFQYDVQQDEw5IZWF0
# aGVyIE1pbGxlcjCCASIwDQYJKoZIhvcNAQEBBQADggEPADCCAQoCggEBANTCIo6b
# lwi73HxoD7HMx/2nCLnKxcBOuJM36tki9I2Jhn3eM20Ux/WJ88uKnR6SJtYYlUjk
# WMKSCz7aOvvml9TJqKD18bWsA8/XyN7gf95TpnE9kkS28obUyUanaq9AgBD6WI01
# 6cAdZvafcn+pRdiC9FYVwYVkAdhmo2z00BhAO5/Vs26ZUvV3ZXjVIffAetV5Paq4
# 7zGtACPvmVZaD3dcNOqRtoSYJSwIM/y4yAF65LZm4x+JI4X+tlEOsNb8LCRB93u7
# Cb6LAlNhNjXxJvLwcb2MrimmwPanGDvHMb2OaoRIWBwLeoX7LKawDc/DWG9wlHmn
# 7TLOAmSo8uUpZVkCAwEAAaOCA30wggN5MDwGCSsGAQQBgjcVBwQvMC0GJSsGAQQB
# gjcVCIGBvUmFvoUTgtWbPIPXjgeG8ckKXIPK9y3C8zICAWQCAR4wEwYDVR0lBAww
# CgYIKwYBBQUHAwMwCwYDVR0PBAQDAgeAMBsGCSsGAQQBgjcVCgQOMAwwCgYIKwYB
# BQUHAwMwIAYDVR0RBBkwF4EVaGVtaWxsZXJAZGVsb2l0dGUuY29tMB0GA1UdDgQW
# BBT07yecOGoljaAhTaU++HB99grLrTAfBgNVHSMEGDAWgBSpxsYK97Sva8XuR10A
# vRNvawomYjCCATsGA1UdHwSCATIwggEuMIIBKqCCASagggEihoHSbGRhcDovLy9D
# Tj1EZWxvaXR0ZSUyMFNIQTIlMjBMZXZlbCUyMDMlMjBDQSUyMDQsQ049dWthdHJh
# bWVlbTAwMixDTj1DRFAsQ049UHVibGljJTIwS2V5JTIwU2VydmljZXMsQ049U2Vy
# dmljZXMsQ049Q29uZmlndXJhdGlvbixEQz1kZWxvaXR0ZSxEQz1jb20/Y2VydGlm
# aWNhdGVSZXZvY2F0aW9uTGlzdD9iYXNlP29iamVjdENsYXNzPWNSTERpc3RyaWJ1
# dGlvblBvaW50hktodHRwOi8vcGtpLmRlbG9pdHRlLmNvbS9DZXJ0ZW5yb2xsL0Rl
# bG9pdHRlJTIwU0hBMiUyMExldmVsJTIwMyUyMENBJTIwNC5jcmwwggFXBggrBgEF
# BQcBAQSCAUkwggFFMIHEBggrBgEFBQcwAoaBt2xkYXA6Ly8vQ049RGVsb2l0dGUl
# MjBTSEEyJTIwTGV2ZWwlMjAzJTIwQ0ElMjA0LENOPUFJQSxDTj1QdWJsaWMlMjBL
# ZXklMjBTZXJ2aWNlcyxDTj1TZXJ2aWNlcyxDTj1Db25maWd1cmF0aW9uLERDPWRl
# bG9pdHRlLERDPWNvbT9jQUNlcnRpZmljYXRlP2Jhc2U/b2JqZWN0Q2xhc3M9Y2Vy
# dGlmaWNhdGlvbkF1dGhvcml0eTB8BggrBgEFBQcwAoZwaHR0cDovL3BraS5kZWxv
# aXR0ZS5jb20vQ2VydGVucm9sbC91a2F0cmFtZWVtMDAyLmF0cmFtZS5kZWxvaXR0
# ZS5jb21fRGVsb2l0dGUlMjBTSEEyJTIwTGV2ZWwlMjAzJTIwQ0ElMjA0KDEpLmNy
# dDANBgkqhkiG9w0BAQsFAAOCAQEAU8AnpnO8UhsU/lryddSmu+29mPFxnS0d/jwI
# Sj3ap/jt66Se8YxittOzMCoqYZXn7M60oB3Crwe1lQFm14IHY0xPaHjAyKv1V6nL
# qq63jYBQqAyEe6dr4K3uRFqGXmqEFy2qkD5eGg02+EWVzduRT0AzEpRAqziYlKKx
# 5LrRIC51OunEH6xlS/840hovKEwRAOxe+nY6rfr6Fsi4fxL63JuuMvs9IifaAXsx
# HtOkXHfA50j6XzX2LMdqPxGrFLOrHs66Lyy14ClaZp0zkZ7+zeP9hxBlwTBwMkD0
# fVLvN6IyrDZ29/t96eQ0AKwORo3uZQHjlayhIcafozstzvlnBDCCBskwggSxoAMC
# AQICEzQAAAAHiSF1iXPNJ/IAAAAAAAcwDQYJKoZIhvcNAQELBQAwUjETMBEGCgmS
# JomT8ixkARkWA2NvbTEYMBYGCgmSJomT8ixkARkWCERlbG9pdHRlMSEwHwYDVQQD
# ExhEZWxvaXR0ZSBTSEEyIExldmVsIDEgQ0EwHhcNMjAwODA1MTczMjU2WhcNMzAw
# ODA1MTc0MjU2WjBUMRMwEQYKCZImiZPyLGQBGRYDY29tMRgwFgYKCZImiZPyLGQB
# GRYIRGVsb2l0dGUxIzAhBgNVBAMTGkRlbG9pdHRlIFNIQTIgTGV2ZWwgMiBDQSAy
# MIICIjANBgkqhkiG9w0BAQEFAAOCAg8AMIICCgKCAgEAmPb6sHLB25JD286NfyR2
# RfuNgmSXaR2dLojx7rPDqiEWKM01mSdquzeXj7Qu/VQsiQLV/9oxwMArSvjJHRje
# Q2L7orPGytxWiO6nNHkKbPUCkBTmRALVcXK0iYmXhQjaypjx5y8bi3K13AR7axTb
# NlPEFy3z9TwFGftmeJOIvle3dBvOCxJre1mxmf544tkzq+Df0ENP8sA41WeQbA5Z
# yDa2C8PWm8XL59X00UgtMJcOq4fCG+xkjl7nnbQ4/AP7lGHGkl0bnYE5Xd/nVA86
# +wO+uTUcmbs0fJ9fKO3bq3wgiUaRyyBbUQ2NzGlgaffxqge2lM3WCmiQeHKyfKsO
# kfg41+6h7qUFywDoDkvnVBjJs2+tgImqqD6iwmgZWHt6PeIiwJA/IIKBf0t1O16G
# 39uim6NSiesSK+wfOMxyxZio/BzKGPOtv4PwosBlPKlhK5bbvMWY2RFsWQJ6LPiR
# XlE5NIYbh/CTyngIdM6Drwr57sIZGWbKCJc9nORteVgx3pgciFAxOFGn1k3zmxM8
# 3qYxxgKi6fql8KCgbo+l6luROLa5rsRfkGPtRXy1HWJ7xwcf8/JxLJGlp1rtnGnZ
# ljvb0Tbtwo8GwDoihSMSh9MoGrJTrtk8tnYf4UpLgGKjGyGOUBFGrRGQcEhWbzDT
# K5qZP/0f31d3CndzQORYAb8CAwEAAaOCAZQwggGQMBAGCSsGAQQBgjcVAQQDAgEC
# MCMGCSsGAQQBgjcVAgQWBBQV4b/ii/DtWsyFdE+p2v+xjwi3MzAdBgNVHQ4EFgQU
# Ry427rSc/1xeGHy4E+G+qSAeFLMwEQYDVR0gBAowCDAGBgRVHSAAMBkGCSsGAQQB
# gjcUAgQMHgoAUwB1AGIAQwBBMAsGA1UdDwQEAwIBhjASBgNVHRMBAf8ECDAGAQH/
# AgEBMB8GA1UdIwQYMBaAFL6LoCtmWVn6kHFRpaeoBkJOkDztMFgGA1UdHwRRME8w
# TaBLoEmGR2h0dHA6Ly9wa2kuZGVsb2l0dGUuY29tL0NlcnRFbnJvbGwvRGVsb2l0
# dGUlMjBTSEEyJTIwTGV2ZWwlMjAxJTIwQ0EuY3JsMG4GCCsGAQUFBwEBBGIwYDBe
# BggrBgEFBQcwAoZSaHR0cDovL3BraS5kZWxvaXR0ZS5jb20vQ2VydEVucm9sbC9T
# SEEyTFZMMUNBX0RlbG9pdHRlJTIwU0hBMiUyMExldmVsJTIwMSUyMENBLmNydDAN
# BgkqhkiG9w0BAQsFAAOCAgEAh56DUZ5xeRvT/JdAM8biqLOI4PLlIhIPufRfMPJl
# mo64dY9G/JVkF7qh8SHWm7umjSTvb357kayJCks5Y3VwS11A9HMsRK11083exB27
# HUBd2W3IyRv2KBZT+SsAsnhtb2slEuPqqrpFZC3u2RZa8XonKVVcX3wfFN0qxE+y
# XkjYMUNxr3kYuclb2kt/4/RggkfV06dL0X2lHktLMYILmr8Tb2/eU2S7//hrdcH/
# tcWZ29hiIzL0qayp0j2MBuXACV/ZDNheEBvD659p14ae23CrgTpXSLL68RwHjaQq
# FVf2EWPXjR1MVJSvjB3QKiGdXTltUu1MBsrRHbFwj83xhiS1nTSWfIxSM+NG0u+t
# j9SJ5fOQSEMlCe0achdoXPvF50uDwaTLxUOoBoDK2DKd8nJFa/x8/Gj35jn7RNp/
# /UuzbmIhOr7YZqdfiBwnGffm4rS577EnBSsQhuOjzujrJbJd3NP2ar293Zupr8d+
# QYUbU51ny+mUYbGQ8VQeZgo72XkFAnzx1vZw9UK5VU7pC0zlBZL/FNV6hbgcnxQ0
# K/qRAudJtx03GpNF5sqhyEC0ndvCSdljKsf4mvgNwrEDTa6HLtEKONisnLSg56Ir
# cPx5W/eD8Ksodlwpfg8UM/A942V8JRZJLgFrc+nqysPi+cINMTd/n40h2wzGRlcj
# DE8xggS6MIIEtgIBATCBgzBsMRMwEQYKCZImiZPyLGQBGRYDY29tMRgwFgYKCZIm
# iZPyLGQBGRYIZGVsb2l0dGUxFjAUBgoJkiaJk/IsZAEZFgZhdHJhbWUxIzAhBgNV
# BAMTGkRlbG9pdHRlIFNIQTIgTGV2ZWwgMyBDQSA0AhNyABdojjNmx1zWqakdAAEA
# F2iOMA0GCWCGSAFlAwQCAQUAoEwwGQYJKoZIhvcNAQkDMQwGCisGAQQBgjcCAQQw
# LwYJKoZIhvcNAQkEMSIEIFn9GsIUatrmtN1tPazCswVHdIOQfWoPqy4cW//PHxi1
# MA0GCSqGSIb3DQEBAQUABIIBAC03fNcrrMxq+kwqqr4bnX95YL/QBpZezZtzJNmk
# 9TrSNPSou1/MpjPdBl5NodRmzWgOzSCMZC6KuKSnLTeCxxBWmcCSkdCnItiZ8cWc
# uix+F374xQX5wZPcasHfQO4d/qgWqbbuWILFIkpuBYgIO9hF4sy7AzGjBR5M7LYt
# CTUTnr0kDydVDM2d+KpdXXdOhzNtViaRfMAAXNL6j9RgD54prdI8wGYQyyw0zGpp
# TcuLOEu0A7SksDY8/esyBRCeO0BQ2GF0b/mYqkemeHuJ+Z1STbGEHr/8kt9ImZAA
# I8vKxnK00jMRuab5USD+CJdXYclOeRGfaAyudhf5JI5iyJChggK5MIICtQYJKoZI
# hvcNAQkGMYICpjCCAqICAQEwazBbMQswCQYDVQQGEwJCRTEZMBcGA1UEChMQR2xv
# YmFsU2lnbiBudi1zYTExMC8GA1UEAxMoR2xvYmFsU2lnbiBUaW1lc3RhbXBpbmcg
# Q0EgLSBTSEEyNTYgLSBHMgIMJFS4fx4UU603+qF4MA0GCWCGSAFlAwQCAQUAoIIB
# DDAYBgkqhkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yMDEy
# MDExNTIyMjBaMC8GCSqGSIb3DQEJBDEiBCB03yhmyL0aU5Aiet0GEzGVtWLix3Pi
# CqlTmH079D69HTCBoAYLKoZIhvcNAQkQAgwxgZAwgY0wgYowgYcEFD7HZtXU1HLi
# Gx8hQ1IcMbeQ2UtoMG8wX6RdMFsxCzAJBgNVBAYTAkJFMRkwFwYDVQQKExBHbG9i
# YWxTaWduIG52LXNhMTEwLwYDVQQDEyhHbG9iYWxTaWduIFRpbWVzdGFtcGluZyBD
# QSAtIFNIQTI1NiAtIEcyAgwkVLh/HhRTrTf6oXgwDQYJKoZIhvcNAQEBBQAEggEA
# gydqkyVAfgLyL4SEhYwDGVsWbSNmZrasVbVrc0Jdb9rCBKVeCZ7+FeHnyqpqjxki
# 9cC1Z0UuHr5V5a+z7gupIUJ255DUNQ5pzcpXwrGlWrJRcPwhUkP/1Sa5dbzd0p6k
# hBR8QEVcSlQ57iYE8lIDd3ccl7UBOB4QN8HO07lniTWuYZJdQQ2NJgSkNsCON0KE
# bslU40PMB5bvv0p/FpCak0oIqUrKZp7X6pvQ/80Esu6DsCtT535ciZzW+9eVUjOL
# vJ5GKy7o7NGjmqwKkcSqbTqokH08bGjf1nKoPOlELPIIe5rQibBs5B2hPs3KIeou
# UNZbFGaXU9gbM89XKwAO3Q==
# SIG # End signature block
