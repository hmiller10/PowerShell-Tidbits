<#

.NOTES
THIS CODE IS MADE AVAILABLE AS IS, WITHOUT WARRANTY OF ANY KIND. THE
RISK OF THE USE OR THE RESULTS FROM THE USE OF THIS CODE REMAINS	WITH
THE USER.

.SYNOPSIS
    Collect list of failed requests from CA database within the 
    defined time period.

.DESCRIPTION
    This script will connect to the Certification Authority server passed
    into the script as a parameter and will utilize the StartDate value and
    EndDate value passed into the script as parameters to locate and report
    on the number of failed certificate requests within the CA database. 
    This script will not compact or cleanup white space in the database.

.PARAMETER CA
    Fully qualified domain name of Certificate Authority server

.PARAMETER StartDate
    Beginning date script should use to define the date range to search for
    failed requests after

.PARAMETER EndDate
    End date script should use to define the date range to search for
    failed requests before

.PARAMETER Report
	Switch: indicates that script should output results of search to a file

.PARAMETER ReportPath
	Path to where script will export CSV file of all failed certificate requests
	found during search of date range defined as input parameters

.OUTPUTS
    Console output for number of failed requests found

.EXAMPLE 
    PS> Get-FailedRequestsByDateRange.ps1 -CA ca.domain.com -StartDate "01/01/2019 00:00:00" -EndDate "12/31/2019 23:59:59"

.EXAMPLE 
    PS> Get-FailedRequestsByDateRange.ps1 -CA ca.domain.com -StartDate "01/01/2019 00:00:00" -EndDate "12/31/2019 23:59:59" -Report -ReportPath <FullPathToFile>

.LINK
    https://www.sysadmins.lv/blog-en/categoryview/powershellpowershellpkimodule.aspx

.LINK
    https://github.com/Crypt32/PSPKI

#>

###########################################################################
#
#
#	AUTHOR:  Heather Miller
#
#	VERSION HISTORY:
#	3.0 10/18/2021 - Updated reporting and output methods
#
# 
###########################################################################
[CmdletBinding()]
Param (
	[Parameter(Mandatory = $false, Position = 0, ValueFromPipeline = $true, HelpMessage = "Specify the fully qualified domain name of the PKI server.")]
	[String]$CA,
	[Parameter(Mandatory = $true, Position = 1, ValueFromPipeline = $true, HelpMessage = "Specify the Request.SubmittedWhen after date range. Format - ""MM/dd/yyyy HH:mm:ss""")]
	[DateTime]$StartDate,
	[Parameter(Mandatory = $true, Position = 2, ValueFromPipeline = $true, HelpMessage = "Specify the Request.SubmittedWhen before date range. Format - ""MM/dd/yyyy HH:mm:ss""")]
	[DateTime]$EndDate,
	[Parameter(Mandatory = $false, Position = 3, HelpMessage = "Switch that specifies whether or not to export results of search to CSV")]
	[Parameter(ParameterSetName = 'ReportParameterSet')]
	[ValidateNotNullOrEmpty()]
	[Switch]$Report,
	[Parameter(Mandatory = $false, Position = 4, HelpMessage = "Specify the file path where the report results should be sent to")]
	[Parameter(ParameterSetName = 'ReportParameterSet')]
	[ValidateNotNullOrEmpty()]
	[String]$ReportPath
)

#Modules
try
{
	Import-Module PSPKI -Force
}
catch
{
	try
	{
		$module = Get-Module -Name PSPKI;
		$modulePath = Split-Path $module.Path;
		$psdPath = "{0}\{1}" -f $modulePath, "PSPKI.psd1"
		Import-Module $psdPath -ErrorAction Stop
	}
	catch
	{
		throw "PSPKI module could not be loaded. $($_.Exception.Message)"
	}
	
}

#Variables
$certProps = @()
$certProps = @("RequestID", "Request.RequesterName", "CommonName", "NotBefore", "NotAfter", "SerialNumber", "CertificateTemplate")
$failedRequests = @()
[int32]$pageSize = 100000






#Script
$Error.Clear()

Write-Verbose -Message "Working on search of $CA"
[int32]$pageNumber = 1
[int32]$LastID = 0

Try
{
	Do
	{
		$r = 0
		Connect-CertificationAuthority -ComputerName $CA -ErrorAction Stop | Out-Null
		Get-FailedRequest -CertificationAuthority $CA -Filter "Request.SubmittedWhen -ge $StartDate", "Request.SubmittedWhen -le $EndDate" -PageSize $pageSize | `
		Select-Object $certProps | ForEach-Object {
			$FailedRequests += $_; $r++
		}
		$pageNumber++
		
	}
	While ($r -eq $pageSize)
	
	If (($FailedRequests.Count -gt 0) -and ($PSBoundParameters.ContainsKey('Report')))
	{
		Write-Output ("Total number of failed certificate requests on {0} is: {1}" -f $CA, $FailedRequests.Count)
		$FailedRequests | Select-Object -Property RequestID, Request.RequesterName, CommonName, ConfigString, NotAfter, NotBefore, SerialNumber, `
								  @{ Name = "CertificateTemplate"; Expression = { $_.CertificateTemplateOID.FriendlyName } } | Export-Csv -Path $ReportPath -Append -NoTypeInformation
	}
	ElseIf (($FailedRequests.Count -gt 0))
	{
		Write-Output ("Total number of failed certificate requests on {0} is: {1}" -f $CA, $FailedRequests.Count)
	}
	Else
	{
		Write-Output ("There are no failed certificate requests {0} from {1} until {2}." -f $CA, $StartDate, $EndDate)
	}
	
}
Catch
{
	$errorMessage = "{0}: {1}" -f $Error[0], $Error[0].InvocationInfo.PositionMessage
	Write-Error $errorMessage -ErrorAction Continue
}
Finally
{
	[System.GC]::GetTotalMemory('forcefullcollection') | Out-Null
}


# SIG # Begin signature block
# MII0VAYJKoZIhvcNAQcCoII0RTCCNEECAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCDKy0k7PReAGGeT
# h88F3/aZib0Y6H0LGMI4l1nJvZhx6KCCLjkwggNfMIICR6ADAgECAgsEAAAAAAEh
# WFMIojANBgkqhkiG9w0BAQsFADBMMSAwHgYDVQQLExdHbG9iYWxTaWduIFJvb3Qg
# Q0EgLSBSMzETMBEGA1UEChMKR2xvYmFsU2lnbjETMBEGA1UEAxMKR2xvYmFsU2ln
# bjAeFw0wOTAzMTgxMDAwMDBaFw0yOTAzMTgxMDAwMDBaMEwxIDAeBgNVBAsTF0ds
# b2JhbFNpZ24gUm9vdCBDQSAtIFIzMRMwEQYDVQQKEwpHbG9iYWxTaWduMRMwEQYD
# VQQDEwpHbG9iYWxTaWduMIIBIjANBgkqhkiG9w0BAQEFAAOCAQ8AMIIBCgKCAQEA
# zCV2kHkGeCIW9cCDtoTKKJ79BXYRxa2IcvxGAkPHsoqdBF8kyy5L4WCCRuFSqwyB
# R3Bs3WTR6/Usow+CPQwrrpfXthSGEHm7OxOAd4wI4UnSamIvH176lmjfiSeVOJ8G
# 1z7JyyZZDXPesMjpJg6DFcbvW4vSBGDKSaYo9mk79svIKJHlnYphVzesdBTcdOA6
# 7nIvLpz70Lu/9T0A4QYz6IIrrlOmOhZzjN1BDiA6wLSnoemyT5AuMmDpV8u5BJJo
# aOU4JmB1sp93/5EU764gSfytQBVI0QIxYRleuJfvrXe3ZJp6v1/BE++bYvsNbOBU
# aRapA9pu6YOTcXbGaYWCFwIDAQABo0IwQDAOBgNVHQ8BAf8EBAMCAQYwDwYDVR0T
# AQH/BAUwAwEB/zAdBgNVHQ4EFgQUj/BLf6guRSSuTVD6Y5qL3uLdG7wwDQYJKoZI
# hvcNAQELBQADggEBAEtA28BQqv7IDO/3llRFSbuWAAlBrLMThoYoBzPKa+Z0uboA
# La6kCtP18fEPir9zZ0qDx0R7eOCvbmxvAymOMzlFw47kuVdsqvwSluxTxi3kJGy5
# lGP73FNoZ1Y+g7jPNSHDyWj+ztrCU6rMkIrp8F1GjJXdelgoGi8d3s0AN0GP7URt
# 11Mol37zZwQeFdeKlrTT3kwnpEwbc3N29BeZwh96DuMtCK0KHCz/PKtVDg+Rfjbr
# w1dJvuEuLXxgi8NBURMjnc73MmuUAaiZ5ywzHzo7JdKGQM47LIZ4yWEvFLru21Vv
# 34TuBQlNvSjYcs7TYlBlHuuSl4Mx2bO1ykdYP18wggVHMIIEL6ADAgECAg0B8kBC
# QM79ItvpbHH8MA0GCSqGSIb3DQEBDAUAMEwxIDAeBgNVBAsTF0dsb2JhbFNpZ24g
# Um9vdCBDQSAtIFIzMRMwEQYDVQQKEwpHbG9iYWxTaWduMRMwEQYDVQQDEwpHbG9i
# YWxTaWduMB4XDTE5MDIyMDAwMDAwMFoXDTI5MDMxODEwMDAwMFowTDEgMB4GA1UE
# CxMXR2xvYmFsU2lnbiBSb290IENBIC0gUjYxEzARBgNVBAoTCkdsb2JhbFNpZ24x
# EzARBgNVBAMTCkdsb2JhbFNpZ24wggIiMA0GCSqGSIb3DQEBAQUAA4ICDwAwggIK
# AoICAQCVB+hzymb57BTKezz3DQjxtEULLIK0SMbrWzyug7hBkjMUpG9/6SrMxrCI
# a8W2idHGsv8UzlEUIexK3RtaxtaH7k06FQbtZGYLkoDKRN5zlE7zp4l/T3hjCMgS
# UG1CZi9NuXkoTVIaihqAtxmBDn7EirxkTCEcQ2jXPTyKxbJm1ZCatzEGxb7ibTIG
# ph75ueuqo7i/voJjUNDwGInf5A959eqiHyrScC5757yTu21T4kh8jBAHOP9msndh
# fuDqjDyqtKT285VKEgdt/Yyyic/QoGF3yFh0sNQjOvddOsqi250J3l1ELZDxgc1X
# kvp+vFAEYzTfa5MYvms2sjnkrCQ2t/DvthwTV5O23rL44oW3c6K4NapF8uCdNqFv
# VIrxclZuLojFUUJEFZTuo8U4lptOTloLR/MGNkl3MLxxN+Wm7CEIdfzmYRY/d9XZ
# kZeECmzUAk10wBTt/Tn7g/JeFKEEsAvp/u6P4W4LsgizYWYJarEGOmWWWcDwNf3J
# 2iiNGhGHcIEKqJp1HZ46hgUAntuA1iX53AWeJ1lMdjlb6vmlodiDD9H/3zAR+YXP
# M0j1ym1kFCx6WE/TSwhJxZVkGmMOeT31s4zKWK2cQkV5bg6HGVxUsWW2v4yb3BPp
# DW+4LtxnbsmLEbWEFIoAGXCDeZGXkdQaJ783HjIH2BRjPChMrwIDAQABo4IBJjCC
# ASIwDgYDVR0PAQH/BAQDAgEGMA8GA1UdEwEB/wQFMAMBAf8wHQYDVR0OBBYEFK5s
# BaOTE+Ki5+LXHNbH8H/IZ1OgMB8GA1UdIwQYMBaAFI/wS3+oLkUkrk1Q+mOai97i
# 3Ru8MD4GCCsGAQUFBwEBBDIwMDAuBggrBgEFBQcwAYYiaHR0cDovL29jc3AyLmds
# b2JhbHNpZ24uY29tL3Jvb3RyMzA2BgNVHR8ELzAtMCugKaAnhiVodHRwOi8vY3Js
# Lmdsb2JhbHNpZ24uY29tL3Jvb3QtcjMuY3JsMEcGA1UdIARAMD4wPAYEVR0gADA0
# MDIGCCsGAQUFBwIBFiZodHRwczovL3d3dy5nbG9iYWxzaWduLmNvbS9yZXBvc2l0
# b3J5LzANBgkqhkiG9w0BAQwFAAOCAQEASaxexYPzWsthKk2XShUpn+QUkKoJ+cR6
# nzUYigozFW1yhyJOQT9tCp4YrtviX/yV0SyYFDuOwfA2WXnzjYHPdPYYpOThaM/v
# f2VZQunKVTm808Um7nE4+tchAw+3TtlbYGpDtH0J0GBh3artAF5OMh7gsmyePLLC
# u5jTkHZqaa0a3KiJ2lhP0sKLMkrOVPs46TsHC3UKEdsLfCUn8awmzxFT5tzG4mE1
# MvTO3YPjGTrrwmijcgDIJDxOuFM8sRer5jUs+dNCKeZfYAOsQmGmsVdqM0LfNTGG
# yj43K9rE2iT1ThLytrm3R+q7IK1hFregM+Mtiae8szwBfyMagAk06TCCBX8wggNn
# oAMCAQICEBi1woRDkBKXQawJijNlphAwDQYJKoZIhvcNAQELBQAwUjETMBEGCgmS
# JomT8ixkARkWA2NvbTEYMBYGCgmSJomT8ixkARkWCERlbG9pdHRlMSEwHwYDVQQD
# ExhEZWxvaXR0ZSBTSEEyIExldmVsIDEgQ0EwHhcNMTUwOTAxMTUwNzI1WhcNMzUw
# OTAxMTUwNzI1WjBSMRMwEQYKCZImiZPyLGQBGRYDY29tMRgwFgYKCZImiZPyLGQB
# GRYIRGVsb2l0dGUxITAfBgNVBAMTGERlbG9pdHRlIFNIQTIgTGV2ZWwgMSBDQTCC
# AiIwDQYJKoZIhvcNAQEBBQADggIPADCCAgoCggIBAJT6jaqlaRONXadbNfp3jPl7
# +c31ACk6w9dEQBfXqRRZiCDB+tlYDFzmGFga8jUGlAGD2Zb7/07dHAinGuVwFiqe
# A2wXm5QpsXx7KhM90f5M5/wh0R50N2DGUTKKwHHKvNg3cSosJ8PaNFuUhg+QPszV
# 4fQ43AsdBuUTgGqor+5tx2M9h6xL6nSLjJDKI+2NonjIsPz7iI1VPMJqHZbZAZhq
# AHjjnFWGlgWU6SstDQv8WC+3h3zCnYEWmt9KRVXmZYjfAzIjf3sW36+ortNH0ALD
# cYmZXh/+gix7jg3YfuyitNFqdlovas/yHjgBpI/I9I1xZ51+Tx0FnFVb/BPLCCtU
# vecQ8ZoTrKzBpBgOhtxe2tanyNVp7b2AU/fHlIGUN8IQ6o8qY+tto41y5FMZX7lH
# rgRwdvk/ZEm2qUByIYwBN3/ewPxTrsxtGnq6+njQuXaDYzl99QX6eVtFpIW1/A9A
# HsX+7/jrG6KhZAJGjc0ktqGl+/ht1wMNq/rJQ5sopq/m6GixKOLzPqBcxVHgUDl7
# 4xgRJxrcV5kiD8LC7vh5YTr47VyDS27HM5dkFswtmq0T8cehy69YT49Hoss+2ygz
# 0PjflPZeTnrS+jTRYKaVQ6KklaORVMIeO8GSmql2fG9Yz4/0kviEbGypwid+/mF3
# 5ANjInfcfKDMkXwX8HnfAgMBAAGjUTBPMAsGA1UdDwQEAwIBhjAPBgNVHRMBAf8E
# BTADAQH/MB0GA1UdDgQWBBS+i6ArZllZ+pBxUaWnqAZCTpA87TAQBgkrBgEEAYI3
# FQEEAwIBADANBgkqhkiG9w0BAQsFAAOCAgEAg070cixu+t++VEqXsSn9ARgrmH1v
# qP+htfBkamXmJbO5TtudcZ+4X12nHxGSBwj6Is2p9Mr2SbWFKHF0TInqO+umcJ/I
# ODoBGYyV7MPNadBpZMpZnVoajvrEbhsmLxmN9jFmdeqM+sfwhvq7HqLth7Cy7ZUl
# e+g9lz5w7rSAByP0uIrOAOxdFfqDV9OGh6VFzG9GNQZDyjRdYCJZn+K3bUYevtq9
# ++ED/kbsPpyYjzcBe0xdvMJM9YlF1ivRUpjNUOvv7l8acfakRUcL/qQnf9OyJ3md
# xpscUzScUs8xCcSH8KCSJ+hy/TWD5bQGewrlmLJuUxWPseU2eakvzuvpwB3d9E+y
# v6vh+YY6fbYrGGod0YgTjlPomb7VLK9V2i0+84s8Z5qeLLBx68i4TZGLtrof3U9f
# 4qvEmzN3yhhJTC/phO2TwINqtge+pIG2PhqxLczZ88DXyMJIfyAymCfZ+iVNskOK
# GubCY03+bh1IcVfWXfPwWSvC9x+UuyMt8glS8qvfMblGdwoS+BaCnkvPbjK4CgAz
# 9sZwHSMNZ53eu2Uj601ZN7GTv7iVSk846VNFddLHNu1ZlcfWAcnOKOMQ6EOvngJV
# GzabqBaXIc3S9nceSVGpS5dk+I4OFCMynInT4RJslNQKSlFxUE28dDfL9VgtlF0T
# sLnNGAy0e7+kQi4wggXcMIIDxKADAgECAhM+AAAABjtO4RBKLGrAAAEAAAAGMA0G
# CSqGSIb3DQEBCwUAMFQxEzARBgoJkiaJk/IsZAEZFgNjb20xGDAWBgoJkiaJk/Is
# ZAEZFghEZWxvaXR0ZTEjMCEGA1UEAxMaRGVsb2l0dGUgU0hBMiBMZXZlbCAyIENB
# IDIwHhcNMTgxMjA0MjAxMjU5WhcNMjMxMjA0MjAyMjU5WjBsMRMwEQYKCZImiZPy
# LGQBGRYDY29tMRgwFgYKCZImiZPyLGQBGRYIZGVsb2l0dGUxFjAUBgoJkiaJk/Is
# ZAEZFgZhdHJhbWUxIzAhBgNVBAMTGkRlbG9pdHRlIFNIQTIgTGV2ZWwgMyBDQSA0
# MIIBIjANBgkqhkiG9w0BAQEFAAOCAQ8AMIIBCgKCAQEAraRsmYh9Mmg9M6+DCxU3
# usOanoa+pIydHFx7NLNEkLezJ1UKljbHX7MUNmpvTuGvLpQuSoW8W/MYOsYxYV/7
# oD/whQV3Y6mWLAQ7p+NEkvjKYS4g8Cx3FMfx9tdmlum7CQ0C36580LxErCSU87PM
# b/mpSqG/qBv/413kfQ29jx0a9/wUZF/sm4yyIAlP4n1/8kQpV2Hrl3UA6ye3docL
# k/mrvVY56ZKU3LFbOp2D3BTPyEfwKV00zUjVrkNUB769MvMSjjKBqEfi7a/JZFqS
# QELz7SsCDjn9TratqfUENmAeEGJCPlmwyYATMhSO7h1tcUnhCAbBuaJwYz1QJbqa
# pQIDAQABo4IBjTCCAYkwEAYJKwYBBAGCNxUBBAMCAQEwIwYJKwYBBAGCNxUCBBYE
# FCjV5nJ/c78sJT3bhfw6B+KwfNAIMB0GA1UdDgQWBBSpxsYK97Sva8XuR10AvRNv
# awomYjAZBgkrBgEEAYI3FAIEDB4KAFMAdQBiAEMAQTALBgNVHQ8EBAMCAYYwEgYD
# VR0TAQH/BAgwBgEB/wIBADAfBgNVHSMEGDAWgBRHLjbutJz/XF4YfLgT4b6pIB4U
# szBcBgNVHR8EVTBTMFGgT6BNhktodHRwOi8vcGtpLmRlbG9pdHRlLmNvbS9DZXJ0
# RW5yb2xsL0RlbG9pdHRlJTIwU0hBMiUyMExldmVsJTIwMiUyMENBJTIwMi5jcmww
# dgYIKwYBBQUHAQEEajBoMGYGCCsGAQUFBzAChlpodHRwOi8vcGtpLmRlbG9pdHRl
# LmNvbS9DZXJ0RW5yb2xsL1NIQTJMVkwyQ0EyX0RlbG9pdHRlJTIwU0hBMiUyMExl
# dmVsJTIwMiUyMENBJTIwMigxKS5jcnQwDQYJKoZIhvcNAQELBQADggIBAFg7X34D
# ud9ee50XH4uh0uQQXK1p1jAncva12tBXpqkJ4R+qfDnGKBzM7gZBiRpaQ2SDG9wW
# K8lOS46dNES9CyvPUupsZZnXjL02wrfc8p0SuJqBS42t6lEwJiPcILjvJqJU5Lzf
# G3MDGx4r6pM6kwSxg42yur+gJfON92kaZSTBYJnBFsCiR3RE/6djR5LDb4ZjbgwW
# BvZWmN206GvgGcVinQ7Czb5Wa7/iIIRxIyohbeYd3clSoW8dZlwYEN0zjwf0d8R9
# IBavaLRjO9ZJXTCYd4xPCotzGXfw/B8jOT4Ve0T3Z8ivUySotkqO0DejyiueuPxG
# IMDk9E0ITUYeF/UUcSSXYCtbdjLqiGHJvK0vjwrmProZWQFRxBLfcAzzFTCUhB2z
# QfL0W9OVuRjH4Ui2VQShf1/7uhvmfnJNcm+27bp0umyerLd7lb7aUie0hRmeytyD
# zi4j93OgAXchRVELnaFg3vh6IgA+GaBaQd25l1csOy3DYzcpCtaBOoqu6quOiElm
# QrPhd8SS5fGc+QpuwWcomQqGVa2+sWj4V349LrGl34izhwDF5Lf4mccmitG1Ooh+
# 4AgeMaqvf58262uvHA1tG1aJdvKsXofe1PCH4Ri7OvFSAr5GzN2r59mvm23SolZA
# pvC6VPrYXp8v0RdfTqHNXvhL2VmV8IhdCE2uMIIGWTCCBEGgAwIBAgINAewckkDe
# /S5AXXxHdDANBgkqhkiG9w0BAQwFADBMMSAwHgYDVQQLExdHbG9iYWxTaWduIFJv
# b3QgQ0EgLSBSNjETMBEGA1UEChMKR2xvYmFsU2lnbjETMBEGA1UEAxMKR2xvYmFs
# U2lnbjAeFw0xODA2MjAwMDAwMDBaFw0zNDEyMTAwMDAwMDBaMFsxCzAJBgNVBAYT
# AkJFMRkwFwYDVQQKExBHbG9iYWxTaWduIG52LXNhMTEwLwYDVQQDEyhHbG9iYWxT
# aWduIFRpbWVzdGFtcGluZyBDQSAtIFNIQTM4NCAtIEc0MIICIjANBgkqhkiG9w0B
# AQEFAAOCAg8AMIICCgKCAgEA8ALiMCP64BvhmnSzr3WDX6lHUsdhOmN8OSN5bXT8
# MeR0EhmW+s4nYluuB4on7lejxDXtszTHrMMM64BmbdEoSsEsu7lw8nKujPeZWl12
# rr9EqHxBJI6PusVP/zZBq6ct/XhOQ4j+kxkX2e4xz7yKO25qxIjw7pf23PMYoEuZ
# HA6HpybhiMmg5ZninvScTD9dW+y279Jlz0ULVD2xVFMHi5luuFSZiqgxkjvyen38
# DljfgWrhsGweZYIq1CHHlP5CljvxC7F/f0aYDoc9emXr0VapLr37WD21hfpTmU1b
# dO1yS6INgjcZDNCr6lrB7w/Vmbk/9E818ZwP0zcTUtklNO2W7/hn6gi+j0l6/5Cx
# 1PcpFdf5DV3Wh0MedMRwKLSAe70qm7uE4Q6sbw25tfZtVv6KHQk+JA5nJsf8sg2g
# lLCylMx75mf+pliy1NhBEsFV/W6RxbuxTAhLntRCBm8bGNU26mSuzv31BebiZtAO
# BSGssREGIxnk+wU0ROoIrp1JZxGLguWtWoanZv0zAwHemSX5cW7pnF0CTGA8zwKP
# Af1y7pLxpxLeQhJN7Kkm5XcCrA5XDAnRYZ4miPzIsk3bZPBFn7rBP1Sj2HYClWxq
# jcoiXPYMBOMp+kuwHNM3dITZHWarNHOPHn18XpbWPRmwl+qMUJFtr1eGfhA3HWsa
# FN8CAwEAAaOCASkwggElMA4GA1UdDwEB/wQEAwIBhjASBgNVHRMBAf8ECDAGAQH/
# AgEAMB0GA1UdDgQWBBTqFsZp5+PLV0U5M6TwQL7Qw71lljAfBgNVHSMEGDAWgBSu
# bAWjkxPioufi1xzWx/B/yGdToDA+BggrBgEFBQcBAQQyMDAwLgYIKwYBBQUHMAGG
# Imh0dHA6Ly9vY3NwMi5nbG9iYWxzaWduLmNvbS9yb290cjYwNgYDVR0fBC8wLTAr
# oCmgJ4YlaHR0cDovL2NybC5nbG9iYWxzaWduLmNvbS9yb290LXI2LmNybDBHBgNV
# HSAEQDA+MDwGBFUdIAAwNDAyBggrBgEFBQcCARYmaHR0cHM6Ly93d3cuZ2xvYmFs
# c2lnbi5jb20vcmVwb3NpdG9yeS8wDQYJKoZIhvcNAQEMBQADggIBAH/iiNlXZytC
# X4GnCQu6xLsoGFbWTL/bGwdwxvsLCa0AOmAzHznGFmsZQEklCB7km/fWpA2PHpby
# hqIX3kG/T+G8q83uwCOMxoX+SxUk+RhE7B/CpKzQss/swlZlHb1/9t6CyLefYdO1
# RkiYlwJnehaVSttixtCzAsw0SEVV3ezpSp9eFO1yEHF2cNIPlvPqN1eUkRiv3I2Z
# OBlYwqmhfqJuFSbqtPl/KufnSGRpL9KaoXL29yRLdFp9coY1swJXH4uc/LusTN76
# 3lNMg/0SsbZJVU91naxvSsguarnKiMMSME6yCHOfXqHWmc7pfUuWLMwWaxjN5Fk3
# hgks4kXWss1ugnWl2o0et1sviC49ffHykTAFnM57fKDFrK9RBvARxx0wxVFWYOh8
# lT0i49UKJFMnl4D6SIknLHniPOWbHuOqhIKJPsBK9SH+YhDtHTD89szqSCd8i3VC
# f2vL86VrlR8EWDQKie2CUOTRe6jJ5r5IqitV2Y23JSAOG1Gg1GOqg+pscmFKyfpD
# xMZXxZ22PLCLsLkcMe+97xTYFEBsIB3CLegLxo1tjLZx7VIh/j72n585Gq6s0i96
# ILH0rKod4i0UnfqWah3GPMrz2Ry/U02kR1l8lcRDQfkl4iwQfoH5DZSnffK1CfXY
# YHJAUJUg1ENEvvqglecgWbZ4xqRqqiKbMIIGZTCCBE2gAwIBAgIQAYTTqM43getX
# 9P2He4OusjANBgkqhkiG9w0BAQsFADBbMQswCQYDVQQGEwJCRTEZMBcGA1UEChMQ
# R2xvYmFsU2lnbiBudi1zYTExMC8GA1UEAxMoR2xvYmFsU2lnbiBUaW1lc3RhbXBp
# bmcgQ0EgLSBTSEEzODQgLSBHNDAeFw0yMTA1MjcxMDAwMTZaFw0zMjA2MjgxMDAw
# MTVaMGMxCzAJBgNVBAYTAkJFMRkwFwYDVQQKDBBHbG9iYWxTaWduIG52LXNhMTkw
# NwYDVQQDDDBHbG9iYWxzaWduIFRTQSBmb3IgTVMgQXV0aGVudGljb2RlIEFkdmFu
# Y2VkIC0gRzQwggGiMA0GCSqGSIb3DQEBAQUAA4IBjwAwggGKAoIBgQDiopu2Sfs0
# SCgjB4b9UhNNusuqNeL5QBwbe2nFmCrMyVzvJ8bsuCVlwz8dROfe4QjvBBcAlZcM
# /dtdg7SI66COm0+DuvnfXhhUagIODuZU8DekHpxnMQW1N3F8en7YgWUz5JrqsDE3
# x2a0o7oFJ+puUoJY2YJWJI3567MU+2QAoXsqH3qeqGOR5tjRIsY/0p04P6+VaVsn
# v+hAJJnHH9l7kgUCfSjGPDn3es33ZSagN68yBXeXauEQG5iFLISt5SWGfHOezYiN
# Syp6nQ9Zeb3y2jZ+Zqwu+LuIl8ltefKz1NXMGvRPi0WVdvKHlYCOKHm6/cVwr7wa
# FAKQfCZbEYtd9brkEQLFgRxmaEveaM6dDIhhqraUI53gpDxGXQRR2z9ZC+fsvtLZ
# EypH70sSEm7INc/uFjK20F+FuE/yfNgJKxJewMLvEzFwNnPc1ldU01dgnhwQlfDm
# qi8Qiht+yc2PzlBLHCWowBdkURULjM/XyV1KbEl0rlrxagZ1Pok3O5ECAwEAAaOC
# AZswggGXMA4GA1UdDwEB/wQEAwIHgDAWBgNVHSUBAf8EDDAKBggrBgEFBQcDCDAd
# BgNVHQ4EFgQUda8nP7jbmuxvHO7DamT2v4Q1sM4wTAYDVR0gBEUwQzBBBgkrBgEE
# AaAyAR4wNDAyBggrBgEFBQcCARYmaHR0cHM6Ly93d3cuZ2xvYmFsc2lnbi5jb20v
# cmVwb3NpdG9yeS8wCQYDVR0TBAIwADCBkAYIKwYBBQUHAQEEgYMwgYAwOQYIKwYB
# BQUHMAGGLWh0dHA6Ly9vY3NwLmdsb2JhbHNpZ24uY29tL2NhL2dzdHNhY2FzaGEz
# ODRnNDBDBggrBgEFBQcwAoY3aHR0cDovL3NlY3VyZS5nbG9iYWxzaWduLmNvbS9j
# YWNlcnQvZ3N0c2FjYXNoYTM4NGc0LmNydDAfBgNVHSMEGDAWgBTqFsZp5+PLV0U5
# M6TwQL7Qw71lljBBBgNVHR8EOjA4MDagNKAyhjBodHRwOi8vY3JsLmdsb2JhbHNp
# Z24uY29tL2NhL2dzdHNhY2FzaGEzODRnNC5jcmwwDQYJKoZIhvcNAQELBQADggIB
# ADiTt301iTTqGtaqes6NhNvhNLd0pf/YXZQ2JY/SgH6hZbGbzzVRXdugS273IUAu
# 7E9vFkByHHUbMAAXOY/IL6RxziQzJpDV5P85uWHvC8o58y1ejaD/TuFWZB/UnHYE
# pERcPWKFcC/5TqT3hlbbekkmQy0Fm+LDibc6oS0nJxjGQ4vcQ6G2ci0/2cY0igLT
# Yjkp8H0o0KnDZIpGbbNDHHSL3bmmCyF7EacfXaLbjOBV02n6d9FdFLmW7JFFGxts
# fkJAJKTtQMZl+kGPSDGc47izF1eCecrMHsLQT08FDg1512ndlaFxXYqe51rCT6gG
# DwiJe9tYyCV9/2i8KKJwnLsMtVPojgaxsoKBhxKpXndMk6sY+ERXWBHL9pMVSTG3
# U1Ah2tX8YH/dMMWsUUQLZ6X61nc0rRIfKPuI2lGbRJredw7uMhJgVgyRnViPvJlX
# 8r7NucNzJBnad6bk7PHeb+C8hB1vw/Hb4dVCUYZREkImPtPqE/QonK1NereiuhRq
# P0BVWE6MZRyz9nXWf64PhIAvvoh4XCcfRxfCPeRpnsuunu8CaIg3EMJsJorIjGWQ
# U02uXdq4RhDUkAqK//QoQIHgUsjyAWRIGIR4aiL6ypyqDh3FjnLDNiIZ6/iUH7/C
# xQFW6aaA6gEdEzUH4rl0FP2aOJ4D0kn2TOuhvRwU0uOZMIIGkTCCBXmgAwIBAgIT
# cgAXaI4zZsdc1qmpHQABABdojjANBgkqhkiG9w0BAQsFADBsMRMwEQYKCZImiZPy
# LGQBGRYDY29tMRgwFgYKCZImiZPyLGQBGRYIZGVsb2l0dGUxFjAUBgoJkiaJk/Is
# ZAEZFgZhdHJhbWUxIzAhBgNVBAMTGkRlbG9pdHRlIFNIQTIgTGV2ZWwgMyBDQSA0
# MB4XDTIwMDkyNDE1MjIwOFoXDTIyMDkyNDE1MjIwOFowGTEXMBUGA1UEAxMOSGVh
# dGhlciBNaWxsZXIwggEiMA0GCSqGSIb3DQEBAQUAA4IBDwAwggEKAoIBAQDUwiKO
# m5cIu9x8aA+xzMf9pwi5ysXATriTN+rZIvSNiYZ93jNtFMf1ifPLip0ekibWGJVI
# 5FjCkgs+2jr75pfUyaig9fG1rAPP18je4H/eU6ZxPZJEtvKG1MlGp2qvQIAQ+liN
# NenAHWb2n3J/qUXYgvRWFcGFZAHYZqNs9NAYQDuf1bNumVL1d2V41SH3wHrVeT2q
# uO8xrQAj75lWWg93XDTqkbaEmCUsCDP8uMgBeuS2ZuMfiSOF/rZRDrDW/CwkQfd7
# uwm+iwJTYTY18Sby8HG9jK4ppsD2pxg7xzG9jmqESFgcC3qF+yymsA3Pw1hvcJR5
# p+0yzgJkqPLlKWVZAgMBAAGjggN9MIIDeTA8BgkrBgEEAYI3FQcELzAtBiUrBgEE
# AYI3FQiBgb1Jhb6FE4LVmzyD144HhvHJClyDyvctwvMyAgFkAgEeMBMGA1UdJQQM
# MAoGCCsGAQUFBwMDMAsGA1UdDwQEAwIHgDAbBgkrBgEEAYI3FQoEDjAMMAoGCCsG
# AQUFBwMDMCAGA1UdEQQZMBeBFWhlbWlsbGVyQGRlbG9pdHRlLmNvbTAdBgNVHQ4E
# FgQU9O8nnDhqJY2gIU2lPvhwffYKy60wHwYDVR0jBBgwFoAUqcbGCve0r2vF7kdd
# AL0Tb2sKJmIwggE7BgNVHR8EggEyMIIBLjCCASqgggEmoIIBIoaB0mxkYXA6Ly8v
# Q049RGVsb2l0dGUlMjBTSEEyJTIwTGV2ZWwlMjAzJTIwQ0ElMjA0LENOPXVrYXRy
# YW1lZW0wMDIsQ049Q0RQLENOPVB1YmxpYyUyMEtleSUyMFNlcnZpY2VzLENOPVNl
# cnZpY2VzLENOPUNvbmZpZ3VyYXRpb24sREM9ZGVsb2l0dGUsREM9Y29tP2NlcnRp
# ZmljYXRlUmV2b2NhdGlvbkxpc3Q/YmFzZT9vYmplY3RDbGFzcz1jUkxEaXN0cmli
# dXRpb25Qb2ludIZLaHR0cDovL3BraS5kZWxvaXR0ZS5jb20vQ2VydGVucm9sbC9E
# ZWxvaXR0ZSUyMFNIQTIlMjBMZXZlbCUyMDMlMjBDQSUyMDQuY3JsMIIBVwYIKwYB
# BQUHAQEEggFJMIIBRTCBxAYIKwYBBQUHMAKGgbdsZGFwOi8vL0NOPURlbG9pdHRl
# JTIwU0hBMiUyMExldmVsJTIwMyUyMENBJTIwNCxDTj1BSUEsQ049UHVibGljJTIw
# S2V5JTIwU2VydmljZXMsQ049U2VydmljZXMsQ049Q29uZmlndXJhdGlvbixEQz1k
# ZWxvaXR0ZSxEQz1jb20/Y0FDZXJ0aWZpY2F0ZT9iYXNlP29iamVjdENsYXNzPWNl
# cnRpZmljYXRpb25BdXRob3JpdHkwfAYIKwYBBQUHMAKGcGh0dHA6Ly9wa2kuZGVs
# b2l0dGUuY29tL0NlcnRlbnJvbGwvdWthdHJhbWVlbTAwMi5hdHJhbWUuZGVsb2l0
# dGUuY29tX0RlbG9pdHRlJTIwU0hBMiUyMExldmVsJTIwMyUyMENBJTIwNCgxKS5j
# cnQwDQYJKoZIhvcNAQELBQADggEBAFPAJ6ZzvFIbFP5a8nXUprvtvZjxcZ0tHf48
# CEo92qf47euknvGMYrbTszAqKmGV5+zOtKAdwq8HtZUBZteCB2NMT2h4wMir9Vep
# y6qut42AUKgMhHuna+Ct7kRahl5qhBctqpA+XhoNNvhFlc3bkU9AMxKUQKs4mJSi
# seS60SAudTrpxB+sZUv/ONIaLyhMEQDsXvp2Oq36+hbIuH8S+tybrjL7PSIn2gF7
# MR7TpFx3wOdI+l819izHaj8RqxSzqx7Oui8steApWmadM5Ge/s3j/YcQZcEwcDJA
# 9H1S7zeiMqw2dvf7fenkNACsDkaN7mUB45WsoSHGn6M7Lc75ZwQwggbJMIIEsaAD
# AgECAhM0AAAAB4khdYlzzSfyAAAAAAAHMA0GCSqGSIb3DQEBCwUAMFIxEzARBgoJ
# kiaJk/IsZAEZFgNjb20xGDAWBgoJkiaJk/IsZAEZFghEZWxvaXR0ZTEhMB8GA1UE
# AxMYRGVsb2l0dGUgU0hBMiBMZXZlbCAxIENBMB4XDTIwMDgwNTE3MzI1NloXDTMw
# MDgwNTE3NDI1NlowVDETMBEGCgmSJomT8ixkARkWA2NvbTEYMBYGCgmSJomT8ixk
# ARkWCERlbG9pdHRlMSMwIQYDVQQDExpEZWxvaXR0ZSBTSEEyIExldmVsIDIgQ0Eg
# MjCCAiIwDQYJKoZIhvcNAQEBBQADggIPADCCAgoCggIBAJj2+rBywduSQ9vOjX8k
# dkX7jYJkl2kdnS6I8e6zw6ohFijNNZknars3l4+0Lv1ULIkC1f/aMcDAK0r4yR0Y
# 3kNi+6KzxsrcVojupzR5Cmz1ApAU5kQC1XFytImJl4UI2sqY8ecvG4tytdwEe2sU
# 2zZTxBct8/U8BRn7ZniTiL5Xt3QbzgsSa3tZsZn+eOLZM6vg39BDT/LAONVnkGwO
# Wcg2tgvD1pvFy+fV9NFILTCXDquHwhvsZI5e5520OPwD+5RhxpJdG52BOV3f51QP
# OvsDvrk1HJm7NHyfXyjt26t8IIlGkcsgW1ENjcxpYGn38aoHtpTN1gpokHhysnyr
# DpH4ONfuoe6lBcsA6A5L51QYybNvrYCJqqg+osJoGVh7ej3iIsCQPyCCgX9LdTte
# ht/bopujUonrEivsHzjMcsWYqPwcyhjzrb+D8KLAZTypYSuW27zFmNkRbFkCeiz4
# kV5ROTSGG4fwk8p4CHTOg68K+e7CGRlmygiXPZzkbXlYMd6YHIhQMThRp9ZN85sT
# PN6mMcYCoun6pfCgoG6PpepbkTi2ua7EX5Bj7UV8tR1ie8cHH/PycSyRpada7Zxp
# 2ZY729E27cKPBsA6IoUjEofTKBqyU67ZPLZ2H+FKS4BioxshjlARRq0RkHBIVm8w
# 0yuamT/9H99Xdwp3c0DkWAG/AgMBAAGjggGUMIIBkDAQBgkrBgEEAYI3FQEEAwIB
# AjAjBgkrBgEEAYI3FQIEFgQUFeG/4ovw7VrMhXRPqdr/sY8ItzMwHQYDVR0OBBYE
# FEcuNu60nP9cXhh8uBPhvqkgHhSzMBEGA1UdIAQKMAgwBgYEVR0gADAZBgkrBgEE
# AYI3FAIEDB4KAFMAdQBiAEMAQTALBgNVHQ8EBAMCAYYwEgYDVR0TAQH/BAgwBgEB
# /wIBATAfBgNVHSMEGDAWgBS+i6ArZllZ+pBxUaWnqAZCTpA87TBYBgNVHR8EUTBP
# ME2gS6BJhkdodHRwOi8vcGtpLmRlbG9pdHRlLmNvbS9DZXJ0RW5yb2xsL0RlbG9p
# dHRlJTIwU0hBMiUyMExldmVsJTIwMSUyMENBLmNybDBuBggrBgEFBQcBAQRiMGAw
# XgYIKwYBBQUHMAKGUmh0dHA6Ly9wa2kuZGVsb2l0dGUuY29tL0NlcnRFbnJvbGwv
# U0hBMkxWTDFDQV9EZWxvaXR0ZSUyMFNIQTIlMjBMZXZlbCUyMDElMjBDQS5jcnQw
# DQYJKoZIhvcNAQELBQADggIBAIeeg1GecXkb0/yXQDPG4qiziODy5SISD7n0XzDy
# ZZqOuHWPRvyVZBe6ofEh1pu7po0k729+e5GsiQpLOWN1cEtdQPRzLEStddPN3sQd
# ux1AXdltyMkb9igWU/krALJ4bW9rJRLj6qq6RWQt7tkWWvF6JylVXF98HxTdKsRP
# sl5I2DFDca95GLnJW9pLf+P0YIJH1dOnS9F9pR5LSzGCC5q/E29v3lNku//4a3XB
# /7XFmdvYYiMy9KmsqdI9jAblwAlf2QzYXhAbw+ufadeGnttwq4E6V0iy+vEcB42k
# KhVX9hFj140dTFSUr4wd0CohnV05bVLtTAbK0R2xcI/N8YYktZ00lnyMUjPjRtLv
# rY/UieXzkEhDJQntGnIXaFz7xedLg8Gky8VDqAaAytgynfJyRWv8fPxo9+Y5+0Ta
# f/1Ls25iITq+2GanX4gcJxn35uK0ue+xJwUrEIbjo87o6yWyXdzT9mq9vd2bqa/H
# fkGFG1OdZ8vplGGxkPFUHmYKO9l5BQJ88db2cPVCuVVO6QtM5QWS/xTVeoW4HJ8U
# NCv6kQLnSbcdNxqTRebKochAtJ3bwknZYyrH+Jr4DcKxA02uhy7RCjjYrJy0oOei
# K3D8eVv3g/CrKHZcKX4PFDPwPeNlfCUWSS4Ba3Pp6srD4vnCDTE3f5+NIdsMxkZX
# IwxPMYIFcTCCBW0CAQEwgYMwbDETMBEGCgmSJomT8ixkARkWA2NvbTEYMBYGCgmS
# JomT8ixkARkWCGRlbG9pdHRlMRYwFAYKCZImiZPyLGQBGRYGYXRyYW1lMSMwIQYD
# VQQDExpEZWxvaXR0ZSBTSEEyIExldmVsIDMgQ0EgNAITcgAXaI4zZsdc1qmpHQAB
# ABdojjANBglghkgBZQMEAgEFAKBMMBkGCSqGSIb3DQEJAzEMBgorBgEEAYI3AgEE
# MC8GCSqGSIb3DQEJBDEiBCALmdttv3WNsvss3oKPD9qRVQaaXT5C5l0ajstXEMta
# LjANBgkqhkiG9w0BAQEFAASCAQB/AlrzLMNoCviwF87EwOaEFgIJ8aYl92Qprock
# 0PSn2IkS/PkC6/6MacxzI/s+upMUGzIOPOwFakUBB/pMcz4d45aJjqcLKbrKYl3+
# cFF+PmmiSbm5eBTjf7GBd4ad1CgN2oUtAxNjbc+sXb96IXDNYMlggtzFaFUy/zQv
# D2+snLHXhraXHJ5RFTE+nPhW1graftUKp/JfnTEgX5J99W+Jqd9uJAa+99ZGubez
# PIq9xSUed3CM1yM8k8V57Azslb1UYE59IkJ77jI+ovIHDgnPqKrYwMI8AXKqsE9p
# Sf8VKXQZwuFiIBPx0BZBBVDpOsQFTLcSTTTWds+2FZmuuUpboYIDcDCCA2wGCSqG
# SIb3DQEJBjGCA10wggNZAgEBMG8wWzELMAkGA1UEBhMCQkUxGTAXBgNVBAoTEEds
# b2JhbFNpZ24gbnYtc2ExMTAvBgNVBAMTKEdsb2JhbFNpZ24gVGltZXN0YW1waW5n
# IENBIC0gU0hBMzg0IC0gRzQCEAGE06jON4HrV/T9h3uDrrIwDQYJYIZIAWUDBAIB
# BQCgggE/MBgGCSqGSIb3DQEJAzELBgkqhkiG9w0BBwEwHAYJKoZIhvcNAQkFMQ8X
# DTIxMTAxODIxMDAzMlowLQYJKoZIhvcNAQk0MSAwHjANBglghkgBZQMEAgEFAKEN
# BgkqhkiG9w0BAQsFADAvBgkqhkiG9w0BCQQxIgQg9ZsRugTlm1vScoxu3fVm+NkG
# e6NMdjxaYADeAUNupUYwgaQGCyqGSIb3DQEJEAIMMYGUMIGRMIGOMIGLBBTdV7Wz
# hzyGGynGrsRzGvvojXXBSTBzMF+kXTBbMQswCQYDVQQGEwJCRTEZMBcGA1UEChMQ
# R2xvYmFsU2lnbiBudi1zYTExMC8GA1UEAxMoR2xvYmFsU2lnbiBUaW1lc3RhbXBp
# bmcgQ0EgLSBTSEEzODQgLSBHNAIQAYTTqM43getX9P2He4OusjANBgkqhkiG9w0B
# AQsFAASCAYCiCUZXF4muxjTIe8The9ThPZM/F0QyVayS+OINlc7zvf7Q7m1IiTZH
# xveFwqnothGVAaa7FxWCqP/CzVcZEPrj1wUHRvO0zR+2cTCHNpvtE+/Xyk2F1Zmx
# L+VLEVcKk/hzQ0xJzEhrvPoC39t7Wkh8vKOzpwdkxB+ffZNtDQE5R+zp4Dj3QJuX
# GeDouOeDUV93Rb9g1/QU/tCz3/Uu9OsmxQIeYFXW5HXML0q+W0rjQDjOjuBE2rEu
# 88IlcKlTZhcSSGySpu34xZtIlqSzqllkZOLKgoqOpijlQq9/RQGwbydRYJq8LPtF
# CVs/Xig/fxx1ZL/edqCUo0EECW5p40jO17KDeGLb0ddeHu/wSubERc9u3MRYx01p
# 8M8yLWjaNoypMJyXefdb8/1TejEhqAK77qGRDYOzu2x2ICSptMPfoFJgQ0Z00LdW
# kR3aoo+2ErRFo/RvB2TiOJTQzyC/2edXTWJ3fu6kz4l/D8HK9xmX7avaSSyEC20D
# q9Jk+arK8PQ=
# SIG # End signature block
