#Requires -Module PSPKI
<#
	.NOTES
	THIS CODE IS MADE AVAILABLE AS IS, WITHOUT WARRANTY OF ANY KIND. THE
	RISK OF THE USE OR THE RESULTS FROM THE USE OF THIS CODE REMAINS WITH
	THE USER.

	.SYNOPSIS
		Script to process pending certificate issuance requests using PSPKI module
		
	.DESCRIPTION
		This script will connect to the certificate authority server passed into
		the script as a parameter and will locate any pending certificate issuance
		requests. Upon finding a pending a request, the person running this script
		will be prompted to either approve or deny the request. The script will then 
		respond accordingly depending on the user's response.

	.PARAMETER CA
		Fully qualified domain name of certificate authority server

	.OUTPUTS
		None

	.EXAMPLE
	PS C:\> Invoke-ProcessPendingRequest.ps1 -CA myca.domain.com -Destination 'C:\Temp'

	.LINK
    	https://www.sysadmins.lv/blog-en/categoryview/powershellpowershellpkimodule.aspx

	.LINK
    	https://github.com/Crypt32/PSPKI
	
#>


[CmdletBinding()]
Param (
	[Parameter(Mandatory = $true, Position = 0, ValueFromPipeline = $true, HelpMessage = "Specify the fully qualified domain name of the PKI server.")]
	[String]$CA,
	[Parameter(Mandatory = $true,
			 HelpMessage = "Enter the path to where the .CER file should be copied to.",
			 ValueFromPipeline = $false)]
	[String]$Destination
)

#Modules
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
$certProps = @()
$certProps = @("RawRequest","RequestID", "Request.RequesterName", "CommonName", "NotBefore", "NotAfter", "SerialNumber", "CertificateTemplate")
$pendingRequests = @()
Add-Type -Assembly System.Windows.Forms
Add-Type -Assembly System.Drawing
$srvHostName = [System.Net.Dns]::GetHostByName("LocalHost").HostName





#Script
try
{
	if ($srvHostName -eq $CA)
	{
		$caServerService = (Get-Service -Name "CertSvc").Status
		if ($caServerService -ne 'Running') { Start-Service -Name "Certsvc" }
		if ($? -eq $true) { Continue }
		else { exit }
	}
	else
	{
		$caServerService = (Get-Service -Name "CertSvc" -ComputerName $CA).Status
		if ($caServerService -ne 'Running')
		{
			Invoke-Command -ComputerName $CA -ScriptBlock {
				Start-Service -Name "CertSvc"; if ($? -eq $true) { Continue }
				else { exit }
			}
		}
	}
	
	$pendingRequests = Get-PendingRequest -CertificationAuthority $CA
	
	if ($pendingRequests.Count -gt 0)
	{
		Write-Output ("Total number of pending certificate requests at this time: $($pendingRequests.Count)")
		
		foreach ($certRequest in $pendingRequests)
		{
			$dbRow = Get-PendingRequest -CertificationAuthority $CA -Property "RawRequest" -RequestID $certRequest.RequestID
			$reqbytes = [convert]::FromBase64String($dbRow."Request.RawRequest")
			$req = New-Object System.Security.Cryptography.X509CertificateRequests.X509CertificateRequest( ,$reqBytes)
			
			Write-Output ("Please carefully review the fields in the below shown request. Pay extra attention to any SANs added to the request.")
			
			$form = New-Object System.Windows.Forms.Form
			$form.width = 800
			$form.height = 700
			$form.StartPosition = "CenterScreen"
			$form.Text = "Select Pending Certificate Request Approval Decision"
			$form.Font = New-Object System.Drawing.Font("Verdana", 10)
			
			
			$approveButton = New-Object System.Windows.Forms.Button
			$approveButton.Location = New-Object System.Drawing.Size(650, 550)
			$approveButton.Size = New-Object System.Drawing.Size(70, 25)
			$approveButton.Text = "Approve"
			$approveButton.DialogResult = [System.Windows.Forms.DialogResult]::Yes
			$form.Controls.Add($approveButton)
			
			$denyButton = New-Object System.Windows.Forms.Button
			$denyButton.Location = New-Object System.Drawing.Size(650, 600)
			$denyButton.Size = New-Object System.Drawing.Size(70, 25)
			$denyButton.Text = "Deny"
			$denyButton.DialogResult = [System.Windows.Forms.DialogResult]::No
			$form.Controls.Add($denyButton)
			
			$MyMultiLineTextBox = New-Object System.Windows.Forms.TextBox
			$MyMultiLineTextBox.Multiline = $true
			$MyMultiLineTextBox.Width = 725
			$MyMultiLineTextBox.Height = 500
			$MyMultiLineTextBox.Location = New-Object System.Drawing.Point(10, 10)
			$MyMultiLineTextBox.Font = 'Verdana,16'
			$MyMultiLineTextBox.ScrollBars = 'Vertical'
			$MyMultiLineTextBox.WordWrap = $true
			$MyMultiLineTextBox.TextAlign = 'Left'
			$form.Controls.Add($MyMultiLineTextBox)
			
			
			$MyMultiLineTextBox.AppendText($req.Format())
			$form.AcceptButton = $approveButton
			$form.CancelButton = $denyButton
			
			$form.Add_Shown({ $form.Activate(); $approveButton.Focus() })
			$selectedReason = $form.ShowDialog()
			
			# Check for ESC presses
			$form.KeyPreview = $True
			$form.Add_KeyDown({
					if ($_.KeyCode -eq "Escape")
					{
						# if escape, exit
						$form.Close()
					}
				})
			
			switch ($selectedReason)
			{
				"Yes" { $certRequest | Approve-CertificateRequest -ErrorAction Continue }
				"No" { $certRequest | Deny-CertificateRequest -ErrorAction Continue }
				
			}
			
			$cert = Get-AdcsDatabaseRow -CertificationAuthority $CA -Table Issued -Filter "RequestID -eq $($certRequest.RequestID)"
			Receive-Certificate -RequestRow $cert -Path $Destination -Force
			
			#To-Do
			# add cert to encrypted .zip and e-mail to caller
			
		}
	}
	else
	{
		Add-Type -AssemblyName PresentationFramework
		[System.Windows.MessageBox]::Show("There are no pending certificate requests at this time on certificate authority: $($CA)")
	}
	
}
catch
{
	$errorMessage = "{0}: {1}" -f $Error[0], $Error[0].InvocationInfo.PositionMessage
	Write-Error $errorMessage -ErrorAction Continue
	$Error.Clear()
}
# SIG # Begin signature block
# MIInCQYJKoZIhvcNAQcCoIIm+jCCJvYCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCAbqp8C9DfjlLny
# z5DnV+07D13NyHp7RwNW+qEeZwCHKaCCIaUwggQVMIIC/aADAgECAgsEAAAAAAEx
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
# EzQAAAAHiSF1iXPNJ/IAAAAAAAcwDQYJKoZIhvcNAQELBQAwUjETMBEGCgmSJomT
# 8ixkARkWA2NvbTEYMBYGCgmSJomT8ixkARkWCERlbG9pdHRlMSEwHwYDVQQDExhE
# ZWxvaXR0ZSBTSEEyIExldmVsIDEgQ0EwHhcNMjAwODA1MTczMjU2WhcNMzAwODA1
# MTc0MjU2WjBUMRMwEQYKCZImiZPyLGQBGRYDY29tMRgwFgYKCZImiZPyLGQBGRYI
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
# P/0f31d3CndzQORYAb8CAwEAAaOCAZQwggGQMBAGCSsGAQQBgjcVAQQDAgECMCMG
# CSsGAQQBgjcVAgQWBBQV4b/ii/DtWsyFdE+p2v+xjwi3MzAdBgNVHQ4EFgQURy42
# 7rSc/1xeGHy4E+G+qSAeFLMwEQYDVR0gBAowCDAGBgRVHSAAMBkGCSsGAQQBgjcU
# AgQMHgoAUwB1AGIAQwBBMAsGA1UdDwQEAwIBhjASBgNVHRMBAf8ECDAGAQH/AgEB
# MB8GA1UdIwQYMBaAFL6LoCtmWVn6kHFRpaeoBkJOkDztMFgGA1UdHwRRME8wTaBL
# oEmGR2h0dHA6Ly9wa2kuZGVsb2l0dGUuY29tL0NlcnRFbnJvbGwvRGVsb2l0dGUl
# MjBTSEEyJTIwTGV2ZWwlMjAxJTIwQ0EuY3JsMG4GCCsGAQUFBwEBBGIwYDBeBggr
# BgEFBQcwAoZSaHR0cDovL3BraS5kZWxvaXR0ZS5jb20vQ2VydEVucm9sbC9TSEEy
# TFZMMUNBX0RlbG9pdHRlJTIwU0hBMiUyMExldmVsJTIwMSUyMENBLmNydDANBgkq
# hkiG9w0BAQsFAAOCAgEAh56DUZ5xeRvT/JdAM8biqLOI4PLlIhIPufRfMPJlmo64
# dY9G/JVkF7qh8SHWm7umjSTvb357kayJCks5Y3VwS11A9HMsRK11083exB27HUBd
# 2W3IyRv2KBZT+SsAsnhtb2slEuPqqrpFZC3u2RZa8XonKVVcX3wfFN0qxE+yXkjY
# MUNxr3kYuclb2kt/4/RggkfV06dL0X2lHktLMYILmr8Tb2/eU2S7//hrdcH/tcWZ
# 29hiIzL0qayp0j2MBuXACV/ZDNheEBvD659p14ae23CrgTpXSLL68RwHjaQqFVf2
# EWPXjR1MVJSvjB3QKiGdXTltUu1MBsrRHbFwj83xhiS1nTSWfIxSM+NG0u+tj9SJ
# 5fOQSEMlCe0achdoXPvF50uDwaTLxUOoBoDK2DKd8nJFa/x8/Gj35jn7RNp//Uuz
# bmIhOr7YZqdfiBwnGffm4rS577EnBSsQhuOjzujrJbJd3NP2ar293Zupr8d+QYUb
# U51ny+mUYbGQ8VQeZgo72XkFAnzx1vZw9UK5VU7pC0zlBZL/FNV6hbgcnxQ0K/qR
# AudJtx03GpNF5sqhyEC0ndvCSdljKsf4mvgNwrEDTa6HLtEKONisnLSg56IrcPx5
# W/eD8Ksodlwpfg8UM/A942V8JRZJLgFrc+nqysPi+cINMTd/n40h2wzGRlcjDE8x
# ggS6MIIEtgIBATCBgzBsMRMwEQYKCZImiZPyLGQBGRYDY29tMRgwFgYKCZImiZPy
# LGQBGRYIZGVsb2l0dGUxFjAUBgoJkiaJk/IsZAEZFgZhdHJhbWUxIzAhBgNVBAMT
# GkRlbG9pdHRlIFNIQTIgTGV2ZWwgMyBDQSAyAhNlADwUlwDjJaP1Xv9CAAAAPBSX
# MA0GCWCGSAFlAwQCAQUAoEwwGQYJKoZIhvcNAQkDMQwGCisGAQQBgjcCAQQwLwYJ
# KoZIhvcNAQkEMSIEICI5PUzxjiMtb7GPZSbXlcEHycKtOBv56Rk2PLB4wizQMA0G
# CSqGSIb3DQEBAQUABIIBALONBfMep95WQgNAHODi3cl6N6hmk26dZRbUmTcGEnI9
# xmT5BTpJK9QV9aFZjW7Du2KexodLEjKyM7KjZqEMs7lHm41QG6Y3Ref4nSnmUCOa
# JLDs6tAlWgI2Tug7yHrjocO6Fz6bf2pDrr2o/eMj/LAkHz12tzjgvYPw9QRhIsxz
# mSd/Sq6X9559VeRfcUFy6pQVslzZaefT7I9nrubHi87klip7RlkQMwADDIJK5i4h
# fchsnoE0lZfKq0jRb8gvbP3QV4Po7OAP1UFpjsYYyzleqYVUfARHxu+/HzjpMxEQ
# EzQpXtHoY9xxeDpCvUzNAfuA6TzPQGzZ51bL4W0qMtyhggK5MIICtQYJKoZIhvcN
# AQkGMYICpjCCAqICAQEwazBbMQswCQYDVQQGEwJCRTEZMBcGA1UEChMQR2xvYmFs
# U2lnbiBudi1zYTExMC8GA1UEAxMoR2xvYmFsU2lnbiBUaW1lc3RhbXBpbmcgQ0Eg
# LSBTSEEyNTYgLSBHMgIMJFS4fx4UU603+qF4MA0GCWCGSAFlAwQCAQUAoIIBDDAY
# BgkqhkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yMDA4MTgx
# NzE5MzJaMC8GCSqGSIb3DQEJBDEiBCDoTRtr1CA9bnWBQ1pxENs2qqxaXdB5n8LX
# jEuxvTF+JTCBoAYLKoZIhvcNAQkQAgwxgZAwgY0wgYowgYcEFD7HZtXU1HLiGx8h
# Q1IcMbeQ2UtoMG8wX6RdMFsxCzAJBgNVBAYTAkJFMRkwFwYDVQQKExBHbG9iYWxT
# aWduIG52LXNhMTEwLwYDVQQDEyhHbG9iYWxTaWduIFRpbWVzdGFtcGluZyBDQSAt
# IFNIQTI1NiAtIEcyAgwkVLh/HhRTrTf6oXgwDQYJKoZIhvcNAQEBBQAEggEAJh4K
# 2VjczOiQ12LOQfvqCQNGkRG8fU8UV92IyMJfHjhGFLaM/AW8EvZmHUncVpuW9aMx
# u0s53u93q2qoQFfc0568/A0p8TlxSFlZTZZ6ZChrJigpvKt+IQAUsyvXqTaKKcXB
# Oxa/9w0x8n+tnI7wxR8VTgdT7hrMCXNLYx3W+2B3eft1GvdWMsvQZlj4SOioBmK3
# jcEP8GhGWK60iEihEnmABz2WcgLYoFgw83Jia7cnAfGY/IzUVJoUkSw3g1BYzH2I
# bG2wUfDD3k2OKf1zVKXwbB8fPCn/gSnj7qfgdXRt14wlslmuLSSTcsxdPQ6AUvnb
# jm+N8QshZZnVkOXjtw==
# SIG # End signature block
