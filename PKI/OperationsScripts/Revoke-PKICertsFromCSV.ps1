#Requires -Module PSPKI
#Requires -RunAsAdministrator
<#

.NOTES
	THIS CODE IS MADE AVAILABLE AS IS, WITHOUT WARRANTY OF ANY KIND. THE
	RISK OF THE USE OR THE RESULTS FROM THE USE OF THIS CODE REMAINS WITH
	THE USER.

.SYNOPSIS
	This script will revoke a PKI certificate as defined in the parameters fed
	into the script.

.DESCRIPTION
	The purpose of this script is to enable the ability to quickly revoke a 
	certificate or certificates issued by multiple certificate authorities

.PARAMETER InputFile
	EG: C:\requests
	
.PARAMETER CertificateAuthorities
	EG: ca1.domain.com
	
.OUTPUTS
	Console output with results from script execution

.EXAMPLE

	Revoccation Reasons - 'Unspecified', 'KeyCompromise', 'CACompromise', 'AffiliationChanged', 'Superseded', 'CeaseOfOperation'
	
	$Requests = @"
     User,SerialNumber,Thumbprint,Before,Reason
     US\jdoe,12345,1a2b3c4d,G/12/20 18:30:,AffiliationChanged
     UK\bsmith,,2b3c4d5e,KeyCompromise
     IN\cjackson,23456,1a3b5d7f,AffiliationChanged
     IL\rwilliams,,,

"@ | ConvertFrom-Csv
	
	PS C:\> .\Revoke-PKICertsFromCSV.ps1 -InputFile <Path\To\My.csv> -CertificateAuthorities ca1.domain.com, ca2.domain.com

#>

###########################################################################
#
#
# AUTHOR:  
#	Heather Miller
#
#
# VERSION HISTORY:
# 	3.0 02/19/2022 - Improved search filtering accuracy
#
# 
###########################################################################

[CmdletBinding()]
Param (
	[Parameter(Mandatory = $true, HelpMessage = "Enter path to input file. EG: Path\To\My.csv")]
	[String]$InputFile,
	[Parameter(Mandatory = $true, HelpMessage = "Enter the FQDN or all certificate authority servers to be searched.")]
	[Array]$CertificateAuthorities
)

#Region Modules
Try
{
	Import-Module PSPKI -Force
}
Catch
{
	Try
	{
		$modulePath = "{0}\{1}\{2}\{3}" -f $env:ProgramFiles, "WindowsPowerShell", "Modules", "PSPKI"
		$moduleVersion = (Get-Module -Name PSPKI).Version
		$strModuleVersion = $moduleVersion.ToString()
		$psdPath = "{0}\{1}\{2}" -f $modulePath, $strModuleVersion, "pspki.psd1"
		Import-Module $psdPath -ErrorAction Stop
	}
	Catch
	{
		Throw "PSPKI module could not be loaded. $($_.Exception.Message)"
	}
	
}

#EndRegion

#Region Variables

$certProps = @("RequestID", "Request.RequesterName", "Serialnumber", "CertificateHash", "NotBefore", "NotAfter")
$revokedProps = @("RequestID", "Request.RevokedWhen", "Request.RevokedReason", "NotBefore", "CommonName", "SerialNumber", "CertificateTemplate")
#EndRegion




#Region Script
$Error.Clear()

$Requests = Import-Csv -Path $InputFile -ErrorAction Stop

$obj = $Requests | Get-Member -MemberType NoteProperty | Select-Object -ExpandProperty Name
	
if (((($obj -contains 'RequestorName') -eq $false) -and ($obj -contains 'SerialNumber' -eq $false) -and ($obj -contains 'Thumbprint' -eq $false))-and ($obj -Contains 'Reason'))
{
     Write-Error -Message "A searchable row value of 'RequestorName' or 'SerialNumber' or 'Thumbprint' is missing. Reformat file to include at least one of these required fields.";break
}

If ($Requests.Count -ge 1)
{

	ForEach ($request In $Requests)
	{
		$filter = ""
		If (([String]::IsNullOrWhiteSpace($request.SerialNumber) -eq $false) -and ([String]::IsNullOrWhiteSpace($request.Before) -eq $false))
		{
			[datetime]$Before = Get-Date $request.Before
			$SerialNumber = $request.SerialNumber
			$filter = "SerialNumber -eq $SerialNumber", "NotBefore -le $Before"
		}
		ElseIf (([string]::IsNullOrWhiteSpace($request.SerialNumber) -eq $false) -and ([String]::IsNullOrWhiteSpace($request.Before) -eq $true))
		{
			$SerialNumber = $request.SerialNumber
			$filter = "SerialNumber -eq $SerialNumber"
		}
		ElseIf (([string]::IsNullOrWhiteSpace($request.SerialNumber) -eq $true) -and ([string]::IsNullOrWhiteSpace($request.Thumbprint) -eq $false) -and ([String]::IsNullOrWhiteSpace($request.Before) -eq $false))
		{
			[datetime]$Before = Get-Date $request.Before
			$Thumbprint = $request.Thumbprint
			$filter = "CertificateHash -eq $Thumbprint", "NotBefore -le $Before"
		}
		ElseIf (([string]::IsNullOrWhiteSpace($request.SerialNumber) -eq $true) -and ([string]::IsNullOrWhiteSpace($request.Thumbprint) -eq $false) -and ([String]::IsNullOrWhiteSpace($request.Before) -eq $true))
		{
			$Thumbprint = $request.Thumbprint
			$filter = "CertificateHash -eq $Thumbprint"
		}
		ElseIf (([string]::IsNullOrWhiteSpace($request.SerialNumber) -eq $true) -and ([string]::IsNullOrWhiteSpace($request.RequestorName) -eq $false) -and ([String]::IsNullOrWhiteSpace($request.Before) -eq $false))
		{
			[datetime]$Before = Get-Date $request.Before
			$User = $request.RequestorName
			$filter = "Request.RequesterName -eq $User", "NotBefore -le $Before"
		}
		Elseif (([string]::IsNullOrWhiteSpace($request.SerialNumber) -eq $true) -and ([string]::IsNullOrWhiteSpace($request.RequestorName) -eq $false) -and ([String]::IsNullOrWhiteSpace($request.Before) -eq $true))
		{
			$User = $request.RequestorName
			$filter = "Request.RequesterName -eq $User"
		}
		
		ForEach ($CA In $CertificateAuthorities)
		{
			$Certs = @()
			#Search and locate certificates to be revoked on all named Certificate Authorities
			Try
			{
				Write-Output ("Connecting to and searching database on $($CA) using filter: $filter")
				$Certs += Get-CertificationAuthority -ComputerName $CA -ErrorAction Stop | Get-IssuedRequest -Filter $filter -Property $certProps -ErrorAction Continue
			}
			Catch
			{
				$errorMessage = "{0}: {1}" -f $Error[0], $Error[0].InvocationInfo.PositionMessage
				Write-Error $errorMessage -ErrorAction Continue
			}
			
			If ($Certs.count -gt 0)
			{
				Try
				{
					#Revoke certificates		
					$Certs.foreach({
							$RequestID = $_.RequestID
							Get-IssuedRequest -CertificationAuthority $CA -Filter "RequestID -eq $RequestID" -Property $CertProps | Revoke-Certificate -Reason $request.Reason -RevocationDate (Get-Date)
							
							If ($? -eq $true)
							{
								$RevokedRequests += Get-RevokedRequest -CertificationAuthority $CA -Filter $filter | Select-Object -Property $revokedProps
								Get-CertificationAuthority -ComputerName $CA | Publish-CRL -UpdateFile
							}
							
							$null  = $RequestID
						})
					$null = $Certs
				}
				Catch
				{
					$errorMessage = "{0}: {1}" -f $Error[0], $Error[0].InvocationInfo.PositionMessage
					Write-Error $errorMessage -ErrorAction Continue
				}
				Finally
				{
					Write-Output $RevokedRequests
				}
				$RevokedRequests = @()
			}
			Else
			{
				Write-Output ("No certificates were found on $CA using the input parameters.")
			}
			$null = $CA
			
		} #end ForEach $CAs
		$null = $request = $User = $SerialNumber = $Thumbprint
	} #end ForEach $Requests
	
	
}#end If $Requests count
#EndRegion