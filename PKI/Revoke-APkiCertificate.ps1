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

.PARAMETER RequestorName
	EG: domain\username

.PARAMETER SerialNumber
	EG: actual serial number of certificate

.PARAMETER CA
	EG: CA FQDN

.PARAMETER Reason
	EG: Reason for revocation

.OUTPUTS
	Console output with results from script execution

.EXAMPLE 
	PS> Revoke-PKICertificate.ps1 -SerialNumber 'abcd1234' -CA myca.domain.com -Reason CeaseOfOperation

.EXAMPLE
	PS> Revoke-PKICertificate.ps1 -RequestorName 'Domain\User' -Reason AffiliationChanged

#>


[CmdletBinding()]
Param (
	[Parameter(ParameterSetName = "RequestorName", Mandatory = $false, HelpMessage = "Enter domain\username")]
	[String]$RequestorName,
	[Parameter(ParameterSetName = "SerialNumber", Mandatory = $false, HelpMessage = "Enter serial number of certificate to be revoked.")]
	[Alias("SN")]
	[String]$SerialNumber,
	[Parameter(ParameterSetName = "RequestorName", Mandatory = $false, HelpMessage = "Enter FQDN of Certificate Authority. EG: myca.domain.com")]
	[Parameter(ParameterSetName = "SerialNumber")]
	[String]$CA,
	[Parameter(ParameterSetName = "RequestorName", Mandatory = $true, HelpMessage = "Enter reason why certificate is being revoked.")]
	[Parameter(ParameterSetName = "SerialNumber")]
	[ValidateSet('Unspecified', 'KeyCompromise', 'CACompromise', 'AffiliationChanged', 'Superseded', 'CeaseOfOperation')]
	[String]$Reason
)

#Region Modules

Try
{
	Import-Module PSPKI -ErrorAction Stop
}
Catch
{
	Try
	{
		Import-Module "C:\Program Files\WindowsPowerShell\Modules\PSPKI.psd1" -ErrorAction Stop
	}
	Catch
	{
		Throw "PSPKI module could not be loaded. $($_.Exception.Message)"
	}
	
}

#EndRegion

#Region Variables

if ($PSBoundParameters.ContainsKey('CA'))
{
	[Array]$CAs = @($CA)
}
else
{
	$ca1 = 'ca1.domain.com'
	$ca2 = 'ca1.domain.com'
	$ca3 = 'ca1.domain.com'
	$ca4 = 'ca1.domain.com'
	
	[Array]$CAs = @($ca1, $ca2, $ca3, $ca4)
}


#EndRegion






#Region Script

if ($PSBoundParameters.ContainsKey('RequesterName'))
{
	$filter = "Request.RequesterName -eq $RequestorName"
}

if ($PSBoundParameters.ContainsKey('SerialNumber'))
{
	$filter = "SerialNumber -eq $SerialNumber"
}


try
{
	foreach ($CA in $CAs)
	{
		Connect-CA -ComputerName $CA
		[Array]$Certs += Get-IssuedRequest -CertificationAuthority $CA -Filter $filter
		
		if ($Certs.Count -eq 1)
		{
			try
			{
				$RequestID = $Certs.RequestID
				Get-IssuedRequest -CertificationAuthority $CA -filter "RequestID -eq $RequestID" | Revoke-Certificate -Reason $Reason -RevocationDate (Get-Date)
				
				if ($?)
				{
					Get-RevokedRequest -CertificationAuthority $CA -Filter $filter | Select-Object -Property RequestID, 'Request.RevokedWhen', 'Request.RevokedReason', CommonName, SerialNumber, CertificateTemplate
					Get-CertificationAuthority -ComputerName $CA | Publish-CRL -DeltaOnly
				}
			}
			catch
			{
				$errorMessage = "{0}: {1}" -f $Error[0], $Error[0].InvocationInfo.PositionMessage
				Write-Error $errorMessage -ErrorAction Continue
			}
		}
		else
		{
			
			try
			{
				$Certs | foreach {
					$RequestID = $_.RequestID
					Get-IssuedRequest -CertificationAuthority $CA -filter "RequestID -eq $RequestID" | Revoke-Certificate -Reason $Reason -RevocationDate (Get-Date)
					
					if ($?)
					{
						Get-RevokedRequest -CertificationAuthority $CA -Filter $filter | Select-Object -Property RequestID, 'Request.RevokedWhen', 'Request.RevokedReason', CommonName, SerialNumber, CertificateTemplate
						Get-CertificationAuthority -ComputerName $CA | Publish-CRL -DeltaOnly
					}
				}
			}
			catch
			{
				$errorMessage = "{0}: {1}" -f $Error[0], $Error[0].InvocationInfo.PositionMessage
				Write-Error $errorMessage -ErrorAction Continue
			}
			
		}
		$CA = $Certs = $RequestID = $null
	}
}
catch
{
	$errorMessage = "{0}: {1}" -f $Error[0], $Error[0].InvocationInfo.PositionMessage
	Write-Error $errorMessage -ErrorAction Continue
}


#EndRegion