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

.PARAMETER Before
	EG: Gathers a list of all certificates issued prior to the defined date.

.PARAMETER Reason
	EG: Reason for revocation

.OUTPUTS
	Console output with results from script execution

.EXAMPLE
	PS> Revoke-PKICertificate.ps1 -RequestorName 'Domain\User' -Before '01/01/2020' -Reason KeyCompromise

.EXAMPLE 
	PS> Revoke-PKICertificate.ps1 -SerialNumber 'abcd1234' -CA myca.domain.com -Reason CeaseOfOperation

.EXAMPLE
	PS> Revoke-PKICertificate.ps1 -Thumbprint '1a2b3c4d' -Reason AffiliationChanged

.EXAMPLE
	PS> Revoke-PKICertificate.ps1 -RequestorName 'Domain\User' -Reason AffiliationChanged

#>

###########################################################################
#
#
# AUTHOR:  
#	Heather Miller, Manager, Identity and Access Management
#
#
# VERSION HISTORY:
# 	2.0 04/27/2020 - Added input parameter for certificate hash or thumbprint
#
# 
###########################################################################



[CmdletBinding(DefaultParameterSetName = 'RequestorName')]
Param (
	[Parameter(ParameterSetName = "RequestorName", Mandatory = $false, HelpMessage = "Enter domain\username")]
	[String]$RequestorName,
	[Parameter(ParameterSetName = "SerialNumber", Mandatory = $false, HelpMessage = "Enter serial number of certificate to be revoked.")]
	[Alias("SN")]
	[String]$SerialNumber,
	[Parameter(ParameterSetName = "CertificateHash", Mandatory = $false, HelpMessage = "Enter CertificateHash value for certificate to be revoked, no spaces.")]
	[String]$Thumbprint,
	[Parameter(ParameterSetName = "CertificateHash", Mandatory = $false, HelpMessage = "Enter the date certificates should have been issued before. EG: 01/01/2020 00:00:00")]
	[Parameter(ParameterSetName = "RequestorName")]
	[Parameter(ParameterSetName = "SerialNumber")]
	[String]$Before,
	[Parameter(ParameterSetName = "CertificateHash", Mandatory = $false, HelpMessage = "Enter FQDN of Certificate Authority. EG: myca.domain.com")]
	[Parameter(ParameterSetName = "RequestorName")]
	[Parameter(ParameterSetName = "SerialNumber")]
	[String]$CA,
	[Parameter(ParameterSetName = "CertificateHash", Mandatory = $true, HelpMessage = "Enter reason why certificate is being revoked.")]
	[Parameter(ParameterSetName = "RequestorName")]
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
$Error.Clear()

switch ($PSCmdlet.ParameterSetName)
{
	"RequestorName" {
		if (($PSBoundParameters.ContainsKey('RequestorName')) -and ($PSBoundParameters.ContainsKey('NotBefore')))
		{
			[DateTime]$sDate = Get-Date $Before
			$filter = "Request.RequesterName -eq $RequestorName", "NotBefore -le $sDate"
		}
		elseif ($PSBoundParameters.ContainsKey('RequestorName'))
		{
			$filter = "Request.RequesterName -eq $RequestorName"
		}
	}
	"SerialNumber" {
		if (($PSBoundParameters.ContainsKey('SerialNumber')) -and ($PSBoundParameters.ContainsKey('NotBefore')))
		{
			[DateTime]$sDate = Get-Date $Before
			$filter = "SerialNumber -eq $SerialNumber", "NotBefore -le $sDate"
		}
		elseif ($PSBoundParameters.ContainsKey('SerialNumber'))
		{
			$filter = "SerialNumber -eq $SerialNumber"
		}
	}
	"CertificateHash" {
		if (($PSBoundParameters.ContainsKey('Thumbprint')) -and ($PSBoundParameters.ContainsKey('Before')))
		{
			[DateTime]$sDate = Get-Date $Before
			$filter = "CertificateHash -eq $Thumbprint", "NotBefore -le $sDate"
		}
		elseif ($PSBoundParameters.ContainsKey('Thumbprint'))
		{
			$filter = "CertificateHash -eq $Thumbprint"
		}
	}
	
}


try
{
	foreach ($CA in $CAs)
	{
		try
		{
			'Filter is: {0}' -f $filter
			Write-Output ("Connecting to and searching database on $($CA) using filter: $($filter)")
			$Certs += Get-CertificationAuthority -ComputerName $CA -ErrorAction Stop | Get-IssuedRequest -Filter $filter -ErrorAction Continue
		}
		catch
		{
			$errorMessage = "{0}: {1}" -f $Error[0], $Error[0].InvocationInfo.PositionMessage
			Write-Error $errorMessage -ErrorAction Continue
		}
		
		if ($Certs.count -gt 0)
		{
			try
			{
				$Certs | foreach {
					$RequestID = $_.RequestID
					Get-IssuedRequest -CertificationAuthority $CA -Filter "RequestID -eq $RequestID" | Revoke-Certificate -Reason $Reason -RevocationDate (Get-Date)
					
					if ($?)
					{
						$RevokedRequests += Get-RevokedRequest -CertificationAuthority $CA -Filter $filter | Select-Object -Property RequestID, 'Request.RevokedWhen', 'Request.RevokedReason', CommonName, SerialNumber, CertificateTemplate
						Publish-Crl -CertificationAuthority $CA
					}
				}
				$Certs = $RequestID = $null
			}
			catch
			{
				$errorMessage = "{0}: {1}" -f $Error[0], $Error[0].InvocationInfo.PositionMessage
				Write-Error $errorMessage -ErrorAction Continue
				$Error.Clear()
			}
			finally
			{
				Write-Output $RevokedRequests
			}
			$RevokedRequests = @()
		}
		else
		{
			Write-Output ("No certificates were found related to the input parameters.")
		}
		$CA = $null
	}
	
}
catch
{
	$errorMessage = "{0}: {1}" -f $Error[0], $Error[0].InvocationInfo.PositionMessage
	Write-Error $errorMessage -ErrorAction Continue
	$Error.Clear()
}

#EndRegion