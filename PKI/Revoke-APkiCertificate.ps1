<#

.NOTES
	THIS CODE IS MADE AVAILABLE AS IS, WITHOUT WARRANTY OF ANY KIND. THE
	RISK OF THE USE OR THE RESULTS FROM THE USE OF THIS CODE REMAINS WITH
	THE USER.

.SYNOPSIS
	This script will connect to a CA, revoke a PKI certificate as defined in the parameters fed
	into the script, and update the CRL to reflect the revoked certificate.

.DESCRIPTION
	The purpose of this script is to enable the ability to quickly revoke a 
	certificate or certificates issued by by the defined certificate authority
	and to subsequently update the CRL for that CA

.PARAMETER RequestorName
	EG: domain\username

.PARAMETER sn
	EG: actual serial number of certificate

.OUTPUTS
	Console output of results from script execution

.EXAMPLE 
	PS> Revoke-APkiCertificate.ps1 -SN 'abcd1234' -Reason CeaseOfOperation

.EXAMPLE
	PS> Revoke-APkiCertificate.ps1 -RequestorName 'Domain\User' -Reason AffiliationChanged

#>
[CmdletBinding()]
    Param (
		[Parameter(
		ParameterSetName="RequestorName",
		Mandatory = $true,
		ValueFromPipeline = $true
		)]
		[String]
		$RequestorName,

		[Parameter(
		ParameterSetName="SerialNumber",
		Mandatory = $true,
		ValueFromPipeline = $true
		)]
		[String]
		$SN,

		[Parameter(Mandatory)]  
		[ValidateSet('Unspecified','KeyCompromise','CACompromise','AffiliationChanged','Superseded','CeaseOfOperation')]
		[String]
		$Reason
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
	    Import-Module "C:\Program Files\WindowsPowerShell\Modules\PSPKI.psd1" -ErrorAction Stop
	}
	Catch
	{
	   Throw "PSPKI module could not be loaded. $($_.Exception.Message)"
	}
	
}


#Script
Connect-CA -ComputerName $CA
if ( $RequestorName )
{
	$filter = "Request.RequesterName -eq $RequestorName"
}
elseif ( $SN )
{
	$filter = "SerialNumber -eq $sn"
}

[Array]$Certs += Get-IssuedRequest -CertificationAuthority $CA -Filter $filter

if ($Certs.Count -eq 1)
{
	$RequestID = $Certs.RequestID
	Get-IssuedRequest -CertificationAuthority $CA -filter "RequestID -eq $RequestID" | Revoke-Certificate -Reason $Reason -RevocationDate (Get-Date)

	if ($?)
	{
    		Get-RevokedRequest -CertificationAuthority $CA -Filter $filter | Select-Object -Property RequestID, 'Request.RevokedWhen', 'Request.RevokedReason', CommonName, SerialNumber, CertificateTemplate
		Get-CertificationAuthority -ComputerName $CA | Publish-CRL -DeltaOnly
	}
}
else
{
	$Certs | foreach {
		$RequestID = $_.RequestID
		Get-IssuedRequest -CertificationAuthority $CA -filter "RequestID -eq $RequestID" | Revoke-Certificate -Reason $reason -RevocationDate (Get-Date)
		
		if ($?)
    		{
        		Get-RevokedRequest -CertificationAuthority $CA -Filter $filter | Select-Object -Property RequestID, 'Request.RevokedWhen', 'Request.RevokedReason', CommonName, SerialNumber, CertificateTemplate
			Get-CertificationAuthority -ComputerName $CA | Publish-CRL -DeltaOnly
    		}
	}
}