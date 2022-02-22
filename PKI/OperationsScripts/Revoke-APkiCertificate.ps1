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
# 	3.0 10/25/2021 - Updated CA connection method and processing loop
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
	[Parameter(ParameterSetName = "RequestorName", Mandatory = $false, HelpMessage = "Enter the date certificates should have been issued before. EG: 01/01/2020 00:00:00")]
	[Parameter(ParameterSetName = "SerialNumber")]
	[Parameter(ParameterSetName = "CertificateHash")]
	[DateTime]$Before,
	[Parameter(ParameterSetName = "RequestorName", Mandatory = $false, HelpMessage = "Enter FQDN of Certificate Authority. EG: myca.domain.com")]
	[Parameter(ParameterSetName = "SerialNumber")]
	[Parameter(ParameterSetName = "CertificateHash")]
	[Array]$CertificateAuthorities,
	[Parameter(ParameterSetName = "RequestorName", Mandatory = $true, HelpMessage = "Enter reason why certificate is being revoked.")]
	[Parameter(ParameterSetName = "SerialNumber")]
	[Parameter(ParameterSetName = "CertificateHash")]
	[ValidateSet('Unspecified', 'KeyCompromise', 'CACompromise', 'AffiliationChanged', 'Superseded', 'CeaseOfOperation')]
	[String]$Reason
)

#Region Modules
try
{
	Import-Module PSPKI -Force
}
catch
{
	try
	{
		$modulePath = "{0}\{1}\{2}\{3}" -f $env:ProgramFiles, "WindowsPowerShell", "Modules", "PSPKI"
		$moduleVersion = (Get-Module -Name PSPKI).Version
		$strModuleVersion = $moduleVersion.ToString()
		$psdPath = "{0}\{1}\{2}" -f $modulePath, $strModuleVersion, "pspki.psd1"
		Import-Module $psdPath -ErrorAction Stop
	}
	catch
	{
		throw "PSPKI module could not be loaded. $($_.Exception.Message)"
	}
	
}

#EndRegion

#Region Variables

#EndRegion




	
	
	
#Region Script
$Error.Clear()

switch ($PSCmdlet.ParameterSetName)
{
	"RequestorName" {
		if ($PSBoundParameters.ContainsKey("Before"))
		{
			$filter = "Request.RequesterName -eq $RequestorName","NotBefore -le $Before";break
		}
		else
		{
			$filter = "Request.RequesterName -eq $RequestorName";break
		}
	}
	"SerialNumber" {
		if ($PSBoundParameters.ContainsKey("Before"))
		{
			$filter = "SerialNumber -eq $SerialNumber", "NotBefore -le $Before";break
		}
		else
		{
			$filter = "SerialNumber -eq $SerialNumber" ;break
		}
	}
	"CertificateHash" {
		if ($PSBoundParameters.ContainsKey("Before"))
		{
			$filter = "CertificateHash -eq $Thumbprint", "NotBefore -le $Before";break
		}
		else
        {
			$filter = "CertificateHash -eq $Thumbprint";break
		}
	}
	
}


try
{
	$Certs = @()
	foreach ($CA in $CertificateAuthorities)
	{
		try
		{
			Write-Output ("Filter is: {0}" -f $filter)
			Write-Output ("Connecting to and searching database on {0} using filter: {1}" -f $CA, $filter)
			$Certs += Get-CertificationAuthority -ComputerName $CA -ErrorAction Stop | Get-IssuedRequest -Filter $filter -ErrorAction Continue
			
		}
		Catch
		{
			$errorMessage = "{0}: {1}" -f $Error[0], $Error[0].InvocationInfo.PositionMessage
			Write-Error $errorMessage -ErrorAction Continue
		}
		
		If ($Certs.count -gt 0)
		{
			$RevokedRequests = @()
			try
			{
				$Certs.foreach({ Write-Output $_ })
#				$Certs.foreach({
#					$RequestID = $_.RequestID
#						
#					Get-IssuedRequest -CertificationAuthority $CertificateAuthority.ComputerName -Filter "RequestID -eq $RequestID" | Revoke-Certificate -Reason $Reason -RevocationDate (Get-Date)
#					
#					if ($?)
#					{
#						$RevokedRequests += Get-RevokedRequest -CertificationAuthority $CertificateAuthority.ComputerName -RequestID $RequestID | Select-Object -Property RequestID, 'Request.RevokedWhen', 'Request.RevokedReason', CommonName, SerialNumber, CertificateTemplate
#						Get-CertificationAuthority -ComputerName $CertificateAuthority.ComputerName | Publish-CRL -UpdateFile
#					}
#					$null = $RequestID
#				})
#				$null = $Certs
			}
			catch
			{
				$errorMessage = "{0}: {1}" -f $Error[0], $Error[0].InvocationInfo.PositionMessage
				Write-Error $errorMessage -ErrorAction Continue
				$Error.Clear()
			}
			finally
			{
				If ($RevokedRequests.Count -gt 0)
				{
					Write-Output $RevokedRequests
				}
				$RevokedRequests = @()
			}
			
		}
		else
		{
			Write-Output ("No certificates were found based on the input parameters.")
		}
		$CA = $null
	}
	
}
catch
{
	$errorMessage = "{0}: {1}" -f $Error[0], $Error[0].InvocationInfo.PositionMessage
	Write-Error $errorMessage -ErrorAction Continue
}

#EndRegion