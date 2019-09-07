<#

.NOTES
THIS CODE IS MADE AVAILABLE AS IS, WITHOUT WARRANTY OF ANY KIND. THE
RISK OF THE USE OR THE RESULTS FROM THE USE OF THIS CODE REMAINS
THE USER.

.SYNOPSIS 
	Remove all X.509 PKI certificates from computer LocalMachine
	certificate store

.DESCRIPTION 
	Remove all X.509 PKI certificates from computer LocalMachine
	certificate store, issued by the named template, using .Net class. This
	script is designed to be run on the local computer itself.

.OUTPUTS 
	Console displays list of certificates deleted by serial number

.EXAMPLE 
    PS> Remove-ConfigMgrCerts.ps1 -TemplateName "<Name of CA template including spaces>"


#>

[CmdletBinding()]
Param (
	[Parameter(Mandatory=$false,Position=0,ValueFromPipeline=$true,HelpMessage="Specify the full name of the certificate template including spaces. EG: ""Web Server - 2 years""")]
	[String]$Template
)

Function Remove-CertByTemplate
{
	# Pass Serial Number of the cert you want to remove
	Param ($TemplateName = $(throw "Please pass a valid certificate name to the script including spaces   "))

	Begin {
		# Access MY store of Local Machine profile 
		$store = New-Object System.Security.Cryptography.X509Certificates.X509Store("My","LocalMachine")
		$store.Open("ReadWrite")

		# Find the cert we want to delete
		$cert = $store.Certificates.Find("FindByTemplateName",$TemplateName,$FALSE)[0]
	}
	Process {
		If ($cert -ne $null)
		{
			# Found the cert. Delete it (need admin permissions to do this)
			$store.Remove($cert)

			Write-Host "Certificate issued by Template $TemplateName has been deleted"
			$FOUND = 1
			return $FOUND
		}
		Else
		{
			# Didn't find the cert. Exit
			Write-Host "Certificate issued by Template $TemplateName could not be found"
			$FOUND = 0
			return $FOUND
		}
	}
	End {
		# We are done
		$store.Close()
	}
}#End function Remove-CertByTemplate


#Variables
$FOUND = 1

While ($FOUND) {
	$FOUND = Remove-CertByTemplate -TemplateName $Template
}