<#

.NOTES
THIS CODE IS MADE AVAILABLE AS IS, WITHOUT WARRANTY OF ANY KIND. THE
RISK OF THE USE OR THE RESULTS FROM THE USE OF THIS CODE REMAINS	WITH
THE USER.

.SYNOPSIS
    Clean expired certificates from CA database within the defined time period.

.DESCRIPTION
    This script will connect to the Certification Authority server passed
    into the script as a parameter and will utilize the StartDate value and
    EndDate value passed into the script as parameters to locate and remove
    expired certificates from the CA database, and if specified will ignore
    certificates issued from EFS, sMIME and key recovery agent templates. 
    This script will not compact or cleanup white space in the database.

.PARAMETER CA
    Fully qualified domain name of Certification Authority server

.PARAMETER StartDate
    Beginning date script should use to define the date range to search for
    expired certificates after

.PARAMETER EndDate
    End date script should use to define the date range to search for
    expired certificates before

.OUTPUTS
    Console output for number of expired certificates that will be removed along
    with the request id of the expired certificate that was removed.

.EXAMPLE 
    PS> Remove-ExpiredCertsByDateRange.ps1 -CA ca.domain.com -StartDate "01/01/2019 00:00:00" -EndDate "12/31/2019 23:59:59"

.EXAMPLE 
    PS> Remove-ExpiredCertsByDateRange.ps1 -CA ca.domain.com -StartDate "01/01/2019 00:00:00" -EndDate "12/31/2019 23:59:59" -Filter

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
# 	1.0 6/15/2019 - Initial release
#
# 
###########################################################################

[CmdletBinding()]
Param (
	[Parameter(Mandatory=$true,Position=0,ValueFromPipeline=$true,HelpMessage="Specify the fully qualified domain name of the PKI server.")]
	[String]$CA,
	
	[Parameter(Mandatory=$true,Position=1,ValueFromPipeline=$true,HelpMessage="Specify the NotBefore date range. Format - ""MM/dd/yyyy HH:mm:ss""")]
	[DateTime]$StartDate,
	
	[Parameter(Mandatory=$true,Position=2,ValueFromPipeline=$true,HelpMessage="Specify the NotAfter date range. Format - ""MM/dd/yyyy HH:mm:ss""")]
	[DateTime]$EndDate,

	[Parameter(Mandatory=$false,Position=3,HelpMessage="Switch to indicate whether or not to apply EFS/SMIME/Recovery Agent filter")]
	[Switch]$Filter
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

#Variables
$certProps = @()
$certProps = @("RequestID", "Request.RequesterName", "CommonName", "NotBefore", "NotAfter", "SerialNumber", "CertificateTemplate")
[String]$sMIME = 'S/MIME'
[String]$efs = 'EFS'
[String]$Recovery = 'Recovery'








#Script
Connect-CertificationAuthority -ComputerName $CA

If ( $PSBoundParameters.ContainsKey("Filter") )
{
	If ( ( Get-Module -Name PSPKI).Version -ge 3.4 )
	{
		Get-IssuedRequest -CertificationAuthority $CA -Filter "NotAfter -ge $StartDate", "NotAfter -le $EndDate" | `
		Where { (($_.CertificateTemplate).Contains($Recovery) -eq $false) -and (($_.CertificateTemplate).Contains($sMIME) -eq $false) -and (($_.CertificateTemplate).Contains($efs) -eq $false) } | ` 
		Remove-AdcsDatabaseRow
	}
	Else
	{
		Get-IssuedRequest -CertificationAuthority $CA -Filter "NotAfter -ge $StartDate", "NotAfter -le $EndDate" | `
		Where { (($_.CertificateTemplate).Contains($Recovery) -eq $false) -and (($_.CertificateTemplate).Contains($sMIME) -eq $false) -and (($_.CertificateTemplate).Contains($efs) -eq $false) } | ` 
		Remove-DatabaseRow
	}
	exit
}
Else
{
	If ( ( Get-Module -Name PSPKI).Version -ge 3.4 )
	{
		Get-IssuedRequest -CertificationAuthority $CA -Filter "NotAfter -ge $StartDate", "NotAfter -le $EndDate" | Remove-AdcsDatabaseRow
	}
	Else
	{
		Get-IssuedRequest -CertificationAuthority $CA -Filter "NotAfter -ge $StartDate", "NotAfter -le $EndDate" | Remove-DatabaseRow
	}
}

#End script