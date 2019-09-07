<#

.NOTES
THIS CODE IS MADE AVAILABLE AS IS, WITHOUT WARRANTY OF ANY KIND. THE
RISK OF THE USE OR THE RESULTS FROM THE USE OF THIS CODE REMAINS	WITH
THE USER.

.SYNOPSIS
    Collect list of expired certificates from CA database within the 
    defined time period.

.DESCRIPTION
    This script will connect to the Certification Authority server passed
    into the script as a parameter and will utilize the StartDate value and
    EndDate value passed into the script as parameters to locate and report
    on the number of expired certificates within the CA database. This script 
    will not compact or cleanup white space in the database.

.PARAMETER CA
    Fully qualified domain name of Certificate Authority server

.PARAMETER StartDate
    Beginning date script should use to define the date range to search for
    expired certs after

.PARAMETER EndDate
    End date script should use to define the date range to search for
    expired certs before

.PARAMETER Report
	Switch: indicates that script should output results of search to a file

.PARAMETER ReportPath
	Path to where script will export CSV file of all expired certificates found 
	during search of date range defined as input parameters

.OUTPUTS
    Console output for number of expired certificates that will be removed and
    request id of the expired certificate that was removed.

.EXAMPLE 
    PS> Get-ExpiredCertsByDateRange.ps1 -CA ca.domain.com -StartDate "01/01/2019 00:00:00" -EndDate "12/31/2019 23:59:59"

.EXAMPLE 
    PS> Get-ExpiredCertsByDateRange.ps1 -CA ca.domain.com -StartDate "01/01/2019 00:00:00" -EndDate "12/31/2019 23:59:59" -Report -ReportPath <FullPathToFile>

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
# 	1.0 6/14/2019 - Initial release
#
# 
###########################################################################

[CmdletBinding()]
Param (
	[Parameter(Mandatory=$false,Position=0,ValueFromPipeline=$true,HelpMessage="Specify the fully qualified domain name of the PKI server.")]
	[String]$CA,
	
	[Parameter(Mandatory=$true,Position=1,HelpMessage="Specify the NotBefore date range. Format - ""MM/dd/yyyy HH:mm:ss""")]
	[DateTime]$StartDate,
	
	[Parameter(Mandatory=$true,Position=2,HelpMessage="Specify the NotAfter date range. Format - ""MM/dd/yyyy HH:mm:ss""")]
	[DateTime]$EndDate,
	
	[Parameter(Mandatory=$false,Position=3,HelpMessage="Switch that specifies whether or not to export results of search to CSV")]
	[Parameter(ParameterSetName = 'ReportParameterSet')]
	[ValidateNotNullOrEmpty()]
	[Switch]$Report,
	
	[Parameter(Mandatory=$false,Position=4,HelpMessage="Specify the file path where the report results should be sent to")]
	[Parameter(ParameterSetName = 'ReportParameterSet')]
	[ValidateNotNullOrEmpty()]
	[String]$ReportPath
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
$expiredCerts = @()
$certProps = @()
$certProps = @("RequestID", "Request.RequesterName", "CommonName", "NotBefore", "NotAfter", "SerialNumber", "CertificateTemplate")







#Script
Connect-CertificationAuthority -ComputerName $CA

[Array]$expiredCerts = Get-IssuedRequest -CertificationAuthority $CA -Filter "NotAfter -ge $StartDate", "NotAfter -le $EndDate" | Select-Object -Property $certProps

Write-Host "Total number of expired certificates for this time period:" -ForegroundColor Cyan
$expiredCerts.Count

IF ( $PSBoundParameters.ContainsKey('Report') )
{
	$expiredCerts | Export-Csv -Path $ReportPath -Append -NoTypeInformation
}

#End script