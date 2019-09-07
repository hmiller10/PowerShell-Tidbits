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
#	1.0 06/04/2019 - Initial release
#
# 
###########################################################################
[CmdletBinding()]
Param (
	[Parameter(Mandatory=$false,Position=0,ValueFromPipeline=$true,HelpMessage="Specify the fully qualified domain name of the PKI server.")]
	[String]$CA,
	
	[Parameter(Mandatory=$true,Position=1,ValueFromPipeline=$true,HelpMessage="Specify the Request.SubmittedWhen after date range. Format - ""MM/dd/yyyy HH:mm:ss""")]
	[DateTime]$StartDate,
	
	[Parameter(Mandatory=$true,Position=2,ValueFromPipeline=$true,HelpMessage="Specify the Rebuest.SubmittedWhen before date range. Format - ""MM/dd/yyyy HH:mm:ss""")]
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
$failedRequests = @()








#Script
Connect-CertificationAuthority -ComputerName $CA

[Array]$failedRequests = Get-FailedRequest -CertificationAuthority $CA -Filter "Request.SubmittedWhen -ge $StartDate", "Request.SubmittedWhen -le $EndDate"

Write-Host "Total number of failed requests for this time period:" -ForegroundColor Cyan
$failedRequests.Count

IF ( $PSBoundParameters.ContainsKey('Report') )
{
	$failedRequests | Export-Csv -Path $ReportPath -Append -NoTypeInformation
}

#End script