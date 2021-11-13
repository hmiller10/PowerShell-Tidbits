<#

.NOTES
THIS CODE IS MADE AVAILABLE AS IS, WITHOUT WARRANTY OF ANY KIND. THE
RISK OF THE USE OR THE RESULTS FROM THE USE OF THIS CODE REMAINS	WITH
THE USER.

.SYNOPSIS
    Clean failed certificate requests from CA database within the 
    defined time period.

.DESCRIPTION
    This script will connect to the Certification Authority server passed
    into the script as a parameter and will utilize the StartDate value and
    EndDate value passed into the script as parameters to locate and remove
    failed requests from the CA database. This script will not compact or
    cleanup white space in the database.

.PARAMETER CA
    Fully qualified domain name of Certification Authority server

.PARAMETER StartDate
    Beginning date script should use to define the date range to search for
    failed requests after

.PARAMETER EndDate
    End date script should use to define the date range to search for
    failed requests before

.OUTPUTS
    Console output for number of failed requests that will be removed along
    with the request id of the failed request that was removed.

.EXAMPLE 
    PS> Remove-FailedRequestsByDateRange.ps1 -CA ca.domain.com -StartDate "01/01/2019 00:00:00" -EndDate "12/31/2019 23:59:59"

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
# 	1.0 6/4/2019 - Initial release
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
	[DateTime]$EndDate
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
Connect-CertificationAuthority -ComputerName $CA


If ( ( Get-Module -Name PSPKI).Version -ge 3.4 ) 
{ 
	Get-FailedRequest -CertificationAuthority $CA -Filter "Request.SubmittedWhen -ge $StartDate", "Request.SubmittedWhen -le $EndDate" | Remove-AdcsDatabaseRow
}
Else
{
	Get-FailedRequest -CertificationAuthority $CA -Filter "Request.SubmittedWhen -ge $StartDate", "Request.SubmittedWhen -le $EndDate" | Remove-DatabaseRow 
}

#End script