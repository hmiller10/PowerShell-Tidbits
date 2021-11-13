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
$StartDate = [DateTime]::UtcNow().AddYears(-1)
$EndDate = [DateTime]::UtcNow
$CA = [System.Net.Dns]::GetHostByName("LocalHost").HostName







#Script
Connect-CertificationAuthority -ComputerName $CA

If ( $PSBoundParameters.ContainsKey("Filter") )
{
	If ( ( Get-Module -Name PSPKI).Version -ge 3.4 )
	{
		Get-IssuedRequest -CertificationAuthority $CA -Filter "NotAfter -ge $StartDate", "NotAfter -le $EndDate" | `
		Where-Object { (($_.CertificateTemplate).Contains($Recovery) -eq $false) -and (($_.CertificateTemplate).Contains($sMIME) -eq $false) -and (($_.CertificateTemplate).Contains($efs) -eq $false) } | ` 
		Remove-AdcsDatabaseRow
	}
	Else
	{
		Get-IssuedRequest -CertificationAuthority $CA -Filter "NotAfter -ge $StartDate", "NotAfter -le $EndDate" | `
		Where-Object { (($_.CertificateTemplate).Contains($Recovery) -eq $false) -and (($_.CertificateTemplate).Contains($sMIME) -eq $false) -and (($_.CertificateTemplate).Contains($efs) -eq $false) } | ` 
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