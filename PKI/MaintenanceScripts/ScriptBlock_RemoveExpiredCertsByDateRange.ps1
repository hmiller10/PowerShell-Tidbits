[CmdletBinding()]
param
(
	[Parameter(Mandatory = $true)]
	[ValidateNotNullOrEmpty()]
	[hashtable]$Arguments
)

begin
{
	try
	{
		Import-Module PSPKI -Force
	}
	catch
	{
		try
		{
			$module = Get-Module -Name PSPKI;
			$modulePath = Split-Path $module.Path;
			$psdPath = "{0}\{1}" -f $modulePath, "PSPKI.psd1"
			Import-Module $psdPath -ErrorAction Stop
		}
		catch
		{
			throw "PSPKI module could not be loaded. $($_.Exception.Message)"
		}
		
	}
	
	#Variables
	$expiredCerts = @()
	[int32]$pageSize = 100000
	$fileStamp = [datetime]::UtcNow.ToString("yyyy-MM-dd_hh-mm-ss")
	$certProps = @("RequestID", "Request.RequesterName", "CommonName", "NotBefore", "NotAfter", "SerialNumber", "CertificateTemplate")
	[String]$sMIME = 'S/MIME'
	[String]$efs = 'EFS'
	[String]$Recovery = 'Recovery'

}
process
{
	$Error.Clear()
	
	Write-Verbose -Message "Working on search of $CA"
	[int32]$pageNumber = 1
	[int32]$LastID = 0
	$StartDate = [DateTime]::ParseExact($Arguments.StartDate, "MM/dd/yyyy HH:mm:ss", $null)
	$EndDate = [DateTime]::ParseExact($Arguments.EndDate, "MM/dd/yyyy HH:mm:ss", $null)
	
	try
	{
		do
		{
			$r = 0
			Connect-CertificationAuthority -ComputerName $Arguments.CertificationAuthority -ErrorAction Stop | Out-Null
			
			if ($PSBoundParameters.ContainsKey("Filter"))
			{
				Get-IssuedRequest -CertificationAuthority $Arguments.CertficationAuthority -Filter "RequestID -gt $LastID", "NotAfter -ge $StartDate", "NotAfter -le $EndDate" -Page $pageNumber -PageSize $pageSize -ErrorAction Continue | `
				Where-Object { ((($_.CertificateTemplateOid.FriendlyName) -notlike "*$Recovery*") -and (($_.CertificateTemplateOid.FriendlyName) -notlike "*$efs*") -and (($_.CertificateTemplateOid.FriendlyName) -notlike "*$sMime*")) } | `
				ForEach-Object {
					$r++
					$LastID = $_.RequestID
					$expiredCerts += $_
				}
			}
			else
			{
				Get-IssuedRequest -CertificationAuthority $Arguments.CertficationAuthority -Filter "RequestID -gt $LastID", "NotAfter -ge $Arguments.StartDate", "NotAfter -le $Arguments.EndDate" -Page $pageNumber -PageSize $pageSize -ErrorAction Continue | `
				ForEach-Object {
					$r++
					$LastID = $_.RequestID
					$expiredCerts += $_
				}
			}
			
		}
		while ($r -eq $pageSize)
		
		Write-Warning -Message "Total number of expired certificates being removed is: $($ExpiredCerts.Count)"
		
		if ($expiredCerts.Count -gt 0)
		{
			if ((Get-Module -Name PSPKI).Version -ge 3.4)
			{
				$expiredCerts.foreach({ Remove-AdcsDatabaseRow -Request $_ })
			}
			else
			{
				$expiredCerts.foreach({ Remove-DatabaseRow -Request $_ })
			}
		}
		
		
	}
	catch
	{
		$errorMessage = "{0}: {1}" -f $Error[0], $Error[0].InvocationInfo.PositionMessage
		Write-Error $errorMessage -ErrorAction Continue
	}
	finally
	{
		[System.GC]::GetTotalMemory('forcefullcollection') | Out-Null
	}
}
end
{
	
}