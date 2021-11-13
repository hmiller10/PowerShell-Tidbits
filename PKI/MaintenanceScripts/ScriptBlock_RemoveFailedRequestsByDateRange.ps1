[CmdletBinding()]
param
(
	[Parameter(Mandatory = $true)]
	[ValidateNotNullOrEmpty()]
	[hashtable]$Arguments
)

begin
{
	Try
	{
		Import-Module PSPKI -Force
	}
	Catch
	{
		Try
		{
			 $module = Get-Module -Name PSPKI;
			 $modulePath = Split-Path $module.Path;
			 $psdPath = "{0}\{1}" -f $modulePath, "PSPKI.psd1"
			 Import-Module $psdPath -ErrorAction Stop
		}
		Catch
		{
			Throw "PSPKI module could not be loaded. $($_.Exception.Message)"
		}
		
	}
	
	#Variables
	$FailedRequests = @()
	[int32]$pageSize = 100000
	$fileStamp = [datetime]::UtcNow.ToString("yyyy-MM-dd_hh-mm-ss")

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
			Connect-CertificationAuthority -ComputerName $Arguments.CertficationAuthority -ErrorAction Stop | Out-Null
			[Array]$FailedRequests += Get-FailedRequest -CertificationAuthority $Arguments.CertficationAuthority -Filter "Request.SubmittedWhen -ge $StartDate", "Request.SubmittedWhen -le $EndDate" -PageSize $pageSize
			ForEach-Object {
				$FailedRequests += $_
				$r++
			}
			$pageNumber++
			
		}
		while ($r -eq $pageSize)
		
		Write-Warning -Message "Total number of failed requests being removed is: $($FailedRequests.Count)"
		if ((Get-Module -Name PSPKI).Version -ge 3.4)
		{
			$FailedRequests.foreach({ Remove-AdcsDatabaseRow -Request $_ })
		}
		else
		{
			$FailedRequests.foreach({ Remove-DatabaseRow -Request $_ })
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