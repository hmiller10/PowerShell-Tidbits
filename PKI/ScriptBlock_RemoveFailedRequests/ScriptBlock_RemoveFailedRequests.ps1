[CmdletBinding(SupportsShouldProcess = $true)]
[OutputType([SysadminsLV.PKI.Management.CertificateServices.Database.AdcsDbRow])]
param
(
[Parameter(Mandatory = $true)]
[ValidateNotNullOrEmpty()]
[hashtable]$Arguments
)

begin
{
	$Error.Clear()
	#Modules
	try
	{
		Import-Module -Name PSPKI -Force -ErrorAction Stop
	}
	catch
	{
		try
		{
			$moduleName = 'PSPKI'
			$ErrorActionPreference = 'Stop';
			$module = Get-Module -ListAvailable -Name $moduleName;
			$ErrorActionPreference = 'Continue';
			$modulePath = Split-Path $module.Path;
			$psdPath = "{0}\{1}" -f $modulePath, "PSPKI.psd1"
			Import-Module $psdPath -ErrorAction Stop
		}
		catch
		{
			Write-Error "PSPKI PS module could not be loaded. $($_.Exception.Message)" -ErrorAction Stop
		}
	}
	
	#Variables
	[int32]$pageSize = 100000
	[int32]$LastID = 0
	try
	{
		#https://docs.microsoft.com/en-us/dotnet/api/system.net.securityprotocoltype?view=netcore-2.0#System_Net_SecurityProtocolType_SystemDefault
		if ($PSVersionTable.PSVersion.Major -lt 6 -and [Net.ServicePointManager]::SecurityProtocol -notmatch 'Tls12')
		{
			Write-Verbose -Message 'Adding support for TLS 1.2'
			[Net.ServicePointManager]::SecurityProtocol += [Net.SecurityProtocolType]::Tls12
		}
	}
	catch
	{
		Write-Warning -Message 'Adding TLS 1.2 to supported security protocols was unsuccessful.'
	}
	
	$certProps = @("RequestID", "Request.StatusCode", "Request.CommonName", "Request.SubmittedWhen", "Request.DispositionMessage", "ConfigString", "CertificateTemplate")
	
}
process
{
	
	[int32]$pageNumber = 1
	$r = 0
	
	try
	{
		$CertificateAuthority = Connect-CertificationAuthority -ErrorAction Stop | Select-Object -ExpandProperty ComputerName
		Write-Verbose -Message ("Working on search of {0}" -f $CertficateAuthority)
	}
	catch
	{
		$errorMessage = "{0}: {1}" -f $Error[0], $Error[0].InvocationInfo.PositionMessage
		Write-Error $errorMessage -ErrorAction Stop
	}
	
	try
	{
		if (($Arguments.Keys -contains ('StartDate')) -and ($Arguments.Keys -contains ('EndDate')))
		{
			$StartDate = $Arguments.StartDate
			$EndDate = $Arguments.EndDate
			
			if ((Get-Module -ListAvailable -Name PSPKI).Version -ge 3.4)
			{
				do
				{
					if ($PSCmdlet.ShouldProcess($CertificateAuthority, "Remove Failed Requests"))
					{
						try
						{
							$CertificateAuthority | Get-FailedRequest -Filter "RequestID -gt $LastID", "Request.SubmittedWhen -ge $StartDate", "Request.SubmittedWhen -le $EndDate" -Properties $certProps -Page $pageNumber -PageSize $pageSize -ErrorAction Stop | `
							Remove-AdcsDatabaseRow -ErrorAction Stop
						}
						catch
						{
							$errorMessage = "{0}: {1}" -f $Error[0], $Error[0].InvocationInfo.PositionMessage
							Write-Error $errorMessage -ErrorAction Continue
						}
					}
					
					$pageNumber++
				}
				while ($r -eq $pageSize)
			}
			else
			{
				do
				{
					if ($PSCmdlet.ShouldProcess($CertificateAuthority, "Remove Failed Requests"))
					{
						try
						{
							$CertificateAuthority | Get-FailedRequest -Filter "RequestID -gt $LastID", "Request.SubmittedWhen -ge $StartDate", "Request.SubmittedWhen -le $EndDate" -Properties $certProps -Page $pageNumber -PageSize $pageSize -ErrorAction Stop | `
							Remove-DatabaseRow -ErrorAction Stop
						}
						catch
						{
							$errorMessage = "{0}: {1}" -f $Error[0], $Error[0].InvocationInfo.PositionMessage
							Write-Error $errorMessage -ErrorAction Continue
						}
					}
					$pageNumber++
				}
				while ($r -eq $pageSize)
			}
		}
		else
		{
			if ((Get-Module -ListAvailable -Name PSPKI).Version -ge 3.4)
			{
				do
				{
					if ($PSCmdlet.ShouldProcess($CertificateAuthority, "Remove Failed Request"))
					{
						try
						{
							$CertificateAuthority | Get-FailedRequest -Filter "RequestID -gt $LastID" -Properties $certProps -Page $pageNumber -PageSize $pageSize -ErrorAction Stop | `
							Remove-AdcsDatabaseRow -ErrorAction Stop
						}
						catch
						{
							$errorMessage = "{0}: {1}" -f $Error[0], $Error[0].InvocationInfo.PositionMessage
							Write-Error $errorMessage -ErrorAction Continue
						}
					}
					$pageNumber++
				}
				while ($r -eq $pageSize)
			}
			else
			{
				do
				{
					if ($PSCmdlet.ShouldProcess($CertificateAuthority, "Remove Failed Request"))
					{
						try
						{
							$CertificateAuthority | Get-FailedRequest -Filter "RequestID -gt $LastID" -Properties $certProps -Page $pageNumber -PageSize $pageSize -ErrorAction Stop | `
							Remove-DatabaseRow -ErrorAction Stop
						}
						catch
						{
							$errorMessage = "{0}: {1}" -f $Error[0], $Error[0].InvocationInfo.PositionMessage
							Write-Error $errorMessage -ErrorAction Continue
						}
					}
					$pageNumber++
				}
				while ($r -eq $pageSize)
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