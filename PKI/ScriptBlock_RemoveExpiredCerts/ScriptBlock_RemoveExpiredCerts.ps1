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
	$certProps = @("RequestID", "Request.RequesterName", "CommonName", "NotBefore", "NotAfter", "SerialNumber", "CertificateTemplate")
	[String]$sMIME = 'S/MIME'
	[String]$efs = 'EFS'
	[String]$Recovery = 'Recovery'
	
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
	
}
process
{
	$Error.Clear()
	[int32]$pageNumber = 1
	[int32]$LastID = 0
	
	try
	{
		try
		{
			if ($Arguments.Keys -contains ('CertificateAuthority'))
			{
				$CertificateAuthority = Connect-CertificationAuthority -ComputerName $Arguments.CertificateAuthority -ErrorAction Stop	
			}
			else
			{
				$CertificateAuthority = Connect-CertificationAuthority -ErrorAction Stop
			}
		}
		catch
		{
			$errorMessage = "{0}: {1}" -f $Error[0], $Error[0].InvocationInfo.PositionMessage
			Write-Error $errorMessage -ErrorAction Stop
		}
		
		if (($Arguments.Keys -contains ('UseFilter')) -and ($Arguments.Keys -contains ('StartDate')) -and ($Arguments.Keys -contains ('EndDate')))
		{
			$StartDate = $Arguments.StartDate
			$EndDate = $Arguments.EndDate
			
			if ((Get-Module -ListAvailable -Name PSPKI).Version -ge 3.4)
			{
				do
				{
					try
					{
						$CertificateAuthority | Get-IssuedRequest -Filter "RequestID -gt $LastID", "NotAfter -ge $StartDate", "NotAfter -le $EndDate" -Properties $certProps -Page $pageNumber -PageSize $pageSize -ErrorAction Stop | `
						Where-Object { ((($_.CertificateTemplateOid.FriendlyName) -notlike "*$Recovery*") -and (($_.CertificateTemplateOid.FriendlyName) -notlike "*$efs*") -and (($_.CertificateTemplateOid.FriendlyName) -notlike "*$sMime*")) } | `
						Select-Object $certProps | Remove-ADCSDatabaseRow
						
					}
					catch
					{
						$errorMessage = "{0}: {1}" -f $Error[0], $Error[0].InvocationInfo.PositionMessage
						Write-Error $errorMessage -ErrorAction Continue
					}
					$r++
					$pageNumber++
				}
				while ($r -eq $pageSize)
			}
			else
			{
				do
				{
					try
					{
						$CertificateAuthority | Get-IssuedRequest -Filter "RequestID -gt $LastID", "NotAfter -ge $StartDate", "NotAfter -le $EndDate" -Properties $certProps -Page $pageNumber -PageSize $pageSize -ErrorAction Stop | `
						Where-Object { ((($_.CertificateTemplateOid.FriendlyName) -notlike "*$Recovery*") -and (($_.CertificateTemplateOid.FriendlyName) -notlike "*$efs*") -and (($_.CertificateTemplateOid.FriendlyName) -notlike "*$sMime*")) } | `
						Select-Object $certProps | Remove-DatabaseRow
						
					}
					catch
					{
						$errorMessage = "{0}: {1}" -f $Error[0], $Error[0].InvocationInfo.PositionMessage
						Write-Error $errorMessage -ErrorAction Continue
					}
					$r++
					$pageNumber++
				}
				while ($r -eq $pageSize)
			}
		}
		elseif ((-not ($Arguments.Keys -contains ('UseFilter'))) -and ($Arguments.Keys -contains ('StartDate')) -and ($Arguments.Keys -contains ('EndDate')))
		{
			if ((Get-Module -ListAvailable -Name PSPKI).Version -ge 3.4)
			{
				do
				{
					try
					{
						$CertificateAuthority | Get-IssuedRequest -Filter "RequestID -gt $LastID", "NotAfter -ge $StartDate", "NotAfter -le $EndDate" -Properties $certProps -Page $pageNumber -PageSize $pageSize -ErrorAction Stop | `
						Select-Object $certProps | Remove-AdcsDatabaseRow
					}
					catch
					{
						$errorMessage = "{0}: {1}" -f $Error[0], $Error[0].InvocationInfo.PositionMessage
						Write-Error $errorMessage -ErrorAction Continue
					}
					
					$r++
					$pageNumber++
				}
				while ($r -eq $pageSize)
			}
			else
			{
				do
				{
					try
					{
						$CertificateAuthority | Get-IssuedRequest -Filter "RequestID -gt $LastID", "NotAfter -ge $StartDate", "NotAfter -le $EndDate" -Properties $certProps -Page $pageNumber -PageSize $pageSize -ErrorAction Stop | `
						Select-Object $certProps | Remove-DatabaseRow
					}
					catch
					{
						$errorMessage = "{0}: {1}" -f $Error[0], $Error[0].InvocationInfo.PositionMessage
						Write-Error $errorMessage -ErrorAction Continue
					}
					
					$r++
					$pageNumber++
				}
				while ($r -eq $pageSize)
			}
			
		} #end else no template filter, only date filter
		else
		{
			if ((Get-Module -ListAvailable -Name PSPKI).Version -ge 3.4)
			{
				do
				{
					try
					{
						$CertificateAuthority | Get-IssuedRequest -Filter "RequestID -gt $LastID" -Properties $certProps -Page $pageNumber -PageSize $pageSize -ErrorAction Stop | `
						Select-Object $certProps | Remove-AdcsDatabaseRow
						
						$r++
						
					}
					catch
					{
						$errorMessage = "{0}: {1}" -f $Error[0], $Error[0].InvocationInfo.PositionMessage
						Write-Error $errorMessage -ErrorAction Continue
					}
					$pageNumber++
				}
				while ($r -eq $pageSize)
			}
			else
			{
				do
				{
					try
					{
						$CertificateAuthority | Get-IssuedRequest -Filter "RequestID -gt $LastID" -Properties $certProps -Page $pageNumber -PageSize $pageSize -ErrorAction Stop | `
						Select-Object $certProps | Remove-DatabaseRow
						
						$r++
						
					}
					catch
					{
						$errorMessage = "{0}: {1}" -f $Error[0], $Error[0].InvocationInfo.PositionMessage
						Write-Error $errorMessage -ErrorAction Continue
					}
					$pageNumber++
				}
				while ($r -eq $pageSize)
			}
			
			
		} #end else no filter
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