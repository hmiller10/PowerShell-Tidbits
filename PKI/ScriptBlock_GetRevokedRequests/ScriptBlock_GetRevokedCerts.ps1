[CmdletBinding()]
param
(
	[Parameter(Mandatory = $false)]
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
	$certProps = @()
	$certProps = @("RequestID", "Request.RequesterName", "CommonName", "ConfigString", "Request.RevokedWhen", "Request.RevokedReason", "SerialNumber", "CertificateTemplate", "CertificateTemplateOID")
	[int32]$pageSize = 50000
	[int32]$pageNumber = 1
	[int32]$LastID = 0
	[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
	$r = 0

	$dtHeaders = ConvertFrom-Csv @"
		ColumnName,DataType
          CertificateAuthority,string
		RequestID,string
		RequestorName,string
		CommonName,string
		ConfigString,string
		RevocationDate,datetime
		RevocationReason,string
          SerialNumber,string
          CertificateTemplate,string
"@
	
	$dt = New-Object System.Data.DataTable
	
	foreach ($header in $dtHeaders)
	{
		[void]$dt.Columns.Add([System.Data.DataColumn]$header.ColumnName.ToString(), $header.DataType)
	}
}

process
{
	$Error.Clear()

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
			Write-Verbose -Message ("Working on search of {0}" -f $CertificateAuthority.ComputerName)
		}
		catch
		{
			$errorMessage = "{0}: {1}" -f $Error[0], $Error[0].InvocationInfo.PositionMessage
			Write-Error $errorMessage -ErrorAction Stop
			break;
		}
		
		Write-Verbose -Message ("Working on search of {0}" -f $Arguments.CertificateAuthority)
			
		if (($Arguments.Keys -Contains ('StartDate')) -and ($Arguments.Keys -Contains ('EndDate')))
		{
			$StartDate = $Arguments.StartDate
			$EndDate = $Arguments.EndDate
			
			do
			{
				$CertificateAuthority | Get-RevokedRequest -Filter "RequestID -gt $LastID", "Request.RevokedWhen -ge $StartDate", "Request.RevokedWhen -le $EndDate" -Properties $certProps -Page $pageNumber -PageSize $pageSize -ErrorAction Continue | `
				Select-Object -Property $certProps | ForEach-Object {
					$LastID = $_.RequestID; $r++
					
					[string]$CertificateAuthority = $Arguments.CertificationAuthority
					[string]$RequestID = $_.RequestID
					[string]$RequestorName = $_."Request.RequesterName"
					[string]$CommonName = $_.CommonName
					[string]$ConfigString = $_.ConfigString
					$RevocationDate = $_."Request.RevokedWhen"
					
					switch ($_."Request.RevokedReason")
					{
						0 { [string]$Reason = "Unspecified: Certificate was revoked for unspecified reason" }
						1 { [string]$Reason = "Key Compromise: Private key of the certificate was compromised" }
						2 { [string]$Reason = "CA Compromise: Private key of the certificate CA was compromized" }
						3 { [string]$Reason = "Affiliation Changed: Certificate owner changed affiliation" }
						4 { [string]$Reason = "Superseded: A new certificate was issued to replace this one" }
						5 { [string]$Reason = "Cessation of Operation: Certificate owner stopped operation" }
						6 { [string]$Reason = "Certificate Hold: Certificate was put on hold" }
					}
					
					[string]$serialNumber = $_.SerialNumber
					[string]$certificateTemplate = $_.CertificateTemplateOID.FriendlyName
					
					$dr = $dt.NewRow()
					
					$dr["CertificateAuthority"] = $CertificateAuthority
					$dr["RequestID"] = $RequestID
					$dr["RequestorName"] = $RequestorName
					$dr["CommonName"] = $CommonName
					$dr["ConfigString"] = $ConfigString
					$dr["RevocationDate"] = $RevocationDate
					$dr["RevocationReason"] = $Reason
					$dr["SerialNumber"] = $SerialNumber
					$dr["CertificateTemplate"] = $certificateTemplate
					
					$dt.Rows.Add($dr)
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
					$CertificateAuthority | Get-RevokedRequest -Filter "RequestID -gt $LastID" -Properties $certProps -Page $pageNumber -PageSize $pageSize -ErrorAction Stop | `
					Select-Object -Property $certProps | ForEach-Object {
						$LastID = $_.RequestID; $r++
						
						[string]$CertificateAuthority = $Arguments.CertificationAuthority
						[string]$RequestID = $_.RequestID
						[string]$RequestorName = $_."Request.RequesterName"
						[string]$CommonName = $_.CommonName
						[string]$ConfigString = $_.ConfigString
						$RevocationDate = $_."Request.RevokedWhen"
						
						switch ($_."Request.RevokedReason")
						{
							0 { [string]$Reason = "Unspecified: Certificate was revoked for unspecified reason" }
							1 { [string]$Reason = "Key Compromise: Private key of the certificate was compromised" }
							2 { [string]$Reason = "CA Compromise: Private key of the certificate CA was compromized" }
							3 { [string]$Reason = "Affiliation Changed: Certificate owner changed affiliation" }
							4 { [string]$Reason = "Superseded: A new certificate was issued to replace this one" }
							5 { [string]$Reason = "Cessation of Operation: Certificate owner stopped operation" }
							6 { [string]$Reason = "Certificate Hold: Certificate was put on hold" }
						}
						
						[string]$serialNumber = $_.SerialNumber
						[string]$certificateTemplate = $_.CertificateTemplateOID.FriendlyName
						
						$dr = $dt.NewRow()
						
						$dr["CertificateAuthority"] = $CertificateAuthority
						$dr["RequestID"] = $RequestID
						$dr["RequestorName"] = $RequestorName
						$dr["CommonName"] = $CommonName
						$dr["ConfigString"] = $ConfigString
						$dr["RevocationDate"] = $RevocationDate
						$dr["RevocationReason"] = $Reason
						$dr["SerialNumber"] = $SerialNumber
						$dr["CertificateTemplate"] = $certificateTemplate
						
						$dt.Rows.Add($dr)
					}
				}
				catch
				{
					$errorMessage = "{0}: {1}" -f $Error[0], $Error[0].InvocationInfo.PositionMessage
					Write-Error $errorMessage -ErrorAction Continue
				}
				$pageNumber++
			}
			while ($r -eq $pageSize)
		}#end else no date filter
		
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
	if ($dt.Rows.Count -ge 1)
	{
		return $dt
	}
	else
	{
		Write-Output ("There were no certificate revokcation requests on {0} from {1} until {2}" -f $Arguments.CertificateAuthority, $StartDate, $EndDate)
	}
}