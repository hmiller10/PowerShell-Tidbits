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
	$certProps = @("RequestID", "Request.StatusCode", "Request.CommonName", "Request.SubmittedWhen", "Request.DispositionMessage", "ConfigString", "CertificateTemplate")
	[int32]$pageSize = 300000
	[int32]$pageNumber = 1
	[int32]$LastID = 0
	[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
	$r = 0

	$dtHeaders = ConvertFrom-Csv @"
		ColumnName,DataType
          CertificateAuthority,string
		RequestID,string
		RequestCommonName,string
		RequestStatusCode,string
		RequestDispositionError,string
		ConfigString,string
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
	Write-Verbose -Message ("Working on search of {0}" -f $Arguments.CertificationAuthority)
	
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
		
		if (($Arguments.Keys -contains ('StartDate')) -and ($Arguments.Keys -contains ('EndDate')))
		{
			$StartDate = $Arguments.StartDate
			$EndDate = $Arguments.EndDate
			
			do
			{
				try
				{
					$CertificateAuthority | Get-FailedRequest -Filter "RequestID -gt $LastID", "Request.SubmittedWhen -ge $StartDate", "Request.SubmittedWhen -le $EndDate" -Properties $certProps -Page $pageNumber -PageSize $pageSize -ErrorAction Stop | `
					Select-Object -Property $certProps | ForEach-Object {
						$LastID = $_.RequestID;$r++
						
						[string]$CertificateAuthority = $Arguments.CertificationAuthority
						[string]$CertificateAuthority = $Arguments.CertificationAuthority
						[string]$RequestID = $_.RequestID
						[string]$RequestCommonName = $_.CommonName
						[string]$RequestStatusCode = $_."Request.StatusCode"
						[string]$RequestDispositionError = $_."Request.DispositionMessage"
						[string]$ConfigString = $_.ConfigString
						[string]$certificateTemplate = $_.CertificateTemplateOID.FriendlyName
						
						$dr = $dt.NewRow()
						
						$dr["CertificateAuthority"] = $CertificateAuthority
						$dr["RequestID"] = $RequestID
						$dr["RequestCommonName"] = $RequestCommonName
						$dr["RequestStatusCode"] = $RequestStatusCode
						$dr["RequestDispositionError"] = $RequestDispositionError
						$dr["ConfigString"] = $ConfigString
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
		}
		else
		{
			do
			{
				try
				{
					$CertificateAuthority | Get-FailedRequest -Filter "RequestID -gt $LastID"  -Properties $certProps -Page $pageNumber -PageSize $pageSize -ErrorAction Stop | `
					Select-Object -Property $certProps | ForEach-Object {
						$LastID = $_.RequestID;$r++
						
						[string]$CertificateAuthority = $Arguments.CertificationAuthority
						[string]$CertificateAuthority = $Arguments.CertificationAuthority
						[string]$RequestID = $_.RequestID
						[string]$RequestCommonName = $_.CommonName
						[string]$RequestStatusCode = $_."Request.StatusCode"
						[string]$RequestDispositionError = $_."Request.DispositionMessage"
						[string]$ConfigString = $_.ConfigString
						[string]$certificateTemplate = $_.CertificateTemplateOID.FriendlyName
						
						$dr = $dt.NewRow()
						
						$dr["CertificateAuthority"] = $CertificateAuthority
						$dr["RequestID"] = $RequestID
						$dr["RequestCommonName"] = $RequestCommonName
						$dr["RequestStatusCode"] = $RequestStatusCode
						$dr["RequestDispositionError"] = $RequestDispositionError
						$dr["ConfigString"] = $ConfigString
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
	if ($dt.Rows.Count -ge 1)
	{
		return $dt
	}
	else
	{
		Write-Output ("There were no failed certificate requests on {0} from {1} until {2}" -f $Arguments.CertificationAuthority, $StartDate, $EndDate)
	}
}