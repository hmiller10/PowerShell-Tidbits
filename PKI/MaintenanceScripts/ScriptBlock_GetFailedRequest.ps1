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
	$certProps = @()
	$certProps = @("RequestID", "Request.StatusCode", "Request.CommonName", "Request.SubmittedWhen", "Request.DispositionMessage", "ConfigString", "CertificateTemplate")
	[int32]$pageSize = 50000
	[int32]$pageNumber = 1
	[int32]$LastID = 0
	$StartDate = [DateTime]::ParseExact($Arguments.StartDate, "MM/dd/yyyy HH:mm:ss", $null)
	$EndDate = [DateTime]::ParseExact($Arguments.EndDate, "MM/dd/yyyy HH:mm:ss", $null)
	
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
		do
		{
			$r = 0
			Connect-CertificationAuthority -ComputerName $Arguments.CertificationAuthority -ErrorAction Stop | Out-Null
			Get-FailedRequest -CertificationAuthority $Arguments.CertificationAuthority -Filter "RequestID -gt $LastID", "Request.SubmittedWhen -ge $StartDate", "Request.SubmittedWhen -le $EndDate" -Page $pageNumber -PageSize $pageSize -ErrorAction Continue | `
			Select-Object -Property $certProps | ForEach-Object {
				$r++
				$LastID = $_.RequestID
				
				[string]$CertificateAuthority = $Arguments.CertificationAuthority
				[string]$RequestID = $_.RequestID
				[string]$RequestorName = $_."Request.RequesterName"
				[string]$RequestCommonName = $_.CommonName
				[string]$RequestStatusCode = $_."Request.StatusCode"
				[string]$RequestDispositionError = $_."Request.DispositionMessage"
				[string]$ConfigString = $_.ConfigString
				#[string]$certificateTemplate = $_ | Select-Object @{ Name = "CertificateTemplate"; Expression = { $_.CertificateTemplateOID.FriendlyName } }
				[string]$certificateTemplate = $_.CertificateTemplateOID.FriendlyName
				
				$dr = $dt.NewRow()
				
				$dr["CertificateAuthority"] = $CertificateAuthority
				$dr["RequestID"] = $RequestID
				$dr["RequestorName"] = $RequestorName
				$dr["RequestCommonName"] = $RequestCommonName
				$dr["RequestStatusCode"] = $RequestStatusCode
				$dr["RequestDispositionError"] = $RequestDispositionError
				$dr["ConfigString"] = $ConfigString
				$dr["CertificateTemplate"] = $certificateTemplate
				
				$dt.Rows.Add($dr)
			}
			
		}
		while ($r -eq $pageSize)
		
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
		Write-Output ("There were no failed certificate requests on {0} from {1} until {2}" -f $CertificateAuthority, $StartDate, $EndDate)
	}
}