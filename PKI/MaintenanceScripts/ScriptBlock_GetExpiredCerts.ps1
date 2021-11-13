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
	$certProps = @("RequestID", "Request.RequesterName", "CommonName", "ConfigString", "NotBefore", "NotAfter", "SerialNumber", "CertificateTemplate", "CertificateTemplateOID")
	[int32]$pageSize = 50000
	[int32]$pageNumber = 1
	[int32]$LastID = 0
	[String]$sMIME = 'S/MIME'
	[String]$efs = 'EFS'
	[String]$Recovery = 'Recovery'
	$StartDate = [DateTime]::ParseExact($Arguments.StartDate, "MM/dd/yyyy HH:mm:ss", $null)
	$EndDate = [DateTime]::ParseExact($Arguments.EndDate, "MM/dd/yyyy HH:mm:ss", $null)
	
	$dtHeaders = ConvertFrom-Csv @"
		ColumnName,DataType
          CertificateAuthority,string
		RequestID,string
		RequestorName,string
		CommonName,string
		ConfigString,string
		NotBefore,datetime
		NotAfter,datetime
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
	Write-Verbose -Message ("Working on search of {0}" -f $Arguments.CertificateAuthority)
	
	try
	{
		do
		{
			$r = 0
			$CA = Connect-CertificationAuthority -ComputerName $Arguments.CertificationAuthority -ErrorAction Stop | Out-Null
			if ($PSBoundParameters.ContainsKey('$Arguments.Filter'))
			{
				Get-IssuedRequest -CertificationAuthority $Arguments.CertificationAuthority -Filter "RequestID -gt $LastID", "NotAfter -ge $StartDate", "NotAfter -le $EndDate" -Page $pageNumber -PageSize $pageSize -ErrorAction Continue | `
				Where-Object { ((($_.CertificateTemplateOid.FriendlyName) -notlike "*$Recovery*") -and (($_.CertificateTemplateOid.FriendlyName) -notlike "*$efs*") -and (($_.CertificateTemplateOid.FriendlyName) -notlike "*$sMime*")) } | `
				Select-Object -Property $certProps | ForEach-Object {
					$r++
					$LastID = $_.RequestID
					
					[string]$CertificateAuthority = $Arguments.CertificationAuthority
					[string]$RequestID = $_.RequestID
					[string]$RequestorName = $_."Request.RequesterName"
					[string]$CommonName = $_.CommonName
					[string]$ConfigString = $_.ConfigString
					$NotBefore = $_.NotBefore
					$NotAfter = $_.NotAfter
					[string]$serialNumber = $_.SerialNumber
					#[string]$certificateTemplate = $_ | Select-Object @{ Name = "CertificateTemplate"; Expression = { $_.CertificateTemplateOID.FriendlyName } }
					[string]$certificateTemplate = $_.CertificateTemplateOID.FriendlyName
					
					$dr = $dt.NewRow()
					
					$dr["CertificateAuthority"] = $CertificateAuthority
					$dr["RequestID"] = $RequestID
					$dr["RequestorName"] = $RequestorName
					$dr["CommonName"] = $CommonName
					$dr["ConfigString"] = $ConfigString
					$dr["NotBefore"] = $NotBefore
					$dr["NotAfter"] = $NotAfter
					$dr["SerialNumber"] = $SerialNumber
					$dr["CertificateTemplate"] = $certificateTemplate
					
					$dt.Rows.Add($dr)
				}
			}
			else
			{
				Get-IssuedRequest -CertificationAuthority $Arguments.CertificationAuthority -Filter "RequestID -gt $LastID", "NotAfter -ge $StartDate", "NotAfter -le $EndDate" -Page $pageNumber -PageSize $pageSize -ErrorAction Continue | `
				Select-Object $certProps | ForEach-Object {
					$r++
					$LastID = $_.RequestID
					
					[string]$CertificateAuthority = $Arguments.CertificationAuthority
					[string]$RequestID = $_.RequestID
					[string]$RequestorName = $_."Request.RequesterName"
					[string]$CommonName = $_.CommonName
					[string]$ConfigString = $_.ConfigString
					$NotBefore = $_.NotBefore
					$NotAfter = $_.NotAfter
					[string]$serialNumber = $_.SerialNumber
					#[string]$certificateTemplate = $_ | Select-Object @{ Name = "CertificateTemplate"; Expression = { $_.CertificateTemplateOID.FriendlyName } }
					[string]$certificateTemplate = $_.CertificateTemplateOID.FriendlyName
					
					$dr = $dt.NewRow()
					
					$dr["CertificateAuthority"] = $CertificateAuthority
					$dr["RequestID"] = $RequestID
					$dr["RequestorName"] = $RequestorName
					$dr["CommonName"] = $CommonName
					$dr["ConfigString"] = $ConfigString
					$dr["NotBefore"] = $NotBefore
					$dr["NotAfter"] = $NotAfter
					$dr["SerialNumber"] = $SerialNumber
					$dr["CertificateTemplate"] = $certificateTemplate
					
					$dt.Rows.Add($dr)
				}
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
		Write-Output ("There were no expired certificates on {0} from {1} until {2}" -f $CertificateAuthority, $StartDate, $EndDate)
	}
}