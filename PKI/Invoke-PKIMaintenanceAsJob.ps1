#Requires -Module  HelperFunctions

<#
.SYNOPSIS
This script is essentially a wrapper script that will call Invoke-Command.

.DESCRIPTION
This script is essentially a wrapper script that will call Invoke-Command.  The target computer could be the localhost or any computer.  The script is designed to enumerate all domain controllers
if a domain name or forest name is specified.

.PARAMETER ComputerName
Specifies the name of the target computer(s).  If not specified, the local machine is used to execute the scriptblock.

.PARAMETER DomainName
Specifies the fully qualified domain name of the target Active Directory domain. If specified, all domain controllers in the target domain will be used to execute the scriptblock.

.PARAMETER ForestName
Specifies the fully qualified domain name of the target Active Directory forest. If specified, all domain controllers in the target forest will be used to execute the scriptblock.

.PARAMETER ScriptBlock
Specifies the commands to run. Enclose the commands in braces ( { } ) to create a script block.

.PARAMETER ArgumentList
Specifies a hashtable of parameters to pass to the scriptblock when this wrapper script calls Invoke-Command.  The Invoke-Command cmdlet requires an object array of arguments. Therefore, all
arguments are passed by Invoke-Command to the scriptblock as a single script level parameter.  Make sure standalone script blocks are written with this in mind.

.PARAMETER Credential
Specifies the username and password of an account with access to all target computer(s), domain(s), or forest.

.EXAMPLE
PS> Get-Help .\Invoke-PKIMaintenanceAsJob.ps1 -Full

.EXAMPLE
PS> .\.\Invoke-PKIMaintenanceAsJob.ps1 -CertificateAuthority "server1.example.com", "server2.example.com", "server3.example.com" -Credential (Get-Credential) -ScriptBlock ([scriptblock]::Create((Get-Content .\ScriptBlock.ps1 -Raw))) -ArgumentList @{StartDate = (Get-Date).AddMonths(-2); EndDate = (Get-Date)}

.INPUTS
[scriptblock], [hashtable]

.OUTPUTS
None

.NOTES
THIS CODE IS MADE AVAILABLE AS IS, WITHOUT WARRANTY OF ANY KIND. THE
ENTIRE RISK OF THE USE OR THE RESULTS FROM THE USE OF THIS CODE REMAINS
WITH THE USER.
#>

[CmdletBinding()]
param
(
	[Parameter(Mandatory = $false, ParameterSetName = "ParamSetComputerName")]
	[ValidateNotNullOrEmpty()]
	[string[]]$CertificateAuthority = [System.Net.Dns]::GetHostByName("LocalHost").HostName,
	[Parameter(Mandatory = $true, ParameterSetName = "ParamSetComputerName")]
	[ValidateNotNullOrEmpty()]
	[scriptblock]$ScriptBlock,
	[Parameter(Mandatory = $false)]
	[ValidateNotNullOrEmpty()]
	[hashtable]$ArgumentList,
	[Parameter(Mandatory = $false)]
	[ValidateNotNullOrEmpty()]
	[System.Management.Automation.PsCredential]$Credential
)


#Modules
try
{
	Import-Module -Name HelperFunctions -Force  -ErrorAction Stop
}
catch
{
	try
	{
		$moduleName = 'HelperFunctions'
		$ErrorActionPreference = 'Stop';
		$module = Get-Module -ListAvailable -Name $moduleName;
		$ErrorActionPreference = 'Continue';
		$modulePath = Split-Path $module.Path;
		$psdPath = "{0}\{1}" -f $modulePath, "HelperFunctions.psd1"
		Import-Module $psdPath -ErrorAction Stop
	}
	catch
	{
		Write-Error "HelperFunctions PS module could not be loaded. $($_.Exception.Message)" -ErrorAction Stop
	}
}  

#Script
$Error.Clear()
$dtmFormatString = "yyyy-MM-dd HH:mm:ss"
$dtmScriptStartTimeUTC = Get-UTCTime
$myInv = Get-MyInvocation
[int32]$sleepDurationSeconds = 5


Write-Verbose ("[{0} UTC] [SCRIPT] Beginning execution of script." -f $dtmScriptStartTimeUTC.ToString($dtmFormatString))
Write-Verbose ("[{0} UTC] [SCRIPT] Script Name:  {1}" -f (Get-UTCTime).ToString($dtmFormatString), $myInv.ScriptName)

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

try
{
	$Results = @()
	$Jobs = @()
	$JobErrors = @()
	$ConnectivityErrors = @()
	
	switch ($PSCmdlet.ParameterSetName)
	{
		"ParamSetComputerName"
		{
			$caCount = 1
			foreach ($CA in $CertificateAuthority)
			{
				$ActivityMessage = "Querying CA database, please wait..."
				$StatusMessage = ("Processing {0} of {1}: {2}" -f $caCount, $CertificateAuthority.Count, $CA)
				$PercentComplete = ($caCount / $CertificateAuthority.Count * 100)
				Write-Progress -Activity $ActivityMessage -Status $StatusMessage -PercentComplete $PercentComplete -id 1
				
				if ((Test-Connection -ComputerName $CA -Count 1 -Quiet) -eq $true)
				{
					$params = @{
						ComputerName = $CA
						ScriptBlock  = $PSBoundParameters["ScriptBlock"]
						AsJob	   = $true
						ErrorAction  = "Stop"
						JobName	   = $CA
					}
					
					if (($PSBoundParameters.ContainsKey("ArgumentList") -eq $true) -and (($PSBoundParameters["ArgumentList"] -ne $null)))
					{
						$params.ArgumentList = $PSBoundParameters["ArgumentList"]
					}
					
					try
					{
						if (($PSBoundParameters.ContainsKey("Credential") -eq $true) -and ($PSBoundParameters["Credential"] -ne $null))
						{
							$Jobs += Invoke-Command @params -Credential $PSBoundParameters["Credential"]
						}
						else
						{
							$Jobs += Invoke-Command @params
						}
					}
					catch
					{
						if ($Error[0].Exception -eq $null)
						{
							$objProperties = [PSCustomObject]@{
								ComputerName = $CA
								ErrorMessage = "Unknown error occured"
							}
						}
						else
						{
							$objProperties = [PSCustomObject]@{
								ComputerName = $CA
								ErrorMessage = $Error[0].Exception.ToString().Trim()
							}
						}
						$JobErrors += $objProperties
					}
				}
				else
				{
					$ConnectivityErrors += [PSCustomObject]@{
						ComputerName = $CA
					}
				}
				$caCount++
			}
			$OutputFile = "{0}_{1}_{2}.csv" -f $dtmScriptStartTimeUTC.ToString("yyyy-MM-dd_HH.mm.ss"), (([System.IO.Path]::GetFileNameWithoutExtension($myinv.ScriptName.ToString())).ToString().Replace(" ", "").Replace("-", "")), "for_$($CA)"
		}
		
		
	}
	
	$runningJobs = (Get-Job -State Running).Count
	
	while ($runningJobs -ne 0)
	{
		$CurrentJobs = Get-Job
		try
		{
			Write-Verbose ("[{0} UTC] Processing PowerShell RSjobs, please wait..." -f (Get-UTCTime).ToString($dtmFormatString))
			$runningJobs = $CurrentJobs.Where({ $PSItem.State -eq 'Running' }).Count
			$waitingJobs = $CurrentJobs.Where({ $PSItem.State -eq 'NotStarted' }).Count
			$completedJobs = $CurrentJobs.Where({ $PSItem.State -eq 'Completed' }).Count
			
			Write-Output "$runningJobs jobs are running"
			Write-Output "$waitingJobs jobs are waiting"
			Write-Output "$completedJobs jobs are complete"
			
			Start-Sleep -Seconds $sleepDurationSeconds
		}
		catch
		{
			$jobError = [PSCustomObject]@{
				ComputerName = $CurrentJobs.Name
				ErrorMessage = "{0}: {1}" -f $Error[0], $Error[0].InvocationInfo.PositionMessage
			}
			$JobErrors += $jobError
		}
	}
	
	Write-Verbose ("[{0} UTC] Receiving PowerShell jobs, please wait..." -f (Get-UTCTime).ToString($dtmFormatString))
	foreach ($job in $jobs)
	{
		if ($Job.HasErrors)
		{
			$job | Receive-Job
			$Properties = [ordered]@{
				CertificateAuthority = Get-Job $job | Select-Object -ExpandProperty Name
				RequesterName	      = $null
				RequestID		      = $null
				CommonName		 = $null
				ConfigString	      = $null
				NotBefore		      = $null
				NotAfter		      = $null
				SerialNumber	      = $null
				CertificateTemplate  = $null
				JobStatus		      = Get-Job $job | Receive-Job | Select-Object -ExpandProperty Error
			}
			$Results += New-Object PSObject -Property $Properties
		}
		else
		{
			$Results += $Job | Receive-Job
		}
	}
	
	# if results count is only 1, just display to console
	if ($Results.Count -gt 1)
	{
		Write-Verbose ("[{0} UTC] Exporting results data to CSV, please wait..." -f [datetime]::UtcNow.ToString($dtmFormatString))
		$OutputFile = "{0}_{1}.csv" -f $dtmScriptStartTimeUTC.ToString("yyyy-MM-dd_HH_mm_ss"), "CA-Maintenance"
		$Results | Select-Object -Exclude "PSComputerName","RunspaceID","PSShowComputerName" | Export-Csv $OutputFile -NoTypeInformation
	}
	else
	{
		$Results
	}
	
	if ($ConnectivityErrors.Count -ge 1)
	{
		Write-Output ("[{0} UTC] Total Connectivity Errors: {1}" -f [datetime]::UtcNow.ToString($dtmFormatString), $colConnectivityErrors.Count)
		$OutputFile = "{0}_{1}.csv" -f $dtmScriptStartTimeUTC.ToString("yyyy-MM-dd_HH.mm.ss"), ("{0}_ConnectivityErrors" -f "CA-Maintenance")
		$ConnectivityErrors | Export-Csv $OutputFile -NoTypeInformation
	}
	
	if ($colJobErrors.Count -ge 1)
	{
		Write-Output ("[{0} UTC] Total Job Errors: {1}" -f [datetime]::UtcNow.ToString($dtmFormatString), $colJobErrors.Count)
		$OutputFile = "{0}_{1}.csv" -f $dtmScriptStartTimeUTC.ToString("yyyy-MM-dd_HH.mm.ss"), ("{0}_JobErrors" -f "CA-Maintenance")
		$JobErrors | Export-Csv $OutputFile -NoTypeInformation
	}
}
catch
{
	$errorMessage = "{0}: {1}" -f $Error[0], $Error[0].InvocationInfo.PositionMessage
	Write-Error $errorMessage
}
finally
{
	# script is complete
	$dtmScriptStopTimeUTC = Get-UTCTime
	$elapsedTime = New-TimeSpan -Start $dtmScriptStartTimeUTC -End $dtmScriptStopTimeUTC
	Write-Verbose ("[{0} UTC] [SCRIPT] Script Complete" -f $dtmScriptStopTimeUTC.ToString($dtmFormatString))
	Write-Verbose ("[{0} UTC] [SCRIPT] Script Start Time :  {1}" -f (Get-UTCTime).ToString($dtmFormatString), $dtmScriptStartTimeUTC.ToString($dtmFormatString))
	Write-Verbose ("[{0} UTC] [SCRIPT] Script Stop Time  :  {1}" -f (Get-UTCTime).ToString($dtmFormatString), $dtmScriptStopTimeUTC.ToString($dtmFormatString))
	Write-Verbose ("[{0} UTC] [SCRIPT] Elapsed Time      :  {1}  (Days.Hours:Minutes:Seconds)" -f (Get-UTCTime).ToString($dtmFormatString), (New-Object System.TimeSpan($elapsedTime.Days, $elapsedTime.Hours, $elapsedTime.Minutes, $elapsedTime.Seconds)))
}