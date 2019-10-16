#Requires -Modules ADCSAdministration, ADCSTemplate, Microsoft.PowerShell.Security, PKI -Version 3.0

#Region Help
<#

.NOTES 
Needs to be run from Certificate Authority server,
Requires third party module, PSPKI, from http://www.sysadmins.lv

.SYNOPSIS 
Backs up Certificate Authority Issuing CA database and other
critical files

.DESCRIPTION 
This script utilizes native PowerShell cmdlets in Windows
Server 2012 R2 and later to perform daily backups for the CA database locally. After
those backups are completed, the static files are moved to tape. Only 4
backups are kept at a time.

.PARAMETER 
BackupFolder

.PARAMETER 
IncludeKey

.PARAMETER 
IncludeIISFiles

.PARAMETER 
IncludeThalesHSM

.OUTPUTS 
Daily e-mail report of backup status from previous day.

.EXAMPLE 
.\Backup-CertAuthority.ps1 -BackupFolder

.EXAMPLE 
.\Backup-CertAuthority.ps1 -BackupFolder -IncludeKey

.EXAMPLE
.\Backup-CertAuthority.ps1 -BackupFolder -IncludeIISFiles

.EXAMPLE 
.\Backup-CertAuthority.ps1 -BackupFolder -IncludeKey -IncludeIISFiles -IncludeThalesHSM

#>
###########################################################################
#
# AUTHOR:  Heather Miller
#
# VERSION HISTORY:
# October 8, 2019 - Version 2.0
# 
###########################################################################
#EndRegion

[CmdletBinding()]
param
(
	[Parameter(Mandatory = $true, HelpMessage = "Specify full path where CA backup should should reside. EG: E:\CABackups")]
	[alias("Folder")]
	[String]$BackupFolder,
	[Parameter(Mandatory = $false)]
	[Switch]$IncludeKey,
	[Parameter(Mandatory = $false)]
	[Switch]$IncludeIISFiles,
	[Parameter(Mandatory = $false)]
	[Switch]$IncludeThalesHSM
	
	
)

#Region ExecutionPolicy
#Set Execution Policy for Powershell
Set-ExecutionPolicy RemoteSigned
#EndRegion

#Region Modules
#Check IF required module is loaded, IF not load import it
IF (-not (Get-Module ADCSAdministration))
{
	Import-Module ADCSAdministration
}
If (-not (Get-Module PKI))
{
	Import-Module -Name PKI
}
If (-not (Get-Module Microsoft.PowerShell.Security))
{
	Import-Module Microsoft.PowerShell.Security
}
If (-not (Get-Module ADCSTemplate))
{
	Import-Module ADCSTemplate
}
#EndRegion

#Region Variables
#Dim variables
$Limit = (Get-Date).AddDays(-1)
$myScriptName = $MyInvocation.MyCommand.Name
$evtProps = @("Index", "TimeWritten", "EntryType", "Source", "InstanceID", "Message")
$LogonServer = $env:LOGONSERVER
$RetentionLimit = (Get-Date).AddDays(-3)
$ServerName = [System.Net.Dns]::GetHostByName("LocalHost").HostName
[Void][Reflection.Assembly]::LoadWithPartialName("System.Web")
Add-Type -AssemblyName WindowsBase
Add-Type -AssemblyName "System.IO.Compression.FileSystem"
$compressionLevel = [System.IO.Compression.CompressionLevel]::Optimal
$thisServer = $env:COMPUTERNAME
$VerbosePreference = 'Continue'
#EndRegion

#Region Functions

Function Get-TodaysDate
{
	#Begin function to get short date
	Get-Date -Format "MM-dd-yyyy"
} #End function Get-TodaysDate

Function Check-Path
{
	#Begin function to check path variable and return results
	[CmdletBinding()]
	Param
	(
		[Parameter(Mandatory, Position = 0)]
		[string]$Path,
		[Parameter(Mandatory, Position = 1)]
		$PathType
	)
	
	
	#Define variables
	$VerbosePreference = "Continue"
	
	Switch ($PathType)
	{
		File {
			If ((Test-Path -Path $Path -PathType Leaf) -eq $true)
			{
				Write-Verbose -Message "File: $Path already exists..."
			}
			Else
			{
				New-Item -Path $Path -ItemType File -Force
				Write-Verbose -Message "File: $Path not present, creating new file..."
			}
		}
		Folder {
			If ((Test-Path -Path $Path -PathType Container) -eq $true)
			{
				Write-Verbose -Message "Folder: $Path already exists..."
			}
			Else
			{
				New-Item -Path $Path -ItemType Directory -Force
				Write-Verbose -Message "Folder: $Path not present, creating new folder"
			}
		}
	}
} #end function Check-Path

Function UTC-Now
{
	#Begin function to get UTC date and time
	[System.DateTime]::UtcNow
} #End UTC-Now

Function Get-LongDate
{
	#Begin function to get date and time in long format
	Get-Date -Format G
} #End function Get-LongDate

Function Get-ReportDate
{
	#Begin function set report date format
	Get-Date -Format "yyyy-MM-dd"
} #End function Get-ReportDate

#EndRegion








#Region Script
#Begin Script
$Error.Clear()

#Start Function timer, to display elapsed time for function. Uses System.Diagnostics.Stopwatch class - see here: https://msdn.microsoft.com/en-us/library/system.diagnostics.stopwatch(v=vs.110).aspx 
$stopWatch = [System.Diagnostics.Stopwatch]::StartNew()
$dtmScriptStartTimeUTC = Utc-Now
$dtmFormatString = "yyyy-MM-dd HH:mm:ss"
$transcriptFileName = "{0}-{1}-Transcript.txt" -f $dtmScriptStartTimeUTC.ToString("yyyy-MM-dd_HH.mm.ss"), "$($thisServer)-CARoleBackup"

$scriptDir = Split-Path $MyInvocation.MyCommand.Path
$scriptName = $MyInvocation.ScriptName

#Region Check folder structures
Check-Path -Path $BackupFolder -PathType Folder
$TodaysFldr = "{0}\{1}" -f $BackupFolder, "CABackup_$(Get-ReportDate)"
Check-Path -Path $TodaysFldr -PathType Folder
#EndRegion

try
{
	# start transcript file
	Start-Transcript ("{0}\{1}" -f $TodaysFldr, $transcriptFileName)
	Write-Verbose -Message ("[{0} UTC] [SCRIPT] Beginning execution of script." -f $dtmScriptStartTimeUTC.ToString($dtmFormatString))
	Write-Verbose -Message ("[{0} UTC] [SCRIPT] Script Name             		:  {1}" -f $(UTC-Now).ToString($dtmFormatString), $scriptName)
	Write-Verbose -Message ("[{0} UTC] [SCRIPT] Script Directory path   		:  {1}" -f $(UTC-Now).ToString($dtmFormatString), $scriptDir)
	Write-Verbose -Message ("[{0} UTC] [SCRIPT] Main Backup Folder path  		:  {1}" -f $(UTC-Now).ToString($dtmFormatString), $BackupFolder)
	Write-Verbose -Message ("[{0} UTC] [SCRIPT] Todays Backup Folder path     	:  {1}" -f $(UTC-Now).ToString($dtmFormatString), $TodaysFldr)
	
	$dtHeadersCsv = @"
	ColumnName,DataType
	Date,datetime
	ServerName,string
	TaskName,string
	TaskStatus,string
	Result,string
	ScriptRunTime,string
"@
	
	$dtHeaders = ConvertFrom-Csv -InputObject $dtHeadersCsv
	
	$dt = New-Object System.Data.DataTable
	
	foreach ($header in $dtHeaders)
	{
		[void]$dt.Columns.Add([System.Data.DataColumn]$header.ColumnName.ToString(), $header.DataType)
	}
	
	#Region Backup-CA-DB-Key
	#Backup Certificate Authority Database and Private Key
	$Error.Clear()
	try
	{
		
		$taskName = "CA DB Backup"
		
		$params = @{
			Path = $TodaysFldr
		}
		
		if ($PSBoundParameters.ContainsKey('IncludeKey'))
		{
			#Get password to be used during backup process
			$pw = [System.Web.Security.Membership]::GeneratePassword(20, 5)
			$BkpPassword = ConvertTo-SecureString $pw -AsPlainText -Force
			
			#Run this command if backing database with key included
			$params.Add('Password', $BkpPassword)
			
			#Export password to encrypted XML file in case restore is required.
			$xmlFile = "{0}\{1}" -f $TodaysFldr, "secPassword.xml"
			$BkpPassword | Export-Clixml -Path $xmlFile -Force
		}
		else
		{
			#Run this command if backing up just the database without the key. EG: Key is stored on HSM
			$params.Add('DatabaseOnly', $true)
		}
		
		Backup-CARoleService @params
		
		$CAEvents = Get-EventLog -LogName Application -Source 'ESENT' -Newest 1 | Where-Object { $_.EventID -eq 213 } | `
		Select-Object -Property TimeWritten, EntryType, Source, InstanceID, Message | Sort-Object -Property TimeWritten
		
		If ($CAEvents -ne $null)
		{
			$EventInfo = "Event Time: " + ($CAEvents).TimeWritten + " Event ID: " + ($CAEvents).InstanceID + " Event Details: " + ($CAEvents).Message
			Write-Verbose -Message $EventInfo
			
			$taskStatus = "Success"
			$Result = $EventInfo | Out-String
		}
		Else
		{
			Write-Warning -Message "Certificate Authority Backups failed $(Get-TodaysDate). Please investigate."
			$taskStatus = "Failed"
			$Rssult = "Certificate Authority Backups failed $(Get-TodaysDate). Please investigate."
		}

	}
	catch
	{
		Write-Warning -Message "Certificate Authority Backups failed $(Get-TodaysDate). Please investigate."
		$taskStatus = "Failed"
		$Rssult = "Certificate Authority Backups failed $(Get-TodaysDate). Please investigate."
		
		$errorMessage = "{0}: {1}" -f $Error[0], $Error[0].InvocationInfo.PositionMessage
		Write-Error $errorMessage
		$Error.Clear()
	}
	
	$dataRow = $dt.NewRow()
	$dataRow["Date"] = (Get-LongDate).ToString()
	$dataRow["ServerName"] = $ServerName
	$dataRow["TaskName"] = $taskName
	$dataRow["TaskStatus"] = $taskStatus
	$dataRow["Result"] = $Result
	$dt.Rows.Add($dataRow)
	
	$taskName = $taskStatus = $Result = $null
	#EndRegion
	
	#region Backup CA Certs
	#Backup CA Certificates
	$Error.Clear()
	$taskName = "Backup CA Certificates"
	
	$certSource = "{0}\{1}\{2}\{3}" -f $env:SystemRoot, "System32", "Certsrv", "CertEnroll"
	$certBkpFldr = "{0}\{1}" -f $TodaysFldr, "CACerts"
	Check-Path -Path $certBkpFldr -PathType Folder
	
	[Array]$certs = Get-ChildItem -Path $certSource -Exclude *.crl
	
	foreach ($cert in $certs)
	{
		Copy-Item -Path $cert.FullName -Destination $certBkpFldr
	}
	
	[Array]$copiedCerts = Get-ChildItem -Path $certBkpFldr -Filter *
	
	if ($copiedCerts.Count -ge 1)
	{
		$taskStatus = "Success"
		$Result = "Successfully copied CA certificates to backup folder."
	}
	else
	{
		$taskStatus = "Failure"
		$Result = "Failed to successfully copy CA certificates to backup folder."
	}
	
	$dataRow = $dt.NewRow()
	$dataRow["Date"] = (Get-LongDate).ToString()
	$dataRow["ServerName"] = $ServerName
	$dataRow["TaskName"] = $taskName
	$dataRow["TaskStatus"] = $taskStatus
	$dataRow["Result"] = $Result
	$dt.Rows.Add($dataRow)
	
	$taskName = $taskStatus = $Result = $null
	
	#EndRegion
	
	#Region Backup-CA-Registry
	$Error.Clear()
	#Backup Certificate Authority Registry Hive
	$RegFldr = "{0}\{1}" -f $TodaysFldr, "RegistryKeys"
	Check-Path -Path $RegFldr -PathType Folder
	$rptDate = (Get-ReportDate).ToString()
	$RegFile = "{0}\{1}_{2}{3}" -f $RegFldr, "CARegistry", $rptDate, ".reg"
	
	try
	{
		$taskName = "Backup CA Registry Keys"
		#Run reg.exe from command line to backup CA registry hive
		reg.exe export HKLM\System\CurrentControlSet\Services\CertSvc $RegFile
		
		If (Test-Path -Path $RegFile -IsValid)
		{
			$taskStatus = "Success"
			$Result = "Certificate Services registry key: $RegFile export was successful."
		}
		Else
		{
			$Status = "Failed"
			$Result = "Certificate Services registry key export failed on $(Get-ReportDate)."
		}
		
		$dataRow = $dt.NewRow()
		$dataRow["Date"] = (Get-LongDate).ToString()
		$dataRow["ServerName"] = $ServerName
		$dataRow["TaskName"] = $taskName
		$dataRow["TaskStatus"] = $taskStatus
		$dataRow["Result"] = $Result
		$dt.Rows.Add($dataRow)
		
		$taskName = $taskStatus = $Result = $null
	}
	catch
	{
		$errorMessage = "{0}: {1}" -f $Error[0], $Error[0].InvocationInfo.PositionMessage
		Write-Error $errorMessage
		$Error.Clear()
	}
	#EndRegion
	
	#Region Backup-Policy-File
	$Error.Clear()
	#If not using a Policy Certificate Authority server and policies are implemented using .INF file, backup configuration file.
	#Backup Certificate Policy .Inf file
	$PolicyFldr = "{0}\{1}" -f $TodaysFldr, "PolicyFile"
	Check-Path -Path $PolicyFldr -PathType Folder
	$PolicyFile = "{0}\{1}" -f $env:SystemRoot, "CAPolicy.inf"
	
	$taskName = "Backup CAPolicy.inf"
	
	if (Test-Path -Path $PolicyFile -PathType Leaf -IsValid)
	{
		Copy-Item -Path $PolicyFile -Destination $PolicyFldr
		$bkpPolicyFile = "{0}\{1}" -f $PolicyFile, $PolicyFile
		
		If (Test-Path -Path $bkpPolicyFile -PathType Leaf -IsValid)
		{
			$taskStatus = "Success"
			$Result = "Backup copy of policy file: $PolicyFile was successful."
		}
		Else
		{
			$taskStatus = "Failed"
			$Result = "Backup copy of policy file: $PolicyFile failed. Please investigate."
		}
		
		$dataRow = $dt.NewRow()
		$dataRow["Date"] = (Get-LongDate).ToString()
		$dataRow["ServerName"] = $ServerName
		$dataRow["TaskName"] = $taskName
		$dataRow["TaskStatus"] = $taskStatus
		$dataRow["Result"] = $Result
		$dt.Rows.Add($dataRow)
		
		$taskName = $taskStatus = $Result = $null
	}
	else
	{
		$errorMessage = "{0}: {1}" -f $Error[0], $Error[0].InvocationInfo.PositionMessage
		Write-Error $errorMessage
		$Error.Clear()
	}
	#EndRegion
	
	#Region Backup-IIS-Files
	$Error.Clear()
	if ($PSBoundParameters.ContainsKey('IncludeIISFiles'))
	{
		if (((Get-WindowsFeature -Name ADCS-Web-Enrollment).InstallState -eq "Installed") -and ((Get-WindowsFeature -name Web-Server).InstallState -eq "Installed"))
		{
			try
			{
				#Backup IIS Custom files
				$taskName = "Backup IIS Custom Files"
				
				If (-not (Get-Module WebAdministration))
				{
					Import-Module WebAdministration
				}

				[Array]$colWebSites = Get-WebSite

				
				foreach ($site in $colWebSites)
				{
					$siteFolder = $site.physicalpath
					if ($siteFolder -match "%systemdrive%")
					{
						$siteFolder = [String]$siteFolder -replace "%systemdrive%", $env:SystemDrive
					}
					
					$IISCustomFldr = "{0}\{1}" -f $TodaysFldr, "IISCustomizations"
					Check-Path -Path $IISCustomFldr -PathType Folder
					
					$archiveFile = "{0}\{1}" -f $IISCustomFldr, "Custom_IIS_files_archive_for_$(Get-ReportDate).zip"
					if ((Test-Path -Path $archiveFile -PathType Leaf) -eq $true) { Remove-Item -Path $archiveFile -Confirm:$false }
					#See https://msdn.microsoft.com/en-us/library/hh875104(v=vs.110).aspx?cs-save-lang=1&cs-lang=vb#code-snippet-1
					[IO.Compression.ZipFile]::CreateFromDirectory($siteFolder, $archiveFile, $compressionLevel, $false)
					#[IO.Compression.ZipFile]::Dispose()
					
					if ((Test-Path -Path $archiveFile -PathType Leaf) -eq $true)
					{
						$taskStatus = "Success"
						$Result = "Backup copy of IIS custom files were successful."
					}
					else
					{
						$taskStatus = "Failed"
						$Result = "Backup copy of IIS custom files failed. Please investigate."
					}
					
					$dataRow = $dt.NewRow()
					$dataRow["Date"] = (Get-LongDate).ToString()
					$dataRow["ServerName"] = $ServerName
					$dataRow["TaskName"] = $taskName
					$dataRow["TaskStatus"] = $taskStatus
					$dataRow["Result"] = $Result
					$dt.Rows.Add($dataRow)
				}

				$taskName = $taskStatus = $Result = $null
				
			}
			catch
			{
				$errorMessage = "{0}: {1}" -f $Error[0], $Error[0].InvocationInfo.PositionMessage
				Write-Error $errorMessage
				$Error.Clear()
			}
			
		}
		
	}
	#EndRegion
	
	$Error.Clear()
	$activeConfig = Get-ItemProperty -Path "HKLM:\System\CurrentControlSet\Services\CertSvc\configuration" -Name Active
	$activeConfig = $activeConfig.Active
	
	$CAType = Get-ItemProperty -Path HKLM:\System\CurrentControlSet\Services\CertSvc\configuration\$activeConfig -Name CAType
	if ($CAType.CAType -eq "1")
	{
		#Region Export-CA-Template-List
		$Error.Clear()
		$taskName = "Gather list of published templates"
		
		Write-Verbose -Message "Backing up list of published templates."
		
		certutil.exe -catemplates > $TodaysFldr\CATemplates.txt
		if ($?)
		{
			$taskStatus = "Success"
			$Result = "Successfully compiled a list of published ADCS templates."
		}
		else
		{
			$taskStatus = "Failed"
			$Result = "Unable to extract a list of published templates."
		}
		
		$dataRow = $dt.NewRow()
		$dataRow["Date"] = (Get-LongDate).ToString()
		$dataRow["ServerName"] = $ServerName
		$dataRow["TaskName"] = $taskName
		$dataRow["TaskStatus"] = $taskStatus
		$dataRow["Result"] = $Result
		$dt.Rows.Add($dataRow)
		
		$taskName = $taskStatus = $Result = $null
		#endregion
		
		######
		#region Export-CA-Templates-To-File
		$taskName = "Backup CA Templates"
		
		try
		{
			##Export (Backup) CA Templates
			
			$templateFldr = "{0}\{1}" -f $TodaysFldr, "Templates"
			Check-Path -Path $templateFldr -PathType Folder
			
			#Get list of templates to process
			$Templates = Get-ADCSTemplate * | Select-Object -Property DisplayName, Name | Sort-Object -Property Name
			
			#export templates to files
			foreach ($Template in $Templates)
			{
				$templateName = ($Template).Name
				$tempDisplayName = ($Template).DisplayName
				
				If ($templateName -match '\/')
				{
					$templateName = $templateName -replace "\/", ""
					Write-Verbose -Message "The template $($Template.Name) had a ""/"" in the name. It has been removed in the output file due to limitations of the filesystem."
					Write-Verbose -Message "If using the templates exported from this script at a later date for import, the ""/"" will need to be added back to the file name."
				}
				ElseIf ($templateName -match '\\')
				{
					$templateName = $templateName -replace "\\", ""
					Write-Verbose -Message "The template $($Template.Name) had a ""\"" in the name. It has been removed in the output file due to limitations of the filesystem."
					Write-Verbose -Message "If using the templates exported from this script at a later date for import, the ""\"" will need to be added back to the file name."
				}
				
				Write-Verbose -Message "Exporting $tempDisplayName..."
				$templateFileName = "{0}.{1}" -f $templateName, "json"
				$templateFile = "{0}\{1}" -f $templateFldr, $templateFileName
				
				Export-ADCSTemplate -DisplayName $tempDisplayName | Out-File -FilePath $templateFile -Force
			}
			
			#Create .Zip backup to save space
			$templatesArchiveFile = "{0}\{1}" -f $TodaysFldr, "CertTemplates.zip"
			if ((Test-Path -Path $templatesArchiveFile -PathType Leaf) -eq $true) { Remove-Item -Path $templatesArchiveFile -Confirm:$false }
			#See https://msdn.microsoft.com/en-us/library/hh875104(v=vs.110).aspx?cs-save-lang=1&cs-lang=vb#code-snippet-1
			[IO.Compression.ZipFile]::CreateFromDirectory($templateFldr, $templatesArchiveFile, $compressionLevel, $false)
				
			if ((Test-Path -Path $templatesArchiveFile -PathType Leaf) -eq $true)
			{
				Get-ChildItem -Path $TodaysFldr | Where-Object { ($_.PsIsContainer) -and ($_.Name -like "Templates*") } | Remove-Item -Recurse -Force -Confirm:$false
				$taskStatus = "Success"
				$Result = "Certficate Template definitions have been exported to {0}\{1}\{2}" -f $TodaysFldr, "Templates", "CertTemplates.zip."
			}
		}
		catch
		{
			$taskStatus = "Failed"
			$Result = "Template backups failed. Please contact Administrator."
			$errorMessage = "{0}: {1}" -f $Error[0], $Error[0].InvocationInfo.PositionMessage
			Write-Error $errorMessage
			$Error.Clear()
		}
		
		$dataRow = $dt.NewRow()
		$dataRow["Date"] = (Get-LongDate).ToString()
		$dataRow["ServerName"] = $ServerName
		$dataRow["TaskName"] = $taskName
		$dataRow["TaskStatus"] = $taskStatus
		$dataRow["Result"] = $Result
		$dt.Rows.Add($dataRow)
		
		$taskName = $taskStatus = $Result = $null
		
		#EndRegion
	}
	
	#Region Backup Thales Data Files
	$Error.Clear()
	if ($PSBoundParameters.ContainsKey('IncludeThalesHSM'))
	{
		try
		{
			#Region Backup-nCipher Directory
			
			$taskName = "Backup Thales HSM Files."
			
			$nCipherFldr1 = "C:\ProgramData\nCipher"
			
			$hsmBkpRoot = "{0}\{1}" -f $TodaysFldr, "HSM"
			Check-Path -Path $hsmBkpRoot -PathType Folder
			
			Write-Verbose -Message "Backing up Thales HSM configuration information..."
			
			#Create .Zip backup to save space
			$nCipherZipFile1 = "{0}\{1}" -f $hsmBkpRoot, "nCipher1.zip"
			if ((Test-Path -Path $nCipherZipFile1 -PathType Leaf) -eq $true) { Remove-Item -Path $nCipherZipFile1 -Confirm:$false }
			#See https://msdn.microsoft.com/en-us/library/hh875104(v=vs.110).aspx?cs-save-lang=1&cs-lang=vb#code-snippet-1
			[IO.Compression.ZipFile]::CreateFromDirectory($nCipherFldr1, $nCipherZipFile1, $compressionLevel, $false)
			
			if (Test-Path -Path $nCipherZipFile1 -IsValid)
			{
				$taskStatus = "Success"
				$Result = "Successfully collected relevant Thales data files."
			}
			else
			{
				$taskStatus = "Failed"
				$Result = "Failed to collect needed Thales data files."
			}
			
			$dataRow = $dt.NewRow()
			$dataRow["Date"] = (Get-LongDate).ToString()
			$dataRow["ServerName"] = $ServerName
			$dataRow["TaskName"] = $taskName
			$dataRow["TaskStatus"] = $taskStatus
			$dataRow["Result"] = $Result
			$dt.Rows.Add($dataRow)
			
			$taskName = $taskStatus = $Result = $null
			
			#EndRegion
			
		}
		catch
		{
			$errorMessage = "{0}: {1}" -f $Error[0], $Error[0].InvocationInfo.PositionMessage
			Write-Error $errorMessage
			$Error.Clear()
		}
		
	}
	
	#EndRegion
	
	#Stop the stopwatch	
	$stopWatch.Stop()
	
	$dtmScriptStopTimeUTC = UTC-Now
	$elapsedTime = New-TimeSpan -Start $dtmScriptStartTimeUTC -End $dtmScriptStopTimeUTC
	$runtime = $stopWatch.Elapsed.ToString('dd\.hh\:mm\:ss')
	
	$reportBody | ConvertTo-Html -Title "$thisServer CA backup report as of $(Get-TodaysDate)" -Body $reportBody -PostContent "Script took: $($runTime) to run."
	
	Write-Verbose -Message ("[{0} UTC] [SCRIPT] Script Complete" -f $(UTC-Now).ToString($dtmFormatString))
	Write-Verbose -Message ("[{0} UTC] [SCRIPT] Script Start Time :  {1}" -f $(UTC-Now).ToString($dtmFormatString), $dtmScriptStartTimeUTC.ToString($dtmFormatString))
	Write-Verbose -Message ("[{0} UTC] [SCRIPT] Script Stop Time  :  {1}" -f $(UTC-Now).ToString($dtmFormatString), $dtmScriptStopTimeUTC.ToString($dtmFormatString))
	Write-Verbose -Message ("[{0} UTC] [SCRIPT] Elapsed Time: {1:N0}.{2:N0}:{3:N0}:{4:N1}  (Days.Hours:Minutes:Seconds)" -f $(UTC-Now).ToString($dtmFormatString), $elapsedTime.Days, $elapsedTime.Hours, $elapsedTime.Minutes, $elapsedTime.Seconds)
	
}
catch
{
	$errorMessage = "{0}: {1}" -f $Error[0], $Error[0].InvocationInfo.PositionMessage
	Write-Error $errorMessage
}
finally
{
	#Region Remove-Old-Backups
	#Cleanup old CA Backups
	$taskName = "Backup and Archive cleanup"
	[Array]$Folders = Get-ChildItem -Path $BackupFolder -Force | Where-Object { $_.LastWriteTime -le $RetentionLimit -and $_.PSisContainer } | Sort-Object -Property LastWriteTime -Descending
	
	if ($Results.Count -ge 1)
	{
		foreach ($folder in $Folders)
		{
			$FolderName = ($Folder).FullName
			$FolderName | Remove-Item -Force -Recurse
			If ($?)
			{
				$taskStatus = "Success"
				$Result = $FolderName + " was deleted as of $(Get-LongDate)"
				
				$dataRow = $dt.NewRow()
				$dataRow["Date"] = (Get-LongDate).ToString()
				$dataRow["ServerName"] = $ServerName
				$dataRow["TaskName"] = $taskName
				$dataRow["TaskStatus"] = $taskStatus
				$dataRow["Result"] = $Result
				$dt.Rows.Add($dataRow)
			}
			else
			{
				$taskStatus = "Failed"
				$Result = "Failed to remove $FolderName. Contact Administrator."
				
				$dataRow = $dt.NewRow()
				$dataRow["Date"] = (Get-LongDate).ToString()
				$dataRow["ServerName"] = $ServerName
				$dataRow["TaskName"] = $taskName
				$dataRow["TaskStatus"] = $taskStatus
				$dataRow["Result"] = $Result
				$dt.Rows.Add($dataRow)
			}
		}
	}
	else
	{
		$taskStatus = "No action required."
		$Result = "There are no backup folders ready for deletion as of $(Get-TodaysDate)."
		
		$dataRow = $dt.NewRow()
		$dataRow["Date"] = (Get-LongDate).ToString()
		$dataRow["ServerName"] = $ServerName
		$dataRow["TaskName"] = $taskName
		$dataRow["TaskStatus"] = $taskStatus
		$dataRow["Result"] = $Result
		$dt.Rows.Add($dataRow)
	}
	
	#EndRegion
	
	$dataRow = $dt.NewRow()
	$dataRow["Date"] = (Get-LongDate).ToString()
	$dataRow["ServerName"] = $ServerName
	$dataRow["ScriptRunTime"] = $runtime
	$dt.Rows.Add($dataRow)
	
	Stop-Transcript
	
	#Send E-mail Report
	
	$Header = @"
	<style>BODY{Background-Color:white;Font-Family: Arial;Font-Size: 12pt;}
	TABLE {border-width: 1px;border-style: solid;border-color: black;border-collapse: collapse;}
	TH {border-width: 1px;padding: 3px;border-style: solid;border-color: black;color:white;background-color: #003366;Font-Family: Arial;Font-Size: 14pt;Text-Align: Center;}
	TD {border-width: 1px;padding: 3px;border-style: solid;border-color: black;Font-Family: Arial;Font-Size: 12pt;Text-Align: Left;}
	.odd  { background-color:#ffffff; }
	.even { background-color:#dddddd; }
	</style>
	<title>Microsoft ADCS Database Backup Report</title>
"@
	
	$Body = $dt | Select Date, ServerName, TaskName, TaskStatus, Result, ScriptRunTime | ConvertTo-Html -Head $Header -PreContent "<h2>Microsoft ADCS Database Backup Report for $($activeConfig)</h2>" -PostContent "<p>For further details, contact the PKI Infrastructure Team.</p>"
	$Body = $Body | Out-String
	
	$From = 'no-reply@yourdomain.com'
	$To = "you@yourdomain.com"
	$ReportSubject = "Active Directory Certificate Authority Backup status for $(Get-TodaysDate)"
	$smtpServer = 'smtprelay.yourdomain.com'
	
	$transcript = "{0}\{1}" -f $TodaysFldr, $transcriptFileName
	$colAttachments = @(Get-ChildItem -Path $transcript -File)
	if ($PSBoundParameters.ContainsKey('IncludeKey'))
	{
		$colAttachments += Get-ChildItem -Path $xmlFile -File
	}
	
	Send-MailMessage -From $From -To $To -Subject $ReportSubject -Body $Body -Attachments $colAttachments -BodyAsHTML -SmtpServer $smtpServer
}



#EndRegion
# SIG # Begin signature block
# MIInCQYJKoZIhvcNAQcCoIIm+jCCJvYCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCBmtwUrHEHAM6cL
# 6//KspmNgtDDWlPgPhsuu0XL1FJ8IqCCIaUwggQVMIIC/aADAgECAgsEAAAAAAEx
# icZQBDANBgkqhkiG9w0BAQsFADBMMSAwHgYDVQQLExdHbG9iYWxTaWduIFJvb3Qg
# Q0EgLSBSMzETMBEGA1UEChMKR2xvYmFsU2lnbjETMBEGA1UEAxMKR2xvYmFsU2ln
# bjAeFw0xMTA4MDIxMDAwMDBaFw0yOTAzMjkxMDAwMDBaMFsxCzAJBgNVBAYTAkJF
# MRkwFwYDVQQKExBHbG9iYWxTaWduIG52LXNhMTEwLwYDVQQDEyhHbG9iYWxTaWdu
# IFRpbWVzdGFtcGluZyBDQSAtIFNIQTI1NiAtIEcyMIIBIjANBgkqhkiG9w0BAQEF
# AAOCAQ8AMIIBCgKCAQEAqpuOw6sRUSUBtpaU4k/YwQj2RiPZRcWVl1urGr/SbFfJ
# MwYfoA/GPH5TSHq/nYeer+7DjEfhQuzj46FKbAwXxKbBuc1b8R5EiY7+C94hWBPu
# TcjFZwscsrPxNHaRossHbTfFoEcmAhWkkJGpeZ7X61edK3wi2BTX8QceeCI2a3d5
# r6/5f45O4bUIMf3q7UtxYowj8QM5j0R5tnYDV56tLwhG3NKMvPSOdM7IaGlRdhGL
# D10kWxlUPSbMQI2CJxtZIH1Z9pOAjvgqOP1roEBlH1d2zFuOBE8sqNuEUBNPxtyL
# ufjdaUyI65x7MCb8eli7WbwUcpKBV7d2ydiACoBuCQIDAQABo4HoMIHlMA4GA1Ud
# DwEB/wQEAwIBBjASBgNVHRMBAf8ECDAGAQH/AgEAMB0GA1UdDgQWBBSSIadKlV1k
# sJu0HuYAN0fmnUErTDBHBgNVHSAEQDA+MDwGBFUdIAAwNDAyBggrBgEFBQcCARYm
# aHR0cHM6Ly93d3cuZ2xvYmFsc2lnbi5jb20vcmVwb3NpdG9yeS8wNgYDVR0fBC8w
# LTAroCmgJ4YlaHR0cDovL2NybC5nbG9iYWxzaWduLm5ldC9yb290LXIzLmNybDAf
# BgNVHSMEGDAWgBSP8Et/qC5FJK5NUPpjmove4t0bvDANBgkqhkiG9w0BAQsFAAOC
# AQEABFaCSnzQzsm/NmbRvjWek2yX6AbOMRhZ+WxBX4AuwEIluBjH/NSxN8RooM8o
# agN0S2OXhXdhO9cv4/W9M6KSfREfnops7yyw9GKNNnPRFjbxvF7stICYePzSdnno
# 4SGU4B/EouGqZ9uznHPlQCLPOc7b5neVp7uyy/YZhp2fyNSYBbJxb051rvE9ZGo7
# Xk5GpipdCJLxo/MddL9iDSOMXCo4ldLA1c3PiNofKLW6gWlkKrWmotVzr9xG2wSu
# kdduxZi61EfEVnSAR3hYjL7vK/3sbL/RlPe/UOB74JD9IBh4GCJdCC6MHKCX8x2Z
# faOdkdMGRE4EbnocIOM28LZQuTCCBMYwggOuoAMCAQICDCRUuH8eFFOtN/qheDAN
# BgkqhkiG9w0BAQsFADBbMQswCQYDVQQGEwJCRTEZMBcGA1UEChMQR2xvYmFsU2ln
# biBudi1zYTExMC8GA1UEAxMoR2xvYmFsU2lnbiBUaW1lc3RhbXBpbmcgQ0EgLSBT
# SEEyNTYgLSBHMjAeFw0xODAyMTkwMDAwMDBaFw0yOTAzMTgxMDAwMDBaMDsxOTA3
# BgNVBAMMMEdsb2JhbFNpZ24gVFNBIGZvciBNUyBBdXRoZW50aWNvZGUgYWR2YW5j
# ZWQgLSBHMjCCASIwDQYJKoZIhvcNAQEBBQADggEPADCCAQoCggEBANl4YaGWrhL/
# o/8n9kRge2pWLWfjX58xkipI7fkFhA5tTiJWytiZl45pyp97DwjIKito0ShhK5/k
# Ju66uPew7F5qG+JYtbS9HQntzeg91Gb/viIibTYmzxF4l+lVACjD6TdOvRnlF4RI
# shwhrexz0vOop+lf6DXOhROnIpusgun+8V/EElqx9wxA5tKg4E1o0O0MDBAdjwVf
# ZFX5uyhHBgzYBj83wyY2JYx7DyeIXDgxpQH2XmTeg8AUXODn0l7MjeojgBkqs2Iu
# YMeqZ9azQO5Sf1YM79kF15UgXYUVQM9ekZVRnkYaF5G+wcAHdbJL9za6xVRsX4ob
# +w0oYciJ8BUCAwEAAaOCAagwggGkMA4GA1UdDwEB/wQEAwIHgDBMBgNVHSAERTBD
# MEEGCSsGAQQBoDIBHjA0MDIGCCsGAQUFBwIBFiZodHRwczovL3d3dy5nbG9iYWxz
# aWduLmNvbS9yZXBvc2l0b3J5LzAJBgNVHRMEAjAAMBYGA1UdJQEB/wQMMAoGCCsG
# AQUFBwMIMEYGA1UdHwQ/MD0wO6A5oDeGNWh0dHA6Ly9jcmwuZ2xvYmFsc2lnbi5j
# b20vZ3MvZ3N0aW1lc3RhbXBpbmdzaGEyZzIuY3JsMIGYBggrBgEFBQcBAQSBizCB
# iDBIBggrBgEFBQcwAoY8aHR0cDovL3NlY3VyZS5nbG9iYWxzaWduLmNvbS9jYWNl
# cnQvZ3N0aW1lc3RhbXBpbmdzaGEyZzIuY3J0MDwGCCsGAQUFBzABhjBodHRwOi8v
# b2NzcDIuZ2xvYmFsc2lnbi5jb20vZ3N0aW1lc3RhbXBpbmdzaGEyZzIwHQYDVR0O
# BBYEFNSHuI3m5UA8nVoGY8ZFhNnduxzDMB8GA1UdIwQYMBaAFJIhp0qVXWSwm7Qe
# 5gA3R+adQStMMA0GCSqGSIb3DQEBCwUAA4IBAQAkclClDLxACabB9NWCak5BX87H
# iDnT5Hz5Imw4eLj0uvdr4STrnXzNSKyL7LV2TI/cgmkIlue64We28Ka/GAhC4evN
# GVg5pRFhI9YZ1wDpu9L5X0H7BD7+iiBgDNFPI1oZGhjv2Mbe1l9UoXqT4bZ3hcD7
# sUbECa4vU/uVnI4m4krkxOY8Ne+6xtm5xc3NB5tjuz0PYbxVfCMQtYyKo9JoRbFA
# uqDdPBsVQLhJeG/llMBtVks89hIq1IXzSBMF4bswRQpBt3ySbr5OkmCCyltk5lXT
# 0gfenV+boQHtm/DDXbsZ8BgMmqAc6WoICz3pZpendR4PvyjXCSMN4hb6uvM0MIIF
# fzCCA2egAwIBAgIQGLXChEOQEpdBrAmKM2WmEDANBgkqhkiG9w0BAQsFADBSMRMw
# EQYKCZImiZPyLGQBGRYDY29tMRgwFgYKCZImiZPyLGQBGRYIRGVsb2l0dGUxITAf
# BgNVBAMTGERlbG9pdHRlIFNIQTIgTGV2ZWwgMSBDQTAeFw0xNTA5MDExNTA3MjVa
# Fw0zNTA5MDExNTA3MjVaMFIxEzARBgoJkiaJk/IsZAEZFgNjb20xGDAWBgoJkiaJ
# k/IsZAEZFghEZWxvaXR0ZTEhMB8GA1UEAxMYRGVsb2l0dGUgU0hBMiBMZXZlbCAx
# IENBMIICIjANBgkqhkiG9w0BAQEFAAOCAg8AMIICCgKCAgEAlPqNqqVpE41dp1s1
# +neM+Xv5zfUAKTrD10RAF9epFFmIIMH62VgMXOYYWBryNQaUAYPZlvv/Tt0cCKca
# 5XAWKp4DbBeblCmxfHsqEz3R/kzn/CHRHnQ3YMZRMorAccq82DdxKiwnw9o0W5SG
# D5A+zNXh9DjcCx0G5ROAaqiv7m3HYz2HrEvqdIuMkMoj7Y2ieMiw/PuIjVU8wmod
# ltkBmGoAeOOcVYaWBZTpKy0NC/xYL7eHfMKdgRaa30pFVeZliN8DMiN/exbfr6iu
# 00fQAsNxiZleH/6CLHuODdh+7KK00Wp2Wi9qz/IeOAGkj8j0jXFnnX5PHQWcVVv8
# E8sIK1S95xDxmhOsrMGkGA6G3F7a1qfI1WntvYBT98eUgZQ3whDqjypj622jjXLk
# UxlfuUeuBHB2+T9kSbapQHIhjAE3f97A/FOuzG0aerr6eNC5doNjOX31Bfp5W0Wk
# hbX8D0Aexf7v+OsboqFkAkaNzSS2oaX7+G3XAw2r+slDmyimr+boaLEo4vM+oFzF
# UeBQOXvjGBEnGtxXmSIPwsLu+HlhOvjtXINLbsczl2QWzC2arRPxx6HLr1hPj0ei
# yz7bKDPQ+N+U9l5OetL6NNFgppVDoqSVo5FUwh47wZKaqXZ8b1jPj/SS+IRsbKnC
# J37+YXfkA2Mid9x8oMyRfBfwed8CAwEAAaNRME8wCwYDVR0PBAQDAgGGMA8GA1Ud
# EwEB/wQFMAMBAf8wHQYDVR0OBBYEFL6LoCtmWVn6kHFRpaeoBkJOkDztMBAGCSsG
# AQQBgjcVAQQDAgEAMA0GCSqGSIb3DQEBCwUAA4ICAQCDTvRyLG76375USpexKf0B
# GCuYfW+o/6G18GRqZeYls7lO251xn7hfXacfEZIHCPoizan0yvZJtYUocXRMieo7
# 66Zwn8g4OgEZjJXsw81p0GlkylmdWhqO+sRuGyYvGY32MWZ16oz6x/CG+rseou2H
# sLLtlSV76D2XPnDutIAHI/S4is4A7F0V+oNX04aHpUXMb0Y1BkPKNF1gIlmf4rdt
# Rh6+2r374QP+Ruw+nJiPNwF7TF28wkz1iUXWK9FSmM1Q6+/uXxpx9qRFRwv+pCd/
# 07IneZ3GmxxTNJxSzzEJxIfwoJIn6HL9NYPltAZ7CuWYsm5TFY+x5TZ5qS/O6+nA
# Hd30T7K/q+H5hjp9tisYah3RiBOOU+iZvtUsr1XaLT7zizxnmp4ssHHryLhNkYu2
# uh/dT1/iq8SbM3fKGElML+mE7ZPAg2q2B76kgbY+GrEtzNnzwNfIwkh/IDKYJ9n6
# JU2yQ4oa5sJjTf5uHUhxV9Zd8/BZK8L3H5S7Iy3yCVLyq98xuUZ3ChL4FoKeS89u
# MrgKADP2xnAdIw1nnd67ZSPrTVk3sZO/uJVKTzjpU0V10sc27VmVx9YByc4o4xDo
# Q6+eAlUbNpuoFpchzdL2dx5JUalLl2T4jg4UIzKcidPhEmyU1ApKUXFQTbx0N8v1
# WC2UXROwuc0YDLR7v6RCLjCCBdwwggPEoAMCAQICEz4AAAAHOQSYtK/MwjQAAQAA
# AAcwDQYJKoZIhvcNAQELBQAwVDETMBEGCgmSJomT8ixkARkWA2NvbTEYMBYGCgmS
# JomT8ixkARkWCERlbG9pdHRlMSMwIQYDVQQDExpEZWxvaXR0ZSBTSEEyIExldmVs
# IDIgQ0EgMjAeFw0xODEyMDQyMDEyNTBaFw0yMzEyMDQyMDIyNTBaMGwxEzARBgoJ
# kiaJk/IsZAEZFgNjb20xGDAWBgoJkiaJk/IsZAEZFghkZWxvaXR0ZTEWMBQGCgmS
# JomT8ixkARkWBmF0cmFtZTEjMCEGA1UEAxMaRGVsb2l0dGUgU0hBMiBMZXZlbCAz
# IENBIDIwggEiMA0GCSqGSIb3DQEBAQUAA4IBDwAwggEKAoIBAQCvao3WR6CSYsMe
# I4sWzR6nXvczKc7voHTVzi/q3LbOD6j6YQNa/WnJeDITb2yf8BcIUXeLqm9dd64S
# in69YS3gTLT7ZFucodBp11g6IaA1R40tbWW9x2WDxYGMDoN+Hvq78bQMsSFEo1Ad
# mZRS/GGCO69u0ROyFtAgRt3E4jLFuzm1RWiNdEl00qNYnmaN4iLz2dEnKtJm+Cl2
# NH1xlB+m47ovgHlejoqJ/eg9kLmwEZam8o2SzgMrBup85GO8UmV55f3mv7zrRNhe
# oL+rdBAqN3NsA3n2a2JmLZAkcRD5Zk5I46EnJhRZpguRoafd4INeOPYH2iKNKqpe
# HFIbyWKPAgMBAAGjggGNMIIBiTAQBgkrBgEEAYI3FQEEAwIBATAjBgkrBgEEAYI3
# FQIEFgQURHtJJiaF3HfA4va/QlnTnPpod7wwHQYDVR0OBBYEFGmVYfUC8O4CaCIJ
# kuTjxIa0u/lpMBkGCSsGAQQBgjcUAgQMHgoAUwB1AGIAQwBBMAsGA1UdDwQEAwIB
# hjASBgNVHRMBAf8ECDAGAQH/AgEAMB8GA1UdIwQYMBaAFEcuNu60nP9cXhh8uBPh
# vqkgHhSzMFwGA1UdHwRVMFMwUaBPoE2GS2h0dHA6Ly9wa2kuZGVsb2l0dGUuY29t
# L0NlcnRFbnJvbGwvRGVsb2l0dGUlMjBTSEEyJTIwTGV2ZWwlMjAyJTIwQ0ElMjAy
# LmNybDB2BggrBgEFBQcBAQRqMGgwZgYIKwYBBQUHMAKGWmh0dHA6Ly9wa2kuZGVs
# b2l0dGUuY29tL0NlcnRFbnJvbGwvU0hBMkxWTDJDQTJfRGVsb2l0dGUlMjBTSEEy
# JTIwTGV2ZWwlMjAyJTIwQ0ElMjAyKDEpLmNydDANBgkqhkiG9w0BAQsFAAOCAgEA
# T4VkpKHJQHX5pk2FaNiXUHQKkZQXs/uD8lbhSdUgPqZCUaD7rml/aqzusVpA2GML
# zrsUcomq7xt4S9kOKIGQabSUeg681nGvzXrp0P8xOsXYUWqR9PIcEkfdDYs3pNce
# S98TAFl8+hKkMm2XMDaOpBz7AT6xb5ISKEybUWf/Gsdfmha1UzfCtIDVQUdWQcFD
# FQnFfVL4gcKfmwp7fq5bZi5l4/4kMM1OP1s10Og8PaAPhRkaYdQapDbaT82czXZS
# v0dqimBXWImTAJx9PbcWc5iqmNtrUxPsYCt2yGNByO3spCIa96MqfkiQQBISZxRr
# NT6pjMGtdR3Kij/rixmEBy/ITd4Ua4Za4TR09C8Lw/+ukmdV3D0G+3zRwqcwURAV
# Bvxwp62sVe0+yUYnckwmIiwbI9X8VYyCURk0YvKqfsXRZjnWtGOhSjT2EnxO87e4
# hrO4G9akInQvzAL6giL/K4UCzpl4qotDlYK8PzvmsceuGWx23nZaQQ3K21FgNduo
# HIvqVuslCf+u7Z/ZYCwguGb6xKIzDS1vpzkqMuSHa1gxsmLm+PzMyM4i9E9FFnbX
# vKf3P6SXyk0yXi7bB/KcG9t7QsITpZ7X+LA2+gWDY2LE1i7XLsOoOn5KaV70sTB6
# PoL5qaqOJAxswoJ0t2j1itrhsG7y/GUPhcG3kWq9V+EwggaOMIIFdqADAgECAhNl
# ADwUlwDjJaP1Xv9CAAAAPBSXMA0GCSqGSIb3DQEBCwUAMGwxEzARBgoJkiaJk/Is
# ZAEZFgNjb20xGDAWBgoJkiaJk/IsZAEZFghkZWxvaXR0ZTEWMBQGCgmSJomT8ixk
# ARkWBmF0cmFtZTEjMCEGA1UEAxMaRGVsb2l0dGUgU0hBMiBMZXZlbCAzIENBIDIw
# HhcNMTgxMTE0MTQxNjU1WhcNMjAxMDI5MTg1OTI0WjAZMRcwFQYDVQQDEw5IZWF0
# aGVyIE1pbGxlcjCCASIwDQYJKoZIhvcNAQEBBQADggEPADCCAQoCggEBANNqwwoQ
# yJmuG4mVQZhiZ9GyeXKwxRsjzRbeDmi5PkDuNKCF0zm/3nJqRyMSAXkNL9KElsqm
# 8lHwphrJfo/XJgxBRSkSY+4Y+Fh4pCeBQAevNXE2wA3A1sEsmaP4uxKgEUtbJEDS
# 35h9SEDvj+esroKB09wa6qFkaTjaWq6GnhYzHWts2BFTaJ3iHu+mNBdZRfYH0jgg
# HEcGGRZaMmrXhGm0mf9UmZxgZZG4/mu9ZFdLOgV3Spwy897XmjMdpzlBZtvgKn44
# UpXwfw5PxEK4Ygx+VbaPIwJ0sKRZrbYyLeaTleVBm1ckK76t+b2/sITPx3gv1Sv9
# 8ECrBEWRrqD2muECAwEAAaOCA3owggN2MDwGCSsGAQQBgjcVBwQvMC0GJSsGAQQB
# gjcVCIGBvUmFvoUTgtWbPIPXjgeG8ckKXIPK9y3C8zICAWQCAR4wEwYDVR0lBAww
# CgYIKwYBBQUHAwMwCwYDVR0PBAQDAgeAMBsGCSsGAQQBgjcVCgQOMAwwCgYIKwYB
# BQUHAwMwIAYDVR0RBBkwF4EVaGVtaWxsZXJAZGVsb2l0dGUuY29tMB0GA1UdDgQW
# BBTYjzXkhwjhstLpzEior3SlOAA+RDAfBgNVHSMEGDAWgBRplWH1AvDuAmgiCZLk
# 48SGtLv5aTCCATsGA1UdHwSCATIwggEuMIIBKqCCASagggEihoHSbGRhcDovLy9D
# Tj1EZWxvaXR0ZSUyMFNIQTIlMjBMZXZlbCUyMDMlMjBDQSUyMDIsQ049dXNhdHJh
# bWVlbTAwNCxDTj1DRFAsQ049UHVibGljJTIwS2V5JTIwU2VydmljZXMsQ049U2Vy
# dmljZXMsQ049Q29uZmlndXJhdGlvbixEQz1kZWxvaXR0ZSxEQz1jb20/Y2VydGlm
# aWNhdGVSZXZvY2F0aW9uTGlzdD9iYXNlP29iamVjdENsYXNzPWNSTERpc3RyaWJ1
# dGlvblBvaW50hktodHRwOi8vcGtpLmRlbG9pdHRlLmNvbS9DZXJ0ZW5yb2xsL0Rl
# bG9pdHRlJTIwU0hBMiUyMExldmVsJTIwMyUyMENBJTIwMi5jcmwwggFUBggrBgEF
# BQcBAQSCAUYwggFCMIHEBggrBgEFBQcwAoaBt2xkYXA6Ly8vQ049RGVsb2l0dGUl
# MjBTSEEyJTIwTGV2ZWwlMjAzJTIwQ0ElMjAyLENOPUFJQSxDTj1QdWJsaWMlMjBL
# ZXklMjBTZXJ2aWNlcyxDTj1TZXJ2aWNlcyxDTj1Db25maWd1cmF0aW9uLERDPWRl
# bG9pdHRlLERDPWNvbT9jQUNlcnRpZmljYXRlP2Jhc2U/b2JqZWN0Q2xhc3M9Y2Vy
# dGlmaWNhdGlvbkF1dGhvcml0eTB5BggrBgEFBQcwAoZtaHR0cDovL3BraS5kZWxv
# aXR0ZS5jb20vQ2VydGVucm9sbC91c2F0cmFtZWVtMDA0LmF0cmFtZS5kZWxvaXR0
# ZS5jb21fRGVsb2l0dGUlMjBTSEEyJTIwTGV2ZWwlMjAzJTIwQ0ElMjAyLmNydDAN
# BgkqhkiG9w0BAQsFAAOCAQEAqFnnDf3WnhUtTZO7fhCSm1vcLN5H7xh55Fhsrapj
# Ku0aCSvHgWlZ9xlH2DboVFoMd589lU6DQujvfcTTpqY9zQu97QdszH8Wfhk9mW2O
# vVA3hDjahCEt+2vahw3aqsoSZaPYAjaRAMmeq23olHjMnFXvYntZImHjJjcSUpe+
# KkWxpdMd9rgKRUj86EQ0CluNC3ro3yrai/IUiDqboZ0lvI7GZYDnNzJMZHI3CtTn
# eDvfgtMY+xU+5ra53hbp93TYgr32bktk7p3Qp2kENBLYV/D59CghE4wxJW0pZ/Sw
# VXaJx3xzOjeO6PyAZ8vQieiBaf+2IDHXIh62x8UqlT1RDDCCBskwggSxoAMCAQIC
# EzQAAAAFqIzfrA2XWTIAAAAAAAUwDQYJKoZIhvcNAQELBQAwUjETMBEGCgmSJomT
# 8ixkARkWA2NvbTEYMBYGCgmSJomT8ixkARkWCERlbG9pdHRlMSEwHwYDVQQDExhE
# ZWxvaXR0ZSBTSEEyIExldmVsIDEgQ0EwHhcNMTUxMDI5MTcyMTAzWhcNMjUxMDI5
# MTczMTAzWjBUMRMwEQYKCZImiZPyLGQBGRYDY29tMRgwFgYKCZImiZPyLGQBGRYI
# RGVsb2l0dGUxIzAhBgNVBAMTGkRlbG9pdHRlIFNIQTIgTGV2ZWwgMiBDQSAyMIIC
# IjANBgkqhkiG9w0BAQEFAAOCAg8AMIICCgKCAgEAmPb6sHLB25JD286NfyR2RfuN
# gmSXaR2dLojx7rPDqiEWKM01mSdquzeXj7Qu/VQsiQLV/9oxwMArSvjJHRjeQ2L7
# orPGytxWiO6nNHkKbPUCkBTmRALVcXK0iYmXhQjaypjx5y8bi3K13AR7axTbNlPE
# Fy3z9TwFGftmeJOIvle3dBvOCxJre1mxmf544tkzq+Df0ENP8sA41WeQbA5ZyDa2
# C8PWm8XL59X00UgtMJcOq4fCG+xkjl7nnbQ4/AP7lGHGkl0bnYE5Xd/nVA86+wO+
# uTUcmbs0fJ9fKO3bq3wgiUaRyyBbUQ2NzGlgaffxqge2lM3WCmiQeHKyfKsOkfg4
# 1+6h7qUFywDoDkvnVBjJs2+tgImqqD6iwmgZWHt6PeIiwJA/IIKBf0t1O16G39ui
# m6NSiesSK+wfOMxyxZio/BzKGPOtv4PwosBlPKlhK5bbvMWY2RFsWQJ6LPiRXlE5
# NIYbh/CTyngIdM6Drwr57sIZGWbKCJc9nORteVgx3pgciFAxOFGn1k3zmxM83qYx
# xgKi6fql8KCgbo+l6luROLa5rsRfkGPtRXy1HWJ7xwcf8/JxLJGlp1rtnGnZljvb
# 0Tbtwo8GwDoihSMSh9MoGrJTrtk8tnYf4UpLgGKjGyGOUBFGrRGQcEhWbzDTK5qZ
# P/0f31d3CndzQORYAb8CAwEAAaOCAZQwggGQMBAGCSsGAQQBgjcVAQQDAgEBMCMG
# CSsGAQQBgjcVAgQWBBRF4tTkKKaihh8hZlZ2wn5W1acT+zAdBgNVHQ4EFgQURy42
# 7rSc/1xeGHy4E+G+qSAeFLMwEQYDVR0gBAowCDAGBgRVHSAAMBkGCSsGAQQBgjcU
# AgQMHgoAUwB1AGIAQwBBMAsGA1UdDwQEAwIBhjASBgNVHRMBAf8ECDAGAQH/AgEB
# MB8GA1UdIwQYMBaAFL6LoCtmWVn6kHFRpaeoBkJOkDztMFgGA1UdHwRRME8wTaBL
# oEmGR2h0dHA6Ly9wa2kuZGVsb2l0dGUuY29tL0NlcnRFbnJvbGwvRGVsb2l0dGUl
# MjBTSEEyJTIwTGV2ZWwlMjAxJTIwQ0EuY3JsMG4GCCsGAQUFBwEBBGIwYDBeBggr
# BgEFBQcwAoZSaHR0cDovL3BraS5kZWxvaXR0ZS5jb20vQ2VydEVucm9sbC9TSEEy
# TFZMMUNBX0RlbG9pdHRlJTIwU0hBMiUyMExldmVsJTIwMSUyMENBLmNydDANBgkq
# hkiG9w0BAQsFAAOCAgEAUIIxw2cOQAxpWz1ZyL6PUsJPtdtzaxKmz4Tsw48uWk/l
# TbmWm7bD0WbFIlWwZ5DREGa9G99F0L3f+CO8Bqn+T6Jcw6xQ6Po53cXG4NSgoL6V
# v6CIfKVg9UwgcIj4J49sjTgiY7pn+wav9EKXM99AxPpNqxjLhRvTBk6Mbdg2ifED
# ljdc12PBWrHOE1M72cngFDkdRNboPpLylH8wUC3PojELdMIWC80//HOqLFsM07FM
# JaHHLB95oDuP+7+B0Q8n22MQVKyPihVAVDE6rhiAI7b2dt0C5vweubo0bTTIWhBA
# x5RO6b7/J1shCGb33HBxoAqX40i6AHaX6t+hapLCwYn1jGI0Ba57U0MeoLTrg77O
# KdqxwaJRauS8pORzZIJMEcJztATZaFf9cTKm8rD7EcvEfJib0I/ydR6chS55RWgD
# h8GlPoikRKW8xIomoA/iCKYMrroq5E6rY3ChgoYb3OwvtiTNpYKLsCVjRn4KieEm
# x4wl4h77RFywMjnGISoj56wrrk4jePpxjfiTHQVGx/6nQYx22IYPkMTEcMqVtT0Y
# Omd0rISvbwdSbyuozw923cC3lF86FoZAz1F5muSdE2VeejZYe7eYBxOeHHKk+/LA
# 3La7TCE/j/wWzN31mpOgQq62ct+HdG9o7EX/ITmwN7EDM4Aa4oMZytupX8iO61gx
# ggS6MIIEtgIBATCBgzBsMRMwEQYKCZImiZPyLGQBGRYDY29tMRgwFgYKCZImiZPy
# LGQBGRYIZGVsb2l0dGUxFjAUBgoJkiaJk/IsZAEZFgZhdHJhbWUxIzAhBgNVBAMT
# GkRlbG9pdHRlIFNIQTIgTGV2ZWwgMyBDQSAyAhNlADwUlwDjJaP1Xv9CAAAAPBSX
# MA0GCWCGSAFlAwQCAQUAoEwwGQYJKoZIhvcNAQkDMQwGCisGAQQBgjcCAQQwLwYJ
# KoZIhvcNAQkEMSIEIFbCkK14wuTgIY6ISubfGQXm3s1ckdeQjnT2Jc2Vux+sMA0G
# CSqGSIb3DQEBAQUABIIBAFBzsfpKb5PMLy1h/Rv3C1uaEcNpRCTBIhMGMLYjkZrs
# VSgyebp/TmrCahjEBFBMo9ZvNO3B+vaCQWbGi4maJ+jFcqW/+KsWuE5gN/cfJL3b
# 8ncZSt2U+K5AnsNy9R9GJ/T19zfxddXYHntKsGT3alFMqcG57W/00gWmHgpYwqie
# 0YkOJGF6mmqh6K/zu4NTkp7RH5Qbd14Bx6YTNJYMx4SRdwHhGtFiZ9JAwS86uUra
# 7oUFPzByXnVoHDq2hPcmvBm/io3MzUaRYvbn9b39jqQ4Mbj+cLiNb4LW7diPc2XJ
# fargQmuIlNo9j4+ww+fScNoHZcRUiy9plKritVG0nlGhggK5MIICtQYJKoZIhvcN
# AQkGMYICpjCCAqICAQEwazBbMQswCQYDVQQGEwJCRTEZMBcGA1UEChMQR2xvYmFs
# U2lnbiBudi1zYTExMC8GA1UEAxMoR2xvYmFsU2lnbiBUaW1lc3RhbXBpbmcgQ0Eg
# LSBTSEEyNTYgLSBHMgIMJFS4fx4UU603+qF4MA0GCWCGSAFlAwQCAQUAoIIBDDAY
# BgkqhkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0xOTEwMTYy
# MDM3NTdaMC8GCSqGSIb3DQEJBDEiBCD31pKSNBoGRu2iOBOrxd4BPzDOwNlZMbFN
# 8R2j1cJrrzCBoAYLKoZIhvcNAQkQAgwxgZAwgY0wgYowgYcEFD7HZtXU1HLiGx8h
# Q1IcMbeQ2UtoMG8wX6RdMFsxCzAJBgNVBAYTAkJFMRkwFwYDVQQKExBHbG9iYWxT
# aWduIG52LXNhMTEwLwYDVQQDEyhHbG9iYWxTaWduIFRpbWVzdGFtcGluZyBDQSAt
# IFNIQTI1NiAtIEcyAgwkVLh/HhRTrTf6oXgwDQYJKoZIhvcNAQEBBQAEggEASBg+
# pOS1yjNNyptYHPfF9q/CR9hLfDEoDs32vQVZLSQVj6BcDWsNr7NIEj5Lg2BI7DKB
# J5WQ8H7ISmpUfNQAPAaD/oGRF6O1667sn9j9P2HPVOpIA5X6sHG7wpGM5lrAht+I
# LmbW4xI/a+SfMEJAm0cJHqk0OcUfbg3IminzpoBUSuiJVKQ+3wBgGFNuaxk4I/X6
# pt6m6ABjnHtZ8+RQO/T37CtcuvDN5s0TdEiFPhstn9qbike3Pohta7+I0/ODKdQo
# JScNzvwa3f4QGTKgsvUE9cnOiplbl80Mcf6Mlz5j8nt7SxUP+6MMQlqFobT5eVb1
# qufu5w3fjls3WHEHQA==
# SIG # End signature block
