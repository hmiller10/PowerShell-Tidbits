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

Function Test-PathExists
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
} #end function Test-PathExists

Function Get-UtcTime
{
	#Begin function to get UTC date and time
	[System.DateTime]::UtcNow
} #End Get-UtcTime

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
$dtmScriptStartTimeUTC = Get-UtcTime
$dtmFormatString = "yyyy-MM-dd HH:mm:ss"
$transcriptFileName = "{0}-{1}-Transcript.txt" -f $dtmScriptStartTimeUTC.ToString("yyyy-MM-dd_HH.mm.ss"), "$($thisServer)-CARoleBackup"

$scriptDir = Split-Path $MyInvocation.MyCommand.Path
$scriptName = $MyInvocation.ScriptName

#Region Check folder structures
Test-PathExists -Path $BackupFolder -PathType Folder
$TodaysFldr = "{0}\{1}" -f $BackupFolder, "CABackup_$(Get-ReportDate)"
Test-PathExists -Path $TodaysFldr -PathType Folder
#EndRegion

try
{
	# start transcript file
	Start-Transcript ("{0}\{1}" -f $TodaysFldr, $transcriptFileName)
	Write-Verbose -Message ("[{0} UTC] [SCRIPT] Beginning execution of script." -f $dtmScriptStartTimeUTC.ToString($dtmFormatString))
	Write-Verbose -Message ("[{0} UTC] [SCRIPT] Script Name             		:  {1}" -f $(Get-UtcTime).ToString($dtmFormatString), $scriptName)
	Write-Verbose -Message ("[{0} UTC] [SCRIPT] Script Directory path   		:  {1}" -f $(Get-UtcTime).ToString($dtmFormatString), $scriptDir)
	Write-Verbose -Message ("[{0} UTC] [SCRIPT] Main Backup Folder path  		:  {1}" -f $(Get-UtcTime).ToString($dtmFormatString), $BackupFolder)
	Write-Verbose -Message ("[{0} UTC] [SCRIPT] Todays Backup Folder path     	:  {1}" -f $(Get-UtcTime).ToString($dtmFormatString), $TodaysFldr)
	
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
	Test-PathExists -Path $certBkpFldr -PathType Folder
	
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
	Test-PathExists -Path $RegFldr -PathType Folder
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
	}
	#EndRegion
	
	#Region Backup-Policy-File
	
	#If not using a Policy Certificate Authority server and policies are implemented using .INF file, backup configuration file.
	#Backup Certificate Policy .Inf file
	$PolicyFldr = "{0}\{1}" -f $TodaysFldr, "PolicyFile"
	Test-PathExists -Path $PolicyFldr -PathType Folder
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
	}
	#EndRegion
	
	#Region Backup-IIS-Files
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
					Test-PathExists -Path $IISCustomFldr -PathType Folder
					
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
			Test-PathExists -Path $templateFldr -PathType Folder
			
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
			Test-PathExists -Path $hsmBkpRoot -PathType Folder
			
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
		}
		
	}
	
	#EndRegion
	
	#Stop the stopwatch	
	$stopWatch.Stop()
	
	$dtmScriptStopTimeUTC = Get-UtcTime
	$elapsedTime = New-TimeSpan -Start $dtmScriptStartTimeUTC -End $dtmScriptStopTimeUTC
	$runtime = $stopWatch.Elapsed.ToString('dd\.hh\:mm\:ss')
	
	$reportBody | ConvertTo-Html -Title "$thisServer CA backup report as of $(Get-TodaysDate)" -Body $reportBody -PostContent "Script took: $($runTime) to run."
	
	Write-Verbose -Message ("[{0} UTC] [SCRIPT] Script Complete" -f $(Get-UtcTime).ToString($dtmFormatString))
	Write-Verbose -Message ("[{0} UTC] [SCRIPT] Script Start Time :  {1}" -f $(Get-UtcTime).ToString($dtmFormatString), $dtmScriptStartTimeUTC.ToString($dtmFormatString))
	Write-Verbose -Message ("[{0} UTC] [SCRIPT] Script Stop Time  :  {1}" -f $(Get-UtcTime).ToString($dtmFormatString), $dtmScriptStopTimeUTC.ToString($dtmFormatString))
	Write-Verbose -Message ("[{0} UTC] [SCRIPT] Elapsed Time: {1:N0}.{2:N0}:{3:N0}:{4:N1}  (Days.Hours:Minutes:Seconds)" -f $(Get-UtcTime).ToString($dtmFormatString), $elapsedTime.Days, $elapsedTime.Hours, $elapsedTime.Minutes, $elapsedTime.Seconds)
	
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