#Region Help

<#

.NOTES
THIS CODE IS MADE AVAILABLE AS IS, WITHOUT WARRANTY OF ANY KIND. THE
ENTIRE RISK OF THE USE OR THE RESULTS FROM THE USE OF THIS CODE REMAINS
WITH THE USER.

.SYNOPSIS
GPO backup script with archive cleanup and e-mail notification.

.DESCRIPTION
This script runs from a scheduled task on one domain controller in the domain for the purpose
of backing up all Group Policy objects for the current domain. The backed up GPOs are in a
dated folder which is the zipped up and archived. The results of script execution are then e-mailed
to the named parties.

.LINK
https://msdn.microsoft.com/en-us/library/hh875104(v=vs.110).aspx?cs-save-lang=1&cs-lang=vb#code-snippet-1

.OUTPUTS
Group policy standard backups to pre-defined destination folder

.EXAMPLE 
	PS C:\>Backup-Domain-GPOs-with-notify.ps1

#>

###########################################################################
#
#
# AUTHOR:  Heather Miller
#          
#
# VERSION HISTORY:
# 6.0 08/21/2018 - Added $messageBody @@CopyrightYear@@ syntax
#
###########################################################################

#EndRegion

#Region ExecutionPolicy
#Set Execution Policy
#Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Confirm:$false -Force
#EndRegion

#Region Modules
#Check if required module is loaded, if not load import it
Try
{
	Import-Module ActiveDirectory -ErrorAction Stop
}
Catch
{
	Throw "Active Directory module could not be loaded. $($_.Exception.Message)"
}

Try
{
	Import-Module GroupPolicy -ErrorAction Stop
}
Catch
{
	Throw "Group Policy module could not be loaded. $($_.Exception.Message)"
}

#EndRegion

#Region Variables
$Domain = Get-ADDomain -Current LocalComputer
$domDNS = ($Domain).dnsRoot
$domainName = $domDNS.Replace('.', '_')
$pdcE = ($Domain).pdcEmulator
$dtmFormatString = "yyyy-MM-dd HH:mm:ss"
$Limit = (Get-Date).AddDays(-60)
$moveLimit = (Get-Date).AddDays(-30)

Add-Type -AssemblyName "System.IO.Compression.FileSystem"
#EndRegion

#Region Functions

Function Get-MyInvocation
{
	#Begin function to define $MyInvocation
	Return $MyInvocation
} #End function Get-MyInvocation

Function Get-ReportDate
{
	#Begin function get report execution date
	Get-Date -Format "yyyy-MM-dd"
} #End function Get-ReportDate

Function Get-LongDate
{
	#Begin function to get date and time in long format
	Get-Date -Format G
} #End function Get-LongDate

Function Get-UTCTime
{
<#
.SYNOPSIS
Gets current date and time in UTC format

.DESCRIPTION
Gets current date and time in UTC format

.INPUTS
None

.OUTPUTS
None

.EXAMPLE
Get-UtcTime

#>
	[System.DateTime]::UtcNow
	
} #End function Get-UTCTime

Function Send-SmtpRelayMessage
{
	[CmdletBinding()]
	Param
	(
		[Parameter(Mandatory = $true)]
		$To,
		[Parameter(Mandatory = $false)]
		$Cc,
		[Parameter(Mandatory = $true)]
		$From,
		[Parameter(Mandatory = $false)]
		$ReplyTo,
		[Parameter(Mandatory = $true)]
		$Subject,
		[Parameter(Mandatory = $true)]
		$Body,
		[Parameter(Mandatory = $true)]
		$SMTPServer,
		[Parameter(Mandatory = $false)]
		$Port = 25,
		[Parameter(Mandatory = $false)]
		$InlineImageAttachments,
		[Parameter(Mandatory = $false)]
		$Attachments
	)
	
	
	$objMailMessage = New-Object System.Net.Mail.MailMessage
	$objSmtpClient = New-Object System.Net.Mail.SmtpClient
	
	$objSmtpClient.host = $SMTPServer
	$objSmtpClient.Port = $port
	
	ForEach ($recipient in $To) { $objMailMessage.To.Add((New-Object System.Net.Mail.MailAddress($recipient))) }
	if ($PSBoundParameters.ContainsKey('CC')) { ForEach ($recipient in $Cc) { $objMailMessage.Cc.Add((New-Object System.Net.Mail.MailAddress($recipient))) } }
	$objMailMessage.From = $From
	$objMailMessage.Sender = $From
	$objMailMessage.ReplyTo = $ReplyTo
	$objMailMessage.Subject = $Subject
	$objMailMessage.Body = $Body
	$objMailMessage.IsBodyHtml = $true
	
	ForEach ($inlineAttachment in $InlineImageAttachments)
	{
		$objAttachment = New-Object System.Net.Mail.Attachment($inlineAttachment.FullName.ToString())
		$objAttachment.ContentDisposition.Inline = $true
		$objAttachment.ContentDisposition.DispositionType = [System.Net.Mime.DispositionTypeNames]::Inline
		$objAttachment.ContentType.MediaType = "image/" + $inlineAttachment.Extension.ToString().Substring(1)
		$objAttachment.ContentType.Name = $inlineAttachment.Name.ToString()
		$objAttachment.ContentId = $inlineAttachment.Name.ToString()
		$objMailMessage.Attachments.Add($objAttachment)
	}
	
	ForEach ($attachment in $Attachments)
	{
		$objAttachment = New-Object System.Net.Mail.Attachment($attachment.FullName.ToString())
		$objAttachment.ContentDisposition.Inline = $false
		$objAttachment.ContentDisposition.DispositionType = [System.Net.Mime.DispositionTypeNames]::Attachment
		$objAttachment.ContentType.MediaType = [System.Net.Mime.MediaTypeNames+Text]::Plain
		$objMailMessage.Attachments.Add($objAttachment)
	}
	
	$objSmtpClient.Send($objMailMessage)
	
	$objAttachment.Dispose()
	$objMailMessage.Dispose()
	$objSmtpClient.Dispose()
} #End function Send-SMTPRelayMessage

Function Get-SMTPServer
{
	#Begin function to get SMTP server for AD forest
	[CmdletBinding()]
	Param
	(
		[Parameter(Mandatory = $true)]
		[string]$Domain
	)
	
	Switch -Wildcard ($Domain)
	{
		'example.com' { $smtpServer = "relaymail.example.com" }
		'child.example.com' { $smtpServer = "relaymail.example.com" }
		
		default { $smtpserver = "appmail.example.com" }
	}
	
	`	$out = [PSCustomObject] @{
		smtpServer = $smtpServer
		Port	      = '25'
	}
	
	Return $out
	
} #end function Get-SMTPServer

Function Test-PathExists
{
<#
.SYNOPSIS
Checks if a path to a file or folder exists, and creates it if it does not exist.

.DESCRIPTION
Checks if a path to a file or folder exists, and creates it if it does not exist.

.PARAMETER Path
Full path to the file or folder to be checked

.PARAMETER PathType
Valid options are "File" and "Folder", depending on which to check.

.OUTPUTS
None

.EXAMPLE
Test-PathExists -Path "C:\temp\SomeFile.txt" -PathType File
	
.EXAMPLE
Test-PathExists -Path "C:\temp" -PathFype Folder

#>
	[CmdletBinding()]
	Param
	(
		[Parameter(Mandatory, Position = 0)]
		[string]$Path,
		[Parameter(Mandatory, Position = 1)]
		[ValidateSet("File", "Folder")]
		[string]$PathType
	)
	
	Switch ($PathType)
	{
		File
		{
			If ((Test-Path -Path $Path -PathType Leaf) -eq $true)
			{
				Write-Information -Message "File: $Path already exists.."
			}
			Else
			{
				New-Item -Path $Path -ItemType File -Force
				Write-Information -Message "File: $Path not present, creating new file..."
			}
		}
		Folder
		{
			If ((Test-Path -Path $Path -PathType Container) -eq $true)
			{
				Write-Information -Message "Folder: $Path already exists..."
			}
			Else
			{
				New-Item -Path $Path -ItemType Directory -Force
				Write-Information -Message "Folder: $Path not present, creating new folder."
			}
		}
	}
	
} #end function Test-PathExists

#EndRegion






#Region Script
$Error.Clear()

#Start Function timer, to display elapsed time for function. Uses System.Diagnostics.Stopwatch class - see here: https://msdn.microsoft.com/en-us/library/system.diagnostics.stopwatch(v=vs.110).aspx 
$stopWatch = [System.Diagnostics.Stopwatch]::StartNew()
$dtmScriptStartTimeUTC = Get-UTCTime
$transcriptFileName = "{0}-{1}-Transcript.txt" -f $dtmScriptStartTimeUTC.ToString("yyyy-MM-dd_HH.mm.ss"), "GroupPolicyBackups"

$myInv = Get-MyInvocation
$scriptDir = Split-Path $MyInvocation.MyCommand.Path
$scriptName = $myInv.ScriptName

#Check to see if destination backup folder is present and accessible
[String]$gpoBkpDestFldr = 'E:\GPOBackups'
Test-PathExists -Path $gpoBkpDestFldr -PathType Folder

#Create working directory for today's backups
[String]$workingDir = "{0}\{1}" -f $scriptDir, "WorkingDir"
Test-PathExists -Path $workingDir -PathType Folder

#Create .zip archive folder if needed
$archiveFolder = "{0}\{1}" -f $gpoBkpDestFldr, "Archives"
Test-PathExists -Path $archiveFolder -PathType Folder

# start transcript file
Start-Transcript ("{0}\{1}" -f $workingDir, $transcriptFileName)
Write-Verbose -Message ("[{0} UTC] [SCRIPT] Beginning execution of script." -f $dtmScriptStartTimeUTC.ToString($dtmFormatString))
Write-Verbose -Message ("[{0} UTC] [SCRIPT] Script Name             :  {1}" -f $(Get-UTCTime).ToString($dtmFormatString), $scriptName)
Write-Verbose -Message ("[{0} UTC] [SCRIPT] Script Directory path   :  {1}" -f $(Get-UTCTime).ToString($dtmFormatString), $scriptDir)
Write-Verbose -Message ("[{0} UTC] [SCRIPT] Working Directory path  :  {1}" -f $(Get-UTCTime).ToString($dtmFormatString), $workingDir)
Write-Verbose -Message ("[{0} UTC] [SCRIPT] Archive folder path     :  {1}" -f $(Get-UTCTime).ToString($dtmFormatString), $archiveFolder)

#Create array of all GPO objects
$Error.Clear()
Try
{
	[Array]$GPOs = Get-GPO -All -Domain $domDNS -Server $pdcE | Select-Object DisplayName, ID
	if (!($?))
	{
		Try
		{
			[Array]$GPOs = Get-GPO -All -Domain $domDNS -Server $domDNS | Select-Object DisplayName, ID
		}
		Catch
		{
			$errorMessage = "{0}: {1}" -f $Error[0], $Error[0].InvocationInfo.PositionMessage
			Write-Error $errorMessage -ErrorAction Continue
		}
	}
	
	#Loop through all GPOs in array and backup to local directory
	ForEach ($GPO in $GPOs)
	{
		$gpoDisplayName = ($GPO).DisplayName -replace (" ", "_")
		$gpoGuid = ($GPO).ID
		$Comment = "{0}_{1}" -f $gpoDisplayName, $(Get-ReportDate)
		$Error.Clear()
		Try
		{
			Backup-GPO -Guid $gpoGuid -Path $workingDir -Comment $Comment -Domain $domDNS -Server $pdcE -WarningAction Continue
			if (!($?))
			{
				Try
				{
					Backup-GPO -Guid $gpoGuid -Path $workingDir -Comment $Comment -Domain $domDNS -Server $domDNS -ErrorAction Stop -WarningAction Continue
				}
				Catch
				{
					$errorMessage = "{0}: {1}" -f $Error[0], $Error[0].InvocationInfo.PositionMessage
					Write-Error $errorMessage -ErrorAction Continue
				}
			}
		}
		Catch
		{
			$errorMessage = "{0}: {1}" -f $Error[0], $Error[0].InvocationInfo.PositionMessage
			Write-Error $errorMessage -ErrorAction Continue
		}
		$null = $GPO = $gpoDisplayName = $gpoGuid = $Comment
	}
	
}
Catch
{
	$errorMessage = "{0}: {1}" -f $Error[0], $Error[0].InvocationInfo.PositionMessage
	Write-Error $errorMessage -ErrorAction Continue
}

[xml]$manifest = Get-Content (Join-Path $workingDir 'manifest.xml')
foreach ($gpBackup in $manifest.Backups.BackupInst)
{
	Rename-Item -Path (Join-Path $workingDir $gpBackup.Id.InnerText) -NewName ($gpBackup.GPODisplayName.InnerText -replace '[:\\/]', '')
}

Stop-Transcript
Start-Sleep -Seconds 10

#Save output
#Compress the GPO folders into a single archive
$archiveFile = "{0}\{1}" -f $gpoBkpDestFldr, "$($domainName)_archive_for_$(Get-ReportDate).zip"
$compressionLevel = [System.IO.Compression.CompressionLevel]::Optimal
If ((Test-Path -Path $archiveFile -PathType Leaf) -eq $true) { Remove-Item -Path $archiveFile -Confirm:$false }
#See https://msdn.microsoft.com/en-us/library/hh875104(v=vs.110).aspx?cs-save-lang=1&cs-lang=vb#code-snippet-1
[IO.Compression.ZipFile]::CreateFromDirectory($workingDir, $archiveFile, $compressionLevel, $false)


#Stop the stopwatch	
$stopWatch.Stop()

#Send e-mail notification	
$emailTemplatePath = "{0}\{1}" -f $scriptDir, "EmailTemplates"
$imageAttachmentPath = "{0}\{1}" -f $scriptDir, "Images"
$emailTemplateFileName = "IAM_Weekly_GPOBkpNotification.html"
$smtpInfo = Get-SMTPServer -Domain $domDNS

$runTime = $stopWatch.Elapsed.ToString('dd\.hh\:mm\:ss')
Write-Verbose -Message ("[{0}] Sending email notification, please wait..." -f $(Get-UTCTime).ToString($dtmFormatString))
$emailTemplate = "{0}\{1}" -f $emailTemplatePath, $emailTemplateFileName

$htmlTemplate = [System.IO.StreamReader]$emailTemplate
$messageBody = $htmlTemplate.ReadToEnd()
$htmlTemplate.Dispose()

$messageBody = $messageBody.Replace("@@Date@@", $(Get-Date -Format MMM-dd-yyyy))
$messageBody = $messageBody.Replace("@@DomainName@@", $domDNS)
$messageBody = $messageBody.Replace("@@ScheduledTaskName@@", "IAM.Weekly.BackupDomainGPOs")
$messageBody = $messageBody.Replace("@@ServerName@@", [System.Net.Dns]::GetHostByName("LocalHost").HostName)
$messageBody = $messageBody.Replace("@@ScriptName@@", "$($myInv.ScriptName)")
$messageBody = $messageBody.Replace("@@ScriptRunTime@@", $runTime)
$messageBody = $messageBody.Replace("@@CopyrightYear@@", $(Get-Date -Format yyyy))

$colInlineImageAttachments = Get-ChildItem -Path $imageAttachmentPath
$colAttachments = Get-ChildItem -Path $workingDir -Filter *.txt

$params = @{
	To	      = "me@example.com"
	CC	      = "you@example.com"
	From	      = "GPOBackupNotifications@example.com"
	ReplyTo    = "GPOBackupNotifications@example.com"
	SMTPServer = $smtpInfo.smtpServer
	Port	      = $smtpInfo.Port
	InlineImageAttachments = $colInlineImageAttachments
	Attachments = $colAttachments
}


If ((Test-Path -Path $archiveFile -PathType Leaf) -eq $true)
{
	$messageBody = $messageBody.Replace("@@ScriptStatus@@", "Success")
	$params.Subject = "SUCCESS: $($domainName)_GPO_Backup_Report_as_of_$(Get-ReportDate)"
	$params.Body = $messageBody
}
Else
{
	$messageBody = $messageBody.Replace("@@ScriptStatus@@", "Failed")
	$params.Subject = "FAILED: $($domainName)_GPO_Backup_Report_as_of_$(Get-ReportDate)"
	$params.Body = $messageBody
}

Send-SmtpRelayMessage @params
Get-ChildItem -Path $scriptDir | Where-Object { ($_.PsIsContainer) -and ($_.Name -like "workingDir*") } | Remove-Item -Recurse -Force -Confirm:$false

#Close out script
$dtmScriptStopTimeUTC = Get-UTCTime
$elapsedTime = New-TimeSpan -Start $dtmScriptStartTimeUTC -End $dtmScriptStopTimeUTC
Write-Verbose -Message ("[{0} UTC] [SCRIPT] Script Complete" -f $(Get-UTCTime).ToString($dtmFormatString))
Write-Verbose -Message ("[{0} UTC] [SCRIPT] Script Start Time :  {1}" -f $(Get-UTCTime).ToString($dtmFormatString), $dtmScriptStartTimeUTC.ToString($dtmFormatString))
Write-Verbose -Message ("[{0} UTC] [SCRIPT] Script Stop Time  :  {1}" -f $(Get-UTCTime).ToString($dtmFormatString), $dtmScriptStopTimeUTC.ToString($dtmFormatString))
Write-Verbose -Message ("[{0} UTC] [SCRIPT] Elapsed Time: {1:N0}.{2:N0}:{3:N0}:{4:N1}  (Days.Hours:Minutes:Seconds)" -f $(Get-UTCTime).ToString($dtmFormatString), $elapsedTime.Days, $elapsedTime.Hours, $elapsedTime.Minutes, $elapsedTime.Seconds)

Get-ChildItem -Path $scriptDir | Where-Object { ($_.PsIsContainer) -and ($_.Name -like "workingDir*") } | Remove-Item -Recurse -Force -Confirm:$false

#Clean up archive files
$filesToMove = Get-ChildItem -Path $gpoBkpDestFldr | Where-Object { (-not($_.PSIsContainer)) -and ($_.CreationTime -lt $moveLimit) }
$filesToMove | ForEach-Object { Move-Item -Path $_.FullName -Destination $archiveFolder -Force -Confirm:$false }

Get-ChildItem -Path $archiveFolder -Recurse | Where-Object { $_.LastWriteTime -lt $Limit } | Remove-Item -Force -Confirm:$false
#EndRegion