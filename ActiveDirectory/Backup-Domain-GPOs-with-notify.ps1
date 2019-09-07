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
.\Backup-Domain-GPOs-with-notify.ps1

#>

###########################################################################
#
#
# AUTHOR:  Heather Miller
#          
#
# VERSION HISTORY:
# 7.0 04/01/2019 - Cleaned up spacing
#
###########################################################################


#Region ExecutionPolicy
#Set Execution Policy
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Confirm:$false -Force
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
#Define variables
$Domain = Get-ADDomain -Current LocalComputer
$domDNS = ($Domain).dnsRoot
$domainName = $domDNS.Replace('.', '_')
$pdcE = ($Domain).pdcEmulator
$dtmFormatString = "yyyy-MM-dd HH:mm:ss"
$Limit = (Get-Date).AddDays(-60)
$moveLimit = (Get-Date).AddDays(-30)
Add-Type -AssemblyName WindowsBase
Add-Type -AssemblyName "System.IO.Compression.FileSystem"
$VerbosePreference = "Continue"
#EndRegion

#Region Functions
Function fnGet-MyInvocation {#Begin function to define $MyInvocation
	Return $MyInvocation
}#End function fnGet-MyInvocation

Function fnGet-ReportDate {#Begin function get report execution date
	Get-Date -Format "yyyy-MM-dd"
}#End function fnGet-ReportDate

Function fnGet-LongDate {#Begin function to get date and time in long format
	Get-Date -Format G
}#End function fnGet-LongDate

Function fnUTC-Now {#Begin function to get date and time in UTC format
	[System.DateTime]::UtcNow
}#End function fnUTC-Now

Function fnSend-SmtpRelayMessage {
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

    ForEach ( $recipient in $To ) { $objMailMessage.To.Add((New-Object System.Net.Mail.MailAddress($recipient))) }
    ForEach ( $recipient in $Cc ) { $objMailMessage.Cc.Add((New-Object System.Net.Mail.MailAddress($recipient))) }
    $objMailMessage.From = $From
    $objMailMessage.Sender = $From
    $objMailMessage.ReplyTo = $ReplyTo
    $objMailMessage.Subject = $Subject
    $objMailMessage.Body = $Body
    $objMailMessage.IsBodyHtml = $true

    ForEach ( $inlineAttachment in $InlineImageAttachments )
    {
        $objAttachment = New-Object System.Net.Mail.Attachment($inlineAttachment.FullName.ToString())
        $objAttachment.ContentDisposition.Inline = $true
        $objAttachment.ContentDisposition.DispositionType = [System.Net.Mime.DispositionTypeNames]::Inline
        $objAttachment.ContentType.MediaType = "image/" + $inlineAttachment.Extension.ToString().Substring(1)
        $objAttachment.ContentType.Name = $inlineAttachment.Name.ToString()
        $objAttachment.ContentId = $inlineAttachment.Name.ToString()
        $objMailMessage.Attachments.Add($objAttachment)
    }

    ForEach ( $attachment in $Attachments )
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
}#End function fnSend-SMTPRelayMessage

Function fnCheck-Path {#Begin function to check path variable and return results
 	[CmdletBinding()]
    Param
    (
        [Parameter(Mandatory,Position=0)]
        [string]$Path,
        [Parameter(Mandatory,Position=1)]
        $PathType
    )
    
	#Define variables
	$VerbosePreference = "Continue" 

    Switch ( $PathType )
    {
    		File	{
				If ( ( Test-Path -Path $Path -PathType Leaf ) -eq $true )
				{
					Write-Verbose -Message "File: $Path already exists..." -Verbose
				}
				Else
				{
					New-Item -Path $Path -ItemType File -Force
					Write-Verbose -Message "File: $Path not present, creating new file..." -Verbose
				}
			}
		Folder
			{
				If ( ( Test-Path -Path $Path -PathType Container ) -eq $true )
				{
					Write-Verbose -Message "Folder: $Path already exists..." -Verbose
				}
				Else
				{
					New-Item -Path $Path -ItemType Directory -Force
					Write-Verbose -Message "Folder: $Path not present, creating new folder" -Verbose
				}
			}
	}
}#end function fnCheck-Path

#EndRegion








#Region Script
#Begin Script

$Error.Clear()
#Start Function timer, to display elapsed time for function. Uses System.Diagnostics.Stopwatch class - see here: https://msdn.microsoft.com/en-us/library/system.diagnostics.stopwatch(v=vs.110).aspx 
$stopWatch 				= [System.Diagnostics.Stopwatch]::StartNew()
$dtmScriptStartTimeUTC 	= fnUTC-Now
$transcriptFileName    	= "{0}-{1}-Transcript.txt" -f $dtmScriptStartTimeUTC.ToString("yyyy-MM-dd_HH.mm.ss"), "GroupPolicyBackups"

$myInv 		= fnGet-MyInvocation
$scriptDir  = Split-Path $MyInvocation.MyCommand.Path
$scriptName = $myInv.ScriptName

#Check to see if destination backup folder is present and accessible
[String]$gpoBkpDestFldr = "E:\GPOBackups"
fnCheck-Path -Path $gpoBkpDestFldr -PathType Folder
#Create working directory for today's backups
[String]$workingDir = "{0}\{1}" -f $scriptDir, "workingDir"
fnCheck-Path -Path $workingDir -PathType Folder
#Create .zip archive folder if needed
[String]$archiveFolder = "{0}\{1}" -f $gpoBkpDestFldr, "Archives"
fnCheck-Path -Path $archiveFolder -PathType Folder

# start transcript file
Start-Transcript ("{0}\{1}" -f  $workingDir, $transcriptFileName)
Write-Verbose -Message ("[{0} UTC] [SCRIPT] Beginning execution of script." -f $dtmScriptStartTimeUTC.ToString($dtmFormatString)) -Verbose
Write-Verbose -Message ("[{0} UTC] [SCRIPT] Script Name             :  {1}" -f $(fnUTC-Now).ToString($dtmFormatString), $scriptName) -Verbose
Write-Verbose -Message ("[{0} UTC] [SCRIPT] Script Directory path   :  {1}" -f $(fnUTC-Now).ToString($dtmFormatString), $scriptDir) -Verbose
Write-Verbose -Message ("[{0} UTC] [SCRIPT] Working Directory path  :  {1}" -f $(fnUTC-Now).ToString($dtmFormatString), $workingDir) -Verbose
Write-Verbose -Message ("[{0} UTC] [SCRIPT] Archive folder path     :  {1}" -f $(fnUTC-Now).ToString($dtmFormatString), $archiveFolder) -Verbose



#Create array of all GPO objects
$Error.Clear()
Try
{
	[Array]$GPOs = Get-GPO -All -Domain $domDNS -Server $pdcE | Select-Object DisplayName, ID
}
Catch
{
	[Array]$GPOs = Get-GPO -All -Domain $domDNS -Server $domDNS | Select-Object DisplayName, ID
}

#Loop through all GPOs in array and backup to local directory
ForEach ($GPO in $GPOs) 
{
	$gpoDisplayName = ($GPO).DisplayName
	$gpoGuid 		= ($GPO).ID
	$Comment 		= $gpoDisplayName + "_" + $(fnGet-ReportDate)
	$Error.Clear()
	Try
	{
		Backup-GPO -Guid $gpoGuid -Path $workingDir -Comment $Comment -Domain $domDNS -Server $pdcE -ErrorAction Stop -WarningAction Continue
	}
	Catch
	{
		Backup-GPO -Guid $gpoGuid -Path $workingDir -Comment $Comment -Domain $domDNS -Server $domDNS -ErrorAction Stop -WarningAction Continue
	}
	
	$GPO = $gpoDisplayName = $gpoGuid = $Comment = $null
}

Stop-Transcript
Sleep -Seconds 10

#Save output
#Compress the GPO folders into a single archive
$archiveFile 		= "{0}\{1}" -f $gpoBkpDestFldr, "$($domainName)_archive_for_$(fnGet-ReportDate).zip"
$compressionLevel 	= [System.IO.Compression.CompressionLevel]::Optimal
If( ( Test-Path -Path $archiveFile -PathType Leaf ) -eq $true ) { Remove-Item -Path $archiveFile -Confirm:$false }
#See https://msdn.microsoft.com/en-us/library/hh875104(v=vs.110).aspx?cs-save-lang=1&cs-lang=vb#code-snippet-1
[IO.Compression.ZipFile]::CreateFromDirectory($workingDir, $archiveFile, $compressionLevel, $false)


#Stop the stopwatch	
$stopWatch.Stop()

#Send e-mail notification	
$emailTemplatePath      = "{0}\{1}" -f $scriptDir, "EmailTemplates"
$imageAttachmentPath    = "{0}\{1}" -f $scriptDir, "Images"
$emailTemplateFileName 	= "IAM_Weekly_GPOBkpNotification.html"
$xmlConfigFile 			= "{0}\{1}" -f $scriptDir, "EmailSettings.xml"
[xml]$objXmlConfig 		= Get-Content $xmlConfigFile

$runTime = $stopWatch.Elapsed.ToString('dd\.hh\:mm\:ss')
Write-Verbose -Message  ("[{0}] Sending email notification, please wait..." -f $(fnUTC-Now)::Now.ToString($dtmFormatString)) -Verbose
$emailTemplate = "{0}\{1}" -f $emailTemplatePath, $emailTemplateFileName

$htmlTemplate 	= [System.IO.StreamReader]$emailTemplate
$messageBody 	= $htmlTemplate.ReadToEnd()
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
	#To = "hemiller@deloitte.com"
	To = "gtsctoiaminfrastructureteam@deloitte.com"
	CC = "dbreeze@deloitte.com"
	From = "IAM-Weekly-GPOBackupNotifications@deloitte.com"
	ReplyTo = "IAM-Weekly-GPOBackupNotifications@deloitte.com"
	SMTPServer = $objXmlConfig.Configuration.EmailSettings.SMTPServer
	Port = $objXmlConfig.Configuration.EmailSettings.Port
	InlineImageAttachments = $colInlineImageAttachments
	Attachments = $colAttachments
}



If ( ( Test-Path -Path $archiveFile -PathType Leaf ) -eq $true )
{
	$messageBody = $messageBody.Replace("@@ScriptStatus@@", "Success")
	$params.Subject = "SUCCESS: $($domainName)_GPO_Backup_Report_as_of_$(fnGet-ReportDate)"
	$params.Body = $messageBody
}
Else
{
	$messageBody = $messageBody.Replace("@@ScriptStatus@@", "Failed")
	$params.Subject = "FAILED: $($domainName)_GPO_Backup_Report_as_of_$(fnGet-ReportDate)"
	$params.Body = $messageBody
}

fnSend-SmtpRelayMessage @params
Get-ChildItem -Path $scriptDir | Where-Object { ($_.PsIsContainer) -and ($_.Name -like "workingDir*") } | Remove-Item -Recurse -Force -Confirm:$false

#Close out script
$dtmScriptStopTimeUTC = fnUTC-Now
$elapsedTime = New-TimeSpan -Start $dtmScriptStartTimeUTC -End $dtmScriptStopTimeUTC
Write-Verbose -Message ("[{0} UTC] [SCRIPT] Script Complete" -f $(fnUTC-Now).ToString($dtmFormatString)) -Verbose
Write-Verbose -Message ("[{0} UTC] [SCRIPT] Script Start Time :  {1}" -f $(fnUTC-Now).ToString($dtmFormatString), $dtmScriptStartTimeUTC.ToString($dtmFormatString)) -Verbose
Write-Verbose -Message ("[{0} UTC] [SCRIPT] Script Stop Time  :  {1}" -f $(fnUTC-Now).ToString($dtmFormatString), $dtmScriptStopTimeUTC.ToString($dtmFormatString)) -Verbose
Write-Verbose -Message ("[{0} UTC] [SCRIPT] Elapsed Time: {1:N0}.{2:N0}:{3:N0}:{4:N1}  (Days.Hours:Minutes:Seconds)" -f $(fnUTC-Now).ToString($dtmFormatString), $elapsedTime.Days, $elapsedTime.Hours, $elapsedTime.Minutes, $elapsedTime.Seconds) -Verbose


#Clean up archive files
$filesToMove = Get-ChildItem -Path $gpoBkpDestFldr | Where-Object { !$_.PSIsContainer -and $_.CreationTime -lt $moveLimit }
$filesToMove | ForEach-Object { Move-Item -Path $_.FullName -Destination $archiveFolder -Force }

Get-ChildItem -Path $archiveFolder -Recurse | Where-Object { $_.LastWriteTime -lt $Limit } | Remove-Item -Force
#EndRegion