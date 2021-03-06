#Requires -Modules ADCSTemplate, PKI -Version 3.0

#Region Help
<#

.NOTES Needs to be run from Certificate Authority server

.SYNOPSIS Backs up Certificate Authority Templates

.DESCRIPTION Exports all CA certificate templates to a serialized 
json file and zips them to a .zip file

.PARAMETER BackupFolder

.OUTPUTS .Zip file containing all PKI templates

.EXAMPLE 
.\Export-AllCertTemplates.ps1 -BackupFolder


#>
###########################################################################
#
#
# AUTHOR:  Heather Miller

# Credit to: Ashley McGlone
# @GoateePFE
#
# VERSION HISTORY:
# 10/15/2019 - Version 1.0
# 
###########################################################################
#EndRegion

[CmdletBinding()]
param
(
	[Parameter(Mandatory = $true, HelpMessage = "Specify full path where CA backup should should reside. EG: E:\TemplateBackups")]
	[alias("Folder")]
	[String]$BackupFolder
	
	
)

#Region ExecutionPolicy
#Set Execution Policy for Powershell
Set-ExecutionPolicy RemoteSigned
#EndRegion

#Region Modules
#Check IF required module is loaded, IF not load import it
If (-not (Get-Module ADCSAdministration))
{
	Import-Module ADCSAdministration
}
If (-not (Get-Module PKI))
{
	Import-Module -Name PKI
}
If (-not (Get-Module ADCSTemplate))
{
	Import-Module -Name ADCSTemplate
}
If (-not (Get-Module Microsoft.PowerShell.Security))
{
	Import-Module Microsoft.PowerShell.Security
}
#EndRegion

#Region Variables
$TempBkpRoot = "{0}\{1}" -f $BackupFolder, "Templates"
$myScriptName = $MyInvocation.MyCommand.Name
Add-Type -AssemblyName "System.IO.Compression.FileSystem"
$compressionLevel = [System.IO.Compression.CompressionLevel]::Optimal
$VerbosePreference = 'Continue'
#EndRegion

#Region Functions
Function Get-UtcTime
{
	#Begin function to get UTC date and time
	[System.DateTime]::UtcNow
} #End Get-UtcTime

Function Get-ReportDate {#Begin function set report date format
	Get-Date -Format "yyyy-MM-dd"
}#End function Get-ReportDate

Function Get-MyInvocation
{
	#Begin function to define $MyInvocation
	Return $MyInvocation
} #End function Get-MyInvocation


#EndRegion








#Region Script
#Begin Script
$Error.Clear()

#Start Function timer, to display elapsed time for function. Uses System.Diagnostics.Stopwatch class - see here: https://msdn.microsoft.com/en-us/library/system.diagnostics.stopwatch(v=vs.110).aspx 
$stopWatch = [System.Diagnostics.Stopwatch]::StartNew()
$dtmScriptStartTimeUTC = Get-UtcTime
$dtmFormatString = "yyyy-MM-dd HH:mm:ss"
$transcriptFileName = "{0}-{1}-Transcript.txt" -f $dtmScriptStartTimeUTC.ToString("yyyy-MM-dd_HH.mm.ss"), "$($thisServer)-CATemplateBackup"

$myInv = Get-MyInvocation
$scriptDir = $myInv.PSScriptRoot
$scriptName = $myInv.ScriptName

#Check pre-requisite folders exist, if not, create them

if ((Test-Path -Path $BackupFolder -PathType Container) -eq $False)
{
	New-Item -ItemType Directory -Path $BackupFolder -Force
}

if ((Test-Path -Path $TempBkpRoot -PathType Container) -eq $False)
{
	New-Item -ItemType Directory -Path $TempBkpRoot -Force
}

# start transcript file
Start-Transcript ("{0}\{1}" -f $BackupFolder, $transcriptFileName)
Write-Verbose -Message ("[{0} UTC] [SCRIPT] Beginning execution of script." -f $dtmScriptStartTimeUTC.ToString($dtmFormatString))
Write-Verbose -Message ("[{0} UTC] [SCRIPT] Script Name             		:  {1}" -f $(Get-UtcTime).ToString($dtmFormatString), $scriptName)
Write-Verbose -Message ("[{0} UTC] [SCRIPT] Script Directory path   		:  {1}" -f $(Get-UtcTime).ToString($dtmFormatString), $scriptDir)
Write-Verbose -Message ("[{0} UTC] [SCRIPT] Main Backup Folder path  		:  {1}" -f $(Get-UtcTime).ToString($dtmFormatString), $BackupFolder)
Write-Verbose -Message ("[{0} UTC] [SCRIPT] Template Backup Folder path    :  {1}" -f $(Get-UtcTime).ToString($dtmFormatString), $TempBkpRoot)


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
		Write-Verbose -Message "The template $($Template.Name) had a ""/"" in the name. It has been removed in the output file due to limitations of the filesystem." -Verbose
		Write-Verbose -Message "If using the templates exported from this script at a later date for import, the ""/"" will need to be added back to the file name." -Verbose
	}
	ElseIf ($templateName -match '\\')
	{
		$templateName = $templateName -replace "\\", ""
		Write-Verbose -Message "The template $($Template.Name) had a ""\"" in the name. It has been removed in the output file due to limitations of the filesystem." -Verbose
		Write-Verbose -Message "If using the templates exported from this script at a later date for import, the ""\"" will need to be added back to the file name." -Verbose
	}
	
	Write-Verbose -Message "Exporting $tempDisplayName..." -Verbose
	$templateFileName = "{0}.{1}" -f $templateName, "json"
	$templateFile = "{0}\{1}" -f $TempBkpRoot, $templateFileName
	
	Export-ADCSTemplate -DisplayName $tempDisplayName | Out-File -FilePath $templateFile -Force
}

#Create .Zip backup to save space
$templatesArchiveFile = "{0}\{1}" -f $BackupFolder, "CertTemplates_$(Get-ReportDate).zip"
if ((Test-Path -Path $templatesArchiveFile -PathType Leaf) -eq $true) { Remove-Item -Path $templatesArchiveFile -Confirm:$false }
#See https://msdn.microsoft.com/en-us/library/hh875104(v=vs.110).aspx?cs-save-lang=1&cs-lang=vb#code-snippet-1
[IO.Compression.ZipFile]::CreateFromDirectory($TempBkpRoot, $templatesArchiveFile, $compressionLevel, $false)
	
if ((Test-Path -Path $templatesArchiveFile -PathType Leaf) -eq $true)
{
	Get-ChildItem -Path $BackupFolder | Where-Object { ($_.PsIsContainer) -and ($_.Name -like "Templates*") } | Remove-Item -Recurse -Force -Confirm:$false
}


#Stop the stopwatch	
$stopWatch.Stop()

$dtmScriptStopTimeUTC = Get-UtcTime
$elapsedTime = New-TimeSpan -Start $dtmScriptStartTimeUTC -End $dtmScriptStopTimeUTC
$runtime = $stopWatch.Elapsed.ToString('dd\.hh\:mm\:ss')

Write-Verbose -Message ("[{0} UTC] [SCRIPT] Script Complete" -f $(Get-UtcTime).ToString($dtmFormatString))
Write-Verbose -Message ("[{0} UTC] [SCRIPT] Script Stop Time  :  {1}" -f $(Get-UtcTime).ToString($dtmFormatString), $dtmScriptStopTimeUTC.ToString($dtmFormatString))
Write-Verbose -Message ("[{0} UTC] [SCRIPT] Elapsed Time: {1:N0}.{2:N0}:{3:N0}:{4:N1}  (Days.Hours:Minutes:Seconds)" -f $(Get-UtcTime).ToString($dtmFormatString), $elapsedTime.Days, $elapsedTime.Hours, $elapsedTime.Minutes, $elapsedTime.Seconds)

#EndRegion