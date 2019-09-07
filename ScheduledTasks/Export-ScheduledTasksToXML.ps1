#Region Help
<#

.NOTES
#------------------------------------------------------------------------------
#
# THIS CODE IS MADE AVAILABLE AS IS, WITHOUT WARRANTY OF ANY KIND. THE
# ENTIRE RISK OF THE USE OR THE RESULTS FROM THE USE OF THIS CODE REMAINS
# WITH THE USER.
#
#------------------------------------------------------------------------------
.SYNOPSIS
Script to export current scheduled tasks using ScheduledTasks module

.DESCRIPTION
This script exports scheduled tasks to XML files for import to another
computer

It will export all tasks in the \ ; \IAM  ; \Deloitte folders on a computer

.OUTPUTS
XML files for each scheduled task found in \ and \IAM folders. Files will
be placed in 'C:\MigrationInfo\$ComputerName\rootTasks',
'C:\MigrationInfo\$ComputerName\iamTasks'
'C:\MigrationInfo\$ComputerName\deloitteTasks'

.EXAMPLE
.\Export-ScheduledTasksToXML.ps1 -ComputerName <NameOfComputer>

###########################################################################
#
#
# AUTHOR:  Heather Miller
#          
#
# VERSION HISTORY:
# 1.0 4/4/2017 - Initial release
#
###########################################################################
#>
#EndRegion

Param
(
	[Parameter(Position=0,Mandatory=$true, ValueFromPipeline=$true)]
	[Alias("CN","Computer")]
	[String[]]$ComputerName="$env:COMPUTERNAME"
)

#Region Modules
#Check if required module is loaded, if not load import it
Try 
{
	Import-Module ScheduledTasks -ErrorAction Continue
}
Catch
{
	Throw "ScheduledTasks module could not be loaded. $($_.Exception.Message)"
}

#EndRegion

#Region Variables
$rootTasks = @()
$iamTasks = @()
$deloitteTasks = @()
$rootTaskOutput = "C:\MigrationInfo\$ComputerName\rootTasks"
$iamTaskOutput = "C:\MigrationInfo\$ComputerName\iamTasks"
$deloitteTaskOutput = "C:\MigrationInfo\$ComputerName\deloitteTasks"
#EndRegion


#Region Script
#Begin script

[Array]$rootTasks = Get-ScheduledTask -TaskName * -TaskPath "\"-ErrorAction Continue | Where-Object { $_.State -eq "Ready" }

[Array]$iamTasks = Get-ScheduledTask -TaskName * -Taskpath "\IAM\" -ErrorAction Continue | Where-Object { $_.State -eq "Ready" }

[Array]$deloitteTasks = Get-ScheduledTask -TaskName * -Taskpath "\Deloitte\" -ErrorAction Continue | Where-Object { $_.State -eq "Ready" }

If ( ( Test-Path -Path "C:\MigrationInfo\$ComputerName" -PathType Container ) -eq $false ) { New-Item -Path "C:\MigrationInfo\$ComputerName" -ItemType Directory -Force }


#Export tasks under \ folder
If ( $rootTasks.Count -gt 0 )
{
	If ( ( Test-Path -Path $rootTaskOutput -PathType Container ) -eq $false ) { New-Item -Path $rootTaskOutput -ItemType Directory -Force }
	ForEach ( $task in $rootTasks )
	{
		$taskName = ($task).TaskName
		$taskPath = ($task).TaskPath
		[String]$taskXML = Export-ScheduledTask -TaskName $taskName -TaskPath $taskPath
		$taskXML | Out-File -FilePath "$rootTaskOutput\$taskName.xml"
	}
}

#Export tasks under \IAM\ folder
If ( $iamTasks.Count -gt 0 )
{
	If ( ( Test-Path -Path $iamTaskOutput -PathType Container ) -eq $false ) { New-Item -Path $iamTaskOutput -ItemType Directory -Force }
	ForEach ( $task in $iamTasks )
	{
		$taskName = ($task).TaskName
		$taskPath = ($task).TaskPath
		[String]$taskXML = Export-ScheduledTask -TaskName $taskName -TaskPath $taskPath
		$taskXML | Out-File -FilePath "$iamTaskOutput\$taskName.xml"
	}
}

#Export tasks under \Deloitte\ folder(s)
If ( $deloitteTasks.Count -gt 0 )
{
	If ( ( Test-Path -Path $deloitteTaskOutput -PathType Container ) -eq $false ) { New-Item -Path $deloitteTaskOutput -ItemType Directory -Force }
	ForEach ( $task in $deloitteTasks )
	{
		$taskName = ($task).TaskName
		$taskPath = ($task).TaskPath
		[String]$taskXML = Export-ScheduledTask -TaskName $taskName -TaskPath $taskPath
		$taskXML | Out-File -FilePath "$deloitteTaskOutput\$taskName.xml"
	}
}

#EndRegion