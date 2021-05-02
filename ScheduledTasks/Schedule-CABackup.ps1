#Requires -RunAsAdministrator
<#

Create a scheduled task to update PowerShell help.

Chris Warwick, @cjwarwickps, 2014
chrisjwarwick.wordpress.com
#>

<#
.Synopsis
   Creates a scheduled task to regularly update PowerShell help.
.DESCRIPTION
   This script creates a scheduled task to regularly update PowerShell help.
.EXAMPLE
   Schedule-UpdateHelp
.OUTPUTS
   Returns the ScheduledTask object that's created
#>




$ActionParams = @{
	Id		       = 'IAM'
	Execute	       = 'PowerShell.exe'
	Argument	       = '-nologo -noprofile -noninteractive -windowstyle hidden -file "E:\scripts\Scheduled Tasks\CABackupTask\Backup-DeloitteCA.ps1"'
	WorkingDirectory = 'E:\'
}

# Change the scheduling below to suit...
$msg = "Enter the username and password that will run the task";
$credential = $Host.UI.PromptForCredential("Task username and password", $msg, "$env:userdomain\$env:username", $env:userdomain)
$username = $credential.UserName
$password = $credential.GetNetworkCredential().Password
$duration = ([Timespan]::MaxValue)
$trigger = New-ScheduledTaskTrigger -Daily -At 3am
$taskName = 'IAM.Semi-Daily.CA.Backup'
$TaskParams = @{
	Taskname = $taskName
	Action   = New-ScheduledTaskAction @ActionParams
	Trigger  = $trigger
	Setting  = New-ScheduledTaskSettingsSet -RunOnlyIfNetworkAvailable -DisallowHardTerminate:$false -StartWhenAvailable -DisallowDemandStart:$false -ExecutionTimeLimit (New-TimeSpan -Hours 11)
	Description = "Scheduled task runs every 12 hours to backup CA database"
	TaskPath = "IAM"
}

Register-ScheduledTask @TaskParams -User $username -Password $password -RunLevel Highest
$task = Get-ScheduledTask -TaskName $taskName
$task.triggers.Repetition.Interval = "PT12H"
$task | Set-ScheduledTask -User $username -Password $password