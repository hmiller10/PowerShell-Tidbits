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

#Requires -RunAsAdministrator
# (Update-Help must be run as admin - so the scheduled task needs to be created from an RunAsAdmin session)


$ActionParams = @{
    Id       = 'IAM'
    Execute  = 'PowerShell.exe'
    Argument = '-nologo -noprofile -noninteractive -windowstyle hidden -file "E:\scripts\Scheduled Tasks\CABackupTask\Backup-DeloitteCA.ps1"'
	WorkingDirectory = 'E:\'
}

# Change the scheduling below to suit...
#$userID = 'Atrame\dtt_sc_pkiservice'
#[String]$pw = '5(qKhccf>4tP6~cXr72t'
#Password = ConvertTo-SecureString -String $pw -AsPlainText -Force
$msg = "Enter the username and password that will run the task"; 
$credential = $Host.UI.PromptForCredential("Task username and password",$msg,"$env:userdomain\$env:username",$env:userdomain)
$username = $credential.UserName
$password = $credential.GetNetworkCredential().Password
$duration = ([Timespan]::MaxValue)
#$interval = ([Timespan]"00:12:00:00")
$trigger = New-ScheduledTaskTrigger -Daily -At 3am
$taskName = 'IAM.Semi-Daily.CA.Backup'
$TaskParams = @{
	Taskname    = $taskName
    Action      = New-ScheduledTaskAction @ActionParams
    Trigger     = $trigger
    Setting     = New-ScheduledTaskSettingsSet -RunOnlyIfNetworkAvailable -DisallowHardTerminate:$false -StartWhenAvailable -DisallowDemandStart:$false -ExecutionTimeLimit (New-TimeSpan -Hours 11)
	#Principal   = New-ScheduledTaskPrincipal "NT AUTHORITY\SYSTEM" -LogonType ServiceAccount -RunLevel Highest
    #Principal   = New-ScheduledTaskPrincipal -UserId $userID -LogonType Password -RunLevel Highest
    Description = "Scheduled task runs every 12 hours to backup CA database"
    TaskPath     = "IAM"
}

Register-ScheduledTask @TaskParams -User $username -Password $password -RunLevel Highest
$task = Get-ScheduledTask -TaskName $taskName
$task.triggers.Repetition.Interval = "PT12H"
#$task.triggers.Repetition.Duration = $duration
$task | Set-ScheduledTask -User $username -Password $password