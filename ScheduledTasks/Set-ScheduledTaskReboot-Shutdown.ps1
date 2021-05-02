<#

.NOTES
THIS CODE IS MADE AVAILABLE AS IS, WITHOUT WARRANTY OF ANY KIND. THE
RISK OF THE USE OR THE RESULTS FROM THE USE OF THIS CODE REMAINS
THE USER.

.SYNOPSIS
	Schedule a computer reboot or shutdown via Powershell script
.DESCRIPTION
	This script consumes 3 parameters that allow the person executing the script to
	specify:
		1) Date and time the scheduled task should execute
		2) Length of time the computer should remain in maintenance mode in SCOM
		3) Whether to use 'Reboot' parameters or 'Shutdown' parameters

.PARAMETER TriggerTime
	Date and Time task execution should occur in format 'yyyy,MM,dd,HH,mm,ss'
	
.PARAMETER OutageDuration
	Length of time, in hours, computer should remain in maintenance mode in SCOM
	
.PARAMETER R
	Switch to tell script to use reboot parameters
	
.PARAMETER S
	Switch to tell script to use shutdown parameters
	
.OUTPUTS
	Console output of new scheduled task path, task name and status

.EXAMPLE 
	PS>.\Set-ScheduledTaskReboot-Shutdown.ps1 -TriggerTime '2019,05,21,11,30,00' -OutageDuration 1 -R

.EXAMPLE 
	PS>.\Set-ScheduledTaskReboot-Shutdown.ps1 -TriggerTime '2019,05,21,11,30,00' -OutageDuration 1 -S
#>

###########################################################################
#
#
# AUTHOR:  
#	Heather Miller, Manager, Identity and Access Management
#
#
# VERSION HISTORY:
# 	1.0 5/29/2019 - Initial release
#
# 
###########################################################################

[CmdletBinding()]
Param (
	[Parameter(Mandatory = $true, Position = 0, ValueFromPipeline = $true, HelpMessage = "Specify the time the task should execute. Format - ""yyyy,MM,dd,HH,mm,ss""")]
	$TriggerTime,
	[Parameter(Mandatory = $true, Position = 1, ValueFromPipeline = $true, HelpMessage = "Specify the length of the outage in hours.")]
	[String]$OutageDuration,
	[Parameter(ParameterSetName = 'Reboot', Mandatory = $true, Position = 2, ValueFromPipeline = $true, HelpMessage = "Use this parameter to reboot the computer.")]
	[Switch]$R,
	[Parameter(ParameterSetName = 'Shutdown', Mandatory = $true, Position = 3, ValueFromPipeline = $true, HelpMessage = "Use this parameter to shutdown the computer.")]
	[Switch]$S
)





#Begin script
# note: trigger time is formatted as - System.DateTime(year, month, day, hours, minutes, seconds)
$TriggerTime = [DateTime]::ParseExact($TriggerTime, "yyyy,MM,dd,HH,mm,ss", $null)
# note: uncomment the other trigger time if setting up the scheduled task across multiple systems to happen at the same UTC time - time is auto converted to local system time
#$TriggerTime = $TriggerTime.ToLocalTime()


$params1 = @{
	TriggerTime	     = $TriggerTime
	OutageDurationHours = $OutageDuration
	RunAsUser		     = "SYSTEM"
}

$params2 = @{ }

If ($PSBoundParameters.ContainsKey('R'))
{
	$params1.ScheduledTaskName = "Scheduled Reboot"
	$params1.Args = '-NoProfile -NonInteractive -Windowstyle Hidden -ExecutionPolicy RemoteSigned -Command "& {Restart-Computer -Force}"'
	$params2.ScheduledTaskName = "Scheduled Reboot - Get DC Health"
}
ElseIf ($PSBoundParameters.ContainsKey('S'))
{
	$params1.ScheduledTaskName = "Scheduled Shutdown"
	$params1.Args = '-NoProfile -NonInteractive -Windowstyle Hidden -ExecutionPolicy RemoteSigned -Command "& {Stop-Computer -Force}"'
	$params2.ScheduledTaskName = "Scheduled Shutdown - Get DC Health"
}

$colActions1 = New-Object System.Collections.ArrayList
$params1ScheduledTaskAction = @{
	Execute  = "EventCreate.exe"
	Argument = ('/L Application /SO MMSchedule /T Information /ID {0} /d "{1}"' -f $params1.OutageDurationHours, $params1.ScheduledTaskName)
}
$action = New-ScheduledTaskAction @params1ScheduledTaskAction
[void]$colActions1.Add($action)

$params1ScheduledTaskAction = @{
	Execute  = "powershell.exe"
	Argument = $params1.Args
}
$action = New-ScheduledTaskAction @params1ScheduledTaskAction
[void]$colActions1.Add($action)

$colActions2 = New-Object System.Collections.ArrayList
$params2ScheduledTaskAction = @{
	Execute  = "powershell.exe"
	Argument = '-NoProfile -ExecutionPolicy RemoteSigned -File "E:\Scripts\Get-DCHealth.ps1"'
}
$action = New-ScheduledTaskAction @params2ScheduledTaskAction
[void]$colActions2.Add($action)


$principal = New-ScheduledTaskPrincipal -UserId $params1.RunAsUser -LogonType Password -RunLevel Highest
$trigger1 = New-ScheduledTaskTrigger -Once -At $params1.TriggerTime
$trigger2 = New-ScheduledTaskTrigger -AtStartup

$paramsScheduledTaskSettings = @{
	Compatibility		       = "Win7"
	DisallowDemandStart	       = $false
	Disable			       = $false
	AllowStartIfOnBatteries    = $false
	DontStopIfGoingOnBatteries = $true
	DontStopOnIdleEnd	       = $true
	StartWhenAvailable	       = $true
	
}

$settingsSet = New-ScheduledTaskSettingsSet @paramsScheduledTaskSettings


$task1 = Register-ScheduledTask $params1.ScheduledTaskName -Action $colActions1 -Principal $principal -Trigger $trigger1 -Settings $settingsSet
$task1.Triggers[0].EndBoundary = $triggerTime.ToLocalTime().AddDays(1).ToString("yyyy-MM-ddTHH:mm:ss")
$task1.Settings.DeleteExpiredTaskAfter = "P10D"
$task1.Settings.ExecutionTimeLimit = "PT1H"
$task1.Author = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name
$task1.Date = ([DateTime]::UtcNow).ToLocalTime().ToString("yyyy-MM-ddThh:mm:ss")

Set-ScheduledTask -InputObject $task1


$task2 = Register-ScheduledTask $params2.ScheduledTaskName -Action $colActions2 -Principal $principal -Trigger $trigger2 -Settings $settingsSet
$task2.Triggers[0].EndBoundary = $triggerTime.ToLocalTime().AddDays(1).ToString("yyyy-MM-ddTHH:mm:ss")
$task2.Triggers[0].Delay = "PT15M"
$task2.Settings.DeleteExpiredTaskAfter = "P10D"
$task2.Settings.ExecutionTimeLimit = "PT1H"
$task2.Author = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name
$task2.Date = ([DateTime]::UtcNow).ToLocalTime().ToString("yyyy-MM-ddThh:mm:ss")

Set-ScheduledTask -InputObject $task2
#End script