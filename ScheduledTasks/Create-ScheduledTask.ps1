<#
.Synopsis
	This PowerShell Script is used to create a scheduled task on a specified machine.
.Description
	This PowerShell Script uses the Schedule.Service COM Object to create a Windows
    Scheduled Task on a specified machine. Using the parameters provided, you can 
    specify the name, hostname, program to launch and the location of the script or file you're
    opening. The person executing the script will have to provide credentials.
.Example
	C:\PS>.\Create-ScheduledTask.ps1 -Hostname myhost.fqdn.com -Description 'Run Pshell Script' -ScriptPath C:\Scripts\myscript.ps1 -SchedTaskCreds (Get-Credential)
	This example creates a scheduled task on host myhost, running a powershell script, under credentials specified by get-credential
.Notes
	Name: .\Create-ScheduledTask.ps1
	Last Change by: Ionut Nica
	Original Author: Ryan Dennis
	Last Edit: 27/10/2012
	Keywords: Create Scheduled Task
.Link
	www.rivnet.ro
#>

[CmdletBinding()]
Param(
[Parameter(Mandatory=$true)][System.String]$Description,
[Parameter(Mandatory=$true)][System.String]$HostName,
[Parameter(Mandatory=$false)][System.String]$Program="PowerShell.exe",
[Parameter(Mandatory=$true)][System.String]$ScriptPath,
[Parameter(Mandatory=$true)][System.Management.Automation.PSCredential]$SchedTaskCreds

)

# Date Variables #
$date = (Get-Date 11:00PM).AddDays(1)
$taskStartTime = $date | Get-Date -Format yyyy-MM-ddTHH:ss:ms

# Get the credentials #
$UserName = $SchedTaskCreds.UserName
$Password = $SchedTaskCreds.GetNetworkCredential().Password

# Build the Argument based on Prefix, Suffix and ScriptPath parameter #
$argPrefix = '-file "'
$argSuffix = '"'
$programArguments = $argPrefix+$ScriptPath+$argSuffix

$service = New-Object -ComObject "Schedule.Service"
$service.Connect($Hostname)
$rootFolder = $service.GetFolder("\")
$taskDefinition = $service.NewTask(0)
$regInfo = $taskDefinition.RegistrationInfo
$regInfo.Description = $Description
$regInfo.Author = [Security.Principal.WindowsIdentity]::GetCurrent().Name
$settings = $taskDefinition.Settings
$settings.Enabled = $True
#allow the task to start on demand
$settings.AllowDemandStart = $True
$settings.StartWhenAvailable = $True
$settings.Hidden = $False
$triggers = $taskDefinition.Triggers
$trigger = $triggers.Create(2)
#$startTime = "2006-05-02T22:00:00"
$trigger.StartBoundary = $taskStartTime
$trigger.DaysInterval = 1
$trigger.Id = "DailyTriggerId"
$trigger.Enabled = $True
$Action = $taskDefinition.Actions.Create(0)
$Action.Path = $Program
$Action.Arguments = $programArguments
$Principal = $taskDefinition.Principal
# Principal.RunLevel -- 0 is least privilege, 1 is highest privilege #
$Principal.RunLevel = 1


$rootFolder.RegisterTaskDefinition($Description, $taskDefinition, 6, $UserName, $Password, 1)
