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
Register scheduled tasks from exported XML files

.DESCRIPTION
This script registers/imports scheduled tasks into the specified folders from
previously exported XML task configuration files. It assumes the script will
undert NT Authority\SYSTEM.

.OUTPUTS
None

.EXAMPLE 
.\Set-ScheduledTasksfromXML.ps1 -ComputerName <NameOfComputer>

###########################################################################
#
#
# AUTHOR:  Heather Miller
#          
#
# VERSION HISTORY:
# 1.0 4/5/2017 - Initial release
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

$rootXMLTaskDir = "C:\MigrationInfo\$ComputerName\rootTasks"
$rootTasks = Get-ChildItem -Path $rootXMLTaskDir -Filter *.xml -Recurse
 
ForEach ($task in $rootTasks)
{
    $TaskName = [io.path]::GetFileNameWithoutExtension($task)
    Register-ScheduledTask -Xml (get-content $task.FullName | out-string) -TaskName $TaskName –User SYSTEM -Taskpath \
}


$iamXMLTaskDir = "C:\MigrationInfo\$ComputerName\iamTasks"
$iamTasks = Get-ChildItem -Path $iamXMLTaskDir -Filter *.xml -Recurse
 
ForEach ($task in $iamTasks)
{
    $TaskName = [io.path]::GetFileNameWithoutExtension($task)
    Register-ScheduledTask -Xml (get-content $task.FullName | out-string) -TaskName $TaskName –User SYSTEM -Taskpath IAM
}