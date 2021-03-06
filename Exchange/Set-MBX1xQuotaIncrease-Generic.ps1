#Region Help
<#
.NOTES
Segments of this script have been taken from the Mailbox Quota Powershell Script
owned by the xxx team, and have been reused for this script.

.SYNOPSIS
E-mail mailbox notifications to a users whose quota has been increased.

.DESCRIPTION 
This script will query Active Directory based on a user's employeeNumber
attribute (HR System Rank Number) and then based on pre-determined 
criteria will initiate a set of processes that will issue a one time, 
one gigabyte (1GB) increase to the user's send quota in Microsoft Exchange
for a 30 day period.

.OUTPUTS
Notification e-mail for each user whose mailbox quota has been granted a
one time increase of 1GB to their Exchange mailbox send quota

.PARAMETER -User or -u <Login ID of user>

.EXAMPLE 
.\Set-MBX1xQuotaIncrease.ps1 -User or -u <Login ID of user>

#>
###########################################################################
#
# AUTHOR:  Heather Miller
#		   
#
# VERSION HISTORY:
# 1.0 4/16/2014 - Initial Release
#
###########################################################################
#EndRegion

Param( 
[Parameter(Position=0, Mandatory=$true)] 
[string] 
[ValidateNotNullOrEmpty()] 
[alias("u")] 
$User
)#End Param

#Region ExecutionPolicy
#Set Execution Policy for Powershell
Set-ExecutionPolicy Unrestricted
#EndRegion

#Region Modules
#Import required modules
IF (-not(Get-Module ActiveDirectory))
{
	Import-Module ActiveDirectory
}
#EndRegion

#Region Variables
#Initialize Variables
$AdminAddresses = "Administrator1@smtp.mailserver.com", "Administrator2@smtp.mailserver.com"
$dnsRoot = (Get-ADDomain).DNSRoot
[int] $intPSQBumpLimit = 1024
[int] $intWQLimit = 308
[Array]$Properties = @("extensionAttribute4", "givenName", "mail", "Name", "samAccountName", "sn", "title")
$pshellCMD = $env:SystemRoot + "\System32\WindowsPowerShell\v1.0\powershell.exe"
$shellPath = $env:ExchangeInstallPath + "\bin\RemoteExchange.ps1"
$ScriptName = $MyInvocation.MyCommand.Name
$smtpServer = 'smtp.mailserver.com'
$TodaysDate = Get-Date -Format "MM-dd-yyyy"
$ToolPath = @('-PSConsoleFile','$env:ExchangeInstallPath\bin\exshell.psc1','-Command','$env:ExchangeInstallPath\bin\RemoteExchange.ps1',';Connect-ExchangeServer -Auto')
$VerbosePreference = "SilentlyContinue"
#EndRegion

#Region Functions
#Define Functions
Function Get-TodaysDate {#Begin function to get the current date
	Get-Date -Format MM-dd-yyyy
}#End function Get-TodaysDate

Function Get-LongDate {#Begin function to get long date
	Get-Date -Format G
}#End function Get-LongDate

Function fnSet-NewMBXQuota  {#Begin function fnSet-NewMBXQuota

	Param( 
	[Parameter(Position=0, Mandatory=$true)] 
	[string] 
	[ValidateNotNullOrEmpty()] 
	$oSam
	)#End Param
	
    #Dim Variables for mailbox increments
    [int] $int1GBIncrement = 1024
    [int] $intWarnIncrement = 308
    
    #Establish connection to Microsoft Exchange Server 2010
	IF ((Test-Path -Path $shellPath -PathType Leaf) -eq $true) 
	{
		Add-PSSnapin Microsoft.Exchange.Management.Powershell.E2010
		& $pshellCMD $ToolPath
		Write-Verbose "Using E2K10 Snapin and E2K10 Mgt. Shell" -Verbose
	}
	ELSEIF ((Test-Path -Path $shellPath -PathType Leaf) -eq $true) 
	{
		$ExSession = New-PSSession –ConfigurationName Microsoft.Exchange –ConnectionUri 'http://exchsrv01.$dnsRoot/PowerShell/?SerializationLevel=Full' –Authentication Kerberos
		Import-PSSession $ExSession -AllowClobber
		Write-Verbose "Connected via original import session code." -Verbose
	}
	ELSE
	{
		$exchangeServer = "exchsrv01","exchsrv02","exchsrv03" | Where-Object {Test-Connection -ComputerName $_ -Count 1 -Quiet} | Get-Random
		$ExSession = New-PSSession –ConfigurationName Microsoft.Exchange –ConnectionUri ("http://{0}.$dnsRoot/PowerShell" -f $exchangeServer)
		Import-PSSession $ExSession -AllowClobber
		Write-Verbose "Connected via imported remote session" -Verbose
	}
	#Get user's mailbox
	$uMailbox = Get-Mailbox -Identity $oSam -ErrorAction SilentlyContinue
    $uDisplayName = ($uMailbox).DisplayName
    $uMailDB = ($uMailbox).Database
	$uOffice = ($uMailbox).Office
	$uName = ($uMailbox).Name
	$uSAM = ($uMailbox).samAccountName
    $uUseDefaultDBQuotas = ($uMailbox).UseDatabaseQuotaDefaults
    IF ($uUseDefaultDBQuotas) {
		#Turn off mailbox database default quota limits
		Set-Mailbox -Identity $uSam -UseDatabaseQuotaDefaults $false
		Write-Verbose "User Mailbox is: $uMailbox" -Verbose
		#Get user's mailbox and mailbox DN
		$uMailboxDB = Get-MailboxDatabase -Identity $uMailDB
        #Calculate new user send/receive/warning limits based on DAG number
		IF ((Get-MailboxDatabase -Identity $uMailDB).ProhibitSendQuota.Value) {
			$uPSQ = (Get-MailboxDatabase -Identity $uMailDB).ProhibitSendQuota.Value.ToMB()
			Write-Verbose "Mailbox database associated with $uName has a ProhibitSendQuota value of $uPSQ." -Verbose
		}
		IF ((Get-MailboxDatabase -Identity $uMailDB).ProhibitSendReceiveQuota.Value) {
			$uPSRQ = (Get-MailboxDatabase -Identity $uMailDB).ProhibitSendReceiveQuota.Value.ToMB()
			Write-Verbose "Mailbox database associated with $uName has a ProhibitSendReceiveQuota value of $uPSRQ." -Verbose
		}
		IF ((Get-MailboxDatabase -Identity $uMailDB).IssueWarningQuota.Value) {
			$uIWQ = (Get-MailboxDatabase -Identity $uMailDB).IssueWarningQuota.Value.ToMB()
			Write-Verbose "Mailbox database associated with $uName has a IssueWarningQuota value of $uIWQ." -Verbose
		}
		$newPSQ = $uPSQ + $intPSQBumpLimit
		Write-Verbose "New send quota will be: $newPSQ" -Verbose
		$newPSRQ = $newPSQ * 2
		Write-Verbose "New send/receive quota is: $newPSRQ" -Verbose
		$newIWQ = $newPSQ - $intWQLimit
		Write-Verbose "New warning quota is: $newIWQ" -Verbose
		$strNewPSQ = [String] $newPSQ + "MB"
		Write-Verbose "New send quota string value is: $strNewPSQ" -Verbose
		$strNewPSRQ = [String] $newPSRQ + "MB"
		Write-Verbose "New send-receive qSet-uota string value is: $strNewPSRQ"
		$strNewIWQ = [String] $newIWQ + "MB"
		Write-Verbose "New warning quota string value is: $strNewIWQ" -Verbose
    }
	ELSE
	{
		IF ((Get-Mailbox -Identity $uSAM -ResultSize Unlimited ).ProhibitSendQuota.Value) {
			$uPSQ = (Get-Mailbox $uSAM -ResultSize Unlimited ).ProhibitSendQuota.Value.ToMB()
			Write-Verbose "Current mailbox send quota is: $uPSQ" -Verbose
		}
		IF ((Get-Mailbox -Identity $uSAM -ResultSize Unlimited ).ProhibitSendReceiveQuota.Value) {
			$uPSRQ = (Get-Mailbox $uSAM -ResultSize Unlimited ).ProhibitSendReceiveQuota.Value.ToMB()
			Write-Verbose "Current mailbox send receive quota is: $uPSRQ" -Verbose
		}
		IF ((Get-Mailbox -Identity $uSAM -ResultSize Unlimited ).IssueWarningQuota.Value) {
			$uIWQ = (Get-Mailbox $uSAM -ResultSize Unlimited ).IssueWarningQuota.Value.ToMB()
			Write-Verbose "Current mailbox warning quota is: $uIWQ" -Verbose
		}
		$newPSQ = $uPSQ + $intPSQBumpLimit
		Write-Verbose "New send quota will be: $newPSQ" -Verbose
		$newPSRQ = $newPSQ * 2
		Write-Verbose "New send/receive quota is: $newPSRQ" -Verbose
		$newIWQ = $newPSQ - $intWQLimit
		Write-Verbose "New warning quota is: $newIWQ" -Verbose
		$strNewPSQ = [String] $newPSQ + "MB"
		Write-Verbose "New send quota string value is: $strNewPSQ" -Verbose
		$strNewPSRQ = [String] $newPSRQ + "MB"
		Write-Verbose "New send-receive quota string value is: $strNewPSRQ"
		$strNewIWQ = [String] $newIWQ + "MB"
		Write-Verbose "New warning quota string value is: $strNewIWQ" -Verbose
	}#End UseDefaultDatabaseQuotas check
	Set-Mailbox -Identity $uMailbox -ProhibitSendQuota $strNewPSQ -ProhibitSendReceiveQuota $strNewPSRQ -IssueWarningQuota $strNewIWQ -CustomAttribute4 $(Get-TodaysDate)
	$uMBXObject | Add-Member -MemberType NoteProperty -Name Mailbox -Value $uMailbox
	$uMBXObject | Add-Member -MemberType NoteProperty -Name ProhibitSendQuota -Value $strNewPSQ
	$uMBXObject | Add-Member -MemberType NoteProperty -Name ProhibitSendReceiveQuota -Value $strNewPSRQ
	$uMBXObject | Add-Member -MemberType NoteProperty -Name IssueWarningQuota -Value $strNewIWQ
}#End Function fnSetNewMBXQuota
#EndRegion


#Region Script
#Begin Script

#Build PSObject for reporting
$uMBXObject = New-Object PSObject
$uMBXObject | Add-Member -MemberType NoteProperty -Name ScriptStartDate -Value Get-LongDate

$ADUser  = Get-ADUser -Filter {samAccountName -eq $User} -Properties $Properties | Select-Object -Property $Properties
Write-Verbose "User is: $ADUser" -Verbose

IF ($ADUser) 
{
    $oXA4 = ($ADUser).extensionAttribute4
    $oEMail = ($ADUser).mail
    $oName = ($ADUser).Name
    $oSam = ($ADUser).samAccountName
    $oTitle = ($ADUser).title
	$uMBXObject | Add-Member -MemberType NoteProperty -Name ValidUser -Value "$oName exists within Active Directory."
	IF ($oTitle -and $oXA4 -eq $null) 
	{
		$uMBXObject | Add-Member -MemberType NoteProperty -Name Eligibility -Value "User $oName is eligible for mailbox quota increase."
		fnSet-NewMBXQuota $oSam
    }
	ELSE
	{
		$uMBXObject | Add-Member -MemberType NoteProperty -Name Eligibility -Value "User $oName is ineligible for mailbox quota increase."
	}
} 
ELSE 
{
	$uMBXObject | Add-Member -MemberType NoteProperty -Name ValidUser -Value "$oName does not exist within Active Directory. Please verify user ID."
}
$uMBXObject | Add-Member -MemberType NoteProperty -Name ScriptEndDate -Value Get-LongDate
Write-Output -InputObject $uMBXObject -ErrorAction Continue
#EndRegion