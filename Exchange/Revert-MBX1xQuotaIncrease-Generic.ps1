#Region Help
<#
.NOTES
Segments of this script have been taken from the Mailbox Quota Powershell Script
owned by the AD/Messaging team, and have been reused for this script.

.SYNOPSIS
E-mail mailbox notifications to a user whose quota has been reverted.

.DESCRIPTION 
This script will query Active Directory for ExtensionAttribute4 and will
examine whether the attribute is null or has a date stamped into the
attribute field.  If a date has been stamped it will be read into a variable
and will then be calculated to see if the date is 30 days or greater past
the date in ExtentionAttribute4. If the date is greater than 30 days the 
user's mailbox quota settings will be decremented by 1GB and then the other
quotas calculated accordingly.  If a revert does occur, the date the user's
mailbox was reverted will be stamped into Microsoft Exchange
ExtensionCustomAttribute4. If ExtensionCustomAttribute4 is not null the 
script will continue processing the mailboxes for additional accounts that
should be reverted.

.OUTPUTS
Notification e-mail for each user whose mailbox quota had been granted a
one time increase of 1GB to their Exchange mailbox send quota whose quota
increase has been reverted

.EXAMPLE 
.\Revert-MBX1xQuotaIncrease.ps1

#>
###########################################################################
#
# AUTHOR:  Heather Miller
#		   
#
# VERSION HISTORY:
# 7.0 4/23/2014
#
###########################################################################
#EndRegion

#Region ExecutionPolicy
#Set Execution Policy for Powershell
Set-ExecutionPolicy Unrestricted
#EndRegion

#Region Modules
#Import Required Modules
IF (-not(Get-Module ActiveDirectory))
{
	Import-Module ActiveDirectory
}
#EndRegion

#Region Variables
#Initialize Variables
[int] $intPSQBumpLimit = 1024
[int] $intWQLimit = 308
$smtpServer = 'smtp.mailserver.com'
$VerbosePreference = "SilentlyContinue"
#EndRegion

#Region Functions
#Define Functions

Function Get-TodaysDate {#Begin function to get date in common format
	Get-Date -Format MM-dd-yyyy
}#End function Get-TodaysDate

Function fnGet-LabExchangeConnection {#Begin function to connect to MS Exchange in lab
	#Dim Variables
	$dnsRoot = (Get-ADDomain).DNSRoot
	$shellPath = $env:ExchangeInstallPath + "\bin\RemoteExchange.ps1"
	$ToolPath = @('-PSConsoleFile','$env:ExchangeInstallPath\bin\exshell.psc1','-Command','$env:ExchangeInstallPath\bin\RemoteExchange.ps1',';Connect-ExchangeServer -Auto')
	$pshellCMD = $env:SystemRoot + "\System32\WindowsPowerShell\v1.0\powershell.exe"
	
	#Establish connection to Microsoft Exchange Server 2010
	IF ((Test-Path -Path $shellPath -PathType Leaf) -eq $true) 
	{
		Add-PSSnapin Microsoft.Exchange.Management.Powershell.E2010
		& $pshellCMD $ToolPath
		Write-Verbose  -Message "Using E2K10 Snapin and E2K10 Mgt. Shell" -Verbose
	}
	ELSEIF ((Test-Path -Path $shellPath -PathType Leaf) -eq $true) 
	{
		$ExSession = New-PSSession –ConfigurationName Microsoft.Exchange –ConnectionUri 'http://exchsrv01.$dnsRoot/PowerShell/?SerializationLevel=Full' –Authentication Kerberos
		Import-PSSession $ExSession -AllowClobber
		Write-Verbose  -Message "Connected via original import session code." -Verbose
	}
	ELSE
	{
		$exchangeServer = "exchsrv01","exchsrv02","exchsrv03" | Where-Object {Test-Connection -ComputerName $_ -Count 1 -Quiet} | Get-Random
		$ExSession = New-PSSession –ConfigurationName Microsoft.Exchange –ConnectionUri ("http://{0}.$dnsRoot/PowerShell" -f $exchangeServer)
		Import-PSSession $ExSession -AllowClobber
		Write-Verbose  -Message "Connected via imported remote session" -Verbose
	}
}#End function fnGet-LabExchangeConnection

Function fnGet-RevertDate {#Begin function to get date quota was reverted to place value into table row
	Param ($uMBX)
	Get-Mailbox -Identity $uMBX | Select-Object -Property @{Name='ExtensionCustomAttribute4';Expression={[string]::Join(";",($_.extensionCustomAttribute4))}} | `
	ForEach-Object {
		$RevertTimeStamp = $_.extensionCustomAttribute4
		Write-Verbose "Date quota was reverted is: $RevertTimeStamp" -Verbose
		Return $RevertTimeStamp
	}
}#End fnGet-RevertDate

Function fnGet-TodaysDate {#Begin function to set report date
	Get-Date -Format "yyyy-MM-dd"
}#End function fnGet-TodaysDate

Function fnSend-UserNotification {#Begin Function to send e-mail notitificaion to user telling them their quota was reverted
	Param ($uFirst, $uLast, $uMail)
	$GreetingName = $uFirst + " " + $uLast
	$Subject = "IMPORTANT: Your temporary e-mail quota increase has expired."
	$Body = @"
		<p>$(Get-TodaysDate)</p>
		
		<p>Dear $GreetingName,</p>

		<p>Your one time e-mail quota limits were temporarily increased thirty (30) days ago.  This 30 day extension has expired.</p>

		<p>As of today your quota limits have been reset to their original limits.</p>
		
		<p>If you have any questions, please contact your local IT HelpDesk for assistance.</p>

		<p>Thank you.</p>

		<p>*** This is an automatically generated email. Please do not reply. ***</p>
"@
	Send-MailMessage -To $uMail -From 'no-reply@domain.com' -Subject $Subject -Body $Body -BodyAsHTML -SmtpServer $smtpServer -Priority High
}#End function fnSend-UserNotification

Function fnSet-Limit {#Begin function to set number of days to calculate revert date
	(Get-Date).AddDays(-30)
}#End function fnSet-Limit
#EndRegion



#Region Script
#Begin Script
#Run function to connect to Microsoft Exchange
Write-Verbose  -Message "Checking for Microsoft.Exchange.Management.PowerShell.E2010 PSSnapin" -Verbose
IF ((Get-PSSnapin -Name Microsoft.Exchange.Management.PowerShell.E2010 -ErrorAction SilentlyContinue) -eq $NULL)
{
	Write-Verbose  -Message "Microsoft.Exchange.Management.PowerShell.E2010 not installed, intalling now" -Verbose
	#Add-PSSnapin -Name Microsoft.Exchange.Management.PowerShell.E2010
	fnGet-LabExchangeConnection
}

#Pre-configure Administrative report notification parameters 
$AdminMsg = New-Object Net.Mail.MailMessage
$smtp = New-Object Net.Mail.SmtpClient($smtpServer)
$AdminMsg.To.Add("Administrator1@smtp.mailserver.com")
$AdminMsg.Cc.Add("team@smtp.mailservercom")
$AdminMsg.From = "MailboxQuotaChecker@domain.com"
$AdminMsg.Subject = "Mailbox Quotas Reverted on $(fnGet-TodaysDate)"
$AdminMsg.Body = @()

#Re-importing AD module into remote session
Import-Module ActiveDirectory
[Array]$Properties = @("distinguishedName", "extensionAttribute4", "givenName", "mail", "Name", "samAccountName", "sn")
$RootDSE = Get-ADRootDSE
$Root = ($RootDSE).defaultNamingContext
$Users = Get-ADUser -Filter {Enabled -eq $true} -Properties $Properties -Searchbase $Root -SearchScope Subtree -ResultSetSize $null | `
Where-Object {$_.distinguishedName -like "*OU=External Accounts*" -or $_.distinguishedName -like "*OU=Users*" `
-or $_.distinguishedName -like "*CN=Users*" -and $_.extensionAttribute4 -ne $null}

#Prepare Admin E-mail HTML Header
$htmlBody += "<!DOCTYPE html>
		<HTML><TITLE>Company Quota Reset Report</TITLE>
		<BODY background-color:white>
		<H2 style=""font-family:arial;color:white;font-size:10px""> Company Exchange 2010 Mailbox Quotas Reset on $(Get-TodaysDate) </H2>
		<Table border=1 cellpadding=0 cellspacing=0>
		<TR bgcolor=navy align=center>
		<font face=""Arial"">
		<TD><B>Employee Name</B></TD>
		<TD><B>Office</B></TD>
		<TD><B>Date Quota Increased</B></TD>
		<TD><B>Old Prohibit Send Quota</B></TD>
		<TD><B>Old Prohibit Send Receive Quota</B></TD>
		<TD><B>Old Issue Warning Quota</B></TD>
		<TD><B>Date Quotas Reverted</B></TD>
		<TD><B>New Prohibit Send Quota</B></TD>
		<TD><B>New Prohibit Send Receive Quota</B></TD>
		<TD><B>New Issue Warning Quota</B></TD></TR></font>"

#Build Admin Report Table
$table = New-Object system.Data.DataTable "Company Mailbox Quota Report"
$col1 = New-Object system.Data.DataColumn Name,([string])
$col2 = New-Object system.Data.DataColumn Office,([string])
$col3 = New-Object system.Data.DataColumn DateQuotaIncreased,([string])
$col4 = New-Object system.Data.DataColumn OldProhibitSendQuotaMB,([string])
$col5 = New-Object system.Data.DataColumn OldProhibitSendReceiveQuotaMB,([string])
$col6 = New-Object system.Data.DataColumn OldIssueWarningQuotaMB,([string])
$col7 = New-Object system.Data.DataColumn DateQuotaReset,([string])
$col8 = New-Object system.Data.DataColumn NewProhibitSendQuotaMB,([string])
$col9 = New-Object system.Data.DataColumn NewProhibitSendReceiveQuotaMB,([string])
$col10 = New-Object system.Data.DataColumn NewIssueWarningQuotaMB,([string])

$table.columns.add($col1)
$table.columns.add($col2)
$table.columns.add($col3)
$table.columns.add($col4)
$table.columns.add($col5)
$table.columns.add($col6)
$table.columns.add($col7)
$table.columns.add($col8)
$table.columns.add($col9)
$table.columns.add($col10)

$iCounter = $null
ForEach ($User in $Users) 
{
	Write-Verbose "Checking user account to verify if mailbox quotas have already been reverted. Stand-By...." -Verbose
	$uFirst = ($User).givenName
	$uLast = ($User).sn
	$uMail = ($User).mail
	$uName = ($User).Name
	$uSAM = ($User).samAccountName
	Write-Verbose "User is $uName" -Verbose
	$uMBX = Get-Mailbox -Identity $uSAM -ErrorAction Stop
	Write-Verbose "User mailbox is: $uMBX" -Verbose
	$uXA4 = ($uMBX).customAttribute4
	Write-Verbose "User has a value in their customAttribute4 AD attribute of: $uXA4" -Verbose
	
	IF (fnGet-RevertDate $uMBX) 
	{
		Write-Verbose "Skipping user $uNAME, MS Exchange extensionCustomAttribute4 is: $(fnGet-RevertDate $uMBX))" -Verbose
	}
	ELSE
	{
		Write-Verbose "Exchange extensionCustomAttribute4 is blank. Continuing process for mailbox reverts..." -Verbose
		$Limit = $(fnSet-Limit)
		$RaiseTimeStamp = [datetime]$uXA4
		IF ($Limit - $RaiseTimeStamp -ge 30)
		{
			IF ((Get-Mailbox -Identity $uSAM).ProhibitSendQuota.Value) {
				$uPSQ = (Get-Mailbox $uSAM).ProhibitSendQuota.Value.ToMB()
				Write-Verbose "Current mailbox send quota is: $uPSQ" -Verbose
			}
			IF ((Get-Mailbox -Identity $uSAM).ProhibitSendReceiveQuota.Value) {
				$uPSRQ = (Get-Mailbox $uSAM).ProhibitSendReceiveQuota.Value.ToMB()
				Write-Verbose "Current mailbox send receive quota is: $uPSRQ" -Verbose
			}
			IF ((Get-Mailbox -Identity $uSAM).IssueWarningQuota.Value) {
				$uIWQ = (Get-Mailbox $uSAM).IssueWarningQuota.Value.ToMB()
				Write-Verbose "Current mailbox warning quota is: $uIWQ" -Verbose
			}
			$newPSQ = $uPSQ - $intPSQBumpLimit
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
			Set-Mailbox -Identity $uMBX -ProhibitSendQuota $strNewPSQ -ProhibitSendReceiveQuota $strNewPSRQ -IssueWarningQuota $strNewIWQ -ExtensionCustomAttribute4 $(Get-TodaysDate)
			Sleep -Seconds 30
			$row = $table.NewRow()
			$row.Name = $uMBX.Name
			$row.Office = $uMBX.Office
			$row.DateQuotaIncreased = $uMBX.customAttribute4
			$row.OldProhibitSendQuotaMB = $uPSQ
			$row.OldProhibitSendReceiveQuotaMB  = $uPSQ * 2
			$row.OldIssueWarningQuotaMB  = $uIWQ
			$row.DateQuotaReset = fnGet-RevertDate $uMBX
			$row.NewProhibitSendQuotaMB = $strNewPSQ
			$row.NewProhibitSendReceiveQuotaMB  = $strNewPSRQ
			$row.NewIssueWarningQuotaMB  = $strNewIWQ
			
			$htmlBody += "<font face=""Arial"">"
			$htmlBody += "<TR align=center>"
			$htmlBody += "<TD>" + $row.Name + "</TD>"
			$htmlBody += "<TD>" + $row.Office + "</TD>"
			$htmlBody += "<TD>" + $row.DateQuotaIncreased + "</TD>"
			$htmlBody += "<TD>" + $row.OldProhibitSendQuotaMB + "MB</TD>"
			$htmlBody += "<TD>" + $row.OldProhibitSendReceiveQuotaMB + "MB</TD>"
			$htmlBody += "<TD>" + $row.OldIssueWarningQuotaMB + "MB</TD>"
			$htmlBody += "<TD>" + $row.DateQuotaReset + "</TD>"
			$htmlBody += "<TD>" + $row.NewProhibitSendQuotaMB + "</TD>"
			$htmlBody += "<TD>" + $row.NewProhibitSendReceiveQuotaMB + "</TD>"
			$htmlBody += "<TD>" + $row.NewIssueWarningQuotaMB + "</TD>"
			$htmlBody += "</TR></font>"
			$table.Rows.Add($row)
		}#End If revert date value check
		fnSend-UserNotification $uFirst $uLast $uMail
	}#End If extensionCustomAttribute4 value check
	$uFirst = $uLast = $uMail = $uName = $uOffice = $uSAM = $uXA4 = $uMBX = $uMBXECA4 = $null
	$Limit = $RaiseTimeStamp = $uPSQ = $uPSRQ = $uIWQ = $newPSQ = $newPSRQ = $newIWQ = $strNewPSQ = $strNewPSRQ = $strNewIWQ = $null
	$iCounter++
}#End ForEach loop
$htmlBody += "</Table></BODY></HTML>"
$AdminMsg.IsBodyHTML = $true
$AdminMsg.Body += $htmlBody
$smtp.Send($AdminMsg)
#EndRegion