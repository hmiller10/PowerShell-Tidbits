#Region Help
<#
.NOTES
Segments of this script have been taken from existing Powershell Scripts
owned by the AD/Messaging team, and have been reused for this script.

.SYNOPSIS
E-mail report indicating all users who have had their mailbox quotas 
increased that day

.DESCRIPTION 
This script will query Active Directory for extensionAttribute4 and will
compare the date stamped in extensionAttribute4 with today's date and if the
dates match will add that user to the daily quota change report along with
current mailbox quota settings.

.OUTPUTS
Notification e-mail for Administrators indicating what users had their 
mailbox quotas increased that day.

.EXAMPLE 
.\Get-DailyQuotaChanges.ps1

#>
###########################################################################
#
# AUTHOR:  Heather Miller
#		   
#
# VERSION HISTORY:
# 5.0 4/23/2014 - Initial Release
#
###########################################################################
#EndRegion

#Region ExecutionPolicy
#Set Execution Policy for Powershell
Set-ExecutionPolicy Unrestricted
#EndRegion

#Region Modules
#Import Required Modules
Import-Module ActiveDirectory
#EndRegion

#Region Variables
#Initialize Variables
[int] $intPSQBumpLimit = 1024
[int] $intWQLimit = 308
[Array]$Properties = @("distinguishedName", "extensionAttribute4", "Name",  "physicalDeliveryOfficeName", "samAccountName")
$RootDSE = Get-ADRootDSE
$Root = ($RootDSE).defaultNamingContext
$ScriptName = $MyInvocation.MyCommand.Name
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

Function fnGet-TodaysDate {#Begin function to set report date
	Get-Date -Format "yyyy-MM-dd"
}#End function fnGet-TodaysDate
#EndRegion





#Region Script
#Begin Script
#Pre-configure Administrative report notification parameters 
$AdminMsg = New-Object Net.Mail.MailMessage
$smtp = New-Object Net.Mail.SmtpClient($smtpServer)
$AdminMsg.To.Add("Administrator1@smtp.mailserver.com")
$AdminMsg.Cc.Add("team@smtp.mailserver.com")
$AdminMsg.From = "MailboxQuotaChecker@domain.com"
$AdminMsg.Subject = "IMPORTANT: Daily mailbox quota alterations for $(fnGet-TodaysDate)"
$AdminMsg.Body = @()


#Get list of users to examine
$Users = Get-ADUser -Filter {Enabled -eq $true} -Properties $Properties -Searchbase $Root -SearchScope Subtree -ResultSetSize $null | `
Where-Object {$_.distinguishedName -like "*OU=External Accounts*" -or $_.distinguishedName -like "*OU=Users*" `
-or $_.distinguishedName -like "*CN=Users*" -and $_.extensionAttribute4 -ne $null}

[Int32]$iCounter = $null
Write-Verbose  -Message "Checking user accounts with quotas to be checked. Stand-By...." -Verbose
ForEach ($User in $Users) 
{
	$uSAM = ($User).samAccountName
	$uName = ($User).Name
	Write-Verbose  -Message "User is $uName" -Verbose
	$uOffice = ($User).physicalDeliveryOfficeName
	#Get contents of user's extensionAttribute4
	$uXA4 = ($User).extensionAttribute4
	Write-Verbose  -Message "User has a value in their extensionAttribute4 AD attribute of: $uXA4" -Verbose
	$AttribDate = [datetime]$uXA4
	Write-Verbose -Message "AttribDate is: $AttribDate" -Verbose
	
	IF ($AttribDate -eq $(Get-TodaysDate))
	{
		Write-Verbose -Message "Date Match" -Verbose
		#Run function to connect to Microsoft Exchange
		Write-Verbose  -Message "Checking for Microsoft.Exchange.Management.PowerShell.E2010 PSSnapin" -Verbose
		IF ( (Get-PSSnapin -Name Microsoft.Exchange.Management.PowerShell.E2010 -ErrorAction SilentlyContinue) -eq $NULL)
		{
			Write-Verbose  -Message "Microsoft.Exchange.Management.PowerShell.E2010 not installed, intalling now" -Verbose
			#Add-PSSnapin -Name Microsoft.Exchange.Management.PowerShell.E2010
			fnGet-LabExchangeConnection
		}

		#Prepare Admin E-mail HTML Header
		$htmlTblData += "<!DOCTYPE html>
			<font face=""Arial"">
			<HTML><TITLE> Company Daily Quota Changes Report</TITLE>
			<BODY background-color:white>
			<H2>Company Exchange 2010 Mailbox Quota Changes for $(Get-TodaysDate)</H2>
			<Table border=1 cellpadding=0 cellspacing=0>
			<TR bgcolor=navy align=center>
			<font face=""Arial"" color:white;font-size:9px>
			<TD><B>Employee Name</B></TD>
			<TD><B>Office</B></TD>
			<TD><B>Date Quota Increased</B></TD>
			<TD><B>Prohibit Send Quota</B></TD>
			<TD><B>Prohibit Send Receive Quota</B></TD>
			<TD><B>Issue Warning Quota</B></TD></TR></font>"

		#Build Admin Report Table
		$table = New-Object system.Data.DataTable "Company Quota Changes Report"
		$col1 = New-Object system.Data.DataColumn EmployeeName,([string])
		$col2 = New-Object system.Data.DataColumn Office,([string])
		$col3 = New-Object system.Data.DataColumn DateQuotaIncreased,([string])
		$col4 = New-Object system.Data.DataColumn ProhibitSendQuotaMB,([string])
		$col5 = New-Object system.Data.DataColumn ProhibitSendReceiveQuotaMB,([string])
		$col6 = New-Object system.Data.DataColumn IssueWarningQuotaMB,([string])

		$table.columns.add($col1)
		$table.columns.add($col2)
		$table.columns.add($col3)
		$table.columns.add($col4)
		$table.columns.add($col5)
		$table.columns.add($col6)

		$uMBX = Get-Mailbox -Identity $uSAM -ErrorAction Stop
		Write-Verbose  -Message "User mailbox is: $uMBX" -Verbose
		
		#Get mailbox quota settings for user
		IF ((Get-Mailbox -Identity $uSAM).ProhibitSendQuota.Value) {
			$uPSQ = (Get-Mailbox $uSAM).ProhibitSendQuota.Value.ToMB()
			Write-Verbose  -Message "Current mailbox send quota is: $uPSQ MB" -Verbose
		}
		IF ((Get-Mailbox -Identity $uSAM).ProhibitSendReceiveQuota.Value) {
			$uPSRQ = (Get-Mailbox $uSAM).ProhibitSendReceiveQuota.Value.ToMB()
			Write-Verbose  -Message "Current mailbox send receive quota is: $uPSRQ MB" -Verbose
		}
		IF ((Get-Mailbox -Identity $uSAM).IssueWarningQuota.Value) {
			$uIWQ = (Get-Mailbox $uSAM).IssueWarningQuota.Value.ToMB()
			Write-Verbose  -Message "Current mailbox warning quota is: $uIWQ MB" -Verbose
		}

		#Add user object information to custom Powershell object
		#Sleep -Seconds 10
		$row = $table.NewRow()
		$row.EmployeeName = $uName
		$row.Office = $uOffice
		$row.DateQuotaIncreased = $uXA4
		$row.ProhibitSendQuotaMB = $uPSQ
		$row.ProhibitSendReceiveQuotaMB  = $uPSRQ
		$row.IssueWarningQuotaMB  = $uIWQ
		
		$htmlTblData += "<font face=""Arial"">"
		$htmlTblData += "<TR align=center>"
		$htmlTblData += "<TD>" + $row.EmployeeName + "</TD>"
		$htmlTblData += "<TD>" + $row.Office + "</TD>"
		$htmlTblData += "<TD>" + $row.DateQuotaIncreased + "</TD>"
		$htmlTblData += "<TD>" + $row.ProhibitSendQuotaMB + "MB</TD>"
		$htmlTblData += "<TD>" + $row.ProhibitSendReceiveQuotaMB + "MB</TD>"
		$htmlTblData += "<TD>" + $row.IssueWarningQuotaMB + "MB</TD>"
		$htmlTblData += "</TR></font>"
		$table.Rows.Add($row)
		$iCounter++
	}
	$uSAM = $uName = $uOffice = $uXA4 = $AttribDate = $null
	$Today = $uMBX = $uPSQ = $uPSRQ = $uIWQ = $null
}#End ForEach loop
$htmlTblData += "</Table></BODY></HTML>"
$AdminMsg.Body += $htmlTblData

$htmlTbl2Data += "<!DOCTYPE html>
	<HTML><TITLE> Total Count of Company Quota Increases</TITLE>
	<BODY background-color:white>
	<H2>Total # Company Exchange 2010 Mailbox Quotas Increased on $(Get-TodaysDate)</H2>
	<Table border=1 cellpadding=0 cellspacing=0>
	<TR bgcolor=navy align=center>
	<font face=""Arial"" color:white;font-size:9px>
	<TD><B>Total Users Increased</B></TD></TR></font>"

$table2 = New-Object system.Data.DataTable "Total Count Daily Mbx Quota Bump"
$col1 = New-Object system.Data.DataColumn TotalUsersIncreased,([Int32])

$table2.columns.add($col1)

$row = $table2.NewRow()
$row.TotalUsersIncreased  = [Int32]$iCounter

$htmlTbl2Data += "<font face=""Arial"">"
$htmlTbl2Data += "<TR align=center>"
$htmlTbl2Data += "<TD>" + [Int32]$iCounter + "</TD>"
$htmlTbl2Data += "</TR></font>"
$table2.Rows.Add($row)
$htmlTbl2Data += "</Table></BODY></HTML>"
$AdminMsg.Body += $htmlTbl2Data
$AdminMsg.IsBodyHTML = $true
$smtp.Send($AdminMsg)

#EndRegion