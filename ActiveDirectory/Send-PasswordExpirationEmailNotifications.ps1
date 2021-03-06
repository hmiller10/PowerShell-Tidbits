#Requires -Modules ActiveDirectory
#Requires -Version 2.0
#Region Help
<#
.NOTES
Portions of this script have been reused from existing code, function and modules
owned by the AD/Messaging team and found on the Internet. Where applicable snippets
have been reused in this script. Special thanks to Martin Pugh,
http://www.thesurlyadmin.com

.SYNOPSIS
.\Email-PwdExpNote.ps1 
E-mail password expiration notifications to all users whose passwords are expiring

.DESCRIPTION 
This script will query Active Directory for all user's whose passwords are going to
expire within 'x' number of days according to the global password age policy set for
the domain

.OUTPUTS
Notification e-mail for each user whose password is about to expire.

.EXAMPLE 
.\Email-PwdExpNote.ps1

#>
###########################################################################
#
# AUTHOR:  Heather Miller
#
# VERSION HISTORY:
# 1.0 7/16/2014 - Initial Release
#
###########################################################################
#EndRegion

#Region ExecutionPolicy
#Set Execution Policy for Powershell
Set-ExecutionPolicy Unrestricted
#EndRegion

#Region Modules
#Check if required module is loaded, if not load import it
IF (-not(Get-Module ActiveDirectory))
{
	Import-Module ActiveDirectory
}
#EndRegion

#Region Variables
#Dim variables
$Days_To_Expiration = 13
$maxPwdAgeTimespan = (Get-ADDefaultDomainPasswordPolicy).MaxPasswordAge.Days
$Properties = @("givenName","lastLogon","mail","Name","PasswordExpired","physicalDeliveryOfficeName","pwdLastSet","samAccountName","sn","userAccountControl")
$RootDSE = Get-ADRootDSE
$Root = $RootDSE.defaultNamingContext
$ScriptName = $MyInvocation.MyCommand.Name
$smtpServer = '<smtp-server-IPaddress'
[PSObject[]]$UserExpiringArray = @()
[PSObject[]]$UserExpiredArray = @()
[PSObject[]]$UserNLOArray = @()
[PSObject[]]$Site1UserExpiringArray = @()
[PSObject[]]$Site2UserExpiringArray = @()
[PSObject[]]$Site3UserExpiringArray = @()
[PSObject[]]$Site4UserExpiringArray = @()
[PSObject[]]$Site5UserExpiringArray = @()
[PSObject[]]$Site6UserExpiringArray = @()
[PSObject[]]$Site7UserExpiringArray = @()
[PSObject[]]$Site8UserExpiringArray = @()
[PSObject[]]$Site9UserExpiringArray = @()
[PSObject[]]$Site10UserExpiringArray = @()
[PSObject[]]$Site11UserExpiringArray = @()
[PSObject[]]$Site12UserExpiringArray = @()
[PSObject[]]$Site13UserExpiringArray = @()
[PSObject[]]$Site14UserExpiringArray = @()
[PSObject[]]$Site15UserExpiringArray = @()
[PSObject[]]$Site16UserExpiringArray = @()
[PSObject[]]$Site17UserExpiringArray = @()
[PSObject[]]$Site18UserExpiringArray = @()
[PSObject[]]$Site19UserExpiringArray = @()
[PSObject[]]$Site20UserExpiringArray = @()
#EndRegion

#Region Functions
#Functions
Function Get-TodaysDate {#Begin function to get short date
	Get-Date -Format "MM-dd-yyyy"
}#End function Get-TodaysDate

Function fnGet-TodaysDate {#Begin function to get today's date
	Get-Date
}#End function fnGet-TodaysDate

Function Get-LongDate {#Begin function to get date and time in long format
	Get-Date -Format G
}#End function Get-LongDate

Function Get-ReportDate {#Begin function set report date format
	Get-Date -Format "yyyy-MM-dd"
}#End function Get-ReportDate

Function Set-AlternatingRows {
	<#
	.SYNOPSIS
		Simple function to alternate the row colors in an HTML table
	.DESCRIPTION
		This function accepts pipeline input from ConvertTo-HTML or any
		string with HTML in it.  It will then search for <tr> and replace 
		it with <tr class=(something)>.  With the combination of CSS it
		can set alternating colors on table rows.
		
		CSS requirements:
		.odd  { background-color:#ffffff; }
		.even { background-color:#dddddd; }
		
		Classnames can be anything and are configurable when executing the
		function.  Colors can, of course, be set to your preference.
		
		This function does not add CSS to your report, so you must provide
		the style sheet, typically part of the ConvertTo-HTML cmdlet using
		the -Head parameter.
	.PARAMETER Line
		String containing the HTML line, typically piped in through the
		pipeline.
	.PARAMETER CSSEvenClass
		Define which CSS class is your "even" row and color.
	.PARAMETER CSSOddClass
		Define which CSS class is your "odd" row and color.
	.EXAMPLE $Report | ConvertTo-HTML -Head $Header | Set-AlternateRows -CSSEvenClass even -CSSOddClass odd | Out-File HTMLReport.html
	
		$Header can be defined with a here-string as:
		$Header = @"
		<style>
		TABLE {border-width: 1px;border-style: solid;border-color: black;border-collapse: collapse;}
		TH {border-width: 1px;padding: 3px;border-style: solid;border-color: black;background-color: #6495ED;}
		TD {border-width: 1px;padding: 3px;border-style: solid;border-color: black;}
		.odd  { background-color:#ffffff; }
		.even { background-color:#dddddd; }
		</style>
		"@
		
		This will produce a table with alternating white and grey rows.  Custom CSS
		is defined in the $Header string and included with the table thanks to the -Head
		parameter in ConvertTo-HTML.
	.NOTES
		Author:         Martin Pugh
		Twitter:        @thesurlyadm1n
		Spiceworks:     Martin9700
		Blog:           www.thesurlyadmin.com
		
		Changelog:
			1.1         Modified replace to include the <td> tag, as it was changing the class
                        for the TH row as well.
            1.0         Initial function release
	.LINK
		http://community.spiceworks.com/scripts/show/1745-set-alternatingrows-function-modify-your-html-table-to-have-alternating-row-colors
    .LINK
        http://thesurlyadmin.com/2013/01/21/how-to-create-html-reports/
	#>
    [CmdletBinding()]
   	Param(
       	[Parameter(Mandatory,ValueFromPipeline)]
        [string]$Line,
       
   	    [Parameter(Mandatory)]
       	[string]$CSSEvenClass,
       
        [Parameter(Mandatory)]
   	    [string]$CSSOddClass
   	)
	Begin {
		$ClassName = $CSSEvenClass
	}
	Process {
		If ($Line.Contains("<tr><td>"))
		{	$Line = $Line.Replace("<tr>","<tr class=""$ClassName"">")
			If ($ClassName -eq $CSSEvenClass)
			{	$ClassName = $CSSOddClass
			}
			Else
			{	$ClassName = $CSSEvenClass
			}
		}
		Return $Line
	}
}#End function Set-AlternatingRows

Function Set-CellColor {   
<#
    .SYNOPSIS
        Function that allows you to set individual cell colors in an HTML table
    .DESCRIPTION
        To be used inconjunction with ConvertTo-HTML this simple function allows you
        to set particular colors for cells in an HTML table.  You provide the criteria
        the script uses to make the determination if a cell should be a particular 
        color (property -gt 5, property -like "*Apple*", etc).
        
        You can add the function to your scripts, dot source it to load into your current
        PowerShell session or add it to your $Profile so it is always available.
        
        To dot source:
            .".\Set-CellColor.ps1"
            
    .PARAMETER Property
        Property, or column that you will be keying on.  
    .PARAMETER Color
        Name or 6-digit hex value of the color you want the cell to be
    .PARAMETER InputObject
        HTML you want the script to process.  This can be entered directly into the
        parameter or piped to the function.
    .PARAMETER Filter
        Specifies a query to determine if a cell should have its color changed.  $true
        results will make the color change while $false result will return nothing.
        
        Syntax
        <Property Name> <Operator> <Value>
        
        <Property Name>::= the same as $Property.  This must match exactly
        <Operator>::= "-eq" | "-le" | "-ge" | "-ne" | "-lt" | "-gt"| "-approx" | "-like" | "-notlike" 
            <JoinOperator> ::= "-and" | "-or"
            <NotOperator> ::= "-not"
        
        The script first attempts to convert the cell to a number, and if it fails it will
        cast it as a string.  So 40 will be a number and you can use -lt, -gt, etc, but 40%
        would be cast as a string so you could only use -eq, -ne, -like, etc.  
    .INPUTS
        HTML with table
    .OUTPUTS
        HTML
    .EXAMPLE
        get-process | convertto-html | set-cellcolor -Propety cpu -Color red -Filter "cpu -gt 1000" | out-file c:\test\get-process.html

        Assuming Set-CellColor has been dot sourced, run Get-Process and convert to HTML.  
        Then change the CPU cell to red only if the CPU field is greater than 1000.
        
    .EXAMPLE
        get-process | convertto-html | set-cellcolor cpu red -filter "cpu -gt 1000 -and cpu -lt 2000" | out-file c:\test\get-process.html
        
        Same as Example 1, but now we will only turn a cell red if CPU is greater than 100 
        but less than 2000.
        
    .EXAMPLE
        $HTML = $Data | sort server | ConvertTo-html -head $header | Set-CellColor cookedvalue red -Filter "cookedvalue -gt 1"
        PS C:\> $HTML = $HTML | Set-CellColor Server green -Filter "server -eq 'dc2'"
        PS C:\> $HTML | Set-CellColor Path Yellow -Filter "Path -like ""*memory*""" | Out-File c:\Test\colortest.html
        
        Takes a collection of objects in $Data, sorts on the property Server and converts to HTML.  From there 
        we set the "CookedValue" property to red if it's greater then 1.  We then send the HTML through Set-CellColor
        again, this time setting the Server cell to green if it's "DC2".  One more time through Set-CellColor
        turns the Path cell to Yellow if it contains the word "memory" in it.
        
    .NOTES
        Author:             Martin Pugh
        Twitter:            @thesurlyadm1n
        Spiceworks:         Martin9700
        Blog:               www.thesurlyadmin.com
          
        Changelog:
            1.03            Added error message in case the $Property field cannot be found in the table header
            1.02            Added some additional text to help.  Added some error trapping around $Filter
                            creation.
            1.01            Added verbose output
            1.0             Initial Release
    .LINK
        http://community.spiceworks.com/scripts/show/2450-change-cell-color-in-html-table-with-powershell-set-cellcolor
    #>

    [CmdletBinding()]
    Param (
        [Parameter(Mandatory,Position=0)]
        [string]$Property,
        [Parameter(Mandatory,Position=1)]
        [string]$Color,
        [Parameter(Mandatory,ValueFromPipeline)]
        [Object[]]$InputObject,
        [Parameter(Mandatory)]
        [string]$Filter
    )
    
    Begin {
        Write-Verbose "$(Get-Date): Function Set-CellColor begins"
        If ($Filter)
        {   If ($Filter.ToUpper().IndexOf($Property.ToUpper()) -ge 0)
            {   $Filter = $Filter.ToUpper().Replace($Property.ToUpper(),"`$Value")
                Try {
                    [scriptblock]$Filter = [scriptblock]::Create($Filter)
                }
                Catch {
                    Write-Warning "$(Get-Date): ""$Filter"" caused an error, stopping script!"
                    Write-Warning $Error[0]
                    Exit
                }
            }
            Else
            {   Write-Warning "Could not locate $Property in the Filter, which is required.  Filter: $Filter"
                Exit
            }
        }
    }
    
    Process {
        ForEach ($Line in $InputObject)
        {   If ($Line.IndexOf("<tr><th") -ge 0)
            {   Write-Verbose "$(Get-Date): Processing headers..."
                $Search = $Line | Select-String -Pattern '<th ?[a-z\-:;"=]*>(.*?)<\/th>' -AllMatches
                $Index = 0
                ForEach ($Match in $Search.Matches)
                {   If ($Match.Groups[1].Value -eq $Property)
                    {   Break
                    }
                    $Index ++
                }
                If ($Index -eq $Search.Matches.Count)
                {   Write-Warning "$(Get-Date): Unable to locate property: $Property in table header"
                    Exit
                }
                Write-Verbose "$(Get-Date): $Property column found at index: $Index"
            }
            If ($Line.IndexOf("<tr><td") -ge 0)
            {   $Search = $Line | Select-String -Pattern '<td ?[a-z\-:;"=]*>(.*?)<\/td>' -AllMatches
                $Value = $Search.Matches[$Index].Groups[1].Value -as [double]
                If (-not $Value)
                {   $Value = $Search.Matches[$Index].Groups[1].Value
                }
                If (Invoke-Command $Filter)
                {   Write-Verbose "$(Get-Date): Criteria met!  Changing cell to $Color..."
                    $Line = $Line.Replace($Search.Matches[$Index].Value,"<td style=""background-color:$Color"">$Value</td>")
                }
            }
            Write-Output $Line
        }
    }
    
    End {
        Write-Verbose "$(Get-Date): Function Set-CellColor completed"
    }
}#End function Set-CellColor

Function Send-EmailNotification {#Begin function to send e-mail to end user regarding password expiration
	Param ($Email,$UserName,$pwExpired,$pwValue,$Days,$Days_To_Expiration,$pwExpirationDate,$uFirst,$uLast)
	$GreetingName = $uFirst + " " + $uLast
	If ($pwExpired -eq "True" -and $pwValue -gt 0) {
		$Body = @"
		<p>$(Get-TodaysDate)</p>
		
		<p>Dear $GreetingName,</p>

		<p>Your logon password has expired. </p>
		
		<p>Please contact your local IT HelpDesk at: 123-4567 | 123-456-7890 | 1-234-567-8900  or at Help@YourDomain.com for assistance.</p>

		<p>Thank you.</p>

		<p>*** This is an automatically generated email. Please do not reply. ***</p>
"@
		$Subject = "Urgent: Your logon password already expired"
		Send-MailMessage -From $From -To $Email -Subject $Subject -Body $Body -BodyAsHtml -SmtpServer $smtpServer
	} ElseIf ($Days -le $Days_To_Expiration) {
		$Body = @"
		<p>$(Get-TodaysDate)</p>
		
		<p>Dear $GreetingName,</p>

		<p>Your logon password expires on: $pwExpirationDate which is in $Days_To_Expiration.</p>
		
		<p>Please consider changing your password.</p>
		
		<p>Please contact your local IT HelpDesk at: 123-4567 | 123-456-7890 | 1-234-567-8900  or at Help@YourDomain.com for assistance.</p>

		<p>Thank you.</p>

		<p>*** This is an automatically generated email. Please do not reply. ***</p>
"@
		$From = 'no-reply@YourDomain.com'
		$Subject = "Password for account: $UserName is set to expire."
		Send-MailMessage -From $From -To $Email -Subject $Subject -Body $Body -BodyAsHtml -SmtpServer $smtpServer
	}
}#End function Send-EmailNotification

Function Send-AdminEmail {#Begin function to send summary e-mail to Administrators
	Param ($UserExpiredArray,$UserNLOArray,$UserExpiringArray)
	$Header = @"
	<style>BODY{Background-Color:lightgrey;Font-Family: Arial;Font-Size: 12pt;}
	TABLE {border-width: 1px;border-style: solid;border-color: black;border-collapse: collapse;}
	TH {border-width: 1px;padding: 3px;border-style: solid;border-color: black;color:white;background-color: #003366;Font-Family: Arial;Font-Size: 14pt;Text-Align: Center;}
	TD {border-width: 1px;padding: 3px;border-style: solid;border-color: black;Font-Family: Arial;Font-Size: 12pt;Text-Align: Center;}
	.odd  { background-color:#ffffff; }
	.even { background-color:#dddddd; }
	</style>
"@
	$Pre0 = @"
	<p><br />
	Users whose passwords are expiring soon<br />
	Key:<br /> 
	Table cells shaded in <tr
	<span style="color: yellow;">Yellow:</span> Password expires in 2-5 days.<br /> 
	Table cells shaded in <span style="color: red;">Red:</span> Password expires today or tomorrow.<br />
	</p>
"@
	$Pre1 = @"
	<p><br />
	Users with expired passwords<br />
	</p>
"@
	$Pre2 = @"
	<p><br />
	Users who have never logged on<br />
	</p>
"@
	$Post = @"
	<p><br />
	This is an automatically generated e-mail. Please do not reply.<br />
	$ScriptName run on: $(Get-LongDate)<br />
	</p>
"@
	$From = 'PasswordNotifications@YourDomain.com'
	$HRTo = "Your.Email@YourDomain.com"
	If ($(fnGet-TodaysDate).DayOfWeek -eq 'Monday')
	{
		$HRReportSubject = "MWE users who have expired passwords or have never logged in for week of: $(fnGet-TodaysDate)"
			$HRPre1 = @"
			<p><br />
			Users with expired passwords<br />
			</p>
"@
			$HRPre2 = @"
			<p><br/>
			Users who have never logged on<br />
			</p>
"@
		$HREmailBody += $UserExpiredArray | Select-Object Office,Name,"Last Logon Date","Password Expiration Date","Password Expired" | Sort Office,Name | `
		ConvertTo-HTML -Head $Header -PreContent $Pre1 | Set-AlternatingRows -CSSEvenClass even -CSSOddClass odd | Out-String
		$HREmailBody += $UserNLOArray | Select-Object Office,Name,"Password Expiration Date","Password Last Set" | Sort Office,Name | `
		ConvertTo-HTML -Head $Header -PreContent $HRPre2 -PostContent $Post | Set-AlternatingRows -CSSEvenClass even -CSSOddClass odd | Out-String
	
		Send-MailMessage -From $From -To $HRTo -Subject $HRReportSubject -Body $HREmailBody -BodyAsHTML -SmtpServer $smtpServer
	}
	
	$ReportSubject="Password notifications for $(Get-ReportDate)"
	$To = "Your.Email@YourDomain.com"
	$AdminMsgBody += $UserExpiringArray | Select-Object Office,Name,"Days To Expiration","Password Expiration Date" | Sort Office,Name | `
	ConvertTo-HTML -Head $Header -PreContent $Pre0 | `
	Set-CellColor -Property "Days To Expiration" -Color Yellow -Filter """Days To Expiration"" -le 5 -and ""Days To Expiration"" -ge 2" | `
	Set-CellColor -Property "Days To Expiration" -Color Red -Filter """Days To Expiration"" -le 1 -and ""Days To Expiration"" -ge 0" | `
	Set-AlternatingRows -CSSEvenClass even -CSSOddClass odd | Out-String
	$AdminMsgBody += $UserExpiredArray | Select-Object Office,Name,"Last Logon Date","Password Expiration Date","Password Expired" | Sort Office,Name | `
	ConvertTo-HTML -Head $Header -PreContent $Pre1 | Set-AlternatingRows -CSSEvenClass even -CSSOddClass odd | Out-String
	$AdminMsgBody += $UserNLOArray | Select-Object Office,Name,"Password Expiration Date","Password Last Set" | Sort Office,Name | `
	ConvertTo-HTML -Head $Header -PreContent $Pre2 -PostContent $Post| Set-AlternatingRows -CSSEvenClass even -CSSOddClass odd | Out-String
	Send-MailMessage -From $From -To $To -Subject $ReportSubject -Body $AdminMsgBody -BodyAsHTML -SmtpServer $smtpServer
}#End function Send-AdminEmail
#EndRegion




#Region Script
#Begin Script
#Get list of AD OUs from containers 
$OUs = @()
$OUs = Get-ADOrganizationalUnit -Filter * -SearchBase $Root -SearchScope Subtree -ResultSetSize $null | `
Where-Object {$_.distinguishedName -like "*OU=Users*" -or $_.distinguishedName -like "*OU=External Accounts*" -or $_.distinguishedName -like "*OU=Shared Accounts*"} | `
Select-Object distinguishedName

ForEach ($OU in $OUs) {#Loop through each Organizational Unit
	$ADOU = $OU.distinguishedName
	Write-Host $ADOU -ForegroundColor Gray

	#Gather list of active users to examine into an array
	$Users = @()
	[Array]$Users = Get-ADuser -Filter {(PasswordNeverExpires -eq $false) -and (Enabled -eq $true)} -Properties $Properties -SearchBase $ADOU -SearchScope Subtree `
	-ResultSetSize $null  | Select-Object $Properties | Sort-Object -Property $Properties[4],$Properties[2]
	
	ForEach ($User in $Users) {#Loop through each user account in the Organizational Unit
		
		$uFirst = $User.givenName
		$UserID = $User.samAccountName
		$UserName = $User.Name
		$UserSite = $User.physicalDeliveryOfficeName
		$uLast = $User.sn
		$Email = $User.mail
		#Calculate date user last logged on to domain
		$uLastLogonTime = ([datetime]::FromFileTime((Get-ADUser -Identity $UserID -Properties "lastLogonTimestamp")."lastLogonTimestamp"))
		#Calculate days to password expiration
		$Days =  (([datetime]::FromFileTime((Get-ADUser -Identity $UserID -Properties "msDS-UserPasswordExpiryTimeComputed")."msDS-UserPasswordExpiryTimeComputed")) `
		-(Get-Date)).Days
		#Convert user's password last set attribute to readable date/time variable
		$pwdLastSet = [datetime]::FromFileTime([Int64]::Parse($User.pwdLastSet).ToString('g'))
		$pwExpired = $User.PasswordExpired
		$pwValue = $User.pwdLastSet
		$pwdExpirationDate = $pwdLastSet.AddDays($maxPwdAgeTimespan)
		#Convert password expiration date to string value to write in table row
		$uPWExpDate = [string] $pwdExpirationDate
		#Check if user password is expired and if so send appropriate notification.
		If ($pwExpired -eq "True" -and $pwValue -gt 0)
		{
			$UserExpiredArray += New-Object -TypeName PSCustomObject -Property @{
				Office = $UserSite
				Name = $UserName
				"Last Logon Date" = $uLastLogonTime
				"Password Expiration Date" = $uPWExpDate
				"Password Expired" = $pwExpired
			}
			Write-Host "$UserName in $UserSite domain account password expired on $pwdExpirationDate and will need to be reset by IT personnel." -ForegroundColor Red
			Send-EmailNotification ($Email,$UserName,$pwExpired,$pwValue,$Days,$Days_To_Expiration,$pwExpirationDate)		
		} 
		ElseIf ($pwExpired -eq "True" -and $pwValue -eq 0) 
		{
			$UserNLOArray += New-Object -TypeName PSCustomObject -Property @{
			Office = $UserSite
			Name = $UserName
			"Password Expiration Date" = $uPWExpDate
			"Password Last Set" = "Never"
			}
			Write-Host "$UserName in $UserSite account has not logged in yet. Please check with Human Resources to verify if his/her account should remain active." -ForegroundColor Cyan
		} 
		ElseIf ($Days -le $Days_To_Expiration)
		{
			Switch ($UserSite) 
			{	
				Site1{ $Site1UserExpiringArray += New-Object -TypeName PSCustomObject -Property @{
						Office = $UserSite
						Name = $UserName
						"Days To Expiration" = $Days
						"Password Expiration Date" = $uPWExpDate}
						Write-Host "Adding $UserName to Site1 Expiring Array" -ForegroundColor Green
						; break}
				Site2{ $Site2UserExpiringArray += New-Object -TypeName PSCustomObject -Property @{
						Office = $UserSite
						Name = $UserName
						"Days To Expiration" = $Days
						"Password Expiration Date" = $uPWExpDate}
						Write-Host "Adding $UserName to Site2 Expiring Array" -ForegroundColor Green
						; break}
				Site3{ $Site3UserExpiringArray += New-Object -TypeName PSCustomObject -Property @{
						Office = $UserSite
						Name = $UserName
						"Days To Expiration" = $Days
						"Password Expiration Date" = $uPWExpDate}
						Write-Host "Adding $UserName to Site3 Expiring Array" -ForegroundColor Green
						; break}
				Site4{ $Site4UserExpiringArray += New-Object -TypeName PSCustomObject -Property @{
						Office = $UserSite
						Name = $UserName
						"Days To Expiration" = $Days
						"Password Expiration Date" = $uPWExpDate}
						Write-Host "Adding $UserName to Site4 Expiring Array" -ForegroundColor Green
						; break}
				Site5{ $Site5UserExpiringArray += New-Object -TypeName PSCustomObject -Property @{
						Office = $UserSite
						Name = $UserName
						"Days To Expiration" = $Days
						"Password Expiration Date" = $uPWExpDate}
						Write-Host "Adding $UserName to Site5 Expiring Array" -ForegroundColor Green
						; break}
				Site6{ $Site6UserExpiringArray += New-Object -TypeName PSCustomObject -Property @{
						Office = $UserSite
						Name = $UserName
						"Days To Expiration" = $Days
						"Password Expiration Date" = $uPWExpDate}
						Write-Host "Adding $UserName to Site6 Expiring Array" -ForegroundColor Green
						; break}
				Site7{ $Site7UserExpiringArray += New-Object -TypeName PSCustomObject -Property @{
						Office = $UserSite
						Name = $UserName
						"Days To Expiration" = $Days
						"Password Expiration Date" = $uPWExpDate}
						Write-Host "Adding $UserName to Site7 Expiring Array" -ForegroundColor Green
						; break}
				Site8{ $Site8UserExpiringArray += New-Object -TypeName PSCustomObject -Property @{
						Office = $UserSite
						Name = $UserName
						"Days To Expiration" = $Days
						"Password Expiration Date" = $uPWExpDate}
						Write-Host "Adding $UserName to Site8 Expiring Array" -ForegroundColor Green
						; break}
				Site9{ $Site9UserExpiringArray += New-Object -TypeName PSCustomObject -Property @{
						Office = $UserSite
						Name = $UserName
						"Days To Expiration" = $Days
						"Password Expiration Date" = $uPWExpDate}
						Write-Host "Adding $UserName to Site9 Expiring Array" -ForegroundColor Green
						; break}
				Site10{ $Site10UserExpiringArray += New-Object -TypeName PSCustomObject -Property @{
						Office = $UserSite
						Name = $UserName
						"Days To Expiration" = $Days
						"Password Expiration Date" = $uPWExpDate}
						Write-Host "Adding $UserName to Site10 Expiring Array" -ForegroundColor Green
						; break}
				Site11{ $Site11UserExpiringArray += New-Object -TypeName PSCustomObject -Property @{
						Office = $UserSite
						Name = $UserName
						"Days To Expiration" = $Days
						"Password Expiration Date" = $uPWExpDate}
						Write-Host "Adding $UserName to Site11 Expiring Array" -ForegroundColor Green
						; break}
				Site12{ $Site12UserExpiringArray += New-Object -TypeName PSCustomObject -Property @{
						Office = $UserSite
						Name = $UserName
						"Days To Expiration" = $Days
						"Password Expiration Date" = $uPWExpDate}
						Write-Host "Adding $UserName to Site12 Expiring Array" -ForegroundColor Green
						; break}
				Site13{ $Site13UserExpiringArray += New-Object -TypeName PSCustomObject -Property @{
						Office = $UserSite
						Name = $UserName
						"Days To Expiration" = $Days
						"Password Expiration Date" = $uPWExpDate}
						Write-Host "Adding $UserName to Site13 Expiring Array" -ForegroundColor Green
						; break}
				Site14{ $Site14UserExpiringArray += New-Object -TypeName PSCustomObject -Property @{
						Office = $UserSite
						Name = $UserName
						"Days To Expiration" = $Days
						"Password Expiration Date" = $uPWExpDate}
						Write-Host "Adding $UserName to Site14 Expiring Array" -ForegroundColor Green
						; break}
				Site15{ $Site15UserExpiringArray += New-Object -TypeName PSCustomObject -Property @{
						Office = $UserSite
						Name = $UserName
						"Days To Expiration" = $Days
						"Password Expiration Date" = $uPWExpDate}
						Write-Host "Adding $UserName to Site15 Expiring Array" -ForegroundColor Green
						; break}
				Site16{ $Site16UserExpiringArray += New-Object -TypeName PSCustomObject -Property @{
					Office = $UserSite
					Name = $UserName
					"Days To Expiration" = $Days
					"Password Expiration Date" = $uPWExpDate}
					Write-Host "Adding $UserName to Site16 Expiring Array" -ForegroundColor Green
					; break}
				Site17{ $Site17UserExpiringArray += New-Object -TypeName PSCustomObject -Property @{
						Office = $UserSite
						Name = $UserName
						"Days To Expiration" = $Days
						"Password Expiration Date" = $uPWExpDate}
						Write-Host "Adding $UserName to Site17 Expiring Array" -ForegroundColor Green
						; break}
				Site18{ $Site18UserExpiringArray += New-Object -TypeName PSCustomObject -Property @{
						Office = $UserSite
						Name = $UserName
						"Days To Expiration" = $Days
						"Password Expiration Date" = $uPWExpDate}
						Write-Host "Adding $UserName to Site18 User Expiring Array" -ForegroundColor Green
						; break}
				Site19{ $Site19UserExpiringArray += New-Object -TypeName PSCustomObject -Property @{
								Office = $UserSite
								Name = $UserName
								"Days To Expiration" = $Days
								"Password Expiration Date" = $uPWExpDate}
								Write-Host "Adding $UserName to Site19 Expiring Array" -ForegroundColor Green
								; break}
				Site20{ $Site20UserExpiringArray += New-Object -TypeName PSCustomObject -Property @{
								Office = $UserSite
								Name = $UserName
								"Days To Expiration" = $Days
								"Password Expiration Date" = $uPWExpDate}
								Write-Host "Adding $UserName to Site20 Expiring Array" -ForegroundColor Green
								; break}
			}#End switch statement
			$UserExpiringArray += New-Object -TypeName PSCustomObject -Property @{
			Office = $UserSite
			Name = $UserName
			"Days To Expiration" = $Days
			"Password Expiration Date" = $uPWExpDate}
			Write-Host "$UserName in $UserSite password is set to expire in " $Days " days." -ForegroundColor Yellow
			Send-EmailNotification ($Email,$UserName,$pwExpired,$pwValue,$Days,$Days_To_Expiration,$pwExpirationDate)
		}#End ElseIf
	}#End Users loop
}#End OU loop
#Send local IT staff a report of user's in their respetive office.
$UserObjects = @()
$UserObjects = ($Site1UserExpiringArray,$Site2UserExpiringArray,$Site3UserExpiringArray,$Site4UserExpiringArray,$Site5UserExpiringArray,$Site6UserExpiringArray,`
$Site7UserExpiringArray,$Site8UserExpiringArray,$Site9UserExpiringArray,$Site10UserExpiringArray,$Site11UserExpiringArray,$Site12UserExpiringArray,$Site13UserExpiringArray, `
$Site14UserExpiringArray,$Site15UserExpiringArray,$Site16UserExpiringArray,$Site17UserExpiringArray,$Site18UserExpiringArray,$Site19UserExpiringArray,$Site20UserExpiringArray)

$From = 'PasswordNotifications@YourDomain.com'
$ReportSubject="Password expiration notifications for $(Get-ReportDate)"
ForEach ($Object in $UserObjects)
{#Begin processing local office reports
	$Header = @"
	<style>BODY{Background-Color:lightgrey;Font-Family: Arial;Font-Size: 12pt;}
	TABLE {border-width: 1px;border-style: solid;border-color: black;border-collapse: collapse;}
	TH {border-width: 1px;padding: 3px;border-style: solid;border-color: black;color:white;background-color: #003366;Font-Family: Arial;Font-Size: 14pt;Text-Align: Center;}
	TD {border-width: 1px;padding: 3px;border-style: solid;border-color: black;Font-Family: Arial;Font-Size: 12pt;Text-Align: Center;}
	.odd  { background-color:#ffffff; }
	.even { background-color:#dddddd; }
	</style>
"@
	$Pre = @"
	<p><br />
	Users whose passwords are expiring soon<br />
	Key:<br /> 
	Table cells shaded in <span style="color: yellow;">Yellow:</span> Password expires in 2-5 days.<br /> 
	Table cells shaded in <span style="color: red;">Red:</span> Password expires today or tomorrow.<br />
	</p>
"@
	$Post = @"
	<p><br />
	This is an automatically generated e-mail. Please do not reply.<br />
	$ScriptName run on: $(Get-LongDate)<br />
	</p>
"@
	
	Switch -wildcard ($Object.Office)
	{
		Site1{Write-Host "Processing local IT report for: Site1"
				$Site1Body += $Site1UserExpiringArray | Select-Object Office,Name,"Days To Expiration","Password Expiration Date" | Sort Office,Name | `
				ConvertTo-HTML -Head $Header -PreContent $Pre -PostContent $Post | `
				Set-CellColor -Property "Days To Expiration" -Color Yellow -Filter """Days To Expiration"" -le 5 -and ""Days To Expiration"" -ge 2" | `
				Set-CellColor -Property "Days To Expiration" -Color Red -Filter """Days To Expiration"" -le 1 -and ""Days To Expiration"" -ge 0" | `
				Set-AlternatingRows -CSSEvenClass even -CSSOddClass odd | Out-String
				$ITTo = "Site1 Help Desk"
				Send-MailMessage -From $From -To $ITTo -Subject $ReportSubject -Body $Site1Body -BodyAsHtml -SmtpServer $smtpServer
			; break}
		Site2{Write-Host "Processing local IT report for: Site2"
				$Site2Body += $Site2UserExpiringArray | Select-Object Office,Name,"Days To Expiration","Password Expiration Date" | Sort Office,Name | `
				ConvertTo-HTML -Head $Header -PreContent $Pre -PostContent $Post | `
				Set-CellColor -Property "Days To Expiration" -Color Yellow -Filter """Days To Expiration"" -le 5 -and ""Days To Expiration"" -ge 2" | `
				Set-CellColor -Property "Days To Expiration" -Color Red -Filter """Days To Expiration"" -le 1 -and ""Days To Expiration"" -ge 0" | `
				Set-AlternatingRows -CSSEvenClass even -CSSOddClass odd | Out-String
				$ITTo = "Site2 Help Desk"
				Send-MailMessage -From $From -To $ITTo -Subject $ReportSubject -Body $Site2Body -BodyAsHtml -SmtpServer $smtpServer
			; break}
		Site3{Write-Host "Processing local IT report for: Site3"
				$Site3Body += $Site3UserExpiringArray | Select-Object Office,Name,"Days To Expiration","Password Expiration Date" | Sort Office,Name | `
				ConvertTo-HTML -Head $Header -PreContent $Pre -PostContent $Post | `
				Set-CellColor -Property "Days To Expiration" -Color Yellow -Filter """Days To Expiration"" -le 5 -and ""Days To Expiration"" -ge 2" | `
				Set-CellColor -Property "Days To Expiration" -Color Red -Filter """Days To Expiration"" -le 1 -and ""Days To Expiration"" -ge 0" | `
				Set-AlternatingRows -CSSEvenClass even -CSSOddClass odd | Out-String
				$ITTo = "Site3 Help Desk"
				Send-MailMessage -From $From -To $ITTo -Subject $ReportSubject -Body $Site3Body -BodyAsHtml -SmtpServer $smtpServer
			; break}
		Site4{Write-Host "Processing local IT report for: Site4"
				$Site4Body += $Site4UserExpiringArray | Select-Object Office,Name,"Days To Expiration","Password Expiration Date" | Sort Office,Name | `
				ConvertTo-HTML -Head $Header -PreContent $Pre -PostContent $Post | `
				Set-CellColor -Property "Days To Expiration" -Color Yellow -Filter """Days To Expiration"" -le 5 -and ""Days To Expiration"" -ge 2" | `
				Set-CellColor -Property "Days To Expiration" -Color Red -Filter """Days To Expiration"" -le 1 -and ""Days To Expiration"" -ge 0" | `
				Set-AlternatingRows -CSSEvenClass even -CSSOddClass odd | Out-String
				$ITTo = "Site4 Help Desk"
				Send-MailMessage -From $From -To $ITTo -Subject $ReportSubject -Body $Site4Body -BodyAsHtml -SmtpServer $smtpServer
			; break}
		Site5{Write-Host "Processing local IT report for: Site5"
				$Site5Body += $Site5UserExpiringArray | Select-Object Office,Name,"Days To Expiration","Password Expiration Date" | Sort Office,Name | `
				ConvertTo-HTML -Head $Header -PreContent $Pre -PostContent $Post | `
				Set-CellColor -Property "Days To Expiration" -Color Yellow -Filter """Days To Expiration"" -le 5 -and ""Days To Expiration"" -ge 2" | `
				Set-CellColor -Property "Days To Expiration" -Color Red -Filter """Days To Expiration"" -le 1 -and ""Days To Expiration"" -ge 0" | `
				Set-AlternatingRows -CSSEvenClass even -CSSOddClass odd | Out-String
				$ITTo = "Site5 Help Desk"
				Send-MailMessage -From $From -To $ITTo -Subject $ReportSubject -Body $Site5Body -BodyAsHtml -SmtpServer $smtpServer
			; break}
		Site6{Write-Host "Processing local IT report for: Site6"
				$Site6Body += $Site6UserExpiringArray | Select-Object Office,Name,"Days To Expiration","Password Expiration Date" | Sort Office,Name | `
				ConvertTo-HTML -Head $Header -PreContent $Pre -PostContent $Post | `
				Set-CellColor -Property "Days To Expiration" -Color Yellow -Filter """Days To Expiration"" -le 5 -and ""Days To Expiration"" -ge 2" | `
				Set-CellColor -Property "Days To Expiration" -Color Red -Filter """Days To Expiration"" -le 1 -and ""Days To Expiration"" -ge 0" | `
				Set-AlternatingRows -CSSEvenClass even -CSSOddClass odd | Out-String
				$ITTo = "Site6 Help Desk"
				Send-MailMessage -From $From -To $ITTo -Subject $ReportSubject -Body $Site6Body -BodyAsHtml -SmtpServer $smtpServer
			; break}
		Site7{Write-Host "Processing local IT report for: Site7"
				$Site7Body += $Site7UserExpiringArray | Select-Object Office,Name,"Days To Expiration","Password Expiration Date" | Sort Office,Name | `
				ConvertTo-HTML -Head $Header -PreContent $Pre -PostContent $Post | `
				Set-CellColor -Property "Days To Expiration" -Color Yellow -Filter """Days To Expiration"" -le 5 -and ""Days To Expiration"" -ge 2" | `
				Set-CellColor -Property "Days To Expiration" -Color Red -Filter """Days To Expiration"" -le 1 -and ""Days To Expiration"" -ge 0" | `
				Set-AlternatingRows -CSSEvenClass even -CSSOddClass odd | Out-String
				$ITTo = "Site7 Help Desk"
				Send-MailMessage -From $From -To $ITTo -Subject $ReportSubject -Body $Site7Body -BodyAsHtml -SmtpServer $smtpServer
			; break}
		Site8{Write-Host "Processing local IT report for: Site8"
				$Site8Body += $Site8UserExpiringArray | Select-Object Office,Name,"Days To Expiration","Password Expiration Date" | Sort Office,Name | `
				ConvertTo-HTML -Head $Header -PreContent $Pre -PostContent $Post | `
				Set-CellColor -Property "Days To Expiration" -Color Yellow -Filter """Days To Expiration"" -le 5 -and ""Days To Expiration"" -ge 2" | `
				Set-CellColor -Property "Days To Expiration" -Color Red -Filter """Days To Expiration"" -le 1 -and ""Days To Expiration"" -ge 0" | `
				Set-AlternatingRows -CSSEvenClass even -CSSOddClass odd | Out-String
				$ITTo = "Site8 Help Desk"
				Send-MailMessage -From $From -To $ITTo -Subject $ReportSubject -Body $Site8Body -BodyAsHtml -SmtpServer $smtpServer
			; break}
		Site9{Write-Host "Processing local IT report for: Site9"
				$Site9 += $Site9UserExpiringArray | Select-Object Office,Name,"Days To Expiration","Password Expiration Date" | Sort Office,Name | `
				ConvertTo-HTML -Head $Header -PreContent $Pre -PostContent $Post | `
				Set-CellColor -Property "Days To Expiration" -Color Yellow -Filter """Days To Expiration"" -le 5 -and ""Days To Expiration"" -ge 2" | `
				Set-CellColor -Property "Days To Expiration" -Color Red -Filter """Days To Expiration"" -le 1 -and ""Days To Expiration"" -ge 0" | `
				Set-AlternatingRows -CSSEvenClass even -CSSOddClass odd | Out-String
				$ITTo = "Site9 Help Desk"
				Send-MailMessage -From $From -To $ITTo -Subject $ReportSubject -Body $Site9Body -BodyAsHtml -SmtpServer $smtpServer
			; break}
		Site10{Write-Host "Processing local IT report for: Site10"
				$Site10Body += $Site10UserExpiringArray | Select-Object Office,Name,"Days To Expiration","Password Expiration Date" | Sort Office,Name | `
				ConvertTo-HTML -Head $Header -PreContent $Pre -PostContent $Post | `
				Set-CellColor -Property "Days To Expiration" -Color Yellow -Filter """Days To Expiration"" -le 5 -and ""Days To Expiration"" -ge 2" | `
				Set-CellColor -Property "Days To Expiration" -Color Red -Filter """Days To Expiration"" -le 1 -and ""Days To Expiration"" -ge 0" | `
				Set-AlternatingRows -CSSEvenClass even -CSSOddClass odd | Out-String
				$ITTo = "Site10 Help Desk"
				Send-MailMessage -From $From -To $ITTo -Subject $ReportSubject -Body $Site10Body -BodyAsHtml -SmtpServer $smtpServer
			; break}
		Site11{Write-Host "Processing local IT report for: Site11"
				$Site11Body += $Site11UserExpiringArray | Select-Object Office,Name,"Days To Expiration","Password Expiration Date" | Sort Office,Name | `
				ConvertTo-HTML -Head $Header -PreContent $Pre -PostContent $Post | `
				Set-CellColor -Property "Days To Expiration" -Color Yellow -Filter """Days To Expiration"" -le 5 -and ""Days To Expiration"" -ge 2" | `
				Set-CellColor -Property "Days To Expiration" -Color Red -Filter """Days To Expiration"" -le 1 -and ""Days To Expiration"" -ge 0" | `
				Set-AlternatingRows -CSSEvenClass even -CSSOddClass odd | Out-String
				$ITTo = "Site11 Help Desk"
				Send-MailMessage -From $From -To $ITTo -Subject $ReportSubject -Body $Site11Body -BodyAsHtml -SmtpServer $smtpServer
			; break}
		Site12{Write-Host "Processing local IT report for: Site12"
				$Site12Body += $Site12UserExpiringArray | Select-Object Office,Name,"Days To Expiration","Password Expiration Date" | Sort Office,Name | `
				ConvertTo-HTML -Head $Header -PreContent $Pre -PostContent $Post | `
				Set-CellColor -Property "Days To Expiration" -Color Yellow -Filter """Days To Expiration"" -le 5 -and ""Days To Expiration"" -ge 2" | `
				Set-CellColor -Property "Days To Expiration" -Color Red -Filter """Days To Expiration"" -le 1 -and ""Days To Expiration"" -ge 0" | `
				Set-AlternatingRows -CSSEvenClass even -CSSOddClass odd | Out-String
				$ITTo = "Site12 Help Desk"
				Send-MailMessage -From $From -To $ITTo -Subject $ReportSubject -Body $Site12Body -BodyAsHtml -SmtpServer $smtpServer
			; break}
		Site13{Write-Host "Processing local IT report for: Site13"
				$Site13Body += $Site13UserExpiringArray | Select-Object Office,Name,"Days To Expiration","Password Expiration Date" | Sort Office,Name | `
				ConvertTo-HTML -Head $Header -PreContent $Pre -PostContent $Post | `
				Set-CellColor -Property "Days To Expiration" -Color Yellow -Filter """Days To Expiration"" -le 5 -and ""Days To Expiration"" -ge 2" | `
				Set-CellColor -Property "Days To Expiration" -Color Red -Filter """Days To Expiration"" -le 1 -and ""Days To Expiration"" -ge 0" | `
				Set-AlternatingRows -CSSEvenClass even -CSSOddClass odd | Out-String
				$ITTo = "Site13 Help Desk"
				Send-MailMessage -From $From -To $ITTo -Subject $ReportSubject -Body $Site13Body -BodyAsHtml -SmtpServer $smtpServer
			; break}
		Site14{Write-Host "Processing local IT report for: Site14"
				$Site14Body += $Site14UserExpiringArray | Select-Object Office,Name,"Days To Expiration","Password Expiration Date" | Sort Office,Name | `
				ConvertTo-HTML -Head $Header -PreContent $Pre -PostContent $Post | `
				Set-CellColor -Property "Days To Expiration" -Color Yellow -Filter """Days To Expiration"" -le 5 -and ""Days To Expiration"" -ge 2" | `
				Set-CellColor -Property "Days To Expiration" -Color Red -Filter """Days To Expiration"" -le 1 -and ""Days To Expiration"" -ge 0" | `
				Set-AlternatingRows -CSSEvenClass even -CSSOddClass odd | Out-String
				$ITTo = "Site14 Help Desk"
				Send-MailMessage -From $From -To $ITTo -Subject $ReportSubject -Body $Site14Body -BodyAsHtml -SmtpServer $smtpServer
			; break}
		Site15{Write-Host "Processing local IT report for: Site15"
				$Site15Body += $Site15UserExpiringArray | Select-Object Office,Name,"Days To Expiration","Password Expiration Date" | Sort Office,Name | `
				ConvertTo-HTML -Head $Header -PreContent $Pre -PostContent $Post | `
				Set-CellColor -Property "Days To Expiration" -Color Yellow -Filter """Days To Expiration"" -le 5 -and ""Days To Expiration"" -ge 2" | `
				Set-CellColor -Property "Days To Expiration" -Color Red -Filter """Days To Expiration"" -le 1 -and ""Days To Expiration"" -ge 0" | `
				Set-AlternatingRows -CSSEvenClass even -CSSOddClass odd | Out-String
				$ITTo = "Site15 Help Desk"
				Send-MailMessage -From $From -To $ITTo -Subject $ReportSubject -Body $Site15Body -BodyAsHtml -SmtpServer $smtpServer
			; break}
		Site16{Write-Host "Processing local IT report for: Site16"
				$Site16Body += $Site16UserExpiringArray | Select-Object Office,Name,"Days To Expiration","Password Expiration Date" | Sort Office,Name | `
				ConvertTo-HTML -Head $Header -PreContent $Pre -PostContent $Post | `
				Set-CellColor -Property "Days To Expiration" -Color Yellow -Filter """Days To Expiration"" -le 5 -and ""Days To Expiration"" -ge 2" | `
				Set-CellColor -Property "Days To Expiration" -Color Red -Filter """Days To Expiration"" -le 1 -and ""Days To Expiration"" -ge 0" | `
				Set-AlternatingRows -CSSEvenClass even -CSSOddClass odd | Out-String
				$ITTo = "Site16 Help Desk"
				Send-MailMessage -From $From -To $ITTo -Subject $ReportSubject -Body $Site16Body -BodyAsHtml -SmtpServer $smtpServer
			; break}
		Site17{Write-Host "Processing local IT report for: Site17"
				$Site17Body += $Site17UserExpiringArray | Select-Object Office,Name,"Days To Expiration","Password Expiration Date" | Sort Office,Name | `
				ConvertTo-HTML -Head $Header -PreContent $Pre -PostContent $Post | `
				Set-CellColor -Property "Days To Expiration" -Color Yellow -Filter """Days To Expiration"" -le 5 -and ""Days To Expiration"" -ge 2" | `
				Set-CellColor -Property "Days To Expiration" -Color Red -Filter """Days To Expiration"" -le 1 -and ""Days To Expiration"" -ge 0" | `
				Set-AlternatingRows -CSSEvenClass even -CSSOddClass odd | Out-String
				$ITTo = "Site17 Help Desk"
				Send-MailMessage -From $From -To $ITTo -Subject $ReportSubject -Body $Site17Body -BodyAsHtml -SmtpServer $smtpServer
			; break}
		Site18{Write-Host "Processing local IT report for: Site18"
				$Site18Body += $Site18UserExpiringArray | Select-Object Office,Name,"Days To Expiration","Password Expiration Date" | Sort Office,Name | `
				ConvertTo-HTML -Head $Header -PreContent $Pre -PostContent $Post | `
				Set-CellColor -Property "Days To Expiration" -Color Yellow -Filter """Days To Expiration"" -le 5 -and ""Days To Expiration"" -ge 2" | `
				Set-CellColor -Property "Days To Expiration" -Color Red -Filter """Days To Expiration"" -le 1 -and ""Days To Expiration"" -ge 0" | `
				Set-AlternatingRows -CSSEvenClass even -CSSOddClass odd | Out-String
				$ITTo = "Site18 Help Desk"
				Send-MailMessage -From $From -To $ITTo -Subject $ReportSubject -Body $Site18Body -BodyAsHtml -SmtpServer $smtpServer
			; break}
		Site19{Write-Host "Processing local IT report for: Site19"
				$Site19Body += $Site19UserExpiringArray | Select-Object Office,Name,"Days To Expiration","Password Expiration Date" | Sort Office,Name | `
				ConvertTo-HTML -Head $Header -PreContent $Pre -PostContent $Post | `
				Set-CellColor -Property "Days To Expiration" -Color Yellow -Filter """Days To Expiration"" -le 5 -and ""Days To Expiration"" -ge 2" | `
				Set-CellColor -Property "Days To Expiration" -Color Red -Filter """Days To Expiration"" -le 1 -and ""Days To Expiration"" -ge 0" | `
				Set-AlternatingRows -CSSEvenClass even -CSSOddClass odd | Out-String
				$ITTo = "Site19 Help Desk"
				Send-MailMessage -From $From -To $ITTo -Subject $ReportSubject -Body $Site19Body -BodyAsHtml -SmtpServer $smtpServer
			; break}
		Site20{Write-Host "Processing local IT report for: Site20"
				$Site20Body += $Site20UserExpiringArray | Select-Object Office,Name,"Days To Expiration","Password Expiration Date" | Sort Office,Name | `
				ConvertTo-HTML -Head $Header -PreContent $Pre -PostContent $Post | `
				Set-CellColor -Property "Days To Expiration" -Color Yellow -Filter """Days To Expiration"" -le 5 -and ""Days To Expiration"" -ge 2" | `
				Set-CellColor -Property "Days To Expiration" -Color Red -Filter """Days To Expiration"" -le 1 -and ""Days To Expiration"" -ge 0" | `
				Set-AlternatingRows -CSSEvenClass even -CSSOddClass odd | Out-String
				$ITTo = "Site20 Help Desk"
				Send-MailMessage -From $From -To $ITTo -Subject $ReportSubject -Body $Site20Body -BodyAsHtml -SmtpServer $smtpServer
			; break}
	}#End Switch statement
}#End ForEach loop
#Send administrators a report of objects in each array
Send-AdminEmail $UserExpiredArray $UserNLOArray $UserExpiringArray
#EndRegion