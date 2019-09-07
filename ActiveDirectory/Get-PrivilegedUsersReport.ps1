<#
  This script will create a report of users that are members of the following
  privileged groups:
  - Enterprise Admins
  - Schema Admins
  - Domain Admins
  - Cert Publishers
  - Administrators
  - Account Operators
  - Server Operators
  - Backup Operators
  - Print Operators
  - DNS Admins

  A summary report is output to the console, whilst a full report is exported
  to a CSV file.

  The original script was written by Doug Symalla from Microsoft:
  - http://blogs.technet.com/b/askpfeplat/archive/2013/04/08/audit-membership-in-privileged-active-directory-groups-a-second-look.aspx
  - http://gallery.technet.microsoft.com/scriptcenter/List-Membership-In-bff89703

  The script was okay, but needed some updates to be more accurate and
  bug free. As Doug had not updated it since 26th April 2013, I though
  that I would. The changes I made are:

  1. Addressed a bug with the member count in the main section.
     Changed...
       $numberofUnique = $uniqueMembers.count
     To...
       $numberofUnique = ($uniqueMembers | measure-object).count
  2. Addressed a bug with the $colOfMembersExpanded variable in the
     getMemberExpanded function 
     Added...
       $colOfMembersExpanded=@()
  3. Enhanced the main section
  4. Enhanced the getForestPrivGroups function
  5. Enhanced the getUserAccountAttribs function
  6. Added script variables
  7. Added the accountExpires and info attributes
  8. Enhanced description of object members (AKA csv headers) so that
     it's easier to read.
  9. Added output to Excel file, included built-in groups 'Replicators', 'Remote Desktop Users', 'Network Configuration Operators', 'DnsAdmins'
  10. Added functionality to export to Excel in a cleaner more readable format.

  Script Name: Get-PrivilegedUsersReport.ps1
  Release v5
  Modified by Jeremy@jhouseconsulting.com 13/06/2014
  Modified again by Heather Miller - hemiller@deloitte.com 09-February-2018
  Modified output results by Heather Miller - hemiller@deloitte.com 21-August-2018

#>
#-------------------------------------------------------------

#Region ExecutionPolicy
#Set Execution Policy
#Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Confirm:$false -Force
#EndRegion

#Region Modules
#Check if required module is loaded, if not load import it
Try 
{
	Import-Module ActiveDirectory -ErrorAction Stop
}
Catch
{
	Throw "Active Directory module could not be loaded. $($_.Exception.Message)"
}

Try
{
	Import-Module ImportExcel -ErrorAction Stop
}
Catch
{
	Throw "Import Excel module could not be loaded. $($_.Exception.Message)"
}

#EndRegion

#Region Variables
$ADForestProperties = @("Domains", "ForestMode", "Name", "RootDomain")
$ADForest = Get-ADForest -ErrorAction SilentlyContinue | Select-Object -Property $ForestProperties
$forestName = ($ADForest).Name
$Domains = ($ADForest).Domains
[PSObject[]]$privGroupsInfo = @()
$GroupProps = @("distinguishedName", "Description", "groupCategory", "groupScope", "groupType", "managedBy", "member", "Name", "owner", "samAccountName", "whenChanged", "whenCreated")
$rootDNC = (Get-ADRootDSE).defaultNamingContext
$rootNS = 'root\CIMv2'
$rptFolder = 'E:\Reports'
$thisDomainName = ([System.DirectoryServices.ActiveDirectory.Domain]::GetCurrentDomain()).Name
# Set this to maximum number of unique members threshold
$MaxUniqueMembers = 100
# Set this to maximum password age threshold
$MaxPasswordAge = 365
# Set this to true to privide a detailed output to the console
[bool]$DetailedConsoleOutput = $True
$dtmFormatString = "yyyy-MM-dd HH:mm:ss"
$Limit = (Get-Date).AddDays(-60)
$moveLimit = (Get-Date).AddDays(-30)
Add-Type -AssemblyName "System.IO.Compression.FileSystem"
#EndRegion

#Region Functions
Function Check-Path {#Begin function to check path variable and return results
 	[CmdletBinding()]
    Param
    (
        [Parameter(Mandatory,Position=0)]
        [String]$Path,
        [Parameter(Mandatory,Position=1)]
        $PathType
    )
    
    Switch ( $PathType )
    {
    	File
			{
				If ( ( Test-Path -Path $Path -PathType Leaf ) -eq $true )
				{
					Write-Host "File: $Path already exists..." -BackgroundColor Black -ForegroundColor Green
				}
				Else
				{
					New-Item -Path $Path -ItemType File -Force
					Write-Host "File: $Path not present, creating new file..." -BackgroundColor Black -ForegroundColor Yellow
				}
			}
		Folder
			{
				If ( ( Test-Path -Path $Path -PathType Container ) -eq $true )
				{
					Write-Host "Folder: $Path already exists..." -BackgroundColor Black -ForegroundColor Green
				}
				Else
				{
					New-Item -Path $Path -ItemType Directory -Force
					Write-Host "Folder: $Path not present, creating new folder" -Background Black -ForegroundColor Yellow
				}
			}
	}
}#end function Check-Path

Function Get-MyInvocation {#Begin function to get $MyInvocation information
    Return $MyInvocation
}#End function Get-MyInvocation

Function Get-MemberExpanded {##################   Function to Expand Group Membership ################

        Param ($dn)

        $colOfMembersExpanded=@()
        $adobject = [adsi]"LDAP://$dn"
        $colMembers = $adobject.properties.item("member")
        ForEach ($objMember in $colMembers)
        {
			$objMembermod = $objMember.replace("/","\/")
			$objAD = [adsi]"LDAP://$objmembermod"
			$attObjClass = $objAD.properties.item("objectClass")
			If ($attObjClass -eq "group")
			{
				Get-MemberExpanded $objMember           
			}   
			Else
			{
			$colOfMembersExpanded += $objMember
			}
        }    
$colOfMembersExpanded 
}    

Function Get-ReportDate {#Begin function get report execution date
	Get-Date -Format "yyyy-MM-dd"
}#End function Get-ReportDate

Function Utc-Now {#Begin function to get date and time in UTC format
	[System.DateTime]::UtcNow
}#End function Utc-Now

Function Get-SmtpServer {#Begin function to get SMTP server for AD forest
	[CmdletBinding()]
    Param
    (
        [Parameter(Mandatory = $true)]
        [string]$Domain
    )
	
	Begin {}
	Process {
		Switch -Wildcard ($Domain){
			'*dmz.dtt' {$smtpServer = "appmail.dmz.dtt"}
			'dmz.fa' {$smtpServer = "10.246.65.208"}
			'*fantasia.qa' {$smtpServer = "appmail.ame.fantasia.qa"}
			'*dttplatform.dev' {$smtpServer = "dttdevmail.us.dttplatform.dev"}
			'us.ead.dev' {$smtpServer = "appmail.us.ead.dev"}
			'ead.dev' {$smtpServer = "dttdevmail.us.dttplatform.dev"}
			
			default {$smtpserver = "appmail.atrema.deloitte.com"}
		}
	}
	End {
		$out = [PSCustomObject] @{
			SmtpServer = $smtpServer
			Port = '25'
		}
		Return $out
	}
}#end function Get-SmtpServer

Function Send-SmtpRelayMessage {
	[CmdletBinding()]
	Param
	(
		[Parameter(Mandatory = $true)]
		$To,

		[Parameter(Mandatory = $false)]
		$Cc,

		[Parameter(Mandatory = $true)]
		$From,
		
		[Parameter(Mandatory = $false)]
		$ReplyTo,

		[Parameter(Mandatory = $true)]
		$Subject,

		[Parameter(Mandatory = $true)]
		$Body,

		[Parameter(Mandatory = $true)]
		$SMTPServer,

		[Parameter(Mandatory = $false)]
		$Port = 25,

        [Parameter(Mandatory = $false)]
        $InlineImageAttachments,

        [Parameter(Mandatory = $false)]
        $Attachments
	)


    $objMailMessage = New-Object System.Net.Mail.MailMessage
    $objSmtpClient = New-Object System.Net.Mail.SmtpClient

    $objSmtpClient.host = $SMTPServer
    $objSmtpClient.Port = $port

    ForEach ( $recipient in $To ) { $objMailMessage.To.Add((New-Object System.Net.Mail.MailAddress($recipient))) }
    ForEach ( $recipient in $Cc ) { $objMailMessage.Cc.Add((New-Object System.Net.Mail.MailAddress($recipient))) }
    $objMailMessage.From = $From
    $objMailMessage.Sender = $From
    $objMailMessage.ReplyTo = $ReplyTo
    $objMailMessage.Subject = $Subject
    $objMailMessage.Body = $Body
    $objMailMessage.IsBodyHtml = $true

    ForEach ( $inlineAttachment in $InlineImageAttachments )
    {
        $objAttachment = New-Object System.Net.Mail.Attachment($inlineAttachment.FullName.ToString())
        $objAttachment.ContentDisposition.Inline = $true
        $objAttachment.ContentDisposition.DispositionType = [System.Net.Mime.DispositionTypeNames]::Inline
        $objAttachment.ContentType.MediaType = "image/" + $inlineAttachment.Extension.ToString().Substring(1)
        $objAttachment.ContentType.Name = $inlineAttachment.Name.ToString()
        $objAttachment.ContentId = $inlineAttachment.Name.ToString()
        $objMailMessage.Attachments.Add($objAttachment)
    }

    ForEach ( $attachment in $Attachments )
    {
        $objAttachment = New-Object System.Net.Mail.Attachment($attachment.FullName.ToString())
        $objAttachment.ContentDisposition.Inline = $false
        $objAttachment.ContentDisposition.DispositionType = [System.Net.Mime.DispositionTypeNames]::Attachment
        $objAttachment.ContentType.MediaType = [System.Net.Mime.MediaTypeNames+Text]::Plain
        $objMailMessage.Attachments.Add($objAttachment)
    }

    $objSmtpClient.Send($objMailMessage)

	$objAttachment.Dispose()
	$objMailMessage.Dispose()
	$objSmtpClient.Dispose()
}#End function Send-SmtpRelayMessage

Function Get-DomainDNfromFQDN {########################## Function to Generate Domain DN from FQDN ########
	
	Param ($domainFQDN)
	
	$colSplit = $domainFQDN.Split(".")
	$FQDNdepth = $colSplit.length
	$DomainDN = ""
	For ($i=0;$i -lt ($FQDNdepth);$i++)
	{
		If ($i -eq ($FQDNdepth - 1))
		{
			$Separator=""
		}
		Else
		{
			$Separator = ","
		}
		[string]$DomainDN += "DC=" + $colSplit[$i] + $Separator
	}

	$DomainDN
}#End function Get-DomainDNfromFQDN

Function Get-UserAccountAttribs	{########################### Function to Calculate Password Age ##############
	
	Param($objADUser,$parentGroup)
	
	$objADUser = $objADUser.replace("/","\/")
    $adsiEntry = New-Object directoryservices.directoryentry("LDAP://$objADUser")
    $adsiSearcher = New-Object directoryservices.directorysearcher($adsientry)
    $adsiSearcher.pagesize=1000
    $adsiSearcher.searchscope="base"
	$adsiSearcher.ServerTimeLimit = 600
    $colUsers=$adsiSearcher.findall()
	ForEach($objuser in $colUsers)
	{
		$dn = $objuser.properties.item("distinguishedname")
		$domain = $dn -Split "," | ? {$_ -like "DC=*"}
		$domain = $domain -join "." -replace ("DC=", "")

		$sam = $objuser.properties.item("samaccountname")
		
		$attObjClass = $objuser.properties.item("objectClass")
		If ($attObjClass -eq "user")
		{
			$description = $objuser.properties.item("description")[0]
			$name = $objuser.properties.item("name")[0]
			$notes = $objuser.properties.item("info")[0]
			$notes = $notes -replace "`r`n", "|"
	 		If (($objuser.properties.item("lastlogontimestamp") | Measure-Object).Count -gt 0) 
			{
				$lastlogontimestamp = $objuser.properties.item("lastlogontimestamp")[0]
				$lastLogon = [System.DateTime]::FromFileTime($lastlogontimestamp)
				$lastLogonInDays = ((Get-Date) - $lastLogon).Days
				If ($lastLogon -match "1/01/1601")
				{
					$lastLogon = "Never logged on before"
	 		    	$lastLogonInDays = "N/A"
	            }
	 		}
			Else
			{
				$lastLogon = "Never logged on before"
				$lastLogonInDays = "N/A"
	 		}
			
	 		$accountexpiration = $objuser.properties.item("accountexpires")[0]
	 		If (($accountexpiration -eq 0) -OR ($accountexpiration -gt [DateTime]::MaxValue.Ticks))
			{
				$accountexpires = "<Never>"
	 		}
			Else
			{
				$accountexpires = [datetime]::FromFileTime([int64]::parse($accountexpiration))
	 		}

	   		$pwdLastSet=$objuser.properties.item("pwdLastSet")
	 		If ($pwdLastSet -gt 0)
         	{
				$pwdLastSet = [datetime]::FromFileTime([int64]::parse($pwdLastSet))
				$PasswordAge = ((Get-Date) - $pwdLastSet).days
         	}
         	Else 
			{
				$PasswordAge = "<Not Set>"
			}
			
			$uac = $objuser.properties.item("useraccountcontrol")
			$uac = $uac.item(0)
			If (($uac -bor 0x0002) -eq $uac)
			{
				$disabled="TRUE"
			}
			Else
			{
				$disabled = "FALSE"
			}
			
			If (($uac -bor 0x10000) -eq $uac)
			{
				$passwordneverexpires="TRUE"
			}
			Else
			{
				$passwordNeverExpires = "FALSE"
			}
	     }
		
		$record = "" | Select-Object -Property samAccountName,DistinguishedName,Name,domain,MemberOf,PasswordAge,LastLogon,LastLogonInDays,Disabled,PasswordNeverExpires,AccountExpires,Description,Notes
		$record.SamAccountName = [string]$sam
		$record.DistinguishedName = [string]$dn
		$record.Name = [string]$name
		$record.Domain = [string]$domain
		$record.MemberOf = [string]$parentGroup
		$record.PasswordAge = $PasswordAge
		$record.LastLogon = $lastLogon
		$record.LastLogonInDays = $lastLogonInDays
		$record.Disabled = $disabled
		$record.PasswordNeverExpires = $passwordNeverExpires
		$record.AccountExpires = $accountexpires
		$record.Description = $description
		$record.Notes = $notes

	 }
	 
$record

}#End function Get-UserAccountAttribs

Function Get-ForestPrivGroups {####### Function to find all Privileged Groups in the Forest ##########
  
	# Privileged Group Membership for the following groups:
	# - Enterprise Admins - SID: S-1-5-21root domain-519
	# - Schema Admins - SID: S-1-5-21root domain-518
	# - Domain Admins - SID: S-1-5-21domain-512
	# - Cert Publishers - SID: S-1-5-21domain-517
	# - Administrators - SID: S-1-5-32-544
	# - Account Operators - SID: S-1-5-32-548
	# - Server Operators - SID: S-1-5-32-549
	# - Backup Operators - SID: S-1-5-32-551
	# - Print Operators - SID: S-1-5-32-550
	# Reference: http://support.microsoft.com/kb/243330

	$colOfDNs = @()
	$Forest = [System.DirectoryServices.ActiveDirectory.Forest]::GetCurrentForest()
	$rootDomain = [string]($Forest.RootDomain.Name)
	$forestDomains = $Forest.domains
	$colDomainNames = @()
	
	ForEach ($domain in $forestDomains)
	{
		$domainName = [string]($domain.name)
		$colDomainNames += $domainName
	}

	$ForestRootDN = Get-DomainDNfromFQDN $rootDomain
	$colDomainDNs = @()
	
	ForEach ($domainname in $colDomainNames)
	{
		$domainDN = Get-DomainDNfromFQDN $domainname
		$colDomainDNs += $domainDN	
	}

	$GC = $forest.FindGlobalCatalog()
	$adObject = [adsi]"GC://$ForestRootDN"
	$rootDomainSid = New-Object System.Security.Principal.SecurityIdentifier($adObject.objectSid[0], 0)
	$rootDomainSid = $rootDomainSid.ToString()
	$colDASids = @()
	
	ForEach ( $domainDN in $colDomainDNs )
	{
		$adObject = [adsi]"GC://$domainDN"
		$DomainSid = New-Object System.Security.Principal.SecurityIdentifier($adObject.objectSid[0], 0)
		$DomainSid = $DomainSid.ToString()
		$daSid = "$DomainSID-512"
		$colDASids += $daSid
		$cpSid = "$DomainSID-517"
		$colDASids += $cpSid
		$dnsAdmSID = (Get-ADGroup -Identity DnsAdmins -Server $domainName).SID
		$colDASids += $dnsAdmSID.Value
	}

	#$colPrivGroups = @("S-1-5-32-544";"S-1-5-32-548";"S-1-5-32-549";"S-1-5-32-551";"S-1-5-32-550";"S-1-5-32-552";"S-1-5-32-554";"S-1-5-32-555";"$rootDomainSid-519";"$rootDomainSid-518")
	$colPrivGroups = @("S-1-5-32-544";"S-1-5-32-548";"S-1-5-32-549";"S-1-5-32-551";"S-1-5-32-550";"S-1-5-32-552";"S-1-5-32-555";"$rootDomainSid-519";"$rootDomainSid-518")
	$colPrivGroups += $colDASids
	      
	$Searcher = $GC.GetDirectorySearcher()
	ForEach($privGroup in $colPrivGroups)
    {
		$Searcher.filter = "(objectSID=$privGroup)"
		$Searcher.ServerTimeLimit = 600
		$Results = $Searcher.FindAll()
		ForEach ($result in $Results)
		{
			$dn = $result.properties.distinguishedname
			$colOfDNs += $dn
		}
    }
	
$colOfDNs

}#End function Get-ForestPrivGroups
   
#EndRegion











#Region Script
$Error.Clear()
#Start Function timer, to display elapsed time for function. Uses System.Diagnostics.Stopwatch class - see here: https://msdn.microsoft.com/en-us/library/system.diagnostics.stopwatch(v=vs.110).aspx 
$stopWatch = [System.Diagnostics.Stopwatch]::StartNew()
$dtmScriptStartTimeUTC = Utc-Now
$transcriptFileName     = "{0}-{1}-Transcript.txt" -f $dtmScriptStartTimeUTC.ToString("yyyy-MM-dd_HH.mm.ss"), "Privileged-Users-Report"

$myInv = Get-MyInvocation
$scriptDir = $myInv.PSScriptRoot
$scriptName = $myInv.ScriptName

#Check required folders and files exist, create if needed
Check-Path -Path $rptFolder -PathType Folder
[String]$privUserRptFldr = "{0}\{1}" -f $rptFolder, "PrivilegedUserReports"
Check-Path -Path $privUserRptFldr -PathType Folder
[String]$workingDir = "{0}\{1}" -f $privUserRptFldr, "workingDir"
Check-Path -Path $workingDir -PathType Folder
[String]$archiveFolder = "{0}\{1}" -f $privUserRptFldr, "Archives"
Check-Path -Path $archiveFolder -PathType Folder
`
# Start transcript file
Start-Transcript ("{0}\{1}" -f  $workingDir, $transcriptFileName)
Write-Verbose -Message ("[{0} UTC] [SCRIPT] Beginning execution of script." -f $dtmScriptStartTimeUTC.ToString($dtmFormatString)) -Verbose
Write-Verbose -Message ("[{0} UTC] [SCRIPT] Script Name             :  {1}" -f $(Utc-Now).ToString($dtmFormatString), $scriptName) -Verbose
Write-Verbose -Message ("[{0} UTC] [SCRIPT] Script Directory path   :  {1}" -f $(Utc-Now).ToString($dtmFormatString), $scriptDir) -Verbose
Write-Verbose -Message ("[{0} UTC] [SCRIPT] Working Directory path  :  {1}" -f $(Utc-Now).ToString($dtmFormatString), $workingDir) -Verbose
Write-Verbose -Message ("[{0} UTC] [SCRIPT] Archive folder path     :  {1}" -f $(Utc-Now).ToString($dtmFormatString), $archiveFolder) -Verbose

$forestPrivGroups = Get-ForestPrivGroups
$colAllPrivUsers = @()

$rootdse = New-Object DirectoryServices.DirectoryEntry("LDAP://rootDse")

ForEach ($privGroup in $forestPrivGroups)
{
	Write-Host ""
	Write-Host "Enumerating $privGroup.." -ForegroundColor Yellow
	$uniqueMembers = @()
	$colOfMembersExpanded = @()
	$colOfUniqueMembers = @()
	$members = @()
	$members += Get-MemberExpanded $privGroup
	If ( $members.count -ge 1 )
	{
		$uniqueMembers = $members | Sort-Object -Unique
		$numberofUnique = ($uniqueMembers | Measure-Object).count
		ForEach ($uniqueMember in $uniqueMembers)
		{
			$objAttribs = Get-UserAccountAttribs $uniqueMember $privGroup
			$colOfuniqueMembers += $objAttribs      
		}
		$colAllPrivUsers += $colOfUniqueMembers
	}
	Else
	{
		$numberOfUnique = 0
	}
                
	If ($numberOfUnique -gt $MaxUniqueMembers)
	{
	    Write-Host "...$privGroup has $numberofUnique unique members" -ForegroundColor Red
	}
	Else
	{
		Write-Host "...$privGroup has $numberofUnique unique members" -ForegroundColor Green
	}

	$pwdNeverExpiresCount = 0
	$pwdAgeCount = 0

	ForEach($user in $colOfUniqueMembers)
	{
		$i = 0
		$userpwdAge = $user.pwdAge
		$userpwdNeverExpires = $user.pWDneverExpires
		$userSAM = $user.SAM
		If ($userpwdneverExpires -eq $True)
		{
			$pwdneverExpiresCount ++
			$i ++
			If ( [bool]$DetailedConsoleOutput -eq $true )
			{
				Write-Host "......$userSAM has a password age of $userpwdAge and the password is set to never expire" -ForegroundColor Green
			}
		}
		
		If ($userpwdAge -gt $MaxPasswordAge)
		{
			$pwdAgeCount ++
			If ($i -gt 0)
			{
				If ( [bool]$DetailedConsoleOutput -eq $true )
				{
					Write-Host "......$userSAM has a password age of $userpwdage days" -ForegroundColor Green
				}
			}
		}
	}

	If ($numberofUnique -gt 0)
	{
		Write-Host "......There are $pwdneverExpiresCount accounts that have the password is set to never expire." -ForegroundColor Green
		Write-Host "......There are $pwdAgeCount accounts that have a password age greater than $MaxPasswordAge days." -ForegroundColor Green
	}
}

If ([bool]$DetailedConsoleOutput -eq $true)
{
	Write-Host "`nComments:" -ForegroundColor Yellow
	Write-Host " - If a privileged group contains more than $MaxUniqueMembers unique members, it's highlighted in red." -ForegroundColor Yellow
	Write-Host " - The privileged user is listed if their password is set to never expire." -ForegroundColor Yellow
	Write-Host " - The privileged user is listed if their password age is greater than $MaxPasswordAge days." -ForegroundColor Yellow
	Write-Host " - Service accounts should not be privileged users in the domain." -ForegroundColor Yellow
}

$EAs = $colAllPrivUsers | Where {$_.memberOf -like "*CN=Enterprise Admins*"}
$DAs = $colAllPrivUsers | Where { ($_.memberOf -like "*Domain Admins*") -or ($_.memberOf -like "*CN=Admins, del dominio*") } 

#Save output

$csvOutfile = "{0}\{1}" -f $workingDir, "$($forestName)_Privileged_User_Report_for_$(Get-ReportDate).csv"
$colAllPrivUsers | Export-CSV -Path $csvOutFile -Append -Delimiter ';' -NoTypeInformation

# Remove the quotes
#(Get-Content -Path "$outFile") | ForEach-Object { $_ -replace '"',"" } | Out-File -FilePath "$outFile" -Force -Encoding ascii

#Create Excel file
[String]$wsName = "Privileged Users"
$xlOutFile = "{0}\{1}" -f $workingDir, "$($forestName)_Privileged_User_Report_for_$(Get-ReportDate).xlsx"



$pt = [Ordered]@{}

$pt."grps by mbrLoginName" =@{
SourceWorksheet = 'Privileged Users';
PivotRows = 'domain','memberOf','samAccountName'
PivotData = @{'samAccountName'='Count'}
}

$pt."grps by mbrName" =@{
SourceWorkSheet = 'Privileged Users';
PivotRows = 'domain','memberOf','Name'
PivotData = @{'samAccountName'='Count'}
}

$pt."user is membeOf"=@{
SourceWorkSheet = 'Privileged Users';
PivotRows = 'domain','samAccountName','memberOf'
PivotData = @{'memberOf'='Count'}
}


$ExcelParams = @{
        Path = $xlOutFile
	    StartRow = 1
	    StartColumn = 1
	    AutoSize = $true
	    AutoFilter = $true
	    BoldTopRow = $true
	    FreezeTopRow = $true
    }



#$colAllPrivUsers | Export-Excel @ExcelParams -WorkSheetName "Privileged Users" -PivotTableDefinition $pt

$colAllPrivUsers | Export-Excel @ExcelParams -WorksheetName "Privileged Users" -Title "Active Directory Privileged Users" -TitleSize 18 -TitleBold -TitleBackgroundColor LightBlue -TitleFillPattern Solid -PivotTableDefinition $pt
$EAs | Export-Excel -Path $xlOutFile -Title "Active Directory Enterprise Admins" -TitleSize 18 -TitleBold -TitleBackgroundColor LightBlue -TitleFillPattern Solid -WorksheetName "Enterprise Admins by Name" -HideSheet "Enterprise Admins by Name" -StartRow 1 -StartColumn 1 -AutoSize -AutoFilter -BoldTopRow -FreezeTopRow -IncludePivotTable -PivotRows Domain, memberOf, Name -PivotData @{memberOf='count'}
$DAs | Export-Excel -Path $xlOutFile -Title "Active Directory Domain Admins" -TitleSize 18 -TitleBold -TitleBackgroundColor LightBlue -TitleFillPattern Solid -WorksheetName "Domain Admins by Name" -HideSheet "Domain Admins by Name" -StartRow 1 -StartColumn 1 -AutoSize -AutoFilter -BoldTopRow -FreezeTopRow -IncludePivotTable -PivotRows Domain, memberOf, Name -PivotData @{memberOf='count'}


Stop-Transcript
Sleep -Seconds 10

#Save output
#Compress the report files into a single archive
$archiveFile = "{0}\{1}" -f $privUserRptFldr, "$($forestName)_archive_for_$(Get-ReportDate).zip"
$compressionLevel = [System.IO.Compression.CompressionLevel]::Optimal
If( ( Test-Path -Path $archiveFile -PathType Leaf ) -eq $true ) { Remove-Item -Path $archiveFile -Confirm:$false }
#See https://msdn.microsoft.com/en-us/library/hh875104(v=vs.110).aspx?cs-save-lang=1&cs-lang=vb#code-snippet-1
[IO.Compression.ZipFile]::CreateFromDirectory($workingDir, $archiveFile, $compressionLevel, $false)


#Stop the stopwatch	
$stopWatch.Stop()

#Send e-mail notification	
$emailTemplatePath      = "{0}\{1}" -f $scriptDir, "EmailTemplates"
$imageAttachmentPath    = "{0}\{1}" -f $scriptDir, "Images"
$emailTemplateFileName = "IAM_Monthly_PrivilegedUserReport.html"
#$xmlConfigFile = "{0}\{1}" -f $scriptDir, "EmailSettings.xml"
#[xml]$objXmlConfig = Get-Content $xmlConfigFile
$smtpInfo = Get-SmtpServer -Domain ($ADForest).RootDomain

$runTime = $stopWatch.Elapsed.ToString('dd\.hh\:mm\:ss')
Write-Verbose -Message  ("[{0}] Sending email notification, please wait..." -f $(Utc-Now).ToString($dtmFormatString)) -Verbose
$emailTemplate = "{0}\{1}" -f $emailTemplatePath, $emailTemplateFileName

$htmlTemplate = [System.IO.StreamReader]$emailTemplate
$messageBody = $htmlTemplate.ReadToEnd()
$htmlTemplate.Dispose()

$messageBody = $messageBody.Replace("@@Date@@", $(Get-Date -Format MMM-dd-yyyy))
$messageBody = $messageBody.Replace("@@ForestName@@", $forestName)
$messageBody = $messageBody.Replace("@@ScheduledTaskName@@", "IAM.Monthly.PrivUsersReport")
$messageBody = $messageBody.Replace("@@ServerName@@", [System.Net.Dns]::GetHostByName("LocalHost").HostName)
$messageBody = $messageBody.Replace("@@ScriptName@@", $scriptName)
$messageBody = $messageBody.Replace("@@ScriptRunTime@@", $runTime)
$messageBody = $messageBody.Replace("@@CopyrightYear@@", $(Get-Date -Format yyyy))

$colInlineImageAttachments = Get-ChildItem -Path $imageAttachmentPath
$colAttachments = @(Get-ChildItem -Path $workingDir -Filter *.txt)
$colAttachments += Get-ChildItem -Path $archiveFile -File


$params = @{
	#To = "hemiller@deloitte.com"
	To = "gtsctoiaminfrastructureteam@deloitte.com"
	CC = "dbreeze@deloitte.com"
	From = "IAM-Monthly-PrivUserRptNotifications@deloitte.com"
	ReplyTo = "IAM-Monthly-PrivUserRptNotifications@deloitte.com"
	#SMTPServer = $objXmlConfig.Configuration.EmailSettings.SMTPServer
	SMTPServer = $smtpInfo.smtpServer
	#Port = $objXmlConfig.Configuration.EmailSettings.Port
	Port = $smtpInfo.Port
	InlineImageAttachments = $colInlineImageAttachments
	Attachments = $colAttachments
}


If ( ( Test-Path -Path $archiveFile -PathType Leaf ) -eq $true )
{
	$messageBody = $messageBody.Replace("@@ScriptStatus@@", "Success")
	$params.Subject = "SUCCESS: $($forestName)_Privileged_Users_Report_as_of_$(Get-ReportDate)"
	$params.Body = $messageBody
}
Else
{
	$messageBody = $messageBody.Replace("@@ScriptStatus@@", "Failed")
	$params.Subject = "FAILED: $($forestName)_Privileged_Users_Report_as_of_$(Get-ReportDate)"
	$params.Body = $messageBody
}

Send-SmtpRelayMessage @params
#Send-MailMessage @params
Get-ChildItem -Path $privUserRptFldr | Where-Object { ($_.PsIsContainer) -and ($_.Name -like "workingDir*") } | Remove-Item -Recurse -Force -Confirm:$false

#Clean up archive files
$filesToMove = Get-ChildItem -Path $privUserRptFldr | Where-Object { !$_.PSIsContainer -and $_.CreationTime -lt $moveLimit }
$filesToMove | ForEach-Object { Move-Item -Path $_.FullName -Destination $archiveFolder -Force }

Get-ChildItem -Path $archiveFolder -Recurse | Where-Object { $_.LastWriteTime -lt $Limit } | Remove-Item -Force

#Close out script
$dtmScriptStopTimeUTC = Utc-Now
$elapsedTime = New-TimeSpan -Start $dtmScriptStartTimeUTC -End $dtmScriptStopTimeUTC
Write-Verbose -Message ("[{0} UTC] [SCRIPT] Script Complete" -f $(Utc-Now).ToString($dtmFormatString)) -Verbose
Write-Verbose -Message ("[{0} UTC] [SCRIPT] Script Start Time :  {1}" -f $(Utc-Now).ToString($dtmFormatString), $dtmScriptStartTimeUTC.ToString($dtmFormatString)) -Verbose
Write-Verbose -Message ("[{0} UTC] [SCRIPT] Script Stop Time  :  {1}" -f $(Utc-Now).ToString($dtmFormatString), $dtmScriptStopTimeUTC.ToString($dtmFormatString)) -Verbose
Write-Verbose -Message ("[{0} UTC] [SCRIPT] Elapsed Time: {1:N0}.{2:N0}:{3:N0}:{4:N1}  (Days.Hours:Minutes:Seconds)" -f $(Utc-Now).ToString($dtmFormatString), $elapsedTime.Days, $elapsedTime.Hours, $elapsedTime.Minutes, $elapsedTime.Seconds) -Verbose
#EndRegion