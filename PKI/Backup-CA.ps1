#Requires -Modules ADCSAdministration, Microsoft.PowerShell.Security, PKI
#Requires -Version 2.0

#Region Help
<#

.NOTES Needs to be run from Certificate Authority server

.SYNOPSIS Backs up CertIFicate Authority Issuing CA database

.DESCRIPTION This script utilizes native PowerShell cmdlets in Windows
Server 2012 R2 to perform daily backups for the CA database locally. After
those backups are completed, the EDC team backs up the static files to tape

.OUTPUTS Log file with System.Exception error that gets e-mailed to
Messaging Team

.EXAMPLE 
	PS> Backup-CA.ps1


#>
###########################################################################
#
#
# AUTHOR:  Heather Miller, Sr. AD/Messaging Engineer
#          IT Operations Group
#
# VERSION HISTORY:
# Get-Date -Format MM-dd-YYYY - Version .1
#
# Usage: .\Backup-CA002.ps1
# 
###########################################################################
#EndRegion

#Region ExecutionPolicy
#Set Execution Policy for Powershell
Set-ExecutionPolicy RemoteSigned
#EndRegion

#Region Modules
#Check IF required module is loaded, IF not load import it
IF (-not(Get-Module ADCSAdministration))
{
	Import-Module ADCSAdministration
}
If (-not(Get-Module PKI))
{
    Import-Module -Name PKI
}
#IF (-not(Get-Module WindowsServerBackup))
#{
#	Import-Module WindowsServerBackup
#}
IF (-not(Get-Module Microsoft.PowerShell.Security))
{
	Import-Module Microsoft.PowerShell.Security
}
#EndRegion

#Region Variables
#Dim variables
$Limit = (Get-Date).AddDays(-1)
$myScriptName = $MyInvocation.MyCommand.Name
$evtProps = @("Index", "TimeWritten", "EntryType","Source", "InstanceID", "Message")
$LogonServer = $env:LOGONSERVER
#$RetentionLimit = (Get-Date).AddDays(-3)
[int]$RetentionLimit = 70
$ServerName = $env:COMPUTERNAME
$ADForest = [System.DirectoryServices.ActiveDirectory.Forest]::GetCurrentForest()
$forestName = ($ADForest).Name
#EndRegion

#Region Functions
Function fnExport-CertificateTemplate {
<#
.Synopsis
    Exports certificate templates to a serialized format.
.Description
    Exports certificate templates to a serialized format. Exported templates can be distributed
    and imported in another forest.
.Parameter Template
    A collection of certificate templates to export. A collection can be retrieved by running
    Get-CertificateTemplate that is a part of PSPKI module: https://pspki.codeplex.com
.Parameter Path
    Specifies the path to export.
.Example
    $Templates = Get-CertificateTemplate -Name SmartCardV2, WebServerV3
    PS C:\> Export-CertificateTemplate $templates c:\temp\templates.dat
#>
[CmdletBinding()]
    Param(
        [Parameter(Mandatory = $true)]
        [PKI.CertificateTemplates.CertificateTemplate[]]$Template,
        [Parameter(Mandatory = $true)]
        [IO.FileInfo]$Path
    )
    If ($Template.Count -lt 1) {Throw "At least one template must be specified in the 'Template' parameter."}
    $ErrorActionPreference = "Stop"

#Region enums
    $HashAlgorithmGroup = 1
    $EncryptionAlgorithmGroup = 2
    $PublicKeyIdGroup = 3
    $SigningAlgorithmIdGroup = 4
    $RDNIdGroup = 5
    $ExtensionAttributeGroup = 6
    $EKUGroup = 7
    $CertificatePolicyGroup = 8
    $EnrollmentObjectGroup = 9
#EndRegion

#Region Inner Functions
    Function Get-OIDid ($OID,$group) {
        $found = $false
        :outer For ($i = 0; $i -lt $oids.Count; $i++) {
            If ($script:oids[$i].Value -eq $OID.Value) {
                $ID = ++$i
                $found = $true
                Break outer
            }
        }
        If (!$found) {
            $script:oids += New-Object psobject -Property @{
                Value = $OID.Value;
                Group = $group;
                Name = $OID.FriendlyName;
            }
            $ID = $script:oids.Count
        }
        $ID
    }
    Function Get-Seconds ($str) {
        [void]("$str" -match "(\d+)\s(\w+)")
        $period = $matches[1] -as [int]
        $units = $matches[2]
        Switch ($units) {
            "hours" {$period * 3600}
            "days" {$period * 3600 * 24}
            "weeks" {$period * 3600 * 168}
            "months" {$period * 3600 * 720}
            "years" {$period * 3600 * 8760}
        }
    }
#EndRegion

    $SB = New-Object Text.StringBuilder
    [void]$SB.Append(
@"
<GetPoliciesResponse xmlns="http://schemas.microsoft.com/windows/pki/2009/01/enrollmentpolicy">
    <response>
        <policyID/>
        <policyFriendlyName/>
        <nextUpdateHours>8</nextUpdateHours>
        <policiesNotChanged a:nil="true" xmlns:a="http://www.w3.org/2001/XMLSchema-instance"/>
        <policies>
"@)
    $script:oids = @()
    ForEach ($temp in $Template) {
         [void]$SB.Append("<policy>")
        $OID = New-Object Security.Cryptography.Oid $temp.OID.Value, $temp.DisplayName
        $tempID = Get-OIDid $OID $EnrollmentObjectGroup
        # validity/renewal
        $validity = Get-Seconds $temp.Settings.ValidityPeriod
        $renewal = Get-Seconds $temp.Settings.RenewalPeriod
        # key usages
        $KU = If ([int]$temp.Settings.Cryptography.CNGKeyUsage -eq 0) {
            '<keyUsageProperty xmlns:a="http://www.w3.org/2001/XMLSchema-instance" a:nil="true"/>'
        } Else {
            "<keyUsageProperty>$([int]$temp.Settings.CNGKeyUsage)</keyUsageProperty>"
        }
        # private key security
        $PKS = If ([string]::IsNullOrEmpty($temp.Settings.Cryptography.PrivateKeySecuritySDDL)) {
            '<permissions xmlns:a="http://www.w3.org/2001/XMLSchema-instance" a:nil="true"/>'
        } Else {
            "<permissions>$($temp.Settings.PrivateKeySecuritySDDL)</permissions>"
        }
        # public key algorithm
        $KeyAlgorithm = If ($temp.Settings.Cryptography.KeyAlgorithm.Value -eq "1.2.840.113549.1.1.1") {
            '<algorithmOIDReference xmlns:a="http://www.w3.org/2001/XMLSchema-instance" a:nil="true"/>'
        } Else {
            $kalgID = Get-OIDid $temp.Settings.Cryptography.KeyAlgorithm $PublicKeyIdGroup
            "<algorithmOIDReference>$kalgID</algorithmOIDReference>"
        }
        # superseded templates
        $superseded = If ($temp.Settings.SupersededTemplates -eq 0) {
            '<supersededPolicies xmlns:a="http://www.w3.org/2001/XMLSchema-instance" a:nil="true"/>'    
        } Else {
            $str = "<supersededPolicies>"
            $temp.Settings.SupersededTemplates | ForEach-Object {$str += "<commonName>$_</commonName>"}
            $str + "</supersededPolicies>"
        }
        # list of CSPs
        $CSPs = If ($temp.Settings.Cryptography.CSPList.Count -eq 0) {
            '<cryptoProviders xmlns:a="http://www.w3.org/2001/XMLSchema-instance" a:nil="true"/>'
        } Else {
            $str = "<cryptoProviders>`n"
            $temp.Settings.Cryptography.CSPList | ForEach-Object {
                $str += "<provider>$_</provider>`n"
            }
            $str + "</cryptoProviders>"
        }
        # version
        [void]($temp.Version -match "(\d+)\.(\d+)")
        $major = $matches[1]
        $minor = $matches[2]
        # hash algorithm
        $hash = If ($temp.Settings.Cryptography.HashAlgorithm.Value -eq "1.3.14.3.2.26") {
            '<hashAlgorithmOIDReference xmlns:a="http://www.w3.org/2001/XMLSchema-instance" a:nil="true"/>'
        } Else {
            $hashID = Get-OIDid $temp.Settings.Cryptography.HashAlgorithm $HashAlgorithmGroup
            "<hashAlgorithmOIDReference>$hashID</hashAlgorithmOIDReference>"
        }
        # enrollment agent
        $RAR = If ($temp.Settings.RegistrationAuthority.SignatureCount -eq 0) {
            '<rARequirements xmlns:a="http://www.w3.org/2001/XMLSchema-instance" a:nil="true"/>'
        } Else {
            $str = @"
<rARequirements>
<rASignatures>$($temp.Settings.RegistrationAuthority.SignatureCount)</rASignatures>
"@
            If ([string]::IsNullOrEmpty($temp.Settings.RegistrationAuthority.ApplicationPolicy.Value)) {
                $str += '<rAEKUs xmlns:a="http://www.w3.org/2001/XMLSchema-instance" a:nil="true"/>'
            } Else {
                $raapID = Get-OIDid $temp.Settings.RegistrationAuthority.ApplicationPolicy $EKUGroup
                $str += @"
<rAEKUs>
    <oIDReference>$raapID</oIDReference>
</rAEKUs>
"@
            }
            If ($temp.Settings.RegistrationAuthority.CertificatePolicies.Count -eq 0) {
                $str += '<rAPolicies xmlns:a="http://www.w3.org/2001/XMLSchema-instance" a:nil="true"/>'
            } Else {
                $str += "                       <rAPolicies>"
                $temp.Settings.RegistrationAuthority.CertificatePolicies | ForEach-Object {
                    $raipID = Get-OIDid $_ $CertificatePolicyGroup
                    $str += "<oIDReference>$raipID</oIDReference>`n"
                }
                $str += "</rAPolicies>`n"
            }
            $str += "</rARequirements>`n"
            $str
        }
        # key archival
        $KAS = If (!$temp.Settings.KeyArchivalSettings.KeyArchival) {
            '<keyArchivalAttributes xmlns:a="http://www.w3.org/2001/XMLSchema-instance" a:nil="true"/>'
        } Else {
            $kasID = Get-OIDid $temp.Settings.KeyArchivalSettings.EncryptionAlgorithm $EncryptionAlgorithmGroup
@"
<keyArchivalAttributes>
    <symmetricAlgorithmOIDReference>$kasID</symmetricAlgorithmOIDReference>
    <symmetricAlgorithmKeyLength>$($temp.Settings.KeyArchivalSettings.KeyLength)</symmetricAlgorithmKeyLength>
</keyArchivalAttributes>
"@
        }
        $sFlags = [Convert]::ToUInt32($("{0:x2}" -f [int]$temp.Settings.SubjectName),16)
        [void]$SB.Append(
@"
<policyOIDReference>$tempID</policyOIDReference>
<cAs>
    <cAReference>0</cAReference>
</cAs>
<attributes>
    <commonName>$($temp.Name)</commonName>
    <policySchema>$($temp.SchemaVersion)</policySchema>
    <certificateValidity>
        <validityPeriodSeconds>$validity</validityPeriodSeconds>
        <renewalPeriodSeconds>$renewal</renewalPeriodSeconds>
    </certificateValidity>
    <permission>
        <enroll>false</enroll>
        <autoEnroll>false</autoEnroll>
    </permission>
    <privateKeyAttributes>
        <minimalKeyLength>$($temp.Settings.Cryptography.MinimalKeyLength)</minimalKeyLength>
        <keySpec>$([int]$temp.Settings.Cryptography.KeySpec)</keySpec>
        $KU
        $PKS
        $KeyAlgorithm
        $CSPs
    </privateKeyAttributes>
    <revision>
        <majorRevision>$major</majorRevision>
        <minorRevision>$minor</minorRevision>
    </revision>
    $superseded
    <privateKeyFlags>$([int]$temp.Settings.Cryptography.PrivateKeyOptions)</privateKeyFlags>
    <subjectNameFlags>$sFlags</subjectNameFlags>
    <enrollmentFlags>$([int]$temp.Settings.EnrollmentOptions)</enrollmentFlags>
    <generalFlags>$([int]$temp.Settings.GeneralFlags)</generalFlags>
    $hash
    $rar
    $KAS
<extensions>
"@)
        ForEach ($ext in $temp.Settings.Extensions) {
            $extID = Get-OIDid ($ext.Oid) $ExtensionAttributeGroup
            $critical = $ext.Critical.ToString().ToLower()
            $value = [Convert]::ToBase64String($ext.RawData)
            [void]$SB.Append("<extension><oIDReference>$extID</oIDReference><critical>$critical</critical><value>$value</value></extension>")
        }
        [void]$SB.Append("</extensions></attributes></policy>")
    }
    [void]$SB.Append("</policies></response>")
    [void]$SB.Append("<oIDs>")
    $n = 1
    $script:oids | ForEach-Object {
        [void]$SB.Append(@"
<oID>
    <value>$($_.Value)</value>
    <group>$($_.Group)</group>
    <oIDReferenceID>$n</oIDReferenceID>
    <defaultName>$($_.Name)</defaultName>
</oID>
"@)
        $n++
    }
    [void]$SB.Append("</oIDs></GetPoliciesResponse>")
    Set-Content -Path $Path -Value $SB.ToString() -Encoding Ascii
}

Function fnGeneratePassword() {
    Param (
    [int]$length = 25,    
    [bool] $includeLowercaseLetters = $true,
    [bool] $includeUppercaseLetters = $true,
    [bool] $includeNumbers = $true,
    [bool] $includeSpecialChars = $true,
    [bool] $noSimilarCharacters = $true
    )
 
    <#
    (c) Morgan de Jonge CC BY SA
    Generates a random password. you're able to specify:
    - The desired password length (minimum = 4)
    - Whether or not to use lowercase characters
    - Whether or not to use uppercase characters
    - Whether or not to use numbers
    - Whether or not to use special characters
    - Whether or not to avoid using similar characters ( e.g. i, l, o, 1, 0, I)
    #>
 
    # Validate params
    If($length -lt 4) {
        $exception = New-Object Exception "The minimum password length is 4"
        Throw $exception
    }
    If ($includeLowercaseLetters -eq $false -and 
            $includeUppercaseLetters -eq $false -and
            $includeNumbers -eq $false -and
            $includeSpecialChars -eq $false) {
        $exception = New-Object Exception "At least one set of included characters must be specified"
        Throw $exception
    }
 
    #Available characters
    $CharsToSkip = [char]"i", [char]"l", [char]"o", [char]"1", [char]"0", [char]"I"
    $AvailableCharsForPassword = $null;
    $uppercaseChars = $null 
    For($a = 65; $a -le 90; $a++) { if($noSimilarCharacters -eq $false -or [char][byte]$a -notin $CharsToSkip) {$uppercaseChars += ,[char][byte]$a }}
    $lowercaseChars = $null
    For($a = 97; $a -le 122; $a++) { if($noSimilarCharacters -eq $false -or [char][byte]$a -notin $CharsToSkip) {$lowercaseChars += ,[char][byte]$a }}
    $digitChars = $null
    For($a = 48; $a -le 57; $a++) { if($noSimilarCharacters -eq $false -or [char][byte]$a -notin $CharsToSkip) {$digitChars += ,[char][byte]$a }}
    $specialChars = $null
    $specialChars += [char]"=", [char]"+", [char]"_", [char]"?", [char]"!", [char]"-", [char]"#", [char]"$", [char]"*", [char]"&", [char]"@"
 
    $TemplateLetters = $null
    If($includeLowercaseLetters) { $TemplateLetters += "L" }
    If($includeUppercaseLetters) { $TemplateLetters += "U" }
    If($includeNumbers) { $TemplateLetters += "N" }
    If($includeSpecialChars) { $TemplateLetters += "S" }
    $PasswordTemplate = @()
    # Set password template, to ensure that required chars are included
    Do {   
        $PasswordTemplate.Clear()
        for($loop = 1; $loop -le $length; $loop++) {
            $PasswordTemplate += $TemplateLetters.Substring((Get-Random -Maximum $TemplateLetters.Length),1)
        }
    }
    While ((
        (($includeLowercaseLetters -eq $false) -or ($PasswordTemplate -contains "L")) -and
        (($includeUppercaseLetters -eq $false) -or ($PasswordTemplate -contains "U")) -and
        (($includeNumbers -eq $false) -or ($PasswordTemplate -contains "N")) -and
        (($includeSpecialChars -eq $false) -or ($PasswordTemplate -contains "S"))) -eq $false
    )
    #$PasswordTemplate now contains an array with at least one of each included character type (uppercase, lowercase, number and/or special)
 
    ForEach($char in $PasswordTemplate) {
        Switch ($char) {
            L { $Password += $lowercaseChars | Get-Random }
            U { $Password += $uppercaseChars | Get-Random }
            N { $Password += $digitChars | Get-Random }
            S { $Password += $specialChars | Get-Random }
        }
    }
    Return $Password
}

Function fnGet-Date {#Begin function to get short date
	Get-Date -Format "MM-dd-yyyy"
}#End function fnGet-Date

Function fnGet-TodaysDate {#Begin function to get today's date
	Get-Date
}#End function fnGet-TodaysDate

Function fnGet-LongDate {#Begin function to get date and time in long format
	Get-Date -Format G
}#End function fnGet-LongDate

Function fnGet-ReportDate {#Begin function set report date format
	Get-Date -Format "yyyy-MM-dd"
}#End function fnGet-ReportDate

Function fnSend-AdminEmail {#Begin function to send summary e-mail to Administrators
	Param ($BodyMessage1, $BodyMessage2, $BodyMessage3, $BodyMessage4, $BodyMessage5, $BodyMessage6, $BodyMessage7, $BodyMessage8, $BodyMessage9, $BodyMessage10)
	
	#Dim function specific variables
	$From = 'no-reply@domain.com'
	$To = "user@domain.com"
	$Body = @"
		<p>$BodyMessage1</p>
		
		<p>$BodyMessage3</p>

		<p>$BodyMessage4</p>
		
		<p>$BodyMessage5</p>
		
		<p>$BodyMessage6</p>
		
		<p>$BodyMessage7</p>

		<p>$BodyMessage8</p>
		
		<p>$BodyMessage9</p>
		
		<p>$BodyMessage10</p>
		
		<p>$BodyMessage2</p>
"@
	$ReportSubject="Active Directory Certificate Authority Backup status for $(fnGet-Date)"
	Switch($forestName){
	"domain1.com" {$smtpServer = "appmail.domain1.com"}
	"domain2.com" {$smtpServer = "appmail.domain2.com"}
	
	default {$smtpserver = "appmail.domain.com"}
	}
	Send-MailMessage -From $From -To $To -Subject $ReportSubject -Body $Body -BodyAsHTML -SmtpServer $smtpServer
}#End function fnSend-AdminEmail
#EndRegion








#Region Script
#Begin Script
$BodyMessage1 = "Script name: $myScriptName started at $(fnGet-LongDate)."

#Region Check folder structures
	$Disk = Get-WmiObject -Class Win32_LogicalDisk -Namespace root\CIMv2 | Where-Object {$_.DeviceID -eq 'D:'} | Select-Object -Property DeviceID
	$LogicalDisk = ($Disk).DeviceID
	$volBkpFldr = $LogicalDisk + "\CABackup\"
	$TodaysFldr = $volBkpFldr + $((Get-Date).ToString('yyyy-MM-dd'))
#EndRegion

#Region Backup-CA-DB-Key
	#Backup Certificate Authority Database and Private Key
	If ((Test-Path -Path $TodaysFldr -PathType Container) -eq $true)
	{
		[String]$pw = fnGeneratePassword -length 25 -includeLowercaseLetters $true -includeUppercaseLetters $true -includeNumbers $true -includeSpecialChars $true
		$BkpPassword = ConvertTo-SecureString -String $pw -AsPlainText -Force

		Backup-CARoleService -Path $TodaysFldr -Password $BkpPassword
		$CAEvents = Get-EventLog -LogName Application -Newest 1 | Where-Object {$_.Source -eq "ESENT" -and $_.Message -like "*certsrv.exe*"} | `
		Select-Object -Property TimeWritten, EntryType, Source, InstanceID, Message | Sort-Object -Property TimeWritten
		If (($CAEvents).InstanceID -eq 213)
		{
			$DBFilePath = $TodaysFldr + "\Database"
			$DBFileInfo = Get-ChildItem -Path $DBFilePath -Recurse -File | Select-Object -Property @{Name="Name"; Expression={$_.Name.ToUpper()}}, CreationTime, @{Name="Kilobytes";Expression={[math]::Round($_.Length / 1Kb,2)}}
			$EventInfo = "Event Time: " + ($CAEvents).TimeWritten + " Entry Type: " + ($CAEvents).EntryType + " Event ID: " + ($CAEvents).InstanceID + " Event Details: " + ($CAEvents).Message
			$BodyMessage3 = "Certificate Authority Backup Results:<br/>"
			$BodyMessage3 += "Event Log Results:<br />"
			$BodyMessage3 += $EventInfo  + "<br />"
			$BodyMessage3 += "Number of files backed up:<br />"
			$BodyMessage3 += ($DBFileInfo).Count
			$BodyMessage3 += $DBFileInfo + "<br />"
			$BodyMessage2 = $pw
		}
		ElseIf (($CAEvents).EntryType -eq "Error")
		{
			$DBFilePath = $TodaysFldr + "\Database"
			$DBFileInfo = Get-ChildItem -Path $DBFilePath -Recurse -File | Select-Object -Property @{Name="Name"; Expression={$_.Name.ToUpper()}}, CreationTime, @{Name="Kilobytes";Expression={[math]::Round($_.Length / 1Kb,2)}}
			$EventInfo = "Event Time: " + ($CAEvents).TimeWritten + " Entry Type: " + ($CAEvents).EntryType + " Event ID: " + ($CAEvents).InstanceID + " Event Details: " + ($CAEvents).Message
			$BodyMessage3 = "Certificate Authority Backup Results:<br/>"
			$BodyMessage3 += "Event Log Results:<br />"
			$BodyMessage3 += $EventInfo  + "<br />"
			$BodyMessage3 += "Number of files backed up:<br />"
			$BodyMessage3 += ($DBFileInfo).Count
			$BodyMessage3 += $DBFileInfo + "<br />"
		}
		Else
		{
			$BodyMessage3 = "Certificate Authority Backup Results:<br/>"
			$BodyMessage3 += "Event Log Results:<br />"
			$BodyMessage3 += "Certificate Authority Database and Private Key Backups failed $(fnGet-Date). Please investigate. <br />"
		}
	}
	Else
	{
		New-Item -ItemType Directory -Path $TodaysFldr -Force
		[String]$pw = fnGeneratePassword -length 25 -includeLowercaseLetters $true -includeUppercaseLetters $true -includeNumbers $true -includeSpecialChars $true
		$BkpPassword = ConvertTo-SecureString -String $pw -AsPlainText -Force
		Backup-CARoleService -Path $TodaysFldr -Password $Password
		$CAEvents = Get-EventLog -LogName Application -Newest 100 | Where-Object {$_.Source -eq "ESENT" -and $_.Message -like "*certsrv.exe*"} | `
		Select-Object -Property TimeWritten, EntryType, Source, InstanceID, Message | Sort-Object -Property TimeWritten
		If (($CAEvents).InstanceID -eq 213)
		{
			$DBFilePath = $TodaysFldr + "\Database"
			$DBFileInfo = Get-ChildItem -Path $DBFilePath -Recurse -File | Select-Object -Property @{Name="Name"; Expression={$_.Name.ToUpper()}}, CreationTime, @{Name="Kilobytes";Expression={[math]::Round($_.Length / 1Kb,2)}}
			$EventInfo = "Event Time: " + ($CAEvents).TimeWritten + " Entry Type: " + ($CAEvents).EntryType + " Event ID: " + ($CAEvents).InstanceID + " Event Details: " + ($CAEvents).Message
			$BodyMessage3 = "Certificate Authority Backup Results:<br/>"
			$BodyMessage3 += "Event Log Results:<br />"
			$BodyMessage3 += $EventInfo  + "<br />"
			$BodyMessage3 += "Number of files backed up:<br />"
			$BodyMessage3 += ($DBFileInfo).Count
			$BodyMessage3 += $DBFileInfo + "<br />"
			$BodyMessage2 = $pw
		}
		ElseIf (($CAEvents).EntryType -eq "Error")
		{
			$DBFilePath = $TodaysFldr + "\Database"
			$DBFileInfo = Get-ChildItem -Path $DBFilePath -Recurse -File | Select-Object -Property @{Name="Name"; Expression={$_.Name.ToUpper()}}, CreationTime, @{Name="Kilobytes";Expression={[math]::Round($_.Length / 1Kb,2)}}
			$EventInfo = "Event Time: " + ($CAEvents).TimeWritten + " Entry Type: " + ($CAEvents).EntryType + " Event ID: " + ($CAEvents).InstanceID + " Event Details: " + ($CAEvents).Message
			$BodyMessage3 = "Certificate Authority Backup Results:<br/>"
			$BodyMessage3 += "Event Log Results:<br />"
			$BodyMessage3 += $EventInfo  + "<br />"
			$BodyMessage3 += "Number of files backed up:<br />"
			$BodyMessage3 += ($DBFileInfo).Count
			$BodyMessage3 += $DBFileInfo + "<br />"
		}
		Else
		{
			$BodyMessage3 = "Certificate Authority Backup Results:<br/>"
			$BodyMessage3 += "Event Log Results:<br />"
			$BodyMessage3 += "Certificate Authority Database and Private Key Backups failed $(fnGet-Date). Please investigate. <br />"
		}
	}
#EndRegion

#Region Backup-CA-Registry
#Backup Certificate Authority Registry Hive
	$RegFldr = $TodaysFldr + "\Registry"

	IF ((Test-Path -Path $RegFldr -PathType Container) -eq $true)
	{
		reg.exe export HKLM\System\CurrentControlSet\Services\CertSvc "$RegFldr\CARegistry_$(fnGet-ReportDate).reg"
		$RegFile = "$RegFldr\CARegistry_$(fnGet-ReportDate).reg"
		If ((Test-Path -Path $RegFile -PathType Leaf) -eq $true)
		{
			$BodyMessage4 = "Registry Key Backup Results:<br />"
			$BodyMessage4 += "Certificate Services registry key: $RegFile export was successful.<br />"
		}
		Else
		{
			$BodyMessage4 = "Registry Key Backup Results:<br />"
			$BodyMessage4 += "Certificate Services registry key export failed on $(fnGet-ReportDate).<br />"
		}
	}
	Else
	{
		New-Item -ItemType Directory -Path $RegFldr -Force
		reg.exe export HKLM\System\CurrentControlSet\Services\CertSvc  "$RegFldr\CARegistry_$(fnGet-ReportDate).reg"
		$RegFile = "$RegFldr\CARegistry_$(fnGet-ReportDate).reg"
		If ((Test-Path -Path $RegFile -PathType Leaf) -eq $true)
		{
			$BodyMessage4 = "Registry Key Backup Results:<br />"
			$BodyMessage4 += "Certificate Services registry key: $RegFile export was successful.<br />"
		}
		Else
		{
			$BodyMessage4 = "Registry Key Backup Results:<br />"
			$BodyMessage4 += "Certificate Services registry key export failed on $(fnGet-ReportDate).<br />"
		}	
	}
#EndRegion

#Region Backup-Policy-File
#Backup Certificate Policy .Inf file
	$PolicyFldr = $TodaysFldr + "\PolicyFile\"
	$PolicyFile = $env:SystemRoot + "\CAPolicy.inf"
	IF ((Test-Path -Path $PolicyFldr -PathType Container) -eq $true)
	{
		Copy-Item -Path $PolicyFile -Destination $PolicyFldr
		If ((Test-Path -Path $PolicyFile -PathType Leaf) -eq $true)
		{
			$BodyMessage5 = "Certificate Authority Policy File Backup Results:<br />"
			$BodyMessage5 += "Backup copy of policy file: CAPolicy.inf was successful."
		}
		Else
		{
			$BodyMessage5 = "Certificate Authority Policy File Backup Results:<br />"
			$BodyMessage5 += "Backup copy of policy file: CAPolicy.inf failed. Please investigate."
		}
	}
	Else
	{
		New-Item -ItemType Directory -Path $PolicyFldr -Force
		Copy-Item -Path $PolicyFile -Destination $PolicyFldr
		If ((Test-Path -Path $PolicyFile -PathType Leaf) -eq $true)
		{
			$BodyMessage5 = "Certificate Authority Policy File Backup Results:<br />"
			$BodyMessage5 += "Backup copy of policy file: CAPolicy.inf was successful."
		}
		Else
		{
			$BodyMessage5 = "Certificate Authority Policy File Backup Results:<br />"
			$BodyMessage5 += "Backup copy of policy file: CAPolicy.inf failed. Please investigate."
		}
	}
#EndRegion

#Region Backup-IIS-Files
#Backup IIS Custom files
	$IISCustomsFldr = $TodaysFldr + "\IISCustomizations\"
	$IISCustomFile1 = $env:SystemDrive + "\Inetpub\wwwroot\Redirmain.html"
	$IISCustomFile2 = $env:SystemDrive + "\Inetpub\wwwroot\default.aspx"

	IF ((Test-Path -Path $IISCustomsFldr -PathType Container) -eq $true)
	{
		Copy-Item -Path $IISCustomFile1 -Destination $IISCustomsFldr
		If ((Test-Path -Path $IISCustomFile1 -PathType Leaf) -eq $true)
		{
			$BodyMessage6 = "IIS Customization File Backups: <br />"
			$BodyMessage6 += "Backup copy of IIS Redirmain.html was successful.<br />"
		}
		Else
		{
			$BodyMessage6 = "IIS Customization File Backups: <br />"
			$BodyMessage6 += "Backup copy of IIS Redirmain.html failed. Please investigate.<br />"
		}
		Copy-Item -Path $IISCustomFile2 -Destination $IISCustomsFldr
		If ((Test-Path -Path $IISCustomFile2 -PathType Leaf) -eq $true)
		{
			$BodyMessage6 += "Backup copy of IIS default.aspx was successful.<br />"
		}
		Else
		{
			$BodyMessage6 += "Backup copy of IIS default.aspx failed. Please investigate.<br />"
		}
	}
	Else
	{
		New-Item -ItemType Directory -Path $IISCustomsFldr -Force
		Copy-Item -Path $IISCustomFile1 -Destination $IISCustomsFldr
		If ((Test-Path -Path $IISCustomFile1 -PathType Leaf) -eq $true)
		{
			$BodyMessage6 = "IIS Customization File Backups: <br />"
			$BodyMessage6 += "Backup copy of IIS Redirmain.html was successful.<br />"
		}
		Else
		{
			$BodyMessage6 = "IIS Customization File Backups: <br />"
			$BodyMessage6 = "Backup copy of IIS Redirmain.html failed. Please investigate.<br />"
		}
		Copy-Item -Path $IISCustomFile2 -Destination $IISCustomsFldr
		If ((Test-Path -Path $IISCustomFile2 -PathType Leaf) -eq $true)
		{
			$BodyMessage6 += "Backup copy of IIS default.aspx was successful.<br />"
		}
		Else
		{
			$BodyMessage6 += "Backup copy of IIS default.aspx failed. Please investigate.<br />"
		}
	}
#EndRegion

#Region Export-CA-Templates
#Export (Backup) CA Templates

	$TempBkpRoot = $TodaysFldr + "\Templates"
	#Check pre-requisite folders exist, if not, create them
	If ((Test-Path -Path $TodaysFldr -PathType Container) -eq $True)
	{
		If ((Test-Path -Path $TempBkpRoot -PathType Container) -eq $False)
		{
			New-Item -ItemType Directory -Path $TempBkpRoot -Force   
		}
	}

	#Get list of templates to process
	$Templates = Get-CertificateTemplate -Name * | Select-Object -Property Name
	ForEach ($Temp in $Templates)
	{
		$Template = ($Temp).Name
		$Path = $TempBkpRoot + "\" + $Template + ".dat"
	    
		fnExport-CertificateTemplate $Template $Path
	}

	#Create .Zip backup to save space
	$ZipFile = $TempBkpRoot + "\CertTemplates.zip"
	Get-ChildItem -Path $TempBkpRoot -Filter *.dat | Copy-ToZip -ZipFile $ZipFile -HideProgress -ErrorAction SilentlyContinue

	If ((Test-Path -Path $ZipFile -PathType Leaf) -eq $true)
	{
	    Get-ChildItem -Path $TempBkpRoot -Filter *.dat | Remove-Item
		$TempBkpRootInfo = Get-ChildItem -Path $TempBkpRoot -Recurse -File | Select-Object -Property @{Name="Name"; Expression={$_.Name.ToUpper()}}, CreationTime, @{Name="Kilobytes";Expression={[math]::Round($_.Length / 1Kb,2)}}  | Out-String
		$BodyMessage7 = "Certificate Template Export Results:<br />"
		$BodyMessage7 += "Certficate Template definitions have been exported to " + $ZipFile + ".<br />"
		$BodyMessage7 += $TempBkpRootInfo + "<br  />"
	}
	Else
	{	
		$BodyMessage7 = "Certificate Template Export Results:<br />"
		$BodyMessage7 += "Template backups failed. Please contact Administrator."
	}
#EndRegion

#Region Get-Backup-Folder-Size
$colSubFldrs = (Get-ChildItem $TodaysFldr -Recurse | Where-Object {$_.PSIsContainer -eq $True} | Sort-Object)
ForEach ($fldr in $colSubFolders)
{
	$fldrSize = (Get-ChildItem ($fldr).FullName -Recurse | Measure-Object -Property Length -Sum).sum
	If ($size -ge 1GB)
	{
		$fldrSize = "{0:N2}" -f  ($fldrSize / 1GB) + " GB"
	}
	ElseIf ($size -ge 1MB)
    {
    	$fldrSize = "{0:N2}" -f  ($fldrSize / 1MB) + " MB"
    }
	Else
    {
    	$fldrSize = "{0:N2}" -f  ($fldrSize / 1KB) + " KB"
    }
	
	$BodyMessage8 = $fldrSize
}
#EndRegion

#Region Remove-Old-Backups
#Cleanup old CA Backups
$Results = Get-ChildItem -Path $volBkpFldr -Force
ForEach ($Result in $Results)
{
	$ExpiredBackupTime = ((Get-Date) - $Result.CreationTime).Hours
	$FolderName += ($Result).FullName
	If ($ExpiredBackupTime -gt [int]$RetentionLimit -and $Result.PSIsContainer -eq $true)
	{
		$Result | Remove-Item -Force -Recurse
	}
	
	If ($?) 
	{
		$BodyMessage9 = $FolderName + " was deleted as of $(fnGet-LongDate)"
	}
	Else 
	{
		$BodyMessage9 = "There are no backup folders ready for deletion as of $(fnGet-Date)."
	}
}
#EndRegion

$BodyMessage10 = "Script name: $myScriptName completed at $(fnGet-LongDate)"

#Send E-mail Report
fnSend-AdminEmail $BodyMessage1 $BodyMessage2 $BodyMessage3 $BodyMessage4 $BodyMessage5 $BodyMessage6 $BodyMessage7 $BodyMessage8 $BodyMessage9 $BodyMessage10
#EndRegion