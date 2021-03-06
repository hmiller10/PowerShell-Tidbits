#Region Help
<#

.NOTES Needs input file with user information

.SYNOPSIS Modify user RDS profile paths to updated target

.DESCRIPTION Update Remote Desktop Services Home Drive and Home Directory
to match Profile Path Home Drive and Home Directory Settings

.OUTPUTS Log file with names of user's whose RDS profile paths were updated

.EXAMPLE 
.\Update-RDSPathFromFile.ps1 -File Path\To\File.ps1


#>
###########################################################################
#
#
# AUTHOR:  Heather Miller
#          
#
# VERSION HISTORY: 1.0 Initial Release
# 
#
# 
###########################################################################
#EndRegion
Param( 
[Parameter(Position=0, Mandatory=$true)] 
[string] 
[ValidateNotNullOrEmpty()] 
[alias("f")] 
$File
)#End Param


Try 
{
	Import-Module ActiveDirectory -ErrorAction Stop
}
Catch
{
	Try
	{
	    Import-Module C:\Windows\System32\WindowsPowerShell\v1.0\Modules\ActiveDirectory\ActiveDirectory.psd1 -ErrorAction Stop
	}
	Catch
	{
	   Throw "Active Directory module could not be loaded. $($_.Exception.Message)"
	}
	
}

#Variables
$aryProperties = "DisplayName","distinguishedName","HomeDirectory"
$aryTSProperties = "allowLogon","TerminalServicesHomeDirectory","TerminalServicesHomeDrive","TerminalServicesProfilePath"
$Date = Get-Date -Format "yyyy-MM-dd"
$path = "D:\Path\To\RDSPath_Updates_for_AAUsers_$Date.txt"

#Create log file
IF(!(Test-Path -Path $path))
{
	New-Item -Path $path –ItemType file -Force
}

#Get list of users from source file 
$Users = Get-Content $File






#Begin Functions
#Get Remote Desktop Profile Settings, Write To Output File
Function fnQueryTSProperties  {#Begin Function to query RDS path attributes for a given user
	 Param ($UserName)
	 ForEach($property in $aryTSProperties)
	 {
	  	$data = "$($Property) value: $($UserName.PSBase.InvokeGet($Property))"
	  	Out-File -FilePath $path -InputObject $data -Append
	 }
} #End Function fnQueryTSProperties

#Update user Remote Desktop Profile Home Drive and Home Directory, Write Output To File
Function fnUpdUserRDSInfo {#Begin Function to Update/Modify RDS Path Information
	Param ($UserName,$TShdValue,$TShdlValue)
	fnQueryTSProperties $UserName
	$UserName.PSBase.InvokeSet("allowLogon",1)
	$UserName.PSBase.InvokeSet("TerminalServicesProfilePath","")
	$UserName.PSBase.InvokeSet("TerminalServicesHomeDirectory",$TShdValue)
	$UserName.PSBase.InvokeSet("TerminalServicesHomeDrive",$TShdlValue)
	$UserName.SetInfo()
	Out-File -FilePath $path -InputObject $Separator -Append
	fnQueryTSProperties $UserName
} #End Function


#Begin processing users from input file
ForEach ($User in $Users) {
	$ADUser = Get-ADUser $User -Properties $aryProperties
	$DisplayName = $ADUser.DisplayName
	$UserDN = $ADUser.distinguishedName
	$UserName = [ADSI]"LDAP://$UserDN"
	$TShdValue = $ADUser.HomeDirectory
	$TShdlValue = "H:"
	$RDSPathHdr = "========== Begin Modifying RDS Path Information for Active Directory User Account: " + $DisplayName + " =========="
	Out-File -FilePath $path -InputObject $RDSPathHdr -Append
	fnUpdUserRDSInfo $UserName $TShdValue $TShdlValue
	$RDSPathFtr = "========== End Modifications of RDS Path Information for Active Directory User Account: " + $DisplayName + " =========="
	Out-File -FilePath $path -InputObject $RDSPathFtr -Append
}