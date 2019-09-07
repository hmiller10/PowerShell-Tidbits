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
Script to cleanup orphaned adminSDHolder objects

.DESCRIPTION
This script uses ADSI/LDAP to bind to an object in Active Directory, reset the
adminCount attribute to null and enable inheritance on the object's ACL. It requires
a .CSV file of the orphaned adminSDHolder user accounts by distinguishedName.

.OUTPUTS
Console output

.EXAMPLE 
.\Cleanup-adminSDHolders-ADSI.ps1 -inputFile <Path\To\inputFile.csv>

###########################################################################
#
#
# AUTHOR:  Heather Miller
#          
#
# VERSION HISTORY:
# 1.0 9/11/2017 - Initial release
#
###########################################################################
#>
#EndRegion

Param (
	[Parameter(Mandatory=$true,Position=0,ValueFromPipeline=$true,HelpMessage="Where is the input file located?")]
	[alias("File")]
	[String]$inputFile
)

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
#EndRegion

#Region Functions
Function fnSet-AdminUser {  
	<# 
	.SYNOPSIS
		Clears adminCount, and enables inherited security on a user account using ADSI.
	
	.DESCRIPTION
		Clears adminCount, and enables inherited security on a user account by binding to the ADSI/LDAP object.
	
	.NOTES
	    Version    	      	: v1.1
	    Rights Required		: Current adminSDHolder
		Author		: Heather Miller, hemiller@deloitte.com
	
	.INPUTS
		Pipeline input to this function
			
	.PARAMETER objectName
	
	.EXAMPLE 
		fnSet-AdminUser -objectName [distinguishedName]
			
		Description
		-----------
		Clears the adminCount of the specified user, and enables ACL security inheritance
	
	.EXAMPLE 
		Get-AdGroupMember [group name] | fnSet-AdminUser
		
		Description
		-----------
		Clears the adminCount of all group members, and enables ACL security inheritance
	
	#>
	
	[CmdletBinding(SupportsShouldProcess = $True)]
	Param (
		[Parameter(ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $True, Mandatory = $true)]
	    [ValidateNotNullOrEmpty()]
		[string]$objectName
	)
	Begin
	{
		## allows inheritance 
		[bool]$isProtected = $false
		## preserves inherited rules 
		[bool]$preserveInheritance = $true
	}
	Process
	{
        #[string]$dn = (Get-ADUser $UserName -Server $domain).DistinguishedName
        #[String]$dn = Get-ADUser $UserName -Server $domain
	    #Set-AdObject -identity $dn -clear adminCount -Server $domain
	    $object = [ADSI]"LDAP://$objectName"
        $object.adminCount.Remove(1)
        $object.SetInfo()
		[String]$objDn = $object.distinguishedName
        Write-Host $objDn -ForegroundColor Yellow
	    $acl = $object.objectSecurity
	    Write-Host "Original permissions blocked:" -ForegroundColor Yellow
	    Write-Host $acl.AreAccessRulesProtected -ForegroundColor Yellow
	    If ( $acl.AreAccessRulesProtected )
		{
		    $acl.SetAccessRuleProtection($isProtected,$preserveInheritance)
		    $inherited = $acl.AreAccessRulesProtected
		    $object.commitchanges()
		    Write-Host "Updated permissions blocked:" -ForegroundColor Green
		    Write-Host $acl.AreAccessRulesProtected -ForegroundColor Green
	    }
    }
	End
	{
		Remove-Variable acl
		Remove-Variable objectName
		Remove-Variable isProtected
		Remove-Variable preserveInheritance
		Remove-Variable objDn
		Remove-Variable object
	}
} # end function fnSet-AdminUser
#EndRegion

#Region Variables
$Domain = Get-ADDomain
$orphans = Import-Csv -Path $inputFile
#EndRegion





#Region Script
#Begin Script
ForEach ($orphan in $orphans)
{
    $orphanDN = ($orphan).distinguishedName
    Write-Host $orphanDN -ForegroundColor Cyan
    fnSet-AdminUser $orphanDN
}
#EndRegion