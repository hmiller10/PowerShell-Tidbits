#Region Help
<#

.NOTES

.SYNOPSIS

.DESCRIPTION

.OUTPUTS

.EXAMPLE 
.\Filter-DNSDebugLogv8.ps1

###########################################################################
#
#
# AUTHOR:  Heather Miller
#
# VERSION HISTORY: 8.0
# 
###########################################################################
#>
#EndRegion

[CmdletBinding()]
param
(
	[Parameter( Mandatory = $false,
			 ValueFromPipeline = $true,
			 HelpMessage = 'Specify the text you wish to filter the DNS log on')]
	[String]$FilterText

#Region Variables
[Array]$DCProps = @("Domain", "Name", "IPv4Address")
[Array]$DCs = (Get-ADForest).Domains | ForEach-Object { Get-ADDomainController -Filter * -Server $_ -ErrorAction SilentlyContinue } | Select-Object -Property $DCProps | `
Sort-Object -Property Domain, Name
$LogFldr = "C:\Temp\DNSLogs"
$myPSObjArray = New-Object System.Collections.ArrayList
[PSObject[]]$Hash = New-Object System.Collections.HashTable
[PSObject[]]$Summary = New-Object System.Collections.HashTable
[reflection.assembly]::loadWithPartialname("Microsoft.Office.Interop.Excel") | Out-Null
$xlConstants = "Microsoft.Office.Interop.Excel.Constants" -as [type]
$xlConditionValues=[Microsoft.Office.Interop.Excel.XLConditionValueTypes]
$xlTheme=[Microsoft.Office.Interop.Excel.XLThemeColor]
$xlChart=[Microsoft.Office.Interop.Excel.XLChartType]
$xlIconSet=[Microsoft.Office.Interop.Excel.XLIconSet]
$xlDirection=[Microsoft.Office.Interop.Excel.XLDirection]
$xlTop10Items = 3
$xlTop10Percent = 5
$xlBottom10Percent = 6
$xlBottom10Items = 4
$xlAnd = 1
$xlOr = 2
$xlNormal = -4143
$xlPasteValues = -4163 # Values only, not formulas
$xlBottom = -4107
$xlCellTypeLastCell = 11 # to find last used cell
$xlCenter = -4108
$xlFilterValues = 7
#EndRegion

#Region Functions
Function Delete-EmptyRows {#Begin function to delete empty rows on spreadsheet
	Param($ws)
	
	$used = $ws.usedRange 
	$lastCell = $used.SpecialCells($xlCellTypeLastCell) 
	$row = $lastCell.row 
	 
	For ($i = 1; $i -le $row; $i++) {
	    IF ($ws.Cells.Item($i, 1).Value() -eq $Null)
		{
	        $Range = $ws.Cells.Item($i, 1).EntireRow
	        $Range.Delete()
	    }
	}
}#End function Delete-EmptyRows

Function Get-ComputerNameByIP {
    Param
	(
        $IPAddress = $null
    )
    BEGIN
	{
    }
    PROCESS
	{
        If ($IPAddress -and $_)
		{
            Throw 'Please use either pipeline or input parameter'
            Break
        }
		ElseIf ($IPAddress) 
		{
            ([System.Net.Dns]::GetHostbyAddress($IPAddress)).HostName
        }
		ElseIf ($_)
		{
            [System.Net.Dns]::GetHostbyAddress($_).HostName
        }
		Else
		{
            $IPAddress = Read-Host "Please supply the IP Address"
            [System.Net.Dns]::GetHostbyAddress($IPAddress).HostName
        }
    }
    END
	{
    }
}#End function Get-ComputerNameByIP

Function Get-Duplicates {#Begin function to get duplicate items
	Param($array, [switch]$count)

	Begin
	{
		$hash = @{}
	}
	Process
	{
		$array | ForEach-Object { $hash[$_] = $hash[$_] + 1 }
		If ($count)
		{
			$hash.GetEnumerator() | Where-Object { $_.value -gt 1 } | ForEach-Object {
	            New-Object PSObject -Property @{
	                Server = $_.key
	                Count = $_.value
	            }
	        }
		}
		Else
		{
			$hash.GetEnumerator() | Where-Object { $_.value -gt 1 } | ForEach-Object { $_.key }
		}
}
	End
	{
	}
}#End function Get-Duplicates

Function Get-LastLogon {
<#

.SYNOPSIS
	This function will list the last user logged on or logged in.

.DESCRIPTION
	This function will list the last user logged on or logged in.  It will detect if the user is currently logged on
	via WMI or the Registry, depending on what version of Windows is running on the target.  There is some "guess" work
	to determine what Domain the user truly belongs to if run against Vista NON SP1 and below, since the function
	is using the profile name initially to detect the user name.  It then compares the profile name and the Security
	Entries (ACE-SDDL) to see if they are equal to determine Domain and if the profile is loaded via the Registry.

.PARAMETER ComputerName
	A single Computer or an array of computer names.  The default is localhost ($env:COMPUTERNAME).

.PARAMETER FilterSID
	Filters a single SID from the results.  For use if there is a service account commonly used.
	
.PARAMETER WQLFilter
	Default WQLFilter defined for the Win32_UserProfile query, it is best to leave this alone, unless you know what
	you are doing.
	Default Value = "NOT SID = 'S-1-5-18' AND NOT SID = 'S-1-5-19' AND NOT SID = 'S-1-5-20'"
	
.EXAMPLE
	$Servers = Get-Content "C:\ServerList.txt"
	fnGet-LastLogon -ComputerName $Servers

	This example will return the last logon information from all the servers in the C:\ServerList.txt file.

	Computer          : SVR01
	User              : WILHITE\BRIAN
	SID               : S-1-5-21-012345678-0123456789-012345678-012345
	Time              : 9/20/2012 1:07:58 PM
	CurrentlyLoggedOn : False

	Computer          : SVR02
	User              : WILHITE\BRIAN
	SID               : S-1-5-21-012345678-0123456789-012345678-012345
	Time              : 9/20/2012 12:46:48 PM
	CurrentlyLoggedOn : True
	
.EXAMPLE
	Get-LastLogon -ComputerName svr01, svr02 -FilterSID S-1-5-21-012345678-0123456789-012345678-012345

	This example will return the last logon information from all the servers in the C:\ServerList.txt file.

	Computer          : SVR01
	User              : WILHITE\ADMIN
	SID               : S-1-5-21-012345678-0123456789-012345678-543210
	Time              : 9/20/2012 1:07:58 PM
	CurrentlyLoggedOn : False

	Computer          : SVR02
	User              : WILHITE\ADMIN
	SID               : S-1-5-21-012345678-0123456789-012345678-543210
	Time              : 9/20/2012 12:46:48 PM
	CurrentlyLoggedOn : True

.LINK
	http://msdn.microsoft.com/en-us/library/windows/desktop/ee886409(v=vs.85).aspx
	http://msdn.microsoft.com/en-us/library/system.security.principal.securityidentifier.aspx

.NOTES
	Author:	 Brian C. Wilhite
	Email:	 bwilhite1@carolina.rr.com
	Date: 	 "09/20/2012"
	Updates: Added FilterSID Parameter
	         Cleaned Up Code, defined fewer variables when creating PSObjects
	ToDo:    Clean up the UserSID Translation, to continue even if the SID is local
#>

[CmdletBinding()]
param(
	[Parameter(Position=0,ValueFromPipeline=$true)]
	[Alias("CN","Computer")]
	[String[]]$ComputerName="$env:COMPUTERNAME",
	[String]$FilterSID,
	[String]$WQLFilter="NOT SID = 'S-1-5-18' AND NOT SID = 'S-1-5-19' AND NOT SID = 'S-1-5-20'"
	)

Begin
	{
		#Adjusting ErrorActionPreference to stop on all errors
		$TempErrAct = $ErrorActionPreference
		$ErrorActionPreference = "Stop"
		#Exclude Local System, Local Service & Network Service
	}#End Begin Script Block

Process
	{
		Foreach ($Computer in $ComputerName)
			{
				$Computer = $Computer.ToUpper().Trim()
				Try
					{
						#Querying Windows version to determine how to proceed.
						$Win32OS = Get-WmiObject -Class Win32_OperatingSystem -ComputerName $Computer
						$Build = $Win32OS.BuildNumber
						
						#Win32_UserProfile exist on Windows Vista and above
						If ($Build -ge 6001)
							{
								If ($FilterSID)
									{
										$WQLFilter = $WQLFilter + " AND NOT SID = `'$FilterSID`'"
									}#End If ($FilterSID)
								$Win32User = Get-WmiObject -Class Win32_UserProfile -Filter $WQLFilter -ComputerName $Computer
								$LastUser = $Win32User | Sort-Object -Property LastUseTime -Descending | Select-Object -First 1
								$Loaded = $LastUser.Loaded
								$Script:Time = ([WMI]'').ConvertToDateTime($LastUser.LastUseTime)
								
								#Convert SID to Account for friendly display
								$Script:UserSID = New-Object System.Security.Principal.SecurityIdentifier($LastUser.SID)
								$User = $Script:UserSID.Translate([System.Security.Principal.NTAccount])
							}#End If ($Build -ge 6001)
							
						If ($Build -le 6000)
							{
								If ($Build -eq 2195)
									{
										$SysDrv = $Win32OS.SystemDirectory.ToCharArray()[0] + ":"
									}#End If ($Build -eq 2195)
								Else
									{
										$SysDrv = $Win32OS.SystemDrive
									}#End Else
								$SysDrv = $SysDrv.Replace(":","$")
								$Script:ProfLoc = "\\$Computer\$SysDrv\Documents and Settings"
								$Profiles = Get-ChildItem -Path $Script:ProfLoc
								$Script:NTUserDatLog = $Profiles | ForEach-Object -Process {$_.GetFiles("ntuser.dat.LOG")}
								
								#Function to grab last profile data, used for allowing -FilterSID to function properly.
								function GetLastProfData ($InstanceNumber)
									{
										$Script:LastProf = ($Script:NTUserDatLog | Sort-Object -Property LastWriteTime -Descending)[$InstanceNumber]							
										$Script:UserName = $Script:LastProf.DirectoryName.Replace("$Script:ProfLoc","").Trim("\").ToUpper()
										$Script:Time = $Script:LastProf.LastAccessTime
										
										#Getting the SID of the user from the file ACE to compare
										$Script:Sddl = $Script:LastProf.GetAccessControl().Sddl
										$Script:Sddl = $Script:Sddl.split("(") | Select-String -Pattern "[0-9]\)$" | Select-Object -First 1
										#Formatting SID, assuming the 6th entry will be the users SID.
										$Script:Sddl = $Script:Sddl.ToString().Split(";")[5].Trim(")")
										
										#Convert Account to SID to detect if profile is loaded via the remote registry
										$Script:TranSID = New-Object System.Security.Principal.NTAccount($Script:UserName)
										$Script:UserSID = $Script:TranSID.Translate([System.Security.Principal.SecurityIdentifier])
									}#End function GetLastProfData
								GetLastProfData -InstanceNumber 0
								
								#If the FilterSID equals the UserSID, rerun GetLastProfData and select the next instance
								If ($Script:UserSID -eq $FilterSID)
									{
										GetLastProfData -InstanceNumber 1
									}#End If ($Script:UserSID -eq $FilterSID)
								
								#If the detected SID via Sddl matches the UserSID, then connect to the registry to detect currently loggedon.
								If ($Script:Sddl -eq $Script:UserSID)
									{
										$Reg = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey([Microsoft.Win32.RegistryHive]"Users",$Computer)
										$Loaded = $Reg.GetSubKeyNames() -contains $Script:UserSID.Value
										#Convert SID to Account for friendly display
										$Script:UserSID = New-Object System.Security.Principal.SecurityIdentifier($Script:UserSID)
										$User = $Script:UserSID.Translate([System.Security.Principal.NTAccount])
									}#End If ($Script:Sddl -eq $Script:UserSID)
								Else
									{
										$User = $Script:UserName
										$Loaded = "Unknown"
									}#End Else

							}#End If ($Build -le 6000)
						
						#Creating Custom PSObject For Output
						New-Object -TypeName PSObject -Property @{
							Computer=$Computer
							User=$User
							SID=$Script:UserSID
							Time=$Script:Time
							CurrentlyLoggedOn=$Loaded
							} | Select-Object -Property User #| Select-Object Computer, User, SID, Time, CurrentlyLoggedOn
							
					}#End Try
					
				Catch
					{
						If ($_.Exception.Message -Like "*Some or all identity references could not be translated*")
							{
								Write-Warning "Unable to Translate $Script:UserSID, try filtering the SID `nby using the -FilterSID parameter."	
								Write-Warning "It may be that $Script:UserSID is local to $Computer, Unable to translate remote SID"
							}
						Else
							{
								Write-Warning $_
							}
					}#End Catch
					
			}#End Foreach ($Computer in $ComputerName)
			
	}#End Process
	
End
	{
		#Resetting ErrorActionPref
		$ErrorActionPreference = $TempErrAct
	}#End End

}# End Function Get-LastLogon

Function Get-LongDate {#Begin function to get date and time in long format
	Get-Date -Format G
}#End function Get-LongDate

Function Get-TodaysDate {#Begin function set Todays date format
	Get-Date -Format "MM-dd-yyyy"
}#End function Get-TodaysDate

Function Get-ReportDate {#Begin function set Report date format
	Get-Date -Format "yyyy-MM-dd"
}#End function Get-ReportDate

Function Get-MMddDate {#Begin function set Todays date format
	Get-Date -Format "MM-dd"
}#End function Get-MMddDate
#EndRegion





#Region Script
#Begin Script
#Clear screen
cls
""
Write-Host "Creating folder for today's log processing ..."
$TodaysFldr = $LogFldr + "\" + $(Get-MMddDate)
If( (Test-Path -Path $TodaysFldr -PathType Container) -eq $false) 
{
	New-Item -ItemType Directory -Force -Path $TodaysFldr
}

Set-Location -Path $TodaysFldr
""
Write-Host "Attempting to load MS Excel ..."
Try
{
	$xl = New-Object -ComObject Excel.Application
}
Catch
{
	Throw "This script requires MS Excel to be installed."
	break;
}
Finally
{
	$xl.Visible = $true
	$xl.DisplayAlerts = $false
	Write-Host "MS Excel was successfully loaded ..."
}

Write-Host "Checking to see if spreadsheet exists, if not create new one ..."
$xlFile = "{0}\{1}" -f $TodaysFldr, "$($FilterText).xlsx"

If ((Test-Path -Path $xlFile -PathType Leaf) -eq $true)
{  
	#Open the document  
	$wb = $xl.WorkBooks.Open($xlFile)
	$wb.Author = "Heather Miller"
	$wb.Title = "$($FilterText) DNS Reference Tracking"
	$wb.Subject = "$($FilterText) Reference Tracking"
	
	""
	Write-Host "Adding new worksheet to MS Excel workbook and format for use ..."
	$ws=$wb.Worksheets.Add()
	$ws.Name = Get-MMddDate
	$ws=$wb.Worksheets.Add()
	$ws.Name = "Summary"

}
Else
{
	""
	Write-Host "File path $xlFile is not accessible. Stand-by while new file is created"
	$wb = $xl.WorkBooks.Add()
	$wb.Author = "Heather Miller"
	$wb.Title = "$($FilterTest) DNS Reference Tracking"
	$wb.Subject = "$($FilterTest) Reference Tracking"
	
	""
	Write-Host "Adding new worksheet to MS Excel workbook and format for use ..."
	$ws=$wb.Worksheets.Add()
	$ws.Name = Get-MMddDate
	$ws=$wb.Worksheets.Add()
	$ws.Name = "Summary"
	
	""
	Write-Host "Deleting initial worksheets ..."
	$xl.Worksheets.Item("Sheet1").Delete()
	$xl.Worksheets.Item("Sheet2").Delete()
	$xl.Worksheets.Item("Sheet3").Delete()
}


$s1row = 1
$s1col = 1
$ws1 = $wb.sheets | Where-Object {$_.Name -match $(Get-MMddDate)}
$Sheet1 = $wb.Worksheets.Item(($ws1).Index)
$Sheet1.Activate()
Start-Sleep 1
Write-Host "Adding Column Headings"
"DNS Server","Date","Time","Query Type","Querying Server Name","IPv4Address","Last Logged On User" | ForEach-Object {
	$ws1.cells.item($s1row,$s1col) = $_
	$ws1.cells.item($s1row,$s1col).font.bold = $true
	$ws1.cells.item($s1row,$s1col).font.size = 14
	$ws1.cells.item($s1row,$s1col).interior.colorindex = 49
	$ws1.cells.item($s1row,$s1col).font.colorindex = 2
	$ws1.cells.item($s1row,$s1col).HorizontalAlignment = $xlConstants::xlCenter
	$ws1.cells.item($s1row,$s1col).VerticalAlignment = $xlConstants::xlBottom
	$s1col++
}

$s1range1 = $ws1.usedRange

Write-Host "Freeze Top Row In Spreadsheet"
$ws1.Application.ActiveWindow.SplitColumn = 0
$ws1.Application.ActiveWindow.SplitRow = 1
$ws1.Application.ActiveWindow.FreezePanes = $true
$s1row++

ForEach ($DnsSrv in $DCs)
{
	$srvName = ($DnsSrv).Name
	""
	Write-Host "Copying DNS log from $srvName to process ..."

	$SourceFile =  "{0}{1}\{2}\{3}" -f "\\", $srvName, C$, "dnslog.txt"
	Write-Host $SourceFile
	IF ( ( Test-Path -Path $SourceFile -PathType Leaf ) -eq $false)
	{
		break;
	}
	ELSE
	{
		$DNSSrvFldr = "{0}\{1}" -f $TodaysFldr, $srvName
		
		IF( (Test-Path -Path $DNSSrvFldr -PathType Container) -eq $false) 
		{
			New-Item -ItemType Directory -Force -Path $DNSSrvFldr
		}
		
		Copy-Item $SourceFile -Destination ($DNSSrvFldr + "\" + $srvName + ".txt")
	}

	#Define working location of log files to be processed
	$DNSLogFile = "{0}\{1}.{2}" -f $DNSSrvFldr, $srvName, "txt"
	$LogFile = "{0}\{1}.{2}" -f $DNSSrvFldr, $($FilterText), "log"
	IF ( -not (Test-Path -Path $LogFile -PathType Leaf) -eq $true)
	{
		New-Item -Path $LogFile -ItemType File -Force
	}
	
	Write-Host "Processing Debug Log $DNSLogFile for $($FilterText) references ..." 
	#Parse $LogFile for specified string pattern
	$filterLength = $FilterText.Length
	if (($FilterText -notmatch ':') -or ($FilterText -notmatch '.'))
	{
    		$parts = $FilterText.Split('')
    		[int32]$i = 0
    		$arrayOfStrings = @()
 
		Do
	    	{	
			#Build formatted filter	
			$string = "({0}){1}" -f $parts[$i].Length, $parts[$i]
			$arrayOfStrings += $string
			$i++					
			
		} While ($i -le $FilterLength)
		
		$arrayOfStrings = $arrayOfStrings -join ('')
	}
	elseif ($FilterText -match ':')
	{
		$parts = $FilterText.Split(':')
    		[int32]$i = 0
    		$arrayOfStrings = @()
 
		Do
	    	{	
			#Build formatted filter	
			$string = "({0}){1}" -f $parts[$i].Length, $parts[$i]
			$arrayOfStrings += $string
			$i++					
			
		} While ($i -le $FilterLength)
		
		$arrayOfStrings = $arrayOfStrings -join ('')
	}
	elseif ($FilterText -match '.')
	{
		$parts = $FilterText.Split('.')
    		[int32]$i = 0
    		$arrayOfStrings = @()
 
		Do
	    	{	
			#Build formatted filter	
			$string = "({0}){1}" -f $parts[$i].Length, $parts[$i]
			$arrayOfStrings += $string
			$i++					
			
		} While ($i -le $FilterLength)
		
		$arrayOfStrings = $arrayOfStrings -join ('')
	}
	
	Get-Content -Path $DNSLogFile | Select-String -Pattern $arrayOfStrings -SimpleMatch | Select-String -Pattern 'Rcv' -SimpleMatch | Out-File -FilePath $LogFile
	(Get-Content $LogFile) | ForEach { $_.Trim() } | Where-Object { $_ -ne "" } | Set-Content -Path $LogFile

	$reader = [System.IO.File]::OpenText($LogFile)
	Try
	{
		For ()
		{
			$line = $reader.ReadLine()
			If ( $null -eq $line )
			{
				break
			}
			$LogDate = ($line).split(" ")[0]
			$LogTime = ($line).split(" ")[1] + ($line).split(" ")[2]
			$LogQueryType = ($line).split(" ")[8]
			Try
			{
				Get-ComputerNameByIP ($line).split(" ")[9]
			}
			Catch
			{
				If ( $Error[0] )
				{
					$LogHostName = ""
				}
			}
			Finally
			{
				If ( $LogHostName -eq "" )
				{
					$LogHostName = "Querying Server isNull"
				}
				Else
				{
					$LogHostName = Get-ComputerNameByIP ($line).split(" ")[9]
				}
			}
			
			$LogIPAdress = ($line).split(" ")[9]
			Try
			{
				$LastLogonUser = Get-LastLogon -ComputerName $LogHostName
				$LogLogonUser = ($LastLogonUser).User
			}
			Catch
			{
				If ( $Error[0] )
				{
					$LastLogonUser = "##ERROR##"
				}
			}
			
			Write-Host "Adding object values to PS Hash from server $srvName ..."
			$Hash += New-Object -TypeName PSCustomObject -Property @{
				DNSServer = $srvName
				Date = $LogDate
				Time = $LogTime
				QueryType = $LogQueryType
				QueryingServer = $LogHostName
				IPv4Address = $LogIPAdress
				LastLogonUser = $LogLogonUser.Value
			}
				
		$null = $line = $LogDate = $LogTime = $LogQueryType = $LogHostName = $LogIPAddress = $LastLogonUser = $LogLogonUser
		}
	}
	Finally
	{
		$reader.Close()
	}

	$null - $srvName = $DNSLogFile = $SourceFile = $DNSSrvFldr = $LogFile = $DNSSrv
}


ForEach ( $obj in $Hash )
{
	$s1col = 1
	$ws1.cells.item($s1row,$s1col) = $obj.DNSServer
	$s1col++
	$ws1.cells.item($s1row,$s1col) = $obj.Date
	$s1col++
	$ws1.cells.item($s1row,$s1col) = $obj.Time
	$s1col++
	$ws1.cells.item($s1row,$s1col) = $obj.QueryType
	$s1col++
	$ws1.cells.item($s1row,$s1col) = $obj.QueryingServer | Out-String
	$s1col++
	$ws1.cells.item($s1row,$s1col) = $obj.IPv4Address | Out-String
	$s1col++
	$ws1.cells.item($s1row,$s1col) = $obj.LastLogonUser | Out-String
	$s1row++
}


#Delete empty rows on Sheet 1
Delete-EmptyRows $ws1

$s1range1.EntireRow.AutoFilter() | Out-Null

$s1range2 = $ws1.UsedRange
$s1range2.EntireColumn.AutoFit()
$s1range2.EntireRow.AutoFit()
$s1range2.HorizontalAlignment = $xlCenter
$s1range2.VerticalAlignment = $xlBottom
Start-Sleep 1


""
Write-Host "Building summary worksheet ..."
$s2row = 1
$s2col = 1
$ws2 = $wb.sheets | Where-Object {$_.Name -match "Summary"}
$Sheet2 = $wb.Worksheets.Item(($ws2).Index)
$Sheet2.Activate()
Start-Sleep 1
Write-Host "Adding Column Headings to Sheet 2"
"Utilization Summary"," "," "," " | ForEach-Object {
	$ws2.cells.item($s2row,$s2col) = $_
	$ws2.cells.item($s2row,$s2col).font.bold = $true
	$ws2.cells.item($s2row,$s2col).font.size = 14
	$ws2.cells.item($s2row,$s2col).interior.colorindex = 49
	$ws2.cells.item($s2row,$s2col).font.colorindex = 2
	$s2col++
}

$s2range1 = $ws2.UsedRange
$s2range1.Select()
$s2range1.MergeCells = $true
$s2range1.HorizontalAlignment = $xlCenter
$s2range1.VerticalAlignment = $xlBottom

$s2row++
$s2col = 1

"Date","Total number of Filter References", "Servers Querying for FilteredText", "Number of Queries by Server" | ForEach-Object {
	$ws2.cells.item($s2row,$s2col) = $_
	$ws2.cells.item($s2row,$s2col).font.bold = $true
	$ws2.cells.item($s2row,$s2col).font.size = 14
	$ws2.cells.item($s2row,$s2col).interior.colorindex = 49
	$ws2.cells.item($s2row,$s2col).font.colorindex = 2
	$s2col++
}

$s2range2 = $ws2.UsedRange
$s2range2.Select()
$s2range2.EntireColumn.AutoFit()
$s2range2.EntireRow.AutoFit()

#Freeze top row of worksheet 2
Write-Host "Freeze Top Rows In Spreadsheet"
$ws2.Application.ActiveWindow.SplitColumn = 0
$ws2.Application.ActiveWindow.SplitRow = 2
$ws2.Application.ActiveWindow.FreezePanes = $true
$s2row++

[int]$Refs = ($Hash).Count
[Array]$QueryingServers = $Hash | Select-Object -Property QueryingServer | Sort-Object -Property QueryingServer
$qNames = ($QueryingServers).QueryingServer
#$qNames
$qServers = Get-Duplicates $qNames -Count


$ws2.cells.item($s2row,1) = Get-TodaysDate
$ws2.cells.item($s2row,2) = $Refs

ForEach ( $qServer in $qServers )
{
	$s2col = 3
	$ws2.cells.item($s2row,$s2col) = ($qServer).Server | Out-String
	$s2col++
	$ws2.cells.item($s2row,$s2col) = ($qServer).Count
	$s2col++
	$s2row++
$qServer = ""
}

$s2range3 = $ws2.UsedRange
$s2range3.EntireColumn.AutoFit()
$s2range3.EntireRow.AutoFit()
$s2range3.HorizontalAlignment = $xlCenter
$s2range3.VerticalAlignment = $xlBottom
$srvRangeStart = $ws2.Range("C3")
#Get the last cell in col C
$srvRangeEnd = $ws2.Range($srvRangeStart,$srvRangeStart.End($xlDirection::xlDown))
$srvRangeEnd.Select()
$srvRangeEnd.AutoFilter() | Out-Null

""
Write-Host "Saving our work ..."
IF((Test-Path -Path $xlFile -PathType Leaf) -eq $true)
{
   $xl.ActiveWorkbook.Save()
   $xl.ActiveWorkbook.Close()
   $xl.Quit()
} 
ELSE
{
   $xl.ActiveWorkbook.SaveAs($xlFile)
   $xl.ActiveWorkbook.Close()
   $xl.Quit()
}
$xl.ActiveWorkbook.
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($xl)
"" 
Write-Host "... Done ..." 
""
#EndRegion