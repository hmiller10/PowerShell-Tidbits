#------------------------------------------------------------------------------
#
# THIS CODE IS MADE AVAILABLE AS IS, WITHOUT WARRANTY OF ANY KIND. THE
# ENTIRE RISK OF THE USE OR THE RESULTS FROM THE USE OF THIS CODE REMAINS
# WITH THE USER.
#
#------------------------------------------------------------------------------
#
# NAME:
#    Zip_and_Purge_IISLogs.ps1
#
# AUTHOR:
#    Heather Miller
#
#------------------------------------------------------------------------------

#Modules

try {
	Import-Module -Name WebAdministration -Force
}
catch {
	Import-Module 'C:\Windows\System32\WindowsPowerShell\v1.0\Modules\WebAdministration\WebAdministration.psd1'
}


#Variables
$driveRoot = (Get-Location).Drive.Root
$logArchives = "{0}{1}" -f $driveRoot, "LogArchives"
[int]$purgeLimit = -90
[int]$archiveLimit = -30
Add-Type -AssemblyName WindowsBase
Add-Type -AssemblyName "System.IO.Compression.Filesystem"
$now = [System.DateTime]::UtcNow
$currentMonth = ($now).Month
$currentYear = ($now).Year
$currentDay = ($now).Day
$previousmonth = ((Get-Date).AddMonths(-1)).Month
$firstdayofpreviousmonth = (Get-Date -Year $currentYear -Month $currentMonth -Day 1).AddMonths(-1)

#Functions

Function Test-PathExists {#Begin function to check path variable and return results
	[CmdletBinding()]
	Param
	(
		[Parameter(Mandatory,Position=0)]
		[string]$Path,
		[Parameter(Mandatory,Position=1)]
		$PathType
	)

	Switch ( $PathType )
	{
    		File	{
				If ( ( Test-Path -Path $Path -PathType Leaf ) -eq $true )
				{
					Write-Verbose -Message "File: $Path already exists..." -Verbose
				}
				Else
				{
					New-Item -Path $Path -ItemType File -Force
					Write-Verbose -Message "File: $Path not present, creating new file..." -Verbose
				}
			}
		Folder
			{
				If ( ( Test-Path -Path $Path -PathType Container ) -eq $true )
				{
					Write-Verbose -Message "Folder: $Path already exists..." -Verbose
				}
				Else
				{
					New-Item -Path $Path -ItemType Directory -Force
					Write-Verbose -Message "Folder: $Path not present, creating new folder" -Verbose
				}
			}
	}
}#end function Test-PathExists


function Get-NetFrameWork45OrHigherInstalled
{
    if (Test-Path “HKLM:\Software\Microsoft\Net Framework Setup\NDP\v4\Full”)
    {
	    # Example values of the Release DWORD
	    # 378389 = .NET Framework 4.5
	    # 378675 = .NET Framework 4.5.1 installed with Windows 8.1
	    # 378758 = .NET Framework 4.5.1 installed on Windows 8, Windows 7 SP1, or Windows Vista SP2
	    $RegRemoteBaseKey = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey([Microsoft.Win32.RegistryHive]::LocalMachine, [String]::Empty)
	    $RegSubKey = $RegRemoteBaseKey.OpenSubKey("Software\Microsoft\Net Framework Setup\NDP\v4\Full\")
	    $RegSubKeyValue = [string]$RegSubKey.GetValue("Release")
	    $RegSubKey.Close()
	    $RegRemoteBaseKey.Close()
	    if ($RegSubKeyValue -ge "378389")
	    {
		    Add-Type -AssemblyName System.IO.Compression, System.IO.Compression.FileSystem
            return $true
	    }
    }
    return $false
}


function Zip-File
{
	[CmdletBinding()]
	param
	(
		[Parameter(Mandatory = $False)]
		[String]$CompressedFileName,

		[Parameter(Mandatory = $False)]
		[String]$FileToCompress,
		
		[Parameter(Mandatory = $False)]
		[String]$EntryName,

		[Parameter(Mandatory = $False)]
		[String]$ArchiveMode
	)
	
	switch ($ArchiveMode)
	{
		"Create" {$objCompressedFile = [System.IO.Compression.ZipFile]::Open($CompressedFileName, [System.IO.Compression.ZipArchiveMode]::Create)}
		"Read" {$objCompressedFile = [System.IO.Compression.ZipFile]::Open($CompressedFileName, [System.IO.Compression.ZipArchiveMode]::Read)}
		"Update" {$objCompressedFile = [System.IO.Compression.ZipFile]::Open($CompressedFileName, [System.IO.Compression.ZipArchiveMode]::Update)}
	}
	$compressionLevel = [System.IO.Compression.CompressionLevel]::Optimal
	$archiveEntry = [System.IO.Compression.ZipFileExtensions]::CreateEntryFromFile($objCompressedFile, $FileToCompress, $EntryName, $compressionLevel)
	$objCompressedFile.Dispose()
}






#Script

Test-PathExists -Path $logArchives -PathType Folder

# Get the IIS log file top level folder, then process the W3SVC* log files folder for each site

try {
	[Array]$colWebSites = Get-WebSite
}
catch {
	exit
}


foreach ( $site In $colWebSites )
{
	$iisLogsFolderPath = $site.logFile.directory
	if ( $iisLogsFolderPath -match "%systemdrive%" ) 
	{ 
		$iisLogsFolderPath = [string]($site.logFile.directory) -replace "%systemdrive%", $env:SystemDrive 
	}

	$iisLogsFolderPath += "{0}{1}{2}" -f "\","W3SVC", $site.id
	
	# add log file to a zip file and delete log file
	$colLogFiles = Get-ChildItem -Path $iisLogsFolderPath -Filter *.log	| Where-Object { ( $_.CreationTime.Month -eq $previousMonth ) -and ( $_.CreationTime.Year -eq $currentYear ) }
	
	$logArchiveFileName = "{0}{1}_{2}_{3}_{4}.{5}" -f "W3SVC", $site.id, $currentYear, $previousmonth, "IISLogs", "zip"
	$logArchiveFile = "{0}\{1}" -f $logArchives, $logArchiveFileName

	if ( $colLogFiles )
	{
		ForEach ($file In $colLogFiles)
		{

			Write-Host ("Processing file for zip: {0}" -f $file.FullName.ToString())
			$FileToCompress = $file.FullName.ToString()
			$FileName = $file.Name.ToString()


			Write-Verbose ("[{0} UTC] [SCRIPT] File to compress:  `"{1}`"." -f [datetime]::UtcNow.ToString($dtmFormatString), $FileToCompress) -Verbose

			try
			{
				if ( Get-NetFrameWork45OrHigherInstalled -eq $true )
				{
					# if compressed file name already exists, use update mode.  otherwise, create new archive
					if ( Test-Path -Path $logArchiveFile )
					{
						Zip-File -CompressedFileName $logArchiveFile -FileToCompress $FileToCompress -EntryName $FileName -ArchiveMode "Update"
					}
					else
					{
						Zip-File -CompressedFileName $logArchiveFile -FileToCompress $FileToCompress -EntryName $FileName -ArchiveMode "Create"
					}
				}
				else
				{
					# do nothing.  assumes script is running on Windows Server 2012 or higher.
					Write-Verbose ("[{0} UTC] [SCRIPT] Missing .Net Framework 4.5 or higher." -f [datetime]::UtcNow.ToString($dtmFormatString)) -Verbose
				}
			}
			catch
			{
				Write-Verbose ("[{0} UTC] [SCRIPT] Error compressing file: {1}" -f [datetime]::UtcNow.ToString($dtmFormatString), $Error[0].Exception.Message) -Verbose
			}

			# before deleting the log file, make sure that the zip file contains the log file entry
			if ( ( Test-Path -Path $logArchiveFile ) -eq $true )
			{
				$zip = [io.compression.zipfile]::OpenRead($logArchiveFile)
				$ZippedFile = $zip.Entries | Where-Object { $_.Name -eq $FileName }
				if ( $ZippedFile )
				{
					Remove-Item -Path $File.FullName -Force -Confirm:$false
				}
				$zip.Dispose()
				
			}

		}#End ForEach $file

	}#End If $colLogFiles

}#end ForEach $site

# remove previously zipped files that are older than $purgeLimit
[Array]$colZipFiles = Get-ChildItem -Path $logArchives -Include *.zip

If ($colZipFiles)
{
	foreach($zipFile In $colZipFiles){
		If ($zipFile.LastWriteTime.ToUniversalTime() -lt [DateTime]::UtcNow.AddDays($purgeLimit))
		{
			Write-Verbose ( "Processing file for delete: {0}" -f $file.FullName.ToString() )
			[System.IO.File]::Delete($file.FullName.ToString())
		}
	}
}

#End script