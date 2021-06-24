<#
	.NOTES
	THIS CODE IS MADE AVAILABLE AS IS, WITHOUT WARRANTY OF ANY KIND. THE
	RISK OF THE USE OR THE RESULTS FROM THE USE OF THIS CODE REMAINS WITH
	THE USER.

	.SYNOPSIS
		Script to process pending certificate issuance requests using PSPKI module
		
	.DESCRIPTION
		This script will connect to the certificate authority server passed into
		the script as a parameter and will locate any pending certificate issuance
		requests. Upon finding a pending a request, the person running this script
		will be prompted to either approve or deny the request. The script will then 
		respond accordingly depending on the user's response.

	.PARAMETER CA
		Fully qualified domain name of certificate authority server

	.OUTPUTS
		None

	.EXAMPLE
	PS C:\> Invoke-ProcessPendingRequest.ps1 -CA myca.domain.com

	.LINK
    	https://www.sysadmins.lv/blog-en/categoryview/powershellpowershellpkimodule.aspx

	.LINK
    	https://github.com/Crypt32/PSPKI
	
#>


[CmdletBinding()]
Param (
	[Parameter(Mandatory = $true, Position = 0, ValueFromPipeline = $true, HelpMessage = "Specify the fully qualified domain name of the PKI server.")]
	[String]$CA
)

#Modules
Try
{
	Import-Module PSPKI -ErrorAction Stop
}
Catch
{
	Try
	{
		$modulePath = "{0}\{1}\{2}\{3}" -f $env:ProgramFiles, "WindowsPowerShell", "Modules", "PSPKI"
		$psdPath = "{0}\{1}\{2}" -f $modulePath, (Get-Module -Name PSPKI).Version, "pspki.psd1"
		Import-Module $psdPath -ErrorAction Stop
	}
	Catch
	{
		Throw "PSPKI module could not be loaded. $($_.Exception.Message)"
	}
	
}

#Variables
$certProps = @()
$certProps = @("RequestID", "Request.RequesterName", "CommonName", "NotBefore", "NotAfter", "SerialNumber", "CertificateTemplate")
$pendingRequests = @()



#Functions
Function Get-AnApprovalDecision
{
	#Begin function to select singular reason for pending certificate request approval or denial
	
	[CmdletBinding()]
	Param ()
	
	Begin
	{
		Add-Type -Assembly System.Windows.Forms
		Add-Type -Assembly System.Drawing
		
		$csvData = ConvertFrom-Csv @"
		Choice
		"Approve Request"
		"Deny Request"
"@
	}
	Process
	{
		$colApprovals = New-Object System.Collections.ArrayList
		foreach ($row in $csvData)
		{
			[void]$colApprovals.Add($row.Choice)
		}
		
		$form = New-Object System.Windows.Forms.Form
		$form.width = 500
		$form.height = 200
		$form.StartPosition = "CenterScreen"
		$form.Text = "Select Pending Certificate Request Approval Decision"
		$form.Font = New-Object System.Drawing.Font("Verdana", 10)
		
		$comboBox = New-Object System.Windows.Forms.ComboBox
		$comboBox.Location = New-Object System.Drawing.Size(50, 20)
		$comboBox.Size = New-Object System.Drawing.Size(400, 40)
		$comboBox.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDown
		
		$okButton = New-Object System.Windows.Forms.Button
		$okButton.Location = New-Object System.Drawing.Size(130, 100)
		$okButton.Size = New-Object System.Drawing.Size(100, 30)
		$okButton.Text = "OK"
		$okButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
		#$okButton.Anchor = [System.Windows.Forms.AnchorStyles]::None
		$form.Controls.Add($okButton)
		
		$cancelButton = New-Object System.Windows.Forms.Button
		$cancelButton.Location = New-Object System.Drawing.Size(255, 100)
		$cancelButton.Size = New-Object System.Drawing.Size(100, 30)
		$cancelButton.Text = "Cancel"
		$cancelButton.Add_Click({ $form.Close() })
		#$cancelButton.Anchor = [System.Windows.Forms.AnchorStyles]::None
		$form.Controls.Add($cancelButton)
		
		[void]$comboBox.Items.AddRange($colApprovals)
		
		$form.Controls.Add($comboBox)
		$form.AcceptButton = $okButton
		$form.CancelButton = $cancelButton
		
		#$form.Add_Shown({$form.Activate()})
		$form.Add_Shown({ $form.Activate(); $okButton.Focus() })
		$selectedReason = $form.ShowDialog()
	}
	End
	{
		if ($selectedReason -eq [System.Windows.Forms.DialogResult]::OK)
		{
			return $comboBox.SelectedItem
		}
		else
		{
			return [string]::Empty
		}
	}
} #End function Get-AnApprovalDecision








#Script
try
{
	$pendingRequests = Get-PendingRequest -CertificationAuthority $CA -Properties *
	
	if ($pendingRequests.Count -gt 0)
	{
		Write-Output ("Total number of pending certificate requests at this time: $($pendingRequests.Count)")
		
		foreach ($certRequest in $pendingRequests)
		{		
			$reqbytes = [convert]::FromBase64String($certRequest."Request.RawRequest")
			$req = New-Object System.Security.Cryptography.X509CertificateRequests.X509CertificateRequest(,$reqBytes)
			
			Write-Output ("Please carefully review the fields in the below shown request. Pay extra attention to any SANs added to the request.")
			
			$req.ToString()
			
			
			$response = Read-Host -Prompt "Ready to proceed? (Y/N)"
			
			while ($response -notcontains "Y")
			{
				if ($response -contains "N") { exit }
				$response = Read-Host -Prompt "Ready to proceed? (Y/N)"
			}
			

			$decision = Get-AnApprovalDecision
			switch ($decision)
			{
				"Approve Request" { $certRequest | Approve-CertificateRequest }
				"Deny Request" { $certRequest | Deny-CertificateRequest }
			}
				
				
		}
	}
	else
	{
		Add-Type -AssemblyName PresentationFramework
		[System.Windows.MessageBox]::Show("There are no pending certificate requests at this time on certificate authority: $($CA)")
	}
	
}
catch
{
	$errorMessage = "{0}: {1}" -f $Error[0], $Error[0].InvocationInfo.PositionMessage
	Write-Error $errorMessage -ErrorAction Continue
	$Error.Clear()
}