########################################################################
# Name: Skype4B / Lync Certificate Checker
# Version: v1.0.0 (1/12/2016)
# Date: 1/12/2016
# Created By: James Cussen
# Web Site: http://www.myskypelab.com
#
# Notes: This is a Powershell tool. To run the tool, open it from the Powershell command line or Right Click and select "Run with Powershell". The tool can be run from Windows or Windows Server.
#		 The tool can be run on any PC or Server with Powershell 2+
# 		 For more information on the requirements for setting up and using this tool please visit http://www.myskypelab.com.
#
# Copyright: Copyright (c) 2016, James Cussen (www.myskypelab.com) All rights reserved.
# Licence: 	Redistribution and use of script, source and binary forms, with or without modification, are permitted provided that the following conditions are met:
#				1) Redistributions of script code must retain the above copyright notice, this list of conditions and the following disclaimer.
#				2) Redistributions of source code must retain the above copyright notice, this list of conditions and the following disclaimer.
#				3) Redistributions in binary form must reproduce the above copyright notice, this list of conditions and the following disclaimer in the documentation and/or other materials provided with the distribution.
#				4) This license does not include any resale or commercial use of this software.
#				5) Any portion of this software may not be reproduced, duplicated, copied, sold, resold, or otherwise exploited for any commercial purpose without express written consent of James Cussen.
#			THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS "AS IS" AND ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE DISCLAIMED. IN NO EVENT SHALL THE COPYRIGHT HOLDER OR CONTRIBUTORS BE LIABLE FOR ANY DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES; LOSS OF USE, DATA, OR PROFITS; LOSS OF GOODWILL OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF THIS SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.
#
# Limitations: 
# - Works on Powershell 2+
# - Only IPv4 DNS lookup is supported.
#
# Import CSV Header Format: Domain,Type,Port
# Example Federation Record: "microsoft.com","FED","",
# Example SIP Record: "microsoft.com","SIP","",
# Example SIP Internal Record: "microsoft.com","SIPINT","",
# Example Direct Record: "sip.microsoft.com","DIR","5061",
#
# Release Notes:
# 1.00 Initial Release.
# 		
########################################################################


$theVersion = $PSVersionTable.PSVersion
$MajorVersion = $theVersion.Major

Write-Host ""
Write-Host "--------------------------------------------------------------"
Write-Host "Powershell Version Check..." -foreground "yellow"
if($MajorVersion -eq  "1")
{
	Write-Host "This machine only has Version 1 Powershell installed.  This version of Powershell is not supported." -foreground "red"
	exit
}
elseif($MajorVersion -eq  "2")
{
	Write-Host "This machine has version 2 Powershell installed. CHECK PASSED!" -foreground "green"
}
elseif($MajorVersion -eq  "3")
{
	Write-Host "This machine has version 3 Powershell installed. CHECK PASSED!" -foreground "green"
}
elseif($MajorVersion -eq  "4")
{
	Write-Host "This machine has version 4 Powershell installed. CHECK PASSED!" -foreground "green"
}
elseif($MajorVersion -eq  "5")
{
	Write-Host "This machine has version 5 Powershell installed. CHECK PASSED!" -foreground "green"
}
else
{
	Write-Host "This machine has version $MajorVersion Powershell installed. Unknown level of support for this version." -foreground "yellow"
}
Write-Host "--------------------------------------------------------------"
Write-Host ""


# boolean for cancelling lookup
$script:CancelScan = $false
$script:FirstCheck = $true
$script:OKFailColumn = $true
$script:RTFDisplayString = ""
$script:RTFStart = "{\rtf1\ansi "
$script:RTFStart += "{\colortbl;\red0\green0\blue0;\red46\green116\blue181;\red70\green70\blue70;\red68\green84\blue106;\red192\green0\blue0;\red112\green173\blue71;\red255\green255\blue0;\red255\green255\blue255;\red0\green0\blue128;\red0\green128\blue128;\red0\green128\blue0;\red128\green0\blue128;\red128\green0\blue0;\red128\green128\blue0;\red128\green128\blue128;\red192\green192\blue192;}"
$script:RTFStart += "{\fonttbl{\f0\fcharset0 Courier New;}}\fs18"
$script:RTFEnd = "}"

$scriptpath = $MyInvocation.MyCommand.Path
$dir = Split-Path $scriptpath
$fso = New-Object -ComObject Scripting.FileSystemObject
$shortname = $fso.GetFolder($dir).Path
Write-host "Script directory: $shortname"


# Set up the form  ============================================================

[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing") 

$objForm = New-Object System.Windows.Forms.Form 
$objForm.Text = "Skype4B / Lync Certificate Checker v1.00"
$objForm.Size = New-Object System.Drawing.Size(880,460) 
$objForm.MinimumSize = New-Object System.Drawing.Size(600,440)
$objForm.StartPosition = "CenterScreen"
[byte[]]$WindowIcon = @(66, 77, 56, 3, 0, 0, 0, 0, 0, 0, 54, 0, 0, 0, 40, 0, 0, 0, 16, 0, 0, 0, 16, 0, 0, 0, 1, 0, 24, 0, 0, 0, 0, 0, 2, 3, 0, 0, 18, 11, 0, 0, 18, 11, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114,0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0,198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 205, 132, 32, 234, 202, 160,255, 255, 255, 244, 229, 208, 205, 132, 32, 202, 123, 16, 248, 238, 224, 198, 114, 0, 205, 132, 32, 234, 202, 160, 255,255, 255, 255, 255, 255, 244, 229, 208, 219, 167, 96, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 248, 238, 224, 198, 114, 0, 198, 114, 0, 223, 176, 112, 255, 255, 255, 219, 167, 96, 198, 114, 0, 198, 114, 0, 255, 255, 255, 255, 255, 255, 227, 185, 128, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 241, 220, 192, 198, 114, 0, 198,114, 0, 248, 238, 224, 255, 255, 255, 244, 229, 208, 198, 114, 0, 198, 114, 0, 255, 255, 255, 255, 255, 255, 227, 185, 128, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 241, 220, 192, 198, 114, 0, 216, 158, 80, 255, 255, 255, 255, 255, 255, 252, 247, 240, 209, 141, 48, 198, 114, 0, 255, 255, 255, 255, 255, 255, 227, 185, 128, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 241, 220, 192, 198, 114, 0, 241, 220, 192, 255, 255, 255, 252, 247, 240, 212, 149, 64, 234, 202, 160, 198, 114, 0, 255, 255, 255, 255, 255, 255, 227, 185, 128, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 241, 220, 192, 205, 132, 32, 255, 255, 255, 255, 255, 255, 227, 185, 128, 198, 114, 0, 248, 238, 224, 202, 123, 16, 255, 255, 255, 255, 255, 255, 227, 185, 128, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 241, 220, 192, 234, 202, 160, 255, 255, 255, 255, 255, 255, 205, 132, 32, 198, 114, 0, 223, 176, 112, 223, 176, 112, 255, 255, 255, 255, 255, 255, 227, 185, 128, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 244, 229, 208, 252, 247, 240, 255, 255, 255, 237, 211, 176, 198, 114, 0, 198, 114, 0, 202, 123, 16, 248, 238, 224, 255, 255, 255, 255, 255, 255, 227, 185, 128, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 212, 149, 64, 255, 255, 255, 255, 255, 255, 255, 255, 255, 212, 149, 64, 198, 114, 0, 198, 114, 0, 198, 114, 0, 234, 202, 160, 255, 255,255, 255, 255, 255, 241, 220, 192, 205, 132, 32, 198, 114, 0, 198, 114, 0, 205, 132, 32, 227, 185, 128, 227, 185, 128, 227, 185, 128, 227, 185, 128, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 205, 132, 32, 227, 185, 128, 227, 185,128, 227, 185, 128, 219, 167, 96, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 198, 114, 0, 0, 0)
$ico = New-Object IO.MemoryStream($WindowIcon, 0, $WindowIcon.Length)
$objForm.Icon = [System.Drawing.Icon]::FromHandle((new-object System.Drawing.Bitmap -argument $ico).GetHIcon())
$objForm.KeyPreview = $True
$objForm.TabStop = $false


#FEDSRVTest Label ============================================================
$FEDSRVTestLabel = New-Object System.Windows.Forms.Label
$FEDSRVTestLabel.Location = New-Object System.Drawing.Size(30,30) 
$FEDSRVTestLabel.Size = New-Object System.Drawing.Size(60,15) 
$FEDSRVTestLabel.Text = "FED SRV:"
$FEDSRVTestLabel.TabStop = $false
$FEDSRVTestLabel.Anchor = [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Top
$FEDSRVTestLabel.Add_Click(
{
	if($FEDSRVCheckBox.Checked -eq $true)
    {
        $PortTextBox.Enabled = $true
		$PortTextBox.text = "443"
		$HostNameLabel.Text = "FQDN / IP:"
		$SIPSRVCheckBox.Checked = $false
		$SIPINTSRVCheckBox.Checked = $false
		$FEDSRVCheckBox.Checked = $false
    }
	else
	{
		$PortTextBox.Enabled = $false
		$PortTextBox.text = ""
		$HostNameLabel.Text = "SIP Domain:"
		$SIPSRVCheckBox.Checked = $false
		$SIPINTSRVCheckBox.Checked = $false
		$FEDSRVCheckBox.Checked = $true
	}
})
$objForm.Controls.Add($FEDSRVTestLabel)


# FED SRV Check Box ============================================================
$FEDSRVCheckBox = New-Object System.Windows.Forms.Checkbox 
$FEDSRVCheckBox.Location = New-Object System.Drawing.Size(90,28) 
$FEDSRVCheckBox.Size = New-Object System.Drawing.Size(20,20)
$FEDSRVCheckBox.TabStop = $false
#$SIPSRVCheckBox.Anchor = [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Bottom
$FEDSRVCheckBox.Add_Click(
{
	if($FEDSRVCheckBox.Checked -eq $true)
    {
        $PortTextBox.Enabled = $false
		$PortTextBox.text = ""
		$HostNameLabel.Text = "SIP Domain:"
		$SIPSRVCheckBox.Checked = $false
		$SIPINTSRVCheckBox.Checked = $false
    }
	else
	{
		$PortTextBox.Enabled = $true
		$PortTextBox.text = "443"
		$HostNameLabel.Text = "FQDN / IP:"
		$SIPSRVCheckBox.Checked = $false
		$SIPINTSRVCheckBox.Checked = $false
	}
})
$objForm.Controls.Add($FEDSRVCheckBox) 
$FEDSRVCheckBox.Checked = $false


#SRVTest Label =================================================================
$SIPSRVTestLabel = New-Object System.Windows.Forms.Label
$SIPSRVTestLabel.Location = New-Object System.Drawing.Size(120,30) 
$SIPSRVTestLabel.Size = New-Object System.Drawing.Size(57,15) 
$SIPSRVTestLabel.Text = "SIP SRV:"
$SIPSRVTestLabel.TabStop = $false
$SIPSRVTestLabel.Anchor = [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Top
$SIPSRVTestLabel.Add_Click(
{
	if($SIPSRVCheckBox.Checked -eq $true)
    {
        $PortTextBox.Enabled = $true
		$PortTextBox.text = "443"
		$HostNameLabel.Text = "FQDN / IP:"
		$FEDSRVCheckBox.Checked = $false
		$SIPINTSRVCheckBox.Checked = $false
		$SIPSRVCheckBox.Checked = $false
    }
	else
	{
		$PortTextBox.Enabled = $false
		$PortTextBox.text = ""
		$HostNameLabel.Text = "SIP Domain:"
		$FEDSRVCheckBox.Checked = $false
		$SIPINTSRVCheckBox.Checked = $false
		$SIPSRVCheckBox.Checked = $true
	}
})
$objForm.Controls.Add($SIPSRVTestLabel)


# SIP SRV Check Box ============================================================
$SIPSRVCheckBox = New-Object System.Windows.Forms.Checkbox 
$SIPSRVCheckBox.Location = New-Object System.Drawing.Size(177,28) 
$SIPSRVCheckBox.Size = New-Object System.Drawing.Size(20,20)
$SIPSRVCheckBox.TabStop = $false
#$SIPSRVCheckBox.Anchor = [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Bottom
$SIPSRVCheckBox.Add_Click(
{
	if($SIPSRVCheckBox.Checked -eq $true)
    {
        $PortTextBox.Enabled = $false
		$PortTextBox.text = ""
		$HostNameLabel.Text = "SIP Domain:"
		$FEDSRVCheckBox.Checked = $false
		$SIPINTSRVCheckBox.Checked = $false
    }
	else
	{
		$PortTextBox.Enabled = $true
		$PortTextBox.text = "443"
		$HostNameLabel.Text = "FQDN / IP:"
		$FEDSRVCheckBox.Checked = $false
		$SIPINTSRVCheckBox.Checked = $false
	}
})
$objForm.Controls.Add($SIPSRVCheckBox) 
$SIPSRVCheckBox.Checked = $false


#SIPINT SRVTest Label =================================================================
$SIPINTSRVTestLabel = New-Object System.Windows.Forms.Label
$SIPINTSRVTestLabel.Location = New-Object System.Drawing.Size(205,30) 
$SIPINTSRVTestLabel.Size = New-Object System.Drawing.Size(75,15) 
$SIPINTSRVTestLabel.Text = "SIP INT SRV:"
$SIPINTSRVTestLabel.TabStop = $false
$SIPINTSRVTestLabel.Anchor = [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Top
$SIPINTSRVTestLabel.Add_Click(
{
	if($SIPINTSRVCheckBox.Checked -eq $true)
    {
        $PortTextBox.Enabled = $true
		$PortTextBox.text = "443"
		$HostNameLabel.Text = "FQDN / IP:"
		$FEDSRVCheckBox.Checked = $false
		$SIPSRVCheckBox.Checked = $false
		$SIPINTSRVCheckBox.Checked = $false
    }
	else
	{
		$PortTextBox.Enabled = $false
		$PortTextBox.text = ""
		$HostNameLabel.Text = "SIP Domain:"
		$FEDSRVCheckBox.Checked = $false
		$SIPSRVCheckBox.Checked = $false
		$SIPINTSRVCheckBox.Checked = $true
	}
})
$objForm.Controls.Add($SIPINTSRVTestLabel)


# SIP INT SRV Check Box ============================================================
$SIPINTSRVCheckBox = New-Object System.Windows.Forms.Checkbox 
$SIPINTSRVCheckBox.Location = New-Object System.Drawing.Size(280,28) 
$SIPINTSRVCheckBox.Size = New-Object System.Drawing.Size(20,20)
$SIPINTSRVCheckBox.TabStop = $false
#$SIPSRVCheckBox.Anchor = [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Bottom
$SIPINTSRVCheckBox.Add_Click(
{
	if($SIPINTSRVCheckBox.Checked -eq $true)
    {
        $PortTextBox.Enabled = $false
		$PortTextBox.text = ""
		$HostNameLabel.Text = "SIP Domain:"
		$FEDSRVCheckBox.Checked = $false
		$SIPSRVCheckBox.Checked = $false
    }
	else
	{
		$PortTextBox.Enabled = $true
		$PortTextBox.text = "443"
		$HostNameLabel.Text = "FQDN / IP:"
		$FEDSRVCheckBox.Checked = $false
		$SIPSRVCheckBox.Checked = $false
	}
})
$objForm.Controls.Add($SIPINTSRVCheckBox) 
$SIPINTSRVCheckBox.Checked = $false



#Host Name Label ============================================================
$HostNameLabel = New-Object System.Windows.Forms.Label
$HostNameLabel.Location = New-Object System.Drawing.Size(5,52) 
$HostNameLabel.Size = New-Object System.Drawing.Size(80,15) 
$HostNameLabel.Text = "FQDN / IP:"
$HostNameLabel.TextAlign = [System.Drawing.ContentAlignment]::TopRight
$HostNameLabel.Anchor = [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Top
$HostNameLabel.TabStop = $false
$objForm.Controls.Add($HostNameLabel)

#Host Name Text box ============================================================
$HostNameTextBox = new-object System.Windows.Forms.textbox
$HostNameTextBox.location = new-object system.drawing.size(90,50)
$HostNameTextBox.size= new-object system.drawing.size(220,15)
$HostNameTextBox.text = "sip.microsoft.com"
$HostNameTextBox.TabIndex = 1
$HostNameTextBox.Anchor = [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Top
$objform.controls.add($HostNameTextBox)
$HostNameTextBox.add_KeyUp(
{
	if ($_.KeyCode -eq "Enter") 
	{	
		Add-Domain
	}
})


#Host Name Label ============================================================
$PortLabel = New-Object System.Windows.Forms.Label
$PortLabel.Location = New-Object System.Drawing.Size(57,72) 
$PortLabel.Size = New-Object System.Drawing.Size(32,15) 
$PortLabel.Text = "Port: "
$PortLabel.TabStop = $false
$PortLabel.Anchor = [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Top
$objForm.Controls.Add($PortLabel)

#Host Name Text box ============================================================
$PortTextBox = new-object System.Windows.Forms.textbox
$PortTextBox.location = new-object system.drawing.size(90,70)
$PortTextBox.size= new-object system.drawing.size(220,15)
$PortTextBox.text = "443"
$PortTextBox.TabIndex = 2
$PortTextBox.Anchor = [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Top
$objform.controls.add($PortTextBox)
$PortTextBox.add_KeyUp(
{
	if ($_.KeyCode -eq "Enter") 
	{	
		Add-Domain
	}
})
$PortTextBox.Enabled = $true



#Domain Add button
$DomainNameAddButton = New-Object System.Windows.Forms.Button
$DomainNameAddButton.Location = New-Object System.Drawing.Size(90,95)
$DomainNameAddButton.Size = New-Object System.Drawing.Size(90,18)
$DomainNameAddButton.Text = "Add"
$DomainNameAddButton.TabStop = $false
$DomainNameAddButton.Anchor = [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Top
$DomainNameAddButton.Add_Click(
{
	Add-Domain
})
$objForm.Controls.Add($DomainNameAddButton)



$ImportCSVButton = New-Object System.Windows.Forms.Button
$ImportCSVButton.Location = New-Object System.Drawing.Size(200,95)
$ImportCSVButton.Size = New-Object System.Drawing.Size(90,18)
$ImportCSVButton.Text = "Import..."
$ImportCSVButton.TabStop = $false
$ImportCSVButton.Anchor = [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Top
$ImportCSVButton.Add_Click(
{
	if($MajorVersion -ne  "2")
	{
		$Filter="CSV Files (*.csv*)|*.csv*"
		[System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") | Out-Null
		$objDialog = New-Object System.Windows.Forms.OpenFileDialog
		$objDialog.InitialDirectory = $shortname
		$objDialog.Filter = $Filter
		$objDialog.Title = "Open CSV File"
		#$objDialog.CheckFileExists = $false
		$Show = $objDialog.ShowDialog()
		if ($Show -eq "OK")
		{
			[string]$InputFilename = $objDialog.FileName
		}
		else
		{
			[string]$InputFilename = ""
			Write-Host "ERROR: No file provided." -foreground "red"
			return
		}

		if(Test-Path $InputFilename)
		{
			Write-Host "Starting Port Testing." -foreground "yellow"

			$Records = Import-Csv $InputFilename
			
			if($Records.length -eq 0)
			{
				Write-Host "ERROR: The file supplied is either not formatted correctly, or does not have any information in it." -foreground "Red"
				return
			}
				
			foreach($Record in $Records)
			{
				
				[string]$DomainName = $Record.Domain
				[string]$DomainType = $Record.Type
				[string]$DomainPort = $Record.Port
				#[string]$FQDNName = $Record.ProxyFqdn
				
				Write-Host "IMPORTING: $DomainName $DomainType $DomainPort" -foreground "yellow"

				$DomainNameListboxItem = new-object System.Windows.Forms.ListViewItem("$DomainName")
				
				if($DomainPort -ne "" -and ($DomainType -ne "" -and $DomainType -ne "DIR"))
				{
					Write-Host "ERROR: Cannot have both Port and Domain type specified at the same time. Not importing $DomainName." -foreground "red"
				}
				else
				{
					$PortError = $false
					$TypeError = $false
					if($DomainPort -match "^(6553[0-5])$|^(655[0-2]\d)$|^(65[0-4]\d{2})$|^(6[0-4]\d{3})$|^([1-5]\d{4})$|^([1-9]\d{1,3})$|^(\d{1})$" -or $DomainPort -eq "")
					{
						[void]$DomainNameListboxItem.SubItems.Add($DomainPort)
					}
					else
					{
						Write-Host "ERROR: Port $DomainPort is not in range 1-65535" -foreground "red"
						$PortError = $true
					}
					
					if($DomainType -eq "SIP" -or $DomainType -eq "SIPINT" -or $DomainType -eq "FED" -or $DomainType -eq "DIR" -or $DomainType -eq "")
					{
						[void]$DomainNameListboxItem.SubItems.Add($DomainType)
					}
					else
					{
						Write-Host "ERROR: Invalid Domain Type $DomainType. Must be either SIP or FED." -foreground "red"
						$TypeError = $true
					}
					
					if($DomainPort -eq "" -and ($DomainType -eq "" -or $DomainType -eq "DIR"))
					{
						Write-Host "ERROR: Both Port and DomainType cannot be blank. Not importing $DomainName." -foreground "red"
					}
					elseif($PortError -or $TypeError)
					{
						#Do nothing, error already reported
					}
					else
					{
						$DomainNameListboxItem.ForeColor = "Black"
						[void]$DomainNameListbox.Items.Add($DomainNameListboxItem)
					}
				}

			}

		}
	}
	else
	{
		Write-Host "ERROR: This feature is not supported on Powershell Version 2" -foreground "red"
	}
})
$objForm.Controls.Add($ImportCSVButton)


#Domain Remove button
$DomainNameRemoveButton = New-Object System.Windows.Forms.Button
$DomainNameRemoveButton.Location = New-Object System.Drawing.Size(100,280)
$DomainNameRemoveButton.Size = New-Object System.Drawing.Size(70,20)
$DomainNameRemoveButton.Text = "Delete"
$DomainNameRemoveButton.TabStop = $false
$DomainNameRemoveButton.Anchor = [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Top
$DomainNameRemoveButton.Add_Click(
{
	$CurrentSelectedItem = $DomainNameListbox.Items.IndexOf($DomainNameListbox.SelectedItems[0])
		
	if($CurrentSelectedItem -ge 0)
	{
		[void]$DomainNameListbox.Items.Remove($DomainNameListbox.SelectedItems[0])
		
		if($CurrentSelectedItem -ge 1)
		{
			$index = $CurrentSelectedItem - 1
			$DomainNameListbox.Items[$index].Selected = $true
			$DomainNameListbox.Select()
		}
		else
		{
			if($DomainNameListbox.Items.Count -ne 0)
			{
				$DomainNameListbox.Items[0].Selected = $true
				$DomainNameListbox.Select()
			}
		}
	}
})
$objForm.Controls.Add($DomainNameRemoveButton)


#Domain Clear button
$DomainNameClearButton = New-Object System.Windows.Forms.Button
$DomainNameClearButton.Location = New-Object System.Drawing.Size(180,280)
$DomainNameClearButton.Size = New-Object System.Drawing.Size(70,20)
$DomainNameClearButton.Text = "Clear All"
$DomainNameClearButton.TabStop = $false
$DomainNameClearButton.Anchor = [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Top
$DomainNameClearButton.Add_Click(
{
		[void]$DomainNameListbox.Items.Clear()
})
$objForm.Controls.Add($DomainNameClearButton)


# Listbox of FQDNS ============================================================
$DomainNameListbox = New-Object windows.forms.ListView
$DomainNameListbox.View = [System.Windows.Forms.View]"Details"
$DomainNameListbox.Size = New-Object System.Drawing.Size(293,155)
$DomainNameListbox.Location = New-Object System.Drawing.Size(30,120)
$DomainNameListbox.FullRowSelect = $true
$DomainNameListbox.GridLines = $true
$DomainNameListbox.HideSelection = $false
$DomainNameListbox.MultiSelect = $false
$DomainNameListbox.Sorting = [System.Windows.Forms.SortOrder]"Ascending"
[void]$DomainNameListbox.Columns.Add("Domain / FQDN / IP Address", 170)
[void]$DomainNameListbox.Columns.Add("Port", 50)
[void]$DomainNameListbox.Columns.Add("Type", 50)
$DomainNameListbox.Anchor = [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Top
$DomainNameListbox.TabStop = $false
$objForm.Controls.Add($DomainNameListbox)


$CSVInfoLabel = New-Object System.Windows.Forms.Label
$CSVInfoLabel.Location = New-Object System.Drawing.Size(340,343) 
$CSVInfoLabel.Size = New-Object System.Drawing.Size(70,15) 
$CSVInfoLabel.Text = "Output CSV:"
$CSVInfoLabel.TabStop = $false
$CSVInfoLabel.Anchor = [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Bottom
$CSVInfoLabel.Add_Click(
{
	if($OutFileCheckBox.checked)
	{
		$OutFileCheckBox.checked = $false
	}
	else
	{
		$OutFileCheckBox.checked = $true
	}
})
$objForm.Controls.Add($CSVInfoLabel)

# Add OutFileCheckBox ============================================================
$OutFileCheckBox = New-Object System.Windows.Forms.Checkbox 
$OutFileCheckBox.Location = New-Object System.Drawing.Size(415,340) 
$OutFileCheckBox.Size = New-Object System.Drawing.Size(20,20)
$OutFileCheckBox.TabStop = $false
$OutFileCheckBox.Anchor = [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Bottom
$objForm.Controls.Add($OutFileCheckBox) 

#File Text box ============================================================
$FileLocationTextBox = New-Object System.Windows.Forms.TextBox
$FileLocationTextBox.location = new-object system.drawing.size(437,340)
$FileLocationTextBox.size = new-object system.drawing.size(320,23)
$FileLocationTextBox.tabIndex = 2
$FileLocationTextBox.text = "${shortname}\CertificateExport.csv" 
$FileLocationTextBox.Anchor = [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right -bor [System.Windows.Forms.AnchorStyles]::Bottom
$DomainNameListbox.TabStop = $false
$objform.controls.add($FileLocationTextBox)
$FileLocationTextBox.SelectionStart = $FileLocationTextBox.Text.Length
$FileLocationTextBox.ScrollToCaret()


#File Browse button
$BrowseButton = New-Object System.Windows.Forms.Button
$BrowseButton.Location = New-Object System.Drawing.Size(765,340)
$BrowseButton.Size = New-Object System.Drawing.Size(70,18)
$BrowseButton.Text = "Browse..."
$BrowseButton.Anchor = [System.Windows.Forms.AnchorStyles]::Right -bor [System.Windows.Forms.AnchorStyles]::Bottom
$BrowseButton.Add_Click(
{
	
	if($MajorVersion -ne  "2")
	{
		#File Dialog
		[string] $pathVar = "C:\"
		$Filter="All Files (*.*)|*.*"
		[System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") | Out-Null
		$objDialog = New-Object System.Windows.Forms.SaveFileDialog
		#$objDialog.InitialDirectory = 
		$objDialog.FileName = "CertificateExport.csv"
		$objDialog.Filter = $Filter
		$objDialog.Title = "Export File Name"
		$objDialog.CheckFileExists = $false
		$Show = $objDialog.ShowDialog()
		if ($Show -eq "OK")
		{
			[string]$content = ""
			$FileLocationTextBox.text = $objDialog.FileName
			$FileLocationTextBox.SelectionStart = $FileLocationTextBox.Text.Length
			$FileLocationTextBox.ScrollToCaret()
		}
		else
		{
			return
		}
	}
	else
	{
		Write-Host "ERROR: You will not be able to use the browse button on Powershell version 2. Manually type the file location." -foreground "red"
	}
	
})
$objForm.Controls.Add($BrowseButton)

# DNSServer label ============================================================
$DNSServerLabel = New-Object System.Windows.Forms.Label
$DNSServerLabel.Location = New-Object System.Drawing.Size(340,368) 
$DNSServerLabel.Size = New-Object System.Drawing.Size(70,15) 
$DNSServerLabel.Text = "DNS Server:"
$DNSServerLabel.TabStop = $false
$DNSServerLabel.Anchor = [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Bottom
$objForm.Controls.Add($DNSServerLabel)

#DNSServer box ============================================================
$DNSServerTextBox = New-Object System.Windows.Forms.TextBox
$DNSServerTextBox.location = new-object system.drawing.size(437,365)
$DNSServerTextBox.size = new-object system.drawing.size(150,23)
$DNSServerTextBox.tabIndex = 15
$DNSServerTextBox.text = "8.8.8.8" 
$DNSServerTextBox.Anchor = [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Bottom
$objform.controls.add($DNSServerTextBox)
#$DNSServerTextBox.SelectionStart = $FileLocationTextBox.Text.Length
#$DNSServerTextBox.ScrollToCaret()

$DNSServerTextBox.Add_LostFocus(
{
	$DNSServerIPTest = $DNSServerTextBox.text
	if($DNSServerIPTest -match "^(([01]?\d?\d|2[0-4]\d|25[0-5])\.){3}([01]?\d?\d|2[0-4]\d|25[0-5])$")
	{
		#DO NOTHING
	}
	else
	{
		Write-Host "ERROR: The value entered is not an IP Address. Changing back to 8.8.8.8" -foreground "red"
		$DNSServerTextBox.text = "8.8.8.8"
	}
})



$ShowAdvancedLabel = New-Object System.Windows.Forms.Label
$ShowAdvancedLabel.Location = New-Object System.Drawing.Size(105,310) 
$ShowAdvancedLabel.Size = New-Object System.Drawing.Size(100,15) 
$ShowAdvancedLabel.Text = "Show Advanced:"
$ShowAdvancedLabel.TabStop = $false
$ShowAdvancedLabel.TextAlign = [System.Drawing.ContentAlignment]::TopRight
#$ShowAdvancedLabel.Anchor = [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Bottom
$ShowAdvancedLabel.Add_Click(
{
	if($ShowAdvancedCheckBox.checked)
	{
		$ShowAdvancedCheckBox.checked = $false
	}
	else
	{
		$ShowAdvancedCheckBox.checked = $true
	}
})
$objForm.Controls.Add($ShowAdvancedLabel)


# Show Advanced ============================================================
$ShowAdvancedCheckBox = New-Object System.Windows.Forms.Checkbox 
$ShowAdvancedCheckBox.Location = New-Object System.Drawing.Size(210,307) 
$ShowAdvancedCheckBox.Size = New-Object System.Drawing.Size(20,20)
$ShowAdvancedCheckBox.TabStop = $false
#$SIPSRVCheckBox.Anchor = [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Bottom
$ShowAdvancedCheckBox.Add_Click(
{

})
$objForm.Controls.Add($ShowAdvancedCheckBox) 
$ShowAdvancedCheckBox.checked = $false

$ShowChainLabel = New-Object System.Windows.Forms.Label
$ShowChainLabel.Location = New-Object System.Drawing.Size(105,330) 
$ShowChainLabel.Size = New-Object System.Drawing.Size(100,15) 
$ShowChainLabel.Text = "Show Root Chain:"
$ShowChainLabel.TextAlign = [System.Drawing.ContentAlignment]::TopRight
$ShowChainLabel.TabStop = $false
$ShowChainLabel.Add_Click(
{
	if($ShowChainCheckBox.checked)
	{
		$ShowChainCheckBox.checked = $false
	}
	else
	{
		$ShowChainCheckBox.checked = $true
	}
})
#$ShowAdvancedLabel.Anchor = [System.Windows.Forms.AnchorStyles]::Left
$objForm.Controls.Add($ShowChainLabel)


# Show Advanced ============================================================
$ShowChainCheckBox = New-Object System.Windows.Forms.Checkbox 
$ShowChainCheckBox.Location = New-Object System.Drawing.Size(210,327) 
$ShowChainCheckBox.Size = New-Object System.Drawing.Size(20,20)
$ShowChainCheckBox.TabStop = $false
#$SIPSRVCheckBox.Anchor = [System.Windows.Forms.AnchorStyles]::Left
$ShowChainCheckBox.Add_Click(
{

})
$objForm.Controls.Add($ShowChainCheckBox) 
$ShowChainCheckBox.checked = $false


$TestPoolLabel = New-Object System.Windows.Forms.Label
$TestPoolLabel.Location = New-Object System.Drawing.Size(105,350) 
$TestPoolLabel.Size = New-Object System.Drawing.Size(100,15) 
$TestPoolLabel.Text = "Test DNSLB Pool:"
$TestPoolLabel.TextAlign = [System.Drawing.ContentAlignment]::TopRight
$TestPoolLabel.TabStop = $false
#$ShowAdvancedLabel.Anchor = [System.Windows.Forms.AnchorStyles]::Left
$TestPoolLabel.Add_Click(
{
	if($TestPoolCheckBox.checked)
	{
		$TestPoolCheckBox.checked = $false
	}
	else
	{
		$TestPoolCheckBox.checked = $true
	}
})
$objForm.Controls.Add($TestPoolLabel)


# Show Advanced ============================================================
$TestPoolCheckBox = New-Object System.Windows.Forms.Checkbox 
$TestPoolCheckBox.Location = New-Object System.Drawing.Size(210,347) 
$TestPoolCheckBox.Size = New-Object System.Drawing.Size(20,20)
$TestPoolCheckBox.TabStop = $false
#$SIPSRVCheckBox.Anchor = [System.Windows.Forms.AnchorStyles]::Left
$TestPoolCheckBox.Add_Click(
{

})
$objForm.Controls.Add($TestPoolCheckBox) 
$TestPoolCheckBox.checked = $true


# Test button
$TestButton = New-Object System.Windows.Forms.Button
$TestButton.Location = New-Object System.Drawing.Size(100,370)
$TestButton.Size = New-Object System.Drawing.Size(150,25)
$TestButton.Text = "Test"
$TestButton.TabIndex = 6
$TestButton.Anchor = [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Top
$TestButton.Add_Click(
{
	$script:CancelScan = $false
	$TestButton.Visible = $false
	$CancelTestButton.Visible = $true
	$TestButton.Enabled = $false
	$DomainNameRemoveButton.Enabled = $false
	$DomainNameClearButton.Enabled = $false
	$DomainNameAddButton.Enabled = $false
	$ImportCSVButton.Enabled = $false
	$OutFileCheckBox.Enabled = $false
	$BrowseButton.Enabled = $false
	$StatusLabel.Text = "Testing..."
	
	$script:RTFDisplayString = ""
	$InformationTextBox.Clear()
	$script:FirstCheck = $true
	[System.Windows.Forms.Application]::DoEvents()
	
	
	foreach($DomainName in $DomainNameListbox.Items)
	{
		[System.Windows.Forms.Application]::DoEvents()
	
		if($script:CancelScan)
		{
			Write-Host "Cancelling Operation..."
			$TestButton.Enabled = $true
			#Start-Sleep -s 1
			break
		}
		
		$SubItems = $DomainName.SubItems
		[string]$DomainNameString = $SubItems[0].Text
		[string]$PortString = $SubItems[1].Text
		[string]$TypeString = $SubItems[2].Text
		
		Write-Host "INFO: Looking up DNS for $DomainNameString  $PortString  $TypeString"
		
		if($DNSServerTextBox.text -match "^(([01]?\d?\d|2[0-4]\d|25[0-5])\.){3}([01]?\d?\d|2[0-4]\d|25[0-5])$")
		{
			$DNSServerIP = $DNSServerTextBox.text
		}
		else
		{
			Write-Host "The DNS Server IP Address is not in a supported format. Using 8.8.8.8 instead." -foreground "red"
		}
		
		if($PortString -eq "")
		{
			if($TypeString -eq "SIP")
			{
				$SRVName = "_sip._tls.${DomainNameString}"
				$DNSResultArray = GetDnsSrv $SRVName $DNSServerIP "SRV"
				
				foreach($DNSResult in $DNSResultArray)
				{
					$DNSResult | FT
					
					Write-Host "DNS Result:" $DNSResult.NameHost
					Write-Host "DNS Result:" $DNSResult.Port
					Write-Host
					
					[string]$Hostname = $DNSResult.NameHost
					[string]$PortString = $DNSResult.Port
					
					Check-DNSRecord "SIP $DomainNameString" $Hostname $PortString
					
					[System.Windows.Forms.Application]::DoEvents()
					
					if($script:CancelScan)
					{
						Write-Host "Cancelling Operation..."
						#Start-Sleep -s 1
						$TestButton.Enabled = $true
						break
					}
				}
			}
			elseif($TypeString -eq "FED")
			{
				$SRVName = "_sipfederationtls._tcp.${DomainNameString}"
				$DNSResultArray = GetDnsSrv $SRVName $DNSServerIP "SRV"
				
				foreach($DNSResult in $DNSResultArray)
				{
					$DNSResult | FT
					
					Write-Host "DNS Result:" $DNSResult.NameHost
					Write-Host "DNS Result:" $DNSResult.Port
					Write-Host
					
					[string]$Hostname = $DNSResult.NameHost
					[string]$PortString = $DNSResult.Port
					
					Check-DNSRecord "FED $DomainNameString" $Hostname $PortString
					
					[System.Windows.Forms.Application]::DoEvents()
					
					if($script:CancelScan)
					{
						Write-Host "Cancelling Operation..."
						#Start-Sleep -s 1
						$TestButton.Enabled = $true
						break
					}
				}
			}
			if($TypeString -eq "SIPINT")
			{
				$SRVName = "_sipinternaltls._tcp.${DomainNameString}"
				$DNSResultArray = GetDnsSrv $SRVName $DNSServerIP "SRV"
				
				foreach($DNSResult in $DNSResultArray)
				{
					$DNSResult | FT
					
					Write-Host "DNS Result:" $DNSResult.NameHost
					Write-Host "DNS Result:" $DNSResult.Port
					Write-Host
					
					[string]$Hostname = $DNSResult.NameHost
					[string]$PortString = $DNSResult.Port
					
					Check-DNSRecord "SIPINT $DomainNameString" $Hostname $PortString
					
					[System.Windows.Forms.Application]::DoEvents()
					
					if($script:CancelScan)
					{
						Write-Host "Cancelling Operation..."
						#Start-Sleep -s 1
						$TestButton.Enabled = $true
						break
					}
				}
			}
		}
		else
		{
			[string]$Hostname = $DomainNameString
			
			Check-DNSRecord "" $Hostname $PortString
			
			
			$InformationTextBox.SelectionStart = $InformationTextBox.Text.length
			$InformationTextBox.ScrollToCaret()
			[System.Windows.Forms.Application]::DoEvents()
			
			Start-Sleep -s 2
			
		}
		
		$InformationTextBox.Text += "`r`n"
		$Script:RTFDisplayString += " \line "
		Start-Sleep -s 1
	}
		
		
	$InformationTextBox.Rtf = $script:RTFStart + $Script:RTFDisplayString + $script:RTFEnd
	$script:FirstCheck = $true
	$TestButton.Enabled = $true
	$DomainNameRemoveButton.Enabled = $true
	$DomainNameClearButton.Enabled = $true
	$DomainNameAddButton.Enabled = $true
	$ImportCSVButton.Enabled = $true
	$OutFileCheckBox.Enabled = $true
	$BrowseButton.Enabled = $true
	$TestButton.Visible = $true
	$CancelTestButton.Visible = $false
	$StatusLabel.Text = ""
})
$objForm.Controls.Add($TestButton)


# Test button
$CancelTestButton = New-Object System.Windows.Forms.Button
$CancelTestButton.Location = New-Object System.Drawing.Size(100,370)
$CancelTestButton.Size = New-Object System.Drawing.Size(150,25)
$CancelTestButton.Text = "Cancel Test"
$CancelTestButton.ForeColor = "red"
$CancelTestButton.TabStop = $false
$CancelTestButton.Anchor = [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Top
$CancelTestButton.Add_Click(
{
	$script:CancelScan = $true
	
	$DomainNameRemoveButton.Enabled = $true
	$DomainNameClearButton.Enabled = $true
	$DomainNameAddButton.Enabled = $true
	#$DNSServerTextBox.Enabled = $true
	$CancelTestButton.Visible = $false
	$TestButton.Visible = $true
	
})
$objForm.Controls.Add($CancelTestButton)


$objInfoLabel = New-Object System.Windows.Forms.Label
$objInfoLabel.Location = New-Object System.Drawing.Size(340,15) 
$objInfoLabel.Size = New-Object System.Drawing.Size(100,15) 
$objInfoLabel.Text = "Information:"
$objInfoLabel.TabStop = $false
$objForm.Controls.Add($objInfoLabel)



#Info Box
$FontCourier = new-object System.Drawing.Font("Courier New",9,[Drawing.FontStyle]'Regular')
$InformationTextBox = New-Object System.Windows.Forms.RichTextBox 
$InformationTextBox.Location = New-Object System.Drawing.Size(340,30)
$InformationTextBox.Size = New-Object System.Drawing.Size(510,300)  
$InformationTextBox.Font = $FontCourier
$InformationTextBox.Multiline = $True	
#$InformationTextBox.ScrollBars = [System.Windows.Forms.RichTextBoxScrollBars]::Vertical
#$InformationTextBox.ScrollBars = [System.Windows.Forms.RichTextBoxScrollBars]::Horizontal
$InformationTextBox.ReadOnly = $true
$InformationTextBox.Wordwrap = $false
$InformationTextBox.BackColor = [System.Drawing.Color]::White
$InformationTextBox.Text = ""
$InformationTextBox.TabStop = $false
$InformationTextBox.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Right -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Bottom
$InformationTextBox.Add_LinkClicked(
{
	Write-Host "LINK CLICKED: " $_.LinkText
	[system.Diagnostics.Process]::start($_.LinkText)
})

$objForm.Controls.Add($InformationTextBox) 


# Add the Status Label ============================================================
$StatusLabel = New-Object System.Windows.Forms.Label
$StatusLabel.Location = New-Object System.Drawing.Size(10,400) 
$StatusLabel.Size = New-Object System.Drawing.Size(420,15) 
$StatusLabel.Text = ""
$StatusLabel.forecolor = "red"
$StatusLabel.TabStop = $false
$StatusLabel.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left
$objForm.Controls.Add($StatusLabel)



# Tool tips for Help!
$ToolTip = New-Object System.Windows.Forms.ToolTip 
$ToolTip.BackColor = [System.Drawing.Color]::LightGoldenrodYellow 
$ToolTip.IsBalloon = $true 
$ToolTip.InitialDelay = 500 
$ToolTip.ReshowDelay = 500 
$ToolTip.AutoPopDelay = 10000
$ToolTip.SetToolTip($FEDSRVCheckBox, "Selecting this checkbox will force a SRV lookup on _sipfederationtls._tcp.<SIP Domain>") 
$ToolTip.SetToolTip($FEDSRVTestLabel, "Selecting this checkbox will force a SRV lookup on _sipfederationtls._tcp.<SIP Domain>") 
$ToolTip.SetToolTip($SIPSRVCheckBox, "Selecting this checkbox will force a SRV lookup on _sip._tls.<SIP Domain>") 
$ToolTip.SetToolTip($SIPSRVTestLabel, "Selecting this checkbox will force a SRV lookup on _sip._tls.<SIP Domain>") 
$ToolTip.SetToolTip($SIPINTSRVCheckBox, "Selecting this checkbox will force a SRV lookup on _sipinternaltls._tcp.<SIP Domain>.`r`nNote: This is an internal DNS record so set the DNS Server to the internal server.") 
$ToolTip.SetToolTip($SIPINTSRVTestLabel, "Selecting this checkbox will force a SRV lookup on _sipinternaltls._tcp.<SIP Domain>.`r`nNote: This is an internal DNS record so set the DNS Server to the internal server.") 
$ToolTip.SetToolTip($PortTextBox, "Enter the port number of the FQDN/IP location that you would like to see the certificate for")
$ToolTip.SetToolTip($TestPoolCheckBox, "If the A Record lookup contains multiple servers then automatically test all servers resolved")
$ToolTip.SetToolTip($TestPoolLabel, "If the A Record lookup contains multiple servers then automatically test all servers resolved")
$ToolTip.SetToolTip($ShowChainCheckBox, "Show the full certificate chain in the Information window")
$ToolTip.SetToolTip($ShowChainLabel, "Show the full certificate chain in the Information window")
$ToolTip.SetToolTip($ShowAdvancedCheckBox, "Show full certificate details in the Information window")
$ToolTip.SetToolTip($ShowAdvancedLabel, "Show full certificate details in the Information window")
$ToolTip.SetToolTip($OutFileCheckBox, "When the test is run also produce a CSV file with all the returned certificate data in it")
$ToolTip.SetToolTip($CSVInfoLabel, "When the test is run also produce a CSV file with all the returned certificate data in it")
$ToolTip.SetToolTip($DNSServerLabel, "Specify your own DNS server for resolving DNS records")
$ToolTip.SetToolTip($DNSServerTextBox, "Specify your own DNS server for resolving DNS records")




function Add-Domain
{
	if($HostNameTextBox.Text -ne "")
	{
		if($HostNameTextBox.Text -match ".*,.*")
		{
			$Sections = $HostNameTextBox.Text -split ","
			$PortString = $PortTextBox.text.Trim()
			
			
			foreach($Section in $Sections)
			{
				[string]$Name = $Section.Trim()
				
				if($PortString -match ".*,.*")
				{
					$PortSections = $PortString -split ","
					foreach($PortSection in $PortSections)
					{
						[string]$testPort = $PortSection.Trim()
						[string]$Name = $Section.Trim()
											
						
						if($SIPSRVCheckBox.Checked)
						{
							$DomainNameListboxItem = new-object System.Windows.Forms.ListViewItem("${Name}")
							[void]$DomainNameListboxItem.SubItems.Add("")
							[void]$DomainNameListboxItem.SubItems.Add("SIP")
							$DomainNameListboxItem.ForeColor = "Black"
							[void]$DomainNameListbox.Items.Add($DomainNameListboxItem)
						}
						elseif($SIPINTSRVCheckBox.Checked)
						{
							$DomainNameListboxItem = new-object System.Windows.Forms.ListViewItem("${Name}")
							[void]$DomainNameListboxItem.SubItems.Add("")
							[void]$DomainNameListboxItem.SubItems.Add("SIPINT")
							$DomainNameListboxItem.ForeColor = "Black"
							[void]$DomainNameListbox.Items.Add($DomainNameListboxItem)
						}
						elseif($FEDSRVCheckBox.Checked)
						{
							$DomainNameListboxItem = new-object System.Windows.Forms.ListViewItem("${Name}")
							[void]$DomainNameListboxItem.SubItems.Add("")
							[void]$DomainNameListboxItem.SubItems.Add("FED")
							$DomainNameListboxItem.ForeColor = "Black"
							[void]$DomainNameListbox.Items.Add($DomainNameListboxItem)
						}
						elseif($testPort -ne "" -and $testPort -match "^(6553[0-5])$|^(655[0-2]\d)$|^(65[0-4]\d{2})$|^(6[0-4]\d{3})$|^([1-5]\d{4})$|^([1-9]\d{1,3})$|^(\d{1})$")
						{
							$DomainNameListboxItem = new-object System.Windows.Forms.ListViewItem("${Name}")
							[void]$DomainNameListboxItem.SubItems.Add($testPort)
							[void]$DomainNameListboxItem.SubItems.Add("DIR")
							$DomainNameListboxItem.ForeColor = "Black"
							[void]$DomainNameListbox.Items.Add($DomainNameListboxItem)
						}
						else
						{
							Write-Host "ERROR: Incorrect port number value." -foreground "red"
						}
					}
				}
				else
				{
					if($SIPSRVCheckBox.Checked)
					{
						$DomainNameListboxItem = new-object System.Windows.Forms.ListViewItem("${Name}")
						[void]$DomainNameListboxItem.SubItems.Add("")
						[void]$DomainNameListboxItem.SubItems.Add("SIP")
						$DomainNameListboxItem.ForeColor = "Black"
						[void]$DomainNameListbox.Items.Add($DomainNameListboxItem)
					}
					elseif($SIPINTSRVCheckBox.Checked)
					{
						$DomainNameListboxItem = new-object System.Windows.Forms.ListViewItem("${Name}")
						[void]$DomainNameListboxItem.SubItems.Add("")
						[void]$DomainNameListboxItem.SubItems.Add("SIPINT")
						$DomainNameListboxItem.ForeColor = "Black"
						[void]$DomainNameListbox.Items.Add($DomainNameListboxItem)
					}
					elseif($FEDSRVCheckBox.Checked)
					{
						$DomainNameListboxItem = new-object System.Windows.Forms.ListViewItem("${Name}")
						[void]$DomainNameListboxItem.SubItems.Add("")
						[void]$DomainNameListboxItem.SubItems.Add("FED")
						$DomainNameListboxItem.ForeColor = "Black"
						[void]$DomainNameListbox.Items.Add($DomainNameListboxItem)
					}
					elseif($PortString -ne "" -and $PortString -match "^(6553[0-5])$|^(655[0-2]\d)$|^(65[0-4]\d{2})$|^(6[0-4]\d{3})$|^([1-5]\d{4})$|^([1-9]\d{1,3})$|^(\d{1})$")
					{
						$DomainNameListboxItem = new-object System.Windows.Forms.ListViewItem("${Name}")
						
						[void]$DomainNameListboxItem.SubItems.Add($PortString)
						$DomainNameListboxItem.ForeColor = "Black"
						[void]$DomainNameListboxItem.SubItems.Add("DIR")
						[void]$DomainNameListbox.Items.Add($DomainNameListboxItem)
						
					}
					else
					{
						Write-Host "ERROR: Incorrect port number value." -foreground "red"
					}
				}
			}
			
		}
		else
		{
			$PortString = $PortTextBox.text.Trim()
			[string]$Name = $HostNameTextBox.Text.Trim()
			
			if($PortString -match ".*,.*")
			{
				$PortSections = $PortString -split ","
				foreach($PortSection in $PortSections)
				{
					[string]$testPort = $PortSection.Trim()
					
										
					if($SIPSRVCheckBox.Checked)
					{
						$DomainNameListboxItem = new-object System.Windows.Forms.ListViewItem("${Name}")
						[void]$DomainNameListboxItem.SubItems.Add("")
						[void]$DomainNameListboxItem.SubItems.Add("SIP")
						$DomainNameListboxItem.ForeColor = "Black"
						[void]$DomainNameListbox.Items.Add($DomainNameListboxItem)
					}
					elseif($SIPINTSRVCheckBox.Checked)
					{
						$DomainNameListboxItem = new-object System.Windows.Forms.ListViewItem("${Name}")
						[void]$DomainNameListboxItem.SubItems.Add("")
						[void]$DomainNameListboxItem.SubItems.Add("SIPINT")
						$DomainNameListboxItem.ForeColor = "Black"
						[void]$DomainNameListbox.Items.Add($DomainNameListboxItem)
					}
					elseif($FEDSRVCheckBox.Checked)
					{
						$DomainNameListboxItem = new-object System.Windows.Forms.ListViewItem("${Name}")
						[void]$DomainNameListboxItem.SubItems.Add("")
						[void]$DomainNameListboxItem.SubItems.Add("FED")
						$DomainNameListboxItem.ForeColor = "Black"
						[void]$DomainNameListbox.Items.Add($DomainNameListboxItem)
					}
					elseif($testPort -ne "" -and $testPort -match "^(6553[0-5])$|^(655[0-2]\d)$|^(65[0-4]\d{2})$|^(6[0-4]\d{3})$|^([1-5]\d{4})$|^([1-9]\d{1,3})$|^(\d{1})$")
					{
						$DomainNameListboxItem = new-object System.Windows.Forms.ListViewItem("${Name}")
					
						[void]$DomainNameListboxItem.SubItems.Add($testPort)
						$DomainNameListboxItem.ForeColor = "Black"
						[void]$DomainNameListboxItem.SubItems.Add("DIR")
						[void]$DomainNameListbox.Items.Add($DomainNameListboxItem)
					}
					else
					{
						Write-Host "ERROR: Incorrect port number value." -foreground "red"
					}
				}
			}
			else
			{
							
				if($SIPSRVCheckBox.Checked)
				{
					$DomainNameListboxItem = new-object System.Windows.Forms.ListViewItem("${Name}")
					[void]$DomainNameListboxItem.SubItems.Add("")
					[void]$DomainNameListboxItem.SubItems.Add("SIP")
					$DomainNameListboxItem.ForeColor = "Black"
					[void]$DomainNameListbox.Items.Add($DomainNameListboxItem)
				}
				elseif($SIPINTSRVCheckBox.Checked)
				{
					$DomainNameListboxItem = new-object System.Windows.Forms.ListViewItem("${Name}")
					[void]$DomainNameListboxItem.SubItems.Add("")
					[void]$DomainNameListboxItem.SubItems.Add("SIPINT")
					$DomainNameListboxItem.ForeColor = "Black"
					[void]$DomainNameListbox.Items.Add($DomainNameListboxItem)
				}
				elseif($FEDSRVCheckBox.Checked)
				{
					$DomainNameListboxItem = new-object System.Windows.Forms.ListViewItem("${Name}")
					[void]$DomainNameListboxItem.SubItems.Add("")
					[void]$DomainNameListboxItem.SubItems.Add("FED")
					$DomainNameListboxItem.ForeColor = "Black"
					[void]$DomainNameListbox.Items.Add($DomainNameListboxItem)
				}
				elseif($PortString -ne "" -and $PortString -match "^(6553[0-5])$|^(655[0-2]\d)$|^(65[0-4]\d{2})$|^(6[0-4]\d{3})$|^([1-5]\d{4})$|^([1-9]\d{1,3})$|^(\d{1})$")
				{
					$DomainNameListboxItem = new-object System.Windows.Forms.ListViewItem("${Name}")
					[void]$DomainNameListboxItem.SubItems.Add($PortString)
					[void]$DomainNameListboxItem.SubItems.Add("DIR")
					$DomainNameListboxItem.ForeColor = "Black"
					[void]$DomainNameListbox.Items.Add($DomainNameListboxItem)
				}
				else
				{
					Write-Host "ERROR: Incorrect port number value." -foreground "red"
				}
			}
		}
	}
	else
	{
		Write-Host "ERROR: No FQDN/IP specified." -foreground "red"
	}
}

function Check-DNSRecord($OriginalFQDN, $Hostname, $Port)
{
	if($OriginalFQDN -ne "" -and $OriginalFQDN -ne $null)
	{
		$InformationTextBox.Text += "++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++`r`n"
		$InformationTextBox.Text += "SRV Record:  ${OriginalFQDN}`r`n"

		$Script:RTFDisplayString += "{\cf1++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++}\line "
		$Script:RTFDisplayString += "{\b {\cf2\fs21 SRV Record:  ${OriginalFQDN} }\b0 \line "

	}
	if($DNSServerTextBox.text -match "^(([01]?\d?\d|2[0-4]\d|25[0-5])\.){3}([01]?\d?\d|2[0-4]\d|25[0-5])$")
	{
		$DNSServerIP = $DNSServerTextBox.text
	}
	else
	{
		Write-Host "The DNS Server IP Address is not in a supported format. Using 8.8.8.8 instead." -foreground "red"
	}
		
	if($Hostname -ne $null -and $Hostname -ne "")
	{
		if($PortString -ne $null -and $PortString -ne "" -and $PortString -ne "0")
		{
			$StatusLabel.Text = "Checking: ${Hostname}:${PortString}"
			
			$LoopCount = 0
			Do
			{
				#LOOKUP THE SINGLE NAME
				$result = CheckCertificate -IPAddress $Hostname -Port $PortString
				$LoopCount++
				if($result -eq $false)
				{
					Write-Host "INFO: Failed to connect." -foreground "yellow"
					Write-Host "INFO: Trying again..." -foreground "yellow"
					Write-Host
					Start-Sleep -s 1
				}
			}
			Until($result -eq $true -or $LoopCount -eq 3)
			#Finished Loop
						
			
			if($result -eq $false)
			{
				
				$InformationTextBox.Text += "--------------------------------------------------------------------------------------`r`n"
				$InformationTextBox.Text += "ERROR: No certificate response from ${Hostname}:${PortString}.`r`n"
				$InformationTextBox.Text += "--------------------------------------------------------------------------------------`r`n"
				$InformationTextBox.Text += "`r`n"
				$InformationTextBox.Text += "`r`n"
				
				$Script:RTFDisplayString += "{\cf5--------------------------------------------------------------------------------------}\line "
				$Script:RTFDisplayString += "{\cf5ERROR: No certificate response from ${Hostname}:${PortString}.} \line "
				$Script:RTFDisplayString += "{\cf5--------------------------------------------------------------------------------------}\line "
				$Script:RTFDisplayString += " \line "
				$Script:RTFDisplayString += " \line "
			}
			
			if($TestPoolCheckBox.checked) #Check for multiple A record responses and check all servers
			{
				Write-Host
				Write-Host "============POOL CHECK============" -foreground "gray"
				$Script:RTFDisplayString += "{\cf2"

				#POOL TEST
				$AResultArray = GetDnsCnameA $Hostname $DNSServerIP "A"
				
				#IF MULTIPLE A RECORDS THEN TRY THEM ALL
				if($AResultArray.count -gt 1)
				{
					[string]$AllTheIPs = ""
					$theLoopNo = 1
					foreach($ARecord in $AResultArray)
					{
						$AllTheIPs += $ARecord.IPAddress
						if($theLoopNo -lt $AResultArray.count)
						{
							$AllTheIPs += ", "
						}
						$theLoopNo++
					}
					
					$InformationTextBox.Text += "--------------------------------------------------------------------------------------`r`n"
					$InformationTextBox.Text += "INFO: $Hostname DNS Resolves to multiple IPs ($AllTheIPs). Testing all servers in pool`r`n"
					$InformationTextBox.Text += "--------------------------------------------------------------------------------------`r`n"
					$InformationTextBox.Text += "`r`n"
					
					$Script:RTFDisplayString += "{\cf2--------------------------------------------------------------------------------------}\line "
					$Script:RTFDisplayString += "{\cf2INFO: $Hostname DNS Resolves to multiple IPs ($AllTheIPs).}\line "
					$Script:RTFDisplayString += "{\cf2Testing all servers in pool }\line "
					$Script:RTFDisplayString += "{\cf2--------------------------------------------------------------------------------------}\line "
					$Script:RTFDisplayString += "\line "
					
					foreach($ARecord in $AResultArray)
					{
						$ArecordIP = $ARecord.IPAddress
						Write-Host "A RECORD: $ArecordIP"
						
						#LOOKUP THE WHOLE POOL
						if($ArecordIP -match "^(([01]?\d?\d|2[0-4]\d|25[0-5])\.){3}([01]?\d?\d|2[0-4]\d|25[0-5])$")
						{
							$StatusLabel.Text = "Checking: ${ArecordIP}:${PortString}"
							
							$LoopCount = 0
							Do
							{
								#LOOKUP THE SINGLE NAME
								$result = CheckCertificate -IPAddress $ArecordIP -Port $PortString
								$LoopCount++
								if($result -eq $false)
								{
									Write-Host "INFO: Failed to connect." -foreground "yellow"
									Write-Host "INFO: Trying again..." -foreground "yellow"
									Write-Host
									Start-Sleep -s 1
								}
							}
							Until($result -eq $true -or $LoopCount -eq 3)
							#Finished Loop
							
							
							if($result -eq $false)
							{
								$InformationTextBox.Text += "--------------------------------------------------------------------------------------`r`n"
								$InformationTextBox.Text += "ERROR: No certificate response from ${ArecordIP}:${PortString}.`r`n"
								$InformationTextBox.Text += "--------------------------------------------------------------------------------------`r`n"
								$InformationTextBox.Text += "`r`n"
								$InformationTextBox.Text += "`r`n"
								
								$Script:RTFDisplayString += "{\cf5--------------------------------------------------------------------------------------}\line "
								$Script:RTFDisplayString += "{\cf5ERROR: No certificate response from ${ArecordIP}:${PortString}.} \line "
								$Script:RTFDisplayString += "{\cf5--------------------------------------------------------------------------------------}\line "
								$Script:RTFDisplayString += " \line "
								$Script:RTFDisplayString += " \line "
							}
					
							Start-Sleep -s 2
						}
						else
						{
							Write-Host "ERROR: $ArecordIP is not an IP Address" -foreground "red"
						}
					}
				}
				else
				{
					Write-Host "INFO: Only 1 server found in DNS response." -foreground "yellow"
				}
				Write-Host "==================================" -foreground "gray"
				$Script:RTFDisplayString += "}"
				$Script:RTFDisplayString += "\line"
				$InformationTextBox.Text += "`r`n"
				#$Script:RTFDisplayString += "{\cf2======================================================================================}\line "
			}
		}
		else
		{
			Write-Host "INFO: No port found..." -foreground "yellow"
			$InformationTextBox.Text += "++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++`r`n"
			$InformationTextBox.Text +=  "No Port resolved for: ${SRVName}`r`n"
			$InformationTextBox.Text += "++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++`r`n"
			$InformationTextBox.Text += "`r`n"
			$InformationTextBox.Text += "`r`n"
			
			$Script:RTFDisplayString += "{\cf5++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++}\line "
			$Script:RTFDisplayString +=  "{\cf5No Port resolved for: ${SRVName}}\line "
			$Script:RTFDisplayString += "{\cf5++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++}\line "
			$Script:RTFDisplayString += "\line "
			$Script:RTFDisplayString += "\line "
		}
	}
	else
	{
		Write-Host "INFO: No hostname found..." -foreground "yellow"
		$InformationTextBox.Text += "++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++`r`n"
		$InformationTextBox.Text +=  "No Hostname resolved for: ${SRVName}`r`n"
		$InformationTextBox.Text += "++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++`r`n"
		$InformationTextBox.Text += "`r`n"
		$InformationTextBox.Text += "`r`n"
		
		$Script:RTFDisplayString += "{\cf5++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++}\line "
		$Script:RTFDisplayString +=  "{\cf5No Hostname resolved for: ${SRVName}}\line "
		$Script:RTFDisplayString += "{\cf5++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++}\line "
		$Script:RTFDisplayString += "\line "
		$Script:RTFDisplayString += "\line "
	}
	
}

## SRV LOOKUP EXAMPLE
##
## Server:  google-public-dns-a.google.com
## Address:  8.8.8.8
##
## Non-authoritative answer:
## _sip._tls.domain.com SRV service location:
##           priority       = 0
##           weight         = 0
##           port           = 443
##           svr hostname   = sip.domain.com

#SRV Records
function GetDnsSrv($DNSQuery, $DNSServer, $DNSType)
{
	#check required variables exist
	if ($DNSType -and $DNSQuery -and $DNSServer)
	{
		$DNSLookup = Invoke-Expression "nslookup -type=`"$DNSType`" $DNSQuery $DNSServer 2>`$null"
		Write-Host "RUNNING: nslookup -type=`"$DNSType`" $DNSQuery $DNSServer" -foreground "green"
	}
	
	#Write-Host "RAW DNS ERROR: " $error[1] -foreground "red"
	#Write-Host "RAW DNS MESSAGE: $DNSLookup" -foreground "yellow"
	
	#if the query contains "SRV service location" then assume a valid response was recieved
	if ($DNSLookup -like "*SRV service location*")
	{
		$item = $null
		$DNSSRVLookupArray = @()
		$error.Clear()
		
		$DNSSRVLookup = $null
		
		Write-Host
		foreach ($item in $DNSLookup)
		{
			#Write-Host
			Write-Host $item -foreground "yellow"
			
			if($item -like "*SRV service location*")
			{
				$DNSSRVLookup = New-Object System.Object
				$DNSSRVLookup | Add-Member -type NoteProperty -name "Type" -Value "SRV"
				
				$SRVName = $item.split("SRV")
				$SRVName = $SRVName[0] -replace '\s+', ''
				$DNSSRVLookup | Add-Member -type NoteProperty -name "Name" -Value $SRVName
			}
			if ($item -like "*port*")
			{
				$SRVPort = $item.split("=")
				$SRVPort = $SRVPort[1] -replace '\s+', ''
				#Write-Host $SRVPort
				$DNSSRVLookup | Add-Member -type NoteProperty -name "Port" -Value $SRVPort
			}
			if ($item -like "*svr hostname*")
			{
				$SRVHostname = $item.split("=")
				$SRVHostname = $SRVHostname[1] -replace '\s+', ''
				#Write-Host $SRVHostname
				$DNSSRVLookup | Add-Member -type NoteProperty -name "NameHost" -Value $SRVHostname
				$DNSSRVLookupArray += $DNSSRVLookup
			}
		}
		Write-Host
		return $DNSSRVLookupArray

	}
	elseif ($error[1] -like "*Timeout*")
	{
		Write-Host "DNS Lookup Error: Server Timeout $DNSQuery" -foreground "red"
		
		#$InformationTextBox.Text += "++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++`r`n"
		#$InformationTextBox.Text +=  "DNS Lookup Error: Server Timeout $DNSQuery $DNSType`r`n"
		#$InformationTextBox.Text += "++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++`r`n"
		#$InformationTextBox.Text += "`r`n"
		#$InformationTextBox.Text += "`r`n"
		
		#$Script:RTFDisplayString += "{\cf5++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++}\line "
		#$Script:RTFDisplayString += "{\cf5DNS Lookup Error: Server Timeout $DNSQuery $DNSType} \line "
		#$Script:RTFDisplayString += "{\cf5++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++}\line "
		#$Script:RTFDisplayString += "\line "
		#$Script:RTFDisplayString += "\line "
	}
	elseif ($error[1] -like "*Timed-out*")
	{
		Write-Host "DNS Lookup Error: Server Timeout $DNSQuery $DNSType" -foreground "red"
		
		$InformationTextBox.Text += "++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++`r`n"
		$InformationTextBox.Text +=  "DNS Lookup Error: Server Timeout $DNSQuery $DNSType`r`n"
		$InformationTextBox.Text += "++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++`r`n"
		$InformationTextBox.Text += "`r`n"
		$InformationTextBox.Text += "`r`n"
		
		$Script:RTFDisplayString += "{\cf5++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++}\line "
		$Script:RTFDisplayString += "{\cf5DNS Lookup Error: Server Timeout $DNSQuery $DNSType} \line "
		$Script:RTFDisplayString += "{\cf5++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++}\line "
		$Script:RTFDisplayString += "\line "
		$Script:RTFDisplayString += "\line "
	}
	elseif ($error[1] -like "*Non-existent domain*")
	{
		Write-Host "DNS Lookup Error: Non-existent domain" -foreground "red"
		
		$InformationTextBox.Text += "++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++`r`n"
		$InformationTextBox.Text +=  "DNS Lookup Error: Non-existent domain $DNSQuery $DNSType`r`n"
		$InformationTextBox.Text += "++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++`r`n"
		$InformationTextBox.Text += "`r`n"
		$InformationTextBox.Text += "`r`n"
		
		$Script:RTFDisplayString += "{\cf5++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++}\line "
		$Script:RTFDisplayString += "{\cf5DNS Lookup Error: Non-existent domain $DNSQuery $DNSType} \line "
		$Script:RTFDisplayString += "{\cf5++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++}\line "
		$Script:RTFDisplayString += "\line "
		$Script:RTFDisplayString += "\line "
	}
	else
	{
		Write-Host "DNS Lookup Error: Unspecified Error Occured" -foreground "red"
		
		Write-Host "THE ERROR: " $error[1] -foreground "red"
		
		#$InformationTextBox.Text += "++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++`r`n"
		#$InformationTextBox.Text +=  "DNS Lookup Error: Unspecified Error Occured $DNSQuery $DNSType`r`n"
		#$InformationTextBox.Text += "++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++`r`n"
		#$InformationTextBox.Text += "`r`n"
		#$InformationTextBox.Text += "`r`n"
		
		#$Script:RTFDisplayString += "{\cf5++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++}\line "
		#$Script:RTFDisplayString += "{\cf5DNS Lookup Error: Unspecified Error Occured $DNSQuery $DNSType} \line "
		#$Script:RTFDisplayString += "{\cf5++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++}\line "
		#$Script:RTFDisplayString += "\line "
		#$Script:RTFDisplayString += "\line "
	}
}
	

## A RECORD LOOKUP EXAMPLE:
##
## Server:  google-public-dns-a.google.com
## Address:  8.8.8.8
##
## Non-authoritative answer:
## Name:    www.domain.com
## Address:  216.58.199.68
	
function GetDnsCnameA($DNSQuery, $DNSServer, $DNSType)
{
	[bool]$Addresses = $false
	[bool]$DNSAContainsAliases = $false

	if ($DNSType -ne "" -and $DNSType -ne $null -and $DNSQuery -ne "" -and $DNSQuery -ne $null -and $DNSServer -ne "" -and $DNSServer -ne $null)
	{
		$DNSLookup = Invoke-Expression "nslookup -type=`"A`" $DNSQuery $DNSServer 2>`$null"
		Write-Host "RUNNING: nslookup -type=`"A`" $DNSQuery $DNSServer" -foreground "green"
		
		#Check if query result is a CNAME record.
		if ($DNSLookup -like "*Aliases:*")
		{
		  $DNSAContainsAliases = $true
		}
	}
	
	#Write-Host "RAW DNS ERROR: " $error[1] -foreground "red"
	#Write-Host "RAW DNS MESSAGE: $DNSLookup" -foreground "yellow"
	
	#if the query contains "Name:" then assume a valid response was recieved
	if ($DNSLookup -like "*Name:*")
	{
		$item = $null
		$DNSALookupArray = @()
		$error.Clear()
		[int]$ANameCount = $null
		[int]$AAddressCount = $null
		[int]$AAliasesCount = $null
		$AName = $null
		$AAddress = $null
		$AAliases = $null
		
		foreach ($item in $DNSLookup)
		{
			Write-Host $item -foreground "yellow"
			
			# If all required variables are populated, reset Address to null
			if ($DNSAContainsAliases)
			{
				if (($ANameCount -eq $AAddressCount) -and ($AAliasesCount -ge "1"))
				{
					$AAddress = $null
				}
			}
			else
			{
				if ($ANameCount -eq $AAddressCount)
				{
					$AAddress = $null
				}
			}
			if ($item -like "*Name:*")
			{
				$ANameCount++
				$AName = $item.split(":")
				$AName = $AName[1] -replace '\s+', ''
			}
			if ($item -like "*Aliases:*")
			{					
				$AAliasesCount++
				$AAliases = $item.split(":")
				$AAliases = $AAliases[1] -replace '\s+', ''                    
			}
			if ($item -like "*Address:*")
			{
				if ($item -notlike "*$DNSServer*")
				{
					$AAddressCount++
					$AAddress = $item.split(":")
					$AAddress = $AAddress[1] -replace '\s+', ''
				}
			}
			if ($item -like "*Addresses:*" -or ($Addresses -eq $true -and $item -match '\b((25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)(\.|$)){4}\b'))
			{
				if ($Addresses -eq $true)
				{
					$AAddress = $item -replace '\s+', ''
				}
				else
				{
					$AAddressCount++
					$AAddress = $item.split(":")
					$AAddress = $AAddress[1] -replace '\s+', ''
				}
				$Addresses = $true
			}
			if ($ANameCount -eq $AAddressCount)
			{

				if ($DNSAContainsAliases)
				{
					if ($AName -and $AAddress -and $AAliases)
					{
						$DNSALookup = New-Object System.Object
						$DNSALookup | Add-Member -type NoteProperty -name "Type" -Value "CNAME"
						$DNSALookup | Add-Member -type NoteProperty -name "Name" -Value $AAliases
						$DNSALookup | Add-Member -type NoteProperty -name "NameHost" -Value $AName
						$DNSALookup | Add-Member -type NoteProperty -name "IPAddress" -Value $AAddress
						$DNSALookupArray += $DNSALookup
					}
				}
				else
				{
					if ($AName -and $AAddress)
					{
						$DNSALookup = New-Object System.Object
						$DNSALookup | Add-Member -type NoteProperty -name "Type" -Value "A"
						$DNSALookup | Add-Member -type NoteProperty -name "Name" -Value $AName
						$DNSALookup | Add-Member -type NoteProperty -name "NameHost" -Value ""
						$DNSALookup | Add-Member -type NoteProperty -name "IPAddress" -Value $AAddress
						
						$DNSALookupArray += $DNSALookup
					}
				}
			}
		}
		foreach($item in $DNSALookupArray)
		{
			Write-Host "A RECORD FOUND: "$item.Name $item.IPAddress -foreground "Green"
		}
		Write-Host
		
		return $DNSALookupArray
	}
	elseif ($error[1] -like "*Timeout*")
	{
		Write-Host "DNS Lookup Error: Server Timeout" -foreground "red"
		
		#$InformationTextBox.Text += "++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++`r`n"
		#$InformationTextBox.Text +=  "DNS Lookup Error: Server Timeout $DNSQuery $DNSType`r`n"
		#$InformationTextBox.Text += "++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++`r`n"
		#$InformationTextBox.Text += "`r`n"
		#$InformationTextBox.Text += "`r`n"
		
		#$Script:RTFDisplayString += "{\cf5++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++}\line "
		#$Script:RTFDisplayString += "{\cf5DNS Lookup Error: Server Timeout $DNSQuery $DNSType} \line "
		#$Script:RTFDisplayString += "{\cf5++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++}\line "
		#$Script:RTFDisplayString += "\line "
		#$Script:RTFDisplayString += "\line "
	}
	elseif ($error[1] -like "*Timed-out*")
	{
		Write-Host "DNS Lookup Error: Server Timeout $DNSQuery $DNSType" -foreground "red"
		
		$InformationTextBox.Text += "++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++`r`n"
		$InformationTextBox.Text +=  "DNS Lookup Error: Server Timeout $DNSQuery $DNSType`r`n"
		$InformationTextBox.Text += "++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++`r`n"
		$InformationTextBox.Text += "`r`n"
		$InformationTextBox.Text += "`r`n"
		
		$Script:RTFDisplayString += "{\cf5++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++}\line "
		$Script:RTFDisplayString += "{\cf5DNS Lookup Error: Server Timeout $DNSQuery $DNSType} \line "
		$Script:RTFDisplayString += "{\cf5++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++}\line "
		$Script:RTFDisplayString += "\line "
		$Script:RTFDisplayString += "\line "
	}
	elseif ($error[1] -like "*Non-existent domain*")
	{
		Write-Host "DNS Lookup Error: Non-existent domain" -foreground "red"
		
		#$InformationTextBox.Text += "++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++`r`n"
		#$InformationTextBox.Text +=  "DNS Lookup Error: Non-existent domain $DNSQuery $DNSType`r`n"
		#$InformationTextBox.Text += "++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++`r`n"
		#$InformationTextBox.Text += "`r`n"
		#$InformationTextBox.Text += "`r`n"
		
		#$Script:RTFDisplayString += "{\cf5++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++}\line "
		#$Script:RTFDisplayString += "{\cf5DNS Lookup Error: Non-existent domain $DNSQuery $DNSType} \line "
		#$Script:RTFDisplayString += "{\cf5++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++}\line "
		#$Script:RTFDisplayString += "\line "
		#$Script:RTFDisplayString += "\line "
	}
	elseif ($error[1] -like "*No response from server*")
	{
		Write-Host "DNS Lookup Error: No response from server" -foreground "red"
		
		#$InformationTextBox.Text += "++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++`r`n"
		#$InformationTextBox.Text +=  "DNS Lookup Error: No response from server $DNSQuery $DNSType`r`n"
		#$InformationTextBox.Text += "++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++`r`n"
		#$InformationTextBox.Text += "`r`n"
		#$InformationTextBox.Text += "`r`n"
		
		#$Script:RTFDisplayString += "{\cf5++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++\}line "
		#$Script:RTFDisplayString += "{\cf5DNS Lookup Error: No response from server $DNSQuery $DNSType} \line "
		#$Script:RTFDisplayString += "{\cf5++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++}\line "
		#$Script:RTFDisplayString += "\line "
		#$Script:RTFDisplayString += "\line "
	}
	else
	{
		Write-Host "DNS Lookup Error: Unspecified Error Occured" -foreground "red"
		
		Write-Host "THE ERROR: " $error[1] -foreground "red"
		
		#$InformationTextBox.Text += "++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++`r`n"
		#$InformationTextBox.Text +=  "DNS Lookup Error: Unspecified Error Occured $DNSQuery $DNSType`r`n"
		#$InformationTextBox.Text += "++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++`r`n"
		#$InformationTextBox.Text += "`r`n"
		#$InformationTextBox.Text += "`r`n"
		
		#$Script:RTFDisplayString += "{\cf5++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++}\line "
		#$Script:RTFDisplayString += "{\cf5DNS Lookup Error: Unspecified Error Occured $DNSQuery $DNSType} \line "
		#$Script:RTFDisplayString += "{\cf5++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++}\line "
		#$Script:RTFDisplayString += "\line "
		#$Script:RTFDisplayString += "\line "
	}
	
}
	
	
function CheckCertificate([string]$IPAddress, [int]$Port )
{
	#$startDTM = $null #Setup variable...
	
	Write-Host "Attempting to connect on $IPAddress port $Port" -foreground "yellow"
	$timeout = 4000
	
	#Can force the use of specific local port... There may be a use for this at some point?
	#$IPEndPoint = New-Object System.Net.IPEndPoint($LocalInterfaceIP, $LocalSourcePort)
	#$tcpclient = New-Object System.Net.Sockets.TcpClient($IPEndPoint)
	
	$tcpclient = New-Object System.Net.Sockets.TcpClient
	
	$iar = $tcpclient.BeginConnect($IPAddress,$Port,$null,$null)
	$wait = $iar.AsyncWaitHandle.WaitOne($timeout,$false)
	
	trap 
	{
		Write-Host "ERROR: TCP General Exception" -foreground "red"
		Write-Host "Exception: $($_.exception.message)" -foreground "red"
		#$InformationTextBox.Text += "++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++`r`n"
		#$InformationTextBox.Text +=  "General Connection Exception connecting to $IPAddress port $Port`r`n"
		#$InformationTextBox.Text +=  "Exception: $($_.exception.message)`r`n"
		#$InformationTextBox.Text += "++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++`r`n"
		#$InformationTextBox.Text += "`r`n"
		
		#$Script:RTFDisplayString += "{\cf5++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++}\line "
		#$Script:RTFDisplayString +=  "{\cf5General Connection Exception connecting to $IPAddress port $Port}\line "
		#$Script:RTFDisplayString +=  "{\cf5Exception: $($_.exception.message)}\line "
		#$Script:RTFDisplayString += "{\cf5++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++}\line "
		#$Script:RTFDisplayString += "\line "
		return $false
	}
	trap [System.Net.Sockets.SocketException]
	{
		Write-Host "ERROR: TCP Exception: $($_.exception.message)" -foreground "red"
		return $false
	}
	if(!$wait)
	{
		$tcpClient.Client.Disconnect($true)
		$tcpclient.Close()
		Write-Host "ERROR: TCP Connection Timeout..." -foreground "red"
		#$InformationTextBox.Text += "++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++`r`n"
		#$InformationTextBox.Text +=  "Failed to connect to: ${IPAddress}:${Port}`r`n"
		#$InformationTextBox.Text +=  "Connection Timeout`r`n"
		#$InformationTextBox.Text += "++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++`r`n"
		#$InformationTextBox.Text += "`r`n"
		#$InformationTextBox.Text += "`r`n"
		
		
		#$Script:RTFDisplayString += "++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++\line "
		#$Script:RTFDisplayString += "{\cf5Failed to connect to: ${IPAddress}:${Port}}\line "
		#$Script:RTFDisplayString += "{\cf5Connection Timeout}\line "
		#$Script:RTFDisplayString += "++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++\line "
		#$Script:RTFDisplayString += "\line "
		#$Script:RTFDisplayString += "\line "
					
		return $false
	}
	else
	{
		
		$Stream = $tcpclient.GetStream()
		$RemoteIP = $tcpclient.Client.RemoteEndPoint
		$socket = $tcpclient.Client
		Write-Host
		Write-Host "Local Address:" ($socket.LocalEndPoint).Address.ToString()
		Write-Host "Local Port: "  ($socket.LocalEndPoint).Port.ToString()
		Write-Host "Remote IP: $RemoteIP"
		
		if($Stream -ne $null)
		{
			try {
					
				$sbCallback = {
					
					#Connection has worked now disconnect
					Write-Host "End TCP connection..."
					Write-Host
					$sslStream.Close()
 					$tcpclient.Close()
					
					if($args.length -ge 3)
					{
						$certificate = [System.Security.Cryptography.X509Certificates.X509Certificate2]$args[1]
						#Debug
						#Write-Host "args 0: " $args[0]
						#Write-Host
						#Write-Host "==================Certificate Response===================" -foreground "Yellow"
						#Write-Host "args 1: " $args[1] -foreground "Yellow"
						#Write-Host "args 2: " $args[2] -foreground "Yellow"
						#Write-Host "=========================================================" -foreground "Yellow"
						#Write-Host
						
						$chain = [System.Security.Cryptography.X509Certificates.X509Chain]$args[2]
					
					
						if ($certificate -ne $null)
						{
				
							[string]$Subject = $certificate.Subject.ToString()
							$SubjectName = $certificate.SubjectName.Name
							[string]$issuer = $certificate.Issuer
							$IssuerName = $certificate.IssuerName.Name
							[string]$NotBefore = $certificate.NotBefore.ToString()
							[string]$NotAfter = $certificate.NotAfter.ToString()
							[string]$SerialNumber = $certificate.SerialNumber.ToString()
							[string]$SignatureAlgorithm = $certificate.SignatureAlgorithm.FriendlyName
							[string]$Thumbprint = $certificate.Thumbprint.ToString()
							[string]$Version = $certificate.Version.ToString()
							[string]$HasPrivateKey = $certificate.HasPrivateKey.ToString()
							[string]$Handle = $certificate.Handle.ToString()
							[string] $isValid = $certificate.Verify()
							$Extensions = $certificate.Extensions
							[string]$Archived = $certificate.Archived.ToString()

							
							[System.DateTime] $NotBeforeDate = $certificate.NotBefore
							[System.DateTime] $NotAfterDate = $certificate.NotAfter
							[System.DateTime] $CurrentDate = Get-Date
							
							
							Write-Host "--------------------------------------------------------------------------------------" -foreground "Yellow"
							Write-Host "Checking: ${IPAddress}:${Port}" -foreground "Yellow"
							Write-Host "IP Address: $RemoteIP" -foreground "Yellow"
							Write-Host
							Write-Host "Certificate Response:" -foreground "Yellow"
							Write-Host							
							Write-Host "Subject: $Subject" -foreground "Yellow"
							
							
							$CNErrorString = ""
							$CNNameArray = $Subject -Split ","
							if($CNNameArray[0] -match "CN=")
							{
								[string]$CNName = $CNNameArray[0]
								$CNName = $CNName -Replace "CN=", ""
								if($CNName -eq $IPAddress)
								{
									Write-Host "Common Name Match found!" -foreground "green"
									$CNErrorString = "- Common Name Match found"
								}
								else
								{
									if($CNName -match "^\*\.")
									{
										Write-Host "Warning: Wildcard Certificate is being used. Not supported for Lync / Skype for Business servers." -foreground "red"
										$CNErrorString = "- Warning: Wildcard Certificate is being used. Not supported for Lync / Skype for Business servers."
									}
									elseif(!($IPAddress -match "^((25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)(\.|$)){4}$"))
									{
										Write-Host "Warning: Common Name does not match. Check the SAN list for a match." -foreground "red"
										$CNErrorString = "- Common Name does not match."
									}
									else
									{
										Write-Host "INFO: Testing a server IP Address so CN check is skipped" -foreground "green"
										$CNErrorString = "- Info: Testing a server IP Address so CN check is skipped."
									}
								}
							}
							
							
							Write-Host "Issuer: $issuer" -foreground "Yellow"
							
							$NotBeforeErrorString = ""
							Write-Host "Not Before: $NotBefore" -foreground "Yellow"
							if($CurrentDate -gt $NotBeforeDate)
							{
								Write-Host "Certificate before date is OK!" -foreground "green"
							}
							else
							{
								Write-Host "Warning: Server date is before certificate creation date! Check server and certificate dates." -foreground "red"
								$NotBeforeErrorString = "- Warning: Server date is before certificate creation date! Check server and certificate dates."
							}
							
							$NotAfterErrorString = ""
							Write-Host "Not After: $NotAfter" -foreground "Yellow"
							if($NotAfterDate -gt $CurrentDate)
							{
								Write-Host "Certificate expiry date is OK!" -foreground "green"
							}
							else
							{
								Write-Host "Warning: Certificate expiry date is expired!" -foreground "red"
								$NotAfterErrorString = "- Warning: Certificate expiry date is expired!"
							}
							
							Write-Host "Serial Number: $SerialNumber" -foreground "Yellow"
							Write-Host "Signature Algorithm: $SignatureAlgorithm" -foreground "Yellow"
							
							$SignatureErrorString = ""
							#Signing Algorithm SHA-1 and SHA-2 suite of digest sizes (224, 256, 384 and 512-bit)
							if($SignatureAlgorithm -match "SHA")
							{
								Write-Host "Signing Algorithm is SHA. OK!" -foreground "green"
							}
							elseif($SignatureAlgorithm -match "RSASSA-PSS")
							{
								#Reference: https://technet.microsoft.com/en-us/library/gg398066(v=ocs.15).aspx
								Write-Host "Signing Algorithm RSASSA-PSS is unsupported by Lync / Skype for Business." -foreground "red"
								$SignatureErrorString = "- Warning: Signing Algorithm RSASSA-PSS is unsupported by Lync / Skype for Business."
							}
							else
							{
								Write-Host "Signing Algorithm is not supported for Lync / Skype for Business!" -foreground "red"
								$SignatureErrorString = "- Warning: Signing Algorithm is not supported for Lync / Skype for Business!"
							}
								
							Write-Host "Thumbprint: $Thumbprint" -foreground "Yellow"
							Write-Host "Version: $Version" -foreground "Yellow"
							Write-Host "HasPrivateKey: $HasPrivateKey" -foreground "Yellow"
							Write-Host "Archived: $Archived" -foreground "Yellow"
							$SANErrorString = ""
							$ServerEKUErrorString = ""
							$CRLErrorString = ""
							foreach($Extension in $Extensions)
							{
								[System.Security.Cryptography.AsnEncodedData] $asndata = New-Object System.Security.Cryptography.AsnEncodedData($extension.Oid, $extension.RawData)
								Write-Host "Extension Type: " $extension.Oid.FriendlyName  -foreground "Yellow"
								Write-Host "Oid Value: " $asndata.Oid.Value  -foreground "Yellow"
								Write-Host "Data: " $asndata.Format($true) -foreground "Yellow"
								
								[bool]$SANMatch = $false
								#CHECK SAN NAMES
																
								if($asndata.Oid.Value -eq "2.5.29.17")
								{
									if(!($IPAddress -match "^((25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)(\.|$)){4}$"))
									{
										[string]$data = $asndata.Format($true)
										$dataArray = $data.Split("`n")
										foreach($line in $dataArray)
										{
											#Write-Host "SAN CHECK: $line"
											[string]$line = $line -Replace "DNS Name=",""
											$line = $line.Trim()
											if($line -eq $IPAddress)
											{
												$SANMatch = $true
												Write-Host "SAN Match found!" -foreground "green"
												$SANErrorString = "- FQDN is in SAN list."
												break
											}
										}
										if($SANMatch -eq $false)
										{
											Write-Host "SAN was not found!" -foreground "red"
											$SANErrorString = "- Warning: SAN was not found!"
										}
									}
									else
									{
										Write-Host "INFO: Testing a server IP Address so SAN check is skipped" -foreground "green"
										$SANErrorString = "- Info: Testing a server IP Address so SAN check is skipped"
									}
								}

								#CHECK THE CRL LOCATION EXISTS
								if($asndata.Oid.Value -eq "2.5.29.31")
								{
									Write-Host "A CRL Location included in cert!" -foreground "green"
									$CRLFound = $true
								}
								
								
								#SERVER EKU CHECK - Server EKU 1.3.6.1.5.5.7.3.1
								if($asndata.Oid.Value -eq "2.5.29.37")
								{
									$data = $asndata.Format($true)
									
									if($data -match "1.3.6.1.5.5.7.3.1")
									{
										Write-Host "Server EKU found!" -foreground "green"
										$ServerEKUFound = $true
									}
									else
									{
										#Write-Host "No server EKU found!" -foreground "red"
									}
								}
								
								#CHECK KEY Length - 1024, 2048, 4096
								#CHECK KEY HASHES - ECDH_P256, ECDH_P384, ECDH_P512, RSA

							}
							if($ServerEKUFound -eq $false)
							{
								$ServerEKUErrorString = "- Warning: No server EKU found in the certificate. Skype for Business / Lync servers require a server EKU."
							}
							if($CRLFound -eq $false)
							{
								$CRLErrorString = "- No CRL found in certificate. Skype for Business / Lync servers require an accessible CRL."
							}
							Write-Host "--------------------------------------------------------------------------------------" -foreground "Yellow"
							Write-Host
							Write-Host
		
							
							$InformationTextBox.Text += "++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++`r`n"
							#$InformationTextBox.Text +=  "`r`n"
							$InformationTextBox.Text +=  "Checking FQDN: ${IPAddress}:${Port}`r`n"
							$InformationTextBox.Text +=  "Checking IP Address: $RemoteIP`r`n"
							$InformationTextBox.Text +=  "`r`n"
							$InformationTextBox.Text +=  "Certificate Response:`r`n"
							$InformationTextBox.Text +=  "`r`n"
							$InformationTextBox.Text +=  "Subject: $Subject`r`n"
							$InformationTextBox.Text +=  "Issuer: $issuer`r`n"
							$InformationTextBox.Text +=  "Not Before: $NotBefore`r`n"
							$InformationTextBox.Text +=  "Not After: $NotAfter`r`n"
							$InformationTextBox.Text +=  "Serial Number: $SerialNumber`r`n"
							$InformationTextBox.Text +=  "Signature Algorithm: $SignatureAlgorithm`r`n"
							$InformationTextBox.Text +=  "Thumbprint: $Thumbprint`r`n"
							$InformationTextBox.Text +=  "Version: $Version`r`n"
							$InformationTextBox.Text +=  "Has Private Key: $HasPrivateKey`r`n"
							$InformationTextBox.Text +=  "Is Valid: $isValid`r`n"
							
							
							$Script:RTFDisplayString += "++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++\line "
							#$Script:RTFDisplayString += "\line "
							$Script:RTFDisplayString += "\b {\cf2\fs21 Checking FQDN:  ${IPAddress}:${Port} }\b0 \line "
							$Script:RTFDisplayString += "\b Checking IP Address: \b0 $RemoteIP\line "
							$Script:RTFDisplayString += "\line "
							$Script:RTFDisplayString += "\b{\cf2Certificate Response:} \b0 \line "
							$Script:RTFDisplayString += "\line "
							$Script:RTFDisplayString +=  "\b Subject: \b0 $Subject\line "
							$Script:RTFDisplayString += "\b Issuer: \b0 $issuer\line "
							$Script:RTFDisplayString +=  "\b Not Before: \b0 $NotBefore\line "
							$Script:RTFDisplayString +=  "\b Not After: \b0 $NotAfter\line "
							$Script:RTFDisplayString +=  "\b Serial Number: \b0 $SerialNumber\line "
							$Script:RTFDisplayString += "\b Signature Algorithm: \b0 $SignatureAlgorithm\line "
							$Script:RTFDisplayString += "\b Thumbprint: \b0 $Thumbprint\line "
							$Script:RTFDisplayString += "\b Version: \b0 $Version\line "
							$Script:RTFDisplayString +=  "\b Has Private Key: \b0 $HasPrivateKey\line "
							$Script:RTFDisplayString +=  "\b Is Valid: \b0 $isValid\line "
							
							
							if($ShowAdvancedCheckBox.checked)
							{
								$InformationTextBox.Text +=  "Archived: $Archived`r`n`r`n"
								$Script:RTFDisplayString +=  "\b Archived: \b0 $Archived\line \line "
								foreach($Extension in $Extensions)
								{
									[System.Security.Cryptography.AsnEncodedData] $asndata = New-Object System.Security.Cryptography.AsnEncodedData($extension.Oid, $extension.RawData)
									$InformationTextBox.Text +=  "Extension Type: " + $extension.Oid.FriendlyName + "`r`n"
									$InformationTextBox.Text +=  "Oid Value: " + $asndata.Oid.Value + "`r`n"
									$InformationTextBox.Text +=  "Data:`r`n" 
									$InformationTextBox.Text +=  $asndata.Format($true) + "`r`n"
									
									$Script:RTFDisplayString +=  "\b Extension Type:\b0  " + $extension.Oid.FriendlyName + "\line "
									$Script:RTFDisplayString +=  "\b Oid Value:\b0  " + $asndata.Oid.Value + "\line "
									$Script:RTFDisplayString +=  "\b Data:\b0 \line " 
									[string]$OIDDATA = $asndata.Format($true)
									$OIDDATA = $OIDDATA.Replace("`n", "\line ")
									$Script:RTFDisplayString +=  $OIDDATA
									$Script:RTFDisplayString += "\line "
									
								}
							}
														
							if($ShowChainCheckBox.checked)
							{
								if($chain -ne $null)
								{
									$LoopNo = 1
									$HighestCertInChain = "Unknown"
									$InformationTextBox.Text +=  "`r`n"
									$InformationTextBox.Text +=  "Certificate Chain:`r`n"
									$InformationTextBox.Text +=  "`r`n"
									
									$Script:RTFDisplayString +=  "\line "
									$Script:RTFDisplayString +=  "\b{\cf2Certificate Chain:}\b0 \line \line "
									
									Write-Host "CERTIFICATE CHAIN"
									Write-Host
									
									foreach ($element in $chain.ChainElements)
									{
										
										$InformationTextBox.Text +=  "Certificate Chain Item $LoopNo`r`n"
										$Script:RTFDisplayString += "\b Certificate Chain Item $LoopNo\b0 \line "
										#X509ChainElement  $element
										[string]$chainSubjectName = $element.Certificate.SubjectName.Name
										[string]$chainIssuer = $element.Certificate.Issuer
										[string]$chainBefore = $element.Certificate.NotBefore
										[string]$chainUntil = $element.Certificate.NotAfter
										[string]$chainValid = $element.Certificate.Verify()
										[string]$chainLength = $element.Certificate.ChainElementStatus.Length
										[string]$chainExtCount = $element.Certificate.Extensions.Count
										[string]$chainSignatureAlgorithm = $element.Certificate.SignatureAlgorithm.FriendlyName
										[string]$chainSerialNumber = $element.Certificate.SerialNumber.ToString()
										[string]$chainThumbprint = $element.Certificate.Thumbprint.ToString()
										[string]$chainVersion = $element.Certificate.Version.ToString()
										
										
										Write-Host "Chain Subject Name: $chainSubjectName"
										Write-Host "Chain Issuer name: $chainIssuer"
										Write-Host "Chain Certificate Not Before: $chainUntil"
										Write-Host "Chain Certificate valid until: $chainUntil"
										Write-Host "Chain error status length: $chainLength"
										Write-Host "Chain Serial Number: $chainSerialNumber" 
										Write-Host "Chain Thumbprint: $chainThumbprint" 
										Write-Host "Chain Version: $chainVersion" 
										Write-Host "Chain Signature Algorithm: $chainSignatureAlgorithm"
										Write-Host "Chain Certificate is valid: $chainValid"
										Write-Host "Number of element extensions: $chainExtCount"
										Write-Host 
										
										$InformationTextBox.Text +=  "Chain Subject Name: $chainSubjectName `r`n"
										$InformationTextBox.Text +=  "Chain Issuer name: $chainIssuer`r`n"
										$InformationTextBox.Text +=  "Chain Not Before: $chainBefore`r`n"
										$InformationTextBox.Text +=  "Chain Not After: $chainUntil`r`n"
										$InformationTextBox.Text +=  "Chain Serial Number: $chainSerialNumber`r`n"
										$InformationTextBox.Text +=  "Chain Signature Algorithm: $chainSignatureAlgorithm`r`n"
										$InformationTextBox.Text +=  "Chain Thumbprint: $chainThumbprint`r`n"
										$InformationTextBox.Text +=  "Chain Version: $chainVersion`r`n"
										$InformationTextBox.Text +=  "Chain is valid: $chainValid`r`n"
										$InformationTextBox.Text +=  "`r`n"
										
										$Script:RTFDisplayString +=  "\b Chain Subject Name:\b0  ${chainSubjectName}\line "
										$Script:RTFDisplayString +=  "\b Chain Issuer name:\b0  ${chainIssuer}\line "
										$Script:RTFDisplayString +=  "\b Chain Not Before:\b0  ${chainBefore}\line "
										$Script:RTFDisplayString +=  "\b Chain Not After:\b0  ${chainUntil}\line "
										$Script:RTFDisplayString +=  "\b Chain Serial Number:\b0  ${chainSerialNumber}\line "
										$Script:RTFDisplayString +=  "\b Chain Signature Algorithm:\b0  ${chainSignatureAlgorithm}\line "
										$Script:RTFDisplayString +=  "\b Chain Thumbprint:\b0  ${chainThumbprint}\line "
										$Script:RTFDisplayString +=  "\b Chain Version:\b0  ${chainVersion}\line "
										$Script:RTFDisplayString +=  "\b Chain is valid:\b0  ${chainValid}\line "
										$Script:RTFDisplayString +=  "\line "
																				
										$HighestCertInChain = $chainIssuer
										$LoopNo++
									}
									Write-Host
									
									$InformationTextBox.Text +=  "--------------------------------------------------------------------------------------`r`n"
									#$InformationTextBox.Text +=  "`r`n"
									$InformationTextBox.Text +=  "Root Certificates:`r`n"
									$InformationTextBox.Text +=  "`r`n"
									
									$Script:RTFDisplayString +=  "--------------------------------------------------------------------------------------\line "
									#$Script:RTFDisplayString +=  "\line "
									$Script:RTFDisplayString +=  "\b{\cf2Root Certificates:} \b0\line "
									$Script:RTFDisplayString +=  "\line "
									
									
									
									if($HighestCertInChain -match "Digicert")
									{
										$issuerSplit = $HighestCertInChain -Split "CN="
										Write-Host "$issuer"
										Write-Host "Issuer Split " $issuerSplit.count 		
										$theIsserName = "Unknown"
										
										if($issuerSplit -is [Array])
										{
											Write-Host "`$issuerSplit[1]" $issuerSplit[1]
											$theIsserName = $issuerSplit[1] -Split ","
											[string]$theIsserNameString = $theIsserName[0].Trim()
											$theIsserNameString = $theIsserNameString.Replace(" ", "")
										}
										elseif($issuerSplit -is [String])
										{
											$theIsserName = $issuerSplit -Split ","
											$theIsserNameString = $theIsserName[0]
											[string]$theIsserNameString = $theIsserNameString.Replace(" ", "")
										}
										
										Write-Host "Get Root Certs here: https://www.digicert.com/digicert-root-certificates.htm" -foreground "green"
										Write-Host
										$InformationTextBox.Text += "Get Root Certs here: https://www.digicert.com/digicert-root-certificates.htm `r`n"
										$InformationTextBox.Text += "Download Root Cert: http://cacerts.digicert.com/${theIsserNameString}.crt `r`n"
									
										$Script:RTFDisplayString += "- Get Root Certs here: https://www.digicert.com/digicert-root-certificates.htm \line "
										$Script:RTFDisplayString += "- Download Root Cert: http://cacerts.digicert.com/${theIsserNameString}.crt \line "
									}
									elseif($HighestCertInChain -match "Entrust")
									{
										Write-Host "Get Root Certs here: http://www.entrust.com/get-support/ssl-certificate-support/root-certificate-downloads/" -foreground "green"
										Write-Host
										$InformationTextBox.Text += "Get Root Certs here: http://www.entrust.com/get-support/ssl-certificate-support/root-certificate-downloads/ `r`n"
									
										$Script:RTFDisplayString += "- Get Root Certs here: http://www.entrust.com/get-support/ssl-certificate-support/root-certificate-downloads/ \line "
										
										if($HighestCertInChain -imatch "Entrust Root Certification Authority - G2")
										{
											$InformationTextBox.Text += "Download Root Cert: http://www.entrust.com/root-certificates/entrust_g2_ca.cer `r`n"
											$Script:RTFDisplayString += "- Download Root Cert: http://www.entrust.com/root-certificates/entrust_g2_ca.cer \line "
										}
										elseif($HighestCertInChain -imatch "Entrust Root Certification Authority - G3")
										{
											$InformationTextBox.Text += "Download Root Cert: http://www.entrust.com/root-certificates/entrust_g3_ca.cer `r`n"
											$Script:RTFDisplayString += "- Download Root Cert: http://www.entrust.com/root-certificates/entrust_g2_ca.cer \line "
										}
										elseif($HighestCertInChain -imatch "Entrust Root Certification Authority - EC1")
										{
											$InformationTextBox.Text += "Download Root Cert: http://www.entrust.com/root-certificates/entrust_ec1_ca.cer `r`n"
											$Script:RTFDisplayString += "- Download Root Cert: http://www.entrust.com/root-certificates/entrust_g2_ca.cer \line "
										}
										elseif($HighestCertInChain -imatch "Entrust.net Certification Authority \(2048\)")
										{
											$InformationTextBox.Text += "Download Root Cert: http://www.entrust.com/root-certificates/entrust_2048_ca.cer `r`n"
											$Script:RTFDisplayString += "- Download Root Cert: http://www.entrust.com/root-certificates/entrust_g2_ca.cer \line "
										}
										elseif($HighestCertInChain -imatch "Entrust Root Certification Authority")
										{
											$InformationTextBox.Text += "Download Root Cert: http://www.entrust.com/root-certificates/entrust_ev_ca.cer `r`n"
											$Script:RTFDisplayString += "- Download Root Cert: http://www.entrust.com/root-certificates/entrust_g2_ca.cer \line "
										}
										<#
										elseif($HighestCertInChain -imatch "Entrust.net Certification Authority (2048)")
										{
											$InformationTextBox.Text += "Download Root Cert: https://www.entrust.net/downloads/binary/entrust_2048_ca.cer `r`n"
										}
										elseif($HighestCertInChain -imatch "Entrust.net Certification Authority")
										{
											$InformationTextBox.Text += "Download Root Cert: https://www.entrust.net/downloads/binary/entrust_ssl_ca.cer `r`n"
										}
										 #>
										
																	
									}
									elseif($HighestCertInChain -match "thawte")
									{
										Write-Host "Get Root Certs here: https://www.thawte.com/roots/" -foreground "green"
										Write-Host
										$InformationTextBox.Text += "Get Root Certs here: https://www.thawte.com/roots/ `r`n"
										
										$Script:RTFDisplayString += "- Get Root Certs here: https://www.thawte.com/roots/ \line "
										
										if($HighestCertInChain -imatch "Thawte Personal Freemail CA")
										{
											$InformationTextBox.Text += "Download Root Cert: https://www.thawte.com/roots/thawte_Personal_Freemail_CA.pem `r`n"
											$Script:RTFDisplayString += "- Download Root Cert: https://www.thawte.com/roots/thawte_Personal_Freemail_CA.pem \line "
										}
										elseif($HighestCertInChain -imatch "thawte Primary Root CA - G3")
										{
											$InformationTextBox.Text += "Download Root Cert: https://www.thawte.com/roots/thawte_Primary_Root_CA-G3_SHA256.pem `r`n"
											$Script:RTFDisplayString += "- Download Root Cert: https://www.thawte.com/roots/thawte_Primary_Root_CA-G3_SHA256.pem \line "
										}
										elseif($HighestCertInChain -imatch "thawte Primary Root CA - G2")
										{
											$InformationTextBox.Text += "Download Root Cert: https://www.thawte.com/roots/thawte_Primary_Root_CA-G2_ECC.pem `r`n"
											$Script:RTFDisplayString += "- Download Root Cert: https://www.thawte.com/roots/thawte_Primary_Root_CA-G2_ECC.pem \line "
										}
										elseif($HighestCertInChain -imatch "Thawte Server CA")
										{
											$InformationTextBox.Text += "Download Root Cert: https://www.thawte.com/roots/thawte_Server_CA.pem `r`n"
											$Script:RTFDisplayString += "- Download Root Cert: https://www.thawte.com/roots/thawte_Server_CA.pem \line "
										}
										elseif($HighestCertInChain -imatch "Thawte Premium Server CA")
										{
											$InformationTextBox.Text += "Download Root Cert: https://www.thawte.com/roots/thawte_Premium_Server_CA.pem `r`n"
											$Script:RTFDisplayString += "- Download Root Cert: https://www.thawte.com/roots/thawte_Premium_Server_CA.pem \line "
										}
										elseif($HighestCertInChain -imatch "thawte Primary Root CA")
										{
											$InformationTextBox.Text += "Download Root Cert: https://www.thawte.com/roots/thawte_Primary_Root_CA.pem `r`n"
											$Script:RTFDisplayString += "- Download Root Cert: https://www.thawte.com/roots/thawte_Primary_Root_CA.pem \line "
										}
										
									}
									elseif($HighestCertInChain -match "symantec")
									{
										
										Write-Host "Get Root Certs here: http://www.symantec.com/page.jsp?id=roots" -foreground "green"
										Write-Host
										$InformationTextBox.Text += "Get Root Certs here: http://www.symantec.com/page.jsp?id=roots"
										$Script:RTFDisplayString += "- Get Root Certs here: http://www.symantec.com/page.jsp?id=roots \line "
										
										$issuerSplit = $HighestCertInChain -Split "CN="
										Write-Host "$issuer"
										Write-Host "Issuer Split " $issuerSplit.count 		
										$theIsserName = "Unknown"
										
										if($issuerSplit -is [Array])
										{
											Write-Host "`$issuerSplit[1]" $issuerSplit[1]
											$theIsserName = $issuerSplit[1] -Split ","
											[string]$theIsserNameString = $theIsserName[0]
											$theIsserNameString = $theIsserNameString.Replace(" - ", " ").Replace(" 3","%203").Replace(" ", "-")
										}
										elseif($issuerSplit -is [String])
										{
											$theIsserName = $issuerSplit -Split ","
											$theIsserNameString = $theIsserName[0]
											$theIsserNameString = $theIsserNameString.Replace(" - ", " ").Replace(" 3","%203").Replace(" ", "-")
										}
										$InformationTextBox.Text += "Download Root Cert: http://www.symantec.com/content/en/us/enterprise/verisign/roots/${theIsserNameString}.pem `r`n"
										$Script:RTFDisplayString += "- Download Root Cert: http://www.symantec.com/content/en/us/enterprise/verisign/roots/${theIsserNameString}.pem \line "
									}
									elseif($HighestCertInChain -match "verisign")
									{
										Write-Host "Get Root Certs here: http://www.symantec.com/page.jsp?id=roots" -foreground "green"
										Write-Host
										$InformationTextBox.Text += "Get Root Certs here: http://www.symantec.com/page.jsp?id=roots `r`n"
										$Script:RTFDisplayString += "- Get Root Certs here: http://www.symantec.com/page.jsp?id=roots \line "
										
										$issuerSplit = $HighestCertInChain -Split "CN="
										Write-Host "$issuer"
										Write-Host "Issuer Split " $issuerSplit.count 		
										$theIsserName = "Unknown"
										
										if($issuerSplit -is [Array])
										{
											Write-Host "`$issuerSplit[1]" $issuerSplit[1]
											$theIsserName = $issuerSplit[1] -Split ","
											[string]$theIsserNameString = $theIsserName[0]
											$theIsserNameString = $theIsserNameString.Replace(" - ", " ").Replace(" 3","%203").Replace(" ", "-")
										}
										elseif($issuerSplit -is [String])
										{
											$theIsserName = $issuerSplit -Split ","
											$theIsserNameString = $theIsserName[0]
											$theIsserNameString = $theIsserNameString.Replace(" - ", " ").Replace(" 3","%203").Replace(" ", "-")
										}
										$InformationTextBox.Text += "Download Root Cert: http://www.symantec.com/content/en/us/enterprise/verisign/roots/${theIsserNameString}.pem `r`n"
										$Script:RTFDisplayString += "- Download Root Cert: http://www.symantec.com/content/en/us/enterprise/verisign/roots/${theIsserNameString}.pem \line "
										#"http://www.symantec.com/content/en/us/enterprise/verisign/roots/VeriSign-Class%203-Public-Primary-Certification-Authority-G5.pem"
									}
									elseif($HighestCertInChain -match "globalsign")
									{
										Write-Host "Get Root Certs here: https://support.globalsign.com/customer/portal/articles/1426602-globalsign-root-certificates" -foreground "green"
										Write-Host
										$InformationTextBox.Text += "Get Root Certs here: https://support.globalsign.com/customer/portal/articles/1426602-globalsign-root-certificates"
										$Script:RTFDisplayString += "- Get Root Certs here: https://support.globalsign.com/customer/portal/articles/1426602-globalsign-root-certificates \line "
									
										if($HighestCertInChain -imatch "GlobalSign Root R1")
										{
											$InformationTextBox.Text += "Download Root Cert: https://secure.globalsign.net/cacert/Root-R1.crt? `r`n"
											$Script:RTFDisplayString += "- Download Root Cert: https://secure.globalsign.net/cacert/Root-R1.crt? \line "
										}
										elseif($HighestCertInChain -imatch "GlobalSign Root R2")
										{
											$InformationTextBox.Text += "Download Root Cert: https://secure.globalsign.net/cacert/Root-R2.crt? `r`n"
											$Script:RTFDisplayString += "- Download Root Cert: https://secure.globalsign.net/cacert/Root-R2.crt? \line "
										}
										elseif($HighestCertInChain -imatch "GlobalSign Root R3")
										{
											$InformationTextBox.Text += "Download Root Cert: https://secure.globalsign.net/cacert/Root-R3.crt? `r`n"
											$Script:RTFDisplayString += "- Download Root Cert: https://secure.globalsign.net/cacert/Root-R3.crt? \line "
										}
										elseif($HighestCertInChain -imatch "GlobalSign ECC Root R4")
										{
											$InformationTextBox.Text += "Download Root Cert: https://secure.globalsign.net/cacert/Root-R4.crt? `r`n"
											$Script:RTFDisplayString += "- Download Root Cert: https://secure.globalsign.net/cacert/Root-R4.crt? \line "
										}
										elseif($HighestCertInChain -imatch "GlobalSign ECC Root R5")
										{
											$InformationTextBox.Text += "Download Root Cert: https://secure.globalsign.net/cacert/Root-R5.crt? `r`n"
											$Script:RTFDisplayString += "- Download Root Cert: https://secure.globalsign.net/cacert/Root-R5.crt? \line "
										}
										
									
									}
									elseif($HighestCertInChain -match "go daddy")
									{
										Write-Host "Get Root Certs here: https://certs.godaddy.com/repository" -foreground "green"
										Write-Host
										$InformationTextBox.Text += "Get Root Certs here: https://certs.godaddy.com/repository `r`n"
										$Script:RTFDisplayString += "- Get Root Certs here: https://certs.godaddy.com/repository \line "
										
										if($HighestCertInChain -imatch "Go Daddy Class 2 Certification Authority - G2")
										{
											$InformationTextBox.Text += "Download Root Cert: https://certs.godaddy.com/repository/gdroot-g2.crt `r`n"
											$Script:RTFDisplayString += "- Download Root Cert: https://certs.godaddy.com/repository/gdroot-g2.crt \line "
										}
										elseif($HighestCertInChain -imatch "Go Daddy Class 2 Certification Authority")
										{
											$InformationTextBox.Text += "Download Root Cert: https://certs.godaddy.com/repository/gd-class2-root.crt `r`n"
											$Script:RTFDisplayString += "- Download Root Cert: https://certs.godaddy.com/repository/gd-class2-root.crt \line "
										}
										
									}
									elseif($HighestCertInChain -match "geotrust" -or $HighestCertInChain -match "Equifax")
									{
										Write-Host "Get Root Certs here: https://www.geotrust.com/resources/root-certificates/" -foreground "green"
										Write-Host
										$InformationTextBox.Text += "Get Root Certs here: https://www.geotrust.com/resources/root-certificates/ `r`n"
										$Script:RTFDisplayString += "- Get Root Certs here: https://www.geotrust.com/resources/root-certificates/ \line "
										
										$issuerSplit = $HighestCertInChain -Split "CN="
										Write-Host "$issuer"
										Write-Host "Issuer Split " $issuerSplit.count 		
										$theIsserName = "Unknown"
										
										if($issuerSplit -is [Array])
										{
											Write-Host "`$issuerSplit[1]" $issuerSplit[1]
											$theIsserName = $issuerSplit[1] -Split ","
											[string]$theIsserNameString = $theIsserName[0]
											$theIsserNameString = $theIsserNameString.Replace(" ", "_")
										}
										elseif($issuerSplit -is [String])
										{
											$theIsserName = $issuerSplit -Split ","
											$theIsserNameString = $theIsserName[0]
											$theIsserNameString = $theIsserNameString.Replace(" ", "_")
										}
										$InformationTextBox.Text += "Download Root Cert: https://www.geotrust.com/resources/root_certificates/certificates/${theIsserNameString}.pem `r`n"
										$Script:RTFDisplayString += "- Download Root Cert: https://www.geotrust.com/resources/root_certificates/certificates/${theIsserNameString}.pem \line "
										
										#GeoTrust Global CA
										#"https://www.geotrust.com/resources/root_certificates/certificates/GeoTrust_Global_CA.pem"
									}
									elseif($HighestCertInChain -match "comodo" -or $HighestCertInChain -match "AddTrust" -or $HighestCertInChain -match "Network Solutions")
									{
										Write-Host "Get Root Certs here: https://support.comodo.com/index.php?/Default/Knowledgebase/List/Index/71" -foreground "green"
										Write-Host
										$InformationTextBox.Text += "Get Root Certs here: https://support.comodo.com/index.php?/Default/Knowledgebase/List/Index/71 `r`n"
										$Script:RTFDisplayString += "- Get Root Certs here: https://support.comodo.com/index.php?/Default/Knowledgebase/List/Index/71 \line "
										
										#Add Trust
										if($HighestCertInChain -imatch "AddTrust External TTP Network")
										{
											$InformationTextBox.Text += "Download Root Cert: https://support.comodo.com/index.php?/Knowledgebase/Article/GetAttachment/917/66 `r`n"
											$Script:RTFDisplayString += "- Download Root Cert: https://support.comodo.com/index.php?/Knowledgebase/Article/GetAttachment/917/66 \line "
										}
										elseif($HighestCertInChain -imatch "AddTrust External CA Root")
										{
											$InformationTextBox.Text += "Download Root Cert: ftp://ftp.networksolutions.com/certs/AddTrustExternalCARoot.crt `r`n"
											$Script:RTFDisplayString += "- Download Root Cert: ftp://ftp.networksolutions.com/certs/AddTrustExternalCARoot.crt \line "
										}
										elseif($HighestCertInChain -imatch "UTN-USERFirst-Hardware")
										{
											$InformationTextBox.Text += "Download Root Cert: ftp://ftp.networksolutions.com/certs/UTNAddTrustServer_CA.crt `r`n"
											$Script:RTFDisplayString += "- Download Root Cert: ftp://ftp.networksolutions.com/certs/UTNAddTrustServer_CA.crt \line "
										}
										elseif($HighestCertInChain -imatch "COMODO RSA Certification Authority")
										{
											$InformationTextBox.Text += "Download Root Cert: https://support.comodo.com/index.php?/Knowledgebase/Article/GetAttachment/969/821026 `r`n"
											$Script:RTFDisplayString += "- Download Root Cert: https://support.comodo.com/index.php?/Knowledgebase/Article/GetAttachment/969/821026 \line "
										}
										elseif($HighestCertInChain -imatch "")
										{
											$InformationTextBox.Text += "Download Root Cert:  `r`n"
											$Script:RTFDisplayString += "- Download Root Cert:  \line "
										}
												
										
									}
									elseif($HighestCertInChain -match "CyberTrust")
									{
										Write-Host "Get Root Certs here: http://cybertrust.omniroot.com/support/sureserver/rootcert_iis.cfm" -foreground "green"
										Write-Host
										$InformationTextBox.Text += "Get Root Certs here: http://cybertrust.omniroot.com/support/sureserver/rootcert_iis.cfm `r`n"
										$Script:RTFDisplayString += "- Get Root Certs here: http://cybertrust.omniroot.com/support/sureserver/rootcert_iis.cfm \line "
										
										if($HighestCertInChain -imatch "GTE CyberTrust Global Root")
										{
											$InformationTextBox.Text += "Download Root Cert: http://secure.omniroot.com/cacert/ct_root.der `r`n"
											$Script:RTFDisplayString += "- Download Root Cert: http://secure.omniroot.com/cacert/ct_root.der \line "
										}
										elseif($HighestCertInChain -imatch "Baltimore CyberTrust Root")
										{
											$InformationTextBox.Text += "Download Root Cert: http://cacert.omniroot.com/bc2025.crt `r`n"
											$Script:RTFDisplayString += "- Download Root Cert: http://cacert.omniroot.com/bc2025.crt \line "
										}
										
									}
									elseif($HighestCertInChain -match "StarField")
									{
										Write-Host "Get Root Certs here: https://certs.secureserver.net/repository" -foreground "green"
										Write-Host
										$InformationTextBox.Text += "Get Root Certs here: https://certs.secureserver.net/repository `r`n"
										$Script:RTFDisplayString += "- Get Root Certs here: https://certs.secureserver.net/repository \line "
										
										if($HighestCertInChain -imatch "Starfield Class 2 Certification Authority")
										{
											$InformationTextBox.Text += "Download Root Cert: https://certs.secureserver.net/repository/sf-class2-root.crt `r`n"
											$Script:RTFDisplayString += "- Download Root Cert: https://certs.secureserver.net/repository/sf-class2-root.crt \line "
										}
										elseif($HighestCertInChain -imatch "Starfield Root Certificate Authority - G2")
										{
											$InformationTextBox.Text += "Download Root Cert: https://certs.secureserver.net/repository/sfroot-g2.crt `r`n"
											$Script:RTFDisplayString += "- Download Root Cert: https://certs.secureserver.net/repository/sfroot-g2.crt \line "
										}
										
									}
									else
									{
										Write-Host "NO MATCH FOR COMMON CERT AUTHORITY"
										Write-Host
										if($HighestCertInChain -match "CN=" -and $HighestCertInChain -match ",")
										{
											$issuerSplit = $HighestCertInChain -Split "CN="
											Write-Host "$issuer"
											Write-Host "Issuer Split " $issuerSplit.count 		
											$theIsserName = "Unknown"
											
											if($issuerSplit -is [Array])
											{
												Write-Host "`$issuerSplit[1]" $issuerSplit[1]
												$theIsserName = $issuerSplit[1] -Split ","
												$theIsserNameString = $theIsserName[0].Trim()
											}
											elseif($issuerSplit -is [String])
											{
												$theIsserName = $issuerSplit -Split ","
												$theIsserNameString = $theIsserName[0]
											}
											Write-Host "Certificate is issued by: $theIsserNameString"
											$InformationTextBox.Text += "Certificate is issued by: $theIsserNameString `r`n"
											$InformationTextBox.Text += "Contact this provider for a copy of the root certificate. `r`n"
											
											$Script:RTFDisplayString += "- Certificate is issued by: $theIsserNameString \line "
											$Script:RTFDisplayString += "- Contact this provider for a copy of the root certificate. \line "
											
										}
										else
										{
											$InformationTextBox.Text += "Certificate is issued by: $HighestCertInChain `r`n"
											$InformationTextBox.Text += "Contact this provider for a copy of the root certificate. `r`n"
											
											$Script:RTFDisplayString += "- Certificate is issued by: $HighestCertInChain \line "
											$Script:RTFDisplayString += "- Contact this provider for a copy of the root certificate. \line "
										}
									}
									
								}
								else
								{
									Write-Host "ERROR: Certificate chain equals null" -foreground "red"
									return $false
								}
							}
							
							
							$Script:RTFDisplayString +=  "\line "
							$Script:RTFDisplayString +=  "--------------------------------------------------------------------------------------\line "
							#$Script:RTFDisplayString +=  "\line "
							$Script:RTFDisplayString +=  "\b {\cf2 Comments:} \b0\line "
							$Script:RTFDisplayString +=  "\line "
							
							$InformationTextBox.Text +=  "`r`n"
							$InformationTextBox.Text +=  "--------------------------------------------------------------------------------------`r`n"
							#f$InformationTextBox.Text +=  "`r`n"
							$InformationTextBox.Text +=  "Comments:`r`n"
							$InformationTextBox.Text +=  "`r`n"
							
							if($CNErrorString -ne "")
							{
								$Script:RTFDisplayString +=  "$CNErrorString \line "
								$InformationTextBox.Text += "$CNErrorString`r`n"
								#$Script:RTFDisplayString +=  "\line "
							}
							if($SANErrorString -ne "")
							{
								$Script:RTFDisplayString +=  "$SANErrorString \line "
								$InformationTextBox.Text += "$SANErrorString`r`n"
								#$Script:RTFDisplayString +=  "\line "
							}
							if($SignatureErrorString -ne "")
							{
								$Script:RTFDisplayString +=  "$SignatureErrorString \line "
								$InformationTextBox.Text += "$SignatureErrorString`r`n"
								#$Script:RTFDisplayString +=  "\line "
							}
							if($NotAfterErrorString -ne "")
							{
								$Script:RTFDisplayString +=  "$NotAfterErrorString \line "
								$InformationTextBox.Text += "$NotAfterErrorString`r`n"
								#$Script:RTFDisplayString +=  "\line "
							}
							if($NotBeforeErrorString -ne "")
							{
								$Script:RTFDisplayString +=  "$NotBeforeErrorString \line "
								$InformationTextBox.Text += "$NotBeforeErrorString`r`n"
								#$Script:RTFDisplayString +=  "\line "
							}
							if($ServerEKUErrorString -ne "")
							{
								$Script:RTFDisplayString +=  "$ServerEKUErrorString \line "
								$InformationTextBox.Text += "$ServerEKUErrorString`r`n"
								#$Script:RTFDisplayString +=  "\line "
							}
							if($CRLErrorString -ne "")
							{
								$Script:RTFDisplayString +=  "$CRLErrorString \line "
								$InformationTextBox.Text += "$CRLErrorString`r`n"
							}
							
							
							$InformationTextBox.Text +=  "`r`n"
							$InformationTextBox.Text += "++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++`r`n"
							$InformationTextBox.Text +=  "`r`n"
							$InformationTextBox.Text +=  "`r`n"
							$InformationTextBox.Text +=  "`r`n"
							
							$Script:RTFDisplayString +=  "\line "
							$Script:RTFDisplayString += "++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++\line "
							$Script:RTFDisplayString +=  "\line "
							$Script:RTFDisplayString +=  "\line "
							$Script:RTFDisplayString +=  "\line "
							
							
							$InformationTextBox.SelectionStart = $InformationTextBox.Text.length
							$InformationTextBox.ScrollToCaret()
							[System.Windows.Forms.Application]::DoEvents()
							
							
							if($OutFileCheckBox.Checked -eq $true)
							{
								$filename = $FileLocationTextBox.Text
								##Export to CSV
								
								if($FirstCheck -eq $true)
								{
									Write-Host "${IPAddress}:${Port} First Check was equal to TRUE" -foreground "red"
									
									$csv = "`"FQDN / IP`",`"IP Address`",`"Subject`",`"Issuer`",`"Not Before`",`"Not After`",`"Serial Number`",`"Signature Algorithm`",`"Thumbprint`",`"Version`",`"HasPrivateKey`",`"Archived`""
									
									$script:FirstCheck = $false
									
									foreach($Extension in $Extensions)
									{
										$csv += ",`"Extension Type`",`"Oid Value`",`"Data`""
									}
									
									#$csv += "`r`n"
									
									#Create new UTF-8 file...
									$csv | out-file -Encoding UTF8 -FilePath $filename -Force
								}
								
								$output = "`"${IPAddress}:${Port}`",`"$RemoteIP`",`"$Subject`",`"$issuer`",`"$NotBefore`",`"$NotAfter`",`"$SerialNumber`",`"$SignatureAlgorithm`",`"$Thumbprint`",`"$Version`",`"$HasPrivateKey`",`"$Archived`"" 		
													
								foreach($Extension in $Extensions)
								{
									[System.Security.Cryptography.AsnEncodedData] $asndata = New-Object System.Security.Cryptography.AsnEncodedData($extension.Oid, $extension.RawData)
									$ExtensionName = $extension.Oid.FriendlyName
									$OID = $asndata.Oid.Value
									$Data = $asndata.Format($false)
									
									$output += ",`"$ExtensionName`",`"$OID`",`"$Data`""
								}
								
								#Append to UTF-8 file...					
								$output | out-file -Encoding UTF8 -FilePath $filename -Force -Append					
								Write-Host "Written to CSV File...." -foreground "yellow"
							}
												
							return $true
							
						}
						else
						{
							Write-Host "ERROR: Certificate equals null" -foreground "red"
							return $false
						}
					}
					else
					{
						Write-Host "ERROR: Incorrect number of arguments in certificate check callback" -foreground "red"
					}
				}
				
				Write-Host				
				Write-Host "Creating new SSLSTREAM..."
				$sslStream = New-Object System.Net.Security.SslStream($Stream,$False,([Net.ServicePointManager]::ServerCertificateValidationCallback = $sbCallback))
				
				$sslStream.ReadTimeout = 5000
				$sslStream.WriteTimeout = 5000
				Write-Host "Authenticate as client... $IPAddress"
				Write-Host	
				
				[byte[]]$LocalCert = @(48, 130, 21, 100, 2, 1, 3, 48, 130, 21, 32, 6, 9, 42, 134, 72, 134, 247, 13, 1, 7, 1, 160, 130, 21, 17, 4, 130, 21, 13, 48, 130, 21, 9, 48, 130, 6, 58, 6, 9, 42, 134, 72, 134, 247, 13, 1, 7, 1, 160, 130, 6, 43, 4, 130, 6, 39, 48, 130, 6, 35, 48, 130, 6, 31, 6, 11, 42, 134, 72, 134, 247, 13, 1, 12, 10, 1, 2, 160, 130, 4, 254, 48, 130, 4, 250, 48, 28, 6, 10, 42, 134, 72, 134, 247, 13, 1, 12, 1, 3, 48, 14, 4, 8, 76, 103, 228, 53, 113, 185, 88, 105, 2, 2, 7, 208, 4, 130, 4, 216, 88, 147, 17, 34, 124, 90, 94, 204, 9, 157, 194, 17, 51, 62, 119, 190, 180, 58, 90, 212, 157, 204, 47, 120, 205, 190, 177, 51, 131, 48, 114, 220, 72, 182, 213, 171, 108, 184, 187, 183, 155, 66, 186, 240, 198, 246, 3, 165, 81, 115, 157, 165, 152, 188, 248, 74, 66, 185, 59, 192, 125, 181, 164, 43, 117, 95, 37, 10, 141, 20, 57, 76, 155, 32, 117, 118, 203, 135, 95, 222, 164, 127, 251, 247, 237, 156, 26, 154, 232, 169, 132, 0, 214, 103, 128, 237, 229, 196, 65, 138, 169, 194, 4, 105, 112, 242, 52, 175, 215, 249, 19, 60, 176, 226, 239, 118, 239, 167, 199, 127, 214, 157, 153, 15, 216, 28, 25, 0, 17, 165, 48, 220, 27, 28, 127, 106, 153, 41, 39, 238, 101, 133, 14, 44, 133, 17, 251, 52, 93, 84, 255, 220, 157, 22, 8, 116, 9, 3, 74, 238, 121, 216, 65, 83, 70, 144, 42, 36, 160, 130, 107, 223, 22, 253, 169, 184, 65, 144, 184, 162, 27, 79, 114, 181, 208, 154, 9, 146, 160, 62, 24, 250, 42, 54, 32, 44, 29, 178, 26, 86, 230, 73, 119, 24, 64, 242, 160, 71, 196, 56, 181, 99, 133, 142, 222, 62, 150, 10, 144, 51, 105, 198, 235, 195, 28, 228, 223, 151, 6, 145, 229, 79, 33, 62, 14, 253, 81, 230, 161, 211, 142, 84, 112, 28, 52, 163, 35, 30, 37, 120, 135, 75, 166, 124, 177, 87, 219, 176, 187, 103, 187, 136, 244, 193, 240, 115, 20, 228, 127, 118, 138, 55, 140, 227, 128, 150, 112, 216, 162, 79, 190, 185, 202, 17, 208, 140, 240, 192, 249, 84, 8, 32, 39, 219, 171, 216, 186, 242, 223, 207, 181, 127, 118, 179, 34, 182, 109, 129, 208, 111, 129, 173, 225, 46, 7, 88, 77, 183, 20, 85, 157, 121, 172, 176, 11, 158, 32, 41, 186, 65, 181, 97, 2, 216, 236, 8, 97, 31, 248, 90, 218, 227, 48, 17, 56, 220, 53, 140, 148, 44, 97, 57, 173, 244, 210, 231, 2, 161, 10, 21, 5, 26, 78, 47, 123, 179, 175, 233, 116, 86, 148, 63, 49, 169, 25, 174, 20, 92, 130, 23, 102, 236, 118, 115, 17, 163, 86, 242, 206, 188, 79, 138, 60, 189, 56, 47, 106, 174, 80, 151, 24, 120, 84, 72, 62, 21, 211, 88, 162, 110, 151, 103, 177, 93, 235, 169, 37, 173, 74, 198, 147, 69, 229, 147, 241, 59, 21, 213, 40, 22, 56, 5, 6, 132, 146, 38, 134, 164, 210, 76, 63, 77, 94, 111, 71, 125, 218, 216, 211, 140, 144, 69, 182, 178, 84, 28, 126, 249, 37, 91, 46, 48, 18, 28, 179, 103, 73, 117, 10, 37, 74, 212, 217, 25, 27, 98, 165, 200, 212, 103, 232, 45, 251, 71, 105, 191, 54, 79, 14, 222, 198, 79, 188, 70, 236, 83, 170, 176, 125, 206, 73, 81, 38, 196, 250, 47, 49, 21, 116, 151, 249, 119, 66, 175, 90, 195, 134, 173, 251, 246, 107, 6, 254, 178, 248, 251, 85, 4, 39, 72, 170, 55, 245, 211, 148, 210, 251, 39, 150, 54, 61, 74, 66, 45, 185, 245, 27, 184, 32, 54, 31, 234, 140,242, 116, 18, 34, 114, 127, 47, 26, 45, 102, 234, 152, 209, 130, 129, 241, 61, 148, 6, 191, 6, 101, 15, 189, 87, 149, 203, 43, 13, 107, 58, 158, 248, 183, 140, 173, 167, 196, 109, 79, 215, 131, 126, 108, 103, 100, 116, 180, 130, 217, 6, 243, 246, 11, 71, 24, 206, 175, 35, 131, 40, 160, 127, 44, 127, 159, 210, 74, 84, 86, 115, 25, 249, 38, 92, 133, 60, 184, 0, 235, 20, 188, 26, 200, 54, 233, 197, 228, 175, 84, 194, 65, 229, 20, 188, 231, 153, 225, 247, 235, 110, 141, 44, 253, 95, 232, 88, 87, 121, 10, 112, 255, 37, 218, 118, 87, 11, 170, 213, 65, 52, 226, 6, 146, 66, 166, 230, 150, 251, 207, 175, 108, 38, 81, 109, 24, 90, 223, 202, 120, 20, 125, 26, 59, 77, 168, 0, 16, 139, 10, 70, 107, 106, 225, 133, 1, 197, 32, 222, 84, 175, 155, 87, 140, 114, 146, 113, 103, 154, 62, 235, 72, 208, 127, 65, 251, 227, 161, 82, 148, 183, 5, 44, 47, 241, 164, 153, 158, 199, 155, 86, 187, 15, 108, 5, 204, 113, 201, 3, 135, 187, 35, 7, 34, 227, 243, 171, 5, 156, 237, 80, 56, 124, 118, 27, 85, 77, 212, 72, 125, 221, 60, 56, 64, 218, 111, 56, 106, 29, 236, 69, 231, 207, 165, 39, 130, 143, 233, 80, 85, 105, 6, 207, 165, 95, 4, 0, 57, 241, 96, 179, 87, 13, 56, 139, 176, 57, 237, 86, 146, 143, 216, 234, 67, 75, 209, 169, 6, 124, 23, 123, 59, 10, 74, 220, 15, 108, 253, 57, 255, 81, 142, 217, 145, 222, 40, 127, 8, 154, 67, 246, 127, 228, 204, 37, 31, 239, 43, 110, 27, 17, 28, 220, 68, 103, 52, 234, 97, 248, 70, 70, 255, 127, 16, 107, 159, 211, 165, 10, 15, 237, 223, 66, 67, 199, 60, 129, 201, 46, 49, 3, 65, 194, 130, 244, 77, 207, 251, 42, 224, 135, 210, 237, 4, 131, 129, 219, 189, 50, 22, 145, 141, 58, 114, 47, 213, 79, 104, 88, 145, 73, 232, 241, 137, 47, 170, 18, 114, 154, 253, 124, 169, 165, 41, 191, 90, 169, 129, 177, 172, 112, 103, 78, 65, 199, 152, 251, 146, 115, 29, 154, 224, 209, 172, 191, 82, 199, 55, 159, 75, 133, 71, 178, 52, 48, 116, 16, 221, 33, 138, 170, 149, 58, 212, 171, 209, 205, 189, 136, 182, 162, 72, 88, 71, 242, 160, 79, 40, 36, 118, 24, 255, 174, 203, 204, 85, 177, 149, 34, 8, 230, 126, 17, 102, 188, 228, 247, 125, 144, 39, 222, 158, 55, 205, 173, 5, 79, 82, 152, 237, 110, 189, 134, 145, 95, 199, 28, 34, 57, 76, 235, 16, 70, 197, 183, 221, 99, 189, 22, 155, 207, 234, 136, 182, 99, 76, 15, 61, 14, 242, 190, 127, 132, 181, 52, 251, 108, 43, 76, 228, 164, 173, 25, 21, 87, 173, 36, 127, 241, 183, 229, 144, 224, 185, 216, 55, 104, 62, 224, 252, 59, 14, 131, 175, 96, 13, 5, 15, 119, 137, 38, 213, 133, 203, 251, 74, 85, 243, 176, 215, 149, 17, 44, 129, 81, 197, 21, 218, 162, 143, 175, 210, 134, 88, 162, 138, 87, 122, 7, 155, 37, 87, 213, 137, 164, 245, 135, 37, 197, 106, 31, 22, 35, 151, 88, 126, 160, 156, 224, 154, 45, 127, 230, 49, 38, 52, 35, 74, 80, 166, 193, 115, 247, 215, 9, 247, 142, 9, 10, 117, 28, 99, 78, 251, 13, 247, 214, 58, 230, 21, 208, 74, 122, 221, 152, 164, 185, 61, 227, 2, 200, 189, 4, 194, 224, 127, 78, 79, 209, 235, 143, 197, 248, 175, 38, 51, 206, 246, 207, 173, 141, 3, 177, 90, 218, 243, 205, 230, 121, 223, 106, 241, 8, 228, 90, 247, 58, 181, 244, 42, 211, 143, 70, 75, 136, 115, 37, 40, 38, 95, 43, 48, 66, 55, 60, 8, 3, 132, 103, 107, 201, 2, 135, 170, 166, 235, 193, 49, 130, 1, 12, 48, 13, 6, 9, 43, 6, 1, 4, 1, 130, 55, 17, 2, 49, 0, 48, 19, 6, 9, 42, 134, 72, 134, 247, 13, 1, 9, 21, 49, 6, 4, 4, 1, 0, 0, 0, 48, 105, 6, 9, 43, 6, 1, 4, 1, 130, 55, 17, 1, 49, 92, 30, 90, 0, 77, 0, 105, 0, 99, 0, 114, 0, 111, 0, 115, 0, 111, 0, 102, 0, 116, 0, 32, 0, 82, 0, 83, 0, 65, 0, 32, 0, 83, 0, 67, 0, 104, 0, 97, 0, 110, 0, 110, 0, 101, 0, 108, 0, 32, 0, 67, 0, 114, 0, 121, 0, 112, 0, 116, 0, 111, 0, 103, 0, 114, 0, 97, 0, 112, 0, 104, 0, 105, 0, 99, 0, 32, 0, 80, 0, 114, 0, 111, 0, 118, 0, 105, 0, 100, 0, 101, 0, 114, 48, 123, 6, 9, 42, 134, 72, 134, 247, 13, 1, 9, 20, 49, 110, 30, 108, 0, 67, 0, 101, 0, 114, 0, 116, 0, 82, 0, 101, 0, 113, 0, 45, 0, 87, 0, 101, 0, 98, 0, 83, 0, 101, 0, 114, 0, 118, 0, 101, 0, 114, 0, 45, 0, 56, 0, 57, 0, 49, 0, 55, 0, 50, 0, 101, 0, 50, 0, 57, 0, 45, 0, 102, 0, 100, 0, 48, 0, 48, 0, 45, 0, 52, 0, 49, 0, 99, 0, 97, 0, 45, 0, 57, 0, 50, 0, 99, 0, 56, 0, 45, 0, 48, 0, 52, 0, 54, 0, 51, 0, 52, 0, 55, 0, 52, 0, 51, 0, 101, 0, 99, 0, 48, 0, 50, 48, 130, 14, 199, 6, 9, 42, 134, 72, 134, 247, 13, 1, 7, 6, 160, 130, 14, 184, 48, 130, 14, 180, 2, 1, 0, 48, 130, 14, 173, 6, 9, 42, 134, 72, 134, 247, 13, 1, 7, 1, 48, 28, 6, 10, 42, 134, 72, 134, 247, 13, 1, 12, 1, 6, 48, 14, 4, 8, 15, 241, 136, 109, 83, 184, 87, 104, 2, 2, 7, 208, 128, 130, 14, 128, 205, 28, 60, 73, 58, 12, 245, 145, 117, 220, 255, 53, 241, 121, 4, 78, 97, 235, 122, 24, 72, 182, 89, 69, 75, 215, 135, 57, 167, 234, 237, 106, 116, 132, 91, 162, 189, 231, 122, 222, 245, 159, 37, 35, 160, 23, 123, 85, 61, 5, 136, 203, 119, 31, 200, 143, 15, 78, 46, 66, 40, 125, 64, 88, 76, 56, 12, 156, 232, 7, 45, 188, 76, 52, 111, 74, 125, 100, 140, 11, 76, 204, 85, 203, 111, 106, 100, 245, 159, 247, 73, 18, 137, 61, 68, 78, 250, 144, 45, 140, 148, 26, 9, 39, 183, 90, 31, 129, 74, 73, 202, 155, 109, 70, 182, 117, 16, 67, 81, 250, 121, 177, 10, 24, 175, 27, 225, 8, 237, 32, 185, 44, 122, 10, 223, 120, 98, 154, 57, 219, 87, 151, 228, 105, 75, 126, 161, 145, 87, 201, 228, 25, 180, 213, 104, 75, 125, 214, 98, 73, 232, 93, 71, 33, 180, 229, 93, 99, 81, 89, 127, 170, 53, 138, 104, 56, 34, 141, 74, 163, 216, 203, 137, 66, 202, 191, 232, 43, 223, 63, 234, 23, 18, 80, 112, 42, 73, 167, 23, 159, 22, 111, 181, 169, 160, 78, 107, 213, 39, 83, 107, 18, 144, 213, 95, 183, 98, 177, 58, 120, 210, 229, 79, 63, 126, 76, 96, 97, 108, 206, 37, 113, 173, 205, 171, 145, 192, 193, 96, 153, 144, 241, 168, 117, 107, 89, 95, 90, 79, 28, 135, 235, 193, 62, 10, 6, 99, 14, 1, 195, 25, 177, 252, 219, 0, 124, 189, 225, 144, 205, 68, 9, 46, 92, 174, 63, 54, 229, 116, 107, 98, 76, 220, 144, 148, 208, 199, 206, 17, 107, 49, 179, 203, 30, 102, 169, 33, 152, 179, 229, 116, 162, 69, 219, 64, 201, 10, 161, 245, 124, 208, 104, 113, 163, 244, 109, 118, 92, 189, 0, 46, 99, 195, 205, 29, 164, 67, 119, 110, 99, 207, 85, 39, 52, 231, 180, 75, 220, 248, 137, 90, 165, 187, 25, 196, 110, 68, 73, 100, 78, 117, 202, 78, 105, 169, 181, 91, 63, 204, 111, 201, 196, 38, 74, 45, 182, 252, 226, 100, 113, 87, 176, 157, 165, 218, 229, 50, 116, 164, 138, 179, 1, 93, 76, 227, 5, 7, 196, 207, 212, 16, 242, 122, 30, 182, 114, 58, 77, 207, 41, 189, 104, 238, 210, 199, 91, 145, 224, 34, 155, 171, 33, 92, 52, 216, 106, 235, 198, 202, 143, 105, 101, 216, 9, 62, 228, 24, 220, 112, 242, 28, 49, 134, 230, 66, 21, 199, 128, 69, 165, 156, 186, 139, 93, 16, 97, 125, 186, 241, 2, 31, 18, 24, 1, 218, 122, 124, 13, 200, 31, 194, 6, 233, 11, 121, 94, 251, 28, 186, 202, 41, 168, 53, 180, 4, 229, 243, 216, 171, 39, 128, 26, 15, 231, 157, 131, 100, 228, 151, 17, 24, 196, 10, 100, 112, 4, 254, 78, 70, 191, 124, 11, 252, 163, 61, 120, 247, 215, 44, 197, 45, 222, 90, 18, 247, 117, 230, 160, 95, 166, 178, 2, 137, 190, 205, 30, 71, 65, 38, 8, 161, 153, 137, 179, 94, 127, 204, 236, 191, 145, 147, 182, 124, 114, 29, 71, 51, 195, 176, 138, 179, 40, 226, 40, 133, 247, 154, 189, 7, 19, 150, 45, 243, 189, 216, 238, 220, 209, 156, 218, 92, 18, 136, 85, 94, 229, 186, 75, 227, 92, 119, 145, 228, 26, 203, 31, 143, 194, 219, 134, 153, 215, 211, 100, 21, 10, 146, 109, 110, 235, 74, 173, 88, 222, 160, 202, 253, 208, 241, 94, 255, 133, 76, 183, 116, 22, 191, 117, 124, 229, 0, 132, 207, 56, 216, 205, 234, 41, 61, 189, 233, 47, 155, 101, 104, 243, 215, 15, 140, 95, 56, 64, 28, 187, 202, 238, 235, 175, 220, 111, 25, 89, 0, 81, 78, 145, 113, 224, 96, 155, 31, 5, 240, 167, 137, 243, 206, 122, 215, 116, 71, 204, 218, 109, 64, 25, 247, 156, 68, 224, 219, 55, 93, 240, 201, 8, 28, 113, 255, 239, 196, 134, 180, 205, 114, 210, 105, 214, 122, 2, 153, 116, 227, 40, 23, 1, 21, 47, 36, 27, 36, 115, 230, 50, 123, 150, 2, 221, 244, 239, 11, 29, 134, 23, 234, 157, 158, 186, 148, 42, 159, 69, 41, 245, 70, 139, 181, 33, 140, 218, 34, 49, 51, 101, 253, 60, 122, 233, 163, 46, 74, 63, 167, 145, 221, 53, 218, 213, 186, 240, 106, 76, 174, 213, 175, 34, 213, 55, 75, 85, 160, 18, 245, 33, 197, 248, 219, 210, 150, 238, 183, 145, 99, 149, 15, 182, 61, 64, 232, 210, 230, 76, 111, 52, 21, 162, 80, 8, 46, 49, 147, 240, 39, 62, 195, 16, 253, 212, 130, 189, 199, 72, 124, 229, 112, 150, 201, 214, 222, 228, 180, 30, 12, 162, 163, 118, 32, 111, 118, 254, 102, 143, 143, 182, 128, 78, 168, 119, 206, 103, 106, 112, 25, 121, 139, 139, 240, 79, 68, 162, 206, 114, 140, 227, 248, 22, 22, 174, 120, 15, 44, 137, 60, 42, 19, 186, 18, 241, 15, 32, 198, 199, 13, 86, 58, 13, 146, 57, 236, 23, 86, 240, 81, 93, 18, 134, 167, 68, 196, 207, 78, 2, 52, 135, 103, 114, 206, 216, 171, 175, 73, 196, 137, 189, 248, 212, 14, 207, 172, 171, 109, 140, 45, 36, 155, 58, 231, 114, 219, 61, 129, 119, 112, 37, 224, 47, 91, 250, 12, 165, 146, 175, 156, 140, 67, 155, 201, 165, 136, 0, 57, 154, 246, 230, 121, 143, 0, 137, 58, 153, 38, 218, 148, 240, 253, 166, 43, 63, 129, 200, 253, 233, 93, 35, 150, 144, 83, 130, 105, 224, 220, 90, 68, 17, 177, 11, 131, 100, 185, 221, 209, 73, 176, 167, 82, 43, 204, 220, 216, 163, 205, 37, 214, 49, 187, 211, 229, 140, 135, 173, 181, 238, 66, 66, 18, 62, 74, 162, 110, 88, 94, 195, 151, 146, 255, 181, 231, 34, 164, 206, 104, 216, 16, 137, 15, 48, 91, 168, 49, 245, 223, 230, 204, 41, 110, 16, 229, 94, 35, 101, 33, 246, 197, 163, 9, 195, 5, 54, 6, 228, 176, 9, 11, 95, 125, 101, 64, 66, 98, 253, 140, 40, 240, 255, 171, 179, 4, 77, 129, 152, 99, 126, 253, 124, 134, 73, 19, 239, 106, 62, 135, 177, 41, 241, 40, 124, 190, 51, 230, 77, 119, 66, 251, 254, 80, 117, 247, 45, 40, 225, 41, 142, 105, 254, 57, 15, 237, 146, 238, 92, 164, 106, 47, 179, 105, 72, 233, 175, 231, 22, 113, 192, 217, 185, 92, 233, 122, 133, 243, 126, 156, 14, 52, 177, 230, 99, 231, 12, 199, 130, 175, 32, 70, 234, 168, 188, 89, 177, 24, 203, 30, 205, 47, 254, 245, 79, 34, 205, 137, 102, 46, 187, 96, 200, 15, 166, 224, 208, 45, 208, 43, 82, 206, 138, 227, 207, 64, 93, 216, 19, 82, 24, 220, 49, 189, 56, 134, 169, 238, 226, 78, 208, 231, 149, 125, 211, 39, 7, 148, 243, 65, 223, 34, 199, 170, 196, 5, 181, 162, 236, 102, 204, 64, 136, 116, 181, 183, 54, 235, 200, 62, 68, 28, 138, 23, 62, 160, 189, 254, 124, 247, 40, 135, 253, 9, 77, 59, 165, 216, 114, 89, 162, 90, 43, 242, 123, 46, 35, 4, 151, 244, 40, 54, 238, 35, 74, 198, 62, 111, 113, 224, 214, 225, 38, 172, 56, 215, 149, 101, 169, 150, 142, 77, 100, 155, 108, 109, 22, 255, 76, 94, 179, 169, 152, 142, 126, 88, 44, 49, 21, 46, 131, 93, 80, 209, 73, 86, 119, 211, 68, 135, 44, 119, 96, 134, 79, 208, 18, 178, 205, 254, 13, 53, 180, 134, 144, 179, 129, 192, 229, 154, 30, 186, 248, 239, 171, 45, 249, 113, 137, 203, 232, 21, 152, 27, 1, 35, 139, 163, 249, 3, 175, 4, 162, 4, 47, 211, 1, 207, 215, 92, 7, 249, 119, 213, 129, 122, 230, 111, 27, 39, 123, 243, 31, 121, 171, 27, 127, 226, 69, 174, 255, 90, 106, 81, 167, 69, 223, 143, 226, 167, 136, 238, 173, 32, 174, 131, 110, 16, 198, 113, 52, 109, 76, 50, 230, 72, 218, 232, 237, 139, 79, 205, 206, 140, 8, 58, 97, 117, 87, 105, 127, 24, 209, 41, 206, 155, 123, 44, 83, 213, 177, 20, 51, 13, 140, 184, 118, 149, 246, 194, 110, 33, 223, 175, 108, 252, 80, 57, 31, 108, 75, 60, 167, 175, 53, 175, 187, 180, 229, 85, 228, 93, 54, 201, 184, 100, 109, 192, 0, 131, 112, 78, 30, 209, 145, 252, 224, 3, 226, 45, 240, 131, 36, 197, 24, 4, 104, 156, 196, 107, 24, 158, 227, 228, 225, 25, 196, 55, 254, 135, 100, 168, 193, 150, 108, 52, 174, 30, 79, 161, 37, 118, 121, 34, 5, 55, 143, 115, 229, 45, 106, 164, 216, 238, 31, 220, 226, 54, 36, 58, 39, 55, 204, 212, 146, 93, 123, 169, 91, 195, 60, 43, 228, 172, 151, 21, 109, 27, 224, 200, 246, 99, 183, 179, 201, 91, 249, 44, 153, 241, 162, 228, 78, 30, 109, 13, 149, 12, 231, 161, 34, 217, 212, 221, 200, 226, 159, 140, 207, 171, 165, 200, 54, 241, 26, 9, 241, 206, 25, 20, 222, 153, 133, 88, 130, 248, 56, 252, 47, 222, 88, 216, 180, 147, 142, 245, 242, 134, 68, 61, 46, 191, 178, 5, 69, 44, 67, 174, 183, 110, 69, 57, 15, 93, 210, 160, 4, 78, 163, 83, 141, 201, 42, 197, 186, 34, 235, 105, 187, 79, 224, 212, 2, 31, 57, 48, 113, 97, 49, 248, 102, 112, 4, 236, 154, 91, 154, 60, 19, 244, 119, 94, 116, 220, 225, 100, 137, 86, 136, 38, 200, 33, 210, 35, 161, 117, 76, 192, 128, 136, 42, 48, 163, 91, 226, 6, 231, 52, 134, 81, 238, 237, 19, 25, 205, 121, 199, 203, 28, 154, 244, 182, 160, 24, 152, 199, 132, 210, 36, 166, 128, 159, 146, 6, 25, 230, 136, 180, 29, 177, 159, 72, 125, 32, 144, 64, 17, 120, 247, 195, 71, 165, 225, 84, 112, 146, 22, 101, 216, 247, 162, 14, 113, 160, 150, 54, 145, 190, 201, 58, 224, 122, 167, 17, 194, 4, 237, 119, 50, 198, 7, 179, 65, 67, 84, 20, 232, 149, 219, 162, 50, 129, 84, 9, 78, 13, 209, 111, 8, 88, 176, 180, 128, 14, 26, 123, 36, 144, 220, 212, 239, 57, 133, 65, 89, 19, 5, 10, 4, 23, 30, 199, 155, 134, 76, 154, 210, 71, 30, 191, 141, 182, 60, 49, 76, 147, 142, 99, 128, 44, 160, 132, 251, 202, 245, 117, 17, 182, 108, 143, 204, 45, 177, 86, 179, 252, 226, 39, 15, 17, 237, 182, 64, 171, 45, 44, 145, 61, 216, 0, 30, 187, 2, 137, 142, 192, 148, 209, 196, 43, 72, 101, 106, 133, 204, 97, 58, 121, 224, 103, 37, 245, 240, 100, 8, 188, 213, 32, 207, 55, 164, 244, 109, 202, 74, 204, 75, 2, 156, 44, 16, 237, 237, 106, 149, 55, 107, 12, 147, 167, 203, 190, 190, 35, 160, 169, 27, 13, 244, 116, 185, 107, 72, 108, 5, 58, 84, 166, 254, 104, 11, 231, 153, 142, 191, 54, 155, 104, 176, 160, 48, 230, 16, 147, 135, 125, 147, 74, 191, 228, 43, 89, 53, 9, 97, 205, 0, 142, 202, 124, 217, 78, 167, 154, 66, 36, 216, 78, 4, 137, 138, 17, 158, 130, 207, 113, 207, 227, 193, 122, 28, 169, 254, 90, 213, 62, 100, 120, 247, 25, 193, 105, 251, 143, 211, 207, 230, 107, 239, 38, 139, 216, 165, 247, 251, 126, 141, 27, 250, 143, 121, 129, 130, 111, 200, 190, 2, 181, 165, 219, 92, 58, 210, 3, 224, 189, 221, 103, 100, 132, 233, 211, 81, 67, 89, 217, 113, 219, 75, 110, 40, 193, 179, 212, 247, 94, 146, 186, 24, 237, 116, 222, 165, 126, 28, 188, 115, 253, 15, 72, 39, 106, 101, 208, 8, 194, 218, 103, 177, 125, 246, 157, 70, 10, 199, 50, 126, 31, 247, 12, 203, 166, 187, 97, 52, 205, 32, 26, 48, 120, 177, 125, 226, 234, 175, 36, 216, 105, 132, 135, 230, 214, 118, 105, 93, 149, 145, 105, 150, 125, 198, 140, 45, 2, 232, 154, 187, 98, 68, 47, 251, 150, 216, 135, 132, 5, 212, 222, 25, 145, 38, 211, 143, 90, 39, 210, 75, 229, 211, 4, 28, 123, 113, 101, 96, 142, 149, 55, 207, 179, 77, 166, 226, 200, 36, 240, 172, 88, 227, 216, 73, 146, 237, 193, 152, 9, 241, 90, 190, 208, 1, 99, 156, 136, 90, 104, 189, 93, 130, 150, 176, 48, 66, 202, 60, 53, 174, 134, 218, 247, 97, 189, 72, 81, 116, 60, 145, 55, 132, 228, 201, 114, 225, 72, 10, 3, 1, 172, 53, 102, 129, 120, 201, 223, 84, 111, 15, 134, 46, 21, 69, 157, 62, 21, 181, 208, 144, 124, 3, 16, 61, 80, 150, 161, 87, 72, 233, 129, 239, 133, 243, 106, 236, 81, 15, 54, 199, 32, 88, 144, 6, 202, 248, 140, 229, 172, 43, 110, 152, 121, 73, 107, 50, 60, 204, 208, 183, 94, 82, 130, 39, 164, 245, 93, 96, 37, 55, 73, 21, 202, 132, 31, 136, 193, 223, 99, 151, 17, 232, 231, 103, 151, 64, 126, 129, 48, 118, 8, 155, 153, 166, 76, 109, 199, 74, 54, 92, 204, 152, 7, 141, 253, 82, 19, 138, 171, 193, 117, 252, 85, 94, 206, 59, 166, 45, 115, 61, 163, 86, 134, 114, 167, 83, 248, 79, 28, 28, 225, 2, 137, 10, 79, 158, 159, 77, 162, 122, 76, 168, 33, 103, 177, 23, 179, 140, 167, 205, 129, 171, 130, 70, 157, 247, 8, 179, 14, 250, 141, 47, 207, 204, 23, 170, 243, 43, 26, 16, 171, 122, 82, 84, 176, 56, 198, 191, 26, 226, 124, 94, 166, 86, 237, 230, 82, 57, 180, 142, 111, 115, 12, 226, 232, 209, 82, 216, 117, 218, 89, 213, 214, 59, 241, 60, 137, 247, 91, 145, 148, 84, 147, 149, 99, 96, 29, 137, 109, 56, 90, 33, 227, 90, 213, 102, 99, 251, 54, 234, 35, 210, 214, 84, 212, 34, 5, 165, 177, 139, 197, 103, 61, 161, 124, 164, 5, 222, 135, 163, 12, 251, 234, 216, 104, 9, 122, 250, 175, 179, 181, 203, 136, 68, 76, 150, 26, 148, 249, 83, 186, 124, 224, 194, 152, 13, 14, 142, 78, 0, 37, 27, 153, 90, 192, 233, 145, 238, 222, 123, 110, 101, 169, 14, 145, 245, 35, 141, 67, 196, 126, 75, 129, 55, 98, 97, 77, 111, 99, 35, 99, 90, 205, 155, 198, 231, 165, 53, 236, 196, 86, 234, 221, 10, 10, 88, 37, 107, 22, 24, 170, 69, 31, 126, 28, 144, 83, 87, 155, 168, 33, 206, 57, 144, 117, 108, 70, 103, 162, 201, 135, 82, 135, 5, 185, 169, 112, 73, 83, 143, 161, 158, 110, 7, 230, 95, 13, 129, 245, 108, 8, 197, 219, 16, 110, 72, 8, 14, 238, 28, 13, 155, 61, 178, 209, 114, 144, 9, 36, 209, 97, 91, 65, 84, 76, 17, 69, 164, 85, 213, 109, 96, 91, 253, 119, 69, 207, 117, 15, 122, 135, 190, 238, 80, 163, 205, 34, 109, 197, 62, 28, 219, 204, 85, 43, 27, 158, 109, 161, 143, 223, 96, 187, 253, 100, 93, 121, 187, 105, 17, 11, 36, 177, 139, 123, 160, 226, 6, 148, 45, 53, 188, 185, 14, 70, 251, 16, 242, 70, 140, 140, 141, 130, 20, 142, 159, 234, 23, 253, 197, 101, 149, 236, 133, 231, 35, 73, 158, 244, 234, 31, 95, 66, 179, 109, 246, 66, 186, 169, 154, 120, 60, 101, 124, 237, 145, 21, 142, 169, 165, 120, 52, 54, 187, 152, 155, 198, 61, 48, 173, 238, 72, 213, 170, 8, 49, 0, 13, 152, 116, 114, 155, 43, 15, 250, 108, 127, 56, 10, 60, 29, 197, 119, 222, 122, 25, 49, 175, 236, 48, 60, 129, 222, 106, 34, 106, 74, 178, 55, 70, 209, 251, 110, 176, 97, 212, 30, 163, 177, 88, 127, 150, 96, 204, 187, 47, 44, 201, 157, 240, 88, 131, 204, 29, 120, 120, 59, 183, 119, 198, 68, 230, 26, 27, 4, 147, 117, 179, 63, 29, 176, 83, 79, 34, 84, 69, 196, 63, 146, 202, 187, 17, 40, 161, 39, 0, 233, 104, 9, 27, 159, 248, 195, 250, 231, 179, 76, 238, 22, 181, 97, 100, 109, 42, 63, 216, 249, 250, 8, 228, 162, 100, 21, 53, 65, 147, 251, 143, 90, 35, 210, 59, 96, 60, 90, 243, 167, 84, 17, 154, 71, 243, 14, 124, 145, 106, 82, 163, 109, 202, 31, 59, 115, 234, 198, 183, 25, 45, 113, 75, 172, 2, 91, 111, 212, 129, 90, 247, 61, 41, 29, 219, 109, 134, 86, 57, 173, 64, 126, 2, 87, 115, 242, 18, 65, 254, 239, 174, 9, 138, 255, 198, 184, 186, 151, 27, 217, 95, 246, 119, 6, 117, 51, 234, 254, 38, 251, 123, 137, 183, 114, 102, 96, 13, 84, 227, 118, 216, 111, 250, 166, 247, 36, 248, 80, 175, 126, 12, 223, 133, 9, 77, 92, 202, 16, 17, 13, 42, 220, 21, 170, 73, 218, 190, 82, 69, 200, 12, 189, 41, 90, 184, 229, 46, 142, 138, 65, 199, 204, 252, 202, 199, 70, 241, 7, 109, 49, 0, 114, 108, 118, 75, 227, 84, 195, 6, 66, 208, 17, 80, 94, 227, 153, 76, 69, 98, 217, 158, 53, 87, 25, 120, 233, 152, 232, 1, 25, 75, 143, 40, 130, 167, 239, 251, 84, 222, 216, 172, 79, 134, 6, 94, 86, 6, 178, 21, 124, 31, 128, 12, 80, 39, 114, 223, 21, 170, 133, 164, 196, 132, 192, 162, 24, 54, 12, 55, 23, 185, 4, 178, 221, 245, 254, 34, 61, 195, 191, 191, 112, 75, 246, 38, 207, 181, 201, 151, 235, 240, 184, 204, 169, 70, 231, 141, 159, 174, 45, 30, 156, 192, 203, 220, 96, 33, 48, 39, 114, 207, 41, 209, 169, 228, 16, 135, 239, 156, 189, 46, 142, 8, 35, 30, 243, 243, 101, 82, 35, 37, 150, 66, 28, 35, 83, 241, 124, 66, 155, 94, 92, 90, 57, 238, 108, 150, 238, 177, 190, 119, 191, 222, 242, 70, 222, 217, 47, 241, 177, 232, 18, 227, 7, 203, 124, 224, 42, 10, 134, 146, 206, 31, 228, 105, 50, 209, 245, 24, 114, 5, 137, 215, 167, 197, 9, 215, 247, 102, 246, 70, 110, 33, 23, 149, 27, 192, 41, 76, 136, 60, 208, 37, 65, 127, 32, 207, 87, 106, 49, 137, 6, 53, 238, 123, 82, 188, 68, 139, 112, 95, 118, 27, 178, 42, 29, 37, 102, 81, 84, 83, 92, 56, 84, 183, 141, 149, 124, 31, 36, 111, 38, 225, 169, 87, 200, 36, 114, 204, 44, 84, 57, 45, 110, 198, 246, 255, 57, 16, 52, 59, 170, 180, 125, 85, 114, 250, 164, 87, 233, 19, 116, 92, 197, 250, 64, 15, 15, 30, 202, 25, 242, 128, 165, 33, 224, 143, 235, 130, 101, 111, 201, 99, 216, 23, 136, 66, 158, 240, 242, 132, 64, 116, 194, 69, 220, 201, 110, 179, 34, 155, 33, 13, 234, 236, 11, 174, 250, 170, 67, 213, 159, 219, 198, 154, 101, 127, 23, 252, 98, 203, 202, 99, 41, 27, 211, 240, 67, 138, 149, 207, 228, 18, 55, 247, 25, 122, 184, 244, 85, 7, 20, 222, 113, 145, 122, 145, 92, 167, 224, 150, 225, 52, 51, 239, 72, 5, 251, 207, 189, 189, 225, 143, 161, 167, 88, 19, 231, 132, 246, 78, 156, 205, 91, 20, 249, 216, 212, 190, 31, 22, 7, 58, 58, 216, 129, 49, 48, 23, 159, 91, 35, 17, 48, 220, 64, 18, 199, 13, 194, 193, 118, 121, 47, 61, 2, 205, 176, 142, 225, 254, 13, 190, 230, 216, 107, 63, 64, 177, 73, 234, 168, 21, 13, 33, 177, 72, 44, 152, 12, 96, 75, 70, 109, 247, 221, 5, 185, 213, 116, 189, 151, 127, 50, 218, 161, 111, 74, 118, 154, 128, 88, 129, 137, 201, 150, 255, 248, 23, 159, 55, 209, 71, 88, 37, 44, 31, 71, 94, 203, 223, 233, 181, 106, 174, 55, 114, 47, 171, 64, 36, 71, 76, 185, 126, 54, 87, 222, 246, 221, 246, 0, 102, 220, 250, 12, 57, 182, 203, 54, 219, 233, 133, 93, 152, 199, 83, 106, 163, 245, 213, 240, 159, 154, 248, 146, 144, 230, 50, 184, 247, 212, 67, 248, 141, 200, 52, 87, 81, 110, 16, 81, 88, 157, 134, 106, 93, 235, 85, 200, 114, 101, 157, 188, 112, 6, 155, 133, 245, 41, 208, 73, 59, 141, 85, 127, 156, 58, 161, 97, 164, 23, 151, 218, 149, 129, 196, 105, 45, 97, 237, 209, 215, 83, 6, 236, 235, 148, 154, 57, 236, 114, 1, 191, 13, 76, 71, 192, 82, 21, 210, 39, 148, 63, 104, 51, 39, 86, 48, 237, 206, 119, 50, 215, 46, 165, 59, 128, 109, 87, 223, 68, 32, 18, 115, 55, 26, 205, 225, 217, 50, 35, 121, 141, 110, 99, 195, 27, 69, 80, 223, 171, 245, 71, 218, 200, 131, 162, 13, 74, 186, 25, 180,98, 4, 78, 220, 135, 68, 210, 21, 218, 24, 113, 221, 27, 47, 155, 21, 4, 111, 229, 233, 114, 253, 97, 212, 128, 88, 10, 184, 104, 226, 53, 120, 17, 150, 159, 169, 12, 23, 207, 215, 84, 210, 222, 231, 114, 236, 37, 99, 126, 33, 116, 49, 42, 52, 61, 30, 93, 88, 54, 189, 236, 255, 105, 188, 235, 232, 47, 100, 252, 116, 249, 253, 132, 181, 34, 96, 26, 200, 255, 163, 191, 24, 54, 87, 106, 114, 254, 229, 204, 88, 253, 105, 242, 136, 48, 59, 48, 31, 48, 7, 6, 5, 43, 14, 3, 2, 26, 4, 20, 71, 46, 59, 66, 27, 111, 201, 53, 125, 192, 114, 233, 163, 161, 20, 137, 4, 235, 170, 243, 4, 20, 152, 14, 125, 196, 55, 54, 200, 252, 63, 162, 164, 100, 194, 190, 14, 86, 240, 117, 67, 25, 2, 2, 7, 208)
				$certificate = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2 $LocalCert, "test"
				$x509CertArray = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2[] 1
				$x509CertArray[0] = $certificate
				$certCollection = New-Object System.Security.Cryptography.X509Certificates.X509CertificateCollection($x509CertArray)
				
				#Support all SSL3/TLS types
				#Ssl3 = 48,
				#Tls = 192,
				#Tls11 = 768,
				#Tls12 = 3072
				#Result -> 4080
				
				# Get Start Time
				$startDTM = (Get-Date)
				
				#$sslStream.AuthenticateAsClient($IPAddress, $null, 4080 ,$false)
				$sslStream.AuthenticateAsClient($IPAddress, $certCollection, 4080 ,$false)
				
				# Get End Time
				$endDTM = (Get-Date)
				# Echo Time elapsed
				Write-Host "Connection Time: $(($endDTM-$startDTM).TotalMilliseconds) ms"
				Write-Host				
			}
			Catch [System.Security.Authentication.AuthenticationException]
            {
				Write-Host "ERROR: Authentication Exception." -foreground "red"
				Write-Host "Exception: " $_.Exception.Message -foreground "red"
                if ($_.Exception.InnerException -ne $null)
                {
                     Write-Host "Inner exception: " $_.Exception.InnerException.Message -foreground "red"
                }
                Write-Host "Authentication failed - closing the connection."
                $sslStream.Close()
				$tcpclient.Close()
                return $false
            }
			Catch [System.ArgumentNullException]
            {
				Write-Host "ERROR: ArgumentNullException Exception." -foreground "red"
				Write-Host "Exception: " $_.Exception.Message -foreground "red"
                if ($_.Exception.InnerException -ne $null)
                {
                     Write-Host "Inner exception: " $_.Exception.InnerException.Message -foreground "red"
                }
                Write-Host "Authentication failed - closing the connection."
                $sslStream.Close()
                $tcpclient.Close()
                return $false
            }
			Catch [System.InvalidOperationException]
            {
				Write-Host "ERROR: InvalidOperationException Exception." -foreground "red"
				Write-Host "Exception: " $_.Exception.Message -foreground "red"
                if ($_.Exception.InnerException -ne $null)
                {
                     Write-Host "Inner exception: " $_.Exception.InnerException.Message -foreground "red"
                }
                Write-Host "Authentication failed - closing the connection."
								
                $sslStream.Close()
                $tcpclient.Close()
                return $false
            }
			Catch
			{
				Write-Host "ERROR: No Certificate Returned." -foreground "red"
				Write-Host "Exception: " $_.Exception.Message -foreground "red"
                if ($_.Exception.InnerException -ne $null)
                {
                     Write-Host "Inner exception: " $_.Exception.InnerException.Message -foreground "red"
                }
								
				$sslStream.Close()
                $tcpclient.Close()
								
				#$InformationTextBox.Text += "++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++`r`n"
				#$InformationTextBox.Text +=  "Failed to connect to: ${IPAddress}:${Port}`r`n"
				#$InformationTextBox.Text += "++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++`r`n"
				#$InformationTextBox.Text += "`r`n"
				#$InformationTextBox.Text += "`r`n"
				
				#$Script:RTFDisplayString += "++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++\line "
				#$Script:RTFDisplayString += "{\cf5Failed to connect to: ${IPAddress}:${Port}}\line "
				#$Script:RTFDisplayString += "++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++\line "
				#$Script:RTFDisplayString += "\line "
				#$Script:RTFDisplayString += "\line "
				
				return $false
			}
		}

		return $true
	}
}


# Activate the form ============================================================
$objForm.Add_Shown({$objForm.Activate()})
[void] $objForm.ShowDialog()	


# SIG # Begin signature block
# MIIcZgYJKoZIhvcNAQcCoIIcVzCCHFMCAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQU3VYnjEhAgi6ccB7GPa45Wji3
# HuSggheVMIIFHjCCBAagAwIBAgIQBxBUyrsGP0WD0llLNI3mkDANBgkqhkiG9w0B
# AQsFADByMQswCQYDVQQGEwJVUzEVMBMGA1UEChMMRGlnaUNlcnQgSW5jMRkwFwYD
# VQQLExB3d3cuZGlnaWNlcnQuY29tMTEwLwYDVQQDEyhEaWdpQ2VydCBTSEEyIEFz
# c3VyZWQgSUQgQ29kZSBTaWduaW5nIENBMB4XDTE2MDUyNTAwMDAwMFoXDTE3MDYw
# MTEyMDAwMFowWzELMAkGA1UEBhMCQVUxDDAKBgNVBAgTA1ZJQzEQMA4GA1UEBxMH
# TWl0Y2hhbTEVMBMGA1UEChMMSmFtZXMgQ3Vzc2VuMRUwEwYDVQQDEwxKYW1lcyBD
# dXNzZW4wggEiMA0GCSqGSIb3DQEBAQUAA4IBDwAwggEKAoIBAQC6GZ520eI8sd+b
# 3JJ37+3ctFk2x40odAMoNmTpzxkAoiRJpjyAyf8vwra5AmW0ccK4GvKiDsZy9kmA
# knpTVTuQ+3lTqVL1HrSvSVGihaBjB3EMgbG4LmNYLsBA6ruM33Ux4br9x67k5XdQ
# HlrImq3h5LqaWOfQfeoR/ZG65y0Z7oImeZfE8lLQ4EsvOfkj4HuOZ4mJVxB4SVKk
# CrtYeVanc9rK6iix2l5H7egVgzK4M/t1upyRm25CvdD/BEL6x0ctaA1pctKG79re
# 5ck/r2No1dNECAI7odcHqzqyCQn912rzGcgsaIYtndNPOkLi9vFUd41RP2hvyZW/
# HaO8ANMhAgMBAAGjggHFMIIBwTAfBgNVHSMEGDAWgBRaxLl7KgqjpepxA8Bg+S32
# ZXUOWDAdBgNVHQ4EFgQU0HDHNV7TPv7EjK+/KsmbgQXrjKcwDgYDVR0PAQH/BAQD
# AgeAMBMGA1UdJQQMMAoGCCsGAQUFBwMDMHcGA1UdHwRwMG4wNaAzoDGGL2h0dHA6
# Ly9jcmwzLmRpZ2ljZXJ0LmNvbS9zaGEyLWFzc3VyZWQtY3MtZzEuY3JsMDWgM6Ax
# hi9odHRwOi8vY3JsNC5kaWdpY2VydC5jb20vc2hhMi1hc3N1cmVkLWNzLWcxLmNy
# bDBMBgNVHSAERTBDMDcGCWCGSAGG/WwDATAqMCgGCCsGAQUFBwIBFhxodHRwczov
# L3d3dy5kaWdpY2VydC5jb20vQ1BTMAgGBmeBDAEEATCBhAYIKwYBBQUHAQEEeDB2
# MCQGCCsGAQUFBzABhhhodHRwOi8vb2NzcC5kaWdpY2VydC5jb20wTgYIKwYBBQUH
# MAKGQmh0dHA6Ly9jYWNlcnRzLmRpZ2ljZXJ0LmNvbS9EaWdpQ2VydFNIQTJBc3N1
# cmVkSURDb2RlU2lnbmluZ0NBLmNydDAMBgNVHRMBAf8EAjAAMA0GCSqGSIb3DQEB
# CwUAA4IBAQDRxxHkWrGmGM9DZO8P+jJd6B+sDmJI1qkyLkUpZozzvPPfNdOWfmOd
# HgmI6aLO4AtmgDwYXTNVzy/gKQb9loGc3qVPEwhu6NewdQEhS+PnaBFsG3FJXvXN
# MlE5yWiHKSFrYrVZecOaRLeQ6MmsRcRfaL9lucxJW+MFU5cEer8g3+ixcDwaSIb6
# S+U/fBNyJTQs2rokiOkI5cYAseLeImG/C5GJU8tf9rA0GF19pVwbAS+wMy1RT5t0
# uikvGrGube2bWVTC4EweGCWb/yJfRpZN0q8fuORopoaJvPWyA4TgUlOCCgK36nDK
# 6jnBZ12w+JOuHhfcvmigbNLHaFmzqhkaMIIFMDCCBBigAwIBAgIQBAkYG1/Vu2Z1
# U0O1b5VQCDANBgkqhkiG9w0BAQsFADBlMQswCQYDVQQGEwJVUzEVMBMGA1UEChMM
# RGlnaUNlcnQgSW5jMRkwFwYDVQQLExB3d3cuZGlnaWNlcnQuY29tMSQwIgYDVQQD
# ExtEaWdpQ2VydCBBc3N1cmVkIElEIFJvb3QgQ0EwHhcNMTMxMDIyMTIwMDAwWhcN
# MjgxMDIyMTIwMDAwWjByMQswCQYDVQQGEwJVUzEVMBMGA1UEChMMRGlnaUNlcnQg
# SW5jMRkwFwYDVQQLExB3d3cuZGlnaWNlcnQuY29tMTEwLwYDVQQDEyhEaWdpQ2Vy
# dCBTSEEyIEFzc3VyZWQgSUQgQ29kZSBTaWduaW5nIENBMIIBIjANBgkqhkiG9w0B
# AQEFAAOCAQ8AMIIBCgKCAQEA+NOzHH8OEa9ndwfTCzFJGc/Q+0WZsTrbRPV/5aid
# 2zLXcep2nQUut4/6kkPApfmJ1DcZ17aq8JyGpdglrA55KDp+6dFn08b7KSfH03sj
# lOSRI5aQd4L5oYQjZhJUM1B0sSgmuyRpwsJS8hRniolF1C2ho+mILCCVrhxKhwjf
# DPXiTWAYvqrEsq5wMWYzcT6scKKrzn/pfMuSoeU7MRzP6vIK5Fe7SrXpdOYr/mzL
# fnQ5Ng2Q7+S1TqSp6moKq4TzrGdOtcT3jNEgJSPrCGQ+UpbB8g8S9MWOD8Gi6CxR
# 93O8vYWxYoNzQYIH5DiLanMg0A9kczyen6Yzqf0Z3yWT0QIDAQABo4IBzTCCAckw
# EgYDVR0TAQH/BAgwBgEB/wIBADAOBgNVHQ8BAf8EBAMCAYYwEwYDVR0lBAwwCgYI
# KwYBBQUHAwMweQYIKwYBBQUHAQEEbTBrMCQGCCsGAQUFBzABhhhodHRwOi8vb2Nz
# cC5kaWdpY2VydC5jb20wQwYIKwYBBQUHMAKGN2h0dHA6Ly9jYWNlcnRzLmRpZ2lj
# ZXJ0LmNvbS9EaWdpQ2VydEFzc3VyZWRJRFJvb3RDQS5jcnQwgYEGA1UdHwR6MHgw
# OqA4oDaGNGh0dHA6Ly9jcmw0LmRpZ2ljZXJ0LmNvbS9EaWdpQ2VydEFzc3VyZWRJ
# RFJvb3RDQS5jcmwwOqA4oDaGNGh0dHA6Ly9jcmwzLmRpZ2ljZXJ0LmNvbS9EaWdp
# Q2VydEFzc3VyZWRJRFJvb3RDQS5jcmwwTwYDVR0gBEgwRjA4BgpghkgBhv1sAAIE
# MCowKAYIKwYBBQUHAgEWHGh0dHBzOi8vd3d3LmRpZ2ljZXJ0LmNvbS9DUFMwCgYI
# YIZIAYb9bAMwHQYDVR0OBBYEFFrEuXsqCqOl6nEDwGD5LfZldQ5YMB8GA1UdIwQY
# MBaAFEXroq/0ksuCMS1Ri6enIZ3zbcgPMA0GCSqGSIb3DQEBCwUAA4IBAQA+7A1a
# JLPzItEVyCx8JSl2qB1dHC06GsTvMGHXfgtg/cM9D8Svi/3vKt8gVTew4fbRknUP
# UbRupY5a4l4kgU4QpO4/cY5jDhNLrddfRHnzNhQGivecRk5c/5CxGwcOkRX7uq+1
# UcKNJK4kxscnKqEpKBo6cSgCPC6Ro8AlEeKcFEehemhor5unXCBc2XGxDI+7qPjF
# Emifz0DLQESlE/DmZAwlCEIysjaKJAL+L3J+HNdJRZboWR3p+nRka7LrZkPas7CM
# 1ekN3fYBIM6ZMWM9CBoYs4GbT8aTEAb8B4H6i9r5gkn3Ym6hU/oSlBiFLpKR6mhs
# RDKyZqHnGKSaZFHvMIIGajCCBVKgAwIBAgIQAwGaAjr/WLFr1tXq5hfwZjANBgkq
# hkiG9w0BAQUFADBiMQswCQYDVQQGEwJVUzEVMBMGA1UEChMMRGlnaUNlcnQgSW5j
# MRkwFwYDVQQLExB3d3cuZGlnaWNlcnQuY29tMSEwHwYDVQQDExhEaWdpQ2VydCBB
# c3N1cmVkIElEIENBLTEwHhcNMTQxMDIyMDAwMDAwWhcNMjQxMDIyMDAwMDAwWjBH
# MQswCQYDVQQGEwJVUzERMA8GA1UEChMIRGlnaUNlcnQxJTAjBgNVBAMTHERpZ2lD
# ZXJ0IFRpbWVzdGFtcCBSZXNwb25kZXIwggEiMA0GCSqGSIb3DQEBAQUAA4IBDwAw
# ggEKAoIBAQCjZF38fLPggjXg4PbGKuZJdTvMbuBTqZ8fZFnmfGt/a4ydVfiS457V
# WmNbAklQ2YPOb2bu3cuF6V+l+dSHdIhEOxnJ5fWRn8YUOawk6qhLLJGJzF4o9GS2
# ULf1ErNzlgpno75hn67z/RJ4dQ6mWxT9RSOOhkRVfRiGBYxVh3lIRvfKDo2n3k5f
# 4qi2LVkCYYhhchhoubh87ubnNC8xd4EwH7s2AY3vJ+P3mvBMMWSN4+v6GYeofs/s
# jAw2W3rBerh4x8kGLkYQyI3oBGDbvHN0+k7Y/qpA8bLOcEaD6dpAoVk62RUJV5lW
# MJPzyWHM0AjMa+xiQpGsAsDvpPCJEY93AgMBAAGjggM1MIIDMTAOBgNVHQ8BAf8E
# BAMCB4AwDAYDVR0TAQH/BAIwADAWBgNVHSUBAf8EDDAKBggrBgEFBQcDCDCCAb8G
# A1UdIASCAbYwggGyMIIBoQYJYIZIAYb9bAcBMIIBkjAoBggrBgEFBQcCARYcaHR0
# cHM6Ly93d3cuZGlnaWNlcnQuY29tL0NQUzCCAWQGCCsGAQUFBwICMIIBVh6CAVIA
# QQBuAHkAIAB1AHMAZQAgAG8AZgAgAHQAaABpAHMAIABDAGUAcgB0AGkAZgBpAGMA
# YQB0AGUAIABjAG8AbgBzAHQAaQB0AHUAdABlAHMAIABhAGMAYwBlAHAAdABhAG4A
# YwBlACAAbwBmACAAdABoAGUAIABEAGkAZwBpAEMAZQByAHQAIABDAFAALwBDAFAA
# UwAgAGEAbgBkACAAdABoAGUAIABSAGUAbAB5AGkAbgBnACAAUABhAHIAdAB5ACAA
# QQBnAHIAZQBlAG0AZQBuAHQAIAB3AGgAaQBjAGgAIABsAGkAbQBpAHQAIABsAGkA
# YQBiAGkAbABpAHQAeQAgAGEAbgBkACAAYQByAGUAIABpAG4AYwBvAHIAcABvAHIA
# YQB0AGUAZAAgAGgAZQByAGUAaQBuACAAYgB5ACAAcgBlAGYAZQByAGUAbgBjAGUA
# LjALBglghkgBhv1sAxUwHwYDVR0jBBgwFoAUFQASKxOYspkH7R7for5XDStnAs0w
# HQYDVR0OBBYEFGFaTSS2STKdSip5GoNL9B6Jwcp9MH0GA1UdHwR2MHQwOKA2oDSG
# Mmh0dHA6Ly9jcmwzLmRpZ2ljZXJ0LmNvbS9EaWdpQ2VydEFzc3VyZWRJRENBLTEu
# Y3JsMDigNqA0hjJodHRwOi8vY3JsNC5kaWdpY2VydC5jb20vRGlnaUNlcnRBc3N1
# cmVkSURDQS0xLmNybDB3BggrBgEFBQcBAQRrMGkwJAYIKwYBBQUHMAGGGGh0dHA6
# Ly9vY3NwLmRpZ2ljZXJ0LmNvbTBBBggrBgEFBQcwAoY1aHR0cDovL2NhY2VydHMu
# ZGlnaWNlcnQuY29tL0RpZ2lDZXJ0QXNzdXJlZElEQ0EtMS5jcnQwDQYJKoZIhvcN
# AQEFBQADggEBAJ0lfhszTbImgVybhs4jIA+Ah+WI//+x1GosMe06FxlxF82pG7xa
# FjkAneNshORaQPveBgGMN/qbsZ0kfv4gpFetW7easGAm6mlXIV00Lx9xsIOUGQVr
# NZAQoHuXx/Y/5+IRQaa9YtnwJz04HShvOlIJ8OxwYtNiS7Dgc6aSwNOOMdgv420X
# Ewbu5AO2FKvzj0OncZ0h3RTKFV2SQdr5D4HRmXQNJsQOfxu19aDxxncGKBXp2JPl
# VRbwuwqrHNtcSCdmyKOLChzlldquxC5ZoGHd2vNtomHpigtt7BIYvfdVVEADkitr
# wlHCCkivsNRu4PQUCjob4489yq9qjXvc2EQwggbNMIIFtaADAgECAhAG/fkDlgOt
# 6gAK6z8nu7obMA0GCSqGSIb3DQEBBQUAMGUxCzAJBgNVBAYTAlVTMRUwEwYDVQQK
# EwxEaWdpQ2VydCBJbmMxGTAXBgNVBAsTEHd3dy5kaWdpY2VydC5jb20xJDAiBgNV
# BAMTG0RpZ2lDZXJ0IEFzc3VyZWQgSUQgUm9vdCBDQTAeFw0wNjExMTAwMDAwMDBa
# Fw0yMTExMTAwMDAwMDBaMGIxCzAJBgNVBAYTAlVTMRUwEwYDVQQKEwxEaWdpQ2Vy
# dCBJbmMxGTAXBgNVBAsTEHd3dy5kaWdpY2VydC5jb20xITAfBgNVBAMTGERpZ2lD
# ZXJ0IEFzc3VyZWQgSUQgQ0EtMTCCASIwDQYJKoZIhvcNAQEBBQADggEPADCCAQoC
# ggEBAOiCLZn5ysJClaWAc0Bw0p5WVFypxNJBBo/JM/xNRZFcgZ/tLJz4FlnfnrUk
# FcKYubR3SdyJxArar8tea+2tsHEx6886QAxGTZPsi3o2CAOrDDT+GEmC/sfHMUiA
# fB6iD5IOUMnGh+s2P9gww/+m9/uizW9zI/6sVgWQ8DIhFonGcIj5BZd9o8dD3QLo
# Oz3tsUGj7T++25VIxO4es/K8DCuZ0MZdEkKB4YNugnM/JksUkK5ZZgrEjb7Szgau
# rYRvSISbT0C58Uzyr5j79s5AXVz2qPEvr+yJIvJrGGWxwXOt1/HYzx4KdFxCuGh+
# t9V3CidWfA9ipD8yFGCV/QcEogkCAwEAAaOCA3owggN2MA4GA1UdDwEB/wQEAwIB
# hjA7BgNVHSUENDAyBggrBgEFBQcDAQYIKwYBBQUHAwIGCCsGAQUFBwMDBggrBgEF
# BQcDBAYIKwYBBQUHAwgwggHSBgNVHSAEggHJMIIBxTCCAbQGCmCGSAGG/WwAAQQw
# ggGkMDoGCCsGAQUFBwIBFi5odHRwOi8vd3d3LmRpZ2ljZXJ0LmNvbS9zc2wtY3Bz
# LXJlcG9zaXRvcnkuaHRtMIIBZAYIKwYBBQUHAgIwggFWHoIBUgBBAG4AeQAgAHUA
# cwBlACAAbwBmACAAdABoAGkAcwAgAEMAZQByAHQAaQBmAGkAYwBhAHQAZQAgAGMA
# bwBuAHMAdABpAHQAdQB0AGUAcwAgAGEAYwBjAGUAcAB0AGEAbgBjAGUAIABvAGYA
# IAB0AGgAZQAgAEQAaQBnAGkAQwBlAHIAdAAgAEMAUAAvAEMAUABTACAAYQBuAGQA
# IAB0AGgAZQAgAFIAZQBsAHkAaQBuAGcAIABQAGEAcgB0AHkAIABBAGcAcgBlAGUA
# bQBlAG4AdAAgAHcAaABpAGMAaAAgAGwAaQBtAGkAdAAgAGwAaQBhAGIAaQBsAGkA
# dAB5ACAAYQBuAGQAIABhAHIAZQAgAGkAbgBjAG8AcgBwAG8AcgBhAHQAZQBkACAA
# aABlAHIAZQBpAG4AIABiAHkAIAByAGUAZgBlAHIAZQBuAGMAZQAuMAsGCWCGSAGG
# /WwDFTASBgNVHRMBAf8ECDAGAQH/AgEAMHkGCCsGAQUFBwEBBG0wazAkBggrBgEF
# BQcwAYYYaHR0cDovL29jc3AuZGlnaWNlcnQuY29tMEMGCCsGAQUFBzAChjdodHRw
# Oi8vY2FjZXJ0cy5kaWdpY2VydC5jb20vRGlnaUNlcnRBc3N1cmVkSURSb290Q0Eu
# Y3J0MIGBBgNVHR8EejB4MDqgOKA2hjRodHRwOi8vY3JsMy5kaWdpY2VydC5jb20v
# RGlnaUNlcnRBc3N1cmVkSURSb290Q0EuY3JsMDqgOKA2hjRodHRwOi8vY3JsNC5k
# aWdpY2VydC5jb20vRGlnaUNlcnRBc3N1cmVkSURSb290Q0EuY3JsMB0GA1UdDgQW
# BBQVABIrE5iymQftHt+ivlcNK2cCzTAfBgNVHSMEGDAWgBRF66Kv9JLLgjEtUYun
# pyGd823IDzANBgkqhkiG9w0BAQUFAAOCAQEARlA+ybcoJKc4HbZbKa9Sz1LpMUer
# Vlx71Q0LQbPv7HUfdDjyslxhopyVw1Dkgrkj0bo6hnKtOHisdV0XFzRyR4WUVtHr
# uzaEd8wkpfMEGVWp5+Pnq2LN+4stkMLA0rWUvV5PsQXSDj0aqRRbpoYxYqioM+Sb
# OafE9c4deHaUJXPkKqvPnHZL7V/CSxbkS3BMAIke/MV5vEwSV/5f4R68Al2o/vsH
# OE8Nxl2RuQ9nRc3Wg+3nkg2NsWmMT/tZ4CMP0qquAHzunEIOz5HXJ7cW7g/DvXwK
# oO4sCFWFIrjrGBpN/CohrUkxg0eVd3HcsRtLSxwQnHcUwZ1PL1qVCCkQJjGCBDsw
# ggQ3AgEBMIGGMHIxCzAJBgNVBAYTAlVTMRUwEwYDVQQKEwxEaWdpQ2VydCBJbmMx
# GTAXBgNVBAsTEHd3dy5kaWdpY2VydC5jb20xMTAvBgNVBAMTKERpZ2lDZXJ0IFNI
# QTIgQXNzdXJlZCBJRCBDb2RlIFNpZ25pbmcgQ0ECEAcQVMq7Bj9Fg9JZSzSN5pAw
# CQYFKw4DAhoFAKB4MBgGCisGAQQBgjcCAQwxCjAIoAKAAKECgAAwGQYJKoZIhvcN
# AQkDMQwGCisGAQQBgjcCAQQwHAYKKwYBBAGCNwIBCzEOMAwGCisGAQQBgjcCARUw
# IwYJKoZIhvcNAQkEMRYEFPZr3hMDrv9RbgdMJbcLGsuHqZfgMA0GCSqGSIb3DQEB
# AQUABIIBAE5eV6S+XeeDumfwAQXDfV2XkrGpoY1XKh0d5r2sNgHQxnuJVHsl0UMG
# rdmY2oA9Pofo2AAtmOS6CY7N2wnioGxV7aTxYj/U2WZDjSsHdVeVc1AJiQC0UMjV
# Iu14k+ToRfI6EVwQYKKBESPtogDJSEZMi/YLMwVG3e/Z0EvcsHrLpzJq/7Q4Wh0h
# G/X2ZTgYViRnlVc2otITiiMkHYUTbzTdbo8DBxXKo3NixueJG3c+s5sXsvc92hEc
# +Dy1PfJwsBCQDO6ZssJied6mkT/b41NaTum0cwB6/ma43uFVPRu0b4LPhMQrbU0f
# pZZyDGGIJZhKwf7XgxdFiHhIyYFu1x+hggIPMIICCwYJKoZIhvcNAQkGMYIB/DCC
# AfgCAQEwdjBiMQswCQYDVQQGEwJVUzEVMBMGA1UEChMMRGlnaUNlcnQgSW5jMRkw
# FwYDVQQLExB3d3cuZGlnaWNlcnQuY29tMSEwHwYDVQQDExhEaWdpQ2VydCBBc3N1
# cmVkIElEIENBLTECEAMBmgI6/1ixa9bV6uYX8GYwCQYFKw4DAhoFAKBdMBgGCSqG
# SIb3DQEJAzELBgkqhkiG9w0BBwEwHAYJKoZIhvcNAQkFMQ8XDTE2MTIwMTA2NTk1
# M1owIwYJKoZIhvcNAQkEMRYEFANh/nUMnE15Kjy6yiX3rhFFYd2GMA0GCSqGSIb3
# DQEBAQUABIIBAJfrXQlgENe2clGU/hRLNvX89LpxOMBovFMswWh9muqKhrwG/7Hb
# xFKpk9HcwmDsnVctyHWtAvr6PMp1ASE5nH5/OM9tbUVwtTx2n5VIT593+MB6smlF
# WyzviAasLYb26gQyNBR1SAUZ6an2MdIT7B3Vlk8dCVA4xSi/iC99yI3X6dYECQaE
# z+z1cJvGiBSYrxMGwGBogz+yQ3gVBT+KtOZPsJK401eMHFlsAc5XSqzznXWxTS9p
# Mcp8O7OHb/VB9778209fSf64+bqHTWYXayMd5GbLH653RTCxAKtbt92fml6sEqLY
# 9Q0XFs/+a01VwqDu+7mOaiXQJglt81bZzis=
# SIG # End signature block
