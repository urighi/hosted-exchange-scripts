<#
.SYNOPSIS 
Import-Mailboxes.ps1 - Imports mailbox data to the destination organization.

.DESCRIPTION
Imports mailboxes and user accounts, and imports PST files to mailboxes. Also imports
Imports user accounts and distribution groups from a CSV file, enables the
distribution groups and adds members to them.
Refer to PrepareHostedDomain.ps1 and MailboxExport.ps1 for the accepted input files 
format.

.OUTPUTS
Results are output to screen.  Errors are output to screen and the MailboxImport.log
file.

.PARAMETER InitialPSTImport
When this parameter is used, an initial import will be made, creating AD accounts,
enabling mailboxes, configuring secondary e-mail addresses and importing the PST
files. The path set on the MailboxListPath parameter will be used for the mailbox
list. Takes precedence over FinalPSTImport.

.PARAMETER FinalPSTImport
When this parameter is used, a final import will be performed, creating a new 
import request in Exchange but not creating any user accounts or mailboxes.
The path set on the MailboxListPath parameter will be used for the mailbox
list.

.PARAMETER EnableADAccount
When this parameter is used, the Active Directory accounts will be enabled after
being created.

.PARAMETER ImportDistributionGroups
When this parameter is used, Distribution Groups will be imported from the file
set on the DistributionGroupListPath parameter.

.PARAMETER ImportPermissions

When this parameter is used, permissions will be set to each mailbox according
to the file set on the PermissionListPath parameter.

.PARAMETER ImportForwardingAddresses
When this parameter is used, the forwarding address will be set to each mailbox
according to the file set on the ForwardingAddressesListPath parameter. Make sure
that distribution groups are being imported on the same script execution or have
been imported before.

.PARAMETER ImportDelegates
When this parameter is used, delegates will be added to the corresponding folders,
according to the the file set on the DelegateListPath parameter.
Accounts must be enabled first in order to allow for rights assignment.

.PARAMETER PSTPath
Path (UNC or local) to the exported PST files.

.PARAMETER MailboxListPath
Path (UNC or local) to the mailbox list (CSV).

.PARAMETER DistributionGroupListPath
Path (UNC or local) to the distribution group list (CSV).

.PARAMETER PermissionListPath
Path (UNC or local) to the permissions list (CSV).

.PARAMETER ForwardingAddressesListPath
Path (UNC or local) to the forwarding address list (CSV).

.PARAMETER DelegateListPath
Path (UNC or local) to the delegate list (CSV).

.PARAMETER ClientName
The client name which will be used for accessing the mailbox database and 
organizational unit in AD.

.PARAMETER MailboxDatabase
The mailbox database to be used. If left empty, the "HOSTED - $ClientName"
mailbox will be used.

.PARAMETER PasswordLength
The password length that will be set for all new Active Directory accounts.
Default is 4 characters. Passwords will be the first 4 letters on the client
name, followed by a dot and 4 lowercase/numeric random characters.

.PARAMETER OUPath
Distinguished name for the Organizational Unit that holds all clients.

.EXAMPLE
.\Import-Mailboxes.ps1 -InitialPSTImport -ImportDistributionGroups -PSTPath "E:\PSTFiles" -MailboxListPath "E:\PSTFiles\Mailboxes.csv" -DistributionGroupListPath "E:\PSTFiles\DG.csv" -ClientName "Righi.com"
Imports mailboxes and distribution groups from the configured paths, adds user
accounts to the "Righi.com" organizational unit in Active Directory, adds user
mailboxes to the "HOSTED - Righi.com" mailbox database.

.EXAMPLE
.\Import-Mailboxes.ps1 -FinalPSTImport -PSTPath "E:\PSTFiles" -MailboxListPath "E:\PSTFiles\Mailboxes.csv" -ClientName "Righi.com"


.NOTES
Requires the Active Directory PowerShell module.
Make sure you assign the Mailbox Import Export role to the Organization Management
group before running the PST exports and imports.
New-ManagementRoleAssignment -Name "Mailbox Import Export" -SecurityGroup "Organization Management" -Role "Mailbox Import Export"

Written by Ulisses Righi
ulisses@ulisoft.com.br
Version 1.9.2
6/20/2018

#>

[CmdletBinding()]
param(
	[Parameter(Mandatory=$false)]
	[switch]$InitialPSTImport,
	
	[Parameter(Mandatory=$false)]
	[switch]$FinalPSTImport,
	
	[Parameter(Mandatory=$false)]
	[switch]$EnableADAccount,
	
	[Parameter(Mandatory=$false)]
	[switch]$ImportDistributionGroups,
	
	[Parameter(Mandatory=$false)]
	[switch]$ImportPermissions,
	
	[Parameter(Mandatory=$false)]
	[switch]$ImportForwardingAddresses,
	
	[Parameter(Mandatory=$false)]
	[switch]$ImportDelegates,
		
	[Parameter(Mandatory=$false)]
	[switch]$UseMigrationSuffix,
	
	[Parameter(Mandatory=$false)] [ValidateNotNullOrEmpty()]
	[string]$PSTPath = "\\mbx1\c$\exchange",
	
	[Parameter(Mandatory=$false)] [ValidateNotNullOrEmpty()]
	[string]$MailboxListPath,
	
	[Parameter(Mandatory=$false)] [ValidateNotNullOrEmpty()]
	[string]$DistributionGroupListPath,
	
	[Parameter(Mandatory=$false)] [ValidateNotNullOrEmpty()]
	[string]$PermissionListPath,
	
	[Parameter(Mandatory=$false)] [ValidateNotNullOrEmpty()]
	[string]$ForwardingAddressesListPath,
	
	[Parameter(Mandatory=$false)] [ValidateNotNullOrEmpty()]
	[string]$DelegateListPath,
	
	[Parameter(Mandatory=$true)] [ValidateNotNullOrEmpty()]
	[string]$ClientName,
	
	[Parameter( Mandatory=$false)]
	[string]$BatchName,
	
	[Parameter(Mandatory=$false)] [ValidateRange(4,30)]
	[int]$PasswordLength = 5,
		
	[Parameter(Mandatory=$false)] [ValidateNotNullOrEmpty()]
	[string]$MailboxDatabase,
	
	[Parameter(Mandatory=$false)] [ValidateNotNullOrEmpty()]
	[string]$OUPath = "OU=Hosted Organizations,DC=hosting,DC=email",
	
	[Parameter(Mandatory=$false)]
	[bool]$Confirm = $true,

	[Parameter(Mandatory=$false)]
	[switch]$SkipADAccountCreation

	)
	
$RecipientContainer = "OU=" + $ClientName + "," + $OUPath
$PSTPath = $PSTPath.TrimEnd("\")

if (!$MailboxDatabase) { $MailboxDatabase = "HOSTED - " + $ClientName }
$ScriptLogFile = "MailboxImport.log"
$ErrorActionPreference = "Stop"

if (!$MailboxListPath) { $MailboxListPath = $PSTPath + "\MailboxList.csv"}

if (!$DistributionGroupListPath) { $DistributionGroupListPath = $PSTPath + "\DistributionGroupList.csv"}

if (!$PermissionListPath) { $PermissionListPath = $PSTPath + "\PermissionList.csv" }

if (!$ForwardingAddressesListPath) { $ForwardingAddressesListPath = $PSTPath + "\ForwardingList.csv" }

if (!$DelegateListPath) { $DelegateListPath = $PSTPath + "\DelegateList.csv" }

. ./New-Password.ps1

Function WriteToLog 
{
   Param ([string]$Details)
   $LogString = (Get-Date -format G) + " " + $Details
   Add-Content $ScriptLogFile -value $LogString
}

Function CreateADAccount {
	Param ($Mailbox)
	#Creates AD accounts
			$Alias = $Mailbox.Alias
			if ((Get-AdObject -LDAPFilter "(sAMAccountName=$Alias)") -ne $null)
			{
				$Alias = ($Alias + "_" + $($ClientName.Replace(" ","")))
			}
			if ($Alias.Length -gt 20) { $Alias = $Alias.Substring(0,20) }
			$AccountPassword = (($ClientName).Substring(0,2)) + (New-Password -Length $PasswordLength -Lowercase -Numeric)
			New-ADUser -Name $Mailbox.DisplayName -GivenName $Mailbox.GivenName -Surname $Mailbox.Surname -SamAccountName $Alias -UserPrincipalName $Mailbox.UserPrincipalName -AccountPassword (ConvertTo-SecureString -AsPlainText $AccountPassword -Force) -Path $RecipientContainer -Confirm:$Confirm
			if (($EnableADAccount) -and ($Mailbox.RecipientTypeDetails -eq "UserMailbox")) 
			{ 
				Enable-AdAccount $Alias 
			}
			
						
			# Adds password to table
			$PasswordTable += New-Object PSObject -Property @{DisplayName=$($Mailbox.DisplayName);EmailAddress=$($Mailbox.PrimarySMTPAddress);Login=$($Mailbox.UserPrincipalName);Password=$($AccountPassword)}
			$PasswordTable | Export-CSV "$PSTPath\$($ClientName)_Passwords.csv" -Append
}

Function ConfigureMailbox {
	Param ($Mailbox)
	#Enables Exchange mailbox
			if ($UseMigrationSuffix) { $Mailbox.PrimarySMTPAddress += ".migration"  }
			if ($Mailbox.RecipientTypeDetails -eq "UserMailbox")
			{
				if (($Mailbox.Database -ne "") -and ($Mailbox.Database -ne $null))
				{
					$LocalMailboxDatabase = $Mailbox.Database
				}
				else
				{
					$LocalMailboxDatabase = $MailboxDatabase
				}
				WriteToLog "Creating mailbox for $($Mailbox.UserPrincipalName)"
				Enable-Mailbox -Identity $Mailbox.UserPrincipalName -Database $LocalMailboxDatabase -DisplayName $Mailbox.DisplayName -PrimarySMTPAddress $Mailbox.PrimarySMTPAddress -AddressBookPolicy $ClientName -Confirm:$Confirm
			}
			elseif ($Mailbox.RecipientTypeDetails -eq "RoomMailbox") 
			{
				WriteToLog "Creating mailbox for $($Mailbox.UserPrincipalName)"
				Enable-Mailbox -Identity $Mailbox.UserPrincipalName -Database $LocalMailboxDatabase -DisplayName $Mailbox.DisplayName -PrimarySMTPAddress $Mailbox.PrimarySMTPAddress -Room -Confirm:$Confirm
			}
			
			#Adds secondary e-mail addresses
			$EmailAddresses = $Mailbox.EmailAddresses -Split ';'
			if ($EmailAddresses.Count -gt 1)
			{
				foreach ($EmailAddress in $EmailAddresses)
				{
					if (($EmailAddress -like '*smtp*') -and ($EmailAddress -notlike '*.local*'))
					{
						if ($UseMigrationSuffix) { $EmailAddress += ".migration" }
						Write-Host "Adding $EmailAddress for $($Mailbox.UserPrincipalName)" -ForegroundColor Cyan
						WriteToLog "Adding $EmailAddress for $($Mailbox.UserPrincipalName)"
						Set-Mailbox $Mailbox.UserPrincipalName -EmailAddresses @{Add=$EmailAddress} -EmailAddressPolicyEnabled $false -Confirm:$Confirm
					}
					if ($EmailAddress -like "*x500*")
					{
						Write-Host "Adding $EmailAddress for $($Mailbox.UserPrincipalName)" -ForegroundColor Cyan
						WriteToLog "Adding $EmailAddress for $($Mailbox.UserPrincipalName)"
						Set-Mailbox $Mailbox.UserPrincipalName -EmailAddresses @{Add=$EmailAddress} -EmailAddressPolicyEnabled $false -Confirm:$Confirm
					}
				}
			}
			
			#Adds X500

			if (($Mailbox.X500 -ne $null) -and ($Mailbox.X500 -ne ""))
			{
				$X500 = "X500:" + $Mailbox.X500
				Write-Host "Adding $X500 for $($Mailbox.UserPrincipalName)"
				WriteToLog "Adding $X500 for $($Mailbox.UserPrincipalName)"
				Set-Mailbox $Mailbox.UserPrincipalName -EmailAddresses @{Add=$X500} -Confirm:$Confirm
			}
			
			ConfigurePublicFolder($Mailbox)
}

Function ConfigurePublicFolder {
	Param ($Mailbox)
	WriteToLog "Setting Public Folder for $($Mailbox.UserPrincipalName)"
	#Set-Mailbox $Mailbox.UserPrincipalName -DefaultPublicFolderMailbox "$ClientName Public Folders"
	Add-PublicFolderClientPermission -Identity "\$ClientName Public Folders" -User $Mailbox.UserPrincipalName -AccessRights Editor
	
}

Function InitialPSTImport {
	try
	{
		Write-Host "Starting initial PST import..." -ForegroundColor Cyan
		
		$PasswordTable = @()
		
		$MailboxList = Import-CSV $MailboxListPath
		
		if ($BatchName -eq "") { $BatchName = "$ClientName Initial Import" }
		
		$i = 0
		foreach ($Mailbox in $MailboxList)
		{
			try
			{
				if (!$SkipADAccountCreation)
				{
					$i++
					Write-Progress -Activity "Creating AD account for" -status "$($Mailbox.UserPrincipalName)" -percentComplete ($i / ((($MailboxList.Count)+1)*100))
					WriteToLog "Creating AD account for $($Mailbox.UserPrincipalName)"
					 CreateADAccount($Mailbox)
				}
			}
			catch
			{
				WriteToLog $_.Exception.Message
				Write-Host $_.Exception.Message -ForegroundColor Red
				Write-Host "Error while creating AD account. Please refer to $ScriptLogFile for details." -ForegroundColor Red
			}
		}

		Write-Host "Waiting for AD replication..." -ForegroundColor Cyan
		
		$AccountCheck = $null
		while ($AccountCheck -eq $null)
		{
			try
			{
				$Upn = $MailboxList[-1].UserPrincipalName
				$AccountCheck = Get-ADUser -Filter { UserPrincipalName -eq $upn }	
			}
			catch
			{
				Start-Sleep -Seconds 10	
			}
		}


		$i = 0
		foreach ($Mailbox in $MailboxList)
		{
			try
			{
				$i++
				Write-Progress -Activity "Configuring mailbox for" -status "$($Mailbox.UserPrincipalName)" -percentComplete ($i / ((($MailboxList.Count)+1)*100))
				WriteToLog "Configuring mailbox for $($Mailbox.UserPrincipalName)"
				
				ConfigureMailbox($Mailbox)
				
				#Imports mailboxes
			}
			catch
			{
				WriteToLog $_.Exception.Message
				Write-Host $_.Exception.Message -ForegroundColor Red
				Write-Host "Error while creating mailbox. Please refer to $ScriptLogFile for details." -ForegroundColor Red
			}
		}

		Write-Host "Waiting for AD replication..." -ForegroundColor Cyan
		Start-Sleep -Seconds 10

		$AccountCheck = $null
		while ($AccountCheck -eq $null)
		{
			try
			{
				$AccountCheck = Get-Mailbox $MailboxList[-1].UserPrincipalName	
			}
			catch
			{
				Start-Sleep -Seconds 10	
			}
		}

		$i = 0
		foreach ($Mailbox in $MailboxList)
		{
			try
			{
				$i++
				Write-Progress -Activity "Importing mailbox for" -status "$($Mailbox.UserPrincipalName)" -percentComplete ($i / ((($MailboxList.Count)+1)*100))
				WriteToLog "Importing mailbox for $($Mailbox.UserPrincipalName)"
				
				#Imports mailboxes
				New-MailboxImportRequest -Mailbox $Mailbox.UserPrincipalName -FilePath "$PSTPath\$($Mailbox.UserPrincipalName).pst" -TargetRootFolder "/" -LargeItemLimit 100 -BadItemLimit 50 -AcceptLargeDataLoss -BatchName $BatchName -Name $BatchName -Confirm:$Confirm
			}
			catch
			{
				WriteToLog $_.Exception.Message
				Write-Host $_.Exception.Message -ForegroundColor Red
				Write-Host "Error while creating mailbox. Please refer to $ScriptLogFile for details." -ForegroundColor Red
			}
		}
		
		
		Get-MailboxImportRequest -BatchName $BatchName | Get-MailboxImportRequestStatistics
		Write-Host "Passwords exported to $PSTPath\$($ClientName)_Passwords.csv" -ForegroundColor Cyan
		Write-Host "Run Get-MailboxImportRequest -BatchName `"$BatchName`" | Get-MailboxImportRequestStatistics to get the import status." -ForegroundColor Cyan
		Write-Host "Run Get-MailboxImportRequest -BatchName `"$BatchName`" -Status Completed | Remove-MailboxImportRequest to remove completed requests." -ForegroundColor Cyan
		
	}
	catch
	{
		WriteToLog $_.Exception.Message
		Write-Host $_.Exception.Message -ForegroundColor Red
		Write-Host "Error while importing mailbox. Please refer to $ScriptLogFile for details." -ForegroundColor Red
	}

}

Function FinalPSTImport {
	try
	{
		Write-Host "Starting final PST import..." -ForegroundColor Cyan
		$i = 0
		$MailboxList = Import-CSV $MailboxListPath
		
		if ($BatchName -eq "") { $BatchName = "$ClientName Final Import" }
		
		foreach ($Mailbox in $MailboxList)
		{
			$i++
			Write-Progress -Activity "Importing mailbox for" -status "$($Mailbox.UserPrincipalName)" -percentComplete ($i / ((($MailboxList.Count)+1)*100))
			
			WriteToLog "Importing mailbox for $($Mailbox.UserPrincipalName) from $PSTPath\$($Mailbox.UserPrincipalName).pst"
			Write-Host "Importing mailbox for $($Mailbox.UserPrincipalName) from $PSTPath\$($Mailbox.UserPrincipalName).pst" -ForegroundColor Cyan
			
			New-MailboxImportRequest -Mailbox $Mailbox.UserPrincipalName -FilePath "$PSTPath\$($Mailbox.UserPrincipalName).pst" -TargetRootFolder "/"  -LargeItemLimit 100 -BadItemLimit 50 -AcceptLargeDataLoss -BatchName $BatchName -Name $BatchName -Priority Emergency -Confirm:$Confirm
			if (($EnableADAccount) -and ($Mailbox.RecipientTypeDetails -eq "UserMailbox")) 
			{ 
				WriteToLog "Enabling Active Directory account for $($Mailbox.UserPrincipalName)"
				$Alias = (Get-Mailbox $Mailbox.UserPrincipalName).SamAccountName
				Enable-AdAccount $Alias
			}
		}
		
		Get-MailboxImportRequest -BatchName $BatchName | Get-MailboxImportRequestStatistics
		Write-Host "Run Get-MailboxImportRequest -BatchName `"$BatchName`" | Get-MailboxImportRequestStatistics to get the import status." -ForegroundColor Cyan
		Write-Host "Run Get-MailboxImportRequest -BatchName `"$BatchName`" -Status Completed | Remove-MailboxImportRequest to remove completed requests." -ForegroundColor Cyan
		Write-Host "Run .\SendPasswords.ps1 -PasswordFilePath $ClientName_Passwords.csv to send passwords to users." -ForegroundColor Cyan
	}
	catch
	{
		WriteToLog $_.Exception.Message
		Write-Host $_.Exception.Message -ForegroundColor Red
		Write-Host "Error while importing mailbox. Please refer to $ScriptLogFile for details." -ForegroundColor Red
	}
}

Function ImportDistributionGroups {
	try
	{
		Write-Host "Importing distribution groups..." -ForegroundColor Cyan
		$i = 1
		$DistributionGroupList = Import-CSV $DistributionGroupListPath
		foreach ($DistributionGroup in $DistributionGroupList)
		{
			if ($UseMigrationSuffix) { $DistributionGroup.PrimarySMTPAddress += ".migration" }
			Write-Progress -Id 1 -Activity "Importing distribution group" -status "$DistributionGroup.Name" -percentComplete ($i / ((($DistributionGroupList.Count)+1)*100))
			WriteToLog "Importing distribution group $($DistributionGroup.Name)"
			New-DistributionGroup -Name $DistributionGroup.Name -Alias $DistributionGroup.Alias -Type Distribution -PrimarySmtpAddress $DistributionGroup.PrimarySMTPAddress -OrganizationalUnit $RecipientContainer -Confirm:$Confirm
			$Members = $DistributionGroup.Members -Split ';'
			$x = 1
			if ($Members -ne "" -and $Members -ne $null)
			{
				foreach ($Member in $Members)
				{
					$MemberSMTP = (($Member -Split '=')[-1]).TrimEnd('}')
					Write-Progress -Id 2 -ParentId 1 -Activity "Importing distribution group member" -Status $Member -PercentComplete ($x++ / ((($DistributionGroupList.Count)+1)*100)) -ErrorAction SilentlyContinue
					if ($MemberSMTP -ne "")
					{ 
						WriteToLog "Adding $MemberSMTP to $($DistributionGroup.Name)"
						Add-DistributionGroupMember -Identity $DistributionGroup.PrimarySMTPAddress -Member $MemberSMTP -Confirm:$Confirm
					}
				}
			}
		}
	}
	catch 
	{
		WriteToLog $_.Exception.Message
		Write-Host $_.Exception.Message -ForegroundColor Red
		Write-Host "Error while importing distribution group. Please refer to $ScriptLogFile for details." -ForegroundColor Red
	}
}

Function ImportPermissions {
	$Permissions = Import-CSV $PermissionListPath
	foreach ($Permission in $Permissions) 
	{
		if ($Permission.User -notlike "*Admin*")
		{ 
			WriteToLog "Adding $($Permission.AccessRights) permissions to $($Permission.User) on $($Permission.Alias)"
			if ($Permission.AccessRights -eq "FullAccess")
			{
				Add-MailboxPermission -Identity $Permission.Alias -User $Permission.User -AccessRights $Permission.AccessRights -InheritanceType All -Confirm:$Confirm
			}
			elseif ($Permission.AccessRights -eq "SendOnBehalf")
			{
				Add-AdPermission -Identity (Get-Mailbox $Permission.Alias).Identity -User (Get-Mailbox $Permission.User).Identity -ExtendedRights "Send-As" -Confirm:$Confirm
			}
		}
	}
}

Function ImportForwardingAddresses {
	$Forwardings = Import-CSV $ForwardingAddressesListPath
	foreach ($Forwarding in $Forwardings)
	{
		$FwdAddress = $Forwarding.ForwardingAddress.Split('/')[-1]
		$logstring = "Forwarding" + $Forwarding.DisplayName + "(" + (Get-Mailbox $Forwarding.DisplayName).PrimarySMTPAddress + ") to" + (Get-Mailbox $FwdAddress).DisplayName +  "(" + (Get-Mailbox $FwdAddress).PrimarySMTPAddress + ")"
		Write-Host $logstring -ForegroundColor Cyan
		WriteToLog $logstring
		Set-Mailbox $Forwarding.DisplayName -DeliverToMailboxAndForward ([System.Convert]::ToBoolean($Forwarding.DeliverToMailboxAndForward)) -ForwardingAddress $FwdAddress -Confirm:$Confirm
	}
}

function ImportDelegates
{
	$Delegates = Import-CSV $DelegateListPath
	foreach ($Delegate in $Delegates)
	{
		$AccessRights = [Microsoft.Exchange.Management.StoreTasks.MailboxFolderAccessRight[]]($Delegate.AccessRights.Split(","))
		WriteToLog "Adding $AccessRights permissions to $($Delegate.User) on $($Delegate.Folderkey)"
		Add-MailboxFolderPermission -Identity $Delegate.FolderKey -User $Delegate.User -AccessRights $AccessRights
	}
}

#Initialize

Import-Module ActiveDirectory
if ($InitialPSTImport) { InitialPSTImport }
elseif ($FinalPSTImport) { FinalPSTImport }
if ($ImportDistributionGroups) { ImportDistributionGroups }
if ($ImportPermissions) { ImportPermissions }
if ($ImportForwardingAddresses) { ImportForwardingAddresses }
if ($ImportDelegates) { ImportDelegates }