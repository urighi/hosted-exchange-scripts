<#
.SYNOPSIS 
Export-Mailboxes.ps1 - Export mailbox or mailbox information from the originating domain.

.DESCRIPTION
Generates a list of accepted domains, mailboxes and distribution groups. Creates export
requests for Exchange mailboxes, according to a set date.

.OUTPUTS
Results are output to screen. Mailbox list is output to the MailboxList.csv file.
Distribution group list is output to the DistributionGroupList.csv file. Initial
PST export date is output to the date.txt file.

.PARAMETER GetExchangeInfo
Displays accepted domains, mailbox databases and e-mail address policies.

.PARAMETER ExportMailboxList
Generates a user and room mailbox list in the CSV format, containing the Alias,
DisplayName, EmailAddresses, X500, GivenName, Surname, PrimarySMTPAddress, 
RecipientTypeDetails, UserPrincipalName properties. Also exports the mailbox
permissions and forwarding address to a CSV file.

.PARAMETER ExportDistributionGroups
Generates a distribution group list in the CSV format, containing the Name,
DisplayName, GroupType, Alias, EmailAddresses, PrimarySMTPAddress and Members
properties.

.PARAMETER InitialPSTExport
Creates export requests for all user and room mailboxes. Will also create a date.txt
file containing the current date and time to be used by the FinalPSTExport function.

.PARAMETER FinalPSTExport
Creates export requests for all user and room mailboxes. Will use the date.txt file
or the ExportDate parameter as the initial date filter for exporting mailbox data.

.PARAMETER Path
UNC path where the PST files will be stored.

.PARAMETER ExportDate
Date that will be used on the FinalPSTExport parameter if no date.txt file is present.

.PARAMETER MailboxList
List of mailboxes to be exported. If this parameter is not set, all mailboxes are exported.

.EXAMPLE
.\Export-Mailboxes.ps1 -GetExchangeInfo -ExportDistributionGroups -Path "\\exmbx1\c$\export"
Displays Exchange info and exports the distribution group list.

.EXAMPLE 
.\Export-Mailboxes.ps1 -ExportMailboxList -InitialPSTExport -Path "\\exmbx1\c$\export" -Confirm:$false
Exports the mailbox list and creates mailbox export requests.

.EXAMPLE 
.\Export-Mailboxes.ps1 -FinalPSTExport -Path "\\exmbx1\c$\export" -Date "4/19/2016 01:00PM" -Confirm:$false
Creates mailbox export requests filtering items older than 4/19/2016 01:00PM.

.NOTES
Requires the Active Directory PowerShell module.
Make sure you assign the Mailbox Import Export role to the Organization Management
group before running the PST exports and imports.
New-ManagementRoleAssignment -Name "Mailbox Import Export" -SecurityGroup "Organization Management" -Role "Mailbox Import Export"

Written by Ulisses Righi
ulisses@ulisoft.com.br
Version 1.9.2
7/5/2018

#>

[CmdletBinding()]
param(
	[Parameter( Mandatory=$false)]
	[switch]$GetExchangeInfo,
	
	[Parameter( Mandatory=$false)]
	[switch]$ExportMailboxList,
	
	[Parameter( Mandatory=$false)]
	[switch]$ExportDistributionGroups,
	
	[Parameter( Mandatory=$false)]
	[switch]$InitialPSTExport,
	
	[Parameter( Mandatory=$false)]
	[switch]$FinalPSTExport,
	
	[Parameter( Mandatory=$false)]
	[string]$MailboxList,
	
	[Parameter( Mandatory=$true)][ValidateScript({Test-Path $_ -PathType Container})]
	[string]$Path,
	
	[Parameter( Mandatory=$false)]
	[string]$ExportDate,
	
	[Parameter( Mandatory=$false)]
	[string]$BatchName,
	
	[Parameter(Mandatory=$false)]
	[bool]$Confirm = $true
	)
	

# Variables

#if ($Path) { $Path = "\\mbx1\c$\exchange" }
$Path = $Path.TrimEnd("\")

Function GetExchangeInfo {
	Write-Host "Accepted Domains:"
	Get-AcceptedDomain
	Write-Host "E-mail address policies:"
	Get-EmailAddressPolicy | select-object EnabledEmailAddressTemplates,DisabledEmailAddressTemplates,RecipientFilter,LdapRecipientFilter,RecipientFilterApplied
	Write-Host "Mailbox databases:"
	$MailboxDBs = Get-MailboxDatabase -Status
	foreach ($MailboxDatabase in $MailboxDBs) 
	{ 
		$count = ($MailboxDatabase | Get-Mailbox).Count
		$out = New-Object PSObject
		$out | Add-Member NoteProperty MailboxName $MailboxDatabase.Name
		$out | Add-Member NoteProperty MailboxCount $count
		$out | Add-Member NoteProperty DatabaseSize $MailboxDatabase.DatabaseSize
		$out | Add-Member NoteProperty ProhibitSendQuota $MailboxDatabase.ProhibitSendQuota
		$out | Add-Member NoteProperty ProhibitSendReceiveQuota $MailboxDatabase.ProhibitSendReceiveQuota
		$out | Add-Member NoteProperty RecoverableItemsQuota $MailboxDatabase.RecoverableItemsQuota
		$out | Add-Member NoteProperty RecoverableItemsWarningQuota $MailboxDatabase.RecoverableItemsWarningQuota
		if ($count -gt 0) {
			$size = ($MailboxDatabase.DatabaseSize.ToMB()/$count).ToString("00.00") + " MB"
			$out | Add-Member noteproperty AverageMailboxSize $size
		}
	Write-Output $out
	}
}

Function ExportMailboxList {
	Write-Host "Exporting mailbox list..." -ForegroundColor Yellow
	$MbxListPath = $Path + '\' + "MailboxList.csv"
	if ($MailboxList)
	{
		$Mailbox = Get-Content $MailboxList | Get-Mailbox
	}
	else
	{
		$Mailbox = Get-Mailbox
	}
	$Mailbox | Where-Object {($_.RecipientTypeDetails -eq "UserMailbox") -or ($_.RecipientTypeDetails -eq "RoomMailbox")} | Select-Object Alias,DisplayName,@{Name='EmailAddresses';Expression={[string]::join(";", ($_.EmailAddresses))}},@{Name='GivenName';Expression={(Get-AdUser $_.Alias).GivenName}},@{Name='Surname';Expression={(Get-AdUser $_.Alias).Surname}},@{Name='X500';Expression={(Get-AdUser $_.Alias -Properties legacyExchangeDN).legacyExchangeDN}},PrimarySMTPAddress,RecipientTypeDetails,UserPrincipalName | Export-Csv $MbxListPath
	Write-Host "Mailbox list saved to $MbxListPath" -ForegroundColor Yellow
	
	Write-Host "Exporting mailbox permissions..." -ForegroundColor Yellow
	$PermissionListPath = $Path + '\' + "PermissionList.csv"
	#TODO: Check if user is a group
	$perms = $Mailbox | Get-MailboxPermission | where {($_.User -notlike "*admin*") -and ($_.User -notlike "*organization*") -and ($_.User -notlike "*exchange*") -and ($_.User -notlike "*nt authority*") -and ($_.User -notlike "*delegated setup*") -and ($_.User -notlike "*public folder*") -and ($_.User -notlike "*mail operators*") -and ($_.User -notlike "*discovery management*") }
	$perms | Select-Object @{Name='Alias';Expression={(Get-Mailbox $_.Identity).PrimarySMTPAddress}},@{Name='User';Expression={(Get-Mailbox $_.User.ToString()).PrimarySMTPAddress}},@{Name='AccessRights';Expression={[string]::join(', ', $_.AccessRights)}} | Export-CSV $PermissionListPath -NoTypeInformation
	Write-Host "Permission list saved to $PermissionListPath"
	
	$ForwardingListPath = $Path + '\' + "ForwardingList.csv"
	
	$Mailbox | where {($_.ForwardingAddress -ne $null) -and ($_.ForwardingAddress -ne "")} | Select-Object PrimarySMTPAddress,@{Name='ForwardingAddress';Expression={(Get-Mailbox $_.ForwardingAddress).PrimarySMTPAddress}} | Export-CSV $ForwardingListPath
	
	ExportDelegates
}

Function ExportDistributionGroups {
		Write-Host "Exporting distribution groups..." -ForegroundColor Yellow
		$DGListPath = $Path + '\' + "DistributionGroupList.csv"
		$DGMemberListPath = $Path + '\' + "DistributionGroupMemberList.csv"
		$DGList = @()
		$DGMemberList = @()
		$DGs = Get-DistributionGroup | Where-Object {($_.RecipientTypeDetails -eq "MailUniversalSecurityGroup") -or ($_.RecipientTypeDetails -eq "MailUniversalDistributionGroup")}
		foreach ($DG in $DGs)
		{
			$GroupMembers = Get-DistributionGroupMember $DG
			$out = New-Object PSObject
			$out | Add-Member noteproperty Name $DG.Name
			$out | Add-Member noteproperty DisplayName $DG.DisplayName
			$out | Add-Member noteproperty Alias $DG.Alias
			$out | Add-Member noteproperty EmailAddresses ([string]::join(";", ($DG.EmailAddresses)))
			$out | Add-Member noteproperty PrimarySMTPAddress $DG.PrimarySMTPAddress
			$out | Add-Member noteproperty Members ([string]::join(";", ($GroupMembers | Select-Object PrimarySMTPAddress)))
			$DGList += $out

			foreach ($Member in ($GroupMembers))
			{
				$out = New-Object PSObject
				$out | Add-Member NoteProperty Group $DG.PrimarySMTPAddress
				$out | Add-Member NoteProperty Member $Member.PrimarySMTPAddress
				$DGMemberList += $out
			}
		}

		$DGList | Export-CSV $DGListPath
		$DGMemberList | Export-CSV $DGMemberListPath
		Write-Host "Distribution group list saved to $DGListPath" -ForegroundColor Yellow
}

function ExportDelegates {
	Write-Host "Exporting delegates..." -ForegroundColor Yellow

	if ($MailboxList)
	{
		$AllUsers = Get-Content $MailboxList | Get-Mailbox -RecipientTypeDetails 'UserMailbox' -ResultSize Unlimited	
	}
	else
	{
		$AllUsers = Get-Mailbox -RecipientTypeDetails 'UserMailbox' -ResultSize Unlimited	
	}
	
	$DelegateListPath = $Path + "\DelegateList.csv"
	$DelegateList = @()

	foreach ($Alias in $AllUsers)
	{
		$Mailbox = $Alias.UserPrincipalName
		$Folders = Get-MailboxFolderStatistics $Mailbox | %{$_.folderpath} | %{$_.replace("/","\")}

		foreach ($F in $Folders)
		{
			try
			{
				$FolderKey = $Mailbox + ":" + $F
				$Permissions = Get-MailboxFolderPermission -Identity $FolderKey -ErrorAction SilentlyContinue | Where-Object {$_.User -notlike "Default" -and $_.User -notlike "Anonymous" -and $_.AccessRights -notlike "None" }
				foreach ($Perm in $Permissions)
				{
					if ($Perm.User -ne $null)
					{
						$User = $Perm.User.ToString()
						if ($User.StartsWith("NT User:"))
						{
							$User = $User.Substring(8)
						}
						
						$User = (Get-Mailbox $User -Verbose).PrimarySMTPAddress
						
						if ($User -ne $Alias.PrimarySMTPAddress)
						{
							$out = New-Object PSObject
							$out | Add-Member noteproperty PrimarySMTPAddress $Alias.PrimarySMTPAddress
							$out | Add-Member noteproperty User $User
							$out | Add-Member noteproperty FolderKey $FolderKey
							$out | Add-Member noteproperty AccessRights ([string]::join(',', $Perm.AccessRights))
							$DelegateList += $out
						}
					}
				}
			}
			catch {
				continue
			}
		}
	}
	$DelegateList | Export-CSV $DelegateListPath
	Write-Host "Delegate list exported to $DelegateListPath" -ForegroundColor Yellow
}
	
Function InitialPSTExport {
		$Date = Get-Date
		$DatePath = $Path + '\' + 'date.txt'
		($Date | Out-String).Trim() | Out-File $DatePath
		if ($MailboxList)
		{
			$Mailboxes = Get-Content $MailboxList | Get-Mailbox
		}
		else
		{
			$Mailboxes = Get-Mailbox | Where-Object {($_.RecipientTypeDetails -eq "UserMailbox") -or ($_.RecipientTypeDetails -eq "RoomMailbox")}
		}
		
		$i = 0
		if ($BatchName -eq "") { $BatchName = "InitialExport" }
		foreach ($Mailbox in $Mailboxes)
		{
			$FilePath = $Path + '\' + $Mailbox.UserPrincipalName + '.pst'
			New-MailboxExportRequest -Mailbox $Mailbox -ContentFilter "Received -le '$Date'" -FilePath $FilePath -BatchName $BatchName -Name $BatchName -BadItemLimit 50 -AcceptLargeDataLoss -Confirm:$Confirm
			$i++
			Write-Progress -Activity "Creating export request for" -status "$Mailbox" -percentComplete ($i / (($Mailboxes.Count)+1)*100)
		}
		Write-Host "Run Get-MailboxExportRequest | Get-MailboxExportRequestStatistics to get the export status."
		Write-Host "Run Get-MailboxExportRequest -Status Completed | Remove-MailboxExportRequest to remove completed requests."
}

Function FinalPSTExport {
		$FinalDate = Get-Date
		$FinalDatePath = $Path + '\' + 'finaldate.txt'
		($FinalDate | Out-String).Trim() | Out-File $FinalDatePath
		
		if (!$ExportDate) { $ExportDate = Get-Content $DatePath }
		
		$i = 0
		if ($BatchName -eq "") { $BatchName = "FinalExport" }
		if ($MailboxList)
		{
			$Mailboxes = $MailboxList | Get-Mailbox
		}
		else
		{
			$Mailboxes = Get-Mailbox | Where-Object {($_.RecipientTypeDetails -eq "UserMailbox") -or ($_.RecipientTypeDetails -eq "RoomMailbox")}
		}
		foreach ($Mailbox in $Mailboxes)
		{
			$FilePath = $Path + '\' + $Mailbox.UserPrincipalName + '.pst'
			New-MailboxExportRequest -Mailbox $Mailbox -ContentFilter "Received -ge '$ExportDate'" -FilePath $FilePath -BatchName $BatchName -Name $BatchName -BadItemLimit 50 -AcceptLargeDataLoss -Confirm:$Confirm
			$i++
			Write-Progress -Activity "Creating export request for" -status "$Mailbox" -percentComplete ($i / $Mailboxes.Count*100)
		}
		Write-Host "Run Get-MailboxExportRequest | Get-MailboxExportRequestStatistics to get the export status." -ForegroundColor Yellow
		Write-Host "Run Get-MailboxExportRequest -Status Completed | Remove-MailboxExportRequest to remove completed requests." -ForegroundColor Yellow
}

# Initialize

Import-Module ActiveDirectory
if ($GetExchangeInfo) { GetExchangeInfo }
if ($ExportMailboxList) { ExportMailboxList }
if ($ExportDistributionGroups) { ExportDistributionGroups }
if ($InitialPSTExport) { InitialPSTExport }
elseif ($FinalPSTExport) { FinalPSTExport }