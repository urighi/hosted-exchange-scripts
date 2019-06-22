<#
.SYNOPSIS 
Prepare-HostedDomain.ps1 - Prepares the Active Directory Forest and Exchange Server for the new hosted domain.

.DESCRIPTION
Adds UPNs to Active Directory and creates an OU for the hosted client. In Exchange,
creates accepted domains, e-mail address policies, mailbox databases and DAG copies,
address lists, GALs and OABs.

.OUTPUTS
Results are output to screen. Errors are output to screen and the PrepareHostedDomain.log
file.

.PARAMETER PrepareForest
Indicates if the UPNs and OU should be configured in Active Directory.

.PARAMETER PrepareExchange
Indicates if accepted domains, address policies, mailbox databases, DAG copies,
address lists, GALs and OABs should be created.


.PARAMETER FinalizeMigration
Updates the accepted domain, email address policies and recipient and distribution 
groups e-mail addresses, removing the .migration suffix.

.PARAMETER ForestUPNs
List of UPNs to be added to Active Directory and added to Exhange as accepted
e-mail addresses, separated by commas.

.PARAMETER ClientName
Client name that will be used for naming in Active Directory and Exchange.

.PARAMETER MailboxServer
Mailbox server that will hold the primary database copy.

.PARAMETER Drive
Path to the mailbox database file.

.PARAMETER LogFolderPath
Path to the database log folder.

.PARAMETER ForestName
Forest name on the destination domain. 

.PARAMETER OUPath
Distinguished name for the Organizational Unit that holds all clients.

.PARAMETER UseMigrationSuffix
Adds the .migration suffix to domain names in Exchange.

.PARAMETER SkipMailboxDatabaseCreation
Skips creating a mailbox database.

.PARAMETER CreatePublicFolder
Creates a public folder for the company.

.EXAMPLE
.\Prepare-HostedDomain.ps1 -PrepareForest -PrepareExchange -ClientName "Righi.com" -ForestUPNs "ulissesrighi.com,righi.it,ulissesrighi.com.br" -OUPath "OU=Hosted Exchange,DC=hosting,DC=email" -Confirm:$false
Prepares both Active Directory and Exchange for the client "Righi.com", creating
an OU named "Righi.com" under "OU=Hosted Exchange,DC=hosting,DC=email"", creates 
an address policy for all recipients in that OU, adding "ulissesrighi.com,righi.it,ulissesrighi.com.br" 
to the list of UPNs in AD and accepted domains in Exchange, creates a mailbox database
called "HOSTED - Righi.com", creates mailbox database copies, creates address list
for all recipients in that OU.

.EXAMPLE
.\Prepare-HostedDomain.ps1 -PrepareExchange -ClientName "Righi.com" -ForestUPNs "righi.it" -MailboxServer "EXMBX02" -Drive F

.NOTES
Requires the Active Directory PowerShell module.

Written by Ulisses Righi
ulisses@ulisoft.com.br
Version 1.9.2
6/20/2018

#>

[CmdletBinding()]
param(
	[Parameter(Mandatory=$false)]
	[switch]$PrepareForest,
	
	[Parameter(Mandatory=$false)]
	[switch]$PrepareExchange,
	
	[Parameter(Mandatory=$false)]
	[switch]$FinalizeMigration,
	
	[Parameter(Mandatory=$true)] [ValidateNotNullOrEmpty()]
	[string]$ForestUPNs,
	
	[Parameter(Mandatory=$true)] [ValidateNotNullOrEmpty()]
	[string]$ClientName,
	
	[Parameter(Mandatory=$false)] [ValidateNotNullOrEmpty()]
	[string]$MailboxServer = "EXMBX01",
	
	[Parameter(Mandatory=$false)]
	[string]$Drive = "E",
	
	[Parameter(Mandatory=$false)]
	[string]$LogFolderPath,
	
	[Parameter(Mandatory=$false)]
	[string]$ForestName = "hosting.email",
	
	[Parameter(Mandatory=$false)]
	[string]$OUPath = "OU=Hosted Organizations,DC=hosting,DC=email",
	
	[Parameter(Mandatory=$false)]
	[switch]$UseMigrationSuffix,
	
	[Parameter(Mandatory=$false)]
	[switch]$SkipMailboxDatabaseCreation,

	[Parameter(Mandatory=$false)]
	[switch]$CreatePublicFolder,
	
	[Parameter(Mandatory=$false)]
	[bool]$Confirm = $true,

	[Parameter(Mandatory=$false)]
	[string[]]$Databases
	)


# Script Configuration Variables

$RecipientContainer = "OU=" + $ClientName + "," + $OUPath
$ForestUPNsArray = $ForestUPNs -Split ','
$MailboxDatabase = "HOSTED - " + $ClientName

$EdbFilePath = "$($Drive):\Program Files\Microsoft\Exchange Server\V15\Mailbox\" + $MailboxDatabase + "\" + $MailboxDatabase + ".edb"
$LogFolderPath = "$($Drive):\Program Files\Microsoft\Exchange Server\V15\Mailbox\" + $MailboxDatabase
$ScriptLogFile = "PrepareHostedDomain.log"
$ErrorActionPreference = "Stop"

Function WriteToLog 
{
   Param ([string]$Details)
   $LogString = (Get-Date -format G) + " " + $Details
   Add-Content $ScriptLogFile -value $LogString
}

Function PrepareForest {
	# Adds UPNs
	try
	{
		foreach ($UPN in $ForestUPNsArray)
		{
			Write-Host "`nAdding $UPN to the list of UPN suffixes for $ForestName" -ForegroundColor Yellow
			Set-AdForest -Identity $ForestName -UPNSuffixes @{Add="$UPN"} -Confirm:$Confirm
		}
		# Creates Organizational Unit
		Write-Host "`nCreating Organizational Unit $ClientName at $OUPath"
		New-AdOrganizationalUnit -Name $ClientName -Path $OUPath -Confirm:$Confirm
	}
	catch
	{
		WriteToLog $_.Exception.Message
		Write-Host "Forest preparation failed. Please refer to $ScriptLogFile for details." -ForegroundColor Red
	}
}

Function FinalizeMigration {

	# Updates accepted domain
	
	try
	{
		foreach ($UPN in $ForestUPNsArray)
		{
			$DomainName = $UPN
			$Name = $ClientName + " " + $DomainName
			Remove-AcceptedDomain -Identity "$Name.migration" -Confirm:$Confirm
			New-AcceptedDomain -Name $Name -DomainName $DomainName -DomainType Authoritative -Confirm:$Confirm
		}

	}
	catch
	{
		WriteToLog $_.Exception.Message
		Write-Host "Accepted domain creation or removal failed. Please refer to $ScriptLogFile for details." -ForegroundColor Red
	}
	
	# Creates e-mail address policies
	try
	{
		$primary = $true
		[Microsoft.Exchange.Data.ProxyAddressTemplateCollection]$SMTPAddresses = @()
		
		foreach ($UPN in $ForestUPNsArray)
		{
			if ($primary) { $Address = "SMTP:%m"; $primary = $false }
			else { $Address = "smtp:%m" }
			$Address += "@" + $UPN
			if ($UseMigrationSuffix) { $Address += ".migration" }
			$SMTPAddresses += $Address
		}
		
		New-EmailAddressPolicy -RecipientContainer $RecipientContainer -IncludedRecipients AllRecipients -EnabledEmailAddressTemplate $SMTPAddresses -Name $ClientName -Confirm:$Confirm | Update-EmailAddressPolicy
	}
	catch
	{
		WriteToLog $_.Exception.Message
		Write-Host "E-mail address policy creation failed. Please refer to $ScriptLogFile for details." -ForegroundColor Red
	}
	
	# Fixes recipient email address
	
	try
	{
		foreach ($Database in $Databases)
		{
			$Mailboxes += Get-Mailbox -Database $Database
		}
		
		#$Mailboxes = Get-Mailbox -Database "HOSTED - $ClientName"
		
		foreach ($mbx in $mailboxes)
		{
			$NewAddresses = @()
			foreach ($address in $mbx.EmailAddresses)
		{
			if ($address.Prefix -ne "X500")
			{
				$NewAddresses += $address.ProxyAddressString.Replace(".migration","")
			}
			else 
			{
			$NewAddresses += $address.ProxyAddressString
			}
		}
			Write-Host $NewAddresses
			Set-Mailbox $mbx -EmailAddresses $NewAddresses -Confirm
		}
	}
	catch
	{
		WriteToLog $_.Exception.Message
		Write-Host "E-mail address creation or removal failed. Please refer to $ScriptLogFile for details." -ForegroundColor Red
	}
	
	try
	{
		foreach ($UPN in $ForestUPNsArray)
		{
			$dgs = Get-DistributionGroup | ? { $_.PrimarySMTPAddress -like "*$UPN*"  } 
			foreach ($dg in $dgs)
			{
				$NewAddresses = @()
				foreach ($address in $dg.EmailAddresses)
				{
					if ($address.Prefix -ne "X500")
					{
						$NewAddresses += $address.ProxyAddressString.Replace(".migration","")
					}
					else 
					{
						$NewAddresses += $address.ProxyAddressString
					}
				}
				Set-DistributionGroup $dg -EmailAddresses $NewAddresses
			}
		}

	}
	catch
	{
		WriteToLog $_.Exception.Message
		Write-Host "E-mail address creation or removal failed. Please refer to $ScriptLogFile for details." -ForegroundColor Red
	}
}

Function PrepareExchange {
	
	# Adds accepted domain
	try
	{
		foreach ($UPN in $ForestUPNsArray)
		{
			$DomainName = $UPN
			if ($UseMigrationSuffix) { $DomainName += ".migration" }
			$Name = $ClientName + " " + $DomainName
			New-AcceptedDomain -Name $Name -DomainName $DomainName -DomainType Authoritative -Confirm:$Confirm
		}
	}
	catch
	{
		WriteToLog $_.Exception.Message
		Write-Host "Accepted Domain creation failed. Please refer to $ScriptLogFile for details." -ForegroundColor Red
	}
	
	# Adds mailbox database
	if (!$SkipMailboxDatabaseCreation)
	{
		foreach ($MailboxDatabase in $Databases)	
		{
			try
			{
				$Database = New-MailboxDatabase -Name $MailboxDatabase -EdbFilePath $EdbFilePath -LogFolderPath $LogFolderPath -Server $MailboxServer -Confirm:$Confirm
				Set-MailboxDatabase $MailboxDatabasen
			}
			catch
			{
				WriteToLog $_.Exception.Message
				Write-Host "Mailbox database creation failed. Please refer to $ScriptLogFile for details." -ForegroundColor Red
			}
			
			try
			{
				Write-Host "Restarting the Microsoft Exchange Information Store Service. The script will run `nRedistributeActiveDatabases.ps1 after all database copies are created, `nto rebalance the DAG." -ForegroundColor Yellow
				Get-Service MSExchangeIS -ComputerName $MailboxServer | Restart-Service -Confirm:$Confirm
				Mount-Database -Identity $MailboxDatabase -Confirm:$Confirm
			}
			catch
			{
				WriteToLog $_.Exception.Message
				Write-Host "Database mount operation failed. Please refer to $ScriptLogFile for details." -ForegroundColor Red
			}
			
			# Adds mailbox database copy
			try
			{
				$MailboxCopyServers = Get-MailboxServer
				$SourceServer = Get-MailboxServer $MailboxServer
				foreach ($MailboxCopyServer in $MailboxCopyServers)
				{	if (($MailboxCopyServer.Name -ne $MailboxServer) -and ($MailboxCopyServer.DatabaseAvailabilityGroup -eq $SourceServer.DatabaseAvailabilityGroup))
					{
						Add-MailboxDatabaseCopy -Identity $MailboxDatabase -MailboxServer $MailboxCopyServer -Confirm:$Confirm
						Get-Service MSExchangeIS -ComputerName $MailboxCopyServer | Restart-Service -Confirm:$Confirm
					}
				}
				
				Write-Host "Rebalancing DBs..." -ForegroundColor Yellow
				& "C:\Program Files\Microsoft\Exchange Server\V15\Scripts\RedistributeActiveDatabases.ps1" -DagName EXDAG01 -BalanceDbsByActivationPreference -Confirm:$Confirm
			}
			catch
			{
				WriteToLog $_.Exception.Message
				Write-Host "Mailbox database copy creation failed. Please refer to $ScriptLogFile for details." -ForegroundColor Red
			}
			
			try
			{
				Set-MailboxDatabase $MailboxDatabase -CircularLoggingEnabled:$true
			}
			catch
			{
				WriteToLog $_.Exception.Message
				Write-Host "Mailbox database log configuration failed. Please refer to $ScriptLogFile for details." -ForegroundColor Red
			}
		}
	}

	# Create Address lists
	try
	{
		$GALName = $ClientName + "_GAL"
	
		New-AddressList -Name "$ClientName - All Users" -RecipientFilter {(RecipientType -eq "UserMailbox")} -RecipientContainer $RecipientContainer -Confirm:$Confirm
		New-AddressList -Name "$ClientName - All Groups" -RecipientFilter {((RecipientType -eq "MailUniversalDistributionGroup") -or (RecipientType -eq "DynamicDistributionGroup"))} -RecipientContainer $RecipientContainer -Confirm:$Confirm
		New-AddressList -Name "$ClientName - All Contacts" -RecipientFilter {((RecipientType -eq "MailContact") -or (RecipientType -eq "MailUser"))} -RecipientContainer $RecipientContainer -Confirm:$Confirm
		New-AddressList -Name "$ClientName - All Rooms" -RecipientFilter {(Alias -ne $null) -and (RecipientDisplayType -eq "ConferenceRoomMailbox") -or (RecipientDisplayType -eq 'SyncedConferenceRoomMailbox')} -RecipientContainer $RecipientContainer -Confirm:$Confirm
		
		New-GlobalAddressList -Name $GALName -RecipientFilter {(Alias -ne $null)} -RecipientContainer $RecipientContainer -Confirm:$Confirm
		
		# Create OAB
		New-OfflineAddressBook -Name "$ClientName - Offline Address Book" -AddressLists $GALName -GlobalWebDistributionEnabled $True -Confirm:$Confirm
		# Create AB policy
		New-AddressBookPolicy -Name $ClientName -GlobalAddressList "\$GALName" -OfflineAddressBook "\$ClientName - Offline Address Book" -RoomList "\$ClientName - All Rooms" -AddressLists "\$ClientName - All Contacts","\$ClientName - All Groups","\$ClientName - All Users" -Confirm:$Confirm
	}
	catch
	{
		WriteToLog $_.Exception.Message
		Write-Host "Address list creation failed. Please refer to $ScriptLogFile for details." -ForegroundColor Red
	}
}

function CreatePublicFolder
{
	try
	{
		if ($CreatePublicFolder)
		{
			New-PublicFolder -Name "$ClientName Public Folders" -Confirm:$Confirm
			Remove-PublicFolderClientPermission "\$ClientName Public Folders" -User Default
		}
	}
	catch
	{
		WriteToLog $_.Exception.Message
		Write-Host "Public folder database creation failed. Please refer to $ScriptLogFile for details." -ForegroundColor Red
	}
}

# Initialize

Import-Module ActiveDirectory
if ($PrepareForest) { PrepareForest }
if ($PrepareExchange) { PrepareExchange }
if ($FinalizeMigration) { FinalizeMigration }
if ($CreatePublicFolder) { CreatePublicFolder }