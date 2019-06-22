<# 
.SYNOPSIS 
Create-MailboxReport.ps1 - Generates a mailbox report.

.NOTES
Requires the Active Directory PowerShell module.

Written by Ulisses Righi
ulisses@ulisoft.com.br
Version 1.9.2
6/20/2018

#>

Add-PSSnapin Microsoft.Exchange.Management.PowerShell.E2010
Import-Module ActiveDirectory

if (!(Test-Path "C:\ReportScript")) { New-Item "C:\ReportScript" -Type Directory }

$ReportPath = "C:\ReportScript\MailboxBillingReport_$(Get-Date -Format MMM_yyyy).csv"

Get-Mailbox | Where-Object {($_.RecipientTypeDetails -eq "UserMailbox") -or ($_.RecipientTypeDetails -eq "RoomMailbox")} | `
    Select-Object Alias,DisplayName,Database,@{Name='EmailAddresses';Expression={[string]::join(";", ($_.EmailAddresses))}}, `
    @{Name='GivenName';Expression={(Get-AdUser $_.SamAccountName).GivenName}},@{Name='Surname';Expression={(Get-AdUser $_.Alias).Surname}},`
    @{Name='ItemCount';Expression={(Get-Mailbox $_ | Get-MailboxStatistics).ItemCount}},@{Name='TotalItemSize';Expression={(Get-Mailbox $_ | Get-MailboxStatistics).TotalItemSize}},`
    PrimarySMTPAddress,RecipientTypeDetails,@{Name='Enabled';Expression={(Get-AdUser $_.SamAccountName -Property Enabled).Enabled}} | Export-Csv $ReportPath -NoTypeInformation