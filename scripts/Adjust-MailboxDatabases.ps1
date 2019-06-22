<#
.SYNOPSIS
Moves mailboxes to the appropriate database according to client.

.AUTHOR
Ulisses Righi
ulisses@ulisoft.com.br
#>

param (
	[Parameter(Mandatory=$false)]
	[switch]$MailboxReportPerAcceptedDomain,
	[Parameter(Mandatory=$false)]
	[switch]$MoveMailboxes,
	[Parameter(Mandatory=$false)]
	[switch]$Confirm
)

#$Domains = Get-AcceptedDomain
#$MailboxList = Get-Mailbox
$DBToDomainsCSV = Import-CSV DBToDomains.csv
$DBToDomains=@{}
foreach($row in $DBToDomainsCSV)
{
    $DBToDomains[$row.Domain]=$row.DatabaseName
}

Function MailboxReportPerAcceptedDomain {
	foreach ($Dom in $Domains) 
	{
		$Name = $Dom.DomainName; 
		Write-Host "$name`r`n" -ForegroundColor Yellow;
		$MailboxList | Where {$_.PrimarySMTPAddress -like "*$Name*"} | ft Alias,PrimarySMTPAddress,Database
	}
}

Function MoveMailboxes {
	foreach ($Dom in $Domains)
	{
		$Name = $Dom.DomainName.Domain; 
		$Database = $DBToDomains[$Dom.DomainName]
		Write-Host "Domain: $Name Database: $Database" -ForegroundColor Yellow
		$Mailboxes = $MailboxList | Where {($_.PrimarySMTPAddress -like "*$Name*") -and ($_.Database -ne $Database)}
		$mbx | New-MoveRequest -Identity $Mailbox -TargetDatabase $Database -BatchName "Client adjustment to DB" -Confirm:$Confirm
	}
}

#Initialize
if ($MailboxReportPerAcceptedDomain) { MailboxReportPerAcceptedDomain }
if ($MoveMailboxes) { MoveMailboxes }