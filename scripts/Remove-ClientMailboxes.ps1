<# 
.SYNOPSIS 
Remove-ClientMailboxes.ps1 - Archives, sets forwarding and removes inactive client mailboxes.

.NOTES
Requires the Active Directory PowerShell module.

CSVProperties:
EmailAddress,ForwardTo,ExportPST

Written by Ulisses Righi
ulisses@ulisoft.com.br
Version 1.9.2
6/20/2018

#>

[CmdletBinding()]
param (
    [Parameter(Mandatory=$true)]
    [string]$CSVPath,

    [Parameter(Mandatory=$false)]
    [string]$PSTPath,

    [Parameter(Mandatory=$false)]
    [switch]$Unattended = $false
    )

$ScriptLogFile = "MailboxCleanup.log"
    Function WriteToLog 
{
   Param ([string]$Details)
   $LogString = (Get-Date -format "dd/MM/yyyy HH:mm") + " - " + $Details
   Add-Content $ScriptLogFile -value $LogString
}

$MailboxList = Import-CSV $CSVPath
if ($PSTPath -eq "" -or $PSTPath -eq $null) { $PSTPath = "\\exmbx01\f$\PSTExports\" }

$i = 0
foreach ($Mailbox in $MailboxList)
{
    $i++
    Write-Progress -Id 1 -Activity "Processing $($Mailbox.EmailAddress) - Archive: $($Mailbox.ExportPST) - Forward to: $($Mailbox.ForwardTo)" -Status "[$i/$($MailboxList.Count)]" -PercentComplete ($i/($MailboxList.Count)*100)
    WriteToLog "Processing $($Mailbox.EmailAddress)"

    if ([bool]::Parse($Mailbox.ExportPST))
    {
        New-MailboxExportRequest -Mailbox $Mailbox.EmailAddress -FilePath "$PSTPath\$($Mailbox.EmailAddress).pst" -BatchName "Pre-removal export" -Name "Pre-removal export" -LargeItemLimit 1000 -BadItemLimit 1000 -AcceptLargeDataLoss -Confirm:(!$Unattended)
        
        while ((Get-MailboxExportRequest -Mailbox $Mailbox.EmailAddress).Status -ne "Completed")
        {
            Write-Host "Waiting for $($Mailbox.EmailAddress) export to complete..." -ForegroundColor Yellow
            $Request = Get-MailboxExportRequest -Mailbox $Mailbox.EmailAddress
            $Request | Get-MailboxExportRequestStatistics | ft Name,StatusDetail,SourceAlias,PercentComplete
            if ($Request.Status -ne "Completed")
            {
                $Request | Resume-MailboxExportRequest -ErrorAction SilentlyContinue
                Start-Sleep -Seconds 15
            }
        }

        Get-MailboxStatistics $Mailbox.EmailAddress | ft DisplayName,ItemCount,TotalItemSize

        Get-ChildItem "$PSTPath\$($Mailbox.EmailAddress).pst" | Select-Object FullName,@{l="Length (MB)";e={("{0:N3}" -f $_.Length/1MB)}}

        Write-Host "$($Mailbox.EmailAddress) exported to $PSTPath\$($Mailbox.EmailAddress).pst`r`n"
        WriteToLog "$($Mailbox.EmailAddress) exported to $PSTPath\$($Mailbox.EmailAddress).pst"
    }

    $EmailAddresses = (Get-Mailbox $Mailbox.EmailAddress).EmailAddresses
    $Forwarders = (Get-Mailbox $Mailbox.EmailAddress).ForwardingAddress
    $SMTPForwarders = (Get-Mailbox $Mailbox.EmailAddress).ForwardingSmtpAddress

    Get-MailboxExportRequest -Mailbox $Mailbox.EmailAddress | Remove-MailboxExportRequest -Confirm:$false
    Remove-Mailbox $Mailbox.EmailAddress -Confirm:(!$Unattended)
    WriteToLog "Removed mailbox $($Mailbox.EmailAddress)"

    if (($Mailbox.ForwardTo -ne "") -and ($Mailbox.ForwardTo -ne $null))
    {
        foreach ($EmailAddress in $EmailAddresses)
        {
            if ($EmailAddress -notlike "*x500*")
            {
                $NewEmail = $EmailAddress.ToString().ToLower()
                Set-Mailbox $Mailbox.ForwardTo -EmailAddresses @{Add="$NewEmail"} -Confirm:(!$Unattended)
                WriteToLog "Added alias $NewEmail to $($Mailbox.ForwardTo)"
            }
        }
    }

    # Handles already existing forwarders
    if ($Forwarders -ne $null)
    {
        WriteToLog "Existing forwarders for $($Mailbox.EmailAddress) found"
        foreach ($EmailAddress in $EmailAddresses)
        {
            if ($EmailAddress -notlike "*x500*")
            {
                $NewEmail = $EmailAddress.ToString().ToLower()
                if ((Get-Recipient $Forwarders).RecipientType -eq "MailUniversalDistributionGroup")
                {
                    Set-DistributionGroup $Forwarders -EmailAddresses @{Add="$NewEmail"} -Confirm:(!$Unattended)
                }
                else
                {
                    Set-Mailbox $Forwarders -EmailAddresses @{Add="$NewEmail"} -Confirm:(!$Unattended)
                }
                WriteToLog "Added alias $NewEmail to $Forwarders"
            }
        }
    }

    # Handles already existing SMTP forwarders
    if ($SMTPForwarders -ne $null)
    {
        Write-Host "SMTP forwarders have been found for $($Mailbox.EmailAddress). Please review the log file and add those manually if necessary."
        foreach ($EmailAddress in $EmailAddresses)
        {
            if ($EmailAddress -notlike "*x500*")
            {
                $NewEmail = $EmailAddress.ToString().ToLower()
                try
                {
                    if ((Get-Recipient $SMTPForwarders).RecipientType -eq "MailUniversalDistributionGroup")
                    {
                        Set-DistributionGroup $SMTPForwarders -EmailAddresses @{Add="$NewEmail"} -Confirm:(!$Unattended)
                    }
                    else
                    {
                        Set-Mailbox $SMTPForwarders -EmailAddresses @{Add="$NewEmail"} -Confirm:(!$Unattended)
                    }
                    WriteToLog "Added alias $NewEmail to $SMTPForwarders"
                }
                catch
                {
                    WriteToLog "Recipient $SMTPForwarders was not found on Exchange. Create a mail contact if necessary. Alias not created: $EmailAddress"
                }
            }
        }
    }
}