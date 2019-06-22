<#
.SYNOPSIS 
Remove-ClientConfig.ps1 - Removes hosted domain configuration from Active Directory and Exchange.

.NOTES
Requires the Active Directory PowerShell module.

Written by Ulisses Righi
ulisses@ulisoft.com.br
Version 1.9.2
6/20/2018

#>


[CmdletBinding()]
param(
	
	[Parameter(Mandatory=$true)] [ValidateNotNullOrEmpty()]
	[string]$ForestUPNs,
	
	[Parameter(Mandatory=$true)] [ValidateNotNullOrEmpty()]
	[string]$ClientName,
	
	[Parameter(Mandatory=$false)]
	[string]$ForestName = "hosting.email",
	
	[Parameter(Mandatory=$false)]
	[string]$OUPath = "OU=Hosted Organizations,DC=hosting,DC=email",
	
    [Parameter(Mandatory=$false)]
	[switch]$RemoveClientMailboxes,
	
	[Parameter(Mandatory=$false)]
	[bool]$Confirm = $true

    )
    
$ScriptLogFile = "RemoveClientConfig.log"
$ForestUPNsArray = $ForestUPNs -Split ','

Function WriteToLog 
{
   Param ([string]$Details)
   $LogString = (Get-Date -format G) + " " + $Details
   Add-Content $ScriptLogFile -value $LogString
}

Function RemoveExchangeConfiguration
{
    Get-DistributionGroup -OrganizationalUnit "hosting.email/Hosted Organizations/$ClientName" | Remove-DistributionGroup

    Get-AddressBookPolicy -Identity "*$ClientName*" | Remove-AddressBookPolicy -Confirm:$Confirm
    Get-OfflineAddressBook -Identity "*$ClientName*" | Remove-OfflineAddressBook -Confirm:$Confirm
    Get-GlobalAddressList -Identity "*$ClientName*" | Remove-GlobalAddressList -Confirm:$Confirm
    Get-AddressList -Identity "*$ClientName*" | Remove-AddressList -Confirm:$Confirm
    Get-EmailAddressPolicy -Identity "*$ClientName*" | Remove-EmailAddressPolicy -Confirm:$Confirm
    

    foreach ($UPN in $ForestUPNsArray)
    {
        Get-AcceptedDomain | ? { $_.DomainName -eq $UPN } | Remove-AcceptedDomain -Confirm:$Confirm
    }

}

Function RemoveADConfiguration
{
    foreach ($UPN in $ForestUPNsArray)
		{
			Write-Host "`nRemoving $UPN from the list of UPN suffixes for $ForestName" -ForegroundColor Yellow
			Set-AdForest -Identity $ForestName -UPNSuffixes @{Remove="$UPN"} -Confirm:$Confirm
        }

    Set-ADOrganizationalUnit "OU=$ClientName,$OUPath" -ProtectedFromAccidentalDeletion $false -Confirm:$Confirm
    Start-Sleep -Seconds 10
    Remove-ADOrganizationalUnit "OU=$ClientName,$OUPath" -Recursive -Confirm:$Confirm
    
}

if ((Get-ADUser -SearchBase "OU=$ClientName,$OUPath" -Filter *) -ne $null)
{
    Write-Error "User accounts or mailboxes still exist for $ClientName. Please use the RemoveClientMailboxes.ps1 script to archive and remove user mailboxes and accounts."
}

else
{
    Write-Warning "Please remove MX records in DNS and the organization in Reflexion."
    Write-Warning "Databases and public folders must be removed manually."
    RemoveExchangeConfiguration
    RemoveADConfiguration    
}

