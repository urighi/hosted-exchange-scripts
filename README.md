# hosted-exchange-scripts
Helper scripts to manage a multi-tenant Exchange environment.

### What are the Hosted Exchange Scripts?
The Exchange Migration Scripts are a set of scripts created to aid migration and management of multi-tenant Exchange 2013 (and higher) environments, supporting DAG.

- **Adjust-MailboxDatabases.ps1**
    - Moves mailboxes to their appropriate databases, based on domain and client name.
- **Create-MailboxReport.ps1**
    - Generates a mailbox report including mailbox size and item count. Useful for billing purposes.
- **Export-Mailboxes.ps1**
    - Generates a CSV list of accepted domains, mailboxes, permissions, delegates, forwarding addresses and distribution groups. Creates export requests for Exchange mailboxes, according to a set date. Supported on Exchange 2010 SP3 sources and above.
- **Import-Mailboxes.ps1**
    - Imports user accounts, distribution groups, permissions, delegates and forwarding addresses from a CSV file, enables the mailboxes and user accounts, and imports PST files to mailboxes. Also adds members to distribution groups. Generates passwords and exports them to a CSV file.
- **New-Password.ps1**
    - Function used to generate new passwords for accounts.
- **Prepare-HostedDomain.ps1**
    - Adds UPNs to Active Directory and creates an OU for the hosted client. In Exchange, creates accepted domains, e-mail address policies, mailbox databases and DAG copies, address lists, GALs and OABs. Updates e-mail address policies, accepted domains and e-mail addresses on the final phase (post-cutover).
- **Remove-ClientConfig.ps1**
    - Removes all client configuration created with *Prepare-HostedDomain* from the organization.
- **Remove-ClientMailboxes.ps1**
    - Archives, sets forwarding addresses and removes client mailboxes.

All scripts support the -Confirm switch and will ask for confirmation on every action if this parameter is not set to $false when the script is called. They do not support -WhatIf.

## Migration Scripts

### Preparation

Make sure you can access Active Directory Module in PowerShell. To do this:
1. Open the Exchange Management Shell.
2. Run the following command:
    ```
    Import-Module ActiveDirectory
    ```
3. If you get an error message, the Active Directory Module needs to be installed. In Windows 2008R2 and newer versions, the module can be installed by running the following commands:
    ```
    Import-Module ServerManager
    Add-WindowsFeature RSAT-AD-PowerShell,RSAT-AD-AdminCenter
    ```

    Alternatively, the module can be installed from Server Manager > Add Features > Remote Server Administration Tools > AD DS and AD LDS Tools > Active Directory Module for Windows PowerShell.
    
Make sure you have permissions to export mailboxes in Exchange before running the script. To do this:
1. Open the Exchange Management Shell.
2. Run the following command:
    ```
    New-ManagementRoleAssignment -Name "Mailbox Import Export" -SecurityGroup "Organization Management" -Role "Mailbox Import Export"
    ```
    This will add the Mailbox Import Export role to the Organization Management group, which EAdmin and Administrator should be a member of.
    You can also assign the role to an individual user by running
    ```
    New-ManagementRoleAssignment –Role "Mailbox Import Export" –User Administrator
    ```
3. Restart the Exchange Management Shell.

### Exporting Mailboxes and Exchange Information
#### Gathering mailbox information
1. Open the Exchange Management Shell.
2. Navigate to the scripts folder. Run the following command:
    ```
    .\Export-Mailboxes.ps1 -Path <UNC Path> -GetExchangeInfo
    ```
3. Take note of the Accepted Domains and e-mail address policies. They will be needed for the Exchange preparation phase.
4. Export the mailbox and distribution group list by running the following command:

    ```
    .\Export-Mailboxes.ps1 -Path <UNC Path> -ExportMailboxList -ExportDistributionGroups
    ```

#### Exporting mailboxes to PST
1. To perform an initial mailbox export, run the following command:
    ```
    .\Export-Mailboxes.ps1 -Path <UNC Path> -InitialPSTExport
    ```
2. To perform a final mailbox export, run the following command:
    ```
    .\Export-Mailboxes.ps1 -Path <UNC Path> -FinalPSTExport
    ```
    You can use the -Date parameter to set a date, or let the script use the date.txt file as long as it is stored on the same folder as the PST files.
3. After all export requests are completed, you can remove them by running the following command:
    ```
    Get-MailboxExportRequest -Status Completed | Remove-MailboxExportRequest
    ```
    This will not delete the PST files.
    
#### Preparing Active Directory and Exchange

To prepare Active Directory and Exchange, the Prepare-HostedDomain.ps1 script is used. You'll need the following information, which you gathered from the Export-Mailboxes.ps1 script:

- *ForestUPNs* - List of UPNs (e.g. "righi.com", "righi.it", "righi.br"
- *ClientName* - Client Name (e.g. "Righi")

You can also set the following values when running Prepare-HostedDomain.ps1:

- **MailboxServer** (Default: EXMBX01)
- **EdbFilePath** (Default: "E:\Program Files\Microsoft\Exchange Server\V15\Mailbox\HOSTED - $ClientName\HOSTED - $ClientName.edb")
- **LogFolderPath** (Default: "E:\Program Files\Microsoft\Exchange Server\V15\Mailbox\HOSTED - $ClientName\")
- **ForestName** (Default: hosting.email)
- **OUPath** (Default: "OU=Hosted Organizations,DC=hosting,DC=email")
- **Confirm** (Default: $true)
- **UseMigrationSuffix** (Default: $false)

The **PrepareForest** parameter will allow for the creation of UPNs and the Organizational Unit in Active Directory.

The **PrepareExchange** parameter will allow for the creation of accepted domains (same as the UPN list), address policies, mailbox databases and DAG copies, address lists, GALs and OABs.

The **UseMigrationSuffix** parameter will add ".migration" to the end of accepted domains, but not alter the UserPrincipalName parameter.

To prepare both Exchange and AD, run:
```
.\Prepare-HostedDomain.ps1 -ForestUPNs "upn1.com,upn2.com,upn3.com" -ClientName "Client" -PrepareForest -PrepareExchange -UseMigrationSuffix
```

**NOTE**: Exchange will create an authoritative domain, which means that e-mails to these domains will be routed internally. If the *UseMigrationSuffix* parameter wasn't used, and if necessary, change the domain type to *InternalRelay* and add a explicit send connector to deliver messages to the external domain. Remove this rule and change the domain back to *Authoritative* when the migration is completed.

#### Importing Mailboxes

To import mailboxes to Exchange, the Import-Mailboxes.ps1 script is used. You'll need the following information:

- **PSTPath** - UNC path to the PST files (e.g. "\\mbx1\c$\exchangePSTs")
- **MailboxListPath** - UNC path to the mailbox list (e.g. "\\mbx1\c$\exchange\MailboxList.csv"). If no path is set, the script will use $PSTPath\MailboxList.csv as the path.
- **ClientName** - Client Name (e.g. "Righi")

You can also set the following values when running Import-Mailboxes.ps1:

- **EnableADAccount**  - Enables the AD account after creation. Can be also used in the final import (Default: $false)
- **PasswordLength** - Password length used for the new accounts (Default: 8)
- **MailboxDatabase** - Mailbox database name, in case it is different from the format "HOSTED - $ClientName"
- **OUPath** (Default: "OU=Hosted Organizations,DC=hosting,DC=email")
- **UseMigrationSuffix** (Default: $false)
- **Confirm** (Default: $true)

The **InitialPSTImport** parameter will tell the script to perform an initial import, which includes:

- User account and password creation
- Mailbox enablement
- E-mail address configuration for the mailbox
- Creating a PST import request

Passwords are exported to *$PSTPath\Passwords_ClientName.csv*.

User accounts are created using the Alias in the *MailboxList.csv* file; if an account with the same SamAccountName exists, it will append "_ClientName" to the user account name.

The **FinalPSTImport** parameter will tell the script to perform a final import, which is creating a PST import request to the already existing mailboxes. Source items will be preserved.

The **ImportDistributionGroups** parameter will tell the script to import distribution groups. When setting this parameter, the UNC path to the distribution groups list can be set by using the **DistributionGroupListPath** parameter. If no path is set, the script will use *$PSTPath\DistributionGroupList.csv* as the path.

The **ImportPermissions** parameter will tell the script to import and apply permissions to the mailboxes. When setting this parameter, the UNC path to the permission list can be set by using the **PermissionsListPath** parameter. If no path is set, the script will use *$PSTPath\PermissionList.csv* as the path.

The **ImportForwardingAddresses** parameter will tell the script to import and apply forwarding addresses to the mailboxes. When setting this parameter, the UNC path to the forwarding list can be set by using the **ForwardingAddressesListPath** parameter. If no path is set, the script will use *$PSTPath\ForwardingList.csv* as the path.

The **ImportDelegates** parameter will tell the script to add permissions to mailbox folders. When setting this parameter, the UNC path to the permission list can be set by using the **DelegateListPath** parameter. If no path is set, the script will use *$PSTPath\DelegateList.csv* as the path. The account listed in the user column must be enabled in Active Directory.

The **UseMigrationSuffix** parameter will add ".migration" to the end of e-mail addresses, but not alter the UserPrincipalName parameter.

Any errors will be logged to the *MailboxImport.log* file, on the same directory as the script.

1. To perform an initial mailbox and distribution group migration, run
    ```
    .\MailboxImport.ps1 -InitialPSTImport -ClientName "Client" -PSTPath "\\EXMBX02\E$\PSTs" -MailboxListPath "\\EXMBX02\E$\PSTs\MailboxList.csv" -DistributionGroupListPath "\\exmbx02\e$\PSTs\DistributionGroupList.csv" -ImportDistributionGroups -UseMigrationSuffix
    ```
2. To perform a final mailbox migration and enable the user account, run
    ```
    .\MailboxImport.ps1 -FinalPSTImport -ClientName "Client" -PSTPath "\\EXMBX02\E$\PSTs" -MailboxListPath "\\EXMBX02\E$\PSTs\MailboxList.csv" -EnableADAccount
    ```
3. To import delegates only, run
    ```
    .\MailboxImport.ps1 -ClientName "Client" -DelegateListPath "\\EXMBX02\E$\PSTs\DelegateList.csv" -ImportDelegates
    ```

License
----

GNU GPLv3
