# Get-RecipientPermissions.ps1

The Get- RecipientPermissions.ps1 is a PowerShell script that will report on permissions for one or many recipients. This script includes the function to remove permissions which are deemed as orphaned such as a Deleted Accounts or Disconnected Mailbox Accounts.

### Why so important?

Understanding recipient permissions is fundamental for a smooth migration to Office 365 (or even to another version of Exchange). In a hybrid migration, users selected for migration should group together by those who work together or where sharing permissions are enabled. 

At the time of writing Mailbox Folder Permissions (Sometimes works but not supported by Microsoft), Send-As, and Mailbox Forwarding are not supported or require additional manual configuration between the Cloud and On-Premise. 

Further to this, the removal of the corrupt/invalid permissions prior to migration will reduce the number of failed Mailbox migrations you will encounter meaning that administrators will spend less time reviewing Mailbox Move Reports and more successful first time mailbox moves.

   1. “You should ensure all permissions are explicitly granted and all objects are mail enabled prior to migration. Therefore, you have to plan for configuring these permissions in Office 365 if applicable for your organization.  In the case of Send As permissions, if the user and the resource attempting to be sent as aren’t moved at the same time, you'll need to explicitly add the Send As permission in Exchange Online.”
   
   2. “While mailbox forwarding is supported in Exchange Online, the forwarding configuration isn't copied to Exchange Online when the mailbox is migrated there. Before you migrate a mailbox to Exchange Online, make sure you export the forwarding configuration for each mailbox. ”
   
   3. “Since we are now incrementing the bad item count for each corrupt/invalid permission, this means that if we encounter more corrupt/invalid permissions than your current bad item limit is set to (default is 10 for a migration batch), the migration will fail. Depending on the state of permissions, you could potentially see a LOT of bad entries being logged.”


Ref:  https://technet.microsoft.com/library/jj200581(v=exchg.150).aspx 

Ref:  https://blogs.technet.microsoft.com/exchange/2017/05/30/toomanybaditemspermanentexception-error-when-migrating-to-exchange-online/  

### Final Overview

1. Creates a point in time log of all recipient permissions.
2. Provides information to assist the migration team identify and resolve issues prior to the migration.
3. Removes corrupt/invalid permissions which may cause a mailbox migration failure.
