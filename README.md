# Get-RecipientPermissions.ps1

The Get-RecipientPermissions.ps1 is a PowerShell script that will report on permissions for one or many recipients. This script includes the function to remove permissions which are deemed as orphaned such as a Deleted Accounts or Disconnected Mailbox Accounts.

## Overview

1. Creates a point in time log of all recipient permissions.
2. Provides information to identify and resolve issues prior to the migration.
3. Removes corrupt/invalid permissions which may cause a mailbox migration failure.

## LINKS

[Exchange Hybrid Deployment Considerations](https://technet.microsoft.com/library/jj200581(v=exchg.150).aspx)

[Invalid Permissions Impact to Bad Item Count](https://blogs.technet.microsoft.com/exchange/2017/05/30/toomanybaditemspermanentexception-error-when-migrating-to-exchange-online/)  

## Important points to note

### Function Grant-PermissionRemoval

This Function acts as a broker for when a PERMISSION (ie. Mailbox Folder, Send-As, Full Mailbox Permission etc) should be removed. Input data is presented from the Find-User Function and passes through a Decisions Matrix as to whether the object is should be removed for not. Function responds with a Boolean value if $True Removal Granted, if $False Removal Denied.  

PERMISSION Removal is only completed when the script is run with the switch -PerformRemoval 
    
Account Type  |  Status | Default Decision | Notes
--- | --- | --- | --- 
Deleted Object     |  No Object Found               |   DELETE |               Only a SID is found by a regex match. Exchange is unable to resolve the SID to a Name. 
User NO Mailbox |     Enabled or Disabled  |            LEAVE     |            User Object at one time had a mailbox and a permission was applied to a Mailbox folder or but this mailbox has now been disconnected but the AD Object still exists. 
Group NO Email   |    Enabled or Disabled |             LEAVE  |               Group at one time was mail enabled and a permission for this group was applied. Since then the group 
ADObjectNotFound    | Object Not found in AD          | LEAVE |                Typically displayed because the user account has recently been removed from AD (+-15Min). This is also a fall back value for Find-User for when SID match Fails along with Get-Recipient Get-User and Get-Group. 
Linked Mailbox    |   No Object Found in Remote Forest | LEAVE |                Will delete misconfigured linked mailbox folder permissions. 
Linked Mailbox    |   No Linked Master Account  SID   | LEAVE     |            Assumed account is misconfigured. 
Normal Mailbox    |   Enabled or Disabled   |           LEAVE        |         Default setting is not to delete any permission unless matches occur. 

To update the default behaviour change the following Variables. 

```powershell
     #Set Function Variables - Change here if different results are required
     $removeDeletedUser? = $True
     $removeDisabledUserNoMailbox? = $False     
     $removeUserNoMailbox? = $False
     $removeGroupNoEmail? = $False

         $removeLinkedMailboxAll? = $False
         $removeLinkedMailboxSuccessCrossForest? = $False
         $removeLinkedMailboxFailedCrossForest? = $False
         $removeLinkedMailboxMissingLinkedMasterAccount? = $False
         $removeADObjectNotFound? = $False
```

### Linked Mailboxes 

Please test this first and report any bugs. We have writen support for the Resource Forest model but _have not completed ANY Testing_. 

### Get-Help Get-RecipientPermissions.ps1 -Full

```Powershell
.SYNOPSIS
#This script will run a series of cmdlets / functions to create a report of user permissions. Useful for Office 365 engagements where you need to remediate permissions issues before migration.

.DESCRIPTION
#This script arranges building-block cmdlets / functions to connect to an Exchange environment and loops through all or a subset of mailboxes,  with an account with at a minimum read only access to exchange and active directory.  

Get-RecipientPermissions.ps1 [-Identity <string[Username]> or <Array[Get-Recipient]>] [-PerformRemoval] [-ExportCSV ] [-ExportXML] [-ExportPath <string[]>] [-EnableTranscript]
 
 Search-MailboxFolderPermission [-Identity <Array[Get-Recipient]> or <Array[Get-Recipient]>] 
 
 Search-FullMailboxPermission [-Identity <Array[Get-Recipient]> or <Array[Get-Recipient]>]
 
 Search-PublicDelegatesPermission [-Identity <Array[Get-Recipient]> or <Array[Get-Recipient]>]
 
 Search-PublicFolderPermission [-Identity <Array[Get-Recipient]> or <Array[Get-Recipient]>]
 
 Search-ReceiveAsPermission [-Identity <Array[Get-Recipient]> or <Array[Get-Recipient]>]
 
 Search-SendAsPermission [-Identity <Array[Get-Recipient]> or <Array[Get-Recipient]>]
 
 Search-SendOnBehalfPermission [-Identity <Array[Get-Recipient]> or <Array[Get-Recipient]>]
 
 Search-MailboxForwarding [-Identity <Array[Get-Recipient]> or <Array[Get-Recipient]>]

.PARAMETER Identity
#The Identity parameter specifies the mailbox that you want to view. You can also enter a user's samaccount name or alias. Bulk request can be completed by first creating a variable and piping this variable into the script  

.PARAMETER ExportCSV
#Specifies that all results will be exported to a CSV file. This is a switch only and the filename will be  set via the  script  in the  format of 20180508T014040Z.csv

.PARAMETER ExportXML
#Specifies that all results will be exported to a XML file. This is a switch only and the filename will be  set via the  script  in the  format of 20180508T014040Z.XML

.PARAMETER ExportPath
#Specifies a Path for all exports. This should notinclude a trailing \ and should be included in ''

.PARAMETER PerformRemoval
#The PerformRemoval parameter switch specifies that any invaild recepients should be removed. **Warning** this switch makes active mailbox folder permission changes. Please test before running in a production enviroment. Testing with a -Whatif switch is always recommended. 

.PARAMETER ExportToEmail
#**PENDING 2.0.0**Requires an email address for a report to be compiled and email to a administrator

.PARAMETER ADServer
#**PENDING 2.0.0**The ADServer parameter specifies the Active Directory Server which is to complete AD related requests. Please note that the Exchange server switch should also be used.  

.PARAMETER ExchangeServer
#**PENDING 2.0.0**The ResultSize parameter specifies the maximum number of results to return. If you want to return all requests that match the query, use unlimited for the value of this parameter. The default value is 10.

.PARAMETER UserNotifcation
#**PENDING 2.0.0**Specifies that user will get a email detailing information about the current state of his mailbox and any action recommendations. 

.EXAMPLE
Get-RecipientPermissions.ps1 -Identity <User>

#-- NO PARAMETERS DEFINED --

#When running the script with no parameters, it will prompt for any values which are mandatory. 

.EXAMPLE
.\Get-RecipientPermissions.ps1 -Identity JBines -ExportCSV

#-- CREATE CSV REPORT OF A SINGLE MAILBOX --

#In this example, the mandatory parameters have been provided and the ACTION(s) -Identity and -ExportCSV has been enabled which will create a CSV file containing all permission infomation for this user in the location where the script is run.

.EXAMPLE
$AllUsers = Get-Recipient -RecipientTypeDetails 'UserMailbox'

$AllUsers | .\Get-RecipientPermissions.ps1 -ExportCSV

#-- FIND PERMISSIONS ON A BULK NUMBER OF MAILBOXES WITH CSV EXPORT --

#In this example, the mandatory parameters have been provided and NO ACTION(s) have been enabled. Results will display to the console and will include all permissions of the users. 

.EXAMPLE
$AllUsers = Get-Recipient -ResultSize 'Unlimited'

$AllUsers | .\Get-RecipientPermissions.ps1 -PerformRemoval -ExportCSV -Verbose -WhatIf

#-- TEST ORPHANED OBJECTS REMOVAL ON ALL RECIPIENTS WITH VERBOUS LOGGING ENABLED--

#In this example, the mandatory parameters have been provided and the ACTION(s) -PerformRemoval, -ExportCSV, -Verbose and -WhatIf have been enabled. This enables a high level of logging to the console. The seach results will include any orphaned user permissions and test the removal function of the script without completing any actions on the selected recipients.   

.EXAMPLE
$AllUsers = Get-Recipient -ResultSize 'Unlimited'

$AllUsers | .\Get-RecipientPermissions.ps1 -PerformRemoval -ExportCSV -Verbose -Confirm:$False

#-- REMOVE ORPHANED OBJECTS ON ALL RECIPIENTS WITH CONSOLE LOGGING ENABLED AND WITHOUT PROMPT--

#In this example, the mandatory parameters have been provided and the ACTION(s) -PerformRemoval, -ExportCSV, -Verbose and -Confirm have been enabled. This enables a high level of logging to the console. The seach results will include any orphaned user permissions and will remove any of the permissions set in the Grant-PermissionRemoval Function.   

.EXAMPLE
Get-UserMailPermissions.ps1 -Identity Bines -ExportCSV -ExportXML -ExportPath 'E:\Scripts\Exports' -Verbose 

#-- CREATE CSV & XML REPORT OF A SINGLE MAILBOX WITH EXPORT PATH --

#In this example, the mandatory parameters have been provided and the ACTION(s) -ExportCSV and -ExportXML have been enabled while the -ExportPath has been set to 'E:\Scripts\Exports'. -Verbose is also enabled and will allow detailed logging information in the console window.

.EXAMPLE
Search-RecipientPermissions -Identity (Get-Recipient jjbin0) | Export-Csv -Path E:\Scripts\Exports\test2.csv

#-- CREATE CSV REPORT OF ONLY MAILBOX FOLDER PERMISSIONS FOR A SINGLE USER --

#In this example, the mandatory parameters have been provided and the ACTION(s) -Identity have been populated. These results are exported to the Export-CSV CMDlet. 

.EXAMPLE
$AllUsers = Get-Recipient -RecipientTypeDetails 'UserMailbox'

$AllUsers | Search-MailboxFolderPermission | Export-Csv -Path E:\Scripts\Exports\test2.csv

#-- CREATE CSV REPORT OF ONLY MAILBOX FOLDER PERMISSIONS FOR BULK RECIPIENTS --

#In this example, the mandatory parameters have been provided and the ACTION(s) -Identity have been populated. These results are exported to the Export-CSV CMDlet. 

.LINK
 
Exchange Hybrid Deployment Considerations - https://technet.microsoft.com/library/jj200581(v=exchg.150).aspx

.NOTES
Large environments will take a significant amount of time to scan (hours/days). You can reduce the run time by running the script in batches or multiple instances

Important: Do not run too many instances or against too many mailboxes at once. Doing so could cause performance issues, affecting users. The Author or Contributors are not responsible for any such performance issues or improper use, or a lack of planning and testing.
```


 

