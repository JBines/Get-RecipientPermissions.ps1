<#
.SYNOPSIS
This script will run a series of cmdlets / functions to create a report of user permissions. Useful for Office 365 engagements where you need to remediate permissions issues before migration.

.DESCRIPTION
This script arranges building-block cmdlets / functions to connect to an Exchange environment and loops through all or a subset of mailboxes,  with an account with at a minimum read only access to exchange and active directory.  

Get-RecipientPermissions.ps1 [-Identity <string[Username]> or <Array[Get-Recipient]>] [-PerformRemoval] [-ExportCSV ] [-ExportXML] [-ExportPath <string[]>] [-EnableTranscript]
 
 Search-MailboxFolderPermission [-Identity <Array[Get-Recipient]> or <Array[Get-Recipient]>] 
 
 Search-FullMailboxPermission [-Identity <Array[Get-Recipient]> or <Array[Get-Recipient]>]
 
 Search-PublicDelegatesPermission [-Identity <Array[Get-Recipient]> or <Array[Get-Recipient]>]
 
 Search-PublicFolderPermission [-Identity <Array[Get-Recipient]> or <Array[Get-Recipient]>]
 
 Search-ReceiveAsPermission [-Identity <Array[Get-Recipient]> or <Array[Get-Recipient]>]
 
 Search-SendAsPermission [-Identity <Array[Get-Recipient]> or <Array[Get-Recipient]>]
 
 Search-SendOnBehalfPermission [-Identity <Array[Get-Recipient]> or <Array[Get-Recipient]>]
 
 Search-RecipientForwarding [-Identity <Array[Get-Recipient]> or <Array[Get-Recipient]>]

.PARAMETER Identity
The Identity parameter specifies the mailbox that you want to view. You can also enter a user's samaccount name or alias. Bulk request can be completed by first creating a variable and piping this variable into the script  

.PARAMETER ExportCSV
Specifies that all results will be exported to a CSV file. This is a switch only and the filename will be  set via the  script  in the  format of 20180508T014040Z.csv

.PARAMETER ExportXML
Specifies that all results will be exported to a XML file. This is a switch only and the filename will be  set via the  script  in the  format of 20180508T014040Z.XML

.PARAMETER ExportPath
Specifies a Path for all exports. This should notinclude a trailing \ and should be included in ''

.PARAMETER PerformRemoval
The PerformRemoval parameter switch specifies that any invaild recepients should be removed. **Warning** this switch makes active mailbox folder permission changes. Please test before running in a production enviroment. Testing with a -Whatif switch is always recommended. 

.PARAMETER ExportToEmail
**PENDING 2.0.0**Requires an email address for a report to be compiled and email to a administrator

.PARAMETER ADServer
**PENDING 2.0.0**The ADServer parameter specifies the Active Directory Server which is to complete AD related requests. Please note that the Exchange server switch should also be used.  

.PARAMETER ExchangeServer
**PENDING 2.0.0**The ResultSize parameter specifies the maximum number of results to return. If you want to return all requests that match the query, use unlimited for the value of this parameter. The default value is 10.

.PARAMETER UserNotifcation
**PENDING 2.0.0**Specifies that user will get a email detailing information about the current state of his mailbox and any action recommendations. 

.EXAMPLE
Get-RecipientPermissions.ps1 -Identity <User>

-- NO PARAMETERS DEFINED --

When running the script with no parameters, it will prompt for any values which are mandatory. 

.EXAMPLE
.\Get-RecipientPermissions.ps1 -Identity JBines -ExportCSV

-- CREATE CSV REPORT OF A SINGLE MAILBOX --

In this example, the mandatory parameters have been provided and the ACTION(s) -Identity and -ExportCSV has been enabled which will create a CSV file containing all permission infomation for this user in the location where the script is run.

.EXAMPLE
$AllUsers = Get-Recipient -RecipientTypeDetails 'UserMailbox'

$AllUsers | .\Get-RecipientPermissions.ps1 -ExportCSV

-- FIND PERMISSIONS ON A BULK NUMBER OF MAILBOXES WITH CSV EXPORT --

In this example, the mandatory parameters have been provided and NO ACTION(s) have been enabled. Results will display to the console and will include all permissions of the users. 

.EXAMPLE
$AllUsers = Get-Recipient -ResultSize 'Unlimited'

$AllUsers | .\Get-RecipientPermissions.ps1 -PerformRemoval -ExportCSV -Verbose -WhatIf

-- TEST ORPHANED OBJECTS REMOVAL ON ALL RECIPIENTS WITH VERBOUS LOGGING ENABLED--

In this example, the mandatory parameters have been provided and the ACTION(s) -PerformRemoval, -ExportCSV, -Verbose and -WhatIf have been enabled. This enables a high level of logging to the console. The seach results will include any orphaned user permissions and test the removal function of the script without completing any actions on the selected recipients.   

.EXAMPLE
$AllUsers = Get-Recipient -ResultSize 'Unlimited'

$AllUsers | .\Get-RecipientPermissions.ps1 -PerformRemoval -ExportCSV -Verbose -Confirm:$False

-- REMOVE ORPHANED OBJECTS ON ALL RECIPIENTS WITH CONSOLE LOGGING ENABLED AND WITHOUT PROMPT--

In this example, the mandatory parameters have been provided and the ACTION(s) -PerformRemoval, -ExportCSV, -Verbose and -Confirm have been enabled. This enables a high level of logging to the console. The seach results will include any orphaned user permissions and will remove any of the permissions set in the Grant-PermissionRemoval Function.   

.EXAMPLE
Get-RecipientPermissions.ps1 -Identity Bines -ExportCSV -ExportXML -ExportPath 'E:\Scripts\Exports' -Verbose 

-- CREATE CSV & XML REPORT OF A SINGLE MAILBOX WITH EXPORT PATH --

In this example, the mandatory parameters have been provided and the ACTION(s) -ExportCSV and -ExportXML have been enabled while the -ExportPath has been set to 'E:\Scripts\Exports'. -Verbose is also enabled and will allow detailed logging information in the console window.

.EXAMPLE
Search-RecipientPermissions -Identity (Get-Recipient jjbin0) | Export-Csv -Path E:\Scripts\Exports\test2.csv

-- CREATE CSV REPORT OF ONLY MAILBOX FOLDER PERMISSIONS FOR A SINGLE USER --

In this example, the mandatory parameters have been provided and the ACTION(s) -Identity have been populated. These results are exported to the Export-CSV CMDlet. 

.EXAMPLE
$AllUsers = Get-Recipient -RecipientTypeDetails 'UserMailbox'

$AllUsers | Search-MailboxFolderPermission | Export-Csv -Path E:\Scripts\Exports\test2.csv

-- CREATE CSV REPORT OF ONLY MAILBOX FOLDER PERMISSIONS FOR BULK RECIPIENTS --

In this example, the mandatory parameters have been provided and the ACTION(s) -Identity have been populated. These results are exported to the Export-CSV CMDlet. 

.EXAMPLE
$AllPublicFolders = Get-PublicFolder -Recurse

$AllPublicFolders | Search-PublicFolderPermission | Export-Csv PF_Export.csv

-- CREATE CSV REPORT OF ONLY PUBLIC FOLDER PERMISSIONS FOR BULK RECIPIENTS CAN BE NON MAIL ENABLED --

In this example, the mandatory parameters have been provided and the ACTION(s) -Identity have been populated. These results are exported to the Export-CSV CMDlet. 

.LINK
 
Exchange Hybrid Deployment Considerations - https://technet.microsoft.com/library/jj200581(v=exchg.150).aspx

.NOTES
IMPORTANT! We recommend you run this script from the Exchange Management Shell for the best results. 

Large environments will take a significant amount of time to scan (hours/days). You can reduce the run time by running the script in batches or multiple instances

Important: Do not run too many instances or against too many mailboxes at once. Doing so could cause performance issues, affecting users. The Author or Contributors are not responsible for issues or improper use, or a lack of planning and testing.

[AUTHOR]
 Joshua Bines, Consultant

Find me on:
* Web:	    https://theinformationstore.com.au
* LinkedIn:	https://www.linkedin.com/in/joshua-bines-4451534
* Github:	https://github.com/jbines
  
[CONTRIBUTORS]
 Mihail Popa, Senior Engineer 
 
[VERSION HISTORY / UPDATES]
 0.0.0 20180503 - JBINES - Created the bare bones
 0.0.3 20180507 - JBINES - Console Output with User Permissions.
 0.0.4 20180510 - JBINES - Added Verbose and File Export.
 0.0.5 20180517 - JBINES - Updated to include Linked Mailbox Objects and Invaild permissions on Mailbox Folders.
                         - BUG FIX: Missing Mailbox Folder 'Top of Information Store'
                         - BUG FIX: Fix Mailbox Folders with Special Characters ie/\ 
 0.0.5 20180522 - JBINES - Added Transcript option for console logging. Change Heading from Target to Source. Ammend Console and Verbous output. Extra RBAC for Permission removal. Added ScriptStopwatch. 
 0.0.6 20180528 - JBINES - BUG FIX: Resolved and tested issues with default folder name changes non english mailboxes. 
                         - Added a check on the SID Object Matched by REGEX as a Deleted User to confirm the object does not exist before it deletes the object.
 0.0.7 20180529 - JBINES - Added Functions Find-SIDObject and Find-User and Script Report Action. 
 0.0.8 20180531 - JBINES - BUG FIX: Find-User Incorrectly ideniftying objects
                         - Added Function Grant-PermissionRemoval
 0.0.9 20180604 - JBINES - Added Function New-ArrayObject, Search-FullMailboxPermission
                         - BUG FIX: SID Check to Function Grant-PermissionRemoval
 0.0.10 20180604- JBINES - Added Function Search-SendOnBehalfPermission, Search-SendAsPermission, Search-ReceiveAsPermission, Search-PublicDelegatesPermission
 0.1.0 20180607 - JBINES - Added Function Search-MailboxFolderPermission, Search-PublicFolderPermission
 0.1.1 20180611 - JBINES - Updated Help information. Enabled piping to Export-CSV. Amened Console output. Allowed strings to be used for Search FUNCTIONS
 0.1.2 20180612 - JBINES - Added Function Search-MailboxForwarding
 0.1.3 20190702 - JBINES - BUG FIX:CommonParameters for some exchange CMDlets are not working correctly instead we have had to change the global VAR $ErrorActionPreference
                         - BUG FIX:Skip Audit Folders in mailboxes "Non-system logon cannot access Audits folder."
 0.1.4 20190715 - JBINES - BUG FIX: Updated Search-MailboxFolderPermission to allow a loop break on mailboxes in dismounted DBs. Also move to Guid where Mailnick and SamAccountName do not match and other dodgy objects. 
 0.1.5 20190716 - JBINES - Added Suport for Exchange 2013 for Search-PublicFolderPermission. Still need to test Excahnge 2016 and 2019 but I belive it should work. 
                         - BUG FIX:When get-recepient returns an array higher than 1. Added Select-Object -First 1"
 0.1.6 20190722 - JBINES - Added Suport for non mail Enabled Search-PublicFolderPermissions. 
 0.1.7 20191230 - JBINES - Added New Recipient_SamAccountname to Function New-ArrayObject. 
                         - BUG FIX: Allowed script to continue on error for selected Functions. 
                         - BUG FIX: To fix a bug fix (1.3) changing the $ErrorActionPreference was a silly idea. Fixed random issues I was seeing with the remote PS but inturn created other issues so a roll back was needed. Added note to make sure the script is run from the Exchange Shell
 0.1.8 20200519 - JBINES - BUG FIX: Function Search-RecipientForwarding was missing the -PerformRemoval Switch which was causing an error at runtime. Thx astavitsky for logging the issue. https://github.com/JBines/Get-RecipientPermissions.ps1/issues/3

[TO DO LIST / PRIORITY]
 HIGH - Add XML backup of removed permissions
 HIGH - Exchange ActiveSync clients
 HIGH - Add Support for EXO
 HIGH - Add Support Remote PS Sessions
 MED - Flag duplicates from the Find-User
 MED - Write Log for troubleshooting (Use Verbose for now with Transcript)
 MED - Expand DLs for with a full user list
 LOW - Feature to Target Exchange and AD Servers By Name
 LOW - SID Check via Resource Forest with Service Account and Pass
 LOW - Add OU Scope or LDAP Filter to find User Permissions not in Azure AD
 LOW - Enable/Test Permissions from Azure AD and Exchange Online
 
#>

[CmdletBinding(SupportsShouldProcess=$True)]
Param 
(
	[Parameter(Position=0, Mandatory = $True, ValueFromPipeline = $True, HelpMessage="Please provide the Samaccount name of the User you would like to check?")]
    [ValidateNotNullOrEmpty()]
    $Identity,
    [Parameter(Mandatory = $False)]
    [Switch]$ExportCSV=$False,
    [Parameter(Mandatory = $False)]
    [Switch]$ExportXML=$False,
    [Parameter(Mandatory = $False)]
    [System.String]$ExportPath=$null,
    [Parameter(Mandatory = $False)]
    [Switch]$EnableTranscript=$False,
    [Parameter(Mandatory = $False)]
    [Switch]$PerformRemoval=$False    
)

Begin{

    #Start Script Timing StopWatch
    $TotalScriptStopWatch = [system.diagnostics.stopwatch]::startNew()
    
    #BUG FIX - Changes Global Action Preference in Find-User if the not set to STOP
    $ErrorActionPreferenceChanged = $False
    $oldActionPreference = $ErrorActionPreference

    if(($ErrorActionPreference -eq 'Stop') -or ($ErrorActionPreference -eq 'Suspend') -or ($ErrorActionPreference -eq 'Inquire')){
    
        Write-Warning "Please note your current ErrorActionPreference is set to $ErrorActionPreference and is not recommended for long running batches"

    }

    
    #If Switch $EnableTranscript used start Console logging via Start-Transcript CMDlet
    
    If($EnableTranscript){
            
            $datelog = ((get-date).ToUniversalTime()).ToString("yyyyMMddThhmmssZ") + ".log"
            Start-Transcript $datelog -Append
            
            #CmdletBinding skiping starting transcipt if WhatIf is enabled
            If($WhatIfPreference -eq $true){
            
            $WhatIfPreference = $False
            
            Start-Transcript $datelog -Append
            
            $WhatIfPreference = $True
            
            }
    
    }

    #Load Script Functions
    Function Test-CommandExists

    {

     Param ($command)

         $oldPreference = $ErrorActionPreference

         $ErrorActionPreference = 'stop'

         try {if(Get-Command $command){RETURN $true}}

         Catch {Write-Host "$command does not exist"; RETURN $false}

         Finally {$ErrorActionPreference=$oldPreference}

    } #end function test-CommandExists
        
    Function Find-SIDObject
    <#
    . [AUTHOR]
    . Joshua Bines, Consultant
    .
    . [DATE]
    . 20180529
    .
    .DESCRIPTION
    This Function attempts to find a user by a SID value in the local domain where the script is run.
    #>
    {
     [CmdletBinding()]
     Param ([Parameter(Mandatory = $True, ValueFromPipeline = $True)][System.String]$SID)
     
     Begin {$SIDArray = @()}
        
     Process{
     
            Try{
            
                $SIDobj_Translate = (New-Object System.Security.Principal.SecurityIdentifier("$SID")).Translate([System.Security.Principal.NTAccount]).Value

                    if($SIDobj_Translate){
                        
                        Write-Verbose "SID Object has been matched to $SIDobj_Translate";
                        
                        $SIDArray += (New-Object psobject -Property @{
                            'SID' = $SID;
                            'Found' = $True;
                            'TranslatedTo' = $SIDobj_Translate;
                        })
                
                    }
                }
                            
                Catch{ 
                                
                    Write-Verbose "Object SID $SIDobj Matched by REGEX and Failed Local AD Lookup";
                        $SIDArray += (New-Object psobject -Property @{
                            'SID' = $SID;
                            'Found' = $False;
                            'TranslatedTo' = "UNKNOWN";
                        })
                }
                            
    }
    
    END {$SIDArray}

    } #end function Find-SIDObject

    Function Find-User
    <#
    . [AUTHOR]
    . Joshua Bines, Consultant
    .
    . [DATE]
    . 20180529
    .
    .DESCRIPTION
    This Function attempts to locate a user object from a Username or Domain\Username in the local domain where the script is run.
    #>

    {
     [CmdletBinding()]
     Param ([Parameter(Mandatory = $True, ValueFromPipeline = $False)]$User)
     
     Begin {$userArray = @()
            
        <# Removed as the bug found for this code does not affect script when running from EMS
        if($ErrorActionPreference -ne "STOP"){
        
            #$ErrorActionPreferenceChanged = $ErrorActionPreference
            Set-Variable -Name ErrorActionPreferenceChanged -Value $ErrorActionPreference -Scope Global

            Write-Verbose "Set Global Variable ErrorActionPreferenceChanged: $ErrorActionPreferenceChanged"
            #Set-Variable -Name ErrorActionPreference -Value "STOP" -Scope Global
            If($?){Write-Verbose "FUNCTION Find-User: Changed $ErrorActionPreference to Stop"}

        }
        #>

     }
        
     Process{
            
            If($User -ne $null){     
                
                #Null Var
                $userObj = $Null
                $adUser = $Null
                $adUserEnabled = $Null
                $userStatus = $Null
                $userDisplayName = $Null
                $userSamAccountName = $Null
                $userRecipientTypeDetails = $Null
                $userDistinguishedName = $Null
                $userIsValid = $Null
                $userEnabled = $Null
                $SIDCheck = $Null
                
                
                Try{
                    
                    #In the the event of duplicates take the first of the array (Look at Flaging this in the report!)
                    $userObj = Get-Recipient $User -ErrorAction STOP | Select-Object -First 1                    

                    If($userObj.DistinguishedName){
                            #Set Var for Array
                            $userDisplayName = $userObj.DisplayName
                            $userSamAccountName = $userObj.SamAccountName
                            $userRecipientTypeDetails = $userObj.RecipientTypeDetails
                            $userDistinguishedName = $userObj.DistinguishedName
                            $userIsValid = $True                                                        
                            $userStatus += 'Get-Recipient-Succeeded;'
                    }                
                }
                
                Catch{
                
                    Write-Verbose "Failed Get-Recipient $($_.Exception.Message)"
                    $userStatus += 'Get-Recipient-Failed;'

                    #Check for Deleted Account
                    If($User -match 'S-\d-\d+-(\d+-){1,14}\d+$'){
                        
                        Write-Verbose "FUNCTION Find-User Found Deleted Account: $User";
                        
                        #Set Var
                        $userDisplayName = "DeletedUser"
                        $userSamAccountName = $User
                        $userRecipientTypeDetails = "DeletedUser"
                        $userIsValid = $False
                        $userStatus += "RegexMatched-DeletedUser;"
                        $adUserEnabled = $False

                    }
                    
                    Else{
                        
                        #Check for Disabled and Disconnected Mailboxes
                        Try{

                            $userObj = Get-User -Identity $User.ToString() -ErrorAction STOP | Select-Object -First 1 
                            
                            If($userObj.DistinguishedName){

                                    #Set Var for Array
                                    $userDisplayName = $userObj.DisplayName
                                    $userSamAccountName = $userObj.SamAccountName
                                    $userRecipientTypeDetails = $userObj.RecipientTypeDetails
                                    $userDistinguishedName = $userObj.DistinguishedName
                                    $userIsValid = $True
                                    $userStatus += 'Get-User-Succeeded;'
                            }
                            
                        }
                        
                        Catch{
                            Write-Verbose "FUNCTION Find-User Failed Get-User $($_.Exception.Message)"
                            $userStatus += 'Get-User-Failed; '
                            $userStatus += $_.Exception.Message
                            
                            #Set Var for Array
                            $userDisplayName = "ADObjectNotFound"
                            $userSamAccountName = $User
                            $userRecipientTypeDetails = "ADObjectNotFound"
                            $userIsValid = $False
                            $userEnabled = $False
                                                    
                        }
                    
                    }
                    
                    
                }
                
                Finally{
                    
                    #Check Status of Linked Accounts
                    If($userObj.RecipientTypeDetails -eq "LinkedMailbox"){
                        
                        Try{
                        
                            $userLinkedMasterAccount = (Get-User $userObj.DistinguishedName).linkedmasteraccount

                        }
                        Catch{
                        
                            Write-Verbose "FUNCTION Find-User Failed LinkedMailbox Check: $($_.Exception.Message)"

                        }
                        
                        #Check cross forest sid resolution is correct. 
                            If($userLinkedMasterAccount.length -gt 0){
                                
                                Write-Verbose "FUNCTION Find-User Linked Object Found: $($User.Name) Linked Master Account SID: $userLinkedMasterAccount"; 

                                $SIDCheck = Find-SIDObject -SID $userLinkedMasterAccount
                                
                                Switch ($SIDCheck.found){
                                
                                    $True {
                                    
                                        Write-Verbose "FUNCTION Find-User SID Translatation Successful for Object $($User.displayname) with Linked Master Account SID: $userLinkedMasterAccount"
                                        $userStatus = "LinkedMailbox-SIDSucceededCrossForestResolution"
                                    
                                    }
                                    
                                    $False {
                                    
                                        Write-Verbose "FUNCTION Find-User SID Translatation Failed for Object $($User.displayname) with Linked Master Account SID: $userLinkedMasterAccount";
                                        $userStatus = "LinkedMailbox-SIDFailedCrossForestResolution"
                                    
                                    }
                                    
                                    Default {
                                    
                                        Write-Verbose "FUNCTION Find-User SID Translatation Failed. Item is is listed by the Find-SIDObject as Neither Enabled or Disabled"
                                    
                                    }
                                    
                                }
                            }
                            
                            Else {
                            
                                Write-Verbose "FUNCTION Find-User Invaild Linked Object Found: MISSING Linked Master Account SID"; 
                                $userStatus = "LinkedMailbox-MissingLinkedMasterAccountSID"
                            }
                                                                                    
                    }
                    
                    
                    #Add to arrary if User is Enabled or Disabled
                    if($userObj.DistinguishedName){
                        Write-Verbose "FUNCTION Find-User Object has DN of $($USERobj.DistinguishedName) Checking Enabled/Disabled Status"
                        
                        Try{                        
                            $adUser = [adsi]"LDAP://$($USERobj.DistinguishedName)" 
                            $uac=$adUser.psbase.invokeget("useraccountcontrol") 
                            If($uac -band 0x2)  
                            { Write-Verbose "DISABLED: $($USERobj.DistinguishedName)" ; $adUserEnabled = $False }
                            Else
                            { Write-Verbose "ENABLED: $($USERobj.DistinguishedName)";$adUserEnabled =$True }
                        
                            #Set Var 
                            $userEnabled = $adUserEnabled                       
                        }
                        Catch{
                            
                            Write-Verbose "FUNCTION Find-User Failed Enabled/Disabled Check: $($_.Exception.Message)"
                        }
                    }

                #Set Variables to Array Object
                $UserArray += (New-Object psobject -Property @{
                    'User' = $User;
                    'DisplayName' = $userDisplayName;
                    'SamAccountName' = $userSamAccountName;
                    'RecipientTypeDetails' = $userRecipientTypeDetails;
                    'DistinguishedName' = $userDistinguishedName;
                    'Enabled' = $userEnabled;
                    'IsValid' = $userIsValid;
                    'Status' = $userStatus;
                })            
                    
                }
            }
            else{
                
               Write-Verbose "Function Find-User: Null or More that 1 object found Skipping lookup"
            }
    }
    
    END {
    
        #Reapply Default ErrorActionPreference Value
        <#
        if($ErrorActionPreferenceChanged -ne $False){
            
            Set-Variable -Name ErrorActionPreference -Value $ErrorActionPreferenceChanged -Scope Global
            #$ErrorActionPreference = $ErrorActionPreferenceChanged
            If($?){Write-Verbose "Function Find-User: Revert $ErrorActionPreference Back To: $ErrorActionPreferenceChanged"}

        }#>

        $UserArray

    }

    } #end function Find-User
    
    Function Grant-PermissionRemoval
    <#
    . [AUTHOR]
    . Joshua Bines, Consultant
    .
    . [DATE]
    . 20180531
    .
    .DESCRIPTION
    This Function acts as a broker for when a PERMISSION (ie. Mailbox Folder, Send-As, Full Mailbox Permission etc) should be removed. Input data is presented from the Find-User Function and passes through a Decisions Matrix as to whether the object is should be removed for not. Function responds with a Boolean value if $True Removal Granted, if $False Removal Denied.  
    
    .NOTES
    Important: Decisions have been made to enhance a smooth transition to Exchange Online, but we recommend steps should be taken to confirm if these decisions reflect your requirements.
    
    #Account Type#       #Status#                        #Default Decision#     #Notes#
     Deleted Object       No Object Found                  DELETE                Only a SID is found by a regex match. Exchange is unable to resolve the SID to a Name. 
     User NO Mailbox      Enabled or Disabled              LEAVE                 User Object at one time had a mailbox and a permission was applied to a Mailbox folder or but this mailbox has now been disconnected but the AD Object still exists. 
     ADObjectNotFound     Object Not found in AD           LEAVE                 Typically displayed because the user account has recently been removed from AD (+-15Min). This is also a fall back value for Find-User for when SID match Fails along with Get-Recipient and Get-User. 
     Linked Mailbox       No Object Found in Remote Forest LEAVE                 Will delete misconfigured linked mailbox folder permissions. 
     Linked Mailbox       No Linked Master Account  SID    LEAVE                 Assumed account is misconfigured. 
     Normal Mailbox       Enabled or Disabled              LEAVE                 Default setting is not to delete any permission unless matches occur. 
     
    #>
    {
     [CmdletBinding()]
     Param (
     
    [Parameter(Mandatory = $False)]
    [System.String]$SamAccountName,
    
    [Parameter(Mandatory = $True)]
    [System.String]$RecipientType,

    [Parameter(Mandatory = $False)]
    [System.String]$Status
     
     )
     
     Begin {
     
     #Set Function Variables - Change here if different results are required
     $removeDeletedUser? = $True
     $removeDisabledUserNoMailbox? = $False
     
         $removeUserNoMailbox? = $False 
         $removeLinkedMailboxAll? = $False
         $removeLinkedMailboxSuccessCrossForest? = $False
         $removeLinkedMailboxFailedCrossForest? = $False
         $removeLinkedMailboxMissingLinkedMasterAccount? = $False
         $removeADObjectNotFound? = $False
          
     }
        
     Process{
     
     #Create Array
     $Result = $False
     $SIDCheck =$Null
        
        If($RecipientType){
            Switch($RecipientType){
                "DeletedUser"{
                                
                                
                                $SIDCheck = Find-SIDObject -SID $SamAccountName
                                If ($SIDCheck.found -eq $False){
                                    
                                    $Result = $removeDeletedUser?
                                
                                }
                                Else{
                                
                                Write-Verbose "FUNCTION Grant-PermissionRemoval SID FOUND SID: $SamAccountName TranslatedTo $($SIDCheck.TranslatedTo)"
                                
                                }
                                
                                }
                "LinkedMailbox"{
                                
                                If($Status -eq "LinkedMailbox-SuccessCrossForestResolution"){$Result = $removeLinkedMailboxSuccessCrossForest?}
                                If($Status -eq "LinkedMailbox-MissingLinkedMasterAccount"){$Result = $removeLinkedMailboxMissingLinkedMasterAccount?}
                                If($Status -eq "LinkedMailbox-SIDFailedCrossForestResolution"){$Result = $removeLinkedMailboxFailedCrossForest?}
                                Else{$Result = $removeLinkedMailboxAll?}
                                
                                If($removeLinkedMailboxAll?){$Result = $removeLinkedMailboxAll?}
                                                                    
                                }
                "ADObjectNotFound"{
                                
                                $Result = $removeADObjectNotFound?
                                
                                }
                "User"{
                                $Result = $removeUserNoMailbox?
                                
                                } 
                "DisabledUser"{
                                
                                $Result = $removeDisabledUserNoMailbox?
                                
                                }
 
                Default{$Result = $False}
            }
            
            Write-Verbose "FUNCTION Grant-PermissionRemoval User: $SamAccountName; RecipientType: $RecipientType; Result: $Result"
        
        }
    }
    
    END {$Result}

    } #end function Grant-PermissionRemoval

    Function New-ArrayObject
    <#
    . [AUTHOR]
    . Joshua Bines, Consultant
    .
    . [DATE]
    . 20180604
    .
    .DESCRIPTION
    This Function creates the table for the export to console, CSV and XML.

        #>
    {
     Param (
     
    [Parameter(Mandatory = $False)]
    [System.String]$RecipientDisplayName,
    [Parameter(Mandatory = $False)]
    [System.String]$RecipientSamAcc,    
    [Parameter(Mandatory = $False)]
    [System.String]$RecipientType,
    [Parameter(Mandatory = $False)]
    [System.String]$PermissionType,
    [Parameter(Mandatory = $False)]
    [System.String]$SourceDisplayName,
    [Parameter(Mandatory = $False)]
    [System.String]$SourceSamAcc,
    [Parameter(Mandatory = $False)]
    [System.String]$SourceRecipientType,
    [Parameter(Mandatory = $False)]
    [System.String]$Action,
    [Parameter(Mandatory = $False)]
    $Removal     
     )
     
                    New-Object psobject -Property @{
                        'Recipient' = $RecipientDisplayName;
                        'RecipientSamAccountName' = $RecipientSamAcc;
                        'RecipientType' = $RecipientType;
                        'PermissionType' = $PermissionType;
                        'SourceRecipient' = "$SourceDisplayName ($SourceSamAcc) ($SourceRecipientType)";
                        'SourceRecipientSamAccountName' = $SourceSamAcc;
                        'SourceRecipientType' = $SourceRecipientType
                        'ScriptAction' = $Action;                        
                        'PerformRemoval' = $Removal;                        
                    }
    
    } #end function New-ArrayObject
    
    Function Search-FullMailboxPermission
    <#
    . [AUTHOR]
    . Joshua Bines, Consultant
    .
    . [DATE]
    . 20180604
    .
    .DESCRIPTION
    This Function searches for Full Mailbox Permissions, Reports, and removes permissions if the -PerformRemoval Switch is set. 

        #>
    {
     [CmdletBinding(SupportsShouldProcess=$True)]
     Param (
     
    [Parameter(Mandatory = $True, ValueFromPipeline = $True)]
    $Identity,
    
    [Parameter(Mandatory = $False)]
    [switch]$PerformRemoval
     
     )
     
     Begin {
    
    Write-Verbose "FUNCTION Search-FullMailboxPermission: Check for a user data entered into the 'Identity' Switch"
    
    if((($Identity.GetType()).name) -eq 'String'){
        
        Write-Verbose "FUNCTION Search-FullMailboxPermission: Confirmed User entered data of $Identity";
        Write-Verbose "FUNCTION Search-FullMailboxPermission: Attempting to resolve to a Recipient to $Identity";
    
            Try{
            
                $Identity = Get-Recipient $Identity -ErrorAction STOP | Select-Object -First 1 
            
            }
            Catch{
                        
                Write-Error "The Get-Recipient CMDlet returned a error and is unable to continue"; EXIT
                
            }
            
            if (($Identity | Measure-Object).count -gt 1){
            
                Write-Error "The Get-Recipient CMDlet returned more than one result after running the Get-Recipient CMDlet. Please use another switch for completing bulk actions";$Identity_STR_Error; Break
            
            }
            Else {Write-Verbose "FUNCTION Search-FullMailboxPermission: Recipient found"}
    }
     
     #Create PS Array
     [PSObject[]] $FMPreport = @()
     
     #$CMDlet_FMP='Get-MailboxPermission -Identity $Identity.DistinguishedName | where {$_.user.tostring() -ne "NT AUTHORITY\SELF" -and $_.IsInherited -eq $false}';
     #$FMP = Invoke-Expression $CMDlet_FMP
     
     }
        
     Process{
     
     $FMP = Get-MailboxPermission -Identity $Identity.DistinguishedName | Where-Object {$_.user.tostring() -ne "NT AUTHORITY\SELF" -and $_.IsInherited -eq $false}
            
            if($null -ne $FMP){
                foreach($FMPobj in $FMP){
                    
                    Write-Verbose "FUNCTION Search-FullMailboxPermission: Found Full mailbox permisison for $FMPobj.User on Source Recipient $($recipientObj.Name)"
                    
                    #Null Var
                    $FMPobj_USER = $Null
                    $FMPobj_DEL = $Null
                    $FMPobj_Action = 'Report Only'
                    
                    #Find User and Check for Orphanded SID or Object
                    
                    $FMPobj_USER = Find-User $FMPobj.User
                    
                    $FMPobj_DEL = Grant-PermissionRemoval -SamAccountName $FMPobj_USER.SamAccountName -RecipientType $FMPobj_USER.RecipientTypeDetails -Status $FMPobj_USER.Status
                                                                                
                    if(($PerformRemoval) -and ($FMPobj_DEL)) {
                            
                                If($PSCmdlet.ShouldProcess($($FMPobj_USER.DisplayName),"Removing Full mailbox permisison for user $($Identity.DisplayName)")){
                                                                        
                                    Try{
                                        
                                        #Add Support for the -Confirm:$False Switch
                                        If(($PerformRemoval)-and($ConfirmPreference -eq 'None')){
                                        
                                            Remove-MailboxPermission -Identity $Identity.DisplayName -User $FMPobj_USER.SamAccountName -AccessRights FullAccess -InheritanceType All -Confirm:$False
                                        
                                        }
                                        Else{
                                        
                                            Remove-MailboxPermission -Identity $Identity.DisplayName -User $FMPobj_USER.SamAccountName -AccessRights FullAccess -InheritanceType All
                                        
                                        }
                                        
                                        If(($?)-and($WhatIfPreference -ne $True)){
                                        
                                            Write-Verbose "FUNCTION Search-FullMailboxPermission Successful CMDlet: Remove-MailboxPermission  $($Identity.DisplayName) -User $($FMPobj_USER.SamAccountName)"
                                            $FMPobj_Action = "Successful Removal"
                                        
                                        }
                                    
                                    }
                                    
                                    Catch{
                                        
                                        Write-Verbose "FUNCTION Search-FullMailboxPermission Failure CMDlet: Remove-MailboxFolderPermission $_.Exception.Message";
                                        Write-Error "$_.Exception.Message"
                                        $FMPobj_Action = "Failed Removal"
                                    
                                    }
                                     
                                }
                                
                        If(($WhatIfPreference -eq $True) -and ($FMPobj_Action -ne 'Removal Failed')){
                        
                        Write-Verbose "FUNCTION Search-FullMailboxPermission What If Successful CMDlet: Remove-MailboxPermission  $($Identity.DisplayName) -User $($FMPobj_USER.SamAccountName)"
                        $FMPobj_Action = "Successful WhatIf"
                        
                        }

                    }
                        
                    $FMPreport = $FMPreport + (New-ArrayObject -RecipientDisplayName $FMPobj_USER.DisplayName -RecipientSamAcc $FMPobj_USER.SamAccountName -RecipientType $FMPobj_USER.RecipientTypeDetails -PermissionType "Full Mailox Permission" -SourceDisplayName $Identity.Name -SourceSamAcc $Identity.SamAccountName -SourceRecipientType $Identity.RecipientTypeDetails -Action $FMPobj_Action -Removal $FMPobj_DEL)                    

                }
            }
        

    }
    
    END {$FMPreport}

    } #end function Search-FullMailboxPermission

    Function Search-SendOnBehalfPermission
    <#
    . [AUTHOR]
    . Joshua Bines, Consultant
    .
    . [DATE]
    . 20180604
    .
    .DESCRIPTION
    This Function searches for Send On Behalf Permissions and Reports. From our testing we found that no deleted account is left in the Send On Behalf list, but disconnected mailboxes remain. 

        #>
    {
     [CmdletBinding(SupportsShouldProcess=$True,ConfirmImpact="High")]
     Param (
     
    [Parameter(Mandatory = $True, ValueFromPipeline = $True)]
    $Identity,
    
    [Parameter(Mandatory = $False)]
    [switch]$PerformRemoval
     
     )
     
     Begin {
    
    Write-Verbose "FUNCTION Search-SendOnBehalfPermission: Check for a user data entered into the 'Identity' Switch"
    
    if((($Identity.GetType()).name) -eq 'String'){
        
        Write-Verbose "FUNCTION Search-SendOnBehalfPermission: Confirmed User entered data of $Identity";
        Write-Verbose "FUNCTION Search-SendOnBehalfPermission: Attempting to resolve to a Recipient to $Identity";
    
            Try{
            
                $Identity = Get-Recipient $Identity -ErrorAction STOP | Select-Object -First 1 
            
            }
            Catch{
                        
                Write-Error "The Get-Recipient CMDlet returned a error and is unable to continue: Error: $($_.Exception.Message)"; EXIT
                
            }
            
            if (($Identity | Measure-Object).count -gt 1){
            
                Write-Error "The Get-Recipient CMDlet returned more than one result after running the Get-Recipient CMDlet. Please use another switch for completing bulk actions";$Identity_STR_Error; EXIT
            
            }
            Else {Write-Verbose "FUNCTION Search-SendOnBehalfPermission: Recipient found"}
    }

     #Create PS Array
     [PSObject[]] $SOBPreport = @()
     #$ScriptAction? = "Script Action"
          
     }
        
     Process{
     
     #Create Local User Array
     [PSObject[]] $SOBPUserReport = @()
     $SOBPobjDeleteCounterFalse? = 0
     $SOBPobjDeleteCounterTrue? = 0
     #
        switch ($Identity.RecipientTypeDetails){
                'UserMailbox' { $CMDlet_Get= '(Get-mailbox -Identity $Identity.DistinguishedName).GrantSendOnBehalfTo';$CMDlet_Set='Set-Mailbox $Identity.DistinguishedName -GrantSendOnBehalfTo $SOBPArray'}
                'SharedMailbox' {$CMDlet_Get= '(Get-mailbox -Identity $Identity.DistinguishedName).GrantSendOnBehalfTo';$CMDlet_Set='Set-Mailbox $Identity.DistinguishedName -GrantSendOnBehalfTo $SOBPArray'}
                'RoomMailbox' {$CMDlet_Get= '(Get-mailbox -Identity $Identity.DistinguishedName).GrantSendOnBehalfTo';$CMDlet_Set='Set-Mailbox $Identity.DistinguishedName -GrantSendOnBehalfTo $SOBPArray'}
                'EquipmentMailbox' {$CMDlet_Get= '(Get-mailbox -Identity $Identity.DistinguishedName).GrantSendOnBehalfTo';$CMDlet_Set='Set-Mailbox $Identity.DistinguishedName -GrantSendOnBehalfTo $SOBPArray'}
                'LinkedMailbox' {$CMDlet_Get= '(Get-mailbox -Identity $Identity.DistinguishedName).GrantSendOnBehalfTo';$CMDlet_Set='Set-Mailbox $Identity.DistinguishedName -GrantSendOnBehalfTo $SOBPArray'}
                'MailUser' {$CMDlet_Get= '(Get-MailUser -Identity $Identity.DistinguishedName).GrantSendOnBehalfTo';$CMDlet_Set='Set-MailUser $Identity.DistinguishedName -GrantSendOnBehalfTo $SOBPArray'}
                'MailContact' {$CMDlet_Get= '(Get-MailContact -Identity $Identity.DistinguishedName).GrantSendOnBehalfTo';$CMDlet_Set='Set-MailContact $Identity.DistinguishedName -GrantSendOnBehalfTo $SOBPArray'}
                'MailNonUniversalGroup' {$CMDlet_Get= '(Get-DistributionGroup -Identity $Identity.DistinguishedName).GrantSendOnBehalfTo';$CMDlet_Set='Set-DistributionGroup $Identity.DistinguishedName -GrantSendOnBehalfTo $SOBPArray'}
                'MailUniversalDistributionGroup' {$CMDlet_Get= '(Get-DistributionGroup -Identity $Identity.DistinguishedName).GrantSendOnBehalfTo';$CMDlet_Set='Set-DistributionGroup $Identity.DistinguishedName -GrantSendOnBehalfTo $SOBPArray'}
                'MailUniversalSecurityGroup' {$CMDlet_Get= '(Get-DistributionGroup -Identity $Identity.DistinguishedName).GrantSendOnBehalfTo';$CMDlet_Set='Set-DistributionGroup $Identity.DistinguishedName -GrantSendOnBehalfTo $SOBPArray'}
                'PublicFolder' {$CMDlet_Get= '(Get-MailPublicFolder -Identity $Identity.DistinguishedName).GrantSendOnBehalfTo';$CMDlet_Set='Set-MailPublicFolder $Identity.DistinguishedName -GrantSendOnBehalfTo $SOBPArray'}
                Default{Write-Verbose "The Recipient Type of $($Identity.RecipientTypeDetails) does not meet the requirements to proceed"; Break}
        }
     
         If($CMDlet_Get){$SOBP =  Invoke-Expression $CMDlet_Get
         
            if($null -ne $SOBP){

                $SOBPArray = [System.Collections.ArrayList]@()
            
                foreach($SOBPobj in $SOBP){
                                        
                    #Null Var
                    $SOBPobj_USER = $Null
                    $SOBPobj_DEL = $Null
                    $SOBPobj_Action = 'Report Only'
                    
                    #Find User and Check for Orphanded SID or Object
                    $SOBPobj_USER = Find-User $SOBPobj.DistinguishedName
                    
                    $SOBPobj_DEL = Grant-PermissionRemoval -SamAccountName $SOBPobj_USER.SamAccountName -RecipientType $SOBPobj_USER.RecipientTypeDetails -Status $SOBPobj_USER.Status
                    
                    If($SOBPobj_DEL){$SOBPobjDeleteCounterTrue? += 1}
                    If($SOBPobj_DEL -eq $False){$SOBPobjDeleteCounterFalse? += 1}

                    Write-Verbose "FUNCTION Search-SendOnBehalfPermission: Found send-on-behalf permisison for $($SOBPobj.Name) on Source Recipient $($recipientObj.Name)"
                    
                    #Create new Send On Behalf Of list without Disconnected Mailboxes users
                    If(($SOBPobj_DEL)-and($PerformRemoval)){
                    
                    #This command fails to apply for disconnected mailboxes... Might suggest MS addnew value IsDisconnected like IsDeleted in class Microsoft.Exchange.Data.Directory.ADObjectId and action cleanup.
                    #Set-Mailbox $Identity.DistinguishedName -GrantSendOnBehalfTo @{Remove="$($SOBPobj.DistinguishedName)"}
                    
                        If($WhatIfPreference -ne $True){
                        
                            Write-Verbose "FUNCTION Search-SendOnBehalfPermission Recipient Permission: $($SOBPobj_USER.DisplayName) on Source Recipient: $($Identity.DisplayName) will be removed"
                            $SOBPobj_Action = "Successful Removal"
                        
                        }
                        
                        Else {
                        
                            Write-Verbose "FUNCTION Search-SendOnBehalfPermission -WHATIF Recipient Permission: $($SOBPobj_USER.DisplayName) on Source Recipient: $($Identity.DisplayName) will be removed"
                            $SOBPobj_Action = "Successful WhatIf"
                        
                        }
                    
                    }
                    Else{
                    
                    #Populate Array with real users
                    $SOBPArray += "$($SOBPobj.DistinguishedName)" 
                    
                    }

                    $SOBPUserReport = $SOBPUserReport + (New-ArrayObject -RecipientDisplayName $SOBPobj_USER.DisplayName -RecipientSamAcc $SOBPobj_USER.SamAccountName -RecipientType $SOBPobj_USER.RecipientTypeDetails -PermissionType "Send-on-Behalf" -SourceDisplayName $Identity.Name -SourceSamAcc $Identity.SamAccountName -SourceRecipientType $Identity.RecipientTypeDetails -Action $SOBPobj_Action -Removal $SOBPobj_DEL)
                    
                } 
                    
            If($PerformRemoval){
                
                #Check Array has all the members before actioning
                 If(($SOBPobjDeleteCounterTrue? -gt 0)-and($SOBPArray.count -eq $SOBPobjDeleteCounterFalse?)){
                        
                        Write-Verbose "FUNCTION Search-SendOnBehalfPermission: Updating Send On Behalf of Permisison for User $($Identity.DisplayName) Delete: $SOBPobjDeleteCounterTrue? Leave: $SOBPobjDeleteCounterFalse?"
                        
                        If($PSCmdlet.ShouldProcess("Delete: $SOBPobjDeleteCounterTrue? Leave: $SOBPobjDeleteCounterFalse?","Updating Send On Behalf of Permisison for user $($Identity.DisplayName)")){
                            
                            #Apply other objects ie Contacts and DL's
                            #Set-Mailbox $Identity.DistinguishedName -GrantSendOnBehalfTo $SOBPArray
                            
                            Try{
                            
                                Invoke-Expression $CMDlet_Set
                            
                            }
                            Catch{
                            
                                        Write-Verbose "FUNCTION Search-SendOnBehalfPermission: Failure CMDlet: Set-Mail* $_.Exception.Message";
                                        Write-Error "$_.Exception.Message"
                                        
                                        #Change Script Action Value to 'Failed Removal'
                                        ($SOBPUserReport) | % {If($_.'Script Action' -eq 'Successful Removal'){$_.'Script Action' = 'Failed Removal'}}
                
                            }
                            
                        }
                                
                 }
                        
            }
            
            #Add to main report for piped items
            $SOBPreport = $SOBPreport + $SOBPUserReport
            
            }
        }        
    }
    
    END {$SOBPreport}

    } #end function Search-SendOnBehalfPermission

Function Search-SendAsPermission
<#
. [AUTHOR]
. Joshua Bines, Consultant
.
. [DATE]
. 20180604
.
.DESCRIPTION
This Function searches for Send As Permissions, Reports, and removes permissions if the -PerformRemoval Switch is set. 
#>
{
 [CmdletBinding(SupportsShouldProcess=$True)]
 Param (
 
[Parameter(Mandatory = $True, ValueFromPipeline = $True)]
$Identity,

[Parameter(Mandatory = $False)]
[switch]$PerformRemoval
 
 )
 
 Begin {
    
    Write-Verbose "FUNCTION Search-SendAsPermission: Check for a user data entered into the 'Identity' Switch"
    
    if((($Identity.GetType()).name) -eq 'String'){
        
        Write-Verbose "FUNCTION Search-SendAsPermission: Confirmed User entered data of $Identity";
        Write-Verbose "FUNCTION Search-SendAsPermission: Attempting to resolve to a Recipient to $Identity";
    
            Try{
            
                $Identity = Get-Recipient $Identity -ErrorAction STOP | Select-Object -First 1 
            
            }
            Catch{
                        
                Write-Error "The Get-Recipient CMDlet returned a error and is unable to continue"; EXIT
                
            }
            
            if (($Identity | Measure-Object).count -gt 1){
            
                Write-Error "The Get-Recipient CMDlet returned more than one result after running the Get-Recipient CMDlet. Please use another switch for completing bulk actions";$Identity_STR_Error; EXIT
            
            }
            Else {Write-Verbose "FUNCTION Search-SendAsPermission: Recipient found"}
    }

 #Create PS Array
 [PSObject[]] $SENDASreport = @()
 
 }
    
 Process{
 
 $SENDAS = Get-ADPermission -Identity $Identity.DistinguishedName |  Where-Object {($_.ExtendedRights -like "*Send-As*") -and -not ($_.User -like "NT AUTHORITY\SELF")}
        
        if($null -ne $SENDAS){
            foreach($SENDASobj in $SENDAS){
                
                Write-Verbose "FUNCTION Search-SendAsPermission: Found Send As permisison for $SENDASobj.User on Source Recipient $($Identity.Name)"
                
                #Null Var
                $SENDASobj_USER = $Null
                $SENDASobj_DEL = $Null
                $SENDASobj_Action = 'Report Only'
                
                #Find User and Check for Orphanded SID or Object
                
                $SENDASobj_USER = Find-User $SENDASobj.User
                
                $SENDASobj_DEL = Grant-PermissionRemoval -SamAccountName $SENDASobj_USER.SamAccountName -RecipientType $SENDASobj_USER.RecipientTypeDetails -Status $SENDASobj_USER.Status
                                                                            
                if(($PerformRemoval) -and ($SENDASobj_DEL)) {
                        
                            If($PSCmdlet.ShouldProcess($($SENDASobj_USER.DisplayName),"Removing Send As permisison for user $($Identity.DisplayName)")){
                                                                    
                                Try{
                                    
                                    #Add Support for the -Confirm:$False Switch
                                    If(($PerformRemoval)-and($ConfirmPreference -eq 'None')){
                                    
                                        Remove-ADPermission -Identity $Identity.DistinguishedName -User $SENDASobj_USER.SamAccountName -ExtendedRights "Send As" -Confirm:$False
                                    
                                    }
                                    Else{
                                    
                                        Remove-ADPermission -Identity $Identity.DistinguishedName -User $SENDASobj_USER.SamAccountName -ExtendedRights "Send As"
                                        
                                    }
                                    
                                    If(($?)-and(-not $PSBoundParameters.ContainsKey('WhatIf'))){
                                    
                                        Write-Verbose "FUNCTION Search-SendAsPermission: Successful CMDlet: Remove-AdPermission  $($Identity.DisplayName) -User $($SENDASobj_USER.SamAccountName)"
                                        $SENDASobj_Action = "Successful Removal"
                                    
                                    }
                                
                                }
                                
                                Catch{
                                    
                                    Write-Verbose "FUNCTION Search-SendAsPermission: Failure CMDlet: Remove-ADPermission $_.Exception.Message";
                                    Write-Error "$_.Exception.Message"
                                    $SENDASobj_Action = "Failed Removal"
                                
                                }
                                 
                            }
                            
                    If(($WhatIfPreference -eq $True) -and ($SENDASobj_Action -ne 'Removal Failed')){
                    
                    Write-Verbose "FUNCTION Search-SendAsPermission: What If Successful CMDlet: Remove-ADPermission $($Identity.DisplayName) -User $($SENDASobj_USER.SamAccountName)"
                    $SENDASobj_Action = "Successful WhatIf"
                    
                    }

                }
                    
                $SENDASreport = $SENDASreport + (New-ArrayObject -RecipientDisplayName $SENDASobj_USER.DisplayName -RecipientSamAcc $SENDASobj_USER.SamAccountName -RecipientType $SENDASobj_USER.RecipientTypeDetails -PermissionType "Send-As" -SourceDisplayName $Identity.Name -SourceSamAcc $Identity.SamAccountName -SourceRecipientType $Identity.RecipientTypeDetails -Action $SENDASobj_Action -Removal $SENDASobj_DEL)

            }
        }
}

END {$SENDASreport}

} #end function Search-SendAsPermission

Function Search-ReceiveAsPermission
<#
. [AUTHOR]
. Joshua Bines, Consultant
.
. [DATE]
. 20180604
.
.DESCRIPTION
This Function searches for Receive As Permissions, Reports, and removes permissions if the -PerformRemoval Switch is set. 

    #>
{
 [CmdletBinding(SupportsShouldProcess=$True)]
 Param (
 
[Parameter(Mandatory = $True, ValueFromPipeline = $True)]
$Identity,

[Parameter(Mandatory = $False)]
[switch]$PerformRemoval
 
 )
 
 Begin {
    
    Write-Verbose "FUNCTION Search-ReceiveAsPermission: Check for a user data entered into the 'Identity' Switch"
    
    if((($Identity.GetType()).name) -eq 'String'){
        
        Write-Verbose "FUNCTION Search-ReceiveAsPermission: Confirmed User entered data of $Identity";
        Write-Verbose "FUNCTION Search-ReceiveAsPermission: Attempting to resolve to a Recipient to $Identity";
    
            Try{
            
                $Identity = Get-Recipient $Identity -ErrorAction STOP | Select-Object -First 1 
            
            }
            Catch{
                        
                Write-Error "The Get-Recipient CMDlet returned a error and is unable to continue"; EXIT
                
            }
            
            if (($Identity | Measure-Object).count -gt 1){
            
                Write-Error "The Get-Recipient CMDlet returned more than one result after running the Get-Recipient CMDlet. Please use another switch for completing bulk actions";$Identity_STR_Error; EXIT
            
            }
            Else {Write-Verbose "FUNCTION Search-ReceiveAsPermission: Recipient found"}
    }
 
 #Create PS Array
 [PSObject[]] $RECEIVEASreport = @()
 
 }
    
 Process{
 
 $RECEIVEAS = Get-ADPermission -Identity $Identity.DistinguishedName |  Where-Object {($_.ExtendedRights -like "*Receive-As*") -and -not ($_.User -like "NT AUTHORITY\SELF")}
        
        if($null -ne $RECEIVEAS){
            foreach($RECEIVEASobj in $RECEIVEAS){
                
                Write-Verbose "FUNCTION Search-ReceiveAsPermission: Found Send As permisison for $RECEIVEASobj.User on Source Recipient $($Identity.Name)"
                
                #Null Var
                $RECEIVEASobj_USER = $Null
                $RECEIVEASobj_DEL = $Null
                $RECEIVEASobj_Action = 'Report Only'
                
                #Find User and Check for Orphanded SID or Object
                
                $RECEIVEASobj_USER = Find-User $RECEIVEASobj.User
                
                $RECEIVEASobj_DEL = Grant-PermissionRemoval -SamAccountName $RECEIVEASobj_USER.SamAccountName -RecipientType $RECEIVEASobj_USER.RecipientTypeDetails -Status $RECEIVEASobj_USER.Status
                                                                            
                if(($PerformRemoval) -and ($RECEIVEASobj_DEL)) {
                        
                            If($PSCmdlet.ShouldProcess($($RECEIVEASobj_USER.DisplayName),"Removing Receive As permisison for user $($Identity.DisplayName)")){
                                                                    
                                Try{
                                    
                                    #Add Support for the -Confirm:$False Switch
                                    If(($PerformRemoval)-and($ConfirmPreference -eq 'None')){
                                    
                                        Remove-ADPermission -Identity $Identity.DistinguishedName -User $RECEIVEASobj_USER.SamAccountName -ExtendedRights "Receive As" -Confirm:$False
                                    
                                    }
                                    Else{
                                    
                                        Remove-ADPermission -Identity $Identity.DistinguishedName -User $RECEIVEASobj_USER.SamAccountName -ExtendedRights "Receive As"
                                        
                                    }
                                    
                                    If(($?)-and(-not $PSBoundParameters.ContainsKey('WhatIf'))){
                                    
                                        Write-Verbose "FUNCTION Search-ReceiveAsPermission: Successful CMDlet: Remove-AdPermission  $($Identity.DisplayName) -User $($RECEIVEASobj_USER.SamAccountName)"
                                        $RECEIVEASobj_Action = "Successful Removal"
                                    
                                    }
                                
                                }
                                
                                Catch{
                                    
                                    Write-Verbose "FUNCTION Search-ReceiveAsPermission: Failure CMDlet: Remove-ADPermission $_.Exception.Message";
                                    Write-Error "$_.Exception.Message"
                                    $RECEIVEASobj_Action = "Failed Removal"
                                
                                }
                                 
                            }
                            
                    If(($WhatIfPreference -eq $True) -and ($RECEIVEASobj_Action -ne 'Removal Failed')){
                    
                    Write-Verbose "FUNCTION Search-ReceiveAsPermission: What If Successful CMDlet: Remove-ADPermission $($Identity.DisplayName) -User $($RECEIVEASobj_USER.SamAccountName)"
                    $RECEIVEASobj_Action = "Successful WhatIf"
                    
                    }

                }
                    
                $RECEIVEASreport = $RECEIVEASreport + (New-ArrayObject -RecipientDisplayName $RECEIVEASobj_USER.DisplayName -RecipientSamAcc $RECEIVEASobj_USER.SamAccountName -RecipientType $RECEIVEASobj_USER.RecipientTypeDetails -PermissionType "Receive-As" -SourceDisplayName $Identity.Name -SourceSamAcc $Identity.SamAccountName -SourceRecipientType $Identity.RecipientTypeDetails -Action $RECEIVEASobj_Action -Removal $RECEIVEASobj_DEL)                    

            }
        }
}

END {$RECEIVEASreport}

} #end function Search-RECEIVEASPermission

Function Search-PublicDelegatesPermission
<#
. [AUTHOR]
. Joshua Bines, Consultant
.
. [DATE]
. 20180604
.
.DESCRIPTION
This Function searches for Public Delegates Permissions, Reports, and removes permissions if the -PerformRemoval Switch is set. 
    #>
{
 [CmdletBinding(SupportsShouldProcess=$True,ConfirmImpact="High")]
 Param (
 
[Parameter(Mandatory = $True, ValueFromPipeline = $True)]
$Identity,

[Parameter(Mandatory = $False)]
[switch]$PerformRemoval
 
 )
 
 Begin {
    
    Write-Verbose "FUNCTION Search-PublicDelegatesPermission: Check for a user data entered into the 'Identity' Switch"
    
    if((($Identity.GetType()).name) -eq 'String'){
        
        Write-Verbose "FUNCTION Search-PublicDelegatesPermission: Confirmed User entered data of $Identity";
        Write-Verbose "FUNCTION Search-PublicDelegatesPermission: Attempting to resolve to a Recipient to $Identity";
    
            Try{
            
                $Identity = Get-Recipient $Identity -ErrorAction STOP | Select-Object -First 1 
            
            }
            Catch{
                        
                Write-Error "The Get-Recipient CMDlet returned a error and is unable to continue"; EXIT
                
            }
            
            if (($Identity | Measure-Object).count -gt 1){
            
                Write-Error "The Get-Recipient CMDlet returned more than one result after running the Get-Recipient CMDlet. Please use another switch for completing bulk actions";$Identity_STR_Error; EXIT
            
            }
            Else {Write-Verbose "FUNCTION Search-PublicDelegatesPermission: Recipient found"}
    }
 #Create PS Array
 [PSObject[]] $PUBDELreport = @()
  
 }
    
 Process{
 
 $PUBDEL = (Get-ADUser -Identity $Identity.DistinguishedName -Properties publicDelegates).publicDelegates
 
        if($null -ne $PUBDEL){
            foreach($PUBDELobj in $PUBDEL){
                
                Write-Verbose "FUNCTION Search-PublicDelegatesPermission: Found Send As permisison for $PUBDELobj.User on Source Recipient $($Identity.Name)"
                
                #Null Var
                $PUBDELobj_USER = $Null
                $PUBDELobj_DEL = $Null
                $PUBDELobj_Action = 'Report Only'
                
                #Find User and Check for Orphanded SID or Object
                
                $PUBDELobj_USER = Find-User $PUBDELobj
                
                $PUBDELobj_DEL = Grant-PermissionRemoval -SamAccountName $PUBDELobj_USER.SamAccountName -RecipientType $PUBDELobj_USER.RecipientTypeDetails -Status $PUBDELobj_USER.Status
                                                                            
                if(($PerformRemoval) -and ($PUBDELobj_DEL)) {
                        
                            If($PSCmdlet.ShouldProcess($($PUBDELobj_USER.DisplayName),"Removing Public Delegates Permission permisison for user $($Identity.DisplayName)")){
                                                                    
                                Try{
                                    
                                    #Add Support for the -Confirm:$False Switch
                                    If(($PerformRemoval)-and($ConfirmPreference -eq 'None')){
                                    
                                        Set-ADUser $Identity.DistinguishedName -Remove @{PublicDelegates="$PUBDELobj"} -Confirm:$False
                                        
                                    }
                                    Else{
                                    
                                        Set-ADUser $Identity.DistinguishedName -Remove @{PublicDelegates="$PUBDELobj"}
                                        
                                    }
                                    
                                    If(($?)-and(-not $PSBoundParameters.ContainsKey('WhatIf'))){
                                    
                                        Write-Verbose "FUNCTION Search-PublicDelegatesPermission Successful CMDlet: Remove-AdPermission  $($Identity.DisplayName) -User $($PUBDELobj_USER.SamAccountName)"
                                        $PUBDELobj_Action = "Successful Removal"
                                    
                                    }
                                
                                }
                                
                                Catch{
                                    
                                    Write-Verbose "FUNCTION Search-PublicDelegatesPermission Failure CMDlet: Remove-ADPermission $_.Exception.Message";
                                    Write-Error "$_.Exception.Message"
                                    $PUBDELobj_Action = "Failed Removal"
                                
                                }
                                 
                            }
                            
                    If(($WhatIfPreference -eq $True) -and ($PUBDELobj_Action -ne 'Removal Failed')){
                    
                    Write-Verbose "FUNCTION Search-PublicDelegatesPermission What If Successful CMDlet: Remove-ADPermission $($Identity.DisplayName) -User $($PUBDELobj_USER.SamAccountName)"
                    $PUBDELobj_Action = "Successful WhatIf"
                    
                    }

                }
                    
                $PUBDELreport = $PUBDELreport + (New-ArrayObject -RecipientDisplayName $PUBDELobj_USER.DisplayName -RecipientSamAcc $PUBDELobj_USER.SamAccountName -RecipientType $PUBDELobj_USER.RecipientTypeDetails -PermissionType "Public Delegate" -SourceDisplayName $Identity.Name -SourceSamAcc $Identity.SamAccountName -SourceRecipientType $Identity.RecipientTypeDetails -Action $PUBDELobj_Action -Removal $PUBDELobj_DEL)                    

            }
        }
    

}

END {$PUBDELreport}

} #end function Search-PublicDelegatesPermission

Function Search-MailboxFolderPermission
<#
. [AUTHOR]
. Joshua Bines, Consultant
.
. [DATE]
. 20180607
.
.DESCRIPTION
This Function searches for Mailbox Folder Permissions, Reports, and removes permissions if the -PerformRemoval Switch is set. 

    #>
{
 [CmdletBinding(SupportsShouldProcess=$True)]
 Param (
 
[Parameter(Mandatory = $True, ValueFromPipeline = $True)]
$Identity,

[Parameter(Mandatory = $False)]
[switch]$PerformRemoval
 
 )
 
 Begin {
 
    Write-Verbose "FUNCTION Search-MailboxFolderPermission: Check for a user data entered into the 'Identity' Switch"
    
    if((($Identity.GetType()).name) -eq 'String'){
        
        Write-Verbose "FUNCTION Search-MailboxFolderPermission: Confirmed User entered data of $Identity";
        Write-Verbose "FUNCTION Search-MailboxFolderPermission: Attempting to resolve to a Recipient to $Identity";
    
            Try{
            
                $Identity = Get-Recipient $Identity -ErrorAction STOP | Select-Object -First 1 
            
            }
            Catch{
                        
                Write-Error "The Get-Recipient CMDlet returned a error and is unable to continue"; EXIT
                
            }
            
            if (($Identity | Measure-Object).count -gt 1){
            
                Write-Error "The Get-Recipient CMDlet returned more than one result after running the Get-Recipient CMDlet. Please use another switch for completing bulk actions";$Identity_STR_Error; EXIT
            
            }
            Else {Write-Verbose "FUNCTION Search-MailboxFolderPermission: Recipient found"}
    }
 
 #Create PS Array
 [PSObject[]] $MBXFoldersreport = @()
 
 }
    
 Process{

     #Create Blank Array for the Mailbox Folders. Null for each piped user to stop false postives.
     $MBXFOLArray = @();
     $MBXFoldersobj_PERMID = $Null


    If($null -ne $Identity){
            
            try{
            
                #Create Array 
                [string[]] $FolderPaths = Get-MailboxfolderStatistics "$($Identity.Guid)" | Where-Object{($_.FolderType -ne "RecoverableItemsRoot")-and($_.FolderType -ne "RecoverableItemsDeletions")-and($_.FolderType -ne "RecoverableItemsPurges")-and($_.Folderpath -ne "RecoverableItemsVersions")-and($_.FolderType -ne "SyncIssues")-and($_.FolderType -ne "Conflicts")-and($_.FolderType -ne "ServerFailures")-and($_.FolderType -ne "LocalFailures")-and($_.FolderType -ne "WorkingSet")-and($_.FolderType -ne "Audits")-and($_.FolderType -ne "CalendarLogging")} | %{$MBXFOLArray += (New-Object psobject -Property @{FolderPath=$_.FolderPath; FolderId=$_.FolderId})}
            
            }
            Catch{
            
                Write-Error "The Get-MailboxFolderStatistics CMDlet Returned an Error: $($_.Exception.Message)"
            }

            $MBXFolders = $MBXFOLArray
            foreach($MBXFoldersobj in $MBXFolders){
                    if($null -ne $MBXFoldersobj){

                    #Set Foreach Var to Null to Stop False Postives
                    $MBXFoldersobj_ID = $null
                    $MBXFoldersobj_Path = $null       
                    
                    #Add SamAccountName: for the Get-MailboxFolderPermission                   
                    $MBXFoldersobj_ID = "$($Identity.Guid)" + ":" + $MBXFoldersobj.FolderId
                    $MBXFoldersobj_Path = "$($Identity.Guid)" + ":" + $MBXFoldersobj.FolderPath

                    $MBXFOLPERM = Get-MailboxFolderPermission "$($MBXFoldersobj_ID)" 
                        foreach($MBXFOLPERMobj in $MBXFOLPERM){
                                
                                #Null Var
                                $MBXFoldersobj_USER = $Null
                                #$MBXFoldersobj_DEL = $Null
                                $MBXFoldersobj_Action = 'Report Only'

                                #Switch for Exchange 2010 vs 2013\2016 
                                if(-not($MBXFoldersobj_PERMID)){

                                    Switch ($MBXFOLPERMobj)
                                    {
                                        {$MBXFOLPERMobj.Identity.usertype -like "*"} { $MBXFoldersobj_PERMID = 'Identity'}
                                        {$MBXFOLPERMobj.User.usertype -like "*"} { $MBXFoldersobj_PERMID = 'User' }
                                        Default{Write-Error "Unable to determine Exchange Version. Could not match either '$MBXFOLPERMobj.Identity.usertype' or $MBXFOLPERMobj.User.usertype."}#; Break}
                                    }
                                    Write-Verbose "FUNCTION Search-MailboxFolderPermission: MBXFoldersobj_PERMID set to $MBXFoldersobj_PERMID"
                                
                                }#End If

                                #Declare Variables
                                #$MBXFOLPERMobj_ID_NAME = $MBXFOLPERMobj.$MBXFoldersobj_PERMID.displayname

                                #Ammended Objects Selected
                                #if(($MBXFOLPERMobj.identity.displayname -ne $recipientObj.Name) -and (($MBXFOLPERMobj.Identity.usertype -eq "Internal") -or ($MBXFOLPERMobj.Identity.usertype -eq "Unknown"))-and(($MBXFOLPERMobj.identity.displayname -ne 'Default')-or($MBXFOLPERMobj.identity.displayname -ne 'Anonymous'))){
                                
                                #Added support for 2013\2016
                                if(($MBXFOLPERMobj.$MBXFoldersobj_PERMID.displayname -ne $recipientObj.Name) -and (($MBXFOLPERMobj.$MBXFoldersobj_PERMID.usertype -eq "Internal") -or ($MBXFOLPERMobj.$MBXFoldersobj_PERMID.usertype -eq "Unknown"))-and(($MBXFOLPERMobj.$MBXFoldersobj_PERMID.displayname -ne 'Default')-or($MBXFOLPERMobj.$MBXFoldersobj_PERMID.UserType -ne 'Default')-or($MBXFOLPERMobj.$MBXFoldersobj_PERMID.displayname -ne 'Anonymous')-or($MBXFOLPERMobj.$MBXFoldersobj_PERMID.UserType -ne 'Anonymous'))){

                                    switch($MBXFOLPERMobj.$MBXFoldersobj_PERMID.usertype){
                                        'Internal' {
                                                                                                                                                                    
                                                            #Find User and Check for Orphanded SID or Object
                                                            $MBXFOLPERMobj_USER = Find-User $MBXFOLPERMobj.$MBXFoldersobj_PERMID.displayname
                                                            
                                                            $MBXFOLPERMobj_DEL = Grant-PermissionRemoval -SamAccountName $MBXFOLPERMobj_USER.SamAccountName -RecipientType $MBXFOLPERMobj_USER.RecipientTypeDetails -Status $MBXFOLPERMobj_USER.Status
                                                            
                                                            Write-Verbose "FUNCTION Search-MBXFolderPermission: Found Mailbox Folder permisison for $($MBXFOLPERMobj.User) on Source Recipient $($recipientObj.Name)"
                                                            $MBXFoldersreport = $MBXFoldersreport + (New-ArrayObject -RecipientDisplayName $MBXFOLPERMobj_USER.DisplayName -RecipientSamAcc $MBXFOLPERMobj_USER.SamAccountName -RecipientType $MBXFOLPERMobj_USER.RecipientTypeDetails -PermissionType "$($MBXFOLPERMobj.AccessRights) on Exchange Mailbox Folder $($MBXFoldersobj_Path)" -SourceDisplayName $Identity.Name -SourceSamAcc $Identity.SamAccountName -SourceRecipientType $Identity.RecipientTypeDetails -Action $MBXFoldersobj_Action -Removal $MBXFOLPERMobj_DEL)                    
                                                    
                                                    }
                                        'Unknown'  {
                                                                                                                                                                                        
                                                        #Find User and Check for Orphanded SID or Object
                                                        $MBXFOLPERMobj_USER = Find-User ($MBXFOLPERMobj.$MBXFoldersobj_PERMID.displayname -replace "NT User:")
                                                        
                                                        $MBXFOLPERMobj_DEL = Grant-PermissionRemoval -SamAccountName $MBXFOLPERMobj_USER.SamAccountName -RecipientType $MBXFOLPERMobj_USER.RecipientTypeDetails -Status $MBXFOLPERMobj_USER.Status
                                                            
                                                        Write-Verbose "FUNCTION Search-MailboxFolderPermission: Found Mailbox Folder permisison for $($MBXFOLPERMobj.User) on Source Recipient $($recipientObj.Name)"

                                                        if(($PerformRemoval) -and ($MBXFOLPERMobj_DEL)) {
                                                                
                                                                    If($PSCmdlet.ShouldProcess($MBXFoldersobj_PATH,"Removing mailbox folder permission for user $($MBXFOLPERMobj.$MBXFoldersobj_PERMID.displayname)")){
                                                                                                            
                                                                        Try{
                                                                            
                                                                            #Add Support for the -Confirm:$False Switch
                                                                            If(($PerformRemoval)-and($ConfirmPreference -eq 'None')){
                                                                            
                                                                                #Remove-ADPermission -Identity $Identity.DistinguishedName -User $MBXFoldersobj_USER.SamAccountName -ExtendedRights "Send As" -Confirm:$False
                                                                                Remove-MailboxFolderPermission $MBXFoldersobj_ID -User $MBXFOLPERMobj.$MBXFoldersobj_PERMID.displayname -Confirm:$false
                                                                            }
                                                                            Else{
                                                                            
                                                                                #Remove-ADPermission -Identity $Identity.DistinguishedName -User $MBXFoldersobj_USER.DistinguishedName -ExtendedRights "Send As"
                                                                                Remove-MailboxFolderPermission $MBXFoldersobj_ID -User $MBXFOLPERMobj.$MBXFoldersobj_PERMID.displayname #-Confirm:$false
                                                                            }
                                                                            
                                                                            If(($?)-and(-not $PSBoundParameters.ContainsKey('WhatIf'))){
                                                                            
                                                                                Write-Verbose "FUNCTION Search-MailboxFolderPermission Successful CMDlet: Remove-AdPermission  $($Identity.DisplayName) -User $($MBXFoldersobj_USER.SamAccountName)"
                                                                                $MBXFoldersobj_Action = "Successful Removal"
                                                                            
                                                                            }
                                                                        
                                                                        }
                                                                        
                                                                        Catch{
                                                                            
                                                                            Write-Verbose "FUNCTION Search-MailboxFolderPermission Failure CMDlet: Remove-ADPermission $_.Exception.Message";
                                                                            Write-Error "$_.Exception.Message"
                                                                            $MBXFoldersobj_Action = "Failed Removal"
                                                                        
                                                                        }
                                                                         
                                                                    }
                                                                    
                                                            If(($WhatIfPreference -eq $True) -and ($MBXFoldersobj_Action -ne 'Removal Failed')){
                                                            
                                                            Write-Verbose "FUNCTION Search-MailboxFolderPermission What If Successful CMDlet: Remove-ADPermission $($Identity.DisplayName) -User $($MBXFoldersobj_USER.SamAccountName)"
                                                            $MBXFoldersobj_Action = "Successful WhatIf"
                                                            
                                                            }
                                                                                                        

                                                }
                                                
                                                #Write Output to array for Identity.usertype 'Unknown'
                                                $MBXFoldersreport = $MBXFoldersreport + (New-ArrayObject -RecipientDisplayName $MBXFOLPERMobj_USER.DisplayName -RecipientSamAcc $MBXFOLPERMobj_USER.SamAccountName -RecipientType $MBXFOLPERMobj_USER.RecipientTypeDetails -PermissionType "$($MBXFOLPERMobj.AccessRights) on Exchange Mailbox Folder $($MBXFoldersobj_Path)" -SourceDisplayName $Identity.Name -SourceSamAcc $Identity.SamAccountName -SourceRecipientType $Identity.RecipientTypeDetails -Action $MBXFoldersobj_Action -Removal $MBXFOLPERMobj_DEL)       
                                                
                                            }
                                                
                                        Default{Write-Error "The Mailbox Folder Permission of 'Identity.usertype' is neither Unknown or Internal."}
                                    
                                    }#Switch End
                                    
                                }#If End
                        
                        }#Foreach End                    
               
                    }#If End
                
            }#Foreach End
    }
                                                        
}

END {$MBXFoldersreport}

} #end function Search-MailboxFolderPermission


Function Search-PublicFolderPermission
<#
. [AUTHOR]
. Joshua Bines, Consultant
.
. [DATE]
. 20180607
.
.DESCRIPTION
This Function searches for Public Folder Permissions, Reports, and removes permissions if the -PerformRemoval Switch is set. 
.
. Important! check for Administrators who have created Public Folders without mailboxes. These listed owners maybe removed depending on the variables set in the Grant-PermissionRemoval Function
#>
{
 [CmdletBinding(SupportsShouldProcess=$True,ConfirmImpact="High")]
 Param (
 
[Parameter(Mandatory = $True, ValueFromPipeline = $True)]
$Identity,

[Parameter(Mandatory = $False)]
[switch]$PerformRemoval
 
 )
 
 Begin {
    
    Write-Verbose "FUNCTION Search-PublicFolderPermission: Check for a user data entered into the 'Identity' Switch"
    
    if((($Identity.GetType()).name) -eq 'String'){
        
        Write-Verbose "FUNCTION Search-PublicFolderPermission: Confirmed User entered data of $Identity";
        Write-Verbose "FUNCTION Search-PublicFolderPermission: Attempting to resolve to a Recipient to $Identity";
    
            Try{
            
                $Identity = Get-Recipient $Identity -ErrorAction STOP | Select-Object -First 1 
            
            }
            Catch{
                        
                Write-Error "The Get-Recipient CMDlet returned a error and is unable to continue"; EXIT
                
            }
            
            if (($Identity | Measure-Object).count -gt 1){
            
                Write-Error "The Get-Recipient CMDlet returned more than one result after running the Get-Recipient CMDlet. Please use another switch for completing bulk actions";$Identity_STR_Error; EXIT
            
            }
            Else {Write-Verbose "FUNCTION Search-PublicFolderPermission: Recipient found"}
    }
 #Create PS Array
 [PSObject[]] $PFCPreport = @()
  
 }
    
 Process{
        $PF =$Null
        #Check for get-PublicFolder input from the pipeline
        If($Identity.EntryId){$PF = Get-PublicFolder $Identity.EntryId -ErrorAction Continue }
        Else{
            $PF = Get-MailPublicFolder $Identity.DistinguishedName | ForEach-Object{If($_.EntryId){Get-PublicFolder $_.EntryId}}
            #Exchange 2010 Work Around
            if(-not $PF){$PF = Get-MailPublicFolder $Identity.DistinguishedName | Get-PublicFolder}
        }

        #Support for Exchange 2010
        If($PF.MapiIdentity){
            Try{
                #$PF = Get-MailPublicFolder $Identity.DistinguishedName | Get-PublicFolder
                $PFCP = $PF | Get-PublicFolderClientPermission | Where-Object {($_.user.IsDefault -eq $false) -and ($_.user.IsAnonymous -eq $false)}
            }
            Catch{
                Write-Verbose "FUNCTION Search-PublicFolderPermission Failure CMDlet: Get-MailPublicFolder $($_.Exception.Message)";
                #Write-Error "$_.Exception.Message"
            }
        }

        #Check for Exchange 2013. 
        Else{
            Try{
                $PFCP = Get-PublicFolderClientPermission $PF.EntryId | Where-Object {($($_.User.UserType) -eq  "Internal")-xor ($_.user.UserType -eq "Unknown")}
            }
            Catch{
                Write-Verbose "FUNCTION Search-PublicFolderPermission Failure CMDlet: Get-MailPublicFolder$($_.Exception.Message)";
                #Write-Error "$_.Exception.Message"
            }
        }

        #$PF = Get-mailPublicFolder $recipientObj.alais | Get-PublicFolder
            
            if($null -ne $PFCP){
                foreach($PFCPobj in $PFCP){
                    
                    #Null Var
                    $PFCPobj_USER = $Null
                    $PFCPobj_DEL = $Null
                    $PFCPobj_Action = 'Report Only'
                    
                    #Find User and Check for Orphanded SID or Object
                    $PFCPobj_USER = Find-User ($PFCPobj.User -replace "NT User:")
                    
                    $PFCPobj_DEL = Grant-PermissionRemoval -SamAccountName $PFCPobj_USER.SamAccountName -RecipientType $PFCPobj_USER.RecipientTypeDetails -Status $PFCPobj_USER.Status
                                      
                    if(($PerformRemoval) -and ($PFCPobj_DEL)) {
                            
                                If($PSCmdlet.ShouldProcess($($PFCPobj_USER.DisplayName),"Removing Public Folder Permission permisison for $($Identity.DisplayName)")){
                                                                        
                                    Try{
                                        
                                        #Add Support for the -Confirm:$False Switch
                                        If(($PerformRemoval)-and($ConfirmPreference -eq 'None')){
                                        

                                            #Check for Exchange 2013. 
                                            If($PF.EntryId){

                                                Remove-PublicFolderClientPermission -Identity $PF.EntryId -User $PFCPobj.User -Confirm:$False
        
                                            }
                                            #Support for Exchange 2010
                                            Else{
                                                Remove-PublicFolderClientPermission -Identity $PF.MapiIdentity.tostring() -User $PFCPobj.User -AccessRights $PFCPobj.AccessRights -Confirm:$False
                                            }
                                            
                                        }
                                        Else{
                                        
                                            #Check for Exchange 2013. 
                                            If($PF.EntryId){

                                                Remove-PublicFolderClientPermission -Identity $PF.EntryId -User $PFCPobj.User

        
                                            }
                                            #Support for Exchange 2010
                                            Else{
                                                Remove-PublicFolderClientPermission -Identity $PF.MapiIdentity.tostring() -User $PFCPobj.User -AccessRights $PFCPobj.AccessRights
                                            }

                                            
                                        }
                                        
                                        If(($?)-and(-not $PSBoundParameters.ContainsKey('WhatIf'))){
                                        
                                            Write-Verbose "FUNCTION Search-PublicFolderPermission Successful CMDlet: Remove-AdPermission  $($Identity.DisplayName) -User $($PFCPobj_USER.SamAccountName)"
                                            $PFCPobj_Action = "Successful Removal"
                                        
                                        }
                                    
                                    }
                                    
                                    Catch{
                                        
                                        Write-Verbose "FUNCTION Search-PublicFolderPermission Failure CMDlet: Remove-ADPermission $($_.Exception.Message)";
                                        Write-Error "$_.Exception.Message"
                                        $PFCPobj_Action = "Failed Removal"
                                    
                                    }
                                     
                                }
                                
                        If(($WhatIfPreference -eq $True) -and ($PFCPobj_Action -ne 'Removal Failed')){
                        
                        Write-Verbose "FUNCTION Search-PublicFolderPermission What If Successful CMDlet: Remove-ADPermission $($Identity.DisplayName) -User $($PFCPobj_USER.SamAccountName)"
                        $PFCPobj_Action = "Successful WhatIf"
                        
                        }

                    }
                
                $PFCPreport = $PFCPreport + (New-ArrayObject -RecipientDisplayName $PFCPobj_USER.DisplayName -RecipientSamAcc $PFCPobj_USER.SamAccountName -RecipientType $PFCPobj_USER.RecipientTypeDetails -PermissionType "$($PFCPobj.AccessRights) on Public Folder Permission $($PF.Identity)" -SourceDisplayName $Identity.Name -SourceSamAcc $Identity.SamAccountName -SourceRecipientType $Identity.RecipientTypeDetails -Action $PFCPobj_Action -Removal $PFCPobj_DEL)                    
                    
                }
            }
}

END {$PFCPreport}

} #end function Search-PublicFolderPermission

Function Search-RecipientForwarding
<#
. [AUTHOR]
. Joshua Bines, Consultant
.
. [DATE]
. 20180612
.
.DESCRIPTION
This Function searches for Recipient Forwarding Permissions and Reports. The Perform Removal Switch has been removed from this function.

    #>
{
 [CmdletBinding(SupportsShouldProcess=$True)]
 Param (
 
[Parameter(Mandatory = $True, ValueFromPipeline = $True)]
$Identity,

[Parameter(Mandatory = $False)]
[switch]$PerformRemoval

 )
 
 Begin {

Write-Verbose "FUNCTION Search-RecipientForwarding: Check for a user data entered into the 'Identity' Switch"

if((($Identity.GetType()).name) -eq 'String'){
    
    Write-Verbose "FUNCTION Search-RecipientForwarding: Confirmed User entered data of $Identity";
    Write-Verbose "FUNCTION Search-RecipientForwarding: Attempting to resolve to a Recipient to $Identity";
    
    Try{
    
        $Identity = Get-Recipient $Identity -ErrorAction STOP | Select-Object -First 1 
    
    }
    Catch{
                
        Write-Error "The Get-Recipient CMDlet returned a error and is unable to continue"; EXIT
        
    }
        
        if (($Identity | Measure-Object).count -gt 1){
        
            Write-Error "The Get-Recipient CMDlet returned more than one result after running the Get-Recipient CMDlet. Please use another switch for completing bulk actions";$Identity_STR_Error; EXIT
        
        }
        Else {Write-Verbose "FUNCTION Search-RecipientForwarding: Recipient found"}
}
 #Create PS Array
 [PSObject[]] $FORWARDreport = @()
 #$ScriptAction? = "Script Action"
      
 }
    
 Process{
 
 #Create Local User Array
 [PSObject[]] $FORWARDUserReport = @()

 #
    switch ($Identity.RecipientTypeDetails){
            'UserMailbox' { $CMDlet_Get= '(Get-mailbox -Identity $Identity.DistinguishedName)'}
            'SharedMailbox' {$CMDlet_Get= '(Get-mailbox -Identity $Identity.DistinguishedName)'}
            'RoomMailbox' {$CMDlet_Get= '(Get-mailbox -Identity $Identity.DistinguishedName)'}
            'EquipmentMailbox' {$CMDlet_Get= '(Get-mailbox -Identity $Identity.DistinguishedName)'}
            'LinkedMailbox' {$CMDlet_Get= '(Get-mailbox -Identity $Identity.DistinguishedName)'}
            'PublicFolder' {$CMDlet_Get= '(Get-MailPublicFolder -Identity $Identity.DistinguishedName)'}
            Default{Write-Verbose "The Recipient Type of $($Identity.RecipientTypeDetails) does not meet the requirements to proceed"}#; break}
    }
 
     If($CMDlet_Get){
     
     $FORWARD =  Invoke-Expression $CMDlet_Get
     
        if(($FORWARD.ForwardingAddress)-or($FORWARD.ForwardingSmtpAddress)){
        
            foreach($FORWARDobj in $FORWARD){
                                    
                #Null Var
                $FORWARDobj_USER = $Null
                $FORWARDobj_DEL = $False
                $FORWARDobj_Action = 'Report Only'

                If($FORWARD.ForwardingSmtpAddress){
                    #Find User and Check for Orphanded SID or Object
                    $FORWARDobj_USER = Find-User $FORWARDobj.ForwardingSmtpAddress.AddressString

                    $FORWARDUserReport = $FORWARDUserReport + (New-ArrayObject -RecipientDisplayName $FORWARDobj_USER.DisplayName -RecipientSamAcc $FORWARDobj_USER.SamAccountName -RecipientType $FORWARDobj_USER.RecipientTypeDetails -PermissionType "Forwarding-Smtp-Address" -SourceDisplayName $Identity.Name -SourceSamAcc $Identity.SamAccountName -SourceRecipientType $Identity.RecipientTypeDetails -Action $FORWARDobj_Action -Removal $FORWARDobj_DEL)
                    
                    
                }

                If($FORWARD.ForwardingAddress){

                #Find User and Check for Orphanded SID or Object
                $FORWARDobj_USER = Find-User $FORWARDobj.ForwardingAddress.DistinguishedName
                
                Write-Verbose "FUNCTION Search-RecipientForwarding: Found Recipient-Forwarding permisison for $($FORWARDobj.Name) on Source Recipient $($recipientObj.Name)"

                $FORWARDUserReport = $FORWARDUserReport + (New-ArrayObject -RecipientDisplayName $FORWARDobj_USER.DisplayName -RecipientSamAcc $FORWARDobj_USER.SamAccountName -RecipientType $FORWARDobj_USER.RecipientTypeDetails -PermissionType "Forwarding-Address" -SourceDisplayName $Identity.Name -SourceSamAcc $Identity.SamAccountName -SourceRecipientType $Identity.RecipientTypeDetails -Action $FORWARDobj_Action -Removal $FORWARDobj_DEL)

                }
                                
                
            } 
        
        #Add to main report for piped items
        $FORWARDreport = $FORWARDreport + $FORWARDUserReport
        
        }
    }
}

END {$FORWARDreport}

} #end function Search-RecipientForwarding

    #Load Modules and complete a prep check break from script if fails
     
     if (!(Get-Module | Where-Object {$_.Name -eq "ActiveDirectory"})) 
     {
        Write-Verbose 'Loading the Active Directory Module'
        try{
        Import-module ActiveDirectory
        }
        Catch{
            if($?){
                
                Write-Verbose 'AD Module fired up! Lets try Exchange'
                
                } 
            
            Else {Write-Warning $_.Exception.Message; Write-Verbose 'AD Module Failed To Load`r`n'; EXIT}
        }

     }
    Else {Write-Verbose "Active Directory Module is already loaded!`r`n"}
     
    #Add Exchange 2010 snapin if not already loaded in the PowerShell session

    <#
    Write-Verbose 'Checking Exchange Snapin is loaded'
    if (!(Get-PSSnapin | Where-Object {$_.Name -eq "Microsoft.Exchange.Management.PowerShell.E2010"}))
    {
        Write-Verbose 'Loading the Exchange Server PowerShell snapin'
        try
        {
            Add-PSSnapin Microsoft.Exchange.Management.PowerShell.E2010 -ErrorAction STOP
        }
        catch
        {
            #Snapin was not loaded
            Write-Verbose "The Exchange Server PowerShell snapin did not load.`r`n"
            Write-Warning $_.Exception.Message
            EXIT
        }
        . $env:ExchangeInstallPath\bin\RemoteExchange.ps1
        Connect-ExchangeServer -auto -AllowClobber
    }
    Else {Write-Verbose "Exchange Snapin is already loaded!"}

    #>

    #Test and Data Validation $Identity Input
    Write-Verbose "Testing Local User $($env:UserName) RBAC Permission"    
    If(Test-CommandExists Get-Recipient,Get-Mailbox,Get-MailboxPermission,Get-ADPermission,Get-ADUser,Get-MailboxFolderStatistics,Get-MailUser,Get-MailPublicFolder,Get-MailboxFolderPermission){
    
        Write-Verbose "Correct RBAC Access Confirmed"
    } 
        
    Else {Write-Error "Script requires a higher level of access. Please Contact IT Support"; EXIT}
    
    if($PerformRemoval){
        Write-Verbose "Testing Local User $($env:UserName) RBAC Permission for Switch PerformRemoval"    
        If(Test-CommandExists Remove-MailboxFolderPermission, Remove-MailboxPermission,Set-Mailbox,Set-MailUser,Set-MailContact,Set-DistributionGroup,Set-MailPublicFolder){
        
            Write-Verbose "Correct RBAC Access Confirmed For Switch PerformRemoval"
        } 
            
        Else {Write-Error "Switch PerformRemoval requires a higher level of access. Please Contact IT Support"; EXIT}
    }

    #Precheck the export file path
    if($exportpath){
        Write-Verbose "Testing Export File Path"
        $ExportPathResult = Test-Path $exportPath
            If($ExportPathResult){
            
            Write-Verbose "Export Path of $exportPath Tested Successfully"
            #Add a Trailing \
            $exportPath = "$($exportPath)\"
            
            }
            Else {
            
            Write-Error "Export File Path Value incorrect. Please enter a valid path the runtime user $($env:UserName) can access"; EXIT
            
            }
        
    }
    
    #Check for a string and attempt to resolve to a Recipient
    if($null -ne $Identity){
        Write-Verbose "Check for a user data entered into the 'Identity' Switch"
        if((($Identity.GetType()).name) -eq 'String'){
            Write-Verbose "Confirmed User entered data of $Identity";
            Write-Verbose "Attempting to resolve to a Recipient to $Identity";
            
            Try{
            
                $Identity = Get-Recipient $Identity -ErrorAction STOP | Select-Object -First 1 
            
            }
            Catch{
                        
                Write-Error "The Get-Recipient CMDlet returned a error and is unable to continue"; EXIT
                
            }
                
            if (($Identity | Measure-Object).count -gt 1){
            
            Write-Error "The Get-Recipient CMDlet returned more than one result after running the Get-Recipient CMDlet. Please use another switch for completing bulk actions";$Identity_STR_Error; EXIT
            
            }
            Else {Write-Verbose "Recipient found"}
        }
        
        #Check and Confirm object is an array
        If(($Identity.GetType()).basetype.name -eq 'Array'){
        
        Write-Verbose "Array found instead of String"
            
        }
    }

    #Create Blank Array for the Report
    [PSObject[]] $report = @();
    
    
}#End_Begin

Process{
    
    $recipientObj = $Identity
    
    #Null out looped values
    $CMDlet_FMP = $null
    $CMDlet_SOBP= $null
    $CMDlet_SENDAS = $null
    $CMDlet_RECEIVEAS=$null
    $CMDlet_PUBDEL=$null
    $CMDlet_FORW= $null
    $CMDlet_MBXFOL=$null
    $CMDlet_PF=$null
    
    #Create Blank Array for the Mailbox Folders. Null for each piped user to stop false postives.
    $MBXFOLArray = @();


    #Identify and define CMDLETS
    switch ($recipientObj.RecipientTypeDetails){
            'UserMailbox' { 
                            $CMDlet_FMP = $True
                            $CMDlet_SOBP = $True
                            $CMDlet_SENDAS = $True
                            $CMDlet_RECEIVEAS= $True
                            $CMDlet_PUBDEL= $True
                            $CMDlet_FORW= $True
                            $CMDlet_MBXFOL= $True}
            'SharedMailbox' {
                            $CMDlet_FMP = $True
                            $CMDlet_SOBP = $True
                            $CMDlet_SENDAS = $True
                            $CMDlet_RECEIVEAS= $True
                            $CMDlet_PUBDEL= $True
                            $CMDlet_FORW= $True
                            $CMDlet_MBXFOL= $True}
            'RoomMailbox' {
                            $CMDlet_FMP = $True
                            $CMDlet_SOBP = $True
                            $CMDlet_SENDAS = $True
                            $CMDlet_RECEIVEAS= $True
                            $CMDlet_PUBDEL= $True
                            $CMDlet_FORW= $True
                            $CMDlet_MBXFOL= $True}
            'EquipmentMailbox' {
                            $CMDlet_FMP = $True
                            $CMDlet_SOBP = $True
                            $CMDlet_SENDAS = $True
                            $CMDlet_RECEIVEAS= $True
                            $CMDlet_PUBDEL= $True
                            $CMDlet_FORW= $True
                            $CMDlet_MBXFOL= $True}
            'LinkedMailbox' {
                            $CMDlet_FMP = $True
                            $CMDlet_SOBP = $True
                            $CMDlet_SENDAS = $True
                            $CMDlet_RECEIVEAS= $True
                            $CMDlet_PUBDEL= $True
                            $CMDlet_FORW= $True
                            $CMDlet_MBXFOL= $True}
            'MailUser' {
                            $CMDlet_SOBP = $True
                            $CMDlet_SENDAS = $True
                            $CMDlet_FORW= $True
                            $CMDlet_RECEIVEAS= $True}
            'MailContact' {
                            $CMDlet_SOBP = $True
                            $CMDlet_SENDAS = $True
                            $CMDlet_RECEIVEAS=$True}
            'MailNonUniversalGroup' {
                            $CMDlet_SOBP = $True
                            $CMDlet_SENDAS = $True}
            'MailUniversalDistributionGroup' {
                            $CMDlet_SOBP = $True
                            $CMDlet_SENDAS = $True}
            'MailUniversalSecurityGroup' {
                            $CMDlet_SOBP = $True
                            $CMDlet_SENDAS = $True}
            'PublicFolder' {
                            $CMDlet_SOBP = $True
                            $CMDlet_SENDAS = $True
                            $CMDlet_FORW= $True
                            $CMDlet_PF = $True}
            Default{Write-Error "The Object Recipient Type of $($recipientObj.RecipientTypeDetails) is not accepted"; Write-Verbose "The Recipient Type of $RecipientType does not meet the requirements to proceed"; Break}
        }

        #Script block to get Full mailbox permisisons
        if ($CMDlet_FMP){
            
            Try{

                If(($PerformRemoval)-and($ConfirmPreference -eq 'None')){$FMP = Search-FullMailboxPermission $recipientObj -PerformRemoval -Confirm:$False -ErrorAction Continue}
                ElseIf(($PerformRemoval)-and($WhatIfPreference -eq $True)){$FMP = Search-FullMailboxPermission $recipientObj -PerformRemoval -WhatIf -ErrorAction Continue}
                ElseIf($PerformRemoval){$FMP = Search-FullMailboxPermission $recipientObj -PerformRemoval -ErrorAction Continue}
                Else{$FMP = Search-FullMailboxPermission $recipientObj -ErrorAction Continue}
            
            }
            Catch{
            
                Write-Error "Failed to call function Search-FullMailboxPermission: $_.Exception.Message"
            
            }
            Finally{
            
                If($FMP){
                
                $FMP | Select-Object 'SourceRecipient','PermissionType',Recipient,'ScriptAction' #Limit Results to allow FT by default
                $report +=  $FMP
                
                }
            
            }
            
        }
        
        Else{Write-Verbose "Skipping Full mailbox permission for Object $($($recipientObj).name) switch missing variable 'CMDlet_FMP'"}

        #Script block to get send-on-behalf permissions
        if($CMDlet_SOBP){
        
            Try{

                If(($PerformRemoval)-and($ConfirmPreference -eq 'None')){$SOBP = Search-SendOnBehalfPermission -Identity $recipientObj -PerformRemoval -Confirm:$False -ErrorAction Continue}
                ElseIf(($PerformRemoval)-and($WhatIfPreference -eq $True)){$SOBP = Search-SendOnBehalfPermission -Identity $recipientObj -PerformRemoval -WhatIf -ErrorAction Continue}
                ElseIf($PerformRemoval){$SOBP = Search-SendOnBehalfPermission -Identity $recipientObj -PerformRemoval -ErrorAction Continue}
                Else{$SOBP = Search-SendOnBehalfPermission -Identity $recipientObj -ErrorAction Continue}
            
            }
            Catch{
            
                Write-Error "Failed to call function Search-SendOnBehalfPermission: $_.Exception.Message"
            
            }
            Finally{
            
                If($SOBP){
                
                $SOBP | Select-Object 'SourceRecipient','PermissionType',Recipient,'ScriptAction' #Limit Results to allow FT by default
                $report +=  $SOBP
                
                }
            
            }

        }
        
        Else{Write-Verbose "Skipping send-on-behalf permissions for Object $($($recipientObj).name) switch missing variable 'CMDlet_SOBP'"}
        
        #Script block to get Send As 
        if($CMDlet_SENDAS){
                        
            Try{

                If(($PerformRemoval)-and($ConfirmPreference -eq 'None')){$SENDAS = Search-SendAsPermission -Identity $recipientObj -PerformRemoval -Confirm:$False -ErrorAction Continue}
                ElseIf(($PerformRemoval)-and($WhatIfPreference -eq $True)){$SENDAS = Search-SendAsPermission -Identity $recipientObj -PerformRemoval -WhatIf -ErrorAction Continue}
                ElseIf($PerformRemoval){$SENDAS = Search-SendAsPermission -Identity $recipientObj -PerformRemoval -ErrorAction Continue}
                Else{$SENDAS = Search-SendAsPermission -Identity $recipientObj -ErrorAction Continue}
            
            }
            Catch{
            
                Write-Error "Failed to call function Search-SendOnBehalfPermission: $_.Exception.Message"
            
            }
            Finally{
                
                If($SENDAS){
                
                $SENDAS | Select 'SourceRecipient','PermissionType',Recipient,'ScriptAction' #Limit Results to allow FT by default
                $report +=  $SENDAS
                
                }
            
            }

        }
        
        Else{Write-Verbose "Skipping Send As permissions for Object $($($recipientObj).name) switch missing variable 'CMDlet_SENDAS'"}
        
        #Script block to get Receive As
        if($CMDlet_RECEIVEAS){
            
            Try{

                If(($PerformRemoval)-and($ConfirmPreference -eq 'None')){$RECEIVEAS = Search-ReceiveAsPermission -Identity $recipientObj -PerformRemoval -Confirm:$False -ErrorAction Continue}
                ElseIf(($PerformRemoval)-and($WhatIfPreference -eq $True)){$RECEIVEAS = Search-ReceiveAsPermission -Identity $recipientObj -PerformRemoval -WhatIf -ErrorAction Continue}
                ElseIf($PerformRemoval){$RECEIVEAS = Search-ReceiveAsPermission -Identity $recipientObj -PerformRemoval -ErrorAction Continue}
                Else{$RECEIVEAS = Search-ReceiveAsPermission -Identity $recipientObj -ErrorAction Continue}
            
            }
            Catch{
            
                Write-Error "Failed to call function Search-ReceiveAsPermission: $_.Exception.Message"
            
            }
            Finally{
            
                If($RECEIVEAS){
                
                $RECEIVEAS | Select-Object 'SourceRecipient','PermissionType',Recipient,'ScriptAction' #Limit Results to allow FT by default
                $report +=  $RECEIVEAS
                
                }
            
            }
        
        }
        
       Else{Write-Verbose "Skipping Receive As permissions for Object $($($recipientObj).name) switch missing variable 'CMDlet_RECEIVEAS'"}
        
        #Script block to get Public Delegates
        if($CMDlet_PUBDEL){
            
            Try{

                If(($PerformRemoval)-and($ConfirmPreference -eq 'None')){$PUBDEL = Search-PublicDelegatesPermission -Identity $recipientObj -PerformRemoval -Confirm:$False -ErrorAction Continue}
                ElseIf(($PerformRemoval)-and($WhatIfPreference -eq $True)){$PUBDEL = Search-PublicDelegatesPermission -Identity $recipientObj -PerformRemoval -WhatIf -ErrorAction Continue}
                ElseIf($PerformRemoval){$PUBDEL = Search-PublicDelegatesPermission -Identity $recipientObj -PerformRemoval -ErrorAction Continue}
                Else{$PUBDEL = Search-PublicDelegatesPermission -Identity $recipientObj -ErrorAction Continue}
            
            }
            Catch{
            
                Write-Error "Failed to call function Search-ReceiveAsPermission: $_.Exception.Message"
            
            }
            Finally{
            
                If($PUBDEL){
                
                $PUBDEL | Select-Object 'SourceRecipient','PermissionType',Recipient,'ScriptAction' #Limit Results to allow FT by default
                $report +=  $PUBDEL
                
                }
            
            }
        
        }

        Else{Write-Verbose "Skipping Public Delegates permissions for Object $($($recipientObj).name) switch missing variable 'CMDlet_PUBDEL'"}

        #Script block to get MailBox Folder Permissions
        if($CMDlet_MBXFOL){

            Try{

                If(($PerformRemoval)-and($ConfirmPreference -eq 'None')){$MBXFOL = Search-MailboxFolderPermission -Identity $recipientObj -PerformRemoval -Confirm:$False -ErrorAction Continue}
                ElseIf(($PerformRemoval)-and($WhatIfPreference -eq $True)){$MBXFOL = Search-MailboxFolderPermission -Identity $recipientObj -PerformRemoval -WhatIf -ErrorAction Continue}
                ElseIf($PerformRemoval){$MBXFOL = Search-MailboxFolderPermission -Identity $recipientObj -PerformRemoval -ErrorAction Continue}
                Else{$MBXFOL = Search-MailboxFolderPermission -Identity $recipientObj -ErrorAction Continue}
            
            }
            Catch{
            
                Write-Error "Failed to call function Search-MailboxFolderPermission: $_.Exception.Message"
            
            }
            Finally{
            
                If($MBXFOL){
                
                $MBXFOL | Select-Object 'SourceRecipient','PermissionType',Recipient,'ScriptAction' #Limit Results to allow FT by default
                $report +=  $MBXFOL
                
                }
            
            }
            

        }
        
        Else{Write-Verbose "Skipping  MailBox Folder Permissions for Object $($($recipientObj).name) switch missing variable 'CMDlet_MBXFOL'"}

        if ($CMDlet_PF){
        
            Try{

                If(($PerformRemoval)-and($ConfirmPreference -eq 'None')){$PF = Search-PublicFolderPermission -Identity $recipientObj -PerformRemoval -Confirm:$False -ErrorAction Continue}
                ElseIf(($PerformRemoval)-and($WhatIfPreference -eq $True)){$PF = Search-PublicFolderPermission -Identity $recipientObj -PerformRemoval -WhatIf -ErrorAction Continue}
                ElseIf($PerformRemoval){$PF = Search-PublicFolderPermission -Identity $recipientObj -PerformRemoval -ErrorAction Continue}
                Else{$PF = Search-PublicFolderPermission -Identity $recipientObj -ErrorAction Continue}
            
            }
            Catch{
            
                Write-Error "Failed to call function Search-PublicFolderPermission: $_.Exception.Message"
            
            }
            Finally{
            
                If($PF){
                
                $PF | Select-Object 'SourceRecipient','PermissionType',Recipient,'ScriptAction' #Limit Results to allow FT by default
                $report +=  $PF
                
                }
            
            }
            
        }
        
        Else{Write-Verbose "Skipping Public Folder Permissions for Object $($($recipientObj).name) switch missing variable 'CMDlet_PF'"}
        
        if ($CMDlet_FORW){
        
            Try{
            
                If(($PerformRemoval)-and($ConfirmPreference -eq 'None')){$FORW = Search-RecipientForwarding -Identity $recipientObj -ErrorAction Continue}
                ElseIf(($PerformRemoval)-and($WhatIfPreference -eq $True)){$FORW = Search-RecipientForwarding -Identity $recipientObj -ErrorAction Continue}
                ElseIf($PerformRemoval){$FORW = Search-RecipientForwarding -Identity $recipientObj -PerformRemoval -ErrorAction Continue}
                Else{$FORW = Search-RecipientForwarding -Identity $recipientObj -ErrorAction Continue}
            
            }
            Catch{
            
                Write-Error "Failed to call function Search-RecipientForwarding: $_.Exception.Message"
            
            }
            Finally{
            
                If($FORW){
                
                $FORW | Select-Object 'SourceRecipient','PermissionType',Recipient,'ScriptAction' #Limit Results to allow FT by default
                $report +=  $FORW
                
                }
            
            }
            
        }
        
        Else{Write-Verbose "Skipping Recipient Forwarding for Object $($($recipientObj).name) switch missing variable 'CMDlet_FORW'"}


    # Reapply Default ErrorActionPreference Value Also at the end if the Function Find-User fails
    if($oldActionPreference -ne $ErrorActionPreference){
            
        Set-Variable -Name ErrorActionPreference -Value $oldActionPreference -Scope Global
        #$ErrorActionPreference = $ErrorActionPreferenceChanged
        If($?){Write-Verbose "Variable ErrorActionPreference: Revert $ErrorActionPreference Back To: $oldActionPreference"}

    }


 }

 END {
  
  #$report 
  
 if($exportCSV){
        
        Write-Verbose "Exporting to CSV with path of $($exportpath)$(((get-date).ToUniversalTime()).ToString("yyyyMMddThhmmssZ")).csv"
            If($WhatIfPreference -eq $True){
            
                $WhatIfPreference = $False
                
                $report | Export-Csv -Path "$($exportpath)$(((get-date).ToUniversalTime()).ToString("yyyyMMddThhmmssZ")).csv" -Encoding UTF8
                
                $WhatIfPreference = $True
            
            }
    
            Else {
            
                $report | Export-Csv -Path "$($exportpath)$(((get-date).ToUniversalTime()).ToString("yyyyMMddThhmmssZ")).csv" -Encoding UTF8
            
            }   
        
 }
 
  if($exportXML){

        Write-Verbose "Exporting to XML with path of $($exportpath)$(((get-date).ToUniversalTime()).ToString("yyyyMMddThhmmssZ")).xml"

            If($WhatIfPreference -eq $True){
            
                $WhatIfPreference = $False
                
                $report | Export-Clixml -Path "$($exportpath)$(((get-date).ToUniversalTime()).ToString("yyyyMMddThhmmssZ")).xml" -Encoding UTF8 
                
                $WhatIfPreference = $True
            
            }
    
            Else {
            
                $report | Export-Clixml -Path "$($exportpath)$(((get-date).ToUniversalTime()).ToString("yyyyMMddThhmmssZ")).xml" -Encoding UTF8
            
            }   
        
 }

#Export to screen 
 #$report | Select 'Source Recipient','Permission Type',Recipient,'Script Action'

 
 If($EnableTranscript){
    
            Stop-Transcript
            
    }


#Stop Script Stopwatch and Report
    $TotalScriptStopWatch.Stop() 
    Write-Host "Script Completed in $($TotalScriptstopwatch.Elapsed.TotalMinutes) Minutes"
    
 }
    
