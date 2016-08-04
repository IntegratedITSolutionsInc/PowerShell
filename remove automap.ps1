

<#
.Synopsis
   Remove's another user's folder/calendar from automap
.DESCRIPTION
   If a user is assigned, or has been, permissions to another user's calendar or mailbox, 365 automatically maps it
   in the user's Outlook.  This will turn off the automap
.EXAMPLE
   Example of how to use this cmdlet
.EXAMPLE
   Another example of how to use this cmdlet
#>
function Remove-Outlook-Automap
{

    Begin {

        #Connect to client's 365
         $LiveCred = Get-Credential
         Import-Module msonline; Connect-MsolService -Credential $livecred
         $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://ps.outlook.com/powershell/?proxymethod=rps -Credential $LiveCred -Authentication Basic -AllowRedirection
         Import-PSSession $Session
           }


    Process {
            $owner=read-host "Enter Owner's email address:"
            $requestor=read-host "Enter email address of person to remove Automapping from:"
            Add-MailboxPermission -Identity $owner -User $requestor -AccessRights FullAccess -AutoMapping:$false
            }
}