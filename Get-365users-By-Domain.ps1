

<#
.Synopsis
   Get a user list from 365 for a specific domain
.DESCRIPTION
   Chartis is an example.  Some folks are only ivantage.com so this procedure would give you just 
   those names if you needed it.
.EXAMPLE
   Example of how to use this cmdlet
.EXAMPLE
   Another example of how to use this cmdlet
#>
function Get-365users-By-Domain
{

    

    Begin {

        $domain=read-host "Enter domain name, ex. (integratedit.com) :"

        #Connect to client's 365
         $LiveCred = Get-Credential
         Import-Module msonline; Connect-MsolService -Credential $livecred
         $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://ps.outlook.com/powershell/?proxymethod=rps -Credential $LiveCred -Authentication Basic -AllowRedirection
         Import-PSSession $Session
           }


    Process {
    get-recipient | where {$_.emailaddresses -match $domain} | select name, emailaddress | export-csv "C:\iits_mgmt\365_domain_users.csv"

            }

    End {
        get-pssession | remove-pssession
        }
}