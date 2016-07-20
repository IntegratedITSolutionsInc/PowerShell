

<#
.Synopsis
   Creates and iVantage user on Chartis' 365
.DESCRIPTION
   Will request appropriate information when run
.EXAMPLE
   Example of how to use this cmdlet
.EXAMPLE
   Another example of how to use this cmdlet
#>
function Verb-Noun
{
    [CmdletBinding()]
    [Alias()]
    [OutputType([int])]
    Param
    (
        # Param1 help description
        [Parameter(Mandatory=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=0)]
        $Param1,

        # Param2 help description
        [int]
        $Param2
    )

    Begin
    {

    $LiveCred = Get-Credential
    Import-Module msonline
    Connect-MsolService -Credential $livecred
    $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://ps.outlook.com/powershell/?proxymethod=rps -Credential $LiveCred -Authentication Basic -AllowRedirection
    Import-PSSession $Session

    }
    Process
    {

# Prompt for First name and Last name of the Chartis (iVantage) user
[System.Reflection.Assembly]::LoadWithPartialName("Microsoft.VisualBasic")
$first=[Microsoft.VisualBasic.Interaction]::InputBox('What is the iVantage user FIRST name?:', 'FIRSTNAME')
$last=[Microsoft.VisualBasic.Interaction]::InputBox('What is the iVantage user LAST name?:', 'LASTNAME')
$pass=[Microsoft.VisualBasic.Interaction]::InputBox(‘What is the user PASSWORD:’, ‘PASSWORD’)
$space = " "
$comma = ", "
$title=[Microsoft.VisualBasic.Interaction]::InputBox(‘What is the user Title:’, ‘TITLE’)
$mobile = [Microsoft.VisualBasic.Interaction]::InputBox(‘What is the user Cell Number:’, ‘MOBILE’)
$work = [Microsoft.VisualBasic.Interaction]::InputBox(‘What is the user Work Number:’, ‘WORK’)

$ivantage_email = [Microsoft.VisualBasic.Interaction]::InputBox(‘What is the user NON-Chartis iVantage email?  Example:  jsmith@ivantage.com’, ‘ivantage’)

# Need "first" and "last" variables twice: once for username and once for Principal name name.  Duplicate them here
$pc_first=$first
$pc_last=$last


# Create Principal Name from the above
# If last name isn't long enough there WILL be an error...it can be ignored
$pc_last=$last
$pc_first=$pc_first.substring(0,1)
$principal="$pc_first$pc_last@chartis.com"

#  Add the user to 365
New-MsolUser -DisplayName "$last$comma$first" -FirstName $first -LastName $last -UserPrincipalName $principal -Password $pass
set-user $principal -mobilephone $mobile -phone $work
set-user -identity $principal -title $title


# Set license
Set-MsolUser -UserPrincipalName $principal -UsageLocation US
Set-MsolUserLicense -UserPrincipalName $principal -addlicenses chartis:EXCHANGESTANDARD

#Hide from GAL
set-mailbox -identity $principal -HiddenFromAddressListsEnabled $true

#Forward emails to their iVantage email account
Set-Mailbox -Identity $principal -DeliverToMailboxAndForward $false -ForwardingSMTPAddress $ivantage_email

# Create iVantage alias in the GAL
New-MailContact -Name "$last$comma$first" -ExternalEmailAddress $ivantage_email
Set-Contact "$last$comma$first" -Company "iVantageHealth"
Set-Contact "$last$comma$first" -title $title
set-mailcontact -identity $last$comma$first -alias "$pc_first$last"


# Add iVantage "tag" to comments field
Set-Mailbox –CustomAttribute1 “ivantage” –Identity $principal


# Add to iVantage Distro
Add-DistributionGroupMember -Identity “iVantage@chartis.com” -Member $principal

    }
    End
    {
    }
}
