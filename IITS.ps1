<#
.Synopsis
   This function will find the Kaseya Machine ID of the computer.  It will find the computer name if there is no kaseya agent installed.
.DESCRIPTION
   This function checks the registry for the Kaseya Machine ID.  It can be used in other scripts to find the name.
.EXAMPLE
  Get-KaseyaMachineID
.EXAMPLE
   Another example of how to use this cmdlet
#>
function Get-KaseyaMachineID
{
    [CmdletBinding()]
    [Alias()]
    [OutputType([int])]
    Param
    (
    )
    Begin
    {
    }
    Process
    {
        if($(Get-Process -Name AgentMon -ErrorAction SilentlyContinue).Name)
        { 
            $(Get-ItemProperty -Path "HKLM:\Software\WOW6432Node\Kaseya\Agent\INTTSL74824010499872" -Name MachineID -ErrorAction Stop -ErrorVariable CurrentError).MachineID
        }
        Else
        {
            $env:computername
        }   
    }
    End
    {
    }
}

<#
.Synopsis
   This script emails MSAlarm@integratedit.com.
.DESCRIPTION
   This sciptneeds 2 parameters to work.  It requires a from address and the subject material.  An optional attachment parameter can be used if you wish to attach a file. 
.EXAMPLE
   Email-MSalarm -From "Dkhan@integratedit.com" -Body "This is my Email" -Attachment "C:\Foo.txt"
.EXAMPLE
   Another example of how to use this cmdlet
#>
function Email-MSalarm
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
        $From,

        # Param2 help description
        [Parameter(Mandatory=$true,ValueFromPipelineByPropertyName=$true, Position=1)]
        $Body,
        #Field to enter attachment
        [Parameter(Mandatory=$false, ValueFromPipelineByPropertyName=$true, Position=2)]
        $Attachment
    )

    Begin
    {
        $key = $null
        $key = Get-Content "C:\Scripts\key.key"
        $password = Get-Content "C:\Scripts\passwd.txt" | ConvertTo-SecureString -Key $key
        $credentials = new-object -typename System.Management.Automation.PSCredential -argumentlist "forecast@integratedit.com",$password
        $ErrorLog = "$env:TEMP\EmailMSalarm_IITS.txt"
    }
    Process
    {
        if($Attachment)
        {
            Try
                {
            Send-MailMessage -To MSalarm@integratedit.com -Subject "[$(Get-KaseyaMachineID)] - Emailed form Powershell Script" -body "
            {Script}
        
            $Body"  -Credential $credentials -SmtpServer outlook.office365.com -UseSsl -From $From -Attachments $Attachment -ErrorAction Stop -ErrorVariable CurrentError
        }
            Catch
                {
            "$(Get-Date) - Couldn't email.  Error= $CurrentError ." | Out-File -FilePath $ErrorLog -Force -Append
        }
        }
        Else
        {
            Try
            {
            Send-MailMessage -To MSalarm@integratedit.com -Subject "[$(Get-KaseyaMachineID)] - Emailed form Powershell Script" -body "
            {Script}
        
            $Body"  -Credential $credentials -SmtpServer outlook.office365.com -UseSsl -From $From -ErrorAction Stop -ErrorVariable CurrentError
            }
            Catch
            {
            "$(Get-Date) - Couldn't email.  Error= $CurrentError ." | Out-File -FilePath $ErrorLog -Force -Append
            }
        }
    }
    End
    {
    }
}
