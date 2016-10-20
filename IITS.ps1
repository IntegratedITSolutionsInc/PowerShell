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
    Param
    ()
    
    Process
    {
        try
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
        Catch
        {
            $env:computername
        }   
    }
}

<#
.Synopsis
   This script emails MSAlarm@integratedit.com.
.DESCRIPTION
   This scipt needs 1 parameter to work.  It requires the subject.  An optional attachment parameter can be used if you wish to attach a file. 
.EXAMPLE
   Email-MSalarm -Body "This is my Email" -Attachment "C:\Foo.txt"
#>
function Email-MSalarm
{
    [CmdletBinding()]
    Param
    (
        [Parameter(Mandatory=$true,ValueFromPipelineByPropertyName=$true, Position=0)]
        $Body,

        [Parameter(Mandatory=$false, ValueFromPipelineByPropertyName=$true, Position=1)]
        $Attachment
    )

    Begin
    {
        try
        {
        $CurrentError = $null
        $ErrorLog = "$env:windir\Temp\EmailMSalarm_IITS.txt"
        $key = Get-Content "C:\IITS_Scripts\Key.key" -ErrorAction Stop -ErrorVariable CurrentError
        $password = Get-Content "C:\IITS_Scripts\passwd.txt" | ConvertTo-SecureString -Key $key -ErrorAction Stop -ErrorVariable CurrentError
        $credentials = new-object -typename System.Management.Automation.PSCredential -argumentlist "forecast@integratedit.com",$password
        }
        Catch
        {
        "$(Get-Date) - Couldn't get a variable.  Error= $CurrentError ." | Out-File -FilePath $ErrorLog -Force -Append
        }
    }
    Process
    {
        if(!$CurrentError)
        {
            if($Attachment)
                {
            Try
                {
                    Send-MailMessage -To MSalarm@integratedit.com -Subject "[$(Get-KaseyaMachineID)] - Emailed from Powershell Script with attachment." -body "
                    {Script}
        
                    $Body"  -Credential $credentials -SmtpServer outlook.office365.com -UseSsl -From forecast@integratedit.com -Attachments $Attachment -ErrorAction Stop -ErrorVariable CurrentError
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
                Send-MailMessage -To MSalarm@integratedit.com -Subject "[$(Get-KaseyaMachineID)] - Emailed from Powershell Script." -body "
                {Script}
        
                $Body"  -Credential $credentials -SmtpServer outlook.office365.com -UseSsl -From forecast@integratedit.com -ErrorAction Stop -ErrorVariable CurrentError
            }
            Catch
            {
                "$(Get-Date) - Couldn't email.  Error= $CurrentError ." | Out-File -FilePath $ErrorLog -Force -Append
            }
        }
        }
        Else
        {
            "$(Get-Date) - Skipped process block due to not having the key file or the password.  Error = $CurrentError" | Out-File -FilePath $ErrorLog -Force -Append
        }
    }
    End
    {
    }
}

<#
.Synopsis
   This function will enable or disable the windows 10 action center.
.DESCRIPTION
   This function checks for the existence of the neccessary registry keys and will create the keys if needed.  Once they are created or verified to be there then it will change the appropriate dword based on user request
.EXAMPLE
   Toggle-ActionCenter -Setting Enable
.EXAMPLE
   Another example of how to use this cmdlet
#>
function Toggle-ActionCenter
{
    [CmdletBinding()]
    Param
    (
        # Param1 help description
        [Parameter(Mandatory=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=0
                   )]
        [ValidateSet("Enable","Disable")]
        [String]$Setting
    )

    Begin
    {
        $regpath = "HKCU:\Software\Policies\Microsoft\Windows\Explorer"
        $namedword = "DisableNotificationCenter"
        $output= "$env:windir\Temp\actioncenter_IITS.txt"
        if($Setting -eq "Enable")
        {
            $Status = 0
        }
        Else
        {
            $Status = 1
        }
    }
    Process
    {
        try
        {
            $machineID  = $(Get-ItemProperty -Path "HKLM:\Software\WOW6432Node\Kaseya\Agent\INTTSL74824010499872" -Name MachineID -ErrorAction Stop -ErrorVariable CurrentError).MachineID
            if (!(Test-Path $regpath))
            {
                "$(Get-Date) - Registry path does not exist." | Out-File -FilePath $output -Force -Append
                New-Item -Path $regpath -Force -ErrorAction Stop -ErrorVariable CurrentError
                "$(Get-Date) - Created new key $regpath." | Out-File -FilePath $output -Force -Append
                New-ItemProperty -Path $regpath -Name $namedword -Value $Status -PropertyType DWORD -Force -ErrorAction Stop -ErrorVariable CurrentError
                "$(Get-Date) - Created new dword $namedword with value of $Setting." | Out-File -FilePath $output -Force -Append
            }
            else
            {
                 "$(Get-Date) - Registry path exists." | Out-File -FilePath $output -Force -Append
                 New-ItemProperty -Path $regpath -Name $namedword -Value $Status -PropertyType DWORD -Force -ErrorAction Stop -ErrorVariable CurrentError
                 "$(Get-Date) - Set new dword $namedword with value of $Setting." | Out-File -FilePath $output -Force -Append
            }
        }
        Catch
        {
            "$(Get-Date) - Ran into problem getting the machineID" | Out-File -FilePath $output -Force -Append
        }
    }
    End
    {
    }
}

<#
.Synopsis
   This procedure will calculate the total messages sent from all users in Office 365.  It can then email the results along with a .csv file for futher data manipulation
.DESCRIPTION
   There are 3 parts to this procedure.  The first part connects to Office 365 after requesting the credentials.  The credentials that are used will dictate what tenant information is gathered.  
.EXAMPLE
   Get-MailFlowStatistics -Errorlog
#>
function Get-MailFlowStats
{
    Param
    (
        # Enables outputing of an error log (.txt) at conclusion of procedure.
        [switch]$errorlog
    )

    Begin
    {
        $file_path_csv = "$env:windir\Temp\Email_stats_$(get-date -f yyyyMMdd).csv"
        $shouldemail = Read-Host -Prompt "Do you want to have the results emailed to you along with a CSV attachment? Enter 'Yes' if desired. If no email is required then output is .csv located at $file_path_csv"
        if($shouldemail -like "yes"){
        $office365 = Read-Host -Prompt "Do you want to send through Office 365? Enter Yes or No"
            if($office365 -like "Yes"){
            $smtpServer = "outlook.office365.com"
            }
            Else{
            $smtpServer = Read-Host -Prompt "Please enter the IP address of an accessible Exchange Server."
            }
        $from = read-host -Prompt "Please enter the From address."
        $to = Read-Host -Prompt "Please enter the To address."
        }
        $credential = Get-Credential -Message "Please enter an account that has administrator privleges on the Office 365 tenant."
        Write-Output "Closing open powershell session: $(Get-PSSession)"
        Get-PSSession | Remove-PSSession
        try{
            $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://ps.outlook.com/powershell/ -Credential $credential -Authentication Basic -AllowRedirection -ErrorAction SilentlyContinue -ErrorVariable session_issue
            Import-PSSession $Session -ErrorAction SilentlyContinue -ErrorVariable $session_issue
            cls
            }
        catch{
            Write-Output "There was an issue connecting to Office 365 with the credentials supplied.  Please try again or check the error log."
            $session_issue | Out-File -force "$env:TEMP\Error_log_$(get-date -f yyyyMMdd).txt"
            throw
            }
    }
    Process
    {
        [int]$days_past = -(Read-Host -Prompt 'Enter the number of days in the past to gather reports. Up to a max of 30 days')
        $date = get-date
        $report = @()
        $entry = 1
        $recipients = @()
        Write-Output "Getting recipient list for users with a usermailbox."
        $recipients = Get-Recipient -RecipientTypeDetails UserMailBox -RecipientType UserMailBox
        $start_time = Get-Date
        foreach ($recipient in $recipients)
        {
            Write-Output "Currently working on $recipient.  Number $entry of $(($recipients | measure-object).Count) total. $(($entry/$(($recipients | measure-object).Count))*100 -as [int])% Complete"
            try
            {
                $messages_received = Get-MessageTrace -RecipientAddress $recipient.PrimarySMTPAddress -StartDate $date.AddDays($days_past) -EndDate $date | Measure-Object -ErrorAction Stop -ErrorVariable issue
                $messages_sent = Get-MessageTrace -SenderAddress $recipient.PrimarySMTPAddress -StartDate $date.AddDays($days_past) -EndDate $date | Measure-Object -ErrorAction Stop -ErrorVariable issue
                $Prop=[ordered]@{
                            'Display Name'=$recipient.DisplayName
                            'Start Date'=$date.AddDays($days_past)
                            'End Date'=$date
                            'Messages Received'=$messages_received.Count
                            'Messages Sent'=$messages_sent.Count
                            'Total Messages'=($messages_received.Count + $messages_sent.Count)
                            }
                $report += New-Object -TypeName psobject -Property $Prop
                $entry = ++$entry
            }
            catch
            {
                Write-Output "Error with $recipient.  Check error log for exact issue"
                $issues += $issue
                continue
            }
        }
    }
    End
    {
        if($shouldemail -like "yes")
        {
            $report | Export-Csv -Path $file_path_csv -force
            $subject = "Script finished in $(((get-date).Subtract($start_time)).TotalMinutes -as [int]) minutes!"
                try
                {
                Send-MailMessage -To $to -Subject $subject -BodyAsHtml "$($report | ConvertTo-Html),$($issues | ConvertTo-Html)" -Credential $credential -SmtpServer $smtpServer -UseSsl -From $from -Attachments $file_path_csv -ErrorAction Stop -ErrorVariable email_error
                Remove-Item -Path $file_path_csv -Force
                }
                catch
                {
                Write-Output "Something went wrong while trying to email file.  Defaulting to file output."
                Write-Output "The report has been generated and is located at $file_path_csv"
                break
                }
        }
        Else
        {
                $report | Export-Csv -Path $file_path_csv -Force
                Write-Output "The report has been generated and is located at $file_path_csv"
        }
    Write-Output "Closing powershell sessions: $(Get-PSSession)"
    Get-PSSession | Remove-PSSession       
    }
}

<#
.Synopsis
   This removes a word from outcorrect in anything that uses Microsoft Word. 
.DESCRIPTION
   This cmdlet will connect to the local machine's Microsoft Word installation and then remove a word so that it does not autocorrect to another word.  ehr will no longer autocorrect to her.
.EXAMPLE
   Remove-AutoCorrect -WordToRemove "ehr"

#>
function Remove-AutoCorrect
{
    [CmdletBinding()]
    Param
    (
        # The word that should be removed.
        [Parameter(Mandatory=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=0)]
        [string]$WordToRemove
    )

    Begin
    {
    try
    {
        $ErrorLog= "$env:windir\Temp\disableautocorrectoutput_IITS.txt"
        $found=0
        $word = New-Object -ComObject word.application -ErrorAction Stop -ErrorVariable CurrentError
        $word.visible = $false
    }
    catch
    {
        "$(Get-Date) - Problem starting WORD as a com object.  Error = $CurrentError" | Out-File -FilePath $ErrorLog -Append
    }
    }
    Process
    {
        if(!$CurrentError)
        {
            Try
            {
                $entries = $word.AutoCorrect.entries
                foreach($e in $entries)
                { 
                    if($e.name -eq $WordToRemove)
                    {
                        "$(Get-Date) - Found $WordToRemove in Auto Correct List." | Out-File -FilePath $ErrorLog -Append
                        $found=1
                        $e.delete()
                        "$(Get-Date) - Deleted $WordToRemove in Auto Correct List." | Out-File -FilePath $ErrorLog -Append
                    }
                }
                if($found -eq 0)
                {
                    "$(Get-Date) - Did not find $WordToRemove in Auto Correct List." | Out-File -FilePath $ErrorLog -Append
                }
            }
            Catch
            {
                "$(Get-Date) - Something went wrong while trying to remove the word $WordToRemove" | Out-File -FilePath $ErrorLog -Append
            }
        }
        Else
        {
            Email-MSalarm -From Script_Genie@integratedit.com -Body "Problem starting WORD as a com object.  Error = $CurrentError" -Attachment $ErrorLog
        }
    }
    End
    {
        $word.Quit()
        $word = $null
        [gc]::collect()
        [gc]::WaitForPendingFinalizers()
    }
}

<#
.Synopsis
   Delete a user's 365 account
.DESCRIPTION
   Will prompt for email address of user and disable after you give the 365 credentials
.EXAMPLE
   Example of how to use this cmdlet
#>
function disable-365-account
{
    [CmdletBinding()]
    Param ([string]$mailbox=(read-host "Enter user's email address:"))

    #Connect to client's 365
    $LiveCred = Get-Credential
    Import-Module msonline; Connect-MsolService -Credential $livecred
    $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://ps.outlook.com/powershell/?proxymethod=rps -Credential $LiveCred -Authentication Basic -AllowRedirection
    Import-PSSession $Session

    Set-MsolUser -UserPrincipalName $mailbox -BlockCredential $true
    
}

<#
   Get App Version
.DESCRIPTION
   Gives you the exact application version for an app in Windows
.EXAMPLE
   You will be prompted for the full path and application name, ie:  c:\windows\apppp.exe

#>
function get-app-version
{

    Param ([string]$app=(Read-host "FULL path with filename:"))

$VersionString = [System.Diagnostics.FileVersionInfo]::GetVersionInfo($app).Fileversion
$OutputString = "$Machine $QueryString AppVersion $VersionString"


$sep="."
$parts=$VersionString.split($sep)
$parts[0]

}

<#
.Synopsis
   External Push of Kaseya via Powershell
.DESCRIPTION
   If a machine on a network doesn't have powershell OR if it's a machine
   from an acquisistion that we don't have access to yet, this will install
   Kaseya
.EXAMPLE
   After execution, you'll be asked for the path of the Kaseya installer for
   the client
#>
function external-kaseya-push
{
    [CmdletBinding()]
    Param
    (
        # Param1 help description
        [Parameter(Mandatory=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=0)]
        [string]$kaseyaLink=(read-host "Input the full URL for the Kaseya client to be installed.  Example: https://www.dropbox.com/s/o6q9bbe91jcsoa7/sccit_KcsSetup.exe?dl=1")

    )

    Begin
    {
    $sccit_url = $kaseyaLink 
    $kaseya_path= "$Env:SystemDrive\iits_mgmt" 
    }
    Process
    {
    #Create Kaseya download path if it doesn't already exist 
    if ((Test-Path $kaseya_path) -eq $false) {New-Item -type directory -Path $kaseya_path | out-null} 
    $tableurl
    #Download Kaseya if it's not already there 
    if ((Test-Path "$kaseya_path\sccit_KcsSetup.exe") -eq $false)
        { 
            $kaseya_dload_file = "$kaseya_path\sccit_KcsSetup.exe" 
            $kaseya_dload = new-object System.Net.WebClient 
            $kaseya_dload.DownloadFile($sccit_url,$kaseya_dload_file) 
        } 

    #Run Kaseya and wait for it to exit 
    $kaseya_launch = new-object Diagnostics.ProcessStartInfo 
    $kaseya_launch.FileName = "$kaseya_path\sccit_KcsSetup.exe" 

    #$kaseya_launch.Arguments = $kaseya_switches 

    $kaseya_process = [Diagnostics.Process]::Start($kaseya_launch) 
    $kaseya_process.WaitForExit() 
    }
    End
    {
    }
}

<#
.Synopsis
   Hides a user from the GAL
#>
function hide-user-from-GAL
{
    [CmdletBinding()]
    Param
    (
        [Parameter(Mandatory=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=0)]
        [string]$mailbox=(read-host "Enter user's email address:")

    )

    Begin
    {

    #Connect to client's 365
    $LiveCred = Get-Credential
    Import-Module msonline; Connect-MsolService -Credential $livecred
    $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://ps.outlook.com/powershell/?proxymethod=rps -Credential $LiveCred -Authentication Basic -AllowRedirection
    Import-PSSession $Session

    }
    Process
    {

    Set-Mailbox -Identity $mailbox -HiddenFromAddressListsEnabled $true

    }

    End
    {
    }
}

<#
.Synopsis
   Gets the account used for all services
.DESCRIPTION
   Uses WMI to find the list of services and the accounts they are using to start.  Added description to aid in figuring out what it does.  The Output will be returned to the screen and can then be sent to a file. 
.EXAMPLE
   Get-ServiceAccount -ComputerName Computron -LogFile
#>

function Get-ServiceAccount
{
    Param
    (
        # Param1 help description
        [Parameter(Mandatory=$false,
                   ValueFromPipelineByPropertyName=$true,
                   Position=0)]
        $ComputerName,

        [switch][parameter(mandatory=$false, Position=1)] $LogFile
    )

    Begin
    {
        if(!$ComputerName)
        {
            $ComputerName=$env:COMPUTERNAME
        }
        Else
        {
        }
    }
    Process
    {
        $services = Get-WmiObject win32_service -ComputerName $ComputerName -ErrorAction Stop -ErrorVariable CurrentError
        $services | Select-Object -Property SystemName, Name, StartName, ServiceType, Description | Format-Table
    }
    End
    {
        if($LogFile)
        {
            $ErrorLog = "$env:windir\Temp\ServiceAccount_IITS.txt"
            if(!$CurrentError)
            {
                "$(Get-Date) - Error= NO ERROR." | Out-File -FilePath $ErrorLog -Force -Append
            }
            Else
            {
                "$(Get-Date) - Error= $CurrentError." | Out-File -FilePath $ErrorLog -Force -Append
            }
        }
        Else
        {
        }
    }
}

<#
.Synopsis
   This command will output a list of all of the install programs from the registry.
.DESCRIPTION
   This command will find the installed applications on any computer.  It will output the name and the uninstall string that can be used to remove the application.
.EXAMPLE
   Get-InstalledPrograms -ComputerName
#>
function Get-InstalledPrograms
{
    Param
    (
        # ComuputerName
        [Parameter(Mandatory=$false,
                   ValueFromPipelineByPropertyName=$true,
                   Position=0)]
        $ComputerName,
        [Parameter(Mandatory=$false,
                   ValueFromPipelineByPropertyName=$true,
                   Position=1)]
        [switch]
        $ErrorLog
    )

    Begin
    {
        $booboos = @()
        $array = @()
        $RegLocations = 'SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\',
                        'SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\'
        if(!$ComputerName)
        {
            $ComputerName=$env:COMPUTERNAME
        }
        Else
        {
        }
    }
    Process
    {
        try
        {
            $reg=[microsoft.win32.registrykey]::OpenRemoteBaseKey([Microsoft.Win32.RegistryHive]::LocalMachine,$ComputerName)
            foreach($RegLocation in $RegLocations)
            {
            $CurrentKey= $reg.opensubkey($RegLocation)
            $subkeys = $CurrentKey.GetSubKeyNames()
            foreach($subkey in $subkeys)
            {
                $Values = $reg.OpenSubKey("$RegLocation$subkey")
                if($Values.GetValue('DisplayName'))
                {
                    $Prop=[ordered]@{
                    'Display_Name'=$Values.GetValue('DisplayName')
                    'Uninstall_Path'=$Values.GetValue('UninstallString')
                    }
                     $array += New-Object -TypeName psobject -Property $Prop -ErrorAction SilentlyContinue -ErrorVariable errors
                } 
            }
        }
            $array | Sort-Object -Property 'Display_Name'
        }
        catch
        {
            $booboos += $error
        }
    }
    End
    {
        if($ErrorLog)
        {
            $LogPath = "$env:windir\Temp\InstalledPrograms_IITS.txt"
            foreach($booboo in $booboos)
            {
                "$(Get-Date) - $booboo ." | Out-File -FilePath $LogPath -Force -Append
            }
        }
    }
}

<#
.Synopsis
   This program will uninstall any application in ADD/Remove that matches the information entered.
.DESCRIPTION
   This program will remove an appliation with MSIEXEC if applicable.  If there is no msiexec uninstall string then it will attempt to use the uninstall path if there is one. There will be an output file located in the windows\temp folder.  
.EXAMPLE
   Remove-Application -UninstallPrograms Kaseya,Microsoft -ErrorLog
#>
function Remove-Application
{
    [CmdletBinding()]
    Param
    (
        # Param1 help description
        [Parameter(Mandatory=$false,
                   ValueFromPipelineByPropertyName=$true,
                   Position=0)]
        [array]$UninstallPrograms,

        [Parameter()]
        [switch]$ErrorLog
    )

    Begin
    {
        [array]$booboos = @()
        [array]$InstalledPrograms = Get-InstalledPrograms
        if(!$UninstallPrograms)
        {
            $UninstallPrograms = Read-Host 'What program(s) would you like to uninstall?'
        }
    }
    Process
    {
        try
        {
            foreach($program in $UninstallPrograms)
            {
                $progs = $InstalledPrograms | Where-Object {($_.Display_Name -match "$program")}
                if($progs)
                {
                    foreach($prog in $progs)
                    {    
                        if($prog.Uninstall_Path)
                        {
                            if($prog.Uninstall_Path -match "msiexec.exe")
                            {
                                $removestring = $prog.Uninstall_Path -Replace "msiexec.exe","" -Replace "/I","" -Replace "/X",""
                                $removestring = $removestring.Trim()
                                $booboos += "$(Get-Date) - Removing $($prog.display_name) using $removestring."
                                start-process "msiexec.exe" -arg "/X $removestring /qn" -Wait -ErrorAction SilentlyContinue
                            }
                            else
                            {
                                $booboos += "$(Get-Date) - Application $($prog.display_name) isn't using MSIEXEC for uninstall."
                                start-process "cmd.exe" -arg "$($prog.Uninstall_Path)" -Wait -ErrorAction SilentlyContinue
                            }
                        }
                        Else
                        {
                            $booboos += "$(Get-Date) - Application $($prog.Display_name) doesn't have an uninstall path."
                        }
                    }
                }
                else
                {
                    $booboos += "$(Get-Date) - Couldn't find application that matched $program."
                }

            }
        }
        catch
        {
            $booboos += $error
        }
    }
    End
    {
        if($ErrorLog)
        {
            $LogPath = "$env:windir\Temp\RemoveApplication_IITS.txt"
            foreach($booboo in $booboos)
            {
                "$booboo" | Out-File -FilePath $LogPath -Force -Append
            }
        }
    }
}

<#
.Synopsis
   VERY simple function to get versions of all installed apps
.DESCRIPTION
   
.EXAMPLE
  Get-All-App-Versions
.EXAMPLE
   Another example of how to use this cmdlet
#>

function Get-All-App-Versions
{
    Get-ItemProperty HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* | Select-Object DisplayName, DisplayVersion, Publisher, InstallDate | Export-Csv "c:\iits_mgmt\all_app_versions.csv"
}

<#
.Synopsis
   Gathers crashplan log files from a computer and zips them up in a folder.
.DESCRIPTION
   Long description
.EXAMPLE
   Get-CrashPlanLogs "C:\IITS_MGMT\CrashPlan.zip" -ErrorLog
#>
function Get-CrashPlanLogs
{
    Param
    (
        [Parameter(Mandatory=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=0)]
        $Output,
        
        [Parameter(Mandatory=$false,
                   ValueFromPipelineByPropertyName=$true,
                   Position=1)]
        [switch]$Errorlog
    )

    Begin
    {
        $booboos = @()
        $CrashLogPath = "C:\ProgramData\CrashPlan\log"

    }
    Process
    {
        if(Test-Path $CrashLogPath)
        {
            Create-Zip -Source $CrashLogPath -Destination $Output
            return $Output
        }
        else
        {
            $booboos += "$(Get-Date) - CrashPlan log directory $CrashLogPath doesn't exist."
        }
    }
    End
    {
        if($ErrorLog)
        {
            $LogPath = "$env:windir\Temp\GetCrashPlanLogs_IITS.txt"
            foreach($booboo in $booboos)
            {
                "$booboo" | Out-File -FilePath $LogPath -Force -Append
            }
        }
    }
}

<#
.Synopsis
   This command will zip a folder
.DESCRIPTION
   Long description
.EXAMPLE
   Create-Zip -Source "C:\Temp\Logs" -Destination "C:\Temp\Logs.zip" -Overwrite -Errorlog
#>
function Create-Zip
{
    [CmdletBinding()]
    Param
    (
        # Source Directory
        [Parameter(Mandatory=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=0)]
        $Source,

        # Destination for Zip file
        [Parameter(Mandatory=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=1)]
        $Destination,

        [Parameter(Mandatory=$false,
                   ValueFromPipelineByPropertyName=$true,
                   Position=2)]
        [switch]$Overwrite,
        [switch]$Errorlog
    )

    Begin
    {
        $Error = $null
        $booboos = @()
        if(!$Overwrite)
        {
            if(Test-Path $Destination)
                {
                    $Destination = "$($Destination).new"
                    $booboos += "$(Get-Date) - File already exists.  Creation new file $Destination."
                }
        }
        Else
        {
            if(test-path $Destination)
            {
                Remove-Item $Destination
                $booboos += "$(Get-Date) - Removing previous file $Destination."
            }
        }
    }
    Process
    {
        Try
        {
        Add-Type -assembly "system.io.compression.filesystem" -ErrorAction Stop
        [io.compression.zipfile]::CreateFromDirectory($Source, $Destination) 
        }
        Catch
        {
            $booboos += "$(Get-Date) - Couldn't load assembly.  Error = $Error"
        }
    }
    End
    {
         if($ErrorLog)
        {
            $LogPath = "$env:windir\Temp\CreateZip_IITS.txt"
            foreach($booboo in $booboos)
            {
                "$booboo" | Out-File -FilePath $LogPath -Force -Append
            }
        }
    }
}

<#
.Synopsis
   Will request a domain name and see if it's blacklisted
.DESCRIPTION
   Long description
.EXAMPLE
   Michael takes no credit for this code!
.EXAMPLE
   Another example of how to use this cmdlet
#>
function Find-If-Domain-Blacklisted
{
    Param
    ()

    Begin
    {
    }
    Process
    {
    write-host "Stand by..."
    $wc=New-Object net.webclient
        #$IP =(Invoke-WebRequest ifconfig.me/ip).Content.Trim()
        #$IP = $wc.downloadstring("http://ifconfig.me/ip") -replace "[^\d\.]"
        Try {
            $IP = $wc.downloadstring("http://checkip.dyndns.com") -replace "[^\d\.]"
        }
        Catch {
            $IP = $wc.downloadstring("http://ifconfig.me/ip") -replace "[^\d\.]"
        }
        $IP = Read-Host -prompt "Enter a domain name to see if it's on a blacklist:"
        Write-Host "Testing Public IP: $IP"
        $reversedIP = ($IP -split '\.')[3..0] -join '.'
 
        $blacklistServers = @(
            'b.barracudacentral.org'
            'spam.rbl.msrbl.net'
            'zen.spamhaus.org'
            'bl.deadbeef.com'
            'bl.spamcannibal.org'
            'bl.spamcop.net'
            'blackholes.five-ten-sg.com'
            'bogons.cymru.com'
            'cbl.abuseat.org'
            'combined.rbl.msrbl.net'
            'db.wpbl.info'
            'dnsbl-1.uceprotect.net'
            'dnsbl-2.uceprotect.net'
            'dnsbl-3.uceprotect.net'
            'dnsbl.cyberlogic.net'
            'dnsbl.sorbs.net'
            'duinv.aupads.org'
            'dul.dnsbl.sorbs.net'
            'dyna.spamrats.com'
            'dynip.rothen.com'
            'http.dnsbl.sorbs.net'
            'images.rbl.msrbl.net'
            'ips.backscatterer.org'
            'misc.dnsbl.sorbs.net'
            'noptr.spamrats.com'
            'orvedb.aupads.org'
            'pbl.spamhaus.org'
            'phishing.rbl.msrbl.net'
            'psbl.surriel.com'
            'rbl.interserver.net'
            'relays.nether.net'
            'sbl.spamhaus.org'
            'smtp.dnsbl.sorbs.net'
            'socks.dnsbl.sorbs.net'
            'spam.dnsbl.sorbs.net'
            'spam.spamrats.com'
            't3direct.dnsbl.net.au'
            'tor.ahbl.org'
            'ubl.lashback.com'
            'ubl.unsubscore.com'
            'virus.rbl.msrbl.net'
            'web.dnsbl.sorbs.net'
            'xbl.spamhaus.org'
            'zombie.dnsbl.sorbs.net'
            'hostkarma.junkemailfilter.com'
        )
 
        $blacklistedOn = @()
 
        foreach ($server in $blacklistServers)
        {
            $fqdn = "$reversedIP.$server"
            #Write-Host "Testing Reverse: $fqdn"
            try
            {
              #Write-Host "Trying Blacklist: $server"
                $result = [System.Net.Dns]::GetHostEntry($fqdn)
             foreach ($addr in $result.AddressList) {
              $line = $Addr.IPAddressToString
             } 
            #IPAddress[] $addr = $result.AddressList;
                #$addr[$addr.Length-1].ToString();
            #Write-Host "Blacklist Result: $line"
            if ($line -eq "127.0.0.1") {
                    $blacklistedOn += $server
            }
            }
            catch { }
        }
 
        if ($blacklistedOn.Count -gt 0)
        {
            # The IP was blacklisted on one or more servers; send your email here.  $blacklistedOn is an array of the servers that returned positive results.
            Write-Host "$IP was blacklisted on one or more Lists: $($blacklistedOn -join ', ')"
            Exit 1010
        }
        else
        {
            Write-Host "$IP is not currently blacklisted on any lists checked."
           
        }
 
    }
    End
    {
    }
}

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

<#
.Synopsis
   This cmdlet will return the current filesystem drives as an object that are drivetype of 3 according to it's WMI object. 
.DESCRIPTION
   This cmdlet gets all of the drives that are marked as filesystem drives and returns them as an object to use in any way needed.
.EXAMPLE
   Get-DriveStatistics -ErrorLog
#>
function Get-DriveStatistics
{
    Param
    (
        [switch]$ErrorLog
    )

    Begin
    {
        $booboos=@()
        $error = $null
        $volumes = @()
        $drives = @()
        try
        {
            $drives = Get-PSDrive | Where-Object {$_.Provider -match 'Filesystem'}
            $booboos += "$(Get-Date) - Obtained volumes."
            $fixeddisks = Get-WmiObject -Class Win32_LogicalDisk -Filter "DriveType=3"
            $booboos += "$(Get-Date) - Obtained WMI Objects"
            #the next block of code will compare the drives found in get-psdrive to the drives found with get-wmiobject.  If they match then that is a drive to use.  If they don't match then it isn't a local disk and will be tossed aside.  This script block can be modified (in a new function) to compare any of the parameteres in PSDrive and WMIObject Logical Disk. 
            foreach($drive in $drives)
            {
                $booboos += "$(Get-Date) - Comparing $($drive.root)."
                foreach($fixeddisk in $fixeddisks)
                {
                    $booboos += "$(Get-Date) - Comparing $($drive.root) with $($fixeddisk.deviceid)\."
                    if($drive.root -notlike "$($fixeddisk.DeviceID)\")
                    {
                        $booboos += "$(Get-Date) - Not adding $($drive.root) to volumes array)."
                    }
                    else
                    {
                        $booboos += "$(Get-Date) - Adding $($drive.root) to volumes array."
                        $volumes += $drive
                    }
                }
            }
        }
        catch
        {
            $error += "$(Get-Date) - Couldn't get drive list."
            $booboos += $error
        }
    }
    Process
    {
        $report=@()
        if(!$error)
        {
            foreach($volume in $volumes)
            {
                $Prop=
                [ordered]@{
                'Name'=$volume.Name
                'Drive'=$volume.root
                'UsedSpace'=[System.Math]::Round($($volume.used / 1GB), 2)
                'FreeSpace'=[System.Math]::Round($($volume.free / 1GB) ,2)
                'TotalSpace'=[System.Math]::Round($($volume.used /1gb + $volume.free/1gb), 2)
                }
                $report += New-Object -TypeName psobject -Property $Prop   
            }
            return $report
        }
        else
        {
            $booboos += "$(Get-Date) - Skipping process block due to not having volumes."
        }
    }
    End
    {
        if($ErrorLog)
        {
            $LogPath = "$env:windir\Temp\DriveStatistics_IITS.txt"
            foreach($booboo in $booboos)
            {
                "$booboo" | Out-File -FilePath $LogPath -Force -Append
            }
        }
    }
}

<#
.Synopsis
   This function will export a csv file to C:\IITS_Scripts\DiskInformation that contains disk information.  There will be one file created for each volume including removable drives.
.DESCRIPTION
   This function gathers the disk information and figures out the change in disk usage as a daily change in GB and that day's change percentage.  
   This is all calculated using the used space of the drive.  There is an error log that is stored in the windows temp file directory.
   This function will also gather VSS information if it is run as an administrator.  This can be useful for figuring out if a VSS aware backup is working correctly
.EXAMPLE
   Get-DiskChanges -ErrorLog
#>
function Get-DiskChanges
{
    Param
    (
        [switch]$ErrorLog
    )

    Begin
    {
        #Set variable so that process block will run
        $stop = 0
        #Getting relevent drive information as well as VSS information
        Try
        {
            $booboos = @()
            $import = @()
            $volumes = Get-DriveStatistics -ErrorLog -ErrorAction Stop
            $Shadows = Get-VSSStatistics -ErrorLog -ErrorAction stop
            if($Shadows -match "ERROR. Need to run as administrator! Error = Initialization failure")
            {
                $Shadows = $null
                $booboos += "$(Get-Date) - Couldn't run Get-VSSStatistics. Due to administrive privleges.  Running without VSS information."
            }
            $drives = @()
            foreach($volume in $volumes)
            {
                $booboos += "$(Get-Date) - Going through each volume.  Starting with $volume."
                $shadow = $Shadows | Where-Object {($_.name -eq $volume.name)}
                $booboos += "$(Get-Date) - Figuring out which shadowstorage object matches volume. Shadow = $shadow"
                if($Shadow) #This runs if there is good object data in Shadow
                {
                    $Prop=
                    [ordered]@{
                    'Name' =$volume.name
                    'Drive'=$volume.drive
                    'UsedSpace' = $volume.UsedSpace
                    'FreeSpace' = $volume.FreeSpace
                    'TotalSpace' = $volume.TotalSpace
                    'VSSAllocatedSpaceGB' = $shadow.VSSAllocatedSpaceGB
                    'VSSUsedSpaceGB'= $shadow.VSSUsedSpaceGB
                    'VSSMaxSpaceGB'= $shadow.VSSMaxSpaceGB
                    }
                    $drives += New-Object -TypeName psobject -Property $Prop
                }
                Else #this runs if there is nothing in Shadow.
                {
                    $booboos += "$(Get-Date) - There was no information in Shadow."
                    $Prop=
                    [ordered]@{
                    'Name' =$volume.name
                    'Drive'=$volume.drive
                    'UsedSpace' = $volume.UsedSpace
                    'FreeSpace' = $volume.FreeSpace
                    'TotalSpace' = $volume.TotalSpace
                    'VSSAllocatedSpaceGB' = "NONE"
                    'VSSUsedSpaceGB'= "NONE"
                    'VSSMaxSpaceGB'= "NONE"
                    }
                    $drives += New-Object -TypeName psobject -Property $Prop
                }
             }     
        }
        Catch
        {
            $booboos += "$(Get-Date) - Couldn't get drive lists."
            #Stops process block from running
            $stop = 1
        }
    }
    Process
    {
        if($stop -eq 0) #stops the script if no drives could be found
        {
            Foreach($drive in $drives)
            {
                $booboos += "$(Get-Date) - Processing $($drive.name)."
                if(Test-Path "C:\IITS_Scripts\DiskInformation\$($drive.name).csv") #checking for the existance of the csv file form a previous run.
                {
                    $booboos += "$(Get-Date) - File Exists: C:\IITS_Scripts\DiskInformation\$($drive.name).csv."
                    #importing csv for manipulation
                    $import += Import-Csv -Path "C:\IITS_Scripts\DiskInformation\$($drive.name).csv"
                    $booboos += "$(Get-Date) - Importing old drive informationfor $($drive.name)."
                    #addming new columns
                    $import += $drive | Select-Object *, "ChangeGBUsed", "ChangeRatePercentUsed", "Date", "Time"
                    $booboos += "$(Get-Date) - Appending new drive information for $($drive.name)."
                    if($import.count -ge 2) #Skipping math if there is only 1 row of information.  IE the information has only been gathered once. 
                    {
                        $new = $import[-1]
                        $old = $import[-2]
                        $new.ChangeGBUsed = $old.UsedSpace - $new.UsedSpace
                        if($old.UsedSpace -eq 0)
                        {
                            $booboos += "$(Get-Date) - No Change for $($drive.name)."
                            $new.ChangeRatePercentUsed = 0
                        }
                        Else
                        {
                            $new.ChangeRatePercentUsed = ((($new.UsedSpace - $old.UsedSpace)/($old.UsedSpace))*100)
                        }
                    }
                    Else #Calculating the changes since there are at least 2 data points.
                    {
                        $booboos += "$(Get-Date) - Only one entry for $($volume.name)."
                    }
                    $new.date =  Get-Date -Format d
                    $new.time =  Get-Date -Format T
                    $booboos += "$(Get-Date) - Outputting new drive information to existing CSV for $($drive.name)."
                    #Appending new information to existing csv file.
                    $new | Export-Csv -Path "C:\IITS_Scripts\DiskInformation\$($drive.name).csv" -Force -Append
                }
                Else
                {
                    $booboos += "$(Get-Date) - Creating CSV for $($drive.name)."
                    #creating csv file since one does not exist
                    $export = $drive | Select-Object *, "ChangeGBUsed", "ChangeRatePercentUsed", "Date" , "Time"
                    $export.date =  Get-Date -Format d
                    $export.time =  Get-Date -Format T
                    $export.ChangeGBUsed = 0
                    $export.ChangeRatePercentUsed = 0
                    try
                    {
                        New-Item -Path "C:\IITS_Scripts\DiskInformation" -ItemType Directory -ErrorAction Stop -ErrorVariable error | Out-Null
                    }
                    Catch
                    {
                        $booboos += "$(Get-Date) - Error creating DiskInformation folder. Error = $error"
                    }
                    $export | Export-Csv -Path "C:\IITS_Scripts\DiskInformation\$($drive.name).csv" -Force -Append
                }
            
        }
        }
        Else
        {
            $booboos += "$(Get-Date) - Errors were detected before process block."
        }
    }
    End
    {
        if($ErrorLog)
        {
            $LogPath = "$env:windir\Temp\DiskChanges_IITS.txt"
            foreach($booboo in $booboos)
            {
                "$booboo" | Out-File -FilePath $LogPath -Force -Append
            }
        }
    }
}

<#
.Synopsis
   Checks for outdated PowerShell version and returns True if it is. Option to e-mail MSAlarm on True.
.DESCRIPTION
   Checks the installed PowerShell version (Major) and sees if it's less than 3. If so, it returns True. If the ticket switch is enabled, it also e-mails MSAlarm with a request to update it.
.EXAMPLE
   Check-PSVersion

   True
.EXAMPLE
   Check-PSVersion -ticket

   True
   [ticket created in ConnectWise]
.INPUTS
   No inputs, optional 'ticket' switch.
.OUTPUTS
   Boolean. If 'ticket' switch called, also sends an e-mail.
#>
function Check-PSVersion
{
    Param
    (
        # Switch parameter; call to send e-mail to MSAlarm, which will make a ticket.
        [Switch]$ticket
    )
    
    if ($PSVersionTable.PSVersion.Major -lt 3)
    {
        echo $true
        if ($ticket)
        {
            $id = Get-KaseyaMachineID
            Email-MSalarm -Body "$id needs a PowerShell upgrade."
        }
    }
    else {echo $false}
}

<#
.Synopsis
   This function will output the size of the VSS store on a machine for all volumes that have shadow volume usage. THIS NEEDS TO BE RUN AS WITH ADMINISTATIVE ACCESS!
.DESCRIPTION
   This function gathers all volumes according to the wwin32_volume class as well as the volumes as reported by win32_shadowstorage class.  It will compare the 2 and weed out any volumes that do not have used vss space.  It will then output the results in a custom object.
.EXAMPLE
   Get-VSSStatistics -ErrorLog
#>
function Get-VSSStatistics
{
    Param
    (
        [switch]$ErrorLog
    )

    Begin
    {
        $booboos = @()
        $errors = $null
        try
        {
            #gather all volumes on the computer
            $Volumes = Get-WmiObject -Class Win32_Volume -ErrorAction Stop -ErrorVariable errors
            #gather all shadowstorage objects.
            $ShadowStorageObjects = Get-WmiObject -Class Win32_ShadowStorage -ErrorAction Stop -ErrorVariable errors
        }
        Catch [System.Management.ManagementException]
        {
            $booboos += "$(Get-Date) - Need to run script as administrator. ERROR = $errors."
            return "ERROR. Need to run as administrator! Error = $($error[0])"
        }
        Catch
        {
            $booboos += "$(Get-Date) - Something went wrong with getting WMIObjects."
        }
        [array]$report = @()
    }
    Process
    {
        foreach($ShadowStorageObject in $ShadowStorageObjects)
        {
            foreach($Volume in $Volumes)
            {
                If($ShadowStorageObject.volume -eq $Volume.__RELPATH)
                {
                    $Prop=
                    [ordered]@{
                    'Name' = $Volume.driveletter.trimend(":")
                    'VSSDrive'=$volume.name
                    'VSSAllocatedSpaceGB' = [System.Math]::Round(($ShadowStorageObject.AllocatedSpace /1GB), 3)
                    'VSSUsedSpaceGB'=[System.Math]::Round(($ShadowStorageObject.usedspace /1GB), 3)
                    'VSSMaxSpaceGB'= [System.Math]::Round(($ShadowStorageObject.maxspace /1GB), 3)
                    }
                    $report += New-Object -TypeName psobject -Property $Prop  
                }
                else
                {
                    $booboos += "$(Get-Date) - $($ShadowStorageObject.volume) didn't match $($Volume.__RELPATH)"
                }
            }
        }
        return $report
     }
    End
    {
        if($ErrorLog)
        {
            $LogPath = "$env:windir\Temp\VSSStatistics_IITS.txt"
            foreach($booboo in $booboos)
            {
                "$booboo" | Out-File -FilePath $LogPath -Force -Append
            }
        }
    }
}

<#
.Synopsis
   This function calculates the projected number of days in which the free space on Non-Removable drive will be used completely. If the space will be used in less than 30 days it will send an e-mail.
   .DESCRIPTION
   This function collects the .csv files for each of the drives located in the C:\IITS_Scripts\DiskInformation folder. Based on the total used space, free size and the number of days 
   the data is collected for it calculates the rate of change for a single day. It then projects the number of days in which the free space will be used completely.
   The script does not have any parameter, however it does use 2 other functions : Get-KaseyaMachineID and Email-MSAlarm.
   .EXAMPLE
    Get-Projection
#>
function Get-Projection {
    <# LOG FILE is created to output the information to.Log file exists at  C:\IITS_Mgmt\Temp\DiskInformation\logs.txt
    checks if the log file exist in the first if block, if it exists it adds a general comment , if it does not exist then runs the code in the else statement where 
    it makes the DiskInformation folder and the logs.txt file  #>
    $logfile = 'C:\IITS_Mgmt\Temp\DiskInformation\logs.txt'
    $date = Get-Date
    $testlogfile = Test-Path $logfile
        if ($testlogfile -eq $true) {
        Add-Content $logfile "logs on $date"
        }
        else {
        New-Item -Path 'C:\IITS_Mgmt\Temp' -ItemType 'Directory' -Name 'DiskInformation'
        New-Item -Path 'C:\IITS_Mgmt\Temp\DiskInformation' -ItemType 'file' -Name 'logs.txt'
        Add-Content $logfile "This is where we collect logs for Disk Usage"
        }
        <# each volume (if not a removable drive) has a .csv file in the DiskInformation folder. The $csv variable contains all the .csv files in that location #>   
        $csv = Get-ChildItem C:\IITS_Scripts\DiskInformation -Filter *.csv

        <# $c is the csv file for each volume #>
        ForEach ($c in $csv) {
        <# each volumes csv file is made into an object represented by $volcsv #>
        $volcsv = Import-csv C:\IITS_Scripts\DiskInformation\$c

        <# This is the total used space for each drive, calculated by adding the ChangeGBUsed column in the csv for each volume #>
        $totalused = ($volcsv.ChangeGBUsed | Measure-Object -Sum).Sum

        <# $Size is the TotalSpace of the volume taken as the Maximum value from the TotalSpace column #>
        $Size = ($volcsv.TotalSpace | Measure-Object -Maximum).Maximum

        <# Runs the Get-KaseyaMachineID FUNCTION to get the name of the machine to be used later #>
        $machine = Get-KaseyaMachineID

        <# The variable is all the points in the csv file for which the date is collected #>
        $countall = $volcsv.count

        <# 1 less than the total count to be used in the next step to get the latest freespace #>
        $count = $volcsv.count-1

        <# from the csv file, the variable gets the name of the drive in the foreach loop #>
        $drive = $volcsv.name[$count]

        <# free space on the drive, last item in the freespace column #>
        $freespace = $volcsv.freespace[$count]

            <# if the total used space is greater than 0 #>
            if ($totalused -gt 0.00) {
            <# gets the current time with the last time and date entry in the sheet #>
            $currtime = $volcsv.date[$count]+" "+$volcsv.time[$count]
            <# gets the start time from the first entry of time and date in the csv file #>
            $starttime = $volcsv.date[0]+" "+$volcsv.time[0]
            <# Calculates the Time Difference from $currtime and $starttime #>
            $timediff = [math]::Round(([datetime]$currtime - [datetime]$starttime).TotalDays,2)
            <# Daily usage is calculated by dividing the totalused by the number of days it was used in #>
            $dailyused = $totalused / $timediff
            <# freespace currently on a volume is divided by the dailyused average to get the projected days #>
            $projusedays = [math]::Round(($freespace / $dailyused),2)
                    <# if $projusedays is less than 1 month sends an e-mail , if not outputs to a log file #>
                    if ($projusedays -le 30 ) {
                    Email-MSalarm -Body "$drive drive on $machine is low on free disk space. In the last $timediff days, $totalused GB was used. Based on this trend $freespace GB will be used in $projusedays days" -Attachment C:\IITS_Scripts\DiskInformation\$c
                    Add-Content $logfile "$drive drive has a free space of $freespace GB. For the last $timediff days the $totalused GB was used, FREE space will exhaust in $projusedays days "
                    }
                    else {
                    Add-Content $logfile "$drive drive has a free space of $freespace GB. For the last $timediff days the $totalused GB was used, space will not exhaust in 30 days instead will exhaust in $projusedays days "
                    }
             }
             <# If the total used space is less than 0 ,is a negative #>
            else {
             Add-Content $logfile "The average space used for $drive DRIVE is $totalused GB. email will not be sent because free space will not exhaust"
            }
               
        }
    }
        
<#
.Synopsis
   Disable a user's 365 account
.DESCRIPTION
   Will prompt for email address of user and disable after you give the 365 credentials
.EXAMPLE
   Example of how to use this cmdlet
.EXAMPLE
   Another example of how to use this cmdlet
#>
function disable-365-account
{
    [CmdletBinding()]
    Param ([string]$mailbox=(read-host "Enter user's email address:"))

    #Connect to client's 365
    $LiveCred = Get-Credential
    Import-Module msonline; Connect-MsolService -Credential $livecred
    $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://ps.outlook.com/powershell/?proxymethod=rps -Credential $LiveCred -Authentication Basic -AllowRedirection
    Import-PSSession $Session

    Set-MsolUser -UserPrincipalName $mailbox -BlockCredential $true
}

<#
.Synopsis
   This cmdlet will schedule get-diskchanges to run as a shceduled task for every hour. 
.DESCRIPTION
   This function will scheduled get-diskchanges to run every 10 minutes so that statistics can be gathered.  
.EXAMPLE
   Deploy-GetDiskChanges
#>
function Deploy-GetDiskChanges
{
    [CmdletBinding()]
    Param
    (
        # Switch for error logging. 
        [Switch]
        $ErrorLog
    )

    Begin
    {
        $booboos = @()
        $PSVersion = Check-PSVersion -ticket #Check the powershell version. Needs to be at least version 3.  Sends email to msalarm if version is less than 3. 
        if($PSVersion -eq $false)
        {
            $stop = 0
            $booboos += "$(Get-Date) - Powershell version is 3 or greater."
        }
        else
        {
            $stop = 1
            $booboos += "$(Get-Date) - Powershell version is less than 3."
        }
    }
    Process
    {
        if($stop -eq "0")
        {
            $booboos += "$(Get-Date) - Executing process block if statement because powershell version is 3 or higher."
            $CurrentScheduledTask = Get-ScheduledTask | Where-Object {($_.TaskName -eq 'GetDiskChanges')} #Getting list of tasks and seeing if there is already a task by the name of GetDiskChanges
            if($CurrentScheduledTask)
            {
                $booboos += "$(Get-Date) - Scheduled task $CurrentScheduledTask already exists.  Skip creation process."
            }
            else
            {
                $booboos += "$(Get-Date) - No GetDiskChanges task found. Creating a new task."
                $TaskName = 'GetDiskChanges'
                $action = New-ScheduledTaskAction -Execute 'powershell.exe'-Argument '-NoProfile -WindowStyle Hidden -verb runas -command "Get-DiskChanges"'
                $trigger = New-ScheduledTaskTrigger -Once -At 9am -RepetitionInterval (New-TimeSpan -Minutes 10) -RepetitionDuration (New-TimeSpan -Days 9000)
                $settings = New-ScheduledTaskSettingsSet -Priority 10
                try
                {
                    Register-ScheduledTask -TaskName $TaskName -Action $action -Trigger $trigger -User <#Username goes here!!!!#> -Password <#Password goes here#> -Description "This gathers disk information every 10 minutes." -Settings $settings -ErrorAction Stop
                }
                Catch
                {
                    $booboos += "$(Get-Date) - Couldn't create new scheduled task.  Error is $error[0]."
                }                
            }
        }
        else
        {
            $booboos += "$(Get-Date) - Skipping process block because powershell version is less than 3 and scheduled tasks can't be created."
        }   
    }
    End
    {
        if($ErrorLog)
        {
            $LogPath = "$env:windir\Temp\DeployDiskChanges_IITS.txt"
            foreach($booboo in $booboos)
            {
                "$booboo" | Out-File -FilePath $LogPath -Force -Append
            }
        }
    }
}

<#
.Synopsis
   THis script will send out the three patching notifications prior to server patching if the day is correct. 
.DESCRIPTION
   This script finds the first day of the month and then extrapolates the 4th tuesday from that day.  It will then compare the current day to figure out if it's the Friday preceding, Monday preceding, or the day of patching.  It will do something specific for each of those days. 
.EXAMPLE
   Send-PatchEmail -ErrorLog
#>
function Send-PatchEmail
{
    [CmdletBinding()]
    Param
    (
        [switch]$ErrorLog
    )

    Begin
    {
        
        $booboos = @()
        $currentdate = Get-Date
        $booboos += "$(Get-Date) - Today's date found as $currentdate."
        $firstofthemonth = Get-Date -Day 1
        $booboos += "$(Get-Date) - Found the first of the month as $firstofthemonth."
        switch ($firstofthemonth.DayOfWeek)
        {
            "Sunday"    {$patchTuesday = $firstofthemonth.AddDays(23); break} 
            "Monday"    {$patchTuesday = $firstofthemonth.AddDays(22); break} 
            "Tuesday"   {$patchTuesday = $firstofthemonth.AddDays(21); break} 
            "Wednesday" {$patchTuesday = $firstofthemonth.AddDays(27); break} 
            "Thursday"  {$patchTuesday = $firstofthemonth.AddDays(26); break} 
            "Friday"    {$patchTuesday = $firstofthemonth.AddDays(25); break} 
            "Saturday"  {$patchTuesday = $firstofthemonth.AddDays(24); break} 
        }
        $booboos += "$(Get-Date) - Found patch tuesday to be $patchTuesday."
    }
    Process
    {
        $securepwd = Get-Content -Path 'C:\PatchEmail\Passwd.txt' | ConvertTo-SecureString
        $credentials = new-object -typename System.Management.Automation.PSCredential -argumentlist "Managed.Services",$securepwd
        if($patchTuesday.AddDays(-4).day -eq $currentdate.day)
        {
            $booboos += "$(Get-Date) - Found today is the Friday before patching."
            $Phrase = "Next week Tuesday"
            $email = $true
        }
        elseif($patchTuesday.AddDays(-1).day -eq $currentdate.Day)
        {
            $booboos += "$(Get-Date) - Found today is the Monday before patching."
            $Phrase = "Tomorrow"
            $email = $true

        }
        Elseif($patchTuesday.day -eq $currentdate.Day)
        {
            $booboos += "$(Get-Date) - Found today is the day of patching."
            $Phrase = "Today"
            $email = $true
            $email_eng = $true
        }
        else
        {
            $booboos += "$(Get-Date) - Found today is not either of the right patching days."
            $email = $false
        }
        if($email -eq $true)
        {
            $Subject = "Reminder: Integrated IT Solutions is patching servers on $($patchtuesday | get-date -format D)."
            $Body = "Hi,

$Phrase $($patchtuesday | get-date -format D), is the fourth Tuesday of the month, so in accordance with our patching schedule, we will be patching your servers. Reboots will happen after hours starting at 9pm. Please respond back to this email if there are conflicts with patching your server(s) $Phrase, $($patchtuesday | get-date -format D)!

Any Machines which have been previously discussed as being excluded from patching will continue to be excluded until you tell us otherwise. As a reminder, workstatations are patched according to your agreed upon schedule as detailed in your Managed Services agreement.

Please contact your account manager if you would like to review or change any of your patching schedules.  Thank you for your continued support of our Managed Services Program!

Managed Services Team
Integrated IT Solutions
781-742-2200 Option 2
ITHelp@intgratedit.com"

            Send-MailMessage -SmtpServer 10.12.0.85 -from Managed.Services@integratedit.com -to Managed.Services@integratedit.com -Bcc IITS_Patching_Clients@integratedit.com -Subject $Subject -Body $Body -Credential $Credentials
        }
        else
        {
            $booboos += "$(Get-Date) - Not Sending email since it's not the right day."
        }
        if($email_eng -eq $true)
        {
            $Subject_eng = "IMPORTANT!!!!  IITS CLIENT PATCHING IS HAPPENING TONIGHT!!!"
            $Body_eng = "Hi All,
$Phrase $($patchTuesday | get-date -format D) is the IITS client patching day.  THe vast majority of servers will be patched tonight starting at 9pm.  Reboots will happen after patching.  Please check Kaseya's Patch Management module if you have any questions pertaining to a specific client.
            
Thanks,
            
Managed Services Team"
            Send-MailMessage -SmtpServer 10.12.0.85 -from Managed.Services@integratedit.com -to Engineers@integratedit.com -Subject $Subject_eng -Body $Body_eng -Credential $Credentials
        }
    }
    End
    {
        if($ErrorLog)
        {
            $LogPath = "$env:windir\Temp\patchemail_IITS.txt"
            foreach($booboo in $booboos)
            {
                "$booboo" | Out-File -FilePath $LogPath -Force -Append
            }
        }
    }
}

<#
.Synopsis
   Automatically install ESET.
.DESCRIPTION
   Installs ESET Agent and either ESET Endpoint or ESET File Security (if workstation or server, accordingly). Will not install components that are already present. If there's any failure, will notify the user and Managed Services.
.INPUTS
   ticket (switch): Enable to have this script send an e-mail to MSAlarm on completion, which will make a ticket on the Monitoring board. Even if called, will ONLY send an e-mail if there was actually an error
   (i.e. will not send an e-mail if everything ran without incident).
.EXAMPLE
   Install-Eset
#>
function Install-Eset
{
    [CmdletBinding()]
    [Alias("Get-Eset")]
    Param
    (
        # Optional switch to enable creating a ticket if any problems occur.
        [Switch]$ticket
    )

    Begin
    {
        # Initialize the logs array
        $logs = @("===================","$(Get-Date) - Beginning Install-Eset.")

        # Initialize the killswitch
        $kill = 0

        # The Kaseya temp folder
        $KasTemp = "C:\IITS_Mgmt\Temp"

        # The download location of the ESET Agent
        $agentPath = "$KasTemp\EsetAgent.msi"

        # The download location of ESET Endpoint
        $endPath = "$KasTemp\EsetEndpoint.msi"

        # The download location of ESET File Security
        $FSPath = "$KasTemp\EsetFS.msi"

        # The location of all generated log files
        $LogRoot = "$env:windir\Temp"
    }

    Process
    {
        $logs += "$(Get-Date) - Checking installed PowerShell version."
        if(Check-PSVersion)
        {
            $logs += "$(Get-Date) - Installed PowerShell version ($($PSVersionTable.PSVersion.Major)) is outdated and incompatible."
            $kill++
        }

        $logs += "$(Get-Date) - Checking local Add/Remove for any known security products."
        $avs = Check-InstalledAv

        # If any security products were found, return that instead of continuing.
        if($avs -ne "none found")
        {
            # Convert the list-of-avs object to an array of strings
            [string]$avlist = @()
            foreach($item in $avs){$avlist += "Name: $($item.Display_Name); Uninstall Path: $($item.Uninstall_Path)`n"}

            $logs += "$(Get-Date) - Existing antivirus detected:`n$avlist"
            $kill++
        }
        # Otherwise, just add the fact that none were found to the previous log.
        else{$logs[-1] += " None found."}

        # Don't bother processing *anything* if there are snags.
        if($kill -eq 0)
        {
            $logs += "$(Get-Date) - Checking if the ESET Agent is already installed."
            if(!(Check-EsetAgent))
            {
                $logs += "$(Get-Date) - Checking if the ESET Agent installer already exists."
                if(!(Test-Path $agentPath))
                {
                    $logs += "$(Get-Date) - Fetching the ESET Agent download link."
                    try{$link = Get-EsetLink}
                    catch
                    {
                        $logs += "$(Get-Date) - There was an unexpected error when fetching the ESET Agent download link: $($error[0])"
                        $kill++
                    }

                    # Only process if Get-EsetLink actually returned a value.
                    if($link -ne $null)
                    {
                        $logs += "$(Get-Date) - Downloading ESET Agent installer."
                        try {wget -uri $link -OutFile $agentPath}
                        catch
                        {
                            $logs += "$(Get-Date) - There was an error when attempting to download the ESET Agent installer: $($error[0])"
                            $kill++
                        }
                    }
                }

                $logs += "$(Get-Date) - Installing the ESET Agent."
                try{Install-EsetAgent}
                catch
                {
                    $logs += "$(Get-Date) - There was a problem installing the ESET Agent: $($error[0])"
                    $kill++
                }

            }
            else {$logs += "$(Get-Date) - ESET Agent already installed."}

            # DO NOT install AV if there's no Agent.
            if(Check-EsetAgent)
            {
                $logs += "$(Get-Date) - Checking whether the machine is a server or workstation."
                if(Check-MachineRole -eq "server")
                {
                    # Variable to be used for ticket creation (see End block)
                    $role = "srv"

                    $logs += "$(Get-Date) - Installing ESET File Security."
                    try{Install-EsetFS}
                    catch
                    {
                        $logs += "$(Get-Date) - There was a problem installing ESET File Security: $($error[0])"
                        $kill++
                    }
                }
                elseif(Check-MachineRole -eq "workstation")
                {
                    # Variable to be used for ticket creation (see End block)
                    $role = "wks"
                    
                    $logs += "$(Get-Date) - Installing ESET Endpoint."
                    try{Install-EsetEndpiont}
                    catch
                    {
                        $logs += "$(Get-Date) - There was a problem installing ESET Endpoint: $($error[0])"
                        $kill++
                    }
                }
                else
                {
                    $logs += "$(Get-Date) - Installed OS determined to not be Windows or otherwise not recognized; cancelling procedure."
                    $kill++
                }
            }
            else
            {
                $logs += "$(Get-Date) - Cannot install an ESET AV - ESET Agent is not installed."
                $kill++
            }
        }

        $logs += "$(Get-Date) - Total killswitch triggers: $kill"
    }

    End
    {
        $logs += "$(Get-Date) - End of log for Install-Eset."
        
        # Update (or create) the log file for this function with the contents of the $logs array.
        $LogPath = "$env:windir\Temp\InstallEset_IITS.txt"
    	foreach($log in $logs)
    	{"$log" | Out-File -FilePath $LogPath -Force -Append}
	    
        # (Optional) Send an e-mail to MSAlarm to create a ticket. Even if called, only triggers if there was actually a problem.
        if(($ticket) -and ($kill -ne 0))
        {
            # Find all ESET log files and add them to $logfiles.
            [array]$logfiles = $null
            [array]$files = Get-ChildItem $LogRoot | % Name
            foreach($file in $files)
            {
                if($file -match "eset.*_IITS"){$logfiles += $file}
            }

            # Sends an alert to MSAlarm, which will make a ticket. DOES include the log files as attachments.
            Email-MSalarm -Body "ESET install failed on $(Get-KaseyaMachineID). See Documents for attached logs." -Attachment $logfiles
        }
    }
}

<#
.Synopsis
   Checks if ESET Agent is installed, returns TRUE if so.
.DESCRIPTION
   Test if the ESET Agent executable exists in the typical installation directory. If it exists, return TRUE, else FALSE.
.EXAMPLE
   Check-EsetAgent

   True
.OUTPUTS
   Boolean
#>
function Check-EsetAgent
{
    Test-Path "$env:ProgramFiles\ESET\RemoteAdministrator\Agent\ERAAgent.exe"
}

<#
.Synopsis
   Match an organization and a Windows OS architecture (32 or 64) to download an installer. Only works on a single machine at a time.
.DESCRIPTION
   Determine the root org (groupName) based on the local Kaseya ID (machName). Determine the OS architecture (machOS) of the machine this script is run on (which will be the same machine in machName). Match machOrg and machOS against key EsetKey.csv to get a Dropbox download link to a company-specific ESET Agent installer.
.EXAMPLE
   Get-EsetLink

   http://www.dropbox.com/s/[uniqueURL]/Agent_sccit_64.msi?dl=1
.OUTPUTS
   URL (string) or null
.FUNCTIONALITY
   Downloads a URL link to an installer.
#>
function Get-EsetLink
{
    [OutputType([String])]
    Param
    ()
    
    Begin
    {
        # Initialize the killswitch.
        $kill = 0

        # Initialize the logs array.
        $logs = @("===================","$(Get-Date) - Beginning Get-EsetLink.")

        # Initialize the URL varible.
        $orgLink = $null
        
        # Verify the local PowerShell version is sufficient to actually run the main function.
        $logs += "$(Get-Date) - Checking installed PowerShell version for compatibility."
        if (Check-PSVersion)
        {
            $logs += "$(Get-Date) - PowerShell version insufficient to run Get-EsetLink. Current version: $($PSVersionTable.PSVersion.Major)"
            $kill++
        }

        # Verify EsetKey actually exists where it is supposed to be.
        $logs += "$(Get-Date) - Checking for existence of EsetKey.csv."
        if(!(Test-Path C:\IITS_Scripts\EsetKey.csv))
        {
            $logs += "$(Get-Date) - EsetKey.csv does not exist! Download fresh copy of EsetKey.csv required."
            $kill++
            Write-Host "EsetKey.csv does not exist! Please download a fresh copy of EsetKey.csv." -BackgroundColor Black -ForegroundColor Red
        }       
    }

    Process
    {
        # This is the name of the machine. It will be converted into the name of the org and then checked against a spreadsheet/key.
        $logs += "$(Get-Date) - Pulling full Kaseya ID."
        $machName = Get-KaseyaMachineID
        
        # RegEx the machine name to extract the group name. DO NOT output the actual match result.
        $logs += "$(Get-Date) - Pulling group name from machine name '$machName'."
        Try
        {
            $machName -match '[\w-]+$' | Out-Null
            [String]$groupName = $matches[0]
        }
        Catch
        {
            $logs += "$(Get-Date) - Could not run RegEx on given machine name '$machName'. Error: $($Error[0])"
            $kill++
        }

        # Get the OS architecture of the target (Windows) machine. DO NOT output the actual match result.
        Try
        {
            $logs += "$(Get-Date) - Fetching OS architecture."
            (Get-WmiObject Win32_OperatingSystem).OSArchitecture -match '\d+' | Out-Null
            [Int]$machOS=$matches[0]
        }
        Catch
        {
            $logs += "$(Get-Date) - Could not determine OS architecture."
            $kill++
        }
        
        # If there have been any issues (which would cause $kill /= 0), don't process the rest of the Process block.
        if($kill -eq 0)
        {
            # Import the key and search for the group and OS architecture. Save the result to a container.
            Try{$orgLink = (Import-Csv "C:\IITS_Scripts\EsetKey.csv" | where{$_.machOrg -eq $groupName} | where{$_.machOS -eq $machOS} | % link)}
            Catch{$logs += "$(Get-Date) - Error while importing EsetKey.csv: $($Error[0])"}
        }
        else {$logs += "$(Get-Date) - Skipping remaining Process block. Killswitch trigger count = $kill"}
        
        # Output the container with the ESET link.
        echo $orgLink
    }

    End
    {
        $logs += "$(Get-Date) - End of log for Get-EsetLink."
        
        # Update (or create) the log file for this function with the contents of the $logs array.
        $LogPath = "$env:windir\Temp\GetEsetLink_IITS.txt"
        foreach($log in $logs){"$log" | Out-File -FilePath $LogPath -Force -Append}
    }
}

<#
.Synopsis
   Installs the ESET Agent.
.DESCRIPTION
   Looks for a specific ESET Agent installer in an expected directory and attempts to execute it silently. DOES NOT check if ESET Agent already installed (will cause repair or in-place upgrade if so).
#>
function Install-EsetAgent
{
    [CmdletBinding()]
    Param
    ()

    Begin
    {
        # Initialize the logs array.
	    $logs=@("===================","$(Get-Date) - Beginning Install-EsetAgent.")

        # ESET Agent installer path.
        $AgentPath = "C:\IITS_Mgmt\Temp\EsetAgent.msi"
    }
    
    Process
    {
        $logs += "$(Get-Date) - Verifying ESET Agent installer exists."
        if(Test-Path $AgentPath)
        {
            $logs += "$(Get-Date) - Attempting to install ESET Agent."
            try
            {
                msiexec /i $AgentPath /qn REBOOT="ReallySuppress"

                # Give the procedure time to actually install.
                sleep 30

                # Verify if agent was installed or not.
                if(Check-EsetAgent){$logs += "$(Get-Date) - ESET Agent is installed."}
            }
            catch{$logs += "$(Get-Date) - Could not install ESET Agent. Error: $($error[0])"}
        }
        else{$logs += "$(Get-Date) - Expected ESET Agent installer does not exist."}
        
    }

    End
    {
        $logs += "$(Get-Date) - End of log for Install-EsetAgent."
        
        # Update (or create) the log file for this function with the contents of the $logs array.
    	$LogPath = "$env:windir\Temp\InstallEsetAgent_IITS.txt"
    	foreach($log in $logs)
    	{"$log" | Out-File -FilePath $LogPath -Force -Append}
    }
}

<#
.Synopsis
   Checks if ESET Endpoint is installed, returns TRUE if so.
.DESCRIPTION
   Test if the ESET Endpoint executable exists in the typical installation directory. If it exists, return TRUE, else FALSE.
.EXAMPLE
   Check-EsetEndpoint

   True
.OUTPUTS
   Boolean
#>
function Check-EsetEndpoint
{
    Test-Path "$env:ProgramFiles\ESET\ESET Endpoint Antivirus\egui.exe"
}

<#
.Synopsis
   Downloads the ESET Endpoint installer.
.DESCRIPTION
   Downloads a client-generic but OS-specific installer for ESET Endpoint. The downloaded file is always named EsetEndpoint.msi.
#>
function Download-EsetEndpoint
{
    [CmdletBinding()]
    Param
    ()

    Begin
    {
        # Initialize the logs array.
	    $logs=@("===================","$(Get-Date) - Beginning Download-EsetEndpoint.")

        # Declaring OS architecture container here so that I can check for (lack of) content later.
        [Int]$OS = $null

        # Download link for 32-bit Endpoint
        $url32 = "http://www.dropbox.com/s/69lmk8ug7bnagyv/endpoint_32.msi?dl=1"
        
        # Download link for 32-bit Endpoint
        $url64 = "http://www.dropbox.com/s/12av22ml3q5jxis/endpoint_64.msi?dl=1"

        # File name and path for downloaded installer
        $out = "C:\IITS_Mgmt\Temp\EsetEndpoint.msi"
    }
    
    Process
    {
        # Get the OS architecture of the target (Windows) machine. DO NOT output the actual match result.
        Try
        {
            $logs += "$(Get-Date) - Fetching OS architecture."
            (Get-WmiObject Win32_OperatingSystem).OSArchitecture -match '\d+' | Out-Null
            $OS=$matches[0]
        }
        Catch{$logs += "$(Get-Date) - Could not determine OS architecture."}

        # Download the installer that matches the OS architecture.
        if($OS)
        {
            $logs += "$(Get-Date) - Attempting to download installer that matches OS architecture."
            
            if(($OS -eq 32) -or ($OS -eq 86))
            {
                try{wget -uri $url32 -outfile $out}
                catch{$logs += "$(Get-Date) - There was a problem downloading the installer: $($error[0])"}
            }
            elseif($OS -eq 64)
            {
                try{wget -uri $url64 -outfile $out}
                catch{$logs += "$(Get-Date) - There was a problem downloading the installer: $($error[0])"}
            }
            else{$logs += "$(Get-Date) - OS architecture misread. Given value: $OS"}
        }
    }
    
    End
    {
        $logs += "$(Get-Date) - End of log for Download-EsetAgent."
        
        # Update (or create) the log file for this function with the contents of the $logs array.
    	$LogPath = "$env:windir\Temp\DownloadEsetEndpoint_IITS.txt"
    	foreach($log in $logs){"$log" | Out-File -FilePath $LogPath -Force -Append}
    }
}

<#
.Synopsis
   Installs ESET Endpoint.
.DESCRIPTION
   Looks for a specific ESET Endpoint installer in an expected directory and attempts to execute it silently. DOES NOT check if ESET Endpoint already installed (will cause repair or in-place upgrade if so).
#>
function Install-EsetEndpoint
{
    [CmdletBinding()]
    Param
    ()

    Begin
    {
        # Initialize the logs array.
	    $logs=@("===================","$(Get-Date) - Beginning Install-EsetEndpoint.")

        # ESET Endpoint installer path.
        $EndpointPath = "C:\IITS_Mgmt\Temp\EsetEndpoint.msi"
    }
    
    Process
    {
        $logs += "$(Get-Date) - Verifying ESET Endpoint installer exists."
        if(!(Test-Path $EndpointPath))
        {
            $logs += "$(Get-Date) - Downloading ESET Endpoint installer."
            try{Download-EsetEndpoint}
            catch{$logs += "$(Get-Date) - There was an error downloading the installer: $($error[0])"}
        }
        
        if(Test-Path $EndpointPath)
        {
            $logs += "$(Get-Date) - Attempting to install ESET Endpoint."
            try
            {
                msiexec /i $EndpointPath /qn REBOOT="ReallySuppress"

                # Give the procedure time to actually install.
                sleep 30

                # Verify if Endpoint was installed or not.
                if(Check-EsetEndpoint){$logs += "$(Get-Date) - ESET Endpoint is installed."}
            }
            catch{$logs += "$(Get-Date) - Could not install ESET Endpoint. Error: $($error[0])"}
        }
    }

    End
    {
        $logs += "$(Get-Date) - End of log for Install-EsetEndpoint."
        
        $LogPath = "$env:windir\Temp\InstallEsetEndpoint_IITS.txt"
    	foreach($log in $logs)
    	{"$log" | Out-File -FilePath $LogPath -Force -Append}
    }
}

<#
.Synopsis
   Checks if ESET File Security is installed, returns TRUE if so.
.DESCRIPTION
   Test if the ESET File Security executable exists in the typical installation directory. If it exists, return TRUE, else FALSE.
.EXAMPLE
   Check-EsetFS

   True
.OUTPUTS
   Boolean
#>
function Check-EsetFS
{
    Test-Path "$env:ProgramFiles\ESET\ESET File Security\egui.exe"
}

<#
.Synopsis
   Downloads the ESET File Security installer.
.DESCRIPTION
   Downloads a client-generic but OS-specific installer for ESET File Security. The downloaded file is always named EsetFS.msi.
#>
function Download-EsetFS
{
    [CmdletBinding()]
    Param
    ()

    Begin
    {
        # Initialize the logs array.
	    $logs=@("===================","$(Get-Date) - Beginning Download-EsetFS.")

        # Declaring OS architecture container here so that I can check for (lack of) content later.
        [Int]$OS = $null

        # Download link for 32-bit File Security
        $url32 = "http://www.dropbox.com/s/tvbbjr13k6fh5j6/file-security_32.msi?dl=1"
        
        # Download link for 32-bit File Security
        $url64 = "http://www.dropbox.com/s/yr7e5r66hdjgduq/file-security_64.msi?dl=1"

        # File name and path for downloaded installer
        $out = "C:\IITS_Mgmt\Temp\EsetFS.msi"
    }
    
    Process
    {
        # Get the OS architecture of the target (Windows) machine. DO NOT output the actual match result.
        Try
        {
            $logs += "$(Get-Date) - Fetching OS architecture."
            (Get-WmiObject Win32_OperatingSystem).OSArchitecture -match '\d+' | Out-Null
            $OS=$matches[0]
        }
        Catch{$logs += "$(Get-Date) - Could not determine OS architecture."}

        # Download the installer that matches the OS architecture.
        if($OS)
        {
            $logs += "$(Get-Date) - Attempting to download installer that matches OS architecture."
            
            if(($OS -eq 32) -or ($OS -eq 86))
            {
                try{wget -uri $url32 -outfile $out}
                catch{$logs += "$(Get-Date) - There was a problem downloading the installer: $($error[0])"}
            }
            elseif($OS -eq 64)
            {
                try{wget -uri $url64 -outfile $out}
                catch{$logs += "$(Get-Date) - There was a problem downloading the installer: $($error[0])"}
            }
            else{$logs += "$(Get-Date) - OS architecture misread. Given value: $OS"}
        }
    }
    
    End
    {
        $logs += "$(Get-Date) - End of log for Download-EsetFS."
        
        # Update (or create) the log file for this function with the contents of the $logs array.
    	$LogPath = "$env:windir\Temp\DownloadEsetFS_IITS.txt"
    	foreach($log in $logs){"$log" | Out-File -FilePath $LogPath -Force -Append}
    }
}

<#
.Synopsis
   Installs the ESET File Security.
.DESCRIPTION
   Looks for a specific ESET File Security installer in an expected directory and attempts to execute it silently. DOES NOT check if ESET File Security already installed (will cause repair or in-place upgrade if so).
#>
function Install-EsetFS
{
    [CmdletBinding()]
    Param
    ()

    Begin
    {
        # Initialize the logs array.
	    $logs=@("===================","$(Get-Date) - Beginning Install-EsetFS.")

        # ESET File Security installer path.
        $FSPath = "C:\IITS_Mgmt\Temp\EsetFS.msi"

        # Killswitch
        $kill = 0
    }
    
    Process
    {
        $logs += "$(Get-Date) - Checking for the installer."
        if(!(Test-Path $FSPath))
        {
            $logs += "$(Get-Date) - Downloading the installer."
            try{Download-EsetFS}
            catch
            {
                $logs += "$(Get-Date) - There was a problem downloading the installer: $($error[0])"
                $kill++
            }
        }
        
        if($kill -eq 0)
        {    
            $logs += "$(Get-Date) - Attempting to install ESET File Security."
            try
            {
                msiexec /i $FSPath /qn REBOOT="ReallySuppress"

                # Give the procedure time to actually install.
                sleep 30

                # Verify if FS was installed or not.
                if(Check-EsetFS){$logs += "$(Get-Date) - ESET File Security is installed."}
            }
            catch{$logs += "$(Get-Date) - Could not install ESET File Security. Error: $($error[0])"}
        }
        
    }

    End
    {
        $logs += "$(Get-Date) - End of log for Install-EsetFS."
        
        $LogPath = "$env:windir\Temp\InstallEsetFS_IITS.txt"
    	foreach($log in $logs)
    	{"$log" | Out-File -FilePath $LogPath -Force -Append}
    }
}

<#
.Synopsis
   Determines if Windows machine is server or workstation. Returns "server", "workstation", or "cannot determine".
.DESCRIPTION
   Looks up the host machine's OS and checks it against a list of known OSes. Recognized Windows OSes will return "server" or "workstation"; anything else (including unrecognized Windows OSes) will return "cannot determine".
.EXAMPLE
   Executed on Windows Server 2012 R2:
    Check-MachineRole

    server
.EXAMPLE
   Executed on Windows 10:
    Check-MachineRole

    workstation
.EXAMPLE
   Executed on a Mac:
    Check-MachineRole

    cannot determine
.OUTPUTS
   string ("server","workstation", or "cannot determine")
.FUNCTIONALITY
   Primarily used as a helper function when there is a need to differentiate between a server and a workstation host without foreknowledge, 
   such as running one of two installers (server or workstation version) without having direct access to or knowledge of either.
#>
function Check-MachineRole
{
    [Alias("cmr","Get-Role")]
    [OutputType([String])]
    Param
    ()

    Process
    {
        # Get the OS proper name.
        $os = (gwmi -Class Win32_OperatingSystem).Caption

        if($os -match "Windows")
        {
            if($os -match "Server"){echo "server"}
            else{echo "workstation"}
        }
        else{echo "cannot determine"}
    }
}

<#
.Synopsis
   Checks if known antivirus programs are currently installed on host machine. Returns any found and their uninstall string, else "none found".
.DESCRIPTION
   Checks for the existence of various known AVs in host machine's Add/Remove. If it detects any, returns a table of products and related uninstall strings. Otherwise, it returns "none found".
.EXAMPLE
   Computer with no AV:
   Check-InstalledAV

   none found
.EXAMPLE
   Computer with AVG:
   Check-InstalledAV

   AVG Antivirus
.EXAMPLE
   Computer with AVG *and* ESET:
   Check-InstalledAv

   AVG
   ESET Endpoint Antivirus
#>
function Check-InstalledAv
{
    [Alias("Check-AV")]
    [OutputType([string])]
    Param
    ()

    Begin
    {
        # Initialize the $avs container.
        [string]$avs = $null

        # List of known AV products.
        $KnownAV = @("Avast","AVG","ESET Endpoint","ESET File Security","Kaspersky","Norton","Symantec")
    }

    Process
    {
        # Go through the local machine's Add/Remove list and check for anything with they keywords given in $KnownAVs. Add any found to $avs.
        foreach($item in $KnownAV)
        {
            $avs += Get-InstalledPrograms | where{($_.display_name -match $item)}
        }
    }

    End
    {
        if($avs){echo $avs}
        else{echo "none found"}
    }
}