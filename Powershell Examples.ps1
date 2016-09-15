#  Get Users based on domain
 get-recipient | where {$_.emailaddresses -match "slipstick.com"}
 
 #  Prevent 365 from automatically mapping Mailboxes that a user has permissions to
 Add-MailboxPermission -Identity ahchm@chartis.com -User npatel@chartis.com -AccessRights FullAccess -AutoMapping:$false
 
 # Set Registry via Powershell
 new-item -path HKLM:\software\motive\m-files\10.2.3920.54\client\mfofficeaddin -name outlookaddindisabled - value 0 -force
 
 # List installed apps in a table
 Get-ItemProperty HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* | Select-Object DisplayName, DisplayVersion, Publisher, InstallDate | 
 Format-Table –AutoSize
 
 #####
 
