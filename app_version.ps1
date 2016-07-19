<#
.Synopsis
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
