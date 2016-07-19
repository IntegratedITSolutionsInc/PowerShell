Function Lock-User-Computer
{
<#
.DESCRIPTION
	Function to Lock a user's computer
.SYNOPSIS
	Function to Lock your computer...if already locked this will do nothing
#>
	
$signature = @"
[DllImport("user32.dll", SetLastError = true)]
public static extern bool LockWorkStation();
"@

$LockComputer = Add-Type -memberDefinition $signature -name "Win32LockWorkStation" -namespace Win32Functions -passthru
$LockComputer::LockWorkStation() | Out-Null
}
