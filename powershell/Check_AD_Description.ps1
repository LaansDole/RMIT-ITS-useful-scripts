#===========================================================================
#region Run script as elevated admin and unrestricted executionpolicy
#===========================================================================

# Check if running as an administrator
if (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
    # Not running as an administrator, so relaunch as administrator
    Start-Process powershell.exe "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`"" -Verb RunAs
    exit
}
Write-Host "Running with elevated privileges."
# We are running as an administrator, so change the title and background colour to indicate this
$Host.UI.RawUI.WindowTitle = $myInvocation.MyCommand.Definition + "(Elevated)";
$Host.UI.RawUI.BackgroundColor = "DarkBlue";
Clear-Host;

#endregion

do {
    $hostname=Read-Host "Input device's hostname"
    $computerDescription = Get-ADComputer $hostname -Properties Description | Select Description
    Write-Host "$computerDescription`n" -BackgroundColor Green -ForegroundColor Yellow
} while(1)
