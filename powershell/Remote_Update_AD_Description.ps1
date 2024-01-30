function Check-RSATComponents {
    # Check and Install RSAT: Server Manager - Dependencies
    $feature = "Rsat.ServerManager.Tools~~~~0.0.1.0"
    $installed = Get-WindowsCapability -Name $feature -Online | Select-Object -Property State

    if ($installed.State -eq "Installed") {
        Write-Host "RSAT: Server Manager is already installed" -ForegroundColor Yellow -BackgroundColor DarkGreen
    } else {
        Write-Host "RSAT: Server Manager will be now installed" -ForegroundColor Yellow
        Add-WindowsCapability -Online -Name $feature
    }

    # Check and Install RSAT: Active Directory Certificate Services - Dependencies
    $feature = "Rsat.CertificateServices.Tools~~~~0.0.1.0"
    $installed = Get-WindowsCapability -Name $feature -Online | Select-Object -Property State

    if ($installed.State -eq "Installed") {
        Write-Host "RSAT: Active Directory Certificate is already installed" -ForegroundColor Yellow -BackgroundColor DarkGreen
    } else {
        Write-Host "RSAT: Active Directory Certificate will be now installed" -ForegroundColor Yellow
        Add-WindowsCapability -Online -Name $feature
    }

    # Check and Install RSAT: Active Directory Domain Servers and Lightweight Directory Services Tools 
    $feature = "Rsat.ActiveDirectory.DS-LDS.Tools~~~~0.0.1.0"
    $installed = Get-WindowsCapability -Name $feature -Online | Select-Object -Property State

    if ($installed.State -eq "Installed") {
        Write-Host "RSAT: Active Directory Domain is already installed" -ForegroundColor Yellow -BackgroundColor DarkGreen
    } else {
        Write-Host "RSAT: Active Directory Domain will be now installed" -ForegroundColor Yellow
        try {
            Add-WindowsCapability -Online -Name $feature
        } catch {
            Write-Host "Ignore Capability Not Present, false alarm" -ForegroundColor Yellow -BackgroundColor DarkGreen
        }
    }

    DISM.exe /Online /Get-CapabilityInfo /CapabilityName:Rsat.ActiveDirectory.DS-LDS.Tools~~~~0.0.1.0
}

function Update-AD-Description (
        [string]$computername, 
        [string]$tagcode, 
        [string]$vnumber, 
        [string]$serialnumber,
        [string]$department
    ) {

    $device_description = "Tagcode:" + $tagcode + "; SN:"+ $serialnumber + "; Department:" + $department + "; Updated by:" + $vnumber
    
    return $device_description
}

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

## Comment if you already have RSAT
Check-RSATComponents
Clear-Host
do {
    $isError = $false
    # Get user credentials
    do {
        $vnumber = Read-Host "Input your Vnumber"
    } 
    while ($vnumber.Length -eq 0)
    
    $vnumber = $vnumber.ToUpper()
    $credentials = Get-Credential -Message "Use your PAM/MGR account: " -UserName $vnumber
    
    try {
        $computerObject = Get-ADComputer $env:COMPUTERNAME -Properties * -Credential $credentials | Select *
    } catch {
        Write-Host "Incorrect Credential, try again" -ForegroundColor White -BackgroundColor Red
        $isError = $true
        Start-Sleep 2
        continue
    }
}
while ($isError)

do {
    do {
    $hostname=Read-Host "Input device's Host name"
    } while ($hostname.Length -eq 0)

    do {
    $tagcode=Read-Host "Input device's Tag Code"
    } while ($tagcode.Length -eq 0)

    do {
    $sn=Read-Host "Input device's Serial Number"
    } while ($sn.Length -eq 0)

    do {
    $department=Read-Host "Input device's Department"
    } while ($department.Length -eq 0)

    $device_description = Update-AD-Description -computername $hostname -tagcode $tagcode -vnumber $vnumber -serialnumber $sn -department $department
    Set-ADComputer -Identity $hostname -Description $device_description -Credential $credentials
    $description = Get-ADComputer $hostname -Properties Description | Select Description
    Write-Host "Device Description: " + $description

    if ($description -ne $null) {
        Write-Host "The description has been sucessfull to updated!" -ForegroundColor Yellow -BackgroundColor DarkGreen

    } else {
        Write-Host "Description $device_description is not set"
    }
    
    Start-Sleep -Seconds 3
    Clear-Host
} while (1)
