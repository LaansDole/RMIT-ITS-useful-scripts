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
