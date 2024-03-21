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

# Function Show List Departments
function Get-Department {
    param (
        [string]$Title = 'RMITs Departments'
    )

    do {
        $error = $false

        Write-Host "================ Department ================"
        Write-Host " "
        Write-Host "01: Press '1' for SCD"
        Write-Host "02: Press '2' for TBS"
        Write-Host "03: Press '3' for SSET"
        Write-Host "04: Press '4' for SEUP"
        Write-Host "05: Press '5' for Marketing"
        Write-Host "06: Press '6' for Student Recruiment"
        Write-Host "07: Press '7' for GMKT, Digital & Recruitment"
        Write-Host "08: Press '8' for International"
        write-host "09: Press '9' for Transnational Security Centre"
        Write-Host "10: Press '10' for Human Resources Vietnam"
        Write-Host "11: Press '11' for Finance & Governance"
        Write-Host "12: Press '12' for Legal & Compliance"
        Write-Host "13: Press '13' for Government Affairs"
        Write-Host "14: Press '14' for IT Services"
        Write-Host "15: Press '15' for ITS-Loan"
        Write-Host "16: Press '16' for OHS & Security"
        Write-Host "17: Press '17' for Property & Campus Operations"
        Write-Host "18: Press '18' for Office for Research & Innovation"
        Write-Host "19: Press '19' for PVC Office"
        Write-Host "20: Press '20' for Dean of Students"
        Write-Host "21: Press '21' for Career Alumni Industry Relation"
        Write-Host "22: Press '22' for Student & Family Connect"
        Write-Host "23: Press '23' for Student Life"
        Write-host "24: Press '24' for Student Success"
        Write-Host "25: Press '25' for Academic Experience & Success"
        Write-Host "26: Press '26' for Wellbeing"
        Write-Host "27: Press '27' for Communications"
        Write-Host "28: Press '28' for Academic Registrar's Group"
        Write-Host "29: Press '29' for Library & Digital Services"
        Write-Host "30: Press '30' for Market Research"
        Write-Host "31: Press '31' for Student Admissions"
        Write-Host "32: Press '32' for Events"
        Write-Host "33: Press '33' for TSC"

        $selection = Read-Host "Please make a selection to choose the department, '0' to choose again"
        
        try {
            $selection = [int]$selection
        } catch {
            Write-Host "Invalid Department!" -ForegroundColor White -BackgroundColor Red
            Start-Sleep 2
            $error = $true
            continue
        }

        if (($selection -lt 0) -or ($selection -gt 33)) {
            Write-Host "Invalid Department!" -ForegroundColor White -BackgroundColor Red
            Start-Sleep 2
            $error = $true
            continue

        } elseif ($selection -eq '0') {
            $error = $true
            continue

        } else  {
            $error = $false
            switch ($selection)
            {
                '1' {
                $department = 'VNM|School of Communication & Design'
            
                }
                '2' {
                $department = 'VNM|The Business School'
            
                }
                '3' {
                $department = 'VNM|SSET'
            
                }
                '4' {
                $department = 'VNM|SEUP'
            
                }
                '5' {
                $department = 'VNM|EXP|EXP_Marketing & EXP_MarketingWeb'
            
                }
                '6' {
                $department = 'VNM|EXP|EXP_Student Recruiment'
            
                }
                '7' {
                $department = 'VNM|EXP|GMKT, Digital & Recruitment'
            
                }
                '8' {
                $department = 'VNM|EXP|International'
            
                }
                '9' {
                $department = 'VNM|TSC|Transnational Security Centre'
            
                }
                '10' {
                $department = 'OPS|Human Resources Vietnam'
            
                }
                '11' {
                $department = 'VNM|F&G|Finance & Governance'
            
                }
                '12' {
                $department = 'VNM|F&G|Legal & Compliance'
            
                }
                '13' {
                $department = 'VNM|GD|Government Affairs'
            
                }
                '14' {
                $department = 'VNM|Ops|IT Services'
            
                }
                '15' {
                $department = 'ITS-Loan'
            
                }
                '16' {
                $department = 'VNM|Ops|OHS & Security'
            
                }
                '17' {
                $department = 'VNM|Ops|Property & Campus Operations'
            
                }
                '18' {
                $department = 'VNM|R&I|Office for Research & Innovation'
            
                }
                '19' {
                $department = 'VNM|PVC|PVC Office'
            
                }
                '20' {
                $department = 'Dean of Students'
            
                }
                '21' {
                $department = 'VNM|SEP|Career Alumni Industry Relation'
            
                }
                '22' {
                $department = 'VNM|Student & Family Connect'
            
                }
                '23' {
                $department = 'VNM|Student Life'
            
                }
                '24' {
                $department = 'VNM|Student Success'
            
                }
                '25' {
                $department = 'VNM|SEP|Academic Experience & Success'
            
                }
                '26' {
                $department = 'VNM|SEP|Wellbeing'
            
                }
                '27' {
                $department = 'VNM|UC|Communications, Vietnam'
            
                }
                '28' {
                $department = 'VNM|SEP|Academic Registrar Group'
            
                }
                '29' {
                $department = 'VNM|SEP|Library & Digital Services'
            
                }
                '30' {
                $department = 'VNM|EXP|Market Research'
            
                }
                '31' {
                $department = 'VNM|EXP|Student Admissions'
            
                }
                '32' {
                $department  = 'VNM|EXP|Events'
            
                }
                '33' {
                $department  = 'VNM|TSC|Transactional Security Centre'
            
                }
            }
            break

        }

    } while ($error)

    return $department
}

function Update-AD-Description (
        [string]$computername, 
        [string]$tagcode, 
        [string]$vnumber,
        [string]$vstaff,
        [string]$serialnumber,
        [string]$department,
        [string]$location
    ) {

    if ($location.Length -ne 0) {
        $device_description = "Tagcode:" + $tagcode + "; SN:"+ $serialnumber + "; Location:" + $location + "; Used by:" + $vstaff + "; Updated by:" + $vnumber
    } else {
        $device_description = "Tagcode:" + $tagcode + "; SN:"+ $serialnumber + "; Department:" + $department + "; Used by:" + $vstaff + "; Updated by:" + $vnumber
    }

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
    $sn=Read-Host "Input device's Serial Number"
    } while ($sn.Length -eq 0)

    do {
    $tagcode=Read-Host "Input device's Tag Code"
    } while ($tagcode.Length -eq 0)

    do {
    $vstaff=Read-Host "Input device's User"
    } while ($vstaff.Length -eq 0)

    $location = Read-Host "Input device's Location (ENTER to get department instead)"

    if ($location.Length -eq 0) {
        $department = Get-Department
    }

    $device_description = Update-AD-Description -computername $hostname -tagcode $tagcode -vnumber $vnumber -vstaff $vstaff -serialnumber $sn -department $department -location $location
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
