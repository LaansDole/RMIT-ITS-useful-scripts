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

# Replace with the file path that you want
$filePath = "C:\Excel\Excel File.xlsx"

$ExcelObj = New-Object -comobject Excel.Application

$ExcelWorkBook = $ExcelObj.Workbooks.Open($filePath)

# Check if the workbook is read-only
if ($ExcelWorkBook.ReadOnly) {

    Write-Host "The work book is in read-only" -ForegroundColor Yellow
    # Close the ExcelWorkBook
    $ExcelWorkBook.Close()

    # Change the file's attributes to Normal
    Set-ItemProperty -Path $filePath -Name IsReadOnly -Value $false

    # Reopen the ExcelWorkBook
    $ExcelWorkBook = $ExcelObj.Workbooks.Open($filePath)
}


# Get the number of worksheets
$sheetCount = $ExcelWorkBook.Sheets.Count

# Print the number of worksheets
Write-Host "The workbook contains $sheetCount worksheets." -ForegroundColor Yellow

foreach ($sheet in $ExcelWorkBook.Sheets) {

    Clear-Host
        
    $ExcelWorkSheet = $ExcelWorkBook.Sheets.Item($sheet.Name)
    $currentsheet = $ExcelWorkSheet.Name
    Write-Host "Current Worksheet: $currentsheet" -ForegroundColor Yellow -BackgroundColor DarkGreen

    $usedRange = $ExcelWorkSheet.UsedRange

    # Iterate over each row in the used range
    for ($row = 2; $row -le $usedRange.Rows.Count; $row++) 
    {
        # Get the values of column A
        $hostname = $usedRange.Cells.Item($row, 1).Value2
            
        # The end of the worksheet
        if ($hostname -eq $null) { break; }

        $sn = $usedRange.Cells.Item($row, 2).Value2
        $tagcode = $usedRange.Cells.Item($row, 3).Value2

        Write-Host "Hostname: $hostname, Serialnumber: $sn, Tag Code: $tagcode"

        try {
            $computerDescription = Get-ADComputer $hostname -Properties Description | Select Description
            if ($computerDescription.Description -ne $null) {
                Write-Host "$computerDescription`n" -BackgroundColor Green -ForegroundColor Yellow
                # Update column I with the device description
                $usedRange.Cells.Item($row, 9).Value2 = "Available"
            } else {
                Write-Host "$hostname does not have description" -BackgroundColor Red -ForegroundColor White
                # Update column I with the device description
                $usedRange.Cells.Item($row, 9).Value2 = "Not Available"
            }


        } catch {
            Write-Host "$hostname is not on AD`n" -BackgroundColor Red -ForegroundColor White
        }
    }

    Write-Host "End of Worksheet: $currentsheet" -ForegroundColor Yellow -BackgroundColor DarkGreen
    pause

}


$ExcelWorkBook.Save()
$ExcelWorkBook.Close()
$ExcelObj.Quit()

pause
