<#
.SYNOPSIS
    Windows Installation Script with Local Account Creation and OOBE Configuration
.DESCRIPTION
    This script automates Windows installation from an ISO file, with options to create a local account
    and configure OOBE settings. It includes all the functionality of the original batch script while
    adding new features for local account creation and OOBE customization.
#>

#Requires -RunAsAdministrator

#region Functions

function Show-Menu {
    param (
        [string]$Title = 'Windows Installation'
    )
    Clear-Host
    Write-Host "=== $Title ==="
    Write-Host "This script will install Windows and configure it with your preferences."
    Write-Host "`n"
}

function Get-YesNo {
    param (
        [string]$Prompt
    )
    do {
        $response = Read-Host -Prompt "$Prompt (Y/N)"
    } while ($response -notmatch '^[YyNn]$')
    return ($response -eq 'Y' -or $response -eq 'y')
}

function Get-ValidatedInput {
    param (
        [string]$Prompt,
        [string]$DefaultValue = "",
        [switch]$IsPassword = $false,
        [switch]$IsRequired = $false
    )

    do {
        if ($IsPassword) {
            $secureInput = Read-Host -AsSecureString -Prompt $Prompt
            $inputValue = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto(
                [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($secureInput)
            )
        } else {
            $inputValue = Read-Host -Prompt $Prompt
            if ([string]::IsNullOrWhiteSpace($inputValue) -and -not $IsRequired) {
                $inputValue = $DefaultValue
            }
        }

        if ([string]::IsNullOrWhiteSpace($inputValue) -and $IsRequired) {
            Write-Host "This field is required. Please enter a value." -ForegroundColor Red
        } else {
            break
        }
    } while ($true)

    return $inputValue
}

function Get-TimeZoneList {
    [System.TimeZoneInfo]::GetSystemTimeZones() | 
    Sort-Object DisplayName | 
    ForEach-Object { 
        [PSCustomObject]@{
            Id = $_.Id
            DisplayName = $_.DisplayName
        }
    }
}

function Get-KeyboardLayoutList {
    @(
        [PSCustomObject]@{ Id = "00000409"; Name = "US" },
        [PSCustomObject]@{ Id = "00000809"; Name = "UK" },
        [PSCustomObject]@{ Id = "0000040C"; Name = "French" },
        [PSCustomObject]@{ Id = "00000407"; Name = "German" },
        [PSCustomObject]@{ Id = "0000040A"; Name = "Spanish" },
        [PSCustomObject]@{ Id = "00000816"; Name = "Portuguese" },
        [PSCustomObject]@{ Id = "00000410"; Name = "Italian" },
        [PSCustomObject]@{ Id = "0000041D"; Name = "Swedish" },
        [PSCustomObject]@{ Id = "00000406"; Name = "Danish" },
        [PSCustomObject]@{ Id = "0000041F"; Name = "Turkish" }
    )
}

function New-AutounattendXml {
    param (
        [string]$Username,
        [string]$Password,
        [string]$KeyboardLayout,
        [string]$SystemLocale,
        [string]$UserLocale,
        [string]$UILanguage,
        [string]$TimeZone
    )

    $passwordXml = ""
    if (-not [string]::IsNullOrWhiteSpace($Password)) {
        $passwordXml = @"
        <Password>
            <Value>$Password</Value>
            <PlainText>false</PlainText>
        </Password>
"@
    }

    $xmlContent = @"
<?xml version="1.0" encoding="utf-8"?>
<unattend xmlns="urn:schemas-microsoft-com:unattend">
    <settings pass="oobeSystem">
        <component name="Microsoft-Windows-Shell-Setup" processorArchitecture="amd64" publicKeyToken="31bf3856ad364e35" language="neutral" versionScope="nonSxS" xmlns:wcm="http://schemas.microsoft.com/WMIConfig/2002/State" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
            <OOBE>
                <HideEULAPage>true</HideEULAPage>
                <HideOEMRegistrationScreen>true</HideOEMRegistrationScreen>
                <HideOnlineAccountScreens>true</HideOnlineAccountScreens>
                <HideWirelessSetupInOOBE>true</HideWirelessSetupInOOBE>
                <NetworkLocation>Home</NetworkLocation>
                <ProtectYourPC>1</ProtectYourPC>
                <SkipUserOOBE>true</SkipUserOOBE>
                <SkipMachineOOBE>true</SkipMachineOOBE>
            </OOBE>
            <UserAccounts>
                <LocalAccounts>
                    <LocalAccount wcm:action="add">
                        $passwordXml
                        <Description>Local Admin Account</Description>
                        <DisplayName>$Username</DisplayName>
                        <Group>Administrators</Group>
                        <Name>$Username</Name>
                    </LocalAccount>
                </LocalAccounts>
            </UserAccounts>
            <TimeZone>$TimeZone</TimeZone>
        </component>
        <component name="Microsoft-Windows-International-Core" processorArchitecture="amd64" publicKeyToken="31bf3856ad364e35" language="neutral" versionScope="nonSxS" xmlns:wcm="http://schemas.microsoft.com/WMIConfig/2002/State" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
            <InputLocale>$KeyboardLayout</InputLocale>
            <SystemLocale>$SystemLocale</SystemLocale>
            <UILanguage>$UILanguage</UILanguage>
            <UILanguageFallback>en-US</UILanguageFallback>
            <UserLocale>$UserLocale</UserLocale>
        </component>
    </settings>
    <settings pass="specialize">
        <component name="Microsoft-Windows-Deployment" processorArchitecture="amd64" publicKeyToken="31bf3856ad364e35" language="neutral" versionScope="nonSxS" xmlns:wcm="http://schemas.microsoft.com/WMIConfig/2002/State" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
            <RunSynchronous>
                <RunSynchronousCommand wcm:action="add">
                    <Order>1</Order>
                    <Path>cmd.exe /c "reg add `"HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\DataCollection`" /v AllowTelemetry /t REG_DWORD /d 0 /f"</Path>
                </RunSynchronousCommand>
                <RunSynchronousCommand wcm:action="add">
                    <Order>2</Order>
                    <Path>cmd.exe /c "reg add `"HKLM\SOFTWARE\Policies\Microsoft\Windows\DataCollection`" /v AllowTelemetry /t REG_DWORD /d 0 /f"</Path>
                </RunSynchronousCommand>
            </RunSynchronous>
        </component>
    </settings>
</unattend>
"@

    $outputPath = Join-Path -Path $PSScriptRoot -ChildPath "autounattend.xml"
    $xmlContent | Out-File -FilePath $outputPath -Encoding utf8
    return $outputPath
}

function Get-WindowsImageInfo {
    param (
        [string]$ImagePath
    )

    $images = @()
    $output = & dism.exe /Get-WimInfo /WimFile:$ImagePath
    $currentIndex = 0
    
    foreach ($line in $output) {
        if ($line -match "Index : (\d+)") {
            $currentIndex = $matches[1]
        } elseif ($line -match "Name : (.+)") {
            $images += [PSCustomObject]@{
                Index = [int]$currentIndex
                Name = $matches[1].Trim()
            }
        }
    }
    
    return $images
}

function Get-TargetDrive {
    $drives = Get-Volume | 
        Where-Object { $_.DriveType -eq 'Fixed' -and $_.DriveLetter -ne $null -and $_.DriveLetter -ne 'C' } |
        Select-Object @{Name="DriveLetter"; Expression={$_.DriveLetter + ":"}}, 
                      @{Name="SizeGB"; Expression={[math]::Round($_.Size / 1GB, 2)}},
                      @{Name="FreeSpaceGB"; Expression={[math]::Round($_.SizeRemaining / 1GB, 2)}} |
        Sort-Object DriveLetter

    if ($drives.Count -eq 0) {
        Write-Host "No suitable drives found. Please connect a drive and try again." -ForegroundColor Red
        return $null
    }

    Write-Host "`nAvailable Drives:"
    Write-Host "================="
    $drives | Format-Table -AutoSize | Out-Host

    $validDrives = $drives.DriveLetter -replace ':', ''
    do {
        $driveLetter = Read-Host "`nEnter the drive letter to install Windows (without colon, e.g., D)"
        $driveLetter = $driveLetter.Trim().ToUpper()
    } while ($driveLetter -notin $validDrives)

    # Check for existing Windows installation
    if (Test-Path "$($driveLetter):\Windows") {
        Write-Host "`nWARNING: Drive $($driveLetter): appears to have an existing Windows installation!" -ForegroundColor Yellow
        Write-Host "DISM will fail if the target partition is not empty/formatted." -ForegroundColor Yellow
        Write-Host "Please format drive $($driveLetter): before proceeding, or choose a different drive." -ForegroundColor Yellow
        return $null
    }

    return $driveLetter
}

#endregion

#region Main Script

try {
    # Check for admin rights
    $currentPrincipal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
    if (-not $currentPrincipal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
        Write-Host "This script must be run as Administrator." -ForegroundColor Red
        exit 1
    }

    Show-Menu

    # Check if 7z is available
    if (-not (Get-Command 7z -ErrorAction SilentlyContinue)) {
        Write-Host "7-Zip not found! Please install it and add it to your PATH." -ForegroundColor Red
        exit 1
    }

    # Cleanup previous installation
    Write-Host "Cleaning up previous installation files and DISM cache..." -ForegroundColor Cyan
    if (Test-Path "C:\WindowsInstallation") {
        Remove-Item -Path "C:\WindowsInstallation" -Recurse -Force -ErrorAction SilentlyContinue
    }
    Start-Process -FilePath "dism.exe" -ArgumentList "/Cleanup-Wim" -NoNewWindow -Wait -ErrorAction SilentlyContinue | Out-Null
    Write-Host "Cleanup complete." -ForegroundColor Green

    # Select ISO file
    Write-Host "`nPlease select the Windows ISO file..." -ForegroundColor Cyan
    $isoPath = Read-Host "Enter the full path to the Windows ISO file"
    $isoFile = Get-Item -Path $isoPath -ErrorAction SilentlyContinue
    
    if (-not $isoFile -or -not $isoFile.Exists) {
        Write-Host "The specified ISO file does not exist or is not accessible." -ForegroundColor Red
        exit 1
    }

    # Extract ISO
    Write-Host "`nExtracting ISO file..." -ForegroundColor Cyan
    $extractPath = "C:\WindowsInstallation"
    New-Item -ItemType Directory -Path $extractPath -Force -ErrorAction Stop | Out-Null
    & 7z x -y "-o$extractPath" $isoFile.FullName | Out-Null

    # Find the image file
    $wimFile = Join-Path -Path $extractPath -ChildPath "sources\install.wim"
    $esdFile = Join-Path -Path $extractPath -ChildPath "sources\install.esd"
    
    $imageFile = if (Test-Path $wimFile) { $wimFile } 
                elseif (Test-Path $esdFile) { $esdFile }
                else { $null }

    if (-not $imageFile) {
        Write-Host "No valid Windows image file found in the ISO." -ForegroundColor Red
        exit 1
    }

    # Show available editions
    Write-Host "`nAvailable Windows editions:" -ForegroundColor Cyan
    $images = Get-WindowsImageInfo -ImagePath $imageFile
    $images | Format-Table -Property @(
                                    @{Name="Index"; Expression={$_.Index}},
                                    @{Name="Edition"; Expression={$_.Name}}
                                ) -AutoSize | Out-Host

    # Select edition
    $edition = Read-Host "`nEnter the index number of the edition you want to install"

    # Get target drive
    $driveLetter = Get-TargetDrive
    if (-not $driveLetter) {
        Write-Host "No valid target drive selected. Exiting." -ForegroundColor Red
        exit 1
    }

    # Ask about local account
    $createLocalAccount = Get-YesNo -Prompt "Do you want to create a local account? (Y/N)"
    $autounattendPath = $null

    if ($createLocalAccount) {
        $username = Get-ValidatedInput -Prompt "Enter username" -IsRequired
        $password = Get-ValidatedInput -Prompt "Enter password (leave blank for no password)" -IsPassword
        $systemLocale = "en-US"
        $userLocale = "en-US"
        $uiLanguage = "en-US"

        # Select keyboard layout
        $keyboardLayouts = Get-KeyboardLayoutList
        Write-Host "`nAvailable Keyboard Layouts:" -ForegroundColor Cyan
        for ($i = 0; $i -lt $keyboardLayouts.Count; $i++) {
            Write-Host "$($i + 1). $($keyboardLayouts[$i].Name) ($($keyboardLayouts[$i].Id))"
        }
        $selectedLayout = Read-Host "`nSelect keyboard layout number (default: 1 - US)"
        if (-not [int]::TryParse($selectedLayout, [ref]$null) -or [int]$selectedLayout -lt 1 -or [int]$selectedLayout -gt $keyboardLayouts.Count) {
            $selectedLayout = 1
        }
        $keyboardLayout = $keyboardLayouts[[int]$selectedLayout - 1].Id

        # Select timezone
        $timeZones = Get-TimeZoneList
        Write-Host "`nAvailable Time Zones:" -ForegroundColor Cyan
        $timeZones | Format-Table -Property @(
                                           @{Name="ID"; Expression={$_.Id}},
                                           @{Name="Display Name"; Expression={$_.DisplayName}}
                                       ) -AutoSize | Out-Host
        $timeZone = Read-Host "`nEnter the timezone ID (e.g., 'Eastern Standard Time')"
        if ([string]::IsNullOrWhiteSpace($timeZone)) {
            $timeZone = "Eastern Standard Time"
        }

        # Create autounattend.xml
        $autounattendPath = New-AutounattendXml -Username $username -Password $password -KeyboardLayout $keyboardLayout `
            -SystemLocale $systemLocale -UserLocale $userLocale -UILanguage $uiLanguage -TimeZone $timeZone
        Write-Host "`nCreated autounattend.xml with your settings." -ForegroundColor Green
    }

    # Apply the image
    Write-Host "`nInstalling selected edition to ${driveLetter}:\ ..." -ForegroundColor Cyan
    
    # Build DISM command as a single string
    $dismCmd = "/Apply-Image /ImageFile:`"$imageFile`" /Index:$edition /ApplyDir:${driveLetter}:\"
    
    if ($autounattendPath) {
        $dismCmd += " /Apply-Unattend:`"$autounattendPath`""
    }

    # Run DISM with call operator for better argument handling
    $dismProcess = Start-Process -FilePath "cmd.exe" -ArgumentList "/c dism.exe $dismCmd > ""$env:TEMP\dism_stdout.log"" 2> ""$env:TEMP\dism_stderr.log""" -NoNewWindow -PassThru -Wait
    
    # Display the DISM output
    if (Test-Path "$env:TEMP\dism_stdout.log") {
        $stdout = Get-Content -Path "$env:TEMP\dism_stdout.log" -Raw
        if (-not [string]::IsNullOrEmpty($stdout)) {
            Write-Host $stdout
        }
    }
    if (Test-Path "$env:TEMP\dism_stderr.log") {
        $stderr = Get-Content -Path "$env:TEMP\dism_stderr.log" -Raw
        if (-not [string]::IsNullOrEmpty($stderr)) {
            Write-Host $stderr -ForegroundColor Red
        }
    }
    
    if ($dismProcess.ExitCode -ne 0) {
        throw "Error applying Windows image. Exit code: $($dismProcess.ExitCode)"
    }

    # Copy unattend.xml to the target drive if it exists
    if ($autounattendPath -and (Test-Path $autounattendPath)) {
        $unattendTargetDir = "${driveLetter}:\Windows\System32\sysprep"
        if (-not (Test-Path $unattendTargetDir)) {
            New-Item -ItemType Directory -Path $unattendTargetDir -Force | Out-Null
        }
        Copy-Item -Path $autounattendPath -Destination "$unattendTargetDir\unattend.xml" -Force
    }

    # Create setup scripts directory if it doesn't exist
    $setupScriptsDir = "${driveLetter}:\Windows\Setup\Scripts"
    if (-not (Test-Path $setupScriptsDir)) {
        New-Item -ItemType Directory -Path $setupScriptsDir -Force | Out-Null
    }

    # Run bcdboot to make the drive bootable
    Write-Host "`nMaking the drive bootable..." -ForegroundColor Cyan
    $bcdbootPath = "${driveLetter}:\Windows\System32\bcdboot.exe"
    if (Test-Path $bcdbootPath) {
        & $bcdbootPath "${driveLetter}:\Windows" | Out-Null
    } else {
        Write-Host "Warning: Could not find bcdboot.exe on the target drive. The system may not be bootable." -ForegroundColor Yellow
    }

    Write-Host "`nInstallation complete!" -ForegroundColor Green
    Write-Host "You can now reboot your computer to start using your new Windows installation." -ForegroundColor Green

    # Cleanup
    Write-Host "`nCleaning up installation files..." -ForegroundColor Cyan
    if (Test-Path $extractPath) {
        Remove-Item -Path $extractPath -Recurse -Force -ErrorAction SilentlyContinue
    }
    Start-Process -FilePath "dism.exe" -ArgumentList "/Cleanup-Wim" -NoNewWindow -Wait -ErrorAction SilentlyContinue | Out-Null

    Write-Host "`nPress any key to exit..." -ForegroundColor Cyan
    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")

} catch {
    Write-Host "`nAn error occurred: $_" -ForegroundColor Red
    Write-Host $_.ScriptStackTrace -ForegroundColor Red
    Write-Host "`nPress any key to exit..." -ForegroundColor Cyan
    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
    exit 1
}

#endregion