#Requires -RunAsAdministrator

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

#region Modern UI Colors
$script:Colors = @{
    Background = [System.Drawing.Color]::FromArgb(30, 30, 30)
    Surface = [System.Drawing.Color]::FromArgb(45, 45, 45)
    SurfaceLight = [System.Drawing.Color]::FromArgb(60, 60, 60)
    Primary = [System.Drawing.Color]::FromArgb(0, 120, 215)
    PrimaryHover = [System.Drawing.Color]::FromArgb(0, 140, 235)
    Secondary = [System.Drawing.Color]::FromArgb(70, 70, 70)
    Text = [System.Drawing.Color]::FromArgb(255, 255, 255)
    TextSecondary = [System.Drawing.Color]::FromArgb(180, 180, 180)
    Success = [System.Drawing.Color]::FromArgb(46, 160, 67)
    Warning = [System.Drawing.Color]::FromArgb(210, 153, 34)
    Error = [System.Drawing.Color]::FromArgb(218, 54, 51)
    Border = [System.Drawing.Color]::FromArgb(80, 80, 80)
}
#endregion

#region Helper Functions

function Get-AvailableDrives {
    $drives = @()
    try {
        $drives = Get-Volume | 
            Where-Object { $_.DriveType -eq 'Fixed' -and $_.DriveLetter -ne $null -and $_.DriveLetter -ne 'C' } |
            Select-Object @{N="DriveLetter";E={$_.DriveLetter}},
                         @{N="Label";E={$_.FileSystemLabel}},
                         @{N="SizeGB";E={[math]::Round($_.Size/1GB,2)}},
                         @{N="FreeSpaceGB";E={[math]::Round($_.SizeRemaining/1GB,2)}} |
            Sort-Object DriveLetter
    } catch {
        $drives = Get-WmiObject -Class Win32_LogicalDisk | 
            Where-Object { $_.DriveType -eq 3 -and $_.DeviceID -ne "C:" } |
            Select-Object @{N="DriveLetter";E={$_.DeviceID -replace ':',''}},
                         @{N="Label";E={$_.VolumeName}},
                         @{N="SizeGB";E={[math]::Round($_.Size/1GB,2)}},
                         @{N="FreeSpaceGB";E={[math]::Round($_.FreeSpace/1GB,2)}} |
            Sort-Object DriveLetter
    }
    return $drives
}

function Test-DriveHasData {
    param([string]$DriveLetter)
    $path = "${DriveLetter}:\"
    if (Test-Path $path) {
        $items = Get-ChildItem -Path $path -Force -ErrorAction SilentlyContinue
        return ($items.Count -gt 0)
    }
    return $false
}

function Get-WindowsImageInfo {
    param([string]$ImagePath)
    $images = @()
    $output = & dism.exe /Get-WimInfo /WimFile:$ImagePath 2>&1
    $currentIndex = 0
    foreach ($line in $output) {
        if ($line -match "Index : (\d+)") { $currentIndex = $matches[1] }
        elseif ($line -match "Name : (.+)") {
            $images += [PSCustomObject]@{Index=[int]$currentIndex; Name=$matches[1].Trim()}
        }
    }
    return $images
}

function New-AutounattendXml {
    param($Username,$Password,$KeyboardLayout,$SystemLocale,$UserLocale,$UILanguage,$TimeZone)
    
    $passwordXml = if ($Password) { "<Password><Value>$Password</Value><PlainText>true</PlainText></Password>" } else { "" }
    
    $xml = @"
<?xml version="1.0" encoding="utf-8"?>
<unattend xmlns="urn:schemas-microsoft-com:unattend" xmlns:wcm="http://schemas.microsoft.com/WMIConfig/2002/State">
    <settings pass="oobeSystem">
        <component name="Microsoft-Windows-Shell-Setup" processorArchitecture="amd64" publicKeyToken="31bf3856ad364e35" language="neutral" versionScope="nonSxS">
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
        <component name="Microsoft-Windows-International-Core" processorArchitecture="amd64" publicKeyToken="31bf3856ad364e35" language="neutral" versionScope="nonSxS">
            <InputLocale>$KeyboardLayout</InputLocale>
            <SystemLocale>$SystemLocale</SystemLocale>
            <UILanguage>$UILanguage</UILanguage>
            <UILanguageFallback>en-US</UILanguageFallback>
            <UserLocale>$UserLocale</UserLocale>
        </component>
    </settings>
    <settings pass="specialize">
        <component name="Microsoft-Windows-Deployment" processorArchitecture="amd64" publicKeyToken="31bf3856ad364e35" language="neutral" versionScope="nonSxS">
            <RunSynchronous>
                <RunSynchronousCommand wcm:action="add">
                    <Order>1</Order>
                    <Path>cmd.exe /c "reg add HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\DataCollection /v AllowTelemetry /t REG_DWORD /d 0 /f"</Path>
                </RunSynchronousCommand>
            </RunSynchronous>
        </component>
    </settings>
</unattend>
"@
    
    $outputPath = Join-Path $PSScriptRoot "autounattend.xml"
    $xml | Out-File -FilePath $outputPath -Encoding utf8
    return $outputPath
}

function Format-DrivePartition {
    param([string]$DriveLetter)
    try {
        $volume = Get-Volume -DriveLetter $DriveLetter -ErrorAction Stop
        Format-Volume -DriveLetter $DriveLetter -FileSystem NTFS -Confirm:$false -Force -ErrorAction Stop
        return $true
    } catch {
        return $false
    }
}

#endregion

#region Keyboard Layouts Data

$script:KeyboardLayouts = @(
    @{Id="00000409";Name="US English"},
    @{Id="00000809";Name="UK English"},
    @{Id="00001009";Name="Canadian French"},
    @{Id="00000C0C";Name="Canadian Multilingual"},
    @{Id="0000040C";Name="French"},
    @{Id="00000407";Name="German"},
    @{Id="0000040A";Name="Spanish"},
    @{Id="00000410";Name="Italian"},
    @{Id="00000416";Name="Portuguese (Brazil)"},
    @{Id="00000816";Name="Portuguese"},
    @{Id="00000413";Name="Dutch"},
    @{Id="00000813";Name="Belgian (Period)"},
    @{Id="0000080C";Name="Belgian French"},
    @{Id="00000414";Name="Norwegian"},
    @{Id="0000041D";Name="Swedish"},
    @{Id="0000040B";Name="Finnish"},
    @{Id="00000406";Name="Danish"},
    @{Id="0000040F";Name="Icelandic"},
    @{Id="00010415";Name="Polish"},
    @{Id="00000405";Name="Czech"},
    @{Id="00010405";Name="Czech (QWERTY)"},
    @{Id="0000041B";Name="Slovak"},
    @{Id="0001041B";Name="Slovak (QWERTY)"},
    @{Id="0000040E";Name="Hungarian"},
    @{Id="00000418";Name="Romanian"},
    @{Id="00020402";Name="Bulgarian"},
    @{Id="00000408";Name="Greek"},
    @{Id="00010408";Name="Greek (220)"},
    @{Id="0000041F";Name="Turkish Q"},
    @{Id="00000419";Name="Russian"},
    @{Id="00000422";Name="Ukrainian"},
    @{Id="00000423";Name="Belarusian"},
    @{Id="00000411";Name="Japanese"},
    @{Id="00000412";Name="Korean"},
    @{Id="00000404";Name="Chinese (Traditional)"},
    @{Id="00000804";Name="Chinese (Simplified)"},
    @{Id="00000439";Name="Hindi Traditional"},
    @{Id="00000401";Name="Arabic (101)"},
    @{Id="00010401";Name="Arabic (102)"},
    @{Id="0000040D";Name="Hebrew"},
    @{Id="00000429";Name="Persian"},
    @{Id="0000041E";Name="Thai Kedmanee"},
    @{Id="0001041E";Name="Thai Pattachote"},
    @{Id="0000042A";Name="Vietnamese"},
    @{Id="00000421";Name="Indonesian"},
    @{Id="0000043A";Name="Maltese 47-Key"},
    @{Id="0000041A";Name="Croatian"},
    @{Id="00000424";Name="Slovenian"},
    @{Id="0000081A";Name="Serbian (Latin)"},
    @{Id="00000C1A";Name="Serbian (Cyrillic)"},
    @{Id="00000427";Name="Lithuanian"},
    @{Id="00010427";Name="Lithuanian IBM"},
    @{Id="00000426";Name="Latvian"},
    @{Id="00000425";Name="Estonian"}
)

#endregion

#region Build GUI

# Helper function to style GroupBox
function Style-GroupBox {
    param($GroupBox)
    $GroupBox.ForeColor = $script:Colors.Text
    $GroupBox.BackColor = $script:Colors.Surface
    $GroupBox.FlatStyle = "Flat"
}

# Helper function to style Button
function Style-Button {
    param($Button, [switch]$Primary, [switch]$Secondary)
    $Button.FlatStyle = "Flat"
    $Button.FlatAppearance.BorderSize = 0
    $Button.Cursor = [System.Windows.Forms.Cursors]::Hand
    if ($Primary) {
        $Button.BackColor = $script:Colors.Primary
        $Button.ForeColor = $script:Colors.Text
    } elseif ($Secondary) {
        $Button.BackColor = $script:Colors.Secondary
        $Button.ForeColor = $script:Colors.Text
    } else {
        $Button.BackColor = $script:Colors.SurfaceLight
        $Button.ForeColor = $script:Colors.Text
    }
}

# Helper function to style TextBox
function Style-TextBox {
    param($TextBox)
    $TextBox.BackColor = $script:Colors.SurfaceLight
    $TextBox.ForeColor = $script:Colors.Text
    $TextBox.BorderStyle = "FixedSingle"
}

# Helper function to style ComboBox
function Style-ComboBox {
    param($ComboBox)
    $ComboBox.BackColor = $script:Colors.SurfaceLight
    $ComboBox.ForeColor = $script:Colors.Text
    $ComboBox.FlatStyle = "Flat"
}

# Helper function to style ListBox
function Style-ListBox {
    param($ListBox)
    $ListBox.BackColor = $script:Colors.SurfaceLight
    $ListBox.ForeColor = $script:Colors.Text
    $ListBox.BorderStyle = "FixedSingle"
}

$form = New-Object System.Windows.Forms.Form
$form.Text = "Windows Installation Tool"
$form.Size = New-Object System.Drawing.Size(700, 950)
$form.StartPosition = "CenterScreen"
$form.FormBorderStyle = "FixedSingle"
$form.MaximizeBox = $false
$form.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$form.BackColor = $script:Colors.Background

# Title Panel
$titlePanel = New-Object System.Windows.Forms.Panel
$titlePanel.Location = New-Object System.Drawing.Point(0, 0)
$titlePanel.Size = New-Object System.Drawing.Size(700, 70)
$titlePanel.BackColor = $script:Colors.Primary
$form.Controls.Add($titlePanel)

# Title Label
$titleLabel = New-Object System.Windows.Forms.Label
$titleLabel.Text = "Windows Installation Tool"
$titleLabel.Font = New-Object System.Drawing.Font("Segoe UI Light", 22)
$titleLabel.ForeColor = $script:Colors.Text
$titleLabel.Location = New-Object System.Drawing.Point(25, 18)
$titleLabel.Size = New-Object System.Drawing.Size(400, 40)
$titleLabel.BackColor = [System.Drawing.Color]::Transparent
$titlePanel.Controls.Add($titleLabel)

# Subtitle
$subtitleLabel = New-Object System.Windows.Forms.Label
$subtitleLabel.Text = "Install Windows without USB"
$subtitleLabel.Font = New-Object System.Drawing.Font("Segoe UI", 10)
$subtitleLabel.ForeColor = $script:Colors.TextSecondary
$subtitleLabel.Location = New-Object System.Drawing.Point(430, 28)
$subtitleLabel.Size = New-Object System.Drawing.Size(250, 20)
$subtitleLabel.BackColor = [System.Drawing.Color]::Transparent
$subtitleLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleRight
$titlePanel.Controls.Add($subtitleLabel)

$yPos = 85

# === ISO Selection ===
$isoGroupBox = New-Object System.Windows.Forms.GroupBox
$isoGroupBox.Text = "  Windows ISO  "
$isoGroupBox.Location = New-Object System.Drawing.Point(20, $yPos)
$isoGroupBox.Size = New-Object System.Drawing.Size(645, 75)
Style-GroupBox $isoGroupBox
$form.Controls.Add($isoGroupBox)

$isoPathTextBox = New-Object System.Windows.Forms.TextBox
$isoPathTextBox.Location = New-Object System.Drawing.Point(15, 30)
$isoPathTextBox.Size = New-Object System.Drawing.Size(520, 28)
$isoPathTextBox.ReadOnly = $true
Style-TextBox $isoPathTextBox
$isoGroupBox.Controls.Add($isoPathTextBox)

$browseButton = New-Object System.Windows.Forms.Button
$browseButton.Text = "Browse..."
$browseButton.Location = New-Object System.Drawing.Point(545, 28)
$browseButton.Size = New-Object System.Drawing.Size(85, 30)
Style-Button $browseButton -Primary
$isoGroupBox.Controls.Add($browseButton)

$yPos += 85

# === Windows Edition & Target Drive (side by side) ===
$editionGroupBox = New-Object System.Windows.Forms.GroupBox
$editionGroupBox.Text = "  Windows Edition  "
$editionGroupBox.Location = New-Object System.Drawing.Point(20, $yPos)
$editionGroupBox.Size = New-Object System.Drawing.Size(310, 70)
Style-GroupBox $editionGroupBox
$form.Controls.Add($editionGroupBox)

$editionComboBox = New-Object System.Windows.Forms.ComboBox
$editionComboBox.Location = New-Object System.Drawing.Point(15, 28)
$editionComboBox.Size = New-Object System.Drawing.Size(280, 28)
$editionComboBox.DropDownStyle = "DropDownList"
$editionComboBox.Enabled = $false
Style-ComboBox $editionComboBox
$editionGroupBox.Controls.Add($editionComboBox)

$driveGroupBox = New-Object System.Windows.Forms.GroupBox
$driveGroupBox.Text = "  Target Drive (C: excluded)  "
$driveGroupBox.Location = New-Object System.Drawing.Point(345, $yPos)
$driveGroupBox.Size = New-Object System.Drawing.Size(320, 70)
Style-GroupBox $driveGroupBox
$form.Controls.Add($driveGroupBox)

$driveComboBox = New-Object System.Windows.Forms.ComboBox
$driveComboBox.Location = New-Object System.Drawing.Point(15, 28)
$driveComboBox.Size = New-Object System.Drawing.Size(200, 28)
$driveComboBox.DropDownStyle = "DropDownList"
Style-ComboBox $driveComboBox
$driveGroupBox.Controls.Add($driveComboBox)

$refreshDrivesButton = New-Object System.Windows.Forms.Button
$refreshDrivesButton.Text = "Refresh"
$refreshDrivesButton.Location = New-Object System.Drawing.Point(225, 26)
$refreshDrivesButton.Size = New-Object System.Drawing.Size(80, 30)
Style-Button $refreshDrivesButton
$driveGroupBox.Controls.Add($refreshDrivesButton)

$yPos += 80

# === Local Account ===
$accountGroupBox = New-Object System.Windows.Forms.GroupBox
$accountGroupBox.Text = "  Local Account Setup  "
$accountGroupBox.Location = New-Object System.Drawing.Point(20, $yPos)
$accountGroupBox.Size = New-Object System.Drawing.Size(645, 75)
Style-GroupBox $accountGroupBox
$form.Controls.Add($accountGroupBox)

$usernameLabel = New-Object System.Windows.Forms.Label
$usernameLabel.Text = "Username:"
$usernameLabel.Location = New-Object System.Drawing.Point(15, 32)
$usernameLabel.Size = New-Object System.Drawing.Size(75, 20)
$usernameLabel.ForeColor = $script:Colors.TextSecondary
$accountGroupBox.Controls.Add($usernameLabel)

$usernameTextBox = New-Object System.Windows.Forms.TextBox
$usernameTextBox.Location = New-Object System.Drawing.Point(95, 30)
$usernameTextBox.Size = New-Object System.Drawing.Size(180, 28)
Style-TextBox $usernameTextBox
$accountGroupBox.Controls.Add($usernameTextBox)

$passwordLabel = New-Object System.Windows.Forms.Label
$passwordLabel.Text = "Password:"
$passwordLabel.Location = New-Object System.Drawing.Point(295, 32)
$passwordLabel.Size = New-Object System.Drawing.Size(70, 20)
$passwordLabel.ForeColor = $script:Colors.TextSecondary
$accountGroupBox.Controls.Add($passwordLabel)

$passwordTextBox = New-Object System.Windows.Forms.TextBox
$passwordTextBox.Location = New-Object System.Drawing.Point(370, 30)
$passwordTextBox.Size = New-Object System.Drawing.Size(180, 28)
$passwordTextBox.UseSystemPasswordChar = $true
Style-TextBox $passwordTextBox
$accountGroupBox.Controls.Add($passwordTextBox)

$showPasswordCheckBox = New-Object System.Windows.Forms.CheckBox
$showPasswordCheckBox.Text = "Show"
$showPasswordCheckBox.Location = New-Object System.Drawing.Point(560, 32)
$showPasswordCheckBox.Size = New-Object System.Drawing.Size(60, 20)
$showPasswordCheckBox.ForeColor = $script:Colors.TextSecondary
$showPasswordCheckBox.FlatStyle = "Flat"
$accountGroupBox.Controls.Add($showPasswordCheckBox)

$yPos += 85

# === Regional Settings ===
$regionalGroupBox = New-Object System.Windows.Forms.GroupBox
$regionalGroupBox.Text = "  Regional Settings  "
$regionalGroupBox.Location = New-Object System.Drawing.Point(20, $yPos)
$regionalGroupBox.Size = New-Object System.Drawing.Size(645, 75)
Style-GroupBox $regionalGroupBox
$form.Controls.Add($regionalGroupBox)

$keyboardLabel = New-Object System.Windows.Forms.Label
$keyboardLabel.Text = "Keyboard:"
$keyboardLabel.Location = New-Object System.Drawing.Point(15, 32)
$keyboardLabel.Size = New-Object System.Drawing.Size(65, 20)
$keyboardLabel.ForeColor = $script:Colors.TextSecondary
$regionalGroupBox.Controls.Add($keyboardLabel)

$keyboardComboBox = New-Object System.Windows.Forms.ComboBox
$keyboardComboBox.Location = New-Object System.Drawing.Point(85, 28)
$keyboardComboBox.Size = New-Object System.Drawing.Size(200, 28)
$keyboardComboBox.DropDownStyle = "DropDownList"
Style-ComboBox $keyboardComboBox
$regionalGroupBox.Controls.Add($keyboardComboBox)

$timezoneLabel = New-Object System.Windows.Forms.Label
$timezoneLabel.Text = "Timezone:"
$timezoneLabel.Location = New-Object System.Drawing.Point(300, 32)
$timezoneLabel.Size = New-Object System.Drawing.Size(70, 20)
$timezoneLabel.ForeColor = $script:Colors.TextSecondary
$regionalGroupBox.Controls.Add($timezoneLabel)

$timezoneComboBox = New-Object System.Windows.Forms.ComboBox
$timezoneComboBox.Location = New-Object System.Drawing.Point(375, 28)
$timezoneComboBox.Size = New-Object System.Drawing.Size(255, 28)
$timezoneComboBox.DropDownStyle = "DropDownList"
$timezoneComboBox.AutoCompleteSource = "ListItems"
$timezoneComboBox.AutoCompleteMode = "SuggestAppend"
Style-ComboBox $timezoneComboBox
$regionalGroupBox.Controls.Add($timezoneComboBox)

$yPos += 85

# === Files to Copy to Desktop ===
$filesGroupBox = New-Object System.Windows.Forms.GroupBox
$filesGroupBox.Text = "  Files to Copy to Desktop  "
$filesGroupBox.Location = New-Object System.Drawing.Point(20, $yPos)
$filesGroupBox.Size = New-Object System.Drawing.Size(310, 140)
Style-GroupBox $filesGroupBox
$form.Controls.Add($filesGroupBox)

$filesListBox = New-Object System.Windows.Forms.ListBox
$filesListBox.Location = New-Object System.Drawing.Point(15, 25)
$filesListBox.Size = New-Object System.Drawing.Size(195, 100)
$filesListBox.SelectionMode = "MultiExtended"
$filesListBox.HorizontalScrollbar = $true
Style-ListBox $filesListBox
$filesGroupBox.Controls.Add($filesListBox)

$addFilesButton = New-Object System.Windows.Forms.Button
$addFilesButton.Text = "Add Files"
$addFilesButton.Location = New-Object System.Drawing.Point(220, 25)
$addFilesButton.Size = New-Object System.Drawing.Size(75, 30)
Style-Button $addFilesButton
$filesGroupBox.Controls.Add($addFilesButton)

$removeFilesButton = New-Object System.Windows.Forms.Button
$removeFilesButton.Text = "Remove"
$removeFilesButton.Location = New-Object System.Drawing.Point(220, 60)
$removeFilesButton.Size = New-Object System.Drawing.Size(75, 30)
Style-Button $removeFilesButton
$filesGroupBox.Controls.Add($removeFilesButton)

$clearFilesButton = New-Object System.Windows.Forms.Button
$clearFilesButton.Text = "Clear All"
$clearFilesButton.Location = New-Object System.Drawing.Point(220, 95)
$clearFilesButton.Size = New-Object System.Drawing.Size(75, 30)
Style-Button $clearFilesButton
$filesGroupBox.Controls.Add($clearFilesButton)

# === Registry Files to Embed ===
$regGroupBox = New-Object System.Windows.Forms.GroupBox
$regGroupBox.Text = "  Registry Files to Embed  "
$regGroupBox.Location = New-Object System.Drawing.Point(345, $yPos)
$regGroupBox.Size = New-Object System.Drawing.Size(320, 140)
Style-GroupBox $regGroupBox
$form.Controls.Add($regGroupBox)

$regListBox = New-Object System.Windows.Forms.ListBox
$regListBox.Location = New-Object System.Drawing.Point(15, 25)
$regListBox.Size = New-Object System.Drawing.Size(205, 100)
$regListBox.SelectionMode = "MultiExtended"
$regListBox.HorizontalScrollbar = $true
Style-ListBox $regListBox
$regGroupBox.Controls.Add($regListBox)

$addRegButton = New-Object System.Windows.Forms.Button
$addRegButton.Text = "Add .reg"
$addRegButton.Location = New-Object System.Drawing.Point(230, 25)
$addRegButton.Size = New-Object System.Drawing.Size(75, 30)
Style-Button $addRegButton
$regGroupBox.Controls.Add($addRegButton)

$removeRegButton = New-Object System.Windows.Forms.Button
$removeRegButton.Text = "Remove"
$removeRegButton.Location = New-Object System.Drawing.Point(230, 60)
$removeRegButton.Size = New-Object System.Drawing.Size(75, 30)
Style-Button $removeRegButton
$regGroupBox.Controls.Add($removeRegButton)

$clearRegButton = New-Object System.Windows.Forms.Button
$clearRegButton.Text = "Clear All"
$clearRegButton.Location = New-Object System.Drawing.Point(230, 95)
$clearRegButton.Size = New-Object System.Drawing.Size(75, 30)
Style-Button $clearRegButton
$regGroupBox.Controls.Add($clearRegButton)

$yPos += 150

# === Progress ===
$progressGroupBox = New-Object System.Windows.Forms.GroupBox
$progressGroupBox.Text = "  Progress  "
$progressGroupBox.Location = New-Object System.Drawing.Point(20, $yPos)
$progressGroupBox.Size = New-Object System.Drawing.Size(645, 85)
Style-GroupBox $progressGroupBox
$form.Controls.Add($progressGroupBox)

$statusLabel = New-Object System.Windows.Forms.Label
$statusLabel.Text = "Ready - Select a Windows ISO to begin"
$statusLabel.Location = New-Object System.Drawing.Point(15, 28)
$statusLabel.Size = New-Object System.Drawing.Size(615, 20)
$statusLabel.ForeColor = $script:Colors.TextSecondary
$progressGroupBox.Controls.Add($statusLabel)

$progressBar = New-Object System.Windows.Forms.ProgressBar
$progressBar.Location = New-Object System.Drawing.Point(15, 52)
$progressBar.Size = New-Object System.Drawing.Size(615, 22)
$progressBar.Style = "Continuous"
$progressGroupBox.Controls.Add($progressBar)

$yPos += 100

# === Buttons ===
$installButton = New-Object System.Windows.Forms.Button
$installButton.Text = "? Install Windows"
$installButton.Location = New-Object System.Drawing.Point(20, $yPos)
$installButton.Size = New-Object System.Drawing.Size(310, 50)
$installButton.Font = New-Object System.Drawing.Font("Segoe UI", 12, [System.Drawing.FontStyle]::Bold)
$installButton.BackColor = $script:Colors.Primary
$installButton.ForeColor = $script:Colors.Text
$installButton.FlatStyle = "Flat"
$installButton.FlatAppearance.BorderSize = 0
$installButton.Cursor = [System.Windows.Forms.Cursors]::Hand
$form.Controls.Add($installButton)

$cancelButton = New-Object System.Windows.Forms.Button
$cancelButton.Text = "Exit"
$cancelButton.Location = New-Object System.Drawing.Point(355, $yPos)
$cancelButton.Size = New-Object System.Drawing.Size(310, 50)
$cancelButton.Font = New-Object System.Drawing.Font("Segoe UI", 12)
$cancelButton.BackColor = $script:Colors.Secondary
$cancelButton.ForeColor = $script:Colors.Text
$cancelButton.FlatStyle = "Flat"
$cancelButton.FlatAppearance.BorderSize = 0
$cancelButton.Cursor = [System.Windows.Forms.Cursors]::Hand
$form.Controls.Add($cancelButton)

# Store file lists in script scope
$script:FilesToCopy = @()
$script:RegFilesToEmbed = @()

#endregion

#region Event Handlers

# Populate keyboard layouts
foreach ($kb in $script:KeyboardLayouts) {
    $keyboardComboBox.Items.Add($kb.Name) | Out-Null
}
$keyboardComboBox.SelectedIndex = 0

# Populate timezones
$script:Timezones = Get-TimeZone -ListAvailable | Sort-Object BaseUtcOffset, DisplayName
foreach ($tz in $script:Timezones) {
    $timezoneComboBox.Items.Add($tz.DisplayName) | Out-Null
}
# Select current timezone
$currentTz = Get-TimeZone
$currentTzIndex = 0
for ($i = 0; $i -lt $script:Timezones.Count; $i++) {
    if ($script:Timezones[$i].Id -eq $currentTz.Id) {
        $currentTzIndex = $i
        break
    }
}
$timezoneComboBox.SelectedIndex = $currentTzIndex

# Populate drives function
function Update-DriveList {
    $driveComboBox.Items.Clear()
    $script:AvailableDrives = Get-AvailableDrives
    foreach ($drive in $script:AvailableDrives) {
        $label = if ($drive.Label) { " - $($drive.Label)" } else { "" }
        $driveComboBox.Items.Add("$($drive.DriveLetter):$label ($($drive.SizeGB) GB, $($drive.FreeSpaceGB) GB free)") | Out-Null
    }
    if ($driveComboBox.Items.Count -gt 0) {
        $driveComboBox.SelectedIndex = 0
    }
}
Update-DriveList

# Show/Hide password
$showPasswordCheckBox.Add_CheckedChanged({
    $passwordTextBox.UseSystemPasswordChar = -not $showPasswordCheckBox.Checked
})

# Refresh drives button
$refreshDrivesButton.Add_Click({
    Update-DriveList
    [System.Windows.Forms.MessageBox]::Show("Drive list refreshed.", "Info", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
})

# === Files to Copy to Desktop Event Handlers ===
$addFilesButton.Add_Click({
    $openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $openFileDialog.Filter = "All Files (*.*)|*.*"
    $openFileDialog.Title = "Select files to copy to desktop"
    $openFileDialog.Multiselect = $true
    
    if ($openFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        foreach ($file in $openFileDialog.FileNames) {
            if ($script:FilesToCopy -notcontains $file) {
                $script:FilesToCopy += $file
                $filesListBox.Items.Add([System.IO.Path]::GetFileName($file)) | Out-Null
            }
        }
        $statusLabel.Text = "Added $($openFileDialog.FileNames.Count) file(s) to copy to desktop"
    }
})

$removeFilesButton.Add_Click({
    $selectedIndices = @($filesListBox.SelectedIndices)
    for ($i = $selectedIndices.Count - 1; $i -ge 0; $i--) {
        $index = $selectedIndices[$i]
        $script:FilesToCopy = @($script:FilesToCopy | Where-Object { $_ -ne $script:FilesToCopy[$index] })
        $filesListBox.Items.RemoveAt($index)
    }
})

$clearFilesButton.Add_Click({
    $script:FilesToCopy = @()
    $filesListBox.Items.Clear()
    $statusLabel.Text = "Cleared files to copy list"
})

# === Registry Files Event Handlers ===
$addRegButton.Add_Click({
    $openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $openFileDialog.Filter = "Registry Files (*.reg)|*.reg|All Files (*.*)|*.*"
    $openFileDialog.Title = "Select registry files to embed"
    $openFileDialog.Multiselect = $true
    
    if ($openFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        foreach ($file in $openFileDialog.FileNames) {
            if ($script:RegFilesToEmbed -notcontains $file) {
                $script:RegFilesToEmbed += $file
                $regListBox.Items.Add([System.IO.Path]::GetFileName($file)) | Out-Null
            }
        }
        $statusLabel.Text = "Added $($openFileDialog.FileNames.Count) registry file(s) to embed"
    }
})

$removeRegButton.Add_Click({
    $selectedIndices = @($regListBox.SelectedIndices)
    for ($i = $selectedIndices.Count - 1; $i -ge 0; $i--) {
        $index = $selectedIndices[$i]
        $script:RegFilesToEmbed = @($script:RegFilesToEmbed | Where-Object { $_ -ne $script:RegFilesToEmbed[$index] })
        $regListBox.Items.RemoveAt($index)
    }
})

$clearRegButton.Add_Click({
    $script:RegFilesToEmbed = @()
    $regListBox.Items.Clear()
    $statusLabel.Text = "Cleared registry files list"
})

# Browse ISO button
$browseButton.Add_Click({
    $openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $openFileDialog.Filter = "ISO Files (*.iso)|*.iso|All Files (*.*)|*.*"
    $openFileDialog.Title = "Select Windows ISO"
    
    if ($openFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $isoPathTextBox.Text = $openFileDialog.FileName
        $statusLabel.Text = "Analyzing ISO..."
        $form.Refresh()
        
        # Check for 7-Zip
        $sevenZipPath = $null
        $local7z = Join-Path $PSScriptRoot "7z.exe"
        if (Test-Path $local7z) {
            $sevenZipPath = $local7z
        } else {
            $cmd = Get-Command "7z.exe" -ErrorAction SilentlyContinue
            if ($cmd) { $sevenZipPath = $cmd.Source }
        }
        
        if (-not $sevenZipPath) {
            [System.Windows.Forms.MessageBox]::Show("7-Zip not found! Place 7z.exe next to this script or install 7-Zip.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            $statusLabel.Text = "Error: 7-Zip not found"
            return
        }
        
        $script:SevenZipPath = $sevenZipPath
        
        # Extract to temp to get image info
        $tempExtract = Join-Path $env:TEMP "WinInstallTemp"
        Remove-Item $tempExtract -Recurse -Force -ErrorAction SilentlyContinue
        New-Item -ItemType Directory -Path $tempExtract -Force | Out-Null
        
        # Extract only sources folder to check for install.wim/esd
        & $sevenZipPath x -y "-o$tempExtract" $openFileDialog.FileName "sources\install.*" 2>&1 | Out-Null
        
        $wimFile = Join-Path $tempExtract "sources\install.wim"
        $esdFile = Join-Path $tempExtract "sources\install.esd"
        $imageFile = if (Test-Path $wimFile) { $wimFile } elseif (Test-Path $esdFile) { $esdFile } else { $null }
        
        if (-not $imageFile) {
            [System.Windows.Forms.MessageBox]::Show("No Windows image found in ISO!", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            $statusLabel.Text = "Error: Invalid ISO"
            Remove-Item $tempExtract -Recurse -Force -ErrorAction SilentlyContinue
            return
        }
        
        # Get editions
        $editionComboBox.Items.Clear()
        $script:WindowsImages = Get-WindowsImageInfo -ImagePath $imageFile
        foreach ($img in $script:WindowsImages) {
            $editionComboBox.Items.Add("$($img.Index). $($img.Name)") | Out-Null
        }
        
        if ($editionComboBox.Items.Count -gt 0) {
            $editionComboBox.SelectedIndex = 0
            $editionComboBox.Enabled = $true
        }
        
        Remove-Item $tempExtract -Recurse -Force -ErrorAction SilentlyContinue
        $statusLabel.Text = "ISO loaded successfully. Found $($script:WindowsImages.Count) edition(s)."
    }
})

# Cancel button
$cancelButton.Add_Click({
    $form.Close()
})

# Install button
$installButton.Add_Click({
    # Validation
    if ([string]::IsNullOrWhiteSpace($isoPathTextBox.Text)) {
        [System.Windows.Forms.MessageBox]::Show("Please select a Windows ISO file.", "Validation Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        return
    }
    
    if ($editionComboBox.SelectedIndex -lt 0) {
        [System.Windows.Forms.MessageBox]::Show("Please select a Windows edition.", "Validation Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        return
    }
    
    if ($driveComboBox.SelectedIndex -lt 0) {
        [System.Windows.Forms.MessageBox]::Show("Please select a target drive.", "Validation Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        return
    }
    
    if ([string]::IsNullOrWhiteSpace($usernameTextBox.Text)) {
        [System.Windows.Forms.MessageBox]::Show("Please enter a username.", "Validation Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        return
    }
    
    # Get selected values
    $selectedDrive = $script:AvailableDrives[$driveComboBox.SelectedIndex]
    $driveLetter = $selectedDrive.DriveLetter
    $selectedEdition = $script:WindowsImages[$editionComboBox.SelectedIndex]
    $selectedKeyboard = $script:KeyboardLayouts[$keyboardComboBox.SelectedIndex]
    $selectedTimezone = $script:Timezones[$timezoneComboBox.SelectedIndex]
    
    # Check if drive has data
    if (Test-DriveHasData -DriveLetter $driveLetter) {
        $formatResult = [System.Windows.Forms.MessageBox]::Show(
            "Drive ${driveLetter}: contains data!`n`nWARNING: All data on this drive will be permanently deleted!`n`nDo you want to format the drive and continue?",
            "Format Required",
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        )
        
        if ($formatResult -ne [System.Windows.Forms.DialogResult]::Yes) {
            return
        }
        
        # Double confirmation
        $confirmResult = [System.Windows.Forms.MessageBox]::Show(
            "Are you ABSOLUTELY SURE you want to format drive ${driveLetter}:?`n`nThis action CANNOT be undone!",
            "Final Confirmation",
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Exclamation
        )
        
        if ($confirmResult -ne [System.Windows.Forms.DialogResult]::Yes) {
            return
        }
        
        $statusLabel.Text = "Formatting drive ${driveLetter}:..."
        $progressBar.Value = 5
        $form.Refresh()
        
        if (-not (Format-DrivePartition -DriveLetter $driveLetter)) {
            [System.Windows.Forms.MessageBox]::Show("Failed to format drive ${driveLetter}:. Please format it manually and try again.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            $statusLabel.Text = "Format failed"
            $progressBar.Value = 0
            return
        }
    }
    
    # Disable controls during installation
    $installButton.Enabled = $false
    $browseButton.Enabled = $false
    $refreshDrivesButton.Enabled = $false
    
    try {
        # Cleanup
        $statusLabel.Text = "Cleaning up previous files..."
        $progressBar.Value = 10
        $form.Refresh()
        
        $extractPath = "C:\WindowsInstallation"
        Remove-Item $extractPath -Recurse -Force -ErrorAction SilentlyContinue
        Start-Process dism.exe -ArgumentList "/Cleanup-Wim" -NoNewWindow -Wait -ErrorAction SilentlyContinue
        
        # Extract ISO
        $statusLabel.Text = "Extracting ISO (this may take several minutes)..."
        $progressBar.Value = 20
        $form.Refresh()
        
        New-Item -ItemType Directory -Path $extractPath -Force | Out-Null
        $extractProcess = Start-Process -FilePath $script:SevenZipPath -ArgumentList "x -y `"-o$extractPath`" `"$($isoPathTextBox.Text)`"" -NoNewWindow -Wait -PassThru
        
        if ($extractProcess.ExitCode -ne 0) {
            throw "Failed to extract ISO"
        }
        
        # Find image file
        $wimFile = Join-Path $extractPath "sources\install.wim"
        $esdFile = Join-Path $extractPath "sources\install.esd"
        $imageFile = if (Test-Path $wimFile) { $wimFile } elseif (Test-Path $esdFile) { $esdFile } else { $null }
        
        if (-not $imageFile) {
            throw "No Windows image found in extracted ISO"
        }
        
        $progressBar.Value = 40
        $form.Refresh()
        
        # Create autounattend.xml
        $statusLabel.Text = "Creating unattend configuration..."
        $progressBar.Value = 45
        $form.Refresh()
        
        $autounattendPath = New-AutounattendXml -Username $usernameTextBox.Text -Password $passwordTextBox.Text `
            -KeyboardLayout $selectedKeyboard.Id -SystemLocale "en-US" -UserLocale "en-US" `
            -UILanguage "en-US" -TimeZone $selectedTimezone.Id
        
        # Apply image
        $statusLabel.Text = "Installing Windows to ${driveLetter}:\ (this will take a while)..."
        $progressBar.Value = 50
        $form.Refresh()
        
        $dismArgs = "/Apply-Image /ImageFile:`"$imageFile`" /Index:$($selectedEdition.Index) /ApplyDir:${driveLetter}:\"
        $dismProcess = Start-Process dism.exe -ArgumentList $dismArgs -NoNewWindow -Wait -PassThru
        
        if ($dismProcess.ExitCode -ne 0) {
            throw "DISM failed with exit code: $($dismProcess.ExitCode)"
        }
        
        $progressBar.Value = 80
        $form.Refresh()
        
        # Copy unattend.xml
        $statusLabel.Text = "Copying configuration files..."
        $progressBar.Value = 85
        $form.Refresh()
        
        $unattendDir = "${driveLetter}:\Windows\System32\sysprep"
        New-Item -ItemType Directory -Path $unattendDir -Force -ErrorAction SilentlyContinue | Out-Null
        Copy-Item $autounattendPath "$unattendDir\unattend.xml" -Force
        
        # Copy files to desktop
        if ($script:FilesToCopy.Count -gt 0) {
            $statusLabel.Text = "Copying files to desktop..."
            $progressBar.Value = 87
            $form.Refresh()
            
            $desktopPath = "${driveLetter}:\Users\Default\Desktop"
            New-Item -ItemType Directory -Path $desktopPath -Force -ErrorAction SilentlyContinue | Out-Null
            
            foreach ($file in $script:FilesToCopy) {
                if (Test-Path $file) {
                    Copy-Item $file $desktopPath -Force -ErrorAction SilentlyContinue
                }
            }
        }
        
        # Embed registry files
        if ($script:RegFilesToEmbed.Count -gt 0) {
            $statusLabel.Text = "Preparing registry files..."
            $progressBar.Value = 88
            $form.Refresh()
            
            # Create a script to apply reg files on first login
            $regScriptPath = "${driveLetter}:\Windows\Setup\Scripts"
            New-Item -ItemType Directory -Path $regScriptPath -Force -ErrorAction SilentlyContinue | Out-Null
            
            # Copy reg files to a temp location
            $regFilesDir = "${driveLetter}:\Windows\Setup\RegFiles"
            New-Item -ItemType Directory -Path $regFilesDir -Force -ErrorAction SilentlyContinue | Out-Null
            
            foreach ($regFile in $script:RegFilesToEmbed) {
                if (Test-Path $regFile) {
                    Copy-Item $regFile $regFilesDir -Force -ErrorAction SilentlyContinue
                }
            }
            
            # Create SetupComplete.cmd to apply reg files
            $setupCompleteContent = "@echo off`r`n"
            foreach ($regFile in $script:RegFilesToEmbed) {
                $regFileName = [System.IO.Path]::GetFileName($regFile)
                $setupCompleteContent += "reg import `"C:\Windows\Setup\RegFiles\$regFileName`"`r`n"
            }
            $setupCompleteContent += "rd /s /q `"C:\Windows\Setup\RegFiles`"`r`n"
            
            $setupCompletePath = Join-Path $regScriptPath "SetupComplete.cmd"
            $setupCompleteContent | Out-File -FilePath $setupCompletePath -Encoding ascii -Force
        }
        
        # Make bootable
        $statusLabel.Text = "Making drive bootable..."
        $progressBar.Value = 90
        $form.Refresh()
        
        $bcdbootPath = "${driveLetter}:\Windows\System32\bcdboot.exe"
        if (Test-Path $bcdbootPath) {
            & $bcdbootPath "${driveLetter}:\Windows"
        }
        
        # Cleanup
        $statusLabel.Text = "Cleaning up..."
        $progressBar.Value = 95
        $form.Refresh()
        
        Remove-Item $extractPath -Recurse -Force -ErrorAction SilentlyContinue
        Start-Process dism.exe -ArgumentList "/Cleanup-Wim" -NoNewWindow -Wait -ErrorAction SilentlyContinue
        
        $progressBar.Value = 100
        $statusLabel.Text = "Installation complete!"
        
        [System.Windows.Forms.MessageBox]::Show(
            "Windows has been installed successfully to drive ${driveLetter}:!`n`nReboot your computer and select the new installation from the boot menu.",
            "Installation Complete",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information
        )
        
    } catch {
        $statusLabel.Text = "Error: $_"
        $progressBar.Value = 0
        [System.Windows.Forms.MessageBox]::Show("Installation failed: $_", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    } finally {
        $installButton.Enabled = $true
        $browseButton.Enabled = $true
        $refreshDrivesButton.Enabled = $true
    }
})

#endregion

# Show form
[void]$form.ShowDialog()
