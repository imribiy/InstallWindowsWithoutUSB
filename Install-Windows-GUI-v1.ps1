<# : batch script
@echo off
setlocal EnableDelayedExpansion
net session >nul 2>&1
if %errorlevel% neq 0 (powershell -Command "Start-Process -FilePath \"%~f0\" -Verb RunAs" & exit /b)
set "TEMPPS1=%TEMP%\WinInstallGUI_%RANDOM%.ps1"
powershell -NoProfile -Command "$c = Get-Content -LiteralPath '%~f0' -Raw; $c = $c -replace '(?s)^.*?#\>', ''; Set-Content -LiteralPath '%TEMPPS1%' -Value $c -Encoding UTF8"
set "SCRIPT_ORIGIN=%~dp0"
powershell -ExecutionPolicy Bypass -NoProfile -File "%TEMPPS1%"
del "%TEMPPS1%" 2>nul
exit /b
#>

Add-Type -AssemblyName System.Windows.Forms,System.Drawing
if(-not ("WinInstall.Native" -as [type])){
Add-Type -Namespace WinInstall -Name Native -MemberDefinition @'
[DllImport("kernel32.dll", CharSet=CharSet.Unicode, SetLastError=true)]
public static extern IntPtr LoadLibraryExW(string lpFileName, IntPtr hFile, uint dwFlags);
[DllImport("user32.dll", CharSet=CharSet.Unicode)]
public static extern int LoadStringW(IntPtr hInstance, uint uID, System.Text.StringBuilder lpBuffer, int nBufferMax);
[DllImport("kernel32.dll")]
public static extern bool FreeLibrary(IntPtr hModule);
'@
}

#region Config & Variables
$script:Colors = @{
    Background=[System.Drawing.Color]::FromArgb(30,30,30); Surface=[System.Drawing.Color]::FromArgb(45,45,45)
    SurfaceLight=[System.Drawing.Color]::FromArgb(60,60,60); Primary=[System.Drawing.Color]::FromArgb(0,120,215)
    Secondary=[System.Drawing.Color]::FromArgb(70,70,70); Text=[System.Drawing.Color]::FromArgb(255,255,255)
    TextSecondary=[System.Drawing.Color]::FromArgb(180,180,180); Success=[System.Drawing.Color]::FromArgb(46,160,67)
    Warning=[System.Drawing.Color]::FromArgb(210,153,34); Error=[System.Drawing.Color]::FromArgb(218,54,51)
}
$script:CurrentPage=0; $script:TotalPages=7; $script:MountPath=$null; $script:FeatureActions=@{}
$script:FilesToCopy=@(); $script:RegFilesToEmbed=@(); $script:ScannedFeatures=@(); $script:ScannedPackages=@()
$script:ScannedServices=@(); $script:HasScanned=$false
$script:SvcChecked=@{}; $script:SvcSortCol=0; $script:SvcSortAsc=$true
$script:PageNames = @("ISO Selection","Basic Settings","Files & Registry","Windows Features","AppX Packages","Services","Install")
$script:KeyboardLayouts = @(
    @{Id="00000409";Name="US English"},@{Id="00000809";Name="UK English"},@{Id="0000040C";Name="French"},
    @{Id="00000407";Name="German"},@{Id="0000040A";Name="Spanish"},@{Id="00000410";Name="Italian"},
    @{Id="00000416";Name="Portuguese (Brazil)"},@{Id="00000419";Name="Russian"},@{Id="00000411";Name="Japanese"},
    @{Id="00000412";Name="Korean"},@{Id="00000804";Name="Chinese (Simplified)"},@{Id="00000401";Name="Arabic"},
    @{Id="0000040D";Name="Hebrew"},@{Id="0000041F";Name="Turkish Q"},@{Id="00010415";Name="Polish"},
    @{Id="0000041D";Name="Swedish"},@{Id="00000414";Name="Norwegian"},@{Id="00000406";Name="Danish"},
    @{Id="0000040B";Name="Finnish"},@{Id="00000413";Name="Dutch"},@{Id="00000405";Name="Czech"},
    @{Id="0000040E";Name="Hungarian"},@{Id="00000408";Name="Greek"},@{Id="00000816";Name="Portuguese"}
)
#endregion

#region UI Factory Functions
function New-Ctrl($Type,$Props) {
    $c = New-Object "System.Windows.Forms.$Type"
    $Props.GetEnumerator() | ForEach-Object { $c.($_.Key) = $_.Value }
    return $c
}
function New-Btn($Text,$Loc,$Size,$Primary=$false,$Secondary=$false) {
    $b = New-Ctrl Button @{Text=$Text;Location=(New-Object System.Drawing.Point $Loc[0],$Loc[1]);Size=(New-Object System.Drawing.Size $Size[0],$Size[1]);FlatStyle="Flat";Cursor=[System.Windows.Forms.Cursors]::Hand}
    $b.FlatAppearance.BorderSize = 0
    $b.BackColor = if($Primary){$script:Colors.Primary}elseif($Secondary){$script:Colors.Secondary}else{$script:Colors.SurfaceLight}
    $b.ForeColor = $script:Colors.Text
    return $b
}
function New-Lbl($Text,$Loc,$Size,$Secondary=$true) {
    return New-Ctrl Label @{Text=$Text;Location=(New-Object System.Drawing.Point $Loc[0],$Loc[1]);Size=(New-Object System.Drawing.Size $Size[0],$Size[1]);ForeColor=if($Secondary){$script:Colors.TextSecondary}else{$script:Colors.Text}}
}
function New-Txt($Loc,$Size,$ReadOnly=$false) {
    $t = New-Ctrl TextBox @{Location=(New-Object System.Drawing.Point $Loc[0],$Loc[1]);Size=(New-Object System.Drawing.Size $Size[0],$Size[1]);BackColor=$script:Colors.SurfaceLight;ForeColor=$script:Colors.Text;BorderStyle="FixedSingle";ReadOnly=$ReadOnly}
    return $t
}
function New-Cmb($Loc,$Size) {
    return New-Ctrl ComboBox @{Location=(New-Object System.Drawing.Point $Loc[0],$Loc[1]);Size=(New-Object System.Drawing.Size $Size[0],$Size[1]);BackColor=$script:Colors.SurfaceLight;ForeColor=$script:Colors.Text;FlatStyle="Flat";DropDownStyle="DropDownList"}
}
function New-Grp($Text,$Loc,$Size) {
    $g = New-Ctrl GroupBox @{Text="  $Text  ";Location=(New-Object System.Drawing.Point $Loc[0],$Loc[1]);Size=(New-Object System.Drawing.Size $Size[0],$Size[1]);ForeColor=$script:Colors.Text;BackColor=$script:Colors.Surface}
    return $g
}
function New-Lsv($Loc,$Size,$Cols,$Chk=$false) {
    $lv = New-Ctrl ListView @{Location=(New-Object System.Drawing.Point $Loc[0],$Loc[1]);Size=(New-Object System.Drawing.Size $Size[0],$Size[1]);View="Details";FullRowSelect=$true;GridLines=$true;CheckBoxes=$Chk;BackColor=$script:Colors.SurfaceLight;ForeColor=$script:Colors.Text;BorderStyle="FixedSingle"}
    $Cols | ForEach-Object { $lv.Columns.Add($_[0],$_[1]) | Out-Null }
    return $lv
}
function New-Lbx($Loc,$Size) {
    return New-Ctrl ListBox @{Location=(New-Object System.Drawing.Point $Loc[0],$Loc[1]);Size=(New-Object System.Drawing.Size $Size[0],$Size[1]);BackColor=$script:Colors.SurfaceLight;ForeColor=$script:Colors.Text;BorderStyle="FixedSingle";SelectionMode="MultiExtended";HorizontalScrollbar=$true}
}
function New-Pg { return New-Ctrl Panel @{Location=(New-Object System.Drawing.Point 0,0);Size=(New-Object System.Drawing.Size 750,460);BackColor=$script:Colors.Background;Visible=$false} }
#endregion

#region Core Functions
function Get-AvailableDrives {
    try { Get-Volume | Where-Object {$_.DriveType -eq 'Fixed' -and $_.DriveLetter -and $_.DriveLetter -ne 'C'} | Select-Object @{N="DriveLetter";E={$_.DriveLetter}},@{N="Label";E={$_.FileSystemLabel}},@{N="SizeGB";E={[math]::Round($_.Size/1GB,2)}},@{N="FreeSpaceGB";E={[math]::Round($_.SizeRemaining/1GB,2)}} | Sort-Object DriveLetter }
    catch { Get-CimInstance Win32_LogicalDisk | Where-Object {$_.DriveType -eq 3 -and $_.DeviceID -ne "C:"} | Select-Object @{N="DriveLetter";E={$_.DeviceID -replace ':'}},@{N="Label";E={$_.VolumeName}},@{N="SizeGB";E={[math]::Round($_.Size/1GB,2)}},@{N="FreeSpaceGB";E={[math]::Round($_.FreeSpace/1GB,2)}} | Sort-Object DriveLetter }
}
function Get-WindowsImageInfo($ImagePath) {
    # DISM PowerShell cmdlet: no process spawn, locale-independent
    try { $r=@(Get-WindowsImage -ImagePath $ImagePath -EA Stop | Sort-Object ImageIndex | ForEach-Object { [PSCustomObject]@{Index=[int]$_.ImageIndex;Name=$_.ImageName} }); if($r.Count -gt 0){return $r} } catch {}
    $imgs=@(); $idx=0
    & dism.exe /Get-WimInfo /WimFile:"$ImagePath" 2>&1 | ForEach-Object {
        $l="$_".Trim()
        if($l -match "^Index\s*:\s*(\d+)"){$idx=[int]$matches[1]}
        elseif($l -match "^Name\s*:\s*(.+)"){$imgs+=[PSCustomObject]@{Index=$idx;Name=$matches[1].Trim()}}
    }
    return $imgs
}
function Get-ImageFeaturesFromMount($MountPath) {
    $feats=@(); $inTbl=$false
    & dism.exe /Image:"$MountPath" /Get-Features /Format:Table 2>&1 | ForEach-Object {
        if($_ -match "^-+"){$inTbl=$true}
        elseif($inTbl -and $_ -match "^\s*(\S+)\s*\|\s*(\S+)"){
            $n=$matches[1].Trim(); $s=$matches[2].Trim()
            if($n -and $n -ne "Feature Name"){$feats+=[PSCustomObject]@{Name=$n;State=$s;Enabled=($s -eq "Enabled")}}
        }
    }
    return $feats
}
function Get-ImagePackagesFromMount($MountPath) {
    $pkgs=@(); $cur=@{}
    & dism.exe /Image:"$MountPath" /Get-ProvisionedAppxPackages 2>&1 | ForEach-Object {
        if($_ -match "DisplayName\s*:\s*(.+)"){$cur.DisplayName=$matches[1].Trim()}
        elseif($_ -match "PackageName\s*:\s*(.+)"){$cur.PackageName=$matches[1].Trim()}
        elseif($_ -match "Version\s*:\s*(.+)"){
            $cur.Version=$matches[1].Trim()
            if($cur.DisplayName -and $cur.PackageName){$pkgs+=[PSCustomObject]$cur}
            $cur=@{}
        }
    }
    return $pkgs
}
function Resolve-ImageResourceString($Ref,$MountPath,$LibCache) {
    # Resolves "@%SystemRoot%\system32\foo.dll,-123" refs against the IMAGE's binaries
    $Ref = "$Ref"
    if(-not $Ref){return ""}
    if(-not $Ref.StartsWith('@')){return $Ref}
    # Driver INF refs carry literal fallback text: "@x.inf,%token%;Actual Name"
    $ci = $Ref.LastIndexOf(',')
    if($ci -lt 1){if($Ref -match ';(.+)$'){return $matches[1]}; return ""}
    $dll = $Ref.Substring(1,$ci-1).Trim('"')
    $id = 0
    if(-not [int]::TryParse($Ref.Substring($ci+1),[ref]$id)){if($Ref -match ';(.+)$'){return $matches[1]}; return ""}
    $id = [math]::Abs($id)
    $dll = $dll -replace '(?i)%SystemRoot%|%windir%',"$MountPath\Windows"
    if($dll -notmatch '[\\/]'){$dll = Join-Path "$MountPath\Windows\System32" $dll}
    if(-not $LibCache.ContainsKey($dll)){
        $LibCache[$dll] = if(Test-Path $dll){[WinInstall.Native]::LoadLibraryExW($dll,[IntPtr]::Zero,0x22)}else{[IntPtr]::Zero}
    }
    if($LibCache[$dll] -eq [IntPtr]::Zero){return ""}
    $sb = New-Object Text.StringBuilder 4096
    $len = [WinInstall.Native]::LoadStringW($LibCache[$dll],$id,$sb,$sb.Capacity)
    if($len -gt 0){return $sb.ToString()}
    return ""
}
function Get-ImageServicesFromMount($MountPath) {
    $svcs=@()
    $hiveSrc = Join-Path $MountPath "Windows\System32\config\SYSTEM"
    if(-not (Test-Path $hiveSrc)){return $svcs}
    # reg load needs write access; the scan mount is read-only, so work on a temp copy
    $hive = Join-Path $env:TEMP "WinInstallScanSYSTEM"
    Copy-Item $hiveSrc $hive -Force
    & reg.exe load "HKLM\WinInstallScanSys" $hive 2>&1|Out-Null
    if($LASTEXITCODE -ne 0){Remove-Item $hive -Force -EA SilentlyContinue; return $svcs}
    $libs=@{}
    try {
        Get-ChildItem "HKLM:\WinInstallScanSys\ControlSet001\Services" -EA SilentlyContinue | ForEach-Object {
            $p = Get-ItemProperty $_.PSPath -EA SilentlyContinue
            if($null -eq $p.Start){return}
            $svcs += [PSCustomObject]@{
                Name=$_.PSChildName
                DisplayName=(Resolve-ImageResourceString $p.DisplayName $MountPath $libs)
                Description=(Resolve-ImageResourceString $p.Description $MountPath $libs)
                Start=$p.Start
            }
        }
    } finally {
        $libs.Values | ForEach-Object { if($_ -ne [IntPtr]::Zero){[WinInstall.Native]::FreeLibrary($_)|Out-Null} }
        [gc]::Collect(); Start-Sleep -Milliseconds 300
        & reg.exe unload "HKLM\WinInstallScanSys" 2>&1|Out-Null
        Remove-Item $hive -Force -EA SilentlyContinue
    }
    return $svcs
}
function Mount-WimForScan($ImageFile,$Index) {
    $mp="C:\WinInstallMount_$([Guid]::NewGuid().ToString('N').Substring(0,8))"
    & dism.exe /Cleanup-Wim 2>&1 | Out-Null
    Remove-Item $mp -Recurse -Force -EA SilentlyContinue; New-Item -ItemType Directory $mp -Force | Out-Null
    & dism.exe /Mount-Wim /WimFile:"$ImageFile" /Index:$Index /MountDir:"$mp" /ReadOnly 2>&1 | Out-Null
    if($LASTEXITCODE -eq 0){return $mp}
    # DISM may refuse to mount from read-only ISO media; copy the WIM locally and retry
    $local = Join-Path $env:TEMP "WinInstallScan.wim"
    if(-not (Test-Path $local) -or (Get-Item $local).Length -ne (Get-Item $ImageFile).Length){ Copy-Item $ImageFile $local -Force }
    & dism.exe /Mount-Wim /WimFile:"$local" /Index:$Index /MountDir:"$mp" /ReadOnly 2>&1 | Out-Null
    if($LASTEXITCODE -eq 0){return $mp}
    Remove-Item $mp -Recurse -Force -EA SilentlyContinue; return $null
}
function Wait-ProcessResponsive($Proc) {
    while(-not $Proc.HasExited){ [System.Windows.Forms.Application]::DoEvents(); Start-Sleep -Milliseconds 150 }
    return $Proc.ExitCode
}
# Start-Process -PassThru without -Wait does not reliably capture ExitCode.
# Launch via System.Diagnostics.Process (through cmd /c) so the exit code is real.
function Start-Dism($ArgLine,$LogFile) {
    $psi = New-Object System.Diagnostics.ProcessStartInfo
    $psi.FileName = $env:ComSpec
    $psi.Arguments = '/c dism.exe ' + $ArgLine + $(if($LogFile){" > `"$LogFile`" 2>&1"}else{' >nul 2>&1'})
    $psi.UseShellExecute = $false; $psi.CreateNoWindow = $true
    $p = New-Object System.Diagnostics.Process; $p.StartInfo = $psi; $null = $p.Start()
    return $p
}
function Wait-DismResponsive($Proc) {
    while(-not $Proc.HasExited){ [System.Windows.Forms.Application]::DoEvents(); Start-Sleep -Milliseconds 150 }
    $Proc.WaitForExit(); return $Proc.ExitCode
}
function Read-LastPercent($LogFile) {
    try {
        $fs = New-Object IO.FileStream($LogFile,[IO.FileMode]::Open,[IO.FileAccess]::Read,[IO.FileShare]::ReadWrite)
        $sr = New-Object IO.StreamReader($fs); $txt = $sr.ReadToEnd(); $sr.Close(); $fs.Close()
        $m = [regex]::Matches($txt,'(\d{1,3}(?:\.\d)?)%')
        if($m.Count -gt 0){ return [double]$m[$m.Count-1].Groups[1].Value }
    } catch {}
    return $null
}
function Dismount-Wim($MountPath) {
    if($MountPath -and (Test-Path $MountPath)){& dism.exe /Unmount-Wim /MountDir:"$MountPath" /Discard 2>&1|Out-Null; Start-Sleep 1; Remove-Item $MountPath -Recurse -Force -EA SilentlyContinue}
    & dism.exe /Cleanup-Wim 2>&1 | Out-Null
}
function Get-FeatureDeps($Name) {
    $deps=@{"DirectPlay"=@("LegacyComponents");"IIS-WebServer"=@("IIS-WebServerRole");"Microsoft-Hyper-V"=@("Microsoft-Hyper-V-All");"WCF-Services45"=@("NetFx4-AdvSrvs");"SMB1Protocol-Client"=@("SMB1Protocol");"MicrosoftWindowsPowerShellV2"=@("MicrosoftWindowsPowerShellV2Root")}
    if($deps[$Name]){return $deps[$Name]}; return @()
}
function Resolve-FeatureDeps($Actions) {
    $res=[ordered]@{}; $done=@{}
    $Actions.Keys|Where-Object{$Actions[$_] -eq "Enable"}|ForEach-Object{Get-FeatureDeps $_|Where-Object{-not $done[$_]}|ForEach-Object{$res[$_]="Enable";$done[$_]=$true}}
    $Actions.Keys|ForEach-Object{$res[$_]=$Actions[$_]}
    return $res
}
function New-AutounattendXml($U,$P,$KB,$TZ) {
    $U = [System.Security.SecurityElement]::Escape($U); $P = [System.Security.SecurityElement]::Escape($P)
    $il = if($KB -and $KB.Length -ge 4){"$($KB.Substring($KB.Length-4)):$KB"}else{$KB}
    $ua=if($U){"<UserAccounts><LocalAccounts><LocalAccount wcm:action=`"add`">$(if($P){"<Password><Value>$P</Value><PlainText>true</PlainText></Password>"})<Description>Admin</Description><DisplayName>$U</DisplayName><Group>Administrators</Group><Name>$U</Name></LocalAccount></LocalAccounts></UserAccounts>"}else{""}
    $xml=@"
<?xml version="1.0" encoding="utf-8"?><unattend xmlns="urn:schemas-microsoft-com:unattend" xmlns:wcm="http://schemas.microsoft.com/WMIConfig/2002/State"><settings pass="oobeSystem"><component name="Microsoft-Windows-Shell-Setup" processorArchitecture="amd64" publicKeyToken="31bf3856ad364e35" language="neutral" versionScope="nonSxS"><OOBE><HideEULAPage>true</HideEULAPage><HideOEMRegistrationScreen>true</HideOEMRegistrationScreen><HideOnlineAccountScreens>true</HideOnlineAccountScreens><HideWirelessSetupInOOBE>true</HideWirelessSetupInOOBE><ProtectYourPC>1</ProtectYourPC><SkipUserOOBE>true</SkipUserOOBE><SkipMachineOOBE>true</SkipMachineOOBE></OOBE>$ua<TimeZone>$TZ</TimeZone></component><component name="Microsoft-Windows-International-Core" processorArchitecture="amd64" publicKeyToken="31bf3856ad364e35" language="neutral" versionScope="nonSxS"><InputLocale>$il</InputLocale><SystemLocale>en-US</SystemLocale><UILanguage>en-US</UILanguage><UserLocale>en-US</UserLocale></component></settings></unattend>
"@
    $op=Join-Path $env:TEMP "autounattend.xml"; $xml|Out-File $op -Encoding utf8; return $op
}
#endregion

#region Build Form
$form = New-Ctrl Form @{Text="Windows Installation Tool";Size=(New-Object System.Drawing.Size 750,620);StartPosition="CenterScreen";FormBorderStyle="FixedSingle";MaximizeBox=$false;BackColor=$script:Colors.Background}
$form.Font = New-Object System.Drawing.Font("Segoe UI",9)
$script:ScriptDir = if ($env:SCRIPT_ORIGIN) { $env:SCRIPT_ORIGIN.TrimEnd('\') } else { Split-Path -Parent ([System.Diagnostics.Process]::GetCurrentProcess().MainModule.FileName) }

# Title Panel
$titlePanel = New-Ctrl Panel @{Location=(New-Object System.Drawing.Point 0,0);Size=(New-Object System.Drawing.Size 750,60);BackColor=$script:Colors.Primary}
$titlePanel.Controls.Add((New-Ctrl Label @{Text="Windows Installation Tool";Font=(New-Object System.Drawing.Font("Segoe UI Light",20));ForeColor=$script:Colors.Text;Location=(New-Object System.Drawing.Point 20,12);Size=(New-Object System.Drawing.Size 350,36);BackColor=[System.Drawing.Color]::Transparent}))
$pageIndicatorLabel = New-Ctrl Label @{Text="Step 1 of 6: ISO Selection";Font=(New-Object System.Drawing.Font("Segoe UI",10));ForeColor=$script:Colors.TextSecondary;Location=(New-Object System.Drawing.Point 450,20);Size=(New-Object System.Drawing.Size 280,20);TextAlign=[System.Drawing.ContentAlignment]::MiddleRight;BackColor=[System.Drawing.Color]::Transparent}
$titlePanel.Controls.Add($pageIndicatorLabel)
$form.Controls.Add($titlePanel)

$contentPanel = New-Ctrl Panel @{Location=(New-Object System.Drawing.Point 0,60);Size=(New-Object System.Drawing.Size 750,460);BackColor=$script:Colors.Background}
$form.Controls.Add($contentPanel)

$navPanel = New-Ctrl Panel @{Location=(New-Object System.Drawing.Point 0,520);Size=(New-Object System.Drawing.Size 750,60);BackColor=$script:Colors.Surface}
$prevButton = New-Btn "< Previous" @(20,12) @(120,36) -Secondary:$true; $prevButton.Enabled=$false
$nextButton = New-Btn "Next >" @(480,12) @(120,36) -Primary:$true
$cancelButton = New-Btn "Exit" @(610,12) @(120,36) -Secondary:$true
$statusLabel = New-Lbl "Ready" @(150,18) @(320,24)
$navPanel.Controls.AddRange(@($prevButton,$nextButton,$cancelButton,$statusLabel))
$form.Controls.Add($navPanel)
$installButton = New-Btn "Install Windows" @(480,12) @(120,36) -Primary:$true; $installButton.Visible=$false
#endregion

#region Page 1 - ISO Selection
$page1 = New-Pg; $page1.Visible=$true
$g1a = New-Grp "Select Windows ISO" @(20,20) @(695,100)
$isoPathTextBox = New-Txt @(15,35) @(560,28) -ReadOnly:$true; $browseButton = New-Btn "Browse..." @(585,33) @(95,32) -Primary:$true
$g1a.Controls.AddRange(@($isoPathTextBox,$browseButton,(New-Lbl "Select a Windows ISO file to begin." @(15,70) @(560,20))))
$g1b = New-Grp "Windows Edition" @(20,130) @(695,80)
$editionComboBox = New-Cmb @(15,35) @(450,28); $editionComboBox.Enabled=$false
$scanButton = New-Btn "Scan Image" @(475,33) @(100,32); $scanButton.Enabled=$false
$g1b.Controls.AddRange(@($editionComboBox,$scanButton,(New-Lbl "Scan to load features/packages" @(585,38) @(100,40))))
$g1c = New-Grp "Target Drive (C: excluded)" @(20,220) @(695,80)
$driveComboBox = New-Cmb @(15,35) @(350,28); $refreshDrivesButton = New-Btn "Refresh" @(375,33) @(90,32)
$driveWarningLabel = New-Lbl "WARNING: All data will be erased!" @(480,40) @(200,20); $driveWarningLabel.ForeColor=$script:Colors.Warning
$g1c.Controls.AddRange(@($driveComboBox,$refreshDrivesButton,$driveWarningLabel))
$scanProgressBar = New-Ctrl ProgressBar @{Location=(New-Object System.Drawing.Point 20,320);Size=(New-Object System.Drawing.Size 695,25);Style="Marquee";Visible=$false}
$scanStatusLabel = New-Lbl "" @(20,350) @(695,20)
$page1.Controls.AddRange(@($g1a,$g1b,$g1c,$scanProgressBar,$scanStatusLabel))
$contentPanel.Controls.Add($page1)
#endregion

#region Page 2 - Basic Settings
$page2 = New-Pg
$g2a = New-Grp "Local Account Setup (Optional)" @(20,20) @(695,120)
$usernameTextBox = New-Txt @(100,38) @(250,28); $passwordTextBox = New-Txt @(445,38) @(180,28); $passwordTextBox.UseSystemPasswordChar=$true
$showPwdChk = New-Ctrl CheckBox @{Text="Show";Location=(New-Object System.Drawing.Point 635,40);Size=(New-Object System.Drawing.Size 55,20);ForeColor=$script:Colors.TextSecondary;FlatStyle="Flat"}
$g2a.Controls.AddRange(@((New-Lbl "Username:" @(15,40) @(80,20)),$usernameTextBox,(New-Lbl "Password:" @(370,40) @(70,20)),$passwordTextBox,$showPwdChk,(New-Lbl "Leave empty for default OOBE." @(15,80) @(665,20))))
$g2b = New-Grp "Regional Settings" @(20,150) @(695,100)
$keyboardComboBox = New-Cmb @(130,38) @(200,28); $timezoneComboBox = New-Cmb @(425,38) @(255,28)
$g2b.Controls.AddRange(@((New-Lbl "Keyboard Layout:" @(15,40) @(110,20)),$keyboardComboBox,(New-Lbl "Timezone:" @(350,40) @(70,20)),$timezoneComboBox))
$g2c = New-Grp "Autounattend.xml" @(20,260) @(695,80)
$useIsoAutoChk = New-Ctrl CheckBox @{Text="Use autounattend.xml from ISO";Location=(New-Object System.Drawing.Point 15,35);Size=(New-Object System.Drawing.Size 350,25);ForeColor=$script:Colors.Text;FlatStyle="Flat";Enabled=$false}
$autoInfoLabel = New-Lbl "Not detected in ISO" @(380,38) @(300,20)
$g2c.Controls.AddRange(@($useIsoAutoChk,$autoInfoLabel))
$page2.Controls.AddRange(@($g2a,$g2b,$g2c))
$contentPanel.Controls.Add($page2)
#endregion

#region Page 3 - Files & Registry
$page3 = New-Pg
$g3a = New-Grp "Files to Copy to Desktop" @(20,20) @(695,200)
$filesListBox = New-Lbx @(15,30) @(560,150)
$addFilesBtn = New-Btn "Add Files" @(590,30) @(90,32); $removeFilesBtn = New-Btn "Remove" @(590,70) @(90,32); $clearFilesBtn = New-Btn "Clear All" @(590,110) @(90,32)
$g3a.Controls.AddRange(@($filesListBox,$addFilesBtn,$removeFilesBtn,$clearFilesBtn))
$g3b = New-Grp "Registry Files to Apply" @(20,230) @(695,200)
$regListBox = New-Lbx @(15,30) @(560,150)
$addRegBtn = New-Btn "Add .reg" @(590,30) @(90,32); $removeRegBtn = New-Btn "Remove" @(590,70) @(90,32); $clearRegBtn = New-Btn "Clear All" @(590,110) @(90,32)
$g3b.Controls.AddRange(@($regListBox,$addRegBtn,$removeRegBtn,$clearRegBtn))
$page3.Controls.AddRange(@($g3a,$g3b))
$contentPanel.Controls.Add($page3)
#endregion

#region Page 4 - Features
$page4 = New-Pg
$g4 = New-Grp "Windows Features (Scanned from Image)" @(20,10) @(695,430)
$featuresListView = New-Lsv @(15,55) @(560,355) @(("Feature Name",280),("State",100),("Action",80))
$enableFeatBtn = New-Btn "Enable" @(590,55) @(90,32); $disableFeatBtn = New-Btn "Disable" @(590,95) @(90,32); $clearFeatBtn = New-Btn "Clear All" @(590,135) @(90,32)
$featuresCountLabel = New-Lbl "Scan image first" @(590,180) @(90,60)
$g4.Controls.AddRange(@((New-Lbl "Select features to Enable/Disable" @(15,25) @(400,20)),$featuresListView,$enableFeatBtn,$disableFeatBtn,$clearFeatBtn,$featuresCountLabel))
$page4.Controls.Add($g4)
$contentPanel.Controls.Add($page4)
#endregion

#region Page 5 - Debloat (Tabbed Interface)
$page5 = New-Pg
$g5 = New-Grp "AppX Packages to Remove (Scanned from Image)" @(20,10) @(695,430)
$packagesListView = New-Lsv @(15,30) @(560,375) @(("Package",350),("Version",130)) -Chk:$true
$selectAllPkgBtn = New-Btn "Select All" @(590,30) @(90,28); $clearPkgBtn = New-Btn "Clear All" @(590,65) @(90,28)
$packagesCountLabel = New-Lbl "Scan first" @(590,100) @(90,50)
$g5.Controls.AddRange(@($packagesListView,$selectAllPkgBtn,$clearPkgBtn,$packagesCountLabel))
$page5.Controls.Add($g5)
$contentPanel.Controls.Add($page5)
#endregion

#region Page 6 - Services
$page5svc = New-Pg
$g5svc = New-Grp "Services to Disable (Scanned from Image)" @(20,10) @(695,430)
$svcSearchBox = New-Txt @(75,28) @(500,28)
$servicesListView = New-Lsv @(15,62) @(560,343) @(("Service",120),("Display Name",185),("Description",235)) -Chk:$true
$selectAllSvcBtn = New-Btn "Select All" @(590,30) @(90,28); $clearSvcBtn = New-Btn "Clear All" @(590,65) @(90,28)
$servicesCountLabel = New-Lbl "Scan first" @(590,100) @(90,50)
$svcHintLabel = New-Lbl "Click a column header to sort. Search filters all columns. Checked services get Start=4 (disabled) in the new install." @(590,160) @(90,180)
$g5svc.Controls.AddRange(@((New-Lbl "Search:" @(15,32) @(55,20)),$svcSearchBox,$servicesListView,$selectAllSvcBtn,$clearSvcBtn,$servicesCountLabel,$svcHintLabel))
$page5svc.Controls.Add($g5svc)
$contentPanel.Controls.Add($page5svc)
#endregion

#region Page 7 - Install
$page6 = New-Pg
$g6a = New-Grp "Installation Summary" @(20,10) @(695,300)
$summaryTextBox = New-Txt @(15,30) @(665,255); $summaryTextBox.Multiline=$true; $summaryTextBox.ScrollBars="Vertical"; $summaryTextBox.Font=New-Object System.Drawing.Font("Consolas",9)
$g6a.Controls.Add($summaryTextBox)
$g6b = New-Grp "Progress" @(20,320) @(695,120)
$installProgressBar = New-Ctrl ProgressBar @{Location=(New-Object System.Drawing.Point 15,35);Size=(New-Object System.Drawing.Size 665,28);Style="Continuous"}
$installStatusLabel = New-Lbl "Ready to install." @(15,75) @(665,30)
$g6b.Controls.AddRange(@($installProgressBar,$installStatusLabel))
$page6.Controls.AddRange(@($g6a,$g6b))
$contentPanel.Controls.Add($page6)

$script:Pages = @($page1,$page2,$page3,$page4,$page5,$page5svc,$page6)
#endregion

#region Navigation & Events
function Show-Page($Idx) {
    $script:Pages | ForEach-Object { $_.Visible = $false }
    $script:Pages[$Idx].Visible = $true
    $script:CurrentPage = $Idx
    $pageIndicatorLabel.Text = "Step $($Idx+1) of $($script:TotalPages): $($script:PageNames[$Idx])"
    $prevButton.Enabled = ($Idx -gt 0)
    if($Idx -eq 6) { $nextButton.Visible=$false; $navPanel.Controls.Add($installButton); $installButton.Visible=$true; Update-Summary }
    else { $nextButton.Visible=$true; $installButton.Visible=$false }
}
function Update-Summary {
    $s = "=== INSTALLATION SUMMARY ===`r`n`r`nISO: $($isoPathTextBox.Text)`r`nEdition: $($editionComboBox.SelectedItem)`r`nDrive: $($driveComboBox.SelectedItem)`r`n"
    $s += "`r`n--- Account ---`r`nUsername: $(if($usernameTextBox.Text){$usernameTextBox.Text}else{'(default OOBE)'})`r`n"
    $s += "`r`n--- Regional ---`r`nKeyboard: $($keyboardComboBox.SelectedItem)`r`nTimezone: $($timezoneComboBox.SelectedItem)`r`n"
    $s += "`r`n--- Files ---`r`n$(if($script:FilesToCopy.Count){($script:FilesToCopy|ForEach-Object{"  - "+[IO.Path]::GetFileName($_)})-join"`r`n"}else{'  (none)'})`r`n"
    $s += "`r`n--- Registry ---`r`n$(if($script:RegFilesToEmbed.Count){($script:RegFilesToEmbed|ForEach-Object{"  - "+[IO.Path]::GetFileName($_)})-join"`r`n"}else{'  (none)'})`r`n"
    $fcs = $script:FeatureActions.Keys|Where-Object{$script:FeatureActions[$_]}
    $s += "`r`n--- Features ---`r`n$(if($fcs.Count){($fcs|ForEach-Object{"  [$($script:FeatureActions[$_])] $_"})-join"`r`n"}else{'  (no changes)'})`r`n"
    $pkgCnt = ($packagesListView.Items|Where-Object{$_.Checked}).Count
    $s += "`r`n--- AppX Packages to Remove ---`r`n  $pkgCnt package(s)`r`n"
    $svcSel = @($script:SvcChecked.Keys|Where-Object{$script:SvcChecked[$_]}|Sort-Object)
    $s += "`r`n--- Services to Disable ---`r`n$(if($svcSel.Count){($svcSel|ForEach-Object{"  - "+$_})-join"`r`n"}else{'  (none)'})`r`n"
    $summaryTextBox.Text = $s
}
function Update-DriveList {
    $driveComboBox.Items.Clear()
    $script:AvailableDrives = Get-AvailableDrives
    $script:AvailableDrives | ForEach-Object { $driveComboBox.Items.Add("$($_.DriveLetter): $(if($_.Label){$_.Label+' '})($($_.SizeGB)GB)") | Out-Null }
    if($driveComboBox.Items.Count -gt 0){$driveComboBox.SelectedIndex=0}
}

# Populate combos
$script:KeyboardLayouts | ForEach-Object { $keyboardComboBox.Items.Add($_.Name) | Out-Null }; $keyboardComboBox.SelectedIndex=0
$script:Timezones = Get-TimeZone -ListAvailable | Sort-Object BaseUtcOffset
$script:Timezones | ForEach-Object { $timezoneComboBox.Items.Add($_.DisplayName) | Out-Null }
$curTz = Get-TimeZone; for($i=0;$i -lt $script:Timezones.Count;$i++){if($script:Timezones[$i].Id -eq $curTz.Id){$timezoneComboBox.SelectedIndex=$i;break}}
if($timezoneComboBox.SelectedIndex -lt 0 -and $timezoneComboBox.Items.Count -gt 0){$timezoneComboBox.SelectedIndex=0}
Update-DriveList

function Update-ServiceList {
    # Rebuild the services list from the scanned data, applying the current search filter and sort
    $q = $svcSearchBox.Text.Trim()
    $rows = @($script:ScannedServices)
    if($q){ $rows = @($rows | Where-Object { $_.Name -like "*$q*" -or $_.DisplayName -like "*$q*" -or $_.Description -like "*$q*" }) }
    $expr = switch($script:SvcSortCol){ 1 {{"$($_.DisplayName)"}} 2 {{"$($_.Description)"}} default {{"$($_.Name)"}} }
    $rows = @($rows | Sort-Object @{Expression=$expr} -Descending:(-not $script:SvcSortAsc))
    $servicesListView.BeginUpdate(); $servicesListView.Items.Clear()
    foreach($svc in $rows){
        $itm = New-Object System.Windows.Forms.ListViewItem($svc.Name)
        $itm.SubItems.Add("$($svc.DisplayName)")|Out-Null; $itm.SubItems.Add("$($svc.Description)")|Out-Null
        $itm.ForeColor = $script:Colors.Text; $itm.Checked = [bool]$script:SvcChecked[$svc.Name]
        $servicesListView.Items.Add($itm)|Out-Null
    }
    $servicesListView.EndUpdate()
    $total = $script:ScannedServices.Count
    $servicesCountLabel.Text = if($q){"$($rows.Count)/$total services"}else{"$total services"}
}

# Button Events
$prevButton.Add_Click({
    if($script:CurrentPage -eq 6 -and -not $script:HasScanned){Show-Page 1;return}
    if($script:CurrentPage -gt 0){Show-Page ($script:CurrentPage-1)}
})
$nextButton.Add_Click({
    if($script:CurrentPage -eq 0){
        if(-not $isoPathTextBox.Text){[System.Windows.Forms.MessageBox]::Show("Select an ISO file.","Validation","OK","Warning");return}
        if($editionComboBox.SelectedIndex -lt 0){[System.Windows.Forms.MessageBox]::Show("Select an edition.","Validation","OK","Warning");return}
        if($driveComboBox.SelectedIndex -lt 0){[System.Windows.Forms.MessageBox]::Show("Select a drive.","Validation","OK","Warning");return}
    }
    # Without a scan, the customization pages are empty; jump straight to Install
    if($script:CurrentPage -eq 1 -and -not $script:HasScanned){Show-Page 6;return}
    if($script:CurrentPage -lt 6){Show-Page ($script:CurrentPage+1)}
})
$cancelButton.Add_Click({ if($script:MountPath){Dismount-Wim $script:MountPath}; $form.Close() })
$refreshDrivesButton.Add_Click({ Update-DriveList; $statusLabel.Text="Drives refreshed" })
$showPwdChk.Add_CheckedChanged({ $passwordTextBox.UseSystemPasswordChar = -not $showPwdChk.Checked })

$browseButton.Add_Click({
    $ofd = New-Object System.Windows.Forms.OpenFileDialog; $ofd.Filter="ISO Files (*.iso)|*.iso"
    if($ofd.ShowDialog() -eq "OK"){
        $isoPathTextBox.Text = $ofd.FileName; $scanStatusLabel.Text="Analyzing..."; $scanProgressBar.Visible=$true; $form.Refresh()
        # Reset state from any previously loaded ISO
        $editionComboBox.Items.Clear(); $editionComboBox.Enabled=$false; $scanButton.Enabled=$false
        $featuresListView.Items.Clear(); $packagesListView.Items.Clear(); $servicesListView.Items.Clear(); $script:FeatureActions=@{}
        $script:ScannedFeatures=@(); $script:ScannedPackages=@(); $script:ScannedServices=@(); $script:HasScanned=$false
        $script:SvcChecked=@{}; $svcSearchBox.Text=""
        $featuresCountLabel.Text="Scan image first"; $packagesCountLabel.Text="Scan first"; $servicesCountLabel.Text="Scan first"
        $useIsoAutoChk.Checked=$false; $useIsoAutoChk.Enabled=$false
        $autoInfoLabel.Text="Not detected in ISO"; $autoInfoLabel.ForeColor=$script:Colors.TextSecondary
        # Mount the ISO natively; fall back to 7-Zip extraction if ISO mounting is unavailable
        # (systems with CDFS/UDFS drivers disabled cannot mount ISOs at all)
        $script:IsoDrive = $null
        try {
            if($script:MountedIsoPath){Dismount-DiskImage -ImagePath $script:MountedIsoPath -EA SilentlyContinue|Out-Null; $script:MountedIsoPath=$null}
            Mount-DiskImage -ImagePath $ofd.FileName -EA Stop | Out-Null
            $script:MountedIsoPath = $ofd.FileName
            for($try=0;$try -lt 5 -and -not $script:IsoDrive;$try++){
                $l = (Get-DiskImage -ImagePath $ofd.FileName | Get-Volume -EA SilentlyContinue).DriveLetter
                if($l){$script:IsoDrive="${l}:"}else{Start-Sleep -Milliseconds 400}
            }
        } catch { $script:MountedIsoPath = $null }
        if($script:IsoDrive){
            $wim = "$($script:IsoDrive)\sources\install.wim"; $esd = "$($script:IsoDrive)\sources\install.esd"
        } else {
            if($script:MountedIsoPath){Dismount-DiskImage -ImagePath $script:MountedIsoPath -EA SilentlyContinue|Out-Null; $script:MountedIsoPath=$null}
            $7z = Join-Path $script:ScriptDir "7z.exe"; if(-not (Test-Path $7z)){$7z=(Get-Command "7z.exe" -EA SilentlyContinue).Source}
            if(-not $7z){[System.Windows.Forms.MessageBox]::Show("ISO mounting is disabled on this system (CDFS/UDFS drivers off) and 7z.exe was not found. Re-enable ISO mounting or place 7z.exe next to this script.","Error","OK","Error");$scanProgressBar.Visible=$false;return}
            $script:SevenZipPath = $7z
            $scanStatusLabel.Text="Native ISO mount unavailable; extracting image with 7-Zip (slower)...";$form.Refresh()
            $tmp = Join-Path $env:TEMP "WinInstallTemp"; Remove-Item $tmp -Recurse -Force -EA SilentlyContinue; New-Item $tmp -ItemType Directory -Force|Out-Null
            $p = Start-Process $7z -ArgumentList "x -y `"-o$tmp`" `"$($ofd.FileName)`" sources\install.*" -NoNewWindow -PassThru
            Wait-ProcessResponsive $p | Out-Null
            & $7z x -y -r "-o$tmp" $ofd.FileName "*unattend*.xml" 2>&1|Out-Null
            $wim = Join-Path $tmp "sources\install.wim"; $esd = Join-Path $tmp "sources\install.esd"
        }
        $script:TempImageFile = if(Test-Path $wim){$wim}elseif(Test-Path $esd){$esd}else{$null}
        if(-not $script:TempImageFile){[System.Windows.Forms.MessageBox]::Show("No image found!","Error","OK","Error");$scanProgressBar.Visible=$false;return}
        $script:IsEsdImage = $script:TempImageFile -like "*.esd"
        $script:WindowsImages = Get-WindowsImageInfo $script:TempImageFile
        $script:WindowsImages | ForEach-Object { $editionComboBox.Items.Add("$($_.Index). $($_.Name)")|Out-Null }
        if($editionComboBox.Items.Count -gt 0){$editionComboBox.SelectedIndex=0;$editionComboBox.Enabled=$true;$scanButton.Enabled=$true}
        $unattendRoot = if($script:IsoDrive){"$($script:IsoDrive)\"}else{Join-Path $env:TEMP "WinInstallTemp"}
        $isoUnattend = Get-ChildItem $unattendRoot -Filter "*unattend*.xml" -Recurse -EA SilentlyContinue|Select-Object -First 1
        $script:IsoHasAutounattend = $isoUnattend -ne $null
        if($script:IsoHasAutounattend){$script:IsoAutounattendPath=$isoUnattend.FullName;$useIsoAutoChk.Enabled=$true;$useIsoAutoChk.Checked=$true;$autoInfoLabel.Text="Found";$autoInfoLabel.ForeColor=$script:Colors.Success}
        $scanProgressBar.Visible=$false
        $scanStatusLabel.Text = if($script:IsEsdImage){"Found $($script:WindowsImages.Count) edition(s). ESD image: feature/AppX scan unavailable (install still works)."}else{"Found $($script:WindowsImages.Count) edition(s). Click Scan Image."}
    }
})

$scanButton.Add_Click({
    if($editionComboBox.SelectedIndex -lt 0){return}
    $scanButton.Enabled=$false;$scanProgressBar.Visible=$true;$scanStatusLabel.Text="Mounting...";$form.Refresh();[System.Windows.Forms.Application]::DoEvents()
    try {
        $ed = $script:WindowsImages[$editionComboBox.SelectedIndex]
        if($script:MountPath){Dismount-Wim $script:MountPath;$script:MountPath=$null}
        $script:MountPath = Mount-WimForScan $script:TempImageFile $ed.Index
        if(-not $script:MountPath){throw "Mount failed"}
        
        $scanStatusLabel.Text="Scanning features...";$form.Refresh();[System.Windows.Forms.Application]::DoEvents()
        $script:ScannedFeatures = Get-ImageFeaturesFromMount $script:MountPath
        $featuresListView.Items.Clear(); $script:FeatureActions=@{}
        $script:ScannedFeatures | ForEach-Object { $itm=New-Object System.Windows.Forms.ListViewItem($_.Name);$itm.SubItems.Add($_.State)|Out-Null;$itm.SubItems.Add("")|Out-Null;$itm.ForeColor=$script:Colors.Text;$featuresListView.Items.Add($itm)|Out-Null }
        $featuresCountLabel.Text="$($script:ScannedFeatures.Count) features"
        
        $scanStatusLabel.Text="Scanning packages...";$form.Refresh();[System.Windows.Forms.Application]::DoEvents()
        $script:ScannedPackages = Get-ImagePackagesFromMount $script:MountPath
        $packagesListView.Items.Clear()
        $script:ScannedPackages | ForEach-Object { $itm=New-Object System.Windows.Forms.ListViewItem($_.DisplayName);$itm.SubItems.Add($_.Version)|Out-Null;$itm.Tag=$_.PackageName;$itm.ForeColor=$script:Colors.Text;$packagesListView.Items.Add($itm)|Out-Null }
        $packagesCountLabel.Text="$($script:ScannedPackages.Count) pkgs"

        $scanStatusLabel.Text="Scanning services...";$form.Refresh();[System.Windows.Forms.Application]::DoEvents()
        $script:ScannedServices = Get-ImageServicesFromMount $script:MountPath
        $script:SvcChecked=@{}; $svcSearchBox.Text=""; $script:SvcSortCol=0; $script:SvcSortAsc=$true
        Update-ServiceList

        Dismount-Wim $script:MountPath; $script:MountPath=$null
        $script:HasScanned = $true
        $scanStatusLabel.Text="Scan complete!"; $statusLabel.Text="Scanned"
    } catch { $scanStatusLabel.Text="Error: $_"; if($script:MountPath){Dismount-Wim $script:MountPath;$script:MountPath=$null} }
    finally { $scanButton.Enabled=$true;$scanProgressBar.Visible=$false;$form.Refresh() }
})

# File handlers
$addFilesBtn.Add_Click({ $ofd=New-Object System.Windows.Forms.OpenFileDialog;$ofd.Multiselect=$true;if($ofd.ShowDialog()-eq"OK"){$ofd.FileNames|Where-Object{$script:FilesToCopy -notcontains $_}|ForEach-Object{$script:FilesToCopy+=$_;$filesListBox.Items.Add([IO.Path]::GetFileName($_))|Out-Null}} })
$removeFilesBtn.Add_Click({ $sel=@($filesListBox.SelectedIndices); $keep=@();for($i=0;$i -lt $script:FilesToCopy.Count;$i++){if($i -notin $sel){$keep+=$script:FilesToCopy[$i]}};$script:FilesToCopy=$keep; @($sel)|Sort-Object -Descending|ForEach-Object{$filesListBox.Items.RemoveAt($_)} })
$clearFilesBtn.Add_Click({ $script:FilesToCopy=@();$filesListBox.Items.Clear() })
$addRegBtn.Add_Click({ $ofd=New-Object System.Windows.Forms.OpenFileDialog;$ofd.Filter="Registry (*.reg)|*.reg";$ofd.Multiselect=$true;if($ofd.ShowDialog()-eq"OK"){$ofd.FileNames|Where-Object{$script:RegFilesToEmbed -notcontains $_}|ForEach-Object{$script:RegFilesToEmbed+=$_;$regListBox.Items.Add([IO.Path]::GetFileName($_))|Out-Null}} })
$removeRegBtn.Add_Click({ $sel=@($regListBox.SelectedIndices); $keep=@();for($i=0;$i -lt $script:RegFilesToEmbed.Count;$i++){if($i -notin $sel){$keep+=$script:RegFilesToEmbed[$i]}};$script:RegFilesToEmbed=$keep; @($sel)|Sort-Object -Descending|ForEach-Object{$regListBox.Items.RemoveAt($_)} })
$clearRegBtn.Add_Click({ $script:RegFilesToEmbed=@();$regListBox.Items.Clear() })

# Feature handlers
$enableFeatBtn.Add_Click({ $featuresListView.SelectedItems|ForEach-Object{$script:FeatureActions[$_.Text]="Enable";$_.SubItems[2].Text="+ Enable";$_.ForeColor=$script:Colors.Success} })
$disableFeatBtn.Add_Click({ $featuresListView.SelectedItems|ForEach-Object{$script:FeatureActions[$_.Text]="Disable";$_.SubItems[2].Text="- Disable";$_.ForeColor=$script:Colors.Error} })
$clearFeatBtn.Add_Click({ $script:FeatureActions=@{};$featuresListView.Items|ForEach-Object{$_.SubItems[2].Text="";$_.ForeColor=$script:Colors.Text} })
$script:FeatSortCol=-1; $script:FeatSortAsc=$true
$featuresListView.Add_ColumnClick({ param($s,$e)
    $col=$e.Column
    if($script:FeatSortCol -eq $col){$script:FeatSortAsc=-not $script:FeatSortAsc}else{$script:FeatSortCol=$col;$script:FeatSortAsc=$true}
    $sorted=@($featuresListView.Items)|Sort-Object @{Expression={$_.SubItems[$col].Text}} -Descending:(-not $script:FeatSortAsc)
    $featuresListView.BeginUpdate(); $featuresListView.Items.Clear(); $sorted|ForEach-Object{$featuresListView.Items.Add($_)|Out-Null}; $featuresListView.EndUpdate()
})

# Package handlers
$selectAllPkgBtn.Add_Click({ $packagesListView.Items|ForEach-Object{$_.Checked=$true} })
$clearPkgBtn.Add_Click({ $packagesListView.Items|ForEach-Object{$_.Checked=$false} })

# Service handlers
# Select All / Clear All act on the currently visible (filtered) rows
$selectAllSvcBtn.Add_Click({ $servicesListView.Items|ForEach-Object{$_.Checked=$true} })
$clearSvcBtn.Add_Click({ $servicesListView.Items|ForEach-Object{$_.Checked=$false} })
$servicesListView.Add_ItemChecked({ param($s,$e) $script:SvcChecked[$e.Item.Text]=$e.Item.Checked })
$svcSearchBox.Add_TextChanged({ Update-ServiceList })
$servicesListView.Add_ColumnClick({ param($s,$e)
    if($script:SvcSortCol -eq $e.Column){$script:SvcSortAsc=-not $script:SvcSortAsc}else{$script:SvcSortCol=$e.Column;$script:SvcSortAsc=$true}
    Update-ServiceList
})

#endregion

#region Install Handler
$installButton.Add_Click({
    if(-not $isoPathTextBox.Text -or $editionComboBox.SelectedIndex -lt 0 -or $driveComboBox.SelectedIndex -lt 0){[System.Windows.Forms.MessageBox]::Show("Missing selections","Error","OK","Warning");return}
    $drive = $script:AvailableDrives[$driveComboBox.SelectedIndex]; $dl = $drive.DriveLetter
    $ed = $script:WindowsImages[$editionComboBox.SelectedIndex]; $kb = $script:KeyboardLayouts[$keyboardComboBox.SelectedIndex]; $tz = $script:Timezones[$timezoneComboBox.SelectedIndex]

    if($drive.SizeGB -lt 20){[System.Windows.Forms.MessageBox]::Show("Drive ${dl}: is only $($drive.SizeGB)GB. Windows needs at least 20GB.","Error","OK","Error");return}

    $isoRoot = [IO.Path]::GetPathRoot($isoPathTextBox.Text).TrimEnd('\',':')
    if($isoRoot -and $isoRoot.ToUpper() -eq $dl.ToString().ToUpper()){[System.Windows.Forms.MessageBox]::Show("The ISO is located on drive ${dl}:, the selected target. Formatting it would erase the ISO. Choose a different target drive or move the ISO.","Error","OK","Error");return}

    if(@(Get-ChildItem "${dl}:\" -Force -EA SilentlyContinue).Count -gt 0){
        if([System.Windows.Forms.MessageBox]::Show("Drive $dl has data. Format?","Warning","YesNo","Warning") -ne "Yes"){return}
        if([System.Windows.Forms.MessageBox]::Show("CONFIRM format $dl!","Warning","YesNo","Exclamation") -ne "Yes"){return}
        $installStatusLabel.Text="Formatting...";$installProgressBar.Value=5;$form.Refresh()
        try{Format-Volume -DriveLetter $dl -FileSystem NTFS -Confirm:$false -Force -EA Stop}catch{[System.Windows.Forms.MessageBox]::Show("Format failed","Error","OK","Error");return}
    }
    
    $installButton.Enabled=$false;$prevButton.Enabled=$false;$cancelButton.Enabled=$false
    try {
        $installStatusLabel.Text="Preparing...";$installProgressBar.Value=5;$form.Refresh();[System.Windows.Forms.Application]::DoEvents()
        & dism.exe /Cleanup-Wim 2>&1|Out-Null
        # Apply directly from the mounted ISO; no extraction step
        if(-not $script:TempImageFile -or -not (Test-Path $script:TempImageFile)){throw "Image not found. Re-select the ISO (it may have been unmounted)."}
        $imgFile = $script:TempImageFile

        $installStatusLabel.Text="Applying image...";$installProgressBar.Value=10;$form.Refresh();[System.Windows.Forms.Application]::DoEvents()
        $applyLog = Join-Path $env:TEMP "WinInstallApply.log"; Remove-Item $applyLog -Force -EA SilentlyContinue
        $sw = [Diagnostics.Stopwatch]::StartNew()
        $p = Start-Dism "/Apply-Image /ImageFile:`"$imgFile`" /Index:$($ed.Index) /ApplyDir:${dl}:\" $applyLog
        while(-not $p.HasExited){
            $pct = Read-LastPercent $applyLog
            if($pct -ne $null){$installProgressBar.Value=[int][math]::Min(70,10+$pct*0.6);$installStatusLabel.Text="Applying image... $pct%"}
            else{$installStatusLabel.Text="Applying image... $([int]$sw.Elapsed.TotalSeconds)s"}
            [System.Windows.Forms.Application]::DoEvents(); Start-Sleep -Milliseconds 400
        }
        $p.WaitForExit(); $applyCode = $p.ExitCode
        Remove-Item $applyLog -Force -EA SilentlyContinue
        if($applyCode -ne 0){throw "DISM apply failed (exit code $applyCode). See C:\Windows\Logs\DISM\dism.log."}

        $installProgressBar.Value=70;$form.Refresh()
        
        # Features
        if($script:FeatureActions.Count -gt 0){
            $installStatusLabel.Text="Applying features...";$form.Refresh();[System.Windows.Forms.Application]::DoEvents()
            $resolved = Resolve-FeatureDeps $script:FeatureActions
            $en=@($resolved.Keys|Where-Object{$resolved[$_] -eq "Enable"}); $dis=@($resolved.Keys|Where-Object{$resolved[$_] -eq "Disable"})
            if($en.Count -gt 0){
                $p=Start-Dism ("/Image:${dl}:\ /Enable-Feature "+(($en|ForEach-Object{"/FeatureName:`"$_`""}) -join ' ')+" /All")
                Wait-DismResponsive $p|Out-Null
            }
            if($dis.Count -gt 0){
                $p=Start-Dism ("/Image:${dl}:\ /Disable-Feature "+(($dis|ForEach-Object{"/FeatureName:`"$_`""}) -join ' '))
                Wait-DismResponsive $p|Out-Null
            }
        }

        # Packages (AppX from packages tab)
        $pkgsToRemove = @($packagesListView.Items|Where-Object{$_.Checked}|ForEach-Object{$_.Tag}|Where-Object{$_})
        if($pkgsToRemove.Count -gt 0){
            $i=0
            foreach($pkg in $pkgsToRemove){
                $i++
                $installStatusLabel.Text="Removing packages ($i/$($pkgsToRemove.Count))...";$installProgressBar.Value=[int](72+8*$i/$pkgsToRemove.Count);$form.Refresh()
                $p=Start-Dism "/Image:${dl}:\ /Remove-ProvisionedAppxPackage /PackageName:`"$pkg`""
                Wait-DismResponsive $p|Out-Null
            }
        }
        
        # Services: set Start=4 in the applied image's SYSTEM hive
        $svcToDisable = @($script:SvcChecked.Keys|Where-Object{$script:SvcChecked[$_]})
        if($svcToDisable.Count -gt 0){
            $installStatusLabel.Text="Disabling services ($($svcToDisable.Count))...";$installProgressBar.Value=82;$form.Refresh();[System.Windows.Forms.Application]::DoEvents()
            $sysHive = "${dl}:\Windows\System32\config\SYSTEM"
            if(Test-Path $sysHive){
                & reg.exe load "HKLM\WinInstallOffSys" $sysHive 2>&1|Out-Null
                if($LASTEXITCODE -eq 0){
                    foreach($svc in $svcToDisable){ & reg.exe add "HKLM\WinInstallOffSys\ControlSet001\Services\$svc" /v Start /t REG_DWORD /d 4 /f 2>&1|Out-Null }
                    [gc]::Collect(); Start-Sleep -Milliseconds 300; & reg.exe unload "HKLM\WinInstallOffSys" 2>&1|Out-Null
                }
            }
        }

        # Autounattend
        $installStatusLabel.Text="Configuring...";$installProgressBar.Value=85;$form.Refresh()
        $unDir = "${dl}:\Windows\System32\sysprep"; New-Item $unDir -ItemType Directory -Force -EA SilentlyContinue|Out-Null
        if($useIsoAutoChk.Checked -and $script:IsoHasAutounattend){Copy-Item $script:IsoAutounattendPath "$unDir\unattend.xml" -Force}
        elseif($usernameTextBox.Text){$autoPath=New-AutounattendXml $usernameTextBox.Text $passwordTextBox.Text $kb.Id $tz.Id;Copy-Item $autoPath "$unDir\unattend.xml" -Force}
        
        # Files
        if($script:FilesToCopy.Count -gt 0){
            $dskPath = "${dl}:\Users\Default\Desktop"; New-Item $dskPath -ItemType Directory -Force -EA SilentlyContinue|Out-Null
            $script:FilesToCopy | ForEach-Object { Copy-Item $_ $dskPath -Force -EA SilentlyContinue }
        }
        
        # Registry
        if($script:RegFilesToEmbed.Count -gt 0){
            $regDir = "${dl}:\Windows\Setup\RegFiles"; $scrDir = "${dl}:\Windows\Setup\Scripts"
            New-Item $regDir,$scrDir -ItemType Directory -Force -EA SilentlyContinue|Out-Null
            $hklm=@(); $hkcu=@()
            $script:RegFilesToEmbed | ForEach-Object {
                Copy-Item $_ $regDir -Force; $fn=[IO.Path]::GetFileName($_); $c=Get-Content $_ -Raw -EA SilentlyContinue
                if($c -match 'HKEY_LOCAL_MACHINE|HKLM'){$hklm+=$fn}
                if($c -match 'HKEY_CURRENT_USER|HKCU'){$hkcu+=@{Name=$fn;Content=$c}}
            }
            if($hkcu.Count -gt 0){
                $hive = "${dl}:\Users\Default\NTUSER.DAT"
                if(Test-Path $hive){
                    & reg.exe load "HKU\OffDef" $hive 2>&1|Out-Null
                    if($LASTEXITCODE -eq 0){
                        $hkcu | ForEach-Object { $mod=$_.Content -replace '\[HKEY_CURRENT_USER','[HKEY_USERS\OffDef' -replace '\[HKCU','[HKEY_USERS\OffDef'; $tmp=Join-Path $env:TEMP "hkcu_$($_.Name)"; [IO.File]::WriteAllText($tmp,$mod,[Text.Encoding]::Unicode); & reg.exe import $tmp 2>&1|Out-Null; Remove-Item $tmp -Force -EA SilentlyContinue }
                        [gc]::Collect(); Start-Sleep -Milliseconds 300; & reg.exe unload "HKU\OffDef" 2>&1|Out-Null
                    }
                }
            }
            if($hklm.Count -gt 0){
                $cmd = "@echo off`r`ntimeout /t 3 /nobreak >nul`r`n" + (($hklm|ForEach-Object{"reg import `"C:\Windows\Setup\RegFiles\$_`" 2>nul"}) -join "`r`n") + "`r`n"
                $cmd | Out-File (Join-Path $scrDir "SetupComplete.cmd") -Encoding ascii
            }
        }
        
        $installStatusLabel.Text="Making bootable...";$installProgressBar.Value=92;$form.Refresh()
        & "${dl}:\Windows\System32\bcdboot.exe" "${dl}:\Windows" 2>&1|Out-Null
        if($LASTEXITCODE -ne 0){throw "bcdboot failed (code $LASTEXITCODE). No boot entry was created; the installation will not appear in the boot menu."}

        $installStatusLabel.Text="Cleanup...";$installProgressBar.Value=96;$form.Refresh()
        & dism.exe /Cleanup-Wim 2>&1|Out-Null
        
        $installProgressBar.Value=100;$installStatusLabel.Text="Complete!"
        [System.Windows.Forms.MessageBox]::Show("Windows installed to ${dl}:!`nReboot and select new installation.","Success","OK","Information")
    } catch { $installStatusLabel.Text="Error: $_";$installProgressBar.Value=0;[System.Windows.Forms.MessageBox]::Show("Failed: $_","Error","OK","Error") }
    finally { $installButton.Enabled=$true;$prevButton.Enabled=$true;$cancelButton.Enabled=$true }
})
#endregion

$form.Add_FormClosed({
    if($script:MountPath){Dismount-Wim $script:MountPath;$script:MountPath=$null}
    if($script:MountedIsoPath){Dismount-DiskImage -ImagePath $script:MountedIsoPath -EA SilentlyContinue|Out-Null}
    Remove-Item (Join-Path $env:TEMP "WinInstallScan.wim") -Force -EA SilentlyContinue
    Remove-Item (Join-Path $env:TEMP "WinInstallTemp") -Recurse -Force -EA SilentlyContinue
})

Show-Page 0
[void]$form.ShowDialog()
