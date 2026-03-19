param(
    [string]$ServerUrl   = "",
    [int]$PollInterval   = 600,
    [switch]$Uninstall   = $false,
    [string]$Title       = "DiskHealth Agent - Master Sofa"
)
function Ensure-Admin {
    $id = [System.Security.Principal.WindowsIdentity]::GetCurrent()
    $p  = New-Object System.Security.Principal.WindowsPrincipal($id)
    if (-not $p.IsInRole([System.Security.Principal.WindowsBuiltInRole]::Administrator)) {
        $argList = "-ExecutionPolicy Bypass -File `"$($MyInvocation.ScriptName)`""
        if ($ServerUrl)           { $argList += " -ServerUrl `"$ServerUrl`"" }
        if ($PollInterval -ne 30) { $argList += " -PollInterval $PollInterval" }
        if ($Uninstall)           { $argList += " -Uninstall" }
        if ($Title)               { $argList += " -Title `"$Title`"" }
        Start-Process powershell.exe -Verb RunAs -ArgumentList $argList
        exit
    }
}
$ServiceName = "DiskHealthAgent"
$InstallDir  = "$env:ProgramFiles\DiskHealthAgent"
$AgentScript = Join-Path $InstallDir "DiskHealthAgent.ps1"
function Write-Step { param([string]$m); Write-Host "  [>] $m" -ForegroundColor Cyan   }
function Write-OK   { param([string]$m); Write-Host "  [OK] $m" -ForegroundColor Green  }
function Write-Fail { param([string]$m); Write-Host "  [!!] $m" -ForegroundColor Red    }
function Write-Warn { param([string]$m); Write-Host "  [!] $m"  -ForegroundColor Yellow }
function Check-Admin {
    $id = [System.Security.Principal.WindowsIdentity]::GetCurrent()
    $p  = New-Object System.Security.Principal.WindowsPrincipal($id)
    return $p.IsInRole([System.Security.Principal.WindowsBuiltInRole]::Administrator)
}
function Get-WinMajor {
    $os = Get-WmiObject Win32_OperatingSystem
    return [int]([version]$os.Version).Major
}
function Test-ServerConn {
    param([string]$Url)
    try {
        $req = [System.Net.WebRequest]::Create("$Url/health"); $req.Timeout=5000
        $resp = $req.GetResponse(); $resp.Close(); return $true
    } catch { return $false }
}
function Uninstall-Agent {
    Write-Step "Removing scheduled task..."
    schtasks /Delete /TN $ServiceName /F 2>&1 | Out-Null
    Write-OK "Task removed."
    $procs=Get-WmiObject Win32_Process -Filter "Name='powershell.exe'" -ErrorAction SilentlyContinue
    foreach ($proc in $procs) {
        if ($proc.CommandLine -and $proc.CommandLine -like "*DiskHealthAgent.ps1*") {
            Stop-Process -Id $proc.ProcessId -Force -ErrorAction SilentlyContinue
        }
    }
    Start-Sleep -Seconds 2
    if (Test-Path $InstallDir) {
        for ($t=1;$t-le5;$t++) {
            try { Remove-Item -Recurse -Force $InstallDir -ErrorAction Stop; break }
            catch { Start-Sleep -Seconds 2 }
        }
    }
    netsh advfirewall firewall delete rule name="DiskHealthAgent" 2>&1 | Out-Null
    schtasks /Delete /TN "DiskHealthTray" /F 2>&1 | Out-Null
    Get-WmiObject Win32_Process -Filter "Name='powershell.exe'" -ErrorAction SilentlyContinue |
        Where-Object { $_.CommandLine -like "*DiskHealthTray*" } |
        ForEach-Object { Stop-Process -Id $_.ProcessId -Force -ErrorAction SilentlyContinue }
    Write-OK "Uninstall complete."
}

function Install-Smartctl {
    $smartctlPath = "$env:ProgramFiles\smartmontools\bin\smartctl.exe"
    if (Test-Path $smartctlPath) {
        Write-OK "smartmontools already installed."
        return $true
    }
    Write-Step "Installing smartmontools (required for SSD/NVMe SMART data)..."
    # Try winget first (Windows 10 1709+)
    $wingetOk = $false
    try {
        $wg = Get-Command winget.exe -ErrorAction Stop
        if ($wg) {
            Write-Step "Trying winget..."
            $proc = Start-Process -FilePath "winget.exe" `
                -ArgumentList "install --id smartmontools.smartmontools --silent --accept-package-agreements --accept-source-agreements" `
                -Wait -PassThru -WindowStyle Hidden
            if ($proc.ExitCode -eq 0 -or $proc.ExitCode -eq -1978335212) {
                # -1978335212 = already installed
                if (Test-Path $smartctlPath) {
                    Write-OK "smartmontools installed via winget."
                    $wingetOk = $true
                }
            }
        }
    } catch {}
    if ($wingetOk) { return $true }
    # Fallback: download MSI directly from smartmontools.org
    Write-Step "winget unavailable or failed, downloading MSI directly..."
    try {
        $msiUrl  = "https://www.smartmontools.org/airfiles/smartmontools-7.4-1.win32-setup.exe"
        $msiPath = Join-Path $env:TEMP "smartmontools-setup.exe"
        $wc = New-Object System.Net.WebClient
        $wc.DownloadFile($msiUrl, $msiPath)
        $proc = Start-Process -FilePath $msiPath -ArgumentList "/S" -Wait -PassThru
        Remove-Item $msiPath -Force -ErrorAction SilentlyContinue
        if (Test-Path $smartctlPath) {
            Write-OK "smartmontools installed via direct download."
            return $true
        } else {
            Write-Warn "smartmontools installer ran but binary not found. SSD/NVMe data may be unavailable."
            return $false
        }
    } catch {
        Write-Warn "Could not install smartmontools: $_"
        Write-Warn "SSD/NVMe SMART data will be unavailable. Install manually: winget install smartmontools.smartmontools"
        return $false
    }
}
function Install-Agent {
    if (-not $ServerUrl) { Write-Fail "ServerUrl is required."; exit 1 }
    if (-not $ServerUrl.StartsWith("http")) { Write-Fail "ServerUrl must start with http://"; exit 1 }
    if (-not (Check-Admin)) { Write-Fail "Must be run as Administrator."; exit 1 }
    $winVer = Get-WinMajor
    Write-Step "Creating install directory..."
    if (-not (Test-Path $InstallDir)) { New-Item -ItemType Directory -Path $InstallDir -Force | Out-Null }
    Write-OK "Directory ready."
    # Step 1: Install bundled script first (guaranteed correct version)
    Write-Step "Installing bundled agent script..."
    $scriptDir    = Split-Path -Parent $MyInvocation.ScriptName
    $sourceScript = Join-Path $scriptDir "DiskHealthAgent.ps1"
    if (-not (Test-Path $sourceScript)) { $sourceScript = Join-Path $PSScriptRoot "DiskHealthAgent.ps1" }
    if (Test-Path $sourceScript) {
        Copy-Item -Path $sourceScript -Destination $AgentScript -Force
        Write-OK "Bundled agent script installed."
    } else {
        Write-Fail "No agent script found - cannot install."; exit 1
    }
    # Step 2: Try to pull latest from server (optional upgrade, non-fatal)
    $agentUrl = "$ServerUrl/agent/agent.ps1"
    try {
        $tmp = "$AgentScript.download"
        $wc = New-Object System.Net.WebClient
        $wc.DownloadFile($agentUrl, $tmp)
        # Only replace if download succeeded and file is non-empty
        if ((Test-Path $tmp) -and (Get-Item $tmp).Length -gt 512) {
            Move-Item -Force $tmp $AgentScript
            Write-OK "Agent script updated from server."
        } else {
            Remove-Item $tmp -Force -ErrorAction SilentlyContinue
            Write-Warn "Server script empty/missing - using bundled version."
        }
    } catch {
        Write-Warn "Could not pull update from server ($_) - using bundled version."
    }
    Write-Step "Checking smartmontools..."
    Install-Smartctl | Out-Null
    $oldIdFile = Join-Path $InstallDir "agent_id.txt"
    if (Test-Path $oldIdFile) { Remove-Item $oldIdFile -Force -ErrorAction SilentlyContinue }
    try { Set-ExecutionPolicy -Scope LocalMachine -ExecutionPolicy RemoteSigned -Force } catch {}
    Write-Step "Creating scheduled task '$ServiceName'..."
    $psExe  = "powershell.exe"
    $psArgs = "-NonInteractive -WindowStyle Hidden -ExecutionPolicy Bypass -File `"$AgentScript`" -ServerUrl `"$ServerUrl`" -PollInterval $PollInterval"
    schtasks /Delete /TN $ServiceName /F 2>&1 | Out-Null
    $taskCreated = $false
    if ($winVer -ge 8) {
        $xmlEscapedArgs = [System.Security.SecurityElement]::Escape($psArgs)
        $xml = @"
<?xml version="1.0" encoding="UTF-16"?>
<Task version="1.3" xmlns="http://schemas.microsoft.com/windows/2004/02/mit/task">
  <RegistrationInfo><Description>DiskHealth Agent</Description></RegistrationInfo>
  <Triggers><BootTrigger><Enabled>true</Enabled><Delay>PT15S</Delay></BootTrigger></Triggers>
  <Principals><Principal id="Author"><UserId>SYSTEM</UserId><RunLevel>HighestAvailable</RunLevel></Principal></Principals>
  <Settings>
    <MultipleInstancesPolicy>IgnoreNew</MultipleInstancesPolicy>
    <DisallowStartIfOnBatteries>false</DisallowStartIfOnBatteries>
    <StopIfGoingOnBatteries>false</StopIfGoingOnBatteries>
    <ExecutionTimeLimit>PT0S</ExecutionTimeLimit>
    <Enabled>true</Enabled>
    <RestartOnFailure><Interval>PT1M</Interval><Count>9999</Count></RestartOnFailure>
  </Settings>
  <Actions><Exec>
    <Command>$psExe</Command>
    <Arguments>$xmlEscapedArgs</Arguments>
    <WorkingDirectory>$InstallDir</WorkingDirectory>
  </Exec></Actions>
</Task>
"@
        $xmlPath = Join-Path $env:TEMP "diskhealth_task.xml"
        [System.IO.File]::WriteAllText($xmlPath, $xml, [System.Text.Encoding]::Unicode)
        $out = schtasks /Create /TN $ServiceName /XML $xmlPath /F 2>&1
        Remove-Item $xmlPath -ErrorAction SilentlyContinue
        if ($LASTEXITCODE -eq 0) { Write-OK "Scheduled task created."; $taskCreated = $true }
        else { Write-Warn "XML method failed: $out" }
    }
    if (-not $taskCreated) {
        $out = schtasks /Create /TN $ServiceName /SC ONSTART /DELAY "0000:15" /TR "$psExe $psArgs" /RU SYSTEM /RL HIGHEST /F 2>&1
        if ($LASTEXITCODE -eq 0) { Write-OK "Scheduled task created."; $taskCreated = $true }
        else { Write-Fail "Could not create scheduled task: $out" }
    }
    Write-Step "Testing server connection..."
    if (Test-ServerConn $ServerUrl) { Write-OK "Server is reachable!" }
    else { Write-Warn "Server not reachable - agent will retry automatically." }
    Write-Step "Starting agent in background..."
    try {
        Start-Process -FilePath $psExe -ArgumentList $psArgs -WorkingDirectory $InstallDir -WindowStyle Hidden
        Write-OK "Agent launched."
    } catch { Write-Warn "Auto-start failed - will start on reboot." }
    Write-Step "Setting up system tray icon..."
    $TrayScript = Join-Path $InstallDir "DiskHealthTray.ps1"
    $TrayTask   = "DiskHealthTray"
    # Step 1: install bundled tray (guaranteed correct version)
    $bundledTray = Join-Path (Split-Path -Parent $MyInvocation.ScriptName) "DiskHealthTray.ps1"
    if (-not (Test-Path $bundledTray)) { $bundledTray = Join-Path $PSScriptRoot "DiskHealthTray.ps1" }
    if (Test-Path $bundledTray) {
        Copy-Item -Path $bundledTray -Destination $TrayScript -Force
        Write-OK "Bundled tray script installed."
    }
    # Step 2: try to pull latest from server (optional, non-fatal)
    try {
        $trayTmp = "$TrayScript.download"
        $wc2 = New-Object System.Net.WebClient
        $wc2.DownloadFile("$ServerUrl/agent/tray.ps1", $trayTmp)
        if ((Test-Path $trayTmp) -and (Get-Item $trayTmp).Length -gt 512) {
            Move-Item -Force $trayTmp $TrayScript
            Write-OK "Tray script updated from server."
        } else {
            Remove-Item $trayTmp -Force -ErrorAction SilentlyContinue
        }
    } catch { Write-Warn "Could not pull tray update from server - using bundled version." }
    if (Test-Path $TrayScript) {
        $trayArgs = "-STA -NonInteractive -ExecutionPolicy Bypass -File `"$TrayScript`""
        $trayCreated = $false
        # Use XML method to handle spaces in path correctly
        try {
            $xmlEscArgs  = [System.Security.SecurityElement]::Escape($trayArgs)
            $xmlEscExe   = [System.Security.SecurityElement]::Escape($psExe)
            $xmlEscDir   = [System.Security.SecurityElement]::Escape($InstallDir)
            $trayXml = @"
<?xml version="1.0" encoding="UTF-16"?>
<Task version="1.3" xmlns="http://schemas.microsoft.com/windows/2004/02/mit/task">
  <RegistrationInfo><Description>DiskHealth Tray Icon</Description></RegistrationInfo>
  <Triggers><LogonTrigger><Enabled>true</Enabled></LogonTrigger></Triggers>
  <Principals><Principal id="Author"><GroupId>S-1-5-32-545</GroupId><RunLevel>LeastPrivilege</RunLevel></Principal></Principals>
  <Settings>
    <MultipleInstancesPolicy>IgnoreNew</MultipleInstancesPolicy>
    <DisallowStartIfOnBatteries>false</DisallowStartIfOnBatteries>
    <StopIfGoingOnBatteries>false</StopIfGoingOnBatteries>
    <ExecutionTimeLimit>PT0S</ExecutionTimeLimit>
    <Enabled>true</Enabled>
  </Settings>
  <Actions><Exec>
    <Command>$xmlEscExe</Command>
    <Arguments>$xmlEscArgs</Arguments>
    <WorkingDirectory>$xmlEscDir</WorkingDirectory>
  </Exec></Actions>
</Task>
"@
            $trayXmlPath = Join-Path $env:TEMP "diskhealthtray_task.xml"
            [System.IO.File]::WriteAllText($trayXmlPath, $trayXml, [System.Text.Encoding]::Unicode)
            schtasks /Delete /TN $TrayTask /F 2>&1 | Out-Null
            $out = schtasks /Create /TN $TrayTask /XML $trayXmlPath /F 2>&1
            Remove-Item $trayXmlPath -ErrorAction SilentlyContinue
            if ($LASTEXITCODE -eq 0) { $trayCreated = $true }
            else { Write-Warn "Tray XML method failed: $out" }
        } catch { Write-Warn "Tray task XML error: $_" }
        if ($trayCreated) {
            Write-OK "Tray logon task created."
            # Kill any running tray instance before starting the new one
            Get-WmiObject Win32_Process -Filter "Name='powershell.exe'" -ErrorAction SilentlyContinue |
                Where-Object { $_.CommandLine -like "*DiskHealthTray*" } |
                ForEach-Object { Stop-Process -Id $_.ProcessId -Force -ErrorAction SilentlyContinue }
            Start-Sleep -Seconds 2
            try {
                Start-Process -FilePath $psExe -ArgumentList $trayArgs -WindowStyle Hidden
                Write-OK "Tray icon launched."
            } catch { Write-Warn "Tray auto-start failed - will appear on next logon." }
        } else { Write-Warn "Could not create tray task - will appear on next logon." }
    } else { Write-Warn "DiskHealthTray.ps1 not found - skipping tray setup." }
    Write-OK "Installation complete! Reporting to: $ServerUrl"
}
Ensure-Admin
if ($Uninstall) { Uninstall-Agent } else { Install-Agent }
