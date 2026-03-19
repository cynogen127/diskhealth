# ==============================================================================
#  DiskHealth Agent Installer v3.2.0
# ==============================================================================
param(
    [string]$ServerUrl    = "",
    [int]   $PollInterval = 600,
    [switch]$Uninstall
)

Set-StrictMode -Off
$ErrorActionPreference = "Continue"

trap {
    Write-Host "[!!] FATAL: $_" -ForegroundColor Red
    Write-Host "     At: $($_.InvocationInfo.PositionMessage)" -ForegroundColor DarkGray
    exit 1
}

# ── Self-elevate if not admin ─────────────────────────────────────────────────
$isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
if (-not $isAdmin) {
    $argList = "-NoProfile -NonInteractive -ExecutionPolicy Bypass -File `"$($MyInvocation.MyCommand.Path)`""
    if ($ServerUrl)    { $argList += " -ServerUrl `"$ServerUrl`"" }
    if ($PollInterval) { $argList += " -PollInterval $PollInterval" }
    if ($Uninstall)    { $argList += " -Uninstall" }
    Start-Process powershell.exe -ArgumentList $argList -Verb runas -Wait
    exit $LASTEXITCODE
}

# ── Paths ─────────────────────────────────────────────────────────────────────
$InstallDir = "$env:ProgramFiles\DiskHealthAgent"
$AgentPs1   = Join-Path $InstallDir "DiskHealthAgent.ps1"
$TrayPs1    = Join-Path $InstallDir "DiskHealthTray.ps1"
$TaskName   = "DiskHealthAgent"
$TrayTask   = "DiskHealthTray"
$LogFile    = Join-Path $InstallDir "agent.log"

# Resolve the directory this script lives in
$ScriptDir = if ($PSScriptRoot -and $PSScriptRoot -ne "" -and (Test-Path $PSScriptRoot)) {
    $PSScriptRoot
} elseif ($MyInvocation.MyCommand.Path -and $MyInvocation.MyCommand.Path -ne "") {
    Split-Path -Parent $MyInvocation.MyCommand.Path
} else {
    $PWD.Path
}

function Write-Step { param([string]$m); Write-Host "[>] $m" -ForegroundColor Cyan }
function Write-OK   { param([string]$m); Write-Host "[OK] $m" -ForegroundColor Green }
function Write-Fail { param([string]$m); Write-Host "[!!] $m" -ForegroundColor Red }
function Write-Note { param([string]$m); Write-Host "     $m" -ForegroundColor DarkGray }

function Get-InteractiveUser {
    try {
        $u = (Get-WmiObject Win32_ComputerSystem -ErrorAction Stop).UserName
        if ($u -and $u -ne "") { return $u }
    } catch { }
    return $null
}

Write-Host ""
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host ("  DiskHealth Agent — " + $(if ($Uninstall) {"Uninstaller"} else {"Installer v3.2.0"})) -ForegroundColor Cyan
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host ""

# ==============================================================================
#  UNINSTALL
# ==============================================================================
if ($Uninstall) {
    Write-Step "Stopping tasks..."
    Stop-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue
    Stop-ScheduledTask -TaskName $TrayTask -ErrorAction SilentlyContinue
    Start-Sleep -Seconds 2
    Get-WmiObject Win32_Process -Filter "Name='powershell.exe'" -ErrorAction SilentlyContinue |
        Where-Object { $_.CommandLine -like "*DiskHealthAgent*" -or $_.CommandLine -like "*DiskHealthTray*" } |
        ForEach-Object { Stop-Process -Id $_.ProcessId -Force -ErrorAction SilentlyContinue }
    schtasks /Delete /TN $TaskName /F 2>&1 | Out-Null
    schtasks /Delete /TN $TrayTask /F 2>&1 | Out-Null
    if (Test-Path $InstallDir) {
        Remove-Item $InstallDir -Recurse -Force -ErrorAction SilentlyContinue
        Write-OK "Removed $InstallDir"
    }
    Write-OK "Agent uninstalled."
    exit 0
}

# ==============================================================================
#  INSTALL
# ==============================================================================

# ── Validate URL ──────────────────────────────────────────────────────────────
if (-not $ServerUrl -or $ServerUrl -eq "") {
    $ServerUrl = Read-Host "Enter DiskHealth server URL (e.g. http://192.168.0.150:8765)"
}
$ServerUrl = $ServerUrl.Trim().TrimEnd("/")
if ($ServerUrl -notmatch "^https?://") {
    Write-Fail "Invalid server URL: '$ServerUrl'"
    exit 1
}
Write-OK "Server URL: $ServerUrl"

# ── Locate source PS1 files ───────────────────────────────────────────────────
# They should be in the same folder as this script (both extracted to temp by C# installer)
Write-Note "ScriptDir = $ScriptDir"
$SrcAgent = Join-Path $ScriptDir "DiskHealthAgent.ps1"
$SrcTray  = Join-Path $ScriptDir "DiskHealthTray.ps1"
Write-Note "Looking for: $SrcAgent"
Write-Note "Looking for: $SrcTray"

if (-not (Test-Path $SrcAgent)) {
    Write-Fail "DiskHealthAgent.ps1 not found at: $SrcAgent"
    exit 1
}
if (-not (Test-Path $SrcTray)) {
    Write-Fail "DiskHealthTray.ps1 not found at: $SrcTray"
    exit 1
}
Write-OK "Source scripts found."

# ── Stop any running agent ────────────────────────────────────────────────────
Write-Step "Stopping existing agent..."
Stop-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue
Stop-ScheduledTask -TaskName $TrayTask -ErrorAction SilentlyContinue
Start-Sleep -Seconds 2
Get-WmiObject Win32_Process -Filter "Name='powershell.exe'" -ErrorAction SilentlyContinue |
    Where-Object { $_.CommandLine -like "*DiskHealthAgent*" -or $_.CommandLine -like "*DiskHealthTray*" } |
    ForEach-Object { Stop-Process -Id $_.ProcessId -Force -ErrorAction SilentlyContinue }
Start-Sleep -Seconds 1
Write-OK "Done."

# ── Install files ─────────────────────────────────────────────────────────────
Write-Step "Installing to $InstallDir..."
if (-not (Test-Path $InstallDir)) {
    New-Item -ItemType Directory -Path $InstallDir -Force | Out-Null
}
Copy-Item -Path $SrcAgent -Destination $AgentPs1 -Force
Copy-Item -Path $SrcTray  -Destination $TrayPs1  -Force
Set-Content -Path (Join-Path $InstallDir "server_url.txt") -Value $ServerUrl -Encoding ASCII
Write-OK "Scripts installed."

# ── Agent scheduled task (SYSTEM, boot) ──────────────────────────────────────
Write-Step "Creating agent scheduled task..."
$psArgs    = "-NoProfile -NonInteractive -WindowStyle Hidden -ExecutionPolicy Bypass -File `"$AgentPs1`" -ServerUrl `"$ServerUrl`" -PollInterval $PollInterval"
$psArgsEsc = [System.Security.SecurityElement]::Escape($psArgs)
$agentXml  = @"
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
    <Command>powershell.exe</Command>
    <Arguments>$psArgsEsc</Arguments>
    <WorkingDirectory>$InstallDir</WorkingDirectory>
  </Exec></Actions>
</Task>
"@
$xmlPath = Join-Path $env:TEMP "dh_agent.xml"
[System.IO.File]::WriteAllText($xmlPath, $agentXml, [System.Text.Encoding]::Unicode)
$global:LASTEXITCODE = 0
$r = schtasks /Create /TN $TaskName /XML $xmlPath /F 2>&1
Remove-Item $xmlPath -Force -ErrorAction SilentlyContinue
if ($LASTEXITCODE -eq 0) {
    Write-OK "Agent task created."
} else {
    Write-Note "XML method failed ($r), trying legacy..."
    $global:LASTEXITCODE = 0
    schtasks /Create /TN $TaskName /SC ONSTART /DELAY "0000:15" /TR "powershell.exe $psArgs" /RU SYSTEM /RL HIGHEST /F 2>&1 | Out-Null
    if ($LASTEXITCODE -eq 0) { Write-OK "Agent task created (legacy)." }
    else { Write-Fail "Could not create agent task."; exit 1 }
}

# ── Tray scheduled task (interactive user logon) ──────────────────────────────
Write-Step "Creating tray scheduled task..."
$trayArgs    = "-NoProfile -STA -NonInteractive -WindowStyle Hidden -ExecutionPolicy Bypass -File `"$TrayPs1`""
$trayArgsEsc = [System.Security.SecurityElement]::Escape($trayArgs)
$iUser       = Get-InteractiveUser

$principal = if ($iUser) {
    "<UserId>$iUser</UserId><LogonType>InteractiveToken</LogonType><RunLevel>LeastPrivilege</RunLevel>"
} else {
    "<GroupId>S-1-5-4</GroupId><RunLevel>LeastPrivilege</RunLevel>"
}

$trayXml = @"
<?xml version="1.0" encoding="UTF-16"?>
<Task version="1.3" xmlns="http://schemas.microsoft.com/windows/2004/02/mit/task">
  <RegistrationInfo><Description>DiskHealth Tray</Description></RegistrationInfo>
  <Triggers><LogonTrigger><Enabled>true</Enabled></LogonTrigger></Triggers>
  <Principals><Principal id="Author">$principal</Principal></Principals>
  <Settings>
    <MultipleInstancesPolicy>IgnoreNew</MultipleInstancesPolicy>
    <DisallowStartIfOnBatteries>false</DisallowStartIfOnBatteries>
    <StopIfGoingOnBatteries>false</StopIfGoingOnBatteries>
    <ExecutionTimeLimit>PT0S</ExecutionTimeLimit>
    <Enabled>true</Enabled>
  </Settings>
  <Actions><Exec>
    <Command>powershell.exe</Command>
    <Arguments>$trayArgsEsc</Arguments>
    <WorkingDirectory>$InstallDir</WorkingDirectory>
  </Exec></Actions>
</Task>
"@
$trayXmlPath = Join-Path $env:TEMP "dh_tray.xml"
[System.IO.File]::WriteAllText($trayXmlPath, $trayXml, [System.Text.Encoding]::Unicode)
$global:LASTEXITCODE = 0
schtasks /Create /TN $TrayTask /XML $trayXmlPath /F 2>&1 | Out-Null
Remove-Item $trayXmlPath -Force -ErrorAction SilentlyContinue
if ($LASTEXITCODE -eq 0) { Write-OK "Tray task created." }
else { Write-Note "Tray task skipped (non-fatal)." }

# ── Start agent ───────────────────────────────────────────────────────────────
Write-Step "Starting agent..."
Start-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue
Write-OK "Agent started via scheduled task."

# ── Launch tray in user session via schtasks /Run ─────────────────────────────
Write-Step "Launching tray icon..."
$global:LASTEXITCODE = 0
schtasks /Run /TN $TrayTask 2>&1 | Out-Null
if ($LASTEXITCODE -eq 0) {
    Write-OK "Tray launched."
} else {
    Write-Note "Tray could not start now — will appear at next logon."
    # Fallback: HKCU Run key via reg add (works even from SYSTEM for current user hive)
    try {
        $regPath = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Run"
        Set-ItemProperty -Path $regPath -Name "DiskHealthTray" -Value "powershell.exe $trayArgs" -ErrorAction SilentlyContinue
        Write-Note "Tray registered as startup item (HKLM Run)."
    } catch { }
}

# ── Wait for agent.log ────────────────────────────────────────────────────────
Write-Host ""
Write-Step "Waiting for agent.log (up to 30s)..."
$waited = 0
$logFound = $false
while ($waited -lt 30) {
    Start-Sleep -Seconds 2
    $waited += 2
    if ((Test-Path $LogFile) -and (Get-Item $LogFile -ErrorAction SilentlyContinue).Length -gt 0) {
        $logFound = $true
        Write-OK "agent.log found — agent is running!"
        Write-Host ""
        Get-Content $LogFile -ErrorAction SilentlyContinue | Select-Object -Last 8 | ForEach-Object { Write-Note $_ }
        break
    }
    Write-Note "  ${waited}s..."
}

if (-not $logFound) {
    Write-Note "No log yet — agent will start on next boot, or run manually:"
    Write-Note "  Start-ScheduledTask -TaskName DiskHealthAgent"
}

Write-Host ""
Write-Host "============================================================" -ForegroundColor Green
Write-Host "  Installation complete!" -ForegroundColor Green
Write-Host "  Server : $ServerUrl" -ForegroundColor Green
Write-Host "  Poll   : every $PollInterval seconds" -ForegroundColor Green
Write-Host "  Log    : $LogFile" -ForegroundColor Green
Write-Host "============================================================" -ForegroundColor Green
Write-Host ""
exit 0