# ==============================================================================
#  DiskHealth Tray Icon v3.2.0
#  Runs at user logon (-STA required for WinForms)
#  Shows agent status, opens web panel, allows restart
# ==============================================================================
param([string]$InstallDir = "$env:ProgramFiles\DiskHealthAgent")

# ── Single-instance guard ─────────────────────────────────────────────────────
$mutex = New-Object System.Threading.Mutex($false, "Global\DiskHealthTrayIcon")
$owned = $false
try     { $owned = $mutex.WaitOne(0, $false) }
catch   [System.Threading.AbandonedMutexException] { $owned = $true }
if (-not $owned) { $mutex.Dispose(); exit 0 }

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# ── File paths ────────────────────────────────────────────────────────────────
$LogFile      = Join-Path $InstallDir "agent.log"
$ServerFile   = Join-Path $InstallDir "server_url.txt"
$AgentIdFile  = Join-Path $InstallDir "agent_id.txt"
$NotifyFile   = Join-Path $InstallDir "update_notify.txt"
$AgentTask    = "DiskHealthAgent"

# ── Panel URL helper ──────────────────────────────────────────────────────────
function Get-PanelUrl {
    try {
        if (-not (Test-Path $ServerFile)) { return $null }
        $base = (Get-Content $ServerFile -Raw -ErrorAction SilentlyContinue).Trim()
        if (-not $base -or $base -notmatch "^https?://") { return $null }
        if (Test-Path $AgentIdFile) {
            $id = (Get-Content $AgentIdFile -Raw -ErrorAction SilentlyContinue).Trim()
            if ($id -match "^[0-9a-f\-]{36}$") { return "$base/?agent_id=$id" }
        }
        return $base
    } catch { return $null }
}

# ── Disk-drive icon builder ───────────────────────────────────────────────────
function Make-Icon {
    param([string]$Color = "#22c55e")
    $bmp = New-Object System.Drawing.Bitmap(32, 32)
    $g   = [System.Drawing.Graphics]::FromImage($bmp)
    $g.SmoothingMode     = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
    $g.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
    $g.Clear([System.Drawing.Color]::Transparent)

    $col   = [System.Drawing.ColorTranslator]::FromHtml($Color)
    $light = [System.Drawing.Color]::FromArgb(255, [math]::Min(255,$col.R+80), [math]::Min(255,$col.G+80), [math]::Min(255,$col.B+80))
    $dark  = [System.Drawing.Color]::FromArgb(255, [math]::Max(0,$col.R-60),  [math]::Max(0,$col.G-60),  [math]::Max(0,$col.B-60))

    $bb = New-Object System.Drawing.SolidBrush($col)
    $hb = New-Object System.Drawing.SolidBrush($light)
    $sb = New-Object System.Drawing.SolidBrush($dark)
    $pn = New-Object System.Drawing.Pen([System.Drawing.Color]::FromArgb(220,0,0,0), 1.5)
    $sl = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::FromArgb(100,0,0,0))
    $gl = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::FromArgb(140,$col.R,$col.G,$col.B))
    $wh = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::White)

    $g.FillRectangle($bb, 2, 6, 28, 20)
    $g.FillRectangle($hb, 2, 6, 28, 5)
    $g.FillRectangle($sb, 2, 21, 28, 5)
    $g.DrawRectangle($pn, 2, 6, 27, 19)
    $g.FillRectangle($sl, 4, 14, 14, 4)
    $g.FillEllipse($gl, 20, 13, 8, 8)
    $g.FillEllipse($wh, 22, 15, 4, 4)

    $g.Dispose()
    foreach ($r in @($bb,$hb,$sb,$pn,$sl,$gl,$wh)) { $r.Dispose() }

    return [System.Drawing.Icon]::FromHandle($bmp.GetHicon())
}

# ── Read agent status from log file ──────────────────────────────────────────
function Get-AgentStatus {
    if (-not (Test-Path $LogFile)) {
        return @{ color="#22c55e"; tip="DiskHealth Agent`nStarting up..."; status="starting" }
    }
    try {
        $fs      = [System.IO.File]::Open($LogFile, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::ReadWrite)
        $reader  = New-Object System.IO.StreamReader($fs)
        $content = $reader.ReadToEnd()
        $reader.Close(); $fs.Close()

        $lines = ($content -split "`r?`n") | Where-Object { $_ } | Select-Object -Last 60
        $last  = $lines | Where-Object { $_ -match "\[INFO \]|\[WARN \]|\[ERROR\]" } | Select-Object -Last 1
        $ts    = if ($last -match "\[(\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2})\]") { $matches[1] } else { "Unknown" }

        if    ($last -match "Report accepted") {
            return @{ color="#22c55e"; tip="DiskHealth Agent`nLast report: $ts`nStatus: OK";      status="ok"      }
        }
        elseif ($last -match "ERROR") {
            return @{ color="#ef4444"; tip="DiskHealth Agent`nLast event: $ts`nStatus: Error";    status="error"   }
        }
        elseif ($last -match "WARN") {
            return @{ color="#f59e0b"; tip="DiskHealth Agent`nLast event: $ts`nStatus: Warning";  status="warning" }
        }
        else {
            return @{ color="#22c55e"; tip="DiskHealth Agent`nLast event: $ts`nStatus: Running";  status="ok"      }
        }
    } catch {
        return @{ color="#22c55e"; tip="DiskHealth Agent`nChecking status..."; status="starting" }
    }
}

# ── Is the agent process alive? ───────────────────────────────────────────────
# SYSTEM-process CommandLine is null from user context — never rely on it.
# Use log file recency (agent writes every ~60s) and task state.
function Is-AgentRunning {
    # Method 1: Log file freshness (primary)
    try {
        if (Test-Path $LogFile) {
            $age = ((Get-Date) - (Get-Item $LogFile -ErrorAction Stop).LastWriteTime).TotalMinutes
            if ($age -lt 4) { return $true }
        }
    } catch { }

    # Method 2: Scheduled task state
    try {
        $task = Get-ScheduledTask -TaskName $AgentTask -ErrorAction SilentlyContinue
        if ($task -and $task.State -eq "Running") { return $true }
    } catch { }

    # Method 3: Other powershell processes + log younger than 10 min
    try {
        $others = @(Get-Process -Name powershell, pwsh -ErrorAction SilentlyContinue | Where-Object { $_.Id -ne $PID })
        if ($others.Count -gt 0 -and (Test-Path $LogFile)) {
            $logAge = ((Get-Date) - (Get-Item $LogFile -ErrorAction Stop).LastWriteTime).TotalMinutes
            if ($logAge -lt 10) { return $true }
        }
    } catch { }

    return $false
}

# ── Tray text helper (max 63 chars) ──────────────────────────────────────────
function Get-TrayText {
    param([string]$Tip)
    $text = (($Tip -split "`n")[0..1] -join " | ")
    return $text.Substring(0, [Math]::Min(63, $text.Length))
}

# ── Open panel in browser ─────────────────────────────────────────────────────
function Open-Panel {
    $url = Get-PanelUrl
    if ($url) {
        Start-Process $url -ErrorAction SilentlyContinue
    } else {
        [System.Windows.Forms.MessageBox]::Show(
            "Server URL not configured yet.`nThe agent may still be starting up.",
            "DiskHealth Agent",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information
        ) | Out-Null
    }
}

# ── Update notification balloon ───────────────────────────────────────────────
function Check-UpdateNotify {
    if (-not (Test-Path $NotifyFile)) { return }
    try {
        $msg = [System.IO.File]::ReadAllText($NotifyFile).Trim()
        Remove-Item $NotifyFile -Force -ErrorAction SilentlyContinue
        if (-not $msg) { return }

        if ($msg -match "started|starting|Downloading") {
            $tray.Icon = Make-Icon -Color "#f59e0b"
            $tray.Text = "DiskHealth Agent | UPDATING..."
            $tray.ShowBalloonTip(7000, "DiskHealth — Updating", $msg, [System.Windows.Forms.ToolTipIcon]::Info)
            $script:UpdateIconResetAt = (Get-Date).AddSeconds(8)
        } elseif ($msg -match "success|completed|Restarting") {
            $tray.ShowBalloonTip(6000, "DiskHealth — Updated", $msg, [System.Windows.Forms.ToolTipIcon]::Info)
        } elseif ($msg -match "FAILED|failed|error") {
            $tray.ShowBalloonTip(8000, "DiskHealth — Update Failed", $msg, [System.Windows.Forms.ToolTipIcon]::Error)
        } else {
            $tray.ShowBalloonTip(5000, "DiskHealth — Update", $msg, [System.Windows.Forms.ToolTipIcon]::Info)
        }
    } catch { }
}

# ── State ─────────────────────────────────────────────────────────────────────
$script:StartTime         = Get-Date
$script:GracePeriodSecs   = 45
$script:LastStatus        = "ok"
$script:UpdateIconResetAt = [DateTime]::MinValue

# ── Build tray icon ───────────────────────────────────────────────────────────
$tray         = New-Object System.Windows.Forms.NotifyIcon
$tray.Icon    = Make-Icon -Color "#22c55e"
$tray.Text    = "DiskHealth Agent | Starting up..."
$tray.Visible = $true

# ── Left-click → open panel ───────────────────────────────────────────────────
$tray.Add_Click({
    if ([System.Windows.Forms.Control]::MouseButtons -eq [System.Windows.Forms.MouseButtons]::Left) {
        Open-Panel
    }
})

# ── Context menu ──────────────────────────────────────────────────────────────
$menu    = New-Object System.Windows.Forms.ContextMenuStrip

$miTitle = $menu.Items.Add("DiskHealth Agent")
$miTitle.Enabled = $false
$miTitle.Font    = New-Object System.Drawing.Font("Segoe UI", 8, [System.Drawing.FontStyle]::Bold)
$menu.Items.Add("-") | Out-Null

$miPanel = $menu.Items.Add("Open Web Panel")
$miPanel.Add_Click({ Open-Panel })

$miLog = $menu.Items.Add("View Agent Log")
$miLog.Add_Click({
    try {
        if (-not (Test-Path $LogFile)) {
            [System.Windows.Forms.MessageBox]::Show(
                "agent.log does not exist yet.`nThe agent may still be starting.",
                "DiskHealth Agent", [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Information) | Out-Null
            return
        }
        $tmp = Join-Path $env:TEMP "DiskHealth_log_view.txt"
        $fs  = [System.IO.File]::Open($LogFile, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::ReadWrite)
        $rd  = New-Object System.IO.StreamReader($fs)
        $c   = $rd.ReadToEnd(); $rd.Close(); $fs.Close()
        [System.IO.File]::WriteAllText($tmp, $c)
        Start-Process notepad.exe -ArgumentList $tmp -ErrorAction SilentlyContinue
    } catch {
        Start-Process notepad.exe -ArgumentList $LogFile -ErrorAction SilentlyContinue
    }
})

$miRestart = $menu.Items.Add("Restart Agent")
$miRestart.Add_Click({
    try {
        Stop-ScheduledTask  -TaskName $AgentTask -ErrorAction SilentlyContinue
        Start-Sleep -Seconds 2
        Start-ScheduledTask -TaskName $AgentTask -ErrorAction SilentlyContinue
        $tray.ShowBalloonTip(3000, "DiskHealth Agent", "Agent restarted.", [System.Windows.Forms.ToolTipIcon]::Info)
    } catch {
        $tray.ShowBalloonTip(3000, "DiskHealth Agent", "Restart failed.", [System.Windows.Forms.ToolTipIcon]::Error)
    }
})

$menu.Items.Add("-") | Out-Null

$miExit = $menu.Items.Add("Exit Tray Icon")
$miExit.Add_Click({
    $tray.Visible = $false
    $tray.Dispose()
    [System.Windows.Forms.Application]::Exit()
})

$tray.ContextMenuStrip = $menu

# ── Timer — update icon every 8s ─────────────────────────────────────────────
$timer          = New-Object System.Windows.Forms.Timer
$timer.Interval = 8000
$timer.Add_Tick({
    # ALL tick logic wrapped — one exception must never stop future ticks
    try {
        # 1. Handle update notify file
        Check-UpdateNotify

        # 2. Expire temporary update icon
        if ($script:UpdateIconResetAt -gt [DateTime]::MinValue -and (Get-Date) -ge $script:UpdateIconResetAt) {
            $script:UpdateIconResetAt = [DateTime]::MinValue
        }

        # 3. Grace period: stay green while agent/log is initialising
        $elapsed = ((Get-Date) - $script:StartTime).TotalSeconds
        if ($elapsed -lt $script:GracePeriodSecs) { return }

        # 4. Check if agent process is alive
        if (-not (Is-AgentRunning)) {
            if ($script:LastStatus -ne "stopped") {
                $tray.Icon = Make-Icon -Color "#ef4444"
                $tray.Text = "DiskHealth Agent | NOT RUNNING"
                $tray.ShowBalloonTip(6000, "DiskHealth Agent",
                    "Agent is not running! Right-click > Restart Agent.",
                    [System.Windows.Forms.ToolTipIcon]::Warning)
            }
            $script:LastStatus = "stopped"
            return
        }

        # 5. Agent is running — read status from log
        $s = Get-AgentStatus

        if ($s.status -eq "starting") {
            # Log exists but no meaningful line yet — show "waiting" instead of freezing
            if ($script:LastStatus -ne "waiting") {
                $tray.Icon = Make-Icon -Color "#6b7280"
                $tray.Text = "DiskHealth Agent | Waiting for log..."
            }
            $script:LastStatus = "waiting"
            return
        }

        # 6. Balloon on meaningful status transitions
        $prevStatus = $script:LastStatus
        if ($s.status -eq "error"   -and $prevStatus -ne "error") {
            $tray.ShowBalloonTip(8000, "DiskHealth Agent — ALERT",  $s.tip, [System.Windows.Forms.ToolTipIcon]::Error)
        } elseif ($s.status -eq "warning" -and $prevStatus -ne "warning") {
            $tray.ShowBalloonTip(6000, "DiskHealth Agent",          $s.tip, [System.Windows.Forms.ToolTipIcon]::Warning)
        } elseif ($s.status -eq "ok" -and $prevStatus -in @("stopped","waiting","error")) {
            $tray.ShowBalloonTip(4000, "DiskHealth Agent", "Agent is online.", [System.Windows.Forms.ToolTipIcon]::Info)
        }

        $script:LastStatus = $s.status
        $tray.Icon = Make-Icon -Color $s.color
        $tray.Text = Get-TrayText $s.tip

    } catch {
        # Uncomment to debug tick errors:
        # [System.IO.File]::AppendAllText("$env:TEMP\DiskHealthTray_err.txt", "$((Get-Date).ToString('HH:mm:ss')) TICK: $_`n")
    }
})
$timer.Start()

# ── Show startup balloon ───────────────────────────────────────────────────────
$tray.ShowBalloonTip(3000, "DiskHealth Agent", "Monitoring disk health. Click icon to open web panel.", [System.Windows.Forms.ToolTipIcon]::Info)

[System.Windows.Forms.Application]::Run()

# ── Cleanup ───────────────────────────────────────────────────────────────────
$timer.Dispose()
$tray.Dispose()
if ($owned) { $mutex.ReleaseMutex() }
$mutex.Dispose()