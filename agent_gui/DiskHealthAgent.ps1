# ==============================================================================
#  DiskHealth Agent v3.2.0
#  Compatible: Windows 7 SP1+ (PS 2.0+), Windows 10/11
#  Run via scheduled task as SYSTEM with -NoProfile -NonInteractive
# ==============================================================================
param(
    [Parameter(Mandatory=$true)]
    [string]$ServerUrl,
    [int]   $PollInterval = 600,
    [string]$AgentVersion = "3.2.0",
    [string]$Title        = "DiskHealth Agent",
    [switch]$ScanOnce,
    [string]$ScanCmdId   = ""
)

Set-StrictMode -Off
$ErrorActionPreference = "Continue"

# ==============================================================================
#  PATHS
# ==============================================================================
$script:AgentDir = if ($PSScriptRoot -and (Test-Path $PSScriptRoot)) {
    $PSScriptRoot
} elseif ($MyInvocation.MyCommand.Path) {
    Split-Path -Parent $MyInvocation.MyCommand.Path
} else {
    "$env:ProgramFiles\DiskHealthAgent"
}

$script:IdFile      = Join-Path $AgentDir "agent_id.txt"
$script:LogFile     = Join-Path $AgentDir "agent.log"
$script:ServerFile  = Join-Path $AgentDir "server_url.txt"
$script:LockFile    = Join-Path $AgentDir "scan.lock"
$script:NotifyFile  = Join-Path $AgentDir "update_notify.txt"

# ==============================================================================
#  LOGGING  — write startup marker immediately so log exists even if we crash
# ==============================================================================
try {
    [System.IO.Directory]::CreateDirectory($AgentDir) | Out-Null
    $banner = "[" + (Get-Date).ToString("yyyy-MM-dd HH:mm:ss") + "] [INFO ] DiskHealth Agent v$AgentVersion starting. AgentDir=$AgentDir"
    [System.IO.File]::AppendAllText($LogFile, $banner + [Environment]::NewLine, [System.Text.Encoding]::UTF8)
} catch { }

function Write-Log {
    param([string]$Level, [string]$Message)
    $line = "[" + (Get-Date).ToString("yyyy-MM-dd HH:mm:ss") + "] [$Level] $Message"
    Write-Host $line
    try { [System.IO.File]::AppendAllText($LogFile, $line + [Environment]::NewLine, [System.Text.Encoding]::UTF8) } catch { }
}
function Log-Info  { param([string]$m); Write-Log "INFO " $m }
function Log-Warn  { param([string]$m); Write-Log "WARN " $m }
function Log-Error { param([string]$m); Write-Log "ERROR" $m }

# ==============================================================================
#  AGENT ID
# ==============================================================================
function Get-AgentId {
    if (Test-Path $IdFile) {
        $id = (Get-Content $IdFile -Raw -ErrorAction SilentlyContinue).Trim()
        if ($id -match '^[0-9a-f\-]{36}$') { return $id }
    }

    $hostname    = $env:COMPUTERNAME.ToLower().Trim()
    $machineGuid = ""
    foreach ($rp in @("HKLM:\SOFTWARE\Microsoft\Cryptography","HKLM:\SOFTWARE\Wow6432Node\Microsoft\Cryptography")) {
        try {
            $g = (Get-ItemProperty -Path $rp -Name MachineGuid -ErrorAction Stop).MachineGuid
            if ($g -and $g.Length -gt 10) { $machineGuid = $g.ToLower().Trim(); break }
        } catch { }
    }

    $id = $null
    if ($machineGuid) {
        try {
            $raw  = [System.Text.Encoding]::UTF8.GetBytes("$hostname|$machineGuid")
            $sha1 = [System.Security.Cryptography.SHA1]::Create()
            $hash = $sha1.ComputeHash($raw); $sha1.Dispose()
            $hash[6] = ($hash[6] -band 0x0F) -bor 0x50
            $hash[8] = ($hash[8] -band 0x3F) -bor 0x80
            $hex = [BitConverter]::ToString($hash[0..15]) -replace '-', ''
            $id  = ("{0}-{1}-{2}-{3}-{4}" -f $hex.Substring(0,8), $hex.Substring(8,4), $hex.Substring(12,4), $hex.Substring(16,4), $hex.Substring(20,12)).ToLower()
        } catch { }
    }

    if (-not $id) {
        $bytes = New-Object byte[] 16
        $rng   = New-Object System.Security.Cryptography.RNGCryptoServiceProvider
        $rng.GetBytes($bytes)
        $bytes[6] = ($bytes[6] -band 0x0F) -bor 0x40
        $bytes[8] = ($bytes[8] -band 0x3F) -bor 0x80
        $hex = [BitConverter]::ToString($bytes) -replace '-', ''
        $id  = ("{0}-{1}-{2}-{3}-{4}" -f $hex.Substring(0,8), $hex.Substring(8,4), $hex.Substring(12,4), $hex.Substring(16,4), $hex.Substring(20,12)).ToLower()
    }

    try { Set-Content -Path $IdFile -Value $id -Encoding ASCII } catch { }
    Log-Info "Generated new agent_id: $id (hostname=$hostname)"
    return $id
}

# ==============================================================================
#  HTTP HELPERS
# ==============================================================================
function Invoke-JsonPost {
    param([string]$Url, [string]$JsonBody)
    try {
        $req              = [System.Net.WebRequest]::Create($Url)
        $req.Method       = "POST"
        $req.ContentType  = "application/json"
        $req.Timeout      = 15000
        $bytes            = [System.Text.Encoding]::UTF8.GetBytes($JsonBody)
        $req.ContentLength = $bytes.Length
        $stream = $req.GetRequestStream()
        $stream.Write($bytes, 0, $bytes.Length)
        $stream.Close()
        $resp   = $req.GetResponse()
        $reader = New-Object System.IO.StreamReader($resp.GetResponseStream())
        $body   = $reader.ReadToEnd()
        $reader.Close(); $resp.Close()
        return $body
    } catch {
        Log-Error "POST $Url failed: $_"
        return $null
    }
}

function Invoke-JsonGet {
    param([string]$Url)
    try {
        $req         = [System.Net.WebRequest]::Create($Url)
        $req.Method  = "GET"
        $req.Timeout = 15000
        $resp   = $req.GetResponse()
        $reader = New-Object System.IO.StreamReader($resp.GetResponseStream())
        $body   = $reader.ReadToEnd()
        $reader.Close(); $resp.Close()
        return $body
    } catch {
        Log-Error "GET $Url failed: $_"
        return $null
    }
}

# ==============================================================================
#  JSON SERIALISER  (PS 2.0 compatible, no ConvertTo-Json dependency)
# ==============================================================================
function ConvertTo-SafeJson {
    param($obj)
    if ($null -eq $obj)   { return "null" }
    if ($obj -is [bool])  { return if ($obj) { "true" } else { "false" } }
    if ($obj -is [int] -or $obj -is [long] -or $obj -is [double] -or $obj -is [float] -or $obj -is [decimal]) {
        $d = [double]$obj
        if ([double]::IsNaN($d) -or [double]::IsInfinity($d)) { return "null" }
        return "$obj"
    }
    if ($obj -is [string]) {
        $s = $obj -replace '\\','\\'  -replace '"','\"' `
                  -replace "`r",'\r'  -replace "`n",'\n' -replace "`t",'\t'
        return "`"$s`""
    }
    if ($obj -is [hashtable] -or $obj -is [System.Collections.Specialized.OrderedDictionary]) {
        $pairs = @()
        foreach ($k in $obj.Keys) {
            $ks = ([string]$k) -replace '\\','\\' -replace '"','\"'
            $pairs += "`"$ks`":" + (ConvertTo-SafeJson $obj[$k])
        }
        return "{" + ($pairs -join ",") + "}"
    }
    if ($obj -is [System.Collections.IEnumerable]) {
        $items = @()
        foreach ($item in $obj) { $items += ConvertTo-SafeJson $item }
        return "[" + ($items -join ",") + "]"
    }
    $s = ([string]$obj) -replace '\\','\\' -replace '"','\"'
    return "`"$s`""
}

# ==============================================================================
#  SYSTEM INFO
# ==============================================================================
function Get-LocalIP {
    try {
        $addrs = [System.Net.Dns]::GetHostAddresses([System.Net.Dns]::GetHostName())
        foreach ($a in $addrs) {
            if ($a.AddressFamily -eq [System.Net.Sockets.AddressFamily]::InterNetwork -and $a.ToString() -ne "127.0.0.1") {
                return $a.ToString()
            }
        }
    } catch { }
    return "127.0.0.1"
}

function Get-MacAddress {
    try {
        $mac = Get-WmiObject Win32_NetworkAdapterConfiguration -ErrorAction SilentlyContinue |
               Where-Object { $_.MACAddress -and $_.IPEnabled } |
               Select-Object -First 1 -ExpandProperty MACAddress
        if ($mac) { return $mac }
    } catch { }
    return ""
}

function Get-LoggedInUsers {
    $users = @()
    try {
        foreach ($s in (Get-WmiObject Win32_LoggedOnUser -ErrorAction Stop)) {
            if ($s.Antecedent -match 'Domain="([^"]+)",Name="([^"]+)"') {
                $domain = $matches[1]; $name = $matches[2]
                if ($name -notmatch '^(SYSTEM|LOCAL SERVICE|NETWORK SERVICE|DWM-\d+|UMFD-\d+)$') {
                    $entry = "$domain\$name"
                    if ($users -notcontains $entry) { $users += $entry }
                }
            }
        }
    } catch { }
    if ($users.Count -eq 0 -and $env:USERNAME) { $users += "$env:USERDOMAIN\$env:USERNAME" }
    return ($users -join ", ")
}

# ==============================================================================
#  SMART DATA  — smartctl (preferred) then WMI fallback
# ==============================================================================
function Find-Smartctl {
    $paths = @(
        "$env:ProgramFiles\smartmontools\bin\smartctl.exe",
        "C:\Program Files\smartmontools\bin\smartctl.exe",
        "C:\Program Files (x86)\smartmontools\bin\smartctl.exe"
    )
    foreach ($p in $paths) { if (Test-Path $p) { return $p } }
    try { return (Get-Command smartctl.exe -ErrorAction Stop).Source } catch { }
    return $null
}

function Normalize-Serial { param([string]$s); return ($s -replace '[\s\-_]', '').ToUpper().Trim() }

function Get-SmartctlData {
    param([string]$Bin)
    $bySerial = @{}
    try {
        $scanLines = & $Bin --scan 2>$null
        $devices   = @()
        foreach ($line in $scanLines) {
            $line = $line.Trim(); if (-not $line) { continue }
            if    ($line -match '^(/dev/\S+)\s+-d\s+(\S+)') { $devices += @{path=$matches[1]; dtype=$matches[2]} }
            elseif ($line -match '^(/dev/\S+)')              { $devices += @{path=$matches[1]; dtype='auto'} }
        }
        foreach ($dev in $devices) {
            try {
                $argList = @('-a', '-j')
                if ($dev.dtype -ne 'auto') { $argList += @('-d', $dev.dtype) }
                $argList  += $dev.path
                $jsonRaw   = & $Bin $argList 2>$null | Out-String
                if (-not $jsonRaw -or $jsonRaw.Trim().Length -lt 20) { continue }
                $d      = $jsonRaw | ConvertFrom-Json
                if (($d.smartctl.exit_status -band 1) -and -not $d.device) { continue }
                $serial = if ($d.serial_number) { $d.serial_number.Trim() } else { '' }
                if (-not $serial) { continue }

                $e = @{
                    serial=$serial; dtype=$dev.dtype; path=$dev.path
                    model=if($d.model_name){$d.model_name.Trim()}else{$null}
                    interface=if($d.device -and $d.device.protocol){$d.device.protocol}else{$dev.dtype.ToUpper()}
                    smart_passed=$true; predict_failure=$false
                    temperature=$null; reallocated=$null; pending=$null; uncorrectable=$null
                    power_on_hours=$null; power_cycles=$null
                    host_reads_gb=$null; host_writes_gb=$null
                    available_spare=$null; percentage_used=$null
                    unsafe_shutdowns=$null; media_errors=$null; critical_warning=$null
                    smartctl_used=$true
                }

                if ($d.smart_status -and $null -ne $d.smart_status.passed) {
                    $e.smart_passed    = [bool]$d.smart_status.passed
                    $e.predict_failure = -not [bool]$d.smart_status.passed
                }
                if ($d.temperature -and $null -ne $d.temperature.current) {
                    $e.temperature = [int]$d.temperature.current
                }

                # NVMe attributes
                $nv = $d.nvme_smart_health_information_log
                if ($nv) {
                    if ($null -ne $nv.power_on_hours)   { $e.power_on_hours   = [long]$nv.power_on_hours }
                    if ($null -ne $nv.power_cycles)     { $e.power_cycles     = [long]$nv.power_cycles }
                    if ($null -ne $nv.unsafe_shutdowns) { $e.unsafe_shutdowns = [long]$nv.unsafe_shutdowns }
                    if ($null -ne $nv.media_errors)     { $e.media_errors     = [long]$nv.media_errors }
                    if ($null -ne $nv.available_spare)  { $e.available_spare  = [int]$nv.available_spare }
                    if ($null -ne $nv.percentage_used)  { $e.percentage_used  = [int]$nv.percentage_used }
                    if ($null -ne $nv.critical_warning) { $e.critical_warning = $nv.critical_warning }
                    if ($null -ne $nv.host_reads)       { $e.host_reads_gb    = [math]::Round([double]$nv.host_reads  * 512000 / 1GB, 2) }
                    if ($null -ne $nv.host_writes)      { $e.host_writes_gb   = [math]::Round([double]$nv.host_writes * 512000 / 1GB, 2) }
                    $cw = $e.critical_warning
                    if ($cw -and $cw -ne 0 -and $cw -ne '0x00')                 { $e.predict_failure = $true }
                    if ($e.available_spare -ne $null -and $e.available_spare -le 10) { $e.predict_failure = $true }
                }

                # ATA attributes
                if ($d.ata_smart_attributes -and $d.ata_smart_attributes.table) {
                    foreach ($attr in $d.ata_smart_attributes.table) {
                        $aid = [int]$attr.id
                        $rv  = if ($attr.raw -and $null -ne $attr.raw.value) { [long]$attr.raw.value } else { 0 }
                        switch ($aid) {
                            5   { $e.reallocated    = [int]$rv }
                            9   { $e.power_on_hours = $rv }
                            12  { $e.power_cycles   = $rv }
                            187 { $e.uncorrectable  = [int]$rv }
                            190 { if ($null -eq $e.temperature) { $e.temperature = [int]($rv -band 0xFF) } }
                            194 { if ($null -eq $e.temperature) { $e.temperature = [int]($rv -band 0xFF) } }
                            197 { $e.pending        = [int]$rv }
                            198 { if ($null -eq $e.uncorrectable) { $e.uncorrectable = [int]$rv } }
                            241 { $e.host_writes_gb = [math]::Round($rv * 512 / 1GB, 2) }
                            242 { $e.host_reads_gb  = [math]::Round($rv * 512 / 1GB, 2) }
                        }
                    }
                    if ($e.reallocated -ne $null -and $e.reallocated -gt 0) { $e.predict_failure = $true }
                }

                if ($null -eq $e.power_on_hours -and $d.power_on_time -and $null -ne $d.power_on_time.hours) {
                    $e.power_on_hours = [long]$d.power_on_time.hours
                }

                $bySerial[(Normalize-Serial $serial)] = $e
            } catch { Log-Warn "smartctl parse error $($dev.path): $_" }
        }
    } catch { Log-Warn "smartctl --scan failed: $_" }
    return $bySerial
}

function Get-WmiSmartData {
    $result = @{}
    try {
        $statuses  = Get-WmiObject -Namespace "root\wmi" -Class MSStorageDriver_FailurePredictStatus -ErrorAction Stop
        $rawData   = Get-WmiObject -Namespace "root\wmi" -Class MSStorageDriver_FailurePredictData   -ErrorAction Stop
        $rawLookup = @{}
        foreach ($r in $rawData) { $rawLookup[$r.InstanceName] = $r.VendorSpecific }
        foreach ($s in $statuses) {
            $inst = $s.InstanceName; $idx = 0
            if ($inst -match '(\d+)$') { $idx = [int]$matches[1] }
            $e = @{ predict_failure=[bool]$s.PredictFailure; temperature=$null; reallocated=$null; pending=$null; uncorrectable=$null; smartctl_used=$false }
            if ($rawLookup.ContainsKey($inst)) {
                $raw = $rawLookup[$inst]
                for ($i = 2; $i -lt ($raw.Count - 12); $i += 12) {
                    $id = $raw[$i]; if ($id -eq 0) { continue }
                    $rv = [long]$raw[$i+5] + ([long]$raw[$i+6] -shl 8) + ([long]$raw[$i+7] -shl 16) + ([long]$raw[$i+8] -shl 24)
                    switch ($id) {
                        0xC2 { $e.temperature  = [int]($rv -band 0xFF) }
                        0x05 { $e.reallocated  = [int]$rv }
                        0xC5 { $e.pending      = [int]$rv }
                        0xC6 { $e.uncorrectable= [int]$rv }
                    }
                }
            }
            $result[$idx] = $e
        }
    } catch { $result['_error'] = $_.ToString() }
    return $result
}

# ==============================================================================
#  DISK HEALTH
# ==============================================================================
function Get-DiskSmartStatus {
    param($SmartEntry)
    if (-not $SmartEntry) { return 'Unknown' }
    if ($SmartEntry.predict_failure) { return 'Critical' }
    $status = 'Healthy'
    if ($SmartEntry.reallocated   -ne $null -and $SmartEntry.reallocated   -gt 0)  { $status = 'Warning'  }
    if ($SmartEntry.media_errors  -ne $null -and $SmartEntry.media_errors  -gt 0)  { $status = 'Warning'  }
    if ($SmartEntry.percentage_used -ne $null -and $SmartEntry.percentage_used -ge 90) { $status = 'Warning' }
    if ($SmartEntry.available_spare -ne $null -and $SmartEntry.available_spare -le 10) { $status = 'Critical' }
    return $status
}

function Get-DiskVolumes {
    param([string]$DeviceId)
    $volumes = @()
    try {
        $escaped = $DeviceId -replace '\\', '\\\\'
        $parts   = Get-WmiObject -Query "ASSOCIATORS OF {Win32_DiskDrive.DeviceID='$escaped'} WHERE AssocClass=Win32_DiskDriveToDiskPartition"
        foreach ($part in $parts) {
            $logs = Get-WmiObject -Query "ASSOCIATORS OF {Win32_DiskPartition.DeviceID='$($part.DeviceID)'} WHERE AssocClass=Win32_LogicalDiskToPartition"
            foreach ($ld in $logs) {
                $totalGb = if ($ld.Size)      { [math]::Round([long]$ld.Size      / 1GB, 2) } else { $null }
                $freeGb  = if ($ld.FreeSpace) { [math]::Round([long]$ld.FreeSpace / 1GB, 2) } else { $null }
                $usedPct = if ($totalGb -and $totalGb -gt 0 -and $freeGb -ne $null) {
                    [math]::Round((($totalGb - $freeGb) / $totalGb) * 100, 1)
                } else { 0 }
                $volumes += @{
                    drive      = $ld.DeviceID
                    label      = if ($ld.VolumeName) { $ld.VolumeName } else { '' }
                    filesystem = if ($ld.FileSystem)  { $ld.FileSystem }  else { 'Unknown' }
                    total_gb   = $totalGb
                    free_gb    = $freeGb
                    used_pct   = $usedPct
                }
            }
        }
    } catch { }
    return $volumes
}

function Get-DiskHealth {
    $scBin = Find-Smartctl
    $scMap = @{}
    if ($scBin) {
        try { $scMap = Get-SmartctlData -Bin $scBin } catch { Log-Warn "smartctl failed: $_" }
    }

    $wmiMap = Get-WmiSmartData
    if ($wmiMap.ContainsKey('_error')) {
        $wmiErr = $wmiMap['_error']; $wmiMap.Remove('_error')
        if ($scMap.Count -eq 0) { Log-Warn "WMI SMART unavailable: $wmiErr" }
        else                    { Log-Info "WMI SMART skipped (smartctl active)." }
    }

    # Build positional index map from smartctl for NVMe serial-mismatch fallback
    $scByIndex = @{}
    if ($scBin -and $scMap.Count -gt 0) {
        $idx = 0
        foreach ($line in (& $scBin --scan 2>$null)) {
            $line = $line.Trim(); if (-not $line) { continue }
            if ($line -match '^(/dev/\S+)') {
                foreach ($kv in $scMap.GetEnumerator()) {
                    if ($kv.Value.path -eq $matches[1]) { $scByIndex[$idx] = $kv.Value; break }
                }
                $idx++
            }
        }
    }

    $disks = @()
    foreach ($pd in (Get-WmiObject Win32_DiskDrive | Sort-Object Index)) {
        $i          = [int]$pd.Index
        $wmiSerial  = if ($pd.SerialNumber) { $pd.SerialNumber.Trim() } else { '' }
        $wmiModel   = if ($pd.Model)        { $pd.Model.Trim() }        else { 'Unknown' }
        $wmiIface   = if ($pd.InterfaceType){ $pd.InterfaceType.Trim() } else { 'Unknown' }

        # Prefer smartctl by serial, then positional, then WMI SMART, then WMI status
        $sc = $null
        if ($wmiSerial) {
            $norm = Normalize-Serial $wmiSerial
            if ($scMap.ContainsKey($norm)) { $sc = $scMap[$norm] }
        }
        if (-not $sc -and $scByIndex.ContainsKey($i)) { $sc = $scByIndex[$i] }

        $smartStatus = 'Unknown'
        $props = @{
            temperature=$null; reallocated=$null; pending=$null; uncorrectable=$null
            power_on_hours=$null; power_cycles=$null
            host_reads_gb=$null; host_writes_gb=$null
            available_spare=$null; percentage_used=$null
            unsafe_shutdowns=$null; media_errors=$null; critical_warning=$null
            smartctl_used=$false
        }
        $model=$wmiModel; $serial=$wmiSerial; $iface=$wmiIface

        if ($sc) {
            $props.smartctl_used = $true
            if ($sc.model)     { $model  = $sc.model }
            if ($sc.serial)    { $serial = $sc.serial }
            if ($sc.interface) { $iface  = $sc.interface }
            foreach ($k in @('temperature','reallocated','pending','uncorrectable','power_on_hours','power_cycles',
                              'host_reads_gb','host_writes_gb','available_spare','percentage_used',
                              'unsafe_shutdowns','media_errors','critical_warning')) {
                $props[$k] = $sc[$k]
            }
            $smartStatus = Get-DiskSmartStatus $sc
        } elseif ($wmiMap.ContainsKey($i)) {
            $w = $wmiMap[$i]
            $props.temperature  = $w.temperature
            $props.reallocated  = $w.reallocated
            $props.pending      = $w.pending
            $props.uncorrectable= $w.uncorrectable
            $smartStatus = Get-DiskSmartStatus $w
        } else {
            $smartStatus = switch ($pd.Status) {
                'OK'      { 'Healthy' }
                'Degraded'{ 'Warning' }
                'Error'   { 'Critical' }
                default   { 'Unknown' }
            }
        }

        $disks += @{
            index=$i; model=$model; serial=$serial; interface=$iface
            size_gb      = if ($pd.Size) { [math]::Round([long]$pd.Size / 1GB, 1) } else { $null }
            smart_status = $smartStatus
        } + $props + @{ volumes = Get-DiskVolumes -DeviceId $pd.DeviceID }
    }
    return $disks
}

# ==============================================================================
#  ALERTS
# ==============================================================================
$script:PrevState = @{}

function Send-Alert {
    param([string]$AgentId, [string]$AlertType, [string]$Message, $Data = @{})
    $payload = @{
        agent_id   = $AgentId
        alert_type = $AlertType
        message    = $Message
        hostname   = $env:COMPUTERNAME
        data       = $Data
        timestamp  = (Get-Date).ToString("o")
    }
    $resp = Invoke-JsonPost "$ServerUrl/api/alert" (ConvertTo-SafeJson $payload)
    if ($resp) { Log-Info "Alert sent: [$AlertType] $Message" }
    else       { Log-Warn  "Alert send failed: $AlertType" }
}

function Check-Alerts {
    param([string]$AgentId, $Disks)
    if (-not $Disks) { return }
    foreach ($disk in $Disks) {
        $key = if ($disk.serial) { $disk.serial } else { "disk$($disk.index)" }

        $changed = { param($field)
            $cur  = $disk[$field]
            $prev = $script:PrevState["$key.$field"]
            return ($cur -ne $null -and $cur -gt 0 -and $prev -ne $cur)
        }

        if (& $changed 'reallocated') {
            Send-Alert $AgentId "smart_warning" "Disk $($disk.index) ($($disk.model)): $($disk.reallocated) reallocated sector(s)" $disk
        }
        if (& $changed 'pending') {
            Send-Alert $AgentId "smart_warning" "Disk $($disk.index) ($($disk.model)): $($disk.pending) pending sector(s)" $disk
        }
        if (& $changed 'media_errors') {
            Send-Alert $AgentId "smart_warning" "Disk $($disk.index) ($($disk.model)): $($disk.media_errors) media error(s)" $disk
        }
        if ($disk.smart_status -eq "Critical" -and $script:PrevState["$key.status"] -ne "Critical") {
            Send-Alert $AgentId "smart_critical" "Disk $($disk.index) ($($disk.model)) SMART status is CRITICAL" $disk
        }

        $script:PrevState["$key.reallocated"]  = $disk.reallocated
        $script:PrevState["$key.pending"]      = $disk.pending
        $script:PrevState["$key.media_errors"] = $disk.media_errors
        $script:PrevState["$key.status"]       = $disk.smart_status

        foreach ($vol in $disk.volumes) {
            $drive   = $vol.drive
            $pct     = $vol.used_pct
            $prevPct = $script:PrevState["$drive.pct"]
            if ($pct -ne $null -and $pct -ge 90 -and ($prevPct -eq $null -or $prevPct -lt 90)) {
                Send-Alert $AgentId "low_disk" "Drive $drive is ${pct}% full ($($vol.free_gb) GB free)" $vol
                Log-Warn "ALERT: Drive $drive is ${pct}% full"
            }
            $script:PrevState["$drive.pct"] = $pct
        }
    }
}

# ==============================================================================
#  SERVER COMMUNICATION
# ==============================================================================
function Register-Agent {
    param([string]$AgentId, [string]$MacAddress = "")
    $os      = Get-WmiObject Win32_OperatingSystem -ErrorAction SilentlyContinue
    $payload = @{
        agent_id      = $AgentId
        hostname      = $env:COMPUTERNAME
        ip            = Get-LocalIP
        os            = "Windows"
        os_version    = if ($os) { $os.Caption } else { "" }
        agent_version = $AgentVersion
        logged_users  = Get-LoggedInUsers
        welcome_title = $Title
        mac_address   = $MacAddress
    }
    $resp = Invoke-JsonPost "$ServerUrl/api/register" (ConvertTo-SafeJson $payload)
    if ($resp) { Log-Info "Registered OK."; return $true }
    Log-Warn "Registration failed."
    return $false
}

function Build-Report {
    param([string]$AgentId, [string]$CmdId = "", [bool]$IsStartup = $false)
    $os = Get-WmiObject Win32_OperatingSystem -ErrorAction SilentlyContinue
    $report = @{
        agent_id      = $AgentId
        command_id    = $CmdId
        hostname      = $env:COMPUTERNAME
        ip            = Get-LocalIP
        os            = "Windows"
        os_version    = if ($os) { $os.Caption } else { "" }
        agent_version = $AgentVersion
        logged_users  = Get-LoggedInUsers
        disks         = @(Get-DiskHealth | ForEach-Object { $_ })
        Title         = $Title
    }
    if ($IsStartup) { $report["startup"] = $true }
    return $report
}

function Send-Report {
    param([string]$AgentId, [string]$CmdId = "", [bool]$IsStartup = $false)
    Log-Info "Collecting disk health data..."
    $report = Build-Report -AgentId $AgentId -CmdId $CmdId -IsStartup $IsStartup
    $json   = ConvertTo-SafeJson $report
    if ($json.Length -lt 10) { Log-Error "JSON too short — skipping."; return }
    Log-Info "Sending report ($($json.Length) bytes)..."
    $resp = Invoke-JsonPost "$ServerUrl/api/report" $json
    if ($resp) {
        Log-Info "Report accepted."
        try { Check-Alerts -AgentId $AgentId -Disks $report["disks"] } catch { }
    } else {
        Log-Warn "Send failed."
    }
}

function Send-Ack {
    param([string]$CmdId, $Result)
    Invoke-JsonPost "$ServerUrl/api/ack" (ConvertTo-SafeJson @{ command_id=$CmdId; result=$Result }) | Out-Null
}

# ==============================================================================
#  SELF-UPDATE
# ==============================================================================
function Update-Self {
    param([string]$AckCmdId = "")
    $dest = Join-Path $AgentDir "DiskHealthAgent.ps1"
    $tmp  = "$dest.new"
    try {
        Log-Info "Downloading updated agent from $ServerUrl/agent/agent.ps1"
        $wc = New-Object System.Net.WebClient
        $wc.DownloadFile("$ServerUrl/agent/agent.ps1", $tmp)
        Move-Item -Force $tmp $dest

        # Update tray script too
        try {
            $trayDest = Join-Path $AgentDir "DiskHealthTray.ps1"
            $trayTmp  = "$trayDest.new"
            $wc.DownloadFile("$ServerUrl/agent/tray.ps1", $trayTmp)
            Move-Item -Force $trayTmp $trayDest
            Log-Info "Tray script updated."
            # Re-register tray logon task if missing
            if (-not (Get-ScheduledTask -TaskName "DiskHealthTray" -ErrorAction SilentlyContinue)) {
                $ta = New-ScheduledTaskAction -Execute "powershell.exe" -Argument "-NoProfile -STA -NonInteractive -ExecutionPolicy Bypass -File `"$trayDest`""
                $tr = New-ScheduledTaskTrigger -AtLogOn
                Register-ScheduledTask -TaskName "DiskHealthTray" -Action $ta -Trigger $tr -RunLevel Limited -Force | Out-Null
                Log-Info "Tray logon task registered."
            }
            # Restart tray for current logged-in user
            Get-Process -Name powershell, pwsh -ErrorAction SilentlyContinue |
                Where-Object { try { (Get-WmiObject Win32_Process -Filter "ProcessId=$($_.Id)" -EA SilentlyContinue).CommandLine -like "*DiskHealthTray*" } catch { $false } } |
                ForEach-Object { Stop-Process -Id $_.Id -Force -ErrorAction SilentlyContinue }
            Start-Sleep -Seconds 1
            Start-Process powershell.exe -ArgumentList "-NoProfile -STA -NonInteractive -WindowStyle Hidden -ExecutionPolicy Bypass -File `"$trayDest`"" -ErrorAction SilentlyContinue
        } catch { Log-Info "Tray update skipped: $_" }

        try { [System.IO.File]::WriteAllText($NotifyFile, "Agent updated successfully! Restarting...", [System.Text.Encoding]::UTF8) } catch { }
        Start-Process powershell.exe -ArgumentList "-NoProfile -NonInteractive -WindowStyle Hidden -ExecutionPolicy Bypass -File `"$dest`" -ServerUrl `"$ServerUrl`" -PollInterval $PollInterval"
        exit 0
    } catch {
        Log-Error "Update failed: $_"
        if (Test-Path $tmp) { Remove-Item $tmp -Force -ErrorAction SilentlyContinue }
        try { [System.IO.File]::WriteAllText($NotifyFile, "Agent update FAILED. Check agent log.", [System.Text.Encoding]::UTF8) } catch { }
        return $false
    }
}

# ==============================================================================
#  BACKGROUND SCAN  (spawns a child process so polling never blocks on disk I/O)
# ==============================================================================
function Start-BackgroundScan {
    param([string]$AgentId, [string]$CmdId = "", [bool]$IsStartup = $false)

    # Skip if a scan is already running (lock file younger than 5 min)
    if (Test-Path $LockFile) {
        try {
            $age = ((Get-Date) - (Get-Item $LockFile -ErrorAction Stop).LastWriteTime).TotalMinutes
            if ($age -lt 5) { Log-Info "Scan already in progress, skipping."; return }
        } catch { }
        Remove-Item $LockFile -Force -ErrorAction SilentlyContinue
    }

    $scriptPath = Join-Path $AgentDir "DiskHealthAgent.ps1"
    if (-not (Test-Path $scriptPath)) {
        Log-Warn "DiskHealthAgent.ps1 not found — running scan inline."
        Send-Report -AgentId $AgentId -CmdId $CmdId -IsStartup $IsStartup
        return
    }

    $args  = "-NoProfile -NonInteractive -WindowStyle Hidden -ExecutionPolicy Bypass -File `"$scriptPath`" -ServerUrl `"$ServerUrl`" -PollInterval $PollInterval -AgentVersion `"$AgentVersion`" -ScanOnce"
    if ($CmdId)     { $args += " -ScanCmdId `"$CmdId`"" }

    Log-Info "Launching background scan (CmdId=$CmdId IsStartup=$IsStartup)"
    try {
        Start-Process powershell.exe -ArgumentList $args -WindowStyle Hidden -ErrorAction Stop
    } catch {
        Log-Warn "Background scan failed to start: $_. Running inline."
        Send-Report -AgentId $AgentId -CmdId $CmdId -IsStartup $IsStartup
    }
}

# ==============================================================================
#  COMMAND POLLING
# ==============================================================================
function Poll-Commands {
    param([string]$AgentId)
    $resp = Invoke-JsonGet "$ServerUrl/api/commands/$AgentId"
    if (-not $resp) { return }
    $commands = $null
    try { $commands = ($resp | ConvertFrom-Json).commands } catch { return }
    if (-not $commands -or $commands.Count -eq 0) { return }

    foreach ($cmd in $commands) {
        $cmdId  = $cmd.command_id
        $action = $cmd.action
        switch ($action) {
            "get_disk_health" {
                Start-BackgroundScan -AgentId $AgentId -CmdId $cmdId
            }
            "ping" {
                Send-Ack $cmdId @{ pong=$true; timestamp=(Get-Date).ToString("o") }
                Log-Info "Ping ack'd."
            }
            "update_agent" {
                try { [System.IO.File]::WriteAllText($NotifyFile, "Agent update started by administrator.", [System.Text.Encoding]::UTF8) } catch { }
                Log-Info "update_agent received — acking then updating."
                Send-Ack $cmdId @{ updating=$true; message="Update started — agent restarting shortly" }
                Update-Self -AckCmdId $cmdId | Out-Null
                Log-Warn "Update-Self returned without exiting."
            }
            "clear_log" {
                try {
                    if (Test-Path $LogFile) {
                        Clear-Content -Path $LogFile -Force -ErrorAction Stop
                        Log-Info "Log cleared by remote command."
                        Send-Ack $cmdId @{ cleared=$true; message="Log cleared successfully" }
                    } else {
                        Send-Ack $cmdId @{ cleared=$false; message="Log file not found" }
                    }
                } catch {
                    Send-Ack $cmdId @{ cleared=$false; message="Failed: $_" }
                }
                Log-Info "clear_log ack'd."
            }
            default {
                Send-Ack $cmdId @{ error="unknown action: $action" }
            }
        }
    }
}

# ==============================================================================
#  SCAN-ONCE MODE  (invoked as child process by Start-BackgroundScan)
# ==============================================================================
if ($ScanOnce) {
    $agentId = Get-AgentId
    Log-Info "ScanOnce: collecting disk health. CmdId=$ScanCmdId"
    try {
        Set-Content -Path $LockFile -Value (Get-Date).ToString("o") -Encoding ASCII -ErrorAction SilentlyContinue
        Send-Report -AgentId $agentId -CmdId $ScanCmdId
        Log-Info "ScanOnce: complete."
    } finally {
        Remove-Item $LockFile -Force -ErrorAction SilentlyContinue
    }
    exit 0
}

# ==============================================================================
#  MAIN LOOP
# ==============================================================================
function Main {
    $ErrorActionPreference = "SilentlyContinue"

    $agentId = Get-AgentId
    $macAddr = Get-MacAddress

    Log-Info "DiskHealth Agent v$AgentVersion starting. ID=$agentId  Server=$ServerUrl"

    # Write server URL so tray can build the panel link without waiting for registration
    try { Set-Content -Path $ServerFile -Value $ServerUrl -Encoding ASCII } catch { }

    # Register — retry until server is reachable
    $registered = $false
    while (-not $registered) {
        $registered = Register-Agent -AgentId $agentId -MacAddress $macAddr
        if (-not $registered) { Log-Warn "Registration failed, retrying in 15s..."; Start-Sleep -Seconds 15 }
    }

    # Initial disk scan (non-blocking)
    Start-BackgroundScan -AgentId $agentId -IsStartup $true
    $lastScan     = Get-Date
    $lastRegister = Get-Date

    # CMD_INTERVAL: how often we poll for commands (fast — no disk I/O)
    # SCAN_INTERVAL: how often we do a full disk scan (clamped to ≥60s)
    $CMD_INTERVAL  = 5
    $SCAN_INTERVAL = [Math]::Max(60, $PollInterval)
    Log-Info "Main loop started. CMD_INTERVAL=${CMD_INTERVAL}s  SCAN_INTERVAL=${SCAN_INTERVAL}s"

    while ($true) {
        Start-Sleep -Seconds $CMD_INTERVAL
        try {
            Poll-Commands -AgentId $agentId

            if (((Get-Date) - $lastRegister).TotalSeconds -ge 60) {
                Register-Agent -AgentId $agentId -MacAddress $macAddr | Out-Null
                $lastRegister = Get-Date
            }

            if (((Get-Date) - $lastScan).TotalSeconds -ge $SCAN_INTERVAL) {
                Start-BackgroundScan -AgentId $agentId
                $lastScan = Get-Date
            }
        } catch {
            Log-Error "Poll error: $_"
        }
    }
}

Main