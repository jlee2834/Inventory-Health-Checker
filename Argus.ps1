param(
    [string[]]$ComputerName = @($env:COMPUTERNAME),
    [string]$InputFile,
    [string]$OutputDir = ".\output",
    [string]$LogoUrl = "",
    [switch]$ExportHtml = $true,
    [switch]$OpenReport = $true,
    [switch]$SkipHotfixes,
    [int]$ThrottleDelayMs = 200
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Ensure-OutputDir {
    param([string]$Path)
    if (-not (Test-Path -LiteralPath $Path)) {
        New-Item -Path $Path -ItemType Directory -Force | Out-Null
    }
}

function Resolve-ComputerList {
    param(
        [string[]]$Names,
        [string]$File
    )

    $targets = New-Object System.Collections.Generic.List[string]

    foreach ($name in $Names) {
        if (-not [string]::IsNullOrWhiteSpace($name)) {
            $targets.Add($name.Trim())
        }
    }

    if ($File) {
        if (-not (Test-Path -LiteralPath $File)) {
            throw "Input file not found: $File"
        }
        $lines = Get-Content -LiteralPath $File | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
        foreach ($line in $lines) {
            $targets.Add($line.Trim())
        }
    }

    return $targets | Sort-Object -Unique
}

function Test-TcpPort {
    param(
        [Parameter(Mandatory)][string]$Computer,
        [Parameter(Mandatory)][int]$Port,
        [int]$TimeoutMs = 1500
    )

    $client = New-Object System.Net.Sockets.TcpClient
    try {
        $async = $client.BeginConnect($Computer, $Port, $null, $null)
        $wait  = $async.AsyncWaitHandle.WaitOne($TimeoutMs, $false)
        if (-not $wait) { return $false }
        $client.EndConnect($async)
        return $true
    }
    catch { return $false }
    finally { $client.Close() }
}

function Get-LastBootReadable {
    param([datetime]$LastBoot)
    $uptime = (Get-Date) - $LastBoot
    return "{0}d {1}h {2}m" -f [math]::Floor($uptime.TotalDays), $uptime.Hours, $uptime.Minutes
}

function Get-HealthStatus {
    param(
        [double]$DiskFreePercent,
        [int]$PendingRebootSignals,
        [bool]$DefenderRealtime,
        [object]$LastHotfixDate,
        [bool]$HasError
    )

    $issues = New-Object System.Collections.Generic.List[string]

    if ($HasError)                   { $issues.Add('Collection error') }
    if ($DiskFreePercent -lt 10)     { $issues.Add('Critical disk space') }
    elseif ($DiskFreePercent -lt 20) { $issues.Add('Low disk space') }
    if ($PendingRebootSignals -gt 0) { $issues.Add('Pending reboot') }
    if (-not $DefenderRealtime)      { $issues.Add('Defender real-time protection off or unavailable') }

    if ($null -ne $LastHotfixDate) {
        try {
            $daysSincePatch = ((Get-Date) - [datetime]$LastHotfixDate).Days
            if ($daysSincePatch -gt 45) { $issues.Add('Patch age over 45 days') }
        }
        catch { $issues.Add('Patch date unavailable') }
    }
    else {
        $issues.Add('Patch date unavailable')
    }

    if ($issues.Count -eq 0) {
        return [pscustomobject]@{ Status = 'Healthy'; Issues = 'None' }
    }

    $status = 'Warning'
    if ($issues -contains 'Collection error' -or
        $issues -contains 'Critical disk space' -or
        $issues -contains 'Defender real-time protection off or unavailable') {
        $status = 'Needs Attention'
    }

    return [pscustomobject]@{
        Status = $status
        Issues = ($issues -join '; ')
    }
}

function Get-PendingRebootSignals {
    param([string]$Computer)

    $signals = 0
    $paths = @(
        'SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\RebootPending',
        'SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired'
    )

    try {
        $base = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey('LocalMachine', $Computer)
        foreach ($path in $paths) {
            $sub = $base.OpenSubKey($path)
            if ($sub) { $signals++; $sub.Close() }
        }
        $sm = $base.OpenSubKey('SYSTEM\CurrentControlSet\Control\Session Manager')
        if ($sm) {
            if ($sm.GetValue('PendingFileRenameOperations', $null)) { $signals++ }
            $sm.Close()
        }
        $base.Close()
    }
    catch {}

    return $signals
}

function Get-InstalledSoftwareCount {
    param([string]$Computer)

    $count = 0
    $paths = @(
        'SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall',
        'SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall'
    )

    try {
        $base = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey('LocalMachine', $Computer)
        foreach ($path in $paths) {
            $sub = $base.OpenSubKey($path)
            if ($sub) {
                $count += ($sub.GetSubKeyNames() | Measure-Object).Count
                $sub.Close()
            }
        }
        $base.Close()
    }
    catch { return $null }

    return $count
}

function Get-RemoteInventory {
    param(
        [Parameter(Mandatory)][string]$Computer,
        [switch]$SkipHotfixes
    )

    $session = $null
    $result  = [ordered]@{
        ComputerName           = $Computer
        Reachable              = $false
        CimSession             = $false
        OS                     = $null
        Version                = $null
        BuildNumber            = $null
        Architecture           = $null
        LastBoot               = $null
        Uptime                 = $null
        LoggedOnUser           = $null
        SerialNumber           = $null
        Manufacturer           = $null
        Model                  = $null
        CPU                    = $null
        RAMGB                  = $null
        DiskC_TotalGB          = $null
        DiskC_FreeGB           = $null
        DiskC_FreePercent      = $null
        IPv4                   = $null
        Antivirus              = $null
        DefenderRealtime       = $false
        LastHotfixId           = $null
        LastHotfixDate         = $null
        PendingRebootSignals   = 0
        InstalledSoftwareCount = $null
        RdpOpen                = $false
        WinRMOpen              = $false
        HealthStatus           = 'Unknown'
        HealthIssues           = 'Unknown'
        Error                  = $null
    }

    try {
        $result.Reachable = Test-Connection -ComputerName $Computer -Count 1 -Quiet -ErrorAction SilentlyContinue
        $result.RdpOpen   = Test-TcpPort -Computer $Computer -Port 3389
        $result.WinRMOpen = Test-TcpPort -Computer $Computer -Port 5985

        $session = New-CimSession -ComputerName $Computer -ErrorAction Stop
        $result.CimSession = $true

        $os   = Get-CimInstance -CimSession $session -ClassName Win32_OperatingSystem
        $cs   = Get-CimInstance -CimSession $session -ClassName Win32_ComputerSystem
        $bios = Get-CimInstance -CimSession $session -ClassName Win32_BIOS
        $cpu  = Get-CimInstance -CimSession $session -ClassName Win32_Processor | Select-Object -First 1
        $disk = Get-CimInstance -CimSession $session -ClassName Win32_LogicalDisk -Filter "DeviceID='C:'"
        $nic  = Get-CimInstance -CimSession $session -ClassName Win32_NetworkAdapterConfiguration -Filter 'IPEnabled = True'

        $result.OS           = $os.Caption
        $result.Version      = $os.Version
        $result.BuildNumber  = $os.BuildNumber
        $result.Architecture = $os.OSArchitecture
        $result.LastBoot     = $os.LastBootUpTime
        $result.Uptime       = Get-LastBootReadable -LastBoot $os.LastBootUpTime
        $result.LoggedOnUser = $cs.UserName
        $result.SerialNumber = $bios.SerialNumber
        $result.Manufacturer = $cs.Manufacturer
        $result.Model        = $cs.Model
        $result.CPU          = $cpu.Name
        $result.RAMGB        = [math]::Round($cs.TotalPhysicalMemory / 1GB, 2)

        if ($disk) {
            $result.DiskC_TotalGB     = [math]::Round($disk.Size / 1GB, 2)
            $result.DiskC_FreeGB      = [math]::Round($disk.FreeSpace / 1GB, 2)
            $result.DiskC_FreePercent = if ($disk.Size -gt 0) {
                [math]::Round(($disk.FreeSpace / $disk.Size) * 100, 2)
            } else { 0 }
        }

        $ipv4List = @()
        foreach ($adapter in $nic) {
            foreach ($ip in $adapter.IPAddress) {
                if ($ip -match '^\d+\.\d+\.\d+\.\d+$') { $ipv4List += $ip }
            }
        }
        $result.IPv4 = ($ipv4List | Sort-Object -Unique) -join ', '

        # Guard Defender query behind WinRM availability check
        if ($result.WinRMOpen) {
            try {
                $defender = Invoke-Command -ComputerName $Computer -ScriptBlock {
                    if (Get-Command Get-MpComputerStatus -ErrorAction SilentlyContinue) {
                        Get-MpComputerStatus | Select-Object AMProductVersion, AntivirusEnabled, RealTimeProtectionEnabled
                    }
                } -ErrorAction Stop

                if ($defender) {
                    $result.Antivirus        = "Microsoft Defender $($defender.AMProductVersion)"
                    $result.DefenderRealtime = [bool]$defender.RealTimeProtectionEnabled
                }
                else {
                    $result.Antivirus = 'Unknown or unavailable'
                }
            }
            catch {
                $result.Antivirus = 'Unknown or unavailable'
            }
        }
        else {
            $result.Antivirus = 'WinRM unavailable'
        }

        if (-not $SkipHotfixes) {
            try {
                $hotfix = Get-HotFix -ComputerName $Computer -ErrorAction Stop |
                    Sort-Object InstalledOn -Descending |
                    Select-Object -First 1

                if ($hotfix) {
                    $result.LastHotfixId   = $hotfix.HotFixID
                    $result.LastHotfixDate = $hotfix.InstalledOn
                }
            }
            catch {
                $result.LastHotfixId = 'Unavailable'
            }
        }

        $result.PendingRebootSignals   = Get-PendingRebootSignals -Computer $Computer
        $result.InstalledSoftwareCount = Get-InstalledSoftwareCount -Computer $Computer
    }
    catch {
        $result.Error = $_.Exception.Message
    }
    finally {
        if ($session) {
            Remove-CimSession -CimSession $session -ErrorAction SilentlyContinue
        }
    }

    # PS 5.1-compatible null coalescing
    $diskFreePercent = if ($null -ne $result.DiskC_FreePercent) { [double]$result.DiskC_FreePercent } else { 0.0 }

    $health = Get-HealthStatus `
        -DiskFreePercent      $diskFreePercent `
        -PendingRebootSignals ([int]$result.PendingRebootSignals) `
        -DefenderRealtime     ([bool]$result.DefenderRealtime) `
        -LastHotfixDate       $result.LastHotfixDate `
        -HasError             ([bool](-not [string]::IsNullOrWhiteSpace($result.Error)))

    $result.HealthStatus = $health.Status
    $result.HealthIssues = $health.Issues

    return [pscustomobject]$result
}

function ConvertTo-PrettyHtmlReport {
    param(
        [Parameter(Mandatory)][object[]]$Data,
        [Parameter(Mandatory)][string]$Title,
        [Parameter(Mandatory)][string]$CsvFileName,
        [Parameter(Mandatory)][string]$JsonFileName,
        [Parameter(Mandatory)][string]$ExcelFileName,
        [string]$LogoUrl = ""
    )

    $json        = @($Data) | ConvertTo-Json -Depth 6 -Compress
    $jsonEscaped = $json.Replace('</script>', '<\/script>')
    $generated   = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
    $total       = @($Data).Count
    $healthy     = @($Data | Where-Object { $_.HealthStatus -eq 'Healthy' }).Count
    $warning     = @($Data | Where-Object { $_.HealthStatus -eq 'Warning' }).Count
    $attention   = @($Data | Where-Object { $_.HealthStatus -eq 'Needs Attention' }).Count

    # -------------------------------------------------------------------------
    # Single-quoted here-string: PowerShell will NOT expand any $ tokens.
    # All JavaScript template literals (${...} and backticks) are safe here.
    # Dynamic values are injected below via -replace with __PLACEHOLDER__ tokens.
    # -------------------------------------------------------------------------
    $html = @'
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>__TITLE__</title>
<style>
    :root {
        --bg: #0b1220;
        --panel: #121a2b;
        --panel-2: #182235;
        --text: #e8eefc;
        --muted: #9fb0d1;
        --line: #293653;
        --green: #22c55e;
        --yellow: #f59e0b;
        --red: #ef4444;
        --blue: #60a5fa;
    }
    * { box-sizing: border-box; }
    body {
        margin: 0;
        font-family: Inter, Segoe UI, Arial, sans-serif;
        background: linear-gradient(180deg, #08101d 0%, #0b1220 100%);
        color: var(--text);
    }
    .wrap { max-width: 1600px; margin: 0 auto; padding: 24px; }
    .topbar {
        display: flex; justify-content: space-between;
        align-items: flex-start; gap: 16px;
        margin-bottom: 20px; flex-wrap: wrap;
    }
    .logo { max-width: 50px; max-height: 50px; margin-right: 12px; }
    .title { display: flex; align-items: flex-start; gap: 12px; }
    .title h1 { margin: 0 0 8px 0; font-size: 30px; }
    .title p  { margin: 0; color: var(--muted); }
    .actions  { display: flex; gap: 10px; flex-wrap: wrap; }
    button {
        background: var(--panel-2); color: var(--text);
        border: 1px solid var(--line); padding: 10px 14px;
        border-radius: 10px; cursor: pointer; font-weight: 600;
        transition: border-color .15s, transform .15s;
    }
    button:hover { border-color: var(--blue); transform: translateY(-1px); }
    .cards {
        display: grid;
        grid-template-columns: repeat(auto-fit, minmax(220px, 1fr));
        gap: 14px; margin-bottom: 18px;
    }
    .card {
        background: rgba(18,26,43,.92); border: 1px solid var(--line);
        border-radius: 16px; padding: 18px;
        box-shadow: 0 10px 30px rgba(0,0,0,.22);
    }
    .card .label { color: var(--muted); font-size: 13px; margin-bottom: 10px; }
    .card .value { font-size: 30px; font-weight: 700; }
    .controls {
        display: grid;
        grid-template-columns: 1.5fr 1fr 1fr 1fr;
        gap: 12px; margin-bottom: 8px;
    }
    .controls input, .controls select {
        width: 100%; background: var(--panel); color: var(--text);
        border: 1px solid var(--line); border-radius: 12px; padding: 12px;
    }
    .summary-bar { font-size: 13px; color: var(--muted); margin-bottom: 10px; }
    .table-wrap {
        background: rgba(18,26,43,.92); border: 1px solid var(--line);
        border-radius: 18px; overflow: hidden;
        box-shadow: 0 10px 30px rgba(0,0,0,.22);
    }
    .table-scroll { overflow: auto; max-height: 72vh; }
    table { width: 100%; border-collapse: collapse; min-width: 1500px; }
    thead th {
        position: sticky; top: 0; z-index: 1;
        background: #14203a; color: var(--text);
        text-align: left; padding: 12px;
        border-bottom: 1px solid var(--line);
        font-size: 13px; white-space: nowrap;
        cursor: pointer; user-select: none;
    }
    thead th:hover { background: #1a2d4a; }
    thead th.sort-asc::after  { content: ' \25B2'; font-size: 10px; }
    thead th.sort-desc::after { content: ' \25BC'; font-size: 10px; }
    tbody td {
        padding: 11px 12px;
        border-bottom: 1px solid rgba(41,54,83,.55);
        color: #dce6fb; font-size: 13px; vertical-align: top;
    }
    tbody tr:hover        { background: rgba(96,165,250,.06); }
    tbody tr.selected-row { background: rgba(96,165,250,.12); outline: 1px solid rgba(96,165,250,.35); }
    .clickable-row        { cursor: pointer; }
    .pill { display: inline-block; padding: 3px 10px; border-radius: 999px; font-size: 12px; font-weight: 600; }
    .pill.healthy   { background: rgba(34,197,94,.18);  color: var(--green); }
    .pill.warning   { background: rgba(245,158,11,.18); color: var(--yellow); }
    .pill.attention { background: rgba(239,68,68,.18);  color: var(--red); }
    .bool-true  { color: var(--green); font-weight: 600; }
    .bool-false { color: var(--red);   font-weight: 600; }
    .panel-overlay {
        position: fixed; inset: 0; background: rgba(0,0,0,.45);
        opacity: 0; pointer-events: none;
        transition: opacity .18s ease; z-index: 20;
    }
    .panel-overlay.open { opacity: 1; pointer-events: auto; }
    .side-panel {
        position: fixed; top: 0; right: 0;
        width: min(520px, 96vw); height: 100vh;
        background: #0f1728; border-left: 1px solid var(--line);
        box-shadow: -14px 0 34px rgba(0,0,0,.35);
        transform: translateX(100%); transition: transform .22s ease;
        z-index: 21; display: flex; flex-direction: column;
    }
    .side-panel.open { transform: translateX(0); }
    .panel-header {
        display: flex; justify-content: space-between;
        align-items: flex-start; gap: 12px;
        padding: 20px 20px 14px 20px;
        border-bottom: 1px solid var(--line); background: #121c31;
    }
    .panel-title h2 { margin: 0 0 8px 0; font-size: 24px; }
    .panel-title p  { margin: 0; color: var(--muted); font-size: 13px; }
    .panel-close    { min-width: 42px; padding: 10px 12px; }
    .panel-body     { overflow: auto; padding: 18px 20px 24px 20px; }
    .panel-actions  { display: flex; flex-wrap: wrap; gap: 8px; margin-bottom: 16px; }
    .detail-grid    { display: grid; grid-template-columns: 1fr 1fr; gap: 12px; }
    .detail-card {
        background: rgba(24,34,53,.88); border: 1px solid var(--line);
        border-radius: 14px; padding: 14px;
    }
    .detail-card.full { grid-column: 1 / -1; }
    .section-card {
        grid-column: 1 / -1; background: rgba(24,34,53,.72);
        border: 1px solid var(--line); border-radius: 16px; overflow: hidden;
    }
    .section-header {
        display: flex; align-items: center;
        justify-content: space-between; gap: 12px;
        padding: 14px 16px; border-bottom: 1px solid rgba(41,54,83,.75);
    }
    .section-title    { font-size: 15px; font-weight: 700; letter-spacing: .01em; }
    .section-subtitle { color: var(--muted); font-size: 12px; }
    .section-content  { display: grid; grid-template-columns: 1fr 1fr; gap: 12px; padding: 14px; }
    .section-card.health-overview  { border-color: rgba(96,165,250,.35); background: rgba(20,32,58,.9); }
    .section-card.health-healthy   { border-color: rgba(34,197,94,.38);  background: linear-gradient(180deg,rgba(10,40,24,.9),rgba(18,26,43,.92)); }
    .section-card.health-warning   { border-color: rgba(245,158,11,.4);  background: linear-gradient(180deg,rgba(56,35,5,.92),rgba(18,26,43,.92)); }
    .section-card.health-attention { border-color: rgba(239,68,68,.42);  background: linear-gradient(180deg,rgba(60,17,17,.94),rgba(18,26,43,.92)); }
    .section-card.section-overview     { border-left: 4px solid rgba(96,165,250,.85); }
    .section-card.section-hardware     { border-left: 4px solid rgba(168,85,247,.85); }
    .section-card.section-security     { border-left: 4px solid rgba(34,197,94,.85); }
    .section-card.section-patching     { border-left: 4px solid rgba(245,158,11,.9); }
    .section-card.section-connectivity { border-left: 4px solid rgba(59,130,246,.9); }
    .detail-label {
        color: var(--muted); font-size: 12px; margin-bottom: 8px;
        text-transform: uppercase; letter-spacing: .04em;
    }
    .detail-value { font-size: 15px; line-height: 1.45; word-break: break-word; }
    .kv { display: grid; grid-template-columns: 135px 1fr; gap: 8px 12px; }
    .kv .k { color: var(--muted); font-size: 13px; }
    .kv .v { font-size: 13px; word-break: break-word; }
    #toast {
        position: fixed; bottom: 28px; left: 50%;
        transform: translateX(-50%) translateY(60px);
        background: #1e2f4d; color: var(--text);
        border: 1px solid var(--blue); border-radius: 12px;
        padding: 12px 22px; font-size: 14px; font-weight: 600;
        opacity: 0; transition: opacity .2s, transform .2s;
        z-index: 50; pointer-events: none;
    }
    #toast.show { opacity: 1; transform: translateX(-50%) translateY(0); }
    .footer { margin-top: 12px; color: var(--muted); font-size: 12px; }
    @media (max-width: 1100px) { .controls { grid-template-columns: 1fr 1fr; } }
    @media (max-width: 720px) {
        .controls { grid-template-columns: 1fr; }
        .title h1 { font-size: 24px; }
        .detail-grid { grid-template-columns: 1fr; }
        .section-content { grid-template-columns: 1fr; }
        .kv { grid-template-columns: 1fr; gap: 4px; }
    }
</style>
</head>
<body>
<div class="wrap">
    <div class="topbar">
        <div class="title">
            <img id="headerLogo" src="__LOGO_URL__" alt="Logo" class="logo" style="display: none;">
            <div>
                <h1>__TITLE__</h1>
                <p>Generated: __GENERATED__</p>
            </div>
        </div>
        <div class="actions">
            <button onclick="copyVisibleTable()">Copy visible table</button>
            <button onclick="downloadCsv()">Export CSV</button>
            <button onclick="downloadJson()">Export JSON</button>
            <button onclick="downloadExcel()">Export Excel</button>
            <button onclick="window.print()">Print</button>
        </div>
    </div>

    <div class="cards">
        <div class="card"><div class="label">Total devices</div><div class="value">__TOTAL__</div></div>
        <div class="card"><div class="label">Healthy</div><div class="value" style="color:var(--green)">__HEALTHY__</div></div>
        <div class="card"><div class="label">Warning</div><div class="value" style="color:var(--yellow)">__WARNING__</div></div>
        <div class="card"><div class="label">Needs Attention</div><div class="value" style="color:var(--red)">__ATTENTION__</div></div>
    </div>

    <div class="controls">
        <input id="searchBox" type="text" placeholder="Search any field...  (Ctrl+F)" oninput="applyFilters()">
        <select id="statusFilter" onchange="applyFilters()">
            <option value="">All health states</option>
            <option value="Healthy">Healthy</option>
            <option value="Warning">Warning</option>
            <option value="Needs Attention">Needs Attention</option>
        </select>
        <select id="reachabilityFilter" onchange="applyFilters()">
            <option value="">All reachability</option>
            <option value="true">Reachable</option>
            <option value="false">Unreachable</option>
        </select>
        <select id="sortBy" onchange="applyFilters()">
            <option value="">-- Sort by column --</option>
            <option value="ComputerName">Computer name</option>
            <option value="HealthStatus">Health status</option>
            <option value="DiskC_FreePercent">Free disk %</option>
            <option value="LastHotfixDate">Last patch date</option>
            <option value="RAMGB">RAM (GB)</option>
            <option value="Uptime">Uptime</option>
        </select>
    </div>
    <div class="summary-bar" id="summaryBar"></div>

    <div class="table-wrap">
        <div class="table-scroll">
            <table id="reportTable">
                <thead>
                    <tr>
                        <th data-col="ComputerName">Computer</th>
                        <th data-col="HealthStatus">Health</th>
                        <th data-col="HealthIssues">Issues</th>
                        <th data-col="Reachable">Reachable</th>
                        <th data-col="CimSession">CIM</th>
                        <th data-col="LoggedOnUser">Logged On User</th>
                        <th data-col="OS">OS</th>
                        <th data-col="Version">Version</th>
                        <th data-col="BuildNumber">Build</th>
                        <th data-col="Uptime">Uptime</th>
                        <th data-col="CPU">CPU</th>
                        <th data-col="RAMGB">RAM (GB)</th>
                        <th data-col="DiskC_TotalGB">C: Total (GB)</th>
                        <th data-col="DiskC_FreeGB">C: Free (GB)</th>
                        <th data-col="DiskC_FreePercent">C: Free %</th>
                        <th data-col="IPv4">IPv4</th>
                        <th data-col="Antivirus">Antivirus</th>
                        <th data-col="DefenderRealtime">Defender RT</th>
                        <th data-col="LastHotfixId">Last Hotfix</th>
                        <th data-col="LastHotfixDate">Last Hotfix Date</th>
                        <th data-col="PendingRebootSignals">Pending Reboot</th>
                        <th data-col="InstalledSoftwareCount">SW Count</th>
                        <th data-col="RdpOpen">RDP Open</th>
                        <th data-col="WinRMOpen">WinRM Open</th>
                        <th data-col="Manufacturer">Manufacturer</th>
                        <th data-col="Model">Model</th>
                        <th data-col="SerialNumber">Serial</th>
                        <th data-col="Error">Error</th>
                    </tr>
                </thead>
                <tbody id="tableBody"></tbody>
            </table>
        </div>
    </div>

    <div class="footer">
        Exports apply to the currently visible filtered rows. Click any row to open the detail panel. Press Esc to close. Ctrl+F to search. Created & Maintained by jlee2834 (https://github.com/jlee2834)
    </div>
</div>

<div id="toast"></div>
<div id="panelOverlay" class="panel-overlay" onclick="closeSidePanel()"></div>
<div id="sidePanel" class="side-panel" aria-hidden="true">
    <div class="panel-header">
        <div class="panel-title">
            <h2 id="panelComputerName">Device Details</h2>
            <p id="panelSubtext">Select a row to view full device information.</p>
        </div>
        <button class="panel-close" onclick="closeSidePanel()">&#x2715;</button>
    </div>
    <div class="panel-body">
        <div class="panel-actions">
            <button onclick="copyPanelField('ComputerName')">Copy hostname</button>
            <button onclick="copyPanelField('IPv4')">Copy IP</button>
            <button onclick="copyPanelField('LoggedOnUser')">Copy user</button>
            <button onclick="copyCurrentDeviceJson()">Copy JSON</button>
        </div>
        <div id="panelContent" class="detail-grid">
            <div class="detail-card full">
                <div class="detail-label">No device selected</div>
                <div class="detail-value">Click a device row in the table to open its details here.</div>
            </div>
        </div>
    </div>
</div>

<script>
// Show logo only if a usable URL was injected
(function() {
    const logoEl = document.getElementById('headerLogo');
    const rawLogo = 'https://imgur.com/a/aBYVB2Y';
    if (logoEl && rawLogo && rawLogo.trim() !== '') {
        logoEl.style.display = 'block';
    }
})();

const rawDataSource = __JSON_DATA__;
const rawData = Array.isArray(rawDataSource) ? rawDataSource : [rawDataSource];
let filteredData = [...rawData];
let selectedDevice = null;
let sortCol = null;
let sortDir = 1;

function esc(value) {
    if (value === null || value === undefined) return '';
    return String(value)
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;')
        .replace(/'/g, '&#39;');
}

function toast(msg) {
    const el = document.getElementById('toast');
    el.textContent = msg;
    el.classList.add('show');
    clearTimeout(el._t);
    el._t = setTimeout(() => el.classList.remove('show'), 2200);
}

function boolCell(val) {
    if (val === true  || val === 'true'  || val === 'True')  return '<span class="bool-true">Yes</span>';
    if (val === false || val === 'false' || val === 'False') return '<span class="bool-false">No</span>';
    return esc(val);
}

function statusClass(status) {
    if (status === 'Healthy') return 'pill healthy';
    if (status === 'Warning') return 'pill warning';
    return 'pill attention';
}

function toComparableValue(item, key) {
    const value = item[key];
    if (value === null || value === undefined || value === '') return '';
    if (key === 'LastHotfixDate') {
        const d = new Date(value);
        return isNaN(d) ? '' : d.getTime();
    }
    if (typeof value === 'number') return value;
    if (value === true  || value === 'true'  || value === 'True')  return 1;
    if (value === false || value === 'false' || value === 'False') return 0;
    return String(value).toLowerCase();
}

function healthOrder(status) {
    if (status === 'Needs Attention') return 0;
    if (status === 'Warning') return 1;
    return 2;
}

function renderTable(data) {
    const body = document.getElementById('tableBody');
    body.innerHTML = data.map(function(item) {
        var sel = (selectedDevice && selectedDevice.ComputerName === item.ComputerName) ? 'selected-row' : '';
        var safeName = esc(item.ComputerName).replace(/'/g, '&#39;');
        return '<tr class="clickable-row ' + sel + '" onclick="openSidePanelByName(\'' + safeName + '\')">' +
            '<td>' + esc(item.ComputerName) + '</td>' +
            '<td><span class="' + statusClass(item.HealthStatus) + '">' + esc(item.HealthStatus) + '</span></td>' +
            '<td>' + esc(item.HealthIssues) + '</td>' +
            '<td>' + boolCell(item.Reachable) + '</td>' +
            '<td>' + boolCell(item.CimSession) + '</td>' +
            '<td>' + esc(item.LoggedOnUser) + '</td>' +
            '<td>' + esc(item.OS) + '</td>' +
            '<td>' + esc(item.Version) + '</td>' +
            '<td>' + esc(item.BuildNumber) + '</td>' +
            '<td>' + esc(item.Uptime) + '</td>' +
            '<td>' + esc(item.CPU) + '</td>' +
            '<td>' + esc(item.RAMGB) + '</td>' +
            '<td>' + esc(item.DiskC_TotalGB) + '</td>' +
            '<td>' + esc(item.DiskC_FreeGB) + '</td>' +
            '<td>' + esc(item.DiskC_FreePercent) + '</td>' +
            '<td>' + esc(item.IPv4) + '</td>' +
            '<td>' + esc(item.Antivirus) + '</td>' +
            '<td>' + boolCell(item.DefenderRealtime) + '</td>' +
            '<td>' + esc(item.LastHotfixId) + '</td>' +
            '<td>' + esc(item.LastHotfixDate) + '</td>' +
            '<td>' + esc(item.PendingRebootSignals) + '</td>' +
            '<td>' + esc(item.InstalledSoftwareCount) + '</td>' +
            '<td>' + boolCell(item.RdpOpen) + '</td>' +
            '<td>' + boolCell(item.WinRMOpen) + '</td>' +
            '<td>' + esc(item.Manufacturer) + '</td>' +
            '<td>' + esc(item.Model) + '</td>' +
            '<td>' + esc(item.SerialNumber) + '</td>' +
            '<td>' + esc(item.Error) + '</td>' +
            '</tr>';
    }).join('');

    document.getElementById('summaryBar').textContent =
        'Showing ' + data.length + ' of ' + rawData.length + ' device(s)';
}

// Column header click-to-sort
document.querySelectorAll('thead th[data-col]').forEach(function(th) {
    th.addEventListener('click', function() {
        var col = th.getAttribute('data-col');
        if (sortCol === col) {
            sortDir *= -1;
        } else {
            sortCol = col;
            sortDir = 1;
        }
        document.querySelectorAll('thead th').forEach(function(t) {
            t.classList.remove('sort-asc', 'sort-desc');
        });
        th.classList.add(sortDir === 1 ? 'sort-asc' : 'sort-desc');
        applyFilters();
    });
});

function applyFilters() {
    var search       = document.getElementById('searchBox').value.trim().toLowerCase();
    var status       = document.getElementById('statusFilter').value;
    var reachability = document.getElementById('reachabilityFilter').value;
    var dropSort     = document.getElementById('sortBy').value;

    filteredData = rawData.filter(function(item) {
        var statusOk       = !status       || item.HealthStatus === status;
        var reachabilityOk = !reachability || String(item.Reachable) === reachability;
        var searchOk       = !search       || Object.values(item).some(function(v) {
            return String(v == null ? '' : v).toLowerCase().indexOf(search) !== -1;
        });
        return statusOk && reachabilityOk && searchOk;
    });

    var effectiveCol = sortCol || dropSort;

    if (effectiveCol) {
        filteredData.sort(function(a, b) {
            if (effectiveCol === 'HealthStatus') {
                var hcmp = healthOrder(a.HealthStatus) - healthOrder(b.HealthStatus);
                return (sortDir * hcmp) || String(a.ComputerName).localeCompare(String(b.ComputerName));
            }
            var av = toComparableValue(a, effectiveCol);
            var bv = toComparableValue(b, effectiveCol);
            var cmp = (typeof av === 'number' && typeof bv === 'number')
                ? av - bv
                : String(av).localeCompare(String(bv));
            return cmp * sortDir;
        });
    }

    if (selectedDevice) {
        var stillVisible = filteredData.find(function(x) { return x.ComputerName === selectedDevice.ComputerName; });
        if (!stillVisible) { selectedDevice = null; closeSidePanel(); }
        else { selectedDevice = stillVisible; renderSidePanel(selectedDevice); }
    }

    renderTable(filteredData);
}

function openSidePanelByName(name) {
    var device = filteredData.find(function(x) { return x.ComputerName === name; }) ||
                 rawData.find(function(x) { return x.ComputerName === name; });
    if (!device) return;
    selectedDevice = device;
    renderSidePanel(device);
    document.getElementById('panelOverlay').classList.add('open');
    var panel = document.getElementById('sidePanel');
    panel.classList.add('open');
    panel.setAttribute('aria-hidden', 'false');
    renderTable(filteredData);
}

function closeSidePanel() {
    document.getElementById('panelOverlay').classList.remove('open');
    var panel = document.getElementById('sidePanel');
    panel.classList.remove('open');
    panel.setAttribute('aria-hidden', 'true');
    renderTable(filteredData);
}

function detailCard(title, value, full) {
    var cls = full ? 'detail-card full' : 'detail-card';
    return '<div class="' + cls + '"><div class="detail-label">' + esc(title) +
           '</div><div class="detail-value">' + esc(value == null ? '' : value) + '</div></div>';
}

function kvRow(label, value) {
    return '<div class="k">' + esc(label) + '</div><div class="v">' + esc(value == null ? '' : value) + '</div>';
}

function sectionCard(cssClass, title, subtitle, innerHtml) {
    return '<section class="section-card ' + cssClass + '">' +
        '<div class="section-header"><div>' +
            '<div class="section-title">' + esc(title) + '</div>' +
            '<div class="section-subtitle">' + esc(subtitle) + '</div>' +
        '</div></div>' +
        '<div class="section-content">' + innerHtml + '</div>' +
        '</section>';
}

function getHealthSectionClass(status) {
    if (status === 'Healthy') return 'health-healthy';
    if (status === 'Warning') return 'health-warning';
    return 'health-attention';
}

function renderSidePanel(device) {
    document.getElementById('panelComputerName').textContent = device.ComputerName || 'Device Details';
    document.getElementById('panelSubtext').textContent =
        (device.OS || 'Unknown OS') + ' \u2022 ' + (device.IPv4 || 'No IP found');

    var healthBadge = '<span class="' + statusClass(device.HealthStatus) + '">' + esc(device.HealthStatus) + '</span>';
    var healthClass = getHealthSectionClass(device.HealthStatus);

    var overviewSection = sectionCard('section-overview', 'Overview', 'Primary device summary',
        '<div class="detail-card full"><div class="detail-label">Device Summary</div><div class="kv">' +
        kvRow('Computer Name', device.ComputerName) +
        kvRow('Logged On User', device.LoggedOnUser) +
        kvRow('Reachable', device.Reachable) +
        kvRow('CIM Session', device.CimSession) +
        kvRow('Uptime', device.Uptime) +
        kvRow('Operating System', ((device.OS || '') + ' ' + (device.Version || '')).trim()) +
        '</div></div>'
    );

    var hardwareSection = sectionCard('section-hardware', 'Hardware', 'System identity and physical resources',
        detailCard('Manufacturer', device.Manufacturer) +
        detailCard('Model', device.Model) +
        detailCard('Serial Number', device.SerialNumber) +
        detailCard('Architecture', device.Architecture) +
        detailCard('CPU', device.CPU, true) +
        detailCard('RAM (GB)', device.RAMGB) +
        detailCard('C: Total (GB)', device.DiskC_TotalGB) +
        detailCard('C: Free (GB)', device.DiskC_FreeGB) +
        detailCard('C: Free %', device.DiskC_FreePercent)
    );

    var securitySection = sectionCard('section-security', 'Security', 'Protection state and exposure checks',
        detailCard('Health State', device.HealthStatus) +
        detailCard('Health Issues', device.HealthIssues, true) +
        detailCard('Antivirus', device.Antivirus, true) +
        detailCard('Defender Realtime', device.DefenderRealtime) +
        detailCard('Installed Software Count', device.InstalledSoftwareCount) +
        detailCard('Collection Error', device.Error || 'None', true)
    );

    var patchingSection = sectionCard('section-patching', 'Patching', 'Update visibility and reboot indicators',
        detailCard('Last Hotfix ID', device.LastHotfixId) +
        detailCard('Last Hotfix Date', device.LastHotfixDate) +
        detailCard('Pending Reboot Signals', device.PendingRebootSignals) +
        detailCard('Build Number', device.BuildNumber)
    );

    var connectivitySection = sectionCard('section-connectivity', 'Connectivity', 'Network addressing and remote access exposure',
        detailCard('IPv4', device.IPv4, true) +
        detailCard('RDP Open', device.RdpOpen) +
        detailCard('WinRM Open', device.WinRMOpen) +
        detailCard('Last Boot Time', device.LastBoot, true)
    );

    document.getElementById('panelContent').innerHTML =
        '<section class="section-card health-overview ' + healthClass + '">' +
            '<div class="section-header"><div>' +
                '<div class="section-title">Health Overview</div>' +
                '<div class="section-subtitle">Immediate status and issue summary</div>' +
            '</div><div>' + healthBadge + '</div></div>' +
            '<div class="section-content"><div class="detail-card full">' +
                '<div class="detail-label">Current Condition</div>' +
                '<div class="detail-value">' + esc(device.HealthIssues) + '</div>' +
            '</div></div>' +
        '</section>' +
        overviewSection + hardwareSection + securitySection + patchingSection + connectivitySection;
}

function copyPanelField(field) {
    if (!selectedDevice) { toast('No device selected.'); return; }
    navigator.clipboard.writeText(String(selectedDevice[field] == null ? '' : selectedDevice[field]))
        .then(function() { toast(field + ' copied.'); })
        .catch(function() { toast('Clipboard copy failed.'); });
}

function copyCurrentDeviceJson() {
    if (!selectedDevice) { toast('No device selected.'); return; }
    navigator.clipboard.writeText(JSON.stringify(selectedDevice, null, 2))
        .then(function() { toast('Device JSON copied.'); })
        .catch(function() { toast('Clipboard copy failed.'); });
}

document.addEventListener('keydown', function(e) {
    if (e.key === 'Escape') { closeSidePanel(); return; }
    if ((e.ctrlKey || e.metaKey) && e.key === 'f') {
        e.preventDefault();
        document.getElementById('searchBox').focus();
    }
});

var EXPORT_HEADERS = [
    'ComputerName','HealthStatus','HealthIssues','Reachable','CimSession','LoggedOnUser',
    'OS','Version','BuildNumber','Uptime','CPU','RAMGB','DiskC_TotalGB','DiskC_FreeGB',
    'DiskC_FreePercent','IPv4','Antivirus','DefenderRealtime','LastHotfixId','LastHotfixDate',
    'PendingRebootSignals','InstalledSoftwareCount','RdpOpen','WinRMOpen',
    'Manufacturer','Model','SerialNumber','Error'
];

function rowsToCsv(rows) {
    var escapeCsv = function(v) {
        var t = String(v == null ? '' : v);
        return /[",\n]/.test(t) ? '"' + t.replace(/"/g, '""') + '"' : t;
    };
    var lines = [EXPORT_HEADERS.join(',')];
    rows.forEach(function(r) {
        lines.push(EXPORT_HEADERS.map(function(h) { return escapeCsv(r[h]); }).join(','));
    });
    return lines.join('\r\n');
}

function downloadBlob(content, filename, mime) {
    var blob = new Blob([content], { type: mime });
    var url  = URL.createObjectURL(blob);
    var a    = document.createElement('a');
    a.href = url; a.download = filename;
    document.body.appendChild(a); a.click(); a.remove();
    URL.revokeObjectURL(url);
}

function copyVisibleTable() {
    navigator.clipboard.writeText(rowsToCsv(filteredData))
        .then(function() { toast('Table copied as CSV.'); })
        .catch(function() { toast('Clipboard copy failed.'); });
}

function downloadCsv()  { 
    downloadBlob(rowsToCsv(filteredData), '__CSV_FILE__', 'text/csv;charset=utf-8;');
    toast('CSV file exported successfully.');
}

function downloadJson() { 
    downloadBlob(JSON.stringify(filteredData, null, 2), '__JSON_FILE__', 'application/json;charset=utf-8;');
    toast('JSON file exported successfully.');
}

function downloadExcel() {
    var rows = filteredData.map(function(r) {
        return '<tr>' + EXPORT_HEADERS.map(function(h) { return '<td>' + esc(r[h]) + '</td>'; }).join('') + '</tr>';
    }).join('');
    var html = '<html xmlns:o="urn:schemas-microsoft-com:office:office" ' +
        'xmlns:x="urn:schemas-microsoft-com:office:excel" xmlns="http://www.w3.org/TR/REC-html40">' +
        '<head><meta charset="UTF-8"></head><body><table border="1">' +
        '<tr>' + EXPORT_HEADERS.map(function(h) { return '<th>' + h + '</th>'; }).join('') + '</tr>' +
        rows + '</table></body></html>';
    downloadBlob(html, '__EXCEL_FILE__', 'application/vnd.ms-excel;charset=utf-8;');
    toast('Excel file exported successfully.');
}

applyFilters();
</script>
</body>
</html>
'@

    $html = $html -replace '__TITLE__',      $Title
    $html = $html -replace '__GENERATED__',  $generated
    $html = $html -replace '__TOTAL__',      [string]$total
    $html = $html -replace '__HEALTHY__',    [string]$healthy
    $html = $html -replace '__WARNING__',    [string]$warning
    $html = $html -replace '__ATTENTION__',  [string]$attention
    $html = $html -replace '__JSON_DATA__',  $jsonEscaped
    $html = $html -replace '__CSV_FILE__',   $CsvFileName
    $html = $html -replace '__JSON_FILE__',  $JsonFileName
    $html = $html -replace '__EXCEL_FILE__', $ExcelFileName
    $html = $html -replace '__LOGO_URL__',   $LogoUrl

    return $html
}

# ── Entry point ────────────────────────────────────────────────────────────
Ensure-OutputDir -Path $OutputDir

$targets = Resolve-ComputerList -Names $ComputerName -File $InputFile
if (-not $targets -or @($targets).Count -eq 0) {
    throw 'No computer names were supplied.'
}

$stamp   = Get-Date -Format 'yyyyMMdd_HHmmss'
$results = New-Object System.Collections.Generic.List[object]

foreach ($target in $targets) {
    Write-Host "Collecting from $target ..." -ForegroundColor Cyan
    $item = Get-RemoteInventory -Computer $target -SkipHotfixes:$SkipHotfixes
    $results.Add($item)
    Start-Sleep -Milliseconds $ThrottleDelayMs
}

$results  = $results | Sort-Object HealthStatus, ComputerName

$htmlPath = $null
if ($ExportHtml) {
    $htmlPath  = Join-Path $OutputDir "inventory_$stamp.html"
    $csvName   = "inventory_$stamp.csv"
    $jsonName  = "inventory_$stamp.json"
    $excelName = "inventory_$stamp.xls"

    $html = ConvertTo-PrettyHtmlReport `
        -Data          $results `
        -Title         'Argus Insight' `
        -CsvFileName   $csvName `
        -JsonFileName  $jsonName `
        -ExcelFileName $excelName `
        -LogoUrl       $LogoUrl

    Set-Content -Path $htmlPath -Value $html -Encoding UTF8
}

$results | Format-Table -AutoSize
Write-Host "`nGenerated HTML report in memory" -ForegroundColor Green
if ($htmlPath) {
    Write-Host "Saved HTML: $htmlPath" -ForegroundColor Green
    Write-Host "Note: CSV/JSON/Excel files are generated on-demand via HTML buttons" -ForegroundColor Cyan
    if ($OpenReport) { Start-Process $htmlPath }
}
