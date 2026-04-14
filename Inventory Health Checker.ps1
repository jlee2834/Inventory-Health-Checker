param(
    [string[]]$ComputerName = @($env:COMPUTERNAME),
    [string]$InputFile,
    [string]$OutputDir = ".\output",
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
        $wait = $async.AsyncWaitHandle.WaitOne($TimeoutMs, $false)
        if (-not $wait) {
            return $false
        }
        $client.EndConnect($async)
        return $true
    }
    catch {
        return $false
    }
    finally {
        $client.Close()
    }
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
        [Nullable[datetime]]$LastHotfixDate,
        [bool]$HasError
    )

    $issues = New-Object System.Collections.Generic.List[string]

    if ($HasError) {
        $issues.Add('Collection error')
    }

    if ($DiskFreePercent -lt 10) {
        $issues.Add('Critical disk space')
    }
    elseif ($DiskFreePercent -lt 20) {
        $issues.Add('Low disk space')
    }

    if ($PendingRebootSignals -gt 0) {
        $issues.Add('Pending reboot')
    }

    if (-not $DefenderRealtime) {
        $issues.Add('Defender real-time protection off or unavailable')
    }

    if ($LastHotfixDate.HasValue) {
        $daysSincePatch = ((Get-Date) - $LastHotfixDate.Value).Days
        if ($daysSincePatch -gt 45) {
            $issues.Add('Patch age over 45 days')
        }
    }
    else {
        $issues.Add('Patch date unavailable')
    }

    if ($issues.Count -eq 0) {
        return [pscustomobject]@{
            Status = 'Healthy'
            Issues = 'None'
        }
    }

    $status = 'Warning'
    if ($issues -contains 'Collection error' -or $issues -contains 'Critical disk space' -or $issues -contains 'Defender real-time protection off or unavailable') {
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
        'SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Component Based Servicing\\RebootPending',
        'SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\WindowsUpdate\\Auto Update\\RebootRequired'
    )

    try {
        $base = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey('LocalMachine', $Computer)
        foreach ($path in $paths) {
            $sub = $base.OpenSubKey($path)
            if ($sub) {
                $signals++
                $sub.Close()
            }
        }

        $sessionManager = $base.OpenSubKey('SYSTEM\\CurrentControlSet\\Control\\Session Manager')
        if ($sessionManager) {
            $pending = $sessionManager.GetValue('PendingFileRenameOperations', $null)
            if ($pending) {
                $signals++
            }
            $sessionManager.Close()
        }

        $base.Close()
    }
    catch {
    }

    return $signals
}

function Get-InstalledSoftwareCount {
    param([string]$Computer)

    $count = 0
    $paths = @(
        'SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Uninstall',
        'SOFTWARE\\WOW6432Node\\Microsoft\\Windows\\CurrentVersion\\Uninstall'
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
    catch {
        return $null
    }

    return $count
}

function Get-RemoteInventory {
    param(
        [Parameter(Mandatory)][string]$Computer,
        [switch]$SkipHotfixes
    )

    $session = $null
    $result = [ordered]@{
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
        $result.RdpOpen = Test-TcpPort -Computer $Computer -Port 3389
        $result.WinRMOpen = Test-TcpPort -Computer $Computer -Port 5985

        $session = New-CimSession -ComputerName $Computer -ErrorAction Stop
        $result.CimSession = $true

        $os = Get-CimInstance -CimSession $session -ClassName Win32_OperatingSystem
        $cs = Get-CimInstance -CimSession $session -ClassName Win32_ComputerSystem
        $bios = Get-CimInstance -CimSession $session -ClassName Win32_BIOS
        $cpu = Get-CimInstance -CimSession $session -ClassName Win32_Processor | Select-Object -First 1
        $disk = Get-CimInstance -CimSession $session -ClassName Win32_LogicalDisk -Filter "DeviceID='C:'"
        $nic = Get-CimInstance -CimSession $session -ClassName Win32_NetworkAdapterConfiguration -Filter 'IPEnabled = True'

        $result.OS = $os.Caption
        $result.Version = $os.Version
        $result.BuildNumber = $os.BuildNumber
        $result.Architecture = $os.OSArchitecture
        $result.LastBoot = $os.LastBootUpTime
        $result.Uptime = Get-LastBootReadable -LastBoot $os.LastBootUpTime
        $result.LoggedOnUser = $cs.UserName
        $result.SerialNumber = $bios.SerialNumber
        $result.Manufacturer = $cs.Manufacturer
        $result.Model = $cs.Model
        $result.CPU = $cpu.Name
        $result.RAMGB = [math]::Round($cs.TotalPhysicalMemory / 1GB, 2)

        if ($disk) {
            $result.DiskC_TotalGB = [math]::Round($disk.Size / 1GB, 2)
            $result.DiskC_FreeGB = [math]::Round($disk.FreeSpace / 1GB, 2)
            $result.DiskC_FreePercent = if ($disk.Size -gt 0) { [math]::Round(($disk.FreeSpace / $disk.Size) * 100, 2) } else { 0 }
        }

        $ipv4List = @()
        foreach ($adapter in $nic) {
            foreach ($ip in $adapter.IPAddress) {
                if ($ip -match '^\d+\.\d+\.\d+\.\d+$') {
                    $ipv4List += $ip
                }
            }
        }
        $result.IPv4 = ($ipv4List | Sort-Object -Unique) -join ', '

        try {
            $defender = Invoke-Command -ComputerName $Computer -ScriptBlock {
                if (Get-Command Get-MpComputerStatus -ErrorAction SilentlyContinue) {
                    Get-MpComputerStatus | Select-Object AMProductVersion, AntivirusEnabled, RealTimeProtectionEnabled
                }
            } -ErrorAction Stop

            if ($defender) {
                $result.Antivirus = "Microsoft Defender $($defender.AMProductVersion)"
                $result.DefenderRealtime = [bool]$defender.RealTimeProtectionEnabled
            }
            else {
                $result.Antivirus = 'Unknown or unavailable'
            }
        }
        catch {
            $result.Antivirus = 'Unknown or unavailable'
        }

        if (-not $SkipHotfixes) {
            try {
                $hotfix = Get-HotFix -ComputerName $Computer -ErrorAction Stop |
                    Sort-Object InstalledOn -Descending |
                    Select-Object -First 1

                if ($hotfix) {
                    $result.LastHotfixId = $hotfix.HotFixID
                    $result.LastHotfixDate = $hotfix.InstalledOn
                }
            }
            catch {
                $result.LastHotfixId = 'Unavailable'
            }
        }

        $result.PendingRebootSignals = Get-PendingRebootSignals -Computer $Computer
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

    $health = Get-HealthStatus `
        -DiskFreePercent ([double]($result.DiskC_FreePercent ?? 0)) `
        -PendingRebootSignals ([int]$result.PendingRebootSignals) `
        -DefenderRealtime ([bool]$result.DefenderRealtime) `
        -LastHotfixDate $result.LastHotfixDate `
        -HasError ([bool](-not [string]::IsNullOrWhiteSpace($result.Error)))

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
        [Parameter(Mandatory)][string]$ExcelFileName
    )

    $json = $Data | ConvertTo-Json -Depth 6 -Compress
    $jsonEscaped = $json.Replace('</script>', '<\/script>')
    $generated = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
    $total = $Data.Count
    $healthy = ($Data | Where-Object { $_.HealthStatus -eq 'Healthy' }).Count
    $warning = ($Data | Where-Object { $_.HealthStatus -eq 'Warning' }).Count
    $attention = ($Data | Where-Object { $_.HealthStatus -eq 'Needs Attention' }).Count

    return @"
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>$Title</title>
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
    .wrap {
        max-width: 1600px;
        margin: 0 auto;
        padding: 24px;
    }
    .topbar {
        display: flex;
        justify-content: space-between;
        align-items: flex-start;
        gap: 16px;
        margin-bottom: 20px;
        flex-wrap: wrap;
    }
    .title h1 {
        margin: 0 0 8px 0;
        font-size: 30px;
    }
    .title p {
        margin: 0;
        color: var(--muted);
    }
    .actions {
        display: flex;
        gap: 10px;
        flex-wrap: wrap;
    }
    button {
        background: var(--panel-2);
        color: var(--text);
        border: 1px solid var(--line);
        padding: 10px 14px;
        border-radius: 10px;
        cursor: pointer;
        font-weight: 600;
    }
    button:hover {
        border-color: var(--blue);
        transform: translateY(-1px);
    }
    .cards {
        display: grid;
        grid-template-columns: repeat(auto-fit, minmax(220px, 1fr));
        gap: 14px;
        margin-bottom: 18px;
    }
    .card {
        background: rgba(18, 26, 43, 0.92);
        border: 1px solid var(--line);
        border-radius: 16px;
        padding: 18px;
        box-shadow: 0 10px 30px rgba(0,0,0,0.22);
    }
    .card .label {
        color: var(--muted);
        font-size: 13px;
        margin-bottom: 10px;
    }
    .card .value {
        font-size: 30px;
        font-weight: 700;
    }
    .controls {
        display: grid;
        grid-template-columns: 1.5fr 1fr 1fr 1fr;
        gap: 12px;
        margin-bottom: 16px;
    }
    .controls input, .controls select {
        width: 100%;
        background: var(--panel);
        color: var(--text);
        border: 1px solid var(--line);
        border-radius: 12px;
        padding: 12px;
    }
    .table-wrap {
        background: rgba(18, 26, 43, 0.92);
        border: 1px solid var(--line);
        border-radius: 18px;
        overflow: hidden;
        box-shadow: 0 10px 30px rgba(0,0,0,0.22);
    }
    .table-scroll {
        overflow: auto;
        max-height: 72vh;
    }
    table {
        width: 100%;
        border-collapse: collapse;
        min-width: 1500px;
    }
    thead th {
        position: sticky;
        top: 0;
        z-index: 1;
        background: #14203a;
        color: var(--text);
        text-align: left;
        padding: 12px;
        border-bottom: 1px solid var(--line);
        font-size: 13px;
        white-space: nowrap;
    }
    tbody td {
        padding: 11px 12px;
        border-bottom: 1px solid rgba(41,54,83,0.55);
        color: #dce6fb;
        font-size: 13px;
        vertical-align: top;
    }
    tbody tr:hover {
        background: rgba(96, 165, 250, 0.06);
    }
    .pill {
        display: inline-block;
        padding: 5px 10px;
        border-radius: 999px;
        font-weight: 700;
        font-size: 12px;
        white-space: nowrap;
    }
    .healthy { background: rgba(34,197,94,0.15); color: #86efac; border: 1px solid rgba(34,197,94,0.35); }
    .warning { background: rgba(245,158,11,0.15); color: #fcd34d; border: 1px solid rgba(245,158,11,0.35); }
    .attention { background: rgba(239,68,68,0.15); color: #fca5a5; border: 1px solid rgba(239,68,68,0.35); }
    .muted { color: var(--muted); }
    .footer {
        margin-top: 12px;
        color: var(--muted);
        font-size: 12px;
    }
    @media (max-width: 1100px) {
        .controls {
            grid-template-columns: 1fr 1fr;
        }
    }
    @media (max-width: 720px) {
        .controls {
            grid-template-columns: 1fr;
        }
        .title h1 {
            font-size: 24px;
        }
    }
</style>
</head>
<body>
<div class="wrap">
    <div class="topbar">
        <div class="title">
            <h1>$Title</h1>
            <p>Generated: $generated</p>
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
        <div class="card"><div class="label">Total devices</div><div class="value">$total</div></div>
        <div class="card"><div class="label">Healthy</div><div class="value">$healthy</div></div>
        <div class="card"><div class="label">Warning</div><div class="value">$warning</div></div>
        <div class="card"><div class="label">Needs Attention</div><div class="value">$attention</div></div>
    </div>

    <div class="controls">
        <input id="searchBox" type="text" placeholder="Search any field..." oninput="applyFilters()">
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
            <option value="ComputerName">Sort by computer name</option>
            <option value="HealthStatus">Sort by health status</option>
            <option value="DiskC_FreePercent">Sort by free disk %</option>
            <option value="LastHotfixDate">Sort by last patch date</option>
            <option value="Uptime">Sort by uptime text</option>
        </select>
    </div>

    <div class="table-wrap">
        <div class="table-scroll">
            <table id="reportTable">
                <thead>
                    <tr>
                        <th>Computer</th>
                        <th>Health</th>
                        <th>Issues</th>
                        <th>Reachable</th>
                        <th>CIM</th>
                        <th>Logged On User</th>
                        <th>OS</th>
                        <th>Version</th>
                        <th>Build</th>
                        <th>Uptime</th>
                        <th>CPU</th>
                        <th>RAM (GB)</th>
                        <th>C: Total (GB)</th>
                        <th>C: Free (GB)</th>
                        <th>C: Free %</th>
                        <th>IPv4</th>
                        <th>Antivirus</th>
                        <th>Defender Realtime</th>
                        <th>Last Hotfix</th>
                        <th>Last Hotfix Date</th>
                        <th>Pending Reboot Signals</th>
                        <th>Installed Software Count</th>
                        <th>RDP Open</th>
                        <th>WinRM Open</th>
                        <th>Manufacturer</th>
                        <th>Model</th>
                        <th>Serial</th>
                        <th>Error</th>
                    </tr>
                </thead>
                <tbody id="tableBody"></tbody>
            </table>
        </div>
    </div>

    <div class="footer">
        Buttons export the currently visible filtered data. Excel export creates an Excel-openable .xls file from the filtered table.
    </div>
</div>

<script>
const rawData = $jsonEscaped;
let filteredData = [...rawData];

function esc(value) {
    if (value === null || value === undefined) return '';
    return String(value)
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;')
        .replace(/'/g, '&#39;');
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
    return String(value).toLowerCase();
}

function healthOrder(status) {
    if (status === 'Needs Attention') return 0;
    if (status === 'Warning') return 1;
    return 2;
}

function renderTable(data) {
    const body = document.getElementById('tableBody');
    body.innerHTML = data.map(item => `
        <tr>
            <td>${esc(item.ComputerName)}</td>
            <td><span class="${statusClass(item.HealthStatus)}">${esc(item.HealthStatus)}</span></td>
            <td>${esc(item.HealthIssues)}</td>
            <td>${esc(item.Reachable)}</td>
            <td>${esc(item.CimSession)}</td>
            <td>${esc(item.LoggedOnUser)}</td>
            <td>${esc(item.OS)}</td>
            <td>${esc(item.Version)}</td>
            <td>${esc(item.BuildNumber)}</td>
            <td>${esc(item.Uptime)}</td>
            <td>${esc(item.CPU)}</td>
            <td>${esc(item.RAMGB)}</td>
            <td>${esc(item.DiskC_TotalGB)}</td>
            <td>${esc(item.DiskC_FreeGB)}</td>
            <td>${esc(item.DiskC_FreePercent)}</td>
            <td>${esc(item.IPv4)}</td>
            <td>${esc(item.Antivirus)}</td>
            <td>${esc(item.DefenderRealtime)}</td>
            <td>${esc(item.LastHotfixId)}</td>
            <td>${esc(item.LastHotfixDate)}</td>
            <td>${esc(item.PendingRebootSignals)}</td>
            <td>${esc(item.InstalledSoftwareCount)}</td>
            <td>${esc(item.RdpOpen)}</td>
            <td>${esc(item.WinRMOpen)}</td>
            <td>${esc(item.Manufacturer)}</td>
            <td>${esc(item.Model)}</td>
            <td>${esc(item.SerialNumber)}</td>
            <td>${esc(item.Error)}</td>
        </tr>
    `).join('');
}

function applyFilters() {
    const search = document.getElementById('searchBox').value.trim().toLowerCase();
    const status = document.getElementById('statusFilter').value;
    const reachability = document.getElementById('reachabilityFilter').value;
    const sortBy = document.getElementById('sortBy').value;

    filteredData = rawData.filter(item => {
        const statusOk = !status || item.HealthStatus === status;
        const reachabilityOk = !reachability || String(item.Reachable) === reachability;
        const searchOk = !search || Object.values(item).some(v => String(v ?? '').toLowerCase().includes(search));
        return statusOk && reachabilityOk && searchOk;
    });

    filteredData.sort((a, b) => {
        if (sortBy === 'HealthStatus') {
            return healthOrder(a.HealthStatus) - healthOrder(b.HealthStatus) || String(a.ComputerName).localeCompare(String(b.ComputerName));
        }
        const av = toComparableValue(a, sortBy);
        const bv = toComparableValue(b, sortBy);
        if (typeof av === 'number' && typeof bv === 'number') return av - bv;
        return String(av).localeCompare(String(bv));
    });

    renderTable(filteredData);
}

function rowsToCsv(rows) {
    const headers = [
        'ComputerName','HealthStatus','HealthIssues','Reachable','CimSession','LoggedOnUser','OS','Version','BuildNumber','Uptime',
        'CPU','RAMGB','DiskC_TotalGB','DiskC_FreeGB','DiskC_FreePercent','IPv4','Antivirus','DefenderRealtime','LastHotfixId',
        'LastHotfixDate','PendingRebootSignals','InstalledSoftwareCount','RdpOpen','WinRMOpen','Manufacturer','Model','SerialNumber','Error'
    ];

    const escapeCsv = value => {
        const text = String(value ?? '');
        if (/[",\n]/.test(text)) return '"' + text.replace(/"/g, '""') + '"';
        return text;
    };

    const lines = [];
    lines.push(headers.join(','));
    for (const row of rows) {
        lines.push(headers.map(h => escapeCsv(row[h])).join(','));
    }
    return lines.join('\r\n');
}

function downloadBlob(content, filename, mime) {
    const blob = new Blob([content], { type: mime });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    a.remove();
    URL.revokeObjectURL(url);
}

function copyVisibleTable() {
    const csv = rowsToCsv(filteredData);
    navigator.clipboard.writeText(csv).then(() => {
        alert('Visible table copied to clipboard as CSV text.');
    }).catch(() => {
        alert('Clipboard copy failed in this browser.');
    });
}

function downloadCsv() {
    downloadBlob(rowsToCsv(filteredData), '$CsvFileName', 'text/csv;charset=utf-8;');
}

function downloadJson() {
    downloadBlob(JSON.stringify(filteredData, null, 2), '$JsonFileName', 'application/json;charset=utf-8;');
}

function downloadExcel() {
    const headers = [
        'ComputerName','HealthStatus','HealthIssues','Reachable','CimSession','LoggedOnUser','OS','Version','BuildNumber','Uptime',
        'CPU','RAMGB','DiskC_TotalGB','DiskC_FreeGB','DiskC_FreePercent','IPv4','Antivirus','DefenderRealtime','LastHotfixId',
        'LastHotfixDate','PendingRebootSignals','InstalledSoftwareCount','RdpOpen','WinRMOpen','Manufacturer','Model','SerialNumber','Error'
    ];

    let tableRows = filteredData.map(row => '<tr>' + headers.map(h => `<td>${esc(row[h])}</td>`).join('') + '</tr>').join('');
    let html = `
        <html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel" xmlns="http://www.w3.org/TR/REC-html40">
        <head><meta charset="UTF-8"></head>
        <body>
            <table border="1">
                <tr>${headers.map(h => `<th>${h}</th>`).join('')}</tr>
                ${tableRows}
            </table>
        </body>
        </html>`;

    downloadBlob(html, '$ExcelFileName', 'application/vnd.ms-excel;charset=utf-8;');
}

applyFilters();
</script>
</body>
</html>
"@
}

Ensure-OutputDir -Path $OutputDir
$targets = Resolve-ComputerList -Names $ComputerName -File $InputFile
if (-not $targets -or $targets.Count -eq 0) {
    throw 'No computer names were supplied.'
}

$stamp = Get-Date -Format 'yyyyMMdd_HHmmss'
$results = New-Object System.Collections.Generic.List[object]

foreach ($target in $targets) {
    Write-Host "Collecting from $target ..." -ForegroundColor Cyan
    $item = Get-RemoteInventory -Computer $target -SkipHotfixes:$SkipHotfixes
    $results.Add($item)
    Start-Sleep -Milliseconds $ThrottleDelayMs
}

$results = $results | Sort-Object HealthStatus, ComputerName
$csvPath = Join-Path $OutputDir "inventory_$stamp.csv"
$jsonPath = Join-Path $OutputDir "inventory_$stamp.json"
$results | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8
$results | ConvertTo-Json -Depth 6 | Out-File -FilePath $jsonPath -Encoding utf8

$htmlPath = $null
if ($ExportHtml) {
    $htmlPath = Join-Path $OutputDir "inventory_$stamp.html"
    $csvName = [System.IO.Path]::GetFileName($csvPath)
    $jsonName = [System.IO.Path]::GetFileName($jsonPath)
    $excelName = "inventory_$stamp.xls"

    $html = ConvertTo-PrettyHtmlReport `
        -Data $results `
        -Title 'IT Help Desk Inventory Dashboard' `
        -CsvFileName $csvName `
        -JsonFileName $jsonName `
        -ExcelFileName $excelName

    Set-Content -Path $htmlPath -Value $html -Encoding UTF8
}

$results | Format-Table -AutoSize
Write-Host "`nSaved CSV:  $csvPath" -ForegroundColor Green
Write-Host "Saved JSON: $jsonPath" -ForegroundColor Green
if ($htmlPath) {
    Write-Host "Saved HTML: $htmlPath" -ForegroundColor Green
    if ($OpenReport) {
        Start-Process $htmlPath
    }
}
