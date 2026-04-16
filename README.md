# Argus Insight

A powerful PowerShell-based Windows system inventory and health assessment tool that provides deep visibility into device infrastructure. Generates beautiful, interactive HTML reports with comprehensive system metrics, security status, and device health insights.

## 🎯 Overview

**Argus Insight** collects detailed information from one or multiple Windows computers and generates an interactive HTML dashboard with:
- Real-time device health status (Healthy, Warning, Needs Attention)
- Hardware inventory and specifications
- Operating system and patch information
- Security & antivirus status
- Network connectivity and accessibility
- Customizable logo and branding
- Advanced filtering, sorting, and search capabilities
- One-click exports (CSV, JSON, Excel)

## ✨ Features

### Data Collection
- **System Information**: OS version, build number, architecture, uptime
- **Hardware Details**: CPU, RAM, disk space (C: drive), manufacturer, model, serial number
- **User & Sessions**: Logged-on user, login time
- **Connectivity**: IPv4 addresses, RDP port (3389) accessibility, WinRM port (5985) accessibility, ping status
- **Security**: Microsoft Defender status, real-time protection status, antivirus version
- **Patch Management**: Last installed hotfix ID and date, pending reboot signals
- **Software**: Count of installed applications
- **Health Assessment**: Automatic health status determination based on multiple factors

### Health Status Logic
The tool automatically evaluates device health into three categories:

| Status | Criteria |
|--------|----------|
| 🟢 **Healthy** | No critical issues, all systems nominal |
| 🟡 **Warning** | Low disk space (10-20%), pending updates, or minor issues |
| 🔴 **Needs Attention** | Critical disk space (<10%), no Defender protection, collection errors, or critical issues |

### Interactive HTML Report
- **Search**: Real-time search across all fields (Ctrl+F)
- **Filter**: By health status, reachability, or custom criteria
- **Sort**: Click column headers to sort ascending/descending
- **Detail Panel**: Click any row to see comprehensive device details
- **Export**: Generate CSV, JSON, or Excel files on-demand from filtered data
- **Copy**: Quick copy individual fields (hostname, IP, username) to clipboard
- **Print**: Print-friendly report format
- **Responsive Design**: Works on desktop and mobile browsers

## 📋 Requirements

- **PowerShell 5.1+** (Windows 7-11, Server 2008+)
- **Administrator privileges** on the execution machine (for CIM sessions)
- **Network access** to target computers
- **WinRM enabled** on remote computers (for advanced queries like Defender status)
  - Enable with: `Enable-PSRemoting -Force`
- **.NET Framework 4.0+** for Blob operations

## 🚀 Installation

1. **Clone or download** the repository:
   ```powershell
   git clone https://github.com/jlee2834/Argus-Inisght.git
   cd Argus-Inisght
   ```

2. **Set execution policy** (if needed):
   ```powershell
   Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
   ```

3. **Run the script**:
   ```powershell
   .\Argus.ps1
   ```

## 📖 Usage

### Basic Usage (Local Computer)
```powershell
.\Argus.ps1
```
Scans the local computer and opens the HTML report.

### Scan Multiple Computers
```powershell
.\Argus.ps1 -ComputerName "PC-001", "PC-002", "SERVER-01"
```

### Scan from File
```powershell
.\Argus.ps1 -InputFile "computers.txt"
```
The file should contain one computer name per line.

### Custom Output Directory
```powershell
.\Argus.ps1 -OutputDir "C:\Reports" -ComputerName "PC-001"
```

### Add Custom Logo
```powershell
.\Argus.ps1 -LogoUrl "https://company.com/logo.png"
```

### Skip Hotfix Collection (Faster)
```powershell
.\Argus.ps1 -SkipHotfixes
```

### Combine Parameters
```powershell
.\Argus.ps1 -ComputerName "PC-001", "PC-002" `
            -OutputDir ".\Reports" `
            -LogoUrl "https://imgur.com/logo.png" `
            -ThrottleDelayMs 300
```

## ⚙️ Parameters

| Parameter | Type | Default | Description |
|-----------|------|---------|-------------|
| `-ComputerName` | String[] | Local computer | One or more computer names to scan. Accepts pipeline input. |
| `-InputFile` | String | (none) | Path to text file with computer names (one per line). Combined with `-ComputerName`. |
| `-OutputDir` | String | `.\output` | Directory where HTML reports are saved. Created if doesn't exist. |
| `-LogoUrl` | String | (empty) | URL to company logo image (PNG, JPG). Displayed in report header. |
| `-ExportHtml` | Switch | $true | Generate HTML report (disable with `-ExportHtml:$false`). |
| `-OpenReport` | Switch | $true | Automatically open HTML in default browser. |
| `-SkipHotfixes` | Switch | $false | Skip querying hotfix information (speeds up collection). |
| `-ThrottleDelayMs` | Int | 200 | Delay in milliseconds between each computer scan (helps network management). |

## 💡 Usage Examples

### Example 1: Scan entire department
```powershell
$computers = @("DEPT-PC-01", "DEPT-PC-02", "DEPT-PC-03", "DEPT-SRV-01")
.\Argus.ps1 -ComputerName $computers -OutputDir "C:\IT\Reports" -LogoUrl "https://company.com/logo.png"
```

### Example 2: Scan from domain file with throttling
```powershell
.\Argus.ps1 -InputFile "C:\lists\domain_computers.txt" `
            -ThrottleDelayMs 500 `
            -OutputDir "C:\IT\Inventory"
```

### Example 3: Quick local scan without hotfixes
```powershell
.\Argus.ps1 -SkipHotfixes -OutputDir ".\temp"
```

### Example 4: Generate report but don't open it
```powershell
.\Argus.ps1 -ComputerName "SERVER-01" -OpenReport:$false
```

### Example 5: Pipeline multiple computers
```powershell
"PC-001", "PC-002", "PC-003" | ForEach-Object {
    Write-Host "Scanning $_..."
    .\Argus.ps1 -ComputerName $_ -ThrottleDelayMs 300
}
```

## 📊 Output Format

The script generates:
- **HTML Report** (interactive dashboard)
  - Saved as: `inventory_YYYYMMDD_HHMMSS.html`
  - Contains all data embedded in memory
  - No external dependencies needed

### Data Format in HTML Report

| Column | Description |
|--------|-------------|
| Computer | Device hostname |
| Health | Current status (Healthy/Warning/Needs Attention) |
| Issues | Summary of detected problems |
| Reachable | Ping response status |
| CIM | WinRM connectivity for CIM queries |
| Logged On User | Current user session |
| OS / Version / Build | Operating system details |
| Uptime | Days, hours, minutes since last boot |
| CPU / RAM | Processor and memory specifications |
| C: Disk Stats | Total, free, and free % of C: drive |
| IPv4 | Network addresses |
| Antivirus | Defender version and status |
| Hotfixes | Last patch installed and date |
| RDP / WinRM | Port accessibility status |
| Manufacturer / Model / Serial | Hardware identity |
| Error | Any collection errors encountered |

## 🎨 Interactive Report Features

### Search & Filter
- **Text Search**: Type in search box to filter across all fields
- **Status Filter**: Show only Healthy, Warning, or Needs Attention devices
- **Reachability Filter**: Show only reachable or unreachable devices
- **Sort Dropdown**: Quick sort by common metrics

### Detail Panel
- Click any device row to open detail panel on the right
- Shows comprehensive information organized by category:
  - Health Overview
  - Overview (device summary)
  - Hardware (specs & identity)
  - Security (protection status)
  - Patching (updates & reboots)
  - Connectivity (network & remote access)

### Export & Copy
- **Copy Visible Table**: Exports filtered results as CSV to clipboard
- **Export CSV**: Download filtered data as CSV file
- **Export JSON**: Download all data as JSON
- **Export Excel**: Download filtered data as Excel spreadsheet
- **Copy Hostname / IP / User**: Quick copy individual device fields
- **Copy JSON**: Export selected device as JSON
- **Print**: Print-friendly format

## 🔒 Security Considerations

- ✅ Script requires **local administrator** rights
- ✅ Uses **secure CIM sessions** (no plaintext)
- ✅ Supports **Kerberos & NTLM** authentication
- ⚠️ Remote queries require **WinRM enabled** on targets
- ⚠️ Network admin can monitor WinRM connections
- 💡 Run from **trusted network** only
- 💡 Secure reports containing sensitive data (IPs, usernames, software inventory)

## 🛠️ Troubleshooting

### "CIM session failed" error
**Solution**: Ensure WinRM is enabled on target computer:
```powershell
# On target computer:
Enable-PSRemoting -Force
```

### "Access denied" when running script
**Solution**: Run PowerShell as Administrator:
```powershell
# Right-click PowerShell → Run as Administrator
```

### "Execution policy" error
**Solution**: Change execution policy:
```powershell
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
```

### Hotfix "Unavailable" on some computers
**Solution**: Some systems don't allow remote hotfix queries. Use `-SkipHotfixes` flag or check manually.

### Defender status showing "Unavailable"
**Solution**: Ensure WinRM is enabled and user has permissions to query security info.

## 📝 Output Examples

```
Collecting from PC-001 ...
Collecting from PC-002 ...
Collecting from SERVER-01 ...

ComputerName OS                     Uptime       HealthStatus DiskC_FreePercent
------------ --                     ------       ------------ -----------
PC-001       Windows 10 Enterprise  45d 3h 22m   Healthy      68.5
PC-002       Windows 10 Pro         2d 14h 8m    Warning      18.3
SERVER-01    Windows Server 2019    128d 5h 47m  Needs Atten… 9.8

Generated HTML report in memory
Saved HTML: .\output\inventory_20260416_142530.html
Note: CSV/JSON/Excel files are generated on-demand via HTML buttons
```

## 🎓 Tips & Best Practices

1. **Regular Scans**: Schedule weekly reports for infrastructure visibility
2. **Logo Branding**: Add company logo for professional reports
3. **Bulk Reports**: Use input file for entire departments or geos
4. **Throttling**: Use `-ThrottleDelayMs 500` for large networks to reduce load
5. **Filtering**: Use HTML filters to focus on computers needing attention
6. **Export Data**: Export to JSON for further analysis or automation
7. **Archive Reports**: Save reports with timestamps for historical tracking

## 📄 License

None

## 👨‍💻 Author

Created by jlee2834

## 🤝 Contributing

Contributions welcome! Please feel free to submit pull requests.

## 📞 Support

For issues or feature requests, please open an issue on GitHub.
