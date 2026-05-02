<#
    .SYNOPSIS
    Exchange Environment Report - 3.0
    Author: B.O
    Modernized for Exchange 2016+ (SE Support)
    Performance: Utilizing lookup tables and bulk collections.
#>
param(
    [parameter(Position = 0, Mandatory = $true)][string]$HTMLReport,
    [parameter(Position = 1)][bool]$SendMail = $false,
    [parameter(Position = 2)][string]$MailFrom,
    [parameter(Position = 3)]$MailTo,
    [parameter(Position = 4)][string]$MailServer,
    [parameter(Position = 5)][string]$ServerFilter = "*",
    [string]$CompanyLogo = "EXCHANGE",
    [string]$ReportTitle = "REPORTING",
    [string]$ThemeColor = "#F27A00" # Default Color (ENSP Orange) - Customizable to #0078D4 (Blue) etc.
)
$Global:Sw = [System.Diagnostics.Stopwatch]::StartNew()
function Log($Msg, $Color = "White") { Write-Host "[$($Global:Sw.Elapsed.ToString("mm\:ss"))] $Msg" -ForegroundColor $Color -NoNewline:$false }

# --- INTERNAL FUNCTIONS ---

function _GetSSLCertStatus {
    param($ServerName)
    try {
        $Certs = Get-ExchangeCertificate -Server $ServerName -ErrorAction SilentlyContinue | Where-Object { $_.Services -match "IIS|SMTP" }
        if (!$Certs) { return @{ Status = "Unknown"; Color = "gray" } }
        $MinExpiry = $Certs | Sort-Object NotAfter | Select-Object -First 1
        $DaysLeft = ($MinExpiry.NotAfter - (Get-Date)).Days
        if ($DaysLeft -lt 0) { return @{ Status = "Expired!"; Color = "red" } }
        if ($DaysLeft -lt 30) { return @{ Status = "Expires in $DaysLeft d"; Color = "orange" } }
        return @{ Status = "OK ($DaysLeft d)"; Color = "green" }
    }
    catch { return @{ Status = "Error"; Color = "red" } }
}

function _GetDB {
    param($Database, $ExSvrData, $MailboxesByDB, $ArchivesByDB)
	
    $DbName = $Database.Name
    $DbIdentity = $Database.Identity.ToString()
    
    # Mailbox Counts from Lookup Tables (Super Fast)
    $MBCount = $(if ($MailboxesByDB.ContainsKey($DbIdentity)) { $MailboxesByDB[$DbIdentity].Count } else { 0 })
    $ArcCount = $(if ($ArchivesByDB.ContainsKey($DbName)) { $ArchivesByDB[$DbName].Count } else { 0 })
    
    # Average Sizes from pre-collected Server Stats
    $AvgMBSize = 0; $AvgArcSize = 0
    if ($ExSvrData.MBStatsByDB.ContainsKey($DbIdentity)) {
        $stats = $ExSvrData.MBStatsByDB[$DbIdentity]
        $total = 0; $stats | ForEach-Object { $total += $_.Size }; $AvgMBSize = $total / $stats.Count
    }
    if ($ExSvrData.ArcStatsByDB.ContainsKey($DbIdentity)) {
        $stats = $ExSvrData.ArcStatsByDB[$DbIdentity]
        $total = 0; $stats | ForEach-Object { $total += $_.Size }; $AvgArcSize = $total / $stats.Count
    }

    # Disk Space (CIM) - DB & Log
    $FreeDBDisk = $null; $FreeLogDisk = $null
    if ($ExSvrData.Disks) {
        foreach ($Disk in $ExSvrData.Disks) {
            if ($Database.EdbFilePath.PathName -like "$($Disk.Name)*") { $FreeDBDisk = $Disk.FreeSpace / $Disk.Capacity * 100 }
            if ($Database.LogFolderPath.PathName -like "$($Disk.Name)*") { $FreeLogDisk = $Disk.FreeSpace / $Disk.Capacity * 100 }
        }
    }

    @{Name = $DbName; ActiveOwner = $Database.Server.Name.ToUpper(); MailboxCount = $MBCount; MailboxAverageSize = $AvgMBSize; 
        ArchiveMailboxCount = $ArcCount; ArchiveAverageSize = $AvgArcSize; Size = $Database.DatabaseSize.ToBytes(); 
        Whitespace = $Database.AvailableNewMailboxSpace.ToBytes(); LastFullBackup = $(if ($Database.LastFullBackup) { $Database.LastFullBackup.ToString() }else { "None" });
        FreeDatabaseDiskSpace = $FreeDBDisk; FreeLogDiskSpace = $FreeLogDisk
    }
}

function _GetExSvr {
    param($Svr, $MailboxesByDB)
    Log "Collecting $($Svr.Name)..." "Gray"
    
    # ExSetup Version (Precise)
    $ExSetupVer = try { Invoke-Command -ComputerName $Svr.Name -ScriptBlock { (Get-Command "C:\Program Files\Microsoft\Exchange Server\V15\bin\ExSetup.exe").FileVersionInfo.FileVersion } -ErrorAction SilentlyContinue } catch { $null }
    
    # CIM Info
    $CimSession = New-CimSession -ComputerName $Svr.Name -SessionOption (New-CimSessionOption -Protocol Dcom) -ErrorAction SilentlyContinue
    if ($CimSession) {
        $OS = (Get-CimInstance Win32_OperatingSystem -CimSession $CimSession -ErrorAction SilentlyContinue).Caption.Replace("Microsoft ", "")
        $Disks = Get-CimInstance Win32_Volume -CimSession $CimSession -ErrorAction SilentlyContinue | Select-Object Name, Capacity, FreeSpace
        Remove-CimSession $CimSession
    }

    # Bulk Stats Collection (Fast)
    $MBStatsByDB = @{}; $ArcStatsByDB = @{}
    Get-MailboxStatistics -Server $Svr.Name -ErrorAction SilentlyContinue | ForEach-Object {
        if (!$MBStatsByDB[$_.Database.ToString()]) { $MBStatsByDB[$_.Database.ToString()] = New-Object System.Collections.Generic.List[PSObject] }
        $MBStatsByDB[$_.Database.ToString()].Add(@{Size = $_.TotalItemSize.Value.ToBytes() })
    }
    Get-MailboxStatistics -Server $Svr.Name -Archive -ErrorAction SilentlyContinue | ForEach-Object {
        if (!$ArcStatsByDB[$_.Database.ToString()]) { $ArcStatsByDB[$_.Database.ToString()] = New-Object System.Collections.Generic.List[PSObject] }
        $ArcStatsByDB[$_.Database.ToString()].Add(@{Size = $_.TotalItemSize.Value.ToBytes() })
    }

    $Roles = [array]($Svr.ServerRole.ToString().Split(",") | ForEach-Object { $_.Trim() } | Where-Object { $_ -match "Mailbox|Edge" })
    $MBTotal = 0; $Databases | Where-Object { $_.Server -eq $Svr.Name } | ForEach-Object { $MBTotal += $(if ($MailboxesByDB.ContainsKey($_.Identity.ToString())) { $MailboxesByDB[$_.Identity.ToString()].Count }else { 0 }) }

    Write-Host " [OK]" -ForegroundColor Green
    @{Name = $Svr.Name.ToUpper(); DisplayVer = $(if ($Svr.AdminDisplayVersion.Major -eq 15 -and $Svr.AdminDisplayVersion.Minor -eq 1) { "2016" }elseif ($Svr.AdminDisplayVersion.Minor -ge 2) { "2019 / SE" }else { "$($Svr.AdminDisplayVersion.Major).$($Svr.AdminDisplayVersion.Minor)" });
        Build = $(if ($ExSetupVer) { $ExSetupVer } else { $Svr.AdminDisplayVersion.ToString() }); Roles = $Roles; Mailboxes = $MBTotal; OSVersion = ($OS); Disks = $Disks;
        CertStatus = _GetSSLCertStatus -ServerName $Svr.Name; MBStatsByDB = $MBStatsByDB; ArcStatsByDB = $ArcStatsByDB; Site = $Svr.Site.Name 
    }
}

# --- PROCESS ---
$ExBin = "C:\Program Files\Microsoft\Exchange Server\V15\bin\RemoteExchange.ps1"
if (!(Get-Command Get-ExchangeServer -ErrorAction SilentlyContinue)) { if (Test-Path $ExBin) { . $ExBin; Connect-ExchangeServer -auto } else { throw "Launch from EMS" } }

Log "Global Collection (Optimized Single-Query V3.0)..." "Cyan"
$AllMbx = Get-Mailbox -ResultSize Unlimited | Select-Object Database, ArchiveDatabase, Identity
$MailboxesByDB = $AllMbx | Group-Object Database -AsHashTable -AsString
$ArchivesByDB = $AllMbx | Where-Object { $_.ArchiveDatabase } | Group-Object ArchiveDatabase -AsHashTable -AsString
$ExchangeServers = Get-ExchangeServer $ServerFilter
$Databases = Get-MailboxDatabase -Status | Where-Object { $_.Server -like $ServerFilter }

$EnvData = @{Sites = @{}; Servers = @{}; DBs = @() }
foreach ($S in $ExchangeServers) {
    $Ex = _GetExSvr -Svr $S -MailboxesByDB $MailboxesByDB
    if ($Ex.Site) { if (!$EnvData.Sites[$Ex.Site]) { $EnvData.Sites[$Ex.Site] = @($Ex) }else { $EnvData.Sites[$Ex.Site] += $Ex } }
    $EnvData.Servers[$Ex.Name] = $Ex
}
foreach ($D in $Databases) { $EnvData.DBs += _GetDB -Database $D -ExSvrData $EnvData.Servers[$D.Server.Name] -MailboxesByDB $MailboxesByDB -ArchivesByDB $ArchivesByDB }

# --- KPI CALCULATIONS V1.9.1 ---
$TotalMB = 0; $TotalArc = 0; $TotalSize = 0; $SvrOK = 0; $SvrTotal = $EnvData.Servers.Count
foreach ($S in $EnvData.Servers.Values) { if ($S.CertStatus.Status -like "*OK*") { $SvrOK++ } }
foreach ($D in $EnvData.DBs) { $TotalMB += $D.MailboxCount; $TotalArc += $D.ArchiveMailboxCount; $TotalSize += $D.Size }
$TotalSizeGB = "{0:N2}" -f ($TotalSize / 1GB)

# --- HTML GENERATION ---
$ReportDate = Get-Date -Format "dd/MM/yyyy HH:mm:ss"
$Output = @"
<!DOCTYPE html><html><head><title>Exchange Report 3.0</title>
<meta charset="UTF-8">
<style>
    body { font-family: 'Segoe UI', 'Roboto', Helvetica, Arial, sans-serif; background-color: #F5F5F5; margin: 0; padding: 20px; color: #333; }
    .header { text-align: center; padding: 40px 0 20px 0; background: transparent; color: #333; margin-bottom: 0; box-shadow: none; }
    .header h1 { margin: 0; font-weight: 300; font-size: 32px; color: #1A1A1A; }
    .header h1 span { color: $ThemeColor; font-weight: 600; }
    .header p { margin: 5px 0 0; color: #999; font-size: 13px; letter-spacing: 2px; text-transform: uppercase; }
    .container { background: white; padding: 30px; border-radius: 8px; box-shadow: 0 4px 15px rgba(0,0,0,0.05); width: 98%; margin: 0 auto; }
    h3 { color: #1A1A1A; border-bottom: 2px solid $ThemeColor; padding-bottom: 10px; margin-top: 30px; font-weight: 600; }
    table { width: 100%; border-collapse: collapse; margin-bottom: 25px; font-size: 14px; }
    th { cursor: pointer; background: #1A1A1A; color: white; padding: 12px 15px; font-weight: 500; text-align: center; border-top: 3px solid $ThemeColor; font-size: 13px; text-transform: uppercase; letter-spacing: 0.5px; }
    td { padding: 10px 15px; border-bottom: 1px solid #eee; text-align: center; color: #444; }
    tbody tr:nth-child(even) { background-color: #fafafa; }
    tbody tr:hover { background-color: #fff8f0; }
    .dashboard { display: flex; justify-content: space-between; margin-bottom: 25px; gap: 20px; flex-wrap: wrap; }
    .card { background: white; padding: 20px; border-radius: 4px; box-shadow: 0 2px 5px rgba(0,0,0,0.05); text-align: center; flex: 1; min-width: 150px; border-top: 3px solid $ThemeColor; }
    .card h2 { margin: 0; font-size: 32px; color: #1A1A1A; }
    .card p { margin: 5px 0 0; color: #666; font-size: 13px; text-transform: uppercase; font-weight: bold; }
    .progress-container { display: flex; align-items: center; gap: 8px; }
    .progress-bg { background: #eee; height: 12px; border-radius: 6px; flex: 1; overflow: hidden; }
    .progress-bar { height: 100%; border-radius: 6px; }
    .progress-text { font-weight: bold; font-size: 11px; min-width: 35px; text-align: right; }
    .footer { text-align: center; font-size: 12px; color: #999; margin-top: 40px; }
</style>
<script>
    function sortTable(tid, n, num) {
        var t = document.getElementById(tid), r = Array.from(t.rows).slice(1), dir = t.dataset.dir === 'asc' ? -1 : 1;
        r.sort((a, b) => {
            let v1 = a.cells[n].innerText, v2 = b.cells[n].innerText;
            if (num) { v1 = parseFloat(v1.replace(/[^\d.-]/g, '')) || 0; v2 = parseFloat(v2.replace(/[^\d.-]/g, '')) || 0; }
            return v1 > v2 ? dir : -dir;
        });
        r.forEach(row => t.tBodies[0].appendChild(row));
        t.dataset.dir = dir === 1 ? 'asc' : 'desc';
    }
</script>
</head>
<body>
<div class="header">
    <h1><span>$CompanyLogo</span> $ReportTitle</h1>
    <p>Exchange Infrastructure &bull; $ReportDate</p>
</div>
<div class="container">
    <div class="dashboard">
        <div class="card"><h2>$TotalMB</h2><p>Active Mailboxes</p></div>
        <div class="card"><h2>$TotalArc</h2><p>Archive Mailboxes</p></div>
        <div class="card"><h2>$TotalSizeGB <small style="font-size:16px;">GB</small></h2><p>Total Volume</p></div>
        <div class="card"><h2>$SvrOK / $SvrTotal</h2><p>Servers Online</p></div>
    </div>
"@

foreach ($Site in $EnvData.Sites.GetEnumerator()) {
    $tid = "t_" + $Site.Key.Replace(" ", "")
    $Output += "<h3>Site: $($Site.Key)</h3><table id='$tid'><thead><tr>
    <th onclick='sortTable(""$tid"",0,0)'>Server</th><th onclick='sortTable(""$tid"",1,0)'>Version</th><th onclick='sortTable(""$tid"",2,0)'>Build</th>
    <th onclick='sortTable(""$tid"",3,0)'>Roles</th><th onclick='sortTable(""$tid"",4,1)'>Mailboxes</th><th onclick='sortTable(""$tid"",5,0)'>Certificate</th>
    <th onclick='sortTable(""$tid"",6,0)'>OS</th></tr></thead><tbody>"
    foreach ($S in $Site.Value) {
        $Output += "<tr><td><b>$($S.Name)</b></td><td>$($S.DisplayVer)</td><td style='font-size:8pt;'>$($S.Build)</td><td>$($S.Roles -join ", ")</td>
        <td>$($S.Mailboxes)</td><td style='color:$($S.CertStatus.Color);font-weight:bold;'>$($S.CertStatus.Status)</td><td style='font-size:8pt;'>$($S.OSVersion)</td></tr>"
    }
    $Output += "</tbody></table>"
}

$Output += "<h3>Database Status</h3><table id='dbt'><thead><tr>
<th onclick='sortTable(""dbt"",0,0)'>Server</th><th onclick='sortTable(""dbt"",1,0)'>Database</th><th onclick='sortTable(""dbt"",2,1)'>Mailboxes</th>
<th onclick='sortTable(""dbt"",3,1)'>Avg. Size</th><th onclick='sortTable(""dbt"",4,1)'>Archives</th><th onclick='sortTable(""dbt"",5,1)'>Avg. Arc. Size</th>
<th onclick='sortTable(""dbt"",6,1)'>DB Size</th><th onclick='sortTable(""dbt"",7,1)'>Whitespace</th>
<th onclick='sortTable(""dbt"",8,1)'>Free DB</th><th onclick='sortTable(""dbt"",9,1)'>Free Log</th><th onclick='sortTable(""dbt"",10,0)'>Last Backup</th></tr></thead><tbody>"
foreach ($D in $EnvData.DBs) {
    $pctDB = $D.FreeDatabaseDiskSpace; $colDB = if ($pctDB -lt 10) { "#d32f2f" }elseif ($pctDB -lt 20) { "#ff9800" }else { "#2e7d32" }
    $pctLog = $D.FreeLogDiskSpace; $colLog = if ($pctLog -lt 10) { "#d32f2f" }elseif ($pctLog -lt 20) { "#ff9800" }else { "#2e7d32" }

    $Output += "<tr><td>$($D.ActiveOwner)</td><td align='left'>$($D.Name)</td><td>$($D.MailboxCount)</td>
    <td>$('{0:N2}' -f ($D.MailboxAverageSize/1GB)) GB</td><td>$($D.ArchiveMailboxCount)</td><td>$('{0:N2}' -f ($D.ArchiveAverageSize/1GB)) GB</td>
    <td style='font-weight:bold;'>$('{0:N2}' -f ($D.Size/1GB)) GB</td><td>$('{0:N2}' -f ($D.Whitespace/1GB)) GB</td>
    <td><div class='progress-container'><div class='progress-bg'><div class='progress-bar' style='width:$($pctDB)%;background:$colDB;'></div></div><div class='progress-text'>$('{0:N0}' -f $pctDB)%</div></div></td>
    <td><div class='progress-container'><div class='progress-bg'><div class='progress-bar' style='width:$($pctLog)%;background:$colLog;'></div></div><div class='progress-text'>$('{0:N0}' -f $pctLog)%</div></div></td>
    <td style='font-size:8pt;color:#666;'>$($D.LastFullBackup)</td></tr>"
}
$Output += "</tbody></table></div><div class='footer'>&copy; $(Get-Date -Format 'yyyy') $CompanyLogo - $ReportTitle</div></body></html>"
$Output | Out-File $HTMLReport -Encoding utf8
Log "Report 3.0 completed : $HTMLReport" "Green"


