<#
    .SYNOPSIS
    Exchange Environment Report - V3.0 (ENSP Edition)
    Modernized for Exchange 2016+ (SE Support)
    Performance: Utilisation de lookup tables et collectes groupées.
#>
param(
    [parameter(Position = 0, Mandatory = $true)][string]$HTMLReport,
    [parameter(Position = 1)][bool]$SendMail = $false,
    [parameter(Position = 2)][string]$MailFrom,
    [parameter(Position = 3)]$MailTo,
    [parameter(Position = 4)][string]$MailServer,
    [parameter(Position = 5)][string]$ServerFilter = "*"
)
$Global:Sw = [System.Diagnostics.Stopwatch]::StartNew()
function Log($Msg, $Color = "White") { Write-Host "[$($Global:Sw.Elapsed.ToString("mm\:ss"))] $Msg" -ForegroundColor $Color -NoNewline:$false }

# --- INTERNAL FUNCTIONS ---

function _GetSSLCertStatus {
    param($ServerName)
    try {
        $Certs = Get-ExchangeCertificate -Server $ServerName -ErrorAction SilentlyContinue | Where-Object { $_.Services -match "IIS|SMTP" }
        if (!$Certs) { return @{ Status = "Inconnu"; Color = "gray" } }
        $MinExpiry = $Certs | Sort-Object NotAfter | Select-Object -First 1
        $DaysLeft = ($MinExpiry.NotAfter - (Get-Date)).Days
        if ($DaysLeft -lt 0) { return @{ Status = "Expiré !"; Color = "red" } }
        if ($DaysLeft -lt 30) { return @{ Status = "Expire dans $DaysLeft j"; Color = "orange" } }
        return @{ Status = "OK ($DaysLeft j)"; Color = "green" }
    }
    catch { return @{ Status = "Erreur"; Color = "red" } }
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
        Whitespace = $Database.AvailableNewMailboxSpace.ToBytes(); LastFullBackup = $(if ($Database.LastFullBackup) { $Database.LastFullBackup.ToString() }else { "Aucune" });
        FreeDatabaseDiskSpace = $FreeDBDisk; FreeLogDiskSpace = $FreeLogDisk
    }
}

function _GetExSvr {
    param($Svr, $MailboxesByDB)
    Log "Collecte $($Svr.Name)..." "Gray"
    
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
if (!(Get-Command Get-ExchangeServer -ErrorAction SilentlyContinue)) { if (Test-Path $ExBin) { . $ExBin; Connect-ExchangeServer -auto } else { throw "Lancer depuis EMS" } }

Log "Collecte Globale (Requête unique optimisée V2.8)..." "Cyan"
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

# --- CALCULS KPI V1.9.1 ---
$TotalMB = 0; $TotalArc = 0; $TotalSize = 0; $SvrOK = 0; $SvrTotal = $EnvData.Servers.Count
foreach ($S in $EnvData.Servers.Values) { if ($S.CertStatus.Status -like "*OK*") { $SvrOK++ } }
foreach ($D in $EnvData.DBs) { $TotalMB += $D.MailboxCount; $TotalArc += $D.ArchiveMailboxCount; $TotalSize += $D.Size }
$TotalSizeGB = "{0:N2}" -f ($TotalSize / 1GB)

# --- HTML GENERATION ---
$ReportDate = Get-Date -Format "dd/MM/yyyy HH:mm:ss"
$Output = @"
<!DOCTYPE html><html><head><title>Exchange Report V3.0 - ENSP</title>
<meta charset="UTF-8">
<style>
    body { font-family: 'Segoe UI', 'Roboto', Helvetica, Arial, sans-serif; background-color: #F5F5F5; margin: 0; padding: 20px; color: #333; }
    .header { text-align: center; padding: 40px 0 20px 0; background: transparent; color: #333; margin-bottom: 0; box-shadow: none; }
    .header h1 { margin: 0; font-weight: 300; font-size: 32px; color: #1A1A1A; }
    .header h1 span { color: #F27A00; font-weight: 600; }
    .header p { margin: 5px 0 0; color: #999; font-size: 13px; letter-spacing: 2px; text-transform: uppercase; }
    .container { background: white; padding: 30px; border-radius: 8px; box-shadow: 0 4px 15px rgba(0,0,0,0.05); width: 98%; margin: 0 auto; }
    h3 { color: #1A1A1A; border-bottom: 2px solid #F27A00; padding-bottom: 10px; margin-top: 30px; font-weight: 600; }
    table { width: 100%; border-collapse: collapse; margin-bottom: 25px; font-size: 14px; }
    th { cursor: pointer; background: #1A1A1A; color: white; padding: 12px 15px; font-weight: 500; text-align: center; border-top: 3px solid #F27A00; font-size: 13px; text-transform: uppercase; letter-spacing: 0.5px; }
    td { padding: 10px 15px; border-bottom: 1px solid #eee; text-align: center; color: #444; }
    tbody tr:nth-child(even) { background-color: #fafafa; }
    tbody tr:hover { background-color: #fff8f0; }
    .dashboard { display: flex; justify-content: space-between; margin-bottom: 25px; gap: 20px; flex-wrap: wrap; }
    .card { background: white; padding: 20px; border-radius: 4px; box-shadow: 0 2px 5px rgba(0,0,0,0.05); text-align: center; flex: 1; min-width: 150px; border-top: 3px solid #F27A00; }
    .card h2 { margin: 0; font-size: 32px; color: #1A1A1A; }
    .card p { margin: 5px 0 0; color: #666; font-size: 13px; text-transform: uppercase; font-weight: bold; }
    .progress-container { display: flex; align-items: center; gap: 8px; }
    .progress-bg { background: #eee; height: 12px; border-radius: 6px; flex: 1; overflow: hidden; }
    .progress-bar { height: 100%; border-radius: 6px; }
    .progress-text { font-weight: bold; font-size: 11px; min-width: 35px; text-align: right; }
    .footer { text-align: center; font-size: 12px; color: #999; margin-top: 40px; }
    
    /* Styles du Filtre */
    .filter-icon { cursor: pointer; color: #888; margin-right: 8px; font-size: 11px; transition: color 0.2s; vertical-align: middle; }
    .filter-icon:hover { color: #F27A00; }
    .filter-active { color: #F27A00 !important; font-weight: bold; }
    .filter-menu {
        position: absolute; background: white; color: #333; border: 1px solid #ccc; border-radius: 4px;
        box-shadow: 0 4px 15px rgba(0,0,0,0.15); padding: 5px 0; z-index: 1000;
        max-height: 250px; overflow-y: auto; font-size: 13px; font-weight: normal; text-transform: none; min-width: 150px;
    }
    .filter-menu div { padding: 8px 15px; cursor: pointer; transition: background 0.2s; text-align: left; display: flex; align-items: center; }
    .filter-menu div:hover { background: #f0f7ff; }
    .filter-menu label { cursor: pointer; flex: 1; margin-left: 8px; }
    .filter-menu input[type="checkbox"] { cursor: pointer; width: 14px; height: 14px; }
    .filter-menu hr { margin: 5px 0; border: 0; border-top: 1px solid #eee; }
    .filter-footer { padding: 10px; text-align: right; background: #fafafa; border-top: 1px solid #eee; }
    .filter-btn { padding: 4px 12px; cursor: pointer; border-radius: 3px; border: 1px solid #ccc; background: white; font-size: 11px; transition: all 0.2s; }
    .filter-btn-primary { background: #F27A00; color: white; border-color: #d66c00; font-weight: bold; }
    .filter-btn:hover { background: #f0f0f0; }
    .filter-btn-primary:hover { background: #d66c00; }
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

    function initFilters(tableId) {
        const table = document.getElementById(tableId);
        if (!table) return;
        const headers = table.querySelectorAll('th');
        const tbody = table.querySelector('tbody');
        table.originalRows = Array.from(tbody.querySelectorAll('tr'));
        table.activeFilters = {};

        headers.forEach((th, index) => {
            const icon = document.createElement('span');
            icon.className = 'filter-icon';
            icon.innerHTML = '\u25BC';
            icon.onclick = function(e) { e.stopPropagation(); showFilterMenu(table, th, index, icon); };
            th.insertBefore(icon, th.firstChild);
        });
    }

    function showFilterMenu(table, th, colIndex, icon) {
        let existing = document.querySelector('.filter-menu');
        if (existing) existing.remove();
        
        const values = new Set();
        table.originalRows.forEach(row => { if(row.cells[colIndex]) values.add(row.cells[colIndex].innerText.trim()); });
        const sortedValues = Array.from(values).sort();

        const menu = document.createElement('div');
        menu.className = 'filter-menu';

        // Header : Tout cocher / Décocher
        const header = document.createElement('div');
        header.style.padding = '8px 15px';
        header.style.background = '#f9f9f9';
        header.innerHTML = '<button class="filter-btn" id="btn-all">Tout cocher</button>' +
                           '<button class="filter-btn" id="btn-none" style="margin-left:5px;">Tout d&eacute;cocher</button>';
        menu.appendChild(header);
        menu.appendChild(document.createElement('hr'));

        // Liste des valeurs avec Checkboxes
        const listContainer = document.createElement('div');
        listContainer.style.maxHeight = '180px';
        listContainer.style.overflowY = 'auto';
        
        const currentFilters = table.activeFilters[colIndex] || [];

        sortedValues.forEach(val => {
            const item = document.createElement('div');
            const cb = document.createElement('input');
            cb.type = 'checkbox';
            cb.value = val;
            cb.checked = currentFilters.length === 0 || currentFilters.includes(val);
            
            const lbl = document.createElement('label');
            lbl.innerText = val || '(Vide)';
            lbl.onclick = (e) => { e.preventDefault(); cb.checked = !cb.checked; };
            
            item.appendChild(cb);
            item.appendChild(lbl);
            listContainer.appendChild(item);
        });
        menu.appendChild(listContainer);

        // Footer : Appliquer
        const footer = document.createElement('div');
        footer.className = 'filter-footer';
        const applyBtn = document.createElement('button');
        applyBtn.className = 'filter-btn filter-btn-primary';
        applyBtn.innerText = 'Appliquer';
        applyBtn.onclick = () => {
            const selected = Array.from(listContainer.querySelectorAll('input:checked')).map(c => c.value);
            const finalSelection = (selected.length === sortedValues.length) ? null : selected;
            applyFilter(table, colIndex, finalSelection, menu, icon);
        };
        footer.appendChild(applyBtn);
        menu.appendChild(footer);

        header.querySelector('#btn-all').onclick = () => listContainer.querySelectorAll('input').forEach(c => c.checked = true);
        header.querySelector('#btn-none').onclick = () => listContainer.querySelectorAll('input').forEach(c => c.checked = false);

        document.body.appendChild(menu);
        const rect = th.getBoundingClientRect();
        menu.style.top = (rect.bottom + window.scrollY) + 'px';
        menu.style.left = (rect.left + window.scrollX) + 'px';
        
        setTimeout(() => { 
            document.onclick = function(e) { 
                if (!menu.contains(e.target) && !icon.contains(e.target)) { menu.remove(); document.onclick = null; } 
            }; 
        }, 0);
    }

    function applyFilter(table, colIndex, values, menu, icon) {
        if (values === null || values.length === 0) { 
            delete table.activeFilters[colIndex]; 
            icon.classList.remove('filter-active'); 
        } else { 
            table.activeFilters[colIndex] = values; 
            icon.classList.add('filter-active'); 
        }
        
        const tbody = table.querySelector('tbody');
        tbody.innerHTML = '';
        table.originalRows.forEach(row => {
            let show = true;
            for (const [cIdx, filterArray] of Object.entries(table.activeFilters)) {
                if (!filterArray.includes(row.cells[cIdx].innerText.trim())) { show = false; break; }
            }
            if (show) tbody.appendChild(row);
        });
        menu.remove();
    }
    
    window.onload = function() { initFilters('dbt'); };
</script>
</head>
<body>
<div class="header">
    <h1><span>ENSP</span> REPORTING</h1>
    <p>Infrastructure Exchange &bull; $ReportDate</p>
</div>
<div class="container">
    <div class="dashboard">
        <div class="card"><h2>$TotalMB</h2><p>Bo&icirc;tes Actives</p></div>
        <div class="card"><h2>$TotalArc</h2><p>Bo&icirc;tes Archives</p></div>
        <div class="card"><h2>$TotalSizeGB <small style="font-size:16px;">GB</small></h2><p>Volum&eacute;trie Totale</p></div>
        <div class="card"><h2>$SvrOK / $SvrTotal</h2><p>Serveurs En Ligne</p></div>
    </div>
"@

foreach ($Site in $EnvData.Sites.GetEnumerator()) {
    $tid = "t_" + $Site.Key.Replace(" ", "")
    $Output += "<h3>Site: $($Site.Key)</h3><table id='$tid'><thead><tr>
    <th onclick='sortTable(""$tid"",0,0)'>Serveur</th><th onclick='sortTable(""$tid"",1,0)'>Version</th><th onclick='sortTable(""$tid"",2,0)'>Build</th>
    <th onclick='sortTable(""$tid"",3,0)'>R&ocirc;les</th><th onclick='sortTable(""$tid"",4,1)'>Bo&icirc;tes</th><th onclick='sortTable(""$tid"",5,0)'>Certificat</th>
    <th onclick='sortTable(""$tid"",6,0)'>OS</th></tr></thead><tbody>"
    foreach ($S in $Site.Value) {
        $Output += "<tr><td><b>$($S.Name)</b></td><td>$($S.DisplayVer)</td><td style='font-size:8pt;'>$($S.Build)</td><td>$($S.Roles -join ", ")</td>
        <td>$($S.Mailboxes)</td><td style='color:$($S.CertStatus.Color);font-weight:bold;'>$($S.CertStatus.Status)</td><td style='font-size:8pt;'>$($S.OSVersion)</td></tr>"
    }
    $Output += "</tbody></table>"
}

$Output += "<h3>&Eacute;tat des Bases de Donn&eacute;es</h3><table id='dbt'><thead><tr>
<th onclick='sortTable(""dbt"",0,0)'>Serveur</th><th onclick='sortTable(""dbt"",1,0)'>Base</th><th onclick='sortTable(""dbt"",2,1)'>Bo&icirc;tes</th>
<th onclick='sortTable(""dbt"",3,1)'>Taille Moy.</th><th onclick='sortTable(""dbt"",4,1)'>Archives</th><th onclick='sortTable(""dbt"",5,1)'>Taille Moy. Arc.</th>
<th onclick='sortTable(""dbt"",6,1)'>Taille DB</th><th onclick='sortTable(""dbt"",7,1)'>Espace Blanc</th>
<th onclick='sortTable(""dbt"",8,1)'>DB Libre</th><th onclick='sortTable(""dbt"",9,1)'>Log Libre</th><th onclick='sortTable(""dbt"",10,0)'>Dernier Backup</th></tr></thead><tbody>"
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
$Output += "</tbody></table></div><div class='footer'>&copy; 2026 ENSP - Exchange Reporting System</div></body></html>"
$Output | Out-File $HTMLReport -Encoding utf8
Log "Rapport V3.0 (ENSP Edition) terminé : $HTMLReport" "Green"

# --- CONFIGURATION (IIS Default Document) ---
try {
    $ReportDir = [System.IO.Path]::GetDirectoryName($HTMLReport)
    $ReportFile = [System.IO.Path]::GetFileName($HTMLReport)
    $WebConfigPath = Join-Path $ReportDir "web.config"
    
    # Configuration "Page par défaut" pour accéder via le dossier
    $WebConfigContent = @"
<?xml version="1.0" encoding="UTF-8"?>
<configuration>
    <system.webServer>
        <defaultDocument enabled="true">
            <files>
                <clear />
                <add value="$ReportFile" />
            </files>
        </defaultDocument>
    </system.webServer>
</configuration>
"@
    $CurrentConfig = if (Test-Path $WebConfigPath) { Get-Content $WebConfigPath -Raw -ErrorAction SilentlyContinue } else { "" }
    # Normalisation pour comparaison (suppression retours chariots)
    if ($CurrentConfig.Trim() -ne $WebConfigContent.Trim()) {
        $WebConfigContent | Out-File $WebConfigPath -Encoding utf8
        Log " - Config IIS mise à jour (Un court arrêt est normal)" "Yellow"
    }
    else {
        Log " - Config IIS déjà optimale (Aucun impact)" "Green"
    }
}
catch {
    Log " - Erreur config IIS : $($_.Exception.Message)" "Red"
}
