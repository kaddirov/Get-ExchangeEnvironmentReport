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
        $Certs = Get-ExchangeCertificate -Server $ServerName -ErrorAction SilentlyContinue
        if (!$Certs) { return @{ Details = @() } }
        
        $AuthThumb = (Get-AuthConfig).CurrentCertificateThumbprint
        $StatusDetails = @()

        foreach ($cert in $Certs) {
            $pills = @()
            if ($cert.Services -match "IIS")  { $pills += @{ Letter = "I"; Name = "IIS";  Class = "pill-iis" } }
            if ($cert.Services -match "SMTP") { $pills += @{ Letter = "S"; Name = "SMTP"; Class = "pill-smtp" } }
            if ($cert.Services -match "POP")  { $pills += @{ Letter = "P"; Name = "POP";  Class = "pill-pop" } }
            if ($cert.Services -match "IMAP") { $pills += @{ Letter = "M"; Name = "IMAP"; Class = "pill-imap" } }
            if ($cert.Thumbprint -eq $AuthThumb) { $pills += @{ Letter = "A"; Name = "AUTH"; Class = "pill-auth" } }

            if ($pills.Count -eq 0) { continue }

            $Days = ($cert.NotAfter - (Get-Date)).Days
            $StatusColor = "green"
            if ($Days -lt 0) { $StatusColor = "red" }
            elseif ($Days -lt 30) { $StatusColor = "orange" }
            
            $CN = if ($cert.Subject -match "CN=([^,]+)") { $Matches[1] } else { $cert.Subject }
            
            $StatusDetails += @{ 
                Name = $CN
                Pills = $pills
                Days = $Days
                StatusColor = $StatusColor
                Expiry = $cert.NotAfter.ToString("dd/MM/yyyy")
                Issuer = $cert.Issuer.Replace("CN=", "")
            }
        }
        return @{ Details = $StatusDetails | Sort-Object Days }
    }
    catch { return @{ Details = @() } }
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
    @{Name = $Svr.Name.ToUpper(); DisplayVer = $(if ($Svr.AdminDisplayVersion.Major -eq 15 -and $Svr.AdminDisplayVersion.Minor -eq 1) { "2016" }elseif ($Svr.AdminDisplayVersion.Minor -ge 2) { if ($Svr.AdminDisplayVersion.Build -ge 2500) { "SE" } else { "2019" } } else { "$($Svr.AdminDisplayVersion.Major).$($Svr.AdminDisplayVersion.Minor)" });
        Build = $(if ($ExSetupVer) { $ExSetupVer } else { $Svr.AdminDisplayVersion.ToString() }); Roles = $Roles; Mailboxes = $MBTotal; OSVersion = ($OS); Disks = $Disks;
        CertStatus = _GetSSLCertStatus -ServerName $Svr.Name; MBStatsByDB = $MBStatsByDB; ArcStatsByDB = $ArcStatsByDB; Site = $Svr.Site.Name 
    }
}

# --- PROCESS ---
$ExBin = "C:\Program Files\Microsoft\Exchange Server\V15\bin\RemoteExchange.ps1"
if (!(Get-Command Get-ExchangeServer -ErrorAction SilentlyContinue)) { if (Test-Path $ExBin) { . $ExBin; Connect-ExchangeServer -auto } else { throw "Lancer depuis EMS" } }

Log "Collecte Globale (Requete unique optimisee V3.0)..." "Cyan"
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
foreach ($S in $EnvData.Servers.Values) { if ($S.OSVersion) { $SvrOK++ } }
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
    th { cursor: pointer; background: #1A1A1A; color: white; padding: 12px 8px; font-weight: 500; text-align: center; border-top: 3px solid #F27A00; font-size: 13px; text-transform: uppercase; letter-spacing: 0.5px; position: relative; transition: background 0.2s; }
    th:hover { background: #222; }
    .th-content { display: flex; align-items: center; justify-content: center; gap: 5px; margin-right: 15px; margin-left: 15px; }
    .sort-indicator { font-size: 10px; opacity: 0.7; }
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
    .filter-icon { 
        position: absolute; right: 6px; top: 50%; transform: translateY(-50%);
        cursor: pointer; color: rgba(255,255,255,0.5); padding: 4px; border-radius: 3px;
        transition: all 0.2s; display: inline-flex; align-items: center; justify-content: center;
        background: transparent; border: 1px solid transparent; z-index: 2;
    }
    .filter-icon:hover { background: rgba(255,255,255,0.15); color: #F27A00; }
    .filter-icon svg { width: 12px; height: 12px; display: block; }
    .filter-icon.filter-active { color: #F27A00 !important; background: rgba(242,122,0,0.2); border-color: rgba(242,122,0,0.4); }
    .filter-menu {
        position: fixed; background: white; color: #333; border: 1px solid #ccc; border-radius: 4px;
        box-shadow: 0 4px 15px rgba(0,0,0,0.15); padding: 5px 0; z-index: 9999;
        font-size: 13px; font-weight: normal; text-transform: none; letter-spacing: normal; min-width: 160px;
    }
    .filter-menu div { padding: 8px 15px; cursor: pointer; transition: background 0.2s; text-align: left; display: flex; align-items: center; }
    .filter-menu div:hover { background: #f0f7ff; }
    .filter-menu label { cursor: pointer; flex: 1; margin-left: 8px; white-space: nowrap; }
    .filter-menu input[type="checkbox"] { cursor: pointer; width: 14px; height: 14px; flex-shrink: 0; }
    .filter-menu hr { margin: 5px 0; border: 0; border-top: 1px solid #eee; padding: 0; display: block; cursor: default; }
    .filter-footer { padding: 8px 10px; text-align: right; background: #fafafa; border-top: 1px solid #eee; cursor: default; }
    .filter-btn { padding: 4px 12px; cursor: pointer; border-radius: 3px; border: 1px solid #ccc; background: white; font-size: 11px; transition: all 0.2s; }
    .filter-btn-primary { background: #F27A00; color: white; border-color: #d66c00; font-weight: bold; }
    .filter-btn:hover { background: #f0f0f0; }
    .filter-btn-primary:hover { background: #d66c00; }
    
    /* Styles Certificats Ultra-Compact */
    .cert-container { text-align: left; min-width: 350px; }
    .cert-item { display: flex; align-items: center; white-space: nowrap; margin-bottom: 3px; padding: 2px 0; border-bottom: 1px solid #f9f9f9; }
    .cert-item:last-child { border-bottom: none; }
    .cert-status-dot { width: 7px; height: 7px; border-radius: 50%; margin-right: 6px; flex-shrink: 0; }
    .cert-name-wrap { position: relative; cursor: help; display: flex; align-items: center; }
    .cert-name { font-weight: 600; font-size: 11px; color: #1A1A1A; margin-right: 8px; overflow: hidden; text-overflow: ellipsis; max-width: 200px; display: inline-block; }
    .cert-pills-wrap { display: flex; gap: 3px; margin-right: 8px; }
    .cert-pill { padding: 1px 6px; border-radius: 3px; font-size: 9px; font-weight: bold; color: white; line-height: 14px; height: 14px; text-transform: uppercase; }
    .pill-iis { background: #0078D4; }
    .pill-smtp { background: #2e7d32; }
    .pill-pop, .pill-imap { background: #666; }
    .pill-auth { background: #F27A00; }
    .cert-expiry-text { font-size: 9px; color: #888; margin-left: auto; padding-left: 10px; }
    
    .cert-name-wrap:hover .cert-tooltip { display: block; }
    .cert-tooltip { 
        display: none; position: absolute; bottom: 22px; left: 0;
        background: #1A1A1A; color: white; padding: 6px 10px; border-radius: 4px; font-size: 10px;
        white-space: nowrap; z-index: 2000; box-shadow: 0 3px 8px rgba(0,0,0,0.4); pointer-events: none; font-weight: normal;
    }
</style>
<script>
    /* === TRI === */
    function sortTable(tid, n, num) {
        var t = document.getElementById(tid);
        if (!t) return;
        var rows = Array.from(t.tBodies[0].rows);
        var dir = (t.dataset.sortCol == n && t.dataset.sortDir === 'asc') ? 'desc' : 'asc';
        var mult = dir === 'asc' ? 1 : -1;

        t.querySelectorAll('.sort-indicator').forEach(function(si) { si.textContent = ''; });

        rows.sort(function(a, b) {
            var v1 = a.cells[n] ? a.cells[n].innerText.trim() : '';
            var v2 = b.cells[n] ? b.cells[n].innerText.trim() : '';
            if (num) {
                v1 = parseFloat(v1.replace(/[^\d.,-]/g, '').replace(',', '.')) || 0;
                v2 = parseFloat(v2.replace(/[^\d.,-]/g, '').replace(',', '.')) || 0;
            }
            if (v1 < v2) return -1 * mult;
            if (v1 > v2) return 1 * mult;
            return 0;
        });
        rows.forEach(function(row) { t.tBodies[0].appendChild(row); });

        t.dataset.sortCol = n;
        t.dataset.sortDir = dir;
        var th = t.querySelectorAll('thead th')[n];
        if (th) {
            var ind = th.querySelector('.sort-indicator');
            if (ind) ind.innerHTML = dir === 'asc' ? ' &#9652;' : ' &#9662;';
        }
    }

    /* === FILTRES === */
    function initFilters(tableId) {
        var table = document.getElementById(tableId);
        if (!table || !table.querySelector('tbody')) return;
        var headers = table.querySelectorAll('thead th');
        if (headers.length === 0) return;

        table._origRows = Array.from(table.tBodies[0].rows);
        table._filters = {};

        headers.forEach(function(th, idx) {
            var colName = th.textContent.trim();
            var isNum = /Boites|Bo.tes|Taille|Espace|Libre|Archives/i.test(
                colName.normalize('NFD').replace(/[\u0300-\u036f]/g, '')
            );

            th.innerHTML = '<div class="th-content"><span>' + colName + '</span><span class="sort-indicator"></span></div>';

            var icon = document.createElement('span');
            icon.className = 'filter-icon';
            icon.title = 'Filtrer';
            icon.innerHTML = '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M22 3H2l8 9.46V19l4 2v-8.54L22 3z"/></svg>';
            th.appendChild(icon);

            icon.addEventListener('click', function(e) {
                e.stopPropagation();
                e.preventDefault();
                showFilterMenu(table, th, idx, icon);
            });

            th.addEventListener('click', function(e) {
                if (e.target.closest('.filter-icon')) return;
                sortTable(tableId, idx, isNum);
            });
        });
    }

    function showFilterMenu(table, th, colIndex, icon) {
        closeFilterMenu();

        var values = [];
        var seen = {};
        table._origRows.forEach(function(row) {
            if (!row.cells[colIndex]) return;
            var v = row.cells[colIndex].innerText.trim();
            if (!seen[v]) { seen[v] = true; values.push(v); }
        });
        values.sort();

        var menu = document.createElement('div');
        menu.className = 'filter-menu';
        menu.id = '_activeFilterMenu';

        var header = document.createElement('div');
        header.style.padding = '8px 12px';
        header.style.background = '#f9f9f9';
        header.style.cursor = 'default';
        var btnAll = document.createElement('button');
        btnAll.className = 'filter-btn';
        btnAll.textContent = 'Tout';
        var btnNone = document.createElement('button');
        btnNone.className = 'filter-btn';
        btnNone.textContent = 'Aucun';
        btnNone.style.marginLeft = '5px';
        header.appendChild(btnAll);
        header.appendChild(btnNone);
        menu.appendChild(header);

        var hr = document.createElement('hr');
        menu.appendChild(hr);

        var list = document.createElement('div');
        list.style.maxHeight = '200px';
        list.style.overflowY = 'auto';

        var currentF = table._filters[colIndex] || null;

        values.forEach(function(val) {
            var item = document.createElement('div');
            var cb = document.createElement('input');
            cb.type = 'checkbox';
            cb.value = val;
            cb.checked = !currentF || currentF.indexOf(val) !== -1;
            var lbl = document.createElement('label');
            lbl.textContent = val || '(Vide)';
            lbl.addEventListener('click', function(e) { e.preventDefault(); cb.checked = !cb.checked; });
            item.appendChild(cb);
            item.appendChild(lbl);
            list.appendChild(item);
        });
        menu.appendChild(list);

        var footer = document.createElement('div');
        footer.className = 'filter-footer';
        var applyBtn = document.createElement('button');
        applyBtn.className = 'filter-btn filter-btn-primary';
        applyBtn.textContent = 'Appliquer';
        footer.appendChild(applyBtn);
        menu.appendChild(footer);

        btnAll.addEventListener('click', function() { list.querySelectorAll('input').forEach(function(c) { c.checked = true; }); });
        btnNone.addEventListener('click', function() { list.querySelectorAll('input').forEach(function(c) { c.checked = false; }); });
        applyBtn.addEventListener('click', function() {
            var sel = [];
            list.querySelectorAll('input:checked').forEach(function(c) { sel.push(c.value); });
            if (sel.length === 0 || sel.length === values.length) {
                delete table._filters[colIndex];
                icon.classList.remove('filter-active');
            } else {
                table._filters[colIndex] = sel;
                icon.classList.add('filter-active');
            }
            applyFilters(table);
            closeFilterMenu();
        });

        document.body.appendChild(menu);
        var rect = th.getBoundingClientRect();
        menu.style.top = (rect.bottom + 2) + 'px';
        menu.style.left = Math.min(rect.left, window.innerWidth - menu.offsetWidth - 10) + 'px';

        setTimeout(function() {
            document._filterClose = function(e) {
                var m = document.getElementById('_activeFilterMenu');
                if (m && !m.contains(e.target) && !icon.contains(e.target)) closeFilterMenu();
            };
            document.addEventListener('click', document._filterClose);
        }, 10);
    }

    function closeFilterMenu() {
        var m = document.getElementById('_activeFilterMenu');
        if (m) m.remove();
        if (document._filterClose) { document.removeEventListener('click', document._filterClose); document._filterClose = null; }
    }

    function applyFilters(table) {
        var tbody = table.tBodies[0];
        while (tbody.firstChild) tbody.removeChild(tbody.firstChild);
        table._origRows.forEach(function(row) {
            var show = true;
            for (var cIdx in table._filters) {
                if (!table._filters.hasOwnProperty(cIdx)) continue;
                var cellText = row.cells[cIdx] ? row.cells[cIdx].innerText.trim() : '';
                if (table._filters[cIdx].indexOf(cellText) === -1) { show = false; break; }
            }
            if (show) tbody.appendChild(row);
        });
    }

    window.onload = function() {
        document.querySelectorAll('table[id]').forEach(function(t) { initFilters(t.id); });
    };
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
    <th>Serveur</th><th>Version</th><th>Build</th>
    <th>R&ocirc;les</th><th>Bo&icirc;tes</th><th>Certificat</th>
    <th>OS</th></tr></thead><tbody>"
    foreach ($S in $Site.Value) {
        $CertHTML = "<div class='cert-container'>"
        foreach ($c in $S.CertStatus.Details) {
            $CertHTML += "<div class='cert-item'>"
            $CertHTML += "<span class='cert-status-dot' style='background:$($c.StatusColor)'></span>"
            $CertHTML += "<span class='cert-name-wrap'><span class='cert-name'>$($c.Name)</span><div class='cert-tooltip'>Emetteur : $($c.Issuer)</div></span>"
            $CertHTML += "<div class='cert-pills-wrap'>"
            foreach ($pill in $c.Pills) {
                $CertHTML += "<span class='cert-pill $($pill.Class)'>$($pill.Name)</span>"
            }
            $CertHTML += "</div>"
            $CertHTML += "<span class='cert-expiry-text'>Exp: $($c.Expiry) ($($c.Days) j)</span>"
            $CertHTML += "</div>"
        }
        $CertHTML += "</div>"
        $Output += "<tr><td><b>$($S.Name)</b></td><td>$($S.DisplayVer)</td><td style='font-size:8pt;'>$($S.Build)</td><td>$($S.Roles -join "<br>")</td>
        <td>$($S.Mailboxes)</td><td>$CertHTML</td><td style='font-size:8pt;'>$($S.OSVersion)</td></tr>"
    }
    $Output += "</tbody></table>"
}

$Output += "<h3>&Eacute;tat des Bases de Donn&eacute;es</h3><table id='dbt'><thead><tr>
<th>Serveur</th><th>Base</th><th>Bo&icirc;tes</th>
<th>Taille Moy.</th><th>Archives</th><th>Taille Moy. Arc.</th>
<th>Taille DB</th><th>Espace Blanc</th>
<th>DB Libre</th><th>Log Libre</th><th>Dernier Backup</th></tr></thead><tbody>"
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
Log "Rapport V3.0 (ENSP Edition) termine : $HTMLReport" "Green"

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
        Log " - Config IIS mise a jour (Un court arret est normal)" "Yellow"
    }
    else {
        Log " - Config IIS deja optimale (Aucun impact)" "Green"
    }
}
catch {
    Log " - Erreur config IIS : $($_.Exception.Message)" "Red"
}
