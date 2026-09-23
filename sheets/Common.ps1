<#
Common.ps1 v2 - Fonctions partagées pour mvp-reporting
Chargé par dot-source depuis le lanceur principal.
Expose: logging, proxy, HTTP, encodage, dates, cache, export, helpers métier,
        maps teams, workload scheme days, calcul heures attendues prorata.
#>

# ======================================================================
# LOGGING
# ======================================================================
function Write-Log {
    param([Parameter(Mandatory)][string]$Message,
          [ValidateSet("INFO","WARN","ERROR","DEBUG")][string]$Level = "INFO")
    $ts = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    try { Add-Content -Path $script:logFile -Value "[$ts] [$Level] $Message" -ErrorAction Stop } catch {}
}
function Write-Info($msg) {
    Write-Host "[INFO] $msg"
    if ($msg) { Write-Log $msg "INFO" }
}
function Write-Warn($msg) {
    Write-Warning $msg
    Write-Log $msg "WARN"
}
function Write-ErrLog($msg) {
    Write-Error $msg
    Write-Log $msg "ERROR"
}

# ======================================================================
# PROXY & TLS
# ======================================================================
function Initialize-Proxy {
    param([switch]$UseSystemProxy, [string]$ProxyUrl)
    try {
        # Support combiné TLS 1.2 et TLS 1.3
        [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls12 -bor [System.Net.SecurityProtocolType]::Tls13
        
        if ($ProxyUrl) {
            [System.Net.WebRequest]::DefaultWebProxy = New-Object System.Net.WebProxy($ProxyUrl, $true)
            Write-Info "Proxy: $ProxyUrl"
        } elseif ($UseSystemProxy) {
            [System.Net.WebRequest]::DefaultWebProxy = [System.Net.WebRequest]::GetSystemWebProxy()
            Write-Info "Proxy: système"
        } else {
            [System.Net.WebRequest]::DefaultWebProxy = $null
            Write-Info "Proxy: désactivé"
        }
    } catch { Write-Warn "Init proxy: $($_.Exception.Message)" }
}

function Get-EffectiveProxyUri {
    param([Parameter(Mandatory)][string]$TargetUrl)
    if ($script:ProxyUrl -and $script:ProxyUrl.Trim()) { return $script:ProxyUrl }
    if (-not $script:UseSystemProxy) { return $null }
    
    $dest = $null
    if (-not [System.Uri]::TryCreate($TargetUrl, [System.UriKind]::Absolute, [ref]$dest)) {
        return $null
    }
    
    $wp = [System.Net.WebRequest]::DefaultWebProxy
    if (-not $wp -or $wp.IsBypassed($dest)) { return $null }
    $proxy = $wp.GetProxy($dest)
    if (-not $proxy -or $proxy.AbsoluteUri -eq $dest.AbsoluteUri) { return $null }
    return $proxy.AbsoluteUri
}

# ======================================================================
# HTTP WRAPPERS (Fiabilisés PowerShell 5.1 / ISE)
# ======================================================================
function Get-WebExceptionBody([System.Net.WebException]$ex) {
    try {
        if (-not $ex.Response) { return $null }
        $s = $ex.Response.GetResponseStream()
        $r = New-Object System.IO.StreamReader($s)
        $b = $r.ReadToEnd(); $r.Dispose(); $s.Dispose()
        return $b
    } catch { return $null }
}

function Invoke-ApiGet {
    param([Parameter(Mandatory)][string]$Url,
          [Parameter(Mandatory)][hashtable]$Headers)
    Write-Log "GET $Url" "DEBUG"

    if (-not $Headers.ContainsKey("User-Agent")) {
        $Headers["User-Agent"] = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
    }
    if (-not $Headers.ContainsKey("Accept")) {
        $Headers["Accept"] = "application/json"
    }

    $params = @{
        Method      = 'GET'
        Uri         = $Url
        Headers     = $Headers
        ContentType = 'application/json'
        ErrorAction = 'Stop'
        TimeoutSec  = 30
    }

    $px = Get-EffectiveProxyUri -TargetUrl $Url
    if ($px) {
        $params.Proxy = $px
        $params.ProxyUseDefaultCredentials = $true
    }

    try { 
        return Invoke-RestMethod @params 
    }
    catch [System.Net.WebException] {
        $body = Get-WebExceptionBody $_.Exception
        # Log technique dans le fichier .log sans faire crier la console en rouge
        if ($body) { Write-Log "GET $Url : $($_.Exception.Message) | Body: $body" "WARN" }
        else       { Write-Log "GET $Url : $($_.Exception.Message)" "WARN" }
        throw
    }
}


function Invoke-AssetsAqlPost {
    param([Parameter(Mandatory)][string]$Url,
          [Parameter(Mandatory)][hashtable]$Headers,
          [Parameter(Mandatory)][string]$AqlQuery)
    Write-Log "POST Assets $Url | AQL: $AqlQuery" "DEBUG"
    
    if (-not $Headers.ContainsKey("User-Agent")) {
        $Headers["User-Agent"] = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
    }

    $jsonBody = '{"qlQuery": "' + $AqlQuery + '"}'
    $bytes = [System.Text.Encoding]::UTF8.GetBytes($jsonBody)
    $params = @{
        Method      = 'POST'
        Uri         = $Url
        Headers     = $Headers
        Body        = $bytes
        ContentType = 'application/json; charset=utf-8'
        ErrorAction = 'Stop'
        TimeoutSec  = 30
    }

    $px = Get-EffectiveProxyUri -TargetUrl $Url
    if ($px) {
        $params.Proxy = $px
        $params.ProxyUseDefaultCredentials = $true
    }

    try { return Invoke-RestMethod @params }
    catch [System.Net.WebException] {
        $body = Get-WebExceptionBody $_.Exception
        if ($body) { Write-ErrLog "POST $Url : $($_.Exception.Message)`nBody:`n$body" }
        else       { Write-ErrLog "POST $Url : $($_.Exception.Message)" }
        throw
    }
}

# ======================================================================
# ENCODAGE UTF-8
# ======================================================================
function Fix-DoubleUtf8([string]$s) {
    if ([string]::IsNullOrEmpty($s)) { return $s }
    try {
        $latin1 = [System.Text.Encoding]::GetEncoding("iso-8859-1")
        $bytes  = $latin1.GetBytes($s)
        return [System.Text.Encoding]::UTF8.GetString($bytes)
    } catch { return $s }
}

# ======================================================================
# DATES
# ======================================================================
function Normalize-AssetDate([string]$raw) {
    if ([string]::IsNullOrWhiteSpace($raw)) { return "" }
    $formats = @(
        "dd-MMM-yy","dd/MMM/yy","dd-MMM-yyyy","dd/MMM/yyyy",
        "dd-MMMM-yy","dd-MMMM-yyyy",
        "dd/MM/yyyy","dd-MM-yyyy","yyyy-MM-dd",
        "dd/MM/yy","dd-MM-yy","M/d/yyyy","MM/dd/yyyy"
    )
    $cultures = @(
        [System.Globalization.CultureInfo]::new("fr-FR"),
        [System.Globalization.CultureInfo]::new("en-US"),
        [System.Globalization.CultureInfo]::InvariantCulture
    )
    $variants = @($raw, ($raw -replace '\.', ''))
    foreach ($variant in $variants) {
        foreach ($culture in $cultures) {
            foreach ($fmt in $formats) {
                $parsed = [datetime]::MinValue
                if ([datetime]::TryParseExact($variant, $fmt, $culture,
                    [System.Globalization.DateTimeStyles]::None, [ref]$parsed)) {
                    return $parsed.ToString("dd/MM/yyyy")
                }
            }
        }
    }
    foreach ($variant in $variants) {
        foreach ($culture in $cultures) {
            $parsed = [datetime]::MinValue
            if ([datetime]::TryParse($variant, $culture,
                [System.Globalization.DateTimeStyles]::None, [ref]$parsed)) {
                return $parsed.ToString("dd/MM/yyyy")
            }
        }
    }
    Write-Log "Date Assets non parsée: '$raw'" "WARN"
    return $raw
}

function Get-JoursFeries([int]$Annee) {
    $fixes = @(
        (Get-Date -Year $Annee -Month 1 -Day 1),
        (Get-Date -Year $Annee -Month 5 -Day 1),
        (Get-Date -Year $Annee -Month 5 -Day 8),
        (Get-Date -Year $Annee -Month 7 -Day 14),
        (Get-Date -Year $Annee -Month 8 -Day 15),
        (Get-Date -Year $Annee -Month 11 -Day 1),
        (Get-Date -Year $Annee -Month 11 -Day 11),
        (Get-Date -Year $Annee -Month 12 -Day 25)
    )
    $a = $Annee % 19
    $b = [math]::Floor($Annee / 100); $c = $Annee % 100
    $dd = [math]::Floor($b / 4); $e = $b % 4
    $f = [math]::Floor(($b + 8) / 25)
    $g = [math]::Floor(($b - $f + 1) / 3)
    $h = (19 * $a + $b - $dd - $g + 15) % 30
    $i = [math]::Floor($c / 4); $k = $c % 4
    $l = (32 + 2 * $e + 2 * $i - $h - $k) % 7
    $m = [math]::Floor(($a + 11 * $h + 22 * $l) / 451)
    $mois = [math]::Floor(($h + $l - 7 * $m + 114) / 31)
    $jour = (($h + $l - 7 * $m + 114) % 31) + 1
    $paques = Get-Date -Year $Annee -Month $mois -Day $jour
    $lundiPaques    = $paques.AddDays(1)
    $ascension      = $paques.AddDays(39)
    $lundiPentecote = $paques.AddDays(50)
    $all = $fixes + @($lundiPaques, $ascension, $lundiPentecote)
    return ($all | ForEach-Object { $_.Date })
}

function Get-MonthWeekBuckets([datetime]$MonthStart, [datetime]$MonthEnd) {
    $buckets = @()
    $current = $MonthStart.Date
    for ($w = 0; $w -lt 5; $w++) {
        if ($current -gt $MonthEnd) {
            $buckets += @{ Start = $null; End = $null }
            continue
        }
        $weekStart = $current
        if ($w -eq 4) {
            $weekEnd = $MonthEnd
        } else {
            if ($current.DayOfWeek -eq [DayOfWeek]::Sunday) {
                $weekEnd = $current
            } else {
                $daysToSunday = 7 - [int]$current.DayOfWeek
                $weekEnd = $current.AddDays($daysToSunday)
            }
            if ($weekEnd -gt $MonthEnd) { $weekEnd = $MonthEnd }
        }
        $buckets += @{ Start = $weekStart; End = $weekEnd }
        $current = $weekEnd.AddDays(1)
    }
    return $buckets
}

function Get-WeekIndexForDate([datetime]$Date, [array]$Buckets) {
    for ($i = 0; $i -lt $Buckets.Count; $i++) {
        if ($null -eq $Buckets[$i].Start) { continue }
        if ($Date.Date -ge $Buckets[$i].Start.Date -and $Date.Date -le $Buckets[$i].End.Date) {
            return $i
        }
    }
    return -1
}

# ======================================================================
# CACHE
# ======================================================================
function Save-Json([string]$Path, $Object) {
    $Object | ConvertTo-Json -Depth 80 | Set-Content -Path $Path -Encoding UTF8
}

function Load-Json([string]$Path) {
    if (-not (Test-Path $Path)) { return $null }
    try { return (Get-Content -Path $Path -Raw -Encoding UTF8 | ConvertFrom-Json) }
    catch { return $null }
}

# ======================================================================
# EXPORT CSV
# ======================================================================
function Export-StrictCsv {
    param(
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)][string[]]$Headers,
        [Parameter(Mandatory)]$Rows
    )
    $orderedRows = foreach ($r in $Rows) {
        $o = [ordered]@{}
        foreach ($h in $Headers) { $o[$h] = $r.$h }
        [pscustomobject]$o
    }
    $orderedRows | Export-Csv -NoTypeInformation -Encoding UTF8 -Delimiter ';' -Path $Path
}

# ======================================================================
# HELPERS MÉTIER
# ======================================================================
function Extract-EmailFromCompteJira([string]$c) {
    if ($c -match "\(([^)]+)\)") { return $Matches[1] }
    return $c
}

function Get-DirectionFromGroups([string[]]$Groups) {
    $directions = @()
    foreach ($g in $Groups) {
        if ($g -match "(?i)dsim"        -and $directions -notcontains "DSIM")        { $directions += "DSIM" }
        if ($g -match "(?i)dsit"        -and $directions -notcontains "DSIT")        { $directions += "DSIT" }
        if ($g -match "(?i)gouvernance" -and $directions -notcontains "Gouvernance") { $directions += "Gouvernance" }
    }
    if ($directions.Count -eq 0) { return "Autre" }
    return ($directions -join " / ")
}

function Transform-Programme([string]$p) {
    if ($p -eq "DSIM_TRANSVERSE") { return "DSIM" }
    if ($p -match "_") { return $p.Substring($p.IndexOf("_") + 1) }
    return $p
}

function New-BasicAuthHeader([string]$User, [string]$Pass) {
    $b64 = [Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes("$User`:$Pass"))
    return @{ Authorization = "Basic $b64"; Accept = "application/json" }
}

function ConvertFrom-SecureStringToPlain([securestring]$Secure) {
    $bstr = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($Secure)
    try { [Runtime.InteropServices.Marshal]::PtrToStringBSTR($bstr) }
    finally { [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($bstr) }
}

# ======================================================================
# XLSX
# ======================================================================
function ConvertTo-ExcelSafeSheetName([string]$Name) {
    $n = $Name -replace '[:\\\/\?\*\[\]]', ' '
    $n = $n.Trim()
    if ($n.Length -eq 0) { $n = "Sheet" }
    if ($n.Length -gt 31) { $n = $n.Substring(0, 31) }
    return $n
}

function Import-CsvToWorksheet($Worksheet, [string]$CsvPath) {
    $qt = $Worksheet.QueryTables.Add("TEXT;$CsvPath", $Worksheet.Range("A1"))
    $qt.TextFileParseType = 1
    $qt.TextFilePlatform = 65001
    $qt.TextFileSemicolonDelimiter = $true
    $qt.TextFileCommaDelimiter = $false
    $qt.TextFileConsecutiveDelimiter = $false
    $qt.AdjustColumnWidth = $true
    $qt.Refresh($false) | Out-Null
    $qt.Delete() | Out-Null
}

function New-ExcelWorkbookFromCsvSheets {
    param(
        [Parameter(Mandatory)][hashtable]$SheetToCsvPath,
        [Parameter(Mandatory)][string]$XlsxPath
    )
    $excel = $null; $wb = $null
    try {
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false; $excel.DisplayAlerts = $false
        $wb = $excel.Workbooks.Add()
        $targets = New-Object System.Collections.Generic.List[object]
        foreach ($k in $SheetToCsvPath.Keys) {
            $p = $SheetToCsvPath[$k]
            if (Test-Path $p) { $targets.Add([pscustomobject]@{ Name = $k; Path = $p }) | Out-Null }
            else { Write-Warn "XLSX: CSV introuvable: $p" }
        }
        if ($targets.Count -eq 0) {
            $wb.Worksheets.Item(1).Name = "Empty"
            $wb.Worksheets.Item(1).Range("A1").Value2 = "Aucun CSV."
            $wb.SaveAs($XlsxPath); return
        }
        $wsFirst = $wb.Worksheets.Item(1)
        $wsFirst.Name = (ConvertTo-ExcelSafeSheetName $targets[0].Name)
        Import-CsvToWorksheet -Worksheet $wsFirst -CsvPath $targets[0].Path
        for ($i = 1; $i -lt $targets.Count; $i++) {
            $ws = $wb.Worksheets.Add([System.Reflection.Missing]::Value,
                $wb.Worksheets.Item($wb.Worksheets.Count))
            $ws.Name = (ConvertTo-ExcelSafeSheetName $targets[$i].Name)
            Import-CsvToWorksheet -Worksheet $ws -CsvPath $targets[$i].Path
        }
        while ($wb.Worksheets.Count -gt $targets.Count) {
            $wb.Worksheets.Item($wb.Worksheets.Count).Delete() | Out-Null
        }
        $wb.Worksheets.Item(1).Activate()
        $wb.SaveAs($XlsxPath)
        Write-Info "XLSX généré: $XlsxPath"
    } finally {
        try { if ($wb) { $wb.Close($true) | Out-Null } } catch {}
        try { if ($excel) { $excel.Quit() | Out-Null } } catch {}
        try { if ($excel) { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null } } catch {}
        [GC]::Collect(); [GC]::WaitForPendingFinalizers()
    }
}

# ======================================================================
# MAPS TEAMS → MEMBRES
# ======================================================================
function Map-MembersToTeamsList($teamsData) {
    $map = @{}
    foreach ($t in $teamsData) {
        $members = @()
        if ($t."Membres ID") { $members = $t."Membres ID" -split ",\s*" }
        foreach ($id in $members) {
            if (-not $map.ContainsKey($id)) {
                $map[$id] = New-Object System.Collections.Generic.List[string]
            }
            $map[$id].Add($t.Nom)
        }
    }
    return $map
}

function Map-MembersToTeamExitDates($teamsData) {
    $map = @{}
    foreach ($t in $teamsData) {
        $memberDates = $t.MemberDates
        if (-not $memberDates) { continue }
        $members = @()
        if ($t."Membres ID") { $members = $t."Membres ID" -split ",\s*" }
        foreach ($mid in $members) {
            $dateTo = ""
            if ($memberDates -is [hashtable] -and $memberDates.ContainsKey($mid)) {
                $dateTo = $memberDates[$mid]
            } elseif ($memberDates.PSObject -and $memberDates.PSObject.Properties[$mid]) {
                $dateTo = [string]$memberDates.PSObject.Properties[$mid].Value
            }
            $label = if ([string]::IsNullOrWhiteSpace($dateTo)) {
                "{0}: (active)" -f $t.Nom
            } else {
                "{0}: {1}" -f $t.Nom, $dateTo
            }
            if (-not $map.ContainsKey($mid)) {
                $map[$mid] = New-Object System.Collections.Generic.List[string]
            }
            $map[$mid].Add($label)
        }
    }
    return $map
}

# ======================================================================
# WORKLOAD SCHEME DAYS (v2)
# ======================================================================
function Load-WorkloadSchemeDays {
    param(
        [Parameter(Mandatory)][hashtable]$TempoHeaders,
        [Parameter(Mandatory)][string]$CacheFilePath,
        [Parameter(Mandatory)][bool]$UseCache
    )
    $userDaysMap = @{}

    if ($UseCache) {
        $cached = Load-Json -Path $CacheFilePath
        if ($cached) {
            foreach ($p in $cached.PSObject.Properties) {
                $days = @{}
                foreach ($dp in $p.Value.PSObject.Properties) {
                    $days[$dp.Name] = [double]$dp.Value
                }
                $userDaysMap[$p.Name] = $days
            }
            Write-Info "Workload scheme days cache: $($userDaysMap.Count) users"
            return $userDaysMap
        }
    }

    $url = "https://api.tempo.io/4/workload-schemes"
    while ($true) {
        $resp = Invoke-ApiGet -Url $url -Headers $TempoHeaders
        foreach ($ws in ($resp.results | ForEach-Object { $_ })) {
            $schemeId   = [string]$ws.id
            $schemeName = Fix-DoubleUtf8 ([string]$ws.name)

            $days = @{
                monday=8.0; tuesday=8.0; wednesday=8.0
                thursday=8.0; friday=8.0; saturday=0.0; sunday=0.0
            }

            try {
                $detail = Invoke-ApiGet `
                    -Url "https://api.tempo.io/4/workload-schemes/$schemeId" `
                    -Headers $TempoHeaders

                if ($detail.days) {
                    foreach ($dayName in @("monday","tuesday","wednesday","thursday","friday","saturday","sunday")) {
                        $dayProp = $detail.days.$dayName
                        if ($null -eq $dayProp) {
                            $dayProp = $detail.days.($dayName.Substring(0,1).ToUpper() + $dayName.Substring(1))
                        }
                        if ($null -eq $dayProp) {
                            $dayProp = $detail.days.($dayName.ToUpper())
                        }
                        if ($null -ne $dayProp) {
                            if ($null -ne $dayProp.requiredHours) {
                                $days[$dayName] = [double]$dayProp.requiredHours
                            } elseif ($null -ne $dayProp.requiredSeconds) {
                                $days[$dayName] = [double]$dayProp.requiredSeconds / 3600.0
                            }
                        }
                    }
                }

                Write-Log ("WS days '{0}' (id={1}): L={2} M={3} Me={4} J={5} V={6} S={7} D={8}" -f `
                    $schemeName, $schemeId,
                    $days.monday, $days.tuesday, $days.wednesday,
                    $days.thursday, $days.friday, $days.saturday, $days.sunday) "DEBUG"
            } catch {
                Write-Warn "Workload scheme detail $schemeId ($schemeName): $($_.Exception.Message)"
            }

            try {
                $mResp = Invoke-ApiGet `
                    -Url "https://api.tempo.io/4/workload-schemes/$schemeId/members" `
                    -Headers $TempoHeaders
                foreach ($m in ($mResp.results | ForEach-Object { $_ })) {
                    $mid = $null
                    if ($m.member -and $m.member.accountId) { $mid = [string]$m.member.accountId }
                    elseif ($m.accountId) { $mid = [string]$m.accountId }
                    if ($mid) { $userDaysMap[$mid] = $days }
                }
            } catch {
                Write-Warn "Workload scheme members $schemeId : $($_.Exception.Message)"
            }
        }
        if ($resp.metadata -and $resp.metadata.next) { $url = $resp.metadata.next }
        else { break }
    }

    Save-Json -Path $CacheFilePath -Object $userDaysMap
    Write-Info "Workload scheme days: $($userDaysMap.Count) users"
    return $userDaysMap
}

# ======================================================================
# CALCUL HEURES ATTENDUES PRORATA (v2 — centralisé)
# ======================================================================
function Get-ExpectedHoursFromWorkload {
    param(
        [Parameter(Mandatory)][string]$AccountId,
        [Parameter(Mandatory)][hashtable]$UserWorkloadDays,
        [Parameter(Mandatory)][hashtable]$AssetByAccountId,
        $MemberDates,
        [Parameter(Mandatory)][datetime]$PeriodeFrom,
        [Parameter(Mandatory)][datetime]$PeriodeTo,
        [Parameter(Mandatory)]$JoursFeries
    )

    $effFrom = $PeriodeFrom
    if ($AssetByAccountId.ContainsKey($AccountId)) {
        $deStr = [string]$AssetByAccountId[$AccountId]."Date Entrée"
        if (-not [string]::IsNullOrWhiteSpace($deStr)) {
            $parsed = [datetime]::MinValue
            if ([datetime]::TryParseExact($deStr, "dd/MM/yyyy",
                [System.Globalization.CultureInfo]::InvariantCulture,
                [System.Globalization.DateTimeStyles]::None, [ref]$parsed)) {
                if ($parsed.Date -gt $effFrom.Date) { $effFrom = $parsed.Date }
            }
            elseif ([datetime]::TryParseExact($deStr, "yyyy-MM-dd",
                [System.Globalization.CultureInfo]::InvariantCulture,
                [System.Globalization.DateTimeStyles]::None, [ref]$parsed)) {
                if ($parsed.Date -gt $effFrom.Date) { $effFrom = $parsed.Date }
            }
        }
    }

    $effTo = $PeriodeTo
    if ($null -ne $MemberDates) {
        $dStr = ""
        if ($MemberDates -is [hashtable] -and $MemberDates.ContainsKey($AccountId)) {
            $dStr = $MemberDates[$AccountId]
        } elseif ($MemberDates.PSObject -and $MemberDates.PSObject.Properties[$AccountId]) {
            $dStr = [string]$MemberDates.PSObject.Properties[$AccountId].Value
        }
        if (-not [string]::IsNullOrWhiteSpace($dStr)) {
            $parsed = [datetime]::MinValue
            if ([datetime]::TryParse($dStr, [ref]$parsed)) {
                if ($parsed.Date -lt $effTo.Date) { $effTo = $parsed.Date }
            }
        }
    }

    if ($AssetByAccountId.ContainsKey($AccountId)) {
        $dsStr = [string]$AssetByAccountId[$AccountId]."Date Sortie"
        if (-not [string]::IsNullOrWhiteSpace($dsStr)) {
            $parsed = [datetime]::MinValue
            if ([datetime]::TryParseExact($dsStr, "dd/MM/yyyy",
                [System.Globalization.CultureInfo]::InvariantCulture,
                [System.Globalization.DateTimeStyles]::None, [ref]$parsed)) {
                if ($parsed.Date -lt $effTo.Date) { $effTo = $parsed.Date }
            }
        }
    }

    if ($effFrom.Date -gt $effTo.Date) { return -1 }

    $days = @{
        monday=8.0; tuesday=8.0; wednesday=8.0
        thursday=8.0; friday=8.0; saturday=0.0; sunday=0.0
    }
    if ($UserWorkloadDays.ContainsKey($AccountId)) {
        $days = $UserWorkloadDays[$AccountId]
    }

    $dayNameMap = @{
        [DayOfWeek]::Monday    = "monday"
        [DayOfWeek]::Tuesday   = "tuesday"
        [DayOfWeek]::Wednesday = "wednesday"
        [DayOfWeek]::Thursday  = "thursday"
        [DayOfWeek]::Friday    = "friday"
        [DayOfWeek]::Saturday  = "saturday"
        [DayOfWeek]::Sunday    = "sunday"
    }

    $totalHours = 0.0
    $d = $effFrom.Date
    while ($d -le $effTo.Date) {
        if ($JoursFeries -notcontains $d.Date) {
            $dayKey = $dayNameMap[$d.DayOfWeek]
            if ($days.ContainsKey($dayKey)) {
                $totalHours += [double]$days[$dayKey]
            }
        }
        $d = $d.AddDays(1)
    }

    return $totalHours
}