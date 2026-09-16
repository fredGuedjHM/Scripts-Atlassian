<#
Audit-VolumeHistoriqueAssets.ps1
Mesure le volume d'historique (audit trail) dans JSM Assets mois par mois.
Identifie les 3 plus vieux mois comme candidats a la purge.
Mode DRY RUN uniquement : aucune suppression n'est effectuee.
#>

[CmdletBinding()]
param(
    [switch]$UseSystemProxy = $true,
    [switch]$ProxyUseDefaultCredentials = $true,
    [string]$ProxyUrl,
    [string]$AssetsWorkspaceId = "2898703d-5f14-4527-9f8a-ad32b9279ca1",
    [string[]]$ObjectTypeIds   = @("68", "69"),
    [int]$MaxObjectsPerType    = 0  # 0 = pas de limite
)

# ============================================================
# 0. DOSSIERS
# ============================================================
$scriptDir  = if ($PSCommandPath) { Split-Path -Parent $PSCommandPath } else { (Get-Location).Path }
$secretsDir = Join-Path $scriptDir "secrets"
$cacheDir   = Join-Path $scriptDir "cache"
$logsDir    = Join-Path $scriptDir "logs"
$exportsDir = Join-Path $scriptDir "exports"

function Ensure-Dir([string]$Path) {
    if (-not (Test-Path $Path)) { New-Item -ItemType Directory -Path $Path -Force | Out-Null }
}
Ensure-Dir $secretsDir; Ensure-Dir $cacheDir; Ensure-Dir $logsDir; Ensure-Dir $exportsDir

$runStamp   = Get-Date -Format "yyyyMMdd_HHmmss"
$scriptName = "Audit-VolumeHistoriqueAssets"

# ============================================================
# 1. LOG
# ============================================================
$logFile = Join-Path $logsDir ($scriptName + "_" + $runStamp + ".log")

function Write-Log {
    param([string]$Message = "",
          [ValidateSet("INFO","WARN","ERROR","DEBUG")][string]$Level = "INFO")
    if ([string]::IsNullOrWhiteSpace($Message)) { return }
    $ts = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    try { Add-Content -Path $logFile -Value ("[$ts] [$Level] " + $Message) -ErrorAction Stop } catch {}
}
function Write-Info($msg)   { Write-Host ("[INFO] " + $msg);  Write-Log $msg "INFO"  }
function Write-Warn($msg)   { Write-Warning $msg;              Write-Log $msg "WARN"  }
function Write-ErrLog($msg) { Write-Error $msg;                Write-Log $msg "ERROR" }

Write-Log ("=== DEBUT " + $scriptName + " ===") "INFO"
Write-Info ("Script dir  : " + $scriptDir)
Write-Info ("Exports dir : " + $exportsDir)
Write-Info ("Logs dir    : " + $logsDir)

# ============================================================
# 2. PROXY
# ============================================================
function Initialize-Proxy {
    param([switch]$UseSystemProxy, [string]$ProxyUrl)
    try {
        [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
        if ($ProxyUrl) {
            [System.Net.WebRequest]::DefaultWebProxy = New-Object System.Net.WebProxy($ProxyUrl, $true)
            Write-Info ("Proxy: " + $ProxyUrl)
        } elseif ($UseSystemProxy) {
            [System.Net.WebRequest]::DefaultWebProxy = [System.Net.WebRequest]::GetSystemWebProxy()
            [System.Net.WebRequest]::DefaultWebProxy.Credentials = [System.Net.CredentialCache]::DefaultCredentials
            Write-Info "Proxy: systeme"
        } else {
            [System.Net.WebRequest]::DefaultWebProxy = $null
            Write-Info "Proxy: desactive"
        }
    } catch { Write-Warn ("Init proxy: " + $_.Exception.Message) }
}
Initialize-Proxy -UseSystemProxy:$UseSystemProxy -ProxyUrl $ProxyUrl

function Get-EffectiveProxyUri {
    param([Parameter(Mandatory)][string]$TargetUrl)
    if ($ProxyUrl -and $ProxyUrl.Trim()) { return $ProxyUrl }
    if (-not $UseSystemProxy) { return $null }
    try { $dest = [uri]$TargetUrl } catch { return $null }
    $wp = [System.Net.WebRequest]::DefaultWebProxy
    if (-not $wp -or $wp.IsBypassed($dest)) { return $null }
    $proxy = $wp.GetProxy($dest)
    if (-not $proxy -or $proxy.AbsoluteUri -eq $dest.AbsoluteUri) { return $null }
    return $proxy.AbsoluteUri
}

# ============================================================
# 3. HTTP WRAPPERS
# ============================================================
function Get-WebExceptionBody([System.Net.WebException]$ex) {
    try {
        if (-not $ex.Response) { return $null }
        $s = $ex.Response.GetResponseStream()
        $r = New-Object System.IO.StreamReader($s)
        $b = $r.ReadToEnd(); $r.Dispose(); $s.Dispose(); return $b
    } catch { return $null }
}

function Invoke-ApiGet {
    param([Parameter(Mandatory)][string]$Url,
          [Parameter(Mandatory)][hashtable]$Headers)
    $params = @{ Method='GET'; Uri=$Url; Headers=$Headers; ContentType='application/json'; ErrorAction='Stop' }
    $px = Get-EffectiveProxyUri -TargetUrl $Url
    if ($px) { $params.Proxy=$px; if ($ProxyUseDefaultCredentials) { $params.ProxyUseDefaultCredentials=$true } }
    try { return Invoke-RestMethod @params }
    catch [System.Net.WebException] {
        $sc = $null
        try { $sc = $_.Exception.Response.StatusCode.value__ } catch {}
        if ($sc -ne 404) {
            $body = Get-WebExceptionBody $_.Exception
            Write-ErrLog ("GET " + $Url + " : " + $_.Exception.Message + "`n" + $body)
        }
        throw
    }
}

function Invoke-ApiGetSafe {
    param([Parameter(Mandatory)][string]$Url,
          [Parameter(Mandatory)][hashtable]$Headers)
    try { return Invoke-ApiGet -Url $Url -Headers $Headers }
    catch { return $null }
}

function Invoke-ApiPostUtf8 {
    param([Parameter(Mandatory)][string]$Url,
          [Parameter(Mandatory)][hashtable]$Headers,
          [Parameter(Mandatory)][string]$JsonBody)
    $params = @{
        Method          = 'POST'
        Uri             = $Url
        Headers         = $Headers
        Body            = [System.Text.Encoding]::UTF8.GetBytes($JsonBody)
        ContentType     = "application/json"
        UseBasicParsing = $true
        ErrorAction     = 'Stop'
    }
    $px = Get-EffectiveProxyUri -TargetUrl $Url
    if ($px) { $params.Proxy=$px; if ($ProxyUseDefaultCredentials) { $params.ProxyUseDefaultCredentials=$true } }
    try {
        $resp   = Invoke-WebRequest @params
        $stream = $resp.RawContentStream
        $stream.Position = 0
        $reader = New-Object System.IO.StreamReader($stream, [System.Text.Encoding]::UTF8, $true)
        $raw    = $reader.ReadToEnd()
        $reader.Close()
        return $raw | ConvertFrom-Json
    }
    catch [System.Net.WebException] {
        $body = Get-WebExceptionBody $_.Exception
        Write-ErrLog ("POST " + $Url + " : " + $_.Exception.Message + "`n" + $body)
        throw
    }
}

# ============================================================
# 4. CREDENTIALS JIRA
# ============================================================
$jiraCredFile = Join-Path $secretsDir "jira-jiradot.cred.xml"
if (-not (Test-Path $jiraCredFile)) {
    Write-ErrLog ("Fichier Jira creds introuvable: " + $jiraCredFile)
    throw "Lance d'abord Save-JiraCredential.ps1"
}
$jiraData    = Import-Clixml -Path $jiraCredFile
$jiraBaseUrl = ($jiraData.JiraBaseUrl).TrimEnd("/")
$jiraCred    = $jiraData.Credential
$jiraEmail   = $jiraCred.UserName
$jiraToken   = $jiraCred.GetNetworkCredential().Password

$b64         = [Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes($jiraEmail + ":" + $jiraToken))
$jiraHeaders = @{ Authorization = "Basic " + $b64; Accept = "application/json" }
Write-Info ("Jira: " + $jiraBaseUrl + " (user=" + $jiraEmail + ")")

# ============================================================
# 5. VARIABLES ASSETS
# ============================================================
$assetsBaseUrl   = $jiraBaseUrl + "/gateway/api/jsm/assets/workspace/" + $AssetsWorkspaceId + "/v1"
$assetsAqlUrl    = $assetsBaseUrl + "/object/aql"
$assetsHistUrl   = $assetsBaseUrl + "/object"  # + /{id}/history
Write-Info ("Assets workspace : " + $AssetsWorkspaceId)
Write-Info ("Object Types     : " + ($ObjectTypeIds -join ", "))

# ============================================================
# 6. RECUPERATION DES OBJETS ASSETS
# ============================================================
Write-Info "=== Recuperation des objets Assets ==="

$allObjects = New-Object System.Collections.Generic.List[object]

foreach ($otId in $ObjectTypeIds) {
    $aqlQuery   = "objectTypeId = " + $otId
    $startAt    = 0
    $maxResults = 50
    $isLast     = $false
    $otCount    = 0

    Write-Info ("  AQL : " + $aqlQuery)

    while (-not $isLast) {
        $url      = $assetsAqlUrl + "?startAt=" + $startAt + "&maxResults=" + $maxResults + "&includeAttributes=false"
        $bodyObj  = @{ qlQuery = $aqlQuery }
        $bodyJson = $bodyObj | ConvertTo-Json -Depth 3

        try {
            $resp = Invoke-ApiPostUtf8 -Url $url -Headers $jiraHeaders -JsonBody $bodyJson
        } catch {
            Write-ErrLog ("Erreur AQL OT=" + $otId + " startAt=" + $startAt + " : " + $_.Exception.Message)
            break
        }
        if (-not $resp) { break }

        $objects = $resp.values
        if (-not $objects -or $objects.Count -eq 0) { break }

        foreach ($obj in $objects) {
            if (-not $obj.id) { continue }
            $allObjects.Add([pscustomobject]@{
                ObjectId   = [string]$obj.id
                ObjectKey  = if ($obj.objectKey) { [string]$obj.objectKey } else { "" }
                Label      = if ($obj.label)     { [string]$obj.label }     else { "" }
                ObjectType = $otId
                Created    = if ($obj.created)   { [string]$obj.created }   else { "" }
                Updated    = if ($obj.updated)   { [string]$obj.updated }   else { "" }
            }) | Out-Null
            $otCount++
        }

        $isLast   = if ($null -ne $resp.isLast) { [bool]$resp.isLast } else { $true }
        $startAt += $maxResults

        if ($startAt -gt 50000) { Write-Warn "Pagination anormale, arret."; break }
        Start-Sleep -Milliseconds 100
    }

    Write-Info ("  OT " + $otId + " : " + $otCount + " objets")

    if ($MaxObjectsPerType -gt 0 -and $otCount -gt $MaxObjectsPerType) {
        Write-Warn ("  Limite MaxObjectsPerType=" + $MaxObjectsPerType + " — troncature")
    }
}

Write-Info ("  Total objets recuperes : " + $allObjects.Count)

# Appliquer limite si definie (pour tests)
if ($MaxObjectsPerType -gt 0) {
    $limited = New-Object System.Collections.Generic.List[object]
    $countByOt = @{}
    foreach ($obj in $allObjects) {
        $ot = $obj.ObjectType
        if (-not $countByOt.ContainsKey($ot)) { $countByOt[$ot] = 0 }
        if ($countByOt[$ot] -lt $MaxObjectsPerType) {
            $limited.Add($obj) | Out-Null
            $countByOt[$ot]++
        }
    }
    $allObjects = $limited
    Write-Info ("  Apres limitation : " + $allObjects.Count + " objets")
}
# ============================================================
# 7. COLLECTE HISTORIQUE OBJET PAR OBJET
# ============================================================
Write-Info "=== Collecte historique (audit trail) par objet ==="

# Structure : mois (YYYY-MM) -> compteur d'entrees + liste d'objets concernes
$historyByMonth    = @{}  # "YYYY-MM" -> int (nb entrees)
$objectsByMonth    = @{}  # "YYYY-MM" -> HashSet[objectId]
$entriesByObjMonth = New-Object System.Collections.Generic.List[object]  # detail pour dry-run

$totalHistEntries = 0
$objProcessed     = 0
$objWithHistory   = 0
$objErrors        = 0

foreach ($obj in $allObjects) {
    $objId  = $obj.ObjectId
    $objKey = $obj.ObjectKey
    $objProcessed++

    if ($objProcessed % 50 -eq 0) {
        Write-Info ("  Progression : " + $objProcessed + "/" + $allObjects.Count + " objets (" + $totalHistEntries + " entrees historique)")
    }

    # Recuperer l'historique de cet objet (pagine)
    $histEntries = New-Object System.Collections.Generic.List[object]
    $histPage    = 1
    $histDone    = $false

    while (-not $histDone) {
        $histUrl = $assetsHistUrl + "/" + $objId + "/history?page=" + $histPage
        $histResp = Invoke-ApiGetSafe -Url $histUrl -Headers $jiraHeaders

        if (-not $histResp) {
            # Essayer format alternatif (asc=true)
            $histUrl2 = $assetsHistUrl + "/" + $objId + "/history?asc=true&abbreviate=true"
            $histResp = Invoke-ApiGetSafe -Url $histUrl2 -Headers $jiraHeaders
            if (-not $histResp) {
                $objErrors++
                break
            }
        }

        # L'API peut retourner un array directement ou un objet avec .values/.entries
        $entries = $null
        if     ($histResp -is [array])    { $entries = $histResp }
        elseif ($histResp.values)         { $entries = $histResp.values }
        elseif ($histResp.entries)        { $entries = $histResp.entries }
        elseif ($histResp.histories)      { $entries = $histResp.histories }
        else                              { $entries = @($histResp) }

        if (-not $entries -or $entries.Count -eq 0) { break }

        foreach ($entry in $entries) {
            $histEntries.Add($entry) | Out-Null
        }

        # Pagination : si moins de 25 resultats, probablement la derniere page
        if ($entries.Count -lt 25) { $histDone = $true }
        else { $histPage++ }

        if ($histPage -gt 100) { Write-Warn ("  Historique trop long pour " + $objKey + ", arret page 100"); break }
        Start-Sleep -Milliseconds 50
    }

    if ($histEntries.Count -eq 0) { continue }
    $objWithHistory++

    # Ventiler par mois
    foreach ($entry in $histEntries) {
        # Chercher la date dans les champs possibles
        $dateStr = ""
        if     ($entry.created)       { $dateStr = [string]$entry.created }
        elseif ($entry.createdDate)   { $dateStr = [string]$entry.createdDate }
        elseif ($entry.updatedDate)   { $dateStr = [string]$entry.updatedDate }
        elseif ($entry.when)          { $dateStr = [string]$entry.when }
        elseif ($entry.date)          { $dateStr = [string]$entry.date }

        if ([string]::IsNullOrWhiteSpace($dateStr)) { continue }

        # Extraire YYYY-MM (les dates Assets sont au format ISO ou dd/MM/yyyy)
        $month = ""
        try {
            if ($dateStr -match "^(\d{4})-(\d{2})") {
                $month = $Matches[1] + "-" + $Matches[2]
            } elseif ($dateStr -match "^(\d{2})/(\d{2})/(\d{4})") {
                $month = $Matches[3] + "-" + $Matches[2]
            } elseif ($dateStr -match "(\d{4})-(\d{2})-(\d{2})") {
                $month = $Matches[1] + "-" + $Matches[2]
            }
        } catch {}

        if ([string]::IsNullOrWhiteSpace($month)) { continue }

        # Compteur par mois
        if (-not $historyByMonth.ContainsKey($month)) { $historyByMonth[$month] = 0 }
        $historyByMonth[$month]++

        # Objets par mois
        if (-not $objectsByMonth.ContainsKey($month)) { $objectsByMonth[$month] = @{} }
        $objectsByMonth[$month][$objId] = $true

        $totalHistEntries++
    }

    # Stocker le resume par objet pour le detail dry-run
    $monthsForObj = @{}
    foreach ($entry in $histEntries) {
        $dateStr = ""
        if     ($entry.created)       { $dateStr = [string]$entry.created }
        elseif ($entry.createdDate)   { $dateStr = [string]$entry.createdDate }
        elseif ($entry.updatedDate)   { $dateStr = [string]$entry.updatedDate }
        elseif ($entry.when)          { $dateStr = [string]$entry.when }
        elseif ($entry.date)          { $dateStr = [string]$entry.date }

        $month = ""
        try {
            if ($dateStr -match "^(\d{4})-(\d{2})") { $month = $Matches[1] + "-" + $Matches[2] }
            elseif ($dateStr -match "^(\d{2})/(\d{2})/(\d{4})") { $month = $Matches[3] + "-" + $Matches[2] }
            elseif ($dateStr -match "(\d{4})-(\d{2})-(\d{2})") { $month = $Matches[1] + "-" + $Matches[2] }
        } catch {}
        if ([string]::IsNullOrWhiteSpace($month)) { continue }
        if (-not $monthsForObj.ContainsKey($month)) { $monthsForObj[$month] = 0 }
        $monthsForObj[$month]++
    }

    foreach ($m in $monthsForObj.Keys) {
        $entriesByObjMonth.Add([pscustomobject]@{
            ObjectId  = $objId
            ObjectKey = $objKey
            Label     = $obj.Label
            Month     = $m
            NbEntries = $monthsForObj[$m]
        }) | Out-Null
    }

    Start-Sleep -Milliseconds 80
}

Write-Info "=== Collecte terminee ==="
Write-Info ("  Objets traites        : " + $objProcessed)
Write-Info ("  Objets avec historique : " + $objWithHistory)
Write-Info ("  Objets en erreur       : " + $objErrors)
Write-Info ("  Total entrees hist.    : " + $totalHistEntries)
Write-Info ("  Mois distincts         : " + $historyByMonth.Count)

# ============================================================
# 8. ANALYSE PAR MOIS
# ============================================================
Write-Info "=== Volume historique par mois ==="

$monthSorted = $historyByMonth.GetEnumerator() | Sort-Object Name

$monthResults = New-Object System.Collections.Generic.List[object]

foreach ($m in $monthSorted) {
    $monthKey  = $m.Name
    $nbEntries = $m.Value
    $nbObjects = if ($objectsByMonth.ContainsKey($monthKey)) { $objectsByMonth[$monthKey].Count } else { 0 }

    $monthResults.Add([pscustomobject]@{
        "Mois"             = $monthKey
        "Nb Entrees"       = $nbEntries
        "Nb Objets"        = $nbObjects
        "Pct Total"        = if ($totalHistEntries -gt 0) { [math]::Round(($nbEntries / $totalHistEntries) * 100, 1) } else { 0 }
    }) | Out-Null

    Write-Info ("  " + $monthKey + " : " + $nbEntries + " entrees / " + $nbObjects + " objets (" + [math]::Round(($nbEntries / [math]::Max($totalHistEntries, 1)) * 100, 1) + "%)")
}

# ============================================================
# 9. IDENTIFICATION DES 3 PLUS VIEUX MOIS
# ============================================================
Write-Info "=== Identification des 3 plus vieux mois (candidats purge) ==="

$oldest3 = @($monthSorted | Select-Object -First 3)

$purgeEntries = 0
$purgeObjects = 0

foreach ($m in $oldest3) {
    $monthKey  = $m.Name
    $nbEntries = $m.Value
    $nbObjects = if ($objectsByMonth.ContainsKey($monthKey)) { $objectsByMonth[$monthKey].Count } else { 0 }
    $purgeEntries += $nbEntries
    $purgeObjects += $nbObjects

    Write-Info ("  [CANDIDAT PURGE] " + $monthKey + " : " + $nbEntries + " entrees / " + $nbObjects + " objets")
}

$pctPurge = if ($totalHistEntries -gt 0) { [math]::Round(($purgeEntries / $totalHistEntries) * 100, 1) } else { 0 }
Write-Info ""
Write-Info ("  Total purge estimee : " + $purgeEntries + " entrees (" + $pctPurge + "% du total) sur " + $purgeObjects + " objets")

# ============================================================
# 10. EXPORT CSV
# ============================================================
Write-Info "=== Export CSV ==="

function Export-CsvStrict {
    param([string]$Path, [string[]]$Headers, $Rows)
    $sb = New-Object System.Text.StringBuilder
    [void]$sb.AppendLine($Headers -join ";")
    foreach ($row in $Rows) {
        $vals = foreach ($h in $Headers) {
            $s = [string]$row.$h
            if ($s.Contains(";") -or $s.Contains('"') -or $s.Contains("`n")) {
                '"' + $s.Replace('"', '""') + '"'
            } else { $s }
        }
        [void]$sb.AppendLine($vals -join ";")
    }
    [System.IO.File]::WriteAllText($Path, $sb.ToString(), [System.Text.Encoding]::UTF8)
    Write-Info ("CSV exporte : " + $Path + " (" + $Rows.Count + " lignes)")
}

# 10a. Volume par mois
$csvVolume = Join-Path $exportsDir ("Assets-VolumeHistorique-ParMois_" + $runStamp + ".csv")
Export-CsvStrict -Path $csvVolume `
    -Headers @("Mois", "Nb Entrees", "Nb Objets", "Pct Total") `
    -Rows $monthResults

# 10b. Detail des objets dans les 3 plus vieux mois (dry-run)
$oldest3Keys      = @($oldest3 | ForEach-Object { $_.Name })
$purgeDetailRows  = $entriesByObjMonth | Where-Object { $oldest3Keys -contains $_.Month } | Sort-Object Month, ObjectKey

$csvDryRun = Join-Path $exportsDir ("Assets-DryRun-Purge3MoisAnciens_" + $runStamp + ".csv")
Export-CsvStrict -Path $csvDryRun `
    -Headers @("Month", "ObjectId", "ObjectKey", "Label", "NbEntries") `
    -Rows $purgeDetailRows

# 10c. Synthese purge
$csvSynthese = Join-Path $exportsDir ("Assets-SynthesePurge_" + $runStamp + ".csv")
$syntheseRows = New-Object System.Collections.Generic.List[object]
foreach ($m in $oldest3) {
    $monthKey  = $m.Name
    $nbEntries = $m.Value
    $nbObjects = if ($objectsByMonth.ContainsKey($monthKey)) { $objectsByMonth[$monthKey].Count } else { 0 }
    $syntheseRows.Add([pscustomobject]@{
        "Mois"       = $monthKey
        "Nb Entrees" = $nbEntries
        "Nb Objets"  = $nbObjects
        "Pct Total"  = if ($totalHistEntries -gt 0) { [math]::Round(($nbEntries / $totalHistEntries) * 100, 1) } else { 0 }
        "Action"     = "DRY RUN - Candidat suppression"
    }) | Out-Null
}
Export-CsvStrict -Path $csvSynthese `
    -Headers @("Mois", "Nb Entrees", "Nb Objets", "Pct Total", "Action") `
    -Rows $syntheseRows

# ============================================================
# 11. RESUME FINAL
# ============================================================
Write-Info "=== RESUME FINAL ==="
Write-Info ("  Object Types audites               : " + ($ObjectTypeIds -join ", "))
Write-Info ("  Objets traites                     : " + $objProcessed)
Write-Info ("  Objets avec historique             : " + $objWithHistory)
Write-Info ("  Objets en erreur                   : " + $objErrors)
Write-Info ""
Write-Info ("  Total entrees historique           : " + $totalHistEntries)
Write-Info ("  Mois distincts                     : " + $historyByMonth.Count)
Write-Info ("  Mois le plus ancien                : " + $(if ($monthSorted.Count -gt 0) { $monthSorted[0].Name } else { "N/A" }))
Write-Info ("  Mois le plus recent                : " + $(if ($monthSorted.Count -gt 0) { $monthSorted[-1].Name } else { "N/A" }))
Write-Info ""
Write-Info ("  === DRY RUN - CANDIDATS PURGE ===")
Write-Info ("  Mois candidats                     : " + ($oldest3Keys -join ", "))
Write-Info ("  Entrees a supprimer                : " + $purgeEntries + " / " + $totalHistEntries + " (" + $pctPurge + "%)")
Write-Info ("  Objets concernes                   : " + $purgeObjects)
Write-Info ""
Write-Info ("  !!! AUCUNE SUPPRESSION EFFECTUEE - MODE DRY RUN !!!")
Write-Info ""
Write-Info ("  Exports : " + $exportsDir)
Write-Info ("  Cache   : " + $cacheDir)
Write-Info ("  Log     : " + $logFile)
Write-Log ("=== FIN " + $scriptName + " ===") "INFO"
