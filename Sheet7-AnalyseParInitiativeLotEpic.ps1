<#
Sheet7-AnalyseParInitiativeLotEpic.ps1
Analyse par Initiative, Lot et Epic.
Pour chaque issue de type Initiative/Lot/Epic ayant des worklogs :
  jours internes, jours externes, budget consommé (TJM × jours externes).

Version standalone — utilise les fichiers cache JSON produits par le lanceur.
Produit 2 CSV :
  - "Analyse Par Initiative Lot Epic.csv"  (agrégé par issue)
  - "Diagnostic Personnes.csv"             (réutilisé si déjà présent)
#>

param(
    [string]$CacheDir   = (Join-Path $PSScriptRoot "cache"),
    [string]$ExportDir  = (Join-Path $PSScriptRoot "exports"),
    [string]$LogDir     = (Join-Path $PSScriptRoot "logs")
)

$ErrorActionPreference = "Stop"
$scriptName = "Sheet7-AnalyseParInitiativeLotEpic"
$timestamp  = Get-Date -Format "yyyyMMdd_HHmmss"

foreach ($dir in @($ExportDir, $LogDir)) {
    if (-not (Test-Path $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
}

$logFile = Join-Path $LogDir "$scriptName`_$timestamp.log"

function Write-Log {
    param([string]$Message, [string]$Level = "INFO")
    $line = "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') [$Level] $Message"
    Add-Content -Path $logFile -Value $line -Encoding UTF8
    if ($Level -eq "ERROR")    { Write-Host $line -ForegroundColor Red }
    elseif ($Level -eq "WARN") { Write-Host $line -ForegroundColor Yellow }
    elseif ($Level -eq "DEBUG"){ Write-Host $line -ForegroundColor Cyan }
    else { Write-Host $line }
}

# ============================================================
# 1. CHARGEMENT DES CACHES
# ============================================================
function Load-JsonCache {
    param([string]$FileName)
    $path = Join-Path $CacheDir $FileName
    if (-not (Test-Path $path)) {
        Write-Log "Cache introuvable : $path" "ERROR"
        throw "Fichier cache requis introuvable : $FileName"
    }
    $size = [math]::Round((Get-Item $path).Length / 1MB, 2)
    Write-Log "Chargement cache : $FileName ($size MB)"
    $raw = Get-Content -Path $path -Raw -Encoding UTF8
    return ($raw | ConvertFrom-Json)
}

Write-Log "=== Démarrage $scriptName ==="

$worklogs   = Load-JsonCache "worklogs.json"
$issuesRaw  = Load-JsonCache "issues.json"
$assetUsers = Load-JsonCache "saisi_temps_asset.json"

Write-Log "Worklogs : $($worklogs.Count)"
Write-Log "Assets   : $($assetUsers.Count)"

# ============================================================
# 2. CONSTRUCTION DES MAPS
# ============================================================

$issuesById = @{}
foreach ($p in $issuesRaw.PSObject.Properties) {
    $issuesById[[string]$p.Name] = $p.Value
}
Write-Log "Map IssuesById : $($issuesById.Count) entrées"

# --- Helpers ---
function Get-AssetProp {
    param($Asset, [string]$PropName, [string]$Default = "")
    $p = $Asset.PSObject.Properties[$PropName]
    if ($p -and -not [string]::IsNullOrWhiteSpace([string]$p.Value)) {
        return ([string]$p.Value).Trim()
    }
    return $Default
}

function Parse-Tjm([string]$TarifStr) {
    if ([string]::IsNullOrWhiteSpace($TarifStr) -or $TarifStr -eq "Non défini") { return 0.0 }
    $clean = $TarifStr -replace '[^\d,\.]', ''
    $clean = $clean -replace ',', '.'
    $val = 0.0
    if ([double]::TryParse($clean,
        [System.Globalization.NumberStyles]::Any,
        [System.Globalization.CultureInfo]::InvariantCulture,
        [ref]$val)) {
        return $val
    }
    return 0.0
}

$typesExternes = @("Externe", "Autres prestation")

function Test-IsExterne([string]$TypeRes) {
    foreach ($t in $typesExternes) {
        if ($TypeRes -eq $t) { return $true }
    }
    return $false
}

# Types d'issues à analyser
$typesAnalyses = @("Initiative", "Lot", "Epic")

# --- Map Assets par accountId ---
$assetByAccountId = @{}
foreach ($u in $assetUsers) {
    $aid = Get-AssetProp $u "Compte Jira ID"
    if ([string]::IsNullOrWhiteSpace($aid)) { continue }
    $assetByAccountId[$aid] = $u
}
Write-Log "Map AssetByAccountId : $($assetByAccountId.Count) entrées"

# ============================================================
# 3. RÉSOLUTION PARENT (Initiative > Lot > Epic)
# ============================================================
# Pour chaque issue, retrouver ses parents dans la hiérarchie
# en utilisant le champ fields.parent (récursif, limité au cache)

function Get-ParentChain {
    param(
        [Parameter(Mandatory)]$Issue,
        [Parameter(Mandatory)][hashtable]$IssuesById
    )
    # Remonte la hiérarchie et retourne @{ Initiative = "..."; Lot = "..."; Epic = "..." }
    $chain = @{ Initiative = ""; Lot = ""; Epic = "" }
    $current = $Issue
    $visited = @{}

    # D'abord, enregistrer l'issue elle-même
    $selfType = ""
    if ($current.fields -and $current.fields.issuetype) {
        $selfType = [string]$current.fields.issuetype.name
    }
    $selfKey = [string]$current.key
    $selfSummary = ""
    if ($current.fields -and $current.fields.summary) {
        $selfSummary = [string]$current.fields.summary
    }
    $selfLabel = "$selfKey - $selfSummary"

    if ($selfType -eq "Initiative") { $chain.Initiative = $selfLabel }
    elseif ($selfType -eq "Lot")    { $chain.Lot = $selfLabel }
    elseif ($selfType -eq "Epic")   { $chain.Epic = $selfLabel }

    # Remonter les parents
    while ($current.fields -and $current.fields.parent) {
        $parentId = ""
        if ($current.fields.parent.id) {
            $parentId = [string]$current.fields.parent.id
        } elseif ($current.fields.parent.key) {
            # Chercher par key
            foreach ($iss in $IssuesById.Values) {
                if ([string]$iss.key -eq [string]$current.fields.parent.key) {
                    $parentId = [string]$iss.id
                    break
                }
            }
        }

        if ([string]::IsNullOrWhiteSpace($parentId) -or $visited.ContainsKey($parentId)) { break }
        $visited[$parentId] = $true

        if (-not $IssuesById.ContainsKey($parentId)) { break }
        $parent = $IssuesById[$parentId]

        $pType = ""
        if ($parent.fields -and $parent.fields.issuetype) {
            $pType = [string]$parent.fields.issuetype.name
        }
        $pKey = [string]$parent.key
        $pSummary = ""
        if ($parent.fields -and $parent.fields.summary) {
            $pSummary = [string]$parent.fields.summary
        }
        $pLabel = "$pKey - $pSummary"

        if ($pType -eq "Initiative" -and [string]::IsNullOrWhiteSpace($chain.Initiative)) {
            $chain.Initiative = $pLabel
        }
        elseif ($pType -eq "Lot" -and [string]::IsNullOrWhiteSpace($chain.Lot)) {
            $chain.Lot = $pLabel
        }
        elseif ($pType -eq "Epic" -and [string]::IsNullOrWhiteSpace($chain.Epic)) {
            $chain.Epic = $pLabel
        }

        $current = $parent
    }

    return $chain
}

# ============================================================
# 4. TRAITEMENT — Agrégation par Issue (Initiative/Lot/Epic)
# ============================================================
# Clé = issueId, Valeur = données agrégées
$issueData = @{}
$diag_filteredOut = 0
$diag_noIssue = 0

foreach ($wl in $worklogs) {
    if ([string]::IsNullOrWhiteSpace([string]$wl.startDate)) { continue }

    $authorId = ""
    try { $authorId = [string]$wl.author.accountId } catch {}
    if ([string]::IsNullOrWhiteSpace($authorId)) { continue }

    $timeH = ([double]$wl.timeSpentSeconds) / 3600.0

    # --- Résolution de l'issue ---
    if (-not $wl.issue -or -not $wl.issue.id) { $diag_noIssue++; continue }
    $issueId = [string]$wl.issue.id

    if (-not $issuesById.ContainsKey($issueId)) { $diag_noIssue++; continue }
    $issue = $issuesById[$issueId]

    # --- Filtrer sur les types Initiative / Lot / Epic ---
    $issueType = ""
    if ($issue.fields -and $issue.fields.issuetype) {
        $issueType = [string]$issue.fields.issuetype.name
    }

    if ($issueType -notin $typesAnalyses) {
        $diag_filteredOut++
        continue
    }

    # Initialiser l'entrée issue si nouvelle
    if (-not $issueData.ContainsKey($issueId)) {
        $key = [string]$issue.key
        $summary = ""
        if ($issue.fields -and $issue.fields.summary) { $summary = [string]$issue.fields.summary }

        # Account
        $accountName = "Non défini"
        if ($issue.fields) {
            $cf = $issue.fields.customfield_10032
            if ($cf) {
                if ($cf.value)    { $accountName = [string]$cf.value }
                elseif ($cf.name) { $accountName = [string]$cf.name }
                else              { $accountName = [string]$cf }
            }
        }

        # Chaîne parent
        $chain = Get-ParentChain -Issue $issue -IssuesById $issuesById

        $issueData[$issueId] = @{
            key            = $key
            summary        = $summary
            issueType      = $issueType
            account        = $accountName
            initiative     = $chain.Initiative
            lot            = $chain.Lot
            epic           = $chain.Epic
            heuresInternes = 0.0
            heuresExternes = 0.0
            budgetExternes = 0.0
        }
    }

    # --- Type + TJM ---
    $typeRes = "Inconnu"
    $tjm = 0.0
    $isExt = $false
    if ($assetByAccountId.ContainsKey($authorId)) {
        $asset   = $assetByAccountId[$authorId]
        $typeRes = Get-AssetProp $asset "Type ressource" "Inconnu"
        $tjm     = Parse-Tjm (Get-AssetProp $asset "Tarif € TTC")
        $isExt   = Test-IsExterne $typeRes
    }

    $entry = $issueData[$issueId]
    if ($isExt) {
        $entry.heuresExternes += $timeH
        $entry.budgetExternes += ($timeH / 8.0) * $tjm
    } else {
        $entry.heuresInternes += $timeH
    }
}

Write-Log "Issues Initiative/Lot/Epic avec worklogs : $($issueData.Count)"
Write-Log "Worklogs sur autres types (filtrés) : $diag_filteredOut"
Write-Log "Worklogs sans issue résolue : $diag_noIssue"

# --- DIAGNOSTIC : répartition par type ---
$diag_byType = @{}
foreach ($entry in $issueData.Values) {
    $t = $entry.issueType
    if (-not $diag_byType.ContainsKey($t)) { $diag_byType[$t] = 0 }
    $diag_byType[$t]++
}
foreach ($t in ($diag_byType.Keys | Sort-Object)) {
    Write-Log "  Type '$t' : $($diag_byType[$t]) issues" "DEBUG"
}

# ============================================================
# 5. CONSTRUCTION LIGNES CSV
# ============================================================
$headers = @(
    "Type Issue", "Issue Key", "Summary", "Account",
    "Initiative", "Lot", "Epic",
    "Jours Internes", "Jours Externes", "Budget Consommé Externes (€)"
)

$rows = New-Object System.Collections.Generic.List[object]

# Tri par Type (Initiative > Lot > Epic) puis par Key
$sortOrder = @{ "Initiative" = 1; "Lot" = 2; "Epic" = 3 }
$sortedEntries = $issueData.Values | Sort-Object {
    $order = 99
    if ($sortOrder.ContainsKey($_.issueType)) { $order = $sortOrder[$_.issueType] }
    $order
}, { $_.key }

foreach ($entry in $sortedEntries) {
    $rows.Add([pscustomobject]@{
        "Type Issue"                   = $entry.issueType
        "Issue Key"                    = $entry.key
        "Summary"                      = $entry.summary
        "Account"                      = $entry.account
        "Initiative"                   = $entry.initiative
        "Lot"                          = $entry.lot
        "Epic"                         = $entry.epic
        "Jours Internes"               = [math]::Round($entry.heuresInternes / 8.0, 2)
        "Jours Externes"               = [math]::Round($entry.heuresExternes / 8.0, 2)
        "Budget Consommé Externes (€)" = [math]::Round($entry.budgetExternes, 2)
    }) | Out-Null
}

# ============================================================
# 6. EXPORT CSV (séparateur ; UTF-8 + gestion fichier verrouillé)
# ============================================================
function Export-CsvStrict {
    param(
        [string]$Path,
        [string[]]$Headers,
        [System.Collections.Generic.List[object]]$Rows
    )
    $sb = New-Object System.Text.StringBuilder
    [void]$sb.AppendLine(($Headers -join ";"))
    foreach ($row in $Rows) {
        $vals = foreach ($h in $Headers) {
            $s = [string]$row.$h
            if ($s.Contains(";") -or $s.Contains('"') -or $s.Contains("`n")) {
                '"' + $s.Replace('"', '""') + '"'
            } else { $s }
        }
        [void]$sb.AppendLine(($vals -join ";"))
    }

    # Tentative d'écriture avec fallback horodaté si fichier verrouillé
    try {
        [System.IO.File]::WriteAllText($Path, $sb.ToString(), [System.Text.Encoding]::UTF8)
        Write-Log "CSV exporté : $Path ($($Rows.Count) lignes)"
    } catch [System.IO.IOException] {
        $dir  = [System.IO.Path]::GetDirectoryName($Path)
        $name = [System.IO.Path]::GetFileNameWithoutExtension($Path)
        $ts   = Get-Date -Format "yyyyMMdd_HHmmss"
        $fallback = Join-Path $dir "$name`_$ts.csv"
        Write-Log "Fichier verrouillé, export vers : $fallback" "WARN"
        [System.IO.File]::WriteAllText($fallback, $sb.ToString(), [System.Text.Encoding]::UTF8)
        Write-Log "CSV exporté (fallback) : $fallback ($($Rows.Count) lignes)"
    }
}

$csvPath = Join-Path $ExportDir "Analyse Par Initiative Lot Epic.csv"
Export-CsvStrict -Path $csvPath -Headers $headers -Rows $rows

# ============================================================
# 7. RÉSUMÉ
# ============================================================
$totalJoursInt = ($rows | Measure-Object -Property "Jours Internes" -Sum).Sum
$totalJoursExt = ($rows | Measure-Object -Property "Jours Externes" -Sum).Sum
$totalBudget   = ($rows | Measure-Object -Property "Budget Consommé Externes (€)" -Sum).Sum

Write-Log "=== Résumé ==="
Write-Log "  Issues          : $($rows.Count)"
foreach ($t in ($diag_byType.Keys | Sort-Object)) {
    $tRows = $rows | Where-Object { $_."Type Issue" -eq $t }
    $tJoursI = ($tRows | Measure-Object -Property "Jours Internes" -Sum).Sum
    $tJoursE = ($tRows | Measure-Object -Property "Jours Externes" -Sum).Sum
    $tBudget = ($tRows | Measure-Object -Property "Budget Consommé Externes (€)" -Sum).Sum
    Write-Log "  $t : $($tRows.Count) issues | Int: $([math]::Round($tJoursI,2))j | Ext: $([math]::Round($tJoursE,2))j | Budget: $([math]::Round($tBudget,2)) €"
}
Write-Log "  TOTAL Jours Int : $([math]::Round($totalJoursInt, 2))"
Write-Log "  TOTAL Jours Ext : $([math]::Round($totalJoursExt, 2))"
Write-Log "  TOTAL Budget    : $([math]::Round($totalBudget, 2)) €"
Write-Log "=== Fin $scriptName ==="