<#
Sheet6-AnalyseParLigneBudgetaire.ps1
Feuille 6 : Analyse par Ligne Budgétaire (Account Tempo).

Version standalone avec diagnostics intégrés.
Produit 2 CSV :
  - "Analyse Par Ligne Budgétaire.csv"  (agrégé par Account)
  - "Diagnostic Personnes.csv"          (détail par personne : type, jours, TJM)
#>

param(
    [string]$CacheDir   = (Join-Path $PSScriptRoot "cache"),
    [string]$ExportDir  = (Join-Path $PSScriptRoot "exports"),
    [string]$LogDir     = (Join-Path $PSScriptRoot "logs")
)

$ErrorActionPreference = "Stop"
$scriptName = "Sheet6-AnalyseParLigneBudgetaire"
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

# Types considérés comme externes (facturables)
$typesExternes = @("Externe", "Autres prestation")

function Test-IsExterne([string]$TypeRes) {
    foreach ($t in $typesExternes) {
        if ($TypeRes -eq $t) { return $true }
    }
    return $false
}

# --- Map Assets par accountId ---
$assetByAccountId = @{}
$diag_noId = 0
foreach ($u in $assetUsers) {
    $aid = Get-AssetProp $u "Compte Jira ID"
    if ([string]::IsNullOrWhiteSpace($aid)) { $diag_noId++; continue }
    $assetByAccountId[$aid] = $u
}
Write-Log "Map AssetByAccountId : $($assetByAccountId.Count) entrées ($diag_noId sans ID)"

# --- DIAGNOSTIC ---
$diag_typeValues = @{}
$diag_withTarif = 0
foreach ($aid in $assetByAccountId.Keys) {
    $asset = $assetByAccountId[$aid]
    $typeRes = Get-AssetProp $asset "Type ressource" "(vide)"
    if (-not $diag_typeValues.ContainsKey($typeRes)) { $diag_typeValues[$typeRes] = 0 }
    $diag_typeValues[$typeRes]++
    $tjm = Parse-Tjm (Get-AssetProp $asset "Tarif € TTC")
    if ($tjm -gt 0) { $diag_withTarif++ }
}
Write-Log "--- DIAGNOSTIC Type ressource ---" "DEBUG"
foreach ($k in ($diag_typeValues.Keys | Sort-Object)) {
    $isExt = Test-IsExterne $k
    $tag = if ($isExt) { " → EXTERNE" } else { " → INTERNE" }
    Write-Log "  '$k' : $($diag_typeValues[$k]) personnes$tag" "DEBUG"
}
Write-Log "  Fiches avec TJM > 0 : $diag_withTarif" "DEBUG"
Write-Log "--- FIN DIAGNOSTIC ---" "DEBUG"

# ============================================================
# 3. HEURES PAR PERSONNE (pour diagnostic CSV)
# ============================================================
$personHeures = @{}
foreach ($wl in $worklogs) {
    if ([string]::IsNullOrWhiteSpace([string]$wl.startDate)) { continue }
    $authorId = ""
    try { $authorId = [string]$wl.author.accountId } catch {}
    if ([string]::IsNullOrWhiteSpace($authorId)) { continue }
    $timeH = ([double]$wl.timeSpentSeconds) / 3600.0
    if (-not $personHeures.ContainsKey($authorId)) { $personHeures[$authorId] = 0.0 }
    $personHeures[$authorId] += $timeH
}
Write-Log "Personnes avec worklogs : $($personHeures.Count)"

$diag_matched = 0; $diag_unmatched = 0
foreach ($personId in $personHeures.Keys) {
    if ($assetByAccountId.ContainsKey($personId)) { $diag_matched++ } else { $diag_unmatched++ }
}
Write-Log "DIAG - Matchées Assets: $diag_matched | Non matchées: $diag_unmatched" "DEBUG"

# --- CSV Diagnostic Personnes ---
$personRows = New-Object System.Collections.Generic.List[object]
foreach ($authorId in ($personHeures.Keys | Sort-Object)) {
    $totalH = $personHeures[$authorId]
    $jours  = [math]::Round($totalH / 8.0, 2)

    $nom = ""; $prenom = ""; $typeRes = "Inconnu"; $tarifStr = "Non défini"; $tjm = 0.0
    $matched = "NON"; $isExt = $false

    if ($assetByAccountId.ContainsKey($authorId)) {
        $matched = "OUI"
        $asset = $assetByAccountId[$authorId]
        $nom      = Get-AssetProp $asset "Nom"
        $prenom   = Get-AssetProp $asset "Prénom"
        $typeRes  = Get-AssetProp $asset "Type ressource" "Inconnu"
        $tarifStr = Get-AssetProp $asset "Tarif € TTC" "Non défini"
        $tjm      = Parse-Tjm $tarifStr
        $isExt    = Test-IsExterne $typeRes
    }

    $classif = if ($isExt) { "Externe" } else { "Interne" }
    $budget  = if ($isExt) { [math]::Round($jours * $tjm, 2) } else { 0 }

    $personRows.Add([pscustomobject]@{
        "Account ID"      = $authorId
        "Nom"             = $nom
        "Prénom"           = $prenom
        "Matched Asset"   = $matched
        "Type ressource"  = $typeRes
        "Classification"  = $classif
        "Tarif brut"      = $tarifStr
        "TJM parsé"        = $tjm
        "Total Heures"    = [math]::Round($totalH, 2)
        "Total Jours"     = $jours
        "Budget (€)"      = $budget
    }) | Out-Null
}

# ============================================================
# 4. AGRÉGATION PAR ACCOUNT
# ============================================================
$accountData = @{}

foreach ($wl in $worklogs) {
    if ([string]::IsNullOrWhiteSpace([string]$wl.startDate)) { continue }

    $authorId = ""
    try { $authorId = [string]$wl.author.accountId } catch {}
    if ([string]::IsNullOrWhiteSpace($authorId)) { continue }

    $timeH = ([double]$wl.timeSpentSeconds) / 3600.0

    # --- Account depuis l'issue ---
    $accountName = "Non défini"
    if ($wl.issue -and $wl.issue.id) {
        $issueId = [string]$wl.issue.id
        if ($issuesById.ContainsKey($issueId)) {
            $issue = $issuesById[$issueId]
            if ($issue.fields) {
                $cf = $issue.fields.customfield_10032
                if ($cf) {
                    if ($cf.value)    { $accountName = [string]$cf.value }
                    elseif ($cf.name) { $accountName = [string]$cf.name }
                    else              { $accountName = [string]$cf }
                }
            }
        }
    }

    if (-not $accountData.ContainsKey($accountName)) {
        $accountData[$accountName] = @{
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

    $entry = $accountData[$accountName]
    if ($isExt) {
        $entry.heuresExternes += $timeH
        $entry.budgetExternes += ($timeH / 8.0) * $tjm
    } else {
        $entry.heuresInternes += $timeH
    }
}

Write-Log "Accounts distincts : $($accountData.Count)"

# ============================================================
# 5. CONSTRUCTION LIGNES ACCOUNT
# ============================================================
$accountRows = New-Object System.Collections.Generic.List[object]

foreach ($accountName in ($accountData.Keys | Sort-Object)) {
    $entry = $accountData[$accountName]
    $accountRows.Add([pscustomobject]@{
        "Account"                      = $accountName
        "Jours Internes"               = [math]::Round($entry.heuresInternes / 8.0, 2)
        "Jours Externes"               = [math]::Round($entry.heuresExternes / 8.0, 2)
        "Budget Consommé Externes (€)" = [math]::Round($entry.budgetExternes, 2)
    }) | Out-Null
}

# ============================================================
# 6. EXPORT CSV (séparateur ; UTF-8)
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
    [System.IO.File]::WriteAllText($Path, $sb.ToString(), [System.Text.Encoding]::UTF8)
    Write-Log "CSV exporté : $Path ($($Rows.Count) lignes)"
}

# CSV 1 : Analyse par Account
$csvAccount = Join-Path $ExportDir "Analyse Par Ligne Budgétaire.csv"
Export-CsvStrict -Path $csvAccount `
    -Headers @("Account", "Jours Internes", "Jours Externes", "Budget Consommé Externes (€)") `
    -Rows $accountRows

# CSV 2 : Diagnostic par Personne
$csvPersonnes = Join-Path $ExportDir "Diagnostic Personnes.csv"
Export-CsvStrict -Path $csvPersonnes `
    -Headers @("Account ID", "Nom", "Prénom", "Matched Asset", "Type ressource",
               "Classification", "Tarif brut", "TJM parsé", "Total Heures", "Total Jours", "Budget (€)") `
    -Rows $personRows

# ============================================================
# 7. RÉSUMÉ
# ============================================================
$totalJoursInt = ($accountRows | Measure-Object -Property "Jours Internes" -Sum).Sum
$totalJoursExt = ($accountRows | Measure-Object -Property "Jours Externes" -Sum).Sum
$totalBudget   = ($accountRows | Measure-Object -Property "Budget Consommé Externes (€)" -Sum).Sum

$nbExt    = ($personRows | Where-Object { $_.Classification -eq "Externe" }).Count
$nbInt    = ($personRows | Where-Object { $_.Classification -eq "Interne" }).Count

Write-Log "=== Résumé ==="
Write-Log "  Accounts        : $($accountRows.Count)"
Write-Log "  Personnes       : $($personRows.Count) ($nbInt internes, $nbExt externes)"
Write-Log "  Jours Internes  : $([math]::Round($totalJoursInt, 2))"
Write-Log "  Jours Externes  : $([math]::Round($totalJoursExt, 2))"
Write-Log "  Budget Externes : $([math]::Round($totalBudget, 2)) €"
Write-Log "=== Fin $scriptName ==="