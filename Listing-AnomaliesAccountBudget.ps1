<#
Listing-AnomaliesAccountBudget.ps1
Détection des anomalies de propagation du champ Account (Tempo) depuis les tickets Budget.

Hiérarchie Jiradot :
  Budget (4) → Initiative (3) → Lot (2) → Epic (1) → Story/Task/Bug → Subtask

CAS 1 : Ticket ayant un ancêtre Budget mais Account (customfield_10032) VIDE
         → l'automation de propagation descendante n'a pas fonctionné.
CAS 2 : Ticket ayant Account renseigné mais AUCUN ancêtre Budget
         → valeur Account orpheline.

FILTRE DOMAINES :
  Seules les issues dont l'assignee ou le reporter a un email dans les
  domaines autorises sont analysees (mutex.fr, harmonie-mutuelle.fr, etc.).
  Les issues sans email identifiable sont incluses par defaut.

Le script remonte la chaîne parentale via fields.parent.
Si un parent n'est pas dans le cache issues.json, il est fetché via l'API Jira
puis sauvegardé dans le cache pour les exécutions suivantes.
#>

[CmdletBinding()]
param(
    [switch]$UseSystemProxy = $true,
    [switch]$ProxyUseDefaultCredentials = $true,
    [string]$ProxyUrl
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
$scriptName = "Listing-AnomaliesAccountBudget"

# ============================================================
# 0b. FILTRE DOMAINES EMAIL
# ============================================================
# Seules les issues dont l'assignee ou le reporter a un email
# dans ces domaines seront analysees.
# Les issues sans email identifiable sont INCLUSES par defaut
# (pour ne pas rater d'anomalies).

$allowedDomains = @(
    "mutex.fr",
    "mutex-exterieur.fr",
    "harmonie-mutuelle.fr",
    "prestataire.sihm.fr",
    "chorum.fr"
)

# ============================================================
# 1. LOG
# ============================================================
$logFile = Join-Path $logsDir ("$scriptName`_{0}.log" -f $runStamp)

function Write-Log {
    param([string]$Message = "",
          [ValidateSet("INFO","WARN","ERROR","DEBUG")][string]$Level = "INFO")
    if ([string]::IsNullOrWhiteSpace($Message)) { return }
    $ts = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    try { Add-Content -Path $logFile -Value "[$ts] [$Level] $Message" -ErrorAction Stop } catch {}
}
function Write-Info($msg)   { Write-Host "[INFO] $msg";  Write-Log $msg "INFO" }
function Write-Warn($msg)   { Write-Warning $msg;        Write-Log $msg "WARN" }
function Write-ErrLog($msg) { Write-Error $msg;          Write-Log $msg "ERROR" }

Write-Log "=== DÉBUT $scriptName ===" "INFO"

# ============================================================
# 2. PROXY
# ============================================================
function Initialize-Proxy {
    param([switch]$UseSystemProxy, [string]$ProxyUrl)
    try {
        [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
        if ($ProxyUrl) {
            [System.Net.WebRequest]::DefaultWebProxy = New-Object System.Net.WebProxy($ProxyUrl,$true)
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
        $s = $ex.Response.GetResponseStream(); $r = New-Object System.IO.StreamReader($s)
        $b = $r.ReadToEnd(); $r.Dispose(); $s.Dispose(); return $b
    } catch { return $null }
}

function Invoke-ApiGet {
    param([Parameter(Mandatory)][string]$Url, [Parameter(Mandatory)][hashtable]$Headers)
    $params = @{ Method='GET'; Uri=$Url; Headers=$Headers; ContentType='application/json'; ErrorAction='Stop' }
    $px = Get-EffectiveProxyUri -TargetUrl $Url
    if ($px) { $params.Proxy=$px; if ($ProxyUseDefaultCredentials) { $params.ProxyUseDefaultCredentials=$true } }
    try { return Invoke-RestMethod @params }
    catch [System.Net.WebException] {
        $body = Get-WebExceptionBody $_.Exception
        if ($body) { Write-ErrLog "GET $Url : $($_.Exception.Message)`nBody:`n$body" }
        else       { Write-ErrLog "GET $Url : $($_.Exception.Message)" }
        throw
    }
}

# ============================================================
# 4. CREDENTIALS JIRA
# ============================================================
$jiraCredFile = Join-Path $secretsDir "jira-jiradot.cred.xml"

if (-not (Test-Path $jiraCredFile)) {
    Write-ErrLog "Fichier Jira creds introuvable: $jiraCredFile"
    throw "Lance d'abord Save-JiraCredential.ps1 ou mvp-reporting.ps1 pour créer les credentials."
}

$jiraData    = Import-Clixml -Path $jiraCredFile
$jiraBaseUrl = ($jiraData.JiraBaseUrl).TrimEnd("/")
$jiraCred    = $jiraData.Credential
$jiraEmail   = $jiraCred.UserName
$jiraToken   = $jiraCred.GetNetworkCredential().Password

function New-BasicAuthHeader([string]$User,[string]$Pass) {
    $b64 = [Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes("$User`:$Pass"))
    return @{ Authorization="Basic $b64"; Accept="application/json" }
}

$jiraHeaders = New-BasicAuthHeader -User $jiraEmail -Pass $jiraToken
Write-Info "Jira: $jiraBaseUrl (user=$jiraEmail)"

# ============================================================
# 5. HELPERS
# ============================================================
function Load-JsonCache {
    param([string]$FileName)
    $path = Join-Path $cacheDir $FileName
    if (-not (Test-Path $path)) { return $null }
    $size = [math]::Round((Get-Item $path).Length / 1MB, 2)
    Write-Info "Chargement cache : $FileName ($size MB)"
    $raw = Get-Content -Path $path -Raw -Encoding UTF8
    return ($raw | ConvertFrom-Json)
}

# ============================================================
# 5b. FILTRE DOMAINES — FONCTIONS
# ============================================================

function Get-EmailDomain([string]$email) {
    if ([string]::IsNullOrWhiteSpace($email)) { return $null }
    $idx = $email.IndexOf("@")
    if ($idx -lt 0) { return $null }
    return $email.Substring($idx + 1).Trim().ToLower()
}

function Test-IssueDomainAllowed {
    <#
    Verifie si l'issue appartient a un utilisateur des domaines autorises.
    Regarde l'assignee puis le reporter (en fallback).
    Si aucun email n'est disponible, retourne $true (inclusion par defaut).
    #>
    param([Parameter(Mandatory)]$Issue)

    if (-not $Issue.fields) { return $true }

    # Tenter l'assignee
    $assigneeEmail = $null
    if ($Issue.fields.assignee) {
        if ($Issue.fields.assignee.emailAddress) {
            $assigneeEmail = [string]$Issue.fields.assignee.emailAddress
        }
    }

    # Tenter le reporter
    $reporterEmail = $null
    if ($Issue.fields.reporter) {
        if ($Issue.fields.reporter.emailAddress) {
            $reporterEmail = [string]$Issue.fields.reporter.emailAddress
        }
    }

    # Si on a au moins un email, verifier le domaine
    $domainFound = $null

    if ($assigneeEmail) {
        $domainFound = Get-EmailDomain $assigneeEmail
    }
    if (-not $domainFound -and $reporterEmail) {
        $domainFound = Get-EmailDomain $reporterEmail
    }

    # Si aucun email n'est disponible, inclure par defaut
    if (-not $domainFound) { return $true }

    # Verifier contre les domaines autorises
    foreach ($d in $allowedDomains) {
        if ($domainFound -eq $d.ToLower()) { return $true }
    }

    return $false
}

# ============================================================
# 6. CHARGEMENT DES ISSUES (cache)
# ============================================================
$issuesRaw = Load-JsonCache "issues.json"
if (-not $issuesRaw) {
    Write-ErrLog "Cache issues.json introuvable. Lancer mvp-reporting.ps1 d'abord."
    throw "Fichier cache requis introuvable : issues.json"
}

# Construire la map par ID ET par Key
$issuesById  = @{}
$issuesByKey = @{}
foreach ($p in $issuesRaw.PSObject.Properties) {
    $issue = $p.Value
    $issuesById[[string]$p.Name] = $issue
    if ($issue.key) { $issuesByKey[[string]$issue.key] = $issue }
}
Write-Info "Issues en cache : $($issuesById.Count)"

# --- Diagnostic : répartition par type dans le cache ---
$typeDistrib = @{}
foreach ($iss in $issuesById.Values) {
    $t = ""
    if ($iss.fields -and $iss.fields.issuetype) { $t = [string]$iss.fields.issuetype.name }
    if ([string]::IsNullOrWhiteSpace($t)) { $t = "(inconnu)" }
    if (-not $typeDistrib.ContainsKey($t)) { $typeDistrib[$t] = 0 }
    $typeDistrib[$t]++
}
Write-Info "Types en cache :"
foreach ($t in ($typeDistrib.Keys | Sort-Object)) {
    Write-Info "  $t : $($typeDistrib[$t])"
}

# ============================================================
# 6b. PRE-FILTRAGE PAR DOMAINE
# ============================================================
Write-Info "=== Filtrage par domaines email ==="
Write-Info "  Domaines autorises : $($allowedDomains -join ', ')"

$issuesFiltered    = @{}
$cFilteredIn       = 0
$cFilteredOut      = 0
$cFilteredNoEmail  = 0

# Stats par domaine
$domainStats = @{}

foreach ($issueId in $issuesById.Keys) {
    $issue = $issuesById[$issueId]

    if (Test-IssueDomainAllowed -Issue $issue) {
        $issuesFiltered[$issueId] = $issue
        $cFilteredIn++

        # Stats domaine
        $emailForStat = $null
        if ($issue.fields) {
            if ($issue.fields.assignee -and $issue.fields.assignee.emailAddress) {
                $emailForStat = [string]$issue.fields.assignee.emailAddress
            } elseif ($issue.fields.reporter -and $issue.fields.reporter.emailAddress) {
                $emailForStat = [string]$issue.fields.reporter.emailAddress
            }
        }
        $domKey = if ($emailForStat) { Get-EmailDomain $emailForStat } else { "(sans email)" }
        if ($domKey) {
            if (-not $domainStats.ContainsKey($domKey)) { $domainStats[$domKey] = 0 }
            $domainStats[$domKey]++
        }
        if (-not $emailForStat) { $cFilteredNoEmail++ }
    } else {
        $cFilteredOut++

        # Log le domaine exclu (pour diagnostic)
        $exclEmail = $null
        if ($issue.fields) {
            if ($issue.fields.assignee -and $issue.fields.assignee.emailAddress) {
                $exclEmail = [string]$issue.fields.assignee.emailAddress
            } elseif ($issue.fields.reporter -and $issue.fields.reporter.emailAddress) {
                $exclEmail = [string]$issue.fields.reporter.emailAddress
            }
        }
        $exclDom = if ($exclEmail) { Get-EmailDomain $exclEmail } else { "?" }
        if ($exclDom) {
            if (-not $domainStats.ContainsKey("EXCLU:$exclDom")) { $domainStats["EXCLU:$exclDom"] = 0 }
            $domainStats["EXCLU:$exclDom"]++
        }
    }
}

Write-Info "  Issues incluses     : $cFilteredIn"
Write-Info "  Issues exclues      : $cFilteredOut"
Write-Info "  Issues sans email   : $cFilteredNoEmail (incluses par defaut)"
Write-Info ""
Write-Info "  Repartition par domaine :"
foreach ($dk in ($domainStats.Keys | Sort-Object)) {
    Write-Info "    $dk : $($domainStats[$dk])"
}

# ============================================================
# 7. FETCH API POUR PARENTS MANQUANTS
# ============================================================
$fetchedParents    = @{}
$fetchedNotFound   = @{}
$apiFetchCount     = 0

function Resolve-Issue {
    <#
    Résout une issue par ID ou Key.
    1. Cache issues.json (filtré)
    2. Cache issues.json (complet, pour les parents)
    3. Cache dynamique (déjà fetché cette session)
    4. API Jira → mise en cache dynamique
    #>
    param([string]$IssueIdOrKey)

    if ([string]::IsNullOrWhiteSpace($IssueIdOrKey)) { return $null }

    # 1. Cache filtré
    if ($issuesFiltered.ContainsKey($IssueIdOrKey)) { return $issuesFiltered[$IssueIdOrKey] }

    # 2. Cache complet (pour remonter la chaine parentale meme si le parent est hors domaine)
    if ($issuesById.ContainsKey($IssueIdOrKey))  { return $issuesById[$IssueIdOrKey] }
    if ($issuesByKey.ContainsKey($IssueIdOrKey)) { return $issuesByKey[$IssueIdOrKey] }

    # 3. Cache dynamique (déjà fetché)
    if ($fetchedParents.ContainsKey($IssueIdOrKey))  { return $fetchedParents[$IssueIdOrKey] }
    if ($fetchedNotFound.ContainsKey($IssueIdOrKey)) { return $null }

    # 4. Fetch via API (inclut assignee et reporter pour le domaine)
    $fields = "summary,issuetype,parent,status,project,customfield_10032,assignee,reporter"
    $url = "$jiraBaseUrl/rest/api/3/issue/$IssueIdOrKey`?fields=$fields"

    try {
        $resp = Invoke-ApiGet -Url $url -Headers $jiraHeaders
        $script:apiFetchCount++

        if ($resp.id)  { $fetchedParents[[string]$resp.id]  = $resp }
        if ($resp.key) { $fetchedParents[[string]$resp.key] = $resp }

        Write-Log "  API fetch: $($resp.key) [$($resp.fields.issuetype.name)] — $($resp.fields.summary)" "DEBUG"
        return $resp
    }
    catch {
        Write-Log "  API fetch FAILED: $IssueIdOrKey — $($_.Exception.Message)" "WARN"
        $fetchedNotFound[$IssueIdOrKey] = $true
        return $null
    }
}

# ============================================================
# 8. EXTRACTION DU CHAMP ACCOUNT
# ============================================================
function Get-AccountValue($Issue) {
    if (-not $Issue.fields) { return "" }
    $cf = $Issue.fields.customfield_10032
    if (-not $cf) { return "" }
    if ($cf.value)    { return ([string]$cf.value).Trim() }
    if ($cf.name)     { return ([string]$cf.name).Trim() }
    $s = ([string]$cf).Trim()
    if ($s -eq "System.Object" -or $s -like "System.Collections*") { return "" }
    return $s
}

# ============================================================
# 9. RECHERCHE ANCÊTRE BUDGET (avec fetch API si manquant)
# ============================================================
function Find-BudgetAncestor {
    param([Parameter(Mandatory)]$Issue)

    $current  = $Issue
    $visited  = @{}
    $maxDepth = 10

    for ($depth = 0; $depth -lt $maxDepth; $depth++) {
        if (-not $current.fields -or -not $current.fields.parent) { break }

        $parentRef = $current.fields.parent
        $parentIdOrKey = ""
        if ($parentRef.id)      { $parentIdOrKey = [string]$parentRef.id }
        elseif ($parentRef.key) { $parentIdOrKey = [string]$parentRef.key }
        if ([string]::IsNullOrWhiteSpace($parentIdOrKey)) { break }
        if ($visited.ContainsKey($parentIdOrKey)) { break }
        $visited[$parentIdOrKey] = $true

        $parent = Resolve-Issue -IssueIdOrKey $parentIdOrKey
        if (-not $parent) { break }

        $pType = ""
        if ($parent.fields -and $parent.fields.issuetype) {
            $pType = [string]$parent.fields.issuetype.name
        }

        if ($pType -eq "Budget") {
            return $parent
        }

        $current = $parent
    }

    return $null
}

function Get-AncestorChainLabel {
    param([Parameter(Mandatory)]$Issue)

    $chain    = New-Object System.Collections.Generic.List[string]
    $current  = $Issue
    $visited  = @{}
    $maxDepth = 10

    for ($depth = 0; $depth -lt $maxDepth; $depth++) {
        if (-not $current.fields -or -not $current.fields.parent) { break }

        $parentRef = $current.fields.parent
        $parentIdOrKey = ""
        if ($parentRef.id)      { $parentIdOrKey = [string]$parentRef.id }
        elseif ($parentRef.key) { $parentIdOrKey = [string]$parentRef.key }
        if ([string]::IsNullOrWhiteSpace($parentIdOrKey)) { break }
        if ($visited.ContainsKey($parentIdOrKey)) { break }
        $visited[$parentIdOrKey] = $true

        $parent = Resolve-Issue -IssueIdOrKey $parentIdOrKey
        if (-not $parent) {
            $chain.Add("(parent $parentIdOrKey — hors accès)")
            break
        }

        $pType = ""; $pKey = [string]$parent.key; $pSummary = ""
        if ($parent.fields) {
            if ($parent.fields.issuetype) { $pType = [string]$parent.fields.issuetype.name }
            if ($parent.fields.summary)   { $pSummary = [string]$parent.fields.summary }
        }
        $chain.Add("[$pType] $pKey — $pSummary")
        $current = $parent
    }

    if ($chain.Count -eq 0) { return "(aucun parent)" }
    return ($chain -join " → ")
}

# ============================================================
# 10. ANALYSE (sur les issues filtrées par domaine)
# ============================================================
Write-Info "=== Début de l'analyse (issues filtrées par domaine) ==="

$cas1 = New-Object System.Collections.Generic.List[object]
$cas2 = New-Object System.Collections.Generic.List[object]

$diag_ok = 0; $diag_budgetSkipped = 0
$total = $issuesFiltered.Count; $processed = 0; $lastPct = -1

foreach ($issueId in $issuesFiltered.Keys) {
    $processed++
    $pct = [math]::Floor(($processed / [Math]::Max($total,1)) * 100)
    if ($pct -ne $lastPct -and ($pct % 10 -eq 0)) {
        Write-Info "  Progression : $processed / $total ($pct%) — API fetches: $apiFetchCount"
        $lastPct = $pct
    }

    $issue = $issuesFiltered[$issueId]

    # Type de l'issue
    $issueType = ""
    if ($issue.fields -and $issue.fields.issuetype) {
        $issueType = [string]$issue.fields.issuetype.name
    }

    # Ignorer les Budget eux-mêmes
    if ($issueType -eq "Budget") {
        $diag_budgetSkipped++
        continue
    }

    $key     = [string]$issue.key
    $summary = ""
    if ($issue.fields -and $issue.fields.summary) { $summary = [string]$issue.fields.summary }

    # Account
    $accountVal = Get-AccountValue $issue
    $hasAccount = -not [string]::IsNullOrWhiteSpace($accountVal)

    # Projet
    $projectKey  = ""
    $projectName = ""
    if ($issue.fields -and $issue.fields.project) {
        $projectKey  = [string]$issue.fields.project.key
        $projectName = [string]$issue.fields.project.name
    }

    # Statut
    $statut = ""
    if ($issue.fields -and $issue.fields.status) {
        $statut = [string]$issue.fields.status.name
    }

    # Assignee / Reporter (pour le CSV)
    $assigneeName = ""
    $assigneeEmail = ""
    if ($issue.fields -and $issue.fields.assignee) {
        if ($issue.fields.assignee.displayName)  { $assigneeName  = [string]$issue.fields.assignee.displayName }
        if ($issue.fields.assignee.emailAddress) { $assigneeEmail = [string]$issue.fields.assignee.emailAddress }
    }
    $reporterName = ""
    $reporterEmail = ""
    if ($issue.fields -and $issue.fields.reporter) {
        if ($issue.fields.reporter.displayName)  { $reporterName  = [string]$issue.fields.reporter.displayName }
        if ($issue.fields.reporter.emailAddress) { $reporterEmail = [string]$issue.fields.reporter.emailAddress }
    }

    # Recherche ancêtre Budget
    $budgetAncestor = Find-BudgetAncestor -Issue $issue
    $hasBudget = ($null -ne $budgetAncestor)

    $budgetKey = ""; $budgetSummary = ""
    if ($hasBudget) {
        $budgetKey = [string]$budgetAncestor.key
        if ($budgetAncestor.fields -and $budgetAncestor.fields.summary) {
            $budgetSummary = [string]$budgetAncestor.fields.summary
        }
    }

    # Classification
    if ($hasBudget -and -not $hasAccount) {
        $ancestorChain = Get-AncestorChainLabel -Issue $issue
        $cas1.Add([pscustomobject]@{
            "Projet"           = $projectKey
            "Nom Projet"       = $projectName
            "Type"             = $issueType
            "Key"              = $key
            "Résumé"           = $summary
            "Statut"           = $statut
            "Assignee"         = $assigneeName
            "Email Assignee"   = $assigneeEmail
            "Reporter"         = $reporterName
            "Account"          = ""
            "Budget Ancêtre"   = "$budgetKey — $budgetSummary"
            "Chaîne Parents"   = $ancestorChain
            "Anomalie"         = "Account VIDE — Budget ancêtre trouvé"
        }) | Out-Null
    }
    elseif (-not $hasBudget -and $hasAccount) {
        $ancestorChain = Get-AncestorChainLabel -Issue $issue
        $cas2.Add([pscustomobject]@{
            "Projet"           = $projectKey
            "Nom Projet"       = $projectName
            "Type"             = $issueType
            "Key"              = $key
            "Résumé"           = $summary
            "Statut"           = $statut
            "Assignee"         = $assigneeName
            "Email Assignee"   = $assigneeEmail
            "Reporter"         = $reporterName
            "Account"          = $accountVal
            "Budget Ancêtre"   = "(aucun)"
            "Chaîne Parents"   = $ancestorChain
            "Anomalie"         = "Account renseigné — Aucun Budget ancêtre"
        }) | Out-Null
    }
    else {
        $diag_ok++
    }
}

Write-Info "=== Analyse terminée ==="
Write-Info "  Issues en cache (total)                     : $($issuesById.Count)"
Write-Info "  Issues filtrées (domaines autorisés)        : $($issuesFiltered.Count)"
Write-Info "  Issues exclues (domaines hors périmètre)    : $cFilteredOut"
Write-Info "  Issues analysées (hors Budget)              : $($total - $diag_budgetSkipped)"
Write-Info "  Tickets Budget (ignorés)                    : $diag_budgetSkipped"
Write-Info "  Parents fetchés via API                     : $apiFetchCount"
Write-Info "  CAS 1 — Account vide + Budget ancêtre       : $($cas1.Count)"
Write-Info "  CAS 2 — Account renseigné + pas de Budget   : $($cas2.Count)"
Write-Info "  OK (cohérents)                              : $diag_ok"

# ============================================================
# 11. TRI
# ============================================================
$cas1Sorted = $cas1 | Sort-Object "Projet", "Type", "Key"
$cas2Sorted = $cas2 | Sort-Object "Projet", "Type", "Key"

# --- Détail par projet et type ---
Write-Info "=== Détail CAS 1 par Projet ==="
$grpCas1 = $cas1Sorted | Group-Object "Projet"
foreach ($g in ($grpCas1 | Sort-Object Name)) {
    $pName = if ([string]::IsNullOrWhiteSpace($g.Name)) { "(sans projet)" } else { $g.Name }
    $byType = $g.Group | Group-Object "Type"
    $detail = ($byType | Sort-Object Name | ForEach-Object { "$($_.Name):$($_.Count)" }) -join ", "
    Write-Info "  $pName : $($g.Group.Count) ($detail)"
}

Write-Info "=== Détail CAS 2 par Projet ==="
$grpCas2 = $cas2Sorted | Group-Object "Projet"
foreach ($g in ($grpCas2 | Sort-Object Name)) {
    $pName = if ([string]::IsNullOrWhiteSpace($g.Name)) { "(sans projet)" } else { $g.Name }
    $byType = $g.Group | Group-Object "Type"
    $detail = ($byType | Sort-Object Name | ForEach-Object { "$($_.Name):$($_.Count)" }) -join ", "
    Write-Info "  $pName : $($g.Group.Count) ($detail)"
}

# ============================================================
# 12. EXPORT CSV
# ============================================================
$headers = @(
    "Projet", "Nom Projet", "Type", "Key", "Résumé", "Statut",
    "Assignee", "Email Assignee", "Reporter",
    "Account", "Budget Ancêtre", "Chaîne Parents", "Anomalie"
)

function Export-CsvStrict {
    param(
        [string]$Path,
        [string[]]$Headers,
        $Rows
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

    try {
        [System.IO.File]::WriteAllText($Path, $sb.ToString(), [System.Text.Encoding]::UTF8)
        Write-Info "CSV exporté : $Path ($($Rows.Count) lignes)"
    } catch [System.IO.IOException] {
        $dir  = [System.IO.Path]::GetDirectoryName($Path)
        $name = [System.IO.Path]::GetFileNameWithoutExtension($Path)
        $ts   = Get-Date -Format "yyyyMMdd_HHmmss"
        $fallback = Join-Path $dir "$name`_$ts.csv"
        Write-Warn "Fichier verrouillé, export vers : $fallback"
        [System.IO.File]::WriteAllText($fallback, $sb.ToString(), [System.Text.Encoding]::UTF8)
        Write-Info "CSV exporté (fallback) : $fallback ($($Rows.Count) lignes)"
    }
}

$csv1 = Join-Path $exportsDir "Anomalies - Account Manquant Avec Budget.csv"
$csv2 = Join-Path $exportsDir "Anomalies - Account Orphelin Sans Budget.csv"

Export-CsvStrict -Path $csv1 -Headers $headers -Rows $cas1Sorted
Export-CsvStrict -Path $csv2 -Headers $headers -Rows $cas2Sorted

# ============================================================
# 13. ENRICHISSEMENT DU CACHE issues.json
# ============================================================
if ($apiFetchCount -gt 0) {
    Write-Info "Mise à jour du cache issues.json avec $apiFetchCount parents fetchés..."

    $cacheFile = Join-Path $cacheDir "issues.json"
    $rawJson   = Get-Content -Path $cacheFile -Raw -Encoding UTF8
    $cacheObj  = $rawJson | ConvertFrom-Json

    $added = 0
    # Dédupliquer : ne garder que les entrées par ID (pas par Key)
    $alreadyAdded = @{}
    foreach ($key in $fetchedParents.Keys) {
        $fetched = $fetchedParents[$key]
        $id = [string]$fetched.id
        if ([string]::IsNullOrWhiteSpace($id)) { continue }
        if ($alreadyAdded.ContainsKey($id)) { continue }

        # Ajouter seulement si absent du cache original
        if (-not $cacheObj.PSObject.Properties[$id]) {
            $cacheObj | Add-Member -NotePropertyName $id -NotePropertyValue $fetched -Force
            $added++
        }
        $alreadyAdded[$id] = $true
    }

    if ($added -gt 0) {
        # Backup avant écriture
        $backupFile = Join-Path $cacheDir "issues.json.bak"
        try {
            Copy-Item -Path $cacheFile -Destination $backupFile -Force
            Write-Info "Backup créé : issues.json.bak"
        } catch {
            Write-Warn "Impossible de créer le backup : $($_.Exception.Message)"
        }

        $cacheObj | ConvertTo-Json -Depth 10 -Compress | Set-Content -Path $cacheFile -Encoding UTF8
        $newSize = [math]::Round((Get-Item $cacheFile).Length / 1MB, 2)
        Write-Info "Cache enrichi : +$added issues ajoutées ($newSize MB)"
    } else {
        Write-Info "Cache déjà à jour (tous les parents fetchés étaient déjà présents)"
    }
} else {
    Write-Info "Aucun fetch API — cache inchangé"
}

# ============================================================
# 14. RÉSUMÉ FINAL
# ============================================================
Write-Info "=== Résumé ==="
Write-Info "  Domaines autorisés                         : $($allowedDomains -join ', ')"
Write-Info "  Issues en cache (total)                    : $($issuesById.Count)"
Write-Info "  Issues filtrées (dans le périmètre)        : $cFilteredIn"
Write-Info "  Issues exclues (hors périmètre)             : $cFilteredOut"
Write-Info "  Issues sans email (incluses par défaut)     : $cFilteredNoEmail"
Write-Info "  CAS 1 — Account vide + Budget ancêtre     : $($cas1.Count)"
Write-Info "  CAS 2 — Account orphelin sans Budget       : $($cas2.Count)"
Write-Info "  API fetches (parents hors cache)           : $apiFetchCount"
Write-Info "CSV : $exportsDir"
Write-Info "Log : $logFile"
Write-Log "=== FIN $scriptName ===" "INFO"