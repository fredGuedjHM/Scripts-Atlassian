# ============================================================
# Manage-IssueLinkTypes.ps1
# v1.9 - Inventaire, deduplication et migration des types de liens Jira
# ============================================================
# Objectif :
#   1. Lister tous les types de liens avec leur usage reel (parcours des tickets)
#   2. Identifier les doublons potentiels (inward/outward identiques)
#   3. Mode DryRun : afficher le plan de migration sans rien modifier
#   4. Mode Execute : migrer les liens vers le type de remplacement puis supprimer le doublon
# ============================================================
# Changements v1.9 :
#   - Strategie de comptage : parcours global "issueLinks is not EMPTY" au lieu de
#     JQL "issuelinktype = X" (non supporte en Cloud -> 410)
#   - Un seul parcours pour TOUS les types simultanement
#   - Colonne IssueKeys dans le CSV (tickets concernes, separes par |)
#   - Fix Measure-Object : propriete Usage correctement nommee
#   - Export CSV dans exports\ avec horodatage
#   - Proxy systeme (prc37cti1.hm.dm.ad:8080)
#   - Credentials : cle Email en priorite
# ============================================================

[Diagnostics.CodeAnalysis.SuppressMessageAttribute(
    'PSAvoidUsingPlainTextForPassword', '',
    Justification = 'XmlSecretsFile contient un chemin vers un fichier XML chiffre (Export-Clixml), pas un mot de passe en clair.'
)]
param(
    [string]$XmlSecretsFile = "",
    [switch]$Execute,
    [string]$DeleteLinkTypeId      = "",
    [string]$ReplacementLinkTypeId = "",
    [int]   $MaxKeysPerType        = 500,
    [int]   $ThrottleMs            = 200
)

# --- Proxy systeme (Harmonie Mutuelle) ---
$sysProxy = [System.Net.WebRequest]::GetSystemWebProxy()
$sysProxy.Credentials = [System.Net.CredentialCache]::DefaultCredentials
[System.Net.WebRequest]::DefaultWebProxy = $sysProxy

# --- Resolution du repertoire du script (ISE compatible) ---
$ScriptDir = $PSScriptRoot
if (-not $ScriptDir) {
    if ($psISE -and $psISE.CurrentFile -and $psISE.CurrentFile.FullPath) {
        $ScriptDir = Split-Path $psISE.CurrentFile.FullPath -Parent
    } else {
        $ScriptDir = Split-Path $MyInvocation.MyCommand.Path -Parent
    }
}
if (-not $ScriptDir) { $ScriptDir = (Get-Location).Path }

# --- Dossier exports ---
$ExportDir = Join-Path $ScriptDir "exports"
if (-not (Test-Path $ExportDir)) { New-Item -ItemType Directory -Path $ExportDir | Out-Null }

# --- Timestamp global ---
$RunStamp  = Get-Date -Format "yyyyMMdd-HHmmss"
$CsvPath   = Join-Path $ExportDir "LinkTypes-Inventory_$RunStamp.csv"
$LogPath   = Join-Path $ExportDir "LinkTypes_$RunStamp.log"

# --- Configuration ---
$SiteUrl        = "https://jiradot.atlassian.net"
$DefaultXmlFile = Join-Path $ScriptDir "secrets\site-admin.xml"
if (-not $XmlSecretsFile) { $XmlSecretsFile = $DefaultXmlFile }

# ============================================================
# FONCTIONS UTILITAIRES
# ============================================================

function Write-Log {
    param([string]$Level, [string]$Message)
    $ts    = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $color = switch ($Level) {
        "INFO"  { "White"    }
        "WARN"  { "Yellow"   }
        "ERROR" { "Red"      }
        "OK"    { "Green"    }
        "DRY"   { "Cyan"     }
        default { "White"    }
    }
    $line = "[$ts][$Level] $Message"
    Write-Host $line -ForegroundColor $color
    Add-Content -Path $LogPath -Value $line -Encoding UTF8
}

function Read-StringKey {
    param([object]$Obj, [string[]]$Keys)
    foreach ($k in $Keys) {
        $v = $null
        try { $v = $Obj[$k] } catch {}
        if (-not $v) { try { $v = $Obj.PSObject.Properties[$k]?.Value } catch {} }
        if ($v -and $v -is [string] -and $v.Trim() -ne '') { return $v.Trim() }
    }
    return ''
}

function Read-SecureToken {
    param([object]$Obj, [string[]]$Keys)
    foreach ($k in $Keys) {
        $v = $null
        try { $v = $Obj[$k] } catch {}
        if (-not $v) { try { $v = $Obj.PSObject.Properties[$k]?.Value } catch {} }
        if ($null -ne $v) {
            if ($v -is [System.Security.SecureString]) {
                return (New-Object PSCredential("x", $v)).GetNetworkCredential().Password
            }
            if ($v -is [string] -and $v.Trim() -ne '') { return $v.Trim() }
        }
    }
    return ''
}

function Import-SiteCredentials {
    param([string]$Path)
    Write-Log "INFO" "Chargement credentials : $Path"
    if (-not (Test-Path $Path)) {
        Write-Log "ERROR" "Fichier credentials introuvable : $Path"
        exit 1
    }
    $obj = Import-Clixml -Path $Path

    $user  = ''
    $token = ''

    if ($obj -is [PSCredential]) {
        $user  = $obj.UserName
        $token = $obj.GetNetworkCredential().Password
    } else {
        $user  = Read-StringKey  -Obj $obj -Keys @('Email','Login','User','Username')
        $token = Read-SecureToken -Obj $obj -Keys @('ApiTokenSecureString','ApiToken','Token','Password')
    }

    if ($user -eq '' -or $token -eq '') {
        $keys = ''
        try { $keys = ($obj.Keys -join ', ') } catch { $keys = ($obj.PSObject.Properties.Name -join ', ') }
        Write-Log "ERROR" "Credentials incomplets (user='$user', token vide=$($token -eq ''))"
        Write-Log "ERROR" "Cles disponibles dans le fichier : $keys"
        exit 1
    }

    Write-Log "INFO" "  Connecte en tant que : $user"
    $pair  = "${user}:${token}"
    $bytes = [Text.Encoding]::ASCII.GetBytes($pair)
    $b64   = [Convert]::ToBase64String($bytes)
    return @{ "Authorization" = "Basic $b64"; "Accept" = "application/json"; "Content-Type" = "application/json" }
}

function Invoke-JiraGet {
    param([string]$Endpoint, [int]$MaxRetries = 3)
    $url = "$SiteUrl$Endpoint"
    for ($attempt = 1; $attempt -le $MaxRetries; $attempt++) {
        try {
            $resp = Invoke-WebRequest -Uri $url -Method GET -Headers $script:Headers `
                        -UseBasicParsing -ErrorAction Stop
            $data = $resp.Content | ConvertFrom-Json
            return @{ ok = $true; data = $data; status = [int]$resp.StatusCode }
        } catch {
            $status = 0
            if ($_.Exception.Response) { $status = [int]$_.Exception.Response.StatusCode }
            if ($attempt -lt $MaxRetries -and $status -in @(429, 500, 502, 503)) {
                Start-Sleep -Seconds (2 * $attempt); continue
            }
            return @{ ok = $false; status = $status; error = $_.Exception.Message }
        }
    }
}

function Invoke-JiraPost {
    param([string]$Endpoint, [object]$Body, [int]$MaxRetries = 3)
    $url     = "$SiteUrl$Endpoint"
    $bodyStr = $Body | ConvertTo-Json -Depth 10 -Compress
    for ($attempt = 1; $attempt -le $MaxRetries; $attempt++) {
        try {
            $resp = Invoke-WebRequest -Uri $url -Method POST -Headers $script:Headers `
                        -Body $bodyStr -ContentType "application/json; charset=utf-8" `
                        -UseBasicParsing -ErrorAction Stop
            $data = $resp.Content | ConvertFrom-Json
            return @{ ok = $true; data = $data; status = [int]$resp.StatusCode }
        } catch {
            $status = 0
            if ($_.Exception.Response) { $status = [int]$_.Exception.Response.StatusCode }
            if ($attempt -lt $MaxRetries -and $status -in @(429, 500, 502, 503)) {
                Start-Sleep -Seconds (2 * $attempt); continue
            }
            return @{ ok = $false; status = $status; error = $_.Exception.Message }
        }
    }
}

function Invoke-JiraDelete {
    param([string]$Endpoint, [int]$MaxRetries = 3)
    $url = "$SiteUrl$Endpoint"
    for ($attempt = 1; $attempt -le $MaxRetries; $attempt++) {
        try {
            $resp = Invoke-WebRequest -Uri $url -Method DELETE -Headers $script:Headers `
                        -UseBasicParsing -ErrorAction Stop
            return @{ ok = $true; status = [int]$resp.StatusCode }
        } catch {
            $status = 0
            if ($_.Exception.Response) { $status = [int]$_.Exception.Response.StatusCode }
            if ($status -eq 204) { return @{ ok = $true; status = 204 } }
            if ($attempt -lt $MaxRetries -and $status -in @(429, 500, 502, 503)) {
                Start-Sleep -Seconds (2 * $attempt); continue
            }
            return @{ ok = $false; status = $status; error = $_.Exception.Message }
        }
    }
}

function ConvertTo-NormalizedText {
    param([string]$t)
    return $t.ToLower().Trim() -replace '\s+', ' '
}

# ============================================================
# COMPTAGE GLOBAL : un seul parcours pour tous les types
# ============================================================
# Strategie : JQL "issueLinks is not EMPTY" + fields=key,issuelinks
# On parcourt toutes les pages et on incremente les compteurs par typeId.
# Avantage : 1 seul parcours au lieu de N appels JQL (qui retournent 410 en Cloud).
# ============================================================

function Get-AllLinkUsage {
    param(
        [hashtable]$LinkTypeMap,   # id -> name
        [int]$MaxKeysPerType = 500,
        [int]$ThrottleMs     = 200
    )

    # Initialisation des compteurs
    $usageCount = @{}   # typeId -> int
    $usageKeys  = @{}   # typeId -> [List[string]]
    foreach ($id in $LinkTypeMap.Keys) {
        $usageCount[$id] = 0
        $usageKeys[$id]  = [System.Collections.Generic.List[string]]::new()
    }

    $pageSize = 100
    $startAt  = 0
    $total    = [int]::MaxValue
    $page     = 0
    $seenLinks = [System.Collections.Generic.HashSet[string]]::new()

    Write-Log "INFO" "  Parcours global des tickets avec liens (JQL: issueLinks is not EMPTY)..."

    while ($startAt -lt $total) {
        $page++
        $body = @{
            jql        = "issueLinks is not EMPTY ORDER BY key ASC"
            startAt    = $startAt
            maxResults = $pageSize
            fields     = @("key", "issuelinks")
        }
        $resp = Invoke-JiraPost -Endpoint "/rest/api/3/search/jql" -Body $body
        if (-not $resp.ok) {
            Write-Log "WARN" "  Erreur page $page (status $($resp.status)) - arret du parcours"
            break
        }

        $data   = $resp.data
        $issues = $data.issues
        if ($null -eq $issues -or $issues.Count -eq 0) { break }

        # Mise a jour du total a la premiere page
        if ($page -eq 1) {
            $total = [int]$data.total
            Write-Log "INFO" "  Total tickets avec liens : $total"
        }

        foreach ($issue in $issues) {
            $links = $issue.fields.issuelinks
            if ($null -eq $links) { continue }
            foreach ($lnk in $links) {
                $lid = $lnk.id
                if (-not $seenLinks.Add($lid)) { continue }   # deja compte (lien bidirectionnel)

                $tid = $lnk.type.id
                if (-not $usageCount.ContainsKey($tid)) {
                    # Type inconnu (ne devrait pas arriver)
                    $usageCount[$tid] = 0
                    $usageKeys[$tid]  = [System.Collections.Generic.List[string]]::new()
                }
                $usageCount[$tid]++

                # Collecter les keys (source + cible)
                $keysToAdd = @($issue.key)
                if ($lnk.inwardIssue  -and $lnk.inwardIssue.key)  { $keysToAdd += $lnk.inwardIssue.key  }
                if ($lnk.outwardIssue -and $lnk.outwardIssue.key) { $keysToAdd += $lnk.outwardIssue.key }
                foreach ($k in $keysToAdd) {
                    if ($usageKeys[$tid].Count -lt $MaxKeysPerType -and -not $usageKeys[$tid].Contains($k)) {
                        $usageKeys[$tid].Add($k)
                    }
                }
            }
        }

        $startAt += $issues.Count
        if ($page % 10 -eq 0) {
            Write-Log "INFO" "  ... page $page / $([Math]::Ceiling($total / $pageSize)) ($startAt / $total tickets traites)"
        }
        if ($ThrottleMs -gt 0) { Start-Sleep -Milliseconds $ThrottleMs }
    }

    Write-Log "INFO" "  Parcours termine : $($seenLinks.Count) liens uniques analyses"
    return @{ Count = $usageCount; Keys = $usageKeys }
}

# ============================================================
# MAIN
# ============================================================

Write-Host ""
Write-Host "  ========================================" -ForegroundColor Cyan
Write-Host "  GESTION DES TYPES DE LIENS JIRA"         -ForegroundColor Cyan
Write-Host "  v1.9 - Inventaire et deduplication"       -ForegroundColor Cyan
if (-not $Execute) {
    Write-Host "  MODE : DRY RUN (aucune modification)"  -ForegroundColor Yellow
} else {
    Write-Host "  MODE : EXECUTION (modifications actives)" -ForegroundColor Red
}
Write-Host "  Script dir : $ScriptDir"                  -ForegroundColor Gray
Write-Host "  ========================================" -ForegroundColor Cyan
Write-Host ""

# --- Chargement credentials ---
$script:Headers = Import-SiteCredentials -Path $XmlSecretsFile

# --- Test de connexion ---
$meResp = Invoke-JiraGet -Endpoint "/rest/api/3/myself"
if (-not $meResp.ok) {
    Write-Log "ERROR" "Echec de connexion (status $($meResp.status)) : $($meResp.error)"
    exit 1
}
Write-Log "INFO" "  Connexion OK : $($meResp.data.displayName) ($($meResp.data.emailAddress))"
Write-Host ""

# ============================================================
# ETAPE 1 : Recuperation des types de liens
# ============================================================
Write-Log "INFO" "=== ETAPE 1 : Recuperation des types de liens ==="
$ltResp = Invoke-JiraGet -Endpoint "/rest/api/3/issueLinkType"
if (-not $ltResp.ok) {
    Write-Log "ERROR" "Impossible de recuperer les types de liens (status $($ltResp.status))"
    exit 1
}
$linkTypes = $ltResp.data.issueLinkTypes
Write-Log "INFO" "  $($linkTypes.Count) types de liens recuperes"
Write-Host ""

# Map id -> name pour le parcours global
$linkTypeMap = @{}
foreach ($lt in $linkTypes) { $linkTypeMap[$lt.id] = $lt.name }

# ============================================================
# ETAPE 2 : Comptage global (un seul parcours)
# ============================================================
Write-Log "INFO" "=== ETAPE 2 : Comptage de l'usage par type (parcours global) ==="
Write-Log "INFO" "  MaxKeysPerType = $MaxKeysPerType"
Write-Host ""

$usageData = Get-AllLinkUsage -LinkTypeMap $linkTypeMap -MaxKeysPerType $MaxKeysPerType -ThrottleMs $ThrottleMs
Write-Host ""

# ============================================================
# Construction de l'inventaire
# ============================================================
$inventory = [System.Collections.Generic.List[PSObject]]::new()
foreach ($lt in $linkTypes) {
    $cnt  = if ($usageData.Count.ContainsKey($lt.id)) { $usageData.Count[$lt.id] } else { 0 }
    $keys = if ($usageData.Keys.ContainsKey($lt.id))  { $usageData.Keys[$lt.id]  } else { @() }

    $keyStr = ''
    if ($keys.Count -gt 0) {
        $keyStr = $keys -join '|'
        if ($cnt -gt $keys.Count) {
            $keyStr += "|(+$($cnt - $keys.Count) autres)"
        }
    }

    $inventory.Add([PSCustomObject]@{
        Id        = $lt.id
        Name      = $lt.name
        Inward    = $lt.inward
        Outward   = $lt.outward
        UsageCount = $cnt
        IssueKeys = $keyStr
    })
}

# ============================================================
# Affichage du tableau
# ============================================================
Write-Log "INFO" "=== INVENTAIRE DES TYPES DE LIENS ==="
Write-Host ""
Write-Host ("{0,-6} {1,-35} {2,-28} {3,-28} {4,8}" -f "ID", "NOM", "INWARD", "OUTWARD", "USAGE")
Write-Host ("{0,-6} {1,-35} {2,-28} {3,-28} {4,8}" -f "---", "---", "------", "-------", "-----")

$inventory | Sort-Object Name | ForEach-Object {
    $color = if ($_.UsageCount -eq 0) { "DarkGray" } else { "White" }
    Write-Host ("{0,-6} {1,-35} {2,-28} {3,-28} {4,8}" -f `
        $_.Id, $_.Name, $_.Inward, $_.Outward, $_.UsageCount) -ForegroundColor $color
}

Write-Host ""
Write-Log "INFO" "  Types non utilises (0 liens) : $(($inventory | Where-Object { $_.UsageCount -eq 0 }).Count)"
Write-Host ""

# ============================================================
# ETAPE 3 : Detection des doublons
# ============================================================
Write-Log "INFO" "=== ETAPE 3 : Detection des doublons ==="
Write-Host ""

$groups = @{}
foreach ($item in $inventory) {
    $key = "$(ConvertTo-NormalizedText $item.Inward)|$(ConvertTo-NormalizedText $item.Outward)"
    if (-not $groups.ContainsKey($key)) { $groups[$key] = @() }
    $groups[$key] += $item
}

$duplicates = $groups.GetEnumerator() | Where-Object { $_.Value.Count -gt 1 }
if ($duplicates) {
    Write-Log "WARN" "[WARN] $(@($duplicates).Count) groupe(s) de doublons detecte(s) :"
    Write-Host ""
    foreach ($dup in $duplicates) {
        $sorted = $dup.Value | Sort-Object -Property UsageCount -Descending
        Write-Host "  Groupe : '$($sorted[0].Inward)' / '$($sorted[0].Outward)'" -ForegroundColor Yellow
        foreach ($item in $sorted) {
            Write-Host ("    - ID={0} | Name='{1}' | Usage={2}" -f $item.Id, $item.Name, $item.UsageCount) -ForegroundColor Yellow
        }
        Write-Host ""
        $keep = $sorted[0]
        $del  = $sorted[1]
        Write-Host "  Suggestion : conserver ID=$($keep.Id) '$($keep.Name)' (usage: $($keep.UsageCount))" -ForegroundColor Green
        Write-Host "  Supprimer  : .\Manage-IssueLinkTypes.ps1 -DeleteLinkTypeId $($del.Id) -ReplacementLinkTypeId $($keep.Id)" -ForegroundColor Cyan
        Write-Host ""
    }
} else {
    Write-Log "OK" "  Aucun doublon detecte (tous les inward/outward sont uniques)"
}

# ============================================================
# ETAPE 4 : Migration et suppression (si demande)
# ============================================================
if ($DeleteLinkTypeId) {
    Write-Host ""
    Write-Log "INFO" "=== ETAPE 4 : Migration et suppression ==="

    $toDelete = $inventory | Where-Object { $_.Id -eq $DeleteLinkTypeId }
    if ($null -eq $toDelete) {
        Write-Log "ERROR" "Type de lien ID=$DeleteLinkTypeId introuvable"
        exit 1
    }

    Write-Log "INFO" "  Type a supprimer : ID=$($toDelete.Id) | '$($toDelete.Name)' ($($toDelete.UsageCount) liens)"

    if ($toDelete.UsageCount -gt 0 -and -not $ReplacementLinkTypeId) {
        Write-Log "ERROR" "Ce type est utilise par $($toDelete.UsageCount) tickets."
        Write-Log "ERROR" "Vous devez specifier -ReplacementLinkTypeId pour migrer les liens."
        Write-Host ""
        Write-Log "INFO" "Types de remplacement possibles :"
        $inventory | Where-Object { $_.Id -ne $DeleteLinkTypeId } | Sort-Object Name | ForEach-Object {
            Write-Host ("    -ReplacementLinkTypeId {0}  # '{1}' (usage: {2})" -f $_.Id, $_.Name, $_.UsageCount) -ForegroundColor Cyan
        }
        exit 1
    }

    if ($toDelete.UsageCount -gt 0) {
        $replacement = $inventory | Where-Object { $_.Id -eq $ReplacementLinkTypeId }
        if ($null -eq $replacement) {
            Write-Log "ERROR" "Type de remplacement ID=$ReplacementLinkTypeId introuvable"
            exit 1
        }

        Write-Log "INFO" "  Remplacement : ID=$($replacement.Id) | '$($replacement.Name)'"
        Write-Host ""

        if (-not $Execute) {
            Write-Log "DRY" "  [DRY RUN] Migration de $($toDelete.UsageCount) liens :"
            Write-Log "DRY" "    '$($toDelete.Name)' (ID=$($toDelete.Id)) --> '$($replacement.Name)' (ID=$($replacement.Id))"
            Write-Log "DRY" "  [DRY RUN] Suppression du type '$($toDelete.Name)' (ID=$($toDelete.Id))"
            Write-Host ""
            Write-Log "INFO" "  Pour executer : relancez avec -Execute"
            Write-Host ("    .\Manage-IssueLinkTypes.ps1 -DeleteLinkTypeId $DeleteLinkTypeId -ReplacementLinkTypeId $ReplacementLinkTypeId -Execute") -ForegroundColor Cyan
        } else {
            # Recuperer les liens a migrer depuis les IssueKeys deja collectees
            Write-Log "INFO" "  Migration en cours (re-lecture des liens depuis les tickets concernes)..."

            $keysToProcess = @()
            if ($toDelete.IssueKeys) {
                $keysToProcess = $toDelete.IssueKeys -split '\|' | Where-Object { $_ -notmatch '^\(' }
            }

            $migrated  = 0
            $errors    = 0
            $seenLinks = [System.Collections.Generic.HashSet[string]]::new()

            foreach ($issueKey in $keysToProcess) {
                $issResp = Invoke-JiraGet -Endpoint "/rest/api/3/issue/${issueKey}?fields=issuelinks"
                if (-not $issResp.ok) { continue }
                $links = $issResp.data.fields.issuelinks
                if (-not $links) { continue }

                foreach ($lnk in $links) {
                    if ($lnk.type.id -ne $DeleteLinkTypeId) { continue }
                    if (-not $seenLinks.Add($lnk.id)) { continue }

                    $inKey  = if ($lnk.inwardIssue)  { $lnk.inwardIssue.key  } else { $issueKey }
                    $outKey = if ($lnk.outwardIssue) { $lnk.outwardIssue.key } else { $issueKey }

                    $delResp = Invoke-JiraDelete -Endpoint "/rest/api/3/issueLink/$($lnk.id)"
                    if (-not $delResp.ok -and $delResp.status -ne 204) {
                        Write-Log "WARN" "    Echec suppression lien $($lnk.id) (status $($delResp.status))"
                        $errors++; continue
                    }

                    $newLink = @{
                        type         = @{ id = $ReplacementLinkTypeId }
                        inwardIssue  = @{ key = $inKey  }
                        outwardIssue = @{ key = $outKey }
                    }
                    $createResp = Invoke-JiraPost -Endpoint "/rest/api/3/issueLink" -Body $newLink
                    if ($createResp.ok -or $createResp.status -eq 201) {
                        $migrated++
                    } else {
                        Write-Log "WARN" "    Echec creation lien $inKey <-> $outKey (status $($createResp.status))"
                        $errors++
                    }

                    if ($migrated % 50 -eq 0 -and $migrated -gt 0) {
                        Write-Log "INFO" "    $migrated liens migres..."
                    }
                    if ($ThrottleMs -gt 0) { Start-Sleep -Milliseconds $ThrottleMs }
                }
            }

            Write-Log "OK" "  Migration terminee : $migrated migres, $errors erreurs"

            if ($errors -eq 0) {
                Write-Log "INFO" "  Suppression du type '$($toDelete.Name)' (ID=$($toDelete.Id))..."
                $delTypeResp = Invoke-JiraDelete -Endpoint "/rest/api/3/issueLinkType/$($toDelete.Id)"
                if ($delTypeResp.ok -or $delTypeResp.status -eq 204) {
                    Write-Log "OK" "  Type '$($toDelete.Name)' supprime avec succes"
                } else {
                    Write-Log "ERROR" "  Echec suppression type (status $($delTypeResp.status))"
                }
            } else {
                Write-Log "WARN" "  $errors erreurs de migration - suppression du type ANNULEE"
                Write-Log "WARN" "  Verifiez les liens en erreur puis relancez"
            }
        }
    } else {
        if (-not $Execute) {
            Write-Log "DRY" "  [DRY RUN] Type '$($toDelete.Name)' non utilise -> suppression directe possible"
            Write-Log "INFO" "  Pour executer : relancez avec -Execute"
        } else {
            Write-Log "INFO" "  Type non utilise -> suppression directe..."
            $delResp = Invoke-JiraDelete -Endpoint "/rest/api/3/issueLinkType/$($toDelete.Id)"
            if ($delResp.ok -or $delResp.status -eq 204) {
                Write-Log "OK" "  Type '$($toDelete.Name)' supprime avec succes"
            } else {
                Write-Log "ERROR" "  Echec suppression (status $($delResp.status))"
            }
        }
    }
}

# ============================================================
# Export CSV
# ============================================================
$inventory | Sort-Object Name | Select-Object Id, Name, Inward, Outward, UsageCount, IssueKeys |
    Export-Csv -Path $CsvPath -NoTypeInformation -Encoding UTF8 -Delimiter ";"

# ============================================================
# RESUME
# ============================================================
$totalUsage = ($inventory | Measure-Object -Property UsageCount -Sum).Sum

Write-Host ""
Write-Log "INFO" "========================================"
Write-Log "INFO" "RESUME"
Write-Log "INFO" "========================================"
Write-Log "INFO" "Site       : $SiteUrl"
Write-Log "INFO" "Types      : $($inventory.Count)"
Write-Log "INFO" "Doublons   : $(@($duplicates).Count) groupe(s)"
Write-Log "INFO" "Usage total: $totalUsage liens uniques"
Write-Log "INFO" "Export CSV : $CsvPath"
Write-Log "INFO" "Log        : $LogPath"
Write-Log "INFO" "========================================"
Write-Log "INFO" "Termine."
Write-Host ""
