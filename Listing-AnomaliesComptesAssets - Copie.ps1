<#
Listing-AnomaliesComptesAssets.ps1
Rapport de reconciliation Comptes Jira <-> Fiches Assets.
Approche Assets : endpoint gateway + pagination startAt/maxResults/isLast
                  (identique a mvp-reporting_v2)
Filtres groupes  : multi-valeurs (DSIM + DSIT par defaut)
#>

[CmdletBinding()]
param(
    [switch]$UseSystemProxy = $true,
    [switch]$ProxyUseDefaultCredentials = $true,
    [string]$ProxyUrl,
    [string[]]$GroupFilter = @("DSIM", "DSIT"),
    [string]$AssetsWorkspaceId = "2898703d-5f14-4527-9f8a-ad32b9279ca1"
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
$scriptName = "Listing-AnomaliesComptesAssets"

$allowedDomains = @("mutex.fr","mutex-exterieur.fr","harmonie-mutuelle.fr","prestataire.sihm.fr","chorum.fr")

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

# ============================================================
# 2. PROXY
# ============================================================
function Initialize-Proxy {
    param([switch]$UseSystemProxy, [string]$ProxyUrl)
    try {
        [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
        if ($ProxyUrl) {
            [System.Net.WebRequest]::DefaultWebProxy = New-Object System.Net.WebProxy($ProxyUrl,$true)
            Write-Info ("Proxy: " + $ProxyUrl)
        } elseif ($UseSystemProxy) {
            [System.Net.WebRequest]::DefaultWebProxy = [System.Net.WebRequest]::GetSystemWebProxy()
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
        $body = Get-WebExceptionBody $_.Exception
        Write-ErrLog ("GET " + $Url + " : " + $_.Exception.Message + "`n" + $body)
        throw
    }
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
        Write-ErrLog ("POST UTF8 " + $Url + " : " + $_.Exception.Message + "`n" + $body)
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

$b64 = [Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes($jiraEmail + ":" + $jiraToken))
$jiraHeaders = @{ Authorization = "Basic " + $b64; Accept = "application/json" }
Write-Info ("Jira: " + $jiraBaseUrl + " (user=" + $jiraEmail + ")")

# ============================================================
# 5. HELPERS GENERAUX
# ============================================================
function Get-AssetProp {
    param($Asset, [string]$PropName, [string]$Default = "")
    $p = $Asset.PSObject.Properties[$PropName]
    if ($p -and -not [string]::IsNullOrWhiteSpace([string]$p.Value)) {
        return ([string]$p.Value).Trim()
    }
    return $Default
}

function Load-JsonCache {
    param([string]$FileName)
    $path = Join-Path $cacheDir $FileName
    if (-not (Test-Path $path)) { return $null }
    $size = [math]::Round((Get-Item $path).Length / 1MB, 2)
    Write-Info ("Chargement cache : " + $FileName + " (" + $size + " MB)")
    $raw = Get-Content -Path $path -Raw -Encoding UTF8
    return ($raw | ConvertFrom-Json)
}

function Save-JsonCache {
    param([string]$FileName, $Object)
    $path = Join-Path $cacheDir $FileName
    $Object | ConvertTo-Json -Depth 5 -Compress | Set-Content -Path $path -Encoding UTF8
    $size = [math]::Round((Get-Item $path).Length / 1MB, 2)
    Write-Info ("Cache sauvegarde : " + $FileName + " (" + $size + " MB)")
}

function Get-EmailDomain([string]$Email) {
    if ([string]::IsNullOrWhiteSpace($Email)) { return "" }
    $idx = $Email.IndexOf("@")
    if ($idx -lt 0) { return "" }
    return $Email.Substring($idx + 1).ToLower().Trim()
}

function Test-AllowedDomain([string]$Email) {
    $domain = Get-EmailDomain $Email
    if ([string]::IsNullOrWhiteSpace($domain)) { return $false }
    foreach ($d in $allowedDomains) {
        if ($domain -eq $d.ToLower()) { return $true }
    }
    return $false
}

function Normalize-AccountId([string]$Id) {
    if ([string]::IsNullOrWhiteSpace($Id)) { return "" }
    $Id = $Id.Trim()
    if ($Id -match "^\d+:(.+)$") { return $Matches[1] }
    return $Id
}
# ============================================================
# 6. REFRESH ASSETS (endpoint gateway, pagination startAt/isLast)
# ============================================================
Write-Info "=== Refresh cache Assets (gateway endpoint) ==="

$assetsBaseUrl = $jiraBaseUrl + "/gateway/api/jsm/assets/workspace/" + $AssetsWorkspaceId + "/v1/object/aql"

$eAcc       = [char]233   # é
$prenomKey  = "Pr" + $eAcc + "nom"
$dateEntKey = "Date Entr" + $eAcc + "e"

# Deux requetes separees : objectTypeId=68 (Employe) et objectTypeId=69 (Prestataire)
# Contourne la limite 1000 et le probleme d'accents dans l'AQL
$assetResults = New-Object System.Collections.Generic.List[object]

foreach ($otId in @("68", "69")) {
    $aqlQuery   = "objectTypeId = " + $otId
    $startAt    = 0
    $maxResults = 50
    $isLast     = $false
    $otCount    = 0

    Write-Info ("  Requete AQL : " + $aqlQuery)

    while (-not $isLast) {
        $url      = $assetsBaseUrl + "?startAt=" + $startAt + "&maxResults=" + $maxResults + "&includeAttributes=true"
        $bodyObj  = @{ qlQuery = $aqlQuery }
        $bodyJson = $bodyObj | ConvertTo-Json -Depth 3

        try {
            $resp = Invoke-ApiPostUtf8 -Url $url -Headers $jiraHeaders -JsonBody $bodyJson
        } catch {
            Write-ErrLog ("Erreur Assets AQL OT=" + $otId + " startAt=" + $startAt + " : " + $_.Exception.Message)
            break
        }

        if (-not $resp) { break }

        # Dictionnaire attribut id -> nom depuis objectTypeAttributes
        $attrDict = @{}
        if ($resp.objectTypeAttributes) {
            foreach ($ota in $resp.objectTypeAttributes) {
                if ($ota.id -and $ota.name) {
                    $attrDict[[string]$ota.id] = [string]$ota.name
                }
            }
        }

        $objects = $resp.values
        if (-not $objects -or $objects.Count -eq 0) { break }

        foreach ($obj in $objects) {
            if (-not $obj.id) { continue }
            $props = @{}

            foreach ($attr in $obj.attributes) {
                $attrName = $null
                if ($attr.objectTypeAttributeId) {
                    $attrName = $attrDict[[string]$attr.objectTypeAttributeId]
                }
                if (-not $attrName -and $attr.objectTypeAttribute -and $attr.objectTypeAttribute.name) {
                    $attrName = [string]$attr.objectTypeAttribute.name
                }
                if (-not $attrName) { continue }

                $vals = $attr.objectAttributeValues
                if (-not $vals -or $vals.Count -eq 0) { continue }
                $v0 = $vals[0]

                # Compte Jira (id=377) : lire user.key pour l'accountId
                if ($attrName -eq "Compte Jira") {
                    if ($v0.displayValue)           { $props["Compte Jira"]    = [string]$v0.displayValue }
                    if ($v0.user -and $v0.user.key) { $props["Compte Jira ID"] = [string]$v0.user.key }
                    continue
                }

                # Attributs standards
                if ($v0.displayValue) { $props[$attrName] = [string]$v0.displayValue }
                elseif ($v0.value)    { $props[$attrName] = [string]$v0.value }
            }

            # Log du premier objet pour verification
            if ($assetResults.Count -eq 0) {
                Write-Info "  === Premier objet - attributs resolus ==="
                foreach ($k in ($props.Keys | Sort-Object)) {
                    Write-Info ("    [" + $k + "] = [" + $props[$k] + "]")
                }
            }

            $assetResults.Add([pscustomobject]@{
                "ObjectKey"      = if ($obj.objectKey)                              { [string]$obj.objectKey }           else { "" }
                "Label"          = if ($obj.label)                                  { [string]$obj.label }               else { "" }
                "ObjectType"     = if ($obj.objectType -and $obj.objectType.name)   { [string]$obj.objectType.name }     else { "" }
                "Statut"         = if ($props.ContainsKey("Statut"))                { $props["Statut"] }                 else { "" }
                "Compte Jira"    = if ($props.ContainsKey("Compte Jira"))           { $props["Compte Jira"] }            else { "" }
                "Compte Jira ID" = if ($props.ContainsKey("Compte Jira ID"))        { $props["Compte Jira ID"] }         else { "" }
                "Direction"      = if ($props.ContainsKey("Direction"))             { $props["Direction"] }              else { "" }
                "Type ressource" = if ($props.ContainsKey("Type ressource"))        { $props["Type ressource"] }         else { "" }
                "Date Entree"    = if ($props.ContainsKey($dateEntKey))             { $props[$dateEntKey] }              else { "" }
                "Date Sortie"    = if ($props.ContainsKey("Date Sortie"))           { $props["Date Sortie"] }            else {
                                   if ($props.ContainsKey("Date PLD"))              { $props["Date PLD"] }               else { "" } }
                "Motif Sortie"   = if ($props.ContainsKey("Motif Sortie"))          { $props["Motif Sortie"] }           else { "" }
                "Prenom"         = if ($props.ContainsKey($prenomKey))              { $props[$prenomKey] }               else { "" }
                "Nom"            = if ($props.ContainsKey("Nom"))                   { $props["Nom"] }                    else { "" }
                "Matricule"      = if ($props.ContainsKey("Matricule"))             { $props["Matricule"] }              else { "" }
            }) | Out-Null
            $otCount++
        }

        $isLast   = if ($null -ne $resp.isLast) { [bool]$resp.isLast } else { $true }
        $startAt += $maxResults
        Write-Info ("  OT " + $otId + " : " + $otCount + " charges (startAt=" + $startAt + ", isLast=" + $isLast + ")")

        if ($startAt -gt 50000) { Write-Warn "  Pagination anormale, arret."; break }
        Start-Sleep -Milliseconds 150
    }

    Write-Info ("  OT " + $otId + " termine : " + $otCount + " objets")
}

Write-Info ("  Assets recuperes total : " + $assetResults.Count + " (attendu ~1138)")

# Sauvegarde cache
Save-JsonCache -FileName "saisi_temps_asset.json" -Object $assetResults

# Maps Assets par accountId (user.key = Compte Jira ID)
$assetUsers           = $assetResults
$assetByAccountId     = @{}
$assetByAccountIdNorm = @{}

foreach ($u in $assetUsers) {
    $aid = Get-AssetProp $u "Compte Jira ID"
    if ([string]::IsNullOrWhiteSpace($aid)) { continue }
    if (-not $assetByAccountId.ContainsKey($aid)) {
        $assetByAccountId[$aid] = $u
    }
    $norm = Normalize-AccountId $aid
    if (-not [string]::IsNullOrWhiteSpace($norm) -and -not $assetByAccountIdNorm.ContainsKey($norm)) {
        $assetByAccountIdNorm[$norm] = $u
    }
}
Write-Info ("Assets charges   : " + $assetUsers.Count)
Write-Info ("Map brute        : " + $assetByAccountId.Count + " entrees")
Write-Info ("Map normalisee   : " + $assetByAccountIdNorm.Count + " entrees")

# ============================================================
# 7. CHARGEMENT TEMPO TEAMS
# ============================================================
$teamsData    = Load-JsonCache "tempo_teams.json"
$userTeamsMap = @{}

if ($teamsData) {
    Write-Info ("Tempo Teams : " + $teamsData.Count + " equipes")
    foreach ($t in $teamsData) {
        $teamName     = [string]$t.Nom
        $memberIdsStr = [string]$t."Membres ID"
        if ([string]::IsNullOrWhiteSpace($memberIdsStr)) { continue }
        $members = $memberIdsStr -split ",\s*"
        foreach ($memberId in $members) {
            $memberId = $memberId.Trim()
            if ([string]::IsNullOrWhiteSpace($memberId)) { continue }
            if (-not $userTeamsMap.ContainsKey($memberId)) {
                $userTeamsMap[$memberId] = New-Object System.Collections.Generic.List[string]
            }
            $userTeamsMap[$memberId].Add($teamName)
        }
    }
    Write-Info ("Map UserTeams : " + $userTeamsMap.Count + " personnes avec equipe(s)")
} else {
    Write-Warn "Cache tempo_teams.json introuvable - colonne Nom equipe sera vide"
}

# ============================================================
# 8. RECUPERATION COMPTES JIRA (filtre domaines)
# ============================================================
$jiraUsersCacheFile = Join-Path $cacheDir "jira_users.json"
$jiraUsers = $null

if (Test-Path $jiraUsersCacheFile) {
    $jiraUsers = Load-JsonCache "jira_users.json"
    Write-Info ("Comptes Jira depuis cache : " + $jiraUsers.Count)
} else {
    Write-Info "Recuperation des comptes Jira via API..."
    $jiraUsers = New-Object System.Collections.Generic.List[object]
    $startAt = 0; $pageSize = 200
    while ($true) {
        $url = $jiraBaseUrl + "/rest/api/3/users/search?startAt=" + $startAt + "&maxResults=" + $pageSize
        try { $resp = Invoke-ApiGet -Url $url -Headers $jiraHeaders }
        catch {
            Write-ErrLog ("Erreur API Jira (startAt=" + $startAt + ") : " + $_.Exception.Message)
            throw
        }
        if ($resp.Count -eq 0) { break }
        foreach ($u in $resp) {
            if ($u.accountType -eq "atlassian") { $jiraUsers.Add($u) | Out-Null }
        }
        if ($resp.Count -lt $pageSize) { break }
        $startAt += $pageSize
    }
    Write-Info ("Comptes Jira recuperes via API : " + $jiraUsers.Count)
    $jiraUsers | ConvertTo-Json -Depth 5 -Compress | Set-Content -Path $jiraUsersCacheFile -Encoding UTF8
    Write-Info "Cache sauvegarde : jira_users.json"
}

# Map Jira par accountId + filtre domaines
$jiraUserById     = @{}
$jiraUserByIdNorm = @{}
$jiraActifs       = 0
$jiraInactifs     = 0
$jiraHorsDomaine  = 0

foreach ($u in $jiraUsers) {
    $aid = [string]$u.accountId
    if ([string]::IsNullOrWhiteSpace($aid)) { continue }
    $email = [string]$u.emailAddress
    if (-not (Test-AllowedDomain $email)) { $jiraHorsDomaine++; continue }
    $jiraUserById[$aid] = $u
    $norm = Normalize-AccountId $aid
    if (-not [string]::IsNullOrWhiteSpace($norm)) { $jiraUserByIdNorm[$norm] = $u }
    if ($u.active -eq $true) { $jiraActifs++ } else { $jiraInactifs++ }
}

Write-Info ("Comptes Jira apres filtre domaines : " + $jiraActifs + " actifs, " + $jiraInactifs + " inactifs (total: " + $jiraUserById.Count + ")")
Write-Info ("Comptes Jira exclus (hors domaines) : " + $jiraHorsDomaine)

# ============================================================
# 8b. RECUPERATION GROUPES (DSIM + DSIT)
# ============================================================
Write-Info ("=== Recuperation des groupes contenant : " + ($GroupFilter -join ", ") + " ===")

$dsimGroups = New-Object System.Collections.Generic.List[object]

foreach ($filter in $GroupFilter) {
    $groupPickerUrl = $jiraBaseUrl + "/rest/api/3/groups/picker?query=" + $filter + "&maxResults=500"
    try {
        $gpResp = Invoke-ApiGet -Url $groupPickerUrl -Headers $jiraHeaders
        if ($gpResp.groups) {
            foreach ($g in $gpResp.groups) {
                $gName   = [string]$g.name
                $already = $dsimGroups | Where-Object { $_.Name -eq $gName }
                if (-not $already) {
                    $dsimGroups.Add(@{
                        Name    = $gName
                        GroupId = if ($g.groupId) { [string]$g.groupId } else { "" }
                    }) | Out-Null
                }
            }
        }
    } catch {
        Write-ErrLog ("Erreur API groups/picker (filtre=" + $filter + ") : " + $_.Exception.Message)
        throw
    }
}

Write-Info ("  Groupes trouves : " + $dsimGroups.Count)
foreach ($g in ($dsimGroups | Sort-Object { $_.Name })) {
    Write-Info ("    - " + $g.Name)
}
if ($dsimGroups.Count -eq 0) { throw ("Aucun groupe trouve pour les filtres : " + ($GroupFilter -join ", ")) }

$userDsimGroupsMap = @{}
$cTotalMembers     = 0

foreach ($g in $dsimGroups) {
    $groupName        = $g.Name
    $groupNameEncoded = [System.Uri]::EscapeDataString($groupName)
    $startAt          = 0; $pageSize = 200; $groupMemberCount = 0
    Write-Info ("  Chargement membres de '" + $groupName + "'...")
    while ($true) {
        $url = $jiraBaseUrl + "/rest/api/3/group/member?groupname=" + $groupNameEncoded + "&startAt=" + $startAt + "&maxResults=" + $pageSize + "&includeInactiveUsers=true"
        try { $resp = Invoke-ApiGet -Url $url -Headers $jiraHeaders }
        catch { Write-Warn ("  Erreur group/member '" + $groupName + "' : " + $_.Exception.Message); break }
        $members = $resp.values
        if (-not $members -or $members.Count -eq 0) { break }
        foreach ($m in $members) {
            $accountId = [string]$m.accountId
            if ([string]::IsNullOrWhiteSpace($accountId)) { continue }
            if (-not $userDsimGroupsMap.ContainsKey($accountId)) {
                $userDsimGroupsMap[$accountId] = New-Object System.Collections.Generic.List[string]
            }
            if (-not $userDsimGroupsMap[$accountId].Contains($groupName)) {
                $userDsimGroupsMap[$accountId].Add($groupName)
            }
            $groupMemberCount++
        }
        if ($resp.isLast -eq $true) { break }
        $startAt += $pageSize
        Start-Sleep -Milliseconds 100
    }
    $cTotalMembers += $groupMemberCount
    Write-Info ("    " + $groupName + " : " + $groupMemberCount + " membres")
}
Write-Info ("  Personnes uniques dans groupes filtres : " + $userDsimGroupsMap.Count)

# ============================================================
# 8c. FILTRAGE PAR GROUPES
# ============================================================
Write-Info "=== Filtrage par groupes ==="

$jiraUserByIdFiltered = @{}
$cFilteredIn = 0; $cFilteredOut = 0

foreach ($accountId in $jiraUserById.Keys) {
    if ($userDsimGroupsMap.ContainsKey($accountId)) {
        $jiraUserByIdFiltered[$accountId] = $jiraUserById[$accountId]
        $cFilteredIn++
    } else { $cFilteredOut++ }
}
Write-Info ("  Comptes Jira dans groupes filtres : " + $cFilteredIn)
Write-Info ("  Comptes Jira hors groupes filtres : " + $cFilteredOut)

$assetByAccountIdFiltered     = @{}
$assetByAccountIdNormFiltered = @{}
$cAssetIn = 0; $cAssetOut = 0

foreach ($aid in $assetByAccountId.Keys) {
    $norm   = Normalize-AccountId $aid
    $inScope = $userDsimGroupsMap.ContainsKey($aid) -or
               $userDsimGroupsMap.ContainsKey($norm) -or
               (-not $jiraUserById.ContainsKey($aid) -and -not $jiraUserByIdNorm.ContainsKey($norm))

    if ($inScope) {
        $assetByAccountIdFiltered[$aid] = $assetByAccountId[$aid]
        if (-not [string]::IsNullOrWhiteSpace($norm)) {
            $assetByAccountIdNormFiltered[$norm] = $assetByAccountId[$aid]
        }
        $cAssetIn++
    } else { $cAssetOut++ }
}
Write-Info ("  Assets inclus (perimetre filtre) : " + $cAssetIn)
Write-Info ("  Assets exclus (hors perimetre)   : " + $cAssetOut)

# ============================================================
# 9. HELPERS CROISEMENT
# ============================================================
function Find-Asset([string]$AccountId) {
    if ($assetByAccountIdFiltered.ContainsKey($AccountId)) {
        return $assetByAccountIdFiltered[$AccountId]
    }
    $norm = Normalize-AccountId $AccountId
    if ($assetByAccountIdNormFiltered.ContainsKey($norm)) {
        return $assetByAccountIdNormFiltered[$norm]
    }
    return $null
}

function Get-UserTeamNames([string]$AccountId) {
    if ($userTeamsMap.ContainsKey($AccountId) -and $userTeamsMap[$AccountId].Count -gt 0) {
        return ($userTeamsMap[$AccountId] -join ", ")
    }
    return ""
}

function Get-UserDsimGroups([string]$AccountId) {
    if ($userDsimGroupsMap.ContainsKey($AccountId) -and $userDsimGroupsMap[$AccountId].Count -gt 0) {
        return (($userDsimGroupsMap[$AccountId] | Sort-Object) -join " | ")
    }
    return ""
}

function New-AnomalyRow {
    param([string]$AccountId, [string]$DisplayName, [string]$Email,
          [bool]$JiraActif, [string]$AssetStatut, [string]$Anomalie, $Asset)

    $direction   = ""
    $typeRes     = ""
    $dateEntree  = ""
    $dateSortie  = ""
    $motifSortie = ""
    $prenom      = ""
    $nom         = ""
    $matricule   = ""

    if ($Asset) {
        $direction   = Get-AssetProp $Asset "Direction"
        $typeRes     = Get-AssetProp $Asset "Type ressource"
        $dateEntree  = Get-AssetProp $Asset "Date Entree"
        $dateSortie  = Get-AssetProp $Asset "Date Sortie"
        $motifSortie = Get-AssetProp $Asset "Motif Sortie"
        $prenom      = Get-AssetProp $Asset "Prenom"
        $nom         = Get-AssetProp $Asset "Nom"
        $matricule   = Get-AssetProp $Asset "Matricule"
    }

    return [pscustomobject]@{
        "Direction"      = $direction
        "Nom equipe"     = Get-UserTeamNames  -AccountId $AccountId
        "Groupes"        = Get-UserDsimGroups -AccountId $AccountId
        "Type ressource" = $typeRes
        "Date Entree"    = $dateEntree
        "Date Sortie"    = $dateSortie
        "Motif Sortie"   = $motifSortie
        "Prenom"         = $prenom
        "Nom"            = $nom
        "Matricule"      = $matricule
        "Display Name"   = $DisplayName
        "Email"          = $Email
        "Account ID"     = $AccountId
        "Jira Actif"     = if ($JiraActif) { "OUI" } else { "NON" }
        "Asset Statut"   = if ([string]::IsNullOrWhiteSpace($AssetStatut)) { "AUCUN" } else { $AssetStatut }
        "Anomalie"       = $Anomalie
    }
}

# ============================================================
# 10. CROISEMENT - Detection des anomalies
# ============================================================
Write-Info "=== Croisement Jira <-> Assets ==="

$anomalies1 = New-Object System.Collections.Generic.List[object]
$anomalies2 = New-Object System.Collections.Generic.List[object]

foreach ($accountId in $jiraUserByIdFiltered.Keys) {
    $jiraUser  = $jiraUserByIdFiltered[$accountId]
    $isActive  = ($jiraUser.active -eq $true)
    $dispName  = [string]$jiraUser.displayName
    $email     = [string]$jiraUser.emailAddress
    $asset     = Find-Asset -AccountId $accountId
    $hasAsset  = ($null -ne $asset)
    $assetStat = if ($hasAsset) { Get-AssetProp $asset "Statut" } else { "" }

    if ($isActive) {
        # CAS 1a : Jira actif, aucune fiche Asset
        if (-not $hasAsset) {
            $anomalies1.Add((New-AnomalyRow `
                -AccountId   $accountId `
                -DisplayName $dispName `
                -Email       $email `
                -JiraActif   $true `
                -AssetStatut "" `
                -Anomalie    "Jira actif - Aucune fiche Asset" `
                -Asset       $null)) | Out-Null
        }
        # CAS 1b : Jira actif, fiche Asset presente mais statut != Actif
        elseif ($assetStat -ne "Actif") {
            $anomalies1.Add((New-AnomalyRow `
                -AccountId   $accountId `
                -DisplayName $dispName `
                -Email       $email `
                -JiraActif   $true `
                -AssetStatut $assetStat `
                -Anomalie    ("Jira actif - Asset " + $assetStat) `
                -Asset       $asset)) | Out-Null
        }
    } else {
        # CAS 2a : Jira inactif, fiche Asset Actif
        if ($hasAsset -and $assetStat -eq "Actif") {
            $anomalies2.Add((New-AnomalyRow `
                -AccountId   $accountId `
                -DisplayName $dispName `
                -Email       $email `
                -JiraActif   $false `
                -AssetStatut $assetStat `
                -Anomalie    "Jira inactif - Asset Actif" `
                -Asset       $asset)) | Out-Null
        }
    }
}

# CAS 2b : Asset Actif dont le compte Jira est totalement absent de Jira
foreach ($aid in $assetByAccountIdFiltered.Keys) {
    $norm   = Normalize-AccountId $aid
    $inJira = $jiraUserById.ContainsKey($aid) -or $jiraUserByIdNorm.ContainsKey($norm)
    if ($inJira) { continue }
    $asset     = $assetByAccountIdFiltered[$aid]
    $assetStat = Get-AssetProp $asset "Statut"
    if ($assetStat -ne "Actif") { continue }
    $prenom = Get-AssetProp $asset "Prenom"
    $nom    = Get-AssetProp $asset "Nom"
    $anomalies2.Add((New-AnomalyRow `
        -AccountId   $aid `
        -DisplayName ($prenom + " " + $nom).Trim() `
        -Email       "" `
        -JiraActif   $false `
        -AssetStatut $assetStat `
        -Anomalie    "Compte Jira introuvable - Asset Actif" `
        -Asset       $asset)) | Out-Null
}

Write-Info "=== Anomalies detectees ==="
Write-Info ("  CAS 1 (Jira actif, asset absent/inactif)  : " + $anomalies1.Count)
Write-Info ("  CAS 2 (Jira inactif/absent, asset actif)  : " + $anomalies2.Count)

# ============================================================
# 11. TRI ET EXPORT CSV
# ============================================================
$anomalies1Sorted = $anomalies1 | Sort-Object "Direction", "Nom equipe", "Type ressource", "Date Entree"
$anomalies2Sorted = $anomalies2 | Sort-Object "Direction", "Nom equipe", "Type ressource", "Date Entree"

$headers = @(
    "Direction", "Nom equipe", "Groupes",
    "Type ressource", "Date Entree", "Date Sortie", "Motif Sortie",
    "Prenom", "Nom", "Matricule",
    "Display Name", "Email", "Account ID",
    "Jira Actif", "Asset Statut", "Anomalie"
)

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

$csv1 = Join-Path $exportsDir ("Anomalies - Jira Actif Sans Asset_" + $runStamp + ".csv")
$csv2 = Join-Path $exportsDir ("Anomalies - Jira Inactif Asset Actif_" + $runStamp + ".csv")

Export-CsvStrict -Path $csv1 -Headers $headers -Rows $anomalies1Sorted
Export-CsvStrict -Path $csv2 -Headers $headers -Rows $anomalies2Sorted

# ============================================================
# 12. RESUMES PAR DIRECTION
# ============================================================
Write-Info "=== Resume CAS 1 - Jira Actif / Asset Absent ou Inactif ==="
$grp1 = $anomalies1Sorted | Group-Object "Direction"
foreach ($g in ($grp1 | Sort-Object Name)) {
    $dirName      = if ([string]::IsNullOrWhiteSpace($g.Name)) { "(non defini)" } else { $g.Name }
    $sansAsset    = ($g.Group | Where-Object { $_."Asset Statut" -eq "AUCUN" }).Count
    $assetInactif = $g.Group.Count - $sansAsset
    Write-Info ("  " + $dirName + " : " + $g.Group.Count + " (" + $sansAsset + " sans asset, " + $assetInactif + " asset inactif)")
}

Write-Info "=== Resume CAS 2 - Jira Inactif ou Absent / Asset Actif ==="
$grp2 = $anomalies2Sorted | Group-Object "Direction"
foreach ($g in ($grp2 | Sort-Object Name)) {
    $dirName = if ([string]::IsNullOrWhiteSpace($g.Name)) { "(non defini)" } else { $g.Name }
    Write-Info ("  " + $dirName + " : " + $g.Group.Count)
}

# ============================================================
# 13. RESUME FINAL
# ============================================================
Write-Info "=== Resume Final ==="
Write-Info ("  Filtres groupes                          : " + ($GroupFilter -join ", "))
Write-Info ("  Filtre domaines                          : " + ($allowedDomains -join ", "))
Write-Info ("  Groupes trouves                          : " + $dsimGroups.Count)
Write-Info ("  Comptes Jira total (domaines autorises)  : " + $jiraUserById.Count)
Write-Info ("  Comptes Jira exclus (hors domaines)      : " + $jiraHorsDomaine)
Write-Info ("  Comptes Jira dans groupes filtres        : " + $cFilteredIn)
Write-Info ("  Comptes Jira hors groupes filtres        : " + $cFilteredOut)
Write-Info ("  Assets total (gateway)                   : " + $assetResults.Count)
Write-Info ("  Assets dans perimetre filtre             : " + $cAssetIn)
Write-Info ("  Assets hors perimetre filtre             : " + $cAssetOut)
Write-Info ("  CAS 1 - Jira actif, asset absent/inactif : " + $anomalies1.Count)
Write-Info ("  CAS 2 - Jira inactif/absent, asset actif : " + $anomalies2.Count)
Write-Info ("CSV    : " + $exportsDir)
Write-Info ("Cache  : " + $cacheDir)
Write-Info ("Log    : " + $logFile)
Write-Log ("=== FIN " + $scriptName + " ===") "INFO"