<#
Audit-AutomationsEtAuteursAssets.ps1
Audit des automations Jira (A4J) activées, de leur complexité (étapes),
de leur scope (projets Jira couverts) et de leurs métriques d'exécution.
Enrichissement de l'auteur via :
  - Le schéma Assets Référentiel Personne (Direction, Matricule)
  - L'API Tempo Teams v4 (Équipes Tempo actives séparées par " | ")
  - L'API Jira User Groups (Groupes d'habilitation Jira séparés par " | ")

Mode LECTURE SEULE : aucune modification n'est effectuée.
#>

[CmdletBinding()]
param(
    [string]$JsonInput,
    [switch]$OnlyEnabled = $true,
    [switch]$UseSystemProxy = $true,
    [switch]$ProxyUseDefaultCredentials = $true,
    [string]$ProxyUrl,
    [string]$AssetsWorkspaceId      = "2898703d-5f14-4527-9f8a-ad32b9279ca1",
    [string]$CloudId                = "518657a0-a98f-4c2d-bddb-c6f6039addaa",
    [string]$AssetsPersonSchemaName = "RP"
)

# ============================================================
# 0. DOSSIERS
# ============================================================
$scriptDir  = if ($PSCommandPath) { Split-Path -Parent $PSCommandPath } else { (Get-Location).Path }
$secretsDir = Join-Path $scriptDir "secrets"
$logsDir    = Join-Path $scriptDir "logs"
$exportsDir = Join-Path $scriptDir "exports"

function Ensure-Dir([string]$Path) {
    if (-not (Test-Path $Path)) { New-Item -ItemType Directory -Path $Path -Force | Out-Null }
}
Ensure-Dir $secretsDir; Ensure-Dir $logsDir; Ensure-Dir $exportsDir

$runStamp   = Get-Date -Format "yyyyMMdd_HHmmss"
$scriptName = "Audit-AutomationsEtAuteursAssets"

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
$filtreLabel = if ($OnlyEnabled) { "Activees uniquement (ENABLED)" } else { "Toutes les automations" }
Write-Info ("Filtre statut    : " + $filtreLabel)
Write-Info ("Cloud ID         : " + $CloudId)
Write-Info ("Workspace Assets : " + $AssetsWorkspaceId)

# ============================================================
# 2. PROXY & HELPERS DATES
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

function Parse-DateObj {
    param($val)
    if (-not $val) { return $null }
    $s = [string]$val
    if ($s -match "^\d{13}$") {
        return (Get-Date "1970-01-01").AddMilliseconds([long]$s)
    } elseif ($s -match "^\d{10}$") {
        return (Get-Date "1970-01-01").AddSeconds([long]$s)
    } else {
        try { return [datetime]::Parse($s) } catch { return $null }
    }
}

function Format-DateString {
    param($val)
    $dt = Parse-DateObj $val
    if ($dt) { return $dt.ToString("yyyy-MM-dd HH:mm:ss") }
    $s = [string]$val
    if (-not [string]::IsNullOrWhiteSpace($s)) { return $s }
    return "(Non disponible)"
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
        Write-ErrLog ("POST " + $Url + " : " + $_.Exception.Message + "`n" + $body)
        throw
    }
}

function Test-Endpoint {
    param([string]$Url, [hashtable]$Headers)
    $params = @{ Method='GET'; Uri=$Url; Headers=$Headers; ContentType='application/json'; ErrorAction='SilentlyContinue' }
    $px = Get-EffectiveProxyUri -TargetUrl $Url
    if ($px) { $params.Proxy=$px; $params.ProxyUseDefaultCredentials=$true }
    try {
        $r = Invoke-RestMethod @params
        return ($null -ne $r)
    } catch { return $false }
}

# ============================================================
# 4. CREDENTIALS JIRA & TEMPO
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

function ConvertFrom-SecureStringToPlain {
    param($Secure)
    if (-not $Secure) { return "" }
    if ($Secure -is [System.Security.SecureString]) {
        $bstr = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($Secure)
        try { return [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($bstr) }
        finally { [System.Runtime.InteropServices.Marshal]::ZeroFreeBSTR($bstr) }
    } else { return [string]$Secure }
}

$tempoToken = $null
$tempoTokenFile = Join-Path $secretsDir "tempo-token.xml"
if (Test-Path $tempoTokenFile) {
    try {
        $obj = Import-Clixml -Path $tempoTokenFile
        if ($obj -and $obj.Token) { $tempoToken = ConvertFrom-SecureStringToPlain $obj.Token }
    } catch {}
}

$assetsAqlUrl           = $jiraBaseUrl + "/gateway/api/jsm/assets/workspace/" + $AssetsWorkspaceId + "/v1/object/aql"
$assetsSchemaUrl        = $jiraBaseUrl + "/gateway/api/jsm/assets/workspace/" + $AssetsWorkspaceId + "/v1/objectschema/list"
$jiraUserUrl            = $jiraBaseUrl + "/rest/api/3/user"
$automationJsonFallback = Join-Path $scriptDir "automations_export.json"

# ============================================================
# 4b. CARTOGRAPHIE DES PROJETS JIRA (ID -> KEY)
# ============================================================
Write-Info "=== Resolution des Projets Jira (Mapping ID -> Key) ==="

$jiraProjectsMap = @{} # ID -> Key

try {
    $allProjects = Invoke-ApiGet -Url ($jiraBaseUrl + "/rest/api/3/project") -Headers $jiraHeaders
    if ($allProjects) {
        foreach ($p in $allProjects) {
            if ($p.id -and $p.key) {
                $jiraProjectsMap[[string]$p.id] = [string]$p.key
            }
        }
    }
    Write-Info ("  Projets Jira charges : " + $jiraProjectsMap.Count)
} catch {
    Write-Warn ("  Impossible de charger la liste des projets Jira : " + $_.Exception.Message)
}

function Get-RuleScopeText {
    param($rObj, $projectMap)

    $scopes = New-Object System.Collections.Generic.List[string]

    if ($rObj.projects) {
        foreach ($p in $rObj.projects) {
            if ($p.projectId) {
                $projId = [string]$p.projectId
                if ($projectMap -and $projectMap.ContainsKey($projId)) {
                    $k = $projectMap[$projId]
                    if (-not $scopes.Contains($k)) { $scopes.Add($k) | Out-Null }
                } else {
                    $lbl = "ID:" + $projId
                    if (-not $scopes.Contains($lbl)) { $scopes.Add($lbl) | Out-Null }
                }
            } elseif ($p.projectType) {
                $lbl = "Type:" + $p.projectType
                if (-not $scopes.Contains($lbl)) { $scopes.Add($lbl) | Out-Null }
            }
        }
    }

    if ($scopes.Count -eq 0 -and $rObj.ruleScopeARIs) {
        foreach ($ari in $rObj.ruleScopeARIs) {
            $ariStr = [string]$ari
            if ($ariStr -match "project/(\d+)$") {
                $projId = $Matches[1]
                if ($projectMap -and $projectMap.ContainsKey($projId)) {
                    $k = $projectMap[$projId]
                    if (-not $scopes.Contains($k)) { $scopes.Add($k) | Out-Null }
                } else {
                    $lbl = "ID:" + $projId
                    if (-not $scopes.Contains($lbl)) { $scopes.Add($lbl) | Out-Null }
                }
            } else {
                if (-not $scopes.Contains($ariStr)) { $scopes.Add($ariStr) | Out-Null }
            }
        }
    }

    if ($scopes.Count -gt 0) { return ($scopes -join " | ") }
    return "Global (Tous les projets)"
}

# ============================================================
# 4c. DETECTION DES ENDPOINTS AUTOMATION & FALLBACK JSON
# ============================================================
Write-Info "=== Detection des sources Automation ==="

$url1 = $jiraBaseUrl + "/rest/cb-automation/latest/rules"
$url2 = $jiraBaseUrl + "/rest/cb-automation/latest/rule/list"
$automationRulesUrl = $null

foreach ($cand in [string[]]@($url1, $url2)) {
    if (Test-Endpoint -Url ($cand + "?limit=1") -Headers $jiraHeaders) {
        $automationRulesUrl = $cand
        Write-Info ("  -> Endpoint Automation actif : " + $automationRulesUrl)
        break
    }
}

if (-not $automationRulesUrl) {
    Write-Info "  API Automation directe non accessible — mode FALLBACK JSON active."

    if (-not [string]::IsNullOrWhiteSpace($JsonInput) -and (Test-Path $JsonInput)) {
        $automationJsonFallback = $JsonInput
    }

    if (-not (Test-Path $automationJsonFallback)) {
        $userDownloads = Join-Path ([Environment]::GetFolderPath("UserProfile")) "Downloads"
        $candidateFiles = @(
            (Join-Path $scriptDir "automations_export.json"),
            (Join-Path $scriptDir "automations_export.json.json"),
            (Join-Path $userDownloads "automations_export.json"),
            (Join-Path $userDownloads "automations_export.json.json")
        )
        foreach ($cFile in $candidateFiles) {
            if (Test-Path $cFile) { $automationJsonFallback = $cFile; break }
        }
    }

    if (-not $automationJsonFallback -or -not (Test-Path $automationJsonFallback)) {
        Write-ErrLog "Aucun fichier JSON d'export trouve."
        throw "Fichier automations_export.json introuvable."
    }

    Write-Info ("  Fichier JSON retenu : " + $automationJsonFallback)
}

# ============================================================
# 5. PARSING DES AUTOMATIONS
# ============================================================
Write-Info "=== Recuperation de la liste des Automations ==="

$allRules = New-Object System.Collections.Generic.List[object]

function Parse-RuleSummary {
    param($rSummary, $projectMap)

    $state    = if ($rSummary.state) { [string]$rSummary.state } else { "UNKNOWN" }
    $ruleUuid = if ($rSummary.uuid)  { [string]$rSummary.uuid }  else { [string]$rSummary.id }
    $name     = if ($rSummary.name)  { [string]$rSummary.name }  else { "(Sans nom)" }

    $authorId = ""
    if     ($rSummary.authorAccountId) { $authorId = [string]$rSummary.authorAccountId }
    elseif ($rSummary.actorAccountId)  { $authorId = [string]$rSummary.actorAccountId }
    elseif ($rSummary.author)          { $authorId = [string]$rSummary.author }

    # Date de derniere modification
    $rawUpdated = if ($rSummary.updated) { $rSummary.updated } else { $rSummary.created }
    $updatedStr = Format-DateString $rawUpdated

    # Date de derniere execution
    $rawLastExec = if     ($rSummary.lastExecuted)      { $rSummary.lastExecuted }
                   elseif ($rSummary.lastExecutionDate) { $rSummary.lastExecutionDate }
                   elseif ($rSummary.lastExecution)     { $rSummary.lastExecution }
                   else                                 { $null }
    $lastExecStr = Format-DateString $rawLastExec

    # Scope des projets
    $scopeText = Get-RuleScopeText -rObj $rSummary -projectMap $projectMap

    # Nombre d'etapes (composants)
    $componentCount = 0
    $triggerType    = ""
    if ($rSummary.components)                           { $componentCount = $rSummary.components.Count }
    if ($rSummary.trigger -and $rSummary.trigger.type) { $triggerType    = [string]$rSummary.trigger.type }

    # Volume mensuel
    $execMois = 0
    if     ($null -ne $rSummary.executionCount)            { $execMois = [int]$rSummary.executionCount }
    elseif ($rSummary.stats -and $rSummary.stats.executions30d) { $execMois = [int]$rSummary.stats.executions30d }
    elseif ($rSummary.stats -and $rSummary.stats.totalExecutions) { $execMois = [int]$rSummary.stats.totalExecutions }

    return [pscustomobject]@{
        RuleUuid          = $ruleUuid
        RuleName          = $name
        State             = $state
        AuthorAccountId   = $authorId
        LastUpdated       = $updatedStr
        DerniereExecution = $lastExecStr
        ExecutionsMois30j = $execMois
        NbEtapes          = $componentCount
        TriggerType       = $triggerType
        Scope             = $scopeText
    }
}

if ($automationRulesUrl) {
    $offset   = 0
    $pageSize = 200
    $hasMore  = $true

    while ($hasMore) {
        $urlRules  = $automationRulesUrl + "?limit=" + $pageSize + "&offset=" + $offset
        $respRules = Invoke-ApiGet -Url $urlRules -Headers $jiraHeaders

        $ruleList = if     ($respRules.rules)       { $respRules.rules }
                    elseif ($respRules.data)        { $respRules.data }
                    elseif ($respRules.values)      { $respRules.values }
                    elseif ($respRules -is [array]) { $respRules }

        if (-not $ruleList -or $ruleList.Count -eq 0) { break }

        foreach ($rSummary in $ruleList) {
            $parsed = Parse-RuleSummary -rSummary $rSummary -projectMap $jiraProjectsMap
            if ($OnlyEnabled -and $parsed.State -ine "ENABLED") { continue }
            $allRules.Add($parsed) | Out-Null
        }
        $offset += $pageSize
        $hasMore = ($ruleList.Count -eq $pageSize)
    }

} else {
    Write-Info "  Parsing du fichier JSON..."
    $jsonRaw    = Get-Content -Path $automationJsonFallback -Raw -Encoding UTF8
    $jsonExport = $jsonRaw | ConvertFrom-Json

    $ruleList = if     ($jsonExport.rules)       { $jsonExport.rules }
                elseif ($jsonExport.data)        { $jsonExport.data }
                elseif ($jsonExport -is [array]) { $jsonExport }

    if (-not $ruleList) {
        Write-ErrLog "Structure JSON non reconnue. Attendu : { rules: [...] }"
        throw "JSON invalide."
    }

    foreach ($rSummary in $ruleList) {
        $parsed = Parse-RuleSummary -rSummary $rSummary -projectMap $jiraProjectsMap
        if ($OnlyEnabled -and $parsed.State -ine "ENABLED") { continue }
        $allRules.Add($parsed) | Out-Null
    }
}

Write-Info ("  Automations retenues : " + $allRules.Count)

if ($allRules.Count -eq 0) {
    Write-Warn "Aucune automation trouvee correspondant aux criteres."
    return
}

# ============================================================
# 6. CHARGEMENT EN MEMOIRE DU REFERENTIEL PERSONNE (ASSETS)
# ============================================================
Write-Info "=== Resolution du Schema Assets Personne ==="

$personSchemaId   = $null
$personSchemaName = $AssetsPersonSchemaName

try {
    $schemasResp = Invoke-ApiGet -Url $assetsSchemaUrl -Headers $jiraHeaders
    $schemas     = if ($schemasResp.values) { $schemasResp.values } else { $schemasResp.objectSchemas }

    foreach ($sch in $schemas) {
        $sName = if ($sch.name) { [string]$sch.name } else { "" }
        if ($sName -and ($sName.Trim() -ieq $AssetsPersonSchemaName.Trim() -or $sName -ilike ("*" + $AssetsPersonSchemaName + "*") -or $sName -ilike "*Personne*" -or $sName -ilike "*Referentiel*")) {
            $personSchemaId   = [string]$sch.id
            $personSchemaName = $sName
            break
        }
    }
} catch {}

if ($personSchemaId) {
    Write-Info ("  Schema Personne identifie : " + $personSchemaName + " (id=" + $personSchemaId + ")")
}

Write-Info "=== Chargement du Referentiel Personne en memoire ==="

$personByEmail = @{}
$personByName  = @{}
$aqlPersonQuery = if ($personSchemaId) { "objectSchemaId = " + $personSchemaId } else { 'objectSchema = "' + $personSchemaName + '"' }

$startAt     = 0
$maxResults  = 200
$isLast      = $false
$totalLoaded = 0

while (-not $isLast) {
    $urlAql  = $assetsAqlUrl + "?startAt=" + $startAt + "&maxResults=" + $maxResults + "&includeAttributes=true"
    $bodyAql = (@{ qlQuery = $aqlPersonQuery } | ConvertTo-Json -Depth 3)
    $respAql = $null

    try {
        $respAql = Invoke-ApiPostUtf8 -Url $urlAql -Headers $jiraHeaders -JsonBody $bodyAql
    } catch { break }

    if (-not $respAql -or -not $respAql.values -or $respAql.values.Count -eq 0) { break }

    $attrDict = @{}
    if ($respAql.objectTypeAttributes) {
        foreach ($ota in $respAql.objectTypeAttributes) {
            if ($ota.id -and $ota.name) { $attrDict[[string]$ota.id] = [string]$ota.name }
        }
    }

    foreach ($pObj in $respAql.values) {
        if (-not $pObj.id) { continue }
        $labelRP = if ($pObj.label) { [string]$pObj.label } else { "" }

        $props = @{}
        foreach ($attr in $pObj.attributes) {
            $aName = $attrDict[[string]$attr.objectTypeAttributeId]
            if (-not $aName -and $attr.objectTypeAttribute -and $attr.objectTypeAttribute.name) {
                $aName = [string]$attr.objectTypeAttribute.name
            }
            if (-not $aName) { continue }

            $vals = $attr.objectAttributeValues
            if (-not $vals -or $vals.Count -eq 0) { continue }
            $vStr = [string]$vals[0].displayValue
            if ([string]::IsNullOrWhiteSpace($vStr)) { $vStr = [string]$vals[0].value }
            $props[$aName] = $vStr
        }

        $direction = "(Non trouve)"
        $matricule = ""
        $email     = ""

        foreach ($k in $props.Keys) {
            if     ($k -ieq "Direction" -or $k -ilike "*Direction*" -or $k -ieq "DSIM") { $direction = $props[$k] }
            elseif ($k -ieq "Matricule") { $matricule = $props[$k] }
            elseif ($k -ieq "Email" -or $k -ieq "Mail" -or $k -ilike "*Email*") { $email = $props[$k] }
        }

        $personRec = [pscustomobject]@{
            Label     = $labelRP
            Direction = $direction
            Matricule = $matricule
            Email     = $email
        }

        if (-not [string]::IsNullOrWhiteSpace($email)) { $personByEmail[$email.Trim().ToLower()] = $personRec }
        if (-not [string]::IsNullOrWhiteSpace($labelRP)) { $personByName[$labelRP.Trim().ToLower()] = $personRec }
        $totalLoaded++
    }

    $isLast   = if ($null -ne $respAql.isLast) { [bool]$respAql.isLast } else { $true }
    $startAt += $maxResults
    Start-Sleep -Milliseconds 50
}

Write-Info ("  Personnes chargees depuis Assets : " + $totalLoaded)

# ============================================================
# 6b. CHARGEMENT DES EQUIPES TEMPO (VIA API TEMPO V4 OU FALLBACK)
# ============================================================
Write-Info "=== Resolution des Equipes Tempo (Tempo Teams) ==="

$tempoUserTeamsMap = @{}

if (-not [string]::IsNullOrWhiteSpace($tempoToken)) {
    Write-Info "  Chargement via API Tempo v4 (Bearer token)..."
    $tempoApiHeaders = @{ Authorization = "Bearer " + $tempoToken; Accept = "application/json" }
    try {
        $teamsUrl = "https://api.tempo.io/4/teams"
        while ($teamsUrl) {
            $respTeams = Invoke-ApiGet -Url $teamsUrl -Headers $tempoApiHeaders
            $tResults  = if ($respTeams.results) { $respTeams.results } else { @() }

            foreach ($t in $tResults) {
                $tName = if ($t.name) { [string]$t.name } else { "" }
                $tSelf = if ($t.self) { [string]$t.self } else { "" }
                if (-not $tSelf -or -not $tName) { continue }

                try {
                    $mResp    = Invoke-ApiGet -Url ($tSelf + "/members") -Headers $tempoApiHeaders
                    $mResults = if ($mResp.results) { $mResp.results } else { @() }

                    foreach ($m in $mResults) {
                        $mAccId = if ($m.member -and $m.member.accountId) { [string]$m.member.accountId } else { "" }
                        if (-not $mAccId) { continue }

                        $todayStr = (Get-Date).ToString("yyyy-MM-dd")
                        $toStr    = if ($m.to) { [string]$m.to } else { "" }
                        if ($toStr -and $toStr -lt $todayStr) { continue }

                        if (-not $tempoUserTeamsMap.ContainsKey($mAccId)) {
                            $tempoUserTeamsMap[$mAccId] = New-Object System.Collections.Generic.List[string]
                        }
                        if (-not $tempoUserTeamsMap[$mAccId].Contains($tName)) {
                            $tempoUserTeamsMap[$mAccId].Add($tName) | Out-Null
                        }
                    }
                } catch {}
            }

            if ($respTeams.metadata -and $respTeams.metadata.next) {
                $teamsUrl = $respTeams.metadata.next
            } else { $teamsUrl = $null }
        }
        Write-Info ("  Equipes Tempo v4 chargees pour " + $tempoUserTeamsMap.Count + " utilisateurs")
    } catch {
        Write-Warn ("  Erreur lors de l'appel a l'API Tempo v4 : " + $_.Exception.Message)
    }
}

function Get-TempoTeamsFallback {
    param([string]$AccountId)
    if ([string]::IsNullOrWhiteSpace($AccountId)) { return @() }
    $activeTeams = New-Object System.Collections.Generic.List[string]

    $eps = @(
        "$jiraBaseUrl/rest/tempo-teams/1/member/$AccountId/teams",
        "$jiraBaseUrl/rest/tempo-teams/2/member/$AccountId/teams"
    )
    foreach ($ep in $eps) {
        if (Test-Endpoint -Url $ep -Headers $jiraHeaders) {
            try {
                $resp = Invoke-ApiGet -Url $ep -Headers $jiraHeaders
                $teamsList = if ($resp -is [array]) { $resp } elseif ($resp.teams) { $resp.teams } else { $null }
                if ($teamsList) {
                    foreach ($t in $teamsList) {
                        $tName = if ($t.name) { [string]$t.name } elseif ($t.team -and $t.team.name) { [string]$t.team.name } else { "" }
                        if (-not [string]::IsNullOrWhiteSpace($tName) -and -not $activeTeams.Contains($tName)) {
                            $activeTeams.Add($tName) | Out-Null
                        }
                    }
                }
                if ($activeTeams.Count -gt 0) { break }
            } catch {}
        }
    }
    return $activeTeams
}

# ============================================================
# 7. RESOLUTION DES AUTEURS (JIRA + ASSETS + TEMPO + GROUPES JIRA)
# ============================================================
Write-Info "=== Resolution des Auteurs (Jira + Assets + Tempo + Groupes Jira) ==="

$authorMap = @{}

$distinctAuthorIds = @($allRules |
    ForEach-Object { $_.AuthorAccountId } |
    Where-Object   { -not [string]::IsNullOrWhiteSpace($_) } |
    Select-Object  -Unique)

Write-Info ("  Auteurs distincts a resoudre : " + $distinctAuthorIds.Count)

foreach ($accId in $distinctAuthorIds) {
    $displayName = "Inconnu/Systeme"
    $email       = ""

    # Profil Jira
    try {
        $uObj = Invoke-ApiGet -Url ($jiraUserUrl + "?accountId=" + $accId) -Headers $jiraHeaders
        if ($uObj) {
            if ($uObj.displayName)  { $displayName = [string]$uObj.displayName }
            if ($uObj.emailAddress) { $email       = [string]$uObj.emailAddress }
        }
    } catch {
        Write-Warn ("  Profil Jira non accessible pour accountId : " + $accId)
    }

    # Direction dans Assets
    $direction = "(Non trouve)"
    $matricule = ""
    $labelRP   = ""

    $matched = $null
    if (-not [string]::IsNullOrWhiteSpace($email) -and $personByEmail.ContainsKey($email.Trim().ToLower())) {
        $matched = $personByEmail[$email.Trim().ToLower()]
    } elseif (-not [string]::IsNullOrWhiteSpace($displayName) -and $personByName.ContainsKey($displayName.Trim().ToLower())) {
        $matched = $personByName[$displayName.Trim().ToLower()]
    }

    if ($matched) {
        $direction = $matched.Direction
        $matricule = $matched.Matricule
        $labelRP   = $matched.Label
    } elseif ($displayName -ilike "*Automation*" -or $displayName -ilike "*System*" -or $displayName -ilike "*Addon*") {
        $direction = "(App / Systeme)"
    }

    # Équipes actives Tempo
    $equipesTempoStr = "(Non trouve)"
    if ($displayName -ilike "*Automation*" -or $displayName -ilike "*System*" -or $displayName -ilike "*Addon*") {
        $equipesTempoStr = "(App / Systeme)"
    } elseif ($tempoUserTeamsMap.ContainsKey($accId)) {
        $equipesTempoStr = ($tempoUserTeamsMap[$accId] -join " | ")
    } else {
        $fbTeams = Get-TempoTeamsFallback -AccountId $accId
        if ($fbTeams -and $fbTeams.Count -gt 0) {
            $equipesTempoStr = ($fbTeams -join " | ")
        }
    }

    # Groupes d'habilitation Jira
    $groupesJiraStr = "(Non trouve)"
    if ($displayName -ilike "*Automation*" -or $displayName -ilike "*System*" -or $displayName -ilike "*Addon*") {
        $groupesJiraStr = "(App / Systeme)"
    } else {
        try {
            $urlGroups  = $jiraBaseUrl + "/rest/api/3/user/groups?accountId=" + $accId
            $respGroups = Invoke-ApiGet -Url $urlGroups -Headers $jiraHeaders
            $gList      = if ($respGroups -is [array]) { $respGroups } elseif ($respGroups.values) { $respGroups.values } else { $null }

            if ($gList) {
                $gNames = @()
                foreach ($g in $gList) {
                    if ($g.name) { $gNames += [string]$g.name }
                }
                if ($gNames.Count -gt 0) {
                    $groupesJiraStr = ($gNames -join " | ")
                }
            }
        } catch {}
    }

    Write-Info ("  " + $displayName + " (" + $email + ") -> Direction: " + $direction + " | Tempo: " + $equipesTempoStr + " | Groupes: " + $groupesJiraStr)

    $authorMap[$accId] = [pscustomobject]@{
        AccountId    = $accId
        DisplayName  = $displayName
        Email        = $email
        Direction    = $direction
        EquipesTempo = $equipesTempoStr
        GroupesJira  = $groupesJiraStr
        Matricule    = $matricule
        LabelRP      = $labelRP
    }
}

# ============================================================
# 8. CONSOLIDATION ET EXPORTS CSV
# ============================================================
Write-Info "=== Consolidation des donnees ==="

$detailedRows = New-Object System.Collections.Generic.List[object]

foreach ($rule in $allRules) {
    $author = $authorMap[$rule.AuthorAccountId]

    $authorName   = if ($author) { $author.DisplayName }  else { "Inconnu" }
    $authorMail   = if ($author) { $author.Email }        else { "" }
    $dirName      = if ($author) { $author.Direction }    else { "(Inconnu)" }
    $eqTempo      = if ($author) { $author.EquipesTempo } else { "(Inconnu)" }
    $grpJira      = if ($author) { $author.GroupesJira }  else { "(Inconnu)" }
    $matr         = if ($author) { $author.Matricule }    else { "" }

    $detailedRows.Add([pscustomobject]@{
        RuleName        = $rule.RuleName
        State           = $rule.State
        ScopeProjets    = $rule.Scope
        NbEtapes        = $rule.NbEtapes
        LastUpdated     = $rule.LastUpdated
        AuthorName      = $authorName
        AuthorEmail     = $authorMail
        AuthorMatricule = $matr
        DirectionAssets = $dirName
        EquipesTempo    = $eqTempo
        GroupesJira     = $grpJira
        TriggerType     = $rule.TriggerType
        RuleUuid        = $rule.RuleUuid
        AuthorAccountId = $rule.AuthorAccountId
    }) | Out-Null
}

Write-Info "=== Export des rapports CSV ==="

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
    Write-Info ("  CSV : " + $Path + " (" + $Rows.Count + " lignes)")
}

# 8a. Liste detaillee — triee par Direction puis par Auteur
$csvDetail = Join-Path $exportsDir ("Automations-ListeDetaillee_" + $runStamp + ".csv")
Export-CsvStrict -Path $csvDetail `
    -Headers @("DirectionAssets","AuthorName","AuthorEmail","AuthorMatricule",
               "EquipesTempo","GroupesJira","RuleName","State","ScopeProjets",
               "NbEtapes","LastUpdated","TriggerType","RuleUuid","AuthorAccountId") `
    -Rows ($detailedRows | Sort-Object DirectionAssets, AuthorName)

# 8b. Synthese par Direction
$grpDirection = $detailedRows | Group-Object DirectionAssets | Sort-Object Name
$synthDirRows = foreach ($g in $grpDirection) {
    $auteursDir = @($g.Group | ForEach-Object { $_.AuthorEmail } | Select-Object -Unique).Count
    [pscustomobject]@{
        DirectionAssets    = $g.Name
        NbAutomations      = $g.Count
        NbAuteursDistincts = $auteursDir
    }
}
$csvSynthDir = Join-Path $exportsDir ("Automations-SyntheseParDirection_" + $runStamp + ".csv")
Export-CsvStrict -Path $csvSynthDir `
    -Headers @("DirectionAssets","NbAutomations","NbAuteursDistincts") `
    -Rows $synthDirRows

# 8c. Synthese par Auteur — triee par Direction puis par Auteur
$grpAuteur = $detailedRows | Group-Object AuthorAccountId
$synthAuteurRows = foreach ($g in $grpAuteur) {
    $a0 = $g.Group[0]
    [pscustomobject]@{
        AuthorName      = $a0.AuthorName
        AuthorEmail     = $a0.AuthorEmail
        DirectionAssets = $a0.DirectionAssets
        EquipesTempo    = $a0.EquipesTempo
        GroupesJira     = $a0.GroupesJira
        NbAutomations   = $g.Count
    }
}
$csvSynthAuteur = Join-Path $exportsDir ("Automations-SyntheseParAuteur_" + $runStamp + ".csv")
Export-CsvStrict -Path $csvSynthAuteur `
    -Headers @("DirectionAssets","AuthorName","AuthorEmail","EquipesTempo","GroupesJira","NbAutomations") `
    -Rows ($synthAuteurRows | Sort-Object DirectionAssets, AuthorName)

# 8d. Synthese par Projet (Scope) — triee par nom de scope
$grpScope = $detailedRows | Group-Object ScopeProjets | Sort-Object Name
$synthScopeRows = foreach ($g in $grpScope) {
    $auteursScope = @($g.Group | ForEach-Object { $_.AuthorEmail } | Select-Object -Unique).Count
    [pscustomobject]@{
        ScopeProjets       = $g.Name
        NbAutomations      = $g.Count
        NbAuteursDistincts = $auteursScope
    }
}
$csvSynthScope = Join-Path $exportsDir ("Automations-SyntheseParScope_" + $runStamp + ".csv")
Export-CsvStrict -Path $csvSynthScope `
    -Headers @("ScopeProjets","NbAutomations","NbAuteursDistincts") `
    -Rows $synthScopeRows

# ============================================================
# 9. RESUME FINAL
# ============================================================
$modeSource = if ($automationRulesUrl) { "API cb-automation" } else { "JSON export (" + (Split-Path $automationJsonFallback -Leaf) + ")" }

Write-Info ""
Write-Info "============================================="
Write-Info "=== RESUME AUDIT AUTOMATIONS ET AUTEURS ==="
Write-Info "============================================="
Write-Info ""
Write-Info ("  Source des donnees            : " + $modeSource)
Write-Info ("  Total automations analysees   : " + $detailedRows.Count)
Write-Info ("  Auteurs distincts identifies  : " + $grpAuteur.Count)
Write-Info ("  Directions distinctes         : " + $grpDirection.Count)
Write-Info ("  Projets/Scopes distincts      : " + $grpScope.Count)
Write-Info ""
Write-Info "  Top 3 Directions contributrices :"
foreach ($d in ($synthDirRows | Sort-Object NbAutomations -Descending | Select-Object -First 3)) {
    Write-Info ("    - " + $d.DirectionAssets + " : " + $d.NbAutomations + " automation(s)")
}
Write-Info ""
Write-Info "  Top 3 Auteurs :"
foreach ($a in ($synthAuteurRows | Sort-Object NbAutomations -Descending | Select-Object -First 3)) {
    Write-Info ("    - " + $a.AuthorName + " : " + $a.NbAutomations + " automation(s)")
}
Write-Info ""
Write-Info ("  Rapports CSV dans : " + $exportsDir)
Write-Info ("  Log               : " + $logFile)
Write-Log ("=== FIN " + $scriptName + " ===") "INFO"