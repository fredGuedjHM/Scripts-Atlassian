<#
Audit-ParametrageJira.ps1
Identifie les objets de parametrage Jira orphelins (non utilises) :
  - Priorites
  - Systemes de workflows / Workflows
  - Systemes d'ecrans / Ecrans
  - Systemes d'autorisations (Permission Schemes)
  - Roles
Pour chaque objet : indique s'il est utilise, par quels projets (avec categorie),
et si c'est un candidat a la suppression.
Aucune suppression n'est effectuee par ce script.
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
$scriptName = "Audit-ParametrageJira"

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
function Invoke-ApiGet {
    param([Parameter(Mandatory)][string]$Url,
          [Parameter(Mandatory)][hashtable]$Headers)
    $params = @{
        Method      = 'GET'
        Uri         = $Url
        Headers     = $Headers
        ContentType = 'application/json'
        ErrorAction = 'Stop'
    }
    $px = Get-EffectiveProxyUri -TargetUrl $Url
    if ($px) { $params.Proxy = $px; if ($ProxyUseDefaultCredentials) { $params.ProxyUseDefaultCredentials = $true } }
    try { return Invoke-RestMethod @params }
    catch {
        $sc = $null
        try { $sc = $_.Exception.Response.StatusCode.value__ } catch {}
        # Ne pas logger les 404 : ils sont geres par l'appelant
        if ($sc -ne 404) {
            Write-ErrLog ("GET " + $Url + " : " + $_.Exception.Message)
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

function Invoke-ApiPostJson {
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
    if ($px) { $params.Proxy = $px; if ($ProxyUseDefaultCredentials) { $params.ProxyUseDefaultCredentials = $true } }
    try {
        $resp   = Invoke-WebRequest @params
        $stream = $resp.RawContentStream
        $stream.Position = 0
        $reader = New-Object System.IO.StreamReader($stream, [System.Text.Encoding]::UTF8, $true)
        $raw    = $reader.ReadToEnd()
        $reader.Close()
        return $raw | ConvertFrom-Json
    }
    catch {
        Write-ErrLog ("POST " + $Url + " : " + $_.Exception.Message)
        return $null
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
# 5. HELPERS
# ============================================================
function Save-JsonCache {
    param([string]$FileName, $Object)
    $path = Join-Path $cacheDir $FileName
    $Object | ConvertTo-Json -Depth 10 -Compress | Set-Content -Path $path -Encoding UTF8
    $size = [math]::Round((Get-Item $path).Length / 1MB, 2)
    Write-Info ("Cache sauvegarde : " + $FileName + " (" + $size + " MB)")
}

function Paginate-ApiGet {
    param([string]$BaseUrl, [hashtable]$Headers, [int]$PageSize = 50, [string]$Label = "")
    $all     = New-Object System.Collections.Generic.List[object]
    $startAt = 0
    while ($true) {
        $sep  = if ($BaseUrl.Contains("?")) { "&" } else { "?" }
        $url  = $BaseUrl + $sep + "startAt=" + $startAt + "&maxResults=" + $PageSize
        $resp = Invoke-ApiGetSafe -Url $url -Headers $Headers
        if (-not $resp) { break }

        $values = $null
        if     ($resp.values)      { $values = $resp.values }
        elseif ($resp -is [array]) { $values = $resp }

        if (-not $values -or $values.Count -eq 0) { break }
        foreach ($v in $values) { $all.Add($v) | Out-Null }

        $total  = if ($resp.total)            { [int]$resp.total }    else { -1 }
        $isLast = if ($null -ne $resp.isLast) { [bool]$resp.isLast }  else { $false }

        if ($isLast)                                  { break }
        if ($total -gt 0 -and $all.Count -ge $total) { break }
        $startAt += $PageSize
        Start-Sleep -Milliseconds 100
    }
    if ($Label) { Write-Info ("  " + $Label + " : " + $all.Count + " elements") }
    return $all
}

function Search-IssuesPost {
    # jiradot : seul POST /rest/api/3/search/jql fonctionne
    param([Parameter(Mandatory)][string]$Jql,
          [Parameter(Mandatory)][hashtable]$Headers,
          [int]$MaxResults = 1,
          [string[]]$Fields = @("updated"))
    $url  = $jiraBaseUrl + "/rest/api/3/search/jql"
    $body = @{ jql = $Jql; maxResults = $MaxResults; fields = $Fields } | ConvertTo-Json -Depth 3
    $resp = Invoke-ApiPostJson -Url $url -Headers $Headers -JsonBody $body

    # /search/jql retourne "issues" + "total" ou "issues" + "nextPageToken"
    # Normaliser vers un objet avec .total et .issues pour compatibilite
    if ($resp -and $null -eq $resp.total -and $resp.issues) {
        $resp | Add-Member -NotePropertyName "total" -NotePropertyValue $resp.issues.Count -Force
    }
    return $resp
}


function Truncate-List([string[]]$Items, [int]$Max = 20) {
    if (-not $Items -or $Items.Count -eq 0) { return "" }
    if ($Items.Count -le $Max) { return ($Items -join ", ") }
    return (($Items[0..($Max - 1)] -join ", ") + " ... (+" + ($Items.Count - $Max) + " autres)")
}
# ============================================================
# 6. RECUPERATION PROJETS (avec categorie)
# ============================================================
Write-Info "=== Recuperation des projets ==="

$allProjects = Paginate-ApiGet `
    -BaseUrl  ($jiraBaseUrl + "/rest/api/3/project/search?expand=description,lead,projectCategory") `
    -Headers  $jiraHeaders `
    -PageSize 50 `
    -Label    "Projets"

$projectById  = @{}
$projectByKey = @{}

foreach ($proj in $allProjects) {
    $cat = ""
    if ($proj.projectCategory -and $proj.projectCategory.name) { $cat = [string]$proj.projectCategory.name }
    $obj = @{
        Id       = [string]$proj.id
        Key      = [string]$proj.key
        Name     = [string]$proj.name
        Category = $cat
    }
    $projectById[[string]$proj.id]   = $obj
    $projectByKey[[string]$proj.key] = $obj
}
Write-Info ("  Projets charges : " + $projectById.Count)
Save-JsonCache -FileName "audit_projects.json" -Object $allProjects

$projectIdList = @($projectById.Keys)
$batchSize     = 20

# Helper : enrichir projKeys / projCats depuis une liste d'IDs
function Get-ProjKeysAndCats {
    param([string[]]$ProjIds)
    $keys = @(); $cats = @()
    foreach ($projId in $ProjIds) {
        if ($projectById.ContainsKey($projId)) {
            $keys += $projectById[$projId].Key
            $cat   = $projectById[$projId].Category
            if (-not [string]::IsNullOrWhiteSpace($cat) -and $cats -notcontains $cat) { $cats += $cat }
        }
    }
    return @{ Keys = $keys; Cats = $cats }
}

# ============================================================
# 7. AUDIT PRIORITES
# ============================================================
Write-Info "=== Audit Priorites ==="

$priorities = Paginate-ApiGet `
    -BaseUrl  ($jiraBaseUrl + "/rest/api/3/priority/search") `
    -Headers  $jiraHeaders `
    -PageSize 50 `
    -Label    "Priorites"

$priorityResults = New-Object System.Collections.Generic.List[object]

foreach ($prio in $priorities) {
    $prioName = [string]$prio.name
    $prioId   = [string]$prio.id
    Write-Info ("  Priorite : " + $prioName + " (id=" + $prioId + ")")

    $jql        = "priority = " + $prioId + " ORDER BY updated DESC"
    $searchResp = Search-IssuesPost -Jql $jql -Headers $jiraHeaders -MaxResults 1 -Fields @("updated")

    $nbIssues     = 0
    $lastActivity = ""
    if ($searchResp) {
        $nbIssues = if ($searchResp.total) { [int]$searchResp.total } else { 0 }
        if ($searchResp.issues -and $searchResp.issues.Count -gt 0) {
            $upd = [string]$searchResp.issues[0].fields.updated
            if ($upd.Length -ge 10) { $lastActivity = $upd.Substring(0, 10) }
        }
    }

    $candidate = if ($nbIssues -eq 0) { "OUI" } else { "NON" }

    $priorityResults.Add([pscustomobject]@{
        "Nom"                  = $prioName
        "ID"                   = $prioId
        "Nb Issues"            = $nbIssues
        "Derniere Activite"    = $lastActivity
        "Candidat Suppression" = $candidate
    }) | Out-Null

    Start-Sleep -Milliseconds 200
}

Write-Info ("  Priorites totales    : " + $priorityResults.Count)
Write-Info ("  Priorites candidates : " + ($priorityResults | Where-Object { $_."Candidat Suppression" -eq "OUI" }).Count)

# ============================================================
# 8. AUDIT WORKFLOW SCHEMES / WORKFLOWS
# ============================================================
Write-Info "=== Audit Workflow Schemes ==="

$wfSchemes = Paginate-ApiGet `
    -BaseUrl  ($jiraBaseUrl + "/rest/api/3/workflowscheme") `
    -Headers  $jiraHeaders `
    -PageSize 50 `
    -Label    "Workflow Schemes"

# Mapping schemeId -> projets
$wfSchemeProjectMap = @{}

for ($i = 0; $i -lt $projectIdList.Count; $i += $batchSize) {
    $batch   = $projectIdList[$i..([math]::Min($i + $batchSize - 1, $projectIdList.Count - 1))]
    $qparams = ($batch | ForEach-Object { "projectId=" + $_ }) -join "&"
    $url     = $jiraBaseUrl + "/rest/api/3/workflowscheme/project?" + $qparams
    $resp    = Invoke-ApiGetSafe -Url $url -Headers $jiraHeaders
    if ($resp -and $resp.values) {
        foreach ($v in $resp.values) {
            $schemeId = ""
            if ($v.workflowScheme -and $v.workflowScheme.id) { $schemeId = [string]$v.workflowScheme.id }
            if ([string]::IsNullOrWhiteSpace($schemeId)) { continue }
            $projId = [string]$v.projectId
            if (-not $wfSchemeProjectMap.ContainsKey($schemeId)) {
                $wfSchemeProjectMap[$schemeId] = New-Object System.Collections.Generic.List[string]
            }
            $wfSchemeProjectMap[$schemeId].Add($projId)
        }
    }
    Start-Sleep -Milliseconds 100
}

$wfSchemeResults   = New-Object System.Collections.Generic.List[object]
$usedWorkflowNames = @{}
$allWorkflowNames  = @{}

foreach ($wfs in $wfSchemes) {
    $schemeId   = [string]$wfs.id
    $schemeName = [string]$wfs.name
    $projIds    = if ($wfSchemeProjectMap.ContainsKey($schemeId)) { @($wfSchemeProjectMap[$schemeId]) } else { @() }
    $nbProjects = $projIds.Count
    $kc         = Get-ProjKeysAndCats -ProjIds $projIds

    $referencedWfs = @()
    if ($wfs.defaultWorkflow) {
        $wfn = [string]$wfs.defaultWorkflow
        if ($referencedWfs -notcontains $wfn) { $referencedWfs += $wfn }
    }
    if ($wfs.issueTypeMappings) {
        foreach ($mapping in $wfs.issueTypeMappings.PSObject.Properties) {
            if (-not [string]::IsNullOrWhiteSpace($mapping.Value)) {
                $wfn = [string]$mapping.Value
                if ($referencedWfs -notcontains $wfn) { $referencedWfs += $wfn }
            }
        }
    }
    foreach ($wfn in $referencedWfs) {
        $allWorkflowNames[$wfn] = $true
        if ($nbProjects -gt 0) { $usedWorkflowNames[$wfn] = $true }
    }

    $candidate = if ($nbProjects -eq 0) { "OUI" } else { "NON" }

    $wfSchemeResults.Add([pscustomobject]@{
        "Nom"                  = $schemeName
        "ID"                   = $schemeId
        "Nb Projets"           = $nbProjects
        "Projets"              = Truncate-List $kc.Keys
        "Categories"           = ($kc.Cats | Sort-Object) -join ", "
        "Candidat Suppression" = $candidate
    }) | Out-Null
}

Write-Info ("  Workflow Schemes totaux    : " + $wfSchemeResults.Count)
Write-Info ("  Workflow Schemes candidats : " + ($wfSchemeResults | Where-Object { $_."Candidat Suppression" -eq "OUI" }).Count)

# Audit Workflows individuels
Write-Info "=== Audit Workflows ==="

$workflows = Paginate-ApiGet `
    -BaseUrl  ($jiraBaseUrl + "/rest/api/3/workflow/search") `
    -Headers  $jiraHeaders `
    -PageSize 50 `
    -Label    "Workflows"

$workflowResults = New-Object System.Collections.Generic.List[object]

foreach ($wf in $workflows) {
    $wfName = ""
    if     ($wf.id -and $wf.id.name) { $wfName = [string]$wf.id.name }
    elseif ($wf.name)                 { $wfName = [string]$wf.name }
    if ([string]::IsNullOrWhiteSpace($wfName)) { continue }

    $entityId       = if ($wf.id -and $wf.id.entityId) { [string]$wf.id.entityId } else { "" }
    $inActiveScheme = if ($usedWorkflowNames.ContainsKey($wfName)) { "OUI" } else { "NON" }
    $inAnyScheme    = if ($allWorkflowNames.ContainsKey($wfName))  { "OUI" } else { "NON" }
    $candidate      = if ($inActiveScheme -eq "NON") { "OUI" } else { "NON" }

    $workflowResults.Add([pscustomobject]@{
        "Nom"                  = $wfName
        "Entity ID"            = $entityId
        "Dans Scheme Actif"    = $inActiveScheme
        "Dans Scheme (tout)"   = $inAnyScheme
        "Candidat Suppression" = $candidate
    }) | Out-Null
}

Write-Info ("  Workflows totaux    : " + $workflowResults.Count)
Write-Info ("  Workflows candidats : " + ($workflowResults | Where-Object { $_."Candidat Suppression" -eq "OUI" }).Count)

# ============================================================
# 9. AUDIT ISSUE TYPE SCREEN SCHEMES / SCREEN SCHEMES / SCREENS
# ============================================================
Write-Info "=== Audit Issue Type Screen Schemes (ITSS) ==="

$allItss = Paginate-ApiGet `
    -BaseUrl  ($jiraBaseUrl + "/rest/api/3/issuetypescreenscheme") `
    -Headers  $jiraHeaders `
    -PageSize 50 `
    -Label    "Issue Type Screen Schemes"

# Mapping ITSS -> Projets
$itssProjectMap = @{}

for ($i = 0; $i -lt $projectIdList.Count; $i += $batchSize) {
    $batch   = $projectIdList[$i..([math]::Min($i + $batchSize - 1, $projectIdList.Count - 1))]
    $qparams = ($batch | ForEach-Object { "projectId=" + $_ }) -join "&"
    $url     = $jiraBaseUrl + "/rest/api/3/issuetypescreenscheme/project?" + $qparams
    $resp    = Invoke-ApiGetSafe -Url $url -Headers $jiraHeaders
    if ($resp -and $resp.values) {
        foreach ($v in $resp.values) {
            $itssId = ""
            if ($v.issueTypeScreenScheme -and $v.issueTypeScreenScheme.id) {
                $itssId = [string]$v.issueTypeScreenScheme.id
            }
            if ([string]::IsNullOrWhiteSpace($itssId)) { continue }
            if ($v.projectIds) {
                foreach ($projId in $v.projectIds) {
                    if (-not $itssProjectMap.ContainsKey($itssId)) {
                        $itssProjectMap[$itssId] = New-Object System.Collections.Generic.List[string]
                    }
                    $itssProjectMap[$itssId].Add([string]$projId)
                }
            }
        }
    }
    Start-Sleep -Milliseconds 100
}

# Mapping ITSS -> Screen Schemes via mappings
$usedScreenSchemeIds = @{}
$itssIds = @($allItss | ForEach-Object { [string]$_.id })

for ($i = 0; $i -lt $itssIds.Count; $i += 10) {
    $batch   = $itssIds[$i..([math]::Min($i + 9, $itssIds.Count - 1))]
    $qparams = ($batch | ForEach-Object { "issueTypeScreenSchemeId=" + $_ }) -join "&"
    $url     = $jiraBaseUrl + "/rest/api/3/issuetypescreenscheme/mapping?" + $qparams + "&startAt=0&maxResults=500"
    $resp    = Invoke-ApiGetSafe -Url $url -Headers $jiraHeaders
    if ($resp -and $resp.values) {
        foreach ($mapping in $resp.values) {
            $parentItssId = [string]$mapping.issueTypeScreenSchemeId
            $isActive     = $itssProjectMap.ContainsKey($parentItssId)
            $ssId         = [string]$mapping.screenSchemeId
            if ($isActive -and -not [string]::IsNullOrWhiteSpace($ssId)) {
                $usedScreenSchemeIds[$ssId] = $true
            }
        }
    }
    Start-Sleep -Milliseconds 100
}

# Resultats ITSS
$itssResults = New-Object System.Collections.Generic.List[object]

foreach ($itss in $allItss) {
    $itssId     = [string]$itss.id
    $itssName   = [string]$itss.name
    $projIds    = if ($itssProjectMap.ContainsKey($itssId)) { @($itssProjectMap[$itssId]) } else { @() }
    $nbProjects = $projIds.Count
    $kc         = Get-ProjKeysAndCats -ProjIds $projIds
    $candidate  = if ($nbProjects -eq 0) { "OUI" } else { "NON" }

    $itssResults.Add([pscustomobject]@{
        "Nom"                  = $itssName
        "ID"                   = $itssId
        "Nb Projets"           = $nbProjects
        "Projets"              = Truncate-List $kc.Keys
        "Categories"           = ($kc.Cats | Sort-Object) -join ", "
        "Candidat Suppression" = $candidate
    }) | Out-Null
}

Write-Info ("  ITSS totaux    : " + $itssResults.Count)
Write-Info ("  ITSS candidats : " + ($itssResults | Where-Object { $_."Candidat Suppression" -eq "OUI" }).Count)

# Audit Screen Schemes
Write-Info "=== Audit Screen Schemes ==="

$allScreenSchemes    = Paginate-ApiGet `
    -BaseUrl  ($jiraBaseUrl + "/rest/api/3/screenscheme") `
    -Headers  $jiraHeaders `
    -PageSize 50 `
    -Label    "Screen Schemes"

$usedScreenIds       = @{}
$screenSchemeResults = New-Object System.Collections.Generic.List[object]

foreach ($ss in $allScreenSchemes) {
    $ssId   = [string]$ss.id
    $ssName = [string]$ss.name
    $isUsed = $usedScreenSchemeIds.ContainsKey($ssId)

    if ($ss.screens) {
        foreach ($screenProp in $ss.screens.PSObject.Properties) {
            $screenId = [string]$screenProp.Value
            if ($isUsed -and -not [string]::IsNullOrWhiteSpace($screenId)) {
                $usedScreenIds[$screenId] = $true
            }
        }
    }

    $candidate = if (-not $isUsed) { "OUI" } else { "NON" }

    $screenSchemeResults.Add([pscustomobject]@{
        "Nom"                  = $ssName
        "ID"                   = $ssId
        "Utilise"              = if ($isUsed) { "OUI" } else { "NON" }
        "Candidat Suppression" = $candidate
    }) | Out-Null
}

Write-Info ("  Screen Schemes totaux    : " + $screenSchemeResults.Count)
Write-Info ("  Screen Schemes candidats : " + ($screenSchemeResults | Where-Object { $_."Candidat Suppression" -eq "OUI" }).Count)

# Audit Screens
Write-Info "=== Audit Screens ==="

$allScreens    = Paginate-ApiGet `
    -BaseUrl  ($jiraBaseUrl + "/rest/api/3/screens") `
    -Headers  $jiraHeaders `
    -PageSize 100 `
    -Label    "Screens"

$screenResults = New-Object System.Collections.Generic.List[object]

foreach ($scr in $allScreens) {
    $scrId   = [string]$scr.id
    $scrName = [string]$scr.name
    $scope   = ""
    if ($scr.scope -and $scr.scope.type) { $scope = [string]$scr.scope.type }

    $isUsed    = $usedScreenIds.ContainsKey($scrId)
    $candidate = if (-not $isUsed) { "OUI" } else { "NON" }

    $screenResults.Add([pscustomobject]@{
        "Nom"                  = $scrName
        "ID"                   = $scrId
        "Scope"                = $scope
        "Utilise"              = if ($isUsed) { "OUI" } else { "NON" }
        "Candidat Suppression" = $candidate
    }) | Out-Null
}

Write-Info ("  Screens totaux    : " + $screenResults.Count)
Write-Info ("  Screens candidats : " + ($screenResults | Where-Object { $_."Candidat Suppression" -eq "OUI" }).Count)

# ============================================================
# 10. AUDIT PERMISSION SCHEMES
# ============================================================
Write-Info "=== Audit Permission Schemes ==="

$allPermSchemes = @()
$permResp = Invoke-ApiGetSafe -Url ($jiraBaseUrl + "/rest/api/3/permissionscheme") -Headers $jiraHeaders
if ($permResp -and $permResp.permissionSchemes) { $allPermSchemes = $permResp.permissionSchemes }
Write-Info ("  Permission Schemes recuperes : " + $allPermSchemes.Count)

$permSchemeProjectMap = @{}

foreach ($proj in $allProjects) {
    $projKey = [string]$proj.key
    $projId  = [string]$proj.id
    $url     = $jiraBaseUrl + "/rest/api/3/project/" + $projKey + "/permissionscheme"
    $resp    = Invoke-ApiGetSafe -Url $url -Headers $jiraHeaders
    if ($resp -and $resp.id) {
        $psId = [string]$resp.id
        if (-not $permSchemeProjectMap.ContainsKey($psId)) {
            $permSchemeProjectMap[$psId] = New-Object System.Collections.Generic.List[string]
        }
        $permSchemeProjectMap[$psId].Add($projId)
    }
    Start-Sleep -Milliseconds 50
}

$permSchemeResults = New-Object System.Collections.Generic.List[object]

foreach ($ps in $allPermSchemes) {
    $psId       = [string]$ps.id
    $psName     = [string]$ps.name
    $projIds    = if ($permSchemeProjectMap.ContainsKey($psId)) { @($permSchemeProjectMap[$psId]) } else { @() }
    $nbProjects = $projIds.Count
    $kc         = Get-ProjKeysAndCats -ProjIds $projIds
    $candidate  = if ($nbProjects -eq 0) { "OUI" } else { "NON" }

    $permSchemeResults.Add([pscustomobject]@{
        "Nom"                  = $psName
        "ID"                   = $psId
        "Nb Projets"           = $nbProjects
        "Projets"              = Truncate-List $kc.Keys
        "Categories"           = ($kc.Cats | Sort-Object) -join ", "
        "Candidat Suppression" = $candidate
    }) | Out-Null
}

Write-Info ("  Permission Schemes totaux    : " + $permSchemeResults.Count)
Write-Info ("  Permission Schemes candidats : " + ($permSchemeResults | Where-Object { $_."Candidat Suppression" -eq "OUI" }).Count)
# ============================================================
# 11. AUDIT ROLES
# ============================================================
Write-Info "=== Audit Roles ==="

# 11a. Recuperer la liste globale des roles (dedupliquee par ID)
$allRolesRaw = @()
$rolesResp   = Invoke-ApiGetSafe -Url ($jiraBaseUrl + "/rest/api/3/role") -Headers $jiraHeaders
if ($rolesResp) { $allRolesRaw = $rolesResp }

$rolesById = @{}
foreach ($r in $allRolesRaw) {
    $rid = [string]$r.id
    if (-not $rolesById.ContainsKey($rid)) { $rolesById[$rid] = $r }
}
Write-Info ("  Roles uniques : " + $rolesById.Count)

# 11b. Pour chaque projet, recuperer ses roles en un seul appel
# GET /rest/api/3/project/{key}/role retourne { roleName: url, ... }
# puis GET sur chaque url retourne les acteurs
# On construit : roleId -> List[projectKey]
$roleProjectMap = @{} # roleId -> List[string] (projKeys ayant au moins 1 acteur)

$projCount = 0
foreach ($proj in $allProjects) {
    $projKey = [string]$proj.key
    $projCount++
    if ($projCount % 20 -eq 0) {
        Write-Info ("  Progression roles : " + $projCount + "/" + $allProjects.Count + " projets traites")
    }

    # Lister les roles du projet
    $url  = $jiraBaseUrl + "/rest/api/3/project/" + $projKey + "/role"
    $resp = $null
    try { $resp = Invoke-ApiGet -Url $url -Headers $jiraHeaders }
    catch {
        $sc = $null
        try { $sc = $_.Exception.Response.StatusCode.value__ } catch {}
        if ($sc -ne 404) {
            Write-Warn ("Roles projet " + $projKey + " : HTTP " + $sc + " - " + $_.Exception.Message)
        }
        continue
    }
    if (-not $resp) { continue }

    # $resp est un objet { "roleName": "https://.../role/10100", ... }
    foreach ($roleProp in $resp.PSObject.Properties) {
        $roleUrl = [string]$roleProp.Value

        # Extraire le roleId depuis l'URL (dernier segment)
        $roleId = ""
        try { $roleId = $roleUrl.Split("/")[-1] } catch {}
        if ([string]::IsNullOrWhiteSpace($roleId)) { continue }

        # Recuperer les acteurs de ce role sur ce projet
        $roleResp = $null
        try { $roleResp = Invoke-ApiGet -Url $roleUrl -Headers $jiraHeaders }
        catch {
            $sc = $null
            try { $sc = $_.Exception.Response.StatusCode.value__ } catch {}
            if ($sc -ne 404) {
                Write-Warn ("Acteurs role " + $roleId + " / projet " + $projKey + " : HTTP " + $sc)
            }
            continue
        }

        if ($roleResp -and $roleResp.actors -and $roleResp.actors.Count -gt 0) {
            if (-not $roleProjectMap.ContainsKey($roleId)) {
                $roleProjectMap[$roleId] = New-Object System.Collections.Generic.List[string]
            }
            $roleProjectMap[$roleId].Add($projKey) | Out-Null
        }
        Start-Sleep -Milliseconds 20
    }
    Start-Sleep -Milliseconds 50
}

Write-Info ("  Roles avec au moins 1 projet : " + $roleProjectMap.Count)

# 11c. Construire les resultats
$roleResults = New-Object System.Collections.Generic.List[object]

foreach ($roleId in $rolesById.Keys) {
    $roleObj    = $rolesById[$roleId]
    $roleName   = [string]$roleObj.name
    $projKeys   = if ($roleProjectMap.ContainsKey($roleId)) { @($roleProjectMap[$roleId]) } else { @() }
    $nbProjects = $projKeys.Count

    $kc = Get-ProjKeysAndCats -ProjIds @($projKeys | ForEach-Object {
        if ($projectByKey.ContainsKey($_)) { $projectByKey[$_].Id } else { "" }
    } | Where-Object { $_ -ne "" })

    $isSystem  = (
        $roleName -eq "atlassian-addons-project-access" -or
        $roleName -eq "Administrators"                  -or
        $roleName -eq "Administrator"                   -or
        $roleName -eq "Service Desk Team"               -or
        $roleName -eq "Service Desk Customers"          -or
        $roleName -eq "Member"                          -or
        $roleName -eq "Viewer"                          -or
        $roleName -eq "jira-guest-member"               -or
        ($roleObj.scope -and $roleObj.scope.type -eq "GLOBAL")
    )
    $candidate = if ($nbProjects -eq 0 -and -not $isSystem) { "OUI" } else { "NON" }

    $roleResults.Add([pscustomobject]@{
        "Nom"                  = $roleName
        "ID"                   = $roleId
        "Nb Projets"           = $nbProjects
        "Projets"              = Truncate-List ([string[]]$projKeys)
        "Categories"           = ($kc.Cats | Sort-Object) -join ", "
        "Candidat Suppression" = $candidate
    }) | Out-Null

    Write-Info ("  Role : " + $roleName + " (id=" + $roleId + ") -> " + $nbProjects + " projets" + $(if ($isSystem) { " [SYSTEME]" } else { "" }))
}

Write-Info ("  Roles totaux    : " + $roleResults.Count)
Write-Info ("  Roles systeme   : " + ($roleResults | Where-Object { $_."Candidat Suppression" -eq "NON" -and $_."Nb Projets" -eq 0 }).Count)
Write-Info ("  Roles candidats : " + ($roleResults | Where-Object { $_."Candidat Suppression" -eq "OUI" }).Count)

# ============================================================
# 12. EXPORT CSV
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

$csvPrio = Join-Path $exportsDir ("Audit-Priorites_" + $runStamp + ".csv")
Export-CsvStrict -Path $csvPrio `
    -Headers @("Nom", "ID", "Nb Issues", "Derniere Activite", "Candidat Suppression") `
    -Rows ($priorityResults | Sort-Object "Candidat Suppression", "Nom")

$csvWfSchemes = Join-Path $exportsDir ("Audit-WorkflowSchemes_" + $runStamp + ".csv")
Export-CsvStrict -Path $csvWfSchemes `
    -Headers @("Nom", "ID", "Nb Projets", "Projets", "Categories", "Candidat Suppression") `
    -Rows ($wfSchemeResults | Sort-Object "Candidat Suppression", "Nom")

$csvWfs = Join-Path $exportsDir ("Audit-Workflows_" + $runStamp + ".csv")
Export-CsvStrict -Path $csvWfs `
    -Headers @("Nom", "Entity ID", "Dans Scheme Actif", "Dans Scheme (tout)", "Candidat Suppression") `
    -Rows ($workflowResults | Sort-Object "Candidat Suppression", "Nom")

$csvItss = Join-Path $exportsDir ("Audit-IssueTypeScreenSchemes_" + $runStamp + ".csv")
Export-CsvStrict -Path $csvItss `
    -Headers @("Nom", "ID", "Nb Projets", "Projets", "Categories", "Candidat Suppression") `
    -Rows ($itssResults | Sort-Object "Candidat Suppression", "Nom")

$csvSS = Join-Path $exportsDir ("Audit-ScreenSchemes_" + $runStamp + ".csv")
Export-CsvStrict -Path $csvSS `
    -Headers @("Nom", "ID", "Utilise", "Candidat Suppression") `
    -Rows ($screenSchemeResults | Sort-Object "Candidat Suppression", "Nom")

$csvScreens = Join-Path $exportsDir ("Audit-Screens_" + $runStamp + ".csv")
Export-CsvStrict -Path $csvScreens `
    -Headers @("Nom", "ID", "Scope", "Utilise", "Candidat Suppression") `
    -Rows ($screenResults | Sort-Object "Candidat Suppression", "Nom")

$csvPerm = Join-Path $exportsDir ("Audit-PermissionSchemes_" + $runStamp + ".csv")
Export-CsvStrict -Path $csvPerm `
    -Headers @("Nom", "ID", "Nb Projets", "Projets", "Categories", "Candidat Suppression") `
    -Rows ($permSchemeResults | Sort-Object "Candidat Suppression", "Nom")

$csvRoles = Join-Path $exportsDir ("Audit-Roles_" + $runStamp + ".csv")
Export-CsvStrict -Path $csvRoles `
    -Headers @("Nom", "ID", "Nb Projets", "Projets", "Categories", "Candidat Suppression") `
    -Rows ($roleResults | Sort-Object "Candidat Suppression", "Nom")

# ============================================================
# 13. SAUVEGARDE CACHE
# ============================================================
Save-JsonCache -FileName "audit_priorities.json"         -Object $priorityResults
Save-JsonCache -FileName "audit_workflow_schemes.json"   -Object $wfSchemeResults
Save-JsonCache -FileName "audit_workflows.json"          -Object $workflowResults
Save-JsonCache -FileName "audit_itss.json"               -Object $itssResults
Save-JsonCache -FileName "audit_screen_schemes.json"     -Object $screenSchemeResults
Save-JsonCache -FileName "audit_screens.json"            -Object $screenResults
Save-JsonCache -FileName "audit_permission_schemes.json" -Object $permSchemeResults
Save-JsonCache -FileName "audit_roles.json"              -Object $roleResults

# ============================================================
# 14. RESUME FINAL
# ============================================================
Write-Info "=== RESUME FINAL ==="
Write-Info ("  Projets audites                     : " + $projectById.Count)
Write-Info ""
Write-Info ("  Priorites totales                   : " + $priorityResults.Count)
Write-Info ("  Priorites candidates suppression    : " + ($priorityResults     | Where-Object { $_."Candidat Suppression" -eq "OUI" }).Count)
Write-Info ""
Write-Info ("  Workflow Schemes totaux             : " + $wfSchemeResults.Count)
Write-Info ("  Workflow Schemes candidats          : " + ($wfSchemeResults     | Where-Object { $_."Candidat Suppression" -eq "OUI" }).Count)
Write-Info ("  Workflows totaux                    : " + $workflowResults.Count)
Write-Info ("  Workflows candidats                 : " + ($workflowResults     | Where-Object { $_."Candidat Suppression" -eq "OUI" }).Count)
Write-Info ""
Write-Info ("  Issue Type Screen Schemes totaux    : " + $itssResults.Count)
Write-Info ("  ITSS candidats                      : " + ($itssResults         | Where-Object { $_."Candidat Suppression" -eq "OUI" }).Count)
Write-Info ("  Screen Schemes totaux               : " + $screenSchemeResults.Count)
Write-Info ("  Screen Schemes candidats            : " + ($screenSchemeResults | Where-Object { $_."Candidat Suppression" -eq "OUI" }).Count)
Write-Info ("  Screens totaux                      : " + $screenResults.Count)
Write-Info ("  Screens candidats                   : " + ($screenResults       | Where-Object { $_."Candidat Suppression" -eq "OUI" }).Count)
Write-Info ""
Write-Info ("  Permission Schemes totaux           : " + $permSchemeResults.Count)
Write-Info ("  Permission Schemes candidats        : " + ($permSchemeResults   | Where-Object { $_."Candidat Suppression" -eq "OUI" }).Count)
Write-Info ""
Write-Info ("  Roles totaux                        : " + $roleResults.Count)
Write-Info ("  Roles candidats                     : " + ($roleResults         | Where-Object { $_."Candidat Suppression" -eq "OUI" }).Count)
Write-Info ""
Write-Info ("  Exports : " + $exportsDir)
Write-Info ("  Cache   : " + $cacheDir)
Write-Info ("  Log     : " + $logFile)
Write-Log ("=== FIN " + $scriptName + " ===") "INFO"
