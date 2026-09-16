<#
Met en regard, par projet Jira:
- Catégorie du projet
- Groupes DSIM présents (uniquement) dans les rôles de projet
- Rôle(s) où chaque groupe DSIM est présent
- Heuristique de correspondance groupe <-> catégorie
- Recommandation: "Gestionnaire projet (candidat)" si match, sinon "Utilisateur (candidat)"

Exports:
  * C:\Temp\dsim_groupes_vs_categorie_detail.csv
  * C:\Temp\dsim_groupes_vs_categorie_resume_projet.csv
  * C:\Temp\dsim_groupes_vs_categorie_resume_groupe.csv

Usage:
  .\dsimGroupesVsCategorieParProjet.ps1
  .\dsimGroupesVsCategorieParProjet.ps1 -Filter "DSIM"
  .\dsimGroupesVsCategorieParProjet.ps1 -MinTokenLength 3 -ResetCreds
#>

param(
    [string]$Filter = "DSIM",
    [int]$MinTokenLength = 3,
    [switch]$ResetCreds
)

# ----------------------------
# Runtime
# ----------------------------
$ErrorActionPreference = "Stop"
$ProgressPreference = "SilentlyContinue"
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

# Proxy système Windows + creds Windows (PAC/WPAD/McAfee)
$systemProxy = [System.Net.WebRequest]::GetSystemWebProxy()
$systemProxy.Credentials = [System.Net.CredentialCache]::DefaultNetworkCredentials

function Get-ProxyForUrl($url) {
    try {
        $u = [Uri]$url
        $p = $systemProxy.GetProxy($u)
        if ($p -and $p.AbsoluteUri -ne $u.AbsoluteUri) { return $p.AbsoluteUri }
        return $null
    } catch { return $null }
}

function Invoke-JiraRest($method, $url, $headers, $body = $null) {
    $proxyUri = Get-ProxyForUrl $url

    $params = @{
        Method      = $method
        Uri         = $url
        Headers     = $headers
        ErrorAction = "Stop"
    }
    if ($null -ne $body) { $params.Body = $body }

    if ($proxyUri) {
        $params.Proxy = $proxyUri
        $params.ProxyUseDefaultCredentials = $true
    }

    try {
        return Invoke-RestMethod @params
    } catch {
        $status = $null
        try {
            if ($_.Exception.Response -and $_.Exception.Response.StatusCode) {
                $status = [int]$_.Exception.Response.StatusCode
            }
        } catch { }

        $msg = $_.Exception.Message
        $details = $null
        if ($_.ErrorDetails -and $_.ErrorDetails.Message) { $details = $_.ErrorDetails.Message }
        if ($_.Exception.InnerException) { $msg = "$msg | Inner: $($_.Exception.InnerException.Message)" }

        if ($details) {
            if ($status) { throw "HTTP $status failed: $method $url => $msg | Details: $details" }
            throw "HTTP failed: $method $url => $msg | Details: $details"
        }
        if ($status) { throw "HTTP $status failed: $method $url => $msg" }
        throw "HTTP failed: $method $url => $msg"
    }
}

function Invoke-JiraRestWithRetry($method, $url, $headers, $body = $null, [int]$maxRetries = 5) {
    $attempt = 0
    while ($true) {
        try {
            return Invoke-JiraRest $method $url $headers $body
        } catch {
            $attempt++
            if ($_.Exception.Message -match '\bHTTP 429\b' -and $attempt -le $maxRetries) {
                $sleep = 30 + (5 * $attempt)
                Write-Warning "Rate limit (429). Retry dans $sleep sec (tentative $attempt/$maxRetries) : $url"
                Start-Sleep -Seconds $sleep
                continue
            }
            throw
        }
    }
}

function Normalize-Text([string]$s) {
    if ([string]::IsNullOrWhiteSpace($s)) { return "" }
    $s = $s.Trim()
    $s = $s -replace '\s+', ' '

    # suppression accents
    $formD = $s.Normalize([Text.NormalizationForm]::FormD)
    $sb = New-Object System.Text.StringBuilder
    foreach ($ch in $formD.ToCharArray()) {
        $cat = [Globalization.CharUnicodeInfo]::GetUnicodeCategory($ch)
        if ($cat -ne [Globalization.UnicodeCategory]::NonSpacingMark) {
            [void]$sb.Append($ch)
        }
    }
    $clean = $sb.ToString().Normalize([Text.NormalizationForm]::FormC)

    $clean = $clean.ToLowerInvariant()
    # remplacer tout séparateur par espace
    $clean = $clean -replace '[^a-z0-9]+', ' '
    $clean = $clean -replace '\s+', ' '
    return $clean.Trim()
}

function Get-Tokens([string]$s, [int]$minLen) {
    $n = Normalize-Text $s
    if ([string]::IsNullOrWhiteSpace($n)) { return @() }
    return ($n -split ' ') | Where-Object { $_.Length -ge $minLen } | Select-Object -Unique
}

function Test-CategoryMatch([string]$categoryName, [string]$groupName, [int]$minTokenLength) {
    # Heuristique: si au moins 1 token "significatif" de la catégorie est contenu dans le nom du groupe
    $catTokens = Get-Tokens $categoryName $minTokenLength
    $gNorm = Normalize-Text $groupName

    if ($catTokens.Count -eq 0 -or [string]::IsNullOrWhiteSpace($gNorm)) {
        return [PSCustomObject]@{
            Match = $false
            MatchTokens = ""
        }
    }

    $hits = @()
    foreach ($t in $catTokens) {
        if ($gNorm -like "*$t*") { $hits += $t }
    }

    return [PSCustomObject]@{
        Match = ($hits.Count -gt 0)
        MatchTokens = ($hits | Sort-Object -Unique) -join ";"
    }
}

# ----------------------------
# Config
# ----------------------------
$siteUrl = "https://jiradot.atlassian.net"

$exportDetailPath   = "C:\Temp\dsim_groupes_vs_categorie_detail.csv"
$exportProjectPath  = "C:\Temp\dsim_groupes_vs_categorie_resume_projet.csv"
$exportGroupPath    = "C:\Temp\dsim_groupes_vs_categorie_resume_groupe.csv"

# DPAPI creds
$credDir  = Join-Path $env:APPDATA "Jira"
$credPath = Join-Path $credDir "jira-cloud-cred.clixml"

function Get-JiraCredential($credPath, $reset) {
    $dir = Split-Path $credPath
    if (-not (Test-Path $dir)) { New-Item -ItemType Directory -Path $dir | Out-Null }

    if (-not $reset -and (Test-Path $credPath)) {
        try { return Import-Clixml -Path $credPath }
        catch { Write-Warning "Impossible de relire le credential, nouvelle saisie requise. Détail: $_" }
    }

    $email = Read-Host "Email Atlassian (ex: prenom.nom@domaine.fr)"
    if ($email -match '^\[(.+?)\]\(mailto:(.+?)\)$') { $email = $Matches[2] }
    $email = ($email -replace '^mailto:', '').Trim()

    $secureToken = Read-Host "API Token Atlassian (saisie masquée)" -AsSecureString
    $cred = New-Object System.Management.Automation.PSCredential($email, $secureToken)
    $cred | Export-Clixml -Path $credPath
    Write-Host "Identifiants sauvegardés dans: $credPath" -ForegroundColor Green
    return $cred
}

# ----------------------------
# Auth header
# ----------------------------
$jiraCred = Get-JiraCredential -credPath $credPath -reset:$ResetCreds

$bstr = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($jiraCred.Password)
try { $apiTokenPlain = [Runtime.InteropServices.Marshal]::PtrToStringBSTR($bstr) }
finally { [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($bstr) }

$pair   = "$($jiraCred.UserName)`:$apiTokenPlain"
$bytes  = [Text.Encoding]::UTF8.GetBytes($pair)
$base64 = [Convert]::ToBase64String($bytes)

$headers = @{
    Authorization = "Basic $base64"
    Accept        = "application/json"
    "Content-Type"= "application/json"
}

# ----------------------------
# Test API
# ----------------------------
Write-Host "Test API /myself ..." -ForegroundColor Cyan
$me = Invoke-JiraRestWithRetry "GET" "$siteUrl/rest/api/3/myself" $headers
Write-Host ("API OK - " + $me.displayName) -ForegroundColor Green

# ----------------------------
# 1) Liste des groupes DSIM (référence)
# ----------------------------
Write-Host "Recherche des groupes contenant: '$Filter' ..." -ForegroundColor Cyan
$encoded = [System.Uri]::EscapeDataString($Filter)
$groupsUrl = "$siteUrl/rest/api/3/groups/picker?query=$encoded&maxResults=1000&caseInsensitive=true"
$groupsResp = Invoke-JiraRestWithRetry "GET" $groupsUrl $headers

$dsimGroups = @()
if ($groupsResp.groups) { $dsimGroups = $groupsResp.groups }

$dsimGroups = $dsimGroups |
    Where-Object { $_.name -and ($_.name.ToUpper().Contains($Filter.ToUpper())) } |
    Sort-Object name

Write-Host ("Groupes DSIM trouvés: " + $dsimGroups.Count) -ForegroundColor Green
if ($dsimGroups.Count -eq 0) { throw "Aucun groupe ne correspond au filtre '$Filter'." }

# Index rapide par nom (pour filtrer strictement sur la liste DSIM)
$dsimGroupNameSet = New-Object System.Collections.Generic.HashSet[string]
foreach ($g in $dsimGroups) { [void]$dsimGroupNameSet.Add($g.name) }

# ----------------------------
# 2) Projets live (+ catégories)
# ----------------------------
Write-Host "Récupération des projets live..." -ForegroundColor Cyan
$allProjects = @()
$startAt = 0
$maxResults = 50

while ($true) {
    $url = "$siteUrl/rest/api/3/project/search?startAt=$startAt&maxResults=$maxResults&status=live"
    $resp = Invoke-JiraRestWithRetry "GET" $url $headers
    if ($resp.values) { $allProjects += $resp.values }
    $startAt += ($resp.values | Measure-Object).Count
    if ($startAt -ge $resp.total) { break }
}

Write-Host ("Nombre de projets live: " + $allProjects.Count) -ForegroundColor Green

# ----------------------------
# 3) Détail projet x rôle x groupe DSIM
# ----------------------------
Write-Host "Analyse des rôles et extraction des groupes DSIM par projet..." -ForegroundColor Cyan

$resultDetail = @()

foreach ($p in $allProjects) {
    $projectKey  = $p.key
    $projectName = $p.name
    if ([string]::IsNullOrWhiteSpace($projectKey)) { continue }

    $categoryName = $null
    if ($p.projectCategory -and $p.projectCategory.name) {
        $categoryName = $p.projectCategory.name
    } else {
        $categoryName = "(Aucune catégorie)"
    }

    Write-Host "Projet $projectKey - $projectName (Catégorie: $categoryName)" -ForegroundColor Yellow

    $rolesResp = $null
    try {
        $rolesResp = Invoke-JiraRestWithRetry "GET" "$siteUrl/rest/api/3/project/$projectKey/role" $headers
    } catch {
        Write-Warning "Impossible de lire les rôles pour $projectKey : $_"
        continue
    }

    foreach ($prop in $rolesResp.PSObject.Properties) {
        $roleUrl = $prop.Value

        $roleDetail = $null
        try {
            $roleDetail = Invoke-JiraRestWithRetry "GET" $roleUrl $headers
        } catch {
            Write-Warning "Détail rôle non récupéré ($($prop.Name)) pour $projectKey : $_"
            continue
        }

        $roleId   = $roleDetail.id
        $roleName = $roleDetail.name

        foreach ($actor in $roleDetail.actors) {
            # on ne garde que les groupes
            if ($actor.type -ne "atlassian-group-role-actor") { continue }

            $groupName = $null
            if ($actor.actorGroup -and $actor.actorGroup.name) { $groupName = $actor.actorGroup.name }
            elseif ($actor.displayName) { $groupName = $actor.displayName }

            if ([string]::IsNullOrWhiteSpace($groupName)) { continue }

            # on ne garde QUE les groupes DSIM
            # (strict: doit être dans la liste des groupes DSIM trouvés)
            if (-not $dsimGroupNameSet.Contains($groupName)) { continue }

            $matchInfo = Test-CategoryMatch -categoryName $categoryName -groupName $groupName -minTokenLength $MinTokenLength
            $suggested = $(if ($matchInfo.Match) { "Gestionnaire projet (candidat)" } else { "Utilisateur (candidat)" })

            $resultDetail += [PSCustomObject]@{
                "Project Key"        = $projectKey
                "Project Name"       = $projectName
                "Project Category"   = $categoryName
                "Role Id"            = $roleId
                "Role Name"          = $roleName
                "DSIM Group Name"    = $groupName
                "Category Match"     = $matchInfo.Match
                "Matched Tokens"     = $matchInfo.MatchTokens
                "Suggested Profile"  = $suggested
            }
        }
    }
}

Write-Host ("Lignes détail: " + $resultDetail.Count) -ForegroundColor Green

# ----------------------------
# 4) Résumé par projet
# ----------------------------
$projectSummary = @()
if ($resultDetail.Count -gt 0) {
    $projectSummary = $resultDetail |
        Group-Object "Project Key" |
        ForEach-Object {
            $pKey = $_.Name
            $one  = $_.Group | Select-Object -First 1

            $matchingGroups = $_.Group |
                Where-Object { $_."Category Match" -eq $true } |
                Select-Object -ExpandProperty "DSIM Group Name" -Unique |
                Sort-Object

            $otherGroups = $_.Group |
                Where-Object { $_."Category Match" -eq $false } |
                Select-Object -ExpandProperty "DSIM Group Name" -Unique |
                Sort-Object

            [PSCustomObject]@{
                "Project Key"                = $pKey
                "Project Name"               = $one."Project Name"
                "Project Category"           = $one."Project Category"
                "DSIM Groups (Matching)"     = ($matchingGroups -join ";")
                "DSIM Groups (Other)"        = ($otherGroups -join ";")
                "Matching Groups Count"      = $matchingGroups.Count
                "Other Groups Count"         = $otherGroups.Count
            }
        } | Sort-Object "Project Key"
}

# ----------------------------
# 5) Résumé par groupe DSIM
# ----------------------------
$groupSummary = @()
if ($resultDetail.Count -gt 0) {
    $groupSummary = $resultDetail |
        Group-Object "DSIM Group Name" |
        ForEach-Object {
            $gName = $_.Name

            $projects = $_.Group |
                Select-Object "Project Key","Project Name" -Unique |
                Sort-Object "Project Key"

            $roles = $_.Group |
                Select-Object "Role Name" -Unique |
                Sort-Object "Role Name"

            [PSCustomObject]@{
                "DSIM Group Name" = $gName
                "Projects Count"  = $projects.Count
                "Projects (Keys)" = ($projects | ForEach-Object { $_."Project Key" }) -join ";"
                "Roles Used"      = ($roles | ForEach-Object { $_."Role Name" }) -join ";"
            }
        } | Sort-Object "DSIM Group Name"
}

# ----------------------------
# Export CSV
# ----------------------------
$exportDir = Split-Path $exportDetailPath
if (-not (Test-Path $exportDir)) { New-Item -ItemType Directory -Path $exportDir | Out-Null }

Write-Host "Export détail         : $exportDetailPath" -ForegroundColor Cyan
$resultDetail | Export-Csv -Path $exportDetailPath -NoTypeInformation -Encoding UTF8

Write-Host "Export résumé projet  : $exportProjectPath" -ForegroundColor Cyan
$projectSummary | Export-Csv -Path $exportProjectPath -NoTypeInformation -Encoding UTF8

Write-Host "Export résumé groupe  : $exportGroupPath" -ForegroundColor Cyan
$groupSummary | Export-Csv -Path $exportGroupPath -NoTypeInformation -Encoding UTF8

Write-Host "Terminé." -ForegroundColor Green
Write-Host " - Détail        : $exportDetailPath"
Write-Host " - Résumé projet : $exportProjectPath"
Write-Host " - Résumé groupe : $exportGroupPath"
Write-Host " - Creds         : $credPath"