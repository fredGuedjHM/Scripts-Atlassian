param(
    [string]$SiteUrl = "https://jiradot.atlassian.net",
    [string]$GroupQuery = "DSIM",
    [string]$OutputCsv = "C:\Temp\DSIM_Groups_x_ProjectCategories.csv",
    [switch]$ResetCreds
)

$ErrorActionPreference = "Stop"
$ProgressPreference = "SilentlyContinue"
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

# Proxy système
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

function Invoke-JiraRest($method, $url, $headers) {
    $proxyUri = Get-ProxyForUrl $url
    $params = @{
        Method      = $method
        Uri         = $url
        Headers     = $headers
        ErrorAction = "Stop"
    }
    if ($proxyUri) {
        $params.Proxy = $proxyUri
        $params.ProxyUseDefaultCredentials = $true
    }
    return Invoke-RestMethod @params
}

function Invoke-JiraRestWithRetry($method, $url, $headers, [int]$maxRetries = 5) {
    $attempt = 0
    while ($true) {
        try {
            return Invoke-JiraRest $method $url $headers
        } catch {
            $attempt++
            if ($_.Exception.Message -match '\b429\b' -and $attempt -le $maxRetries) {
                $sleep = 10 + (5 * $attempt)
                Write-Warning "HTTP 429. Retry dans $sleep sec (tentative $attempt/$maxRetries): $url"
                Start-Sleep -Seconds $sleep
                continue
            }
            throw
        }
    }
}

# ---- DPAPI creds (email + token) ----
$credDir  = Join-Path $env:APPDATA "Jira"
$credPath = Join-Path $credDir "jira-cloud-cred.clixml"

function Get-JiraCredential($credPath, $reset) {
    $dir = Split-Path $credPath
    if (-not (Test-Path $dir)) { New-Item -ItemType Directory -Path $dir | Out-Null }

    if (-not $reset -and (Test-Path $credPath)) {
        try { return Import-Clixml -Path $credPath } catch { }
    }

    $email = Read-Host "Email Atlassian"
    $secureToken = Read-Host "API Token Atlassian (saisie masquée)" -AsSecureString
    $cred = New-Object System.Management.Automation.PSCredential($email, $secureToken)
    $cred | Export-Clixml -Path $credPath
    return $cred
}

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
}

# Test API
Write-Host "Test API /myself ..." -ForegroundColor Cyan
$me = Invoke-JiraRestWithRetry "GET" "$SiteUrl/rest/api/3/myself" $headers
Write-Host ("API OK - " + $me.displayName) -ForegroundColor Green

# ---- 1) Catégories de projet (avec description) ----
Write-Host "Récupération des catégories de projets..." -ForegroundColor Cyan
$categories = Invoke-JiraRestWithRetry "GET" "$SiteUrl/rest/api/3/projectCategory" $headers
if (-not $categories) { $categories = @() }
Write-Host ("Catégories: " + $categories.Count) -ForegroundColor Green

# ---- 2) Groupes filtrés par 'DSIM' ----
Write-Host "Récupération des groupes (query='$GroupQuery')..." -ForegroundColor Cyan
$groups = @()
$startAt = 0
$maxResults = 200
$q = [System.Uri]::EscapeDataString($GroupQuery)

while ($true) {
    $url = "$SiteUrl/rest/api/3/groups/picker?query=$q&startAt=$startAt&maxResults=$maxResults&caseInsensitive=true"
    $resp = Invoke-JiraRestWithRetry "GET" $url $headers

    $pageGroups = @()
    if ($resp.groups) { $pageGroups = $resp.groups }

    $groups += $pageGroups

    if ($pageGroups.Count -lt $maxResults) { break }
    $startAt += $pageGroups.Count
}

# Filtre strict "contient DSIM" sur le nom
$groups = $groups |
    Where-Object { $_.name -and $_.name.ToUpper().Contains($GroupQuery.ToUpper()) } |
    Sort-Object name -Unique

Write-Host ("Groupes trouvés (nom contenant '$GroupQuery'): " + $groups.Count) -ForegroundColor Green

# ---- 2b) Enrichir les groupes avec leur description ----
Write-Host "Enrichissement des groupes avec leur description..." -ForegroundColor Cyan
$groupDetailsById = @{}

foreach ($g in $groups) {
    $gid = $g.groupId
    if (-not $gid) { continue }

    $encGid = [System.Uri]::EscapeDataString($gid)
    $url = "$SiteUrl/rest/api/3/group?groupId=$encGid"

    try {
        $detail = Invoke-JiraRestWithRetry "GET" $url $headers
        # Certains schémas ont un champ "description", d'autres pas -> on gère le cas null
        $groupDetailsById[$gid] = @{
            Name        = $detail.name
            Description = $detail.description
        }
    } catch {
        Write-Warning "Impossible de récupérer les détails du groupe $($g.name) (groupId=$gid) : $_"
        $groupDetailsById[$gid] = @{
            Name        = $g.name
            Description = $null
        }
    }
}

# ---- 3) Produit cartésien avec descriptions ----
Write-Host "Génération produit cartésien (Groupes x Catégories)..." -ForegroundColor Cyan
$out = @()

foreach ($g in $groups) {
    $gid = $g.groupId
    $gInfo = $null
    if ($gid -and $groupDetailsById.ContainsKey($gid)) {
        $gInfo = $groupDetailsById[$gid]
    }

    $gName = if ($gInfo) { $gInfo.Name } else { $g.name }
    $gDesc = if ($gInfo) { $gInfo.Description } else { $null }

    foreach ($c in $categories) {
        $out += [PSCustomObject]@{
            "DSIM Group Name"           = $gName
            "DSIM Group Id"             = $gid
            "DSIM Group Description"    = $gDesc
            "Project Category"          = $c.name
            "Project Category Id"       = $c.id
            "Project Category Description" = $c.description
        }
    }
}

Write-Host ("Lignes générées: " + $out.Count) -ForegroundColor Green

# ---- Export ----
$dir = Split-Path -Parent $OutputCsv
if (-not (Test-Path $dir)) { New-Item -ItemType Directory -Path $dir | Out-Null }

$out | Export-Csv -Path $OutputCsv -Delimiter ';' -NoTypeInformation -Encoding UTF8
Write-Host "Export OK: $OutputCsv" -ForegroundColor Green
Write-Host "Creds: $credPath"