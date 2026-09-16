<#
.SYNOPSIS
    Script de diagnostic autonome des habilitations d'un espace Confluence Cloud.
.DESCRIPTION
    1. Interroge un espace Confluence via l'API v1 (expand multi-niveaux) et l'API v2 (avec pagination complète).
    2. Affiche la liste exhaustive des groupes et utilisateurs ayant des droits sur l'espace.
    3. Résout et compare les identifiants textuels (Name) et UUIDs (GroupId / principalId).
.VERSION
    1.1 — 2026-09-14
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [string]$SpaceKey,

    [Parameter(Mandatory = $false)]
    [string]$GroupToTest,

    [Parameter(Mandatory = $false)]
    [string]$CredentialsPath
)

[System.Threading.Thread]::CurrentThread.CurrentCulture   = [System.Globalization.CultureInfo]::GetCultureInfo("fr-FR")
[System.Threading.Thread]::CurrentThread.CurrentUICulture = [System.Globalization.CultureInfo]::GetCultureInfo("fr-FR")

Set-StrictMode -Off
$ErrorActionPreference = "Continue"

# ============================================================
# 0. CONSTANTES & CHEMINS
# ============================================================
$scriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
if (-not $scriptRoot) { $scriptRoot = Get-Location }

$secretsDir = Join-Path $scriptRoot "secrets"

$ProxyUrl                   = ""
$UseSystemProxy             = $true
$ProxyUseDefaultCredentials = $true

# ============================================================
# 1. HELPERS LOGGING & PROXY REST API
# ============================================================
function Write-Info([string]$Message) { Write-Host ("[INFO] " + $Message) -ForegroundColor Cyan }
function Write-Warn([string]$Message) { Write-Warning $Message }
function Write-ErrLog([string]$Message) { Write-Error $Message }

function Fix-Encoding([string]$val) {
    if ([string]::IsNullOrWhiteSpace($val)) { return $val }
    if ($val.Contains("Ã")) {
        try {
            $bytes   = [System.Text.Encoding]::GetEncoding(1252).GetBytes($val)
            $decoded = [System.Text.Encoding]::UTF8.GetString($bytes)
            if (-not $decoded.Contains("")) { return $decoded }
        } catch {}
    }
    return $val
}

function Get-ProxyParams([string]$TargetUrl) {
    $params = @{}
    if ($ProxyUrl -and $ProxyUrl.Trim()) {
        $params.Proxy = $ProxyUrl
        if ($ProxyUseDefaultCredentials) { $params.ProxyUseDefaultCredentials = $true }
        return $params
    }
    if (-not $UseSystemProxy) { return $params }
    try {
        $dest = [uri]$TargetUrl
        $wp   = [System.Net.WebRequest]::DefaultWebProxy
        if ($wp -and -not $wp.IsBypassed($dest)) {
            $proxy = $wp.GetProxy($dest)
            if ($proxy -and $proxy.AbsoluteUri -ne $dest.AbsoluteUri) {
                $params.Proxy = $proxy.AbsoluteUri
                if ($ProxyUseDefaultCredentials) { $params.ProxyUseDefaultCredentials = $true }
            }
        }
    } catch {}
    return $params
}

function Invoke-AtlassianApiGet([string]$Url, [hashtable]$Headers) {
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    $params = @{
        Method          = 'GET'
        Uri             = $Url
        Headers         = $Headers
        ContentType     = 'application/json'
        UseBasicParsing = $true
        ErrorAction     = 'Stop'
    }
    $px = Get-ProxyParams $Url
    foreach ($k in $px.Keys) { $params[$k] = $px[$k] }
    
    $resp   = Invoke-WebRequest @params
    $stream = $resp.RawContentStream
    $stream.Position = 0
    $reader = New-Object System.IO.StreamReader($stream, [System.Text.Encoding]::UTF8, $true)
    $raw    = $reader.ReadToEnd()
    $reader.Close()
    return ($raw | ConvertFrom-Json)
}

# ============================================================
# 2. CHARGEMENT DES CREDENTIALS
# ============================================================
if (-not $CredentialsPath) {
    $CredentialsPath = Join-Path $secretsDir "jira-jiradot.cred.xml"
}
if (-not (Test-Path $CredentialsPath)) {
    throw "Fichier credentials introuvable : $CredentialsPath"
}

$jiraData    = Import-Clixml -Path $CredentialsPath
$jiraBaseUrl = ([string]$jiraData.JiraBaseUrl).TrimEnd("/")
$jiraCred    = $jiraData.Credential
$b64         = [Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes($jiraCred.UserName + ":" + $jiraCred.GetNetworkCredential().Password))
$authHeaders = @{ Authorization = "Basic " + $b64; Accept = "application/json" }

# ============================================================
# 3. SAISIE DE L'ESPACE ET DU GROUPE À TESTER
# ============================================================
Write-Host ""
Write-Host ("=" * 80) -ForegroundColor DarkCyan
Write-Host "  DIAGNOSTIC DES HABILITATIONS D'UN ESPACE CONFLUENCE CLOUD" -ForegroundColor White
Write-Host ("=" * 80) -ForegroundColor DarkCyan
Write-Host ""

if ([string]::IsNullOrWhiteSpace($SpaceKey)) {
    $SpaceKey = Read-Host "Entrez la clé de l'espace Confluence à analyser (ex: OAV, DSIM, ODP, ...)"
}
$SpaceKey = $SpaceKey.Trim().ToUpper()

if ([string]::IsNullOrWhiteSpace($SpaceKey)) {
    throw "La clé de l'espace est obligatoire."
}

if ([string]::IsNullOrWhiteSpace($GroupToTest)) {
    $GroupToTest = Read-Host "Entrez un groupe à tester en particulier (laisser vide pour tout lister)"
}
$GroupToTest = $GroupToTest.Trim()

# Résolution préalable de l'UUID du groupe si renseigné
$groupToTestId = ""
if (-not [string]::IsNullOrWhiteSpace($GroupToTest)) {
    try {
        $gUrl  = $jiraBaseUrl + "/rest/api/3/group/bulk?groupName=" + [Uri]::EscapeDataString($GroupToTest)
        $gResp = Invoke-AtlassianApiGet -Url $gUrl -Headers $authHeaders
        if ($gResp -and $gResp.values -and $gResp.values.Count -gt 0) {
            $groupToTestId = [string]$gResp.values[0].groupId
        }
    } catch {}

    $dispGId = if ($groupToTestId) { $groupToTestId } else { "Non trouvé / Inconnu" }
    Write-Info ("Groupe ciblé pour test : [" + $GroupToTest + "] (UUID: " + $dispGId + ")")
}

Write-Host ""
Write-Info ("Interrogation de l'espace : [" + $SpaceKey + "]...")

$sId = ""

# ============================================================
# 4. TEST 1 : API CONFLUENCE V1 (/wiki/rest/api/space/{key})
# ============================================================
Write-Host ""
Write-Host ("-" * 80) -ForegroundColor DarkGray
Write-Host "1. ANALYSE VIA L'API CONFLUENCE V1 (/wiki/rest/api/space/{key})" -ForegroundColor Cyan
Write-Host ("-" * 80) -ForegroundColor DarkGray

$v1Url = $jiraBaseUrl + "/wiki/rest/api/space/" + [Uri]::EscapeDataString($SpaceKey) + "?expand=permissions.subjects.group,permissions.subjects.user,description.plain"

try {
    $spaceV1 = Invoke-AtlassianApiGet -Url $v1Url -Headers $authHeaders
    $sName   = Fix-Encoding ([string]$spaceV1.name)
    $sId     = [string]$spaceV1.id
    
    Write-Host ("  Nom de l'espace : ") -NoNewline -ForegroundColor White
    Write-Host ($sName) -ForegroundColor Yellow
    Write-Host ("  ID interne      : ") -NoNewline -ForegroundColor White
    Write-Host ($sId) -ForegroundColor Gray

    $permsList = $spaceV1.permissions
    $nbPermsV1 = if ($permsList) { $permsList.Count } else { 0 }
    Write-Host ("  Nombre total d'éléments de permissions retournés : " + $nbPermsV1) -ForegroundColor White
    Write-Host ""

    if ($permsList -and $permsList.Count -gt 0) {
        $groupPermMap = @{}

        foreach ($p in $permsList) {
            $opKey    = if ($p.operation -and $p.operation.key) { [string]$p.operation.key } else { [string]$p.type }
            $opTarget = if ($p.operation -and $p.operation.target) { [string]$p.operation.target } else { "" }
            $opLabel  = if (-not [string]::IsNullOrWhiteSpace($opTarget)) { "$opKey ($opTarget)" } else { $opKey }

            $foundGroups = @()
            if ($p.subjects -and $p.subjects.group -and $p.subjects.group.results) {
                $foundGroups = $p.subjects.group.results
            } elseif ($p.subjects -and $p.subjects.groups -and $p.subjects.groups.results) {
                $foundGroups = $p.subjects.groups.results
            }

            foreach ($g in $foundGroups) {
                $gName = [string]$g.name
                $gId   = [string]$g.id
                $gKey  = if (-not [string]::IsNullOrWhiteSpace($gName)) { "$gName|$gId" } else { "ID:$gId|$gId" }

                if (-not $groupPermMap.ContainsKey($gKey)) {
                    $groupPermMap[$gKey] = New-Object System.Collections.ArrayList
                }
                [void]$groupPermMap[$gKey].Add($opLabel)
            }
        }

        Write-Host "  Groupes habilités identifiés (API v1) :" -ForegroundColor Yellow
        if ($groupPermMap.Keys.Count -gt 0) {
            foreach ($k in ($groupPermMap.Keys | Sort-Object)) {
                $parts = $k -split '\|'
                $name  = $parts[0]
                $id    = $parts[1]
                $perms = ($groupPermMap[$k] | Select-Object -Unique) -join ", "

                $isTarget = $false
                if ($GroupToTest) {
                    if ($name -ieq $GroupToTest -or ($groupToTestId -and $id -ieq $groupToTestId)) { $isTarget = $true }
                }

                $color = if ($isTarget) { "Green" } else { "White" }
                $tag   = if ($isTarget) { " [CIBLE TROUVÉE]" } else { "" }

                Write-Host ("    * Groupe : ") -NoNewline -ForegroundColor Gray
                Write-Host ($name.PadRight(35)) -NoNewline -ForegroundColor $color
                Write-Host (" (ID: " + $id + ")" + $tag) -ForegroundColor DarkGray
                Write-Host ("      Permissions : " + $perms) -ForegroundColor DarkCyan
            }
        } else {
            Write-Host "    -> Aucun groupe résolu dans la structure v1." -ForegroundColor Yellow
        }
    }

} catch {
    Write-Warn ("Erreur API v1 sur l'espace '$SpaceKey' : " + $_.Exception.Message)
}

# ============================================================
# 5. TEST 2 : API CONFLUENCE V2 (/wiki/api/v2/spaces/{id}/permissions)
# ============================================================
Write-Host ""
Write-Host ("-" * 80) -ForegroundColor DarkGray
Write-Host "2. ANALYSE VIA L'API CONFLUENCE V2 (/wiki/api/v2/spaces/{id}/permissions)" -ForegroundColor Cyan
Write-Host ("-" * 80) -ForegroundColor DarkGray

# Résolution de l'ID d'espace si non obtenu en v1
if ([string]::IsNullOrWhiteSpace($sId)) {
    try {
        $spLookupUrl = $jiraBaseUrl + "/wiki/api/v2/spaces?keys=" + [Uri]::EscapeDataString($SpaceKey)
        $spLookupResp = Invoke-AtlassianApiGet -Url $spLookupUrl -Headers $authHeaders
        if ($spLookupResp -and $spLookupResp.results -and $spLookupResp.results.Count -gt 0) {
            $sId = [string]$spLookupResp.results[0].id
        }
    } catch {}
}

if (-not [string]::IsNullOrWhiteSpace($sId)) {
    try {
        $v2AllPermissions = New-Object System.Collections.ArrayList
        $nextCursor = ""
        $hasMoreV2  = $true

        # Boucle de pagination v2 complète
        do {
            $v2Url = $jiraBaseUrl + "/wiki/api/v2/spaces/" + $sId + "/permissions?limit=250"
            if (-not [string]::IsNullOrWhiteSpace($nextCursor)) {
                $v2Url += "&cursor=" + [Uri]::EscapeDataString($nextCursor)
            }

            $spaceV2 = Invoke-AtlassianApiGet -Url $v2Url -Headers $authHeaders
            if ($spaceV2 -and $spaceV2.results) {
                foreach ($item in $spaceV2.results) {
                    [void]$v2AllPermissions.Add($item)
                }
            }

            if ($spaceV2._links -and $spaceV2._links.next) {
                $match = [regex]::Match($spaceV2._links.next, "cursor=([^&]+)")
                if ($match.Success) {
                    $nextCursor = $match.Groups[1].Value
                } else {
                    $hasMoreV2 = $false
                }
            } else {
                $hasMoreV2 = $false
            }
        } while ($hasMoreV2)

        Write-Host ("  Nombre total de permissions v2 récupérées : " + $v2AllPermissions.Count) -ForegroundColor White
        Write-Host ""

        $v2GroupMap = @{}
        foreach ($permItem in $v2AllPermissions) {
            if ($permItem.principal -and [string]$permItem.principal.type -ieq "group") {
                $groupIdVal = [string]$permItem.principal.id
                $opKey      = [string]$permItem.operation.key
                $opTarget   = [string]$permItem.operation.target
                $opFull     = if (-not [string]::IsNullOrWhiteSpace($opTarget)) { "$opKey ($opTarget)" } else { $opKey }

                if (-not $v2GroupMap.ContainsKey($groupIdVal)) {
                    $v2GroupMap[$groupIdVal] = New-Object System.Collections.ArrayList
                }
                [void]$v2GroupMap[$groupIdVal].Add($opFull)
            }
        }

        Write-Host "  Groupes (UUID / GroupId) identifiés en API v2 :" -ForegroundColor Yellow
        if ($v2GroupMap.Keys.Count -gt 0) {
            foreach ($groupIdKey in ($v2GroupMap.Keys | Sort-Object)) {
                $permsStr = ($v2GroupMap[$groupIdKey] | Select-Object -Unique) -join ", "
                $isTarget = ($groupToTestId -and $groupIdKey -ieq $groupToTestId) -or ($GroupToTest -and $groupIdKey -ieq $GroupToTest)
                $color    = if ($isTarget) { "Green" } else { "White" }

                # Tentative de résolution du nom lisible via Jira Bulk
                $resolvedName = ""
                try {
                    $lookupNameUrl = $jiraBaseUrl + "/rest/api/3/group/bulk?groupId=" + [Uri]::EscapeDataString($groupIdKey)
                    $nameResp      = Invoke-AtlassianApiGet -Url $lookupNameUrl -Headers $authHeaders
                    if ($nameResp -and $nameResp.values -and $nameResp.values.Count -gt 0) {
                        $resolvedName = [string]$nameResp.values[0].name
                    }
                } catch {}

                $displayHeader = if ($resolvedName) { "$resolvedName ($groupIdKey)" } else { "UUID: $groupIdKey" }

                Write-Host ("    * ") -NoNewline -ForegroundColor Gray
                Write-Host ($displayHeader.PadRight(45)) -NoNewline -ForegroundColor $color
                if ($isTarget) { Write-Host " [CIBLE TROUVÉE]" -ForegroundColor Green } else { Write-Host "" }
                Write-Host ("      Permissions : " + $permsStr) -ForegroundColor DarkCyan
            }
        } else {
            Write-Host "    -> Aucun groupe retourné par l'endpoint v2." -ForegroundColor Yellow
        }

    } catch {
        Write-Warn ("API v2 non disponible ou erreur : " + $_.Exception.Message)
    }
} else {
    Write-Host "  -> ID de l'espace non résolu, test v2 ignoré." -ForegroundColor Yellow
}

Write-Host ""
Write-Host ("=" * 80) -ForegroundColor DarkCyan
Write-Host "  FIN DU DIAGNOSTIC" -ForegroundColor White
Write-Host ("=" * 80) -ForegroundColor DarkCyan
Write-Host ""