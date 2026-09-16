<#
.SYNOPSIS
    Audit différentiel exhaustif et en temps réel des habilitations entre deux groupes Jira & Confluence (Mode Dry-Run).
.DESCRIPTION
    - Choix du périmètre : Jira seul, Confluence seul, ou les deux.
    - Utilise dynamiquement les noms réels des groupes dans tous les affichages, logs et recommandations.
    - Scanne les Rôles de projets Jira, Schémas d'autorisations Jira et Espaces Confluence (Moteur natif API v2).
    - Compare les permissions du groupe source avec celles du groupe cible.
    - Restitution en temps réel dans la console et export CSV d'impact.
.VERSION
    2.3 — 2026-09-14
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [string]$Group1,

    [Parameter(Mandatory = $false)]
    [string]$Group2,

    [Parameter(Mandatory = $false)]
    [ValidateSet("JIRA", "CONFLUENCE", "BOTH", "TOUS", "")]
    [string]$Scope = "",

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
$exportsDir = Join-Path $scriptRoot "exports"
$runStamp   = Get-Date -Format "yyyyMMdd_HHmmss"

if (-not (Test-Path $exportsDir)) { New-Item -ItemType Directory -Path $exportsDir -Force | Out-Null }

$ProxyUrl                   = ""
$UseSystemProxy             = $true
$ProxyUseDefaultCredentials = $true

# ============================================================
# 1. HELPERS LOGGING & PROXY REST API
# ============================================================
function Write-Info([string]$Message) { Write-Host ("[INFO] " + $Message) -ForegroundColor Cyan }
function Write-Warn([string]$Message) { Write-Warning $Message }
function Write-ErrLog([string]$Message) { Write-Error $Message }

function Write-DiffFound {
    param(
        [string]$Product,
        [string]$Key,
        [string]$Name,
        [string]$Location,
        [string]$SecurityElement,
        [string]$Status,
        [string]$TargetGroupName,
        [string]$MissingPerms,
        [string]$Action
    )
    $statusColor = if ($Status -like "*DEJA PRESENT*") {
        "Green"
    } elseif ($Status -like "*PARTIELLEMENT*") {
        "Yellow"
    } elseif ($Status -like "*ABSENT*") {
        "Red"
    } else {
        "Cyan"
    }

    Write-Host ""
    Write-Host ("  [" + $Product.ToUpper() + "] ") -NoNewline -ForegroundColor White
    Write-Host ($Key + " (" + $Name + ")") -ForegroundColor Yellow
    Write-Host ("      -> Emplacement  : " + $Location + " | Élément : " + $SecurityElement) -ForegroundColor Gray
    Write-Host ("      -> Diagnostic   : ") -NoNewline -ForegroundColor Gray
    Write-Host ("[" + $Status + "]") -ForegroundColor $statusColor
    if (-not [string]::IsNullOrWhiteSpace($MissingPerms)) {
        Write-Host ("      -> Manquant dans [" + $TargetGroupName + "] : " + $MissingPerms) -ForegroundColor Magenta
    }
    Write-Host ("      -> Action       : " + $Action) -ForegroundColor DarkCyan
}

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
if (-not (Test-Path $CredentialsPath)) { throw "Fichier credentials introuvable : $CredentialsPath" }

$jiraData    = Import-Clixml -Path $CredentialsPath
$jiraBaseUrl = ([string]$jiraData.JiraBaseUrl).TrimEnd("/")
$jiraCred    = $jiraData.Credential
$b64         = [Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes($jiraCred.UserName + ":" + $jiraCred.GetNetworkCredential().Password))
$authHeaders = @{ Authorization = "Basic " + $b64; Accept = "application/json" }

# ============================================================
# 3. SAISIE DES GROUPES, PÉRIMÈTRE & RÉSOLUTION DES GROUP IDs
# ============================================================
Write-Host ""
Write-Host ("=" * 80) -ForegroundColor DarkCyan
Write-Host "  AUDIT DIFFÉRENTIEL DES HABILITATIONS JIRA & CONFLUENCE (DRY-RUN)" -ForegroundColor White
Write-Host ("=" * 80) -ForegroundColor DarkCyan
Write-Host ""

if ([string]::IsNullOrWhiteSpace($Group1)) { $Group1 = Read-Host "Entrez le nom du groupe source (a auditer/remplacer)" }
if ([string]::IsNullOrWhiteSpace($Group2)) { $Group2 = Read-Host "Entrez le nom du groupe cible (remplacement futur)" }

$Group1 = $Group1.Trim()
$Group2 = $Group2.Trim()

if ([string]::IsNullOrWhiteSpace($Group1)) { throw "Le nom du groupe source est obligatoire." }
if ([string]::IsNullOrWhiteSpace($Group2)) { throw "Le nom du groupe cible est obligatoire pour l'analyse differentielle." }
if ($Group1 -ieq $Group2) { throw "Le groupe source et le groupe cible sont identiques ('$Group1')." }

if ([string]::IsNullOrWhiteSpace($Scope)) {
    Write-Host ""
    Write-Host "Quel périmètre souhaitez-vous exécuter ?" -ForegroundColor Cyan
    Write-Host "  [1] Jira uniquement" -ForegroundColor Yellow
    Write-Host "  [2] Confluence uniquement" -ForegroundColor Yellow
    Write-Host "  [3] Les deux (Jira & Confluence)" -ForegroundColor Green
    $scopeChoice = Read-Host "Votre choix [1/2/3] (défaut: 3)"
    
    $Scope = switch -Regex ($scopeChoice.Trim()) {
        "^1|J|Jira"       { "JIRA" }
        "^2|C|Conf"       { "CONFLUENCE" }
        default           { "BOTH" }
    }
} else {
    $Scope = switch ($Scope.ToUpper()) {
        "TOUS"  { "BOTH" }
        default { $Scope.ToUpper() }
    }
}

Write-Host ""
Write-Info ("Périmètre sélectionné          : [" + $Scope + "]")

function Get-JiraGroupId([string]$GroupName) {
    if ([string]::IsNullOrWhiteSpace($GroupName)) { return "" }
    try {
        $url  = $jiraBaseUrl + "/rest/api/3/group/bulk?groupName=" + [Uri]::EscapeDataString($GroupName)
        $resp = Invoke-AtlassianApiGet -Url $url -Headers $authHeaders
        if ($resp -and $resp.values -and $resp.values.Count -gt 0) {
            foreach ($g in $resp.values) {
                if ([string]$g.name -ieq $GroupName -and -not [string]::IsNullOrWhiteSpace($g.groupId)) {
                    return [string]$g.groupId
                }
            }
            if ($resp.values[0].groupId) { return [string]$resp.values[0].groupId }
        }
    } catch {}

    try {
        $urlFind  = $jiraBaseUrl + "/rest/api/3/groups/picker?query=" + [Uri]::EscapeDataString($GroupName)
        $respFind = Invoke-AtlassianApiGet -Url $urlFind -Headers $authHeaders
        if ($respFind -and $respFind.groups) {
            foreach ($g in $respFind.groups) {
                if ([string]$g.name -ieq $GroupName -and -not [string]::IsNullOrWhiteSpace($g.groupId)) {
                    return [string]$g.groupId
                }
            }
        }
    } catch {}

    return ""
}

Write-Info "Resolution des identifiants Cloud (Group ID / UUID)..."

$groupId1 = Get-JiraGroupId -GroupName $Group1
$groupId2 = Get-JiraGroupId -GroupName $Group2

$g1DisplayId = if (-not [string]::IsNullOrWhiteSpace($groupId1)) { $groupId1 } else { "Nom seul" }
$g2DisplayId = if (-not [string]::IsNullOrWhiteSpace($groupId2)) { $groupId2 } else { "Nom seul" }

Write-Host ("  [Groupe Source] : ") -NoNewline -ForegroundColor White
Write-Host ($Group1) -NoNewline -ForegroundColor Yellow
Write-Host ("  (GroupId: " + $g1DisplayId + ")") -ForegroundColor DarkGray

Write-Host ("  [Groupe Cible]  : ") -NoNewline -ForegroundColor White
Write-Host ($Group2) -NoNewline -ForegroundColor Yellow
Write-Host ("  (GroupId: " + $g2DisplayId + ")") -ForegroundColor DarkGray
Write-Host ""

$auditFindings = New-Object System.Collections.ArrayList
$projects      = New-Object System.Collections.ArrayList

# ============================================================
# 4. AUDIT JIRA — PROJETS & RÔLES DE PROJETS
# ============================================================
if ($Scope -eq "JIRA" -or $Scope -eq "BOTH") {
    Write-Info ("=== 1/3 Audit des Roles de projets Jira (" + $Group1 + " vs " + $Group2 + ") ===")

    $startAt = 0
    $maxRes  = 50
    $isLast  = $false

    do {
        $pUrl = $jiraBaseUrl + "/rest/api/3/project/search?startAt=" + $startAt + "&maxResults=" + $maxRes + "&expand=permissionscheme"
        try {
            $pResp = Invoke-AtlassianApiGet -Url $pUrl -Headers $authHeaders
            if ($pResp.values) {
                foreach ($p in $pResp.values) { [void]$projects.Add($p) }
            }
            $isLast  = if ($pResp.isLast -ne $null) { [bool]$pResp.isLast } else { ($pResp.values.Count -lt $maxRes) }
            $startAt += $maxRes
        } catch {
            Write-Warn ("Erreur recherche projets Jira : " + $_.Exception.Message)
            break
        }
    } while (-not $isLast)

    Write-Info ("  -> " + $projects.Count + " projets Jira detectes. Scan des roles...")

    $pIdx = 0
    foreach ($proj in $projects) {
        $pIdx++
        $pKey  = [string]$proj.key
        $pName = Fix-Encoding ([string]$proj.name)

        Write-Progress -Id 1 -Activity ("Scan Roles Jira (" + $Group1 + ")") `
            -Status ("Projet " + $pIdx + "/" + $projects.Count + " : " + $pKey + " (" + $pName + ")") `
            -PercentComplete ([math]::Round($pIdx / $projects.Count * 100))

        try {
            $rolesDictUrl = $jiraBaseUrl + "/rest/api/3/project/" + $pKey + "/role"
            $rolesDict    = Invoke-AtlassianApiGet -Url $rolesDictUrl -Headers $authHeaders
            
            if ($rolesDict) {
                foreach ($prop in $rolesDict.PSObject.Properties) {
                    $roleName = Fix-Encoding ([string]$prop.Name)
                    $roleUrl  = [string]$prop.Value

                    try {
                        $roleDetail = Invoke-AtlassianApiGet -Url $roleUrl -Headers $authHeaders
                        if ($roleDetail -and $roleDetail.actors) {
                            $hasG1 = $false
                            $hasG2 = $false
                            $actorG1Detail = ""
                            $actorG2Detail = ""

                            foreach ($actor in $roleDetail.actors) {
                                $actorName = Fix-Encoding ([string]$actor.name)
                                $actorDisp = Fix-Encoding ([string]$actor.displayName)
                                $grpName   = if ($actor.actorGroup -and $actor.actorGroup.name) { [string]$actor.actorGroup.name } else { "" }

                                if ($actorName -ieq $Group1 -or $actorDisp -ieq $Group1 -or $grpName -ieq $Group1) {
                                    $hasG1 = $true
                                    $actorG1Detail = $actorDisp
                                }
                                if ($actorName -ieq $Group2 -or $actorDisp -ieq $Group2 -or $grpName -ieq $Group2) {
                                    $hasG2 = $true
                                    $actorG2Detail = $actorDisp
                                }
                            }

                            if ($hasG1) {
                                $diffStatus   = if ($hasG2) { ($Group2 + " DEJA PRESENT") } else { ($Group2 + " ABSENT") }
                                $missingPerms = if ($hasG2) { "" } else { "Role entier '$roleName'" }
                                $action       = if ($hasG2) {
                                    ("Retirer uniquement '" + $Group1 + "' du role ('" + $Group2 + "' deja present)")
                                } else {
                                    ("Ajouter '" + $Group2 + "' dans le role puis retirer '" + $Group1 + "'")
                                }

                                Write-DiffFound -Product "Jira" -Key $pKey -Name $pName -Location "Role de Projet" `
                                    -SecurityElement $roleName -Status $diffStatus -TargetGroupName $Group2 -MissingPerms $missingPerms -Action $action

                                [void]$auditFindings.Add([pscustomobject]@{
                                    Produit                = "Jira"
                                    EspaceCle              = $pKey
                                    EspaceNom              = $pName
                                    TypeEmplacement        = "Role de Projet"
                                    ElementSecurite        = $roleName
                                    GroupeSource           = $Group1
                                    GroupeCible            = $Group2
                                    StatutComparaison      = $diffStatus
                                    PermissionsSource      = "Membre du role"
                                    PermissionsCible       = if ($hasG2) { "Membre du role" } else { "Non membre" }
                                    PermissionsManquantesCible = $missingPerms
                                    ActionRecommandee      = $action
                                })
                            }
                        }
                    } catch {}
                }
            }
        } catch {}
    }
    Write-Progress -Id 1 -Activity ("Scan Roles Jira (" + $Group1 + ")") -Completed

    # ============================================================
    # 5. AUDIT JIRA — SCHÉMAS D'AUTORISATIONS
    # ============================================================
    Write-Host ""
    Write-Info ("=== 2/3 Audit des Schemas d'autorisations Jira (" + $Group1 + " vs " + $Group2 + ") ===")

    try {
        $schemesUrl  = $jiraBaseUrl + "/rest/api/3/permissionscheme?expand=permissions,user,group"
        $schemesResp = Invoke-AtlassianApiGet -Url $schemesUrl -Headers $authHeaders

        if ($schemesResp.permissionSchemes) {
            foreach ($ps in $schemesResp.permissionSchemes) {
                $psName = Fix-Encoding ([string]$ps.name)
                $psId   = [string]$ps.id

                if ($ps.permissions) {
                    $permsG1 = New-Object System.Collections.ArrayList
                    $permsG2 = New-Object System.Collections.ArrayList

                    foreach ($perm in $ps.permissions) {
                        $hGroup = ""
                        if ($perm.holder) {
                            if ($perm.holder.parameter)                             { $hGroup = [string]$perm.holder.parameter }
                            elseif ($perm.holder.value)                             { $hGroup = [string]$perm.holder.value }
                            elseif ($perm.holder.group -and $perm.holder.group.name) { $hGroup = [string]$perm.holder.group.name }
                        }

                        if ($hGroup -ieq $Group1) { [void]$permsG1.Add([string]$perm.permission) }
                        if ($hGroup -ieq $Group2) { [void]$permsG2.Add([string]$perm.permission) }
                    }

                    if ($permsG1.Count -gt 0) {
                        $distinctG1  = @($permsG1 | Select-Object -Unique)
                        $distinctG2  = @($permsG2 | Select-Object -Unique)
                        $missingInG2 = @($distinctG1 | Where-Object { $distinctG2 -notcontains $_ })

                        $diffStatus = ($Group2 + " ABSENT")
                        if ($missingInG2.Count -eq 0) {
                            $diffStatus = ($Group2 + " DEJA PRESENT")
                        } elseif ($distinctG2.Count -gt 0) {
                            $diffStatus = ($Group2 + " PARTIELLEMENT VU")
                        }

                        $missingStr = if ($missingInG2.Count -gt 0) { $missingInG2 -join ", " } else { "" }
                        $action     = if ($diffStatus -eq ($Group2 + " DEJA PRESENT")) {
                            ("Retirer '" + $Group1 + "' du schema ('" + $Group2 + "' a deja toutes les permissions)")
                        } elseif ($diffStatus -eq ($Group2 + " PARTIELLEMENT VU")) {
                            ("Accorder les permissions manquantes a '" + $Group2 + "' puis retirer '" + $Group1 + "'")
                        } else {
                            ("Accorder les permissions a '" + $Group2 + "' dans le schema puis retirer '" + $Group1 + "'")
                        }

                        $linkedProjects = New-Object System.Collections.ArrayList
                        foreach ($proj in $projects) {
                            if ($proj.permissionScheme -and [string]$proj.permissionScheme.id -eq $psId) {
                                [void]$linkedProjects.Add($proj.key)
                            }
                        }
                        $pListStr = if ($linkedProjects.Count -gt 0) { $linkedProjects -join ", " } else { "(Aucun projet directement assigne)" }

                        Write-DiffFound -Product "Jira" -Key $pListStr -Name ("Schema : " + $psName) -Location "Permission Scheme" `
                            -SecurityElement $psName -Status $diffStatus -TargetGroupName $Group2 -MissingPerms $missingStr -Action $action

                        [void]$auditFindings.Add([pscustomobject]@{
                            Produit                = "Jira"
                            EspaceCle              = $pListStr
                            EspaceNom              = ("Projets lies au schema : " + $psName)
                            TypeEmplacement        = "Permission Scheme"
                            ElementSecurite        = $psName
                            GroupeSource           = $Group1
                            GroupeCible            = $Group2
                            StatutComparaison      = $diffStatus
                            PermissionsSource      = ($distinctG1 -join ", ")
                            PermissionsCible       = ($distinctG2 -join ", ")
                            PermissionsManquantesCible = $missingStr
                            ActionRecommandee      = $action
                        })
                    }
                }
            }
        }
    } catch {
        Write-Warn ("Erreur audit Permission Schemes : " + $_.Exception.Message)
    }
} else {
    Write-Info "Périmètre Jira ignoré (exécuté en mode CONFLUENCE seul)."
}

# ============================================================
# 6. AUDIT CONFLUENCE — ESPACES (MOTEUR NATIF API V2)
# ============================================================
if ($Scope -eq "CONFLUENCE" -or $Scope -eq "BOTH") {
    Write-Host ""
    Write-Info ("=== 3/3 Audit des Espaces Confluence (" + $Group1 + " vs " + $Group2 + ") ===")

    $confSpaces  = New-Object System.Collections.ArrayList
    $nextCursor  = ""
    $hasMoreList = $true
    $batchNum    = 0

    do {
        $batchNum++
        $spacesUrl = $jiraBaseUrl + "/wiki/api/v2/spaces?limit=250&status=current"
        if (-not [string]::IsNullOrWhiteSpace($nextCursor)) {
            $spacesUrl += "&cursor=" + [Uri]::EscapeDataString($nextCursor)
        }

        try {
            $spResp = Invoke-AtlassianApiGet -Url $spacesUrl -Headers $authHeaders
            if ($spResp.results) {
                foreach ($sp in $spResp.results) { [void]$confSpaces.Add($sp) }
            }
            if ($spResp._links -and $spResp._links.next) {
                $match = [regex]::Match($spResp._links.next, "cursor=([^&]+)")
                if ($match.Success) { $nextCursor = $match.Groups[1].Value } else { $hasMoreList = $false }
            } else {
                $hasMoreList = $false
            }
        } catch {
            Write-Warn ("Erreur récupération liste espaces v2 : " + $_.Exception.Message)
            $hasMoreList = $false
        }
    } while ($hasMoreList)

    Write-Info ("  -> Total : " + $confSpaces.Count + " espaces Confluence détectés. Scan des permissions...")

    $cIdx = 0
    foreach ($sp in $confSpaces) {
        $cIdx++
        $sKey  = [string]$sp.key
        $sName = Fix-Encoding ([string]$sp.name)
        $sId   = [string]$sp.id

        Write-Progress -Id 2 -Activity ("Scan Espaces Confluence (" + $Group1 + ")") `
            -Status ("Espace " + $cIdx + "/" + $confSpaces.Count + " : " + $sKey + " (" + $sName + ")") `
            -PercentComplete ([math]::Round($cIdx / $confSpaces.Count * 100))

        $permsG1 = New-Object System.Collections.ArrayList
        $permsG2 = New-Object System.Collections.ArrayList

        try {
            $permCursor = ""
            $hasMorePerms = $true

            do {
                $permUrl = $jiraBaseUrl + "/wiki/api/v2/spaces/" + $sId + "/permissions?limit=250"
                if (-not [string]::IsNullOrWhiteSpace($permCursor)) {
                    $permUrl += "&cursor=" + [Uri]::EscapeDataString($permCursor)
                }

                $pResp = Invoke-AtlassianApiGet -Url $permUrl -Headers $authHeaders
                if ($pResp.results) {
                    foreach ($pItem in $pResp.results) {
                        if ($pItem.principal -and [string]$pItem.principal.type -ieq "group") {
                            $grpIdVal = [string]$pItem.principal.id
                            $opKey    = [string]$pItem.operation.key
                            $opTarget = [string]$pItem.operation.target
                            $opFull   = if (-not [string]::IsNullOrWhiteSpace($opTarget)) { "$opKey ($opTarget)" } else { $opKey }

                            if (($groupId1 -and $grpIdVal -ieq $groupId1) -or $grpIdVal -ieq $Group1) {
                                [void]$permsG1.Add($opFull)
                            }
                            if (($groupId2 -and $grpIdVal -ieq $groupId2) -or $grpIdVal -ieq $Group2) {
                                [void]$permsG2.Add($opFull)
                            }
                        }
                    }
                }

                if ($pResp._links -and $pResp._links.next) {
                    $m = [regex]::Match($pResp._links.next, "cursor=([^&]+)")
                    if ($m.Success) { $permCursor = $m.Groups[1].Value } else { $hasMorePerms = $false }
                } else {
                    $hasMorePerms = $false
                }
            } while ($hasMorePerms)

        } catch {
            # Erreur 403 / Espace restreint
        }

        if ($permsG1.Count -gt 0) {
            $distinctG1  = @($permsG1 | Select-Object -Unique)
            $distinctG2  = @($permsG2 | Select-Object -Unique)
            $missingInG2 = @($distinctG1 | Where-Object { $distinctG2 -notcontains $_ })

            $diffStatus = ($Group2 + " ABSENT")
            if ($missingInG2.Count -eq 0) {
                $diffStatus = ($Group2 + " DEJA PRESENT")
            } elseif ($distinctG2.Count -gt 0) {
                $diffStatus = ($Group2 + " PARTIELLEMENT VU")
            }

            $missingStr = if ($missingInG2.Count -gt 0) { $missingInG2 -join ", " } else { "" }
            $action     = if ($diffStatus -eq ($Group2 + " DEJA PRESENT")) {
                ("Retirer '" + $Group1 + "' de l'espace ('" + $Group2 + "' a deja toutes les permissions)")
            } elseif ($diffStatus -eq ($Group2 + " PARTIELLEMENT VU")) {
                ("Ajouter sur '" + $Group2 + "' les permissions manquantes puis retirer '" + $Group1 + "'")
            } else {
                ("Ajouter '" + $Group2 + "' avec les permissions de '" + $Group1 + "' puis retirer '" + $Group1 + "'")
            }

            Write-DiffFound -Product "Confluence" -Key $sKey -Name $sName -Location "Space Permission" `
                -SecurityElement "Droits d'espace" -Status $diffStatus -TargetGroupName $Group2 -MissingPerms $missingStr -Action $action

            [void]$auditFindings.Add([pscustomobject]@{
                Produit                = "Confluence"
                EspaceCle              = $sKey
                EspaceNom              = $sName
                TypeEmplacement        = "Space Permission"
                ElementSecurite        = "Droits d'espace"
                GroupeSource           = $Group1
                GroupeCible            = $Group2
                StatutComparaison      = $diffStatus
                PermissionsSource      = ($distinctG1 -join ", ")
                PermissionsCible       = ($distinctG2 -join ", ")
                PermissionsManquantesCible = $missingStr
                ActionRecommandee      = $action
            })
        }
    }
    Write-Progress -Id 2 -Activity ("Scan Espaces Confluence (" + $Group1 + ")") -Completed
} else {
    Write-Info "Périmètre Confluence ignoré (exécuté en mode JIRA seul)."
}

# ============================================================
# 7. BILAN FINAL ET STATISTIQUES
# ============================================================
Write-Host ""
Write-Host ("=" * 80) -ForegroundColor Green
Write-Host ("  SYNTHÈSE DU PLAN DE BASCULE : [" + $Group1 + "] -> [" + $Group2 + "]") -ForegroundColor Green
Write-Host ("  Périmètre audité            : " + $Scope) -ForegroundColor Gray
Write-Host ("=" * 80) -ForegroundColor Green

$nbTotal   = $auditFindings.Count
$nbDejaOK  = ($auditFindings | Where-Object { $_.StatutComparaison -like "*DEJA PRESENT*" }).Count
$nbPartiel = ($auditFindings | Where-Object { $_.StatutComparaison -like "*PARTIELLEMENT*" }).Count
$nbAbsent  = ($auditFindings | Where-Object { $_.StatutComparaison -like "*ABSENT*" }).Count

Write-Host ""
Write-Host ("  Total occurrences [" + $Group1 + "] trouvées : " + $nbTotal) -ForegroundColor White
Write-Host ("    -> 🟢 " + $Group2 + " Déjà présent à 100%      : " + $nbDejaOK  + " (retrait direct de " + $Group1 + ")") -ForegroundColor Green
Write-Host ("    -> 🟡 " + $Group2 + " Présent mais incomplet   : " + $nbPartiel + " (permissions à compléter sur " + $Group2 + ")") -ForegroundColor Yellow
Write-Host ("    -> 🔴 " + $Group2 + " Totalement absent        : " + $nbAbsent  + " (ajout préalable de " + $Group2 + " avant retrait)") -ForegroundColor Red
Write-Host ""

if ($nbTotal -gt 0) {
    $auditFindings | Format-Table -Property @(
        @{Label="Produit"; Expression={$_.Produit}; Width=11},
        @{Label="Espace / Clé"; Expression={$_.EspaceCle}; Width=18},
        @{Label="Emplacement"; Expression={$_.TypeEmplacement}; Width=17},
        @{Label="Diagnostic"; Expression={$_.StatutComparaison}; Width=25},
        @{Label="Action Recommandée"; Expression={$_.ActionRecommandee}; Width=42}
    ) | Out-Host
} else {
    Write-Host ("Aucune habilitation directe trouvée pour le groupe '" + $Group1 + "' sur le périmètre " + $Scope + ".") -ForegroundColor Yellow
}

# ============================================================
# 8. EXPORT CSV DIFFÉRENTIEL (Dry-Run)
# ============================================================
$sanitizedG1 = $Group1 -replace '[\\/:*?"<>|]', '_'
$sanitizedG2 = $Group2 -replace '[\\/:*?"<>|]', '_'
$csvPath     = Join-Path $exportsDir ("Plan_Bascule_" + $Scope + "_" + $sanitizedG1 + "_VERS_" + $sanitizedG2 + "_" + $runStamp + ".csv")

$utf8Bom = New-Object System.Text.UTF8Encoding($true)
$writer  = New-Object System.IO.StreamWriter($csvPath, $false, $utf8Bom)
try {
    $writer.WriteLine("Produit;EspaceCle;EspaceNom;TypeEmplacement;ElementSecurite;GroupeSource;GroupeCible;StatutComparaison;PermissionsSource;PermissionsCible;PermissionsManquantesCible;ActionRecommandee")
    foreach ($r in ($auditFindings | Sort-Object StatutComparaison, Produit, EspaceCle)) {
        $line = ('"{0}";"{1}";"{2}";"{3}";"{4}";"{5}";"{6}";"{7}";"{8}";"{9}";"{10}";"{11}"' -f `
            $r.Produit, $r.EspaceCle, $r.EspaceNom, $r.TypeEmplacement, $r.ElementSecurite, $r.GroupeSource, $r.GroupeCible, `
            $r.StatutComparaison, $r.PermissionsSource, $r.PermissionsCible, $r.PermissionsManquantesCible, $r.ActionRecommandee)
        $writer.WriteLine($line)
    }
} finally {
    $writer.Close()
    $writer.Dispose()
}

Write-Host ""
Write-Info ("Fichier CSV du plan de bascule généré : " + $csvPath)
Write-Host ("=" * 80) -ForegroundColor DarkCyan