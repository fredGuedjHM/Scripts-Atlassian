<#
.SYNOPSIS
    Audit ultra-rapide et parallélisé des accès Jira & Confluence pour les collaborateurs hors DSIM / DSIT.
.DESCRIPTION
    - Filtrage strict sur les domaines :
        * @harmonie-mutuelle.fr
        * @prestataire.sihm.fr
        * @prestataire.harmonie-mutuelle.fr
    - Exclusion des comptes techniques / fantômes sans aucun groupe ni droit.
    - Exclusion automatique en mémoire O(1) des membres de groupes DSIM/DSIT.
    - Tri alphabétique systématique par DisplayName.
    - Confluence simplifié : Espaces globaux uniquement (personnels ~ exclus), format "CLE [C|R|G]".
    - Parallélisation multi-threads (RunspacePool x15).
    - Export CSV UTF-8 avec BOM (compatible Excel).
#>

[CmdletBinding()]
param(
    [string]$ProxyUrl = "http://prc37cti1.hm.dm.ad:8080",
    [int]$ThrottleLimit = 15
)

# ======================================================================
# 1. INITIALISATION PROXY, TLS & ENCODAGE (COMPATIBLE ISE & CONSOLE)
# ======================================================================
[System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls12 -bor [System.Net.SecurityProtocolType]::Tls13

try {
    [Console]::OutputEncoding = [System.Text.Encoding]::UTF8
} catch {}

if (-not $ProxyUrl) {
    try {
        $sysProxy = [System.Net.WebRequest]::GetSystemWebProxy()
        $testUri = New-Object System.Uri("https://api.atlassian.com")
        $pUri = $sysProxy.GetProxy($testUri)
        if ($pUri -and $pUri.AbsoluteUri -ne $testUri.AbsoluteUri) {
            $ProxyUrl = $pUri.AbsoluteUri
        }
    } catch {}
}

function Repair-DoubleUtf8([string]$str) {
    if ([string]::IsNullOrWhiteSpace($str)) { return "" }
    try {
        if ($str -match '[\xC2-\xC3][\x80-\xBF]') {
            $bytes = [System.Text.Encoding]::GetEncoding("ISO-8859-1").GetBytes($str)
            return [System.Text.Encoding]::UTF8.GetString($bytes)
        }
    } catch {}
    return $str
}

function Invoke-SafeGet {
    param([string]$Url, [hashtable]$Headers)
    $params = @{
        Uri     = $Url
        Method  = "Get"
        Headers = $Headers
    }
    if ($ProxyUrl) {
        $params["Proxy"] = $ProxyUrl
        $params["ProxyUseDefaultCredentials"] = $true
    }
    return (Invoke-RestMethod @params)
}

# ======================================================================
# 2. DOSSIERS & CREDENTIALS
# ======================================================================
$scriptDir  = if ($PSCommandPath) { Split-Path -Parent $PSCommandPath } else { (Get-Location).Path }
$secretsDir = Join-Path $scriptDir "secrets"
$exportsDir = Join-Path $scriptDir "exports"

if (-not (Test-Path $exportsDir)) { [void](New-Item -ItemType Directory -Path $exportsDir -Force) }

function Write-Info($msg) { Write-Host "[$((Get-Date).ToString('yyyy-MM-dd HH:mm:ss'))] [INFO] $msg" -ForegroundColor Cyan }
function Write-Warn($msg) { Write-Host "[$((Get-Date).ToString('yyyy-MM-dd HH:mm:ss'))] [WARN] $msg" -ForegroundColor Yellow }
function Write-ErrLog($msg) { Write-Host "[$((Get-Date).ToString('yyyy-MM-dd HH:mm:ss'))] [ERROR] $msg" -ForegroundColor Red }

$jiraCredFile = Join-Path $secretsDir "jira-jiradot.cred.xml"
if (-not (Test-Path $jiraCredFile)) { throw "Fichier Jira creds introuvable: $jiraCredFile" }

$jiraData    = Import-Clixml -Path $jiraCredFile
$baseUrl     = ($jiraData.JiraBaseUrl).TrimEnd("/")
$jiraCred    = $jiraData.Credential
$jiraEmail   = $jiraCred.UserName
$jiraToken   = $jiraCred.GetNetworkCredential().Password

$pair        = [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes("${jiraEmail}:${jiraToken}"))
$authHeaders = @{
    "Authorization" = "Basic $pair"
    "Accept"        = "application/json"
}

Write-Info "Connexion Atlassian : $baseUrl (user=$jiraEmail)"
if ($ProxyUrl) { Write-Info "Proxy actif : $ProxyUrl" }

$totalSteps = 5

# ======================================================================
# ÉTAPE 1/5 : DÉTECTION DES GROUPES DSIM & DSIT (POUR EXCLUSION)
# ======================================================================
Write-Progress -Activity "Audit Collaborateurs hors DSIM/DSIT" -Status "Étape 1/$totalSteps : Recherche des groupes DSIM et DSIT..." -PercentComplete 10
Write-Info "Recherche des groupes contenant DSIM ou DSIT..."

$dsiGroupNames = New-Object System.Collections.Generic.HashSet[string]
$searchQueries = @("DSIM", "DSIT", "dsim", "dsit")

foreach ($q in $searchQueries) {
    try {
        $pUrl = "$baseUrl/rest/api/3/groups/picker?query=$q&maxResults=500"
        $pResp = Invoke-SafeGet -Url $pUrl -Headers $authHeaders
        if ($pResp -and $pResp.groups) {
            foreach ($g in $pResp.groups) {
                $gName = Repair-DoubleUtf8 ([string]$g.name)
                if ($gName -match '(?i)(DSIM|DSIT)') {
                    [void]$dsiGroupNames.Add($gName)
                }
            }
        }
    } catch {}
}

Write-Info "$($dsiGroupNames.Count) groupes DSIM/DSIT identifiés pour l'exclusion :"
foreach ($gn in ($dsiGroupNames | Sort-Object)) {
    Write-Host "   -> $gn" -ForegroundColor DarkGray
}

# ======================================================================
# ÉTAPE 2/5 : RÉCUPÉRATION DES MEMBRES DSIM/DSIT (INDEX O(1))
# ======================================================================
$dsiUserAccountIds = New-Object System.Collections.Generic.HashSet[string]
$gIdx = 0

foreach ($gName in $dsiGroupNames) {
    $gIdx++
    $pct = [Math]::Round(($gIdx / [Math]::Max(1, $dsiGroupNames.Count)) * 100)
    Write-Progress -Activity "Audit Collaborateurs hors DSIM/DSIT" `
                   -Status "Étape 2/$totalSteps : Indexation membres DSIM/DSIT ($gIdx / $($dsiGroupNames.Count))" `
                   -PercentComplete ([int](10 + ($pct * 0.15))) `
                   -CurrentOperation "$gName"

    $startAt = 0
    while ($true) {
        $encGroup = [Uri]::EscapeDataString($gName)
        $gmUrl = "$baseUrl/rest/api/3/group/member?groupname=$encGroup&startAt=$startAt&maxResults=50&includeInactiveUsers=false"
        try {
            $gmResp = Invoke-SafeGet -Url $gmUrl -Headers $authHeaders
            if (-not $gmResp -or -not $gmResp.values -or $gmResp.values.Count -eq 0) { break }
            foreach ($u in $gmResp.values) {
                if ($u.accountId) { [void]$dsiUserAccountIds.Add([string]$u.accountId) }
            }
            if ($gmResp.isLast -or $gmResp.values.Count -lt 50) { break }
            $startAt += $gmResp.values.Count
        } catch { break }
    }
}

Write-Info "$($dsiUserAccountIds.Count) utilisateurs uniques DSIM/DSIT indexés (exclus automatiquement)."

# ======================================================================
# ÉTAPE 3/5 : EXTRACTION CIBLÉE & PARALLÉLISATION DES GROUPES
# ======================================================================
Write-Progress -Activity "Audit Collaborateurs hors DSIM/DSIT" -Status "Étape 3/$totalSteps : Extraction des utilisateurs actifs..." -PercentComplete 30
Write-Info "Extraction des comptes actifs hors DSIM/DSIT..."

$targetCandidates = New-Object System.Collections.ArrayList
$startAt = 0; $maxRes = 50

# Domaines autorisés
$validEmailRegex = '(?i)@(harmonie-mutuelle\.fr|prestataire\.sihm\.fr|prestataire\.harmonie-mutuelle\.fr)$'
# Domaines parasites à rejeter immédiatement
$forbiddenDomainRegex = '(?i)@(groupevyv\.onmicrosoft\.com|atlassian\.com|bot|automation)'

while ($true) {
    $uUrl = "$baseUrl/rest/api/3/users/search?startAt=$startAt&maxResults=$maxRes"
    try {
        $uResp = Invoke-SafeGet -Url $uUrl -Headers $authHeaders
    } catch { break }

    if (-not $uResp -or $uResp.Count -eq 0) { break }

    foreach ($u in $uResp) {
        if (-not $u.active -or ($u.accountType -and $u.accountType -ine "atlassian")) { continue }

        $accId = [string]$u.accountId
        if ($dsiUserAccountIds.Contains($accId)) { continue }

        $dispName = Repair-DoubleUtf8 ([string]$u.displayName)
        $email    = [string]$u.emailAddress

        # Rejet immédiat si le DisplayName est une adresse hors domaine cible (ex: gpafi@groupevyv...)
        if ($dispName -match $forbiddenDomainRegex) { continue }
        if ($email -and $email -match $forbiddenDomainRegex) { continue }

        [void]$targetCandidates.Add([pscustomobject]@{
            AccountId   = $accId
            DisplayName = $dispName
            Email       = $email
            Groups      = @()
        })
    }

    if ($uResp.Count -lt $maxRes) { break }
    $startAt += $uResp.Count
}

Write-Info "$($targetCandidates.Count) candidats actifs à analyser..."

if ($targetCandidates.Count -eq 0) {
    Write-Progress -Activity "Audit Collaborateurs" -Completed
    Write-Warn "Aucun candidat trouvé."
    exit 0
}

Write-Info "Résolution multi-threadée des groupes et vérification des profils ($ThrottleLimit threads)..."

$workerUserBlock = {
    param(
        [string]$AccountId,
        [string]$BaseUrl,
        [hashtable]$Headers,
        [string]$Proxy,
        [bool]$UseDefaultProxyAuth
    )

    function Repair-TextEncoding([string]$str) {
        if ([string]::IsNullOrWhiteSpace($str)) { return "" }
        try {
            if ($str -match '[\xC2-\xC3][\x80-\xBF]') {
                $b = [System.Text.Encoding]::GetEncoding("ISO-8859-1").GetBytes($str)
                return [System.Text.Encoding]::UTF8.GetString($b)
            }
        } catch {}
        return $str
    }

    $groups = @()
    $email  = ""

    $params = @{
        Uri     = "$BaseUrl/rest/api/3/user?accountId=$AccountId&expand=groups"
        Method  = "Get"
        Headers = $Headers
    }
    if ($Proxy) {
        $params["Proxy"] = $Proxy
        $params["ProxyUseDefaultCredentials"] = $UseDefaultProxyAuth
    }

    try {
        $resp = Invoke-RestMethod @params
        if ($resp.emailAddress) { $email = [string]$resp.emailAddress }
        if ($resp.groups -and $resp.groups.items) {
            $groups = @($resp.groups.items | ForEach-Object { Repair-TextEncoding ([string]$_.name) })
        }
    } catch {}

    return [pscustomobject]@{
        AccountId = $AccountId
        Email     = $email
        Groups    = $groups
    }
}

$userPool = [runspacefactory]::CreateRunspacePool(1, $ThrottleLimit)
$userPool.Open()

$tasks = New-Object System.Collections.ArrayList
foreach ($tc in $targetCandidates) {
    $ps = [powershell]::Create()
    $ps.RunspacePool = $userPool
    [void]$ps.AddScript($workerUserBlock)
    [void]$ps.AddArgument($tc.AccountId)
    [void]$ps.AddArgument($baseUrl)
    [void]$ps.AddArgument($authHeaders)
    [void]$ps.AddArgument($ProxyUrl)
    [void]$ps.AddArgument($true)

    [void]$tasks.Add([pscustomobject]@{
        Candidate   = $tc
        PowerShell  = $ps
        AsyncResult = $ps.BeginInvoke()
    })
}

$doneCount = 0
$targetUsers = New-Object System.Collections.ArrayList
$allTargetGroupNamesLower = New-Object System.Collections.Generic.HashSet[string]

while ($doneCount -lt $tasks.Count) {
    Start-Sleep -Milliseconds 50
    $completedThisTurn = 0

    foreach ($t in $tasks) {
        if ($null -ne $t.AsyncResult -and $t.AsyncResult.IsCompleted) {
            $completedThisTurn++
            $result = $t.PowerShell.EndInvoke($t.AsyncResult)
            $t.PowerShell.Dispose()
            $t.AsyncResult = $null

            if ($result -and $result[0]) {
                $resObj = $result[0]
                $resolvedEmail = if ($resObj.Email) { $resObj.Email } else { $t.Candidate.Email }
                $userGroups    = @($resObj.Groups)

                # RÈGLE DE FILTRAGE STRICT :
                # 1. Si un email est connu, il doit appartenir aux domaines HM/SIHM
                # 2. Rejet des comptes rejetés par regex ou sans groupe Jira
                $isValidMail = ($resolvedEmail -match $validEmailRegex)
                $isForbidden = ($resolvedEmail -match $forbiddenDomainRegex) -or ($t.Candidate.DisplayName -match $forbiddenDomainRegex)

                if ($isValidMail -and -not $isForbidden) {
                    $t.Candidate.Email = $resolvedEmail
                    $t.Candidate.Groups = $userGroups
                    [void]$targetUsers.Add($t.Candidate)

                    foreach ($g in $userGroups) {
                        [void]$allTargetGroupNamesLower.Add($g.Trim().ToLower())
                    }
                }
            }
            $doneCount++
        }
    }

    if ($completedThisTurn -gt 0 -or $doneCount -eq $tasks.Count) {
        $pct = [Math]::Round(($doneCount / $tasks.Count) * 100)
        Write-Progress -Activity "Audit Collaborateurs hors DSIM/DSIT" `
                       -Status "Étape 3/$totalSteps : Groupes utilisateurs ($doneCount / $($tasks.Count))" `
                       -PercentComplete ([int](30 + ($pct * 0.25))) `
                       -CurrentOperation "$pct % terminé"
    }
}

$userPool.Close()
$userPool.Dispose()

Write-Info "Collaborateurs HM/Prestataires validés : $($targetUsers.Count)"

if ($targetUsers.Count -eq 0) {
    Write-Warn "Aucun collaborateur correspondant aux critères."
    exit 0
}

# ======================================================================
# ÉTAPE 4/5 : CARTOGRAPHIE PROJETS JIRA & RÔLES (PARALLÉLISÉE)
# ======================================================================
Write-Progress -Activity "Audit Collaborateurs hors DSIM/DSIT" -Status "Étape 4/$totalSteps : Projets Jira parallélisés..." -PercentComplete 55
Write-Info "Analyse multi-threadée des rôles sur les projets Jira..."

$projRoleByAccId = @{}
$projRoleByGroup = @{}
$targetAccIdsList = @($targetUsers | ForEach-Object { $_.AccountId })
$targetGroupsList = @($allTargetGroupNamesLower)

$projects = @()
try {
    $projects = @(Invoke-SafeGet -Url "$baseUrl/rest/api/3/project" -Headers $authHeaders)
} catch {}

if ($projects.Count -gt 0) {
    $workerProjectBlock = {
        param(
            [string]$ProjectKey,
            [string]$BaseUrl,
            [hashtable]$Headers,
            [string]$Proxy,
            [bool]$UseDefaultProxyAuth,
            [array]$TargetAccIds,
            [array]$TargetGroupsLower
        )

        function Repair-TextEncoding([string]$str) {
            if ([string]::IsNullOrWhiteSpace($str)) { return "" }
            try {
                if ($str -match '[\xC2-\xC3][\x80-\xBF]') {
                    $b = [System.Text.Encoding]::GetEncoding("ISO-8859-1").GetBytes($str)
                    return [System.Text.Encoding]::UTF8.GetString($b)
                }
            } catch {}
            return $str
        }

        $targetAccSet = New-Object System.Collections.Generic.HashSet[string]
        if ($TargetAccIds) { foreach ($id in $TargetAccIds) { [void]$targetAccSet.Add([string]$id) } }

        $targetGrpSet = New-Object System.Collections.Generic.HashSet[string]
        if ($TargetGroupsLower) { foreach ($g in $TargetGroupsLower) { [void]$targetGrpSet.Add([string]$g) } }

        $userRoles  = New-Object System.Collections.ArrayList
        $groupRoles = New-Object System.Collections.ArrayList

        $params = @{ Method = "Get"; Headers = $Headers }
        if ($Proxy) {
            $params["Proxy"] = $Proxy
            $params["ProxyUseDefaultCredentials"] = $UseDefaultProxyAuth
        }

        try {
            $params["Uri"] = "$BaseUrl/rest/api/3/project/$ProjectKey/role"
            $rolesDict = Invoke-RestMethod @params
            foreach ($prop in $rolesDict.PSObject.Properties) {
                $roleName = Repair-TextEncoding ([string]$prop.Name)
                $roleUrl  = [string]$prop.Value
                if (-not $roleUrl) { continue }

                try {
                    $params["Uri"] = $roleUrl
                    $roleDetail = Invoke-RestMethod @params
                    if ($roleDetail.actors) {
                        foreach ($actor in $roleDetail.actors) {
                            $aType = [string]$actor.type
                            if ($aType -eq "atlassian-user-role-actor" -and $actor.actorUser -and $actor.actorUser.accountId) {
                                $aid = [string]$actor.actorUser.accountId
                                if ($targetAccSet.Contains($aid)) {
                                    [void]$userRoles.Add([pscustomobject]@{ AccountId = $aid; Role = $roleName })
                                }
                            }
                            elseif ($aType -eq "atlassian-group-role-actor" -and $actor.actorGroup -and $actor.actorGroup.name) {
                                $gname = Repair-TextEncoding ([string]$actor.actorGroup.name).Trim().ToLower()
                                if ($targetGrpSet.Contains($gname)) {
                                    [void]$groupRoles.Add([pscustomobject]@{ Group = $gname; Role = $roleName })
                                }
                            }
                        }
                    }
                } catch {}
            }
        } catch {}

        return [pscustomobject]@{
            ProjectKey = $ProjectKey
            UserRoles  = $userRoles
            GroupRoles = $groupRoles
        }
    }

    $projectPool = [runspacefactory]::CreateRunspacePool(1, $ThrottleLimit)
    $projectPool.Open()

    $projTasks = New-Object System.Collections.ArrayList
    foreach ($p in $projects) {
        $pKey = [string]$p.key
        $ps = [powershell]::Create()
        $ps.RunspacePool = $projectPool
        [void]$ps.AddScript($workerProjectBlock)
        [void]$ps.AddArgument($pKey)
        [void]$ps.AddArgument($baseUrl)
        [void]$ps.AddArgument($authHeaders)
        [void]$ps.AddArgument($ProxyUrl)
        [void]$ps.AddArgument($true)
        [void]$ps.AddArgument($targetAccIdsList)
        [void]$ps.AddArgument($targetGroupsList)

        [void]$projTasks.Add([pscustomobject]@{
            ProjectKey  = $pKey
            PowerShell  = $ps
            AsyncResult = $ps.BeginInvoke()
        })
    }

    $pDone = 0
    while ($pDone -lt $projTasks.Count) {
        Start-Sleep -Milliseconds 50
        foreach ($pt in $projTasks) {
            if ($null -ne $pt.AsyncResult -and $pt.AsyncResult.IsCompleted) {
                $pResult = $pt.PowerShell.EndInvoke($pt.AsyncResult)
                $pt.PowerShell.Dispose()
                $pt.AsyncResult = $null

                if ($pResult -and $pResult[0]) {
                    $pData = $pResult[0]
                    $pk = $pData.ProjectKey

                    if ($pData.UserRoles) {
                        foreach ($ur in $pData.UserRoles) {
                            $aid = $ur.AccountId; $rName = $ur.Role
                            if (-not $projRoleByAccId.ContainsKey($aid)) { $projRoleByAccId[$aid] = @{} }
                            if (-not $projRoleByAccId[$aid].ContainsKey($pk)) { 
                                $projRoleByAccId[$aid][$pk] = New-Object System.Collections.Generic.HashSet[string] 
                            }
                            [void]$projRoleByAccId[$aid][$pk].Add($rName)
                        }
                    }
                    if ($pData.GroupRoles) {
                        foreach ($gr in $pData.GroupRoles) {
                            $grp = $gr.Group; $rName = $gr.Role
                            if (-not $projRoleByGroup.ContainsKey($grp)) { $projRoleByGroup[$grp] = @{} }
                            if (-not $projRoleByGroup[$grp].ContainsKey($pk)) { 
                                $projRoleByGroup[$grp][$pk] = New-Object System.Collections.Generic.HashSet[string] 
                            }
                            [void]$projRoleByGroup[$grp][$pk].Add($rName)
                        }
                    }
                }
                $pDone++
            }
        }
    }
    $projectPool.Close()
    $projectPool.Dispose()
}

# ======================================================================
# ÉTAPE 5/5 : ESPACES CONFLUENCE SIMPLIFIÉS (C, R, G)
# ======================================================================
Write-Progress -Activity "Audit Collaborateurs hors DSIM/DSIT" -Status "Étape 5/$totalSteps : Espaces Confluence globaux..." -PercentComplete 80
Write-Info "Analyse des permissions Confluence (espaces globaux uniquement)..."

$spacePermByAccId = @{}
$spacePermByGroup = @{}
$targetAccSet = New-Object System.Collections.Generic.HashSet[string]
foreach ($tu in $targetUsers) { [void]$targetAccSet.Add($tu.AccountId) }

$start = 0; $limit = 50
while ($true) {
    $spUrl = "$baseUrl/wiki/rest/api/space?type=global&limit=$limit&start=$start&status=current&expand=permissions"
    try {
        $spResp = Invoke-SafeGet -Url $spUrl -Headers $authHeaders
    } catch { break }

    if (-not $spResp -or -not $spResp.results -or $spResp.results.Count -eq 0) { break }

    foreach ($sp in $spResp.results) {
        $spKey = [string]$sp.key
        if ($spKey.StartsWith("~")) { continue }

        if ($sp.permissions) {
            foreach ($perm in $sp.permissions) {
                $op = [string]$perm.operation

                if ($perm.subject -and $perm.subject.type -eq "user" -and $perm.subject.accountId) {
                    $uId = [string]$perm.subject.accountId
                    if ($targetAccSet.Contains($uId)) {
                        if (-not $spacePermByAccId.ContainsKey($uId)) { $spacePermByAccId[$uId] = @{} }
                        if (-not $spacePermByAccId[$uId].ContainsKey($spKey)) { 
                            $spacePermByAccId[$uId][$spKey] = New-Object System.Collections.Generic.HashSet[string] 
                        }
                        [void]$spacePermByAccId[$uId][$spKey].Add($op)
                    }
                }

                $groupsInPerm = @()
                if ($perm.subjects -and $perm.subjects.group -and $perm.subjects.group.results) {
                    $groupsInPerm = $perm.subjects.group.results
                } elseif ($perm.subject -and $perm.subject.type -eq "group") {
                    $groupsInPerm = @($perm.subject)
                }

                foreach ($g in $groupsInPerm) {
                    $gName = Repair-DoubleUtf8 ([string]$g.name).Trim().ToLower()
                    if ($allTargetGroupNamesLower.Contains($gName)) {
                        if (-not $spacePermByGroup.ContainsKey($gName)) { $spacePermByGroup[$gName] = @{} }
                        if (-not $spacePermByGroup[$gName].ContainsKey($spKey)) { 
                            $spacePermByGroup[$gName][$spKey] = New-Object System.Collections.Generic.HashSet[string] 
                        }
                        [void]$spacePermByGroup[$gName][$spKey].Add($op)
                    }
                }
            }
        }
    }

    if ($spResp.size -lt $limit) { break }
    $start += $spResp.size
}

function Get-SimplifiedPermission([System.Collections.Generic.HashSet[string]]$ops) {
    if (-not $ops -or $ops.Count -eq 0) { return "" }
    
    foreach ($op in $ops) {
        if ($op -match "(editspace|adminspace|createpage|createattachment|comment)") {
            return "C"
        }
    }
    foreach ($op in $ops) {
        if ($op -match "(guest)") {
            return "G"
        }
    }
    return "R"
}

# ======================================================================
# COMPILATION DU RAPPORT FINAL (TRI ALPHABÉTIQUE STRICT & EXCLUSION FANTÔMES)
# ======================================================================
Write-Progress -Activity "Audit Collaborateurs hors DSIM/DSIT" -Status "Finalisation et tri alphabétique..." -PercentComplete 95
$reportRows = New-Object System.Collections.ArrayList

$sortedUsers = @($targetUsers | Sort-Object DisplayName)

foreach ($tu in $sortedUsers) {
    $aid = $tu.AccountId
    $userGroupNames = @($tu.Groups)
    $userGroupLower = @($userGroupNames | ForEach-Object { $_.Trim().ToLower() })

    # Projets Jira
    $jiraProjMap = @{}
    if ($projRoleByAccId.ContainsKey($aid)) {
        foreach ($pKey in $projRoleByAccId[$aid].Keys) {
            if (-not $jiraProjMap.ContainsKey($pKey)) { $jiraProjMap[$pKey] = New-Object System.Collections.Generic.HashSet[string] }
            foreach ($r in $projRoleByAccId[$aid][$pKey]) { [void]$jiraProjMap[$pKey].Add($r) }
        }
    }
    foreach ($g in $userGroupLower) {
        if ($projRoleByGroup.ContainsKey($g)) {
            foreach ($pKey in $projRoleByGroup[$g].Keys) {
                if (-not $jiraProjMap.ContainsKey($pKey)) { $jiraProjMap[$pKey] = New-Object System.Collections.Generic.HashSet[string] }
                foreach ($r in $projRoleByGroup[$g][$pKey]) { [void]$jiraProjMap[$pKey].Add($r) }
            }
        }
    }

    $jiraAccessList = @()
    foreach ($pKey in ($jiraProjMap.Keys | Sort-Object)) {
        $rolesStr = ($jiraProjMap[$pKey] | Sort-Object) -join ", "
        $jiraAccessList += "$pKey ($rolesStr)"
    }
    $jiraAccessStr = if ($jiraAccessList.Count -gt 0) { $jiraAccessList -join " | " } else { "(aucun rôle projet attribué)" }

    # Espaces Confluence (Format CLE [C|R|G])
    $confSpaceMap = @{}
    if ($spacePermByAccId.ContainsKey($aid)) {
        foreach ($sKey in $spacePermByAccId[$aid].Keys) {
            if (-not $confSpaceMap.ContainsKey($sKey)) { $confSpaceMap[$sKey] = New-Object System.Collections.Generic.HashSet[string] }
            foreach ($op in $spacePermByAccId[$aid][$sKey]) { [void]$confSpaceMap[$sKey].Add($op) }
        }
    }
    foreach ($g in $userGroupLower) {
        if ($spacePermByGroup.ContainsKey($g)) {
            foreach ($sKey in $spacePermByGroup[$g].Keys) {
                if (-not $confSpaceMap.ContainsKey($sKey)) { $confSpaceMap[$sKey] = New-Object System.Collections.Generic.HashSet[string] }
                foreach ($op in $spacePermByGroup[$g][$sKey]) { [void]$confSpaceMap[$sKey].Add($op) }
            }
        }
    }

    $confAccessList = @()
    foreach ($sKey in ($confSpaceMap.Keys | Sort-Object)) {
        $letter = Get-SimplifiedPermission $confSpaceMap[$sKey]
        if ($letter) {
            $confAccessList += "$sKey [$letter]"
        }
    }
    $confAccessStr = if ($confAccessList.Count -gt 0) { $confAccessList -join " | " } else { "(aucun espace)" }

    # Exclusion des comptes fantômes sans aucun groupe ni droit
    if ($userGroupNames.Count -eq 0 -and $jiraProjMap.Count -eq 0 -and $confSpaceMap.Count -eq 0) {
        continue
    }

    $groupsStr = if ($userGroupNames.Count -gt 0) { ($userGroupNames | Sort-Object) -join " | " } else { "(aucun groupe)" }

    [void]$reportRows.Add([pscustomobject]@{
        "AccountId"                   = $aid
        "DisplayName"                 = $tu.DisplayName
        "Email"                       = $tu.Email
        "Nombre de Groupes"           = $userGroupNames.Count
        "Groupes Jira"                = $groupsStr
        "Nombre Projets Jira"         = $jiraProjMap.Count
        "Projets Jira & Rôles"        = $jiraAccessStr
        "Nombre Espaces Confluence"   = $confSpaceMap.Count
        "Espaces Confluence (C/R/G)"  = $confAccessStr
    })
}

Write-Progress -Activity "Audit Collaborateurs hors DSIM/DSIT" -Completed

# ======================================================================
# EXPORT CSV (BOM UTF-8 STRICT)
# ======================================================================
$dateStamp = Get-Date -Format "yyyyMMdd_HHmmss"
$exportCsv = Join-Path $exportsDir "Audit_Users_Hors_DSIM_DSIT_${dateStamp}.csv"

$headers = @(
    "AccountId", "DisplayName", "Email", "Nombre de Groupes", "Groupes Jira",
    "Nombre Projets Jira", "Projets Jira & Rôles", "Nombre Espaces Confluence", "Espaces Confluence (C/R/G)"
)

$csvLines = New-Object System.Collections.Generic.List[string]
$csvLines.Add(($headers -join ";"))

foreach ($r in $reportRows) {
    $lineVals = @()
    foreach ($h in $headers) {
        $val = [string]$r.$h
        if ($val -match '[;"\r\n]') {
            $val = '"' + ($val -replace '"', '""') + '"'
        }
        $lineVals += $val
    }
    $csvLines.Add(($lineVals -join ";"))
}

$utf8WithBom = New-Object System.Text.UTF8Encoding($true)
[System.IO.File]::WriteAllLines($exportCsv, $csvLines, $utf8WithBom)

Write-Host ""
Write-Info "=== AUDIT TERMINÉ AVEC SUCCÈS ==="
Write-Info "Fichier exporté : $exportCsv"
Write-Host ""

$reportRows | Select-Object "DisplayName", "Email", "Nombre de Groupes", "Nombre Projets Jira", "Nombre Espaces Confluence" | Format-Table -AutoSize