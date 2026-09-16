<#
Get-ShadowITUsersHM.ps1
Extraction de la liste nominative complete des utilisateurs inscrits sur les instances Shadow IT Harmonie Mutuelle.

Instances cibles :
  1. https://harmonie-mutuelle.atlassian.net
  2. https://harmonie-mutuelle-filiere-entreprises.atlassian.net

Enrichissement automatique pour chaque utilisateur :
  - Identite Jira (Nom, Email, AccountId, Statut)
  - Schema Assets Referentiel Personne (Direction / DSIM, Matricule RH)
  - API Tempo Teams v4 (Equipes Tempo actives)

Mode LECTURE SEULE : aucune modification n'est effectuee.
#>

[CmdletBinding()]
param(
    [switch]$UseSystemProxy = $true,
    [switch]$ProxyUseDefaultCredentials = $true,
    [string]$ProxyUrl,
    [string]$AssetsWorkspaceId      = "2898703d-5f14-4527-9f8a-ad32b9279ca1",
    [string]$CloudId                = "518657a0-a98f-4c2d-bddb-c6f6039addaa",
    [string]$AssetsPersonSchemaName = "RP"
)

# ============================================================
# 0. DOSSIERS ET INITIALISATION
# ============================================================
$scriptDir  = if ($PSCommandPath) { Split-Path -Parent $PSCommandPath } else { (Get-Location).Path }
$secretsDir = Join-Path $scriptDir "secrets"
$logsDir    = Join-Path $scriptDir "logs"
$exportsDir = Join-Path $scriptDir "exports"

function Initialize-Directory([string]$Path) {
    if (-not (Test-Path $Path)) { New-Item -ItemType Directory -Path $Path -Force | Out-Null }
}
Initialize-Directory $secretsDir; Initialize-Directory $logsDir; Initialize-Directory $exportsDir

$runStamp   = Get-Date -Format "yyyyMMdd_HHmmss"
$scriptName = "Get-ShadowITUsersHM"

# ============================================================
# 1. SYSTEME DE LOG
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
Write-Info ("Cloud ID         : " + $CloudId)
Write-Info ("Workspace Assets : " + $AssetsWorkspaceId)

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
Write-Info ("Compte de connexion : " + $jiraEmail)

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

$assetsAqlUrl    = $jiraBaseUrl + "/gateway/api/jsm/assets/workspace/" + $AssetsWorkspaceId + "/v1/object/aql"
$assetsSchemaUrl = $jiraBaseUrl + "/gateway/api/jsm/assets/workspace/" + $AssetsWorkspaceId + "/v1/objectschema/list"

# ============================================================
# 5. CHARGEMENT MEMOIRE ASSETS (SCHEMA RP) & TEMPO TEAMS
# ============================================================
Write-Info "=== Chargement du Referentiel Personne Assets en memoire ==="

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

$personByEmail  = @{}
$personByName   = @{}
$aqlPersonQuery = if ($personSchemaId) { "objectSchemaId = " + $personSchemaId } else { 'objectSchema = "' + $personSchemaName + '"' }

$startAt     = 0
$maxResults  = 200
$isLast      = $false
$totalLoaded = 0

while (-not $isLast) {
    $urlAql  = $assetsAqlUrl + "?startAt=" + $startAt + "&maxResults=" + $maxResults + "&includeAttributes=true"
    $bodyAql = (@{ qlQuery = $aqlPersonQuery } | ConvertTo-Json -Depth 3)
    $respAql = $null

    try { $respAql = Invoke-ApiPostUtf8 -Url $urlAql -Headers $jiraHeaders -JsonBody $bodyAql } catch { break }

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

# Charger Tempo Teams si token disponible
$tempoUserTeamsMap = @{}
if (-not [string]::IsNullOrWhiteSpace($tempoToken)) {
    Write-Info "=== Chargement des Equipes Tempo (API Tempo v4) ==="
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

                        if (-not $tempoUserTeamsMap.ContainsKey($mAccId)) {
                            $tempoUserTeamsMap[$mAccId] = New-Object System.Collections.Generic.List[string]
                        }
                        if (-not $tempoUserTeamsMap[$mAccId].Contains($tName)) {
                            $tempoUserTeamsMap[$mAccId].Add($tName) | Out-Null
                        }
                    }
                } catch {}
            }
            $teamsUrl = if ($respTeams.metadata -and $respTeams.metadata.next) { $respTeams.metadata.next } else { $null }
        }
    } catch {}
}

# ============================================================
# 6. INTERROGATION DES 2 INSTANCES SHADOW IT HARMONIE MUTUELLE
# ============================================================
Write-Info "=== Interrogation des instances Shadow IT Harmonie Mutuelle ==="

$targetInstances = @(
    @{ Name = "harmonie-mutuelle";                    Url = "https://harmonie-mutuelle.atlassian.net" },
    @{ Name = "harmonie-mutuelle-filiere-entreprises"; Url = "https://harmonie-mutuelle-filiere-entreprises.atlassian.net" }
)

$detailedUsers = New-Object System.Collections.Generic.List[object]

foreach ($inst in $targetInstances) {
    $instName = $inst.Name
    $instUrl  = $inst.Url
    Write-Info ("`n--- Extraction des utilisateurs pour : " + $instUrl + " ---")

    $instUserMap = @{} # accountId -> userObj

    # Strategie 1 : API Confluence CQL User Search
    $cqlUrl = $instUrl + "/wiki/rest/api/user/search?cql=type=user&limit=200"
    try {
        $cqlResp = Invoke-ApiGet -Url $cqlUrl -Headers $jiraHeaders
        $uList   = if ($cqlResp.results) { $cqlResp.results } else { @() }

        foreach ($u in $uList) {
            $accId = if ($u.accountId) { [string]$u.accountId } else { "" }
            if (-not $accId -or $instUserMap.ContainsKey($accId)) { continue }

            $instUserMap[$accId] = [pscustomobject]@{
                AccountId   = $accId
                DisplayName = if ($u.displayName) { [string]$u.displayName } else { "Inconnu" }
                Email       = if ($u.email) { [string]$u.email } else { "" }
                Type        = if ($u.type) { [string]$u.type } else { "user" }
                Active      = $true
            }
        }
        Write-Info ("  API CQL Confluence : " + $instUserMap.Count + " utilisateur(s) trouve(s)")
    } catch {
        Write-Warn ("  Echec API CQL Confluence sur " + $instUrl + " : " + $_.Exception.Message)
    }

    # Strategie 2 : API User Search Jira / REST v3 (si actif)
    $jiraUserUrl = $instUrl + "/rest/api/3/users/search?startAt=0&maxResults=200"
    try {
        $jResp = Invoke-ApiGet -Url $jiraUserUrl -Headers $jiraHeaders
        if ($jResp -is [array]) {
            foreach ($u in $jResp) {
                $accId = if ($u.accountId) { [string]$u.accountId } else { "" }
                if (-not $accId -or $instUserMap.ContainsKey($accId)) { continue }

                $instUserMap[$accId] = [pscustomobject]@{
                    AccountId   = $accId
                    DisplayName = if ($u.displayName) { [string]$u.displayName } else { "Inconnu" }
                    Email       = if ($u.emailAddress) { [string]$u.emailAddress } else { "" }
                    Type        = if ($u.accountType) { [string]$u.accountType } else { "atlassian" }
                    Active      = if ($null -ne $u.active) { [bool]$u.active } else { $true }
                }
            }
        }
        Write-Info ("  Total apres API Jira Users : " + $instUserMap.Count + " utilisateur(s)")
    } catch {}

    # Strategie 3 : Membres du groupe confluence-users
    $groupMembersUrl = $instUrl + "/wiki/rest/api/group/confluence-users/member?limit=200"
    try {
        $gResp  = Invoke-ApiGet -Url $groupMembersUrl -Headers $jiraHeaders
        $gUsers = if ($gResp.results) { $gResp.results } else { @() }

        foreach ($u in $gUsers) {
            $accId = if ($u.accountId) { [string]$u.accountId } else { "" }
            if (-not $accId -or $instUserMap.ContainsKey($accId)) { continue }

            $instUserMap[$accId] = [pscustomobject]@{
                AccountId   = $accId
                DisplayName = if ($u.displayName) { [string]$u.displayName } else { "Inconnu" }
                Email       = if ($u.email) { [string]$u.email } else { "" }
                Type        = if ($u.type) { [string]$u.type } else { "user" }
                Active      = $true
            }
        }
        Write-Info ("  Total consolide pour " + $instName + " : " + $instUserMap.Count + " utilisateur(s)")
    } catch {}

    if ($instUserMap.Count -eq 0) {
        Write-Warn ("  Aucun utilisateur recupere sur " + $instUrl + ". Avez-vous bien clique sur 'Rejoindre en tant qu'administrateur' dans admin.atlassian.com ?")
    }

    # Enrichissement avec Assets RH et Tempo Teams
    foreach ($uKey in $instUserMap.Keys) {
        $u = $instUserMap[$uKey]

        # Ignorer les bots / app systeme
        if ($u.Type -eq "app" -or $u.DisplayName -ilike "*Automation*" -or $u.DisplayName -ilike "*System*") { continue }

        $direction = "(Non trouve)"
        $matricule = ""
        $labelRP   = ""

        $mailLower = if ($u.Email) { $u.Email.Trim().ToLower() } else { "" }
        $nameLower = if ($u.DisplayName) { $u.DisplayName.Trim().ToLower() } else { "" }

        $matched = $null
        if ($mailLower -and $personByEmail.ContainsKey($mailLower)) {
            $matched = $personByEmail[$mailLower]
        } elseif ($nameLower -and $personByName.ContainsKey($nameLower)) {
            $matched = $personByName[$nameLower]
        }

        if ($matched) {
            $direction = $matched.Direction
            $matricule = $matched.Matricule
            $labelRP   = $matched.Label
            if (-not $u.Email -and $matched.Email) { $u.Email = $matched.Email }
        } elseif ($mailLower -ilike "*@harmonie-mutuelle.fr*") {
            $direction = "Harmonie Mutuelle (Hors RP Assets)"
        } elseif ($mailLower -ilike "*@prestataire.sihm.fr*") {
            $direction = "Prestataire SIHM (Hors RP Assets)"
        }

        # Equipes Tempo
        $eqTempo = "(Non trouve)"
        if ($tempoUserTeamsMap.ContainsKey($u.AccountId)) {
            $eqTempo = ($tempoUserTeamsMap[$u.AccountId] -join " | ")
        }

        $detailedUsers.Add([pscustomobject]@{
            InstanceName    = $instName
            InstanceUrl     = $instUrl
            DisplayName     = $u.DisplayName
            Email           = $u.Email
            DirectionAssets = $direction
            EquipesTempo    = $eqTempo
            Matricule       = $matricule
            AccountId       = $u.AccountId
            AccountType     = $u.Type
            StatutCompte    = if ($u.Active) { "Actif" } else { "Inactif" }
        }) | Out-Null
    }
}

# ============================================================
# 7. EXPORTS CSV CONSOLIDES
# ============================================================
Write-Info "`n=== Export des rapports CSV des utilisateurs Shadow IT ==="

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

# 7a. Liste DetaillÃ©e des Utilisateurs
$csvDetail = Join-Path $exportsDir ("ShadowIT-UtilisateursHM-ListeDetaillee_" + $runStamp + ".csv")
Export-CsvStrict -Path $csvDetail `
    -Headers @("InstanceName","DirectionAssets","DisplayName","Email","EquipesTempo","Matricule","StatutCompte","AccountType","AccountId","InstanceUrl") `
    -Rows ($detailedUsers | Sort-Object InstanceName, DirectionAssets, DisplayName)

# 7b. Synthese par Direction
$grpDirection = $detailedUsers | Group-Object DirectionAssets | Sort-Object Name
$synthDirRows = foreach ($g in $grpDirection) {
    $usersDir  = @($g.Group | ForEach-Object { $_.Email } | Where-Object { $_ } | Select-Object -Unique).Count
    $instCount = @($g.Group | ForEach-Object { $_.InstanceName } | Select-Object -Unique).Count
    [pscustomobject]@{
        DirectionAssets      = $g.Name
        NbUtilisateursHM     = if ($usersDir -gt 0) { $usersDir } else { $g.Count }
        NbInstancesImpactees = $instCount
    }
}
$csvSynthDir = Join-Path $exportsDir ("ShadowIT-UtilisateursHM-SyntheseParDirection_" + $runStamp + ".csv")
Export-CsvStrict -Path $csvSynthDir `
    -Headers @("DirectionAssets","NbUtilisateursHM","NbInstancesImpactees") `
    -Rows $synthDirRows

# 7c. Synthese par Instance
$grpInstance = $detailedUsers | Group-Object InstanceName
$synthInstRows = foreach ($g in $grpInstance) {
    $actifs = ($g.Group | Where-Object { $_.StatutCompte -eq "Actif" }).Count
    [pscustomobject]@{
        InstanceName            = $g.Name
        TotalUsersIdentifies    = $g.Count
        UsersActifs             = $actifs
    }
}
$csvSynthInst = Join-Path $exportsDir ("ShadowIT-UtilisateursHM-SyntheseParInstance_" + $runStamp + ".csv")
Export-CsvStrict -Path $csvSynthInst `
    -Headers @("InstanceName","TotalUsersIdentifies","UsersActifs") `
    -Rows $synthInstRows

# ============================================================
# 8. RESUME FINAL EN CONSOLE
# ============================================================
Write-Info ""
Write-Info "=========================================================="
Write-Info "=== RESUME UTILISATEURS SHADOW IT HARMONIE MUTUELLE ==="
Write-Info "=========================================================="
Write-Info ""
Write-Info ("  Total utilisateurs identifies  : " + $detailedUsers.Count)
Write-Info ("  Directions RH impactees        : " + $grpDirection.Count)
Write-Info ""
Write-Info "  Detail par instance :"
foreach ($i in $synthInstRows) {
    Write-Info ("    - " + $i.InstanceName + " : " + $i.TotalUsersIdentifies + " utilisateur(s) dont " + $i.UsersActifs + " actif(s)")
}
Write-Info ""
Write-Info "  Top Directions impactees :"
foreach ($d in ($synthDirRows | Sort-Object NbUtilisateursHM -Descending | Select-Object -First 5)) {
    Write-Info ("    - " + $d.DirectionAssets + " : " + $d.NbUtilisateursHM + " utilisateur(s)")
}
Write-Info ""
Write-Info ("  Rapports CSV generes dans : " + $exportsDir)
Write-Info ("  Journal de log             : " + $logFile)
Write-Log ("=== FIN " + $scriptName + " ===") "INFO"