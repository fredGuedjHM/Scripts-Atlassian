<#
Audit-ShadowITInstances.ps1
Audit des instances Atlassian Cloud hors périmètre officiel (Shadow IT / Discovered Products).
Adapté au format d'export Atlassian Guard : product-discovery.csv

Enrichissement des créateurs / administrateurs via :
  - Parsing avancé de la colonne 'Admins' : "Nom (Rôle) <email>"
  - Le schéma Assets Référentiel Personne (Direction, Matricule)
  - L'API Tempo Teams v4 (Équipes Tempo actives)
  - Analyse du statut d'activité et d'inactivité

Mode LECTURE SEULE : aucune modification n'est effectuée.
#>

[CmdletBinding()]
param(
    [string]$CsvInput,
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

function Ensure-Dir([string]$Path) {
    if (-not (Test-Path $Path)) { New-Item -ItemType Directory -Path $Path -Force | Out-Null }
}
Ensure-Dir $secretsDir; Ensure-Dir $logsDir; Ensure-Dir $exportsDir

$runStamp   = Get-Date -Format "yyyyMMdd_HHmmss"
$scriptName = "Audit-ShadowITInstances"

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
    if ($dt) { return $dt.ToString("yyyy-MM-dd") }
    $s = [string]$val
    if (-not [string]::IsNullOrWhiteSpace($s)) { return $s }
    return "(Non disponible)"
}

# ============================================================
# 2b. PARSER SPÉCIFIQUE DE LA COLONNE ADMINS
# ============================================================
function Parse-AdminString {
    param([string]$RawAdmin)

    if ([string]::IsNullOrWhiteSpace($RawAdmin) -or $RawAdmin -eq "(aucun admin)" -or $RawAdmin -eq "No admins") {
        return [pscustomobject]@{
            Email = ""
            Name  = "(Aucun admin)"
        }
    }

    $emails = New-Object System.Collections.Generic.List[string]
    $names  = New-Object System.Collections.Generic.List[string]

    $adminEntries = $RawAdmin -split '[,;\r\n]+'

    foreach ($entry in $adminEntries) {
        $entryStr = $entry.Trim()
        if ([string]::IsNullOrWhiteSpace($entryStr)) { continue }

        $email = ""
        if ($entryStr -match '<([^>]+)>') {
            $email = $Matches[1].Trim()
        } elseif ($entryStr -match '[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}') {
            $email = $Matches[0].Trim()
        }

        $cleanName = $entryStr
        $cleanName = $cleanName -replace '\([^)]*\)', '' # Supprime le rôle (Organisation Admin)
        $cleanName = $cleanName -replace '<[^>]*>', ''    # Supprime l'email <...>
        $cleanName = $cleanName.Trim()

        if (-not $cleanName -and $email) {
            $cleanName = $email.Split('@')[0]
        }

        if ($email -and -not $emails.Contains($email)) { $emails.Add($email) | Out-Null }
        if ($cleanName -and -not $names.Contains($cleanName)) { $names.Add($cleanName) | Out-Null }
    }

    $finalEmail = if ($emails.Count -gt 0) { ($emails -join " | ") } else { "" }
    $finalName  = if ($names.Count -gt 0) { ($names -join " | ") } else { if ($finalEmail) { $finalEmail } else { "(Aucun admin)" } }

    return [pscustomobject]@{
        Email = $finalEmail
        Name  = $finalName
    }
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

$assetsAqlUrl    = $jiraBaseUrl + "/gateway/api/jsm/assets/workspace/" + $AssetsWorkspaceId + "/v1/object/aql"
$assetsSchemaUrl = $jiraBaseUrl + "/gateway/api/jsm/assets/workspace/" + $AssetsWorkspaceId + "/v1/objectschema/list"
$jiraUserSearch  = $jiraBaseUrl + "/rest/api/3/user/search"

# ============================================================
# 5. CHARGEMENT DU FICHIER EXPORT SHADOW IT (PRODUCT-DISCOVERY.CSV)
# ============================================================
Write-Info "=== Recherche du fichier d'export Shadow IT ==="

$selectedCsvPath = $null

if (-not [string]::IsNullOrWhiteSpace($CsvInput) -and (Test-Path $CsvInput)) {
    $selectedCsvPath = $CsvInput
}

if (-not $selectedCsvPath) {
    $userDownloads = Join-Path ([Environment]::GetFolderPath("UserProfile")) "Downloads"
    $candidateFiles = @(
        (Join-Path $scriptDir "product-discovery.csv"),
        (Join-Path $scriptDir "product-discovery.csv.csv"),
        (Join-Path $scriptDir "shadow_it_apps.csv"),
        (Join-Path $userDownloads "product-discovery.csv"),
        (Join-Path $userDownloads "product-discovery.csv.csv"),
        (Join-Path $userDownloads "shadow_it_apps.csv")
    )
    foreach ($cFile in $candidateFiles) {
        if (Test-Path $cFile) { $selectedCsvPath = $cFile; break }
    }
}

if (-not $selectedCsvPath) {
    Write-Info "  Ouverture de la boite de dialogue pour selectionner l'export Shadow IT..."
    Add-Type -AssemblyName System.Windows.Forms
    $dialog = New-Object System.Windows.Forms.OpenFileDialog
    $dialog.Title            = "Selectionner l'export CSV Shadow IT (product-discovery.csv)"
    $dialog.Filter           = "Fichiers CSV (*.csv)|*.csv|Tous les fichiers (*.*)|*.*"
    $dialog.InitialDirectory = Join-Path ([Environment]::GetFolderPath("UserProfile")) "Downloads"
    $dialog.Multiselect      = $false

    if ($dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK -and $dialog.FileName) {
        $selectedCsvPath = $dialog.FileName
    }
}

if (-not $selectedCsvPath -or -not (Test-Path $selectedCsvPath)) {
    Write-ErrLog "Aucun fichier d'export Shadow IT trouve."
    throw "Exporte le CSV product-discovery.csv depuis admin.atlassian.com > Sécurité > Produits découverts."
}

Write-Info ("  Fichier d'export retenu : " + $selectedCsvPath)

# Import du CSV avec auto-détection du séparateur (virgule ou point-virgule)
$rawLines = Get-Content -Path $selectedCsvPath -Encoding UTF8
$delimiter = if ($rawLines[0].Contains(";")) { ";" } else { "," }
$csvData = Import-Csv -Path $selectedCsvPath -Delimiter $delimiter

Write-Info ("  Lignes d'instances Shadow IT trouvees : " + $csvData.Count)

# ============================================================
# 6. CHARGEMENT ASSETS & TEMPO EN MEMOIRE
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
    Write-Info "  Chargement des equipes via API Tempo v4..."
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
# 7. TRAITEMENT DES INSTANCES SHADOW IT
# ============================================================
Write-Info "=== Traitement et enrichissement des instances ==="

$detailedRows = New-Object System.Collections.Generic.List[object]
$nowDate      = Get-Date

foreach ($row in $csvData) {
    # Match des colonnes product-discovery.csv
    $appName        = if ($row.'Product') { $row.'Product' } elseif ($row.'App') { $row.'App' } elseif ($row.'Name') { $row.'Name' } else { "(Produit sans nom)" }
    $appUrl         = if ($row.'URL') { $row.'URL' } elseif ($row.'App URL') { $row.'App URL' } else { "" }
    $userCount      = if ($row.'User Count') { [int]$row.'User Count' } elseif ($row.'User count') { [int]$row.'User count' } else { 0 }
    $statusInst     = if ($row.'Status') { [string]$row.'Status' } else { "Inconnu" }
    $statusDetails  = if ($row.'Status details') { [string]$row.'Status details' } else { "" }

    $rawCreated     = if ($row.'Created on') { $row.'Created on' } elseif ($row.'Created') { $row.'Created' } else { $null }
    $rawLastAct     = if ($row.'Last Active') { $row.'Last Active' } elseif ($row.'Last active') { $row.'Last active' } else { $null }

    $createdStr     = Format-DateString $rawCreated
    $lastActStr     = Format-DateString $rawLastAct

    # Calcul des jours d'inactivité
    $lastActDt       = Parse-DateObj $rawLastAct
    $joursInactivite = 0
    $statutActivite  = "Actif récent"

    if ($lastActDt) {
        $joursInactivite = [math]::Max(0, [int]($nowDate - $lastActDt).TotalDays)
        if     ($joursInactivite -gt 180) { $statutActivite = "Inactif > 6 mois" }
        elseif ($joursInactivite -gt 90)  { $statutActivite = "Inactif > 3 mois" }
        elseif ($joursInactivite -gt 30)  { $statutActivite = "Inactif > 1 mois" }
    } else {
        $statutActivite = if ($statusInst -eq "Suspended") { "Suspendu (Inactif)" } else { "Activité inconnue" }
    }

    # Parsing de la colonne Admins
    $rawAdmins    = if ($row.'Admins') { [string]$row.'Admins' } elseif ($row.'Admin email') { [string]$row.'Admin email' } else { "" }
    $parsedAdmins = Parse-AdminString -RawAdmin $rawAdmins

    $creatorEmail = $parsedAdmins.Email
    $creatorName  = $parsedAdmins.Name

    # Résolution RH via Assets
    $direction = "(Non trouve)"
    $matricule = ""
    $labelRP   = ""

    if ($creatorEmail) {
        $firstMail = ($creatorEmail -split ' \| ')[0].Trim().ToLower()
        if ($personByEmail.ContainsKey($firstMail)) {
            $matched   = $personByEmail[$firstMail]
            $direction = $matched.Direction
            $matricule = $matched.Matricule
            $labelRP   = $matched.Label
            if ($creatorName -eq $creatorEmail -or $creatorName -eq $firstMail) { $creatorName = $labelRP }
        }
    }

    # Détection intelligente des entités partenaires si absent d'Assets HM
    if ($direction -eq "(Non trouve)" -or [string]::IsNullOrWhiteSpace($direction)) {
        if (-not $creatorEmail) {
            $direction = "(Sans Admin Déclaré)"
        } elseif ($creatorEmail -ilike "*@mnt.fr*") {
            $direction = "MNT (Entité Partenaire)"
        } elseif ($creatorEmail -ilike "*@mgen.fr*") {
            $direction = "MGEN (Entité Partenaire)"
        } elseif ($creatorEmail -ilike "*@harmonie-mutuelle.fr*") {
            $direction = "Harmonie Mutuelle (Hors RP Assets)"
        }
    }

    # Équipes Tempo
    $eqTempo = "(Non trouve)"
    if ($creatorEmail) {
        $firstMail = ($creatorEmail -split ' \| ')[0].Trim()
        try {
            $uResp = Invoke-ApiGet -Url ($jiraUserSearch + "?query=" + [uri]::EscapeDataString($firstMail)) -Headers $jiraHeaders
            if ($uResp -and $uResp.Count -gt 0 -and $uResp[0].accountId) {
                $accId = [string]$uResp[0].accountId
                if ($tempoUserTeamsMap.ContainsKey($accId)) {
                    $eqTempo = ($tempoUserTeamsMap[$accId] -join " | ")
                }
            }
        } catch {}
    }

    Write-Info ("  Site: " + $appUrl + " (" + $appName + ") | Admin: " + $creatorName + " | Direction: " + $direction + " | Statut: " + $statusInst)

    $detailedRows.Add([pscustomobject]@{
        DirectionAssets  = $direction
        CreatorName      = $creatorName
        CreatorEmail     = $creatorEmail
        CreatorMatricule = $matricule
        EquipesTempo     = $eqTempo
        AppName          = $appName
        AppUrl           = $appUrl
        UserCount        = $userCount
        StatusInst       = $statusInst
        StatusDetails    = $statusDetails
        CreatedOn        = $createdStr
        LastActive       = $lastActStr
        JoursInactivite  = $joursInactivite
        StatutActivite   = $statutActivite
    }) | Out-Null
}

# ============================================================
# 8. EXPORTS CSV CONSOLIDES
# ============================================================
Write-Info "=== Export des rapports CSV Shadow IT ==="

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

# 8a. Liste Détaillée des Instances Shadow IT
$csvDetail = Join-Path $exportsDir ("ShadowIT-ListeDetaillee_" + $runStamp + ".csv")
Export-CsvStrict -Path $csvDetail `
    -Headers @("DirectionAssets","CreatorName","CreatorEmail","CreatorMatricule","EquipesTempo",
               "AppName","AppUrl","UserCount","StatusInst","StatusDetails","CreatedOn","LastActive","JoursInactivite","StatutActivite") `
    -Rows ($detailedRows | Sort-Object DirectionAssets, CreatorName)

# 8b. Synthèse par Direction / Entité
$grpDirection = $detailedRows | Group-Object DirectionAssets | Sort-Object Name
$synthDirRows = foreach ($g in $grpDirection) {
    $createursDir = @($g.Group | ForEach-Object { $_.CreatorEmail } | Where-Object { $_ } | Select-Object -Unique).Count
    $totalUsers   = ($g.Group | Measure-Object -Property UserCount -Sum).Sum
    [pscustomobject]@{
        DirectionAssets      = $g.Name
        NbInstancesShadowIT  = $g.Count
        NbCreateursDistincts = $createursDir
        TotalUtilisateurs    = $totalUsers
    }
}
$csvSynthDir = Join-Path $exportsDir ("ShadowIT-SyntheseParDirection_" + $runStamp + ".csv")
Export-CsvStrict -Path $csvSynthDir `
    -Headers @("DirectionAssets","NbInstancesShadowIT","NbCreateursDistincts","TotalUtilisateurs") `
    -Rows $synthDirRows

# 8c. Synthèse par Créateur / Administrateur
$grpCreator = $detailedRows | Group-Object CreatorEmail | Sort-Object Count -Descending
$synthCreatorRows = foreach ($g in $grpCreator) {
    $c0 = $g.Group[0]
    [pscustomobject]@{
        DirectionAssets     = $c0.DirectionAssets
        CreatorName         = $c0.CreatorName
        CreatorEmail        = $c0.CreatorEmail
        EquipesTempo        = $c0.EquipesTempo
        NbInstancesCrees    = $g.Count
    }
}
$csvSynthCreator = Join-Path $exportsDir ("ShadowIT-SyntheseParCreateur_" + $runStamp + ".csv")
Export-CsvStrict -Path $csvSynthCreator `
    -Headers @("DirectionAssets","CreatorName","CreatorEmail","EquipesTempo","NbInstancesCrees") `
    -Rows ($synthCreatorRows | Sort-Object DirectionAssets, CreatorName)

# ============================================================
# 9. RESUME FINAL EN CONSOLE
# ============================================================
Write-Info ""
Write-Info "============================================="
Write-Info "=== RESUME AUDIT SHADOW IT (ATLASSIAN) ==="
Write-Info "============================================="
Write-Info ""
Write-Info ("  Total instances Shadow IT identifiees : " + $detailedRows.Count)
Write-Info ("  Auteurs / Admins distincts           : " + $grpCreator.Count)
Write-Info ("  Directions / Entites impactees       : " + $grpDirection.Count)
Write-Info ""
Write-Info "  Répartition du Statut d'Activité :"
$grpStatut = $detailedRows | Group-Object StatutActivite
foreach ($st in $grpStatut) {
    Write-Info ("    - " + $st.Name + " : " + $st.Count + " instance(s)")
}
Write-Info ""
Write-Info "  Top Directions / Entités impactées :"
foreach ($d in ($synthDirRows | Sort-Object NbInstancesShadowIT -Descending | Select-Object -First 5)) {
    Write-Info ("    - " + $d.DirectionAssets + " : " + $d.NbInstancesShadowIT + " instance(s), " + $d.TotalUtilisateurs + " utilisateur(s)")
}
Write-Info ""
Write-Info ("  Rapports CSV générés dans : " + $exportsDir)
Write-Info ("  Journal de log             : " + $logFile)
Write-Log ("=== FIN " + $scriptName + " ===") "INFO"