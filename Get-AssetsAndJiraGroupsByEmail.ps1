<#
Get-AssetsAndJiraGroupsByEmail.ps1
Pour une liste d'utilisateurs donnee dans un CSV d'entree (username, email) :
  - Extrait l'accountId Jira (userID Atlassian) en COLONNE 1
  - Extrait les donnees RH d'Assets (Schema RP) :
      * Statut
      * Date d'entree  (format DD/MM/YYYY)
      * Direction
      * Affectation 2
      * Manager
      * Date de sortie (format DD/MM/YYYY)
      * Motif de sortie
  - Extrait les groupes d'habilitation / securite Jira (separes par " | ")

Encodage : UTF-8 avec BOM sur tous les CSV de sortie.
Correction transcoding : Windows-1252 -> UTF-8.
Mode LECTURE SEULE.
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
$scriptName = "Get-AssetsAndJiraGroupsByEmail"

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
# 2. PROXY, HELPERS ENCODAGE ET DATES
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

# Correction transcoding Windows-1252 -> UTF-8 (ex: StÃ©phane HEUZÃ‰ -> Stéphane HEUZÉ)
function Fix-Encoding {
    param([string]$val)
    if ([string]::IsNullOrWhiteSpace($val)) { return $val }
    try {
        $bytes = [System.Text.Encoding]::GetEncoding(1252).GetBytes($val)
        $fixed = [System.Text.Encoding]::UTF8.GetString($bytes)
        if ($fixed -match "&#65533;" -or $fixed -match "") { return $val }
        return $fixed
    } catch { return $val }
}

# Normalisation des dates au format DD/MM/YYYY
function ConvertTo-DateFR {
    param($val)
    if ([string]::IsNullOrWhiteSpace($val)) { return "" }
    $s = [string]$val.Trim()

    if ($s -match "^\d{13}$") {
        $dt = (Get-Date "1970-01-01").AddMilliseconds([long]$s)
        return $dt.ToString("dd/MM/yyyy")
    }
    if ($s -match "^\d{10}$") {
        $dt = (Get-Date "1970-01-01").AddSeconds([long]$s)
        return $dt.ToString("dd/MM/yyyy")
    }
    $dt = $null
    $formats = @(
        "yyyy-MM-dd",
        "yyyy-MM-ddTHH:mm:ss",
        "yyyy-MM-ddTHH:mm:ssZ",
        "dd/MM/yyyy",
        "dd-MM-yyyy",
        "MM/dd/yyyy",
        "d MMM yyyy",
        "dd MMM yyyy"
    )
    foreach ($fmt in $formats) {
        try {
            $dt = [datetime]::ParseExact($s, $fmt, [System.Globalization.CultureInfo]::InvariantCulture)
            break
        } catch {}
    }
    if (-not $dt) { try { $dt = [datetime]::Parse($s) } catch {} }
    if ($dt) { return $dt.ToString("dd/MM/yyyy") }
    return $s
}

# ============================================================
# 3. HTTP WRAPPERS (lecture UTF-8 forcee)
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
    $params = @{
        Method          = 'GET'
        Uri             = $Url
        Headers         = $Headers
        ContentType     = 'application/json'
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
Write-Info ("Connexion a Jira : " + $jiraBaseUrl + " (" + $jiraEmail + ")")

$assetsAqlUrl    = $jiraBaseUrl + "/gateway/api/jsm/assets/workspace/" + $AssetsWorkspaceId + "/v1/object/aql"
$assetsSchemaUrl = $jiraBaseUrl + "/gateway/api/jsm/assets/workspace/" + $AssetsWorkspaceId + "/v1/objectschema/list"
$jiraUserSearch  = $jiraBaseUrl + "/rest/api/3/user/search"

# ============================================================
# 5. DEMANDE DU FICHIER CSV D'ENTREE
# ============================================================
Write-Info "=== Selection du fichier CSV d'entree ==="

$selectedCsvPath = $null

if (-not [string]::IsNullOrWhiteSpace($CsvInput) -and (Test-Path $CsvInput)) {
    $selectedCsvPath = $CsvInput
}

if (-not $selectedCsvPath) {
    Write-Info "  Ouverture de la boite de dialogue pour choisir le fichier CSV..."
    Add-Type -AssemblyName System.Windows.Forms
    $dialog = New-Object System.Windows.Forms.OpenFileDialog
    $dialog.Title            = "Selectionner le fichier CSV (colonnes: username, email)"
    $dialog.Filter           = "Fichiers CSV (*.csv)|*.csv|Tous les fichiers (*.*)|*.*"
    $dialog.InitialDirectory = Join-Path ([Environment]::GetFolderPath("UserProfile")) "Downloads"
    $dialog.Multiselect      = $false

    if ($dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK -and $dialog.FileName) {
        $selectedCsvPath = $dialog.FileName
    }
}

if (-not $selectedCsvPath -or -not (Test-Path $selectedCsvPath)) {
    Write-ErrLog "Aucun fichier CSV selectionne."
    throw "Veuillez fournir un fichier CSV valide contenant les colonnes 'username' et 'email'."
}

Write-Info ("  Fichier CSV retenu : " + $selectedCsvPath)

# Lecture du CSV d'entree en UTF-8 force
$rawBytes  = [System.IO.File]::ReadAllBytes($selectedCsvPath)
$rawText   = [System.Text.Encoding]::UTF8.GetString($rawBytes).TrimStart([char]0xFEFF)
$rawLines  = $rawText -split "`r?`n"
$delimiter = if ($rawLines[0].Contains(";")) { ";" } else { "," }
$inputList = $rawText | ConvertFrom-Csv -Delimiter $delimiter

Write-Info ("  Nombre de lignes a traiter : " + $inputList.Count)

# ============================================================
# 6. CHARGEMENT MEMOIRE DU REFERENTIEL PERSONNE (ASSETS)
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
    Write-Info ("  Schema Personne trouve : " + $personSchemaName + " (id=" + $personSchemaId + ")")
}

Write-Info "=== Chargement du Referentiel Personne Assets en memoire ==="

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
            if ($ota.id -and $ota.name) { $attrDict[[string]$ota.id] = Fix-Encoding ([string]$ota.name) }
        }
    }

    foreach ($pObj in $respAql.values) {
        if (-not $pObj.id) { continue }
        $labelRP = if ($pObj.label) { Fix-Encoding ([string]$pObj.label) } else { "" }

        $props = @{}
        foreach ($attr in $pObj.attributes) {
            $aName = $attrDict[[string]$attr.objectTypeAttributeId]
            if (-not $aName -and $attr.objectTypeAttribute -and $attr.objectTypeAttribute.name) {
                $aName = Fix-Encoding ([string]$attr.objectTypeAttribute.name)
            }
            if (-not $aName) { continue }

            $vals = $attr.objectAttributeValues
            if (-not $vals -or $vals.Count -eq 0) { continue }
            $vStr = [string]$vals[0].displayValue
            if ([string]::IsNullOrWhiteSpace($vStr)) { $vStr = [string]$vals[0].value }
            $props[$aName] = Fix-Encoding $vStr
        }

        $statut       = "(Non trouve)"
        $dateEntree   = ""
        $direction    = "(Non trouve)"
        $affectation2 = ""
        $manager      = ""
        $dateSortie   = ""
        $motifSortie  = ""
        $email        = ""

        foreach ($k in $props.Keys) {
            $val = $props[$k]
            if ($k -ieq "Statut" -or $k -ilike "*Statut*") {
                $statut = Fix-Encoding $val
            } elseif ($k -ieq "DateEntree" -or $k -ilike "*Date*entree*" -or $k -ilike "*Arrivee*") {
                $dateEntree = ConvertTo-DateFR $val
            } elseif ($k -ieq "Direction" -or $k -ilike "*Direction*" -or $k -ieq "DSIM") {
                $direction = Fix-Encoding $val
            } elseif ($k -ieq "Affectation 2" -or $k -ieq "Affectation2" -or $k -ilike "*Affectation*") {
                $affectation2 = Fix-Encoding $val
            } elseif ($k -ieq "Manager" -or $k -ilike "*Manager*" -or $k -ilike "*Responsable*") {
                $manager = Fix-Encoding $val
            } elseif ($k -ieq "DateSortie" -or $k -ilike "*Date*sortie*" -or $k -ilike "*Depart*") {
                $dateSortie = ConvertTo-DateFR $val
            } elseif ($k -ieq "MotifSortie" -or $k -ilike "*Motif*sortie*" -or $k -ilike "*Motif*depart*") {
                $motifSortie = Fix-Encoding $val
            } elseif ($k -ieq "Email" -or $k -ieq "Mail" -or $k -ilike "*Email*") {
                $email = $val
            }
        }

        $personRec = [pscustomobject]@{
            Label        = $labelRP
            Statut       = $statut
            DateEntree   = $dateEntree
            Direction    = $direction
            Affectation2 = $affectation2
            Manager      = $manager
            DateSortie   = $dateSortie
            MotifSortie  = $motifSortie
            Email        = $email
        }

        if (-not [string]::IsNullOrWhiteSpace($email))   { $personByEmail[$email.Trim().ToLower()] = $personRec }
        if (-not [string]::IsNullOrWhiteSpace($labelRP)) { $personByName[$labelRP.Trim().ToLower()] = $personRec }
        $totalLoaded++
    }

    $isLast   = if ($null -ne $respAql.isLast) { [bool]$respAql.isLast } else { $true }
    $startAt += $maxResults
    Start-Sleep -Milliseconds 50
}

Write-Info ("  Personnes chargees depuis Assets : " + $totalLoaded)

# ============================================================
# 7. TRAITEMENT DES UTILISATEURS D'ENTREE
# ============================================================
Write-Info "=== Croisement des donnees (Assets + Groupes Jira) ==="

$results = New-Object System.Collections.Generic.List[object]

foreach ($row in $inputList) {
    $uName  = if ($row.username)    { Fix-Encoding ([string]$row.username) } `
          elseif ($row.'User name') { Fix-Encoding ([string]$row.'User name') } `
          else { "" }
    $uEmail = if ($row.email)        { [string]$row.email }                        elseif ($row.'User email') { [string]$row.'User email' }                else { "" }

    if ([string]::IsNullOrWhiteSpace($uEmail) -and [string]::IsNullOrWhiteSpace($uName)) { continue }

    $mailLower = if ($uEmail) { $uEmail.Trim().ToLower() } else { "" }
    $nameLower = if ($uName)  { $uName.Trim().ToLower() }  else { "" }

    # 1. Résolution Assets RH
    $matched = $null
    if ($mailLower -and $personByEmail.ContainsKey($mailLower)) {
        $matched = $personByEmail[$mailLower]
    } elseif ($nameLower -and $personByName.ContainsKey($nameLower)) {
        $matched = $personByName[$nameLower]
    }

    $statut       = if ($matched) { $matched.Statut }       else { "(Non trouve)" }
    $dateEntree   = if ($matched) { $matched.DateEntree }   else { "" }
    $direction    = if ($matched) { $matched.Direction }    else { "(Non trouve)" }
    $affectation2 = if ($matched) { $matched.Affectation2 } else { "" }
    $manager      = if ($matched) { $matched.Manager }      else { "" }
    $dateSortie   = if ($matched) { $matched.DateSortie }   else { "" }
    $motifSortie  = if ($matched) { $matched.MotifSortie }  else { "" }

    # 2. Résolution Profil & Groupes Jira + accountId (userID)
    $groupesJiraStr = "(Non trouve)"
    $accId          = $null

    $querySearch = if ($mailLower) { $mailLower } else { $nameLower }
    try {
        $uResp = Invoke-ApiGet -Url ($jiraUserSearch + "?query=" + [uri]::EscapeDataString($querySearch)) -Headers $jiraHeaders
        if ($uResp -and $uResp.Count -gt 0 -and $uResp[0].accountId) {
            $accId = [string]$uResp[0].accountId
        }
    } catch {}

    $jiraAccountId = if ($accId) { $accId } else { "(Non trouve)" }

    if ($accId) {
        try {
            $urlGroups  = $jiraBaseUrl + "/rest/api/3/user/groups?accountId=" + $accId
            $respGroups = Invoke-ApiGet -Url $urlGroups -Headers $jiraHeaders
            $gList      = if ($respGroups -is [array]) { $respGroups } elseif ($respGroups.values) { $respGroups.values } else { $null }

            if ($gList) {
                $gNames = @()
                foreach ($g in $gList) {
                    if ($g.name) { $gNames += Fix-Encoding ([string]$g.name) }
                }
                if ($gNames.Count -gt 0) {
                    $groupesJiraStr = ($gNames -join " | ")
                }
            }
        } catch {}
    }

    Write-Info ("  " + $uEmail + " (" + $uName + ") -> AccountId: " + $jiraAccountId + " | Direction: " + $direction + " | Statut: " + $statut + " | Sortie: " + $dateSortie + " | Groupes: " + $groupesJiraStr)

    $results.Add([pscustomobject]@{
        accountId    = $jiraAccountId
        username     = $uName
        email        = $uEmail
        Statut       = $statut
        DateEntree   = $dateEntree
        Direction    = $direction
        Affectation2 = $affectation2
        Manager      = $manager
        DateSortie   = $dateSortie
        MotifSortie  = $motifSortie
        GroupesJira  = $groupesJiraStr
    }) | Out-Null
}

# ============================================================
# 8. EXPORT CSV (UTF-8 avec BOM pour Excel)
# ============================================================
Write-Info "=== Export du rapport CSV ==="

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
    $utf8Bom = New-Object System.Text.UTF8Encoding($true)
    [System.IO.File]::WriteAllText($Path, $sb.ToString(), $utf8Bom)
    Write-Info ("  CSV : " + $Path + " (" + $Rows.Count + " lignes)")
}

$csvOutput = Join-Path $exportsDir ("Assets-Et-GroupesJira_" + $runStamp + ".csv")
Export-CsvStrict -Path $csvOutput `
    -Headers @("accountId","username","email","Statut","DateEntree","Direction","Affectation2","Manager","DateSortie","MotifSortie","GroupesJira") `
    -Rows $results

# ============================================================
# 9. RESUME FINAL EN CONSOLE
# ============================================================
Write-Info ""
Write-Info "=========================================="
Write-Info "=== RESUME EXTRACTION ASSETS ET JIRA ==="
Write-Info "=========================================="
Write-Info ""
Write-Info ("  Lignes traitees                : " + $results.Count)
Write-Info ("  AccountId Jira trouves         : " + ($results | Where-Object { $_.accountId -ne "(Non trouve)" }).Count)
Write-Info ("  Trouves dans Assets            : " + ($results | Where-Object { $_.Statut -ne "(Non trouve)" }).Count)
Write-Info ("  Groupes Jira trouves           : " + ($results | Where-Object { $_.GroupesJira -ne "(Non trouve)" }).Count)
Write-Info ("  Avec date de sortie renseignee : " + ($results | Where-Object { $_.DateSortie -ne "" }).Count)
Write-Info ""
Write-Info ("  Rapport CSV genere dans : " + $csvOutput)
Write-Info ("  Journal de log           : " + $logFile)
Write-Log ("=== FIN " + $scriptName + " ===") "INFO"