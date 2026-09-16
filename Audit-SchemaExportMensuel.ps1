<#
Audit-SchemaExportMensuel.ps1
Audit mensuel du contenu du schema Assets "Export".
Regroupe les objets par mois (YYYY-MM) selon leur date de creation, mise a jour ou attribut date.
Mode LECTURE SEULE : aucune modification n'est effectuee.
#>

[CmdletBinding()]
param(
    [switch]$UseSystemProxy = $true,
    [switch]$ProxyUseDefaultCredentials = $true,
    [string]$ProxyUrl,
    [string]$AssetsWorkspaceId = "2898703d-5f14-4527-9f8a-ad32b9279ca1",
    [string]$SchemaName        = "Export",
    [string]$SchemaId,                          # Optionnel : ID du schema si connu
    [string]$DateRef           = "Created",     # "Created", "Updated", ou nom d'un attribut (ex: "Date Export")
    [int]$DerniersMois         = 24             # Nombre de mois d'historique a analyser
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
$scriptName = "Audit-SchemaExportMensuel"

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
Write-Info ("Schema vise          : " + $SchemaName)
Write-Info ("Date de reference    : " + $DateRef)

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

$assetsAqlUrl    = $jiraBaseUrl + "/gateway/api/jsm/assets/workspace/" + $AssetsWorkspaceId + "/v1/object/aql"
$assetsSchemaUrl = $jiraBaseUrl + "/gateway/api/jsm/assets/workspace/" + $AssetsWorkspaceId + "/v1/objectschema/list"

# ============================================================
# 5. PARSEUR DE DATES ROBUSTE
# ============================================================
function Parse-AssetDate([string]$DateStr) {
    if ([string]::IsNullOrWhiteSpace($DateStr)) { return $null }
    $s  = $DateStr.Trim()
    $ci = [System.Globalization.CultureInfo]::InvariantCulture

    # ISO 8601 natif Assets (Created/Updated) : 2025-09-09T16:01:43.241Z
    try { return [datetime]::Parse($s, $ci, [System.Globalization.DateTimeStyles]::RoundtripKind) } catch {}

    $s9  = if ($s.Length -ge 9)  { $s.Substring(0, 9).Trim()  } else { $s }
    $s10 = if ($s.Length -ge 10) { $s.Substring(0, 10).Trim() } else { $s }

    # Format Assets natif attribut : 31/Dec/24
    try { return [datetime]::ParseExact($s9,  "dd/MMM/yy",   $ci) } catch {}
    # Format Assets long  : 31/Dec/2024
    try { return [datetime]::ParseExact($s9,  "dd/MMM/yyyy", $ci) } catch {}
    # ISO date seule : yyyy-MM-dd
    try { return [datetime]::ParseExact($s10, "yyyy-MM-dd",  $ci) } catch {}
    # FR : dd/MM/yyyy
    try { return [datetime]::ParseExact($s10, "dd/MM/yyyy",  $ci) } catch {}
    # Timestamp epoch ms
    if ($s -match "^\d{13}$") {
        try { return (Get-Date "1970-01-01").AddMilliseconds([long]$s) } catch {}
    }

    return $null
}

# ============================================================
# 6. RESOLUTION DU SCHEMA ASSETS
# ============================================================
Write-Info "=== Resolution du Schema Assets ==="

if ([string]::IsNullOrWhiteSpace($SchemaId)) {
    try {
        $schemasResp = Invoke-ApiGet -Url $assetsSchemaUrl -Headers $jiraHeaders
        $schemas = if ($schemasResp.values) { $schemasResp.values } else { $schemasResp.objectSchemas }

        foreach ($sch in $schemas) {
            if ($sch.name -and ($sch.name.Trim() -ieq $SchemaName.Trim() -or $sch.name -ilike ("*" + $SchemaName + "*"))) {
                $SchemaId   = [string]$sch.id
                $SchemaName = [string]$sch.name
                break
            }
        }
    } catch {
        Write-Warn ("Impossible de lister les schemas via API : " + $_.Exception.Message)
    }
}

$SchemaIdQuery = $null

if ([string]::IsNullOrWhiteSpace($SchemaId)) {
    Write-Info "Recherche par AQL direct du schema..."
    $aqlTest  = 'objectSchema = "' + $SchemaName + '"'
    $urlTest  = $assetsAqlUrl + "?startAt=0" + "&maxResults=1" + "&includeAttributes=false"
    $bodyTest = (@{ qlQuery = $aqlTest } | ConvertTo-Json)
    try {
        $respTest = Invoke-ApiPostUtf8 -Url $urlTest -Headers $jiraHeaders -JsonBody $bodyTest
        if ($respTest -and $respTest.values -and $respTest.values.Count -gt 0) {
            $SchemaIdQuery = $aqlTest
        }
    } catch {
        Write-Warn ("AQL direct echoue : " + $_.Exception.Message)
    }
}

if ([string]::IsNullOrWhiteSpace($SchemaId) -and [string]::IsNullOrWhiteSpace($SchemaIdQuery)) {
    Write-ErrLog ("Schema '" + $SchemaName + "' introuvable dans le workspace " + $AssetsWorkspaceId)
    throw ("Schema introuvable : " + $SchemaName)
}

$aqlQuery = if (-not [string]::IsNullOrWhiteSpace($SchemaId)) {
    "objectSchemaId = " + $SchemaId
} else {
    $SchemaIdQuery
}

Write-Info ("Schema identifie : " + $SchemaName + " (AQL: " + $aqlQuery + ")")

# ============================================================
# 7. RECUPERATION DES OBJETS DU SCHEMA EXPORT (OPTIMISE)
# ============================================================
Write-Info "=== Recuperation des objets Assets ==="

# Choix du mode : avec ou sans attributs
# includeAttributes=false si DateRef = Created ou Updated (champs natifs)
# includeAttributes=true  si DateRef = nom d'un attribut custom
$needAttributes = ($DateRef -ine "Created" -and $DateRef -ine "Updated")
$includeAttr    = if ($needAttributes) { "true" } else { "false" }
Write-Info ("  includeAttributes : " + $includeAttr + " (DateRef=" + $DateRef + ")")

$allObjects = New-Object System.Collections.Generic.List[object]
$startAt    = 0
$maxResults = 200        # max accepte par l'API Assets
$isLast     = $false
$totalCount = 0

while (-not $isLast) {
    $url      = $assetsAqlUrl + "?startAt=" + $startAt `
                              + "&maxResults=" + $maxResults `
                              + "&includeAttributes=" + $includeAttr
    $bodyJson = (@{ qlQuery = $aqlQuery } | ConvertTo-Json -Depth 3)

    try {
        $resp = Invoke-ApiPostUtf8 -Url $url -Headers $jiraHeaders -JsonBody $bodyJson
    } catch {
        Write-ErrLog ("Erreur AQL startAt=" + $startAt + " : " + $_.Exception.Message)
        break
    }
    if (-not $resp) { break }

    # Dictionnaire attributs (utile uniquement si includeAttributes=true)
    $attrDict = @{}
    if ($needAttributes -and $resp.objectTypeAttributes) {
        foreach ($ota in $resp.objectTypeAttributes) {
            if ($ota.id -and $ota.name) { $attrDict[[string]$ota.id] = [string]$ota.name }
        }
    }

    $objects = $resp.values
    if (-not $objects -or $objects.Count -eq 0) { break }

    foreach ($obj in $objects) {
        if (-not $obj.id) { continue }

        # Extraction des attributs custom (si necessaire)
        $props = @{}
        if ($needAttributes) {
            foreach ($attr in $obj.attributes) {
                $attrName = $null
                if ($attr.objectTypeAttributeId) {
                    $attrName = $attrDict[[string]$attr.objectTypeAttributeId]
                }
                if (-not $attrName -and $attr.objectTypeAttribute -and $attr.objectTypeAttribute.name) {
                    $attrName = [string]$attr.objectTypeAttribute.name
                }
                if (-not $attrName) { continue }
                $vals = $attr.objectAttributeValues
                if (-not $vals -or $vals.Count -eq 0) { continue }
                $v0 = $vals[0]
                if ($v0.displayValue) { $props[$attrName] = [string]$v0.displayValue }
                elseif ($v0.value)    { $props[$attrName] = [string]$v0.value }
            }
        }

        # Date de reference
        $rawDateRef = ""
        if ($DateRef -ieq "Created") {
            $rawDateRef = if ($obj.created) { [string]$obj.created } else { "" }
        } elseif ($DateRef -ieq "Updated") {
            $rawDateRef = if ($obj.updated) { [string]$obj.updated } else { "" }
        } else {
            if ($props.ContainsKey($DateRef)) { $rawDateRef = $props[$DateRef] }
            else { $rawDateRef = if ($obj.created) { [string]$obj.created } else { "" } }
        }

        $parsedDt = Parse-AssetDate $rawDateRef
        $moisRef  = if ($parsedDt) { $parsedDt.ToString("yyyy-MM") } else { "(non defini)" }
        $anneeRef = if ($parsedDt) { $parsedDt.ToString("yyyy") }    else { "(non defini)" }

        # Statut (si attributs charges)
        $statut = ""
        if ($needAttributes -and $props.ContainsKey("Statut")) { $statut = $props["Statut"] }

        $allObjects.Add([pscustomobject]@{
            ObjectId       = [string]$obj.id
            ObjectKey      = if ($obj.objectKey)                            { [string]$obj.objectKey }       else { "" }
            Label          = if ($obj.label)                                { [string]$obj.label }           else { "" }
            ObjectTypeName = if ($obj.objectType -and $obj.objectType.name) { [string]$obj.objectType.name } else { "" }
            Statut         = $statut
            MoisRef        = $moisRef
            AnneeRef       = $anneeRef
            DateRefValeur  = $rawDateRef
            Created        = if ($obj.created) { [string]$obj.created } else { "" }
            Updated        = if ($obj.updated) { [string]$obj.updated } else { "" }
        }) | Out-Null
        $totalCount++
    }

    $isLast   = if ($null -ne $resp.isLast) { [bool]$resp.isLast } else { $true }
    $startAt += $maxResults

    if ($totalCount % 500 -eq 0 -and $totalCount -gt 0) {
        Write-Info ("  Objets charges : " + $totalCount + " (page startAt=" + $startAt + ")")
    }
    if ($startAt -gt 100000) { Write-Warn "Pagination anormale (>100k), arret."; break }

    Start-Sleep -Milliseconds 50
}

Write-Info ("Total objets recuperes dans le schema '" + $SchemaName + "' : " + $allObjects.Count)

if ($allObjects.Count -eq 0) {
    Write-Warn "Aucun objet trouve dans ce schema."
    return
}

# ============================================================
# 8. ANALYSE ET AGRÉGATION MENSUELLE
# ============================================================
Write-Info "=== Analyse par periode mensuelle ==="

# 8a. Groupement par mois YYYY-MM
$grpMensuel = $allObjects | Group-Object MoisRef | Sort-Object Name -Descending

Write-Info ""
Write-Info "=== VENTILATION PAR MOIS (" + $DateRef + ") ==="
foreach ($g in $grpMensuel) {
    Write-Info ("  " + $g.Name + " : " + $g.Count + " objet(s)")
}

# 8b. Groupement par mois x Object Type
$grpMoisOTRows = New-Object System.Collections.Generic.List[object]
foreach ($gMois in $grpMensuel) {
    $subGrpOT = $gMois.Group | Group-Object ObjectTypeName | Sort-Object Count -Descending
    foreach ($gOT in $subGrpOT) {
        $grpMoisOTRows.Add([pscustomobject]@{
            MoisRef        = $gMois.Name
            ObjectTypeName = $gOT.Name
            NbObjets       = $gOT.Count
        }) | Out-Null
    }
}

# 8c. Groupement par mois x Statut (si statut renseigne)
$grpMoisStatutRows = New-Object System.Collections.Generic.List[object]
foreach ($gMois in $grpMensuel) {
    $subGrpStat = $gMois.Group | Group-Object Statut | Sort-Object Count -Descending
    foreach ($gStat in $subGrpStat) {
        $statName = if ([string]::IsNullOrWhiteSpace($gStat.Name)) { "(vide)" } else { $gStat.Name }
        $grpMoisStatutRows.Add([pscustomobject]@{
            MoisRef  = $gMois.Name
            Statut   = $statName
            NbObjets = $gStat.Count
        }) | Out-Null
    }
}

# ============================================================
# 9. EXPORTS CSV
# ============================================================
Write-Info ""
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

# 9a. Export de la liste detaillee
$csvDetail = Join-Path $exportsDir ("Assets-Export-ListeDetaillee_" + $runStamp + ".csv")
Export-CsvStrict -Path $csvDetail `
    -Headers @("ObjectKey", "ObjectId", "Label", "ObjectTypeName", "MoisRef", "AnneeRef", "Statut", "DateRefValeur", "Created", "Updated") `
    -Rows ($allObjects | Sort-Object MoisRef -Descending)

# 9b. Export de la synthese mensuelle globale
$synthMensuelleRows = foreach ($g in $grpMensuel) {
    [pscustomobject]@{
        MoisRef  = $g.Name
        NbObjets = $g.Count
        PctTotal = [math]::Round(($g.Count / $allObjects.Count) * 100, 1)
    }
}
$csvSynthMensuelle = Join-Path $exportsDir ("Assets-Export-SyntheseMensuelle_" + $runStamp + ".csv")
Export-CsvStrict -Path $csvSynthMensuelle `
    -Headers @("MoisRef", "NbObjets", "PctTotal") `
    -Rows $synthMensuelleRows

# 9c. Export de la synthese Mois x Object Type
$csvSynthOT = Join-Path $exportsDir ("Assets-Export-SyntheseParObjectType_" + $runStamp + ".csv")
Export-CsvStrict -Path $csvSynthOT `
    -Headers @("MoisRef", "ObjectTypeName", "NbObjets") `
    -Rows $grpMoisOTRows

# 9d. Export de la synthese Mois x Statut
$csvSynthStatut = Join-Path $exportsDir ("Assets-Export-SyntheseParStatut_" + $runStamp + ".csv")
Export-CsvStrict -Path $csvSynthStatut `
    -Headers @("MoisRef", "Statut", "NbObjets") `
    -Rows $grpMoisStatutRows

# ============================================================
# 10. RESUME FINAL
# ============================================================
Write-Info ""
Write-Info "============================================="
Write-Info "=== RESUME AUDIT SCHEMA EXPORT ==="
Write-Info "============================================="
Write-Info ""
Write-Info ("  Schema audite             : " + $SchemaName + " (id=" + $SchemaId + ")")
Write-Info ("  Champ date de reference   : " + $DateRef)
Write-Info ("  Total objets dans le schema: " + $allObjects.Count)
Write-Info ("  Nombre de mois identifies : " + $grpMensuel.Count)
Write-Info ""
Write-Info ("  Dernier mois le plus recent: " + $grpMensuel[0].Name + " (" + $grpMensuel[0].Count + " objets)")
if ($grpMensuel.Count -gt 1) {
    Write-Info ("  Mois le plus ancien       : " + $grpMensuel[-1].Name + " (" + $grpMensuel[-1].Count + " objets)")
}
Write-Info ""
Write-Info ("  Exports CSV disponibles dans : " + $exportsDir)
Write-Info ("  Fichier log                  : " + $logFile)
Write-Log ("=== FIN " + $scriptName + " ===") "INFO"