<#
Audit-ObjetsAssetsObsoletes.ps1
Identifie les objets Assets candidats a la suppression pour liberation de quota.
Critere : Statut = "Inactif" ET Date Sortie < 01/10/2025
Mode DRY RUN uniquement : aucune suppression n'est effectuee.
#>

[CmdletBinding()]
param(
    [switch]$UseSystemProxy = $true,
    [switch]$ProxyUseDefaultCredentials = $true,
    [string]$ProxyUrl,
    [string]$AssetsWorkspaceId = "2898703d-5f14-4527-9f8a-ad32b9279ca1",
    [string[]]$ObjectTypeIds   = @("68", "69"),
    [datetime]$DateSortieSeuil = [datetime]::ParseExact("2025-10-01", "yyyy-MM-dd", $null)
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
$scriptName = "Audit-ObjetsAssetsObsoletes"

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
Write-Info ("Script dir       : " + $scriptDir)
Write-Info ("Date seuil sortie: " + $DateSortieSeuil.ToString("yyyy-MM-dd"))

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
        $sc = $null
        try { $sc = $_.Exception.Response.StatusCode.value__ } catch {}
        if ($sc -ne 404) {
            $body = Get-WebExceptionBody $_.Exception
            Write-ErrLog ("GET " + $Url + " : " + $_.Exception.Message + "`n" + $body)
        }
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

# ============================================================
# 5. VARIABLES ASSETS + HELPERS
# ============================================================
$assetsAqlUrl = $jiraBaseUrl + "/gateway/api/jsm/assets/workspace/" + $AssetsWorkspaceId + "/v1/object/aql"
Write-Info ("Assets workspace : " + $AssetsWorkspaceId)
Write-Info ("Object Types     : " + ($ObjectTypeIds -join ", "))

$eAcc       = [char]233
$dateEntKey = "Date Entr" + $eAcc + "e"
$prenomKey  = "Pr" + $eAcc + "nom"

function Parse-DateSortie([string]$DateStr) {
    if ([string]::IsNullOrWhiteSpace($DateStr)) { return $null }
    $s  = $DateStr.Trim()
    $ci = [System.Globalization.CultureInfo]::InvariantCulture

    # Extraire les 9 premiers chars pour dd/MMM/yy (ex: "31/Dec/24")
    $s9  = if ($s.Length -ge 9)  { $s.Substring(0, 9).Trim()  } else { $s }
    # Extraire les 10 premiers chars pour yyyy-MM-dd et dd/MM/yyyy
    $s10 = if ($s.Length -ge 10) { $s.Substring(0, 10).Trim() } else { $s }

    # Format Assets natif : 31/Dec/24
    try { return [datetime]::ParseExact($s9, "dd/MMM/yy",   $ci) } catch {}
    # Format Assets long  : 31/Dec/2024
    try { return [datetime]::ParseExact($s9, "dd/MMM/yyyy", $ci) } catch {}
    # ISO : yyyy-MM-dd
    try { return [datetime]::ParseExact($s10, "yyyy-MM-dd",  $ci) } catch {}
    # FR  : dd/MM/yyyy
    try { return [datetime]::ParseExact($s10, "dd/MM/yyyy",  $ci) } catch {}
    # Timestamp epoch ms
    if ($s -match "^\d{13}$") {
        try { return (Get-Date "1970-01-01").AddMilliseconds([long]$s) } catch {}
    }

    return $null
}

# ============================================================
# 6. RECUPERATION DES OBJETS ASSETS (avec attributs)
# ============================================================
Write-Info "=== Recuperation des objets Assets ==="

$allObjects = New-Object System.Collections.Generic.List[object]

foreach ($otId in $ObjectTypeIds) {
    $aqlQuery   = "objectTypeId = " + $otId
    $startAt    = 0
    $maxResults = 50
    $isLast     = $false
    $otCount    = 0

    Write-Info ("  AQL : " + $aqlQuery)

    while (-not $isLast) {
        $url      = $assetsAqlUrl + "?startAt=" + $startAt + "&maxResults=" + $maxResults + "&includeAttributes=true"
        $bodyObj  = @{ qlQuery = $aqlQuery }
        $bodyJson = $bodyObj | ConvertTo-Json -Depth 3

        try {
            $resp = Invoke-ApiPostUtf8 -Url $url -Headers $jiraHeaders -JsonBody $bodyJson
        } catch {
            Write-ErrLog ("Erreur AQL OT=" + $otId + " startAt=" + $startAt + " : " + $_.Exception.Message)
            break
        }
        if (-not $resp) { break }

        # Dictionnaire attribut id -> nom
        $attrDict = @{}
        if ($resp.objectTypeAttributes) {
            foreach ($ota in $resp.objectTypeAttributes) {
                if ($ota.id -and $ota.name) { $attrDict[[string]$ota.id] = [string]$ota.name }
            }
        }

        $objects = $resp.values
        if (-not $objects -or $objects.Count -eq 0) { break }

        foreach ($obj in $objects) {
            if (-not $obj.id) { continue }
            $props = @{}

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

            $dateSortieStr = ""
            if ($props.ContainsKey("Date Sortie")) { $dateSortieStr = $props["Date Sortie"] }
            elseif ($props.ContainsKey("Date PLD")) { $dateSortieStr = $props["Date PLD"] }

            $allObjects.Add([pscustomobject]@{
                ObjectId       = [string]$obj.id
                ObjectKey      = if ($obj.objectKey) { [string]$obj.objectKey } else { "" }
                Label          = if ($obj.label)     { [string]$obj.label }     else { "" }
                ObjectType     = $otId
                ObjectTypeName = if ($obj.objectType -and $obj.objectType.name) { [string]$obj.objectType.name } else { "" }
                Statut         = if ($props.ContainsKey("Statut"))        { $props["Statut"] }        else { "" }
                Direction      = if ($props.ContainsKey("Direction"))     { $props["Direction"] }     else { "" }
                TypeRessource  = if ($props.ContainsKey("Type ressource")){ $props["Type ressource"] }else { "" }
                DateEntree     = if ($props.ContainsKey($dateEntKey))     { $props[$dateEntKey] }     else { "" }
                DateSortie     = $dateSortieStr
                MotifSortie    = if ($props.ContainsKey("Motif Sortie"))  { $props["Motif Sortie"] }  else { "" }
                Prenom         = if ($props.ContainsKey($prenomKey))      { $props[$prenomKey] }      else { "" }
                Nom            = if ($props.ContainsKey("Nom"))           { $props["Nom"] }           else { "" }
                Matricule      = if ($props.ContainsKey("Matricule"))     { $props["Matricule"] }     else { "" }
            }) | Out-Null
            $otCount++
        }

        $isLast   = if ($null -ne $resp.isLast) { [bool]$resp.isLast } else { $true }
        $startAt += $maxResults

        if ($startAt % 200 -eq 0) {
            Write-Info ("  OT " + $otId + " : " + $otCount + " charges (startAt=" + $startAt + ")")
        }
        if ($startAt -gt 50000) { Write-Warn "Pagination anormale, arret."; break }
        Start-Sleep -Milliseconds 150
    }

    Write-Info ("  OT " + $otId + " termine : " + $otCount + " objets")
}

Write-Info ("  Total objets recuperes : " + $allObjects.Count)

# ============================================================
# 7. ANALYSE : IDENTIFICATION DES CANDIDATS SUPPRESSION
# ============================================================
Write-Info "=== Identification des candidats a la suppression ==="
Write-Info ("  Critere : Statut = 'Inactif' ET Date Sortie < " + $DateSortieSeuil.ToString("yyyy-MM-dd"))
Write-Info ("  Note    : Inactifs SANS date de sortie = exclus du perimetre (non fiables)")

$candidats      = New-Object System.Collections.Generic.List[object]
$actifs         = 0
$inactifsSansDt = 0
$inactifsAvant  = 0   # date sortie < seuil = candidats
$inactifsApres  = 0   # date sortie >= seuil = conserves
$autresStatuts  = 0

foreach ($obj in $allObjects) {
    $statut = $obj.Statut

    if ($statut -eq "Actif") { $actifs++; continue }

    if ($statut -eq "Inactif" -or $statut -eq "inactif") {
        $dtSortie = Parse-DateSortie $obj.DateSortie

        if (-not $dtSortie) {
            $inactifsSansDt++
            continue   # Exclus du perimetre — ni candidat, ni comptabilise
        }

        if ($dtSortie -lt $DateSortieSeuil) {
            $inactifsAvant++
            $candidats.Add([pscustomobject]@{
                ObjectId      = $obj.ObjectId
                ObjectKey     = $obj.ObjectKey
                Label         = $obj.Label
                ObjectType    = $obj.ObjectTypeName
                Statut        = $obj.Statut
                Direction     = $obj.Direction
                TypeRessource = $obj.TypeRessource
                DateEntree    = $obj.DateEntree
                DateSortie    = $obj.DateSortie
                MotifSortie   = $obj.MotifSortie
                Prenom        = $obj.Prenom
                Nom           = $obj.Nom
                Matricule     = $obj.Matricule
            }) | Out-Null
        } else {
            $inactifsApres++
        }
    } else {
        $autresStatuts++
    }
}

# Seuls les inactifs avec date sont dans le perimetre
$inactifsPerimetre = $inactifsAvant + $inactifsApres
$nbCandidats       = $candidats.Count
$pctReduction      = if ($allObjects.Count -gt 0) { [math]::Round(($nbCandidats / $allObjects.Count) * 100, 1) } else { 0 }

Write-Info ""
Write-Info "=== Ventilation des objets ==="
Write-Info ("  Total objets                            : " + $allObjects.Count)
Write-Info ("  Actifs                                  : " + $actifs)
Write-Info ("  Inactifs avec date de sortie            : " + $inactifsPerimetre)
Write-Info ("    - Date Sortie < " + $DateSortieSeuil.ToString("yyyy-MM-dd") + "       : " + $inactifsAvant + " [CANDIDATS SUPPRESSION]")
Write-Info ("    - Date Sortie >= " + $DateSortieSeuil.ToString("yyyy-MM-dd") + "      : " + $inactifsApres + " [CONSERVES]")
Write-Info ("  Inactifs SANS date de sortie (exclus)   : " + $inactifsSansDt)
Write-Info ("  Autres statuts                          : " + $autresStatuts)

# ============================================================
# 8. IMPACT QUOTA
# ============================================================
Write-Info ""
Write-Info "=== IMPACT QUOTA ESTIME ==="

$totalObjets      = $allObjects.Count
$nbCandidats      = $candidats.Count
$objetsApresActif = $totalObjets - $nbCandidats
$pctReduction     = if ($totalObjets -gt 0) { [math]::Round(($nbCandidats / $totalObjets) * 100, 1) } else { 0 }

Write-Info ("  Objets actuels (OT " + ($ObjectTypeIds -join "+") + ")  : " + $totalObjets)
Write-Info ("  Objets a supprimer (dry run)             : " + $nbCandidats)
Write-Info ("  Objets restants apres purge              : " + $objetsApresActif)
Write-Info ("  Reduction quota estimee                  : -" + $nbCandidats + " objets (" + $pctReduction + "%)")

# Ventilation par direction
Write-Info ""
Write-Info "=== Candidats par Direction ==="
$grpDir = $candidats | Group-Object Direction | Sort-Object Count -Descending
foreach ($g in $grpDir) {
    $dirName = if ([string]::IsNullOrWhiteSpace($g.Name)) { "(non defini)" } else { $g.Name }
    Write-Info ("  " + $dirName + " : " + $g.Count)
}

# Ventilation par type de ressource
Write-Info ""
Write-Info "=== Candidats par Type ressource ==="
$grpType = $candidats | Group-Object TypeRessource | Sort-Object Count -Descending
foreach ($g in $grpType) {
    $typeName = if ([string]::IsNullOrWhiteSpace($g.Name)) { "(non defini)" } else { $g.Name }
    Write-Info ("  " + $typeName + " : " + $g.Count)
}

# Ventilation par mois de sortie
Write-Info ""
Write-Info "=== Candidats par mois de sortie ==="
$grpMonth = $candidats | ForEach-Object {
    $dt = Parse-DateSortie $_.DateSortie
    if ($dt) { [pscustomobject]@{ MoisSortie = $dt.ToString("yyyy-MM"); Obj = $_ } }
} | Group-Object MoisSortie | Sort-Object Name

foreach ($g in $grpMonth) {
    Write-Info ("  " + $g.Name + " : " + $g.Count + " objets")
}

# ============================================================
# 9. EXPORT CSV
# ============================================================
Write-Info ""
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

$csvHeaders = @(
    "ObjectKey", "Label", "ObjectType", "Statut",
    "Direction", "TypeRessource",
    "DateEntree", "DateSortie", "MotifSortie",
    "Prenom", "Nom", "Matricule"
)

# 9a. Liste des candidats suppression
$candidatsSorted = $candidats | Sort-Object DateSortie, ObjectKey
$csvCandidats    = Join-Path $exportsDir ("Assets-Candidats-Suppression_" + $runStamp + ".csv")
Export-CsvStrict -Path $csvCandidats -Headers $csvHeaders -Rows $candidatsSorted

# 9b. Synthese par direction
$syntheseDir = New-Object System.Collections.Generic.List[object]
foreach ($g in $grpDir) {
    $dirName = if ([string]::IsNullOrWhiteSpace($g.Name)) { "(non defini)" } else { $g.Name }
    $syntheseDir.Add([pscustomobject]@{
        "Direction"       = $dirName
        "Nb Candidats"    = $g.Count
        "Pct du Total"    = [math]::Round(($g.Count / [math]::Max($nbCandidats, 1)) * 100, 1)
    }) | Out-Null
}
$csvSynthDir = Join-Path $exportsDir ("Assets-Synthese-ParDirection_" + $runStamp + ".csv")
Export-CsvStrict -Path $csvSynthDir -Headers @("Direction", "Nb Candidats", "Pct du Total") -Rows $syntheseDir

# 9c. Synthese globale
$syntheseGlob = New-Object System.Collections.Generic.List[object]
$syntheseGlob.Add([pscustomobject]@{
    "Metrique"       = "Objets actuels (OT " + ($ObjectTypeIds -join "+") + ")"
    "Valeur"         = $totalObjets
}) | Out-Null
$syntheseGlob.Add([pscustomobject]@{
    "Metrique"       = "Objets Actifs"
    "Valeur"         = $actifs
}) | Out-Null
$syntheseGlob.Add([pscustomobject]@{
    "Metrique"       = "Objets Inactifs (total)"
    "Valeur"         = $inactifs
}) | Out-Null
$syntheseGlob.Add([pscustomobject]@{
    "Metrique"       = "Candidats suppression (Inactif + Date Sortie < " + $DateSortieSeuil.ToString("yyyy-MM-dd") + ")"
    "Valeur"         = $nbCandidats
}) | Out-Null
$syntheseGlob.Add([pscustomobject]@{
    "Metrique"       = "Inactifs sans date de sortie"
    "Valeur"         = $inactifsSansDt
}) | Out-Null
$syntheseGlob.Add([pscustomobject]@{
    "Metrique"       = "Inactifs recents (Date Sortie >= seuil)"
    "Valeur"         = $inactifsApres
}) | Out-Null
$syntheseGlob.Add([pscustomobject]@{
    "Metrique"       = "Objets restants apres purge"
    "Valeur"         = $objetsApresActif
}) | Out-Null
$syntheseGlob.Add([pscustomobject]@{
    "Metrique"       = "Reduction quota estimee (%)"
    "Valeur"         = [string]$pctReduction + "%"
}) | Out-Null

$csvSynthGlob = Join-Path $exportsDir ("Assets-Synthese-Globale_" + $runStamp + ".csv")
Export-CsvStrict -Path $csvSynthGlob -Headers @("Metrique", "Valeur") -Rows $syntheseGlob

# ============================================================
# 10. RESUME FINAL
# ============================================================
Write-Info ""
Write-Info "============================================="
Write-Info "=== RESUME FINAL - DRY RUN ==="
Write-Info "============================================="
Write-Info ""
Write-Info ("  Object Types audites                : " + ($ObjectTypeIds -join ", "))
Write-Info ("  Critere suppression                 : Statut='Inactif' ET DateSortie < " + $DateSortieSeuil.ToString("yyyy-MM-dd"))
Write-Info ""
Write-Info ("  Total objets dans le perimetre      : " + $totalObjets)
Write-Info ("  Objets Actifs                       : " + $actifs)
Write-Info ("  Objets Inactifs                     : " + $inactifs)
Write-Info ("    dont candidats suppression        : " + $nbCandidats + " (" + $pctReduction + "%)")
Write-Info ("    dont inactifs sans date sortie    : " + $inactifsSansDt)
Write-Info ("    dont inactifs recents (>= seuil)  : " + $inactifsApres)
Write-Info ""
Write-Info ("  >>> GAIN QUOTA ESTIME : -" + $nbCandidats + " objets (" + $pctReduction + "%) <<<")
Write-Info ""
Write-Info ("  !!! AUCUNE SUPPRESSION EFFECTUEE - MODE DRY RUN !!!")
Write-Info ""
Write-Info ("  Exports : " + $exportsDir)
Write-Info ("  Log     : " + $logFile)
Write-Log ("=== FIN " + $scriptName + " ===") "INFO"
