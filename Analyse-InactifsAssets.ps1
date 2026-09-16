<#
Analyse-InactifsAssets.ps1
Analyse exhaustive de TOUS les objets Assets au statut Inactif.
Inclut les objets sans date de sortie.
Produit un inventaire complet avec categorisation pour decision de purge.
Mode LECTURE SEULE : aucune modification n'est effectuee.
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
$scriptName = "Analyse-InactifsAssets"

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

    $s9  = if ($s.Length -ge 9)  { $s.Substring(0, 9).Trim()  } else { $s }
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
                Created        = if ($obj.created) { [string]$obj.created } else { "" }
                Updated        = if ($obj.updated) { [string]$obj.updated } else { "" }
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
# 7. CATEGORISATION DE TOUS LES INACTIFS
# ============================================================
Write-Info "=== Categorisation des objets ==="

$actifs         = New-Object System.Collections.Generic.List[object]
$inactifsCat1   = New-Object System.Collections.Generic.List[object]  # Date Sortie < seuil
$inactifsCat2   = New-Object System.Collections.Generic.List[object]  # Date Sortie >= seuil
$inactifsCat3   = New-Object System.Collections.Generic.List[object]  # Sans date de sortie
$autresStatuts  = New-Object System.Collections.Generic.List[object]

foreach ($obj in $allObjects) {
    $statut = $obj.Statut

    if ($statut -eq "Actif") {
        $actifs.Add($obj) | Out-Null
        continue
    }

    if ($statut -eq "Inactif" -or $statut -eq "inactif") {
        $dtSortie = Parse-DateSortie $obj.DateSortie

        if (-not $dtSortie) {
            $inactifsCat3.Add($obj) | Out-Null
        } elseif ($dtSortie -lt $DateSortieSeuil) {
            $inactifsCat1.Add($obj) | Out-Null
        } else {
            $inactifsCat2.Add($obj) | Out-Null
        }
        continue
    }

    # Autre statut (vide, autre valeur)
    $autresStatuts.Add($obj) | Out-Null
}

$totalInactifs = $inactifsCat1.Count + $inactifsCat2.Count + $inactifsCat3.Count

Write-Info ""
Write-Info "=== VENTILATION GLOBALE ==="
Write-Info ("  Total objets                                    : " + $allObjects.Count)
Write-Info ("  Actifs                                          : " + $actifs.Count)
Write-Info ("  Inactifs (total)                                : " + $totalInactifs)
Write-Info ("    Cat.1 - Date Sortie < " + $DateSortieSeuil.ToString("yyyy-MM-dd") + "          : " + $inactifsCat1.Count)
Write-Info ("    Cat.2 - Date Sortie >= " + $DateSortieSeuil.ToString("yyyy-MM-dd") + "         : " + $inactifsCat2.Count)
Write-Info ("    Cat.3 - Sans date de sortie                   : " + $inactifsCat3.Count)
Write-Info ("  Autres statuts                                  : " + $autresStatuts.Count)

# ============================================================
# 8. ANALYSE DETAILLEE PAR CATEGORIE
# ============================================================
Write-Info ""
Write-Info "=== ANALYSE Cat.1 : Inactifs avec Date Sortie < seuil (candidats prioritaires) ==="

if ($inactifsCat1.Count -gt 0) {
    $grpDir1 = $inactifsCat1 | Group-Object Direction | Sort-Object Count -Descending
    Write-Info "  Par Direction :"
    foreach ($g in $grpDir1) {
        $n = if ([string]::IsNullOrWhiteSpace($g.Name)) { "(non defini)" } else { $g.Name }
        Write-Info ("    " + $n + " : " + $g.Count)
    }

    $grpType1 = $inactifsCat1 | Group-Object TypeRessource | Sort-Object Count -Descending
    Write-Info "  Par Type ressource :"
    foreach ($g in $grpType1) {
        $n = if ([string]::IsNullOrWhiteSpace($g.Name)) { "(non defini)" } else { $g.Name }
        Write-Info ("    " + $n + " : " + $g.Count)
    }

    $grpMonth1 = $inactifsCat1 | ForEach-Object {
        $dt = Parse-DateSortie $_.DateSortie
        if ($dt) { [pscustomobject]@{ MoisSortie = $dt.ToString("yyyy-MM") } }
    } | Group-Object MoisSortie | Sort-Object Name
    Write-Info "  Par mois de sortie :"
    foreach ($g in $grpMonth1) { Write-Info ("    " + $g.Name + " : " + $g.Count) }
} else {
    Write-Info "  Aucun objet dans cette categorie."
}

Write-Info ""
Write-Info "=== ANALYSE Cat.2 : Inactifs avec Date Sortie >= seuil (conserves) ==="

if ($inactifsCat2.Count -gt 0) {
    $grpDir2 = $inactifsCat2 | Group-Object Direction | Sort-Object Count -Descending
    Write-Info "  Par Direction :"
    foreach ($g in $grpDir2) {
        $n = if ([string]::IsNullOrWhiteSpace($g.Name)) { "(non defini)" } else { $g.Name }
        Write-Info ("    " + $n + " : " + $g.Count)
    }
} else {
    Write-Info "  Aucun objet dans cette categorie."
}

Write-Info ""
Write-Info "=== ANALYSE Cat.3 : Inactifs SANS date de sortie ==="

if ($inactifsCat3.Count -gt 0) {
    $grpDir3 = $inactifsCat3 | Group-Object Direction | Sort-Object Count -Descending
    Write-Info "  Par Direction :"
    foreach ($g in $grpDir3) {
        $n = if ([string]::IsNullOrWhiteSpace($g.Name)) { "(non defini)" } else { $g.Name }
        Write-Info ("    " + $n + " : " + $g.Count)
    }

    $grpType3 = $inactifsCat3 | Group-Object TypeRessource | Sort-Object Count -Descending
    Write-Info "  Par Type ressource :"
    foreach ($g in $grpType3) {
        $n = if ([string]::IsNullOrWhiteSpace($g.Name)) { "(non defini)" } else { $g.Name }
        Write-Info ("    " + $n + " : " + $g.Count)
    }

    # Derniere mise a jour pour evaluer l'anciennete
    Write-Info "  Par annee de derniere mise a jour (Updated) :"
    $grpUpdated3 = $inactifsCat3 | ForEach-Object {
        $dt = Parse-DateSortie $_.Updated
        if ($dt) { [pscustomobject]@{ Annee = $dt.ToString("yyyy") } }
        else     { [pscustomobject]@{ Annee = "(non parsable)" } }
    } | Group-Object Annee | Sort-Object Name
    foreach ($g in $grpUpdated3) { Write-Info ("    " + $g.Name + " : " + $g.Count) }

    # Derniere mise a jour pour evaluer l'anciennete
    Write-Info "  Par annee de creation (Created) :"
    $grpCreated3 = $inactifsCat3 | ForEach-Object {
        $dt = Parse-DateSortie $_.Created
        if ($dt) { [pscustomobject]@{ Annee = $dt.ToString("yyyy") } }
        else     { [pscustomobject]@{ Annee = "(non parsable)" } }
    } | Group-Object Annee | Sort-Object Name
    foreach ($g in $grpCreated3) { Write-Info ("    " + $g.Name + " : " + $g.Count) }
} else {
    Write-Info "  Aucun objet dans cette categorie."
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
    Write-Info ("  CSV exporte : " + $Path + " (" + $Rows.Count + " lignes)")
}

$csvHeaders = @(
    "ObjectKey", "Label", "ObjectType", "Categorie", "Statut",
    "Direction", "TypeRessource",
    "DateEntree", "DateSortie", "MotifSortie",
    "Prenom", "Nom", "Matricule",
    "Created", "Updated"
)

# Construire la liste complete avec la categorie
$allInactifs = New-Object System.Collections.Generic.List[object]

foreach ($obj in $inactifsCat1) {
    $allInactifs.Add([pscustomobject]@{
        ObjectKey     = $obj.ObjectKey;      Label         = $obj.Label
        ObjectType    = $obj.ObjectTypeName; Categorie     = "Cat.1 - Sortie < seuil"
        Statut        = $obj.Statut;         Direction     = $obj.Direction
        TypeRessource = $obj.TypeRessource;  DateEntree    = $obj.DateEntree
        DateSortie    = $obj.DateSortie;     MotifSortie   = $obj.MotifSortie
        Prenom        = $obj.Prenom;         Nom           = $obj.Nom
        Matricule     = $obj.Matricule;      Created       = $obj.Created
        Updated       = $obj.Updated
    }) | Out-Null
}
foreach ($obj in $inactifsCat2) {
    $allInactifs.Add([pscustomobject]@{
        ObjectKey     = $obj.ObjectKey;      Label         = $obj.Label
        ObjectType    = $obj.ObjectTypeName; Categorie     = "Cat.2 - Sortie >= seuil"
        Statut        = $obj.Statut;         Direction     = $obj.Direction
        TypeRessource = $obj.TypeRessource;  DateEntree    = $obj.DateEntree
        DateSortie    = $obj.DateSortie;     MotifSortie   = $obj.MotifSortie
        Prenom        = $obj.Prenom;         Nom           = $obj.Nom
        Matricule     = $obj.Matricule;      Created       = $obj.Created
        Updated       = $obj.Updated
    }) | Out-Null
}
foreach ($obj in $inactifsCat3) {
    $allInactifs.Add([pscustomobject]@{
        ObjectKey     = $obj.ObjectKey;      Label         = $obj.Label
        ObjectType    = $obj.ObjectTypeName; Categorie     = "Cat.3 - Sans date sortie"
        Statut        = $obj.Statut;         Direction     = $obj.Direction
        TypeRessource = $obj.TypeRessource;  DateEntree    = $obj.DateEntree
        DateSortie    = $obj.DateSortie;     MotifSortie   = $obj.MotifSortie
        Prenom        = $obj.Prenom;         Nom           = $obj.Nom
        Matricule     = $obj.Matricule;      Created       = $obj.Created
        Updated       = $obj.Updated
    }) | Out-Null
}

# 9a. Liste exhaustive de tous les inactifs
$csvAll = Join-Path $exportsDir ("Assets-TousInactifs_" + $runStamp + ".csv")
Export-CsvStrict -Path $csvAll -Headers $csvHeaders -Rows ($allInactifs | Sort-Object Categorie, DateSortie, ObjectKey)

# 9b. Synthese par categorie et direction
$syntheseRows = New-Object System.Collections.Generic.List[object]
foreach ($cat in @("Cat.1 - Sortie < seuil", "Cat.2 - Sortie >= seuil", "Cat.3 - Sans date sortie")) {
    $catObjs = $allInactifs | Where-Object { $_.Categorie -eq $cat }
    $grp     = $catObjs | Group-Object Direction | Sort-Object Count -Descending
    foreach ($g in $grp) {
        $dirName = if ([string]::IsNullOrWhiteSpace($g.Name)) { "(non defini)" } else { $g.Name }
        $syntheseRows.Add([pscustomobject]@{
            "Categorie"    = $cat
            "Direction"    = $dirName
            "Nb Objets"    = $g.Count
        }) | Out-Null
    }
}
$csvSynth = Join-Path $exportsDir ("Assets-Synthese-InactifsParCategorie_" + $runStamp + ".csv")
Export-CsvStrict -Path $csvSynth -Headers @("Categorie", "Direction", "Nb Objets") -Rows $syntheseRows

# 9c. Synthese globale
$csvGlob = Join-Path $exportsDir ("Assets-Synthese-Globale_" + $runStamp + ".csv")
$globRows = New-Object System.Collections.Generic.List[object]
$globRows.Add([pscustomobject]@{ Metrique = "Total objets (OT " + ($ObjectTypeIds -join "+") + ")"; Valeur = [string]$allObjects.Count }) | Out-Null
$globRows.Add([pscustomobject]@{ Metrique = "Actifs"; Valeur = [string]$actifs.Count }) | Out-Null
$globRows.Add([pscustomobject]@{ Metrique = "Inactifs (total)"; Valeur = [string]$totalInactifs }) | Out-Null
$globRows.Add([pscustomobject]@{ Metrique = "  Cat.1 - Date Sortie < " + $DateSortieSeuil.ToString("yyyy-MM-dd"); Valeur = [string]$inactifsCat1.Count }) | Out-Null
$globRows.Add([pscustomobject]@{ Metrique = "  Cat.2 - Date Sortie >= " + $DateSortieSeuil.ToString("yyyy-MM-dd"); Valeur = [string]$inactifsCat2.Count }) | Out-Null
$globRows.Add([pscustomobject]@{ Metrique = "  Cat.3 - Sans date de sortie"; Valeur = [string]$inactifsCat3.Count }) | Out-Null
$globRows.Add([pscustomobject]@{ Metrique = "Autres statuts"; Valeur = [string]$autresStatuts.Count }) | Out-Null
$globRows.Add([pscustomobject]@{ Metrique = "---"; Valeur = "---" }) | Out-Null
$globRows.Add([pscustomobject]@{ Metrique = "Gain quota si suppression Cat.1"; Valeur = ("-" + $inactifsCat1.Count + " objets (" + $(if($allObjects.Count -gt 0){[math]::Round(($inactifsCat1.Count/$allObjects.Count)*100,1)}else{0}) + "%)") }) | Out-Null
$globRows.Add([pscustomobject]@{ Metrique = "Gain quota si suppression Cat.1 + Cat.3"; Valeur = ("-" + ($inactifsCat1.Count + $inactifsCat3.Count) + " objets (" + $(if($allObjects.Count -gt 0){[math]::Round((($inactifsCat1.Count+$inactifsCat3.Count)/$allObjects.Count)*100,1)}else{0}) + "%)") }) | Out-Null
$globRows.Add([pscustomobject]@{ Metrique = "Gain quota si suppression TOUS inactifs"; Valeur = ("-" + $totalInactifs + " objets (" + $(if($allObjects.Count -gt 0){[math]::Round(($totalInactifs/$allObjects.Count)*100,1)}else{0}) + "%)") }) | Out-Null
Export-CsvStrict -Path $csvGlob -Headers @("Metrique", "Valeur") -Rows $globRows

# ============================================================
# 10. RESUME FINAL
# ============================================================
Write-Info ""
Write-Info "============================================="
Write-Info "=== RESUME FINAL ==="
Write-Info "============================================="
Write-Info ""
Write-Info ("  Object Types audites                            : " + ($ObjectTypeIds -join ", "))
Write-Info ("  Seuil Date Sortie                               : " + $DateSortieSeuil.ToString("yyyy-MM-dd"))
Write-Info ""
Write-Info ("  Total objets                                    : " + $allObjects.Count)
Write-Info ("  Actifs                                          : " + $actifs.Count)
Write-Info ("  Inactifs (total)                                : " + $totalInactifs)
Write-Info ("    Cat.1 - Date Sortie < seuil                   : " + $inactifsCat1.Count)
Write-Info ("    Cat.2 - Date Sortie >= seuil                  : " + $inactifsCat2.Count)
Write-Info ("    Cat.3 - Sans date de sortie                   : " + $inactifsCat3.Count)
Write-Info ("  Autres statuts                                  : " + $autresStatuts.Count)
Write-Info ""
Write-Info "  === SCENARIOS DE GAIN QUOTA ==="
$gainCat1     = $inactifsCat1.Count
$gainCat1et3  = $inactifsCat1.Count + $inactifsCat3.Count
$gainTous     = $totalInactifs
$pctCat1      = if ($allObjects.Count -gt 0) { [math]::Round(($gainCat1 / $allObjects.Count) * 100, 1) } else { 0 }
$pctCat1et3   = if ($allObjects.Count -gt 0) { [math]::Round(($gainCat1et3 / $allObjects.Count) * 100, 1) } else { 0 }
$pctTous      = if ($allObjects.Count -gt 0) { [math]::Round(($gainTous / $allObjects.Count) * 100, 1) } else { 0 }

Write-Info ("  Scenario A (Cat.1 seule)        : -" + $gainCat1 + " objets (" + $pctCat1 + "%)")
Write-Info ("  Scenario B (Cat.1 + Cat.3)      : -" + $gainCat1et3 + " objets (" + $pctCat1et3 + "%)")
Write-Info ("  Scenario C (tous inactifs)      : -" + $gainTous + " objets (" + $pctTous + "%)")
Write-Info ""
Write-Info ("  !!! MODE LECTURE SEULE - AUCUNE MODIFICATION EFFECTUEE !!!")
Write-Info ""
Write-Info ("  Exports : " + $exportsDir)
Write-Info ("  Log     : " + $logFile)
Write-Log ("=== FIN " + $scriptName + " ===") "INFO"
