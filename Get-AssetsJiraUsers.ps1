# ============================================================
# Nom          : Get-AssetsJiraUsers.ps1
# Description  : Extraction exhaustive des fiches personnes /
#                utilisateurs du Referentiel Personne (RP)
#                dans JSM Assets (CMDB Jiradot).
# Version      : 1.1
# Date         : 11/09/2026
# Auteur       : DSIM / HM_DSIM_PACT
# Prerequis    : Fichier secrets/jira-jiradot.cred.xml valide
# Usage        : .\Get-AssetsJiraUsers.ps1 [-IncludeInactifs]
# ============================================================

[CmdletBinding()]
param(
    [switch]$IncludeInactifs = $false
)

# Culture FR — séparateur décimal virgule et formatage fr-FR
[System.Threading.Thread]::CurrentThread.CurrentCulture   = [System.Globalization.CultureInfo]::GetCultureInfo("fr-FR")
[System.Threading.Thread]::CurrentThread.CurrentUICulture = [System.Globalization.CultureInfo]::GetCultureInfo("fr-FR")

Set-StrictMode -Off
$ErrorActionPreference = "Continue"

# ============================================================
# 0. CONSTANTES ET CHEMINS (identique aux autres scripts)
# ============================================================
$scriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
if (-not $scriptRoot) { $scriptRoot = Get-Location }

$secretsDir  = Join-Path $scriptRoot "secrets"
$exportsDir  = Join-Path $scriptRoot "exports"
$logsDir     = Join-Path $scriptRoot "logs"
$runStamp    = Get-Date -Format "yyyyMMdd_HHmmss"
$logFile     = Join-Path $logsDir ("Get-AssetsJiraUsers_" + $runStamp + ".log")

$ProxyUrl                   = ""
$UseSystemProxy             = $true
$ProxyUseDefaultCredentials = $true

foreach ($d in @($exportsDir, $secretsDir, $logsDir)) {
    if (-not (Test-Path $d)) { New-Item -ItemType Directory -Path $d -Force | Out-Null }
}

# ============================================================
# 1. LOGGING ET HELPERS
# ============================================================
function Write-Log([string]$Message, [string]$Level = "INFO") {
    $ts   = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $line = "[" + $ts + "] [" + $Level + "] " + $Message
    try { Add-Content -Path $logFile -Value $line -Encoding UTF8 } catch {}
}
function Write-Info([string]$Message) {
    Write-Host ("[INFO] " + $Message) -ForegroundColor Cyan
    Write-Log $Message "INFO"
}
function Write-Warn([string]$Message) {
    Write-Warning $Message
    Write-Log $Message "WARN"
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

# ============================================================
# 2. GESTION PROXY
# ============================================================
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

function Invoke-ApiGet([string]$Url, [hashtable]$Headers) {
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    $params = @{ Method='GET'; Uri=$Url; Headers=$Headers; ContentType='application/json'; UseBasicParsing=$true; ErrorAction='Stop' }
    $px = Get-ProxyParams $Url; foreach ($k in $px.Keys) { $params[$k] = $px[$k] }
    $resp   = Invoke-WebRequest @params
    $stream = $resp.RawContentStream; $stream.Position = 0
    $reader = New-Object System.IO.StreamReader($stream, [System.Text.Encoding]::UTF8, $true)
    $raw    = $reader.ReadToEnd(); $reader.Close()
    return ($raw | ConvertFrom-Json)
}

function Invoke-ApiPostJson([string]$Url, [hashtable]$Headers, [object]$BodyObj) {
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    $json   = $BodyObj | ConvertTo-Json -Depth 10 -Compress
    $params = @{ Method='POST'; Uri=$Url; Headers=$Headers; ContentType='application/json'; Body=[Text.Encoding]::UTF8.GetBytes($json); UseBasicParsing=$true; ErrorAction='Stop' }
    $px = Get-ProxyParams $Url; foreach ($k in $px.Keys) { $params[$k] = $px[$k] }
    $resp   = Invoke-WebRequest @params
    $stream = $resp.RawContentStream; $stream.Position = 0
    $reader = New-Object System.IO.StreamReader($stream, [System.Text.Encoding]::UTF8, $true)
    $raw    = $reader.ReadToEnd(); $reader.Close()
    return ($raw | ConvertFrom-Json)
}

# ============================================================
# 3. CHARGEMENT CREDENTIALS JIRA & WORKSPACE ASSETS
# ============================================================
Write-Info "=== 1/3 Chargement credentials et Workspace Assets ==="

$jiraCredFile = Join-Path $secretsDir "jira-jiradot.cred.xml"
if (-not (Test-Path $jiraCredFile)) { throw "Fichier Jira creds introuvable : $jiraCredFile" }

$jiraData    = Import-Clixml -Path $jiraCredFile
$jiraBaseUrl = ([string]$jiraData.JiraBaseUrl).TrimEnd("/")
$jiraCred    = $jiraData.Credential
$b64         = [Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes($jiraCred.UserName + ":" + $jiraCred.GetNetworkCredential().Password))
$jiraHeaders = @{ Authorization = "Basic " + $b64; Accept = "application/json" }

# Récupération dynamique du Workspace ID Assets JSM
$wsUrl = $jiraBaseUrl + "/rest/servicedeskapi/assets/workspace"
$assetsWorkspaceId = "2898703d-5f14-4527-9f8a-ad32b9279ca1" # Default Fallback
try {
    $respWs = Invoke-ApiGet -Url $wsUrl -Headers $jiraHeaders
    if ($respWs.values -and $respWs.values.Count -gt 0) {
        $assetsWorkspaceId = [string]$respWs.values[0].workspaceId
    }
} catch {
    Write-Warn ("Workspace ID Assets non recupere via API, utilisation de la valeur par defaut : " + $assetsWorkspaceId)
}

$assetsAqlUrl    = $jiraBaseUrl + "/gateway/api/jsm/assets/workspace/" + $assetsWorkspaceId + "/v1/object/aql"
$assetsSchemaUrl = $jiraBaseUrl + "/gateway/api/jsm/assets/workspace/" + $assetsWorkspaceId + "/v1/objectschema/list"

Write-Info ("Connecte sur : " + $jiraBaseUrl)
Write-Info ("Workspace Assets ID : " + $assetsWorkspaceId)
Write-Log ("=== DEBUT EXECUTION Get-AssetsJiraUsers v1.1 — " + $runStamp + " ===") "INFO"

# ============================================================
# 4. RECHERCHE SCHEMA "RÉFÉRENTIEL PERSONNE" (RP)
# ============================================================
Write-Info "=== 2/3 Chargement des fiches Assets RP (Referentiel Personne) ==="

$rpSchemaId = "6"
try {
    $respSchemas = Invoke-ApiGet -Url $assetsSchemaUrl -Headers $jiraHeaders
    $sList = if ($respSchemas.values) { $respSchemas.values } else { $respSchemas.objectSchemas }
    foreach ($sch in $sList) {
        if ([string]$sch.name -ilike "*Referentiel personne*" -or
            [string]$sch.name -ilike "*Référentiel personne*" -or
            [string]$sch.name -ieq "RP") {
            $rpSchemaId = [string]$sch.id; break
        }
    }
} catch {}

Write-Info ("Schema RP identifie (ID: " + $rpSchemaId + ")")

# ============================================================
# 5. EXTRACTION ET PAGINATION AQL ASSETS
# ============================================================
$aqlQuery = "objectSchemaId = " + $rpSchemaId
$page     = 0
$maxRes   = 100
$rawObjs  = New-Object System.Collections.ArrayList

while ($true) {
    $page++
    $startAt = ($page - 1) * $maxRes
    # Correctif AQL : le champ JSON requis par l'API Gateway Assets est "qlQuery"
    $body = @{
        qlQuery           = $aqlQuery
        startAt           = $startAt
        maxResults        = $maxRes
        includeAttributes = $true
    }
    
    try {
        $resp = Invoke-ApiPostJson -Url $assetsAqlUrl -Headers $jiraHeaders -BodyObj $body
    } catch {
        Write-Warn ("Erreur lors de la requete AQL page " + $page + " : " + $_.Exception.Message)
        break
    }

    $objs = if ($resp.values) { $resp.values } else { $resp.objectEntries }
    if (-not $objs -or $objs.Count -eq 0) { break }

    foreach ($o in $objs) { [void]$rawObjs.Add($o) }
    Write-Info ("Page " + $page + " : " + $objs.Count + " objets charges (Total provisoire : " + $rawObjs.Count + ")")

    if ($resp.isLast -eq $true -or $objs.Count -lt $maxRes) { break }
}

Write-Info ("Extraction brute terminee : " + $rawObjs.Count + " fiches personnes recuperees")

# ============================================================
# 6. STRUCTURATION ET TRAITEMENT DES ATTRIBUTS
# ============================================================
$finalRows = New-Object System.Collections.ArrayList

foreach ($pObj in $rawObjs) {
    if (-not $pObj.id) { continue }

    $accId=""; $nomPrenom=""; $statut="Actif"; $typeRes="Prestataire"
    $codeSimp=""; $directionStr=""; $serviceStr=""
    $matricule=""; $societe=""; $codeCigref=""; $compteJiraRaw=""
    $attr375=""; $attr376=""

    $labelAssets = if ($pObj.label) { Fix-Encoding ([string]$pObj.label) } else { "" }
    $objTypeName = if ($pObj.objectType -and $pObj.objectType.name) { Fix-Encoding ([string]$pObj.objectType.name) } else { "" }
    $objectKey   = [string]$pObj.objectKey

    foreach ($attr in $pObj.attributes) {
        $attrId = [string]$attr.objectTypeAttributeId
        $vals   = $attr.objectAttributeValues
        if (-not $vals -or $vals.Count -eq 0) { continue }
        $vVal  = Fix-Encoding ([string]$vals[0].value)
        $vDisp = Fix-Encoding ([string]$vals[0].displayValue)
        $vBest = if (-not [string]::IsNullOrWhiteSpace($vDisp)) { $vDisp } else { $vVal }

        switch ($attrId) {
            "375" { $attr375  = $vBest }   # Attribut 375 = NOM
            "376" { $attr376  = $vBest }   # Attribut 376 = Prénom
            "407" { if ([string]::IsNullOrWhiteSpace($nomPrenom)) { $nomPrenom = $vBest } } # Fallback NOM Prénom
            "378" { $codeSimp = $vBest }   # Attribut 378 = Domaine SIMP
            "383" {
                # Attribut 383 = Société (référence objet ou valeur texte)
                foreach ($v in $vals) {
                    if ($v.referencedObject -and $v.referencedObject.label) {
                        $societe = Fix-Encoding ([string]$v.referencedObject.label); break
                    } elseif ($v.referencedObject -and $v.referencedObject.name) {
                        $societe = Fix-Encoding ([string]$v.referencedObject.name); break
                    } elseif (-not [string]::IsNullOrWhiteSpace($v.displayValue)) {
                        $societe = Fix-Encoding ([string]$v.displayValue); break
                    } elseif (-not [string]::IsNullOrWhiteSpace($v.value)) {
                        $societe = Fix-Encoding ([string]$v.value); break
                    }
                }
            }
            default {
                $aName = if ($attr.objectTypeAttribute -and $attr.objectTypeAttribute.name) {
                    Fix-Encoding ([string]$attr.objectTypeAttribute.name)
                } else { "" }

                if ($aName -ieq "Compte Jira" -or $aName -ilike "*AccountId*" -or $aName -ilike "*User*ID*") {
                    if ([string]::IsNullOrWhiteSpace($accId)) {
                        foreach ($v in $vals) {
                            if ($v.user -and $v.user.accountId) {
                                $accId = [string]$v.user.accountId; break
                            }
                        }
                    }
                    if ([string]::IsNullOrWhiteSpace($accId)) {
                        foreach ($v in $vals) {
                            $raw = [string]$v.value
                            if ($raw -match "^[0-9a-f]{24}$" -or $raw -match "^[0-9a-z]{5,12}:[0-9a-f-]{20,50}$") {
                                $accId = $raw.Trim(); break
                            }
                        }
                    }
                    if ([string]::IsNullOrWhiteSpace($accId) -and $vDisp -match "\(([0-9a-f]{24}|[0-9a-z]{5,12}:[0-9a-f-]{20,50})\)") {
                        $accId = $matches[1].Trim()
                    }
                    $compteJiraRaw = $vDisp
                }
                elseif ($aName -ieq "Statut")                            { $statut       = $vBest }
                elseif ($aName -ilike "*Type ressource*")                 { $typeRes      = $vBest }
                elseif ($aName -ieq "Matricule")                         { $matricule    = $vBest }
                elseif ($aName -ieq "Direction")                         { $directionStr = $vBest }
                elseif ($aName -ilike "*Affectation #2*")                { $serviceStr   = $vBest }
                elseif ($aName -ilike "*CIGREF*")                        { $codeCigref   = $vBest }
            }
        }
    }

    # Reconstitution NOM Prénom (attributs 375 + 376)
    if (-not [string]::IsNullOrWhiteSpace($attr375) -or -not [string]::IsNullOrWhiteSpace($attr376)) {
        $nomPrenom = ($attr375.Trim() + " " + $attr376.Trim()).Trim()
    }
    if ([string]::IsNullOrWhiteSpace($nomPrenom)) { $nomPrenom = $labelAssets }

    # Règle Métier Société : Attribut 383 prioritaire, fallback selon type ressource
    if ([string]::IsNullOrWhiteSpace($societe)) {
        if ($objTypeName -ilike "*Employe*" -or $objTypeName -ilike "*Employé*" -or $typeRes -ieq "Interne") {
            $societe = "Harmonie Mutuelle"
        } else {
            $societe = "(Société non renseignée)"
        }
    }

    # Filtrage des comptes inactifs sauf si -IncludeInactifs est spécifié
    if (-not $IncludeInactifs -and $statut -ieq "Inactif") { continue }

    $emailExtrait = ""
    if ($compteJiraRaw -match "\(([^)]+@[^)]+)\)") { $emailExtrait = $matches[1].Trim().ToLower() }

    [void]$finalRows.Add([pscustomobject]@{
        "Clé Objet"              = $objectKey
        "NOM Prénom"             = $nomPrenom
        "NOM"                    = $attr375
        "Prénom"                 = $attr376
        "Type Ressource"         = $typeRes
        "Société"                = $societe
        "Code Domaine SIMP"      = $codeSimp
        "Direction"              = $directionStr
        "Service (Affectation #2)"= $serviceStr
        "Matricule"              = $matricule
        "Code CIGREF"            = $codeCigref
        "Statut RP"              = $statut
        "Email Extrait"          = $emailExtrait
        "AccountId Jira"         = $accId
    })
}

# Tri par NOM Prénom
$finalRows = @($finalRows | Sort-Object "NOM Prénom")

# ============================================================
# 7. EXPORT CSV (FORMAT EXCEL FR)
# ============================================================
Write-Info "=== 3/3 Exportation CSV ==="

$exportCsv = Join-Path $exportsDir ("Utilisateurs_Assets_Jiradot_" + $runStamp + ".csv")

$utf8Bom = New-Object System.Text.UTF8Encoding($true)
$writer  = New-Object System.IO.StreamWriter($exportCsv, $false, $utf8Bom)
$headers = @(
    "Clé Objet","NOM Prénom","NOM","Prénom","Type Ressource","Société",
    "Code Domaine SIMP","Direction","Service (Affectation #2)","Matricule",
    "Code CIGREF","Statut RP","Email Extrait","AccountId Jira"
)

$writer.WriteLine(($headers -join ";"))

foreach ($row in $finalRows) {
    $vals = New-Object System.Collections.ArrayList
    foreach ($h in $headers) {
        $s = [string]$row.$h
        if ($s.Contains(";") -or $s.Contains('"') -or $s.Contains("`n")) {
            [void]$vals.Add('"' + $s.Replace('"','""') + '"')
        } else { [void]$vals.Add($s) }
    }
    $writer.WriteLine(($vals -join ";"))
}
$writer.Close(); $writer.Dispose()

Write-Info ("Export termine avec succes !")
Write-Info ("Nombre total de fiches personnes exportees : " + $finalRows.Count)
Write-Info ("Fichier CSV genere dans : " + $exportCsv)
Write-Log  ("=== FIN EXECUTION Get-AssetsJiraUsers v1.1 — " + (Get-Date -Format "yyyy-MM-dd HH:mm:ss") + " ===") "INFO"

# Affichage récapitulatif console (Top 20)
$finalRows | Select-Object -First 20 | Format-Table -Property "Clé Objet","NOM Prénom","Type Ressource","Société","Code Domaine SIMP","Statut RP" -AutoSize