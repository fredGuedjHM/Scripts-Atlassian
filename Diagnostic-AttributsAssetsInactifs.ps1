<#
Diagnostic-AttributsAssetsInactifs.ps1
Identifie les noms exacts des attributs des objets Assets Inactifs.
Objectif : trouver le nom du champ "Date Sortie" tel qu'il existe dans le schema.
#>

[CmdletBinding()]
param(
    [switch]$UseSystemProxy = $true,
    [switch]$ProxyUseDefaultCredentials = $true,
    [string]$ProxyUrl,
    [string]$AssetsWorkspaceId = "2898703d-5f14-4527-9f8a-ad32b9279ca1",
    [string[]]$ObjectTypeIds   = @("68", "69"),
    [int]$NbEchantillon        = 3   # Nb d'objets Inactifs a inspecter par OT
)

# ============================================================
# 0. DOSSIERS + LOG
# ============================================================
$scriptDir = if ($PSCommandPath) { Split-Path -Parent $PSCommandPath } else { (Get-Location).Path }
$logsDir   = Join-Path $scriptDir "logs"
if (-not (Test-Path $logsDir)) { New-Item -ItemType Directory -Path $logsDir -Force | Out-Null }

$runStamp = Get-Date -Format "yyyyMMdd_HHmmss"
$logFile  = Join-Path $logsDir ("Diagnostic-AttributsAssetsInactifs_" + $runStamp + ".log")

function Write-Log {
    param([string]$Message, [string]$Level = "INFO")
    $ts = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    try { Add-Content -Path $logFile -Value ("[$ts] [$Level] " + $Message) -ErrorAction Stop } catch {}
}
function Write-Info($msg) { Write-Host ("[INFO] " + $msg); Write-Log $msg "INFO"  }
function Write-Warn($msg) { Write-Warning $msg;             Write-Log $msg "WARN"  }
function Write-Err($msg)  { Write-Error $msg;               Write-Log $msg "ERROR" }

Write-Info "=== Diagnostic-AttributsAssetsInactifs ==="
Write-Info ("Log : " + $logFile)

# ============================================================
# 1. PROXY
# ============================================================
try {
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    if ($ProxyUrl) {
        [System.Net.WebRequest]::DefaultWebProxy = New-Object System.Net.WebProxy($ProxyUrl, $true)
        Write-Info ("Proxy explicite : " + $ProxyUrl)
    } elseif ($UseSystemProxy) {
        [System.Net.WebRequest]::DefaultWebProxy = [System.Net.WebRequest]::GetSystemWebProxy()
        [System.Net.WebRequest]::DefaultWebProxy.Credentials = [System.Net.CredentialCache]::DefaultCredentials
        Write-Info "Proxy : systeme"
    } else {
        [System.Net.WebRequest]::DefaultWebProxy = $null
        Write-Info "Proxy : desactive"
    }
} catch { Write-Warn ("Init proxy : " + $_.Exception.Message) }

function Get-EffectiveProxyUri([string]$TargetUrl) {
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
# 2. HTTP WRAPPERS
# ============================================================
function Invoke-ApiGet {
    param([string]$Url, [hashtable]$Headers)
    $params = @{ Method='GET'; Uri=$Url; Headers=$Headers; ContentType='application/json'; ErrorAction='Stop' }
    $px = Get-EffectiveProxyUri $Url
    if ($px) { $params.Proxy=$px; if ($ProxyUseDefaultCredentials) { $params.ProxyUseDefaultCredentials=$true } }
    return Invoke-RestMethod @params
}

function Invoke-ApiPostUtf8 {
    param([string]$Url, [hashtable]$Headers, [string]$JsonBody)
    $params = @{
        Method='POST'; Uri=$Url
        Headers=$Headers
        Body=[System.Text.Encoding]::UTF8.GetBytes($JsonBody)
        ContentType="application/json"
        UseBasicParsing=$true
        ErrorAction='Stop'
    }
    $px = Get-EffectiveProxyUri $Url
    if ($px) { $params.Proxy=$px; if ($ProxyUseDefaultCredentials) { $params.ProxyUseDefaultCredentials=$true } }
    $resp   = Invoke-WebRequest @params
    $stream = $resp.RawContentStream; $stream.Position = 0
    $reader = New-Object System.IO.StreamReader($stream, [System.Text.Encoding]::UTF8, $true)
    $raw    = $reader.ReadToEnd(); $reader.Close()
    return $raw | ConvertFrom-Json
}

# ============================================================
# 3. CREDENTIALS
# ============================================================
$jiraCredFile = Join-Path $scriptDir "secrets/jira-jiradot.cred.xml"
if (-not (Test-Path $jiraCredFile)) { throw "Credential introuvable : " + $jiraCredFile }
$jiraData    = Import-Clixml -Path $jiraCredFile
$jiraBaseUrl = ($jiraData.JiraBaseUrl).TrimEnd("/")
$jiraCred    = $jiraData.Credential
$b64         = [Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes(
                    $jiraCred.UserName + ":" + $jiraCred.GetNetworkCredential().Password))
$jiraHeaders = @{ Authorization = "Basic " + $b64; Accept = "application/json" }
Write-Info ("Jira : " + $jiraBaseUrl)

$assetsAqlUrl  = $jiraBaseUrl + "/gateway/api/jsm/assets/workspace/" + $AssetsWorkspaceId + "/v1/object/aql"
$assetsObjBase = $jiraBaseUrl + "/gateway/api/jsm/assets/workspace/" + $AssetsWorkspaceId + "/v1/object"

# ============================================================
# 4. DIAGNOSTIC PAR OBJECT TYPE
# ============================================================
foreach ($otId in $ObjectTypeIds) {
    Write-Info ""
    Write-Info ("============================================================")
    Write-Info ("=== Object Type : " + $otId + " ===")
    Write-Info ("============================================================")

    # Recuperer un echantillon d'objets avec attributs
    $aqlQuery = "objectTypeId = " + $otId
    $url      = $assetsAqlUrl + "?startAt=0&maxResults=100&includeAttributes=true"
    $bodyJson = (@{ qlQuery = $aqlQuery } | ConvertTo-Json -Depth 3)

    $resp = $null
    try { $resp = Invoke-ApiPostUtf8 -Url $url -Headers $jiraHeaders -JsonBody $bodyJson }
    catch { Write-Err ("AQL OT=" + $otId + " : " + $_.Exception.Message); continue }
    if (-not $resp -or -not $resp.values) { Write-Warn "Aucun objet"; continue }

    # Dictionnaire attributId -> nom
    $attrDict = @{}
    if ($resp.objectTypeAttributes) {
        foreach ($ota in $resp.objectTypeAttributes) {
            if ($ota.id -and $ota.name) { $attrDict[[string]$ota.id] = [string]$ota.name }
        }
    }
    Write-Info ("  Attributs connus via objectTypeAttributes : " + $attrDict.Count)
    foreach ($k in ($attrDict.Keys | Sort-Object)) {
        Write-Info ("    [" + $k + "] => '" + $attrDict[$k] + "'")
    }

    # Identifier les objets Inactifs dans l'echantillon
    $inactifsEch = New-Object System.Collections.Generic.List[object]
    foreach ($obj in $resp.values) {
        $statut = ""
        foreach ($attr in $obj.attributes) {
            $attrName = ""
            if ($attr.objectTypeAttributeId -and $attrDict.ContainsKey([string]$attr.objectTypeAttributeId)) {
                $attrName = $attrDict[[string]$attr.objectTypeAttributeId]
            } elseif ($attr.objectTypeAttribute -and $attr.objectTypeAttribute.name) {
                $attrName = [string]$attr.objectTypeAttribute.name
            }
            if ($attrName -eq "Statut") {
                $vals = $attr.objectAttributeValues
                if ($vals -and $vals.Count -gt 0) {
                    if ($vals[0].displayValue) { $statut = [string]$vals[0].displayValue }
                    elseif ($vals[0].value)    { $statut = [string]$vals[0].value }
                }
            }
        }
        if ($statut -eq "Inactif" -or $statut -eq "inactif") {
            $inactifsEch.Add($obj) | Out-Null
            if ($inactifsEch.Count -ge $NbEchantillon) { break }
        }
    }

    Write-Info ("  Objets Inactifs trouves dans l'echantillon : " + $inactifsEch.Count)

    if ($inactifsEch.Count -eq 0) {
        Write-Warn "  Aucun objet Inactif dans les 100 premiers — augmenter l'echantillon AQL si necessaire"
        continue
    }

    # Pour chaque objet inactif : dump complet des attributs
    foreach ($obj in $inactifsEch) {
        Write-Info ""
        Write-Info ("  --- Objet : " + $obj.objectKey + " / " + $obj.label + " (id=" + $obj.id + ") ---")

        # Appel direct sur l'objet pour avoir tous les attributs
        $objDetail = $null
        try { $objDetail = Invoke-ApiGet -Url ($assetsObjBase + "/" + $obj.id) -Headers $jiraHeaders }
        catch { Write-Warn ("  GET objet " + $obj.id + " : " + $_.Exception.Message) }

        $source = if ($objDetail -and $objDetail.attributes) { $objDetail } else { $obj }

        foreach ($attr in $source.attributes) {
            $attrName = ""
            if ($attr.objectTypeAttributeId -and $attrDict.ContainsKey([string]$attr.objectTypeAttributeId)) {
                $attrName = $attrDict[[string]$attr.objectTypeAttributeId]
            } elseif ($attr.objectTypeAttribute -and $attr.objectTypeAttribute.name) {
                $attrName = [string]$attr.objectTypeAttribute.name
            }
            if ([string]::IsNullOrWhiteSpace($attrName)) {
                $attrName = "(id=" + $attr.objectTypeAttributeId + ")"
            }

            $vals = $attr.objectAttributeValues
            if (-not $vals -or $vals.Count -eq 0) {
                Write-Info ("    '" + $attrName + "' = (vide)")
                continue
            }
            foreach ($v in $vals) {
                $display = ""
                if     ($v.displayValue) { $display = "[displayValue] " + [string]$v.displayValue }
                elseif ($v.value)        { $display = "[value] "        + [string]$v.value }
                elseif ($v.referencedObject -and $v.referencedObject.label) {
                                           $display = "[ref] "          + [string]$v.referencedObject.label }
                else                     { $display = "(format inconnu)" }
                Write-Info ("    '" + $attrName + "' = " + $display)
            }
        }
        Start-Sleep -Milliseconds 200
    }
}

Write-Info ""
Write-Info "=== FIN DIAGNOSTIC ==="
Write-Info ("Log complet : " + $logFile)