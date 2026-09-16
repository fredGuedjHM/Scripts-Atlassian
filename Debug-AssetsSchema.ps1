<#
Debug-AssetsSchema.ps1
1. Liste les schémas Assets JSM pour obtenir leurs IDs exacts (RP, RS, EXCOMP)
2. Dumpe les attributs RAW des schémas RP, RS, EXCOMP avec objectSchemaId = <ID>
3. Teste la résolution de l'issue ID Tempo (255916) -> Clé Jira / Parent via JQL
#>

[CmdletBinding()]
param(
    [string]$AssetsWorkspaceId = "2898703d-5f14-4527-9f8a-ad32b9279ca1",
    [switch]$UseSystemProxy    = $true,
    [switch]$ProxyUseDefaultCredentials = $true,
    [string]$ProxyUrl,
    [int]$MaxObjectsPerSchema  = 3
)

$scriptDir  = if ($PSCommandPath) { Split-Path -Parent $PSCommandPath } else { (Get-Location).Path }
$secretsDir = Join-Path $scriptDir "secrets"
$debugDir   = Join-Path $scriptDir "exports\debug"
if (-not (Test-Path $debugDir)) { New-Item -ItemType Directory -Path $debugDir -Force | Out-Null }
$runStamp   = Get-Date -Format "yyyyMMdd_HHmmss"

# ---- Auth Jira ----
$jiraData    = Import-Clixml -Path (Join-Path $secretsDir "jira-jiradot.cred.xml")
$jiraBaseUrl = ($jiraData.JiraBaseUrl).TrimEnd("/")
$jiraCred    = $jiraData.Credential
$b64         = [Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes($jiraCred.UserName + ":" + $jiraCred.GetNetworkCredential().Password))
$jiraHeaders = @{ Authorization = "Basic " + $b64; Accept = "application/json" }
$assetsAqlUrl    = $jiraBaseUrl + "/gateway/api/jsm/assets/workspace/" + $AssetsWorkspaceId + "/v1/object/aql"
$assetsSchemaUrl = $jiraBaseUrl + "/gateway/api/jsm/assets/workspace/" + $AssetsWorkspaceId + "/v1/objectschema/list"

# ---- Proxy ----
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
if ($UseSystemProxy) {
    [System.Net.WebRequest]::DefaultWebProxy = [System.Net.WebRequest]::GetSystemWebProxy()
    [System.Net.WebRequest]::DefaultWebProxy.Credentials = [System.Net.CredentialCache]::DefaultCredentials
}

function Get-EffectiveProxyUri {
    param([string]$TargetUrl)
    if ($ProxyUrl) { return $ProxyUrl }
    if (-not $UseSystemProxy) { return $null }
    try {
        $dest = [uri]$TargetUrl
        $wp   = [System.Net.WebRequest]::DefaultWebProxy
        if (-not $wp -or $wp.IsBypassed($dest)) { return $null }
        $px = $wp.GetProxy($dest)
        if (-not $px -or $px.AbsoluteUri -eq $dest.AbsoluteUri) { return $null }
        return $px.AbsoluteUri
    } catch { return $null }
}

function Invoke-ApiGet {
    param([string]$Url, [hashtable]$Headers)
    $params = @{ Method='GET'; Uri=$Url; Headers=$Headers; ContentType='application/json'; UseBasicParsing=$true; ErrorAction='Stop' }
    $px = Get-EffectiveProxyUri -TargetUrl $Url
    if ($px) { $params.Proxy=$px; if ($ProxyUseDefaultCredentials) { $params.ProxyUseDefaultCredentials=$true } }
    $resp   = Invoke-WebRequest @params
    $stream = $resp.RawContentStream; $stream.Position = 0
    $reader = New-Object System.IO.StreamReader($stream, [System.Text.Encoding]::UTF8, $true)
    $raw    = $reader.ReadToEnd(); $reader.Close()
    return $raw | ConvertFrom-Json
}

function Invoke-ApiPostUtf8 {
    param([string]$Url, [hashtable]$Headers, [string]$JsonBody)
    $params = @{ Method='POST'; Uri=$Url; Headers=$Headers; Body=[System.Text.Encoding]::UTF8.GetBytes($JsonBody); ContentType="application/json"; UseBasicParsing=$true; ErrorAction='Stop' }
    $px = Get-EffectiveProxyUri -TargetUrl $Url
    if ($px) { $params.Proxy=$px; if ($ProxyUseDefaultCredentials) { $params.ProxyUseDefaultCredentials=$true } }
    $resp   = Invoke-WebRequest @params
    $stream = $resp.RawContentStream; $stream.Position = 0
    $reader = New-Object System.IO.StreamReader($stream, [System.Text.Encoding]::UTF8, $true)
    $raw    = $reader.ReadToEnd(); $reader.Close()
    return $raw | ConvertFrom-Json
}

function Export-CsvStrict {
    param([string]$Path, [string[]]$Headers, $Rows)
    $sb = New-Object System.Text.StringBuilder
    [void]$sb.AppendLine($Headers -join ";")
    foreach ($row in $Rows) {
        $vals = foreach ($h in $Headers) {
            $s = [string]$row.$h
            if ($s.Contains(";") -or $s.Contains('"') -or $s.Contains("`n")) { '"' + $s.Replace('"','""') + '"' } else { $s }
        }
        [void]$sb.AppendLine($vals -join ";")
    }
    $utf8Bom = New-Object System.Text.UTF8Encoding($true)
    [System.IO.File]::WriteAllText($Path, $sb.ToString(), $utf8Bom)
    Write-Host ("[INFO] -> CSV : " + $Path + " (" + $Rows.Count + " lignes)") -ForegroundColor Green
}

# ============================================================
# 1. LISTE ET RÉSOLUTION DES SCHÉMAS ASSETS
# ============================================================
Write-Host "`n[INFO] === 1/3 LISTE DES SCHÉMAS ASSETS ===" -ForegroundColor Cyan
$schemasList = @()
try {
    $respSchemas = Invoke-ApiGet -Url $assetsSchemaUrl -Headers $jiraHeaders
    $schemasList = if ($respSchemas.values) { $respSchemas.values } else { $respSchemas.objectSchemas }
} catch {
    Write-Warning ("Impossible de lister les schémas Assets : " + $_.Exception.Message)
}

$schemaMap = @{}
foreach ($s in $schemasList) {
    Write-Host ("  Schema ID=" + $s.id + "  Name='" + $s.name + "'") -ForegroundColor Yellow
    $schemaMap[[string]$s.name] = [string]$s.id
}

# ============================================================
# 2. DUMP DE DÉBOGAGE DES SCHÉMAS
# ============================================================
function Dump-SchemaById {
    param([string]$SchemaName, [string]$SchemaId, [string]$OutputCsv)

    Write-Host ("`n[INFO] === DUMP SCHEMA : " + $SchemaName + " (ID=" + $SchemaId + ") ===") -ForegroundColor Cyan
    $rows    = New-Object System.Collections.Generic.List[object]
    $urlAql  = $assetsAqlUrl + "?startAt=0&maxResults=" + $MaxObjectsPerSchema + "&includeAttributes=true"
    $aql     = "objectSchemaId = " + $SchemaId
    $bodyAql = (@{ qlQuery = $aql } | ConvertTo-Json -Depth 3)
    $respAql = $null

    try { $respAql = Invoke-ApiPostUtf8 -Url $urlAql -Headers $jiraHeaders -JsonBody $bodyAql }
    catch { Write-Warning ("Erreur AQL " + $SchemaName + " : " + $_.Exception.Message); return }

    if (-not $respAql -or -not $respAql.values) {
        Write-Warning ("Aucun objet pour le schema : " + $SchemaName)
        return
    }

    Write-Host ("[INFO]   Objets trouves : " + $respAql.values.Count + " (total=" + $respAql.total + ")") -ForegroundColor Green

    $attrDict = @{}
    if ($respAql.objectTypeAttributes) {
        foreach ($ota in $respAql.objectTypeAttributes) {
            if ($ota.id -and $ota.name) { $attrDict[[string]$ota.id] = [string]$ota.name }
        }
    }

    foreach ($pObj in $respAql.values) {
        $objId    = [string]$pObj.id
        $objLabel = [string]$pObj.label
        $objType  = if ($pObj.objectType -and $pObj.objectType.name) { [string]$pObj.objectType.name } else { "" }

        Write-Host ("`n  Objet id=" + $objId + "  label='" + $objLabel + "'  type='" + $objType + "'") -ForegroundColor White

        foreach ($attr in $pObj.attributes) {
            $attrId = [string]$attr.objectTypeAttributeId
            $aName  = $attrDict[$attrId]
            if (-not $aName -and $attr.objectTypeAttribute -and $attr.objectTypeAttribute.name) {
                $aName = [string]$attr.objectTypeAttribute.name
            }
            if (-not $aName) { $aName = "(id=" + $attrId + ")" }

            $vals = $attr.objectAttributeValues
            if (-not $vals -or $vals.Count -eq 0) { continue }

            foreach ($v in $vals) {
                $vVal  = [string]$v.value
                $vDisp = [string]$v.displayValue
                $vRef  = if ($v.referencedObject) { "REF_ID=" + [string]$v.referencedObject.id + " / " + [string]$v.referencedObject.label } else { "" }

                Write-Host ("    [" + $aName + "] val='" + $vVal + "'  disp='" + $vDisp + "'" + $(if ($vRef) { "  " + $vRef } else { "" })) -ForegroundColor Cyan

                $rows.Add([pscustomobject]@{
                    Schema       = $SchemaName
                    ObjectId     = $objId
                    ObjectLabel  = $objLabel
                    ObjectType   = $objType
                    AttrId       = $attrId
                    AttrName     = $aName
                    Value        = $vVal
                    DisplayValue = $vDisp
                    RefObject    = $vRef
                }) | Out-Null
            }
        }
    }

    Export-CsvStrict -Path $OutputCsv `
        -Headers @("Schema","ObjectId","ObjectLabel","ObjectType","AttrId","AttrName","Value","DisplayValue","RefObject") `
        -Rows $rows
}

foreach ($sName in @("RP","RS","EXCOMP")) {
    $matchedKey = $schemaMap.Keys | Where-Object { $_ -ieq $sName -or $_ -ilike ("*" + $sName + "*") } | Select-Object -First 1
    if ($matchedKey) {
        $sId = $schemaMap[$matchedKey]
        Dump-SchemaById -SchemaName $matchedKey -SchemaId $sId -OutputCsv (Join-Path $debugDir ("Debug_Schema_" + $sName + "_" + $runStamp + ".csv"))
    } else {
        Write-Warning ("Schema non trouve dans la liste : " + $sName)
    }
}

# ============================================================
# 3. TEST DE RÉSOLUTION TEMPO ISSUE.ID (255916) -> JIRA KEY
# ============================================================
Write-Host "`n[INFO] === 3/3 TEST RESOLUTION TEMPO ISSUE.ID -> JIRA KEY ===" -ForegroundColor Cyan
$sampleIssueId = "255916"
$jqlTest       = "id = " + $sampleIssueId
$urlJql        = $jiraBaseUrl + "/rest/api/3/search/jql?jql=" + [uri]::EscapeDataString($jqlTest) + "&fields=id,key,summary,parent,customfield_10183,customfield_10124"

try {
    $respJql = Invoke-ApiGet -Url $urlJql -Headers $jiraHeaders
    if ($respJql.issues -and $respJql.issues.Count -gt 0) {
        $iss = $respJql.issues[0]
        Write-Host ("  Issue ID      : " + $iss.id) -ForegroundColor Green
        Write-Host ("  Issue Key     : " + $iss.key) -ForegroundColor Green
        Write-Host ("  Summary       : " + $iss.fields.summary) -ForegroundColor Green
        Write-Host ("  Parent        : " + ($iss.fields.parent | ConvertTo-Json -Compress)) -ForegroundColor Yellow
        Write-Host ("  Budget Link   : " + $iss.fields.customfield_10183) -ForegroundColor Yellow
    } else {
        Write-Warning ("Aucun ticket Jira trouve pour ID=" + $sampleIssueId)
    }
} catch {
    Write-Warning ("Erreur recherche Jira JQL pour issue ID " + $sampleIssueId + " : " + $_.Exception.Message)
}

Write-Host "`n[INFO] === DIAGNOSTIC TERMINÉ ===" -ForegroundColor Green