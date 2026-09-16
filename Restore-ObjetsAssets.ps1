<#
Restore-ObjetsAssets.ps1
Restaure des objets Assets supprimes a partir des fichiers JSON de backup.
Peut restaurer un objet unique, une liste, ou tout un dossier de backup.
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [string]$BackupPath,              # Dossier backup OU fichier JSON unique — si vide, boite de dialogue

    [switch]$UseSystemProxy = $true,
    [switch]$ProxyUseDefaultCredentials = $true,
    [string]$ProxyUrl,
    [string]$AssetsWorkspaceId = "2898703d-5f14-4527-9f8a-ad32b9279ca1"
)

# ============================================================
# 0. DOSSIERS + LOG
# ============================================================
$scriptDir  = if ($PSCommandPath) { Split-Path -Parent $PSCommandPath } else { (Get-Location).Path }
$secretsDir = Join-Path $scriptDir "secrets"
$logsDir    = Join-Path $scriptDir "logs"
$exportsDir = Join-Path $scriptDir "exports"

function Ensure-Dir([string]$Path) {
    if (-not (Test-Path $Path)) { New-Item -ItemType Directory -Path $Path -Force | Out-Null }
}
Ensure-Dir $logsDir; Ensure-Dir $exportsDir

$runStamp   = Get-Date -Format "yyyyMMdd_HHmmss"
$scriptName = "Restore-ObjetsAssets"
$logFile    = Join-Path $logsDir ($scriptName + "_" + $runStamp + ".log")

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
# 1. SELECTION DU BACKUP A RESTAURER
# ============================================================
Write-Host ""
Write-Host "=============================================" -ForegroundColor Cyan
Write-Host "  Restore-ObjetsAssets"                        -ForegroundColor Cyan
Write-Host "=============================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "  [1] Restaurer TOUT un dossier de backup"    -ForegroundColor Green
Write-Host "  [2] Restaurer UN objet (fichier JSON)"      -ForegroundColor Yellow
Write-Host ""
$choixMode = Read-Host "Votre choix (1 ou 2)"

if ([string]::IsNullOrWhiteSpace($BackupPath)) {
    Add-Type -AssemblyName System.Windows.Forms

    if ($choixMode -eq "2") {
        # Selection d'un fichier JSON unique
        $dialog = New-Object System.Windows.Forms.OpenFileDialog
        $dialog.Title            = "Selectionner le fichier JSON de l'objet a restaurer"
        $dialog.Filter           = "Fichiers JSON (*.json)|*.json"
        $dialog.InitialDirectory = Join-Path $scriptDir "backup"
        $dialog.Multiselect      = $false
        $result = $dialog.ShowDialog()
        if ($result -ne [System.Windows.Forms.DialogResult]::OK) { throw "Aucun fichier selectionne." }
        $BackupPath = $dialog.FileName
    } else {
        # Selection d'un dossier
        $dialog = New-Object System.Windows.Forms.FolderBrowserDialog
        $dialog.Description  = "Selectionner le dossier Purge_{timestamp} a restaurer"
        $dialog.RootFolder   = [System.Environment+SpecialFolder]::MyComputer
        $dialog.SelectedPath = Join-Path $scriptDir "backup"
        $result = $dialog.ShowDialog()
        if ($result -ne [System.Windows.Forms.DialogResult]::OK) { throw "Aucun dossier selectionne." }
        $BackupPath = $dialog.SelectedPath
    }
}

if (-not (Test-Path $BackupPath)) { throw ("Chemin introuvable : " + $BackupPath) }
Write-Info ("Source backup : " + $BackupPath)

# Determiner les fichiers JSON a restaurer
$jsonFiles = @()
if ((Get-Item $BackupPath).PSIsContainer) {
    # Dossier : prendre tous les JSON sauf les fichiers prefixes _
    $jsonFiles = Get-ChildItem -Path $BackupPath -Filter "*.json" |
        Where-Object { -not $_.Name.StartsWith("_") } |
        Sort-Object Name
    Write-Info ("  Fichiers JSON trouves : " + $jsonFiles.Count)
} else {
    # Fichier unique
    $jsonFiles = @(Get-Item $BackupPath)
    Write-Info ("  Fichier JSON unique : " + $jsonFiles[0].Name)
}

if ($jsonFiles.Count -eq 0) {
    Write-ErrLog "Aucun fichier JSON a restaurer."
    throw "Aucun fichier JSON."
}

# ============================================================
# 2. PROXY
# ============================================================
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
    }
} catch { Write-Warn ("Init proxy: " + $_.Exception.Message) }

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
    $px = Get-EffectiveProxyUri $Url
    if ($px) { $params.Proxy=$px; if ($ProxyUseDefaultCredentials) { $params.ProxyUseDefaultCredentials=$true } }
    try {
        $resp   = Invoke-WebRequest @params
        $stream = $resp.RawContentStream; $stream.Position = 0
        $reader = New-Object System.IO.StreamReader($stream, [System.Text.Encoding]::UTF8, $true)
        $raw    = $reader.ReadToEnd(); $reader.Close()
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
if (-not (Test-Path $jiraCredFile)) { throw "Credential introuvable : " + $jiraCredFile }
$jiraData    = Import-Clixml -Path $jiraCredFile
$jiraBaseUrl = ($jiraData.JiraBaseUrl).TrimEnd("/")
$jiraCred    = $jiraData.Credential
$b64         = [Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes(
                    $jiraCred.UserName + ":" + $jiraCred.GetNetworkCredential().Password))
$jiraHeaders = @{ Authorization = "Basic " + $b64; Accept = "application/json" }
Write-Info ("Jira: " + $jiraBaseUrl)

$assetsCreateUrl = $jiraBaseUrl + "/gateway/api/jsm/assets/workspace/" + $AssetsWorkspaceId + "/v1/object/create"

# ============================================================
# 5. EXTRACTION DES ATTRIBUTS DEPUIS LE JSON DE BACKUP
# ============================================================
function Build-CreatePayload($objData) {
    # L'API create attend : { objectTypeId, attributes: [ { objectTypeAttributeId, objectAttributeValues: [...] } ] }
    $objectTypeId = $null
    if ($objData.objectType -and $objData.objectType.id) {
        $objectTypeId = [string]$objData.objectType.id
    }
    if (-not $objectTypeId) {
        return $null
    }

    $attributes = New-Object System.Collections.Generic.List[object]

    foreach ($attr in $objData.attributes) {
        $attrId = $null
        if ($attr.objectTypeAttributeId) {
            $attrId = [string]$attr.objectTypeAttributeId
        } elseif ($attr.objectTypeAttribute -and $attr.objectTypeAttribute.id) {
            $attrId = [string]$attr.objectTypeAttribute.id
        }
        if (-not $attrId) { continue }

        # Reconstruire les valeurs
        $vals = $attr.objectAttributeValues
        if (-not $vals -or $vals.Count -eq 0) { continue }

        $attrValues = New-Object System.Collections.Generic.List[object]

        foreach ($v in $vals) {
            if ($v.referencedObject -and $v.referencedObject.id) {
                # Attribut de type reference (lien vers un autre objet)
                $attrValues.Add(@{ value = [string]$v.referencedObject.id }) | Out-Null
            } elseif ($v.value -ne $null) {
                $attrValues.Add(@{ value = $v.value }) | Out-Null
            } elseif ($v.displayValue) {
                $attrValues.Add(@{ value = $v.displayValue }) | Out-Null
            }
        }

        if ($attrValues.Count -gt 0) {
            $attributes.Add(@{
                objectTypeAttributeId = $attrId
                objectAttributeValues = $attrValues.ToArray()
            }) | Out-Null
        }
    }

    return @{
        objectTypeId = $objectTypeId
        attributes   = $attributes.ToArray()
    }
}

# ============================================================
# 6. CONFIRMATION
# ============================================================
Write-Host ""
Write-Host "================================================================" -ForegroundColor Yellow
Write-Host ("  " + $jsonFiles.Count + " objet(s) vont etre recrees dans Assets.") -ForegroundColor Yellow
Write-Host "  Les objets recreeront un NOUVEL ObjectId"                          -ForegroundColor Yellow
Write-Host "  (l'ancien ObjectId/ObjectKey ne sera pas conserve)."               -ForegroundColor Yellow
Write-Host "================================================================" -ForegroundColor Yellow
Write-Host ""
$confirm = Read-Host "Tapez RESTAURER pour lancer (autre chose pour annuler)"

if ($confirm -ne "RESTAURER") {
    Write-Info "Restauration ANNULEE par l'utilisateur."
    Write-Log "Restauration annulee." "WARN"
    return
}

# ============================================================
# 7. RESTAURATION
# ============================================================
Write-Info "=== Restauration en cours ==="

$restoreOk    = 0
$restoreError = 0
$restoreIdx   = 0
$reportRows   = New-Object System.Collections.Generic.List[object]

foreach ($jsonFile in $jsonFiles) {
    $restoreIdx++
    $objData = $null

    # Lire le JSON
    try {
        $raw     = Get-Content -Path $jsonFile.FullName -Raw -Encoding UTF8
        $objData = $raw | ConvertFrom-Json
    } catch {
        Write-Warn ("  [" + $restoreIdx + "/" + $jsonFiles.Count + "] Erreur lecture JSON : " +
            $jsonFile.Name + " — " + $_.Exception.Message)
        $restoreError++
        $reportRows.Add([pscustomobject]@{
            FichierSource   = $jsonFile.Name
            AncienObjectKey = ""
            AncienObjectId  = ""
            NouvelObjectId  = ""
            NouvelObjectKey = ""
            Statut          = "ERREUR LECTURE"
            Erreur          = $_.Exception.Message
        }) | Out-Null
        continue
    }

    $ancienKey = if ($objData.objectKey) { [string]$objData.objectKey } else { "" }
    $ancienId  = if ($objData.id)        { [string]$objData.id }        else { "" }
    $label     = if ($objData.label)     { [string]$objData.label }     else { "" }

    # Construire le payload de creation
    $payload = Build-CreatePayload $objData
    if (-not $payload) {
        Write-Warn ("  [" + $restoreIdx + "/" + $jsonFiles.Count + "] " +
            $ancienKey + " : impossible de construire le payload (objectTypeId manquant)")
        $restoreError++
        $reportRows.Add([pscustomobject]@{
            FichierSource   = $jsonFile.Name
            AncienObjectKey = $ancienKey
            AncienObjectId  = $ancienId
            NouvelObjectId  = ""
            NouvelObjectKey = ""
            Statut          = "ERREUR PAYLOAD"
            Erreur          = "objectTypeId manquant"
        }) | Out-Null
        continue
    }

    $jsonBody = $payload | ConvertTo-Json -Depth 10

    # Appel API create
    $newObj = $null
    try {
        $newObj = Invoke-ApiPostUtf8 -Url $assetsCreateUrl -Headers $jiraHeaders -JsonBody $jsonBody
        $restoreOk++

        $newId  = if ($newObj.id)        { [string]$newObj.id }        else { "" }
        $newKey = if ($newObj.objectKey) { [string]$newObj.objectKey } else { "" }

        Write-Info ("  [" + $restoreIdx + "/" + $jsonFiles.Count + "] RESTAURE : " +
            $ancienKey + " -> " + $newKey + " (id=" + $newId + ")")

        $reportRows.Add([pscustomobject]@{
            FichierSource   = $jsonFile.Name
            AncienObjectKey = $ancienKey
            AncienObjectId  = $ancienId
            NouvelObjectId  = $newId
            NouvelObjectKey = $newKey
            Statut          = "RESTAURE"
            Erreur          = ""
        }) | Out-Null

    } catch {
        $restoreError++
        $errMsg = $_.Exception.Message
        Write-Warn ("  [" + $restoreIdx + "/" + $jsonFiles.Count + "] ERREUR : " +
            $ancienKey + " — " + $errMsg)

        $reportRows.Add([pscustomobject]@{
            FichierSource   = $jsonFile.Name
            AncienObjectKey = $ancienKey
            AncienObjectId  = $ancienId
            NouvelObjectId  = ""
            NouvelObjectKey = ""
            Statut          = "ERREUR CREATE"
            Erreur          = $errMsg
        }) | Out-Null
    }

    Start-Sleep -Milliseconds 200
}

# ============================================================
# 8. EXPORT RAPPORT
# ============================================================
Write-Info ""
Write-Info "=== Export rapport de restauration ==="

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

$csvReport = Join-Path $exportsDir ("Restore-Rapport_" + $runStamp + ".csv")
Export-CsvStrict -Path $csvReport `
    -Headers @("FichierSource", "AncienObjectKey", "AncienObjectId",
               "NouvelObjectKey", "NouvelObjectId", "Statut", "Erreur") `
    -Rows $reportRows

# ============================================================
# 9. RESUME FINAL
# ============================================================
Write-Info ""
Write-Info "============================================="
Write-Info "=== RESUME RESTAURATION ==="
Write-Info "============================================="
Write-Info ""
Write-Info ("  Source backup                : " + $BackupPath)
Write-Info ("  Fichiers JSON traites        : " + $jsonFiles.Count)
Write-Info ""
Write-Info ("  Objets restaures avec succes : " + $restoreOk)
Write-Info ("  Erreurs restauration         : " + $restoreError)
Write-Info ""
Write-Info ("  Rapport : " + $csvReport)
Write-Info ("  Log     : " + $logFile)

if ($restoreOk -gt 0) {
    Write-Info ""
    Write-Warn "  ATTENTION : les objets restaures ont de nouveaux ObjectKey/ObjectId."
    Write-Warn "  Consultez le rapport CSV pour la correspondance ancien -> nouveau."
}

Write-Log ("=== FIN " + $scriptName + " ===") "INFO"