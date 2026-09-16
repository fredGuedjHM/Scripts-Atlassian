<#
Purge-ObjetsAssetsObsoletes.ps1
Supprime les objets Assets identifies dans un CSV d'entree (colonne ObjectKey).
Resout ObjectKey -> ObjectId via AQL avant suppression.
Sauvegarde integrale de chaque objet AVANT suppression (JSON + CSV).
Mode DRY RUN par defaut — prompt interactif pour choisir le mode.
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [string]$CsvInput,

    [switch]$DryRun = $true,
    [switch]$UseSystemProxy = $true,
    [switch]$ProxyUseDefaultCredentials = $true,
    [string]$ProxyUrl,
    [string]$AssetsWorkspaceId = "2898703d-5f14-4527-9f8a-ad32b9279ca1"
)

# ============================================================
# 0. DOSSIERS
# ============================================================
$scriptDir  = if ($PSCommandPath) { Split-Path -Parent $PSCommandPath } else { (Get-Location).Path }
$secretsDir = Join-Path $scriptDir "secrets"
$logsDir    = Join-Path $scriptDir "logs"
$exportsDir = Join-Path $scriptDir "exports"
$backupDir  = Join-Path $scriptDir "backup"

function Ensure-Dir([string]$Path) {
    if (-not (Test-Path $Path)) { New-Item -ItemType Directory -Path $Path -Force | Out-Null }
}
Ensure-Dir $secretsDir; Ensure-Dir $logsDir; Ensure-Dir $exportsDir; Ensure-Dir $backupDir

$runStamp   = Get-Date -Format "yyyyMMdd_HHmmss"
$scriptName = "Purge-ObjetsAssetsObsoletes"

# ============================================================
# 0b. CHOIX DU MODE D'EXECUTION
# ============================================================
Write-Host ""
Write-Host "=============================================" -ForegroundColor Cyan
Write-Host "  Purge-ObjetsAssetsObsoletes"                 -ForegroundColor Cyan
Write-Host "=============================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "  [1] DRY RUN   — aucune suppression (defaut)" -ForegroundColor Green
Write-Host "  [2] EXECUTION — suppression reelle"           -ForegroundColor Red
Write-Host ""
$choix = Read-Host "Votre choix (1 ou 2, Entree = DRY RUN)"

if ($choix -eq "2") {
    $DryRun = $false
    Write-Host ""
    Write-Host "  >>> MODE EXECUTION selectionne <<<" -ForegroundColor Red
    Write-Host ""
} else {
    $DryRun = $true
    Write-Host ""
    Write-Host "  >>> MODE DRY RUN selectionne <<<" -ForegroundColor Green
    Write-Host ""
}

$mode = if ($DryRun) { "DRY RUN" } else { "PURGE REELLE" }

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

Write-Log ("=== DEBUT " + $scriptName + " [" + $mode + "] ===") "INFO"
Write-Info ("Mode : " + $mode)

# ============================================================
# 2. SELECTION DU FICHIER CSV D'ENTREE
# ============================================================
if ([string]::IsNullOrWhiteSpace($CsvInput)) {
    Write-Info "Aucun fichier CSV specifie — ouverture de la boite de dialogue..."
    Add-Type -AssemblyName System.Windows.Forms
    $dialog = New-Object System.Windows.Forms.OpenFileDialog
    $dialog.Title            = "Selectionner le CSV des objets a purger"
    $dialog.Filter           = "Fichiers CSV (*.csv)|*.csv|Tous les fichiers (*.*)|*.*"
    $dialog.InitialDirectory = $exportsDir
    $dialog.Multiselect      = $false

    $result = $dialog.ShowDialog()
    if ($result -ne [System.Windows.Forms.DialogResult]::OK -or
        [string]::IsNullOrWhiteSpace($dialog.FileName)) {
        Write-ErrLog "Aucun fichier selectionne — arret."
        throw "Aucun fichier CSV selectionne."
    }
    $CsvInput = $dialog.FileName
}

if (-not (Test-Path $CsvInput)) {
    Write-ErrLog ("Fichier introuvable : " + $CsvInput)
    throw ("Fichier introuvable : " + $CsvInput)
}
Write-Info ("CSV d'entree : " + $CsvInput)

# ============================================================
# 3. PROXY
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
# 4. HTTP WRAPPERS
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

function Invoke-ApiDelete {
    param([Parameter(Mandatory)][string]$Url,
          [Parameter(Mandatory)][hashtable]$Headers)
    $params = @{ Method='DELETE'; Uri=$Url; Headers=$Headers; ContentType='application/json'; ErrorAction='Stop' }
    $px = Get-EffectiveProxyUri -TargetUrl $Url
    if ($px) { $params.Proxy=$px; if ($ProxyUseDefaultCredentials) { $params.ProxyUseDefaultCredentials=$true } }
    try { return Invoke-RestMethod @params }
    catch [System.Net.WebException] {
        $sc   = $null
        try { $sc = $_.Exception.Response.StatusCode.value__ } catch {}
        $body = Get-WebExceptionBody $_.Exception
        Write-ErrLog ("DELETE " + $Url + " : HTTP " + $sc + " — " + $_.Exception.Message + "`n" + $body)
        throw
    }
}

# ============================================================
# 5. CREDENTIALS JIRA
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

$assetsAqlUrl  = $jiraBaseUrl + "/gateway/api/jsm/assets/workspace/" + $AssetsWorkspaceId + "/v1/object/aql"
$assetsObjBase = $jiraBaseUrl + "/gateway/api/jsm/assets/workspace/" + $AssetsWorkspaceId + "/v1/object"
# ============================================================
# 6. LECTURE DU CSV D'ENTREE
# ============================================================
Write-Info "=== Lecture du CSV d'entree ==="

$firstLine = (Get-Content -Path $CsvInput -TotalCount 1 -Encoding UTF8)
$delimiter = if ($firstLine -match ";") { ";" } else { "," }
Write-Info ("  Delimiteur detecte : '" + $delimiter + "'")

$csvRows = Import-Csv -Path $CsvInput -Delimiter $delimiter -Encoding UTF8

$objectKeyCol = $null
foreach ($candidate in @("ObjectKey", "objectKey", "objectkey", "Object Key")) {
    if ($csvRows[0].PSObject.Properties.Name -contains $candidate) {
        $objectKeyCol = $candidate
        break
    }
}
if (-not $objectKeyCol) {
    Write-ErrLog ("Colonne ObjectKey introuvable. Colonnes presentes : " +
        ($csvRows[0].PSObject.Properties.Name -join ", "))
    throw "Le CSV doit contenir une colonne ObjectKey (ex. RP-83084)."
}

$objectKeys = @($csvRows | ForEach-Object { $_.$objectKeyCol } |
    Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Select-Object -Unique)

Write-Info ("  Colonne utilisee     : " + $objectKeyCol)
Write-Info ("  ObjectKeys a traiter : " + $objectKeys.Count)

if ($objectKeys.Count -eq 0) {
    Write-ErrLog "Le CSV ne contient aucun ObjectKey valide."
    throw "CSV vide."
}

# ============================================================
# 7. RESOLUTION ObjectKey -> ObjectId via AQL
# ============================================================
Write-Info "=== Resolution ObjectKey -> ObjectId via AQL ==="

$resolvedMap   = @{}
$resolveErrors = New-Object System.Collections.Generic.List[object]

$batchSize = 25
$batches   = [System.Collections.Generic.List[string[]]]::new()
$batch     = New-Object System.Collections.Generic.List[string]

foreach ($key in $objectKeys) {
    $batch.Add($key) | Out-Null
    if ($batch.Count -ge $batchSize) {
        $batches.Add($batch.ToArray()) | Out-Null
        $batch = New-Object System.Collections.Generic.List[string]
    }
}
if ($batch.Count -gt 0) { $batches.Add($batch.ToArray()) | Out-Null }

Write-Info ("  Batches AQL : " + $batches.Count + " (taille max " + $batchSize + ")")

$batchIdx = 0
foreach ($batchKeys in $batches) {
    $batchIdx++

    $quotedKeys = $batchKeys | ForEach-Object { '"' + $_ + '"' }
    $aqlQuery   = "Key IN (" + ($quotedKeys -join ", ") + ")"

    $url      = $assetsAqlUrl + "?startAt=0&maxResults=" + $batchSize + "&includeAttributes=false"
    $bodyJson = (@{ qlQuery = $aqlQuery } | ConvertTo-Json -Depth 3)

    $resp = $null
    try {
        $resp = Invoke-ApiPostUtf8 -Url $url -Headers $jiraHeaders -JsonBody $bodyJson
    } catch {
        Write-Warn ("  Batch " + $batchIdx + " : erreur AQL — " + $_.Exception.Message)
        foreach ($k in $batchKeys) {
            $resolveErrors.Add([pscustomobject]@{
                ObjectKey = $k; Etape = "RESOLVE"
                Erreur    = "Erreur AQL batch : " + $_.Exception.Message
            }) | Out-Null
        }
        continue
    }

    if (-not $resp -or -not $resp.values) {
        Write-Warn ("  Batch " + $batchIdx + " : aucun resultat AQL")
        continue
    }

    foreach ($obj in $resp.values) {
        if (-not $obj.id -or -not $obj.objectKey) { continue }
        $resolvedMap[[string]$obj.objectKey] = [pscustomobject]@{
            ObjectId  = [string]$obj.id
            ObjectKey = [string]$obj.objectKey
            Label     = if ($obj.label) { [string]$obj.label } else { "" }
        }
    }

    Write-Info ("  Batch " + $batchIdx + "/" + $batches.Count + " : " +
        $resp.values.Count + " objets resolus")
    Start-Sleep -Milliseconds 200
}

foreach ($key in $objectKeys) {
    if (-not $resolvedMap.ContainsKey($key)) {
        Write-Warn ("  Non resolu : " + $key)
        $resolveErrors.Add([pscustomobject]@{
            ObjectKey = $key; Etape = "RESOLVE"
            Erreur    = "ObjectKey introuvable dans Assets"
        }) | Out-Null
    }
}

Write-Info ("  Resolus avec succes : " + $resolvedMap.Count + "/" + $objectKeys.Count)
Write-Info ("  Non resolus         : " + $resolveErrors.Count)

# ============================================================
# 8. SAUVEGARDE INTEGRALE AVANT SUPPRESSION
# ============================================================
Write-Info "=== Sauvegarde integrale des objets (backup) ==="

$backupSubDir = Join-Path $backupDir ("Purge_" + $runStamp)
Ensure-Dir $backupSubDir

$backupObjects  = New-Object System.Collections.Generic.List[object]
$allErrors      = New-Object System.Collections.Generic.List[object]

foreach ($e in $resolveErrors) { $allErrors.Add($e) | Out-Null }

$objBackedUp    = 0
$objBackupError = 0
$objIdx         = 0

foreach ($key in $resolvedMap.Keys) {
    $resolved = $resolvedMap[$key]
    $objId    = $resolved.ObjectId
    $objIdx++

    if ($objIdx % 25 -eq 0) {
        Write-Info ("  Backup : " + $objIdx + "/" + $resolvedMap.Count)
    }

    $url     = $assetsObjBase + "/" + $objId + "?includeExtendedInfo=true"
    $objData = $null
    try {
        $objData = Invoke-ApiGet -Url $url -Headers $jiraHeaders
    } catch {
        $errMsg = $_.Exception.Message
        Write-Warn ("  " + $key + " (id=" + $objId + ") : backup impossible — " + $errMsg)
        $objBackupError++
        $allErrors.Add([pscustomobject]@{
            ObjectKey = $key; ObjectId = $objId
            Etape     = "BACKUP"; Erreur = $errMsg
        }) | Out-Null
        continue
    }

    $jsonFile = Join-Path $backupSubDir ($objId + "_" + $key + ".json")
    try {
        $objData | ConvertTo-Json -Depth 20 |
            Out-File -FilePath $jsonFile -Encoding UTF8 -Force
    } catch {
        Write-Warn ("  " + $key + " : erreur ecriture JSON — " + $_.Exception.Message)
    }

    $backupObjects.Add($objData) | Out-Null
    $objBackedUp++
    Start-Sleep -Milliseconds 100
}

Write-Info ("  Objets sauvegardes : " + $objBackedUp)
Write-Info ("  Erreurs backup     : " + $objBackupError)

$globalBackup = Join-Path $backupSubDir ("_BACKUP_COMPLET_" + $runStamp + ".json")
try {
    $backupObjects | ConvertTo-Json -Depth 20 |
        Out-File -FilePath $globalBackup -Encoding UTF8 -Force
    Write-Info ("  Backup global : " + $globalBackup)
} catch {
    Write-ErrLog ("  Erreur ecriture backup global : " + $_.Exception.Message)
}

Copy-Item -Path $CsvInput -Destination (Join-Path $backupSubDir ("_CSV_INPUT_" + $runStamp + ".csv")) -Force

$idsToDelete = New-Object System.Collections.Generic.List[object]
foreach ($obj in $backupObjects) {
    $objId  = if ($obj.id)        { [string]$obj.id }        else { "" }
    $objKey = if ($obj.objectKey) { [string]$obj.objectKey } else { "" }
    if (-not [string]::IsNullOrWhiteSpace($objId)) {
        $idsToDelete.Add([pscustomobject]@{
            ObjectId  = $objId
            ObjectKey = $objKey
            Label     = if ($obj.label) { [string]$obj.label } else { "" }
        }) | Out-Null
    }
}

Write-Info ("  Objets eligibles suppression (backup OK) : " + $idsToDelete.Count)

if ($objBackupError -gt 0 -or $resolveErrors.Count -gt 0) {
    Write-Warn ""
    Write-Warn "============================================="
    Write-Warn ("  " + ($objBackupError + $resolveErrors.Count) + " objet(s) exclus de la purge")
    Write-Warn "  (non resolus ou backup en erreur)"
    Write-Warn "  Consulter le rapport d'erreurs pour le detail."
    Write-Warn "============================================="
    Write-Warn ""
}
# ============================================================
# 9. PURGE
# ============================================================
$deleteOk    = 0
$deleteError = 0
$confirm     = ""

if ($DryRun) {
    Write-Info ""
    Write-Info "============================================="
    Write-Info "=== MODE DRY RUN — AUCUNE SUPPRESSION ==="
    Write-Info "============================================="
    Write-Info ("  Objets qui SERAIENT supprimes : " + $idsToDelete.Count)
    Write-Info ("  Backup disponible dans        : " + $backupSubDir)
    Write-Info ""

} else {
    Write-Info ""
    Write-Info "============================================="
    Write-Info "=== PURGE REELLE EN COURS ==="
    Write-Info "============================================="
    Write-Info ""

    Write-Host ""
    Write-Host "================================================================" -ForegroundColor Red
    Write-Host ("  ATTENTION : " + $idsToDelete.Count + " objets vont etre DEFINITIVEMENT supprimes.") -ForegroundColor Red
    Write-Host ("  Backup disponible dans : " + $backupSubDir) -ForegroundColor Yellow
    Write-Host "================================================================" -ForegroundColor Red
    Write-Host ""
    $confirm = Read-Host "Tapez CONFIRMER pour executer la purge (autre chose pour annuler)"

    if ($confirm -ne "CONFIRMER") {
        Write-Info "Purge ANNULEE par l'utilisateur."
    } else {
        Write-Info "Confirmation recue — demarrage de la purge..."
        $deleteIdx = 0

        foreach ($item in $idsToDelete) {
            $deleteIdx++
            $url = $assetsObjBase + "/" + $item.ObjectId

            try {
                Invoke-ApiDelete -Url $url -Headers $jiraHeaders
                $deleteOk++
                Write-Info ("  [" + $deleteIdx + "/" + $idsToDelete.Count + "]" +
                    " SUPPRIME : " + $item.ObjectKey + " (id=" + $item.ObjectId + ")")
            } catch {
                $deleteError++
                $errMsg = $_.Exception.Message
                Write-Warn ("  [" + $deleteIdx + "/" + $idsToDelete.Count + "]" +
                    " ERREUR   : " + $item.ObjectKey + " (id=" + $item.ObjectId + ") — " + $errMsg)
                $allErrors.Add([pscustomobject]@{
                    ObjectKey = $item.ObjectKey; ObjectId = $item.ObjectId
                    Etape     = "DELETE"; Erreur = $errMsg
                }) | Out-Null
            }

            Start-Sleep -Milliseconds 200
        }

        Write-Info ""
        Write-Info ("  Supprimes avec succes : " + $deleteOk)
        Write-Info ("  Erreurs suppression   : " + $deleteError)
    }
}

# ============================================================
# 10. EXPORT RAPPORTS
# ============================================================
Write-Info ""
Write-Info "=== Export rapports ==="

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

# 10a. Rapport backup
$backupReportRows = New-Object System.Collections.Generic.List[object]
foreach ($obj in $backupObjects) {
    $objId  = if ($obj.id)        { [string]$obj.id }        else { "" }
    $objKey = if ($obj.objectKey) { [string]$obj.objectKey } else { "" }
    $label  = if ($obj.label)     { [string]$obj.label }     else { "" }
    $backupReportRows.Add([pscustomobject]@{
        ObjectKey = $objKey; ObjectId = $objId; Label = $label
        Statut    = "SAUVEGARDE OK"
        JsonFile  = ($objId + "_" + $objKey + ".json")
    }) | Out-Null
}
$csvBackupReport = Join-Path $exportsDir ("Purge-RapportBackup_" + $runStamp + ".csv")
Export-CsvStrict -Path $csvBackupReport `
    -Headers @("ObjectKey", "ObjectId", "Label", "Statut", "JsonFile") `
    -Rows ($backupReportRows | Sort-Object ObjectKey)

# 10b. Rapport d'erreurs
if ($allErrors.Count -gt 0) {
    $csvErrors = Join-Path $exportsDir ("Purge-Erreurs_" + $runStamp + ".csv")
    Export-CsvStrict -Path $csvErrors `
        -Headers @("ObjectKey", "ObjectId", "Etape", "Erreur") `
        -Rows $allErrors
}

# 10c. Synthese globale
$synthRows = New-Object System.Collections.Generic.List[object]
$synthRows.Add([pscustomobject]@{ Metrique = "Mode d'execution";                Valeur = $mode })                         | Out-Null
$synthRows.Add([pscustomobject]@{ Metrique = "CSV d'entree";                    Valeur = $CsvInput })                     | Out-Null
$synthRows.Add([pscustomobject]@{ Metrique = "ObjectKeys dans le CSV";          Valeur = [string]$objectKeys.Count })     | Out-Null
$synthRows.Add([pscustomobject]@{ Metrique = "ObjectKeys resolus (AQL)";        Valeur = [string]$resolvedMap.Count })    | Out-Null
$synthRows.Add([pscustomobject]@{ Metrique = "ObjectKeys non resolus";          Valeur = [string]$resolveErrors.Count })  | Out-Null
$synthRows.Add([pscustomobject]@{ Metrique = "Objets sauvegardes (backup OK)"; Valeur = [string]$objBackedUp })          | Out-Null
$synthRows.Add([pscustomobject]@{ Metrique = "Erreurs backup";                  Valeur = [string]$objBackupError })       | Out-Null
$synthRows.Add([pscustomobject]@{ Metrique = "Dossier backup";                  Valeur = $backupSubDir })                 | Out-Null
$synthRows.Add([pscustomobject]@{ Metrique = "Backup global JSON";              Valeur = $globalBackup })                 | Out-Null

if ($DryRun) {
    $synthRows.Add([pscustomobject]@{ Metrique = "Suppression";                 Valeur = "NON EXECUTEE (DRY RUN)" })      | Out-Null
    $synthRows.Add([pscustomobject]@{ Metrique = "Objets qui seraient supprimes"; Valeur = [string]$idsToDelete.Count })  | Out-Null
} else {
    if ($confirm -eq "CONFIRMER") {
        $synthRows.Add([pscustomobject]@{ Metrique = "Objets supprimes avec succes"; Valeur = [string]$deleteOk })        | Out-Null
        $synthRows.Add([pscustomobject]@{ Metrique = "Erreurs suppression";          Valeur = [string]$deleteError })     | Out-Null
    } else {
        $synthRows.Add([pscustomobject]@{ Metrique = "Suppression";                  Valeur = "ANNULEE par l'utilisateur" }) | Out-Null
    }
}

$csvSynth = Join-Path $exportsDir ("Purge-Synthese_" + $runStamp + ".csv")
Export-CsvStrict -Path $csvSynth -Headers @("Metrique", "Valeur") -Rows $synthRows

# ============================================================
# 11. RESUME FINAL
# ============================================================
Write-Info ""
Write-Info "============================================="
Write-Info ("=== RESUME FINAL [" + $mode + "] ===")
Write-Info "============================================="
Write-Info ""
Write-Info ("  CSV d'entree                          : " + $CsvInput)
Write-Info ("  ObjectKeys dans le CSV                : " + $objectKeys.Count)
Write-Info ("  ObjectKeys resolus en ObjectId        : " + $resolvedMap.Count)
Write-Info ("  ObjectKeys non resolus                : " + $resolveErrors.Count)
Write-Info ""
Write-Info ("  Objets sauvegardes (backup)           : " + $objBackedUp)
Write-Info ("  Erreurs backup                        : " + $objBackupError)
Write-Info ("  Dossier backup                        : " + $backupSubDir)
Write-Info ""

if ($DryRun) {
    Write-Info "  >>> MODE DRY RUN — AUCUNE SUPPRESSION EFFECTUEE <<<"
    Write-Info ("  >>> Objets qui seraient supprimes : " + $idsToDelete.Count)
} else {
    if ($confirm -eq "CONFIRMER") {
        Write-Info ("  Objets supprimes avec succes          : " + $deleteOk)
        Write-Info ("  Erreurs suppression                   : " + $deleteError)
    } else {
        Write-Info "  >>> PURGE ANNULEE PAR L'UTILISATEUR <<<"
    }
}

Write-Info ""
Write-Info ("  Exports : " + $exportsDir)
Write-Info ("  Backup  : " + $backupSubDir)
Write-Info ("  Log     : " + $logFile)
Write-Log ("=== FIN " + $scriptName + " [" + $mode + "] ===") "INFO"
