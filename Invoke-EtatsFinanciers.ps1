<#
Invoke-EtatsFinanciers.ps1 — Partie 1 / 3
Script unifié d'extraction des États Financiers (SIMP, DÉTERMINÉ, SUIVI-ACTIVITE & SIMP XLSX)

v2.5 - Nouveautés :
  1. Ajout de barres de progression natives (Write-Progress) à toutes les étapes longues
  2. Assouplissement de la détection accountId CMDB Assets (résolution du 0 par accountId)
  3. Résolution robuste des libellés SIMP (Tickets Jira BUD-* + Objets Assets CMDB) avec gestion d'erreurs silencieuse
  4. Optimisation des boucles d'enrichissement sur gros volumes de données (>25k worklogs)

Sorties générées dans ./exports/ :
  - EXPORT_JIRADOT_SUIVI-ACTIVITE_{timestamp}.xlsx
  - EXPORT_JIRADOT_SIMP_{timestamp}.xlsx
  - EtatsFinanciers_DETERMINE_{MM-YYYY}_{timestamp}.csv
  - EtatsFinanciers_SIMP_{MM-YYYY}_{timestamp}.csv
#>

[CmdletBinding()]
param(
    [datetime]$From,
    [datetime]$To,
    [ValidateSet("DETERMINE","SIMP","BOTH")][string]$Format,
    [string]$AssetsWorkspaceId = "2898703d-5f14-4527-9f8a-ad32b9279ca1",
    [switch]$UseSystemProxy    = $true,
    [switch]$ProxyUseDefaultCredentials = $true,
    [string]$ProxyUrl,
    [switch]$SkipCachePrompt
)

# ============================================================
# 0. DOSSIERS ET INITIALISATION
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
$scriptName = "Invoke-EtatsFinanciers"
$logFile    = Join-Path $logsDir ($scriptName + "_" + $runStamp + ".log")

function Write-Log {
    param([string]$Message = "", [ValidateSet("INFO","WARN","ERROR","DEBUG")][string]$Level = "INFO")
    if ([string]::IsNullOrWhiteSpace($Message)) { return }
    $ts = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    try { Add-Content -Path $logFile -Value ("[$ts] [$Level] " + $Message) -ErrorAction Stop } catch {}
}
function Write-Info($msg)   { Write-Host ("[INFO] " + $msg) -ForegroundColor Green;  Write-Log $msg "INFO"  }
function Write-Warn($msg)   { Write-Warning $msg;                                     Write-Log $msg "WARN"  }
function Write-ErrLog($msg) { Write-Error $msg;                                       Write-Log $msg "ERROR" }

Write-Log ("=== DEBUT EXÉCUTION " + $scriptName + " ===") "INFO"

# Cache mémoire global pour les libellés SIMP/Budgets
$global:budgetLibelleCache = @{}

# ============================================================
# 1. HELPERS CACHE LOCAL JSON (24h)
# ============================================================
function Load-JsonCache([string]$Path, [int]$MaxAgeHours = 24) {
    if (-not (Test-Path $Path)) { return $null }
    $fileInfo = Get-Item $Path
    $age = (Get-Date) - $fileInfo.LastWriteTime
    if ($age.TotalHours -gt $MaxAgeHours) {
        Write-Warn ("Cache expiré (" + [math]::Round($age.TotalHours, 1) + "h > " + $MaxAgeHours + "h) : " + (Split-Path $Path -Leaf))
        return $null
    }
    try {
        $raw = [System.IO.File]::ReadAllText($Path, [System.Text.Encoding]::UTF8)
        return ($raw | ConvertFrom-Json)
    } catch { return $null }
}

function Save-JsonCache([string]$Path, $Object) {
    try {
        $json = ($Object | ConvertTo-Json -Depth 10)
        $utf8NoBom = New-Object System.Text.UTF8Encoding($false)
        [System.IO.File]::WriteAllText($Path, $json, $utf8NoBom)
    } catch { Write-Warn ("Erreur sauvegarde cache JSON : " + $_.Exception.Message) }
}

# ============================================================
# 2. PROXY ET HELPERS HTTP / ENCODAGE / CUSTOM FIELDS / BUDGET
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

function ConvertTo-DateFR {
    param($val)
    if ([string]::IsNullOrWhiteSpace($val)) { return "" }
    $s = [string]$val.Trim()
    if ($s -match "^\d{13}$") { return (Get-Date "1970-01-01").AddMilliseconds([long]$s).ToString("dd/MM/yyyy") }
    if ($s -match "^\d{10}$") { return (Get-Date "1970-01-01").AddSeconds([long]$s).ToString("dd/MM/yyyy") }
    $dt = $null
    $formats = @("yyyy-MM-dd","yyyy-MM-ddTHH:mm:ss","yyyy-MM-ddTHH:mm:ssZ","dd/MM/yyyy","dd-MM-yyyy","MM/dd/yyyy")
    foreach ($fmt in $formats) {
        try { $dt = [datetime]::ParseExact($s, $fmt, [System.Globalization.CultureInfo]::InvariantCulture); break } catch {}
    }
    if (-not $dt) { try { $dt = [datetime]::Parse($s) } catch {} }
    if ($dt) { return $dt.ToString("dd/MM/yyyy") }
    return $s
}

function Get-CustomFieldValue {
    param($cf)
    if ($null -eq $cf) { return "" }
    if ($cf -is [string]) { return $cf.Trim() }
    if ($cf.key)   { return [string]$cf.key }
    if ($cf.value) { return [string]$cf.value }
    if ($cf.name)  { return [string]$cf.name }
    if ($cf -is [array] -and $cf.Count -gt 0) {
        $first = $cf[0]
        if ($first -is [string]) { return $first.Trim() }
        if ($first.key)   { return [string]$first.key }
        if ($first.value) { return [string]$first.value }
        if ($first.name)  { return [string]$first.name }
    }
    return [string]$cf
}

# Extraction Clé et Libellé Budget SIMP depuis customfield_10183
function Get-BudgetInfo {
    param($cfBudget)
    $bKey  = ""
    $bSumm = ""

    if ($null -eq $cfBudget) { return [pscustomobject]@{ Key = ""; Summary = "" } }

    function Get-ObjSummary($obj) {
        if ($null -eq $obj) { return "" }
        if ($obj.summary)      { return [string]$obj.summary }
        if ($obj.name)         { return [string]$obj.name }
        if ($obj.label)        { return [string]$obj.label }
        if ($obj.displayName)  { return [string]$obj.displayName }
        if ($obj.displayValue) { return [string]$obj.displayValue }
        if ($obj.value -and [string]$obj.value -ne [string]$obj.key -and [string]$obj.value -ne [string]$obj.objectKey) { return [string]$obj.value }
        return ""
    }

    function Get-ObjKey($obj) {
        if ($null -eq $obj) { return "" }
        if ($obj.objectKey) { return [string]$obj.objectKey }
        if ($obj.key)       { return [string]$obj.key }
        if ($obj.code)      { return [string]$obj.code }
        if ($obj.value)     { return [string]$obj.value }
        if ($obj.id)        { return [string]$obj.id }
        return ""
    }

    if ($cfBudget -is [string]) {
        $str = $cfBudget.Trim()
        if ($str -match "^([A-Z0-9_-]+)\s*[\-\|—:]\s*(.+)$") {
            $bKey  = $matches[1].Trim()
            $bSumm = $matches[2].Trim()
        } elseif ($str -match "^(.+)\s*\(([B0-9][A-Z0-9_-]+)\)$") {
            $bSumm = $str
            $bKey  = $matches[2].Trim()
        } else {
            $bKey = $str
        }
    } elseif ($cfBudget -is [array] -and $cfBudget.Count -gt 0) {
        $first = $cfBudget[0]
        if ($first -is [string]) {
            $bKey = $first.Trim()
        } else {
            $bKey  = Get-ObjKey $first
            $bSumm = Get-ObjSummary $first
        }
    } else {
        $bKey  = Get-ObjKey $cfBudget
        $bSumm = Get-ObjSummary $cfBudget
    }

    return [pscustomobject]@{
        Key     = Fix-Encoding $bKey
        Summary = Fix-Encoding $bSumm
    }
}

function Invoke-ApiGet {
    param([Parameter(Mandatory)][string]$Url, [Parameter(Mandatory)][hashtable]$Headers)
    $params = @{ Method='GET'; Uri=$Url; Headers=$Headers; ContentType='application/json'; UseBasicParsing=$true; ErrorAction='Stop' }
    $px = Get-EffectiveProxyUri -TargetUrl $Url
    if ($px) { $params.Proxy=$px; if ($ProxyUseDefaultCredentials) { $params.ProxyUseDefaultCredentials=$true } }
    try {
        $resp   = Invoke-WebRequest @params
        $stream = $resp.RawContentStream; $stream.Position = 0
        $reader = New-Object System.IO.StreamReader($stream, [System.Text.Encoding]::UTF8, $true)
        $raw    = $reader.ReadToEnd(); $reader.Close()
        return $raw | ConvertFrom-Json
    } catch { Write-ErrLog ("GET " + $Url + " : " + $_.Exception.Message); throw }
}

function Invoke-ApiPostUtf8 {
    param([Parameter(Mandatory)][string]$Url, [Parameter(Mandatory)][hashtable]$Headers, [Parameter(Mandatory)][string]$JsonBody)
    $params = @{ Method='POST'; Uri=$Url; Headers=$Headers; Body=[System.Text.Encoding]::UTF8.GetBytes($JsonBody); ContentType="application/json"; UseBasicParsing=$true; ErrorAction='Stop' }
    $px = Get-EffectiveProxyUri -TargetUrl $Url
    if ($px) { $params.Proxy=$px; if ($ProxyUseDefaultCredentials) { $params.ProxyUseDefaultCredentials=$true } }
    try {
        $resp   = Invoke-WebRequest @params
        $stream = $resp.RawContentStream; $stream.Position = 0
        $reader = New-Object System.IO.StreamReader($stream, [System.Text.Encoding]::UTF8, $true)
        $raw    = $reader.ReadToEnd(); $reader.Close()
        return $raw | ConvertFrom-Json
    } catch { Write-ErrLog ("POST " + $Url + " : " + $_.Exception.Message); throw }
}

# ============================================================
# 3. DIALOGUES INTERACTIFS (Mois & Gestion du Cache 1 par 1)
# ============================================================
function Show-EtatsFinanciersDialog {
    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing

    $form = New-Object System.Windows.Forms.Form
    $form.Text            = "États Financiers Atlassian — Paramètres d'extraction"
    $form.Size            = New-Object System.Drawing.Size(460, 310)
    $form.StartPosition   = "CenterScreen"
    $form.FormBorderStyle = "FixedDialog"
    $form.MaximizeBox     = $false; $form.MinimizeBox = $false; $form.TopMost = $true

    $lblTitre = New-Object System.Windows.Forms.Label
    $lblTitre.Text     = "Extraction États Financiers — Export CSV & XLSX"
    $lblTitre.Font     = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
    $lblTitre.Location = New-Object System.Drawing.Point(20, 12)
    $lblTitre.Size     = New-Object System.Drawing.Size(400, 22)
    $form.Controls.Add($lblTitre)

    $lblMois = New-Object System.Windows.Forms.Label
    $lblMois.Text     = "Période mensuelle à extraire :"
    $lblMois.Font     = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $lblMois.Location = New-Object System.Drawing.Point(20, 48)
    $lblMois.Size     = New-Object System.Drawing.Size(400, 20)
    $form.Controls.Add($lblMois)

    $comboMois = New-Object System.Windows.Forms.ComboBox
    $comboMois.DropDownStyle = "DropDownList"
    $comboMois.Location      = New-Object System.Drawing.Point(20, 72)
    $comboMois.Size          = New-Object System.Drawing.Size(160, 25)
    @("Janvier","Février","Mars","Avril","Mai","Juin","Juillet","Août","Septembre","Octobre","Novembre","Décembre") |
        ForEach-Object { $comboMois.Items.Add($_) | Out-Null }
    $today = Get-Date
    $comboMois.SelectedIndex = $today.AddMonths(-1).Month - 1
    $form.Controls.Add($comboMois)

    $comboAnnee = New-Object System.Windows.Forms.ComboBox
    $comboAnnee.DropDownStyle = "DropDownList"
    $comboAnnee.Location      = New-Object System.Drawing.Point(195, 72)
    $comboAnnee.Size          = New-Object System.Drawing.Size(90, 25)
    @(($today.Year - 1), $today.Year, ($today.Year + 1)) |
        ForEach-Object { $comboAnnee.Items.Add($_) | Out-Null }
    $comboAnnee.SelectedItem = $today.AddMonths(-1).Year
    $form.Controls.Add($comboAnnee)

    $lblFormat = New-Object System.Windows.Forms.Label
    $lblFormat.Text     = "Périmètre d'extraction / Format :"
    $lblFormat.Font     = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $lblFormat.Location = New-Object System.Drawing.Point(20, 115)
    $lblFormat.Size     = New-Object System.Drawing.Size(400, 20)
    $form.Controls.Add($lblFormat)

    $rbBoth = New-Object System.Windows.Forms.RadioButton
    $rbBoth.Text     = "LES DEUX  (DÉTERMINÉ/XLSX + SIMP/XLSX)"
    $rbBoth.Checked  = $true
    $rbBoth.Location = New-Object System.Drawing.Point(30, 140)
    $rbBoth.Size     = New-Object System.Drawing.Size(380, 22)
    $form.Controls.Add($rbBoth)

    $rbDetermine = New-Object System.Windows.Forms.RadioButton
    $rbDetermine.Text     = "DÉTERMINÉ  (Détail Initiatives/Collaborateurs + XLSX)"
    $rbDetermine.Location = New-Object System.Drawing.Point(30, 165)
    $rbDetermine.Size     = New-Object System.Drawing.Size(390, 22)
    $form.Controls.Add($rbDetermine)

    $rbSimp = New-Object System.Windows.Forms.RadioButton
    $rbSimp.Text     = "SIMP  (Agrégation macro par Domaine SIMP + XLSX)"
    $rbSimp.Location = New-Object System.Drawing.Point(30, 190)
    $rbSimp.Size     = New-Object System.Drawing.Size(380, 22)
    $form.Controls.Add($rbSimp)

    $btnOK = New-Object System.Windows.Forms.Button
    $btnOK.Text         = "▶  Lancer l'extraction"
    $btnOK.Font         = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $btnOK.Location     = New-Object System.Drawing.Point(80, 238)
    $btnOK.Size         = New-Object System.Drawing.Size(155, 32)
    $btnOK.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $form.AcceptButton  = $btnOK
    $form.Controls.Add($btnOK)

    $btnCancel = New-Object System.Windows.Forms.Button
    $btnCancel.Text         = "Annuler"
    $btnCancel.Location     = New-Object System.Drawing.Point(250, 238)
    $btnCancel.Size         = New-Object System.Drawing.Size(100, 32)
    $btnCancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $form.CancelButton      = $btnCancel
    $form.Controls.Add($btnCancel)

    $result = $form.ShowDialog()
    if ($result -ne [System.Windows.Forms.DialogResult]::OK) { Write-Info "Extraction annulée par l'utilisateur."; exit 0 }

    $mIdx   = $comboMois.SelectedIndex + 1
    $yVal   = [int]$comboAnnee.SelectedItem
    $dtFrom = Get-Date -Year $yVal -Month $mIdx -Day 1
    $dtTo   = $dtFrom.AddMonths(1).AddDays(-1)

    $fmtChoice = "BOTH"
    if ($rbDetermine.Checked) { $fmtChoice = "DETERMINE" }
    elseif ($rbSimp.Checked)  { $fmtChoice = "SIMP" }

    $form.Dispose()
    return [pscustomobject]@{ From = $dtFrom.Date; To = $dtTo.Date; Format = $fmtChoice }
}

function Show-CacheRefreshDialog {
    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing

    $form = New-Object System.Windows.Forms.Form
    $form.Text            = "États Financiers — Gestion du Cache Local"
    $form.Size            = New-Object System.Drawing.Size(420, 290)
    $form.StartPosition   = "CenterScreen"
    $form.FormBorderStyle = "FixedDialog"
    $form.MaximizeBox     = $false; $form.MinimizeBox = $false; $form.TopMost = $true

    $lt = New-Object System.Windows.Forms.Label
    $lt.Text     = "Sélectionnez les données à rafraîchir (1 par 1) :"
    $lt.Font     = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $lt.Location = New-Object System.Drawing.Point(20, 15)
    $lt.Size     = New-Object System.Drawing.Size(370, 20)
    $form.Controls.Add($lt)

    $li = New-Object System.Windows.Forms.Label
    $li.Text      = "(Coché = Rechargé depuis les API. Décoché = Cache JSON local < 24h)"
    $li.ForeColor = [System.Drawing.Color]::Gray
    $li.Location  = New-Object System.Drawing.Point(20, 38)
    $li.Size      = New-Object System.Drawing.Size(370, 20)
    $form.Controls.Add($li)

    $y = 68
    $chkW = New-Object System.Windows.Forms.CheckBox
    $chkW.Text     = "Worklogs Tempo (saisies de temps du mois)"
    $chkW.Checked  = $false
    $chkW.Location = New-Object System.Drawing.Point(30, $y)
    $chkW.Size     = New-Object System.Drawing.Size(350, 22)
    $form.Controls.Add($chkW); $y += 28

    $chkA = New-Object System.Windows.Forms.CheckBox
    $chkA.Text     = "Fiches Assets CMDB (Référentiel Personne)"
    $chkA.Checked  = $false
    $chkA.Location = New-Object System.Drawing.Point(30, $y)
    $chkA.Size     = New-Object System.Drawing.Size(350, 22)
    $form.Controls.Add($chkA); $y += 28

    $chkI = New-Object System.Windows.Forms.CheckBox
    $chkI.Text     = "Hiérarchie Jira (Tickets, Initiatives & Budgets)"
    $chkI.Checked  = $false
    $chkI.Location = New-Object System.Drawing.Point(30, $y)
    $chkI.Size     = New-Object System.Drawing.Size(350, 22)
    $form.Controls.Add($chkI); $y += 28

    $chkU = New-Object System.Windows.Forms.CheckBox
    $chkU.Text     = "Profils Utilisateurs Jira (Comptes Atlassian)"
    $chkU.Checked  = $false
    $chkU.Location = New-Object System.Drawing.Point(30, $y)
    $chkU.Size     = New-Object System.Drawing.Size(350, 22)
    $form.Controls.Add($chkU); $y += 35

    $btnOK = New-Object System.Windows.Forms.Button
    $btnOK.Text         = "Continuer"
    $btnOK.Font         = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $btnOK.Location     = New-Object System.Drawing.Point(100, $y)
    $btnOK.Size         = New-Object System.Drawing.Size(95, 30)
    $btnOK.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $form.AcceptButton  = $btnOK
    $form.Controls.Add($btnOK)

    $btnCancel = New-Object System.Windows.Forms.Button
    $btnCancel.Text         = "Annuler"
    $btnCancel.Location     = New-Object System.Drawing.Point(210, $y)
    $btnCancel.Size         = New-Object System.Drawing.Size(90, 30)
    $btnCancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $form.CancelButton      = $btnCancel
    $form.Controls.Add($btnCancel)

    $result = $form.ShowDialog()
    if ($result -ne [System.Windows.Forms.DialogResult]::OK) { Write-Info "Extraction annulée."; exit 0 }

    $choices = [pscustomobject]@{
        RefreshWorklogs = $chkW.Checked
        RefreshAssets   = $chkA.Checked
        RefreshIssues   = $chkI.Checked
        RefreshUsers    = $chkU.Checked
    }
    $form.Dispose()
    return $choices
}
# ============================================================
# 4. INITIALISATION PÉRIODE ET CREDENTIALS
# ============================================================
if (-not $From -or -not $To -or -not $Format) {
    $uiParams = Show-EtatsFinanciersDialog
    $From     = $uiParams.From
    $To       = $uiParams.To
    $Format   = $uiParams.Format
}

$fromStr = $From.ToString("yyyy-MM-dd")
$toStr   = $To.ToString("yyyy-MM-dd")
$moisKey = $From.ToString("MM-yyyy")
$fyKey   = "FY" + $From.ToString("yy")
$mKey    = "M"  + $From.ToString("MM")

Write-Info ("Période sélectionnée : " + $From.ToString("dd/MM/yyyy") + " -> " + $To.ToString("dd/MM/yyyy"))
Write-Info ("Périmètre / Format   : " + $Format)

if (-not $SkipCachePrompt) {
    $cacheChoices = Show-CacheRefreshDialog
} else {
    $cacheChoices = [pscustomobject]@{ RefreshWorklogs=$false; RefreshAssets=$false; RefreshIssues=$false; RefreshUsers=$false }
}

# Jira Credentials
$jiraCredFile = Join-Path $secretsDir "jira-jiradot.cred.xml"
if (-not (Test-Path $jiraCredFile)) { throw "Fichier Jira creds introuvable: $jiraCredFile" }
$jiraData    = Import-Clixml -Path $jiraCredFile
$jiraBaseUrl = ($jiraData.JiraBaseUrl).TrimEnd("/")
$jiraCred    = $jiraData.Credential
$b64         = [Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes($jiraCred.UserName + ":" + $jiraCred.GetNetworkCredential().Password))
$jiraHeaders = @{ Authorization = "Basic " + $b64; Accept = "application/json" }

# Tempo Credentials
$tempoTokenFile = Join-Path $secretsDir "tempo-token.xml"
if (-not (Test-Path $tempoTokenFile)) { throw "Fichier Tempo token introuvable: $tempoTokenFile" }
$tempoData    = Import-Clixml -Path $tempoTokenFile
$tempoToken   = [Runtime.InteropServices.Marshal]::PtrToStringAuto([Runtime.InteropServices.Marshal]::SecureStringToBSTR($tempoData.Token))
$tempoHeaders = @{ Authorization = "Bearer " + $tempoToken; Accept = "application/json" }

$assetsAqlUrl    = $jiraBaseUrl + "/gateway/api/jsm/assets/workspace/" + $AssetsWorkspaceId + "/v1/object/aql"
$assetsSchemaUrl = $jiraBaseUrl + "/gateway/api/jsm/assets/workspace/" + $AssetsWorkspaceId + "/v1/objectschema/list"

# ============================================================
# 5. FONCTIONS D'EXPORTATION (CSV & XLSX)
# ============================================================
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
    Write-Info ("  -> CSV généré : " + $Path + " (" + $Rows.Count + " lignes)")
}

function Export-GenericXlsx {
    param(
        [Parameter(Mandatory=$true)][string]$Path,
        [Parameter(Mandatory=$true)][string]$SheetName,
        [Parameter(Mandatory=$true)][string[]]$Headers,
        [Parameter(Mandatory=$true)]$Rows
    )

    $exportObjects = New-Object System.Collections.Generic.List[object]
    foreach ($r in $Rows) {
        $obj = [ordered]@{}
        foreach ($h in $Headers) {
            $val = $r.$h
            $obj[$h] = if ($null -eq $val) { "" } else { $val }
        }
        $exportObjects.Add([pscustomobject]$obj) | Out-Null
    }

    if ($exportObjects.Count -eq 0) {
        Write-Warn ("  -> Aucune ligne à exporter dans le fichier XLSX : " + $Path)
        return
    }

    # Option 1 : Module Import-Excel
    if (Get-Module -ListAvailable -Name Import-Excel) {
        try {
            Import-Module Import-Excel -ErrorAction Stop
            $exportObjects | Export-Excel -Path $Path -WorksheetName $SheetName -AutoSize -TableStyle Medium6 -Show:$false
            Write-Info ("  -> XLSX généré (via Import-Excel) : " + $Path + " (" + $exportObjects.Count + " lignes)")
            return
        } catch { Write-Warn ("Import-Excel disponible mais erreur : " + $_.Exception.Message) }
    }

    # Option 2 : Excel COM Object
    try {
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false; $excel.DisplayAlerts = $false
        $wb = $excel.Workbooks.Add()
        $ws = $wb.Worksheets.Item(1)
        $ws.Name = $SheetName

        for ($col = 0; $col -lt $Headers.Count; $col++) {
            $cell = $ws.Cells.Item(1, $col + 1)
            $cell.Value2 = $Headers[$col]
            $cell.Font.Bold = $true
            $cell.Font.ColorIndex = 2
            $cell.Interior.ColorIndex = 23
        }

        for ($r = 0; $r -lt $exportObjects.Count; $r++) {
            $rowObj = $exportObjects[$r]
            for ($c = 0; $c -lt $Headers.Count; $c++) {
                $cellVal = $rowObj.($Headers[$c])
                $ws.Cells.Item($r + 2, $c + 1) = if ($null -eq $cellVal) { "" } else { [string]$cellVal }
            }
        }

        $ws.Columns.AutoFit() | Out-Null
        $wb.SaveAs($Path, 51)
        $wb.Close($false)
        $excel.Quit()
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
        Write-Info ("  -> XLSX généré (via Excel COM) : " + $Path + " (" + $exportObjects.Count + " lignes)")
        return
    } catch { Write-Warn ("Excel COM non disponible : " + $_.Exception.Message) }

    # Option 3 : Auto-installation Import-Excel
    try {
        Write-Info "Installation automatique de Import-Excel pour la génération XLSX..."
        Install-Module -Name Import-Excel -Scope CurrentUser -Force -AllowClobber -ErrorAction Stop
        Import-Module Import-Excel -ErrorAction Stop
        $exportObjects | Export-Excel -Path $Path -WorksheetName $SheetName -AutoSize -TableStyle Medium6 -Show:$false
        Write-Info ("  -> XLSX généré (via Import-Excel auto-installé) : " + $Path + " (" + $exportObjects.Count + " lignes)")
    } catch {
        Write-ErrLog ("Erreur génération XLSX : " + $_.Exception.Message)
    }
}

# ============================================================
# 6. CHARGEMENT CMDB ASSETS PERSONNE (EXTRACTION LARGE ACCOUNTID)
# ============================================================
Write-Info "=== 1/4 Chargement de la CMDB Assets Référentiel Personne ==="

$assetsCacheFile        = Join-Path $cacheDir "assets_referentiel_personne.json"
$assetsUsersByAccountId = @{}
$assetsUsersByEmail     = @{}
$assetsUsersByName      = @{}
$assetsList             = $null

if (-not $cacheChoices.RefreshAssets) {
    $assetsList = Load-JsonCache -Path $assetsCacheFile
    if ($assetsList) { Write-Info ("  -> Données Assets chargées depuis le cache local (" + $assetsList.Count + " fiches)") }
}

if (-not $assetsList) {
    Write-Info "  -> Interrogation API JSM Assets (Schéma Référentiel personne ID=6)..."
    $assetsList = New-Object System.Collections.Generic.List[object]
    $rpSchemaId = "6"

    try {
        $respSchemas = Invoke-ApiGet -Url $assetsSchemaUrl -Headers $jiraHeaders
        $sList = if ($respSchemas.values) { $respSchemas.values } else { $respSchemas.objectSchemas }
        foreach ($sch in $sList) {
            if ([string]$sch.name -ilike "*Référentiel personne*" -or [string]$sch.name -ieq "RP") {
                $rpSchemaId = [string]$sch.id; break
            }
        }
    } catch {}

    Write-Info ("  -> Schéma RP résolu : ID=" + $rpSchemaId)
    $aqlRP   = "objectSchemaId = " + $rpSchemaId
    $startAt = 0; $maxResults = 200; $isLast = $false; $pageNum = 0

    while (-not $isLast) {
        $pageNum++
        Write-Progress -Id 0 -Activity "Chargement CMDB Assets Référentiel Personne" `
            -Status ("Page " + $pageNum + " — " + $assetsList.Count + " fiches chargées...") `
            -PercentComplete -1

        $urlAql  = $assetsAqlUrl + "?startAt=" + $startAt + "&maxResults=" + $maxResults + "&includeAttributes=true"
        $bodyAql = (@{ qlQuery = $aqlRP } | ConvertTo-Json -Depth 3)
        $respAql = $null
        try { $respAql = Invoke-ApiPostUtf8 -Url $urlAql -Headers $jiraHeaders -JsonBody $bodyAql } catch { break }
        if (-not $respAql -or -not $respAql.values) { break }

        $attrDict = @{}
        if ($respAql.objectTypeAttributes) {
            foreach ($ota in $respAql.objectTypeAttributes) {
                if ($ota.id -and $ota.name) { $attrDict[[string]$ota.id] = Fix-Encoding ([string]$ota.name) }
            }
        }

        foreach ($pObj in $respAql.values) {
            if (-not $pObj.id) { continue }

            $accId = ""; $nom = ""; $prenom = ""; $statut = "Actif"; $typeRes = "Prestataire"
            $dateEntree = ""; $codeSimp = ""; $directionStr = ""; $serviceStr = ""
            $matricule = ""; $societe = "Harmonie Mutuelle"; $codeCigref = ""; $compteJiraRaw = ""

            foreach ($attr in $pObj.attributes) {
                $aName = $attrDict[[string]$attr.objectTypeAttributeId]
                if (-not $aName -and $attr.objectTypeAttribute -and $attr.objectTypeAttribute.name) {
                    $aName = Fix-Encoding ([string]$attr.objectTypeAttribute.name)
                }
                if (-not $aName) { continue }

                $vals  = $attr.objectAttributeValues
                if (-not $vals -or $vals.Count -eq 0) { continue }
                $vVal  = Fix-Encoding ([string]$vals[0].value)
                $vDisp = Fix-Encoding ([string]$vals[0].displayValue)

                if ($aName -ieq "Compte Jira" -or $aName -ilike "*AccountId*" -or $aName -ilike "*User*ID*") {
                    if ($vals[0].user -and $vals[0].user.accountId) {
                        $accId = [string]$vals[0].user.accountId
                    } elseif ($vVal -and $vVal.Trim() -notmatch "\s" -and $vVal.Length -ge 12) {
                        $accId = $vVal.Trim()
                    } elseif ($vDisp -match "([0-9a-f]{24}|[0-9a-z]{5,10}:[0-9a-f-]{30,})") {
                        $accId = $matches[1].Trim()
                    }
                    $compteJiraRaw = $vDisp
                }
                elseif ($aName -ieq "Nom")                              { $nom = $vDisp }
                elseif ($aName -ieq "Prénom" -or $aName -ieq "Prenom") { $prenom = $vDisp }
                elseif ($aName -ieq "Statut")                          { $statut = $vDisp }
                elseif ($aName -ilike "*Type ressource*")               { $typeRes = $vDisp }
                elseif ($aName -ieq "Date Entrée")                     { $dateEntree = ConvertTo-DateFR $vDisp }
                elseif ($aName -ieq "Matricule")                       { $matricule = $vDisp }
                elseif ($aName -ieq "Société")                         { $societe = $vDisp }
                elseif ($aName -ieq "Domaine SIMP")                    { $codeSimp = $vDisp }
                elseif ($aName -ieq "Direction")                       { $directionStr = $vDisp }
                elseif ($aName -ilike "*Affectation #2*")              { $serviceStr = $vDisp }
                elseif ($aName -ilike "*CIGREF*")                      { $codeCigref = $vDisp }
            }

            $labelAssets = if ($pObj.label) { Fix-Encoding ([string]$pObj.label) } else { "" }
            if ([string]::IsNullOrWhiteSpace($nom)) { $nom = $labelAssets }

            $extractedEmail = ""
            if ($compteJiraRaw -match "\(([^)]+@[^)]+)\)") {
                $extractedEmail = $matches[1].Trim().ToLower()
            }

            $assetsList.Add([pscustomobject]@{
                Label           = $labelAssets
                AccountId       = $accId
                ExtractedEmail  = $extractedEmail
                Statut          = $statut
                TypeRessource   = $typeRes
                DateEntree      = $dateEntree
                CodeDomaineSIMP = $codeSimp
                Nom             = $nom
                Prenom          = $prenom
                Matricule       = $matricule
                Direction       = $directionStr
                Service         = $serviceStr
                Societe         = $societe
                CodeCIGREF      = $codeCigref
            }) | Out-Null
        }
        $isLast   = if ($null -ne $respAql.isLast) { [bool]$respAql.isLast } else { $true }
        $startAt += $maxResults
    }
    Write-Progress -Id 0 -Activity "Chargement CMDB Assets Référentiel Personne" -Completed

    Save-JsonCache -Path $assetsCacheFile -Object $assetsList
    Write-Info ("  -> Fiches Assets sauvegardées dans le cache local (" + $assetsList.Count + " fiches)")
}

# Indexation triple en mémoire
foreach ($uRec in $assetsList) {
    if (-not [string]::IsNullOrWhiteSpace($uRec.AccountId)) {
        $assetsUsersByAccountId[$uRec.AccountId] = $uRec
    }
    if (-not [string]::IsNullOrWhiteSpace($uRec.ExtractedEmail)) {
        $assetsUsersByEmail[$uRec.ExtractedEmail] = $uRec
    }
    if (-not [string]::IsNullOrWhiteSpace($uRec.Label)) {
        $assetsUsersByName[$uRec.Label.Trim().ToLower()] = $uRec
    }
}
Write-Info ("  Indexation CMDB -> " + $assetsUsersByAccountId.Count + " par accountId | " +
            $assetsUsersByEmail.Count + " par email | " + $assetsUsersByName.Count + " par nom")

# ============================================================
# 7. EXTRACTION WORKLOGS TEMPO
# ============================================================
Write-Info "=== 2/4 Récupération des Worklogs Tempo ==="

$tempoCacheFile = Join-Path $cacheDir ("worklogs_tempo_" + $fromStr + "_" + $toStr + ".json")
$worklogs = $null

if (-not $cacheChoices.RefreshWorklogs) {
    $worklogs = Load-JsonCache -Path $tempoCacheFile
    if ($worklogs) { Write-Info ("  -> Worklogs Tempo chargés depuis le cache local (" + $worklogs.Count + " enregistrements)") }
}

if (-not $worklogs) {
    Write-Info "  -> Interrogation API Tempo v4..."
    $worklogs  = New-Object System.Collections.Generic.List[object]
    $tempoUrl  = "https://api.eu.tempo.io/4/worklogs?from=" + $fromStr + "&to=" + $toStr
    $tempoPage = 0

    while ($true) {
        $tempoPage++
        Write-Progress -Id 0 -Activity "Récupération Worklogs Tempo" `
            -Status ("Page " + $tempoPage + " — " + $worklogs.Count + " worklogs chargés...") `
            -PercentComplete -1
        $resp = Invoke-ApiGet -Url $tempoUrl -Headers $tempoHeaders
        if ($resp.results) { foreach ($wl in $resp.results) { $worklogs.Add($wl) | Out-Null } }
        if ($resp.metadata -and $resp.metadata.next) { $tempoUrl = $resp.metadata.next } else { break }
    }
    Write-Progress -Id 0 -Activity "Récupération Worklogs Tempo" -Completed

    Save-JsonCache -Path $tempoCacheFile -Object $worklogs
    Write-Info ("  -> Worklogs Tempo sauvegardés dans le cache (" + $worklogs.Count + " enregistrements)")
}

# ============================================================
# 8. ENRICHISSEMENT JIRA (UTILISATEURS & HIÉRARCHIE TICKETS)
# ============================================================
Write-Info "=== 3/4 Enrichissement Profils Utilisateurs & Remontée Hiérarchique Jira ==="

# 8a. Profils Utilisateurs Jira
$usersCacheFile = Join-Path $cacheDir "jira_users_cache.json"
$jiraUserCache  = @{}

if (-not $cacheChoices.RefreshUsers) {
    $cachedUsers = Load-JsonCache -Path $usersCacheFile
    if ($cachedUsers) {
        foreach ($p in $cachedUsers.PSObject.Properties) { $jiraUserCache[$p.Name] = $p.Value }
        Write-Info ("  -> Profils Utilisateurs Jira chargés depuis le cache (" + $jiraUserCache.Count + " utilisateurs)")
    }
}

$uniqueAccIds = $worklogs | ForEach-Object { $_.author.accountId } | Select-Object -Unique
$totalAccIds  = @($uniqueAccIds).Count
$usersFetched = 0
$idxAcc       = 0

foreach ($accId in $uniqueAccIds) {
    $idxAcc++
    if ([string]::IsNullOrWhiteSpace($accId)) { continue }
    Write-Progress -Id 1 -Activity "Profils Utilisateurs Jira" `
        -Status ("Compte " + $idxAcc + " / " + $totalAccIds + " — Nouveaux : " + $usersFetched) `
        -PercentComplete ([math]::Round($idxAcc / $totalAccIds * 100))

    if (-not $jiraUserCache.ContainsKey($accId)) {
        try {
            $uUrl  = $jiraBaseUrl + "/rest/api/3/user?accountId=" + $accId
            $uResp = Invoke-ApiGet -Url $uUrl -Headers $jiraHeaders
            $jiraUserCache[$accId] = [pscustomobject]@{
                DisplayName  = Fix-Encoding ([string]$uResp.displayName)
                EmailAddress = if ($uResp.emailAddress) { [string]$uResp.emailAddress.Trim().ToLower() } else { "" }
            }
            $usersFetched++
        } catch {
            $jiraUserCache[$accId] = [pscustomobject]@{
                DisplayName  = "Utilisateur Jira (" + $accId + ")"
                EmailAddress = ""
            }
        }
    }
}
Write-Progress -Id 1 -Activity "Profils Utilisateurs Jira" -Completed

if ($usersFetched -gt 0 -or -not (Test-Path $usersCacheFile)) {
    Save-JsonCache -Path $usersCacheFile -Object $jiraUserCache
    Write-Info ("  -> " + $usersFetched + " nouveaux profils Jira récupérés et sauvegardés")
}
Write-Info ("  Profils Jira total en mémoire : " + $jiraUserCache.Count + " utilisateurs")

# 8b. Hiérarchie Tickets Jira avec barres de progression
$issuesCacheFile = Join-Path $cacheDir "jira_issues_hierarchy.json"
$jiraIssueCache  = @{}

if (-not $cacheChoices.RefreshIssues) {
    $cachedIssues = Load-JsonCache -Path $issuesCacheFile
    if ($cachedIssues) {
        foreach ($p in $cachedIssues.PSObject.Properties) { $jiraIssueCache[$p.Name] = $p.Value }
        Write-Info ("  -> Hiérarchies Tickets Jira chargées depuis le cache (" + $jiraIssueCache.Count + " tickets)")
    }
}

$uniqueIssueIds  = $worklogs | ForEach-Object { $_.issue.id } | Select-Object -Unique
$totalIssues     = @($uniqueIssueIds).Count
$issuesFetched   = 0
$issuesFromCache = 0
$idxIssue        = 0

Write-Info ("  Résolution hiérarchie Jira : " + $totalIssues + " tickets uniques à traiter...")

foreach ($issId in $uniqueIssueIds) {
    $idxIssue++
    if ([string]::IsNullOrWhiteSpace($issId)) { continue }

    $pct = [math]::Round($idxIssue / $totalIssues * 100)
    Write-Progress -Id 2 -Activity "Hiérarchie Jira — Remontée Tickets / Initiatives / Budgets" `
        -Status ("Ticket " + $idxIssue + " / " + $totalIssues + " | Cache: " + $issuesFromCache + " | API: " + $issuesFetched + " | " + $pct + "%") `
        -PercentComplete $pct

    if ($jiraIssueCache.ContainsKey([string]$issId)) { $issuesFromCache++; continue }

    $currId       = [string]$issId
    $visited      = @{}
    $depth        = 0
    $issueKey     = ""; $issueSummary = ""
    $initKey      = ""; $initSummary  = ""
    $budgetKey    = ""; $budgetSumm   = ""
    $codeFDR      = ""

    while ($currId -and -not $visited.ContainsKey($currId) -and $depth -lt 5) {
        $visited[$currId] = $true
        $depth++

        $iUrl  = $jiraBaseUrl + "/rest/api/3/issue/" + $currId +
                 "?fields=key,summary,parent,issuetype,customfield_10183,customfield_10124,customfield_10014"
        $iResp = $null
        try { $iResp = Invoke-ApiGet -Url $iUrl -Headers $jiraHeaders } catch { break }
        if (-not $iResp) { break }

        $key     = [string]$iResp.key
        $summary = Fix-Encoding ([string]$iResp.fields.summary)
        $iType   = if ($iResp.fields.issuetype -and $iResp.fields.issuetype.name) {
                       [string]$iResp.fields.issuetype.name } else { "" }

        $budgetInfo = Get-BudgetInfo $iResp.fields.customfield_10183
        $cfFDR      = Get-CustomFieldValue $iResp.fields.customfield_10124

        if (-not $issueKey) { $issueKey = $key; $issueSummary = $summary }
        if (-not $codeFDR -and $cfFDR) { $codeFDR = $cfFDR }

        if (-not $budgetKey -and $budgetInfo.Key) {
            $budgetKey  = $budgetInfo.Key
            $budgetSumm = $budgetInfo.Summary
        }

        if (($iType -ieq "Initiative" -or $iType -ilike "*Initiative*") -and -not $initKey) {
            $initKey = $key; $initSummary = $summary
        }

        if (($iType -ieq "Budget" -or $iType -ilike "*Budget*" -or $key -ilike "BUD-*") -and -not $budgetKey) {
            $budgetKey  = $key
            $budgetSumm = $summary
        }

        $nextParent = $null
        if ($iResp.fields.parent -and $iResp.fields.parent.id) {
            $nextParent = [string]$iResp.fields.parent.id
        } elseif ($iResp.fields.parent -and $iResp.fields.parent.key) {
            $nextParent = [string]$iResp.fields.parent.key
        } elseif ($iResp.fields.customfield_10014) {
            $nextParent = Get-CustomFieldValue $iResp.fields.customfield_10014
        }
        $currId = $nextParent
    }

    if (-not $initKey)     { $initKey     = if ($issueKey) { $issueKey } else { "AUTRE" } }
    if (-not $initSummary) { $initSummary = $issueSummary }

    $jiraIssueCache[[string]$issId] = [pscustomobject]@{
        IssueKey          = $issueKey
        IssueSummary      = $issueSummary
        InitiativeKey     = $initKey
        InitiativeSummary = $initSummary
        BudgetKey         = $budgetKey
        BudgetSummary     = $budgetSumm
        CodeFDR           = $codeFDR
    }
    $issuesFetched++

    # Sauvegarde intermédiaire tous les 200 nouveaux tickets
    if ($issuesFetched % 200 -eq 0) {
        Save-JsonCache -Path $issuesCacheFile -Object $jiraIssueCache
        Write-Info ("  -> Sauvegarde intermédiaire cache hiérarchie : " + $jiraIssueCache.Count + " tickets")
    }
}
Write-Progress -Id 2 -Activity "Hiérarchie Jira — Remontée Tickets / Initiatives / Budgets" -Completed

if ($issuesFetched -gt 0 -or -not (Test-Path $issuesCacheFile)) {
    Save-JsonCache -Path $issuesCacheFile -Object $jiraIssueCache
    Write-Info ("  -> " + $issuesFetched + " nouvelles hiérarchies résolues | " + $issuesFromCache + " depuis le cache")
}
Write-Info ("  Hiérarchies Jira total en mémoire : " + $jiraIssueCache.Count + " tickets")

# ============================================================
# 9. EXPORTATION DES FICHIERS CSV ET XLSX
# ============================================================
Write-Info "=== 4/4 Génération des fichiers CSV & XLSX ==="

# Résolution CMDB commune (Clé #1 accountId -> #2 email -> #3 nom)
function Resolve-Cmdb {
    param([string]$AccId, [string]$UserMail, [string]$DispName)
    if ($AccId -and $assetsUsersByAccountId.ContainsKey($AccId)) {
        return $assetsUsersByAccountId[$AccId]
    }
    if ($UserMail -and $assetsUsersByEmail.ContainsKey($UserMail)) {
        return $assetsUsersByEmail[$UserMail]
    }
    if ($DispName -and $assetsUsersByName.ContainsKey($DispName.Trim().ToLower())) {
        return $assetsUsersByName[$DispName.Trim().ToLower()]
    }
    return $null
}

# Résolution garantie du Libellé SIMP
# Distingue Tickets Jira (BUD-113) et Objets Assets CMDB (R1R03300_NDAPP_101F)
function Resolve-LibelleSIMP {
    param([string]$BudgetKey, [string]$BudgetSummary)

    if (-not [string]::IsNullOrWhiteSpace($BudgetSummary)) { return $BudgetSummary }
    if ([string]::IsNullOrWhiteSpace($BudgetKey))          { return "" }

    if ($global:budgetLibelleCache.ContainsKey($BudgetKey)) {
        return $global:budgetLibelleCache[$BudgetKey]
    }

    $lib = ""

    # Cas 1 : Clé de Ticket Jira (ex: BUD-113, AU-184)
    if ($BudgetKey -match "^[A-Za-z0-9]+-\d+$") {
        try {
            $bUrl   = $jiraBaseUrl + "/rest/api/3/issue/" + $BudgetKey + "?fields=summary"
            $params = @{ Method='GET'; Uri=$bUrl; Headers=$jiraHeaders; ContentType='application/json'; UseBasicParsing=$true; ErrorAction='Stop' }
            $px     = Get-EffectiveProxyUri -TargetUrl $bUrl
            if ($px) { $params.Proxy = $px; if ($ProxyUseDefaultCredentials) { $params.ProxyUseDefaultCredentials = $true } }
            $resp   = Invoke-WebRequest @params
            $stream = $resp.RawContentStream; $stream.Position = 0
            $reader = New-Object System.IO.StreamReader($stream, [System.Text.Encoding]::UTF8, $true)
            $json   = $reader.ReadToEnd() | ConvertFrom-Json; $reader.Close()
            if ($json.fields -and $json.fields.summary) {
                $lib = Fix-Encoding ([string]$json.fields.summary)
            }
        } catch {
            Write-Log ("Ticket Jira introuvable pour libellé : " + $BudgetKey) "DEBUG"
        }
    }
    # Cas 2 : Clé d'Objet Assets CMDB (ex: R1R03300_NDAPP_101F, B1TP0001_NDAPP_101I)
    else {
        try {
            $aqlQuery = "objectKey = `"" + $BudgetKey + "`""
            $bodyAql  = (@{ qlQuery = $aqlQuery } | ConvertTo-Json -Depth 3)
            $aqlUrlNa = $assetsAqlUrl + "?includeAttributes=false"
            $params   = @{ Method='POST'; Uri=$aqlUrlNa; Headers=$jiraHeaders; Body=[System.Text.Encoding]::UTF8.GetBytes($bodyAql); ContentType="application/json"; UseBasicParsing=$true; ErrorAction='Stop' }
            $px       = Get-EffectiveProxyUri -TargetUrl $aqlUrlNa
            if ($px) { $params.Proxy = $px; if ($ProxyUseDefaultCredentials) { $params.ProxyUseDefaultCredentials = $true } }
            $resp     = Invoke-WebRequest @params
            $stream   = $resp.RawContentStream; $stream.Position = 0
            $reader   = New-Object System.IO.StreamReader($stream, [System.Text.Encoding]::UTF8, $true)
            $respAql  = $reader.ReadToEnd() | ConvertFrom-Json; $reader.Close()

            if ($respAql -and $respAql.values -and $respAql.values.Count -gt 0) {
                $pObj = $respAql.values[0]
                $lib  = Fix-Encoding ([string]$pObj.label)
                if ([string]::IsNullOrWhiteSpace($lib) -and $pObj.name) {
                    $lib = Fix-Encoding ([string]$pObj.name)
                }
            }
        } catch {
            Write-Log ("Objet Assets introuvable pour libellé : " + $BudgetKey) "DEBUG"
        }
    }

    if (-not [string]::IsNullOrWhiteSpace($lib)) {
        $global:budgetLibelleCache[$BudgetKey] = $lib
    }
    return $lib
}

# ============================================================
# 9a. FORMAT DÉTERMINÉ : CSV + XLSX SUIVI-ACTIVITÉ
# ============================================================
if ($Format -eq "DETERMINE" -or $Format -eq "BOTH") {

    $determineRows = New-Object System.Collections.Generic.List[object]
    $xlsxGrouped   = @{}
    $totalWL       = $worklogs.Count
    $idxWL         = 0

    foreach ($wl in $worklogs) {
        $idxWL++
        if ($idxWL % 500 -eq 0 -or $idxWL -eq $totalWL) {
            Write-Progress -Id 3 -Activity "Génération DÉTERMINÉ" `
                -Status ("Worklog " + $idxWL + " / " + $totalWL + " — " + $xlsxGrouped.Count + " lignes agrégées") `
                -PercentComplete ([math]::Round($idxWL / $totalWL * 100))
        }

        $accId    = [string]$wl.author.accountId
        $issId    = [string]$wl.issue.id
        $jUser    = $jiraUserCache[$accId]
        $dispName = if ($jUser -and $jUser.DisplayName)  { [string]$jUser.DisplayName }  else { "" }
        $userMail = if ($jUser -and $jUser.EmailAddress) { [string]$jUser.EmailAddress } else { "" }

        $cmdb    = Resolve-Cmdb -AccId $accId -UserMail $userMail -DispName $dispName
        $jIssue  = $jiraIssueCache[[string]$issId]
        $initKey = if ($jIssue -and $jIssue.InitiativeKey) { [string]$jIssue.InitiativeKey } else { "AUTRE" }
        $heures  = [math]::Round([double]$wl.timeSpentSeconds / 3600.0, 5)
        $jours   = [math]::Round($heures / 8.0, 5)
        $typeRes = if ($cmdb -and $cmdb.TypeRessource) { $cmdb.TypeRessource } else { "Prestataire" }

        $codeSIMP         = if ($jIssue -and $jIssue.BudgetKey)     { [string]$jIssue.BudgetKey }     else { "" }
        $budgetSummaryRaw = if ($jIssue -and $jIssue.BudgetSummary) { [string]$jIssue.BudgetSummary } else { "" }
        $libelleSIMP      = Resolve-LibelleSIMP -BudgetKey $codeSIMP -BudgetSummary $budgetSummaryRaw

        $determineRows.Add([pscustomobject]@{
            idCle             = ($initKey + "_" + $typeRes + "_" + $moisKey + "_" + $accId)
            exercice          = $From.Year
            mois              = $From.Month.ToString("D2")
            dateSaisie        = ConvertTo-DateFR $wl.startDate
            accountIdJira     = $accId
            nomAffichage      = $dispName
            matricule         = if ($cmdb) { $cmdb.Matricule }        else { "" }
            nom               = if ($cmdb) { $cmdb.Nom }              else { "" }
            prenom            = if ($cmdb) { $cmdb.Prenom }           else { "" }
            typeRessource     = $typeRes
            societe           = if ($cmdb) { $cmdb.Societe }          else { "Harmonie Mutuelle" }
            direction         = if ($cmdb) { $cmdb.Direction }         else { "(Non trouve)" }
            codeDomaineSIMP   = if ($cmdb) { $cmdb.CodeDomaineSIMP }  else { "" }
            codeBudget        = $codeSIMP
            libelleBudget     = $libelleSIMP
            cleInitiative     = $initKey
            libelleInitiative = if ($jIssue) { [string]$jIssue.InitiativeSummary } else { "" }
            codeFDR           = if ($jIssue) { [string]$jIssue.CodeFDR }            else { "" }
            heuresTotales     = $heures
            joursTotaux       = $jours
        }) | Out-Null

        $groupKey = ($accId + "|" + $initKey)
        if (-not $xlsxGrouped.ContainsKey($groupKey)) {
            $nomPrenom = if ($cmdb -and $cmdb.Nom) { ($cmdb.Nom + " " + $cmdb.Prenom).Trim() } else { $dispName }
            $initLib   = if ($jIssue -and $jIssue.InitiativeSummary) {
                $initKey + " - " + [string]$jIssue.InitiativeSummary
            } else { $initKey }

            $xlsxGrouped[$groupKey] = [ordered]@{
                "Année"                     = $From.Year
                "Mois"                      = $From.Month
                "NOM Prénom"                = $nomPrenom
                "Type ressource"            = $typeRes
                "Société"                   = if ($cmdb) { $cmdb.Societe }        else { "Harmonie Mutuelle" }
                "Direction"                 = if ($cmdb) { $cmdb.Direction }       else { "(Non trouve)" }
                "Service"                   = if ($cmdb) { $cmdb.Service }         else { "" }
                "Matricule"                 = if ($cmdb) { $cmdb.Matricule }       else { "" }
                "Code domaine Ressource"    = if ($cmdb) { $cmdb.CodeDomaineSIMP } else { "" }
                "Code SIMP"                 = $codeSIMP
                "Libellé SIMP"              = $libelleSIMP
                "Code - Libellé initiative" = $initLib
                "Nb jours"                  = 0.0
                "Code CIGREF"               = if ($cmdb) { $cmdb.CodeCIGREF }         else { "" }
                "Code Gepetto"              = if ($jIssue) { [string]$jIssue.CodeFDR } else { "" }
                "Identifiant Jira"          = $accId
            }
        } else {
            if ([string]::IsNullOrWhiteSpace($xlsxGrouped[$groupKey]["Libellé SIMP"]) -and -not [string]::IsNullOrWhiteSpace($libelleSIMP)) {
                $xlsxGrouped[$groupKey]["Libellé SIMP"] = $libelleSIMP
            }
        }
        $xlsxGrouped[$groupKey]["Nb jours"] += $jours
    }
    Write-Progress -Id 3 -Activity "Génération DÉTERMINÉ" -Completed

    # Génération CSV DÉTERMINÉ
    $csvDetermine = Join-Path $exportsDir ("EtatsFinanciers_DETERMINE_" + $moisKey + "_" + $runStamp + ".csv")
    Export-CsvStrict -Path $csvDetermine `
        -Headers @("idCle","exercice","mois","dateSaisie","accountIdJira","nomAffichage","matricule","nom","prenom",
                   "typeRessource","societe","direction","codeDomaineSIMP","codeBudget","libelleBudget",
                   "cleInitiative","libelleInitiative","codeFDR","heuresTotales","joursTotaux") `
        -Rows $determineRows

    # Génération XLSX SUIVI-ACTIVITÉ
    $xlsxSuiviRows = New-Object System.Collections.Generic.List[object]
    foreach ($gKey in $xlsxGrouped.Keys) {
        $r = $xlsxGrouped[$gKey]
        $r["Nb jours"] = [math]::Round([double]$r["Nb jours"], 2)
        $xlsxSuiviRows.Add([pscustomobject]$r) | Out-Null
    }

    $xlsxSuiviHeaders = @(
        "Année","Mois","NOM Prénom","Type ressource","Société","Direction","Service","Matricule",
        "Code domaine Ressource","Code SIMP","Libellé SIMP","Code - Libellé initiative",
        "Nb jours","Code CIGREF","Code Gepetto","Identifiant Jira"
    )
    $xlsxSuiviFile = Join-Path $exportsDir ("EXPORT_JIRADOT_SUIVI-ACTIVITE_" + $runStamp + ".xlsx")
    Export-GenericXlsx -Path $xlsxSuiviFile -SheetName "Suivi Activité" -Headers $xlsxSuiviHeaders -Rows $xlsxSuiviRows
    Write-Info ("  SUIVI-ACTIVITE : " + $xlsxSuiviRows.Count + " lignes agrégées (Collaborateur x Initiative)")
}

# ============================================================
# 9b. FORMAT SIMP : CSV + XLSX EXPORT_JIRADOT_SIMP
# ============================================================
if ($Format -eq "SIMP" -or $Format -eq "BOTH") {

    $simpGrouped = @{}
    $totalWL2    = $worklogs.Count
    $idxWL2      = 0

    foreach ($wl in $worklogs) {
        $idxWL2++
        if ($idxWL2 % 500 -eq 0 -or $idxWL2 -eq $totalWL2) {
            Write-Progress -Id 4 -Activity "Génération SIMP" `
                -Status ("Worklog " + $idxWL2 + " / " + $totalWL2 + " — " + $simpGrouped.Count + " groupes") `
                -PercentComplete ([math]::Round($idxWL2 / $totalWL2 * 100))
        }

        $accId  = [string]$wl.author.accountId
        $issId  = [string]$wl.issue.id
        $jUser  = $jiraUserCache[$accId]
        $uMail  = if ($jUser -and $jUser.EmailAddress) { [string]$jUser.EmailAddress } else { "" }
        $dName  = if ($jUser -and $jUser.DisplayName)  { [string]$jUser.DisplayName }  else { "" }

        $cmdb   = Resolve-Cmdb -AccId $accId -UserMail $uMail -DispName $dName
        $jIssue = $jiraIssueCache[[string]$issId]

        $codeSimpRessource = if ($cmdb -and $cmdb.CodeDomaineSIMP) { $cmdb.CodeDomaineSIMP } else { "AUTRE" }
        $typeRes           = if ($cmdb -and $cmdb.TypeRessource)   { $cmdb.TypeRessource }   else { "Régie" }
        $budgetKey         = if ($jIssue -and $jIssue.BudgetKey)     { [string]$jIssue.BudgetKey }     else { "" }
        $budgetSummaryRaw  = if ($jIssue -and $jIssue.BudgetSummary) { [string]$jIssue.BudgetSummary } else { "" }
        $budgetSumm        = Resolve-LibelleSIMP -BudgetKey $budgetKey -BudgetSummary $budgetSummaryRaw
        $heures            = [double]$wl.timeSpentSeconds / 3600.0

        $groupKey = ($codeSimpRessource + "|" + $budgetKey + "|" + $typeRes)
        if (-not $simpGrouped.ContainsKey($groupKey)) {
            $simpGrouped[$groupKey] = [ordered]@{
                CodeSimpRessource = $codeSimpRessource
                BudgetKey         = $budgetKey
                BudgetSummary     = $budgetSumm
                TypeRessource     = $typeRes
                TotalHeures       = 0.0
            }
        }
        if ([string]::IsNullOrWhiteSpace($simpGrouped[$groupKey]["BudgetSummary"]) -and -not [string]::IsNullOrWhiteSpace($budgetSumm)) {
            $simpGrouped[$groupKey]["BudgetSummary"] = $budgetSumm
        }
        $simpGrouped[$groupKey]["TotalHeures"] += $heures
    }
    Write-Progress -Id 4 -Activity "Génération SIMP" -Completed
    Write-Info ("  SIMP : " + $simpGrouped.Count + " groupes d'agrégation générés")

    # Génération CSV SIMP
    $simpRows = New-Object System.Collections.Generic.List[object]
    foreach ($gKey in $simpGrouped.Keys) {
        $sg        = $simpGrouped[$gKey]
        $totHeures = [math]::Round($sg.TotalHeures, 5)
        $totJours  = [math]::Round($sg.TotalHeures / 8.0, 5)

        $simpRows.Add([pscustomobject]@{
            idCle             = ($sg.CodeSimpRessource + "_" + $sg.BudgetKey + "_" + $sg.TypeRessource + "_" + $fyKey + "_" + $mKey)
            codeSIMPRessource = $sg.CodeSimpRessource
            codeBudget        = $sg.BudgetKey
            libelleBudget     = $sg.BudgetSummary
            typeRessource     = $sg.TypeRessource
            exercice          = $fyKey
            mois              = $mKey
            heuresTotales     = $totHeures
            joursTotaux       = $totJours
        }) | Out-Null
    }

    $csvSimp = Join-Path $exportsDir ("EtatsFinanciers_SIMP_" + $moisKey + "_" + $runStamp + ".csv")
    Export-CsvStrict -Path $csvSimp `
        -Headers @("idCle","codeSIMPRessource","codeBudget","libelleBudget","typeRessource","exercice","mois","heuresTotales","joursTotaux") `
        -Rows $simpRows

    # Génération XLSX EXPORT_JIRADOT_SIMP
    $xlsxSimpRows = New-Object System.Collections.Generic.List[object]
    foreach ($gKey in $simpGrouped.Keys) {
        $sg       = $simpGrouped[$gKey]
        $totJours = [math]::Round($sg.TotalHeures / 8.0, 2)

        $activiteInfo = if ($sg.BudgetKey -and $sg.BudgetSummary) {
            $sg.BudgetKey + " — " + $sg.BudgetSummary
        } elseif ($sg.BudgetKey) { $sg.BudgetKey } else { "(Sans budget)" }

        $xlsxSimpRows.Add([pscustomobject]@{
            "Contribution"          = $sg.CodeSimpRessource
            "Activité informatique" = $activiteInfo
            "Domaine applicatif"    = $sg.TypeRessource
            "Indicateur"            = "Nb jours"
            "Année"                 = $From.Year
            "Période"               = "M" + $From.Month.ToString("D2") + "/" + $From.Year
            "Données"               = $totJours
        }) | Out-Null
    }

    $xlsxSimpHeaders = @(
        "Contribution","Activité informatique","Domaine applicatif",
        "Indicateur","Année","Période","Données"
    )
    $xlsxSimpFile = Join-Path $exportsDir ("EXPORT_JIRADOT_SIMP_" + $runStamp + ".xlsx")
    Export-GenericXlsx -Path $xlsxSimpFile -SheetName "SIMP" -Headers $xlsxSimpHeaders -Rows $xlsxSimpRows
    Write-Info ("  SIMP XLSX : " + $xlsxSimpRows.Count + " lignes générées")
}

# ============================================================
# 10. RÉSUMÉ FINAL ET CONTRÔLE DE COUVERTURE CMDB
# ============================================================
Write-Info ""
Write-Info "=========================================="
Write-Info "=== EXTRACTION TERMINÉE — RÉSUMÉ      ==="
Write-Info "=========================================="
Write-Info ("  Période          : " + $From.ToString("dd/MM/yyyy") + " -> " + $To.ToString("dd/MM/yyyy"))
Write-Info ("  Périmètre        : " + $Format)
Write-Info ("  Worklogs Tempo   : " + $worklogs.Count)
Write-Info ("  Comptes Jira     : " + $jiraUserCache.Count)
Write-Info ("  Tickets Résolus  : " + $jiraIssueCache.Count)
Write-Info ("  Libellés Budget  : " + $global:budgetLibelleCache.Count + " résolus via fallback API")
Write-Info ("  Fiches Assets    : " + $assetsUsersByAccountId.Count + " (accountId) | " +
                                      $assetsUsersByEmail.Count     + " (email) | " +
                                      $assetsUsersByName.Count      + " (nom)")
Write-Info ""
Write-Info ("  Fichiers générés dans : " + $exportsDir)

if ($Format -eq "DETERMINE" -or $Format -eq "BOTH") {
    Write-Info ("    - EtatsFinanciers_DETERMINE_" + $moisKey + "_" + $runStamp + ".csv")
    Write-Info ("    - EXPORT_JIRADOT_SUIVI-ACTIVITE_" + $runStamp + ".xlsx")
}
if ($Format -eq "SIMP" -or $Format -eq "BOTH") {
    Write-Info ("    - EtatsFinanciers_SIMP_" + $moisKey + "_" + $runStamp + ".csv")
    Write-Info ("    - EXPORT_JIRADOT_SIMP_" + $runStamp + ".xlsx")
}

Write-Info ""
Write-Info ("  Cache local dans  : " + $cacheDir)
Write-Info ("    - worklogs_tempo_" + $fromStr + "_" + $toStr + ".json")
Write-Info ("    - assets_referentiel_personne.json")
Write-Info ("    - jira_users_cache.json")
Write-Info ("    - jira_issues_hierarchy.json")
Write-Info ""

# Contrôle de couverture CMDB
$matchedById    = 0
$matchedByEmail = 0
$matchedByName  = 0
$notMatched     = 0

foreach ($wl in $worklogs) {
    $accId = [string]$wl.author.accountId
    $jUser = $jiraUserCache[$accId]
    $uMail = if ($jUser -and $jUser.EmailAddress) { [string]$jUser.EmailAddress } else { "" }
    $dName = if ($jUser -and $jUser.DisplayName)  { [string]$jUser.DisplayName }  else { "" }

    if ($accId -and $assetsUsersByAccountId.ContainsKey($accId)) {
        $matchedById++
    } elseif ($uMail -and $assetsUsersByEmail.ContainsKey($uMail)) {
        $matchedByEmail++
    } elseif ($dName -and $assetsUsersByName.ContainsKey($dName.Trim().ToLower())) {
        $matchedByName++
    } else {
        $notMatched++
        Write-Log ("Non rattaché CMDB : accountId=" + $accId + " | displayName=" + $dName + " | email=" + $uMail) "WARN"
    }
}

$totalWorklogs = $worklogs.Count
$pctCoverage   = if ($totalWorklogs -gt 0) {
    [math]::Round(($totalWorklogs - $notMatched) / $totalWorklogs * 100, 1)
} else { 0 }

Write-Info ("  Couverture CMDB RP :")
Write-Info ("    Clé #1 accountId   : " + $matchedById    + " worklogs")
Write-Info ("    Clé #2 email       : " + $matchedByEmail + " worklogs")
Write-Info ("    Clé #3 nom         : " + $matchedByName  + " worklogs")
Write-Info ("    Non rattachés      : " + $notMatched      + " worklogs")
Write-Info ("    Taux de couverture : " + $pctCoverage + "%")

if ($notMatched -gt 0) {
    Write-Warn ("  " + $notMatched + " worklog(s) non rattaché(s) à la CMDB RP — détail dans : " + $logFile)
}

Write-Info ""
Write-Info ("  Log complet dans  : " + $logFile)
Write-Info "=========================================="
Write-Log ("=== FIN EXÉCUTION " + $scriptName + " ===") "INFO"
