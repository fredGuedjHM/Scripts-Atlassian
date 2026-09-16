<#
.SYNOPSIS
    Extraction et consolidation des etats financiers Jiradot (Tempo + Jira + CMDB Assets).
.DESCRIPTION
    Script de consolidation v3.0 avec interface graphique WPF :
    - Choix de la periode (24 mois glissants)
    - Choix du perimetre (DETERMINE, SIMP, LES DEUX)
    - Choix des formats de sortie (Excel XLSX, CSV Plat, JSON Assets EXCOMP)
.VERSION
    3.0 — 2026-09-11
.NOTES
    Auteur  : DSIM / HM_DSIM_PACT
    Sorties : CSV (DETERMINE, SIMP) + XLSX (SUIVI-ACTIVITE, SIMP) + JSON (Assets EXCOMP)
#>
[CmdletBinding()]
param(
    [DateTime]$From,
    [DateTime]$To,
    [ValidateSet("DETERMINE","SIMP","BOTH")]
    [string]$Format,
    [string[]]$ExportFormats = @("XLSX","CSV","JSON"),
    [switch]$SkipCachePrompt
)

# Culture FR — séparateur décimal virgule pour les exports CSV
[System.Threading.Thread]::CurrentThread.CurrentCulture   = [System.Globalization.CultureInfo]::GetCultureInfo("fr-FR")
[System.Threading.Thread]::CurrentThread.CurrentUICulture = [System.Globalization.CultureInfo]::GetCultureInfo("fr-FR")

Set-StrictMode -Off
$ErrorActionPreference = "Continue"

# ============================================================
# 0. CONSTANTES ET CHEMINS
# ============================================================
$scriptName  = "Invoke-EtatsFinanciers_v3"
$scriptRoot  = Split-Path -Parent $MyInvocation.MyCommand.Path
if (-not $scriptRoot) { $scriptRoot = Get-Location }

$exportsDir  = Join-Path $scriptRoot "exports"
$cacheDir    = Join-Path $scriptRoot "cache"
$secretsDir  = Join-Path $scriptRoot "secrets"
$logsDir     = Join-Path $scriptRoot "logs"
$runStamp    = Get-Date -Format "yyyyMMdd_HHmmss"
$logFile     = Join-Path $logsDir ($scriptName + "_" + $runStamp + ".log")

foreach ($d in @($exportsDir,$cacheDir,$secretsDir,$logsDir)) {
    if (-not (Test-Path $d)) { New-Item -ItemType Directory -Path $d -Force | Out-Null }
}

$AssetsWorkspaceId          = "2898703d-5f14-4527-9f8a-ad32b9279ca1"
$global:assetsApiDelayMs    = 150
$global:budgetLibelleCache  = New-Object System.Collections.Hashtable

$ProxyUrl                   = ""
$UseSystemProxy             = $true
$ProxyUseDefaultCredentials = $true

# ============================================================
# 1. LOGGING ET HELPERS ENCODAGE
# ============================================================
Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem

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
function Write-ErrLog([string]$Message) {
    Write-Error $Message
    Write-Log $Message "ERROR"
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

function ConvertTo-DateFR([string]$val) {
    if ([string]::IsNullOrWhiteSpace($val)) { return "" }
    try { return ([DateTime]::Parse($val)).ToString("dd/MM/yyyy") } catch { return $val }
}

# ============================================================
# 2. HELPERS CACHE JSON
# ============================================================
function Load-JsonCache([string]$Path, [int]$MaxAgeHours = 168) {
    if (-not (Test-Path $Path)) { return $null }
    $age = (Get-Date) - (Get-Item $Path).LastWriteTime
    if ($age.TotalHours -gt $MaxAgeHours) {
        Write-Warn ("Cache expire (" + [math]::Round($age.TotalHours,1) + "h > " + $MaxAgeHours + "h) : " + (Split-Path $Path -Leaf))
        return $null
    }
    try {
        $raw = [System.IO.File]::ReadAllText($Path, [System.Text.Encoding]::UTF8)
        return ($raw | ConvertFrom-Json)
    } catch { return $null }
}
function Save-JsonCache([string]$Path, $Object) {
    try {
        $json      = ($Object | ConvertTo-Json -Depth 10)
        $utf8NoBom = New-Object System.Text.UTF8Encoding($false)
        [System.IO.File]::WriteAllText($Path, $json, $utf8NoBom)
    } catch { Write-Warn ("Erreur sauvegarde cache JSON : " + $_.Exception.Message) }
}
function Test-CacheValide([string]$Path, [int]$MaxAgeHours = 168) {
    if (-not (Test-Path $Path)) { return $false }
    return (((Get-Date) - (Get-Item $Path).LastWriteTime).TotalHours -le $MaxAgeHours)
}
function Get-CacheAge([string]$Path) {
    if (-not (Test-Path $Path)) { return -1 }
    return [math]::Round(((Get-Date) - (Get-Item $Path).LastWriteTime).TotalHours, 1)
}

# ============================================================
# 3. HELPERS API REST & ASSETS AQL
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

function Invoke-ApiPostUtf8([string]$Url, [hashtable]$Headers, [string]$JsonBody) {
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    $bodyBytes = [System.Text.Encoding]::UTF8.GetBytes($JsonBody)
    $params    = @{ Method='POST'; Uri=$Url; Headers=$Headers; Body=$bodyBytes; ContentType='application/json'; UseBasicParsing=$true; ErrorAction='Stop' }
    $px = Get-ProxyParams $Url; foreach ($k in $px.Keys) { $params[$k] = $px[$k] }
    $resp   = Invoke-WebRequest @params
    $stream = $resp.RawContentStream; $stream.Position = 0
    $reader = New-Object System.IO.StreamReader($stream, [System.Text.Encoding]::UTF8, $true)
    $raw    = $reader.ReadToEnd(); $reader.Close()
    return ($raw | ConvertFrom-Json)
}

function Invoke-AssetsAql {
    param(
        [Parameter(Mandatory=$true)][string]$AqlQuery,
        [Parameter(Mandatory=$true)][hashtable]$Headers,
        [Parameter(Mandatory=$true)][string]$AqlUrl,
        [int]$MaxResults       = 100,
        [int]$ProgressId       = 0,
        [string]$ProgressLabel = "Chargement Assets",
        [switch]$IncludeAttributes
    )
    $results  = New-Object System.Collections.ArrayList
    $startAt  = 0
    $isLast   = $false
    $pageNum  = 0
    $inclAttr = if ($IncludeAttributes) { "true" } else { "false" }

    do {
        $pageNum++
        Write-Progress -Id $ProgressId -Activity $ProgressLabel `
            -Status ("Page " + $pageNum.ToString() + " — " + $results.Count.ToString() + " objets charges...")

        $urlAql  = $AqlUrl + "?startAt=" + $startAt + "&maxResults=" + $MaxResults + "&includeAttributes=" + $inclAttr
        $bodyAql = (@{ qlQuery = $AqlQuery } | ConvertTo-Json -Depth 3)
        $respAql = $null; $tentative = 0

        while ($tentative -lt 3 -and $null -eq $respAql) {
            $tentative++
            try {
                $respAql = Invoke-ApiPostUtf8 -Url $urlAql -Headers $Headers -JsonBody $bodyAql
            } catch {
                $errMsg     = $_.Exception.Message
                $isThrottle = ($errMsg -ilike "*429*" -or $errMsg -ilike "*too many*" -or
                               $errMsg -ilike "*throttl*" -or $errMsg -ilike "*timeout*")
                if ($tentative -lt 3) {
                    $waitMs = if ($isThrottle) { 2000 * $tentative } else { 500 * $tentative }
                    if ($isThrottle) { $global:assetsApiDelayMs = [math]::Min($global:assetsApiDelayMs * 2, 2000) }
                    Write-Warn ("  Retry " + $tentative + "/3 - attente " + ($waitMs/1000) + "s : " + $errMsg)
                    Start-Sleep -Milliseconds $waitMs
                } else {
                    Write-Warn ("  Echec definitif page " + $pageNum + " : " + $errMsg)
                    $isLast = $true
                }
            }
        }

        if ($null -eq $respAql -or -not $respAql.values) { break }

        if (-not $script:lastAqlAttrDict) { $script:lastAqlAttrDict = @{} }
        if ($respAql.objectTypeAttributes) {
            foreach ($ota in $respAql.objectTypeAttributes) {
                if ($ota.id -and $ota.name) {
                    $script:lastAqlAttrDict[[string]$ota.id] = Fix-Encoding ([string]$ota.name)
                }
            }
        }
        foreach ($obj in $respAql.values) { [void]$results.Add($obj) }

        $isLast   = if ($null -ne $respAql.isLast) { [bool]$respAql.isLast } else { $true }
        $startAt += $MaxResults
        if (-not $isLast -and $global:assetsApiDelayMs -gt 0) {
            Start-Sleep -Milliseconds $global:assetsApiDelayMs
        }
    } while (-not $isLast)

    Write-Progress -Id $ProgressId -Activity $ProgressLabel -Completed
    return ,$results
}

# ============================================================
# 4. INTERFACE UTILISATEUR (Dialogs WPF v3.0)
# ============================================================
function Show-EtatsFinanciersDialog {
    Add-Type -AssemblyName PresentationFramework
    $now = Get-Date
    $xaml = @"
<Window
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    Title="Extraction Etats Financiers v3.0" Width="460" Height="360"
    WindowStartupLocation="CenterScreen" ResizeMode="NoResize">
  <StackPanel Margin="20">
    <TextBlock Text="1. Mois d extraction :" FontWeight="Bold" Margin="0,0,0,5"/>
    <ComboBox x:Name="cbMois" Width="400" Height="28" HorizontalAlignment="Left" Margin="0,0,0,15"/>
    <TextBlock Text="2. Perimetre a generer :" FontWeight="Bold" Margin="0,0,0,5"/>
    <StackPanel Orientation="Horizontal" Margin="0,0,0,15">
      <RadioButton x:Name="rbDetermine" Content="DETERMINE"  Margin="0,0,15,0"/>
      <RadioButton x:Name="rbSimp"      Content="SIMP"        Margin="0,0,15,0"/>
      <RadioButton x:Name="rbBoth"      Content="LES DEUX"    IsChecked="True"/>
    </StackPanel>
    <TextBlock Text="3. Formats de sortie a generer :" FontWeight="Bold" Margin="0,0,0,5"/>
    <StackPanel Orientation="Vertical" Margin="10,0,0,15">
      <CheckBox x:Name="cbXlsx" Content="Excel (.xlsx) - Visualisation et controle" IsChecked="True" Margin="0,3"/>
      <CheckBox x:Name="cbCsv"  Content="CSV (.csv) - Format plat donnees"          IsChecked="True" Margin="0,3"/>
      <CheckBox x:Name="cbJson" Content="JSON (Assets EXCOMP) - Import CMDB Export"  IsChecked="True" Margin="0,3"/>
    </StackPanel>
    <Button x:Name="btnOk" Content="Lancer l extraction" Width="220" Height="35"
            Margin="0,5,0,0" HorizontalAlignment="Center" FontWeight="Bold"/>
  </StackPanel>
</Window>
"@
    $reader = New-Object System.Xml.XmlNodeReader([xml]$xaml)
    $win    = [Windows.Markup.XamlReader]::Load($reader)
    $cbMois     = $win.FindName("cbMois")
    $monthsList = New-Object System.Collections.ArrayList
    for ($i = 0; $i -lt 24; $i++) {
        $mDate    = $now.AddMonths(-$i)
        $culture  = [System.Globalization.CultureInfo]::GetCultureInfo("fr-FR")
        $mName    = $culture.DateTimeFormat.GetMonthName($mDate.Month)
        $mName    = $mName.Substring(0,1).ToUpper() + $mName.Substring(1)
        $label    = $mDate.ToString("MM/yyyy") + "  -  " + $mName + " " + $mDate.Year
        $firstDay = Get-Date -Year $mDate.Year -Month $mDate.Month -Day 1 -Hour 0 -Minute 0 -Second 0
        $lastDay  = $firstDay.AddMonths(1).AddDays(-1)
        [void]$cbMois.Items.Add($label)
        [void]$monthsList.Add([pscustomobject]@{ FirstDay=$firstDay; LastDay=$lastDay })
    }
    $cbMois.SelectedIndex = 1
    $result = @{ From=$null; To=$null; Format="BOTH"; ExportFormats=@() }
    $win.FindName("btnOk").Add_Click({
        $idx = $cbMois.SelectedIndex; if ($idx -lt 0) { $idx = 1 }
        $sel           = $monthsList[$idx]
        $result.From   = $sel.FirstDay
        $result.To     = $sel.LastDay
        $result.Format = if ($win.FindName("rbDetermine").IsChecked) { "DETERMINE" }
                         elseif ($win.FindName("rbSimp").IsChecked)  { "SIMP" }
                         else { "BOTH" }
        $expFmts = New-Object System.Collections.ArrayList
        if ($win.FindName("cbXlsx").IsChecked) { [void]$expFmts.Add("XLSX") }
        if ($win.FindName("cbCsv").IsChecked)  { [void]$expFmts.Add("CSV") }
        if ($win.FindName("cbJson").IsChecked) { [void]$expFmts.Add("JSON") }
        if ($expFmts.Count -eq 0) { [void]$expFmts.Add("XLSX"); [void]$expFmts.Add("CSV"); [void]$expFmts.Add("JSON") }
        $result.ExportFormats = @($expFmts)
        $win.DialogResult = $true; $win.Close()
    })
    [void]$win.ShowDialog()
    return [pscustomobject]$result
}

function Show-CacheRefreshDialog {
    Add-Type -AssemblyName PresentationFramework
    $xaml = @"
<Window
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    Title="Options de cache" Width="400" Height="270"
    WindowStartupLocation="CenterScreen" ResizeMode="NoResize">
  <StackPanel Margin="20">
    <TextBlock Text="Forcer le rechargement du cache :" FontWeight="Bold" Margin="0,0,0,10"/>
    <CheckBox x:Name="cbWorklogs" Content="Worklogs Tempo (mois selectionne)"  Margin="0,5"/>
    <CheckBox x:Name="cbAssets"   Content="CMDB Assets (RP + SIMP)"            Margin="0,5"/>
    <CheckBox x:Name="cbIssues"   Content="Hierarchie Tickets Jira"             Margin="0,5"/>
    <CheckBox x:Name="cbUsers"    Content="Profils Utilisateurs Jira"           Margin="0,5"/>
    <TextBlock Text="Laisser decoche = utiliser le cache local existant"
               FontStyle="Italic" Foreground="Gray" Margin="0,10,0,10" FontSize="11"/>
    <Button x:Name="btnOk" Content="OK" Width="120" Height="30" HorizontalAlignment="Center"/>
  </StackPanel>
</Window>
"@
    $reader = New-Object System.Xml.XmlNodeReader([xml]$xaml)
    $win    = [Windows.Markup.XamlReader]::Load($reader)
    $result = @{ RefreshWorklogs=$false; RefreshAssets=$false; RefreshIssues=$false; RefreshUsers=$false }
    $win.FindName("btnOk").Add_Click({
        $result.RefreshWorklogs = [bool]$win.FindName("cbWorklogs").IsChecked
        $result.RefreshAssets   = [bool]$win.FindName("cbAssets").IsChecked
        $result.RefreshIssues   = [bool]$win.FindName("cbIssues").IsChecked
        $result.RefreshUsers    = [bool]$win.FindName("cbUsers").IsChecked
        $win.DialogResult = $true; $win.Close()
    })
    [void]$win.ShowDialog()
    return [pscustomobject]$result
}

# ============================================================
# 5. FONCTIONS D'EXPORTATION (CSV, XLSX, JSON ASSETS EXCOMP)
# ============================================================
function Export-CsvStrict {
    param([string]$Path, [string[]]$Headers, $Rows)
    $utf8Bom = New-Object System.Text.UTF8Encoding($true)
    $writer  = New-Object System.IO.StreamWriter($Path, $false, $utf8Bom)
    try {
        $writer.WriteLine(($Headers -join ";"))
        foreach ($row in $Rows) {
            $vals = New-Object System.Collections.ArrayList
            foreach ($h in $Headers) {
                $raw = $row.$h
                if ($raw -is [double] -or $raw -is [float] -or $raw -is [decimal]) {
                    $s = $raw.ToString("G", [System.Globalization.CultureInfo]::GetCultureInfo("fr-FR"))
                } else { $s = [string]$raw }
                if ($s.Contains(";") -or $s.Contains('"') -or $s.Contains("`n")) {
                    [void]$vals.Add('"' + $s.Replace('"','""') + '"')
                } else { [void]$vals.Add($s) }
            }
            $writer.WriteLine(($vals -join ";"))
        }
    } finally { $writer.Close(); $writer.Dispose() }
    Write-Info ("  -> CSV genere : " + $Path + " (" + $Rows.Count + " lignes)")
}

function Export-CsvToXlsx {
    param(
        [Parameter(Mandatory=$true)][string]$CsvPath,
        [Parameter(Mandatory=$true)][string]$XlsxPath,
        [string]$SheetName = "Feuille1"
    )
    if (-not (Test-Path -LiteralPath $CsvPath)) { Write-Warn ("CSV source introuvable : " + $CsvPath); return }
    $fullCsvPath  = (Get-Item -LiteralPath $CsvPath).FullName
    $fullXlsxPath = [System.IO.Path]::GetFullPath($XlsxPath)
    if (Get-Module -ListAvailable -Name Import-Excel) {
        try {
            Import-Module Import-Excel -ErrorAction Stop
            Import-Csv -Path $fullCsvPath -Delimiter ";" |
                Export-Excel -Path $fullXlsxPath -WorksheetName $SheetName -AutoSize -TableStyle Medium6 -Show:$false
            Write-Info ("  -> XLSX converti (Import-Excel) : " + $XlsxPath); return
        } catch { Write-Warn ("Import-Excel erreur : " + $_.Exception.Message) }
    }
    try {
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false; $excel.DisplayAlerts = $false
        $excel.Workbooks.OpenText($fullCsvPath, 65001, 1, 1, 1, $false, $false, $true, $false, $false, $false)
        $wb = $excel.ActiveWorkbook; $ws = $wb.Worksheets.Item(1)
        try { $ws.Name = $SheetName } catch {}
        $usedCols    = $ws.UsedRange.Columns.Count
        $headerRange = $ws.Range($ws.Cells.Item(1,1), $ws.Cells.Item(1,$usedCols))
        $headerRange.Font.Bold = $true; $headerRange.Font.ColorIndex = 2; $headerRange.Interior.ColorIndex = 23
        $ws.Columns.AutoFit() | Out-Null
        if (Test-Path -LiteralPath $fullXlsxPath) { Remove-Item -LiteralPath $fullXlsxPath -Force -ErrorAction SilentlyContinue }
        $wb.SaveAs($fullXlsxPath, 51); $wb.Close($false); $excel.Quit()
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
        Write-Info ("  -> XLSX converti (Excel COM) : " + $XlsxPath); return
    } catch { Write-Warn ("Excel COM erreur : " + $_.Exception.Message) }
    Write-ErrLog ("Impossible de convertir le CSV en XLSX : " + $XlsxPath)
}

function Export-GenericXlsx {
    param(
        [Parameter(Mandatory=$true)][string]$Path,
        [Parameter(Mandatory=$true)][string]$SheetName,
        [Parameter(Mandatory=$true)][string[]]$Headers,
        [Parameter(Mandatory=$true)]$Rows
    )
    if ($Rows.Count -eq 0) { Write-Warn ("  -> Aucune ligne a exporter : " + $Path); return }
    $tempCsv = Join-Path $env:TEMP ("temp_xlsx_" + [System.Guid]::NewGuid().ToString("N") + ".csv")
    try {
        Export-CsvStrict -Path $tempCsv -Headers $Headers -Rows $Rows
        Export-CsvToXlsx -CsvPath $tempCsv -XlsxPath $Path -SheetName $SheetName
    } finally {
        if (Test-Path -LiteralPath $tempCsv) { Remove-Item -LiteralPath $tempCsv -Force -ErrorAction SilentlyContinue }
    }
}

function Export-JsonAssetsExcomp {
    param(
        [Parameter(Mandatory=$true)][string]$Path,
        [Parameter(Mandatory=$true)]$DataArray
    )
    $payload   = [ordered]@{ "export" = $DataArray }
    $json      = $payload | ConvertTo-Json -Depth 10
    $utf8NoBom = New-Object System.Text.UTF8Encoding($false)
    [System.IO.File]::WriteAllText($Path, $json, $utf8NoBom)
    Write-Info ("  -> JSON Assets EXCOMP genere : " + $Path + " (" + $DataArray.Count + " objets)")
}

# ============================================================
# 6. INITIALISATION PÉRIODE, CREDENTIALS ET CACHE CHOICES
# ============================================================
if (-not $From -or -not $To -or -not $Format) {
    $uiParams      = Show-EtatsFinanciersDialog
    $From          = $uiParams.From
    $To            = $uiParams.To
    $Format        = $uiParams.Format
    $ExportFormats = $uiParams.ExportFormats
}

$fromStr = $From.ToString("yyyy-MM-dd")
$toStr   = $To.ToString("yyyy-MM-dd")
$moisKey = $From.ToString("MM-yyyy")
$fyKey   = "FY" + $From.ToString("yy")
$mKey    = "M"  + $From.ToString("MM")

Write-Info ("Periode          : " + $From.ToString("dd/MM/yyyy") + " -> " + $To.ToString("dd/MM/yyyy"))
Write-Info ("Perimetre        : " + $Format)
Write-Info ("Formats de sortie: " + ($ExportFormats -join ", "))

if (-not $SkipCachePrompt) {
    $cacheChoices = Show-CacheRefreshDialog
} else {
    $cacheChoices = [pscustomobject]@{ RefreshWorklogs=$false; RefreshAssets=$false; RefreshIssues=$false; RefreshUsers=$false }
}

$jiraCredFile = Join-Path $secretsDir "jira-jiradot.cred.xml"
if (-not (Test-Path $jiraCredFile)) { throw "Fichier Jira creds introuvable : $jiraCredFile" }
$jiraData    = Import-Clixml -Path $jiraCredFile
$jiraBaseUrl = ([string]$jiraData.JiraBaseUrl).TrimEnd("/")
$jiraCred    = $jiraData.Credential
$b64         = [Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes($jiraCred.UserName + ":" + $jiraCred.GetNetworkCredential().Password))
$jiraHeaders = @{ Authorization = "Basic " + $b64; Accept = "application/json" }

$tempoTokenFile = Join-Path $secretsDir "tempo-token.xml"
if (-not (Test-Path $tempoTokenFile)) { throw "Fichier Tempo token introuvable : $tempoTokenFile" }
$tempoData    = Import-Clixml -Path $tempoTokenFile
$tempoToken   = [Runtime.InteropServices.Marshal]::PtrToStringAuto(
                    [Runtime.InteropServices.Marshal]::SecureStringToBSTR($tempoData.Token))
$tempoHeaders = @{ Authorization = "Bearer " + $tempoToken; Accept = "application/json" }

$assetsAqlUrl    = $jiraBaseUrl + "/gateway/api/jsm/assets/workspace/" + $AssetsWorkspaceId + "/v1/object/aql"
$assetsSchemaUrl = $jiraBaseUrl + "/gateway/api/jsm/assets/workspace/" + $AssetsWorkspaceId + "/v1/objectschema/list"

Write-Log ("=== DEBUT EXECUTION " + $scriptName + " v3.0 — " + $runStamp + " ===") "INFO"
# ============================================================
# 7a. CMDB ASSETS — RP (Référentiel Personne - Chargement)
# ============================================================
Write-Info "=== 1a/4 Chargement CMDB Assets — RP (Referentiel Personne) ==="

$assetsCacheFile        = Join-Path $cacheDir "assets_referentiel_personne.json"
$assetsUsersByAccountId = @{}
$assetsUsersByEmail     = @{}
$assetsUsersByName      = @{}
$assetsList             = $null

$rpAge = Get-CacheAge $assetsCacheFile
if ($rpAge -ge 0) { Write-Info ("  Cache RP : age = " + $rpAge + "h") }

if (-not $cacheChoices.RefreshAssets -and (Test-CacheValide $assetsCacheFile -MaxAgeHours 168)) {
    $assetsList = Load-JsonCache -Path $assetsCacheFile -MaxAgeHours 168
    if ($assetsList) { Write-Info ("  -> RP charge depuis le cache (" + $assetsList.Count + " fiches — 0 appel API)") }
}

if (-not $assetsList) {
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
    Write-Info ("  -> Schema RP : ID=" + $rpSchemaId + " | Delai : " + $global:assetsApiDelayMs + "ms")

    $script:lastAqlAttrDict = @{}
    $rawObjects = Invoke-AssetsAql `
        -AqlQuery          ("objectSchemaId = " + $rpSchemaId) `
        -Headers           $jiraHeaders `
        -AqlUrl            $assetsAqlUrl `
        -MaxResults        100 `
        -ProgressId        0 `
        -ProgressLabel     "Chargement CMDB Assets RP" `
        -IncludeAttributes

    $assetsList = New-Object System.Collections.ArrayList
    foreach ($pObj in $rawObjects) {
        if (-not $pObj.id) { continue }

        $accId=""; $nomPrenom=""; $statut="Actif"; $typeRes="Prestataire"
        $codeSimp=""; $directionStr=""; $serviceStr=""
        $matricule=""; $societe=""; $codeCigref=""; $compteJiraRaw=""
        $attr375=""; $attr376=""

        $labelAssets = if ($pObj.label) { Fix-Encoding ([string]$pObj.label) } else { "" }
        $objTypeName = if ($pObj.objectType -and $pObj.objectType.name) { Fix-Encoding ([string]$pObj.objectType.name) } else { "" }

        foreach ($attr in $pObj.attributes) {
            $attrId = [string]$attr.objectTypeAttributeId
            $vals   = $attr.objectAttributeValues
            if (-not $vals -or $vals.Count -eq 0) { continue }
            $vVal  = Fix-Encoding ([string]$vals[0].value)
            $vDisp = Fix-Encoding ([string]$vals[0].displayValue)
            $vBest = if (-not [string]::IsNullOrWhiteSpace($vDisp)) { $vDisp } else { $vVal }

            switch ($attrId) {
                "375" { $attr375  = $vBest }
                "376" { $attr376  = $vBest }
                "407" { if ([string]::IsNullOrWhiteSpace($nomPrenom)) { $nomPrenom = $vBest } }
                "378" { $codeSimp = $vBest }
                "383" {
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
                    $aName = $script:lastAqlAttrDict[$attrId]
                    if (-not $aName -and $attr.objectTypeAttribute -and $attr.objectTypeAttribute.name) {
                        $aName = Fix-Encoding ([string]$attr.objectTypeAttribute.name)
                    }
                    if (-not $aName) { continue }
                    if ($aName -ieq "Compte Jira" -or $aName -ilike "*AccountId*" -or $aName -ilike "*User*ID*") {
                        if ([string]::IsNullOrWhiteSpace($accId)) {
                            foreach ($v in $vals) {
                                if ($v.user -and $v.user.accountId) { $accId = [string]$v.user.accountId; break }
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
                    elseif ($aName -ieq "Statut")               { $statut       = $vBest }
                    elseif ($aName -ilike "*Type ressource*")    { $typeRes      = $vBest }
                    elseif ($aName -ieq "Matricule")            { $matricule    = $vBest }
                    elseif ($aName -ieq "Direction")            { $directionStr = $vBest }
                    elseif ($aName -ilike "*Affectation #2*")   { $serviceStr   = $vBest }
                    elseif ($aName -ilike "*CIGREF*")           { $codeCigref   = $vBest }
                }
            }
        }

        if (-not [string]::IsNullOrWhiteSpace($attr375) -or -not [string]::IsNullOrWhiteSpace($attr376)) {
            $nomPrenom = ($attr375.Trim() + " " + $attr376.Trim()).Trim()
        }
        if ([string]::IsNullOrWhiteSpace($nomPrenom)) { $nomPrenom = $labelAssets }

        if ([string]::IsNullOrWhiteSpace($societe)) {
            if ($objTypeName -ilike "*Employe*" -or $objTypeName -ilike "*Employé*" -or $typeRes -ieq "Interne") {
                $societe = "Harmonie Mutuelle"
            } else {
                $societe = "(Société non renseignée)"
            }
        }

        $extractedEmail = ""
        if ($compteJiraRaw -match "\(([^)]+@[^)]+)\)") { $extractedEmail = $matches[1].Trim().ToLower() }

        [void]$assetsList.Add([pscustomobject]@{
            Label           = $labelAssets
            AccountId       = $accId
            ExtractedEmail  = $extractedEmail
            Statut          = $statut
            TypeRessource   = $typeRes
            CodeDomaineSIMP = $codeSimp
            NomPrenom       = $nomPrenom
            NOM             = $attr375
            Prenom          = $attr376
            Matricule       = $matricule
            Direction       = $directionStr
            Service         = $serviceStr
            Societe         = $societe
            CodeCIGREF      = $codeCigref
        })
    }
    Save-JsonCache -Path $assetsCacheFile -Object $assetsList
    Write-Info ("  -> " + $assetsList.Count + " fiches RP sauvegardees dans le cache")
}

# ============================================================
# 7b. CMDB ASSETS — RP (Indexation et contrôle qualité)
# ============================================================
Write-Info "=== 1b/4 Indexation et controle qualite RP ==="

foreach ($uRec in $assetsList) {
    if (-not [string]::IsNullOrWhiteSpace($uRec.AccountId))      { $assetsUsersByAccountId[$uRec.AccountId]            = $uRec }
    if (-not [string]::IsNullOrWhiteSpace($uRec.ExtractedEmail)) { $assetsUsersByEmail[$uRec.ExtractedEmail.ToLower()]  = $uRec }
    if (-not [string]::IsNullOrWhiteSpace($uRec.NomPrenom))      { $assetsUsersByName[$uRec.NomPrenom.ToLower().Trim()] = $uRec }
    if (-not [string]::IsNullOrWhiteSpace($uRec.Label))          { $assetsUsersByName[$uRec.Label.ToLower().Trim()]     = $uRec }
}

Write-Info ("  Indexation RP -> " + $assetsUsersByAccountId.Count + " accountId | " +
            $assetsUsersByEmail.Count + " email | " + $assetsUsersByName.Count + " nom")

$nbNomPrenom = ($assetsList | Where-Object { -not [string]::IsNullOrWhiteSpace($_.NomPrenom) }).Count
$nbSIMP      = ($assetsList | Where-Object { -not [string]::IsNullOrWhiteSpace($_.CodeDomaineSIMP) }).Count
$nbSociete   = ($assetsList | Where-Object { -not [string]::IsNullOrWhiteSpace($_.Societe) }).Count
Write-Info ("  Qualite RP : " + $nbNomPrenom + " fiches NomPrenom | " +
            $nbSIMP + " fiches Domaine SIMP (attr 378) | " +
            $nbSociete + " fiches Societe (attr 383)")

# ============================================================
# HELPERS DE RÉSOLUTION — (V2 intégraux restaurés)
# ============================================================
function Resolve-Cmdb([string]$AccId, [string]$UserMail, [string]$DispName) {
    if (-not [string]::IsNullOrWhiteSpace($AccId) -and $assetsUsersByAccountId.ContainsKey($AccId)) {
        return $assetsUsersByAccountId[$AccId]
    }
    if (-not [string]::IsNullOrWhiteSpace($UserMail)) {
        $ml = $UserMail.ToLower()
        if ($assetsUsersByEmail.ContainsKey($ml)) { return $assetsUsersByEmail[$ml] }
    }
    if (-not [string]::IsNullOrWhiteSpace($DispName)) {
        $nl = $DispName.ToLower().Trim()
        if ($assetsUsersByName.ContainsKey($nl)) { return $assetsUsersByName[$nl] }
    }
    return $null
}

function Resolve-NomPrenom([string]$CmdbNomPrenom, [string]$CmdbLabel, [string]$JiraDisplayName) {
    if (-not [string]::IsNullOrWhiteSpace($CmdbNomPrenom)) { return $CmdbNomPrenom }
    if (-not [string]::IsNullOrWhiteSpace($CmdbLabel))     { return $CmdbLabel }
    return $JiraDisplayName
}

function Resolve-LibelleSIMP([string]$BudgetKey, [string]$BudgetSummary) {
    # Priorité 1 : BudgetSummary passé en paramètre (vient du cache issue)
    if (-not [string]::IsNullOrWhiteSpace($BudgetSummary) -and $BudgetSummary -ne $BudgetKey) {
        if (-not [string]::IsNullOrWhiteSpace($BudgetKey)) {
            $global:budgetLibelleCache[$BudgetKey] = $BudgetSummary
        }
        return $BudgetSummary
    }
    # Priorité 2 : cache budgetLibelleCache alimenté par la Section 7c
    if (-not [string]::IsNullOrWhiteSpace($BudgetKey) -and $global:budgetLibelleCache.ContainsKey($BudgetKey)) {
        return $global:budgetLibelleCache[$BudgetKey]
    }
    # Fallback : BudgetSummary brut (même si égal à BudgetKey)
    return $BudgetSummary
}

# ============================================================
# 7c. CMDB ASSETS — Référentiel "Domaine SIMP" & Cache Budgets (V2 restauré)
# ============================================================
Write-Info "=== 1c/4 Chargement CMDB Assets — Referentiel Domaine SIMP ==="

$domaineSimpCacheFile = Join-Path $cacheDir "assets_domaine_simp.json"
$domaineSimpList      = $null

$dsAge = Get-CacheAge $domaineSimpCacheFile
if ($dsAge -ge 0) { Write-Info ("  Cache Domaine SIMP : age = " + $dsAge + "h") }

if (-not $cacheChoices.RefreshAssets -and (Test-CacheValide $domaineSimpCacheFile -MaxAgeHours 168)) {
    $domaineSimpList = Load-JsonCache -Path $domaineSimpCacheFile -MaxAgeHours 168
    if ($domaineSimpList) {
        Write-Info ("  -> Domaines SIMP charges depuis le cache (" + $domaineSimpList.Count + " objets — 0 appel API)")
    }
}

if (-not $domaineSimpList) {
    $domaineSimpList = New-Object System.Collections.ArrayList
    $aqlQueries      = @(
        'objectType = "Domaine SIMP"',
        'objectType = "Activite"',
        'objectType = "Budget"',
        'objectType = "Initiative"'
    )
    $qIdx = 0
    foreach ($aql in $aqlQueries) {
        $qIdx++
        Write-Info ("  -> AQL " + $qIdx + "/" + $aqlQueries.Count + " : " + $aql)
        $script:lastAqlAttrDict = @{}
        try {
            $rawObjs = Invoke-AssetsAql `
                -AqlQuery          $aql `
                -Headers           $jiraHeaders `
                -AqlUrl            $assetsAqlUrl `
                -MaxResults        100 `
                -ProgressId        0 `
                -ProgressLabel     ("Assets : " + $aql) `
                -IncludeAttributes

            foreach ($pObj in $rawObjs) {
                if (-not $pObj.id) { continue }
                $objKey  = [string]$pObj.objectKey
                $label   = Fix-Encoding ([string]$pObj.label)
                $codeVal = ""; $nameVal = $label

                if ($pObj.attributes) {
                    foreach ($attr in $pObj.attributes) {
                        $aName = $script:lastAqlAttrDict[[string]$attr.objectTypeAttributeId]
                        if (-not $aName -and $attr.objectTypeAttribute -and $attr.objectTypeAttribute.name) {
                            $aName = Fix-Encoding ([string]$attr.objectTypeAttribute.name)
                        }
                        $vals = $attr.objectAttributeValues
                        if (-not $vals -or $vals.Count -eq 0) { continue }
                        $vVal = Fix-Encoding ([string]$vals[0].value)
                        if ($aName -ieq "Code" -or $aName -ilike "*Code*SIMP*") {
                            if (-not [string]::IsNullOrWhiteSpace($vVal)) { $codeVal = $vVal }
                        } elseif ($aName -ieq "Nom" -or $aName -ieq "Name" -or
                                  $aName -ilike "*Libel*" -or $aName -ilike "*Domaine*") {
                            if (-not [string]::IsNullOrWhiteSpace($vVal)) { $nameVal = $vVal }
                        }
                    }
                }
                [void]$domaineSimpList.Add([pscustomobject]@{
                    ObjectKey = $objKey
                    Code      = $codeVal
                    Label     = $label
                    Name      = $nameVal
                })
            }
            Write-Info ("    -> " + $rawObjs.Count + " objets charges")
        } catch { Write-Warn ("  -> Type inexistant ou erreur AQL : " + $_.Exception.Message) }
    }
    Save-JsonCache -Path $domaineSimpCacheFile -Object $domaineSimpList
    Write-Info ("  -> " + $domaineSimpList.Count + " objets Domaine SIMP sauvegardes dans le cache")
}

# Alimentation du cache global budgetLibelleCache depuis la CMDB Assets
foreach ($item in $domaineSimpList) {
    if (-not [string]::IsNullOrWhiteSpace($item.ObjectKey) -and
        -not [string]::IsNullOrWhiteSpace($item.Label)) {
        $global:budgetLibelleCache[$item.ObjectKey] = $item.Label
    }
    if (-not [string]::IsNullOrWhiteSpace($item.Code) -and
        -not [string]::IsNullOrWhiteSpace($item.Label)) {
        $global:budgetLibelleCache[$item.Code] = $item.Label
    }
    if (-not [string]::IsNullOrWhiteSpace($item.ObjectKey) -and
        -not [string]::IsNullOrWhiteSpace($item.Name)) {
        if (-not $global:budgetLibelleCache.ContainsKey($item.ObjectKey)) {
            $global:budgetLibelleCache[$item.ObjectKey] = $item.Name
        }
    }
}
Write-Info ("  -> Cache Libelle SIMP alimente (" + $global:budgetLibelleCache.Count + " entrees)")

# ============================================================
# 7d. CHARGEMENT WORKLOGS TEMPO
# ============================================================
Write-Info "=== 1d/4 Chargement Worklogs Tempo ==="

$wlCacheFile = Join-Path $cacheDir ("worklogs_" + $moisKey + ".json")
$worklogs    = $null

$wlAge = Get-CacheAge $wlCacheFile
if ($wlAge -ge 0) { Write-Info ("  Cache Worklogs : age = " + $wlAge + "h") }

if (-not $cacheChoices.RefreshWorklogs -and (Test-CacheValide $wlCacheFile -MaxAgeHours 48)) {
    $worklogs = Load-JsonCache -Path $wlCacheFile -MaxAgeHours 48
    if ($worklogs) { Write-Info ("  -> Worklogs charges depuis le cache (" + $worklogs.Count + " worklogs — 0 appel API)") }
}

if (-not $worklogs) {
    $worklogs    = New-Object System.Collections.ArrayList
    $tempoOffset = 0
    $tempoLimit  = 1000

    Write-Info ("  -> Interrogation API Tempo : " + $fromStr + " -> " + $toStr)
    do {
        $tempoUrl = "https://api.eu.tempo.io/4/worklogs?from=" + $fromStr + "&to=" + $toStr +
                    "&limit=" + $tempoLimit + "&offset=" + $tempoOffset
        try {
            $respTempo = Invoke-ApiGet -Url $tempoUrl -Headers $tempoHeaders
        } catch {
            Write-Warn ("Erreur Tempo page offset=" + $tempoOffset + " : " + $_.Exception.Message)
            break
        }
        $wlPage = if ($respTempo.results) { $respTempo.results } else { $respTempo.worklogs }
        if (-not $wlPage -or $wlPage.Count -eq 0) { break }

        foreach ($wl in $wlPage) { [void]$worklogs.Add($wl) }
        # PS 5.1 : barre indéterminée — pas de -PercentComplete sur pagination inconnue
        Write-Progress -Id 1 -Activity "Worklogs Tempo" `
            -Status ($worklogs.Count.ToString() + " worklogs charges (offset=" + $tempoOffset.ToString() + ")...")

        $hasNext = $false
        if     ($respTempo.metadata -and $respTempo.metadata.next) { $hasNext = $true; $tempoOffset += $tempoLimit }
        elseif ($respTempo.next)                                   { $hasNext = $true; $tempoOffset += $tempoLimit }
        else   { $hasNext = $false }
    } while ($hasNext)

    Write-Progress -Id 1 -Activity "Worklogs Tempo" -Completed
    Save-JsonCache -Path $wlCacheFile -Object $worklogs
    Write-Info ("  -> " + $worklogs.Count + " worklogs sauvegardes dans le cache")
}

# ============================================================
# 8. RÉSOLUTION DES UTILISATEURS JIRA (RunspacePool 10 threads)
# ============================================================
Write-Info "=== 2/4 Chargement des profils utilisateurs Jira ==="

$usersCacheFile = Join-Path $cacheDir "jira_users.json"
$jiraUserCache  = @{}

$usersAge = Get-CacheAge $usersCacheFile
if ($usersAge -ge 0) { Write-Info ("  Cache Users : age = " + $usersAge + "h") }

if (-not $cacheChoices.RefreshUsers -and (Test-CacheValide $usersCacheFile -MaxAgeHours 168)) {
    $cachedUsers = Load-JsonCache -Path $usersCacheFile -MaxAgeHours 168
    if ($cachedUsers) {
        foreach ($cu in $cachedUsers) { $jiraUserCache[[string]$cu.AccountId] = $cu }
        Write-Info ("  -> Users charges depuis le cache (" + $jiraUserCache.Count + " profils — 0 appel API)")
    }
}

$unknownAccIds = New-Object System.Collections.ArrayList
foreach ($wl in $worklogs) {
    $aid = [string]$wl.author.accountId
    if (-not [string]::IsNullOrWhiteSpace($aid) -and -not $jiraUserCache.ContainsKey($aid)) {
        if (-not $unknownAccIds.Contains($aid)) { [void]$unknownAccIds.Add($aid) }
    }
}

$resolveUserScript = {
    param($AccountIds, $JiraBaseUrl, $B64, $ProxyUrl, $UseSystemProxy, $ProxyUseDefaultCredentials)

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

    function Fix-Enc($v) {
        if ([string]::IsNullOrWhiteSpace($v)) { return $v }
        if ($v.Contains("Ã")) {
            try {
                $b = [System.Text.Encoding]::GetEncoding(1252).GetBytes($v)
                $d = [System.Text.Encoding]::UTF8.GetString($b)
                if (-not $d.Contains("")) { return $d }
            } catch {}
        }
        return $v
    }

    $headers = @{ Authorization = "Basic " + $B64; Accept = "application/json" }
    $results = @{}

    foreach ($accId in $AccountIds) {
        $url    = $JiraBaseUrl + "/rest/api/3/user?accountId=" + [Uri]::EscapeDataString($accId)
        $params = @{ Method='GET'; Uri=$url; Headers=$headers; ContentType='application/json'; UseBasicParsing=$true; ErrorAction='Stop' }
        $px = Get-ProxyParams $url; foreach ($k in $px.Keys) { $params[$k] = $px[$k] }
        try {
            [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
            $resp   = Invoke-WebRequest @params
            $stream = $resp.RawContentStream; $stream.Position = 0
            $reader = New-Object System.IO.StreamReader($stream, [System.Text.Encoding]::UTF8, $true)
            $raw    = $reader.ReadToEnd(); $reader.Close()
            $u      = $raw | ConvertFrom-Json
            $results[$accId] = [pscustomobject]@{
                AccountId    = $accId
                DisplayName  = Fix-Enc ([string]$u.displayName)
                EmailAddress = Fix-Enc ([string]$u.emailAddress)
            }
        } catch { $results[$accId] = [pscustomobject]@{ AccountId=$accId; DisplayName=""; EmailAddress="" } }
    }
    return $results
}

if ($unknownAccIds.Count -gt 0) {
    Write-Info ("  -> Resolution multi-thread (RunspacePool 10) de " + $unknownAccIds.Count + " profils via API Jira...")

    $threads      = 10
    $runspacePool = [runspacefactory]::CreateRunspacePool(1, $threads)
    $runspacePool.Open()

    $chunkSize = [math]::Ceiling($unknownAccIds.Count / $threads)
    $jobs      = New-Object System.Collections.ArrayList

    for ($i = 0; $i -lt $unknownAccIds.Count; $i += $chunkSize) {
        $count = [math]::Min($chunkSize, $unknownAccIds.Count - $i)
        $chunk = $unknownAccIds.GetRange($i, $count)

        $ps = [powershell]::Create()
        $ps.RunspacePool = $runspacePool
        [void]$ps.AddScript($resolveUserScript)
        [void]$ps.AddArgument($chunk)
        [void]$ps.AddArgument($jiraBaseUrl)
        [void]$ps.AddArgument($b64)
        [void]$ps.AddArgument($ProxyUrl)
        [void]$ps.AddArgument($UseSystemProxy)
        [void]$ps.AddArgument($ProxyUseDefaultCredentials)

        [void]$jobs.Add([pscustomobject]@{ PowerShell=$ps; AsyncResult=$ps.BeginInvoke() })
    }

    foreach ($j in $jobs) {
        $res = $j.PowerShell.EndInvoke($j.AsyncResult)
        if ($res) {
            foreach ($dict in $res) {
                foreach ($k in $dict.Keys) { $jiraUserCache[$k] = $dict[$k] }
            }
        }
        $j.PowerShell.Dispose()
    }
    $runspacePool.Close(); $runspacePool.Dispose()

    $allUsers = New-Object System.Collections.ArrayList
    foreach ($k in $jiraUserCache.Keys) { [void]$allUsers.Add($jiraUserCache[$k]) }
    Save-JsonCache -Path $usersCacheFile -Object $allUsers
    Write-Info ("  -> Cache users mis a jour (" + $jiraUserCache.Count + " profils)")
} else {
    Write-Info ("  -> Tous les profils en cache (" + $jiraUserCache.Count + " profils — 0 appel API)")
}
# ============================================================
# 9a. HIÉRARCHIE JIRA — Setup & Cache
# ============================================================
Write-Info "=== 3/4 Chargement de la hierarchie des tickets Jira ==="

$issuesCacheFile = Join-Path $cacheDir "jira_issues_hierarchy.json"
$jiraIssueCache  = @{}

$issuesAge = Get-CacheAge $issuesCacheFile
if ($issuesAge -ge 0) { Write-Info ("  Cache Hierarchie Jira : age = " + $issuesAge + "h") }

if (-not $cacheChoices.RefreshIssues -and (Test-CacheValide $issuesCacheFile -MaxAgeHours 168)) {
    $cachedIssues = Load-JsonCache -Path $issuesCacheFile -MaxAgeHours 168
    if ($cachedIssues) {
        foreach ($p in $cachedIssues.PSObject.Properties) { $jiraIssueCache[$p.Name] = $p.Value }
        Write-Info ("  -> Hierarchies chargees depuis le cache (" + $jiraIssueCache.Count + " tickets)")
    }
}

$uniqueIssueIds  = @($worklogs | ForEach-Object { $_.issue.id } | Select-Object -Unique)
$issueIdsToFetch = @($uniqueIssueIds | Where-Object { -not $jiraIssueCache.ContainsKey([string]$_) })
Write-Info ("  Tickets : " + $uniqueIssueIds.Count + " uniques | " +
            ($uniqueIssueIds.Count - $issueIdsToFetch.Count) + " cache | " +
            $issueIdsToFetch.Count + " a resoudre")

# ============================================================
# 9b. RÉSOLUTION HIÉRARCHIE — 1 job par ticket, 10 threads (V2 conforme)
# ============================================================
$resolveScript = {
    param($IssId, $JiraBaseUrl, $JiraHeaders, $ProxyUrl, $UseSystemProxy, $ProxyUseDefaultCredentials)

    function Get-ProxyUri([string]$T) {
        if ($ProxyUrl -and $ProxyUrl.Trim()) { return $ProxyUrl }
        if (-not $UseSystemProxy) { return $null }
        try {
            $d  = [uri]$T
            $wp = [System.Net.WebRequest]::DefaultWebProxy
            if ($wp -and -not $wp.IsBypassed($d)) {
                $px = $wp.GetProxy($d)
                if ($px -and $px.AbsoluteUri -ne $d.AbsoluteUri) { return $px.AbsoluteUri }
            }
        } catch {}
        return $null
    }

    function Fix-Enc($v) {
        if ([string]::IsNullOrWhiteSpace($v)) { return $v }
        if ($v.Contains("Ã")) {
            try {
                $b = [System.Text.Encoding]::GetEncoding(1252).GetBytes($v)
                $d = [System.Text.Encoding]::UTF8.GetString($b)
                if (-not $d.Contains("")) { return $d }
            } catch {}
        }
        return $v
    }

    function Get-BudgetKey($cf) {
        if ($null -eq $cf) { return "" }
        $o = if ($cf -is [array] -and $cf.Count -gt 0) { $cf[0] } else { $cf }
        if ($o -is [string])    { return $o.Trim() }
        if ($o.objectKey)       { return [string]$o.objectKey }
        if ($o.key)             { return [string]$o.key }
        if ($o.value)           { return [string]$o.value }
        if ($o.code)            { return [string]$o.code }
        if ($o.id)              { return [string]$o.id }
        if ($o.label)           { return [string]$o.label }
        if ($o.name)            { return [string]$o.name }
        if ($o.displayValue)    { return [string]$o.displayValue }
        return ""
    }

    function Get-BudgetSummary($cf) {
        if ($null -eq $cf) { return "" }
        $o = if ($cf -is [array] -and $cf.Count -gt 0) { $cf[0] } else { $cf }
        if ($o -is [string])    { return "" }
        if ($o.label)           { return [string]$o.label }
        if ($o.summary)         { return [string]$o.summary }
        if ($o.name)            { return [string]$o.name }
        if ($o.displayValue)    { return [string]$o.displayValue }
        if ($o.value)           { return [string]$o.value }
        return ""
    }

    function Get-ObjKey($cf) {
        if ($null -eq $cf) { return "" }
        $o = if ($cf -is [array] -and $cf.Count -gt 0) { $cf[0] } else { $cf }
        if ($o -is [string])    { return $o.Trim() }
        if ($o.objectKey)       { return [string]$o.objectKey }
        if ($o.key)             { return [string]$o.key }
        if ($o.value)           { return [string]$o.value }
        return ""
    }

    function Get-ScalarField($cf) {
        if ($null -eq $cf) { return "" }
        if ($cf -is [string]) { return $cf.Trim() }
        $o = if ($cf -is [array] -and $cf.Count -gt 0) { $cf[0] } else { $cf }
        if ($o -is [string])  { return $o.Trim() }
        if ($o.value)         { return [string]$o.value }
        if ($o.name)          { return [string]$o.name }
        return ""
    }

    function CallJira([string]$Url) {
        $p = @{ Method='GET'; Uri=$Url; Headers=$JiraHeaders;
                ContentType='application/json'; UseBasicParsing=$true; ErrorAction='Stop' }
        $px = Get-ProxyUri $Url
        if ($px) { $p.Proxy = $px; if ($ProxyUseDefaultCredentials) { $p.ProxyUseDefaultCredentials = $true } }
        $r  = Invoke-WebRequest @p
        $s  = $r.RawContentStream; $s.Position = 0
        $rd = New-Object System.IO.StreamReader($s, [System.Text.Encoding]::UTF8, $true)
        $raw = $rd.ReadToEnd(); $rd.Close()
        return ($raw | ConvertFrom-Json)
    }

    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

    $currId       = [string]$IssId
    $visited      = @{}
    $depth        = 0
    $issueKey     = ""; $issueSummary  = ""
    $initKey      = ""; $initSummary   = ""
    $budgetKey    = ""; $budgetSumm    = ""
    $codeFDR      = ""; $codeCigrefInit= ""

    while ($currId -and -not $visited.ContainsKey($currId) -and $depth -lt 6) {
        $visited[$currId] = $true
        $depth++
        $iResp = $null
        try {
            $iUrl  = $JiraBaseUrl + "/rest/api/3/issue/" + $currId +
                     "?fields=key,summary,parent,issuetype," +
                     "customfield_10183,customfield_10124,customfield_10014,customfield_10310"
            $iResp = CallJira $iUrl
        } catch { break }
        if (-not $iResp) { break }

        $key   = [string]$iResp.key
        $summ  = Fix-Enc ([string]$iResp.fields.summary)
        $iType = if ($iResp.fields.issuetype -and $iResp.fields.issuetype.name) {
                     [string]$iResp.fields.issuetype.name } else { "" }
        $bKey  = Fix-Enc (Get-BudgetKey     $iResp.fields.customfield_10183)
        $bSumm = Fix-Enc (Get-BudgetSummary $iResp.fields.customfield_10183)
        $cfFDR = Get-ObjKey    $iResp.fields.customfield_10124
        $cfCIG = Get-ScalarField $iResp.fields.customfield_10310

        # Premier niveau = issue de départ
        if (-not $issueKey) { $issueKey = $key; $issueSummary = $summ }

        # Code FDR / Gepetto
        if (-not $codeFDR -and $cfFDR) { $codeFDR = $cfFDR }

        # Code CIGREF porté par l'Initiative (ou tout niveau qui le porte)
        if (-not [string]::IsNullOrWhiteSpace($cfCIG) -and
            [string]::IsNullOrWhiteSpace($codeCigrefInit)) {
            $codeCigrefInit = $cfCIG
        }

        # -------------------------------------------------------
        # Capture Initiative — V2 STRICTE
        # Condition : iType -ieq "Initiative" OU iType -ilike "*Initiative*"
        # Prise uniquement si pas encore capturée ($initKey vide)
        # -------------------------------------------------------
        if (($iType -ieq "Initiative" -or $iType -ilike "*Initiative*") -and
            [string]::IsNullOrWhiteSpace($initKey)) {
            $initKey     = $key
            $initSummary = $summ
            if (-not [string]::IsNullOrWhiteSpace($cfCIG)) { $codeCigrefInit = $cfCIG }
        }

        # -------------------------------------------------------
        # Capture Budget — V2 STRICTE
        # Cas 1 : ticket courant EST un Budget (type ou clé BUD-*)
        #         → budgetSumm = summary du ticket BUD (= Libellé SIMP)
        #         → budgetKey  = CF10183 si présent, sinon clé du ticket
        # Cas 2 : ticket courant porte un CF10183
        #         → prise temporaire si budgetKey pas encore rempli
        # -------------------------------------------------------
        if ($iType -ieq "Budget" -or $key -ilike "BUD-*" -or $key -ilike "*BUD*") {
            $budgetSumm = $summ
            if (-not [string]::IsNullOrWhiteSpace($bKey)) {
                $budgetKey = $bKey
            } else {
                if ([string]::IsNullOrWhiteSpace($budgetKey)) { $budgetKey = $key }
            }
        } else {
            if ([string]::IsNullOrWhiteSpace($budgetKey) -and
                -not [string]::IsNullOrWhiteSpace($bKey)) {
                $budgetKey = $bKey
                $budgetSumm = $bSumm
            }
        }

        # Remontée vers le parent
        $nextParent = $null
        if      ($iResp.fields.parent -and $iResp.fields.parent.id)  { $nextParent = [string]$iResp.fields.parent.id }
        elseif  ($iResp.fields.parent -and $iResp.fields.parent.key) { $nextParent = [string]$iResp.fields.parent.key }
        elseif  ($iResp.fields.customfield_10014)                    { $nextParent = Get-ObjKey $iResp.fields.customfield_10014 }
        $currId = $nextParent
    }

    # Interrogation secondaire du ticket BUD-* pour CF10183 + summary (V2)
    if (-not [string]::IsNullOrWhiteSpace($budgetKey) -and $budgetKey -ilike "BUD-*") {
        try {
            $budUrl  = $JiraBaseUrl + "/rest/api/3/issue/" + $budgetKey + "?fields=summary,customfield_10183"
            $budResp = CallJira $budUrl
            if ($budResp -and $budResp.fields) {
                if ($budResp.fields.summary) {
                    $budgetSumm = Fix-Enc ([string]$budResp.fields.summary)
                }
                $budCf = Fix-Enc (Get-BudgetKey $budResp.fields.customfield_10183)
                if (-not [string]::IsNullOrWhiteSpace($budCf)) { $budgetKey = $budCf }
            }
        } catch {}
    }

    # Fallback InitiativeKey V2 : IssueKey si rien trouvé, "AUTRE" si IssueKey vide
    if ([string]::IsNullOrWhiteSpace($initKey)) {
        $initKey     = if ($issueKey) { $issueKey } else { "AUTRE" }
        $initSummary = $issueSummary
    }

    return [pscustomobject]@{
        Id                = [string]$IssId
        IssueKey          = $issueKey
        IssueSummary      = $issueSummary
        InitiativeKey     = $initKey
        InitiativeSummary = $initSummary
        BudgetKey         = $budgetKey
        BudgetSummary     = $budgetSumm
        CodeFDR           = $codeFDR
        CodeCigrefInit    = $codeCigrefInit
    }
}

# ============================================================
# 9c. LANCEMENT — 1 job par ticket, pool 10 threads, boucle active V2
# ============================================================
if ($issueIdsToFetch.Count -gt 0) {
    Write-Info ("  -> Resolution parallele (10 threads) de " +
                $issueIdsToFetch.Count + " tickets via API Jira...")

    $pool = [RunspaceFactory]::CreateRunspacePool(1, 10)
    $pool.ApartmentState = "MTA"
    $pool.Open()

    $jobs = New-Object System.Collections.ArrayList
    foreach ($issId in $issueIdsToFetch) {
        $ps = [PowerShell]::Create()
        $ps.RunspacePool = $pool
        [void]$ps.AddScript($resolveScript)
        [void]$ps.AddArgument([string]$issId)
        [void]$ps.AddArgument($jiraBaseUrl)
        [void]$ps.AddArgument($jiraHeaders)
        [void]$ps.AddArgument($ProxyUrl)
        [void]$ps.AddArgument($UseSystemProxy)
        [void]$ps.AddArgument($ProxyUseDefaultCredentials)
        [void]$jobs.Add([pscustomobject]@{ IssId=[string]$issId; PS=$ps; Handle=$ps.BeginInvoke() })
    }

    $completed  = 0
    $totalFetch = $issueIdsToFetch.Count
    $nextSave   = 200

    # Boucle active V2 — collecte au fil de la complétion
    while ($jobs.Count -gt 0) {
        $done = @($jobs | Where-Object { $_.Handle.IsCompleted })
        foreach ($job in $done) {
            $completed++
            try {
                $r = $job.PS.EndInvoke($job.Handle)
                if ($r -and $r.Id) {
                    $jiraIssueCache[$r.Id] = [pscustomobject]@{
                        IssueKey          = $r.IssueKey
                        IssueSummary      = $r.IssueSummary
                        InitiativeKey     = $r.InitiativeKey
                        InitiativeSummary = $r.InitiativeSummary
                        BudgetKey         = $r.BudgetKey
                        BudgetSummary     = $r.BudgetSummary
                        CodeFDR           = $r.CodeFDR
                        CodeCigrefInit    = $r.CodeCigrefInit
                    }
                }
            } catch {
                Write-Log ("Erreur ticket " + $job.IssId + " : " + $_.Exception.Message) "WARN"
            }
            $job.PS.Dispose()
            [void]$jobs.Remove($job)

            Write-Progress -Id 2 -Activity "Hierarchie Jira (10 threads)" `
                -Status ("Termines : " + $completed + "/" + $totalFetch) `
                -PercentComplete ([math]::Round($completed / $totalFetch * 100))

            # Sauvegarde intermédiaire tous les 200 tickets (V2)
            if ($completed -ge $nextSave) {
                Save-JsonCache -Path $issuesCacheFile -Object $jiraIssueCache
                $nextSave += 200
            }
        }
        if ($jobs.Count -gt 0) { Start-Sleep -Milliseconds 100 }
    }

    $pool.Close()
    $pool.Dispose()
    Write-Progress -Id 2 -Activity "Hierarchie Jira (10 threads)" -Completed

    # Sauvegarde finale
    Save-JsonCache -Path $issuesCacheFile -Object $jiraIssueCache
    Write-Info ("  -> " + $completed + " tickets resolus en parallele")

} else {
    Write-Info ("  -> Toutes les issues en cache (" + $jiraIssueCache.Count + " tickets — 0 appel API)")
}

Write-Info ("  Hierarchies Jira total : " + $jiraIssueCache.Count + " tickets")

Write-Info ("=== Donnees sources chargees : " +
    $worklogs.Count       + " worklogs | " +
    $jiraUserCache.Count  + " users | " +
    $jiraIssueCache.Count + " issues | " +
    $assetsList.Count     + " fiches RP | " +
    $global:budgetLibelleCache.Count + " libelles SIMP ===")

# ============================================================
# 10a. FORMAT DÉTERMINÉ : Consolidation worklogs (V2)
# ============================================================
if ($Format -eq "DETERMINE" -or $Format -eq "BOTH") {
    Write-Info "=== 4a/4 Generation DETERMINE (CSV / XLSX / JSON) ==="

    $determineRows = New-Object System.Collections.ArrayList
    $xlsxGrouped   = @{}
    $jsonDetGroups = @{}
    $totalWL       = $worklogs.Count
    $idxWL         = 0

    foreach ($wl in $worklogs) {
        $idxWL++
        if ($idxWL % 200 -eq 0 -or $idxWL -eq $totalWL) {
            Write-Progress -Id 4 -Activity "Generation DETERMINE" `
                -Status ("Worklog " + $idxWL + "/" + $totalWL + " — " + $xlsxGrouped.Count + " lignes agregees") `
                -PercentComplete ([math]::Round($idxWL / $totalWL * 100))
        }

        $accId = [string]$wl.author.accountId
        $issId = [string]$wl.issue.id
        $jUser = $jiraUserCache[$accId]
        $dName = ""; if ($jUser -and $jUser.DisplayName)  { $dName = [string]$jUser.DisplayName }
        $uMail = ""; if ($jUser -and $jUser.EmailAddress) { $uMail = [string]$jUser.EmailAddress }

        $cmdb   = Resolve-Cmdb -AccId $accId -UserMail $uMail -DispName $dName
        $jIssue = $jiraIssueCache[[string]$issId]

        $cmdbNP           = if ($cmdb -and $cmdb.NomPrenom) { $cmdb.NomPrenom } else { "" }
        $cmdbLbl          = if ($cmdb -and $cmdb.Label)     { $cmdb.Label }     else { "" }
        $nomPrenomAffiche = Resolve-NomPrenom -CmdbNomPrenom $cmdbNP -CmdbLabel $cmdbLbl -JiraDisplayName $dName

        $nomBrut    = if ($cmdb -and $cmdb.NOM)    { $cmdb.NOM }    else { "" }
        $prenomBrut = if ($cmdb -and $cmdb.Prenom) { $cmdb.Prenom } else { "" }

        # InitiativeKey : valeur issue du Split 3 (Initiative > Epic > IssueKey)
        $initKey = "AUTRE"
        if ($jIssue -and -not [string]::IsNullOrWhiteSpace($jIssue.InitiativeKey)) {
            $initKey = [string]$jIssue.InitiativeKey
        }

        $heures = [math]::Round([double]$wl.timeSpentSeconds / 3600.0, 5)
        $jours  = [math]::Round($heures / 8.0, 5)

        $typeRes = "Prestataire"
        if ($cmdb -and $cmdb.TypeRessource) { $typeRes = $cmdb.TypeRessource }

        $codeSIMP = ""
        if ($jIssue -and $jIssue.BudgetKey) { $codeSIMP = [string]$jIssue.BudgetKey }

        $bKey        = if ($jIssue -and $jIssue.BudgetKey)     { $jIssue.BudgetKey }     else { "" }
        $bSum        = if ($jIssue -and $jIssue.BudgetSummary) { $jIssue.BudgetSummary } else { "" }
        $libelleSIMP = Resolve-LibelleSIMP -BudgetKey $bKey -BudgetSummary $bSum

        $codeCigrefInit = ""
        if ($jIssue -and $jIssue.CodeCigrefInit) { $codeCigrefInit = [string]$jIssue.CodeCigrefInit }

        $mat=""; $soc="Harmonie Mutuelle"; $dir="(Non trouve)"; $srv=""; $cds=""; $cfd=""
        if ($cmdb) {
            if ($cmdb.Matricule)       { $mat = $cmdb.Matricule }
            if ($cmdb.Direction)       { $dir = $cmdb.Direction }
            if ($cmdb.Service)         { $srv = $cmdb.Service }
            if ($cmdb.Societe)         { $soc = $cmdb.Societe }
            if ($cmdb.CodeDomaineSIMP) { $cds = $cmdb.CodeDomaineSIMP }
        }
        if ($jIssue -and $jIssue.CodeFDR) { $cfd = [string]$jIssue.CodeFDR }

        $initLib = ""
        if ($jIssue -and $jIssue.InitiativeSummary) { $initLib = [string]$jIssue.InitiativeSummary }

        # ---- Ligne CSV DETERMINE (détail brut par worklog) ----
        [void]$determineRows.Add([pscustomobject]@{
            idCle             = ($initKey + "_" + $typeRes + "_" + $moisKey + "_" + $accId)
            exercice          = $From.Year
            mois              = $From.Month.ToString("D2")
            dateSaisie        = ConvertTo-DateFR $wl.startDate
            accountIdJira     = $accId
            nomAffichage      = $nomPrenomAffiche
            matricule         = $mat
            nom               = $nomBrut
            prenom            = $prenomBrut
            typeRessource     = $typeRes
            societe           = $soc
            direction         = $dir
            codeDomaineSIMP   = $cds
            codeBudget        = $codeSIMP
            libelleBudget     = $libelleSIMP
            cleInitiative     = $initKey
            libelleInitiative = $initLib
            codeFDR           = $cfd
            codeCigrefInit    = $codeCigrefInit
            heuresTotales     = $heures
            joursTotaux       = $jours
        })

        # ---- Groupement XLSX SUIVI-ACTIVITE (V2 : Collaborateur x Initiative) ----
        $groupKey = ($accId + "|" + $initKey)
        if (-not $xlsxGrouped.ContainsKey($groupKey)) {
            $initLbl = $initKey
            if (-not [string]::IsNullOrWhiteSpace($initLib) -and $initLib -ne $initKey) {
                $initLbl = $initKey + " - " + $initLib
            }
            $xlsxGrouped[$groupKey] = [ordered]@{
                "Annee"                     = $From.Year
                "Mois"                      = $From.Month
                "NOM Prenom"                = $nomPrenomAffiche
                "Type ressource"            = $typeRes
                "Societe"                   = $soc
                "Direction"                 = $dir
                "Service"                   = $srv
                "Matricule"                 = $mat
                "Code domaine Ressource"    = $cds
                "Code SIMP"                 = $codeSIMP
                "Libelle SIMP"              = $libelleSIMP
                "Code - Libelle initiative" = $initLbl
                "Nb jours"                  = 0.0
                "Code CIGREF"               = $codeCigrefInit
                "Code Gepetto"              = $cfd
                "Identifiant Jira"          = $accId
            }
        } else {
            # Enrichissement progressif V2 si valeurs manquantes sur la 1re entrée
            if ([string]::IsNullOrWhiteSpace($xlsxGrouped[$groupKey]["Code SIMP"]) -and
                -not [string]::IsNullOrWhiteSpace($codeSIMP)) {
                $xlsxGrouped[$groupKey]["Code SIMP"] = $codeSIMP
            }
            if ([string]::IsNullOrWhiteSpace($xlsxGrouped[$groupKey]["Libelle SIMP"]) -and
                -not [string]::IsNullOrWhiteSpace($libelleSIMP)) {
                $xlsxGrouped[$groupKey]["Libelle SIMP"] = $libelleSIMP
            }
            if ([string]::IsNullOrWhiteSpace($xlsxGrouped[$groupKey]["Code CIGREF"]) -and
                -not [string]::IsNullOrWhiteSpace($codeCigrefInit)) {
                $xlsxGrouped[$groupKey]["Code CIGREF"] = $codeCigrefInit
            }
            if ([string]::IsNullOrWhiteSpace($xlsxGrouped[$groupKey]["Code Gepetto"]) -and
                -not [string]::IsNullOrWhiteSpace($cfd)) {
                $xlsxGrouped[$groupKey]["Code Gepetto"] = $cfd
            }
            if ($xlsxGrouped[$groupKey]["Code - Libelle initiative"] -eq $initKey -and
                -not [string]::IsNullOrWhiteSpace($initLib)) {
                $xlsxGrouped[$groupKey]["Code - Libelle initiative"] = $initKey + " - " + $initLib
            }
        }
        $xlsxGrouped[$groupKey]["Nb jours"] += $jours

        # ---- Groupement JSON Assets EXCOMP DETERMINE ----
        $jsonDetKey = ($accId + "|" + $initKey + "|" + $moisKey)
        if (-not $jsonDetGroups.ContainsKey($jsonDetKey)) {
            $jsonDetGroups[$jsonDetKey] = [ordered]@{
                "objectType"        = "DETERMINE"
                "exercice"          = $From.Year
                "anneeFiscale"      = $fyKey
                "moisCode"          = $mKey
                "moisLabel"         = $From.ToString("MM/yyyy")
                "accountId"         = $accId
                "nomPrenom"         = $nomPrenomAffiche
                "nom"               = $nomBrut
                "prenom"            = $prenomBrut
                "matricule"         = $mat
                "typeRessource"     = $typeRes
                "societe"           = $soc
                "direction"         = $dir
                "service"           = $srv
                "codeDomaine"       = $cds
                "cleInitiative"     = $initKey
                "libelleInitiative" = $initLib
                "codeBudget"        = $codeSIMP
                "libelleBudget"     = $libelleSIMP
                "codeCIGREF"        = $codeCigrefInit
                "codeGepetto"       = $cfd
                "nbJours"           = 0.0
                "nbHeures"          = 0.0
            }
        } else {
            if ([string]::IsNullOrWhiteSpace($jsonDetGroups[$jsonDetKey]["codeBudget"]) -and
                -not [string]::IsNullOrWhiteSpace($codeSIMP)) {
                $jsonDetGroups[$jsonDetKey]["codeBudget"]    = $codeSIMP
                $jsonDetGroups[$jsonDetKey]["libelleBudget"] = $libelleSIMP
            }
            if ([string]::IsNullOrWhiteSpace($jsonDetGroups[$jsonDetKey]["codeCIGREF"]) -and
                -not [string]::IsNullOrWhiteSpace($codeCigrefInit)) {
                $jsonDetGroups[$jsonDetKey]["codeCIGREF"] = $codeCigrefInit
            }
            if ([string]::IsNullOrWhiteSpace($jsonDetGroups[$jsonDetKey]["codeGepetto"]) -and
                -not [string]::IsNullOrWhiteSpace($cfd)) {
                $jsonDetGroups[$jsonDetKey]["codeGepetto"] = $cfd
            }
        }
        $jsonDetGroups[$jsonDetKey]["nbJours"]  += $jours
        $jsonDetGroups[$jsonDetKey]["nbHeures"] += $heures
    }
    Write-Progress -Id 4 -Activity "Generation DETERMINE" -Completed

    $determineRows = @($determineRows | Sort-Object nomAffichage)

    # ============================================================
    # 10b. EXPORT CSV DETERMINE
    # ============================================================
    if ($ExportFormats -contains "CSV") {
        $csvDetermine = Join-Path $exportsDir ("EtatsFinanciers_DETERMINE_" + $moisKey + "_" + $runStamp + ".csv")
        Export-CsvStrict -Path $csvDetermine `
            -Headers @(
                "idCle","exercice","mois","dateSaisie","accountIdJira","nomAffichage",
                "matricule","nom","prenom","typeRessource","societe","direction",
                "codeDomaineSIMP","codeBudget","libelleBudget","cleInitiative",
                "libelleInitiative","codeFDR","codeCigrefInit","heuresTotales","joursTotaux"
            ) `
            -Rows $determineRows
    }

    # ============================================================
    # 10c. EXPORT XLSX DETERMINE (SUIVI-ACTIVITE) — V2 Strict
    # ============================================================
    if ($ExportFormats -contains "XLSX") {
        $xlsxSuiviRows = New-Object System.Collections.ArrayList
        foreach ($gKey in $xlsxGrouped.Keys) {
            $r = $xlsxGrouped[$gKey]
            $r["Nb jours"] = [math]::Round([double]$r["Nb jours"], 2)
            [void]$xlsxSuiviRows.Add([pscustomobject]$r)
        }
        $xlsxSuiviRows = @($xlsxSuiviRows | Sort-Object "NOM Prenom", "Code - Libelle initiative")

        $xlsxSuiviFile = Join-Path $exportsDir ("EXPORT_JIRADOT_SUIVI-ACTIVITE_" + $runStamp + ".xlsx")
        Export-GenericXlsx -Path $xlsxSuiviFile -SheetName "Suivi Activite" `
            -Headers @(
                "Annee","Mois","NOM Prenom","Type ressource","Societe","Direction","Service",
                "Matricule","Code domaine Ressource","Code SIMP","Libelle SIMP",
                "Code - Libelle initiative","Nb jours","Code CIGREF","Code Gepetto","Identifiant Jira"
            ) `
            -Rows $xlsxSuiviRows
        Write-Info ("  SUIVI-ACTIVITE XLSX : " + $xlsxSuiviRows.Count +
                    " lignes agregees (Collaborateur x Initiative) — triees par NOM Prenom / Initiative")
    }

    # ============================================================
    # 10d. EXPORT JSON ASSETS EXCOMP — DETERMINE
    # ============================================================
    if ($ExportFormats -contains "JSON") {
        $jsonDetArray = New-Object System.Collections.ArrayList
        foreach ($jKey in $jsonDetGroups.Keys) {
            $obj = $jsonDetGroups[$jKey]
            $obj["nbJours"]  = [math]::Round([double]$obj["nbJours"],  5)
            $obj["nbHeures"] = [math]::Round([double]$obj["nbHeures"], 5)
            [void]$jsonDetArray.Add([pscustomobject]$obj)
        }
        $jsonDetArray = @($jsonDetArray | Sort-Object nomPrenom, cleInitiative)
        $jsonDetFile  = Join-Path $exportsDir ("EXCOMP_ASSETS_DETERMINE_" + $moisKey + "_" + $runStamp + ".json")
        Export-JsonAssetsExcomp -Path $jsonDetFile -DataArray $jsonDetArray
        Write-Info ("  JSON EXCOMP DETERMINE : " + $jsonDetArray.Count + " objets Assets")
    }
}

# ============================================================
# 11a. FORMAT SIMP : Consolidation worklogs
# ============================================================
if ($Format -eq "SIMP" -or $Format -eq "BOTH") {
    Write-Info "=== 4b/4 Generation SIMP (CSV / XLSX / JSON) ==="

    $simpGrouped = @{}
    $totalWL     = $worklogs.Count
    $idxWL       = 0

    foreach ($wl in $worklogs) {
        $idxWL++
        if ($idxWL % 200 -eq 0 -or $idxWL -eq $totalWL) {
            Write-Progress -Id 5 -Activity "Generation SIMP" `
                -Status ("Worklog " + $idxWL + "/" + $totalWL + " — " + $simpGrouped.Count + " lignes SIMP") `
                -PercentComplete ([math]::Round($idxWL / $totalWL * 100))
        }

        $accId = [string]$wl.author.accountId
        $issId = [string]$wl.issue.id
        $jUser = $jiraUserCache[$accId]
        $dName = ""; if ($jUser -and $jUser.DisplayName)  { $dName = [string]$jUser.DisplayName }
        $uMail = ""; if ($jUser -and $jUser.EmailAddress) { $uMail = [string]$jUser.EmailAddress }

        $cmdb   = Resolve-Cmdb -AccId $accId -UserMail $uMail -DispName $dName
        $jIssue = $jiraIssueCache[[string]$issId]

        $heures = [math]::Round([double]$wl.timeSpentSeconds / 3600.0, 5)
        $jours  = [math]::Round($heures / 8.0, 5)

        $typeRes = "Prestataire"
        if ($cmdb -and $cmdb.TypeRessource) { $typeRes = $cmdb.TypeRessource }

        $codeDomaine = ""
        if ($cmdb -and $cmdb.CodeDomaineSIMP) { $codeDomaine = $cmdb.CodeDomaineSIMP }

        # Type SIMP : Régie (Prestataire/Externe) / Interne (Employé HM)
        $typeSIMP = if ($typeRes -ieq "Interne" -or
                        $typeRes -ieq "Employe" -or
                        $typeRes -ieq "Employé") { "Interne" } else { "Régie" }

        # Code activité SIMP = Code Budget (BUD-*)
        $codeSIMP = ""
        if ($jIssue -and $jIssue.BudgetKey) { $codeSIMP = [string]$jIssue.BudgetKey }
        if ([string]::IsNullOrWhiteSpace($codeSIMP)) { $codeSIMP = "SANS-SIMP" }

        $bKey        = if ($jIssue -and $jIssue.BudgetKey)     { $jIssue.BudgetKey }     else { "" }
        $bSum        = if ($jIssue -and $jIssue.BudgetSummary) { $jIssue.BudgetSummary } else { "" }
        $libelleSIMP = Resolve-LibelleSIMP -BudgetKey $bKey -BudgetSummary $bSum

        # Clé de groupement SIMP : typeSIMP + codeDomaine + codeSIMP
        $simpKey = ($typeSIMP + "|" + $codeDomaine + "|" + $codeSIMP)

        if (-not $simpGrouped.ContainsKey($simpKey)) {
            $simpGrouped[$simpKey] = [ordered]@{
                "Exercice"          = $From.Year
                "Mois"              = $From.Month.ToString("D2")
                "Type Ressource"    = $typeSIMP
                "Code Domaine SIMP" = $codeDomaine
                "Code SIMP"         = $codeSIMP
                "Libelle SIMP"      = $libelleSIMP
                "Nb Jours"          = 0.0
                "Nb Heures"         = 0.0
                "_anneeFiscale"     = $fyKey
                "_moisCode"         = $mKey
                "_moisLabel"        = $From.ToString("MM/yyyy")
            }
        } else {
            if ([string]::IsNullOrWhiteSpace($simpGrouped[$simpKey]["Libelle SIMP"]) -and
                -not [string]::IsNullOrWhiteSpace($libelleSIMP)) {
                $simpGrouped[$simpKey]["Libelle SIMP"] = $libelleSIMP
            }
        }
        $simpGrouped[$simpKey]["Nb Jours"]  += $jours
        $simpGrouped[$simpKey]["Nb Heures"] += $heures
    }
    Write-Progress -Id 5 -Activity "Generation SIMP" -Completed

    foreach ($k in $simpGrouped.Keys) {
        $simpGrouped[$k]["Nb Jours"]  = [math]::Round([double]$simpGrouped[$k]["Nb Jours"],  2)
        $simpGrouped[$k]["Nb Heures"] = [math]::Round([double]$simpGrouped[$k]["Nb Heures"], 2)
    }

    $simpRows = New-Object System.Collections.ArrayList
    foreach ($k in $simpGrouped.Keys) { [void]$simpRows.Add([pscustomobject]$simpGrouped[$k]) }
    $simpRows = @($simpRows | Sort-Object "Code SIMP", "Type Ressource")

    Write-Info ("  SIMP : " + $simpRows.Count + " lignes agregees (Type x Domaine x Code SIMP)")

    # ============================================================
    # 11b. EXPORT CSV SIMP
    # ============================================================
    if ($ExportFormats -contains "CSV") {
        $csvSimp = Join-Path $exportsDir ("EtatsFinanciers_SIMP_" + $moisKey + "_" + $runStamp + ".csv")
        Export-CsvStrict -Path $csvSimp `
            -Headers @(
                "Exercice","Mois","Type Ressource","Code Domaine SIMP",
                "Code SIMP","Libelle SIMP","Nb Jours","Nb Heures"
            ) `
            -Rows $simpRows
    }

    # ============================================================
    # 11c. EXPORT XLSX SIMP — V2 Strict (Restauration En-têtes V2)
    # En-têtes V2 : Contribution, Activite informatique, Domaine applicatif,
    #               Indicateur, Annee, Periode, Donnees
    # ============================================================
    if ($ExportFormats -contains "XLSX") {
        $xlsxSimpRows = New-Object System.Collections.ArrayList

        foreach ($k in $simpGrouped.Keys) {
            $r        = $simpGrouped[$k]
            $codeS    = [string]$r["Code SIMP"]
            $libS     = [string]$r["Libelle SIMP"]
            $typeS    = [string]$r["Type Ressource"]
            $domaineS = [string]$r["Code Domaine SIMP"]
            $nbJ      = [math]::Round([double]$r["Nb Jours"], 2)

            $actInfo = $codeS
            if (-not [string]::IsNullOrWhiteSpace($libS) -and $libS -ne $codeS) {
                $actInfo = $codeS + " - " + $libS
            }

            [void]$xlsxSimpRows.Add([ordered]@{
                "Contribution"          = $typeS
                "Activite informatique" = $actInfo
                "Domaine applicatif"    = $domaineS
                "Indicateur"            = "Jours"
                "Annee"                 = $From.Year
                "Periode"               = ("M" + $From.Month.ToString("D2"))
                "Donnees"               = $nbJ
            })
        }
        $xlsxSimpRows = @($xlsxSimpRows | Sort-Object "Activite informatique", "Contribution")

        $xlsxSimpFile = Join-Path $exportsDir ("EXPORT_JIRADOT_SIMP_" + $runStamp + ".xlsx")
        Export-GenericXlsx -Path $xlsxSimpFile -SheetName "SIMP" `
            -Headers @(
                "Contribution","Activite informatique","Domaine applicatif",
                "Indicateur","Annee","Periode","Donnees"
            ) `
            -Rows $xlsxSimpRows
        Write-Info ("  XLSX SIMP genere (Format V2 Strict) : " + $xlsxSimpFile)
    }

    # ============================================================
    # 11d. EXPORT JSON ASSETS EXCOMP — SIMP
    # ============================================================
    if ($ExportFormats -contains "JSON") {
        $jsonSimpArray = New-Object System.Collections.ArrayList

        foreach ($k in $simpGrouped.Keys) {
            $r             = $simpGrouped[$k]
            $codeS         = [string]$r["Code SIMP"]
            $typeS         = [string]$r["Type Ressource"]
            $domaine       = [string]$r["Code Domaine SIMP"]
            $lib           = [string]$r["Libelle SIMP"]
            $nbJ           = [math]::Round([double]$r["Nb Jours"],  2)
            $nbH           = [math]::Round([double]$r["Nb Heures"], 2)
            $fy            = [string]$r["_anneeFiscale"]
            $mk            = [string]$r["_moisCode"]
            $ml            = [string]$r["_moisLabel"]
            $exo           = [int]$r["Exercice"]

            $typeSanitized = $typeS -replace "é","e" -replace "è","e" -replace "ê","e" -replace " ","-"

            [void]$jsonSimpArray.Add([ordered]@{
                "objectType"  = "SIMP"
                "id"          = ($codeS + "_" + $typeSanitized + "_" + $mk + "_" + $fy)
                "code"        = $codeS
                "type"        = $typeS
                "totalTime"   = $nbJ
                "totalHours"  = $nbH
                "activity"    = $lib
                "domain"      = $domaine
                "year"        = $fy
                "month"       = $mk
                "moisLabel"   = $ml
                "exercice"    = $exo
            })
        }

        $jsonSimpArray = @($jsonSimpArray | Sort-Object { $_["code"] }, { $_["type"] })
        $jsonSimpFile  = Join-Path $exportsDir ("EXCOMP_ASSETS_SIMP_" + $moisKey + "_" + $runStamp + ".json")
        Export-JsonAssetsExcomp -Path $jsonSimpFile -DataArray $jsonSimpArray
        Write-Info ("  JSON EXCOMP SIMP : " + $jsonSimpArray.Count + " objets Assets")
    }
}

# ============================================================
# 12. RÉCAPITULATIF FINAL
# ============================================================
Write-Log ("=== FIN EXECUTION " + $scriptName + " v3.0 — " + (Get-Date -Format "yyyy-MM-dd HH:mm:ss") + " ===") "INFO"

Write-Host ""
Write-Host ("=" * 65) -ForegroundColor DarkCyan
Write-Host ("  EXECUTION TERMINEE — " + $scriptName + " v3.0") -ForegroundColor White
Write-Host ("=" * 65) -ForegroundColor DarkCyan
Write-Host ("  Periode    : " + $From.ToString("dd/MM/yyyy") + " -> " + $To.ToString("dd/MM/yyyy")) -ForegroundColor Gray
Write-Host ("  Perimetre  : " + $Format)                                                              -ForegroundColor Gray
Write-Host ("  Formats    : " + ($ExportFormats -join " | "))                                         -ForegroundColor Gray
Write-Host ("  Worklogs   : " + $worklogs.Count)                                                      -ForegroundColor Gray
Write-Host ("  Users      : " + $jiraUserCache.Count)                                                 -ForegroundColor Gray
Write-Host ("  Issues     : " + $jiraIssueCache.Count)                                               -ForegroundColor Gray
Write-Host ("  Fiches RP  : " + $assetsList.Count)                                                   -ForegroundColor Gray
Write-Host ("  Libelles   : " + $global:budgetLibelleCache.Count + " entrees cache SIMP")            -ForegroundColor Gray
Write-Host ""
Write-Host "  Fichiers generes dans : " -NoNewline -ForegroundColor DarkCyan
Write-Host $exportsDir                  -ForegroundColor Yellow
Write-Host ""

Get-ChildItem -Path $exportsDir -Filter ("*" + $runStamp + "*") |
    Sort-Object Name |
    ForEach-Object {
        $ext  = $_.Extension.ToUpper().TrimStart(".")
        $col  = switch ($ext) {
            "XLSX" { "Green"  }
            "CSV"  { "Cyan"   }
            "JSON" { "Yellow" }
            default{ "Gray"   }
        }
        $size = [math]::Round($_.Length / 1KB, 1)
        Write-Host ("  [" + $ext.PadRight(4) + "] " + $_.Name + "  (" + $size + " Ko)") -ForegroundColor $col
    }

Write-Host ""
Write-Host ("=" * 65) -ForegroundColor DarkCyan