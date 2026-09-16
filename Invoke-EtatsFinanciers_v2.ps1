<#
.SYNOPSIS
    Extraction des états financiers Jiradot depuis Tempo + Jira + Assets CMDB
.VERSION
    2.5 — 2026-09-10
.NOTES
    Auteur  : DSIM / HM_DSIM_PACT
    Sorties : CSV (DETERMINE, SIMP) + XLSX (SUIVI-ACTIVITE, SIMP)
#>
[CmdletBinding()]

param(
    [DateTime]$From,
    [DateTime]$To,
    [ValidateSet("DETERMINE","SIMP","BOTH")]
    [string]$Format,
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
$scriptName  = "Invoke-EtatsFinanciers_v2"
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
            -Status ("Page " + $pageNum + " - " + $results.Count + " objets...") -PercentComplete -1

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
# 4. INTERFACE UTILISATEUR (Dialogs WPF)
# ============================================================
function Show-EtatsFinanciersDialog {
    Add-Type -AssemblyName PresentationFramework

    $now = Get-Date

    $xaml = @"
<Window
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    Title="Extraction Etats Financiers v2.5" Width="420" Height="260"
    WindowStartupLocation="CenterScreen" ResizeMode="NoResize">
  <StackPanel Margin="20">
    <TextBlock Text="Mois d extraction :" FontWeight="Bold" Margin="0,0,0,5"/>
    <ComboBox x:Name="cbMois" Width="280" Height="28" HorizontalAlignment="Left" Margin="0,0,0,15"/>
    <TextBlock Text="Format d export :" FontWeight="Bold" Margin="0,0,0,5"/>
    <StackPanel Orientation="Horizontal" Margin="0,0,0,15">
      <RadioButton x:Name="rbDetermine" Content="DETERMINE"  Margin="0,0,15,0" IsChecked="True"/>
      <RadioButton x:Name="rbSimp"      Content="SIMP"        Margin="0,0,15,0"/>
      <RadioButton x:Name="rbBoth"      Content="LES DEUX"/>
    </StackPanel>
    <Button x:Name="btnOk" Content="Lancer l extraction" Width="200" Height="35"
            Margin="0,5,0,0" HorizontalAlignment="Center"/>
  </StackPanel>
</Window>
"@

    $reader = New-Object System.Xml.XmlNodeReader([xml]$xaml)
    $win    = [Windows.Markup.XamlReader]::Load($reader)

    $cbMois     = $win.FindName("cbMois")
    $monthsList = New-Object System.Collections.ArrayList

    # Génération des 24 derniers mois — M-1 présélectionné par défaut
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

    $result = @{ From=$null; To=$null; Format="BOTH" }

    $win.FindName("btnOk").Add_Click({
        $idx = $cbMois.SelectedIndex
        if ($idx -lt 0) { $idx = 1 }
        $sel           = $monthsList[$idx]
        $result.From   = $sel.FirstDay
        $result.To     = $sel.LastDay
        $result.Format = if ($win.FindName("rbDetermine").IsChecked) { "DETERMINE" }
                         elseif ($win.FindName("rbSimp").IsChecked)  { "SIMP" }
                         else { "BOTH" }
        $win.DialogResult = $true
        $win.Close()
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
# 5. FONCTIONS D'EXPORTATION (CSV + CONVERSION CSV -> XLSX)
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
                } else {
                    $s = [string]$raw
                }
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
    if (-not (Test-Path -LiteralPath $CsvPath)) {
        Write-Warn ("CSV source introuvable : " + $CsvPath); return
    }
    $fullCsvPath  = (Get-Item -LiteralPath $CsvPath).FullName
    $fullXlsxPath = [System.IO.Path]::GetFullPath($XlsxPath)

    if (Get-Module -ListAvailable -Name Import-Excel) {
        try {
            Import-Module Import-Excel -ErrorAction Stop
            Import-Csv -Path $fullCsvPath -Delimiter ";" |
                Export-Excel -Path $fullXlsxPath -WorksheetName $SheetName -AutoSize -TableStyle Medium6 -Show:$false
            Write-Info ("  -> XLSX converti (Import-Excel) : " + $XlsxPath)
            return
        } catch { Write-Warn ("Import-Excel erreur : " + $_.Exception.Message) }
    }

    try {
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible       = $false
        $excel.DisplayAlerts = $false
        $excel.Workbooks.OpenText(
            $fullCsvPath,
            65001, 1, 1, 1, $false, $false, $true, $false, $false, $false
        )
        $wb = $excel.ActiveWorkbook
        $ws = $wb.Worksheets.Item(1)
        try { $ws.Name = $SheetName } catch {}
        $usedCols    = $ws.UsedRange.Columns.Count
        $headerRange = $ws.Range($ws.Cells.Item(1,1), $ws.Cells.Item(1,$usedCols))
        $headerRange.Font.Bold           = $true
        $headerRange.Font.ColorIndex     = 2
        $headerRange.Interior.ColorIndex = 23
        $ws.Columns.AutoFit() | Out-Null
        if (Test-Path -LiteralPath $fullXlsxPath) {
            Remove-Item -LiteralPath $fullXlsxPath -Force -ErrorAction SilentlyContinue
        }
        $wb.SaveAs($fullXlsxPath, 51)
        $wb.Close($false)
        $excel.Quit()
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
        Write-Info ("  -> XLSX converti (Excel COM OpenText) : " + $XlsxPath)
        return
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
        if (Test-Path -LiteralPath $tempCsv) {
            Remove-Item -LiteralPath $tempCsv -Force -ErrorAction SilentlyContinue
        }
    }
}
# ============================================================
# 6. INITIALISATION PÉRIODE, CREDENTIALS ET CACHE CHOICES
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

Write-Info ("Periode : " + $From.ToString("dd/MM/yyyy") + " -> " + $To.ToString("dd/MM/yyyy"))
Write-Info ("Format  : " + $Format)

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
$tempoToken   = [Runtime.InteropServices.Marshal]::PtrToStringAuto([Runtime.InteropServices.Marshal]::SecureStringToBSTR($tempoData.Token))
$tempoHeaders = @{ Authorization = "Bearer " + $tempoToken; Accept = "application/json" }

$assetsAqlUrl    = $jiraBaseUrl + "/gateway/api/jsm/assets/workspace/" + $AssetsWorkspaceId + "/v1/object/aql"
$assetsSchemaUrl = $jiraBaseUrl + "/gateway/api/jsm/assets/workspace/" + $AssetsWorkspaceId + "/v1/objectschema/list"

Write-Log ("=== DEBUT EXECUTION " + $scriptName + " v2.5 — " + $runStamp + " ===") "INFO"

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
                "375" { $attr375  = $vBest }   # Attribut 375 = NOM
                "376" { $attr376  = $vBest }   # Attribut 376 = Prénom
                "407" { if ([string]::IsNullOrWhiteSpace($nomPrenom)) { $nomPrenom = $vBest } }  # Fallback NOM Prénom
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
                    $aName = $script:lastAqlAttrDict[$attrId]
                    if (-not $aName -and $attr.objectTypeAttribute -and $attr.objectTypeAttribute.name) {
                        $aName = Fix-Encoding ([string]$attr.objectTypeAttribute.name)
                    }
                    if (-not $aName) { continue }

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

        # Reconstruction NOM Prénom : priorité aux attributs 375 (NOM) + 376 (Prénom)
        if (-not [string]::IsNullOrWhiteSpace($attr375) -or -not [string]::IsNullOrWhiteSpace($attr376)) {
            $nomPrenom = ($attr375.Trim() + " " + $attr376.Trim()).Trim()
        }
        # Fallbacks si ni 375/376 ni 407 ne sont renseignés
        if ([string]::IsNullOrWhiteSpace($nomPrenom)) { $nomPrenom = $labelAssets }

        # Règle Métier Société :
        # Priorité absolue à l'attribut 383. Fallback uniquement si vide.
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
            CodeDomaineSIMP = $codeSimp     # Attribut 378
            NomPrenom       = $nomPrenom    # Attributs 375 (NOM) + 376 (Prénom)
            Matricule       = $matricule
            Direction       = $directionStr
            Service         = $serviceStr
            Societe         = $societe      # Attribut 383
            CodeCIGREF      = $codeCigref
        })
    }
    Save-JsonCache -Path $assetsCacheFile -Object $assetsList
    Write-Info ("  -> " + $assetsList.Count + " fiches RP sauvegardees dans le cache")
}


# ============================================================
# 7b. CMDB ASSETS — RP (Indexation et contrôle qualité)
# ============================================================
Write-Info "=== 1b/4 Indexation et contrôle qualité RP ==="

foreach ($uRec in $assetsList) {
    if (-not [string]::IsNullOrWhiteSpace($uRec.AccountId))      { $assetsUsersByAccountId[$uRec.AccountId]         = $uRec }
    if (-not [string]::IsNullOrWhiteSpace($uRec.ExtractedEmail)) { $assetsUsersByEmail[$uRec.ExtractedEmail]        = $uRec }
    if (-not [string]::IsNullOrWhiteSpace($uRec.Label))          { $assetsUsersByName[$uRec.Label.Trim().ToLower()] = $uRec }
}

Write-Info ("  Indexation RP -> " + $assetsUsersByAccountId.Count + " accountId | " +
            $assetsUsersByEmail.Count + " email | " + $assetsUsersByName.Count + " nom")

$nbNomPrenom = ($assetsList | Where-Object { -not [string]::IsNullOrWhiteSpace($_.NomPrenom) }).Count
$nbSIMP      = ($assetsList | Where-Object { -not [string]::IsNullOrWhiteSpace($_.CodeDomaineSIMP) }).Count
$nbSociete   = ($assetsList | Where-Object { -not [string]::IsNullOrWhiteSpace($_.Societe) }).Count

Write-Info ("  Qualité RP : " + $nbNomPrenom + " fiches avec NomPrenom (attr 407) | " +
            $nbSIMP + " fiches avec Domaine SIMP (attr 378) | " +
            $nbSociete + " fiches avec Société (attr 383)")

# ============================================================
# 7c. CMDB ASSETS — Référentiel "Domaine SIMP"
# ============================================================
Write-Info "=== 1c/4 Chargement CMDB Assets — Referentiel Domaine SIMP ==="

$domaineSimpCacheFile = Join-Path $cacheDir "assets_domaine_simp.json"
$domaineSimpList      = $null

$dsAge = Get-CacheAge $domaineSimpCacheFile
if ($dsAge -ge 0) { Write-Info ("  Cache Domaine SIMP : age = " + $dsAge + "h") }

if (-not $cacheChoices.RefreshAssets -and (Test-CacheValide $domaineSimpCacheFile -MaxAgeHours 168)) {
    $domaineSimpList = Load-JsonCache -Path $domaineSimpCacheFile -MaxAgeHours 168
    if ($domaineSimpList) { Write-Info ("  -> Domaines SIMP charges depuis le cache (" + $domaineSimpList.Count + " objets — 0 appel API)") }
}

if (-not $domaineSimpList) {
    $domaineSimpList = New-Object System.Collections.ArrayList

    $aqlQueries = @(
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
                    ObjectKey=$objKey; Code=$codeVal; Label=$label; Name=$nameVal
                })
            }
            Write-Info ("    -> " + $rawObjs.Count + " objets charges")
        } catch { Write-Warn ("    -> Type inexistant ou erreur : " + $_.Exception.Message) }
        Start-Sleep -Milliseconds 300
    }
    Save-JsonCache -Path $domaineSimpCacheFile -Object $domaineSimpList
    Write-Info ("  -> " + $domaineSimpList.Count + " objets budgetaires sauvegardes dans le cache")
}

$mapCount = 0
foreach ($item in $domaineSimpList) {
    $libelle = if (-not [string]::IsNullOrWhiteSpace($item.Name)) { $item.Name } else { $item.Label }
    if (-not [string]::IsNullOrWhiteSpace($item.Code)) {
        $global:budgetLibelleCache[$item.Code] = $libelle; $mapCount++
    }
    if (-not [string]::IsNullOrWhiteSpace($item.ObjectKey)) {
        $global:budgetLibelleCache[$item.ObjectKey] = $libelle; $mapCount++
    }
    if (-not [string]::IsNullOrWhiteSpace($item.Label)) {
        $global:budgetLibelleCache[$item.Label] = $libelle; $mapCount++
        if ($item.Label -match "^([^\s-]+)\s*-\s*(.+)$") {
            $global:budgetLibelleCache[$matches[1].Trim()] = $matches[2].Trim(); $mapCount++
        }
    }
}

$rpSimpCount = 0
foreach ($uRec in $assetsList) {
    if (-not [string]::IsNullOrWhiteSpace($uRec.CodeDomaineSIMP)) {
        $code = $uRec.CodeDomaineSIMP.Trim()
        if (-not $global:budgetLibelleCache.ContainsKey($code)) {
            $global:budgetLibelleCache[$code] = $code
            $rpSimpCount++
        }
    }
}
Write-Info ("  Dictionnaire SIMP : " + $mapCount + " paires domaines | " + $rpSimpCount + " codes RP supplementaires | Delai API : " + $global:assetsApiDelayMs + "ms")

# ============================================================
# 8. EXTRACTION WORKLOGS TEMPO
# ============================================================
Write-Info "=== 2/4 Recuperation des Worklogs Tempo ==="

$tempoCacheFile = Join-Path $cacheDir ("worklogs_tempo_" + $moisKey + ".json")
$worklogs       = $null

if (-not $cacheChoices.RefreshWorklogs) {
    $worklogs = Load-JsonCache -Path $tempoCacheFile -MaxAgeHours 720
    if ($worklogs) { Write-Info ("  -> Worklogs charges depuis le cache (" + $worklogs.Count + " enregistrements)") }
}

if (-not $worklogs) {
    $worklogs  = New-Object System.Collections.ArrayList
    $tempoUrl  = "https://api.eu.tempo.io/4/worklogs?from=" + $fromStr + "&to=" + $toStr
    $tempoPage = 0
    while ($true) {
        $tempoPage++
        Write-Progress -Id 0 -Activity "Worklogs Tempo" `
            -Status ("Page " + $tempoPage + " — " + $worklogs.Count + " worklogs...") -PercentComplete -1
        $resp = Invoke-ApiGet -Url $tempoUrl -Headers $tempoHeaders
        if ($resp.results) { foreach ($wl in $resp.results) { [void]$worklogs.Add($wl) } }
        if ($resp.metadata -and $resp.metadata.next) { $tempoUrl = $resp.metadata.next } else { break }
    }
    Write-Progress -Id 0 -Activity "Worklogs Tempo" -Completed
    Save-JsonCache -Path $tempoCacheFile -Object $worklogs
    Write-Info ("  -> " + $worklogs.Count + " worklogs sauvegardes dans le cache")
}

# ============================================================
# 9a. PROFILS UTILISATEURS JIRA — PARALLÈLE (10 threads)
# ============================================================
Write-Info "=== 3/4 Enrichissement Profils & Hierarchie Jira (Parallele) ==="

$usersCacheFile = Join-Path $cacheDir "jira_users_cache.json"
$jiraUserCache  = @{}

if (-not $cacheChoices.RefreshUsers) {
    $cachedUsers = Load-JsonCache -Path $usersCacheFile -MaxAgeHours 168
    if ($cachedUsers) {
        foreach ($p in $cachedUsers.PSObject.Properties) { $jiraUserCache[$p.Name] = $p.Value }
        Write-Info ("  -> Profils Jira charges depuis le cache (" + $jiraUserCache.Count + " utilisateurs)")
    }
}

$uniqueAccIds  = @($worklogs | ForEach-Object { $_.author.accountId } | Select-Object -Unique)
$accIdsToFetch = @($uniqueAccIds | Where-Object { -not [string]::IsNullOrWhiteSpace($_) -and -not $jiraUserCache.ContainsKey($_) })
Write-Info ("  Profils : " + $uniqueAccIds.Count + " uniques | " + ($uniqueAccIds.Count - $accIdsToFetch.Count) + " cache | " + $accIdsToFetch.Count + " a resoudre")

if ($accIdsToFetch.Count -gt 0) {
    $resolveUserScript = {
        param($AccId, $JiraBaseUrl, $JiraHeaders, $ProxyUrl, $UseSystemProxy, $ProxyUseDefaultCredentials)
        function Get-ProxyUri { param($T)
            if ($ProxyUrl -and $ProxyUrl.Trim()) { return $ProxyUrl }
            if (-not $UseSystemProxy) { return $null }
            try { $d=[uri]$T; $wp=[System.Net.WebRequest]::DefaultWebProxy
                if ($wp -and -not $wp.IsBypassed($d)) { $px=$wp.GetProxy($d)
                    if ($px -and $px.AbsoluteUri -ne $d.AbsoluteUri) { return $px.AbsoluteUri } } } catch {}
            return $null }
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
        [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
        $dn = "Utilisateur Jira (" + $AccId + ")"; $em = ""
        try {
            $u  = $JiraBaseUrl + "/rest/api/3/user?accountId=" + $AccId
            $p  = @{ Method='GET'; Uri=$u; Headers=$JiraHeaders; ContentType='application/json'; UseBasicParsing=$true; ErrorAction='Stop' }
            $px = Get-ProxyUri $u
            if ($px) { $p.Proxy=$px; if ($ProxyUseDefaultCredentials) { $p.ProxyUseDefaultCredentials=$true } }
            $r  = Invoke-WebRequest @p; $s=$r.RawContentStream; $s.Position=0
            $rd = New-Object System.IO.StreamReader($s,[System.Text.Encoding]::UTF8,$true)
            $j  = $rd.ReadToEnd()|ConvertFrom-Json; $rd.Close()
            if ($j.displayName)  { $dn = Fix-Enc([string]$j.displayName) }
            if ($j.emailAddress) { $em = [string]$j.emailAddress.Trim().ToLower() }
        } catch {}
        return [pscustomobject]@{ AccountId=$AccId; DisplayName=$dn; EmailAddress=$em }
    }

    $poolU    = [RunspaceFactory]::CreateRunspacePool(1, 10)
    $poolU.ApartmentState = "MTA"; $poolU.Open()
    $userJobs = New-Object System.Collections.ArrayList
    foreach ($accId in $accIdsToFetch) {
        $ps = [PowerShell]::Create(); $ps.RunspacePool = $poolU
        [void]$ps.AddScript($resolveUserScript)
        [void]$ps.AddArgument([string]$accId); [void]$ps.AddArgument($jiraBaseUrl)
        [void]$ps.AddArgument($jiraHeaders);   [void]$ps.AddArgument($ProxyUrl)
        [void]$ps.AddArgument($UseSystemProxy); [void]$ps.AddArgument($ProxyUseDefaultCredentials)
        [void]$userJobs.Add([pscustomobject]@{ AccId=[string]$accId; PS=$ps; Handle=$ps.BeginInvoke() })
    }

    $completedU=0; $totalU=$accIdsToFetch.Count
    while ($userJobs.Count -gt 0) {
        $doneU = @($userJobs | Where-Object { $_.Handle.IsCompleted })
        foreach ($job in $doneU) {
            $completedU++
            try {
                $r = $job.PS.EndInvoke($job.Handle)
                if ($r -and $r.AccountId) {
                    $jiraUserCache[$r.AccountId] = [pscustomobject]@{ DisplayName=$r.DisplayName; EmailAddress=$r.EmailAddress }
                }
            } catch {
                $jiraUserCache[$job.AccId] = [pscustomobject]@{ DisplayName=("Utilisateur Jira (" + $job.AccId + ")"); EmailAddress="" }
            }
            $job.PS.Dispose(); [void]$userJobs.Remove($job)
            Write-Progress -Id 1 -Activity "Profils Jira (10 threads)" `
                -Status ("Termines : " + $completedU + "/" + $totalU) `
                -PercentComplete ([math]::Round($completedU/$totalU*100))
        }
        if ($userJobs.Count -gt 0) { Start-Sleep -Milliseconds 100 }
    }
    $poolU.Close(); $poolU.Dispose()
    Write-Progress -Id 1 -Activity "Profils Jira (10 threads)" -Completed
    Save-JsonCache -Path $usersCacheFile -Object $jiraUserCache
    Write-Info ("  -> " + $completedU + " profils resolus")
}
Write-Info ("  Profils Jira total : " + $jiraUserCache.Count)

# ============================================================
# 9b. HIÉRARCHIE TICKETS JIRA — PARALLÈLE (10 threads)
# ============================================================
$issuesCacheFile = Join-Path $cacheDir "jira_issues_hierarchy.json"
$jiraIssueCache  = @{}

if (-not $cacheChoices.RefreshIssues) {
    $cachedIssues = Load-JsonCache -Path $issuesCacheFile -MaxAgeHours 168
    if ($cachedIssues) {
        foreach ($p in $cachedIssues.PSObject.Properties) { $jiraIssueCache[$p.Name] = $p.Value }
        Write-Info ("  -> Hierarchies chargees depuis le cache (" + $jiraIssueCache.Count + " tickets)")
    }
}

$uniqueIssueIds  = @($worklogs | ForEach-Object { $_.issue.id } | Select-Object -Unique)
$issueIdsToFetch = @($uniqueIssueIds | Where-Object { -not $jiraIssueCache.ContainsKey([string]$_) })
Write-Info ("  Tickets : " + $uniqueIssueIds.Count + " uniques | " + ($uniqueIssueIds.Count - $issueIdsToFetch.Count) + " cache | " + $issueIdsToFetch.Count + " a resoudre")

if ($issueIdsToFetch.Count -gt 0) {
    $resolveScript = {
        param($IssId, $JiraBaseUrl, $JiraHeaders, $ProxyUrl, $UseSystemProxy, $ProxyUseDefaultCredentials)
        function Get-ProxyUri { param($T)
            if ($ProxyUrl -and $ProxyUrl.Trim()) { return $ProxyUrl }
            if (-not $UseSystemProxy) { return $null }
            try { $d=[uri]$T; $wp=[System.Net.WebRequest]::DefaultWebProxy
                if ($wp -and -not $wp.IsBypassed($d)) { $px=$wp.GetProxy($d)
                    if ($px -and $px.AbsoluteUri -ne $d.AbsoluteUri) { return $px.AbsoluteUri } } } catch {}
            return $null }
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
            if ($o -is [string]) { return $o.Trim() }
            if ($o.objectKey)    { return [string]$o.objectKey }
            if ($o.key)          { return [string]$o.key }
            if ($o.value)        { return [string]$o.value }
            if ($o.code)         { return [string]$o.code }
            if ($o.id)           { return [string]$o.id }
            if ($o.label)        { return [string]$o.label }
            if ($o.name)         { return [string]$o.name }
            if ($o.displayValue) { return [string]$o.displayValue }
            return "" }
        function Get-BudgetSummary($cf) {
            if ($null -eq $cf) { return "" }
            $o = if ($cf -is [array] -and $cf.Count -gt 0) { $cf[0] } else { $cf }
            if ($o -is [string]) { return "" }
            if ($o.label)        { return [string]$o.label }
            if ($o.summary)      { return [string]$o.summary }
            if ($o.name)         { return [string]$o.name }
            if ($o.displayValue) { return [string]$o.displayValue }
            if ($o.value)        { return [string]$o.value }
            return "" }
        function Get-ObjKey($cf) {
            if ($null -eq $cf) { return "" }
            $o = if ($cf -is [array] -and $cf.Count -gt 0) { $cf[0] } else { $cf }
            if ($o -is [string]) { return $o.Trim() }
            if ($o.objectKey) { return [string]$o.objectKey }
            if ($o.key)       { return [string]$o.key }
            if ($o.value)     { return [string]$o.value }
            return "" }
        function Get-ScalarField($cf) {
            if ($null -eq $cf) { return "" }
            if ($cf -is [string]) { return $cf.Trim() }
            $o = if ($cf -is [array] -and $cf.Count -gt 0) { $cf[0] } else { $cf }
            if ($o -is [string]) { return $o.Trim() }
            if ($o.value)        { return [string]$o.value }
            if ($o.name)         { return [string]$o.name }
            return "" }
        function CallJira($Url) {
            $p  = @{ Method='GET'; Uri=$Url; Headers=$JiraHeaders; ContentType='application/json'; UseBasicParsing=$true; ErrorAction='Stop' }
            $px = Get-ProxyUri $Url
            if ($px) { $p.Proxy=$px; if ($ProxyUseDefaultCredentials) { $p.ProxyUseDefaultCredentials=$true } }
            $r  = Invoke-WebRequest @p; $s=$r.RawContentStream; $s.Position=0
            $rd = New-Object System.IO.StreamReader($s,[System.Text.Encoding]::UTF8,$true)
            $raw = $rd.ReadToEnd(); $rd.Close(); return ($raw | ConvertFrom-Json) }

        [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
        $currId=[string]$IssId; $visited=@{}; $depth=0
        $issueKey=""; $issueSummary=""; $initKey=""; $initSummary=""
        $budgetKey=""; $budgetSumm=""; $codeFDR=""; $codeCigrefInit=""

        while ($currId -and -not $visited.ContainsKey($currId) -and $depth -lt 6) {
            $visited[$currId]=$true; $depth++
            $iResp = $null
            try {
                # customfield_10310 = Code CIGREF Initiative
                # customfield_10183 = Lien / Code SIMP
                $iUrl  = $JiraBaseUrl + "/rest/api/3/issue/" + $currId +
                         "?fields=key,summary,parent,issuetype," +
                         "customfield_10183,customfield_10124,customfield_10014,customfield_10310"
                $iResp = CallJira $iUrl
            } catch { break }
            if (-not $iResp) { break }

            $key    = [string]$iResp.key
            $summ   = Fix-Enc([string]$iResp.fields.summary)
            $iType  = if ($iResp.fields.issuetype -and $iResp.fields.issuetype.name) { [string]$iResp.fields.issuetype.name } else { "" }
            $bKey   = Fix-Enc(Get-BudgetKey     $iResp.fields.customfield_10183)
            $bSumm  = Fix-Enc(Get-BudgetSummary $iResp.fields.customfield_10183)
            $cfFDR  = Get-ObjKey $iResp.fields.customfield_10124
            $cfCIG  = Get-ScalarField $iResp.fields.customfield_10310

            if (-not $issueKey) { $issueKey=$key; $issueSummary=$summ }
            if (-not $codeFDR -and $cfFDR) { $codeFDR=$cfFDR }

            # Code CIGREF porté par l'Initiative
            if (-not [string]::IsNullOrWhiteSpace($cfCIG) -and [string]::IsNullOrWhiteSpace($codeCigrefInit)) {
                $codeCigrefInit = $cfCIG
            }

            if (($iType -ieq "Initiative" -or $iType -ilike "*Initiative*") -and -not $initKey) {
                $initKey=$key; $initSummary=$summ
                if (-not [string]::IsNullOrWhiteSpace($cfCIG)) { $codeCigrefInit = $cfCIG }
            }

            # Si le ticket courant est le ticket Budget Parent (issuetype = Budget ou clé BUD-*)
            if (($iType -ieq "Budget" -or $key -ilike "BUD-*" -or $key -ilike "*BUD*")) {
                $budgetSumm = $summ  # Libellé SIMP = résumé du ticket BUD
                if (-not [string]::IsNullOrWhiteSpace($bKey)) {
                    $budgetKey = $bKey  # Code SIMP = customfield_10183 du ticket BUD
                } else {
                    if ([string]::IsNullOrWhiteSpace($budgetKey)) { $budgetKey = $key }
                }
            } else {
                # Niveaux sous-jacents : récupération temporaire si présent
                if (-not $budgetKey -and $bKey) { $budgetKey=$bKey; $budgetSumm=$bSumm }
            }

            $nextParent = $null
            if      ($iResp.fields.parent -and $iResp.fields.parent.id)  { $nextParent=[string]$iResp.fields.parent.id }
            elseif  ($iResp.fields.parent -and $iResp.fields.parent.key) { $nextParent=[string]$iResp.fields.parent.key }
            elseif  ($iResp.fields.customfield_10014)                    { $nextParent=Get-ObjKey $iResp.fields.customfield_10014 }
            $currId = $nextParent
        }

        # Si $budgetKey correspond à une clé de ticket BUD (ex: BUD-113),
        # interroger directement le ticket BUD pour extraire son summary (Libellé SIMP) ET son customfield_10183 (Code SIMP)
        if (-not [string]::IsNullOrWhiteSpace($budgetKey) -and $budgetKey -ilike "BUD-*") {
            try {
                $budUrl  = $JiraBaseUrl + "/rest/api/3/issue/" + $budgetKey + "?fields=summary,customfield_10183"
                $budResp = CallJira $budUrl
                if ($budResp -and $budResp.fields) {
                    if ($budResp.fields.summary) {
                        $budgetSumm = Fix-Enc([string]$budResp.fields.summary)
                    }
                    $budCf10183 = Fix-Enc(Get-BudgetKey $budResp.fields.customfield_10183)
                    if (-not [string]::IsNullOrWhiteSpace($budCf10183)) {
                        $budgetKey = $budCf10183
                    }
                }
            } catch {}
        }

        if (-not $initKey)     { $initKey     = if ($issueKey) { $issueKey } else { "AUTRE" } }
        if (-not $initSummary) { $initSummary = $issueSummary }

        return [pscustomobject]@{
            Id               = [string]$IssId
            IssueKey         = $issueKey
            IssueSummary     = $issueSummary
            InitiativeKey    = $initKey
            InitiativeSummary= $initSummary
            BudgetKey        = $budgetKey        # Code SIMP (cf[10183] du budget parent)
            BudgetSummary    = $budgetSumm     # Libellé SIMP (summary du budget parent)
            CodeFDR          = $codeFDR
            CodeCigrefInit   = $codeCigrefInit # customfield_10310
        }
    }

    $pool = [RunspaceFactory]::CreateRunspacePool(1, 10)
    $pool.ApartmentState = "MTA"; $pool.Open()
    $jobs = New-Object System.Collections.ArrayList
    foreach ($issId in $issueIdsToFetch) {
        $ps = [PowerShell]::Create(); $ps.RunspacePool = $pool
        [void]$ps.AddScript($resolveScript)
        [void]$ps.AddArgument([string]$issId); [void]$ps.AddArgument($jiraBaseUrl)
        [void]$ps.AddArgument($jiraHeaders);   [void]$ps.AddArgument($ProxyUrl)
        [void]$ps.AddArgument($UseSystemProxy); [void]$ps.AddArgument($ProxyUseDefaultCredentials)
        [void]$jobs.Add([pscustomobject]@{ IssId=[string]$issId; PS=$ps; Handle=$ps.BeginInvoke() })
    }

    $completed=0; $totalFetch=$issueIdsToFetch.Count; $nextSave=200
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
            } catch { Write-Log ("Erreur ticket " + $job.IssId + " : " + $_.Exception.Message) "WARN" }
            $job.PS.Dispose(); [void]$jobs.Remove($job)
            Write-Progress -Id 2 -Activity "Hierarchie Jira (10 threads)" `
                -Status ("Termines : " + $completed + "/" + $totalFetch) `
                -PercentComplete ([math]::Round($completed/$totalFetch*100))
            if ($completed -ge $nextSave) {
                Save-JsonCache -Path $issuesCacheFile -Object $jiraIssueCache
                $nextSave += 200
            }
        }
        if ($jobs.Count -gt 0) { Start-Sleep -Milliseconds 100 }
    }
    $pool.Close(); $pool.Dispose()
    Write-Progress -Id 2 -Activity "Hierarchie Jira (10 threads)" -Completed
    Save-JsonCache -Path $issuesCacheFile -Object $jiraIssueCache
    Write-Info ("  -> " + $completed + " tickets resolus en parallele")
}
Write-Info ("  Hierarchies Jira total : " + $jiraIssueCache.Count + " tickets")

# ============================================================
# Fonctions de résolution
# ============================================================
function Resolve-Cmdb {
    param([string]$AccId, [string]$UserMail, [string]$DispName)
    if (-not [string]::IsNullOrWhiteSpace($AccId) -and $assetsUsersByAccountId.ContainsKey($AccId)) {
        return $assetsUsersByAccountId[$AccId]
    }
    if (-not [string]::IsNullOrWhiteSpace($UserMail) -and $assetsUsersByEmail.ContainsKey($UserMail)) {
        return $assetsUsersByEmail[$UserMail]
    }
    if (-not [string]::IsNullOrWhiteSpace($DispName) -and $assetsUsersByName.ContainsKey($DispName.Trim().ToLower())) {
        return $assetsUsersByName[$DispName.Trim().ToLower()]
    }
    return $null
}

function Resolve-NomPrenom {
    param([string]$CmdbNomPrenom, [string]$CmdbLabel, [string]$JiraDisplayName)
    if (-not [string]::IsNullOrWhiteSpace($CmdbNomPrenom)) { return $CmdbNomPrenom.Trim() }
    if (-not [string]::IsNullOrWhiteSpace($CmdbLabel))     { return $CmdbLabel.Trim() }
    if (-not [string]::IsNullOrWhiteSpace($JiraDisplayName)) { return $JiraDisplayName.Trim() }
    return "(Inconnu)"
}

function Resolve-LibelleSIMP {
    param([string]$BudgetKey, [string]$BudgetSummary)

    if (-not [string]::IsNullOrWhiteSpace($BudgetSummary)) {
        return $BudgetSummary.Trim()
    }
    if (-not [string]::IsNullOrWhiteSpace($BudgetKey)) {
        if ($global:budgetLibelleCache.ContainsKey($BudgetKey)) {
            return $global:budgetLibelleCache[$BudgetKey]
        }
        return $BudgetKey
    }
    return ""
}

# ============================================================
# 10. EXPORTATION DES FICHIERS CSV ET XLSX
# ============================================================
Write-Info "=== 4/4 Generation des fichiers CSV & XLSX ==="

# ============================================================
# 10a. FORMAT DÉTERMINÉ : CSV + XLSX SUIVI-ACTIVITÉ
# ============================================================
if ($Format -eq "DETERMINE" -or $Format -eq "BOTH") {

    $determineRows = New-Object System.Collections.ArrayList
    $xlsxGrouped   = @{}
    $totalWL       = $worklogs.Count
    $idxWL         = 0

    foreach ($wl in $worklogs) {
        $idxWL++
        if ($idxWL % 1000 -eq 0 -or $idxWL -eq $totalWL) {
            Write-Progress -Id 4 -Activity "Generation DETERMINE" `
                -Status ("Worklog " + $idxWL + "/" + $totalWL + " — " + $xlsxGrouped.Count + " lignes agregees") `
                -PercentComplete ([math]::Round($idxWL / $totalWL * 100))
        }

        $accId  = [string]$wl.author.accountId
        $issId  = [string]$wl.issue.id
        $jUser  = $jiraUserCache[$accId]
        $dName  = ""; if ($jUser -and $jUser.DisplayName)  { $dName = [string]$jUser.DisplayName }
        $uMail  = ""; if ($jUser -and $jUser.EmailAddress) { $uMail = [string]$jUser.EmailAddress }

        $cmdb   = Resolve-Cmdb -AccId $accId -UserMail $uMail -DispName $dName
        $jIssue = $jiraIssueCache[[string]$issId]

        $cmdbNP           = if ($cmdb -and $cmdb.NomPrenom) { $cmdb.NomPrenom } else { "" }
        $cmdbLbl          = if ($cmdb -and $cmdb.Label)     { $cmdb.Label }     else { "" }
        $nomPrenomAffiche = Resolve-NomPrenom -CmdbNomPrenom $cmdbNP -CmdbLabel $cmdbLbl -JiraDisplayName $dName

        $initKey = "AUTRE"
        if ($jIssue -and $jIssue.InitiativeKey) { $initKey = [string]$jIssue.InitiativeKey }

        $heures  = [math]::Round([double]$wl.timeSpentSeconds / 3600.0, 5)
        $jours   = [math]::Round($heures / 8.0, 5)

        $typeRes = "Prestataire"
        if ($cmdb -and $cmdb.TypeRessource) { $typeRes = $cmdb.TypeRessource }

        $codeSIMP = ""; $budgetSummaryRaw = ""
        if ($jIssue -and $jIssue.BudgetKey)     { $codeSIMP         = [string]$jIssue.BudgetKey }
        if ($jIssue -and $jIssue.BudgetSummary) { $budgetSummaryRaw = [string]$jIssue.BudgetSummary }

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

        [void]$determineRows.Add([pscustomobject]@{
            idCle             = ($initKey + "_" + $typeRes + "_" + $moisKey + "_" + $accId)
            exercice          = $From.Year
            mois              = $From.Month.ToString("D2")
            dateSaisie        = ConvertTo-DateFR $wl.startDate
            accountIdJira     = $accId
            nomAffichage      = $nomPrenomAffiche
            matricule         = $mat
            nom               = $nomPrenomAffiche
            prenom            = ""
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

        $groupKey = ($accId + "|" + $initKey)
        if (-not $xlsxGrouped.ContainsKey($groupKey)) {
            $initLbl = $initKey
            if ($jIssue -and $jIssue.InitiativeSummary) {
                $initLbl = $initKey + " - " + [string]$jIssue.InitiativeSummary
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
            if ([string]::IsNullOrWhiteSpace($xlsxGrouped[$groupKey]["Libelle SIMP"]) -and
                -not [string]::IsNullOrWhiteSpace($libelleSIMP)) {
                $xlsxGrouped[$groupKey]["Libelle SIMP"] = $libelleSIMP
            }
            if ([string]::IsNullOrWhiteSpace($xlsxGrouped[$groupKey]["Code CIGREF"]) -and
                -not [string]::IsNullOrWhiteSpace($codeCigrefInit)) {
                $xlsxGrouped[$groupKey]["Code CIGREF"] = $codeCigrefInit
            }
        }
        $xlsxGrouped[$groupKey]["Nb jours"] += $jours
    }
    Write-Progress -Id 4 -Activity "Generation DETERMINE" -Completed

    # Tri par NOM Prénom (Col. C) avant export
    $determineRows = @($determineRows | Sort-Object nomAffichage)

    $csvDetermine = Join-Path $exportsDir ("EtatsFinanciers_DETERMINE_" + $moisKey + "_" + $runStamp + ".csv")
    Export-CsvStrict -Path $csvDetermine `
        -Headers @("idCle","exercice","mois","dateSaisie","accountIdJira","nomAffichage","matricule","nom","prenom",
                   "typeRessource","societe","direction","codeDomaineSIMP","codeBudget","libelleBudget",
                   "cleInitiative","libelleInitiative","codeFDR","codeCigrefInit","heuresTotales","joursTotaux") `
        -Rows $determineRows

    # Construction + tri des lignes XLSX par NOM Prénom (Col. C)
    $xlsxSuiviRows = New-Object System.Collections.ArrayList
    foreach ($gKey in $xlsxGrouped.Keys) {
        $r = $xlsxGrouped[$gKey]
        $r["Nb jours"] = [math]::Round([double]$r["Nb jours"], 2)
        [void]$xlsxSuiviRows.Add([pscustomobject]$r)
    }
    $xlsxSuiviRows = @($xlsxSuiviRows | Sort-Object "NOM Prenom")

    $xlsxSuiviFile = Join-Path $exportsDir ("EXPORT_JIRADOT_SUIVI-ACTIVITE_" + $runStamp + ".xlsx")
    Export-GenericXlsx -Path $xlsxSuiviFile -SheetName "Suivi Activite" `
        -Headers @("Annee","Mois","NOM Prenom","Type ressource","Societe","Direction","Service","Matricule",
                   "Code domaine Ressource","Code SIMP","Libelle SIMP","Code - Libelle initiative",
                   "Nb jours","Code CIGREF","Code Gepetto","Identifiant Jira") `
        -Rows $xlsxSuiviRows
    Write-Info ("  SUIVI-ACTIVITE : " + $xlsxSuiviRows.Count + " lignes agregees (Collaborateur x Initiative) — triees par NOM Prenom")
}

# ============================================================
# 10b. FORMAT SIMP : CSV + XLSX EXPORT_JIRADOT_SIMP
# ============================================================
if ($Format -eq "SIMP" -or $Format -eq "BOTH") {

    $simpGrouped = @{}
    $totalWL2    = $worklogs.Count
    $idxWL2      = 0

    foreach ($wl in $worklogs) {
        $idxWL2++
        if ($idxWL2 % 1000 -eq 0 -or $idxWL2 -eq $totalWL2) {
            Write-Progress -Id 5 -Activity "Generation SIMP" `
                -Status ("Worklog " + $idxWL2 + "/" + $totalWL2 + " — " + $simpGrouped.Count + " groupes") `
                -PercentComplete ([math]::Round($idxWL2 / $totalWL2 * 100))
        }

        $accId  = [string]$wl.author.accountId
        $issId  = [string]$wl.issue.id
        $jUser  = $jiraUserCache[$accId]
        $uMail  = ""; if ($jUser -and $jUser.EmailAddress) { $uMail = [string]$jUser.EmailAddress }
        $dName  = ""; if ($jUser -and $jUser.DisplayName)  { $dName = [string]$jUser.DisplayName }

        $cmdb   = Resolve-Cmdb -AccId $accId -UserMail $uMail -DispName $dName
        $jIssue = $jiraIssueCache[[string]$issId]

        $codeSimpRessource = "AUTRE"
        if ($cmdb -and $cmdb.CodeDomaineSIMP) { $codeSimpRessource = $cmdb.CodeDomaineSIMP }
        $typeRes = "Regie"
        if ($cmdb -and $cmdb.TypeRessource)   { $typeRes = $cmdb.TypeRessource }

        $bKey       = if ($jIssue -and $jIssue.BudgetKey)     { $jIssue.BudgetKey }     else { "" }
        $bSum       = if ($jIssue -and $jIssue.BudgetSummary) { $jIssue.BudgetSummary } else { "" }
        $budgetSumm = Resolve-LibelleSIMP -BudgetKey $bKey -BudgetSummary $bSum

        $heures   = [double]$wl.timeSpentSeconds / 3600.0
        $groupKey = ($codeSimpRessource + "|" + $bKey + "|" + $typeRes)

        if (-not $simpGrouped.ContainsKey($groupKey)) {
            $simpGrouped[$groupKey] = [ordered]@{
                CodeSimpRessource = $codeSimpRessource
                BudgetKey         = $bKey
                BudgetSummary     = $budgetSumm
                TypeRessource     = $typeRes
                TotalHeures       = 0.0
            }
        }
        if ([string]::IsNullOrWhiteSpace($simpGrouped[$groupKey]["BudgetSummary"]) -and
            -not [string]::IsNullOrWhiteSpace($budgetSumm)) {
            $simpGrouped[$groupKey]["BudgetSummary"] = $budgetSumm
        }
        $simpGrouped[$groupKey]["TotalHeures"] += $heures
    }
    Write-Progress -Id 5 -Activity "Generation SIMP" -Completed
    Write-Info ("  SIMP : " + $simpGrouped.Count + " groupes d agregation generes")

    $simpRows = New-Object System.Collections.ArrayList
    foreach ($gKey in $simpGrouped.Keys) {
        $sg   = $simpGrouped[$gKey]
        $totH = [math]::Round($sg.TotalHeures, 5)
        $totJ = [math]::Round($sg.TotalHeures / 8.0, 5)
        [void]$simpRows.Add([pscustomobject]@{
            idCle             = ($sg.CodeSimpRessource + "_" + $sg.BudgetKey + "_" + $sg.TypeRessource + "_" + $fyKey + "_" + $mKey)
            codeSIMPRessource = $sg.CodeSimpRessource
            codeBudget        = $sg.BudgetKey
            libelleBudget     = $sg.BudgetSummary
            typeRessource     = $sg.TypeRessource
            exercice          = $fyKey
            mois              = $mKey
            heuresTotales     = $totH
            joursTotaux       = $totJ
        })
    }

    $csvSimp = Join-Path $exportsDir ("EtatsFinanciers_SIMP_" + $moisKey + "_" + $runStamp + ".csv")
    Export-CsvStrict -Path $csvSimp `
        -Headers @("idCle","codeSIMPRessource","codeBudget","libelleBudget","typeRessource","exercice","mois","heuresTotales","joursTotaux") `
        -Rows $simpRows

    $xlsxSimpRows = New-Object System.Collections.ArrayList
    foreach ($gKey in $simpGrouped.Keys) {
        $sg     = $simpGrouped[$gKey]
        $totJ   = [math]::Round($sg.TotalHeures / 8.0, 2)
        $actInfo = "(Sans budget)"
        if     ($sg.BudgetKey -and $sg.BudgetSummary) { $actInfo = $sg.BudgetKey + " - " + $sg.BudgetSummary }
        elseif ($sg.BudgetKey)                         { $actInfo = $sg.BudgetKey }
        [void]$xlsxSimpRows.Add([pscustomobject]@{
            "Contribution"          = $sg.CodeSimpRessource
            "Activite informatique" = $actInfo
            "Domaine applicatif"    = $sg.TypeRessource
            "Indicateur"            = "Nb jours"
            "Annee"                 = $From.Year
            "Periode"               = "M" + $From.Month.ToString("D2") + "/" + $From.Year
            "Donnees"               = $totJ
        })
    }

    $xlsxSimpFile = Join-Path $exportsDir ("EXPORT_JIRADOT_SIMP_" + $runStamp + ".xlsx")
    Export-GenericXlsx -Path $xlsxSimpFile -SheetName "SIMP" `
        -Headers @("Contribution","Activite informatique","Domaine applicatif","Indicateur","Annee","Periode","Donnees") `
        -Rows $xlsxSimpRows
    Write-Info ("  SIMP XLSX : " + $xlsxSimpRows.Count + " lignes generees")
}

# ============================================================
# 11. RÉSUMÉ FINAL ET CONTRÔLE DE COUVERTURE
# ============================================================
Write-Info ""
Write-Info "=========================================="
Write-Info "=== EXTRACTION TERMINEE — RESUME v2.5 ==="
Write-Info "=========================================="
Write-Info ("  Periode          : " + $From.ToString("dd/MM/yyyy") + " -> " + $To.ToString("dd/MM/yyyy"))
Write-Info ("  Perimetre        : " + $Format)
Write-Info ("  Worklogs Tempo   : " + $worklogs.Count)
Write-Info ("  Comptes Jira     : " + $jiraUserCache.Count)
Write-Info ("  Tickets resolus  : " + $jiraIssueCache.Count)
Write-Info ("  Fiches Assets RP : " + $assetsUsersByAccountId.Count + " (accountId) | " +
                                      $assetsUsersByEmail.Count     + " (email) | " +
                                      $assetsUsersByName.Count      + " (nom)")
Write-Info ""
Write-Info ("  Fichiers generes dans : " + $exportsDir)
if ($Format -eq "DETERMINE" -or $Format -eq "BOTH") {
    Write-Info ("    CSV  : EtatsFinanciers_DETERMINE_" + $moisKey + "_" + $runStamp + ".csv")
    Write-Info ("    XLSX : EXPORT_JIRADOT_SUIVI-ACTIVITE_" + $runStamp + ".xlsx")
}
if ($Format -eq "SIMP" -or $Format -eq "BOTH") {
    Write-Info ("    CSV  : EtatsFinanciers_SIMP_" + $moisKey + "_" + $runStamp + ".csv")
    Write-Info ("    XLSX : EXPORT_JIRADOT_SIMP_" + $runStamp + ".xlsx")
}
Write-Info ""
Write-Info ("  Cache local dans  : " + $cacheDir)
Write-Info ("    assets_referentiel_personne.json : age = " + (Get-CacheAge $assetsCacheFile)      + "h")
Write-Info ("    assets_domaine_simp.json         : age = " + (Get-CacheAge $domaineSimpCacheFile)  + "h")
Write-Info ("    jira_users_cache.json            : age = " + (Get-CacheAge $usersCacheFile)        + "h")
Write-Info ("    jira_issues_hierarchy.json       : age = " + (Get-CacheAge $issuesCacheFile)       + "h")
Write-Info ("    worklogs_tempo_" + $moisKey + ".json")
Write-Info ""

# Contrôle couverture CMDB + Libellé SIMP + Code CIGREF
$matchedById=0; $matchedByEmail=0; $matchedByName=0; $notMatched=0
$libelleRempli=0; $libelleVide=0; $cigrefRempli=0

foreach ($wl in $worklogs) {
    $accId = [string]$wl.author.accountId
    $jUser = $jiraUserCache[$accId]
    $uMail = ""; if ($jUser -and $jUser.EmailAddress) { $uMail = [string]$jUser.EmailAddress }
    $dName = ""; if ($jUser -and $jUser.DisplayName)  { $dName = [string]$jUser.DisplayName }

    if      ($accId -and $assetsUsersByAccountId.ContainsKey($accId))             { $matchedById++ }
    elseif  ($uMail -and $assetsUsersByEmail.ContainsKey($uMail))                 { $matchedByEmail++ }
    elseif  ($dName -and $assetsUsersByName.ContainsKey($dName.Trim().ToLower())) { $matchedByName++ }
    else {
        $notMatched++
        Write-Log ("Non rattache CMDB : accountId=" + $accId + " | displayName=" + $dName + " | email=" + $uMail) "WARN"
    }

    $issId  = [string]$wl.issue.id
    $jIssue = $jiraIssueCache[$issId]
    $bKey = if ($jIssue -and $jIssue.BudgetKey)     { $jIssue.BudgetKey }     else { "" }
    $bSum = if ($jIssue -and $jIssue.BudgetSummary) { $jIssue.BudgetSummary } else { "" }
    $lib  = Resolve-LibelleSIMP -BudgetKey $bKey -BudgetSummary $bSum
    if (-not [string]::IsNullOrWhiteSpace($lib)) { $libelleRempli++ } else { $libelleVide++ }

    $cig = if ($jIssue -and $jIssue.CodeCigrefInit) { $jIssue.CodeCigrefInit } else { "" }
    if (-not [string]::IsNullOrWhiteSpace($cig)) { $cigrefRempli++ }
}

$total  = $worklogs.Count
$pctCov = if ($total -gt 0) { [math]::Round(($total - $notMatched) / $total * 100, 1) } else { 0 }
$pctLib = if (($libelleRempli + $libelleVide) -gt 0) {
              [math]::Round($libelleRempli / ($libelleRempli + $libelleVide) * 100, 1) } else { 100 }
$pctCig = if ($total -gt 0) { [math]::Round($cigrefRempli / $total * 100, 1) } else { 0 }

Write-Info ("  Couverture CMDB RP :")
Write-Info ("    Cle #1 accountId     : " + $matchedById    + " worklogs")
Write-Info ("    Cle #2 email         : " + $matchedByEmail + " worklogs")
Write-Info ("    Cle #3 nom           : " + $matchedByName  + " worklogs")
Write-Info ("    Non rattaches        : " + $notMatched      + " worklogs")
Write-Info ("    Taux de couverture   : " + $pctCov + "%")
Write-Info ""
Write-Info ("  Couverture Libelle SIMP depuis Tickets BUD (Col. K) :")
Write-Info ("    Remplis              : " + $libelleRempli + " worklogs")
Write-Info ("    Vides                : " + $libelleVide   + " worklogs")
Write-Info ("    Taux de remplissage  : " + $pctLib + "%")
Write-Info ""
Write-Info ("  Couverture Code CIGREF Initiative (Col. N) :")
Write-Info ("    Remplis              : " + $cigrefRempli + " worklogs")
Write-Info ("    Taux de remplissage  : " + $pctCig + "%")

if ($notMatched -gt 0) {
    Write-Warn ("  " + $notMatched + " worklog(s) non rattache(s) CMDB — detail dans : " + $logFile)
}

Write-Info ""
Write-Info ("  Delai API Assets final : " + $global:assetsApiDelayMs + "ms")
Write-Info ("  Log complet dans       : " + $logFile)
Write-Info "=========================================="
Write-Log  ("=== FIN EXECUTION " + $scriptName + " v2.5 — " + (Get-Date -Format "yyyy-MM-dd HH:mm:ss") + " ===") "INFO"
