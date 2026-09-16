<#
Liste des espaces Jira (projets) sur lesquels sont habilitées des personnes (liste fournie)
- Se base sur les rôles de projet (GET /project/{key}/role + actors)
- TLS 1.2
- Proxy système Windows (PAC/WPAD/McAfee) + creds Windows
- Stockage local chiffré (DPAPI) de l'email+token (Export-Clixml)
- 3 exports CSV:
    * C:\Temp\habilitations_detail.csv
    * C:\Temp\habilitations_resume.csv
    * C:\Temp\habilitations_non_trouvees.csv

Usage:
  .\listeHabilitationsJira.ps1                     # demande de choisir le CSV
  .\listeHabilitationsJira.ps1 -ResetCreds         # + ressaisie email/token
  .\listeHabilitationsJira.ps1 -InputCsv "C:\chemin\Personnes_habilitées.csv"
#>

param(
    [string]$InputCsv,
    [switch]$ResetCreds
)

# ----------------------------
# Fonction sélection du fichier CSV
# ----------------------------
function Select-InputCsvFile([string]$initialDir) {

    try {
        Add-Type -AssemblyName System.Windows.Forms | Out-Null

        $dlg = New-Object System.Windows.Forms.OpenFileDialog
        $dlg.Title = "Sélectionner le fichier CSV des personnes habilitées"
        $dlg.Filter = "Fichiers CSV (*.csv)|*.csv|Tous les fichiers (*.*)|*.*"
        $dlg.Multiselect = $false

        if ($initialDir -and (Test-Path $initialDir)) {
            $dlg.InitialDirectory = $initialDir
        }

        $result = $dlg.ShowDialog()
        if ($result -ne [System.Windows.Forms.DialogResult]::OK) {
            throw "Aucun fichier sélectionné."
        }

        return $dlg.FileName
    }
    catch {
        Write-Warning "Boîte de dialogue indisponible, saisie du chemin en fallback."
        $p = Read-Host "Chemin complet du CSV"
        if (-not (Test-Path $p)) { throw "Fichier introuvable: $p" }
        return $p
    }
}

# Si -InputCsv n'est pas fourni, on demande via une fenêtre
if ([string]::IsNullOrWhiteSpace($InputCsv)) {
    # Dossier par défaut : ton dossier OneDrive PowerShell
    $defaultDir = "C:\Users\GUEDJ-F\OneDrive - Harmonie Mutuelle\Documents\powershell"
    $InputCsv = Select-InputCsvFile -initialDir $defaultDir
}

Write-Host "Fichier CSV sélectionné : $InputCsv" -ForegroundColor Cyan

# ----------------------------
# Runtime
# ----------------------------
$ErrorActionPreference = "Stop"
$ProgressPreference = "SilentlyContinue"

[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

# Proxy système Windows + creds Windows (utile derrière McAfee)
$systemProxy = [System.Net.WebRequest]::GetSystemWebProxy()
$systemProxy.Credentials = [System.Net.CredentialCache]::DefaultNetworkCredentials

function Get-ProxyForUrl($url) {
    try {
        $u = [Uri]$url
        $p = $systemProxy.GetProxy($u)
        if ($p -and $p.AbsoluteUri -ne $u.AbsoluteUri) { return $p.AbsoluteUri }
        return $null
    } catch { return $null }
}

function Invoke-JiraRest($method, $url, $headers, $body = $null) {
    $proxyUri = Get-ProxyForUrl $url

    $params = @{
        Method      = $method
        Uri         = $url
        Headers     = $headers
        ErrorAction = "Stop"
    }
    if ($null -ne $body) { $params.Body = $body }

    if ($proxyUri) {
        $params.Proxy = $proxyUri
        $params.ProxyUseDefaultCredentials = $true
    }

    try {
        return Invoke-RestMethod @params
    } catch {
        $msg = $_.Exception.Message
        $details = $null
        if ($_.ErrorDetails -and $_.ErrorDetails.Message) { $details = $_.ErrorDetails.Message }
        if ($_.Exception.InnerException) { $msg = "$msg | Inner: $($_.Exception.InnerException.Message)" }

        if ($details) { throw "HTTP failed: $method $url => $msg | Details: $details" }
        throw "HTTP failed: $method $url => $msg"
    }
}

# ----------------------------
# Config
# ----------------------------
$siteUrl = "https://jiradot.atlassian.net"

$exportDetailPath   = "C:\Temp\habilitations_detail.csv"
$exportSummaryPath  = "C:\Temp\habilitations_resume.csv"
$exportNotFoundPath = "C:\Temp\habilitations_non_trouvees.csv"

# DPAPI credential file (user scope)
$credDir  = Join-Path $env:APPDATA "Jira"
$credPath = Join-Path $credDir "jira-cloud-cred.clixml"

function Get-JiraCredential($credPath, $reset) {
    $dir = Split-Path $credPath
    if (-not (Test-Path $dir)) {
        New-Item -ItemType Directory -Path $dir | Out-Null
    }

    if (-not $reset -and (Test-Path $credPath)) {
        try {
            return Import-Clixml -Path $credPath
        } catch {
            Write-Warning "Impossible de relire le fichier credential, nouvelle saisie requise. Détail: $_"
        }
    }

    $email = Read-Host "Email Atlassian (ex: prenom.nom@domaine.fr)"
    if ($email -match '^\[(.+?)\]\(mailto:(.+?)\)$') { $email = $Matches[2] }
    $email = ($email -replace '^mailto:', '').Trim()

    $secureToken = Read-Host "API Token Atlassian (saisie masquée)" -AsSecureString

    $cred = New-Object System.Management.Automation.PSCredential($email, $secureToken)
    $cred | Export-Clixml -Path $credPath

    Write-Host "Identifiants sauvegardés dans: $credPath" -ForegroundColor Green
    return $cred
}

function Normalize-Text([string]$s) {
    if ([string]::IsNullOrWhiteSpace($s)) { return "" }

    $s = $s.Trim()
    $s = $s -replace '\s+', ' '

    $formD = $s.Normalize([Text.NormalizationForm]::FormD)
    $sb = New-Object System.Text.StringBuilder
    foreach ($ch in $formD.ToCharArray()) {
        $cat = [Globalization.CharUnicodeInfo]::GetUnicodeCategory($ch)
        if ($cat -ne [Globalization.UnicodeCategory]::NonSpacingMark) {
            [void]$sb.Append($ch)
        }
    }
    $clean = $sb.ToString().Normalize([Text.NormalizationForm]::FormC)

    $clean = $clean.ToLowerInvariant()
    $clean = $clean -replace '[^a-z0-9 @._-]', ' '
    $clean = $clean -replace '\s+', ' '
    return $clean.Trim()
}

function Read-ActorList([string]$path) {
    if (-not (Test-Path $path)) { throw "Fichier introuvable: $path" }

    try {
        $rows = Import-Csv -LiteralPath $path -Encoding UTF8
        if ($rows -and ($rows[0].PSObject.Properties.Name -contains "Actor Name")) {
            $names = $rows |
                ForEach-Object { $_."Actor Name" } |
                Where-Object { $_ -and $_.Trim() -and $_.Trim() -ne "Actor Name" } |
                ForEach-Object { $_.Trim() } |
                Sort-Object -Unique
            if ($names.Count -gt 0) { return $names }
        }
    } catch {
        # fallback plus bas
    }

    $lines = Get-Content -LiteralPath $path -Encoding UTF8 |
        ForEach-Object { $_.Trim() } |
        Where-Object { $_ -and $_ -ne "Actor Name" } |
        Sort-Object -Unique
    return $lines
}

# ----------------------------
# Auth header
# ----------------------------
$jiraCred = Get-JiraCredential -credPath $credPath -reset:$ResetCreds

$bstr = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($jiraCred.Password)
try { $apiTokenPlain = [Runtime.InteropServices.Marshal]::PtrToStringBSTR($bstr) }
finally { [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($bstr) }

$pair   = "$($jiraCred.UserName)`:$apiTokenPlain"
$bytes  = [Text.Encoding]::UTF8.GetBytes($pair)
$base64 = [Convert]::ToBase64String($bytes)

$headers = @{
    Authorization = "Basic $base64"
    Accept        = "application/json"
    "Content-Type"= "application/json"
}

# ----------------------------
# Test API
# ----------------------------
Write-Host "Test API /myself ..." -ForegroundColor Cyan
$me = Invoke-JiraRest "GET" "$siteUrl/rest/api/3/myself" $headers
Write-Host ("API OK - " + $me.displayName) -ForegroundColor Green

# ----------------------------
# Chargement liste personnes
# ----------------------------
Write-Host "Lecture liste personnes: $InputCsv" -ForegroundColor Cyan
$requestedActors = Read-ActorList $InputCsv
Write-Host ("Nb entrées liste: " + $requestedActors.Count) -ForegroundColor Green

$requestedIndex = @{}
foreach ($n in $requestedActors) {
    $k = Normalize-Text $n
    if (-not [string]::IsNullOrWhiteSpace($k)) {
        if (-not $requestedIndex.ContainsKey($k)) { $requestedIndex[$k] = $n }
    }
}

# ----------------------------
# Récupération projets (status=live)
# ----------------------------
Write-Host "Récupération des projets (status=live) ..." -ForegroundColor Cyan
$allProjects = @()
$startAt = 0
$maxResults = 50

while ($true) {
    $url = "$siteUrl/rest/api/3/project/search?startAt=$startAt&maxResults=$maxResults&status=live"
    $resp = Invoke-JiraRest "GET" $url $headers

    if ($resp.values) { $allProjects += $resp.values }
    $startAt += ($resp.values | Measure-Object).Count
    if ($startAt -ge $resp.total) { break }
}

Write-Host ("Nombre de projets live: " + $allProjects.Count) -ForegroundColor Green

# ----------------------------
# Extraction rôles / acteurs & filtrage
# ----------------------------
Write-Host "Analyse rôles/acteurs par projet (filtrage sur ta liste)..." -ForegroundColor Cyan

$resultDetail = @()

foreach ($p in $allProjects) {
    $projectKey  = $p.key
    $projectName = $p.name
    if ([string]::IsNullOrWhiteSpace($projectKey)) { continue }

    Write-Host "Projet $projectKey - $projectName" -ForegroundColor Yellow

    $rolesResp = $null
    try {
        $rolesResp = Invoke-JiraRest "GET" "$siteUrl/rest/api/3/project/$projectKey/role" $headers
    } catch {
        Write-Warning "Impossible de lire les rôles pour $projectKey : $_"
        continue
    }

    foreach ($prop in $rolesResp.PSObject.Properties) {
        $roleUrl = $prop.Value

        $roleDetail = $null
        try {
            $roleDetail = Invoke-JiraRest "GET" $roleUrl $headers
        } catch {
            Write-Warning "Détail rôle non récupéré ($($prop.Name)) pour $projectKey : $_"
            continue
        }

        foreach ($actor in $roleDetail.actors) {
            $jiraActorName = $actor.displayName
            if ([string]::IsNullOrWhiteSpace($jiraActorName)) { $jiraActorName = $actor.name }

            $norm = Normalize-Text $jiraActorName
            if ([string]::IsNullOrWhiteSpace($norm)) { continue }

            if ($requestedIndex.ContainsKey($norm)) {
                $requestedLabel = $requestedIndex[$norm]

                $accountId = $null
                if ($actor.actorUser -and $actor.actorUser.accountId) { $accountId = $actor.actorUser.accountId }
                elseif ($actor.accountId) { $accountId = $actor.accountId }

                $groupName = $null
                if ($actor.actorGroup -and $actor.actorGroup.name) { $groupName = $actor.actorGroup.name }

                $resultDetail += [PSCustomObject]@{
                    "Requested Actor"   = $requestedLabel
                    "Jira Actor Name"   = $jiraActorName
                    "Actor Type"        = $actor.type
                    "Actor AccountId"   = $accountId
                    "Actor Group"       = $groupName
                    "Project Key"       = $projectKey
                    "Project Name"      = $projectName
                    "Role Name"         = $roleDetail.name
                    "Role Id"           = $roleDetail.id
                }
            }
        }
    }
}

Write-Host ("Lignes détail habilitations: " + $resultDetail.Count) -ForegroundColor Green

# ----------------------------
# Résumé par personne (liste projets)
# ----------------------------
$summary = @()
if ($resultDetail.Count -gt 0) {
    $summary = $resultDetail |
        Group-Object "Requested Actor" |
        ForEach-Object {
            $actor = $_.Name
            $projects = $_.Group |
                Select-Object -Property "Project Key","Project Name" -Unique |
                Sort-Object "Project Key"

            [PSCustomObject]@{
                "Requested Actor" = $actor
                "Projects Count"  = $projects.Count
                "Projects (Keys)" = ($projects | ForEach-Object { $_."Project Key" }) -join ";"
                "Projects (Names)"= ($projects | ForEach-Object { $_."Project Name" }) -join ";"
            }
        } |
        Sort-Object "Requested Actor"
}

# ----------------------------
# Personnes non trouvées
# ----------------------------
$foundActors = @()
if ($resultDetail.Count -gt 0) {
    $foundActors = $resultDetail | Select-Object -ExpandProperty "Requested Actor" -Unique
}

$notFound = $requestedActors |
    Where-Object { $_ -notin $foundActors } |
    Sort-Object -Unique |
    ForEach-Object { [PSCustomObject]@{ "Requested Actor" = $_ } }

# ----------------------------
# Export CSV
# ----------------------------
$exportDir = Split-Path $exportDetailPath
if (-not (Test-Path $exportDir)) { New-Item -ItemType Directory -Path $exportDir | Out-Null }

Write-Host "Export CSV détail : $exportDetailPath" -ForegroundColor Cyan
$resultDetail | Export-Csv -Path $exportDetailPath -NoTypeInformation -Encoding UTF8

Write-Host "Export CSV résumé : $exportSummaryPath" -ForegroundColor Cyan
$summary | Export-Csv -Path $exportSummaryPath -NoTypeInformation -Encoding UTF8

Write-Host "Export CSV non trouvés : $exportNotFoundPath" -ForegroundColor Cyan
$notFound | Export-Csv -Path $exportNotFoundPath -NoTypeInformation -Encoding UTF8

Write-Host "Terminé." -ForegroundColor Green
Write-Host " - Détail      : $exportDetailPath"
Write-Host " - Résumé      : $exportSummaryPath"
Write-Host " - Non trouvés : $exportNotFoundPath"
Write-Host " - Creds       : $credPath"