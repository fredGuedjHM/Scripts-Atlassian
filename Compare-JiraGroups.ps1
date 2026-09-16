<#
Compare-JiraGroups.ps1
.SYNOPSIS
    Compare les membres de deux groupes Jira Cloud (G1 et G2) et identifie les utilisateurs de G1 absents de G2.
.DESCRIPTION
    1. Demande interactivement (ou par paramètres) le nom du groupe source (G1) et du groupe cible (G2).
    2. Récupère la liste complète des membres de G1 et G2 via l'API REST Jira Cloud v3 (avec gestion de pagination).
    3. Affiche les détails des membres (AccountId, DisplayName, Email, Statut).
    4. Calcule le delta : membres présents dans G1 mais absents de G2 (pour préparer leur ajout / fusion).
    5. Génère un export CSV récapitulatif avec encodage UTF-8 BOM.
.VERSION
    1.0 — 2026-09-14
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [string]$Group1,

    [Parameter(Mandatory = $false)]
    [string]$Group2,

    [Parameter(Mandatory = $false)]
    [string]$CredentialsPath,

    [switch]$ExportCsv
)

[System.Threading.Thread]::CurrentThread.CurrentCulture   = [System.Globalization.CultureInfo]::GetCultureInfo("fr-FR")
[System.Threading.Thread]::CurrentThread.CurrentUICulture = [System.Globalization.CultureInfo]::GetCultureInfo("fr-FR")

Set-StrictMode -Off
$ErrorActionPreference = "Continue"

# ============================================================
# 0. CONSTANTES & CHEMINS
# ============================================================
$scriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
if (-not $scriptRoot) { $scriptRoot = Get-Location }

$secretsDir = Join-Path $scriptRoot "secrets"
$exportsDir = Join-Path $scriptRoot "exports"
$runStamp   = Get-Date -Format "yyyyMMdd_HHmmss"

if (-not (Test-Path $exportsDir)) { New-Item -ItemType Directory -Path $exportsDir -Force | Out-Null }

$ProxyUrl                   = ""
$UseSystemProxy             = $true
$ProxyUseDefaultCredentials = $true

# ============================================================
# 1. HELPERS PROXY & REST API
# ============================================================
function Write-Info([string]$Message) {
    Write-Host ("[INFO] " + $Message) -ForegroundColor Cyan
}
function Write-Warn([string]$Message) {
    Write-Warning $Message
}
function Write-ErrLog([string]$Message) {
    Write-Error $Message
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

function Invoke-JiraApiGet([string]$Url, [hashtable]$Headers) {
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    $params = @{
        Method          = 'GET'
        Uri             = $Url
        Headers         = $Headers
        ContentType     = 'application/json'
        UseBasicParsing = $true
        ErrorAction     = 'Stop'
    }
    $px = Get-ProxyParams $Url
    foreach ($k in $px.Keys) { $params[$k] = $px[$k] }
    
    $resp   = Invoke-WebRequest @params
    $stream = $resp.RawContentStream
    $stream.Position = 0
    $reader = New-Object System.IO.StreamReader($stream, [System.Text.Encoding]::UTF8, $true)
    $raw    = $reader.ReadToEnd()
    $reader.Close()
    return ($raw | ConvertFrom-Json)
}

function Get-JiraGroupMembers {
    param(
        [Parameter(Mandatory = $true)][string]$GroupName,
        [Parameter(Mandatory = $true)][string]$BaseUrl,
        [Parameter(Mandatory = $true)][hashtable]$Headers
    )

    $members = New-Object System.Collections.ArrayList
    $startAt = 0
    $maxResults = 50
    $isLast = $false

    Write-Info ("Recuperation des membres du groupe : [" + $GroupName + "]...")

    do {
        $escapedGroup = [Uri]::EscapeDataString($GroupName)
        $url = $BaseUrl + "/rest/api/3/group/member?groupname=" + $escapedGroup +
               "&includeInactiveUsers=true&startAt=" + $startAt + "&maxResults=" + $maxResults

        try {
            $resp = Invoke-JiraApiGet -Url $url -Headers $Headers
        } catch {
            Write-ErrLog ("Impossible de recuperer les membres du groupe '" + $GroupName + "' : " + $_.Exception.Message)
            return $null
        }

        if ($resp -and $resp.values) {
            foreach ($u in $resp.values) {
                [void]$members.Add([pscustomobject]@{
                    AccountId    = [string]$u.accountId
                    DisplayName  = Fix-Encoding ([string]$u.displayName)
                    EmailAddress = if ($u.emailAddress) { [string]$u.emailAddress } else { "(Masque/Non renseigne)" }
                    AccountType  = [string]$u.accountType
                    Active       = [bool]$u.active
                })
            }
        }

        if ($resp.isLast -ne $null) {
            $isLast = [bool]$resp.isLast
        } else {
            if (-not $resp.values -or $resp.values.Count -lt $maxResults) { $isLast = $true }
        }

        $startAt += $maxResults
    } while (-not $isLast)

    Write-Info ("  -> " + $members.Count + " membre(s) trouve(s) dans [" + $GroupName + "]")
    return ,$members
}

# ============================================================
# 2. CHARGEMENT DES CREDENTIALS
# ============================================================
if (-not $CredentialsPath) {
    $CredentialsPath = Join-Path $secretsDir "jira-jiradot.cred.xml"
}

if (-not (Test-Path $CredentialsPath)) {
    throw "Fichier de credentials Jira introuvable : $CredentialsPath"
}

$jiraData    = Import-Clixml -Path $CredentialsPath
$jiraBaseUrl = ([string]$jiraData.JiraBaseUrl).TrimEnd("/")
$jiraCred    = $jiraData.Credential
$b64         = [Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes($jiraCred.UserName + ":" + $jiraCred.GetNetworkCredential().Password))
$jiraHeaders = @{ Authorization = "Basic " + $b64; Accept = "application/json" }

# ============================================================
# 3. SAISIE DES GROUPES
# ============================================================
Write-Host ""
Write-Host ("=" * 70) -ForegroundColor DarkCyan
Write-Host "  COMPARAISON ET FUSION DE GROUPES D'HABILITATION JIRA" -ForegroundColor White
Write-Host ("=" * 70) -ForegroundColor DarkCyan
Write-Host ""

if ([string]::IsNullOrWhiteSpace($Group1)) {
    $Group1 = Read-Host "Entrez le nom du groupe source G1 (a fusionner/migrer)"
}
if ([string]::IsNullOrWhiteSpace($Group2)) {
    $Group2 = Read-Host "Entrez le nom du groupe cible G2 (destinataire)"
}

$Group1 = $Group1.Trim()
$Group2 = $Group2.Trim()

if ([string]::IsNullOrWhiteSpace($Group1) -or [string]::IsNullOrWhiteSpace($Group2)) {
    throw "Les deux noms de groupes (G1 et G2) doivent etre renseignes."
}

if ($Group1 -ieq $Group2) {
    throw "Le groupe G1 et le groupe G2 sont identiques ($Group1)."
}

# ============================================================
# 4. RÉCUPÉRATION DES MEMBRES G1 ET G2
# ============================================================
$membersG1 = Get-JiraGroupMembers -GroupName $Group1 -BaseUrl $jiraBaseUrl -Headers $jiraHeaders
if ($null -eq $membersG1) { return }

$membersG2 = Get-JiraGroupMembers -GroupName $Group2 -BaseUrl $jiraBaseUrl -Headers $jiraHeaders
if ($null -eq $membersG2) { return }

# ============================================================
# 5. COMPARAISON & CALCUL DU DELTA (G1 \ G2)
# ============================================================
# Indexation de G2 par AccountId
$g2AccountIds = @{}
foreach ($u in $membersG2) {
    $g2AccountIds[$u.AccountId] = $u
}

$missingInG2 = New-Object System.Collections.ArrayList
$alreadyInG2 = New-Object System.Collections.ArrayList

foreach ($u in $membersG1) {
    if (-not $g2AccountIds.ContainsKey($u.AccountId)) {
        [void]$missingInG2.Add($u)
    } else {
        [void]$alreadyInG2.Add($u)
    }
}

# ============================================================
# 6. AFFICHAGE DES RÉSULTATS
# ============================================================
Write-Host ""
Write-Host ("-" * 70) -ForegroundColor DarkGray
Write-Host ("1. MEMBRES DU GROUPE SOURCE G1 : [" + $Group1 + "] (" + $membersG1.Count + ")") -ForegroundColor Yellow
Write-Host ("-" * 70) -ForegroundColor DarkGray
$membersG1 | Sort-Object DisplayName | Format-Table -Property @(
    @{Label="Nom / Prenom (DisplayName)"; Expression={$_.DisplayName}; Width=32},
    @{Label="Email"; Expression={$_.EmailAddress}; Width=35},
    @{Label="Actif"; Expression={if ($_.Active) {"Oui"} else {"Non"}}; Width=6},
    @{Label="User ID (AccountId)"; Expression={$_.AccountId}; Width=30}
) | Out-Host

Write-Host ("-" * 70) -ForegroundColor DarkGray
Write-Host ("2. MEMBRES DU GROUPE CIBLE G2 : [" + $Group2 + "] (" + $membersG2.Count + ")") -ForegroundColor Yellow
Write-Host ("-" * 70) -ForegroundColor DarkGray
$membersG2 | Sort-Object DisplayName | Format-Table -Property @(
    @{Label="Nom / Prenom (DisplayName)"; Expression={$_.DisplayName}; Width=32},
    @{Label="Email"; Expression={$_.EmailAddress}; Width=35},
    @{Label="Actif"; Expression={if ($_.Active) {"Oui"} else {"Non"}}; Width=6},
    @{Label="User ID (AccountId)"; Expression={$_.AccountId}; Width=30}
) | Out-Host

Write-Host ("=" * 70) -ForegroundColor Green
Write-Host ("3. BILAN : COMPTES PRESENTS DANS G1 MAIS ABSENTS DE G2 (" + $missingInG2.Count + ")") -ForegroundColor Green
Write-Host ("   -> Ces comptes doivent etre ajoutes dans [" + $Group2 + "] pour finaliser la fusion.") -ForegroundColor Gray
Write-Host ("=" * 70) -ForegroundColor Green

if ($missingInG2.Count -gt 0) {
    $missingInG2 | Sort-Object DisplayName | Format-Table -Property @(
        @{Label="Nom / Prenom (DisplayName)"; Expression={$_.DisplayName}; Width=32},
        @{Label="Email"; Expression={$_.EmailAddress}; Width=35},
        @{Label="Actif"; Expression={if ($_.Active) {"Oui"} else {"Non"}}; Width=6},
        @{Label="User ID (AccountId)"; Expression={$_.AccountId}; Width=30}
    ) | Out-Host
} else {
    Write-Host "Tous les membres de G1 sont deja presents dans G2 ! Aucun ajout necessaire." -ForegroundColor Green
}

# ============================================================
# 7. EXPORT CSV (Optionnel / Automatique)
# ============================================================
$sanitizedG1 = $Group1 -replace '[\\/:*?"<>|]', '_'
$sanitizedG2 = $Group2 -replace '[\\/:*?"<>|]', '_'
$csvPath     = Join-Path $exportsDir ("Delta_Groupes_" + $sanitizedG1 + "_VERS_" + $sanitizedG2 + "_" + $runStamp + ".csv")

$csvRows = New-Object System.Collections.ArrayList
foreach ($u in $missingInG2) {
    [void]$csvRows.Add([pscustomobject]@{
        GroupeSourceG1 = $Group1
        GroupeCibleG2  = $Group2
        StatutFusion   = "A AJOUTER DANS G2"
        DisplayName    = $u.DisplayName
        EmailAddress   = $u.EmailAddress
        AccountId      = $u.AccountId
        CompteActif    = if ($u.Active) { "Oui" } else { "Non" }
        TypeCompte     = $u.AccountType
    })
}
foreach ($u in $alreadyInG2) {
    [void]$csvRows.Add([pscustomobject]@{
        GroupeSourceG1 = $Group1
        GroupeCibleG2  = $Group2
        StatutFusion   = "DEJA PRESENT DANS G2"
        DisplayName    = $u.DisplayName
        EmailAddress   = $u.EmailAddress
        AccountId      = $u.AccountId
        CompteActif    = if ($u.Active) { "Oui" } else { "Non" }
        TypeCompte     = $u.AccountType
    })
}

$utf8Bom = New-Object System.Text.UTF8Encoding($true)
$writer  = New-Object System.IO.StreamWriter($csvPath, $false, $utf8Bom)
try {
    $writer.WriteLine("GroupeSourceG1;GroupeCibleG2;StatutFusion;DisplayName;EmailAddress;AccountId;CompteActif;TypeCompte")
    foreach ($r in ($csvRows | Sort-Object StatutFusion, DisplayName)) {
        $line = ('"{0}";"{1}";"{2}";"{3}";"{4}";"{5}";"{6}";"{7}"' -f `
            $r.GroupeSourceG1, $r.GroupeCibleG2, $r.StatutFusion, $r.DisplayName, $r.EmailAddress, $r.AccountId, $r.CompteActif, $r.TypeCompte)
        $writer.WriteLine($line)
    }
} finally {
    $writer.Close()
    $writer.Dispose()
}

Write-Host ""
Write-Info ("Rapport CSV genere avec succes : " + $csvPath)
Write-Host ""