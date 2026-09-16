<#
.SYNOPSIS
    Audit des Espaces Confluence associés à un groupe Jira/Confluence (G1) dans l'optique de son décommissionnement.
.DESCRIPTION
    - Parcourt tous les espaces Confluence de l'instance.
    - Identifie les espaces où le groupe G1 possède des habilitations (Lecture, Écriture, Admin, etc.).
    - Pour chaque espace :
        * Liste les autres groupes habilités en lecture/accès.
        * Calcule le niveau de risque lors de la suppression de G1 :
          -> "⚠️ CRITIQUE : Seul G1 donne l'accès" (aucun autre groupe métier n'a accès).
          -> "⚠️ ATTENTION : Uniquement admins en plus de G1".
          -> "✅ COUVERT : D'autres groupes métiers garantissent l'accès".
    - Exporte le résultat en CSV (séparateur ;) encodé en UTF-8 avec BOM (Excel ready).
#>

[CmdletBinding()]
param(
    [string]$GroupName,
    [string]$ProxyUrl = "http://prc37cti1.hm.dm.ad:8080"
)

# ======================================================================
# 1. INITIALISATION PROXY, TLS & ENCODAGE
# ======================================================================
[System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls12 -bor [System.Net.SecurityProtocolType]::Tls13
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

if (-not $ProxyUrl) {
    try {
        $sysProxy = [System.Net.WebRequest]::GetSystemWebProxy()
        $testUri = New-Object System.Uri("https://api.atlassian.com")
        $pUri = $sysProxy.GetProxy($testUri)
        if ($pUri -and $pUri.AbsoluteUri -ne $testUri.AbsoluteUri) {
            $ProxyUrl = $pUri.AbsoluteUri
        }
    } catch {}
}

function Fix-DoubleUtf8([string]$str) {
    if ([string]::IsNullOrWhiteSpace($str)) { return "" }
    try {
        if ($str -match '[\xC2-\xC3][\x80-\xBF]') {
            $bytes = [System.Text.Encoding]::GetEncoding("ISO-8859-1").GetBytes($str)
            return [System.Text.Encoding]::UTF8.GetString($bytes)
        }
    } catch {}
    return $str
}

function Invoke-SafeGet {
    param([string]$Url, [hashtable]$Headers)
    $params = @{
        Uri     = $Url
        Method  = "Get"
        Headers = $Headers
    }
    if ($ProxyUrl) {
        $params["Proxy"] = $ProxyUrl
        $params["ProxyUseDefaultCredentials"] = $true
    }
    return (Invoke-RestMethod @params)
}

# ======================================================================
# 2. DOSSIERS & LOGS
# ======================================================================
$scriptDir  = if ($PSCommandPath) { Split-Path -Parent $PSCommandPath } else { (Get-Location).Path }
$secretsDir = Join-Path $scriptDir "secrets"
$exportsDir = Join-Path $scriptDir "exports"

if (-not (Test-Path $exportsDir)) { [void](New-Item -ItemType Directory -Path $exportsDir -Force) }

function Write-Info($msg) { Write-Host "[$((Get-Date).ToString('yyyy-MM-dd HH:mm:ss'))] [INFO] $msg" -ForegroundColor Cyan }
function Write-Warn($msg) { Write-Host "[$((Get-Date).ToString('yyyy-MM-dd HH:mm:ss'))] [WARN] $msg" -ForegroundColor Yellow }
function Write-ErrLog($msg) { Write-Host "[$((Get-Date).ToString('yyyy-MM-dd HH:mm:ss'))] [ERROR] $msg" -ForegroundColor Red }

# ======================================================================
# 3. CREDENTIALS ATLASSIAN
# ======================================================================
$jiraCredFile = Join-Path $secretsDir "jira-jiradot.cred.xml"

if (-not (Test-Path $jiraCredFile)) {
    throw "Fichier Jira creds introuvable: $jiraCredFile"
}

$jiraData    = Import-Clixml -Path $jiraCredFile
$baseUrl     = ($jiraData.JiraBaseUrl).TrimEnd("/")
$jiraCred    = $jiraData.Credential
$jiraEmail   = $jiraCred.UserName
$jiraToken   = $jiraCred.GetNetworkCredential().Password

$pair        = [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes("${jiraEmail}:${jiraToken}"))
$authHeaders = @{
    "Authorization" = "Basic $pair"
    "Accept"        = "application/json"
}

# ======================================================================
# 4. SÉLECTION DU GROUPE CIBLE (G1)
# ======================================================================
if ([string]::IsNullOrWhiteSpace($GroupName)) {
    [void][System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
    [void][System.Reflection.Assembly]::LoadWithPartialName("System.Drawing")
    
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Audit Espaces Confluence par Groupe"
    $form.Size = New-Object System.Drawing.Size(460, 180)
    $form.StartPosition = "CenterScreen"
    $form.FormBorderStyle = "FixedDialog"
    $form.MaximizeBox = $false; $form.MinimizeBox = $false; $form.TopMost = $true

    $lbl = New-Object System.Windows.Forms.Label
    $lbl.Text = "Entrez le nom exact du groupe Jira/Confluence (G1) :"
    $lbl.Location = New-Object System.Drawing.Point(20, 20)
    $lbl.Size = New-Object System.Drawing.Size(400, 20)
    [void]$form.Controls.Add($lbl)

    $txt = New-Object System.Windows.Forms.TextBox
    $txt.Location = New-Object System.Drawing.Point(20, 45)
    $txt.Size = New-Object System.Drawing.Size(400, 25)
    [void]$form.Controls.Add($txt)

    $btnOK = New-Object System.Windows.Forms.Button
    $btnOK.Text = "Valider"
    $btnOK.Location = New-Object System.Drawing.Point(140, 85)
    $btnOK.Size = New-Object System.Drawing.Size(80, 30)
    $btnOK.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $form.AcceptButton = $btnOK
    [void]$form.Controls.Add($btnOK)

    $btnCancel = New-Object System.Windows.Forms.Button
    $btnCancel.Text = "Annuler"
    $btnCancel.Location = New-Object System.Drawing.Point(230, 85)
    $btnCancel.Size = New-Object System.Drawing.Size(80, 30)
    $btnCancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $form.CancelButton = $btnCancel
    [void]$form.Controls.Add($btnCancel)

    $txt.Focus()
    $res = $form.ShowDialog()
    if ($res -eq [System.Windows.Forms.DialogResult]::OK -and -not [string]::IsNullOrWhiteSpace($txt.Text)) {
        $GroupName = $txt.Text.Trim()
    } else {
        $GroupName = Read-Host "Nom du groupe cible"
    }
    $form.Dispose()
}

if ([string]::IsNullOrWhiteSpace($GroupName)) {
    Write-ErrLog "Aucun groupe spécifié. Fin du script."
    exit 1
}

$GroupName = Fix-DoubleUtf8 $GroupName
Write-Info "Groupe Confluence cible : $GroupName"
if ($ProxyUrl) { Write-Info "Proxy actif : $ProxyUrl" }

# Groupes d'administration système à exclure lors de l'évaluation de l'accès exclusif
$adminGroupPatterns = @(
    "site-admins", "org-admins", "confluence-administrators", 
    "administrators", "jira-administrators", "system-administrators",
    "atlassian-addons-admin", "audit-confluence"
)

# ======================================================================
# 5. EXTRACTION & ANALYSE DES PERMISSIONS PAR ESPACE CONFLUENCE
# ======================================================================
Write-Info "Confluence: parcours et analyse de l'ensemble des espaces..."

$spacesUrl = "$baseUrl/wiki/rest/api/space?limit=50&status=current&type=global&expand=permissions"
$confluenceSpaces = New-Object System.Collections.ArrayList
$start = 0
$limit = 50

while ($true) {
    $urlWithPaging = "$baseUrl/wiki/rest/api/space?limit=$limit&start=$start&status=current&type=global&expand=permissions"
    try {
        $resp = Invoke-SafeGet -Url $urlWithPaging -Headers $authHeaders
    } catch {
        Write-ErrLog "Erreur chargement espaces Confluence (start=$start) : $($_.Exception.Message)"
        break
    }

    if (-not $resp -or -not $resp.results -or $resp.results.Count -eq 0) { break }

    foreach ($sp in $resp.results) {
        [void]$confluenceSpaces.Add($sp)
    }

    if ($resp.size -lt $limit) { break }
    $start += $resp.size
    Write-Info "Confluence: $($confluenceSpaces.Count) espaces analysés..."
}

Write-Info "Confluence: Total de $($confluenceSpaces.Count) espaces récupérés. Filtrage sur le groupe '$GroupName'..."

# ======================================================================
# 6. ÉVALUATION DES RISQUES POUR LE GROUPE G1
# ======================================================================
$auditResults = New-Object System.Collections.ArrayList
$exclusiveCount = 0

foreach ($sp in $confluenceSpaces) {
    $spaceKey  = [string]$sp.key
    $spaceName = Fix-DoubleUtf8 ([string]$sp.name)
    $spaceUrl  = "$baseUrl/wiki/spaces/$spaceKey"

    $g1HasAccess = $false
    $g1Ops = New-Object System.Collections.Generic.HashSet[string]
    $allPermittedGroups = New-Object System.Collections.Generic.HashSet[string]
    $otherBusinessGroups = New-Object System.Collections.Generic.HashSet[string]

    if ($sp.permissions) {
        foreach ($perm in $sp.permissions) {
            $op = [string]$perm.operation

            # Extraction des groupes associés à cette permission
            $groupsInPerm = @()
            if ($perm.subjects -and $perm.subjects.group -and $perm.subjects.group.results) {
                $groupsInPerm = $perm.subjects.group.results
            } elseif ($perm.subject -and $perm.subject.type -eq "group") {
                $groupsInPerm = @($perm.subject)
            }

            foreach ($g in $groupsInPerm) {
                $gName = Fix-DoubleUtf8 ([string]$g.name)
                if ([string]::IsNullOrWhiteSpace($gName)) { continue }

                [void]$allPermittedGroups.Add($gName)

                if ($gName -ieq $GroupName) {
                    $g1HasAccess = $true
                    [void]$g1Ops.Add($op)
                } else {
                    # Vérifier si c'est un groupe métier ou un groupe admin système
                    $isAdmin = $false
                    foreach ($pattern in $adminGroupPatterns) {
                        if ($gName -ilike "*$pattern*") { $isAdmin = $true; break }
                    }
                    if (-not $isAdmin) {
                        [void]$otherBusinessGroups.Add($gName)
                    }
                }
            }
        }
    }

    # Si le groupe G1 a des droits sur l'espace
    if ($g1HasAccess) {
        $otherGroupsStr = if ($otherBusinessGroups.Count -gt 0) {
            ($otherBusinessGroups | Sort-Object) -join " | "
        } else {
            "(aucun autre groupe métier)"
        }

        $allOtherGroups = @($allPermittedGroups | Where-Object { $_ -ine $GroupName })
        $allOtherGroupsStr = if ($allOtherGroups.Count -gt 0) {
            ($allOtherGroups | Sort-Object) -join " | "
        } else {
            "(aucun)"
        }

        # Détermination du point d'attention
        $pointAttention = ""
        $impactNiveau   = ""

        if ($otherBusinessGroups.Count -eq 0 -and $allOtherGroups.Count -eq 0) {
            $pointAttention = "⚠️ CRITIQUE : Seul $GroupName a accès à cet espace (Risque d'espace orphelin si G1 supprimé)"
            $impactNiveau   = "CRITIQUE (Exclusif)"
            $exclusiveCount++
        } elseif ($otherBusinessGroups.Count -eq 0) {
            $pointAttention = "⚠️ ATTENTION : Seul $GroupName + Admins système ont accès (Aucun autre groupe métier)"
            $impactNiveau   = "ÉLEVÉ (Métier Exclusif)"
            $exclusiveCount++
        } else {
            $pointAttention = "✅ COUVERT : Accès partagé avec d'autres groupes métiers"
            $impactNiveau   = "FAIBLE"
        }

        $g1OpsStr = ($g1Ops | Sort-Object) -join ", "

        [void]$auditResults.Add([pscustomobject]@{
            "Groupe Cible (G1)"                  = $GroupName
            "Clé Espace"                         = $spaceKey
            "Nom Espace"                         = $spaceName
            "Point d'attention / Risque"         = $pointAttention
            "Niveau d'impact"                    = $impactNiveau
            "Permissions de G1"                  = $g1OpsStr
            "Autres Groupes Métiers Habilités"   = $otherGroupsStr
            "Tous les Autres Groupes Habilités"  = $allOtherGroupsStr
            "URL Espace Confluence"              = $spaceUrl
        })
    }
}

Write-Info "Confluence: $($auditResults.Count) espace(s) trouvé(s) avec des habilitations pour '$GroupName'."
if ($exclusiveCount -gt 0) {
    Write-Warn "Points d'attention détectés : $exclusiveCount espace(s) dépendent exclusivement de '$GroupName' !"
}

# ======================================================================
# 7. EXPORT CSV (AVEC BOM UTF-8 STRICT)
# ======================================================================
$dateStamp = Get-Date -Format "yyyyMMdd_HHmmss"
$safeGroupName = $GroupName -replace '[\\/:*?"<>| ]', '_'
$exportCsv = Join-Path $exportsDir "Audit_Confluence_Espaces_${safeGroupName}_${dateStamp}.csv"

$headers = @(
    "Groupe Cible (G1)", "Clé Espace", "Nom Espace", "Point d'attention / Risque",
    "Niveau d'impact", "Permissions de G1", "Autres Groupes Métiers Habilités",
    "Tous les Autres Groupes Habilités", "URL Espace Confluence"
)

$csvLines = New-Object System.Collections.Generic.List[string]
$csvLines.Add(($headers -join ";"))

foreach ($r in $auditResults) {
    $lineVals = @()
    foreach ($h in $headers) {
        $val = [string]$r.$h
        if ($val -match '[;"\r\n]') {
            $val = '"' + ($val -replace '"', '""') + '"'
        }
        $lineVals += $val
    }
    $csvLines.Add(($lineVals -join ";"))
}

$utf8WithBom = New-Object System.Text.UTF8Encoding($true)
[System.IO.File]::WriteAllLines($exportCsv, $csvLines, $utf8WithBom)

Write-Host ""
Write-Info "=== AUDIT CONFLUENCE TERMINÉ ==="
Write-Info "Fichier généré : $exportCsv"
Write-Host ""

# Affichage synthèse console
$auditResults | Sort-Object "Niveau d'impact" | Select-Object "Clé Espace", "Nom Espace", "Niveau d'impact", "Autres Groupes Métiers Habilités" | Format-Table -AutoSize