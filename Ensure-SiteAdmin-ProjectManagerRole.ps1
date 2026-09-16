<# ---------------------------------------------------------------------------
Ensure-site-admins-Only-ProjectManager.ps1

But:
- Sur tous les projets listés par le CSV de matrice (ProjectKey ou Category),
  garantir que le groupe "site-admins" :
  - est présent dans "1. Gestionnaire Projet"
  - est absent de TOUS les autres rôles du projet

Modes:
- DRY-RUN (simulation) / EXECUTE (appliquer) via UI si -Apply non fourni
- Log dans .\logs\Ensure-site-admins-Only-ProjectManager_yyyyMMdd_HHmmss.log
--------------------------------------------------------------------------- #>

[CmdletBinding()]
param(
  [string] $SiteUrl   = "https://jiradot.atlassian.net",
  [string] $UserEmail = "frederic.guedj@harmonie-mutuelle.fr",

  [string] $CsvPath = "C:\Temp\DSIM_Groups_x_ProjectCategories.csv",
  [char]   $Delimiter = ';',

  [string] $CredentialFile = $null,

  # Libellé exact du groupe admin
  [string] $AdminGroupName = "site-admins",

  # Si tu passes -Apply, on saute la boite de dialogue et on exécute
  [switch] $Apply,

  # reset DPAPI credential
  [switch] $ResetCredential
)

# Neutraliser StrictMode (au cas où le profil l'active)
try { Set-StrictMode -Off } catch {}

# ------------------------- Déterminer le répertoire du script -------------------------
$ScriptDir = $null
if ($PSCommandPath) {
  $ScriptDir = Split-Path -Parent $PSCommandPath
} elseif ($MyInvocation -and $MyInvocation.MyCommand -and $MyInvocation.MyCommand.Path) {
  $ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
}
if ([string]::IsNullOrWhiteSpace($ScriptDir)) { $ScriptDir = (Get-Location).Path }
if ([string]::IsNullOrWhiteSpace($CredentialFile)) {
  $CredentialFile = Join-Path $ScriptDir "jira-credential.xml"
}

# ------------------------- Logging fichier -------------------------
$global:LogFile = $null

function Initialize-LogFile {
  param([string]$BaseDir)
  $logsDir = Join-Path $BaseDir "logs"
  if (-not (Test-Path $logsDir)) { New-Item -ItemType Directory -Path $logsDir -Force | Out-Null }
  $name = "Ensure-site-admins-Only-ProjectManager_{0}.log" -f (Get-Date -Format "yyyyMMdd_HHmmss")
  $path = Join-Path $logsDir $name
  Set-Content -Path $path -Value "" -Encoding UTF8
  return $path
}

function Write-Log {
  param(
    [Parameter(Mandatory=$true)][string]$Message,
    [ValidateSet('INFO','WARN','ERROR','OK','DRYRUN')][string]$Level = 'INFO'
  )
  $ts = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
  $line = "[$ts][$Level] $Message"
  Write-Host $line
  if ($global:LogFile) { Add-Content -Path $global:LogFile -Value $line -Encoding UTF8 }
}

$global:LogFile = Initialize-LogFile -BaseDir $ScriptDir
Write-Log "LogFile: $global:LogFile" "INFO"

# ------------------------- Pré-requis réseau -------------------------
try { [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12 } catch {}
try {
  $proxy = [System.Net.WebRequest]::GetSystemWebProxy()
  $proxy.Credentials = [System.Net.CredentialCache]::DefaultCredentials
  [System.Net.WebRequest]::DefaultWebProxy = $proxy
} catch {}

$SiteUrl = $SiteUrl.TrimEnd('/')

# ------------------------- Helpers safe properties -------------------------
function Get-SafeProp {
  param(
    [Parameter(Mandatory=$true)]$Object,
    [Parameter(Mandatory=$true)][string]$PropName
  )
  if ($null -eq $Object) { return $null }
  $p = $Object.PSObject.Properties[$PropName]
  if ($p) { return $p.Value }
  return $null
}

# ------------------------- Normalisations -------------------------
function Normalize-RoleName {
  param([AllowNull()][string]$s)
  if ([string]::IsNullOrWhiteSpace($s)) { return "" }
  $t = $s.Trim()
  $t = $t -replace '^\s*(\d+)\s*\.\s*', '$1 '       # "1." -> "1 "
  $t = $t -replace '[\.\-_/]+', ' '
  $t = $t -replace '\s+', ' '
  $t = $t.Trim().ToLowerInvariant()
  return $t
}

function Normalize-HeaderName {
  param([AllowNull()][string]$s)
  if ([string]::IsNullOrWhiteSpace($s)) { return "" }
  $t = $s.Trim().ToLowerInvariant()
  $t = $t.Normalize([Text.NormalizationForm]::FormD)
  $sb = New-Object System.Text.StringBuilder
  foreach ($ch in $t.ToCharArray()) {
    if ([Globalization.CharUnicodeInfo]::GetUnicodeCategory($ch) -ne [Globalization.UnicodeCategory]::NonSpacingMark) {
      [void]$sb.Append($ch)
    }
  }
  $t = $sb.ToString()
  $t = $t -replace '[^a-z0-9]+', ''
  return $t
}

# ------------------------- Credentials (DPAPI XML) -------------------------
function Get-JiraCredential {
  param(
    [Parameter(Mandatory=$true)][string]$CredentialFile,
    [Parameter(Mandatory=$true)][string]$UserEmail,
    [switch]$Reset
  )

  if ($Reset -and (Test-Path $CredentialFile)) {
    Remove-Item -Path $CredentialFile -Force -ErrorAction SilentlyContinue
  }

  if (Test-Path $CredentialFile) {
    $cred = Import-Clixml -Path $CredentialFile
    if ($cred -isnot [pscredential]) { throw "Fichier credential invalide: $CredentialFile" }
    return $cred
  }

  Write-Log "Credential introuvable : saisie unique du token (stocké DPAPI dans '$CredentialFile')." "WARN"
  $secureToken = Read-Host -Prompt "Jira API token pour $UserEmail" -AsSecureString
  $cred = New-Object System.Management.Automation.PSCredential ($UserEmail, $secureToken)

  $dir = Split-Path -Parent $CredentialFile
  if (-not (Test-Path $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }

  $cred | Export-Clixml -Path $CredentialFile
  Write-Log "Credential enregistré: $CredentialFile" "OK"
  return $cred
}

function Get-JiraAuthHeader {
  param([Parameter(Mandatory=$true)][pscredential]$Credential)
  $tokenPlain = [System.Net.NetworkCredential]::new("", $Credential.Password).Password
  $pair = "{0}:{1}" -f $Credential.UserName, $tokenPlain
  $b64  = [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes($pair))
  return @{
    Authorization = "Basic $b64"
    Accept        = "application/json"
  }
}

# ------------------------- Invoke Jira (JSON + UTF-8) -------------------------
function Invoke-Jira {
  param(
    [Parameter(Mandatory=$true)][ValidateSet('GET','POST','PUT','DELETE')][string]$Method,
    [Parameter(Mandatory=$true)][string]$Url,
    [hashtable]$Headers,
    $Body
  )

  $json = $null
  try {
    if ($null -ne $Body) {
      $json  = $Body | ConvertTo-Json -Depth 20 -Compress
      $bytes = [System.Text.Encoding]::UTF8.GetBytes($json)
      return Invoke-RestMethod `
        -Method $Method `
        -Uri $Url `
        -Headers $Headers `
        -ContentType "application/json; charset=utf-8" `
        -Body $bytes `
        -ErrorAction Stop
    } else {
      return Invoke-RestMethod -Method $Method -Uri $Url -Headers $Headers -ErrorAction Stop
    }
  } catch {
    $details = $_.Exception.Message
    if ($_.ErrorDetails -and $_.ErrorDetails.Message) {
      $details = "$details | $($_.ErrorDetails.Message)"
    }
    if ($null -ne $json) {
      try { Write-Log "Body JSON envoyé: $json" "ERROR" } catch {}
    }
    throw "Erreur API Jira ($Method $Url) : $details"
  }
}

# ------------------------- UI: sélection CSV si besoin -------------------------
function Resolve-CsvPath {
  param([string]$Path, [string]$InitialDirectory)

  if (Test-Path $Path) { return $Path }

  Write-Log "CSV introuvable: $Path" "WARN"
  Add-Type -AssemblyName System.Windows.Forms | Out-Null
  $dlg = New-Object System.Windows.Forms.OpenFileDialog
  $dlg.Filter = "CSV (*.csv)|*.csv|All files (*.*)|*.*"
  $dlg.Title  = "Sélectionner le CSV Matrice"
  if (Test-Path $InitialDirectory) { $dlg.InitialDirectory = $InitialDirectory }
  if ($dlg.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) { throw "Aucun fichier CSV sélectionné." }
  return $dlg.FileName
}

# ------------------------- UI: choix DRYRUN/EXECUTE -------------------------
function Select-RunMode {
  param([bool]$ApplyAlreadyProvided)

  if ($ApplyAlreadyProvided) {
    return [pscustomobject]@{ Apply = $true }
  }

  Add-Type -AssemblyName System.Windows.Forms | Out-Null
  Add-Type -AssemblyName System.Drawing | Out-Null

  $form = New-Object System.Windows.Forms.Form
  $form.Text = "site-admins - Rôle unique"
  $form.Size = New-Object System.Drawing.Size(620, 210)
  $form.StartPosition = "CenterScreen"
  $form.TopMost = $true

  $lbl = New-Object System.Windows.Forms.Label
  $lbl.AutoSize = $true
  $lbl.Location = New-Object System.Drawing.Point(15, 15)
  $lbl.Text = "Garantit que '$AdminGroupName' n'a QUE '1. Gestionnaire Projet' (retire de tous les autres rôles)."
  $form.Controls.Add($lbl)

  $rbDry = New-Object System.Windows.Forms.RadioButton
  $rbDry.Text = "DRY-RUN (simulation)"
  $rbDry.Location = New-Object System.Drawing.Point(18, 55)
  $rbDry.Checked = $true
  $rbDry.AutoSize = $true
  $form.Controls.Add($rbDry)

  $rbExec = New-Object System.Windows.Forms.RadioButton
  $rbExec.Text = "EXECUTE (applique dans Jira)"
  $rbExec.Location = New-Object System.Drawing.Point(18, 80)
  $rbExec.AutoSize = $true
  $form.Controls.Add($rbExec)

  $btnOk = New-Object System.Windows.Forms.Button
  $btnOk.Text = "OK"
  $btnOk.Location = New-Object System.Drawing.Point(410, 120)
  $btnOk.Add_Click({ $form.Tag = "OK"; $form.Close() })
  $form.Controls.Add($btnOk)

  $btnCancel = New-Object System.Windows.Forms.Button
  $btnCancel.Text = "Annuler"
  $btnCancel.Location = New-Object System.Drawing.Point(490, 120)
  $btnCancel.Add_Click({ $form.Tag = "CANCEL"; $form.Close() })
  $form.Controls.Add($btnCancel)

  $form.ShowDialog() | Out-Null
  if ($form.Tag -ne "OK") { throw "Exécution annulée par l'utilisateur." }

  return [pscustomobject]@{ Apply = $rbExec.Checked }
}

# ------------------------- Jira: projets + catégories -------------------------
function Get-AllProjectsWithCategory {
  param([string]$SiteUrl, [hashtable]$Headers)

  $all = New-Object System.Collections.Generic.List[object]
  $startAt = 0
  $max = 50

  while ($true) {
    $url = "$SiteUrl/rest/api/3/project/search?startAt=$startAt&maxResults=$max&expand=projectCategory"
    $resp = Invoke-Jira -Method GET -Url $url -Headers $Headers
    foreach ($p in @($resp.values)) {
      $pc = Get-SafeProp -Object $p -PropName "projectCategory"
      $all.Add([pscustomobject]@{
        key          = (Get-SafeProp -Object $p -PropName "key")
        name         = (Get-SafeProp -Object $p -PropName "name")
        categoryId   = (Get-SafeProp -Object $pc -PropName "id")
        categoryName = (Get-SafeProp -Object $pc -PropName "name")
      })
    }
    $startAt += $resp.maxResults
    if ($startAt -ge $resp.total) { break }
  }
  return $all
}

# ------------------------- Helpers rôle: groupes d'un rôle + suppression groupe -------------------------
function Get-GroupActorsFromRoleDetail {
  param([object]$RoleDetail)

  $groups = @{}
  foreach ($a in @($RoleDetail.actors)) {
    $actorGroup = Get-SafeProp -Object $a -PropName "actorGroup"
    if ($actorGroup) {
      $name = Get-SafeProp -Object $actorGroup -PropName "name"
      if ($name) { $groups[[string]$name] = $true }
      continue
    }

    $type = Get-SafeProp -Object $a -PropName "type"
    if ($type -and ($type -like "*group*")) {
      $name = Get-SafeProp -Object $a -PropName "name"
      if ($name) { $groups[[string]$name] = $true; continue }
      $dn = Get-SafeProp -Object $a -PropName "displayName"
      if ($dn) { $groups[[string]$dn] = $true; continue }
    }
  }
  return $groups
}

function Remove-GroupFromRole {
  param(
    [Parameter(Mandatory=$true)][string]$RoleUrl,
    [Parameter(Mandatory=$true)][string]$GroupName,
    [Parameter(Mandatory=$true)][hashtable]$Headers
  )

  $enc = [Uri]::EscapeDataString($GroupName)
  $sep = if ($RoleUrl.Contains('?')) { '&' } else { '?' }
  $url = "$RoleUrl$sep" + "group=$enc"

  try {
    Invoke-Jira -Method DELETE -Url $url -Headers $Headers | Out-Null
    return [pscustomobject]@{ Success = $true; Url = $url; Error = $null }
  } catch {
    return [pscustomobject]@{ Success = $false; Url = $url; Error = $_.Exception.Message }
  }
}

# ------------------------- Constantes rôle cible -------------------------
$RoleGestionnaire = "1. Gestionnaire Projet"
$normGestionnaire = Normalize-RoleName $RoleGestionnaire

# ------------------------- Main -------------------------
Write-Log "ScriptDir: $ScriptDir" "INFO"
Write-Log "SiteUrl: $SiteUrl" "INFO"
Write-Log "CredentialFile: $CredentialFile" "INFO"
Write-Log "AdminGroupName: $AdminGroupName" "INFO"

$choice = Select-RunMode -ApplyAlreadyProvided:$Apply.IsPresent
$ApplyMode = [bool]$choice.Apply

$modeText  = if ($ApplyMode) { "EXECUTE (écriture)" } else { "DRY-RUN (simulation)" }
$modeLevel = if ($ApplyMode) { "WARN" } else { "DRYRUN" }
Write-Log "Mode: $modeText" $modeLevel

$CsvPath = Resolve-CsvPath -Path $CsvPath -InitialDirectory $ScriptDir
Write-Log "CSV: $CsvPath" "INFO"

$JiraCredential = Get-JiraCredential -CredentialFile $CredentialFile -UserEmail $UserEmail -Reset:$ResetCredential
$Headers        = Get-JiraAuthHeader -Credential $JiraCredential

$me = Invoke-Jira -Method GET -Url "$SiteUrl/rest/api/3/myself" -Headers $Headers
Write-Log ("Authentifié: {0}" -f (Get-SafeProp -Object $me -PropName "displayName")) "OK"

# ---- Lire CSV + déterminer format ----
$rows = Import-Csv -Path $CsvPath -Delimiter $Delimiter
if (-not $rows -or $rows.Count -eq 0) { throw "CSV vide: $CsvPath" }

$colMap = @{}
$rows[0].PSObject.Properties.Name | ForEach-Object { $colMap[(Normalize-HeaderName $_)] = $_ }

function Get-Col {
  param([string[]]$Candidates)
  foreach ($c in $Candidates) {
    $k = Normalize-HeaderName $c
    if ($colMap.ContainsKey($k)) { return $colMap[$k] }
  }
  return $null
}

$projectKeyCol   = Get-Col @("ProjectKey")
$categoryNameCol = Get-Col @("Project Category", "ProjectCategory")
$categoryIdCol   = Get-Col @("Project Category Id", "ProjectCategoryId")

$format = $null
if ($projectKeyCol) {
  $format = "PROJECTKEY"
} elseif ($categoryIdCol -or $categoryNameCol) {
  $format = "CATEGORY"
} else {
  throw ("CSV invalide: impossible de trouver ProjectKey ou Project Category. Colonnes: {0}" -f ($rows[0].PSObject.Properties.Name -join ", "))
}
Write-Log "CSV format détecté: $format" "INFO"

# ---- Construire liste projets uniques ----
$projectKeys = New-Object System.Collections.Generic.HashSet[string]

if ($format -eq "PROJECTKEY") {
  foreach ($r in $rows) {
    $pk = ([string]$r.($projectKeyCol)).Trim()
    if ($pk) { [void]$projectKeys.Add($pk) }
  }
}

if ($format -eq "CATEGORY") {
  Write-Log "Chargement de tous les projets + catégories..." "INFO"
  $projects = Get-AllProjectsWithCategory -SiteUrl $SiteUrl -Headers $Headers
  Write-Log ("Projets récupérés: {0}" -f $projects.Count) "OK"

  foreach ($r in $rows) {
    $catId = if ($categoryIdCol) { ([string]$r.($categoryIdCol)).Trim() } else { "" }
    $catNm = if ($categoryNameCol) { ([string]$r.($categoryNameCol)).Trim() } else { "" }

    $targets = @()
    if ($catId) {
      $targets = @($projects | Where-Object { $_.categoryId -eq $catId })
    } else {
      $targets = @($projects | Where-Object {
        $_.categoryName -and ($_.categoryName.Trim().ToLowerInvariant() -eq $catNm.Trim().ToLowerInvariant())
      })
    }

    foreach ($p in $targets) {
      if ($p.key) { [void]$projectKeys.Add([string]$p.key) }
    }
  }
}

$allProjectKeys = @($projectKeys)
Write-Log ("Projets ciblés (uniques): {0}" -f $allProjectKeys.Count) "INFO"

# ---- Traitement par projet ----
foreach ($pk in $allProjectKeys) {
  Write-Log "---- Projet $pk ----" "INFO"

  try {
    $roles = Invoke-Jira -Method GET -Url "$SiteUrl/rest/api/3/project/$pk/role" -Headers $Headers
  } catch {
    Write-Log "Impossible de lire les rôles du projet $pk (skip). Err=$($_.Exception.Message)" "ERROR"
    continue
  }

  # index: normRole -> {Name, Url}
  $roleIndex = @{}
  foreach ($prop in $roles.PSObject.Properties) {
    $roleIndex[(Normalize-RoleName $prop.Name)] = [pscustomobject]@{ Name = $prop.Name; Url = $prop.Value }
  }

  if (-not $roleIndex.ContainsKey($normGestionnaire)) {
    Write-Log "Rôle '$RoleGestionnaire' introuvable sur $pk (skip)" "ERROR"
    continue
  }

  # 1) Assurer présence dans le rôle gestionnaire
  $urlGestionnaire = $roleIndex[$normGestionnaire].Url
  try {
    $detailGest = Invoke-Jira -Method GET -Url $urlGestionnaire -Headers $Headers
    $groupsGest = Get-GroupActorsFromRoleDetail -RoleDetail $detailGest
  } catch {
    Write-Log "Impossible de lire détail '$RoleGestionnaire' ($pk) (skip). Err=$($_.Exception.Message)" "ERROR"
    continue
  }

  if ($groupsGest.ContainsKey($AdminGroupName)) {
    Write-Log "OK déjà présent: '$AdminGroupName' dans '$RoleGestionnaire' ($pk)" "OK"
  } else {
    if (-not $ApplyMode) {
      Write-Log "Ajouterait: '$AdminGroupName' -> '$RoleGestionnaire' ($pk)" "DRYRUN"
    } else {
      try {
        $body = @{ group = @($AdminGroupName) }
        Invoke-Jira -Method POST -Url $urlGestionnaire -Headers $Headers -Body $body | Out-Null
        Write-Log "Ajouté: '$AdminGroupName' -> '$RoleGestionnaire' ($pk)" "OK"
      } catch {
        Write-Log "POST KO (on continue): '$AdminGroupName' -> '$RoleGestionnaire' ($pk) | Err=$($_.Exception.Message)" "ERROR"
      }
    }
  }

  # 2) Retirer de TOUS les autres rôles
  foreach ($normRole in $roleIndex.Keys) {
    if ($normRole -eq $normGestionnaire) { continue }

    $otherName = $roleIndex[$normRole].Name
    $otherUrl  = $roleIndex[$normRole].Url

    # Lire détail rôle pour savoir si le groupe est présent
    try {
      $detail = Invoke-Jira -Method GET -Url $otherUrl -Headers $Headers
      $groups = Get-GroupActorsFromRoleDetail -RoleDetail $detail
    } catch {
      Write-Log "Impossible de lire détail '$otherName' ($pk) (skip). Err=$($_.Exception.Message)" "ERROR"
      continue
    }

    if (-not $groups.ContainsKey($AdminGroupName)) { continue }

    if (-not $ApplyMode) {
      Write-Log "Retirerait: '$AdminGroupName' de '$otherName' ($pk)" "DRYRUN"
    } else {
      $res = Remove-GroupFromRole -RoleUrl $otherUrl -GroupName $AdminGroupName -Headers $Headers
      if ($res.Success) {
        Write-Log "Retiré: '$AdminGroupName' de '$otherName' ($pk)" "OK"
      } else {
        Write-Log "DELETE KO (on continue): '$AdminGroupName' de '$otherName' ($pk) | Url=$($res.Url) | Err=$($res.Error)" "ERROR"
      }
    }
  }
}

Write-Log "Terminé." "OK"
Write-Log "Log sauvegardé: $global:LogFile" "OK"