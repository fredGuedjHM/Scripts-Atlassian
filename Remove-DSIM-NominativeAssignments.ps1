<# ---------------------------------------------------------------------------
Remove-DSIM-NominativeAssignments.ps1

But:
- Identifier les users membres de groupes dont le nom contient "DSIM" (ou token)
- Pour les projets issus d'une matrice (CSV catégorie/groupe/role ou projectKey/role/group),
  supprimer les affectations NOMINATIVES redondantes en tenant compte d'une hiérarchie :
    1. Gestionnaire Projet  ⊃  2. Utilisateur  ⊃  3. Lecture seule

Règle finale :
- Pour un user membre d'au moins un groupe DSIM de la matrice sur ce projet,
  on calcule son niveau max via ces groupes (power):
    role 1 => 3 ; role 2 => 2 ; role 3 => 1
- On supprime alors ses affectations nominatives dans les rôles 1/2/3 dont
  le power est <= au power max obtenu via ses groupes (redondant ou inférieur).

Modes:
- DRY-RUN (simulation) / EXECUTE (appliquer) via UI si -Apply non fourni
- Log dans .\logs\Remove-DSIM-NominativeAssignments_yyyyMMdd_HHmmss.log
--------------------------------------------------------------------------- #>

[CmdletBinding()]
param(
  [string] $SiteUrl   = "https://jiradot.atlassian.net",
  [string] $UserEmail = "frederic.guedj@harmonie-mutuelle.fr",

  [string] $CsvPath = "C:\Temp\DSIM_Groups_x_ProjectCategories.csv",
  [char]   $Delimiter = ';',

  [string] $CredentialFile = $null,

  # Si tu passes -Apply, on saute la boite de dialogue et on exécute
  [switch] $Apply,

  # reset DPAPI credential
  [switch] $ResetCredential,

  # Filtre groupe DSIM (par défaut: contient "DSIM")
  [string] $DsimGroupToken = "DSIM",

  # Pagination membres groupes
  [int] $MaxResults = 1000
)

# IMPORTANT: neutraliser StrictMode (si activé dans le profil)
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
  $name = "Remove-DSIM-NominativeAssignments_{0}.log" -f (Get-Date -Format "yyyyMMdd_HHmmss")
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

function Get-RolePowerFromRoleName {
  param([string]$RoleName)

  # Retourne 3 (role1), 2 (role2), 1 (role3), 0 sinon
  if ([string]::IsNullOrWhiteSpace($RoleName)) { return 0 }

  $m = [regex]::Match($RoleName.Trim(), '^\s*(\d)')
  if (-not $m.Success) { return 0 }

  switch ([int]$m.Groups[1].Value) {
    1 { return 3 }
    2 { return 2 }
    3 { return 1 }
    default { return 0 }
  }
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
    try {
      $cred = Import-Clixml -Path $CredentialFile
      if ($cred -isnot [pscredential]) { throw "Fichier credential invalide." }
      return $cred
    } catch {
      throw "Impossible de lire '$CredentialFile' (DPAPI / fichier corrompu / autre compte Windows). Détail: $($_.Exception.Message)"
    }
  }

  Write-Log "Credential introuvable : saisie unique du token (stocké chiffré DPAPI dans '$CredentialFile')." "WARN"
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
  $b64  = [Convert]::ToBase64String([System.Text.Encoding]::ASCII.GetBytes($pair))
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
  Write-Log "Ouverture d'une fenêtre de sélection de fichier CSV..." "INFO"

  Add-Type -AssemblyName System.Windows.Forms | Out-Null
  $dlg = New-Object System.Windows.Forms.OpenFileDialog
  $dlg.Filter = "CSV (*.csv)|*.csv|All files (*.*)|*.*"
  $dlg.Title  = "Sélectionner le CSV Matrice DSIM"
  if (Test-Path $InitialDirectory) { $dlg.InitialDirectory = $InitialDirectory }

  if ($dlg.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) {
    throw "Aucun fichier CSV sélectionné."
  }
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
  $form.Text = "DSIM - Suppression nominatives (hiérarchie rôles)"
  $form.Size = New-Object System.Drawing.Size(640, 230)
  $form.StartPosition = "CenterScreen"
  $form.TopMost = $true

  $lbl = New-Object System.Windows.Forms.Label
  $lbl.AutoSize = $true
  $lbl.Location = New-Object System.Drawing.Point(15, 15)
  $lbl.Text = "Supprime les nominatives redondantes pour users DSIM (1⊃2⊃3, inférieur OU égal au rôle de groupe)."
  $form.Controls.Add($lbl)

  $rbDry = New-Object System.Windows.Forms.RadioButton
  $rbDry.Text = "DRY-RUN (simulation)"
  $rbDry.Location = New-Object System.Drawing.Point(18, 55)
  $rbDry.Checked = $true
  $rbDry.AutoSize = $true
  $form.Controls.Add($rbDry)

  $rbExec = New-Object System.Windows.Forms.RadioButton
  $rbExec.Text = "EXECUTE (supprime réellement)"
  $rbExec.Location = New-Object System.Drawing.Point(18, 80)
  $rbExec.AutoSize = $true
  $form.Controls.Add($rbExec)

  $btnOk = New-Object System.Windows.Forms.Button
  $btnOk.Text = "OK"
  $btnOk.Location = New-Object System.Drawing.Point(420, 130)
  $btnOk.Add_Click({ $form.Tag = "OK"; $form.Close() })
  $form.Controls.Add($btnOk)

  $btnCancel = New-Object System.Windows.Forms.Button
  $btnCancel.Text = "Annuler"
  $btnCancel.Location = New-Object System.Drawing.Point(500, 130)
  $btnCancel.Add_Click({ $form.Tag = "CANCEL"; $form.Close() })
  $form.Controls.Add($btnCancel)

  $form.ShowDialog() | Out-Null
  if ($form.Tag -ne "OK") { throw "Exécution annulée par l'utilisateur." }

  return [pscustomobject]@{ Apply = $rbExec.Checked }
}

# ------------------------- Mapping rôles -------------------------
$RoleNameMap = @{
  "1-gestionnaire"  = "1. Gestionnaire Projet"
  "2-utilisateur"   = "2. Utilisateur"
  "3-lecture seule" = "3. Lecture seule"
  "ne pas toucher"  = $null
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
        key          = [string](Get-SafeProp -Object $p -PropName "key")
        name         = [string](Get-SafeProp -Object $p -PropName "name")
        categoryId   = [string](Get-SafeProp -Object $pc -PropName "id")
        categoryName = [string](Get-SafeProp -Object $pc -PropName "name")
      })
    }

    $startAt += $resp.maxResults
    if ($startAt -ge $resp.total) { break }
  }

  return $all
}

# ------------------------- Jira: Groups picker (DSIM) -------------------------
function Get-GroupsByQuery {
  param([string]$SiteUrl, [hashtable]$Headers, [string]$Query)

  $out = New-Object System.Collections.Generic.List[string]
  $startAt = 0
  $max = 50

  while ($true) {
    $encQ = [Uri]::EscapeDataString($Query)
    $url = "$SiteUrl/rest/api/3/groups/picker?query=$encQ&startAt=$startAt&maxResults=$max"
    $resp = Invoke-Jira -Method GET -Url $url -Headers $Headers

    foreach ($g in @($resp.groups)) {
      $name = [string](Get-SafeProp -Object $g -PropName "name")
      if ($name) { $out.Add($name) }
    }

    $startAt += $max
    if (-not $resp.groups -or $resp.groups.Count -lt $max) { break }
  }

  return $out | Sort-Object -Unique
}

function Get-GroupMembers {
  param(
    [string]$SiteUrl,
    [hashtable]$Headers,
    [string]$GroupName,
    [int]$MaxResults = 1000
  )

  $members = New-Object System.Collections.Generic.List[object]
  $startAt = 0

  while ($true) {
    $enc = [Uri]::EscapeDataString($GroupName)
    $url = "$SiteUrl/rest/api/3/group/member?groupname=$enc&startAt=$startAt&maxResults=$MaxResults"
    $resp = Invoke-Jira -Method GET -Url $url -Headers $Headers

    foreach ($u in @($resp.values)) {
      $aid = [string](Get-SafeProp -Object $u -PropName "accountId")
      if ($aid) { $members.Add($u) }
    }

    $startAt += $resp.maxResults
    if ($startAt -ge $resp.total) { break }
  }

  return $members
}

# ------------------------- Helpers roles: users nominativement dans un rôle -------------------------
function Get-RoleActorsUsers {
  param([object]$RoleDetail)

  $users = New-Object System.Collections.Generic.List[object]
  foreach ($a in @($RoleDetail.actors)) {
    $actorUser = Get-SafeProp -Object $a -PropName "actorUser"
    if ($actorUser) {
      $aid = [string](Get-SafeProp -Object $actorUser -PropName "accountId")
      if ($aid) {
        $dn = [string](Get-SafeProp -Object $actorUser -PropName "displayName")
        $users.Add([pscustomobject]@{
          accountId   = $aid
          displayName = $dn
        })
      }
    }
  }
  return $users
}

function Remove-UserFromRole {
  <#
    Supprime l'affectation NOMINATIVE d'un user sur un rôle de projet.
    API: DELETE {roleUrl}?user={accountId}
  #>
  param(
    [Parameter(Mandatory=$true)][string]$RoleUrl,
    [Parameter(Mandatory=$true)][string]$AccountId,
    [Parameter(Mandatory=$true)][hashtable]$Headers
  )

  $enc = [Uri]::EscapeDataString($AccountId)
  $sep = if ($RoleUrl.Contains('?')) { '&' } else { '?' }
  $url = "$RoleUrl$sep" + "user=$enc"

  try {
    Invoke-Jira -Method DELETE -Url $url -Headers $Headers | Out-Null
    return [pscustomobject]@{ Success = $true; Url = $url; Error = $null }
  } catch {
    return [pscustomobject]@{ Success = $false; Url = $url; Error = $_.Exception.Message }
  }
}

# ------------------------- MAIN -------------------------
Write-Log "ScriptDir: $ScriptDir" "INFO"
Write-Log "SiteUrl: $SiteUrl" "INFO"
Write-Log "CredentialFile: $CredentialFile" "INFO"

$choice    = Select-RunMode -ApplyAlreadyProvided:$Apply.IsPresent
$ApplyMode = [bool]$choice.Apply

$modeText  = if ($ApplyMode) { "EXECUTE (écriture)" } else { "DRY-RUN (simulation)" }
$modeLevel = if ($ApplyMode) { "WARN" } else { "DRYRUN" }
Write-Log "Mode: $modeText" $modeLevel

$CsvPath = Resolve-CsvPath -Path $CsvPath -InitialDirectory $ScriptDir
Write-Log "CSV: $CsvPath" "INFO"

$JiraCredential = Get-JiraCredential -CredentialFile $CredentialFile -UserEmail $UserEmail -Reset:$ResetCredential
$Headers        = Get-JiraAuthHeader -Credential $JiraCredential

$me = Invoke-Jira -Method GET -Url "$SiteUrl/rest/api/3/myself" -Headers $Headers
Write-Log ("Authentifié: {0}" -f [string](Get-SafeProp -Object $me -PropName "displayName")) "OK"

# ---- 1) Charger matrice CSV -> entries (ProjectKey, Group, RoleNorm, RolePower) ----
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

$projectKeyCol = Get-Col @("ProjectKey")
$roleColA      = Get-Col @("Role")
$groupColA     = Get-Col @("Group")

$categoryNameCol = Get-Col @("Project Category", "ProjectCategory")
$categoryIdCol   = Get-Col @("Project Category Id", "ProjectCategoryId")
$groupColB       = Get-Col @("DSIM Group Name", "Group")
$roleColB        = Get-Col @("Rôle", "Role")

$format = $null
if ($projectKeyCol -and $roleColA -and $groupColA) {
  $format = "PROJECTKEY"
} elseif ( ($categoryIdCol -or $categoryNameCol) -and $groupColB -and $roleColB ) {
  $format = "CATEGORY"
} else {
  throw ("CSV invalide: format non reconnu. Colonnes trouvées: {0}" -f ($rows[0].PSObject.Properties.Name -join ", "))
}
Write-Log "CSV format détecté: $format" "INFO"

$UnifiedEntries = New-Object System.Collections.Generic.List[object]

function Add-UnifiedEntry {
  param([string]$ProjectKey, [string]$Group, [string]$RoleRaw)

  if ([string]::IsNullOrWhiteSpace($ProjectKey) -or [string]::IsNullOrWhiteSpace($Group) -or [string]::IsNullOrWhiteSpace($RoleRaw)) { return }

  $mapped = $RoleRaw
  $k = $RoleRaw.Trim().ToLowerInvariant()
  if ($RoleNameMap.ContainsKey($k)) { $mapped = $RoleNameMap[$k] }
  if ($null -eq $mapped) { return } # "Ne pas toucher"

  $power = Get-RolePowerFromRoleName -RoleName $mapped
  if ($power -le 0) { return }

  $UnifiedEntries.Add([pscustomobject]@{
    ProjectKey = $ProjectKey
    Group      = $Group
    RoleName   = $mapped
    NormRole   = (Normalize-RoleName $mapped)
    RolePower  = $power
  })
}

if ($format -eq "PROJECTKEY") {
  foreach ($r in $rows) {
    $pk = ([string]$r.($projectKeyCol)).Trim()
    $gr = ([string]$r.($groupColA)).Trim()
    $ro = ([string]$r.($roleColA)).Trim()
    Add-UnifiedEntry -ProjectKey $pk -Group $gr -RoleRaw $ro
  }
}

if ($format -eq "CATEGORY") {
  Write-Log "Chargement de tous les projets + catégories..." "INFO"
  $projects = Get-AllProjectsWithCategory -SiteUrl $SiteUrl -Headers $Headers
  Write-Log ("Projets récupérés: {0}" -f $projects.Count) "OK"

  foreach ($r in $rows) {
    $catId = if ($categoryIdCol) { ([string]$r.($categoryIdCol)).Trim() } else { "" }
    $catNm = if ($categoryNameCol) { ([string]$r.($categoryNameCol)).Trim() } else { "" }

    $gr = ([string]$r.($groupColB)).Trim()
    $ro = ([string]$r.($roleColB)).Trim()
    if ([string]::IsNullOrWhiteSpace($gr) -or [string]::IsNullOrWhiteSpace($ro)) { continue }

    $targets = @()
    if (-not [string]::IsNullOrWhiteSpace($catId)) {
      $targets = @($projects | Where-Object { $_.categoryId -eq $catId })
    } else {
      $targets = @($projects | Where-Object {
        $_.categoryName -and ($_.categoryName.Trim().ToLowerInvariant() -eq $catNm.Trim().ToLowerInvariant())
      })
    }

    foreach ($p in $targets) {
      Add-UnifiedEntry -ProjectKey ([string]$p.key) -Group $gr -RoleRaw $ro
    }
  }
}

# Garder uniquement les groupes DSIM
$token = $DsimGroupToken.Trim().ToUpperInvariant()
$UnifiedEntries = @($UnifiedEntries | Where-Object { $_.Group -and ($_.Group.ToUpperInvariant().Contains($token)) })

if (-not $UnifiedEntries -or $UnifiedEntries.Count -eq 0) {
  throw "Aucune entrée DSIM trouvée dans la matrice (après filtre groupe contenant '$DsimGroupToken')."
}

$EntriesByProject = $UnifiedEntries | Group-Object ProjectKey
Write-Log ("Projets à analyser (issus matrice, DSIM only): {0}" -f $EntriesByProject.Count) "INFO"

# ---- 2) Groupes DSIM + membres (users DSIM) ----
Write-Log "Recherche des groupes contenant '$DsimGroupToken'..." "INFO"
$dsimGroups = Get-GroupsByQuery -SiteUrl $SiteUrl -Headers $Headers -Query $DsimGroupToken
Write-Log ("Groupes trouvés via picker: {0}" -f $dsimGroups.Count) "OK"

$matrixGroups = @($UnifiedEntries | Select-Object -ExpandProperty Group -Unique)
$dsimGroupsToUse = @($dsimGroups | Where-Object { $matrixGroups -contains $_ })

Write-Log ("Groupes DSIM utilisés (intersection matrice): {0}" -f $dsimGroupsToUse.Count) "INFO"
if ($dsimGroupsToUse.Count -eq 0) {
  throw "Aucun groupe DSIM de la matrice n'a été retrouvé via groups/picker. Vérifie noms des groupes."
}

$UserToGroups = @{}   # accountId -> HashSet(groupName)
$Users        = @{}   # accountId -> {displayName, email}

foreach ($g in $dsimGroupsToUse) {
  Write-Log "Lecture membres du groupe: $g" "INFO"
  try {
    $members = Get-GroupMembers -SiteUrl $SiteUrl -Headers $Headers -GroupName $g -MaxResults $MaxResults
    Write-Log ("Membres: {0}" -f $members.Count) "OK"

    foreach ($m in $members) {
      $aid = [string](Get-SafeProp -Object $m -PropName "accountId")
      if (-not $aid) { continue }

      if (-not $UserToGroups.ContainsKey($aid)) {
        $UserToGroups[$aid] = New-Object System.Collections.Generic.HashSet[string]
      }
      [void]$UserToGroups[$aid].Add($g)

      if (-not $Users.ContainsKey($aid)) {
        $Users[$aid] = [pscustomobject]@{
          displayName = [string](Get-SafeProp -Object $m -PropName "displayName")
          email       = [string](Get-SafeProp -Object $m -PropName "emailAddress")
        }
      }
    }
  } catch {
    Write-Log "Impossible de lire les membres de '$g' (on continue). Err=$($_.Exception.Message)" "ERROR"
  }
}

Write-Log ("Users DSIM uniques (accountId): {0}" -f $Users.Keys.Count) "INFO"

# ---- 3) Analyse projets / suppression nominatives ----
# Rôles hiérarchiques 1/2/3 à traiter
$Role1 = Normalize-RoleName "1. Gestionnaire Projet"
$Role2 = Normalize-RoleName "2. Utilisateur"
$Role3 = Normalize-RoleName "3. Lecture seule"

$stats = [pscustomobject]@{
  ProjectsAnalyzed = 0
  RoleDetailsRead  = 0
  CandidatesFound  = 0
  WouldDelete      = 0
  Deleted          = 0
  DeleteErrors     = 0
}

foreach ($projGroup in $EntriesByProject) {
  $projectKey = [string]$projGroup.Name
  $entries    = @($projGroup.Group)

  $stats.ProjectsAnalyzed++
  Write-Log "---- Projet $projectKey ----" "INFO"

  # index roles du projet
  try {
    $roles = Invoke-Jira -Method GET -Url "$SiteUrl/rest/api/3/project/$projectKey/role" -Headers $Headers
  } catch {
    Write-Log "Impossible de lire les rôles du projet $projectKey (skip). Err=$($_.Exception.Message)" "ERROR"
    continue
  }

  $roleIndex = @{} # norm -> {Name, Url}
  foreach ($prop in $roles.PSObject.Properties) {
    $roleIndex[(Normalize-RoleName $prop.Name)] = [pscustomobject]@{ Name = $prop.Name; Url = $prop.Value }
  }

  # Mapping: groupe -> power max (via matrice) sur ce projet
  $GroupToPower = @{}
  foreach ($e in $entries) {
    $g   = [string]$e.Group
    $pwr = [int]$e.RolePower
    if (-not $GroupToPower.ContainsKey($g)) { $GroupToPower[$g] = $pwr }
    else { if ($pwr -gt $GroupToPower[$g]) { $GroupToPower[$g] = $pwr } }
  }

  # Rôles 1/2/3 réellement présents sur ce projet
  $rolesToCheck = @()
  foreach ($nr in @($Role1,$Role2,$Role3)) {
    if ($roleIndex.ContainsKey($nr)) { $rolesToCheck += $nr }
  }
  if ($rolesToCheck.Count -eq 0) {
    Write-Log "Aucun rôle 1/2/3 détecté sur $projectKey (skip)" "WARN"
    continue
  }

  # Charger nominatives (users) par rôle 1/2/3
  $NominativeByRole = @{} # normRole -> HashSet(accountId)
  $RoleUrlByRole    = @{} # normRole -> url
  $RoleNameByRole   = @{} # normRole -> display role name

  foreach ($nr in $rolesToCheck) {
    $roleUrl  = [string]$roleIndex[$nr].Url
    $roleName = [string]$roleIndex[$nr].Name

    $RoleUrlByRole[$nr]  = $roleUrl
    $RoleNameByRole[$nr] = $roleName

    try {
      $detail = Invoke-Jira -Method GET -Url $roleUrl -Headers $Headers
      $stats.RoleDetailsRead++
    } catch {
      Write-Log "Impossible de lire détail rôle '$roleName' ($projectKey) => skip role. Err=$($_.Exception.Message)" "ERROR"
      continue
    }

    $usersInRole = Get-RoleActorsUsers -RoleDetail $detail
    $set = New-Object System.Collections.Generic.HashSet[string]
    foreach ($u in $usersInRole) {
      if ($u.accountId) { [void]$set.Add([string]$u.accountId) }
    }
    $NominativeByRole[$nr] = $set
  }

  # Candidats = union des accountId nominativement dans 1/2/3
  $candidateSet = New-Object System.Collections.Generic.HashSet[string]
  foreach ($nr in $NominativeByRole.Keys) {
    foreach ($aid in $NominativeByRole[$nr]) { [void]$candidateSet.Add($aid) }
  }
  if ($candidateSet.Count -eq 0) { continue }

  # Cache : maxPower par user sur ce projet (via groupes DSIM)
  $UserMaxPowerCache = @{}

  foreach ($aid in $candidateSet) {
    # limiter aux users DSIM
    if (-not $UserToGroups.ContainsKey($aid)) { continue }

    # max power via groupes DSIM de la matrice pour CE projet
    if (-not $UserMaxPowerCache.ContainsKey($aid)) {
      $maxPower = 0
      $userGroups = $UserToGroups[$aid]
      foreach ($g in $GroupToPower.Keys) {
        if ($userGroups.Contains($g)) {
          $pwr = [int]$GroupToPower[$g]
          if ($pwr -gt $maxPower) { $maxPower = $pwr }
        }
      }
      $UserMaxPowerCache[$aid] = $maxPower
    }

    $maxPower = [int]$UserMaxPowerCache[$aid]
    if ($maxPower -le 0) { continue } # pas de rôle via groupes DSIM pour ce projet

    $stats.CandidatesFound++

    $u  = $Users[$aid]
    $dn = if ($u) { $u.displayName } else { "" }
    $em = if ($u) { $u.email }      else { "" }
    $who = if ($em) { "$dn <$em>" } else { "$dn (accountId=$aid)" }

    # Pour chaque rôle nominatif 1/2/3, supprimer si power(role) <= maxPower (redondant ou inférieur)
    foreach ($nr in $NominativeByRole.Keys) {
      $roleName = $RoleNameByRole[$nr]
      $rolePower = Get-RolePowerFromRoleName -RoleName $roleName
      if ($rolePower -le 0) { continue }

      if (-not $NominativeByRole[$nr].Contains($aid)) { continue }

      if ($rolePower -le $maxPower) {
        $roleUrl = $RoleUrlByRole[$nr]

        if (-not $ApplyMode) {
          $stats.WouldDelete++
          Write-Log "Supprimerait nominatif: $who | Projet=$projectKey | Rôle=$roleName | Groupe DSIM donne power=$maxPower (rolePower=$rolePower <= max)" "DRYRUN"
        } else {
          $res = Remove-UserFromRole -RoleUrl $roleUrl -AccountId $aid -Headers $Headers
          if ($res.Success) {
            $stats.Deleted++
            Write-Log "Supprimé nominatif: $who | Projet=$projectKey | Rôle=$roleName | Groupe DSIM donne power=$maxPower (rolePower=$rolePower <= max)" "OK"
          } else {
            $stats.DeleteErrors++
            Write-Log "DELETE KO (on continue): $who | Projet=$projectKey | Rôle=$roleName | Url=$($res.Url) | Err=$($res.Error)" "ERROR"
          }
        }
      }
    }
  }
}

Write-Log "---- Récap ----" "INFO"
Write-Log ("ProjectsAnalyzed={0} RoleDetailsRead={1} CandidatesFound={2} WouldDelete={3} Deleted={4} DeleteErrors={5}" -f `
  $stats.ProjectsAnalyzed, $stats.RoleDetailsRead, $stats.CandidatesFound, $stats.WouldDelete, $stats.Deleted, $stats.DeleteErrors) "OK"

Write-Log "Terminé." "OK"
Write-Log "Log sauvegardé: $global:LogFile" "OK"