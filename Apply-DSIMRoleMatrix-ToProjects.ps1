<# ---------------------------------------------------------------------------
Apply-DSIMRoleMatrix-ToProjects.ps1

- DRY-RUN ou EXECUTE via boite de dialogue (si -Apply n’est pas fourni)
- Ajout des groupes manquants dans le rôle cible
- Option ForceUnique : 1 seul rôle par groupe d’habilitation (sur rôles 1/2/3)
- Log fichier automatique: .\logs\Apply-DSIMRoleMatrix_yyyyMMdd_HHmmss.log

CSV supportés:

A) Format "par projet"
ProjectKey;Role;Group

B) Format "par catégorie"
Project Category;DSIM Group Name;...;Project Category Id;...;Rôle
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

  # Force : 1 seul rôle parmi les rôles gérés (1/2/3) par groupe
  [switch] $ForceUnique,

  # reset DPAPI credential
  [switch] $ResetCredential
)

# ------------------------- Déterminer le répertoire du script (robuste) -------------------------
$ScriptDir = $null
if ($PSCommandPath) {
  $ScriptDir = Split-Path -Parent $PSCommandPath
} elseif ($MyInvocation -and $MyInvocation.MyCommand -and $MyInvocation.MyCommand.Path) {
  $ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
}
if ([string]::IsNullOrWhiteSpace($ScriptDir)) {
  $ScriptDir = (Get-Location).Path
}
if ([string]::IsNullOrWhiteSpace($CredentialFile)) {
  $CredentialFile = Join-Path $ScriptDir "jira-credential.xml"
}

# ------------------------- Logging fichier -------------------------
$global:LogFile = $null

function Initialize-LogFile {
  param([string]$BaseDir)

  $logsDir = Join-Path $BaseDir "logs"
  if (-not (Test-Path $logsDir)) { New-Item -ItemType Directory -Path $logsDir -Force | Out-Null }

  $name = "Apply-DSIMRoleMatrix_{0}.log" -f (Get-Date -Format "yyyyMMdd_HHmmss")
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
  if ($global:LogFile) {
    Add-Content -Path $global:LogFile -Value $line -Encoding UTF8
  }
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

# ------------------------- Normalisation -------------------------
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

  # retire accents (ex: "Rôle" -> "role")
  $t = $t.Normalize([Text.NormalizationForm]::FormD)
  $sb = New-Object System.Text.StringBuilder
  foreach ($ch in $t.ToCharArray()) {
    if ([Globalization.CharUnicodeInfo]::GetUnicodeCategory($ch) -ne [Globalization.UnicodeCategory]::NonSpacingMark) {
      [void]$sb.Append($ch)
    }
  }
  $t = $sb.ToString()

  # garde seulement a-z0-9
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
  $b64  = [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes($pair))

  return @{
    Authorization = "Basic $b64"
    Accept        = "application/json"
    # Ne PAS mettre Content-Type ici (on le passe via -ContentType dans Invoke-RestMethod)
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
      # JSON compact
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
  param(
    [string]$Path,
    [string]$InitialDirectory
  )

  if (Test-Path $Path) { return $Path }

  Write-Log "CSV introuvable: $Path" "WARN"
  Write-Log "Ouverture d'une fenêtre de sélection de fichier CSV..." "INFO"

  Add-Type -AssemblyName System.Windows.Forms | Out-Null
  $dlg = New-Object System.Windows.Forms.OpenFileDialog
  $dlg.Filter = "CSV (*.csv)|*.csv|All files (*.*)|*.*"
  $dlg.Title  = "Sélectionner le CSV DSIM"
  if (Test-Path $InitialDirectory) { $dlg.InitialDirectory = $InitialDirectory }

  if ($dlg.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) {
    throw "Aucun fichier CSV sélectionné."
  }
  return $dlg.FileName
}

# ------------------------- UI: choix mode DRYRUN/EXECUTE + ForceUnique -------------------------
function Select-RunMode {
  param(
    [bool]$ApplyAlreadyProvided,
    [bool]$ForceUniqueAlreadyProvided
  )

  if ($ApplyAlreadyProvided) {
    return [pscustomobject]@{ Apply = $true; ForceUnique = $ForceUniqueAlreadyProvided }
  }

  Add-Type -AssemblyName System.Windows.Forms | Out-Null
  Add-Type -AssemblyName System.Drawing | Out-Null

  $form = New-Object System.Windows.Forms.Form
  $form.Text = "DSIM - Mode d'exécution"
  $form.Size = New-Object System.Drawing.Size(520, 220)
  $form.StartPosition = "CenterScreen"
  $form.TopMost = $true

  $lbl = New-Object System.Windows.Forms.Label
  $lbl.AutoSize = $true
  $lbl.Location = New-Object System.Drawing.Point(15, 15)
  $lbl.Text = "Choisis le mode:"
  $form.Controls.Add($lbl)

  $rbDry = New-Object System.Windows.Forms.RadioButton
  $rbDry.Text = "DRY-RUN (simulation)"
  $rbDry.Location = New-Object System.Drawing.Point(18, 45)
  $rbDry.Checked = $true
  $form.Controls.Add($rbDry)

  $rbExec = New-Object System.Windows.Forms.RadioButton
  $rbExec.Text = "EXECUTE (applique dans Jira)"
  $rbExec.Location = New-Object System.Drawing.Point(18, 70)
  $form.Controls.Add($rbExec)

  $cbForce = New-Object System.Windows.Forms.CheckBox
  $cbForce.Text = "Forcer 1 seul rôle par groupe (sur rôles 1/2/3)"
  $cbForce.Location = New-Object System.Drawing.Point(18, 105)
  $cbForce.AutoSize = $true
  $cbForce.Checked = $ForceUniqueAlreadyProvided
  $form.Controls.Add($cbForce)

  $btnOk = New-Object System.Windows.Forms.Button
  $btnOk.Text = "OK"
  $btnOk.Location = New-Object System.Drawing.Point(330, 135)
  $btnOk.Add_Click({
    $form.Tag = "OK"
    $form.Close()
  })
  $form.Controls.Add($btnOk)

  $btnCancel = New-Object System.Windows.Forms.Button
  $btnCancel.Text = "Annuler"
  $btnCancel.Location = New-Object System.Drawing.Point(410, 135)
  $btnCancel.Add_Click({
    $form.Tag = "CANCEL"
    $form.Close()
  })
  $form.Controls.Add($btnCancel)

  $form.ShowDialog() | Out-Null

  if ($form.Tag -ne "OK") {
    throw "Exécution annulée par l'utilisateur."
  }

  return [pscustomobject]@{
    Apply       = $rbExec.Checked
    ForceUnique = $cbForce.Checked
  }
}

# ------------------------- Mapping rôles -------------------------
$RoleNameMap = @{
  "1-gestionnaire"  = "1. Gestionnaire Projet"
  "2-utilisateur"   = "2. Utilisateur"
  "3-lecture seule" = "3. Lecture seule"
  "ne pas toucher"  = $null
}

$ManagedRoleNames = @(
  "1. Gestionnaire Projet",
  "2. Utilisateur",
  "3. Lecture seule"
)
$ManagedRoleNorms = @{}
foreach ($rn in $ManagedRoleNames) { $ManagedRoleNorms[(Normalize-RoleName $rn)] = $true }

# ------------------------- Jira: liste projets (pour mode "par catégorie") -------------------------
function Get-AllProjectsWithCategory {
  param(
    [Parameter(Mandatory=$true)][string]$SiteUrl,
    [Parameter(Mandatory=$true)][hashtable]$Headers
  )

  $all = New-Object System.Collections.Generic.List[object]
  $startAt = 0
  $max = 50

  while ($true) {
    $url = "$SiteUrl/rest/api/3/project/search?startAt=$startAt&maxResults=$max&expand=projectCategory"
    $resp = Invoke-Jira -Method GET -Url $url -Headers $Headers

    foreach ($p in $resp.values) {
      $all.Add([pscustomobject]@{
        key          = $p.key
        name         = $p.name
        categoryId   = if ($p.projectCategory) { $p.projectCategory.id } else { $null }
        categoryName = if ($p.projectCategory) { $p.projectCategory.name } else { $null }
      })
    }

    $startAt += $resp.maxResults
    if ($startAt -ge $resp.total) { break }
  }

  return $all
}

# ------------------------- Helpers rôles/actors -------------------------
function Get-GroupActorsFromRoleDetail {
  param([object]$RoleDetail)

  $groups = @{}
  foreach ($a in @($RoleDetail.actors)) {

    if ($a.PSObject.Properties['actorGroup']) {
      $ag = $a.actorGroup
      if ($ag -and $ag.PSObject.Properties['name'] -and -not [string]::IsNullOrWhiteSpace($ag.name)) {
        $groups[$ag.name] = $true
        continue
      }
    }

    if ($a.PSObject.Properties['type'] -and ($a.type -like "*group*")) {
      if ($a.PSObject.Properties['name'] -and $a.name) { $groups[$a.name] = $true; continue }
      if ($a.PSObject.Properties['displayName'] -and $a.displayName) { $groups[$a.displayName] = $true; continue }
    }
  }
  return $groups
}

function Remove-GroupFromRole {
  <#
    Sécurité: n'échoue pas le script si le DELETE échoue.
    Retourne un objet: Success, Url, Error
  #>
  param(
    [Parameter(Mandatory=$true)][string]$RoleUrl,
    [Parameter(Mandatory=$true)][string]$GroupName,
    [Parameter(Mandatory=$true)][hashtable]$Headers
  )

  $enc = [Uri]::EscapeDataString($GroupName)

  # IMPORTANT: ne pas utiliser -like "*?*" (le ? est un joker)
  $sep = if ($RoleUrl.Contains('?')) { '&' } else { '?' }

  $url = "$RoleUrl$sep" + "group=$enc"

  try {
    Invoke-Jira -Method DELETE -Url $url -Headers $Headers | Out-Null
    return [pscustomobject]@{ Success = $true; Url = $url; Error = $null }
  } catch {
    return [pscustomobject]@{ Success = $false; Url = $url; Error = $_.Exception.Message }
  }
}

# ------------------------- Traitement d'un projet -------------------------
function Apply-Entries-ToProject {
  param(
    [Parameter(Mandatory=$true)][string]$ProjectKey,
    [Parameter(Mandatory=$true)][object[]]$Entries,   # each: Role, Group
    [Parameter(Mandatory=$true)][string]$SiteUrl,
    [Parameter(Mandatory=$true)][hashtable]$Headers,
    [switch]$ApplyMode,
    [switch]$ForceUniqueMode
  )

  Write-Log "---- Projet $ProjectKey ----" "INFO"

  $roles = Invoke-Jira -Method GET -Url "$SiteUrl/rest/api/3/project/$ProjectKey/role" -Headers $Headers

  $roleIndex = @{}
  foreach ($prop in $roles.PSObject.Properties) {
    $roleName = $prop.Name
    $roleUrl  = $prop.Value
    $roleIndex[(Normalize-RoleName $roleName)] = [pscustomobject]@{ Name = $roleName; Url = $roleUrl }
  }

  # Pré-chargement des rôles gérés (1/2/3) si ForceUnique activé
  $managedDetails = @{} # norm -> @{ Name, Url, Groups = @{...} }
  if ($ForceUniqueMode) {
    foreach ($norm in $ManagedRoleNorms.Keys) {
      if ($roleIndex.ContainsKey($norm)) {
        $ri = $roleIndex[$norm]
        $detail = Invoke-Jira -Method GET -Url $ri.Url -Headers $Headers
        $managedDetails[$norm] = @{
          Name   = $ri.Name
          Url    = $ri.Url
          Groups = (Get-GroupActorsFromRoleDetail -RoleDetail $detail)
        }
      }
    }
  }

  foreach ($e in $Entries) {
    $rawRole = [string]$e.Role
    $group   = [string]$e.Group

    if ([string]::IsNullOrWhiteSpace($rawRole) -or [string]::IsNullOrWhiteSpace($group)) {
      Write-Log "Entrée ignorée (Role/Group vide) sur $ProjectKey" "WARN"
      continue
    }

    $mappedRole = $rawRole
    $k = $rawRole.Trim().ToLowerInvariant()
    if ($RoleNameMap.ContainsKey($k)) { $mappedRole = $RoleNameMap[$k] }

    if ($null -eq $mappedRole) {
      Write-Log "Role='$rawRole' => 'Ne pas toucher' => skip ($ProjectKey / $group)" "INFO"
      continue
    }

    $normDesired = Normalize-RoleName $mappedRole
    if (-not $roleIndex.ContainsKey($normDesired)) {
      $available = ($roles.PSObject.Properties.Name -join " | ")
      Write-Log "Rôle introuvable: '$mappedRole' (depuis '$rawRole'). Rôles dispo: $available" "ERROR"
      continue
    }

    $desiredRoleName = $roleIndex[$normDesired].Name
    $desiredRoleUrl  = $roleIndex[$normDesired].Url

    # Vérifier présence dans rôle désiré
    $detailDesired   = Invoke-Jira -Method GET -Url $desiredRoleUrl -Headers $Headers
    $groupsDesired   = Get-GroupActorsFromRoleDetail -RoleDetail $detailDesired
    $alreadyInDesired = $groupsDesired.ContainsKey($group)

    if ($alreadyInDesired) {
      Write-Log "OK déjà présent: '$group' dans '$desiredRoleName' ($ProjectKey)" "OK"
    } else {
      if (-not $ApplyMode) {
        Write-Log "Ajouterait: '$group' -> '$desiredRoleName' ($ProjectKey)" "DRYRUN"
      } else {
        $body = @{ group = @($group) }

        try {
          Invoke-Jira -Method POST -Url $desiredRoleUrl -Headers $Headers -Body $body | Out-Null
          Write-Log "Ajouté: '$group' -> '$desiredRoleName' ($ProjectKey)" "OK"

          if ($ForceUniqueMode -and $managedDetails.ContainsKey($normDesired)) {
            $managedDetails[$normDesired].Groups[$group] = $true
          }
        } catch {
          # Sécurité POST : on log et on continue
          Write-Log "POST KO (on continue): '$group' -> '$desiredRoleName' ($ProjectKey) | Url=$desiredRoleUrl | Err=$($_.Exception.Message)" "ERROR"
          continue
        }
      }
    }

    # ForceUnique : retirer le groupe des autres rôles gérés (1/2/3)
    if ($ForceUniqueMode) {
      foreach ($normOther in $managedDetails.Keys) {
        if ($normOther -eq $normDesired) { continue }

        $other = $managedDetails[$normOther]
        if ($other.Groups.ContainsKey($group)) {

          if (-not $ApplyMode) {
            Write-Log "Retirerait: '$group' de '$($other.Name)' (garder uniquement '$desiredRoleName') ($ProjectKey)" "DRYRUN"
          } else {
            $res = Remove-GroupFromRole -RoleUrl $other.Url -GroupName $group -Headers $Headers
            if ($res.Success) {
              Write-Log "Retiré: '$group' de '$($other.Name)' (unicité: '$desiredRoleName') ($ProjectKey)" "OK"
              $other.Groups.Remove($group) | Out-Null
            } else {
              # Sécurité DELETE: on continue malgré l'erreur
              Write-Log "DELETE KO (on continue): '$group' de '$($other.Name)' ($ProjectKey) | Url=$($res.Url) | Err=$($res.Error)" "ERROR"
            }
          }
        }
      }
    }
  }
}

# ------------------------- Main -------------------------
Write-Log "ScriptDir: $ScriptDir" "INFO"
Write-Log "SiteUrl: $SiteUrl" "INFO"
Write-Log "CredentialFile: $CredentialFile" "INFO"

$choice          = Select-RunMode -ApplyAlreadyProvided:$Apply.IsPresent -ForceUniqueAlreadyProvided:$ForceUnique.IsPresent
$ApplyMode       = [bool]$choice.Apply
$ForceUniqueMode = [bool]$choice.ForceUnique

$modeText  = if ($ApplyMode) { "EXECUTE (écriture)" } else { "DRY-RUN (simulation)" }
$modeLevel = if ($ApplyMode) { "WARN" } else { "DRYRUN" }
Write-Log "Mode: $modeText" $modeLevel
Write-Log ("ForceUnique: {0}" -f $ForceUniqueMode) "INFO"

$CsvPath = Resolve-CsvPath -Path $CsvPath -InitialDirectory $ScriptDir
Write-Log "CSV: $CsvPath" "INFO"

$JiraCredential = Get-JiraCredential -CredentialFile $CredentialFile -UserEmail $UserEmail -Reset:$ResetCredential
$Headers        = Get-JiraAuthHeader -Credential $JiraCredential

$me = Invoke-Jira -Method GET -Url "$SiteUrl/rest/api/3/myself" -Headers $Headers
Write-Log ("Authentifié: {0} <{1}>" -f $me.displayName, $me.emailAddress) "OK"

$rows = Import-Csv -Path $CsvPath -Delimiter $Delimiter
if (-not $rows -or $rows.Count -eq 0) { throw "CSV vide: $CsvPath" }

$colMap = @{}
$rows[0].PSObject.Properties.Name | ForEach-Object {
  $colMap[(Normalize-HeaderName $_)] = $_
}
function Get-Col {
  param([string[]]$Candidates)
  foreach ($c in $Candidates) {
    $k = Normalize-HeaderName $c
    if ($colMap.ContainsKey($k)) { return $colMap[$k] }
  }
  return $null
}

# Détection format CSV
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

if ($format -eq "PROJECTKEY") {
  $byProject = $rows | Group-Object { ([string]$_.($projectKeyCol)).Trim() }

  foreach ($p in $byProject) {
    $projectKey = $p.Name
    if ([string]::IsNullOrWhiteSpace($projectKey)) { continue }

    $entries = @()
    foreach ($r in $p.Group) {
      $entries += [pscustomobject]@{
        Role  = ([string]$r.($roleColA)).Trim()
        Group = ([string]$r.($groupColA)).Trim()
      }
    }

    Apply-Entries-ToProject -ProjectKey $projectKey -Entries $entries -SiteUrl $SiteUrl -Headers $Headers -ApplyMode:$ApplyMode -ForceUniqueMode:$ForceUniqueMode
  }
}

if ($format -eq "CATEGORY") {
  Write-Log "Chargement de tous les projets + catégories (peut prendre un peu de temps)..." "INFO"
  $projects = Get-AllProjectsWithCategory -SiteUrl $SiteUrl -Headers $Headers
  Write-Log ("Projets récupérés: {0}" -f $projects.Count) "OK"

  $byCategory = $rows | Group-Object {
    $id = if ($categoryIdCol) { ([string]$_.($categoryIdCol)).Trim() } else { "" }
    $nm = if ($categoryNameCol) { ([string]$_.($categoryNameCol)).Trim() } else { "" }

    if (-not [string]::IsNullOrWhiteSpace($id)) { "ID:$id" }
    else { "NAME:$nm" }
  }

  foreach ($cat in $byCategory) {
    $catKey = $cat.Name

    $catId = $null
    $catName = $null
    if ($catKey -like "ID:*")   { $catId = $catKey.Substring(3) }
    if ($catKey -like "NAME:*") { $catName = $catKey.Substring(5) }

    # IMPORTANT: @() pour avoir .Count même si 0/1 résultat
    if ($catId) {
      $targetProjects = @($projects | Where-Object { $_.categoryId -eq $catId })
      Write-Log ("Catégorie ID={0} => {1} projets" -f $catId, $targetProjects.Count) "INFO"
    } else {
      $targetProjects = @($projects | Where-Object {
        $_.categoryName -and ($_.categoryName.Trim().ToLowerInvariant() -eq $catName.Trim().ToLowerInvariant())
      })
      Write-Log ("Catégorie Name='{0}' => {1} projets" -f $catName, $targetProjects.Count) "INFO"
    }

    if ($targetProjects.Count -eq 0) {
      Write-Log "Aucun projet trouvé pour $catKey (skip)" "WARN"
      continue
    }

    $entries = @()
    foreach ($r in $cat.Group) {
      $entries += [pscustomobject]@{
        Role  = ([string]$r.($roleColB)).Trim()
        Group = ([string]$r.($groupColB)).Trim()
      }
    }

    foreach ($proj in $targetProjects) {
      Apply-Entries-ToProject -ProjectKey $proj.key -Entries $entries -SiteUrl $SiteUrl -Headers $Headers -ApplyMode:$ApplyMode -ForceUniqueMode:$ForceUniqueMode
    }
  }
}

Write-Log "Terminé." "OK"
Write-Log "Log sauvegardé: $global:LogFile" "OK"