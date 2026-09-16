<# ---------------------------------------------------------------------------
List-ProjectManagers-ByCategory.ps1

But:
- Pour une catégorie de projet (ou toutes), lister pour chaque projet
  les utilisateurs & groupes ayant le rôle "1. Gestionnaire Projet".
- Ajouter une ALERTE si le projet n'a pas de gestionnaire projet
  autre que "site-admins" :
    => ligne CSV ActorType=Alert, ActorName="Beware Project Manager"

Sorties:
- Log: .\logs\List-ProjectManagers-ByCategory_yyyyMMdd_HHmmss.log
- Export CSV: .\exports\ProjectManagers_yyyyMMdd_HHmmss.csv

Colonnes CSV:
- CategoryId, CategoryName, ProjectKey, ProjectName, RoleName,
  ActorType (Group/User/Alert), ActorName, AccountId, ActorEmail

UI:
- Si ni -AllCategories ni -CategoryId/-CategoryName => formulaire de sélection
--------------------------------------------------------------------------- #>

[CmdletBinding()]
param(
  [string] $SiteUrl   = "https://jiradot.atlassian.net",
  [string] $UserEmail = "frederic.guedj@harmonie-mutuelle.fr",

  [string] $CredentialFile = $null,
  [switch] $ResetCredential,

  # Ciblage catégorie
  [switch] $AllCategories,
  [string] $CategoryId,
  [string] $CategoryName,

  # Rôle à analyser
  [string] $RoleNameToFind = "1. Gestionnaire Projet",

  # Groupe "admin" à ignorer pour la détection "pas de PM"
  [string] $AdminGroupName = "site-admins",

  # Texte d'alerte à mettre dans le CSV
  [string] $BewareText = "Beware Project Manager"
)

try { Set-StrictMode -Off } catch {}

# ------------------------- ScriptDir / Credential file -------------------------
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

# ------------------------- Logging + Export paths -------------------------
$global:LogFile = $null
$global:ExportCsv = $null

function Ensure-Dir([string]$path) {
  if (-not (Test-Path $path)) { New-Item -ItemType Directory -Path $path -Force | Out-Null }
}

function Initialize-Paths {
  param([string]$BaseDir)

  $logsDir = Join-Path $BaseDir "logs"
  $expDir  = Join-Path $BaseDir "exports"
  Ensure-Dir $logsDir
  Ensure-Dir $expDir

  $ts = Get-Date -Format "yyyyMMdd_HHmmss"
  $log = Join-Path $logsDir ("List-ProjectManagers-ByCategory_{0}.log" -f $ts)
  $csv = Join-Path $expDir  ("ProjectManagers_{0}.csv" -f $ts)

  Set-Content -Path $log -Value "" -Encoding UTF8
  return @{ Log = $log; Csv = $csv }
}

function Write-Log {
  param(
    [Parameter(Mandatory=$true)][string]$Message,
    [ValidateSet('INFO','WARN','ERROR','OK')][string]$Level = 'INFO'
  )
  $ts = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
  $line = "[$ts][$Level] $Message"
  Write-Host $line
  if ($global:LogFile) { Add-Content -Path $global:LogFile -Value $line -Encoding UTF8 }
}

$paths = Initialize-Paths -BaseDir $ScriptDir
$global:LogFile  = $paths.Log
$global:ExportCsv = $paths.Csv

Write-Log "LogFile: $global:LogFile" "INFO"
Write-Log "ExportCsv: $global:ExportCsv" "INFO"

# ------------------------- Network prerequisites -------------------------
try { [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12 } catch {}
try {
  $proxy = [System.Net.WebRequest]::GetSystemWebProxy()
  $proxy.Credentials = [System.Net.CredentialCache]::DefaultCredentials
  [System.Net.WebRequest]::DefaultWebProxy = $proxy
} catch {}

$SiteUrl = $SiteUrl.TrimEnd('/')

# ------------------------- Helpers -------------------------
function Get-SafeProp {
  param([Parameter(Mandatory=$true)]$Object, [Parameter(Mandatory=$true)][string]$PropName)
  if ($null -eq $Object) { return $null }
  $p = $Object.PSObject.Properties[$PropName]
  if ($p) { return $p.Value }
  return $null
}

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

# ------------------------- Credentials (DPAPI) -------------------------
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
  Ensure-Dir $dir
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
    if ($_.ErrorDetails -and $_.ErrorDetails.Message) { $details = "$details | $($_.ErrorDetails.Message)" }
    throw "Erreur API Jira ($Method $Url) : $details"
  }
}

# ------------------------- Jira functions -------------------------
function Get-AllProjectCategories {
  param([string]$SiteUrl, [hashtable]$Headers)
  return @(Invoke-Jira -Method GET -Url "$SiteUrl/rest/api/3/projectCategory" -Headers $Headers)
}

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

    $startAt += [int]$resp.maxResults
    if ($startAt -ge [int]$resp.total) { break }
  }

  return $all
}

function Get-GroupActorsFromRoleDetail {
  param([object]$RoleDetail)
  $groups = New-Object System.Collections.Generic.HashSet[string]
  foreach ($a in @($RoleDetail.actors)) {
    $actorGroup = Get-SafeProp -Object $a -PropName "actorGroup"
    if ($actorGroup) {
      $name = [string](Get-SafeProp -Object $actorGroup -PropName "name")
      if ($name) { [void]$groups.Add($name) }
      continue
    }
    $type = [string](Get-SafeProp -Object $a -PropName "type")
    if ($type -and ($type -like "*group*")) {
      $name = [string](Get-SafeProp -Object $a -PropName "name")
      if ($name) { [void]$groups.Add($name); continue }
      $dn = [string](Get-SafeProp -Object $a -PropName "displayName")
      if ($dn) { [void]$groups.Add($dn); continue }
    }
  }
  return $groups
}

function Get-UserActorsFromRoleDetail {
  param([object]$RoleDetail)
  $users = New-Object System.Collections.Generic.List[object]
  foreach ($a in @($RoleDetail.actors)) {
    $actorUser = Get-SafeProp -Object $a -PropName "actorUser"
    if ($actorUser) {
      $aid = [string](Get-SafeProp -Object $actorUser -PropName "accountId")
      if ($aid) {
        $users.Add([pscustomobject]@{
          accountId   = $aid
          displayName = [string](Get-SafeProp -Object $actorUser -PropName "displayName")
        })
      }
    }
  }
  return $users
}

# Cache email: accountId -> email
$UserEmailCache = @{}

function Get-UserEmailByAccountId {
  param(
    [string]$SiteUrl,
    [hashtable]$Headers,
    [string]$AccountId
  )

  if (-not $AccountId) { return "" }

  if ($UserEmailCache.ContainsKey($AccountId)) {
    return $UserEmailCache[$AccountId]
  }

  # Jira Cloud: GET /rest/api/3/user?accountId=...
  $enc = [Uri]::EscapeDataString($AccountId)
  $url = "$SiteUrl/rest/api/3/user?accountId=$enc"

  try {
    $u = Invoke-Jira -Method GET -Url $url -Headers $Headers
    $mail = [string](Get-SafeProp -Object $u -PropName "emailAddress")
    # Certain sites peuvent masquer l'email; dans ce cas: string vide
    $UserEmailCache[$AccountId] = $mail
    return $mail
  } catch {
    Write-Log "Impossible de récupérer l'email pour accountId=$AccountId (on continue) : $($_.Exception.Message)" "WARN"
    $UserEmailCache[$AccountId] = ""
    return ""
  }
}

# ------------------------- UI selection of category -------------------------
function Select-CategoryUi {
  param([object[]]$Categories)

  Add-Type -AssemblyName System.Windows.Forms | Out-Null
  Add-Type -AssemblyName System.Drawing | Out-Null

  $form = New-Object System.Windows.Forms.Form
  $form.Text = "Sélection catégorie (ou Toutes)"
  $form.Size = New-Object System.Drawing.Size(640, 160)
  $form.StartPosition = "CenterScreen"
  $form.TopMost = $true

  $label = New-Object System.Windows.Forms.Label
  $label.AutoSize = $true
  $label.Location = New-Object System.Drawing.Point(15, 15)
  $label.Text = "Choisis une catégorie (ou 'TOUTES') :"
  $form.Controls.Add($label)

  $combo = New-Object System.Windows.Forms.ComboBox
  $combo.Location = New-Object System.Drawing.Point(18, 45)
  $combo.Width = 590
  $combo.DropDownStyle = "DropDownList"
  [void]$combo.Items.Add("TOUTES")
  foreach ($c in $Categories) {
    $combo.Items.Add(("{0} (id={1})" -f $c.name, $c.id)) | Out-Null
  }
  $combo.SelectedIndex = 0
  $form.Controls.Add($combo)

  $ok = New-Object System.Windows.Forms.Button
  $ok.Text = "OK"
  $ok.Location = New-Object System.Drawing.Point(450, 80)
  $ok.Add_Click({ $form.Tag = "OK"; $form.Close() })
  $form.Controls.Add($ok)

  $cancel = New-Object System.Windows.Forms.Button
  $cancel.Text = "Annuler"
  $cancel.Location = New-Object System.Drawing.Point(530, 80)
  $cancel.Add_Click({ $form.Tag = "CANCEL"; $form.Close() })
  $form.Controls.Add($cancel)

  $form.ShowDialog() | Out-Null
  if ($form.Tag -ne "OK") { throw "Exécution annulée." }

  $sel = [string]$combo.SelectedItem
  if ($sel -eq "TOUTES") {
    return [pscustomobject]@{ All = $true; CategoryId = $null; CategoryName = $null }
  }

  $m = [regex]::Match($sel, '\(id=(\d+)\)\s*$')
  $id = if ($m.Success) { $m.Groups[1].Value } else { $null }
  $name = $sel -replace '\s*\(id=\d+\)\s*$', ''

  return [pscustomobject]@{ All = $false; CategoryId = $id; CategoryName = $name }
}

# ------------------------- Main -------------------------
Write-Log "SiteUrl: $SiteUrl" "INFO"
Write-Log "CredentialFile: $CredentialFile" "INFO"
Write-Log "RoleNameToFind: $RoleNameToFind" "INFO"
Write-Log "AdminGroupName (ignored for beware): $AdminGroupName" "INFO"
Write-Log "BewareText: $BewareText" "INFO"

$cred = Get-JiraCredential -CredentialFile $CredentialFile -UserEmail $UserEmail -Reset:$ResetCredential
$headers = Get-JiraAuthHeader -Credential $cred

$me = Invoke-Jira -Method GET -Url "$SiteUrl/rest/api/3/myself" -Headers $headers
Write-Log ("Authentifié: {0}" -f [string](Get-SafeProp -Object $me -PropName "displayName")) "OK"

$categories = Get-AllProjectCategories -SiteUrl $SiteUrl -Headers $headers
Write-Log ("Catégories trouvées: {0}" -f $categories.Count) "INFO"

# Déterminer le périmètre catégorie
if (-not $AllCategories -and [string]::IsNullOrWhiteSpace($CategoryId) -and [string]::IsNullOrWhiteSpace($CategoryName)) {
  $sel = Select-CategoryUi -Categories $categories
  $AllCategories = [bool]$sel.All
  $CategoryId = $sel.CategoryId
  $CategoryName = $sel.CategoryName
}

if ($AllCategories) {
  $targetCategories = $categories | Sort-Object name
  Write-Log "Mode: Toutes catégories" "INFO"
} else {
  $targetCategories = @()

  if ($CategoryId) {
    $targetCategories = @($categories | Where-Object { [string]$_.id -eq [string]$CategoryId })
  } elseif ($CategoryName) {
    $targetCategories = @($categories | Where-Object { $_.name -and ($_.name.Trim().ToLowerInvariant() -eq $CategoryName.Trim().ToLowerInvariant()) })
  }

  if (-not $targetCategories -or $targetCategories.Count -eq 0) {
    throw "Catégorie non trouvée (CategoryId='$CategoryId', CategoryName='$CategoryName')."
  }
  Write-Log ("Mode: Catégorie ciblée => {0} (id={1})" -f $targetCategories[0].name, $targetCategories[0].id) "INFO"
}

# Charger tous les projets (1 seule fois) puis filtrer
Write-Log "Chargement de tous les projets (avec catégories)..." "INFO"
$allProjects = Get-AllProjectsWithCategory -SiteUrl $SiteUrl -Headers $headers
Write-Log ("Projets récupérés: {0}" -f $allProjects.Count) "OK"

$roleNormToFind = Normalize-RoleName $RoleNameToFind
$adminUpper = $AdminGroupName.Trim().ToUpperInvariant()

$results = New-Object System.Collections.Generic.List[object]

foreach ($cat in $targetCategories) {
  $catId = [string]$cat.id
  $catName = [string]$cat.name

  $projects = @($allProjects | Where-Object { $_.categoryId -eq $catId })
  Write-Log ("--- Catégorie: {0} (id={1}) => {2} projets ---" -f $catName, $catId, $projects.Count) "INFO"

  foreach ($p in $projects) {
    $pk = $p.key
    $pn = $p.name

    Write-Log ("Projet {0} - {1}" -f $pk, $pn) "INFO"

    # Rôles du projet
    $roles = $null
    try {
      $roles = Invoke-Jira -Method GET -Url "$SiteUrl/rest/api/3/project/$pk/role" -Headers $headers
    } catch {
      Write-Log "KO lecture rôles projet $pk : $($_.Exception.Message)" "ERROR"
      # Alarme : impossible d'assurer qu'il y a un PM => beware
      $results.Add([pscustomobject]@{
        CategoryId   = $catId
        CategoryName = $catName
        ProjectKey   = $pk
        ProjectName  = $pn
        RoleName     = $RoleNameToFind
        ActorType    = "Alert"
        ActorName    = $BewareText
        AccountId    = ""
        ActorEmail   = ""
      })
      continue
    }

    # Trouver l'URL du rôle
    $roleUrl = $null
    $roleDisplayName = $null
    foreach ($prop in $roles.PSObject.Properties) {
      if ((Normalize-RoleName $prop.Name) -eq $roleNormToFind) {
        $roleUrl = [string]$prop.Value
        $roleDisplayName = [string]$prop.Name
        break
      }
    }

    if (-not $roleUrl) {
      Write-Log "Rôle '$RoleNameToFind' introuvable sur $pk" "WARN"
      # Alarme
      $results.Add([pscustomobject]@{
        CategoryId   = $catId
        CategoryName = $catName
        ProjectKey   = $pk
        ProjectName  = $pn
        RoleName     = $RoleNameToFind
        ActorType    = "Alert"
        ActorName    = $BewareText
        AccountId    = ""
        ActorEmail   = ""
      })
      continue
    }

    # Détail du rôle => acteurs
    $detail = $null
    try {
      $detail = Invoke-Jira -Method GET -Url $roleUrl -Headers $headers
    } catch {
      Write-Log "KO lecture détail rôle '$roleDisplayName' ($pk) : $($_.Exception.Message)" "ERROR"
      # Alarme
      $results.Add([pscustomobject]@{
        CategoryId   = $catId
        CategoryName = $catName
        ProjectKey   = $pk
        ProjectName  = $pn
        RoleName     = $roleDisplayName
        ActorType    = "Alert"
        ActorName    = $BewareText
        AccountId    = ""
        ActorEmail   = ""
      })
      continue
    }

    $groups = Get-GroupActorsFromRoleDetail -RoleDetail $detail
    $users  = Get-UserActorsFromRoleDetail  -RoleDetail $detail

    # Export acteurs (groupes + users)
    foreach ($g in $groups) {
      $results.Add([pscustomobject]@{
        CategoryId   = $catId
        CategoryName = $catName
        ProjectKey   = $pk
        ProjectName  = $pn
        RoleName     = $roleDisplayName
        ActorType    = "Group"
        ActorName    = $g
        AccountId    = ""
        ActorEmail   = ""
      })
    }

    foreach ($u in $users) {
      $mail = Get-UserEmailByAccountId -SiteUrl $SiteUrl -Headers $headers -AccountId $u.accountId
      $results.Add([pscustomobject]@{
        CategoryId   = $catId
        CategoryName = $catName
        ProjectKey   = $pk
        ProjectName  = $pn
        RoleName     = $roleDisplayName
        ActorType    = "User"
        ActorName    = $u.displayName
        AccountId    = $u.accountId
        ActorEmail   = $mail
      })
    }

    # ---- RÈGLE "Beware" ----
    # On alerte si : aucun user, et aucun groupe autre que site-admins
    $hasNonAdminGroup = $false
    foreach ($g in $groups) {
      if ($g -and ($g.Trim().ToUpperInvariant() -ne $adminUpper)) {
        $hasNonAdminGroup = $true
        break
      }
    }

    $hasAnyUser = ($users -and $users.Count -gt 0)

    if ((-not $hasAnyUser) -and (-not $hasNonAdminGroup)) {
      Write-Log "ALERTE: $pk n'a pas de PM autre que '$AdminGroupName' => '$BewareText'" "WARN"

      $results.Add([pscustomobject]@{
        CategoryId   = $catId
        CategoryName = $catName
        ProjectKey   = $pk
        ProjectName  = $pn
        RoleName     = $roleDisplayName
        ActorType    = "Alert"
        ActorName    = $BewareText
        AccountId    = ""
        ActorEmail   = ""
      })
    }
  }
}

# Export CSV
$results | Export-Csv -Path $global:ExportCsv -Delimiter ';' -NoTypeInformation -Encoding UTF8
Write-Log ("Export CSV terminé: {0} lignes -> {1}" -f $results.Count, $global:ExportCsv) "OK"

Write-Log "Terminé." "OK"
Write-Log "Log sauvegardé: $global:LogFile" "OK"