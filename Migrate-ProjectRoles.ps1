<# ---------------------------------------------------------------------------
Migrate-ProjectRoles_TargetProjects_AndPermissionScheme.ps1

INPUTS (via OpenFileDialog):
1) CSV Mapping rôles :
   - Colonne "Id" (id rôle source)
   - Colonne "IdRôle cible" :
        - "Ne pas changer" => ignore la ligne
        - sinon entier => id numérique du rôle cible

2) CSV Projets cibles :
   - Colonne "ProjectKey"

RÈGLES Permission Scheme:
- Si ProjectName commence par z_ / Z_ => appliquer "z_DSIM_Projets archivés"
- Sinon => appliquer "DSIM_Projets en cours v1.0"

Migration rôles:
- Ajout au rôle cible (si manquants)
- Suppression du rôle source

MODE:
- DRY-RUN / EXECUTE via formulaire, ou -Apply pour EXECUTE direct

OUTPUT:
- Log  : .\logs\Migrate-ProjectRoles_yyyyMMdd_HHmmss.log
- CSV  : .\exports\Migrate-ProjectRoles_Actions_yyyyMMdd_HHmmss.csv
--------------------------------------------------------------------------- #>

[CmdletBinding()]
param(
  [string] $SiteUrl   = "https://jiradot.atlassian.net",
  [string] $UserEmail = "frederic.guedj@harmonie-mutuelle.fr",
  [char]   $Delimiter = ';',

  [string] $CredentialFile = $null,
  [switch] $ResetCredential,

  # EXECUTE direct (sans formulaire)
  [switch] $Apply,

  # Scheme par défaut (projets "normaux")
  [string] $DefaultPermissionSchemeName = "DSIM_Projets en cours v1.0",

  # Scheme pour projets archivés (ProjectName commence par z_)
  [string] $ArchivedPermissionSchemeName = "z_DSIM_Projets archivés"
)

try { Set-StrictMode -Off } catch {}

# ------------------------- ScriptDir / Credential file -------------------------
$ScriptDir = $null
if ($PSCommandPath) { $ScriptDir = Split-Path -Parent $PSCommandPath }
elseif ($MyInvocation -and $MyInvocation.MyCommand -and $MyInvocation.MyCommand.Path) { $ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path }
if ([string]::IsNullOrWhiteSpace($ScriptDir)) { $ScriptDir = (Get-Location).Path }
if ([string]::IsNullOrWhiteSpace($CredentialFile)) { $CredentialFile = Join-Path $ScriptDir "jira-credential.xml" }

# ------------------------- Logging + Exports -------------------------
$global:LogFile = $null
$global:ActionsCsv = $null
$global:Actions = New-Object System.Collections.Generic.List[object]

function Ensure-Dir([string]$path) { if (-not (Test-Path $path)) { New-Item -ItemType Directory -Path $path -Force | Out-Null } }

function Init-Paths([string]$BaseDir) {
  $logsDir = Join-Path $BaseDir "logs"
  $expDir  = Join-Path $BaseDir "exports"
  Ensure-Dir $logsDir
  Ensure-Dir $expDir

  $ts = Get-Date -Format "yyyyMMdd_HHmmss"
  $log = Join-Path $logsDir ("Migrate-ProjectRoles_{0}.log" -f $ts)
  $csv = Join-Path $expDir  ("Migrate-ProjectRoles_Actions_{0}.csv" -f $ts)

  Set-Content -Path $log -Value "" -Encoding UTF8
  return @{ Log = $log; Csv = $csv }
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

function Add-Action {
  param(
    [string]$ProjectKey,
    [string]$ProjectName,

    [string]$ActionType,      # PermissionScheme / RoleMigration
    [string]$Action,          # SetScheme / Skip / AddToTarget / RemoveFromSource / Error
    [string]$Mode,            # DRYRUN / EXECUTE
    [string]$Result,          # OK / KO
    [string]$Details,

    [int]$FromRoleId = 0,
    [string]$FromRoleName = "",
    [int]$ToRoleId = 0,
    [string]$ToRoleName = "",

    [string]$ActorType = "",  # Group/User
    [string]$ActorName = "",
    [string]$AccountId = ""
  )

  $global:Actions.Add([pscustomobject]@{
    ProjectKey   = $ProjectKey
    ProjectName  = $ProjectName
    ActionType   = $ActionType
    Action       = $Action
    Mode         = $Mode
    Result       = $Result
    Details      = $Details

    FromRoleId   = $FromRoleId
    FromRoleName = $FromRoleName
    ToRoleId     = $ToRoleId
    ToRoleName   = $ToRoleName

    ActorType    = $ActorType
    ActorName    = $ActorName
    AccountId    = $AccountId
  })
}

$paths = Init-Paths -BaseDir $ScriptDir
$global:LogFile = $paths.Log
$global:ActionsCsv = $paths.Csv
Write-Log "LogFile: $global:LogFile" "INFO"
Write-Log "ActionsCsv: $global:ActionsCsv" "INFO"

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

function Normalize-HeaderName {
  param([AllowNull()][string]$s)
  if ([string]::IsNullOrWhiteSpace($s)) { return "" }
  $t = $s.Trim().ToLowerInvariant()
  $t = $t.Normalize([Text.NormalizationForm]::FormD)
  $sb = New-Object System.Text.StringBuilder
  foreach ($ch in $t.ToCharArray()) {
    if ([Globalization.CharUnicodeInfo]::GetUnicodeCategory($ch) -ne [Globalization.UnicodeCategory]::NonSpacingMark) { [void]$sb.Append($ch) }
  }
  $t = $sb.ToString()
  $t = $t -replace '[^a-z0-9]+', ''
  return $t
}

function Get-RoleIdFromRoleUrl([string]$roleUrl) {
  if ([string]::IsNullOrWhiteSpace($roleUrl)) { return $null }
  $m = [regex]::Match($roleUrl, '/role/(\d+)$')
  if ($m.Success) { return [int]$m.Groups[1].Value }
  return $null
}

# ------------------------- UI: file pickers -------------------------
function Pick-CsvFile([string]$Title, [string]$InitialDirectory) {
  Add-Type -AssemblyName System.Windows.Forms | Out-Null
  $dlg = New-Object System.Windows.Forms.OpenFileDialog
  $dlg.Filter = "CSV (*.csv)|*.csv|All files (*.*)|*.*"
  $dlg.Title  = $Title
  if ($InitialDirectory -and (Test-Path $InitialDirectory)) { $dlg.InitialDirectory = $InitialDirectory }
  if ($dlg.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) { throw "Aucun fichier sélectionné ($Title)." }
  return $dlg.FileName
}

# ------------------------- UI: DRYRUN/EXECUTE -------------------------
function Select-RunMode {
  param([bool]$ApplyAlreadyProvided)

  if ($ApplyAlreadyProvided) { return [pscustomobject]@{ Apply = $true } }

  Add-Type -AssemblyName System.Windows.Forms | Out-Null
  Add-Type -AssemblyName System.Drawing | Out-Null

  $form = New-Object System.Windows.Forms.Form
  $form.Text = "Migration rôles + Permission scheme"
  $form.Size = New-Object System.Drawing.Size(820, 240)
  $form.StartPosition = "CenterScreen"
  $form.TopMost = $true

  $lbl = New-Object System.Windows.Forms.Label
  $lbl.AutoSize = $true
  $lbl.Location = New-Object System.Drawing.Point(15, 15)
  $lbl.Text = "Permission scheme: z_* => '$ArchivedPermissionSchemeName' ; sinon => '$DefaultPermissionSchemeName'."
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
  $btnOk.Location = New-Object System.Drawing.Point(620, 150)
  $btnOk.Add_Click({ $form.Tag = "OK"; $form.Close() })
  $form.Controls.Add($btnOk)

  $btnCancel = New-Object System.Windows.Forms.Button
  $btnCancel.Text = "Annuler"
  $btnCancel.Location = New-Object System.Drawing.Point(700, 150)
  $btnCancel.Add_Click({ $form.Tag = "CANCEL"; $form.Close() })
  $form.Controls.Add($btnCancel)

  $form.ShowDialog() | Out-Null
  if ($form.Tag -ne "OK") { throw "Exécution annulée." }

  return [pscustomobject]@{ Apply = $rbExec.Checked }
}

# ------------------------- Credentials (DPAPI) -------------------------
function Get-JiraCredential {
  param([string]$CredentialFile, [string]$UserEmail, [switch]$Reset)

  if ($Reset -and (Test-Path $CredentialFile)) { Remove-Item -Path $CredentialFile -Force -ErrorAction SilentlyContinue }

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
  param([pscredential]$Credential)
  $tokenPlain = [System.Net.NetworkCredential]::new("", $Credential.Password).Password
  $pair = "{0}:{1}" -f $Credential.UserName, $tokenPlain
  $b64  = [Convert]::ToBase64String([System.Text.Encoding]::ASCII.GetBytes($pair))
  return @{ Authorization = "Basic $b64"; Accept = "application/json" }
}

function Invoke-Jira {
  param(
    [ValidateSet('GET','POST','PUT','DELETE')][string]$Method,
    [string]$Url,
    [hashtable]$Headers,
    $Body
  )
  $json = $null
  try {
    if ($null -ne $Body) {
      $json  = $Body | ConvertTo-Json -Depth 20 -Compress
      $bytes = [System.Text.Encoding]::UTF8.GetBytes($json)
      return Invoke-RestMethod -Method $Method -Uri $Url -Headers $Headers -ContentType "application/json; charset=utf-8" -Body $bytes -ErrorAction Stop
    } else {
      return Invoke-RestMethod -Method $Method -Uri $Url -Headers $Headers -ErrorAction Stop
    }
  } catch {
    $details = $_.Exception.Message
    if ($_.ErrorDetails -and $_.ErrorDetails.Message) { $details = "$details | $($_.ErrorDetails.Message)" }
    throw "Erreur API Jira ($Method $Url) : $details"
  }
}

# ------------------------- CSV readers -------------------------
function Read-RoleMappings {
  param([string]$RoleMappingCsv, [char]$Delimiter)

  $rows = Import-Csv -Path $RoleMappingCsv -Delimiter $Delimiter
  if (-not $rows -or $rows.Count -eq 0) { throw "RoleMappingCsv vide: $RoleMappingCsv" }

  $headers = @($rows[0].PSObject.Properties.Name)
  $colMap = @{}
  foreach ($h in $headers) { $colMap[(Normalize-HeaderName $h)] = $h }

  function Get-Col([string[]]$cands) {
    foreach ($c in $cands) {
      $k = Normalize-HeaderName $c
      if ($colMap.ContainsKey($k)) { return $colMap[$k] }
    }
    return $null
  }

  $idSrcCol = Get-Col @("Id","RoleId","IdRoleSource","IdRôleSource")
  $idDstCol = Get-Col @("IdRôle cible","IdRoleCible","IdRôleCible","TargetRoleId","IdRoleTarget")

  if (-not $idSrcCol -or -not $idDstCol) {
    throw ("CSV mapping invalide: colonnes requises 'Id' et 'IdRôle cible'. Colonnes trouvées: {0}" -f ($headers -join ", "))
  }

  $out = New-Object System.Collections.Generic.List[object]
  foreach ($r in $rows) {
    $idStr  = ([string]$r.($idSrcCol)).Trim()
    $dstStr = ([string]$r.($idDstCol)).Trim()

    if ([string]::IsNullOrWhiteSpace($idStr)) { continue }

    [int]$idSrc = 0
    if (-not [int]::TryParse($idStr, [ref]$idSrc)) { continue }

    if ([string]::IsNullOrWhiteSpace($dstStr)) { continue }
    if ($dstStr.Trim().ToLowerInvariant() -eq "ne pas changer") { continue }

    [int]$idDst = 0
    if (-not [int]::TryParse($dstStr, [ref]$idDst)) {
      throw "Valeur invalide dans 'IdRôle cible' pour Id=$idStr : '$dstStr' (attendu entier ou 'Ne pas changer')."
    }

    if ($idSrc -eq $idDst) { continue }

    $out.Add([pscustomobject]@{ SourceId = $idSrc; TargetId = $idDst })
  }

  return $out
}

function Read-TargetProjects {
  param([string]$ProjectsCsv, [char]$Delimiter)

  $rows = Import-Csv -Path $ProjectsCsv -Delimiter $Delimiter
  if (-not $rows -or $rows.Count -eq 0) { throw "CSV projets vide: $ProjectsCsv" }

  $headers = @($rows[0].PSObject.Properties.Name)
  $colMap = @{}
  foreach ($h in $headers) { $colMap[(Normalize-HeaderName $h)] = $h }

  $pkCol = $null
  foreach ($c in @("ProjectKey")) {
    $k = Normalize-HeaderName $c
    if ($colMap.ContainsKey($k)) { $pkCol = $colMap[$k]; break }
  }
  if (-not $pkCol) {
    throw ("CSV projets invalide: il faut une colonne 'ProjectKey'. Colonnes trouvées: {0}" -f ($headers -join ", "))
  }

  $set = New-Object System.Collections.Generic.HashSet[string]
  foreach ($r in $rows) {
    $pk = ([string]$r.($pkCol)).Trim()
    if ($pk) { [void]$set.Add($pk) }
  }
  return @($set)
}

# ------------------------- Permission scheme functions -------------------------
function Get-PermissionSchemeIdByName {
  param([string]$SiteUrl, [hashtable]$Headers, [string]$SchemeName)

  $schemes = Invoke-Jira -Method GET -Url "$SiteUrl/rest/api/3/permissionscheme" -Headers $Headers
  foreach ($s in @($schemes.permissionSchemes)) {
    $name = [string](Get-SafeProp -Object $s -PropName "name")
    if ($name -and ($name.Trim().ToLowerInvariant() -eq $SchemeName.Trim().ToLowerInvariant())) {
      $idStr = [string](Get-SafeProp -Object $s -PropName "id")
      [int]$id = 0
      if ([int]::TryParse($idStr, [ref]$id)) { return $id }
    }
  }
  return $null
}

function Get-ProjectPermissionScheme {
  param([string]$SiteUrl, [hashtable]$Headers, [string]$ProjectKey)
  return Invoke-Jira -Method GET -Url "$SiteUrl/rest/api/3/project/$ProjectKey/permissionscheme" -Headers $Headers
}

function Set-ProjectPermissionScheme {
  param([string]$SiteUrl, [hashtable]$Headers, [string]$ProjectKey, [int]$SchemeId)
  $body = @{ id = $SchemeId }
  return Invoke-Jira -Method PUT -Url "$SiteUrl/rest/api/3/project/$ProjectKey/permissionscheme" -Headers $Headers -Body $body
}

# ------------------------- Role actors + CRUD -------------------------
function Get-RoleActors {
  param([object]$RoleDetail)

  $groups = New-Object System.Collections.Generic.HashSet[string]
  $users  = New-Object System.Collections.Generic.HashSet[string] # accountId

  foreach ($a in @($RoleDetail.actors)) {
    $ag = Get-SafeProp -Object $a -PropName "actorGroup"
    if ($ag) {
      $gn = [string](Get-SafeProp -Object $ag -PropName "name")
      if ($gn) { [void]$groups.Add($gn) }
      continue
    }

    $au = Get-SafeProp -Object $a -PropName "actorUser"
    if ($au) {
      $aid = [string](Get-SafeProp -Object $au -PropName "accountId")
      if ($aid) { [void]$users.Add($aid) }
      continue
    }
  }

  return [pscustomobject]@{ Groups = $groups; Users = $users }
}

function Add-ActorsToRole {
  param([string]$RoleUrl, [string[]]$Groups, [string[]]$Users, [hashtable]$Headers)

  $body = @{}
  if ($Groups -and $Groups.Count -gt 0) { $body.group = @($Groups) }
  if ($Users  -and $Users.Count  -gt 0) { $body.user  = @($Users)  }
  if ($body.Keys.Count -eq 0) { return $null }

  return Invoke-Jira -Method POST -Url $RoleUrl -Headers $Headers -Body $body
}

function Remove-GroupFromRole {
  param([string]$RoleUrl, [string]$GroupName, [hashtable]$Headers)
  $enc = [Uri]::EscapeDataString($GroupName)
  $sep = if ($RoleUrl.Contains('?')) { '&' } else { '?' }
  $url = "$RoleUrl$sep" + "group=$enc"
  Invoke-Jira -Method DELETE -Url $url -Headers $Headers | Out-Null
  return $url
}

function Remove-UserFromRole {
  param([string]$RoleUrl, [string]$AccountId, [hashtable]$Headers)
  $enc = [Uri]::EscapeDataString($AccountId)
  $sep = if ($RoleUrl.Contains('?')) { '&' } else { '?' }
  $url = "$RoleUrl$sep" + "user=$enc"
  Invoke-Jira -Method DELETE -Url $url -Headers $Headers | Out-Null
  return $url
}

# ------------------------- MAIN -------------------------
Write-Log "SiteUrl: $SiteUrl" "INFO"
Write-Log "CredentialFile: $CredentialFile" "INFO"
Write-Log "DefaultPermissionSchemeName: $DefaultPermissionSchemeName" "INFO"
Write-Log "ArchivedPermissionSchemeName: $ArchivedPermissionSchemeName" "INFO"

$choice = Select-RunMode -ApplyAlreadyProvided:$Apply.IsPresent
$ApplyMode = [bool]$choice.Apply
$modeTxt = if ($ApplyMode) { "EXECUTE" } else { "DRYRUN" }
$modeLevel = if ($ApplyMode) { "WARN" } else { "DRYRUN" }
Write-Log "Mode: $modeTxt" $modeLevel

# Upload des 2 CSV
$RoleMappingCsv = Pick-CsvFile -Title "Uploader le CSV de mapping des rôles (Id / IdRôle cible)" -InitialDirectory "C:\Temp"
Write-Log "RoleMappingCsv: $RoleMappingCsv" "OK"

$ProjectsCsv = Pick-CsvFile -Title "Uploader le CSV des Projets Cibles (ProjetsCibles.csv avec ProjectKey)" -InitialDirectory "C:\Temp"
Write-Log "ProjectsCsv: $ProjectsCsv" "OK"

# Auth
$cred = Get-JiraCredential -CredentialFile $CredentialFile -UserEmail $UserEmail -Reset:$ResetCredential
$headers = Get-JiraAuthHeader -Credential $cred

$me = Invoke-Jira -Method GET -Url "$SiteUrl/rest/api/3/myself" -Headers $headers
Write-Log ("Authentifié: {0}" -f [string](Get-SafeProp -Object $me -PropName "displayName")) "OK"

# Lire mapping + projets
$mappings = Read-RoleMappings -RoleMappingCsv $RoleMappingCsv -Delimiter $Delimiter
Write-Log ("Mappings à appliquer (Id source -> Id cible): {0}" -f $mappings.Count) "INFO"
if ($mappings.Count -eq 0) { throw "Aucun mapping à appliquer (CSV vide ou 'Ne pas changer' partout)." }

$targetProjects = Read-TargetProjects -ProjectsCsv $ProjectsCsv -Delimiter $Delimiter
Write-Log ("Projets cibles: {0}" -f $targetProjects.Count) "INFO"
if ($targetProjects.Count -eq 0) { throw "Aucun projectKey dans le CSV projets cibles." }

# Rôles globaux (pour noms)
Write-Log "Récupération des rôles globaux via /rest/api/3/role..." "INFO"
$globalRoles = Invoke-Jira -Method GET -Url "$SiteUrl/rest/api/3/role" -Headers $headers
$roleNameById = @{}
foreach ($r in @($globalRoles)) {
  $idStr = [string](Get-SafeProp -Object $r -PropName "id")
  [int]$idInt = 0
  if ([int]::TryParse($idStr, [ref]$idInt)) {
    $roleNameById[$idInt] = [string](Get-SafeProp -Object $r -PropName "name")
  }
}

# Permission schemes cibles (2 IDs)
Write-Log "Recherche des permission schemes..." "INFO"
$defaultSchemeId  = Get-PermissionSchemeIdByName -SiteUrl $SiteUrl -Headers $headers -SchemeName $DefaultPermissionSchemeName
$archivedSchemeId = Get-PermissionSchemeIdByName -SiteUrl $SiteUrl -Headers $headers -SchemeName $ArchivedPermissionSchemeName

if ($null -eq $defaultSchemeId)  { throw "Permission scheme introuvable: '$DefaultPermissionSchemeName'." }
if ($null -eq $archivedSchemeId) { throw "Permission scheme introuvable: '$ArchivedPermissionSchemeName'." }

Write-Log ("Default scheme: '{0}' => id={1}" -f $DefaultPermissionSchemeName, $defaultSchemeId) "OK"
Write-Log ("Archived scheme: '{0}' => id={1}" -f $ArchivedPermissionSchemeName, $archivedSchemeId) "OK"

# Traitement projet par projet
foreach ($pk in $targetProjects) {
  Write-Log "---- Projet $pk ----" "INFO"

  # récupérer projet (nom)
  $proj = $null
  try {
    $proj = Invoke-Jira -Method GET -Url "$SiteUrl/rest/api/3/project/$pk" -Headers $headers
  } catch {
    Write-Log "Projet introuvable / inaccessible: $pk : $($_.Exception.Message)" "ERROR"
    Add-Action -ProjectKey $pk -ProjectName "" -ActionType "PermissionScheme" -Action "Error" -Mode $modeTxt -Result "KO" -Details "Project not found/inaccessible"
    continue
  }

  $projectName = [string](Get-SafeProp -Object $proj -PropName "name")

  # Déterminer le scheme cible en fonction du NOM
  $isArchived = $false
  if (($projectName.Trim()) -match '^(?i)z_') { $isArchived = $true }

  $targetSchemeName = $DefaultPermissionSchemeName
  $targetSchemeId   = $defaultSchemeId
  if ($isArchived) {
    $targetSchemeName = $ArchivedPermissionSchemeName
    $targetSchemeId   = $archivedSchemeId
  }

  # 1) Appliquer permission scheme (ciblé)
  try {
    $current = Get-ProjectPermissionScheme -SiteUrl $SiteUrl -Headers $headers -ProjectKey $pk
    $currentName = [string](Get-SafeProp -Object $current -PropName "name")

    if ($currentName -and ($currentName.Trim().ToLowerInvariant() -eq $targetSchemeName.Trim().ToLowerInvariant())) {
      Write-Log "Permission scheme déjà OK: '$currentName' ($pk)" "OK"
      Add-Action -ProjectKey $pk -ProjectName $projectName -ActionType "PermissionScheme" -Action "Skip" -Mode $modeTxt -Result "OK" `
        -Details ("Already on target scheme ({0})" -f $targetSchemeName)
    } else {
      if (-not $ApplyMode) {
        Write-Log "Appliquerait permission scheme: '$targetSchemeName' (id=$targetSchemeId) sur $pk (actuel='$currentName')" "DRYRUN"
        Add-Action -ProjectKey $pk -ProjectName $projectName -ActionType "PermissionScheme" -Action "SetScheme" -Mode $modeTxt -Result "OK" `
          -Details ("Would set scheme to '{0}' (archived={1})" -f $targetSchemeName, $isArchived)
      } else {
        Set-ProjectPermissionScheme -SiteUrl $SiteUrl -Headers $headers -ProjectKey $pk -SchemeId $targetSchemeId | Out-Null
        Write-Log "Permission scheme appliqué: '$targetSchemeName' sur $pk (ancien='$currentName')" "OK"
        Add-Action -ProjectKey $pk -ProjectName $projectName -ActionType "PermissionScheme" -Action "SetScheme" -Mode $modeTxt -Result "OK" `
          -Details ("Scheme set to '{0}' (archived={1})" -f $targetSchemeName, $isArchived)
      }
    }
  } catch {
    Write-Log "KO application permission scheme sur $pk : $($_.Exception.Message)" "ERROR"
    Add-Action -ProjectKey $pk -ProjectName $projectName -ActionType "PermissionScheme" -Action "SetScheme" -Mode $modeTxt -Result "KO" -Details $_.Exception.Message
    # Sécurité: on ne migre pas les rôles si le scheme n'a pas pu être appliqué
    continue
  }

  # 2) Migration rôles sur ce projet
  $roles = $null
  try {
    $roles = Invoke-Jira -Method GET -Url "$SiteUrl/rest/api/3/project/$pk/role" -Headers $headers
  } catch {
    Write-Log "KO lecture rôles projet $pk : $($_.Exception.Message)" "ERROR"
    Add-Action -ProjectKey $pk -ProjectName $projectName -ActionType "RoleMigration" -Action "Error" -Mode $modeTxt -Result "KO" -Details "Cannot read project roles"
    continue
  }

  # index id -> {Name, Url}
  $roleIndexById = @{}
  foreach ($prop in $roles.PSObject.Properties) {
    $rName = $prop.Name
    $rUrl  = [string]$prop.Value
    $rid = Get-RoleIdFromRoleUrl $rUrl
    if ($null -ne $rid) {
      $roleIndexById[[int]$rid] = [pscustomobject]@{ Name = $rName; Url = $rUrl }
    }
  }

  foreach ($map in $mappings) {
    $fromId = [int]$map.SourceId
    $toId   = [int]$map.TargetId

    $fromName = if ($roleIndexById.ContainsKey($fromId)) { $roleIndexById[$fromId].Name } else { ($roleNameById[$fromId]) }
    $toName   = if ($roleIndexById.ContainsKey($toId))   { $roleIndexById[$toId].Name }   else { ($roleNameById[$toId]) }

    if (-not $roleIndexById.ContainsKey($fromId)) {
      Add-Action -ProjectKey $pk -ProjectName $projectName -ActionType "RoleMigration" -Action "Skip" -Mode $modeTxt -Result "OK" `
        -Details "Source role not on project" -FromRoleId $fromId -FromRoleName $fromName -ToRoleId $toId -ToRoleName $toName
      continue
    }
    if (-not $roleIndexById.ContainsKey($toId)) {
      Add-Action -ProjectKey $pk -ProjectName $projectName -ActionType "RoleMigration" -Action "Error" -Mode $modeTxt -Result "KO" `
        -Details "Target role not on project" -FromRoleId $fromId -FromRoleName $fromName -ToRoleId $toId -ToRoleName $toName
      continue
    }

    $fromUrl = $roleIndexById[$fromId].Url
    $toUrl   = $roleIndexById[$toId].Url

    $fromDetail = $null
    $toDetail   = $null
    try { $fromDetail = Invoke-Jira -Method GET -Url $fromUrl -Headers $headers } catch {
      Add-Action -ProjectKey $pk -ProjectName $projectName -ActionType "RoleMigration" -Action "Error" -Mode $modeTxt -Result "KO" `
        -Details "Cannot read source role detail" -FromRoleId $fromId -FromRoleName $fromName -ToRoleId $toId -ToRoleName $toName
      continue
    }
    try { $toDetail = Invoke-Jira -Method GET -Url $toUrl -Headers $headers } catch {
      Add-Action -ProjectKey $pk -ProjectName $projectName -ActionType "RoleMigration" -Action "Error" -Mode $modeTxt -Result "KO" `
        -Details "Cannot read target role detail" -FromRoleId $fromId -FromRoleName $fromName -ToRoleId $toId -ToRoleName $toName
      continue
    }

    $fromActors = Get-RoleActors -RoleDetail $fromDetail
    $toActors   = Get-RoleActors -RoleDetail $toDetail

    if (($fromActors.Groups.Count + $fromActors.Users.Count) -eq 0) {
      Add-Action -ProjectKey $pk -ProjectName $projectName -ActionType "RoleMigration" -Action "Skip" -Mode $modeTxt -Result "OK" `
        -Details "Source role empty" -FromRoleId $fromId -FromRoleName $fromName -ToRoleId $toId -ToRoleName $toName
      continue
    }

    $groupsToAdd = @()
    foreach ($g in $fromActors.Groups) { if (-not $toActors.Groups.Contains($g)) { $groupsToAdd += $g } }

    $usersToAdd = @()
    foreach ($u in $fromActors.Users) { if (-not $toActors.Users.Contains($u)) { $usersToAdd += $u } }

    # Add to target
    if ($groupsToAdd.Count -gt 0 -or $usersToAdd.Count -gt 0) {
      if (-not $ApplyMode) {
        foreach ($g in $groupsToAdd) {
          Add-Action -ProjectKey $pk -ProjectName $projectName -ActionType "RoleMigration" -Action "AddToTarget" -Mode $modeTxt -Result "OK" -Details "Would add" `
            -FromRoleId $fromId -FromRoleName $fromName -ToRoleId $toId -ToRoleName $toName -ActorType "Group" -ActorName $g
        }
        foreach ($u in $usersToAdd) {
          Add-Action -ProjectKey $pk -ProjectName $projectName -ActionType "RoleMigration" -Action "AddToTarget" -Mode $modeTxt -Result "OK" -Details "Would add" `
            -FromRoleId $fromId -FromRoleName $fromName -ToRoleId $toId -ToRoleName $toName -ActorType "User" -AccountId $u
        }
      } else {
        try {
          Add-ActorsToRole -RoleUrl $toUrl -Groups $groupsToAdd -Users $usersToAdd -Headers $headers | Out-Null
          foreach ($g in $groupsToAdd) {
            Add-Action -ProjectKey $pk -ProjectName $projectName -ActionType "RoleMigration" -Action "AddToTarget" -Mode $modeTxt -Result "OK" -Details "Added" `
              -FromRoleId $fromId -FromRoleName $fromName -ToRoleId $toId -ToRoleName $toName -ActorType "Group" -ActorName $g
          }
          foreach ($u in $usersToAdd) {
            Add-Action -ProjectKey $pk -ProjectName $projectName -ActionType "RoleMigration" -Action "AddToTarget" -Mode $modeTxt -Result "OK" -Details "Added" `
              -FromRoleId $fromId -FromRoleName $fromName -ToRoleId $toId -ToRoleName $toName -ActorType "User" -AccountId $u
          }
        } catch {
          Add-Action -ProjectKey $pk -ProjectName $projectName -ActionType "RoleMigration" -Action "AddToTarget" -Mode $modeTxt -Result "KO" -Details $_.Exception.Message `
            -FromRoleId $fromId -FromRoleName $fromName -ToRoleId $toId -ToRoleName $toName
          continue
        }
      }
    }

    # Remove from source
    if (-not $ApplyMode) {
      foreach ($g in $fromActors.Groups) {
        Add-Action -ProjectKey $pk -ProjectName $projectName -ActionType "RoleMigration" -Action "RemoveFromSource" -Mode $modeTxt -Result "OK" -Details "Would remove" `
          -FromRoleId $fromId -FromRoleName $fromName -ToRoleId $toId -ToRoleName $toName -ActorType "Group" -ActorName $g
      }
      foreach ($u in $fromActors.Users) {
        Add-Action -ProjectKey $pk -ProjectName $projectName -ActionType "RoleMigration" -Action "RemoveFromSource" -Mode $modeTxt -Result "OK" -Details "Would remove" `
          -FromRoleId $fromId -FromRoleName $fromName -ToRoleId $toId -ToRoleName $toName -ActorType "User" -AccountId $u
      }
    } else {
      foreach ($g in $fromActors.Groups) {
        try {
          $delUrl = Remove-GroupFromRole -RoleUrl $fromUrl -GroupName $g -Headers $headers
          Add-Action -ProjectKey $pk -ProjectName $projectName -ActionType "RoleMigration" -Action "RemoveFromSource" -Mode $modeTxt -Result "OK" -Details $delUrl `
            -FromRoleId $fromId -FromRoleName $fromName -ToRoleId $toId -ToRoleName $toName -ActorType "Group" -ActorName $g
        } catch {
          Add-Action -ProjectKey $pk -ProjectName $projectName -ActionType "RoleMigration" -Action "RemoveFromSource" -Mode $modeTxt -Result "KO" -Details $_.Exception.Message `
            -FromRoleId $fromId -FromRoleName $fromName -ToRoleId $toId -ToRoleName $toName -ActorType "Group" -ActorName $g
        }
      }
      foreach ($u in $fromActors.Users) {
        try {
          $delUrl = Remove-UserFromRole -RoleUrl $fromUrl -AccountId $u -Headers $headers
          Add-Action -ProjectKey $pk -ProjectName $projectName -ActionType "RoleMigration" -Action "RemoveFromSource" -Mode $modeTxt -Result "OK" -Details $delUrl `
            -FromRoleId $fromId -FromRoleName $fromName -ToRoleId $toId -ToRoleName $toName -ActorType "User" -AccountId $u
        } catch {
          Add-Action -ProjectKey $pk -ProjectName $projectName -ActionType "RoleMigration" -Action "RemoveFromSource" -Mode $modeTxt -Result "KO" -Details $_.Exception.Message `
            -FromRoleId $fromId -FromRoleName $fromName -ToRoleId $toId -ToRoleName $toName -ActorType "User" -AccountId $u
        }
      }
    }
  }
}

# Export actions
$global:Actions | Export-Csv -Path $global:ActionsCsv -Delimiter ';' -NoTypeInformation -Encoding UTF8
Write-Log ("Export actions: {0} lignes -> {1}" -f $global:Actions.Count, $global:ActionsCsv) "OK"
Write-Log "Terminé." "OK"
Write-Log "Log sauvegardé: $global:LogFile" "OK"