<#
.SYNOPSIS
  Diag-Projet-Jira.ps1
  Diagnostic rapide d'acces a un projet Jira via API
#>

param(
  [string] $ProjectKey = "CPT"
)

$ScriptDir  = Split-Path -Parent $PSCommandPath
$SecretsDir = Join-Path $ScriptDir "secrets"

# --- Credentials ---
$data  = Import-Clixml -Path (Join-Path $SecretsDir "site-admin.xml")
$url   = [string]$data.SiteUrl
$email = [string]$data.Email
$token = [System.Net.NetworkCredential]::new("", $data.ApiTokenSecureString).Password
$auth  = [Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes("${email}:${token}"))
$headers = @{ Authorization = "Basic $auth"; Accept = "application/json" }
$baseUrl = "https://$url"

try { [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12 } catch {}
try {
  $proxy = [System.Net.WebRequest]::GetSystemWebProxy()
  $proxy.Credentials = [System.Net.CredentialCache]::DefaultCredentials
  [System.Net.WebRequest]::DefaultWebProxy = $proxy
} catch {}

function Api([string]$path) {
  try {
    $resp = Invoke-WebRequest -Uri "$baseUrl$path" -Headers $headers -UseBasicParsing -ErrorAction Stop
    return @{ ok=$true; status=[int]$resp.StatusCode; content=$resp.Content }
  } catch {
    $status = 0
    try { $status = [int]$_.Exception.Response.StatusCode } catch {}
    return @{ ok=$false; status=$status; error=$_.Exception.Message }
  }
}

Write-Host ""
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "  DIAGNOSTIC PROJET JIRA : $ProjectKey" -ForegroundColor Cyan
Write-Host "  Site : $baseUrl" -ForegroundColor Cyan
Write-Host "  Email : $email" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

# --- Test 1 : Qui suis-je ? ---
Write-Host "1. IDENTITE DU COMPTE API" -ForegroundColor Yellow
$r = Api "/rest/api/3/myself"
if ($r.ok) {
  $me = $r.content | ConvertFrom-Json
  Write-Host ("   Nom          : {0}" -f $me.displayName) -ForegroundColor Green
  Write-Host ("   Email        : {0}" -f $me.emailAddress) -ForegroundColor Green
  Write-Host ("   AccountId    : {0}" -f $me.accountId) -ForegroundColor DarkGray
  Write-Host ("   Active       : {0}" -f $me.active) -ForegroundColor Green
  $myAccountId = $me.accountId
} else {
  Write-Host ("   ERREUR : status={0} {1}" -f $r.status, $r.error) -ForegroundColor Red
  Write-Host "   => Verifiez les credentials dans site-admin.xml" -ForegroundColor Red
  return
}

# --- Test 2 : Le projet existe ? ---
Write-Host ""
Write-Host "2. PROJET $ProjectKey" -ForegroundColor Yellow
$r = Api "/rest/api/3/project/$ProjectKey"
if ($r.ok) {
  $proj = $r.content | ConvertFrom-Json
  Write-Host ("   Nom          : {0}" -f $proj.name) -ForegroundColor Green
  Write-Host ("   Key          : {0}" -f $proj.key) -ForegroundColor Green
  Write-Host ("   Type         : {0}" -f $proj.projectTypeKey) -ForegroundColor Green
  Write-Host ("   Style        : {0}" -f $proj.style) -ForegroundColor Green
  $isNextGen = ($proj.style -eq "next-gen" -or $proj.simplified -eq $true)
  if ($isNextGen) {
    Write-Host "   ATTENTION    : Projet Team-managed (next-gen)" -ForegroundColor DarkYellow
  } else {
    Write-Host "   Type gestion : Company-managed (classic)" -ForegroundColor Green
  }
  if ($proj.lead) {
    Write-Host ("   Lead         : {0}" -f $proj.lead.displayName) -ForegroundColor DarkGray
  }
} else {
  Write-Host ("   ERREUR : status={0}" -f $r.status) -ForegroundColor Red
  if ($r.status -eq 404) { Write-Host "   => Le projet n'existe pas ou est invisible pour ce compte" -ForegroundColor Red }
  return
}

# --- Test 3 : Permissions sur le projet ---
Write-Host ""
Write-Host "3. PERMISSIONS SUR LE PROJET" -ForegroundColor Yellow
$permsToCheck = @("BROWSE_PROJECTS", "VIEW_ISSUES", "SEARCH_ISSUES", "ADMINISTER_PROJECTS")
foreach ($perm in $permsToCheck) {
  $r = Api "/rest/api/3/mypermissions?projectKey=$ProjectKey&permissions=$perm"
  if ($r.ok) {
    $permJson = $r.content | ConvertFrom-Json
    $hasPerm = $permJson.permissions.$perm.havePermission
    $color = if ($hasPerm) { "Green" } else { "Red" }
    $symbol = if ($hasPerm) { "OK" } else { "NON" }
    Write-Host ("   {0,-25} : {1}" -f $perm, $symbol) -ForegroundColor $color
  } else {
    Write-Host ("   {0,-25} : ERREUR (status={1})" -f $perm, $r.status) -ForegroundColor Red
  }
}

# --- Test 4 : Roles du compte sur le projet ---
Write-Host ""
Write-Host "4. ROLES SUR LE PROJET" -ForegroundColor Yellow
$r = Api "/rest/api/3/project/$ProjectKey/role"
if ($r.ok) {
  $roles = $r.content | ConvertFrom-Json
  foreach ($roleProp in $roles.PSObject.Properties) {
    $roleName = $roleProp.Name
    $roleUrl = [string]$roleProp.Value
    # Extraire l'ID du role depuis l'URL
    $roleId = ""
    if ($roleUrl -match "/role/(\d+)") { $roleId = $matches[1] }
    if ($roleId) {
      $rRole = Api "/rest/api/3/project/$ProjectKey/role/$roleId"
      if ($rRole.ok) {
        $roleData = $rRole.content | ConvertFrom-Json
        $actors = $roleData.actors
        $isMember = $false
        foreach ($actor in $actors) {
          if ($actor.actorUser -and $actor.actorUser.accountId -eq $myAccountId) {
            $isMember = $true
          }
        }
        $actorCount = ($actors | Measure-Object).Count
        $color = if ($isMember) { "Green" } else { "DarkGray" }
        $symbol = if ($isMember) { "<= MON COMPTE" } else { "" }
        Write-Host ("   {0,-25} : {1} membres {2}" -f $roleName, $actorCount, $symbol) -ForegroundColor $color
      }
    }
  }
} else {
  Write-Host ("   ERREUR : status={0}" -f $r.status) -ForegroundColor Red
}

# --- Test 5 : Scheme de securite ---
Write-Host ""
Write-Host "5. SECURITY SCHEME" -ForegroundColor Yellow
$r = Api "/rest/api/3/project/$ProjectKey/issuesecuritylevelscheme"
if ($r.ok) {
  $secScheme = $r.content | ConvertFrom-Json
  if ($secScheme.name) {
    Write-Host ("   Scheme       : {0}" -f $secScheme.name) -ForegroundColor DarkYellow
    Write-Host "   => Des tickets peuvent etre masques par un niveau de securite" -ForegroundColor DarkYellow
  } else {
    Write-Host "   Aucun scheme de securite" -ForegroundColor Green
  }
} else {
  if ($r.status -eq 404) {
    Write-Host "   Aucun scheme de securite" -ForegroundColor Green
  } else {
    Write-Host ("   ERREUR : status={0}" -f $r.status) -ForegroundColor DarkGray
  }
}

# --- Test 6 : JQL sans filtre ---
Write-Host ""
Write-Host "6. COMPTAGE JQL" -ForegroundColor Yellow

$jqlTests = @(
  @{ label="project = $ProjectKey"; jql="project=$ProjectKey" },
  @{ label="project = $ProjectKey (guillemets)"; jql="project=""$ProjectKey""" },
  @{ label="project = $ProjectKey (tous statuts)"; jql="project=$ProjectKey ORDER BY created DESC" },
  @{ label="project = $ProjectKey AND statusCategory in (""To Do"",""In Progress"",""Done"")"; jql="project=$ProjectKey AND statusCategory in (""To Do"",""In Progress"",""Done"")" }
)

foreach ($test in $jqlTests) {
  $encodedJql = [System.Uri]::EscapeDataString($test.jql)
  $r = Api "/rest/api/3/search?jql=$encodedJql&maxResults=0"
  if ($r.ok) {
    $searchJson = $r.content | ConvertFrom-Json
    $total = [int]$searchJson.total
    $color = if ($total -gt 0) { "Green" } else { "Red" }
    Write-Host ("   {0,-60} => {1} tickets" -f $test.label, $total) -ForegroundColor $color
  } else {
    Write-Host ("   {0,-60} => ERREUR status={1}" -f $test.label, $r.status) -ForegroundColor Red
  }
}

# --- Test 7 : Dernier ticket cree ---
Write-Host ""
Write-Host "7. DERNIER TICKET DU PROJET" -ForegroundColor Yellow
$encodedJql = [System.Uri]::EscapeDataString("project=$ProjectKey ORDER BY created DESC")
$r = Api "/rest/api/3/search?jql=$encodedJql&maxResults=1&fields=key,summary,status,created"
if ($r.ok) {
  $searchJson = $r.content | ConvertFrom-Json
  if ($searchJson.issues -and ($searchJson.issues | Measure-Object).Count -gt 0) {
    $lastIssue = $searchJson.issues[0]
    Write-Host ("   Key     : {0}" -f $lastIssue.key) -ForegroundColor Green
    Write-Host ("   Summary : {0}" -f $lastIssue.fields.summary) -ForegroundColor Green
    Write-Host ("   Status  : {0}" -f $lastIssue.fields.status.name) -ForegroundColor Green
    Write-Host ("   Created : {0}" -f $lastIssue.fields.created) -ForegroundColor Green
  } else {
    Write-Host "   Aucun ticket retourne" -ForegroundColor Red
  }
} else {
  Write-Host ("   ERREUR : status={0}" -f $r.status) -ForegroundColor Red
}

# --- Test 8 : Boards associes (pour les projets next-gen/Kanban) ---
Write-Host ""
Write-Host "8. BOARDS ASSOCIES" -ForegroundColor Yellow
$r = Api "/rest/agile/1.0/board?projectKeyOrId=$ProjectKey"
if ($r.ok) {
  $boardsJson = $r.content | ConvertFrom-Json
  $boardCount = ($boardsJson.values | Measure-Object).Count
  Write-Host ("   {0} board(s) trouve(s)" -f $boardCount) -ForegroundColor Green
  foreach ($board in $boardsJson.values) {
    Write-Host ("     - {0} (type: {1}, id: {2})" -f $board.name, $board.type, $board.id) -ForegroundColor DarkGray
  }
} else {
  Write-Host ("   ERREUR : status={0}" -f $r.status) -ForegroundColor DarkGray
}

# --- Test 9 : Filtres JQL du board (si le board a un filtre restrictif) ---
Write-Host ""
Write-Host "9. FILTRE DU BOARD" -ForegroundColor Yellow
if ($boardsJson -and $boardsJson.values) {
  foreach ($board in $boardsJson.values) {
    $r = Api "/rest/agile/1.0/board/$($board.id)/configuration"
    if ($r.ok) {
      $boardConfig = $r.content | ConvertFrom-Json
      if ($boardConfig.filter) {
        $filterId = $boardConfig.filter.id
        Write-Host ("   Board '{0}' utilise le filtre ID={1}" -f $board.name, $filterId) -ForegroundColor DarkGray
        $rFilter = Api "/rest/api/3/filter/$filterId"
        if ($rFilter.ok) {
          $filterJson = $rFilter.content | ConvertFrom-Json
          Write-Host ("   Filtre JQL : {0}" -f $filterJson.jql) -ForegroundColor DarkYellow
        }
      }
    }
  }
} else {
  Write-Host "   Pas de board a analyser" -ForegroundColor DarkGray
}

# --- Resume ---
Write-Host ""
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "  FIN DU DIAGNOSTIC" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "  Si tous les JQL retournent 0 tickets :" -ForegroundColor White
Write-Host "    - Le projet est peut-etre vide (aucun ticket cree)" -ForegroundColor DarkGray
Write-Host "    - Le compte API n'a pas BROWSE_PROJECTS" -ForegroundColor DarkGray
Write-Host "    - Un scheme de securite masque tous les tickets" -ForegroundColor DarkGray
Write-Host "    - Le projet est Team-managed avec des restrictions" -ForegroundColor DarkGray
Write-Host ""
Write-Host "  Collez la sortie ci-dessus pour analyse." -ForegroundColor White