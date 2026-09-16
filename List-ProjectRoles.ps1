<# ---------------------------------------------------------------------------
List-ProjectRoles.ps1

But:
- Lister tous les rôles de projet définis dans Jira (globaux),
  avec:
    - Id
    - Name
    - Description

Sorties:
- Log: .\logs\List-ProjectRoles_yyyyMMdd_HHmmss.log
- Export CSV: .\exports\ProjectRoles_yyyyMMdd_HHmmss.csv

API utilisée:
- GET /rest/api/3/role
--------------------------------------------------------------------------- #>

[CmdletBinding()]
param(
  [string] $SiteUrl   = "https://jiradot.atlassian.net",
  [string] $UserEmail = "frederic.guedj@harmonie-mutuelle.fr",

  [string] $CredentialFile = $null,
  [switch] $ResetCredential
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
  $log = Join-Path $logsDir ("List-ProjectRoles_{0}.log" -f $ts)
  $csv = Join-Path $expDir  ("ProjectRoles_{0}.csv" -f $ts)

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

# ------------------------- Main -------------------------
Write-Log "SiteUrl: $SiteUrl" "INFO"
Write-Log "CredentialFile: $CredentialFile" "INFO"

$cred = Get-JiraCredential -CredentialFile $CredentialFile -UserEmail $UserEmail -Reset:$ResetCredential
$headers = Get-JiraAuthHeader -Credential $cred

$me = Invoke-Jira -Method GET -Url "$SiteUrl/rest/api/3/myself" -Headers $headers
Write-Log ("Authentifié: {0}" -f [string](Get-SafeProp -Object $me -PropName "displayName")) "OK"

Write-Log "Récupération des rôles via /rest/api/3/role ..." "INFO"

$roles = Invoke-Jira -Method GET -Url "$SiteUrl/rest/api/3/role" -Headers $headers

if (-not $roles -or $roles.Count -eq 0) {
  Write-Log "Aucun rôle retourné par l'API." "WARN"
}

$results = New-Object System.Collections.Generic.List[object]

foreach ($r in $roles) {
  $id   = [string](Get-SafeProp -Object $r -PropName "id")
  $name = [string](Get-SafeProp -Object $r -PropName "name")

  # description n'est pas toujours renvoyée selon les clients / versions;
  # sinon => rappelera GET /role/{id} si besoin dans une V2.
  $desc = [string](Get-SafeProp -Object $r -PropName "description")

  $results.Add([pscustomobject]@{
    Id          = $id
    Name        = $name
    Description = $desc
  })
}

$results | Sort-Object Name | Export-Csv -Path $global:ExportCsv -Delimiter ';' -NoTypeInformation -Encoding UTF8

Write-Log ("Export CSV terminé: {0} rôles -> {1}" -f $results.Count, $global:ExportCsv) "OK"
Write-Log "Terminé." "OK"
Write-Log "Log sauvegardé: $global:LogFile" "OK"