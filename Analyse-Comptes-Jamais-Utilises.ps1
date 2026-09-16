<#
.SYNOPSIS
  Analyse-Comptes-Jamais-Utilises.ps1 v1.6
.DESCRIPTION
  Identifie les comptes jamais utilises (lastActive=null) sur Jiradot/mutexfr.
  Enrichissement Assets via gateway/api pour les comptes DSIM/DSIT.
  Colonnes Assets : Statut, DateEntree, DateSortie, DatePLD, MotifSortie
.NOTES
  Auteur  : Frederic GUEDJ
  Version : 1.6 — Aout 2026
#>

[CmdletBinding()]
param(
  [int] $SeuilMois = 2,
  [int] $MaxRetries = 5
)

Add-Type -AssemblyName System.Windows.Forms | Out-Null

$ScriptDir  = Split-Path -Parent $PSCommandPath
$SecretsDir = Join-Path $ScriptDir "secrets"
$ExportsDir = Join-Path $ScriptDir "exports"
if (-not (Test-Path $ExportsDir)) { New-Item -ItemType Directory -Path $ExportsDir -Force | Out-Null }

$ts = Get-Date -Format "yyyyMMdd_HHmmss"
$logFile      = Join-Path $ExportsDir ("ComptesJamaisUtilises_{0}.log" -f $ts)
$csvConsolide = Join-Path $ExportsDir ("ComptesJamaisUtilises_TOUS_{0}.csv" -f $ts)
$csvJiradot   = Join-Path $ExportsDir ("ComptesJamaisUtilises_Jiradot_{0}.csv" -f $ts)
$csvMutexfr   = Join-Path $ExportsDir ("ComptesJamaisUtilises_mutexfr_{0}.csv" -f $ts)
$csvRolesFile = Join-Path $ExportsDir ("GroupesProduit-ApplicationRoles_{0}.csv" -f $ts)
"" | Set-Content -Path $logFile -Encoding UTF8

# ============================================================
# FONCTIONS UTILITAIRES
# ============================================================

function Log([string]$msg, [string]$level="INFO") {
  $line = "[{0}][{1}] {2}" -f (Get-Date).ToString("yyyy-MM-dd HH:mm:ss"), $level, $msg
  Write-Host $line
  Add-Content -Path $logFile -Value $line -Encoding UTF8
}

function Format-DateFr([string]$isoDate) {
  if (-not $isoDate) { return "" }
  try {
    $dt = [DateTimeOffset]::Parse($isoDate)
    return $dt.ToLocalTime().ToString("dd/MM/yyyy HH:mm")
  } catch { return $isoDate }
}

# ============================================================
# FILTRE DOMAINES EMAIL
# ============================================================

$allowedDomains = @(
    "mutex.fr", "mutex-exterieur.fr", "harmonie-mutuelle.fr",
    "prestataire.sihm.fr", "chorum.fr"
)

function Get-EmailDomain([string]$email) {
    if ([string]::IsNullOrWhiteSpace($email)) { return $null }
    $idx = $email.IndexOf("@")
    if ($idx -lt 0) { return $null }
    return $email.Substring($idx + 1).Trim().ToLower()
}

function Test-EmailDomainAllowed([string]$email) {
    if ([string]::IsNullOrWhiteSpace($email)) { return $true }
    $domain = Get-EmailDomain $email
    if (-not $domain) { return $true }
    foreach ($d in $script:allowedDomains) {
        if ($domain -eq $d.ToLower()) { return $true }
    }
    return $false
}

function Test-IsDsimGroup([string]$groupesStr) {
    if ([string]::IsNullOrWhiteSpace($groupesStr)) { return $false }
    $lower = $groupesStr.ToLower()
    return ($lower.Contains("dsim") -or $lower.Contains("dsit"))
}

# ============================================================
# RESEAU
# ============================================================

try { [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12 } catch {}
try {
  $proxy = [System.Net.WebRequest]::GetSystemWebProxy()
  $proxy.Credentials = [System.Net.CredentialCache]::DefaultCredentials
  [System.Net.WebRequest]::DefaultWebProxy = $proxy
} catch {}

# ============================================================
# HTTP HELPER
# ============================================================

function Invoke-ApiCall {
  param([string]$Method, [string]$Url, [hashtable]$Headers, [string]$Body = $null)
  $attempt = 0
  while ($true) {
    $attempt++
    try {
      $params = @{ Method=$Method; Uri=$Url; Headers=$Headers; UseBasicParsing=$true; ErrorAction="Stop" }
      if ($Body) {
        $params["ContentType"] = "application/json; charset=utf-8"
        $params["Body"] = [System.Text.Encoding]::UTF8.GetBytes($Body)
      }
      $resp = Invoke-WebRequest @params
      $contentUtf8 = $resp.Content
      try {
        $stream = $resp.RawContentStream
        if ($stream -and $stream.CanSeek) {
          $stream.Position = 0
          $reader = New-Object System.IO.StreamReader($stream, [System.Text.Encoding]::UTF8, $true)
          $contentUtf8 = $reader.ReadToEnd(); $reader.Close()
        }
      } catch {
        try {
          $isoBytes = [System.Text.Encoding]::GetEncoding("ISO-8859-1").GetBytes($resp.Content)
          $contentUtf8 = [System.Text.Encoding]::UTF8.GetString($isoBytes)
        } catch {}
      }
      return @{ ok=$true; status=[int]$resp.StatusCode; content=$contentUtf8 }
    } catch {
      $status = 0; $errBody = ""
      try {
        $status = [int]$_.Exception.Response.StatusCode
        $rd = New-Object System.IO.StreamReader($_.Exception.Response.GetResponseStream())
        $errBody = $rd.ReadToEnd(); $rd.Close()
      } catch {}
      if ($attempt -gt $MaxRetries) { return @{ ok=$false; status=$status; error=$_.Exception.Message; body=$errBody } }
      if ($status -eq 429 -or ($status -ge 500 -and $status -le 599) -or $status -eq 0) {
        $sleepSec = [Math]::Min(60, [Math]::Pow(2, [Math]::Min(5, $attempt)))
        Log ("Retry {0} status={1} in {2}s ({3}/{4})" -f $Method, $status, $sleepSec, $attempt, $MaxRetries) "WARN"
        Start-Sleep -Seconds $sleepSec; continue
      }
      return @{ ok=$false; status=$status; error=$_.Exception.Message; body=$errBody }
    }
  }
}

# ============================================================
# CREDENTIALS
# ============================================================

function Import-SiteCredentials([string]$FilePath, [string]$SiteName) {
  if (-not (Test-Path $FilePath)) {
    [System.Windows.Forms.MessageBox]::Show(
      ("Credentials pour {0} manquants." -f $SiteName), "Credentials $SiteName",
      [System.Windows.Forms.MessageBoxButtons]::OK,
      [System.Windows.Forms.MessageBoxIcon]::Information) | Out-Null
    $inputUrl = Read-Host "URL du site"
    $inputUrl = $inputUrl -replace "^https?://", "" -replace "/.*$", "" -replace "/$", ""
    $adminEmail = Read-Host "Email administrateur"
    $apiTokenSecure = Read-Host "API Token" -AsSecureString
    @{ SiteUrl=$inputUrl; Email=$adminEmail; ApiTokenSecureString=$apiTokenSecure } | Export-Clixml -Path $FilePath
  }
  $data = Import-Clixml -Path $FilePath
  $url   = [string]$data.SiteUrl
  $email = [string]$data.Email
  $token = [System.Net.NetworkCredential]::new("", $data.ApiTokenSecureString).Password
  $auth  = [Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes("${email}:${token}"))
  return @{
    BaseUrl = "https://$url"
    Headers = @{ Authorization = "Basic $auth"; Accept = "application/json" }
    Name    = $SiteName
  }
}

function Get-SiteUsers([hashtable]$Site) {
  $usersMap = @{}
  $startAt = 0; $pageSize = 200
  while ($true) {
    $url = "{0}/rest/api/3/users/search?startAt={1}&maxResults={2}" -f $Site.BaseUrl, $startAt, $pageSize
    $resp = Invoke-ApiCall -Method "GET" -Url $url -Headers $Site.Headers
    if (-not $resp.ok) { Log ("Erreur listing users {0}" -f $Site.Name) "ERROR"; break }
    $json = $resp.content | ConvertFrom-Json
    $count = ($json | Measure-Object).Count
    if ($count -eq 0) { break }
    foreach ($u in $json) {
      $accId = [string]$u.accountId
      $accType = ""; if ($u.accountType) { $accType = [string]$u.accountType }
      if ($accType -eq "app") { continue }
      if (-not $usersMap.ContainsKey($accId)) {
        $usersMap[$accId] = @{ displayName=[string]$u.displayName; active=[bool]$u.active; accountType=$accType }
      }
    }
    if ($count -lt $pageSize) { break }
    $startAt += $pageSize
    Start-Sleep -Milliseconds 200
  }
  Log ("  {0} : {1} utilisateurs charges" -f $Site.Name, $usersMap.Count)
  return $usersMap
}

# ============================================================
# GROUPES PRODUIT (applicationrole)
# ============================================================

function Get-ApplicationRoles([hashtable]$Site) {
    $roles = New-Object System.Collections.Generic.List[object]
    $url = "{0}/rest/api/3/applicationrole" -f $Site.BaseUrl
    $resp = Invoke-ApiCall -Method "GET" -Url $url -Headers $Site.Headers
    if (-not $resp.ok) { Log ("WARN : applicationrole indisponible sur {0}" -f $Site.Name) "WARN"; return $roles }
    $json = $resp.content | ConvertFrom-Json
    foreach ($role in $json) {
        $roleKey  = ""; if ($role.key)  { $roleKey  = [string]$role.key }
        $roleName = ""; if ($role.name) { $roleName = [string]$role.name }
        foreach ($g in @($role.groups) + @($role.defaultGroups)) {
            if (-not $g) { continue }
            $gName = [string]$g
            if ([string]::IsNullOrWhiteSpace($gName)) { continue }
            $isDefault = $false
            if ($role.defaultGroups) { $isDefault = ($role.defaultGroups | ForEach-Object { [string]$_ }) -contains $gName }
            $exists = $roles | Where-Object { $_.GroupName -eq $gName -and $_.Site -eq $Site.Name -and $_.RoleKey -eq $roleKey }
            if (-not $exists) {
                $roles.Add(@{ RoleKey=$roleKey; RoleName=$roleName; GroupName=$gName; IsDefault=$isDefault; Site=$Site.Name }) | Out-Null
            }
        }
    }
    return $roles
}

function Test-ConfluenceAccess([hashtable]$Site, [string]$AccountId) {
    $url = "{0}/wiki/rest/api/user?accountId={1}" -f $Site.BaseUrl, $AccountId
    $resp = Invoke-ApiCall -Method "GET" -Url $url -Headers $Site.Headers
    return $resp.ok
}

# ============================================================
# JIRA ASSETS v1.6 — endpoint gateway/api + champs modele
# ============================================================

function Get-AssetsWorkspaceId([hashtable]$Site) {
    $url = "{0}/rest/servicedeskapi/assets/workspace" -f $Site.BaseUrl
    $resp = Invoke-ApiCall -Method "GET" -Url $url -Headers $Site.Headers
    if ($resp.ok) {
        $json = $resp.content | ConvertFrom-Json
        if ($json.values -and $json.values.Count -gt 0) { return [string]$json.values[0].workspaceId }
    }
    $url2 = "{0}/rest/assets/1.0/workspace" -f $Site.BaseUrl
    $resp2 = Invoke-ApiCall -Method "GET" -Url $url2 -Headers $Site.Headers
    if ($resp2.ok) {
        $json2 = $resp2.content | ConvertFrom-Json
        if ($json2.values -and $json2.values.Count -gt 0) { return [string]$json2.values[0].workspaceId }
    }
    return $null
}

function ConvertFrom-AssetResponse([string]$Content) {
    $json = $Content | ConvertFrom-Json
    $objects = $null
    if ($json.values)            { $objects = $json.values }
    elseif ($json.objectEntries) { $objects = $json.objectEntries }
    elseif ($json -is [array])   { $objects = $json }
    if (-not $objects -or ($objects | Measure-Object).Count -eq 0) { return $null }

    $asset       = $objects[0]
    $statut      = "Trouve (statut inconnu)"
    $dateEntree  = ""
    $dateSortie  = ""
    $dateStatut  = ""
    $motifSortie = ""
    $typeAsset   = ""

    if ($asset.objectType -and $asset.objectType.name) {
        $typeAsset = [string]$asset.objectType.name
    }

    if ($asset.attributes) {
        foreach ($attr in $asset.attributes) {
            $attrName = ""
            if ($attr.objectTypeAttribute -and $attr.objectTypeAttribute.name) {
                $attrName = [string]$attr.objectTypeAttribute.name
            }
            $val0 = $null
            if ($attr.objectAttributeValues -and $attr.objectAttributeValues.Count -gt 0) {
                $val0 = $attr.objectAttributeValues[0]
            }

            switch ($attrName) {
                "Statut" {
                    if ($val0) {
                        if ($val0.displayValue)                      { $statut = [string]$val0.displayValue }
                        elseif ($val0.value)                         { $statut = [string]$val0.value }
                        elseif ($val0.status -and $val0.status.name) { $statut = [string]$val0.status.name }
                    }
                }
                # Nom exact modele : "Date Entrée"
                { $_ -in @("Date Entrée","Date Entree","Date entrée","Date entree","DateEntree") } {
                    if ($val0 -and $val0.displayValue) { $dateEntree = Format-DateFr ([string]$val0.displayValue) }
                    elseif ($val0 -and $val0.value)    { $dateEntree = Format-DateFr ([string]$val0.value) }
                }
                # Nom exact modele : "Date Sortie"
                { $_ -in @("Date Sortie","Date sortie","DateSortie","Date de sortie") } {
                    if ($val0 -and $val0.displayValue) { $dateSortie = Format-DateFr ([string]$val0.displayValue) }
                    elseif ($val0 -and $val0.value)    { $dateSortie = Format-DateFr ([string]$val0.value) }
                }
                # Nom exact modele : "Date PLD" (date changement statut)
                { $_ -in @("Date PLD","DatePLD","Date pld") } {
                    if ($val0 -and $val0.displayValue) { $dateStatut = Format-DateFr ([string]$val0.displayValue) }
                    elseif ($val0 -and $val0.value)    { $dateStatut = Format-DateFr ([string]$val0.value) }
                }
                # Nom exact modele : "Motif Sortie"
                { $_ -in @("Motif Sortie","Motif sortie","MotifSortie") } {
                    if ($val0) {
                        if ($val0.displayValue) { $motifSortie = [string]$val0.displayValue }
                        elseif ($val0.value)    { $motifSortie = [string]$val0.value }
                    }
                }
            }
        }
    }

    return @{
        Statut      = if ($typeAsset) { "$statut ($typeAsset)" } else { $statut }
        DateEntree  = $dateEntree
        DateSortie  = $dateSortie
        DateStatut  = $dateStatut
        MotifSortie = $motifSortie
    }
}

function Search-AssetByIdentity([hashtable]$Site, [string]$WorkspaceId, [string]$DisplayName, [string]$Email) {
    <#
    v1.7 : AQL pour trouver l ID -> GET /object/{id} pour les attributs
    L AQL ne retourne pas les attributs via gateway — il faut un GET separe.
    #>
    if ([string]::IsNullOrWhiteSpace($WorkspaceId)) { return $null }

    $aqlUrl = "{0}/gateway/api/jsm/assets/workspace/{1}/v1/object/aql" -f $Site.BaseUrl, $WorkspaceId
    $objUrl = "{0}/gateway/api/jsm/assets/workspace/{1}/v1/object" -f $Site.BaseUrl, $WorkspaceId

    # Fonction interne : GET attributs par ID
    function Get-AssetById([string]$ObjId) {
        $getUrl = "{0}/{1}" -f $objUrl, $ObjId
        $resp = Invoke-ApiCall -Method "GET" -Url $getUrl -Headers $Site.Headers
        if ($resp.ok) {
            # Wrapper dans values[] pour ConvertFrom-AssetResponse
            $obj = $resp.content | ConvertFrom-Json
            $wrapped = ConvertTo-Json @{ values = @($obj) } -Depth 20 -Compress
            return ConvertFrom-AssetResponse $wrapped
        }
        return $null
    }

    # Fonction interne : extraire l ID depuis une reponse AQL
    function Get-IdFromAqlResponse([string]$Content) {
        try {
            $json = $Content | ConvertFrom-Json
            $objects = $null
            if ($json.values)            { $objects = $json.values }
            elseif ($json.objectEntries) { $objects = $json.objectEntries }
            elseif ($json -is [array])   { $objects = $json }
            if ($objects -and ($objects | Measure-Object).Count -gt 0) {
                $obj = $objects[0]
                if ($obj.id) { return [string]$obj.id }
            }
        } catch {}
        return $null
    }

    # Fonction interne : AQL -> ID -> attributs
    function Try-AqlThenGet([string]$Aql, [string]$Label) {
        $body = @{ qlQuery=$Aql; maxResults=10 } | ConvertTo-Json -Depth 5 -Compress
        $resp = Invoke-ApiCall -Method "POST" -Url $aqlUrl -Headers $Site.Headers -Body $body
        if (-not $resp.ok) { return $null }

        # Cas 1 : un seul resultat -> GET direct
        $objId = Get-IdFromAqlResponse $resp.content
        if ($objId) {
            Log ("    {0} -> id={1}" -f $Label, $objId) "DEBUG"
            return Get-AssetById $objId
        }

        # Cas 2 : plusieurs resultats -> filtrer par label PS
        if (-not [string]::IsNullOrWhiteSpace($DisplayName)) {
            try {
                $json = $resp.content | ConvertFrom-Json
                $objects = $null
                if ($json.values)            { $objects = $json.values }
                elseif ($json.objectEntries) { $objects = $json.objectEntries }
                if ($objects -and ($objects | Measure-Object).Count -gt 0) {
                    $parts = $DisplayName.Trim() -split '\s+'
                    foreach ($obj in $objects) {
                        $label = ""
                        if ($obj.label) { $label = [string]$obj.label }
                        elseif ($obj.name) { $label = [string]$obj.name }
                        $allMatch = $true
                        foreach ($p in $parts) {
                            if (-not $label.ToUpper().Contains($p.ToUpper())) { $allMatch=$false; break }
                        }
                        if ($allMatch -and $obj.id) {
                            Log ("    {0} (filtre PS) -> id={1} label={2}" -f $Label, $obj.id, $label) "DEBUG"
                            return Get-AssetById ([string]$obj.id)
                        }
                    }
                }
            } catch {}
        }
        return $null
    }

    # ================================================================
    # STRATEGIES DE RECHERCHE
    # ================================================================

    # --- Strategie 1 : Name exact ---
    if (-not [string]::IsNullOrWhiteSpace($DisplayName)) {
        $nameEsc = $DisplayName -replace '"','\"'
        $r = Try-AqlThenGet ('Name = "{0}"' -f $nameEsc) "strat1 Name="
        if ($r) { return $r }
    }

    # --- Strategie 2 : Name LIKE chaque mot (filtre PS sur tous les mots) ---
    if (-not [string]::IsNullOrWhiteSpace($DisplayName)) {
        $parts = $DisplayName.Trim() -split '\s+'
        if ($parts.Count -ge 2) {
            foreach ($mot in $parts) {
                $motEsc = $mot -replace '"','\"'
                $r = Try-AqlThenGet ('Name LIKE "{0}"' -f $motEsc) ("strat2 LIKE {0}" -f $mot)
                if ($r) { return $r }
            }
        }
    }

    # --- Strategie 3 : Nom + Prenom attributs separes (NOM majuscules) ---
    if (-not [string]::IsNullOrWhiteSpace($DisplayName)) {
        $parts = $DisplayName.Trim() -split '\s+'
        if ($parts.Count -ge 2) {
            $combos = @(
                @{ Nom=$parts[-1].ToUpper(); Prenom=($parts[0..($parts.Count-2)] -join " ") },
                @{ Nom=$parts[0].ToUpper();  Prenom=($parts[1..($parts.Count-1)] -join " ") }
            )
            foreach ($combo in $combos) {
                $nomEsc = $combo.Nom -replace '"','\"'; $prenomEsc = $combo.Prenom -replace '"','\"'
                foreach ($prenomAttr in @("Prénom","Prenom")) {
                    $r = Try-AqlThenGet ('"Nom" = "{0}" AND "{1}" = "{2}"' -f $nomEsc, $prenomAttr, $prenomEsc) "strat3 Nom+Prenom MAJ"
                    if ($r) { return $r }
                }
            }
        }
    }

    # --- Strategie 4 : Nom + Prenom casse originale ---
    if (-not [string]::IsNullOrWhiteSpace($DisplayName)) {
        $parts = $DisplayName.Trim() -split '\s+'
        if ($parts.Count -ge 2) {
            $combos = @(
                @{ Nom=$parts[-1]; Prenom=($parts[0..($parts.Count-2)] -join " ") },
                @{ Nom=$parts[0];  Prenom=($parts[1..($parts.Count-1)] -join " ") }
            )
            foreach ($combo in $combos) {
                $nomEsc = $combo.Nom -replace '"','\"'; $prenomEsc = $combo.Prenom -replace '"','\"'
                foreach ($prenomAttr in @("Prénom","Prenom")) {
                    $r = Try-AqlThenGet ('"Nom" = "{0}" AND "{1}" = "{2}"' -f $nomEsc, $prenomAttr, $prenomEsc) "strat4 Nom+Prenom orig"
                    if ($r) { return $r }
                }
            }
        }
    }

    # --- Strategie 5 : Email ---
    if (-not [string]::IsNullOrWhiteSpace($Email)) {
        $emailEsc = $Email -replace '"','\"'
        foreach ($attrName in @("Email","Mail","E-mail","Adresse mail")) {
            $r = Try-AqlThenGet ('"{0}" = "{1}"' -f $attrName, $emailEsc) ("strat5 Email {0}" -f $attrName)
            if ($r) { return $r }
        }
    }

    return $null
}

function ConvertFrom-AssetResponse([string]$Content) {
    <#
    v1.7 : parse l objet retourne par GET /object/{id}
    Structure : objet unique avec cle "attributes" (tableau)
    Champs modele : Statut, Date Entree, Date Sortie, Date PLD, Motif Sortie
    #>
    $json = $Content | ConvertFrom-Json

    # Accepter objet direct OU wrapper {values:[...]}
    $asset = $null
    if ($json.values -and ($json.values | Measure-Object).Count -gt 0) {
        $asset = $json.values[0]
    } elseif ($json.objectEntries -and ($json.objectEntries | Measure-Object).Count -gt 0) {
        $asset = $json.objectEntries[0]
    } elseif ($json.id) {
        # Objet direct (retourne par GET /object/{id})
        $asset = $json
    } elseif ($json -is [array] -and $json.Count -gt 0) {
        $asset = $json[0]
    }

    if (-not $asset) { return $null }

    $statut      = "Trouve (statut inconnu)"
    $dateEntree  = ""
    $dateSortie  = ""
    $dateStatut  = ""
    $motifSortie = ""
    $typeAsset   = ""

    if ($asset.objectType -and $asset.objectType.name) {
        $typeAsset = [string]$asset.objectType.name
    }

    # Les attributs peuvent etre dans $asset.attributes
    # ou directement dans le tableau racine (endpoint /attributes)
    $attrs = $null
    if ($asset.attributes) {
        $attrs = $asset.attributes
    }

    if ($attrs) {
        foreach ($attr in $attrs) {
            $attrName = ""
            if ($attr.objectTypeAttribute -and $attr.objectTypeAttribute.name) {
                $attrName = [string]$attr.objectTypeAttribute.name
            }

            $val0 = $null
            if ($attr.objectAttributeValues -and $attr.objectAttributeValues.Count -gt 0) {
                $val0 = $attr.objectAttributeValues[0]
            }

            switch ($attrName) {
                "Statut" {
                    if ($val0) {
                        if ($val0.displayValue)                      { $statut = [string]$val0.displayValue }
                        elseif ($val0.value)                         { $statut = [string]$val0.value }
                        elseif ($val0.status -and $val0.status.name) { $statut = [string]$val0.status.name }
                    }
                }
                { $_ -in @("Date Entrée","Date Entree","Date entrée","Date entree","DateEntree") } {
                    if ($val0 -and $val0.displayValue) { $dateEntree = Format-DateFr ([string]$val0.displayValue) }
                    elseif ($val0 -and $val0.value)    { $dateEntree = Format-DateFr ([string]$val0.value) }
                }
                { $_ -in @("Date Sortie","Date sortie","DateSortie") } {
                    if ($val0 -and $val0.displayValue) { $dateSortie = Format-DateFr ([string]$val0.displayValue) }
                    elseif ($val0 -and $val0.value)    { $dateSortie = Format-DateFr ([string]$val0.value) }
                }
                { $_ -in @("Date PLD","DatePLD","Date pld") } {
                    if ($val0 -and $val0.displayValue) { $dateStatut = Format-DateFr ([string]$val0.displayValue) }
                    elseif ($val0 -and $val0.value)    { $dateStatut = Format-DateFr ([string]$val0.value) }
                }
                { $_ -in @("Motif Sortie","Motif sortie","MotifSortie") } {
                    if ($val0) {
                        if ($val0.displayValue) { $motifSortie = [string]$val0.displayValue }
                        elseif ($val0.value)    { $motifSortie = [string]$val0.value }
                    }
                }
            }
        }
    }

    return @{
        Statut      = if ($typeAsset) { "$statut ($typeAsset)" } else { $statut }
        DateEntree  = $dateEntree
        DateSortie  = $dateSortie
        DateStatut  = $dateStatut
        MotifSortie = $motifSortie
    }
}
# ============================================================
# CONFIGURATION + BANNIERE
# ============================================================

$dateSeuil    = (Get-Date).AddMonths(-$SeuilMois)
$dateSeuilStr = $dateSeuil.ToString("dd/MM/yyyy")

Write-Host ""
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "  ANALYSE COMPTES JAMAIS UTILISES v1.6" -ForegroundColor Cyan
Write-Host ("  Seuil : > {0} mois (avant le {1})" -f $SeuilMois, $dateSeuilStr) -ForegroundColor Cyan
Write-Host "  Endpoint Assets : gateway/api/jsm/assets" -ForegroundColor Green
Write-Host "  Champs Assets : Statut, Date Entree, Date Sortie, Date PLD, Motif Sortie" -ForegroundColor Green
Write-Host "  Mode : DRY-RUN" -ForegroundColor Green
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

Log "================================================================"
Log ("  ANALYSE COMPTES JAMAIS UTILISES v1.6 — seuil {0} mois" -f $SeuilMois)
Log "================================================================"

# ============================================================
# ETAPE 1 : CREDENTIALS
# ============================================================

Log "=== ETAPE 1 : Chargement des credentials ==="

$orgCredFile = Join-Path $SecretsDir "org-admin.xml"
if (-not (Test-Path $orgCredFile)) {
    $inputOrgId   = Read-Host "OrgId"
    $apiKeySecure = Read-Host "API Key organisation" -AsSecureString
    @{ OrgId=$inputOrgId; ApiKeySecureString=$apiKeySecure } | Export-Clixml -Path $orgCredFile
}
$orgData   = Import-Clixml -Path $orgCredFile
$orgId     = [string]$orgData.OrgId
$orgApiKey = [System.Net.NetworkCredential]::new("", $orgData.ApiKeySecureString).Password
$orgHeaders = @{ Authorization="Bearer $orgApiKey"; Accept="application/json" }

$siteJiradot = Import-SiteCredentials -FilePath (Join-Path $SecretsDir "site-admin.xml")   -SiteName "Jiradot"
$siteMutexfr = Import-SiteCredentials -FilePath (Join-Path $SecretsDir "site-mutexfr.xml") -SiteName "mutexfr"

Log ("  OrgId   : {0}" -f $orgId)
Log ("  Jiradot : {0}" -f $siteJiradot.BaseUrl)
Log ("  mutexfr : {0}" -f $siteMutexfr.BaseUrl)

# ============================================================
# ETAPE 2 : PRE-CHARGEMENT UTILISATEURS
# ============================================================

Log "=== ETAPE 2 : Pre-chargement des utilisateurs ==="
$jiradotUsers = Get-SiteUsers -Site $siteJiradot
$mutexfrUsers = Get-SiteUsers -Site $siteMutexfr

# ============================================================
# ETAPE 2b : GROUPES PRODUIT
# ============================================================

Log "=== ETAPE 2b : Groupes Jira produit ==="
$rolesJiradot = Get-ApplicationRoles -Site $siteJiradot
$rolesMutexfr = Get-ApplicationRoles -Site $siteMutexfr

$allRoles = New-Object System.Collections.Generic.List[object]
foreach ($r in $rolesJiradot) { $allRoles.Add($r) | Out-Null }
foreach ($r in $rolesMutexfr) { $allRoles.Add($r) | Out-Null }

$productGroupsJiradot = New-Object System.Collections.Generic.HashSet[string]([System.StringComparer]::OrdinalIgnoreCase)
foreach ($r in $rolesJiradot) { [void]$productGroupsJiradot.Add($r.GroupName) }
$productGroupsMutexfr = New-Object System.Collections.Generic.HashSet[string]([System.StringComparer]::OrdinalIgnoreCase)
foreach ($r in $rolesMutexfr) { [void]$productGroupsMutexfr.Add($r.GroupName) }

Log ("  Jiradot : {0} groupes produit" -f $productGroupsJiradot.Count)
Log ("  mutexfr : {0} groupes produit" -f $productGroupsMutexfr.Count)

$rolesCsvColumns = @("Site","ApplicationRole","RoleKey","GroupName","IsDefault")
$utf8Bom = New-Object System.Text.UTF8Encoding($true)
$rolesSb = New-Object System.Text.StringBuilder
[void]$rolesSb.AppendLine(($rolesCsvColumns | ForEach-Object { '"{0}"' -f $_ }) -join ";")
foreach ($r in ($allRoles | Sort-Object { $_.Site + $_.RoleName + $_.GroupName })) {
    $row = [ordered]@{ Site=$r.Site; ApplicationRole=$r.RoleName; RoleKey=$r.RoleKey; GroupName=$r.GroupName; IsDefault=if($r.IsDefault){"OUI"}else{"NON"} }
    [void]$rolesSb.AppendLine(($rolesCsvColumns | ForEach-Object { '"{0}"' -f ([string]$row[$_] -replace '"','""') }) -join ";")
}
[System.IO.File]::WriteAllText($csvRolesFile, $rolesSb.ToString(), $utf8Bom)
Log ("  CSV roles -> {0}" -f $csvRolesFile)

# ============================================================
# ETAPE 2c : WORKSPACE ASSETS
# ============================================================

Log "=== ETAPE 2c : Workspace Assets ==="
$assetsWorkspaceId = Get-AssetsWorkspaceId -Site $siteJiradot
if ($assetsWorkspaceId) { Log ("  Workspace : {0}" -f $assetsWorkspaceId) }
else { Log "  WARN : Workspace Assets non trouve" "WARN" }

# ============================================================
# ETAPE 3 : BOUCLE PRINCIPALE
# ============================================================

Log "=== ETAPE 3 : Identification des comptes jamais utilises ==="

$comptesJamaisUtilises = New-Object System.Collections.Generic.List[object]
$cursor = $null; $startTime = Get-Date
$cTotal=0; $cSkipInactive=0; $cSkipApp=0; $cSkipHorsDomaine=0
$cSkipHorsEmail=0; $cSkipSiteInactif=0; $cSkipAvecActivite=0
$cSkipTropRecent=0; $cJamaisUtilise=0; $cErreurs=0

while ($true) {
  $url = "https://api.atlassian.com/admin/v1/orgs/$orgId/users?maxResults=100"
  if ($cursor) { $url += "&cursor=$cursor" }
  $resp = Invoke-ApiCall -Method "GET" -Url $url -Headers $orgHeaders
  if (-not $resp.ok) { Log ("Erreur API users : status={0}" -f $resp.status) "ERROR"; break }
  $json = $resp.content | ConvertFrom-Json

  foreach ($user in $json.data) {
    $cTotal++
    $accId       = [string]$user.account_id
    $accType     = ""; if ($user.account_type)   { $accType   = [string]$user.account_type }
    $accStatus   = ""; if ($user.account_status) { $accStatus = [string]$user.account_status }
    $displayName = ""; if ($user.name)            { $displayName = [string]$user.name }
    $email       = ""; if ($user.email)           { $email = [string]$user.email }
    $lastActive  = ""; if ($user.last_active)     { $lastActive = [string]$user.last_active }

    if ($accStatus -in @("inactive","closed","suspended")) { $cSkipInactive++; continue }
    if ($accType -eq "app") { $cSkipApp++; continue }

    $onJiradot = $jiradotUsers.ContainsKey($accId)
    $onMutexfr = $mutexfrUsers.ContainsKey($accId)
    if (-not $onJiradot -and -not $onMutexfr) { $cSkipHorsDomaine++; continue }
    if ($email -and -not (Test-EmailDomainAllowed $email)) { $cSkipHorsEmail++; continue }

    $alreadyInactive = $false
    if ($onJiradot -and $jiradotUsers[$accId].active -eq $false -and (-not $onMutexfr -or $mutexfrUsers[$accId].active -eq $false)) { $alreadyInactive = $true }
    if (-not $onJiradot -and $onMutexfr -and $mutexfrUsers[$accId].active -eq $false) { $alreadyInactive = $true }
    if ($alreadyInactive) { $cSkipSiteInactif++; continue }
    if ($lastActive) { $cSkipAvecActivite++; continue }

    $profileUrl = "https://api.atlassian.com/users/$accId/manage/profile"
    $profResp = Invoke-ApiCall -Method "GET" -Url $profileUrl -Headers $orgHeaders
    if (-not $profResp.ok) {
      if ($profResp.status -eq 404) { continue }
      $cErreurs++; Start-Sleep -Milliseconds 100; continue
    }
    $profData = $profResp.content | ConvertFrom-Json

    if (-not $email) {
      if ($profData.account -and $profData.account.email) { $email = [string]$profData.account.email }
      elseif ($profData.email) { $email = [string]$profData.email }
      if ($email -and -not (Test-EmailDomainAllowed $email)) { $cSkipHorsEmail++; continue }
    }

    $dateCreation = ""; $dateInvitation = ""
    if ($profData.account -and $profData.account.created) { $dateCreation = [string]$profData.account.created }
    elseif ($profData.created) { $dateCreation = [string]$profData.created }
    if ($profData.account -and $profData.account.invited) { $dateInvitation = [string]$profData.account.invited }
    elseif ($profData.invited) { $dateInvitation = [string]$profData.invited }

    $dateReference = $dateCreation
    if ($dateInvitation) { $dateReference = $dateInvitation }

    $isOldEnough = $false
    if ($dateReference) {
      try { $dtRef = [DateTimeOffset]::Parse($dateReference); if ($dtRef.DateTime -lt $dateSeuil) { $isOldEnough = $true } } catch {}
    } else { $isOldEnough = $true }
    if (-not $isOldEnough) { $cSkipTropRecent++; continue }

    $productList = ""
    if ($profData.product_access) {
      foreach ($pa in $profData.product_access) {
        if ($pa.products -and ($pa.products | Measure-Object).Count -gt 0) {
          $siteName = ""; if ($pa.url) { $siteName=[string]$pa.url } elseif ($pa.name) { $siteName=[string]$pa.name }
          $prods = ($pa.products | ForEach-Object { $pName=""; if ($_.name){$pName=[string]$_.name}; $pName }) -join ", "
          if ($productList) { $productList += " | " }
          $productList += ("{0}: {1}" -f $siteName, $prods)
        }
      }
    }

    $cJamaisUtilise++
    $domaines = @(); if ($onJiradot) { $domaines += "Jiradot" }; if ($onMutexfr) { $domaines += "mutexfr" }

    $comptesJamaisUtilises.Add(@{
      accountId=$accId; displayName=$displayName; email=$email; accountType=$accType; status=$accStatus
      domaines=($domaines -join " + "); onJiradot=$onJiradot; onMutexfr=$onMutexfr
      dateCreation=$dateCreation; dateCreationFr=Format-DateFr $dateCreation
      dateInvitation=$dateInvitation; dateInvitationFr=Format-DateFr $dateInvitation
      dateReferenceFr=Format-DateFr $dateReference; lastActiveFr="(jamais)"
      productsAPI=$productList; groupes=""
      accesJiraJiradot=$false; accesJiraMutexfr=$false
      accesConfluenceJiradot=$false; accesConfluenceMutexfr=$false; accesResume=""
      isDsim=$false; statutAsset=""; dateEntree=""; dateSortie=""
      dateStatut=""; motifSortie=""
    }) | Out-Null

    Start-Sleep -Milliseconds 100
  }

  Log ("  {0} traites, {1} jamais utilises" -f $cTotal, $cJamaisUtilise)
  $cursor = $null
  if ($json.links -and $json.links.next) {
    $nextUrl = [string]$json.links.next
    if ($nextUrl -match "cursor=([^&]+)") { $cursor = $matches[1] }
  }
  if (-not $cursor) { break }
  Start-Sleep -Milliseconds 200
}

Log ("  JAMAIS UTILISES : {0}" -f $cJamaisUtilise)
# ============================================================
# ETAPE 4 : GROUPES + ACCES PRODUIT + STATUT ASSET v1.6
# ============================================================

Log "=== ETAPE 4 : Groupes + acces produit + statut Asset ==="

$cFetched=0; $cCount=$comptesJamaisUtilises.Count
$cDsim=0; $cAssetTrouve=0; $cAssetNonTrouve=0

foreach ($c in $comptesJamaisUtilises) {
  $cFetched++
  if ($cFetched % 10 -eq 0) {
    Write-Progress -Activity "Verification groupes + Asset" `
      -Status ("{0}/{1} — DSIM:{2} trouve:{3}" -f $cFetched, $cCount, $cDsim, $cAssetTrouve) `
      -PercentComplete ([int](100*$cFetched/$cCount))
  }

  $groupNames = New-Object System.Collections.Generic.List[string]

  if ($c.onJiradot) {
    $url = "{0}/rest/api/3/user/groups?accountId={1}" -f $siteJiradot.BaseUrl, $c.accountId
    $resp = Invoke-ApiCall -Method "GET" -Url $url -Headers $siteJiradot.Headers
    if ($resp.ok) {
      foreach ($g in ($resp.content | ConvertFrom-Json)) {
        $gName = [string]$g.name
        if ($gName -and -not $groupNames.Contains($gName)) { $groupNames.Add($gName) | Out-Null }
      }
    }
    foreach ($gn in $groupNames) { if ($productGroupsJiradot.Contains($gn)) { $c.accesJiraJiradot=$true; break } }
    $c.accesConfluenceJiradot = Test-ConfluenceAccess -Site $siteJiradot -AccountId $c.accountId
  }

  if ($c.onMutexfr) {
    $url = "{0}/rest/api/3/user/groups?accountId={1}" -f $siteMutexfr.BaseUrl, $c.accountId
    $resp = Invoke-ApiCall -Method "GET" -Url $url -Headers $siteMutexfr.Headers
    if ($resp.ok) {
      foreach ($g in ($resp.content | ConvertFrom-Json)) {
        $gName = [string]$g.name
        if ($gName -and -not $groupNames.Contains($gName)) { $groupNames.Add($gName) | Out-Null }
      }
    }
    foreach ($gn in $groupNames) { if ($productGroupsMutexfr.Contains($gn)) { $c.accesJiraMutexfr=$true; break } }
    $c.accesConfluenceMutexfr = Test-ConfluenceAccess -Site $siteMutexfr -AccountId $c.accountId
  }

  $c["groupes"] = if ($groupNames.Count -gt 0) { ($groupNames | Sort-Object) -join " | " } else { "" }

  $accesItems = @()
  if ($c.accesJiraJiradot)       { $accesItems += "Jira(Jiradot)" }
  if ($c.accesJiraMutexfr)       { $accesItems += "Jira(mutexfr)" }
  if ($c.accesConfluenceJiradot) { $accesItems += "Confluence(Jiradot)" }
  if ($c.accesConfluenceMutexfr) { $accesItems += "Confluence(mutexfr)" }
  $c["accesResume"] = if ($accesItems.Count -gt 0) { $accesItems -join " | " } else { "(aucun acces detecte)" }

  $c["isDsim"] = Test-IsDsimGroup $c.groupes

  if ($c.isDsim -and $assetsWorkspaceId) {
    $cDsim++
    $assetResult = Search-AssetByIdentity -Site $siteJiradot -WorkspaceId $assetsWorkspaceId `
                     -DisplayName $c.displayName -Email $c.email
    if ($assetResult) {
      $c["statutAsset"]  = $assetResult.Statut
      $c["dateEntree"]   = $assetResult.DateEntree
      $c["dateSortie"]   = $assetResult.DateSortie
      $c["dateStatut"]   = $assetResult.DateStatut
      $c["motifSortie"]  = $assetResult.MotifSortie
      $cAssetTrouve++
      Log ("  ASSET : {0} -> {1}" -f $c.displayName, $assetResult.Statut) "DEBUG"
    } else {
      $c["statutAsset"]="Non trouve dans Assets"
      $c["dateEntree"]=""; $c["dateSortie"]=""; $c["dateStatut"]=""; $c["motifSortie"]=""
      $cAssetNonTrouve++
      Log ("  ASSET : {0} -> Non trouve" -f $c.displayName) "DEBUG"
    }
  } elseif ($c.isDsim -and -not $assetsWorkspaceId) {
    $c["statutAsset"]="(Assets indisponible)"
    $c["dateEntree"]=""; $c["dateSortie"]=""; $c["dateStatut"]=""; $c["motifSortie"]=""
    $cDsim++
  } else {
    $c["statutAsset"]=""; $c["dateEntree"]=""; $c["dateSortie"]=""; $c["dateStatut"]=""; $c["motifSortie"]=""
  }

  Start-Sleep -Milliseconds 100
}

Write-Progress -Activity "Verification" -Completed
Log ("  DSIM/DSIT : {0} | Asset trouve : {1} | Non trouve : {2}" -f $cDsim, $cAssetTrouve, $cAssetNonTrouve)

# ============================================================
# ETAPE 5 : EXPORT CSV
# ============================================================

Log "=== ETAPE 5 : Export CSV ==="

$csvColumns = @(
  "AccountId","DisplayName","Email","AccountType","Status",
  "Domaines","Groupes","DateCreation","DateInvitation","DerniereActivite",
  "AccesProduit","ProduitsAPI",
  "StatutAsset","DateEntree","DateSortie","DatePLD","MotifSortie","Site"
)

function Export-CsvFile([string]$FilePath, [System.Collections.Generic.List[object]]$Data) {
  $utf8Bom = New-Object System.Text.UTF8Encoding($true)
  $sb = New-Object System.Text.StringBuilder
  [void]$sb.AppendLine(($csvColumns | ForEach-Object { '"{0}"' -f $_ }) -join ";")
  foreach ($c in ($Data | Sort-Object { $_.displayName })) {
    $row = [ordered]@{
      AccountId=$c.accountId; DisplayName=$c.displayName; Email=$c.email
      AccountType=$c.accountType; Status=$c.status; Domaines=$c.domaines
      Groupes=$c.groupes; DateCreation=$c.dateCreationFr; DateInvitation=$c.dateInvitationFr
      DerniereActivite=$c.lastActiveFr; AccesProduit=$c.accesResume; ProduitsAPI=$c.productsAPI
      StatutAsset=$c.statutAsset; DateEntree=$c.dateEntree; DateSortie=$c.dateSortie
      DatePLD=$c.dateStatut; MotifSortie=$c.motifSortie; Site=$c.domaines
    }
    [void]$sb.AppendLine(($csvColumns | ForEach-Object { '"{0}"' -f ([string]$row[$_] -replace '"','""') }) -join ";")
  }
  [System.IO.File]::WriteAllText($FilePath, $sb.ToString(), $utf8Bom)
}

Export-CsvFile -FilePath $csvConsolide -Data $comptesJamaisUtilises
Log ("  CSV consolide -> {0}" -f $csvConsolide)

$dataJiradot = New-Object System.Collections.Generic.List[object]
foreach ($c in $comptesJamaisUtilises) {
  if ($c.onJiradot -and ($c.accesJiraJiradot -or $c.accesConfluenceJiradot)) { $dataJiradot.Add($c) | Out-Null }
}
Export-CsvFile -FilePath $csvJiradot -Data $dataJiradot
Log ("  CSV Jiradot -> {0} ({1} comptes)" -f $csvJiradot, $dataJiradot.Count)

$dataMutexfr = New-Object System.Collections.Generic.List[object]
foreach ($c in $comptesJamaisUtilises) {
  if ($c.onMutexfr -and ($c.accesJiraMutexfr -or $c.accesConfluenceMutexfr)) { $dataMutexfr.Add($c) | Out-Null }
}
Export-CsvFile -FilePath $csvMutexfr -Data $dataMutexfr
Log ("  CSV mutexfr -> {0} ({1} comptes)" -f $csvMutexfr, $dataMutexfr.Count)

# ============================================================
# ETAPE 6 : AFFICHAGE CONSOLE
# ============================================================

Write-Host ""
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ("  {0} COMPTES JAMAIS UTILISES" -f $comptesJamaisUtilises.Count) -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ("  Jiradot avec acces : {0}" -f $dataJiradot.Count) -ForegroundColor Yellow
Write-Host ("  mutexfr avec acces : {0}" -f $dataMutexfr.Count) -ForegroundColor Yellow
Write-Host ("  DSIM/DSIT          : {0}  (trouve:{1} / non trouve:{2})" -f $cDsim, $cAssetTrouve, $cAssetNonTrouve) -ForegroundColor Cyan
Write-Host ""

$displayList = $comptesJamaisUtilises | Sort-Object { $_.displayName }
$displayMax  = [Math]::Min(50, $displayList.Count)
$idx = 0

foreach ($c in $displayList) {
  $idx++
  if ($idx -gt $displayMax) {
    Write-Host ("  ... et {0} autres (voir CSV)" -f ($displayList.Count - $displayMax)) -ForegroundColor DarkGray
    break
  }
  $dateCr  = if ($c.dateCreationFr)   { $c.dateCreationFr }   else { "-" }
  $dateInv = if ($c.dateInvitationFr) { $c.dateInvitationFr } else { "-" }
  $grpDisp = if ($c.groupes)          { $c.groupes }          else { "(aucun groupe)" }
  $assetDisp = if ($c.statutAsset)    { $c.statutAsset }      else { "-" }

  Write-Host ("{0,4}. {1}" -f $idx, $c.displayName) -ForegroundColor $(if ($c.isDsim) { "Cyan" } else { "Yellow" })
  Write-Host ("      Email    : {0}" -f $c.email) -ForegroundColor DarkGray
  Write-Host ("      Domaines : {0}" -f $c.domaines) -ForegroundColor DarkGray
  Write-Host ("      Creation : {0}  |  Invitation : {1}" -f $dateCr, $dateInv) -ForegroundColor DarkGray
  Write-Host ("      Acces    : {0}" -f $c.accesResume) -ForegroundColor DarkGray
  if ($c.isDsim) {
    $assetColor = if ($c.statutAsset -match "Actif") { "Green" } elseif ($c.statutAsset -match "Non trouve") { "Red" } else { "Yellow" }
    Write-Host ("      Asset    : {0}" -f $assetDisp) -ForegroundColor $assetColor
    if ($c.dateEntree)  { Write-Host ("      Entree   : {0}" -f $c.dateEntree) -ForegroundColor DarkGray }
    if ($c.dateSortie)  { Write-Host ("      Sortie   : {0}" -f $c.dateSortie) -ForegroundColor DarkGray }
    if ($c.dateStatut)  { Write-Host ("      DatePLD  : {0}" -f $c.dateStatut) -ForegroundColor DarkGray }
    if ($c.motifSortie) { Write-Host ("      Motif    : {0}" -f $c.motifSortie) -ForegroundColor DarkGray }
  }
  Write-Host ("      Groupes  : {0}" -f $grpDisp) -ForegroundColor DarkCyan
  Write-Host ""
}

# ============================================================
# RESUME FINAL
# ============================================================

$elapsed = (Get-Date) - $startTime

Write-Host "========================================" -ForegroundColor Cyan
Write-Host "  RESUME" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ("  Total scannes          : {0}" -f $cTotal)
Write-Host ("  Skip inactifs          : {0}" -f $cSkipInactive) -ForegroundColor DarkGray
Write-Host ("  Skip apps              : {0}" -f $cSkipApp) -ForegroundColor DarkGray
Write-Host ("  Skip hors sites        : {0}" -f $cSkipHorsDomaine) -ForegroundColor DarkGray
Write-Host ("  Skip email hors domaine: {0}" -f $cSkipHorsEmail) -ForegroundColor DarkGray
Write-Host ("  Skip inactif site      : {0}" -f $cSkipSiteInactif) -ForegroundColor DarkGray
Write-Host ("  Skip avec activite     : {0}" -f $cSkipAvecActivite) -ForegroundColor DarkGray
Write-Host ("  Skip trop recents      : {0}" -f $cSkipTropRecent) -ForegroundColor DarkGray
Write-Host ("  JAMAIS UTILISES        : {0}" -f $comptesJamaisUtilises.Count) -ForegroundColor Yellow
Write-Host ("    Jiradot avec acces   : {0}" -f $dataJiradot.Count) -ForegroundColor Yellow
Write-Host ("    mutexfr avec acces   : {0}" -f $dataMutexfr.Count) -ForegroundColor Yellow
Write-Host ("    DSIM/DSIT            : {0}" -f $cDsim) -ForegroundColor Cyan
Write-Host ("      Asset trouve       : {0}" -f $cAssetTrouve) -ForegroundColor Green
Write-Host ("      Asset non trouve   : {0}" -f $cAssetNonTrouve) -ForegroundColor Yellow
Write-Host ("  Erreurs API            : {0}" -f $cErreurs)
Write-Host ""
Write-Host ("  Endpoint Assets : gateway/api/jsm/assets (v1.6)") -ForegroundColor Green
Write-Host ("  Champs Assets   : Statut, Date Entree, Date Sortie, Date PLD, Motif Sortie") -ForegroundColor Green
Write-Host ("  Duree           : {0:N1} min" -f $elapsed.TotalMinutes)
Write-Host ("  Horodatage      : {0}" -f (Get-Date).ToString("dd/MM/yyyy HH:mm:ss"))
Write-Host ""
Write-Host "  Fichiers :" -ForegroundColor White
Write-Host ("    CSV consolide  : {0}" -f $csvConsolide)
Write-Host ("    CSV Jiradot    : {0}" -f $csvJiradot)
Write-Host ("    CSV mutexfr    : {0}" -f $csvMutexfr)
Write-Host ("    CSV roles      : {0}" -f $csvRolesFile)
Write-Host ("    LOG            : {0}" -f $logFile)
Write-Host "========================================" -ForegroundColor Cyan

Log "Termine."