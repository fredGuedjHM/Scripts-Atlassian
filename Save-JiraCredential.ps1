<#
Save-JiraCredential.ps1
- Stocke l'URL Jira + un PSCredential (email + API token) dans un CLIXML
- Fichier chiffré pour ton compte Windows (DPAPI)
- Emplacement : .\secrets\jira-jiradot.cred.xml (à côté de ce script)
#>

$scriptDir = if ($PSCommandPath) { Split-Path -Parent $PSCommandPath } else { (Get-Location).Path }
$secretsDir = Join-Path $scriptDir "secrets"
if (-not (Test-Path $secretsDir)) { New-Item -ItemType Directory -Path $secretsDir -Force | Out-Null }

$outFile = Join-Path $secretsDir "jira-jiradot.cred.xml"

# 1) Demander l'URL une seule fois
$jiraBaseUrl = Read-Host -Prompt "Jira base URL (ex: https://jiradot.atlassian.net)"
$jiraBaseUrl = $jiraBaseUrl.TrimEnd("/")

# 2) Demander les creds (email + API token)
$cred = Get-Credential -Message "Saisis ton email Atlassian (username) + ton API token (password)"

# 3) Stocker un objet contenant URL + Cred
$data = [pscustomobject]@{
    JiraBaseUrl = $jiraBaseUrl
    Credential  = $cred
}

$data | Export-Clixml -Path $outFile

Write-Host "OK. Credential + URL sauvegardés ici :"
Write-Host "  $outFile"
Write-Host "Lisible uniquement par ton compte Windows sur ce poste."