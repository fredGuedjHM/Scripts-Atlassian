<#
Save-AtlassianOrgAdminKey.ps1
Stocke orgId + Organization API key (Admin) dans .\secrets\org-admin.xml
Chiffrement DPAPI -> lisible uniquement par ton user Windows sur ce poste.
#>

$scriptDir  = Split-Path -Parent $PSCommandPath
$secretsDir = Join-Path $scriptDir "secrets"
if (-not (Test-Path $secretsDir)) { New-Item -ItemType Directory -Path $secretsDir -Force | Out-Null }

$outFile = Join-Path $secretsDir "org-admin.xml"

$orgId = Read-Host -Prompt "OrgId (admin.atlassian.com/o/<OrgId>/...)"

Write-Host "Colle l'Organization API key (Admin)."
$apiKey = Read-Host -Prompt "Organization API key" -AsSecureString

$data = [pscustomobject]@{
  OrgId = $orgId
  ApiKeySecureString = $apiKey
}

$data | Export-Clixml -Path $outFile

Write-Host "OK -> $outFile"