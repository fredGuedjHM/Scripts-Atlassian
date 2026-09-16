$data = Import-Clixml "C:\Users\GUEDJ-F\OneDrive - Harmonie Mutuelle\Documents\powershell\secrets\org-admin.xml"
$orgId = $data.OrgId
$apiKey = [System.Net.NetworkCredential]::new("", $data.ApiKeySecureString).Password
$h = @{Authorization="Bearer $apiKey";Accept="application/json"}

[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

# Structure d'un user + pagination
$url = "https://api.atlassian.com/admin/v1/orgs/$orgId/users?limit=3"
$r = Invoke-WebRequest -Uri $url -Headers $h -UseBasicParsing
$json = $r.Content | ConvertFrom-Json

Write-Host "=== STRUCTURE D'UN USER ===" -ForegroundColor Cyan
$json.data[0] | ConvertTo-Json -Depth 5

Write-Host "`n=== PAGINATION ===" -ForegroundColor Cyan
if ($json.links) { $json.links | ConvertTo-Json -Depth 3 }
if ($json.meta)  { $json.meta  | ConvertTo-Json -Depth 3 }

Write-Host "`n=== TOTAL ===" -ForegroundColor Cyan
Write-Host "data.Count = $($json.data.Count)"