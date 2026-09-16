[CmdletBinding()]
param()

Add-Type -AssemblyName System.Windows.Forms | Out-Null
Add-Type -AssemblyName System.Globalization | Out-Null
$culture = [System.Globalization.CultureInfo]::InvariantCulture

function Pick-File([string]$title) {
  $dlg = New-Object System.Windows.Forms.OpenFileDialog
  $dlg.Title = $title
  $dlg.Filter = "CSV (*.csv)|*.csv|Tous les fichiers (*.*)|*.*"
  $dlg.Multiselect = $false
  $dlg.CheckFileExists = $true
  $dlg.CheckPathExists = $true
  $null = $dlg.ShowDialog()
  return $dlg.FileName
}

function Detect-Delimiter([string]$headerLine) {
  $counts = @{
    ';' = ([regex]::Matches($headerLine,';').Count)
    ',' = ([regex]::Matches($headerLine,',').Count)
    "`t" = ([regex]::Matches($headerLine,"`t").Count)
  }
  ($counts.GetEnumerator() | Sort-Object Value -Descending | Select-Object -First 1).Key
}

function Import-CsvWithDelim([string]$path, [string]$delim) {
  return Import-Csv -Path $path -Delimiter $delim
}

$path = Pick-File "Sélectionner l'export Audit log (CSV)"
if (-not $path) { throw "Aucun fichier sélectionné." }

$header = (Get-Content -Path $path -TotalCount 1 -Encoding UTF8)
$delim = Detect-Delimiter $header
Write-Host "File: $path"
Write-Host "Delimiter detected: [$delim]"
Write-Host "Header: $header"

$rows = Import-CsvWithDelim -path $path -delim $delim
Write-Host ("Rows: {0}" -f $rows.Count)

$headers = @($rows[0].PSObject.Properties.Name)
Write-Host "Columns:"
$headers | ForEach-Object { " - $_" } | Write-Host

# Filtre invite (assez large)
$inviteRegex = '(?i)\b(invite|invited|invitation|invit[eé]|invité|ajouté|added)\b'

$match = 0
$sample = New-Object System.Collections.Generic.List[object]
foreach ($r in $rows) {
  $all = ($r.PSObject.Properties | ForEach-Object { [string]$_.Value }) -join " | "
  if ($all -match $inviteRegex) {
    $match++
    if ($sample.Count -lt 50) { $sample.Add($r) | Out-Null }
  }
}

Write-Host ("Rows matching inviteRegex: {0}" -f $match)

$ScriptDir  = Split-Path -Parent $PSCommandPath
$ExportsDir = Join-Path $ScriptDir "exports"
if (-not (Test-Path $ExportsDir)) { New-Item -ItemType Directory -Path $ExportsDir -Force | Out-Null }
$out = Join-Path $ExportsDir ("AuditInviteRows_Sample_{0}.csv" -f (Get-Date -Format "yyyyMMdd_HHmmss"))
$sample | Export-Csv -Path $out -Delimiter $delim -NoTypeInformation -Encoding UTF8
Write-Host "Sample export -> $out"