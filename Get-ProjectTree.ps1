[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [string]$Path = '.',

    [Parameter(Mandatory = $false)]
    [string]$GitIgnorePath = '.\.gitignore'
)

Set-StrictMode -Off
$ErrorActionPreference = 'Continue'

$targetPath = (Resolve-Path $Path).Path
$rootName   = (Get-Item $targetPath).Name

# 1. Règles d'exclusion par défaut
$ignoredList = @('.git', '.vscode')

# 2. Lecture du fichier .gitignore
if (Test-Path $GitIgnorePath) {
    $lines = Get-Content $GitIgnorePath
    foreach ($l in $lines) {
        $clean = $l.Trim().TrimEnd('/').TrimEnd('\')
        if ($clean -ne '' -and -not $clean.StartsWith('#')) {
            $ignoredList += $clean.Replace('*', '.*')
        }
    }
}

function Test-ItemIgnored {
    param(
        [string]$Name,
        [string]$RelativePath
    )
    foreach ($pattern in $ignoredList) {
        if ($Name -ieq $pattern -or $Name -imatch "^$pattern$" -or $RelativePath -imatch "^$pattern") {
            return $true
        }
    }
    return $false
}

$script:treeResult = @('```text', "$rootName/")

function Build-Tree {
    param(
        [string]$Dir,
        [string]$Indent = '',
        [string]$Rel = ''
    )

    $children = Get-ChildItem -Path $Dir -Force -ErrorAction SilentlyContinue
    $valid = @()

    foreach ($c in $children) {
        $cRel = if ($Rel -ne '') { "$Rel/$($c.Name)" } else { $c.Name }
        if (-not (Test-ItemIgnored -Name $c.Name -RelativePath $cRel)) {
            $valid += $c
        }
    }

    $sorted = $valid | Sort-Object @{Expression={-not $_.PSIsContainer}}, @{Expression={$_.Name}}
    $count = $sorted.Count
    $i = 0

    foreach ($item in $sorted) {
        $i++
        $isLast = ($i -eq $count)
        $branch = if ($isLast) { '+--- ' } else { '+--- ' }
        $nextIndent = if ($isLast) { '     ' } else { '|    ' }
        $suffix = if ($item.PSIsContainer) { '/' } else { '' }

        $script:treeResult += ($Indent + $branch + $item.Name + $suffix)

        if ($item.PSIsContainer) {
            $newRel = if ($Rel -ne '') { "$Rel/$($item.Name)" } else { $item.Name }
            Build-Tree -Dir $item.FullName -Indent ($Indent + $nextIndent) -Rel $newRel
        }
    }
}

# 3. Construction de l'arbre
Build-Tree -Dir $targetPath
$script:treeResult += '```'

# 4. Affichage console
Write-Host ''
Write-Host 'Arborescence projet (filtree par .gitignore) :' -ForegroundColor Cyan
$script:treeResult | ForEach-Object { Write-Host $_ -ForegroundColor White }
Write-Host ''

# 5. Copie dans le presse-papiers
$script:treeResult | Out-String | Set-Clipboard
Write-Host '[OK] Arborescence copiee dans le presse-papiers !' -ForegroundColor Green