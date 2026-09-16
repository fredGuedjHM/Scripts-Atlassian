<#
Sheet3-SaisiTempsAsset.ps1
Feuille 3 : Saisie hebdomadaire par user/équipe.
Semaines ISO (lundi→dimanche), 5 colonnes fixes (Semaine 1 à 5).
Inclut le Workload Scheme Tempo.
Direction = "N/A" si non renseignée.
#>

function Build-SaisiTempsAsset {
    param(
        [Parameter(Mandatory)]$Worklogs,
        [Parameter(Mandatory)][hashtable]$UserTeamsMap,
        [Parameter(Mandatory)][hashtable]$AssetByAccountId,
        [Parameter(Mandatory)][hashtable]$WorkloadSchemeMap,
        [Parameter(Mandatory)][array]$WeekBuckets
    )

    $headers = @(
        "Compte Jira", "Author ID", "Nom d'Équipe", "Workload Scheme",
        "Nom", "Prénom", "Direction", "Type ressource", "Date Entrée",
        "Issue (Key - Summary)", "Issue Type", "Account", "Initiative (Parent Issue)",
        "Semaine 1", "Semaine 2", "Semaine 3", "Semaine 4", "Semaine 5"
    )

    # Construire les heures par semaine par user
    $weeklyWorklogs = @{}
    foreach ($wl in $Worklogs) {
        $aid = ""
        try { $aid = [string]$wl.author.accountId } catch {}
        if ([string]::IsNullOrWhiteSpace($aid)) { continue }

        $sd = [string]$wl.startDate
        if ([string]::IsNullOrWhiteSpace($sd)) { continue }

        try { $wlDate = [datetime]::Parse($sd) } catch { continue }

        $wIdx = Get-WeekIndexForDate -Date $wlDate -Buckets $WeekBuckets
        if ($wIdx -lt 0 -or $wIdx -ge 5) { continue }

        if (-not $weeklyWorklogs.ContainsKey($aid)) {
            $weeklyWorklogs[$aid] = @(0.0, 0.0, 0.0, 0.0, 0.0)
        }
        $weeklyWorklogs[$aid][$wIdx] += ([double]$wl.timeSpentSeconds) / 3600.0
    }

    # Tous les users concernés
    $allUsers = New-Object System.Collections.Generic.HashSet[string]
    foreach ($k in $UserTeamsMap.Keys)    { [void]$allUsers.Add($k) }
    foreach ($k in $weeklyWorklogs.Keys)  { [void]$allUsers.Add($k) }

    $rows = New-Object System.Collections.Generic.List[object]

    foreach ($authorId in $allUsers) {
        $teams = @("Non attribué")
        if ($UserTeamsMap.ContainsKey($authorId) -and $UserTeamsMap[$authorId].Count -gt 0) {
            $teams = $UserTeamsMap[$authorId]
        }

        $wh = @(0.0, 0.0, 0.0, 0.0, 0.0)
        if ($weeklyWorklogs.ContainsKey($authorId)) { $wh = $weeklyWorklogs[$authorId] }

        $userData = $null
        if ($AssetByAccountId.ContainsKey($authorId)) { $userData = $AssetByAccountId[$authorId] }

        $compteJira = if ($userData) { Extract-EmailFromCompteJira ([string]$userData."Compte Jira") } else { "Inconnu" }
        $nom        = if ($userData) { [string]$userData.Nom } else { "" }
        $prenom     = if ($userData) { [string]$userData."Prénom" } else { "" }
        $typeRes    = if ($userData) { [string]$userData."Type ressource" } else { "" }
        $dateEntree = if ($userData) { [string]$userData."Date Entrée" } else { "" }

        $direction = "N/A"
        if ($userData -and -not [string]::IsNullOrWhiteSpace([string]$userData.Direction)) {
            $direction = [string]$userData.Direction
        }

        $wsName = "Non défini"
        if ($WorkloadSchemeMap.ContainsKey($authorId)) { $wsName = $WorkloadSchemeMap[$authorId] }

        foreach ($teamName in $teams) {
            $rows.Add([pscustomobject]@{
                "Compte Jira"               = $compteJira
                "Author ID"                 = $authorId
                "Nom d'Équipe"              = $teamName
                "Workload Scheme"           = $wsName
                "Nom"                       = $nom
                "Prénom"                    = $prenom
                "Direction"                 = $direction
                "Type ressource"            = $typeRes
                "Date Entrée"               = $dateEntree
                "Issue (Key - Summary)"     = "N/A"
                "Issue Type"                = "N/A"
                "Account"                   = "N/A"
                "Initiative (Parent Issue)" = "N/A"
                "Semaine 1" = [math]::Round($wh[0], 2)
                "Semaine 2" = [math]::Round($wh[1], 2)
                "Semaine 3" = [math]::Round($wh[2], 2)
                "Semaine 4" = [math]::Round($wh[3], 2)
                "Semaine 5" = [math]::Round($wh[4], 2)
            }) | Out-Null
        }
    }

    return @{ Headers = $headers; Rows = $rows }
}
