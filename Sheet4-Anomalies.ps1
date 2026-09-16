<#
Sheet4-Anomalies.ps1 v2
Feuille 4 : Anomalies équipe / saisie.
Filtre: DSIM / DSIT / Gouvernance uniquement.
Exclut: inactifs, dates d'entrée futures.
Anomalies: sans équipe, sans saisie, pas de fiche Asset, saisie excessive.
Utilise Get-ExpectedHoursFromWorkload (Common.ps1) pour le prorata Workload Scheme.
Pas de seuil de tolérance sur la saisie excessive.
#>

function Build-Anomalies {
    param(
        [Parameter(Mandatory)]$AssetUsers,
        [Parameter(Mandatory)][hashtable]$UserTeamsMap,
        [Parameter(Mandatory)][hashtable]$UserTeamExitMap,
        [Parameter(Mandatory)][hashtable]$TimeLoggedMap,
        [Parameter(Mandatory)][hashtable]$AssetByAccountId,
        [Parameter(Mandatory)][string]$JiraBaseUrl,
        [Parameter(Mandatory)][hashtable]$JiraHeaders,
        [Parameter(Mandatory)][bool]$UseCacheGroups,
        [Parameter(Mandatory)][string]$GroupsCachePath,
        [Parameter(Mandatory)][bool]$UseCacheWorklogs,
        [Parameter(Mandatory)][string]$JiraUsersCachePath,
        [Parameter(Mandatory)][datetime]$PeriodeFrom,
        [Parameter(Mandatory)][datetime]$PeriodeTo,
        [Parameter(Mandatory)]$JoursFeries,
        [Parameter(Mandatory)][hashtable]$UserWorkloadDays
    )

    $headers = @(
        "Direction", "Groupes Jira", "Compte Jira", "Compte Jira ID",
        "Nom Complet (Jira)", "Nom", "Prénom", "Direction (Asset)",
        "Type ressource", "Statut", "Date Entrée", "Nom d'Équipe",
        "Sortie Équipe Tempo", "Temps saisi (heures)",
        "Heures attendues (prorata)", "Anomalie"
    )

    $rows = New-Object System.Collections.Generic.List[object]
    $todayDate = (Get-Date).Date

    # Cache groupes
    $groupsMap = @{}
    if ($UseCacheGroups) {
        $cg = Load-Json -Path $GroupsCachePath
        if ($cg) { foreach ($p in $cg.PSObject.Properties) { $groupsMap[$p.Name] = @($p.Value) } }
    }

    function Local-GetUserGroups([string]$AccountId) {
        if ($groupsMap.ContainsKey($AccountId)) { return @($groupsMap[$AccountId]) }
        $url = "$JiraBaseUrl/rest/api/3/user/groups?accountId=$AccountId"
        try {
            $resp = Invoke-ApiGet -Url $url -Headers $JiraHeaders
            $gn = @(); foreach ($g in $resp) { if ($g.name) { $gn += [string]$g.name } }
            $groupsMap[$AccountId] = $gn; return $gn
        } catch {
            Write-Warn "Groupes $AccountId : $($_.Exception.Message)"
            return @()
        }
    }

    function Local-IsDirection([string[]]$Groups) {
        foreach ($g in $Groups) {
            if ($g -match "(?i)dsim") { return $true }
            if ($g -match "(?i)dsit") { return $true }
            if ($g -match "(?i)gouvernance") { return $true }
        }
        return $false
    }

    function Local-TestDateFuture([string]$DateStr) {
        if ([string]::IsNullOrWhiteSpace($DateStr)) { return $false }
        $parsed = [datetime]::MinValue
        $cultures = @(
            [System.Globalization.CultureInfo]::new("fr-FR"),
            [System.Globalization.CultureInfo]::InvariantCulture
        )
        foreach ($cu in $cultures) {
            if ([datetime]::TryParse($DateStr, $cu,
                [System.Globalization.DateTimeStyles]::None, [ref]$parsed)) {
                return ($parsed.Date -gt $todayDate)
            }
        }
        return $false
    }

    # ================================================================
    # Cas 1, 2, 3 + saisie excessive : fiches Assets actives
    # ================================================================
    foreach ($u in $AssetUsers) {
        $aid = ([string]$u."Compte Jira ID").Trim()
        if ([string]::IsNullOrWhiteSpace($aid)) { continue }

        $statut = [string]$u.Statut
        if ($statut -notmatch "(?i)^actif$") { continue }

        $dateEntree = [string]$u."Date Entrée"
        if (Local-TestDateFuture -DateStr $dateEntree) {
            Write-Log "Exclu F4 (futur): $($u.Nom) $($u."Prénom") [$aid] entrée=$dateEntree" "DEBUG"
            continue
        }

        $hours = 0.0
        if ($TimeLoggedMap.ContainsKey($aid)) { $hours = [double]$TimeLoggedMap[$aid] }
        $hasTeam = ($UserTeamsMap.ContainsKey($aid) -and $UserTeamsMap[$aid].Count -gt 0)
        $teamName = if ($hasTeam) { ($UserTeamsMap[$aid] | Sort-Object) -join " | " } else { "Aucune équipe" }
        $hoursRounded = [math]::Round($hours, 2)

        # Calcul heures attendues prorata via Workload Scheme
        $heuresAttendue = Get-ExpectedHoursFromWorkload `
            -AccountId $aid `
            -UserWorkloadDays $UserWorkloadDays `
            -AssetByAccountId $AssetByAccountId `
            -MemberDates $null `
            -PeriodeFrom $PeriodeFrom `
            -PeriodeTo $PeriodeTo `
            -JoursFeries $JoursFeries
        $heuresAttendueAff = if ($heuresAttendue -ge 0) { [math]::Round($heuresAttendue, 2) } else { 0 }

        $anomalies = @()
        if (-not $hasTeam -and $hours -eq 0) { $anomalies += "Sans équipe + sans saisie" }
        if (-not $hasTeam -and $hours -gt 0) { $anomalies += "Sans équipe (saisie présente)" }
        if ($hasTeam -and $hours -eq 0)      { $anomalies += "Sans saisie (équipe présente)" }

        # Saisie excessive (pas de tolérance)
        if ($heuresAttendue -gt 0 -and $hours -gt $heuresAttendue) {
            $anomalies += ("Saisie excessive ({0}h / {1}h attendues)" -f $hoursRounded, $heuresAttendueAff)
            Write-Log ("F4 excessive: $aid — {0}h saisies / {1}h attendues" -f $hoursRounded, $heuresAttendueAff) "WARN"
        }

        if ($anomalies.Count -eq 0) { continue }

        $groups = Local-GetUserGroups -AccountId $aid
        if (-not (Local-IsDirection -Groups $groups)) { continue }

        $direction = Get-DirectionFromGroups -Groups $groups
        $groupesConcat = ($groups | Sort-Object) -join " | "
        $sortieEquipe = ""
        if ($UserTeamExitMap.ContainsKey($aid)) { $sortieEquipe = ($UserTeamExitMap[$aid]) -join " | " }

        $rows.Add([pscustomobject]@{
            "Direction"                    = $direction
            "Groupes Jira"                 = $groupesConcat
            "Compte Jira"                  = [string]$u."Compte Jira"
            "Compte Jira ID"               = $aid
            "Nom Complet (Jira)"           = ""
            "Nom"                          = [string]$u.Nom
            "Prénom"                       = [string]$u."Prénom"
            "Direction (Asset)"            = [string]$u.Direction
            "Type ressource"               = [string]$u."Type ressource"
            "Statut"                       = $statut
            "Date Entrée"                  = $dateEntree
            "Nom d'Équipe"                 = $teamName
            "Sortie Équipe Tempo"          = $sortieEquipe
            "Temps saisi (heures)"         = $hoursRounded
            "Heures attendues (prorata)"   = $heuresAttendueAff
            "Anomalie"                     = ($anomalies -join " | ")
        }) | Out-Null
    }

    # ================================================================
    # Cas 4 : users Jira actifs sans fiche Asset
    # ================================================================
    $jiraUsersForSheet4 = $null
    if ($UseCacheWorklogs) {
        $jiraUsersForSheet4 = Load-Json -Path $JiraUsersCachePath
    }
    if (-not $jiraUsersForSheet4) {
        $jiraUsersForSheet4 = @()
        $startAt = 0; $maxResults = 200
        while ($true) {
            $url = "$JiraBaseUrl/rest/api/3/users/search?startAt=$startAt&maxResults=$maxResults"
            $resp = Invoke-ApiGet -Url $url -Headers $JiraHeaders
            if (-not $resp -or $resp.Count -eq 0) { break }
            foreach ($ju in $resp) {
                if ($ju.accountType -eq "atlassian" -and $ju.active -eq $true) {
                    $email = if ($ju.emailAddress) { [string]$ju.emailAddress } else { "" }
                    if ($email -notmatch "@harmonie-mutuelle\.fr$" -and
                        $email -notmatch "@prestataire\.sihm\.fr$") { continue }
                    $jiraUsersForSheet4 += [pscustomobject]@{
                        AccountId   = [string]$ju.accountId
                        DisplayName = [string]$ju.displayName
                        Email       = $email
                    }
                }
            }
            if ($resp.Count -lt $maxResults) { break }
            $startAt += $maxResults
        }
        Write-Info "Feuille 4 - Users Jira: $($jiraUsersForSheet4.Count)"
    }

    foreach ($ju in $jiraUsersForSheet4) {
        $aid = [string]$ju.AccountId
        if ($AssetByAccountId.ContainsKey($aid)) { continue }

        $groups = Local-GetUserGroups -AccountId $aid
        if (-not (Local-IsDirection -Groups $groups)) { continue }

        $direction = Get-DirectionFromGroups -Groups $groups
        $groupesConcat = ($groups | Sort-Object) -join " | "

        $hours = 0.0
        if ($TimeLoggedMap.ContainsKey($aid)) { $hours = [double]$TimeLoggedMap[$aid] }
        $hasTeam = ($UserTeamsMap.ContainsKey($aid) -and $UserTeamsMap[$aid].Count -gt 0)
        $teamName = if ($hasTeam) { ($UserTeamsMap[$aid] | Sort-Object) -join " | " } else { "Aucune équipe" }
        $hoursRounded = [math]::Round($hours, 2)

        # Sans fiche Asset → pas de date d'entrée connue → période complète
        $heuresAttendueFullPeriode = Get-ExpectedHoursFromWorkload `
            -AccountId $aid `
            -UserWorkloadDays $UserWorkloadDays `
            -AssetByAccountId @{} `
            -MemberDates $null `
            -PeriodeFrom $PeriodeFrom `
            -PeriodeTo $PeriodeTo `
            -JoursFeries $JoursFeries
        $heuresAttendueAff = if ($heuresAttendueFullPeriode -ge 0) {
            [math]::Round($heuresAttendueFullPeriode, 2)
        } else { 0 }

        $anomalies = @("Pas de fiche Asset")

        # Saisie excessive (pas de tolérance)
        if ($heuresAttendueFullPeriode -gt 0 -and $hours -gt $heuresAttendueFullPeriode) {
            $anomalies += ("Saisie excessive ({0}h / {1}h attendues)" -f $hoursRounded, $heuresAttendueAff)
        }

        $sortieEquipe = ""
        if ($UserTeamExitMap.ContainsKey($aid)) { $sortieEquipe = ($UserTeamExitMap[$aid]) -join " | " }

        $rows.Add([pscustomobject]@{
            "Direction"                    = $direction
            "Groupes Jira"                 = $groupesConcat
            "Compte Jira"                  = [string]$ju.Email
            "Compte Jira ID"               = $aid
            "Nom Complet (Jira)"           = [string]$ju.DisplayName
            "Nom"                          = ""
            "Prénom"                       = ""
            "Direction (Asset)"            = ""
            "Type ressource"               = ""
            "Statut"                       = "Pas de fiche Asset"
            "Date Entrée"                  = ""
            "Nom d'Équipe"                 = $teamName
            "Sortie Équipe Tempo"          = $sortieEquipe
            "Temps saisi (heures)"         = $hoursRounded
            "Heures attendues (prorata)"   = $heuresAttendueAff
            "Anomalie"                     = ($anomalies -join " | ")
        }) | Out-Null
    }

    # Sauvegarder cache groupes
    Save-Json -Path $GroupsCachePath -Object $groupsMap

    Write-Info "Feuille 4 - Anomalies: $($rows.Count)"
    return @{ Headers = $headers; Rows = $rows }
}
