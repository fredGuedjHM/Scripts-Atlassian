<#
Sheet3-SaisiTempsAsset.ps1 (v2.3 - Ventilation par Ticket & Initiative)
Feuille 3 : Saisie hebdomadaire par User x Equipe x Ticket Jira croisee avec Assets.
#>

function Build-SaisiTempsAsset {
    param(
        [Parameter(Mandatory)]$Worklogs,
        [Parameter(Mandatory)][hashtable]$UserTeamsMap,
        [Parameter(Mandatory)]$AssetByAccountId,
        [Parameter(Mandatory)][hashtable]$WorkloadSchemeMap,
        [Parameter(Mandatory)][array]$WeekBuckets,
        [hashtable]$IssuesById = @{},
        [hashtable]$JiraHeaders = @{}
    )

    $headers = @(
        "Compte Jira", "Author ID", "Nom d'Équipe", "Workload Scheme",
        "Nom", "Prénom", "Direction", "Type ressource", "Date Entrée",
        "Issue (Key - Summary)", "Issue Type", "Account", "Initiative (Parent Issue)",
        "Semaine 1", "Semaine 2", "Semaine 3", "Semaine 4", "Semaine 5"
    )

    # 1. Indexation normalisee du referentiel Assets
    $assetMap = @{}
    if ($AssetByAccountId) {
        if ($AssetByAccountId -is [System.Collections.IDictionary]) {
            foreach ($k in $AssetByAccountId.Keys) { $assetMap[[string]$k] = $AssetByAccountId[$k] }
        } elseif ($AssetByAccountId.PSObject -and $AssetByAccountId.PSObject.Properties) {
            foreach ($prop in $AssetByAccountId.PSObject.Properties) { $assetMap[[string]$prop.Name] = $prop.Value }
        } elseif ($AssetByAccountId -is [System.Collections.IEnumerable]) {
            foreach ($item in $AssetByAccountId) {
                $aid = [string]$item."Compte Jira ID"
                if (-not [string]::IsNullOrWhiteSpace($aid)) { $assetMap[$aid] = $item }
            }
        }
    }

    # 2. Agregation des heures par (AuthorId + IssueId) et par Semaine
    $userIssueWeekly = @{}
    $userLoggedSet   = New-Object System.Collections.Generic.HashSet[string]
    $dateWeekCache   = @{}

    foreach ($wl in $Worklogs) {
        $aid = ""
        try { $aid = [string]$wl.author.accountId } catch {}
        if ([string]::IsNullOrWhiteSpace($aid)) { continue }

        $sd = [string]$wl.startDate
        if ([string]::IsNullOrWhiteSpace($sd)) { continue }

        $wIdx = -1
        if ($dateWeekCache.ContainsKey($sd)) {
            $wIdx = $dateWeekCache[$sd]
        } else {
            try { 
                $wlDate = [datetime]::ParseExact($sd, "yyyy-MM-dd", [System.Globalization.CultureInfo]::InvariantCulture)
                $wIdx = Get-WeekIndexForDate -Date $wlDate -Buckets $WeekBuckets
            } catch {
                try {
                    $wlDate = [datetime]::Parse($sd)
                    $wIdx = Get-WeekIndexForDate -Date $wlDate -Buckets $WeekBuckets
                } catch { $wIdx = -1 }
            }
            $dateWeekCache[$sd] = $wIdx
        }

        if ($wIdx -lt 0 -or $wIdx -ge 5) { continue }

        $issueId = ""
        if ($wl.issue -and $wl.issue.id) { $issueId = [string]$wl.issue.id }
        
        $key = "$aid|$issueId"
        if (-not $userIssueWeekly.ContainsKey($key)) {
            $userIssueWeekly[$key] = @(0.0, 0.0, 0.0, 0.0, 0.0)
        }
        $userIssueWeekly[$key][$wIdx] += ([double]$wl.timeSpentSeconds) / 3600.0
        [void]$userLoggedSet.Add($aid)
    }

    $initiativeCache = @{}

    # 3. Tous les collaborateurs a traiter
    $allUsers = New-Object System.Collections.Generic.HashSet[string]
    foreach ($k in $UserTeamsMap.Keys)   { [void]$allUsers.Add([string]$k) }
    foreach ($k in $userLoggedSet)       { [void]$allUsers.Add([string]$k) }

    $rows = New-Object System.Collections.Generic.List[object]
    $userTotal = $allUsers.Count
    $userIndex = 0

    foreach ($authorId in $allUsers) {
        $userIndex++
        if ($userIndex % 50 -eq 0 -or $userIndex -eq $userTotal) {
            $pct = [int](($userIndex / [Math]::Max(1, $userTotal)) * 100)
            Write-Progress -Activity "Construction Feuille 3 (Saisie Hebdomadaire par Ticket)" `
                           -Status "Collaborateur $userIndex / $userTotal ($pct%)" `
                           -PercentComplete $pct
        }

        $teams = @("Non attribué")
        if ($UserTeamsMap.ContainsKey($authorId) -and $UserTeamsMap[$authorId].Count -gt 0) {
            $teams = $UserTeamsMap[$authorId]
        }

        $userData   = if ($assetMap.ContainsKey($authorId)) { $assetMap[$authorId] } else { $null }
        $compteJira = if ($userData) { Extract-EmailFromCompteJira ([string]$userData."Compte Jira") } else { "Inconnu" }
        $nom        = if ($userData -and $userData.Nom) { [string]$userData.Nom } else { "" }
        $prenom     = if ($userData -and $userData.Prénom) { [string]$userData.Prénom } else { "" }
        $typeRes    = if ($userData -and $userData."Type ressource") { [string]$userData."Type ressource" } else { "" }
        $dateEntree = if ($userData -and $userData."Date Entrée") { [string]$userData."Date Entrée" } else { "" }
        
        $direction  = "N/A"
        if ($userData -and -not [string]::IsNullOrWhiteSpace([string]$userData.Direction)) {
            $direction = [string]$userData.Direction
        }

        $wsName = "Non défini"
        if ($WorkloadSchemeMap.ContainsKey($authorId)) { $wsName = $WorkloadSchemeMap[$authorId] }

        $userIssueKeys = @($userIssueWeekly.Keys | Where-Object { $_ -like "$authorId|*" })

        foreach ($teamName in $teams) {
            if ($userIssueKeys.Count -eq 0) {
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
                    "Semaine 1"                 = 0
                    "Semaine 2"                 = 0
                    "Semaine 3"                 = 0
                    "Semaine 4"                 = 0
                    "Semaine 5"                 = 0
                }) | Out-Null
                continue
            }

            foreach ($key in $userIssueKeys) {
                $issueId = $key.Split('|')[1]
                $wh      = $userIssueWeekly[$key]

                $issueKeySummary = "N/A"
                $issueType       = "N/A"
                $account         = "N/A"
                $parentIssue     = "N/A"

                if (-not [string]::IsNullOrWhiteSpace($issueId) -and $IssuesById.ContainsKey($issueId)) {
                    $issue = $IssuesById[$issueId]
                    if ($issue) {
                        $keyStr  = [string]$issue.key
                        $summStr = if ($issue.fields -and $issue.fields.summary) { [string]$issue.fields.summary } else { "" }
                        $issueKeySummary = "$keyStr - $summStr"

                        if ($issue.fields -and $issue.fields.issuetype) {
                            $issueType = [string]$issue.fields.issuetype.name
                        }

                        if ($issue.fields -and $issue.fields.customfield_10032) {
                            $acctField = $issue.fields.customfield_10032
                            if ($acctField.value) { $account = [string]$acctField.value }
                            elseif ($acctField.name) { $account = [string]$acctField.name }
                            else { $account = [string]$acctField }
                        }

                        if (Get-Command Get-InitiativeParentString -ErrorAction SilentlyContinue) {
                            $parentIssue = Get-InitiativeParentString -Issue $issue `
                                -IssuesById $IssuesById -JiraHeaders $JiraHeaders -InitiativeCache $initiativeCache
                        }
                    }
                }

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
                    "Issue (Key - Summary)"     = $issueKeySummary
                    "Issue Type"                = $issueType
                    "Account"                   = $account
                    "Initiative (Parent Issue)" = $parentIssue
                    "Semaine 1"                 = [math]::Round($wh[0], 2)
                    "Semaine 2"                 = [math]::Round($wh[1], 2)
                    "Semaine 3"                 = [math]::Round($wh[2], 2)
                    "Semaine 4"                 = [math]::Round($wh[3], 2)
                    "Semaine 5"                 = [math]::Round($wh[4], 2)
                }) | Out-Null
            }
        }
    }

    Write-Progress -Activity "Construction Feuille 3 (Saisie Hebdomadaire par Ticket)" -Completed
    return @{ Headers = $headers; Rows = $rows }
}
