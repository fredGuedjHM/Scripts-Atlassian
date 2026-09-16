<#
Sheet1-WorklogsIssues.ps1
Feuille 1 : Détail des worklogs Tempo croisés avec les issues Jira.
1 ligne par worklog par équipe de l'auteur.
#>

function Get-TempoAccountMetadata {
    param(
        [Parameter(Mandatory)][string]$AccountId,
        [Parameter(Mandatory)][hashtable]$Cache,
        [Parameter(Mandatory)][hashtable]$TempoHeaders
    )
    if ($Cache.ContainsKey($AccountId)) { return $Cache[$AccountId] }
    $resp = Invoke-ApiGet -Url "https://api.tempo.io/4/accounts/$AccountId" -Headers $TempoHeaders
    $meta = @{
        category = $(if ($resp.category) { Fix-DoubleUtf8 $resp.category.name } else { "Non défini" })
        client   = $(if ($resp.customer) { Fix-DoubleUtf8 $resp.customer.name } else { "Non défini" })
    }
    $Cache[$AccountId] = $meta
    return $meta
}

function Get-InitiativeParentString {
    param(
        [Parameter(Mandatory)]$Issue,
        [Parameter(Mandatory)][hashtable]$IssuesById,
        [Parameter(Mandatory)][hashtable]$JiraHeaders
    )
    $current = $Issue; $iter = 0
    while ($Issue.fields -and $Issue.fields.parent) {
        $parent = $Issue.fields.parent
        $parentId = [string]$parent.id
        $parentType = ""
        try { $parentType = $parent.fields.issuetype.name } catch {}

        if ($parentType -and $parentType.ToLower() -eq "initiative") {
            return ("{0} - {1}" -f $parent.key, $parent.fields.summary)
        }
        if ($parentType -and ($parentType.ToLower() -eq "lot" -or $parentType.ToLower() -eq "epic")) {
            if (-not $IssuesById.ContainsKey($parentId)) {
                if ($parent.self) {
                    try {
                        $fetched = Invoke-ApiGet -Url $parent.self -Headers $JiraHeaders
                        if ($fetched -and $fetched.id) { $IssuesById[[string]$fetched.id] = $fetched }
                    } catch { break }
                } else { break }
            }
            if ($IssuesById.ContainsKey($parentId)) { $Issue = $IssuesById[$parentId] }
            else { break }
        } else { break }
        $iter++; if ($iter -gt 20) { break }
    }
    return ("{0} - {1}" -f $current.key, $current.fields.summary)
}

function Build-WorklogsIssues {
    param(
        [Parameter(Mandatory)]$AllUsers,
        [Parameter(Mandatory)][hashtable]$UserTeamsMap,
        [Parameter(Mandatory)][hashtable]$UserWorklogs,
        [Parameter(Mandatory)][hashtable]$IssuesById,
        [Parameter(Mandatory)][hashtable]$AssetByAccountId,
        [Parameter(Mandatory)][hashtable]$TempoAccountMetaCache,
        [Parameter(Mandatory)][hashtable]$TempoHeaders,
        [Parameter(Mandatory)][hashtable]$JiraHeaders
    )

    $headers = @(
        "Start Date","Author ID","Nom d'Équipe","Type ressource","Tarif € TTC",
        "Time Spent (Hours)","Issue (Key - Summary)","Issue Type",
        "Account","Catégorie Account","Client Account","Parent Issue"
    )

    $rows = New-Object System.Collections.Generic.List[object]

    foreach ($authorId in $AllUsers) {
        $teams = @("Non attribué")
        if ($UserTeamsMap.ContainsKey($authorId) -and $UserTeamsMap[$authorId].Count -gt 0) {
            $teams = $UserTeamsMap[$authorId]
        }

        $logs = @()
        if ($UserWorklogs.ContainsKey($authorId)) { $logs = $UserWorklogs[$authorId] }

        $typeRes = "Inconnu"; $tarifTtc = "Non défini"
        if ($AssetByAccountId.ContainsKey($authorId)) {
            $typeRes  = $AssetByAccountId[$authorId]."Type ressource"
            $tarifTtc = $AssetByAccountId[$authorId]."Tarif € TTC"
            if ([string]::IsNullOrWhiteSpace($typeRes))  { $typeRes = "Inconnu" }
            if ([string]::IsNullOrWhiteSpace($tarifTtc)) { $tarifTtc = "Non défini" }
        }

        foreach ($teamName in $teams) {
            if (-not $logs -or $logs.Count -eq 0) {
                $rows.Add([pscustomobject]@{
                    "Start Date"="";"Author ID"=$authorId;"Nom d'Équipe"=$teamName
                    "Type ressource"=$typeRes;"Tarif € TTC"=$tarifTtc
                    "Time Spent (Hours)"="N/A";"Issue (Key - Summary)"="N/A"
                    "Issue Type"="N/A";"Account"="N/A"
                    "Catégorie Account"="N/A";"Client Account"="N/A";"Parent Issue"="N/A"
                }) | Out-Null
                continue
            }

            foreach ($wl in $logs) {
                $startDate = [string]$wl.startDate
                $timeSpentH = ([double]$wl.timeSpentSeconds) / 3600.0
                $issueKeySummary = ""; $issueType = ""; $account = ""
                $accountCategory = "Non défini"; $accountClient = "Non défini"
                $parentIssue = ""

                if ($wl.issue -and $wl.issue.id) {
                    $issueId = [string]$wl.issue.id
                    if ($IssuesById.ContainsKey($issueId)) {
                        $issue = $IssuesById[$issueId]
                        $issueKeySummary = "{0} - {1}" -f $issue.key, $issue.fields.summary
                        $issueType = $issue.fields.issuetype.name

                        $cf = $issue.fields.customfield_10032
                        if ($cf) {
                            if ($cf.value) {
                                $account = [string]$cf.value
                                $tempoAccountId = [string]$cf.id
                                if (-not [string]::IsNullOrWhiteSpace($tempoAccountId)) {
                                    try {
                                        $meta = Get-TempoAccountMetadata -AccountId $tempoAccountId `
                                            -Cache $TempoAccountMetaCache -TempoHeaders $TempoHeaders
                                        $accountCategory = $meta.category
                                        $accountClient   = $meta.client
                                    } catch {}
                                }
                            } else { $account = [string]$cf }
                        }

                        $parentIssue = Get-InitiativeParentString -Issue $issue `
                            -IssuesById $IssuesById -JiraHeaders $JiraHeaders
                    }
                }

                $rows.Add([pscustomobject]@{
                    "Start Date"=$startDate;"Author ID"=$authorId;"Nom d'Équipe"=$teamName
                    "Type ressource"=$typeRes;"Tarif € TTC"=$tarifTtc
                    "Time Spent (Hours)"=$timeSpentH
                    "Issue (Key - Summary)"=$issueKeySummary;"Issue Type"=$issueType
                    "Account"=$account;"Catégorie Account"=$accountCategory
                    "Client Account"=$accountClient;"Parent Issue"=$parentIssue
                }) | Out-Null
            }
        }
    }

    return @{ Headers = $headers; Rows = $rows }
}
